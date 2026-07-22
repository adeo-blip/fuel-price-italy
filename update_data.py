"""Fetch today's Italian fuel data from the MIMIT open-data feeds — prices
(prezzo_alle_8) and the active station registry (anagrafica_impianti_attivi) —
and update the national, Regione, Provincia, Comune and Gestore level JSON
datasets that power the Fuel Dashboard.

Deliberately stdlib-only (no pandas) to minimize failure modes when run from an
unattended scheduled context where extra packages may not be installed.

Regione/Provincia/Comune/Gestore level history is being built fresh starting
from the day this was introduced (no historical backfill). To keep repo size
bounded, Comune and Gestore level files only keep a rolling window of the most
recent COMUNE_GESTORE_RETENTION_DAYS daily entries; Regione, Provincia and the
national data.json keep unbounded history (they're small — tens vs thousands
of keys).
"""
import urllib.request
import json
import os
import sys
import csv
import statistics
from collections import defaultdict, Counter
from datetime import datetime, timezone

PRICE_URL = "https://www.mimit.gov.it/images/exportCSV/prezzo_alle_8.csv"
REGISTRY_URL = "https://www.mimit.gov.it/images/exportCSV/anagrafica_impianti_attivi.csv"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(BASE_DIR, "data.json")
PROVINCE_MAP_CSV = os.path.join(BASE_DIR, "data", "province_regione.csv")
SCOPE_FILES = {
    'regione': os.path.join(BASE_DIR, "data", "data_regione.json"),
    'provincia': os.path.join(BASE_DIR, "data", "data_provincia.json"),
    'comune': os.path.join(BASE_DIR, "data", "data_comune.json"),
    'gestore': os.path.join(BASE_DIR, "data", "data_gestore.json"),
}
COMUNE_RETENTION_DAYS = 60
GESTORE_RETENTION_DAYS = 35
DRILLDOWN_FUELS = ['Benzina', 'Gasolio']

FUEL_MAP = {
    'Benzina': 'Benzina', 'Benzina speciale': 'Benzina', 'Benzina WR 100': 'Benzina',
    'Benzina Plus 98': 'Benzina', 'Benzina Energy 98 ottani': 'Benzina', 'Benzina 100 ottani': 'Benzina',
    'Gasolio': 'Gasolio', 'Gasolio Premium': 'Gasolio', 'Gasolio speciale': 'Gasolio',
    'Gasolio artico': 'Gasolio', 'Gasolio Artico': 'Gasolio', 'Gasolio Oro Diesel': 'Gasolio',
    'Gasolio Alpino': 'Gasolio', 'Gasolio Ecoplus': 'Gasolio', 'Gasolio Gelo': 'Gasolio',
    'Gasolio Energy D': 'Gasolio',
    'Blue Diesel': 'Gasolio', 'Blue Super': 'Benzina', 'Hi-Q Diesel': 'Gasolio',
    'HiQ Perform+': 'Benzina', 'Supreme Diesel': 'Gasolio', 'Excellium Diesel': 'Gasolio',
    'Excellium diesel': 'Gasolio', 'DieselMax': 'Gasolio', 'S-Diesel': 'Gasolio',
    'Diesel e+10': 'Gasolio', 'GP DIESEL': 'Gasolio', 'Blu Diesel Alpino': 'Gasolio',
    'E-DIESEL': 'Gasolio', 'V-Power': 'Benzina', 'V-Power Diesel': 'Gasolio',
    'F101': 'Gasolio', 'R100': 'Gasolio', 'SSP98': 'Benzina',
    'GPL': 'GPL', 'Metano': 'Metano', 'GNL': 'Metano', 'L-GNC': 'Metano',
}
FLOORS = {'Benzina': 1.0, 'Gasolio': 1.0, 'GPL': 0.3, 'Metano': 0.5}
MAIN_FUELS = ['Benzina', 'Gasolio', 'GPL', 'Metano']
UNSPECIFIED = 'Altro / Non specificato'


def fetch_text(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode('utf-8', errors='replace')


def load_province_map():
    sigla_to_provincia = {}
    sigla_to_regione = {}
    name_to_sigla = {}
    with open(PROVINCE_MAP_CSV, encoding='utf-8') as f:
        for row in csv.DictReader(f):
            sigla = row['Sigla della provincia'].strip()
            provincia = row['Provincia'].strip()
            regione = row['Regione'].strip()
            sigla_to_provincia[sigla] = provincia
            sigla_to_regione[sigla] = regione
            name_to_sigla[provincia.upper()] = sigla
    return sigla_to_provincia, sigla_to_regione, name_to_sigla


def resolve_provincia(raw, sigla_to_provincia, sigla_to_regione, name_to_sigla):
    key = (raw or '').strip().upper()
    if key in sigla_to_provincia:
        return sigla_to_provincia[key], sigla_to_regione[key]
    if key in name_to_sigla:
        sigla = name_to_sigla[key]
        return sigla_to_provincia[sigla], sigla_to_regione[sigla]
    return UNSPECIFIED, UNSPECIFIED


def load_registry(sigla_to_provincia, sigla_to_regione, name_to_sigla):
    raw = fetch_text(REGISTRY_URL)
    lines = raw.splitlines()
    if len(lines) < 2:
        print("ERROR: registry feed has no data rows", file=sys.stderr)
        sys.exit(1)
    header = lines[1].split('|')
    idx = {name: i for i, name in enumerate(header)}
    required = ['idImpianto', 'Gestore', 'Comune', 'Provincia']
    missing = [c for c in required if c not in idx]
    if missing:
        print(f"ERROR: registry feed missing columns {missing}. Header: {header}", file=sys.stderr)
        sys.exit(1)

    registry = {}
    for line in lines[2:]:
        if not line.strip():
            continue
        parts = line.split('|')
        if len(parts) <= max(idx.values()):
            continue
        station_id = parts[idx['idImpianto']].strip()
        if not station_id:
            continue
        gestore = parts[idx['Gestore']].strip() or UNSPECIFIED
        comune = parts[idx['Comune']].strip() or UNSPECIFIED
        provincia, regione = resolve_provincia(parts[idx['Provincia']], sigla_to_provincia, sigla_to_regione, name_to_sigla)
        registry[station_id] = {'gestore': gestore, 'comune': comune, 'provincia': provincia, 'regione': regione}
    return registry


def new_bucket():
    return {fuel: {'self': [], 'served': []} for fuel in MAIN_FUELS}


def summarize_buckets(buckets, include_served=True, fuels=MAIN_FUELS, decimals=4):
    out = {}
    for fuel in fuels:
        self_prices = buckets[fuel]['self']
        entry = {"self": round(statistics.mean(self_prices), decimals) if self_prices else None}
        if include_served:
            served_prices = buckets[fuel]['served']
            entry["served"] = round(statistics.mean(served_prices), decimals) if served_prices else None
        out[fuel] = entry
    return out


def load_json(path, default):
    if os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    return default


def write_scope_file(path, date_str, buckets_map, registered_counter, reporting_map, include_served, retention_days=None):
    """Regione/Provincia: small key sets (~20 / ~110) — plain names as keys, unbounded history."""
    doc = load_json(path, {"history": []})
    entries = {}
    for key, buckets in buckets_map.items():
        summary = summarize_buckets(buckets, include_served=include_served)
        summary["stations"] = {
            "registered": registered_counter.get(key, 0),
            "reporting": len(reporting_map.get(key, ())),
        }
        entries[key] = summary
    day_entry = {"date": date_str, "entries": entries}

    history = [h for h in doc['history'] if h['date'] != date_str]
    history.append(day_entry)
    history.sort(key=lambda h: h['date'])
    if retention_days is not None and len(history) > retention_days:
        history = history[-retention_days:]
    doc['history'] = history
    doc['last_updated'] = datetime.now(timezone.utc).isoformat()

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, separators=(',', ':'))


def write_scope_file_indexed(path, date_str, buckets_map, registered_counter, reporting_map, include_served, retention_days=None, fuels=MAIN_FUELS, decimals=4):
    """Comune/Gestore: thousands of keys — repeating the full name every day would bloat the
    file past GitHub's file-size limits within weeks, so names are interned once into a
    top-level, append-only `names` array and daily entries key on their (stable) index instead.
    Also restricted to fewer fuels / coarser precision than Regione/Provincia to keep the
    per-day footprint (the real driver of long-run file size, not the keys) affordable."""
    doc = load_json(path, {"names": [], "history": []})
    name_to_idx = {name: i for i, name in enumerate(doc['names'])}
    entries = {}
    for key, buckets in buckets_map.items():
        if key not in name_to_idx:
            name_to_idx[key] = len(doc['names'])
            doc['names'].append(key)
        idx = name_to_idx[key]
        summary = summarize_buckets(buckets, include_served=include_served, fuels=fuels, decimals=decimals)
        summary["stations"] = {
            "registered": registered_counter.get(key, 0),
            "reporting": len(reporting_map.get(key, ())),
        }
        entries[str(idx)] = summary
    day_entry = {"date": date_str, "entries": entries}

    history = [h for h in doc['history'] if h['date'] != date_str]
    history.append(day_entry)
    history.sort(key=lambda h: h['date'])
    if retention_days is not None and len(history) > retention_days:
        history = history[-retention_days:]
    doc['history'] = history
    doc['last_updated'] = datetime.now(timezone.utc).isoformat()

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(doc, f, ensure_ascii=False, separators=(',', ':'))


def main():
    sigla_to_provincia, sigla_to_regione, name_to_sigla = load_province_map()
    registry = load_registry(sigla_to_provincia, sigla_to_regione, name_to_sigla)

    raw = fetch_text(PRICE_URL)
    lines = raw.splitlines()
    if len(lines) < 2:
        print("ERROR: fetched file has no data rows", file=sys.stderr)
        sys.exit(1)

    date_str = lines[0].split('Estrazione del')[-1].strip()
    header = lines[1].split('|')
    idx = {name: i for i, name in enumerate(header)}
    required = ['idImpianto', 'descCarburante', 'prezzo', 'isSelf']
    missing = [c for c in required if c not in idx]
    if missing:
        print(f"ERROR: expected columns missing from feed header: {missing}. Header was: {header}", file=sys.stderr)
        sys.exit(1)

    national = new_bucket()
    by_regione = defaultdict(new_bucket)
    by_provincia = defaultdict(new_bucket)
    by_comune = defaultdict(new_bucket)
    by_gestore = defaultdict(new_bucket)

    reporting_national = set()
    reporting_regione = defaultdict(set)
    reporting_provincia = defaultdict(set)
    reporting_comune = defaultdict(set)
    reporting_gestore = defaultdict(set)

    row_count = 0
    for line in lines[2:]:
        if not line.strip():
            continue
        parts = line.split('|')
        if len(parts) <= max(idx.values()):
            continue
        row_count += 1

        station_id = parts[idx['idImpianto']].strip()
        station = registry.get(station_id)
        if station:
            reporting_national.add(station_id)
            reporting_regione[station['regione']].add(station_id)
            reporting_provincia[station['provincia']].add(station_id)
            reporting_comune[station['comune']].add(station_id)
            reporting_gestore[station['gestore']].add(station_id)

        fuel_raw = parts[idx['descCarburante']]
        fuel = FUEL_MAP.get(fuel_raw)
        if fuel is None:
            continue
        try:
            price = float(parts[idx['prezzo']])
            is_self = parts[idx['isSelf']].strip() == '1'
        except ValueError:
            continue
        floor = FLOORS[fuel]
        if not (floor < price < 5):
            continue

        bucket_key = 'self' if is_self else 'served'
        national[fuel][bucket_key].append(price)
        if station:
            by_regione[station['regione']][fuel][bucket_key].append(price)
            by_provincia[station['provincia']][fuel][bucket_key].append(price)
            by_comune[station['comune']][fuel][bucket_key].append(price)
            by_gestore[station['gestore']][fuel][bucket_key].append(price)

    if row_count == 0:
        print("ERROR: parsed zero data rows from feed — format may have changed", file=sys.stderr)
        sys.exit(1)

    registered_regione, registered_provincia = Counter(), Counter()
    registered_comune, registered_gestore = Counter(), Counter()
    for st in registry.values():
        registered_regione[st['regione']] += 1
        registered_provincia[st['provincia']] += 1
        registered_comune[st['comune']] += 1
        registered_gestore[st['gestore']] += 1

    # --- national entry (data.json — backward-compatible shape, now with station counts) ---
    entry = {"date": date_str}
    entry.update(summarize_buckets(national, include_served=True))
    entry["stations"] = {"registered": len(registry), "reporting": len(reporting_national)}

    data = load_json(DATA_JSON, {"history": []})
    history = [h for h in data['history'] if h['date'] != date_str]
    history.append(entry)
    history.sort(key=lambda h: h['date'])
    data['history'] = history
    data['last_updated'] = datetime.now(timezone.utc).isoformat()
    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

    # --- Regione / Provincia / Comune / Gestore level ---
    write_scope_file(SCOPE_FILES['regione'], date_str, by_regione, registered_regione, reporting_regione, include_served=True)
    write_scope_file(SCOPE_FILES['provincia'], date_str, by_provincia, registered_provincia, reporting_provincia, include_served=True)
    write_scope_file_indexed(SCOPE_FILES['comune'], date_str, by_comune, registered_comune, reporting_comune,
                              include_served=False, retention_days=COMUNE_RETENTION_DAYS, fuels=DRILLDOWN_FUELS, decimals=3)
    write_scope_file_indexed(SCOPE_FILES['gestore'], date_str, by_gestore, registered_gestore, reporting_gestore,
                              include_served=False, retention_days=GESTORE_RETENTION_DAYS, fuels=DRILLDOWN_FUELS, decimals=3)

    print(f"Processed {row_count} price rows across {len(registry)} registered stations ({len(reporting_national)} reporting today)")
    print(json.dumps(entry, indent=2))


if __name__ == '__main__':
    main()
