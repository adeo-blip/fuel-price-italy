"""Fetch today's Italian fuel prices from the MIMIT open-data feed and append/update
today's national self-service and servito averages in data.json.

Deliberately stdlib-only (no pandas) to minimize failure modes when run from an
unattended scheduled context where extra packages may not be installed.
"""
import urllib.request
import json
import os
import sys
import statistics
from datetime import datetime, timezone

PRICE_URL = "https://www.mimit.gov.it/images/exportCSV/prezzo_alle_8.csv"
DATA_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")

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


def fetch_text(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode('utf-8', errors='replace')


def main():
    raw = fetch_text(PRICE_URL)
    lines = raw.splitlines()
    if len(lines) < 2:
        print("ERROR: fetched file has no data rows", file=sys.stderr)
        sys.exit(1)

    date_str = lines[0].split('Estrazione del')[-1].strip()
    header = lines[1].split('|')
    idx = {name: i for i, name in enumerate(header)}
    required = ['descCarburante', 'prezzo', 'isSelf']
    missing = [c for c in required if c not in idx]
    if missing:
        print(f"ERROR: expected columns missing from feed header: {missing}. Header was: {header}", file=sys.stderr)
        sys.exit(1)

    buckets = {fuel: {'self': [], 'served': []} for fuel in MAIN_FUELS}
    row_count = 0
    for line in lines[2:]:
        if not line.strip():
            continue
        parts = line.split('|')
        if len(parts) <= max(idx.values()):
            continue
        row_count += 1
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
        buckets[fuel]['self' if is_self else 'served'].append(price)

    if row_count == 0:
        print("ERROR: parsed zero data rows from feed — format may have changed", file=sys.stderr)
        sys.exit(1)

    entry = {"date": date_str}
    for fuel in MAIN_FUELS:
        self_prices = buckets[fuel]['self']
        served_prices = buckets[fuel]['served']
        entry[fuel] = {
            "self": round(statistics.mean(self_prices), 4) if self_prices else None,
            "served": round(statistics.mean(served_prices), 4) if served_prices else None,
        }

    if os.path.exists(DATA_JSON):
        with open(DATA_JSON, encoding='utf-8') as f:
            data = json.load(f)
    else:
        data = {"history": []}

    existing_dates = [h['date'] for h in data['history']]
    if date_str in existing_dates:
        for i, h in enumerate(data['history']):
            if h['date'] == date_str:
                data['history'][i] = entry
        print(f"Updated existing entry for {date_str} ({row_count} rows parsed)")
    else:
        data['history'].append(entry)
        data['history'].sort(key=lambda h: h['date'])
        print(f"Added new entry for {date_str} ({row_count} rows parsed)")

    data['last_updated'] = datetime.now(timezone.utc).isoformat()

    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

    print(json.dumps(entry, indent=2))


if __name__ == '__main__':
    main()
