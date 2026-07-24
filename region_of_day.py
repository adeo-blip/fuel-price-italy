"""Pick "today's" Regione for the daily LinkedIn article and dump its recent
figures into data/region_of_day.json for the article-writing step to read.

The rotation is alphabetical over Italy's 20 regions (derived from
data/province_regione.csv, not hardcoded) and keyed off the calendar date via
date.toordinal() % 20 — so it always lands on the "right" region for that date
even if a previous day's run was skipped, with no separate counter to drift
out of sync.
"""
import csv
import json
import os
import sys
from datetime import date, datetime, timezone

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROVINCE_MAP_CSV = os.path.join(BASE_DIR, "data", "province_regione.csv")
REGIONE_JSON = os.path.join(BASE_DIR, "data", "data_regione.json")
NATIONAL_JSON = os.path.join(BASE_DIR, "data.json")
OUT_JSON = os.path.join(BASE_DIR, "data", "region_of_day.json")

FUELS = ['Benzina', 'Gasolio', 'GPL', 'Metano']


def load_regions():
    with open(PROVINCE_MAP_CSV, encoding='utf-8') as f:
        return sorted({row['Regione'].strip() for row in csv.DictReader(f)})


def main():
    regions = load_regions()
    today = date.today()
    region = regions[today.toordinal() % len(regions)]

    with open(REGIONE_JSON, encoding='utf-8') as f:
        regione_doc = json.load(f)
    history = [
        {"date": h["date"], **h["entries"][region]}
        for h in regione_doc.get("history", [])
        if region in h.get("entries", {})
    ]
    if not history:
        print(f"ERROR: no history found for region '{region}' in {REGIONE_JSON}", file=sys.stderr)
        sys.exit(1)
    history.sort(key=lambda h: h["date"])
    latest = history[-1]
    previous = history[-2] if len(history) > 1 else None

    national_latest = None
    if os.path.exists(NATIONAL_JSON):
        with open(NATIONAL_JSON, encoding='utf-8') as f:
            national_doc = json.load(f)
        if national_doc.get("history"):
            national_latest = national_doc["history"][-1]

    out = {
        "date": today.isoformat(),
        "region": region,
        "region_index": today.toordinal() % len(regions),
        "region_cycle_length": len(regions),
        "history": history,
        "latest": latest,
        "previous": previous,
        "national_latest": national_latest,
        "generated_at": datetime.now(timezone.utc).isoformat(),
    }
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"Region of the day: {region} ({len(history)} day(s) of history)")
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == '__main__':
    main()
