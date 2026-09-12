"""Pick "today's" Regione for the daily LinkedIn article and dump its recent
figures — including a weekly trend for all 4 main fuels — into
data/region_of_day.json for the article-writing step to read.

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
from datetime import date, datetime, timedelta, timezone

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROVINCE_MAP_CSV = os.path.join(BASE_DIR, "data", "province_regione.csv")
REGIONE_JSON = os.path.join(BASE_DIR, "data", "data_regione.json")
NATIONAL_JSON = os.path.join(BASE_DIR, "data.json")
OUT_JSON = os.path.join(BASE_DIR, "data", "region_of_day.json")

FUELS = ['Benzina', 'Gasolio', 'GPL', 'Metano']


def load_regions():
    with open(PROVINCE_MAP_CSV, encoding='utf-8') as f:
        return sorted({row['Regione'].strip() for row in csv.DictReader(f)})


def iso_week_key(iso_date_str):
    y, w, _ = date.fromisoformat(iso_date_str).isocalendar()
    return f"{y}-W{w:02d}"


def iso_week_monday(iso_date_str):
    d = date.fromisoformat(iso_date_str)
    return (d - timedelta(days=d.weekday())).isoformat()


def weekly_buckets(history, fuels):
    """One row per ISO week, averaging each fuel's self-service price over the
    days seen that week (a week still in progress is simply averaged over
    however many days have landed so far — days_counted says how many)."""
    groups = {}
    for h in history:
        wk = iso_week_key(h["date"])
        g = groups.setdefault(wk, {"dates": [], **{f: [] for f in fuels}})
        g["dates"].append(h["date"])
        for f in fuels:
            v = (h.get(f) or {}).get("self")
            if v is not None:
                g[f].append(v)
    out = []
    for wk in sorted(groups.keys()):
        g = groups[wk]
        entry = {
            "week": wk,
            "week_start": iso_week_monday(min(g["dates"])),
            "days_counted": len(g["dates"]),
        }
        for f in fuels:
            entry[f] = round(sum(g[f]) / len(g[f]), 4) if g[f] else None
        out.append(entry)
    return out


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
    region_weekly = weekly_buckets(history, FUELS)

    national_latest = None
    national_weekly = []
    if os.path.exists(NATIONAL_JSON):
        with open(NATIONAL_JSON, encoding='utf-8') as f:
            national_doc = json.load(f)
        national_history = national_doc.get("history", [])
        if national_history:
            national_latest = national_history[-1]
        region_week_keys = {w["week"] for w in region_weekly}
        national_weekly = [w for w in weekly_buckets(national_history, FUELS) if w["week"] in region_week_keys]

    out = {
        "date": today.isoformat(),
        "region": region,
        "region_index": today.toordinal() % len(regions),
        "region_cycle_length": len(regions),
        "history": history,
        "latest": latest,
        "previous": previous,
        "weekly": region_weekly,
        "national_latest": national_latest,
        "national_weekly": national_weekly,
        "generated_at": datetime.now(timezone.utc).isoformat(),
    }
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"Region of the day: {region} ({len(history)} day(s) of history)")
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == '__main__':
    main()
