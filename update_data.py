import urllib.request
import pandas as pd
import json, os
from io import StringIO
from datetime import datetime, timezone

PRICE_URL = "https://www.mimit.gov.it/images/exportCSV/prezzo_alle_8.csv"
DATA_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")

FUEL_MAP = {
    'Benzina':'Benzina', 'Benzina speciale':'Benzina', 'Benzina WR 100':'Benzina',
    'Benzina Plus 98':'Benzina', 'Benzina Energy 98 ottani':'Benzina', 'Benzina 100 ottani':'Benzina',
    'Gasolio':'Gasolio', 'Gasolio Premium':'Gasolio', 'Gasolio speciale':'Gasolio',
    'Gasolio artico':'Gasolio', 'Gasolio Artico':'Gasolio', 'Gasolio Oro Diesel':'Gasolio',
    'Gasolio Alpino':'Gasolio', 'Gasolio Ecoplus':'Gasolio', 'Gasolio Gelo':'Gasolio',
    'Gasolio Energy D':'Gasolio',
    'Blue Diesel':'Gasolio', 'Blue Super':'Benzina', 'Hi-Q Diesel':'Gasolio',
    'HiQ Perform+':'Benzina', 'Supreme Diesel':'Gasolio', 'Excellium Diesel':'Gasolio',
    'Excellium diesel':'Gasolio', 'DieselMax':'Gasolio', 'S-Diesel':'Gasolio',
    'Diesel e+10':'Gasolio', 'GP DIESEL':'Gasolio', 'Blu Diesel Alpino':'Gasolio',
    'E-DIESEL':'Gasolio', 'V-Power':'Benzina', 'V-Power Diesel':'Gasolio',
    'F101':'Gasolio', 'R100':'Gasolio', 'SSP98':'Benzina',
    'GPL':'GPL', 'Metano':'Metano', 'GNL':'Metano', 'L-GNC':'Metano',
}
FLOORS = {'Benzina':1.0, 'Gasolio':1.0, 'GPL':0.3, 'Metano':0.5}
MAIN_FUELS = ['Benzina','Gasolio','GPL','Metano']


def fetch_text(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode('utf-8', errors='replace')


def main():
    raw = fetch_text(PRICE_URL)
    lines = raw.splitlines()
    date_str = lines[0].split('Estrazione del')[-1].strip()

    df = pd.read_csv(StringIO('\n'.join(lines[1:])), sep='|')
    df = df.rename(columns={'idImpianto': 'id', 'descCarburante': 'fuel_raw', 'prezzo': 'price', 'isSelf': 'isSelf'})
    df['fuel'] = df['fuel_raw'].map(FUEL_MAP)
    df = df[df['fuel'].isin(MAIN_FUELS)].copy()
    floor = df['fuel'].map(FLOORS)
    df = df[(df['price'] > floor) & (df['price'] < 5)]

    entry = {"date": date_str}
    for fuel in MAIN_FUELS:
        sub = df[df['fuel'] == fuel]
        self_avg = sub[sub['isSelf'] == 1]['price'].mean()
        served_avg = sub[sub['isSelf'] == 0]['price'].mean()
        entry[fuel] = {
            "self": round(float(self_avg), 4) if pd.notna(self_avg) else None,
            "served": round(float(served_avg), 4) if pd.notna(served_avg) else None,
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
        print(f"Updated existing entry for {date_str}")
    else:
        data['history'].append(entry)
        data['history'].sort(key=lambda h: h['date'])
        print(f"Added new entry for {date_str}")

    data['last_updated'] = datetime.now(timezone.utc).isoformat()

    with open(DATA_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

    print(json.dumps(entry, indent=2))


if __name__ == '__main__':
    main()
