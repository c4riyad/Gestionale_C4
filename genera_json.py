"""
Leggi il foglio dal file Programmazione TUC4.xlsx e genera data.json per il portale web.
Uso: python genera_json.py <percorso_excel>
"""
import sys
import json
import openpyxl
from datetime import datetime

def parse(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Riga 1: colonna 0 = "PERSONALE", colonne 1..N = date
    header = rows[0]
    dates = []
    for v in header[1:]:
        if isinstance(v, datetime):
            dates.append(v.strftime('%Y-%m-%d'))
        elif v is None:
            dates.append(None)
        else:
            dates.append(str(v))

    personnel = []
    for row in rows[1:]:
        name = row[0]
        if not name or str(name).strip() == '':
            continue
        shifts = []
        for i, v in enumerate(row[1:len(dates)+1]):
            shifts.append({
                'date': dates[i],
                'pos':  str(v).strip() if v else ''
            })
        personnel.append({'name': str(name).strip(), 'shifts': shifts})

    return {'generated': datetime.now().strftime('%Y-%m-%d %H:%M'), 'personnel': personnel}

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python genera_json.py <percorso_excel>")
        sys.exit(1)
    data = parse(sys.argv[1])
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"data.json generato — {len(data['personnel'])} persone, {len(data['personnel'][0]['shifts']) if data['personnel'] else 0} giorni")
