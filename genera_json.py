"""
Leggi il foglio DETTAGLIO C4 da Programmazione TASK FORCE.xlsm e genera data.json.
Uso: python genera_json.py <percorso_excel>
"""
import sys, json, openpyxl
from datetime import datetime

KNOWN_POS = {'KHURAIS', 'HOTEL', 'LICENZA', 'ITALIA'}

def parse(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    # Cerca il foglio C4 (può chiamarsi DETTAGLIO C4 o simile)
    sheet_name = None
    for name in wb.sheetnames:
        if 'C4' in name.upper():
            sheet_name = name
            break
    if not sheet_name:
        raise ValueError("Foglio C4 non trovato nel file Excel")

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))

    # Riga 2 (index 1): date
    header = rows[1]
    dates, date_indices = [], []
    for i, v in enumerate(header):
        if isinstance(v, datetime):
            dates.append(v.strftime('%Y-%m-%d'))
            date_indices.append(i)

    personnel = []
    for row in rows[2:]:
        name_val = row[0] if row else None
        if not name_val or str(name_val).strip() == '':
            continue
        person_name = str(name_val).strip().upper()
        shifts = []
        for col_i, date_str in zip(date_indices, dates):
            v = row[col_i] if col_i < len(row) else None
            pos = str(v).strip().upper() if v else ''
            if pos not in KNOWN_POS:
                pos = ''
            shifts.append({'date': date_str, 'pos': pos})
        personnel.append({'name': person_name, 'shifts': shifts})

    return {'generated': datetime.now().strftime('%Y-%m-%d %H:%M'), 'personnel': personnel}

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python genera_json.py <percorso_excel>")
        sys.exit(1)
    data = parse(sys.argv[1])
    with open('data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    days = len(data['personnel'][0]['shifts']) if data['personnel'] else 0
    print(f"data.json generato — {len(data['personnel'])} persone, {days} giorni")
