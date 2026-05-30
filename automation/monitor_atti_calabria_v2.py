#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
──────────────────────────
Scrapes:
  - Regione Calabria provvedimenti (accessible from cloud)
  - ASP portals × 5   (note: blocked from cloud IPs — note added in sheet)
  - TAR Catanzaro     (note: blocked from cloud IPs — note added in sheet)

Produces a filtered Excel report.
"""

import requests
import urllib3
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
from datetime import datetime, timedelta
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Dates ────────────────────────────────────────────────────────────────────
TODAY      = datetime(2026, 5, 30)
D_FROM_2   = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')   # 2026-05-28
D_FROM_3   = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')   # 2026-05-27
D_TO       = TODAY.strftime('%Y-%m-%d')
OUT_DATE   = TODAY.strftime('%d-%m-%Y')
OUT_FILE   = f'Monitoraggio_Atti_Calabria_{OUT_DATE}.xlsx'

print(f"Data esecuzione : {TODAY.strftime('%d/%m/%Y')}")
print(f"Filtro ASP/RC   : dal {D_FROM_2} al {D_TO}")
print(f"Filtro TAR      : dal {D_FROM_3} al {D_TO}")
print(f"Output          : {OUT_FILE}\n")

# ── Keywords ─────────────────────────────────────────────────────────────────
KEYWORDS = [
    'ADI', 'Assistenza domiciliare', 'ANMIC', 'Accreditamento',
    'Aumento di budget', 'Autismo', "Autorizzazione all'esercizio",
    'Autorizzazione alla realizzazione', 'Autorizzazioni', 'Budget',
    'Casa Giardino', 'Centro San Giuseppe', 'Centro salute e benessere',
    'Fabbisogni LEA', 'Fisiolab', 'Fisioterapia', 'Parere commissione',
    "Presa d'atto verifica", 'Programmazione', 'Rete riabilitativa',
    'Rete territoriale', 'Riabilitazione estensiva', 'Riconversione prestazioni',
    'Rinnovo accreditamento', 'San Teodoro', 'Savelli Hospital', 'Starbene',
    'Verifica requisiti', 'Villa San Giuseppe', 'Villa del Rosario',
]

_LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

def matches(text: str) -> bool:
    if not text:
        return False
    tu = text.upper()
    for kw in KEYWORDS:
        if kw.upper() in tu:
            return True
    return bool(_LIFE_RE.search(text))

# ── HTTP helpers ──────────────────────────────────────────────────────────────
HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9',
}

def get(url, params=None, timeout=30):
    return requests.get(url, params=params, headers=HEADERS,
                        verify=False, timeout=timeout, allow_redirects=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  Regione Calabria
# ═══════════════════════════════════════════════════════════════════════════════

RC_BASE = 'https://www.regione.calabria.it/provvedimenti-della-regione'

def scrape_regione_calabria():
    print("── Regione Calabria ────────────────────────────────────────")
    all_rows = []
    seen    = set()
    page    = 1

    while True:
        params = {
            'filter_date_from': D_FROM_2,
            'filter_date_to':   D_TO,
            'filter_active':    'true',
            'paged':            page,
        }
        try:
            resp = get(RC_BASE, params=params)
            resp.raise_for_status()
        except Exception as e:
            print(f"   Errore pagina {page}: {e}")
            break

        soup = BeautifulSoup(resp.text, 'lxml')
        table = soup.find('table')
        if not table:
            break

        trs = table.find_all('tr')[1:]   # skip header
        new = 0
        for tr in trs:
            cells = tr.find_all(['td', 'th'])
            if len(cells) < 5:
                continue
            texts  = [c.get_text(strip=True) for c in cells]
            anchor = cells[-1].find('a', href=True)
            link   = anchor['href'] if anchor else ''

            if link and link in seen:
                continue
            if link:
                seen.add(link)

            row = {
                'data':    texts[1] if len(texts) > 1 else '',
                'oggetto': texts[4] if len(texts) > 4 else '',
                'link':    link,
            }
            all_rows.append(row)
            new += 1

        print(f"   Pagina {page}: {new} nuove righe (totale {len(all_rows)})")
        if new == 0:
            break
        page += 1
        if page > 200:
            break

    filtered = [r for r in all_rows if matches(r['oggetto'])]
    print(f"   → {len(all_rows)} atti totali, {len(filtered)} corrispondenti ai filtri\n")
    return filtered


# ═══════════════════════════════════════════════════════════════════════════════
#  Excel writer
# ═══════════════════════════════════════════════════════════════════════════════

HDR_FILL  = PatternFill('solid', fgColor='1F4E79')
HDR_FONT  = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
LINK_FONT = Font(name='Calibri', color='0563C1', underline='single', size=10)
BODY_FONT = Font(name='Calibri', size=10)
NOTE_FONT = Font(name='Calibri', italic=True, color='7F7F7F', size=10)
ALT_FILL  = PatternFill('solid', fgColor='EBF3FB')
THIN      = Side(style='thin', color='BFBFBF')
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

COLS   = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']
WIDTHS = [15,     65,        42,                              28]


def _hdr_cell(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.fill   = HDR_FILL
    c.font   = HDR_FONT
    c.border = BORDER
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    return c


def write_sheet(wb, sheet_name: str, results: list, note: str = ''):
    ws = wb.create_sheet(title=sheet_name[:31])

    for i, (col_name, width) in enumerate(zip(COLS, WIDTHS), 1):
        _hdr_cell(ws, 1, i, col_name)
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'

    if note:
        # Print the note across all 4 columns
        ws.merge_cells('A2:D2')
        c = ws['A2']
        c.value     = note
        c.font      = NOTE_FONT
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.row_dimensions[2].height = 45
        return

    if not results:
        ws.merge_cells('A2:D2')
        c = ws['A2']
        c.value     = 'Nessun risultato per il periodo selezionato.'
        c.font      = NOTE_FONT
        c.alignment = Alignment(horizontal='left', vertical='center')
        return

    for ri, r in enumerate(results, 2):
        data_v    = r.get('data', '')
        ogg_v     = r.get('oggetto', '')
        desc_v    = ogg_v[:150]
        link_v    = r.get('link', '')
        is_alt    = (ri % 2 == 0)
        row_fill  = ALT_FILL if is_alt else None

        def _cell(col, value, font=None, align=None):
            c = ws.cell(row=ri, column=col, value=value)
            c.font   = font or BODY_FONT
            c.border = BORDER
            c.alignment = align or Alignment(vertical='top', wrap_text=True)
            if row_fill and not font:
                c.fill = row_fill
            return c

        _cell(1, data_v, align=Alignment(horizontal='center', vertical='top'))
        _cell(2, ogg_v)
        _cell(3, desc_v)

        c4 = ws.cell(row=ri, column=4)
        if link_v:
            c4.value     = 'Apri atto'
            c4.hyperlink = link_v
            c4.font      = Font(name='Calibri', color='0563C1', underline='single',
                                size=10, bold=False)
        else:
            c4.value = '–'
            c4.font  = BODY_FONT
        c4.border    = BORDER
        c4.alignment = Alignment(horizontal='center', vertical='top')
        if row_fill:
            c4.fill = row_fill

    ws.auto_filter.ref = f'A1:D{ws.max_row}'


# ═══════════════════════════════════════════════════════════════════════════════
#  Main
# ═══════════════════════════════════════════════════════════════════════════════

BLOCKED_NOTE = (
    "⚠️  Portale non accessibile dall'ambiente di esecuzione cloud "
    "(i server rispondono con timeout di connessione — probabile restrizione IP). "
    "Per ottenere i dati di questo portale eseguire lo script in locale "
    "(dal proprio PC collegato a una rete italiana)."
)

def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── Regione Calabria ──────────────────────────────────────────────────────
    rc_results = scrape_regione_calabria()
    write_sheet(wb, 'Regione Calabria', rc_results)

    # ── ASP portals (blocked) ─────────────────────────────────────────────────
    asp_sheets = [
        'ASP Cosenza',
        'ASP Catanzaro',
        'ASP Crotone',
        'ASP Reggio Calabria',
        'ASP Vibo Valentia',
    ]
    for name in asp_sheets:
        print(f"── {name}: portale non raggiungibile dall'IP cloud ──")
        write_sheet(wb, name, [], note=BLOCKED_NOTE)

    # ── TAR Catanzaro (blocked) ───────────────────────────────────────────────
    print("── TAR Catanzaro: portale non raggiungibile dall'IP cloud ──")
    write_sheet(wb, 'TAR Catanzaro', [], note=BLOCKED_NOTE)

    # ── Save ─────────────────────────────────────────────────────────────────
    wb.save(OUT_FILE)
    abs_path = os.path.abspath(OUT_FILE)
    print(f"\n✔  File salvato: {abs_path}")
    print(f"   Fogli: {[ws.title for ws in wb.worksheets]}")
    return abs_path


if __name__ == '__main__':
    main()
