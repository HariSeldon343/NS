#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Portali: ASP CO/CZ/KR/RC/VV · Regione Calabria · TAR Catanzaro
"""

import re
import sys
import os
import time
import urllib.parse
from datetime import datetime, timedelta

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

requests.packages.urllib3.disable_warnings()

# ──────────────────────────────────────────────
# DATE SETUP
# ──────────────────────────────────────────────
TODAY = datetime.now()
DATE_FROM_2D = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_FROM_3D = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')
RUN_DATE_STR = TODAY.strftime('%d-%m-%Y')

print(f"Data esecuzione : {TODAY.strftime('%d/%m/%Y %H:%M')}")
print(f"Filtro ASP/Reg  : dal {DATE_FROM_2D}")
print(f"Filtro TAR      : dal {DATE_FROM_3D}")
print()

# ──────────────────────────────────────────────
# KEYWORDS
# ──────────────────────────────────────────────
KEYWORDS = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento", "Aumento di budget",
    "Autismo", "Autorizzazione all'esercizio", "Autorizzazione alla realizzazione",
    "Autorizzazioni", "Budget", "Casa Giardino", "Centro San Giuseppe",
    "Centro salute e benessere", "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento", "San Teodoro",
    "Savelli Hospital", "Starbene", "Verifica requisiti", "Villa San Giuseppe",
    "Villa del Rosario",
    "Life",  # word-boundary match applied in matches_keywords()
]

LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)


def matches_keywords(text: str) -> bool:
    if not text:
        return False
    for kw in KEYWORDS:
        if kw.upper() == "LIFE":
            if LIFE_RE.search(text):
                return True
        else:
            if kw.upper() in text.upper():
                return True
    return False


# ──────────────────────────────────────────────
# HTTP HELPERS
# ──────────────────────────────────────────────
BROWSER_HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9,en;q=0.8',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive',
    'Cache-Control': 'no-cache',
}


def new_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(BROWSER_HEADERS)
    s.verify = False
    return s


def safe_get(session, url, **kw):
    for attempt in range(3):
        try:
            r = session.get(url, timeout=25, **kw)
            r.raise_for_status()
            return r
        except requests.exceptions.Timeout:
            print(f"    [TIMEOUT] GET {url} (tentativo {attempt+1})")
        except requests.exceptions.HTTPError as e:
            print(f"    [HTTP {e.response.status_code}] GET {url}")
            return None
        except Exception as e:
            print(f"    [ERR] GET {url}: {e}")
        time.sleep(2 ** attempt)
    return None


def safe_post(session, url, data=None, **kw):
    for attempt in range(3):
        try:
            r = session.post(url, data=data, timeout=25, **kw)
            r.raise_for_status()
            return r
        except requests.exceptions.Timeout:
            print(f"    [TIMEOUT] POST {url} (tentativo {attempt+1})")
        except requests.exceptions.HTTPError as e:
            print(f"    [HTTP {e.response.status_code}] POST {url}")
            return None
        except Exception as e:
            print(f"    [ERR] POST {url}: {e}")
        time.sleep(2 ** attempt)
    return None


# ──────────────────────────────────────────────
# ASP ALBO ONLINE SCRAPER
# ──────────────────────────────────────────────

def _build_base(url: str) -> str:
    p = urllib.parse.urlparse(url)
    return f"{p.scheme}://{p.netloc}"


def _parse_asp_table(soup: BeautifulSoup, base: str) -> list[dict]:
    rows = []
    table = (
        soup.find('table', id='tabella')
        or soup.find('table', class_=re.compile(r'result|albo|atti', re.I))
    )
    if table is None:
        for t in soup.find_all('table'):
            if len(t.find_all('tr')) > 2:
                table = t
                break
    if table is None:
        return rows

    header_row = table.find('tr')
    if not header_row:
        return rows
    headers = [th.get_text(strip=True).lower()
               for th in header_row.find_all(['th', 'td'])]

    col_data = col_oggetto = col_link = -1
    for i, h in enumerate(headers):
        if 'data' in h and col_data == -1:
            col_data = i
        if any(x in h for x in ('oggetto', 'descri', 'titol')):
            col_oggetto = i
        if any(x in h for x in ('link', 'azion', 'dettagl', 'visu')):
            col_link = i

    if col_data == -1:
        col_data = 0
    if col_oggetto == -1:
        col_oggetto = 1

    for tr in table.find_all('tr')[1:]:
        cells = tr.find_all(['td', 'th'])
        if len(cells) < 2:
            continue

        data_v = cells[col_data].get_text(strip=True) if col_data < len(cells) else ''
        obj_v  = cells[col_oggetto].get_text(strip=True) if col_oggetto < len(cells) else ''

        link_v = ''
        if col_link != -1 and col_link < len(cells):
            a = cells[col_link].find('a', href=True)
            if a:
                link_v = a['href'] if a['href'].startswith('http') else base + a['href']
        if not link_v:
            for cell in cells:
                a = cell.find('a', href=True)
                if a and a['href'] not in ('#', '', 'javascript:void(0)'):
                    link_v = a['href'] if a['href'].startswith('http') else base + a['href']
                    break

        if obj_v:
            rows.append({'data': data_v, 'oggetto': obj_v, 'link': link_v})
    return rows


def _next_page_asp(soup: BeautifulSoup, base: str):
    for a in soup.find_all('a'):
        text = a.get_text(strip=True).lower()
        if any(t in text for t in ['successiv', 'next', '>', '»', 'avanti']):
            href = a.get('href', '')
            if href and href not in ('#', 'javascript:void(0)'):
                return href if href.startswith('http') else base + href
    return None


def scrape_asp(base_url: str, label: str) -> list[dict]:
    print(f"\n{'─'*55}")
    print(f"  [{label}]  {base_url}")
    session = new_session()

    r = safe_get(session, base_url)
    if r is None:
        print(f"  → Non raggiungibile (porta o IP bloccati)")
        return [{'_unreachable': True}]

    soup = BeautifulSoup(r.text, 'lxml')
    base = _build_base(base_url)

    # Collect hidden fields + form action
    form = soup.find('form')
    payload = {}
    form_action = base_url
    if form:
        for inp in form.find_all('input', type='hidden'):
            if inp.get('name'):
                payload[inp['name']] = inp.get('value', '')
        if form.get('action'):
            act = form['action']
            form_action = act if act.startswith('http') else base + act
        btn = form.find('input', type='submit') or form.find('button', type='submit')
        if btn and btn.get('name'):
            payload[btn['name']] = btn.get('value', 'Cerca')

    payload['dataPubblicazioneDal'] = DATE_FROM_2D

    r2 = safe_post(session, form_action, data=payload)
    if r2 is None:
        r2 = safe_get(session, base_url, params={'dataPubblicazioneDal': DATE_FROM_2D})
    if r2 is None:
        print(f"  → Ricerca fallita")
        return [{'_unreachable': True}]

    all_rows = []
    page = 1
    current = r2
    while True:
        s2 = BeautifulSoup(current.text, 'lxml')
        rows = _parse_asp_table(s2, base)
        print(f"  → pag {page}: {len(rows)} atti")
        all_rows.extend(rows)
        nxt = _next_page_asp(s2, base)
        if not nxt:
            break
        current = safe_get(session, nxt)
        if not current:
            break
        page += 1
        time.sleep(0.4)

    matched = [r for r in all_rows if matches_keywords(r.get('oggetto', ''))]
    print(f"  → Totale: {len(all_rows)} atti  |  con keyword: {len(matched)}")
    return matched


# ──────────────────────────────────────────────
# REGIONE CALABRIA SCRAPER
# ──────────────────────────────────────────────

def _parse_regione_table(soup: BeautifulSoup) -> list[dict]:
    """
    Table columns (find_all td+th per row):
      0 td  Tipologia
      1 td  Data Repertoriazione
      2 th  N (numero)
      3 td  Dipartimento
      4 td  Oggetto
      5 td  → link
    """
    rows = []
    table = soup.find('table')
    if not table:
        return rows

    for tr in table.find_all('tr')[1:]:
        cells = tr.find_all(['td', 'th'])
        if len(cells) < 5:
            continue
        tipologia = cells[0].get_text(strip=True) if len(cells) > 0 else ''
        data_v    = cells[1].get_text(strip=True) if len(cells) > 1 else ''
        numero    = cells[2].get_text(strip=True) if len(cells) > 2 else ''
        dip       = cells[3].get_text(strip=True) if len(cells) > 3 else ''
        oggetto_v = cells[4].get_text(strip=True) if len(cells) > 4 else ''

        link_v = ''
        for cell in cells:
            a = cell.find('a', href=True)
            if a and a['href'].startswith('http'):
                link_v = a['href']
                break

        # Build composite oggetto for keyword matching
        full_obj = f"{tipologia} {numero} - {dip} - {oggetto_v}".strip(' -')
        if oggetto_v:
            rows.append({
                'data': data_v,
                'oggetto': full_obj,
                'oggetto_raw': oggetto_v,
                'link': link_v,
            })
    return rows


def scrape_regione_calabria() -> list[dict]:
    print(f"\n{'─'*55}")
    print(f"  [Regione Calabria]")
    base_url = "https://www.regione.calabria.it/provvedimenti-della-regione/"
    session = new_session()
    all_rows = []
    page = 1

    while True:
        params = {
            'filter_active': 'true',
            'filter_date_from': DATE_FROM_2D,
        }
        if page > 1:
            params['paged'] = str(page)

        print(f"  → pag {page} (dal {DATE_FROM_2D})...")
        r = safe_get(session, base_url, params=params)
        if r is None:
            print(f"  → Non raggiungibile")
            return [{'_unreachable': True}]

        soup = BeautifulSoup(r.text, 'lxml')
        rows = _parse_regione_table(soup)
        print(f"     {len(rows)} atti trovati")
        if not rows:
            break
        all_rows.extend(rows)

        # Find max page from pagination links
        pag_links = soup.find_all('a', href=lambda h: h and 'paged=' in str(h))
        max_page = page
        for a in pag_links:
            m = re.search(r'paged=(\d+)', a['href'])
            if m:
                max_page = max(max_page, int(m.group(1)))

        if page >= max_page:
            break
        page += 1
        time.sleep(0.5)

    matched = [r for r in all_rows if matches_keywords(r.get('oggetto', ''))]
    print(f"  → Totale: {len(all_rows)} atti  |  con keyword: {len(matched)}")
    return matched


# ──────────────────────────────────────────────
# TAR CATANZARO SCRAPER
# ──────────────────────────────────────────────

def _parse_tar_table(soup: BeautifulSoup) -> list[dict]:
    rows = []
    table = soup.find('table')
    if not table:
        return rows

    header_row = table.find('tr')
    if not header_row:
        return rows
    headers = [th.get_text(strip=True).lower()
               for th in header_row.find_all(['th', 'td'])]

    col_data = col_parte = col_tipo = col_num = -1
    for i, h in enumerate(headers):
        if 'data' in h and col_data == -1:
            col_data = i
        if 'parte' in h:
            col_parte = i
        if 'tipo' in h:
            col_tipo = i
        if any(x in h for x in ('num', 'n.', 'n°')):
            col_num = i

    if col_data == -1: col_data = 0
    if col_parte == -1: col_parte = 2

    base = 'https://www.giustizia-amministrativa.it'
    for tr in table.find_all('tr')[1:]:
        cells = tr.find_all(['td', 'th'])
        if len(cells) < 2:
            continue
        data_v  = cells[col_data].get_text(strip=True) if col_data < len(cells) else ''
        parte_v = cells[col_parte].get_text(strip=True) if col_parte < len(cells) else ''
        tipo_v  = cells[col_tipo].get_text(strip=True) if col_tipo != -1 and col_tipo < len(cells) else ''
        num_v   = cells[col_num].get_text(strip=True) if col_num != -1 and col_num < len(cells) else ''

        parts = [x for x in [tipo_v, num_v, parte_v] if x]
        oggetto_v = ' | '.join(parts) if parts else parte_v

        link_v = ''
        for cell in cells:
            a = cell.find('a', href=True)
            if a:
                link_v = a['href'] if a['href'].startswith('http') else base + a['href']
                break

        if data_v or parte_v:
            rows.append({'data': data_v, 'oggetto': oggetto_v, 'parte': parte_v, 'link': link_v})
    return rows


def scrape_tar_catanzaro() -> list[dict]:
    print(f"\n{'─'*55}")
    print(f"  [TAR Catanzaro]")
    base_url = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
    session = new_session()
    all_rows = []
    page = 1

    while True:
        params = {'publishDateFrom': DATE_FROM_3D}
        if page > 1:
            params['page'] = str(page)
        print(f"  → pag {page} (dal {DATE_FROM_3D})...")
        r = safe_get(session, base_url, params=params)
        if r is None:
            print(f"  → Non raggiungibile")
            return [{'_unreachable': True}]

        soup = BeautifulSoup(r.text, 'lxml')
        rows = _parse_tar_table(soup)
        print(f"     {len(rows)} atti trovati")
        if not rows:
            break
        all_rows.extend(rows)

        # Pagination
        nxt = soup.find('a', attrs={'aria-label': re.compile(r'next|successiv', re.I)})
        if not nxt:
            nxt = soup.find('li', class_=re.compile(r'next', re.I))
            if nxt:
                nxt = nxt.find('a')
        if not nxt:
            break
        page += 1
        time.sleep(0.5)

    # For TAR: match on parte OR oggetto
    matched = [r for r in all_rows
               if matches_keywords(r.get('parte', '')) or matches_keywords(r.get('oggetto', ''))]
    print(f"  → Totale: {len(all_rows)} atti  |  con keyword: {len(matched)}")
    return matched


# ──────────────────────────────────────────────
# EXCEL WRITER
# ──────────────────────────────────────────────

HDR_FILL  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HDR_FONT  = Font(name='Calibri', bold=True, color="FFFFFF", size=11)
DATA_FONT = Font(name='Calibri', size=10)
LINK_FONT = Font(name='Calibri', size=10, color="0563C1", underline="single")
ALT_FILL  = PatternFill(start_color="EEF3F7", end_color="EEF3F7", fill_type="solid")
THIN = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)
UNREACHABLE_FONT = Font(name='Calibri', size=10, italic=True, color="C00000")


def write_sheet(wb: openpyxl.Workbook, name: str, data: list[dict]):
    ws = wb.create_sheet(title=name[:31])

    # ── Header
    for col, label in enumerate(['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link'], 1):
        c = ws.cell(row=1, column=col, value=label)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = THIN
    ws.row_dimensions[1].height = 25

    # ── Check unreachable flag
    if data and data[0].get('_unreachable'):
        msg = ("⚠  Portale non raggiungibile dall'ambiente di esecuzione remota.\n"
               "I server accettano connessioni solo da IP italiani (VPN/rete regionale).\n"
               "Eseguire lo script localmente oppure tramite VPN italiana.")
        c = ws.cell(row=2, column=1, value=msg)
        c.font = UNREACHABLE_FONT
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells('A2:D2')
        ws.row_dimensions[2].height = 50
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 55
        ws.column_dimensions['C'].width = 45
        ws.column_dimensions['D'].width = 60
        ws.freeze_panes = 'A2'
        print(f"    Foglio '{name}': portale non raggiungibile")
        return

    # ── No results
    if not data:
        c = ws.cell(row=2, column=1, value='Nessun risultato')
        c.font = Font(name='Calibri', size=10, italic=True, color="888888")
        ws.merge_cells('A2:D2')
        ws.cell(row=2, column=1).alignment = Alignment(horizontal='center')
    else:
        for ri, item in enumerate(data, 2):
            obj  = item.get('oggetto', '')
            desc = obj[:150]
            link = item.get('link', '')

            c1 = ws.cell(row=ri, column=1, value=item.get('data', ''))
            c2 = ws.cell(row=ri, column=2, value=obj)
            c3 = ws.cell(row=ri, column=3, value=desc)
            c4 = ws.cell(row=ri, column=4)

            for c in (c1, c2, c3, c4):
                c.font = DATA_FONT
                c.border = THIN
                c.alignment = Alignment(vertical='top', wrap_text=(c.column == 2 or c.column == 3))

            if link:
                c4.value = link
                c4.hyperlink = link
                c4.font = LINK_FONT
            else:
                c4.value = '—'

            if ri % 2 == 0:
                for col in range(1, 5):
                    ws.cell(row=ri, column=col).fill = ALT_FILL

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 60
    ws.freeze_panes = 'A2'
    print(f"    Foglio '{name}': {len(data)} righe")


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

PORTALS = [
    ("https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo", "ASP Cosenza"),
    ("https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo", "ASP Catanzaro"),
    ("https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo", "ASP Crotone"),
    ("https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo", "ASP Reggio Calabria"),
    ("https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo", "ASP Vibo Valentia"),
]
SHEET_ORDER = [
    "ASP Cosenza", "ASP Catanzaro", "ASP Crotone",
    "ASP Reggio Calabria", "ASP Vibo Valentia",
    "Regione Calabria", "TAR Catanzaro",
]


def main():
    print("=" * 55)
    print("AVVIO SCRAPING")
    print("=" * 55)

    results = {}

    # ASP portals
    for url, label in PORTALS:
        results[label] = scrape_asp(url, label)

    # Regione Calabria
    results["Regione Calabria"] = scrape_regione_calabria()

    # TAR Catanzaro
    results["TAR Catanzaro"] = scrape_tar_catanzaro()

    # Build workbook
    print(f"\n{'='*55}")
    print("CREAZIONE EXCEL")
    wb = openpyxl.Workbook()
    del wb['Sheet']  # remove default

    for name in SHEET_ORDER:
        write_sheet(wb, name, results.get(name, []))

    out_file = f"Monitoraggio_Atti_Calabria_{RUN_DATE_STR}.xlsx"
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), out_file)
    wb.save(out_path)

    # Summary
    print(f"\n{'='*55}")
    print("RIEPILOGO")
    print(f"{'─'*55}")
    total_match = 0
    for name in SHEET_ORDER:
        data = results.get(name, [])
        is_unreachable = data and data[0].get('_unreachable')
        if is_unreachable:
            print(f"  {name:<25}  ⚠  non raggiungibile")
        else:
            n = len(data)
            total_match += n
            print(f"  {name:<25}  {n:>3} atti con keyword")
    print(f"{'─'*55}")
    print(f"  {'TOTALE MATCH':<25}  {total_match:>3}")
    print(f"\n✓  File: {out_path}")
    return out_path, out_file


if __name__ == '__main__':
    main()
