#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraper per portali ASP, Regione Calabria e TAR Catanzaro
"""

import re
import sys
import requests
import warnings
from datetime import date, timedelta
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

TODAY = date.today()
DATE_FROM_2D = TODAY - timedelta(days=2)   # last 2 days
DATE_FROM_3D = TODAY - timedelta(days=3)   # last 3 days (TAR)

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                  '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
}

KEYWORDS = [
    "ADI",
    "Assistenza domiciliare",
    "ANMIC",
    "Accreditamento",
    "Aumento di budget",
    "Autismo",
    "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione",
    "Autorizzazioni",
    "Budget",
    "Casa Giardino",
    "Centro San Giuseppe",
    "Centro salute e benessere",
    "Fabbisogni LEA",
    "Fisiolab",
    "Fisioterapia",
    "Parere commissione",
    "Presa d'atto verifica",
    "Programmazione",
    "Rete riabilitativa",
    "Rete territoriale",
    "Riabilitazione estensiva",
    "Riconversione prestazioni",
    "Rinnovo accreditamento",
    "San Teodoro",
    "Savelli Hospital",
    "Starbene",
    "Verifica requisiti",
    "Villa San Giuseppe",
    "Villa del Rosario",
]
# "Life" uses word-boundary match separately
KEYWORD_LIFE = re.compile(r'\bLIFE\b', re.IGNORECASE)

# Build a combined regex for all keywords (case-insensitive)
_kw_patterns = [re.escape(k) for k in KEYWORDS]
KEYWORD_RE = re.compile('|'.join(_kw_patterns), re.IGNORECASE)


def matches_keywords(text: str) -> bool:
    if not text:
        return False
    if KEYWORD_RE.search(text):
        return True
    if KEYWORD_LIFE.search(text):
        return True
    return False


def get_session():
    s = requests.Session()
    s.headers.update(HEADERS)
    s.verify = False
    return s


# ─────────────────────────────────────────────
# ASP portals (AlboOnline)
# ─────────────────────────────────────────────

ASP_PORTALS = {
    'ASP Cosenza':        'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Catanzaro':      'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Crotone':        'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Reggio Calabria':'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Vibo Valentia':  'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
}


def scrape_asp(portal_name: str, base_url: str) -> list[dict]:
    """Scrape a single ASP AlboOnline portal, paginating all results."""
    print(f"\n[{portal_name}] Scraping {base_url} ...")
    session = get_session()
    results = []

    # Step 1: GET the search form to grab any hidden fields / CSRF tokens
    try:
        resp = session.get(base_url, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        print(f"  ERROR fetching form page: {e}")
        return _error_result(f"Portale non raggiungibile: {e}")

    soup = BeautifulSoup(resp.text, 'lxml')
    form = soup.find('form')
    if not form:
        print("  ERROR: No form found on page")
        return _error_result("Struttura pagina non riconosciuta")

    # Collect hidden fields
    post_data = {}
    for inp in form.find_all('input'):
        name = inp.get('name')
        val  = inp.get('value', '')
        if name:
            post_data[name] = val

    # Set date filter
    post_data['dataPubblicazioneDal'] = DATE_FROM_2D.isoformat()
    post_data['dataPubblicazioneAl']  = TODAY.isoformat()

    form_action = form.get('action', base_url)
    if not form_action.startswith('http'):
        from urllib.parse import urljoin
        form_action = urljoin(base_url, form_action)

    method = form.get('method', 'post').lower()

    page = 1
    while True:
        try:
            if method == 'get':
                r = session.get(form_action, params=post_data, timeout=30)
            else:
                r = session.post(form_action, data=post_data, timeout=30)
            r.raise_for_status()
        except Exception as e:
            print(f"  ERROR on page {page}: {e}")
            break

        page_soup = BeautifulSoup(r.text, 'lxml')
        rows_found = _parse_asp_table(page_soup, base_url, results)
        if rows_found == 0:
            break

        # Check for next page
        next_link = _asp_next_page(page_soup)
        if not next_link:
            break

        # Navigate to next page (often a GET link)
        if next_link.startswith('http'):
            form_action = next_link
            post_data = {}
            method = 'get'
        else:
            from urllib.parse import urljoin
            form_action = urljoin(base_url, next_link)
            post_data = {}
            method = 'get'
        page += 1
        if page > 200:
            break

    print(f"  Found {len(results)} total records")
    return results


def _parse_asp_table(soup, base_url: str, results: list) -> int:
    """Parse results table from ASP page. Returns number of data rows added."""
    table = soup.find('table', id=lambda x: x and 'result' in x.lower() if x else False)
    if not table:
        table = soup.find('table')
    if not table:
        return 0

    rows = table.find_all('tr')
    if len(rows) <= 1:
        return 0

    # Determine column indices from header row
    header_cells = rows[0].find_all(['th', 'td'])
    col_map = {}
    for i, cell in enumerate(header_cells):
        txt = cell.get_text(strip=True).lower()
        if 'oggetto' in txt or 'descrizione' in txt:
            col_map['oggetto'] = i
        elif 'data' in txt and 'pubbl' in txt:
            col_map['data'] = i
        elif 'data' in txt:
            col_map.setdefault('data', i)
        elif 'numero' in txt or 'prot' in txt:
            col_map['numero'] = i

    if 'oggetto' not in col_map and len(header_cells) >= 2:
        col_map['oggetto'] = 1   # fallback
    if 'data' not in col_map:
        col_map['data'] = 0

    count = 0
    for row in rows[1:]:
        cells = row.find_all('td')
        if not cells:
            continue
        obj_idx = col_map.get('oggetto', 1)
        dat_idx = col_map.get('data', 0)
        obj = cells[obj_idx].get_text(strip=True) if obj_idx < len(cells) else ''
        dat = cells[dat_idx].get_text(strip=True) if dat_idx < len(cells) else ''

        # Find link (detail link)
        link = ''
        for cell in reversed(cells):
            a = cell.find('a', href=True)
            if a:
                href = a.get('href', '')
                if href and not href.startswith('#'):
                    if not href.startswith('http'):
                        from urllib.parse import urljoin
                        href = urljoin(base_url, href)
                    link = href
                    break

        results.append({'data': dat, 'oggetto': obj, 'link': link})
        count += 1
    return count


def _asp_next_page(soup) -> str:
    """Return href of next-page link or empty string."""
    for a in soup.find_all('a'):
        txt = a.get_text(strip=True).lower()
        if txt in ('successiva', 'next', '>', '»', 'pagina successiva', 'avanti'):
            return a.get('href', '')
    # Check aria-label
    for a in soup.find_all('a', attrs={'aria-label': True}):
        if 'next' in a.get('aria-label', '').lower():
            return a.get('href', '')
    return ''


def _error_result(msg: str) -> list[dict]:
    # Shorten 503 error messages
    short = msg
    if '503' in msg:
        short = "Portale non raggiungibile (servizio 503 - server backend non disponibile)"
    elif 'TLS' in msg or 'SSL' in msg or 'tls' in msg:
        short = "Portale non raggiungibile (errore 503/TLS - certificato o connessione)"
    return [{'data': '', 'oggetto': short, 'link': '', '_error': True}]


# ─────────────────────────────────────────────
# Regione Calabria
# ─────────────────────────────────────────────

def scrape_regione_calabria() -> list[dict]:
    print("\n[Regione Calabria] Scraping provvedimenti...")
    base = 'https://www.regione.calabria.it/provvedimenti-della-regione/'
    session = get_session()
    all_records = []
    page = 1

    while True:
        params = {
            'filter_active': 'true',
            'filter_date_from': DATE_FROM_2D.isoformat(),
            'filter_date_to': TODAY.isoformat(),
            'sort_order': 'DESC',
            'paged': page,
        }
        try:
            r = session.get(base, params=params, timeout=30)
            r.raise_for_status()
        except Exception as e:
            print(f"  ERROR page {page}: {e}")
            break

        soup = BeautifulSoup(r.text, 'lxml')
        table = soup.find('table')
        if not table:
            break

        rows = table.find_all('tr')
        # Detect header row to find column indices
        header_row = rows[0] if rows else None
        col_data, col_oggetto, col_link = 1, 3, 4  # defaults
        if header_row:
            hdrs = [c.get_text(strip=True).lower() for c in header_row.find_all(['th','td'])]
            for i, h in enumerate(hdrs):
                if 'data' in h and 'repert' in h:
                    col_data = i
                elif 'oggetto' in h:
                    col_oggetto = i
            col_link = len(hdrs) - 1  # last column = Dettaglio

        data_rows = rows[1:]
        if not data_rows:
            break

        found_any = False
        for row in data_rows:
            cells = row.find_all('td')
            if not cells:
                continue

            n = len(cells)
            # Flexible extraction: Oggetto is always second-to-last before link
            # Link is in last cell
            dat  = cells[col_data].get_text(strip=True)   if col_data < n else ''
            tip  = cells[0].get_text(strip=True)
            dept = cells[min(col_oggetto - 1, n-1)].get_text(strip=True) if n > 2 else ''

            # Find the Oggetto: 5-col layout = [Tipo,Data,Dept,Oggetto,→]
            #                   6-col layout = [Tipo,Data,N,Dept,Oggetto,→]
            if n == 6:
                obj = cells[4].get_text(strip=True)
                dat = cells[1].get_text(strip=True)
                link_cell = cells[5]
            elif n == 5:
                obj = cells[3].get_text(strip=True)
                dat = cells[1].get_text(strip=True)
                link_cell = cells[4]
            elif n >= 3:
                obj = cells[-2].get_text(strip=True)
                dat = cells[1].get_text(strip=True) if n > 1 else ''
                link_cell = cells[-1]
            else:
                continue

            link = ''
            a = link_cell.find('a', href=True)
            if a:
                link = a.get('href', '')

            all_records.append({'data': dat, 'oggetto': obj, 'link': link})
            found_any = True

        if not found_any:
            break

        # Check for next page
        next_btn = soup.find('a', class_='next')
        if not next_btn:
            # look for "Pagina successiva" text
            for a in soup.find_all('a'):
                if 'successiva' in a.get_text(strip=True).lower():
                    next_btn = a
                    break
        if not next_btn:
            break
        page += 1
        if page > 500:
            break

    print(f"  Fetched {len(all_records)} total records from Regione Calabria")
    return all_records


# ─────────────────────────────────────────────
# TAR Catanzaro
# ─────────────────────────────────────────────

def scrape_tar_catanzaro() -> list[dict]:
    print("\n[TAR Catanzaro] Scraping provvedimenti...")
    session = get_session()

    # Try different possible URL patterns for Liferay portals
    candidate_urls = [
        'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro',
        'https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro',
    ]

    for base_url in candidate_urls:
        try:
            r = session.get(base_url, timeout=30)
            if r.status_code == 200 and len(r.text) > 500:
                print(f"  Connected to: {base_url}")
                return _parse_tar_pages(session, base_url, r)
        except Exception as e:
            print(f"  ERROR {base_url}: {e}")

    print("  TAR portal not reachable (503/TLS error)")
    return _error_result("Portale TAR non raggiungibile dal server (503/TLS)")


def _parse_tar_pages(session, base_url: str, first_response) -> list[dict]:
    all_records = []
    soup = BeautifulSoup(first_response.text, 'lxml')

    # Look for date filter form
    form = soup.find('form')
    date_param_name = 'publishDateFrom'

    # Try to submit form with date filter
    post_data = {}
    if form:
        for inp in form.find_all('input'):
            n = inp.get('name')
            v = inp.get('value', '')
            if n:
                post_data[n] = v
        post_data[date_param_name] = DATE_FROM_3D.isoformat()
        form_action = form.get('action', base_url)
        if not form_action.startswith('http'):
            from urllib.parse import urljoin
            form_action = urljoin(base_url, form_action)

        try:
            r = session.post(form_action, data=post_data, timeout=30)
            soup = BeautifulSoup(r.text, 'lxml')
        except Exception as e:
            print(f"  Form submit error: {e}")

    page = 1
    while True:
        rows_found = _parse_tar_table(soup, all_records)
        # Check next page
        next_link = _asp_next_page(soup)
        if not next_link or rows_found == 0:
            break
        try:
            if not next_link.startswith('http'):
                from urllib.parse import urljoin
                next_link = urljoin(base_url, next_link)
            r = session.get(next_link, timeout=30)
            soup = BeautifulSoup(r.text, 'lxml')
        except Exception as e:
            print(f"  Pagination error: {e}")
            break
        page += 1
        if page > 200:
            break

    print(f"  Fetched {len(all_records)} total records from TAR")
    return all_records


def _parse_tar_table(soup, results: list) -> int:
    table = soup.find('table')
    if not table:
        return 0
    rows = table.find_all('tr')
    count = 0
    for row in rows[1:]:
        cells = row.find_all('td')
        if not cells:
            continue
        # Find 'Parte' column (TAR specific) and date
        dat, parte, link = '', '', ''
        for cell in cells:
            txt = cell.get_text(strip=True)
            a = cell.find('a', href=True)
            if a:
                link = a.get('href', '')
        if len(cells) >= 2:
            dat  = cells[0].get_text(strip=True)
            parte = cells[1].get_text(strip=True)
        results.append({'data': dat, 'oggetto': parte, 'link': link})
        count += 1
    return count


# ─────────────────────────────────────────────
# Excel output
# ─────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", fgColor="1F3864")
HEADER_FONT  = Font(bold=True, color="FFFFFF", size=11)
LINK_FONT    = Font(color="0563C1", underline="single")
ALT_FILL     = PatternFill("solid", fgColor="DCE6F1")
BORDER_SIDE  = Side(style="thin", color="AAAAAA")
CELL_BORDER  = Border(left=BORDER_SIDE, right=BORDER_SIDE,
                       top=BORDER_SIDE, bottom=BORDER_SIDE)
COL_WIDTHS   = [15, 70, 55, 40]  # Data, Oggetto, Descrizione breve, Link


def write_sheet(ws, records: list[dict]):
    # Headers
    ws.append(["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"])
    header_row = ws[1]
    for cell in header_row:
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws.row_dimensions[1].height = 22

    if not records or (len(records) == 1 and records[0].get('_error')):
        # No data or error
        msg = records[0]['oggetto'] if records else 'Nessun risultato'
        ws.append([msg])
        cell = ws['A2']
        cell.font = Font(italic=True, color="666666")
        is_error = records and records[0].get('_error') and 'non raggiungibile' in msg
        if is_error:
            cell.fill = PatternFill("solid", fgColor="FFD7D7")
            cell.font = Font(italic=True, bold=True, color="CC0000")
        return

    for i, rec in enumerate(records, start=2):
        obj   = rec.get('oggetto', '')
        dat   = rec.get('data', '')
        link  = rec.get('link', '')
        brief = obj[:150]

        row = [dat, obj, brief, link if link else '']
        ws.append(row)
        r = ws[i]

        # Alternate row fill
        if i % 2 == 0:
            for cell in r:
                cell.fill = ALT_FILL

        # Borders and alignment
        for cell in r:
            cell.border    = CELL_BORDER
            cell.alignment = Alignment(vertical='top', wrap_text=True)

        ws.row_dimensions[i].height = 40

        # Hyperlink in column D
        link_cell = ws.cell(row=i, column=4)
        if link:
            link_cell.value     = link
            link_cell.hyperlink = link
            link_cell.font      = LINK_FONT
        else:
            link_cell.value = 'N/D'

    # Column widths
    for col_idx, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze header row
    ws.freeze_panes = 'A2'


def generate_excel(sheet_data: dict[str, list[dict]], output_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    sheet_names = [
        'ASP Cosenza',
        'ASP Catanzaro',
        'ASP Crotone',
        'ASP Reggio Calabria',
        'ASP Vibo Valentia',
        'Regione Calabria',
        'TAR Catanzaro',
    ]

    for sheet_name in sheet_names:
        ws = wb.create_sheet(title=sheet_name)
        records = sheet_data.get(sheet_name, [])

        if not records:
            filtered = []
        else:
            has_error = len(records) == 1 and records[0].get('_error')
            if has_error:
                filtered = records  # show error message
            else:
                filtered = [r for r in records if matches_keywords(r.get('oggetto', ''))]

        if not filtered and not (records and records[0].get('_error')):
            write_sheet(ws, [{'data': '', 'oggetto': 'Nessun risultato', 'link': '', '_error': True}])
        else:
            write_sheet(ws, filtered)

        print(f"  Sheet '{sheet_name}': {len(filtered) if not (filtered and filtered[0].get('_error')) else 0} matching records")

    wb.save(output_path)
    print(f"\nFile saved: {output_path}")


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────

def main():
    print(f"=== Monitoraggio Atti Calabria ===")
    print(f"Data di esecuzione: {TODAY.strftime('%d-%m-%Y')}")
    print(f"Filtro date: dal {DATE_FROM_2D} al {TODAY}")
    print(f"TAR date: dal {DATE_FROM_3D} al {TODAY}")

    sheet_data = {}

    # Scrape ASP portals
    for name, url in ASP_PORTALS.items():
        sheet_data[name] = scrape_asp(name, url)

    # Scrape Regione Calabria
    sheet_data['Regione Calabria'] = scrape_regione_calabria()

    # Scrape TAR Catanzaro
    sheet_data['TAR Catanzaro'] = scrape_tar_catanzaro()

    # Generate Excel
    filename = f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
    output_path = f"/home/user/NS/{filename}"
    print(f"\n=== Generating Excel: {filename} ===")
    generate_excel(sheet_data, output_path)

    # Summary
    print("\n=== SUMMARY ===")
    for sheet_name in ['ASP Cosenza', 'ASP Catanzaro', 'ASP Crotone',
                       'ASP Reggio Calabria', 'ASP Vibo Valentia',
                       'Regione Calabria', 'TAR Catanzaro']:
        records = sheet_data.get(sheet_name, [])
        has_error = len(records) == 1 and records[0].get('_error')
        if has_error:
            print(f"  {sheet_name}: ERRORE - {records[0]['oggetto'][:80]}")
        else:
            matched = [r for r in records if matches_keywords(r.get('oggetto', ''))]
            print(f"  {sheet_name}: {len(records)} atti totali, {len(matched)} matching keywords")

    return output_path


if __name__ == '__main__':
    out = main()
    print(f"\nFile output: {out}")
