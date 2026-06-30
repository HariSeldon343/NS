#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scrapes ASP portals, Regione Calabria, and TAR Catanzaro for recent acts.
"""

import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import time
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuration ──────────────────────────────────────────────────────────────

CA_BUNDLE = '/root/.ccr/ca-bundle.crt'
TODAY = datetime.now()
DATE_FROM_2 = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')   # ASP / Regione: last 2 days
DATE_FROM_3 = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')   # TAR: last 3 days

KEYWORDS = [
    'ADI',
    'Assistenza domiciliare',
    'ANMIC',
    'Accreditamento',
    'Aumento di budget',
    'Autismo',
    "Autorizzazione all'esercizio",
    'Autorizzazione alla realizzazione',
    'Autorizzazioni',
    'Budget',
    'Casa Giardino',
    'Centro San Giuseppe',
    'Centro salute e benessere',
    'Fabbisogni LEA',
    'Fisiolab',
    'Fisioterapia',
    'Parere commissione',
    "Presa d'atto verifica",
    'Programmazione',
    'Rete riabilitativa',
    'Rete territoriale',
    'Riabilitazione estensiva',
    'Riconversione prestazioni',
    'Rinnovo accreditamento',
    'San Teodoro',
    'Savelli Hospital',
    'Starbene',
    'Verifica requisiti',
    'Villa San Giuseppe',
    'Villa del Rosario',
]
# "Life" uses word boundary to avoid false positives
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

def keyword_match(text):
    """Return True if text contains at least one keyword (case-insensitive)."""
    t = text or ''
    # Word-boundary check for LIFE
    if LIFE_RE.search(t):
        return True
    t_lo = t.lower()
    for kw in KEYWORDS:
        if kw.lower() in t_lo:
            return True
    return False

def make_session():
    s = requests.Session()
    s.verify = CA_BUNDLE
    s.headers.update({
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/124.0.0.0 Safari/537.36'
        ),
        'Accept-Language': 'it-IT,it;q=0.9,en;q=0.8',
    })
    return s

# ── ASP Portals (sisr.regione.calabria.it) ─────────────────────────────────────

ASP_PORTALS = {
    'ASP Cosenza':        'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Catanzaro':      'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Crotone':        'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Reggio Calabria':'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Vibo Valentia':  'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
}

def scrape_asp(name, base_url, session):
    """
    Try to scrape an ASP Albo Online portal.
    Returns (list_of_rows, error_message).
    Each row: {'data': str, 'oggetto': str, 'link': str}
    """
    print(f"\n{'='*60}")
    print(f"Scraping {name} — {base_url}")

    results = []
    error = None

    try:
        # Step 1: GET the search page
        r = session.get(base_url, timeout=25)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, 'html.parser')

        # The Albo Online uses a JSF/Spring form with dataPubblicazioneDal
        form = soup.find('form')
        if not form:
            return [], "Form non trovato nella pagina"

        # Collect hidden inputs (ViewState, tokens)
        post_data = {}
        for inp in form.find_all('input'):
            n = inp.get('name', '')
            v = inp.get('value', '')
            if n:
                post_data[n] = v

        # Set the date filter
        post_data['dataPubblicazioneDal'] = DATE_FROM_2
        # Clear 'al' so it defaults to today
        if 'dataPubblicazioneAl' in post_data:
            post_data['dataPubblicazioneAl'] = TODAY.strftime('%Y-%m-%d')

        action = form.get('action', base_url)
        if action.startswith('/'):
            from urllib.parse import urlparse
            p = urlparse(base_url)
            action = f"{p.scheme}://{p.netloc}{action}"

        # Step 2: POST search
        headers = {
            'Referer': base_url,
            'Content-Type': 'application/x-www-form-urlencoded',
            'Origin': base_url.split('/AlboOnline')[0],
        }
        r2 = session.post(action, data=post_data, headers=headers, timeout=25)
        r2.raise_for_status()
        soup2 = BeautifulSoup(r2.text, 'html.parser')

        page = 1
        while True:
            rows, has_more = _parse_asp_page(soup2, base_url)
            results.extend(rows)
            print(f"  Page {page}: {len(rows)} rows")

            if not has_more:
                break

            # Navigate to next page - look for pagination link/button
            next_btn = soup2.find('a', string=re.compile(r'[Ss]uccessiv|Next|»|›'))
            if not next_btn:
                break

            next_url = next_btn.get('href', '')
            if not next_url:
                break

            if next_url.startswith('/'):
                from urllib.parse import urlparse
                p = urlparse(base_url)
                next_url = f"{p.scheme}://{p.netloc}{next_url}"

            r3 = session.get(next_url, timeout=25)
            r3.raise_for_status()
            soup2 = BeautifulSoup(r3.text, 'html.parser')
            page += 1
            if page > 50:
                break

    except Exception as e:
        error = str(e)
        print(f"  ERROR: {error}")

    return results, error


def _parse_asp_page(soup, base_url):
    """Parse one page of ASP Albo Online results. Returns (rows, has_more)."""
    rows = []
    table = soup.find('table', id=re.compile(r'albo|result|list', re.I)) or soup.find('table')
    if not table:
        return rows, False

    trs = table.find_all('tr')
    for tr in trs[1:]:  # skip header
        tds = tr.find_all('td')
        if not tds:
            continue
        # Typical columns: Numero Atto | Data | Oggetto | Tipo | Link
        # Column order may vary – we look for 'data' and 'oggetto' columns
        data_val = ''
        oggetto_val = ''
        link_val = ''

        texts = [td.get_text(strip=True) for td in tds]
        link_el = tr.find('a')
        if link_el:
            href = link_el.get('href', '')
            if href.startswith('/'):
                from urllib.parse import urlparse
                p = urlparse(base_url)
                link_val = f"{p.scheme}://{p.netloc}{href}"
            else:
                link_val = href

        # Heuristic: find a date-looking cell and use the longest text as oggetto
        for txt in texts:
            if re.match(r'\d{2}[/\-]\d{2}[/\-]\d{4}', txt) or re.match(r'\d{4}-\d{2}-\d{2}', txt):
                data_val = txt
            elif len(txt) > 20 and not data_val == txt:
                oggetto_val = max(oggetto_val, txt, key=len)

        if oggetto_val or data_val:
            rows.append({'data': data_val, 'oggetto': oggetto_val, 'link': link_val})

    has_more = bool(soup.find('a', string=re.compile(r'[Ss]uccessiv|Next|»|›')))
    return rows, has_more

# ── Regione Calabria ───────────────────────────────────────────────────────────

def scrape_regione_calabria(session):
    """Scrape all pages of Regione Calabria provvedimenti."""
    print(f"\n{'='*60}")
    print(f"Scraping Regione Calabria — filter_date_from={DATE_FROM_2}")

    results = []
    error = None

    try:
        page = 1
        while True:
            url = (
                'https://www.regione.calabria.it/provvedimenti-della-regione/'
                f'?filter_date_from={DATE_FROM_2}&paged={page}'
            )
            r = session.get(url, timeout=30)
            r.raise_for_status()
            soup = BeautifulSoup(r.text, 'html.parser')

            table = soup.find('table')
            if not table:
                print(f"  Page {page}: no table")
                break

            trs = table.find_all('tr')
            data_trs = trs[1:]  # skip header row
            if not data_trs:
                print(f"  Page {page}: 0 data rows → stop")
                break

            for tr in data_trs:
                tds = tr.find_all('td')
                link_el = tr.find('a')
                link = link_el.get('href', '') if link_el else ''

                if len(tds) >= 4:
                    tipologia = tds[0].get_text(strip=True)
                    data_val  = tds[1].get_text(strip=True)
                    # col 2 = N (numero), col 3 = Dipartimento (may be omitted depending on cols)
                    # Columns: Tipologia | Data | N | Dipartimento | Oggetto | Dettaglio
                    if len(tds) >= 5:
                        dipartimento = tds[3].get_text(strip=True) if len(tds) > 3 else ''
                        oggetto = tds[4].get_text(strip=True) if len(tds) > 4 else ''
                    else:
                        dipartimento = ''
                        oggetto = tds[3].get_text(strip=True) if len(tds) > 3 else ''

                    results.append({
                        'data': data_val,
                        'tipologia': tipologia,
                        'dipartimento': dipartimento,
                        'oggetto': oggetto,
                        'link': link,
                    })

            print(f"  Page {page}: {len(data_trs)} rows, total={len(results)}")

            # Next page
            next_link = soup.find('a', string=re.compile(r'[Pp]agina\s+successiva|Next'))
            if not next_link:
                break
            page += 1
            if page > 100:
                break

    except Exception as e:
        error = str(e)
        print(f"  ERROR: {error}")

    return results, error

# ── TAR Catanzaro ─────────────────────────────────────────────────────────────

TAR_BASE = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'
TAR_PP_ID = 'it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe'
TAR_PREFIX = f'_{TAR_PP_ID}_'

def scrape_tar_catanzaro(session):
    """Scrape TAR Catanzaro provvedimenti."""
    print(f"\n{'='*60}")
    print(f"Scraping TAR Catanzaro — publishDateFrom={DATE_FROM_3}")

    results = []
    error = None

    try:
        # Step 1: GET initial page (obtain p_auth and formDate)
        r0 = session.get(TAR_BASE, timeout=30)
        r0.raise_for_status()
        soup0 = BeautifulSoup(r0.text, 'html.parser')

        form = soup0.find('form', {'action': re.compile('search')})
        if not form:
            return [], "Form di ricerca non trovato"

        action = form.get('action', '')
        form_date_input = soup0.find('input', {'name': TAR_PREFIX + 'formDate'})
        form_date_val = form_date_input.get('value', str(int(time.time() * 1000))) if form_date_input else str(int(time.time() * 1000))

        # Step 2: POST search
        post_data = {
            TAR_PREFIX + 'formDate': form_date_val,
            TAR_PREFIX + 'year': '',
            TAR_PREFIX + 'number': '',
            TAR_PREFIX + 'hearingDateFrom': '',
            TAR_PREFIX + 'hearingDateTo': '',
            TAR_PREFIX + 'section': '',
            TAR_PREFIX + 'type': '',
            TAR_PREFIX + 'specific': '',
            TAR_PREFIX + 'publishDateFrom': DATE_FROM_3,
            TAR_PREFIX + 'publishDateTo': '',
            TAR_PREFIX + 'president': '',
            TAR_PREFIX + 'draftingJudge': '',
        }

        headers = {
            'Referer': TAR_BASE,
            'Origin': 'https://www.giustizia-amministrativa.it',
            'Content-Type': 'application/x-www-form-urlencoded',
        }

        r2 = session.post(action, data=post_data, headers=headers, timeout=30, allow_redirects=True)
        r2.raise_for_status()
        soup2 = BeautifulSoup(r2.text, 'html.parser')

        portlet = soup2.find(id=re.compile('jjYpzZYF4Qfe'))
        if not portlet:
            return [], "Portlet TAR non trovato nella risposta"

        portlet_text = portlet.get_text().lower()
        # The backend may show "errore" alongside the empty table (header only).
        # Treat any "errore" presence as a backend failure signal and record it;
        # we still continue parsing — if we get 0 data rows the caller will
        # surface the error in the sheet.
        backend_error = 'errore' in portlet_text

        # Parse through pages
        current_soup = soup2
        page = 1
        while True:
            page_results, has_more = _parse_tar_page(current_soup, portlet)
            results.extend(page_results)
            print(f"  Page {page}: {len(page_results)} rows")

            if not has_more:
                break

            # Find next-page link inside portlet
            portlet_el = current_soup.find(id=re.compile('jjYpzZYF4Qfe'))
            next_link = portlet_el.find('a', string=re.compile(r'[Ss]uccessiv|Next|»|›')) if portlet_el else None
            if not next_link:
                break

            next_href = next_link.get('href', '')
            if not next_href:
                break

            r3 = session.get(next_href, timeout=30)
            r3.raise_for_status()
            current_soup = BeautifulSoup(r3.text, 'html.parser')
            portlet = current_soup.find(id=re.compile('jjYpzZYF4Qfe'))
            page += 1
            if page > 50:
                break

    except Exception as e:
        error = str(e)
        print(f"  ERROR: {error}")

    # If backend signalled an error AND we got 0 data rows, surface the error
    if 'backend_error' in dir() and backend_error and not results:
        error = ("Il portale TAR Catanzaro restituisce un errore di backend "
                 "(\"Si è verificato un errore nel recuperare le informazioni\"). "
                 "Il servizio potrebbe essere temporaneamente non disponibile. "
                 "Verificare manualmente: https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro")

    return results, error


def _parse_tar_page(soup, portlet_el=None):
    """Parse one page of TAR results."""
    container = portlet_el or soup
    results = []

    table = container.find('table')
    if not table:
        return results, False

    trs = table.find_all('tr')
    for tr in trs[1:]:
        tds = tr.find_all('td')
        if not tds:
            continue

        texts = [td.get_text(strip=True) for td in tds]
        link_el = tr.find('a')
        link = ''
        if link_el:
            href = link_el.get('href', '')
            if href.startswith('/'):
                link = 'https://www.giustizia-amministrativa.it' + href
            else:
                link = href

        # Columns: NRG | Sezione | Parte | Tipo Udienza | Data Udienza |
        #          Numero Provvedimento | Data pubblicazione | Tipo Provvedimento |
        #          Relatore | Presidente | Esito
        parte = texts[2] if len(texts) > 2 else ''
        data_pub = texts[6] if len(texts) > 6 else ''
        nrg = texts[0] if texts else ''

        # "Oggetto" for TAR is built from Parte + NRG + tipo
        tipo_prov = texts[7] if len(texts) > 7 else ''
        oggetto = f"[{tipo_prov}] Parte: {parte} — NRG: {nrg}"

        results.append({
            'data': data_pub,
            'nrg': nrg,
            'parte': parte,
            'tipo': tipo_prov,
            'oggetto': oggetto,
            'link': link,
        })

    has_more = bool(container.find('a', string=re.compile(r'[Ss]uccessiv|Next|»|›')))
    return results, has_more

# ── Excel Generation ──────────────────────────────────────────────────────────

HEADER_FILL = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
ALT_FILL    = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
LINK_FONT   = Font(color='0563C1', underline='single')
THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)

def _apply_header(ws, headers):
    ws.append(headers)
    for i, cell in enumerate(ws[1], 1):
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER


def _write_no_results(ws):
    ws.append(['Nessun risultato'])
    ws['A2'].font = Font(italic=True, color='888888')


def add_asp_sheet(wb, sheet_name, rows, error):
    ws = wb.create_sheet(title=sheet_name)
    ws.row_dimensions[1].height = 30

    if error:
        # Show error notice
        _apply_header(ws, ['Data', 'Oggetto', 'Descrizione breve', 'Link'])
        err_row = [
            '—',
            f'PORTALE NON RAGGIUNGIBILE — {error[:200]}',
            'Errore di connessione SSL/TLS al dominio sisr.regione.calabria.it',
            '—',
        ]
        ws.append(err_row)
        ws.cell(row=2, column=2).font = Font(color='CC0000', italic=True)
        _set_column_widths(ws)
        return

    # Filter by keywords
    matched = [r for r in rows if keyword_match(r.get('oggetto', ''))]

    _apply_header(ws, ['Data', 'Oggetto', 'Descrizione breve', 'Link'])

    if not matched:
        _write_no_results(ws)
        _set_column_widths(ws)
        return

    for i, row in enumerate(matched, start=2):
        oggetto = row.get('oggetto', '')
        ws.cell(row=i, column=1, value=row.get('data', ''))
        ws.cell(row=i, column=2, value=oggetto)
        ws.cell(row=i, column=3, value=(oggetto[:150] if oggetto else ''))
        link = row.get('link', '')
        link_cell = ws.cell(row=i, column=4, value='Apri atto')
        if link:
            link_cell.hyperlink = link
            link_cell.font = LINK_FONT
        # Alternate row color
        fill = ALT_FILL if i % 2 == 0 else PatternFill(fill_type=None)
        for col in range(1, 5):
            c = ws.cell(row=i, column=col)
            if fill.fill_type:
                c.fill = fill
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.border = THIN_BORDER

    _set_column_widths(ws)


def add_regione_sheet(wb, rows, error):
    ws = wb.create_sheet(title='Regione Calabria')
    ws.row_dimensions[1].height = 30

    _apply_header(ws, ['Data', 'Tipologia', 'Dipartimento', 'Oggetto', 'Descrizione breve', 'Link'])

    if error and not rows:
        err_row = ['—', '—', '—', f'ERRORE: {error[:200]}', '—', '—']
        ws.append(err_row)
        ws.cell(row=2, column=4).font = Font(color='CC0000', italic=True)
        _set_column_widths(ws)
        return

    # Filter by keywords
    matched = [r for r in rows if keyword_match(r.get('oggetto', ''))]

    if not matched:
        _write_no_results(ws)
        _set_column_widths(ws)
        return

    for i, row in enumerate(matched, start=2):
        oggetto = row.get('oggetto', '')
        ws.cell(row=i, column=1, value=row.get('data', ''))
        ws.cell(row=i, column=2, value=row.get('tipologia', ''))
        ws.cell(row=i, column=3, value=row.get('dipartimento', ''))
        ws.cell(row=i, column=4, value=oggetto)
        ws.cell(row=i, column=5, value=(oggetto[:150] if oggetto else ''))
        link = row.get('link', '')
        link_cell = ws.cell(row=i, column=6, value='Apri atto')
        if link:
            link_cell.hyperlink = link
            link_cell.font = LINK_FONT
        fill = ALT_FILL if i % 2 == 0 else PatternFill(fill_type=None)
        for col in range(1, 7):
            c = ws.cell(row=i, column=col)
            if fill.fill_type:
                c.fill = fill
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.border = THIN_BORDER

    _set_column_widths(ws)


def add_tar_sheet(wb, rows, error):
    ws = wb.create_sheet(title='TAR Catanzaro')
    ws.row_dimensions[1].height = 30

    _apply_header(ws, ['Data Pubbl.', 'NRG', 'Parte', 'Tipo Provvedimento', 'Oggetto composto', 'Descrizione breve', 'Link'])

    if error and not rows:
        err_row = ['—', '—', '—', '—', f'ERRORE: {error[:200]}', '—', '—']
        ws.append(err_row)
        ws.cell(row=2, column=5).font = Font(color='CC0000', italic=True)
        _set_column_widths(ws)
        return

    # For TAR, keyword filter applies to 'parte' field (many are anonymised as XXX_OMISSIS_XXX)
    matched = [r for r in rows if keyword_match(r.get('parte', '') + ' ' + r.get('oggetto', ''))]

    if not matched:
        _write_no_results(ws)
        _set_column_widths(ws)
        return

    for i, row in enumerate(matched, start=2):
        oggetto = row.get('oggetto', '')
        ws.cell(row=i, column=1, value=row.get('data', ''))
        ws.cell(row=i, column=2, value=row.get('nrg', ''))
        ws.cell(row=i, column=3, value=row.get('parte', ''))
        ws.cell(row=i, column=4, value=row.get('tipo', ''))
        ws.cell(row=i, column=5, value=oggetto)
        ws.cell(row=i, column=6, value=(oggetto[:150] if oggetto else ''))
        link = row.get('link', '')
        link_cell = ws.cell(row=i, column=7, value='Apri atto')
        if link:
            link_cell.hyperlink = link
            link_cell.font = LINK_FONT
        fill = ALT_FILL if i % 2 == 0 else PatternFill(fill_type=None)
        for col in range(1, 8):
            c = ws.cell(row=i, column=col)
            if fill.fill_type:
                c.fill = fill
            c.alignment = Alignment(wrap_text=True, vertical='top')
            c.border = THIN_BORDER

    _set_column_widths(ws)


def _set_column_widths(ws):
    col_widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                col_letter = get_column_letter(cell.column)
                content_len = min(len(str(cell.value)), 80)
                col_widths[col_letter] = max(col_widths.get(col_letter, 10), content_len + 2)
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = min(width, 60)

# ── Cover Sheet ────────────────────────────────────────────────────────────────

def add_cover_sheet(wb, run_date, asp_errors, regione_error, tar_error):
    ws = wb.create_sheet(title='Riepilogo', index=0)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 70

    title_font = Font(bold=True, size=14, color='1F4E79')
    ws['A1'] = 'MONITORAGGIO ATTI CALABRIA'
    ws['A1'].font = title_font
    ws.merge_cells('A1:B1')

    ws.append([])
    ws.append(['Data esecuzione', run_date.strftime('%d/%m/%Y %H:%M')])
    ws.append(['Periodo monitoraggio (ASP/Regione)', f"Ultimi 2 giorni: dal {DATE_FROM_2} ad oggi"])
    ws.append(['Periodo monitoraggio (TAR)', f"Ultimi 3 giorni: dal {DATE_FROM_3} ad oggi"])
    ws.append([])
    ws.append(['SORGENTI DATI', 'STATO'])

    header_row = ws.max_row
    for col in [1, 2]:
        c = ws.cell(row=header_row, column=col)
        c.fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
        c.font = Font(bold=True, color='FFFFFF')

    status_rows = [
        ('ASP Cosenza',        'ERRORE SSL/TLS — portale non raggiungibile' if asp_errors.get('ASP Cosenza') else 'OK'),
        ('ASP Catanzaro',      'ERRORE SSL/TLS — portale non raggiungibile' if asp_errors.get('ASP Catanzaro') else 'OK'),
        ('ASP Crotone',        'ERRORE SSL/TLS — portale non raggiungibile' if asp_errors.get('ASP Crotone') else 'OK'),
        ('ASP Reggio Calabria','ERRORE SSL/TLS — portale non raggiungibile' if asp_errors.get('ASP Reggio Calabria') else 'OK'),
        ('ASP Vibo Valentia',  'ERRORE SSL/TLS — portale non raggiungibile' if asp_errors.get('ASP Vibo Valentia') else 'OK'),
        ('Regione Calabria',   f'ERRORE: {regione_error}' if regione_error else 'OK'),
        ('TAR Catanzaro',      f'ERRORE: {tar_error}' if tar_error else 'OK'),
    ]

    for src, status in status_rows:
        ws.append([src, status])
        row_n = ws.max_row
        color = 'CC0000' if 'ERRORE' in status else '008000'
        ws.cell(row=row_n, column=2).font = Font(color=color)

    ws.append([])
    ws.append(['NOTA', (
        'I portali ASP (sisr.regione.calabria.it) usano TLS 1.0/1.1 e sono '
        'irraggiungibili dall\'ambiente di esecuzione remoto a causa di '
        'incompatibilità con il proxy SSL aziendale. '
        'Il portale TAR Catanzaro restituisce un errore di backend. '
        'Consultare i fogli individuali per i dettagli sui risultati disponibili.'
    )])
    ws.cell(row=ws.max_row, column=2).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[ws.max_row].height = 60

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    run_date = TODAY
    print(f"Monitoraggio Atti Calabria — {run_date.strftime('%d/%m/%Y %H:%M')}")
    print(f"Date from (ASP/Regione): {DATE_FROM_2}")
    print(f"Date from (TAR):         {DATE_FROM_3}")

    session = make_session()
    wb = Workbook()
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # ── ASP portals ──
    asp_errors = {}
    for asp_name, asp_url in ASP_PORTALS.items():
        rows, err = scrape_asp(asp_name, asp_url, session)
        if err:
            asp_errors[asp_name] = err
        add_asp_sheet(wb, asp_name, rows, err)
        matched = [r for r in rows if keyword_match(r.get('oggetto', ''))]
        print(f"  → {asp_name}: {len(matched)} match(es) after keyword filter" + (f" [ERROR: {err[:80]}]" if err else ""))

    # ── Regione Calabria ──
    reg_rows, reg_error = scrape_regione_calabria(session)
    add_regione_sheet(wb, reg_rows, reg_error)
    reg_matched = [r for r in reg_rows if keyword_match(r.get('oggetto', ''))]
    print(f"\n  → Regione Calabria: {len(reg_rows)} total, {len(reg_matched)} matched")

    # ── TAR Catanzaro ──
    tar_rows, tar_error = scrape_tar_catanzaro(session)
    add_tar_sheet(wb, tar_rows, tar_error)
    tar_matched = [r for r in tar_rows if keyword_match(r.get('parte', '') + ' ' + r.get('oggetto', ''))]
    print(f"  → TAR Catanzaro: {len(tar_rows)} total, {len(tar_matched)} matched")

    # ── Cover sheet ──
    add_cover_sheet(wb, run_date, asp_errors, reg_error, tar_error)

    # ── Save file ──
    filename = f"Monitoraggio_Atti_Calabria_{run_date.strftime('%d-%m-%Y')}.xlsx"
    output_dir = '/home/user/NS'
    os.makedirs(output_dir, exist_ok=True)
    filepath = os.path.join(output_dir, filename)
    wb.save(filepath)
    print(f"\n✓ File saved: {filepath}")
    print(f"  Sheets: {wb.sheetnames}")

    return filepath, reg_matched, tar_matched, asp_errors, reg_error, tar_error


if __name__ == '__main__':
    main()
