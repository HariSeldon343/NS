#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scarica atti da: 5 portali ASP, Regione Calabria, TAR Catanzaro
Filtra per keyword e produce file Excel.
"""

import re
import json
import time
import logging
from datetime import date, timedelta
from urllib.parse import urlencode, quote

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

# ─── Date ranges ──────────────────────────────────────────────────────────────
TODAY = date.today()
DATE_FROM_2 = TODAY - timedelta(days=2)          # ASP + Regione Calabria
DATE_FROM_3 = TODAY - timedelta(days=3)          # TAR Catanzaro

DATE_FROM_2_STR = DATE_FROM_2.strftime('%Y-%m-%d')
DATE_FROM_3_STR = DATE_FROM_3.strftime('%Y-%m-%d')
DATE_TO_STR     = TODAY.strftime('%Y-%m-%d')
TODAY_STR       = TODAY.strftime('%d-%m-%Y')

# ─── Keywords ─────────────────────────────────────────────────────────────────
KEYWORDS = [
    'ADI', 'Assistenza domiciliare', 'ANMIC', 'Accreditamento',
    'Aumento di budget', 'Autismo', 'Autorizzazione all\'esercizio',
    'Autorizzazione alla realizzazione', 'Autorizzazioni', 'Budget',
    'Casa Giardino', 'Centro San Giuseppe', 'Centro salute e benessere',
    'Fabbisogni LEA', 'Fisiolab', 'Fisioterapia',
    'Parere commissione', 'Presa d\'atto verifica', 'Programmazione',
    'Rete riabilitativa', 'Rete territoriale', 'Riabilitazione estensiva',
    'Riconversione prestazioni', 'Rinnovo accreditamento',
    'San Teodoro', 'Savelli Hospital', 'Starbene',
    'Verifica requisiti', 'Villa San Giuseppe', 'Villa del Rosario',
]
# Life: word-boundary match; all others: case-insensitive contains
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

KW_PATTERNS = [re.compile(re.escape(kw), re.IGNORECASE) for kw in KEYWORDS]


def matches_keywords(text: str) -> bool:
    if LIFE_RE.search(text):
        return True
    return any(p.search(text) for p in KW_PATTERNS)


# ─── HTTP session ─────────────────────────────────────────────────────────────
def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0.0.0 Safari/537.36'
        ),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'it-IT,it;q=0.9,en-US;q=0.8',
        'Connection': 'keep-alive',
    })
    return s


# ═══════════════════════════════════════════════════════════════════════════════
# ASP Portals (sisr.regione.calabria.it)
# ═══════════════════════════════════════════════════════════════════════════════

ASP_PORTALS = {
    'ASP Cosenza':        'aspco',
    'ASP Catanzaro':      'aspcz',
    'ASP Crotone':        'aspkr',
    'ASP Reggio Calabria':'asprc',
    'ASP Vibo Valentia':  'aspvv',
}


def scrape_asp(acronym: str, label: str):
    """
    Scrape the AlboOnline portal for one ASP.
    Returns list of dicts: {data, oggetto, link}
    """
    base_url = f'https://online-{acronym}.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'
    session = make_session()
    results = []

    try:
        # ── Step 1: GET the search page to obtain form tokens ──────────────
        log.info(f'[{label}] GET {base_url}')
        resp = session.get(base_url, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        log.warning(f'[{label}] Impossibile raggiungere il portale: {e}')
        return None  # None = portale irraggiungibile

    soup = BeautifulSoup(resp.text, 'lxml')

    # ── Detect form action and method ──────────────────────────────────────
    form = soup.find('form')
    form_action = base_url
    form_method = 'post'
    if form:
        form_action = form.get('action', base_url) or base_url
        if not form_action.startswith('http'):
            form_action = base_url.rsplit('/', 1)[0] + '/' + form_action.lstrip('/')
        form_method = (form.get('method', 'post') or 'post').lower()

    # ── Collect all hidden / default inputs ───────────────────────────────
    base_data = {}
    if form:
        for inp in form.find_all('input'):
            n, v = inp.get('name', ''), inp.get('value', '')
            if n:
                base_data[n] = v

    # ── Date field discovery ───────────────────────────────────────────────
    # Known id: dataPubblicazioneDal  (from task description)
    date_from_field = 'dataPubblicazioneDal'
    date_to_field   = 'dataPubblicazioneAl'

    # Try to find actual input names by id
    if form:
        for inp in form.find_all('input', id='dataPubblicazioneDal'):
            date_from_field = inp.get('name', date_from_field)
        for inp in form.find_all('input', id='dataPubblicazioneAl'):
            date_to_field = inp.get('name', date_to_field)

    page = 1
    while True:
        payload = dict(base_data)
        payload[date_from_field] = DATE_FROM_2_STR
        payload[date_to_field]   = DATE_TO_STR
        # Common pagination parameter names
        payload['page']     = str(page)
        payload['pagina']   = str(page)
        payload['pageNum']  = str(page)

        try:
            if form_method == 'post':
                r = session.post(form_action, data=payload, timeout=30)
            else:
                r = session.get(form_action, params=payload, timeout=30)
            r.raise_for_status()
        except Exception as e:
            log.warning(f'[{label}] Errore pagina {page}: {e}')
            break

        soup_p = BeautifulSoup(r.text, 'lxml')

        # ── Extract rows ───────────────────────────────────────────────
        new_rows = _parse_asp_rows(soup_p, f'https://online-{acronym}.sisr.regione.calabria.it')
        if not new_rows:
            break
        results.extend(new_rows)

        # ── Check for next page ────────────────────────────────────────
        has_next = _asp_has_next_page(soup_p, page)
        if not has_next:
            break
        page += 1
        time.sleep(0.5)

    return results


def _parse_asp_rows(soup, base_url: str):
    """Parse result rows from an ASP albo page."""
    rows = []
    table = soup.find('table')
    if not table:
        return rows

    tbody = table.find('tbody') or table
    for tr in tbody.find_all('tr'):
        tds = tr.find_all(['td', 'th'])
        if len(tds) < 2:
            continue
        cells = [td.get_text(separator=' ', strip=True) for td in tds]

        # Try to identify data/oggetto columns heuristically
        data_val  = ''
        oggetto   = ''
        link_href = ''

        for i, td in enumerate(tds):
            txt = cells[i]
            # Date pattern DD/MM/YYYY or YYYY-MM-DD
            if re.match(r'\d{2}/\d{2}/\d{4}', txt) or re.match(r'\d{4}-\d{2}-\d{2}', txt):
                if not data_val:
                    data_val = txt
            # Link
            a = td.find('a')
            if a and a.get('href'):
                href = a['href']
                if not href.startswith('http'):
                    href = base_url + href
                link_href = href

        # Oggetto: longest cell
        if cells:
            oggetto = max(cells, key=len)

        if not data_val and not oggetto:
            continue

        rows.append({'data': data_val, 'oggetto': oggetto, 'link': link_href})
    return rows


def _asp_has_next_page(soup, current_page: int) -> bool:
    """Check whether there is a next page."""
    # Look for "next" links
    for a in soup.find_all('a'):
        txt = a.get_text(strip=True).lower()
        href = a.get('href', '')
        if any(x in txt for x in ['successiv', 'next', '>>', '›', str(current_page + 1)]):
            return True
        if f'page={current_page + 1}' in href or f'pagina={current_page + 1}' in href:
            return True
    return False


# ═══════════════════════════════════════════════════════════════════════════════
# Regione Calabria
# ═══════════════════════════════════════════════════════════════════════════════

def scrape_regione_calabria():
    """Scrape www.regione.calabria.it/provvedimenti-della-regione/"""
    base_url = 'https://www.regione.calabria.it/provvedimenti-della-regione/'
    session = make_session()
    results = []
    page = 1

    while True:
        params = {
            'filter_date_from': DATE_FROM_2_STR,
            'filter_date_to':   DATE_TO_STR,
            'filter_active':    'true',
        }
        if page > 1:
            params['paged'] = str(page)

        log.info(f'[Regione Calabria] Pagina {page}')
        try:
            resp = session.get(base_url, params=params, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            log.warning(f'[Regione Calabria] Errore pagina {page}: {e}')
            break

        soup = BeautifulSoup(resp.text, 'lxml')
        table = soup.find('table')
        if not table:
            break

        tbody = table.find('tbody')
        if not tbody:
            break

        rows = tbody.find_all('tr')
        if not rows:
            break

        for tr in rows:
            tds = tr.find_all(['td', 'th'])
            if len(tds) < 5:
                continue
            tipologia  = tds[0].get_text(strip=True)
            data_str   = tds[1].get_text(strip=True)
            numero     = tds[2].get_text(strip=True)
            dip        = tds[3].get_text(strip=True)
            oggetto    = tds[4].get_text(separator=' ', strip=True)
            link_td    = tds[5] if len(tds) > 5 else None
            link_href  = ''
            if link_td:
                a = link_td.find('a')
                if a:
                    link_href = a.get('href', '')

            results.append({
                'data':    data_str,
                'oggetto': oggetto,
                'link':    link_href,
            })

        # Check pagination
        pag = soup.find('ul', class_='page-numbers')
        has_next = False
        if pag:
            for a in pag.find_all('a'):
                href = a.get('href', '')
                txt  = a.get_text(strip=True).lower()
                if 'next' in txt or 'successiv' in txt or f'paged={page + 1}' in href:
                    has_next = True
                    break
        if not has_next:
            break
        page += 1
        time.sleep(0.5)

    return results


# ═══════════════════════════════════════════════════════════════════════════════
# TAR Catanzaro
# ═══════════════════════════════════════════════════════════════════════════════

PORTLET_PREFIX = (
    '_it_indra_ga_institutional_area_'
    'JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_'
)

def scrape_tar_catanzaro():
    """Scrape www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"""
    page_url = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'
    session = make_session()

    # ── Step 1: GET page to capture p_auth and formDate ───────────────────
    log.info('[TAR Catanzaro] GET form page')
    try:
        resp = session.get(page_url, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        log.warning(f'[TAR Catanzaro] Impossibile raggiungere il portale: {e}')
        return None

    soup = BeautifulSoup(resp.text, 'lxml')
    form = soup.find('form', action=re.compile('jjYpzZYF4Qfe'))
    if not form:
        log.warning('[TAR Catanzaro] Form non trovato')
        return None

    action = form.get('action', '')
    form_data = {
        inp.get('name', ''): inp.get('value', '')
        for inp in form.find_all('input') if inp.get('name', '')
    }
    for sel in form.find_all('select'):
        n = sel.get('name', '')
        if n:
            form_data[n] = ''

    form_data[f'{PORTLET_PREFIX}publishDateFrom'] = DATE_FROM_3_STR
    form_data[f'{PORTLET_PREFIX}publishDateTo']   = DATE_TO_STR

    # ── Step 2: POST the search form to set session state ─────────────────
    log.info('[TAR Catanzaro] POST search form')
    try:
        post_resp = session.post(action, data=form_data, timeout=30, allow_redirects=True)
        post_resp.raise_for_status()
    except Exception as e:
        log.warning(f'[TAR Catanzaro] Errore POST: {e}')
        return None

    # ── Step 3: Extract AJAX searchURL ────────────────────────────────────
    m = re.search(r'searchURL\s*=\s*"(https://[^"]+)"', post_resp.text)
    if not m:
        log.warning('[TAR Catanzaro] searchURL non trovato')
        return None
    search_url = m.group(1)

    # ── Step 4: Build additionalInfo ──────────────────────────────────────
    additional_info = json.dumps({
        'schema':               'TAR_CATANZARO',
        'type':                 None,
        'year':                 '',
        'number':               '',
        'hearingDateFrom':      None,
        'hearingDateTo':        None,
        'publishDateFrom':      DATE_FROM_3_STR,
        'publishDateTo':        DATE_TO_STR,
        'hearingType':          None,
        'nrg':                  None,
        'section':              '',
        'provisionSpecification': '',
        'president':            '',
        'draftingJudge':        '',
        'subjectMatter':        None,
        'page':                 None,
        'size':                 None,
        'orderBy':              None,
        'orderStrategy':        None,
        'queryString':          None,
    })

    # ── Step 5: Paginate through DataTables AJAX ──────────────────────────
    results = []
    start = 0
    page_size = 100

    while True:
        dt_payload = {
            'draw':         start // page_size + 1,
            'columns': [
                {'data': col, 'name': '', 'searchable': True, 'orderable': True,
                 'search': {'value': '', 'regex': False}}
                for col in ['nrgFascicolo', 'sezione', 'parte', 'tipoUdienza',
                            'dataUdienza', 'numProvvedimento', 'dataPubblicazione',
                            'tipoProvvedimento', 'relatore', 'presidente', 'esito']
            ],
            'order':  [{'column': 6, 'dir': 'desc'}, {'column': 5, 'dir': 'desc'}],
            'start':  start,
            'length': page_size,
            'search': {'value': '', 'regex': False},
            'additionalInfo': additional_info,
        }

        log.info(f'[TAR Catanzaro] AJAX start={start}')
        try:
            ajax_resp = session.post(
                search_url,
                data=json.dumps(dt_payload),
                timeout=30,
                headers={
                    'Content-Type':    'application/json',
                    'X-Requested-With': 'XMLHttpRequest',
                    'Accept':           'application/json, text/javascript, */*; q=0.01',
                    'Referer':          post_resp.url,
                },
            )
            ajax_resp.raise_for_status()
        except Exception as e:
            log.warning(f'[TAR Catanzaro] Errore AJAX: {e}')
            break

        try:
            data = ajax_resp.json()
        except Exception as e:
            log.warning(f'[TAR Catanzaro] JSON parse error: {e}')
            break

        records = data.get('data', [])
        if not records:
            break

        for rec in records:
            nrg       = rec.get('nrgFascicolo', '')
            nome_file = rec.get('nomeFile', '')
            parte     = rec.get('parte', '')
            data_pub  = rec.get('dataPubblicazione', '')
            tipo      = rec.get('tipoProvvedimento', '')
            num_prov  = rec.get('numProvvedimento', '')

            link = (
                'https://mdp.giustizia-amministrativa.it/visualizza/'
                f'?nodeRef=&schema=tar_cz&nrg={nrg}&nomeFile={nome_file}'
                '&subDir=Provvedimenti'
            ) if nome_file else ''

            oggetto = f'{tipo} n.{num_prov} – {parte}'.strip(' –')

            results.append({
                'data':    data_pub,
                'oggetto': oggetto,
                'parte':   parte,
                'link':    link,
            })

        total = data.get('recordsFiltered', data.get('recordsTotal', 0))
        start += page_size
        if start >= total:
            break
        time.sleep(0.3)

    return results


# ═══════════════════════════════════════════════════════════════════════════════
# Excel writer
# ═══════════════════════════════════════════════════════════════════════════════

HEADER_FILL   = PatternFill('solid', fgColor='1F497D')
HEADER_FONT   = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
LINK_FONT     = Font(name='Calibri', color='0563C1', underline='single', size=10)
NORMAL_FONT   = Font(name='Calibri', size=10)
ALT_FILL      = PatternFill('solid', fgColor='DCE6F1')
THIN_BORDER   = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

COLUMNS = ['Data', 'Oggetto', 'Descrizione breve', 'Link']
COL_WIDTHS = [14, 60, 40, 15]


def write_sheet(ws, rows_data: list, sheet_label: str):
    """Write one portal's results to a worksheet."""
    # Header row
    for col_idx, (header, width) in enumerate(zip(COLUMNS, COL_WIDTHS), start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border    = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'

    if not rows_data:
        ws.cell(row=2, column=1, value='Nessun risultato').font = Font(italic=True, color='888888')
        return

    for row_idx, item in enumerate(rows_data, start=2):
        fill = ALT_FILL if row_idx % 2 == 0 else None

        # Col A – Data
        c_data = ws.cell(row=row_idx, column=1, value=item.get('data', ''))
        c_data.font      = NORMAL_FONT
        c_data.alignment = Alignment(horizontal='center', vertical='top')
        c_data.border    = THIN_BORDER
        if fill:
            c_data.fill = fill

        # Col B – Oggetto
        oggetto = item.get('oggetto', '')
        c_obj = ws.cell(row=row_idx, column=2, value=oggetto)
        c_obj.font      = NORMAL_FONT
        c_obj.alignment = Alignment(vertical='top', wrap_text=True)
        c_obj.border    = THIN_BORDER
        if fill:
            c_obj.fill = fill

        # Col C – Descrizione breve
        c_desc = ws.cell(row=row_idx, column=3, value=oggetto[:150])
        c_desc.font      = NORMAL_FONT
        c_desc.alignment = Alignment(vertical='top', wrap_text=True)
        c_desc.border    = THIN_BORDER
        if fill:
            c_desc.fill = fill

        # Col D – Link cliccabile
        link = item.get('link', '')
        c_link = ws.cell(row=row_idx, column=4, value='Apri' if link else '')
        if link:
            c_link.hyperlink = link
            c_link.font      = LINK_FONT
        else:
            c_link.font = NORMAL_FONT
        c_link.alignment = Alignment(horizontal='center', vertical='top')
        c_link.border    = THIN_BORDER
        if fill:
            c_link.fill = fill

        ws.row_dimensions[row_idx].height = max(15, min(60, len(oggetto) // 3))


# ═══════════════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    # ── 1. ASP Portals ────────────────────────────────────────────────────
    for label, acronym in ASP_PORTALS.items():
        log.info(f'═══ {label} ═══')
        raw = scrape_asp(acronym, label)

        if raw is None:
            # portal unreachable
            ws = wb.create_sheet(title=label)
            for col_idx, (header, width) in enumerate(zip(COLUMNS, COL_WIDTHS), start=1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = HEADER_FONT; cell.fill = HEADER_FILL
                cell.alignment = Alignment(horizontal='center')
                cell.border = THIN_BORDER
                ws.column_dimensions[get_column_letter(col_idx)].width = width
            msg = ws.cell(row=2, column=1, value=f'Portale non raggiungibile dall\'ambiente remoto (HTTP 503)')
            msg.font = Font(italic=True, color='FF0000')
            ws.merge_cells('A2:D2')
            log.warning(f'[{label}] Portale non raggiungibile')
            continue

        # Filter by keyword
        filtered = [r for r in raw if matches_keywords(r.get('oggetto', ''))]
        log.info(f'[{label}] {len(raw)} atti totali → {len(filtered)} match keyword')

        ws = wb.create_sheet(title=label)
        write_sheet(ws, filtered, label)

    # ── 2. Regione Calabria ───────────────────────────────────────────────
    log.info('═══ Regione Calabria ═══')
    rc_raw = scrape_regione_calabria()
    if rc_raw is None:
        rc_filtered = []
        rc_error = True
    else:
        rc_filtered = [r for r in rc_raw if matches_keywords(r.get('oggetto', ''))]
        rc_error = False
        log.info(f'[Regione Calabria] {len(rc_raw)} atti totali → {len(rc_filtered)} match')

    ws_rc = wb.create_sheet(title='Regione Calabria')
    if rc_error:
        ws_rc.cell(row=1, column=1, value='Portale non raggiungibile').font = Font(italic=True, color='FF0000')
    else:
        write_sheet(ws_rc, rc_filtered, 'Regione Calabria')

    # ── 3. TAR Catanzaro ──────────────────────────────────────────────────
    log.info('═══ TAR Catanzaro ═══')
    tar_raw = scrape_tar_catanzaro()
    if tar_raw is None:
        tar_filtered = []
        tar_error = True
    else:
        # For TAR filter on 'parte' (the party name column) AND the oggetto (type + n. + parte)
        tar_filtered = [
            r for r in tar_raw
            if matches_keywords(r.get('parte', '')) or matches_keywords(r.get('oggetto', ''))
        ]
        tar_error = False
        log.info(f'[TAR Catanzaro] {len(tar_raw)} atti totali → {len(tar_filtered)} match')

    ws_tar = wb.create_sheet(title='TAR Catanzaro')
    if tar_error:
        ws_tar.cell(row=1, column=1, value='Portale non raggiungibile').font = Font(italic=True, color='FF0000')
    else:
        write_sheet(ws_tar, tar_filtered, 'TAR Catanzaro')

    # ── Save ──────────────────────────────────────────────────────────────
    filename = f'Monitoraggio_Atti_Calabria_{TODAY_STR}.xlsx'
    filepath = f'/home/user/NS/{filename}'
    wb.save(filepath)
    log.info(f'File salvato: {filepath}')
    print(f'\n✓ File: {filepath}')
    print(f'  Data esecuzione: {TODAY_STR}')

    # Summary
    print('\n── Riepilogo risultati ──')
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        nrows = ws.max_row - 1
        first_cell = ws.cell(row=2, column=1).value or ''
        if 'non raggiungibile' in str(first_cell):
            print(f'  {sheet:30s}: PORTALE NON RAGGIUNGIBILE')
        elif 'Nessun risultato' in str(first_cell):
            print(f'  {sheet:30s}: nessun match')
        else:
            print(f'  {sheet:30s}: {nrows} righe')


if __name__ == '__main__':
    main()
