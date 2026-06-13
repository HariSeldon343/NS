#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scrapes ASP portals, Regione Calabria, and TAR Catanzaro
and produces a filtered Excel report.
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
from datetime import datetime, timedelta
import time
import json
import sys
import os
import urllib.parse

# ─── Configuration ──────────────────────────────────────────────────────────

TODAY = datetime(2026, 5, 30)
DATE_FROM_2DAYS = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_FROM_3DAYS = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')
OUTPUT_DATE = TODAY.strftime('%d-%m-%Y')
OUTPUT_FILENAME = f'Monitoraggio_Atti_Calabria_{OUTPUT_DATE}.xlsx'

print(f"Data esecuzione : {TODAY.strftime('%d/%m/%Y')}")
print(f"Data dal (2gg)  : {DATE_FROM_2DAYS}")
print(f"Data dal (3gg)  : {DATE_FROM_3DAYS}")
print(f"File output     : {OUTPUT_FILENAME}")

# ─── Keywords ────────────────────────────────────────────────────────────────

KEYWORDS = [
    'ADI', 'Assistenza domiciliare', 'ANMIC', 'Accreditamento',
    'Aumento di budget', 'Autismo', "Autorizzazione all'esercizio",
    'Autorizzazione alla realizzazione', 'Autorizzazioni', 'Budget',
    'Casa Giardino', 'Centro San Giuseppe', 'Centro salute e benessere',
    'Fabbisogni LEA', 'Fisiolab', 'Fisioterapia', 'Parere commissione',
    "Presa d'atto verifica", 'Programmazione', 'Rete riabilitativa',
    'Rete territoriale', 'Riabilitazione estensiva', 'Riconversione prestazioni',
    'Rinnovo accreditamento', 'San Teodoro', 'Savelli Hospital', 'Starbene',
    'Verifica requisiti', 'Villa San Giuseppe', 'Villa del Rosario'
]

def matches_keywords(text):
    if not text:
        return False
    for kw in KEYWORDS:
        if kw.upper() in text.upper():
            return True
    # Word-boundary match for "Life"
    if re.search(r'\bLIFE\b', text, re.IGNORECASE):
        return True
    return False

# ─── HTTP Session ────────────────────────────────────────────────────────────

def make_session():
    s = requests.Session()
    s.headers.update({
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/124.0.0.0 Safari/537.36'
        ),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'it-IT,it;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
    })
    return s

# ─── ASP Albo Online ─────────────────────────────────────────────────────────

ASP_PORTALS = [
    ('ASP Cosenza',        'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Catanzaro',      'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Crotone',        'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Reggio Calabria','https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Vibo Valentia',  'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
]

def extract_hidden_fields(soup):
    """Extract all hidden form fields."""
    fields = {}
    for inp in soup.find_all('input', type='hidden'):
        name = inp.get('name')
        value = inp.get('value', '')
        if name:
            fields[name] = value
    return fields

def extract_form_id(soup):
    form = soup.find('form')
    return form.get('id', 'form') if form else 'form'

def parse_asp_table(soup, base_url):
    """Parse result table rows from ASP Albo Online page."""
    rows = []
    table = soup.find('table', id=lambda x: x and 'risultati' in x.lower())
    if not table:
        # Try any table with data rows
        tables = soup.find_all('table')
        for t in tables:
            trs = t.find_all('tr')
            if len(trs) > 1:
                # Check if it looks like the results table
                headers = [th.get_text(strip=True).lower() for th in trs[0].find_all(['th', 'td'])]
                if any(h in headers for h in ['oggetto', 'data', 'tipo', 'numero']):
                    table = t
                    break

    if not table:
        return rows

    trs = table.find_all('tr')
    if not trs:
        return rows

    # Determine column indices from header
    header_row = trs[0]
    headers = [th.get_text(strip=True).lower() for th in header_row.find_all(['th', 'td'])]

    col_data = col_oggetto = col_link = -1
    for i, h in enumerate(headers):
        if 'data' in h and 'pubbl' in h:
            col_data = i
        elif 'data' in h:
            col_data = i
        if 'oggetto' in h:
            col_oggetto = i
        if col_data == -1 and i == 0:
            col_data = 0

    for tr in trs[1:]:
        cells = tr.find_all(['td', 'th'])
        if not cells:
            continue
        # Flatten cell texts
        texts = [c.get_text(strip=True) for c in cells]
        if not texts:
            continue

        data_val = ''
        oggetto_val = ''
        link_val = ''

        if col_data >= 0 and col_data < len(texts):
            data_val = texts[col_data]
        if col_oggetto >= 0 and col_oggetto < len(texts):
            oggetto_val = texts[col_oggetto]
        else:
            # Fallback: use the longest cell text
            oggetto_val = max(texts, key=len) if texts else ''

        # Look for link (anchor tag)
        for cell in cells:
            a = cell.find('a', href=True)
            if a:
                href = a['href']
                if href.startswith('http'):
                    link_val = href
                else:
                    link_val = urllib.parse.urljoin(base_url, href)
                break

        if not data_val and not oggetto_val:
            continue

        rows.append({
            'data': data_val,
            'oggetto': oggetto_val,
            'link': link_val,
        })

    return rows

def has_next_page_asp(soup, form_id):
    """Check if there is a 'next page' button/link."""
    # Look for pagination controls
    pagination = soup.find(id=lambda x: x and ('paginat' in x.lower() or 'paginator' in x.lower()))
    if pagination:
        # Look for a next button that is not disabled
        nxt = pagination.find('a', class_=lambda x: x and 'next' in x.lower())
        if nxt and 'ui-state-disabled' not in (nxt.get('class') or []):
            return True
        nxt = pagination.find('span', class_=lambda x: x and 'next' in x.lower())
        if nxt and 'ui-state-disabled' not in (nxt.get('class') or []):
            return True
    return False

def get_next_page_source_asp(soup, form_id):
    """Return the j_idt source for 'next page' button if available."""
    # PrimeFaces paginator: look for aria-label="Next Page" or title="Next Page"
    for tag in soup.find_all(['a', 'span'], attrs={'aria-label': re.compile(r'next|succe', re.I)}):
        tid = tag.get('id')
        if tid:
            return tid
    # Try onclick attributes
    for tag in soup.find_all(attrs={'onclick': re.compile(r'next|succe', re.I)}):
        tid = tag.get('id')
        if tid:
            return tid
    return None

def scrape_asp_portal(portal_name, base_url, date_from):
    """Scrape a single ASP Albo Online portal."""
    print(f"\n{'='*60}")
    print(f"  {portal_name}")
    print(f"  URL: {base_url}")
    print(f"  Dal: {date_from}")
    print(f"{'='*60}")

    session = make_session()
    all_results = []

    try:
        # ── Step 1: GET the page ──────────────────────────────────────────
        resp = session.get(base_url, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'lxml')

        hidden = extract_hidden_fields(soup)
        form_id = extract_form_id(soup)
        print(f"  Form ID: {form_id}")
        print(f"  ViewState present: {'javax.faces.ViewState' in hidden}")

        # ── Step 2: Submit search form ────────────────────────────────────
        # Build POST payload for PrimeFaces/JSF search
        # We set dataPubblicazioneDal and submit
        search_source = f'{form_id}:ricerca'
        # Fallback: look for any submit button
        submit_btn = soup.find('input', type='submit') or soup.find('button', type='submit')
        if submit_btn:
            btn_id = submit_btn.get('id', '')
            btn_name = submit_btn.get('name', btn_id)
            search_source = btn_id or search_source

        # Try to find the date field
        date_field = soup.find('input', id='dataPubblicazioneDal')
        if not date_field:
            date_field = soup.find('input', id=lambda x: x and 'dataPubbl' in x)
        if date_field:
            date_field_name = date_field.get('name', date_field.get('id', f'{form_id}:dataPubblicazioneDal_input'))
            print(f"  Campo data trovato: {date_field_name}")
        else:
            date_field_name = f'{form_id}:dataPubblicazioneDal_input'
            print(f"  Campo data non trovato, uso default: {date_field_name}")

        post_data = dict(hidden)
        post_data.update({
            'javax.faces.partial.ajax': 'true',
            'javax.faces.source': search_source,
            'javax.faces.partial.execute': '@all',
            'javax.faces.partial.render': '@all',
            search_source: search_source,
            date_field_name: date_from,
        })

        # Also try non-AJAX submit
        headers_ajax = dict(session.headers)
        headers_ajax['Faces-Request'] = 'partial/ajax'
        headers_ajax['X-Requested-With'] = 'XMLHttpRequest'
        headers_ajax['Content-Type'] = 'application/x-www-form-urlencoded; charset=UTF-8'

        resp2 = session.post(base_url, data=post_data, headers=headers_ajax, timeout=30)
        resp2.raise_for_status()

        # The response might be XML (partial response) or HTML
        content_type = resp2.headers.get('Content-Type', '')
        if 'xml' in content_type or resp2.text.strip().startswith('<?xml'):
            # Parse partial response
            try:
                xml_soup = BeautifulSoup(resp2.text, 'lxml-xml')
                updates = xml_soup.find_all('update')
                combined_html = ''
                for upd in updates:
                    combined_html += upd.get_text()
                search_soup = BeautifulSoup(combined_html, 'lxml')
            except Exception:
                search_soup = BeautifulSoup(resp2.text, 'lxml')
        else:
            search_soup = BeautifulSoup(resp2.text, 'lxml')

        # ── Step 3: Try fallback — non-AJAX POST ─────────────────────────
        # If no table found, try a plain form submit
        test_rows = parse_asp_table(search_soup, base_url)
        if not test_rows:
            print("  Nessuna riga da AJAX, provo submit normale...")
            post_data2 = dict(hidden)
            post_data2[date_field_name] = date_from
            # Also set the display value
            post_data2[f'{form_id}:dataPubblicazioneDal'] = date_from

            resp3 = session.post(base_url, data=post_data2, timeout=30)
            resp3.raise_for_status()
            search_soup = BeautifulSoup(resp3.text, 'lxml')
            test_rows = parse_asp_table(search_soup, base_url)

        # ── Step 4: Try direct GET with params ───────────────────────────
        if not test_rows:
            print("  Provo GET con parametri query string...")
            params = {
                'dataPubblicazioneDal': date_from,
                'dataPubblicazioneAl': TODAY.strftime('%Y-%m-%d'),
            }
            resp4 = session.get(base_url, params=params, timeout=30)
            resp4.raise_for_status()
            search_soup = BeautifulSoup(resp4.text, 'lxml')
            test_rows = parse_asp_table(search_soup, base_url)

        # ── Step 5: Collect results & paginate ────────────────────────────
        current_soup = search_soup
        page = 1
        while True:
            rows = parse_asp_table(current_soup, base_url)
            print(f"  Pagina {page}: {len(rows)} righe trovate")
            all_results.extend(rows)

            if not has_next_page_asp(current_soup, form_id):
                break

            # Try to go to next page
            next_src = get_next_page_source_asp(current_soup, form_id)
            if not next_src:
                break

            hidden2 = extract_hidden_fields(current_soup)
            post_next = dict(hidden2)
            post_next.update({
                'javax.faces.partial.ajax': 'true',
                'javax.faces.source': next_src,
                'javax.faces.partial.execute': '@all',
                'javax.faces.partial.render': '@all',
                next_src: next_src,
            })

            resp_next = session.post(base_url, data=post_next, headers=headers_ajax, timeout=30)
            resp_next.raise_for_status()

            if 'xml' in resp_next.headers.get('Content-Type', '') or resp_next.text.strip().startswith('<?xml'):
                xml_s = BeautifulSoup(resp_next.text, 'lxml-xml')
                updates = xml_s.find_all('update')
                combined = ''
                for upd in updates:
                    combined += upd.get_text()
                current_soup = BeautifulSoup(combined, 'lxml')
            else:
                current_soup = BeautifulSoup(resp_next.text, 'lxml')

            page += 1
            if page > 50:  # safety cap
                break

    except requests.exceptions.SSLError as e:
        print(f"  SSL Error: {e} — provo senza verifica SSL...")
        try:
            session.verify = False
            import urllib3
            urllib3.disable_warnings()
            resp = session.get(base_url, timeout=30)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, 'lxml')
            rows = parse_asp_table(soup, base_url)
            all_results.extend(rows)
            print(f"  {len(rows)} righe trovate (senza SSL)")
        except Exception as e2:
            print(f"  Errore fallback: {e2}")

    except Exception as e:
        print(f"  Errore: {type(e).__name__}: {e}")

    # Filter by keywords
    filtered = [r for r in all_results if matches_keywords(r.get('oggetto', ''))]
    print(f"  Totale: {len(all_results)} atti, {len(filtered)} corrispondenti ai filtri")
    return filtered

# ─── Regione Calabria ─────────────────────────────────────────────────────────

def scrape_regione_calabria(date_from):
    """Scrape Regione Calabria provvedimenti."""
    base_url = 'https://www.regione.calabria.it/provvedimenti-della-regione/'
    print(f"\n{'='*60}")
    print(f"  Regione Calabria")
    print(f"  URL: {base_url}")
    print(f"  Dal: {date_from}")
    print(f"{'='*60}")

    session = make_session()
    all_results = []
    page = 1

    try:
        while True:
            params = {
                'filter_date_from': date_from,
                'filter_date_to': TODAY.strftime('%Y-%m-%d'),
                'paged': page,
            }
            resp = session.get(base_url, params=params, timeout=30)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, 'lxml')

            rows = parse_regione_table(soup, base_url)
            print(f"  Pagina {page}: {len(rows)} righe")
            all_results.extend(rows)

            if not rows:
                break

            # Look for next page link
            next_link = find_next_page_link(soup, page, base_url, params)
            if not next_link:
                break
            page += 1
            if page > 100:
                break

    except Exception as e:
        print(f"  Errore: {type(e).__name__}: {e}")

    filtered = [r for r in all_results if matches_keywords(r.get('oggetto', ''))]
    print(f"  Totale: {len(all_results)} atti, {len(filtered)} corrispondenti")
    return filtered

def parse_regione_table(soup, base_url):
    """Parse Regione Calabria results."""
    rows = []

    # Try common WordPress/CMS table or list structures
    # Look for table
    table = soup.find('table')
    if table:
        trs = table.find_all('tr')
        if trs:
            header_cells = [th.get_text(strip=True).lower() for th in trs[0].find_all(['th', 'td'])]
            col_data = col_oggetto = -1
            for i, h in enumerate(header_cells):
                if 'data' in h:
                    col_data = i
                if 'oggetto' in h or 'titolo' in h or 'descrizione' in h:
                    col_oggetto = i

            for tr in trs[1:]:
                cells = tr.find_all(['td', 'th'])
                if not cells:
                    continue
                texts = [c.get_text(strip=True) for c in cells]
                data_val = texts[col_data] if col_data >= 0 and col_data < len(texts) else ''
                oggetto_val = texts[col_oggetto] if col_oggetto >= 0 and col_oggetto < len(texts) else max(texts, key=len) if texts else ''
                link_val = ''
                for cell in cells:
                    a = cell.find('a', href=True)
                    if a:
                        href = a['href']
                        link_val = href if href.startswith('http') else urllib.parse.urljoin(base_url, href)
                        break
                if oggetto_val:
                    rows.append({'data': data_val, 'oggetto': oggetto_val, 'link': link_val})
        return rows

    # Try list items / article tags
    articles = soup.find_all('article') or soup.find_all('li', class_=re.compile(r'post|item|row'))
    for art in articles:
        title_tag = (art.find(['h2', 'h3', 'h4', 'h1']) or
                     art.find(class_=re.compile(r'title|oggetto|titolo')))
        oggetto_val = title_tag.get_text(strip=True) if title_tag else art.get_text(strip=True)[:200]

        date_tag = art.find(class_=re.compile(r'date|data|time')) or art.find('time')
        data_val = date_tag.get_text(strip=True) if date_tag else ''

        a_tag = art.find('a', href=True)
        link_val = a_tag['href'] if a_tag else ''
        if link_val and not link_val.startswith('http'):
            link_val = urllib.parse.urljoin(base_url, link_val)

        if oggetto_val:
            rows.append({'data': data_val, 'oggetto': oggetto_val, 'link': link_val})

    return rows

def find_next_page_link(soup, current_page, base_url, current_params):
    """Check if there's a next page."""
    next_link = soup.find('a', rel='next')
    if next_link:
        href = next_link.get('href', '')
        return href if href else None

    # PrimeFaces paginator or numbered pages
    pag = soup.find(class_=re.compile(r'pagination|paginat|nav-links'))
    if pag:
        current_a = pag.find('a', class_=re.compile(r'current|active'))
        if current_a:
            next_sib = current_a.find_next_sibling('a')
            if next_sib:
                return next_sib.get('href')

    # WordPress: check if page has content (non-empty)
    return None

# ─── TAR Catanzaro ────────────────────────────────────────────────────────────

def scrape_tar_catanzaro(date_from):
    """Scrape TAR Catanzaro provvedimenti."""
    base_url = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'
    print(f"\n{'='*60}")
    print(f"  TAR Catanzaro")
    print(f"  URL: {base_url}")
    print(f"  Dal: {date_from}")
    print(f"{'='*60}")

    session = make_session()
    all_results = []

    # TAR uses a specific API endpoint
    api_variants = [
        f'{base_url}?publishDateFrom={date_from}&publishDateTo={TODAY.strftime("%Y-%m-%d")}',
        f'{base_url}?start=0&publishDateFrom={date_from}',
    ]

    try:
        # First try the main page with query params
        resp = session.get(base_url, params={
            'publishDateFrom': date_from,
            'publishDateTo': TODAY.strftime('%Y-%m-%d'),
        }, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'lxml')

        rows = parse_tar_page(soup, base_url)
        print(f"  Pagina 1: {len(rows)} righe")
        all_results.extend(rows)

        # Paginate
        page = 2
        while True:
            resp_p = session.get(base_url, params={
                'publishDateFrom': date_from,
                'publishDateTo': TODAY.strftime('%Y-%m-%d'),
                'p_p_col_count': '1',
                'start': (page - 1) * 10,
                'page': page,
            }, timeout=30)
            resp_p.raise_for_status()
            soup_p = BeautifulSoup(resp_p.text, 'lxml')
            rows_p = parse_tar_page(soup_p, base_url)
            print(f"  Pagina {page}: {len(rows_p)} righe")
            if not rows_p:
                break
            all_results.extend(rows_p)
            page += 1
            if page > 50:
                break

    except Exception as e:
        print(f"  Errore: {type(e).__name__}: {e}")

    # Filter by keywords on "Parte" field
    filtered = [r for r in all_results if matches_keywords(r.get('oggetto', ''))]
    print(f"  Totale: {len(all_results)} atti, {len(filtered)} corrispondenti")
    return filtered

def parse_tar_page(soup, base_url):
    """Parse TAR page results."""
    rows = []

    # TAR Catanzaro typically shows a table or list
    table = soup.find('table')
    if table:
        trs = table.find_all('tr')
        if len(trs) > 1:
            headers = [th.get_text(strip=True).lower() for th in trs[0].find_all(['th', 'td'])]
            col_data = col_parte = col_num = -1
            for i, h in enumerate(headers):
                if 'data' in h:
                    col_data = i
                if 'parte' in h or 'ricorrente' in h or 'controinteressato' in h:
                    col_parte = i
                if 'numero' in h or 'n.' in h:
                    col_num = i

            for tr in trs[1:]:
                cells = tr.find_all(['td', 'th'])
                if not cells:
                    continue
                texts = [c.get_text(strip=True) for c in cells]
                if not any(texts):
                    continue

                data_val = texts[col_data] if col_data >= 0 and col_data < len(texts) else ''
                parte_val = texts[col_parte] if col_parte >= 0 and col_parte < len(texts) else ' '.join(texts)
                num_val = texts[col_num] if col_num >= 0 and col_num < len(texts) else ''

                oggetto_combined = f"{num_val} - {parte_val}".strip(' -')

                link_val = ''
                for cell in cells:
                    a = cell.find('a', href=True)
                    if a:
                        href = a['href']
                        link_val = href if href.startswith('http') else urllib.parse.urljoin(base_url, href)
                        break

                if parte_val or oggetto_combined:
                    rows.append({
                        'data': data_val,
                        'oggetto': oggetto_combined,
                        'link': link_val,
                    })
        return rows

    # List/article structure
    items = (soup.find_all('article') or
             soup.find_all('li', class_=re.compile(r'post|item|provvedimento', re.I)) or
             soup.find_all('div', class_=re.compile(r'provvedimento|row|item', re.I)))

    for item in items:
        text = item.get_text(strip=True)
        if not text:
            continue
        a_tag = item.find('a', href=True)
        link_val = ''
        if a_tag:
            href = a_tag['href']
            link_val = href if href.startswith('http') else urllib.parse.urljoin(base_url, href)

        date_tag = item.find(class_=re.compile(r'date|data', re.I)) or item.find('time')
        data_val = date_tag.get_text(strip=True) if date_tag else ''

        rows.append({'data': data_val, 'oggetto': text[:300], 'link': link_val})

    return rows

# ─── Excel Output ─────────────────────────────────────────────────────────────

HEADER_FILL = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
HEADER_FONT = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
LINK_FONT = Font(name='Calibri', color='0563C1', underline='single', size=10)
NORMAL_FONT = Font(name='Calibri', size=10)
THIN_BORDER = Border(
    left=Side(style='thin', color='BFBFBF'),
    right=Side(style='thin', color='BFBFBF'),
    top=Side(style='thin', color='BFBFBF'),
    bottom=Side(style='thin', color='BFBFBF'),
)

COLUMNS = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']
COL_WIDTHS = [14, 60, 40, 30]

def write_sheet(wb, sheet_name, results):
    ws = wb.create_sheet(title=sheet_name[:31])

    # Header row
    for col_idx, (col_name, width) in enumerate(zip(COLUMNS, COL_WIDTHS), start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = 'A2'

    if not results:
        cell = ws.cell(row=2, column=1, value='Nessun risultato')
        cell.font = Font(name='Calibri', italic=True, color='7F7F7F', size=10)
        return

    for row_idx, r in enumerate(results, start=2):
        data_val = r.get('data', '')
        oggetto_val = r.get('oggetto', '')
        desc_breve = oggetto_val[:150] if oggetto_val else ''
        link_val = r.get('link', '')

        # Data
        c1 = ws.cell(row=row_idx, column=1, value=data_val)
        c1.font = NORMAL_FONT
        c1.alignment = Alignment(horizontal='center', vertical='top')
        c1.border = THIN_BORDER

        # Oggetto
        c2 = ws.cell(row=row_idx, column=2, value=oggetto_val)
        c2.font = NORMAL_FONT
        c2.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        c2.border = THIN_BORDER

        # Descrizione breve
        c3 = ws.cell(row=row_idx, column=3, value=desc_breve)
        c3.font = NORMAL_FONT
        c3.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        c3.border = THIN_BORDER

        # Link (clickable hyperlink)
        c4 = ws.cell(row=row_idx, column=4)
        if link_val:
            c4.value = 'Apri documento'
            c4.hyperlink = link_val
            c4.font = LINK_FONT
        else:
            c4.value = '–'
            c4.font = NORMAL_FONT
        c4.alignment = Alignment(horizontal='center', vertical='top')
        c4.border = THIN_BORDER

        # Alternate row fill for readability
        if row_idx % 2 == 0:
            fill = PatternFill(start_color='EBF3FB', end_color='EBF3FB', fill_type='solid')
            for col_idx in range(1, 5):
                ws.cell(row=row_idx, column=col_idx).fill = fill

    # Auto-filter
    ws.auto_filter.ref = f'A1:D{ws.max_row}'

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    sheet_map = [
        ('ASP Cosenza',         'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
        ('ASP Catanzaro',       'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
        ('ASP Crotone',         'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
        ('ASP Reggio Calabria', 'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
        ('ASP Vibo Valentia',   'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ]

    # Scrape ASP portals
    for name, url in sheet_map:
        results = scrape_asp_portal(name, url, DATE_FROM_2DAYS)
        write_sheet(wb, name, results)

    # Scrape Regione Calabria
    rc_results = scrape_regione_calabria(DATE_FROM_2DAYS)
    write_sheet(wb, 'Regione Calabria', rc_results)

    # Scrape TAR Catanzaro
    tar_results = scrape_tar_catanzaro(DATE_FROM_3DAYS)
    write_sheet(wb, 'TAR Catanzaro', tar_results)

    # Save file
    out_path = OUTPUT_FILENAME
    wb.save(out_path)
    print(f"\n{'='*60}")
    print(f"  File salvato: {out_path}")
    print(f"  Percorso assoluto: {os.path.abspath(out_path)}")
    print(f"{'='*60}")

if __name__ == '__main__':
    main()
