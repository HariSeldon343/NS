#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraping ASP portals, Regione Calabria, TAR Catanzaro
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import re
from datetime import datetime, timedelta
import time
import os
import urllib3
import json
from urllib.parse import urljoin, urlencode, urlparse, parse_qs

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Date ──────────────────────────────────────────────────────────────────────
TODAY       = datetime.now()
DATE_FROM_2 = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_FROM_3 = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')
DATE_TO     = TODAY.strftime('%Y-%m-%d')
FILE_DATE   = TODAY.strftime('%d-%m-%Y')

OUTPUT_FILE = f'/home/user/NS/Monitoraggio_Atti_Calabria_{FILE_DATE}.xlsx'

print(f"Date range (2 days): {DATE_FROM_2} → {DATE_TO}")
print(f"Date range (3 days): {DATE_FROM_3} → {DATE_TO}")
print(f"Output: {OUTPUT_FILE}\n")

# ── Keywords ──────────────────────────────────────────────────────────────────
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
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

def matches_keywords(text):
    if not text:
        return False
    t = text.upper()
    for kw in KEYWORDS:
        if kw.upper() in t:
            return True
    return bool(LIFE_RE.search(text))

# ── HTTP session ──────────────────────────────────────────────────────────────
BASE_HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9,en-US;q=0.8',
    'Connection': 'keep-alive',
}

def make_session():
    s = requests.Session()
    s.headers.update(BASE_HEADERS)
    return s

# ── ASP portal scraper ────────────────────────────────────────────────────────
ASP_PORTALS = [
    ('ASP Cosenza',         'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Catanzaro',       'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Crotone',         'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Reggio Calabria', 'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
    ('ASP Vibo Valentia',   'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo'),
]

def collect_form_data(form, exclude_types=('submit', 'button', 'image', 'reset')):
    data = {}
    for tag in form.find_all(['input', 'select', 'textarea']):
        name = tag.get('name') or ''
        val  = tag.get('value', '')
        typ  = tag.get('type', 'text').lower()
        if name and typ not in exclude_types:
            if tag.name == 'select':
                # pick first option or selected option
                selected = tag.find('option', selected=True)
                val = selected['value'] if selected and selected.get('value') else ''
            data[name] = val
    return data

def find_next_page_url(soup, base_url):
    """Return URL of next page or None."""
    # rel="next"
    a = soup.find('a', rel=lambda r: r and 'next' in r)
    if a and a.get('href'):
        return urljoin(base_url, a['href'])
    # text-based
    for a in soup.find_all('a', href=True):
        txt = a.get_text(strip=True)
        if txt in ('Successiva', '>', '»', 'Next', 'Avanti', '›'):
            return urljoin(base_url, a['href'])
    return None

def parse_results_table(soup, base_url):
    """
    Find the results table and return list of dicts {data, oggetto, link}.
    Tries to identify columns by header text.
    """
    rows_found = []

    # Find best matching table
    best_table = None
    for tbl in soup.find_all('table'):
        headers = [th.get_text(strip=True).lower() for th in tbl.find_all('th')]
        if any(h in headers for h in ('oggetto', 'data', 'numero', 'tipo', 'atto', 'descrizione')):
            best_table = tbl
            break
        body_text = tbl.get_text().lower()
        if 'oggetto' in body_text and ('data' in body_text or 'numero' in body_text):
            best_table = tbl
            break

    if best_table is None:
        tables = soup.find_all('table')
        if tables:
            best_table = tables[0]

    if best_table is None:
        return rows_found

    all_rows = best_table.find_all('tr')
    if not all_rows:
        return rows_found

    # Detect header row
    header_cells = all_rows[0].find_all(['th', 'td'])
    headers = [c.get_text(strip=True).lower() for c in header_cells]

    def col_idx(patterns):
        for p in patterns:
            for i, h in enumerate(headers):
                if p in h:
                    return i
        return -1

    data_col    = col_idx(['data', 'date'])
    oggetto_col = col_idx(['oggetto', 'descrizione', 'titolo', 'object'])

    for row in all_rows[1:]:
        cells = row.find_all(['td', 'th'])
        if not cells:
            continue
        texts = [c.get_text(strip=True) for c in cells]

        data_val    = texts[data_col] if 0 <= data_col < len(texts) else (texts[0] if texts else '')
        oggetto_val = texts[oggetto_col] if 0 <= oggetto_col < len(texts) else ' | '.join(texts)

        # Extract link
        link = None
        for c in cells:
            a = c.find('a', href=True)
            if a:
                href = a['href']
                if href and not href.startswith('#'):
                    link = urljoin(base_url, href)
                    break

        rows_found.append({
            'data':    data_val,
            'oggetto': oggetto_val,
            'link':    link or base_url,
            '_all':    ' | '.join(texts),
        })

    return rows_found

def scrape_asp(name, base_url):
    results = []
    s = make_session()
    print(f"\n{'='*60}")
    print(f"[{name}] {base_url}")

    try:
        r = s.get(base_url, timeout=30, verify=False)
        print(f"[{name}] GET {r.status_code}")
        if r.status_code != 200:
            print(f"[{name}] Non-200 response – skipping")
            return results

        soup = BeautifulSoup(r.text, 'lxml')

        # Identify form
        form = soup.find('form')
        if not form:
            # Maybe the page itself is already a result listing without a form
            print(f"[{name}] No <form> found. Trying to parse results directly.")
            rows = parse_results_table(soup, base_url)
            for row in rows:
                if matches_keywords(row['oggetto']) or matches_keywords(row['_all']):
                    results.append(row)
            print(f"[{name}] Direct parse: {len(results)} matches")
            return results

        form_action = form.get('action') or base_url
        if not form_action.startswith('http'):
            form_action = urljoin(base_url, form_action)
        form_method = (form.get('method') or 'post').lower()

        form_data = collect_form_data(form)
        print(f"[{name}] Form fields ({len(form_data)}): {list(form_data.keys())[:15]}")

        # Inject date filter
        set_date = False
        for key in list(form_data.keys()):
            lkey = key.lower()
            if 'datapubblicazionedal' in lkey or lkey == 'datadal' or lkey == 'dal':
                form_data[key] = DATE_FROM_2
                set_date = True
            if 'datapubblicazioneal' in lkey or lkey == 'dataal' or lkey == 'al':
                form_data[key] = DATE_TO

        if not set_date:
            form_data['dataPubblicazioneDal'] = DATE_FROM_2
            form_data['dataPubblicazioneAl']  = DATE_TO
            print(f"[{name}] Injected dataPubblicazioneDal / Al manually")

        page = 1
        while True:
            print(f"[{name}] Submitting page {page} to {form_action}")
            if form_method == 'post':
                r2 = s.post(form_action, data=form_data, timeout=30, verify=False)
            else:
                r2 = s.get(form_action, params=form_data, timeout=30, verify=False)

            print(f"[{name}] Response {r2.status_code} ({len(r2.content)} bytes)")
            if r2.status_code != 200:
                break

            soup2 = BeautifulSoup(r2.text, 'lxml')

            if page == 1:
                title = soup2.title.string.strip() if soup2.title else 'N/A'
                print(f"[{name}] Page title: {title}")

            rows = parse_results_table(soup2, base_url)
            print(f"[{name}] Page {page}: {len(rows)} rows in table")

            if not rows:
                # Debug first 500 chars of body
                body_snippet = soup2.get_text()[:300].replace('\n', ' ')
                print(f"[{name}] Body snippet: {body_snippet}")
                break

            page_matches = 0
            for row in rows:
                search_text = row['oggetto'] + ' ' + row.get('_all', '')
                if matches_keywords(search_text):
                    page_matches += 1
                    results.append({
                        'data':    row['data'],
                        'oggetto': row['oggetto'],
                        'link':    row['link'],
                    })
            print(f"[{name}] Page {page}: {page_matches} keyword matches")

            next_url = find_next_page_url(soup2, base_url)
            if not next_url or next_url == form_action:
                break

            form_action = next_url
            form_method = 'get'
            form_data   = {}
            page += 1
            time.sleep(0.5)

    except Exception as exc:
        import traceback
        print(f"[{name}] ERROR: {exc}")
        traceback.print_exc()

    print(f"[{name}] TOTAL MATCHES: {len(results)}")
    return results

# ── Regione Calabria ──────────────────────────────────────────────────────────

def scrape_regione_calabria():
    results = []
    base_url = 'https://www.regione.calabria.it/provvedimenti-della-regione/'
    s = make_session()

    print(f"\n{'='*60}")
    print(f"[Regione Calabria] {base_url}")

    try:
        # Try multiple date param conventions
        param_variants = [
            {'filter_date_from': DATE_FROM_2},
            {'data_dal': DATE_FROM_2},
            {'from': DATE_FROM_2},
            {'date_from': DATE_FROM_2},
        ]

        page_url = base_url
        params   = param_variants[0]
        page     = 1

        while True:
            url_to_fetch = page_url if page > 1 else base_url
            p = params if page == 1 else {}

            r = s.get(url_to_fetch, params=p, timeout=30, verify=False)
            print(f"[Regione Calabria] Page {page}: {r.status_code} ({len(r.content)} bytes)")
            if r.status_code not in (200, 301, 302):
                break

            soup = BeautifulSoup(r.text, 'lxml')

            if page == 1:
                title = soup.title.string.strip() if soup.title else 'N/A'
                print(f"[Regione Calabria] Title: {title}")

            page_rows = 0
            page_matches = 0

            # --- Table rows ---
            for tbl in soup.find_all('table'):
                rows = tbl.find_all('tr')
                headers = [th.get_text(strip=True).lower() for th in rows[0].find_all(['th', 'td'])] if rows else []
                obj_col  = next((i for i, h in enumerate(headers) if 'oggetto' in h or 'descrizione' in h or 'titolo' in h), -1)
                data_col = next((i for i, h in enumerate(headers) if 'data' in h), 0)
                for row in rows[1:]:
                    cells = row.find_all(['td', 'th'])
                    if not cells:
                        continue
                    page_rows += 1
                    texts = [c.get_text(strip=True) for c in cells]
                    full  = ' | '.join(texts)
                    data_val    = texts[data_col] if data_col < len(texts) else ''
                    oggetto_val = texts[obj_col]  if 0 <= obj_col < len(texts) else full
                    link = None
                    for c in cells:
                        a = c.find('a', href=True)
                        if a and a['href'] and not a['href'].startswith('#'):
                            link = urljoin(base_url, a['href'])
                            break
                    if matches_keywords(full):
                        page_matches += 1
                        results.append({'data': data_val, 'oggetto': oggetto_val, 'link': link or base_url})

            # --- Article/div-based layout ---
            selectors = [
                soup.find_all('article'),
                soup.find_all('li', class_=re.compile(r'provvedimento|atto|item|post', re.I)),
                soup.find_all('div', class_=re.compile(r'provvedimento|atto|item|views-row', re.I)),
            ]
            for items in selectors:
                for item in items:
                    page_rows += 1
                    text = item.get_text(' ', strip=True)
                    link = None
                    a = item.find('a', href=True)
                    if a and a['href'] and not a['href'].startswith('#'):
                        link = urljoin(base_url, a['href'])
                    if matches_keywords(text):
                        page_matches += 1
                        # try to extract date
                        date_tag = item.find(class_=re.compile(r'date|data', re.I))
                        data_val = date_tag.get_text(strip=True) if date_tag else ''
                        title_tag = item.find(['h2', 'h3', 'h4', 'strong', 'a'])
                        oggetto_val = title_tag.get_text(strip=True) if title_tag else text[:200]
                        results.append({'data': data_val, 'oggetto': oggetto_val, 'link': link or base_url})

            print(f"[Regione Calabria] Page {page}: {page_rows} items found, {page_matches} matches")

            if page_rows == 0:
                snippet = soup.get_text()[:500].replace('\n', ' ')
                print(f"[Regione Calabria] Body snippet: {snippet}")
                break

            next_url = find_next_page_url(soup, base_url)
            if not next_url or next_url == page_url:
                break
            page_url = next_url
            page += 1
            time.sleep(0.5)

    except Exception as exc:
        import traceback
        print(f"[Regione Calabria] ERROR: {exc}")
        traceback.print_exc()

    print(f"[Regione Calabria] TOTAL MATCHES: {len(results)}")
    return results

# ── TAR Catanzaro ─────────────────────────────────────────────────────────────

def scrape_tar():
    results = []
    base_url = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'
    s = make_session()
    # GA site sometimes needs a Referer
    s.headers.update({'Referer': 'https://www.giustizia-amministrativa.it/'})

    print(f"\n{'='*60}")
    print(f"[TAR Catanzaro] {base_url}")

    # Known param names used by GA portal
    param_variants = [
        {'publishDateFrom': DATE_FROM_3, 'publishDateTo': DATE_TO},
        {'dateFrom': DATE_FROM_3, 'dateTo': DATE_TO},
        {'dataDal': DATE_FROM_3, 'dataAl': DATE_TO},
        {'filter_date_from': DATE_FROM_3},
    ]

    try:
        page     = 1
        page_url = base_url
        params   = param_variants[0]

        while True:
            r = s.get(page_url, params=params if page == 1 else {}, timeout=30, verify=False)
            print(f"[TAR] Page {page}: {r.status_code} ({len(r.content)} bytes)")
            if r.status_code not in (200,):
                break

            soup = BeautifulSoup(r.text, 'lxml')
            if page == 1:
                title = soup.title.string.strip() if soup.title else 'N/A'
                print(f"[TAR] Title: {title}")

            page_rows    = 0
            page_matches = 0

            for tbl in soup.find_all('table'):
                rows    = tbl.find_all('tr')
                if not rows:
                    continue
                headers = [th.get_text(strip=True).lower() for th in rows[0].find_all(['th', 'td'])]
                # TAR: columns might be: Data, Numero, Sezione, Tipo, Parte, Oggetto, PDF
                data_col  = next((i for i, h in enumerate(headers) if 'data' in h), 0)
                parte_col = next((i for i, h in enumerate(headers) if 'parte' in h or 'ricorrente' in h), -1)
                obj_col   = next((i for i, h in enumerate(headers) if 'oggetto' in h or 'descrizione' in h), -1)

                for row in rows[1:]:
                    cells = row.find_all(['td', 'th'])
                    if not cells:
                        continue
                    page_rows += 1
                    texts   = [c.get_text(strip=True) for c in cells]
                    full    = ' | '.join(texts)
                    data_val    = texts[data_col] if data_col < len(texts) else ''
                    check_cols  = []
                    if 0 <= parte_col < len(texts):
                        check_cols.append(texts[parte_col])
                    if 0 <= obj_col < len(texts):
                        check_cols.append(texts[obj_col])
                    if not check_cols:
                        check_cols.append(full)
                    search_text = ' '.join(check_cols)

                    link = None
                    for c in cells:
                        a = c.find('a', href=True)
                        if a and a['href'] and not a['href'].startswith('#'):
                            link = urljoin(base_url, a['href'])
                            break

                    if matches_keywords(search_text) or matches_keywords(full):
                        page_matches += 1
                        oggetto_val = texts[obj_col] if 0 <= obj_col < len(texts) else full
                        results.append({'data': data_val, 'oggetto': oggetto_val, 'link': link or base_url})

            # div/article layout fallback
            for item in soup.find_all(['article', 'li', 'div'],
                                       class_=re.compile(r'provvedimento|atto|item|row|result', re.I)):
                page_rows += 1
                text = item.get_text(' ', strip=True)
                link = None
                a = item.find('a', href=True)
                if a and a['href']:
                    link = urljoin(base_url, a['href'])
                if matches_keywords(text):
                    page_matches += 1
                    date_tag = item.find(class_=re.compile(r'date|data', re.I))
                    data_val = date_tag.get_text(strip=True) if date_tag else ''
                    t_tag    = item.find(['h2', 'h3', 'h4', 'strong', 'a'])
                    oggetto_val = t_tag.get_text(strip=True) if t_tag else text[:200]
                    results.append({'data': data_val, 'oggetto': oggetto_val, 'link': link or base_url})

            print(f"[TAR] Page {page}: {page_rows} rows, {page_matches} matches")

            if page_rows == 0:
                snippet = soup.get_text()[:500].replace('\n', ' ')
                print(f"[TAR] Body snippet: {snippet}")
                break

            next_url = find_next_page_url(soup, base_url)
            if not next_url or next_url == page_url:
                break
            page_url = next_url
            page += 1
            time.sleep(0.5)

    except Exception as exc:
        import traceback
        print(f"[TAR] ERROR: {exc}")
        traceback.print_exc()

    print(f"[TAR] TOTAL MATCHES: {len(results)}")
    return results

# ── Excel output ──────────────────────────────────────────────────────────────

SHEET_NAMES = [
    'ASP Cosenza', 'ASP Catanzaro', 'ASP Crotone',
    'ASP Reggio Calabria', 'ASP Vibo Valentia',
    'Regione Calabria', 'TAR Catanzaro',
]

def write_excel(all_results):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_font  = Font(bold=True)
    link_font    = Font(color='0563C1', underline='single')
    wrap_align   = Alignment(wrap_text=True, vertical='top')

    for sheet_name, rows in zip(SHEET_NAMES, all_results):
        ws = wb.create_sheet(title=sheet_name)

        # Column headers
        for col_idx, hdr in enumerate(['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link'], 1):
            c = ws.cell(row=1, column=col_idx, value=hdr)
            c.font      = header_font
            c.alignment = wrap_align

        ws.row_dimensions[1].height = 18

        if not rows:
            ws.cell(row=2, column=1, value='Nessun risultato')
        else:
            for r_idx, item in enumerate(rows, 2):
                data_val    = item.get('data', '')
                oggetto_val = item.get('oggetto', '')
                link_val    = item.get('link', '')
                descr_val   = oggetto_val[:150] if oggetto_val else ''

                ws.cell(row=r_idx, column=1, value=data_val).alignment    = wrap_align
                ws.cell(row=r_idx, column=2, value=oggetto_val).alignment = wrap_align
                ws.cell(row=r_idx, column=3, value=descr_val).alignment   = wrap_align

                lc = ws.cell(row=r_idx, column=4, value=link_val)
                lc.alignment = wrap_align
                if link_val and link_val.startswith('http'):
                    lc.hyperlink = link_val
                    lc.font      = link_font

        # Column widths
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 70
        ws.column_dimensions['C'].width = 55
        ws.column_dimensions['D'].width = 55

        # Freeze header row
        ws.freeze_panes = 'A2'

    wb.save(OUTPUT_FILE)
    print(f"\n✓ File salvato: {OUTPUT_FILE}")

# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    all_results = []

    # 5 ASP portals
    for name, url in ASP_PORTALS:
        all_results.append(scrape_asp(name, url))

    # Regione Calabria
    all_results.append(scrape_regione_calabria())

    # TAR Catanzaro
    all_results.append(scrape_tar())

    # Write Excel
    write_excel(all_results)

    # Summary
    print(f"\n{'='*60}")
    print("RIEPILOGO FINALE")
    for name, res in zip(SHEET_NAMES, all_results):
        print(f"  {name:<25} {len(res):>4} atti")
    print(f"\nFile: {OUTPUT_FILE}")
