#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scarica atti da portali ASP, Regione Calabria e TAR Catanzaro,
filtra per keyword e genera file Excel.

Requisiti:
    pip install requests beautifulsoup4 lxml openpyxl

Uso:
    python monitoraggio_atti_calabria.py

NOTA: I portali ASP Calabria (sisr.regione.calabria.it) e TAR Catanzaro
(giustizia-amministrativa.it) sono accessibili SOLO da IP italiani.
Eseguire sempre lo script da un PC con connessione italiana (ufficio / VPN IT).
"""

import os
import requests
import re
import sys
import time
from bs4 import BeautifulSoup
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

# ─── Configurazione percorso output ───────────────────────────────────────────
# Modifica OUTPUT_DIR con il percorso desiderato (es. su Windows):
#   OUTPUT_DIR = r"C:\Users\aoedo\kDrive\01_Lavoro\Clienti\Starbene\Atti"
# Lascia '' per salvare nella cartella corrente.
OUTPUT_DIR = ''

# ─── Date ────────────────────────────────────────────────────────────────────
TODAY       = date.today()
DATE_FROM   = TODAY - timedelta(days=2)        # ultimi 2 giorni
DATE_FROM3  = TODAY - timedelta(days=3)        # ultimi 3 giorni (TAR)
DATE_FROM_S = DATE_FROM.strftime('%Y-%m-%d')
DATE_FROM3_S= DATE_FROM3.strftime('%Y-%m-%d')

EXEC_DATE   = TODAY.strftime('%d-%m-%Y')
_fname      = f"Monitoraggio_Atti_Calabria_{EXEC_DATE}.xlsx"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, _fname) if OUTPUT_DIR else _fname

print(f"Monitoraggio al {TODAY}  |  Da: {DATE_FROM_S}  |  File: {OUTPUT_FILE}")

# ─── Keywords ────────────────────────────────────────────────────────────────
KEYWORDS = [
    'ADI', 'Assistenza domiciliare', 'ANMIC', 'Accreditamento',
    'Aumento di budget', 'Autismo', "Autorizzazione all'esercizio",
    'Autorizzazione alla realizzazione', 'Autorizzazioni', 'Budget',
    'Casa Giardino', 'Centro San Giuseppe', 'Centro salute e benessere',
    'Fabbisogni LEA', 'Fisiolab', 'Fisioterapia',
    'Parere commissione', "Presa d'atto verifica", 'Programmazione',
    'Rete riabilitativa', 'Rete territoriale', 'Riabilitazione estensiva',
    'Riconversione prestazioni', 'Rinnovo accreditamento', 'San Teodoro',
    'Savelli Hospital', 'Starbene', 'Verifica requisiti',
    'Villa San Giuseppe', 'Villa del Rosario',
]
_LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

def matches_keywords(text: str) -> bool:
    t = text.upper()
    for kw in KEYWORDS:
        if kw.upper() in t:
            return True
    return bool(_LIFE_RE.search(text))

# ─── HTTP Session ─────────────────────────────────────────────────────────────
HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
        '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'it-IT,it;q=0.9,en;q=0.8',
}

def make_session(verify_ssl=True):
    s = requests.Session()
    s.headers.update(HEADERS)
    s.verify = verify_ssl
    return s

# ─── ASP portals ─────────────────────────────────────────────────────────────
ASP_PORTALS = {
    'ASP Cosenza':         'https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Catanzaro':       'https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Crotone':         'https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Reggio Calabria': 'https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
    'ASP Vibo Valentia':   'https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo',
}

def scrape_asp_portal(name: str, base_url: str) -> list[dict]:
    """Scrape one ASP AlboOnline portal."""
    print(f"\n  [{name}] {base_url}")
    results = []
    session = make_session()

    try:
        # Step 1: GET homepage to get JSF ViewState (if present)
        r = session.get(base_url, timeout=20)
        if r.status_code == 503:
            raise Exception("503 – Portale non raggiungibile dall'IP corrente (restrizione geografica). "
                            "Eseguire lo script da una rete italiana.")
        r.raise_for_status()
        soup = BeautifulSoup(r.text, 'lxml')

        # Extract ViewState (JSF)
        vs_input = soup.find('input', {'name': re.compile(r'javax\.faces\.ViewState|com\.sun\.faces\.VIEW', re.I)})
        viewstate = vs_input['value'] if vs_input else ''

        # Extract form action
        form = soup.find('form')
        form_action = base_url
        if form:
            action = form.get('action', '')
            if action and not action.startswith('http'):
                from urllib.parse import urljoin
                form_action = urljoin(base_url, action)

        page_num = 0
        max_pages = 500
        per_page_estimate = 10

        while page_num < max_pages:
            # Build POST payload
            payload = {
                'dataPubblicazioneDal': DATE_FROM_S,
                'dataPubblicazioneAl':  TODAY.strftime('%Y-%m-%d'),
                'pageNum':              str(page_num),
                'javax.faces.ViewState': viewstate,
            }
            # Also include all hidden fields from form
            if form:
                for hidden in form.find_all('input', type='hidden'):
                    n = hidden.get('name', '')
                    v = hidden.get('value', '')
                    if n and n not in payload:
                        payload[n] = v

            resp = session.post(form_action, data=payload, timeout=25)
            resp.raise_for_status()
            page_soup = BeautifulSoup(resp.text, 'lxml')

            rows_found = 0
            table = page_soup.find('table')
            if table:
                rows = table.find_all('tr')
                for row in rows[1:]:  # skip header
                    cols = row.find_all(['td', 'th'])
                    if len(cols) < 3:
                        continue
                    oggetto = ''
                    data_val = ''
                    link_url = ''
                    for i, col in enumerate(cols):
                        txt = col.get_text(strip=True)
                        if len(txt) > len(oggetto) and len(txt) > 10:
                            oggetto = txt
                        if re.match(r'\d{2}/\d{2}/\d{4}', txt) and not data_val:
                            data_val = txt
                        a = col.find('a', href=True)
                        if a and not link_url:
                            from urllib.parse import urljoin
                            link_url = urljoin(base_url, a['href'])

                    if oggetto and matches_keywords(oggetto):
                        results.append({'data': data_val, 'oggetto': oggetto, 'link': link_url})
                    rows_found += 1

            print(f"    Pagina {page_num+1}: {rows_found} righe trovate")
            if rows_found < per_page_estimate:
                break
            page_num += 1
            time.sleep(0.3)

    except requests.exceptions.Timeout:
        msg = ('Portale non raggiungibile: timeout di connessione. '
               'I portali SISR Calabria sono accessibili solo da IP italiani. '
               f'Eseguire lo script localmente. URL: {base_url}')
        print(f"    TIMEOUT")
        results = [{'_error': msg}]
    except requests.exceptions.ConnectionError as e:
        msg = (f'Errore di connessione: {str(e)[:200]}. '
               'I portali SISR richiedono accesso da rete italiana.')
        print(f"    CONNECTION ERROR")
        results = [{'_error': msg}]
    except Exception as e:
        msg = str(e)[:400]
        print(f"    ERROR: {msg}")
        results = [{'_error': msg}]

    return results


# ─── Regione Calabria ─────────────────────────────────────────────────────────
RC_BASE = 'https://www.regione.calabria.it/provvedimenti-della-regione/'

def scrape_regione_calabria() -> list[dict]:
    """Scrape provvedimenti Regione Calabria."""
    print(f"\n  [Regione Calabria] {RC_BASE}")
    results = []
    session = make_session()
    paged = 1

    while True:
        params = {
            'paged':            str(paged),
            'pr':               '',
            'sort_order':       '0',
            'filter_active':    'true',
            'filter_department':'',
            'filter_tematic':   '',
            'filter_type':      '',
            'filter_item':      '',
            'filter_number':    '',
            'filter_date_from': DATE_FROM_S,
            'filter_date_to':   '',
            'filter_date_accept_from': '',
            'filter_date_accept_to':   '',
            'pageNum':          '0',
            'searchButton':     'Cerca',
        }

        try:
            url = RC_BASE
            if paged > 1:
                url = RC_BASE + f'?paged={paged}'
            r = session.get(RC_BASE, params=params, timeout=25)
            r.raise_for_status()
        except Exception as e:
            print(f"    ERROR page {paged}: {e}")
            break

        soup = BeautifulSoup(r.text, 'lxml')

        # Total count
        count_span = soup.find(string=re.compile(r'\d+\s+elementi utili'))
        if paged == 1 and count_span:
            count_str = re.search(r'(\d[\d.]*)\s+elementi utili', count_span)
            if count_str:
                cnt = int(count_str.group(1).replace('.',''))
                print(f"    Totale elementi: {cnt}")

        table = soup.find('table', class_='table')
        if not table:
            print(f"    Pagina {paged}: nessuna tabella")
            break

        rows = table.find_all('tr')
        page_count = 0
        for row in rows[1:]:
            cols = row.find_all('td')
            if len(cols) < 5:
                continue
            tipologia   = cols[0].get_text(strip=True)
            data_rep    = cols[1].get_text(strip=True)
            num         = cols[2].get_text(strip=True) if len(cols) > 2 else ''
            dipartimento= cols[3].get_text(strip=True) if len(cols) > 3 else ''
            oggetto     = cols[4].get_text(strip=True) if len(cols) > 4 else ''
            # Link from last column
            link_url = ''
            if len(cols) > 5:
                a = cols[5].find('a', href=True)
                if a:
                    link_url = a['href']
                    if not link_url.startswith('http'):
                        link_url = 'https://www.regione.calabria.it' + link_url

            if matches_keywords(oggetto):
                results.append({
                    'data': data_rep,
                    'oggetto': oggetto,
                    'link': link_url,
                    'extra': f"{tipologia} n.{num} - {dipartimento}",
                })
            page_count += 1

        print(f"    Pagina {paged}: {page_count} righe, {len(results)} match cumulativi")

        # Check for next page
        next_link = soup.find('a', class_='next page-numbers')
        if not next_link:
            break

        paged += 1
        time.sleep(0.4)

    return results


# ─── TAR Catanzaro ───────────────────────────────────────────────────────────
TAR_BASE = 'https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro'

def scrape_tar_catanzaro() -> list[dict]:
    """Scrape TAR Catanzaro."""
    print(f"\n  [TAR Catanzaro] {TAR_BASE}")
    results = []

    # Try multiple SSL configs
    for verify in [True, False]:
        try:
            session = make_session(verify_ssl=verify)
            r = session.get(TAR_BASE, timeout=20)
            r.raise_for_status()
            print(f"    Connessione OK (verify={verify}), size={len(r.text)}")
            soup = BeautifulSoup(r.text, 'lxml')
            break
        except requests.exceptions.SSLError as e:
            print(f"    SSL error (verify={verify}): {str(e)[:120]}")
            soup = None
            continue
        except requests.exceptions.Timeout:
            print(f"    TIMEOUT")
            return [{'_error': 'Portale non raggiungibile: timeout'}]
        except Exception as e:
            print(f"    Errore: {str(e)[:120]}")
            soup = None
            continue

    if soup is None:
        return [{'_error': (
            'Portale non raggiungibile: errore SSL/TLS. '
            'Il sito giustizia-amministrativa.it utilizza una configurazione TLS '
            'incompatibile con questo ambiente cloud. '
            'Eseguire lo script localmente da rete italiana.'
        )}]

    # TAR uses its own pagination/search API - look for the form or API
    # Try with publishDateFrom parameter
    page = 1
    while True:
        try:
            params = {
                'publishDateFrom': DATE_FROM3_S,
                'page': str(page),
            }
            r = session.get(TAR_BASE, params=params, timeout=20)
            r.raise_for_status()
        except Exception as e:
            print(f"    Errore pagina {page}: {e}")
            break

        soup = BeautifulSoup(r.text, 'lxml')
        table = soup.find('table')
        if not table:
            # Try looking for article/card structure
            items = soup.find_all(['article', 'div'], class_=re.compile(r'result|item|row'))
            if not items:
                print(f"    Pagina {page}: nessuna tabella trovata")
                break
            for item in items:
                text = item.get_text(strip=True)
                if matches_keywords(text):
                    a = item.find('a', href=True)
                    link = a['href'] if a else ''
                    results.append({
                        'data': '',
                        'oggetto': text[:500],
                        'link': link,
                    })
            break

        rows = table.find_all('tr')
        page_count = 0
        for row in rows[1:]:
            cols = row.find_all(['td', 'th'])
            texts = [c.get_text(strip=True) for c in cols]
            full_text = ' '.join(texts)
            if not full_text.strip():
                continue

            # Find "Parte" column and date
            data_val = ''
            parte = ''
            oggetto_full = full_text
            for t in texts:
                if re.match(r'\d{2}/\d{2}/\d{4}', t):
                    data_val = t
                elif t and t != 'XXX_OMISSIS_XXX' and not re.match(r'\d', t):
                    if len(t) > len(parte):
                        parte = t
            link_url = ''
            a_tag = row.find('a', href=True)
            if a_tag:
                href = a_tag['href']
                if not href.startswith('http'):
                    link_url = 'https://www.giustizia-amministrativa.it' + href
                else:
                    link_url = href

            if matches_keywords(oggetto_full) or matches_keywords(parte):
                results.append({
                    'data': data_val,
                    'oggetto': oggetto_full[:500],
                    'link': link_url,
                })
            page_count += 1

        print(f"    Pagina {page}: {page_count} righe, {len(results)} match")

        # Pagination
        next_btn = soup.find('a', string=re.compile(r'success|next|>|›', re.I))
        if not next_btn or page >= 50:
            break
        page += 1
        time.sleep(0.3)

    return results


# ─── Excel output ─────────────────────────────────────────────────────────────
def write_excel(all_results: dict, filename: str):
    from openpyxl.styles import PatternFill, Alignment
    from openpyxl.styles.borders import Border, Side

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    BLUE_FONT   = Font(color='0563C1', underline='single')
    HEADER_FONT = Font(bold=True, color='FFFFFF')
    ERROR_FONT  = Font(italic=True, color='CC0000')
    HEADER_FILL = PatternFill('solid', fgColor='17375E')
    ALT_FILL    = PatternFill('solid', fgColor='DCE6F1')
    WRAP        = Alignment(wrap_text=True, vertical='top')

    # ── RIEPILOGO sheet (first) ─────────────────────────────────────────────
    ws0 = wb.create_sheet('RIEPILOGO', 0)
    ws0['A1'] = f'Monitoraggio Atti Calabria – {TODAY.strftime("%d/%m/%Y")}'
    ws0['A1'].font = Font(bold=True, size=14)
    ws0['A2'] = f'Periodo: dal {DATE_FROM_S} al {TODAY}'
    ws0['A2'].font = Font(italic=True, size=11)
    ws0['A4'] = 'Portale'
    ws0['B4'] = 'Stato'
    ws0['C4'] = 'Match trovati'
    for c in ['A4','B4','C4']:
        ws0[c].font = HEADER_FONT
        ws0[c].fill = HEADER_FILL
    ws0.column_dimensions['A'].width = 28
    ws0.column_dimensions['B'].width = 70
    ws0.column_dimensions['C'].width = 15

    riepilogo_row = 5
    for sheet_name, rows in all_results.items():
        ws0.cell(row=riepilogo_row, column=1, value=sheet_name)
        if rows and '_error' in rows[0]:
            ws0.cell(row=riepilogo_row, column=2, value=rows[0]['_error'][:200])
            ws0.cell(row=riepilogo_row, column=2).font = ERROR_FONT
            ws0.cell(row=riepilogo_row, column=3, value='N/A')
        else:
            stato = 'OK – nessun match' if not rows else f'OK – {len(rows)} match'
            ws0.cell(row=riepilogo_row, column=2, value=stato)
            ws0.cell(row=riepilogo_row, column=3, value=len(rows))
        riepilogo_row += 1

    ws0['A' + str(riepilogo_row + 1)] = (
        'NOTA: I portali ASP (sisr.regione.calabria.it) e TAR Catanzaro (giustizia-amministrativa.it) '
        'sono accessibili solo da IP italiani. Eseguire lo script localmente dalla rete dell\'ufficio.'
    )
    ws0['A' + str(riepilogo_row + 1)].font = Font(italic=True, color='666666')
    ws0.merge_cells(f'A{riepilogo_row+1}:C{riepilogo_row+1}')

    # ── Data sheets ────────────────────────────────────────────────────────
    for sheet_name, rows in all_results.items():
        ws = wb.create_sheet(title=sheet_name[:31])

        if rows and '_error' in rows[0]:
            ws['A1'] = '⚠ Portale non disponibile'
            ws['A1'].font = Font(bold=True, color='CC0000', size=12)
            ws['A2'] = rows[0]['_error']
            ws['A2'].font = ERROR_FONT
            ws['A2'].alignment = WRAP
            ws.column_dimensions['A'].width = 100
            ws.row_dimensions[2].height = 60
            continue

        headers = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.font    = HEADER_FONT
            cell.fill    = HEADER_FILL
            cell.alignment = Alignment(horizontal='center', vertical='center')

        if not rows:
            ws.cell(row=2, column=1, value='Nessun risultato nel periodo di riferimento')
        else:
            for row_idx, item in enumerate(rows, 2):
                oggetto  = item.get('oggetto', '')
                data_val = item.get('data',    '')
                link_url = item.get('link',    '')

                c1 = ws.cell(row=row_idx, column=1, value=data_val)
                c2 = ws.cell(row=row_idx, column=2, value=oggetto)
                c3 = ws.cell(row=row_idx, column=3, value=oggetto[:150])
                c2.alignment = WRAP
                c3.alignment = WRAP

                if row_idx % 2 == 0:
                    for c in [c1, c2, c3]:
                        c.fill = ALT_FILL

                link_cell = ws.cell(row=row_idx, column=4)
                if link_url:
                    link_cell.value     = 'Apri →'
                    link_cell.hyperlink = link_url
                    link_cell.font      = BLUE_FONT
                    link_cell.alignment = Alignment(horizontal='center')

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 85
        ws.column_dimensions['C'].width = 55
        ws.column_dimensions['D'].width = 10
        ws.freeze_panes = 'A2'
        ws.row_dimensions[1].height = 20

    wb.save(filename)
    print(f"\n  File salvato: {filename}")


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    all_results = {}

    # 1. ASP portals
    for name, url in ASP_PORTALS.items():
        res = scrape_asp_portal(name, url)
        # Filter out only real matches (not error dicts with _error key)
        all_results[name] = res

    # 2. Regione Calabria
    res = scrape_regione_calabria()
    all_results['Regione Calabria'] = res

    # 3. TAR Catanzaro
    res = scrape_tar_catanzaro()
    all_results['TAR Catanzaro'] = res

    # Summary
    print("\n=== RIEPILOGO ===")
    for name, rows in all_results.items():
        if rows and '_error' in rows[0]:
            print(f"  {name}: ERRORE - {rows[0]['_error'][:80]}")
        else:
            print(f"  {name}: {len(rows)} match")

    # Write Excel
    write_excel(all_results, OUTPUT_FILE)
    print(f"\nCompletato. File: {OUTPUT_FILE}")


if __name__ == '__main__':
    main()
