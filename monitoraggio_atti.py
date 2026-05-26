#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Accede ai portali ASP, Regione Calabria e TAR Catanzaro,
filtra per keyword e produce un file xlsx.
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
from datetime import datetime, timedelta
import time
import urllib.parse
import logging
import sys
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger(__name__)

# ─── Config ────────────────────────────────────────────────────────────────────
TODAY        = datetime.now()
DATE_STR_ISO = TODAY.strftime('%Y-%m-%d')
DATE_FROM_2  = (TODAY - timedelta(days=2)).strftime('%Y-%m-%d')
DATE_FROM_3  = (TODAY - timedelta(days=3)).strftime('%Y-%m-%d')
DATE_FILE    = TODAY.strftime('%d-%m-%Y')

OUTPUT_FILENAME = f"Monitoraggio_Atti_Calabria_{DATE_FILE}.xlsx"
OUTPUT_PATH     = f"/home/user/NS/{OUTPUT_FILENAME}"

KEYWORDS = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo", "Autorizzazione all'esercizio",
    "Autorizzazione alla realizzazione", "Autorizzazioni", "Budget",
    "Casa Giardino", "Centro San Giuseppe", "Centro salute e benessere",
    "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale", "Riabilitazione estensiva",
    "Riconversione prestazioni", "Rinnovo accreditamento", "San Teodoro",
    "Savelli Hospital", "Starbene", "Verifica requisiti",
    "Villa San Giuseppe", "Villa del Rosario",
]
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)
KW_UPPER = [k.upper() for k in KEYWORDS]

def keyword_match(text: str) -> bool:
    if not text:
        return False
    t = text.upper()
    for kw in KW_UPPER:
        if kw in t:
            return True
    return bool(LIFE_RE.search(text))

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
    'Upgrade-Insecure-Requests': '1',
}

def make_session():
    s = requests.Session()
    s.headers.update(BROWSER_HEADERS)
    return s


# ─── ASP Portals ───────────────────────────────────────────────────────────────
ASP_PORTALS = [
    ("ASP Cosenza",         "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Catanzaro",       "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Crotone",         "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Reggio Calabria", "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
    ("ASP Vibo Valentia",   "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo"),
]

def scrape_asp_portal(name: str, base_url: str):
    """
    Portali AlboOnline SISR Regione Calabria.
    Struttura: form POST con dataPubblicazioneDal / dataPubblicazioneAl,
    risultati in tabella HTML, paginazione con parametro pagina/page.
    """
    log.info(f"[{name}] Avvio scraping: {base_url}")
    results = []
    session = make_session()
    error_msg = None

    # ── GET iniziale per raccogliere hidden fields ──
    try:
        resp = session.get(base_url, timeout=30, verify=False)
        resp.raise_for_status()
    except requests.exceptions.Timeout:
        error_msg = "Timeout di connessione (portale non raggiungibile dall'ambiente cloud)"
        log.warning(f"[{name}] {error_msg}")
        return results, error_msg
    except requests.exceptions.ConnectionError as e:
        error_msg = f"Errore di connessione: {e}"
        log.warning(f"[{name}] {error_msg}")
        return results, error_msg
    except Exception as e:
        if resp.status_code == 503:
            error_msg = "Portale non raggiungibile (HTTP 503 - possibile blocco IP cloud)"
        else:
            error_msg = f"Errore HTTP {resp.status_code}: {e}"
        log.warning(f"[{name}] {error_msg}")
        return results, error_msg

    soup0 = BeautifulSoup(resp.text, 'lxml')
    form  = soup0.find('form')
    action = base_url
    method = 'post'
    if form:
        rel = form.get('action', '')
        if rel:
            action = urllib.parse.urljoin(base_url, rel)
        method = form.get('method', 'post').lower()

    # Hidden inputs (JSF ViewState, token, ecc.)
    hidden = {}
    if form:
        for inp in form.find_all('input', type='hidden'):
            n = inp.get('name') or ''
            if n:
                hidden[n] = inp.get('value', '')

    base_payload = {
        **hidden,
        'dataPubblicazioneDal': DATE_FROM_2,
        'dataPubblicazioneAl':  DATE_STR_ISO,
    }

    page = 1
    max_pages = 100

    while page <= max_pages:
        payload = {**base_payload, 'pagina': page, 'page': page, 'currentPage': page}
        try:
            if method == 'get':
                r = session.get(action, params=payload, timeout=30, verify=False)
            else:
                r = session.post(action, data=payload, timeout=30, verify=False)
            r.raise_for_status()
        except Exception as e:
            log.warning(f"[{name}] Errore pag.{page}: {e}")
            break

        page_soup = BeautifulSoup(r.text, 'lxml')
        rows = _parse_asp_table(page_soup, base_url)

        if not rows:
            if page == 1:
                log.info(f"[{name}] Nessun risultato a pag.1")
            else:
                log.info(f"[{name}] Fine risultati a pag.{page}")
            break

        results.extend(rows)
        log.info(f"[{name}] Pag.{page}: {len(rows)} atti, totale: {len(results)}")

        if not _has_next(page_soup):
            break
        page += 1
        time.sleep(0.3)

    return results, None


def _parse_asp_table(soup, base_url: str):
    rows = []
    for table in soup.find_all('table'):
        all_rows = table.find_all('tr')
        if len(all_rows) < 2:
            continue

        header_row = all_rows[0]
        ths = header_row.find_all(['th', 'td'])
        header_texts = [th.get_text(strip=True).lower() for th in ths]

        col = {'oggetto': -1, 'data': -1}
        for i, h in enumerate(header_texts):
            if any(x in h for x in ['oggetto', 'descr', 'titolo', 'denominaz']):
                col['oggetto'] = i
            elif any(x in h for x in ['data', 'pubbl', 'registr']):
                col['data'] = i

        for tr in all_rows[1:]:
            tds = tr.find_all('td')
            if not tds:
                continue

            oggetto = link_url = ''
            data_val = ''

            if col['oggetto'] >= 0 and col['oggetto'] < len(tds):
                td = tds[col['oggetto']]
                oggetto = td.get_text(strip=True)
                a = td.find('a') or tr.find('a')
                if a and a.get('href'):
                    link_url = urllib.parse.urljoin(base_url, a['href'])
            else:
                # Fallback: testo più lungo
                texts = [td.get_text(strip=True) for td in tds]
                oggetto = max(texts, key=len) if texts else ''
                a = tr.find('a')
                if a and a.get('href'):
                    link_url = urllib.parse.urljoin(base_url, a['href'])

            if col['data'] >= 0 and col['data'] < len(tds):
                data_val = tds[col['data']].get_text(strip=True)

            if oggetto:
                rows.append({'data': data_val, 'oggetto': oggetto, 'link': link_url})
    return rows


def _has_next(soup) -> bool:
    for a in soup.find_all('a'):
        t = a.get_text(strip=True).lower()
        if any(x in t for x in ['successiv', 'next', 'avanti', '>']):
            return True
    return False


# ─── Regione Calabria ──────────────────────────────────────────────────────────
def scrape_regione_calabria():
    """
    Portale Regione Calabria - Provvedimenti.
    Struttura verificata:
      - GET /provvedimenti-della-regione/ con params:
          filter_date_from (YYYY-MM-DD), filter_date_to (YYYY-MM-DD),
          filter_active=true, paged=N
      - Tabella con colonne: Tipologia, Data Repertoriazione, N, Dipartimento,
          Oggetto, Dettaglio (il link "→" è in questa cella)
      - Paginazione: ul.page-numbers con href contenenti paged=N
    """
    base_url = "https://www.regione.calabria.it/provvedimenti-della-regione/"
    log.info(f"[Regione Calabria] Avvio scraping: {base_url}")
    results = []
    session = make_session()

    page = 1
    max_pages = 200

    while page <= max_pages:
        params = {
            'filter_date_from': DATE_FROM_2,
            'filter_date_to':   DATE_STR_ISO,
            'filter_active':    'true',
            'pr':               '',
            'paged':            page,
        }
        try:
            r = session.get(base_url, params=params, timeout=30, verify=False)
            r.raise_for_status()
        except requests.exceptions.Timeout:
            log.warning(f"[Regione Calabria] Timeout pag.{page}")
            break
        except Exception as e:
            log.warning(f"[Regione Calabria] Errore pag.{page}: {e}")
            break

        soup = BeautifulSoup(r.text, 'lxml')
        table = soup.find('table', class_=re.compile(r'table', re.I))
        if not table:
            log.info(f"[Regione Calabria] Nessuna tabella a pag.{page}")
            break

        # Header detection
        # Struttura verificata del portale Regione Calabria:
        #   Header: <th>Tipologia</th><th>Data Rep.</th><th>N</th><th>Dipartimento</th><th>Oggetto</th><th>Dettaglio</th>
        #   Data row: <td>Decreto</td><td>25/05/2026</td><th scope="row">9104</th><td>Settore...</td><td>Oggetto...</td><td><a>→</a></td>
        # Nota: nella riga dati la colonna N è <th scope="row">, NON <td>.
        # Usiamo find_all(['td','th']) sia per header che per righe dati,
        # così gli indici coincidono e non slittano.

        all_rows = table.find_all('tr')
        if len(all_rows) < 2:
            break

        header_cells = all_rows[0].find_all(['th', 'td'])
        header_texts = [c.get_text(strip=True).lower() for c in header_cells]

        col = {'tipologia': -1, 'data': -1, 'numero': -1, 'dipart': -1, 'oggetto': -1, 'dettaglio': -1}
        for i, h in enumerate(header_texts):
            if 'tipolog' in h:
                col['tipologia'] = i
            elif 'data' in h or 'repertor' in h:
                col['data'] = i
            elif h in ('n', 'numero'):
                col['numero'] = i
            elif 'dipart' in h or 'settore' in h:
                col['dipart'] = i
            elif 'oggetto' in h or 'descriz' in h:
                col['oggetto'] = i
            elif 'dettaglio' in h or 'link' in h or h == '→':
                col['dettaglio'] = i

        page_rows = []
        for tr in table.find_all('tr')[1:]:
            cells = tr.find_all(['td', 'th'])
            if not cells:
                continue

            def cell_text(c):
                if c >= 0 and c < len(cells):
                    return cells[c].get_text(strip=True)
                return ''

            tipologia = cell_text(col['tipologia'])
            data_val  = cell_text(col['data'])
            numero    = cell_text(col['numero'])
            dipart    = cell_text(col['dipart'])
            oggetto   = cell_text(col['oggetto'])

            # Link: cerca in tutte le celle
            link_url = ''
            for cell in cells:
                a = cell.find('a', href=True)
                if a:
                    link_url = urllib.parse.urljoin(base_url, a['href'])
                    break

            # Oggetto completo = tipologia + numero + oggetto
            full_oggetto = ' '.join(filter(None, [tipologia, f"n.{numero}" if numero else '', oggetto]))
            if not full_oggetto:
                continue

            page_rows.append({
                'data':    data_val,
                'oggetto': full_oggetto,
                'link':    link_url,
                '_raw_oggetto': oggetto,
            })

        if not page_rows:
            log.info(f"[Regione Calabria] Nessuna riga a pag.{page}")
            break

        results.extend(page_rows)
        log.info(f"[Regione Calabria] Pag.{page}: {len(page_rows)} atti, totale: {len(results)}")

        # Verifica paginazione: cerca <a class="next page-numbers">
        next_link = soup.find('a', class_=re.compile(r'next.*page-numbers|page-numbers.*next', re.I))
        if not next_link:
            break
        page += 1
        time.sleep(0.3)

    log.info(f"[Regione Calabria] Totale atti: {len(results)}")
    return results, None


# ─── TAR Catanzaro ─────────────────────────────────────────────────────────────
def scrape_tar_catanzaro():
    """
    Portale TAR Catanzaro - giustizia-amministrativa.it
    Struttura: form GET con publishDateFrom/publishDateTo (YYYY-MM-DD o dd/mm/yyyy).
    Colonna 'Parte' contiene il nome delle parti (spesso anonimizzate XXX_OMISSIS_XXX).
    """
    base_url  = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
    log.info(f"[TAR Catanzaro] Avvio scraping: {base_url}")
    results   = []
    session   = make_session()
    error_msg = None

    date_configs = [
        {'publishDateFrom': DATE_FROM_3, 'publishDateTo': DATE_STR_ISO},
        {'publishDateFrom': (TODAY - timedelta(days=3)).strftime('%d/%m/%Y'),
         'publishDateTo':   TODAY.strftime('%d/%m/%Y')},
    ]

    for dc in date_configs:
        try:
            r = session.get(base_url, params=dc, timeout=30, verify=False)
            if r.status_code == 503:
                error_msg = "Portale non raggiungibile (HTTP 503 - possibile blocco IP cloud)"
                log.warning(f"[TAR Catanzaro] {error_msg}")
                return results, error_msg
            r.raise_for_status()
        except requests.exceptions.Timeout:
            error_msg = "Timeout di connessione (portale non raggiungibile dall'ambiente cloud)"
            log.warning(f"[TAR Catanzaro] {error_msg}")
            return results, error_msg
        except Exception as e:
            if hasattr(e, 'response') and e.response is not None and e.response.status_code == 503:
                error_msg = "Portale non raggiungibile (HTTP 503)"
            else:
                error_msg = f"Errore: {e}"
            log.warning(f"[TAR Catanzaro] {error_msg}")
            return results, error_msg

        soup = BeautifulSoup(r.text, 'lxml')
        rows_first = _parse_tar_table(soup, base_url)
        if rows_first:
            results.extend(rows_first)
            log.info(f"[TAR Catanzaro] Pag.1 con config={dc}: {len(rows_first)} atti")

            # Paginazione
            page = 2
            while True:
                next_a = soup.find('a', class_=re.compile(r'next|successiv', re.I))
                if not next_a:
                    break
                next_url = urllib.parse.urljoin(base_url, next_a['href'])
                try:
                    r2 = session.get(next_url, timeout=30, verify=False)
                    r2.raise_for_status()
                    soup = BeautifulSoup(r2.text, 'lxml')
                    more = _parse_tar_table(soup, base_url)
                    if not more:
                        break
                    results.extend(more)
                    log.info(f"[TAR Catanzaro] Pag.{page}: {len(more)} atti")
                    page += 1
                    time.sleep(0.3)
                except Exception:
                    break
            break  # successo con questa config

    log.info(f"[TAR Catanzaro] Totale atti: {len(results)}")
    return results, error_msg if not results else None


def _parse_tar_table(soup, base_url: str):
    rows = []
    for table in soup.find_all('table'):
        all_trs = table.find_all('tr')
        if len(all_trs) < 2:
            continue
        ths = table.find_all('th')
        header_texts = [th.get_text(strip=True).lower() for th in ths]

        col = {'parte': -1, 'data': -1, 'tipo': -1, 'numero': -1}
        for i, h in enumerate(header_texts):
            if 'parte' in h or 'ricorrente' in h or 'appellante' in h:
                col['parte'] = i
            elif 'data' in h or 'deposit' in h or 'pubbl' in h:
                col['data'] = i
            elif 'tipo' in h or 'provvedimento' in h:
                col['tipo'] = i
            elif 'numero' in h or 'n.' in h or h == 'n':
                col['numero'] = i

        for tr in all_trs[1:]:
            tds = tr.find_all('td')
            if not tds:
                continue

            parte  = tds[col['parte']].get_text(strip=True) if col['parte'] >= 0 and col['parte'] < len(tds) else ''
            data_v = tds[col['data']].get_text(strip=True)  if col['data']  >= 0 and col['data']  < len(tds) else ''
            tipo   = tds[col['tipo']].get_text(strip=True)  if col['tipo']  >= 0 and col['tipo']  < len(tds) else ''
            numero = tds[col['numero']].get_text(strip=True) if col['numero'] >= 0 and col['numero'] < len(tds) else ''

            link_url = ''
            for td in tds:
                a = td.find('a', href=True)
                if a:
                    link_url = urllib.parse.urljoin(base_url, a['href'])
                    break

            # Testo principale: Parte è la colonna da filtrare per TAR
            # Oggetto comprensivo
            oggetto = ' | '.join(filter(None, [tipo, f"n.{numero}" if numero else '', parte]))
            if not oggetto:
                oggetto = ' '.join(td.get_text(strip=True) for td in tds)

            rows.append({'data': data_v, 'oggetto': oggetto, 'link': link_url, '_parte': parte})
    return rows


# ─── Excel ─────────────────────────────────────────────────────────────────────
def write_excel(all_data: dict, output_path: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    H_FONT  = Font(bold=True, color='FFFFFF', size=11)
    H_FILL  = PatternFill('solid', fgColor='1F4E79')
    H_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LINK_FONT = Font(color='0563C1', underline='single', size=10)
    WRAP = Alignment(wrap_text=True, vertical='top')
    NORM = Font(size=10)

    for sheet_name, (rows, err_msg) in all_data.items():
        safe = re.sub(r'[/\\?*\[\]:]', '_', sheet_name)[:31]
        ws = wb.create_sheet(title=safe)
        ws.row_dimensions[1].height = 28

        headers = ['Data', 'Oggetto', 'Descrizione breve (150 car.)', 'Link']
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.font = H_FONT; c.fill = H_FILL; c.alignment = H_ALIGN

        if err_msg and not rows:
            c = ws.cell(row=2, column=1, value=f"⚠ {err_msg}")
            c.font = Font(color='C00000', italic=True, size=10)
        elif not rows:
            ws.cell(row=2, column=1, value='Nessun risultato per il periodo selezionato')
        else:
            for ri, row in enumerate(rows, 2):
                data_v  = row.get('data', '')
                oggetto = row.get('oggetto', '')
                descr   = oggetto[:150]
                link    = row.get('link', '')

                c1 = ws.cell(row=ri, column=1, value=data_v)
                c1.font = NORM; c1.alignment = WRAP

                c2 = ws.cell(row=ri, column=2, value=oggetto)
                c2.font = NORM; c2.alignment = WRAP

                c3 = ws.cell(row=ri, column=3, value=descr)
                c3.font = NORM; c3.alignment = WRAP

                c4 = ws.cell(row=ri, column=4)
                if link:
                    c4.value = link
                    c4.hyperlink = link
                    c4.font = LINK_FONT
                else:
                    c4.value = ''
                    c4.font = NORM
                c4.alignment = WRAP

        # Larghezze
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 65
        ws.column_dimensions['C'].width = 42
        ws.column_dimensions['D'].width = 65
        ws.freeze_panes = 'A2'

    wb.save(output_path)
    log.info(f"Salvato: {output_path}")


# ─── Main ──────────────────────────────────────────────────────────────────────
def main():
    log.info(f"Avvio monitoraggio - data odierna: {TODAY.strftime('%d/%m/%Y')}")
    log.info(f"Periodo ASP/RC: {DATE_FROM_2} → {DATE_STR_ISO} | TAR: {DATE_FROM_3} → {DATE_STR_ISO}")

    all_data = {}  # sheet_name -> (matched_rows, error_msg)

    # ── ASP Portals ──
    for asp_name, asp_url in ASP_PORTALS:
        raw, err = scrape_asp_portal(asp_name, asp_url)
        matched = [r for r in raw if keyword_match(r.get('oggetto', ''))]
        log.info(f"[{asp_name}] {len(raw)} atti → {len(matched)} match")
        all_data[asp_name] = (matched, err)

    # ── Regione Calabria ──
    raw_rc, err_rc = scrape_regione_calabria()
    # Per la Regione, filtra su _raw_oggetto se disponibile, altrimenti oggetto
    def rc_match(r):
        obj = r.get('_raw_oggetto') or r.get('oggetto', '')
        return keyword_match(obj)
    matched_rc = [r for r in raw_rc if rc_match(r)]
    log.info(f"[Regione Calabria] {len(raw_rc)} atti → {len(matched_rc)} match")
    all_data['Regione Calabria'] = (matched_rc, err_rc)

    # ── TAR Catanzaro (filtra su _parte) ──
    raw_tar, err_tar = scrape_tar_catanzaro()
    def tar_match(r):
        # TAR: keyword nella colonna Parte
        return keyword_match(r.get('_parte') or r.get('oggetto', ''))
    matched_tar = [r for r in raw_tar if tar_match(r)]
    log.info(f"[TAR Catanzaro] {len(raw_tar)} atti → {len(matched_tar)} match")
    all_data['TAR Catanzaro'] = (matched_tar, err_tar)

    # ── Excel ──
    write_excel(all_data, OUTPUT_PATH)

    # ── Riepilogo ──
    sep = "=" * 62
    print(f"\n{sep}")
    print(f"  MONITORAGGIO ATTI CALABRIA - {TODAY.strftime('%d/%m/%Y')}")
    print(sep)
    total = 0
    for sheet, (rows, err) in all_data.items():
        cnt = len(rows)
        total += cnt
        if err and not rows:
            status = f"⚠ Non raggiungibile"
        else:
            status = f"{cnt} match" if cnt else "Nessun risultato"
        print(f"  {sheet:<26} {status}")
    print(f"  {'TOTALE MATCH':<26} {total}")
    print(sep)
    print(f"\n  File: {OUTPUT_PATH}")
    print(f"  Per uso locale: {OUTPUT_FILENAME}\n")
    return 0


if __name__ == '__main__':
    sys.exit(main())
