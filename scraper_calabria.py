#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraper per portali ASP, Regione Calabria e TAR Catanzaro
"""

import re
import json
import time
import requests
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Configurazione date ────────────────────────────────────────────────────────
TODAY = datetime.now()
DATE_FROM_2D = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_FROM_3D = (TODAY - timedelta(days=3)).strftime("%Y-%m-%d")
DATE_TO = TODAY.strftime("%Y-%m-%d")
DATE_FROM_DD_MM = (TODAY - timedelta(days=2)).strftime("%d/%m/%Y")
EXECUTION_DATE = TODAY.strftime("%d-%m-%Y")

print(f"Esecuzione: {TODAY.strftime('%d/%m/%Y %H:%M')}")
print(f"Filtro date (2gg): dal {DATE_FROM_2D} al {DATE_TO}")
print(f"Filtro date TAR (3gg): dal {DATE_FROM_3D}")

# ── Keywords ──────────────────────────────────────────────────────────────────
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
# "Life" con word boundary (evita falsi positivi dentro parole composte)
LIFE_PATTERN = re.compile(r'\bLIFE\b', re.IGNORECASE)


def matches_keywords(text: str) -> bool:
    """True se il testo contiene almeno una keyword (case-insensitive)."""
    upper = text.upper()
    for kw in KEYWORDS:
        if kw.upper() in upper:
            return True
    if LIFE_PATTERN.search(text):
        return True
    return False


# ── Sessione HTTP comune ──────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0.0.0 Safari/537.36"),
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.8",
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)


# ═══════════════════════════════════════════════════════════════════════════════
# SCRAPER ASP (portali SISR — rete regionale)
# ═══════════════════════════════════════════════════════════════════════════════
ASP_PORTALS = {
    "ASP Cosenza":        "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":      "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":        "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria":"https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia":  "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}


def scrape_asp(portal_name: str, base_url: str) -> list[dict]:
    """
    Esegue la ricerca sul portale AlboOnline ASP.
    Restituisce lista di dict {data, oggetto, link}.
    Restituisce None se il portale è irraggiungibile.
    """
    print(f"\n  [{portal_name}] Connessione a {base_url} ...")

    # 1) GET pagina per ottenere VIEWSTATE e altri hidden fields
    try:
        r = SESSION.get(base_url, timeout=25)
    except Exception as e:
        print(f"    ERROR: {e}")
        return None

    if r.status_code == 503:
        print(f"    503 – Portale non raggiungibile (rete SISR privata)")
        return None
    if r.status_code != 200:
        print(f"    HTTP {r.status_code}")
        return None

    soup = BeautifulSoup(r.text, "lxml")

    # Raccoglie tutti gli hidden field del form
    form = soup.find("form")
    if not form:
        print("    Form non trovato")
        return None

    form_data = {}
    for inp in form.find_all("input"):
        name = inp.get("name", "")
        val = inp.get("value", "")
        if name:
            form_data[name] = val

    # Imposta il filtro data
    form_data["dataPubblicazioneDal"] = DATE_FROM_2D
    form_data["dataPubblicazioneAl"] = DATE_TO

    # Determina action URL
    action = form.get("action", "")
    if not action.startswith("http"):
        from urllib.parse import urljoin
        action = urljoin(base_url, action)
    method = form.get("method", "post").upper()

    all_results = []
    page = 1

    while True:
        print(f"    Pagina {page}...", end=" ")
        try:
            if method == "POST":
                resp = SESSION.post(action, data=form_data, timeout=30)
            else:
                resp = SESSION.get(action, params=form_data, timeout=30)
        except Exception as e:
            print(f"ERRORE: {e}")
            break

        if resp.status_code != 200:
            print(f"HTTP {resp.status_code}")
            break

        soup2 = BeautifulSoup(resp.text, "lxml")
        rows = _parse_asp_table(soup2, base_url)
        print(f"{len(rows)} righe trovate")
        all_results.extend(rows)

        # Cerca link pagina successiva
        next_url = _find_next_asp_page(soup2, base_url, page)
        if not next_url:
            break

        # Per la pagina successiva potremmo dover aggiornare form_data
        # (dipende dall'implementazione — tentativamente aggiorniamo i viewstate)
        soup_next = BeautifulSoup(resp.text, "lxml")
        for inp in soup_next.find_all("input", type="hidden"):
            name = inp.get("name", "")
            if name:
                form_data[name] = inp.get("value", "")
        form_data["dataPubblicazioneDal"] = DATE_FROM_2D
        form_data["dataPubblicazioneAl"] = DATE_TO

        # Imposta paginazione se c'è un parametro pagina
        if "paginaCorrente" in form_data:
            form_data["paginaCorrente"] = str(page + 1)
        elif "page" in form_data:
            form_data["page"] = str(page + 1)

        page += 1
        time.sleep(0.5)

    return all_results


def _parse_asp_table(soup: BeautifulSoup, base_url: str) -> list[dict]:
    """Estrae le righe della tabella risultati dal portale ASP."""
    results = []
    table = soup.find("table", {"class": re.compile(r"table|result", re.I)})
    if not table:
        table = soup.find("table")
    if not table:
        return results

    rows = table.find_all("tr")
    for row in rows[1:]:  # salta header
        cells = row.find_all(["td", "th"])
        if len(cells) < 2:
            continue

        texts = [c.get_text(strip=True) for c in cells]

        # Cerca la data (formato DD/MM/YYYY o YYYY-MM-DD)
        data = ""
        oggetto = ""
        link = ""

        for t in texts:
            if re.match(r"\d{2}/\d{2}/\d{4}", t) or re.match(r"\d{4}-\d{2}-\d{2}", t):
                data = t
            elif len(t) > 10 and not data:
                oggetto = t

        # Se non trovati con euristica, usa le prime celle significative
        if not oggetto and len(texts) >= 2:
            data = texts[0]
            oggetto = texts[-2] if len(texts) >= 3 else texts[1]

        # Link all'atto
        a_tag = row.find("a", href=True)
        if a_tag:
            href = a_tag["href"]
            if not href.startswith("http"):
                from urllib.parse import urljoin
                href = urljoin(base_url, href)
            link = href

        if oggetto:
            results.append({"data": data, "oggetto": oggetto, "link": link})

    return results


def _find_next_asp_page(soup: BeautifulSoup, base_url: str, current_page: int) -> str | None:
    """Cerca il link alla pagina successiva nell'impaginazione ASP."""
    # Cerca pattern comuni di paginazione
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True)
        href = a["href"]
        if text in ("Successiva", "Avanti", "»", ">", "Next", str(current_page + 1)):
            if not href.startswith("http"):
                from urllib.parse import urljoin
                href = urljoin(base_url, href)
            return href
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# SCRAPER REGIONE CALABRIA
# ═══════════════════════════════════════════════════════════════════════════════
REGIONE_BASE = "https://www.regione.calabria.it/provvedimenti-della-regione/"


def scrape_regione() -> list[dict]:
    """Scarica tutti i provvedimenti della Regione degli ultimi 2 giorni."""
    print("\n  [Regione Calabria] Avvio scraping ...")
    all_results = []
    page = 1

    while True:
        url = REGIONE_BASE
        params = {
            "filter_date_from": DATE_FROM_2D,
            "filter_date_to": DATE_TO,
        }
        if page > 1:
            params["paged"] = page

        print(f"    Pagina {page}...", end=" ")
        try:
            r = SESSION.get(url, params=params, timeout=30)
        except Exception as e:
            print(f"ERRORE: {e}")
            break

        if r.status_code != 200:
            print(f"HTTP {r.status_code}")
            break

        soup = BeautifulSoup(r.text, "lxml")
        rows, has_more = _parse_regione_table(soup)
        print(f"{len(rows)} righe")
        all_results.extend(rows)

        if not has_more:
            break

        # Controlla che ci sia effettivamente la pagina successiva
        next_link = soup.find("a", class_="next page-numbers")
        if not next_link:
            break

        page += 1
        time.sleep(0.3)

    return all_results


def _parse_regione_table(soup: BeautifulSoup) -> tuple[list[dict], bool]:
    """Estrae righe dalla tabella provvedimenti Regione Calabria."""
    results = []
    table = soup.find("table", class_=re.compile(r"table", re.I))
    if not table:
        return results, False

    rows = table.find_all("tr")
    for row in rows[1:]:  # salta header
        cells = row.find_all(["td", "th"])
        if len(cells) < 5:
            continue

        tipologia = cells[0].get_text(strip=True)
        data_rep = cells[1].get_text(strip=True)
        # numero = cells[2].get_text(strip=True)
        # dip = cells[3].get_text(strip=True)
        oggetto = cells[4].get_text(strip=True)

        # Link (ultima cella o cella dettaglio)
        link = ""
        a_tag = row.find("a", href=True)
        if a_tag:
            href = a_tag["href"]
            link = href if href.startswith("http") else "https://www.regione.calabria.it" + href

        # Converti data da GG/MM/AAAA a YYYY-MM-DD per confronto
        data_iso = _parse_date_it(data_rep)

        if oggetto:
            results.append({
                "data": data_rep,
                "data_iso": data_iso,
                "oggetto": oggetto,
                "tipologia": tipologia,
                "link": link,
            })

    has_more = bool(soup.find("a", class_="next page-numbers"))
    return results, has_more


def _parse_date_it(date_str: str) -> str:
    """Converte data italiana GG/MM/AAAA in YYYY-MM-DD."""
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", date_str)
    if m:
        return f"{m.group(3)}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    return date_str


# ═══════════════════════════════════════════════════════════════════════════════
# SCRAPER TAR CATANZARO
# ═══════════════════════════════════════════════════════════════════════════════
TAR_HOME = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
NS = "_it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_"
TAR_COLS = [
    "nrgFascicolo", "sezione", "parte", "tipoUdienza", "dataUdienza",
    "numProvvedimento", "dataPubblicazione", "tipoProvvedimento",
    "relatore", "presidente", "esito",
]


def _get_tar_search_url() -> tuple[str, str, str]:
    """
    Carica la pagina TAR per estrarre p_auth, formDate e searchURL freschi.
    Restituisce (search_url, p_auth, form_date).
    """
    print("    Ottenimento token TAR...", end=" ")
    r = SESSION.get(TAR_HOME, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"TAR home HTTP {r.status_code}")

    auth_m = re.search(r'p_auth=([a-zA-Z0-9]+)', r.text)
    fd_m = re.search(r'INSTANCE_jjYpzZYF4Qfe_formDate".*?value="(\d+)"', r.text, re.DOTALL)
    url_m = re.search(r'var searchURL\s*=\s*"([^"]+)"', r.text)

    p_auth = auth_m.group(1) if auth_m else ""
    form_date = fd_m.group(1) if fd_m else str(int(time.time() * 1000))
    search_url = url_m.group(1) if url_m else (
        "https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro"
        "?p_p_id=it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe"
        "&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view"
        "&p_p_resource_id=%2Fadministrative-acts%2Fsearch%2Fresults&p_p_cacheability=cacheLevelPage"
    )
    print(f"p_auth={p_auth}")
    return search_url, p_auth, form_date


def _tar_detail_url(nrg: str, nome_file: str) -> str:
    """Costruisce URL di dettaglio TAR usando il viewer ufficiale MDP."""
    if nome_file:
        return (
            f"https://mdp.giustizia-amministrativa.it/visualizza/"
            f"?nodeRef=&schema=tar_cz&nrg={nrg}&nomeFile={nome_file}&subDir=Provvedimenti"
        )
    base = "https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro"
    return f"{base}?nrg={nrg}"


def scrape_tar() -> list[dict]:
    """Scarica i provvedimenti TAR Catanzaro degli ultimi 3 giorni via DataTables API."""
    print("\n  [TAR Catanzaro] Avvio scraping ...")

    search_url, p_auth, form_date = _get_tar_search_url()

    all_data = []
    start = 0
    page_size = 100

    tar_headers = {
        "Content-Type": "application/json",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Origin": "https://www.giustizia-amministrativa.it",
        "Referer": TAR_HOME,
        "X-Requested-With": "XMLHttpRequest",
    }

    additional_info = {
        "schema": "TAR_CATANZARO",
        "type": None,
        "year": "",
        "number": "",
        "hearingDateFrom": None,
        "hearingDateTo": None,
        "publishDateFrom": DATE_FROM_3D,
        "publishDateTo": DATE_TO,
        "hearingType": None,
        "nrg": None,
        "section": "",
        "provisionSpecification": "",
        "president": "",
        "draftingJudge": "",
        "subjectMatter": None,
        "page": None,
        "size": None,
        "orderBy": None,
        "orderStrategy": None,
        "queryString": None,
    }

    page = 1
    total_records = None

    while True:
        print(f"    Pagina {page} (start={start})...", end=" ")
        payload = {
            "draw": page,
            "columns": [
                {
                    "data": c,
                    "name": "",
                    "searchable": True,
                    "orderable": True,
                    "search": {"value": "", "regex": False},
                }
                for c in TAR_COLS
            ],
            "order": [{"column": 6, "dir": "desc"}],
            "start": start,
            "length": page_size,
            "search": {"value": "", "regex": False},
            "additionalInfo": json.dumps(additional_info),
        }

        try:
            resp = SESSION.post(
                search_url,
                json=payload,
                headers=tar_headers,
                timeout=30,
            )
        except Exception as e:
            print(f"ERRORE: {e}")
            break

        if resp.status_code != 200:
            print(f"HTTP {resp.status_code}")
            break

        data = resp.json()
        records = data.get("data", [])
        print(f"{len(records)} record")

        if total_records is None:
            total_records = data.get("recordsFiltered", 0)
            print(f"    Totale TAR: {total_records} record nel periodo")

        for rec in records:
            num_prov = rec.get("numProvvedimento", "")
            tipo_prov = rec.get("tipoProvvedimento", "")
            sezione = rec.get("sezione", "")
            nome_file = rec.get("nomeFile", "")

            # Costruzione link al documento
            link = _tar_detail_url(rec.get("nrgFascicolo", ""), nome_file)

            # Per TAR l'oggetto è nella colonna "parte" + tipo provvedimento
            parte = rec.get("parte", "")
            oggetto_tar = f"[{tipo_prov}] NRG {rec.get('nrgFascicolo','')} - {parte}"

            all_data.append({
                "data": rec.get("dataPubblicazione", ""),
                "nrg": rec.get("nrgFascicolo", ""),
                "sezione": sezione,
                "parte": parte,
                "oggetto": oggetto_tar,
                "tipo": tipo_prov,
                "num_provvedimento": num_prov,
                "link": link,
                "nome_file": nome_file,
            })

        start += len(records)
        if start >= total_records or len(records) == 0:
            break

        page += 1
        time.sleep(0.3)

    return all_data


# ═══════════════════════════════════════════════════════════════════════════════
# GENERAZIONE EXCEL
# ═══════════════════════════════════════════════════════════════════════════════
BLUE_FONT = Font(color="0000FF", underline="single")
HEADER_FILL = PatternFill(start_color="17324D", end_color="17324D", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
GREY_FILL = PatternFill(start_color="EEF0F6", end_color="EEF0F6", fill_type="solid")


def _write_sheet(ws, rows: list[dict] | None, sheet_type: str = "asp"):
    """Scrive un foglio Excel con i risultati."""
    if sheet_type == "asp":
        headers = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]
    elif sheet_type == "regione":
        headers = ["Data", "Tipologia", "Oggetto", "Descrizione breve (150 car.)", "Link"]
    else:  # tar
        headers = ["Data pubbl.", "NRG", "Sezione", "Parte", "Tipo provvedimento",
                   "N. Provvedimento", "Descrizione breve (150 car.)", "Link"]

    # Header row
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    ws.row_dimensions[1].height = 25

    if rows is None:
        ws.cell(row=2, column=1, value="⚠️ Portale non raggiungibile (rete SISR privata – HTTP 503)")
        ws.column_dimensions["A"].width = 70
        return

    if not rows:
        ws.cell(row=2, column=1, value="Nessun risultato")
        ws.column_dimensions["A"].width = 40
        return

    for row_idx, rec in enumerate(rows, 2):
        fill = GREY_FILL if row_idx % 2 == 0 else None

        if sheet_type == "asp":
            data_vals = [
                rec.get("data", ""),
                rec.get("oggetto", ""),
                rec.get("oggetto", "")[:150],
                rec.get("link", ""),
            ]
            link_col = 4
        elif sheet_type == "regione":
            data_vals = [
                rec.get("data", ""),
                rec.get("tipologia", ""),
                rec.get("oggetto", ""),
                rec.get("oggetto", "")[:150],
                rec.get("link", ""),
            ]
            link_col = 5
        else:  # tar
            data_vals = [
                rec.get("data", ""),
                rec.get("nrg", ""),
                rec.get("sezione", ""),
                rec.get("parte", ""),
                rec.get("tipo", ""),
                rec.get("num_provvedimento", ""),
                rec.get("oggetto", "")[:150],
                rec.get("link", ""),
            ]
            link_col = 8

        for col_idx, val in enumerate(data_vals, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if fill:
                cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")

        # Imposta hyperlink cliccabile sulla cella link
        link_url = rec.get("link", "")
        if link_url:
            link_cell = ws.cell(row=row_idx, column=link_col)
            link_cell.value = link_url
            link_cell.hyperlink = link_url
            link_cell.font = BLUE_FONT

    # Larghezze colonne
    if sheet_type == "asp":
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 60
        ws.column_dimensions["C"].width = 55
        ws.column_dimensions["D"].width = 80
    elif sheet_type == "regione":
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 65
        ws.column_dimensions["D"].width = 55
        ws.column_dimensions["E"].width = 80
    else:
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 10
        ws.column_dimensions["D"].width = 50
        ws.column_dimensions["E"].width = 28
        ws.column_dimensions["F"].width = 18
        ws.column_dimensions["G"].width = 55
        ws.column_dimensions["H"].width = 80

    # Freeze header
    ws.freeze_panes = "A2"


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # rimuovi foglio default

    # ── ASP Portals ──────────────────────────────────────────────────────────
    print("\n═══ PORTALI ASP ═══")
    for portal_name, portal_url in ASP_PORTALS.items():
        print(f"\n► {portal_name}")
        raw = scrape_asp(portal_name, portal_url)

        ws = wb.create_sheet(title=portal_name)

        if raw is None:
            # Portale irraggiungibile
            _write_sheet(ws, None, sheet_type="asp")
            print(f"  → Foglio '{portal_name}': portale non raggiungibile")
        else:
            # Filtra per keywords
            matched = [r for r in raw if matches_keywords(r.get("oggetto", ""))]
            print(f"  → {len(raw)} atti totali, {len(matched)} con keyword match")
            _write_sheet(ws, matched, sheet_type="asp")

    # ── Regione Calabria ─────────────────────────────────────────────────────
    print("\n═══ REGIONE CALABRIA ═══")
    try:
        raw_rc = scrape_regione()
        matched_rc = [r for r in raw_rc if matches_keywords(r.get("oggetto", ""))]
        print(f"  → {len(raw_rc)} atti totali, {len(matched_rc)} con keyword match")
    except Exception as e:
        print(f"  ERRORE: {e}")
        raw_rc = []
        matched_rc = []

    ws_rc = wb.create_sheet(title="Regione Calabria")
    _write_sheet(ws_rc, matched_rc, sheet_type="regione")

    # ── TAR Catanzaro ────────────────────────────────────────────────────────
    print("\n═══ TAR CATANZARO ═══")
    try:
        raw_tar = scrape_tar()
        # Per TAR la keyword si cerca sulla colonna "parte" e "oggetto"
        matched_tar = []
        for r in raw_tar:
            combined = r.get("parte", "") + " " + r.get("oggetto", "")
            if matches_keywords(combined):
                matched_tar.append(r)
        print(f"  → {len(raw_tar)} atti totali, {len(matched_tar)} con keyword match")
    except Exception as e:
        print(f"  ERRORE: {e}")
        raw_tar = []
        matched_tar = []

    ws_tar = wb.create_sheet(title="TAR Catanzaro")
    _write_sheet(ws_tar, matched_tar, sheet_type="tar")

    # ── Salvataggio ──────────────────────────────────────────────────────────
    output_filename = f"Monitoraggio_Atti_Calabria_{EXECUTION_DATE}.xlsx"
    output_path = f"/home/user/NS/{output_filename}"

    wb.save(output_path)
    print(f"\n✔ File salvato: {output_path}")
    print(f"\nRiepilogo:")
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        n = ws.max_row - 1
        val = ws.cell(2, 1).value or ""
        if "non raggiungibile" in str(val) or "Nessun risultato" in str(val):
            print(f"  {sheet}: {val}")
        else:
            print(f"  {sheet}: {n} risultati con match")


if __name__ == "__main__":
    main()
