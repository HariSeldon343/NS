#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
Scraping: 5 ASP + Regione Calabria + TAR Catanzaro
Output: Excel con un foglio per portale
"""

import re
import json
import time
import warnings
from datetime import date, timedelta

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ─── DATE ──────────────────────────────────────────────────────────────────────
TODAY = date.today()
DATE_FROM_2D = (TODAY - timedelta(days=2)).isoformat()   # 2026-06-13
DATE_FROM_3D = (TODAY - timedelta(days=3)).isoformat()   # 2026-06-12
DATE_TO      = TODAY.isoformat()                          # 2026-06-15

print(f"Data oggi: {TODAY.isoformat()}")
print(f"Periodo 2 giorni: dal {DATE_FROM_2D} al {DATE_TO}")
print(f"Periodo 3 giorni: dal {DATE_FROM_3D} al {DATE_TO}")

# ─── KEYWORDS ──────────────────────────────────────────────────────────────────
KEYWORDS = [
    "ADI",
    "Assistenza domiciliare",
    "ANMIC",
    "Accreditamento",
    "Aumento di budget",
    "Autismo",
    r"Autorizzazione all['’]esercizio",
    "Autorizzazione alla realizzazione",
    "Autorizzazioni",
    "Budget",
    "Casa Giardino",
    "Centro San Giuseppe",
    "Centro salute e benessere",
    "Fabbisogni LEA",
    "Fisiolab",
    "Fisioterapia",
    r"\bLIFE\b",
    "Parere commissione",
    r"Presa d['’]atto verifica",
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

# Pre-compila pattern regex
PATTERNS = [re.compile(kw, re.IGNORECASE | re.UNICODE) for kw in KEYWORDS]


def matches_keywords(text: str) -> bool:
    for pat in PATTERNS:
        if pat.search(text):
            return True
    return False


# ─── HTTP SESSION ──────────────────────────────────────────────────────────────
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "it-IT,it;q=0.9,en;q=0.7",
}


def new_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    return s


# ─── SCRAPER ASP (sisr.regione.calabria.it) ────────────────────────────────────
ASP_PORTALS = {
    "ASP Cosenza":         "https://online-aspco.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Catanzaro":       "https://online-aspcz.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Crotone":         "https://online-aspkr.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Reggio Calabria": "https://online-asprc.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
    "ASP Vibo Valentia":   "https://online-aspvv.sisr.regione.calabria.it/AlboOnline/ricercaAlbo",
}


def scrape_asp(name: str, base_url: str) -> list[dict]:
    """
    Tenta di scaricare atti dall'Albo Online ASP.
    Ritorna lista di dict {data, oggetto, link} oppure
    [{'_error': 'messaggio'}] se il portale non risponde.
    """
    print(f"\n[ASP] Scraping {name} …")
    s = new_session()

    # 1. GET iniziale per cookie e ViewState
    try:
        r = s.get(base_url, timeout=20, verify=False)
        r.raise_for_status()
    except requests.exceptions.Timeout:
        print(f"  TIMEOUT su {base_url}")
        return [{"_error": f"Portale non raggiungibile – timeout (20 s)"}]
    except requests.exceptions.HTTPError as e:
        print(f"  HTTP {e.response.status_code}")
        return [{"_error": f"Portale non raggiungibile – HTTP {e.response.status_code}"}]
    except Exception as e:
        print(f"  ERRORE: {e}")
        return [{"_error": f"Portale non raggiungibile – {e}"}]

    soup = BeautifulSoup(r.text, "lxml")

    # Cerca campo data di pubblicazione
    date_field = soup.find("input", {"id": "dataPubblicazioneDal"})
    if not date_field:
        # Prova name
        date_field = soup.find("input", attrs={"name": re.compile(r"dataPubblicazioneDal", re.I)})

    # Costruisce form data
    form = soup.find("form")
    form_data = {}
    if form:
        for inp in form.find_all(["input", "select", "textarea"]):
            n = inp.get("name")
            v = inp.get("value", "")
            if n:
                form_data[n] = v

    # Imposta data
    if date_field:
        form_data[date_field.get("name") or "dataPubblicazioneDal"] = DATE_FROM_2D
    else:
        form_data["dataPubblicazioneDal"] = DATE_FROM_2D

    # Submit
    action = form.get("action", base_url) if form else base_url
    if action.startswith("/"):
        from urllib.parse import urljoin
        action = urljoin(base_url, action)

    results = []
    page = 0
    seen_first = None

    while True:
        page += 1
        try:
            resp = s.post(action, data=form_data, timeout=20, verify=False)
            resp.raise_for_status()
        except Exception as e:
            if not results:
                return [{"_error": f"Errore POST pagina {page}: {e}"}]
            break

        soup2 = BeautifulSoup(resp.text, "lxml")
        rows = _parse_asp_table(soup2, base_url)

        if not rows:
            break

        # Detect loop (stessa prima riga)
        first_key = rows[0].get("oggetto", "")[:50]
        if first_key == seen_first:
            break
        seen_first = first_key

        results.extend(rows)

        # Paginazione: cerca link "Successiva" o numero pagina
        next_link = (
            soup2.find("a", string=re.compile(r"success", re.I))
            or soup2.find("a", attrs={"title": re.compile(r"success", re.I)})
        )
        if not next_link:
            break

        next_url = next_link.get("href", "")
        if next_url.startswith("/"):
            from urllib.parse import urljoin
            next_url = urljoin(base_url, next_url)

        # Per JSF: potrebbe essere un submit js; se è GET url semplice, usalo
        if next_url.startswith("http"):
            try:
                r2 = s.get(next_url, timeout=20, verify=False)
                soup_tmp = BeautifulSoup(r2.text, "lxml")
                next_rows = _parse_asp_table(soup_tmp, base_url)
                if not next_rows or next_rows[0].get("oggetto", "")[:50] == first_key:
                    break
                results.extend(next_rows)
            except Exception:
                break
        break  # se non riusciamo a paginare, usciamo dopo la prima pagina

    print(f"  → {len(results)} atti totali")
    return results


def _parse_asp_table(soup: BeautifulSoup, base_url: str) -> list[dict]:
    """Estrae righe dalla tabella risultati ASP."""
    table = soup.find("table")
    if not table:
        return []
    rows = table.find_all("tr")[1:]  # skip header
    out = []
    for row in rows:
        cells = row.find_all(["td", "th"])
        if len(cells) < 3:
            continue
        data_val = cells[0].get_text(strip=True)
        oggetto = cells[-2].get_text(strip=True) if len(cells) >= 2 else ""
        link_tag = row.find("a")
        link = ""
        if link_tag and link_tag.get("href"):
            href = link_tag["href"]
            if href.startswith("/"):
                from urllib.parse import urljoin
                href = urljoin(base_url, href)
            link = href
        out.append({"data": data_val, "oggetto": oggetto, "link": link})
    return out


# ─── SCRAPER REGIONE CALABRIA ──────────────────────────────────────────────────
RC_URL = "https://www.regione.calabria.it/provvedimenti-della-regione"


def scrape_regione_calabria() -> list[dict]:
    print("\n[RC] Scraping Regione Calabria …")
    s = new_session()
    results = []
    page_num = 0
    seen_first = None

    while True:
        data = {
            "filter_date_from": DATE_FROM_2D,
            "filter_date_to": DATE_TO,
            "filter_active": "true",
            "sort_order": "",
            "pageNum": str(page_num),
        }
        try:
            r = s.post(RC_URL, data=data, timeout=30, verify=False)
            r.raise_for_status()
        except Exception as e:
            print(f"  Errore pagina {page_num}: {e}")
            break

        soup = BeautifulSoup(r.text, "lxml")
        table = soup.find("table")
        if not table:
            break

        rows = table.find_all("tr")[1:]
        if not rows:
            break

        first_key = rows[0].find_all(["td", "th"])[4].get_text(strip=True)[:50] if len(rows[0].find_all(["td","th"])) >= 5 else ""
        if first_key == seen_first:
            break
        seen_first = first_key

        for row in rows:
            cells = row.find_all(["td", "th"])
            if len(cells) < 6:
                continue
            data_val = cells[1].get_text(strip=True)   # Data Repertoriazione
            oggetto = cells[4].get_text(strip=True)     # Oggetto
            link_tag = cells[5].find("a")
            link = link_tag["href"] if link_tag and link_tag.get("href") else ""
            results.append({"data": data_val, "oggetto": oggetto, "link": link})

        page_num += 1

    print(f"  → {len(results)} atti totali")
    return results


# ─── SCRAPER TAR CATANZARO ────────────────────────────────────────────────────
TAR_BASE = "https://www.giustizia-amministrativa.it/provvedimenti-tar-catanzaro"
TAR_PP = (
    "_it_indra_ga_institutional_area_"
    "JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe_"
)
TAR_AJAX = (
    "https://www.giustizia-amministrativa.it/web/guest/provvedimenti-tar-catanzaro"
    "?p_p_id=it_indra_ga_institutional_area_JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe"
    "&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view"
    "&p_p_resource_id=%2Fadministrative-acts%2Fsearch%2Fresults"
    "&p_p_cacheability=cacheLevelPage"
)
TAR_DOC_BASE = "https://mdp.giustizia-amministrativa.it/visualizza/"


def scrape_tar_catanzaro() -> list[dict]:
    print("\n[TAR] Scraping TAR Catanzaro …")
    s = new_session()

    # Step 1 – GET pagina iniziale
    try:
        r = s.get(TAR_BASE, timeout=30, verify=False)
        r.raise_for_status()
    except Exception as e:
        print(f"  Errore GET: {e}")
        return [{"_error": str(e)}]

    soup = BeautifulSoup(r.text, "lxml")
    form = next(
        (f for f in soup.find_all("form") if "jjYpzZYF4Qfe" in str(f.get("action", ""))),
        None,
    )
    if not form:
        return [{"_error": "Form TAR non trovata"}]

    form_date_inp = form.find("input", {"name": TAR_PP + "formDate"})
    form_date = form_date_inp.get("value", "") if form_date_inp else ""
    action = form.get("action", "")

    # Step 2 – POST form (trigger sessione di ricerca)
    post_data = {
        TAR_PP + "formDate": form_date,
        TAR_PP + "year": "",
        TAR_PP + "number": "",
        TAR_PP + "hearingDateFrom": "",
        TAR_PP + "hearingDateTo": "",
        TAR_PP + "section": "",
        TAR_PP + "type": "",
        TAR_PP + "specific": "",
        TAR_PP + "publishDateFrom": DATE_FROM_3D,
        TAR_PP + "publishDateTo": DATE_TO,
        TAR_PP + "president": "",
        TAR_PP + "draftingJudge": "",
    }
    try:
        s.post(
            action,
            data=post_data,
            headers={"Referer": TAR_BASE, "Origin": "https://www.giustizia-amministrativa.it"},
            timeout=30,
            verify=False,
        )
    except Exception as e:
        print(f"  Errore POST form: {e}")

    # Step 3 – DataTables AJAX (fetch tutto in una sola chiamata)
    cols = [
        "nrgFascicolo", "sezione", "parte", "tipoUdienza",
        "dataUdienza", "numProvvedimento", "dataPubblicazione",
        "tipoProvvedimento", "relatore", "presidente", "esito",
    ]
    additional_info = json.dumps({
        "schema": "TAR_CATANZARO",
        "type": None, "year": "", "number": "",
        "hearingDateFrom": None, "hearingDateTo": None,
        "publishDateFrom": DATE_FROM_3D, "publishDateTo": DATE_TO,
        "hearingType": None, "nrg": None,
        "section": "", "provisionSpecification": "",
        "president": "", "draftingJudge": "",
        "subjectMatter": None, "page": None, "size": None,
        "orderBy": None, "orderStrategy": None, "queryString": None,
    })
    dt_payload = {
        "draw": 1,
        "columns": [
            {"data": c, "name": "", "searchable": True, "orderable": True,
             "search": {"value": "", "regex": False}}
            for c in cols
        ],
        "order": [{"column": 6, "dir": "desc"}, {"column": 5, "dir": "desc"}],
        "start": 0,
        "length": 1000,
        "search": {"value": "", "regex": False},
        "additionalInfo": additional_info,
    }

    ajax_headers = {
        "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": TAR_BASE,
    }
    try:
        resp = s.post(
            TAR_AJAX,
            json=dt_payload,
            headers=ajax_headers,
            timeout=30,
            verify=False,
        )
        resp.raise_for_status()
        js = resp.json()
    except Exception as e:
        print(f"  Errore AJAX: {e}")
        return [{"_error": str(e)}]

    raw = js.get("data", [])
    print(f"  → {len(raw)} record totali (recordsFiltered={js.get('recordsFiltered',0)})")

    results = []
    for rec in raw:
        nrg = rec.get("nrgFascicolo", "")
        nome_file = rec.get("nomeFile", "")
        parte = rec.get("parte", "")
        data_pub = rec.get("dataPubblicazione", "")
        tipo = rec.get("tipoProvvedimento", "")
        num_prov = rec.get("numProvvedimento", "")

        doc_link = (
            f"{TAR_DOC_BASE}?nodeRef=&schema=tar_cz"
            f"&nrg={nrg}&nomeFile={nome_file}&subDir=Provvedimenti"
            if nome_file else TAR_BASE
        )
        oggetto_tar = f"{tipo} n.{num_prov} – Parte: {parte} – NRG: {nrg}"

        results.append({
            "data": data_pub,
            "oggetto": oggetto_tar,
            "parte": parte,
            "link": doc_link,
        })

    return results


# ─── EXCEL OUTPUT ─────────────────────────────────────────────────────────────
HYPERLINK_FONT = Font(color="0563C1", underline="single")
HEADER_FILL   = PatternFill("solid", fgColor="4472C4")
HEADER_FONT   = Font(color="FFFFFF", bold=True)
ALT_FILL      = PatternFill("solid", fgColor="DCE6F1")


def write_sheet(ws, rows: list[dict], is_tar: bool = False):
    """Popola un foglio Excel con i risultati."""
    headers = ["Data", "Oggetto", "Descrizione breve", "Link"]
    col_widths = [15, 60, 40, 50]

    # Intestazioni
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 20

    # Controlla se ci sono errori
    if rows and "_error" in rows[0]:
        ws.cell(row=2, column=1, value=rows[0]["_error"])
        ws.cell(row=2, column=1).font = Font(color="FF0000", italic=True)
        ws.merge_cells("A2:D2")
        _set_col_widths(ws, col_widths)
        return

    # Filtra per keyword
    if is_tar:
        filtered = [r for r in rows if matches_keywords(r.get("parte", ""))]
    else:
        filtered = [r for r in rows if matches_keywords(r.get("oggetto", ""))]

    if not filtered:
        ws.cell(row=2, column=1, value="Nessun risultato")
        ws.merge_cells("A2:D2")
        _set_col_widths(ws, col_widths)
        return

    for row_idx, rec in enumerate(filtered, 2):
        fill = ALT_FILL if row_idx % 2 == 0 else None

        data_val = rec.get("data", "")
        oggetto  = rec.get("oggetto", "")
        desc     = oggetto[:150]
        link     = rec.get("link", "")

        c_data = ws.cell(row=row_idx, column=1, value=data_val)
        c_obj  = ws.cell(row=row_idx, column=2, value=oggetto)
        c_desc = ws.cell(row=row_idx, column=3, value=desc)
        c_link = ws.cell(row=row_idx, column=4, value="Apri atto")

        # Hyperlink cliccabile
        if link:
            c_link.hyperlink = link
            c_link.font = HYPERLINK_FONT
            c_link.value = "Apri atto"
        else:
            c_link.value = ""

        for cell in [c_data, c_obj, c_desc, c_link]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if fill:
                cell.fill = fill

    _set_col_widths(ws, col_widths)


def _set_col_widths(ws, widths):
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # rimuove foglio vuoto di default

    # — ASP portali —
    for sheet_name, url in ASP_PORTALS.items():
        ws = wb.create_sheet(sheet_name)
        data = scrape_asp(sheet_name, url)
        write_sheet(ws, data)

    # — Regione Calabria —
    ws_rc = wb.create_sheet("Regione Calabria")
    rc_data = scrape_regione_calabria()
    write_sheet(ws_rc, rc_data)

    # — TAR Catanzaro —
    ws_tar = wb.create_sheet("TAR Catanzaro")
    tar_data = scrape_tar_catanzaro()
    write_sheet(ws_tar, tar_data, is_tar=True)

    # Salva file
    date_str = TODAY.strftime("%d-%m-%Y")
    filename = f"Monitoraggio_Atti_Calabria_{date_str}.xlsx"
    out_path = f"/home/user/NS/{filename}"
    wb.save(out_path)
    print(f"\n✓ File salvato: {out_path}")
    return out_path, filename


if __name__ == "__main__":
    out_path, filename = main()
