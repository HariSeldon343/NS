#!/usr/bin/env python3
"""
Monitoraggio Atti Calabria
- 5 portali ASP (sisr.regione.calabria.it) → irraggiungibili da rete remota (SSL/IP)
- Regione Calabria (regione.calabria.it)    → requests GET, tabella HTML
- TAR Catanzaro (giustizia-amministrativa.it) → AJAX DataTables JSON API
"""

import re, os, json
from datetime import datetime, timedelta
from urllib.parse import urljoin
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
TODAY      = datetime(2026, 6, 25)
DATE_2D    = (TODAY - timedelta(days=2)).strftime("%Y-%m-%d")
DATE_3D    = (TODAY - timedelta(days=3)).strftime("%Y-%m-%d")
TODAY_STR  = TODAY.strftime("%Y-%m-%d")
OUTPUT_DIR = "/home/user/NS"
OUTPUT_FILE = os.path.join(
    OUTPUT_DIR, f"Monitoraggio_Atti_Calabria_{TODAY.strftime('%d-%m-%Y')}.xlsx"
)
CA_BUNDLE = "/root/.ccr/ca-bundle.crt"
PROXIES = {
    "http":  os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY"),
    "https": os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY"),
}
BASE_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

# ── KEYWORDS ──────────────────────────────────────────────────────────────────
KEYWORDS_PLAIN = [
    "ADI", "Assistenza domiciliare", "ANMIC", "Accreditamento",
    "Aumento di budget", "Autismo",
    "Autorizzazione all'esercizio", "Autorizzazione alla realizzazione",
    "Autorizzazioni", "Budget",
    "Casa Giardino", "Centro San Giuseppe", "Centro salute e benessere",
    "Fabbisogni LEA", "Fisiolab", "Fisioterapia",
    "Parere commissione", "Presa d'atto verifica", "Programmazione",
    "Rete riabilitativa", "Rete territoriale",
    "Riabilitazione estensiva", "Riconversione prestazioni",
    "Rinnovo accreditamento",
    "San Teodoro", "Savelli Hospital", "Starbene",
    "Verifica requisiti", "Villa San Giuseppe", "Villa del Rosario",
]
LIFE_RE = re.compile(r'\bLIFE\b', re.IGNORECASE)

def kw_match(text: str) -> bool:
    if not text:
        return False
    tu = text.upper()
    for kw in KEYWORDS_PLAIN:
        if kw.upper() in tu:
            return True
    return bool(LIFE_RE.search(text))

def strip_tags(html: str) -> str:
    return re.sub(r'<[^>]+>', '', html or '').strip()

def make_abs(href: str, base: str) -> str:
    if not href:
        return ""
    return href if href.startswith("http") else urljoin(base, href)

def new_session() -> requests.Session:
    s = requests.Session()
    s.verify  = CA_BUNDLE
    s.proxies = PROXIES
    s.headers.update(BASE_HEADERS)
    return s


# ═══════════════════════════════════════════════════════════════════════════════
# REGIONE CALABRIA  (GET form → HTML table)
# ═══════════════════════════════════════════════════════════════════════════════
RC_BASE = "https://www.regione.calabria.it"
RC_PATH = "/provvedimenti-della-regione/"

def _rc_parse(html: str) -> list:
    """Return (data, oggetto, link) tuples from table."""
    results = []
    table_m = re.search(
        r'<table\b[^>]*class=["\'][^"\']*table[^"\']*["\'][^>]*>(.*?)</table>',
        html, re.DOTALL | re.IGNORECASE
    )
    if not table_m:
        return results
    tbody_m = re.search(r'<tbody>(.*?)</tbody>', table_m.group(1), re.DOTALL | re.IGNORECASE)
    if not tbody_m:
        return results

    for row in re.findall(r'<tr\b[^>]*>(.*?)</tr>', tbody_m.group(1), re.DOTALL | re.IGNORECASE):
        cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL | re.IGNORECASE)
        if len(cells) < 4:
            continue
        # Cols: Tipologia | Data Repertoriazione | Dipartimento | Oggetto | [DataPub] | →
        data_rep = strip_tags(cells[1])
        oggetto  = strip_tags(cells[3])
        link_m   = re.search(r'href=["\']([^"\']+)["\']', row)
        link     = make_abs(link_m.group(1) if link_m else "", RC_BASE)
        if kw_match(oggetto):
            results.append({"data": data_rep, "oggetto": oggetto, "link": link})
    return results

def scrape_regione() -> list:
    print(f"\n[Regione Calabria] Filtro: {DATE_2D} → {TODAY_STR}")
    sess = new_session()
    url  = RC_BASE + RC_PATH
    params = {
        "filter_date_from": DATE_2D,
        "filter_date_to":   TODAY_STR,
        "sort_order": "0",
    }
    results  = []
    page_num = 0

    while url:
        page_num += 1
        try:
            r = sess.get(url, params=(params if page_num == 1 else None), timeout=25)
            r.raise_for_status()
        except Exception as e:
            print(f"  Page {page_num} error: {e}")
            break

        matched = _rc_parse(r.text)
        results.extend(matched)
        print(f"  Page {page_num}: {len(matched)} match(es) (totale {len(results)})")

        url = None
        for pat in [
            r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>\s*(?:Successivo|›|&rsaquo;|&gt;)\s*</a>',
            r'<a[^>]+rel=["\']next["\'][^>]+href=["\']([^"\']+)["\']',
        ]:
            m = re.search(pat, r.text, re.IGNORECASE)
            if m:
                url = make_abs(m.group(1), RC_BASE)
                break

    print(f"[Regione Calabria] Totale match: {len(results)}")
    return results


# ═══════════════════════════════════════════════════════════════════════════════
# TAR CATANZARO  (AJAX DataTables JSON – p_p_lifecycle=2)
# ═══════════════════════════════════════════════════════════════════════════════
TAR_BASE   = "https://www.giustizia-amministrativa.it"
TAR_PAGE   = "/provvedimenti-tar-catanzaro"
TAR_NS     = ("it_indra_ga_institutional_area_"
              "JurisdictionalActivityAdministrativeActsWebPortlet_INSTANCE_jjYpzZYF4Qfe")
TAR_COLUMNS = [
    "nrgFascicolo", "sezione", "parte", "tipoUdienza",
    "dataUdienza", "numProvvedimento", "dataPubblicazione",
    "tipoProvvedimento", "relatore", "presidente", "esito",
]

def _tar_doc_url(num_provv: str, nome_file: str) -> str:
    """Build best-effort document URL for a TAR act."""
    if nome_file:
        return (f"{TAR_BASE}/portale/pages/istituzionale/visualizza"
                f"?codice={num_provv}")
    return f"{TAR_BASE}{TAR_PAGE}"

def scrape_tar() -> list:
    print(f"\n[TAR Catanzaro] Filtro: {DATE_3D} → {TODAY_STR}")
    sess = new_session()

    # Establish session (cookie + JSESSIONID)
    try:
        sess.get(TAR_BASE + TAR_PAGE, timeout=25)
    except Exception as e:
        print(f"  Session init error: {e}")
        return []

    ajax_url = (
        f"{TAR_BASE}/web/guest/provvedimenti-tar-catanzaro"
        f"?p_p_id={TAR_NS}"
        f"&p_p_lifecycle=2&p_p_state=normal&p_p_mode=view"
        f"&p_p_resource_id=%2Fadministrative-acts%2Fsearch%2Fresults"
        f"&p_p_cacheability=cacheLevelPage"
    )

    results  = []
    start    = 0
    length   = 200   # fetch in batches of 200
    draw     = 0

    while True:
        draw += 1
        additional_info = {
            "schema": "TAR_CATANZARO",
            "type": None, "year": "", "number": None,
            "hearingDateFrom": None, "hearingDateTo": None,
            "publishDateFrom": DATE_3D, "publishDateTo": TODAY_STR,
            "hearingType": None, "nrg": None, "section": "",
            "provisionSpecification": "", "president": None,
            "draftingJudge": None, "subjectMatter": None,
            "page": None, "size": None, "orderBy": None,
            "orderStrategy": None, "queryString": None,
        }
        payload = {
            "draw": draw,
            "columns": [
                {"data": col, "name": "", "searchable": True, "orderable": True,
                 "search": {"value": "", "regex": False}}
                for col in TAR_COLUMNS
            ],
            "order": [{"column": 6, "dir": "desc"}, {"column": 5, "dir": "desc"}],
            "start":  start,
            "length": length,
            "search": {"value": "", "regex": False},
            "additionalInfo": json.dumps(additional_info),
        }

        try:
            r = sess.post(
                ajax_url,
                data=json.dumps(payload),
                headers={
                    "Content-Type":     "application/json",
                    "X-Requested-With": "XMLHttpRequest",
                    "Accept":           "application/json",
                    "Referer":          TAR_BASE + TAR_PAGE,
                    "Origin":           TAR_BASE,
                },
                timeout=30,
            )
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"  AJAX error (start={start}): {e}")
            break

        records = data.get("data", [])
        total   = data.get("recordsTotal", 0)
        print(f"  Batch start={start}: {len(records)} records (total={total})")

        for rec in records:
            parte = rec.get("parte", "") or ""
            if kw_match(parte):
                results.append({
                    "data":    rec.get("dataPubblicazione", ""),
                    "oggetto": parte,
                    "link":    _tar_doc_url(
                        rec.get("numProvvedimento", ""),
                        rec.get("nomeFile", ""),
                    ),
                })

        start += len(records)
        if start >= total or not records:
            break

    print(f"[TAR Catanzaro] Totale match: {len(results)}")
    return results


# ═══════════════════════════════════════════════════════════════════════════════
# XLSX WRITER
# ═══════════════════════════════════════════════════════════════════════════════
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
LNK_FONT = Font(color="0563C1", underline="single", name="Calibri", size=10)
DAT_FONT = Font(name="Calibri", size=10)
ERR_FONT = Font(italic=True, color="808080", name="Calibri", size=10)
THIN     = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

SHEET_ORDER = [
    "ASP Cosenza", "ASP Catanzaro", "ASP Crotone",
    "ASP Reggio Calabria", "ASP Vibo Valentia",
    "Regione Calabria", "TAR Catanzaro",
]
ASP_NOTE = (
    "Portale non raggiungibile dall'ambiente di esecuzione remota "
    "(SSL_ERROR_SYSCALL su sisr.regione.calabria.it – possibile restrizione IP "
    "a range italiani o mTLS richiesto). Verificare manualmente da browser locale."
)

def write_xlsx(all_data: dict):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for sheet_name in SHEET_ORDER:
        rows = all_data.get(sheet_name)
        ws   = wb.create_sheet(title=sheet_name[:31])

        if rows is None:
            ws["A1"] = ASP_NOTE
            ws["A1"].font = ERR_FONT
            ws.column_dimensions["A"].width = 90
            continue

        if not rows:
            ws["A1"] = "Nessun risultato nel periodo"
            ws["A1"].font = ERR_FONT
            ws.column_dimensions["A"].width = 40
            continue

        headers    = ["Data", "Oggetto", "Descrizione breve (150 car.)", "Link"]
        col_widths = [18, 65, 55, 55]
        for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.font      = HDR_FONT
            c.fill      = HDR_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = THIN
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[1].height = 20
        ws.freeze_panes = "A2"

        for ri, rec in enumerate(rows, 2):
            data    = rec.get("data", "")
            oggetto = rec.get("oggetto", "")
            desc    = oggetto[:150]
            link    = rec.get("link", "")

            for ci, val in enumerate([data, oggetto, desc, link], 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.font      = DAT_FONT
                c.border    = THIN
                c.alignment = Alignment(vertical="top", wrap_text=(ci == 2))

            if link:
                lc           = ws.cell(row=ri, column=4)
                lc.value     = link
                lc.hyperlink = link
                lc.font      = LNK_FONT

        ws.auto_filter.ref = ws.dimensions

    wb.save(OUTPUT_FILE)
    print(f"\n✅ Salvato: {OUTPUT_FILE}")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    print("╔══ Monitoraggio Atti Calabria ══════════════════════════════════╗")
    print(f"║  Esecuzione: {TODAY.strftime('%d/%m/%Y')}")
    print(f"║  Filtro 2gg:  {DATE_2D} → {TODAY_STR}")
    print(f"║  Filtro 3gg:  {DATE_3D} → {TODAY_STR}")
    print("╚════════════════════════════════════════════════════════════════╝\n")

    all_data = {}

    for name in ["ASP Cosenza", "ASP Catanzaro", "ASP Crotone",
                 "ASP Reggio Calabria", "ASP Vibo Valentia"]:
        print(f"[{name}] ⚠  sisr.regione.calabria.it non raggiungibile (SSL/IP restriction)")
        all_data[name] = None

    all_data["Regione Calabria"] = scrape_regione()
    all_data["TAR Catanzaro"]    = scrape_tar()

    print("\n╔══ RIEPILOGO ═══════════════════════════════════════════════════╗")
    for sh in SHEET_ORDER:
        v = all_data.get(sh)
        if v is None:
            print(f"║  {sh:<25} ⚠  Non raggiungibile")
        else:
            print(f"║  {sh:<25} ✓  {len(v)} match(es)")
    print("╚════════════════════════════════════════════════════════════════╝")

    write_xlsx(all_data)


if __name__ == "__main__":
    main()
