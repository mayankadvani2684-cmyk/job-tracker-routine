#!/usr/bin/env python3
"""
LinkedIn job scraper → Excel report.
Uses Apify curious_coder/linkedin-jobs-scraper.
Writes XLSX without openpyxl (stdlib zipfile + XML only).

Actor field names confirmed from live test:
  companyName, link, postedAt (YYYY-MM-DD), seniorityLevel, employmentType, location
"""
import io, os, re, sys, time, json, zipfile, requests
from datetime import date, datetime

# ─── Config ──────────────────────────────────────────────────────────────────

APIFY_TOKEN = os.environ.get("APIFY_API_TOKEN", "")
ACTOR       = "curious_coder~linkedin-jobs-scraper"

BASE = "https://www.linkedin.com/jobs/search/?f_TPR=r86400&"

KEYWORDS = [
    "FP%26A%20analyst",
    "strategy%20analyst",
    "business%20finance%20analyst",
    "credit%20analyst",
    "real%20estate%20analyst",
    "valuation%20analyst",
    "deals%20analyst",
    "transaction%20advisory%20analyst",
    "corporate%20finance%20analyst",
]
LOCATIONS = [
    "Mumbai%2C%20Maharashtra%2C%20India",
    "Delhi%2C%20India",
    "Bengaluru%2C%20Karnataka%2C%20India",
]
OMIT = {
    ("business%20finance%20analyst",      "Delhi%2C%20India"),
    ("deals%20analyst",                   "Bengaluru%2C%20Karnataka%2C%20India"),
    ("transaction%20advisory%20analyst",  "Bengaluru%2C%20Karnataka%2C%20India"),
    ("corporate%20finance%20analyst",     "Bengaluru%2C%20Karnataka%2C%20India"),
    ("valuation%20analyst",               "Bengaluru%2C%20Karnataka%2C%20India"),
}

# Actor accepts plain URL strings (not objects)
START_URLS = [
    f"{BASE}keywords={kw}&location={loc}"
    for kw in KEYWORDS
    for loc in LOCATIONS
    if (kw, loc) not in OMIT
]

# ─── Apify ────────────────────────────────────────────────────────────────────

def apify_post(path, **kwargs):
    r = requests.post(f"https://api.apify.com/v2/{path}",
                      params={"token": APIFY_TOKEN}, timeout=60, **kwargs)
    r.raise_for_status()
    return r.json()

def apify_get(path, extra_params=None):
    params = {"token": APIFY_TOKEN}
    if extra_params:
        params.update(extra_params)
    r = requests.get(f"https://api.apify.com/v2/{path}", params=params, timeout=120)
    r.raise_for_status()
    return r.json()

def run_actor():
    data = apify_post(f"acts/{ACTOR}/runs",
                      json={"urls": START_URLS, "count": 50, "scrapeCompany": False})
    run_id = data["data"]["id"]
    print(f"Run started: {run_id}", flush=True)
    return run_id

def wait_for_run(run_id):
    for attempt in range(240):
        d = apify_get(f"actor-runs/{run_id}")["data"]
        status = d["status"]
        print(f"  [{attempt*30}s] status={status}", flush=True)
        if status == "SUCCEEDED":
            return d["defaultDatasetId"]
        if status in ("FAILED", "ABORTED", "TIMED-OUT"):
            sys.exit(f"Actor run ended: {status}")
        time.sleep(30)
    sys.exit("Timed out waiting for actor")

def fetch_dataset(dataset_id):
    items, offset, limit = [], 0, 1000
    while True:
        d = apify_get(f"datasets/{dataset_id}/items",
                      extra_params={"limit": limit, "offset": offset, "format": "json"})
        batch = d if isinstance(d, list) else d.get("items", d.get("data", []))
        if not batch:
            break
        items.extend(batch)
        if len(batch) < limit:
            break
        offset += limit
    return items

# ─── Filtering ────────────────────────────────────────────────────────────────

BAD_SENIORITY = re.compile(
    r'\b(director|vp|vice.?president|c-suite|chief|ceo|cfo|coo|cto|internship|intern)\b', re.I)
BAD_TYPE      = re.compile(r'\b(internship|contract)\b', re.I)
BAD_TITLE     = re.compile(
    r'administrative|executive assistant|\bea to\b|front office|guest relations|'
    r'compliance officer|sharepoint|\berp\b|\br2r\b|record to report|vat analyst|'
    r'production controller|chartered accountant|sales executive|'
    r'office administration|process associate', re.I)
BAD_COMPANIES = {"mygwork", "scoutit", "alignerr", "datamark", "aditi consulting"}

def keep(job):
    title     = job.get("title", "")
    company   = (job.get("companyName") or job.get("company") or "").strip()
    seniority = job.get("seniorityLevel", "") or ""
    emp_type  = job.get("employmentType", "") or ""
    if BAD_SENIORITY.search(seniority):       return False
    if BAD_TYPE.search(emp_type):             return False
    if BAD_TITLE.search(title):               return False
    if company.lower() in BAD_COMPANIES:      return False
    return True

# ─── Tier & Resume ───────────────────────────────────────────────────────────

IT_RE = re.compile(
    r'\btcs\b|tata consultancy|infosys|\bwipro\b|\bhcl\b|hcltech|tech mahindra|'
    r'mphasis|l&t infotech|ltimindtree|\bmindtree\b|hexaware|niit tech|'
    r'persistent systems|accenture|cognizant|capgemini|\bibm\b|'
    r'zensar|cyient|\bkpit\b|mastech|birlasoft|sonata software|coforge|eclerx|'
    r'\bepam\b|firstsource|\bwns\b|igate|patni|'
    r'\bamazon\b|\bgoogle\b|\bmicrosoft\b|\bapple\b|\bmeta\b|\bnetflix\b|'
    r'flipkart|paytm|\bola\b|\buber\b|\bzomato\b|swiggy|byju|unacademy|'
    r'razorpay|phonepe|freshworks|\bzoho\b|'
    r'software solutions|it consulting|digital services', re.I)

FIN_RE = re.compile(
    r'\bbank\b|\bnbfc\b|financial services|finance|capital|investment|securities|'
    r'insurance|asset management|wealth|advisory|brokerage|'
    r'private equity|venture capital|hedge fund|mutual fund|'
    r'kotak|hdfc|icici|\baxis\b|\bsbi\b|yes bank|indusind|\brbl\b|'
    r'bajaj|tata capital|aditya birla|l&t finance|shriram|'
    r'jm financial|motilal|edelweiss|\biifl\b|angel broking|zerodha|'
    r'nomura|goldman sachs|morgan stanley|barclays|deutsche|citi|hsbc|'
    r'standard chartered|jp morgan|credit suisse|\bubs\b|macquarie|'
    r'rothschild|lazard|mckinsey|bain &|bcg|kearney|roland berger|'
    r'deloitte|pwc|\bkpmg\b|\bey\b|ernst & young|grant thornton|'
    r'management consulting|'
    r'northern trust|bnp paribas|paribas|cr[eé]dit agricole|agricole|'
    r'société générale|socgen|natwest|lloyds|rbs|santander|anz\b|'
    r'fidelity|blackrock|vanguard|state street|wellington|'
    r'pimco|bridgewater|citadel|two sigma|renaissance|'
    r'd\.?\s*e\.?\s*shaw|de shaw|nuvama|nuveen|'
    r'xl dynamics|xfactrs|ifc\b|adb\b|world bank|'
    r'ambit|emkay|systematix|prabhudas|anand rathi|dsp\b|'
    r'muthoot|manappuram|chola|sundaram|'
    r'aon|mercer|willis|marsh|gallagher|'
    r'kroll|duff & phelps|alvarez|houlihan|lazard|evercore|moelis|'
    r'pa consulting|oliver wyman|'
    r'stock.*exchang|nse\b|bse\b|sebi\b', re.I)

TIER1_T = re.compile(
    r'fp&a|fpa|financial planning|business finance|'
    r'credit analyst|real estate analyst|real estate finance|'
    r'valuation analyst|deals analyst|corporate finance analyst|'
    r'transaction advisory|'
    r'corporate development|business finance analyst|'
    r'finance business partner|fp\s*&\s*a', re.I)

TIER2_T = re.compile(
    r'\bfinancial analyst\b|\bfinance analyst\b|strategy analyst|'
    r'investment analyst|research analyst|equity analyst|'
    r'\bib\b|investment banking|senior analyst|associate analyst|'
    r'fund anal|portfolio anal', re.I)

def assign_tier(job):
    title   = job.get("title", "")
    company = (job.get("companyName") or job.get("company") or "")
    is_it   = bool(IT_RE.search(company))
    is_fin  = bool(FIN_RE.search(company))
    if TIER1_T.search(title) and not is_it:
        return 1
    if TIER2_T.search(title) and is_fin:
        return 2
    return 3

RE_PAT     = re.compile(r'real estate|valuation|development finance|property|asset management', re.I)
ECM_PAT    = re.compile(r'\bib\b|tas\b|deals|m&a|structured finance|capital markets', re.I)
STRAT_PAT  = re.compile(r'fp&a|fpa|business finance|strategy|corporate finance|financial planning', re.I)
CREDIT_PAT = re.compile(r'credit analyst', re.I)
FINAN_PAT  = re.compile(r'financial analyst|finance analyst', re.I)

def assign_resume(job):
    title   = job.get("title", "")
    company = (job.get("companyName") or job.get("company") or "")
    is_fin  = bool(FIN_RE.search(company))
    if CREDIT_PAT.search(title): return "RE/Strategy"
    if RE_PAT.search(title):     return "RE Resume"
    if ECM_PAT.search(title):    return "ECM Resume"
    if STRAT_PAT.search(title):  return "Strategy Resume"
    if FINAN_PAT.search(title) and is_fin: return "Strategy/ECM"
    return "Strategy Resume"

def extract_city(loc):
    if not loc: return ""
    # "Mumbai Metropolitan Region" → "Mumbai Metropolitan Region" (keep as-is, trim at comma)
    return loc.split(",")[0].strip()

def parse_posted(job):
    raw = (job.get("postedAt") or job.get("publishedAt") or
           job.get("posted_at") or job.get("date") or "")
    if not raw:
        return None
    raw = str(raw).strip()
    try:
        return datetime.strptime(raw[:10], "%Y-%m-%d").date()
    except Exception:
        return None

def get_job_url(job):
    return (job.get("link") or job.get("jobUrl") or
            job.get("job_url") or job.get("url") or "")

def get_company(job):
    return (job.get("companyName") or job.get("company") or "")

# ─── Minimal XLSX writer ─────────────────────────────────────────────────────

def xml_escape(s):
    s = str(s)
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;")
             .replace("'", "&apos;"))

def date_to_excel(d):
    if d is None:
        return None
    return (d - date(1899, 12, 30)).days

# Style indices for cellXfs:
# 0: Arial10, no fill               → default text (tier3 / no fill)
# 1: Arial10 bold white, header fill → header
# 2: Arial10 link blue underline, no fill → link (tier3)
# 3: Arial10, tier1 fill            → tier1 text
# 4: Arial10, tier2 fill            → tier2 text
# 5: Arial10 link, tier1 fill       → link (tier1)
# 6: Arial10 link, tier2 fill       → link (tier2)
# 7: Arial10, no fill, date fmt     → date (tier3)
# 8: Arial10, tier1 fill, date fmt  → date (tier1)
# 9: Arial10, tier2 fill, date fmt  → date (tier2)

STYLES_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="YYYY-MM-DD"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="10"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><name val="Arial"/><color rgb="FFFFFFFF"/></font>
    <font><sz val="10"/><name val="Arial"/><color rgb="FF0563C1"/><u val="single"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC8E6C9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF9C4"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0"   fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0"   fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0"   fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0"   fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0"   fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>
  </cellXfs>
</styleSheet>
"""

def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

COLS   = ["Rank", "Job Title", "Company", "City", "Resume", "Tier", "Link", "Status", "Posted"]
WIDTHS = [6, 48, 22, 16, 20, 6, 10, 12, 13]

def _style(col_idx, tier, is_date=False):
    """Return cellXf index for a data cell."""
    if col_idx == 7:              # Link column
        return {1: 5, 2: 6}.get(tier, 2)
    if is_date:
        return {1: 8, 2: 9}.get(tier, 7)
    return {1: 3, 2: 4}.get(tier, 0)

def sheet_xml(rows):
    col_defs = "".join(
        f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>'
        for i, w in enumerate(WIDTHS)
    )

    def c_str(row, col, val, style):
        addr = f"{col_letter(col)}{row}"
        return f'<c r="{addr}" s="{style}"><v>{xml_escape(val)}</v></c>'

    def c_inline(row, col, val, style):
        addr = f"{col_letter(col)}{row}"
        return f'<c r="{addr}" s="{style}" t="inlineStr"><is><t>{xml_escape(val)}</t></is></c>'

    # Header row
    hdr = "".join(c_inline(1, ci, h, 1) for ci, h in enumerate(COLS, 1))
    row_xmls = [f'<row r="1" ht="28" customHeight="1">{hdr}</row>']

    hyperlinks = []
    hl_id = 1

    for ri, row in enumerate(rows, 2):
        tier = row.get("tier", 3)
        cells = ""

        cells += c_str(ri, 1, str(row.get("rank", "")), _style(1, tier))
        cells += c_inline(ri, 2, row.get("title", ""), _style(2, tier))
        cells += c_inline(ri, 3, row.get("company", ""), _style(3, tier))
        cells += c_inline(ri, 4, row.get("city", ""), _style(4, tier))
        cells += c_inline(ri, 5, row.get("resume", ""), _style(5, tier))
        cells += c_str(ri, 6, str(tier), _style(6, tier))

        # Link column (7)
        addr7  = f"{col_letter(7)}{ri}"
        job_url = row.get("url", "")
        lk_style = _style(7, tier)
        if job_url:
            rel_id = f"rId{hl_id}"; hl_id += 1
            hyperlinks.append((addr7, job_url, row.get("title", ""), rel_id))
        cells += c_inline(ri, 7, "Open", lk_style)

        # Status (8)
        cells += c_inline(ri, 8, "", _style(8, tier))

        # Posted (9)
        d = row.get("posted_date")
        dt_style = _style(9, tier, is_date=True)
        if d is not None:
            cells += c_str(ri, 9, str(date_to_excel(d)), dt_style)
        else:
            cells += c_inline(ri, 9, "", dt_style)

        row_xmls.append(f'<row r="{ri}" ht="20" customHeight="1">{cells}</row>')

    hl_xml = ""
    if hyperlinks:
        parts = [
            f'<hyperlink ref="{addr}" r:id="{rid}" tooltip="{xml_escape(tip)}"/>'
            for addr, _, tip, rid in hyperlinks
        ]
        hl_xml = "<hyperlinks>" + "".join(parts) + "</hyperlinks>"

    rels_xml = ""
    if hyperlinks:
        parts = [
            f'<Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            f'Target="{xml_escape(url)}" TargetMode="External"/>'
            for _, url, _, rid in hyperlinks
        ]
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(parts) + "</Relationships>"
        )

    ws = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheetViews><sheetView workbookViewId=\"0\">"
        "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
        "</sheetView></sheetViews>"
        f"<cols>{col_defs}</cols>"
        "<sheetData>" + "\n".join(row_xmls) + "</sheetData>"
        + hl_xml +
        "</worksheet>"
    )
    return ws, rels_xml

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>'
)

RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)

def workbook_xml(names):
    sheets = "".join(
        f'<sheet name="{xml_escape(n)}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        for i, n in enumerate(names)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets}</sheets></workbook>"
    )

def workbook_rels_xml(n):
    parts = [
        f'<Relationship Id="rId{i+1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{i+1}.xml"/>'
        for i in range(n)
    ]
    parts.append(
        f'<Relationship Id="rId{n+1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        f'Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(parts) + "</Relationships>"
    )

def build_xlsx(all_rows):
    priority = []
    for i, r in enumerate([x for x in all_rows if x["tier"] in (1, 2)], 1):
        priority.append({**r, "rank": i})

    s1_xml, s1_rels = sheet_xml(priority)
    s2_xml, s2_rels = sheet_xml(all_rows)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("xl/workbook.xml", workbook_xml(["Priority Jobs", "All Filtered Jobs"]))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(2))
        zf.writestr("xl/styles.xml", STYLES_XML)
        zf.writestr("xl/worksheets/sheet1.xml", s1_xml)
        zf.writestr("xl/worksheets/sheet2.xml", s2_xml)
        if s1_rels:
            zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", s1_rels)
        if s2_rels:
            zf.writestr("xl/worksheets/_rels/sheet2.xml.rels", s2_rels)
    return buf.getvalue()

# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    if not APIFY_TOKEN:
        sys.exit("APIFY_API_TOKEN not set")

    print(f"Generated {len(START_URLS)} search URLs", flush=True)

    # Allow re-use of existing dataset via env var for quick rebuilds
    dataset_id = os.environ.get("REUSE_DATASET_ID", "")
    if dataset_id:
        print(f"Re-using dataset: {dataset_id}", flush=True)
    else:
        run_id     = run_actor()
        dataset_id = wait_for_run(run_id)
    print(f"Dataset: {dataset_id}", flush=True)

    raw = fetch_dataset(dataset_id)
    print(f"Raw jobs fetched: {len(raw)}", flush=True)
    if raw:
        print("Sample keys:", list(raw[0].keys())[:15], flush=True)

    # Filter
    filtered = [j for j in raw if keep(j)]
    print(f"After filter: {len(filtered)}", flush=True)

    # Dedup by company+title
    seen, deduped = set(), []
    for j in filtered:
        key = (get_company(j).lower().strip(), j.get("title", "").lower().strip())
        if key not in seen:
            seen.add(key)
            deduped.append(j)
    print(f"After dedup: {len(deduped)}", flush=True)

    # Enrich
    for j in deduped:
        j["_tier"]   = assign_tier(j)
        j["_resume"] = assign_resume(j)
        j["_city"]   = extract_city(j.get("location", ""))
        j["_posted"] = parse_posted(j)
        j["_url"]    = get_job_url(j)

    # Sort: tier → company
    deduped.sort(key=lambda j: (j["_tier"], get_company(j).lower()))

    # Build row dicts (rank by overall order)
    rows = []
    for i, j in enumerate(deduped, 1):
        rows.append({
            "rank":        i,
            "title":       j.get("title", ""),
            "company":     get_company(j),
            "city":        j["_city"],
            "resume":      j["_resume"],
            "tier":        j["_tier"],
            "url":         j["_url"],
            "posted_date": j["_posted"],
        })

    today    = datetime.now().strftime("%Y-%m-%d")
    filename = f"LinkedIn_Jobs_{today}.xlsx"
    path     = f"/home/user/job-tracker-routine/{filename}"

    xlsx_bytes = build_xlsx(rows)
    with open(path, "wb") as f:
        f.write(xlsx_bytes)
    print(f"Saved {len(xlsx_bytes):,} bytes → {path}", flush=True)

    meta = {"filename": filename, "path": path}
    with open("/home/user/job-tracker-routine/.xlsx_meta.json", "w") as f:
        json.dump(meta, f)

    print("DONE", flush=True)

if __name__ == "__main__":
    main()
