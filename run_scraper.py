#!/usr/bin/env python3
"""
LinkedIn Jobs Pipeline
  1. Calls Apify curious_coder/linkedin-jobs-scraper
  2. Filters / classifies results
  3. Writes LinkedIn_Jobs_YYYY-MM-DD.xlsx (manual xlsx writer – no openpyxl)
"""
import io, json, os, re, sys, time, zipfile
from datetime import date as Date, datetime
from xml.sax.saxutils import escape as xe

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
import requests

APIFY_TOKEN = os.environ["APIFY_API_TOKEN"]
ACTOR_ID    = "curious_coder~linkedin-jobs-scraper"
TODAY       = datetime.now().strftime("%Y-%m-%d")
OUTPUT      = f"/home/user/job-tracker-routine/LinkedIn_Jobs_{TODAY}.xlsx"

# ─────────────────────────────────────────────────────────────────────────────
# URL GENERATION
# ─────────────────────────────────────────────────────────────────────────────
BASE = "https://www.linkedin.com/jobs/search/?f_TPR=r86400&"
KEYWORDS = [
    "FP%26A%20analyst", "strategy%20analyst", "business%20finance%20analyst",
    "credit%20analyst", "real%20estate%20analyst", "valuation%20analyst",
    "deals%20analyst", "transaction%20advisory%20analyst", "corporate%20finance%20analyst",
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
SEARCH_URLS = [
    f"{BASE}keywords={kw}&location={loc}"
    for kw in KEYWORDS for loc in LOCATIONS
    if (kw, loc) not in OMIT
]
print(f"Search URLs generated: {len(SEARCH_URLS)}")

# ─────────────────────────────────────────────────────────────────────────────
# APIFY RUN
# ─────────────────────────────────────────────────────────────────────────────
actor_input = {
    "urls":          SEARCH_URLS,
    "count":         50,
    "scrapeCompany": False,
}

print("Starting Apify actor…")
r = requests.post(
    f"https://api.apify.com/v2/acts/{ACTOR_ID}/runs",
    params={"token": APIFY_TOKEN},
    json=actor_input,
    timeout=60,
)
r.raise_for_status()
run       = r.json()["data"]
run_id    = run["id"]
ds_id     = run["defaultDatasetId"]
print(f"  run_id={run_id}  dataset={ds_id}")

# poll
for _ in range(150):          # up to 50 minutes
    time.sleep(20)
    st = requests.get(
        f"https://api.apify.com/v2/actor-runs/{run_id}",
        params={"token": APIFY_TOKEN}, timeout=30,
    ).json()["data"]["status"]
    print(f"  {st}", flush=True)
    if st in ("SUCCEEDED", "FAILED", "ABORTED", "TIMED-OUT"):
        break
if st != "SUCCEEDED":
    sys.exit(f"Run ended: {st}")

items = requests.get(
    f"https://api.apify.com/v2/datasets/{ds_id}/items",
    params={"token": APIFY_TOKEN, "limit": 100000},
    timeout=120,
).json()
print(f"Raw items: {len(items)}")
if items:
    print("Keys:", list(items[0].keys())[:25])
    print("Sample:\n" + json.dumps(items[0], indent=2, default=str)[:1200])

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def gf(job, *keys):
    for k in keys:
        v = job.get(k)
        if v is not None and v != "":
            return str(v).strip()
    return ""

def parse_date(s):
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%SZ",
                "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:len(fmt.replace("%Y","0000").replace("%m","00")
                                        .replace("%d","00").replace("%H","00")
                                        .replace("%M","00").replace("%S","00")
                                        .replace("%f","000000").replace("%Z",""))], fmt).date()
        except Exception:
            pass
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except Exception:
        return None

def city_from(loc):
    if not loc:
        return ""
    for c in ("Mumbai","Delhi","Bengaluru","Bangalore","Pune",
              "Hyderabad","Chennai","Gurugram","Gurgaon","Noida"):
        if c.lower() in loc.lower():
            return "Bengaluru" if c == "Bangalore" else c
    return loc.split(",")[0].strip()

# ─────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────
IT_RE = re.compile(
    r"\b(TCS|Tata Consultancy|Infosys|Wipro|HCL Tech|Tech Mahindra|Accenture|Cognizant|"
    r"Capgemini|IBM|Oracle|SAP|Unisys|Mphasis|Hexaware|NIIT Technologies|"
    r"Mindtree|LTIMindtree|L&T Infotech|Zensar|Persistent|Cyient)\b", re.I)

FINSERV_RE = re.compile(
    r"\b(Bank|Banking|NBFC|Finance|Financial|Capital|Investment|Securities|"
    r"Asset Management|Wealth|Insurance|Deloitte|KPMG|EY|Ernst|PwC|"
    r"PricewaterhouseCoopers|McKinsey|BCG|Boston Consulting|Bain|"
    r"Alvarez|Avendus|Kotak|HDFC|ICICI|Axis|Yes Bank|IndusInd|RBL|SBI|"
    r"State Bank|Citi|HSBC|Barclays|JPMorgan|Morgan Stanley|Goldman|"
    r"Deutsche Bank|Nomura|Ambit|IIFL|Motilal Oswal|Edelweiss|Nuvama|"
    r"360 One|Anand Rathi|Mirae|Nippon|DSP|Sundaram|Tata Capital|"
    r"Bajaj Finance|Mahindra Finance|Muthoot|Shriram|JM Financial|"
    r"ICRA|CRISIL|CARE Ratings|India Ratings)\b", re.I)

TIER1_RE = re.compile(
    r"fp.?a|business finance|credit analyst|real estate|valuation|"
    r"\bdeals?\b|corporate finance|transaction advisory", re.I)

TIER2_RE = re.compile(
    r"financial analyst|strategy analyst|investment analyst|ib analyst|"
    r"senior analyst|associate analyst|research analyst", re.I)

TITLE_EXCL = re.compile(
    r"Administrative|Executive Assistant|\bEA to\b|Front Office|Guest Relations|"
    r"Compliance Officer|SharePoint|\bERP\b|\bR2R\b|Record to Report|VAT Analyst|"
    r"Production Controller|Chartered Accountant|Sales Executive|"
    r"Office Administration|Process Associate", re.I)

SENIORITY_EXCL = re.compile(r"Director|Vice President|\bVP\b|C-Suite|Chief|Internship", re.I)
TYPE_EXCL      = re.compile(r"Internship|Contract", re.I)
CO_EXCL        = {"mygwork","scoutit","alignerr","datamark","aditi consulting"}

def assign_resume(title, company):
    t, c = title.lower(), company.lower()
    if re.search(r"real estate|valuation|development finance|property|asset management", t):
        return "RE Resume"
    if re.search(r"\btas\b|\bib\b|investment bank|deals|m&a|m & a|structured finance|capital markets|transaction advisory", t):
        return "ECM Resume"
    if re.search(r"fp.?a|business finance|strategy|corporate finance|financial planning", t):
        return "Strategy Resume"
    if re.search(r"financial analyst", t) and FINSERV_RE.search(company):
        return "Strategy/ECM"
    if re.search(r"credit analyst", t):
        return "RE/Strategy"
    return "Strategy Resume"

def assign_tier(title, company):
    if TIER1_RE.search(title) and not IT_RE.search(company):
        return 1
    if TIER2_RE.search(title) and FINSERV_RE.search(company):
        return 2
    return 3

# ─────────────────────────────────────────────────────────────────────────────
# FILTER
# ─────────────────────────────────────────────────────────────────────────────
seen, filtered = set(), []
for job in items:
    title    = gf(job, "title","jobTitle","position","name")
    company  = gf(job, "companyName","company","employer","hiringOrganization")
    loc      = gf(job, "location","jobLocation","city","locationName")
    seniority= gf(job, "seniorityLevel","seniority","jobLevel","experienceLevel","level")
    emp_type = gf(job, "employmentType","jobType","type","contractType")
    url      = gf(job, "link","applyUrl","url","jobUrl","externalApplyLink","jobPostingUrl")
    posted_r = gf(job, "postedAt","publishedAt","datePosted","listedAt","posted","postingDate")

    if not title or not company:
        continue
    key = (company.lower(), title.lower())
    if key in seen:
        continue
    if SENIORITY_EXCL.search(seniority) or SENIORITY_EXCL.search(title):
        continue
    if TYPE_EXCL.search(emp_type):
        continue
    if TITLE_EXCL.search(title):
        continue
    if company.lower() in CO_EXCL:
        continue

    seen.add(key)
    city   = city_from(loc)
    posted = parse_date(posted_r)
    resume = assign_resume(title, company)
    tier   = assign_tier(title, company)
    filtered.append(dict(title=title, company=company, city=city,
                         resume=resume, tier=tier, url=url,
                         posted=posted, status="", rank=0))

filtered.sort(key=lambda x: (x["tier"], x["company"].lower()))
for i, j in enumerate(filtered, 1):
    j["rank"] = i

t1 = sum(1 for j in filtered if j["tier"]==1)
t2 = sum(1 for j in filtered if j["tier"]==2)
t3 = sum(1 for j in filtered if j["tier"]==3)
print(f"Filtered: {len(filtered)}  Tier1={t1}  Tier2={t2}  Tier3={t3}")

# ─────────────────────────────────────────────────────────────────────────────
# XLSX WRITER  (pure zipfile / XML — no openpyxl)
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_EPOCH = Date(1899, 12, 30)
def date_serial(d):
    return (d - EXCEL_EPOCH).days

def col_letter(n):           # 1-based → A, B, …, Z, AA, …
    s = ""
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def cell_ref(row, col):      # both 1-based
    return f"{col_letter(col)}{row}"

# ── Style constants ──────────────────────────────────────────────────────────
# Font indexes:  0=Arial10  1=Arial10 bold white  2=Arial10 blue underline
# Fill indexes:  0=none  1=gray125  2=C8E6C9  3=FFF9C4  4=1F4E79  5=FFFFFF
# numFmt 164 = "YYYY-MM-DD"
#
# XF index mapping:
XF_DEFAULT  = 0   # Arial10 / no fill
XF_HEADER   = 1   # bold white / navy / center
XF_T1_DATA  = 2   # Arial10 / green
XF_T2_DATA  = 3   # Arial10 / yellow
XF_T3_DATA  = 4   # Arial10 / white
XF_T1_LINK  = 5   # blue underline / green
XF_T2_LINK  = 6   # blue underline / yellow
XF_T3_LINK  = 7   # blue underline / white
XF_T1_DATE  = 8   # Arial10 / green  / YYYY-MM-DD
XF_T2_DATE  = 9   # Arial10 / yellow / YYYY-MM-DD
XF_T3_DATE  = 10  # Arial10 / white  / YYYY-MM-DD

def xf_for_tier(tier):          return [XF_T1_DATA, XF_T2_DATA, XF_T3_DATA][tier-1]
def xf_link_for_tier(tier):     return [XF_T1_LINK, XF_T2_LINK, XF_T3_LINK][tier-1]
def xf_date_for_tier(tier):     return [XF_T1_DATE, XF_T2_DATE, XF_T3_DATE][tier-1]

def styles_xml():
    return """\
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
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC8E6C9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF9C4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="11">
    <xf numFmtId="0"   fontId="0" fillId="5" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="1" fillId="4" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="0" fillId="2" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="0" fillId="3" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="0" fillId="5" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="2" fillId="2" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="2" fillId="3" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="2" fillId="5" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="5" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"""

def content_types_xml(n_sheets):
    overrides = ""
    for i in range(1, n_sheets+1):
        overrides += f'  <Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{overrides}  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

def root_rels_xml():
    return """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

def workbook_xml(sheet_names):
    sheets = "".join(
        f'    <sheet name="{xe(n)}" sheetId="{i}" r:id="rId{i}"/>\n'
        for i, n in enumerate(sheet_names, 1)
    )
    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
{sheets}  </sheets>
</workbook>"""

def workbook_rels_xml(n_sheets):
    rels = ""
    for i in range(1, n_sheets+1):
        rels += f'  <Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>\n'
    rels += f'  <Relationship Id="rId{n_sheets+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>\n'
    rels += f'  <Relationship Id="rId{n_sheets+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>\n'
    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels}</Relationships>"""

def shared_strings_xml(strings):
    items = "".join(f"  <si><t>{xe(s)}</t></si>\n" for s in strings)
    n = len(strings)
    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{n}" uniqueCount="{n}">
{items}</sst>"""

# ── Sheet builder ─────────────────────────────────────────────────────────────
HEADERS    = ["Rank","Job Title","Company","City","Resume","Tier","Link","Status","Posted"]
COL_WIDTHS = [6, 48, 22, 16, 20, 6, 10, 12, 13]   # same order

def build_sheet_xml(rows, ss_idx, sheet_hyperlinks):
    """
    rows: list of job dicts
    ss_idx: dict str→int (shared string registry — shared across sheets)
    sheet_hyperlinks: list to append (cell_ref, url, tooltip) tuples
    Returns: sheet XML string
    """
    def ss(s):
        if s not in ss_idx:
            ss_idx[s] = len(ss_idx)
        return ss_idx[s]

    # cols element
    cols_xml = ""
    for ci, w in enumerate(COL_WIDTHS, 1):
        cols_xml += f'    <col min="{ci}" max="{ci}" width="{w}" customWidth="1"/>\n'

    row_xmls = []

    # Header row (row 1, height=28)
    cells = ""
    for ci, h in enumerate(HEADERS, 1):
        ref = cell_ref(1, ci)
        cells += f'      <c r="{ref}" s="{XF_HEADER}" t="s"><v>{ss(h)}</v></c>\n'
    row_xmls.append(f'    <row r="1" ht="28" customHeight="1">\n{cells}    </row>')

    # Data rows
    for ri, job in enumerate(rows, 2):
        tier = job["tier"]
        xf_d = xf_for_tier(tier)
        xf_l = xf_link_for_tier(tier)
        xf_dt= xf_date_for_tier(tier)
        cells = ""

        # col 1: Rank (number)
        cells += f'      <c r="{cell_ref(ri,1)}" s="{xf_d}"><v>{job["rank"]}</v></c>\n'
        # col 2: Job Title (string)
        cells += f'      <c r="{cell_ref(ri,2)}" s="{xf_d}" t="s"><v>{ss(job["title"])}</v></c>\n'
        # col 3: Company (string)
        cells += f'      <c r="{cell_ref(ri,3)}" s="{xf_d}" t="s"><v>{ss(job["company"])}</v></c>\n'
        # col 4: City (string)
        cells += f'      <c r="{cell_ref(ri,4)}" s="{xf_d}" t="s"><v>{ss(job["city"])}</v></c>\n'
        # col 5: Resume (string)
        cells += f'      <c r="{cell_ref(ri,5)}" s="{xf_d}" t="s"><v>{ss(job["resume"])}</v></c>\n'
        # col 6: Tier (number)
        cells += f'      <c r="{cell_ref(ri,6)}" s="{xf_d}"><v>{tier}</v></c>\n'
        # col 7: Link (string "Open" + hyperlink)
        href_ref = cell_ref(ri, 7)
        cells += f'      <c r="{href_ref}" s="{xf_l}" t="s"><v>{ss("Open")}</v></c>\n'
        if job["url"]:
            sheet_hyperlinks.append((href_ref, job["url"], job["title"]))
        # col 8: Status (string)
        cells += f'      <c r="{cell_ref(ri,8)}" s="{xf_d}" t="s"><v>{ss("")}</v></c>\n'
        # col 9: Posted (date as number or empty)
        d = job["posted"]
        if isinstance(d, Date):
            serial = date_serial(d)
            cells += f'      <c r="{cell_ref(ri,9)}" s="{xf_dt}"><v>{serial}</v></c>\n'
        else:
            cells += f'      <c r="{cell_ref(ri,9)}" s="{xf_dt}" t="s"><v>{ss("")}</v></c>\n'

        row_xmls.append(f'    <row r="{ri}" ht="20" customHeight="1">\n{cells}    </row>')

    sheet_data = "\n".join(row_xmls)

    # Hyperlinks element
    if sheet_hyperlinks:
        hl_xml = "  <hyperlinks>\n"
        for (ref, url, tip) in sheet_hyperlinks:
            # rId will be assigned later based on index
            hl_xml += f'    <hyperlink ref="{ref}" r:id="hl_{ref}" tooltip="{xe(tip)}"/>\n'
        hl_xml += "  </hyperlinks>"
    else:
        hl_xml = ""

    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
{cols_xml}  </cols>
  <sheetData>
{sheet_data}
  </sheetData>
{hl_xml}
</worksheet>"""

def sheet_rels_xml(hyperlinks):
    """hyperlinks: list of (cell_ref, url, tooltip)"""
    if not hyperlinks:
        return ""
    rels = ""
    for ref, url, tip in hyperlinks:
        rels += (f'  <Relationship Id="hl_{ref}" '
                 f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                 f'Target="{xe(url)}" TargetMode="External"/>\n')
    return f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels}</Relationships>"""

# ─────────────────────────────────────────────────────────────────────────────
# BUILD XLSX
# ─────────────────────────────────────────────────────────────────────────────
priority = [j for j in filtered if j["tier"] in (1, 2)]
all_jobs  = filtered

sheet_defs = [
    ("Priority Jobs",     priority),
    ("All Filtered Jobs", all_jobs),
]

ss_idx = {}  # shared string registry (mutable, passed into both sheets)
sheets_data = []
for name, rows in sheet_defs:
    hls = []
    xml = build_sheet_xml(rows, ss_idx, hls)
    sheets_data.append((name, xml, hls))

# rebuild shared strings list in index order
shared_list = [""] * len(ss_idx)
for s, i in ss_idx.items():
    shared_list[i] = s

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml",       content_types_xml(len(sheet_defs)))
    zf.writestr("_rels/.rels",               root_rels_xml())
    zf.writestr("xl/workbook.xml",           workbook_xml([n for n,_,_ in sheets_data]))
    zf.writestr("xl/_rels/workbook.xml.rels",workbook_rels_xml(len(sheet_defs)))
    zf.writestr("xl/styles.xml",             styles_xml())
    zf.writestr("xl/sharedStrings.xml",      shared_strings_xml(shared_list))
    for i, (name, xml, hls) in enumerate(sheets_data, 1):
        zf.writestr(f"xl/worksheets/sheet{i}.xml", xml)
        rels = sheet_rels_xml(hls)
        if rels:
            zf.writestr(f"xl/worksheets/_rels/sheet{i}.xml.rels", rels)

with open(OUTPUT, "wb") as f:
    f.write(buf.getvalue())

print(f"Saved → {OUTPUT}  ({os.path.getsize(OUTPUT):,} bytes)")
print(f"Priority Jobs sheet: {len(priority)} rows")
print(f"All Filtered Jobs sheet: {len(all_jobs)} rows")
