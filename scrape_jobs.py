#!/usr/bin/env python3
"""LinkedIn Jobs → Filter → XLSX (pure Python, no openpyxl needed)"""

import os, re, sys, time, json, zipfile, io
from datetime import datetime, timedelta, date as date_cls
import requests

APIFY_TOKEN = os.environ['APIFY_API_TOKEN']
TODAY = datetime.now().date()
TODAY_STR = TODAY.strftime('%Y-%m-%d')
OUTPUT = f"/home/user/job-tracker-routine/LinkedIn_Jobs_{TODAY_STR}.xlsx"

# ─── 1. Build search URLs ─────────────────────────────────────────────────────
BASE = "https://www.linkedin.com/jobs/search/?f_TPR=r86400&"
KWS = [
    "FP%26A%20analyst", "strategy%20analyst", "business%20finance%20analyst",
    "credit%20analyst", "real%20estate%20analyst", "valuation%20analyst",
    "deals%20analyst", "transaction%20advisory%20analyst", "corporate%20finance%20analyst",
]
LOCS = [
    "Mumbai%2C%20Maharashtra%2C%20India",
    "Delhi%2C%20India",
    "Bengaluru%2C%20Karnataka%2C%20India",
]
OMIT = {
    ("business%20finance%20analyst", "Delhi%2C%20India"),
    ("deals%20analyst",              "Bengaluru%2C%20Karnataka%2C%20India"),
    ("transaction%20advisory%20analyst", "Bengaluru%2C%20Karnataka%2C%20India"),
    ("corporate%20finance%20analyst", "Bengaluru%2C%20Karnataka%2C%20India"),
    ("valuation%20analyst",          "Bengaluru%2C%20Karnataka%2C%20India"),
}
URLS = [f"{BASE}keywords={k}&location={l}" for k in KWS for l in LOCS if (k, l) not in OMIT]
print(f"Built {len(URLS)} search URLs")

# ─── 2. Start Apify run ───────────────────────────────────────────────────────
ACTOR = "curious_coder~linkedin-jobs-scraper"

resp = requests.post(
    f"https://api.apify.com/v2/acts/{ACTOR}/runs?token={APIFY_TOKEN}",
    json={"urls": URLS, "count": 50, "scrapeCompany": False},
    timeout=60,
)
print(f"Run start: {resp.status_code}")
if not resp.ok:
    print(resp.text[:500]); sys.exit(1)

rd = resp.json()["data"]
RUN_ID, DS_ID = rd["id"], rd["defaultDatasetId"]
print(f"Run={RUN_ID}  Dataset={DS_ID}")

# ─── 3. Poll until done ───────────────────────────────────────────────────────
print("Polling…")
status = "RUNNING"
for i in range(180):
    time.sleep(20)
    try:
        s = requests.get(
            f"https://api.apify.com/v2/acts/{ACTOR}/runs/{RUN_ID}?token={APIFY_TOKEN}",
            timeout=30,
        ).json()["data"]
        status = s["status"]
        n = s.get("stats", {}).get("outputItems", "?")
        print(f"  [{(i+1)*20}s] {status} | items={n}")
    except Exception as e:
        print(f"  [{(i+1)*20}s] poll error: {e}")
        continue
    if status not in ("RUNNING", "READY", "CREATED"):
        break

if status != "SUCCEEDED":
    print(f"FATAL: run ended {status}"); sys.exit(1)

# ─── 4. Fetch results ─────────────────────────────────────────────────────────
print("Fetching dataset…")
raw = requests.get(
    f"https://api.apify.com/v2/datasets/{DS_ID}/items?token={APIFY_TOKEN}&limit=10000",
    timeout=120,
).json()
print(f"Fetched {len(raw)} raw items")
if raw:
    print("Keys:", list(raw[0].keys())[:15])
    print("Sample:", json.dumps(raw[0], default=str)[:600])

# ─── 5. Helpers ───────────────────────────────────────────────────────────────
def gf(j, *keys):
    for k in keys:
        v = j.get(k)
        if v not in (None, "", [], {}):
            return str(v)
    return ""

def extract_city(loc):
    c = str(loc).split(",")[0].strip() if loc else ""
    lc = c.lower()
    if "mumbai" in lc: return "Mumbai"
    if "delhi" in lc:  return "Delhi"
    if "bengaluru" in lc or "bangalore" in lc: return "Bengaluru"
    return c

def parse_date(raw):
    if not raw: return None
    s = str(raw).strip()
    try: return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except: pass
    m = re.match(r"(\d+)\s+day", s, re.I)
    if m: return TODAY - timedelta(days=int(m[1]))
    if re.search(r"today|hour|just\s*now", s, re.I): return TODAY
    if re.search(r"yesterday", s, re.I):              return TODAY - timedelta(1)
    m = re.match(r"(\d+)\s+week", s, re.I)
    if m: return TODAY - timedelta(weeks=int(m[1]))
    return None

# ─── 6. Filters & Tiers ───────────────────────────────────────────────────────
IT = re.compile(
    r"\b(infosys|tcs|tata consultancy|wipro|hcl\s*tech|accenture|cognizant|capgemini|"
    r"tech mahindra|hexaware|mphasis|mindtree|ltimindtree|persistent|zensar|mastek|"
    r"niit|birlasoft|sonata software|happiest minds|dxc|ntt data|virtusa|coforge|kpit)\b",
    re.I,
)
FIN = re.compile(
    r"\b(bank|finance|financial|capital|securities|insurance|investment|wealth|"
    r"asset.?management|mutual.?fund|nbfc|hdfc|icici|axis|kotak|bajaj|mahindra|"
    r"shriram|muthoot|iifl|aditya.?birla|tata.?capital|l&t.?finance|"
    r"jp.?morgan|goldman|morgan.?stanley|citi|barclays|hsbc|deutsche|ubs|bnp|"
    r"kpmg|deloitte|pwc|ernst|grant.?thornton|bdo|mckinsey|bcg|bain|"
    r"crisil|icra|angel|motilal|edelweiss|sbicap|ambit|nomura|dsp|nippon|franklin|mirae)\b",
    re.I,
)
T1_RE = re.compile(
    r"\b(fp[&\s]?a|financial.?planning|business.?finance|credit.?analyst|credit.?risk|"
    r"real.?estate|valuation|deals.?analyst|corporate.?finance|transaction.?advisory|tas.?analyst)\b",
    re.I,
)
T2_RE = re.compile(
    r"\b(financial.?analyst|finance.?analyst|strategy.?analyst|investment.?analyst|"
    r"equity.?analyst|research.?analyst|senior.?analyst|ib.?analyst|"
    r"investment.?banking.?analyst|associate.?analyst)\b",
    re.I,
)
DROP_T = re.compile(
    r"\b(administrative|executive.?assistant|ea\s+to|front.?office|guest.?relations|"
    r"compliance.?officer|sharepoint|erp|r2r|record.?to.?report|vat.?analyst|"
    r"production.?controller|chartered.?accountant|sales.?executive|"
    r"office.?administration|process.?associate)\b",
    re.I,
)
DROP_C = re.compile(r"^(mygwork|scoutit|alignerr|datamark|aditi\s+consulting)$", re.I)
DROP_S = re.compile(
    r"\b(director|vice.?president|\bvp\b|svp|evp|chief|cxo|cto|cfo|ceo|coo|internship|intern)\b",
    re.I,
)

RE_R  = re.compile(r"\b(real.?estate|valuation|development.?finance|property|asset.?management)\b", re.I)
ECM_R = re.compile(r"\b(investment.?banking|\bib\b|\btas\b|deals|m[&\s]?a|structured.?finance|capital.?markets)\b", re.I)
ST_R  = re.compile(r"\b(fp[&\s]?a|business.?finance|strategy|corporate.?finance|financial.?planning)\b", re.I)

def get_tier(t, c):
    if T1_RE.search(t) and not IT.search(c): return 1
    if T2_RE.search(t) and FIN.search(c):   return 2
    return 3

def get_resume(t, c):
    if re.search(r"\bcredit\b", t, re.I): return "RE/Strategy"
    if RE_R.search(t):  return "RE Resume"
    if ECM_R.search(t): return "ECM Resume"
    if ST_R.search(t):  return "Strategy Resume"
    if re.search(r"financial.?analyst|finance.?analyst", t, re.I) and FIN.search(c):
        return "Strategy/ECM"
    return "Strategy Resume"

# ─── 7. Filter & deduplicate ──────────────────────────────────────────────────
jobs, seen = [], set()
for j in raw:
    t = gf(j, "title", "jobTitle", "position")
    c = gf(j, "company", "companyName", "employer")
    if not t or not c: continue
    if DROP_T.search(t) or DROP_C.search(c.strip()) or DROP_S.search(t): continue
    if re.search(r"\b(director|vp|vice.?president|c.?suite|internship)\b",
                 gf(j, "seniorityLevel", "seniority"), re.I): continue
    if re.search(r"\b(internship|contract)\b",
                 gf(j, "employmentType", "type"), re.I): continue
    key = (c.lower(), t.lower())
    if key in seen: continue
    seen.add(key)
    jobs.append({
        "title":  t,
        "company": c,
        "city":   extract_city(gf(j, "location", "jobLocation")),
        "url":    gf(j, "link", "url", "jobUrl", "applyUrl"),
        "posted": parse_date(gf(j, "postedAt", "publishedAt", "listingDate", "datePosted", "date")),
        "tier":   get_tier(t, c),
        "resume": get_resume(t, c),
    })

tc = {1: 0, 2: 0, 3: 0}
for j in jobs: tc[j["tier"]] += 1
print(f"\nFiltered: {len(jobs)} jobs  T1={tc[1]}  T2={tc[2]}  T3={tc[3]}")

pri  = sorted([j for j in jobs if j["tier"] in (1, 2)], key=lambda j: (j["tier"], j["company"]))
all_ = sorted(jobs, key=lambda j: (j["tier"], j["company"]))

# ─── 8. XLSX helpers ──────────────────────────────────────────────────────────
def xe(s):
    """XML-escape a string."""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def cr(row, col):
    return f"{col_letter(col)}{row}"

_EPOCH = date_cls(1899, 12, 31)
def excel_serial(d):
    delta = (d - _EPOCH).days
    if d >= date_cls(1900, 3, 1):
        delta += 1
    return delta

# Style map: (cell_type, tier) → xf index
# xf 0: data no-fill  1: header  2: data T1  3: data T2
# xf 4: link no-fill  5: link T1  6: link T2
# xf 7: date no-fill  8: date T1  9: date T2
_S = {
    ("h",  0): 1, ("h",  1): 1, ("h",  2): 1, ("h",  3): 1,
    ("d",  0): 0, ("d",  1): 2, ("d",  2): 3, ("d",  3): 0,
    ("lk", 0): 4, ("lk", 1): 5, ("lk", 2): 6, ("lk", 3): 4,
    ("dt", 0): 7, ("dt", 1): 8, ("dt", 2): 9, ("dt", 3): 7,
}

class SS:
    """Shared strings table."""
    def __init__(self):
        self.lst = []; self.idx = {}
    def add(self, s):
        s = "" if s is None else str(s)
        if s not in self.idx:
            self.idx[s] = len(self.lst); self.lst.append(s)
        return self.idx[s]
    def xml(self):
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            f' count="{len(self.lst)}" uniqueCount="{len(self.lst)}">',
        ]
        for s in self.lst:
            parts.append(f'<si><t xml:space="preserve">{xe(s)}</t></si>')
        parts.append("</sst>")
        return "".join(parts)

HDRS = ["Rank", "Job Title", "Company", "City", "Resume", "Tier", "Link", "Status", "Posted"]
WIDTHS = [6, 48, 22, 16, 20, 6, 10, 12, 13]

def build_sheet(jobs, ss):
    rows_xml = []
    hls = []  # (cell_ref, rId, url, tooltip)

    # Header row
    hcells = "".join(
        f'<c r="{cr(1, ci)}" t="s" s="1"><v>{ss.add(h)}</v></c>'
        for ci, h in enumerate(HDRS, 1)
    )
    rows_xml.append(f'<row r="1" ht="28" customHeight="1">{hcells}</row>')

    for ri, j in enumerate(jobs, 2):
        tier = j["tier"]
        sd   = _S[("d",  tier)]
        sl   = _S[("lk", tier)]
        sdt  = _S[("dt", tier)]
        cells = []

        # Rank
        cells.append(f'<c r="{cr(ri,1)}" s="{sd}"><v>{ri-1}</v></c>')
        # Title
        cells.append(f'<c r="{cr(ri,2)}" t="s" s="{sd}"><v>{ss.add(j["title"])}</v></c>')
        # Company
        cells.append(f'<c r="{cr(ri,3)}" t="s" s="{sd}"><v>{ss.add(j["company"])}</v></c>')
        # City
        cells.append(f'<c r="{cr(ri,4)}" t="s" s="{sd}"><v>{ss.add(j["city"])}</v></c>')
        # Resume
        cells.append(f'<c r="{cr(ri,5)}" t="s" s="{sd}"><v>{ss.add(j["resume"])}</v></c>')
        # Tier
        cells.append(f'<c r="{cr(ri,6)}" s="{sd}"><v>{tier}</v></c>')
        # Link
        lref = cr(ri, 7)
        cells.append(f'<c r="{lref}" t="s" s="{sl}"><v>{ss.add("Open")}</v></c>')
        if j["url"]:
            rid = f"rId{len(hls)+1}"
            hls.append((lref, rid, j["url"], j["title"][:255]))
        # Status (empty, still styled)
        cells.append(f'<c r="{cr(ri,8)}" t="s" s="{sd}"><v>{ss.add("")}</v></c>')
        # Posted
        if j["posted"]:
            cells.append(f'<c r="{cr(ri,9)}" s="{sdt}"><v>{excel_serial(j["posted"])}</v></c>')
        else:
            cells.append(f'<c r="{cr(ri,9)}" s="{sd}"/>')

        rows_xml.append(
            f'<row r="{ri}" ht="20" customHeight="1">' + "".join(cells) + "</row>"
        )

    cols_xml = "<cols>" + "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
        for i, w in enumerate(WIDTHS, 1)
    ) + "</cols>"

    hl_xml = ""
    if hls:
        hl_xml = "<hyperlinks>" + "".join(
            f'<hyperlink ref="{ref}" r:id="{rid}" tooltip="{xe(tip)}"/>'
            for ref, rid, _, tip in hls
        ) + "</hyperlinks>"

    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheetViews>"
        '<sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="topLeft"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        "</sheetView></sheetViews>"
        '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>'
        + cols_xml
        + "<sheetData>" + "".join(rows_xml) + "</sheetData>"
        + hl_xml
        + "</worksheet>"
    )

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            f'<Relationship Id="{rid}"'
            f' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"'
            f' Target="{xe(url)}" TargetMode="External"/>'
            for _, rid, url, _ in hls
        )
        + "</Relationships>"
    )
    return sheet, rels

# ─── 9. Static XML blobs ──────────────────────────────────────────────────────
CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    "</Types>"
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    "</Relationships>"
)
WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    "<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"16384\" windowHeight=\"8192\"/></bookViews>"
    "<sheets>"
    '<sheet name="Priority Jobs" sheetId="1" r:id="rId1"/>'
    '<sheet name="All Filtered Jobs" sheetId="2" r:id="rId2"/>'
    "</sheets></workbook>"
)
WB_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)
# Styles:
# Font 0=Arial10 black  1=Arial10 bold white  2=Arial10 blue underline
# Fill 0=none  1=gray125  2=solid#1F4E79  3=solid#C8E6C9  4=solid#FFF9C4
# xf  0=data   1=header   2=data-T1  3=data-T2  4=link  5=link-T1  6=link-T2
#     7=date   8=date-T1  9=date-T2
STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="YYYY-MM-DD"/></numFmts>'
    '<fonts count="3">'
    '<font><sz val="10"/><name val="Arial"/></font>'
    '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>'
    '<font><u/><sz val="10"/><color rgb="FF0563C1"/><name val="Arial"/></font>'
    '</fonts>'
    '<fills count="5">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC8E6C9"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF9C4"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="10">'
    # 0: data, no fill
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    # 1: header
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">'
    '<alignment horizontal="center" vertical="center"/></xf>'
    # 2: data T1
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    # 3: data T2
    '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    # 4: link no fill
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    # 5: link T1
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    # 6: link T2
    '<xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    # 7: date no fill
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>'
    # 8: date T1
    '<xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>'
    # 9: date T2
    '<xf numFmtId="164" fontId="0" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>'
)

# ─── 10. Build & write XLSX ───────────────────────────────────────────────────
ss = SS()
s1_xml, s1_rels = build_sheet(pri,  ss)
s2_xml, s2_rels = build_sheet(all_, ss)

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml",               CONTENT_TYPES)
    zf.writestr("_rels/.rels",                        ROOT_RELS)
    zf.writestr("xl/workbook.xml",                    WORKBOOK)
    zf.writestr("xl/_rels/workbook.xml.rels",         WB_RELS)
    zf.writestr("xl/styles.xml",                      STYLES)
    zf.writestr("xl/sharedStrings.xml",               ss.xml())
    zf.writestr("xl/worksheets/sheet1.xml",           s1_xml)
    zf.writestr("xl/worksheets/sheet2.xml",           s2_xml)
    zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", s1_rels)
    zf.writestr("xl/worksheets/_rels/sheet2.xml.rels", s2_rels)

with open(OUTPUT, "wb") as f:
    f.write(buf.getvalue())

size = os.path.getsize(OUTPUT)
print(f"\nSaved: {OUTPUT}  ({size:,} bytes)")
print(f"Priority Jobs: {len(pri)} rows")
print(f"All Filtered Jobs: {len(all_)} rows")
print("DONE")
