"""
LinkedIn jobs scraper → Excel → Google Drive
Uses Apify actor curious_coder/linkedin-jobs-scraper
"""
import os, json, time, re, base64, zipfile
from datetime import date, datetime
import requests

sys_import = __import__("sys")
sys_import.path.insert(0, "/home/user/job-tracker-routine")
from xlsx_builder import Workbook, _date_serial, _col_letter

# ── config ──────────────────────────────────────────────────────────────────
APIFY_TOKEN = os.environ["APIFY_API_TOKEN"]
ACTOR_ID    = "curious_coder~linkedin-jobs-scraper"
COUNT       = 75
TODAY       = date.today()
FILENAME    = f"LinkedIn_Jobs_{TODAY.strftime('%Y-%m-%d')}.xlsx"
OUTPUT_PATH = FILENAME

SEARCH_URLS = [
    "https://www.linkedin.com/jobs/search/?keywords=FP%26A%20analyst&location=Mumbai%2C%20Maharashtra%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=strategy%20analyst&location=Mumbai%2C%20Maharashtra%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=business%20finance%20analyst&location=Mumbai%2C%20Maharashtra%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=credit%20analyst&location=Mumbai%2C%20Maharashtra%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=real%20estate%20analyst&location=Mumbai%2C%20Maharashtra%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=deals%20analyst&location=Mumbai%2C%20Maharashtra%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=corporate%20finance%20analyst&location=Mumbai%2C%20Maharashtra%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=FP%26A%20analyst&location=Delhi%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=strategy%20analyst&location=Delhi%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=credit%20analyst&location=Delhi%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=real%20estate%20analyst&location=Delhi%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=corporate%20finance%20analyst&location=Delhi%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=FP%26A%20analyst&location=Bengaluru%2C%20Karnataka%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=strategy%20analyst&location=Bengaluru%2C%20Karnataka%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=business%20finance%20analyst&location=Bengaluru%2C%20Karnataka%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=real%20estate%20analyst&location=Bengaluru%2C%20Karnataka%2C%20India&f_TPR=r86400",
    "https://www.linkedin.com/jobs/search/?keywords=credit%20analyst&location=Bengaluru%2C%20Karnataka%2C%20India&f_TPR=r86400",
]

# ── filters ──────────────────────────────────────────────────────────────────
BAD_SENIORITY = {"director", "vp", "vice president", "c-suite", "ceo", "cfo", "coo",
                 "chief", "intern", "internship"}
BAD_TYPE      = {"internship", "contract"}
BAD_TITLE_RE  = re.compile(
    r"administrative|executive assistant|ea to|front office|guest relations|"
    r"compliance officer|sharepoint|erp|r2r|record to report|vat analyst|"
    r"production controller|chartered accountant|sales executive|"
    r"office administration|process associate",
    re.I
)
BAD_COMPANIES = {"mygwork", "scoutit", "alignerr", "datamark", "aditi consulting"}

def clean_url(url):
    """Strip query params from LinkedIn URL."""
    return url.split("?")[0]

def should_filter(job):
    title     = (job.get("title") or "").lower()
    company   = (job.get("companyName") or job.get("company") or "").lower().strip()
    seniority = (job.get("seniorityLevel") or job.get("seniority") or "").lower()
    job_type  = (job.get("employmentType") or job.get("jobType") or "").lower()

    if any(s in seniority for s in BAD_SENIORITY):
        return True
    if any(s in job_type for s in BAD_TYPE):
        return True
    if BAD_TITLE_RE.search(title):
        return True
    if company in BAD_COMPANIES:
        return True
    return False

# ── tiering ─────────────────────────────────────────────────────────────────
TIER1_TITLE_RE = re.compile(
    r"fp.?a|financial planning|business finance|credit analyst|real estate.*anal|"
    r"valuation|deals analyst|corporate finance|transaction advisory",
    re.I
)
TIER1_EXCL_INDUSTRY_RE = re.compile(r"software|tech|it |information technology", re.I)

TIER2_TITLE_RE = re.compile(
    r"financial analyst|strategy analyst|investment analyst|ib analyst|"
    r"senior analyst|associate analyst|strategy.*analyst|analyst.*strategy",
    re.I
)
TIER2_INDUSTRY_RE = re.compile(
    r"bank|nbfc|consult|financial service|insurance|asset management|private equity|"
    r"venture|capital market|brokerage|investment",
    re.I
)

def tier(job):
    title   = job.get("title") or ""
    company = job.get("companyName") or job.get("company") or ""
    industry = job.get("industries") or job.get("companyIndustry") or job.get("industry") or ""

    t1 = TIER1_TITLE_RE.search(title)
    if t1:
        # exclude IT firms if it's not explicitly finance
        if TIER1_EXCL_INDUSTRY_RE.search(industry):
            pass  # fall through to tier2/3
        else:
            return 1

    if TIER2_TITLE_RE.search(title) and TIER2_INDUSTRY_RE.search(industry + " " + company):
        return 2

    return 3

# ── resume mapping ───────────────────────────────────────────────────────────
def resume(job):
    title   = (job.get("title") or "").lower()
    industry = (job.get("industries") or job.get("companyIndustry") or job.get("industry") or "").lower()

    if re.search(r"real estate|valuation|development finance|property|asset management", title, re.I):
        return "RE Resume"
    if re.search(r"\bib\b|tas\b|deals|m&a|structured finance|capital markets", title, re.I):
        return "ECM Resume"
    if re.search(r"fp.?a|business finance|strategy|corporate finance|financial planning", title, re.I):
        return "Strategy Resume"
    if re.search(r"financial analyst", title, re.I):
        if re.search(r"bank|financial service|nbfc|capital", industry, re.I):
            return "Strategy/ECM"
    if re.search(r"credit analyst", title, re.I):
        return "RE/Strategy"
    return "Strategy Resume"

# ── Apify helpers ────────────────────────────────────────────────────────────
BASE = "https://api.apify.com/v2"

def run_actor(search_url):
    payload = {
        "urls": [search_url],
        "count": COUNT,
        "scrapeCompany": False,
    }
    r = requests.post(
        f"{BASE}/acts/{ACTOR_ID}/runs?token={APIFY_TOKEN}&waitForFinish=300",
        json=payload,
        timeout=360,
    )
    r.raise_for_status()
    data = r.json()
    run_id = data["data"]["id"]
    dataset_id = data["data"]["defaultDatasetId"]
    # wait for SUCCEEDED
    for _ in range(60):
        status_r = requests.get(f"{BASE}/actor-runs/{run_id}?token={APIFY_TOKEN}", timeout=30)
        status = status_r.json()["data"]["status"]
        if status == "SUCCEEDED":
            break
        if status in ("FAILED","ABORTED","TIMED-OUT"):
            print(f"  Run {run_id} ended with status {status}")
            return []
        time.sleep(5)

    items_r = requests.get(
        f"{BASE}/datasets/{dataset_id}/items?token={APIFY_TOKEN}&format=json&limit=200",
        timeout=60,
    )
    items_r.raise_for_status()
    return items_r.json()

# ── main ─────────────────────────────────────────────────────────────────────
def main():
    all_jobs = []
    seen = set()   # (company_lower, title_lower) dedup key

    for i, url in enumerate(SEARCH_URLS, 1):
        kw = re.search(r"keywords=([^&]+)", url)
        loc = re.search(r"location=([^&]+)", url)
        label = f"{kw.group(1) if kw else '?'} / {loc.group(1) if loc else '?'}"
        print(f"[{i:02d}/{len(SEARCH_URLS)}] {label} ...", flush=True)
        try:
            jobs = run_actor(url)
            print(f"  → {len(jobs)} raw results", flush=True)
        except Exception as e:
            print(f"  ERROR: {e}", flush=True)
            jobs = []

        for job in jobs:
            if should_filter(job):
                continue
            title_l   = (job.get("title") or "").strip().lower()
            company_l = (job.get("companyName") or job.get("company") or "").strip().lower()
            key = (company_l, title_l)
            if key in seen:
                continue
            seen.add(key)
            all_jobs.append(job)

    print(f"\nTotal after filter+dedup: {len(all_jobs)}", flush=True)

    # Assign tier, resume, cleaned url, city
    for job in all_jobs:
        job["_tier"]   = tier(job)
        job["_resume"] = resume(job)
        raw_url = job.get("link") or job.get("jobUrl") or job.get("url") or job.get("applyUrl") or ""
        job["_url"] = clean_url(raw_url)
        # city extraction
        loc = job.get("location") or job.get("city") or ""
        city_match = re.match(r"([^,]+)", loc)
        job["_city"] = city_match.group(1).strip() if city_match else loc

    # Sort: tier → company
    all_jobs.sort(key=lambda j: (j["_tier"], (j.get("companyName") or j.get("company") or "").lower()))

    # Rank
    for rank, job in enumerate(all_jobs, 1):
        job["_rank"] = rank

    priority = [j for j in all_jobs if j["_tier"] in (1, 2)]

    print(f"Priority (T1+T2): {len(priority)}  |  All: {len(all_jobs)}", flush=True)

    # ── build xlsx ────────────────────────────────────────────────────────────
    COL_WIDTHS = [6, 48, 22, 16, 20, 6, 10, 12, 13]
    HEADERS    = ["Rank","Job Title","Company","City","Resume","Tier","Link","Status","Posted"]

    def fill_sheet(ws, jobs, priority_only=False):
        # freeze row 1
        ws.freeze_row(1)
        # col widths
        for ci, w in enumerate(COL_WIDTHS, 1):
            ws.set_col_width(ci, w)
        # header row
        ws.set_row_height(1, 28)
        for ci, h in enumerate(HEADERS, 1):
            # col 9 (Posted) is a date in header → still text label
            ws.write(1, ci, h, style=1)

        for ri_offset, job in enumerate(jobs, 2):
            t = job["_tier"]
            # base fill style index: tier1=2, tier2=3, else=0
            base_s = 2 if t == 1 else (3 if t == 2 else 0)
            hyp_s  = 5 if t == 1 else (6 if t == 2 else 4)
            date_s = 9 if t == 1 else (10 if t == 2 else 7)

            ws.set_row_height(ri_offset, 20)

            ws.write(ri_offset, 1, job["_rank"],            style=base_s)
            ws.write(ri_offset, 2, job.get("title") or "",  style=base_s)
            ws.write(ri_offset, 3, job.get("companyName") or job.get("company") or "", style=base_s)
            ws.write(ri_offset, 4, job["_city"],             style=base_s)
            ws.write(ri_offset, 5, job["_resume"],           style=base_s)
            ws.write(ri_offset, 6, t,                        style=base_s)

            # hyperlink col 7
            url = job["_url"]
            if url:
                ws.write_hyperlink(ri_offset, 7, "Open", url, style=hyp_s)
            else:
                ws.write(ri_offset, 7, "", style=base_s)

            ws.write(ri_offset, 8, "", style=base_s)   # Status blank

            # Posted date
            raw_date = job.get("postedDate") or job.get("publishedAt") or job.get("postedAt") or ""
            if raw_date and len(raw_date) >= 10:
                try:
                    d = datetime.strptime(raw_date[:10], "%Y-%m-%d").date()
                    ws.write(ri_offset, 9, d, style=date_s)
                except:
                    ws.write(ri_offset, 9, raw_date[:10], style=base_s)
            else:
                ws.write(ri_offset, 9, "", style=base_s)

    wb = Workbook()
    ws1 = wb.add_sheet("Priority Jobs")
    ws2 = wb.add_sheet("All Filtered Jobs")

    fill_sheet(ws1, priority, priority_only=True)
    fill_sheet(ws2, all_jobs)

    data = wb.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}  ({len(data):,} bytes)", flush=True)

    # verify integrity
    with zipfile.ZipFile(OUTPUT_PATH) as z:
        bad = z.testzip()
        if bad:
            os.remove(OUTPUT_PATH)
            raise RuntimeError(f"ZIP corrupt: {bad}")
    print("ZIP integrity OK", flush=True)

    return data, all_jobs

if __name__ == "__main__":
    data, jobs = main()
    print(json.dumps({"status":"ok","rows":len(jobs),"file":OUTPUT_PATH}))
