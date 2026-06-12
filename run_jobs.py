"""
LinkedIn jobs scraper → Excel → Google Drive
Apify actor: curious_coder/linkedin-jobs-scraper — 17 URLs, parallel max_workers=5
"""
import os, re, time
from datetime import date, datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests

import sys
sys.path.insert(0, "/home/user/job-tracker-routine")
from xlsx_builder import Workbook, _date_serial, _col_letter

# ── config ────────────────────────────────────────────────────────────────────
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

# ── filters ───────────────────────────────────────────────────────────────────
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

# ── tiering ───────────────────────────────────────────────────────────────────
TIER1_TITLE_RE = re.compile(
    r"fp.?a|financial planning|business finance|credit analyst|real estate.*anal|"
    r"valuation|deals analyst|corporate finance|transaction advisory",
    re.I
)
TIER1_EXCL_RE  = re.compile(r"software|tech|it |information technology", re.I)
TIER2_TITLE_RE = re.compile(
    r"financial analyst|strategy analyst|investment analyst|ib analyst|"
    r"senior analyst|associate analyst|strategy.*analyst|analyst.*strategy",
    re.I
)
TIER2_IND_RE   = re.compile(
    r"bank|nbfc|consult|financial service|insurance|asset management|private equity|"
    r"venture|capital market|brokerage|investment",
    re.I
)

def get_tier(job):
    title    = job.get("title") or ""
    company  = job.get("companyName") or job.get("company") or ""
    industry = job.get("industries") or job.get("companyIndustry") or job.get("industry") or ""
    if TIER1_TITLE_RE.search(title) and not TIER1_EXCL_RE.search(industry):
        return 1
    if TIER2_TITLE_RE.search(title) and TIER2_IND_RE.search(industry + " " + company):
        return 2
    return 3

def get_resume(job):
    title    = (job.get("title") or "").lower()
    industry = (job.get("industries") or job.get("companyIndustry") or job.get("industry") or "").lower()
    if re.search(r"real estate|valuation|development finance|property|asset management", title, re.I):
        return "RE Resume"
    if re.search(r"\bib\b|tas\b|deals|m&a|structured finance|capital markets", title, re.I):
        return "ECM Resume"
    if re.search(r"fp.?a|business finance|strategy|corporate finance|financial planning", title, re.I):
        return "Strategy Resume"
    if re.search(r"financial analyst", title, re.I) and re.search(r"bank|financial service|nbfc|capital", industry, re.I):
        return "Strategy/ECM"
    if re.search(r"credit analyst", title, re.I):
        return "RE/Strategy"
    return "Strategy Resume"

# ── Apify ─────────────────────────────────────────────────────────────────────
BASE = "https://api.apify.com/v2"

def run_actor(search_url):
    r = requests.post(
        f"{BASE}/acts/{ACTOR_ID}/runs?token={APIFY_TOKEN}&waitForFinish=300",
        json={"urls": [search_url], "count": COUNT, "scrapeCompany": False},
        timeout=360,
    )
    r.raise_for_status()
    run_data   = r.json()["data"]
    run_id     = run_data["id"]
    dataset_id = run_data["defaultDatasetId"]

    for _ in range(60):
        sr     = requests.get(f"{BASE}/actor-runs/{run_id}?token={APIFY_TOKEN}", timeout=30)
        status = sr.json()["data"]["status"]
        if status == "SUCCEEDED":
            break
        if status in ("FAILED", "ABORTED", "TIMED-OUT"):
            print(f"[APIFY] Run {run_id} terminal status: {status}", flush=True)
            return []
        time.sleep(5)

    items_r = requests.get(
        f"{BASE}/datasets/{dataset_id}/items?token={APIFY_TOKEN}&format=json&limit=200",
        timeout=60,
    )
    items_r.raise_for_status()
    return items_r.json()

# ── main ──────────────────────────────────────────────────────────────────────
def main():
    results = {}

    def fetch(url):
        jobs = run_actor(url)
        print(f"[APIFY] {len(jobs)} raw results from URL {url}", flush=True)
        return url, jobs

    with ThreadPoolExecutor(max_workers=5) as pool:
        futures = {pool.submit(fetch, url): url for url in SEARCH_URLS}
        for fut in as_completed(futures):
            url = futures[fut]
            try:
                _, jobs = fut.result()
            except Exception as e:
                print(f"[APIFY] ERROR {url}: {e}", flush=True)
                jobs = []
            results[url] = jobs

    # filter + dedup in original URL order for deterministic output
    all_jobs, seen = [], set()
    for url in SEARCH_URLS:
        for job in results.get(url, []):
            if should_filter(job):
                continue
            key = (
                (job.get("companyName") or job.get("company") or "").strip().lower(),
                (job.get("title") or "").strip().lower(),
            )
            if key in seen:
                continue
            seen.add(key)
            all_jobs.append(job)

    print(f"[FILTER] {len(all_jobs)} jobs remaining after filter+dedup", flush=True)

    # annotate
    for job in all_jobs:
        job["_tier"]   = get_tier(job)
        job["_resume"] = get_resume(job)
        raw_url = job.get("link") or job.get("jobUrl") or job.get("url") or job.get("applyUrl") or ""
        job["_url"]  = clean_url(raw_url)
        loc = job.get("location") or job.get("city") or ""
        m = re.match(r"([^,]+)", loc)
        job["_city"] = m.group(1).strip() if m else loc

    all_jobs.sort(key=lambda j: (j["_tier"], (j.get("companyName") or j.get("company") or "").lower()))
    for rank, job in enumerate(all_jobs, 1):
        job["_rank"] = rank

    priority = [j for j in all_jobs if j["_tier"] in (1, 2)]
    print(f"Priority (T1+T2): {len(priority)}  |  All: {len(all_jobs)}", flush=True)

    # ── xlsx ──────────────────────────────────────────────────────────────────
    COL_WIDTHS = [6, 48, 22, 16, 20, 6, 10, 12, 13]
    HEADERS    = ["Rank", "Job Title", "Company", "City", "Resume", "Tier", "Link", "Status", "Posted"]

    def fill_sheet(ws, jobs):
        ws.freeze_row(1)
        for ci, w in enumerate(COL_WIDTHS, 1):
            ws.set_col_width(ci, w)
        ws.set_row_height(1, 28)
        for ci, h in enumerate(HEADERS, 1):
            ws.write(1, ci, h, style=1)

        for ri, job in enumerate(jobs, 2):
            t      = job["_tier"]
            base_s = 2 if t == 1 else (3 if t == 2 else 0)
            hyp_s  = 5 if t == 1 else (6 if t == 2 else 4)
            date_s = 9 if t == 1 else (10 if t == 2 else 7)
            ws.set_row_height(ri, 20)

            ws.write(ri, 1, job["_rank"],                                                style=base_s)
            ws.write(ri, 2, job.get("title") or "",                                      style=base_s)
            ws.write(ri, 3, job.get("companyName") or job.get("company") or "",          style=base_s)
            ws.write(ri, 4, job["_city"],                                                style=base_s)
            ws.write(ri, 5, job["_resume"],                                              style=base_s)
            ws.write(ri, 6, t,                                                           style=base_s)

            url = job["_url"]
            if url:
                ws.write_hyperlink(ri, 7, "Open", url, style=hyp_s)
            else:
                ws.write(ri, 7, "", style=base_s)

            ws.write(ri, 8, "", style=base_s)

            raw_date = job.get("postedDate") or job.get("publishedAt") or job.get("postedAt") or ""
            if raw_date and len(raw_date) >= 10:
                try:
                    d = datetime.strptime(raw_date[:10], "%Y-%m-%d").date()
                    ws.write(ri, 9, d, style=date_s)
                except Exception:
                    ws.write(ri, 9, raw_date[:10], style=base_s)
            else:
                ws.write(ri, 9, "", style=base_s)

    wb  = Workbook()
    ws1 = wb.add_sheet("Priority Jobs")
    ws2 = wb.add_sheet("All Filtered Jobs")
    fill_sheet(ws1, priority)
    fill_sheet(ws2, all_jobs)

    data = wb.save(OUTPUT_PATH)
    print(f"[EXCEL] File saved, size={len(data):,} bytes", flush=True)

    return data

if __name__ == "__main__":
    main()
