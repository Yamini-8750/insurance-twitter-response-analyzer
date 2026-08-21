# Insurance Twitter Response-Time Analyzer

Automatically measures how fast insurance companies respond to customer complaints on Twitter/X — turning a slow, manual, small-sample task into a repeatable pipeline that scores responsiveness at scale.

> **Impact:** Replaced a manual process (visiting each tweet, eyeballing timestamps, recording gaps by hand) that was slow, error-prone, and capped at tiny sample sizes. The pipeline now reads full complaint threads, identifies insurer replies, computes response times, and exports a scoring-ready CSV — **~150 complaint threads processed in one run for under $2**, at effectively zero cost for the first ~650 threads.

---

## What it does

- Reads the full reply thread for each customer-complaint tweet
- Detects which replies came from the **insurer** (vs the customer or others)
- Calculates **first response time**, **average thread response time**, and **per-reply gaps**
- Flags threads where the insurer **never replied**
- Exports a clean CSV ready to paste straight into the scoring sheet

## Why it exists

Insurance ratings depend on responsiveness, but there was no reliable way to measure it. Off-the-shelf change-detection tooling couldn't handle Twitter threads, and manual tracking didn't scale past a handful of tweets. This tool makes responsiveness a measurable, monthly metric.

## How it works

```
Twitter Advanced Search  ─►  Export complaint URLs (CSV)
                                      │
                                      ▼
        twitterapi.io  ◄──  Python script (fetch_threads.py)
        (reads full thread, exact timestamps)
                                      │
                                      ▼
        Identify insurer replies  ─►  Compute response gaps
                                      │
                                      ▼
        twitter_summary.csv  +  twitter_timestamps.csv  +  threads.json
```

1. **Collect** — pull complaint tweet URLs for an insurer/month via Twitter Advanced Search (root tweets only) and export with a browser extension.
2. **Fetch** — the Python script calls `twitterapi.io` (official API access, not scraping) to retrieve each full conversation thread with exact timestamps.
3. **Analyze** — it matches insurer replies, computes first/average/per-reply response times, and flags no-reply threads.
4. **Export** — three outputs: a per-tweet summary CSV for scoring, a per-turn timestamp CSV for audit, and a JSON for monthly trend charts.

## Tech stack

`Python` · `twitterapi.io` (Twitter data API) · `requests` · CSV/JSON export · Twitter Advanced Search + browser-export extension for collection

## Output

| File | Contents | Use |
|---|---|---|
| `twitter_summary.csv` | One row per complaint: replied Y/N, first & avg response time, per-reply gaps, turn counts | Scoring sheet |
| `twitter_timestamps.csv` | Every individual turn in every thread | Detailed audit |
| `twitter_threads.json` | Structured threads | Monthly trend charts |

## Design notes (V1 → V2)

The first version used a no-code cloud scraper (Apify). It worked but hit hard limits: a 5-URL cap per run (≈60 manual runs for 300 tweets), paid tiers for bulk access, manual URL collection, and breakage whenever Twitter changed its UI. **V2 moved to a direct API + local Python script** — no URL limit, near-zero cost, full multi-turn thread capture, and far more resilient. This trade-off analysis (cost, scale, accuracy, maintainability) is the core product decision behind the project.

## Setup

```bash
pip3 install requests
# In fetch_threads.py, set:
#   API_KEY        = "your_twitterapi_io_key"   (via .env — do NOT commit)
#   INSURER_HANDLE = "e.g. ExampleSupport"
python3 fetch_threads.py
```

> ⚠️ Store your API key in a `.env` file (see `.env.example`) and keep it out of version control. Rate limit: ~1 request / 5s on the free tier; the script waits 7s between calls.

## Repository contents

```
fetch_threads.py     # main pipeline
.env.example         # names of required secrets (no real values)
sample_output.csv    # anonymized example output
README.md
```
