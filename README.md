# Citadele App Store Reviews Fetcher

Fetches Apple App Store reviews for **Citadele Bank** (app ID `495139240`) from
the Estonia, Lithuania and Latvia storefronts and saves them as JSON and Excel.

## What you get

Per run (in `./output/`):

- `citadele_reviews_<timestamp>.json` — full machine-readable archive
- `citadele_reviews_<timestamp>.xlsx` — Excel with two sheets:
  - **Reviews** — every review, filterable, with country / rating / title / body / author / app version / date
  - **Summary** — review counts and average rating per country, plus a rating distribution (5★…1★)
- `citadele_reviews_latest.json` / `.xlsx` — overwritten on every run, useful for automation

## How much data does this actually return?

Apple's public RSS feed for customer reviews returns roughly **the most recent
~500 reviews per country** (10 pages × 50 entries, sorted most-recent-first).
For Citadele's three Baltic storefronts that's typically 1,000–1,500 reviews
across EE / LT / LV — enough for sentiment analysis, version-over-version
comparison, etc.

If you need the **complete history** (every review ever, including ones older
than the RSS window) you need the **App Store Connect API**. That requires:

- Being an admin/account holder on Citadele's App Store Connect account
- Generating an API key (Users and Access → Integrations → App Store Connect API)
- Signing JWTs to call `GET /v1/apps/{id}/customerReviews`

The RSS approach below needs none of that.

## Run locally (Mac)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python fetch_reviews.py
```

That's it. Open the file in `output/`.

## Run on GitHub Actions (free, scheduled)

If you want this to run automatically (say, every Monday morning) and commit
the latest snapshot back to the repo:

1. Push this folder to a GitHub repo
2. Add `.github/workflows/fetch-reviews.yml` (template below)
3. Done. Each run will commit the new `output/citadele_reviews_latest.*` files

```yaml
name: Fetch Citadele App Store Reviews

on:
  schedule:
    - cron: "0 6 * * 1"   # every Monday 06:00 UTC
  workflow_dispatch:        # also lets you run it manually from the Actions tab

permissions:
  contents: write           # needed to commit the output files

jobs:
  fetch:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - run: pip install -r requirements.txt
      - run: python fetch_reviews.py

      - name: Commit snapshot
        run: |
          git config user.name  "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add output/
          git diff --cached --quiet || git commit -m "chore: refresh reviews $(date -u +%F)"
          git push
```

That's the "Claude + GitHub" pattern you were asking about: Claude wrote the
code, GitHub Actions runs it on a schedule, and Git keeps a history of every
snapshot so you can diff week-over-week.

## Connect this to Claude

A few ways to plug the output into Claude:

1. **Upload `citadele_reviews_latest.xlsx`** to a Claude chat and ask things
   like "summarise the top complaint themes in the last 30 days for LV" or
   "show me 1- and 2-star reviews mentioning Face ID".
2. **Put the repo into a Claude Project** with the workflow above — every fresh
   snapshot is then a project file Claude can read.
3. **Use Claude Code** locally, point it at the JSON, and have it generate ad
   hoc analyses (trend charts, theme extraction, etc.).

## Notes / gotchas

- Apple throttles aggressive callers — the script sleeps 1 s between page
  requests and bails politely on a 403.
- Country code is the App Store storefront, not the user's nationality.
- The feed sorts most-recent-first; for trend analysis sort by the `updated`
  field after loading.
- Review text comes back in whatever language the reviewer wrote in (Latvian,
  Lithuanian, Estonian, Russian, English…). If you want everything in one
  language, run a translation pass downstream — Claude is good at this.
