# Twitch streamer scanner

A self-hosted service that finds and qualifies small Twitch streamers in a target market,
scores them, and publishes the results as a spreadsheet, an interactive board and a change
log — without anyone touching it.

Built to replace a manual sourcing process that involved reading third-party analytics
sites by hand. It has tracked **~9,900 channels** and produced a working shortlist of
**437 candidates**, updated hourly.

> This is the public code-only mirror. The private repository holds the collected data.

---

## What it does

| | |
|---|---|
| **Collect** | Sweeps every live English-language stream inside a configurable viewer range via the Twitch Helix API, hourly. |
| **Enrich** | Pulls profile, follower count, Twitch status and account age for channels seen in more than one sweep. |
| **Qualify** | Infers location from profile text, timezone mentions and go-live hours; excludes 24/7 reruns, watch parties and badge farms; detects giveaway language, discount codes and brand names. |
| **Score** | Ranks each channel 0–100 across engagement, consistency, brand experience, location certainty and audience size, with the breakdown exposed per row. |
| **Report** | Generates a multi-sheet Excel workbook with live formulas and charts, a self-contained HTML board with search, filters and per-candidate notes, and CSV exports. |
| **Track** | Diffs each run against the previous one: who is new, who dropped off, whose score moved, who started advertising a giveaway or a sponsor code. |
| **Notify** | Opens a GitHub issue when a run finds something worth acting on; optional webhook for Slack, Discord, Telegram or ntfy. |

## How it runs

```
GitHub Actions (hourly cron)
  └── docker run scanner ci
        ├── Twitch Helix API   (app access token, client credentials)
        ├── SQLite             (sightings + channels + run snapshots)
        └── outputs            (xlsx, html, csv, markdown notes)
  └── commit results back to the repo  →  the git history is the research log
```

Stack: Python 3.12, `requests`, `openpyxl`, SQLite, Docker, GitHub Actions. No framework,
no third-party data source, no scraping.

## Design decisions worth explaining

**Sampling over time, not snapshots.** The Helix API only exposes streams that are live at
the moment of the call. One call tells you who happened to be streaming that minute. Sampling
hourly and storing every sighting turns that into consistency data: who streams regularly,
when they go live, how their audience holds. The whole qualification model depends on it.

**App access token, not user OAuth.** Follower totals are readable with a client-credentials
token, so the scheduled job needs no browser login and no refresh token in CI secrets. The
script falls back to a user token automatically when one is present locally.

**Excluding the easy wins.** 24/7 rerun channels and drop farms show very high viewer-to-follower
ratios and would dominate any naive engagement ranking. They are detected by name patterns,
title patterns and average stream length, and excluded with the reason recorded, so the
filter can be audited and reversed.

**Every exclusion keeps its reason.** The `All channels` sheet lists every channel scanned
along with why it did not make the shortlist. Filters that can't be challenged don't get
trusted.

**Location is inferred, and labelled as such.** Twitch publishes no country. Profile text and
timezone mentions are treated as high and medium confidence; streaming hours alone are marked
as a guess. The distinction is carried through to every output rather than hidden.

## Running it

```bash
pip install requests openpyxl
export TWITCH_CLIENT_ID=... TWITCH_CLIENT_SECRET=...

python streamer_finder_v3.py collect    # one sweep
python streamer_finder_v3.py loop 60    # sweep hourly
python streamer_finder_v3.py all        # regenerate every output
python streamer_finder_v3.py notes      # diff against the previous run
python streamer_finder_v3.py weekly     # roll 7 days into a summary
```

Or containerised:

```bash
docker build -t scanner .
docker run --rm -v "$PWD/data:/data" \
  -e TWITCH_CLIENT_ID -e TWITCH_CLIENT_SECRET scanner ci
```

All thresholds — follower range, viewer range, minimum days seen, target regions, retention —
are constants at the top of the script.

## Known limits

- Sponsorships mostly live in Twitch panels and creator socials, which the API does not expose.
  Detected signals are real; their absence proves nothing.
- Location inference cannot distinguish a UK streamer from a European one streaming in English
  at the same hour, which is why hours-only guesses are labelled separately.
- Scheduled GitHub Actions runs are best-effort and often delayed by 5–30 minutes. Harmless for
  hourly sampling.

## Repository

```
streamer_finder_v3.py            the scanner
Dockerfile                       runtime image
.github/workflows/twitch-scan.yml    hourly collection
.github/workflows/weekly-summary.yml weekly roll-up as a GitHub issue
```
