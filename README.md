# Battery Smart DFE Leaderboards

A single-file [Streamlit](https://streamlit.io/) app that reads live activity from a
Google Sheet and ranks DFEs on a points system.

## Points

| Task | Points |
|------|--------|
| KYC | 3 |
| Referral from DFE | 2 |
| Retrofitment | 1 |

Streak points are added on top of task points:

- **KYC streak** — a +Rs.50 bonus = 1 pt, a +Rs.100 bonus = 2 pts
- **Referral streak** — Part A = 1 pt, Part B = 2 pts each (Under-performers not eligible)

A streak needs 3 qualifying days in a row; a missed day resets it.

## Views (left menu)

1. **My Performance** — search a DFE and see their Pan India, Cohort and Zone ranks, streaks,
   a day-by-day chart, and each leaderboard with that DFE highlighted.
2. **Pan India Leaderboard** — every DFE ranked together.
3. **Cohort Basis Leaderboard** — pick a cohort from a dropdown.
4. **Zone-wise Leaderboard** — search a zone and view its leaderboard.

## Data source

The app reads two tabs from one Google workbook: `OB_RAW` (July activity) and
`Membership` (roster + cohort + June baseline). The workbook must be shared as
**"Anyone with the link → Viewer"** so the app can download it. The sheet URL is set in
`app_v6.py` (`GSHEET_URL`). Data is cached for 5 minutes; use the **Refresh data** button
in the sidebar to force a reload.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app_v6.py
```

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub (see steps below).
2. Go to https://share.streamlit.io and sign in with GitHub.
3. **Create app → Deploy a public app from GitHub.**
4. Pick your repo and branch, and set **Main file path** to `app_v6.py`.
5. Click **Deploy**. Streamlit installs `requirements.txt` automatically.

No secrets are required because the sheet is link-shared for viewing.
