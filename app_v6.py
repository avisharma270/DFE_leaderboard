"""
Battery Smart DFE Leaderboards - single-file Streamlit app (streaks edition).

WHAT CHANGED vs app_v4
    1. Targets are GONE. The 'membership_targets' tab is no longer read (that
       logic will be decided later). No target columns, bars or metric switch.
    2. Data tabs renamed to match the current workbook:
           Membership     -> roster + tier + June baseline (OB_June / KYC_June)
           OB_RAW          -> July live activity rows (scoring + streaks)
       OB_from_June (June) and DFE_details are REFERENCE ONLY and not scored.
    3. New task points:  KYC = 3,  Referral from DFE = 2,  Retrofitment = 1.
    4. NEW: streak points, added on top of task points.
         KYC streak  : a +Rs.50 streak bonus  -> 1 point
                       a +Rs.100 streak bonus -> 2 points
         OB  streak  : Part A (extra OB inside the 3-day window) -> 1 point
                       Part B (each OB from Day 4 while active)  -> 2 points
                       (tier-agnostic; Under-performers are NOT eligible)
    5. UI flow when a DFE is selected:
           rank + performance card
           ongoing streaks card  (+ a details/info block)
           ONE combined leaderboard (all cohorts together)
           the DFE's Cohort (tier) leaderboard
           the DFE's Zone leaderboard

DATA TABS USED (all in one Google workbook)
    Membership      -> roster + tier + June baseline
    OB_RAW          -> July activity rows (scoring + streaks)

SCORING (counted per USC Emp Code from OB_RAW, on/after LIVE_FROM)
    Onboarding    = rows whose 'sourceId'      holds that USC  -> 2 pts (Referral from DFE)
    Retrofitment  = rows whose 'retro_by'       holds that USC  -> 1 pt
    KYC           = rows whose 'newassignedto'  holds that USC  -> 3 pts
    + KYC streak points + OB streak points (see above)
    Channel-partner 'CP...' codes and blanks are ignored.

RUN LOCALLY
    pip install streamlit pandas openpyxl
    streamlit run app_v5.py
"""

import html
import re

import pandas as pd
import streamlit as st

# ===========================================================================
#  SOURCE + SETTINGS  --  edit these
# ===========================================================================
# "gsheet" -> read the live Google Sheet below (default)
# "csv"    -> read local files (see DATA_*_CSV below)
DATA_SOURCE = "gsheet"

# --- Google Sheet ("Anyone with the link -> Viewer") -----------------------
GSHEET_URL = "docs.google.com/spreadsheets/d/11n6G7t9xI6JgbLZP6ErGw0iPoCIaxB1j5SlO8QJFYmQ/edit?usp=sharing"
GSHEET_ID = ""             # fallback, only used if GSHEET_URL is left empty

# Tab names (must match the workbook exactly).
OB_RAW_SHEET_NAME = "OB_RAW"             # July live activity rows (scoring + streaks)
MEMBERSHIP_SHEET_NAME = "Membership"     # roster + tier + June baseline
# NOTE: 'OB_from_June' (June activity) and 'DFE_details' are reference tabs only
#       and are intentionally NOT read by the leaderboard.

# --- Local CSV fallback ----------------------------------------------------
DATA_OB_RAW_CSV = "data/ob_raw.csv"
DATA_MEMBERSHIP_CSV = "data/membership.csv"

# --- Task points -----------------------------------------------------------
POINTS = {"verification": 3, "onboarding": 2, "retrofit": 1}
CREDIT_COLUMNS = {                 # task -> OB_RAW column holding the DFE's USC id
    "onboarding":   "sourceId",
    "retrofit":     "retro_by",
    "verification": "newassignedto",
}
ZONE_COL = "zoneId"
DATE_COL = "liveDate"          # OB_RAW date column (used for cutoff + streaks)
LIVE_FROM = "2026-07-01"       # competition counts activity on/after this date only
DFE_CODE_PATTERN = r"^USC"     # only values like 'USC-1234' count (excludes 'CP...')

# --- Streak points ---------------------------------------------------------
# A "day" counts toward a streak if the DFE did >=1 of that activity that day.
# A run must reach STREAK_MIN_DAYS consecutive days before ANY streak point is
# earned. Once it does:
#   KYC:  days 1-3 -> each EXTRA KYC (2nd onward that day) = KYC_50_PTS
#         day 4+   -> 1st KYC = KYC_50_PTS, each extra = KYC_100_PTS
#   OB :  days 1-3 -> each EXTRA OB  (2nd onward that day) = OB_PARTA_PTS
#         day 4+   -> EVERY OB that day               = OB_PARTB_PTS
# A missed day (zero activity) breaks the run; the 3-day window must rebuild.
STREAK_MIN_DAYS = 3
KYC_50_PTS = 1                 # +Rs.50 KYC streak bonus  -> 1 point
KYC_100_PTS = 2                # +Rs.100 KYC streak bonus -> 2 points
OB_PARTA_PTS = 1               # OB Part A (extra OB in 3-day window) -> 1 point
OB_PARTB_PTS = 2               # OB Part B (each OB from Day 4)       -> 2 points
OB_STREAK_EXCLUDE = {"Underperformer"}   # tiers NOT eligible for OB streak points

# --- Membership tiers (read straight from the Membership sheet) ------------
TIERS = [
    {"key": "Gold",           "match": "gold",           "slug": "gold",   "prize": "Top prize"},
    {"key": "Silver",         "match": "silver",         "slug": "silver", "prize": "Mid prize"},
    {"key": "Bronze",         "match": "bronze",         "slug": "bronze", "prize": "Entry prize"},
    {"key": "Underperformer", "match": "underperformer", "slug": "under",  "prize": "Improvement plan"},
    {"key": "to be assigned", "match": "to be assigned", "slug": "tba",    "prize": "Pending review"},
]
TIER_ORDER = {t["key"]: i for i, t in enumerate(TIERS)}
DEFAULT_TIER = "to be assigned"

# --- Membership sheet columns ----------------------------------------------
MEM_EMPCODE_COL = "Emp Code"
MEM_NAME_COL = "Name"
MEM_ZONE_COL = "Zone ID"
MEM_CITY_COL = "City"
MEM_TIER_COL = "Membership"
MEM_JUNE_OB_COL = "OB_June"
MEM_JUNE_KYC_COL = "KYC_June"

# --- Branding --------------------------------------------------------------
APP_TITLE = "Battery Smart DFE Leaderboards"
APP_TAGLINE = "Charged Up. Powered Up. Winning Big."
PRIMARY_COLOR = "#0B8A3D"

# ===========================================================================
#  DATA LOADING
# ===========================================================================
TASKS = list(CREDIT_COLUMNS.keys())
_PATTERN = DFE_CODE_PATTERN or "^USC"
_SPLIT = re.compile(r"[\s,;/|]+")


def _sheet_id():
    if GSHEET_URL.strip():
        m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", GSHEET_URL)
        if m:
            return m.group(1)
    return GSHEET_ID.strip()


def _gsheet_xlsx_url():
    return f"https://docs.google.com/spreadsheets/d/{_sheet_id()}/export?format=xlsx"


@st.cache_data(ttl=300, show_spinner=False)
def _workbook_bytes():
    import urllib.request
    req = urllib.request.Request(_gsheet_xlsx_url(), headers={"User-Agent": "Mozilla/5.0"})
    return urllib.request.urlopen(req, timeout=180).read()


def _read_named(sheet_name, csv_path):
    if DATA_SOURCE == "gsheet":
        import io
        book = pd.ExcelFile(io.BytesIO(_workbook_bytes()))   # requires openpyxl
        if sheet_name not in book.sheet_names:
            raise KeyError(f"Tab '{sheet_name}' not found. Tabs present: {book.sheet_names}")
        df = pd.read_excel(book, sheet_name=sheet_name, dtype=str)
    else:
        df = pd.read_csv(csv_path, dtype=str)
    df = df.fillna("")
    df.columns = [c.strip() for c in df.columns]
    return df


def load_ob():
    return _read_named(OB_RAW_SHEET_NAME, DATA_OB_RAW_CSV)


def load_membership():
    return _read_named(MEMBERSHIP_SHEET_NAME, DATA_MEMBERSHIP_CSV)


def _norm(v):
    return str(v).strip().upper()


def _region(zone):
    m = re.match(r"[A-Za-z]+", str(zone))
    return m.group(0).upper() if m else ""


def _tier_key(value):
    v = str(value).strip().lower()
    for t in TIERS:
        if v == t["match"]:
            return t["key"]
    return DEFAULT_TIER


def _to_int(v):
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "nan", "NaN", "None"):
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


# ===========================================================================
#  TASK TALLY
# ===========================================================================
def _credit_pairs(ob, col):
    """One (code, zone) row per USC code found in `col`; group-size = row count."""
    df = ob[[col, ZONE_COL]].copy()
    df.columns = ["code", "zone"]
    df["code"] = df["code"].astype(str).str.upper().str.split(_SPLIT)
    df = df.explode("code")
    df["code"] = df["code"].str.strip()
    df = df[df["code"].str.match(_PATTERN, na=False)]
    df["zone"] = df["zone"].astype(str).str.strip()
    return df


def _tally(ob):
    """Return (per-DFE task counts dict, zone_mode Series, missing list)."""
    counts, zone_frames, missing = {}, [], []
    for task, col in CREDIT_COLUMNS.items():
        if col in ob.columns:
            pairs = _credit_pairs(ob, col)
            counts[task] = pairs.groupby("code").size()
            zone_frames.append(pairs)
        else:
            counts[task] = pd.Series(dtype="int64")
            missing.append((task, col))
    if zone_frames:
        allpairs = pd.concat(zone_frames, ignore_index=True)
        allpairs = allpairs[allpairs["zone"] != ""]
        zone_mode = (allpairs.groupby("code")["zone"]
                     .agg(lambda s: s.value_counts().index[0])) if not allpairs.empty else pd.Series(dtype=str)
    else:
        zone_mode = pd.Series(dtype=str)
    return counts, zone_mode, missing


def filter_live(ob):
    """Keep only OB rows on/after LIVE_FROM (by DATE_COL)."""
    if DATE_COL not in ob.columns:
        return ob
    d = pd.to_datetime(ob[DATE_COL], errors="coerce")
    return ob[d >= pd.Timestamp(LIVE_FROM)]


# ===========================================================================
#  STREAK ENGINE
# ===========================================================================
def _daily_counts(ob, col):
    """emp_code -> {date(Timestamp): count} for the given credit column."""
    if col not in ob.columns or DATE_COL not in ob.columns:
        return {}
    df = ob[[col, DATE_COL]].copy()
    df.columns = ["code", "date"]
    df["code"] = df["code"].astype(str).str.upper().str.split(_SPLIT)
    df = df.explode("code")
    df["code"] = df["code"].str.strip()
    df = df[df["code"].str.match(_PATTERN, na=False)]
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.normalize()
    df = df.dropna(subset=["date"])
    out = {}
    for (code, date), n in df.groupby(["code", "date"]).size().items():
        out.setdefault(code, {})[date] = int(n)
    return out


def _runs(day_map):
    """List of maximal consecutive-day runs; each run is a list of (date, count)."""
    if not day_map:
        return []
    days = sorted(day_map.keys())
    runs, cur = [], [(days[0], day_map[days[0]])]
    for d in days[1:]:
        if (d - cur[-1][0]).days == 1:
            cur.append((d, day_map[d]))
        else:
            runs.append(cur)
            cur = [(d, day_map[d])]
    runs.append(cur)
    return runs


def _kyc_streak_points(day_map):
    total = 0
    for run in _runs(day_map):
        if len(run) < STREAK_MIN_DAYS:
            continue
        for i, (_, c) in enumerate(run):
            if i < STREAK_MIN_DAYS:                       # days 1-3 (Part A)
                total += max(c - 1, 0) * KYC_50_PTS
            else:                                         # day 4+ (Part B)
                total += KYC_50_PTS + max(c - 1, 0) * KYC_100_PTS
    return total


def _ob_streak_points(day_map, tier):
    if tier in OB_STREAK_EXCLUDE:
        return 0
    total = 0
    for run in _runs(day_map):
        if len(run) < STREAK_MIN_DAYS:
            continue
        for i, (_, c) in enumerate(run):
            if i < STREAK_MIN_DAYS:                       # days 1-3 (Part A)
                total += max(c - 1, 0) * OB_PARTA_PTS
            else:                                         # day 4+ (Part B)
                total += c * OB_PARTB_PTS
    return total


def _current_run(day_map, as_of):
    """(length, last_active_day, is_ongoing) for the run ending at the latest day."""
    if not day_map:
        return 0, None, False
    last = _runs(day_map)[-1]
    last_day = last[-1][0]
    ongoing = as_of is not None and last_day == as_of
    return len(last), last_day, ongoing


def compute_streaks(ob, roster):
    """Return (info dict keyed by emp_code, as_of Timestamp)."""
    kyc_daily = _daily_counts(ob, CREDIT_COLUMNS["verification"])
    ob_daily = _daily_counts(ob, CREDIT_COLUMNS["onboarding"])
    all_dates = [d for m in (kyc_daily, ob_daily) for dm in m.values() for d in dm]
    as_of = max(all_dates) if all_dates else None
    tier_of = dict(zip(roster["emp_code"], roster["tier"]))
    info = {}
    for code in set(kyc_daily) | set(ob_daily):
        km, om = kyc_daily.get(code, {}), ob_daily.get(code, {})
        klen, kday, kon = _current_run(km, as_of)
        olen, oday, oon = _current_run(om, as_of)
        info[code] = {
            "kyc_streak_pts": _kyc_streak_points(km),
            "ob_streak_pts": _ob_streak_points(om, tier_of.get(code, DEFAULT_TIER)),
            "kyc_len": klen, "kyc_day": kday, "kyc_on": kon,
            "ob_len": olen, "ob_day": oday, "ob_on": oon,
        }
    return info, as_of


@st.cache_data(ttl=300, show_spinner=False)
def get_daily_activity():
    """Per-DFE day-by-day counts. Returns (kyc, onb, ret) dicts:
    each is {emp_code: {date(Timestamp): count}} over the live window."""
    ob = filter_live(load_ob())
    kyc = _daily_counts(ob, CREDIT_COLUMNS["verification"])
    onb = _daily_counts(ob, CREDIT_COLUMNS["onboarding"])
    ret = _daily_counts(ob, CREDIT_COLUMNS["retrofit"])
    return kyc, onb, ret


def _streak_states(day_map):
    """date -> 'off' | 'build' (day 1-3 of a qualifying run) | 'active' (day 4+)."""
    states = {}
    for run in _runs(day_map):
        if len(run) < STREAK_MIN_DAYS:
            for d, _ in run:
                states[d] = "off"
        else:
            for i, (d, _) in enumerate(run):
                states[d] = "build" if i < STREAK_MIN_DAYS else "active"
    return states


# ===========================================================================
#  ROSTER
# ===========================================================================
@st.cache_data(ttl=300, show_spinner="Reading memberships...")
def get_roster():
    mem = load_membership()
    need = [MEM_EMPCODE_COL, MEM_NAME_COL, MEM_ZONE_COL, MEM_TIER_COL]
    for c in need:
        if c not in mem.columns:
            raise KeyError(f"Column '{c}' not found in '{MEMBERSHIP_SHEET_NAME}'. "
                           f"Columns present: {list(mem.columns)}")
    r = pd.DataFrame()
    r["emp_code"] = mem[MEM_EMPCODE_COL].map(_norm)
    r["name"] = mem[MEM_NAME_COL].astype(str).str.strip()
    r["zone"] = mem[MEM_ZONE_COL].astype(str).str.strip()
    r["city"] = mem[MEM_CITY_COL].astype(str).str.strip() if MEM_CITY_COL in mem.columns else ""
    r["tier"] = mem[MEM_TIER_COL].map(_tier_key)
    r["june_ob"] = mem[MEM_JUNE_OB_COL].map(lambda v: _to_int(v) or 0) if MEM_JUNE_OB_COL in mem.columns else 0
    r["june_kyc"] = mem[MEM_JUNE_KYC_COL].map(lambda v: _to_int(v) or 0) if MEM_JUNE_KYC_COL in mem.columns else 0
    r = r[r["emp_code"] != ""].drop_duplicates("emp_code").reset_index(drop=True)
    r["region"] = r["zone"].map(_region)
    return r


# ===========================================================================
#  LIVE BOARD
# ===========================================================================
@st.cache_data(ttl=300, show_spinner="Loading leaderboard...")
def get_live_board():
    """Every DFE in Membership; task points + streak points from OB_RAW."""
    ob = filter_live(load_ob())
    counts, zone_mode, missing = _tally(ob)
    roster = get_roster()
    streaks, as_of = compute_streaks(ob, roster)

    board = roster.copy()
    for task in TASKS:
        board[task] = board["emp_code"].map(counts[task]).fillna(0).astype(int)
    board["zone"] = board["zone"].where(board["zone"] != "",
                                        board["emp_code"].map(zone_mode)).fillna("").replace("", "Unassigned")
    board["region"] = board["zone"].map(_region)

    board["base_points"] = sum(board[t] * POINTS[t] for t in TASKS)
    board["kyc_streak_pts"] = board["emp_code"].map(lambda c: streaks.get(c, {}).get("kyc_streak_pts", 0)).astype(int)
    board["ob_streak_pts"] = board["emp_code"].map(lambda c: streaks.get(c, {}).get("ob_streak_pts", 0)).astype(int)
    board["streak_pts"] = board["kyc_streak_pts"] + board["ob_streak_pts"]
    board["points"] = board["base_points"] + board["streak_pts"]

    # current-streak status for the card
    board["kyc_len"] = board["emp_code"].map(lambda c: streaks.get(c, {}).get("kyc_len", 0)).astype(int)
    board["ob_len"] = board["emp_code"].map(lambda c: streaks.get(c, {}).get("ob_len", 0)).astype(int)
    board["kyc_on"] = board["emp_code"].map(lambda c: bool(streaks.get(c, {}).get("kyc_on", False)))
    board["ob_on"] = board["emp_code"].map(lambda c: bool(streaks.get(c, {}).get("ob_on", False)))
    board["kyc_day"] = board["emp_code"].map(lambda c: streaks.get(c, {}).get("kyc_day"))
    board["ob_day"] = board["emp_code"].map(lambda c: streaks.get(c, {}).get("ob_day"))

    board["tier_order"] = board["tier"].map(TIER_ORDER).fillna(len(TIERS)).astype(int)
    tie = ["points", "onboarding", "verification", "retrofit", "emp_code"]
    asc = [False, False, False, False, True]

    # single combined ranking (all cohorts together)
    board = board.sort_values(tie, ascending=asc).reset_index(drop=True)
    board["overall_rank"] = range(1, len(board) + 1)
    # cohort (tier) ranking, pan-India
    board["cohort_rank"] = board.groupby("tier").cumcount() + 1
    # zone ranking (all tiers within a zone)
    board = board.sort_values(["zone"] + tie, ascending=[True] + asc)
    board["zone_rank"] = board.groupby("zone").cumcount() + 1

    board = board.sort_values("overall_rank").reset_index(drop=True)
    return board, missing, as_of


def _esc(v):
    return html.escape(str(v))


def _tier_meta(key):
    for t in TIERS:
        if t["key"] == key:
            return t
    return TIERS[-1]


def _fmt_day(d):
    try:
        return pd.Timestamp(d).strftime("%d %b")
    except Exception:  # noqa: BLE001
        return "-"


# ===========================================================================
#  UI
# ===========================================================================
st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")

STYLE = """
<style>
:root{
--green:#0B8A3D; --green2:#12A94B; --dark:#0A5C2C;
--v:#6C5CE7; --o:#2E86DE; --r:#E1A200; --streak:#F2542D; --streak2:#FF8A3D;
--gold:#F2B705; --silver:#9AA7B2; --bronze:#C77B30; --under:#E15554; --tba:#7A8794;
--ink:#12261B; --muted:#6b7c72; --line:#e6ede8;
}
.block-container{padding-top:1.1rem; padding-bottom:2rem; max-width:1080px;}
.app-header{background:linear-gradient(120deg,var(--green) 0%,var(--green2) 55%,#16c057 100%);
color:#fff; border-radius:18px; padding:1.15rem 1.35rem; box-shadow:0 10px 26px rgba(11,138,61,.28);}
.app-header h1{font-size:1.75rem; font-weight:800; margin:0; letter-spacing:-.02em;}
.app-header p{margin:.28rem 0 0; opacity:.93; font-size:.95rem;}
.stat-grid{display:flex; flex-wrap:wrap; gap:.6rem; margin:.9rem 0 .2rem;}
.stat-card{flex:1 1 120px; border-radius:14px; padding:.72rem .95rem; color:#fff; box-shadow:0 5px 15px rgba(0,0,0,.09);}
.stat-card .n{font-size:1.55rem; font-weight:800; line-height:1;}
.stat-card .l{font-size:.78rem; opacity:.95; margin-top:.22rem;}
.sc-dfe{background:linear-gradient(135deg,#0A5C2C,#0B8A3D);}
.sc-o{background:linear-gradient(135deg,#1c63b0,#2E86DE);}
.sc-v{background:linear-gradient(135deg,#5544cf,#6C5CE7);}
.sc-r{background:linear-gradient(135deg,#b98200,#E1A200);}
.sc-s{background:linear-gradient(135deg,#d63a17,#FF8A3D);}
.tier-grid{display:flex; flex-wrap:wrap; gap:.5rem; margin:.35rem 0 .2rem;}
.tier-card{flex:1 1 100px; border-radius:12px; padding:.55rem .7rem; color:#fff;}
.tier-card .n{font-size:1.35rem; font-weight:800; line-height:1;}
.tier-card .l{font-size:.72rem; opacity:.96; margin-top:.18rem;}
.tc-gold{background:linear-gradient(135deg,#caa100,#F2B705);}
.tc-silver{background:linear-gradient(135deg,#7c8894,#9AA7B2);}
.tc-bronze{background:linear-gradient(135deg,#a2621f,#C77B30);}
.tc-under{background:linear-gradient(135deg,#b73f3d,#E15554);}
.tc-tba{background:linear-gradient(135deg,#5e6b76,#7A8794);}
.me-card{border-radius:16px; padding:1rem 1.1rem; margin:.5rem 0 .2rem;
background:linear-gradient(135deg,#ffffff,#eef8f1); border:1px solid var(--line); box-shadow:0 8px 20px rgba(11,138,61,.12);}
.me-top{display:flex; flex-wrap:wrap; align-items:center; gap:.5rem 1rem;}
.me-name{font-size:1.25rem; font-weight:800; color:var(--ink);}
.me-sub{color:var(--muted); font-size:.85rem;}
.me-ranks{display:flex; gap:.55rem; flex-wrap:wrap; margin-left:auto;}
.rk{background:var(--green); color:#fff; border-radius:12px; padding:.4rem .75rem; text-align:center; min-width:92px;}
.rk.alt{background:var(--dark);} .rk.z{background:#127a56;}
.rk .b{font-size:1.3rem; font-weight:800; line-height:1;}
.rk .s{font-size:.68rem; opacity:.92; margin-top:.15rem;}
.me-chips{display:flex; gap:.5rem; flex-wrap:wrap; margin-top:.75rem;}
.mchip{border-radius:12px; padding:.5rem .75rem; color:#fff; flex:1 1 84px; min-width:84px;}
.mchip .n{font-size:1.2rem; font-weight:800; line-height:1;}
.mchip .l{font-size:.7rem; opacity:.95; margin-top:.18rem;}
.bg-v{background:var(--v);} .bg-o{background:var(--o);} .bg-r{background:var(--r);}
.bg-s{background:linear-gradient(135deg,#d63a17,#FF8A3D);}
.bg-t{background:linear-gradient(135deg,#0A5C2C,#0B8A3D);}
.reason{background:#f3f6f4; border-left:4px solid var(--green); border-radius:8px; padding:.6rem .8rem; margin-top:.65rem; font-size:.85rem; color:var(--ink); line-height:1.5;}
/* streak card */
.streak-card{border-radius:16px; padding:1rem 1.1rem; margin:.6rem 0 .2rem;
background:linear-gradient(135deg,#fff7f2,#fff); border:1px solid #ffd9c4; box-shadow:0 8px 20px rgba(242,84,45,.10);}
.streak-head{font-weight:800; color:#b5350f; font-size:1.02rem; display:flex; align-items:center; gap:.4rem; margin-bottom:.55rem;}
.streak-tiles{display:flex; gap:.6rem; flex-wrap:wrap;}
.stile{flex:1 1 220px; border-radius:14px; padding:.75rem .85rem; border:1px solid var(--line); background:#fff;}
.stile .h{font-weight:800; color:var(--ink); font-size:.9rem; display:flex; align-items:center; gap:.4rem;}
.stile .flame{filter:grayscale(1); opacity:.35;}
.stile.on{border-color:#ffc7a8; background:linear-gradient(135deg,#fff2ea,#fff);}
.stile.on .flame{filter:none; opacity:1;}
.stile.build{border-color:#ffe6b0; background:linear-gradient(135deg,#fffaf0,#fff);}
.stile .state{font-size:.8rem; margin-top:.3rem;}
.stile .state b{color:#b5350f;}
.stile .foot{font-size:.72rem; color:var(--muted); margin-top:.25rem;}
.stile .pts{float:right; font-weight:800; color:#d63a17; font-size:1.15rem;}
.stile .pts small{font-size:.6rem; color:var(--muted); font-weight:700;}
.lb-scroll{max-height:620px; overflow-y:auto; padding:.15rem;}
.lb-row{display:flex; align-items:center; gap:.55rem; padding:.5rem .6rem; border-radius:12px; margin-bottom:.42rem;
background:#fff; border:1px solid var(--line);}
.lb-row:nth-child(even){background:#f7faf8;}
.lb-row.top1{border:1px solid #f4d770; background:linear-gradient(90deg,#fff8e2,#fff);}
.lb-row.top2{border:1px solid #d7dee4; background:linear-gradient(90deg,#f3f6f8,#fff);}
.lb-row.top3{border:1px solid #e6c39a; background:linear-gradient(90deg,#fdf0e3,#fff);}
.lb-row.me{border:2px solid var(--green); background:linear-gradient(90deg,#e8faef,#fff);}
.badge{width:34px; height:34px; border-radius:50%; display:flex; align-items:center; justify-content:center;
font-weight:800; font-size:.9rem; color:var(--dark); background:#e7f3ec; flex:0 0 auto;}
.badge.g{background:var(--gold); color:#5a4600;} .badge.s{background:var(--silver); color:#243040;} .badge.b{background:var(--bronze); color:#fff;}
.lb-main{flex:1 1 auto; min-width:0;}
.lb-name{font-weight:700; color:var(--ink); font-size:.95rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.lb-sub{color:var(--muted); font-size:.75rem;}
.lb-bar{height:6px; background:#eef2ef; border-radius:6px; margin-top:.28rem; overflow:hidden;}
.lb-bar span{display:block; height:100%; background:linear-gradient(90deg,var(--green),var(--green2)); border-radius:6px;}
.lb-chips{display:flex; gap:.32rem; flex:0 0 auto;}
.chip{font-size:.72rem; font-weight:800; color:#fff; border-radius:8px; padding:.22rem .42rem; min-width:36px; text-align:center;}
.chip.v{background:var(--v);} .chip.o{background:var(--o);} .chip.r{background:var(--r);}
.chip.s{background:linear-gradient(135deg,#d63a17,#FF8A3D);}
.lb-pts{flex:0 0 auto; text-align:right; min-width:54px; font-weight:800; color:var(--green); font-size:1.08rem;}
.lb-pts small{display:block; font-size:.62rem; color:var(--muted); font-weight:600;}
.legend{display:flex; gap:.8rem; flex-wrap:wrap; margin:.2rem 0 .6rem; font-size:.78rem; color:var(--muted);}
.legend b{color:var(--ink);}
.dot{display:inline-block; width:10px; height:10px; border-radius:3px; margin-right:.3rem; vertical-align:middle;}
@media (max-width:640px){
.app-header h1{font-size:1.35rem;}
.stat-card{flex:1 1 44%;}
.tier-card{flex:1 1 30%;}
.me-ranks{margin-left:0;}
.lb-row{flex-wrap:wrap;}
.lb-main{flex:1 1 62%;}
.lb-chips{order:3; width:100%; margin-top:.4rem;}
.chip{flex:1 1 auto;}
.lb-pts{min-width:44px;}
}
.co-head{border-radius:14px; padding:.7rem 1rem; margin:1rem 0 .5rem; color:#fff;}
.co-head .t{font-size:1.05rem; font-weight:800;}
.co-head .m{font-size:.78rem; opacity:.95; margin-top:.15rem;}
.co-gold{background:linear-gradient(135deg,#caa100,#F2B705);}
.co-silver{background:linear-gradient(135deg,#7c8894,#9AA7B2);}
.co-bronze{background:linear-gradient(135deg,#a2621f,#C77B30);}
.co-under{background:linear-gradient(135deg,#b73f3d,#E15554);}
.co-tba{background:linear-gradient(135deg,#5e6b76,#7A8794);}
.cobadge{display:inline-block; padding:.1rem .45rem; border-radius:7px; font-weight:800; color:#fff; font-size:.72rem;}
.cobadge.gold{background:#caa100;} .cobadge.silver{background:#7c8894;} .cobadge.bronze{background:#a2621f;}
.cobadge.under{background:#b73f3d;} .cobadge.tba{background:#5e6b76;}
.sec-title{font-weight:800; color:var(--ink); margin:1rem 0 .3rem; font-size:1.0rem;}
.tracker{border-radius:16px; padding:.8rem .9rem; margin:.4rem 0 .2rem; background:#fff; border:1px solid var(--line); box-shadow:0 6px 16px rgba(11,138,61,.08);}
.tk-scroll{overflow-x:auto; padding-bottom:.2rem;}
.tk-scroll svg{display:block;}
.tk-legend{display:flex; gap:.7rem; flex-wrap:wrap; align-items:center; font-size:.74rem; color:var(--muted); margin-bottom:.5rem;}
.tk-legend span{display:inline-flex; align-items:center; gap:.32rem;}
.tk-legend i{width:11px; height:11px; border-radius:3px; display:inline-block;}
.tk-legend .sep{width:1px; height:14px; background:var(--line); margin:0 .1rem;}
.tk-empty{color:var(--muted); font-size:.86rem; padding:.6rem .2rem;}
</style>
"""


def me_card_html(me, n_total, tier_size, zone_size):
    key = me["tier"]
    meta = _tier_meta(key)
    name = _esc(me["name"] or me["emp_code"])
    loc = _esc(me["city"] or me["zone"])
    return (
        "<div class='me-card'><div class='me-top'>"
        f"<div><div class='me-name'>{name}</div>"
        f"<div class='me-sub'>{_esc(me['emp_code'])} &middot; {loc} &middot; "
        f"<span class='cobadge {meta['slug']}'>{_esc(key)}</span></div></div>"
        "<div class='me-ranks'>"
        f"<div class='rk'><div class='b'>#{int(me['overall_rank'])}</div><div class='s'>Overall ({n_total})</div></div>"
        f"<div class='rk alt'><div class='b'>#{int(me['cohort_rank'])}</div><div class='s'>{_esc(key)} ({tier_size})</div></div>"
        f"<div class='rk z'><div class='b'>#{int(me['zone_rank'])}</div><div class='s'>Zone {_esc(me['zone'])} ({zone_size})</div></div>"
        "</div></div>"
        "<div class='me-chips'>"
        f"<div class='mchip bg-o'><div class='n'>{int(me['onboarding'])}</div><div class='l'>Referral</div></div>"
        f"<div class='mchip bg-v'><div class='n'>{int(me['verification'])}</div><div class='l'>KYC</div></div>"
        f"<div class='mchip bg-r'><div class='n'>{int(me['retrofit'])}</div><div class='l'>Retrofit</div></div>"
        f"<div class='mchip bg-s'><div class='n'>{int(me['streak_pts'])}</div><div class='l'>Streak pts</div></div>"
        f"<div class='mchip bg-t'><div class='n'>{int(me['points'])}</div><div class='l'>Total points</div></div>"
        "</div>"
        f"<div class='reason'>Cohort <b>{_esc(key)}</b> (from the sheet) &middot; "
        f"Task points <b>{int(me['base_points'])}</b> + streak points <b>{int(me['streak_pts'])}</b> "
        f"= <b>{int(me['points'])}</b> total &middot; "
        f"June baseline: <b>{int(me['june_ob'])}</b> OB / <b>{int(me['june_kyc'])}</b> KYC</div>"
        "</div>"
    )


def _stile(label, length, day, ongoing, pts):
    if length >= STREAK_MIN_DAYS:
        cls, state = "on", f"<b>Active</b> &middot; {length}-day streak"
    elif length > 0:
        cls, state = "build", f"Building &middot; {length}/{STREAK_MIN_DAYS} days"
    else:
        cls, state = "", "No active streak"
    if length > 0 and ongoing:
        foot = "Live right now &mdash; keep it going today."
    elif length > 0:
        foot = f"Last active {_fmt_day(day)} &middot; a missed day breaks it."
    else:
        foot = "Do this on 3 days in a row to start earning."
    return (
        f"<div class='stile {cls}'>"
        f"<span class='pts'>{int(pts)}<small> pts</small></span>"
        f"<div class='h'><span class='flame'>&#128293;</span>{_esc(label)}</div>"
        f"<div class='state'>{state}</div>"
        f"<div class='foot'>{foot}</div>"
        "</div>"
    )


def streak_card_html(me):
    return (
        "<div class='streak-card'>"
        "<div class='streak-head'><span>&#9889;</span>Your ongoing streaks</div>"
        "<div class='streak-tiles'>"
        + _stile("KYC streak", int(me["kyc_len"]), me["kyc_day"], bool(me["kyc_on"]), int(me["kyc_streak_pts"]))
        + _stile("Referral streak", int(me["ob_len"]), me["ob_day"], bool(me["ob_on"]), int(me["ob_streak_pts"]))
        + "</div></div>"
    )


# tracker geometry
_TK_SLOT = 34        # px per day column
_TK_BARW = 8         # px per bar
_TK_H = 130          # px bar area height
_TK_TOP = 14         # px headroom above tallest bar
_TK_RIBBON = 15      # px per streak ribbon row
_TK_XLAB = 26        # px x-axis label strip


def daily_tracker_html(emp_code, kyc, onb, ret, start, end):
    """Grouped daily bars (KYC/Onb/Ret) + KYC and OB streak ribbons, as SVG."""
    km = kyc.get(emp_code, {})
    om = onb.get(emp_code, {})
    rm = ret.get(emp_code, {})
    if not (km or om or rm):
        return ("<div class='tracker'><div class='tk-empty'>"
                "No July activity recorded yet. Do a KYC, onboarding or retrofit to start the chart.</div></div>")

    start = pd.Timestamp(start).normalize()
    end = pd.Timestamp(end).normalize()
    if end < start:
        end = start
    days = list(pd.date_range(start, end, freq="D"))
    n = len(days)

    maxc = 1
    for d in days:
        maxc = max(maxc, km.get(d, 0), om.get(d, 0), rm.get(d, 0))

    kstate = _streak_states(km)
    ostate = _streak_states(om)

    left, right = 30, 12
    width = left + right + n * _TK_SLOT
    ribbons_y = _TK_TOP + _TK_H + _TK_XLAB
    height = ribbons_y + 2 * (_TK_RIBBON + 6) + 24

    series = [("verification", km, "var(--v)"), ("onboarding", om, "var(--o)"), ("retrofit", rm, "var(--r)")]
    parts = [f"<svg viewBox='0 0 {width} {height}' width='{width}' height='{height}' "
             "xmlns='http://www.w3.org/2000/svg' font-family='inherit'>"]

    # y grid: baseline + max line
    base_y = _TK_TOP + _TK_H
    parts.append(f"<line x1='{left}' y1='{base_y}' x2='{width-right}' y2='{base_y}' stroke='#dce6df'/>")
    parts.append(f"<line x1='{left}' y1='{_TK_TOP}' x2='{width-right}' y2='{_TK_TOP}' stroke='#eef2ef' stroke-dasharray='3 3'/>")
    parts.append(f"<text x='{left-6}' y='{_TK_TOP+4}' text-anchor='end' font-size='9' fill='#9aa8a0'>{maxc}</text>")
    parts.append(f"<text x='{left-6}' y='{base_y+3}' text-anchor='end' font-size='9' fill='#9aa8a0'>0</text>")

    for i, d in enumerate(days):
        cx = left + i * _TK_SLOT + _TK_SLOT / 2
        k, o, r = km.get(d, 0), om.get(d, 0), rm.get(d, 0)
        # three grouped bars
        group_w = 3 * _TK_BARW + 2 * 2
        x0 = cx - group_w / 2
        for j, (_key, m, col) in enumerate(series):
            c = m.get(d, 0)
            bh = (c / maxc) * (_TK_H - 2) if maxc else 0
            bx = x0 + j * (_TK_BARW + 2)
            by = base_y - bh
            if c > 0:
                parts.append(f"<rect x='{bx:.1f}' y='{by:.1f}' width='{_TK_BARW}' height='{bh:.1f}' "
                             f"rx='2' fill='{col}'><title>{d.strftime('%d %b')} &#183; "
                             f"KYC {k}, Onb {o}, Ret {r}</title></rect>")
        # x label (day number; show month on the 1st and every 5th)
        lab = d.strftime("%d")
        parts.append(f"<text x='{cx:.1f}' y='{base_y + 15}' text-anchor='middle' font-size='9' fill='#6b7c72'>{lab}</text>")

    # ribbons
    def ribbon(y, label, state_map, base_col):
        parts.append(f"<text x='{left-6}' y='{y+11}' text-anchor='end' font-size='9' fill='#6b7c72'>{label}</text>")
        for i, d in enumerate(days):
            x = left + i * _TK_SLOT + 3
            w = _TK_SLOT - 6
            s = state_map.get(d, "off")
            if s == "active":
                fill, op = base_col, "1"
            elif s == "build":
                fill, op = base_col, "0.4"
            else:
                fill, op = "#e6ede8", "1"
            parts.append(f"<rect x='{x:.1f}' y='{y}' width='{w:.1f}' height='{_TK_RIBBON}' rx='3' "
                         f"fill='{fill}' fill-opacity='{op}'/>")

    ry1 = ribbons_y
    ry2 = ribbons_y + _TK_RIBBON + 6
    ribbon(ry1, "KYC streak", kstate, "#6C5CE7")
    ribbon(ry2, "OB streak", ostate, "#2E86DE")

    parts.append("</svg>")
    svg = "".join(parts)
    legend = (
        "<div class='tk-legend'>"
        "<span><i style='background:var(--v)'></i>KYC</span>"
        "<span><i style='background:var(--o)'></i>Referral</span>"
        "<span><i style='background:var(--r)'></i>Retrofit</span>"
        "<span class='sep'></span>"
        "<span><i style='background:#8579ec'></i>streak building (day 1-3)</span>"
        "<span><i style='background:#6C5CE7'></i>streak active (day 4+)</span>"
        "</div>"
    )
    return f"<div class='tracker'>{legend}<div class='tk-scroll'>{svg}</div></div>"


LEGEND = (
    "<div class='legend'>"
    "<span><span class='dot' style='background:#6C5CE7'></span><b>KYC</b> = 3 pts</span>"
    "<span><span class='dot' style='background:#2E86DE'></span><b>Ref</b> = 2 pts</span>"
    "<span><span class='dot' style='background:#E1A200'></span><b>Ret</b> = 1 pt</span>"
    "<span><span class='dot' style='background:#F2542D'></span><b>Streak</b> bonus</span>"
    "</div>"
)


def leaderboard_html(df, rank_col, highlight_id):
    if df.empty:
        return "<div class='lb-scroll'><div class='lb-sub' style='padding:1rem'>No one found.</div></div>"
    maxpts = max(int(df["points"].max()), 1)
    parts = ["<div class='lb-scroll'>"]
    for rec in df.to_dict("records"):
        rank = int(rec[rank_col])
        pts = int(rec["points"])
        pct = int(round(pts / maxpts * 100))
        name = _esc(rec["name"] or rec["emp_code"])
        rowcls, badgecls = "lb-row", "badge"
        if rank == 1:
            rowcls += " top1"; badgecls += " g"
        elif rank == 2:
            rowcls += " top2"; badgecls += " s"
        elif rank == 3:
            rowcls += " top3"; badgecls += " b"
        if highlight_id and rec["emp_code"] == highlight_id:
            rowcls += " me"
        streak_chip = ""
        if int(rec.get("streak_pts", 0)) > 0:
            streak_chip = f"<span class='chip s' title='Streak points'>&#128293;{int(rec['streak_pts'])}</span>"
        parts.append(
            f"<div class='{rowcls}'>"
            f"<div class='{badgecls}'>{rank}</div>"
            f"<div class='lb-main'><div class='lb-name'>{name}</div>"
            f"<div class='lb-sub'>{_esc(rec['emp_code'])} &middot; {_esc(rec['zone'])} &middot; "
            f"<span class='cobadge {_tier_meta(rec['tier'])['slug']}'>{_esc(rec['tier'])}</span></div>"
            f"<div class='lb-bar'><span style='width:{pct}%'></span></div></div>"
            "<div class='lb-chips'>"
            f"<span class='chip v' title='KYC'>KYC {int(rec['verification'])}</span>"
            f"<span class='chip o' title='Referral from DFE'>Ref {int(rec['onboarding'])}</span>"
            f"<span class='chip r' title='Retrofit'>Ret {int(rec['retrofit'])}</span>"
            f"{streak_chip}"
            "</div>"
            f"<div class='lb-pts'>{pts}<small>pts</small></div>"
            "</div>"
        )
    parts.append("</div>")
    return "".join(parts)


def tier_header_html(meta, sub):
    if len(sub):
        pts = sub["points"]
        rng = f"{int(pts.min())}-{int(pts.max())}"
        avg = f"{pts.mean():.0f}"
    else:
        rng, avg = "-", "-"
    return (
        f"<div class='co-head co-{meta['slug']}'>"
        f"<div class='t'>{_esc(meta['key'])}</div>"
        f"<div class='m'>{len(sub)} DFEs &middot; points {rng} (avg {avg}) &middot; Prize: {_esc(meta['prize'])}</div>"
        "</div>"
    )


st.markdown(STYLE, unsafe_allow_html=True)
st.markdown(f"<div class='app-header'><h1>{_esc(APP_TITLE)}</h1><p>{_esc(APP_TAGLINE)}</p></div>",
            unsafe_allow_html=True)

try:
    board, missing, as_of = get_live_board()
except Exception as exc:  # noqa: BLE001
    hint = ""
    if DATA_SOURCE == "gsheet":
        hint = ("\n\nIt tried to download: "
                f"{_gsheet_xlsx_url()}\n\n"
                f"This app needs the '{OB_RAW_SHEET_NAME}' and '{MEMBERSHIP_SHEET_NAME}' tabs.\n"
                "1) Make sure GSHEET_URL is your exact sheet URL and the sheet is link-shared.\n"
                "2) Confirm both tab names are exact.\n"
                "3) CSV mode: put the two CSVs in a 'data' folder and set DATA_SOURCE='csv'.")
    st.error(f"Could not load the data (DATA_SOURCE = '{DATA_SOURCE}').{hint}\n\nDetails: {exc}")
    st.stop()

n_comp = len(board)

# ===========================================================================
#  LEFT MENU (sidebar navigation) + shared helpers
# ===========================================================================
with st.sidebar:
    st.markdown("### 📊 Leaderboards")
    view = st.radio(
        "Choose a view",
        ["Pan India Leaderboard", "Cohort Specific Leaderboard", "Zone-wise Leaderboard"],
        label_visibility="collapsed",
    )
    st.divider()
    st.subheader("How points work")
    st.markdown(
        f"**Task points**\n"
        f"- KYC = **{POINTS['verification']} pts**\n"
        f"- Referral from DFE = **{POINTS['onboarding']} pts**\n"
        f"- Retrofitment = **{POINTS['retrofit']} pt**\n\n"
        f"**Streak points** (added on top)\n"
        f"- KYC: +Rs.50 bonus = **{KYC_50_PTS} pt**, +Rs.100 bonus = **{KYC_100_PTS} pts**\n"
        f"- Referral: Part A = **{OB_PARTA_PTS} pt/extra**, Part B = **{OB_PARTB_PTS} pts each**\n"
        f"- A streak needs **{STREAK_MIN_DAYS} days in a row**; a missed day resets it.\n"
        f"- Under-performers don't earn Referral streak points."
    )
    st.caption(f"Counted from {LIVE_FROM}" + (f" - latest activity {_fmt_day(as_of)}" if as_of is not None else ""))
    st.caption(f"Cohorts from '{MEMBERSHIP_SHEET_NAME}' - activity from '{OB_RAW_SHEET_NAME}'")
    st.caption(f"Data: {DATA_SOURCE}")
    if st.button("Refresh data", width='stretch'):
        st.cache_data.clear()
        st.rerun()


def stats_strip(df):
    """A small summary strip of totals for whatever slice of the board is passed."""
    return (
        "<div class='stat-grid'>"
        f"<div class='stat-card sc-dfe'><div class='n'>{len(df):,}</div><div class='l'>DFEs</div></div>"
        f"<div class='stat-card sc-o'><div class='n'>{int(df['onboarding'].sum()):,}</div><div class='l'>Referrals</div></div>"
        f"<div class='stat-card sc-v'><div class='n'>{int(df['verification'].sum()):,}</div><div class='l'>KYC</div></div>"
        f"<div class='stat-card sc-r'><div class='n'>{int(df['retrofit'].sum()):,}</div><div class='l'>Retrofit</div></div>"
        f"<div class='stat-card sc-s'><div class='n'>{int(df['points'].sum()):,}</div><div class='l'>Total pts</div></div>"
        "</div>"
    )


if missing:
    cols = ", ".join(f"'{c}' (for {t})" for t, c in missing)
    st.info(f"Credit columns not found in '{OB_RAW_SHEET_NAME}', so those tasks score 0: {cols}.")


# ---------------------------------------------------------------- Pan India
if view == "Pan India Leaderboard":
    st.markdown("<div class='sec-title'>Pan India Leaderboard &mdash; all cohorts combined</div>",
                unsafe_allow_html=True)
    st.markdown(stats_strip(board), unsafe_allow_html=True)
    st.caption(f"Every DFE across India ranked together by total points. {n_comp} DFEs, counted from {LIVE_FROM}.")
    st.markdown(LEGEND, unsafe_allow_html=True)

    q = st.text_input("Search", key="india_q",
                      placeholder="Filter by name, ID, zone or cohort (optional)")
    combined = board.sort_values("overall_rank")
    if q.strip():
        mask = combined.apply(lambda r: r.astype(str).str.contains(q, case=False, na=False).any(), axis=1)
        combined = combined[mask]
        st.caption(f"{len(combined)} DFE(s) match '{q}'.")
    st.markdown(leaderboard_html(combined, "overall_rank", ""), unsafe_allow_html=True)

# ---------------------------------------------------------------- Cohort
elif view == "Cohort Specific Leaderboard":
    st.markdown("<div class='sec-title'>Cohort Specific Leaderboard</div>", unsafe_allow_html=True)
    st.caption("Pick a cohort to see only that cohort's DFEs, ranked within the cohort.")

    # cohort overview cards (counts per tier)
    tier_counts = board["tier"].value_counts().to_dict()
    tcards = ["<div class='tier-grid'>"]
    for t in TIERS:
        c = tier_counts.get(t["key"], 0)
        tcards.append(
            f"<div class='tier-card tc-{t['slug']}'><div class='n'>{c}</div>"
            f"<div class='l'>{_esc(t['key'])}</div></div>")
    tcards.append("</div>")
    st.markdown("".join(tcards), unsafe_allow_html=True)

    present = [t["key"] for t in TIERS if (board["tier"] == t["key"]).any()]
    if not present:
        st.info("No cohorts found in the data yet.")
    else:
        cohort = st.selectbox("Select a cohort", options=present, index=0)
        meta = _tier_meta(cohort)
        sub = board[board["tier"] == cohort].sort_values("cohort_rank")
        st.markdown(tier_header_html(meta, sub), unsafe_allow_html=True)
        st.markdown(LEGEND, unsafe_allow_html=True)
        st.caption(f"{len(sub)} DFEs in the {cohort} cohort, ranked by points within the cohort.")
        st.markdown(leaderboard_html(sub, "cohort_rank", ""), unsafe_allow_html=True)

# ---------------------------------------------------------------- Zone-wise
else:
    st.markdown("<div class='sec-title'>Zone-wise Leaderboard</div>", unsafe_allow_html=True)
    st.caption("Search for a zone below, pick it, and see that zone's leaderboard (all cohorts together).")

    zones = sorted(z for z in board["zone"].astype(str).unique() if z.strip() and z != "Unassigned")
    if (board["zone"] == "Unassigned").any():
        zones.append("Unassigned")

    query = st.text_input("Search Zone", key="zone_q",
                          placeholder="Type part of a zone, e.g. DEL, MUM, BLR ...")
    if query.strip():
        matches = [z for z in zones if query.strip().lower() in z.lower()]
    else:
        matches = zones

    if not matches:
        st.warning(f"No zone matches '{query}'. Try a shorter or different search term.")
    else:
        label = f"Select a zone ({len(matches)} match)" if query.strip() else f"Select a zone ({len(matches)} total)"
        zone = st.selectbox(label, options=matches, index=0)
        zl = board[board["zone"] == zone].sort_values("zone_rank")
        st.markdown(stats_strip(zl), unsafe_allow_html=True)
        st.caption(f"{len(zl)} DFEs in zone {zone}, ranked by points within the zone.")
        st.markdown(LEGEND, unsafe_allow_html=True)
        st.markdown(leaderboard_html(zl, "zone_rank", ""), unsafe_allow_html=True)

st.caption(f"Cohorts fixed from '{MEMBERSHIP_SHEET_NAME}'. Activity + streaks from '{OB_RAW_SHEET_NAME}'. "
           f"Points counted from {LIVE_FROM}.")
