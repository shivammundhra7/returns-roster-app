"""
Putaway Roster Generator  —  CP-SAT engine (drop-in replacement for the greedy V16 loop)

WHAT CHANGED vs the old code
  - The whole month is solved at once by a constraint solver (Google OR-Tools CP-SAT),
    so it can never paint itself into a month-end corner the way day-by-day greedy did.
  - Hard rules (compliance, contracts) are HARD constraints. Everything else is a
    weighted objective, so the model ALWAYS returns a roster and a plain-language report
    of anything it had to bend — instead of silently violating a rule.
  - Date parsing is day-first (so "01-07-2026" is 1 July, not 7 Jan), load values with
    commas are handled, and night preferences read clean Min/Max columns if present.

HOW THE THREE GOALS ARE MODELLED
  1. Night split 47-49%  -> top-weighted soft target, per role per day (a FLOOR and a ceiling).
  2. Night preference     -> per-person monthly band, weighted BELOW staffing (worst-case only).
  3. Load-weighted WOs    -> per day per role, week-offs kept in 8-16% of active, with the
                             target inside that band interpolated from the day's load
                             (highest load -> 8%, lowest load -> 16%).

Run with:
    pip install streamlit pandas numpy ortools xlsxwriter openpyxl
    streamlit run putaway_roster_cpsat.py
"""

import io
import re
import math
from datetime import datetime, date
import pandas as pd
import numpy as np
import streamlit as st

try:
    from ortools.sat.python import cp_model
    ORTOOLS_OK = True
except Exception:
    ORTOOLS_OK = False

# ======================================================================
# CONFIG  —  every knob the maintainer needs is here, with no logic below
#            depending on numbers buried in the middle of the file.
# ======================================================================

# --- Shift vocabulary: everything non-night that means "working" collapses to Day ---
DAY_TOKENS   = {"P-M", "P-E", "P-D", "M", "E", "D"}      # working, counts as Day
NIGHT_TOKENS = {"P-N", "N"}                              # working, counts as Night
OFF_TOKENS   = {"WOD", "WO"}                             # planned week-off
LEAVE_TOKENS = {"L", "L-D", "L-N", "A", "A-D", "A-N"}    # leave / absent (breaks streaks)

# --- Hard policy (compliance + contract) ---
MAX_CONSEC_WORK_DAYS = 9       # legal cap on consecutive working days (HARD)
WO_DAYS_PER_OFF      = 7.0     # ~1 week-off per 7 active days  -> 4 for a full 31-day month
ENFORCE_SHIFT_BLOCKS = True    # ASSUMPTION: no Day<->Night switch without a WO/leave between.
                               # Matches your current practice. Set False to allow free flipping.
ALLOW_WO_FLEX        = 0       # 0 = exact monthly WO count per person; 1 = allow +/-1

# --- Target percentages ---
NIGHT_LO, NIGHT_HI = 0.47, 0.49   # night share of WORKING staff, per role per day
WO_LO,    WO_HI    = 0.08, 0.16   # week-off share of ACTIVE staff, per role per day

# --- Day-of-week and per-role week-off tweaks ---
SUNDAY_WO_BUMP = 0.03             # extra WO fraction on Sundays (added to target, still clamped to 8-16%)
FLAT_WO_ROLES  = {"RSTO Putter"}  # roles whose WOs spread evenly across the month, ignoring daily load
DAY_ONLY_ROLES = {"RSTO Putter"}  # roles worked only during the day: never assign nights, no night split

# --- Soft-constraint weights (relative priority). Higher = the solver tries harder. ---
# Read in "cost per person misplaced". Night-staff is scaled (x100) internally, so its
# raw weight looks small but is comparable to the others per person.
W_NIGHT_STAFF = 12      # ~1200 / person outside the 47-49% night band   (top priority)
W_WO_RANGE    = 1000    # per person outside the 8-16% WO band            (top priority)
W_PREF        = 200     # per night a person is outside their band        (below staffing)
W_STREAK      = 25      # per working day beyond the 6th in a 7-day run   (nudge WO by day 6-7)
W_WO_LOAD     = 4       # per person/day from the load-weighted WO target (gentle shaping)

# --- Split nights for light-night staff (avoid long back-to-back night blocks) ---
SPLIT_NIGHTS_BAND_MAX = 12  # apply to people whose night band tops out at <= this
MAX_NIGHT_RUN         = 3   # soft cap on consecutive nights for those people (the dial to tune)
W_NIGHT_RUN           = 30  # per night beyond the run cap (gentle; never changes night totals)

PREF_DEFAULT_MAX  = 20  # night ceiling for a male with no stated preference
SOLVER_TIME_LIMIT = 120 # seconds
SOLVER_WORKERS    = 8

NLO_PCT = int(round(NIGHT_LO * 100))
NHI_PCT = int(round(NIGHT_HI * 100))

META_COLS = {"Emp_ID", "Name", "NAME", "Gender", "Job Role", "Role"}


# ======================================================================
# SMALL HELPERS
# ======================================================================
def parse_dates(series, dayfirst=True):
    """Parse a date column robustly. Retries with the opposite convention if most fail."""
    out = pd.to_datetime(series, dayfirst=dayfirst, errors="coerce")
    if out.isna().mean() > 0.5:
        out = pd.to_datetime(series, dayfirst=not dayfirst, errors="coerce")
    return out


def clean_num(x):
    """'534,000' -> 534000.0 ; blanks -> NaN."""
    if pd.isna(x):
        return float("nan")
    s = str(x).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return float("nan")


def classify(val):
    """Map a raw cell to one of: 'D', 'N', 'OFF', 'LEAVE', 'NONE'."""
    if val is None:
        return "NONE"
    s = str(val).strip().upper()
    if s in ("", "NAN", "NONE", "0", "0.0"):
        return "NONE"
    if s in NIGHT_TOKENS:
        return "N"
    if s in DAY_TOKENS:
        return "D"
    if s in OFF_TOKENS:
        return "OFF"
    if s in LEAVE_TOKENS:
        return "LEAVE"
    # Fallbacks for anything unexpected
    if s.endswith("-N"):
        return "N"
    return "D"


def parse_band_value(val):
    """Recover (min_nights, max_nights) from a night-preference cell.

    Handles plain text like '7-12' or '21-26', AND the common Excel glitch where a
    value like '7-12' is silently stored as the date 12-July. In that date, the month
    and the day ARE the two numbers the user typed (month 7, day 12 -> band 7-12), so
    we rebuild from .month and .day and ignore the year/time entirely.
    """
    if val is None:
        return None
    if isinstance(val, float) and math.isnan(val):
        return None

    # Case 1: Excel turned 'M-N' into a real date -> month & day are the intended numbers.
    if isinstance(val, (pd.Timestamp, datetime, date)):
        lo, hi = sorted((int(val.month), int(val.day)))
        if 0 <= lo <= 31 and 0 <= hi <= 31:
            return lo, hi
        return None

    # Case 2: plain text or number. Keep only 0-31 values, then take the first two in order.
    nums = [int(x) for x in re.findall(r"\d+", str(val))]
    nums = [n for n in nums if 0 <= n <= 31]
    if len(nums) >= 2:
        return min(nums[0], nums[1]), max(nums[0], nums[1])
    if len(nums) == 1:
        return nums[0], nums[0]
    return None


# ======================================================================
# UI
# ======================================================================
st.set_page_config(page_title="Putaway Roster (CP-SAT)", page_icon="📅", layout="centered")
st.title("📅 Automated Putaway Roster — CP-SAT engine")
st.markdown(
    "Upload your standard **Putaway_Roster.xlsx**.  \n"
    "*Whole-month optimisation · 47-49% nights (floor + ceiling) · monthly night-preference bands · "
    "load-weighted 8-16% week-offs · hard 9-day cap.*"
)

if not ORTOOLS_OK:
    st.error("OR-Tools is not installed. Run:  `pip install ortools`  and restart the app.")

uploaded_file = st.file_uploader("Upload Input Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None and ORTOOLS_OK and st.button("🚀 Generate Roster", use_container_width=True):
    with st.spinner("Solving the whole month at once… this takes a few seconds."):
        try:
            # ==========================================================
            # 1. LOAD + CLEAN
            # ==========================================================
            df_prev    = pd.read_excel(uploaded_file, sheet_name="Previous_Month_Attendance")
            df_targets = pd.read_excel(uploaded_file, sheet_name="Daily_Targets")
            df_emp     = pd.read_excel(uploaded_file, sheet_name="Employee_Master")
            try:
                df_leaves = pd.read_excel(uploaded_file, sheet_name="Planned_Leaves")
            except Exception:
                df_leaves = pd.DataFrame(columns=["Emp_ID", "Date"])
            try:
                df_prefs = pd.read_excel(uploaded_file, sheet_name="Night_Preferences")
            except Exception:
                df_prefs = pd.DataFrame(columns=["Emp_ID"])

            # Normalise headers: strip whitespace on every column name, and accept
            # common aliases ("Job Role", etc.) as the canonical "Role".
            ROLE_ALIASES = {"job role", "jobrole", "role", "department", "dept"}
            for df in (df_emp, df_prev, df_targets, df_leaves, df_prefs):
                df.columns = [str(c).strip() for c in df.columns]
                if "Role" not in df.columns:
                    for c in list(df.columns):
                        if c.strip().lower() in ROLE_ALIASES:
                            df.rename(columns={c: "Role"}, inplace=True)
                            break
                if "Emp_ID" in df.columns:
                    df["Emp_ID"] = df["Emp_ID"].astype(str).str.strip()

            # Clear, friendly errors if a required column is still missing.
            if "Role" not in df_emp.columns:
                st.error(
                    "Employee_Master has no 'Role' column. Columns found: "
                    f"{list(df_emp.columns)}. Rename the role column to 'Role' (or 'Job Role'), "
                    "and check there is no title row sitting above the header row."
                )
                st.stop()
            if "Role" not in df_targets.columns:
                st.error(
                    "Daily_Targets has no 'Role' column. Columns found: "
                    f"{list(df_targets.columns)}. Rename the role column to 'Role'."
                )
                st.stop()

            df_emp["Role"]     = df_emp["Role"].astype(str).str.strip()
            df_targets["Role"] = df_targets["Role"].astype(str).str.strip()

            df_targets["Date"] = parse_dates(df_targets["Date"], dayfirst=True)
            df_targets = df_targets.dropna(subset=["Date"])
            df_targets["Daily_Load"] = df_targets["Daily_Load"].apply(clean_num)
            if "Zero_WO_Day" in df_targets.columns:
                df_targets["Zero_WO_Day"] = (
                    df_targets["Zero_WO_Day"].astype(str).str.strip().str.upper().isin(["TRUE", "1", "YES"])
                )
            else:
                df_targets["Zero_WO_Day"] = False

            if not df_leaves.empty and "Date" in df_leaves.columns:
                df_leaves["Date"] = parse_dates(df_leaves["Date"], dayfirst=True)

            df_emp["DOJ"] = parse_dates(df_emp["Date_of_Joining"], dayfirst=False)

            # Horizon
            days = sorted(df_targets["Date"].dt.normalize().unique())
            days = [pd.Timestamp(d) for d in days]
            num_days = len(days)
            day_index = {d: i for i, d in enumerate(days)}

            load_by_date = {
                d: df_targets.loc[df_targets["Date"].dt.normalize() == d, "Daily_Load"].dropna().iloc[0]
                if not df_targets.loc[df_targets["Date"].dt.normalize() == d, "Daily_Load"].dropna().empty
                else float("nan")
                for d in days
            }
            zero_wo = {}
            for _, r in df_targets.iterrows():
                zero_wo[(pd.Timestamp(r["Date"]).normalize(), r["Role"])] = bool(r["Zero_WO_Day"])

            loads = [load_by_date[d] for d in days if not math.isnan(load_by_date[d])]
            lmin, lmax = (min(loads), max(loads)) if loads else (0.0, 0.0)

            avg_wo_frac = (WO_LO + WO_HI) / 2.0

            def wo_target_frac(d, role=None):
                # Flat-WO roles ignore load entirely: even spread across the month.
                if role in FLAT_WO_ROLES:
                    base = avg_wo_frac
                else:
                    ld = load_by_date[d]
                    if lmax == lmin or math.isnan(ld):
                        base = avg_wo_frac
                    else:
                        f = (ld - lmin) / (lmax - lmin)    # 0 at lowest load, 1 at highest
                        base = WO_HI - (WO_HI - WO_LO) * f  # lowest load -> more offs (WO_HI)
                    # Sundays run lighter than the sheet says -> push toward more offs.
                    if d.weekday() == 6:                    # Monday=0 ... Sunday=6
                        base += SUNDAY_WO_BUMP
                return min(max(base, WO_LO), WO_HI)         # never leave the 8-16% band

            roles = sorted(df_emp["Role"].unique())

            # ==========================================================
            # 2. NIGHT PREFERENCE BANDS
            # ==========================================================
            has_minmax = {"Min_Nights", "Max_Nights"}.issubset(df_prefs.columns)
            pref_str = {}
            if "Night_Shift_Pref" in df_prefs.columns:
                d2 = df_prefs.dropna(subset=["Emp_ID"]).drop_duplicates("Emp_ID", keep="last")
                pref_str = d2.set_index("Emp_ID")["Night_Shift_Pref"].to_dict()
            pref_minmax = {}
            if has_minmax:
                d3 = df_prefs.dropna(subset=["Emp_ID"]).drop_duplicates("Emp_ID", keep="last")
                pref_minmax = d3.set_index("Emp_ID")[["Min_Nights", "Max_Nights"]].to_dict("index")

            def get_band(emp_id, gender):
                """Return (min_nights, max_nights, is_default)."""
                if gender != "Male":
                    return 0, 0, False                       # women: Day only
                if emp_id in pref_minmax:
                    row = pref_minmax[emp_id]
                    try:
                        return int(row["Min_Nights"]), int(row["Max_Nights"]), False
                    except (ValueError, TypeError):
                        pass
                if emp_id in pref_str:
                    parsed = parse_band_value(pref_str[emp_id])
                    if parsed:
                        return parsed[0], parsed[1], False
                return 0, PREF_DEFAULT_MAX, True

            # ==========================================================
            # 3. PER-EMPLOYEE STATE (carry-over from June, eligibility, entitlement)
            # ==========================================================
            prev_idx = (
                df_prev.dropna(subset=["Emp_ID"]).drop_duplicates("Emp_ID", keep="first").set_index("Emp_ID")
                if "Emp_ID" in df_prev.columns else pd.DataFrame()
            )
            prev_date_cols = [c for c in df_prev.columns if c not in META_COLS]

            leave_by_emp = {}
            if not df_leaves.empty and "Date" in df_leaves.columns:
                for _, r in df_leaves.dropna(subset=["Date"]).iterrows():
                    leave_by_emp.setdefault(r["Emp_ID"], set()).add(pd.Timestamp(r["Date"]).normalize())

            emp = {}   # emp_id -> dict of state
            for _, row in df_emp.iterrows():
                eid = row["Emp_ID"]
                gender = row["Gender"]
                doj = row["DOJ"] if pd.notna(row["DOJ"]) else None
                mn, mx, is_def = get_band(eid, gender)
                eligible = (gender == "Male") and (mx > 0) and (row["Role"] not in DAY_ONLY_ROLES)
                if not eligible:
                    mn, mx = 0, 0          # can't work nights -> no night-preference penalty

                # schedulable day indices = joined and not on leave
                sched = []
                for i, d in enumerate(days):
                    joined = (doj is None) or (doj.normalize() <= d)
                    on_leave = d in leave_by_emp.get(eid, set())
                    if joined and not on_leave:
                        sched.append(i)
                active_days = sum(1 for d in days if (doj is None) or (doj.normalize() <= d))
                entitlement = int(round(active_days / WO_DAYS_PER_OFF))
                entitlement = max(0, min(entitlement, len(sched)))

                # June carry-over: trailing working streak + last working shift
                june_work = []           # 1 if that June day was working, in chronological order
                last_class = "NONE"
                if eid in prev_idx.index and prev_date_cols:
                    prow = prev_idx.loc[eid]
                    for c in prev_date_cols:
                        june_work.append(1 if classify(prow[c]) in ("D", "N") else 0)
                    last_class = classify(prow[prev_date_cols[-1]])
                june_tail = june_work[-MAX_CONSEC_WORK_DAYS:] if june_work else []
                prev_last_working = last_class in ("D", "N")
                prev_shift = last_class if prev_last_working else None

                emp[eid] = dict(
                    role=row["Role"], gender=gender, doj=doj, eligible=eligible,
                    mn=mn, mx=mx, is_def=is_def, sched=sched, sched_set=set(sched),
                    entitlement=entitlement, june_tail=june_tail,
                    prev_last_working=prev_last_working, prev_shift=prev_shift,
                )

            # Roles that actually have night-eligible staff. Day-only roles (e.g. RSTO Putter)
            # are excluded, so the night split is neither enforced nor reported for them.
            night_roles = {r for r in roles if any(e["eligible"] and e["role"] == r for e in emp.values())}

            # ==========================================================
            # 4. BUILD THE CP-SAT MODEL
            # ==========================================================
            model = cp_model.CpModel()
            day_v, night_v, wo_v = {}, {}, {}

            for eid, e in emp.items():
                for i in e["sched"]:
                    dv = model.NewBoolVar(f"d_{eid}_{i}")
                    ov = model.NewBoolVar(f"o_{eid}_{i}")
                    if e["eligible"]:
                        nv = model.NewBoolVar(f"n_{eid}_{i}")
                        model.Add(dv + nv + ov == 1)          # exactly one state
                        night_v[(eid, i)] = nv
                    else:
                        model.Add(dv + ov == 1)               # Day or week-off only
                    day_v[(eid, i)] = dv
                    wo_v[(eid, i)] = ov

            def NV(eid, i):  # night expression (0 if not eligible)
                return night_v.get((eid, i), 0)

            def WV(eid, i):  # working expression (day or night)
                n = night_v.get((eid, i))
                return day_v[(eid, i)] + n if n is not None else day_v[(eid, i)]

            obj_night, obj_wo_range, obj_wo_load, obj_pref, obj_streak = [], [], [], [], []

            # ---- HARD: shift blocks (no Day<->Night flip without a break) ----
            if ENFORCE_SHIFT_BLOCKS:
                for eid, e in emp.items():
                    # boundary with June
                    if e["prev_last_working"] and 0 in e["sched_set"]:
                        if e["prev_shift"] == "N" and (eid, 0) in day_v:
                            model.Add(day_v[(eid, 0)] == 0)        # was night -> stay night or take WO
                        elif e["prev_shift"] == "D" and (eid, 0) in night_v:
                            model.Add(night_v[(eid, 0)] == 0)
                    # within July, between calendar-adjacent schedulable days
                    for i in range(num_days - 1):
                        if i in e["sched_set"] and (i + 1) in e["sched_set"]:
                            model.Add(NV(eid, i) + day_v[(eid, i + 1)] <= 1)
                            model.Add(day_v[(eid, i)] + NV(eid, i + 1) <= 1)

            # ---- HARD: <= 9 consecutive working days (spans the June boundary) ----
            WIN = MAX_CONSEC_WORK_DAYS + 1
            for eid, e in emp.items():
                seq = [("c", v) for v in e["june_tail"]]
                for i in range(num_days):
                    seq.append(("e", WV(eid, i)) if i in e["sched_set"] else ("c", 0))
                for s in range(0, len(seq) - WIN + 1):
                    window = seq[s:s + WIN]
                    exprs = [v for t, v in window if t == "e"]
                    if exprs:
                        const = sum(v for t, v in window if t == "c")
                        model.Add(sum(exprs) + const <= MAX_CONSEC_WORK_DAYS)

            # ---- HARD: monthly week-off count per person ----
            for eid, e in emp.items():
                wos = [wo_v[(eid, i)] for i in e["sched"]]
                ent = e["entitlement"]
                if not wos:
                    continue
                if ALLOW_WO_FLEX:
                    model.Add(sum(wos) >= max(0, ent - 1))
                    model.Add(sum(wos) <= ent + 1)
                else:
                    model.Add(sum(wos) == ent)

            # ---- Precompute who is schedulable per (day, role) ----
            sched_by_day_role = {(i, r): [] for i in range(num_days) for r in roles}
            for eid, e in emp.items():
                for i in e["sched"]:
                    sched_by_day_role[(i, e["role"])].append(eid)

            # ---- SOFT: night split 47-49% (per day, per role) ----
            for i in range(num_days):
                for r in roles:
                    members = sched_by_day_role[(i, r)]
                    if not members:
                        continue
                    # Day-only roles (no night-eligible staff that day) have no night split.
                    if not any(emp[m]["eligible"] for m in members):
                        continue
                    Wexpr = sum(WV(m, i) for m in members)        # working count (variable)
                    Nexpr = sum(NV(m, i) for m in members)        # night count (variable)
                    cap = 100 * len(members)
                    under = model.NewIntVar(0, cap, f"nu_{i}_{r}")
                    over = model.NewIntVar(0, cap, f"no_{i}_{r}")
                    model.Add(under >= NLO_PCT * Wexpr - 100 * Nexpr)   # nights below 47%
                    model.Add(over >= 100 * Nexpr - NHI_PCT * Wexpr)    # nights above 49%
                    obj_night += [under, over]

            # ---- SOFT: week-offs in 8-16% of ACTIVE, target interpolated by load ----
            for i in range(num_days):
                d = days[i]
                for r in roles:
                    members = sched_by_day_role[(i, r)]
                    A = len(members)
                    if A == 0:
                        continue
                    woexpr = sum(wo_v[(m, i)] for m in members)
                    if zero_wo.get((d, r), False):
                        model.Add(woexpr == 0)
                        continue
                    lo = math.floor(WO_LO * A)
                    hi = math.ceil(WO_HI * A)
                    u = model.NewIntVar(0, A, f"wu_{i}_{r}")
                    o = model.NewIntVar(0, A, f"wo_{i}_{r}")
                    model.Add(u >= lo - woexpr)                  # below 8%
                    model.Add(o >= woexpr - hi)                  # above 16%
                    obj_wo_range += [u, o]
                    tgt = int(round(wo_target_frac(d, r) * A))
                    tgt = min(max(tgt, lo), hi)
                    dev = model.NewIntVar(0, A, f"wd_{i}_{r}")
                    model.Add(dev >= woexpr - tgt)
                    model.Add(dev >= tgt - woexpr)
                    obj_wo_load += [dev]

            # ---- SOFT: monthly night-preference band (per person) ----
            for eid, e in emp.items():
                if not e["sched"]:
                    continue
                tot = sum(NV(eid, i) for i in e["sched"])
                L = len(e["sched"])
                u = model.NewIntVar(0, L, f"pu_{eid}")
                o = model.NewIntVar(0, L, f"po_{eid}")
                model.Add(u >= e["mn"] - tot)                    # fewer nights than wanted
                model.Add(o >= tot - e["mx"])                    # more nights than wanted
                obj_pref += [u, o]

            # ---- SOFT: discourage runs of 7+ working days (nudge WO by day 6-7) ----
            W7 = 7
            for eid, e in emp.items():
                seq = [("c", v) for v in e["june_tail"]]
                for i in range(num_days):
                    seq.append(("e", WV(eid, i)) if i in e["sched_set"] else ("c", 0))
                for s in range(0, len(seq) - W7 + 1):
                    window = seq[s:s + W7]
                    exprs = [v for t, v in window if t == "e"]
                    if exprs:
                        const = sum(v for t, v in window if t == "c")
                        sv = model.NewIntVar(0, W7, f"st_{eid}_{s}")
                        model.Add(sv >= sum(exprs) + const - 6)
                        obj_streak.append(sv)

            # ---- SOFT: split nights for light-night staff (band max <= SPLIT_NIGHTS_BAND_MAX) ----
            # Penalise any run of more than MAX_NIGHT_RUN consecutive night shifts so their
            # nights are spread across the month rather than kept in one long block. A rest day
            # or a day shift breaks the run. Heavy-night staff are not affected.
            obj_night_run = []
            WN = MAX_NIGHT_RUN + 1
            for eid, e in emp.items():
                if (not e["eligible"]) or e["mx"] == 0 or e["mx"] > SPLIT_NIGHTS_BAND_MAX:
                    continue
                sched = e["sched"]
                for s in range(0, len(sched) - WN + 1):
                    window = sched[s:s + WN]
                    if window[-1] - window[0] == WN - 1:        # calendar-consecutive days only
                        run = sum(NV(eid, i) for i in window)
                        sv = model.NewIntVar(0, WN, f"nr_{eid}_{s}")
                        model.Add(sv >= run - MAX_NIGHT_RUN)
                        obj_night_run.append(sv)

            # ---- OBJECTIVE ----
            model.Minimize(
                W_NIGHT_STAFF * sum(obj_night)
                + W_WO_RANGE * sum(obj_wo_range)
                + W_PREF * sum(obj_pref)
                + W_STREAK * sum(obj_streak)
                + W_WO_LOAD * sum(obj_wo_load)
                + W_NIGHT_RUN * sum(obj_night_run)
            )

            # ==========================================================
            # 5. SOLVE
            # ==========================================================
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT
            solver.parameters.num_search_workers = SOLVER_WORKERS
            status = solver.Solve(model)
            status_name = solver.StatusName(status)

            if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                st.error(
                    f"Solver returned {status_name}. This usually means the HARD constraints "
                    f"conflict (e.g. exact WO count + 9-day cap for someone with a long June streak). "
                    f"Try setting ALLOW_WO_FLEX = 1 or ENFORCE_SHIFT_BLOCKS = False and re-run."
                )
                st.stop()

            # ==========================================================
            # 6. EXTRACT SCHEDULE
            # ==========================================================
            def cell(eid, i):
                e = emp[eid]
                d = days[i]
                if (e["doj"] is not None) and (e["doj"].normalize() > d):
                    return "Not Joined"
                if d in leave_by_emp.get(eid, set()):
                    return "L"
                if i not in e["sched_set"]:
                    return ""
                if solver.Value(wo_v[(eid, i)]) == 1:
                    return "WOD"
                if (eid, i) in night_v and solver.Value(night_v[(eid, i)]) == 1:
                    return "N"
                return "D"

            day_labels = [d.strftime("%d-%b") for d in days]
            rows = []
            for eid, e in emp.items():
                rec = {"Emp_ID": eid, "Role": e["role"]}
                nights = wos = 0
                for i in range(num_days):
                    c = cell(eid, i)
                    rec[day_labels[i]] = c
                    nights += (c == "N")
                    wos += (c == "WOD")
                rec["Total_WOs"] = wos
                rec["Total_Nights"] = nights
                rec["Pref_Band"] = f"{e['mn']}-{e['mx']}" if e["eligible"] else "Day only"
                rows.append(rec)
            roster_df = pd.DataFrame(rows).sort_values(["Role", "Emp_ID"]).reset_index(drop=True)

            # ---- Daily summary (per day, per role) ----
            summ = []
            for i in range(num_days):
                d = days[i]
                for r in roles:
                    members = sched_by_day_role[(i, r)]
                    if not members:
                        continue
                    work = sum(1 for m in members if cell(m, i) in ("D", "N"))
                    nig = sum(1 for m in members if cell(m, i) == "N")
                    woc = sum(1 for m in members if cell(m, i) == "WOD")
                    active = len(members)
                    summ.append({
                        "Date": d.strftime("%d-%b"), "Role": r, "Active": active,
                        "Working": work, "Day": work - nig, "Night": nig,
                        "Night_%": round(100 * nig / work, 1) if work else 0,
                        "Night_target": f"{NLO_PCT}-{NHI_PCT}%" if r in night_roles else "Day only",
                        "WeekOffs": woc,
                        "WO_%": round(100 * woc / active, 1) if active else 0,
                        "WO_target_%": round(100 * wo_target_frac(d, r), 1),
                    })
            summary_df = pd.DataFrame(summ)

            # ---- Plain-language constraint report (what the solver had to bend) ----
            report = []
            for eid, e in emp.items():
                if e["gender"] != "Male":
                    continue
                tot = sum(1 for i in range(num_days) if cell(eid, i) == "N")
                if tot < e["mn"]:
                    report.append({"Type": "Night preference", "Who": eid,
                                   "Detail": f"got {tot} nights, below preferred minimum {e['mn']} "
                                             f"(short {e['mn'] - tot}) — likely needed to staff the floor"})
                elif tot > e["mx"]:
                    report.append({"Type": "Night preference", "Who": eid,
                                   "Detail": f"got {tot} nights, above preferred maximum {e['mx']} "
                                             f"(over {tot - e['mx']}) — likely needed to staff the floor"})
            for _, s in summary_df.iterrows():
                if s["Role"] in night_roles and s["Working"] and not (NLO_PCT <= s["Night_%"] <= NHI_PCT):
                    report.append({"Type": "Night split", "Who": f"{s['Date']} / {s['Role']}",
                                   "Detail": f"nights at {s['Night_%']}% of {int(s['Working'])} working "
                                             f"(target {NLO_PCT}-{NHI_PCT}%)"})
                wo_lo_n = math.floor(WO_LO * s["Active"])
                wo_hi_n = math.ceil(WO_HI * s["Active"])
                if not (wo_lo_n <= s["WeekOffs"] <= wo_hi_n):
                    report.append({"Type": "Week-off range", "Who": f"{s['Date']} / {s['Role']}",
                                   "Detail": f"{int(s['WeekOffs'])} offs vs allowed {wo_lo_n}-{wo_hi_n}"})
            report_df = pd.DataFrame(report) if report else pd.DataFrame(
                [{"Type": "—", "Who": "—", "Detail": "All soft targets met."}]
            )

            # ==========================================================
            # 7. WRITE WORKBOOK
            # ==========================================================
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                roster_df.to_excel(writer, index=False, sheet_name="Generated_Roster")
                summary_df.to_excel(writer, index=False, sheet_name="Daily_Summary")
                report_df.to_excel(writer, index=False, sheet_name="Constraint_Report")

            month_name = days[0].strftime("%B_%Y")
            st.success(f"✅ {month_name.replace('_', ' ')} roster generated — solver status: {status_name}.")

            # Headline diagnostics in the app
            c1, c2, c3 = st.columns(3)
            outside_night = sum(
                1 for _, s in summary_df.iterrows()
                if s["Role"] in night_roles and s["Working"] and not (NLO_PCT <= s["Night_%"] <= NHI_PCT)
            )
            pref_viol = sum(
                1 for eid, e in emp.items() if e["gender"] == "Male"
                and not (e["mn"] <= sum(1 for i in range(num_days) if cell(eid, i) == "N") <= e["mx"])
            )
            c1.metric("Day/role cells off night band", outside_night)
            c2.metric("People outside night band", pref_viol)
            c3.metric("Objective penalty", int(solver.ObjectiveValue()))
            st.caption("Lower is better. Zeros mean every soft target was met.")
            st.dataframe(summary_df, use_container_width=True, height=320)

            st.download_button(
                "⬇️ Download Roster + Diagnostics (.xlsx)",
                data=buf.getvalue(),
                file_name=f"{month_name}_Putaway_Roster.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as exc:
            st.error(f"⚠️ An error occurred: {exc}")
            st.exception(exc)
