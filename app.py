"""
Roster Generator — CP-SAT operational tool (multi-department, Excel-configurable)

Everything that used to be hard-coded now lives in optional Excel sheets, so any
department configures and runs it without touching this file. With none of the new
sheets present, behaviour is identical to the previous version — your current
workbook still runs unchanged.

INPUT SHEETS
  Required : Employee_Master, Daily_Targets
  Usual    : Previous_Month_Attendance, Night_Preferences, Planned_Leaves
  Optional : Config, Roles_Config, Shift_Map, Current_Roster, History

WHAT EACH OPTIONAL SHEET DOES
  Config        : key/value overrides for every number/weight (see DEFAULTS below).
  Roles_Config  : per-role flags — Day_Only (Y/N), Flat_WO (Y/N), Night_Lo/Night_Hi.
  Shift_Map     : Token -> Type (Day/Night/Off/Leave) so each dept maps its own codes.
  Current_Roster: a previously generated roster, used to FREEZE part of the plan:
                    - FREEZE_THROUGH_DATE blank -> lock the whole existing roster; only
                      employees NOT in it (new joiners) are built, from their join date.
                    - FREEZE_THROUGH_DATE set   -> lock days on/before that date for
                      everyone, re-optimise the rest (mid-month re-plan).
  History       : Emp_ID, Cumulative_Nights, Cumulative_WeekendOffs from prior months,
                  for cross-month fairness (enable with W_CUM_NIGHT_FAIR > 0).

Run with:
    pip install streamlit pandas numpy ortools xlsxwriter openpyxl
    streamlit run app.py
"""

import io
import re
import math
import warnings as pywarnings
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
# DEFAULTS — overridable from the optional "Config" sheet (key/value).
# ======================================================================
DEFAULTS = {
    "MAX_CONSEC_WORK_DAYS": 9,       # hard cap on consecutive working days
    "WO_PER_MONTH":         4,       # week-offs for someone present the whole period (fixed, not scaled by period length)
    "ENFORCE_SHIFT_BLOCKS": True,    # no Day<->Night switch without a rest day between
    "ALLOW_WO_FLEX":        0,       # 0 = exact monthly WO count; 1 = allow +/-1
    "NIGHT_LO":             0.47,    # night share of WORKING staff, per role per day
    "NIGHT_HI":             0.49,
    "WO_LO":                0.08,    # week-off share of ACTIVE staff, per role per day
    "WO_HI":                0.16,
    "SUNDAY_WO_BUMP":       0.03,    # extra WO fraction on Sundays (still clamped to band)
    "PREF_DEFAULT_MAX":     20,      # night ceiling for a male with no stated preference
    "MAX_NIGHT_RUN":        6,       # longest night stretch before it's broken by a day stretch
    "W_NIGHT_STAFF":        12,      # weights (priority). Higher = solver tries harder.
    "W_WO_RANGE":           1000,
    "W_PREF":               200,
    "W_STREAK":             25,
    "W_WO_LOAD":            4,
    "W_NIGHT_RUN":          30,
    "W_CUM_NIGHT_FAIR":     0,       # cross-month night fairness (needs History; 0 = off)
    "W_STABILITY":          0,       # month-to-month day/night stability (0 = off)
    "SOLVER_TIME_LIMIT":    120,
    "SOLVER_WORKERS":       8,
    "SOLVER_SEED":          1,       # fixed seed for reproducible runs
}
# FREEZE_THROUGH_DATE handled separately (it's a date, default None).

DEFAULT_SHIFT_MAP = {
    "P-M": "Day", "P-E": "Day", "P-D": "Day", "M": "Day", "E": "Day", "D": "Day",
    "P-N": "Night", "N": "Night",
    "WOD": "Off", "WO": "Off",
    "L": "Leave", "L-D": "Leave", "L-N": "Leave", "A": "Leave", "A-D": "Leave", "A-N": "Leave",
}
DEFAULT_DAY_ONLY = {"RSTO Putter"}    # used only if no Roles_Config sheet is supplied
DEFAULT_FLAT_WO  = {"RSTO Putter"}

META_COLS = {"Emp_ID", "Name", "NAME", "Gender", "Job Role", "Role"}
ROLE_ALIASES = {"job role", "jobrole", "role", "department", "dept"}


# ======================================================================
# SMALL HELPERS
# ======================================================================
def parse_dates(series, dayfirst=True):
    """Robust to mixed date formats in one column (e.g. '01-07-2026' and '7/22/2026').
    Fast path assumes one format; if anything fails, infer per element, then fall back
    to element-wise parsing that tries both day-first and month-first."""
    with pywarnings.catch_warnings():
        pywarnings.simplefilter("ignore")
        out = pd.to_datetime(series, dayfirst=dayfirst, errors="coerce")
        if out.isna().any():
            try:
                out = out.fillna(pd.to_datetime(series, format="mixed", dayfirst=dayfirst, errors="coerce"))
            except Exception:
                pass
            if out.isna().any():
                def _one(x):
                    if pd.isna(x):
                        return pd.NaT
                    for df in (dayfirst, not dayfirst):
                        try:
                            return pd.Timestamp(pd.to_datetime(str(x).strip(), dayfirst=df))
                        except Exception:
                            continue
                    return pd.NaT
                out = out.fillna(series.map(_one))
    return out


def clean_num(x):
    if pd.isna(x):
        return float("nan")
    s = str(x).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return float("nan")


def coerce_to(default, raw):
    """Coerce a Config value to the type of its default. A blank value keeps the default,
    so an empty cell never silently turns a setting off."""
    if raw is None or (isinstance(raw, float) and math.isnan(raw)):
        return default
    s = str(raw).strip()
    if s == "":                                        # blank cell -> keep the default
        return default
    if isinstance(default, bool):                      # bool must be checked before int
        return s.lower() in ("true", "1", "yes", "y", "t")
    if isinstance(default, int):
        try:
            return int(float(s.replace(",", "")))
        except ValueError:
            return default
    if isinstance(default, float):
        try:
            return float(s.replace(",", ""))
        except ValueError:
            return default
    return s


def parse_band_value(val):
    """Recover (min, max) night band. Handles 'A-B' text and the Excel-date glitch
    where e.g. '7-12' is stored as 12-July (month & day are the two numbers typed)."""
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, (pd.Timestamp, datetime, date)):
        lo, hi = sorted((int(val.month), int(val.day)))
        return (lo, hi) if 0 <= lo <= 31 and 0 <= hi <= 31 else None
    nums = [int(x) for x in re.findall(r"\d+", str(val)) if 0 <= int(x) <= 31]
    if len(nums) >= 2:
        return min(nums[0], nums[1]), max(nums[0], nums[1])
    if len(nums) == 1:
        return nums[0], nums[0]
    return None


def normalize_headers(df):
    df.columns = [str(c).strip() for c in df.columns]
    if "Role" not in df.columns:
        for c in list(df.columns):
            if c.strip().lower() in ROLE_ALIASES:
                df.rename(columns={c: "Role"}, inplace=True)
                break
    if "Emp_ID" in df.columns:
        df["Emp_ID"] = df["Emp_ID"].astype(str).str.strip()
    return df


def read_sheet(xls, name, required_cols=None):
    try:
        df = pd.read_excel(xls, sheet_name=name)
        return normalize_headers(df)
    except Exception:
        cols = list(required_cols) if required_cols else []
        return pd.DataFrame(columns=cols)


# ======================================================================
# CONFIG / ROLES / SHIFT-MAP loaders
# ======================================================================
def load_config(xls):
    cfg = dict(DEFAULTS)
    freeze = None
    df = read_sheet(xls, "Config")
    if not df.empty:
        cols = {c.lower(): c for c in df.columns}
        pcol = cols.get("parameter") or cols.get("key") or df.columns[0]
        vcol = cols.get("value") or (df.columns[1] if len(df.columns) > 1 else None)
        if vcol is not None:
            for _, r in df.iterrows():
                k = str(r[pcol]).strip()
                if k in cfg:
                    cfg[k] = coerce_to(cfg[k], r[vcol])
                elif k.upper() == "FREEZE_THROUGH_DATE":
                    d = parse_dates(pd.Series([r[vcol]]), dayfirst=True).iloc[0]
                    freeze = None if pd.isna(d) else pd.Timestamp(d).normalize()
    cfg["FREEZE_THROUGH_DATE"] = freeze
    return cfg


def load_roles_config(xls):
    df = read_sheet(xls, "Roles_Config")
    if df.empty or "Role" not in df.columns:
        return DEFAULT_DAY_ONLY.copy(), DEFAULT_FLAT_WO.copy(), {}
    cols = {c.lower(): c for c in df.columns}
    day_only, flat_wo, band = set(), set(), {}

    def yes(v):
        return str(v).strip().lower() in ("y", "yes", "true", "1", "t")

    for _, r in df.iterrows():
        role = str(r["Role"]).strip()
        if not role or role.lower() == "nan":
            continue
        if "day_only" in cols and yes(r[cols["day_only"]]):
            day_only.add(role)
        if "flat_wo" in cols and yes(r[cols["flat_wo"]]):
            flat_wo.add(role)
        lo = r[cols["night_lo"]] if "night_lo" in cols else None
        hi = r[cols["night_hi"]] if "night_hi" in cols else None
        if pd.notna(lo) and pd.notna(hi):
            try:
                band[role] = (float(lo), float(hi))
            except ValueError:
                pass
    return day_only, flat_wo, band


def load_shift_map(xls):
    df = read_sheet(xls, "Shift_Map")
    if df.empty:
        return dict(DEFAULT_SHIFT_MAP)
    cols = {c.lower(): c for c in df.columns}
    tcol = cols.get("token") or df.columns[0]
    ycol = cols.get("type") or (df.columns[1] if len(df.columns) > 1 else None)
    if ycol is None:
        return dict(DEFAULT_SHIFT_MAP)
    m = {}
    for _, r in df.iterrows():
        tok = str(r[tcol]).strip().upper()
        typ = str(r[ycol]).strip().capitalize()
        if tok and typ in ("Day", "Night", "Off", "Leave"):
            m[tok] = typ
    return m or dict(DEFAULT_SHIFT_MAP)


def make_classifier(shift_map):
    upmap = {k.upper(): v for k, v in shift_map.items()}

    def classify(val):
        if val is None:
            return "NONE"
        s = str(val).strip().upper()
        if s in ("", "NAN", "NONE", "0", "0.0"):
            return "NONE"
        if s in upmap:
            return {"Day": "D", "Night": "N", "Off": "OFF", "Leave": "LEAVE"}[upmap[s]]
        if s.endswith("-N"):
            return "N"
        return "D"
    return classify


# ======================================================================
# VALIDATION (Tier 1) — runs before the solver
# ======================================================================
def validate(df_emp, df_targets, df_prev, df_prefs, df_leaves, days, emp_preview):
    errors, warnings = [], []

    for col in ("Emp_ID", "Role", "Gender", "Date_of_Joining"):
        if col not in df_emp.columns:
            errors.append(f"Employee_Master is missing required column '{col}'.")
    for col in ("Date", "Role", "Daily_Load"):
        if col not in df_targets.columns:
            errors.append(f"Daily_Targets is missing required column '{col}'.")
    if errors:
        return errors, warnings  # can't go further without these

    if not days:
        errors.append("Daily_Targets has no readable dates.")
        return errors, warnings

    # contiguous calendar
    expected = pd.date_range(days[0], days[-1])
    missing = [d.strftime("%d-%b") for d in expected if pd.Timestamp(d) not in set(days)]
    if missing:
        errors.append(f"Daily_Targets dates are not contiguous. Missing day(s): {', '.join(missing[:8])}"
                      + (" ..." if len(missing) > 8 else ""))

    # duplicate employees
    dups = df_emp["Emp_ID"][df_emp["Emp_ID"].duplicated()].unique().tolist()
    if dups:
        warnings.append(f"Duplicate Emp_IDs in Employee_Master: {', '.join(map(str, dups[:8]))}.")

    # gender values
    bad_g = sorted(set(df_emp["Gender"].dropna().astype(str)) - {"Male", "Female"})
    if bad_g:
        warnings.append(f"Unexpected Gender values (expected Male/Female): {', '.join(bad_g[:8])}.")

    master_ids = set(df_emp["Emp_ID"])
    if "Emp_ID" in df_prefs.columns:
        miss = sorted(set(df_prefs["Emp_ID"].dropna()) - master_ids)
        if miss:
            warnings.append(f"{len(miss)} Emp_ID(s) in Night_Preferences are not in Employee_Master "
                            f"(ignored): {', '.join(map(str, miss[:6]))}.")
    if "Emp_ID" in df_leaves.columns and not df_leaves.empty:
        miss = sorted(set(df_leaves["Emp_ID"].dropna()) - master_ids)
        if miss:
            warnings.append(f"{len(miss)} Emp_ID(s) in Planned_Leaves are not in Employee_Master "
                            f"(ignored): {', '.join(map(str, miss[:6]))}.")

    # role coverage
    roles_master = set(df_emp["Role"].dropna().astype(str).str.strip())
    roles_targets = set(df_targets["Role"].dropna().astype(str).str.strip())
    for r in sorted(roles_master - roles_targets):
        warnings.append(f"Role '{r}' has staff but no Daily_Targets rows — it will be rostered "
                        f"but week-offs won't track load.")
    for r in sorted(roles_targets - roles_master):
        warnings.append(f"Role '{r}' has Daily_Targets but no staff in Employee_Master.")

    # loads numeric/positive
    bad_loads = df_targets["Daily_Load"].apply(clean_num)
    if bad_loads.isna().any() or (bad_loads.dropna() <= 0).any():
        warnings.append("Some Daily_Load values are missing, non-numeric, or <= 0 — check the load column.")

    # night bands recovered from dates / inverted
    date_recovered, inverted = [], []
    for eid, info in emp_preview.items():
        if info.get("from_date"):
            date_recovered.append(eid)
        if info["mn"] > info["mx"]:
            inverted.append(eid)
    if date_recovered:
        warnings.append(f"{len(date_recovered)} night preference(s) were recovered from an Excel date "
                        f"(verify): {', '.join(date_recovered[:6])}. Tip: use Min_Nights/Max_Nights columns.")
    if inverted:
        warnings.append(f"Night band min > max for: {', '.join(inverted[:6])}.")

    # per-role night feasibility pre-scan
    active_avg = {}
    for r in roles_master:
        n_role = sum(1 for v in emp_preview.values() if v["role"] == r)
        n_elig = sum(1 for v in emp_preview.values() if v["role"] == r and v["eligible"])
        if n_role == 0:
            continue
        need_night = math.floor(0.47 * n_role * 0.88)  # rough: 47% of ~88% working
        if n_elig < need_night:
            warnings.append(f"Role '{r}': only {n_elig} night-eligible of {n_role}; the ~47-49% night "
                            f"split may be unreachable on busy days (needs ≈{need_night}).")
    return errors, warnings


# ======================================================================
# UI
# ======================================================================
st.set_page_config(page_title="Roster Generator (CP-SAT)", page_icon="📅", layout="centered")
st.title("📅 Roster Generator — operational tool")
st.markdown("Upload your department workbook. Optional **Config / Roles_Config / Shift_Map / "
            "Current_Roster / History** sheets unlock configuration, freezing/re-planning, and "
            "cross-month fairness. Without them, defaults apply.")

if not ORTOOLS_OK:
    st.error("OR-Tools is not installed. Run:  `pip install ortools`  and restart the app.")

uploaded_file = st.file_uploader("Upload Input Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None and ORTOOLS_OK and st.button("🚀 Generate Roster", use_container_width=True):
    with st.spinner("Validating, then solving the month…"):
        try:
            cfg = load_config(uploaded_file)
            day_only_roles, flat_wo_roles, role_band = load_roles_config(uploaded_file)
            shift_map = load_shift_map(uploaded_file)
            classify = make_classifier(shift_map)

            # ---------- load data ----------
            df_emp     = read_sheet(uploaded_file, "Employee_Master")
            df_targets = read_sheet(uploaded_file, "Daily_Targets")
            df_prev    = read_sheet(uploaded_file, "Previous_Month_Attendance")
            df_leaves  = read_sheet(uploaded_file, "Planned_Leaves", ["Emp_ID", "Date"])
            df_prefs   = read_sheet(uploaded_file, "Night_Preferences", ["Emp_ID"])
            df_current = read_sheet(uploaded_file, "Current_Roster")
            df_hist    = read_sheet(uploaded_file, "History")

            need = [c for c in ("Emp_ID", "Role", "Gender", "Date_of_Joining") if c not in df_emp.columns]
            need += [c for c in ("Date", "Role", "Daily_Load") if c not in df_targets.columns]
            if need:
                st.error(f"Missing required column(s): {need}. Check Employee_Master and Daily_Targets.")
                st.stop()

            df_emp["Role"] = df_emp["Role"].astype(str).str.strip()
            df_targets["Role"] = df_targets["Role"].astype(str).str.strip()
            _raw_dates = df_targets["Date"].astype(str)
            date_fmt_warn = None
            if _raw_dates.str.contains("/").any() and _raw_dates.str.contains("-").any():
                date_fmt_warn = ("Daily_Targets 'Date' column mixes '/' and '-' formats. It was parsed "
                                 "safely, but please standardize to one day-first format "
                                 "(e.g. 01-07-2026) so the roster period is never misread.")
            df_targets["Date"] = parse_dates(df_targets["Date"], dayfirst=True)
            df_targets = df_targets.dropna(subset=["Date"])
            df_targets["Daily_Load"] = df_targets["Daily_Load"].apply(clean_num)
            df_targets["Zero_WO_Day"] = (
                df_targets["Zero_WO_Day"].astype(str).str.strip().str.upper().isin(["TRUE", "1", "YES"])
                if "Zero_WO_Day" in df_targets.columns else False
            )
            if not df_leaves.empty and "Date" in df_leaves.columns:
                df_leaves["Date"] = parse_dates(df_leaves["Date"], dayfirst=True)
            df_emp["DOJ"] = parse_dates(df_emp["Date_of_Joining"], dayfirst=False)

            # ---------- horizon ----------
            days = [pd.Timestamp(d) for d in sorted(df_targets["Date"].dt.normalize().unique())]
            num_days = len(days)
            day_labels = [d.strftime("%d-%b") for d in days]
            label_to_idx = {lab: i for i, lab in enumerate(day_labels)}

            load_by_date, zero_wo = {}, {}
            for d in days:
                sub = df_targets.loc[df_targets["Date"].dt.normalize() == d, "Daily_Load"].dropna()
                load_by_date[d] = sub.iloc[0] if not sub.empty else float("nan")
            for _, r in df_targets.iterrows():
                zero_wo[(pd.Timestamp(r["Date"]).normalize(), str(r["Role"]).strip())] = bool(r["Zero_WO_Day"])
            loads = [v for v in load_by_date.values() if not math.isnan(v)]
            lmin, lmax = (min(loads), max(loads)) if loads else (0.0, 0.0)

            def wo_target_frac(d, role):
                if role in flat_wo_roles:
                    base = (cfg["WO_LO"] + cfg["WO_HI"]) / 2.0
                else:
                    ld = load_by_date[d]
                    if lmax == lmin or math.isnan(ld):
                        base = (cfg["WO_LO"] + cfg["WO_HI"]) / 2.0
                    else:
                        f = (ld - lmin) / (lmax - lmin)
                        base = cfg["WO_HI"] - (cfg["WO_HI"] - cfg["WO_LO"]) * f
                    if d.weekday() == 6:
                        base += cfg["SUNDAY_WO_BUMP"]
                return min(max(base, cfg["WO_LO"]), cfg["WO_HI"])

            def night_band(role):
                lo, hi = role_band.get(role, (cfg["NIGHT_LO"], cfg["NIGHT_HI"]))
                return int(round(lo * 100)), int(round(hi * 100))

            roles = sorted(df_emp["Role"].unique())

            # ---------- preferences ----------
            pref_str, pref_minmax = {}, {}
            if "Night_Shift_Pref" in df_prefs.columns:
                d2 = df_prefs.dropna(subset=["Emp_ID"]).drop_duplicates("Emp_ID", keep="last")
                pref_str = d2.set_index("Emp_ID")["Night_Shift_Pref"].to_dict()
            if {"Min_Nights", "Max_Nights"}.issubset(df_prefs.columns):
                d3 = df_prefs.dropna(subset=["Emp_ID"]).drop_duplicates("Emp_ID", keep="last")
                pref_minmax = d3.set_index("Emp_ID")[["Min_Nights", "Max_Nights"]].to_dict("index")

            def get_band(emp_id, gender):
                """(min, max, is_default, from_date)."""
                if gender != "Male":
                    return 0, 0, False, False
                if emp_id in pref_minmax:
                    row = pref_minmax[emp_id]
                    try:
                        return int(row["Min_Nights"]), int(row["Max_Nights"]), False, False
                    except (ValueError, TypeError):
                        pass
                if emp_id in pref_str:
                    raw = pref_str[emp_id]
                    parsed = parse_band_value(raw)
                    if parsed:
                        from_date = isinstance(raw, (pd.Timestamp, datetime, date))
                        return parsed[0], parsed[1], False, from_date
                return 0, cfg["PREF_DEFAULT_MAX"], True, False

            # ---------- previous-month carry-over ----------
            prev_idx = (df_prev.dropna(subset=["Emp_ID"]).drop_duplicates("Emp_ID", keep="first")
                        .set_index("Emp_ID") if "Emp_ID" in df_prev.columns else pd.DataFrame())
            prev_date_cols = [c for c in df_prev.columns if c not in META_COLS]

            leave_by_emp = {}
            if not df_leaves.empty and "Date" in df_leaves.columns:
                for _, r in df_leaves.dropna(subset=["Date"]).iterrows():
                    leave_by_emp.setdefault(r["Emp_ID"], set()).add(pd.Timestamp(r["Date"]).normalize())

            # ---------- current roster (for freeze / new-joiner) ----------
            current = {}
            if not df_current.empty and "Emp_ID" in df_current.columns:
                cur_date_cols = [c for c in df_current.columns if c in label_to_idx]
                for _, r in df_current.iterrows():
                    eid = r["Emp_ID"]
                    for c in cur_date_cols:
                        current.setdefault(eid, {})[label_to_idx[c]] = str(r[c]).strip().upper()
            freeze_date = cfg["FREEZE_THROUGH_DATE"]

            def lock_for(eid, i):
                """Return 'D'/'N'/'WOD' if cell (eid, day i) must be frozen, else None."""
                if eid not in current:
                    return None
                val = current[eid].get(i)
                if val not in ("D", "N", "WOD"):
                    return None
                if freeze_date is None:
                    return val                      # lock whole existing roster (new-joiner mode)
                return val if days[i] <= freeze_date else None

            # ---------- per-employee state ----------
            emp = {}
            for _, row in df_emp.iterrows():
                eid, gender = row["Emp_ID"], row["Gender"]
                doj = row["DOJ"] if pd.notna(row["DOJ"]) else None
                mn, mx, is_def, from_date = get_band(eid, gender)
                eligible = (gender == "Male") and (mx > 0) and (row["Role"] not in day_only_roles)
                if not eligible:
                    mn, mx = 0, 0

                sched = [i for i, d in enumerate(days)
                         if ((doj is None) or (doj.normalize() <= d)) and d not in leave_by_emp.get(eid, set())]
                active_days = sum(1 for d in days if (doj is None) or (doj.normalize() <= d))
                # Week-offs = the department's monthly figure, prorated by how much of the period
                # the person is present. A full-period employee gets exactly WO_PER_MONTH, whether
                # the period is 30, 31, or 35 days — period length never silently changes it.
                entitlement = max(0, min(int(round(cfg["WO_PER_MONTH"] * active_days / num_days)), len(sched)))

                june_work, prev_nights = [], 0
                last_class = "NONE"
                if eid in prev_idx.index and prev_date_cols:
                    prow = prev_idx.loc[eid]
                    for c in prev_date_cols:
                        cl = classify(prow[c])
                        june_work.append(1 if cl in ("D", "N") else 0)
                        prev_nights += (cl == "N")
                    last_class = classify(prow[prev_date_cols[-1]])
                june_tail = june_work[-cfg["MAX_CONSEC_WORK_DAYS"]:] if june_work else []

                emp[eid] = dict(
                    role=row["Role"], gender=gender, doj=doj, eligible=eligible,
                    mn=mn, mx=mx, is_def=is_def, from_date=from_date,
                    sched=sched, sched_set=set(sched), entitlement=entitlement,
                    june_tail=june_tail, prev_nights=prev_nights, prev_work=sum(june_work),
                    prev_last_working=last_class in ("D", "N"),
                    prev_shift=last_class if last_class in ("D", "N") else None,
                )

            night_roles = {r for r in roles if any(e["eligible"] and e["role"] == r for e in emp.values())}

            # ---------- validation ----------
            errors, warnings = validate(df_emp, df_targets, df_prev, df_prefs, df_leaves, days, emp)
            if date_fmt_warn:
                warnings.append(date_fmt_warn)
            if errors:
                st.error("Cannot generate — fix these first:")
                for e in errors:
                    st.write(f"• {e}")
                st.stop()
            if warnings:
                with st.expander(f"⚠️ {len(warnings)} warning(s) — review, but generation will proceed"):
                    for w in warnings:
                        st.write(f"• {w}")

            # ==========================================================
            # BUILD MODEL
            # ==========================================================
            model = cp_model.CpModel()
            day_v, night_v, wo_v = {}, {}, {}
            for eid, e in emp.items():
                for i in e["sched"]:
                    dv = model.NewBoolVar(f"d_{eid}_{i}")
                    ov = model.NewBoolVar(f"o_{eid}_{i}")
                    if e["eligible"]:
                        nv = model.NewBoolVar(f"n_{eid}_{i}")
                        model.Add(dv + nv + ov == 1)
                        night_v[(eid, i)] = nv
                    else:
                        model.Add(dv + ov == 1)
                    day_v[(eid, i)] = dv
                    wo_v[(eid, i)] = ov
                    lk = lock_for(eid, i)               # freeze cells from a Current_Roster
                    if lk == "D":
                        model.Add(dv == 1)
                    elif lk == "WOD":
                        model.Add(ov == 1)
                    elif lk == "N" and e["eligible"]:
                        model.Add(nv == 1)

            def NV(eid, i):
                return night_v.get((eid, i), 0)

            def WV(eid, i):
                n = night_v.get((eid, i))
                return day_v[(eid, i)] + n if n is not None else day_v[(eid, i)]

            obj_night, obj_wo_range, obj_wo_load, obj_pref, obj_streak, obj_night_run = [], [], [], [], [], []
            obj_cum, obj_stab = [], []

            # HARD: shift blocks
            if cfg["ENFORCE_SHIFT_BLOCKS"]:
                for eid, e in emp.items():
                    if e["prev_last_working"] and 0 in e["sched_set"]:
                        if e["prev_shift"] == "N" and (eid, 0) in day_v:
                            model.Add(day_v[(eid, 0)] == 0)
                        elif e["prev_shift"] == "D" and (eid, 0) in night_v:
                            model.Add(night_v[(eid, 0)] == 0)
                    for i in range(num_days - 1):
                        if i in e["sched_set"] and (i + 1) in e["sched_set"]:
                            model.Add(NV(eid, i) + day_v[(eid, i + 1)] <= 1)
                            model.Add(day_v[(eid, i)] + NV(eid, i + 1) <= 1)

            # HARD: <= max consecutive working days (spans previous month)
            WIN = cfg["MAX_CONSEC_WORK_DAYS"] + 1
            for eid, e in emp.items():
                seq = [("c", v) for v in e["june_tail"]]
                for i in range(num_days):
                    seq.append(("e", WV(eid, i)) if i in e["sched_set"] else ("c", 0))
                for s in range(0, len(seq) - WIN + 1):
                    w = seq[s:s + WIN]
                    ex = [v for t, v in w if t == "e"]
                    if ex:
                        model.Add(sum(ex) + sum(v for t, v in w if t == "c") <= cfg["MAX_CONSEC_WORK_DAYS"])

            # HARD: monthly week-off count
            for eid, e in emp.items():
                wos = [wo_v[(eid, i)] for i in e["sched"]]
                if not wos:
                    continue
                if cfg["ALLOW_WO_FLEX"]:
                    model.Add(sum(wos) >= max(0, e["entitlement"] - 1))
                    model.Add(sum(wos) <= e["entitlement"] + 1)
                else:
                    model.Add(sum(wos) == e["entitlement"])

            sched_by_day_role = {(i, r): [] for i in range(num_days) for r in roles}
            for eid, e in emp.items():
                for i in e["sched"]:
                    sched_by_day_role[(i, e["role"])].append(eid)

            # SOFT: night split (per role band)
            for i in range(num_days):
                for r in roles:
                    members = sched_by_day_role[(i, r)]
                    if not members or not any(emp[m]["eligible"] for m in members):
                        continue
                    nlo, nhi = night_band(r)
                    Wexpr = sum(WV(m, i) for m in members)
                    Nexpr = sum(NV(m, i) for m in members)
                    cap = 100 * len(members)
                    under = model.NewIntVar(0, cap, f"nu_{i}_{r}")
                    over = model.NewIntVar(0, cap, f"no_{i}_{r}")
                    model.Add(under >= nlo * Wexpr - 100 * Nexpr)
                    model.Add(over >= 100 * Nexpr - nhi * Wexpr)
                    obj_night += [under, over]

            # SOFT: week-offs in band + load-weighted target
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
                    lo, hi = math.floor(cfg["WO_LO"] * A), math.ceil(cfg["WO_HI"] * A)
                    u = model.NewIntVar(0, A, f"wu_{i}_{r}")
                    o = model.NewIntVar(0, A, f"wo_{i}_{r}")
                    model.Add(u >= lo - woexpr)
                    model.Add(o >= woexpr - hi)
                    obj_wo_range += [u, o]
                    tgt = min(max(int(round(wo_target_frac(d, r) * A)), lo), hi)
                    dev = model.NewIntVar(0, A, f"wd_{i}_{r}")
                    model.Add(dev >= woexpr - tgt)
                    model.Add(dev >= tgt - woexpr)
                    obj_wo_load.append(dev)

            # SOFT: monthly night-preference band
            for eid, e in emp.items():
                if not e["sched"]:
                    continue
                tot = sum(NV(eid, i) for i in e["sched"])
                L = len(e["sched"])
                u = model.NewIntVar(0, L, f"pu_{eid}")
                o = model.NewIntVar(0, L, f"po_{eid}")
                model.Add(u >= e["mn"] - tot)
                model.Add(o >= tot - e["mx"])
                obj_pref += [u, o]

            # SOFT: discourage 7+ working-day runs
            for eid, e in emp.items():
                seq = [("c", v) for v in e["june_tail"]]
                for i in range(num_days):
                    seq.append(("e", WV(eid, i)) if i in e["sched_set"] else ("c", 0))
                for s in range(0, len(seq) - 7 + 1):
                    w = seq[s:s + 7]
                    ex = [v for t, v in w if t == "e"]
                    if ex:
                        sv = model.NewIntVar(0, 7, f"st_{eid}_{s}")
                        model.Add(sv >= sum(ex) + sum(v for t, v in w if t == "c") - 6)
                        obj_streak.append(sv)

            # SOFT: break long NIGHT stretches with a DAY stretch (everyone who works nights).
            # nphase[i] = consecutive nights up to day i. A DAY shift resets it to 0; a week-off
            # CARRIES the count forward (a rest day alone is NOT a break). So the only way to end
            # a long night stretch is to actually work a stretch of day shifts — which is the ask.
            CAP = cfg["MAX_NIGHT_RUN"]
            for eid, e in emp.items():
                if (not e["eligible"]) or e["mx"] <= CAP:     # can't exceed the cap -> nothing to split
                    continue
                nphase = {}
                L = len(e["sched"])
                for i in e["sched"]:
                    base = nphase[i - 1] if (i - 1) in e["sched_set"] else 0
                    ph = model.NewIntVar(0, L, f"ph_{eid}_{i}")
                    model.Add(ph == base + 1).OnlyEnforceIf(night_v[(eid, i)])
                    model.Add(ph == 0).OnlyEnforceIf(day_v[(eid, i)])
                    model.Add(ph == base).OnlyEnforceIf(wo_v[(eid, i)])
                    nphase[i] = ph
                    pe = model.NewIntVar(0, L, f"pe_{eid}_{i}")
                    model.Add(pe >= ph - CAP)                  # penalise nights beyond the cap
                    obj_night_run.append(pe)

            # SOFT (optional): cross-month cumulative night fairness — minimise spread of totals
            cum_nights = {}
            if not df_hist.empty and "Emp_ID" in df_hist.columns and "Cumulative_Nights" in df_hist.columns:
                for _, r in df_hist.iterrows():
                    cum_nights[r["Emp_ID"]] = clean_num(r["Cumulative_Nights"])
            if cfg["W_CUM_NIGHT_FAIR"] > 0 and cum_nights:
                elig_ids = [eid for eid, e in emp.items() if e["eligible"] and e["sched"]]
                if elig_ids:
                    maxT = model.NewIntVar(0, 10000, "cum_max")
                    minT = model.NewIntVar(0, 10000, "cum_min")
                    for eid in elig_ids:
                        base = int(cum_nights.get(eid, 0) or 0)
                        tot = base + sum(NV(eid, i) for i in emp[eid]["sched"])
                        model.Add(maxT >= tot)
                        model.Add(minT <= tot)
                    spread = model.NewIntVar(0, 10000, "cum_spread")
                    model.Add(spread == maxT - minT)
                    obj_cum.append(spread)

            # SOFT (optional): month-to-month day/night stability anchor
            if cfg["W_STABILITY"] > 0:
                for eid, e in emp.items():
                    if not e["eligible"] or e["prev_work"] <= 0 or not e["sched"]:
                        continue
                    share = e["prev_nights"] / e["prev_work"]
                    target = int(round(share * max(0, len(e["sched"]) - e["entitlement"])))
                    tot = sum(NV(eid, i) for i in e["sched"])
                    dev = model.NewIntVar(0, len(e["sched"]), f"stab_{eid}")
                    model.Add(dev >= tot - target)
                    model.Add(dev >= target - tot)
                    obj_stab.append(dev)

            # OBJECTIVE
            model.Minimize(
                cfg["W_NIGHT_STAFF"] * sum(obj_night)
                + cfg["W_WO_RANGE"] * sum(obj_wo_range)
                + cfg["W_PREF"] * sum(obj_pref)
                + cfg["W_STREAK"] * sum(obj_streak)
                + cfg["W_WO_LOAD"] * sum(obj_wo_load)
                + cfg["W_NIGHT_RUN"] * sum(obj_night_run)
                + cfg["W_CUM_NIGHT_FAIR"] * sum(obj_cum)
                + cfg["W_STABILITY"] * sum(obj_stab)
            )

            # ---------- SOLVE ----------
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = float(cfg["SOLVER_TIME_LIMIT"])
            solver.parameters.num_search_workers = int(cfg["SOLVER_WORKERS"])
            solver.parameters.random_seed = int(cfg["SOLVER_SEED"])
            status = solver.Solve(model)
            status_name = solver.StatusName(status)
            if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                st.error(f"Solver returned {status_name}. The HARD constraints likely conflict "
                         f"(e.g. exact WO count + 9-day cap with a long prior streak, or frozen cells "
                         f"that break shift blocks). Try ALLOW_WO_FLEX=1 or relax the freeze.")
                st.stop()

            # ---------- EXTRACT ----------
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
                    nlo, nhi = night_band(r)
                    summ.append({
                        "Date": day_labels[i], "Role": r, "Active": active, "Working": work,
                        "Day": work - nig, "Night": nig,
                        "Night_%": round(100 * nig / work, 1) if work else 0,
                        "Night_target": f"{nlo}-{nhi}%" if r in night_roles else "Day only",
                        "WeekOffs": woc, "WO_%": round(100 * woc / active, 1) if active else 0,
                        "WO_target_%": round(100 * wo_target_frac(d, r), 1),
                    })
            summary_df = pd.DataFrame(summ)

            # ---------- constraint report ----------
            report = []
            for eid, e in emp.items():
                if not e["eligible"]:
                    continue
                tot = sum(1 for i in range(num_days) if cell(eid, i) == "N")
                if tot < e["mn"]:
                    report.append({"Type": "Night preference", "Who": eid,
                                   "Detail": f"{tot} nights, below preferred min {e['mn']} (short {e['mn']-tot}) "
                                             f"— likely needed to staff the floor"})
                elif tot > e["mx"]:
                    report.append({"Type": "Night preference", "Who": eid,
                                   "Detail": f"{tot} nights, above preferred max {e['mx']} (over {tot-e['mx']}) "
                                             f"— likely needed to staff the floor"})
            for _, s in summary_df.iterrows():
                nlo, nhi = night_band(s["Role"])
                if s["Role"] in night_roles and s["Working"] and not (nlo <= s["Night_%"] <= nhi):
                    report.append({"Type": "Night split", "Who": f"{s['Date']} / {s['Role']}",
                                   "Detail": f"nights at {s['Night_%']}% of {int(s['Working'])} working "
                                             f"(target {nlo}-{nhi}%)"})
                wlo, whi = math.floor(cfg["WO_LO"] * s["Active"]), math.ceil(cfg["WO_HI"] * s["Active"])
                if not (wlo <= s["WeekOffs"] <= whi):
                    report.append({"Type": "Week-off range", "Who": f"{s['Date']} / {s['Role']}",
                                   "Detail": f"{int(s['WeekOffs'])} offs vs allowed {wlo}-{whi}"})
            report_df = pd.DataFrame(report) if report else pd.DataFrame(
                [{"Type": "—", "Who": "—", "Detail": "All soft targets met."}])

            # ---------- integrity checks (verify HARD rules actually hold) ----------
            checks = []

            def add_check(name, ok, detail=""):
                checks.append({"Check": name, "Result": "PASS" if ok else "FAIL", "Detail": detail})

            max_run, run_bad = 0, []
            for eid, e in emp.items():
                seq = list(e["june_tail"]) + [1 if cell(eid, i) in ("D", "N") else 0 for i in range(num_days)]
                run = 0
                for v in seq:
                    run = run + 1 if v else 0
                    max_run = max(max_run, run)
                    if run > cfg["MAX_CONSEC_WORK_DAYS"]:
                        run_bad.append(eid)
            add_check(f"No working run > {cfg['MAX_CONSEC_WORK_DAYS']} days", not run_bad,
                      f"longest run seen: {max_run}" if not run_bad else f"violations: {set(run_bad)}")

            wo_bad = [eid for eid, e in emp.items()
                      if not cfg["ALLOW_WO_FLEX"] and e["sched"]
                      and sum(1 for i in range(num_days) if cell(eid, i) == "WOD") != e["entitlement"]]
            add_check("Monthly week-off count == entitlement", not wo_bad,
                      "" if not wo_bad else f"off for: {set(wo_bad)}")

            no_night_bad = [eid for eid, e in emp.items()
                            if not e["eligible"] and any(cell(eid, i) == "N" for i in range(num_days))]
            add_check("Day-only / female staff have no nights", not no_night_bad,
                      "" if not no_night_bad else f"violations: {set(no_night_bad)}")

            if cfg["ENFORCE_SHIFT_BLOCKS"]:
                flip_bad = []
                for eid, e in emp.items():
                    for i in range(num_days - 1):
                        if i in e["sched_set"] and (i + 1) in e["sched_set"]:
                            if {cell(eid, i), cell(eid, i + 1)} == {"D", "N"}:
                                flip_bad.append((eid, day_labels[i]))
                add_check("No Day/Night switch without a rest day", not flip_bad,
                          "" if not flip_bad else f"flips at: {flip_bad[:6]}")

            lock_bad = []
            for eid in current:
                if eid in emp:
                    for i in range(num_days):
                        lk = lock_for(eid, i)
                        if lk and cell(eid, i) != lk:
                            lock_bad.append((eid, day_labels[i]))
            add_check("Frozen cells preserved", not lock_bad,
                      "" if not lock_bad else f"changed: {lock_bad[:6]}")
            checks_df = pd.DataFrame(checks)

            # ---------- objective breakdown ----------
            def comp(terms):
                return int(sum(solver.Value(v) for v in terms)) if terms else 0
            breakdown = [
                ("Night staffing", cfg["W_NIGHT_STAFF"] * comp(obj_night)),
                ("Week-off range", cfg["W_WO_RANGE"] * comp(obj_wo_range)),
                ("Night preference", cfg["W_PREF"] * comp(obj_pref)),
                ("Long work streaks", cfg["W_STREAK"] * comp(obj_streak)),
                ("Week-off load shaping", cfg["W_WO_LOAD"] * comp(obj_wo_load)),
                ("Night-run splitting", cfg["W_NIGHT_RUN"] * comp(obj_night_run)),
                ("Cross-month fairness", cfg["W_CUM_NIGHT_FAIR"] * comp(obj_cum)),
                ("Month-to-month stability", cfg["W_STABILITY"] * comp(obj_stab)),
            ]
            breakdown_df = pd.DataFrame(
                [{"Component": n, "Penalty": p} for n, p in breakdown if p > 0]
                or [{"Component": "—", "Penalty": 0}])

            # ---------- write workbook ----------
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                roster_df.to_excel(writer, index=False, sheet_name="Generated_Roster")
                summary_df.to_excel(writer, index=False, sheet_name="Daily_Summary")
                report_df.to_excel(writer, index=False, sheet_name="Constraint_Report")
                checks_df.to_excel(writer, index=False, sheet_name="Integrity_Checks")
                breakdown_df.to_excel(writer, index=False, sheet_name="Objective_Breakdown")
                if warnings:
                    pd.DataFrame({"Warning": warnings}).to_excel(writer, index=False, sheet_name="Validation")

            month_name = days[0].strftime("%B_%Y")
            new_joiners = [eid for eid in emp if current and eid not in current]
            mode = ("fresh full solve" if not current else
                    (f"new-joiner build ({len(new_joiners)} new)" if freeze_date is None
                     else f"re-plan from {freeze_date.strftime('%d-%b')}"))
            blocks_state = "ON" if cfg["ENFORCE_SHIFT_BLOCKS"] else "OFF"
            st.success(f"✅ {month_name.replace('_', ' ')} — {status_name} · mode: {mode} · "
                       f"{num_days}-day period · {cfg['WO_PER_MONTH']} week-offs each · "
                       f"Day↔Night blocks: {blocks_state}")
            if not cfg["ENFORCE_SHIFT_BLOCKS"]:
                st.warning("Day↔Night shift blocks are OFF, so shifts can switch without a rest day. "
                           "To require a rest day between Day and Night, set ENFORCE_SHIFT_BLOCKS = TRUE "
                           "in the Config sheet (or remove that row to use the default, which is ON).")

            all_pass = all(c["Result"] == "PASS" for c in checks)
            if not all_pass:
                st.error("⚠️ Integrity check FAILED — see the Integrity_Checks tab before using this roster.")

            c1, c2, c3 = st.columns(3)
            outside = sum(1 for _, s in summary_df.iterrows()
                          if s["Role"] in night_roles and s["Working"]
                          and not (night_band(s["Role"])[0] <= s["Night_%"] <= night_band(s["Role"])[1]))
            pref_v = sum(1 for eid, e in emp.items() if e["eligible"]
                         and not (e["mn"] <= sum(1 for i in range(num_days) if cell(eid, i) == "N") <= e["mx"]))
            c1.metric("Day/role cells off night band", outside)
            c2.metric("People outside night band", pref_v)
            c3.metric("Integrity", "PASS" if all_pass else "FAIL")
            st.caption("Objective penalty by component:")
            st.dataframe(breakdown_df, use_container_width=True, hide_index=True)
            st.dataframe(summary_df, use_container_width=True, height=300)

            st.download_button("⬇️ Download Roster + Diagnostics (.xlsx)", data=buf.getvalue(),
                               file_name=f"{month_name}_Roster.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

        except Exception as exc:
            st.error(f"⚠️ An error occurred: {exc}")
            st.exception(exc)
