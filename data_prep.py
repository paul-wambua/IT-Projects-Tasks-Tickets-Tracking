"""Data layer for the IT Projects / Tasks / Tickets dashboard.

Pure pandas, no Streamlit imports, so it can be tested on its own.
"""
from __future__ import annotations

import datetime as dt
import io
import re
from collections import Counter

import numpy as np
import pandas as pd

CLOSED_STATUSES = ["Closed", "Resolved"]
TRACKER_COLUMNS = [
    "No", "Type", "Ticket #", "Created By", "Creation Date", "Issue",
    "Action/Comment", "Status", "Agent", "Group", "Deadline",
    "Resolved Date", "Total Time Taken",
]
DEFAULT_LISTS = {
    "Type": ["Project", "Task", "Ticket"],
    "Status": ["Open", "In Progress", "Pending", "Awaiting Customer Response",
               "Awaiting 3rd Party Response", "Resolved", "Closed"],
    "Agent": [], "Group": [],
}
STOPWORDS = set("""
the and for with from not are was has have this that you your our out all any can
new re fw fwd please kindly issue error unable when does did its into off per via
""".split())


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #
def _open(source):
    """Return a fresh file-like/path for pandas each time (BytesIO can't be re-read)."""
    if isinstance(source, (bytes, bytearray)):
        return io.BytesIO(source)
    return source


def read_workbook(source, sheet: str = "Tracker"):
    """Read the Tracker sheet (header row auto-detected) and the Lists sheet.

    `source` may be a file path or raw bytes. Returns (raw_df, lists_dict).
    """
    probe = pd.read_excel(_open(source), sheet_name=sheet, header=None, nrows=15)
    header_row = 2
    for i, row in probe.iterrows():
        if "Creation Date" in row.astype(str).values:
            header_row = i
            break
    raw = pd.read_excel(_open(source), sheet_name=sheet, header=header_row)
    raw.columns = [str(c).strip() for c in raw.columns]

    lists = {}
    try:
        lst = pd.read_excel(_open(source), sheet_name="Lists", usecols="A:D")
        for col in ["Type", "Status", "Agent", "Group"]:
            if col in lst.columns:
                lists[col] = lst[col].dropna().astype(str).str.strip().tolist()
    except Exception:
        pass
    return raw, lists


# --------------------------------------------------------------------------- #
# Cleaning
# --------------------------------------------------------------------------- #
def _to_minutes(value) -> float:
    """Ticket 'Total Time Taken' (timedelta / time / text) -> minutes."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return np.nan
    try:
        if isinstance(value, dt.time):
            return (value.hour * 3600 + value.minute * 60 + value.second) / 60
        if isinstance(value, str):
            value = pd.to_timedelta(value.strip())
        if isinstance(value, (pd.Timedelta, dt.timedelta)):
            secs = pd.Timedelta(value).total_seconds()
            return secs / 60 if secs >= 0 else np.nan
    except Exception:
        return np.nan
    return np.nan


def prepare(raw: pd.DataFrame, as_of: dt.date, fix_deadlines: bool = True) -> pd.DataFrame:
    """Clean the raw tracker and add the analytical columns used by the app."""
    df = raw.copy()
    for col in TRACKER_COLUMNS:
        if col not in df.columns:
            df[col] = np.nan
    df = df[df["Type"].notna() & df["Status"].notna()].copy()

    for col in ["Type", "Status", "Agent", "Group"]:
        df[col] = df[col].astype(str).str.strip()
    df["Created By"] = df["Created By"].astype(str).str.strip().str.title()
    df["Issue"] = df["Issue"].fillna("").astype(str).str.strip()
    df["Action/Comment"] = df["Action/Comment"].fillna("").astype(str).str.strip()

    for col in ["Creation Date", "Deadline", "Resolved Date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["Ticket #"] = pd.to_numeric(df["Ticket #"], errors="coerce").astype("Int64")

    as_of_ts = pd.Timestamp(as_of)
    df["Created Day"] = df["Creation Date"].dt.normalize()

    # Deadline typos (e.g. 2016 instead of 2026): shift by +10 years if that makes sense.
    df["Deadline Fixed"] = False
    if fix_deadlines:
        far_past = df["Deadline"].notna() & (
            (df["Created Day"] - df["Deadline"]).dt.days > 365
        )
        shifted = df.loc[far_past, "Deadline"] + pd.DateOffset(years=10)
        ok = shifted >= df.loc[far_past, "Created Day"]
        idx = shifted[ok].index
        df.loc[idx, "Deadline"] = shifted[ok]
        df.loc[idx, "Deadline Fixed"] = True

    df["Is Active"] = ~df["Status"].isin(CLOSED_STATUSES)

    # Resolution time: tickets -> minutes (from the system export); tasks/projects -> days.
    is_ticket = df["Type"].eq("Ticket")
    df["Resolution (min)"] = np.where(
        is_ticket, df["Total Time Taken"].map(_to_minutes), np.nan
    )
    days = (df["Resolved Date"] - df["Created Day"]).dt.days
    df["Resolution (days)"] = np.where(~is_ticket & df["Resolved Date"].notna(), days, np.nan)

    df["Overdue"] = (
        df["Is Active"] & df["Deadline"].notna() & (df["Deadline"].dt.normalize() < as_of_ts)
    )
    df["Days Overdue"] = np.where(
        df["Overdue"], (as_of_ts - df["Deadline"].dt.normalize()).dt.days, 0
    ).astype(int)
    df["Age (days)"] = (as_of_ts - df["Created Day"]).dt.days.clip(lower=0)
    df["Weekday"] = df["Creation Date"].dt.day_name()

    df = df.drop(columns=["Total Time Taken"])
    return df.sort_values("Creation Date", ascending=False).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# KPIs, trend, helpers
# --------------------------------------------------------------------------- #
def fmt_minutes(m: float) -> str:
    if m is None or pd.isna(m):
        return "n/a"
    if m < 1:
        return f"{m * 60:.0f}s"
    if m < 60:
        return f"{int(m)}m {int(round((m % 1) * 60)):02d}s"
    return f"{int(m // 60)}h {int(round(m % 60)):02d}m"


def compute_kpis(df: pd.DataFrame, as_of: dt.date) -> dict:
    total = len(df)
    active = int(df["Is Active"].sum())
    closed = total - active
    last7 = int((df["Created Day"] > pd.Timestamp(as_of) - pd.Timedelta(days=7)).sum())
    tick = df.loc[df["Type"].eq("Ticket"), "Resolution (min)"].dropna()
    return {
        "total": total,
        "last7": last7,
        "active": active,
        "closed": closed,
        "overdue": int(df["Overdue"].sum()),
        "completion": (closed / total) if total else 0.0,
        "median_res": float(tick.median()) if len(tick) else np.nan,
    }


def trend_table(df: pd.DataFrame, granularity: str = "Week"):
    """Items created per period by type, plus a rolling average of the total."""
    code = {"Day": "D", "Week": "W-SUN", "Month": "M"}[granularity]
    window = {"Day": 7, "Week": 4, "Month": 2}[granularity]
    d = df.dropna(subset=["Creation Date"])
    if d.empty:
        return pd.DataFrame(columns=["Period", "Type", "Items"]), pd.DataFrame(columns=["Period", "Total", "Rolling"])
    periods = d["Creation Date"].dt.to_period(code)
    full = pd.period_range(periods.min(), periods.max(), freq=code)
    pivot = (
        d.assign(Period=periods).groupby(["Period", "Type"]).size().unstack(fill_value=0)
        .reindex(full, fill_value=0)
    )
    total = pivot.sum(axis=1)
    rolling = total.rolling(window, min_periods=1).mean()
    long = pivot.reset_index(names="Period").melt(id_vars="Period", var_name="Type", value_name="Items")
    long["Period"] = long["Period"].dt.start_time
    totals = pd.DataFrame({"Period": full.start_time, "Total": total.values, "Rolling": rolling.values})
    return long, totals


def top_keywords(df: pd.DataFrame, n: int = 8) -> pd.DataFrame:
    words = Counter()
    for text in df.loc[df["Type"].eq("Ticket"), "Issue"]:
        for w in set(re.findall(r"[a-z]{3,}", text.lower())):
            if w not in STOPWORDS:
                words[w] += 1
    return pd.DataFrame(words.most_common(n), columns=["Keyword", "Tickets"])


def data_quality(df: pd.DataFrame) -> list[str]:
    notes = []
    fixed = int(df["Deadline Fixed"].sum())
    if fixed:
        nos = ", ".join(str(x) for x in df.loc[df["Deadline Fixed"], "No"].tolist()[:5])
        notes.append(f"{fixed} deadline(s) were years in the past vs. creation date (item No {nos}) "
                     "and were auto-corrected by +10 years. Check the source sheet.")
    no_deadline = df[df["Is Active"] & df["Type"].ne("Ticket") & df["Deadline"].isna()]
    if len(no_deadline):
        notes.append(f"{len(no_deadline)} active task/project(s) have no deadline, so they can never show as overdue.")
    t = df[df["Type"].eq("Ticket")]
    miss = int(t["Ticket #"].isna().sum())
    if miss:
        notes.append(f"{miss} ticket(s) have no ticket number.")
    dup = t["Ticket #"].dropna()
    dup = int(dup.duplicated().sum())
    if dup:
        notes.append(f"{dup} duplicate ticket number(s) found.")
    missing_res = int(t["Resolution (min)"].isna().sum())
    if missing_res:
        notes.append(f"{missing_res} ticket(s) have no usable resolution time.")
    return notes


# --------------------------------------------------------------------------- #
# Insights
# --------------------------------------------------------------------------- #
def build_insights(df: pd.DataFrame, as_of: dt.date, sla_minutes: float = 15) -> list[tuple[str, str]]:
    """Plain-language findings computed from whatever is currently filtered."""
    if df.empty:
        return [("ℹ️", "No items match the current filters.")]
    out: list[tuple[str, str]] = []
    total = len(df)

    mix = df["Type"].value_counts()
    out.append(("📦", f"**{total} items in view.** {mix.index[0]}s dominate with {mix.iloc[0]} "
                      f"({mix.iloc[0] / total:.0%} of the workload)."))

    active = df[df["Is Active"]]
    if len(active):
        top_status = active["Status"].value_counts()
        out.append(("🧱", f"**{len(active)} items are still open.** The biggest bucket is "
                          f"“{top_status.index[0]}” ({top_status.iloc[0]})."))
        waiting = int(active["Status"].str.contains("Awaiting").sum())
        if waiting:
            out.append(("⏳", f"**{waiting} open items are blocked waiting on someone else** "
                              "(customer or 3rd party), so they need follow-up rather than work."))
        top_agent = active["Agent"].value_counts()
        out.append(("👷", f"**{top_agent.index[0]} carries the most open work** "
                          f"({top_agent.iloc[0]} of {len(active)} open items)."))
    else:
        out.append(("✅", "**Nothing is open** in the current view."))

    od = df[df["Overdue"]].sort_values("Days Overdue", ascending=False)
    if len(od):
        worst = od.iloc[0]
        by_agent = od["Agent"].value_counts()
        out.append(("🚨", f"**{len(od)} overdue item(s).** The oldest is “{worst['Issue'][:60]}” "
                          f"({worst['Agent']}), {int(worst['Days Overdue'])} days past deadline. "
                          f"{by_agent.index[0]} has the most overdue ({by_agent.iloc[0]})."))
    else:
        out.append(("🟢", "**No overdue items** as of the selected date."))

    grp = df["Group"].value_counts()
    out.append(("🏢", f"**{grp.index[0]} generates the most demand** ({grp.iloc[0]} items, "
                      f"{grp.iloc[0] / total:.0%})."))

    long, totals = trend_table(df, "Week")
    if len(totals) >= 3:
        peak = totals.loc[totals["Total"].idxmax()]
        avg = totals["Total"].mean()
        latest = totals.iloc[-1]
        ratio = latest["Total"] / avg if avg else 0
        out.append(("📈", f"**Weekly volume averages {avg:.1f} items.** Peak week started "
                          f"{peak['Period']:%d %b} with {int(peak['Total'])}. The latest week "
                          f"(may be partial) has {int(latest['Total'])} ({ratio:.1f}x average)."))

    tk = df[df["Type"].eq("Ticket")]
    res = tk["Resolution (min)"].dropna()
    if len(res):
        within = (res <= sla_minutes).mean()
        text = (f"**Median ticket resolution is {fmt_minutes(res.median())}**, "
                f"90th percentile {fmt_minutes(res.quantile(0.9))}. "
                f"{within:.0%} of tickets finish within the {sla_minutes:g}-minute target.")
        by = tk.groupby("Agent")["Resolution (min)"].agg(["median", "count"])
        by = by[by["count"] >= 3].sort_values("median", ascending=False)
        if len(by) >= 2:
            text += f" Slowest median: {by.index[0]} ({fmt_minutes(by['median'].iloc[0])})."
        out.append(("⚡", text))

    if len(tk):
        wd = tk["Weekday"].value_counts()
        req = tk["Created By"].value_counts()
        top3 = req.head(3).sum() / len(tk)
        out.append(("🙋", f"**{req.index[0]} raises the most tickets** ({req.iloc[0]}); the top 3 "
                          f"requesters account for {top3:.0%} of tickets. Busiest day: {wd.index[0]}."))
        kw = top_keywords(df, 3)
        if len(kw):
            names = ", ".join(f"“{k}” ({n})" for k, n in zip(kw["Keyword"], kw["Tickets"]))
            out.append(("🔁", f"**Recurring ticket themes:** {names}. Repeat themes are candidates "
                              "for a permanent fix or user training."))

    slow_proj = df[df["Type"].ne("Ticket") & df["Resolution (days)"].notna()]
    if len(slow_proj):
        out.append(("🗓️", f"**Closed tasks/projects took {slow_proj['Resolution (days)'].mean():.0f} "
                          f"days on average** ({len(slow_proj)} completed)."))
    return out
