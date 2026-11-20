"""IT Projects, Tasks & Tickets dashboard.

Run:  streamlit run app.py
"""
from __future__ import annotations

import datetime as dt
import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import data_prep as dp

DEFAULT_FILE = Path(__file__).parent / "Projects_Tracking-IT_Department.xlsx"

STATUS_COLORS = {
    "Closed": "#7f8c8d", "Resolved": "#2ecc71", "In Progress": "#3498db",
    "Pending": "#f1c40f", "Awaiting 3rd Party Response": "#e67e22",
    "Awaiting Customer Response": "#9b59b6", "Open": "#e74c3c",
}
TYPE_COLORS = {"Ticket": "#3498db", "Task": "#e67e22", "Project": "#2ecc71"}
WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

st.set_page_config(page_title="IT Department Dashboard", page_icon="🛠️", layout="wide")


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner="Reading workbook…")
def load_workbook_bytes(content: bytes):
    return dp.read_workbook(content)


def tidy(fig, height=340):
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=40, b=10),
                      legend=dict(orientation="h", y=-0.2), title_font_size=15)
    return fig


def style_status(styler, column="Status"):
    def colour(val):
        c = STATUS_COLORS.get(val)
        return f"background-color: {c}33; color: inherit; font-weight: 600" if c else ""
    fn = getattr(styler, "map", None) or styler.applymap  # pandas >= 2.1 vs older
    return fn(colour, subset=[column])


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Tracker")
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Sidebar layout (containers keep the visual order while code runs in a
# different order: data -> settings -> new item -> filters)
# --------------------------------------------------------------------------- #
st.sidebar.title("🛠️ Controls")
src_box = st.sidebar.container()
filter_box = st.sidebar.container()
settings_box = st.sidebar.container()
add_box = st.sidebar.container()

# ---- 1. Data source -------------------------------------------------------
with src_box:
    st.subheader("Data")
    upload = st.file_uploader("Upload tracker (.xlsx)", type=["xlsx"],
                              help="Needs a 'Tracker' sheet like your IT tracking workbook.")
if upload is not None:
    content = upload.getvalue()
elif DEFAULT_FILE.exists():
    content = DEFAULT_FILE.read_bytes()
else:
    st.info("Upload your tracker workbook in the sidebar to get started.")
    st.stop()

try:
    raw_base, lists = load_workbook_bytes(content)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read the 'Tracker' sheet: {exc}")
    st.stop()

# ---- 2. Settings (inputs) -------------------------------------------------
with settings_box:
    st.subheader("Settings")
    as_of = st.date_input("Report date (for overdue/age)", value=dt.date.today())
    sla = st.number_input("Ticket target (minutes)", min_value=1.0, value=15.0, step=1.0,
                          help="Used for the '% within target' insight.")
    fix_deadlines = st.checkbox("Auto-fix deadline typos (+10 years)", value=True,
                                help="e.g. a deadline in 2016 on an item created in 2026.")

# ---- 3. Add a new item (session only) ------------------------------------
if "extra" not in st.session_state:
    st.session_state["extra"] = []


def options(col, series):
    base = lists.get(col) or dp.DEFAULT_LISTS.get(col, [])
    return sorted(set(base) | set(series.dropna().astype(str)), key=lambda x: (x not in base, x))


with add_box:
    with st.expander("➕ Add an item"):
        with st.form("add_item", clear_on_submit=True):
            n_type = st.selectbox("Type", options("Type", raw_base["Type"]))
            n_issue = st.text_input("Issue")
            n_by = st.text_input("Created by")
            n_status = st.selectbox("Status", options("Status", raw_base["Status"]))
            n_agent = st.selectbox("Agent", options("Agent", raw_base["Agent"]))
            n_group = st.selectbox("Group", options("Group", raw_base["Group"]))
            n_created = st.date_input("Creation date", value=dt.date.today())
            has_deadline = st.checkbox("Has deadline")
            n_deadline = st.date_input("Deadline", value=dt.date.today() + dt.timedelta(days=7))
            n_comment = st.text_area("Action / comment", height=70)
            if st.form_submit_button("Add"):
                if not n_issue.strip():
                    st.warning("Please describe the issue.")
                else:
                    next_no = int(pd.to_numeric(raw_base["No"], errors="coerce").max() or 0) + 1 \
                        + len(st.session_state["extra"])
                    st.session_state["extra"].append({
                        "No": next_no, "Type": n_type, "Ticket #": None,
                        "Created By": n_by or "Unknown",
                        "Creation Date": pd.Timestamp(n_created), "Issue": n_issue,
                        "Action/Comment": n_comment, "Status": n_status, "Agent": n_agent,
                        "Group": n_group,
                        "Deadline": pd.Timestamp(n_deadline) if has_deadline else pd.NaT,
                        "Resolved Date": pd.NaT, "Total Time Taken": None,
                    })
        if st.session_state["extra"]:
            st.caption(f"{len(st.session_state['extra'])} new item(s) added this session "
                       "(not saved to the source file; use the download in the Data tab).")
            if st.button("Clear added items"):
                st.session_state["extra"] = []
                st.rerun()

raw_all = raw_base
if st.session_state["extra"]:
    raw_all = pd.concat([raw_base, pd.DataFrame(st.session_state["extra"])], ignore_index=True)

df_all = dp.prepare(raw_all, as_of, fix_deadlines)

# ---- 4. Filters -----------------------------------------------------------
with filter_box:
    st.subheader("Filters")
    types = st.multiselect("Type", sorted(df_all["Type"].unique()), default=sorted(df_all["Type"].unique()))
    statuses = st.multiselect("Status", sorted(df_all["Status"].unique()), default=sorted(df_all["Status"].unique()))
    agents = st.multiselect("Agent", sorted(df_all["Agent"].unique()), default=sorted(df_all["Agent"].unique()))
    groups = st.multiselect("Group", sorted(df_all["Group"].unique()), default=sorted(df_all["Group"].unique()))
    dmin = df_all["Created Day"].min().date()
    dmax = max(df_all["Created Day"].max().date(), as_of)
    date_range = st.date_input("Created between", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    only_active = st.toggle("Open items only", value=False)
    only_overdue = st.toggle("Overdue only", value=False)
    search = st.text_input("Search issue / comment / requester")

start, end = (date_range if isinstance(date_range, (tuple, list)) and len(date_range) == 2
              else (dmin, dmax))
mask = (
    df_all["Type"].isin(types) & df_all["Status"].isin(statuses)
    & df_all["Agent"].isin(agents) & df_all["Group"].isin(groups)
    & (df_all["Created Day"] >= pd.Timestamp(start)) & (df_all["Created Day"] <= pd.Timestamp(end))
)
if only_active:
    mask &= df_all["Is Active"]
if only_overdue:
    mask &= df_all["Overdue"]
if search.strip():
    s = search.strip().lower()
    mask &= (df_all["Issue"].str.lower().str.contains(s, regex=False)
             | df_all["Action/Comment"].str.lower().str.contains(s, regex=False)
             | df_all["Created By"].str.lower().str.contains(s, regex=False))
df = df_all[mask].copy()

# --------------------------------------------------------------------------- #
# Header + KPIs
# --------------------------------------------------------------------------- #
st.title("🛠️ IT Projects, Tasks & Tickets Dashboard")
st.caption(f"Report date: {as_of:%d %b %Y}  •  Showing {len(df)} of {len(df_all)} items")

k = dp.compute_kpis(df, as_of)
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total items", k["total"], delta=f"{k['last7']} in last 7 days", border=True)
c2.metric("Open", k["active"], border=True)
c3.metric("Closed / resolved", k["closed"], border=True)
c4.metric("Overdue", k["overdue"], delta="needs action" if k["overdue"] else None,
          delta_color="inverse", border=True)
c5.metric("Completion rate", f"{k['completion']:.0%}", border=True)
c6.metric("Median ticket resolution", dp.fmt_minutes(k["median_res"]), border=True)

if df.empty:
    st.warning("No items match the current filters.")
    st.stop()

tab_dash, tab_ins, tab_data = st.tabs(["📊 Dashboard", "💡 Insights", "🗂️ Data"])

# --------------------------------------------------------------------------- #
# Dashboard tab
# --------------------------------------------------------------------------- #
with tab_dash:
    st.subheader("Trend")
    gran = st.radio("Group by", ["Day", "Week", "Month"], index=1, horizontal=True)
    long, totals = dp.trend_table(df, gran)
    fig = px.bar(long, x="Period", y="Items", color="Type", color_discrete_map=TYPE_COLORS,
                 barmode="stack", title=f"Items created per {gran.lower()}")
    fig.add_trace(go.Scatter(x=totals["Period"], y=totals["Rolling"], mode="lines",
                             name="Rolling average", line=dict(color="#c0392b", width=3)))
    st.plotly_chart(tidy(fig, 360))

    a, b = st.columns(2)
    with a:
        by_status = df["Status"].value_counts().reset_index()
        by_status.columns = ["Status", "Items"]
        fig = px.pie(by_status, names="Status", values="Items", hole=0.55, color="Status",
                     color_discrete_map=STATUS_COLORS, title="Status mix")
        st.plotly_chart(tidy(fig))
    with b:
        wl = df.groupby(["Agent", "Status"]).size().reset_index(name="Items")
        order = df["Agent"].value_counts().index.tolist()[::-1]
        fig = px.bar(wl, y="Agent", x="Items", color="Status", orientation="h",
                     color_discrete_map=STATUS_COLORS, category_orders={"Agent": order},
                     title="Workload by agent")
        st.plotly_chart(tidy(fig))

    a, b = st.columns(2)
    with a:
        g = df.groupby(["Group", "Type"]).size().reset_index(name="Items")
        g_order = df["Group"].value_counts().index.tolist()
        fig = px.bar(g, x="Group", y="Items", color="Type", color_discrete_map=TYPE_COLORS,
                     category_orders={"Group": g_order}, title="Demand by business group")
        st.plotly_chart(tidy(fig))
    with b:
        tk = df[df["Type"].eq("Ticket")].dropna(subset=["Resolution (min)"])
        if tk.empty:
            st.info("No tickets with resolution time in the current view.")
        else:
            r = tk.groupby("Agent")["Resolution (min)"].median().reset_index().sort_values("Resolution (min)")
            fig = px.bar(r, x="Agent", y="Resolution (min)", title="Median ticket resolution (minutes)",
                         color_discrete_sequence=["#3498db"])
            fig.add_hline(y=sla, line_dash="dash", line_color="#c0392b",
                          annotation_text=f"target {sla:g} min")
            st.plotly_chart(tidy(fig))

    a, b = st.columns(2)
    tickets = df[df["Type"].eq("Ticket")]
    with a:
        if tickets.empty:
            st.info("No tickets in the current view.")
        else:
            wd = tickets["Weekday"].value_counts().reindex(WEEKDAYS, fill_value=0).reset_index()
            wd.columns = ["Weekday", "Tickets"]
            fig = px.bar(wd, x="Weekday", y="Tickets", title="Tickets by weekday",
                         color_discrete_sequence=["#3498db"])
            st.plotly_chart(tidy(fig))
    with b:
        if tickets.empty:
            st.info("No tickets in the current view.")
        else:
            req = tickets["Created By"].value_counts().head(10).reset_index()
            req.columns = ["Requester", "Tickets"]
            fig = px.bar(req.iloc[::-1], y="Requester", x="Tickets", orientation="h",
                         title="Top 10 requesters", color_discrete_sequence=["#3498db"])
            st.plotly_chart(tidy(fig))

    a, b = st.columns(2)
    with a:
        kw = dp.top_keywords(df, 10)
        if kw.empty:
            st.info("No ticket text to analyse.")
        else:
            fig = px.bar(kw.iloc[::-1], y="Keyword", x="Tickets", orientation="h",
                         title="Recurring words in ticket titles", color_discrete_sequence=["#9b59b6"])
            st.plotly_chart(tidy(fig))
    with b:
        st.markdown("**Needs attention** (open items, most overdue first)")
        att = df[df["Is Active"]].sort_values(["Days Overdue", "Age (days)"], ascending=False)
        att = att[["No", "Type", "Issue", "Status", "Agent", "Deadline", "Days Overdue"]]
        st.dataframe(style_status(att.style.format({"Deadline": "{:%d %b %Y}"}, na_rep="-")),
                     hide_index=True, height=300)

# --------------------------------------------------------------------------- #
# Insights tab
# --------------------------------------------------------------------------- #
with tab_ins:
    st.subheader("What the data is telling you")
    st.caption("Generated automatically from the filtered data; change the filters to focus on a team or period.")
    for icon, text in dp.build_insights(df, as_of, sla):
        st.markdown(f"{icon}  {text}")
    notes = dp.data_quality(df)
    if notes:
        st.subheader("Data quality")
        for n in notes:
            st.warning(n)

# --------------------------------------------------------------------------- #
# Data tab
# --------------------------------------------------------------------------- #
with tab_data:
    st.subheader("Filtered data")
    cols = ["No", "Type", "Ticket #", "Created By", "Creation Date", "Issue", "Action/Comment",
            "Status", "Agent", "Group", "Deadline", "Resolved Date", "Resolution (min)",
            "Resolution (days)", "Overdue", "Days Overdue"]
    view = df[cols]
    fmt = {"Creation Date": "{:%d %b %Y %H:%M}", "Deadline": "{:%d %b %Y}",
           "Resolved Date": "{:%d %b %Y}", "Resolution (min)": "{:.1f}",
           "Resolution (days)": "{:.0f}", "Ticket #": "{}"}
    st.dataframe(style_status(view.style.format(fmt, na_rep="-")), hide_index=True, height=520)

    d1, d2 = st.columns(2)
    d1.download_button("⬇️ Download filtered data (CSV)", view.to_csv(index=False).encode("utf-8"),
                       file_name="it_tracker_filtered.csv", mime="text/csv")
    d2.download_button("⬇️ Download all data incl. added items (Excel)", to_excel_bytes(df_all[cols]),
                       file_name="it_tracker_all.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
