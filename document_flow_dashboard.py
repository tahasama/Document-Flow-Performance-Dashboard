import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import date

st.set_page_config(
    page_title="Document Control Audit Report",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

C_NAVY  = "#1a3a5c"
C_BLUE  = "#2271b3"
C_GREEN = "#1a7a45"
C_AMBER = "#b8860b"
C_RED   = "#9b2335"
C_GREY  = "#5a6475"
C_LIGHT = "#f5f6f8"
C_BDR   = "#dde2e8"

matplotlib.rcParams.update({
    "font.family":       "sans-serif",
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.titleweight":  "bold",
    "axes.titlesize":    12,
    "axes.labelsize":    9.5,
    "xtick.labelsize":   9,
    "ytick.labelsize":   9,
    "figure.facecolor":  "white",
    "axes.facecolor":    "white",
    "grid.color":        "#e8eaed",
    "grid.linewidth":    0.7,
})

CSS = f"""
<style>
* {{ box-sizing: border-box; }}
.report-title {{
    font-size: 1.75rem; font-weight: 800; color: {C_NAVY};
    margin: 0; line-height: 1.2; letter-spacing: -0.02em;
}}
.report-sub {{
    font-size: 0.74rem; color: {C_GREY}; margin: 0.18rem 0 0 0;
    text-transform: uppercase; letter-spacing: 0.08em;
}}
.meta-block {{
    background: white; border: 1px solid {C_BDR};
    border-left: 4px solid {C_BLUE}; border-radius: 5px;
    padding: 0.65rem 1.2rem; margin: 0.8rem 0 1.5rem 0; font-size: 0.86rem;
}}
.meta-block table {{ border-collapse: collapse; width: 100%; }}
.meta-block td {{ padding: 0.15rem 1.1rem 0.15rem 0.4rem; vertical-align: top; }}
.meta-block td:nth-child(odd) {{
    font-weight: 700; color: {C_GREY}; white-space: nowrap;
    width: 130px; font-size: 0.77rem; text-transform: uppercase; letter-spacing: 0.05em;
}}
.hs-wrap {{
    background: white; border: 1px solid {C_BDR}; border-radius: 6px;
    padding: 0.7rem 0.9rem; text-align: center;
}}
.hs-score {{
    font-size: 2.1rem; font-weight: 800; line-height: 1.05; margin: 0;
}}
.hs-denom {{ font-size: 0.85rem; font-weight: 400; opacity: 0.6; }}
.hs-lbl {{
    font-size: 0.66rem; text-transform: uppercase; letter-spacing: 0.09em;
    color: {C_GREY}; margin: 0.18rem 0 0.12rem 0;
}}
.hs-status {{ font-size: 0.8rem; font-weight: 700; margin: 0; }}
.kpi {{
    background: white; border: 1px solid {C_BDR};
    border-top: 3px solid; border-radius: 5px; padding: 0.8rem 1rem 0.85rem 1rem;
}}
.kpi .kv {{ font-size: 1.6rem; font-weight: 800; line-height: 1.1; margin: 0; }}
.kpi .kl {{
    font-size: 0.7rem; color: {C_GREY}; margin: 0.25rem 0 0 0;
    text-transform: uppercase; letter-spacing: 0.06em;
}}
.sec {{
    font-size: 0.68rem; font-weight: 800; color: {C_BLUE};
    text-transform: uppercase; letter-spacing: 0.12em;
    border-bottom: 2px solid {C_BLUE}; padding-bottom: 0.35rem;
    margin: 2.2rem 0 0.75rem 0;
}}
.fc {{
    border-radius: 5px; padding: 1rem 1.3rem;
    margin-bottom: 0.85rem; border-left: 5px solid;
}}
.fc.red   {{ background: #fdf1f1; border-color: {C_RED}; }}
.fc.amber {{ background: #fefbee; border-color: #c8940a; }}
.fc.green {{ background: #f0faf4; border-color: {C_GREEN}; }}
.fc-badge {{
    display: inline-block; font-size: 0.62rem; font-weight: 800;
    text-transform: uppercase; letter-spacing: 0.1em;
    padding: 0.1rem 0.5rem; border-radius: 3px; margin-bottom: 0.35rem;
}}
.fc-badge.red   {{ background: #f5c6cb; color: {C_RED}; }}
.fc-badge.amber {{ background: #fde9a0; color: #7a5200; }}
.fc-badge.green {{ background: #b8e8c8; color: {C_GREEN}; }}
.fc h4 {{ margin: 0 0 0.4rem 0; font-size: 0.94rem; font-weight: 700; color: #111; }}
.fc .fc-body {{
    font-size: 0.87rem; line-height: 1.65; color: #2c2c2c; margin: 0 0 0.5rem 0;
}}
.fc ul.fc-list {{
    margin: 0.15rem 0 0 1.1rem; padding: 0;
    font-size: 0.85rem; line-height: 1.7; color: #333;
}}
.fc ul.fc-list li {{ margin-bottom: 0.1rem; }}
.fc ul.rnames {{
    margin: 0.08rem 0 0.35rem 1.2rem; padding: 0; list-style: none;
    font-size: 0.82rem; color: #555; line-height: 1.6;
}}
.fc ul.rnames li::before {{ content: "— "; color: #999; }}
.ibox {{
    border-left: 3px solid {C_BLUE}; background: #eef5fb;
    border-radius: 0 4px 4px 0; padding: 0.65rem 0.95rem;
    font-size: 0.86rem; color: #192a3a; margin: 0.4rem 0 0.9rem 0; line-height: 1.65;
}}
.tnote {{
    background: {C_LIGHT}; border-left: 3px solid #aab5c0;
    border-radius: 0 4px 4px 0; padding: 0.45rem 0.85rem;
    font-size: 0.81rem; color: #444; margin: 0.2rem 0 0.8rem 0; line-height: 1.6;
}}
.sc {{
    background: white; border: 1px solid {C_BDR}; border-radius: 5px;
    padding: 1rem 1.1rem; font-size: 0.86rem; line-height: 1.7; height: 100%;
}}
.sc h5 {{
    margin: 0 0 0.55rem 0; font-size: 0.68rem; color: {C_BLUE};
    text-transform: uppercase; letter-spacing: 0.1em; font-weight: 800;
    border-bottom: 1px solid {C_BDR}; padding-bottom: 0.3rem;
}}
.mrow {{
    display: flex; justify-content: space-between; padding: 0.2rem 0;
    border-bottom: 1px solid #f0f2f4; font-size: 0.85rem;
}}
.mrow:last-child {{ border-bottom: none; }}
.mval {{ font-weight: 700; }}
.scope {{
    background: {C_LIGHT}; border: 1px solid {C_BDR}; border-radius: 5px;
    padding: 0.75rem 1rem; font-size: 0.79rem; color: {C_GREY};
    margin-top: 2.5rem; line-height: 1.7;
}}
.welcome {{
    background: white; border: 1px solid {C_BDR}; border-radius: 7px;
    padding: 1.8rem 2.2rem; margin-top: 0.8rem; max-width: 660px;
}}
.welcome h3 {{ color: {C_NAVY}; margin: 0 0 0.7rem 0; font-size: 1rem; }}
.welcome li {{ margin-bottom: 0.35rem; font-size: 0.88rem; line-height: 1.6; }}
.dataframe table {{ width: 100%; font-size: 9.5pt; }}
.dataframe th {{ font-size: 10pt; padding: 6px 9px; background: {C_LIGHT}; }}
.dataframe td {{ font-size: 9.5pt; padding: 6px 9px; }}
.dataframe th.row_heading {{ min-width: 180px; text-align: left; }}
.stApp header button[kind="header"]:not(:last-child) {{ display: none !important; }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def clean_header(df):
    markers = {"document no", "document no.", "file", "step name", "assigned to",
               "workflow name", "date in", "planned submission", "revision"}
    header_row = 0
    for i in range(min(12, len(df))):
        row_vals = {str(v).strip().lower() for v in df.iloc[i].dropna()}
        if row_vals & markers:
            header_row = i
            break
    df = df.iloc[header_row:].copy()
    df.columns = df.iloc[0]
    df = df[1:]
    df.columns = df.columns.str.strip()
    df.reset_index(drop=True, inplace=True)
    return df


def kpi_color(v, good=80, warn=60):
    if v >= good: return C_GREEN
    if v >= warn: return C_AMBER
    return C_RED


def compute_health_score(supplier_on_time, review_on_time, backlog, risk_df):
    backlog_score = max(0, 100 - (backlog / 40) * 100)
    avg_breach    = risk_df[["Sub_Breach_%","Rev_Breach_%"]].max(axis=1).mean() if not risk_df.empty and "Sub_Breach_%" in risk_df.columns else 0
    breach_score  = max(0, 100 - avg_breach)
    return round(
        0.25 * supplier_on_time + 0.35 * review_on_time
        + 0.25 * backlog_score  + 0.15 * breach_score, 1
    )


def _fig(w=10, h=4.5):
    fig, ax = plt.subplots(figsize=(w, h))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    return fig, ax


def sec(label):
    st.markdown(f'<div class="sec">{label}</div>', unsafe_allow_html=True)

def ibox(html):
    st.markdown(f'<div class="ibox">{html}</div>', unsafe_allow_html=True)

def tnote(html):
    st.markdown(f'<div class="tnote">{html}</div>', unsafe_allow_html=True)

def kpi_tile(val, label, color):
    return (f'<div class="kpi" style="border-top-color:{color};">'
            f'<p class="kv" style="color:{color};">{val}</p>'
            f'<p class="kl">{label}</p></div>')


def process_data(supplier_docs, workflow):
    workflow = workflow.rename(columns={
        "Document No.": "Document No", "Assigned To": "Assigned_To",
        "Date Completed": "Date Completed", "Date Due": "Date Due",
        "Step Status": "Step Status", "Step Outcome": "Step Outcome",
        "Date In": "Date In",
    })
    supplier_docs["Document No"] = supplier_docs["Document No"].astype(str).str.strip()
    workflow["Document No"]      = workflow["Document No"].astype(str).str.strip()
    supplier_docs["Planned Submission Date"] = pd.to_datetime(
        supplier_docs["Planned Submission Date"], errors="coerce")
    for col in ["Date In", "Date Completed", "Date Due"]:
        workflow[col] = pd.to_datetime(workflow[col], errors="coerce")

    supplier_docs["_Rev"] = 0
    for col in ["Revision", "Document Revision"]:
        if col in supplier_docs.columns:
            supplier_docs["_Rev"] = pd.to_numeric(supplier_docs[col], errors="coerce").fillna(0).astype(int)
            break

    wf_done = (workflow[workflow["Step Status"] == "Completed"].copy()
               if "Step Status" in workflow.columns else workflow.copy())

    outcome_dist = pd.Series(dtype=int)
    if "Step Outcome" in wf_done.columns:
        def _bucket(s):
            s = str(s).strip().lower()
            if s.startswith("c1"): return "C1 — Approved"
            if s.startswith("c2"): return "C2 — Approved with Comments"
            if s.startswith("c3"): return "C3 — Rejected"
            if s.startswith("c4"): return "C4 — For Information"
            return "Other"
        outcome_dist = wf_done["Step Outcome"].dropna().astype(str).apply(_bucket).value_counts()

    agg = (wf_done.groupby("Document No").agg(
        FirstDateIn  = ("Date In",        "min"),
        ApprovedDate = ("Date Completed", "max"),
        DueDate      = ("Date Due",       "max"),
        Assigned_To  = ("Assigned_To",    "last"),
    ).reset_index())
    df = supplier_docs.merge(agg, on="Document No", how="left")

    has_sub = df["FirstDateIn"].notna() & df["Planned Submission Date"].notna()
    df["SupplierOnTime"] = False
    df.loc[has_sub, "SupplierOnTime"] = (
        df.loc[has_sub, "FirstDateIn"] <= df.loc[has_sub, "Planned Submission Date"])
    supplier_on_time = df.loc[has_sub, "SupplierOnTime"].mean() * 100 if has_sub.any() else 0

    has_rev = df["ApprovedDate"].notna() & df["DueDate"].notna()
    df["ReviewerOnTime"] = np.nan
    df.loc[has_rev, "ReviewerOnTime"] = (
        df.loc[has_rev, "ApprovedDate"] <= df.loc[has_rev, "DueDate"]).astype(float)
    review_on_time = df.loc[has_rev, "ReviewerOnTime"].mean() * 100 if has_rev.any() else 0

    df["CycleTime"] = (df["ApprovedDate"] - df["Planned Submission Date"]).dt.days
    median_cycle    = df["CycleTime"].dropna().median()

    today   = pd.Timestamp.today().normalize()
    in_wf   = df["DueDate"].notna()
    df["Overdue"] = False
    df.loc[in_wf, "Overdue"] = (
        (df.loc[in_wf, "DueDate"] < today) &
        (df.loc[in_wf, "ApprovedDate"].isna() |
         (df.loc[in_wf, "ApprovedDate"] > df.loc[in_wf, "DueDate"])))
    backlog = df.loc[in_wf, "Overdue"].mean() * 100 if in_wf.any() else 0

    valid = df.dropna(subset=["Planned Submission Date", "FirstDateIn", "ApprovedDate"]).copy()
    valid["SupplierDelay"]  = (valid["FirstDateIn"] - valid["Planned Submission Date"]).dt.days.clip(lower=0)
    valid["ReviewDuration"] = (valid["ApprovedDate"] - valid["FirstDateIn"]).dt.days.clip(lower=0)
    valid["Total"]          = valid["SupplierDelay"] + valid["ReviewDuration"]
    valid = valid[valid["Total"] > 0]
    total_sum    = valid["Total"].sum()
    supplier_pct = (valid["SupplierDelay"].sum()  / total_sum * 100) if total_sum > 0 else 0
    review_pct   = (valid["ReviewDuration"].sum() / total_sum * 100) if total_sum > 0 else 0

    sp_base = df[has_sub].copy()
    sp_base["SupplierName"] = (
        sp_base["Select List 5"].fillna("Unknown") if "Select List 5" in sp_base.columns else "Unknown")
    sp_base["SupplierDelay_sp"] = (
        sp_base["FirstDateIn"] - sp_base["Planned Submission Date"]).dt.days.clip(lower=0)
    def _avg_late(s):
        late = s[s > 0]; return late.mean() if len(late) > 0 else 0.0
    supplier_perf = (sp_base.groupby("SupplierName").agg(
        Total_Docs  = ("Document No",       "count"),
        OnTime_Docs = ("SupplierOnTime",    lambda x: int(x.sum())),
        Avg_Delay   = ("SupplierDelay_sp",  _avg_late),
        Max_Delay   = ("SupplierDelay_sp",  "max"),
    ).reset_index())
    supplier_perf["Delayed_Docs"] = supplier_perf["Total_Docs"] - supplier_perf["OnTime_Docs"]
    supplier_perf["OnTime_%"]     = (supplier_perf["OnTime_Docs"] / supplier_perf["Total_Docs"] * 100).round(1)
    supplier_perf["Avg_Delay"]    = supplier_perf["Avg_Delay"].round(1)
    supplier_perf["Max_Delay"]    = supplier_perf["Max_Delay"].fillna(0).astype(int)
    supplier_perf = supplier_perf.sort_values("Delayed_Docs", ascending=False).set_index("SupplierName")

    flow = df.dropna(subset=["FirstDateIn", "ApprovedDate"]).copy()
    flow["Week_Sub"] = flow["FirstDateIn"].dt.to_period("W").dt.start_time
    flow["Week_App"] = flow["ApprovedDate"].dt.to_period("W").dt.start_time
    trend = pd.concat(
        [flow.groupby("Week_Sub").size(), flow.groupby("Week_App").size()],
        axis=1, sort=True).fillna(0)
    trend.columns = ["Submitted", "Approved"]
    trend["Backlog"] = (trend["Submitted"] - trend["Approved"]).cumsum()

    df_sla   = df[df["DueDate"].notna()].copy()
    disc_col = "Select List 1" if "Select List 1" in df_sla.columns else None

    # Submission late: supplier delivered after the planned submission date
    df_sla["Sub_Late"] = (
        df_sla["FirstDateIn"].notna() & df_sla["Planned Submission Date"].notna() &
        (df_sla["FirstDateIn"] > df_sla["Planned Submission Date"])
    )
    # Review late: arrived on time from supplier but approved after due date
    # (or still pending past due date) — isolates the reviewer-side delay
    df_sla["Rev_Late"] = (
        ~df_sla["Sub_Late"] & (
            (df_sla["ApprovedDate"].notna() & (df_sla["ApprovedDate"] > df_sla["DueDate"])) |
            (df_sla["ApprovedDate"].isna()  & (df_sla["DueDate"] < today))
        )
    )

    if disc_col:
        risk = df_sla.groupby(disc_col).agg(
            Total    = ("Document No", "count"),
            Sub_Late = ("Sub_Late",    "sum"),
            Rev_Late = ("Rev_Late",    "sum"),
        )
        risk["Sub_Breach_%"] = (risk["Sub_Late"] / risk["Total"] * 100).round(1)
        risk["Rev_Breach_%"] = (risk["Rev_Late"] / risk["Total"] * 100).round(1)
        # Sort by whichever breach rate is higher per discipline
        risk["_max_breach"]  = risk[["Sub_Breach_%","Rev_Breach_%"]].max(axis=1)
        risk = risk.sort_values("_max_breach", ascending=False).drop(columns=["_max_breach"])
    else:
        risk = pd.DataFrame()

    df_rev = df[has_rev].dropna(subset=["Assigned_To"]).copy()
    df_rev["ReviewDays"] = (df_rev["ApprovedDate"] - df_rev["FirstDateIn"]).dt.days

    # Explode multi-person cells: "Alice, Bob, Carol" → one row per person
    # Each person in the cell shares equal credit for the review result.
    df_rev["Assigned_To"] = df_rev["Assigned_To"].astype(str).str.split(r"\s*,\s*")
    df_rev = df_rev.explode("Assigned_To")
    df_rev["Assigned_To"] = df_rev["Assigned_To"].str.strip()
    df_rev = df_rev[df_rev["Assigned_To"] != ""]

    global_weeks = max(
        (df_rev["ApprovedDate"].max() - df_rev["FirstDateIn"].min()).days / 7, 1
    ) if len(df_rev) > 0 else 1
    reviewer_data = []
    for rev in df_rev["Assigned_To"].unique():
        rdocs = df_rev[df_rev["Assigned_To"] == rev]
        reviewer_data.append({
            "Reviewer":      rev,
            "Count":         len(rdocs),
            "Avg_Days":      round(rdocs["ReviewDays"].mean(), 1),
            "OnTime_%":      round(rdocs["ReviewerOnTime"].mean() * 100, 1),
            "Docs_per_Week": round(len(rdocs) / global_weeks, 2),
        })
    review_perf = (
        pd.DataFrame(reviewer_data).sort_values("Docs_per_Week", ascending=False).set_index("Reviewer")
        if reviewer_data else pd.DataFrame())

    dist_labels = ["0-5", "6-10", "11-20", "20+"]
    bucket      = pd.cut(df_rev["ReviewDays"].dropna(), bins=[0,5,10,20,9999],
                         labels=dist_labels, right=True)
    dist_counts = bucket.value_counts().sort_index()
    dist_pct    = (dist_counts / dist_counts.sum() * 100).round(1) if dist_counts.sum() > 0 else dist_counts

    rev_df = df.copy()
    rev_df["Revision"]     = rev_df["_Rev"]
    name_map               = (sp_base[["Document No","SupplierName"]]
                               .drop_duplicates("Document No").set_index("Document No")["SupplierName"])
    rev_df["SupplierName"] = rev_df["Document No"].map(name_map).fillna("Unknown")
    if disc_col and disc_col in df.columns:
        rev_df["Discipline"] = rev_df["Document No"].map(df.set_index("Document No")[disc_col]).fillna("Unknown")
    else:
        rev_df["Discipline"] = "Unknown"

    supp_rev = rev_df.groupby("SupplierName")["Revision"].agg(
        Total_Docs=("count"), Avg_Revision=("mean"), Max_Revision=("max"),
        First_Time_Accepted=(lambda x: int((x == 0).sum())),
    ).reset_index()
    supp_rev["Avg_Revision"] = supp_rev["Avg_Revision"].round(2)
    supp_rev["First_Time_%"] = (supp_rev["First_Time_Accepted"] / supp_rev["Total_Docs"] * 100).round(1)
    supp_rev["Rework_Docs"]  = supp_rev["Total_Docs"] - supp_rev["First_Time_Accepted"]
    supp_rev = supp_rev.sort_values("Avg_Revision", ascending=False).set_index("SupplierName")

    disc_rev = rev_df.groupby("Discipline")["Revision"].agg(
        Total_Docs=("count"), Avg_Revision=("mean"), Max_Revision=("max"),
    ).reset_index()
    disc_rev["Avg_Revision"] = disc_rev["Avg_Revision"].round(2)
    disc_rev = disc_rev.sort_values("Avg_Revision", ascending=False).set_index("Discipline")

    top_cols    = (["Document No","SupplierName","Discipline","Revision","Title"]
                   if "Title" in rev_df.columns else ["Document No","SupplierName","Discipline","Revision"])
    top_revised = rev_df[top_cols].sort_values("Revision", ascending=False).head(10).reset_index(drop=True)

    return {
        "df": df, "valid": valid,
        "supplier_on_time": supplier_on_time,
        "review_on_time":   review_on_time,
        "median_cycle":     median_cycle,
        "backlog":          backlog,
        "supplier_pct":     supplier_pct,
        "review_pct":       review_pct,
        "supplier_perf":    supplier_perf,
        "trend":            trend,
        "risk":             risk,
        "review_perf":      review_perf,
        "dist_counts":      dist_counts,
        "dist_pct":         dist_pct,
        "dist_labels":      dist_labels,
        "outcome_dist":     outcome_dist,
        "supp_rev":         supp_rev,
        "disc_rev":         disc_rev,
        "top_revised":      top_revised,
        "total_revisions":  int(rev_df["Revision"].sum()),
        "avg_revision":     round(rev_df["Revision"].mean(), 2),
        "pct_first_time":   round((rev_df["Revision"] == 0).mean() * 100, 1),
        "max_revision":     int(rev_df["Revision"].max()),
        "rev_df":           rev_df,
    }


def get_findings(R):
    actions = []
    sp = R["supplier_perf"]

    if R["review_pct"] > R["supplier_pct"]:
        actions.append({
            "sev": "red" if R["review_pct"] > 70 else "amber",
            "sev_label": "CRITICAL" if R["review_pct"] > 70 else "ATTENTION",
            "title": "Review Is the Primary Bottleneck",
            "body": (
                f"Internal review accounts for <strong>{R['review_pct']:.1f}%</strong> of total cycle time "
                f"vs {R['supplier_pct']:.1f}% from supplier delay. The constraint is inside the team, not upstream."
            ),
            "items": [
                "Rebalance reviewer workload — check who is carrying the most volume",
                "Introduce a priority queue: oldest pending documents reviewed first",
                "Audit for redundant sign-off steps that add time without adding value",
            ],
        })
    else:
        top = sp[sp["Delayed_Docs"] == sp["Delayed_Docs"].max()].index.tolist()
        nev = sp[sp["OnTime_%"] == 0].index.tolist()
        items = [f"<strong>Highest delayed count:</strong> {', '.join(top)} — escalate for schedule alignment"]
        if nev:
            items.append(f"<strong>Never submitted on time:</strong> {', '.join(nev)} — formal performance review warranted")
        actions.append({
            "sev": "red" if R["supplier_pct"] > 70 else "amber",
            "sev_label": "CRITICAL" if R["supplier_pct"] > 70 else "ATTENTION",
            "title": "Supplier Delays Are Driving Cycle Time",
            "body": (
                f"Supplier-side delays account for <strong>{R['supplier_pct']:.1f}%</strong> of total cycle time. "
                "The review team is not the bottleneck — corrective action must focus upstream."
            ),
            "items": items,
        })

    if R["backlog"] > 20:
        actions.append({
            "sev": "red", "sev_label": "CRITICAL",
            "title": "Backlog Is Critical",
            "body": (
                f"<strong>{R['backlog']:.1f}%</strong> of documents are overdue and still pending approval — "
                "more than double the 10% acceptable limit. This is an active schedule and compliance risk."
            ),
            "items": [
                "Clear session needed immediately — all active reviewers, oldest documents first",
                "Redirect available capacity to the review queue until backlog falls below 15%",
            ],
        })
    elif R["backlog"] > 10:
        actions.append({
            "sev": "amber", "sev_label": "ATTENTION",
            "title": "Backlog Trending Above Threshold",
            "body": (
                f"Backlog is at <strong>{R['backlog']:.1f}%</strong> — above the 10% watch level. "
                "Manageable now, but an upward trend from here escalates quickly."
            ),
            "items": [
                "Monitor weekly; escalate immediately if it approaches 20%",
                "Confirm reviewer capacity can absorb current submission volume",
            ],
        })

    if not R["risk"].empty:
        risk       = R["risk"]
        # Use whichever breach rate is higher as the severity driver
        risk["_max"] = risk[["Sub_Breach_%","Rev_Breach_%"]].max(axis=1)
        top_disc   = risk["_max"].idxmax()
        top_breach = risk["_max"].max()
        total_d    = len(risk)
        crit_d     = risk[risk["_max"] > 50]
        att_d      = risk[(risk["_max"] > 30) & (risk["_max"] <= 50)]
        top5_rows  = "".join(
            f"<li><strong>{d}</strong> — "
            f"Sub: {risk.loc[d,'Sub_Breach_%']:.1f}% &nbsp;|&nbsp; "
            f"Rev: {risk.loc[d,'Rev_Breach_%']:.1f}% "
            f"({int(risk.loc[d,'Total'])} docs)</li>"
            for d in risk.nlargest(5,"_max").index
        )
        if top_breach > 50:
            actions.append({
                "sev": "red", "sev_label": "CRITICAL",
                "title": f"SLA Breach — {len(crit_d)} of {total_d} Disciplines Above 50%",
                "body": (
                    f"Highest breach: <strong>{top_disc}</strong> at <strong>{top_breach:.1f}%</strong>. "
                    "This is not an isolated incident — it reflects structural resourcing or workflow issues."
                ),
                "top5_html": top5_rows,
                "items": [
                    f"Targeted audit of pending documents in {top_disc} and other critical disciplines",
                    "Engage discipline leads to identify root cause: under-resourcing, complexity, or process gaps",
                ],
            })
        elif top_breach > 30:
            actions.append({
                "sev": "amber", "sev_label": "ATTENTION",
                "title": f"SLA Breach — {len(att_d)} Discipline(s) Above 30%",
                "body": (
                    f"Highest breach: <strong>{top_disc}</strong> at <strong>{top_breach:.1f}%</strong>. "
                    "Early intervention prevents this from becoming critical next cycle."
                ),
                "top5_html": top5_rows,
                "items": [
                    "Increase review monitoring for high-breach disciplines",
                    "Identify whether delays cluster around specific reviewers or document types",
                ],
            })

    if not R["review_perf"].empty:
        rp    = R["review_perf"]
        med   = rp["Docs_per_Week"].median()
        low_t = rp[rp["Docs_per_Week"] < med * 0.5].index.tolist()
        bsla  = rp[rp["OnTime_%"] < 80].index.tolist()
        if low_t or bsla:
            def chunks(lst, n=4):
                return [lst[i:i+n] for i in range(0, len(lst), n)]
            # Critical if: ≥30% of team is below 80% SLA, OR any single reviewer fails both
            dual      = [n for n in low_t if n in bsla]
            pct_below = len(bsla) / len(rp) * 100
            sev       = "red"      if (dual or pct_below >= 30) else "amber"
            sev_label = "CRITICAL" if (dual or pct_below >= 30) else "ATTENTION"
            parts = []
            if bsla:
                parts.append(
                    f"<strong>{len(bsla)} of {len(rp)} reviewers ({pct_below:.0f}%)</strong> "
                    f"have an on-time rate below the 80% SLA threshold."
                )
            if low_t:
                parts.append(
                    f"<strong>{len(low_t)} reviewer(s)</strong> are also below 50% of team median throughput "
                    f"({med:.2f} docs/week)."
                )
            if dual:
                parts.append(
                    f"<strong>{len(dual)} reviewer(s) fail both criteria</strong> — highest priority for support or rebalancing."
                )
            actions.append({
                "sev": sev, "sev_label": sev_label,
                "title": "Reviewer Performance Gaps",
                "body": " ".join(parts),
                "reviewer_section": {
                    "low_throughput": {"chunks": chunks(low_t), "count": len(low_t)},
                    "below_sla":      {"chunks": chunks(bsla),  "count": len(bsla)},
                },
                "items": [
                    "Check whether affected reviewers need support or carry a disproportionate complexity load",
                    "Consider temporary workload rebalancing until performance recovers",
                ],
            })

    avg_rev  = R["avg_revision"];  pct_ft = R["pct_first_time"]
    max_rev  = R["max_revision"];  sr     = R["supp_rev"]
    worst_s  = sr.index[0]       if len(sr) > 0 else "—"
    worst_sa = sr.iloc[0]["Avg_Revision"] if len(sr) > 0 else 0
    if avg_rev > 1.5 or pct_ft < 50:
        sev = "red" if (avg_rev > 3 or pct_ft < 30) else "amber"
        actions.append({
            "sev": sev, "sev_label": "CRITICAL" if sev == "red" else "ATTENTION",
            "title": "Revision Rate Is Consuming Review Capacity",
            "body": (
                f"Only <strong>{pct_ft:.1f}%</strong> of documents accepted on first submission "
                f"(avg {avg_rev:.2f} revisions/doc). "
                f"<strong>{worst_s}</strong> has the highest rework rate at {worst_sa:.2f} revisions/doc. "
                "Every revision cycle is a full review pass that could have been avoided."
            ),
            "items": [
                f"Investigate root cause for {worst_s}: unclear scope, insufficient QA, or ambiguous review criteria",
                "Issue a submission quality advisory to all suppliers averaging above 2 revisions",
                f"The most-revised document has reached revision {max_rev} — assign a responsible reviewer to close it",
            ],
        })
    elif avg_rev <= 0.5 and pct_ft >= 80:
        actions.append({
            "sev": "green", "sev_label": "HEALTHY",
            "title": "Submission Quality Is Good",
            "body": (
                f"<strong>{pct_ft:.1f}%</strong> of documents accepted on first submission "
                f"(avg {avg_rev:.2f} revisions/doc). No systemic rework burden detected."
            ),
            "items": ["Maintain current submission quality standards and benchmarks"],
        })

    if R["backlog"] <= 10:
        actions.append({
            "sev": "green", "sev_label": "HEALTHY",
            "title": "Backlog Is Under Control",
            "body": f"Current backlog: <strong>{R['backlog']:.1f}%</strong> — within the ≤10% target.",
            "items": ["Maintain current capacity; flag any week-on-week increase above 3–5 percentage points"],
        })
    if R["risk"].empty or R["risk"][["Sub_Breach_%","Rev_Breach_%"]].max(axis=1).max() <= 30:
        actions.append({
            "sev": "green", "sev_label": "HEALTHY",
            "title": "Discipline SLA Performance Is Acceptable",
            "body": "All disciplines operating below the 30% SLA breach threshold. Record as audit baseline.",
            "items": ["Document current breach rates as the baseline for the next reporting cycle"],
        })

    return actions


def render_finding(action):
    sev  = action["sev"]
    html = (
        f'<div class="fc {sev}">'
        f'<span class="fc-badge {sev}">{action["sev_label"]}</span>'
        f'<h4>{action["title"]}</h4>'
        f'<p class="fc-body">{action["body"]}</p>'
    )
    if "top5_html" in action:
        html += ('<p style="font-weight:700;font-size:0.82rem;margin:0.4rem 0 0.2rem 0;">'
                 'Disciplines by breach rate:</p>'
                 f'<ol style="margin:0 0 0.45rem 1.15rem;font-size:0.84rem;line-height:1.75;">'
                 f'{action["top5_html"]}</ol>')
    if "reviewer_section" in action:
        rs = action["reviewer_section"]; lt = rs["low_throughput"]; bs = rs["below_sla"]
        html += '<ul class="fc-list">'
        if lt["count"] > 0:
            html += f'<li><strong>Low throughput — {lt["count"]} reviewer(s):</strong><ul class="rnames">'
            for chunk in lt["chunks"]: html += f"<li>{', '.join(chunk)}</li>"
            html += '</ul></li>'
        if bs["count"] > 0:
            html += f'<li><strong>Below 80% on-time — {bs["count"]} reviewer(s):</strong><ul class="rnames">'
            for chunk in bs["chunks"]: html += f"<li>{', '.join(chunk)}</li>"
            html += '</ul></li>'
        for item in action.get("items", []): html += f"<li>{item}</li>"
        html += '</ul>'
    elif "items" in action:
        html += '<ul class="fc-list">'
        for item in action["items"]: html += f"<li>{item}</li>"
        html += '</ul>'
    html += '</div>'
    return html


def chart_time_split(supplier_pct, review_pct):
    fig, ax = plt.subplots(figsize=(6, 5))
    fig.patch.set_facecolor("white")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    labels = ["Supplier\nDelay", "Review\nDuration"]
    vals   = [supplier_pct, review_pct]
    colors = [C_NAVY, C_AMBER]
    bars   = ax.bar(labels, vals, color=colors, width=0.42)
    ax.axhline(50, color="#bbb", linestyle="--", linewidth=1.0, label="50% balance line")
    ax.set_ylabel("% of Total Cycle Time")
    ax.set_ylim(0, 115)
    ax.set_title("Where Is Time Being Lost?", pad=10)
    ax.grid(axis="y", linestyle="--", alpha=0.5)
    ax.grid(axis="x", visible=False)
    ax.legend(fontsize=8.5, loc="upper right")
    for bar, val in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width() / 2, val + 2,
                f"{val:.1f}%", ha="center", va="bottom", fontweight="bold", fontsize=12)
    dominant = "Review" if review_pct > supplier_pct else "Supplier"
    fig.text(0.5, -0.02,
             f"Primary driver: {dominant} phase ({max(supplier_pct, review_pct):.1f}% of total cycle time)",
             ha="center", fontsize=9, color=C_GREY, style="italic")
    fig.tight_layout()
    return fig


def chart_flow_trend(trend):
    fig, ax = _fig(11, 4.5)
    ax.plot(trend.index, trend["Submitted"], marker="o", label="Submitted",
            color=C_NAVY, linewidth=2, markersize=5)
    ax.plot(trend.index, trend["Approved"],  marker="s", label="Approved",
            color=C_GREEN, linewidth=2, markersize=5)
    ax.plot(trend.index, trend["Backlog"],   marker="d", label="Cumulative Backlog",
            color=C_RED, linewidth=2, markersize=5, linestyle="--")
    ax.fill_between(trend.index, trend["Backlog"], alpha=0.08, color=C_RED)
    ax.set_xlabel("Week"); ax.set_ylabel("Documents")
    ax.set_title("Submission vs. Approval — Weekly Flow")
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.5)
    return fig


def chart_revision_dist(rev_df):
    rev_bins = rev_df["Revision"].value_counts().sort_index()
    rev_bins = rev_bins[rev_bins.index <= 10]
    fig, ax  = _fig(8, 3.8)
    clrs = [C_GREEN if i == 0 else C_AMBER if i <= 2 else C_RED for i in rev_bins.index]
    ax.bar(rev_bins.index.astype(str), rev_bins.values, color=clrs, width=0.52)
    ax.set_xlabel("Revision Number"); ax.set_ylabel("Documents")
    ax.set_title("Revision Distribution Across the Register")
    ax.legend(handles=[
        mpatches.Patch(color=C_GREEN, label="Rev 0 — accepted first time"),
        mpatches.Patch(color=C_AMBER, label="Rev 1–2 — moderate rework"),
        mpatches.Patch(color=C_RED,   label="Rev 3+ — significant rework"),
    ], fontsize=8.5)
    for i, (idx, val) in enumerate(zip(rev_bins.index, rev_bins.values)):
        ax.text(i, val + rev_bins.max() * 0.02, str(val), ha="center", fontsize=9)
    return fig


def main():
    with st.sidebar:
        st.markdown("### Upload Data")
        supplier_file = st.file_uploader("Supplier Documents (.xlsx)", type=["xlsx"])
        workflow_file = st.file_uploader("Workflow Export (.xlsx)",    type=["xlsx"])
        st.markdown("---")
        st.markdown("### Report Details")
        report_date    = st.date_input("Report Date",    value=date.today())
        project_name   = st.text_input("Project Name",   value="")
        project_number = st.text_input("Project Number", value="")
        prepared_by    = st.text_input("Prepared By",    value="")
        show_supplier   = False; show_reviewers = False; show_discipline = False
        show_revision   = False; show_trend     = False; show_outcomes   = False
        if supplier_file and workflow_file:
            st.markdown("---")
            st.markdown("### Supporting Detail")
            st.markdown("<small style='color:#888;'>Show additional tables and charts below the findings.</small>", unsafe_allow_html=True)
            show_supplier   = st.checkbox("Supplier performance table",   value=True)
            show_reviewers  = st.checkbox("Reviewer performance table",   value=True)
            show_discipline = st.checkbox("Discipline risk table",        value=True)
            show_trend      = st.checkbox("Flow trend chart",             value=True)
            show_revision   = st.checkbox("Revision detail tables",       value=False)
            show_outcomes   = st.checkbox("Review outcome distribution",  value=False)

    title_col, score_col = st.columns([3, 1])
    with title_col:
        st.markdown(
            '<p class="report-title">Document Control Audit Report</p>'
            '<p class="report-sub">Performance Analysis — Document Intelligence Layer</p>',
            unsafe_allow_html=True)

    def _v(v): return v if v else '<em style="color:#aaa;">—</em>'
    st.markdown(f"""
    <div class="meta-block"><table>
        <tr>
            <td>Date</td><td>{report_date.strftime("%d %B %Y")}</td>
            <td style="width:30px;"></td>
            <td>Project Name</td><td>{_v(project_name)}</td>
        </tr><tr>
            <td>Project Number</td><td>{_v(project_number)}</td>
            <td></td>
            <td>Prepared By</td><td>{_v(prepared_by)}</td>
        </tr>
    </table></div>""", unsafe_allow_html=True)

    if not (supplier_file and workflow_file):
        st.markdown(f"""
        <div class="welcome">
            <h3 style="margin:0 0 0.3rem 0;">Document Control Audit — Report Plan</h3>
            <p style="font-size:0.83rem;color:{C_GREY};margin:0 0 1rem 0;line-height:1.6;">
            Structured around one question: <strong>where is time being lost, and what must be done about it?</strong>
            Designed for a monthly or bi-monthly audit review meeting — findings up front, evidence on demand.
            </p>
            <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
            <tr>
                <td style="padding:0.55rem 0.9rem 0.55rem 0;vertical-align:top;width:24px;color:{C_BLUE};font-weight:800;">①</td>
                <td style="padding:0.55rem 0 0.55rem 0;vertical-align:top;border-bottom:1px solid {C_BDR};">
                    <strong>Key Performance Indicators</strong><br>
                    <span style="color:{C_GREY};font-size:0.81rem;">Four headline metrics — supplier on-time rate, review SLA compliance,
                    median cycle time, and backlog pressure — colour-coded against fixed thresholds.</span>
                </td>
            </tr>
            <tr>
                <td style="padding:0.55rem 0.9rem 0.55rem 0;vertical-align:top;color:{C_BLUE};font-weight:800;">②</td>
                <td style="padding:0.55rem 0 0.55rem 0;vertical-align:top;border-bottom:1px solid {C_BDR};">
                    <strong>Revision Cycle Snapshot</strong><br>
                    <span style="color:{C_GREY};font-size:0.81rem;">Compact view of rework burden — first-time acceptance rate,
                    average revisions per document, and the supplier driving the most revision cycles.</span>
                </td>
            </tr>
            <tr>
                <td style="padding:0.55rem 0.9rem 0.55rem 0;vertical-align:top;color:{C_BLUE};font-weight:800;">③</td>
                <td style="padding:0.55rem 0 0.55rem 0;vertical-align:top;border-bottom:1px solid {C_BDR};">
                    <strong>Audit Findings &amp; Recommended Actions</strong><br>
                    <span style="color:{C_GREY};font-size:0.81rem;">Every finding derived from the data, labelled Critical / Attention / Healthy,
                    and paired with specific corrective actions. This is the section the meeting runs on.</span>
                </td>
            </tr>
            <tr>
                <td style="padding:0.55rem 0.9rem 0 0;vertical-align:top;color:{C_BLUE};font-weight:800;">④</td>
                <td style="padding:0.55rem 0 0 0;vertical-align:top;">
                    <strong>Supporting Detail</strong> <span style="color:{C_GREY};font-size:0.79rem;">(sidebar toggle)</span><br>
                    <span style="color:{C_GREY};font-size:0.81rem;">Supplier table, reviewer table, discipline SLA rates,
                    flow trend, and revision detail — shown only when needed.</span>
                </td>
            </tr>
            </table>
            <p style="font-size:0.79rem;color:{C_GREY};margin:1rem 0 0 0;border-top:1px solid {C_BDR};padding-top:0.7rem;">
            ← Upload both Excel exports via the sidebar to generate the report.
            </p>
        </div>""", unsafe_allow_html=True)
        return

    progress = st.progress(0, text="Reading supplier documents…")
    sdocs = clean_header(pd.read_excel(supplier_file))
    progress.progress(30, text="Reading workflow data…")
    wflow = clean_header(pd.read_excel(workflow_file))
    progress.progress(60, text="Running analysis…")
    R = process_data(sdocs, wflow)
    progress.progress(100, text="Done.")
    progress.empty()

    hs = compute_health_score(R["supplier_on_time"], R["review_on_time"], R["backlog"], R["risk"])
    if hs >= 80:   hc, ht, hbg = C_GREEN, "Good",      "#e8f8ee"
    elif hs >= 60: hc, ht, hbg = C_AMBER, "Attention", "#fef8e5"
    else:          hc, ht, hbg = C_RED,   "Critical",  "#fdf0f0"
    with score_col:
        st.markdown(f"""
        <div class="hs-wrap" style="background:{hbg};">
            <p class="hs-lbl">Process Health Score</p>
            <p class="hs-score" style="color:{hc};">{hs}<span class="hs-denom"> /100</span></p>
            <p class="hs-status" style="color:{hc};">{ht}</p>
        </div>""", unsafe_allow_html=True)

    # KPIs
    sec("Key Performance Indicators")
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.markdown(kpi_tile(f"{R['supplier_on_time']:.1f}%", "Supplier On-Time",     kpi_color(R['supplier_on_time'])), unsafe_allow_html=True)
    with k2: st.markdown(kpi_tile(f"{R['review_on_time']:.1f}%",   "Review On-Time (SLA)", kpi_color(R['review_on_time'])),   unsafe_allow_html=True)
    with k3: st.markdown(kpi_tile(f"{R['median_cycle']:.0f} days", "Median Cycle Time",    C_BLUE),                           unsafe_allow_html=True)
    with k4:
        bc = C_RED if R['backlog'] > 20 else C_AMBER if R['backlog'] > 10 else C_GREEN
        st.markdown(kpi_tile(f"{R['backlog']:.1f}%", "Backlog Pressure", bc), unsafe_allow_html=True)
    ibox("Target ≥80% for all rate metrics. &nbsp;"
         "<strong>Backlog Pressure</strong>: % of documents past due date still awaiting approval — "
         "target &lt;10%; ≥20% is critical. &nbsp;"
         "<strong>Median Cycle Time</strong>: midpoint of days from planned submission to final approval.")

    # Supporting detail (sidebar-controlled)
    any_detail = any([show_supplier, show_reviewers, show_discipline,
                      show_revision, show_trend, show_outcomes])
    if any_detail:
        sec("Supporting Detail")

    if show_trend or show_supplier:
        if show_trend and show_supplier:
            col1, gap, col2 = st.columns([5, 1, 9])
            with col1:
                st.markdown("**Where is time being lost?**")
                fig = chart_time_split(R["supplier_pct"], R["review_pct"])
                st.pyplot(fig); plt.close(fig)
            with col2:
                st.markdown("**Weekly submission vs. approval flow**")
                fig = chart_flow_trend(R["trend"])
                st.pyplot(fig); plt.close(fig)
        elif show_supplier:
            _, col_c, _ = st.columns([1, 3, 1])
            with col_c:
                st.markdown("**Where is time being lost?**")
                fig = chart_time_split(R["supplier_pct"], R["review_pct"])
                st.pyplot(fig); plt.close(fig)
        elif show_trend:
            st.markdown("**Weekly submission vs. approval flow**")
            fig = chart_flow_trend(R["trend"])
            st.pyplot(fig); plt.close(fig)

    if show_supplier:
        with st.expander("Supplier Submission Performance", expanded=True):
            st.caption("`OnTime_%` target ≥80%. `Avg_Delay` = mean days late for delayed docs only. Sorted by delayed count descending.")
            disp = R["supplier_perf"][["Total_Docs","OnTime_Docs","Delayed_Docs","OnTime_%","Avg_Delay","Max_Delay"]].copy()
            st.dataframe(
                disp.style
                .format({"OnTime_%":"{:.1f}%","Avg_Delay":"{:.1f}","Max_Delay":"{:.0f}"})
                .background_gradient(subset=["Delayed_Docs"], cmap="Reds")
                .background_gradient(subset=["OnTime_%"],     cmap="RdYlGn")
                .background_gradient(subset=["Avg_Delay"],    cmap="Oranges")
                .set_properties(**{"text-align":"center","min-width":"100px"}),
                use_container_width=True)
            mx  = R["supplier_perf"]["Delayed_Docs"].max()
            top = R["supplier_perf"][R["supplier_perf"]["Delayed_Docs"] == mx].index.tolist()
            nev = R["supplier_perf"][R["supplier_perf"]["OnTime_%"] == 0].index.tolist()
            nev_l = (f"Suppliers with 0% on-time: <strong>{', '.join(nev)}</strong>."
                     if nev else "All suppliers have at least one on-time submission.")
            tnote(f"Highest delayed count ({int(mx)} docs): <strong>{', '.join(top)}</strong>. &nbsp; {nev_l}")

    if show_discipline and not R["risk"].empty:
        with st.expander("Discipline SLA Breach Rates", expanded=True):
            st.caption(
                "**Submission Late**: supplier delivered after the planned submission date. "
                "**Review Late**: arrived on time but approved after the due date (or still pending). "
                "The two breach rates show *where* the delay sits — supplier side or review side."
            )
            risk_disp = R["risk"][["Total","Sub_Late","Sub_Breach_%","Rev_Late","Rev_Breach_%"]].copy()
            risk_disp.columns = ["Total","Submission Late","Sub Breach %","Review Late","Rev Breach %"]
            st.dataframe(
                risk_disp.style
                .format({"Sub Breach %":"{:.1f}%","Rev Breach %":"{:.1f}%",
                         "Total":"{:.0f}","Submission Late":"{:.0f}","Review Late":"{:.0f}"})
                .background_gradient(subset=["Sub Breach %"], cmap="Oranges")
                .background_gradient(subset=["Rev Breach %"], cmap="Reds")
                .set_properties(**{"text-align":"center","min-width":"110px"}),
                use_container_width=True)
            top_sub = R["risk"]["Sub_Breach_%"].idxmax() if not R["risk"].empty else "—"
            top_rev = R["risk"]["Rev_Breach_%"].idxmax() if not R["risk"].empty else "—"
            top_sub_v = R["risk"]["Sub_Breach_%"].max()
            top_rev_v = R["risk"]["Rev_Breach_%"].max()
            tnote(
                f"Sorted by highest breach rate (submission or review) descending.<br>"
                f"— Highest submission breach: <strong>{top_sub}</strong> ({top_sub_v:.1f}%) &nbsp;|&nbsp; "
                f"Highest review breach: <strong>{top_rev}</strong> ({top_rev_v:.1f}%)"
            )

    if show_reviewers and not R["review_perf"].empty:
        with st.expander("Reviewer Performance", expanded=True):
            st.caption("`OnTime_%` target ≥80%. `Docs_per_Week` uses project-wide window — comparable across all reviewers.")
            st.dataframe(
                R["review_perf"][["Count","Docs_per_Week","Avg_Days","OnTime_%"]].style
                .format({"Docs_per_Week":"{:.2f}","Avg_Days":"{:.1f}","OnTime_%":"{:.1f}%"})
                .background_gradient(cmap="RdYlGn", subset=["OnTime_%"])
                .background_gradient(cmap="Blues",  subset=["Docs_per_Week"])
                .set_properties(**{"text-align":"center","min-width":"100px"}),
                use_container_width=True)
            b80 = int((R["review_perf"]["OnTime_%"] < 80).sum())
            avg_ot = R["review_perf"]["OnTime_%"].mean()
            tnote(f"Team avg on-time: <strong>{avg_ot:.1f}%</strong>. &nbsp; "
                  f"<span style='color:{C_RED if b80>0 else C_GREEN};font-weight:700;'>"
                  f"{b80} reviewer(s) below 80% on-time threshold.</span>")

    if show_revision:
        with st.expander("Revision Cycle Detail", expanded=False):
            rv_c = C_RED if R["avg_revision"]>2 else C_AMBER if R["avg_revision"]>1 else C_GREEN
            ft_c = C_GREEN if R["pct_first_time"]>=80 else C_AMBER if R["pct_first_time"]>=50 else C_RED
            mx_c = C_RED if R["max_revision"]>5 else C_AMBER if R["max_revision"]>2 else C_GREEN
            r1,r2,r3,r4 = st.columns(4)
            with r1: st.markdown(kpi_tile(f"{R['avg_revision']:.2f}", "Avg Revisions / Doc", rv_c), unsafe_allow_html=True)
            with r2: st.markdown(kpi_tile(f"{R['pct_first_time']:.1f}%", "First-Time Acceptance", ft_c), unsafe_allow_html=True)
            with r3: st.markdown(kpi_tile(str(R["max_revision"]), "Highest Single Document", mx_c), unsafe_allow_html=True)
            with r4: st.markdown(kpi_tile(str(R["total_revisions"]), "Total Revision Cycles", C_GREY), unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            fig = chart_revision_dist(R["rev_df"]); st.pyplot(fig); plt.close(fig)
            ibox(
                f"Of {len(R['rev_df'])} documents, <strong>{int((R['rev_df']['Revision']==0).sum())}</strong> "
                f"({R['pct_first_time']:.1f}%) were accepted on first submission. "
                f"The remaining <strong>{int((R['rev_df']['Revision']>0).sum())}</strong> generated "
                f"<strong>{R['total_revisions']}</strong> additional review cycles.")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**By Supplier**")
                st.dataframe(
                    R["supp_rev"][["Total_Docs","First_Time_Accepted","Rework_Docs","First_Time_%","Avg_Revision","Max_Revision"]].style
                    .format({"First_Time_%":"{:.1f}%","Avg_Revision":"{:.2f}","Max_Revision":"{:.0f}"})
                    .background_gradient(subset=["Avg_Revision"], cmap="Reds")
                    .background_gradient(subset=["First_Time_%"], cmap="RdYlGn")
                    .set_properties(**{"text-align":"center","min-width":"90px"}),
                    use_container_width=True)
            with col2:
                st.markdown("**By Discipline**")
                st.dataframe(
                    R["disc_rev"].style
                    .format({"Avg_Revision":"{:.2f}","Max_Revision":"{:.0f}"})
                    .background_gradient(subset=["Avg_Revision"], cmap="Reds")
                    .set_properties(**{"text-align":"center","min-width":"90px"}),
                    use_container_width=True)
            st.markdown("**Top 10 Most-Revised Documents**")
            st.dataframe(
                R["top_revised"].style.background_gradient(subset=["Revision"], cmap="Reds")
                .set_properties(**{"text-align":"left"}), use_container_width=True)

    if show_outcomes and not R["outcome_dist"].empty:
        with st.expander("Review Outcome Distribution (C-Codes)", expanded=False):
            oc = R["outcome_dist"]; tot_oc = oc.sum()
            clr_map = {"C1 — Approved":C_GREEN,"C2 — Approved with Comments":C_AMBER,
                       "C3 — Rejected":C_RED,"C4 — For Information":C_BLUE,"Other":C_GREY}
            col1, col2 = st.columns([2, 1])
            with col1:
                fig, ax = _fig(8, max(3, len(oc)*0.6))
                colors  = [clr_map.get(k, C_GREY) for k in oc.index[::-1]]
                ax.barh(oc.index[::-1], oc.values[::-1], color=colors, height=0.48)
                ax.set_xlabel("Count"); ax.set_title("Review Outcome Distribution")
                ax.grid(axis="x", linestyle="--", alpha=0.6); ax.grid(axis="y", visible=False)
                for i,(k,v) in enumerate(zip(oc.index[::-1], oc.values[::-1])):
                    ax.text(v+oc.max()*0.01, i, str(int(v)), va="center", fontsize=9, fontweight="bold")
                ax.set_xlim(0, oc.max()*1.14)
                st.pyplot(fig); plt.close(fig)
            with col2:
                def _pct(n): return f"{n/tot_oc*100:.1f}%" if tot_oc>0 else "—"
                c1v=oc.get("C1 — Approved",0); c2v=oc.get("C2 — Approved with Comments",0)
                c3v=oc.get("C3 — Rejected",0); c4v=oc.get("C4 — For Information",0)
                st.markdown(f"""
                <div class="sc"><h5>Outcome Breakdown</h5>
                <div class="mrow"><span style="color:{C_GREEN};">C1 Approved</span>
                    <span class="mval">{int(c1v)} ({_pct(c1v)})</span></div>
                <div class="mrow"><span style="color:{C_AMBER};">C2 w/ Comments</span>
                    <span class="mval">{int(c2v)} ({_pct(c2v)})</span></div>
                <div class="mrow"><span style="color:{C_RED};">C3 Rejected</span>
                    <span class="mval">{int(c3v)} ({_pct(c3v)})</span></div>
                <div class="mrow"><span style="color:{C_BLUE};">C4 For Info</span>
                    <span class="mval">{int(c4v)} ({_pct(c4v)})</span></div>
                </div>""", unsafe_allow_html=True)

    # ── Revision Cycle Snapshot ──────────────────────────────────────────────
    sec("Revision Cycle Snapshot")
    sr       = R["supp_rev"]
    worst_s  = sr.index[0]               if len(sr) > 0 else "—"
    worst_sa = sr.iloc[0]["Avg_Revision"] if len(sr) > 0 else 0
    rev0     = int((R["rev_df"]["Revision"] == 0).sum())
    rev_pos  = int((R["rev_df"]["Revision"] >  0).sum())
    rv_c  = C_RED   if R["avg_revision"]   > 2  else C_AMBER if R["avg_revision"]   > 1  else C_GREEN
    ft_c  = C_GREEN if R["pct_first_time"] >= 80 else C_AMBER if R["pct_first_time"] >= 50 else C_RED
    mx_c  = C_RED   if R["max_revision"]   > 5  else C_AMBER if R["max_revision"]   > 2  else C_GREEN
    st.markdown(f"""
    <div style="background:white;border:1px solid {C_BDR};border-radius:5px;padding:1rem 1.3rem;line-height:1.9;font-size:0.88rem;">
        <table style="width:100%;border-collapse:collapse;">
        <tr>
            <td style="width:50%;padding:0.15rem 1.5rem 0.15rem 0;vertical-align:top;border-right:1px solid {C_BDR};">
                <span style="color:{C_GREY};font-size:0.75rem;text-transform:uppercase;letter-spacing:0.06em;margin:0.5rem;">Avg Revisions / Document</span><br>
                <span style="font-size:1.45rem;font-weight:800;color:{rv_c};margin:0.5rem;">{R["avg_revision"]:.2f}</span>
                &nbsp;<span style="font-size:0.8rem;color:{C_GREY};">revisions per doc</span>
            </td>
            <td style="width:50%;padding:0.15rem 0 0.15rem 1.5rem;vertical-align:top;">
                <span style="color:{C_GREY};font-size:0.75rem;text-transform:uppercase;letter-spacing:0.06em;">First-Time Acceptance Rate</span><br>
                <span style="font-size:1.45rem;font-weight:800;color:{ft_c};">{R["pct_first_time"]:.1f}%</span>
                &nbsp;<span style="font-size:0.8rem;color:{C_GREY};">accepted without resubmission</span>
            </td>
        </tr>
        <tr>
            <td style="padding:0.15rem 1.5rem 0.15rem 0;vertical-align:top;border-right:1px solid {C_BDR};border-top:1px solid {C_BDR};">
                <span style="color:{C_GREY};font-size:0.75rem;text-transform:uppercase;letter-spacing:0.06em;margin:0.5rem;">Highest Single-Document Revision</span><br>
                <span style="font-size:1.45rem;font-weight:800;color:{mx_c};margin:0.5rem;">{R["max_revision"]}</span>
                &nbsp;<span style="font-size:0.8rem;color:{C_GREY};">resubmissions on one document</span>
            </td>
            <td style="padding:0.15rem 0 0.15rem 1.5rem;vertical-align:top;border-top:1px solid {C_BDR};">
                <span style="color:{C_GREY};font-size:0.75rem;text-transform:uppercase;letter-spacing:0.06em;">Total Revision Cycles — Register</span><br>
                <span style="font-size:1.45rem;font-weight:800;color:{C_GREY};">{R["total_revisions"]}</span>
                &nbsp;<span style="font-size:0.8rem;color:{C_GREY};">additional review passes from rework</span>
            </td>
        </tr>
        </table>
    </div>""", unsafe_allow_html=True)
    ibox(
        f"Of <strong>{len(R['rev_df'])}</strong> documents, <strong>{rev0}</strong> ({R['pct_first_time']:.1f}%) "
        f"were accepted on first submission. The remaining <strong>{rev_pos}</strong> generated "
        f"<strong>{R['total_revisions']}</strong> additional review cycles — capacity consumed by rework "
        f"rather than new documents. Highest rework supplier: <strong>{worst_s}</strong> "
        f"({worst_sa:.2f} avg revisions/doc). "
        f"See <em>Revision Detail</em> in Supporting Detail for the full breakdown."
    )

    # ── Audit Findings ────────────────────────────────────────────────────────
    sec("Audit Findings & Recommended Actions")
    st.markdown(
        "<small style='color:#666;'>"
        "<strong>Critical</strong> — immediate action required. &nbsp;"
        "<strong>Attention</strong> — proactive intervention needed before it escalates. &nbsp;"
        "<strong>Healthy</strong> — acceptable performance, retain as audit evidence."
        "</small>", unsafe_allow_html=True)
    for finding in get_findings(R):
        st.markdown(render_finding(finding), unsafe_allow_html=True)

    st.markdown(f"""
    <div class="scope"><strong>Scope &amp; Limitations</strong> —
    SLA on-time metrics are from <em>completed</em> workflow steps only; in-progress documents
    are excluded from on-time calculations but included in backlog and breach figures.
    Revision analysis uses the revision number in the supplier documents export.
    Findings indicate where to focus investigation and should be read alongside project context.
    </div>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()