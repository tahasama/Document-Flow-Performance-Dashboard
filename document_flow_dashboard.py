import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import base64

# Set page configuration
st.set_page_config(
    page_title="Document Flow Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .chart-container {
        background-color: white;
        padding: 1.5rem;
        border-radius: 0.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
    }
    .dataframe {
        font-size: 0.9rem;
    }
    .section-header {
        color: #1f77b4;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #1f77b4;
    }
    .explanation-box {
        background-color: #e8f4f8;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
        margin-bottom: 1rem;
    }
    .action-box {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ffc107;
        margin-bottom: 1rem;
    }
    .warning-box {
        background-color: #f8d7da;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #dc3545;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
        margin-bottom: 1rem;
    }
    .stImage {
        border-radius: 0.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .performer-badge {
        background-color: #f8f9fa;
        padding: 0.5rem;
        border-radius: 0.3rem;
        margin-bottom: 0.3rem;
        border-left: 3px solid #28a745;
    }
    .duration-summary {
        width: 100%;
    }
    .insight-text {
        margin: 0.5rem 0;
        font-size: 0.95rem;
    }
    /* Table styling to match Jupyter */
    .dataframe table {
        width: 100%;
        font-size: 11pt;
    }
    .dataframe th {
        font-size: 12pt;
        padding: 8px;
    }
    .dataframe td {
        font-size: 11pt;
        padding: 8px;
    }
    .dataframe th.row_heading {
        min-width: 220px;
        text-align: left;
    }
</style>
""", unsafe_allow_html=True)

# Helper functions
def clean_header(df):
    """Clean Excel headers by removing first 7 rows and setting proper columns"""
    df = df.iloc[7:].copy()
    df.columns = df.iloc[0]
    df = df[1:]
    df.columns = df.columns.str.strip()
    df.reset_index(drop=True, inplace=True)
    return df

def process_data(supplier_docs, workflow):
    """Process and merge the dataframes"""
    # Rename columns for consistency
    supplier_docs = supplier_docs.rename(columns={
        "Document No": "Document No",
        "Planned Submission Date": "Planned Submission Date"
    })
    
    workflow = workflow.rename(columns={
        "Document No.": "Document No",
        "Assigned To": "Assigned_To",
        "Date In": "Date In",
        "Date Completed": "Date Completed",
        "Date Due": "Date Due",
        "Step Status": "Step Status"
    })
    
    # Convert to string for consistent merging
    supplier_docs["Document No"] = supplier_docs["Document No"].astype(str)
    workflow["Document No"] = workflow["Document No"].astype(str)
    
    # Convert date columns
    supplier_docs["Planned Submission Date"] = pd.to_datetime(
        supplier_docs["Planned Submission Date"], errors="coerce"
    )
    
    workflow["Date In"] = pd.to_datetime(workflow["Date In"], errors="coerce")
    workflow["Date Completed"] = pd.to_datetime(workflow["Date Completed"], errors="coerce")
    workflow["Date Due"] = pd.to_datetime(workflow["Date Due"], errors="coerce")
    
    # Keep only COMPLETED reviews for workflow aggregation
    workflow_completed = workflow[workflow.get("Step Status", "Completed") == "Completed"].copy()
    
    # Aggregate workflow data from COMPLETED reviews only
    agg = workflow_completed.groupby("Document No").agg({
        "Date In": "min",
        "Date Completed": "max",
        "Date Due": "max",
        "Assigned_To": "first"
    }).reset_index()
    
    agg = agg.rename(columns={
        "Date In": "FirstDateIn",
        "Date Completed": "ApprovedDate",
        "Date Due": "DueDate"
    })
    
    # Merge dataframes
    df = supplier_docs.merge(agg, on="Document No", how="left")
    
    # Calculate KPIs
    df["SupplierOnTime"] = df["FirstDateIn"] <= df["Planned Submission Date"]
    supplier_on_time = df["SupplierOnTime"].mean() * 100
    
    df["ReviewerOnTime"] = df["ApprovedDate"] <= df["DueDate"]
    review_on_time = df["ReviewerOnTime"].mean() * 100
    
    df["CycleTime"] = (df["ApprovedDate"] - df["Planned Submission Date"]).dt.days
    median_cycle = df["CycleTime"].median()
    
    df["Overdue"] = (df["DueDate"] < pd.Timestamp.today()) & (
        df["ApprovedDate"].isna() | (df["ApprovedDate"] > df["DueDate"])
    )
    backlog = df["Overdue"].mean() * 100
    
    # Calculate time contributions
    valid = df.dropna(subset=[
        "Planned Submission Date",
        "FirstDateIn",
        "ApprovedDate"
    ]).copy()
    
    valid["SupplierDelay"] = (
        valid["FirstDateIn"] - valid["Planned Submission Date"]
    ).dt.days.clip(lower=0)
    
    valid["ReviewDuration"] = (
        valid["ApprovedDate"] - valid["FirstDateIn"]
    ).dt.days.clip(lower=0)
    
    valid["Total"] = valid["SupplierDelay"] + valid["ReviewDuration"]
    valid = valid[valid["Total"] > 0]
    
    supplier_pct = valid["SupplierDelay"].sum() / valid["Total"].sum() * 100
    review_pct = valid["ReviewDuration"].sum() / valid["Total"].sum() * 100
    
    # Supplier Performance Analysis
    valid["SupplierName"] = valid["Select List 5"].fillna("Internal")
    
    supplier_perf = valid.groupby("SupplierName").agg(
        Total_Docs=("Document No", "count"),
        OnTime_Docs=("SupplierOnTime", lambda x: x.sum()),
        Delayed_Docs=("SupplierOnTime", lambda x: (~x).sum()),
        Avg_Delay=("SupplierDelay", "mean"),
        Max_Delay=("SupplierDelay", "max")
    ).reset_index()
    
    supplier_perf = supplier_perf.sort_values("Delayed_Docs", ascending=False)
    supplier_perf["OnTime_%"] = (supplier_perf["OnTime_Docs"] / supplier_perf["Total_Docs"] * 100).round(1)
    supplier_perf["Avg_Delay"] = supplier_perf["Avg_Delay"].round(1)
    supplier_perf["Max_Delay"] = supplier_perf["Max_Delay"].fillna(0).astype(int)
    supplier_perf.set_index("SupplierName", inplace=True)
    
    # Prepare flow trend data
    flow = df.dropna(subset=["FirstDateIn", "ApprovedDate"]).copy()
    flow["Week_Sub"] = flow["FirstDateIn"].dt.to_period("W").dt.start_time
    flow["Week_App"] = flow["ApprovedDate"].dt.to_period("W").dt.start_time
    
    submitted = flow.groupby("Week_Sub").size()
    approved = flow.groupby("Week_App").size()
    
    trend = pd.concat([submitted, approved], axis=1, sort=True).fillna(0)
    trend.columns = ["Submitted", "Approved"]
    trend["Backlog"] = (trend["Submitted"] - trend["Approved"]).cumsum()
    
    # Prepare discipline risk data
    df_sla = df.dropna(subset=["DueDate", "ApprovedDate"]).copy()
    df_sla["Delay"] = (df_sla["ApprovedDate"] - df_sla["DueDate"]).dt.days.clip(lower=0)
    
    risk = df_sla.groupby("Select List 1").agg(
        Total=("Document No", "count"),
        Late=("Delay", lambda x: (x > 0).sum())
    )
    risk["Breach_%"] = risk["Late"] / risk["Total"] * 100
    risk = risk.sort_values("Breach_%", ascending=False)
    
    # Prepare reviewer performance data
    df_rev = df.dropna(subset=["FirstDateIn", "ApprovedDate", "Assigned_To"]).copy()
    df_rev["ReviewDays"] = (df_rev["ApprovedDate"] - df_rev["FirstDateIn"]).dt.days
    
    reviewer_data = []
    for reviewer in df_rev["Assigned_To"].unique():
        reviewer_docs = df_rev[df_rev["Assigned_To"] == reviewer]
        if len(reviewer_docs) == 0:
            continue
        first_review = reviewer_docs["FirstDateIn"].min()
        last_review = reviewer_docs["ApprovedDate"].max()
        days_active = max((last_review - first_review).days, 1)
        weeks_active = days_active / 7
        throughput = len(reviewer_docs) / weeks_active
        reviewer_data.append({
            "Reviewer": reviewer,
            "Count": len(reviewer_docs),
            "Avg_Days": reviewer_docs["ReviewDays"].mean(),
            "OnTime_%": reviewer_docs["ReviewerOnTime"].mean() * 100,
            "Docs_per_Week": throughput,
            "Days_Active": days_active,
            "Weeks_Active": weeks_active
        })
    
    if reviewer_data:
        review_perf = pd.DataFrame(reviewer_data)
        review_perf = review_perf.sort_values("Docs_per_Week", ascending=False)
        review_perf.set_index("Reviewer", inplace=True)
    else:
        review_perf = pd.DataFrame()
    
    # Prepare review duration distribution
    dist = df_rev["ReviewDays"].dropna()
    bins = [0, 5, 10, 20, 100]
    labels = ["0-5", "6-10", "11-20", "20+"]
    bucket = pd.cut(dist, bins=bins, labels=labels, right=False)
    counts = bucket.value_counts().sort_index()
    percentages = (counts / counts.sum() * 100).round(1)
    
    return {
        "df": df,
        "valid": valid,
        "supplier_on_time": supplier_on_time,
        "review_on_time": review_on_time,
        "median_cycle": median_cycle,
        "backlog": backlog,
        "supplier_pct": supplier_pct,
        "review_pct": review_pct,
        "supplier_perf": supplier_perf,
        "trend": trend,
        "risk": risk,
        "review_perf": review_perf,
        "dist_counts": counts,
        "dist_percentages": percentages,
        "dist_labels": labels
    }

def create_download_link(df, filename="data.csv"):
    """Generate a download link for a dataframe"""
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def get_dynamic_actions(results, supplier_perf):
    """Generate dynamic action items with supplier context and readable insights"""
    actions = []
    
    # --- 1. Time contribution analysis ---
    if results['review_pct'] > results['supplier_pct']:
        actions.append({
            "priority": "🔴" if results['review_pct'] > 70 else "🟠",
            "category": "Review Process Bottleneck",
            "action": f"Review contribution ({results['review_pct']:.1f}%) exceeds supplier contribution ({results['supplier_pct']:.1f}%). Consider:",
            "items": [
                "Increase reviewer capacity or redistribute workload",
                "Implement priority queuing for urgent documents",
                "Review and streamline approval workflows"
            ]
        })
    else:
        # --- 2. Supplier Delay Dominant with text explanation ---
        max_delays = supplier_perf["Delayed_Docs"].max()
        top_late_suppliers = supplier_perf[supplier_perf["Delayed_Docs"] == max_delays].index.tolist()
        never_ontime = supplier_perf[supplier_perf["OnTime_%"] == 0].index.tolist()
        
        if len(never_ontime) > 0:
            insight_text = (
                f"- Supplier(s) with the **most delayed documents ({max_delays})**: {', '.join(top_late_suppliers)}. "
                f"These suppliers may need follow-up to improve submission timeliness.\n"
                f"- Suppliers that were **never on time**: {', '.join(never_ontime)}. Might check the schedule with them."
            )
        else:
            insight_text = (
                "From the supplier submission performance, you can see:\n"
                f"- Supplier(s) with the **most delayed documents ({max_delays})**: {', '.join(top_late_suppliers)}.\n"
                "- All suppliers have at least some on-time submissions."
            )
        
        actions.append({
            "priority": "🔴" if results['supplier_pct'] > 70 else "🟠",
            "category": "Supplier Delay Dominant",
            "action": f"Supplier delay ({results['supplier_pct']:.1f}%) is the main contributor.",
            "insight": insight_text
        })
    
    # --- 3. Backlog analysis ---
    if results['backlog'] > 20:
        actions.append({
            "priority": "🔴",
            "category": "High Backlog Pressure",
            "action": f"*Backlog pressure* is {results['backlog']:.1f}% (target <20%). Immediate actions required:",
            "items": [
                "Organize a backlog clearing session",
                "Temporarily reassign resources to review queue",
                "Prioritize oldest documents first"
            ]
        })
    elif results['backlog'] > 10:
        actions.append({
            "priority": "🟠",
            "category": "Moderate Backlog Pressure",
            "action": f"*Backlog pressure* is {results['backlog']:.1f}%. Monitor closely:",
            "items": [
                "Track weekly trend to ensure it doesn't increase",
                "Review if current capacity can handle incoming volume"
            ]
        })
    
    # --- 4. Discipline risk analysis ---
    if not results['risk'].empty:
        top_discipline = results['risk'].index[0]
        top_breach = results['risk'].iloc[0]['Breach_%']
        
        if top_breach > 50:
            actions.append({
                "priority": "🔴",
                "category": "Critical Discipline Risk",
                "action": f"'{top_discipline}' has {top_breach:.1f}% breach rate. Immediate focus required:",
                "items": [
                    f"Review all documents from {top_discipline} discipline",
                    "Meet with discipline leads to understand root causes",
                    "Consider temporary additional support for this area"
                ]
            })
        elif top_breach > 30:
            actions.append({
                "priority": "🟠",
                "category": "Discipline Risk",
                "action": f"'{top_discipline}' has {top_breach:.1f}% breach rate. Action needed:",
                "items": [
                    f"Monitor {top_discipline} documents closely",
                    "Identify patterns in delayed reviews",
                    "Provide targeted support to reviewers in this discipline"
                ]
            })
    
    # --- 5. Reviewer performance analysis ---
    if not results['review_perf'].empty:
        avg_throughput = results['review_perf']['Docs_per_Week'].mean()
        slow_reviewers = results['review_perf'][results['review_perf']['Docs_per_Week'] < avg_throughput * 0.5]
        
        if len(slow_reviewers) > 0:
            names_list = slow_reviewers.index.tolist()
            chunk_size = 4
            
            # Split names into lines of 4
            name_lines = []
            for i in range(0, len(names_list), chunk_size):
                chunk = names_list[i:i+chunk_size]
                name_lines.append(", ".join(chunk))
            
            # Join lines with <br> for proper multi-line display
            reviewers_text = "Reviewers:<br>" + "<br>".join(name_lines)
            
            actions.append({
                "priority": "🟠",
                "category": "Low Throughput Reviewers",
                "action": f"{len(slow_reviewers)} reviewers have below-average throughput:",
                "items": [
                    reviewers_text,
                    "Check if they need additional training or support",
                    "Review their workload distribution"
                ]
            })
    
    # --- Backlog good section ---
    if results['backlog'] <= 10:
        actions.append({
            "priority": "🟢",
            "category": "Backlog Pressure Low",
            "action": f"Backlog is {results['backlog']:.1f}%, which is within target. Keep monitoring.",
            "items": [
                "Maintain current review capacity",
                "Continue tracking weekly trends"
            ]
        })
    
    # --- Discipline good section ---
    if results['risk'].empty or results['risk']['Breach_%'].max() <= 30:
        actions.append({
            "priority": "🟢",
            "category": "Discipline Performance Healthy",
            "action": "All disciplines have breach rates within acceptable limits.",
            "items": [
                "No immediate action required",
                "Maintain current document review standards"
            ]
        })
    
    return actions

# Main Streamlit app
def main():
    st.markdown('<h1 class="main-header">Document Flow Performance Dashboard</h1>', unsafe_allow_html=True)
    
    st.sidebar.header("Upload Your Data")
    st.sidebar.info("""
    Please upload two Excel files:
    1. Supplier Documents (ExportSupplierDocuments.xlsx)
    2. Workflow Data (ExportWorkflow.xlsx)
    """)
    
    # File uploaders
    supplier_file = st.sidebar.file_uploader("Upload Supplier Documents", type=["xlsx"])
    workflow_file = st.sidebar.file_uploader("Upload Workflow Data", type=["xlsx"])
    
    if supplier_file and workflow_file:
        # Read and process data
        supplier_docs_raw = pd.read_excel(supplier_file)
        workflow_raw = pd.read_excel(workflow_file)
        
        supplier_docs = clean_header(supplier_docs_raw)
        workflow = clean_header(workflow_raw)
        
        # Process data
        results = process_data(supplier_docs, workflow)
        
        # KPI Explanations
        st.markdown("""
        <div class="explanation-box">
            <h4>📊 Understanding the KPIs:</h4>
            <ul>
                <li><strong>SLA (Service Level Agreement):</strong> A target time frame for completing reviews. Documents are considered "on-time" if approved by the due date.</li>
                <li><strong>Supplier On-Time:</strong> Percentage of documents submitted by suppliers before or on the planned submission date.</li>
                <li><strong>Review On-Time (SLA):</strong> Percentage of documents reviewed and approved within the SLA due date.</li>
                <li><strong>Median Cycle Time:</strong> The middle value of total days from planned submission to final approval (less sensitive to outliers than average).</li>
                <li><strong>Backlog Pressure:</strong> Percentage of documents that are overdue (past due date and not yet approved). Higher values indicate growing workflow bottlenecks.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        # Display KPIs
        st.markdown("### Key Performance Indicators")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown(f"""
            <div class="metric-container">
                <h3 style="color: #1f77b4;">{results['supplier_on_time']:.1f}%</h3>
                <p>Supplier On-Time</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"""
            <div class="metric-container">
                <h3 style="color: #ff7f0e;">{results['review_on_time']:.1f}%</h3>
                <p>Review On-Time (SLA)</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown(f"""
            <div class="metric-container">
                <h3 style="color: #2ca02c;">{results['median_cycle']:.0f} days</h3>
                <p>Median Cycle Time</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            backlog_color = "#d62728" if results['backlog'] > 20 else "#ff7f0e" if results['backlog'] > 10 else "#2ca02c"
            st.markdown(f"""
            <div class="metric-container">
                <h3 style="color: {backlog_color};">{results['backlog']:.1f}%</h3>
                <p>Backlog Pressure</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Time Contribution Analysis
        st.markdown('<h2 class="section-header">⏱️ Time Contribution Analysis</h2>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.bar(["Supplier", "Review"], [results['supplier_pct'], results['review_pct']], 
                   color=['#1f77b4', '#ff7f0e'])
            ax.set_title("Contribution to Total Cycle Time", fontsize=14, fontweight='bold')
            ax.set_ylabel("% of Total Duration")
            ax.set_ylim(0, 100)
            ax.grid(True, axis='y', linestyle='--', alpha=0.7)
            
            for i, v in enumerate([results['supplier_pct'], results['review_pct']]):
                ax.text(i, v + 1, f"{v:.1f}%", ha='center', fontweight='bold')
            
            st.pyplot(fig)
            plt.close(fig)
        
        with col2:
            bottleneck = "Review" if results['review_pct'] > results['supplier_pct'] else "Supplier"
            st.markdown(f"""
            <div class="chart-container">
                <h4>Key Insight</h4>
                <p><strong>Main bottleneck:</strong> {bottleneck} phase ({max(results['review_pct'], results['supplier_pct']):.1f}% of time)</p>
                <p><strong>Balance ratio:</strong> {results['supplier_pct']/results['review_pct']:.2f}:1 (Supplier:Review)</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Supplier Submission Performance - UPDATED with Jupyter styling
        st.markdown('<h2 class="section-header">📋 Supplier Submission Performance</h2>', unsafe_allow_html=True)
        
        st.markdown("This table shows how each supplier (or internal documents) performed in terms of submissions relative to the planned dates:")
        
        # Apply the same styling as Jupyter notebook
        supplier_display = results['supplier_perf'][["Total_Docs", "OnTime_Docs", "Delayed_Docs", "OnTime_%", "Avg_Delay", "Max_Delay"]].copy()
        
        # Create styled dataframe
        styled_supplier = supplier_display.style.format({
            "OnTime_%": "{:.1f}%",
            "Avg_Delay": "{:.1f}",
            "Max_Delay": "{:.0f}"
        }).background_gradient(subset=["Delayed_Docs"], cmap="Reds"
        ).background_gradient(subset=["OnTime_%"], cmap="RdYlGn"
        ).background_gradient(subset=["Avg_Delay"], cmap="Oranges"
        ).set_properties(**{"text-align": "center", "min-width": "120px"})
        
        st.dataframe(styled_supplier, use_container_width=True)
        
        # Text explanation
        max_delays = results['supplier_perf']["Delayed_Docs"].max()
        top_late_suppliers = results['supplier_perf'][results['supplier_perf']['Delayed_Docs'] == max_delays].index.tolist()
        never_ontime = results['supplier_perf'][results['supplier_perf']["OnTime_%"] == 0].index.tolist()
        
        if never_ontime:
            st.markdown(f"""
            <div class="explanation-box">
                <strong>From the table above, you can see:</strong><br>
                - Supplier(s) with the <strong>most delayed documents ({max_delays})</strong>: {', '.join(top_late_suppliers)}. These suppliers may need follow-up to improve submission timeliness.<br>
                - Suppliers that were <strong>never on time</strong>: {', '.join(never_ontime)}. Might check the schedule with them.
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="explanation-box">
                <strong>From the table above, you can see:</strong><br>
                - Supplier(s) with the <strong>most delayed documents ({max_delays})</strong>: {', '.join(top_late_suppliers)}. These suppliers may need follow-up to improve submission timeliness.<br>
                - ✅ All suppliers have at least some on-time submissions.
            </div>
            """, unsafe_allow_html=True)
        
        # Flow Trend
        st.markdown('<h2 class="section-header">📈 Submission vs Approval Trend</h2>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            fig, ax = plt.subplots(figsize=(14, 6))
            ax.plot(results['trend'].index, results['trend']['Submitted'], 
                    marker='o', label='Submitted', color='#1f77b4', linewidth=2, markersize=8)
            ax.plot(results['trend'].index, results['trend']['Approved'], 
                    marker='s', label='Approved', color='#ff7f0e', linewidth=2, markersize=8)
            ax.plot(results['trend'].index, results['trend']['Backlog'], 
                    marker='d', label='Backlog', color='#d62728', linewidth=2, markersize=8)
            ax.set_title("Document Flow Over Time", fontsize=14, fontweight='bold')
            ax.set_xlabel("Week")
            ax.set_ylabel("Number of Documents")
            ax.legend(loc='best')
            ax.grid(True, linestyle='--', alpha=0.7)
            
            st.pyplot(fig)
            plt.close(fig)
        
        with col2:
            current_backlog = int(results['trend']['Backlog'].iloc[-1]) if not results['trend'].empty else 0
            approval_rate = results['trend']['Approved'].mean() if not results['trend'].empty else 0
            submission_rate = results['trend']['Submitted'].mean() if not results['trend'].empty else 0
            
            st.markdown(f"""
            <div class="chart-container">
                <h4>Flow Metrics</h4>
                <p><strong>Current Backlog:</strong> {current_backlog} docs</p>
                <p><strong>Avg Weekly Submissions:</strong> {submission_rate:.0f}</p>
                <p><strong>Avg Weekly Approvals:</strong> {approval_rate:.0f}</p>
                <p><strong>Processing Gap:</strong> {submission_rate - approval_rate:.0f} docs/week</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Discipline Risk Analysis - UPDATED with Jupyter styling
        st.markdown('<h2 class="section-header">⚠️ Discipline Risk Analysis</h2>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            fig, ax = plt.subplots(figsize=(12, 7))
            top_risk = results['risk'].head(10)
            y_pos = np.arange(len(top_risk.index))
            
            # Use GnBu colormap like in Jupyter
            norm = plt.Normalize(0, top_risk['Breach_%'].max())
            cmap = plt.cm.GnBu
            colors = cmap(norm(top_risk['Breach_%']))
            
            ax.barh(y_pos, top_risk['Breach_%'], color=colors)
            ax.set_yticks(y_pos)
            ax.set_yticklabels(top_risk.index)
            ax.set_xlabel("Breach %")
            ax.set_title("Top 10 Disciplines by SLA Breach %", fontsize=14, fontweight='bold')
            ax.grid(True, axis='x', linestyle='--', alpha=0.7)
            
            for i, v in enumerate(top_risk['Breach_%']):
                ax.text(v + 0.5, i, f"{v:.1f}%", va='center', fontweight='bold')
            
            st.pyplot(fig)
            plt.close(fig)
        
        with col2:
            high_risk_count = len(results['risk'][results['risk']['Breach_%'] > 50])
            med_risk_count = len(results['risk'][results['risk']['Breach_%'] > 30]) - high_risk_count
            
            st.markdown(f"""
            <div class="chart-container">
                <h4>Risk Summary</h4>
                <p><strong>🔴 Critical (>50%):</strong> {high_risk_count}</p>
                <p><strong>🟠 High (30-50%):</strong> {med_risk_count}</p>
                <p><strong>🟢 Low (<30%):</strong> {len(results['risk']) - high_risk_count - med_risk_count}</p>
                <hr>
                <h4>Top 3 Risk Areas:</h4>
                <ol>
                    {''.join([f"<li><strong>{results['risk'].index[idx]}</strong> ({results['risk'].iloc[idx]['Breach_%']:.1f}%)</li>" for idx in range(min(3, len(results['risk'])))])}
                </ol>
            </div>
            """, unsafe_allow_html=True)
        
        # Top 3 Risk Disciplines Styled Table
        st.markdown("#### 🏆 Top 3 Risk Disciplines")
        top3 = results['risk'].head(3).copy()
        styled_top3 = top3.style.format({
            "Breach_%": "{:.1f}%",
            "Total": "{:.0f}",
            "Late": "{:.0f}"
        }).background_gradient(subset=["Breach_%"], cmap="Reds"
        ).background_gradient(subset=["Late"], cmap="Oranges"
        ).background_gradient(subset=["Total"], cmap="Blues"
        ).set_properties(**{"text-align": "center", "min-width": "120px"})
        
        st.dataframe(styled_top3, use_container_width=True)
        
        # Detailed Discipline Risk Data - UPDATED with Jupyter styling
        st.markdown("#### 📋 Detailed Discipline Risk Data")
        styled_risk = results['risk'].style.format({
            "Breach_%": "{:.1f}%",
            "Total": "{:.0f}",
            "Late": "{:.0f}"
        }).background_gradient(subset=["Breach_%"], cmap="Reds"
        ).background_gradient(subset=["Late"], cmap="Oranges"
        ).background_gradient(subset=["Total"], cmap="Blues"
        ).set_properties(**{"text-align": "center", "min-width": "120px"})
        
        st.dataframe(styled_risk, use_container_width=True)
        
        # Reviewer Performance
        st.markdown('<h2 class="section-header">👥 Reviewer Performance</h2>', unsafe_allow_html=True)
        
        if not results['review_perf'].empty:
            st.markdown("Reviewer throughput, average review days, and SLA compliance.")
            
            # Display reviewer performance table with green gradient on OnTime_%
            display_cols = ['Count', 'Docs_per_Week', 'Avg_Days', 'OnTime_%', 'Days_Active', 'Weeks_Active']
            styled_reviewer = results['review_perf'][display_cols].style.format({
                "Docs_per_Week": "{:.2f}",
                "Avg_Days": "{:.1f}",
                "OnTime_%": "{:.1f}%",
                "Days_Active": "{:.0f}",
                "Weeks_Active": "{:.1f}"
            }).background_gradient(cmap='Greens', subset=['OnTime_%'])
            
            st.dataframe(styled_reviewer, use_container_width=True)
            
            # Add explanation
            avg_throughput = results['review_perf']['Docs_per_Week'].mean()
            avg_ontime = results['review_perf']['OnTime_%'].mean()
            
            st.markdown(f"""
            <div class="explanation-box">
                <h4>📊 Reviewer Performance Summary</h4>
                <ul>
                    <li><strong>Average throughput:</strong> {avg_throughput:.2f} documents per week</li>
                    <li><strong>Average on-time rate:</strong> {avg_ontime:.1f}%</li>
                    <li><strong>Total active reviewers:</strong> {len(results['review_perf'])}</li>
                    <li><strong>Total documents reviewed:</strong> {int(results['review_perf']['Count'].sum())}</li>
                </ul>
                <p><em>Higher throughput and on-time % indicate better performance. Reviewers with low throughput may need attention.</em></p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.info("No completed reviews found to analyze reviewer performance.")
        
        # Review Duration Distribution
        st.markdown('<h2 class="section-header">📊 Review Duration Distribution</h2>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            fig, ax = plt.subplots(figsize=(10, 6))
            bars = ax.bar(results['dist_labels'], results['dist_percentages'].values, 
                          color=['#2ca02c', '#ff7f0e', '#ff7f0e', '#d62728'])
            ax.set_title("Distribution of Review Durations", fontsize=14, fontweight='bold')
            ax.set_xlabel("Duration Range (days)")
            ax.set_ylabel("Percentage of Documents (%)")
            ax.grid(True, axis='y', linestyle='--', alpha=0.7)
            
            for i, v in enumerate(results['dist_percentages']):
                ax.text(i, v + 1, f"{v:.1f}%", ha='center', va='bottom', fontweight='bold')
            
            st.pyplot(fig)
            plt.close(fig)
        
        with col2:
            fast_percent = results['dist_percentages'].get('0-5', 0)
            slow_percent = results['dist_percentages'].get('20+', 0)
            
            st.markdown(f"""
            <div class="chart-container duration-summary">
                <h4>Duration Summary</h4>
                <p>The chart shows the percentage of documents completed within different review durations:</p>
                <ul>
                    <li><strong style="color:#2ca02c;">✅ 0-5 days:</strong> Fast reviews ({fast_percent:.1f}%)</li>
                    <li><strong style="color:#ff7f0e;">🟡 6-10 days:</strong> Normal duration</li>
                    <li><strong style="color:#ff7f0e;">🟡 11-20 days:</strong> Slow reviews</li>
                    <li><strong style="color:#d62728;">🔴 20+ days:</strong> Very slow reviews ({slow_percent:.1f}%)</li>
                </ul>
                <p><strong>Action:</strong> Investigate documents in the 20+ days range for potential bottlenecks.</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Detailed Distribution Table with Jupyter styling
        st.markdown("#### 📋 Detailed Distribution Data")
        
        # Ensure arrays match
        min_len = min(len(results['dist_labels']), len(results['dist_counts']), len(results['dist_percentages']))
        dist_table = pd.DataFrame({
            "Duration Range": results['dist_labels'][:min_len],
            "Document Count": results['dist_counts'].values[:min_len],
            "Percentage (%)": results['dist_percentages'].values[:min_len]
        })
        
        # Apply subtle highlighting like in Jupyter
        styled_dist = dist_table.style.format({
            "Percentage (%)": "{:.2f}%"
        }).highlight_max(subset=["Document Count"], color="#d4edda"
        ).highlight_max(subset=["Percentage (%)"], color="#fff3cd"
        ).set_caption("📊 Document Duration Distribution")
        
        st.dataframe(styled_dist, use_container_width=True)
        
        # Dynamic Action Plan - UPDATED with Jupyter styling
        st.markdown('<h2 class="section-header">🎯 Dynamic Action Plan</h2>', unsafe_allow_html=True)
        
        actions = get_dynamic_actions(results, results['supplier_perf'])
        
        if actions:
            for action in actions:
                # Determine box class based on priority emoji
                if action["priority"] == "🔴":
                    box_class = "warning-box"
                elif action["priority"] == "🟠":
                    box_class = "action-box"
                else:  # 🟢
                    box_class = "success-box"
                
                action_html = f"""
                <div class="{box_class}">
                    <h4>{action['priority']} {action['category']}</h4>
                    <p>{action['action']}</p>
                """
                
                if "insight" in action:
                    insight_lines = action['insight'].split("\n")
                    for line in insight_lines:
                        if line.strip():
                            action_html += f'<div class="insight-text">{line}</div>'
                
                if "items" in action:
                    action_html += "<ul>"
                    for item in action["items"]:
                        # Handle HTML line breaks in items
                        if "<br>" in item:
                            parts = item.split("<br>")
                            for part in parts:
                                if part.strip():
                                    action_html += f"<li>{part}</li>"
                        else:
                            action_html += f"<li>{item}</li>"
                    action_html += "</ul>"
                
                action_html += "</div>"
                st.markdown(action_html, unsafe_allow_html=True)
                st.markdown("---")
        else:
            st.markdown("""
            <div class="success-box">
                <h4>✅ All Systems Normal</h4>
                <p>No critical issues detected. Continue monitoring and maintaining current performance levels.</p>
            </div>
            """, unsafe_allow_html=True)
        
        # # Data download section
        # st.markdown('<h2 class="section-header">📥 Download Processed Data</h2>', unsafe_allow_html=True)
        # col1, col2, col3, col4 = st.columns(4)
        
        # with col1:
        #     st.markdown(create_download_link(results['df'], "merged_data.csv"), unsafe_allow_html=True)
        
        # with col2:
        #     st.markdown(create_download_link(results['risk'], "discipline_risk.csv"), unsafe_allow_html=True)
        
        # with col3:
        #     if not results['review_perf'].empty:
        #         st.markdown(create_download_link(results['review_perf'], "reviewer_performance.csv"), unsafe_allow_html=True)
        
        # with col4:
            st.markdown(create_download_link(results['supplier_perf'], "supplier_performance.csv"), unsafe_allow_html=True)
    
    else:
        st.info("Please upload both Excel files to begin the analysis.")
        st.image("https://cdn-icons-png.flaticon.com/512/3135/3135711.png", width=200)

if __name__ == "__main__":
    main()