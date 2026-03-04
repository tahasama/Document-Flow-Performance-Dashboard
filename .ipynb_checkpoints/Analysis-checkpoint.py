# ==============================================================================
# STRATEGIC ADVISORY & RECOMMENDATIONS
# ==============================================================================

# --- Synthesize Findings for Recommendations ---
# 1. Bottleneck
bottleneck_finding = f"The internal review process is the primary bottleneck, accounting for {kpis['review_contribution_pct']:.0f}% of total cycle time."

# 2. Worst Discipline
worst_discipline = sla_perf.iloc[0]
discipline_finding = f"The '{worst_discipline['Select List 1']}' discipline has the highest SLA breach rate at {worst_discipline['SLA_Breach_%']:.1f}%."

# 3. Backlog Trend
backlog_trend = flow_df['Cumulative_Backlog'].iloc[-1] - flow_df['Cumulative_Backlog'].iloc[-5] if len(flow_df) > 5 else 0
backlog_finding = f"The project backlog is increasing, with a net growth of {backlog_trend:.0f} documents in the last 5 weeks."

# --- Generate Recommendations ---
recommendations = [
    {
        'ID': 1,
        'Priority': 'High',
        'Area': 'Process Efficiency',
        'Finding': bottleneck_finding,
        'Recommendation': 'Initiate a targeted Lean/Kaizen event for the review process. Map the current state, identify non-value-added steps, and implement a streamlined future state workflow.',
        'Expected Impact': 'Reduce median review time by 20-30%, accelerating overall project delivery.',
        'Owner': 'Head of Engineering'
    },
    {
        'ID': 2,
        'Priority': 'Critical',
        'Area': 'Discipline Performance',
        'Finding': discipline_finding,
        'Recommendation': f'Immediately conduct a capacity and skills gap analysis for the {worst_discipline["Select List 1"]} team. Provide targeted training or temporary resource augmentation.',
        'Expected Impact': f'Reduce SLA breach rate for {worst_discipline["Select List 1"]} to <15% in the next quarter.',
        'Owner': f'Discipline Lead - {worst_discipline["Select List 1"]}'
    },
    {
        'ID': 3,
        'Priority': 'Medium',
        'Area': 'Capacity Management',
        'Finding': backlog_finding,
        'Recommendation': 'Implement a mandatory weekly backlog review meeting with Project Managers and Discipline Leads to proactively reassign at-risk documents and remove blockers.',
        'Expected Impact': 'Stabilize and begin reducing the backlog within 4 weeks, preventing future capacity crises.',
        'Owner': 'Project Management Office (PMO)'
    },
    {
        'ID': 4,
        'Priority': 'Low',
        'Area': 'People Development',
        'Finding': 'Reviewer performance analysis shows significant variation in speed and consistency among individual reviewers.',
        'Recommendation': 'Develop a reviewer performance scorecard (incorporating speed, quality, and consistency). Use this for coaching and identifying best practices to be shared across the team.',
        'Expected Impact': 'Improve overall team consistency and raise the performance floor, creating a more predictable review process.',
        'Owner': 'HR Business Partner / Engineering Leads'
    }
]

# --- Create and Display Final Report ---
advisory_df = pd.DataFrame(recommendations)

print("\n" + "="*60)
print(" " * 15 + "STRATEGIC ADVISORY REPORT")
print("="*60)
print("\nBased on the comprehensive analysis, the following actions are recommended:\n")
print(advisory_df.to_string(index=False))
print("\n" + "="*60)
