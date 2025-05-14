import streamlit as st
import pandas as pd
import plotly.express as px

# Load data
file_path = 'Project_Management_Template_Updated.xlsx'
df_workstreams = pd.read_excel(file_path, sheet_name='Workstreams')
df_resources = pd.read_excel(file_path, sheet_name='Resources')
df_budget = pd.read_excel(file_path, sheet_name='Budget_vs_Actual')
df_risks = pd.read_excel(file_path, sheet_name='Risk_Register')
df_issues = pd.read_excel(file_path, sheet_name='Issue_Tracker')
df_milestones = pd.read_excel(file_path, sheet_name='Milestones')

st.set_page_config(page_title="Project Management Dashboard", layout="wide")
st.title("📊 Project Management Dashboard")

# --- KPIs ---
col1, col2, col3, col4 = st.columns(4)
col1.metric("Workstreams", df_workstreams['Work-stream'].nunique())
col2.metric("Open Risks", df_risks[df_risks['Status'] == 'Open'].shape[0])
col3.metric("Open Issues", df_issues[df_issues['Status'] == 'Open'].shape[0] if not df_issues.empty else 0)
col4.metric("Milestones", df_milestones.shape[0])

st.markdown("---")

# --- Milestone Timeline ---
st.subheader("📅 Milestone Timeline")
df_milestones['Planned Date'] = pd.to_datetime(df_milestones['Planned Date'], errors='coerce')
milestone_fig = px.timeline(
    df_milestones.dropna(subset=['Planned Date']),
    x_start='Planned Date',
    x_end='Actual Date',
    y='Milestone Name',
    color='Workstream',
    title='Planned vs Actual Dates'
)
milestone_fig.update_yaxes(autorange="reversed")
st.plotly_chart(milestone_fig, use_container_width=True)

# --- Risk Overview ---
st.subheader("⚠️ Risk Overview")
risk_count = df_risks['Risk Score'].value_counts()
st.bar_chart(risk_count)

# --- Issue Tracker ---
st.subheader("🛑 Issue Tracker")
if not df_issues.empty:
    open_issues = df_issues[df_issues['Status'] == 'Open']
    st.dataframe(open_issues[['Issue ID', 'Workstream', 'Issue Description', 'Assigned To', 'Status']])
else:
    st.info("No issues have been reported yet.")

# --- Budget Tracker ---
st.subheader("💰 Budget Tracker")
if 'Planned Budget ($)' in df_budget.columns:
    df_budget_summary = df_budget[['Activity Name', 'Planned Budget ($)', 'Actual Cost ($)', 'Variance ($)']]
    st.dataframe(df_budget_summary)

# --- Resource Overview ---
st.subheader("👥 Resource Overview")
if 'Person Name' in df_resources.columns:
    resource_summary = df_resources[['Person Name', 'Role', 'Total Available Hours', 'Allocated/Used Hours']]
    st.dataframe(resource_summary)
