import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Load data
file_path = 'Project_Management_Template_Updated.xlsx'
df_workstreams = pd.read_excel(file_path, sheet_name='Workstreams')
df_risks = pd.read_excel(file_path, sheet_name='Risk_Register')
df_resources = pd.read_excel(file_path, sheet_name='Resources')

st.set_page_config(page_title="Project Dashboard", layout="wide")
st.title("📊 Project Management Dashboard")

# --- 1. Workstream Progress ---
st.subheader("📈 Workstream Progress")
df_workstreams['Progress %'] = pd.to_numeric(df_workstreams['Progress %'], errors='coerce')
workstream_progress = df_workstreams.groupby('Work-stream')['Progress %'].mean().reset_index()
fig_progress = px.bar(workstream_progress, x='Work-stream', y='Progress %', text='Progress %', title='Workstream Progress')
st.plotly_chart(fig_progress, use_container_width=True)

# --- 2. Project Risk Overview ---
st.subheader("⚠️ Project Risks")
total_risks = df_risks.shape[0]
risk_distribution = df_risks['Risk Score'].value_counts().reset_index()
risk_distribution.columns = ['Risk Score', 'Count']
col1, col2 = st.columns(2)
col1.metric("Total Risks", total_risks)
fig_risks = px.bar(risk_distribution, x='Risk Score', y='Count', title='Risks by Score')
col2.plotly_chart(fig_risks, use_container_width=True)

# --- 3. Planned Tasks This Month ---
st.subheader("📅 Tasks Planned for This Month")
df_workstreams['Planned Start Date'] = pd.to_datetime(df_workstreams['Planned Start Date'], errors='coerce')
df_workstreams['Planned End Date'] = pd.to_datetime(df_workstreams['Planned End Date'], errors='coerce')

now = datetime.now()
start_of_month = now.replace(day=1)
end_of_month = now.replace(day=28) + pd.offsets.MonthEnd(1)
tasks_this_month = df_workstreams[
    (df_workstreams['Planned Start Date'] <= end_of_month) & 
    (df_workstreams['Planned End Date'] >= start_of_month)
]
st.dataframe(tasks_this_month[['Activity Code', 'Activity Name', 'Planned Start Date', 'Planned End Date']])

# --- 4. Team Members Active This Month ---
st.subheader("👥 Team Members Active This Month")
if 'Allocated/Used Hours' in df_resources.columns:
    df_resources = df_resources[pd.to_numeric(df_resources['Allocated/Used Hours'], errors='coerce').notnull()]
    st.dataframe(df_resources[['Person Name', 'Role', 'Allocated/Used Hours']])
