import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st

# 1. Page Configuration
st.set_page_config(
    page_title="FY31 State-wise Dispatch & Capacity Dashboard",
    page_icon="⚡",
    layout="wide",
)


# 2. Robust Data Loading & Parsing Function (Cached)
@st.cache_data
def load_and_process_data():
  excel_path = "../S2_Dispatch_FY31_statewise_output.xlsx"
  wb = openpyxl.load_workbook(excel_path, read_only=True)

  gen_records = []
  stor_records = []

  for state in wb.sheetnames:
    ws = wb[state]
    rows = list(ws.iter_rows(values_only=True))

    # Locate all 'Collection' header rows in this sheet
    hdr_indices = [i for i, r in enumerate(rows) if r and r[0] == "Collection"]

    # Parse Generators Table
    if len(hdr_indices) >= 1:
      g_idx = hdr_indices[0]
      g_headers = list(rows[g_idx])
      cat_col = g_headers.index("Category")
      cap_col = g_headers.index("Installed Capacity (MW)")

      end_row = hdr_indices[1] if len(hdr_indices) > 1 else len(rows)
      for r in rows[g_idx + 1 : end_row]:
        if r and r[0] == "Generator" and r[cat_col]:
          gen_records.append({
              "State": state,
              "Raw_Category": str(r[cat_col]).strip(),
              "Capacity_MW": float(r[cap_col] or 0.0),
          })

    # Parse Storage Table
    if len(hdr_indices) >= 2:
      s_idx = hdr_indices[1]
      s_headers = list(rows[s_idx])
      cat_col = s_headers.index("Category")
      cap_col = s_headers.index("Installed Capacity (MWh)")

      for r in rows[s_idx + 1 :]:
        if r and r[0] in ["Battery", "Storage"] and r[cat_col]:
          stor_records.append({
              "State": state,
              "Raw_Category": str(r[cat_col]).strip(),
              "Capacity_MWh": float(r[cap_col] or 0.0),
          })

  df_gen = pd.DataFrame(gen_records)
  df_stor = pd.DataFrame(stor_records)

  # Classification helper functions
  def classify_generator(cat):
    c = cat.lower()
    if "existing" in c:
      status = "Existing"
    elif "upcoming" in c:
      status = "Upcoming"
    elif "candidate" in c:
      status = "Candidates"
    elif "retirement" in c:
      status = "Retirement"
    else:
      status = "Other"

    if "thermal" in c:
      tech = "Thermal"
    elif "solar" in c:
      tech = "Solar"
    elif "wind" in c:
      tech = "Wind"
    elif "hydro" in c:
      tech = "Hydro"
    elif "nuclear" in c:
      tech = "Nuclear"
    elif "dre" in c:
      tech = "DRE"
    elif "other re" in c:
      tech = "Other RE"
    else:
      tech = "Other"
    return pd.Series([tech, status])

  def classify_storage(cat):
    c = cat.lower()
    if "existing" in c:
      status = "Existing"
    elif "upcoming" in c:
      status = "Upcoming"
    elif "candidate" in c:
      status = "Candidates"
    else:
      status = "Other"

    tech = "PSP" if "psp" in c else "BESS"
    return pd.Series([tech, status])

  if not df_gen.empty:
    df_gen[["Technology", "Status"]] = df_gen["Raw_Category"].apply(
        classify_generator
    )
  if not df_stor.empty:
    df_stor[["Technology", "Status"]] = df_stor["Raw_Category"].apply(
        classify_storage
    )

  return df_gen, df_stor


df_gen, df_stor = load_and_process_data()

# 3. Sidebar Navigation
st.sidebar.title("⚡ FY31 Capacity Dashboard")
page = st.sidebar.radio(
    "Select View", ["State Explorer", "All-India Overview"]
)
st.sidebar.markdown("---")
st.sidebar.info(
    "Source: `S2_Dispatch_FY31_statewise_output.xlsx` (PLEXOS Model)"
)

# --- VIEW 1: STATE EXPLORER ---
if page == "State Explorer":
  st.title("🗺️ State-wise Installed Capacity Explorer")
  st.markdown(
      "Explore generation (MW) and storage energy capacity (MWh) by state and"
      " technology."
  )

  states = sorted(df_gen["State"].unique())
  selected_state = st.selectbox("Select State / Region", states)

  # Filter data
  st_gen = df_gen[df_gen["State"] == selected_state]
  st_stor = (
      df_stor[df_stor["State"] == selected_state]
      if not df_stor.empty
      else pd.DataFrame()
  )

  # Top KPIs for Selected State
  total_mw = st_gen["Capacity_MW"].sum()
  exist_mw = st_gen[st_gen["Status"] == "Existing"]["Capacity_MW"].sum()
  up_mw = st_gen[st_gen["Status"] == "Upcoming"]["Capacity_MW"].sum()
  ret_mw = st_gen[st_gen["Status"] == "Retirement"]["Capacity_MW"].sum()

  k1, k2, k3, k4 = st.columns(4)
  k1.metric("Total Generation (MW)", f"{total_mw:,.1f}")
  k2.metric("Existing Capacity (MW)", f"{exist_mw:,.1f}")
  k3.metric("Upcoming Capacity (MW)", f"{up_mw:,.1f}")
  k4.metric("Retirement Capacity (MW)", f"{ret_mw:,.1f}")

  st.markdown("---")

  # Generator Charts Layout
  col1, col2 = st.columns(2)

  with col1:
    st.subheader(f"📊 Generator Technology Mix ({selected_state})")
    if not st_gen.empty:
      tech_df = (
          st_gen.groupby("Technology")["Capacity_MW"].sum().reset_index()
      )
      fig_tech = px.bar(
          tech_df,
          x="Technology",
          y="Capacity_MW",
          text="Capacity_MW",
          color="Technology",
      )
      fig_tech.update_traces(texttemplate="%{text:,.1f}", textposition="outside")
      fig_tech.update_layout(
          showlegend=False,
          yaxis_title="Capacity (MW)",
          yaxis=dict(tickformat=","),
      )
      st.plotly_chart(fig_tech, use_container_width=True)

  with col2:
    st.subheader(f"🏷️ Generator Status Breakdown ({selected_state})")
    if not st_gen.empty:
      status_df = (
          st_gen.groupby(["Technology", "Status"])["Capacity_MW"]
          .sum()
          .reset_index()
      )
      fig_status = px.bar(
          status_df,
          x="Technology",
          y="Capacity_MW",
          color="Status",
          barmode="group",
      )
      fig_status.update_layout(
          yaxis_title="Capacity (MW)", yaxis=dict(tickformat=",")
      )
      st.plotly_chart(fig_status, use_container_width=True)

  # Storage Charts & Table
  st.markdown("---")
  st.subheader(f"🔋 Storage Assets ({selected_state} - MWh)")
  if not st_stor.empty:
    scol1, scol2 = st.columns(2)

    with scol1:
      stor_tech_df = (
          st_stor.groupby("Technology")["Capacity_MWh"].sum().reset_index()
      )
      fig_stor_tech = px.bar(
          stor_tech_df,
          x="Technology",
          y="Capacity_MWh",
          text="Capacity_MWh",
          color="Technology",
          title=f"Storage Energy Capacity by Technology ({selected_state})",
      )
      fig_stor_tech.update_traces(
          texttemplate="%{text:,.1f}", textposition="outside"
      )
      fig_stor_tech.update_layout(
          showlegend=False,
          yaxis_title="Capacity (MWh)",
          yaxis=dict(tickformat=","),
      )
      st.plotly_chart(fig_stor_tech, use_container_width=True)

    with scol2:
      stor_stat_df = (
          st_stor.groupby(["Technology", "Status"])["Capacity_MWh"]
          .sum()
          .reset_index()
      )
      fig_stor_stat = px.bar(
          stor_stat_df,
          x="Technology",
          y="Capacity_MWh",
          color="Status",
          barmode="group",
          title=f"Storage Capacity by Status ({selected_state})",
      )
      fig_stor_stat.update_layout(
          yaxis_title="Capacity (MWh)", yaxis=dict(tickformat=",")
      )
      st.plotly_chart(fig_stor_stat, use_container_width=True)

    with st.expander("View Storage Data Table"):
      st.dataframe(st_stor, use_container_width=True)
  else:
    st.info(f"No dedicated storage records found for `{selected_state}`.")

# --- VIEW 2: ALL-INDIA OVERVIEW ---
elif page == "All-India Overview":
  st.title("🌐 All-India FY31 Capacity Overview")
  st.markdown(
      "Aggregated simulation output across all 27 states and union"
      " territories."
  )

  tot_gen = df_gen["Capacity_MW"].sum()
  tot_exist = df_gen[df_gen["Status"] == "Existing"]["Capacity_MW"].sum()
  tot_up = df_gen[df_gen["Status"] == "Upcoming"]["Capacity_MW"].sum()
  tot_ret = df_gen[df_gen["Status"] == "Retirement"]["Capacity_MW"].sum()

  c1, c2, c3, c4 = st.columns(4)
  c1.metric("National Total Gen (MW)", f"{tot_gen:,.1f}")
  c2.metric("National Existing (MW)", f"{tot_exist:,.1f}")
  c3.metric("National Upcoming (MW)", f"{tot_up:,.1f}")
  c4.metric("National Retirement (MW)", f"{tot_ret:,.1f}")

  st.markdown("---")

  # State-wise Stacked Bar Chart (No 'k' abbreviation on Y-axis)
  st.subheader("📈 State-wise Installed Capacity Breakdown (MW)")
  state_agg = (
      df_gen.groupby(["State", "Technology"])["Capacity_MW"].sum().reset_index()
  )
  fig_state = px.bar(
      state_agg,
      x="State",
      y="Capacity_MW",
      color="Technology",
      barmode="stack",
  )
  fig_state.update_layout(
      xaxis_title="State / Region",
      yaxis_title="Installed Capacity (MW)",
      yaxis=dict(tickformat=","),  # Disables 'k' abbreviation and shows full comma-separated numbers
  )
  st.plotly_chart(fig_state, use_container_width=True)

  # Interactive Visual Explorer (Charts instead of raw table first)
  st.markdown("---")
  st.subheader("🔍 Interactive Technology & Status Explorer")

  fcol1, fcol2 = st.columns(2)
  with fcol1:
    selected_techs = st.multiselect(
        "Filter by Technology",
        sorted(df_gen["Technology"].unique()),
        default=sorted(df_gen["Technology"].unique()),
    )
  with fcol2:
    selected_statuses = st.multiselect(
        "Filter by Status",
        sorted(df_gen["Status"].unique()),
        default=sorted(df_gen["Status"].unique()),
    )

  filtered_df = df_gen[
      df_gen["Technology"].isin(selected_techs)
      & df_gen["Status"].isin(selected_statuses)
  ]

  if not filtered_df.empty:
    vcol1, vcol2 = st.columns(2)

    with vcol1:
      filt_tech_agg = (
          filtered_df.groupby("Technology")["Capacity_MW"].sum().reset_index()
      )
      fig_filt_tech = px.bar(
          filt_tech_agg,
          x="Technology",
          y="Capacity_MW",
          text="Capacity_MW",
          color="Technology",
          title="Filtered Capacity by Technology (MW)",
      )
      fig_filt_tech.update_traces(
          texttemplate="%{text:,.1f}", textposition="outside"
      )
      fig_filt_tech.update_layout(
          showlegend=False,
          yaxis_title="Capacity (MW)",
          yaxis=dict(tickformat=","),
      )
      st.plotly_chart(fig_filt_tech, use_container_width=True)

    with vcol2:
      filt_stat_agg = (
          filtered_df.groupby("Status")["Capacity_MW"].sum().reset_index()
      )
      fig_filt_stat = px.bar(
          filt_stat_agg,
          x="Status",
          y="Capacity_MW",
          text="Capacity_MW",
          color="Status",
          title="Filtered Capacity by Status (MW)",
      )
      fig_filt_stat.update_traces(
          texttemplate="%{text:,.1f}", textposition="outside"
      )
      fig_filt_stat.update_layout(
          showlegend=False,
          yaxis_title="Capacity (MW)",
          yaxis=dict(tickformat=","),
      )
      st.plotly_chart(fig_filt_stat, use_container_width=True)

    with st.expander("View Underlying Filtered Data Table"):
      st.dataframe(filtered_df, use_container_width=True)
  else:
    st.warning("No records match your selected filters.")