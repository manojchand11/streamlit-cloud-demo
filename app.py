from pathlib import Path

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st


# ---------------------------------------------------------------------
# 1. PAGE CONFIGURATION
# ---------------------------------------------------------------------
st.set_page_config(
    page_title="FY31 State-wise Dispatch & Capacity Dashboard",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ FY31 State-wise Dispatch & Capacity Dashboard")


# ---------------------------------------------------------------------
# 2. FILE PATHS AND COLOUR SCHEMES
# Keep app.py and the Excel workbook in the same GitHub repository folder.
# ---------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "S2_Dispatch_FY31_statewise_output.xlsx"

# Consistent technology colours across all charts that display technology.
# Thermal is included in grey because the existing classification logic
# identifies thermal generation separately from an explicitly named Coal row.
TECHNOLOGY_COLORS = {
    "Solar": "#FFD700",       # Yellow
    "Wind": "#2E8B57",        # Green
    "Nuclear": "#000000",     # Black
    "Coal": "#808080",        # Grey
    "Thermal": "#808080",     # Grey (existing workbook classification)
    "DRE": "#90EE90",         # Light green
    "Hydro": "#87CEFA",       # Light blue
    "PSP": "#800080",         # Purple
    "BESS": "#FF69B4",        # Pink
    "Other RE": "#FF8C00",    # Orange (fallback category)
    "Other": "#A9A9A9",       # Light grey (fallback category)
}

STATUS_COLORS = {
    "Existing": "#4C78A8",
    "Upcoming": "#F2A541",
    "Candidates": "#72B7B2",
    "Retirement": "#E45756",
    "Other": "#A0A0A0",
}


# ---------------------------------------------------------------------
# 3. CLASSIFICATION AND DATA-UTILITY FUNCTIONS
# ---------------------------------------------------------------------
def classify_generator(category: str) -> tuple[str, str]:
    """Identify generator technology and project status."""
    text = str(category).strip().lower()

    if "existing" in text:
        status = "Existing"
    elif "upcoming" in text:
        status = "Upcoming"
    elif "candidate" in text:
        status = "Candidates"
    elif "retirement" in text:
        status = "Retirement"
    else:
        status = "Other"

    # Check specific technologies before generic categories.
    if "nuclear" in text:
        technology = "Nuclear"
    elif "solar" in text:
        technology = "Solar"
    elif "wind" in text:
        technology = "Wind"
    elif "hydro" in text:
        technology = "Hydro"
    elif "dre" in text:
        technology = "DRE"
    elif "coal" in text:
        technology = "Coal"
    elif "thermal" in text:
        # Retained as Thermal if that is how the source workbook labels it.
        technology = "Thermal"
    elif "other re" in text:
        technology = "Other RE"
    else:
        technology = "Other"

    return technology, status


def classify_storage(category: str) -> tuple[str, str]:
    """Identify storage technology and project status."""
    text = str(category).strip().lower()

    if "existing" in text:
        status = "Existing"
    elif "upcoming" in text:
        status = "Upcoming"
    elif "candidate" in text:
        status = "Candidates"
    else:
        status = "Other"

    technology = "PSP" if "psp" in text else "BESS"
    return technology, status


def numeric_value(value) -> float:
    """Convert Excel values to numbers; treat blanks/non-numeric cells as zero."""
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def apply_number_format(fig, y_title: str):
    """Apply consistent axis formatting to a Plotly chart."""
    fig.update_layout(
        yaxis_title=y_title,
        yaxis=dict(tickformat=","),
        margin=dict(t=55, b=45),
    )
    return fig


# ---------------------------------------------------------------------
# 4. LOAD AND PROCESS THE EXCEL WORKBOOK
# ---------------------------------------------------------------------
@st.cache_data(show_spinner="Loading and processing the Excel workbook...")
def load_and_process_data(file_path: str):
    workbook_path = Path(file_path)

    if not workbook_path.is_file():
        raise FileNotFoundError(
            f"Excel workbook not found: {workbook_path}\n\n"
            "Upload S2_Dispatch_FY31_statewise_output.xlsx to the same "
            "GitHub repository folder as app.py."
        )

    generator_records = []
    storage_records = []

    workbook = openpyxl.load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
    )

    try:
        for state in workbook.sheetnames:
            worksheet = workbook[state]
            rows = list(worksheet.iter_rows(values_only=True))

            # Locate the Collection header rows in this worksheet.
            header_indices = [
                index
                for index, row in enumerate(rows)
                if row and str(row[0]).strip() == "Collection"
            ]

            # -------------------------
            # Generator capacity table
            # -------------------------
            if header_indices:
                generator_header_index = header_indices[0]
                generator_headers = list(rows[generator_header_index])

                try:
                    category_col = generator_headers.index("Category")
                    capacity_col = generator_headers.index(
                        "Installed Capacity (MW)"
                    )
                except ValueError as error:
                    raise ValueError(
                        f"Generator table headers are missing in worksheet "
                        f"'{state}'. Expected 'Category' and "
                        "'Installed Capacity (MW)'."
                    ) from error

                generator_end = (
                    header_indices[1]
                    if len(header_indices) > 1
                    else len(rows)
                )

                for row in rows[generator_header_index + 1 : generator_end]:
                    if (
                        row
                        and len(row) > max(category_col, capacity_col)
                        and str(row[0]).strip() == "Generator"
                        and row[category_col]
                    ):
                        generator_records.append(
                            {
                                "State": state,
                                "Raw_Category": str(
                                    row[category_col]
                                ).strip(),
                                "Capacity_MW": numeric_value(
                                    row[capacity_col]
                                ),
                            }
                        )

            # -------------------------
            # Storage capacity table
            # -------------------------
            if len(header_indices) >= 2:
                storage_header_index = header_indices[1]
                storage_headers = list(rows[storage_header_index])

                try:
                    category_col = storage_headers.index("Category")
                    capacity_col = storage_headers.index(
                        "Installed Capacity (MWh)"
                    )
                except ValueError as error:
                    raise ValueError(
                        f"Storage table headers are missing in worksheet "
                        f"'{state}'. Expected 'Category' and "
                        "'Installed Capacity (MWh)'."
                    ) from error

                for row in rows[storage_header_index + 1 :]:
                    if (
                        row
                        and len(row) > max(category_col, capacity_col)
                        and str(row[0]).strip() in {"Battery", "Storage"}
                        and row[category_col]
                    ):
                        storage_records.append(
                            {
                                "State": state,
                                "Raw_Category": str(
                                    row[category_col]
                                ).strip(),
                                "Capacity_MWh": numeric_value(
                                    row[capacity_col]
                                ),
                            }
                        )
    finally:
        workbook.close()

    df_gen = pd.DataFrame(
        generator_records,
        columns=["State", "Raw_Category", "Capacity_MW"],
    )
    df_stor = pd.DataFrame(
        storage_records,
        columns=["State", "Raw_Category", "Capacity_MWh"],
    )

    if not df_gen.empty:
        classified = df_gen["Raw_Category"].apply(classify_generator)
        df_gen[["Technology", "Status"]] = pd.DataFrame(
            classified.tolist(),
            index=df_gen.index,
        )

    if not df_stor.empty:
        classified = df_stor["Raw_Category"].apply(classify_storage)
        df_stor[["Technology", "Status"]] = pd.DataFrame(
            classified.tolist(),
            index=df_stor.index,
        )

    return df_gen, df_stor


# ---------------------------------------------------------------------
# 5. LOAD DATA AND HANDLE ERRORS
# ---------------------------------------------------------------------
try:
    df_gen, df_stor = load_and_process_data(str(EXCEL_FILE))
except Exception as error:
    st.error("The dashboard could not load or process its Excel data.")
    st.exception(error)
    st.stop()

if df_gen.empty:
    st.error(
        "No generator records were found. Check the workbook's worksheets "
        "and the 'Collection' header rows."
    )
    st.stop()


# ---------------------------------------------------------------------
# 6. SIDEBAR NAVIGATION
# ---------------------------------------------------------------------
st.sidebar.title("⚡ FY31 Capacity Dashboard")

page = st.sidebar.radio(
    "Select View",
    ["State Explorer", "All-India Overview"],
)

st.sidebar.markdown("---")
st.sidebar.info(
    "Source: `S2_Dispatch_FY31_statewise_output.xlsx` (PLEXOS Model)"
)

st.sidebar.markdown("### Technology colour key")
for technology, colour in TECHNOLOGY_COLORS.items():
    st.sidebar.markdown(
        f'<div style="display:flex;align-items:center;margin:4px 0;">'
        f'<span style="background:{colour};width:13px;height:13px;'
        f'display:inline-block;border:1px solid #777;margin-right:8px;'
        f'border-radius:2px;"></span>{technology}</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------
# 7. VIEW 1 — STATE EXPLORER
# ---------------------------------------------------------------------
if page == "State Explorer":
    st.title("🗺️ State-wise Installed Capacity Explorer")
    st.markdown(
        "Explore generation capacity (MW) and storage energy capacity "
        "(MWh) by state and technology."
    )

    states = sorted(df_gen["State"].dropna().unique().tolist())
    selected_state = st.selectbox("Select State / Region", states)

    state_gen = df_gen[df_gen["State"] == selected_state].copy()

    if not df_stor.empty:
        state_stor = df_stor[
            df_stor["State"] == selected_state
        ].copy()
    else:
        state_stor = pd.DataFrame()

    # State-level generation KPIs
    total_mw = state_gen["Capacity_MW"].sum()
    existing_mw = state_gen.loc[
        state_gen["Status"] == "Existing", "Capacity_MW"
    ].sum()
    upcoming_mw = state_gen.loc[
        state_gen["Status"] == "Upcoming", "Capacity_MW"
    ].sum()
    retirement_mw = state_gen.loc[
        state_gen["Status"] == "Retirement", "Capacity_MW"
    ].sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Generation (MW)", f"{total_mw:,.1f}")
    k2.metric("Existing Capacity (MW)", f"{existing_mw:,.1f}")
    k3.metric("Upcoming Capacity (MW)", f"{upcoming_mw:,.1f}")
    k4.metric("Retirement Capacity (MW)", f"{retirement_mw:,.1f}")

    st.markdown("---")

    # Generator charts
    col1, col2 = st.columns(2)

    with col1:
        st.subheader(f"📊 Generator Technology Mix ({selected_state})")

        technology_df = (
            state_gen.groupby("Technology", as_index=False)["Capacity_MW"]
            .sum()
        )

        fig_technology = px.bar(
            technology_df,
            x="Technology",
            y="Capacity_MW",
            text="Capacity_MW",
            color="Technology",
            color_discrete_map=TECHNOLOGY_COLORS,
        )
        fig_technology.update_traces(
            texttemplate="%{text:,.1f}",
            textposition="outside",
        )
        fig_technology.update_layout(
            showlegend=False,
            yaxis_title="Capacity (MW)",
            yaxis=dict(tickformat=","),
        )
        st.plotly_chart(fig_technology, use_container_width=True)

    with col2:
        st.subheader(f"🏷️ Generator Status Breakdown ({selected_state})")

        status_df = (
            state_gen.groupby(
                ["Technology", "Status"], as_index=False
            )["Capacity_MW"]
            .sum()
        )

        fig_status = px.bar(
            status_df,
            x="Technology",
            y="Capacity_MW",
            color="Status",
            barmode="group",
            color_discrete_map=STATUS_COLORS,
        )
        apply_number_format(fig_status, "Capacity (MW)")
        st.plotly_chart(fig_status, use_container_width=True)

    # Storage charts and table
    st.markdown("---")
    st.subheader(f"🔋 Storage Assets ({selected_state} - MWh)")

    if not state_stor.empty:
        storage_col1, storage_col2 = st.columns(2)

        with storage_col1:
            storage_technology_df = (
                state_stor.groupby(
                    "Technology", as_index=False
                )["Capacity_MWh"]
                .sum()
            )

            fig_storage_technology = px.bar(
                storage_technology_df,
                x="Technology",
                y="Capacity_MWh",
                text="Capacity_MWh",
                color="Technology",
                color_discrete_map=TECHNOLOGY_COLORS,
                title=(
                    "Storage Energy Capacity by Technology "
                    f"({selected_state})"
                ),
            )
            fig_storage_technology.update_traces(
                texttemplate="%{text:,.1f}",
                textposition="outside",
            )
            fig_storage_technology.update_layout(
                showlegend=False,
                yaxis_title="Capacity (MWh)",
                yaxis=dict(tickformat=","),
            )
            st.plotly_chart(
                fig_storage_technology,
                use_container_width=True,
            )

        with storage_col2:
            storage_status_df = (
                state_stor.groupby(
                    ["Technology", "Status"], as_index=False
                )["Capacity_MWh"]
                .sum()
            )

            fig_storage_status = px.bar(
                storage_status_df,
                x="Technology",
                y="Capacity_MWh",
                color="Status",
                barmode="group",
                color_discrete_map=STATUS_COLORS,
                title=f"Storage Capacity by Status ({selected_state})",
            )
            apply_number_format(fig_storage_status, "Capacity (MWh)")
            st.plotly_chart(
                fig_storage_status,
                use_container_width=True,
            )

        with st.expander("View Storage Data Table"):
            st.dataframe(state_stor, use_container_width=True)
    else:
        st.info(f"No dedicated storage records found for {selected_state}.")


# ---------------------------------------------------------------------
# 8. VIEW 2 — ALL-INDIA OVERVIEW
# ---------------------------------------------------------------------
elif page == "All-India Overview":
    st.title("🌐 All-India FY31 Capacity Overview")
    st.markdown(
        "Aggregated simulation output across all states and union territories."
    )

    total_generation = df_gen["Capacity_MW"].sum()
    total_existing = df_gen.loc[
        df_gen["Status"] == "Existing", "Capacity_MW"
    ].sum()
    total_upcoming = df_gen.loc[
        df_gen["Status"] == "Upcoming", "Capacity_MW"
    ].sum()
    total_retirement = df_gen.loc[
        df_gen["Status"] == "Retirement", "Capacity_MW"
    ].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("National Total Gen (MW)", f"{total_generation:,.1f}")
    c2.metric("National Existing (MW)", f"{total_existing:,.1f}")
    c3.metric("National Upcoming (MW)", f"{total_upcoming:,.1f}")
    c4.metric("National Retirement (MW)", f"{total_retirement:,.1f}")

    st.markdown("---")

    # State-wise stacked generation capacity chart
    st.subheader("📈 State-wise Installed Capacity Breakdown (MW)")

    state_aggregate = (
        df_gen.groupby(
            ["State", "Technology"], as_index=False
        )["Capacity_MW"]
        .sum()
    )

    fig_state = px.bar(
        state_aggregate,
        x="State",
        y="Capacity_MW",
        color="Technology",
        barmode="stack",
        color_discrete_map=TECHNOLOGY_COLORS,
    )
    fig_state.update_layout(
        xaxis_title="State / Region",
        yaxis_title="Installed Capacity (MW)",
        yaxis=dict(tickformat=","),
    )
    st.plotly_chart(fig_state, use_container_width=True)

    # Interactive technology and status explorer
    st.markdown("---")
    st.subheader("🔍 Interactive Technology & Status Explorer")

    filter_col1, filter_col2 = st.columns(2)

    available_technologies = sorted(
        df_gen["Technology"].dropna().unique().tolist()
    )
    available_statuses = sorted(
        df_gen["Status"].dropna().unique().tolist()
    )

    with filter_col1:
        selected_technologies = st.multiselect(
            "Filter by Technology",
            options=available_technologies,
            default=available_technologies,
        )

    with filter_col2:
        selected_statuses = st.multiselect(
            "Filter by Status",
            options=available_statuses,
            default=available_statuses,
        )

    filtered_df = df_gen[
        df_gen["Technology"].isin(selected_technologies)
        & df_gen["Status"].isin(selected_statuses)
    ]

    if not filtered_df.empty:
        visual_col1, visual_col2 = st.columns(2)

        with visual_col1:
            filtered_technology_aggregate = (
                filtered_df.groupby(
                    "Technology", as_index=False
                )["Capacity_MW"]
                .sum()
            )

            fig_filtered_technology = px.bar(
                filtered_technology_aggregate,
                x="Technology",
                y="Capacity_MW",
                text="Capacity_MW",
                color="Technology",
                color_discrete_map=TECHNOLOGY_COLORS,
                title="Filtered Capacity by Technology (MW)",
            )
            fig_filtered_technology.update_traces(
                texttemplate="%{text:,.1f}",
                textposition="outside",
            )
            fig_filtered_technology.update_layout(
                showlegend=False,
                yaxis_title="Capacity (MW)",
                yaxis=dict(tickformat=","),
            )
            st.plotly_chart(
                fig_filtered_technology,
                use_container_width=True,
            )

        with visual_col2:
            filtered_status_aggregate = (
                filtered_df.groupby(
                    "Status", as_index=False
                )["Capacity_MW"]
                .sum()
            )

            fig_filtered_status = px.bar(
                filtered_status_aggregate,
                x="Status",
                y="Capacity_MW",
                text="Capacity_MW",
                color="Status",
                color_discrete_map=STATUS_COLORS,
                title="Filtered Capacity by Status (MW)",
            )
            fig_filtered_status.update_traces(
                texttemplate="%{text:,.1f}",
                textposition="outside",
            )
            apply_number_format(
                fig_filtered_status,
                "Capacity (MW)",
            )
            st.plotly_chart(
                fig_filtered_status,
                use_container_width=True,
            )

        with st.expander("View Underlying Filtered Data Table"):
            st.dataframe(filtered_df, use_container_width=True)
    else:
        st.warning("No records match your selected filters.")
