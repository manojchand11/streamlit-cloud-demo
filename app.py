from pathlib import Path

import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =============================================================================
# 1. PAGE CONFIGURATION
# =============================================================================
st.set_page_config(
    page_title="PLEXOS | State-wise Capacity Dashboard",
    page_icon="⚡",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent

WORKBOOKS = {
    "FY31": BASE_DIR / "S2_Dispatch_FY31_statewise_output.xlsx",
    "FY36": BASE_DIR / "S2_Dispatch_FY36_statewise_output.xlsx",
}

# Manager-approved technology colours and display order.
# Generation technologies are shown first; storage technologies follow.
TECHNOLOGY_ORDER = [
    "Thermal", "Nuclear", "Hydro", "Other RE", "Solar",
    "Wind", "BESS", "PSP", "Other",
]
GENERATION_TECHNOLOGY_ORDER = [
    "Thermal", "Nuclear", "Hydro", "Other RE", "Solar",
    "Wind", "Other",
]
STORAGE_TECHNOLOGY_ORDER = ["BESS", "PSP"]

TECHNOLOGY_COLORS = {
    "Thermal": "#808080",   # Grey
    "Nuclear": "#D62728",   # Red
    "Hydro": "#1F77B4",     # Blue
    "Other RE": "#9467BD",  # Purple
    "Solar": "#FFD700",     # Yellow
    "Wind": "#2E8B57",      # Green
    "BESS": "#C2185B",      # Magenta / deep pink
    "PSP": "#17A2A4",       # Teal / cyan
    "Other": "#A9A9A9",     # Neutral fallback
}

STATUS_COLORS = {
    "Existing": "#4C78A8",
    "Upcoming": "#F2A541",
    "Candidates": "#72B7B2",
    "Retirement": "#E45756",
    "Other": "#A0A0A0",
}

st.title("⚡ PLEXOS | State-wise Capacity Dashboard")
st.caption(
    "Installed generation capacity (MW) and storage energy capacity (MWh) "
    "— FY31 and FY36 scenario analysis"
)


# =============================================================================
# 2. CLASSIFICATION AND FORMATTING HELPERS
# =============================================================================
def classify_generator(category: str) -> tuple[str, str]:
    """Map the source category into a consistent technology/status taxonomy."""
    text = str(category).strip().lower()

    if "retirement" in text:
        status = "Retirement"
    elif "candidate" in text:
        status = "Candidates"
    elif "upcoming" in text:
        status = "Upcoming"
    elif "existing" in text:
        status = "Existing"
    else:
        status = "Other"

    if "nuclear" in text:
        technology = "Nuclear"
    elif "solar" in text:
        technology = "Solar"
    elif "wind" in text:
        technology = "Wind"
    elif "hydro" in text:
        technology = "Hydro"
    elif "dre" in text or "other re" in text:
        technology = "Other RE"
    elif "thermal" in text:
        technology = "Thermal"
    else:
        technology = "Other"

    return technology, status


def classify_storage(category: str) -> tuple[str, str]:
    """Map storage categories while retaining the source's project status."""
    text = str(category).strip().lower()

    if "retirement" in text:
        status = "Retirement"
    elif "candidate" in text:
        status = "Candidates"
    elif "upcoming" in text:
        status = "Upcoming"
    elif "existing" in text:
        status = "Existing"
    else:
        status = "Other"

    technology = "PSP" if "psp" in text else "BESS"
    return technology, status


def to_number(value) -> float:
    """Convert workbook values to numeric values without breaking on blanks."""
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def style_capacity_chart(fig, y_title: str):
    fig.update_layout(
        yaxis_title=y_title,
        yaxis=dict(tickformat=","),
        margin=dict(t=65, b=45),
        legend_title_text="",
        hovermode="x unified",
    )
    return fig


def format_delta(current: float, previous: float) -> str:
    """Return a readable absolute and percentage change."""
    change = current - previous
    if previous == 0:
        pct = "New / no FY31 base" if current != 0 else "0.0%"
    else:
        pct = f"{change / abs(previous) * 100:+.1f}%"
    return f"{change:+,.1f} MW ({pct})"


def add_comparison_columns(
    base: pd.DataFrame,
    future: pd.DataFrame,
    group_columns: list[str],
    value_column: str,
    value_label: str,
) -> pd.DataFrame:
    """Full outer comparison so categories/states present in only one year remain visible."""
    left = (
        base.groupby(group_columns, as_index=False)[value_column]
        .sum()
        .rename(columns={value_column: "FY31"})
    )
    right = (
        future.groupby(group_columns, as_index=False)[value_column]
        .sum()
        .rename(columns={value_column: "FY36"})
    )

    result = left.merge(right, on=group_columns, how="outer").fillna(0)
    result["Change"] = result["FY36"] - result["FY31"]
    result["Change_%"] = result.apply(
        lambda row: (
            (row["Change"] / abs(row["FY31"])) * 100
            if row["FY31"] != 0
            else (None if row["FY36"] != 0 else 0.0)
        ),
        axis=1,
    )
    result["Metric"] = value_label
    return result


# =============================================================================
# 3. WORKBOOK READER
# =============================================================================
@st.cache_data(show_spinner="Reading and processing PLEXOS workbooks...")
def load_and_process_data(file_path: str):
    """
    Read each state's generator and storage Collection tables.

    Generator capacity is measured in MW.
    Storage installed capacity is measured in MWh.
    The parser locates Collection rows dynamically, so different header row
    positions across state worksheets are supported.
    """
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(
            f"Workbook not found: {path.name}. Keep both FY31 and FY36 Excel "
            "workbooks in the same GitHub repository folder as app.py."
        )

    generator_records = []
    storage_records = []
    sheets_without_generator_table = []
    sheets_without_storage_table = []

    workbook = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=True,
    )

    try:
        for state in workbook.sheetnames:
            worksheet = workbook[state]
            rows = list(worksheet.iter_rows(values_only=True))

            collection_rows = [
                i for i, row in enumerate(rows)
                if row and str(row[0]).strip() == "Collection"
            ]

            if not collection_rows:
                sheets_without_generator_table.append(state)
                continue

            # Generator table: first Collection header in the sheet.
            gen_header_idx = collection_rows[0]
            gen_headers = list(rows[gen_header_idx])

            try:
                gen_category_col = gen_headers.index("Category")
                gen_capacity_col = gen_headers.index("Installed Capacity (MW)")
                gen_sent_out_col = gen_headers.index("Generation Sent Out (GWh)")
            except ValueError as exc:
                raise ValueError(
                    f"Generator headers in sheet '{state}' of {path.name} "
                    "do not contain Category, Installed Capacity (MW), and Generation Sent Out (GWh)."
                ) from exc

            gen_end_idx = (
                collection_rows[1] if len(collection_rows) > 1 else len(rows)
            )

            for row in rows[gen_header_idx + 1:gen_end_idx]:
                if (
                    row
                    and len(row) > max(gen_category_col, gen_capacity_col, gen_sent_out_col)
                    and str(row[0]).strip() == "Generator"
                    and row[gen_category_col] is not None
                ):
                    generator_records.append(
                        {
                            "State": state,
                            "Raw_Category": str(row[gen_category_col]).strip(),
                            "Capacity_MW": to_number(row[gen_capacity_col]),
                            "Generation_Sent_Out_GWh": to_number(row[gen_sent_out_col]),
                        }
                    )

            # Storage table: second Collection header, where available.
            if len(collection_rows) < 2:
                sheets_without_storage_table.append(state)
                continue

            stor_header_idx = collection_rows[1]
            stor_headers = list(rows[stor_header_idx])

            try:
                stor_category_col = stor_headers.index("Category")
                stor_capacity_col = stor_headers.index("Installed Capacity (MWh)")
            except ValueError as exc:
                raise ValueError(
                    f"Storage headers in sheet '{state}' of {path.name} "
                    "do not contain Category and Installed Capacity (MWh)."
                ) from exc

            for row in rows[stor_header_idx + 1:]:
                if (
                    row
                    and len(row) > max(stor_category_col, stor_capacity_col)
                    and str(row[0]).strip() in {"Battery", "Storage"}
                    and row[stor_category_col] is not None
                ):
                    storage_records.append(
                        {
                            "State": state,
                            "Raw_Category": str(row[stor_category_col]).strip(),
                            "Capacity_MWh": to_number(row[stor_capacity_col]),
                        }
                    )
    finally:
        workbook.close()

    df_gen = pd.DataFrame(
        generator_records,
        columns=["State", "Raw_Category", "Capacity_MW", "Generation_Sent_Out_GWh"],
    )
    df_stor = pd.DataFrame(
        storage_records,
        columns=["State", "Raw_Category", "Capacity_MWh"],
    )

    if not df_gen.empty:
        classified = df_gen["Raw_Category"].apply(classify_generator)
        df_gen[["Technology", "Status"]] = pd.DataFrame(
            classified.tolist(), index=df_gen.index
        )

    if not df_stor.empty:
        classified = df_stor["Raw_Category"].apply(classify_storage)
        df_stor[["Technology", "Status"]] = pd.DataFrame(
            classified.tolist(), index=df_stor.index
        )

    metadata = {
        "sheets": workbook_sheet_names(path),
        "sheets_without_generator_table": sheets_without_generator_table,
        "sheets_without_storage_table": sheets_without_storage_table,
    }
    return df_gen, df_stor, metadata


def workbook_sheet_names(path: Path) -> list[str]:
    """Read sheet names separately for concise workbook metadata."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


# =============================================================================
# 4. LOAD BOTH YEARS
# =============================================================================
data = {}
load_errors = {}

for year, path in WORKBOOKS.items():
    try:
        data[year] = load_and_process_data(str(path))
    except Exception as exc:
        load_errors[year] = exc

if load_errors:
    for year, error in load_errors.items():
        st.error(f"Could not load {year} data: {error}")
    st.info(
        "Check that both Excel files are committed to the same GitHub "
        "repository folder as app.py, with the exact filenames shown in "
        "the sidebar."
    )
    st.stop()

for year in ("FY31", "FY36"):
    if data[year][0].empty:
        st.error(f"No generator records were found in the {year} workbook.")
        st.stop()

gen31, stor31, meta31 = data["FY31"]
gen36, stor36, meta36 = data["FY36"]



def render_all_state_comparison():
    """Render the All State Overview with FY31 and FY36 shown together."""
    st.subheader("All-state comparison: FY31 vs FY36")
    st.markdown(
        "Compare installed generation capacity, generation sent out and storage "
        "across states for both planning years."
    )

    total31 = gen31["Capacity_MW"].sum()
    total36 = gen36["Capacity_MW"].sum()
    sent31 = gen31["Generation_Sent_Out_GWh"].sum()
    sent36 = gen36["Generation_Sent_Out_GWh"].sum()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("FY31 Installed Capacity", f"{total31:,.1f} MW")
    k2.metric("FY36 Installed Capacity", f"{total36:,.1f} MW", f"{total36-total31:+,.1f} MW")
    k3.metric("FY31 Generation Sent Out", f"{sent31:,.1f} GWh")
    k4.metric("FY36 Generation Sent Out", f"{sent36:,.1f} GWh", f"{sent36-sent31:+,.1f} GWh")

    st.markdown("---")
    st.subheader("State-wise Installed Generation Capacity")
    c1, c2 = st.columns(2)
    for col, year, frame in [(c1, "FY31", gen31), (c2, "FY36", gen36)]:
        with col:
            agg = frame.groupby(["State", "Technology"], as_index=False)["Capacity_MW"].sum()
            fig = px.bar(
                agg, x="State", y="Capacity_MW", color="Technology", barmode="stack",
                color_discrete_map=TECHNOLOGY_COLORS,
                category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                title=f"Installed Capacity by State — {year}",
            )
            fig.update_layout(xaxis_title="State / Region", yaxis_title="Installed Capacity (MW)",
                              yaxis=dict(tickformat=","), legend_title_text="", hovermode="x unified")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("State-wise Generation Sent Out")
    c1, c2 = st.columns(2)
    for col, year, frame in [(c1, "FY31", gen31), (c2, "FY36", gen36)]:
        with col:
            agg = frame.groupby(["State", "Technology"], as_index=False)["Generation_Sent_Out_GWh"].sum()
            order = sorted(agg["State"].dropna().unique().tolist())
            fig = px.bar(
                agg, x="State", y="Generation_Sent_Out_GWh", color="Technology", barmode="stack",
                color_discrete_map=TECHNOLOGY_COLORS,
                category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER, "State": order},
                title=f"Generation Sent Out by State — {year}",
            )
            fig.update_layout(
                xaxis_title="State / Region", yaxis_title="Generation Sent Out (GWh)",
                xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=order),
                yaxis=dict(tickformat=","), legend_title_text="",
                height=520, margin=dict(b=110, t=65),
            )
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("National Generation Capacity Mix")
    c1, c2 = st.columns(2)
    for col, year, frame in [(c1, "FY31", gen31), (c2, "FY36", gen36)]:
        with col:
            agg = frame.groupby("Technology", as_index=False)["Capacity_MW"].sum()
            fig = px.pie(agg, names="Technology", values="Capacity_MW", color="Technology",
                         color_discrete_map=TECHNOLOGY_COLORS,
                         category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                         hole=0.42, title=f"Capacity Mix — {year}")
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("Storage Portfolio Comparison")
    c1, c2 = st.columns(2)
    for col, year, frame in [(c1, "FY31", stor31), (c2, "FY36", stor36)]:
        with col:
            if frame.empty:
                st.info(f"No storage records found in {year}.")
                continue
            agg = frame.groupby("Technology", as_index=False)["Capacity_MWh"].sum()
            fig = px.bar(agg, x="Technology", y="Capacity_MWh", color="Technology",
                         color_discrete_map=TECHNOLOGY_COLORS,
                         category_orders={"Technology": STORAGE_TECHNOLOGY_ORDER},
                         title=f"Storage Energy Capacity — {year}")
            style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
            st.plotly_chart(fig, use_container_width=True)


# =============================================================================
# 5. SIDEBAR AND NAVIGATION
# =============================================================================
st.sidebar.title("⚡ PLEXOS Dashboard")

view = st.sidebar.radio(
    "Select Analysis",
    [
        "State Explorer",
        "All State Overview",
        "All India Overview",
    ],
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Source workbooks**")
st.sidebar.caption("FY31: S2_Dispatch_FY31_statewise_output.xlsx")
st.sidebar.caption("FY36: S2_Dispatch_FY36_statewise_output.xlsx")

st.sidebar.markdown("---")
st.sidebar.markdown("### Technology colour key")
for technology, colour in TECHNOLOGY_COLORS.items():
    st.sidebar.markdown(
        f'<div style="display:flex;align-items:center;margin:4px 0;">'
        f'<span style="background:{colour};width:13px;height:13px;'
        f'display:inline-block;border:1px solid #777;margin-right:8px;'
        f'border-radius:2px;"></span>{technology}</div>',
        unsafe_allow_html=True,
    )

st.sidebar.markdown("---")
st.sidebar.caption(
    "Interpretation: generation installed capacity is reported in MW; "
    "storage energy capacity is reported in MWh. These are different units "
    "and are not added together."
)


# =============================================================================
# 6. VIEW — STATE EXPLORER
# =============================================================================
if view == "State Explorer":
    st.header("🗺️ State Explorer")
    st.markdown(
        "Explore the generation portfolio and storage position of an "
        "individual state or region. Switch between FY31 and FY36, or view "
        "both years side by side."
    )

    year_choice = st.radio(
        "Select financial year",
        ["FY31", "FY36", "Compare FY31 and FY36"],
        horizontal=True,
        key="state_year_choice",
    )

    state_options = sorted(
        set(gen31["State"].dropna().unique())
        | set(gen36["State"].dropna().unique())
    )
    selected_state = st.selectbox("Select State / Region", state_options)

    def render_state_year(year: str):
        gdf = data[year][0]
        sdf = data[year][1]
        state_gen = gdf[gdf["State"] == selected_state].copy()
        state_stor = sdf[sdf["State"] == selected_state].copy()

        total = state_gen["Capacity_MW"].sum()
        existing = state_gen.loc[
            state_gen["Status"] == "Existing", "Capacity_MW"
        ].sum()
        upcoming = state_gen.loc[
            state_gen["Status"] == "Upcoming", "Capacity_MW"
        ].sum()
        candidates = state_gen.loc[
            state_gen["Status"] == "Candidates", "Capacity_MW"
        ].sum()
        retirement = state_gen.loc[
            state_gen["Status"] == "Retirement", "Capacity_MW"
        ].sum()

        st.subheader(f"{selected_state} — {year}")
        a, b, c, d, e = st.columns(5)
        a.metric("Total Generation", f"{total:,.1f} MW")
        b.metric("Existing", f"{existing:,.1f} MW")
        c.metric("Upcoming", f"{upcoming:,.1f} MW")
        d.metric("Candidates", f"{candidates:,.1f} MW")
        e.metric("Retirement", f"{retirement:,.1f} MW")

        left, right = st.columns(2)

        with left:
            tech = (
                state_gen.groupby("Technology", as_index=False)["Capacity_MW"]
                .sum()
                .sort_values("Capacity_MW", ascending=False)
            )
            fig = px.bar(
                tech,
                x="Technology",
                y="Capacity_MW",
                color="Technology",
                text="Capacity_MW",
                color_discrete_map=TECHNOLOGY_COLORS,
                category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                title=f"Generation Capacity Mix — {year}",
            )
            fig.update_traces(
                texttemplate="%{text:,.0f}",
                textposition="outside",
            )
            fig.update_layout(showlegend=False)
            style_capacity_chart(fig, "Installed Capacity (MW)")
            st.plotly_chart(fig, use_container_width=True)

        with right:
            status = (
                state_gen.groupby(
                    ["Technology", "Status"], as_index=False
                )["Capacity_MW"].sum()
            )
            fig = px.bar(
                status,
                x="Technology",
                y="Capacity_MW",
                color="Status",
                barmode="stack",
                color_discrete_map=STATUS_COLORS,
                category_orders={
                    "Technology": GENERATION_TECHNOLOGY_ORDER,
                    "Status": ["Existing", "Upcoming", "Candidates", "Retirement", "Other"],
                },
                title=f"Project Status by Technology — {year}",
            )
            style_capacity_chart(fig, "Installed Capacity (MW)")
            st.plotly_chart(fig, use_container_width=True)

        # Generation Sent Out for the selected state and financial year.
        st.markdown("---")
        st.subheader(f"Generation Sent Out — {year}")
        sent_out_total = state_gen["Generation_Sent_Out_GWh"].sum()
        st.metric(
            f"{selected_state} — Generation Sent Out",
            f"{sent_out_total:,.1f} GWh",
        )

        sent_out_technology = (
            state_gen.groupby("Technology", as_index=False)["Generation_Sent_Out_GWh"]
            .sum()
        )
        sent_out_technology = sent_out_technology[
            sent_out_technology["Generation_Sent_Out_GWh"] != 0
        ]
        sent_out_status = (
            state_gen.groupby(["Technology", "Status"], as_index=False)["Generation_Sent_Out_GWh"]
            .sum()
        )
        sent_out_status = sent_out_status[
            sent_out_status["Generation_Sent_Out_GWh"] != 0
        ]
        if not sent_out_technology.empty:
            sent_mix_col, sent_status_col = st.columns(2)
            with sent_mix_col:
                fig = px.bar(
                    sent_out_technology,
                    x="Technology",
                    y="Generation_Sent_Out_GWh",
                    color="Technology",
                    text="Generation_Sent_Out_GWh",
                    color_discrete_map=TECHNOLOGY_COLORS,
                    category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                    title=f"Generation Sent Out Mix — {selected_state}, {year}",
                )
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                fig.update_layout(showlegend=False)
                style_capacity_chart(fig, "Generation Sent Out (GWh)")
                st.plotly_chart(fig, use_container_width=True)
            with sent_status_col:
                fig = px.bar(
                    sent_out_status,
                    x="Technology",
                    y="Generation_Sent_Out_GWh",
                    color="Status",
                    barmode="stack",
                    color_discrete_map=STATUS_COLORS,
                    category_orders={
                        "Technology": GENERATION_TECHNOLOGY_ORDER,
                        "Status": ["Existing", "Upcoming", "Candidates", "Retirement", "Other"],
                    },
                    title=f"Generation Sent Out by Project Status — {selected_state}, {year}",
                )
                style_capacity_chart(fig, "Generation Sent Out (GWh)")
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"No Generation Sent Out data is available for {selected_state} in {year}.")

        st.markdown("**Storage portfolio**")
        if not state_stor.empty:
            x, y = st.columns(2)
            with x:
                stor_mix = (
                    state_stor.groupby("Technology", as_index=False)["Capacity_MWh"]
                    .sum()
                )
                fig = px.bar(
                    stor_mix,
                    x="Technology",
                    y="Capacity_MWh",
                    color="Technology",
                    text="Capacity_MWh",
                    color_discrete_map=TECHNOLOGY_COLORS,
                    category_orders={"Technology": STORAGE_TECHNOLOGY_ORDER},
                    title=f"Storage Energy Capacity — {year}",
                )
                fig.update_traces(
                    texttemplate="%{text:,.0f}",
                    textposition="outside",
                )
                fig.update_layout(showlegend=False)
                style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
                st.plotly_chart(fig, use_container_width=True)

            with y:
                stor_status = (
                    state_stor.groupby(
                        ["Technology", "Status"], as_index=False
                    )["Capacity_MWh"].sum()
                )
                fig = px.bar(
                    stor_status,
                    x="Technology",
                    y="Capacity_MWh",
                    color="Status",
                    barmode="group",
                    color_discrete_map=STATUS_COLORS,
                    category_orders={
                        "Technology": STORAGE_TECHNOLOGY_ORDER,
                        "Status": ["Existing", "Upcoming", "Candidates", "Retirement", "Other"],
                    },
                    title=f"Storage Status — {year}",
                )
                style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"No storage records are available for {selected_state} in {year}.")

        with st.expander(f"View {year} source-category data"):
            st.dataframe(
                state_gen.sort_values(["Technology", "Status"]),
                use_container_width=True,
            )
            if not state_stor.empty:
                st.dataframe(state_stor, use_container_width=True)

        return state_gen, state_stor

    if year_choice == "Compare FY31 and FY36":
        # Dedicated comparison view: do not repeat the two full State Explorer pages.
        state_gen31 = gen31[gen31["State"] == selected_state].copy()
        state_gen36 = gen36[gen36["State"] == selected_state].copy()
        state_stor31 = stor31[stor31["State"] == selected_state].copy()
        state_stor36 = stor36[stor36["State"] == selected_state].copy()

        st.subheader(f"{selected_state} — FY31 vs FY36")

        # 1. Generation capacity summary for both years.
        def status_capacity(df, status):
            return df.loc[df["Status"] == status, "Capacity_MW"].sum()

        capacity_summary = []
        for label, selector in [
            ("Total Generation Capacity", None),
            ("Existing", "Existing"),
            ("Upcoming", "Upcoming"),
            ("Candidates", "Candidates"),
            ("Retirement", "Retirement"),
        ]:
            value31 = (
                state_gen31["Capacity_MW"].sum()
                if selector is None else status_capacity(state_gen31, selector)
            )
            value36 = (
                state_gen36["Capacity_MW"].sum()
                if selector is None else status_capacity(state_gen36, selector)
            )
            capacity_summary.append({
                "Metric": label,
                "FY31 (MW)": value31,
                "FY36 (MW)": value36,
                "Change (MW)": value36 - value31,
                "Change (%)": ((value36 - value31) / abs(value31) * 100)
                    if value31 != 0 else (None if value36 != 0 else 0.0),
            })

        st.markdown("**Generation capacity summary**")
        summary_df = pd.DataFrame(capacity_summary)
        st.dataframe(
            summary_df.style.format({
                "FY31 (MW)": "{:,.1f}",
                "FY36 (MW)": "{:,.1f}",
                "Change (MW)": "{:+,.1f}",
                "Change (%)": lambda value: "—" if pd.isna(value) else f"{value:+.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")
        st.subheader("Generation Portfolio")

        # 2A. Generation capacity mix — one chart per year, side by side.
        col31, col36 = st.columns(2)
        for col, year, frame in [
            (col31, "FY31", state_gen31),
            (col36, "FY36", state_gen36),
        ]:
            with col:
                tech_mix = (
                    frame.groupby("Technology", as_index=False)["Capacity_MW"]
                    .sum()
                )
                fig = px.bar(
                    tech_mix,
                    x="Technology",
                    y="Capacity_MW",
                    color="Technology",
                    text="Capacity_MW",
                    color_discrete_map=TECHNOLOGY_COLORS,
                    category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                    title=f"Generation Capacity Mix — {year}",
                )
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                fig.update_layout(showlegend=False)
                style_capacity_chart(fig, "Installed Capacity (MW)")
                st.plotly_chart(fig, use_container_width=True)

        # 2B. Project status by technology — stacked bars for each year.
        col31, col36 = st.columns(2)
        status_order = ["Existing", "Upcoming", "Candidates", "Retirement", "Other"]
        for col, year, frame in [
            (col31, "FY31", state_gen31),
            (col36, "FY36", state_gen36),
        ]:
            with col:
                status_mix = (
                    frame.groupby(["Technology", "Status"], as_index=False)["Capacity_MW"]
                    .sum()
                )
                fig = px.bar(
                    status_mix,
                    x="Technology",
                    y="Capacity_MW",
                    color="Status",
                    barmode="stack",
                    color_discrete_map=STATUS_COLORS,
                    category_orders={
                        "Technology": GENERATION_TECHNOLOGY_ORDER,
                        "Status": status_order,
                    },
                    title=f"Project Status by Technology — {year}",
                )
                style_capacity_chart(fig, "Installed Capacity (MW)")
                st.plotly_chart(fig, use_container_width=True)

        # 3. Generation sent out — aggregate KPI plus technology-wise comparison.
        st.markdown("---")
        st.subheader("Generation Sent Out")
        sent31 = state_gen31["Generation_Sent_Out_GWh"].sum()
        sent36 = state_gen36["Generation_Sent_Out_GWh"].sum()
        sent_change = sent36 - sent31
        sent_pct = sent_change / abs(sent31) * 100 if sent31 else None
        k1, k2, k3 = st.columns(3)
        k1.metric("FY31 Generation Sent Out", f"{sent31:,.1f} GWh")
        k2.metric("FY36 Generation Sent Out", f"{sent36:,.1f} GWh")
        k3.metric(
            "Change in Generation Sent Out",
            f"{sent_change:+,.1f} GWh",
            delta=f"{sent_pct:+.1f}%" if sent_pct is not None else "No FY31 base",
        )

        sent_comp = add_comparison_columns(
            state_gen31, state_gen36, ["Technology"],
            "Generation_Sent_Out_GWh", "Generation sent out (GWh)"
        )
        sent_comp["Technology"] = pd.Categorical(
            sent_comp["Technology"],
            categories=GENERATION_TECHNOLOGY_ORDER,
            ordered=True,
        )
        sent_comp = sent_comp.sort_values("Technology")
        fig = px.bar(
            sent_comp,
            x="Technology",
            y=["FY31", "FY36"],
            barmode="group",
            color_discrete_sequence=["#4C78A8", "#F2A541"],
            category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
            title="Generation Sent Out by Technology — FY31 vs FY36",
        )
        style_capacity_chart(fig, "Generation Sent Out (GWh)")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(
            sent_comp[["Technology", "FY31", "FY36", "Change", "Change_%"]]
            .rename(columns={"Change": "Change (GWh)", "Change_%": "Change (%)"})
            .style.format({
                "FY31": "{:,.1f}", "FY36": "{:,.1f}",
                "Change (GWh)": "{:+,.1f}",
                "Change (%)": lambda value: "—" if pd.isna(value) else f"{value:+.1f}%",
            }),
            use_container_width=True,
            hide_index=True,
        )

        # 4. Storage portfolio — side-by-side FY31 and FY36 charts.
        st.markdown("---")
        st.subheader("Storage Portfolio")
        storage_summary = pd.DataFrame([
            {
                "Metric": "Total Storage Energy Capacity",
                "FY31 (MWh)": state_stor31["Capacity_MWh"].sum(),
                "FY36 (MWh)": state_stor36["Capacity_MWh"].sum(),
            },
        ])
        storage_summary["Change (MWh)"] = (
            storage_summary["FY36 (MWh)"] - storage_summary["FY31 (MWh)"]
        )
        st.dataframe(
            storage_summary.style.format({
                "FY31 (MWh)": "{:,.1f}", "FY36 (MWh)": "{:,.1f}",
                "Change (MWh)": "{:+,.1f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        # Storage energy capacity mix by year.
        col31, col36 = st.columns(2)
        for col, year, frame in [
            (col31, "FY31", state_stor31),
            (col36, "FY36", state_stor36),
        ]:
            with col:
                if frame.empty:
                    st.info(f"No storage records are available for {year}.")
                    continue
                stor_mix = (
                    frame.groupby("Technology", as_index=False)["Capacity_MWh"]
                    .sum()
                )
                fig = px.bar(
                    stor_mix,
                    x="Technology",
                    y="Capacity_MWh",
                    color="Technology",
                    text="Capacity_MWh",
                    color_discrete_map=TECHNOLOGY_COLORS,
                    category_orders={"Technology": STORAGE_TECHNOLOGY_ORDER},
                    title=f"Storage Energy Capacity — {year}",
                )
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                fig.update_layout(showlegend=False)
                style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
                st.plotly_chart(fig, use_container_width=True)

        # Storage status by technology — stacked bars by year.
        col31, col36 = st.columns(2)
        for col, year, frame in [
            (col31, "FY31", state_stor31),
            (col36, "FY36", state_stor36),
        ]:
            with col:
                if frame.empty:
                    st.info(f"No storage status records are available for {year}.")
                    continue
                stor_status = (
                    frame.groupby(["Technology", "Status"], as_index=False)["Capacity_MWh"]
                    .sum()
                )
                fig = px.bar(
                    stor_status,
                    x="Technology",
                    y="Capacity_MWh",
                    color="Status",
                    barmode="stack",
                    color_discrete_map=STATUS_COLORS,
                    category_orders={
                        "Technology": STORAGE_TECHNOLOGY_ORDER,
                        "Status": status_order,
                    },
                    title=f"Storage Portfolio by Project Status — {year}",
                )
                style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
                st.plotly_chart(fig, use_container_width=True)

        st.caption(
            "Generation Sent Out is taken from the corresponding state worksheet "
            "and aggregated across generator categories. Generation capacity is "
            "reported in MW; storage energy capacity is reported in MWh."
        )

        with st.expander("View FY31 and FY36 source-category records"):
            st.markdown("**FY31 generation records**")
            st.dataframe(state_gen31, use_container_width=True)
            st.markdown("**FY36 generation records**")
            st.dataframe(state_gen36, use_container_width=True)
            if not state_stor31.empty:
                st.markdown("**FY31 storage records**")
                st.dataframe(state_stor31, use_container_width=True)
            if not state_stor36.empty:
                st.markdown("**FY36 storage records**")
                st.dataframe(state_stor36, use_container_width=True)
    else:
        render_state_year(year_choice)

# =============================================================================
# 7. VIEW — ALL-INDIA OVERVIEW
# =============================================================================
elif view == "All State Overview":
    st.header("🌐 All State Overview")
    st.markdown(
        "National portfolio view of installed generation capacity, project "
        "pipeline and storage energy capacity."
    )

    year_choice = st.radio(
        "Select financial year",
        ["FY31", "FY36", "Compare FY31 and FY36"],
        horizontal=True,
        key="national_year_choice",
    )
    if year_choice == "Compare FY31 and FY36":
        render_all_state_comparison()
        st.stop()

    df_gen, df_stor, metadata = data[year_choice]

    total = df_gen["Capacity_MW"].sum()
    existing = df_gen.loc[df_gen["Status"] == "Existing", "Capacity_MW"].sum()
    upcoming = df_gen.loc[df_gen["Status"] == "Upcoming", "Capacity_MW"].sum()
    candidates = df_gen.loc[df_gen["Status"] == "Candidates", "Capacity_MW"].sum()
    retirement = df_gen.loc[df_gen["Status"] == "Retirement", "Capacity_MW"].sum()

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total Generation", f"{total:,.1f} MW")
    k2.metric("Existing", f"{existing:,.1f} MW")
    k3.metric("Upcoming", f"{upcoming:,.1f} MW")
    k4.metric("Candidates", f"{candidates:,.1f} MW")
    k5.metric("Retirement", f"{retirement:,.1f} MW")

    st.markdown("---")
    st.subheader(f"State-wise Installed Generation Capacity — {year_choice}")

    state_agg = (
        df_gen.groupby(["State", "Technology"], as_index=False)["Capacity_MW"]
        .sum()
    )
    fig = px.bar(
        state_agg,
        x="State",
        y="Capacity_MW",
        color="Technology",
        barmode="stack",
        color_discrete_map=TECHNOLOGY_COLORS,
        category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
    )
    fig.update_layout(
        xaxis_title="State / Region",
        yaxis_title="Installed Capacity (MW)",
        yaxis=dict(tickformat=","),
        legend_title_text="Technology",
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True)

    # State-wise Generation Sent Out for the selected financial year.
    st.markdown("---")
    st.subheader(f"State-wise Generation Sent Out (GWh) — {year_choice}")
    sent_out_state = (
        df_gen.groupby(["State", "Technology"], as_index=False)[
            "Generation_Sent_Out_GWh"
        ].sum()
    )
    state_order_sent = sorted(sent_out_state["State"].dropna().unique().tolist())
    fig = px.bar(
        sent_out_state,
        x="State",
        y="Generation_Sent_Out_GWh",
        color="Technology",
        barmode="stack",
        color_discrete_map=TECHNOLOGY_COLORS,
        category_orders={
            "Technology": GENERATION_TECHNOLOGY_ORDER,
            "State": state_order_sent,
        },
        title=f"State-wise Generation Sent Out by Technology — {year_choice}",
    )
    fig.update_layout(
        xaxis_title="State / Region",
        yaxis_title="Generation Sent Out (GWh)",
        xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=state_order_sent),
        yaxis=dict(tickformat=","),
        legend_title_text="Technology",
        hovermode="x unified",
        height=520,
        margin=dict(b=110, t=65),
    )
    st.plotly_chart(fig, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        tech_agg = (
            df_gen.groupby("Technology", as_index=False)["Capacity_MW"]
            .sum()
            .sort_values("Capacity_MW", ascending=False)
        )
        fig = px.pie(
            tech_agg,
            names="Technology",
            values="Capacity_MW",
            color="Technology",
            color_discrete_map=TECHNOLOGY_COLORS,
            category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
            hole=0.45,
            title=f"National Generation Technology Mix — {year_choice}",
        )
        fig.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        status_agg = (
            df_gen.groupby("Status", as_index=False)["Capacity_MW"]
            .sum()
        )
        fig = px.bar(
            status_agg,
            x="Status",
            y="Capacity_MW",
            color="Status",
            text="Capacity_MW",
            color_discrete_map=STATUS_COLORS,
            title=f"National Capacity by Project Status — {year_choice}",
        )
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        style_capacity_chart(fig, "Installed Capacity (MW)")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("Storage Portfolio")
    if not df_stor.empty:
        s1, s2, s3 = st.columns(3)
        storage_total = df_stor["Capacity_MWh"].sum()
        s1.metric("Total Storage Energy Capacity", f"{storage_total:,.1f} MWh")
        s1.caption("MWh is energy capacity; it is not directly additive to MW.")

        # Combined state-wise storage portfolio: all states in one chart, alphabetically.
        st.markdown("**State-wise Storage Portfolio**")
        state_storage = (
            df_stor.groupby(["State", "Technology"], as_index=False)["Capacity_MWh"]
            .sum()
        )
        storage_state_order = sorted(state_storage["State"].dropna().unique().tolist())
        fig = px.bar(
            state_storage,
            x="State",
            y="Capacity_MWh",
            color="Technology",
            barmode="stack",
            color_discrete_map=TECHNOLOGY_COLORS,
            category_orders={
                "Technology": STORAGE_TECHNOLOGY_ORDER,
                "State": storage_state_order,
            },
            title=f"State-wise Storage Energy Capacity — {year_choice}",
        )
        fig.update_layout(
            xaxis_title="State / Region",
            yaxis_title="Installed Energy Capacity (MWh)",
            xaxis=dict(tickangle=-45, categoryorder="array", categoryarray=storage_state_order),
            yaxis=dict(tickformat=","),
            legend_title_text="Storage Technology",
            hovermode="x unified",
            height=520,
            margin=dict(b=110, t=65),
        )
        st.plotly_chart(fig, use_container_width=True)

        stor_tech = (
            df_stor.groupby("Technology", as_index=False)["Capacity_MWh"].sum()
        )
        fig = px.bar(
            stor_tech,
            x="Technology",
            y="Capacity_MWh",
            color="Technology",
            text="Capacity_MWh",
            color_discrete_map=TECHNOLOGY_COLORS,
            category_orders={"Technology": STORAGE_TECHNOLOGY_ORDER},
            title=f"National Storage Energy Capacity — {year_choice}",
        )
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig.update_layout(showlegend=False)
        style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
        st.plotly_chart(fig, use_container_width=True)

        stor_status = (
            df_stor.groupby(["Technology", "Status"], as_index=False)["Capacity_MWh"]
            .sum()
        )
        fig = px.bar(
            stor_status,
            x="Technology",
            y="Capacity_MWh",
            color="Status",
            barmode="stack",
            color_discrete_map=STATUS_COLORS,
            category_orders={
                "Technology": STORAGE_TECHNOLOGY_ORDER,
                "Status": ["Existing", "Upcoming", "Candidates", "Retirement", "Other"],
            },
            title=f"Storage Energy Capacity by Status — {year_choice}",
        )
        style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info(f"No storage records found in {year_choice}.")

    st.markdown("---")
    st.subheader("Interactive Portfolio Filter")
    f1, f2 = st.columns(2)
    technologies = sorted(df_gen["Technology"].dropna().unique())
    statuses = sorted(df_gen["Status"].dropna().unique())
    with f1:
        selected_techs = st.multiselect(
            "Technology", technologies, default=technologies, key="national_tech_filter"
        )
    with f2:
        selected_status = st.multiselect(
            "Project status", statuses, default=statuses, key="national_status_filter"
        )

    filtered = df_gen[
        df_gen["Technology"].isin(selected_techs)
        & df_gen["Status"].isin(selected_status)
    ]
    if not filtered.empty:
        filtered_state = (
            filtered.groupby(["State", "Technology"], as_index=False)["Capacity_MW"]
            .sum()
        )
        fig = px.bar(
            filtered_state,
            x="State",
            y="Capacity_MW",
            color="Technology",
            barmode="stack",
            color_discrete_map=TECHNOLOGY_COLORS,
            category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
            title="Filtered State-wise Capacity (MW)",
        )
        style_capacity_chart(fig, "Installed Capacity (MW)")
        st.plotly_chart(fig, use_container_width=True)
        with st.expander("View filtered records"):
            st.dataframe(filtered, use_container_width=True)
    else:
        st.warning("No records match the selected filters.")

    with st.expander("Data coverage and workbook checks"):
        st.write(f"Worksheets read: {len(metadata['sheets'])}")
        st.write(f"Generator records: {len(df_gen):,}")
        st.write(f"Storage records: {len(df_stor):,}")
        st.write(
            "Worksheets without a storage table: "
            + (
                ", ".join(metadata["sheets_without_storage_table"])
                if metadata["sheets_without_storage_table"] else "None"
            )
        )


# =============================================================================
# 8. VIEW — FY31 VS FY36 COMPARISON
# =============================================================================
elif view == "All India Overview":
    st.header("🌐 All India Overview")
    st.markdown(
        "National-level generation capacity, generation sent out and storage portfolio."
    )
    india_year_choice = st.radio(
        "Select financial year",
        ["FY31", "FY36", "Compare FY31 and FY36"],
        horizontal=True,
        key="india_overview_year_choice",
    )

    if india_year_choice != "Compare FY31 and FY36":
        india_gen, india_stor, india_metadata = data[india_year_choice]
        india_total = india_gen["Capacity_MW"].sum()
        india_existing = india_gen.loc[india_gen["Status"] == "Existing", "Capacity_MW"].sum()
        india_upcoming = india_gen.loc[india_gen["Status"] == "Upcoming", "Capacity_MW"].sum()
        india_candidates = india_gen.loc[india_gen["Status"] == "Candidates", "Capacity_MW"].sum()
        india_retirement = india_gen.loc[india_gen["Status"] == "Retirement", "Capacity_MW"].sum()
        india_sent_out = india_gen["Generation_Sent_Out_GWh"].sum()

        st.subheader(f"National Generation Capacity — {india_year_choice}")
        ik1, ik2, ik3, ik4, ik5 = st.columns(5)
        ik1.metric("Total Generation", f"{india_total:,.1f} MW")
        ik2.metric("Existing", f"{india_existing:,.1f} MW")
        ik3.metric("Upcoming", f"{india_upcoming:,.1f} MW")
        ik4.metric("Candidates", f"{india_candidates:,.1f} MW")
        ik5.metric("Retirement", f"{india_retirement:,.1f} MW")

        india_mix, india_status = st.columns(2)
        with india_mix:
            mix = india_gen.groupby("Technology", as_index=False)["Capacity_MW"].sum()
            fig = px.pie(
                mix, names="Technology", values="Capacity_MW", color="Technology",
                color_discrete_map=TECHNOLOGY_COLORS,
                category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                hole=0.42, title=f"National Generation Capacity Mix — {india_year_choice}",
            )
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)
        with india_status:
            capacity_status_tech = india_gen.groupby(["Technology", "Status"], as_index=False)["Capacity_MW"].sum()
            fig = px.bar(
                capacity_status_tech, x="Technology", y="Capacity_MW", color="Status",
                barmode="stack", color_discrete_map=STATUS_COLORS,
                category_orders={
                    "Technology": GENERATION_TECHNOLOGY_ORDER,
                    "Status": ["Existing", "Upcoming", "Candidates", "Retirement", "Other"],
                },
                title=f"Project Status by Technology — {india_year_choice}",
            )
            style_capacity_chart(fig, "Installed Capacity (MW)")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.subheader(f"Generation Sent Out — {india_year_choice}")
        sg1, sg2, sg3 = st.columns(3)
        sg1.metric("Total Generation Sent Out", f"{india_sent_out:,.1f} GWh")
        sent_mix = india_gen.groupby("Technology", as_index=False)["Generation_Sent_Out_GWh"].sum()
        sent_mix = sent_mix[sent_mix["Generation_Sent_Out_GWh"] != 0]
        sent_status = india_gen.groupby(["Technology", "Status"], as_index=False)["Generation_Sent_Out_GWh"].sum()
        sent_status = sent_status[sent_status["Generation_Sent_Out_GWh"] != 0]
        if not sent_mix.empty:
            smix_col, sstatus_col = st.columns(2)
            with smix_col:
                fig = px.bar(
                    sent_mix, x="Technology", y="Generation_Sent_Out_GWh", color="Technology",
                    text="Generation_Sent_Out_GWh", color_discrete_map=TECHNOLOGY_COLORS,
                    category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                    title=f"Generation Sent Out Mix — {india_year_choice}",
                )
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                fig.update_layout(showlegend=False)
                style_capacity_chart(fig, "Generation Sent Out (GWh)")
                st.plotly_chart(fig, use_container_width=True)
            with sstatus_col:
                fig = px.bar(
                    sent_status, x="Technology", y="Generation_Sent_Out_GWh", color="Status",
                    barmode="stack", color_discrete_map=STATUS_COLORS,
                    category_orders={
                        "Technology": GENERATION_TECHNOLOGY_ORDER,
                        "Status": ["Existing", "Upcoming", "Candidates", "Retirement", "Other"],
                    },
                    title=f"Generation Sent Out by Project Status — {india_year_choice}",
                )
                style_capacity_chart(fig, "Generation Sent Out (GWh)")
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.subheader(f"Storage Portfolio — {india_year_choice}")
        if not india_stor.empty:
            storage_total = india_stor["Capacity_MWh"].sum()
            st.metric("Total Storage Energy Capacity", f"{storage_total:,.1f} MWh")
            storage_mix = india_stor.groupby("Technology", as_index=False)["Capacity_MWh"].sum()
            fig = px.bar(
                storage_mix, x="Technology", y="Capacity_MWh", color="Technology",
                text="Capacity_MWh", color_discrete_map=TECHNOLOGY_COLORS,
                category_orders={"Technology": STORAGE_TECHNOLOGY_ORDER},
                title=f"National Storage Energy Capacity — {india_year_choice}",
            )
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
            fig.update_layout(showlegend=False)
            style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
            st.plotly_chart(fig, use_container_width=True)
            storage_status = india_stor.groupby(["Technology", "Status"], as_index=False)["Capacity_MWh"].sum()
            fig = px.bar(
                storage_status, x="Technology", y="Capacity_MWh", color="Status",
                barmode="stack", color_discrete_map=STATUS_COLORS,
                category_orders={
                    "Technology": STORAGE_TECHNOLOGY_ORDER,
                    "Status": ["Existing", "Upcoming", "Candidates", "Retirement", "Other"],
                },
                title=f"Storage Portfolio by Project Status — {india_year_choice}",
            )
            style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"No storage records are available in {india_year_choice}.")

        with st.expander("View national source-category records"):
            st.dataframe(india_gen, use_container_width=True)
            if not india_stor.empty:
                st.dataframe(india_stor, use_container_width=True)
        st.stop()

    st.header("🌐 All India Overview | FY31 vs FY36 Comparison")
    st.markdown(
        "A like-for-like comparison of the modelled installed capacity position "
        "between the two planning years. Use the technology, status and state "
        "views together to understand both the scale and composition of change."
    )

    total31 = gen31["Capacity_MW"].sum()
    total36 = gen36["Capacity_MW"].sum()
    change_total = total36 - total31
    change_pct = change_total / abs(total31) * 100 if total31 else None

    existing31 = gen31.loc[gen31["Status"] == "Existing", "Capacity_MW"].sum()
    existing36 = gen36.loc[gen36["Status"] == "Existing", "Capacity_MW"].sum()
    pipeline31 = gen31.loc[
        gen31["Status"].isin(["Upcoming", "Candidates"]), "Capacity_MW"
    ].sum()
    pipeline36 = gen36.loc[
        gen36["Status"].isin(["Upcoming", "Candidates"]), "Capacity_MW"
    ].sum()
    retirement31 = gen31.loc[
        gen31["Status"] == "Retirement", "Capacity_MW"
    ].sum()
    retirement36 = gen36.loc[
        gen36["Status"] == "Retirement", "Capacity_MW"
    ].sum()

    st.subheader("National capacity position")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("FY31 Installed Capacity", f"{total31:,.1f} MW")
    m2.metric("FY36 Installed Capacity", f"{total36:,.1f} MW")
    m3.metric(
        "Net Change, FY31–FY36",
        f"{change_total:+,.1f} MW",
        delta=(f"{change_pct:+.1f}%" if change_pct is not None else "No FY31 base"),
    )
    m4.metric(
        "Existing Capacity Change",
        f"{existing36 - existing31:+,.1f} MW",
        delta=(
            f"{(existing36-existing31)/abs(existing31)*100:+.1f}%"
            if existing31 else "No FY31 base"
        ),
    )

    st.caption(
        "The net change is FY36 minus FY31. It represents the difference "
        "between the two modelled capacity snapshots, not a measure of "
        "electricity generated over the period."
    )

    # Side-by-side capacity mix pies make the change in portfolio composition visible.
    st.markdown("---")
    st.subheader("Installed Generation Capacity Mix — FY31 vs FY36")
    pie31, pie36 = st.columns(2)
    for col, year, frame in [(pie31, "FY31", gen31), (pie36, "FY36", gen36)]:
        with col:
            mix = frame.groupby("Technology", as_index=False)["Capacity_MW"].sum()
            fig = px.pie(
                mix, names="Technology", values="Capacity_MW", color="Technology",
                color_discrete_map=TECHNOLOGY_COLORS,
                category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
                hole=0.42, title=f"Generation Capacity Mix — {year}",
            )
            fig.update_traces(textposition="inside", textinfo="percent+label",
                              hovertemplate="%{label}<br>%{value:,.1f} MW<br>%{percent}<extra></extra>")
            st.plotly_chart(fig, use_container_width=True)

    # National technology comparison.
    st.markdown("---")
    st.subheader("1. Technology-wise capacity movement")
    technology_comp = add_comparison_columns(
        gen31, gen36, ["Technology"], "Capacity_MW",
        "Installed generation capacity (MW)"
    )
    technology_comp["Technology"] = pd.Categorical(
        technology_comp["Technology"],
        categories=GENERATION_TECHNOLOGY_ORDER,
        ordered=True,
    )
    technology_comp = technology_comp.sort_values("Technology")

    fig = go.Figure()
    for yr, colour in [("FY31", "#4C78A8"), ("FY36", "#F2A541")]:
        fig.add_trace(
            go.Bar(
                name=yr,
                x=technology_comp["Technology"],
                y=technology_comp[yr],
                marker_color=colour,
                hovertemplate="%{x}<br>" + yr + ": %{y:,.1f} MW<extra></extra>",
            )
        )
    fig.update_layout(
        barmode="group",
        xaxis={"categoryorder": "array", "categoryarray": GENERATION_TECHNOLOGY_ORDER},
        title="Installed Generation Capacity by Technology",
        xaxis_title="Technology",
        yaxis_title="Installed Capacity (MW)",
        yaxis=dict(tickformat=","),
        legend_title_text="",
    )
    st.plotly_chart(fig, use_container_width=True)

    table = technology_comp[
        ["Technology", "FY31", "FY36", "Change", "Change_%"]
    ].copy()
    table["Change_%"] = table["Change_%"].map(
        lambda x: "—" if pd.isna(x) else f"{x:+.1f}%"
    )
    table = table.rename(
        columns={
            "FY31": "FY31 (MW)",
            "FY36": "FY36 (MW)",
            "Change": "Change (MW)",
            "Change_%": "Change (%)",
        }
    )
    st.dataframe(
        table.style.format(
            {
                "FY31 (MW)": "{:,.1f}",
                "FY36 (MW)": "{:,.1f}",
                "Change (MW)": "{:+,.1f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        "Download technology comparison (CSV)",
        technology_comp.to_csv(index=False).encode("utf-8"),
        file_name="FY31_FY36_Technology_Comparison.csv",
        mime="text/csv",
    )

    # Detailed Generation Sent Out comparison (energy output, GWh).
    st.markdown("---")
    st.subheader("2. Generation Sent Out (GWh)")
    sent31 = gen31["Generation_Sent_Out_GWh"].sum()
    sent36 = gen36["Generation_Sent_Out_GWh"].sum()
    sent_change = sent36 - sent31
    sent_pct = sent_change / abs(sent31) * 100 if sent31 else None
    gs1, gs2, gs3 = st.columns(3)
    gs1.metric("FY31 Generation Sent Out", f"{sent31:,.1f} GWh")
    gs2.metric("FY36 Generation Sent Out", f"{sent36:,.1f} GWh")
    gs3.metric("Net Change", f"{sent_change:+,.1f} GWh",
               delta=(f"{sent_pct:+.1f}%" if sent_pct is not None else "No FY31 base"))

    sent_tech = add_comparison_columns(
        gen31, gen36, ["Technology"], "Generation_Sent_Out_GWh",
        "Generation sent out (GWh)"
    )
    sent_tech["Technology"] = pd.Categorical(
        sent_tech["Technology"], categories=GENERATION_TECHNOLOGY_ORDER, ordered=True
    )
    sent_tech = sent_tech.sort_values("Technology")
    sent_long = sent_tech.melt(
        id_vars=["Technology"], value_vars=["FY31", "FY36"],
        var_name="Financial Year", value_name="Generation_Sent_Out_GWh"
    )
    fig = px.bar(
        sent_long, x="Technology", y="Generation_Sent_Out_GWh",
        color="Financial Year", barmode="group",
        color_discrete_map={"FY31": "#4C78A8", "FY36": "#F2A541"},
        category_orders={"Technology": GENERATION_TECHNOLOGY_ORDER},
        title="Generation Sent Out by Technology — FY31 vs FY36",
    )
    style_capacity_chart(fig, "Generation Sent Out (GWh)")
    st.plotly_chart(fig, use_container_width=True)

    sent_tech_table = sent_tech[["Technology", "FY31", "FY36", "Change", "Change_%"]].rename(
        columns={"FY31": "FY31 (GWh)", "FY36": "FY36 (GWh)",
                 "Change": "Change (GWh)", "Change_%": "Change (%)"}
    )
    st.dataframe(
        sent_tech_table.style.format({
            "FY31 (GWh)": "{:,.1f}", "FY36 (GWh)": "{:,.1f}",
            "Change (GWh)": "{:+,.1f}",
            "Change (%)": lambda x: "—" if pd.isna(x) else f"{x:+.1f}%",
        }), use_container_width=True, hide_index=True
    )
    st.download_button(
        "Download technology-wise Generation Sent Out (CSV)",
        sent_tech.to_csv(index=False).encode("utf-8"),
        file_name="FY31_FY36_Generation_Sent_Out_Technology.csv", mime="text/csv",
    )

    sent_state = add_comparison_columns(
        gen31, gen36, ["State"], "Generation_Sent_Out_GWh",
        "Generation sent out (GWh)"
    ).sort_values("Change", ascending=False)
    sent_state_long = sent_state.melt(
        id_vars=["State"], value_vars=["FY31", "FY36"],
        var_name="Financial Year", value_name="Generation_Sent_Out_GWh"
    )
    fig = px.bar(
        sent_state_long, x="Generation_Sent_Out_GWh", y="State",
        color="Financial Year", barmode="group", orientation="h",
        color_discrete_map={"FY31": "#4C78A8", "FY36": "#F2A541"},
        title="State-wise Generation Sent Out — FY31 vs FY36",
    )
    fig.update_layout(xaxis_title="Generation Sent Out (GWh)", yaxis_title="State / Region",
                      xaxis=dict(tickformat=","), height=max(500, 22 * len(sent_state)))
    st.plotly_chart(fig, use_container_width=True)

    sent_state_table = sent_state[["State", "FY31", "FY36", "Change", "Change_%"]].rename(
        columns={"FY31": "FY31 (GWh)", "FY36": "FY36 (GWh)",
                 "Change": "Change (GWh)", "Change_%": "Change (%)"}
    )
    st.dataframe(
        sent_state_table.style.format({
            "FY31 (GWh)": "{:,.1f}", "FY36 (GWh)": "{:,.1f}",
            "Change (GWh)": "{:+,.1f}",
            "Change (%)": lambda x: "—" if pd.isna(x) else f"{x:+.1f}%",
        }), use_container_width=True, hide_index=True
    )
    st.download_button(
        "Download state-wise Generation Sent Out (CSV)",
        sent_state.to_csv(index=False).encode("utf-8"),
        file_name="FY31_FY36_Generation_Sent_Out_Statewise.csv", mime="text/csv",
    )

    # Status comparison.
    st.markdown("---")
    st.subheader("3. Capacity composition by project status")
    status_comp = add_comparison_columns(
        gen31, gen36, ["Status"], "Capacity_MW",
        "Installed generation capacity (MW)"
    )
    status_long = status_comp.melt(
        id_vars=["Status"],
        value_vars=["FY31", "FY36"],
        var_name="Financial Year",
        value_name="Capacity_MW",
    )
    fig = px.bar(
        status_long,
        x="Status",
        y="Capacity_MW",
        color="Financial Year",
        barmode="group",
        color_discrete_map={"FY31": "#4C78A8", "FY36": "#F2A541"},
        title="Existing, Upcoming, Candidate and Retirement Capacity",
    )
    style_capacity_chart(fig, "Installed Capacity (MW)")
    st.plotly_chart(fig, use_container_width=True)

    status_table = status_comp[
        ["Status", "FY31", "FY36", "Change", "Change_%"]
    ].rename(
        columns={
            "FY31": "FY31 (MW)",
            "FY36": "FY36 (MW)",
            "Change": "Change (MW)",
            "Change_%": "Change (%)",
        }
    )
    st.dataframe(
        status_table.style.format(
            {
                "FY31 (MW)": "{:,.1f}",
                "FY36 (MW)": "{:,.1f}",
                "Change (MW)": "{:+,.1f}",
                "Change (%)": lambda x: "—" if pd.isna(x) else f"{x:+.1f}%",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        f"Upcoming + Candidates pipeline: FY31 {pipeline31:,.1f} MW; "
        f"FY36 {pipeline36:,.1f} MW. Retirement category: FY31 "
        f"{retirement31:,.1f} MW; FY36 {retirement36:,.1f} MW. "
        "These are source-category aggregates and should be interpreted "
        "according to the model's category definitions."
    )

    # State-wise comparison.
    st.markdown("---")
    st.subheader("4. State-wise change in installed generation capacity")
    state_comp = add_comparison_columns(
        gen31, gen36, ["State"], "Capacity_MW",
        "Installed generation capacity (MW)"
    ).sort_values("Change", ascending=False)

    fig = px.bar(
        state_comp.sort_values("Change", ascending=True),
        x="Change",
        y="State",
        orientation="h",
        color="Change",
        color_continuous_scale="RdYlGn",
        title="Net Capacity Change by State / Region (FY36 minus FY31)",
        hover_data=["FY31", "FY36", "Change_%"],
    )
    fig.update_layout(
        xaxis_title="Change in Installed Capacity (MW)",
        yaxis_title="State / Region",
        xaxis=dict(tickformat=","),
        coloraxis_colorbar_title="Change (MW)",
        height=max(500, 22 * len(state_comp)),
    )
    st.plotly_chart(fig, use_container_width=True)

    state_table = state_comp[
        ["State", "FY31", "FY36", "Change", "Change_%"]
    ].copy()
    state_table["Change_%"] = state_table["Change_%"].map(
        lambda x: "—" if pd.isna(x) else f"{x:+.1f}%"
    )
    state_table = state_table.rename(
        columns={
            "FY31": "FY31 (MW)",
            "FY36": "FY36 (MW)",
            "Change": "Change (MW)",
            "Change_%": "Change (%)",
        }
    )
    st.dataframe(
        state_table.style.format(
            {
                "FY31 (MW)": "{:,.1f}",
                "FY36 (MW)": "{:,.1f}",
                "Change (MW)": "{:+,.1f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        "Download state-wise comparison (CSV)",
        state_comp.to_csv(index=False).encode("utf-8"),
        file_name="FY31_FY36_Statewise_Comparison.csv",
        mime="text/csv",
    )

    # State selector for a focused portfolio comparison.
    st.markdown("---")
    st.subheader("5. Drill down: state technology mix")
    states_both = sorted(set(gen31["State"]) | set(gen36["State"]))
    focus_state = st.selectbox(
        "Choose a state / region for technology-level comparison",
        states_both,
        key="comparison_focus_state",
    )
    focus31 = gen31[gen31["State"] == focus_state]
    focus36 = gen36[gen36["State"] == focus_state]
    focus_comp = add_comparison_columns(
        focus31, focus36, ["Technology"], "Capacity_MW",
        "Installed generation capacity (MW)"
    )
    focus_comp["Technology"] = pd.Categorical(
        focus_comp["Technology"],
        categories=GENERATION_TECHNOLOGY_ORDER,
        ordered=True,
    )
    focus_comp = focus_comp.sort_values("Technology")

    fig = go.Figure()
    for yr, colour in [("FY31", "#4C78A8"), ("FY36", "#F2A541")]:
        fig.add_trace(
            go.Bar(
                name=yr,
                x=focus_comp["Technology"],
                y=focus_comp[yr],
                marker_color=colour,
            )
        )
    fig.update_layout(
        barmode="group",
        xaxis={"categoryorder": "array", "categoryarray": GENERATION_TECHNOLOGY_ORDER},
        title=f"{focus_state}: Technology Capacity Mix",
        xaxis_title="Technology",
        yaxis_title="Installed Capacity (MW)",
        yaxis=dict(tickformat=","),
    )
    st.plotly_chart(fig, use_container_width=True)

    # Storage is deliberately compared separately in MWh.
    st.markdown("---")
    st.subheader("6. Storage energy capacity comparison")
    if not stor31.empty or not stor36.empty:
        storage_comp = add_comparison_columns(
            stor31, stor36, ["Technology"], "Capacity_MWh",
            "Installed storage energy capacity (MWh)"
        )
        storage_comp["Technology"] = pd.Categorical(
            storage_comp["Technology"],
            categories=STORAGE_TECHNOLOGY_ORDER,
            ordered=True,
        )
        storage_comp = storage_comp.sort_values("Technology")
        storage_total31 = stor31["Capacity_MWh"].sum()
        storage_total36 = stor36["Capacity_MWh"].sum()
        sm1, sm2, sm3 = st.columns(3)
        sm1.metric("FY31 Storage Energy Capacity", f"{storage_total31:,.1f} MWh")
        sm2.metric("FY36 Storage Energy Capacity", f"{storage_total36:,.1f} MWh")
        sm3.metric(
            "Net Change",
            f"{storage_total36-storage_total31:+,.1f} MWh",
            delta=(
                f"{(storage_total36-storage_total31)/abs(storage_total31)*100:+.1f}%"
                if storage_total31 else "No FY31 base"
            ),
        )

        storage_long = storage_comp.melt(
            id_vars=["Technology"],
            value_vars=["FY31", "FY36"],
            var_name="Financial Year",
            value_name="Capacity_MWh",
        )
        fig = px.bar(
            storage_long,
            x="Technology",
            y="Capacity_MWh",
            color="Financial Year",
            barmode="group",
            color_discrete_map={"FY31": "#4C78A8", "FY36": "#F2A541"},
            category_orders={"Technology": STORAGE_TECHNOLOGY_ORDER},
            title="BESS and PSP Installed Energy Capacity",
        )
        style_capacity_chart(fig, "Installed Energy Capacity (MWh)")
        st.plotly_chart(fig, use_container_width=True)

        storage_table = storage_comp[
            ["Technology", "FY31", "FY36", "Change", "Change_%"]
        ].rename(
            columns={
                "FY31": "FY31 (MWh)",
                "FY36": "FY36 (MWh)",
                "Change": "Change (MWh)",
                "Change_%": "Change (%)",
            }
        )
        st.dataframe(
            storage_table.style.format(
                {
                    "FY31 (MWh)": "{:,.1f}",
                    "FY36 (MWh)": "{:,.1f}",
                    "Change (MWh)": "{:+,.1f}",
                    "Change (%)": lambda x: "—" if pd.isna(x) else f"{x:+.1f}%",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "Download storage comparison (CSV)",
            storage_comp.to_csv(index=False).encode("utf-8"),
            file_name="FY31_FY36_Storage_Comparison.csv",
            mime="text/csv",
        )
    else:
        st.info("No storage records are available in either workbook.")

    # Short methodology and interpretation note.
    with st.expander("Methodology, definitions and interpretation"):
        st.markdown(
            """
            - **Comparison basis:** FY31 and FY36 workbook records are aggregated
              from each worksheet using the `Installed Capacity (MW)` field for
              generators and `Installed Capacity (MWh)` for storage.
            - **Technology mapping:** Source category labels are mapped into a
              consistent technology taxonomy (Thermal, Nuclear, Hydro, Other RE (including DRE),
              Solar, Wind, BESS and PSP). The original source category is
              retained in the underlying records.
            - **Status mapping:** Existing, Upcoming, Candidates and Retirement
              are classified from the source category text.
            - **Change:** `FY36 − FY31`. A positive value indicates an increase
              in the FY36 modelled capacity position; a negative value indicates
              a decrease.
            - **Percentage change:** Calculated relative to FY31. Where the FY31
              base is zero, the percentage is shown as unavailable rather than
              presenting a misleading percentage.
            - **Units:** MW (power capacity) and MWh (energy capacity) are
              presented separately and are never summed together.
            - **Caution:** Differences between planning-year outputs should be
              interpreted in the context of scenario assumptions, model inputs,
              and category definitions. A change in installed capacity does not
              by itself establish a change in actual generation, utilisation,
              or project commissioning.
            """
        )

st.markdown("---")
st.caption(
    "PLEXOS model output | FY31 and FY36 installed capacity analysis | "
    "Use the comparison view to identify portfolio shifts, state-level "
    "concentration and changes in the storage mix."
)