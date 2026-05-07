import re
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# FILE PATHS
# ============================================================
RESULTS_FILE = "praktikum_unibern_group4.xls"
GEOMETRY_FILE = "uniber_4.xlsx"

st.set_page_config(
    page_title="UniBern 3-Point Ball Contact Loading Dashboard",
    layout="wide"
)

st.title("UniBern Group 3-Point Ball Contact Loading Test Dashboard")

# ============================================================
# LOAD DATA
# ============================================================
@st.cache_data
def load_results_workbook(path):
    return pd.ExcelFile(path).sheet_names


@st.cache_data
def load_results_table(path):
    df = pd.read_excel(path, sheet_name="Results")

    df = df[df["Specimen no."].notna()].copy()

    df["Specimen no."] = df["Specimen no."].astype(int)
    df["Fmax"] = pd.to_numeric(df["Fmax"], errors="coerce")
    df["tTest"] = pd.to_numeric(df["tTest"], errors="coerce")
    df["Specimen ID"] = df["Specimen ID"].astype(str)

    def comment_from_id(x):
        x_low = str(x).lower()
        comments = []
        if "calibration" in x_low:
            comments.append("Calibration specimen")
        if "white" in x_low:
            comments.append("White specimen / special visual note")
        return "; ".join(comments) if comments else "Normal specimen"

    df["Comment"] = df["Specimen ID"].apply(comment_from_id)

    return df


@st.cache_data
def load_geometry(path):
    df = pd.read_excel(path)

    df["Sample"] = pd.to_numeric(df["Sample"], errors="coerce")
    df = df[df["Sample"].notna()].copy()
    df["Sample"] = df["Sample"].astype(int)

    for col in ["d1", "d2", "d3", "h1", "h2", "h3"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["diameter_avg"] = df[["d1", "d2", "d3"]].mean(axis=1)
    df["height_avg"] = df[["h1", "h2", "h3"]].mean(axis=1)

    # Cross-sectional area from average diameter
    df["area_mm2"] = 3.141592653589793 * (df["diameter_avg"] / 2) ** 2

    return df


@st.cache_data
def load_specimen_curve(path, specimen_no):
    sheet = f"Specimen {int(specimen_no)}"

    df = pd.read_excel(path, sheet_name=sheet, header=None)

    df = df.iloc[2:, :2].copy()
    df.columns = ["Time_s", "Standard_force_N"]

    df["Time_s"] = pd.to_numeric(df["Time_s"], errors="coerce")
    df["Standard_force_N"] = pd.to_numeric(df["Standard_force_N"], errors="coerce")
    df = df.dropna()

    return df


sheet_names = load_results_workbook(RESULTS_FILE)
results_df = load_results_table(RESULTS_FILE)
geometry_df = load_geometry(GEOMETRY_FILE)

available_specimens = [int(x) for x in sorted(results_df["Specimen no."].unique())]

missing_specimens = [
    int(x) for x in sorted(
        set(range(min(available_specimens), max(available_specimens) + 1))
        - set(available_specimens)
    )
]

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("Controls")

mode = st.sidebar.radio(
    "Visualization mode",
    ["Single specimen", "Compare specimens", "Results overview"]
)

st.sidebar.info(
    f"Available specimens: {available_specimens}\n\n"
    f"Missing specimens in Results sheet: {missing_specimens}"
)

# ============================================================
# RESULTS OVERVIEW
# ============================================================
if mode == "Results overview":
    st.header("Results Overview")

    display_cols = ["Specimen no.", "Specimen ID", "Fmax", "tTest", "Comment"]
    st.subheader("Results table")
    st.dataframe(results_df[display_cols], use_container_width=True)

    col1, col2, col3 = st.columns(3)

    col1.metric("Number of valid specimens", len(results_df))
    col2.metric("Maximum Fmax [N]", f"{results_df['Fmax'].max():.2f}")
    col3.metric("Mean Fmax [N]", f"{results_df['Fmax'].mean():.2f}")

    st.subheader("Fmax by Specimen")

    fig_fmax = px.bar(
        results_df,
        x="Specimen no.",
        y="Fmax",
        text="Specimen no.",
        hover_data=["Specimen ID", "tTest", "Comment"],
        title="Maximum Force Fmax per Specimen",
        labels={
            "Specimen no.": "Specimen no.",
            "Fmax": "Fmax [N]"
        }
    )
    fig_fmax.update_traces(
        texttemplate="S%{text}",
        textposition="outside",
        cliponaxis=False
    )
    fig_fmax.update_layout(
        xaxis=dict(
            tickmode="array",
            tickvals=available_specimens,
            ticktext=[str(x) for x in available_specimens]
        ),
        uniformtext_minsize=9,
        uniformtext_mode="show"
    )
    st.plotly_chart(fig_fmax, use_container_width=True)

    # ========================================================
    # AREA / DIAMETER / HEIGHT VS FORCE VISUALIZATION
    # ========================================================
    st.subheader("Specimen Geometry vs Maximum Force")

    geometry_force_df = results_df.merge(
        geometry_df[["Sample", "diameter_avg", "height_avg", "area_mm2"]],
        left_on="Specimen no.",
        right_on="Sample",
        how="left"
    ).drop(columns=["Sample"])

    # Apparent compressive stress / normalized load
    geometry_force_df["apparent_compressive_stress_N_per_mm2"] = (
        geometry_force_df["Fmax"] / geometry_force_df["area_mm2"]
    )

    geometry_force_df["apparent_compressive_stress_MPa"] = (
        geometry_force_df["apparent_compressive_stress_N_per_mm2"]
    )

    st.dataframe(
        geometry_force_df[
            [
                "Specimen no.",
                "Specimen ID",
                "diameter_avg",
                "height_avg",
                "area_mm2",
                "Fmax",
                "apparent_compressive_stress_MPa",
                "tTest",
                "Comment"
            ]
        ],
        use_container_width=True
    )

    fig_area_force = px.scatter(
        geometry_force_df,
        x="area_mm2",
        y="Fmax",
        text="Specimen no.",
        size="height_avg",
        hover_data=[
            "Specimen ID",
            "diameter_avg",
            "height_avg",
            "area_mm2",
            "apparent_compressive_stress_MPa",
            "tTest",
            "Comment"
        ],
        title="Relationship Between Cross-Sectional Area and Maximum Force",
        labels={
            "area_mm2": "Average Cross-Sectional Area [mm²]",
            "Fmax": "Fmax [N]",
            "height_avg": "Average Height [mm]",
            "diameter_avg": "Average Diameter [mm]",
            "apparent_compressive_stress_MPa": "Apparent Compressive Stress [MPa]"
        }
    )

    fig_area_force.update_traces(
        texttemplate="S%{text}",
        textposition="top center"
    )

    st.plotly_chart(fig_area_force, use_container_width=True)

    st.caption(
        "Each point represents one specimen. The x-axis shows calculated average cross-sectional area "
        "from the average measured diameter, the y-axis shows maximum force, and marker size represents "
        "average specimen height."
    )

    # ========================================================
    # NEW: APPARENT COMPRESSIVE STRESS
    # ========================================================
    st.subheader("Normalized Load / Apparent Compressive Stress")

    st.write(
        "The apparent compressive stress was calculated as the maximum force divided by "
        "the calculated cross-sectional area of each circular specimen."
    )

    fig_stress = px.bar(
        geometry_force_df,
        x="Specimen no.",
        y="apparent_compressive_stress_MPa",
        text="Specimen no.",
        hover_data=[
            "Specimen ID",
            "Fmax",
            "area_mm2",
            "diameter_avg",
            "height_avg",
            "Comment"
        ],
        title="Apparent Compressive Stress per Specimen",
        labels={
            "Specimen no.": "Specimen no.",
            "apparent_compressive_stress_MPa": "Apparent Compressive Stress [MPa]"
        }
    )

    fig_stress.update_traces(
        texttemplate="S%{text}",
        textposition="outside",
        cliponaxis=False
    )

    fig_stress.update_layout(
        xaxis=dict(
            tickmode="array",
            tickvals=available_specimens,
            ticktext=[str(x) for x in available_specimens]
        ),
        uniformtext_minsize=9,
        uniformtext_mode="show"
    )

    st.plotly_chart(fig_stress, use_container_width=True)

    # ========================================================
    # NEW: MEAN, STANDARD DEVIATION, AND CV
    # ========================================================
    st.subheader("Statistical Summary")

    def coefficient_of_variation(series):
        mean_value = series.mean()
        std_value = series.std()
        if pd.isna(mean_value) or mean_value == 0:
            return None
        return (std_value / mean_value) * 100

    statistical_summary = pd.DataFrame({
        "Parameter": [
            "Maximum force, Fmax",
            "Test time, tTest",
            "Average diameter",
            "Average height",
            "Cross-sectional area",
            "Apparent compressive stress"
        ],
        "Mean": [
            geometry_force_df["Fmax"].mean(),
            geometry_force_df["tTest"].mean(),
            geometry_force_df["diameter_avg"].mean(),
            geometry_force_df["height_avg"].mean(),
            geometry_force_df["area_mm2"].mean(),
            geometry_force_df["apparent_compressive_stress_MPa"].mean()
        ],
        "Standard deviation": [
            geometry_force_df["Fmax"].std(),
            geometry_force_df["tTest"].std(),
            geometry_force_df["diameter_avg"].std(),
            geometry_force_df["height_avg"].std(),
            geometry_force_df["area_mm2"].std(),
            geometry_force_df["apparent_compressive_stress_MPa"].std()
        ],
        "Coefficient of variation [%]": [
            coefficient_of_variation(geometry_force_df["Fmax"]),
            coefficient_of_variation(geometry_force_df["tTest"]),
            coefficient_of_variation(geometry_force_df["diameter_avg"]),
            coefficient_of_variation(geometry_force_df["height_avg"]),
            coefficient_of_variation(geometry_force_df["area_mm2"]),
            coefficient_of_variation(geometry_force_df["apparent_compressive_stress_MPa"])
        ],
        "Unit": [
            "N",
            "s",
            "mm",
            "mm",
            "mm²",
            "MPa"
        ]
    })

    st.dataframe(
        statistical_summary.style.format({
            "Mean": "{:.3f}",
            "Standard deviation": "{:.3f}",
            "Coefficient of variation [%]": "{:.2f}"
        }),
        use_container_width=True
    )

    st.caption(
        "The coefficient of variation describes the relative spread of the data. "
        "A lower CV indicates more consistent specimen behavior, while a higher CV indicates greater variability."
    )

    st.subheader("Diameter and Height vs Fmax")

    fig_diameter_height = px.scatter(
        geometry_force_df,
        x="diameter_avg",
        y="height_avg",
        size="Fmax",
        text="Specimen no.",
        hover_data=[
            "Specimen ID",
            "Fmax",
            "area_mm2",
            "apparent_compressive_stress_MPa",
            "diameter_avg",
            "height_avg",
            "tTest",
            "Comment"
        ],
        title="Average Diameter vs Average Height with Fmax as Marker Size",
        labels={
            "diameter_avg": "Average Diameter [mm]",
            "height_avg": "Average Height [mm]",
            "Fmax": "Fmax [N]",
            "apparent_compressive_stress_MPa": "Apparent Compressive Stress [MPa]"
        }
    )

    fig_diameter_height.update_traces(
        texttemplate="S%{text}",
        textposition="top center"
    )

    st.plotly_chart(fig_diameter_height, use_container_width=True)

    st.subheader("tTest by Specimen")

    fig_time = px.bar(
        results_df,
        x="Specimen no.",
        y="tTest",
        text="Specimen no.",
        hover_data=["Specimen ID", "Fmax", "Comment"],
        title="Test Time per Specimen",
        labels={
            "Specimen no.": "Specimen no.",
            "tTest": "tTest [s]"
        }
    )
    fig_time.update_traces(
        texttemplate="S%{text}",
        textposition="outside",
        cliponaxis=False
    )
    fig_time.update_layout(
        xaxis=dict(
            tickmode="array",
            tickvals=available_specimens,
            ticktext=[str(x) for x in available_specimens]
        ),
        uniformtext_minsize=9,
        uniformtext_mode="show"
    )
    st.plotly_chart(fig_time, use_container_width=True)

    st.subheader("Fmax vs tTest")

    fig_scatter = px.scatter(
        results_df,
        x="tTest",
        y="Fmax",
        text="Specimen no.",
        hover_data=["Specimen ID", "Comment"],
        title="Relationship Between Test Time and Maximum Force",
        labels={
            "tTest": "tTest [s]",
            "Fmax": "Fmax [N]"
        }
    )
    fig_scatter.update_traces(
        texttemplate="S%{text}",
        textposition="top center"
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

    st.info(
        "This overview keeps the original specimen numbering from the Results sheet. "
        "Specimens 8 and 14 are missing and are not renumbered."
    )

# ============================================================
# SINGLE SPECIMEN VIEW
# ============================================================
elif mode == "Single specimen":
    specimen_no = st.sidebar.selectbox(
        "Select specimen",
        available_specimens
    )

    specimen_no = int(specimen_no)

    st.header(f"Specimen {specimen_no}")

    curve_df = load_specimen_curve(RESULTS_FILE, specimen_no)

    result_row = results_df[results_df["Specimen no."] == specimen_no].iloc[0]

    geometry_row = geometry_df[geometry_df["Sample"] == specimen_no]

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Specimen no.", specimen_no)
    col2.metric("Fmax [N]", f"{result_row['Fmax']:.2f}")
    col3.metric("tTest [s]", f"{result_row['tTest']:.2f}")
    col4.metric("Specimen ID", str(result_row["Specimen ID"]))

    st.write(f"**Comment:** {result_row['Comment']}")

    if not geometry_row.empty:
        g = geometry_row.iloc[0]

        st.subheader("Geometry Summary Before Loading")

        geom_summary = pd.DataFrame({
            "Parameter": [
                "d1", "d2", "d3",
                "Average diameter",
                "h1", "h2", "h3",
                "Average height",
                "Calculated area"
            ],
            "Value": [
                g["d1"], g["d2"], g["d3"],
                g["diameter_avg"],
                g["h1"], g["h2"], g["h3"],
                g["height_avg"],
                g["area_mm2"]
            ],
            "Unit": [
                "mm", "mm", "mm",
                "mm",
                "mm", "mm", "mm",
                "mm",
                "mm²"
            ]
        })

        st.dataframe(geom_summary, use_container_width=True)

    else:
        st.warning("No matching diameter/height data found for this specimen in uniber_4.xlsx.")

    st.subheader("Force-Time Curve")

    fig = px.line(
        curve_df,
        x="Time_s",
        y="Standard_force_N",
        title=f"Specimen {specimen_no}: Time vs Standard Force",
        labels={
            "Time_s": "Time [s]",
            "Standard_force_N": "Standard Force [N]"
        }
    )

    fig.add_hline(
        y=result_row["Fmax"],
        line_dash="dash",
        annotation_text=f"Fmax = {result_row['Fmax']:.2f} N",
        annotation_position="top left"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Raw Data Table")
    st.dataframe(curve_df, use_container_width=True)

    st.subheader("Automatic Short Interpretation")

    peak_idx = curve_df["Standard_force_N"].idxmax()
    peak_time = curve_df.loc[peak_idx, "Time_s"]
    peak_force = curve_df.loc[peak_idx, "Standard_force_N"]

    st.write(
        f"Specimen **{specimen_no}** reached a maximum measured force of "
        f"**{peak_force:.2f} N** at approximately **{peak_time:.2f} s**. "
        f"The Results sheet reports Fmax = **{result_row['Fmax']:.2f} N** "
        f"and total test time tTest = **{result_row['tTest']:.2f} s**. "
        f"Comment classification: **{result_row['Comment']}**."
    )

# ============================================================
# COMPARE SPECIMENS
# ============================================================
elif mode == "Compare specimens":
    selected_specimens = st.sidebar.multiselect(
        "Select specimens to compare",
        available_specimens,
        default=available_specimens[:3]
    )

    selected_specimens = [int(x) for x in selected_specimens]

    st.header("Compare Specimens")

    if len(selected_specimens) == 0:
        st.warning("Select at least one specimen.")
        st.stop()

    combined = []

    for specimen_no in selected_specimens:
        df = load_specimen_curve(RESULTS_FILE, specimen_no)
        df["Specimen no."] = specimen_no
        combined.append(df)

    compare_df = pd.concat(combined, ignore_index=True)

    st.subheader("Combined Force-Time Curves")

    fig = px.line(
        compare_df,
        x="Time_s",
        y="Standard_force_N",
        color="Specimen no.",
        title="Comparison of Force-Time Curves",
        labels={
            "Time_s": "Time [s]",
            "Standard_force_N": "Standard Force [N]",
            "Specimen no.": "Specimen"
        }
    )

    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Selected Specimens Results Summary")

    summary_df = results_df[
        results_df["Specimen no."].isin(selected_specimens)
    ][["Specimen no.", "Specimen ID", "Fmax", "tTest", "Comment"]].copy()

    summary_df = summary_df.merge(
        geometry_df[["Sample", "diameter_avg", "height_avg", "area_mm2"]],
        left_on="Specimen no.",
        right_on="Sample",
        how="left"
    ).drop(columns=["Sample"])

    summary_df["apparent_compressive_stress_MPa"] = (
        summary_df["Fmax"] / summary_df["area_mm2"]
    )

    st.dataframe(summary_df, use_container_width=True)

    st.subheader("Fmax Comparison")

    fig_bar = px.bar(
        summary_df,
        x="Specimen no.",
        y="Fmax",
        text="Specimen no.",
        hover_data=[
            "Specimen ID",
            "tTest",
            "Comment",
            "diameter_avg",
            "height_avg",
            "area_mm2",
            "apparent_compressive_stress_MPa"
        ],
        title="Fmax Comparison of Selected Specimens",
        labels={
            "Specimen no.": "Specimen no.",
            "Fmax": "Fmax [N]"
        }
    )
    fig_bar.update_traces(
        texttemplate="S%{text}",
        textposition="outside",
        cliponaxis=False
    )
    fig_bar.update_layout(
        xaxis=dict(
            tickmode="array",
            tickvals=selected_specimens,
            ticktext=[str(x) for x in selected_specimens]
        ),
        uniformtext_minsize=9,
        uniformtext_mode="show"
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    st.subheader("Short Comparison Description")

    best = summary_df.loc[summary_df["Fmax"].idxmax()]
    weakest = summary_df.loc[summary_df["Fmax"].idxmin()]

    st.write(
        f"Among the selected specimens, **Specimen {int(best['Specimen no.'])}** "
        f"had the highest Fmax of **{best['Fmax']:.2f} N**, while "
        f"**Specimen {int(weakest['Specimen no.'])}** had the lowest Fmax of "
        f"**{weakest['Fmax']:.2f} N**. "
        f"This comparison helps identify which specimens resisted the applied contact loading force more strongly "
        f"before failure or test termination."
    )