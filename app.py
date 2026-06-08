import os
import glob
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# =========================
# PAGE SETUP
# =========================
st.set_page_config(
    page_title="SAMRIDH Dashboard",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================
# PREMIUM CSS
# =========================
st.markdown(
    """
    <style>
        .main {background-color: #f7f9fc;}
        div[data-testid="stMetric"] {
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            border: 1px solid #e5e7eb;
            padding: 18px;
            border-radius: 18px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        }
        div[data-testid="stMetric"] label {
            color: #475569 !important;
            font-weight: 700 !important;
        }
        .section-card {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 20px;
            padding: 20px 22px;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
            margin-bottom: 18px;
        }
        .big-title {
            font-size: 34px;
            font-weight: 900;
            color: #0f172a;
            margin-bottom: 0px;
        }
        .sub-title {
            font-size: 15px;
            color: #64748b;
            margin-top: 2px;
            margin-bottom: 18px;
        }
        .small-note {
            font-size: 13px;
            color: #64748b;
        }
        .stTabs [data-baseweb="tab-list"] {gap: 8px;}
        .stTabs [data-baseweb="tab"] {
            background-color: #ffffff;
            border-radius: 12px;
            border: 1px solid #e5e7eb;
            padding: 10px 18px;
            font-weight: 700;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================
# CONSTANTS
# =========================
COLOR_SEQ = [
    "#2563eb", "#0f766e", "#f97316", "#7c3aed", "#dc2626",
    "#0891b2", "#16a34a", "#ca8a04", "#9333ea", "#ea580c"
]

STATE_CENTROIDS = {
    "Andhra Pradesh": (15.9129, 79.7400),
    "Assam": (26.2006, 92.9376),
    "Bihar": (25.0961, 85.3131),
    "Chandigarh": (30.7333, 76.7794),
    "Delhi": (28.7041, 77.1025),
    "Goa": (15.2993, 74.1240),
    "Gujarat": (22.2587, 71.1924),
    "Haryana": (29.0588, 76.0856),
    "Himachal Pradesh": (31.1048, 77.1734),
    "Jammu & Kashmir": (33.7782, 76.5762),
    "Jharkhand": (23.6102, 85.2799),
    "Karnataka": (15.3173, 75.7139),
    "Kerala": (10.8505, 76.2711),
    "Madhya Pradesh": (22.9734, 78.6569),
    "Maharashtra": (19.7515, 75.7139),
    "Manipur": (24.6637, 93.9063),
    "Odisha": (20.9517, 85.0985),
    "Punjab": (31.1471, 75.3412),
    "Rajasthan": (27.0238, 74.2179),
    "Tamil Nadu": (11.1271, 78.6569),
    "Telangana": (18.1124, 79.0193),
    "Uttar Pradesh": (26.8467, 80.9462),
    "Uttarakhand": (30.0668, 79.0193),
    "West Bengal": (22.9868, 87.8550),
}

# =========================
# DATA LOADING
# =========================
def find_excel_file() -> str | None:
    """Folder me dashboard-ready Excel automatically find karega."""
    preferred = [
        "SAMRIDH_2nd_Cohort_Cleaned_Dashboard_Ready_Checked.xlsx",
        "SAMRIDH_2nd_Cohort_Cleaned_Dashboard_Ready_Checked(1).xlsx",
        "SAMRIDH_2nd_Cohort_Cleaned_Dashboard_Ready.xlsx",
    ]
    for file in preferred:
        if os.path.exists(file):
            return file

    excel_files = glob.glob("*.xlsx")
    excel_files = [f for f in excel_files if not Path(f).name.startswith("~$")]
    if excel_files:
        return excel_files[0]
    return None


def clean_text(x):
    if pd.isna(x):
        return "Unknown"
    x = str(x).strip()
    x = re.sub(r"\s+", " ", x)
    return x if x else "Unknown"


@st.cache_data(show_spinner=False)
def load_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name="Clean_Data", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    required_cols = [
        "Cohort Label", "Accelerator", "Startup Name", "Sector Clean",
        "Technology Tags", "Primary Technology", "City Clean", "State Clean", "Tier Label"
    ]
    for col in required_cols:
        if col not in df.columns:
            st.error(f"Missing column: {col}. Please use checked cleaned Excel file.")
            st.stop()

    text_cols = [
        "Cohort Label", "Accelerator", "Startup Name", "Sector Clean",
        "Technology Tags", "Primary Technology", "City Clean", "State Clean", "Tier Label"
    ]
    for col in text_cols:
        df[col] = df[col].apply(clean_text)

    if "Duplicate Startup Flag" in df.columns:
        df["Duplicate Startup Flag"] = df["Duplicate Startup Flag"].fillna("No")

    return df


def filter_df(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.markdown("### 🔎 Filters")

    cohort = st.sidebar.multiselect(
        "Cohort",
        sorted(df["Cohort Label"].dropna().unique()),
        default=sorted(df["Cohort Label"].dropna().unique()),
    )
    state = st.sidebar.multiselect(
        "State",
        sorted(df["State Clean"].dropna().unique()),
        default=[],
    )
    sector = st.sidebar.multiselect(
        "Sector",
        sorted(df["Sector Clean"].dropna().unique()),
        default=[],
    )
    accelerator = st.sidebar.multiselect(
        "Accelerator",
        sorted(df["Accelerator"].dropna().unique()),
        default=[],
    )
    tier = st.sidebar.multiselect(
        "Tier",
        sorted(df["Tier Label"].dropna().unique()),
        default=[],
    )

    out = df[df["Cohort Label"].isin(cohort)].copy()
    if state:
        out = out[out["State Clean"].isin(state)]
    if sector:
        out = out[out["Sector Clean"].isin(sector)]
    if accelerator:
        out = out[out["Accelerator"].isin(accelerator)]
    if tier:
        out = out[out["Tier Label"].isin(tier)]
    return out


# =========================
# CHART HELPERS
# =========================
def apply_layout(fig, height=420, showlegend=True):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial", size=13, color="#0f172a"),
        margin=dict(l=20, r=20, t=55, b=30),
        showlegend=showlegend,
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
    )
    fig.update_xaxes(showgrid=False, zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="#e5e7eb", zeroline=False)
    return fig


def bar_count(df, col, title, top_n=10, horizontal=True):
    data = df[col].value_counts().head(top_n).reset_index()
    data.columns = [col, "Startups"]
    if horizontal:
        data = data.sort_values("Startups")
        fig = px.bar(data, x="Startups", y=col, orientation="h", text="Startups", color="Startups", color_continuous_scale="Blues")
    else:
        fig = px.bar(data, x=col, y="Startups", text="Startups", color=col, color_discrete_sequence=COLOR_SEQ)
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_layout(title=title, coloraxis_showscale=False)
    return apply_layout(fig, showlegend=False)


def explode_technology(df):
    temp = df[["Startup Name", "Sector Clean", "Technology Tags", "Cohort Label", "State Clean", "Accelerator"]].copy()
    temp["Technology"] = temp["Technology Tags"].str.split("|")
    temp = temp.explode("Technology")
    temp["Technology"] = temp["Technology"].apply(clean_text)
    return temp[temp["Technology"] != "Unknown"]

# =========================
# MAIN APP
# =========================
st.markdown('<div class="big-title">SAMRIDH 2nd Cohort Dashboard</div>', unsafe_allow_html=True)


file_path = find_excel_file()
if not file_path:
    st.error("Excel file nahi mila. Is app.py ke same folder me cleaned Excel file rakho.")
    st.stop()

raw_df = load_data(file_path)
df = filter_df(raw_df)
tech_df = explode_technology(df)

st.sidebar.markdown("---")

if df.empty:
    st.warning("Selected filters ke according koi data nahi mila.")
    st.stop()

# =========================
# KPI ROW
# =========================
total_startups = df["Startup Name"].nunique()
total_accelerators = df["Accelerator"].nunique()
total_states = df["State Clean"].nunique()
total_cities = df["City Clean"].nunique()
top_sector = df["Sector Clean"].value_counts().idxmax()
top_tech = tech_df["Technology"].value_counts().idxmax() if not tech_df.empty else "NA"

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Startups", f"{total_startups:,}")
k2.metric("Accelerators", f"{total_accelerators:,}")
k3.metric("States Covered", f"{total_states:,}")
k4.metric("Cities Covered", f"{total_cities:,}")

k5, k6, k7, k8 = st.columns(4)
k5.metric("Top Sector", top_sector)
k6.metric("Top Technology", top_tech)
k7.metric("Cohort 1", f"{(df['Cohort Label'] == 'Cohort 1').sum():,}")
k8.metric("Cohort 2", f"{(df['Cohort Label'] == 'Cohort 2').sum():,}")

st.markdown("---")

# =========================
# TABS
# =========================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📌 Executive View",
    "🗺️ Geography",
    "🏭 Sector & Tech",
    "🏢 Accelerator",
    "📋 Data Explorer",
])

with tab1:
    c1, c2 = st.columns([1, 1])
    with c1:
        cohort_data = df["Cohort Label"].value_counts().reset_index()
        cohort_data.columns = ["Cohort", "Startups"]
        fig = px.pie(
            cohort_data,
            names="Cohort",
            values="Startups",
            hole=0.55,
            title="Cohort-wise Startup Distribution",
            color_discrete_sequence=COLOR_SEQ,
        )
        fig.update_traces(textinfo="label+percent+value")
        st.plotly_chart(apply_layout(fig, height=410), use_container_width=True)

    with c2:
        tier_data = df["Tier Label"].value_counts().reset_index()
        tier_data.columns = ["Tier", "Startups"]
        fig = px.bar(tier_data, x="Tier", y="Startups", text="Startups", title="Tier-wise Startup Spread", color="Tier", color_discrete_sequence=COLOR_SEQ)
        fig.update_traces(textposition="outside")
        st.plotly_chart(apply_layout(fig, height=410), use_container_width=True)

    c3, c4 = st.columns([1, 1])
    with c3:
        st.plotly_chart(bar_count(df, "State Clean", "Top States by Startups", top_n=10), use_container_width=True)
    with c4:
        st.plotly_chart(bar_count(df, "Sector Clean", "Top Sectors by Startups", top_n=10), use_container_width=True)

with tab2:
    state_counts = df["State Clean"].value_counts().reset_index()
    state_counts.columns = ["State", "Startups"]
    state_counts["lat"] = state_counts["State"].map(lambda x: STATE_CENTROIDS.get(x, (None, None))[0])
    state_counts["lon"] = state_counts["State"].map(lambda x: STATE_CENTROIDS.get(x, (None, None))[1])
    state_counts = state_counts.dropna(subset=["lat", "lon"])

    fig = px.scatter_geo(
        state_counts,
        lat="lat",
        lon="lon",
        size="Startups",
        color="Startups",
        hover_name="State",
        hover_data={"Startups": True, "lat": False, "lon": False},
        title="India Startup Coverage by State",
        projection="natural earth",
        color_continuous_scale="Viridis",
        size_max=42,
    )
    fig.update_geos(
        visible=False,
        showcountries=True,
        countrycolor="#cbd5e1",
        showland=True,
        landcolor="#f8fafc",
        fitbounds="locations",
        lataxis_range=[6, 38],
        lonaxis_range=[68, 98],
    )
    fig.update_layout(coloraxis_showscale=True)
    st.plotly_chart(apply_layout(fig, height=520, showlegend=False), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(bar_count(df, "State Clean", "Top 15 States", top_n=15), use_container_width=True)
    with c2:
        st.plotly_chart(bar_count(df, "City Clean", "Top 15 Cities", top_n=15), use_container_width=True)

with tab3:
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(bar_count(df, "Sector Clean", "Sector-wise Startup Count", top_n=15), use_container_width=True)
    with c2:
        tech_counts = tech_df["Technology"].value_counts().head(15).reset_index()
        tech_counts.columns = ["Technology", "Startups"]
        tech_counts = tech_counts.sort_values("Startups")
        fig = px.bar(tech_counts, x="Startups", y="Technology", orientation="h", text="Startups", title="Top Technology Tags", color="Startups", color_continuous_scale="Teal")
        fig.update_traces(textposition="outside", cliponaxis=False)
        st.plotly_chart(apply_layout(fig, showlegend=False), use_container_width=True)

    heat = tech_df.groupby(["Sector Clean", "Technology"]).size().reset_index(name="Startups")
    top_sectors = df["Sector Clean"].value_counts().head(10).index
    top_techs = tech_df["Technology"].value_counts().head(10).index
    heat = heat[heat["Sector Clean"].isin(top_sectors) & heat["Technology"].isin(top_techs)]
    pivot = heat.pivot(index="Sector Clean", columns="Technology", values="Startups").fillna(0)
    fig = px.imshow(
        pivot,
        text_auto=True,
        aspect="auto",
        title="Sector vs Technology Heatmap",
        color_continuous_scale="Blues",
    )
    st.plotly_chart(apply_layout(fig, height=560, showlegend=False), use_container_width=True)

with tab4:
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(bar_count(df, "Accelerator", "Top Accelerators by Startup Count", top_n=15), use_container_width=True)
    with c2:
        acc_cohort = df.groupby(["Accelerator", "Cohort Label"]).size().reset_index(name="Startups")
        top_acc = df["Accelerator"].value_counts().head(12).index
        acc_cohort = acc_cohort[acc_cohort["Accelerator"].isin(top_acc)]
        fig = px.bar(
            acc_cohort,
            x="Startups",
            y="Accelerator",
            color="Cohort Label",
            orientation="h",
            title="Top Accelerators: Cohort Split",
            text="Startups",
            color_discrete_sequence=COLOR_SEQ,
        )
        fig.update_traces(textposition="inside")
        st.plotly_chart(apply_layout(fig), use_container_width=True)

    acc_sector = df.groupby(["Accelerator", "Sector Clean"]).size().reset_index(name="Startups")
    top_acc2 = df["Accelerator"].value_counts().head(10).index
    acc_sector = acc_sector[acc_sector["Accelerator"].isin(top_acc2)]
    fig = px.treemap(
        acc_sector,
        path=["Accelerator", "Sector Clean"],
        values="Startups",
        title="Accelerator-wise Sector Portfolio",
        color="Startups",
        color_continuous_scale="Blues",
    )
    st.plotly_chart(apply_layout(fig, height=560), use_container_width=True)

with tab5:
    st.markdown("### Search / Export Data")
    search = st.text_input("Startup / Accelerator / City / State search")
    show_df = df.copy()
    if search.strip():
        s = search.strip().lower()
        show_df = show_df[
            show_df["Startup Name"].str.lower().str.contains(s, na=False)
            | show_df["Accelerator"].str.lower().str.contains(s, na=False)
            | show_df["City Clean"].str.lower().str.contains(s, na=False)
            | show_df["State Clean"].str.lower().str.contains(s, na=False)
            | show_df["Sector Clean"].str.lower().str.contains(s, na=False)
        ]

    display_cols = [
        "Cohort Label", "Accelerator", "Startup Name", "Sector Clean",
        "Primary Technology", "Technology Tags", "City Clean", "State Clean", "Tier Label"
    ]
    st.dataframe(show_df[display_cols], use_container_width=True, hide_index=True, height=520)

    csv = show_df[display_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Download Filtered CSV",
        data=csv,
        file_name="samridh_filtered_dashboard_data.csv",
        mime="text/csv",
    )

