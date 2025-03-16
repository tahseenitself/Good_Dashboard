import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Set page config for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ---- Sidebar Filters ----
st.sidebar.title("Filters")

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)

df_filtered = df_original.copy()
if selected_region != "All":
    df_filtered = df_filtered[df_filtered["Region"] == selected_region]

# State Filter
all_states = sorted(df_filtered["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)

if selected_state != "All":
    df_filtered = df_filtered[df_filtered["State"] == selected_state]

# Category Filter
all_categories = sorted(df_filtered["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)

if selected_category != "All":
    df_filtered = df_filtered[df_filtered["Category"] == selected_category]

# Sub-Category Filter
all_subcats = sorted(df_filtered["Sub-Category"].dropna().unique())
selected_subcat = st.sidebar.selectbox("Select Sub-Category", options=["All"] + all_subcats)

if selected_subcat != "All":
    df_filtered = df_filtered[df_filtered["Sub-Category"] == selected_subcat]

# ---- Sidebar Date Range ----
min_date = df_filtered["Order Date"].min()
max_date = df_filtered["Order Date"].max()

from_date = st.sidebar.date_input("From Date", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("To Date", value=max_date, min_value=min_date, max_value=max_date)

df_filtered = df_filtered[
    (df_filtered["Order Date"] >= pd.to_datetime(from_date)) &
    (df_filtered["Order Date"] <= pd.to_datetime(to_date))
]

# ---- KPI Selection (MODIFIED) ----
st.subheader("Visualize KPI Across Time & Top Products")

kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
kpi_selection = st.sidebar.multiselect(
    "Select KPIs to Display",
    kpi_options,
    default=["Sales", "Profit"]
)

# ---- KPI Calculation ----
if df_filtered.empty:
    total_sales, total_quantity, total_profit, margin_rate = 0, 0, 0, 0
else:
    total_sales = df_filtered["Sales"].sum()
    total_quantity = df_filtered["Quantity"].sum()
    total_profit = df_filtered["Profit"].sum()
    margin_rate = (total_profit / total_sales) if total_sales != 0 else 0

# ---- KPI Display (MODIFIED) ----
st.title("Enhanced Superstore Dashboard")

kpi_values = {
    "Sales": f"${total_sales:,.2f}",
    "Quantity": f"{total_quantity:,.0f}",
    "Profit": f"${total_profit:,.2f}",
    "Margin Rate": f"{(margin_rate * 100):,.2f}%"
}

kpi_columns = st.columns(len(kpi_selection))

for i, kpi in enumerate(kpi_selection):
    with kpi_columns[i]:  
        st.markdown(
            f"""
            <div class='kpi-box'>
                <div class='kpi-title'>{kpi}</div>
                <div class='kpi-value'>{kpi_values[kpi]}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

# ---- Prepare Data for Charts ----
if not df_filtered.empty:
    daily_grouped = df_filtered.groupby("Order Date").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)

    product_grouped = df_filtered.groupby("Product Name").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)

    # ---- Sales Trend Chart (MODIFIED) ----
    st.subheader("Sales Trend Over Time")

    fig_line = px.line(
        daily_grouped,
        x="Order Date",
        y=kpi_selection,  # Supports multiple KPI selection
        title="KPI Trends Over Time",
        labels={"Order Date": "Date"},
        template="plotly_white",
        markers=True  # Adds markers for better visibility
    )
    fig_line.update_traces(line=dict(width=2))  # Make the line thicker
    st.plotly_chart(fig_line, use_container_width=True)

    # ---- Top 10 Products Chart ----
    st.subheader("Top 10 Products by KPI")

    for kpi in kpi_selection:
        top_10 = product_grouped.nlargest(10, kpi)

        fig_bar = px.bar(
            top_10,
            x=kpi,
            y="Product Name",
            orientation="h",
            title=f"Top 10 Products by {kpi}",
            labels={kpi: kpi, "Product Name": "Product"},
            color=kpi,
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        st.plotly_chart(fig_bar, use_container_width=True)
