import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Set page configuration for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# 1. Load the Data
@st.cache_data
def load_data():
    # Adjust the path if needed (e.g., "data/Sample - Superstore.xlsx")
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    # Ensure Order Date is datetime
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df = load_data()

st.title("SuperStore KPI Dashboard")

# ------------------------------------------------------------------------------
# 2. Cascading Filters
# ------------------------------------------------------------------------------
# Region Filter
regions = sorted(df["Region"].dropna().unique().tolist())
selected_region = st.selectbox("Select Region:", options=["All"] + regions)
if selected_region != "All":
    df = df[df["Region"] == selected_region]

# State Filter (depends on Region selection)
states = sorted(df["State"].dropna().unique().tolist())
selected_state = st.selectbox("Select State:", options=["All"] + states)
if selected_state != "All":
    df = df[df["State"] == selected_state]

# Category Filter
categories = sorted(df["Category"].dropna().unique().tolist())
selected_category = st.selectbox("Select Category:", options=["All"] + categories)
if selected_category != "All":
    df = df[df["Category"] == selected_category]

# Sub-Category Filter
subcats = sorted(df["Sub-Category"].dropna().unique().tolist())
selected_subcat = st.selectbox("Select Sub-Category:", options=["All"] + subcats)
if selected_subcat != "All":
    df = df[df["Sub-Category"] == selected_subcat]

# ------------------------------------------------------------------------------
# 3. KPI Tiles
# ------------------------------------------------------------------------------
# Calculate metrics
total_sales = df["Sales"].sum()
total_quantity = df["Quantity"].sum()
total_profit = df["Profit"].sum()
margin_rate = total_profit / total_sales if total_sales != 0 else 0

# Display KPIs in a row
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("Sales", f"${total_sales:,.2f}")
kpi2.metric("Quantity Sold", f"{total_quantity:,.0f}")
kpi3.metric("Profit", f"${total_profit:,.2f}")
kpi4.metric("Margin Rate", f"{margin_rate*100:,.2f}%")

# ------------------------------------------------------------------------------
# 4. Line Chart with Time Range + KPI Selection
# ------------------------------------------------------------------------------
# Date Range Slider (min to max Order Date in the filtered df)
if not df.empty:
    min_date = df["Order Date"].min()
    max_date = df["Order Date"].max()
else:
    # Fallback in case there's no data after filtering
    min_date = datetime(2020, 1, 1)
    max_date = datetime(2020, 12, 31)

date_range = st.slider(
    "Select Date Range:",
    min_value=min_date,
    max_value=max_date,
    value=(min_date, max_date),
    format="YYYY-MM-DD"
)

# Filter the dataframe by the selected date range
df_time_filtered = df[(df["Order Date"] >= date_range[0]) & (df["Order Date"] <= date_range[1])]

# Let the user pick which KPI to plot
kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)

# Prepare data for line chart
# We'll group by Order Date. For Margin Rate, we do sum(profit)/sum(sales) for each date.
if not df_time_filtered.empty:
    daily_grouped = df_time_filtered.groupby("Order Date").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"]
else:
    daily_grouped = pd.DataFrame(columns=["Order Date", "Sales", "Quantity", "Profit", "Margin Rate"])

# Plotly line chart
if not daily_grouped.empty:
    fig_line = px.line(
        daily_grouped,
        x="Order Date",
        y=selected_kpi,
        title=f"{selected_kpi} Over Time",
        labels={"Order Date": "Date", selected_kpi: selected_kpi}
    )
    fig_line.update_layout(height=400)
    st.plotly_chart(fig_line, use_container_width=True)
else:
    st.warning("No data available for the selected filters and date range.")

# ------------------------------------------------------------------------------
# 5. Horizontal Bar Chart for Top 10 Items by Sales
# ------------------------------------------------------------------------------
# We'll use df_time_filtered so it respects all filters + date range
if not df_time_filtered.empty:
    # Group by Product Name, sum sales, sort descending
    product_sales = (
        df_time_filtered.groupby("Product Name")["Sales"]
        .sum()
        .reset_index()
        .sort_values(by="Sales", ascending=False)
        .head(10)
    )

    fig_bar = px.bar(
        product_sales,
        x="Sales",
        y="Product Name",
        orientation="h",
        title="Top 10 Products by Sales",
        labels={"Sales": "Sales", "Product Name": "Product"},
        color="Sales",
        color_continuous_scale="Blues"
    )
    fig_bar.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
    st.plotly_chart(fig_bar, use_container_width=True)
else:
    st.warning("No data to display for Top 10 items.")
