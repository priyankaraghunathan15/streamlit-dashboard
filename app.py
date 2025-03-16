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
    # Convert Order Date to datetime if not already
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ---- Sidebar Filters ----
st.sidebar.title("Filters")

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)

# Filter data by Region
if selected_region != "All":
    df_filtered_region = df_original[df_original["Region"] == selected_region]
else:
    df_filtered_region = df_original

# State Filter
all_states = sorted(df_filtered_region["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)

# Filter data by State
if selected_state != "All":
    df_filtered_state = df_filtered_region[df_filtered_region["State"] == selected_state]
else:
    df_filtered_state = df_filtered_region

# City Filter
all_cities = sorted(df_filtered_state["City"].dropna().unique())
selected_city = st.sidebar.selectbox("Select City", options=["All"] + all_cities)

# Filter data by City
if selected_city != "All":
    df_filtered_city = df_filtered_state[df_filtered_state["City"] == selected_city]
else:
    df_filtered_city = df_filtered_state

# Segment Filter
all_segments = sorted(df_filtered_city["Segment"].dropna().unique())
selected_segment = st.sidebar.selectbox("Select Segment", options=["All"] + all_segments)

# Filter data by Segment
if selected_segment != "All":
    df_filtered_segment = df_filtered_city[df_filtered_city["Segment"] == selected_segment]
else:
    df_filtered_segment = df_filtered_city

# Category Filter
all_categories = sorted(df_filtered_segment["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)

# Filter data by Category
if selected_category != "All":
    df_filtered_category = df_filtered_segment[df_filtered_segment["Category"] == selected_category]
else:
    df_filtered_category = df_filtered_segment

# Ship Mode Filter
all_ship_modes = sorted(df_filtered_category["Ship Mode"].dropna().unique())
selected_ship_mode = st.sidebar.selectbox("Select Ship Mode", options=["All"] + all_ship_modes)

# Filter data by Ship Mode
if selected_ship_mode != "All":
    df_filtered_ship_mode = df_filtered_category[df_filtered_category["Ship Mode"] == selected_ship_mode]
else:
    df_filtered_ship_mode = df_filtered_category

# Final filtered data
df = df_filtered_ship_mode.copy()


# ---- Sidebar Date Range (From and To) ----
if df.empty:
    min_date = df_original["Order Date"].min()
    max_date = df_original["Order Date"].max()
else:
    min_date = df["Order Date"].min()
    max_date = df["Order Date"].max()

from_date = st.sidebar.date_input(
    "From Date", value=min_date, min_value=min_date, max_value=max_date
)
to_date = st.sidebar.date_input(
    "To Date", value=max_date, min_value=min_date, max_value=max_date
)

# Ensure from_date <= to_date
if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

# Apply date range filter
df = df[
    (df["Order Date"] >= pd.to_datetime(from_date))
    & (df["Order Date"] <= pd.to_datetime(to_date))
]

# ---- Page Title ----
st.title("SuperStore KPI Dashboard")

# ---- Custom CSS for KPI Tiles ----
st.markdown(
    """
    <style>
    .kpi-box {
        background-color: #FFFFFF;
        border: 2px solid #EAEAEA;
        border-radius: 8px;
        padding: 16px;
        margin: 8px;
        text-align: center;
    }
    .kpi-title {
        font-weight: 600;
        color: #333333;
        font-size: 22px;  /* Sales and Quantity Sold header size */
        margin-bottom: 8px;
    }
    .kpi-title-small {
    font-size: 18px;
    font-weight: 600;
    color: #333333;
    margin-bottom: 0;
    }
    .kpi-value {
        font-weight: 700;
        font-size: 27px;  /* Sales and Quantity Sold value size */
        color: #1E90FF;
    }
    .margin-rate-positive {
        color: #28A745;  /* Green color for positive margin */
        font-weight: 700;
        font-size: 27px;
        vertical-align: middle;
    }
    .margin-rate-negative {
        color: #DC3545;  /* Red color for negative margin */
        font-weight: 700;
        font-size: 27px;
        vertical-align: middle;
    }
    .quantity-positive {
        color: #28A745;  /* Green for above avg */
        font-weight: 700;
        font-size: 20px;
    }
    .quantity-negative {
        color: #DC3545;  /* Red for below avg */
        font-weight: 700;
        font-size: 20px;
    }
    .kpi-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 16px;
    }
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---- KPI Calculation ----
if df.empty:
    total_sales = 0
    total_quantity = 0
    total_profit = 0
    margin_rate = 0
else:
    total_sales = df["Sales"].sum()
    total_quantity = df["Quantity"].sum()
    total_profit = df["Profit"].sum()
    margin_rate = (total_profit / total_sales) if total_sales != 0 else 0

# Sales format
sales_formatted = f"${total_sales / 1_000_000:,.2f}M"
profit_formatted = f"${total_profit / 1_000_000:,.2f}M"
margin_rate_formatted = f"{(margin_rate * 100):,.2f}%"
if margin_rate >= 0:
    margin_rate_display = f"↑ {margin_rate_formatted}"
else:
    margin_rate_display = f"↓ {margin_rate_formatted}"

# Quantity Sold and Above/Below Average Calculation
avg_quantity_per_transaction = df["Quantity"].mean()
orders_above_avg = df[df["Quantity"] > avg_quantity_per_transaction].shape[0]
orders_below_avg = df[df["Quantity"] < avg_quantity_per_transaction].shape[0]

# ---- KPI Display (Rectangles) ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
with kpi_col1:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Sales</div>
            <div class='kpi-value'>{sales_formatted}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col2:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Quantity Sold</div>
            <div class='kpi-value'>{total_quantity:,.0f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col3:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Profit</div>
            <div class='kpi-value'>{profit_formatted}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col4:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Margin Rate</div>
            <div class="{'margin-rate-positive' if margin_rate >= 0 else 'margin-rate-negative'}">{margin_rate_display}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---- KPI Selection ----
st.subheader("Visualize KPI Across Time")

if df.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    # Radio button for KPI
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
    selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)

    # Aggregation level selection (Week, Month, Quarter, Year)
    aggregation_level = st.selectbox(
        "Select Aggregation Level",
        ["Week", "Month", "Quarter", "Year"]
    )

# ---- Prepare Data for Chart ----
# ---- Add Margin Rate Column to df ----
df["Margin Rate"] = df["Profit"] / df["Sales"].replace(0, 1)  # Avoid dividing by zero

# ---- Prepare Data for Chart ----
if aggregation_level == "Week":
    df["Week"] = df["Order Date"].dt.strftime('%Y-%U')  # Year-Week format
    daily_grouped = df.groupby("Week").agg({selected_kpi: "sum", "Sales": "sum", "Profit": "sum", "Margin Rate": "sum"}).reset_index()

    # Ensure that all weeks are represented, even if there are missing weeks
    all_weeks = pd.date_range(start=df["Order Date"].min(), end=df["Order Date"].max(), freq='W-MON')
    all_weeks_str = all_weeks.strftime('%Y-%U')  # Week format

    # Merge with missing weeks and fill missing values with 0
    all_weeks_df = pd.DataFrame({"Week": all_weeks_str})
    daily_grouped = pd.merge(all_weeks_df, daily_grouped, on="Week", how="left").fillna(0)

    # Create a range slider inside the chart for week selection
    fig_line = px.line(
        daily_grouped,
        x="Week",
        y=selected_kpi,
        title=f"{selected_kpi} Over Week",
        labels={"Week": "Week", selected_kpi: selected_kpi},
        template="plotly_white",
    )

    fig_line.update_layout(
        xaxis=dict(
            rangeslider=dict(visible=True),  # Enable range slider
            type="category",  # Treat week as a category for proper slider behavior
        )
    )

    # Customize the tooltip to include Week, Sales, Profit, and Margin Rate
    fig_line.update_traces(
        hovertemplate="Week: %{x}<br>Sales: %{customdata[0]}<br>Profit: %{customdata[1]}<br>Margin Rate: %{customdata[2]}"
    )

    # Set custom data for hover tooltips (week number, sales, profit, margin rate)
    fig_line.update_traces(customdata=daily_grouped[["Sales", "Profit", "Margin Rate"]].values)

    fig_line.update_layout(height=400, showlegend=False, xaxis_showgrid=False, yaxis_showgrid=False)

elif aggregation_level == "Month":
    df["Month"] = df["Order Date"].dt.to_period("M").astype(str)  # Convert Period to string
    daily_grouped = df.groupby("Month").agg({selected_kpi: "sum", "Sales": "sum", "Profit": "sum", "Margin Rate": "sum"}).reset_index()
    
    # Create the line chart
    fig_line = px.line(
        daily_grouped,
        x="Month",
        y=selected_kpi,
        title=f"{selected_kpi} Over Month",
        labels={"Month": "Month and Year", selected_kpi: selected_kpi},
        template="plotly_white",
    )

    # Manually adjust x-axis tickvals and ticktext to avoid packed labels
    tickvals = daily_grouped["Month"][::3]  # Show every 3rd month
    ticktext = daily_grouped["Month"].apply(lambda x: pd.to_datetime(x).strftime('%b %Y'))[::3]  # Month-Year format
    
    fig_line.update_layout(
        xaxis=dict(
            rangeslider=dict(visible=True),
            tickmode="array",
            tickvals=tickvals,
            ticktext=ticktext,
        )
    )

elif aggregation_level == "Quarter":
    df["Quarter"] = df["Order Date"].dt.to_period("Q").astype(str)  # Convert Period to string
    daily_grouped = df.groupby("Quarter").agg({selected_kpi: "sum", "Sales": "sum", "Profit": "sum", "Margin Rate": "sum"}).reset_index()

    fig_line = px.line(
        daily_grouped,
        x="Quarter",
        y=selected_kpi,
        title=f"{selected_kpi} Over Quarter",
        labels={"Quarter": "Quarter", selected_kpi: selected_kpi},
        template="plotly_white",
    )

    # Add range slider for Quarter aggregation
    fig_line.update_layout(
        xaxis=dict(
            rangeslider=dict(visible=True),
            tickmode="array",
            tickvals=daily_grouped["Quarter"],
            ticktext=daily_grouped["Quarter"],  # Showing Quarter as labels
            type="category",  # Treat as categorical data
        )
    )

elif aggregation_level == "Year":
    df["Year"] = df["Order Date"].dt.year  # Ensure no .5 for year
    daily_grouped = df.groupby("Year").agg({selected_kpi: "sum", "Sales": "sum", "Profit": "sum", "Margin Rate": "sum"}).reset_index()

    fig_line = px.line(
        daily_grouped,
        x="Year",
        y=selected_kpi,
        title=f"{selected_kpi} Over Year",
        labels={"Year": "Year", selected_kpi: selected_kpi},
        template="plotly_white",
    )

    # Add range slider for Year aggregation
    fig_line.update_layout(
        xaxis=dict(
            rangeslider=dict(visible=True),
            tickmode="array",
            tickvals=daily_grouped["Year"],
            ticktext=daily_grouped["Year"].astype(str),  # Showing year as labels
        )
    )



# ---- Prepare Bar Chart (Top 10 Sub-Categories) ----
subcategory_grouped = df.groupby("Sub-Category").agg({
    "Sales": "sum",
    "Quantity": "sum",
    "Profit": "sum"
}).reset_index()
subcategory_grouped["Margin Rate"] = subcategory_grouped["Profit"] / subcategory_grouped["Sales"].replace(0, 1)

# Sort for top 10 by selected KPI
subcategory_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
top_10_subcategories = subcategory_grouped.head(10)

# ---- Side-by-Side Layout for Charts ----
col_left, col_right = st.columns(2)

with col_left:
    # Line Chart
    st.plotly_chart(fig_line, use_container_width=True)

with col_right:
    # Horizontal Bar Chart for Top 10 Sub-Categories
    fig_bar = px.bar(
        top_10_subcategories,
        x=selected_kpi,
        y="Sub-Category",
        orientation="h",
        title=f"Top 10 Product Types by {selected_kpi}",
        labels={selected_kpi: selected_kpi, "Sub-Category": "Sub-Category"},
        color=selected_kpi,
        color_continuous_scale="Blues",
        template="plotly_white",
    )

    # Hide y-axis title (Sub-Category) but keep the labels
    fig_bar.update_layout(
        height=400,
        yaxis={"categoryorder": "total ascending", "title": None}  # Hides the y-axis title
    )

    # Edit the tooltips to show the Sub-Category and the selected KPI value (e.g., Sales)
    fig_bar.update_traces(
        hovertemplate='%{y}: %{x:.2f}'  # Tooltip displays the Sub-Category and the selected KPI (formatted to 2 decimal places)
    )

    # Display the Bar Chart
    st.plotly_chart(fig_bar, use_container_width=True)