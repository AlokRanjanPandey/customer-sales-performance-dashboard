
from pathlib import Path
import io
import re

import pandas as pd
import streamlit as st

# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------
st.set_page_config(
    page_title="Customer Sales Performance Dashboard",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Customer Sales Performance Dashboard")
st.caption(
    "Interactive analysis of sales, profit, products, "
    "countries, and customer segments."
)

EXPECTED_FILE = (
    Path(__file__).parent /
    "Customer_Sales_Performance_Dashboard.xlsx"
)


# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------
def normalize(value):
    return re.sub(
        r"[^a-z0-9]+",
        "",
        str(value).strip().lower()
    )


@st.cache_data(show_spinner=False)
def read_workbook(raw):
    excel = pd.ExcelFile(io.BytesIO(raw))
    sheets = excel.sheet_names

    # Prefer a worksheet named Data; otherwise use the first.
    sheet = next(
        (s for s in sheets if normalize(s) == "data"),
        sheets[0]
    )

    data = pd.read_excel(
        io.BytesIO(raw),
        sheet_name=sheet
    )

    data.columns = [str(c).strip() for c in data.columns]
    data = data.dropna(how="all")

    # Convert numeric-looking text columns to numbers.
    for col in data.columns:
        if data[col].dtype == "object":
            cleaned = (
                data[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("$", "", regex=False)
            )

            numbers = pd.to_numeric(
                cleaned,
                errors="coerce"
            )

            count = data[col].notna().sum()

            if count and numbers.notna().sum() / count >= 0.8:
                data[col] = numbers

    return data, sheets, sheet


def find_column(columns, possible_names):
    lookup = {normalize(c): c for c in columns}

    for name in possible_names:
        if normalize(name) in lookup:
            return lookup[normalize(name)]

    for col in columns:
        if any(
            normalize(name) in normalize(col)
            for name in possible_names
        ):
            return col

    return None


# --------------------------------------------------
# LOAD THE EXCEL WORKBOOK
# --------------------------------------------------
if EXPECTED_FILE.exists():
    try:
        df, sheet_names, used_sheet = read_workbook(
            EXPECTED_FILE.read_bytes()
        )
    except Exception as error:
        st.error(
            "The Excel file could not be opened. "
            "Check that it is a valid .xlsx workbook."
        )
        st.exception(error)
        st.stop()

    with st.sidebar:
        st.success(f"Workbook loaded: {EXPECTED_FILE.name}")
        st.caption(f"Worksheet: {used_sheet}")

        if st.checkbox("Upload a different workbook"):
            upload = st.file_uploader(
                "Choose an Excel file",
                type=["xlsx"]
            )

            if upload is not None:
                try:
                    df, sheet_names, used_sheet = read_workbook(
                        upload.getvalue()
                    )
                except Exception as error:
                    st.error("Could not read the uploaded workbook.")
                    st.exception(error)
                    st.stop()

else:
    st.warning(
        "The Excel workbook is not in the repository root. "
        "Upload it below to continue."
    )

    upload = st.file_uploader(
        "Upload Customer Sales Excel workbook",
        type=["xlsx"]
    )

    if upload is None:
        st.stop()

    try:
        df, sheet_names, used_sheet = read_workbook(
            upload.getvalue()
        )
    except Exception as error:
        st.error("Could not read the uploaded workbook.")
        st.exception(error)
        st.stop()


if df.empty:
    st.warning("The selected worksheet contains no data.")
    st.stop()


# --------------------------------------------------
# DETECT COMMON DATA COLUMNS
# --------------------------------------------------
columns = list(df.columns)

sales = find_column(
    columns, ["Sales", "Total Sales", "Revenue"]
)
profit = find_column(
    columns, ["Profit", "Net Profit"]
)
units = find_column(
    columns, ["Units Sold", "Units", "Quantity"]
)
discounts = find_column(
    columns, ["Discounts", "Discount"]
)
country = find_column(
    columns, ["Country", "Market"]
)
product = find_column(
    columns, ["Product", "Product Name"]
)
segment = find_column(
    columns, ["Segment", "Customer Segment"]
)
year = find_column(columns, ["Year"])
date = find_column(
    columns, ["Date", "Order Date", "Sales Date"]
)

numeric_columns = df.select_dtypes(
    include="number"
).columns.tolist()

categorical_columns = [
    col for col in columns
    if col not in numeric_columns
    and df[col].nunique(dropna=True) <= 100
]


# --------------------------------------------------
# SIDEBAR FILTERS
# --------------------------------------------------
filtered = df.copy()

with st.sidebar:
    st.header("🔎 Dashboard Filters")

    filter_columns = list(
        dict.fromkeys(
            col for col in [country, product, segment, year]
            if col is not None
        )
    )

    for col in filter_columns:
        options = sorted(
            filtered[col].dropna().unique().tolist(),
            key=lambda value: str(value)
        )

        selected = st.multiselect(
            f"Select {col}",
            options=options,
            default=[]
        )

        if selected:
            filtered = filtered[
                filtered[col].isin(selected)
            ]

    st.caption(
        f"{len(filtered):,} rows selected "
        f"out of {len(df):,} total rows."
    )


if filtered.empty:
    st.warning(
        "No records match the selected filters. "
        "Clear one or more filters in the sidebar."
    )
    st.stop()


# --------------------------------------------------
# KPI CARDS
# --------------------------------------------------
st.subheader("📌 Key Performance Indicators")

kpi_columns = st.columns(4)

kpi_specs = [
    ("Total Sales", sales),
    ("Total Profit", profit),
    ("Units Sold", units),
    ("Total Discounts", discounts),
]

displayed = 0

for label, col in kpi_specs:
    if (
        col is not None
        and col in filtered.columns
        and pd.api.types.is_numeric_dtype(filtered[col])
    ):
        value = filtered[col].sum()
        kpi_columns[displayed].metric(
            label,
            f"{value:,.2f}"
        )
        displayed += 1

if displayed == 0:
    kpi_columns[0].metric("Total Records", f"{len(filtered):,}")
    kpi_columns[1].metric("Total Columns", f"{len(columns):,}")
    kpi_columns[2].metric(
        "Numeric Fields", f"{len(numeric_columns):,}"
    )
    kpi_columns[3].metric(
        "Category Fields", f"{len(categorical_columns):,}"
    )

if (
    sales is not None
    and profit is not None
    and sales in filtered.columns
    and profit in filtered.columns
    and filtered[sales].sum() != 0
):
    margin = (
        filtered[profit].sum()
        / filtered[sales].sum()
        * 100
    )

    st.caption(
        f"**Profit margin for the selected data:** "
        f"{margin:.2f}%"
    )


# --------------------------------------------------
# CATEGORY COMPARISON CHARTS
# --------------------------------------------------
st.divider()

left, right = st.columns(2)

groups = list(
    dict.fromkeys(
        [
            col for col in [
                country, product, segment, year
            ]
            if col is not None
        ] + categorical_columns
    )
)

measures = list(
    dict.fromkeys(
        [
            col for col in [
                sales, profit, units, discounts
            ]
            if col is not None and col in numeric_columns
        ] + numeric_columns
    )
)

with left:
    st.subheader("📊 Compare Categories")

    if groups and measures:
        group = st.selectbox(
            "Group categories by",
            options=groups,
            key="comparison_group"
        )

        measure = st.selectbox(
            "Choose a metric",
            options=measures,
            key="comparison_measure"
        )

        limit = st.slider(
            "Number of categories",
            min_value=3,
            max_value=20,
            value=10
        )

        chart_data = (
            filtered.groupby(group, dropna=False)[measure]
            .sum()
            .sort_values(ascending=False)
            .head(limit)
        )

        st.bar_chart(chart_data)

    else:
        st.info(
            "A category column and a numeric column "
            "are needed to create this chart."
        )


with right:
    st.subheader("💰 Profit by Category")

    if profit is not None and profit in filtered.columns and groups:
        profit_group = st.selectbox(
            "Choose a category",
            options=groups,
            key="profit_group"
        )

        profit_data = (
            filtered.groupby(
                profit_group,
                dropna=False
            )[profit]
            .sum()
            .sort_values(ascending=False)
            .head(10)
        )

        st.bar_chart(profit_data)

    else:
        st.info(
            "A numeric Profit column and a category "
            "column are needed for this chart."
        )


# --------------------------------------------------
# SALES TREND OVER TIME
# --------------------------------------------------
if date is not None and date in filtered.columns:
    parsed_dates = pd.to_datetime(
        filtered[date],
        errors="coerce"
    )

    trend_measure = sales or (
        numeric_columns[0] if numeric_columns else None
    )

    if parsed_dates.notna().any() and trend_measure:
        st.divider()
        st.subheader("📈 Trend Over Time")

        trend_data = filtered.copy()
        trend_data["_date"] = parsed_dates

        trend_data = (
            trend_data
            .dropna(subset=["_date"])
            .set_index("_date")
        )

        monthly_trend = (
            trend_data[trend_measure]
            .resample("MS")
            .sum()
        )

        st.line_chart(monthly_trend)


# --------------------------------------------------
# DATA TABLE AND DOWNLOAD
# --------------------------------------------------
st.divider()
st.subheader("📋 Explore Sales Data")

with st.expander("View data records", expanded=True):
    st.dataframe(
        filtered,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        label="⬇️ Download Filtered Data (CSV)",
        data=filtered.to_csv(index=False).encode("utf-8"),
        file_name="customer_sales_filtered.csv",
        mime="text/csv"
    )


st.caption(
    "Customer Sales Performance Dashboard | "
    "Built with Streamlit and pandas."
)

