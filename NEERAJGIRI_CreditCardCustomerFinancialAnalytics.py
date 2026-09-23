"""Credit Card Customer & Financial Behavior Analytics.

AICTE | IBM SkillsBuild Data Analytics with AI Internship 2026
Student: NEERAJGIRI

Run the dashboard with:
    streamlit run NEERAJGIRI_CreditCardCustomerFinancialAnalytics.py
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
import seaborn as sns
import streamlit as st


PROJECT_DIR = Path(__file__).resolve().parent
RAW_DATA_PATHS = [PROJECT_DIR / "data" / "raw" / "CC GENERAL.csv", PROJECT_DIR / "CC GENERAL.csv"]
PROCESSED_DATA_PATH = PROJECT_DIR / "data" / "processed" / "cleaned_credit_card_customers.csv"
REPORT_PATH = PROJECT_DIR / "NEERAJGIRI_CreditCardCustomerFinancialAnalytics_ProjectReport.docx"

REQUIRED_COLUMNS = [
    "CUST_ID", "BALANCE", "BALANCE_FREQUENCY", "PURCHASES", "ONEOFF_PURCHASES",
    "INSTALLMENTS_PURCHASES", "CASH_ADVANCE", "PURCHASES_FREQUENCY",
    "ONEOFF_PURCHASES_FREQUENCY", "PURCHASES_INSTALLMENTS_FREQUENCY",
    "CASH_ADVANCE_FREQUENCY", "CASH_ADVANCE_TRX", "PURCHASES_TRX", "CREDIT_LIMIT",
    "PAYMENTS", "MINIMUM_PAYMENTS", "PRC_FULL_PAYMENT", "TENURE",
]
NUMERIC_COLUMNS = [column for column in REQUIRED_COLUMNS if column != "CUST_ID"]


def find_data_path() -> Path:
    """Find the supplied dataset using project-relative paths only."""
    for path in RAW_DATA_PATHS:
        if path.exists():
            return path
    raise FileNotFoundError(
        "Dataset not found. Download 'CC GENERAL.csv' from Kaggle and place it in data/raw/."
    )


def load_raw_data(path: Path | None = None) -> pd.DataFrame:
    data_path = path or find_data_path()
    data = pd.read_csv(data_path)
    missing_columns = sorted(set(REQUIRED_COLUMNS) - set(data.columns))
    if missing_columns:
        raise ValueError(f"Dataset is missing required columns: {missing_columns}")
    return data[REQUIRED_COLUMNS].copy()


def clean_data(raw: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, object]]:
    """Clean types and missing values without deleting valid extreme behavior."""
    data = raw.copy()
    rows_before = len(data)
    missing_before = int(data.isna().sum().sum())
    duplicate_count = int(data.duplicated().sum())
    data["CUST_ID"] = data["CUST_ID"].astype(str).str.strip()
    for column in NUMERIC_COLUMNS:
        data[column] = pd.to_numeric(data[column], errors="coerce")
    data = data.drop_duplicates().copy()
    numeric_missing = data[NUMERIC_COLUMNS].isna().sum()
    imputation_details = {}
    for column, count in numeric_missing.items():
        if count:
            median_value = float(data[column].median())
            data[column] = data[column].fillna(median_value)
            imputation_details[column] = {"missing": int(count), "method": "median", "value": median_value}
    invalid_numeric_count = int((data[NUMERIC_COLUMNS] < 0).sum().sum())
    for column in NUMERIC_COLUMNS:
        data[column] = data[column].clip(lower=0)
    report = {
        "rows_before": rows_before,
        "rows_after": len(data),
        "columns": len(data.columns),
        "missing_before": missing_before,
        "missing_after": int(data.isna().sum().sum()),
        "duplicate_count": duplicate_count,
        "invalid_numeric_count": invalid_numeric_count,
        "imputation_details": imputation_details,
        "operations": [
            "Validated the expected one-row-per-customer grain using unique CUST_ID values.",
            "Converted analytical fields to numeric types and removed exact duplicate rows.",
            "Imputed missing CREDIT_LIMIT and MINIMUM_PAYMENTS values with their observed medians.",
            "Checked for negative financial values; none were present in the supplied data.",
            "Retained extreme valid observations and flagged them with IQR/percentile methods.",
        ],
    }
    return data, report


def safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    return numerator.div(denominator.replace(0, np.nan)).replace([np.inf, -np.inf], np.nan).fillna(0)


def engineer_features(data: pd.DataFrame) -> pd.DataFrame:
    data = data.copy()
    data["CREDIT_UTILIZATION"] = safe_divide(data["BALANCE"], data["CREDIT_LIMIT"])
    data["PAYMENT_RATIO"] = safe_divide(data["PAYMENTS"], data["BALANCE"])
    data["PURCHASE_TO_LIMIT"] = safe_divide(data["PURCHASES"], data["CREDIT_LIMIT"])
    data["CASH_ADVANCE_INTENSITY"] = safe_divide(data["CASH_ADVANCE"], data["CREDIT_LIMIT"])
    data["AVG_PURCHASE_VALUE"] = safe_divide(data["PURCHASES"], data["PURCHASES_TRX"])
    data["INSTALLMENT_PURCHASE_SHARE"] = safe_divide(data["INSTALLMENTS_PURCHASES"], data["PURCHASES"])
    return data


def add_segments(data: pd.DataFrame) -> pd.DataFrame:
    data = data.copy()
    high_purchase = data["PURCHASES"] >= data["PURCHASES"].quantile(0.75)
    active = data["PURCHASES_FREQUENCY"] >= data["PURCHASES_FREQUENCY"].quantile(0.75)
    high_utilization = data["CREDIT_UTILIZATION"] >= data["CREDIT_UTILIZATION"].quantile(0.75)
    cash_heavy = data["CASH_ADVANCE_INTENSITY"] >= data["CASH_ADVANCE_INTENSITY"].quantile(0.75)
    installment_focused = (
        data["INSTALLMENT_PURCHASE_SHARE"] >= 0.60
    ) & (data["PURCHASES"] > 0)
    low_engagement = (data["PURCHASES_FREQUENCY"] <= 0.25) & (data["PURCHASES_TRX"] <= data["PURCHASES_TRX"].median())
    data["SEGMENT"] = np.select(
        [high_purchase & active, high_utilization, cash_heavy, installment_focused, low_engagement],
        ["High-Value Active", "High-Utilization", "Cash-Advance Heavy", "Installment-Focused", "Low-Engagement"],
        default="Balanced Usage",
    )
    return data


def prepare_data(raw_path: Path | None = None, save_processed: bool = False) -> Tuple[pd.DataFrame, Dict[str, object]]:
    raw = load_raw_data(raw_path)
    cleaned, quality = clean_data(raw)
    data = add_segments(engineer_features(cleaned))
    if save_processed:
        PROCESSED_DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
        data.to_csv(PROCESSED_DATA_PATH, index=False)
    quality["unique_customers"] = int(data["CUST_ID"].nunique())
    quality["raw_path"] = str((raw_path or find_data_path()).relative_to(PROJECT_DIR))
    return data, quality


def calculate_kpis(data: pd.DataFrame) -> Dict[str, float]:
    if data.empty:
        return {"Total Customers": 0}
    return {
        "Total Customers": int(data["CUST_ID"].nunique()),
        "Average Credit Limit": data["CREDIT_LIMIT"].mean(),
        "Median Credit Limit": data["CREDIT_LIMIT"].median(),
        "Average Balance": data["BALANCE"].mean(),
        "Average Purchases": data["PURCHASES"].mean(),
        "Average Payments": data["PAYMENTS"].mean(),
        "Average Cash Advance": data["CASH_ADVANCE"].mean(),
        "Average Purchase Frequency": data["PURCHASES_FREQUENCY"].mean(),
        "Average Credit Utilization": data["CREDIT_UTILIZATION"].mean(),
        "Average Tenure": data["TENURE"].mean(),
        "Average Full Payment Percentage": data["PRC_FULL_PAYMENT"].mean(),
    }


def segment_summary(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame()
    summary = data.groupby("SEGMENT", as_index=False).agg(
        Customers=("CUST_ID", "nunique"),
        Purchases=("PURCHASES", "mean"),
        Balance=("BALANCE", "mean"),
        Credit_Limit=("CREDIT_LIMIT", "mean"),
        Payments=("PAYMENTS", "mean"),
        Utilization=("CREDIT_UTILIZATION", "mean"),
        Cash_Advance=("CASH_ADVANCE", "mean"),
        Purchase_Frequency=("PURCHASES_FREQUENCY", "mean"),
        Tenure=("TENURE", "mean"),
    )
    summary["Customer Share"] = summary["Customers"] / summary["Customers"].sum()
    return summary.sort_values("Customers", ascending=False)


def iqr_outlier_counts(data: pd.DataFrame, columns: Iterable[str]) -> pd.DataFrame:
    rows = []
    for column in columns:
        q1, q3 = data[column].quantile([0.25, 0.75])
        upper = q3 + 1.5 * (q3 - q1)
        rows.append({"Metric": column, "Upper IQR Threshold": upper, "Flagged Customers": int((data[column] > upper).sum())})
    return pd.DataFrame(rows)


def money(value: float) -> str:
    return f"${value:,.0f}"


def pct(value: float) -> str:
    return f"{value:.1%}"


def empty_message() -> None:
    st.warning("No customers match the selected filters. Widen one or more ranges to continue.")


def chart_layout(fig: go.Figure, title: str | None = None) -> go.Figure:
    layout_options = {
        "height": 390,
        "margin": dict(l=20, r=20, t=55, b=20),
        "template": "plotly_white",
    }
    if title is not None:
        layout_options["title"] = title
    if fig.layout.yaxis.title.text in (None, "count"):
        layout_options["yaxis"] = {"title": "Customer Count"}
    fig.update_layout(**layout_options)
    return fig


def overview_charts(data: pd.DataFrame) -> Tuple[go.Figure, go.Figure, go.Figure, go.Figure, go.Figure, go.Figure]:
    segment = px.bar(segment_summary(data), x="SEGMENT", y="Customers", color="SEGMENT", title="Customer Count by Segment", labels={"SEGMENT": "Customer Segment", "Customers": "Customer Count"})
    utilization = px.histogram(data, x="CREDIT_UTILIZATION", nbins=35, title="Credit Utilization Distribution", labels={"CREDIT_UTILIZATION": "Balance / Credit Limit"})
    purchases_limit = px.scatter(data, x="CREDIT_LIMIT", y="PURCHASES", color="SEGMENT", opacity=0.65, hover_data=["CUST_ID"], title="Purchases vs Credit Limit by Customer Segment", labels={"CREDIT_LIMIT": "Credit Limit", "PURCHASES": "Purchases", "SEGMENT": "Customer Segment"})
    balance_limit = px.scatter(data, x="CREDIT_LIMIT", y="BALANCE", color="SEGMENT", opacity=0.65, hover_data=["CUST_ID"], title="Balance vs Credit Limit by Customer Segment", labels={"CREDIT_LIMIT": "Credit Limit", "BALANCE": "Balance", "SEGMENT": "Customer Segment"})
    credit_limit = px.histogram(data, x="CREDIT_LIMIT", nbins=30, title="Credit Limit Distribution", labels={"CREDIT_LIMIT": "Credit Limit", "count": "Customer Count"})
    balance = px.histogram(data, x="BALANCE", nbins=30, title="Balance Distribution", labels={"BALANCE": "Balance", "count": "Customer Count"})
    return tuple(chart_layout(fig) for fig in (segment, utilization, purchases_limit, balance_limit, credit_limit, balance))


def business_insights(data: pd.DataFrame) -> List[str]:
    if data.empty:
        return []
    segment = segment_summary(data)
    largest = segment.iloc[0]
    high_util = data["CREDIT_UTILIZATION"].ge(data["CREDIT_UTILIZATION"].quantile(0.75)).mean()
    zero_purchase = data["PURCHASES"].eq(0).mean()
    purchase_corr = data[["PURCHASES", "CREDIT_LIMIT"]].corr().iloc[0, 1]
    return [
        f"The largest filtered segment is {largest['SEGMENT']} with {int(largest['Customers']):,} customers ({pct(largest['Customer Share'])}). This supports segment-specific engagement planning.",
        f"{pct(high_util)} of filtered customers are at or above the filtered 75th-percentile utilization level. The observed pattern suggests an opportunity to monitor utilization and payment activity together.",
        f"{pct(zero_purchase)} of filtered customers report no purchases in the observation period. This is a low-engagement signal for targeted activation analysis, not a judgment about customer quality.",
        f"The Pearson correlation between purchases and credit limit is {purchase_corr:.2f}. This is an observed association and should not be interpreted as causation.",
    ]


def render_sidebar(data: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filters")
    selected_segments = st.sidebar.multiselect("Customer Segment", sorted(data["SEGMENT"].unique()), default=sorted(data["SEGMENT"].unique()))
    util_min, util_max = float(data["CREDIT_UTILIZATION"].min()), float(data["CREDIT_UTILIZATION"].max())
    purchase_min, purchase_max = float(data["PURCHASES"].min()), float(data["PURCHASES"].max())
    limit_min, limit_max = float(data["CREDIT_LIMIT"].min()), float(data["CREDIT_LIMIT"].max())
    tenure_min, tenure_max = int(data["TENURE"].min()), int(data["TENURE"].max())
    util_range = st.sidebar.slider("Credit Utilization Range", util_min, max(util_max, util_min + 0.01), (util_min, util_max), format="%.2f")
    purchase_range = st.sidebar.slider("Purchase Range", purchase_min, max(purchase_max, purchase_min + 1), (purchase_min, purchase_max), format="$%.0f")
    limit_range = st.sidebar.slider("Credit Limit Range", limit_min, max(limit_max, limit_min + 1), (limit_min, limit_max), format="$%.0f")
    tenure_range = st.sidebar.slider("Tenure Range", tenure_min, tenure_max, (tenure_min, tenure_max))
    mask = (
        data["SEGMENT"].isin(selected_segments)
        & data["CREDIT_UTILIZATION"].between(*util_range)
        & data["PURCHASES"].between(*purchase_range)
        & data["CREDIT_LIMIT"].between(*limit_range)
        & data["TENURE"].between(*tenure_range)
    )
    return data.loc[mask].copy()


def render_kpis(data: pd.DataFrame, labels: List[str]) -> None:
    kpis = calculate_kpis(data)
    columns = st.columns(len(labels))
    for column, label in zip(columns, labels):
        value = kpis.get(label, 0)
        if label == "Total Customers":
            display = f"{int(value):,}"
        elif "Percentage" in label or "Frequency" in label or "Utilization" in label:
            display = pct(value)
        else:
            display = money(value)
        column.metric(label, display)


def render_overview(data: pd.DataFrame) -> None:
    st.subheader("Executive Overview")
    render_kpis(data, ["Total Customers", "Average Credit Limit", "Average Balance", "Average Purchases", "Average Payments", "Average Credit Utilization"])
    if data.empty:
        empty_message()
        return
    charts = overview_charts(data)
    left, right = st.columns(2)
    left.plotly_chart(charts[0], use_container_width=True, key="overview_segments")
    right.plotly_chart(charts[1], use_container_width=True, key="overview_utilization")
    left.plotly_chart(charts[2], use_container_width=True, key="overview_purchases_limit")
    right.plotly_chart(charts[3], use_container_width=True, key="overview_balance_limit")
    left.plotly_chart(charts[4], use_container_width=True, key="overview_credit_limit")
    right.plotly_chart(charts[5], use_container_width=True, key="overview_balance")
    st.subheader("Key Insights")
    for insight in business_insights(data):
        st.info(insight)


def render_spending(data: pd.DataFrame) -> None:
    st.subheader("Spending Analysis")
    if data.empty:
        empty_message()
        return
    left, right = st.columns(2)
    left.plotly_chart(chart_layout(px.histogram(data, x="PURCHASES", nbins=35, title="Purchases Distribution", labels={"PURCHASES": "Purchases", "count": "Customer Count"})), use_container_width=True, key="spending_purchases")
    right.plotly_chart(chart_layout(px.histogram(data, x="PURCHASES_FREQUENCY", nbins=25, title="Purchases Frequency Distribution", labels={"PURCHASES_FREQUENCY": "Purchase Frequency", "count": "Customer Count"})), use_container_width=True, key="spending_frequency")
    left.plotly_chart(chart_layout(px.scatter(data, x="PURCHASES", y="ONEOFF_PURCHASES", color="SEGMENT", title="One-off Purchases by Customer Segment", labels={"PURCHASES": "Total Purchases", "ONEOFF_PURCHASES": "One-off Purchases", "SEGMENT": "Customer Segment"})), use_container_width=True, key="spending_oneoff")
    right.plotly_chart(chart_layout(px.scatter(data, x="PURCHASES", y="INSTALLMENTS_PURCHASES", color="SEGMENT", title="Installment Purchases by Customer Segment", labels={"PURCHASES": "Total Purchases", "INSTALLMENTS_PURCHASES": "Installment Purchases", "SEGMENT": "Customer Segment"})), use_container_width=True, key="spending_installments")
    st.plotly_chart(chart_layout(px.scatter(data, x="TENURE", y="PURCHASES", color="SEGMENT", title="Tenure vs Spending by Customer Segment", labels={"TENURE": "Tenure (months)", "PURCHASES": "Purchases", "SEGMENT": "Customer Segment"})), use_container_width=True, key="spending_tenure")
    st.plotly_chart(chart_layout(px.box(data, x="SEGMENT", y="CASH_ADVANCE", color="SEGMENT", title="Cash Advance by Segment", labels={"SEGMENT": "Customer Segment", "CASH_ADVANCE": "Cash Advance", "color": "Customer Segment"})), use_container_width=True, key="spending_cash_advance")


def render_payment_credit(data: pd.DataFrame) -> None:
    st.subheader("Payment & Credit Analysis")
    if data.empty:
        empty_message()
        return
    left, right = st.columns(2)
    left.plotly_chart(chart_layout(px.scatter(data, x="PURCHASES", y="PAYMENTS", color="SEGMENT", title="Purchases vs Payments by Segment", labels={"PURCHASES": "Purchases", "PAYMENTS": "Payments", "SEGMENT": "Customer Segment"})), use_container_width=True, key="payment_purchases")
    right.plotly_chart(chart_layout(px.scatter(data, x="CREDIT_LIMIT", y="BALANCE", color="SEGMENT", title="Balance vs Credit Limit by Segment", labels={"CREDIT_LIMIT": "Credit Limit", "BALANCE": "Balance", "SEGMENT": "Customer Segment"})), use_container_width=True, key="payment_balance_limit")
    left.plotly_chart(chart_layout(px.histogram(data, x="PRC_FULL_PAYMENT", nbins=25, title="Full Payment Percentage Distribution", labels={"PRC_FULL_PAYMENT": "Full Payment Percentage", "count": "Customer Count"})), use_container_width=True, key="payment_full_percentage")
    right.plotly_chart(chart_layout(px.histogram(data, x="PAYMENT_RATIO", nbins=35, title="Payment Ratio Distribution", labels={"PAYMENT_RATIO": "Payment Ratio", "count": "Customer Count"})), use_container_width=True, key="payment_ratio")
    correlation_columns = ["BALANCE", "PURCHASES", "CASH_ADVANCE", "CREDIT_LIMIT", "PAYMENTS", "PURCHASES_TRX", "CREDIT_UTILIZATION"]
    correlation = data[correlation_columns].corr()
    st.plotly_chart(chart_layout(px.imshow(correlation, text_auto=".2f", color_continuous_scale="RdBu_r", zmin=-1, zmax=1, title="Correlation Heatmap of Key Financial Measures", labels={"x": "Financial Measure", "y": "Financial Measure", "color": "Correlation"})), use_container_width=True, key="payment_correlation")
    with st.expander("Matplotlib / Seaborn correlation view"):
        figure, axis = plt.subplots(figsize=(9, 5))
        sns.heatmap(correlation, cmap="vlag", center=0, annot=True, fmt=".2f", ax=axis)
        axis.set_title("Correlation Heatmap")
        st.pyplot(figure, clear_figure=True)
    st.dataframe(data[["CUST_ID", "MINIMUM_PAYMENTS", "PAYMENTS", "BALANCE", "PAYMENT_RATIO"]].sort_values("MINIMUM_PAYMENTS", ascending=False).head(20), use_container_width=True, hide_index=True)


def render_segments(data: pd.DataFrame) -> None:
    st.subheader("Customer Segments")
    if data.empty:
        empty_message()
        return
    summary = segment_summary(data)
    left, right = st.columns(2)
    left.plotly_chart(chart_layout(px.bar(summary, x="SEGMENT", y="Customers", color="SEGMENT", title="Customer Count by Segment", labels={"SEGMENT": "Customer Segment", "Customers": "Customer Count", "color": "Customer Segment"})), use_container_width=True, key="segments_size")
    right.plotly_chart(chart_layout(px.bar(summary, x="SEGMENT", y="Purchases", color="SEGMENT", title="Average Purchases by Segment", labels={"SEGMENT": "Customer Segment", "Purchases": "Average Purchases", "color": "Customer Segment"})), use_container_width=True, key="segments_purchases")
    st.plotly_chart(chart_layout(px.bar(summary, x="SEGMENT", y=["Balance", "Credit_Limit", "Payments"], barmode="group", title="Average Balance, Credit Limit, and Payments by Segment", labels={"SEGMENT": "Customer Segment", "value": "Average Amount", "variable": "Financial Measure"})), use_container_width=True, key="segments_comparison")
    st.dataframe(summary.style.format({"Customer Share": "{:.1%}", "Purchases": "${:,.2f}", "Balance": "${:,.2f}", "Credit_Limit": "${:,.2f}", "Payments": "${:,.2f}", "Utilization": "{:.2%}", "Cash_Advance": "${:,.2f}", "Purchase_Frequency": "{:.2%}", "Tenure": "{:.1f}"}), use_container_width=True, hide_index=True)


def render_explorer(data: pd.DataFrame) -> None:
    st.subheader("Customer Explorer")
    if data.empty:
        empty_message()
        return
    options = data["CUST_ID"].sort_values().tolist()
    selected = st.selectbox("Select a customer", options)
    customer = data.loc[data["CUST_ID"] == selected].iloc[0]
    explorer_metrics = {
        "Credit Limit": customer.CREDIT_LIMIT,
        "Balance": customer.BALANCE,
        "Purchases": customer.PURCHASES,
        "Payments": customer.PAYMENTS,
        "Credit Utilization": customer.CREDIT_UTILIZATION,
        "Purchase Frequency": customer.PURCHASES_FREQUENCY,
    }
    columns = st.columns(len(explorer_metrics))
    for column, (label, value) in zip(columns, explorer_metrics.items()):
        display = pct(value) if "Utilization" in label or "Frequency" in label else money(value)
        column.metric(label, display)
    st.dataframe(pd.DataFrame({"Metric": ["Credit Limit", "Balance", "Purchases", "Payments", "Credit Utilization", "Purchase Frequency", "Cash Advance", "Tenure"], "Value": [money(customer.CREDIT_LIMIT), money(customer.BALANCE), money(customer.PURCHASES), money(customer.PAYMENTS), pct(customer.CREDIT_UTILIZATION), pct(customer.PURCHASES_FREQUENCY), money(customer.CASH_ADVANCE), f"{customer.TENURE:.0f} months"]}), use_container_width=True, hide_index=True)


def render_data_quality(quality: Dict[str, object], data: pd.DataFrame) -> None:
    st.subheader("Data Quality")
    cols = st.columns(5)
    cols[0].metric("Rows before cleaning", f"{quality['rows_before']:,}")
    cols[1].metric("Rows after cleaning", f"{quality['rows_after']:,}")
    cols[2].metric("Columns", quality["columns"])
    cols[3].metric("Missing before", quality["missing_before"])
    cols[4].metric("Missing after", quality["missing_after"])
    st.write(f"Unique customers: **{quality['unique_customers']:,}** | Exact duplicate rows: **{quality['duplicate_count']}** | Raw source: `{quality['raw_path']}`")
    st.write("Major cleaning operations")
    for operation in quality["operations"]:
        st.write(f"- {operation}")
    with st.expander("Outlier flags (retained for analysis)"):
        st.dataframe(iqr_outlier_counts(data, ["BALANCE", "PURCHASES", "CASH_ADVANCE", "CREDIT_UTILIZATION", "PAYMENTS"]), use_container_width=True, hide_index=True)


def generate_project_report(data: pd.DataFrame, quality: Dict[str, object], output_path: Path = REPORT_PATH) -> None:
    """Generate a DOCX report; screenshots are inserted when available."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    styles = document.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10)
    styles["Normal"].paragraph_format.space_after = Pt(6)
    styles["Normal"].paragraph_format.line_spacing = 1.08
    styles["Heading 1"].font.name = "Aptos Display"
    styles["Heading 1"].font.size = Pt(16)
    styles["Heading 1"].font.color.rgb = RGBColor(31, 78, 121)
    styles["Heading 1"].paragraph_format.space_before = Pt(12)
    styles["Heading 1"].paragraph_format.space_after = Pt(5)
    styles["Heading 2"].font.name = "Aptos Display"
    styles["Heading 2"].font.size = Pt(12)
    styles["Heading 2"].font.color.rgb = RGBColor(47, 84, 150)

    def add_bullet(text: str) -> None:
        document.add_paragraph(text, style="List Bullet")

    def add_takeaway(text: str) -> None:
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(3)
        run = paragraph.add_run("In short: ")
        run.bold = True
        run.font.color.rgb = RGBColor(31, 78, 121)
        paragraph.add_run(text)

    def add_metric(label: str, value: str, meaning: str) -> None:
        paragraph = document.add_paragraph(style="List Bullet")
        run = paragraph.add_run(f"{label}: {value}. ")
        run.bold = True
        paragraph.add_run(meaning)

    kpis = calculate_kpis(data)
    summary = segment_summary(data)
    largest_segment = summary.iloc[0]
    zero_purchase_share = data["PURCHASES"].eq(0).mean()
    high_utilization_share = data["CREDIT_UTILIZATION"].ge(data["CREDIT_UTILIZATION"].quantile(0.75)).mean()
    purchase_limit_correlation = data[["PURCHASES", "CREDIT_LIMIT"]].corr().iloc[0, 1]

    title = document.add_heading("Credit Card Customer & Financial Behavior Analytics", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph = document.add_paragraph("AICTE | IBM SkillsBuild Data Analytics with AI Internship 2026\nStudent: NEERAJGIRI")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    one_line = document.add_paragraph()
    one_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = one_line.add_run("What this project does in one line: ")
    run.bold = True
    one_line.add_run("It turns credit-card customer behavior into clear metrics, segments, charts, and business insights.")
    document.add_page_break()

    document.add_heading("1. Problem Statement", level=1)
    document.add_paragraph("Credit-card customers do not all use their accounts in the same way. Some spend regularly, some mainly use installments, some use cash advances, and some show very little purchase activity. Looking at these behaviors together helps an analyst describe the portfolio more clearly.")
    document.add_paragraph("This project turns customer-level account measures into an understandable analytical story for portfolio monitoring and engagement planning. It describes patterns in the data; it does not make credit, fraud, or personal-finance decisions.")
    add_takeaway("The project makes a large customer table easier to understand and discuss using evidence from the data.")

    document.add_heading("2. Project Objectives", level=1)
    for item in [
        "Inspect the supplied dataset instead of assuming its size, columns, or quality.",
        "Clean the data transparently while keeping the original raw CSV unchanged.",
        "Create useful financial behavior measures, such as credit utilization and payment ratio.",
        "Calculate key performance indicators, or KPIs, from the active data selection.",
        "Compare understandable, rule-based customer segments.",
        "Identify unusual observations without calling customers fraudulent or bad.",
        "Present the analysis in a professional Streamlit dashboard and DOCX report.",
    ]:
        add_bullet(item)
    add_takeaway("The objective is a complete data analytics workflow, from raw data to business communication.")

    document.add_heading("3. Dataset Description", level=1)
    document.add_paragraph("The project uses Kaggle's Credit Card Dataset for Clustering. The supplied file contains one row for each credit-card customer, so each row represents one customer record rather than one transaction.")
    add_metric("8,950 customers", f"{len(data):,}", "This is the size of the customer population analyzed; it is large enough to compare broad behavior patterns.")
    add_metric("18 source columns", f"{len(REQUIRED_COLUMNS):,}", "These columns describe balances, purchases, payments, credit limits, frequencies, cash advances, and tenure.")
    add_metric("Unique customer IDs", f"{data['CUST_ID'].nunique():,}", "The matching row and ID counts verify the expected one-row-per-customer grain.")
    document.add_paragraph("The source includes one identifier column, CUST_ID, and 17 numerical behavior columns. There are no source categorical behavior fields; customer segments are created later for analysis.")
    document.add_heading("Dataset Source", level=2)
    document.add_paragraph("Kaggle: https://www.kaggle.com/datasets/arjunbhasin2013/ccdata")
    add_takeaway("The dataset is a customer-level snapshot, so the analysis compares customer profiles rather than individual purchases.")

    document.add_heading("4. Tools and Technologies", level=1)
    for item in [
        "Python: the programming language used for the complete project.",
        "Pandas: reads, cleans, transforms, and summarizes the table.",
        "NumPy: supports safe numerical calculations.",
        "Matplotlib and Seaborn: create traditional analytical charts, including the correlation heatmap.",
        "Plotly: creates interactive charts in the dashboard.",
        "Streamlit: provides the interactive dashboard interface.",
        "python-docx: creates this formatted project report.",
    ]:
        add_bullet(item)
    add_takeaway("The project uses a standard Python analytics stack, with Streamlit providing the user-facing application.")

    document.add_heading("5. Data Quality and Cleaning", level=1)
    document.add_paragraph("Data cleaning means checking whether the values are usable and making only documented changes. The raw CSV is preserved in data/raw/, so the original input remains available for review.")
    document.add_heading("What was found", level=2)
    add_metric("Missing values before cleaning", f"{quality['missing_before']:,}", "The missing-value count shows how much information needed attention before analysis.")
    add_metric("Missing values after cleaning", f"{quality['missing_after']:,}", "Zero missing values means the dashboard can calculate all displayed metrics consistently.")
    add_metric("Duplicate rows", f"{quality['duplicate_count']:,}", "No duplicate rows were found, so no customer record was repeated exactly.")
    add_metric("Negative numeric values", f"{quality['invalid_numeric_count']:,}", "No negative financial values were found; this supports retaining the observed numerical records.")
    document.add_heading("How it was handled", level=2)
    for item in [
        "Numeric fields were converted carefully so the calculations use numbers rather than text.",
        "One missing CREDIT_LIMIT value was filled with the observed median credit limit of $3,000. This is a middle-value replacement that avoids inventing an extreme limit.",
        "313 missing MINIMUM_PAYMENTS values were filled with the observed median of $312.34. The method keeps all customer rows available while making the assumption explicit.",
        "Exact duplicate rows would be removed because they would count the same record twice; none were present in this dataset.",
        "Valid extreme values were retained. They were flagged with the IQR method rather than automatically deleted.",
    ]:
        add_bullet(item)
    document.add_heading("Plain-language definition: IQR", level=2)
    document.add_paragraph("IQR, or interquartile range, describes the middle 50% of values. A value far above the upper IQR boundary is flagged as unusual for review, but it is not automatically considered wrong.")
    add_takeaway("Cleaning removed ambiguity from missing fields without hiding genuine high-balance, high-purchase, or high-cash-advance behavior.")

    document.add_heading("6. Feature Engineering", level=1)
    document.add_paragraph("Feature engineering means creating useful measures from the original columns. These ratios put raw amounts into context and make customer behaviors easier to compare.")
    feature_rows = [
        ("Credit utilization", "BALANCE / CREDIT_LIMIT", "The share of the assigned limit represented by the balance. A higher value means more of the available limit is being used."),
        ("Payment ratio", "PAYMENTS / BALANCE", "Payments compared with the balance. A zero balance is handled safely so the dashboard does not divide by zero."),
        ("Purchase-to-limit ratio", "PURCHASES / CREDIT_LIMIT", "Purchases compared with the assigned limit, helping compare spending across customers with different limits."),
        ("Cash-advance intensity", "CASH_ADVANCE / CREDIT_LIMIT", "Cash advances compared with the assigned limit."),
        ("Average purchase value", "PURCHASES / PURCHASES_TRX", "Approximate purchase amount per purchase transaction."),
        ("Installment purchase share", "INSTALLMENTS_PURCHASES / PURCHASES", "The share of purchases made through installments. Customers with no purchases safely receive zero."),
    ]
    table = document.add_table(rows=1, cols=3)
    table.style = "Light Shading Accent 1"
    for cell, text in zip(table.rows[0].cells, ["Metric", "Formula", "Plain-language meaning"]):
        cell.text = text
    for metric, formula, meaning in feature_rows:
        cells = table.add_row().cells
        cells[0].text, cells[1].text, cells[2].text = metric, formula, meaning
    add_takeaway("The derived metrics make raw account amounts comparable across customers and support the later segments and charts.")

    document.add_heading("7. Executive KPI Analysis", level=1)
    document.add_paragraph("A KPI is a key performance indicator: a short measure used to understand the main shape of the data. The dashboard recalculates these values whenever filters change, so they are never hard-coded.")
    add_metric("Average credit limit", money(kpis["Average Credit Limit"]), "This is the typical assigned spending capacity across the analyzed customers.")
    add_metric("Median credit limit", money(kpis["Median Credit Limit"]), "The median is the middle customer value and is less affected by unusually large limits than an average.")
    add_metric("Average balance", money(kpis["Average Balance"]), "This shows the typical recorded balance in the customer population.")
    add_metric("Average purchases", money(kpis["Average Purchases"]), "This summarizes typical purchase value during the dataset's observation period.")
    add_metric("Average payments", money(kpis["Average Payments"]), "This shows typical payment activity and gives context for balances and purchases.")
    add_metric("Average cash advance", money(kpis["Average Cash Advance"]), "This shows how much cash-advance activity appears on average.")
    add_metric("Average purchase frequency", pct(kpis["Average Purchase Frequency"]), "This is the average share of the observation period with purchase activity.")
    add_metric("Average credit utilization", pct(kpis["Average Credit Utilization"]), "This indicates how much of the available credit is represented by balances on average.")
    add_metric("Average tenure", f"{kpis['Average Tenure']:.1f} months", "This describes how long customers have been represented in the account history.")
    add_metric("Average full-payment percentage", pct(kpis["Average Full Payment Percentage"]), "This is the average share of periods in which the full payment amount was made.")
    add_takeaway("The KPI panel provides a quick portfolio snapshot, while filters allow the same measures to be compared across subsets.")

    document.add_heading("8. Exploratory Data Analysis", level=1)
    document.add_paragraph("Exploratory data analysis, or EDA, means looking for distributions and relationships before drawing conclusions. The dashboard includes:")
    for item in [
        "Credit limit, balance, purchase, purchase-frequency, utilization, and full-payment distributions.",
        "Purchases versus credit limit and balance versus credit limit scatter charts.",
        "Payments versus purchases and tenure versus spending relationships.",
        "One-off purchases, installment purchases, and cash-advance behavior.",
        "A correlation heatmap. Correlation shows how two measures move together; it does not prove that one causes the other.",
    ]:
        add_bullet(item)
    add_takeaway("EDA shows where customer behavior is concentrated and where measures move together, while keeping interpretation descriptive.")

    document.add_heading("9. Customer Segmentation", level=1)
    document.add_paragraph("Segmentation means grouping customers with similar observed behavior so the portfolio can be discussed more clearly. This project uses transparent rules before any optional machine-learning approach.")
    for item in [
        "High-Value Active: customers with both high purchase amounts and high purchase frequency relative to this dataset.",
        "High-Utilization: customers at or above the dataset's 75th-percentile credit utilization.",
        "Cash-Advance Heavy: customers at or above the dataset's 75th-percentile cash-advance intensity.",
        "Installment-Focused: customers with at least 60% of purchases represented by installments and at least one purchase.",
        "Low-Engagement: customers with low purchase frequency and no more than the median number of purchase transactions.",
        "Balanced Usage: customers who do not meet one of the specific rules above.",
    ]:
        add_bullet(item)
    document.add_paragraph("These labels are analytical descriptions, not rankings. They do not identify fraudulent customers, bad customers, or credit defaulters.")
    add_takeaway(f"The largest segment in the analyzed dataset is {largest_segment['SEGMENT']} with {int(largest_segment['Customers']):,} customers ({pct(largest_segment['Customer Share'])}); this matters because it identifies the most common observed behavior pattern.")

    document.add_heading("10. Segment Analysis", level=1)
    document.add_paragraph("For every segment, the dashboard compares customer count, share of customers, average purchases, balance, credit limit, payments, utilization, cash advance, purchase frequency, and tenure. These comparisons help explain how the groups differ without declaring one segment universally better or worse.")
    segment_table = document.add_table(rows=1, cols=5)
    segment_table.style = "Light Shading Accent 1"
    for cell, text in zip(segment_table.rows[0].cells, ["Segment", "Customers", "Share", "Avg purchases", "Avg utilization"]):
        cell.text = text
    for _, row in summary.iterrows():
        cells = segment_table.add_row().cells
        cells[0].text = str(row["SEGMENT"])
        cells[1].text = f"{int(row['Customers']):,}"
        cells[2].text = pct(row["Customer Share"])
        cells[3].text = money(row["Purchases"])
        cells[4].text = pct(row["Utilization"])
    add_takeaway("Segment comparison turns many individual measures into a practical portfolio view while avoiding unsupported labels about customer quality.")

    document.add_heading("11. Outlier Analysis", level=1)
    document.add_paragraph("An outlier is an observation that is unusually high or low compared with the rest of the dataset. Outlier analysis is useful for review, but unusual does not automatically mean incorrect or harmful.")
    outliers = iqr_outlier_counts(data, ["BALANCE", "PURCHASES", "CASH_ADVANCE", "CREDIT_UTILIZATION", "PAYMENTS"])
    for _, row in outliers.iterrows():
        add_bullet(f"{row['Metric']}: {int(row['Flagged Customers']):,} customers are above the upper IQR threshold of {row['Upper IQR Threshold']:.2f}. This highlights records for further review.")
    document.add_paragraph("The project uses neutral terms such as high-utilization customer, unusual spending pattern, high cash-advance usage, and low payment activity. It does not make unsupported fraud or default claims.")
    add_takeaway("Outlier flags prioritize analytical attention while preserving valid observations in the dataset.")

    document.add_heading("12. Dashboard Overview", level=1)
    for item in [
        "Executive Overview: filtered KPIs, segment distribution, utilization distribution, and core relationship charts.",
        "Spending Analysis: purchase distributions, purchase frequency, one-off and installment purchases, cash advance, and tenure versus spending.",
        "Payment & Credit: payments versus purchases, utilization, balance versus credit limit, full-payment percentage, payment ratio, and correlation heatmap.",
        "Customer Segments: segment size, average behavior, and a detailed comparison table.",
        "Customer Explorer: selected-customer metrics for credit limit, balance, purchases, payments, utilization, purchase frequency, cash advance, and tenure.",
        "Data Quality: rows before and after cleaning, missing values, duplicates, cleaning operations, and outlier flags.",
        "Sidebar filters: segment, utilization range, purchase range, credit-limit range, and tenure range. All relevant charts, KPIs, and tables respond to the selection.",
    ]:
        add_bullet(item)
    add_takeaway("The dashboard turns the analysis into an interactive experience where readers can move from portfolio-level patterns to one customer record.")

    document.add_heading("13. Business Insights", level=1)
    insights = [
        ("Largest observed segment", f"{largest_segment['SEGMENT']} contains {int(largest_segment['Customers']):,} customers, or {pct(largest_segment['Customer Share'])} of the dataset.", "This identifies the most common behavior pattern in the portfolio.", "Use the segment profile as a starting point for targeted engagement analysis."),
        ("Utilization concentration", f"{pct(high_utilization_share)} of customers are at or above the dataset's 75th-percentile utilization level.", "Higher utilization deserves context alongside payment behavior and credit limits.", "Monitor utilization and payment measures together rather than using either measure alone."),
        ("No-purchase activity", f"{pct(zero_purchase_share)} of customers have recorded purchases equal to zero.", "This is a measurable low-engagement signal in this observation period.", "Consider activation analysis while recognizing that the snapshot does not explain why purchases are zero."),
        ("Purchases and credit limit", f"The observed Pearson correlation is {purchase_limit_correlation:.2f}.", "The value describes association in this dataset, not a causal effect.", "Use the relationship for descriptive portfolio analysis and validate it with time-based or business data before acting."),
    ]
    for finding, evidence, implication, action in insights:
        document.add_heading(finding, level=2)
        add_bullet(f"Finding: {finding}.")
        add_bullet(f"Supporting evidence: {evidence}")
        add_bullet(f"Business implication: {implication}")
        add_bullet(f"Possible business action: {action}")
    add_takeaway("The findings are evidence-based descriptions of this dataset, not predictions or causal claims.")

    document.add_heading("14. Business Recommendations", level=1)
    for item in [
        "Use the segment profiles to design and compare engagement experiments rather than sending one message to every customer.",
        "Review utilization together with payments, balances, and credit limits so one metric is not interpreted in isolation.",
        "Study cash-advance-heavy behavior with additional context before deciding on any operational response.",
        "Treat low purchase activity as a question for further investigation, not as a judgment about the customer.",
        "Add time-series data and repeat the analysis before making decisions about change, trend, or intervention impact.",
    ]:
        add_bullet(item)
    add_takeaway("Recommendations focus on responsible analysis and follow-up evidence, not personal financial advice.")

    document.add_heading("15. Data Dictionary", level=1)
    document.add_paragraph("A data dictionary is a plain-language guide to the columns used in the project.")
    dictionary_table = document.add_table(rows=1, cols=3)
    dictionary_table.style = "Light Shading Accent 1"
    for cell, text in zip(dictionary_table.rows[0].cells, ["Column", "Type", "Meaning"]):
        cell.text = text
    meanings = {
        "CUST_ID": "Customer identifier.", "BALANCE": "Statement balance.", "BALANCE_FREQUENCY": "How often a balance is recorded.", "PURCHASES": "Total purchases.", "ONEOFF_PURCHASES": "Purchases made as one-off transactions.", "INSTALLMENTS_PURCHASES": "Purchases made through installments.", "CASH_ADVANCE": "Cash advance amount.", "PURCHASES_FREQUENCY": "Frequency of purchase activity.", "ONEOFF_PURCHASES_FREQUENCY": "Frequency of one-off purchases.", "PURCHASES_INSTALLMENTS_FREQUENCY": "Frequency of installment purchases.", "CASH_ADVANCE_FREQUENCY": "Frequency of cash advances.", "CASH_ADVANCE_TRX": "Number of cash-advance transactions.", "PURCHASES_TRX": "Number of purchase transactions.", "CREDIT_LIMIT": "Assigned credit limit.", "PAYMENTS": "Payments made.", "MINIMUM_PAYMENTS": "Recorded minimum payment amount.", "PRC_FULL_PAYMENT": "Percentage of periods paid in full.", "TENURE": "Account tenure in months.",
    }
    for column in REQUIRED_COLUMNS:
        cells = dictionary_table.add_row().cells
        cells[0].text = column
        cells[1].text = str(data[column].dtype)
        cells[2].text = meanings[column]

    document.add_heading("16. Conclusion", level=1)
    document.add_paragraph("This project demonstrates a complete data analytics workflow: inspect the raw data, clean it transparently, create interpretable measures, calculate KPIs, compare customer segments, visualize relationships, and communicate business implications through an interactive dashboard.")
    add_takeaway("The result is a reproducible portfolio project that connects technical analysis with accessible business communication.")

    document.add_heading("17. Limitations", level=1)
    for item in [
        "The dataset is a historical snapshot and does not provide a time series.",
        "It does not include demographic context, business outcomes, or labels for causal analysis.",
        "Correlations describe association and do not prove that one behavior causes another.",
        "Segment rules use thresholds derived from this portfolio, so they may change on a different dataset.",
        "Median imputation makes the table complete but introduces an assumption for the missing values.",
    ]:
        add_bullet(item)
    add_takeaway("The dashboard is useful for descriptive analysis, but stronger decisions require richer and time-based data.")

    document.add_heading("18. Future Improvements", level=1)
    for item in [
        "Add monthly account history so trends and changes can be measured.",
        "Validate segment thresholds with business experts and monitor whether they remain useful.",
        "Compare the transparent rules with an optional clustering analysis without letting machine learning dominate the story.",
        "Add automated tests for data loading, formulas, filters, and empty-result behavior.",
        "Add controlled deployment, access governance, and monitoring for a production setting.",
    ]:
        add_bullet(item)
    add_takeaway("The next improvements would add context, validation, and operational reliability while preserving interpretability.")

    document.add_heading("19. How to Run the Project", level=1)
    for item in [
        "Install the dependencies with: pip install -r requirements.txt",
        "Place the manually downloaded Kaggle file at: data/raw/CC GENERAL.csv",
        "Start the dashboard with: streamlit run NEERAJGIRI_CreditCardCustomerFinancialAnalytics.py",
        "Generate this report with: python NEERAJGIRI_CreditCardCustomerFinancialAnalytics.py --generate-report",
    ]:
        add_bullet(item)
    add_takeaway("A new user can reproduce the dashboard and report from the documented project files and the manually supplied dataset.")

    document.add_heading("20. Dashboard Screenshots", level=1)
    for filename in ["dashboard_overview.png", "spending_analysis.png", "payment_credit_analysis.png", "customer_segments.png"]:
        image_path = PROJECT_DIR / "screenshots" / filename
        if image_path.exists():
            label = document.add_paragraph()
            label.add_run(filename).bold = True
            document.add_picture(str(image_path), width=Inches(6.3))
        else:
            document.add_paragraph(f"Screenshot unavailable at report-generation time: {filename}")
    document.save(output_path)


def run_report_command() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--generate-report", action="store_true")
    args, _ = parser.parse_known_args()
    if args.generate_report:
        data, quality = prepare_data(save_processed=True)
        generate_project_report(data, quality)
        print(f"Report written to {REPORT_PATH}")


def main() -> None:
    st.set_page_config(page_title="Credit Card Customer & Financial Behavior Analytics", page_icon="CC", layout="wide")
    st.title("Credit Card Customer & Financial Behavior Analytics")
    st.caption("Data Analytics Portfolio Project — NEERAJGIRI")
    try:
        data, quality = prepare_data(save_processed=True)
    except (FileNotFoundError, ValueError) as error:
        st.error(str(error))
        st.stop()
    filtered = render_sidebar(data)
    st.caption(f"Showing {len(filtered):,} of {len(data):,} customers")
    tabs = st.tabs(["Executive Overview", "Spending Analysis", "Payment & Credit", "Customer Segments", "Customer Explorer", "Data Quality"])
    with tabs[0]: render_overview(filtered)
    with tabs[1]: render_spending(filtered)
    with tabs[2]: render_payment_credit(filtered)
    with tabs[3]: render_segments(filtered)
    with tabs[4]: render_explorer(filtered)
    with tabs[5]: render_data_quality(quality, filtered)
    st.markdown("---")
    st.caption("Analytical results are calculated from the supplied dataset. Associations are descriptive and do not establish causation.")


if __name__ == "__main__":
    run_report_command()
    main()
