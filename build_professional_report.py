"""Build the submission-ready Word report from the actual project data and dashboard assets."""
from __future__ import annotations

import importlib.util
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "NEERAJGIRI_CreditCardCustomerFinancialBehaviorAnalytics_Report.docx"
ASSET_DIR = ROOT / "report_assets"
ASSET_DIR.mkdir(exist_ok=True)
SOURCE = ROOT / "NEERAJGIRI_CreditCardCustomerFinancialAnalytics.py"


def load_project_module():
    spec = importlib.util.spec_from_file_location("credit_analytics", SOURCE)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


analytics = load_project_module()
RAW = analytics.load_raw_data(ROOT / "CC GENERAL.csv")
DATA, QUALITY = analytics.prepare_data(ROOT / "CC GENERAL.csv")
KPIS = analytics.calculate_kpis(DATA)
SEGMENTS = analytics.segment_summary(DATA)

NAVY = RGBColor(23, 55, 94)
BLUE = RGBColor(31, 78, 121)
TEAL = RGBColor(0, 112, 128)
GREY = RGBColor(89, 89, 89)
LIGHT_BLUE = "D9EAF7"
LIGHT_GREY = "F2F5F7"


def money(value: float) -> str:
    return f"${value:,.2f}"


def pct(value: float) -> str:
    return f"{value:.1%}"


def shade(cell, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        properties.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, bold: bool = False, color: RGBColor | None = None) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(str(text))
    run.bold = bold
    if color:
        run.font.color.rgb = color
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def table(document: Document, headers: Iterable[str], rows: Iterable[Iterable[str]], widths=None):
    result = document.add_table(rows=1, cols=len(list(headers)))
    # headers is materialized again because it may be a generator.
    headers = list(headers)
    result.style = "Table Grid"
    result.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell, text in zip(result.rows[0].cells, headers):
        set_cell_text(cell, text, bold=True, color=RGBColor(255, 255, 255))
        shade(cell, "1F4E79")
    for row in rows:
        cells = result.add_row().cells
        for cell, text in zip(cells, row):
            set_cell_text(cell, text)
    if widths:
        for row in result.rows:
            for cell, width in zip(row.cells, widths):
                cell.width = Inches(width)
    document.add_paragraph().paragraph_format.space_after = Pt(0)
    return result


def add_field(paragraph, instruction: str) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])


def add_caption(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(8)
    run = paragraph.add_run(text)
    run.italic = True
    run.font.size = Pt(9)
    run.font.color.rgb = GREY


def add_bullet(document: Document, text: str, level: int = 0) -> None:
    paragraph = document.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.add_run(text)


def add_takeaway(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(8)
    run = paragraph.add_run("In short: ")
    run.bold = True
    run.font.color.rgb = TEAL
    paragraph.add_run(text)


def add_insight_box(document: Document, title: str, body: str) -> None:
    box = document.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = box.cell(0, 0)
    shade(cell, "EAF4F7")
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(2)
    run = paragraph.add_run(title)
    run.bold = True
    run.font.color.rgb = NAVY
    cell.add_paragraph(body)
    document.add_paragraph().paragraph_format.space_after = Pt(0)


def heading(document: Document, number: str, title: str, level: int = 1) -> None:
    label = f"{number}. {title}" if number else title
    document.add_heading(label, level=level)


def make_charts() -> dict[str, Path]:
    sns.set_theme(style="whitegrid", palette="deep")
    paths = {}
    chart_specs = [
        ("credit_limit", DATA["CREDIT_LIMIT"], "Distribution of Customer Credit Limits", "Credit limit", "Customers", "hist"),
        ("balance", DATA["BALANCE"], "Distribution of Customer Balances", "Balance", "Customers", "hist"),
        ("purchases", DATA["PURCHASES"], "Distribution of Customer Purchases", "Purchases", "Customers", "hist"),
        ("utilization", DATA["CREDIT_UTILIZATION"], "Distribution of Credit Utilization", "Balance / Credit limit", "Customers", "hist"),
        ("cash_advance", DATA["CASH_ADVANCE"], "Distribution of Cash Advance Usage", "Cash advance", "Customers", "hist"),
        ("full_payment", DATA["PRC_FULL_PAYMENT"], "Distribution of Full Payment Percentage", "Full payment percentage", "Customers", "hist"),
    ]
    for name, series, title, xlabel, ylabel, kind in chart_specs:
        fig, axis = plt.subplots(figsize=(8.5, 4.2), dpi=170)
        clipped = series.clip(upper=series.quantile(0.99))
        axis.hist(clipped, bins=30, color="#1F4E79", edgecolor="white")
        axis.set_title(title, loc="left", fontsize=13, fontweight="bold", color="#17375E")
        axis.set_xlabel(xlabel)
        axis.set_ylabel(ylabel)
        axis.spines[["top", "right"]].set_visible(False)
        fig.tight_layout()
        path = ASSET_DIR / f"{name}.png"
        fig.savefig(path, bbox_inches="tight")
        plt.close(fig)
        paths[name] = path

    fig, axis = plt.subplots(figsize=(8.5, 4.6), dpi=170)
    sns.scatterplot(data=DATA.sample(min(2500, len(DATA)), random_state=42), x="CREDIT_LIMIT", y="PURCHASES", hue="SEGMENT", alpha=.55, s=25, ax=axis)
    axis.set_title("Purchases vs Credit Limit", loc="left", fontsize=13, fontweight="bold", color="#17375E")
    axis.set_xlabel("Credit limit")
    axis.set_ylabel("Purchases")
    axis.legend(title="Segment", fontsize=7, title_fontsize=8, loc="upper right")
    fig.tight_layout()
    paths["purchases_limit"] = ASSET_DIR / "purchases_limit.png"
    fig.savefig(paths["purchases_limit"], bbox_inches="tight")
    plt.close(fig)

    corr_cols = ["BALANCE", "PURCHASES", "CASH_ADVANCE", "CREDIT_LIMIT", "PAYMENTS", "PURCHASES_TRX", "CREDIT_UTILIZATION"]
    fig, axis = plt.subplots(figsize=(8.5, 5), dpi=170)
    sns.heatmap(DATA[corr_cols].corr(), cmap="RdBu_r", center=0, annot=True, fmt=".2f", ax=axis)
    axis.set_title("Correlation Heatmap of Key Financial Measures", loc="left", fontsize=13, fontweight="bold", color="#17375E")
    fig.tight_layout()
    paths["correlation"] = ASSET_DIR / "correlation.png"
    fig.savefig(paths["correlation"], bbox_inches="tight")
    plt.close(fig)

    fig, axis = plt.subplots(figsize=(8.8, 3.7), dpi=170)
    axis.axis("off")
    boxes = ["CSV Dataset", "Data Validation", "Cleaning", "Feature Engineering", "EDA", "Segmentation", "Dashboard", "Insights"]
    x_positions = np.linspace(.07, .93, len(boxes))
    for index, (x, label) in enumerate(zip(x_positions, boxes)):
        axis.text(x, .5, label, ha="center", va="center", fontsize=9, color="white", fontweight="bold", bbox=dict(boxstyle="round,pad=.65", facecolor="#1F4E79" if index % 2 == 0 else "#008080", edgecolor="none"))
        if index < len(boxes) - 1:
            axis.annotate("", xy=(x_positions[index + 1] - .055, .5), xytext=(x + .055, .5), arrowprops=dict(arrowstyle="->", color="#7F8C8D", lw=1.5))
    axis.set_title("Project Architecture and Analytical Flow", loc="left", fontsize=13, fontweight="bold", color="#17375E")
    fig.tight_layout()
    paths["architecture"] = ASSET_DIR / "architecture.png"
    fig.savefig(paths["architecture"], bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return paths


def configure_document(document: Document) -> None:
    styles = document.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor(45, 45, 45)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.08
    for style_name, size, color in [("Heading 1", 17, NAVY), ("Heading 2", 12, BLUE), ("Heading 3", 10, TEAL)]:
        style = styles[style_name]
        style.font.name = "Aptos Display"
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.font.bold = True
        style.paragraph_format.space_before = Pt(12 if style_name == "Heading 1" else 7)
        style.paragraph_format.space_after = Pt(5)
    for section in document.sections:
        section.top_margin = Inches(.72)
        section.bottom_margin = Inches(.65)
        section.left_margin = Inches(.72)
        section.right_margin = Inches(.72)
        header = section.header.paragraphs[0]
        header.text = "NEERAJGIRI  |  Credit Card Customer & Financial Behavior Analytics"
        header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        header.runs[0].font.size = Pt(8)
        header.runs[0].font.color.rgb = GREY
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run("AICTE | IBM SkillsBuild Data Analytics with AI Internship 2026   •   Page ")
        add_field(footer, "PAGE")
        for run in footer.runs:
            run.font.size = Pt(8)
            run.font.color.rgb = GREY


def add_toc(document: Document) -> None:
    heading(document, "", "Table of Contents")
    paragraph = document.add_paragraph()
    add_field(paragraph, 'TOC \\o "1-3" \\h \\z \\u')
    document.add_paragraph("The table of contents updates automatically when the document is opened in Microsoft Word. The section list below is included as a stable reading guide.")
    for number, title in [
        ("1", "Executive Summary"), ("2", "Problem Statement"), ("3", "Project Objectives"), ("4", "Dataset Description"), ("5", "Data Dictionary"), ("6", "Data Quality Assessment"), ("7", "Data Cleaning & Preprocessing"), ("8", "Exploratory Data Analysis"), ("9", "Feature Engineering"), ("10", "Customer Segmentation"), ("11", "Financial Behavior Analysis"), ("12", "Key Findings & Insights"), ("13", "Business Recommendations"), ("14", "Dashboard Overview"), ("15", "Technical Architecture"), ("16", "Project Workflow"), ("17", "Limitations"), ("18", "Future Scope"), ("19", "Conclusion"), ("20", "References")]:
        add_bullet(document, f"{number}. {title}")


def build_report() -> None:
    charts = make_charts()
    document = Document()
    configure_document(document)

    # Cover page.
    cover_title = document.add_heading("Credit Card Customer & Financial Behavior Analytics", 0)
    cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_title.runs[0].font.name = "Aptos Display"
    cover_title.runs[0].font.size = Pt(25)
    cover_title.runs[0].font.color.rgb = NAVY
    subtitle = document.add_paragraph("DATA ANALYTICS PROJECT")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = Pt(13)
    subtitle.runs[0].font.bold = True
    subtitle.runs[0].font.color.rgb = TEAL
    document.add_picture(str(charts["architecture"]), width=Inches(6.6))
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_table = table(document, ["Prepared by", "Technology", "Dataset"], [["NEERAJGIRI", "Python | Pandas | NumPy | Matplotlib | Seaborn | Plotly | Streamlit", "Credit Card Dataset for Clustering"]], widths=[1.4, 3.4, 1.6])
    document.add_paragraph("Project title: Credit Card Customer & Financial Behavior Analytics")
    document.add_paragraph()
    statement = document.add_paragraph()
    statement.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = statement.add_run("What this project does in one line: ")
    run.bold = True
    run.font.color.rgb = TEAL
    statement.add_run("It transforms customer-level credit-card behavior into practical metrics, segments, visual analysis, and an interactive business dashboard.")
    document.add_paragraph("AICTE | IBM SkillsBuild Data Analytics with AI Internship 2026").alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_page_break()

    add_toc(document)
    document.add_page_break()

    # Executive summary.
    heading(document, "1", "Executive Summary")
    document.add_paragraph("This project analyzes customer-level credit-card behavior using the Kaggle Credit Card Dataset for Clustering. The analysis focuses on balances, purchases, installment purchases, one-off purchases, cash advances, payments, credit limits, utilization, payment frequency, full-payment behavior, and account tenure.")
    document.add_paragraph("The verified dataset contains one row per customer. It includes 8,950 unique customers and 18 source columns. This matters because the dashboard compares customer profiles rather than individual transaction records.")
    add_insight_box(document, "Portfolio snapshot", f"The source contains {len(DATA):,} customers, {len(RAW.columns):,} original variables, {QUALITY['missing_before']:,} missing cells before cleaning, and {QUALITY['missing_after']:,} after cleaning. The zero-after-cleaning result means the dashboard can calculate its metrics consistently.")
    document.add_paragraph("The project uses transparent data cleaning, ratio-based feature engineering, exploratory analysis, IQR-based outlier flagging, and rule-based customer segmentation. No predictive model or causal claim is presented. The Streamlit dashboard allows a reader to filter the portfolio and explore both segment-level and customer-level behavior.")
    document.add_paragraph("The analysis can support descriptive business decisions such as designing segment-specific engagement experiments, monitoring credit utilization alongside payment activity, and identifying customer groups that merit deeper review. It should not be used by itself to approve credit, label customers as risky, or make personal financial recommendations.")
    add_takeaway(document, "The project turns a wide customer table into an evidence-based, interactive portfolio view while keeping the analytical assumptions visible.")

    heading(document, "2", "Problem Statement")
    document.add_paragraph("Credit-card organizations collect many measures of customer behavior, but raw measures can be difficult to compare. A balance of $2,000 may represent very different behavior for customers with different credit limits. Similarly, a payment amount is easier to interpret when viewed alongside purchases, balances, and utilization.")
    document.add_paragraph("The business problem addressed here is how to organize these measures into clear customer behavior patterns. The analysis examines spending intensity, purchase frequency, cash-advance use, installment behavior, payment activity, credit utilization, and tenure in one consistent framework.")
    add_bullet(document, "The problem is analytical: understand patterns and differences in the observed portfolio.")
    add_bullet(document, "The output is descriptive: KPIs, charts, segments, insights, and dashboard exploration.")
    add_bullet(document, "The project deliberately avoids unsupported claims about fraud, default, causation, or customer quality.")
    add_takeaway(document, "The project addresses the gap between raw financial behavior fields and a clear, decision-supporting portfolio story.")

    heading(document, "3", "Project Objectives")
    objectives = ["Analyze customer financial behavior.", "Understand credit utilization and credit-limit context.", "Analyze purchase, installment, one-off, and cash-advance patterns.", "Study payment activity and full-payment behavior.", "Identify understandable customer segments using the existing rule-based method.", "Compare financial characteristics across segments.", "Detect unusual or extreme behavior without deleting valid observations.", "Build an interactive Streamlit dashboard with responsive filters.", "Generate evidence-based business insights and practical recommendations."]
    for item in objectives:
        add_bullet(document, item)
    add_takeaway(document, "The objectives cover the full analytics lifecycle while keeping the project understandable as a data analytics project rather than an ML project.")

    heading(document, "4", "Dataset Description")
    document.add_paragraph("Dataset name: Credit Card Dataset for Clustering. Source: Kaggle, https://www.kaggle.com/datasets/arjunbhasin2013/ccdata. The project reads the supplied CC GENERAL.csv file and keeps the original raw file unchanged.")
    dataset_rows = [
        ["Records", f"{len(RAW):,}", "Each record represents one customer; the count defines the population being compared."],
        ["Source variables", f"{len(RAW.columns):,}", "These fields describe account balances, spending, payments, frequencies, and tenure."],
        ["Unique customers", f"{RAW['CUST_ID'].nunique():,}", "Unique IDs match the row count, verifying the one-row-per-customer grain."],
        ["Categorical fields", "CUST_ID", "The identifier is text; analytical behavior fields are numerical."],
        ["Tenure range", f"{int(RAW.TENURE.min())} to {int(RAW.TENURE.max())} months", "The account-history window varies between six and twelve months."],
    ]
    table(document, ["Dataset item", "Observed value", "Why it matters"], dataset_rows, widths=[1.5, 1.5, 3.7])
    document.add_paragraph("The largest recorded source values include a balance of $19,043.14, purchases of $49,039.57, cash advances of $47,137.21, payments of $50,721.48, and a credit limit of $30,000. These are not automatically errors; they are retained as valid observations and considered in outlier analysis.")
    add_takeaway(document, "The dataset is a static, customer-level portfolio snapshot with enough behavioral detail for descriptive comparison, but not enough context for causal or predictive conclusions.")

    heading(document, "5", "Data Dictionary")
    document.add_paragraph("The table below explains every source column in everyday language. Percentages and frequencies are represented as numeric values between zero and one where the dataset uses that scale.")
    meanings = {
        "CUST_ID": "Customer identifier", "BALANCE": "Statement balance", "BALANCE_FREQUENCY": "How often a balance is recorded", "PURCHASES": "Total purchases", "ONEOFF_PURCHASES": "Purchases made as one-off transactions", "INSTALLMENTS_PURCHASES": "Purchases made through installments", "CASH_ADVANCE": "Cash advance amount", "PURCHASES_FREQUENCY": "Frequency of purchase activity", "ONEOFF_PURCHASES_FREQUENCY": "Frequency of one-off purchases", "PURCHASES_INSTALLMENTS_FREQUENCY": "Frequency of installment purchases", "CASH_ADVANCE_FREQUENCY": "Frequency of cash advances", "CASH_ADVANCE_TRX": "Number of cash-advance transactions", "PURCHASES_TRX": "Number of purchase transactions", "CREDIT_LIMIT": "Assigned credit limit", "PAYMENTS": "Payments made", "MINIMUM_PAYMENTS": "Recorded minimum payment amount", "PRC_FULL_PAYMENT": "Percentage of periods paid in full", "TENURE": "Account tenure in months",
    }
    dictionary_rows = []
    for column in RAW.columns:
        series = RAW[column]
        if column == "CUST_ID":
            example = f"{series.iloc[0]} to {series.iloc[-1]}"
        elif pd.api.types.is_numeric_dtype(series):
            example = f"min {series.min():,.2f}; max {series.max():,.2f}"
        else:
            example = str(series.iloc[0])
        dictionary_rows.append([column, meanings.get(column, "Source field"), str(series.dtype), example])
    table(document, ["Column", "Business meaning", "Data type", "Observed range/example"], dictionary_rows, widths=[1.35, 2.55, .85, 2.0])
    add_takeaway(document, "The data dictionary connects technical column names to the business behavior each field represents.")

    heading(document, "6", "Data Quality Assessment")
    document.add_paragraph("Data quality assessment means checking whether the source is complete, consistent, correctly typed, and suitable for analysis before building charts or conclusions.")
    quality_rows = [["Rows before cleaning", f"{QUALITY['rows_before']:,}", "The starting population."], ["Rows after cleaning", f"{QUALITY['rows_after']:,}", "No rows were lost because there were no exact duplicates."], ["Columns", f"{QUALITY['columns']:,}", "Source fields loaded for analysis."], ["Missing cells before", f"{QUALITY['missing_before']:,}", "These values required documented treatment."], ["Missing cells after", f"{QUALITY['missing_after']:,}", "Zero means all displayed calculations have usable inputs."], ["Duplicate rows", f"{QUALITY['duplicate_count']:,}", "No exact duplicate records were found."], ["Negative numeric values", f"{QUALITY['invalid_numeric_count']:,}", "No invalid negative financial values were found."]]
    table(document, ["Quality check", "Result", "Why it matters"], quality_rows, widths=[1.8, 1.2, 3.7])
    missing_rows = []
    for column in RAW.columns:
        count = int(RAW[column].isna().sum())
        if count or column in ["CREDIT_LIMIT", "MINIMUM_PAYMENTS"]:
            treatment = "Median imputation" if count else "No treatment required"
            missing_rows.append([column, f"{count:,}", pct(count / len(RAW)), treatment])
    table(document, ["Column", "Missing count", "Missing percentage", "Treatment"], missing_rows, widths=[2.1, 1.2, 1.4, 2.0])
    outliers = analytics.iqr_outlier_counts(DATA, ["BALANCE", "PURCHASES", "CASH_ADVANCE", "CREDIT_UTILIZATION", "PAYMENTS"])
    outlier_rows = [[row["Metric"], f"{row['Upper IQR Threshold']:,.2f}", f"{int(row['Flagged Customers']):,}"] for _, row in outliers.iterrows()]
    table(document, ["Metric", "Upper IQR threshold", "Customers flagged"], outlier_rows, widths=[2.2, 2.0, 1.8])
    document.add_paragraph("IQR means interquartile range: the distance between the 25th and 75th percentiles, or the middle half of the values. The upper IQR threshold is a review boundary, not a proof that a record is wrong.")
    add_takeaway(document, "The source is structurally clean, with two missing-value fields requiring transparent median treatment and several valid extreme values requiring contextual review.")

    heading(document, "7", "Data Cleaning & Preprocessing")
    cleaning_rows = [["Step", "What was done", "Why it was done"], ["Grain validation", "Confirmed unique CUST_ID values and one row per customer.", "Prevents accidental double-counting."], ["Type conversion", "Converted all analytical columns to numeric values.", "Makes formulas and summaries reliable."], ["Duplicate handling", "Removed exact duplicate rows if present; observed count was zero.", "Avoids counting one source record twice."], ["Missing values", "Filled one CREDIT_LIMIT and 313 MINIMUM_PAYMENTS values with observed medians.", "Keeps the customer population available for comparable calculations."], ["Invalid values", "Checked for negative numerical values; none were found. Values are clipped at zero defensively.", "Financial amounts and counts should not be negative in this dataset."], ["Extreme values", "Retained valid extremes and flagged them with IQR thresholds.", "Avoids deleting meaningful customer behavior."], ["Raw data", "Kept the supplied CSV unchanged in the raw-data location.", "Preserves reproducibility and auditability."]]
    table(document, cleaning_rows[0], cleaning_rows[1:], widths=[1.25, 3.05, 2.0])
    document.add_paragraph("Imputation means filling a missing value using a documented replacement rule. Median imputation uses the middle observed value; it is less influenced by extreme values than an average. The project documents this assumption rather than presenting imputed values as originally observed.")
    add_takeaway(document, "Preprocessing improves calculation reliability while preserving the raw source and avoiding blind outlier deletion.")

    heading(document, "8", "Exploratory Data Analysis")
    document.add_paragraph("Exploratory data analysis, or EDA, is the process of examining distributions and relationships before forming business interpretations. The report charts below are generated from the actual cleaned project data. Histograms are clipped at the 99th percentile only for visual readability; the underlying analysis retains all values.")
    for figure, key, interpretation in [("Figure 8.1", "credit_limit", "The distribution is concentrated around common credit-limit bands, with a long upper tail showing that a smaller number of customers have much larger limits."), ("Figure 8.2", "balance", "Balances are right-skewed: many customers have relatively modest balances while a smaller group has much higher balances."), ("Figure 8.3", "purchases", "Purchase values are also concentrated at lower levels, with a smaller number of high-purchase observations."), ("Figure 8.4", "utilization", "Most customers show lower utilization, while the upper tail identifies customers whose balances represent a larger share of their limit."), ("Figure 8.5", "cash_advance", "Cash advances are zero or low for many customers, with a visible group showing higher cash-advance activity."), ("Figure 8.6", "full_payment", "Full-payment percentages are concentrated near zero for many customers, indicating that full payment is not common across every record in the snapshot.")]:
        document.add_picture(str(charts[key]), width=Inches(6.45))
        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_caption(document, f"{figure} – {dict([('credit_limit','Distribution of Customer Credit Limits'),('balance','Distribution of Customer Balances'),('purchases','Distribution of Customer Purchases'),('utilization','Distribution of Credit Utilization'),('cash_advance','Distribution of Cash Advance Usage'),('full_payment','Distribution of Full Payment Percentage')])[key]}")
        document.add_paragraph(interpretation)
    document.add_page_break()
    document.add_picture(str(charts["purchases_limit"]), width=Inches(6.45))
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_caption(document, "Figure 8.7 – Purchases versus Credit Limit by Customer Segment")
    document.add_paragraph("The scatter chart compares purchase amounts with credit limits and colors customers by the rule-based segment. The pattern is an observed relationship, not proof that a larger credit limit causes higher purchases.")
    document.add_picture(str(charts["correlation"]), width=Inches(6.45))
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_caption(document, "Figure 8.8 – Correlation Heatmap of Key Financial Measures")
    document.add_paragraph("The heatmap summarizes pairwise association among important measures. It helps identify relationships worth exploring, but correlation alone cannot establish causation.")
    add_takeaway(document, "The EDA shows skewed financial distributions, meaningful variation across customers, and relationships that are useful for descriptive comparison.")

    heading(document, "9", "Feature Engineering")
    document.add_paragraph("Feature engineering means creating new analytical measures from existing columns. The project uses safe division: when a denominator is zero, the derived metric is set to zero rather than producing an error or an infinite value.")
    feature_rows = [["Metric", "Formula", "What it tells the reader"], ["Credit utilization", "BALANCE / CREDIT_LIMIT", "How much of the assigned limit is represented by the balance."], ["Payment ratio", "PAYMENTS / BALANCE", "Payment activity relative to balance; zero balances are handled safely."], ["Purchase-to-limit ratio", "PURCHASES / CREDIT_LIMIT", "Purchases in relation to assigned credit capacity."], ["Cash-advance intensity", "CASH_ADVANCE / CREDIT_LIMIT", "Cash advances in relation to assigned credit capacity."], ["Average purchase value", "PURCHASES / PURCHASES_TRX", "Approximate amount per purchase transaction."], ["Installment purchase share", "INSTALLMENTS_PURCHASES / PURCHASES", "Share of purchases represented by installments."]]
    table(document, feature_rows[0], feature_rows[1:], widths=[1.7, 2.0, 2.6])
    add_takeaway(document, "The derived measures put raw amounts into context, making comparisons across customers and segments more meaningful.")

    heading(document, "10", "Customer Segmentation")
    document.add_paragraph("Segmentation means grouping customers with similar observed characteristics. The existing project uses rule-based business segments, not K-Means or another machine-learning clustering algorithm. Thresholds are calculated from the observed dataset, so the labels describe this portfolio and may change on another dataset.")
    segment_definitions = [["Segment", "Implemented rule", "Business interpretation"]]
    for name, rule, meaning in [("High-Value Active", "PURCHASES and PURCHASES_FREQUENCY are both at or above their dataset 75th percentiles.", "High observed spending combined with frequent purchase activity."), ("High-Utilization", "CREDIT_UTILIZATION is at or above its dataset 75th percentile.", "Balances use a relatively large share of assigned limits."), ("Cash-Advance Heavy", "CASH_ADVANCE_INTENSITY is at or above its dataset 75th percentile.", "Cash advances represent a relatively large share of assigned limits."), ("Installment-Focused", "INSTALLMENT_PURCHASE_SHARE is at least 60% and purchases are positive.", "Installments represent most purchase value."), ("Low-Engagement", "PURCHASES_FREQUENCY is at most 0.25 and PURCHASES_TRX is at or below the median.", "Low observed purchase frequency and transaction activity."), ("Balanced Usage", "Does not meet an earlier segment rule.", "No single specialized behavior dominates the implemented rules.")]:
        segment_definitions.append([name, rule, meaning])
    table(document, segment_definitions[0], segment_definitions[1:], widths=[1.45, 3.0, 1.85])
    segment_rows = [[row["SEGMENT"], f"{int(row['Customers']):,}", pct(row["Customer Share"]), money(row["Purchases"]), money(row["Balance"]), pct(row["Utilization"]), money(row["Cash_Advance"])] for _, row in SEGMENTS.iterrows()]
    table(document, ["Segment", "Customers", "Share", "Avg purchases", "Avg balance", "Avg utilization", "Avg cash advance"], segment_rows, widths=[1.35, .8, .75, 1.1, 1.0, 1.0, 1.25])
    add_takeaway(document, f"{SEGMENTS.iloc[0]['SEGMENT']} is the largest implemented segment with {int(SEGMENTS.iloc[0]['Customers']):,} customers ({pct(SEGMENTS.iloc[0]['Customer Share'])}); segment labels are descriptive groupings, not credit decisions.")

    heading(document, "11", "Financial Behavior Analysis")
    subsections = [
        ("A. Credit Utilization", "Credit utilization is balance divided by credit limit. The average is " + pct(KPIS["Average Credit Utilization"]) + ". A higher value means a larger share of assigned credit is represented by the balance; the measure should be read alongside payment activity and credit limits.", "Relevant visuals: Figure 8.4, Figure 8.7, and the dashboard's Balance vs Credit Limit chart."),
        ("B. Purchase Behavior", "Average purchases are " + money(KPIS["Average Purchases"]) + ". Purchase totals are highly uneven, so the average should be read with the distribution and segment table rather than treated as a typical value for every customer.", "Relevant visuals: Figure 8.3 and the dashboard's Purchases Distribution chart."),
        ("C. Cash Advance Behavior", "Average cash advance is " + money(KPIS["Average Cash Advance"]) + ". Cash-advance intensity compares this amount with credit limit, which makes customers with different limits more comparable.", "Relevant visual: dashboard Cash Advance by Segment chart."),
        ("D. Payment Behavior", "Average payments are " + money(KPIS["Average Payments"]) + ". The payment ratio compares payments with balance and is safely calculated when a balance is zero.", "Relevant visual: dashboard Purchases vs Payments by Segment chart."),
        ("E. Credit Limit", "The average credit limit is " + money(KPIS["Average Credit Limit"]) + ", while the median is " + money(KPIS["Median Credit Limit"]) + ". The difference between average and median shows why both measures are useful when the distribution has a high-value tail.", "Relevant visuals: Figure 8.1 and the dashboard Credit Limit Distribution."),
        ("F. Customer Engagement", "Average purchase frequency is " + pct(KPIS["Average Purchase Frequency"]) + ", while " + pct(DATA["PURCHASES"].eq(0).mean()) + " of customers have zero recorded purchases. This is a descriptive engagement signal for the observation period, not an explanation of customer intent.", "Relevant visual: dashboard Purchases Frequency Distribution."),
        ("G. Installment Purchases", "The installment purchase share identifies customers whose purchase value is primarily installment-based. This supports product and engagement analysis without assuming why a customer selected installments.", "Relevant visual: dashboard Installment Purchases chart."),
        ("H. Full Payment Behavior", "Average full-payment percentage is " + pct(KPIS["Average Full Payment Percentage"]) + ". The distribution shows that full-payment behavior varies substantially across customers.", "Relevant visual: Figure 8.6 and the dashboard Full Payment Percentage Distribution."),
        ("I. Tenure", "Average tenure is " + f"{KPIS['Average Tenure']:.1f} months" + ". Tenure provides account-history context for spending and payment measures, although this snapshot cannot show how behavior changed over time.", "Relevant visual: dashboard Tenure vs Spending chart."),
    ]
    for label, body, visual in subsections:
        document.add_heading(label, level=2)
        document.add_paragraph(body)
        document.add_paragraph(visual)
        add_takeaway(document, "The metric is useful for comparison, but it should be interpreted with the other financial behavior measures rather than in isolation.")

    heading(document, "12", "Key Findings & Insights")
    purchase_corr = DATA[["PURCHASES", "CREDIT_LIMIT"]].corr().iloc[0, 1]
    findings = [
        ("1", f"The portfolio contains {len(DATA):,} unique customers and one row per customer.", "This establishes the correct analytical grain and makes customer counts directly interpretable."),
        ("2", f"{SEGMENTS.iloc[0]['SEGMENT']} is the largest segment with {int(SEGMENTS.iloc[0]['Customers']):,} customers ({pct(SEGMENTS.iloc[0]['Customer Share'])}).", "The largest group is the first audience for understanding the portfolio's most common observed pattern."),
        ("3", f"The average credit limit is {money(KPIS['Average Credit Limit'])}, compared with a median of {money(KPIS['Median Credit Limit'])}.", "Reporting both measures prevents the upper tail from hiding the middle customer experience."),
        ("4", f"Average credit utilization is {pct(KPIS['Average Credit Utilization'])}.", "Utilization provides context for balances by relating them to assigned credit capacity."),
        ("5", f"{pct(DATA['PURCHASES'].eq(0).mean())} of customers have zero recorded purchases.", "The result identifies a sizable low-purchase group for further engagement analysis."),
        ("6", f"The observed purchases-to-credit-limit correlation is {purchase_corr:.2f}.", "The association is descriptive and should not be interpreted as a causal relationship."),
        ("7", f"Average cash advance is {money(KPIS['Average Cash Advance'])}, with cash-advance-heavy customers identified using the portfolio's upper intensity quartile.", "Cash advances deserve a separate behavioral view because they differ from purchase activity."),
        ("8", f"Average payments are {money(KPIS['Average Payments'])}, while average full-payment percentage is {pct(KPIS['Average Full Payment Percentage'])}.", "Payment amount and full-payment frequency capture different aspects of payment behavior."),
        ("9", f"Average tenure is {KPIS['Average Tenure']:.1f} months and the observed range is six to twelve months.", "Tenure gives context but does not provide a time series of behavior changes."),
        ("10", "The source has no exact duplicate rows and no negative numerical values.", "The basic structural checks support using the records after the documented missing-value treatment."),
    ]
    for number, finding, meaning in findings:
        add_insight_box(document, f"Finding {number}", f"{finding} {meaning}")
    add_takeaway(document, "The strongest findings describe concentration, variation, and association in the observed portfolio without overstating what the data can prove.")

    heading(document, "13", "Business Recommendations")
    recommendations = [
        ("Segment concentration", "The largest segment represents the most common observed behavior.", "Use segment profiles to design targeted engagement experiments and compare responses."),
        ("Utilization monitoring", "Utilization is more interpretable when read with balance, limit, and payment measures.", "Monitor these measures together instead of using a single ratio as a decision rule."),
        ("Low purchase activity", "A measurable share of customers has zero purchases in the snapshot.", "Consider activation research and message testing, while investigating whether the snapshot misses relevant context."),
        ("Cash-advance behavior", "Cash-advance-heavy customers have a distinct behavior profile.", "Create a separate monitoring view and review the pattern with business context before action."),
        ("Installment behavior", "Installment-focused customers represent a distinct purchase mix.", "Evaluate installment product communication and customer experience using controlled analysis."),
        ("Static data limitation", "The dataset has no time dimension.", "Add monthly history before claiming that behavior is changing or that an intervention worked."),
    ]
    for finding, interpretation, action in recommendations:
        add_insight_box(document, "Finding → Business interpretation → Recommended action", f"{finding}\n\nBusiness interpretation: {interpretation}\nRecommended action: {action}")
    add_takeaway(document, "Recommendations are practical next steps for analysis and engagement planning, not guaranteed outcomes or personal financial advice.")

    heading(document, "14", "Dashboard Overview")
    document.add_paragraph("The Streamlit dashboard is the interactive application layer of the project. All displayed KPIs, tables, charts, and generated insights are recalculated from the active filtered data.")
    screenshot_info = [("01_executive_overview.png", "Figure 14.1 – Executive Overview Dashboard", "The overview combines the main KPIs with segment distribution, utilization distribution, purchases versus credit limit, balance versus credit limit, and generated key insights."), ("02_spending_analysis.png", "Figure 14.2 – Spending Analysis Dashboard", "This view compares purchase amount, purchase frequency, one-off purchases, installment purchases, tenure versus spending, and cash advance by segment."), ("03_payment_credit.png", "Figure 14.3 – Payment & Credit Analysis", "This view connects purchases with payments, balances with credit limits, full-payment percentage, payment ratio, and the correlation heatmap."), ("04_customer_segments.png", "Figure 14.4 – Customer Segments", "This view compares segment size, average purchases, balance, credit limit, payments, utilization, cash advance, frequency, and tenure."), ("05_customer_explorer.png", "Figure 14.5 – Customer Explorer", "The explorer allows one customer record to be selected and displays only relevant financial behavior metrics, without exposing personal information."), ("06_data_quality.png", "Figure 14.6 – Data Quality Dashboard", "The data-quality view shows rows before and after cleaning, missing values, duplicate counts, cleaning operations, and retained outlier flags." )]
    for index, (filename, caption, explanation) in enumerate(screenshot_info):
        image_path = ROOT / "screenshots" / filename
        if image_path.exists():
            document.add_picture(str(image_path), width=Inches(6.35))
            document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_caption(document, caption)
            document.add_paragraph(explanation)
        else:
            document.add_paragraph(f"{caption}: screenshot unavailable.")
        if index in [1, 3]:
            document.add_page_break()
    add_takeaway(document, "The dashboard links portfolio-level analysis to filtered segment views and customer-level exploration through one consistent application.")

    heading(document, "15", "Technical Architecture")
    document.add_paragraph("The project is Python-based from data loading through dashboard rendering. The main source file is intentionally self-contained so another user can run the analytical application after installing requirements.txt and placing the raw CSV in the documented location.")
    document.add_picture(str(charts["architecture"]), width=Inches(6.55))
    document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_caption(document, "Figure 15.1 – Technical architecture and analytical flow")
    architecture_rows = [["File or component", "Role"], ["NEERAJGIRI_CreditCardCustomerFinancialAnalytics.py", "Loads data, cleans it, engineers features, calculates KPIs, assigns segments, creates charts, and runs the Streamlit dashboard."], ["CC GENERAL.csv", "Original manually supplied Kaggle dataset; kept unchanged as the raw input."], ["data/processed/cleaned_credit_card_customers.csv", "Generated cleaned and feature-enriched data output."], ["requirements.txt", "Lists the Python libraries needed to run the project."], ["README.md", "Documents setup, functionality, project structure, and limitations."], ["screenshots/", "Stores actual rendered dashboard screenshots used in the portfolio and report."]]
    table(document, architecture_rows[0], architecture_rows[1:], widths=[2.8, 3.7])
    add_takeaway(document, "The architecture keeps the analytical logic visible in one main Python file while separating raw data, processed output, documentation, and screenshots.")

    heading(document, "16", "Project Workflow")
    document.add_paragraph("The workflow below shows how the project moves from a raw file to business-facing insight. Each stage has a distinct purpose and produces evidence for the next stage.")
    workflow_rows = [["Stage", "Output"], ["Dataset", "Verified source CSV and one-row-per-customer understanding."], ["Data validation", "Column, type, missing-value, duplicate, and range checks."], ["Cleaning", "Complete analytical table with documented imputation."], ["Feature engineering", "Contextual ratios and behavior measures."], ["EDA", "Distributions, relationships, and correlation views."], ["Segmentation", "Six transparent rule-based behavior groups."], ["Visualization", "Readable Plotly, Matplotlib, and Seaborn charts."], ["Dashboard", "Interactive filters, tabs, tables, and customer explorer."], ["Business insights", "Evidence-based findings and possible actions."]]
    table(document, workflow_rows[0], workflow_rows[1:], widths=[1.7, 4.8])
    add_takeaway(document, "The workflow creates a traceable path from data to information, information to insight, and insight to possible business action.")

    heading(document, "17", "Limitations")
    for item in ["The dataset is a historical/static snapshot without a time series.", "Customer behavior may change after the observation period.", "Correlation does not establish causation.", "The dataset does not include demographic, campaign, or contextual variables that could explain behavior.", "The analysis uses only the variables supplied in CC GENERAL.csv.", "Rule-based segment thresholds are relative to this dataset and may not transfer unchanged to another portfolio.", "Median imputation makes the table complete but adds an assumption for missing values.", "The dashboard is descriptive and does not provide prediction, credit approval, fraud detection, or personal financial advice."]:
        add_bullet(document, item)
    add_takeaway(document, "The results are suitable for descriptive portfolio analysis, while stronger operational decisions require richer, time-based, and validated business data.")

    heading(document, "18", "Future Scope")
    for item in ["Automated data pipelines and scheduled data-quality checks.", "Real-time or near-real-time dashboard refreshes.", "Validated thresholds based on business expertise and monitoring.", "Optional advanced customer clustering for comparison with the transparent rules.", "Predictive modeling for churn or other clearly labeled future outcomes.", "Credit-risk modeling only with appropriate labels, governance, and validation.", "Anomaly detection for unusual changes over time.", "Recommendation systems for controlled engagement experiments.", "Database integration and cloud deployment with access controls."]:
        add_bullet(document, item)
    add_takeaway(document, "Future improvements should add context and validation without losing the transparency of the current analytical workflow.")

    heading(document, "19", "Conclusion")
    document.add_paragraph("This project demonstrates the ability to transform raw financial customer data into a professional analytical product. The source was inspected, cleaned transparently, enriched with meaningful ratios, compared through rule-based segments, visualized in multiple ways, and presented in a responsive Streamlit dashboard.")
    document.add_paragraph("The result is a descriptive decision-support tool: Data becomes organized information; information becomes evidence-based insight; and insight suggests business questions and possible actions. The project does not overstate what the dataset can prove, which is an important part of responsible analytics.")
    add_takeaway(document, "The project connects technical Python work with clear business communication and a reproducible portfolio deliverable.")

    heading(document, "20", "References")
    references = ["Kaggle. Credit Card Dataset for Clustering. https://www.kaggle.com/datasets/arjunbhasin2013/ccdata", "Python Software Foundation. Python Documentation. https://docs.python.org/3/", "The pandas development team. pandas documentation. https://pandas.pydata.org/docs/", "NumPy Developers. NumPy documentation. https://numpy.org/doc/", "Matplotlib Development Team. Matplotlib documentation. https://matplotlib.org/stable/", "Waskom, M. Seaborn documentation. https://seaborn.pydata.org/", "Plotly Technologies Inc. Plotly Python documentation. https://plotly.com/python/", "Streamlit Inc. Streamlit documentation. https://docs.streamlit.io/", "python-docx documentation. https://python-docx.readthedocs.io/"]
    for reference in references:
        add_bullet(document, reference)
    add_takeaway(document, "The report uses the supplied dataset and the documented Python analytics stack as its source material.")

    # Update fields on open in Word.
    settings = document.settings.element
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.append(update_fields)
    document.save(OUTPUT)
    print(f"Created {OUTPUT}")


if __name__ == "__main__":
    build_report()
