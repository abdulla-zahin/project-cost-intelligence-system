import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import date
from io import BytesIO
from pathlib import Path
import textwrap

st.set_page_config(
    page_title="Project Cost Intelligence System",
    page_icon="🏗️",
    layout="wide",
)

PROJECT_TYPE_SPLITS = {
    "General Contracting": {
        "Materials": 40,
        "Labor": 25,
        "Transportation": 5,
        "Office Expense": 5,
        "Salaries / Overheads": 10,
        "Company Profit": 10,
        "Contingency": 5,
    },
    "Pipeline Project": {
        "Materials": 50,
        "Labor": 20,
        "Transportation": 10,
        "Office Expense": 5,
        "Salaries / Overheads": 5,
        "Company Profit": 7,
        "Contingency": 3,
    },
    "Civil Construction": {
        "Materials": 45,
        "Labor": 30,
        "Transportation": 5,
        "Office Expense": 5,
        "Salaries / Overheads": 8,
        "Company Profit": 5,
        "Contingency": 2,
    },
    "Mechanical Installation": {
        "Materials": 38,
        "Labor": 32,
        "Transportation": 6,
        "Office Expense": 5,
        "Salaries / Overheads": 9,
        "Company Profit": 7,
        "Contingency": 3,
    },
    "Maintenance Contract": {
        "Materials": 25,
        "Labor": 40,
        "Transportation": 8,
        "Office Expense": 7,
        "Salaries / Overheads": 10,
        "Company Profit": 7,
        "Contingency": 3,
    },
    "EPC Project": {
        "Materials": 42,
        "Labor": 22,
        "Transportation": 6,
        "Office Expense": 6,
        "Salaries / Overheads": 12,
        "Company Profit": 8,
        "Contingency": 4,
    },
}

CURRENCY_SYMBOLS = {
    "AED": "AED",
    "USD": "$",
    "INR": "₹",
}


def build_budget_dataframe(total_budget: float, allocations: dict) -> pd.DataFrame:
    rows = []
    for category, pct in allocations.items():
        rows.append(
            {
                "Category": category,
                "Allocation %": round(float(pct), 2),
                "Amount": round(total_budget * float(pct) / 100, 2),
            }
        )
    return pd.DataFrame(rows)



def get_allocation_status(total_pct: float) -> str:
    if abs(total_pct - 100) < 1e-9:
        return "Balanced"
    if total_pct < 100:
        return "Under Allocated"
    return "Over Allocated"



def analyze_budget(allocations: dict, total_pct: float) -> list[str]:
    insights = []
    status = get_allocation_status(total_pct)

    if status == "Balanced":
        insights.append("✅ Total allocation is perfectly balanced at 100%.")
    elif status == "Under Allocated":
        insights.append("⚠️ Total allocation is below 100%. Some budget is still unassigned.")
    else:
        insights.append("❌ Total allocation exceeds 100%. Adjust the percentages to match the budget.")

    profit = allocations["Company Profit"]
    labor = allocations["Labor"]
    materials = allocations["Materials"]
    contingency = allocations["Contingency"]
    transport = allocations["Transportation"]

    if profit < 8:
        insights.append("⚠️ Profit margin is below 8%. This may be too tight for comfortable execution.")
    elif profit <= 15:
        insights.append("✅ Profit margin looks healthy for a practical contracting estimate.")
    else:
        insights.append("⚠️ Profit margin is quite high. Double-check competitiveness and client expectations.")

    if labor < 20:
        insights.append("⚠️ Labor allocation is low. Execution quality or manpower coverage may suffer.")
    elif labor > 35:
        insights.append("⚠️ Labor allocation is very high. Verify productivity assumptions and manpower planning.")

    if materials < 30:
        insights.append("⚠️ Material allocation is low. Review BOQ realism and procurement assumptions.")
    elif materials > 60:
        insights.append("⚠️ Material allocation is very high. Check whether supply cost is dominating the project.")

    if contingency < 5:
        insights.append("⚠️ Contingency is low. This leaves little room for site surprises or change orders.")
    else:
        insights.append("✅ Contingency reserve provides a reasonable safety buffer.")

    if transport > 10:
        insights.append("⚠️ Transportation cost is high. Recheck logistics, fuel, and delivery assumptions.")

    return insights



def generate_report_text(
    project_name: str,
    project_type: str,
    company_name: str,
    client_name: str,
    project_reference: str,
    currency_symbol: str,
    total_budget: float,
    allocations: dict,
    total_pct: float,
) -> dict:
    profit_pct = allocations["Company Profit"]
    profit_amount = total_budget * profit_pct / 100
    allocation_status = get_allocation_status(total_pct).lower()

    estimate_summary = (
        f"Project Reference: {project_reference}. Client: {client_name}. "
        f"Based on the submitted project details for {project_name}, categorized under {project_type}, "
        f"{company_name} has prepared a preliminary budget estimate with a total projected value of {currency_symbol} {total_budget:,.2f}. "
        f"The estimate has been distributed across key execution heads including materials, labor, transportation, office expenses, overheads, contingency, and company profit. "
        f"The current model reflects an expected profit margin of {profit_pct:.2f}% with an estimated profit value of {currency_symbol} {profit_amount:,.2f}. "
        f"Overall allocation status is presently {allocation_status}, subject to final technical and commercial review."
    )

    proposal_note = (
        f"Reference {project_reference} has been prepared for {client_name}. "
        f"We are pleased to submit this preliminary commercial budget consideration for {project_name}. "
        f"This estimate has been developed based on the selected project type, expected execution requirements, material demand, labor deployment, logistics considerations, and administrative overheads. "
        f"The proposed allocation is intended to maintain practical project execution while preserving commercial viability for {company_name}. "
        f"Any final quotation or commercial submission should be validated against approved scope, drawings, specifications, and prevailing market rates before issue."
    )

    disclaimer_note = (
        "This budget estimate is prepared solely for planning, review, and preliminary commercial evaluation purposes. "
        "Final values may vary depending on site conditions, actual quantities, client requirements, specification changes, transportation conditions, authority approvals, and material price fluctuations at the time of execution. "
        "Any variation in scope, timeline, procurement conditions, or execution methodology may lead to revision of the final commercial offer and agreement terms."
    )

    internal_note = (
        f"Internal review reference: {project_reference}. "
        f"From an internal review standpoint, the current allocation for {project_name} appears {allocation_status} with a projected profit margin of {profit_pct:.2f}%. "
        "Management is advised to review category distribution, execution risks, and contingency adequacy before final approval or client submission."
    )

    return {
        "Estimate Summary": estimate_summary,
        "Commercial Proposal Note": proposal_note,
        "Terms / Disclaimer Note": disclaimer_note,
        "Internal Approval Note": internal_note,
    }



from dataclasses import dataclass

@dataclass
class ProjectReportParams:
    project_name: str
    project_type: str
    company_name: str
    client_name: str
    project_reference: str
    currency_code: str
    currency_symbol: str
    author_name: str
    today: str
    total_budget: float
    allocations: dict
    df: pd.DataFrame
    insights: list[str]
    report_sections: dict

def create_pdf_report(params: ProjectReportParams) -> BytesIO:
    """
    Generate a comprehensive PDF report with budget breakdown and analysis.

    Returns:
        BytesIO: A BytesIO object containing the PDF data.
    """
    pdf_buffer = BytesIO()
    profit_pct = params.allocations.get("Company Profit", 0)
    pdf_buffer = BytesIO()
    # Ensure Amount column is numeric for plotting
    numeric_amounts = pd.to_numeric(params.df["Amount"], errors="coerce")
    profit_pct = params.allocations.get("Company Profit", 0)
    profit_amount = params.total_budget * profit_pct / 100
    total_pct = sum(params.allocations.values())
    allocation_status = get_allocation_status(total_pct)

    # Page 1
    fig1 = plt.figure(figsize=(8.27, 11.69))
    fig1.patch.set_facecolor("white")
    ax1 = fig1.add_axes([0.05, 0.05, 0.9, 0.9])
    ax1.axis("off")
    y = 0.96
    ax1.text(0.5, y, params.company_name, ha="center", va="top", fontsize=20, fontweight="bold")
    y -= 0.035
    ax1.text(0.5, y, "PROJECT COST INTELLIGENCE REPORT", ha="center", va="top", fontsize=14, fontweight="bold")
    y -= 0.06

    left_lines = [
        f"Project Name: {params.project_name}",
        f"Client Name: {params.client_name}",
        f"Project Reference: {params.project_reference}",
        f"Project Type: {params.project_type}",
        f"Date: {params.today}",
        f"Prepared By: {params.author_name}",
        f"Currency: {params.currency_code}",
    ]
    right_lines = [
        f"Total Budget: {params.currency_symbol} {params.total_budget:,.2f}",
        f"Profit Margin: {params.allocations['Company Profit']:.2f}%",
        f"Profit Amount: {params.currency_symbol} {profit_amount:,.2f}",
        f"Allocation Status: {allocation_status}",
    ]

    y_left = y
    for line in left_lines:
        ax1.text(0.08, y_left, line, fontsize=10, va="top")
        y_left -= 0.03

    y_right = y
    for line in right_lines:
        ax1.text(0.65, y_right, line, fontsize=10, va="top")
        y_right -= 0.03

    display_df = params.df.copy()
    display_df["Allocation %"] = display_df["Allocation %"].map(lambda x: f"{x:.2f}%")
    display_df["Amount"] = display_df["Amount"].map(lambda x: f"{params.currency_symbol} {x:,.2f}")
    y = min(y_left, y_right) - 0.03
    ax1.text(0.08, y, "Budget Intelligence", fontsize=12, fontweight="bold", va="top")
    y -= 0.03

    for insight in params.insights:
        wrapped = textwrap.fill(insight, width=90)
        ax1.text(0.09, y, wrapped, fontsize=9, va="top")
        y -= 0.04 + 0.012 * wrapped.count("\n")

    pdf = PdfPages(pdf_buffer)
    pdf.savefig(fig1, bbox_inches="tight")
    plt.close(fig1)

    # Page 2
    fig2 = plt.figure(figsize=(8.27, 11.69))
    fig2.patch.set_facecolor("white")
    ax2 = fig2.add_axes([0.05, 0.05, 0.9, 0.9])
    ax2.axis("off")
    ax2.set_title("Detailed Budget Breakdown", fontsize=14, fontweight="bold", pad=20)

    display_df2 = params.df.copy()
    display_df2["Allocation %"] = display_df2["Allocation %"].map(lambda x: f"{x:.2f}%")
    display_df2["Amount"] = display_df2["Amount"].map(lambda x: f"{params.currency_symbol} {x:,.2f}")

    table = ax2.table(
        cellText=display_df2.values,
        colLabels=display_df2.columns,
        cellLoc="center",
        colLoc="center",
        loc="upper center",
    )

    # Highlight rows for Company Profit and Contingency
    for row in range(1, len(display_df2) + 1):
        category = display_df2.iloc[row - 1, 0]
        for col in range(len(display_df2.columns)):
            cell = table[row, col]
            if category == "Company Profit":
                cell.set_facecolor("#FFF2CC")
            elif category == "Contingency":
                cell.set_facecolor("#E2F0D9")

    pdf.savefig(fig2, bbox_inches="tight")
    plt.close(fig2)

    # Page 3
    fig3 = plt.figure(figsize=(8.27, 11.69))
    fig3.patch.set_facecolor("white")
    # Pie chart
    ax3 = fig3.add_axes([0.08, 0.53, 0.84, 0.36])
    ax3.pie(numeric_amounts, labels=params.df["Category"], autopct="%1.1f%%", startangle=90)
    ax3.set_title("Budget Distribution")
    # Bar chart
    ax4 = fig3.add_axes([0.08, 0.09, 0.84, 0.32])
    ax4.bar(params.df["Category"], numeric_amounts)
    ax4.set_title(f"Category-wise Budget Amount ({params.currency_code})")
    ax4.set_ylabel(params.currency_code)
    ax4.tick_params(axis="x", rotation=45)

    pdf.savefig(fig3, bbox_inches="tight")
    plt.close(fig3)

    # Page 4+
    fig4 = plt.figure(figsize=(8.27, 11.69))
    fig4.patch.set_facecolor("white")
    ax5 = fig4.add_axes([0, 0, 1, 1])
    ax5.axis("off")
    ax5.text(0.5, 0.96, "AUTO-GENERATED REPORT DRAFT", ha="center", va="top", fontsize=14, fontweight="bold")
    y = 0.9

    for title, content in params.report_sections.items():
        ax5.text(0.07, y, title, fontsize=11, fontweight="bold", va="top")
        y -= 0.03
        wrapped = textwrap.fill(content, width=100)
        ax5.text(0.08, y, wrapped, fontsize=9, va="top")
        y -= 0.08 + 0.014 * wrapped.count("\n")

        if y < 0.12:
            pdf.savefig(fig4, bbox_inches="tight")
            plt.close(fig4)
            fig4 = plt.figure(figsize=(8.27, 11.69))
            fig4.patch.set_facecolor("white")
            ax5 = fig4.add_axes([0, 0, 1, 1])
            ax5.axis("off")
            y = 0.95

    pdf.savefig(fig4, bbox_inches="tight")
    plt.close(fig4)

    pdf_buffer.seek(0)
    return pdf_buffer
        ws.sheet_view.showGridLines = False
        for col, width in {"A": 16, "B": 24, "C": 24, "D": 24, "E": 24, "F": 24}.items():
            ws.column_dimensions[col].width = width

        ws["A1"].fill = dark_fill
        ws["A1"].border = border
        ws["A1"].alignment = center
        ws.merge_cells("B1:F1")
        ws["B1"] = f"{company_name} - PROJECT COST INTELLIGENCE REPORT"
        ws["B1"].fill = dark_fill
        ws["B1"].font = title_font
        ws["B1"].alignment = center
        ws.row_dimensions[1].height = 34

        if logo_path and Path(logo_path).exists():
            try:
                logo = XLImage(logo_path)
                logo.width = 72
                logo.height = 72
                ws.add_image(logo, "A1")
            except Exception:
                ws["A1"] = "LOGO"
                ws["A1"].font = header_font
                ws["A1"].alignment = center
        else:
            ws["A1"] = "LOGO"
            ws["A1"].font = header_font
            ws["A1"].alignment = center

        ws.merge_cells("A3:C3")
        ws["A3"] = "PROJECT INFORMATION"
        ws["A3"].fill = mid_fill
        ws["A3"].font = section_font
        ws["A3"].alignment = center

        ws.merge_cells("D3:F3")
        ws["D3"] = "KEY FINANCIAL METRICS"
        ws["D3"].fill = gold_fill
        ws["D3"].font = section_font
        ws["D3"].alignment = center

        info_rows = [
            ("Project Name", project_name),
            ("Project Reference", project_reference),
            ("Client Name", client_name),
            ("Project Type", project_type),
            ("Company Name", company_name),
            ("Author", author_name),
            ("Date", today),
            ("Currency", currency_code),
        ]
        metric_rows = [
            ("Total Budget", total_budget),
            ("Profit %", profit_pct / 100),
            ("Profit Amount", profit_amount),
            ("Allocation %", total_allocation_pct / 100),
            ("Status", allocation_status),
        ]

        start_row = 5
        for i, (label, value) in enumerate(info_rows, start=start_row):
            ws[f"A{i}"] = label
            ws[f"B{i}"] = value
            ws[f"A{i}"].fill = soft_fill
            ws[f"A{i}"].font = header_font
            ws[f"A{i}"].border = border
            ws[f"B{i}"].border = border
            ws[f"C{i}"].border = border
            ws[f"A{i}"].alignment = left
            ws[f"B{i}"].alignment = left

        for i, (label, value) in enumerate(metric_rows, start=start_row):
            ws[f"D{i}"] = label
            ws[f"E{i}"] = value
            ws[f"F{i}"] = ""
            ws[f"D{i}"].fill = gray_fill
            ws[f"D{i}"].font = header_font
            ws[f"D{i}"].border = border
            ws[f"E{i}"].border = border
            ws[f"F{i}"].border = border
            ws[f"D{i}"].alignment = left
            ws[f"E{i}"].alignment = center if label != "Status" else left

            if label in {"Total Budget", "Profit Amount"}:
                ws[f"E{i}"].number_format = f'"{currency_symbol}" #,##0.00'
                ws[f"E{i}"].font = big_font
            elif label in {"Profit %", "Allocation %"}:
                ws[f"E{i}"].number_format = "0.00%"
                ws[f"E{i}"].font = big_font
            elif label == "Status":
                ws[f"E{i}"].fill = green_fill if allocation_status == "Balanced" else red_fill
                ws[f"E{i}"].font = header_font

        note_header_row = start_row + max(len(info_rows), len(metric_rows)) + 2
        note_body_row = note_header_row + 1
        ws.merge_cells(f"A{note_header_row}:F{note_header_row}")
        ws[f"A{note_header_row}"] = "EXECUTIVE NOTE"
        ws[f"A{note_header_row}"].fill = mid_fill
        ws[f"A{note_header_row}"].font = section_font
        ws[f"A{note_header_row}"].alignment = center

        note_text = (
            f"This report summarizes the recommended allocation for a {project_type.lower()} under {company_name} for client {client_name}. "
            f"Current allocation status is {allocation_status.lower()}, with an estimated profit of {profit_pct:.2f}% and profit amount of {currency_symbol} {profit_amount:,.2f}."
        )
        ws.merge_cells(f"A{note_body_row}:F{note_body_row + 2}")
        ws[f"A{note_body_row}"] = note_text
        ws[f"A{note_body_row}"].alignment = wrap_left
        ws[f"A{note_body_row}"].border = border

        bd = wb["Budget Breakdown"]
        bd.sheet_view.showGridLines = False
        for col, width in {"A": 30, "B": 16, "C": 22, "D": 18}.items():
            bd.column_dimensions[col].width = width

        bd.merge_cells("A1:D1")
        bd["A1"] = "DETAILED BUDGET BREAKDOWN"
        bd["A1"].fill = dark_fill
        bd["A1"].font = title_font
        bd["A1"].alignment = center

        bd["A3"] = "Project Name"
        bd["B3"] = project_name
        bd["A4"] = "Project Reference"
        bd["B4"] = project_reference
        bd["A5"] = "Client Name"
        bd["B5"] = client_name
        bd["A6"] = "Project Type"
        bd["B6"] = project_type
        bd["D3"] = "Prepared By"
        bd["D4"] = author_name
        bd["D5"] = "Prepared On"
        bd["D6"] = today

        for cell in ["A3", "A4", "A5", "A6", "D3", "D5"]:
            bd[cell].fill = soft_fill
            bd[cell].font = header_font
        for cell in ["A3", "B3", "A4", "B4", "A5", "B5", "A6", "B6", "D3", "D4", "D5", "D6"]:
            bd[cell].border = border

        table_header_row = 8
        bd["D8"] = "Remarks"
        for cell in bd[table_header_row]:
            cell.fill = mid_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border

        for row in range(table_header_row + 1, table_header_row + 1 + len(df)):
            bd[f"B{row}"].number_format = "0.00"
            bd[f"C{row}"].number_format = f'"{currency_symbol}" #,##0.00'
            bd[f"D{row}"] = "Reviewed"
            for col in "ABCD":
                bd[f"{col}{row}"].border = border
            category = bd[f"A{row}"].value
            if category == "Company Profit":
                for col in "ABCD":
                    bd[f"{col}{row}"].fill = gold_fill
                    bd[f"{col}{row}"].font = header_font
            elif category == "Contingency":
                for col in "ABCD":
                    bd[f"{col}{row}"].fill = green_fill

        total_row = table_header_row + 1 + len(df)
        bd[f"A{total_row}"] = "Total Allocation"
        bd[f"B{total_row}"] = total_allocation_pct
        bd[f"C{total_row}"] = df["Amount"].sum()
        bd[f"D{total_row}"] = allocation_status
        for col in "ABCD":
            bd[f"{col}{total_row}"].fill = gray_fill
            bd[f"{col}{total_row}"].font = header_font
            bd[f"{col}{total_row}"].border = border
        bd[f"B{total_row}"].number_format = "0.00"
        bd[f"C{total_row}"].number_format = f'"{currency_symbol}" #,##0.00'
        bd.freeze_panes = "A8"

        ins = wb["Budget Insights"]
        ins.sheet_view.showGridLines = False
        ins.column_dimensions["A"].width = 110
        ins["A1"] = "BUDGET INTELLIGENCE & RISK NOTES"
        ins["A1"].fill = dark_fill
        ins["A1"].font = title_font
        ins["A1"].alignment = center
        ins["A4"].fill = mid_fill
        ins["A4"].font = header_font
        ins["A4"].border = border
        for row in range(5, 5 + len(insights_with_status)):
            ins[f"A{row}"].border = border
            ins[f"A{row}"].alignment = wrap_left
            value = str(ins[f"A{row}"].value)
            if "✅" in value or "Balanced" in value:
                ins[f"A{row}"].fill = green_fill
            elif "⚠️" in value or "❌" in value or "Over" in value or "Under" in value:
                ins[f"A{row}"].fill = red_fill
        ins.freeze_panes = "A4"

        rpt = wb.create_sheet("Report Draft")
        rpt.sheet_view.showGridLines = False
        rpt.column_dimensions["A"].width = 28
        rpt.column_dimensions["B"].width = 110
        rpt.merge_cells("A1:B1")
        rpt["A1"] = "AUTO-GENERATED REPORT DRAFT"
        rpt["A1"].fill = dark_fill
        rpt["A1"].font = title_font
        rpt["A1"].alignment = center
        report_row = 3
        for title, content in report_sections.items():
            rpt[f"A{report_row}"] = title
            rpt[f"A{report_row}"].fill = mid_fill
            rpt[f"A{report_row}"].font = header_font
            rpt[f"A{report_row}"].border = border
            rpt[f"B{report_row}"] = content
            rpt[f"B{report_row}"].alignment = wrap_left
            rpt[f"B{report_row}"].border = border
            rpt.row_dimensions[report_row].height = 54
            report_row += 2

        ch = wb.create_sheet("Charts")
        ch.sheet_view.showGridLines = False
        ch.merge_cells("A1:H1")
        ch["A1"] = "BUDGET VISUALS"
        ch["A1"].fill = dark_fill
        ch["A1"].font = title_font
        ch["A1"].alignment = center
        ch["A2"] = "Category"
        ch["B2"] = "Amount"
        ch["C2"] = "Allocation %"
        for cell in ch[2]:
            cell.fill = mid_fill
            cell.font = header_font
            cell.border = border

        for idx, (_, row) in enumerate(df.iterrows(), start=3):
            ch[f"A{idx}"] = row["Category"]
            ch[f"B{idx}"] = row["Amount"]
            ch[f"C{idx}"] = row["Allocation %"] / 100
            ch[f"B{idx}"].number_format = f'"{currency_symbol}" #,##0.00'
            ch[f"C{idx}"].number_format = "0.00%"

        pie = PieChart()
        pie_labels = Reference(ch, min_col=1, min_row=3, max_row=2 + len(df))
        pie_data = Reference(ch, min_col=2, min_row=2, max_row=2 + len(df))
        pie.add_data(pie_data, titles_from_data=True)
        pie.set_categories(pie_labels)
        pie.title = "Budget Distribution"
        pie.height = 9
        pie.width = 12
        ch.add_chart(pie, "E3")

        bar = BarChart()
        bar.type = "col"
        bar.style = 10
        bar.title = "Allocation by Category"
        bar.y_axis.title = f"Amount ({currency_code})"
        bar.x_axis.title = "Category"
        bar_data = Reference(ch, min_col=2, min_row=2, max_row=2 + len(df))
        bar_labels = Reference(ch, min_col=1, min_row=3, max_row=2 + len(df))
        bar.add_data(bar_data, titles_from_data=True)
        bar.set_categories(bar_labels)
        bar.height = 9
        bar.width = 14
        ch.add_chart(bar, "E20")

    output.seek(0)
    return output


st.title("🏗️ Project Cost Intelligence System")
st.caption("Smart budget allocation, profit analysis, and planning dashboard for engineering / contracting projects.")

st.subheader("Project Details")
col1, col2 = st.columns(2)

with col1:
    project_name = st.text_input("Project Name", value="New Engineering Project")
    client_name = st.text_input("Client Name", value="Client Name")
    company_name = st.text_input("Company Name", value="TKM")
    logo_path_input = st.text_input("Company Logo Path (optional)", value="")

with col2:
    author_name = "Abdulla Zahin"
    today = date.today().strftime("%d %B %Y")
    project_reference = st.text_input("Project Reference", value="TKM-EST-2026-001")
    currency_code = st.selectbox("Currency", options=["AED", "USD", "INR"], index=0)
    st.text_input("Author", value=author_name, disabled=True)
    st.text_input("Date", value=today, disabled=True)

currency_symbol = CURRENCY_SYMBOLS[currency_code]

st.subheader("Budget Input")
project_type = st.selectbox("Project Type", options=list(PROJECT_TYPE_SPLITS.keys()), index=0)

if "last_project_type" not in st.session_state:
    st.session_state.last_project_type = project_type
if "allocations_state" not in st.session_state:
    st.session_state.allocations_state = PROJECT_TYPE_SPLITS[project_type].copy()
if project_type != st.session_state.last_project_type:
    st.session_state.allocations_state = PROJECT_TYPE_SPLITS[project_type].copy()
    st.session_state.last_project_type = project_type

selected_split = st.session_state.allocations_state

total_budget = st.number_input(
    "Total Project Budget",
    min_value=0.0,
    value=500000.0,
    step=10000.0,
    format="%.2f",
)

st.markdown("### Recommended Allocation")
st.write(f"Smart recommendation is loaded for **{project_type}**. You can still adjust the percentages if needed.")

alloc_col1, alloc_col2 = st.columns(2)
allocations = {}
items = list(selected_split.items())
mid = len(items) // 2 + len(items) % 2

with alloc_col1:
    for category, value in items[:mid]:
        widget_key = f"alloc_{category}"
        allocations[category] = st.number_input(
            f"{category} (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(st.session_state.allocations_state[category]),
            step=1.0,
            key=widget_key,
        )
        st.session_state.allocations_state[category] = allocations[category]

with alloc_col2:
    for category, value in items[mid:]:
        widget_key = f"alloc_{category}"
        allocations[category] = st.number_input(
            f"{category} (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(st.session_state.allocations_state[category]),
            step=1.0,
            key=widget_key,
        )
        st.session_state.allocations_state[category] = allocations[category]
# Validate allocations to ensure all values are numeric
total_pct = sum(float(v) for v in allocations.values() if isinstance(v, (int, float)) or (isinstance(v, str) and v.replace('.', '', 1).isdigit()))
remaining_pct = 100 - total_pct
remaining_pct = 100 - total_pct

df = build_budget_dataframe(total_budget, allocations)
insights = analyze_budget(allocations, total_pct)
report_sections = generate_report_text(
    project_name=project_name,
    project_type=project_type,
    company_name=company_name,
    client_name=client_name,
    project_reference=project_reference,
    currency_symbol=currency_symbol,
    total_budget=total_budget,
    allocations=allocations,
    total_pct=total_pct,
)

info1, info2, info3, info4 = st.columns(4)
info1.metric("Total Allocation %", f"{total_pct:.2f}%")
info2.metric("Unallocated / Excess %", f"{remaining_pct:.2f}%")
info3.metric("Estimated Profit %", f"{allocations['Company Profit']:.2f}%")
info4.metric("Estimated Profit Amount", f"{currency_symbol} {total_budget * allocations['Company Profit'] / 100:,.2f}")

st.subheader("Budget Breakdown")
st.dataframe(df, use_container_width=True, hide_index=True)

st.subheader("Budget Intelligence")
for item in insights:
    st.write(item)

st.subheader("Project Summary")
summary1, summary2, summary3, summary4 = st.columns(4)
profit_amount = total_budget * allocations["Company Profit"] / 100
materials_amount = total_budget * allocations["Materials"] / 100
labor_amount = total_budget * allocations["Labor"] / 100
contingency_amount = total_budget * allocations["Contingency"] / 100
summary1.metric("Total Budget", f"{currency_symbol} {total_budget:,.2f}")
summary2.metric("Materials", f"{currency_symbol} {materials_amount:,.2f}")
summary3.metric("Labor", f"{currency_symbol} {labor_amount:,.2f}")
summary4.metric("Contingency", f"{currency_symbol} {contingency_amount:,.2f}")
st.metric("Company Profit", f"{currency_symbol} {profit_amount:,.2f}", f"{allocations['Company Profit']:.2f}% margin")

st.subheader("Visual Dashboard")
chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    fig1, ax1 = plt.subplots(figsize=(6, 6))
    ax1.pie(df["Amount"], labels=df["Category"], autopct="%1.1f%%", startangle=90)
    ax1.set_title("Budget Distribution")
    st.pyplot(fig1)
    plt.close(fig1)
with chart_col2:
    fig2, ax2 = plt.subplots(figsize=(8, 5))
    ax2.bar(df["Category"], df["Amount"])
    ax2.set_title("Category-wise Budget Amount")
    ax2.set_xlabel("Category")
    ax2.set_ylabel(f"Amount ({currency_code})")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig2)
    plt.close(fig2)

st.subheader("Auto Report Writing")
report_option = st.selectbox("Select Report Draft Type", options=list(report_sections.keys()))
st.markdown("### Generated Draft")
st.text_area("Report Text", value=report_sections[report_option], height=220)
full_report_text = "\n\n".join([f"{title}\n{content}" for title, content in report_sections.items()])
st.download_button(
    label="📝 Download Report Draft (.txt)",
    data=full_report_text.encode("utf-8"),
    file_name="Project_Budget_Report.txt",
    mime="text/plain",
)

pdf_params = ProjectReportParams(
    project_name=project_name,
    project_type=project_type,
    company_name=company_name,
    client_name=client_name,
    project_reference=project_reference,
    currency_code=currency_code,
    currency_symbol=currency_symbol,
    author_name=author_name,
    today=today,
    total_budget=total_budget,
    allocations=allocations,
    df=df,
    insights=insights,
    report_sections=report_sections,
)
pdf_file = create_pdf_report(pdf_params)

st.markdown("---")
st.markdown("### Report Metadata")
st.write(f"**Project Name:** {project_name}")
st.write(f"**Project Type:** {project_type}")
st.write(f"**Project Reference:** {project_reference}")
st.write(f"**Client Name:** {client_name}")
st.write(f"**Company Name:** {company_name}")
st.write(f"**Currency:** {currency_code}")
st.write(f"**Author:** {author_name}")
st.write(f"**Date:** {today}")
st.write(f"**Logo Path:** {logo_path_input if logo_path_input else 'Not provided'}")

export_col1, export_col2 = st.columns(2)
with export_col1:
    st.download_button(
        label="📄 Download PDF Report",
        data=pdf_file,
        file_name="Project_Budget_Report.pdf",
        mime="application/pdf",
    )
with export_col2:
    st.download_button(
        label="📝 Download Report Draft (.txt)",
        data=full_report_text.encode("utf-8"),
        file_name="Project_Budget_Report.txt",
        mime="text/plain",
    )

st.success("Version 7 rebuilt cleanly with Excel export, PDF export, report writing, and branding support.")