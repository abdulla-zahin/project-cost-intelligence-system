import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import date

st.set_page_config(page_title="Project Cost Intelligence System", page_icon="🏗️", layout="wide")

# -----------------------------
# Default configuration
# -----------------------------
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


def build_budget_dataframe(total_budget: float, allocations: dict) -> pd.DataFrame:
    rows = []
    for category, pct in allocations.items():
        amount = total_budget * (pct / 100)
        rows.append({
            "Category": category,
            "Allocation %": round(pct, 2),
            "Amount": round(amount, 2),
        })
    return pd.DataFrame(rows)


def analyze_budget(allocations: dict, total_pct: float) -> list[str]:
    insights = []

    if total_pct < 100:
        insights.append("⚠️ Total allocation is below 100%. Some budget is still unassigned.")
    elif total_pct > 100:
        insights.append("❌ Total allocation exceeds 100%. Adjust the percentages to match the budget.")
    else:
        insights.append("✅ Total allocation is perfectly balanced at 100%.")

    profit = allocations["Company Profit"]
    labor = allocations["Labor"]
    materials = allocations["Materials"]
    contingency = allocations["Contingency"]
    transport = allocations["Transportation"]

    if profit < 8:
        insights.append("⚠️ Profit margin is below 8%. This may be too tight for comfortable execution.")
    elif 8 <= profit <= 15:
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


def create_excel_report(project_name: str, project_type: str, company_name: str, author_name: str, today: str, total_budget: float, allocations: dict, df: pd.DataFrame, insights: list[str]) -> BytesIO:
    output = BytesIO()

    project_info_df = pd.DataFrame(
        {
            "Field": [
                "Project Name",
                "Project Type",
                "Company Name",
                "Author",
                "Date",
                "Total Budget",
                "Total Allocation %",
                "Estimated Profit %",
                "Estimated Profit Amount",
            ],
            "Value": [
                project_name,
                project_type,
                company_name,
                author_name,
                today,
                round(total_budget, 2),
                round(sum(allocations.values()), 2),
                round(allocations["Company Profit"], 2),
                round(total_budget * allocations["Company Profit"] / 100, 2),
            ],
        }
    )

    insights_df = pd.DataFrame({"Budget Insights": insights})

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        project_info_df.to_excel(writer, index=False, sheet_name="Project Info")
        df.to_excel(writer, index=False, sheet_name="Budget Breakdown")
        insights_df.to_excel(writer, index=False, sheet_name="Budget Insights")

        for sheet_name, worksheet in writer.sheets.items():
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                worksheet.column_dimensions[column_letter].width = max_length + 2

    output.seek(0)
    return output


# -----------------------------
# Header
# -----------------------------
st.title("🏗️ Project Cost Intelligence System")
st.caption("Smart budget allocation, profit analysis, and planning dashboard for engineering / contracting projects.")

# -----------------------------
# Project details
# -----------------------------
st.subheader("Project Details")
col1, col2 = st.columns(2)

with col1:
    project_name = st.text_input("Project Name", value="New Engineering Project")
    company_name = st.text_input("Company Name", value="TKM")

with col2:
    author_name = "Abdulla Zahin"
    today = date.today().strftime("%d %B %Y")
    st.text_input("Author", value=author_name, disabled=True)
    st.text_input("Date", value=today, disabled=True)

# -----------------------------
# Budget input
# -----------------------------
st.subheader("Budget Input")

project_type = st.selectbox(
    "Project Type",
    options=list(PROJECT_TYPE_SPLITS.keys()),
    index=0,
)

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
        if widget_key not in st.session_state:
            st.session_state[widget_key] = float(value)
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
        if widget_key not in st.session_state:
            st.session_state[widget_key] = float(value)
        allocations[category] = st.number_input(
            f"{category} (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(st.session_state.allocations_state[category]),
            step=1.0,
            key=widget_key,
        )
        st.session_state.allocations_state[category] = allocations[category]

total_pct = sum(allocations.values())
remaining_pct = 100 - total_pct

info1, info2, info3, info4 = st.columns(4)
info1.metric("Total Allocation %", f"{total_pct:.2f}%")
info2.metric("Unallocated / Excess %", f"{remaining_pct:.2f}%")
info3.metric("Estimated Profit %", f"{allocations['Company Profit']:.2f}%")
info4.metric("Estimated Profit Amount", f"{total_budget * allocations['Company Profit'] / 100:,.2f}")

# -----------------------------
# Data table
# -----------------------------
df = build_budget_dataframe(total_budget, allocations)

st.subheader("Budget Breakdown")
st.dataframe(df, use_container_width=True, hide_index=True)

excel_file = create_excel_report(
    project_name=project_name,
    project_type=project_type,
    company_name=company_name,
    author_name=author_name,
    today=today,
    total_budget=total_budget,
    allocations=allocations,
    df=df,
    insights=analyze_budget(allocations, total_pct),
)

# -----------------------------
# Analysis
# -----------------------------
st.subheader("Budget Intelligence")
insights = analyze_budget(allocations, total_pct)
for item in insights:
    st.write(item)

# -----------------------------
# Summary cards
# -----------------------------
st.subheader("Project Summary")
summary1, summary2, summary3, summary4 = st.columns(4)

profit_amount = total_budget * allocations["Company Profit"] / 100
materials_amount = total_budget * allocations["Materials"] / 100
labor_amount = total_budget * allocations["Labor"] / 100
contingency_amount = total_budget * allocations["Contingency"] / 100

summary1.metric("Total Budget", f"{total_budget:,.2f}")
summary2.metric("Materials", f"{materials_amount:,.2f}")
summary3.metric("Labor", f"{labor_amount:,.2f}")
summary4.metric("Contingency", f"{contingency_amount:,.2f}")

st.metric("Company Profit", f"{profit_amount:,.2f}", f"{allocations['Company Profit']:.2f}% margin")

# -----------------------------
# Charts
# -----------------------------
st.subheader("Visual Dashboard")
chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    fig1, ax1 = plt.subplots(figsize=(6, 6))
    ax1.pie(df["Amount"], labels=df["Category"], autopct="%1.1f%%", startangle=90)
    ax1.set_title("Budget Distribution")
    st.pyplot(fig1)

with chart_col2:
    fig2, ax2 = plt.subplots(figsize=(8, 5))
    ax2.bar(df["Category"], df["Amount"])
    ax2.set_title("Category-wise Budget Amount")
    ax2.set_xlabel("Category")
    ax2.set_ylabel("Amount")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig2)

# -----------------------------
# Report footer
# -----------------------------
st.markdown("---")
st.markdown("### Report Metadata")
st.write(f"**Project Name:** {project_name}")
st.write(f"**Project Type:** {project_type}")
st.write(f"**Company Name:** {company_name}")
st.write(f"**Author:** {author_name}")
st.write(f"**Date:** {today}")

st.download_button(
    label="📥 Download Excel Report",
    data=excel_file,
    file_name="Project_Budget_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("Project Cost Intelligence System is ready for estimation, analysis, and Excel reporting.")
