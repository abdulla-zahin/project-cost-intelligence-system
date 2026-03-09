import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date

st.set_page_config(page_title="Project Cost Intelligence System", page_icon="🏗️", layout="wide")

# -----------------------------
# Default configuration
# -----------------------------
DEFAULT_SPLIT = {
    "Materials": 40,
    "Labor": 25,
    "Transportation": 5,
    "Office Expense": 5,
    "Salaries / Overheads": 10,
    "Company Profit": 10,
    "Contingency": 5,
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

total_budget = st.number_input(
    "Total Project Budget",
    min_value=0.0,
    value=500000.0,
    step=10000.0,
    format="%.2f",
)

st.markdown("### Recommended Allocation")
st.write("Default smart recommendation is loaded below. You can adjust the percentages if needed.")

alloc_col1, alloc_col2 = st.columns(2)

allocations = {}
items = list(DEFAULT_SPLIT.items())
mid = len(items) // 2 + len(items) % 2

with alloc_col1:
    for category, value in items[:mid]:
        allocations[category] = st.number_input(
            f"{category} (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(value),
            step=1.0,
            key=f"alloc_{category}",
        )

with alloc_col2:
    for category, value in items[mid:]:
        allocations[category] = st.number_input(
            f"{category} (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(value),
            step=1.0,
            key=f"alloc_{category}",
        )

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
st.write(f"**Company Name:** {company_name}")
st.write(f"**Author:** {author_name}")
st.write(f"**Date:** {today}")

st.success("Project Cost Intelligence System is ready for estimation and analysis.")
