"""
Financial Data Analysis Project
Author: Shashikant Yadav
Description: Analyzes financial datasets to identify trends and performance indicators.
Tech Used: Python, Pandas, Matplotlib, Excel
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import os

# ─────────────────────────────────────────
# 1. Load Data
# ─────────────────────────────────────────
df = pd.read_csv("financial_data.csv")
print("=" * 50)
print("FINANCIAL DATA ANALYSIS REPORT")
print("=" * 50)
print(f"\nDataset Shape: {df.shape[0]} rows x {df.shape[1]} columns")
print("\n--- Raw Data Preview ---")
print(df.to_string(index=False))

# ─────────────────────────────────────────
# 2. Data Cleaning & Validation
# ─────────────────────────────────────────
print("\n--- Data Quality Check ---")
print(f"Missing values: {df.isnull().sum().sum()}")
print(f"Duplicate rows: {df.duplicated().sum()}")

# ─────────────────────────────────────────
# 3. Key Performance Indicators (KPIs)
# ─────────────────────────────────────────
total_revenue  = df["Revenue"].sum()
total_expenses = df["Expenses"].sum()
total_profit   = df["Profit"].sum()
avg_margin     = (df["Profit"] / df["Revenue"] * 100).mean()
best_month     = df.loc[df["Profit"].idxmax(), "Month"]
worst_month    = df.loc[df["Profit"].idxmin(), "Month"]
revenue_growth = ((df["Revenue"].iloc[-1] - df["Revenue"].iloc[0]) / df["Revenue"].iloc[0]) * 100

print("\n--- Key Performance Indicators ---")
print(f"Total Annual Revenue  : ₹{total_revenue:,.0f}")
print(f"Total Annual Expenses : ₹{total_expenses:,.0f}")
print(f"Total Annual Profit   : ₹{total_profit:,.0f}")
print(f"Average Profit Margin : {avg_margin:.2f}%")
print(f"Best Performing Month : {best_month}")
print(f"Worst Performing Month: {worst_month}")
print(f"Revenue Growth (YTD)  : {revenue_growth:.2f}%")

# ─────────────────────────────────────────
# 4. Trend Analysis
# ─────────────────────────────────────────
df["Profit_Margin_%"] = (df["Profit"] / df["Revenue"] * 100).round(2)
df["MoM_Revenue_Growth_%"] = df["Revenue"].pct_change() * 100
df["MoM_Revenue_Growth_%"] = df["MoM_Revenue_Growth_%"].round(2)

print("\n--- Month-over-Month Trend ---")
print(df[["Month", "Revenue", "Profit", "Profit_Margin_%", "MoM_Revenue_Growth_%"]].to_string(index=False))

# ─────────────────────────────────────────
# 5. Visualizations
# ─────────────────────────────────────────
os.makedirs("charts", exist_ok=True)
months = df["Month"]
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle("Financial Performance Analysis — FY 2024", fontsize=16, fontweight="bold", y=1.01)

# Chart 1: Revenue vs Expenses vs Profit
ax1 = axes[0, 0]
x = range(len(months))
width = 0.28
ax1.bar([i - width for i in x], df["Revenue"],   width=width, label="Revenue",   color="#2563eb")
ax1.bar([i          for i in x], df["Expenses"],  width=width, label="Expenses",  color="#ef4444")
ax1.bar([i + width  for i in x], df["Profit"],    width=width, label="Profit",    color="#16a34a")
ax1.set_title("Revenue vs Expenses vs Profit", fontweight="bold")
ax1.set_xticks(list(x))
ax1.set_xticklabels(months, rotation=45, fontsize=8)
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"₹{v/1e5:.1f}L"))
ax1.legend()
ax1.grid(axis="y", alpha=0.3)

# Chart 2: Profit Trend Line
ax2 = axes[0, 1]
ax2.plot(months, df["Profit"], marker="o", color="#16a34a", linewidth=2.5, markersize=6)
ax2.fill_between(range(len(months)), df["Profit"], alpha=0.15, color="#16a34a")
ax2.set_title("Profit Trend (Monthly)", fontweight="bold")
ax2.set_xticks(range(len(months)))
ax2.set_xticklabels(months, rotation=45, fontsize=8)
ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"₹{v/1e5:.1f}L"))
ax2.grid(alpha=0.3)

# Chart 3: Profit Margin %
ax3 = axes[1, 0]
colors = ["#16a34a" if m >= avg_margin else "#f97316" for m in df["Profit_Margin_%"]]
bars = ax3.bar(months, df["Profit_Margin_%"], color=colors)
ax3.axhline(y=avg_margin, color="red", linestyle="--", linewidth=1.5, label=f"Avg: {avg_margin:.1f}%")
ax3.set_title("Profit Margin % by Month", fontweight="bold")
ax3.set_xticklabels(months, rotation=45, fontsize=8)
ax3.set_ylabel("Margin %")
ax3.legend()
ax3.grid(axis="y", alpha=0.3)
for bar, val in zip(bars, df["Profit_Margin_%"]):
    ax3.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3, f"{val:.0f}%", ha="center", fontsize=7)

# Chart 4: Sales Units vs Customer Count
ax4 = axes[1, 1]
ax4.plot(months, df["Sales_Units"],    marker="s", color="#7c3aed", linewidth=2, label="Sales Units")
ax4_twin = ax4.twinx()
ax4_twin.plot(months, df["Customer_Count"], marker="^", color="#db2777", linewidth=2, linestyle="--", label="Customers")
ax4.set_title("Sales Units vs Customer Count", fontweight="bold")
ax4.set_xticks(range(len(months)))
ax4.set_xticklabels(months, rotation=45, fontsize=8)
ax4.set_ylabel("Sales Units", color="#7c3aed")
ax4_twin.set_ylabel("Customer Count", color="#db2777")
ax4.grid(alpha=0.3)
lines1, labels1 = ax4.get_legend_handles_labels()
lines2, labels2 = ax4_twin.get_legend_handles_labels()
ax4.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=8)

plt.tight_layout()
plt.savefig("charts/financial_analysis_charts.png", dpi=150, bbox_inches="tight")
plt.close()
print("\n✅ Charts saved to: charts/financial_analysis_charts.png")

# ─────────────────────────────────────────
# 6. Export to Excel Report
# ─────────────────────────────────────────
with pd.ExcelWriter("Financial_Analysis_Report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Raw Data", index=False)

    summary = pd.DataFrame({
        "KPI": ["Total Revenue", "Total Expenses", "Total Profit", "Avg Profit Margin %",
                "Best Month", "Worst Month", "Revenue Growth %"],
        "Value": [f"₹{total_revenue:,.0f}", f"₹{total_expenses:,.0f}", f"₹{total_profit:,.0f}",
                  f"{avg_margin:.2f}%", best_month, worst_month, f"{revenue_growth:.2f}%"]
    })
    summary.to_excel(writer, sheet_name="KPI Summary", index=False)
    df[["Month","Profit_Margin_%","MoM_Revenue_Growth_%"]].to_excel(writer, sheet_name="Trend Analysis", index=False)

print("✅ Excel report saved: Financial_Analysis_Report.xlsx")
print("\n✅ Analysis Complete!")
