"""
Comprehensive Data Analysis for E-Commerce Superstore Dataset
Complete data cleaning, analysis, and visualization
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
import warnings
import os
from pathlib import Path

warnings.filterwarnings('ignore')
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (15, 8)

# Create graphs folder
graphs_dir = Path('graphs')
graphs_dir.mkdir(exist_ok=True)

print("=" * 80)
print("COMPREHENSIVE DATA ANALYSIS - E-COMMERCE SUPERSTORE DATASET")
print("=" * 80)

# ============================================================================
# 1. LOAD AND EXPLORE DATA
# ============================================================================
print("\n1. LOADING AND EXPLORING DATA...")
df = pd.read_csv('dataset.csv', encoding='latin-1')
print(f"Dataset shape: {df.shape}")
print(f"\nFirst few rows:\n{df.head()}")

# ============================================================================
# 2. DATA CLEANING
# ============================================================================
print("\n2. DATA CLEANING...")
initial_rows = len(df)

# Convert date columns
df['Order Date'] = pd.to_datetime(df['Order Date'], format='%m/%d/%Y')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], format='%m/%d/%Y')

# Remove duplicates
df = df.drop_duplicates()
print(f"Duplicates removed: {initial_rows - len(df)}")

# Handle missing values
df = df.dropna(subset=['Sales', 'Profit', 'Quantity'])

# Remove invalid data
df = df[df['Sales'] > 0]
df = df[df['Quantity'] > 0]

print(f"Final dataset shape: {df.shape}")
print(f"\nData types:\n{df.dtypes}")

# ============================================================================
# 3. BASIC STATISTICS
# ============================================================================
print("\n3. GENERATING STATISTICS...")
print(f"\nNumeric columns statistics:\n{df.describe()}")

# ============================================================================
# 4. CREATE VISUALIZATIONS
# ============================================================================
print("\n4. GENERATING VISUALIZATIONS...")

# 4.1 Sales Distribution
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Sales Distribution Analysis', fontsize=16, fontweight='bold')
axes[0, 0].hist(df['Sales'], bins=50, color='skyblue', edgecolor='black')
axes[0, 0].set_title('Distribution of Sales', fontweight='bold')
axes[0, 0].set_xlabel('Sales ($)')
axes[0, 0].set_ylabel('Frequency')
axes[0, 1].hist(df['Profit'], bins=50, color='lightgreen', edgecolor='black')
axes[0, 1].set_title('Distribution of Profit', fontweight='bold')
axes[0, 1].set_xlabel('Profit ($)')
axes[0, 1].set_ylabel('Frequency')
axes[1, 0].hist(df['Quantity'], bins=30, color='salmon', edgecolor='black')
axes[1, 0].set_title('Distribution of Quantity', fontweight='bold')
axes[1, 0].set_xlabel('Quantity')
axes[1, 0].set_ylabel('Frequency')
axes[1, 1].hist(df['Discount'], bins=30, color='lightyellow', edgecolor='black')
axes[1, 1].set_title('Distribution of Discount', fontweight='bold')
axes[1, 1].set_xlabel('Discount')
axes[1, 1].set_ylabel('Frequency')
plt.tight_layout()
plt.savefig(graphs_dir / '01_sales_distribution.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 01_sales_distribution.png")

# 4.2 Box plots
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Outlier Detection - Box Plots', fontsize=16, fontweight='bold')
axes[0, 0].boxplot(df['Sales'])
axes[0, 0].set_title('Sales Outliers', fontweight='bold')
axes[0, 0].set_ylabel('Sales ($)')
axes[0, 1].boxplot(df['Profit'])
axes[0, 1].set_title('Profit Outliers', fontweight='bold')
axes[0, 1].set_ylabel('Profit ($)')
axes[1, 0].boxplot(df['Quantity'])
axes[1, 0].set_title('Quantity Outliers', fontweight='bold')
axes[1, 0].set_ylabel('Quantity')
axes[1, 1].boxplot(df['Discount'])
axes[1, 1].set_title('Discount Outliers', fontweight='bold')
axes[1, 1].set_ylabel('Discount')
plt.tight_layout()
plt.savefig(graphs_dir / '02_box_plots_outliers.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 02_box_plots_outliers.png")

# 4.3 Sales by Category
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle('Sales Analysis by Category', fontsize=16, fontweight='bold')
category_sales = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
axes[0].bar(category_sales.index, category_sales.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
axes[0].set_title('Total Sales by Category', fontweight='bold')
axes[0].set_ylabel('Sales ($)')
for i, v in enumerate(category_sales.values):
    axes[0].text(i, v, f'${v:,.0f}', ha='center', va='bottom', fontweight='bold')
axes[1].pie(category_sales.values, labels=category_sales.index, autopct='%1.1f%%')
axes[1].set_title('Sales Distribution by Category', fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '03_sales_by_category.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 03_sales_by_category.png")

# 4.4 Sales by Segment
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle('Sales Analysis by Customer Segment', fontsize=16, fontweight='bold')
segment_sales = df.groupby('Segment')['Sales'].sum().sort_values(ascending=False)
axes[0].bar(segment_sales.index, segment_sales.values, color=['#95E1D3', '#F38181', '#AA96DA'])
axes[0].set_title('Total Sales by Segment', fontweight='bold')
axes[0].set_ylabel('Sales ($)')
for i, v in enumerate(segment_sales.values):
    axes[0].text(i, v, f'${v:,.0f}', ha='center', va='bottom', fontweight='bold')
axes[1].pie(segment_sales.values, labels=segment_sales.index, autopct='%1.1f%%')
axes[1].set_title('Sales Distribution by Segment', fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '04_sales_by_segment.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 04_sales_by_segment.png")

# 4.5 Sales by Region
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle('Sales Analysis by Region', fontsize=16, fontweight='bold')
region_sales = df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
axes[0].barh(region_sales.index, region_sales.values, color=['#FCBAD3', '#A8D8EA', '#AA96DA', '#FFB6C1'])
axes[0].set_title('Total Sales by Region', fontweight='bold')
axes[0].set_xlabel('Sales ($)')
axes[1].pie(region_sales.values, labels=region_sales.index, autopct='%1.1f%%')
axes[1].set_title('Sales Distribution by Region', fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '05_sales_by_region.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 05_sales_by_region.png")

# 4.6 Profit by Category and Segment
fig, ax = plt.subplots(figsize=(14, 7))
profit_data = df.groupby(['Category', 'Segment'])['Profit'].sum().unstack()
profit_data.plot(kind='bar', ax=ax, color=['#95E1D3', '#F38181', '#AA96DA'])
ax.set_title('Profit by Category and Segment', fontsize=14, fontweight='bold')
ax.set_ylabel('Profit ($)')
ax.set_xlabel('Category')
ax.legend(title='Segment')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(graphs_dir / '06_profit_category_segment.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 06_profit_category_segment.png")

# 4.7 Top 10 Sub-Categories by Sales
fig, ax = plt.subplots(figsize=(14, 8))
subcategory_sales = df.groupby('Sub-Category')['Sales'].sum().sort_values(ascending=False).head(10)
ax.barh(subcategory_sales.index, subcategory_sales.values, color='steelblue')
ax.set_title('Top 10 Sub-Categories by Sales', fontsize=14, fontweight='bold')
ax.set_xlabel('Sales ($)')
for i, v in enumerate(subcategory_sales.values):
    ax.text(v, i, f' ${v:,.0f}', ha='left', va='center', fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '07_top_subcategories_sales.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 07_top_subcategories_sales.png")

# 4.8 Discount vs Profit
fig, ax = plt.subplots(figsize=(12, 8))
scatter = ax.scatter(df['Discount'], df['Profit'], c=df['Sales'], cmap='viridis', alpha=0.6, s=50)
ax.set_title('Discount vs Profit Relationship', fontsize=14, fontweight='bold')
ax.set_xlabel('Discount')
ax.set_ylabel('Profit ($)')
cbar = plt.colorbar(scatter, ax=ax)
cbar.set_label('Sales ($)')
z = np.polyfit(df['Discount'], df['Profit'], 1)
p = np.poly1d(z)
ax.plot(np.sort(df['Discount']), p(np.sort(df['Discount'])), "r--", alpha=0.8, linewidth=2)
plt.tight_layout()
plt.savefig(graphs_dir / '08_discount_vs_profit.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 08_discount_vs_profit.png")

# 4.9 Sales Over Time
fig, ax = plt.subplots(figsize=(16, 6))
monthly_sales = df.groupby(df['Order Date'].dt.to_period('M'))['Sales'].sum()
monthly_sales.index = monthly_sales.index.to_timestamp()
ax.plot(monthly_sales.index, monthly_sales.values, marker='o', linewidth=2, color='steelblue')
ax.fill_between(monthly_sales.index, monthly_sales.values, alpha=0.3)
ax.set_title('Monthly Sales Trend', fontsize=14, fontweight='bold')
ax.set_xlabel('Date')
ax.set_ylabel('Sales ($)')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(graphs_dir / '09_sales_trend.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 09_sales_trend.png")

# 4.10 Shipping Mode Performance
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle('Shipping Mode Performance', fontsize=16, fontweight='bold')
ship_sales = df.groupby('Ship Mode')['Sales'].sum().sort_values(ascending=False)
ship_profit = df.groupby('Ship Mode')['Profit'].sum().sort_values(ascending=False)
axes[0].bar(ship_sales.index, ship_sales.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A'])
axes[0].set_title('Sales by Shipping Mode', fontweight='bold')
axes[0].set_ylabel('Sales ($)')
axes[1].bar(ship_profit.index, ship_profit.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A'])
axes[1].set_title('Profit by Shipping Mode', fontweight='bold')
axes[1].set_ylabel('Profit ($)')
plt.tight_layout()
plt.savefig(graphs_dir / '10_shipping_mode.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 10_shipping_mode.png")

# 4.11 Top 20 States by Sales
fig, ax = plt.subplots(figsize=(14, 10))
state_sales = df.groupby('State')['Sales'].sum().sort_values(ascending=False).head(20)
ax.barh(state_sales.index, state_sales.values, color='teal')
ax.set_title('Top 20 States by Sales', fontsize=14, fontweight='bold')
ax.set_xlabel('Sales ($)')
plt.tight_layout()
plt.savefig(graphs_dir / '11_top_states.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 11_top_states.png")

# 4.12 Correlation Heatmap
fig, ax = plt.subplots(figsize=(10, 8))
numeric_cols = ['Sales', 'Quantity', 'Discount', 'Profit']
correlation = df[numeric_cols].corr()
sns.heatmap(correlation, annot=True, cmap='coolwarm', center=0, ax=ax, square=True, linewidths=1)
ax.set_title('Correlation Matrix - Numeric Variables', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '12_correlation_heatmap.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 12_correlation_heatmap.png")

# 4.13 Profit Margin Analysis
df['Profit_Margin'] = (df['Profit'] / df['Sales'] * 100).round(2)
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle('Profit Margin Analysis', fontsize=16, fontweight='bold')
profit_margin_cat = df.groupby('Category')['Profit_Margin'].mean().sort_values(ascending=False)
axes[0].bar(profit_margin_cat.index, profit_margin_cat.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
axes[0].set_title('Average Profit Margin by Category', fontweight='bold')
axes[0].set_ylabel('Profit Margin (%)')
profit_margin_seg = df.groupby('Segment')['Profit_Margin'].mean().sort_values(ascending=False)
axes[1].bar(profit_margin_seg.index, profit_margin_seg.values, color=['#95E1D3', '#F38181', '#AA96DA'])
axes[1].set_title('Average Profit Margin by Segment', fontweight='bold')
axes[1].set_ylabel('Profit Margin (%)')
plt.tight_layout()
plt.savefig(graphs_dir / '13_profit_margin.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 13_profit_margin.png")

# 4.14 Customer Count Analysis
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Customer Analysis', fontsize=16, fontweight='bold')
cust_cat = df.groupby('Category')['Customer ID'].nunique()
axes[0, 0].bar(cust_cat.index, cust_cat.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
axes[0, 0].set_title('Unique Customers by Category', fontweight='bold')
cust_seg = df.groupby('Segment')['Customer ID'].nunique()
axes[0, 1].bar(cust_seg.index, cust_seg.values, color=['#95E1D3', '#F38181', '#AA96DA'])
axes[0, 1].set_title('Unique Customers by Segment', fontweight='bold')
cust_reg = df.groupby('Region')['Customer ID'].nunique().sort_values(ascending=False)
axes[1, 0].barh(cust_reg.index, cust_reg.values)
axes[1, 0].set_title('Unique Customers by Region', fontweight='bold')
avg_cust_sales = df.groupby('Segment').apply(lambda x: x['Sales'].sum() / x['Customer ID'].nunique()).sort_values(ascending=False)
axes[1, 1].bar(avg_cust_sales.index, avg_cust_sales.values, color=['#95E1D3', '#F38181', '#AA96DA'])
axes[1, 1].set_title('Average Sales per Customer by Segment', fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '14_customer_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 14_customer_analysis.png")

# 4.15 Sub-Category Profitability
fig, ax = plt.subplots(figsize=(14, 10))
subcat_profit = df.groupby('Sub-Category')['Profit'].sum().sort_values()
colors_list = ['red' if x < 0 else 'green' for x in subcat_profit.values]
ax.barh(subcat_profit.index, subcat_profit.values, color=colors_list)
ax.set_title('Profitability by Sub-Category', fontsize=14, fontweight='bold')
ax.set_xlabel('Profit ($)')
ax.axvline(x=0, color='black', linestyle='-', linewidth=0.8)
plt.tight_layout()
plt.savefig(graphs_dir / '15_subcategory_profitability.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 15_subcategory_profitability.png")

# 4.16 Quantity Analysis
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Quantity Analysis', fontsize=16, fontweight='bold')
qty_cat = df.groupby('Category')['Quantity'].sum().sort_values(ascending=False)
axes[0, 0].bar(qty_cat.index, qty_cat.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
axes[0, 0].set_title('Total Quantity by Category', fontweight='bold')
qty_seg = df.groupby('Segment')['Quantity'].sum().sort_values(ascending=False)
axes[0, 1].bar(qty_seg.index, qty_seg.values, color=['#95E1D3', '#F38181', '#AA96DA'])
axes[0, 1].set_title('Total Quantity by Segment', fontweight='bold')
axes[1, 0].scatter(df['Quantity'], df['Sales'], alpha=0.5)
axes[1, 0].set_title('Quantity vs Sales Relationship', fontweight='bold')
axes[1, 0].set_xlabel('Quantity')
axes[1, 0].set_ylabel('Sales ($)')
avg_qty_sub = df.groupby('Sub-Category')['Quantity'].mean().sort_values(ascending=False).head(10)
axes[1, 1].barh(avg_qty_sub.index, avg_qty_sub.values, color='steelblue')
axes[1, 1].set_title('Top 10 Sub-Categories by Average Quantity', fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '16_quantity_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 16_quantity_analysis.png")

# 4.17 ROI Analysis
df['ROI'] = (df['Profit'] / (df['Sales'] - df['Profit'] + 0.001) * 100).replace([np.inf, -np.inf], 0)
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle('ROI Analysis by Category and Segment', fontsize=16, fontweight='bold')
roi_cat = df.groupby('Category')['ROI'].mean().sort_values(ascending=False)
axes[0].bar(roi_cat.index, roi_cat.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1'])
axes[0].set_title('Average ROI by Category', fontweight='bold')
axes[0].set_ylabel('ROI (%)')
roi_seg = df.groupby('Segment')['ROI'].mean().sort_values(ascending=False)
axes[1].bar(roi_seg.index, roi_seg.values, color=['#95E1D3', '#F38181', '#AA96DA'])
axes[1].set_title('Average ROI by Segment', fontweight='bold')
axes[1].set_ylabel('ROI (%)')
plt.tight_layout()
plt.savefig(graphs_dir / '17_roi_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 17_roi_analysis.png")

# 4.18 Shipment Days Analysis
df['Shipment_Days'] = (df['Ship Date'] - df['Order Date']).dt.days
fig, axes = plt.subplots(2, 2, figsize=(16, 10))
fig.suptitle('Shipment Days Analysis', fontsize=16, fontweight='bold')
axes[0, 0].hist(df['Shipment_Days'], bins=30, color='skyblue', edgecolor='black')
axes[0, 0].set_title('Distribution of Shipment Days', fontweight='bold')
shipment_mode = df.groupby('Ship Mode')['Shipment_Days'].mean().sort_values(ascending=False)
axes[0, 1].bar(shipment_mode.index, shipment_mode.values, color=['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A'])
axes[0, 1].set_title('Average Shipment Days by Mode', fontweight='bold')
shipment_reg = df.groupby('Region')['Shipment_Days'].mean().sort_values(ascending=False)
axes[1, 0].barh(shipment_reg.index, shipment_reg.values)
axes[1, 0].set_title('Average Shipment Days by Region', fontweight='bold')
axes[1, 1].boxplot([df[df['Ship Mode'] == mode]['Shipment_Days'].values for mode in df['Ship Mode'].unique()], labels=df['Ship Mode'].unique())
axes[1, 1].set_title('Shipment Days Distribution by Mode', fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '18_shipment_analysis.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 18_shipment_analysis.png")

# 4.19 Monthly Profit Trend
fig, ax = plt.subplots(figsize=(16, 6))
monthly_profit = df.groupby(df['Order Date'].dt.to_period('M'))['Profit'].sum()
monthly_profit.index = monthly_profit.index.to_timestamp()
ax.plot(monthly_profit.index, monthly_profit.values, marker='o', linewidth=2, color='green')
ax.fill_between(monthly_profit.index, monthly_profit.values, alpha=0.3, color='green')
ax.set_title('Monthly Profit Trend', fontsize=14, fontweight='bold')
ax.set_xlabel('Date')
ax.set_ylabel('Profit ($)')
ax.axhline(y=0, color='red', linestyle='--', linewidth=1)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(graphs_dir / '19_profit_trend.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 19_profit_trend.png")

# 4.20 Category vs Sub-Category Heatmap
fig, ax = plt.subplots(figsize=(14, 10))
cat_subcat = df.pivot_table(values='Sales', index='Sub-Category', columns='Category', aggfunc='sum', fill_value=0)
sns.heatmap(cat_subcat, annot=True, fmt='.0f', cmap='YlGnBu', ax=ax)
ax.set_title('Sales Heatmap: Sub-Category vs Category', fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig(graphs_dir / '20_category_subcategory_heatmap.png', dpi=300, bbox_inches='tight')
plt.close()
print("✓ Saved: 20_category_subcategory_heatmap.png")

# ============================================================================
# 5. SAVE RESULTS
# ============================================================================
print("\n5. SAVING RESULTS...")
df.to_csv(graphs_dir / 'cleaned_data.csv', index=False)
print("✓ Saved: cleaned_data.csv")

# Save summary statistics
with open(graphs_dir / 'analysis_summary.txt', 'w') as f:
    f.write("=" * 80 + "\n")
    f.write("DATA ANALYSIS SUMMARY\n")
    f.write("=" * 80 + "\n\n")
    f.write(f"Total Records: {len(df):,}\n")
    f.write(f"Date Range: {df['Order Date'].min().date()} to {df['Order Date'].max().date()}\n")
    f.write(f"Total Sales: ${df['Sales'].sum():,.2f}\n")
    f.write(f"Total Profit: ${df['Profit'].sum():,.2f}\n")
    f.write(f"Profit Margin: {(df['Profit'].sum() / df['Sales'].sum() * 100):.2f}%\n")
    f.write(f"Unique Customers: {df['Customer ID'].nunique():,}\n\n")
    
    for cat in df['Category'].unique():
        cat_df = df[df['Category'] == cat]
        f.write(f"\n{cat}:\n")
        f.write(f"  Sales: ${cat_df['Sales'].sum():,.2f}\n")
        f.write(f"  Profit: ${cat_df['Profit'].sum():,.2f}\n")

print("✓ Saved: analysis_summary.txt")

print("\n" + "=" * 80)
print("ANALYSIS COMPLETE!")
print("=" * 80)
print(f"\n✓ Created 'graphs' folder with 20 visualizations")
print(f"✓ Cleaned data saved as cleaned_data.csv")
print(f"✓ Ready for detailed report generation!")
