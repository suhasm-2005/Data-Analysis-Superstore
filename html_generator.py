#!/usr/bin/env python3
"""
HTML Report Generator for E-Commerce Superstore Analysis
Converts analysis content and visualizations into an interactive HTML website
"""

import os
import base64
from pathlib import Path

def encode_image_to_base64(image_path):
    """Convert image to base64 string for embedding in HTML"""
    try:
        with open(image_path, 'rb') as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8')
    except Exception as e:
        print(f"Error encoding image {image_path}: {e}")
        return None

def generate_html():
    """Generate comprehensive HTML report"""
    
    # Get script directory
    script_dir = Path(__file__).parent
    graphs_dir = script_dir / 'graphs'
    
    # Image file names in order
    images = [
        '01_sales_distribution.png',
        '02_box_plots_outliers.png',
        '03_sales_by_category.png',
        '04_sales_by_segment.png',
        '05_sales_by_region.png',
        '06_profit_category_segment.png',
        '07_top_subcategories_sales.png',
        '08_discount_vs_profit.png',
        '09_sales_trend.png',
        '10_shipping_mode.png',
        '11_top_states.png',
        '12_correlation_heatmap.png',
        '13_profit_margin.png',
        '14_customer_analysis.png',
        '15_subcategory_profitability.png',
        '16_quantity_analysis.png',
        '17_roi_analysis.png',
        '18_shipment_analysis.png',
        '19_profit_trend.png',
        '20_category_subcategory_heatmap.png',
    ]
    
    # Encode all images
    encoded_images = {}
    for img_name in images:
        img_path = graphs_dir / img_name
        if img_path.exists():
            encoded_images[img_name] = encode_image_to_base64(str(img_path))
            print(f"✓ Encoded {img_name}")
        else:
            print(f"✗ Image not found: {img_name}")
    
    # HTML Content
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>E-Commerce Superstore Analysis Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background: #f5f7fa;
        }
        
        /* Header and Navigation */
        header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 2rem 0;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            position: sticky;
            top: 0;
            z-index: 100;
        }
        
        .header-content {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 20px;
        }
        
        h1 {
            font-size: 2.5rem;
            margin-bottom: 0.5rem;
            font-weight: 600;
        }
        
        .report-meta {
            font-size: 0.95rem;
            opacity: 0.95;
        }
        
        nav {
            background: #f5f7fa;
            padding: 1rem 0;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            position: sticky;
            top: 80px;
            z-index: 99;
        }
        
        .nav-content {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 20px;
        }
        
        .nav-links {
            display: flex;
            flex-wrap: wrap;
            gap: 2rem;
            list-style: none;
            overflow-x: auto;
        }
        
        .nav-links a {
            color: #667eea;
            text-decoration: none;
            font-weight: 500;
            transition: color 0.3s;
            white-space: nowrap;
        }
        
        .nav-links a:hover {
            color: #764ba2;
        }
        
        /* Main Content */
        main {
            max-width: 1200px;
            margin: 0 auto;
            padding: 2rem 20px;
        }
        
        section {
            background: white;
            padding: 2rem;
            margin-bottom: 2rem;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        
        h2 {
            color: #667eea;
            font-size: 2rem;
            margin-bottom: 1.5rem;
            border-bottom: 3px solid #667eea;
            padding-bottom: 0.5rem;
        }
        
        h3 {
            color: #764ba2;
            font-size: 1.5rem;
            margin-top: 1.5rem;
            margin-bottom: 1rem;
        }
        
        p {
            margin-bottom: 1rem;
            text-align: justify;
        }
        
        /* KPI Cards */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2rem;
        }
        
        .kpi-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 1.5rem;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.2);
            transition: transform 0.3s;
        }
        
        .kpi-card:hover {
            transform: translateY(-5px);
        }
        
        .kpi-value {
            font-size: 2rem;
            font-weight: bold;
            margin: 0.5rem 0;
        }
        
        .kpi-label {
            font-size: 0.9rem;
            opacity: 0.9;
        }
        
        /* Tables */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 1.5rem 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        
        th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }
        
        td {
            padding: 12px;
            border-bottom: 1px solid #f0f0f0;
        }
        
        tr:hover {
            background: #f9f9f9;
        }
        
        /* Insights Boxes */
        .insight {
            background: #f0f4ff;
            border-left: 4px solid #667eea;
            padding: 1rem;
            margin: 1rem 0;
            border-radius: 4px;
        }
        
        .insight strong {
            color: #667eea;
        }
        
        .warning-insight {
            background: #fff3cd;
            border-left-color: #ffc107;
        }
        
        .warning-insight strong {
            color: #ff6b6b;
        }
        
        /* Images */
        .chart-container {
            margin: 2rem 0;
            padding: 1rem;
            background: #f9f9f9;
            border-radius: 8px;
            text-align: center;
        }
        
        .chart-container img {
            max-width: 100%;
            height: auto;
            border-radius: 4px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .chart-title {
            margin-top: 1rem;
            color: #667eea;
            font-weight: 600;
        }
        
        /* Footer */
        footer {
            background: #f5f7fa;
            color: #666;
            text-align: center;
            padding: 2rem;
            margin-top: 3rem;
            border-top: 1px solid #ddd;
        }
        
        .scroll-top {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: #667eea;
            color: white;
            width: 50px;
            height: 50px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
            opacity: 0;
            transition: opacity 0.3s;
        }
        
        .scroll-top.show {
            opacity: 1;
        }
        
        .scroll-top:hover {
            background: #764ba2;
        }
        
        /* Responsive */
        @media (max-width: 768px) {
            h1 {
                font-size: 1.8rem;
            }
            
            .nav-links {
                gap: 1rem;
            }
            
            .kpi-grid {
                grid-template-columns: 1fr;
            }
            
            section {
                padding: 1.5rem;
            }
        }
    </style>
</head>
<body>
    <!-- Header -->
    <header>
        <div class="header-content">
            <h1>📊 E-Commerce Superstore Analysis Report</h1>
            <div class="report-meta">
                <strong>Analysis Period:</strong> January 3, 2014 - December 30, 2017 | 
                <strong>Report Generated:</strong> March 26, 2026 | 
                <strong>Records Analyzed:</strong> 9,994 transactions
            </div>
        </div>
    </header>
    
    <!-- Navigation -->
    <nav>
        <div class="nav-content">
            <ul class="nav-links">
                <li><a href="#summary">Executive Summary</a></li>
                <li><a href="#overview">Dataset Overview</a></li>
                <li><a href="#sales">Sales Analysis</a></li>
                <li><a href="#profit">Profit Analysis</a></li>
                <li><a href="#segments">Segment Analysis</a></li>
                <li><a href="#visualizations">All Visualizations</a></li>
                <li><a href="#recommendations">Recommendations</a></li>
            </ul>
        </div>
    </nav>
    
    <!-- Main Content -->
    <main>
        <!-- Executive Summary Section -->
        <section id="summary">
            <h2>📈 Executive Summary</h2>
            <p>This comprehensive analysis examines 9,994 transactions from an e-commerce superstore dataset spanning from January 2014 to December 2017. The analysis reveals critical insights into sales performance, profitability, customer behavior, and operational efficiency across multiple geographic markets and customer segments.</p>
            
            <h3>Key Performance Indicators</h3>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Total Sales Revenue</div>
                    <div class="kpi-value">$2,297,201</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Total Profit</div>
                    <div class="kpi-value">$286,397</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Profit Margin</div>
                    <div class="kpi-value">12.46%</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Unique Customers</div>
                    <div class="kpi-value">793</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Average Order Value</div>
                    <div class="kpi-value">$229.95</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Profitable Orders</div>
                    <div class="kpi-value">8,181 (81.9%)</div>
                </div>
            </div>
            
            <h3>Executive Insights</h3>
            <div class="insight">
                <strong>✓ Strong Revenue Base:</strong> The superstore generated total sales of $2,297,200.86 with an overall profit of $286,397.02
            </div>
            <div class="insight">
                <strong>✓ Healthy Order Performance:</strong> 81.9% of orders are profitable, indicating generally sound operational execution
            </div>
            <div class="insight warning-insight">
                <strong>⚠ Profitability Concern:</strong> Overall profit margin of 12.46% suggests significant room for optimization through operational efficiency and pricing strategy refinement
            </div>
            <div class="insight">
                <strong>✓ Customer Loyalty:</strong> 793 unique customers with approximately 70% repeat purchase rate demonstrate strong customer retention
            </div>
        </section>
        
        <!-- Dataset Overview -->
        <section id="overview">
            <h2>📋 Dataset Overview & Structure</h2>
            
            <h3>Dataset Dimensions</h3>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Total Records</div>
                    <div class="kpi-value">9,994</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Data Attributes</div>
                    <div class="kpi-value">21</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Date Range</div>
                    <div class="kpi-value">1,458 Days</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Geographic Scope</div>
                    <div class="kpi-value">49 States, 4 Regions</div>
                </div>
            </div>
            
            <h3>Data Cleaning Process</h3>
            <ul style="margin-left: 1rem; margin-bottom: 1rem;">
                <li style="margin-bottom: 0.5rem;"><strong>Duplicate Removal:</strong> All duplicate records were identified and systematically removed</li>
                <li style="margin-bottom: 0.5rem;"><strong>Missing Value Treatment:</strong> Records with missing values in critical fields were removed</li>
                <li style="margin-bottom: 0.5rem;"><strong>Data Type Conversion:</strong> Date columns converted to proper datetime format</li>
                <li style="margin-bottom: 0.5rem;"><strong>Outlier Detection:</strong> Statistical methods used to identify and document outliers</li>
                <li style="margin-bottom: 0.5rem;"><strong>Completeness Validation:</strong> 100% completeness achieved across all critical fields</li>
            </ul>
        </section>
        
        <!-- Sales Analysis -->
        <section id="sales">
            <h2>💰 Sales Performance Analysis</h2>
            
            <h3>Overall Sales Statistics</h3>
            <table>
                <tr>
                    <th>Metric</th>
                    <th>Value</th>
                </tr>
                <tr>
                    <td>Total Sales Revenue</td>
                    <td><strong>$2,297,200.86</strong></td>
                </tr>
                <tr>
                    <td>Average Sale per Transaction</td>
                    <td><strong>$229.95</strong></td>
                </tr>
                <tr>
                    <td>Median Sale Value</td>
                    <td><strong>$140.53</strong></td>
                </tr>
                <tr>
                    <td>Sales Range</td>
                    <td><strong>$0.44 - $22,638.48</strong></td>
                </tr>
            </table>
            
            <h3>Sales by Category</h3>
            <table>
                <tr>
                    <th>Category</th>
                    <th>Total Sales</th>
                    <th>Orders</th>
                    <th>Avg Sale</th>
                    <th>Customers</th>
                </tr>
                <tr>
                    <td>Technology</td>
                    <td>$836,153.36</td>
                    <td>1,847</td>
                    <td>$452.74</td>
                    <td>509</td>
                </tr>
                <tr>
                    <td>Furniture</td>
                    <td>$742,000.05</td>
                    <td>2,121</td>
                    <td>$349.67</td>
                    <td>419</td>
                </tr>
                <tr>
                    <td>Office Supplies</td>
                    <td>$719,047.45</td>
                    <td>6,026</td>
                    <td>$119.32</td>
                    <td>637</td>
                </tr>
            </table>
            
            <h3>Key Sales Insights</h3>
            <div class="insight">
                <strong>✓ Premium Category:</strong> Technology category generates the highest average order value at $452.74, indicating strong demand for high-value products
            </div>
            <div class="insight">
                <strong>✓ Volume Leader:</strong> Office Supplies dominates transaction volume with 6,026 orders, representing 60.3% of all transactions
            </div>
            <div class="insight">
                <strong>✓ Market Diversity:</strong> Three distinct product categories show different performance characteristics, suggesting need for category-specific strategies
            </div>
        </section>
        
        <!-- Profit Analysis -->
        <section id="profit">
            <h2>📊 Profit Analysis & Margins</h2>
            
            <h3>Overall Profitability Metrics</h3>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Total Profit</div>
                    <div class="kpi-value">$286,397</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Profit Margin</div>
                    <div class="kpi-value">12.46%</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Avg Profit/Order</div>
                    <div class="kpi-value">$28.67</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Profitable Orders</div>
                    <div class="kpi-value">81.9%</div>
                </div>
            </div>
            
            <h3>Profit Analysis by Category</h3>
            <table>
                <tr>
                    <th>Category</th>
                    <th>Total Profit</th>
                    <th>Profit Margin</th>
                    <th>Avg Profit/Order</th>
                    <th>Loss Count</th>
                </tr>
                <tr style="background: #fff3cd;">
                    <td><strong>Furniture</strong></td>
                    <td><strong style="color: #d32f2f;">-$17,725.48</strong></td>
                    <td><strong style="color: #d32f2f;">-2.39%</strong></td>
                    <td>-$8.36</td>
                    <td>548 orders</td>
                </tr>
                <tr>
                    <td><strong>Office Supplies</strong></td>
                    <td><strong>$123,843.05</strong></td>
                    <td>17.23%</td>
                    <td>$20.55</td>
                    <td>523 orders</td>
                </tr>
                <tr style="background: #e8f5e9;">
                    <td><strong>Technology</strong></td>
                    <td><strong>$180,279.45</strong></td>
                    <td>21.56%</td>
                    <td>$97.59</td>
                    <td>742 orders</td>
                </tr>
            </table>
            
            <h3>Critical Findings</h3>
            <div class="insight warning-insight">
                <strong>🚨 Critical Issue:</strong> Furniture category operating at -2.39% margin loss, representing major profitability concern requiring immediate intervention
            </div>
            <div class="insight">
                <strong>✓ Technology Leadership:</strong> Technology delivers highest margins at 21.56%, driven by premium pricing and lower discount rates
            </div>
            <div class="insight">
                <strong>✓ Office Supplies Stability:</strong> Office Supplies shows solid 17.23% margin, providing stable profit contribution and customer loyalty
            </div>
            <div class="insight warning-insight">
                <strong>⚠ Loss Orders Issue:</strong> 1,813 orders (18.1%) operate at a loss across all categories, indicating pricing or cost control problems
            </div>
        </section>
        
        <!-- Segment Analysis -->
        <section id="segments">
            <h2>👥 Customer & Geographic Analysis</h2>
            
            <h3>Customer Segments</h3>
            <table>
                <tr>
                    <th>Segment</th>
                    <th>Customer Count</th>
                    <th>Avg Repeat Rate</th>
                    <th>Business Impact</th>
                </tr>
                <tr>
                    <td>Consumer</td>
                    <td>Largest</td>
                    <td>~70%</td>
                    <td>Primary revenue driver</td>
                </tr>
                <tr>
                    <td>Corporate</td>
                    <td>Medium</td>
                    <td>~70%</td>
                    <td>Stable, reliable revenue</td>
                </tr>
                <tr>
                    <td>Home Office</td>
                    <td>Smallest</td>
                    <td>~70%</td>
                    <td>Growing segment opportunity</td>
                </tr>
            </table>
            
            <h3>Geographic Distribution</h3>
            <div class="insight">
                <strong>✓ Regional Presence:</strong> Operations across 4 regions (East, West, South, Central) and 49 states
            </div>
            <div class="insight">
                <strong>✓ Market Concentration:</strong> West region shows particularly strong performance and market penetration
            </div>
            <div class="insight">
                <strong>✓ State Leadership:</strong> California leads in sales contribution as the single largest state market
            </div>
        </section>
        
        <!-- All Visualizations -->
        <section id="visualizations">
            <h2>📊 Complete Visualization Suite</h2>
            <p>Below are all 20 comprehensive visualizations generated from the superstore dataset, providing detailed insights into various aspects of business performance.</p>
"""
    
    # Add image sections
    image_descriptions = {
        '01_sales_distribution.png': ('Sales Distribution', 'Distribution of sales values across all transactions, showing market spread and concentration'),
        '02_box_plots_outliers.png': ('Outlier Detection', 'Box plots identifying statistical outliers and anomalies in key metrics'),
        '03_sales_by_category.png': ('Sales by Category', 'Comparative sales performance across Furniture, Office Supplies, and Technology categories'),
        '04_sales_by_segment.png': ('Sales by Customer Segment', 'Sales breakdown across Consumer, Corporate, and Home Office segments'),
        '05_sales_by_region.png': ('Regional Sales Distribution', 'Geographic sales distribution across East, West, South, and Central regions'),
        '06_profit_category_segment.png': ('Profit Analysis', 'Detailed profit breakdown by category and customer segment combinations'),
        '07_top_subcategories_sales.png': ('Top Subcategories', 'Sales performance of top-performing product subcategories'),
        '08_discount_vs_profit.png': ('Discount Impact', 'Critical analysis showing negative correlation between discount levels and profitability'),
        '09_sales_trend.png': ('Sales Trends Over Time', 'Monthly and yearly sales trends showing seasonal patterns and growth trajectory'),
        '10_shipping_mode.png': ('Shipping Mode Analysis', 'Comparison of sales volume and margins across different shipping methods'),
        '11_top_states.png': ('Top States by Sales', 'Ranking of top-performing states in terms of sales contribution'),
        '12_correlation_heatmap.png': ('Variable Correlations', 'Correlation matrix showing relationships between key business variables'),
        '13_profit_margin.png': ('Profit Margin Trends', 'Evolution of profit margins over time and across product categories'),
        '14_customer_analysis.png': ('Customer Metrics', 'Customer lifetime value, repeat purchase rates, and segment characteristics'),
        '15_subcategory_profitability.png': ('Subcategory Profitability', 'Detailed profitability analysis for all product subcategories'),
        '16_quantity_analysis.png': ('Quantity-Sales Relationship', 'How order quantity impacts sales and profit metrics'),
        '17_roi_analysis.png': ('Return on Investment Analysis', 'ROI metrics by product category and customer segment'),
        '18_shipment_analysis.png': ('Shipment Performance', 'Shipping mode performance, delivery times, and cost impact analysis'),
        '19_profit_trend.png': ('Profit Trends', 'Detailed profit trends over the analysis period showing seasonality and growth'),
        '20_category_subcategory_heatmap.png': ('Category-Subcategory Matrix', 'Heatmap showing relationships and performance across categories and subcategories'),
    }
    
    for img_name in images:
        if img_name in image_descriptions:
            title, description = image_descriptions[img_name]
            if img_name in encoded_images and encoded_images[img_name]:
                html_content += f"""
            <div class="chart-container">
                <h3 style="text-align: left; color: #667eea; margin-bottom: 0.5rem;">{title}</h3>
                <p style="text-align: left; font-size: 0.9rem; color: #666; margin-bottom: 1rem;">{description}</p>
                <img src="data:image/png;base64,{encoded_images[img_name]}" alt="{title}">
            </div>
"""
    
    # Add Recommendations Section
    html_content += """
        </section>
        
        <!-- Recommendations -->
        <section id="recommendations">
            <h2>💡 Strategic Recommendations</h2>
            
            <h3>1. Furniture Category Rehabilitation (Highest Priority)</h3>
            <p><strong>Current Status:</strong> Operating at -2.39% profit margin with 548 losing orders</p>
            <p><strong>Recommendations:</strong></p>
            <ul style="margin-left: 1rem; margin-bottom: 1rem;">
                <li style="margin-bottom: 0.5rem;">Conduct comprehensive cost analysis to identify margin leakage points</li>
                <li style="margin-bottom: 0.5rem;">Review pricing strategy for high-margin furniture subcategories</li>
                <li style="margin-bottom: 0.5rem;">Implement stricter discount caps (max 10%) for furniture products</li>
                <li style="margin-bottom: 0.5rem;">Negotiate better supplier terms to reduce cost of goods sold</li>
            </ul>
            <p><strong>Impact:</strong> Potential margin improvement to +8-12%, generating additional $74,000-148,000 in annual profit</p>
            
            <h3>2. Discount Optimization Strategy</h3>
            <p><strong>Current Issue:</strong> Orders with &gt;20% discount show -8.93% margins</p>
            <p><strong>Recommendations:</strong></p>
            <ul style="margin-left: 1rem; margin-bottom: 1rem;">
                <li style="margin-bottom: 0.5rem;">Implement dynamic discount capping at 15% maximum across all categories</li>
                <li style="margin-bottom: 0.5rem;">Use targeted promotions instead of blanket discounts to preserve margins</li>
                <li style="margin-bottom: 0.5rem;">Create loyalty programs to encourage repeat purchases without deep discounting</li>
            </ul>
            <p><strong>Impact:</strong> Could improve overall margins by 2-3%, adding $45,000-68,000 in annual profit</p>
            
            <h3>3. Loss-Making Orders Management</h3>
            <p><strong>Current Issue:</strong> 1,813 orders (18.1%) operate at a loss</p>
            <p><strong>Recommendations:</strong></p>
            <ul style="margin-left: 1rem; margin-bottom: 1rem;">
                <li style="margin-bottom: 0.5rem;">Implement break-even analysis for all orders before fulfillment</li>
                <li style="margin-bottom: 0.5rem;">Create pricing floors to prevent unprofitable transactions</li>
                <li style="margin-bottom: 0.5rem;">Review shipping costs for small/low-value orders</li>
            </ul>
            <p><strong>Impact:</strong> Eliminate losses on 50-75% of loss-making orders, recovering $27,000-41,000 annually</p>
            
            <h3>4. Category-Specific Strategies</h3>
            <p><strong>Technology (21.56% margin):</strong> Maintain premium positioning; increase market share aggressively</p>
            <p><strong>Office Supplies (17.23% margin):</strong> Boost repeat purchase programs; aim for 5-10% volume increase</p>
            <p><strong>Furniture (-2.39% margin):</strong> Priority rehabilitation needed; consider restructuring or outsourcing options</p>
            
            <h3>5. Regional Expansion</h3>
            <p><strong>Current Performance:</strong> Variable margins across regions (West 14.91% vs Central 7.93%)</p>
            <p><strong>Recommendations:</strong></p>
            <ul style="margin-left: 1rem; margin-bottom: 1rem;">
                <li style="margin-bottom: 0.5rem;">Analyze West region success factors and replicate in underperforming regions</li>
                <li style="margin-bottom: 0.5rem;">Allocate additional resources to high-potential Central and South markets</li>
                <li style="margin-bottom: 0.5rem;">Implement regional pricing strategies based on local market conditions</li>
            </ul>
            
            <h3>Expected Cumulative Impact</h3>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Conservative Profit Increase</div>
                    <div class="kpi-value">$100K+</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Profit Margin Improvement</div>
                    <div class="kpi-value">+2-4%</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Implementation Timeline</div>
                    <div class="kpi-value">6 Months</div>
                </div>
            </div>
        </section>
        
        <!-- Conclusion -->
        <section>
            <h2>🎯 Conclusion</h2>
            <p>The e-commerce superstore demonstrates solid revenue generation with $2.3M in annual sales, but significant profitability challenges exist. The most critical issues—Furniture category losses, excessive discounting, and loss-making orders—are addressable through targeted operational improvements.</p>
            
            <p>By implementing the recommended strategies, particularly focusing on category restructuring and discount optimization, the business can realistically achieve 15-25% profit improvement within 6 months. The strong customer loyalty (70% repeat rate) and diverse geographic presence provide a solid foundation for sustainable growth.</p>
            
            <p>Priority should be given to understanding and rehabilitating the Furniture category while simultaneously implementing discount and pricing discipline across all categories. These actions alone could add $100,000+ to annual profits.</p>
        </section>
    </main>
    
    <!-- Scroll to Top Button -->
    <button class="scroll-top" id="scrollTop">↑</button>
    
    <!-- Footer -->
    <footer>
        <p>&copy; 2026 E-Commerce Superstore Data Analysis. All rights reserved.</p>
        <p>Report generated on March 26, 2026 | Data period: January 3, 2014 - December 30, 2017</p>
    </footer>
    
    <script>
        // Scroll to top functionality
        const scrollTopBtn = document.getElementById('scrollTop');
        
        window.addEventListener('scroll', () => {
            if (window.pageYOffset > 300) {
                scrollTopBtn.classList.add('show');
            } else {
                scrollTopBtn.classList.remove('show');
            }
        });
        
        scrollTopBtn.addEventListener('click', () => {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
        
        // Smooth scrolling for navigation links
        document.querySelectorAll('a[href^="#"]').forEach(anchor => {
            anchor.addEventListener('click', function (e) {
                e.preventDefault();
                const target = document.querySelector(this.getAttribute('href'));
                if (target) {
                    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            });
        });
    </script>
</body>
</html>
"""
    
    return html_content

if __name__ == '__main__':
    print("🚀 Generating HTML Report...")
    print("=" * 60)
    
    html_content = generate_html()
    
    # Save HTML to file
    output_path = Path(__file__).parent / 'index.html'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    file_size = output_path.stat().st_size / (1024 * 1024)  # Convert to MB
    print(f"\n✅ HTML Report generated successfully!")
    print(f"📁 Output file: {output_path}")
    print(f"📊 File size: {file_size:.2f} MB")
    print("=" * 60)
