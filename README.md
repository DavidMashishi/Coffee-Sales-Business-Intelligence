# Coffee Sales — Business Intelligence Dashboard

## 📊 Project Overview

This project presents a **Coffee Sales Analysis Dashboard** developed using Microsoft Excel. The objective was to transform raw transactional sales data into an interactive Business Intelligence dashboard that enables stakeholders to monitor sales performance, identify customer trends, compare coffee varieties, and support data-driven business decisions.

The dataset covers coffee sales from **2019 to 2022** and includes customer information, product details, order history, roast types, coffee varieties, package sizes, and loyalty card status.

The final deliverable is an interactive Microsoft Excel dashboard featuring:

- Slicers
- Timeline filtering
- Pivot Tables
- Pivot Charts
- Excel functions
- Feature engineering
- Data preparation and analysis

---

## 📑 Table of Contents

- [Project Overview](#-project-overview)
- [Business Task](#-business-task-ask)
- [Dataset](#-dataset)
- [Tools & Technologies](#-tools--technologies)
- [BI Development Process](#-bi-development-process)
- [Data Preparation](#-data-preparation)
- [Process & Feature Engineering](#-process--feature-engineering)
- [Business Analysis](#-analyse--business-analysis)
- [Dashboard Features](#-dashboard-features)
- [Key Insights](#-key-insights)
- [Business Impact](#-business-impact)
- [Recommendations](#-recommendations)
- [Conclusion](#-conclusion)

---

## 🎯 Business Task — Ask

The primary objective of the project was to answer several key business questions:

1. Which country generates the highest sales revenue?
2. Who are the top-performing customers?
3. Which coffee varieties generate the highest sales?
4. Which roast type performs best?
5. Does customer loyalty influence purchasing behaviour?
6. How have sales changed over time?

---

# 📂 Dataset

The project combines data from multiple worksheets:

- **Orders**
- **Customers**
- **Products**

### Key Data Fields

| Field | Description |
|---|---|
| **Order Date** | Date of customer transaction |
| **Customer Details** | Customer information |
| **Country** | Customer's country |
| **Coffee Type** | Coffee variety |
| **Package Size** | Product package size |
| **Price Per Unit** | Product unit price |
| **Quantity** | Quantity purchased |
| **Sales** | Sales value |
| **Loyalty Card Status** | Customer loyalty membership |

The dataset contains **13,013 records**, which were reviewed and confirmed as unique during the data preparation process.

---

# 🛠️ Tools & Technologies

| Tool / Feature | Purpose |
|---|---|
| **Microsoft Excel** | Data preparation, analysis and dashboard development |
| **Pivot Tables** | Business performance analysis |
| **Pivot Charts** | Data visualization |
| **Timeline Filter** | Time-based analysis |
| **Slicers** | Interactive filtering |
| **XLOOKUP** | Customer information retrieval |
| **INDEX + MATCH** | Product information retrieval |
| **IF Statements** | Data transformation |
| **Excel Tables** | Structured data management |
| **Conditional Formatting** | Visual identification of patterns |
| **Feature Engineering** | Improving analytical usability |

---

# 🔍 BI Development Process

The project followed a structured Business Intelligence workflow:

**Ask → Prepare → Process → Analyse → Share → Act**

---

## 1. Prepare — Data Preparation

The preparation stage focused on ensuring the dataset was accurate, complete, and suitable for analysis.

### Data Preparation Activities

- Reviewed the dataset for duplicate records
- Confirmed that all **13,013 records were unique**
- Verified appropriate data types
- Combined data from multiple worksheets into a single analytical table
- Standardized column formatting
- Prepared the data model for dashboard creation



---

# ⚙️ Process & Feature Engineering

Several Excel functions and transformations were used to enrich the dataset before analysis.

## Customer Information

Customer information was retrieved using **XLOOKUP**, including:

- Customer Name
- Email Address
- Country
- Loyalty Card Status

### Example — XLOOKUP

```excel
=IF(XLOOKUP(C5;customers!$A$1:$A$1001;customers!$C$1:$C$1001;;0)=0;"";XLOOKUP(C5;customers!$A$1:$A$1001;customers!$C$1:$C$1001;;0))
```

This formula was used to prevent unwanted zero values from appearing in the email column.

Another XLOOKUP was used to populate country information from the customer worksheet.

---

## Product Information

Product attributes were populated using a combination of **INDEX and MATCH**, including:

- Coffee Type
- Roast Type
- Package Size
- Unit Price

### Example — INDEX + MATCH

```excel
=INDEX(products!$A$1:$G$49;MATCH(orders!$D2;products!$A$1:$A$49;0);MATCH(orders!I$1;products!$A$1:$G$1;0))
```



---

# 🧮 Feature Engineering

Additional transformations were performed to improve data readability and reporting.

## Roast Type Standardization

| Original | Updated |
|---|---|
| Ara | Arabica |
| Exc | Excelsa |
| Lib | Liberica |
| Rob | Robusta |

## Package Size Formatting

| Original | Updated |
|---:|---|
| 0.2 | 0.2kg |
| 0.5 | 0.5kg |
| 1.0 | 1.0kg |
| 2.5 | 2.5kg |

### Additional Enhancements

- Converted the dataset into an Excel Table using `Ctrl + T`
- Assigned meaningful Table names
- Refreshed Pivot Tables after adding new fields
- Added the Loyalty Card field to enable dashboard filtering



---

# 📈 Analyse — Business Analysis

Pivot Tables were created to summarize business performance from multiple analytical perspectives.

### Sales Timeline

The analysis included:

- Sales over time
- Coffee Type comparison
- Monthly trends
- Yearly trends

### Sales by Country

Countries were ranked according to total sales.

### Top 5 Customers

Customers were ranked according to total sales generated.

---

# 📊 Dashboard Features

The interactive dashboard provides management with a consolidated view of sales performance.

### Key Visualizations

- **Total Sales Timeline**
- **Sales by Country**
- **Top 5 Customers**

### Interactive Filters

Users can dynamically filter the dashboard using:

- **Roast Type Slicer**
- **Package Size Slicer**
- **Loyalty Card Filter**
- **Timeline Filter**

The timeline enables users to filter sales by **month and year**.

---

# 💡 Key Insights

## 🌍 Country Performance

- The **United States** generated the highest sales revenue by a significant margin.
- **Ireland** ranked second.
- The **United Kingdom** generated the lowest sales among the three countries.

## ☕ Coffee Performance

- **Light Roast** products generated the highest overall revenue.
- **Arabica** and **Excelsa** were among the strongest-performing coffee varieties.

## 👥 Customer Analysis

- A small group of repeat customers contributed a significant share of total revenue.
- The **Top 5 customers** consistently outperformed the rest of the customer base.

## 📈 Sales Trends

- Sales fluctuated throughout the reporting period, with several seasonal spikes.
- Monitoring these spikes can support inventory planning and promotional campaign scheduling.

## 🎫 Loyalty Card Analysis

The dashboard suggests that purchases made by loyalty card holders did **not significantly outperform** purchases from customers without loyalty cards.

This indicates an opportunity to review and improve the effectiveness of the loyalty programme.

---

# 💼 Business Impact

The dashboard transforms raw sales data into actionable Business Intelligence, enabling stakeholders to make faster, data-driven decisions.

## 💰 Revenue Growth

- Identifies high-performing countries, allowing marketing and sales teams to focus investment on profitable markets.
- Highlights top-selling coffee products and roast types, supporting product optimization and promotional campaigns.

## 👥 Customer Relationship Management

- Identifies high-value customers to support loyalty and retention planning.
- Reveals customer purchasing patterns.

## ⚡ Operational Efficiency

- Reduces the time required to analyse sales performance through dynamic slicers and timeline filtering.
- Enables management to monitor sales trends without creating multiple manual reports.

## 📦 Inventory Planning

- Identifies popular coffee varieties and package sizes to support optimum stock levels.
- Supports demand planning through monthly and yearly sales trend analysis.

## 🎯 Strategic Decision-Making

The dashboard supports:

- Sales performance monitoring
- Customer behaviour evaluation
- Future sales planning
- Marketing strategy development using historical sales data



---

# 🚀 Recommendations — Act

Based on the analysis, the following actions are recommended:

### 1. Increase Marketing in the United States

The United States demonstrates the strongest demand and should remain a key focus for marketing and sales investment.

### 2. Promote Light Roast Products

Light Roast coffee consistently generates the highest sales and should receive continued marketing and promotional attention.

### 3. Target High-Value Customers

Develop targeted campaigns for repeat purchases from high-performing customers to strengthen retention and customer lifetime value.

### 4. Review the Loyalty Programme

The analysis indicates that loyalty card holders do not significantly outperform non-members.

The loyalty programme should therefore be reviewed and strengthened to improve customer engagement and retention.

### 5. Use Seasonal Trends for Planning

Historical seasonal sales patterns should be incorporated into:

- Inventory planning
- Promotional scheduling
- Demand forecasting

### 6. Explore UK Growth Opportunities

The United Kingdom recorded the lowest sales among the three countries. Additional analysis should be conducted to identify opportunities to increase demand and market penetration.



---

# 🖼️ Dashboard

![Coffee Sales](https://github.com/DavidMashishi/Coffee-Sales-Business-Intelligence/blob/dfbcdaed7b2e36e4180bc4b9d1bd9cdfc261bcd2/images/Coffee%20sales%20dashboard%20Image.png)

```markdown

```

---

# 🧠 Skills Demonstrated

- Microsoft Excel
- Data Cleaning
- Data Preparation
- Feature Engineering
- XLOOKUP
- INDEX + MATCH
- IF Statements
- Pivot Tables
- Pivot Charts
- Slicers
- Timeline Filtering
- Data Visualization
- Sales Analysis
- Customer Analysis
- Business Intelligence Reporting
- KPI Analysis
- Data-Driven Decision-Making

---

# 🏁 Conclusion

This project demonstrates how **Microsoft Excel can transform raw transactional data into an interactive Business Intelligence dashboard**.

By following the **Ask → Prepare → Process → Analyse → Share → Act** framework, the dashboard provides meaningful insights into product performance, customer behaviour, sales demand, and trends.

---

## 📌 Project Focus

**Business Intelligence | Data Analytics | Microsoft Excel | Sales Analytics | Customer Analytics | Dashboard Development**
