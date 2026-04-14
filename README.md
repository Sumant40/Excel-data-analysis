# Excel Analytics Project: Retail Sales Dashboard
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)
## Complete Step-by-Step Guide for Data Analyst Portfolio

---

## Project Overview

**Project Name:** Retail Sales Performance Dashboard and Analysis  
**Workbook File:** `calculated excel/work in progress.xlsx`  
**Source Data:** `output/retail_sales_data.csv`  
**Duration:** 3-4 hours (beginner) | 2-3 hours (intermediate)  
**Difficulty:** Beginner to Intermediate  
**Tools Required:** Microsoft Excel 2016 or later (Excel 365 recommended), Python (optional for regenerating sample data)

**What You'll Build:**
- Interactive sales dashboard
- Cleaned and enriched Excel data table
- KPI summary analysis
- Pivot-based category, regional, and rep analysis
- Business insights suitable for a portfolio project

**Skills Demonstrated:**
- Data cleaning and transformation
- Excel formulas (`TEXT`, `IF`, `SUMIFS`, structured references)
- Pivot Tables and dashboard reporting
- KPI design and business analysis
- Project documentation

---

## Learning Objectives

By completing this project, you will:
1. Import and clean transactional retail data in Excel.
2. Build formula-driven columns for sales, discounts, calendar grouping, and revenue bands.
3. Create KPI summaries and pivot views for category, region, and sales rep performance.
4. Validate workbook results against the source CSV.
5. Turn spreadsheet outputs into clear business recommendations and portfolio-ready documentation.

---

## Dataset Description

**Business Context:**  
You are analyzing sales for "TechGear Retail," a consumer electronics business operating across multiple regions in India. The goal is to understand revenue performance, discount impact, category contribution, and sales trends across 2023.

**Dataset Specifications:**
- Time period: January 1, 2023 to December 31, 2023
- Records: 1,000 transactions
- Raw columns: 12
- Cleaned columns: 19
- Unique customers: 40

**Raw Data Fields:**
1. `Order_ID` - Unique transaction identifier
2. `Order_Date` - Date of purchase
3. `Customer_ID` - Customer identifier
4. `Customer_Name` - Customer full name
5. `Region` - Geographic region
6. `City` - City name
7. `Product_Category` - Product category
8. `Product_Name` - Specific product
9. `Quantity` - Units sold
10. `Unit_Price` - Price per unit
11. `Discount_Percent` - Discount applied
12. `Sales_Rep` - Assigned sales representative

**Calculated Fields Added in `Cleaned_data`:**
- `Total_sales = Quantity * Unit_Price`
- `Discount_Amount = Total_sales * (Discount_Percent / 100)`
- `Net_Sales = Total_sales - Discount_Amount`
- `Month = TEXT(Order_Date,"MMM-YYYY")`
- `Quarter = "Q" & ROUNDUP(MONTH(Order_Date)/3,0) & "-2023"`
- `Day_of_Week = TEXT(Order_Date,"DDD")`
- `Revenue_Category = IF(Net_Sales>=100000,"High",IF(Net_Sales>=50000,"Medium","Low"))`

---

## Workbook Structure

The workbook currently contains these sheets:

1. `retail_sales_data` - Raw imported dataset
2. `Cleaned_data` - Cleaned table with calculated fields
3. `Summary_Analysis` - KPI summary metrics
4. `Pivot_Analysis` - Pivot-based analysis by category and region
5. `Dashboard` - Presentation layer for the Excel dashboard
6. `Business_Insights` - Written interpretation of the analysis

---

## Confirmed Results

The current workbook values in `Summary_Analysis` are:

- Total Revenue: `30,205,420`
- Total Orders: `1000`
- Average Order Value: `30,205.42`
- Total Discount Given: `3,029,480`
- Discount Percent of Gross Revenue: `9.12%`
- Unique Customers: `40`
- Units Sold: `1292`
- Average Discount Percent: `9.435%`

These totals match the source CSV when the calculated columns are recomputed from the raw data.

---

## Key Findings

- Laptops are the top revenue category with `11,074,910`, contributing about 37% of total net revenue.
- East is the highest revenue region with `6,483,340`.
- Priya Patel is the top sales representative by net revenue with `4,336,320`.
- August 2023 is the strongest month with `3,129,925` in net revenue.
- The `Business_Insights` sheet already captures an important category-mix insight: laptops generate a large share of revenue with a smaller share of orders, which creates both upside and category concentration risk.

---

## Step-by-Step Implementation

### Phase 1: Prepare the source data

1. Use `output/retail_sales_data.csv` as the base dataset.
2. If needed, regenerate the sample dataset with:

```bash
python generate_sales_data.py
```

3. Open Excel and import the CSV into the workbook.

### Phase 2: Build the cleaned data table

1. Copy the imported data into `Cleaned_data`.
2. Convert the range into an Excel table named `Cleaned_Data`.
3. Add the calculated fields listed above.
4. Format date and currency columns consistently.

### Phase 3: Create summary metrics

Use formulas such as:

```excel
=SUM(Cleaned_Data[Net_Sales])
=COUNTA(Cleaned_Data[Order_ID])
=SUM(Cleaned_Data[Discount_Amount])
=AVERAGE(Cleaned_Data[Discount_Percent])
```

Populate the `Summary_Analysis` sheet with KPI outputs for revenue, orders, discounts, customers, and units sold.

### Phase 4: Build pivot analysis

Create pivot tables from `Cleaned_Data` for:
- Revenue by `Product_Category`
- Regional performance by `Region`
- Sales rep performance by `Sales_Rep`
- Time-based trend analysis by `Month` or `Quarter`

Use these pivot tables to support the dashboard and validate headline findings.

### Phase 5: Create the dashboard and insights

1. Use the KPI cells and pivot outputs to populate the `Dashboard` sheet.
2. Add charts for category performance, regional contribution, and trends.
3. Write business takeaways in `Business_Insights`.

---

## Supporting Files

- `generate_sales_data.py` - Generates the sample retail dataset
- `output/retail_sales_data.csv` - Source dataset used by the workbook
- `excel_formulas_cheatsheet.md` - Formula reference for the Excel build
- `project_checklist_resume_template.md` - Checklist and resume-oriented phrasing ideas

---

## Suggested Next Improvements

- Add more written insights to the `Business_Insights` sheet
- Expand monthly trend commentary in the workbook
- Add screenshots of the dashboard for portfolio use
- Review the checklist file and complete any unfinished dashboard polish items

---

## Resume-Ready Project Summary

This project demonstrates the ability to clean raw sales data, create formula-based metrics, build pivot-driven Excel analysis, and communicate business findings clearly. It is suitable for entry-level data analyst, reporting analyst, and business analyst portfolios.
