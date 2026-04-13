# Excel Formulas Quick Reference Guide
## For Retail Sales Dashboard Project

---

## 🎯 Core Formulas You'll Use

### **1. VLOOKUP - Lookup Product Details**

**Syntax:**
```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

**Example Use Case:** Find product price from product master list
```excel
=VLOOKUP(A2, Products!$A$2:$D$100, 3, FALSE)
```
- A2: Product name to find
- Products!$A$2:$D$100: Master table with products
- 3: Return value from 3rd column (price)
- FALSE: Exact match

**Common Errors:**
- #N/A: Lookup value not found → Use IFERROR
- #REF!: Column index exceeds table range → Check col_index_num

**Pro Tip:** Always use FALSE for exact match in sales data

---

### **2. INDEX-MATCH - Better Than VLOOKUP**

**Syntax:**
```excel
=INDEX(return_array, MATCH(lookup_value, lookup_array, 0))
```

**Example:** Find revenue for best-selling product
```excel
=INDEX(Revenue_Range, MATCH(MAX(Sales_Range), Sales_Range, 0))
```

**Why Better Than VLOOKUP:**
- Can look left (VLOOKUP can't)
- Faster for large datasets
- More flexible

**Common Pattern - Two-Way Lookup:**
```excel
=INDEX(Data_Range, MATCH(Row_Value, Row_Range, 0), MATCH(Col_Value, Col_Range, 0))
```

---

### **3. SUMIFS - Conditional Summing**

**Syntax:**
```excel
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2], [criteria2]...)
```

**Example 1:** Total revenue for North region in Q1
```excel
=SUMIFS(Net_Sales, Region, "North", Quarter, "Q1-2023")
```

**Example 2:** Revenue for Laptops with discount > 10%
```excel
=SUMIFS(Net_Sales, Product_Category, "Laptops", Discount_Percent, ">10")
```

**Example 3:** Sales in January 2023
```excel
=SUMIFS(Net_Sales, Order_Date, ">=1/1/2023", Order_Date, "<=1/31/2023")
```

**Comparison:**
- SUMIF: One condition
- SUMIFS: Multiple conditions (use this!)

---

### **4. COUNTIFS - Count with Conditions**

**Syntax:**
```excel
=COUNTIFS(criteria_range1, criteria1, [criteria_range2], [criteria2]...)
```

**Example 1:** Count orders in North region
```excel
=COUNTIFS(Region, "North")
```

**Example 2:** Count high-value orders (>₹50,000) in South region
```excel
=COUNTIFS(Net_Sales, ">50000", Region, "South")
```

**Example 3:** Count laptop sales with >15% discount
```excel
=COUNTIFS(Product_Category, "Laptops", Discount_Percent, ">15")
```

---

### **5. AVERAGEIFS - Conditional Averaging**

**Syntax:**
```excel
=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
```

**Example:** Average order value for smartphones in Q4
```excel
=AVERAGEIFS(Net_Sales, Product_Category, "Smartphones", Quarter, "Q4-2023")
```

---

### **6. TEXT - Format Dates & Numbers**

**Syntax:**
```excel
=TEXT(value, format_text)
```

**Common Date Formats:**
```excel
=TEXT(A2, "MMM-YYYY")          → "Jan-2023"
=TEXT(A2, "DD-MM-YYYY")        → "15-01-2023"
=TEXT(A2, "DDD")               → "Mon"
=TEXT(A2, "MMMM DD, YYYY")     → "January 15, 2023"
```

**Number Formats:**
```excel
=TEXT(12345.67, "₹#,##0.00")   → "₹12,345.67"
=TEXT(0.1234, "0.00%")         → "12.34%"
=TEXT(A2, "0.0") & " K"        → "12.3 K"
```

**Project Use:**
```excel
=TEXT(Order_Date, "MMM-YYYY")  → Extract month for grouping
=TEXT(Order_Date, "DDD")       → Day of week analysis
```

---

### **7. DATE Functions**

**Extract Components:**
```excel
=YEAR(A2)           → 2023
=MONTH(A2)          → 1
=DAY(A2)            → 15
=WEEKDAY(A2)        → 1 (Sunday=1, Monday=2...)
```

**Calculate Quarter:**
```excel
="Q" & ROUNDUP(MONTH(A2)/3, 0)  → "Q1"
="Q" & ROUNDUP(MONTH(A2)/3, 0) & "-" & YEAR(A2)  → "Q1-2023"
```

**Date Differences:**
```excel
=DATEDIF(Start_Date, End_Date, "D")     → Days
=DATEDIF(Start_Date, End_Date, "M")     → Months
=DATEDIF(Start_Date, End_Date, "Y")     → Years
```

**Today's Date:**
```excel
=TODAY()            → Current date
=NOW()              → Current date & time
```

---

### **8. IF Statements - Conditional Logic**

**Basic Syntax:**
```excel
=IF(logical_test, value_if_true, value_if_false)
```

**Example 1:** Revenue category
```excel
=IF(Net_Sales >= 100000, "High", "Low")
```

**Nested IF:**
```excel
=IF(Net_Sales >= 100000, "High", 
   IF(Net_Sales >= 50000, "Medium", "Low"))
```

**Example 2:** Discount tier
```excel
=IF(Discount_Percent >= 20, "Heavy", 
   IF(Discount_Percent >= 10, "Moderate", 
   IF(Discount_Percent > 0, "Light", "None")))
```

**With AND/OR:**
```excel
=IF(AND(Region="North", Net_Sales>50000), "Target Met", "Below Target")
=IF(OR(Product_Category="Laptops", Product_Category="Smartphones"), "Electronics", "Other")
```

---

### **9. IFERROR - Handle Errors Gracefully**

**Syntax:**
```excel
=IFERROR(value, value_if_error)
```

**Example 1:** VLOOKUP with error handling
```excel
=IFERROR(VLOOKUP(A2, Products!A:D, 3, FALSE), "Not Found")
```

**Example 2:** Division by zero
```excel
=IFERROR(Revenue/Orders, 0)
```

**Example 3:** Clean display
```excel
=IFERROR(B2/C2, "-")  → Shows "-" instead of #DIV/0!
```

---

### **10. CONCATENATE / TEXTJOIN - Combine Text**

**Simple Concatenation:**
```excel
=A2 & " " & B2                        → "Rahul Sharma"
=A2 & " - " & TEXT(B2, "DD-MMM")      → "ORD-001 - 15-Jan"
```

**CONCATENATE Function:**
```excel
=CONCATENATE(A2, " (", B2, ")")       → "Laptops (North)"
```

**TEXTJOIN (Excel 2019+):**
```excel
=TEXTJOIN(", ", TRUE, A2:A10)         → Joins with commas, ignores blanks
```

---

### **11. UNIQUE & COUNT Unique Values**

**Count Unique (Formula for all Excel versions):**
```excel
=SUMPRODUCT(1/COUNTIF(A2:A100, A2:A100))
```

**Count Unique with Criteria:**
```excel
=SUMPRODUCT((Region="North")/COUNTIFS(Customer_ID, Customer_ID, Region, "North"))
```

**UNIQUE Function (Excel 365):**
```excel
=UNIQUE(A2:A100)
```

---

### **12. Array Formulas - Advanced**

**Find Maximum Value with Criteria:**
```excel
{=MAX(IF(Region="North", Net_Sales))}
```
Press Ctrl+Shift+Enter to create array formula

**Sum Top 5 Values:**
```excel
=SUMPRODUCT(LARGE(Revenue_Range, ROW(1:5)))
```

**Dynamic Top Product:**
```excel
=INDEX(Product_Range, MATCH(LARGE(Revenue_Range, 1), Revenue_Range, 0))
```

---

## 🔢 Calculated Columns for Dashboard

### **Column M: Total Sales (Before Discount)**
```excel
=[@Quantity] * [@Unit_Price]
```
Or without table:
```excel
=I2 * J2
```

---

### **Column N: Discount Amount**
```excel
=[@Total_Sales] * ([@Discount_Percent]/100)
```
Or:
```excel
=M2 * (K2/100)
```

---

### **Column O: Net Sales (After Discount)**
```excel
=[@Total_Sales] - [@Discount_Amount]
```
Or:
```excel
=M2 - N2
```

Alternative single formula:
```excel
=I2 * J2 * (1 - K2/100)
```

---

### **Column P: Month**
```excel
=TEXT([@Order_Date], "MMM-YYYY")
```

---

### **Column Q: Quarter**
```excel
="Q" & ROUNDUP(MONTH([@Order_Date])/3, 0) & "-" & YEAR([@Order_Date])
```

---

### **Column R: Day of Week**
```excel
=TEXT([@Order_Date], "DDD")
```

---

### **Column S: Revenue Category**
```excel
=IF([@Net_Sales]>=100000, "High", IF([@Net_Sales]>=50000, "Medium", "Low"))
```

---

### **Column T: Performance Flag**
```excel
=IF(AND([@Net_Sales]>50000, [@Discount_Percent]<15), "Excellent", 
  IF([@Net_Sales]>50000, "Good", 
  IF([@Discount_Percent]<15, "Low Revenue", "High Discount")))
```

---

## 📊 Dashboard KPI Formulas

### **Total Revenue**
```excel
=SUM(Cleaned_Data[Net_Sales])
```

### **Total Orders**
```excel
=COUNTA(Cleaned_Data[Order_ID])
```

### **Average Order Value**
```excel
=AVERAGE(Cleaned_Data[Net_Sales])
```
Or:
```excel
=SUM(Cleaned_Data[Net_Sales]) / COUNTA(Cleaned_Data[Order_ID])
```

### **Total Customers (Unique Count)**
```excel
=SUMPRODUCT(1/COUNTIF(Cleaned_Data[Customer_ID], Cleaned_Data[Customer_ID]))
```

### **Total Discount Given**
```excel
=SUM(Cleaned_Data[Discount_Amount])
```

### **Discount as % of Revenue**
```excel
=SUM(Cleaned_Data[Discount_Amount]) / (SUM(Cleaned_Data[Net_Sales]) + SUM(Cleaned_Data[Discount_Amount]))
```
Format as percentage

### **Units Sold**
```excel
=SUM(Cleaned_Data[Quantity])
```

### **Average Discount %**
```excel
=AVERAGE(Cleaned_Data[Discount_Percent])
```

### **Repeat Customer Rate**
```excel
=(COUNTA(Cleaned_Data[Customer_ID]) - SUMPRODUCT(1/COUNTIF(Cleaned_Data[Customer_ID], Cleaned_Data[Customer_ID]))) / COUNTA(Cleaned_Data[Customer_ID])
```
This gives: (Total Orders - Unique Customers) / Total Orders

---

## 🎯 Regional Analysis Formulas

### **Revenue by Region**
```excel
=SUMIF(Cleaned_Data[Region], "North", Cleaned_Data[Net_Sales])
```

### **Orders by Region**
```excel
=COUNTIF(Cleaned_Data[Region], "North")
```

### **Average Order Value by Region**
```excel
=AVERAGEIF(Cleaned_Data[Region], "North", Cleaned_Data[Net_Sales])
```

### **Top Region (Dynamic)**
```excel
=INDEX(Region_List, MATCH(MAX(Region_Revenue), Region_Revenue, 0))
```

### **Region % Contribution**
```excel
=SUMIF(Cleaned_Data[Region], "North", Cleaned_Data[Net_Sales]) / SUM(Cleaned_Data[Net_Sales])
```
Format as percentage

---

## 📅 Time-Based Analysis

### **Monthly Revenue**
```excel
=SUMIFS(Cleaned_Data[Net_Sales], Cleaned_Data[Month], "Jan-2023")
```

### **Quarter Revenue**
```excel
=SUMIFS(Cleaned_Data[Net_Sales], Cleaned_Data[Quarter], "Q1-2023")
```

### **Growth Rate (MoM)**
```excel
=(Current_Month_Revenue - Previous_Month_Revenue) / Previous_Month_Revenue
```
Example:
```excel
=(B32 - B31) / B31
```

### **Year-to-Date Revenue**
```excel
=SUMIFS(Cleaned_Data[Net_Sales], Cleaned_Data[Order_Date], ">=1/1/2023", Cleaned_Data[Order_Date], "<="&TODAY())
```

### **Best Month (Dynamic)**
```excel
=INDEX(Month_List, MATCH(MAX(Month_Revenue), Month_Revenue, 0))
```

---

## 🏆 Top Performers

### **Top Product by Revenue**
```excel
=INDEX(Cleaned_Data[Product_Name], MATCH(MAX(Product_Revenue), Product_Revenue, 0))
```

### **Top Sales Rep**
```excel
=INDEX(Cleaned_Data[Sales_Rep], MATCH(MAX(Rep_Revenue), Rep_Revenue, 0))
```

### **Most Sold Product (by Quantity)**
```excel
=INDEX(Product_Range, MATCH(MAX(COUNTIF(Product_Range, Product_Range)), COUNTIF(Product_Range, Product_Range), 0))
```

---

## 🔍 Data Validation Formulas

### **Check for Blanks**
```excel
=COUNTBLANK(A2:L2)
```
If > 0, row has missing data

### **Check for Duplicates**
```excel
=COUNTIF($A$2:$A$1000, A2) > 1
```
TRUE = duplicate

### **Validate Date Range**
```excel
=AND(A2>=DATE(2023,1,1), A2<=DATE(2023,12,31))
```

### **Validate Discount Range**
```excel
=AND(K2>=0, K2<=50)
```

---

## 💡 Business Insights Formulas

### **Correlation: Discount vs Revenue**
```excel
=CORREL(Cleaned_Data[Discount_Percent], Cleaned_Data[Net_Sales])
```
Value between -1 and 1:
- Close to 1: Strong positive correlation
- Close to -1: Strong negative correlation
- Close to 0: No correlation

### **Revenue per Sales Rep**
```excel
=SUMIF(Cleaned_Data[Sales_Rep], "Rahul Sharma", Cleaned_Data[Net_Sales])
```

### **Conversion Rate (if you add leads data)**
```excel
=Orders / Leads
```

### **Customer Lifetime Value (simple)**
```excel
=Average_Order_Value * Purchase_Frequency
```

---

## 🎨 Conditional Formatting Formulas

### **Highlight High Revenue Rows**
Select data range → Conditional Formatting → New Rule → Use Formula:
```excel
=$O2 >= 100000
```
Format: Green fill

### **Highlight Heavy Discounts**
```excel
=$K2 > 20
```
Format: Red fill

### **Highlight Weekend Sales**
```excel
=OR(WEEKDAY($B2)=1, WEEKDAY($B2)=7)
```
Format: Yellow fill

### **Alternate Row Coloring**
```excel
=MOD(ROW(), 2) = 0
```
Format: Light gray fill

---

## ⚙️ Helper Formulas

### **Remove Duplicates (Mark for Deletion)**
```excel
=IF(COUNTIF($A$1:A1, A2)>0, "DELETE", "KEEP")
```

### **Clean Text (Remove Extra Spaces)**
```excel
=TRIM(A2)
```

### **Proper Case (First Letter Capital)**
```excel
=PROPER(A2)
```
"rahul sharma" → "Rahul Sharma"

### **Extract Numbers from Text**
```excel
=VALUE(LEFT(A2, FIND(" ", A2)-1))
```

---

## 🚀 Performance Tips

### **1. Use Tables Instead of Ranges**
```excel
=SUM(Sales_Data[Net_Sales])  ✅ Better
=SUM($O$2:$O$1000)           ❌ Harder to maintain
```

### **2. Absolute vs Relative References**
```excel
=$A$1    → Always cell A1 (locked row & column)
=$A1     → Column A locked, row changes
=A$1     → Row 1 locked, column changes
=A1      → Both change when copied
```

### **3. Named Ranges**
Instead of:
```excel
=SUM($O$2:$O$1000)
```
Define name "Revenue" for O2:O1000, then:
```excel
=SUM(Revenue)  → Much cleaner!
```

### **4. Avoid Volatile Functions in Large Files**
These recalculate every time Excel recalculates (slow):
- NOW()
- TODAY()
- RAND()
- OFFSET()
- INDIRECT()

Use sparingly or calculate once and convert to values.

---

## 📋 Common Error Messages

| Error | Meaning | Solution |
|-------|---------|----------|
| #DIV/0! | Division by zero | Use IFERROR or check denominator |
| #N/A | Value not found | Check VLOOKUP/MATCH lookup value |
| #REF! | Invalid cell reference | Check if you deleted referenced cells |
| #VALUE! | Wrong data type | Check if text in number formula |
| #NAME? | Excel doesn't recognize text | Check formula name spelling |
| #NUM! | Invalid numeric value | Check for negative in SQRT, etc. |
| ##### | Column too narrow | Widen column |

---

## 🎯 Formula Debugging Tips

### **1. Evaluate Formula Step-by-Step**
- Select cell with formula
- Formulas tab → Evaluate Formula
- Click "Evaluate" repeatedly to see each step

### **2. Use F9 to Test Parts**
- In formula bar, select part of formula
- Press F9 to see result
- Press Esc (don't press Enter or it becomes value!)

### **3. Trace Precedents/Dependents**
- Formulas tab → Trace Precedents (shows what feeds this cell)
- Trace Dependents (shows what uses this cell)

### **4. Show Formulas**
- Ctrl + ` (backtick) → Toggle show/hide formulas
- Or: Formulas tab → Show Formulas

---

## 📚 Practice Exercises

### **Exercise 1: Regional Analysis**
Create formula to answer:
- Which region has highest average order value?
- Which region gives most discounts?
- Which region has best revenue-to-discount ratio?

### **Exercise 2: Time Analysis**
- Which month had highest growth rate?
- Which quarter performed best?
- What's the average weekend vs weekday revenue?

### **Exercise 3: Product Analysis**
- Which product category has highest profit margin?
- Which product has best sell-through rate?
- Which product-region combination performs best?

---

## ✅ Formula Checklist for Dashboard

Before submitting your project, ensure you've used:

**Basic Calculations:**
- [ ] SUM
- [ ] AVERAGE
- [ ] COUNT / COUNTA

**Conditional:**
- [ ] SUMIFS
- [ ] COUNTIFS
- [ ] AVERAGEIFS
- [ ] IF (with nesting)

**Lookup:**
- [ ] VLOOKUP or INDEX-MATCH

**Text/Date:**
- [ ] TEXT
- [ ] Date functions (YEAR, MONTH, QUARTER)

**Error Handling:**
- [ ] IFERROR

**Advanced:**
- [ ] Array formula (at least one)
- [ ] SUMPRODUCT (for unique count)

---

## 🎓 Next Level Formulas (Optional)

### **XLOOKUP (Excel 365 only)**
```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```
Better than VLOOKUP - can search any direction, has built-in error handling.

### **FILTER (Excel 365 only)**
```excel
=FILTER(array, include, [if_empty])
```
Example: Filter all North region sales
```excel
=FILTER(Sales_Data, Sales_Data[Region]="North", "No results")
```

### **LET (Excel 365 only)**
Define variables in formulas:
```excel
=LET(revenue, SUM(Net_Sales), 
     orders, COUNTA(Order_ID),
     revenue/orders)
```

---

**END OF QUICK REFERENCE**

Print this guide and keep it handy while building your dashboard!

Good luck! 🚀
