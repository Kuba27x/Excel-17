# 📊 Excel-17

![Status](https://img.shields.io/badge/status-active-brightgreen.svg)
![Excel](https://img.shields.io/badge/Microsoft-Excel-blue.svg)

## ✨ Project Description

**Excel-17** is your guide to mastering Pivot Tables in Microsoft Excel.  
Learn practical tips, step-by-step instructions, and see illustrated examples for:
- Pivot Tables
- Grouping
- Multi-level Pivot Tables
- Multiple Report Filter Fields
- Frequency Distribution
- Pivot Charts
- Slicers
- Calculated Field/Item
- GETPIVOTDATA function

> 📚 **Goal:** Help you analyze data efficiently and create insightful Excel reports—ideal for learners and power users!

---

## 📒 Table of Contents

- [Pivot Tables](#-pivot-tables)
- [Sort & Filter Pivot Table](#-sort--filter-pivot-table)
- [Change Summary Calculation](#-change-summary-calculation)
- [Two-dimensional & Multi-level Pivot Tables](#-two-dimensional--multi-level-pivot-tables)
- [Grouping Items & Dates](#-grouping-items--dates)
- [Multiple Value & Report Filter Fields](#-multiple-value--report-filter-fields)
- [Frequency Distribution](#-frequency-distribution)
- [Pivot Charts](#-pivot-charts)
- [Slicers](#-slicers)
- [Update Pivot Table](#-update-pivot-table)
- [Calculated Field/Item](#-calculated-fielditem)
- [GETPIVOTDATA Function](#-getpivotdata-function)
- [Screenshots](#-screenshots)
- [Requirements](#-requirements)
- [Author](#-author)

---

## 📊 Pivot Tables

Pivot tables are a powerful Excel feature for summarizing large data sets.

To create a pivot table:
1. Click any cell in your data set.
2. Go to the **Insert** tab → **Tables** group → **PivotTable**.
   *Excel auto-selects your data. Default is a new worksheet.*

![Pivot Table Wizard](Screenshots/Pivot1.png)

3. Click **OK**.

The PivotTable Fields pane opens.  
To summarize exported amounts by product:
- Drag **Product** to Rows
- Drag **Amount** to Values
- Drag **Country** to Filters

![Pivot Fields](Screenshots/Pivot2.png)

**Result:**

![Pivot Table Result](Screenshots/Pivot3.png)

---

## 🔍 Sort & Filter Pivot Table

**Sort:**
1. Click a cell in the *Sum of Amount* column.
2. Right-click → Sort → *Sort Largest to Smallest*.

![Sort Dialog](Screenshots/Pivot4.png)

**Filter:**
Click the filter drop-down and select *France*.

**Result:**

![Filtered Pivot](Screenshots/Pivot5.png)

---

## 🧮 Change Summary Calculation

1. Click any cell in *Sum of Amount*.
2. Right-click → Value Field Settings.
3. Choose calculation type (e.g., *Count*).

![Value Field Settings](Screenshots/Pivot6.png)

4. Click OK.

![Count Result](Screenshots/Pivot7.png)

> 16 out of the 28 orders to France were 'Apple' orders.

---

## 🔀 Two-dimensional & Multi-level Pivot Tables

**Two-dimensional:**
- Drag **Country** to Rows
- Drag **Product** to Columns
- Drag **Amount** to Values
- Drag **Category** to Filters

![2D Fields](Screenshots/Pivot8.png)

**Result:**

![2D Result](Screenshots/Pivot9.png)

**Multi-level:**
Drag more than one field to Rows (e.g., *Category* and *Country*).

![Multi-level Fields](Screenshots/Pivot15.png)

**Result:**

![Multi-level Result](Screenshots/Pivot16.png)

---

## 🔗 Grouping Items & Dates

**Group Products:**
- Select items, right-click → Group.

![Group Dialog](Screenshots/Pivot10.png)
![Group Result](Screenshots/Pivot11.png)
Click minus signs to collapse.

![Collapse Groups](Screenshots/Pivot12.png)

**Group Dates:**
Add *Date* to Rows.  
Right-click a date, Group → Quarters.

![Date Group Dialog](Screenshots/Pivot13.png)
![Date Group Result](Screenshots/Pivot14.png)

> Quarter 2 is the best!

---

## ➕ Multiple Value & Report Filter Fields

**Multiple Values:**
- Country to Rows
- Amount to Values (twice)

![Multi-Value Fields](Screenshots/Pivot17.png)
![Multi-Value Result](Screenshots/Pivot18.png)

To show % of Grand Total:
- Value Field Settings → Custom Name: *Percentage*
- Show Values As: % of Grand Total

![Show Values As](Screenshots/Pivot19.png)
![Percentage Result](Screenshots/Pivot20.png)

**Multiple Filters:**
- Order ID to Rows
- Amount to Values
- Country & Product to Filters

![Multiple Filters](Screenshots/Pivot21.png)

Select *United Kingdom* and *Broccoli* in filters.

**Result:**

![Multiple Filters Result](Screenshots/Pivot22.png)

---

## 📈 Frequency Distribution

Pivot tables can create frequency distributions (histograms).

- Amount to Rows
- Amount to Values

![Freq Fields](Screenshots/Pivot23.png)

Change Values to *Count*.

![Count Values](Screenshots/Pivot24.png)

Group amounts (e.g., by 1000):

![Group Amounts](Screenshots/Pivot25.png)

**Result:**

![Freq Result](Screenshots/Pivot26.png)

Create a pivot chart for visualization:

![Freq Chart](Screenshots/Pivot27.png)

---

## 📊 Pivot Charts

Pivot charts visualize pivot tables.

Use your 2D pivot table:
1. Click a cell in the pivot table
2. PivotTable Analyze → Tools → PivotChart → OK

![Pivot Chart](Screenshots/Pivot28.png)

> Changes sync between chart and table!

**Filter Pivot Chart:**
Use Country filter to show only US exports.

![Country Filter](Screenshots/Pivot29.png)

Use Category filter to show vegetables.

![Category Filter](Screenshots/Pivot30.png)

**Change Chart Type:**
Select chart → Design → Change Chart Type → Pie → OK

![Pie Chart](Screenshots/Pivot31.png)
![Pie Result](Screenshots/Pivot32.png)

> Pie charts show one data series.

---

## 🥒 Slicers

Slicers make filtering pivot tables intuitive.

Insert a slicer:
1. Click a cell in your pivot table
2. PivotTable Analyze → Filters → Insert Slicer
3. Check *Country* → OK

![Country Slicer](Screenshots/Pivot33.png)

Click *United States* for product details.

![Slicer Filtered](Screenshots/Pivot34.png)

Insert a *Product* slicer, select style, use CTRL to select multiple.

![Product Slicer](Screenshots/Pivot35.png)

**Connect slicers to multiple tables:**
- Insert a second pivot table
- Select slicer → Slicer tab → Report Connections → select table

![Report Connections](Screenshots/Pivot36.png)
![Connected Slicers](Screenshots/Pivot37.png)

Click the icon in a slicer to clear its filter:

![Clear Slicer](Screenshots/Pivot38.png)

> No beans or carrots were exported to Canada!

---

## 🔄 Update Pivot Table

Pivot tables do **not** auto-refresh.  
To update after changing data:
1. Click any cell in the pivot table
2. Right-click → Refresh

> Or set "Refresh data when opening file" in PivotTable Options.

---

## 🧮 Calculated Field/Item

**Calculated Field:**  
Uses values from other fields.

1. Click a cell in the pivot table
2. PivotTable Analyze → Calculations → Fields, Items & Sets → Calculated Field

![Calc Field](Screenshots/Pivot39.png)

Insert formula, e.g.,
```
=IF(Amount>100000, 3%*Amount, 0)
```
Excel adds the field to Values.

![Calc Field Result](Screenshots/Pivot41.png)

**Calculated Item:**  
Uses values from other items.

1. Click a Country in the pivot table
2. PivotTable Analyze → Calculations → Fields, Items & Sets → Calculated Item

Insert formula, e.g.,
```
=3%*(Australia+'New Zealand')
```

![Calc Item Dialog](Screenshots/Pivot42.png)
![Calc Item Result](Screenshots/Pivot43.png)

> Created groups: Sales and Taxes.

---

## 📑 GETPIVOTDATA Function

**GETPIVOTDATA** returns visible data from a PivotTable.

## 📚 Official Documentation

- [GETPIVOTDATA function — Microsoft Docs](https://support.microsoft.com/en-us/office/getpivotdata-function-8c083b99-a922-4ca0-af5e-3af55960761f)

---

## 📷 Screenshots

All screenshots are in the `/Screenshots` folder.

---

## ℹ️ Requirements

- Microsoft Excel (recommended: 2021/365)
- Windows OS

---

## 👨‍💻 Author

Project and documentation by **Kuba27x**  
Repository: [Kuba27x/Excel-17](https://github.com/Kuba27x/Excel-17)

---
