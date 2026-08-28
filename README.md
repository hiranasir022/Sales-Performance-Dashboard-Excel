
# 🍫 Chocolate Sales Dashboard (Excel and power BI)

> **End-to-end business intelligence solution** built using both **Microsoft Excel** and **Microsoft Power BI**.  
> A complete BI solution for tracking chocolate sales across 6 countries, 4 teams, and 22 products.


## 📊 Dashboard Previews

### Excel Dashboard
![Excel Dashboard](Product-sale-dashboard.png)

### Power BI — Sales Performance Report
![Power BI Sales Performance](Sales-performance-report.png)

### Power BI — Sales Trend & Forecast
![Power BI Sales Trend](Sales-trend-and-forcast.png)

---

## 📊 Project Overview

This project provides a **360-degree view** of chocolate sales performance, implemented in **two platforms**:

| Platform | Focus | Key Features |
|----------|-------|--------------|
| **Excel** | Advanced Excel capabilities, formulas, Power Pivot | Dynamic arrays, sparklines, custom formatting, slicers |
| **Power BI** | Interactive visualizations, DAX, data modeling | Rich visuals, forecasting, drill-through, bookmarks |

### Business Context
A chocolate company sells 22 different products across 6 countries through 25 salespeople organized into 4 teams. The dashboard answers:
- How is overall business performing? (KPIs + MoM trends)
- Which countries are top performers? (Geography + Maps)
- Which team and people are driving sales? (Bubble chart + rankings)
- Which products are most profitable? (Category slicer + sparklines)
- What are the future sales trends? (Forecasting)

---

## 🛠️ Excel Version

Built entirely in Excel (no external tools) to demonstrate advanced Excel capabilities including:

- **Power Pivot Data Modeling**
- **Dynamic Array Functions** (Excel 365)
- **Advanced Conditional Formatting**
- **Custom Number Formatting**
- **Sparklines & Mini Charts**
- **Interactive Slicers**
- **Named Ranges & Structured References**

### Data Model (Star Schema)

The project uses a **Star Schema** with 1 fact table and 3 dimension tables, connected via Power Pivot relationships.

| Table | Type | Rows | Columns | Key Field |
|-------|------|------|---------|-----------|
| **`sales`** | Fact Table | 3,791 | Sales Person, Geography, Product, Date, Amount, Customers, Boxes | — |
| **`products`** | Dimension | 22 | Product, Category, Size, Cost per Box | Product |
| **`locations`** | Dimension | 6 | Geo, Region | Geo |
| **`people`** | Dimension | 25 | Sales Person, Team | Sales Person |

### Relationships
```
sales[Product]      ──→ products[Product]
sales[Geography]    ──→ locations[Geo]
sales[Sales Person] ──→ people[Sales Person]
```

### Profit Calculation Logic
```
Total Cost      = SUM(Boxes × Cost per Box)   [from products table via relationship]
Total Profit    = Total Amount − Total Cost
Profit %        = Total Profit / Total Amount
```

### Excel Formulas & Logic — Section by Section

#### 1. Slicer & Dynamic Text (Row 2–4)

| Cell | Formula | Purpose |
|------|---------|---------|
| `I3` | `=MAX(sales[Date])` | Gets the latest date in the dataset |
| `I4` | `=I3-28` | Calculates date 28 days ago for "Last 28 Days" trends |
| `E4` | `=COUNTA(C4:C6)` | Counts how many categories are selected in the slicer |
| `F4` | `=CHOOSE(E4, C4, "(multiple)", "(All)")` | Displays the slicer status |

#### 2. Monthly Summary Pivot (Row 6–17)

| Metric | Calculation |
|--------|-------------|
| Sum of Amount | `=SUM(sales[Amount])` grouped by Year-Month |
| Sum of Boxes | `=SUM(sales[Boxes])` grouped by Year-Month |
| Total shipments | `=COUNT(sales[Date])` per month |
| Total Cost | `=SUMPRODUCT(sales[Boxes], products[Cost per Box])` |
| Total Profit | `=Amount − Cost` |
| Profit % | `=Profit / Amount` |

#### 3. KPI Cards — Top Summary (Row 19–30)

| KPI | Formula | Format Code |
|-----|---------|-------------|
| **Total Amount** | `=SUM(sales[Amount])` | `$#,,.00 "m"` → `$21.70 m` |
| **Boxes** | `=SUM(sales[Boxes])` | `#,#,.0 "k"` → `1,344.6 k` |
| **Shipments** | `=COUNTA(sales[Date])` | `#,##` → `3,791` |
| **Total Cost** | `=SUMPRODUCT(Boxes, Cost per Box)` | `$#,,.00 "m"` → `$6.68 m` |
| **Total Profit** | `=Amount − Cost` | `$#,,.00 "m"` → `$15.02 m` |
| **Profit %** | `=Profit / Amount` | `#0.0%` → `69.2%` |

**Month-over-Month (MoM) %:**
```excel
MoM % = (This Month − Previous Month) / Previous Month
```
- Positive = 🟢 Green arrow (business growing)
- Negative = 🔴 Red arrow (business declining)

#### 4. Geography / Maps Section (Row 52–68)

**6 countries** dynamically sorted by user-selected criteria:

| Column | Purpose |
|--------|---------|
| Row Labels | Country names (Australia, Canada, India, New Zealand, UK, USA) |
| Sum of Amount | Total sales per country |
| Profit % | Profit margin per country |
| Map ID# | Assigned number (1–6) for image lookup |

**Dynamic Sorting:**
- Named range `sort.option.countrylist` controls sort order
- User can sort by **Sales** or **Profit %**

**Map Images:**
```excel
=INDEX(country.images, calculation_sheet!P63)
```

#### 5. Team & People Performance (Row 80–100)

**Bubble Chart Data:**

| Team | Sales | Profit | x | y | Size |
|------|-------|--------|---|---|------|
| Yummies | $7.89m | $5.47m | 1 | 2 | 7,887,334 |
| Delish | $6.14m | $4.23m | 2 | 2 | 6,136,837 |
| Jucies | $4.24m | $2.91m | 1 | 1 | 4,237,639 |
| Tempo | $3.44m | $2.41m | 2 | 1 | 3,439,912 |

**Why x, y coordinates?**  
To place 4 teams in a **2×2 grid** on a scatter/bubble chart.

#### 6. Product Performance Table (Row 110–168)

| Feature | Implementation |
|---------|---------------|
| **Category Slicer** | `Slicer_Category` — filters Bars / Bites / Other / All |
| **Dynamic Sort** | `sort.option.Product.table` — sort by Sales, Boxes, Shipments, or Profit % |
| **Sort Order** | `-1` (Descending / Top to Bottom) |
| **Filter Logic** | `=FILTER(products, IF(category="All", TRUE, products[Category]=category))` |

---

## 📈 Power BI Version

The Power BI version elevates the Excel dashboard with richer interactivity, advanced DAX measures, and built-in forecasting capabilities.

### Power BI Features

| Feature | Description |
|---------|-------------|
| **Interactive Visuals** | Click any visual to cross-filter all others |
| **Drill-Through** | Right-click on a country/team to see detailed breakdown |
| **Bookmarks & Buttons** | Toggle between different dashboard views |
| **Forecasting** | Built-in trend analysis with confidence intervals |
| **Tooltips** | Rich hover cards with additional KPIs |
| **Mobile Layout** | Optimized phone view for on-the-go insights |

### Power BI Data Model

Same **Star Schema** as Excel, loaded into Power BI via:
- **Import Mode** (recommended for this dataset size)
- Relationships auto-detected or manually configured

### DAX Measures

```dax
// Total Sales
Total Amount = SUM(sales[Amount])

// Total Boxes
Total Boxes = SUM(sales[Boxes])

// Total Cost (using RELATED)
Total Cost = SUMX(sales, sales[Boxes] * RELATED(products[Cost per Box]))

// Total Profit
Total Profit = [Total Amount] - [Total Cost]

// Profit %
Profit % = DIVIDE([Total Profit], [Total Amount], 0)

// MoM Change
MoM % = 
VAR CurrentMonth = [Total Amount]
VAR PreviousMonth = CALCULATE([Total Amount], DATEADD(sales[Date], -1, MONTH))
RETURN DIVIDE(CurrentMonth - PreviousMonth, PreviousMonth, 0)

// Amount per Box
Amount per Box = DIVIDE([Total Amount], [Total Boxes], 0)
```

### Power BI Visuals

| Visual | Purpose | Interactivity |
|--------|---------|---------------|
| **KPI Cards** | Top-level metrics with trend indicators | Click to filter |
| **Bar Chart — Team Sales** | Compare 4 teams side-by-side | Cross-filter |
| **Bar Chart — Category Performance** | Bars vs Bites vs Other | Drill-down |
| **Bar Chart — Amount per Box by Geo** | Compare profitability by country | Tooltip details |
| **Line Chart — Sales Trend** | Daily sales over time | Forecast overlay |
| **Line Chart — Customer Trend** | Daily customer count | Forecast overlay |
| **Slicers** | Team & Category filters | Multi-select |
| **Table — Salesperson Detail** | Individual performance | Sortable |

### Forecasting in Power BI

The **Sales Trend & Forecast** page uses Power BI's built-in forecasting:

```
Forecast Settings:
├── Forecast length: 7 days
├── Confidence interval: 95%
├── Seasonality: Auto-detected
└── Ignore last: 0 days
```

> **Insight:** The forecast shows a potential dip in early February, suggesting the need for a promotional push.

---

## 🎨 Dashboard Sheets / Pages

### Excel Dashboard Sheet

The **Dashboard** sheet is purely a presentation layer. All calculations live in the **Calculation Sheet**.

| Component | Source |
|-----------|--------|
| 🔢 Top KPI Cards | Linked cells from Calculation Sheet |
| 🗺️ Country Sales List | Sorted geography table + map images |
| 🫧 Team Bubble Chart | x, y, size columns → Scatter chart |
| 👥 Top People Table | Team-filtered salespeople rankings |
| 📦 Product Table | Sorted & filtered product list |
| 📈 Sparklines | Last 28 days trend data |
| 🎛️ Slicers | Category + Team interactive filters |

### Power BI Report Pages

| Page | Focus | Key Visuals |
|------|-------|-------------|
| **Sales Performance** | Team & individual performance | Team bar chart, salesperson table, category chart, geo chart |
| **Sales Trend & Forecast** | Time-series analysis | Customer trend line, amount trend line, 7-day forecast |

---

## 🖼️ Visual Previews

### Excel Dashboard
| Dashboard View | Description |
|---------------|-------------|
| (https://1drv.ms/i/c/5c6c50a2fd2850a8/IQA_qXA4g69SRq9y3pCNkM9KARUTAS3PyCygllGOlyJCDXU?e=zA347h) | Full Excel dashboard overview |

### Power BI Reports
| Report Page | Description |
|-------------|-------------|
| **Sales Performance Report** | Team-wise sales, category performance, amount per box by geography, salesperson rankings |
| **Sales Trend & Forecast** | Daily customer and amount trends with 7-day forecast and confidence intervals |

---

## 📋 Complete Formula Reference

### Excel Formulas

| Formula | Where Used | Why |
|---------|-----------|-----|
| `MAX()` | Latest date | Dynamic "latest month" display |
| `CHOOSE()` | Slicer text | Switch between 1/multiple/All |
| `COUNTA()` | Slicer count | Count selected categories |
| `SUMIFS()` | KPIs, product table | Conditional summing |
| `SUMPRODUCT()` | Total Cost | Multiply Boxes × Cost per Box |
| `COUNTIFS()` | Shipments | Count transactions per product |
| `INDEX()` | Map images | Dynamic image lookup |
| `MATCH()` | Sorting | Find position in sorted list |
| `SORT()` | Product table | Dynamic array sorting (Excel 365) |
| `FILTER()` | Product table | Show only selected category |
| `IF()` | Slicer logic | Handle All vs specific category |
| `TEXT()` | Formatting | Custom number display |
| `UNIQUE()` | Dropdown lists | Remove duplicates |

### Power BI DAX

| Measure | Purpose |
|---------|---------|
| `SUM()` | Basic aggregation |
| `SUMX()` | Iterative calculation (Boxes × Cost) |
| `CALCULATE()` | Context modification |
| `DATEADD()` | Time intelligence (MoM, YoY) |
| `DIVIDE()` | Safe division (handles divide by zero) |
| `RELATED()` | Pull values from related dimension tables |

---

## 🚀 How to Use

### Excel Version
1. **Open the Excel file**
2. **Use Category Slicer** (top right) to filter Bars / Bites / Other / All
3. **Use Team Slicer** to focus on specific team performance
4. **Hover over KPIs** to see MoM % change
5. **Check Product Table** — sparklines show 28-day trends
6. **Click on Country Maps** section to see regional breakdown

### Power BI Version
1. **Open the `.pbix` file** in Power BI Desktop
2. **Use Team Slicer** (left panel) to filter by Delish / Jucies / Tempo / Yummies
3. **Click on any bar** in charts to cross-filter the entire page
4. **Hover over line charts** to see daily values and forecast ranges
5. **Right-click on a country** in visuals for drill-through options
6. **Toggle between pages** using navigation buttons

---

## 🛠️ Technical Requirements

### Excel
| Requirement | Version |
|-------------|---------|
| Microsoft Excel | 365 or Excel 2021+ |
| Power Pivot | Enabled (COM Add-in) |
| Dynamic Arrays | Required for SORT, FILTER, UNIQUE |
| OS | Windows 10/11 or macOS |

### Power BI
| Requirement | Version |
|-------------|---------|
| Power BI Desktop | Latest (monthly updates recommended) |
| Power BI Service | Optional (for sharing & scheduling) |
| DAX Knowledge | Intermediate |
| OS | Windows 10/11 |

---

## 📁 File Structure

```
Chocolate Sales Dashboard/
│
├── Excel/
│   ├── Product project.xlsx
│   │   ├── Data sheet          ← Raw data (4 tables)
│   │   ├── calculation sheet   ← All formulas & logic
│   │   ├── Dashboard           ← Visual presentation layer
│   │   └── images              ← Design assets & color palette
│   └── README.md
│
└── Power BI/
    ├── Chocolate Sales.pbix
    │   ├── Sales Performance Report    ← Page 1
    │   ├── Sales Trend & Forecast      ← Page 2
    │   └── Data Model                  ← Star schema + DAX
    └── README.md
```

---

## 🎯 Key Takeaways (Interview Ready)

> **Q: How is profit calculated?**  
> "Cost per Box comes from the Products table via relationship (Power Pivot in Excel, RELATED in Power BI). In the Sales table, Boxes are multiplied by Cost per Box to get Total Cost. Amount minus Cost gives Profit, and Profit divided by Amount gives Profit Percentage."

> **Q: How do map images work dynamically in Excel?**  
> "Six country map PNGs are stored in a named range called `country.images`. The calculation sheet assigns a Map ID# (1–6) to each country based on sort order. The dashboard uses `INDEX(country.images, MapID)` to pull the correct image dynamically."

> **Q: How does the Category slicer work in Excel?**  
> "`COUNTA` counts selected categories. `CHOOSE` displays 'All', '(multiple)', or the specific name. `FILTER` function updates the product table to show only the selected category's products."

> **Q: What are Sparklines showing?**  
> "The calculation sheet has daily sales data for the last 28 days. Each product gets its own 28-day trend row. The dashboard's Sparkline charts reference this data to show mini line graphs."

> **Q: Why x,y coordinates for the bubble chart?**  
> "Excel scatter charts need numeric x and y values. I assigned (1,1), (1,2), (2,1), (2,2) to create a 2×2 grid layout for the 4 teams. Bubble size represents total sales volume."

> **Q: What forecasting method does Power BI use?**  
> "Power BI uses exponential smoothing for forecasting. It analyzes historical patterns, detects seasonality automatically, and projects future values with confidence intervals. This helps identify potential dips or peaks in sales."

> **Q: What's the difference between Excel and Power BI for this project?**  
> "Excel showcases advanced formula skills, Power Pivot modeling, and dynamic arrays — great for spreadsheet-based analysis. Power BI adds interactive cross-filtering, DAX time intelligence, built-in forecasting, and richer visual storytelling — better for stakeholder presentations and scalable BI."

---

## 👤 Author

**Hira Nasir**  
Data Analyst | Excel & Power BI Enthusiast  
[LinkedIn](https://www.linkedin.com/in/hira-nasir-448635a5/) • [GitHub](https://github.com/hiranasir022)

---

*Built with ❤️ in Microsoft Excel & Power BI*  
*Dataset: Chocolate sales simulation for educational purposes*
