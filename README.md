
# 🍫 Chocolate Sales Dashboard (Excel)

> **End-to-end Excel dashboard** built using Power Pivot, advanced formulas, dynamic arrays, and custom visual design.  
> A complete business intelligence solution for tracking chocolate sales across 6 countries, 4 teams, and 22 products.

---

## 📊 Project Overview

This dashboard provides a **360-degree view** of chocolate sales performance. It was built entirely in Excel (no Power BI, no external tools) to demonstrate advanced Excel capabilities including:

- **Power Pivot Data Modeling**
- **Dynamic Array Functions** (Excel 365)
- **Advanced Conditional Formatting**
- **Custom Number Formatting**
- **Sparklines & Mini Charts**
- **Interactive Slicers**
- **Named Ranges & Structured References**

### Business Context
A chocolate company sells 22 different products across 6 countries through 25 salespeople organized into 4 teams. The dashboard answers:
- How is overall business performing? (KPIs + MoM trends)
- Which countries are top performers? (Geography + Maps)
- Which team and people are driving sales? (Bubble chart + rankings)
- Which products are most profitable? (Category slicer + sparklines)

---

## 🗂️ Data Model (Star Schema)

The project uses a **Star Schema** with 1 fact table and 3 dimension tables, connected via Power Pivot relationships.

| Table | Type | Rows | Columns | Key Field |
|-------|------|------|---------|-----------|
| **`sales`** | Fact Table | 3,791 | Sales Person, Geography, Product, Date, Amount, Customers, Boxes | — |
| **`products`** | Dimension | 22 | Product, Category, Size, Cost per Box | Product |
| **`locations`** | Dimension | 6 | Geo, Region | Geo |
| **`people`** | Dimension | 25 | Sales Person, Team | Sales Person |

### Relationships
```
sales[Product]     ──→ products[Product]
sales[Geography]   ──→ locations[Geo]
sales[Sales Person] ──→ people[Sales Person]
```

### Profit Calculation Logic
```
Total Cost      = SUM(Boxes × Cost per Box)   [from products table via relationship]
Total Profit    = Total Amount − Total Cost
Profit %        = Total Profit / Total Amount
```

---

## 🧮 Formulas & Logic — Section by Section

### 1. Slicer & Dynamic Text (Row 2–4)

| Cell | Formula | Purpose |
|------|---------|---------|
| `I3` | `=MAX(sales[Date])` | Gets the latest date in the dataset (e.g., 44592 = Jan 2022) |
| `I4` | `=I3-28` | Calculates date 28 days ago for "Last 28 Days" trends |
| `E4` | `=COUNTA(C4:C6)` | Counts how many categories are selected in the slicer |
| `F4` | `=CHOOSE(E4, C4, "(multiple)", "(All)")` | Displays the slicer status: single name, "(multiple)", or "(All)" |

**Why CHOOSE?**  
Dashboard dynamically shows what the user selected. If 1 category → show its name. If 2+ → show "(multiple)". If all → show "(All)".

---

### 2. Monthly Summary Pivot (Row 6–17)

This section extracts monthly totals using **Power Pivot measures** or `SUMIFS`:

| Metric | Calculation |
|--------|-------------|
| Sum of Amount | `=SUM(sales[Amount])` grouped by Year-Month |
| Sum of Boxes | `=SUM(sales[Boxes])` grouped by Year-Month |
| Total shipments | `=COUNT(sales[Date])` per month |
| Total Cost | `=SUMPRODUCT(sales[Boxes], products[Cost per Box])` |
| Total Profit | `=Amount − Cost` |
| Profit % | `=Profit / Amount` |

> **Note:** Dates are Excel serial numbers (44197 ≈ Jan 2021, 44592 ≈ Jan 2022).

---

### 3. KPI Cards — Top Summary (Row 19–30)

Six KPI cards show overall business health with **Month-over-Month % change**:

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

---

### 4. Geography / Maps Section (Row 52–68)

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
- `country.images` = Named range containing 6 country map PNGs
- `P63` = Map ID# from sorted list
- Result: Each country gets its correct map icon dynamically

**Dashboard Display Format:**
```
New Zealand    $3.78 m    $1.27 m    [map icon]
India          $3.76 m    $1.14 m    [map icon]
UK             $3.59 m    $1.13 m    [map icon]
```

---

### 5. Team & People Performance (Row 80–100)

#### Bubble Chart Data

| Team | Sales | Profit | x | y | Size |
|------|-------|--------|---|---|------|
| Yummies | $7.89m | $5.47m | 1 | 2 | 7,887,334 |
| Delish | $6.14m | $4.23m | 2 | 2 | 6,136,837 |
| Jucies | $4.24m | $2.91m | 1 | 1 | 4,237,639 |
| Tempo | $3.44m | $2.41m | 2 | 1 | 3,439,912 |

**Why x, y coordinates?**  
To place 4 teams in a **2×2 grid** on a scatter/bubble chart:
```
        x=1          x=2
y=2   [Jucies]    [Tempo]
y=1   [Yummies]   [Delish]
```
Bubble **size** = Total Sales (bigger bubble = more sales).

#### Top Sales People Table

```excel
=SUMIFS(sales[Amount], sales[Sales Person], person_name, people[Team], selected_team)
```
Or via **Power Pivot DAX measure** with team slicer context.

**Sparkline Data (Last 28 Days):**  
Columns `44565` to `44592` contain daily sales per person for sparkline mini-charts.

---

### 6. Product Performance Table (Row 110–168)

This is the **most advanced section** of the dashboard.

#### Features:

| Feature | Implementation |
|---------|---------------|
| **Category Slicer** | `Slicer_Category` — filters Bars / Bites / Other / All |
| **Dynamic Sort** | `sort.option.Product.table` — sort by Sales, Boxes, Shipments, or Profit % |
| **Sort Order** | `-1` (Descending / Top to Bottom) |
| **Filter Logic** | `=FILTER(products, IF(category="All", TRUE, products[Category]=category))` |

#### Product Table Columns:

| Column | Formula / Source |
|--------|---------------|
| Product | Product name from `products` table |
| Sum of Amount | `=SUMIFS(sales[Amount], sales[Product], product_name)` |
| Sum of Boxes | `=SUMIFS(sales[Boxes], sales[Product], product_name)` |
| Total shipments | `=COUNTIFS(sales[Product], product_name)` |
| Profit % | `=(Amount − Cost) / Amount` |

#### "In Last 28 Days" Sparklines

- `starting.date.for.trend` = `I4` (28 days ago from latest date)
- Daily sales per product for last 28 days
- Feeds **Sparkline** charts in the dashboard

#### Conditional Formatting & Icons

| Element | Logic |
|---------|-------|
| Data Bars | `amount.bar.limit` sets max value for bar scaling |
| Profit Icon | `profit.icon.status` shows ▲/▼ arrows |
| Highlight | Selected category rows get highlighted color |

---

## 🎨 Dashboard Sheet — Visual Layer

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

### Conditional Formatting Used:
- 🟢🔴 **Icon Sets** (arrows) for MoM changes
- 📊 **Data Bars** for visual comparison in product table
- 🎨 **Color Scales** for profit margins
- ✨ **Highlight Rules** for selected categories

---

## 🖼️ Images Sheet — Design Assets

| Asset | Purpose |
|-------|---------|
| Country Maps | 6 map icons from [mapsicon](https://github.com/djaiss/mapsicon) |
| Color Palette | Background, Highlight, Chart colors, Text colors |
| Sample Tile | Dashboard tile design reference |

---

## 📋 Complete Formula Reference

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

---

## 🚀 How to Use This Dashboard

1. **Open the Excel file**
2. **Use Category Slicer** (top right) to filter Bars / Bites / Other / All
3. **Use Team Slicer** to focus on specific team performance
4. **Hover over KPIs** to see MoM % change
5. **Check Product Table** — sparklines show 28-day trends
6. **Click on Country Maps** section to see regional breakdown

---

## 🛠️ Technical Requirements

| Requirement | Version |
|-------------|---------|
| Microsoft Excel | 365 or Excel 2021+ |
| Power Pivot | Enabled (COM Add-in) |
| Dynamic Arrays | Required for SORT, FILTER, UNIQUE |
| OS | Windows 10/11 or macOS |

---

## 📁 File Structure

```
Product project.xlsx
├── Data sheet          ← Raw data (4 tables)
├── calculation sheet   ← All formulas & logic
├── Dashboard           ← Visual presentation layer
└── images              ← Design assets & color palette
```

---

## 🎯 Key Takeaways (Interview Ready)

> **Q: How is profit calculated?**  
> "Cost per Box comes from the Products table via Power Pivot relationship. In the Sales table, Boxes are multiplied by Cost per Box to get Total Cost. Amount minus Cost gives Profit, and Profit divided by Amount gives Profit Percentage."

> **Q: How do map images work dynamically?**  
> "Six country map PNGs are stored in a named range called `country.images`. The calculation sheet assigns a Map ID# (1–6) to each country based on sort order. The dashboard uses `INDEX(country.images, MapID)` to pull the correct image dynamically."

> **Q: How does the Category slicer work?**  
> `COUNTA` counts selected categories. `CHOOSE` displays 'All', '(multiple)', or the specific name. `FILTER` function updates the product table to show only the selected category's products."

> **Q: What are Sparklines showing?**  
> "The calculation sheet has daily sales data for the last 28 days (columns 44565–44592). Each product gets its own 28-day trend row. The dashboard's Sparkline charts reference this data to show mini line graphs."

> **Q: Why x,y coordinates for the bubble chart?**  
> "Excel scatter charts need numeric x and y values. I assigned (1,1), (1,2), (2,1), (2,2) to create a 2×2 grid layout for the 4 teams. Bubble size represents total sales volume."

---


| Dashboard View | Description |
|---------------|-------------|
[![Overall](.png)
| (https://1drv.ms/i/c/5c6c50a2fd2850a8/IQA_qXA4g69SRq9y3pCNkM9KARUTAS3PyCygllGOlyJCDXU?e=zA347h) | Full dashboard overview |

---

## 👤 Author

**Hira Nasir
Data Analyst | Excel & Power BI Enthusiast  
[[LinkedIn](your-linkedin-url)](https://www.linkedin.com/in/hira-nasir-448635a5/) • [[GitHub](your-github-url)](https://github.com/hiranasir022)

---

*Built with ❤️ in Microsoft Excel*  
*Dataset: Chocolate sales simulation for educational purposes*
