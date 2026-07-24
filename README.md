# 🍫 Product Sales Performance Dashboard (Excel)

An advanced, single-page interactive Excel sales dashboard covering **overall business summary, geographical performance, team & people performance, and product performance** — built on a relational data model (Power Pivot style) with dynamic sorting, KPI trend cards, in-cell mini-charts, and map visuals. Built to demonstrate advanced Excel / BI skills for Data Analyst roles.

---

## 📌 Project Overview

The workbook analyzes **3,791 sales transactions** (Jan 2021 – Jan 2022, ~$21.7M in sales) for a chocolate/confectionery business sold across 6 countries by 4 sales teams. Instead of one flat table, the data is split into **4 related tables** (Sales, Products, Locations, People) linked together like a mini database — the same design principle used in Power Pivot / Power BI data models — and the dashboard lets the viewer resort and filter the view without touching a single chart manually.

---

## 🗂️ Workbook Structure

| Sheet | Purpose |
|---|---|
| **Data sheet** | 4 linked Excel Tables: `sales` (transactions), `products`, `locations`, `people` |
| **calculation sheet** | The engine — 200+ formulas computing every KPI, sparkline series, sort order, and map lookup used on the dashboard |
| **Dashboard** | The final interactive report: KPI cards, country map tiles, team/people leaderboard, and product performance table |
| **images** | Country map icons (used as visual tiles per country) + the dashboard's color/style palette |

**Relational data model** (like a star schema):
- `sales` — Sales Person, Geography, Product, Date, Amount, Customers, Boxes (3,791 rows)
- `products` — Product, Category, Size, Cost per Box (22 products across Bars / Bites / Other)
- `locations` — Geo, Region (6 countries mapped to APAC / Americas / Europe)
- `people` — Sales person, Team (26 people across 4 teams: Yummies, Delish, Jucies, Tempo)

---

## 🧮 Formulas Used

### 1. Latest date & rolling trend window
```excel
=MAX(sales[Date])
=I3-28
```
`MAX()` finds the most recent transaction date in the whole `sales` Table. Subtracting 28 pulls a "4 weeks ago" date — used to define the rolling trend window shown in the KPI sparklines, so the dashboard always trends the most recent month regardless of when it's opened.

### 2. Slicer selection label — `COUNTA` + `CHOOSE`
```excel
=COUNTA(C4:C6)
=CHOOSE(E4,C4,"(multiple)","(All)")
```
`COUNTA` counts how many category values are currently selected in the slicer. `CHOOSE` then picks what label to display: if exactly 1 is selected → show that category's name; if 2 → show "(multiple)"; if all 3 → show "(All)". This is how the dashboard's title bar dynamically shows what you're currently filtering on.

### 3. Month-over-Month % change
```excel
=D30/E30-1
```
Classic percentage-change formula: (this period ÷ previous period) − 1. Used across every KPI card (Sales, Boxes, Shipments, Cost, Profit, Profit %) to show the up/down trend arrow.

### 4. Conditional text formatting — `TEXT` + `IF`
```excel
=TEXT(H30,$K30)
=IF($E$4=3,"",TEXT(I30,$K30))
```
`TEXT(value, format_code)` renders a raw number using a custom number format stored in another cell (so currency, %, or plain-number formatting can be swapped centrally). The `IF($E$4=3,"",...)` hides the "selected category" figure entirely when "(All)" categories are chosen, since a selected-vs-all comparison wouldn't make sense.

### 5. `XMATCH` — modern lookup for sorting & map selection
```excel
=XMATCH(K56, sort.option.countrylist)
=XMATCH($K$56, country.names)
```
`XMATCH` finds the position of a value inside a list — used here (a) to find where the user's chosen sort option sits in a list of options, and (b) to find which row a selected country occupies in the country list, so the correct map image can be pulled for it.

### 6. `INDEX` — pulling the matching map image
```excel
map.1 = INDEX(country.images, 'calculation sheet'!$P$63)
```
Once `XMATCH` finds *which* country is selected, `INDEX` returns the image stored at that position from a named range of country map icons — this is how the geography section swaps in the correct country map automatically.

### 7. Dynamic labels with `&` and `UPPER`
```excel
="COUNTRY " & UPPER(K56)
```
Builds a dynamic title like "COUNTRY UK" by joining static text with the selected country name, forced to uppercase for consistent styling.

### 8. Highlighting the selected series — `IF` + `NA()`
```excel
=IF(G84=$H$89, L84, NA())
```
Compares each row's label against the one currently selected (e.g., selected team). If it matches, the real value is returned; otherwise `NA()` is returned, which Excel's in-cell mini-charts/conditional formatting simply skip — this is how the dashboard "highlights" just the selected team or product without extra chart series.

### 9. Building a multi-line data-label — `TEXTJOIN` + `CHAR(10)`
```excel
=TEXTJOIN(CHAR(10), , G84, H92, I92)
```
Joins several values with a line-break character (`CHAR(10)`) between them instead of a normal separator — producing a stacked, multi-line label inside a single cell (name on line 1, value on line 2, etc.), a common trick for compact KPI tiles.

### 10. Building continuous date sequences for trend charts
```excel
=starting.date.for.trend+1
(then each next cell = the cell before it + 1)
```
Starting from a named "trend start date," each subsequent cell just adds one day — generating a clean, continuous date axis for the sparkline/trend visuals across many columns.

### 11. `MROUND` — clean axis scaling
```excel
=MROUND(MAX(J100:J112)*2, 100000)
```
Finds the largest value in a range, doubles it (to leave headroom), then rounds it to the nearest 100,000 using `MROUND` — used to set a clean, round maximum scale for the in-cell bar/gauge visuals so bars never look cramped or overflow.

### 12. Counting selections — `COUNTIFS`
```excel
=COUNTIFS($C$4:$C$7, C135)
```
Counts how many times a category appears within the current slicer selection range — used to check whether a given category/team/product is part of what's currently filtered, driving highlight logic elsewhere.

### 13. Toggle logic — `IF` returning ±1
```excel
=IF(K124=1, 1, -1)
=IF(K153="Product", 1, -1)
```
Converts a selection into a `+1`/`-1` flag — commonly used to flip a sort order (ascending vs. descending) or a chart's direction based on what the user picked from a dropdown, without needing separate formulas for each direction.

### 14. Building sortable/labelled lists dynamically
```excel
=C135 & IF(H135, $J$133, "")
```
Appends a marker (e.g., an arrow or highlight symbol) onto a label only when a condition is true (`H135` is truthy) — used to flag the currently-active sort column in list headers.

---

## 📊 Dashboard Components

The **Dashboard** sheet is organized into 4 sections, matching the brief noted in the workbook itself:

1. **Overall Business Summary** — KPI cards for Sales, Boxes, Shipments, Cost, Profit, and Profit % with month-over-month trend indicators and mini sparkline trends
2. **Geographical Performance** — country tiles (Australia, Canada, India, New Zealand, UK, USA) with map icons that update based on the country selected, sorted by Sales or Profit
3. **Team & People Performance** — a sortable leaderboard of sales teams (Yummies, Delish, Jucies, Tempo) and individual salespeople, with the selected team highlighted
4. **Product Performance** — product-level breakdown by category (Bars, Bites, Other), sortable and filterable by category slicer, with a "Profit Indicator" toggle

Every section responds to the shared slicers/selection cells, so choosing a category, team, or sort order updates KPI cards, highlighted bars, map tiles, and labels together — without needing separate charts per filter combination.

---

## 📈 Key Dataset Stats

- **3,791 transactions**, Jan 2021 – Jan 2022, across **6 countries** (Australia, Canada, India, New Zealand, UK, USA)
- **Total sales: ~$21.7M** | Total profit: ~$15.0M (≈69% profit margin)
- **22 products** across 3 categories: Bars (11), Bites (7), Other (4)
- **26 salespeople** across **4 teams**: Yummies (9), Delish (7), Jucies (5), Tempo (4)
- Countries roll up into 3 regions: APAC (Australia, India, New Zealand), Americas (USA, Canada), Europe (UK)

---

## 🛠️ Tools & Skills Used

- Microsoft Excel
- Relational data modeling (multiple linked Tables, star-schema style)
- Advanced formulas: `XMATCH`, `INDEX`, `MROUND`, `TEXTJOIN`, `COUNTIFS`, `CHOOSE`, `COUNTA`, nested `IF`, dynamic named ranges
- KPI cards with month-over-month % trend logic
- In-cell mini-visuals (sparklines, conditional-formatting bars) instead of standard charts
- Dynamic image/map lookups tied to slicer selections
- Custom number formats & centralized style/color palette
- Interactive, formula-driven filtering (no VBA)

---

## 📂 How to Use

1. Download/clone this repository.
2. Open `Product_project.xlsx` in Microsoft Excel.
3. Go to the **Dashboard** sheet to view the final report.
4. Use the slicers/dropdowns (Category, Team, Sort by) to see the KPI cards, highlighted bars, and map tiles update live.
5. Explore the **calculation sheet** to see how each visual and KPI is built from the underlying `sales`, `products`, `locations`, and `people` Tables.

---

## 👤 Author

**[Hira Nasir]**
 🔗 https://www.linkedin.com/in/hira-nasir-448635a5

*If you found this project useful, feel free to ⭐ star the repo!*
