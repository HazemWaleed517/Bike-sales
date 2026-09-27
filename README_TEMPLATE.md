# Bike Sales Performance Analysis 
Q4 2021 (Excel PivotTable Dashboard)

> Analyzed 752 Q4 2021 bicycle sales transactions across six countries using Excel PivotTables and PivotCharts to uncover the revenue, profitability, and customer drivers behind a $1.41M quarter, delivered as an interactive dashboard plus a structured business-question framework.

---

## Project Type

- [x] Exploratory Data Analysis (EDA)
- [x] Dashboard / Data Visualization
- [x] Data Cleaning / Wrangling

---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Objectives](#2-objectives)
3. [Project Scope & Tools](#3-project-scope--tools)
4. [Repository Structure](#4-repository-structure)
5. [Data Workflow](#5-data-workflow)
6. [Data Model & Schema](#6-data-model--schema)
7. [Analysis & Metrics](#7-analysis--metrics)
8. [Key Insights](#8-key-insights)
9. [Recommendations](#9-recommendations)
10. [Assumptions & Limitations](#10-assumptions--limitations)
11. [Future Enhancements](#11-future-enhancements)
12. [Deliverables](#12-deliverables)
13. [Author](#13-author)

---

## 1. Project Overview

**Context:** A raw bicycle sales export covering three months of transactions (October–December 2021) across six countries arrived as a flat, un-enriched log — no revenue, profit, or age-segment fields, just unit prices, costs, and quantities.

**Problem Statement:** Which product category, country, and customer segment actually drove Q4 revenue and profit, and were profit margins healthy and consistent across the business — or hiding weak spots?

**Approach:** Built a structured analytical-questions framework spanning six themes (overall performance, customer behavior, profitability, seasonality, geography, and data quality), enriched the raw data with calculated fields (Revenue, Total_Cost, Profit, Profit_Margin_%, Age_Group), and answered the framework using an interactive Excel PivotTable/PivotChart dashboard.

**Outcome:** The quarter closed at **$1,409,500** in revenue and **$513,460** in profit (a **36.5%** average margin) across 752 transactions and 1,483 units sold. December alone accounted for 44% of quarterly revenue, Road Bikes outsold Mountain Bikes roughly 3-to-1, and Australia was the single largest market.

---

## 2. Objectives

- **Primary Objective:** Determine which product category, country, and customer segment drove Q4 2021 revenue and profit.
- **Secondary Objective 1:** Quantify the monthly (seasonal) revenue and profit trend across October–December 2021.
- **Secondary Objective 2:** Identify differences in profitability (Profit_Margin_%) across product categories, colors, and countries.
- **Secondary Objective 3:** Build a reusable, slicer-driven Excel PivotTable dashboard for interactive exploration of the data.

---

## 3. Project Scope & Tools

### Scope

| Dimension | Details |
|-----------|---------|
| **In Scope** | Transaction-level bike sales data — country, state, sub-category, color, product, quantity, unit cost, and unit price — for six countries: Australia, Canada, France, Germany, United Kingdom, and United States. |
| **Out of Scope** | Customer identity / repeat-purchase behavior, marketing spend, and store-level or channel data — none of these existed in the source export. |
| **Time Period** | October 1, 2021 – December 31, 2021 (Q4 2021) |
| **Granularity** | Row-level — one row per sales transaction (752 rows total) |

### Tools & Technologies

| Category | Tool(s) Used |
|----------|-------------|
| Data Storage | Excel workbooks (.xlsx) |
| Data Processing | Excel (calculated columns, formulas) |
| Analysis | Excel PivotTables |
| Visualization | Excel PivotCharts (bar, line, doughnut, 3D pie) + slicer |
| Documentation | Markdown (this README) + Word (analytical-questions framework) |

---

## 4. Repository Structure

```
bike-sales-analysis/
│
├── data/
│   ├── raw/                  # raw_data.xlsx — original, unmodified export
│   └── processed/            # Processed_data_.xlsx — enriched dataset
│
├── reports/
│   └── Professional_Analytical_Questions_Report.docx   # 20+ structured business questions
│
├── visuals/
│   └── Dashboard.xlsx         # Interactive PivotTable dashboard (10 charts + slicer)
│
└── README.md                  # You are here
```

---

## 5. Data Workflow

```
[Raw sales export]
      ↓
[Opened directly in Excel]
      ↓
[Added calculated fields: Revenue, Total_Cost, Profit, Profit_Margin_%, Age_Group]
      ↓
[PivotTables segmented by Country, Sub_Category, Bike_Color, Gender, Age_Group, Month]
      ↓
[Interactive dashboard: 10 PivotCharts + slicer]
```

1. **Source:** `raw_data.xlsx` — a single "Sales Data" sheet of 752 rows and 11 columns (Date, Customer_Age, Customer_Gender, Country, State, Sub_Category, Bike_Color, Product, Order_Quantity, Unit_Cost, Unit_Price), plus an "Info" sheet documenting each column.
2. **Ingestion:** Opened and worked on directly in Excel — no external database or scripting layer.
3. **Cleaning:** Reviewed for consistency (no duplicate transaction keys or missing core fields found in the 752-row export).
4. **Transformation:** Added five calculated columns to produce `Processed_data_.xlsx`:
   - `Revenue` = Unit_Price × Order_Quantity
   - `Total_Cost` = Unit_Cost × Order_Quantity
   - `Profit` = Revenue − Total_Cost
   - `Profit_Margin_%` = Profit ÷ Revenue
   - `Age_Group` = Customer_Age binned into Under 25 / 25–34 / 35–44 / 45–54 / 55+
5. **Analysis:** PivotTables aggregating Revenue, Profit, Total_Cost, and Order_Quantity by month, country, gender, age group, sub-category, and bike color.
6. **Output:** An interactive Excel dashboard (10 PivotCharts + a slicer) and a structured Word document framing the analysis as 20+ business questions across six themes.

---

## 6. Data Model & Schema

### Dataset: `Processed_data_.xlsx`

| Field Name | Data Type | Description | Example Value |
|------------|-----------|--------------|---------------|
| `Date` | date | Date of the sale | 2021-10-01 |
| `Customer_Age` | int | Customer's age at time of purchase | 17 |
| `Customer_Gender` | string | Customer gender (M/F) | M |
| `Country` | string | Country where the sale took place | Australia |
| `State` | string | State/region within the country | Victoria |
| `Sub_Category` | string | Product sub-category | Road Bikes |
| `Bike_Color` | string | Bike color | Black |
| `Product` | string | Product name and size | Road-750 Black, 48 |
| `Order_Quantity` | int | Units sold in the transaction | 3 |
| `Unit_Cost` | float | Company's cost per unit | 700 |
| `Unit_Price` | float | Customer-facing price per unit | 1100 |
| `Revenue` | float | Unit_Price × Order_Quantity | 3300 |
| `Total_Cost` | float | Unit_Cost × Order_Quantity | 2100 |
| `Profit` | float | Revenue − Total_Cost | 1200 |
| `Profit_Margin_%` | float | Profit ÷ Revenue | 0.364 |
| `Age_Group` | string | Customer_Age binned | Under 25 |

> **Row count:** 752 transactions (1,483 units sold)
> **Date range:** 2021-10-01 – 2021-12-31

---

## 7. Analysis & Metrics

### Analytical Approach

Exploratory — the dataset was interrogated through a self-built framework of 20+ questions across six themes (overall performance, customer behavior, profitability, seasonality, geography/market mix, and data quality), then answered through PivotTables and PivotCharts rather than a fixed hypothesis test.

### Key Metrics Defined

| Metric | Plain-Language Definition | Why It Matters |
|--------|---------------------------|-----------------|
| `Revenue` | Order_Quantity × Unit_Price | Top-line measure of sales activity |
| `Profit` | Revenue − Total_Cost | What the business actually keeps per sale |
| `Profit_Margin_%` | Profit ÷ Revenue | Whether growth is coming at healthy or thin margins |
| `Average Order Value` | Revenue ÷ number of transactions | Whether growth comes from more orders or bigger ones |

### Methods Used

- Descriptive aggregation (sum, average, count) via PivotTables
- Monthly trend analysis (Oct–Dec 2021)
- Segmentation by country, gender, age group, sub-category, and bike color
- Cross-tabulation (e.g., color × country, category × age group)
- Interactive filtering via a slicer across all charts

---

## 8. Key Insights

**Insight 1: December carried the quarter**
Revenue rose from $429,770 (Oct) to $352,760 (Nov) to $626,970 (Dec) — December alone made up 44% of the quarter. The dip in November followed by a sharp December spike points to a clear holiday-driven seasonal pattern rather than steady organic growth.

**Insight 2: Road Bikes dominate revenue, but the margin gap between categories is small**
Road Bikes generated $1,029,110 (73%) of revenue versus $380,390 for Mountain Bikes, yet their average margins are close (36.8% vs 35.8%). Category strategy should be judged mainly on volume and mix, not on a meaningful margin advantage.

**Insight 3: Australia is a disproportionately large market**
Australia alone produced $511,050 (36%) of quarterly revenue — more than the next two countries (US $331,080, France $175,860) combined. Canada, at $71,100, is the smallest market by a wide margin.

**Insight 4: The 25–34 age group is the real revenue engine, not gender**
Male and female customers split revenue almost evenly ($699,830 vs $709,670), so gender is not a meaningful lever. The 25–34 age bracket, however, generated $512,840 (36%) of revenue — clearly ahead of every other age group, including 35–44 ($423,450).

**Insight 5: Black is the runaway top-selling color**
Black-colored bikes generated $577,790 (41%) of revenue — about 1.5× the next-highest color, Red ($380,690) — and this pattern holds across most countries, suggesting color preference is a durable driver rather than a one-market fluke.

**Insight 6: Profit margin is unusually flat across the whole business**
Margin holds steady around 36–37% regardless of month, country, age group, or color. This uniformity suggests pricing and cost structure are set centrally rather than varying by segment — meaning profit growth tracks revenue growth almost one-to-one, with little segment-level margin upside currently being captured.

---

## 9. Recommendations

| Priority | Recommendation | Based On | Suggested Owner |
|----------|-----------------|----------|-------------------|
| High | Plan inventory and staffing ahead of the Oct→Dec demand ramp, especially for Road Bikes in Black, to avoid stockouts during the peak month. | Insight 1, 2, 5 | Operations / Inventory |
| Medium | Weight marketing spend toward Australia and the 25–34 age segment, given their outsized share of revenue relative to other markets and age groups. | Insight 3, 4 | Marketing |
| Low | Investigate whether pricing could be adjusted by category or country to create healthier margin variance, since margins are currently almost identical everywhere. | Insight 6 | Pricing / Finance |

---

## 10. Assumptions & Limitations

### Assumptions
- Each row represents one completed sales transaction, with no duplicates or cancellations to filter out.
- `Unit_Price` is the customer-facing price and `Unit_Cost` is the company's cost; `Revenue`, `Total_Cost`, and `Profit` were derived from these since the raw file didn't include them directly.
- The `Age_Group` bins (Under 25, 25–34, 35–44, 45–54, 55+) were adopted as a standard, business-relevant segmentation.

### Limitations
- Only one quarter (Q4 2021) of data is available — there's no way to confirm the December spike is a recurring seasonal pattern versus a one-off event without a full year or prior-year comparison.
- There's no customer ID, so repeat customers can't be distinguished from one-time buyers, and retention/lifetime value can't be measured.
- `Profit` reflects gross margin only (Unit_Cost vs Unit_Price) — it excludes shipping, marketing, and other operating costs, so it likely overstates true net profitability.
- Country-level revenue isn't normalized against population or market size, so cross-country comparisons reflect absolute sales volume rather than relative market penetration.

---

## 11. Future Enhancements

- [ ] Extend the dataset to a full year to confirm whether the Q4 spike is a true recurring seasonal pattern.
- [ ] Add a customer ID field to analyze repeat-purchase rate and customer lifetime value.
- [ ] Rebuild the enriched dataset as a Power BI report for a shareable, web-based interactive dashboard.
- [ ] Incorporate marketing spend by country to calculate true ROI rather than gross margin alone.

---

## 12. Deliverables

| Deliverable | Description | Location |
|-------------|--------------|-----------|
| `raw_data.xlsx` | Original, unmodified sales export (752 transactions, Oct–Dec 2021) | `data/raw/raw_data.xlsx` |
| `Processed_data_.xlsx` | Enriched dataset with Revenue, Total_Cost, Profit, Profit_Margin_%, and Age_Group | `data/processed/Processed_data_.xlsx` |
| `Dashboard.xlsx` | Interactive PivotTable dashboard — 10 PivotCharts plus a slicer | `visuals/Dashboard.xlsx` |
| `Professional_Analytical_Questions_Report.docx` | Structured framework of 20+ analytical questions across 6 themes | `reports/Professional_Analytical_Questions_Report.docx` |

---

## 13. Author

**Hazem**
Junior Data Analyst

- 🔗 [LinkedIn](https://www.linkedin.com/in/7azem-waleed-ai/?isSelfProfile=true)
- 💼 [GitHub — https://github.com/HazemWaleed517)
- 📧 reem.aweys21@gmail.com

---

*Last updated: September 2026*
