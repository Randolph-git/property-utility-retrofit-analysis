# property-utility-retrofit-analysis

**An Excel-first sustainability analytics case study: from utility-data quality checks to building comparisons, anomaly investigations and conditional retrofit decisions.**

**Author:** Randolph O. Ibe

> **Synthetic-data disclaimer:** All properties, utility records, weather inputs, project estimates, incentive programs, benchmarks and results are fictional practice data. This is an independent portfolio project, not an official company project, company endorsement or disclosure of confidential company information.

## Project overview

This project uses **12 fictional Ontario apartment properties** and **24 months of data, January 2024–December 2025**, to answer three management questions:

1. **Can the utility records be trusted for comparison?**
2. **Which consumption patterns warrant investigation, and what evidence is still missing?**
3. **How dependent are proposed retrofit returns on unverified savings assumptions?**

The workflow deliberately separates **data-quality exceptions**, **operational investigation priorities** and **capital-project approval**. A high consumption rank is not proof of inefficiency; a positive modelled NPV is not proof of achievable savings.

## Main deliverables

| File | Purpose |
|---|---|
| `Property_Utility_Retrofit_Model.xlsx` | Working Excel model containing source tables, queries, checks, PivotTables, investigations, financial scenarios and charts |
| `Property_Utility_Retrofit_Executive_Summary.pdf` | One-page management summary of the two proposals and conditional recommendations |
| `Property_Utility_Retrofit_Dashboard.pdf` | Four-chart visual explanation of the findings and financial sensitivity |

## Tools and skills demonstrated

- **Excel Tables:** structured inputs, stable identifiers and auditable references.
- **Power Query:** data types, text cleanup, duplicate checks, quality flags and multi-key merges.
- **PivotTables:** raw and cleaned record-coverage summaries.
- **Excel formulas:** `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MAXIFS`, `IF`, `XLOOKUP`, `FILTER`, `SORTBY`, `MEDIAN`, `SEQUENCE`, `DATE` and `TEXT`.
- **Financial modelling:** simple payback, net present value (NPV), break-even savings and downside scenarios using `PV` and `PMT`.
- **Data visualization:** sorted bar charts, monthly line comparisons, an HDD-context scatter plot and clustered scenario bars.
- **Analytical judgement:** separating screening estimates from verified causes, savings and approval decisions.

## Analysis workflow

### 1. Preserve and audit the source data

`Utility_Raw` was kept unchanged. `Data_Quality`, `DQ_Completeness` and `DQ_Issue_Log` were added to check coverage, duplicates, reading types, billing periods, measurement units, critical fields and cost reconciliation.

The checks identified:

- **Three missing monthly records:** APT-001 gas in March 2024, APT-003 electricity in November 2024, and APT-007 water in August 2025.
- **One extra duplicate invoice occurrence:** APT-012 electricity in March 2025.
- **Eight estimated readings**, retained with caution flags.
- **One 62-day electricity bill** at APT-005, also flagged for billing-month alignment review.

The raw dataset contains **862 records**. Removing the extra duplicate occurrence produces **861 cleaned records**, without inventing missing bills or silently repairing unusual records.

### 2. Build repeatable Power Query processing

| Query | Role |
|---|---|
| `q_UtilityRaw` | Prepare source types and clean key text |
| `q_Invoice_Counts` | Count occurrences of each invoice ID |
| `q_UtilityRaw_Audit` | Add duplicate, reading-quality, billing-period, unit, cost and critical-field flags |
| `Utility_Clean` | Remove the extra duplicate occurrence and load the cleaned worksheet table |
| `q_Properties` | Prepare property-size and location lookups |
| `q_Occupancy` | Prepare monthly occupied-suite lookups |
| `q_Weather` | Prepare city-month weather lookups |
| `q_Utility_Analysis` | Enrich the cleaned records and load `tblUtilityAnalysis` |

Connection-only staging queries keep preparation separate from worksheet analysis. The final analysis uses left-outer merges on:

- **Property:** `Property_ID`
- **Occupancy:** `Property_ID + Month_Start`
- **Weather:** `City + Month_Start`

Unique lookup keys, row counts and blank expanded fields were checked so the merges did not multiply records or silently lose context. The enriched analysis retained **861 records**.

### 3. Check coverage before comparing buildings

Two PivotTables answer different questions:

- **`DQ_Completeness`** counts raw records by property, reporting period and utility. Its total is **862**, including the duplicate occurrence.
- **`Benchmark_Coverage`** checks cleaned analysis coverage by property, year and utility. Its total is **861**.

Those totals are intentionally different. Incomplete years and the flagged long-bill property-year are withheld from ordinary annual rankings.

### 4. Calculate utility-specific intensity measures

| Sheet | Measure | Reason |
|---|---|---|
| `Electricity_Benchmark` | Annual kWh ÷ floor area in m² | Compare use relative to building size |
| `Water_Benchmark` | Annual water m³ ÷ average monthly occupied suites | Account for occupied residential capacity |
| `Gas_Benchmark` | Annual gas m³ ÷ floor area in m² | Screen gas intensity with weather context |

Comparable property-years are ranked within the same year and assessed against the comparable 2025 portfolio median. These medians are **internal screening references**, not regulatory standards or verified efficiency targets.

Annual heating degree days (**HDD**) are summed once per gas property-month. HDD provides heating-demand context; this project does **not** claim a fitted weather-normalization model.

### 5. Investigate unusual water and gas use

#### APT-006: water

APT-006's 2025 water intensity was approximately **203.13 m³ per average occupied suite per year**, versus a comparable-property median of **153.34**.

A monthly review identified an April–June increase and a subsequent return toward normal. `Water_Cost_Investigation` compared the affected months with the same months in 2024:

```text
Expected 2025 use = 2024 use / 2024 occupied suites × 2025 occupied suites
Estimated excess = actual 2025 use − expected 2025 use
Consumption effect = estimated excess × 2025 unit rate
Rate effect = expected 2025 use × (2025 rate − 2024 rate)
```

| April–June screening output | Result |
|---|---:|
| Estimated excess use | 8,572.98 m³ |
| Consumption-driven variable-cost effect | CAD 39,007.04 |
| Rate effect on expected use | CAD 2,048.74 |
| Combined modelled variable-cost effect | CAD 41,055.78 |

**Interpretation:** unexplained consumption requiring meter, billing and maintenance checks—not confirmed leak volume or recoverable savings. The cost decomposition excludes fixed/other charges and is not the difference between two actual total bills.

#### APT-004: gas

February gas use increased from **46,630.3 m³ in 2024** to **60,821.2 m³ in 2025**, while HDD fell from **627 to 565**. Both February readings were actual, but January 2025 was estimated.

`Investigation_Log` records the need to reconcile adjacent bills and review meter, boiler/control and maintenance information before attributing the increase to an equipment fault or proposing verified savings.

### 6. Test two retrofit proposals

`Retrofit_business_Case` uses project-ID lookups to evaluate a boiler replacement and a water leak-detection system. Proposal savings remain **assumed, not metered**. The base case uses **zero confirmed incentives**, rather than treating illustrative potential funding as secured.

| Measure | PRJ-002: boiler | PRJ-003: leak detection |
|---|---:|---:|
| Gross capital cost | CAD 240,000 | CAD 76,000 |
| Assumed annual utility savings | CAD 43,000 | CAD 24,500 |
| Annual maintenance-cost change | CAD −2,500 | CAD +800 |
| Assumed net annual benefit | CAD 45,500 | CAD 23,700 |
| Useful life / discount rate | 20 years / 7% | 10 years / 7% |
| Simple payback | 5.27 years | 3.21 years |
| Assumed base-case NPV | CAD 242,027.65 | CAD 90,458.88 |
| Break-even annual utility savings | CAD 20,154.30 | CAD 11,620.69 |
| NPV at 50% of proposed utility savings | CAD 14,256.34 | CAD 4,420.01 |
| NPV at 40% of proposed utility savings | CAD −31,297.92 | CAD −12,787.77 |

The model assumes constant year-end benefits. Only utility savings change in the 50% and 40% scenarios; maintenance changes and other assumptions remain fixed.

**Decision:** both proposals warrant further evaluation, but neither is approval-ready. Their positive proposed-case NPVs do not establish that the required recurring savings are achievable. A temporary water spike cannot automatically substantiate years of leak-detection savings.

## Dashboard

| Chart | What it communicates |
|---|---|
| **2025 water intensity — comparable properties** | Identifies the APT-006 outlier using sorted bars, an orange investigation highlight and a linked median text note; incomplete APT-007 coverage is excluded |
| **APT-006 monthly water use per occupied suite** | Compares 2024 and 2025 patterns and highlights the April–June unexplained increase |
| **2025 gas intensity and heating demand** | Plots annual HDD against gas m³/m²/year, labels properties and highlights APT-004 without claiming weather-normalized inefficiency |
| **Modelled retrofit NPV — utility-savings scenarios** | Shows positive and negative NPV outcomes under proposed, 50% and 40% savings assumptions |

`Utility_Dashboard` presents the charts. `Gas_Chart_Data` and `Retrofit_Chart_Data` keep supporting ranges traceable. The water chart-source ranges are on the dashboard sheet itself. Labels, units and synthetic-data caveats travel with the charts when copied.

## How to review the project

1. **Start with the executive-summary PDF** for the recommendations.
2. **Review the dashboard PDF** for the visual story.
3. **Download and open the workbook in desktop Excel** with support for Power Query and the dynamic-array functions listed above.
4. Read `Start_Here` and `Data_Dictionary`, then follow the source → quality checks → cleaned analysis → benchmarks → investigations → business case → summary/dashboard sequence.

No Python runtime or database is needed to review the completed Excel analysis. Preserve an unchanged copy before experimenting with formulas or source data. Refreshing queries does not automatically reconcile manually entered investigation inputs or rewrite narrative findings; those must be checked when data changes.


## Validation and limitations

Development checks covered preservation of source values, duplicate handling, unique merge keys, row counts, joined fields, benchmark results, investigation calculations, financial scenarios and formula consistency in a reviewed workbook snapshot. Later cleanup, chart layout and final PDF readability were checked manually during guided development; the final revised workbook was not subjected to a second independent audit.

This project does **not** establish:

- A confirmed leak, boiler fault or recoverable savings amount.
- Verified engineering savings or persistence over project life.
- Formal regression-based weather normalization.
- Official benchmark, emissions-factor or incentive-program validity.
- A confirmed incentive award or approved investment.
- A full life-cycle cost model, common-horizon capital optimization or automated native-Excel refresh test.

**Main takeaway:** reliable sustainability analysis requires a visible chain from data quality, to qualified comparisons, to investigation evidence, to conditional financial decisions—not simply an attractive dashboard or positive projected return.
