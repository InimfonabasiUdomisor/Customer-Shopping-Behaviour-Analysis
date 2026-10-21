# Customer Shopping Behavior Analysis

**Revenue, delivery, retention and customer segmentation for a fashion retailer (Excel + interactive HTML dashboard)**

An end-to-end analysis of 10,000 retail orders from 2,581 customers across 20 Indian cities, 1 January 2023 to 31 December 2024. The project takes a raw transaction sheet through documented cleaning, a nine-sheet Excel analysis and an interactive dashboard, and ends with quantified recommendations.

> **Headline:** Revenue is flat at ₹18.90M (−1.4% year on year) because new-customer acquisition fell 72% in 2024 and 30% of all orders arrive six or more days late. Champions and Loyal customers are 31% of the base but generate 66% of revenue.

---

## Table of contents

1. [Business questions](#business-questions)
2. [Dataset](#dataset)
3. [Repository contents](#repository-contents)
4. [Workbook guide](#workbook-guide)
5. [Data cleaning and quality](#data-cleaning-and-quality)
6. [Methodology](#methodology)
7. [Data validation](#data-validation)
8. [Key findings](#key-findings)
9. [Recommendations](#recommendations)
10. [Limitations](#limitations)
11. [How to use](#how-to-use)
12. [Tools and skills](#tools-and-skills)
13. [Data source and licence](#data-source-and-licence)
14. [Author](#author)

---

## Business questions

1. How is revenue trending, which months and festivals matter, and what should 2025 look like?
2. What do sale events do to order volume, order value and returns?
3. Does delivery speed affect ratings and returns, and what does late delivery cost?
4. Who buys what, by age, gender, category, wear type and city?
5. Which customers generate the revenue (RFM segmentation)?
6. Is the business still acquiring new customers, and keeping existing ones?
7. Where are returns concentrated?

## Dataset

| Item | Detail |
|---|---|
| Grain | One row per order |
| Size | 10,000 orders, 2,581 customers, 26 source fields (32 after cleaning) |
| Period | 1 Jan 2023 to 31 Dec 2024 |
| Currency | Indian rupees (₹) |
| Scope | Clothing, footwear and accessories; online (six marketplaces) and in-store; 20 cities |
| Customer-level fields | Age, gender, city, subscription status and shopping frequency are constant for each customer |
| Personal data | None: no names, emails or phone numbers |

Key fields: Transaction ID, Customer ID, Purchase Date, Age, Gender, Location, Online/Offline, Online Store, Category, Item Purchased, Brand, Color, Size, Quantity, Purchase Amount (₹), Discount (%), Festival/Sale, Shipping Charge (₹), Delivery Speed, Delivery Time (Days), Subscription Status, Payment Method, Review Rating, Return Status, Previous Purchases, Frequency of Purchases. The full field list, with notes on every cleaned and derived column, is in the **Data Dictionary** sheet.

## Repository contents

```
.
├── README.md
├── customer_shopping_behavior_Analysis.xlsx      # Cleaned data + all analysis sheets
├── customer_shopping_dashboard.html              # Interactive dashboard (open in any browser)
├── Customer_Shopping_Behavior_Findings_and_Recommendations.docx   # 2-page findings report
└── images/                                       # Dashboard screenshots (optional)
```

## Workbook guide

| Sheet | Purpose |
|---|---|
| Data Dictionary | Every original and derived field, with cleaning notes and flags |
| customer_shopping_behavior | Original raw data, left untouched |
| Customer_Shopping_Cleaned | Cleaned structured table (`Data`) with six helper columns |
| Time Series | Monthly revenue, orders, AOV, 3-month moving average, festival-peak flag, YoY growth, forecast back-test, 2025 baseline, sale-event windows and event impact |
| Logistics | Delivery speed and delivery days vs rating and returns, late vs on-time comparison, platform comparison, rating vs return, subscription impact on shipping |
| Demographics | Age bracket, gender, category, Western vs Traditional wear, item and city views |
| RFM Analysis | Recency, frequency, monetary scoring for 2,581 customers and a segment summary |
| Acquisition & Retention | New vs returning customers, retention of the 2023 base, revenue by cohort, new customers by month |
| Sheet5 | Customer-level pivot feeding the RFM table |

## Data cleaning and quality

The raw sheet is preserved; all fixes are made in a separate cleaned table and documented in the Data Dictionary.

| Issue | Treatment |
|---|---|
| Garbled currency symbol in raw headers (â‚¹) | Encoding repaired to ₹ |
| 2,249 offline orders with no online store | Labelled `In-store` |
| 2,249 offline orders with no delivery speed | Labelled `Offline Purchase` (0 shipping, 0 delivery days) |
| 6,788 orders with no festival/sale value | Labelled `Normal Day` |
| 1,035 accessory orders with no size | Marked not applicable |
| Outliers (largest order ₹22,306) | Reviewed and retained: high-quantity footwear orders |
| Discount pattern: Normal Day average 45.0% vs sale events 7.4% | Flagged, kept, not used to judge discount effectiveness |
| Brands appearing in several categories | Flagged; brand not used as a category proxy |

**Derived columns** (Excel formulas in the cleaned table):

```text
Year-Month         =TEXT([@[Purchase Date]],"yyyy-mm")
Age Bracket        =IFS([@Age]<=30,"18-30",[@Age]<=45,"31-45",[@Age]>45,"46-59")
Wear Type          =IF([@Category]<>"Clothing","N/A",IF(OR([@[Item Purchased]]="Kurta",[@[Item Purchased]]="Saree"),"Traditional Wear","Western Wear"))
Returned Flag      =IF([@[Return Status]]="Returned",1,0)
Late Delivery Flag =IF(AND([@[Online/Offline]]="Online",[@[Delivery Time (Days)]]>=6),1,0)
Sale Flag          =IF([@[Festival/Sale]]<>"Normal Day",1,0)
```

## Methodology

- **Revenue trend:** monthly revenue, orders and average order value; 3-month moving average; a month is a festival peak when revenue exceeds 108% of its moving average; year-over-year growth.
- **Forecast:** a 2024 back-test compares a seasonal-naive method with a linear trend using mean absolute percentage error (MAPE). Seasonal-naive wins (6.3% vs 8.2%), so the 2025 baseline is 2024 × the 2024 ÷ 2023 growth factor (0.986), shown with a ±MAPE range.
- **Sale events:** seven events with 2023 and 2024 date windows. Impact is measured as revenue per day inside sale windows vs normal days, so events of different length compare fairly.
- **Late delivery:** the 6-day cut-off was found in the data: average ratings stay near 4.0 up to day 5 and fall to about 3.4 from day 6.
- **RFM:** recency, frequency and monetary value are each converted to a percentile rank scaled 0 to 10 using `PERCENTRANK.EXC`, summed to a 0 to 30 score, then banded: up to 6 Lost/Hibernating, up to 12 At Risk, up to 18 Potential Loyalist, up to 24 Loyal Customer, up to 30 Champions. Snapshot date: 1 Jan 2025.
- **Retention:** first purchase date per customer (`MINIFS`) separates new from returning customers; retention is the share of 2023 customers who ordered again in 2024.

## Data validation

- 10,000 rows in raw and cleaned sheets; zero duplicate transaction IDs; 2,581 unique customers.
- Total revenue of ₹18,898,820 reconciles across the monthly table, annual split, age brackets, gender, category and city tables; order counts reconcile to 10,000 across stores and months.
- Range checks: ages 18–59, quantity 1–4, ratings 1–5, discounts 0–70%, no order at or below zero, dates inside the stated period.
- Customer attributes never change within a customer; Previous Purchases matches chronological order on every row.
- Offline orders carry zero shipping and zero delivery days; Same Day orders take 1 day, Express 2–4 days, Standard 4–8 days.
- Forecast back-tested on 2024 before use; dashboard headline tiles match the workbook.

## Key findings

| # | Finding | Evidence |
|---|---|---|
| 1 | Revenue is flat and peaks are sale-driven | ₹9.51M (2023) vs ₹9.38M (2024), −1.4%; five peak months, all sale-event months |
| 2 | Sale events add volume, not value | 32% of orders in 16% of days; revenue per day +44%; orders per day 2.4×; average order value ₹1,286 vs ₹2,175 (−41%) |
| 3 | Late delivery hurts ratings and returns | 3,007 orders took 6+ days: rating 3.42 vs 4.08, return rate 21.8% vs 16.5%, about ₹299K excess returned revenue; all late orders are Standard shipping |
| 4 | Acquisition has collapsed | New customers 2,011 → 570 (−72%); retention 72.7%; 83% of 2024 revenue from the 2023 base |
| 5 | A small segment carries the business | Champions (10%) and Loyal (20%) customers generate 66% of revenue; 629 customers At Risk |
| 6 | Footwear punches above its weight | 28% of orders, 42% of revenue (AOV ₹2,785) |
| 7 | Demand is concentrated in seven metros | 64.5% of revenue; the other 13 cities contribute 2–3.6% each |
| 8 | Returns are material | 18.0% of orders, ₹3.51M (18.6% of revenue); 1–2 star orders are returned 48% of the time |
| 9 | 2025 baseline | About ₹9.26M (range ₹8.67M to ₹9.84M) if nothing changes |

## Recommendations

1. **Fix Standard delivery first:** 5-day service target; default or low-cost Express upgrade (Express averages ₹56 vs ₹53 for Standard).
2. **Rebuild acquisition:** move part of the sale-event budget into acquisition channels and track new customers monthly.
3. **Protect and win back:** loyalty perks for Champions and Loyal customers; win-back campaign for At Risk customers; target 80% retention.
4. **Reshape the sale calendar:** keep Independence Day, Amazon Great Summer and Holi; review the long Diwali and End of Season windows; add bundles and spend thresholds.
5. **Lean into footwear:** prioritise stock, bundle accessories, add size guidance to reduce clothing returns.
6. **Tighten governance and feedback:** audit the discount data; trigger proactive support for late and 1–2 star orders.

## Limitations

- Revenue is **gross of returns**; net revenue after returns is about ₹15.4M.
- Two years of history support only a single-factor baseline forecast.
- The fall in new customers may partly reflect how the dataset was assembled; confirm with the data owner.
- The discount pattern is counter-intuitive, so no conclusions are drawn about discount effectiveness.
- No cost, margin, stock or marketing data: findings describe revenue and behaviour, not profit or return on investment.
- RFM segment sizes are partly mechanical because they are rank-based; revenue shares carry the insight.
- Delivery and ratings are linked by association, not proven cause.

## How to use

1. **Workbook:** open `customer_shopping_behavior_Analysis.xlsx` in Excel 2019 or Microsoft 365 (the formulas use `IFS`, `MINIFS` and `PERCENTRANK.EXC`). Start at the **Data Dictionary** sheet, then move through the analysis sheets in order.
2. **Dashboard:** open `customer_shopping_dashboard.html` in any modern browser; no installation or internet connection is needed. Filter by channel and time period; all charts update together.
3. **Report:** `Customer_Shopping_Behavior_Findings_and_Recommendations.docx` is a two-page summary for non-technical readers.
4. **Reproduce:** the analysis sheets are driven by live formulas against the `Data` table, so changes to the cleaned data flow through automatically.

## Tools and skills

Excel (structured tables, IFS, COUNTIFS, SUMIFS, AVERAGEIFS, MINIFS, PERCENTRANK.EXC, PivotTables, moving averages, forecast back-testing) · interactive HTML dashboard · data cleaning and documentation · RFM segmentation · cohort-style acquisition and retention analysis · logistics and customer-experience analytics · executive reporting.

## Data source and licence

- **Source:** (https://www.kaggle.com/datasets/rewantbhriguvanshi/customer-shopping-trends-indian)
- **Nature of data:** Public
- No confidential employer data is used.

## Author

**Inimfonabasi Udomisor** | Data & Business Intelligence Analyst
inimfonabasiudomisor@gmail.com
