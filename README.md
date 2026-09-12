# Excel Assignment 1 - Data Exploration & Analysis

## Overview
This repository contains the completed **Excel Assignment 1**, focusing on exploratory data analysis, summary statistics, logical categorizations, and text manipulation on a product dataset consisting of **34 items** (Rows 2–35).

## Summary of Tasks Completed

* **Summary Aggregations:** Calculated total dataset price (`$10,100.00`), total product count (`34`), and average price (`$297.06`) using `SUM`, `COUNTA`, and `AVERAGE`.
* **Min & Max Analysis:** Determined the minimum (`$30.00`) and maximum (`$1,000.00`) product unit prices using `MIN` and `MAX`.
* **Price Tier Categorization:** Applied an `IF` statement (`=IF(D2>=500, "High Price", "Standard Price")`) to classify products into pricing tiers.
* **Conditional Functions:** Calculated total Electronics category price (`$8,050.00`) via `SUMIF` and counted budget items under $100 (`11`) using `COUNTIF`.
* **Text Extraction:** Parsed structured `Product ID` codes (formatted as `DD-MMM-CC`) using text formulas:
  * `Day`: `=LEFT(A2, 2)`
  * `Country Code`: `=RIGHT(A2, 2)`
  * `Month`: `=MID(A2, 4, 3)`
* **Revenue Metrics:** Computed extended total product revenue (`$289,700.00`), total quantity sold (`900 units`), and weighted average unit price (`$321.89`) using `SUMPRODUCT`.

## Files Included
* `Excel Assignment 1 - Data Exploration.xlsx` — Excel workbook containing raw data and calculated `AnswerSheet`.
* `Excel_Assignment_1_AnswerSheet_Documentation.docx` — Complete formula documentation guide.
