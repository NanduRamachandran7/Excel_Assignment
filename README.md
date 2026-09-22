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









# Excel Assignment 2 - Data Cleaning & Transformation

## Overview
This repository contains the completed **Excel Assignment 2**, focusing on data cleaning, missing value imputation, structural data transformation, deduplication, and conditional formatting on a product dataset initially consisting of **34 items** (Rows 2–35)[cite: 7].

## Summary of Tasks Completed

* **Missing Value Imputation:** 
  * Imputed missing prices for Headphones (`$250.00`) and Sunglasses (`$245.00`) using `=AVERAGEIF`[cite: 7].
  * Calculated conditional median price (`$65`) for Outdoor equipment using `=MEDIAN(IF(F2:F32="Outdoor", D2:D32))` to avoid outlier distortion[cite: 7].
  * Filled all 4 blank category cells with `"Not available"` using `Go To Special` -> `Blanks` (`Ctrl + G`)[cite: 7].
* **Text Standardization & Typo Correction:** Standardized lowercase product titles (`laptop`, `smartphone`, `headphones`) to Title Case and corrected widespread category typos (`Electroni` -> `Electronics`) using `Find & Replace` (`Ctrl + H`)[cite: 7].
* **Duplicate Purging:** Purged **3 exact duplicate rows** (HP Laptop, Bose Headphones, Samsonite Bag) via `Data Tab` -> `Remove Duplicates`, bringing the clean row count from **34** to **31 items**[cite: 7].
* **Data Splitting & Merging:**
  * Decomposed composite `Product ID` codes (e.g., `28-JAN-US`) into independent `Manufacturing Date` and `Country Code` columns using `Flash Fill` (`Ctrl + E`)[cite: 7].
  * Combined Brand Name and Product Name into a unified `Product Brand` column using `=TEXTJOIN("-", TRUE, B2, C2)`[cite: 7].
* **Formatting & Visual Rules:** Standardized prices to Currency (`$#,##0`), formatted dates to `DD-MM-YYYY`, applied gradient blue `Data Bars` across Price values, and configured a `Highlight Cells Rule` for the `Electronics` category[cite: 7].

## Files Included
* `Excel Assignment 2 - Data Cleaning.xlsx` – Excel workbook containing raw data, cleaned dataset, and step-by-step formulas.
* `Data_Cleaning_Documentation.docx` – Comprehensive step-by-step documentation report[cite: 7].
