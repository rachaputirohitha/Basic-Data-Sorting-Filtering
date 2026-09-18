# Task 11 – Basic Data Sorting & Filtering 📊

## Project Overview

This project focuses on **basic data sorting and filtering** using the **Sample Superstore dataset**.

The objective of this task was to organize retail data, apply sorting and filtering techniques, calculate totals for filtered records, and identify useful business insights from Sales and Profit data using Microsoft Excel.


##  Dataset

**Dataset:** Sample Superstore

The dataset contains **9,994 retail transaction records** with the following fields:

- Ship Mode
- Segment
- Country
- City
- State
- Postal Code
- Region
- Category
- Sub-Category
- Sales
- Quantity
- Discount
- Profit

##  Objectives

The main objectives of this task were:

1. Understand and work with a retail sales dataset.
2. Convert raw data into an Excel Table.
3. Apply sorting techniques.
4. Apply filters to analyze specific groups of data.
5. Calculate totals for filtered records.
6. Answer business-related questions using the dataset.
7. Present the findings in a separate worksheet.


##  Tools Used

- Microsoft Excel
- Sample Superstore Dataset
- Excel Tables
- Sort & Filter
- `SUBTOTAL` Function


##  Tasks Performed

### 1. Created a Working Copy

A working copy of the original dataset was created so that the raw data remained unchanged.

### 2. Converted Data into an Excel Table

The dataset was converted into an Excel Table to enable:

- Column filtering
- Sorting
- Easier data management
- Structured analysis

### 3. Applied Sorting

Sales and Profit columns were sorted from **Largest to Smallest** to help identify high-value records.

### 4. Applied Filters

Filters were applied to:

- Category
- Region
- Sub-Category
- Segment

### 5. Calculated Filtered Totals

The `SUBTOTAL` function was used to calculate totals while filters were applied.


##  Excel Formulas Used

### Total Sales for Filtered Data

```excel
=SUBTOTAL(9,J2:J9995)
