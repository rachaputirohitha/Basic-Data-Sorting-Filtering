# Task 11 – Basic Data Sorting & Filtering 📊

##  Project Overview

This project focuses on **Basic Data Sorting and Filtering** using the **Sample Superstore dataset** in Microsoft Excel.

The objective of this task was to organize the dataset, apply sorting and filtering techniques, calculate totals for filtered data, and identify important business insights from Sales and Profit data.


##  Objectives

- Work with a real-world retail sales dataset.
- Convert raw data into an Excel Table.
- Apply sorting techniques.
- Apply filters to different columns.
- Calculate Sales and Profit totals for filtered data.
- Answer business-related questions.
- Present the findings in a separate worksheet.
- Keep the original/raw data unchanged.


##  Dataset Information

**Dataset Name:** Sample Superstore

**Number of Records:** 9,994

### Columns Used

| Column | Description |
|---|---|
| Ship Mode | Shipping method |
| Segment | Customer segment |
| Country | Country |
| City | Customer city |
| State | Customer state |
| Postal Code | Postal/ZIP code |
| Region | Sales region |
| Category | Product category |
| Sub-Category | Product sub-category |
| Sales | Sales amount |
| Quantity | Quantity sold |
| Discount | Discount applied |
| Profit | Profit amount |


##  Tools Used

- Microsoft Excel
- Excel Tables
- Sort & Filter
- `SUBTOTAL` Function
- Sample Superstore Dataset


#  Task Implementation

## Step 1 – Create a Working Copy

The original `SampleSuperstore.csv` file was kept unchanged.

A separate Excel working file was created for performing sorting, filtering, and analysis.

**Working File:**

`Task_11_Superstore_Filtering.xlsx`


## Step 2 – Convert Data into an Excel Table

The complete dataset was selected and converted into an Excel Table using:

**Ctrl + A → Ctrl + T**

The option:

**My table has headers**

was selected.

This enabled filter dropdowns for each column.


#  Step 3 – Data Sorting

Sorting was performed to organize values from **Largest to Smallest**.

### Sales Sorting

The `Sales` column was sorted using:

**Sales → Sort Largest to Smallest**

### Profit Sorting

The `Profit` column was also sorted using:

**Profit → Sort Largest to Smallest**

Sorting helped identify high-value Sales and Profit records.


#  Step 4 – Data Filtering

Filters were applied to the following columns:

- Category
- Region
- Sub-Category
- Segment

Filtering allowed specific groups of records to be analyzed without changing the original data.


#  Step 5 – Excel Formulas

## 1. Calculate Filtered Sales

The following formula was used:

1.Total Sales:
=SUM(J2:J9995)

2. Total Profit:
=SUM(M2:M9995)

3. Filtered Total Sales:
=SUBTOTAL(9,J2:J9995)

4. Filtered Total Profit:
=SUBTOTAL(9,M2:M9995)

5. Total Quantity:
=SUM(K2:K9995)

6. Average Sales:
=AVERAGE(J2:J9995)

7. Average Profit:
=AVERAGE(M2:M9995)

8. Average Discount:
=AVERAGE(L2:L9995)

9. Number of Records:
=COUNTA(A2:A9995)

10. Technology Sales:
=SUMIF(H2:H9995,"Technology",J2:J9995)

11. Furniture Sales:
=SUMIF(H2:H9995,"Furniture",J2:J9995)

12. Office Supplies Sales:
=SUMIF(H2:H9995,"Office Supplies",J2:J9995)

13. Technology Profit:
=SUMIF(H2:H9995,"Technology",M2:M9995)

14. Furniture Profit:
=SUMIF(H2:H9995,"Furniture",M2:M9995)

15. Office Supplies Profit:
=SUMIF(H2:H9995,"Office Supplies",M2:M9995)

# Region-Based Formulas:
West Sales:
=SUMIF(G2:G9995,"West",J2:J9995)

West Profit:
=SUMIF(G2:G9995,"West",M2:M9995)

# Segment-Based Formulas:
Consumer Sales:
=SUMIF(B2:B9995,"Consumer",J2:J9995)

Corporate Sales:
=SUMIF(B2:B9995,"Corporate",J2:J9995)

Home Office Sales:
=SUMIF(B2:B9995,"Home Office",J2:J9995)

 # Sub-Category Formula
Phones Sales:
=SUMIF(I2:I9995,"Phones",J2:J9995)
