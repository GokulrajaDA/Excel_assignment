# Excel Functions Assignment - README

## **Overview**
This assignment demonstrates core Excel functions for data analysis and text manipulation using a product inventory dataset with 34 items across multiple categories and regions.

## **Dataset Structure**
| Column A | Column B | Column C | Column D | Column E | Column F |
| --- | --- | --- | --- |
| Product ID | Product Name | Brand Name | Price ($) | Quantity | Category |

Data range: `A1:F35` with headers in row 1.

## **Function Implementation**

### **1. Mathematical & Statistical Functions**
Applied on `Price ($)` column D2:D35:

| Function | Formula | Result | Purpose |
| --- | --- | --- | --- |
| SUM | `=SUM(D2:D35)` | $11,490 | Total inventory value |
| COUNT | `=COUNT(D2:D35)` | 34 | Count of products with price |
| AVERAGE | `=AVERAGE(D2:D35)` | $337.94 | Average product price |
| MIN | `=MIN(D2:D35)` | $30 | Lowest priced product |
| MAX | `=MAX(D2:D35)` | $1,000 | Highest priced product |

### **2. Logical Function – IF**
Added column G: `Price Category`  
Formula in G2: `=IF(D2>=500,"High Price","Standard Price")`  
Classifies products as "High Price" if price ≥ $500, else "Standard Price".

### **3. Conditional Functions – SUMIF and COUNTIF**
| Function | Formula | Result | Purpose |
| --- | --- | --- | --- |
| SUMIF | `=SUMIF(F2:F35,"Electronics",D2:D35)` | $7,660 | Total price of Electronics category |
| SUMIF | `=SUMIF(F2:F35,"Kitchen",E2:E35)` | 145 | Total quantity of Kitchen items |
| COUNTIF | `=COUNTIF(D2:D35,"<100")` | 12 | Count of products under $100 |
| COUNTIF | `=COUNTIF(F2:F35,"Fashion")` | 7 | Count of Fashion category items |

### **4. Text Formatting – LEFT, RIGHT, MID**
Applied on `Product ID` column A to extract date components:

| Function | Formula in H2 | Purpose | Example from A2 |
| --- | --- | --- | --- |
| LEFT | `=LEFT(A2,2)` | Extract day | "28" from 28-JAN-US |
| MID | `=MID(A2,4,3)` | Extract month | "JAN" from 28-JAN-US |
| RIGHT | `=RIGHT(A2,2)` | Extract country code | "US" from 28-JAN-US |

Add columns H: `Day`, I: `Month`, J: `Country` and drag formulas down to row 35.

## **How to Replicate**
1. Open the Excel file with the product table
2. Insert new columns G-J for `Price Category`, `Day`, `Month`, `Country`
3. Enter formulas in row 2 as shown above
4. Drag all formulas down to row 35 to apply to entire dataset
5. Add summary formulas below the table in D37:D41 for SUM, COUNT, AVERAGE, MIN, MAX

## **Key Learning Outcomes**
- Aggregate data using SUM, COUNT, AVERAGE, MIN, MAX
- Apply business logic with IF for price classification
- Perform conditional analysis using SUMIF and COUNTIF
- Parse structured text with LEFT, RIGHT, MID functions

## **Notes**
- All price values are numeric despite `$` formatting in Excel
- Product ID format is `DD-MMM-CC` where DD=day, MMM=month, CC=country
- Use absolute references `$D$2:$D$35` when creating dashboard summaries

---
*Excel Assignment | 2026*
