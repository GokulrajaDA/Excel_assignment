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
# Excel Data Cleaning Assignment using Power Query - README

## **Project Overview**
This assignment performs end-to-end data cleaning and transformation using **Power Query Editor** in Excel. The dataset contains 34 product records. All steps are documented as applied transformations for reproducibility.

**Source Data**: `Table1` range `A1:F35`
**Tool Used**: Power Query → `Data > Get & Transform Data > From Table/Range`

---

## **Tasks Performed in Power Query**

### **1. Handling Missing Values**

#### **A. Price Column Check**
**Step in Power Query**:
1. Select `Price ($)` column → `Home > Replace Values`
2. Value to Find: `null` → Replace with: `150`
3. Or use `Transform > Replace Errors` if prices are errors

**M Code**:
```M
= Table.ReplaceValue(#"PreviousStep", null, 150, Replacer.ReplaceValue, {"Price ($)"})Handling Strategies Applied:Impute with Median: Home > Statistics > Median → Replace nulls with $150Remove Rows: Home > Remove Rows > Remove Blank Rows if price is mandatoryFlag for Review: Add Column > Custom Column → = if [Price ($)] = null then "Missing" else "OK"B. Category Column Missing Values
Step: Select Category → Transform > Replace Values
Replace: null with "Uncategorized" or use ModeM Code:M= Table.ReplaceValue(#"PreviousStep", null, "Uncategorized", Replacer.ReplaceValue, {"Category"})Imputation Strategy: Use Group By to find mode, then Merge Queries to fill blanks.2. Correcting Inconsistent Data
A. Product Name Standardization
Steps:Select Product Name → Transform > Format > Trim → Removes extra spacesTransform > Format > Capitalize Each Word → Title CaseM Code:M= Table.TransformColumns(#"PreviousStep", {{"Product Name", each Text.Proper(Text.Trim(_)), type text}})B. Category Typos Correction
Steps:Select Category → Transform > Replace ValuesReplace: "Electonics" → "Electronics", "Accesories" → "Accessories"M Code:M= Table.ReplaceValue(#"PrevStep", "Electonics", "Electronics", Replacer.ReplaceText, {"Category"})Alternative: Transform > Replace Values using List for bulk corrections.3. Removing Duplicates
Steps:Home > Remove Rows > Remove DuplicatesPower Query checks all columns by defaultM Code:M= Table.Distinct(#"PreviousStep")Result: 1 duplicate removed. 33 unique rows loaded.4. Splitting and Merging Data
A. Split Product ID → Manufacturing Date & Country Code
Product ID format: 28-JAN-USSteps:Select Product ID → Transform > Split Column > By DelimiterDelimiter: - → Split at: Each occurrence → 3 columns createdRename: Product ID.1 → Day, Product ID.2 → Month, Product ID.3 → Country CodeManufacturing Date: Add Column > Custom Column
Formula: = Date.FromText([Day] & "-" & [Month] & "-2024")Delete Day and Month columns → Home > Remove ColumnsM Code:M= Table.SplitColumn(#"PreviousStep", "Product ID", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"Day", "Month", "Country Code"}),
#"Added Date" = Table.AddColumn(#"Split", "Manufacturing Date", each Date.FromText([Day] & "-" & [Month] & "-2024"), type date)B. Merge Brand Name + Product Name → Product Brand
Steps:Select Brand Name then Product Name using Ctrl → Transform > Merge ColumnsSeparator: Space → New column name: Product BrandM Code:M= Table.CombineColumns(#"PreviousStep", {"Brand Name", "Product Name"}, Combiner.CombineTextByDelimiter(" ", QuoteStyle.None), "Product Brand")5. Number Formatting in Power Query
A. Price Column → Currency
Steps: Select Price ($) → Transform > Data Type > Currency
M Code: = Table.TransformColumnTypes(#"PrevStep", {{"Price ($)", Currency.Type}})B. Manufacturing Date → DD-MM-YYYY
Steps: Select Manufacturing Date → Transform > Data Type > Date
Display format handled in Excel after Close & Load. In sheet: Ctrl+1 > Custom > DD-MM-YYYY6. Conditional Formatting (Applied in Excel after Load)
Power Query loads clean data. Formatting applied in Excel sheet:A. Price Column - Data Bars
Select loaded Price ($) column in ExcelHome > Conditional Formatting > Data Bars > Blue Data BarB. Category Column - Highlight Electronics
Select Category column in ExcelHome > Conditional Formatting > New Rule > Use a formulaFormula: =$F2="Electronics" → Format: Fill Light BluePower Query Load Steps
Home > Close & Load To → Table → Existing worksheetCheck Add this data to the Data Model if needed for PivotTablesFinal Table Structure After Power QueryProduct IDManufacturing DateCountry CodeProduct BrandPrice ($) | Quantity | CategoryKey M Code SummaryMlet
    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Price ($)", type number}}),
    #"Replaced Price Nulls" = Table.ReplaceValue(#"Changed Type",null,150,Replacer.ReplaceValue,{"Price ($)"}),
    #"Trimmed Text" = Table.TransformColumns(#"Replaced Price Nulls",{{"Product Name", Text.Proper, type text}}),
    #"Removed Duplicates" = Table.Distinct(#"Trimmed Text"),
    #"Split ID" = Table.SplitColumn(#"Removed Duplicates", "Product ID", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"Day", "Month", "Country Code"}),
    #"Added Date" = Table.AddColumn(#"Split ID", "Manufacturing Date", each Date.FromText([Day] & "-" & [Month] & "-2024")),
    #"Merged Brand" = Table.CombineColumns(#"Added Date",{"Brand Name", "Product Name"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Product Brand"),
    #"Changed Types" = Table.TransformColumnTypes(#"Merged Brand",{{"Price ($)", Currency.Type}, {"Manufacturing Date", type date}})
in
    #"Changed Types"Advantages of Power Query Approach
Reproducible: All steps saved. Refresh data with Data > Refresh AllNon-destructive: Original data untouchedScalable: Same query works if rows increase from 34 to 10,000Documented: Each step visible in Applied Steps paneKey Learning Outcomes
Handle nulls/errors using Replace Values and Imputation in Power QueryStandardize text with Text.Trim, Text.Proper transformationsRemove duplicates using Table.DistinctSplit columns by delimiter and create date fieldsMerge columns and set proper data types for downstream analysisApply Excel conditional formatting after loading clean dataData Cleaning with Power Query | Excel Assignment | 2026
