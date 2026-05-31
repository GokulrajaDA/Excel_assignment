markdown# Excel Functions Assignment - README

This project demonstrates the use of core Excel functions for data analysis and text manipulation. The assignment covers mathematical, statistical, logical, and text functions applied to a sample dataset.

## **Objective**
To understand and implement the following Excel functions:
**SUM, MIN, MAX, COUNT, AVERAGE, AVERAGEIF, SUMIF, COUNTIF, IF, RIGHT, UPPER, LOWER**

## **File Structure**
- **Sheet Name**: `Data`
- **Sample Columns**: `Student Name`, `Department`, `Score`, `Region`
- **Formula Columns**: `SUM`, `MIN`, `MAX`, `COUNT`, `AVERAGE`, `AVERAGEIF`, `SUMIF`, `COUNTIF`, `IF`, `RIGHT`, `UPPER`, `LOWER`

## **Function Explanations & Usage**

1. **SUM**  
   Adds all numeric values in a selected range.  
   *Example*: `=SUM(C2:C20)` → Total of all scores.

2. **MIN**  
   Returns the smallest numeric value in a range.  
   *Example*: `=MIN(C2:C20)` → Lowest score.

3. **MAX**  
   Returns the largest numeric value in a range.  
   *Example*: `=MAX(C2:C20)` → Highest score.

4. **COUNT**  
   Counts the number of cells that contain numbers in a range.  
   *Example*: `=COUNT(C2:C20)` → Total number of scored entries.

5. **AVERAGE**  
   Calculates the arithmetic mean of a range.  
   *Example*: `=AVERAGE(C2:C20)` → Average score.

6. **AVERAGEIF**  
   Calculates the average of cells that meet a given condition.  
   *Example*: `=AVERAGEIF(B2:B20,"Sales",C2:C20)` → Average score for Sales department.

7. **SUMIF**  
   Adds cells in a range that meet a given condition.  
   *Example*: `=SUMIF(B2:B20,"North",C2:C20)` → Total score for North region.

8. **COUNTIF**  
   Counts cells that meet a given condition.  
   *Example*: `=COUNTIF(B2:B20,"IT")` → Number of entries in IT department.

9. **IF**  
   Performs a logical test and returns one value if TRUE and another if FALSE.  
   *Example*: `=IF(C2>=50,"Pass","Fail")` → Checks if a student passed based on score.

10. **RIGHT**  
    Extracts a specified number of characters from the right side of a text string.  
    *Example*: `=RIGHT(A2,3)` → Extracts last 3 characters from the student name.

11. **UPPER**  
    Converts all text to uppercase.  
    *Example*: `=UPPER(A2)` → Converts student name to uppercase.

12. **LOWER**  
    Converts all text to lowercase.  
    *Example*: `=LOWER(A2)` → Converts student name to lowercase.

## **How to Run**
1. Open the Excel file `Excel_Assignment.xlsx`.
2. Enter sample data in columns A-D from row 2 onwards.
3. In columns E-P, apply the respective formulas as shown in the examples above.
4. Drag the formulas down to apply them to all rows.

## **Key Learning Outcomes**
- Basic arithmetic and statistical analysis using Excel.
- Conditional calculations with IF and IF-based functions.
- Text cleaning and formatting using text functions.

## **Notes**
- Ensure there are no blank cells in the range when using COUNT, MIN, and MAX for accurate results.
- Use absolute references `$C$2:$C$20` if you plan to copy formulas across multiple columns.

---
*Created for Excel Assignment | 2026*
