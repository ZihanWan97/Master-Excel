## Common Excel Statistical Functions

### 1.Basic Statistical Functions

| Function Name      | Example/Syntax           |
| ------------------ | ------------------------ |
| Count              | =COUNT(A2:A20)           |
| Average            | =AVERAGE(A2:A20)         |
| Median             | =MEDIAN(A2:A20)          |
| Mode               | =MODE(A2:A20)            |
| Max                | =MAX(A2:A20)             |
| Min                | =MIN(A2:A20)             |
| Percentile         | =PERCENTILE(A2:A20, .75) |
| Standard Deviation | =STDEV(A2:A20)           |
| Variance           | =VAR(A2:A20)             |

------

### 2.SMALL/LARGE & RANK/PERCENTRANK

| Function Name | Example/Syntax            | Description                                                  |
| ------------- | ------------------------- | ------------------------------------------------------------ |
| Rank          | =RANK(A2,A2:A8)           | Returns the rank of a particular number among a list of values |
| PercentRank   | =PERCENTRANK(array,value) | Returns the rank of a value as a percentage of a given array(range of data) or dataset |
| Small         | =SMALL(A2:A8,2)           | Returns the nth smallest value within an array               |
| Large         | =LARGE(A2:A8,3)           | Returns the nth largest value within an array                |

------

### 3.RAND() & RANDBETWEEN

| Function Name | Example/Syntax      | Description                                            |
| ------------- | ------------------- | ------------------------------------------------------ |
| Rand          | =RAND()             | Returns a random value between 0 and 1 (to 15 digits)  |
| RandBetween   | =RANDBETWEEN(0,100) | Returns an integer between two values that you specify |

------

### 4.SUMPRODUCT

| Function Name | Example/Syntax                         | Description                                                  |
| ------------- | -------------------------------------- | ------------------------------------------------------------ |
| SumProduct    | =SUMPRODUCT(array1,array2,... array_N) | multiplies corresponding cells from multiple arrays and returns the sum of the products(*Note: all arrays must have the same dimensions*) |

SUMPRODUCT is often used with **filters** to calculate products only for rows that meet certain criteria, for example:

```excel
# A:Storre;B:Product;C:Quantity;D:Price
# To get revenue from apples sold at Shaws
=SUMPRODUCT((A2:A17="Shaws")*(B2:B17="Apple")*C2:C17*D2:D17)
```

**Note:**

- When you add filters to a SUMPRODUCT, you need to change the commas to **multiplication** signs.
- When you apply a condition or filter to a column, Excel translates those cells as 0's(if false) and 1's(if true); If you multiply all columns, only rows that satisfy all conditions will produce a non-zero sum.

------

### 5.COUNTIF/SUMIF/AVERAGEIF

| Function Name | Example/Syntax                           | Description                                          |
| ------------- | ---------------------------------------- | ---------------------------------------------------- |
| CountIf       | =COUNTIF(range,criteria)                 | Calculates a **count** based on specific criteria    |
| SumIf         | =SUMIF(range,criteria,sum_range)         | Calculates a **sum** based on specific criteria      |
| AverageIf     | =AVERAGEIF(range,criteria,average_range) | Calculates an **average** based on specific criteria |

**Note:**

- range: Which cells need to match your criteria?
- criteria: Under what condition do I want to sum, count, or average?
- sum/average_range: Where are the values that I want to sum or average?

| Function Name | Example/Syntax                                               | Description                                                  |
| ------------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| CountIfs      | =COUNTIFS(criteria_range1,criteria1,criteria_range2,criteria2...) | Calculates a **count** based on multiple conditions or criteria |
| SumIfs        | =SUMIFS(sum_range,criteria_range1,criteria1,criteria_range2,criteria2...) | Calculates a **sum** based on multiple conditions or criteria |
| AverageIfs    | =AVERAGEIFS(average_range,criteria_range1,criteria1,criteria_range2,criteria2...) | Calculates an **average** based on multiple conditions or criteria |

**Note:**

- If you use < or >, you need to add quatation marks as you would with text(i.e. ">200")

**Other function:**

- =COUNTBLANK(range)

