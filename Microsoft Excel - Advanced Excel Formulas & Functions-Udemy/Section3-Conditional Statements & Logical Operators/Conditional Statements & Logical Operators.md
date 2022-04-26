## Conditional Statements & Logical Operators

### 1.IF Statement

```excel
=IF(logical_test,[Value if True],[Value if False])
```

-------

### 2.Nested IF Statement

```excel
=IF(logical_test_1,[Value if True],IF(logical_test_2,[Value if True],[Value if False]))
```

- By using Nested IF Statements, you can include multiple logical tests within a single formula

------

### 3.Conditional AND/OR Operators

**AND** and **OR** statements allow you to include multiple logical tests at once.

```excel
=IF(OR(logical_test_1,logical_test_2),[Value if OR True],[Value if OR False])
```

```excel
=IF(AND(logical_test_1,logical_test_2),[Value if AND1 True],IF(AND(logical_test_3,logical_test_4),[Value if AND2 True],[Value if AND2 False]))
```

------

### 4.NOT & <>

If you want to evaluate a case where a logical statement is not true, you can use either the **NOT** statement or a **<>** operator.

```excel
=IF(NOT(logical_test),[Value if True],[Value if False])
```

```excel
=IF(...<>...,[Value if True],[Value if False])
```

------

### 5.IFERROR

The **IFERROR** statement is an excellent tool to eliminate annoying error messages(**#N/A, #DIV/0!, #REF!, etc.**), which is particularly useful for front-end formatting.

```excel
=IFERROR(value_or_formula,value_if_error)
```

------

### 6.IS Statement

Excel offers a number of different **IS formulas**, each of which checks whether a certain conditions is true.

| Formula   | Description                                                  |
| --------- | ------------------------------------------------------------ |
| ISBLANK   | Checks whether the reference cell or value is ***blank***    |
| ISNUMBER  | Checks whether the reference cell or value is ***numerical*** |
| ISTEXT    | Checks whether the reference cell or value is ***a text string*** |
| ISERROR   | Checks whether the reference cell or value ***returns an error*** |
| ISEVEN    | Checks whether the reference cell or value is ***even***     |
| ISODD     | Checks whether the reference cell or value is ***odd***      |
| ISLOGICAL | Checks whether the reference cell or value is ***a logical operator*** |
| ISFORMULA | Checks whether the reference cell or value is  ***a formula*** |
