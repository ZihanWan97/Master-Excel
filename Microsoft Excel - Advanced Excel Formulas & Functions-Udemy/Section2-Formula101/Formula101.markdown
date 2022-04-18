## Formula101

### 1.Formula Syntax

| Syntax                                       | Example                      | Output | Description                                          |
| -------------------------------------------- | ---------------------------- | ------ | ---------------------------------------------------- |
| =YEAR(serial_number)                         | =YEAR("2/22/2022")           | 2022   | Get year of date                                     |
| =LEFT(text, [num_chars])                     | =LEFT("Hello", 2)            | He     | Get characters of text from left                     |
| SEARCH(find_text, within_text, [start_num])) | =SEARCH("@","wzh@gmail.com") | 4      | Get the order of the specified character in the text |

- How to get username from email?

  - =LEFT("wzh@gmail.com",SEARCH("@","wzh@gmail.com")-1)
  - Output: wzh

------

### 2.Reference Types

- Cells references are **relative** by default(**A1**). This allows the reference to change as the formula is copies to new cells.

- The **$** symbol is used to create **fixed** references. You can fix entire cells(**$A$1**) or just the column(**$A1**) or row(**A$1**), which prevents references from changing as the formula is copied to new cells.


------

### 3.Common Error Types

| Error Type | Description                                  | Solution                                                     |
| ---------- | -------------------------------------------- | ------------------------------------------------------------ |
| ######     | Column isn't wide enough to display values   | Drag or double-click column border to increase width, or right-click to set custom column width |
| #NAME?     | Excels does not recognize text in a formula  | Make sure that function names are correct, references are valid and spelled properly, and quotation marks and colons are in place |
| #VALUE!    | Formula has the wrong type of argument       | Check that your formula isn’t trying to perform an arithmetic operation on text strings or cells formatted as text |
| #DIV/0!    | Formula is dividing by zero or an empty cell | Check the value of your divisor; if 0 is correct, use an IF statement to display an alternate value if you choose |
| #REF!      | Formula refers to a cell that is not valid   | Make sure that you didn’t move, delete, or replace cells that are referenced in your formula |
| #N/A       | Formula can't find a referenced value        | Check that all references and formula arguments evaluate properly (the most common cause is a lookup value with no match) |

------

### 4.Auditing Tools

Steps: "**Formulas**" menu ---> "**Formula Auditing**" sub-menu ---> Tools(Buttons) below:

| Button Name      | Function                                                     |
| ---------------- | ------------------------------------------------------------ |
| Trace Precedents | Identifies cells which ***affect*** the value of the one selected. |
| Trace Dependents | Identifies cells which ***are affected by*** the value of the one selected. |
| Show Formulas    | Temporarily displays all formulas in the worksheet as ***text strings**.* |
| Remove Arrows    | Remove the arrows drawn by trace precedents or trace dependents. |
| Evaluate Formula | Allows you to cycle through ***each individual calculation step*** within a formula, see how each component evaluates, and ***pinpoint the source of the error***. |
| Error Checking   | Scans the sheet for errors and provides a summary with ***options*** to trace the source, ignore the error, modify your options, or link out to Microsoft support. |

------

### 5.Ctrl Shortcuts

| Shortcuts            | Function                                                     |
| -------------------- | ------------------------------------------------------------ |
| Ctrl + Arrow         | ***Jumps to the last cell*** in a data region, in the direction of the arrow. |
| Ctrl + Shift + Arrow | ***Select to the last cell*** in a data region, in the direction of the arrow. |
| Ctrl + Home/End      | Jumps to the ***Home(top-left)*** or ***End(bottom_right)*** cell in a region. |
| Ctrl + .             | Jumps straight to ***each corner*** within a selected cell range. |
| Ctrl + PgUp/PgDn     | Switches worksheet tabs, either to the ***left(PgUp)*** or ***right(PgDn)***. |

> Note: Excel determines **a data region** based on "**Blanks**"

------

### 6.Function Shortcuts

| Shortcuts | Function                                                     |
| --------- | ------------------------------------------------------------ |
| F1        | Launches the ***Excel help*** pane(default); Links to the ***Microsoft Support*** websites(tool-specific). |
| F2        | Allows you to ***edit*** the active cell; Highlights cells referenced by the active formula. |
| F4        | Repeats the ***last action*** taken(default); ***Toggles*** absolute/relative cell references within a formula**(A1->$A$1->A$1->$A1)**. |
| F9        | Calculates all workbook formulas(when in ***manual*** mode); Evaluates ***each function argument*** within the formula bar. |

[🌐Click here to get full list of keyboard shortcuts](https://exceljet.net/keyboard-shortcuts)

------

### 7.Alt Key Tips

- Allows you to **quickly access** tools from ribbon menus and sub-menus using **only the keyboard**.
- Each keystroke takes you **a layer deeper**, until you land on the tool you need.
- Focus on the **combinations(tools)** that you use **most frequently**.

------

### 8.Data Validation

Steps: "**Data**" menu ---> "**Data Tools**" sub-menu ---> "**Data Validation**" button

**Restricts the values** that a user can enter into **a given cell**, based on:

- Number Type(Whole vs. Decimal)
- Value(Between, Less than, Equal to, etc)
- List of Items(Based on cell range or manual list)
- Date/Time(Between, Less than, Equal to, etc)
- Text Length(Between, Less than, Equal to, etc)
- Custom(Formula-Driven)
