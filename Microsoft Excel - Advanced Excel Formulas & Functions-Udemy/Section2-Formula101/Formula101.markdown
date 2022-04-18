## Formula101

### Formula Syntax

| Syntax                                       | Example                      | Output | Description                                          |
| -------------------------------------------- | ---------------------------- | ------ | ---------------------------------------------------- |
| =YEAR(serial_number)                         | =YEAR("2/22/2022")           | 2022   | Get year of date                                     |
| =LEFT(text, [num_chars])                     | =LEFT("Hello", 2)            | He     | Get characters of text from left                     |
| SEARCH(find_text, within_text, [start_num])) | =SEARCH("@","wzh@gmail.com") | 4      | Get the order of the specified character in the text |

- How to get username from email?

  - ```excel
    =LEFT("wzh@gmail.com",SEARCH("@","wzh@gmail.com")-1)
    ```

  - Output: wzh

------

### Reference Types

Cells references are **relative** by default(**A1**). This allows the reference to change as the formula is copies to new cells.

The **$** symbol is used to create **fixed** references. You can fix entire cells(**$A$1**) or just the column(**$A1**) or row(**A$1**), which prevents references from changing as the formula is copied to new cells.

------

### Common Error Types

| Error Type | Description                                  | Solution                                                     |
| ---------- | -------------------------------------------- | ------------------------------------------------------------ |
| ######     | Colum isn't wide enough to display values    | Drag or double-click column border to increase width, or right-click to set custom column width |
| #NAME?     | Excels does not recognize text in a formula  | Make sure that function names are correct, references are valid and spelled properly, and quotation marks and colons are in place |
| #VALUE!    | Formula has the wrong type of argument       | Check that your formula isn’t trying to perform an arithmetic operation on text strings or cells formatted as text |
| #DIV/0!    | Formula is dividing by zero or an empty cell | Check the value of your divisor; if 0 is correct, use an IF statement to display an alternate value if you choose |
| #REF!      | Formula refers to a cell that is not valid   | Make sure that you didn’t move, delete, or replace cells that are referenced in your formula |
| #N/A       | Formula can't find a referenced value        | Check that all references and formula arguments evaluate properly (the most common cause is a lookup value with no match) |
