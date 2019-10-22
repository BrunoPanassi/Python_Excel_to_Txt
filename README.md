# Formatting data from Excel to .txt with Python
A algorithm that takes data from a Excel file to make a formatted text in a .txt file using Python.
This project can also be seen as an example of an improvement in a code.
# What is it and what for it?
A few time ago, a need arose on my work, a client ask us monthly a file which has separated formatted information in a specific range of columns, for financial purposes.
To build this file manually takes around 2-4 hours of work, and you still risk the information not being in the right places.
# The solution
For this, i create a Python algorithm that does all of this automatically, with the extraction and manipulation of data to put in their respective columns respecting the rules of building the file.
### The 1º Algorithm
The 1º algorithm is the [FechCont.py](https://github.com/BrunoPanassi/Python_Excel_to_Txt/blob/master/FechCont.py) builded with the **Jupyter Notebook**, create the file correctly, but in a Excel file with **16.000** rows takes around **20** minutes to build. Much as faster than building manually, the algorithm had some improvements to be made.
### The 2º Algorithm
The [FechCont_V2.py](https://github.com/BrunoPanassi/Python_Excel_to_Txt/blob/master/FechCont_V2.py) file was also builded in **Jupyter Notebook** and had some adjustments in **VS Code**.
Fixing the bugs of the 1º algorithm and having some few lines less, the algorithm process a Excel file with **16.000** rows in **0,10** seconds, having a 99% improvement in time.
### What were the changes made?
The improvements that i did in the 2º algorithm:
* **The Excel File Load** - 1º Algorithm: The load is made with no columns selected, besides being used two diferent functions to load the Excel file (*ExcelFile* and *read_excel*) | 2º Algorithm: The load is made only by *read_excel* giving the sheet name.
* **The Selected Columns** - 1º Algorithm: The selected columns is given by a list with the related indexes of each column. | 2º Algorithm: A new Data Frame is created ordering the columns giving their related names.
* **1º Column Validation** - 1º Algorithm: Here i validate if has diference with the 1º and 2º column value of the last row to apply the rule and has a loop to write all the zeros before writing the value. | 2º Algorithm: The validations about the size of the value it is made in fewer lines and there is no need to unnecessary values convertions. The loop to fill the value with zeros at the left it is made with the function *zfill*.
* **4º Column Validation** - 1º Algorithm: Has two loops to convert the value to float and fill the value with zeros at left. | 2º Algorithm: Both loops are replaced by the functions *float*, *format* and *zfill* reducing of 21 to 5 lines of code.
* **5º Column Validation** - 1º Algorithm: If the value has the right size it is filled with a right value of spaces after the value, and a loop made this. | 2º Algorithm: Only one concatenation is required.
* **10º Column Validation** - 1º Algorithm: To fill in the missing spaces a loop it is made. | 2º Algorithm: Besides the loop begin replace by the function *rjust*, there is a validation to leave the value with a fixed size with the function *format*, even if the size exceeds this value.
* **11º Column Validation** - 1º Algorithm: The validation here it is made in the 1º column, taking the values of last row to compare | 2º Algorithm: Since it´s the last column to check, it has to be applied the last rule, so the rule to compare values between the rows, it is made comparing the value of the next row.
# Conclusion
From a need to automate a task to a review and changes to improve a code, has be seen that some points need of an attention when is coding a algorithm.
