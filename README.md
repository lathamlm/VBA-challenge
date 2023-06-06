# VBA-challenge
Module 2 Challenge

_____________________________________________________________________________________________________________________________________________________________
This code utilizes a For loop to identify the first and last rows of each ticker and and then uses those rows to assign values for the ticker names, opening values, and closing values.  It also translates the first and last instances of each ticker into a range value (i.e. "J3") to use in the sum function to produce total stock volume for each ticker.  

A summary table is created from the information gathered in the first For loop, and a second For loop utilizes the information in the summary table to identify the greatest increase, greatest decrease, and greatest volume as indicated in the prompt instructions.  This code also loops through each worksheet so all sheets are updated at once.

The final code runs in roughly one to three minutes in the larger Multiple_year_stock_data file
_____________________________________________________________________________________________________________________________________________________________

Resources:

#summary_row concept:  discussed in class with examples given credit_charges.xlsm

#row_count resource: https://stackoverflow.com/questions/18088729/row-count-where-data-exists by kyoya007

#cells to range resource: https://stackoverflow.com/questions/6262743/convert-cells1-1-into-a1-and-vice-versa by Anders Lindahl

#run all worksheets resource: https://stackoverflow.com/questions/43738802/how-to-apply-vba-code-to-all-worksheets-in-the-workbook by Scott Holtzman

#remove duplicates resource: https://learn.microsoft.com/en-us/office/vba/api/excel.range.removeduplicates (DIDN'T USE IN FINAL CODE - unnecessary)

#last of each ticker resouce: https://stackoverflow.com/questions/28132471/excel-vba-find-last-row-number-where-column-c-contains-a-known-value by Gary's Student (DIDN'T USE IN FINAL CODE - unnecessary)


