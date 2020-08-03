# MYOB with WMS Stock Analyzer Based on Sales (Stock Analysis)
Video Documentation Link: https://youtu.be/UavIboqfViA

# What is it?
This is a program that will analyze the following:
- The list of stocks that will need to be ordered to maintain the supply based on the record. For example, if your analysis on 1 month, the program will determine the number of items to maintain 1 month of stock.
- The list of stocks that are considered "dead stocks", in other words, stocks that exists in the warehouse, but no supply chain established on the product
- The details on each stocks that are currently in the warehouse and the duration to last.
- Name and count of customers that are associated with a specific stock

# How To Use
The program will require 6 main inputs, 5 of which is from MYOB, and 1 of which is from WMS.
*Note: The name of the file has to match the name specified here for the program to work. Otherwise you will just be notified for missing files. The names of the file are based on the default name of the file when exported for convenience.*

## Installation
1. Make sure that you have python installed in your computer
2. Download this file, and place it where it pleases

## Input/Setup
*MYOB Input files from Import/Export Assistance has an extra row that has the content "{}" whenever being exported. The python script can handle this, so* **DO NOT CHANGE THE FILE***

### 3 Sales Report ([ITEMSALE-2.csv,ITEMSALE-4.csv,ITEMSALE-12.csv]) [MYOB]
Data that comes from here are related to sales/ items sold.
1. Locate the file - This can be found in MYOB:  File (Top-left corner) >> Import/Export Assistance
2. Select **Export Data**. Press Next
3. Select "Sales" as the file to export.
4. Select "Item Sales" as the type of sales.
5. Select the date range (recommended is [2 weeks, 4 weeks, 12 weeks] before up to Today). Press Next
6. Select the seperate data using "Commas", and Make sure the field "Include field headers in the file" is checked. Press next
7. Export the fields:
   1. Item Number
   2. Quantity
   3. Co./Last Name
   4. Item Description
8. Press Next to export and save as "ITEMSALE-#.csv" inside the input folder of the program replacing "#" with number of weeks (2, 4 or 12). (You may need to change the save as type to "All Files", and change the format name)

### Inventory Report (itmls1.xlsx) [MYOB]
Data that comes from here are related to the current state of the inventory/warehouse
1. Locate the file - This can be found in MYOB: Reports (Bottom Middle) >> Inventory >> Items >> Item  List [Summary]
2. Press Export To Excel
3. Save excel file as "itmls1.xlsx" inside the input folder of the program.

### Purchase Report (rpt_purchasesgeneral.csv) [WMS]
Data that comes from here are related to backorder to ensure that the items that will be ordered are not yet ordered before
1. Locate the file - This can be found in WMS: Reports >> Purchases >> Purchase Summary >> By Purchase Order
2. Tweak the settings as you please, but the most important are the start and end date (recommmended to be around 3 months before today upto Today)
3. Report Type should be detail
4. Under the advanced tab, the check field in Receive Status "Open" and "Split" should be checked.
5. Press Generate Report
6. Click Export Report on the icon that can be found in the top-left corner.
7. Change the format of the file to "csv".
8. Save the file as "rpt_purchasesgeneral.csv" inside the input folder of the program.

### Item Details report  (ITEM.txt) [MYOB]
This file does not need to be changed often if there are no new product codes in the system. This essentially contains extra information on the items in the warehouse such as the supplier, description, the amount of selling units per buying units.With this specific type of export in CSV, MYOB sometimes put "," in terms the fields, so this is better to be exported as txt file to be separated by tabs,
1. Locate the file - This can be found in MYOB:  File (Top-left corner) >> Import/Export Assistance
2. Select **Export Data**. Press Next
3. Select "Items" as the file to export.
4. Select the seperate data using "tabs", and Make sure the field "Include field headers in the file" is checked. Press next
5. Export All fields
6. Press Next to export and save as "ITEM.txt" inside the input/sets folder of the program. (You may need to change the save as type to "All Files", and change the format name)

## Execution
After finishing the setup - exporting of the files. You can start the program either by running "start.bat" or typing in the terminal that is pointed inside the directory of this file "python main.py"

## Output
The execution of the program will output a files called "SA-[DateToday]-[Type].xlsx" that can be found inside the output folder of the program. The file will have the date today and the type:
- Based2WeekSale
- Based4WeekSale
- Based12WeekSale
- DifferenceAnalysis


The "Based[NoWeeks]WeekSale" will determine the stock analysis based on **NoWeeks** (no of weeks) sales record. The **Difference Analysis** will analyze the different NoWeeks sale to determine the variation and helps in picking the correct type for a specific item.

