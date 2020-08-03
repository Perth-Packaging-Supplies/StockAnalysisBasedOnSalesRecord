import pandas as pd
import numpy as np
import math
import datetime

# Constants
TODAY = datetime.datetime.now()
DECIMAL_PLACES = 3


print("START")

# ALL TIME MANIPULATION WILL BE IN WEEKS
# CONFIGURATION FOR CALCULATION OF STOCK
STOCK_TARGET = [2, 2, 2]
SALES_REPORT_FILES = ["ITEMSALE-2.csv", "ITEMSALE-4.csv", "ITEMSALE-12.csv"]
SALES_REPORT_SPAN = [2, 4, 12]

# GENERATION SORTING -THE HIGHEST REPORT SPAN SHOULD BE ANALYZED FIRST
zippedCombine =sorted(zip(SALES_REPORT_SPAN,SALES_REPORT_FILES,STOCK_TARGET),reverse=True)
SALES_REPORT_SPAN, SALES_REPORT_FILES, STOCK_TARGET = zip(*zippedCombine)
summarizedStockAnalysis = None
for index, salesReportFile in enumerate(SALES_REPORT_FILES):
    # SPECIFIC LOOP CONSTANTS
    TARGET = STOCK_TARGET[index]
    REPORT_SPAN = SALES_REPORT_SPAN[index]

    # INPUT FILES
    try:
        # DATA: ITEM NUMBER, COMPANY NAME, QUANTITY SOLD (THIS IS THE MOST TIME CRITICAL FILE)
        sale = pd.read_csv("./input/{}".format(salesReportFile),skip_blank_lines=True,skiprows=1).dropna()
    except:
        print("{} - Sales Report cannot be found".format(salesReportFile))
        continue

    try:
        # DATA: ITEM NUMBER, ITEM NAME, SUPPLIER, UNITS ON HAND
        itemSummary = pd.read_excel("./input/itmls1.xlsx",skiprows=9,usecols="B:H").dropna()
    except:
        print("itmls1.xlsx - Item Summary Report cannot be found")
        continue

    try:
        # DATA: backOrder
        backOrder = pd.read_csv("./input/rpt_purchasesgeneral.csv",header=None,usecols=[32,34])
    except:
        print("rpt_purchasesgeneral.csv - Purchase Report cannot be found")
        continue

    try:
        products = pd.read_csv("./input/sets/ITEM.txt",skip_blank_lines=True,skiprows=1,sep='\t', engine='python')
    except:
        print("sets/ITEM.csv - All Products cannot be found")
        continue

    #Indices Initializations
    backOrder.rename(columns={32:"Item Number",34:"Units On Order"},inplace=True)
    backOrder.set_index("Item Number",inplace=True)

    itemSummary.rename(columns={"Item No.":"Item Number"},inplace=True)
    itemSummary.set_index("Item Number",inplace=True)

    products.set_index("Item Number",inplace=True)
    products.rename(columns={"Primary Supplier":"Supplier"},inplace=True)

    # REMOVING ADJUSTMENTS
    sale = sale[sale["Item Number"] != "\c"]
    uniqueItemCustomer = sale.drop_duplicates(["Item Number","Co./Last Name"])

    # Item Number With Associated Customers
    itemCustomers = uniqueItemCustomer.groupby("Item Number")["Co./Last Name"].apply(lambda name: ",".join(name)).reset_index()
    itemCustomers["No. Customers"] = itemCustomers["Co./Last Name"].str.count(",") + 1
    itemCustomers.rename(columns={"Co./Last Name":"Company Names"},inplace=True)
    itemCustomers.set_index("Item Number",inplace=True,drop=False)
    itemCustomers = itemCustomers.join(products[["Item Name","Supplier","Supplier Item Number"]])

    # Rearrange Columns
    itemCustomers = itemCustomers[["Item Number","Item Name", "Supplier","Supplier Item Number","Company Names","No. Customers"]]

    # Item Number With Quantity Sold
    itemQuantitySold = sale.groupby("Item Number").sum().reset_index()
    itemQuantitySold.set_index("Item Number",inplace=True,drop=False)

    stockAnalysis = itemSummary.join(itemQuantitySold,how="outer").join(backOrder)
    stockAnalysis.fillna(0,inplace=True) # Some Items are not in the joins3
    stockAnalysis.drop(columns=["Item Name", "Supplier"],inplace=True)
    stockAnalysis = stockAnalysis.join(products[["Item Name","Supplier","Supplier Item Number","Sell Unit Measure","No. Items/Buy Unit","Buy Unit Measure"]])

    # Standardize No. Weeks To Last From Quantity
    stockAnalysis["No. Weeks to Last"] = (stockAnalysis["Units On Hand"] + stockAnalysis["Units On Order"]) / (stockAnalysis["Quantity"]/REPORT_SPAN)

    deadStock = stockAnalysis.loc[stockAnalysis["No. Weeks to Last"]==np.inf]
    deadStock = deadStock[["Item Number","Item Name", "Supplier","Supplier Item Number","Units On Hand","Total Value"]]

    stockAnalysis = stockAnalysis.loc[stockAnalysis["No. Weeks to Last"]!=np.inf]
    stockToBuy = stockAnalysis.loc[stockAnalysis["No. Weeks to Last"]<=TARGET]
    stockToBuy.rename(columns={"Quantity":"Sold"},inplace=True)
    stockToBuy = stockToBuy[["Item Name", "Supplier","Supplier Item Number","Units On Hand","Units On Order","Sold","Total Value","No. Weeks to Last","Sell Unit Measure","No. Items/Buy Unit","Buy Unit Measure"]]
    stockToBuy["No. Items To Buy Per Unit"] =  stockToBuy["Sold"] - stockToBuy["Units On Hand"] - stockToBuy["Units On Order"]
    stockToBuy["No. Items To Buy Per Buying Unit"] = stockToBuy["No. Items To Buy Per Unit"] / stockToBuy["No. Items/Buy Unit"]


    # Put Units on the Operations
    stockToBuy["No. Items To Buy Per Unit"] = stockToBuy.agg("{0[No. Items To Buy Per Unit]} {0[Sell Unit Measure]}".format,axis=1)
    stockToBuy["No. Items To Buy Per Buying Unit"] = stockToBuy.agg("{0[No. Items To Buy Per Buying Unit]} {0[Buy Unit Measure]}".format,axis=1)
    stockToBuy.drop(columns=["Sell Unit Measure", "Buy Unit Measure"],inplace=True)

    # Rearrange Columns
    itemQuantitySold = stockAnalysis[["Item Name", "Supplier","Quantity"]]

    OUTPUT_FILENAME = "SA-{}-Based{}WeekSale.xlsx".format(TODAY.strftime("%m-%d-%Y"),REPORT_SPAN)
    with pd.ExcelWriter("./output/{}".format(OUTPUT_FILENAME)) as writer:
        stockAnalysis.to_excel(writer,sheet_name="Stock Analysis",index=True)
        deadStock.to_excel(writer,sheet_name="Stock Analysis - Dead Stock",index=True)
        stockToBuy.to_excel(writer,sheet_name="Stock Analysis - To Buy",index=True)
        itemCustomers.to_excel(writer, sheet_name="Item With Associated Customers",index=False)
        itemQuantitySold.to_excel(writer, sheet_name="Item Quantity Sold", index=False)

    newColumnName = "No. Weeks to Last ({})".format(REPORT_SPAN)
    stockAnalysis.rename(columns={"No. Weeks to Last": newColumnName}, inplace=True)
    if summarizedStockAnalysis is None:
        summarizedStockAnalysis = stockAnalysis
    else:
        summarizedStockAnalysis = summarizedStockAnalysis.join(stockAnalysis[[newColumnName]])

    print("Successfully Generated {}".format(OUTPUT_FILENAME))

# Difference Analysis

# Helper Functions
# Calculates average of a 1D List
def average(input):
    return sum(input) / len(input)

# Calculates Standard Deviation of a 1D List
def standardDeviation(input):
    mean = average(input)
    var = sum((x - mean) ** 2 for x in input) / len(input)
    return var ** (0.5)

summarizedStockAnalysis.fillna(0,inplace=True)
OUTPUT_FILENAME = "SA-{}-DifferenceAnalysis.xlsx".format(TODAY.strftime("%m-%d-%Y"))
VALUE_COLUMNS = ["No. Weeks to Last ({})".format(reportSpanValue) for reportSpanValue in SALES_REPORT_SPAN]

colRange = summarizedStockAnalysis.loc[:, VALUE_COLUMNS[0] : VALUE_COLUMNS[-1]]
summarizedStockAnalysis["MEAN of No. Weeks To Last"] = colRange.mean(axis=1)
summarizedStockAnalysis["STD of No. Weeks To Last"] = colRange.std(axis=1)
summarizedStockAnalysis["VAR of No. Weeks To Last"] = summarizedStockAnalysis["STD of No. Weeks To Last"] ** 2

# Individual Analysis Z-Score
for valueColumn in VALUE_COLUMNS:
    summarizedStockAnalysis["Zscore of {}".format(valueColumn)] = (summarizedStockAnalysis[valueColumn] - summarizedStockAnalysis["MEAN of No. Weeks To Last"]) / summarizedStockAnalysis["STD of No. Weeks To Last"]

with pd.ExcelWriter("./output/{}".format(OUTPUT_FILENAME)) as writer:
    summarizedStockAnalysis.to_excel(writer, sheet_name="Difference Analysis", index=True)

print("Successfully Generated {}".format(OUTPUT_FILENAME))




print("END")