import pandas as pd
import numpy as np
import math
import datetime

# Constants
TODAY = datetime.datetime.now()
DECIMAL_PLACES = 3

# INPUT FILES

try:
    # DATA: ITEM NUMBER, COMPANY NAME, QUANTITY SOLD
    sale = pd.read_csv("./input/ITEMSALE.csv",skip_blank_lines=True,skiprows=1).dropna()
except:
    print("ITEMSALE.csv - Sales Report cannot be found")

try:
    # DATA: ITEM NUMBER, ITEM NAME, SUPPLIER, UNITS ON HAND
    itemSummary = pd.read_excel("./input/itmls1.xlsx",skiprows=9,usecols="B:H").dropna()
except:
    print("itmls1.xlsx - Item Summary Report cannot be found")

itemSummary.rename(columns={"Item No.":"Item Number"},inplace=True)
itemSummary.set_index("Item Number",inplace=True)

# REMOVING ADJUSTMENTS
sale = sale[sale["Item Number"] != "\c"]
uniqueItemCustomer = sale.drop_duplicates(["Item Number","Co./Last Name"])

# Item Number With Associated Customers
itemCustomers = uniqueItemCustomer.groupby("Item Number")["Co./Last Name"].apply(lambda name: ",".join(name)).reset_index()
itemCustomers["No. Customers"] = itemCustomers["Co./Last Name"].str.count(",") + 1
itemCustomers.rename(columns={"Co./Last Name":"Company Names"},inplace=True)
itemCustomers.set_index("Item Number",inplace=True,drop=False)
itemCustomers = itemCustomers.join(itemSummary[["Item Name","Supplier"]])

# Rearrange Columns
itemCustomers = itemCustomers[["Item Number","Item Name", "Supplier","Company Names","No. Customers"]]

# Item Number With Quantity Sold
itemQuantitySold = sale.groupby("Item Number").sum().reset_index()
itemQuantitySold.set_index("Item Number",inplace=True,drop=False)

with pd.ExcelWriter("./output/StockAnalysis.xlsx") as writer:
    itemCustomers.to_excel(writer, sheet_name="Item With Associated Customers",index=False)
    itemQuantitySold.to_excel(writer,sheet_name="Item Quantity Sold",index=False)
