import pandas as pd
import numpy as np
import math
import datetime

# Constants
TODAY = datetime.datetime.now()
DECIMAL_PLACES = 3


print("START")
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

try:
    # DATA: backOrder
    backOrder = pd.read_csv("./input/rpt_purchasesgeneral.csv",header=None,usecols=[32,34])
except:
    print("rpt_purchasesgeneral.csv - Purchase Report cannot be found")

try:
    products = pd.read_csv("./input/sets/products.csv",skip_blank_lines=True,skiprows=0)
except:
    print("sets/products.csv - All Products cannot be found")

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
stockAnalysis = stockAnalysis.join(products[["Item Name","Supplier","Supplier Item Number"]])
stockAnalysis["No. Months To Last"] = (stockAnalysis["Units On Hand"] + stockAnalysis["Units On Order"]) / stockAnalysis["Quantity"]

deadStock = stockAnalysis.loc[stockAnalysis["No. Months To Last"]==np.inf]
deadStock = deadStock[["Item Number","Item Name", "Supplier","Supplier Item Number","Units On Hand","Total Value"]]

stockAnalysis = stockAnalysis.loc[stockAnalysis["No. Months To Last"]!=np.inf]
stockToBuy = stockAnalysis.loc[stockAnalysis["No. Months To Last"]<=0.5]
stockToBuy.rename(columns={"Quantity":"Sold"},inplace=True)
stockToBuy = stockToBuy[["Item Name", "Supplier","Supplier Item Number","Units On Hand","Units On Order","Sold","Total Value","No. Months To Last"]]
stockToBuy["No. Items To Buy"] =  stockToBuy["Sold"] - stockToBuy["Units On Hand"] - stockToBuy["Units On Order"]

# Rearrange Columns
itemQuantitySold = stockAnalysis[["Item Name", "Supplier","Quantity"]]


with pd.ExcelWriter("./output/StockAnalysis.xlsx") as writer:
    itemCustomers.to_excel(writer, sheet_name="Item With Associated Customers",index=False)
    itemQuantitySold.to_excel(writer,sheet_name="Item Quantity Sold",index=False)
    stockAnalysis.to_excel(writer,sheet_name="Stock Analysis",index=True)
    deadStock.to_excel(writer,sheet_name="Stock Analysis - Dead Stock",index=True)
    stockToBuy.to_excel(writer,sheet_name="Stock Analysis - To Buy",index=True)

print("END")