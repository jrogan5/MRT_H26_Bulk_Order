import pandas as pd
import numpy as np

def parseSpreadsheet(filename, sheetNum, verbose=False):
    df_perF = [] # init output df. List of df, grouped by family part

    if(verbose):
        print("Sheet Num:", sheetNum, "\n###########################")
    df = pd.read_excel(filename, sheet_name=sheetNum)
    df = df.ffill(axis=0) # forward-fill empty cells

    # assign spreadsheet columns to variable names
    partFamily = df.columns.values[0]
    manufacturerPart = df.columns.values[1]
    LCSC_Code = df.columns.values[4]
    purchaseQuantity = df.columns.values[5]
    df[df.columns[5]] = df[df.columns[5]].astype("int64") # convert purchase quantity to int

    # extract list of family names
    df_groups = df.groupby([partFamily], observed=False).head(1)
    families = np.array(df_groups[partFamily])

    for f in families:
        df_F = df[df[partFamily] == f].get([partFamily, manufacturerPart, LCSC_Code, purchaseQuantity])
        df_perF.append(df_F)
    if(verbose):
        for F in df_perF:
            print(F)
            print('------------------------------------------------')
    
    print("###########################\n")
    return df_perF

###########################

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import time

def scrapePartNumbers(df_perF, df_template, df_unavailableList, verbose=False):
    waitDuration = 1 # seconds. Don't DDOS LCSC

    df_exportList = df_template.copy() # df of cheapest parts

    for df in df_perF:
        partFamily, purchaseQuantity = df.iloc[0, [0, 3]]
        if(purchaseQuantity == 0): # 0 quantity, so skip
            continue
        partNumbers = df[df.columns[2]].drop_duplicates().astype(str).tolist()
        if(verbose):
            print("Part Family:", partFamily, "\n###########################")
            print(partNumbers)

        df = df_template.copy() # init output df for this family
        
        # loop through each part in family and web-scrape
        for partNum in partNumbers:
            if(verbose):
                print(partNum)
            if(partNum.lower() == "--"): # no LCSC code, so skip iteration
                continue
            # navigate to product page
            productUrl = f"https://www.lcsc.com/search?q={partNum}&s_z=n_{partNum}"
            driver.get(productUrl)
            time.sleep(waitDuration) # Don't DDOS LCSC

            # get LCSC stock status and available quantity
            stockStatusElementClass = driver.find_element(By.CLASS_NAME, "detailRightPanelWrap")
            # print(stockStatusElementClass.get_property("innerHTML"))
            stockStatusElement = stockStatusElementClass.find_element(By.XPATH, "//span[contains(text(), 'Stock')]")
            # print(stockStatusElement.get_property("innerHTML"))
            stockStatus = "In-Stock" if "In-Stock" in stockStatusElement.text else "Out of Stock"
            stockQuantity = stockStatusElement.text.split(":")[1].strip() if stockStatus == "In-Stock" else "0"
            stockQuantity = int(stockQuantity.strip().replace(",", ""))
            # print("Stock Status:\t" + stockStatus)
            # print("Stock Quantity:\t" + str(stockQuantity))
            if(stockStatus == "Out of Stock" or stockQuantity < purchaseQuantity): # skip iteration if out of stock or not enough stock
                continue
            # input desired purchase quantity into price calculator
            priceCalculatorElement = driver.find_element(By.CLASS_NAME, "quantityReelWrap")
            priceCalculatorInputElement = priceCalculatorElement.find_element(By.CSS_SELECTOR, "input[maxlength='9']")
            priceCalculatorInputElement.clear()  # Clear any existing text
            priceCalculatorInputElement.send_keys(str(purchaseQuantity))
            time.sleep(0.1) # Wait for input to register
            priceCalculatorInputElement.send_keys(Keys.RETURN)

            # get LCSC unit price for the given quantity
            purchaseElement = driver.find_element(By.CLASS_NAME, 'reelResultWrap')
            dataRow = purchaseElement.find_element(By.CSS_SELECTOR, '.row.mt6')
            dataColumns = dataRow.find_elements(By.CSS_SELECTOR, '.col.col-3')
            roundPurchaseQuantity = int(dataColumns[1].find_element(By.CLASS_NAME, 'major2--text').text.strip().replace(",", ""))
            # print("Purchase Quantity:\t" + str(roundPurchaseQuantity))
            unitPriceElement = dataColumns[2].find_element(By.XPATH, ".//div[contains(text(), '$')]")
            unitPrice = float(unitPriceElement.text.replace("$", "").strip())
            # print("Unit Price:\t" + str(unitPrice))

            # print("") # new line
            time.sleep(waitDuration) # Don't DDOS LCSC

            # export to df
            df.loc[len(df)] = [partFamily, partNum, stockStatus, stockQuantity, unitPrice, purchaseQuantity, roundPurchaseQuantity]

        # select cheapest part in family
        if (not df.empty) and df[df.columns[3]].notna().any():
            minPart = df.loc[df[df.columns[3]].idxmin()].to_frame().T
            df_exportList = pd.concat([df_exportList, minPart], ignore_index = True)
        else: # no parts available
            data = [partFamily, "N/A", "Out of Stock", "0", "9999", "0", "0"]
            df_unavailableList.loc[len(df_unavailableList)] = data


        # export dataframe for LCSC BOM list
        if(verbose):
            print("\nAll parts:\n", df)
            print("\nCheapest:\n", df_exportList, "\n###########################\n")

    return df_exportList

###########################



startTime = time.time()



############# Program Settings ##############

verbose = False

sourceFilename = r'LCSC BOM H26.xlsx'
outputFilename = r'MRT_H26_Nov14_LCSCBulkOrder.csv'
unavailableFilename = r'unavailable.csv'

sheetStart = 1 # Must skip sheet 1 (info sheet)
totalNumSheets = len(pd.ExcelFile(sourceFilename).sheet_names) # determine total number of sheets
sheetEnd = totalNumSheets
# sheetEnd = 2 # override if desired

# output df template. Check code if you modify this, since other parts may break
df_template = pd.DataFrame(columns=['Part Family', 'LCSC Code',
                                        'Stock Status', 'Stock Quantity', 'Unit Price',
                                        'Purchase Quantity', 'Rounded Purchase Quantity'])

############# Program Settings ##############



# delete if old output files exist
import os
if os.path.exists(outputFilename):
    os.remove(outputFilename)
if os.path.exists(unavailableFilename):
    os.remove(unavailableFilename)

# start Selenium
service = Service(executable_path=r'./driver/geckodriver.exe', log_output=r"./driver/gecko_log.txt")
driver = webdriver.Firefox(service=service)

for sheetNum in range(sheetStart, totalNumSheets): 
    df_perF = parseSpreadsheet(sourceFilename, sheetNum, verbose=verbose)

    df_unavailableList = df_template.copy()
    df_exportList = scrapePartNumbers(df_perF, df_template, df_unavailableList, verbose=verbose)

    # export dataframe for LCSC BOM list. Append to existing csv
    print("Export List:\n", df_exportList, "\n")
    print("Out of Stock List:\n", df_unavailableList, "\n###########################\n")
    df_exportList.to_csv(outputFilename, mode='a', header=False, index=False)
    df_unavailableList.to_csv(unavailableFilename, mode='a', header=False, index=False)

    sheetNum += 1

# stop Selenium (close the browser)
driver.quit()
print("Total execution time:", time.time() - startTime)