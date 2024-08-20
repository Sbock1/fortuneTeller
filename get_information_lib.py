from alpha_vantage.fundamentaldata import FundamentalData
from alpha_vantage.timeseries import TimeSeries
import matplotlib.pyplot as plt
import seaborn as sb
import pandas as pd
import numpy as np
import datetime as dt
import time
import sqlite3 as sql
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.utils.dataframe import dataframe_to_rows

import tkinter as tk
import os

################################################################## 
# Global variables based on API requirements
API_key="HNNMDBOG55BREC5P" #Account API key to validate access
time_delay = 1 #The free API access only allows 25 requests per day without any delay, but for test
FundamentalDataAPIPull = FundamentalData(API_key,output_format='pandas')
TimeSeriesPull = TimeSeries(key=API_key, output_format="pandas")
# IMPORTANT:
# The AlphaVantage API documentation can be found here:
# https://www.alphavantage.co/documentation/
##################################################################


##Update first row -> needs to be called separately, not included yet -> TO-DO: Better implementation
def insertFirstRowColumnNames():
    firstRowData = FundamentalDataAPIPull.get_company_overview("A")[0]
      
    # Adding of additional columns at the end
    firstRowData["Downloaded_at_datetime"] = timestampToday()
    firstRowData["Downloaded_at_quarter"] = timestampQuarter()

    # Path of file
    excel_path = "output.xlsx"

    # Check if file exists
    if not os.path.exists(excel_path):
        wb = Workbook()
        wb.save(excel_path)

    book = load_workbook(excel_path)

    # Check if worksheet exists
    if "Overview" not in book.sheetnames:
        balanceSheet = book.create_sheet("Overview")
    else:
        balanceSheet = book["Overview"]

    for row in dataframe_to_rows(firstRowData, index=False, header=True):
        balanceSheet.append(row)

    # Speichere die Änderungen
    book.save(excel_path)

def insertFirstRowColumnNamesBalanceQuarterly():
    # Load balance data from Alpha Vantage
    stock = "A"
    firstRowData = FundamentalDataAPIPull.get_balance_sheet_quarterly(stock)[0]
    
    # Adding of additional columns at the end
    firstRowData["Downloaded_at_datetime"] = timestampToday()
    firstRowData["Downloaded_at_quarter"] = timestampQuarter()

    # Check and if not Dataframe datatype, convert to Dataframe
    if not isinstance(firstRowData, pd.DataFrame):
        firstRowData = pd.DataFrame([firstRowData])

    firstRowData.insert(0, "Symbol", stock)

    # Path of file
    excel_path = "output.xlsx"

    # Check if file exists
    if not os.path.exists(excel_path):
        wb = Workbook()
        wb.save(excel_path)

    book = load_workbook(excel_path)

    # Check if worksheet exists
    if "Balance_Quarterly" not in book.sheetnames:
        balanceSheet = book.create_sheet("Balance_Quarterly")
    else:
        balanceSheet = book["Balance_Quarterly"]

    for row in dataframe_to_rows(firstRowData, index=False, header=True):
        balanceSheet.append(row)

    # Speichere die Änderungen
    book.save(excel_path)
    APIRequestDelay()


def insertFirstRowColumnNamesBalanceYearly():
    # Load balance data from Alpha Vantage
    stock = "A"
    firstRowData = FundamentalDataAPIPull.get_balance_sheet_annual(stock)[0]
    
    # Adding of additional columns at the end
    firstRowData["Downloaded_at_datetime"] = timestampToday()
    firstRowData["Downloaded_at_quarter"] = timestampQuarter()

    # Check and if not Dataframe datatype, convert to Dataframe
    if not isinstance(firstRowData, pd.DataFrame):
        firstRowData = pd.DataFrame([firstRowData])

    firstRowData.insert(0, "Symbol", stock)

    # Path of file
    excel_path = "output.xlsx"

    # Check if file exists
    if not os.path.exists(excel_path):
        wb = Workbook()
        wb.save(excel_path)

    book = load_workbook(excel_path)

    # Check if worksheet exists
    if "Balance_Yearly" not in book.sheetnames:
        balanceSheet = book.create_sheet("Balance_Yearly")
    else:
        balanceSheet = book["Balance_Yearly"]

    for row in dataframe_to_rows(firstRowData, index=False, header=True):
        balanceSheet.append(row)

    # Speichere die Änderungen
    book.save(excel_path)
    APIRequestDelay()

def insertFirstRowColumnNamesIncomeQuarterly():
    # Load balance data from Alpha Vantage
    stock = "A"
    firstRowData = FundamentalDataAPIPull.get_income_statement_quarterly(stock)[0]
    
    # Adding of additional columns at the end
    firstRowData["Downloaded_at_datetime"] = timestampToday()
    firstRowData["Downloaded_at_quarter"] = timestampQuarter()

    # Check and if not Dataframe datatype, convert to Dataframe
    if not isinstance(firstRowData, pd.DataFrame):
        firstRowData = pd.DataFrame([firstRowData])

    firstRowData.insert(0, "Symbol", stock)

    # Path of file
    excel_path = "output.xlsx"

    # Check if file exists
    if not os.path.exists(excel_path):
        wb = Workbook()
        wb.save(excel_path)

    book = load_workbook(excel_path)

    # Check if worksheet exists
    if "Income_Quarterly" not in book.sheetnames:
        balanceSheet = book.create_sheet("Income_Quarterly")
    else:
        balanceSheet = book["Income_Quarterly"]

    for row in dataframe_to_rows(firstRowData, index=False, header=True):
        balanceSheet.append(row)

    # Speichere die Änderungen
    book.save(excel_path)
    APIRequestDelay()


def insertFirstRowColumnNamesIncomeYearly():
    # Load balance data from Alpha Vantage
    stock = "A"
    firstRowData = FundamentalDataAPIPull.get_income_statement_annual(stock)[0]
    
    # Adding of additional columns at the end
    firstRowData["Downloaded_at_datetime"] = timestampToday()
    firstRowData["Downloaded_at_quarter"] = timestampQuarter()

    # Check and if not Dataframe datatype, convert to Dataframe
    if not isinstance(firstRowData, pd.DataFrame):
        firstRowData = pd.DataFrame([firstRowData])

    firstRowData.insert(0, "Symbol", stock)

    # Path of file
    excel_path = "output.xlsx"

    # Check if file exists
    if not os.path.exists(excel_path):
        wb = Workbook()
        wb.save(excel_path)

    book = load_workbook(excel_path)

    # Check if worksheet exists
    if "Income_Yearly" not in book.sheetnames:
        balanceSheet = book.create_sheet("Income_Yearly")
    else:
        balanceSheet = book["Income_Yearly"]

    for row in dataframe_to_rows(firstRowData, index=False, header=True):
        balanceSheet.append(row)

    # Speichere die Änderungen
    book.save(excel_path)
    APIRequestDelay()


def insertFirstRowColumnNamesStockDaily():
    # Load balance data from Alpha Vantage
    stock = "A"
    firstRowData = TimeSeriesPull.get_daily(stock, outputsize="full")[0]
    
    # Adding of additional columns at the end
    firstRowData["Downloaded_at_datetime"] = timestampToday()
    firstRowData["Downloaded_at_quarter"] = timestampQuarter()

    # Check and if not Dataframe datatype, convert to Dataframe
    if not isinstance(firstRowData, pd.DataFrame):
        firstRowData = pd.DataFrame([firstRowData])

    # Convert index to a column to preserve the datetime of the actual stock value
    firstRowData.reset_index(inplace=True)

    # Rename the column to Date
    firstRowData.rename(columns={'index':'Date'}, inplace=True)
    firstRowData["date"] = firstRowData["date"].dt.date

    firstRowData.insert(0, "Symbol", stock)

    # Path of file
    excel_path = "output.xlsx"

    # Check if file exists
    if not os.path.exists(excel_path):
        wb = Workbook()
        wb.save(excel_path)

    book = load_workbook(excel_path)

    # Check if worksheet exists
    if "Stocks_Daily" not in book.sheetnames:
        balanceSheet = book.create_sheet("Stocks_Daily")
    else:
        balanceSheet = book["Stocks_Daily"]

    for row in dataframe_to_rows(firstRowData, index=False, header=True):
        balanceSheet.append(row)

    # Speichere die Änderungen
    book.save(excel_path)
    APIRequestDelay()


def deleletExcelPredefinedSheet():
    # Path of file
    excel_path = "output.xlsx"

    # Check if file exists
    if not os.path.exists(excel_path):
        wb = Workbook()
        wb.save(excel_path)

    book = load_workbook(excel_path)

    # Check if predefined worksheet exists and delete it
    if "Tabelle1" in book.sheetnames:
        del book["Tabelle1"]
    elif "Sheet1" in book.sheetnames:
        del book["Sheet"]

    # Speichere die Änderungen
    book.save(excel_path)


# Reads any column of an excel worksheet and gives out a list of the contents of the rows of that worksheet
def loadOneColumnRowDataAsList(filename, sheetname, column):
    listFromColumn = pd.read_excel(filename, sheet_name=sheetname, usecols=column, header=None, index_col=0) #Read column A / No header - Take first row NOT as replacement header / No index / returns Dataframe
    listFromColumn = listFromColumn.index.tolist() # Convert Dataframe to list in order for easier processing later / use index as the column A is seen as index column
    return listFromColumn


# Compare mainList and deleteList, delete the items in mainList, which are not in deleteList, and return modified mainList
# Compare mainList and addList, add the items to mainList, which are in addList, and return modified mainList
# The goal of the function is helping with having a uptodate SymbolList
def compareDeleteAddListWithMainList(mainList, deleteList=None, addList=None):
    if deleteList == None:
        deleteList = []
    if addList == None:
        addList =[]
    
    # KEEP IN MIND: First the deletion is done, THEN the addition!
    mainList = [item for item in mainList if item not in deleteList]
    mainList.extend([item for item in addList if item not in mainList])
    mainList.sort()

    return mainList


def getExcelSheetInformation(filename, sheetname):
    mainSheetDataframe = pd.read_excel(filename, sheetname)


    excelSymbolsExisting = list(set(mainSheetDataframe["Symbol"]))
    excelQuartersExisting = list(mainSheetDataframe["Downloaded_at_quarter"])
    
    return mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting



def checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting):
    # This func checks three things:
    # 1. Which symbols (stocks) are already in the main excel written
    # 2. Which symbols are already in the main excel existing, but the download quarter is old, and they needRefresh
    # 3. Which symbols do not exist anymore, meaning which companies are not longer on the stock exchange registered
    # This func does not eliminate any stock from the main stock list, but it gives the information to Writer functions
    existingSymbols = []
    symbolsNeedRefresh = []
    updateNotExistingSymbols = []
    
    if excelSymbolsExisting != []:
        for symbol in symbolList:
            if symbol in excelSymbolsExisting:
                rowOnWorksheet = excelSymbolsExisting.index(symbol)
                if timestampQuarter() == excelQuartersExisting[rowOnWorksheet]:
                    existingSymbols.append(symbol + " in" + str(timestampQuarter()) +" uptodate")
                else:
                    symbolsNeedRefresh.append([symbol, rowOnWorksheet])
            else:
                updateNotExistingSymbols.append(symbol)
                
        return existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols

    else:
        updateNotExistingSymbols=symbolList

        return existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols



def APIRequestDelay():
    time.sleep(time_delay)

    return None


def timestampToday():
    # Get the current datetime
    today = pd.to_datetime(dt.datetime.now())

    return today


def timestampQuarter():
    # Get the current quarter
    today = pd.to_datetime(dt.datetime.today())
    currentQuarter = (today.month - 1) // 3 + 1

    return currentQuarter


def identifyLastColumnWithContents(DataFrame):
    # To identify the last column, the DataFrame is given as input, than it is check which cells hold at not a NaN value
    # than the columns with at least one not NaN value are marked true. The last one of these is the output translated into a number,
    # this number is returned by this function.
    lastColumnWithContentsName = DataFrame.notna().any().index[-1]
    lastColumnWithContentsNumber = DataFrame.columns.get_loc(lastColumnWithContentsName)

    return lastColumnWithContentsNumber


def updateCompanyOverview(updateNotExistingSymbols):
    stocksNotExisting = []
    allStockOverviewData = pd.DataFrame()
    counter = 0

    for elem in updateNotExistingSymbols:
        try:
            #the API command get_company_overview retrives stock data like Revenue, Cashflow, Ebit, etc. 
            stockOverviewData = FundamentalDataAPIPull.get_company_overview(elem)[0] 
            #The function identifyLastColumnWithContents checks the number of the last column with content in it, then two columns with timestamps are appended
            stockOverviewData["Downloaded_at_datetime"] = timestampToday()
            stockOverviewData["Downloaded_at_quarter"] = timestampQuarter()

            # Check and if not Dataframe datatype, convert to Dataframe
            if not isinstance(stockOverviewData, pd.DataFrame):
                stockOverviewData = pd.DataFrame(stockOverviewData)

            allStockOverviewData = pd.concat([allStockOverviewData, stockOverviewData], ignore_index=True)

            APIRequestDelay()
        except ValueError:
            stocksNotExisting.append(elem)
            print("Error" + str(counter) + " " + str(elem))
            counter = counter + 1
            APIRequestDelay()
            pass
    
    return allStockOverviewData, stocksNotExisting

def update_balance_sheet_quarterly(updateNotExistingSymbols):
    stocksNotExisting = []
    counter = 0
        
    for stock in updateNotExistingSymbols:
        
        try:
            #the API command get_company_overview retrives stock data like Revenue, Cashflow, Ebit, etc. 
            stockBalanceData = FundamentalDataAPIPull.get_balance_sheet_quarterly(stock)[0]
            #The function identifyLastColumnWithContents checks the number of the last column with content in it, then two columns with timestamps are appended
            stockBalanceData["Downloaded_at_datetime"] = timestampToday()
            stockBalanceData["Downloaded_at_quarter"] = timestampQuarter()
            
            # Check and if not Dataframe datatype, convert to Dataframe
            if not isinstance(stockBalanceData, pd.DataFrame):
                stockBalanceData = pd.DataFrame(stockBalanceData)
            
            stockBalanceData.insert(0, "Symbol", stock)
            APIRequestDelay()
        
        except ValueError:
            stocksNotExisting.append(stock)
            print("Error" + str(counter) + " " + str(stock))
            counter = counter + 1
            pass
        
    return stockBalanceData, stocksNotExisting 


def update_balance_sheet_annual(updateNotExistingSymbols):
    stocksNotExisting = []
    counter = 0
        
    for stock in updateNotExistingSymbols:
        
        try:
            #the API command get_company_overview retrives stock data like Revenue, Cashflow, Ebit, etc. 
            stockBalanceData = FundamentalDataAPIPull.get_balance_sheet_annual(stock)[0]
            #The function identifyLastColumnWithContents checks the number of the last column with content in it, then two columns with timestamps are appended
            stockBalanceData["Downloaded_at_datetime"] = timestampToday()
            stockBalanceData["Downloaded_at_quarter"] = timestampQuarter()
            
            # Check and if not Dataframe datatype, convert to Dataframe
            if not isinstance(stockBalanceData, pd.DataFrame):
                stockBalanceData = pd.DataFrame(stockBalanceData)
            
            stockBalanceData.insert(0, "Symbol", stock)
            APIRequestDelay()
        
        except ValueError:
            stocksNotExisting.append(stock)
            print("Error" + str(counter) + " " + str(stock))
            counter = counter + 1
            pass
        
    return stockBalanceData, stocksNotExisting 


def update_income_statement_quarterly(updateNotExistingSymbols):
    stocksNotExisting = []
    counter = 0
        
    for stock in updateNotExistingSymbols:
        
        try:
            #the API command get_company_overview retrives stock data like Revenue, Cashflow, Ebit, etc. 
            stockBalanceData = FundamentalDataAPIPull.get_income_statement_quarterly(stock)[0]
            #The function identifyLastColumnWithContents checks the number of the last column with content in it, then two columns with timestamps are appended
            stockBalanceData["Downloaded_at_datetime"] = timestampToday()
            stockBalanceData["Downloaded_at_quarter"] = timestampQuarter()
            
            # Check and if not Dataframe datatype, convert to Dataframe
            if not isinstance(stockBalanceData, pd.DataFrame):
                stockBalanceData = pd.DataFrame(stockBalanceData)
            
            stockBalanceData.insert(0, "Symbol", stock)
            APIRequestDelay()
        
        except ValueError:
            stocksNotExisting.append(stock)
            print("Error" + str(counter) + " " + str(stock))
            counter = counter + 1
            pass
        
    return stockBalanceData, stocksNotExisting  



def update_income_statement_annual(updateNotExistingSymbols):
    stocksNotExisting = []
    counter = 0
        
    for stock in updateNotExistingSymbols:
        
        try:
            #the API command get_company_overview retrives stock data like Revenue, Cashflow, Ebit, etc. 
            stockBalanceData = FundamentalDataAPIPull.get_income_statement_annual(stock)[0]
            #The function identifyLastColumnWithContents checks the number of the last column with content in it, then two columns with timestamps are appended
            stockBalanceData["Downloaded_at_datetime"] = timestampToday()
            stockBalanceData["Downloaded_at_quarter"] = timestampQuarter()
            
            # Check and if not Dataframe datatype, convert to Dataframe
            if not isinstance(stockBalanceData, pd.DataFrame):
                stockBalanceData = pd.DataFrame(stockBalanceData)
            
            stockBalanceData.insert(0, "Symbol", stock)
            APIRequestDelay()
        
        except ValueError:
            stocksNotExisting.append(stock)
            print("Error" + str(counter) + " " + str(stock))
            counter = counter + 1
            pass
        
    return stockBalanceData, stocksNotExisting  


def getTimeSeriesData(updateNotExistingSymbols):
    stocksNotExisting = []
    counter = 0

    for stock in updateNotExistingSymbols:
        try:
            #the API command get_company_overview retrives stock data like Revenue, Cashflow, Ebit, etc. 
            dailySeriesPerStock = TimeSeriesPull.get_daily(stock, outputsize="full")[0]
            #The function identifyLastColumnWithContents checks the number of the last column with content in it, then two columns with timestamps are appended
            dailySeriesPerStock["Downloaded_at_datetime"] = timestampToday()
            dailySeriesPerStock["Downloaded_at_quarter"] = timestampQuarter()
            
            # Check and if not Dataframe datatype, convert to Dataframe
            if not isinstance(dailySeriesPerStock, pd.DataFrame):
                dailySeriesPerStock = pd.DataFrame(dailySeriesPerStock)
            
            # Convert index to a column to preserve the datetime of the actual stock value
            dailySeriesPerStock.reset_index(inplace=True)

            # Rename the column to Date
            dailySeriesPerStock.rename(columns={'index':'Date'}, inplace=True)
            dailySeriesPerStock["date"] = dailySeriesPerStock["date"].dt.date

            dailySeriesPerStock.insert(0, "Symbol", stock)
            APIRequestDelay()
        
        except ValueError:
            stocksNotExisting.append(stock)
            print("Error" + str(counter) + " " + str(stock))
            counter = counter + 1
            pass
        

    return dailySeriesPerStock, stocksNotExisting  


"""
def refresh(excel_data, refresh_list):
    today = pd.to_datetime(dt.datetime.today())
    today_quarter = int(today.quarter)
    for elem in refresh_list:
        try:
            stock1 = fd.get_company_overview(elem)[0]
            stock1.insert(59, "Downloaded_at", pd.to_datetime(dt.datetime.today()))
            stock1.insert(60, "Downloaded_quarter", today_quarter)
            excel_data.loc[excel_data["Symbol"] == elem] = stock_transform(stock1)
            time.sleep(time_delay)
        except ValueError:
            excel_data_refreshed = excel_data
            return excel_data_ref
    excel_data_refreshed = excel_data
    return excel_data_ref

def load_stock_price_yf(stock_list_new):
    stock_price_df = yf.download(stock_list_new, period="1d")
    stock_price_df_t = stock_price_df["Close"].T
    stock_price_df_t.insert(1, "Date", pd.to_datetime(dt.datetime.today()))
    stock_price_df_t1 = stock_price_df_t.reset_index()
    stock_price_df_t1.columns = ["Symbol","Price","Date"]
    
    return stock_price_df_t1
"""
           
def writeToExcel(mainExcel, worksheet):
    
    excelPath = "output.xlsx"

    book = load_workbook(excelPath)

    balanceSheet = book[worksheet]

    for row in dataframe_to_rows(mainExcel, index=False, header=False):
        balanceSheet.append(row)

    # Speichere die Änderungen
    book.save(excelPath)

def readFromDataBase(database):
    # Connection to Database
    conn = sql.connect("mainDatabase.db")

    try:
        mainSheetDataframe = pd.read_sql(f"SELECT Symbol, Downloaded_at_quarter FROM {database}", conn)
        
        excelSymbolsExisting = list(set(mainSheetDataframe["Symbol"]))
        excelQuartersExisting = list(mainSheetDataframe["Downloaded_at_quarter"])
        
        return excelSymbolsExisting, excelQuartersExisting

    except:

        excelSymbolsExisting = ["A"]
        excelQuartersExisting = timestampQuarter()

        return excelSymbolsExisting, excelQuartersExisting

    finally:
        conn.close()


######## TO-DO: Update based on Database column symbols, NOT the excel sheet. Write new funcs for all features overview/balance/income
def writeToDataBase(mainExcel, database):

    conn = sql.connect("mainDatabase.db")
    mainExcel.to_sql(database, conn, if_exists="append", index=False)
    

    conn.close()


def sortDatabaseBySymbolName(table):
    
    conn = sql.connect("mainDatabase.db")

    cursor = conn.cursor()

    # Create temp table differently sorted
    cursor.execute(f"CREATE TEMPORARY TABLE temp_{table} AS SELECT * FROM {table} ORDER BY Symbol ASC")
    
    # Delete old table
    cursor.execute(f"DROP TABLE {table}")
    
    cursor.execute(f"ALTER TABLE temp_{table} RENAME TO {table}")


    conn.commit()

    conn.close()



"""

def write_to_excel_daily_stock_price(result):
    book = load_workbook("output.xlsx")
    book_stock_price = book["Daily_Stock_Price"]

    writer = pd.ExcelWriter("output.xlsx", engine='openpyxl') 

    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
  
    result.to_excel(writer, index=False, header=True, sheet_name="Daily_Stock_Price")
    writer.save() 

"""

# NOT USED: Implementation as is, leads to deletion of Symbol entries in Stock_symbol_list, because the API request does get canceled from provide due to exhaustion of free API requests per day.
# TO-DO: When this particular error comes back from server, it is needed to ignore this delete func.
def deleteNoneUpdatableSymbols(symbol_list, stocksNotExisting):
    filename = "Stock_symbols_list.xlsx"
    sheetname = "Overview"
     
    # Compare symbol list against stocks not existing anymore at the stock exchange list
    filtered_symbols = [symbol for symbol in symbol_list if symbol not in stocksNotExisting]

    excelDataFrame = pd.DataFrame(filtered_symbols, columns=[0])

    excelDataFrame.to_excel(filename, sheet_name=sheetname, index=False, header=False)







"""

def prep_df_for_calc(df_to_prep):
    df_to_prep = df_to_prep.dropna()
    for elem in range(len(df_to_prep)):
        if df_to_prep[elem] == "None":
            df_to_prep[elem] = 1
        if type(amount_stocks[elem]) == str:
            df_to_prep[elem] = float(df_to_prep[elem])
        if type(amount_stocks[elem]) == float:
            df_to_prep[elem] = int(df_to_prep[elem])

    return df_to_prep

def calc_price_per_share():
    book = load_workbook("output.xlsx")
    book_overview = book["Overview"]
    book_balance_quarterly = book["Balance_Quarterly"]
    book_income_quarterly = book["Income_Quarterly"]
    book_stock_price = book["Daily_Stock_Price"]

    data_balance = book_balance.values
    data_income = book_income.values
    data_stock_price = book_stock_price.values


    header_bal = next(data_balance)
    header_inc = next(data_income)
    header_pri = next(data_stock_price)

    table_bal = pd.DataFrame(data_balance,columns=header_bal, dtype="int64")
    table_bal["commonStockSharesOutstanding"][table_bal["commonStockSharesOutstanding"]=="None"] = 1
    table_inc = pd.DataFrame(data_income,columns=header_inc, dtype="int64")
    table_inc["netIncome"][table_inc["netIncome"]=="None"] = 1
    table_pri = pd.DataFrame(data_stock_price,columns=header_pri, dtype="int64")

    table_pri = table_pri[["Symbol","Price"]]
    table_pri = table_pri.dropna()
    amount_stocks = table_bal["commonStockSharesOutstanding"]
    net_income = table_inc["netIncome"]

    amount_stocks = prep_df_for_calc(amount_stocks)
    net_income = prep_df_for_calc(net_income)

    multipliers = pd.DataFrame(columns=["Symbol", "fiscalDateEnding","EarningsPerShare","SharePrice"])
    multipliers["EarningsPerShare"] = net_income/amount_stocks*4
    multipliers["Symbol"] = table_inc["Symbol"]
    multipliers["fiscalDateEnding"] = table_inc["fiscalDateEnding"]

    for pos,elem in enumerate(multipliers["Symbol"]):
        for pri_elem in table_pri["Symbol"]:
            if elem == pri_elem:
                multipliers["SharePrice"][pos] = table_pri["Price"][table_pri["Symbol"]==elem].values[0]
    
    return multipliers
"""


   