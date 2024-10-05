import APIandDatabaseLib as gil
import time
import tkinter as tk
from tkinter import messagebox

#ONLY USE if not already initialized "output.xlsx" is present. Otherwise header gets downloaded again -> TO-DO: Change func to check if file and header per worksheet are exisitng
def initializeExcelSheet():
    gil.insertFirstRowColumnNames()
    gil.insertFirstRowColumnNamesBalanceQuarterly()
    gil.insertFirstRowColumnNamesBalanceYearly()
    gil.insertFirstRowColumnNamesIncomeQuarterly()
    gil.insertFirstRowColumnNamesIncomeYearly()
    gil.insertFirstRowColumnNamesStockDaily()
    gil.insertFirstRowColumnNamesCashflowYearly()
    gil.insertFirstRowColumnNamesCashflowQuarterly()
    gil.deleletExcelPredefinedSheet()


def sortAllTables():
    
    for table in ["Overview", "Balance_Quarterly", "Balance_Yearly", "Income_Quarterly", "Income_Yearly", "Stocks_Daily", "Cashflow_Yearly", "Cashflow_Quarterly"]:
        gil.sortDatabaseBySymbolName(table)


def createWindow():
    # Main window
    root = tk.Tk()
    # Window title
    root.title("Load data window")
    # Fenstergröße setzen
    root.geometry("400x300")
    # Ein Label-Widget hinzufügen
    label = tk.Label(root, text="You can decide how many stock overview data you want to download. Currently each row will take 1 seconds.")
    label.pack(pady=20)  # Pack-Layout-Manager verwenden
        
    # Ein Button-Widget hinzufügen
    def load_overview_button_clicked():
        try:
            amount = int(entry.get())
            loadOverviewExcel(amount)
            messagebox.showinfo("Info", f"Overview was expanded by {amount} entries.")
        except ValueError:
            messagebox.showerror("Error", "Please insert a valid int value.")

    entry = tk.Entry(root)
    entry.pack(pady=10)

    button = tk.Button(root, text="Download entries", command=load_overview_button_clicked)
    button.pack(pady=10)

    # Hauptschleife starten
    root.mainloop()

#createWindow()


def loadOverviewExcel(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Overview")
    print("Step 1: Loading initial data from Excel: ## Overview ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Following symbols exist already {excelSymbolsExisting}")
    stockOverviewData, stocksNotExisting = gil.updateCompanyOverview(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of myExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockOverviewData, "Overview")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadOverviewDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Overview")
    print("Step 1: Loading initial data from Database: ## Overview ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Following symbols exist already {excelSymbolsExisting}")
    stockOverviewData, stocksNotExisting = gil.updateCompanyOverview(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockOverviewData, "Overview")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_balance_quartely(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Balance_Quarterly")
    print("Step 1: Loading initial data from Excel: ## Balance_Quarterly ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Following symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_balance_sheet_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Balance_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadBalanceQuarterlyDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Balance_Quarterly")
    print("Step 1: Loading initial data from Database: ## Balance_Quarterly ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Following symbols exist already {excelSymbolsExisting}")
    stockOverviewData, stocksNotExisting = gil.update_balance_sheet_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockOverviewData, "Balance_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)
    

def load_balance_annual(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Balance_Yearly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_balance_sheet_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Balance_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)
    

def loadBalanceAnnuallyDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Balance_Yearly")
    print("Step 1: Loading initial data from Database: ## Balance_Yearly ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_balance_sheet_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Balance_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_income_quartely(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Income_Quarterly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Income_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadIncomeQuarterlyDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Income_Quarterly")
    print("Step 1: Loading initial data from Database: ## Income_Quarterly ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Income_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_income_annual(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Income_Yearly")
    print("Step 1: Loading initial data from Excel: ## Income_Annually ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Income_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadIncomeAnnuallyDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Income_Yearly")
    print("Step 1: Loading initial data from Database: ## Income_Annually ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Income_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_cashflow_annual(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Cashflow_Yearly")
    print("Step 1: Loading initial data from Excel: ## Cashflow_Annually ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.updateCashflowStatementAnually(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Cashflow_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadCashflowAnuallyDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Cashflow_Yearly")
    print("Step 1: Loading initial data from Database: ## Cashflow_Yearly ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.updateCashflowStatementAnually(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Cashflow_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_cashflow_quarterly(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Cashflow_Quarterly")
    print("Step 1: Loading initial data from Excel: ## Cashflow_Quarterly ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.updateCashflowStatementQuarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Cashflow_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadCashflowQuarterlyDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Cashflow_Quarterly")
    print("Step 1: Loading initial data from Database: ## Cashflow_Quarterly ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.updateCashflowStatementQuarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Cashflow_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)

def load_daily_stock(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Stocks_Daily")
    print("Step 1: Loading initial data from Excel: ## Stocks_Daily ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.getTimeSeriesData(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Stocks_Daily")
    #gil.writeToDataBase(stockBalanceData, "Stocks_Daily")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadDailyStockDatabase(amount: int):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Stocks_Daily")
    print("Step 1: Loading initial data from Database: ## Stocks_Daily ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.getTimeSeriesData(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of Database with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Stocks_Daily")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)





#gil.insertFirstRowColumnNamesStockDaily()
#loadOverviewExcel(2)
#load_balance_quartely(10)

#load_balance_annual(1)
#loadBalanceAnnuallyDatabase(11)
#load_income_quartely(1)
#load_income_annual(1)
#loadIncomeAnnuallyDatabase(11)
#load_cashflow_annual(1)
#loadCashflowAnuallyDatabase(4)
#load_cashflow_quarterly(1)
#load_daily_stock(1)

#gil.createAnalysisYearlyTable()

loadOverviewDatabase(3) 
loadBalanceQuarterlyDatabase(3)
loadIncomeQuarterlyDatabase(3)
loadCashflowQuarterlyDatabase(3)
loadDailyStockDatabase(3)


gil.createAnalysisQuarterlyTable()
gil.cleaningDatabaseNulltoZero()
gil.addingPrimaryKeyColumn()
sortAllTables()

'''
def load_daily_stock_prices():
    stock_list_new = gil.load_stock_symbol_list()
    stock_list_con = " ".join(str(elem) for elem in stock_list_new)

    daily_price_stock_list = gil.load_stock_price_yf(stock_list_con)
    gil.write_to_excel_daily_stock_price(daily_price_stock_list)



'''

#load_daily_stock_prices()
