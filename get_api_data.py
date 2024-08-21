import get_information_lib as gil
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
    gil.deleletExcelPredefinedSheet()


def initializeSQLDatabase(): # TO-DO
    pass


def sortAllTables():
    
    for table in ["Overview", "Balance_Quarterly", "Balance_Yearly", "Income_Quarterly", "Income_Yearly", "Stocks_Daily"]:
        gil.sortDatabaseBySymbolName(table)

sortAllTables()


def createWindow():
    # Main window
    root = tk.Tk()
    # Window title
    root.title("Load data window")
    # Fenstergröße setzen
    root.geometry("400x300")
    # Ein Label-Widget hinzufügen
    label = tk.Label(root, text="You can decide how many stock overview data you want to download. Currently each row will take 18 seconds.")
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


def loadOverviewExcel(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Overview")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    stockOverviewData, stocksNotExisting = gil.updateCompanyOverview(updateNotExistingSymbols[0:amount])
    print(f"Step 2: Creation DataFrame with update of mainExcel with {amount} entries")
    #gil.writeToDataBase(stockOverviewData, "Overview")
    gil.writeToExcel(stockOverviewData, "Overview")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadOverviewDatabase(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Overview")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    stockOverviewData, stocksNotExisting = gil.updateCompanyOverview(updateNotExistingSymbols[0:amount])
    print(f"Step 2: Creation DataFrame with update of mainExcel with {amount} entries")
    gil.writeToDataBase(stockOverviewData, "Overview")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_balance_quartely(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Balance_Quarterly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_balance_sheet_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Balance_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadBalanceQuarterlyDatabase(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Balance_Quarterly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    stockOverviewData, stocksNotExisting = gil.update_balance_sheet_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 2: Creation DataFrame with update of mainExcel with {amount} entries")
    gil.writeToDataBase(stockOverviewData, "Balance_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)
    

def load_balance_annual(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Balance_Yearly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_balance_sheet_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Balance_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)
    

def loadBalanceAnnuallyDatabase(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Balance_Yearly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_balance_sheet_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Balance_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_income_quartely(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Income_Quarterly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Income_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadIncomeQuarterlyDatabase(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Income_Quarterly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Income_Quarterly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_income_annual(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Income_Yearly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Income_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadIncomeAnnuallyDatabase(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Income_Yearly")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_income_statement_annual(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Income_Yearly")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_daily_stock(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Stocks_Daily")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.getTimeSeriesData(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcel(stockBalanceData, "Stocks_Daily")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def loadDailyStockDatabase(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase("Stocks_Daily")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.getTimeSeriesData(updateNotExistingSymbols[0:amount])
    print(f"Step 3: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToDataBase(stockBalanceData, "Stocks_Daily")
    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)

   
#loadOverviewExcel(2)
#loadOverviewDatabase(1)   
#load_balance_quartely(1)
#loadBalanceQuarterlyDatabase(2)
#load_balance_annual(1)
#loadBalanceAnnuallyDatabase(2)
#load_income_quartely(1)
#loadIncomeQuarterlyDatabase(2)
#load_income_annual(1)
#loadIncomeAnnuallyDatabase(2)
#load_daily_stock(1)
loadDailyStockDatabase(1)


'''
def load_daily_stock_prices():
    stock_list_new = gil.load_stock_symbol_list()
    stock_list_con = " ".join(str(elem) for elem in stock_list_new)

    daily_price_stock_list = gil.load_stock_price_yf(stock_list_con)
    gil.write_to_excel_daily_stock_price(daily_price_stock_list)



'''

#load_daily_stock_prices()
