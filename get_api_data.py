import APIandDatabaseLib as gil
import time
import tkinter as tk
from tkinter import messagebox
import argparse as ap

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


def loadAndStoreAPIDataPull(amount, destination: str, source: str):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    if destination == "Excel":
        excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", source)
    elif destination == "DB":
        excelSymbolsExisting, excelQuartersExisting = gil.readFromDataBase(source)
    print(f"Step 1: Loading initial data from ## {source} ## -> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Following symbols exist already {excelSymbolsExisting}")

    if isinstance(amount, int):
        symbolsToUpdate = updateNotExistingSymbols[0:amount]
    elif isinstance(amount, list):
        symbolsToUpdate = [symbol for symbol in updateNotExistingSymbols if symbol in amount] 
    elif isinstance(amount, str):
        symbolsToUpdate = [amount]
    else:
        raise ValueError(f"Invalid input for 'amount' {amount}. Expected int, list, or str.")

    if source == "Overview":
        stockAPIData, stocksNotExisting = gil.updateCompanyOverview(symbolsToUpdate)
    elif source == "Balance_Quarterly":
        stockAPIData, stocksNotExisting = gil.update_balance_sheet_quarterly(symbolsToUpdate)
    elif source == "Balance_Yearly":
        stockAPIData, stocksNotExisting = gil.update_balance_sheet_annual(symbolsToUpdate)
    elif source == "Income_Quarterly":
        stockAPIData, stocksNotExisting = gil.update_income_statement_quarterly(symbolsToUpdate)
    elif source == "Income_Yearly":
        stockAPIData, stocksNotExisting = gil.update_income_statement_annual(symbolsToUpdate)
    elif source == "Cashflow_Quarterly":
        stockAPIData, stocksNotExisting = gil.updateCashflowStatementQuarterly(symbolsToUpdate)
    elif source == "Cashflow_Yearly":
        stockAPIData, stocksNotExisting = gil.updateCashflowStatementAnually(symbolsToUpdate)      
    elif source == "Stocks_Daily":
        stockAPIData, stocksNotExisting = gil.getTimeSeriesData(symbolsToUpdate)

    print(f"Step 3: Creation DataFrame with update of {destination} with {amount} entries which are: {symbolsToUpdate}")
    
    # Where to write the dataframe to: Either Excel file named "output.xlsx" or into a SQL Database
    if destination == "Excel":
        gil.writeToExcel(stockAPIData, source)
    elif destination == "DB":
        gil.writeToDataBase(stockAPIData, source)

    #gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)

allDataGroup = ["Overview", "Balance_Quarterly", "Balance_Yearly", "Income_Quarterly", "Income_Yearly", "Cashflow_Quarterly", "Cashflow_Yearly", "Stock_Daily"]

for elem in allDataGroup:
    loadAndStoreAPIDataPull("MSFT", "DB", elem)

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
