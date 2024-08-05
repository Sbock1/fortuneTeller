import get_information_lib as gil
import time
import tkinter as tk
from tkinter import messagebox

def initializeExcelSheet():
    gil.insertFirstRowColumnNames()
    gil.insertFirstRowColumnNamesBalance()
    gil.deleletExcelPredefinedSheet()

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
            load_overview(amount)
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


def load_overview(amount):
    
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Overview")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    stockOverviewData, stocksNotExisting = gil.updateCompanyOverview(mainSheetDataframe, updateNotExistingSymbols[0:amount])
    print(f"Step 2: Creation DataFrame with update of mainExcel with {amount} entries")
    gil.writeToExcelToUpdateOverview(stockOverviewData, "Overview")
    gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)


def load_balance_quartely(amount):
    symbolList = gil.loadOneColumnRowDataAsList("Stock_symbols_list.xlsx", "Overview", "A")
    mainSheetDataframe, excelSymbolsExisting, excelQuartersExisting = gil.getExcelSheetInformation("output.xlsx", "Balance")
    print("Step 1: Loading initial data from Excel-> Done")
    existingSymbols, symbolsNeedRefresh, updateNotExistingSymbols = gil.checkSymbolCurrentQuarterExisting(symbolList, excelSymbolsExisting, excelQuartersExisting)
    print(f"Step 2: Follwing symbols exist already {excelSymbolsExisting}")
    stockBalanceData, stocksNotExisting = gil.update_balance_sheet_quarterly(updateNotExistingSymbols[0:amount])
    print(f"Step 4: Creation DataFrame with update of mainExcel with {amount} entries which are {updateNotExistingSymbols[0:amount]}")
    gil.writeToExcelToUpdateOverview(stockBalanceData, "Balance")
    gil.deleteNoneUpdatableSymbols(symbolList, stocksNotExisting)

#load_overview(1)
#load_balance_quartely(1)

'''
def load_balance_annual(amount):
    stock_list_new = gil.load_stock_symbol_list()
    excel_data_bal, excel_symbol_output, excel_quarter_output = gil.get_output_data_to_pandas("output.xlsx", "Balance")
    check_list, refresh_list, update_list = gil.search_symbol(stock_list_new, excel_symbol_output, excel_quarter_output)
    excel_data_up_bal, not_updatable_bal = gil.update_balance_sheet_annual(excel_data_bal, update_list[0:amount])

    gil.write_to_excel_update_balance(excel_data_up_bal)
    gil.delete_not_upgradable_symbols(stock_list_new, not_updatable_bal)

def load_income_quartely(amount):
    stock_list_new = gil.load_stock_symbol_list()
    excel_data_inc, excel_symbol_output, excel_quarter_output = gil.get_output_data_to_pandas("output.xlsx", "Income")
    check_list, refresh_list, update_list = gil.search_symbol(stock_list_new, excel_symbol_output, excel_quarter_output)
    excel_data_up_inc, not_updatable_inc = gil.update_income_statement_quarterly(excel_data_inc, update_list[0:amount])

    gil.write_to_excel_update_income(excel_data_up_inc)
    gil.delete_not_upgradable_symbols(stock_list_new, not_updatable_inc)

def load_income_annual(amount):
    stock_list_new = gil.load_stock_symbol_list()
    excel_data_inc, excel_symbol_output, excel_quarter_output = gil.get_output_data_to_pandas("output.xlsx", "Income")
    check_list, refresh_list, update_list = gil.search_symbol(stock_list_new, excel_symbol_output, excel_quarter_output)
    excel_data_up_inc, not_updatable_inc = gil.update_income_statement_annual(excel_data_inc, update_list[0:amount])

    gil.write_to_excel_update_income(excel_data_up_inc)
    gil.delete_not_upgradable_symbols(stock_list_new, not_updatable_inc)



def load_daily_stock_prices():
    stock_list_new = gil.load_stock_symbol_list()
    stock_list_con = " ".join(str(elem) for elem in stock_list_new)

    daily_price_stock_list = gil.load_stock_price_yf(stock_list_con)
    gil.write_to_excel_daily_stock_price(daily_price_stock_list)

#Write function that create the "output.xlsx" file if not already existing


'''

#load_overview(1)
#load_income_annual(60)
#load_balance_annual(230)

#load_daily_stock_prices()
