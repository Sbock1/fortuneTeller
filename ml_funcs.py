import sqlite3 as sql
import pandas as pd

from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
import matplotlib.pyplot as plt
from statsmodels.tsa.stattools import adfuller
#from pmdarima.arima.utils import ndiffs

###########################################
### This module consists of functions, which allow machine learning techniques to be prepared and deployed on the existing database
### Author: Sebastian Bock
### Date: 29.09.2024
###########################################

### Load database data for further processing 
def loadStockDataFromDatabase(stock: chr):
    # Connect to sql database, it is the mainDatabase.db
    conn = sql.connect("mainDatabase.db")

    # Load data from SQL database into pandas dataframe datatype
    query = f"SELECT Symbol, fiscalDateEnding, Earnings_Per_Share FROM Analysis_Quarterly WHERE Symbol is '{stock}'"
    dataframe = pd.read_sql(query, conn)
    #dataframe["fiscalDateEnding"] = pd.to_datetime(dataframe["fiscalDateEnding"])
    #dataframe.sort_values("fiscalDateEnding")

    return dataframe

### ARIMA Model (time-based methode)
def analysisARIMA(dataframe: pd.DataFrame):
    #Using fiscalDateEnding as Index
    dataframe.set_index("fiscalDateEnding", inplace=True)
    print(dataframe)

    # Dickey-Fuller-Test 
    result = adfuller(dataframe["Earnings_Per_Share"].dropna())

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 4))
    ax1.plot(dataframe["Earnings_Per_Share"])
    plot_acf(dataframe["Earnings_Per_Share"], ax=ax2)

    print("ADF Statistic: ", result[0])
    print("p-value: ", result[1])

    dataframe['diff_EPS'] = dataframe['Earnings_Per_Share'].diff().dropna()
    dataframe['diff_EPS'] = dataframe["diff_EPS"].diff().dropna()

    '''fig, (ax3, ax4) = plt.subplots(1, 2, figsize=(16, 4))

    ax3.plot(dataframe["diff_EPS"])

    # ACF und PACF Diagramme erstellen
    plot_acf(dataframe['diff_EPS'].dropna())
    plot_pacf(dataframe['diff_EPS'].dropna())
    plt.show()
    '''
    
    #ndiffs(dataframe["Earnings_Per_Share"], test="adf")


dataframe = loadStockDataFromDatabase("A")
analysisARIMA(dataframe)
