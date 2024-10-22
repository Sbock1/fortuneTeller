import sqlite3 as sql
import pandas as pd

import statsmodels.api as sm
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
import matplotlib.pyplot as plt
from statsmodels.tsa.stattools import adfuller
from pmdarima import auto_arima
from statsmodels.tsa.arima.model import ARIMA
from sklearn.metrics import mean_squared_error
#from pmdarima.arima.utils import ndiffs

###########################################
### This module consists of functions, which allow different analytics algos and machine learning techniques to be prepared and deployed on the existing database
### Author: Sebastian Bock
### Created Date: 29.09.2024
###########################################

### Load database data for further processing 
def loadStockDataFromDatabase(stock: str):
    # Connect to sql database, it is the mainDatabase.db
    conn = sql.connect("mainDatabase.db")

    # Load data from SQL database into pandas dataframe datatype
    query = f"SELECT Symbol, fiscalDateEnding, Earnings_Per_Share FROM Analysis_Quarterly WHERE Symbol is '{stock}'"
    dataframe = pd.read_sql(query, conn)
    #dataframe["fiscalDateEnding"] = pd.to_datetime(dataframe["fiscalDateEnding"])
    #dataframe.sort_values("fiscalDateEnding")

    return dataframe

### ARIMA Model (time-based methode) TO-DO: Complexity in order to properly implement the model much higher than anticipated,
#  MUST BE REFINED!!!
def analysisARIMA(dataframe: pd.DataFrame):
    #Using fiscalDateEnding as Index
    dataframe['fiscalDateEnding'] = pd.to_datetime(dataframe['fiscalDateEnding'])
    dataframe.set_index("fiscalDateEnding", inplace=True)
    print(dataframe)

    # Dickey-Fuller-Test 
    result = adfuller(dataframe["Earnings_Per_Share"].dropna())
    print("ADF Statistic: ", result[0])
    print("p-value: ", result[1])
 
    # Generate actual ARIMA model automatically
    model_fit = auto_arima(dataframe["Earnings_Per_Share"], start_p=3, start_q=2, seasonal=False, stepwise=False, trace=True)
    print(model_fit.summary())

    ############################ Residual analysis ###############################
    residuals = dataframe["Earnings_Per_Share"] - model_fit.predict_in_sample()
    residuals = pd.DataFrame(residuals, columns=['Residuals'])

    # Residuen plotten
    residuals.plot(title="Residuals")
    plt.show()

    # Histogramm der Residuen mit Normalverteilungskurve
    residuals.plot(kind='kde', title="Density of Residuals")
    plt.show()

    # QQ-Plot der Residuen
    sm.qqplot(residuals.squeeze(), line='s')
    plt.title("QQ-Plot of Residuals")
    plt.show()
    
    ############################# Forecasting #####################################

    # Anzahl der Perioden, für die du Vorhersagen treffen möchtest (z.B. 10 Schritte in die Zukunft)
    forecast_steps = 20
    forecast = model_fit.predict(n_periods=forecast_steps)

    print(forecast)

    # Mean Squared Error berechnen
    actual_values = dataframe["Earnings_Per_Share"][-forecast_steps:]
    mse = mean_squared_error(actual_values, forecast[:len(actual_values)])
    print(f'Mean Squared Error: {mse}')

    # Visualisierung der Vorhersagen
    plt.plot(dataframe.index, dataframe["Earnings_Per_Share"], label='Original')

    # Neuer Index für die Prognosen, der nach dem letzten Datum beginnt
    forecast_index = pd.date_range(start=dataframe.index[-1], periods=forecast_steps + 1, freq='ME')[1:]
    
    plt.plot(forecast_index, forecast, label='Forecast', color='red')
    plt.legend()
    plt.show()


dataframe = loadStockDataFromDatabase("A")
analysisARIMA(dataframe)
