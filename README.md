# fortuneTeller

**Goal of the project: **
Create a stock forecast model that takes into account:
1. Module: Fundamental Stock Data
2. Module: Sentiment Data
3. Module: Economic and environmental data

**Step One: (ongoing)**
In a first project step, the fundamental data needs to be gathered and stored. This is done by using the AlphaVantage API.
The API allows downloads of:
- Balance Sheets
- Income statements
- Cashflow statements
- Daily stock data

The data will be pulled based on the API wrappers predefined by AV (Python->pandas->openpyxl) and pushed into an excel workbook for storage.
(a next step will be an mySQL database implementation to allow better scaling)

**Step Two:**
The data from the created database will be analyzed by:
- Step a: Fundamental stock data analysis methods (Debt Structure / Liquidity Ratios / Profitability Ratios / Efficiency Ratios / Comparative Analysis [segment/industry] )
- Step b: Stock time series data analysis methods (Time Series Analysis [ARIMA / Seasonal ARIMA] / Regression Models / Machine Learning [Random Forest / Gradient Boosting Machines]  )
- Step c: Sentiment analysis

