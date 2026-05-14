################################################################## 
# Author: Sebastian Bock (Sbock)
# Created Date: 29.09.2024
# Description: 
# This file manages the request of data from the alpha vantage API

# IMPORTANT:
# The AlphaVantage API documentation can be found here:
# https://www.alphavantage.co/documentation/
##################################################################

from alpha_vantage.fundamentaldata import FundamentalData
from alpha_vantage.timeseries import TimeSeries

# Trial Account API key to validate access
API_key: str = "HNNMDBOG55BREC5P"
# The free API access only allows 25 requests per day without any delay, 
# but for testing purposes this is sufficient.
Time_delay: int = 1 
# Initialize the API clients for FundamentalData and TimeSeries
FundamentalDataAPIPull = FundamentalData(API_key, output_format='pandas')
TimeSeriesAPIPull = TimeSeries(key=API_key, output_format="pandas")