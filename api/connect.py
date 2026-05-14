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
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Get API key from environment variables (raises error if not set)
API_key: str = os.getenv("ALPHAVANTAGE_API_KEY")
if not API_key:
    raise ValueError(
        "ALPHAVANTAGE_API_KEY not found in environment variables. "
        "Please set it in your .env file or as an environment variable."
    )

# The free API access only allows 25 requests per day without any delay, 
# but for testing purposes this is sufficient.
APIRequestDelay: int = int(os.getenv("API_REQUEST_DELAY", "1"))

# Initialize the API clients for FundamentalData and TimeSeries
FundamentalDataAPIPull = FundamentalData(API_key, output_format='pandas')
TimeSeriesAPIPull = TimeSeries(key=API_key, output_format="pandas")