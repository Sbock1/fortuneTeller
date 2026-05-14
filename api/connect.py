##################################################################
# Alpha Vantage API Client Configuration and Initialization
# 
# Author: Sebastian Bock (Sbock)
# Created Date: 29.09.2024
# Last Modified: May 14, 2026
#
# Description:
# This module manages all API requests to the Alpha Vantage service.
# It initializes and exposes two main API clients:
# - FundamentalDataAPIPull: For fetching company fundamentals (balance sheets, 
#   income statements, cash flows)
# - TimeSeriesAPIPull: For fetching historical stock price data
#
# Configuration is loaded from environment variables to prevent exposing 
# sensitive API credentials in source code.
#
# IMPORTANT - API Documentation:
# Full Alpha Vantage API documentation: https://www.alphavantage.co/documentation/
#
# Rate Limits:
# - Free tier: 25 API requests per day
# - Premium tiers available at https://www.alphavantage.co/premium/
##################################################################

from alpha_vantage.fundamentaldata import FundamentalData
from alpha_vantage.timeseries import TimeSeries
import os
from dotenv import load_dotenv


# Load all environment variables from the .env file into the application.
# This allows sensitive configuration (like API keys) to be stored outside 
# of version control.
load_dotenv()


# Retrieve the Alpha Vantage API key from environment variables.
# The API key is required for all API requests and must be stored in the 
# .env file under the variable name 'ALPHAVANTAGE_API_KEY'.
# If the key is not found, a ValueError is raised immediately to prevent 
# any silent failures later during API calls.
API_key: str = os.getenv("ALPHAVANTAGE_API_KEY")
if not API_key:
    raise ValueError(
        "ALPHAVANTAGE_API_KEY not found in environment variables. "
        "Please set it in your .env file or as an environment variable."
    )


# Delay (in seconds) between consecutive API requests.
# Alpha Vantage's free tier allows 25 requests per day, so a delay helps 
# prevent hitting rate limits. The value can be configured via the 
# API_REQUEST_DELAY environment variable (default is 1 second).
# Note: Production code should implement proper rate limiting with 
# exponential backoff for failed requests.
APIRequestDelay: int = int(os.getenv("API_REQUEST_DELAY", "1"))


# Initialize the FundamentalData API client.
# This client provides access to company fundamental data including:
# - Company overview (structure, market cap, description, etc.)
# - Balance sheets (quarterly and annual)
# - Income statements (quarterly and annual)
# - Cash flow statements (quarterly and annual)
# Output format is set to 'pandas' to return data as pandas DataFrames.
FundamentalDataAPIPull = FundamentalData(API_key, output_format='pandas')


# Initialize the TimeSeries API client.
# This client provides access to historical stock price data including:
# - Daily stock prices (open, high, low, close, volume)
# - Intraday data (if available for the symbol)
# Output format is set to 'pandas' to return data as pandas DataFrames.
TimeSeriesAPIPull = TimeSeries(key=API_key, output_format="pandas")