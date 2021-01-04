# nse-data-analysis
OI data analysis of NSE Option chain


Before running these scripts update google_sheet ID and Google API token, also enable API for google sheets.

---------

Option chain contains very important information but the trend formation about the underlying. This script is scraping data from NSE API and calculating sum of open interest change in near ATM strikes for calls and Puts.

If the Call OI change is more than Put OI change then we assume that Big players are doing Call writing and trend formation is bearish. (Note: the difference should be substantial), and vice-a-versa sum of near ATM strike OI change of Puts is greater than Call than it is assumed that Puts are being sold more than Calls and view is bullish.
