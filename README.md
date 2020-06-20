# Python NSE-Option-Chain-Analyzer

### [Downloads](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases)

This program web scrapes and then generates useful analysis of the Option Chain for the specified security from the NSE website.
It also continuously refreshes the Option Chain and visually displays the trend in various indicators and useful for Technical Analysis.

## Usage:

-Open Option Chain (Equity Derivatives) page on https://www1.nseindia.com/

-Select your Index or Underlying Stock and Expiry Date on the website

-Run the program and enter the final URL from the browser and your preferred Strike Price

## Note:

-Required modules:

```
bs4
requests
pandas
time
```

-Install missing modules using pip

-In case of network or connection errors the program doesn't crash and will keep retrying infinitely

-If a ZeroDivisionError occurs or some data doesn't exist the value of the variable will be defaulted to 0

## Features:

-The program infinitely refreshes every 30 seconds

-New data is printed only if the server time has changed (To prevent printing duplicate data)

-PEP 8 format

-Data printed:

Data | How it's calculated
--- | ---
Server Time | *Web Scraped*
Underlying Index Points | *Web Scraped*
Max Call Open Interest and corresponding Strike Price | Highest Call Open Interest (in lacs)
Max Put Open Interest and corresponding Strike Price | Highest Put Open Interest (in lacs)
Put Call Ratio(PCR) | Total Put Open Interest divided by Total Call Open Interest
Call Sum | Sum of the Change in Call Open Interests of the given Strike Price and the next two Strike Prices (in lacs)
Put Sum | Sum of the Change in Put Open Interests of the given Strike Price and the next two Strike Prices (in lacs)
Difference | Difference between the Call Sum and Put Sum
Call Boundary | Change in Call Open Interest for 2 Strike Prices above the given Strike Price. This is used to determine if Call writers are exiting their positions.
Put Boundary | Change in Put Open Interest for the given Strike Price. This is used to determine if Put writers are exiting their positions.
Call In The Money(ITM) | This indicates that bullish trend could continue and Value could cross 4 Strike Prices above given Strike Price. It's the ratio of Put writing and Call writing at the 4th Strike Price above the given Strike price.
Put In The Money(ITM) | This indicates that bearish trend could continue and Value could cross 2 Strike Prices below given Strike Price. It's the ratio of Call writing and Put writing at the 2nd Strike Price below the given Strike price.
