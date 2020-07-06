# Python NSE-Option-Chain-Analyzer

### [Downloads](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

This program web scrapes and then generates useful analysis of the Option Chain for the specified security from the NSE website.
It also continuously refreshes the Option Chain and visually displays the trend in various indicators and useful for Technical Analysis.

## Usage:

-Open Option Chain (Equity Derivatives) page on https://www1.nseindia.com/

-Select your Index or Underlying Stock and Expiry Date on the website

-Run the program and enter the current URL from the browser and your preferred Strike Price

-You can select all table data using Ctrl+A or select individual cells, rows and columns 

-Then you can copy it using Ctrl+C or right click menu

-You can then paste it in any spreadsheet application (Tested with MS Excel and Google Sheets)

-You can also export all data to a .csv file from the menu

## Note:

-Required modules:

```
tkinter
tksheet
bs4
requests
pandas
datetime
webbrowser
csv
```

-Install missing modules using pip

-In case of network or connection errors the program doesn't crash and will keep retrying infinitely

-If a ZeroDivisionError occurs or some data doesn't exist the value of the variable will be defaulted to 0

## Features:

-The program infinitely refreshes every 30 seconds

-New data is displayed only if the server time has increased (To prevent printing duplicate data)

-Red and Green colour indication for data based on trends

-Program title format: NSE-Option-Chain-Analyzer - {underlying_stock} - {expiry_date} - {strike_price}

-Stop and Start functionality

-Export to .csv file option

-About window with version and links for developer GitHub profile, README, license, releases and sources

-PEP 8 format

-Table Data displayed:

Data | How it's calculated
--- | ---
Server Time | *Web Scraped*
Underlying Index Points | *Web Scraped*
Call Sum | Sum of the Change in Call Open Interests of the given Strike Price and the next two Strike Prices (in lacs)
Put Sum | Sum of the Change in Put Open Interests of the given Strike Price and the next two Strike Prices (in lacs)
Difference | Difference between the Call Sum and Put Sum
Call Boundary | Change in Call Open Interest for 2 Strike Prices above the given Strike Price. This is used to determine if Call writers are exiting their positions.
Put Boundary | Change in Put Open Interest for the given Strike Price. This is used to determine if Put writers are exiting their positions.
Call In The Money(ITM) | This indicates that bullish trend could continue and Value could cross 4 Strike Prices above given Strike Price. It's the ratio of Put writing and Call writing at the 4th Strike Price above the given Strike price.
Put In The Money(ITM) | This indicates that bearish trend could continue and Value could cross 2 Strike Prices below given Strike Price. It's the ratio of Call writing and Put writing at the 2nd Strike Price below the given Strike price.

-Label Data displayed:

Data | How it's calculated
--- | ---
Max Call Open Interest and Strike Price | Highest Call Open Interest (in lacs) and it's corresponding Strike Price
Max Put Open Interest and Strike Price | Highest Put Open Interest (in lacs) and it's corresponding Strike Price
Put Call Ratio(PCR) | Total Put Open Interest divided by Total Call Open Interest
Open Interest | This indicates if the latest OI data record indicates Bearish or Bullish signs near Indicated Strike Price. If the Call Sum is more than the Put Sum then the OI is considered Bearish as there is more Call writing than Puts. If the Put Sum is more than the Call sum then the OI Is considered Bullish as the Put writing exceeds the Call writing.
Call Exits | This indicates if the Call writers are exiting near given Strike Price in the latest OI data record. If the Call sum is < 0 or if the change in Call OI at the Call boundary (2 Strike Prices above the given Strike Price) is < 0, then Call writers are exiting their positions and the Bulls have a clear path.
Put Exits | This indicates if the Put writers are exiting near given Strike Price in the latest OI data record. If the Put sum is < 0 or if the change in Put OI at the Put boundary (the given Strike Price) is < 0, then Put writers are exiting their positions and the Bears have a clear path.
Call In The Money(ITM) | This indicates if the Call writers are also exiting far OTM strike prices (4 Strike Prices above the given Strike Price) showing extreme bullishness. Conditions are if the Call writers are exiting their far OTM positions and the Put writers are writing at the same Strike Price. This signal also changes to Yes if the change in Call OI at the far OTM is < 0.
Put In The Money(ITM) | This indicates if the Put writers are also exiting far OTM strike prices (2 Strike Prices below the given Strike Price) showing extreme bearishness. Conditions are if the Put writers are exiting their far OTM positions and the Call writers are writing at the same Strike Price. This signal also changes to Yes if the change in Put OI at the far OTM is < 0.


## Screenshots:

![Screenshot_1](https://i.imgur.com/JEUKcMp.png)

![Screenshot_2](https://i.imgur.com/rwJeMmT.png)

![Screenshot_3](https://i.imgur.com/O4kNI2Y.png)

![Screenshot_4](https://i.imgur.com/Hwbep1G.png)
