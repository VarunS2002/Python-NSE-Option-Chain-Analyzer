# Python NSE-Option-Chain-Analyzer

## [Downloads](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases)
[![New-Site: v3.4](https://img.shields.io/badge/New--Site-v3.4-brightgreen)](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.4)
[![Old-Site: v2.0](https://img.shields.io/badge/Old--Site-v2.0-brightgreen)](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/2.0)
![Download-Count](https://img.shields.io/github/downloads/VarunS2002/Python-NSE-Option-Chain-Analyzer/total?color=blue)
![Build: passing](https://img.shields.io/badge/build-passing-brightgreen)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

For doing technical analysis for option traders, the Option Chain is the most important tool for deciding entry and exit strategies.
The National Stock Exchange (NSE) has a website which displays the option chain for traders in near real-time. This program scrapes this data from the NSE site and then generates useful analysis of the Option Chain for the specified Security or Index from the NSE website.
It also continuously refreshes the Option Chain and visually displays the trend in various indicators and useful for Technical Analysis.

## Installation:
Easy Indtallation process
-Types of variants available:
 
 1. `.py` (Python Source Code)
 
 2. `.pyw` (Compiled Python file without Console)
 
 3. `.exe` (Windows Executable)

-Requirements for 3:
 
 - Windows OS  

-Requirements for 1 and 2:
 
 - Python 3.6+ 
 
 - For Windows https://www.python.org/downloads/ is recommended

 - Add Python to PATH/Environment Variables during installation in Windows (recommended)
 
 - To run the Compiled Python file with Console change the extension to `.pyc` 

 - Required modules:

    ```
    sys
    datetime
    webbrowser
    json
    csv
    tkinter
    tksheet
    pandas
    requests
    streamtologger
    ```
    
  - Install missing modules using `pip install module_name`

## Usage:

1. Select your Index or Security option and it's Expiry Date

2. Enter your preferred Strike Price 

3. Click Start

## Note:

-If there is an error in fetching dates then try refreshing

-In you face any issue then feel free to open an issue. 

-It is recommended to enable logging and then send the nse.log file or the console output  

-In case of network or connection errors the program doesn't crash and will keep retrying until manually stopped

-If a ZeroDivisionError occurs or some data doesn't exist the value of the variable will be defaulted to 0

-All data is retrieved from `https://www.nseindia.com/api/option-chain-indices?symbol=*index_name*`

-[stream-to-logger](https://pypi.org/project/streamtologger/) is used for debug logging

-[auto-py-to-exe](https://pypi.org/project/auto-py-to-exe/) is used for compiling the program to a .exe file

## Features:

-The program continuously retrieves and refreshes the option chain every minute giving near real-time analysis to the traders

-New data row is added only if the NSE server updates its time or data (To prevent printing duplicate data)

-Supported Indices :
 * NIFTY
 * BANKNIFTY
 * NIFTYIT

-Red and Green colour indication for data based on trends

-Program title format: `NSE-Option-Chain-Analyzer - {index} - {expiry_date} - {strike_price}`

-Stop and Start functionality

-You can select all table data using Ctrl+A or select individual cells, rows and columns 

-Then you can copy it using Ctrl+C or right click menu

-You can then paste it in any spreadsheet application (Tested with MS Excel and Google Sheets)

-Export to .csv file option

-Debug Logging toggle 

-About window with version and links for developer GitHub profile, README, license, releases and sources

-PEP 8 format

-Object Oriented

-Table Data displayed:

Data | How it's calculated
--- | ---
Server Time | *Web Scraped*. Indicates last data update time by NSE server 
Value | *Web Scraped*. Underlying Instrument Value indicates the value of the underlying Security or Index
Call Sum | Calculated. Sum of the Changes in Call Open Interest contracts of the given Strike Price and the next immediate two Strike Prices (In Thousands)
Put Sum | Calculated. Sum of the Change in Put Open Interest contracts of the given Strike Price and the next two Strike Prices (In Thousands)
Difference | Calculated. Difference between the Call Sum and Put Sum. If its very -ve its bullish, if its very +ve then its bearish else its a sideways session. 
Call Boundary | Change in Call Open Interest contracts for 2 Strike Prices above the given Strike Price. This is used to determine if Call writers are taking new positions (Bearish) or exiting their positions (Bullish). (In Thousands)
Put Boundary | Change in Put Open Interest for the given Strike Price. This is used to determine if Put writers are taking new positions (Bullish) or exiting their positions(Bearish). (In Thousands)
Call In The Money(ITM) | This indicates that bullish trend could continue and Value could cross 4 Strike Prices above given Strike Price. It's calculated as the ratio of Put writing and Call writing at the 4th Strike Price above the given Strike price. If the absolute ratio > 1.5 then its bullish sign. 
Put In The Money(ITM) | This indicates that bearish trend could continue and Value could cross 2 Strike Prices below given Strike Price. It's calculated as the ratio of Call writing and Put writing at the 2nd Strike Price below the given Strike price. If the absolute ratio > 1.5 then its bearish sign. 

-Label Data displayed:

Data | How it's calculated
--- | ---
Max Call Open Interest and Strike Price | Highest Call Open Interest contracts (in thousands) and it's corresponding Strike Price
Max Put Open Interest and Strike Price | Highest Put Open Interest contracts (in thousands) and it's corresponding Strike Price
Put Call Ratio(PCR) | Sum Total of Put Open Interest contracts divided by Sum Total of Call Open Interest contracts
Open Interest | This indicates if the latest OI data record indicates Bearish or Bullish signs near Indicated Strike Price. If the Call Sum is more than the Put Sum then the OI is considered Bearish as there is more Call writing than Puts. If the Put Sum is more than the Call sum then the OI Is considered Bullish as the Put writing exceeds the Call writing.
Call Exits | This indicates if the Call writers are exiting near given Strike Price in the latest OI data record. If the Call sum is < 0 or if the change in Call OI at the Call boundary (2 Strike Prices above the given Strike Price) is < 0, then Call writers are exiting their positions and the Bulls have a clear path.
Put Exits | This indicates if the Put writers are exiting near given Strike Price in the latest OI data record. If the Put sum is < 0 or if the change in Put OI at the Put boundary (the given Strike Price) is < 0, then Put writers are exiting their positions and the Bears have a clear path.
Call In The Money(ITM) | This indicates if the Call writers are also exiting far OTM strike prices (4 Strike Prices above the given Strike Price) showing extreme bullishness. Conditions are if the Call writers are exiting their far OTM positions and the Put writers are writing at the same Strike Price & if the absolute ratio > 1.5 then its bullish sign. This signal also changes to Yes if the change in Call OI at the far OTM is < 0.
Put In The Money(ITM) | This indicates if the Put writers are also exiting far OTM strike prices (2 Strike Prices below the given Strike Price) showing extreme bearishness. Conditions are if the Put writers are exiting their far OTM positions and the Call writers are writing at the same Strike Price & if the absolute ratio > 1.5 then its a bearish sign. . This signal also changes to Yes if the change in Put OI at the far OTM is < 0.


## Screenshots:

-Login Page:

![Screenshot_1](https://i.imgur.com/2heigvk.png)

-Main Window:

![Screenshot_2](https://i.imgur.com/2JO5BuT.png)

-Selecting data:

![Screenshot_3](https://i.imgur.com/wxoEyPZ.png)

-Option Menu

![Screenshot_4](https://i.imgur.com/wWTLWK6.png)
