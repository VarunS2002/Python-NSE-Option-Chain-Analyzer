# Python NSE-Option-Chain-Analyzer

## [Downloads](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases)
[![Latest: v5.0](https://img.shields.io/badge/release-v5.0-brightgreen)](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/download/5.0/NSE_Option_Chain_Analyzer_5.0.exe)
![Download-Count](https://img.shields.io/github/downloads/VarunS2002/Python-NSE-Option-Chain-Analyzer/total?color=blue)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

For doing technical analysis for option traders, the Option Chain is the most important tool for deciding entry and exit strategies.
The National Stock Exchange (NSE) has a website which displays the option chain for traders in near real-time. This program retrieves this data from the NSE site and then generates useful analysis of the Option Chain for the specified Index or Stock from the NSE website.
It also continuously refreshes the Option Chain and visually displays the trend in various indicators and useful for Technical Analysis.
Calculations are based on [Mr. Sameer Dharaskar's Course](http://advancesharetrading.com/).

## Installation:

>#### Types of variants available:
 
 1. `.exe` (Windows Executable)

 2. `.py` (Python Source Code)

- Does not support Linux

- Requirements for 1:
 
    - Windows OS  

- Requirements for 2:
 
     - Python 3.6+ 
     
     - For Windows https://www.python.org/downloads/ is recommended
    
     - Add Python to PATH/Environment Variables during installation in Windows (recommended)
    
     - Required modules: [requirements.txt](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/blob/master/requirements.txt)
       
     - Install missing modules using `pip install -r requirements.txt`

## Usage:

1. Set Index Mode or Stock Mode

2. Select your Index or Stock 
   
3. Select it's Expiry Date

4. Enter your preferred Strike Price 

5. Set the interval you want the program to refresh (Optional : Defaults to 1 minute)

6. Click Start

## Note:

- If there is an error in fetching dates then try refreshing

- If you face any issue or have a suggestion then feel free to open an issue. 

- It is recommended to enable logging and then send the NSE-OCA.log file or the console output for reporting issues  

- In case of network or connection errors the program doesn't crash and will keep retrying until manually stopped

- If a ZeroDivisionError occurs or some data doesn't exist the value of the variable will be defaulted to 0

## Dependencies:

- [auto-py-to-exe](https://pypi.org/project/auto-py-to-exe/) is used for compiling the program to a .exe file

- [numpy](https://pypi.org/project/numpy/) is used for data types

- [pandas](https://pypi.org/project/pandas/) is used for storing and manipulating the data

- [requests](https://pypi.org/project/requests/) is used for accessing and retrieving data from the NSE website 

- [stream-to-logger](https://pypi.org/project/streamtologger/) is used for debug logging

- [tksheet](https://pypi.org/project/tksheet/) is used for the table containing the data

- [win10toast](https://pypi.org/project/win10toast/) is used for Windows Toast notifications

## Features:

- The program continuously retrieves and refreshes the option chain giving near real-time analysis to the traders

- New data rows are added only if the NSE server updates its time or data (To prevent displaying duplicate data)

- Supported Indices :
     * NIFTY
     * BANKNIFTY
     * FINNIFTY

- Supported Stocks :
  * AARTIIND
  * ACC
  * ADANIENT
  * ADANIPORTS
  * AMARAJABAT
  * AMBUJACEM
  * APOLLOHOSP
  * APOLLOTYRE
  * ASHOKLEY
  * ASIANPAINT
  * AUROPHARMA
  * AXISBANK
  * BAJAJ-AUTO
  * BAJAJFINSV
  * BAJFINANCE
  * BALKRISIND
  * BANDHANBNK
  * BANKBARODA
  * BATAINDIA
  * BEL
  * BERGEPAINT
  * BHARATFORG
  * BHARTIARTL
  * BHEL
  * BIOCON
  * BOSCHLTD
  * BPCL
  * BRITANNIA
  * CADILAHC
  * CANBK
  * CHOLAFIN
  * CIPLA
  * COALINDIA
  * COFORGE
  * COLPAL
  * CONCOR
  * CUMMINSIND
  * DABUR
  * DIVISLAB
  * DLF
  * DRREDDY
  * EICHERMOT
  * ESCORTS
  * EXIDEIND
  * FEDERALBNK
  * GAIL
  * GLENMARK
  * GMRINFRA
  * GODREJCP
  * GODREJPROP
  * GRASIM
  * HAVELLS
  * HCLTECH
  * HDFC
  * HDFCAMC
  * HDFCBANK
  * HDFCLIFE
  * HEROMOTOCO
  * HINDALCO
  * HINDPETRO
  * HINDUNILVR
  * IBULHSGFIN
  * ICICIBANK
  * ICICIGI
  * ICICIPRULI
  * IDEA
  * IDFCFIRSTB
  * IGL
  * INDIGO
  * INDUSINDBK
  * INDUSTOWER
  * INFRATEL
  * INFY
  * IOC
  * ITC
  * JINDALSTEL
  * JSWSTEEL
  * JUBLFOOD
  * KOTAKBANK
  * L&TFH
  * LALPATHLAB
  * LICHSGFIN
  * LT
  * LUPIN
  * M&M
  * M&MFIN
  * MANAPPURAM
  * MARICO
  * MARUTI
  * MCDOWELL-N
  * MFSL
  * MGL
  * MINDTREE
  * MOTHERSUMI
  * MRF
  * MUTHOOTFIN
  * NATIONALUM
  * NAUKRI
  * NESTLEIND
  * NMDC
  * NTPC
  * ONGC
  * PAGEIND
  * PEL
  * PETRONET
  * PFC
  * PIDILITIND
  * PNB
  * POWERGRID
  * PVR
  * RAMCOCEM
  * RBLBANK
  * RECLTD
  * RELIANCE
  * SAIL
  * SBILIFE
  * SBIN
  * SHREECEM
  * SIEMENS
  * SRF
  * SRTRANSFIN
  * SUNPHARMA
  * SUNTV
  * TATACHEM
  * TATACONSUM
  * TATAMOTORS
  * TATAPOWER
  * TATASTEEL
  * TCS
  * TECHM
  * TITAN
  * TORNTPHARM
  * TORNTPOWER
  * TVSMOTOR
  * UBL
  * ULTRACEMCO
  * UPL
  * VEDL
  * VOLTAS
  * WIPRO
  * ZEEL

- Supports multiple instances with different indices/stocks and/or strike prices selected

- Red and Green colour indication for data based on trends

- Toast Notifications for notifying when trend changes. Notified changes:
     * Open Interest: Bullish/Bearish
     * Open Interest Upper Boundary Strike Prices: Change in Value
     * Open Interest Lower Boundary Strike Prices: Change in Value
     * Call Exits: Yes/No
     * Put Exits: Yes/No
     * Call ITM: Yes/No
     * Put ITM: Yes/No

- Program title format: `NSE-Option-Chain-Analyzer - {index/stock} - {expiry_date} - {strike_price}`

- Stop and Start manually

- You can select all table data using Ctrl+A or select individual cells, rows and columns 

- Then you can copy it using Ctrl+C or right click menu

- You can then paste it in any spreadsheet application (Tested with MS Excel and Google Sheets)

- Export table data to .csv file

- Real time exporting data rows to .csv file

- Dumping entire Option Chain data to a .csv file

- Auto stop the program at 3:30pm when the market closes

- Auto Checking for updates

- Debug Logging

- Saves certain settings in a configuration file for subsequent runs. Saved Settings:
     * Index/Stock Mode
     * Selected Index
     * Selected Stock
     * Refresh Interval
     * Live Export
     * Notifications
     * Dump entire Option Chain
     * Auto stop at 3:30pm
     * Auto Check for Updates
     * Debug Logging

- Keyboard shortcuts for all options

## Data Displayed

>#### Table Data:

Data | Description
--- | ---
Server Time | *Web Scraped*. Indicates last data update time by NSE server 
Value | *Web Scraped*. Underlying Instrument Value indicates the value of the underlying Security or Index
Call Sum | Calculated. Sum of the Changes in Call Open Interest contracts of the given Strike Price and the next immediate two Strike Prices (In Thousands for Index Mode and Tens for Stock Mode)
Put Sum | Calculated. Sum of the Change in Put Open Interest contracts of the given Strike Price and the next two Strike Prices (In Thousands for Index Mode and Tens for Stock Mode)
Difference | Calculated. Difference between the Call Sum and Put Sum. If its very -ve its bullish, if its very +ve then its bearish else its a sideways session. 
Call Boundary | Change in Call Open Interest contracts for 2 Strike Prices above the given Strike Price. This is used to determine if Call writers are taking new positions (Bearish) or exiting their positions (Bullish). (In Thousands for Index Mode and Tens for Stock Mode)
Put Boundary | Change in Put Open Interest for the given Strike Price. This is used to determine if Put writers are taking new positions (Bullish) or exiting their positions(Bearish). (In Thousands for Index Mode and Tens for Stock Mode)
Call In The Money(ITM) | This indicates that bullish trend could continue and Value could cross 4 Strike Prices above given Strike Price. It's calculated as the ratio of Put writing and Call writing at the 4th Strike Price above the given Strike price. If the absolute ratio > 1.5 then its bullish sign. 
Put In The Money(ITM) | This indicates that bearish trend could continue and Value could cross 2 Strike Prices below given Strike Price. It's calculated as the ratio of Call writing and Put writing at the 2nd Strike Price below the given Strike price. If the absolute ratio > 1.5 then its bearish sign. 

>#### Label Data:

Data | Description
--- | ---
Open Interest Upper Boundary | Highest and 2nd Highest(highest in OI boundary range) Call Open Interest contracts (In Thousands for Index Mode and Tens for Stock Mode) and their corresponding Strike Prices
Open Interest Lower Boundary | Highest and 2nd Highest(highest in OI boundary range) Put Open Interest contracts (In Thousands for Index Mode and Tens for Stock Mode) and their corresponding Strike Prices
Open Interest | This indicates if the latest OI data record indicates Bearish or Bullish signs near Indicated Strike Price. If the Call Sum is more than the Put Sum then the OI is considered Bearish as there is more Call writing than Puts. If the Put Sum is more than the Call sum then the OI Is considered Bullish as the Put writing exceeds the Call writing.
Put Call Ratio(PCR) | Sum Total of Put Open Interest contracts divided by Sum Total of Call Open Interest contracts
Call Exits | This indicates if the Call writers are exiting near given Strike Price in the latest OI data record. If the Call sum is < 0 or if the change in Call OI at the Call boundary (2 Strike Prices above the given Strike Price) is < 0, then Call writers are exiting their positions and the Bulls have a clear path.
Put Exits | This indicates if the Put writers are exiting near given Strike Price in the latest OI data record. If the Put sum is < 0 or if the change in Put OI at the Put boundary (the given Strike Price) is < 0, then Put writers are exiting their positions and the Bears have a clear path.
Call In The Money(ITM) | This indicates if the Call writers are also exiting far OTM strike prices (4 Strike Prices above the given Strike Price) showing extreme bullishness. Conditions are if the Call writers are exiting their far OTM positions and the Put writers are writing at the same Strike Price & if the absolute ratio > 1.5 then its bullish sign. This signal also changes to Yes if the change in Call OI at the far OTM is < 0.
Put In The Money(ITM) | This indicates if the Put writers are also exiting far OTM strike prices (2 Strike Prices below the given Strike Price) showing extreme bearishness. Conditions are if the Put writers are exiting their far OTM positions and the Call writers are writing at the same Strike Price & if the absolute ratio > 1.5 then its a bearish sign. . This signal also changes to Yes if the change in Put OI at the far OTM is < 0.


## Screenshots:

- Login Page:

  <br>![Login_Window](https://i.imgur.com/TfETQkz.png) <br><br>

- Main Window Index Mode:

  <br>![Main_Window_Index](https://i.imgur.com/JHn58gn.png) <br><br>

- Main Window Stock Mode:
  
  <br>![Main_Window_Stock](https://i.imgur.com/jwL1zrU.png) <br><br>

- Selecting Data:

  <br>![Selecting_Data](https://i.imgur.com/qYoy2iO.png) <br><br>

- Option Menu

  <br>![Option_Menu](https://i.imgur.com/jtrjCvY.png)
