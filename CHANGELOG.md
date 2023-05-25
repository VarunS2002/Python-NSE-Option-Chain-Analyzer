# CHANGELOG

<br>

> # [5.5](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/5.5)

## Bug Fix Update

- Fixed `IndexError` caused by NSE updating the method they use for loading symbols on their website.
  Issue: [#54](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/54)
  , [#55](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/55)
  , [#56](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/56)
- No longer uses the web scraping technique to load symbols
    - Instead uses the API to load symbols as that is what the new NSE website uses
    - Removed `beautifulsoup4` from `requirements.txt` as it is no longer required
- Updated documentation:
    - Removed `beautifulsoup4` from dependencies

<br>

> # [5.4](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/5.4)

## Bug Fix Update

- Fixed `IndexError` with tksheet versions `>=5.0.29`.
  Issue: [#24](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/24)
  , [#25](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/25)
  , [#28](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/28)
  , [#32](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/32)
  and [#33](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/33). Pull
  Request: [#34](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/pull/34) (Thanks
  to [@yjagota](https://github.com/yjagota/))
- Fixed redundant code. Pull Request: [#19](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/pull/19)
  (Thanks to [@QuickLearner171998](https://github.com/QuickLearner171998/))
- Bumped dependencies to fix known vulnerabilities
- Updated documentation:
    - Added new contributors to the list
    - Updated table formatting
    - Fixed typos

<br>

> # [5.3](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/5.3)

## Feature and Bug Fix Update

- New App Icon:

  <img width="64" height="64" src="https://i.imgur.com/OGHZnUu.png" alt="icon_square">

- Improved Documentation
- Added Option to alert if last time the data was updated is 5 minutes or more called `Warn Late Server Updates`
- Downloading the app icon can be disabled now changing the value of `load_nse_icon` to `False` in the config file
    - This is to speed up loading times while using the .py version
- Added Saving New settings for subsequent runs
    - New Saved settings:
        * Load App Icon
        * Warn Late Server Updates
- Improved Logging:
    - New Logging messages:
        * Beta status
        * Beta version number
- Disabled resizability for About window
- Bumped dependencies for `.py` version
- Removed numpy import by casting all numpy.float64 and numpy.int64 types to float and int
    - Removed numpy from requirements but will be installed by pandas anyways
- Fixed icon not loading in `.exe` version
- Prevents crashing the program if it fails to check for updates
- Prevents crashing the program if icon fails to load correctly
    - Runs the program with the default icon instead
- Program will exit if an error occurs while trying to load symbols
    - Error message window will be displayed before exiting

<br>

> # [5.2](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/5.2)

## Compatibility and Feature Update

- Added Support for Linux
    - Notifications only supports Windows 10
- List of Stocks and Indices is loaded dynamically
  from https://www.nseindia.com/products-services/equity-derivatives-list-underlyings-information everytime
    - Program is now not required to be updated when there are changes in Stocks and Indices
- Added name of stock or index in notifications
- Reworked configuration reading and writing
    - Settings will stay persistent across updates
    - Only invalid values are corrected and missing values are appended
    - Unless a user tampers with it, the config does not reset completely
- If your running the .py version, icons are now temporarily downloaded every time you run the program and need not be
  saved separately
- Disabled text input in Dropdown menus on Login page to prevent errors
- Improved Logging:
    - New Logging messages:
        * OS Name and Version
        * Error Class along with Error Message
- Note: Users who are still running version 4.x are recommended to directly update to 5.2 to prevent settings from
  resetting

<br>

> # [5.0](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/5.0)

## Major Feature Update

- Added Support for Stocks (Total 141 Stocks).
  Issue: [#8](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/8)
    - Set Index Mode or Stock Mode then select your preferred Index or Stock
    - Values are in 10s in Stock mode and 1000s in Index mode (to compensate for low values of stocks)
    - Symbol of the Stock is present in the names of the .csv files
    - Appropriate Units are reflected in the headers of .csv files
- Added Saving New settings for subsequent runs
    - New Saved settings:
        * Index/Stock Mode
        * Selected Stock
- Fixed possible issues while reading configuration
- Optimised retrieving data from the API

<br>

> # [4.1](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/4.1)

## Feature and Bug Fix Update

- Added Dumping Entire Option Chain data to a .csv file.
  Issues: [#3](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/3)
  and [#4](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/4)
    - Dump Entire Option Chain is disabled by default (Enable from Option menu or Ctrl+O)
    - Saves this setting for subsequent runs
- Added Notifications for changes in value of OI Upper and Lower Boundary Strike Prices
- Renamed 'Export all to CSV' option to 'Export Table to CSV'
- Fixed Call and Put OI for 2nd Strike Price not being displayed in K when the Strike Prices were consecutive
- Fixed issues where export would fail and program would stop if the .csv file is open in some other program or is
  inaccessible
- Fixed issue where program stops immediately if you start it before market opens when you have auto stop enabled
- Prevents crash during Checking for updates due to poor internet connection
- Fixed 'Quitting Program' being logged even if Debug Logging was off
- Fixed possible issues while reading configuration

<br>

> # [4.0](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/4.0)

## Major Feature Update

- Added support for FINNIFTY index
- Dropped support for NIFTYIT index
- Added Live Exporting of Data rows to a .csv file
    - Live Exporting is disabled by default (Enable from Option menu or Ctrl+B)
- Supports exporting data while running multiple instances with different indices and/or expiry dates selected
    - Filename contains the selected index and expiry date. For eg. NSE-OCA-NIFTY-14-Jan-2021.csv will only have the
      data for NIFTY and 14 Jan 21 regardless of the instance running
- Adds Column Names to the .csv file if it is created for the first time
- Added Toast Notifications on Windows when a state of a label changes (except PCR label)
    - Notifications are disabled by default (Enable from Option menu or Ctrl+N)
- Added option to automatically stop the program at 3:30pm when market closes
    - Auto Stop is disabled by default (Enable from Option menu or Ctrl+K)
- Added Auto and Manual Check for updates
    - Auto Check for updates are enabled by default (Disable from Option menu or Ctrl+U)
- Added Saving settings for subsequent runs
    - Saved settings:
        * Selected Index
        * Refresh Interval
        * Live Export
        * Notifications
        * Auto stop at 3:30pm
        * Auto Check for Updates
        * Debug Logging
    - Settings are saved to NSE-OCA.ini
    - Resets NSE-OCA.ini file if incorrectly configured
- Fixed issue where Points would be 0 for some strike prices.
  Issue: [#6](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/issues/6)
- Added NSE icon to all windows (if icon file is missing, default icon will be used)
- Improved Option Menu
- Improved messages in Alert boxes
- Improved Buttons
- Modified some Labels
- Improved Logging:
    - New Logging messages:
        * Whether running instance is .py version or .exe
        * Version number
        * Logging Started
        * Logging Stopped
        * Program Quitting
    - Removed unnecessary messages:
        * 'Nse' object has no attribute 'options' 10
        * module 'sys' has no attribute '_MEIPASS' 0
        * invalid command name ".!combobox2" 4
    - Changed name of the log file from nse.log to NSE-OCA.log
- Many Code Improvements

<br>

> # [3.7](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.7)

## Feature Update

- Added 2nd Highest Call and Put Open Interest and their corresponding Strike Price
    - It is calculated between the OI boundary range (highest in the range)
- Added option to change refresh interval
- Reworked Login Screen
- Added Type Hints in code everywhere
- Added requirements.txt
- Reduced size of .exe by ~10%
- Ceased releasing Python Compiled Files (.pyc/.pyw)
    - Since it runs only on specific versions of Python

<br>

> # [3.5](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.5)

## Bug Fix Update

- Fixed an issue when program would stop refreshing after a few hours (creates a new session everytime)

<br>

> # [3.4](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.4)

## Bug Fix Update

- Fixed an issue when program would stop refreshing after every ~2 hours (creates a new session everytime)
- Fixed an issue when the program would stop refreshing after one connection error
- Added Debug Logging (Use this to report any issues)

<br>

> # [3.3](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.3)

## Bug Fix Update

- Now uses sessions and cookies to access the website to solve many connection errors
- Added units in bottom labels
- Changed label "Points" to "Value"
- Refactored code
- Drastically reduced .exe size (~98MB to ~35MB)

<br>

> # [3.2](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.2)

## Bug Fix Update

- Fixed an issue when the program would stop responding if the internet connection is poor
- Fixed an issue when no data would be retrieved if the program is run before the website updates for the first time on
  a day (~9:15 am)
- Fixed an issue when no data would be retrieved if the expiry date selected expires when the program is running. This
  now throws an error and stops the program.
- Added units in column labels
- Slightly reduced the width of the program's main window

<br>

> # [3.1](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.1)

## Bug Fix and .exe Update

- Fixed error when the entered strike price is towards the upper limit or the lower limit of the table
- The missing values will be set to 0
- Used to auto-py-to-exe to compile the program to a .exe file (beta)
- README.md updated

<br>

> # [3.0](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/3.0)

## Major Compatibility Update

- Now compatible with the new NSE website
- Instead of scraping the data from the html the program now calculates the data from a json file which is also the
  implementation of the new website (Thanks to [@medknecth](https://github.com/medknecth/))
- Since the values on the new website display contracts instead of shares, the values in the program have been updated
  to display in thousands instead of lacs
- Completely reworked the main code
- Updated Login window
- Dropped support for shares
- Now only supports the following indexes: NIFTY, BANKNIFTY and NIFTYIT
- Refreshes every 1 min now, increased from the earlier 30 sec
- Requires a new module: json
- Instructions about the data is included in README.md
- Object Oriented
- Miscellaneous bug fixes
- Since the new website disallows and is made for preventing web scraping, you may encounter more connection errors

<br>

> # [2.0](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/2.0)

## Major Feature Update

- Added GUI
- Red and Green colour indication for data based on trends
- Added Stop and Start functionality
- Added Export to .csv option
- Added About window with version and links for developer GitHub profile, README, license, releases and sources
- Instructions about the data is included in README.md

<br>

> # [1.0](https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/tag/1.0)

## First Release

- Enter the final URL from the browser and your preferred Strike Price
- In case of network or connection errors the program doesn't crash and will keep retrying infinitely
- If a ZeroDivisionError occurs or some data doesn't exist the value of the variable will be defaulted to 0
- The program infinitely refreshes every 30 seconds
- New data is printed only if the server time has changed (To prevent printing duplicate data)
- Check README.md to see what data is printed and how it's calculated
- PEP 8 format
