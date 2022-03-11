import configparser
import csv
import datetime
import os
import platform
import sys
import time
import webbrowser
from tkinter import Tk, Toplevel, Event, TclError, StringVar, Frame, Menu, Label, Entry, SOLID, RIDGE, \
    DISABLED, NORMAL, N, S, E, W, LEFT, messagebox, PhotoImage
from tkinter.ttk import Combobox, Button
from typing import Union, Optional, List, Dict, Tuple, TextIO, Any

import bs4
import pandas
import requests
import streamtologger
import tksheet

is_windows: bool = platform.system() == "Windows"
is_windows_10: bool = is_windows and platform.release() == "10"
if is_windows_10:
    # noinspection PyUnresolvedReferences
    import win10toast


# noinspection PyAttributeOutsideInit
class Nse:
    version: str = '5.4'
    beta: Tuple[bool, int] = (False, 0)

    def __init__(self, window: Tk) -> None:
        self.intervals: List[int] = [1, 2, 3, 5, 10, 15]
        self.stdout: TextIO = sys.stdout
        self.stderr: TextIO = sys.stderr
        self.previous_date: Optional[datetime.date] = None
        self.previous_time: Optional[datetime.time] = None
        self.time_difference_factor: int = 5
        self.first_run: bool = True
        self.stop: bool = False
        self.dates: List[str] = [""]
        self.indices: List[str] = []
        self.stocks: List[str] = []
        self.url_oc: str = "https://www.nseindia.com/option-chain"
        self.url_index: str = "https://www.nseindia.com/api/option-chain-indices?symbol="
        self.url_stock: str = "https://www.nseindia.com/api/option-chain-equities?symbol="
        self.url_symbols: str = "https://www.nseindia.com/products-services/" \
                                "equity-derivatives-list-underlyings-information"
        self.url_icon_png: str = "https://raw.githubusercontent.com/VarunS2002/" \
                                 "Python-NSE-Option-Chain-Analyzer/master/nse_logo.png"
        self.url_icon_ico: str = "https://raw.githubusercontent.com/VarunS2002/" \
                                 "Python-NSE-Option-Chain-Analyzer/master/nse_logo.ico"
        self.url_update: str = "https://api.github.com/repos/VarunS2002/" \
                               "Python-NSE-Option-Chain-Analyzer/releases/latest"
        self.headers: Dict[str, str] = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, '
                          'like Gecko) Chrome/80.0.3987.149 Safari/537.36',
            'accept-language': 'en,gu;q=0.9,hi;q=0.8',
            'accept-encoding': 'gzip, deflate, br'}
        self.get_symbols(window)
        self.config_parser: configparser.ConfigParser = configparser.ConfigParser()
        self.create_config(new=True) if not os.path.isfile('NSE-OCA.ini') else None
        self.get_config()
        self.log() if self.logging else None
        self.units_str: str = 'in K' if self.option_mode == 'Index' else 'in 10s'
        self.output_columns: Tuple[str, str, str, str, str, str, str, str, str] = (
            'Time', 'Value', f'Call Sum\n({self.units_str})', f'Put Sum\n({self.units_str})',
            f'Difference\n({self.units_str})', f'Call Boundary\n({self.units_str})',
            f'Put Boundary\n({self.units_str})', 'Call ITM', 'Put ITM')
        self.csv_headers: Tuple[str, str, str, str, str, str, str, str, str] = (
            'Time', 'Value', f'Call Sum ({self.units_str})', f'Put Sum ({self.units_str})',
            f'Difference ({self.units_str})',
            f'Call Boundary ({self.units_str})', f'Put Boundary ({self.units_str})', 'Call ITM', 'Put ITM')
        self.session: requests.Session = requests.Session()
        self.cookies: Dict[str, str] = {}
        self.toaster: win10toast.ToastNotifier = win10toast.ToastNotifier() if is_windows_10 else None
        self.get_icon()
        self.login_win(window)

    def get_symbols(self, window: Tk) -> None:
        def create_error_window(error_window: Tk):
            error_window.title("NSE-Option-Chain-Analyzer")
            window_width: int = error_window.winfo_reqwidth()
            window_height: int = error_window.winfo_reqheight()
            position_right: int = int(error_window.winfo_screenwidth() / 2 - window_width / 2)
            position_down: int = int(error_window.winfo_screenheight() / 2 - window_height / 2)
            error_window.geometry("320x160+{}+{}".format(position_right, position_down))
            messagebox.showerror(title="Error", message="Failed to fetch Symbols.\nThe program will now exit.")
            error_window.destroy()

        try:
            symbols_information: requests.Response = requests.get(self.url_symbols, headers=self.headers)
        except Exception as err:
            print(err, sys.exc_info()[0], "19")
            create_error_window(window)
            sys.exit()
        symbols_information_soup: bs4.BeautifulSoup = bs4.BeautifulSoup(symbols_information.content, "html.parser")
        try:
            symbols_table: bs4.element.Tag = symbols_information_soup.findChildren('table')[0]
        except IndexError as err:
            print(err, sys.exc_info()[0], "20")
            create_error_window(window)
            sys.exit()
        symbols_table_rows: List[bs4.element.Tag] = list(symbols_table.findChildren(['th', 'tr']))
        symbols_table_rows_str: List[str] = ['' for _ in range(len(symbols_table_rows) - 1)]
        for column in range(len(symbols_table_rows) - 1):
            symbols_table_rows_str[column] = str(symbols_table_rows[column])
        divider_row: str = '<tr>\n' \
                           '<td colspan="3"><strong>Derivatives on Individual Securities</strong></td>\n' \
                           '</tr>'
        for column in range(4, symbols_table_rows_str.index(divider_row) + 1):
            cells: bs4.element.ResultSet = symbols_table_rows[column].findChildren('td')
            column: int = 0
            for cell in cells:
                if column == 2:
                    self.indices.append(cell.string)
                column += 1
        for column in reversed(range(symbols_table_rows_str.index(divider_row) + 1)):
            symbols_table_rows.pop(column)
        for row in symbols_table_rows:
            cells: bs4.element.ResultSet = row.findChildren('td')
            column: int = 0
            for cell in cells:
                if column == 2:
                    self.stocks.append(cell.string)
                column += 1

    def get_icon(self) -> None:
        self.icon_png_path: str
        self.icon_ico_path: str
        try:
            # noinspection PyProtectedMember,PyUnresolvedReferences
            base_path: str = sys._MEIPASS
            self.icon_png_path = os.path.join(base_path, 'nse_logo.png')
            self.icon_ico_path = os.path.join(base_path, 'nse_logo.ico')
            self.load_nse_icon = True
        except AttributeError:
            if self.load_nse_icon:
                try:
                    icon_png_raw: requests.Response = requests.get(self.url_icon_png, headers=self.headers, stream=True)
                    with open('.NSE-OCA.png', 'wb') as f:
                        for chunk in icon_png_raw.iter_content(1024):
                            f.write(chunk)
                    self.icon_png_path = '.NSE-OCA.png'
                    PhotoImage(file=self.icon_png_path)
                except Exception as err:
                    print(err, sys.exc_info()[0], "17")
                    self.load_nse_icon = False
                    return
                if is_windows_10:
                    try:
                        icon_ico_raw: requests.Response = requests.get(self.url_icon_ico,
                                                                       headers=self.headers, stream=True)
                        with open('.NSE-OCA.ico', 'wb') as f:
                            for chunk in icon_ico_raw.iter_content(1024):
                                f.write(chunk)
                        self.icon_ico_path = '.NSE-OCA.ico'
                    except Exception as err:
                        print(err, sys.exc_info()[0], "18")
                        self.icon_ico_path = None
                        return

    def check_for_updates(self, auto: bool = True) -> None:
        try:
            release_data: requests.Response = requests.get(self.url_update, headers=self.headers, timeout=5)
            latest_version: str = release_data.json()['tag_name']
            float(latest_version)
        except Exception as err:
            print(err, sys.exc_info()[0], "21")
            if not auto:
                self.info.attributes('-topmost', False)
                messagebox.showerror(title="Error", message="Failed to check for updates.")
                self.info.attributes('-topmost', True)
            return

        if float(latest_version) > float(Nse.version):
            self.info.attributes('-topmost', False) if not auto else None
            update: bool = messagebox.askyesno(
                title="New Update Available",
                message=f"You are running version: {Nse.version}\n"
                        f"Latest version: {latest_version}\n"
                        f"Do you want to update now ?\n"
                        f"{'You can disable auto check for updates from the menu.' if auto and self.update else ''}")
            if update:
                webbrowser.open_new("https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/latest")
                self.info.attributes('-topmost', False) if not auto else None
            else:
                self.info.attributes('-topmost', True) if not auto else None
        else:
            if not auto:
                self.info.attributes('-topmost', False)
                messagebox.showinfo(title="No Updates Available", message=f"You are running the latest version.\n"
                                                                          f"Version: {Nse.version}")
                self.info.attributes('-topmost', True)

    def get_config(self) -> None:
        try:
            self.config_parser.read('NSE-OCA.ini')
            try:
                self.load_nse_icon: bool = self.config_parser.getboolean('main', 'load_nse_icon')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="load_nse_icon")
                self.load_nse_icon: bool = self.config_parser.getboolean('main', 'load_nse_icon')
            try:
                self.index: str = self.config_parser.get('main', 'index')
                if self.index not in self.indices:
                    raise ValueError(f'{self.index} is not a valid index')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="index")
                self.index: str = self.config_parser.get('main', 'index')
            try:
                self.stock: str = self.config_parser.get('main', 'stock')
                if self.stock not in self.stocks:
                    raise ValueError(f'{self.stock} is not a valid stock')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="stock")
                self.stock: str = self.config_parser.get('main', 'stock')
            try:
                self.option_mode: str = self.config_parser.get('main', 'option_mode')
                if self.option_mode not in ('Index', 'Stock'):
                    raise ValueError(f'{self.option_mode} is not a valid option mode')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="option_mode")
                self.option_mode: str = self.config_parser.get('main', 'option_mode')
            try:
                self.seconds: int = self.config_parser.getint('main', 'seconds')
                if self.seconds not in (60, 120, 180, 300, 600, 900):
                    raise ValueError(f'{self.seconds} is not a refresh interval')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="seconds")
                self.seconds: int = self.config_parser.getint('main', 'seconds')
            try:
                self.live_export: bool = self.config_parser.getboolean('main', 'live_export')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="live_export")
                self.live_export: bool = self.config_parser.getboolean('main', 'live_export')
            try:
                self.save_oc: bool = self.config_parser.getboolean('main', 'save_oc')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="save_oc")
                self.save_oc: bool = self.config_parser.getboolean('main', 'save_oc')
            try:
                self.notifications: bool = self.config_parser.getboolean('main', 'notifications') \
                    if is_windows_10 else False
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="notifications")
                self.notifications: bool = self.config_parser.getboolean('main', 'notifications') \
                    if is_windows_10 else False
            try:
                self.auto_stop: bool = self.config_parser.getboolean('main', 'auto_stop')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="auto_stop")
                self.auto_stop: bool = self.config_parser.getboolean('main', 'auto_stop')
            try:
                self.update: bool = self.config_parser.getboolean('main', 'update')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="update")
                self.update: bool = self.config_parser.getboolean('main', 'update')
            try:
                self.logging: bool = self.config_parser.getboolean('main', 'logging')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="logging")
                self.logging: bool = self.config_parser.getboolean('main', 'logging')
            try:
                self.warn_late_update: bool = self.config_parser.getboolean('main', 'warn_late_update')
            except (configparser.NoOptionError, ValueError) as err:
                print(err, sys.exc_info()[0], "0")
                self.create_config(attribute="warn_late_update")
                self.warn_late_update: bool = self.config_parser.getboolean('main', 'warn_late_update')
        except (configparser.NoSectionError, configparser.MissingSectionHeaderError,
                configparser.DuplicateSectionError, configparser.DuplicateOptionError) as err:
            print(err, sys.exc_info()[0], "0")
            self.create_config(corrupted=True)
            return self.get_config()

    def create_config(self, new: bool = False, corrupted: bool = False, attribute: Optional[str] = None) -> None:
        if new or corrupted:
            if corrupted:
                os.remove('NSE-OCA.ini')
                self.config_parser = configparser.ConfigParser()
            self.config_parser.read('NSE-OCA.ini')
            self.config_parser.add_section('main')
            self.config_parser.set('main', 'load_nse_icon', 'True')
            self.config_parser.set('main', 'index', self.indices[0])
            self.config_parser.set('main', 'stock', self.stocks[0])
            self.config_parser.set('main', 'option_mode', 'Index')
            self.config_parser.set('main', 'seconds', '60')
            self.config_parser.set('main', 'live_export', 'False')
            self.config_parser.set('main', 'save_oc', 'False')
            self.config_parser.set('main', 'notifications', 'False')
            self.config_parser.set('main', 'auto_stop', 'False')
            self.config_parser.set('main', 'update', 'True')
            self.config_parser.set('main', 'logging', 'False')
            self.config_parser.set('main', 'warn_late_update', 'False')
        elif attribute is not None:
            if attribute == "load_nse_icon":
                self.config_parser.set('main', 'load_nse_icon', 'True')
            elif attribute == "index":
                self.config_parser.set('main', 'index', self.indices[0])
            elif attribute == "stock":
                self.config_parser.set('main', 'stock', self.stocks[0])
            elif attribute == "option_mode":
                self.config_parser.set('main', 'option_mode', 'Index')
            elif attribute == "seconds":
                self.config_parser.set('main', 'seconds', '60')
            elif attribute in ("live_export", "save_oc", "notifications", "auto_stop", "logging"):
                self.config_parser.set('main', attribute, 'False')
            elif attribute == "update":
                self.config_parser.set('main', 'update', 'True')
            elif attribute == "warn_late_update":
                self.config_parser.set('main', 'warn_late_update', 'False')

        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def get_data(self, event: Optional[Event] = None) -> Optional[Tuple[Optional[requests.Response], Any]]:
        if self.first_run:
            return self.get_data_first_run()
        else:
            return self.get_data_refresh()

    def get_data_first_run(self) -> Optional[Tuple[Optional[requests.Response], Any]]:
        request: Optional[requests.Response] = None
        response: Optional[requests.Response] = None
        self.units_str = 'in K' if self.option_mode == 'Index' else 'in 10s'
        self.output_columns: Tuple[str, str, str, str, str, str, str, str, str] = (
            'Time', 'Value', f'Call Sum\n({self.units_str})', f'Put Sum\n({self.units_str})',
            f'Difference\n({self.units_str})', f'Call Boundary\n({self.units_str})',
            f'Put Boundary\n({self.units_str})', 'Call ITM', 'Put ITM')
        self.csv_headers: Tuple[str, str, str, str, str, str, str, str, str] = (
            'Time', 'Value', f'Call Sum ({self.units_str})', f'Put Sum ({self.units_str})',
            f'Difference ({self.units_str})',
            f'Call Boundary ({self.units_str})', f'Put Boundary ({self.units_str})', 'Call ITM', 'Put ITM')
        self.round_factor: int = 1000 if self.option_mode == 'Index' else 10
        if self.option_mode == 'Index':
            self.index = self.index_var.get()
            self.config_parser.set('main', 'index', f'{self.index}')
        else:
            self.stock = self.stock_var.get()
            self.config_parser.set('main', 'stock', f'{self.stock}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

        url: str = self.url_index + self.index if self.option_mode == 'Index' else self.url_stock + self.stock
        try:
            request = self.session.get(self.url_oc, headers=self.headers, timeout=5)
            self.cookies = dict(request.cookies)
            response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
        except Exception as err:
            print(request)
            print(response)
            print(err, sys.exc_info()[0], "1")
            messagebox.showerror(title="Error", message="Error in fetching dates.\nPlease retry.")
            self.dates.clear()
            self.dates = [""]
            self.date_menu.config(values=tuple(self.dates))
            self.date_menu.current(0)
            return
        json_data: Any
        if response is not None:
            try:
                json_data = response.json()
            except Exception as err:
                print(response)
                print(err, sys.exc_info()[0], "2")
                json_data = {}
        else:
            json_data = {}
        if json_data == {}:
            messagebox.showerror(title="Error", message="Error in fetching dates.\nPlease retry.")
            self.dates.clear()
            self.dates = [""]
            try:
                self.date_menu.config(values=tuple(self.dates))
                self.date_menu.current(0)
            except TclError as err:
                print(err, sys.exc_info()[0], "3")
            return
        self.dates.clear()
        for dates in json_data['records']['expiryDates']:
            self.dates.append(dates)
        try:
            self.date_menu.config(values=tuple(self.dates))
            self.date_menu.current(0)
        except TclError:
            pass

        return response, json_data

    def get_data_refresh(self) -> Optional[Tuple[Optional[requests.Response], Any]]:
        request: Optional[requests.Response] = None
        response: Optional[requests.Response] = None
        url: str = self.url_index + self.index if self.option_mode == 'Index' else self.url_stock + self.stock
        try:
            response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
            if response.status_code == 401:
                self.session.close()
                self.session = requests.Session()
                request = self.session.get(self.url_oc, headers=self.headers, timeout=5)
                self.cookies = dict(request.cookies)
                response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
                print("reset cookies")
        except Exception as err:
            print(request)
            print(response)
            print(err, sys.exc_info()[0], "4")
            try:
                self.session.close()
                self.session = requests.Session()
                request = self.session.get(self.url_oc, headers=self.headers, timeout=5)
                self.cookies = dict(request.cookies)
                response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
                print("reset cookies")
            except Exception as err:
                print(request)
                print(response)
                print(err, sys.exc_info()[0], "5")
                return
        if response is not None:
            try:
                json_data: Any = response.json()
            except Exception as err:
                print(response)
                print(err, sys.exc_info()[0], "6")
                json_data = {}
        else:
            json_data = {}
        if json_data == {}:
            return

        return response, json_data

    def login_win(self, window: Tk) -> None:
        self.login: Tk = window
        self.login.title("NSE-Option-Chain-Analyzer")
        self.login.protocol('WM_DELETE_WINDOW', self.close_login)
        window_width: int = self.login.winfo_reqwidth()
        window_height: int = self.login.winfo_reqheight()
        position_right: int = int(self.login.winfo_screenwidth() / 2 - window_width / 2)
        position_down: int = int(self.login.winfo_screenheight() / 2 - window_height / 2)
        self.login.geometry("320x160+{}+{}".format(position_right, position_down))
        self.login.resizable(False, False)
        self.login.iconphoto(True, PhotoImage(file=self.icon_png_path)) if self.load_nse_icon else None
        self.login.rowconfigure(0, weight=1)
        self.login.rowconfigure(1, weight=1)
        self.login.rowconfigure(2, weight=1)
        self.login.rowconfigure(3, weight=1)
        self.login.rowconfigure(4, weight=1)
        self.login.rowconfigure(5, weight=1)
        self.login.columnconfigure(0, weight=1)
        self.login.columnconfigure(1, weight=1)
        self.login.columnconfigure(2, weight=1)

        self.intervals_var: StringVar = StringVar()
        self.intervals_var.set(str(self.intervals[0]))
        self.index_var: StringVar = StringVar()
        self.index_var.set(self.indices[0])
        self.stock_var: StringVar = StringVar()
        self.stock_var.set(self.stocks[0])
        self.dates_var: StringVar = StringVar()
        self.dates_var.set(self.dates[0])

        option_mode_label: Label = Label(self.login, text="Mode: ", justify=LEFT)
        option_mode_label.grid(row=0, column=0, sticky=N + S + W)
        self.option_mode_btn: Button = Button(self.login, text=f"{'Index' if self.option_mode == 'Index' else 'Stock'}",
                                              command=self.change_option_mode, width=10)
        self.option_mode_btn.grid(row=0, column=1, sticky=N + S + E + W)
        index_label: Label = Label(self.login, text="Index: ", justify=LEFT)
        index_label.grid(row=1, column=0, sticky=N + S + W)
        self.index_menu: Combobox = Combobox(self.login, textvariable=self.index_var, values=self.indices)
        self.index_menu.config(width=15, state='readonly' if self.option_mode == 'Index' else DISABLED)
        self.index_menu.grid(row=1, column=1, sticky=N + S + E)
        self.index_menu.current(self.indices.index(self.index))
        stock_label: Label = Label(self.login, text="Stock: ", justify=LEFT)
        stock_label.grid(row=2, column=0, sticky=N + S + W)
        self.stock_menu: Combobox = Combobox(self.login, textvariable=self.stock_var, values=self.stocks)
        self.stock_menu.config(width=15, state='readonly' if self.option_mode == 'Stock' else DISABLED)
        self.stock_menu.grid(row=2, column=1, sticky=N + S + E)
        self.stock_menu.current(self.stocks.index(self.stock))
        date_label: Label = Label(self.login, text="Expiry Date: ", justify=LEFT)
        date_label.grid(row=3, column=0, sticky=N + S + W)
        self.date_menu: Combobox = Combobox(self.login, textvariable=self.dates_var, state="readonly")
        self.date_menu.config(width=15)
        self.date_menu.grid(row=3, column=1, sticky=N + S + E)
        self.date_get: Button = Button(self.login, text="Refresh", command=self.get_data, width=10)
        self.date_get.grid(row=3, column=2, sticky=N + S + E + W)
        sp_label: Label = Label(self.login, text="Strike Price (eg. 14750): ")
        sp_label.grid(row=4, column=0, sticky=N + S + W)
        self.sp_entry = Entry(self.login, width=18, relief=SOLID)
        self.sp_entry.grid(row=4, column=1, sticky=N + S + E)
        start_btn: Button = Button(self.login, text="Start", command=self.start, width=10)
        start_btn.grid(row=4, column=2, rowspan=2, sticky=N + S + E + W)
        intervals_label: Label = Label(self.login, text="Refresh Interval (in min): ", justify=LEFT)
        intervals_label.grid(row=5, column=0, sticky=N + S + W)
        self.intervals_menu: Combobox = Combobox(self.login, textvariable=self.intervals_var,
                                                 values=[str(interval) for interval in self.intervals],
                                                 state="readonly")
        self.intervals_menu.config(width=15)
        self.intervals_menu.grid(row=5, column=1, sticky=N + S + E)
        self.intervals_menu.current(self.intervals.index(int(self.seconds / 60)))
        self.sp_entry.focus_set()
        self.get_data()

        # noinspection PyUnusedLocal
        def focus_widget(event: Event, mode: int) -> None:
            if mode == 1:
                self.get_data()
                self.date_menu.focus_set()
            elif mode == 2:
                self.sp_entry.focus_set()

        self.index_menu.bind('<Return>', lambda event, a=1: focus_widget(event, a))
        self.index_menu.bind("<<ComboboxSelected>>", self.get_data)
        self.stock_menu.bind('<Return>', lambda event, a=1: focus_widget(event, a))
        self.stock_menu.bind("<<ComboboxSelected>>", self.get_data)
        self.date_menu.bind('<Return>', lambda event, a=2: focus_widget(event, a))
        self.sp_entry.bind('<Return>', self.start)

        self.login.mainloop()

    def change_option_mode(self) -> None:
        if self.option_mode_btn['text'] == 'Index':
            self.option_mode = 'Stock'
            self.option_mode_btn.config(text='Stock')
            self.index_menu.config(state=DISABLED)
            self.stock_menu.config(state='readonly')
        else:
            self.option_mode = 'Index'
            self.option_mode_btn.config(text='Index')
            self.index_menu.config(state='readonly')
            self.stock_menu.config(state=DISABLED)

        self.get_data()

        self.config_parser.set('main', 'option_mode', f'{self.option_mode}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def start(self, event: Optional[Event] = None) -> None:
        self.seconds = int(self.intervals_var.get()) * 60
        self.config_parser.set('main', 'seconds', f'{self.seconds}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)
        self.expiry_date: str = self.dates_var.get()
        if self.expiry_date == "":
            messagebox.showerror(title="Error", message="Incorrect Expiry Date.\nPlease enter correct Expiry Date.")
            return
        if self.live_export:
            self.export_row(None)
        try:
            self.sp: int = int(self.sp_entry.get())
            self.login.destroy()
            self.main_win()
        except ValueError as err:
            print(err, sys.exc_info()[0], "7")
            messagebox.showerror(title="Error", message="Incorrect Strike Price.\nPlease enter correct Strike Price.")

    # noinspection PyUnusedLocal
    def change_state(self, event: Optional[Event] = None) -> None:

        if not self.stop:
            self.stop = True
            self.options.entryconfig(self.options.index(0), label="Start")
            messagebox.showinfo(title="Stopped", message="Retrieving new data has been stopped.")
        else:
            self.stop = False
            self.options.entryconfig(self.options.index(0), label="Stop")
            messagebox.showinfo(title="Started", message="Retrieving new data has been started.")

            self.main()

    # noinspection PyUnusedLocal
    def export(self, event: Optional[Event] = None) -> None:
        sheet_data: List[List[str]] = self.sheet.get_sheet_data()
        csv_exists: bool = os.path.isfile(
            f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-{self.expiry_date}.csv")
        try:
            if not csv_exists:
                with open(f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-{self.expiry_date}.csv",
                          "a", newline="") as row:
                    data_writer: csv.writer = csv.writer(row)
                    data_writer.writerow(self.csv_headers)

            with open(f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-{self.expiry_date}.csv",
                      "a", newline="") as row:
                data_writer: csv.writer = csv.writer(row)
                data_writer.writerows(sheet_data)

            messagebox.showinfo(title="Export Successful",
                                message=f"Data has been exported to NSE-OCA-"
                                        f"{self.index if self.option_mode == 'Index' else self.stock}-"
                                        f"{self.expiry_date}.csv.")
        except PermissionError as err:
            print(err, sys.exc_info()[0], "12")
            messagebox.showerror(title="Export Failed",
                                 message=f"Failed to access NSE-OCA-"
                                         f"{self.index if self.option_mode == 'Index' else self.stock}-"
                                         f"{self.expiry_date}.csv.\n"
                                         f"Permission Denied. Try closing any apps using it.")
        except Exception as err:
            print(err, sys.exc_info()[0], "8")
            messagebox.showerror(title="Export Failed",
                                 message="An error occurred while exporting the data.")

    def export_row(self, values: Optional[List[Union[str, float]]]) -> None:
        if values is None:
            csv_exists: bool = os.path.isfile(
                f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-{self.expiry_date}.csv")
            try:
                if not csv_exists:
                    with open(
                            f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-"
                            f"{self.expiry_date}.csv",
                            "a", newline="") as row:
                        data_writer: csv.writer = csv.writer(row)
                        data_writer.writerow(self.csv_headers)
            except PermissionError as err:
                print(err, sys.exc_info()[0], "13")
                messagebox.showerror(title="Export Failed",
                                     message=f"Failed to access NSE-OCA-"
                                             f"{self.index if self.option_mode == 'Index' else self.stock}-"
                                             f"{self.expiry_date}.csv.\n"
                                             f"Permission Denied. Try closing any apps using it.")
            except Exception as err:
                print(err, sys.exc_info()[0], "9")
        else:
            try:
                with open(f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-{self.expiry_date}.csv",
                          "a", newline="") as row:
                    data_writer: csv.writer = csv.writer(row)
                    data_writer.writerow(values)
            except PermissionError as err:
                print(err, sys.exc_info()[0], "14")
                messagebox.showerror(title="Export Failed",
                                     message=f"Failed to access NSE-OCA-"
                                             f"{self.index if self.option_mode == 'Index' else self.stock}-"
                                             f"{self.expiry_date}.csv.\n"
                                             f"Permission Denied. Try closing any apps using it.")
            except Exception as err:
                print(err, sys.exc_info()[0], "15")

    # noinspection PyUnusedLocal
    def toggle_live_export(self, event: Optional[Event] = None) -> None:
        if self.live_export:
            self.live_export = False
            self.options.entryconfig(self.options.index(2), label="Live Exporting to CSV: Off")
            messagebox.showinfo(title="Live Exporting Disabled",
                                message="Data rows will not be exported.")
        else:
            self.live_export = True
            self.options.entryconfig(self.options.index(2), label="Live Exporting to CSV: On")
            messagebox.showinfo(title="Live Exporting Enabled",
                                message=f"Data rows will be exported in real time to "
                                        f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-"
                                        f"{self.expiry_date}.csv.")

        self.config_parser.set('main', 'live_export', f'{self.live_export}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)
        self.export_row(None)

    # noinspection PyUnusedLocal
    def toggle_save_oc(self, event: Optional[Event] = None) -> None:
        if self.save_oc:
            self.save_oc = False
            self.options.entryconfig(self.options.index(3), label="Dump Entire Option Chain to CSV: Off")
            messagebox.showinfo(title="Dump Entire Option Chain Disabled",
                                message=f"Entire Option Chain data will not be exported.")
        else:
            self.save_oc = True
            self.options.entryconfig(self.options.index(3), label="Dump Entire Option Chain to CSV: On")
            messagebox.showinfo(title="Dump Entire Option Chain Enabled",
                                message=f"Entire Option Chain data will be exported to "
                                        f"NSE-OCA-"
                                        f"{self.index if self.option_mode == 'Index' else self.stock}-"
                                        f"{self.expiry_date}-Full.csv.")

        self.config_parser.set('main', 'save_oc', f'{self.save_oc}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def toggle_notifications(self, event: Optional[Event] = None) -> None:
        if self.notifications:
            self.notifications = False
            self.options.entryconfig(self.options.index(4), label="Notifications: Off")
            messagebox.showinfo(title="Notifications Disabled",
                                message="You will not receive any Notifications.")
        else:
            self.notifications = True
            self.options.entryconfig(self.options.index(4), label="Notifications: On")
            messagebox.showinfo(title="Notifications Enabled",
                                message="You will receive Notifications when the state of a label changes.")

        self.config_parser.set('main', 'notifications', f'{self.notifications}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def toggle_auto_stop(self, event: Optional[Event] = None) -> None:
        if self.auto_stop:
            self.auto_stop = False
            self.options.entryconfig(self.options.index(5), label="Stop automatically at 3:30pm: Off")
            messagebox.showinfo(title="Auto Stop Disabled", message="Program will not automatically stop at 3:30pm")
        else:
            self.auto_stop = True
            self.options.entryconfig(self.options.index(5), label="Stop automatically at 3:30pm: On")
            messagebox.showinfo(title="Auto Stop Enabled", message="Program will automatically stop at 3:30pm")

        self.config_parser.set('main', 'auto_stop', f'{self.auto_stop}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def toggle_warn_late_update(self, event: Optional[Event] = None) -> None:
        if self.warn_late_update:
            self.warn_late_update = False
            self.options.entryconfig(self.options.index(6), label="Warn Late Server Updates: Off")
            messagebox.showinfo(title="Warn Late Server Updates Disabled",
                                message="Program will not alert you if the server updates late.")
        else:
            self.warn_late_update = True
            self.options.entryconfig(self.options.index(6), label="Warn Late Server Updates: On")
            messagebox.showinfo(title="Warn Late Server Updates Enabled",
                                message="Program will alert you if the server update time is 5 minutes or more.")

        self.config_parser.set('main', 'warn_late_update', f'{self.warn_late_update}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def toggle_updates(self, event: Optional[Event] = None) -> None:
        if self.update:
            self.update = False
            self.options.entryconfig(self.options.index(8), label="Auto Check for Updates: Off")
            messagebox.showinfo(title="Auto Checking for Updates Disabled",
                                message="Program will not check for updates at start.")
        else:
            self.update = True
            self.options.entryconfig(self.options.index(8), label="Auto Check for Updates: On")
            messagebox.showinfo(title="Auto Checking for Updates Enabled",
                                message="Program will check for updates at start.")

        self.config_parser.set('main', 'update', f'{self.update}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def log(self, event: Optional[Event] = None) -> None:
        if self.first_run and self.logging or not self.logging:
            streamtologger.redirect(target="NSE-OCA.log",
                                    header_format="[{timestamp:%Y-%m-%d %H:%M:%S} - {level:5}] ")
            self.logging = True
            print('----------Logging Started----------')

            try:
                # noinspection PyProtectedMember,PyUnresolvedReferences
                base_path: str = sys._MEIPASS
                print(platform.system() + ' ' + platform.release() + ' .exe version ' + Nse.version, end=' ')
                if Nse.beta[0]:
                    print(f"beta {Nse.beta[1]}")
                else:
                    print()
            except AttributeError:
                print(platform.system() + ' ' + platform.release() + ' .py version ' + Nse.version, end=' ')
                if Nse.beta[0]:
                    print(f"beta {Nse.beta[1]}")
                else:
                    print()
                if not self.load_nse_icon:
                    print("NSE icon loading disabled")

            try:
                self.options.entryconfig(self.options.index(9), label="Debug Logging: On")
                messagebox.showinfo(title="Debug Logging Enabled",
                                    message="Errors will be logged to NSE-OCA.log.")
            except AttributeError:
                pass
        elif self.logging:
            print('----------Logging Stopped----------')
            sys.stdout = self.stdout
            sys.stderr = self.stderr
            streamtologger._is_redirected = False
            self.logging = False
            self.options.entryconfig(self.options.index(9), label="Debug Logging: Off")
            messagebox.showinfo(title="Debug Logging Disabled", message="Errors will not be logged.")

        self.config_parser.set('main', 'logging', f'{self.logging}')
        with open('NSE-OCA.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def links(self, link: str, event: Optional[Event] = None) -> None:

        if link == "developer":
            webbrowser.open_new("https://github.com/VarunS2002/")
        elif link == "readme":
            webbrowser.open_new("https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/blob/master/README.md/")
        elif link == "license":
            webbrowser.open_new("https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/blob/master/LICENSE/")
        elif link == "releases":
            webbrowser.open_new("https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/releases/")
        elif link == "sources":
            webbrowser.open_new("https://github.com/VarunS2002/Python-NSE-Option-Chain-Analyzer/")

        self.info.attributes('-topmost', False)

    def about_window(self) -> Toplevel:
        self.info: Toplevel = Toplevel()
        self.info.title("About")
        window_width: int = self.info.winfo_reqwidth()
        window_height: int = self.info.winfo_reqheight()
        position_right: int = int(self.info.winfo_screenwidth() / 2 - window_width / 2)
        position_down: int = int(self.info.winfo_screenheight() / 2 - window_height / 2)
        self.info.geometry("250x150+{}+{}".format(position_right, position_down))
        self.info.resizable(False, False)
        self.info.iconphoto(True, PhotoImage(file=self.icon_png_path)) if self.load_nse_icon else None
        self.info.attributes('-topmost', True)
        self.info.grab_set()
        self.info.focus_force()

        return self.info

    # noinspection PyUnusedLocal
    def about(self, event: Optional[Event] = None) -> None:
        self.info: Toplevel = self.about_window()
        self.info.rowconfigure(0, weight=1)
        self.info.rowconfigure(1, weight=1)
        self.info.rowconfigure(2, weight=1)
        self.info.rowconfigure(3, weight=1)
        self.info.rowconfigure(4, weight=1)
        self.info.columnconfigure(0, weight=1)
        self.info.columnconfigure(1, weight=1)

        heading: Label = Label(self.info, text="NSE-Option-Chain-Analyzer", relief=RIDGE,
                               font=("TkDefaultFont", 10, "bold"))
        heading.grid(row=0, column=0, columnspan=2, sticky=N + S + W + E)
        version_label: Label = Label(self.info, text="Version:", relief=RIDGE)
        version_label.grid(row=1, column=0, sticky=N + S + W + E)
        version_val: Label = Label(self.info, text=f"{Nse.version}", relief=RIDGE)
        version_val.grid(row=1, column=1, sticky=N + S + W + E)
        dev_label: Label = Label(self.info, text="Developer:", relief=RIDGE)
        dev_label.grid(row=2, column=0, sticky=N + S + W + E)
        dev_val: Label = Label(self.info, text="Varun Shanbhag", fg="blue", cursor="hand2", relief=RIDGE)
        dev_val.bind("<Button-1>", lambda click, link="developer": self.links(link, click))
        dev_val.grid(row=2, column=1, sticky=N + S + W + E)
        readme: Label = Label(self.info, text="README", fg="blue", cursor="hand2", relief=RIDGE)
        readme.bind("<Button-1>", lambda click, link="readme": self.links(link, click))
        readme.grid(row=3, column=0, sticky=N + S + W + E)
        licenses: Label = Label(self.info, text="LICENSE", fg="blue", cursor="hand2", relief=RIDGE)
        licenses.bind("<Button-1>", lambda click, link="license": self.links(link, click))
        licenses.grid(row=3, column=1, sticky=N + S + W + E)
        releases: Label = Label(self.info, text="Releases", fg="blue", cursor="hand2", relief=RIDGE)
        releases.bind("<Button-1>", lambda click, link="releases": self.links(link, click))
        releases.grid(row=4, column=0, sticky=N + S + W + E)
        sources: Label = Label(self.info, text="Sources", fg="blue", cursor="hand2", relief=RIDGE)
        sources.bind("<Button-1>", lambda click, link="sources": self.links(link, click))
        sources.grid(row=4, column=1, sticky=N + S + W + E)
        updates: Button = Button(self.info, text="Check for Updates",
                                 command=lambda auto=False: self.check_for_updates(auto))
        updates.grid(row=5, column=0, columnspan=2, sticky=N + S + W + E)
        self.info.mainloop()

    def close_login(self) -> None:
        self.session.close()
        if self.logging:
            print('----------Quitting Program----------')
        os.remove('.NSE-OCA.png') if os.path.isfile('.NSE-OCA.png') else None
        os.remove('.NSE-OCA.ico') if os.path.isfile('.NSE-OCA.ico') else None
        self.login.destroy()
        sys.exit()

    # noinspection PyUnusedLocal
    def close_main(self, event: Optional[Event] = None) -> None:
        ask_quit: bool = messagebox.askyesno("Quit", "All unsaved data will be lost.\nProceed to quit?", icon='warning',
                                             default='no')
        if ask_quit:
            self.session.close()
            if self.logging:
                print('----------Quitting Program----------')
            os.remove('.NSE-OCA.png') if os.path.isfile('.NSE-OCA.png') else None
            os.remove('.NSE-OCA.ico') if os.path.isfile('.NSE-OCA.ico') else None
            self.root.destroy()
            sys.exit()
        elif not ask_quit:
            pass

    def main_win(self) -> None:
        self.root: Tk = Tk()
        self.root.focus_force()
        self.root.title("NSE-Option-Chain-Analyzer")
        self.root.protocol('WM_DELETE_WINDOW', self.close_main)
        window_width: int = self.root.winfo_reqwidth()
        window_height: int = self.root.winfo_reqheight()
        position_right: int = int(self.root.winfo_screenwidth() / 3 - window_width / 2)
        position_down: int = int(self.root.winfo_screenheight() / 3 - window_height / 2)
        self.root.geometry("815x560+{}+{}".format(position_right, position_down))
        self.root.iconphoto(True, PhotoImage(file=self.icon_png_path)) if self.load_nse_icon else None
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        menubar: Menu = Menu(self.root)
        self.options: Menu = Menu(menubar, tearoff=0)
        self.options.add_command(label="Stop", accelerator="(Ctrl+X)", command=self.change_state)
        self.options.add_command(label="Export Table to CSV", accelerator="(Ctrl+S)", command=self.export)
        self.options.add_command(label=f"Live Exporting to CSV: {'On' if self.live_export else 'Off'}",
                                 accelerator="(Ctrl+B)", command=self.toggle_live_export)
        self.options.add_command(label=f"Dump Entire Option Chain to CSV: {'On' if self.save_oc else 'Off'}",
                                 accelerator="(Ctrl+O)", command=self.toggle_save_oc)
        self.options.add_command(label=f"Notifications: {'On' if self.notifications else 'Off'}",
                                 accelerator="(Ctrl+N)", command=self.toggle_notifications,
                                 state=NORMAL if is_windows_10 else DISABLED)
        self.options.add_command(label=f"Stop automatically at 3:30pm: {'On' if self.auto_stop else 'Off'}",
                                 accelerator="(Ctrl+K)", command=self.toggle_auto_stop)
        self.options.add_command(label=f"Warn Late Server Updates: {'On' if self.warn_late_update else 'Off'}",
                                 accelerator="(Ctrl+W)", command=self.toggle_warn_late_update)
        self.options.add_separator()
        self.options.add_command(label=f"Auto Check for Updates: {'On' if self.update else 'Off'}",
                                 accelerator="(Ctrl+U)", command=self.toggle_updates)
        self.options.add_command(label=f"Debug Logging: {'On' if self.logging else 'Off'}", accelerator="(Ctrl+L)",
                                 command=self.log)
        self.options.add_command(label="About", accelerator="(Ctrl+M)", command=self.about)
        self.options.add_command(label="Quit", accelerator="(Ctrl+Q)", command=self.close_main)
        menubar.add_cascade(label="Menu", menu=self.options)
        self.root.config(menu=menubar)

        self.root.bind('<Control-x>', self.change_state)
        self.root.bind('<Control-s>', self.export)
        self.root.bind('<Control-b>', self.toggle_live_export)
        self.root.bind('<Control-o>', self.toggle_save_oc)
        self.root.bind('<Control-n>', self.toggle_notifications) if is_windows_10 else None
        self.root.bind('<Control-k>', self.toggle_auto_stop)
        self.root.bind('<Control-w>', self.toggle_warn_late_update)
        self.root.bind('<Control-u>', self.toggle_updates)
        self.root.bind('<Control-l>', self.log)
        self.root.bind('<Control-m>', self.about)
        self.root.bind('<Control-q>', self.close_main)

        top_frame: Frame = Frame(self.root)
        top_frame.rowconfigure(0, weight=1)
        top_frame.columnconfigure(0, weight=1)
        top_frame.pack(fill="both", expand=True)

        self.sheet: tksheet.Sheet = tksheet.Sheet(top_frame, column_width=85, align="center",
                                                  headers=self.output_columns, header_font=("TkDefaultFont", 9, "bold"),
                                                  empty_horizontal=0, empty_vertical=20, header_height=35)
        self.sheet.enable_bindings(
            ("toggle_select", "drag_select", "column_select", "row_select", "column_width_resize",
             "arrowkeys", "right_click_popup_menu", "rc_select", "copy", "select_all"))
        self.sheet.grid(row=0, column=0, sticky=N + S + W + E)

        bottom_frame: Frame = Frame(self.root)
        bottom_frame.rowconfigure(0, weight=1)
        bottom_frame.rowconfigure(1, weight=1)
        bottom_frame.rowconfigure(2, weight=1)
        bottom_frame.rowconfigure(3, weight=1)
        bottom_frame.rowconfigure(4, weight=1)
        bottom_frame.rowconfigure(5, weight=1)
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)
        bottom_frame.columnconfigure(2, weight=1)
        bottom_frame.columnconfigure(3, weight=1)
        bottom_frame.columnconfigure(4, weight=1)
        bottom_frame.columnconfigure(5, weight=1)
        bottom_frame.columnconfigure(6, weight=1)
        bottom_frame.columnconfigure(7, weight=1)
        bottom_frame.pack(fill="both", expand=True)

        oi_ub_label: Label = Label(bottom_frame, text="Open Interest Upper Boundary", relief=RIDGE,
                                   font=("TkDefaultFont", 10, "bold"))
        oi_ub_label.grid(row=0, column=0, columnspan=4, sticky=N + S + W + E)
        max_call_oi_sp_label: Label = Label(bottom_frame, text="Strike Price 1:", relief=RIDGE,
                                            font=("TkDefaultFont", 9, "bold"))
        max_call_oi_sp_label.grid(row=1, column=0, sticky=N + S + W + E)
        self.max_call_oi_sp_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_call_oi_sp_val.grid(row=1, column=1, sticky=N + S + W + E)
        max_call_oi_label: Label = Label(bottom_frame, text=f"OI ({self.units_str}):", relief=RIDGE,
                                         font=("TkDefaultFont", 9, "bold"))
        max_call_oi_label.grid(row=1, column=2, sticky=N + S + W + E)
        self.max_call_oi_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_call_oi_val.grid(row=1, column=3, sticky=N + S + W + E)
        oi_lb_label: Label = Label(bottom_frame, text="Open Interest Lower Boundary", relief=RIDGE,
                                   font=("TkDefaultFont", 10, "bold"))
        oi_lb_label.grid(row=0, column=4, columnspan=4, sticky=N + S + W + E)
        max_put_oi_sp_label: Label = Label(bottom_frame, text="Strike Price 1:", relief=RIDGE,
                                           font=("TkDefaultFont", 9, "bold"))
        max_put_oi_sp_label.grid(row=1, column=4, sticky=N + S + W + E)
        self.max_put_oi_sp_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_put_oi_sp_val.grid(row=1, column=5, sticky=N + S + W + E)
        max_put_oi_label: Label = Label(bottom_frame, text=f"OI ({self.units_str}):", relief=RIDGE,
                                        font=("TkDefaultFont", 9, "bold"))
        max_put_oi_label.grid(row=1, column=6, sticky=N + S + W + E)
        self.max_put_oi_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_put_oi_val.grid(row=1, column=7, sticky=N + S + W + E)
        max_call_oi_sp_2_label: Label = Label(bottom_frame, text="Strike Price 2:", relief=RIDGE,
                                              font=("TkDefaultFont", 9, "bold"))
        max_call_oi_sp_2_label.grid(row=2, column=0, sticky=N + S + W + E)
        self.max_call_oi_sp_2_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_call_oi_sp_2_val.grid(row=2, column=1, sticky=N + S + W + E)
        max_call_oi_2_label: Label = Label(bottom_frame, text=f"OI ({self.units_str}):", relief=RIDGE,
                                           font=("TkDefaultFont", 9, "bold"))
        max_call_oi_2_label.grid(row=2, column=2, sticky=N + S + W + E)
        self.max_call_oi_2_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_call_oi_2_val.grid(row=2, column=3, sticky=N + S + W + E)
        oi_lb_2_label: Label = Label(bottom_frame, text="Open Interest Lower Boundary", relief=RIDGE,
                                     font=("TkDefaultFont", 10, "bold"))
        oi_lb_2_label.grid(row=2, column=4, columnspan=4, sticky=N + S + W + E)
        max_put_oi_sp_2_label: Label = Label(bottom_frame, text="Strike Price 2:", relief=RIDGE,
                                             font=("TkDefaultFont", 9, "bold"))
        max_put_oi_sp_2_label.grid(row=2, column=4, sticky=N + S + W + E)
        self.max_put_oi_sp_2_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_put_oi_sp_2_val.grid(row=2, column=5, sticky=N + S + W + E)
        max_put_oi_2_label: Label = Label(bottom_frame, text=f"OI ({self.units_str}):", relief=RIDGE,
                                          font=("TkDefaultFont", 9, "bold"))
        max_put_oi_2_label.grid(row=2, column=6, sticky=N + S + W + E)
        self.max_put_oi_2_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_put_oi_2_val.grid(row=2, column=7, sticky=N + S + W + E)

        oi_label: Label = Label(bottom_frame, text="Open Interest:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
        oi_label.grid(row=3, column=0, columnspan=2, sticky=N + S + W + E)
        self.oi_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.oi_val.grid(row=3, column=2, columnspan=2, sticky=N + S + W + E)
        pcr_label: Label = Label(bottom_frame, text="PCR:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
        pcr_label.grid(row=3, column=4, columnspan=2, sticky=N + S + W + E)
        self.pcr_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.pcr_val.grid(row=3, column=6, columnspan=2, sticky=N + S + W + E)
        call_exits_label: Label = Label(bottom_frame, text="Call Exits:", relief=RIDGE,
                                        font=("TkDefaultFont", 9, "bold"))
        call_exits_label.grid(row=4, column=0, columnspan=2, sticky=N + S + W + E)
        self.call_exits_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.call_exits_val.grid(row=4, column=2, columnspan=2, sticky=N + S + W + E)
        put_exits_label: Label = Label(bottom_frame, text="Put Exits:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
        put_exits_label.grid(row=4, column=4, columnspan=2, sticky=N + S + W + E)
        self.put_exits_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.put_exits_val.grid(row=4, column=6, columnspan=2, sticky=N + S + W + E)
        call_itm_label: Label = Label(bottom_frame, text="Call ITM:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
        call_itm_label.grid(row=5, column=0, columnspan=2, sticky=N + S + W + E)
        self.call_itm_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.call_itm_val.grid(row=5, column=2, columnspan=2, sticky=N + S + W + E)
        put_itm_label: Label = Label(bottom_frame, text="Put ITM:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
        put_itm_label.grid(row=5, column=4, columnspan=2, sticky=N + S + W + E)
        self.put_itm_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.put_itm_val.grid(row=5, column=6, columnspan=2, sticky=N + S + W + E)

        self.root.after(100, self.main)

        self.root.mainloop()

    def get_dataframe(self) -> Optional[Tuple[pandas.DataFrame, str, float]]:
        try:
            response: Optional[requests.Response]
            json_data: Any
            response, json_data = self.get_data()
        except TypeError:
            return
        if response is None or json_data is None:
            return

        pandas.set_option('display.max_rows', None)
        pandas.set_option('display.max_columns', None)
        pandas.set_option('display.width', 400)

        df: pandas.DataFrame = pandas.read_json(response.text)
        df = df.transpose()

        ce_values: List[dict] = [data['CE'] for data in json_data['records']['data'] if
                                 "CE" in data and data['expiryDate'].lower() == self.expiry_date.lower()]
        pe_values: List[dict] = [data['PE'] for data in json_data['records']['data'] if
                                 "PE" in data and data['expiryDate'].lower() == self.expiry_date.lower()]
        points: float = pe_values[0]['underlyingValue']
        if points == 0:
            for item in pe_values:
                if item['underlyingValue'] != 0:
                    points = item['underlyingValue']
                    break
        ce_data_f: pandas.DataFrame = pandas.DataFrame(ce_values)
        pe_data_f: pandas.DataFrame = pandas.DataFrame(pe_values)

        if ce_data_f.empty:
            messagebox.showerror(title="Error",
                                 message="Invalid Expiry Date.\nPlease restart and enter a new Expiry Date.")
            self.change_state()
            return
        columns_ce: List[str] = ['openInterest', 'changeinOpenInterest', 'totalTradedVolume', 'impliedVolatility',
                                 'lastPrice',
                                 'change', 'bidQty', 'bidprice', 'askPrice', 'askQty', 'strikePrice']
        columns_pe: List[str] = ['strikePrice', 'bidQty', 'bidprice', 'askPrice', 'askQty', 'change', 'lastPrice',
                                 'impliedVolatility', 'totalTradedVolume', 'changeinOpenInterest', 'openInterest']
        ce_data_f = ce_data_f[columns_ce]
        pe_data_f = pe_data_f[columns_pe]
        merged_inner: pandas.DataFrame = pandas.merge(left=ce_data_f, right=pe_data_f, left_on='strikePrice',
                                                      right_on='strikePrice')
        merged_inner.columns = ['Open Interest', 'Change in Open Interest', 'Traded Volume', 'Implied Volatility',
                                'Last Traded Price', 'Net Change', 'Bid Quantity', 'Bid Price', 'Ask Price',
                                'Ask Quantity', 'Strike Price', 'Bid Quantity', 'Bid Price', 'Ask Price',
                                'Ask Quantity', 'Net Change', 'Last Traded Price', 'Implied Volatility',
                                'Traded Volume', 'Change in Open Interest', 'Open Interest']
        current_time: str = df['timestamp']['records']
        return merged_inner, current_time, points

    def set_values(self) -> None:
        if self.first_run:
            self.root.title(f"NSE-Option-Chain-Analyzer - {self.index if self.option_mode == 'Index' else self.stock} "
                            f"- {self.expiry_date} - {self.sp}")

        self.old_max_call_oi_sp: float
        self.old_max_call_oi_sp_2: float
        self.old_max_put_oi_sp: float
        self.old_max_put_oi_sp_2: float

        self.max_call_oi_val.config(text=self.max_call_oi)
        self.max_call_oi_sp_val.config(text=self.max_call_oi_sp)
        self.max_call_oi_2_val.config(text=self.max_call_oi_2)
        self.max_call_oi_sp_2_val.config(text=self.max_call_oi_sp_2)
        self.max_put_oi_val.config(text=self.max_put_oi)
        self.max_put_oi_sp_val.config(text=self.max_put_oi_sp)
        self.max_put_oi_2_val.config(text=self.max_put_oi_2)
        self.max_put_oi_sp_2_val.config(text=self.max_put_oi_sp_2)

        if self.first_run or self.old_max_call_oi_sp == self.max_call_oi_sp:
            self.old_max_call_oi_sp = self.max_call_oi_sp
        else:
            if self.notifications:
                self.toaster.show_toast("Upper Boundary Strike Price changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_max_call_oi_sp} to {self.max_call_oi_sp}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_max_call_oi_sp = self.max_call_oi_sp

        if self.first_run or self.old_max_call_oi_sp_2 == self.max_call_oi_sp_2:
            self.old_max_call_oi_sp_2 = self.max_call_oi_sp_2
        else:
            if self.notifications:
                self.toaster.show_toast("Upper Boundary Strike Price 2 changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_max_call_oi_sp_2} to {self.max_call_oi_sp_2}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_max_call_oi_sp_2 = self.max_call_oi_sp_2

        if self.first_run or self.old_max_put_oi_sp == self.max_put_oi_sp:
            self.old_max_put_oi_sp = self.max_put_oi_sp
        else:
            if self.notifications:
                self.toaster.show_toast("Lower Boundary Strike Price changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_max_put_oi_sp} to {self.max_put_oi_sp}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_max_put_oi_sp = self.max_put_oi_sp

        if self.first_run or self.old_max_put_oi_sp_2 == self.max_put_oi_sp_2:
            self.old_max_put_oi_sp_2 = self.max_put_oi_sp_2
        else:
            if self.notifications:
                self.toaster.show_toast("Lower Boundary Strike Price 2 changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_max_put_oi_sp_2} to {self.max_put_oi_sp_2}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_max_put_oi_sp_2 = self.max_put_oi_sp_2

        red: str = "#e53935"
        green: str = "#00e676"
        default: str = "SystemButtonFace" if is_windows else "#d9d9d9"

        bg: str

        self.old_oi_label: str
        oi_label: str

        if self.call_sum >= self.put_sum:
            oi_label = "Bearish"
            bg = red
        else:
            oi_label = "Bullish"
            bg = green
        self.oi_val.config(text=oi_label, bg=bg)

        if self.first_run or self.old_oi_label == oi_label:
            self.old_oi_label = oi_label
        else:
            if self.notifications:
                self.toaster.show_toast("Open Interest changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_oi_label} to {oi_label}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_oi_label = oi_label

        if self.put_call_ratio >= 1:
            self.pcr_val.config(text=self.put_call_ratio, bg=green)
        else:
            self.pcr_val.config(text=self.put_call_ratio, bg=red)

        def set_itm_labels(call_change: float, put_change: float) -> str:
            label: str = "No"
            if put_change > call_change:
                if put_change >= 0:
                    if call_change <= 0:
                        label = "Yes"
                    elif put_change / call_change > 1.5:
                        label = "Yes"
                else:
                    if put_change / call_change < 0.5:
                        label = "Yes"
            if call_change <= 0:
                label = "Yes"
            return label

        self.old_call_label: str
        call: str = set_itm_labels(call_change=self.p5, put_change=self.p4)

        if call == "No":
            self.call_itm_val.config(text="No", bg=default)
        else:
            self.call_itm_val.config(text="Yes", bg=green)

        if self.first_run or self.old_call_label == call:
            self.old_call_label = call
        else:
            if self.notifications:
                self.toaster.show_toast("Call ITM changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_call_label} to {call}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_call_label = call

        self.old_put_label: str
        put: str = set_itm_labels(call_change=self.p7, put_change=self.p6)

        if put == "No":
            self.put_itm_val.config(text="No", bg=default)
        else:
            self.put_itm_val.config(text="Yes", bg=red)

        if self.first_run or self.old_put_label == put:
            self.old_put_label = put
        else:
            if self.notifications:
                self.toaster.show_toast("Put ITM changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_put_label} to {put}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_put_label = put

        self.old_call_exits_label: str
        call_exits_label: str

        if self.call_boundary <= 0:
            call_exits_label = "Yes"
            bg = green
        elif self.call_sum <= 0:
            call_exits_label = "Yes"
            bg = green
        else:
            call_exits_label = "No"
            bg = default

        self.call_exits_val.config(text=call_exits_label, bg=bg)
        if self.first_run or self.old_call_exits_label == call_exits_label:
            self.old_call_exits_label = call_exits_label
        else:
            if self.notifications:
                self.toaster.show_toast("Call Exits changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_call_exits_label} to {call_exits_label}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_call_exits_label = call_exits_label

        self.old_put_exits_label: str
        put_exits_label: str

        if self.put_boundary <= 0:
            put_exits_label = "Yes"
            bg = red
        elif self.put_sum <= 0:
            put_exits_label = "Yes"
            bg = red
        else:
            put_exits_label = "No"
            bg = default

        self.put_exits_val.config(text=put_exits_label, bg=bg)
        if self.first_run or self.old_put_exits_label == put_exits_label:
            self.old_put_exits_label = put_exits_label
        else:
            if self.notifications:
                self.toaster.show_toast("Put Exits changed "
                                        f"for {self.index if self.option_mode == 'Index' else self.stock}",
                                        f"Changed from {self.old_put_exits_label} to {put_exits_label}",
                                        duration=4, threaded=True,
                                        icon_path=self.icon_ico_path if self.load_nse_icon else None)
            self.old_put_exits_label = put_exits_label

        output_values: List[Union[str, float]] = [self.str_current_time, self.points, self.call_sum,
                                                  self.put_sum, self.difference,
                                                  self.call_boundary, self.put_boundary, self.call_itm,
                                                  self.put_itm]
        self.sheet.insert_row(values=output_values, add_columns=True)
        if self.live_export:
            self.export_row(output_values)

        last_row: int = self.sheet.get_total_rows() - 1

        self.old_points: float
        if self.first_run or self.points == self.old_points:
            self.old_points = self.points
        elif self.points > self.old_points:
            self.sheet.highlight_cells(row=last_row, column=1, bg=green)
            self.old_points = self.points
        else:
            self.sheet.highlight_cells(row=last_row, column=1, bg=red)
            self.old_points = self.points
        self.old_call_sum: float
        if self.first_run or self.old_call_sum == self.call_sum:
            self.old_call_sum = self.call_sum
        elif self.call_sum > self.old_call_sum:
            self.sheet.highlight_cells(row=last_row, column=2, bg=red)
            self.old_call_sum = self.call_sum
        else:
            self.sheet.highlight_cells(row=last_row, column=2, bg=green)
            self.old_call_sum = self.call_sum
        self.old_put_sum: float
        if self.first_run or self.old_put_sum == self.put_sum:
            self.old_put_sum = self.put_sum
        elif self.put_sum > self.old_put_sum:
            self.sheet.highlight_cells(row=last_row, column=3, bg=green)
            self.old_put_sum = self.put_sum
        else:
            self.sheet.highlight_cells(row=last_row, column=3, bg=red)
            self.old_put_sum = self.put_sum
        self.old_difference: float
        if self.first_run or self.old_difference == self.difference:
            self.old_difference = self.difference
        elif self.difference > self.old_difference:
            self.sheet.highlight_cells(row=last_row, column=4, bg=red)
            self.old_difference = self.difference
        else:
            self.sheet.highlight_cells(row=last_row, column=4, bg=green)
            self.old_difference = self.difference
        self.old_call_boundary: float
        if self.first_run or self.old_call_boundary == self.call_boundary:
            self.old_call_boundary = self.call_boundary
        elif self.call_boundary > self.old_call_boundary:
            self.sheet.highlight_cells(row=last_row, column=5, bg=red)
            self.old_call_boundary = self.call_boundary
        else:
            self.sheet.highlight_cells(row=last_row, column=5, bg=green)
            self.old_call_boundary = self.call_boundary
        self.old_put_boundary: float
        if self.first_run or self.old_put_boundary == self.put_boundary:
            self.old_put_boundary = self.put_boundary
        elif self.put_boundary > self.old_put_boundary:
            self.sheet.highlight_cells(row=last_row, column=6, bg=green)
            self.old_put_boundary = self.put_boundary
        else:
            self.sheet.highlight_cells(row=last_row, column=6, bg=red)
            self.old_put_boundary = self.put_boundary
        self.old_call_itm: float
        if self.first_run or self.old_call_itm == self.call_itm:
            self.old_call_itm = self.call_itm
        elif self.call_itm > self.old_call_itm:
            self.sheet.highlight_cells(row=last_row, column=7, bg=green)
            self.old_call_itm = self.call_itm
        else:
            self.sheet.highlight_cells(row=last_row, column=7, bg=red)
            self.old_call_itm = self.call_itm
        self.old_put_itm: float
        if self.first_run or self.old_put_itm == self.put_itm:
            self.old_put_itm = self.put_itm
        elif self.put_itm > self.old_put_itm:
            self.sheet.highlight_cells(row=last_row, column=8, bg=red)
            self.old_put_itm = self.put_itm
        else:
            self.sheet.highlight_cells(row=last_row, column=8, bg=green)
            self.old_put_itm = self.put_itm

        if self.sheet.get_yview()[1] >= 0.9:
            self.sheet.see(last_row)
            self.sheet.set_yview(1)
        self.sheet.refresh()

    def main(self) -> None:
        if self.stop:
            return

        try:
            entire_oc: pandas.DataFrame
            current_time: str
            self.points: float
            entire_oc, current_time, self.points = self.get_dataframe()
        except TypeError:
            self.root.after((self.seconds * 1000), self.main)
            return

        self.str_current_time: str = current_time.split(" ")[1]
        current_date: datetime.date = datetime.datetime.strptime(current_time.split(" ")[0], '%d-%b-%Y').date()
        current_time: datetime.time = datetime.datetime.strptime(current_time.split(" ")[1], '%H:%M:%S').time()
        if self.first_run:
            self.previous_date = current_date
            self.previous_time = current_time
        elif current_date > self.previous_date:
            self.previous_date = current_date
            self.previous_time = current_time
        elif current_date == self.previous_date:
            if current_time > self.previous_time:
                time_difference: float = 0
                if current_time.hour > self.previous_time.hour:
                    time_difference = (60 - self.previous_time.minute) + current_time.minute + \
                                      ((60 - self.previous_time.second) + current_time.second) / 60
                elif current_time.hour == self.previous_time.hour:
                    time_difference = current_time.minute - self.previous_time.minute + \
                                      (current_time.second - self.previous_time.second) / 60
                if time_difference >= self.time_difference_factor and self.warn_late_update:
                    self.root.after(2000,
                                    (lambda title="Late Update", message=f"The data from the server was last updated "
                                                                         f"about {int(time_difference)} minutes ago.":
                                     messagebox.showinfo(title=title, message=message)))
                self.previous_time = current_time
            else:
                self.root.after((self.seconds * 1000), self.main)
                return

        call_oi_list: List[int] = []
        for i in range(len(entire_oc)):
            int_call_oi: int = int(entire_oc.iloc[i, [0]][0])
            call_oi_list.append(int_call_oi)
        call_oi_index: int = call_oi_list.index(max(call_oi_list))
        self.max_call_oi: float = round(max(call_oi_list) / self.round_factor, 1)
        self.max_call_oi_sp: float = float(entire_oc.iloc[call_oi_index]['Strike Price'])

        put_oi_list: List[int] = []
        for i in range(len(entire_oc)):
            int_put_oi: int = int(entire_oc.iloc[i, [20]][0])
            put_oi_list.append(int_put_oi)
        put_oi_index: int = put_oi_list.index(max(put_oi_list))
        self.max_put_oi: float = round(max(put_oi_list) / self.round_factor, 1)
        self.max_put_oi_sp: float = float(entire_oc.iloc[put_oi_index]['Strike Price'])

        sp_range_list: List[float] = []
        for i in range(put_oi_index, call_oi_index + 1):
            sp_range_list.append(float(entire_oc.iloc[i]['Strike Price']))

        self.max_call_oi_2: float
        self.max_call_oi_sp_2: float
        self.max_put_oi_2: float
        self.max_put_oi_sp_2: float
        if self.max_call_oi_sp == self.max_put_oi_sp:
            self.max_call_oi_2 = self.max_call_oi
            self.max_call_oi_sp_2 = self.max_call_oi_sp
            self.max_put_oi_2 = self.max_put_oi
            self.max_put_oi_sp_2 = self.max_put_oi_sp
        elif len(sp_range_list) == 2:
            self.max_call_oi_2 = round((entire_oc[entire_oc['Strike Price'] == self.max_put_oi_sp].iloc[0, 0]) /
                                       self.round_factor, 1)
            self.max_call_oi_sp_2 = self.max_put_oi_sp
            self.max_put_oi_2 = round((entire_oc[entire_oc['Strike Price'] == self.max_call_oi_sp].iloc[0, 20]) /
                                      self.round_factor, 1)
            self.max_put_oi_sp_2 = self.max_call_oi_sp
        else:
            call_oi_list_2: List[int] = []
            for i in range(put_oi_index, call_oi_index):
                int_call_oi_2: int = int(entire_oc.iloc[i, [0]][0])
                call_oi_list_2.append(int_call_oi_2)
            call_oi_index_2: int = put_oi_index + call_oi_list_2.index(max(call_oi_list_2))
            self.max_call_oi_2 = round(max(call_oi_list_2) / self.round_factor, 1)
            self.max_call_oi_sp_2 = float(entire_oc.iloc[call_oi_index_2]['Strike Price'])

            put_oi_list_2: List[int] = []
            for i in range(put_oi_index + 1, call_oi_index + 1):
                int_put_oi_2: int = int(entire_oc.iloc[i, [20]][0])
                put_oi_list_2.append(int_put_oi_2)
            put_oi_index_2: int = put_oi_index + 1 + put_oi_list_2.index(max(put_oi_list_2))
            self.max_put_oi_2 = round(max(put_oi_list_2) / self.round_factor, 1)
            self.max_put_oi_sp_2 = float(entire_oc.iloc[put_oi_index_2]['Strike Price'])

        total_call_oi: int = sum(call_oi_list)
        total_put_oi: int = sum(put_oi_list)
        self.put_call_ratio: float
        try:
            self.put_call_ratio = round(total_put_oi / total_call_oi, 2)
        except ZeroDivisionError:
            self.put_call_ratio = 0

        try:
            index: int = int(entire_oc[entire_oc['Strike Price'] == self.sp].index.tolist()[0])
        except IndexError as err:
            print(err, sys.exc_info()[0], "10")
            messagebox.showerror(title="Error",
                                 message="Incorrect Strike Price.\nPlease enter correct Strike Price.")
            self.root.destroy()
            return

        a: pandas.DataFrame = entire_oc[['Change in Open Interest']][entire_oc['Strike Price'] == self.sp]
        b1: pandas.Series = a.iloc[:, 0]
        c1: int = int(b1.get(index))
        b2: pandas.Series = entire_oc.iloc[:, 1]
        c2: int = int(b2.get((index + 1), 'Change in Open Interest'))
        b3: pandas.Series = entire_oc.iloc[:, 1]
        c3: int = int(b3.get((index + 2), 'Change in Open Interest'))
        if isinstance(c2, str):
            c2 = 0
        if isinstance(c3, str):
            c3 = 0
        self.call_sum: float = round((c1 + c2 + c3) / self.round_factor, 1)
        if self.call_sum == -0:
            self.call_sum = 0.0
        self.call_boundary: float = round(c3 / self.round_factor, 1)

        o1: pandas.Series = a.iloc[:, 1]
        p1: int = int(o1.get(index))
        o2: pandas.Series = entire_oc.iloc[:, 19]
        p2: int = int(o2.get((index + 1), 'Change in Open Interest'))
        p3: int = int(o2.get((index + 2), 'Change in Open Interest'))
        self.p4: int = int(o2.get((index + 4), 'Change in Open Interest'))
        o3: pandas.Series = entire_oc.iloc[:, 1]
        self.p5: int = int(o3.get((index + 4), 'Change in Open Interest'))
        self.p6: int = int(o3.get((index - 2), 'Change in Open Interest'))
        self.p7: int = int(o2.get((index - 2), 'Change in Open Interest'))
        if isinstance(p2, str):
            p2 = 0
        if isinstance(p3, str):
            p3 = 0
        if isinstance(self.p4, str):
            self.p4 = 0
        if isinstance(self.p5, str):
            self.p5 = 0
        self.put_sum: float = round((p1 + p2 + p3) / self.round_factor, 1)
        self.put_boundary: float = round(p1 / self.round_factor, 1)
        self.difference: float = round(self.call_sum - self.put_sum, 1)
        self.call_itm: float
        if self.p5 == 0:
            self.call_itm = 0.0
        else:
            self.call_itm = round(self.p4 / self.p5, 1)
            if self.call_itm == -0:
                self.call_itm = 0.0
        if isinstance(self.p6, str):
            self.p6 = 0
        if isinstance(self.p7, str):
            self.p7 = 0
        self.put_itm: float
        if self.p7 == 0:
            self.put_itm = 0.0
        else:
            self.put_itm = round(self.p6 / self.p7, 1)
            if self.put_itm == -0:
                self.put_itm = 0.0

        if self.stop:
            return

        self.set_values()

        if self.save_oc:
            try:
                entire_oc.to_csv(
                    f"NSE-OCA-{self.index if self.option_mode == 'Index' else self.stock}-{self.expiry_date}-Full.csv",
                    index=False)
            except PermissionError as err:
                print(err, sys.exc_info()[0], "11")
                messagebox.showerror(title="Export Failed",
                                     message=f"Failed to access NSE-OCA-"
                                             f"{self.index if self.option_mode == 'Index' else self.stock}-"
                                             f"{self.expiry_date}-Full.csv.\n"
                                             f"Permission Denied. Try closing any apps using it.")
            except Exception as err:
                print(err, sys.exc_info()[0], "16")

        if self.first_run:
            if self.update:
                self.check_for_updates()
            self.first_run = False
        if self.str_current_time == '15:30:00' and not self.stop and self.auto_stop \
                and self.previous_date == datetime.datetime.strptime(time.strftime("%d-%b-%Y", time.localtime()),
                                                                     "%d-%b-%Y").date():
            self.stop = True
            self.options.entryconfig(self.options.index(0), label="Start")
            messagebox.showinfo(title="Market Closed", message="Retrieving new data has been stopped.")
            return
        self.root.after((self.seconds * 1000), self.main)
        return

    @staticmethod
    def create_instance() -> None:
        master_window: Tk = Tk()
        Nse(master_window)
        master_window.mainloop()


if __name__ == '__main__':
    Nse.create_instance()
