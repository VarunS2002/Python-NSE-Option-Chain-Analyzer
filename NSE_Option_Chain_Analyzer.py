from typing import Union, Optional, List, Dict, Tuple, TextIO, Any
import sys
import os
import datetime
import webbrowser
import csv
import configparser
from tkinter import Tk, Toplevel, Event, TclError, StringVar, Frame, Menu, \
    Label, Entry, SOLID, RIDGE, N, S, E, W, LEFT, messagebox
from tkinter.ttk import Combobox, Button
import tksheet
import numpy
import pandas
import requests
import win10toast
import streamtologger


# noinspection PyAttributeOutsideInit
class Nse:
    def __init__(self, window: Tk) -> None:
        self.seconds: int = 60
        self.intervals: List[int] = [1, 2, 3, 5, 10, 15]
        self.stdout: TextIO = sys.stdout
        self.stderr: TextIO = sys.stderr
        self.previous_date: Optional[datetime.date] = None
        self.previous_time: Optional[datetime.time] = None
        self.first_run: bool = True
        self.stop: bool = False
        self.config_parser: configparser.ConfigParser = configparser.ConfigParser()
        self.config_parser.read('NSE-Option-Chain-Analyzer.ini')
        if not os.path.isfile('NSE-Option-Chain-Analyzer.ini'):
            self.config_parser.add_section('main')
            self.config_parser.set('main', 'notifications', 'False')
            self.config_parser.set('main', 'auto_stop', 'False')
            self.config_parser.set('main', 'logging', 'False')
            with open('NSE-Option-Chain-Analyzer.ini', 'w') as f:
                self.config_parser.write(f)
        self.notifications: bool = self.config_parser.getboolean('main', 'notifications')
        self.auto_stop: bool = self.config_parser.getboolean('main', 'auto_stop')
        self.logging: bool = not self.config_parser.getboolean('main', 'logging')
        if not self.logging:
            self.log()
        else:
            self.logging = False
        self.dates: List[str] = [""]
        self.indices: List[str] = ["NIFTY", "BANKNIFTY", "NIFTYIT"]
        self.headers: Dict[str, str] = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, '
                          'like Gecko) '
                          'Chrome/80.0.3987.149 Safari/537.36',
            'accept-language': 'en,gu;q=0.9,hi;q=0.8', 'accept-encoding': 'gzip, deflate, br'}
        self.url_oc: str = "https://www.nseindia.com/option-chain"
        self.session: requests.Session = requests.Session()
        self.cookies: Dict[str, str] = {}
        self.toaster: win10toast.ToastNotifier = win10toast.ToastNotifier()
        self.icon_path: str = Nse.get_icon_path() if os.path.isfile(Nse.get_icon_path()) else ''
        self.login_win(window)

    @staticmethod
    def get_icon_path() -> str:
        try:
            # noinspection PyProtectedMember,PyUnresolvedReferences
            base_path = sys._MEIPASS
        except AttributeError as err:
            print(err, "0")
            base_path = os.path.abspath(".")
        return os.path.join(base_path, 'nse_logo.ico')

    # noinspection PyUnusedLocal
    def get_data(self, event: Optional[Event] = None) -> Optional[Tuple[Optional[requests.Response], Any]]:
        if self.first_run:
            return self.get_data_first_run()
        else:
            return self.get_data_refresh()

    def get_data_first_run(self) -> Optional[Tuple[Optional[requests.Response], Any]]:
        request: Optional[requests.Response] = None
        response: Optional[requests.Response] = None
        self.index: str = self.index_var.get()
        try:
            url: str = f"https://www.nseindia.com/api/option-chain-indices?symbol={self.index}"
            request = self.session.get(self.url_oc, headers=self.headers, timeout=5)
            self.cookies = dict(request.cookies)
            response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
        except Exception as err:
            print(request)
            print(response)
            print(err, "1")
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
                print(err, "2")
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
                print(err, "3")
            return
        self.dates.clear()
        for dates in json_data['records']['expiryDates']:
            self.dates.append(dates)
        try:
            self.date_menu.config(values=tuple(self.dates))
            self.date_menu.current(0)
        except TclError as err:
            print(err, "4")

        return response, json_data

    def get_data_refresh(self) -> Optional[Tuple[Optional[requests.Response], Any]]:
        request: Optional[requests.Response] = None
        response: Optional[requests.Response] = None
        try:
            url: str = f"https://www.nseindia.com/api/option-chain-indices?symbol={self.index}"
            response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
            if response.status_code == 401:
                self.session.close()
                self.session = requests.Session()
                url = f"https://www.nseindia.com/api/option-chain-indices?symbol={self.index}"
                request = self.session.get(self.url_oc, headers=self.headers, timeout=5)
                self.cookies = dict(request.cookies)
                response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
                print("reset cookies")
        except Exception as err:
            print(request)
            print(response)
            print(err, "5")
            try:
                self.session.close()
                self.session = requests.Session()
                url = f"https://www.nseindia.com/api/option-chain-indices?symbol={self.index}"
                request = self.session.get(self.url_oc, headers=self.headers, timeout=5)
                self.cookies = dict(request.cookies)
                response = self.session.get(url, headers=self.headers, timeout=5, cookies=self.cookies)
                print("reset cookies")
            except Exception as err:
                print(request)
                print(response)
                print(err, "6")
                return
        if response is not None:
            try:
                json_data: Any = response.json()
            except Exception as err:
                print(response)
                print(err, "7")
                json_data = {}
        else:
            json_data = {}
        if json_data == {}:
            return

        return response, json_data

    def login_win(self, window: Tk) -> None:
        self.login: Tk = window
        self.login.title("NSE-Option-Chain-Analyzer")
        window_width: int = self.login.winfo_reqwidth()
        window_height: int = self.login.winfo_reqheight()
        position_right: int = int(self.login.winfo_screenwidth() / 2 - window_width / 2)
        position_down: int = int(self.login.winfo_screenheight() / 2 - window_height / 2)
        self.login.geometry("320x110+{}+{}".format(position_right, position_down))
        if self.icon_path:
            self.login.iconbitmap(self.icon_path)
        self.login.rowconfigure(0, weight=1)
        self.login.rowconfigure(1, weight=1)
        self.login.rowconfigure(2, weight=1)
        self.login.rowconfigure(3, weight=1)
        self.login.columnconfigure(0, weight=1)
        self.login.columnconfigure(1, weight=1)
        self.login.columnconfigure(2, weight=1)

        self.intervals_var: StringVar = StringVar()
        self.intervals_var.set(self.intervals[0])
        self.index_var: StringVar = StringVar()
        self.index_var.set(self.indices[0])
        self.dates_var: StringVar = StringVar()
        self.dates_var.set(self.dates[0])

        index_label: Label = Label(self.login, text="Index: ", justify=LEFT)
        index_label.grid(row=0, column=0, sticky=N + S + W)
        self.index_menu: Combobox = Combobox(self.login, textvariable=self.index_var, values=self.indices)
        self.index_menu.config(width=15)
        self.index_menu.grid(row=0, column=1, sticky=N + S + E)
        date_label: Label = Label(self.login, text="Expiry Date: ", justify=LEFT)
        date_label.grid(row=1, column=0, sticky=N + S + W)
        self.date_menu: Combobox = Combobox(self.login, textvariable=self.dates_var)
        self.date_menu.config(width=15)
        self.date_menu.grid(row=1, column=1, sticky=N + S + E)
        self.date_get: Button = Button(self.login, text="Refresh", command=self.get_data, width=10)
        self.date_get.grid(row=1, column=2, sticky=N + S + E + W)
        sp_label: Label = Label(self.login, text="Strike Price (eg. 11850): ")
        sp_label.grid(row=2, column=0, sticky=N + S + W)
        self.sp_entry = Entry(self.login, width=18, relief=SOLID)
        self.sp_entry.grid(row=2, column=1, sticky=N + S + E)
        start_btn: Button = Button(self.login, text="Start", command=self.start, width=10)
        start_btn.grid(row=2, column=2, rowspan=2, sticky=N + S + E + W)
        intervals_label: Label = Label(self.login, text="Refresh Interval (in min): ", justify=LEFT)
        intervals_label.grid(row=3, column=0, sticky=N + S + W)
        self.intervals_menu: Combobox = Combobox(self.login, textvariable=self.intervals_var,
                                                 values=tuple(self.intervals))
        self.intervals_menu.config(width=15)
        self.intervals_menu.grid(row=3, column=1, sticky=N + S + E)
        self.intervals_menu.current(0)
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
        self.date_menu.bind('<Return>', lambda event, a=2: focus_widget(event, a))
        self.sp_entry.bind('<Return>', self.start)

        self.login.mainloop()

    # noinspection PyUnusedLocal
    def start(self, event: Optional[Event] = None) -> None:
        self.seconds = int(self.intervals_var.get()) * 60
        self.expiry_date: str = self.dates_var.get()
        if self.expiry_date == "":
            messagebox.showerror(title="Error", message="Incorrect Expiry Date.\nPlease enter correct Expiry Date.")
            return
        try:
            self.sp: int = int(self.sp_entry.get())
            self.login.destroy()
            self.main_win()
        except ValueError as err:
            print(err, "8")
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
        csv_exists: bool = os.path.isfile(f"NSE-OCA-{self.index}-{self.expiry_date}.csv")
        try:
            if not csv_exists:
                with open(f"NSE-OCA-{self.index}-{self.expiry_date}.csv", "a", newline="") as row:
                    data_writer: csv.writer = csv.writer(row)
                    data_writer.writerow((
                        'Time', 'Value', 'Call Sum (in K)', 'Put Sum (in K)', 'Difference (in K)',
                        'Call Boundary (in K)', 'Put Boundary (in K)', 'Call ITM', 'Put ITM'))

            with open(f"NSE-OCA-{self.index}-{self.expiry_date}.csv", "a", newline="") as row:
                data_writer: csv.writer = csv.writer(row)
                data_writer.writerows(sheet_data)

            messagebox.showinfo(title="Export Successful",
                                message=f"Data has been exported to NSE-OCA-{self.index}-{self.expiry_date}.csv.")
        except Exception as err:
            print(err, "9")
            messagebox.showerror(title="Export Failed",
                                 message="An error occurred while exporting the data.")

    # noinspection PyUnusedLocal
    def toggle_notifications(self, event: Optional[Event] = None) -> None:
        if self.notifications:
            self.notifications = False
            self.config_parser.set('main', 'notifications', 'False')
            self.options.entryconfig(self.options.index(2), label="Notifications: Off")
            messagebox.showinfo(title="Notifications Disabled",
                                message="You will not receive any Notifications.")
        else:
            self.notifications = True
            self.config_parser.set('main', 'notifications', 'True')
            self.options.entryconfig(self.options.index(2), label="Notifications: On")
            messagebox.showinfo(title="Notifications Enabled",
                                message="You will receive Notifications when the state of a label changes.")

        with open('NSE-Option-Chain-Analyzer.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def toggle_auto_stop(self, event: Optional[Event] = None) -> None:
        if self.auto_stop:
            self.auto_stop = False
            self.config_parser.set('main', 'auto_stop', 'False')
            self.options.entryconfig(self.options.index(3), label="Stop automatically at 3:30pm: Off")
            messagebox.showinfo(title="Auto Stop Disabled", message="Program will not automatically stop at 3:30pm")
        else:
            self.auto_stop = True
            self.config_parser.set('main', 'auto_stop', 'True')
            self.options.entryconfig(self.options.index(3), label="Stop automatically at 3:30pm: On")
            messagebox.showinfo(title="Auto Stop Enabled", message="Program will automatically stop at 3:30pm")

        with open('NSE-Option-Chain-Analyzer.ini', 'w') as f:
            self.config_parser.write(f)

    # noinspection PyUnusedLocal
    def log(self, event: Optional[Event] = None) -> None:
        if not self.logging:
            streamtologger.redirect(target="NSE-Option-Chain-Analyzer.log",
                                    header_format="[{timestamp:%Y-%m-%d %H:%M:%S} - {level:5}] ")
            self.logging = True
            print('----------Logging Started----------')
            self.config_parser.set('main', 'logging', 'True')
            try:
                self.options.entryconfig(self.options.index(5), label="Logging: On")
                messagebox.showinfo(title="Debug Logging Enabled",
                                    message="Errors will be logged to NSE-Option-Chain-Analyzer.log.")
            except AttributeError as err:
                print(err, "10")
        elif self.logging:
            sys.stdout = self.stdout
            sys.stderr = self.stderr
            streamtologger._is_redirected = False
            self.logging = False
            print('----------Logging Stopped----------')
            self.config_parser.set('main', 'logging', 'False')
            self.options.entryconfig(self.options.index(5), label="Logging: Off")
            messagebox.showinfo(title="Debug Logging Disabled", message="Errors will not be logged.")

        with open('NSE-Option-Chain-Analyzer.ini', 'w') as f:
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
        if self.icon_path:
            self.info.iconbitmap(self.icon_path)
        self.info.attributes('-topmost', True)
        self.info.grab_set()
        self.info.focus_force()

        return self.info

    def about(self) -> None:
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
        version_val: Label = Label(self.info, text="3.7", relief=RIDGE)
        version_val.grid(row=1, column=1, sticky=N + S + W + E)
        dev_label: Label = Label(self.info, text="Developer:", relief=RIDGE)
        dev_label.grid(row=2, column=0, sticky=N + S + W + E)
        dev_val: Label = Label(self.info, text="Varun Shanbhag", fg="blue", cursor="hand2", relief=RIDGE)
        dev_val.bind("<Button-1>", lambda click, link="developer": self.links(link, click))
        dev_val.grid(row=2, column=1, sticky=N + S + W + E)
        readme: Label = Label(self.info, text="README", fg="blue", cursor="hand2", relief=RIDGE)
        readme.bind("<Button-1>", lambda click, link="readme": self.links(link, click))
        readme.grid(row=3, column=0, sticky=N + S + W + E)
        licenses: Label = Label(self.info, text="License", fg="blue", cursor="hand2", relief=RIDGE)
        licenses.bind("<Button-1>", lambda click, link="license": self.links(link, click))
        licenses.grid(row=3, column=1, sticky=N + S + W + E)
        releases: Label = Label(self.info, text="Releases", fg="blue", cursor="hand2", relief=RIDGE)
        releases.bind("<Button-1>", lambda click, link="releases": self.links(link, click))
        releases.grid(row=4, column=0, sticky=N + S + W + E)
        sources: Label = Label(self.info, text="Sources", fg="blue", cursor="hand2", relief=RIDGE)
        sources.bind("<Button-1>", lambda click, link="sources": self.links(link, click))
        sources.grid(row=4, column=1, sticky=N + S + W + E)

        self.info.mainloop()

    # noinspection PyUnusedLocal
    def close(self, event: Optional[Event] = None) -> None:
        ask_quit: bool = messagebox.askyesno("Quit", "All unsaved data will be lost.\nProceed to quit?", icon='warning',
                                             default='no')
        if ask_quit:
            self.session.close()
            print('----------Quitting Program----------')
            self.root.destroy()
            sys.exit()
        elif not ask_quit:
            pass

    def main_win(self) -> None:
        self.root: Tk = Tk()
        self.root.focus_force()
        self.root.title("NSE-Option-Chain-Analyzer")
        self.root.protocol('WM_DELETE_WINDOW', self.close)
        window_width: int = self.root.winfo_reqwidth()
        window_height: int = self.root.winfo_reqheight()
        position_right: int = int(self.root.winfo_screenwidth() / 3 - window_width / 2)
        position_down: int = int(self.root.winfo_screenheight() / 3 - window_height / 2)
        self.root.geometry("815x560+{}+{}".format(position_right, position_down))
        if self.icon_path:
            self.root.iconbitmap(self.icon_path)
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        menubar: Menu = Menu(self.root)
        self.options: Menu = Menu(menubar, tearoff=0)
        self.options.add_command(label="Stop", accelerator="(Ctrl+X)", command=self.change_state)
        self.options.add_command(label="Export to CSV", accelerator="(Ctrl+S)", command=self.export)
        self.options.add_command(label=f"Notifications: {'On' if self.notifications else 'Off'}",
                                 accelerator="(Ctrl+N)", command=self.toggle_notifications)
        self.options.add_command(label=f"Stop automatically at 3:30pm: {'On' if self.auto_stop else 'Off'}",
                                 accelerator="(Ctrl+K)", command=self.toggle_auto_stop)
        self.options.add_separator()
        self.options.add_command(label=f"Debug Logging: {'On' if self.logging else 'Off'}", accelerator="(Ctrl+L)",
                                 command=self.log)
        self.options.add_command(label="About", accelerator="(Ctrl+M)", command=self.about)
        self.options.add_command(label="Quit", accelerator="(Ctrl+Q)", command=self.close)
        menubar.add_cascade(label="Menu", menu=self.options)
        self.root.config(menu=menubar)

        self.root.bind('<Control-x>', self.change_state)
        self.root.bind('<Control-s>', self.export)
        self.root.bind('<Control-n>', self.toggle_notifications)
        self.root.bind('<Control-k>', self.toggle_auto_stop)
        self.root.bind('<Control-l>', self.log)
        self.root.bind('<Control-m>', self.about)
        self.root.bind('<Control-q>', self.close)

        top_frame: Frame = Frame(self.root)
        top_frame.rowconfigure(0, weight=1)
        top_frame.columnconfigure(0, weight=1)
        top_frame.pack(fill="both", expand=True)

        output_columns: Tuple[str, str, str, str, str, str, str, str, str] = (
            'Time', 'Value', 'Call Sum\n(in K)', 'Put Sum\n(in K)', 'Difference\n(in K)', 'Call Boundary\n(in K)',
            'Put Boundary\n(in K)', 'Call ITM', 'Put ITM')
        self.sheet: tksheet.Sheet = tksheet.Sheet(top_frame, column_width=85, align="center", headers=output_columns,
                                                  header_font=("TkDefaultFont", 9, "bold"), empty_horizontal=0,
                                                  empty_vertical=20, header_height=35)
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
        max_call_oi_label: Label = Label(bottom_frame, text="OI (in K):", relief=RIDGE,
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
        max_put_oi_label: Label = Label(bottom_frame, text="OI (in K):", relief=RIDGE,
                                        font=("TkDefaultFont", 9, "bold"))
        max_put_oi_label.grid(row=1, column=6, sticky=N + S + W + E)
        self.max_put_oi_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_put_oi_val.grid(row=1, column=7, sticky=N + S + W + E)
        max_call_oi_sp_2_label: Label = Label(bottom_frame, text="Strike Price 2:", relief=RIDGE,
                                              font=("TkDefaultFont", 9, "bold"))
        max_call_oi_sp_2_label.grid(row=2, column=0, sticky=N + S + W + E)
        self.max_call_oi_sp_2_val: Label = Label(bottom_frame, text="", relief=RIDGE)
        self.max_call_oi_sp_2_val.grid(row=2, column=1, sticky=N + S + W + E)
        max_call_oi_2_label: Label = Label(bottom_frame, text="OI (in K):", relief=RIDGE,
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
        max_put_oi_2_label: Label = Label(bottom_frame, text="OI (in K):", relief=RIDGE,
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

    def get_dataframe(self) -> Optional[Tuple[pandas.DataFrame, str, str, float]]:
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
                                 "CE" in data and str(data['expiryDate'].lower() == str(self.expiry_date).lower())]
        pe_values: List[dict] = [data['PE'] for data in json_data['records']['data'] if
                                 "PE" in data and str(data['expiryDate'].lower() == str(self.expiry_date).lower())]
        underlying_stock: str = ce_values[0]['underlying']
        points: float = pe_values[0]['underlyingValue']
        ce_data: pandas.DataFrame = pandas.DataFrame(ce_values)
        pe_data: pandas.DataFrame = pandas.DataFrame(pe_values)
        ce_data_f: pandas.DataFrame = ce_data.loc[ce_data['expiryDate'] == self.expiry_date]
        pe_data_f: pandas.DataFrame = pe_data.loc[pe_data['expiryDate'] == self.expiry_date]
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
        return merged_inner, current_time, underlying_stock, points

    def set_values(self) -> None:
        self.root.title(f"NSE-Option-Chain-Analyzer - {self.underlying_stock} - {self.expiry_date} - {self.sp}")

        self.max_call_oi_val.config(text=self.max_call_oi)
        self.max_call_oi_sp_val.config(text=self.max_call_oi_sp)
        self.max_call_oi_2_val.config(text=self.max_call_oi_2)
        self.max_call_oi_sp_2_val.config(text=self.max_call_oi_sp_2)
        self.max_put_oi_val.config(text=self.max_put_oi)
        self.max_put_oi_sp_val.config(text=self.max_put_oi_sp)
        self.max_put_oi_2_val.config(text=self.max_put_oi_2)
        self.max_put_oi_sp_2_val.config(text=self.max_put_oi_sp_2)

        red: str = "#e53935"
        green: str = "#00e676"
        default: str = "SystemButtonFace"

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
                self.toaster.show_toast("Open Interest changed", f"Changed from {self.old_oi_label} to {oi_label}",
                                        duration=4, threaded=True, icon_path=self.icon_path)
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
                self.toaster.show_toast("Call ITM changed", f"Changed from {self.old_call_label} to {call}",
                                        duration=4, threaded=True, icon_path=self.icon_path)
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
                self.toaster.show_toast("Put ITM changed", f"Changed from {self.old_put_label} to {put}",
                                        duration=4, threaded=True, icon_path=self.icon_path)
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
                self.toaster.show_toast("Call Exits changed",
                                        f"Changed from {self.old_call_exits_label} to {call_exits_label}",
                                        duration=4, threaded=True, icon_path=self.icon_path)
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
                self.toaster.show_toast("Put Exits changed",
                                        f"Changed from {self.old_put_exits_label} to {put_exits_label}",
                                        duration=4, threaded=True, icon_path=self.icon_path)
            self.old_put_exits_label = put_exits_label

        output_values: List[Union[str, float, numpy.float64]] = [self.str_current_time, self.points, self.call_sum,
                                                                 self.put_sum, self.difference,
                                                                 self.call_boundary, self.put_boundary, self.call_itm,
                                                                 self.put_itm]
        self.sheet.insert_row(values=output_values)

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
        self.old_call_sum: numpy.float64
        if self.first_run or self.old_call_sum == self.call_sum:
            self.old_call_sum = self.call_sum
        elif self.call_sum > self.old_call_sum:
            self.sheet.highlight_cells(row=last_row, column=2, bg=red)
            self.old_call_sum = self.call_sum
        else:
            self.sheet.highlight_cells(row=last_row, column=2, bg=green)
            self.old_call_sum = self.call_sum
        self.old_put_sum: numpy.float64
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
        self.old_call_boundary: numpy.float64
        if self.first_run or self.old_call_boundary == self.call_boundary:
            self.old_call_boundary = self.call_boundary
        elif self.call_boundary > self.old_call_boundary:
            self.sheet.highlight_cells(row=last_row, column=5, bg=red)
            self.old_call_boundary = self.call_boundary
        else:
            self.sheet.highlight_cells(row=last_row, column=5, bg=green)
            self.old_call_boundary = self.call_boundary
        self.old_put_boundary: numpy.float64
        if self.first_run or self.old_put_boundary == self.put_boundary:
            self.old_put_boundary = self.put_boundary
        elif self.put_boundary > self.old_put_boundary:
            self.sheet.highlight_cells(row=last_row, column=6, bg=green)
            self.old_put_boundary = self.put_boundary
        else:
            self.sheet.highlight_cells(row=last_row, column=6, bg=red)
            self.old_put_boundary = self.put_boundary
        self.old_call_itm: numpy.float64
        if self.first_run or self.old_call_itm == self.call_itm:
            self.old_call_itm = self.call_itm
        elif self.call_itm > self.old_call_itm:
            self.sheet.highlight_cells(row=last_row, column=7, bg=green)
            self.old_call_itm = self.call_itm
        else:
            self.sheet.highlight_cells(row=last_row, column=7, bg=red)
            self.old_call_itm = self.call_itm
        self.old_put_itm: numpy.float64
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
            df: pandas.DataFrame
            current_time: str
            self.underlying_stock: str
            self.points: float
            df, current_time, self.underlying_stock, self.points = self.get_dataframe()
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
                self.previous_time = current_time
            else:
                self.root.after((self.seconds * 1000), self.main)
                return

        call_oi_list: List[int] = []
        for i in range(len(df)):
            int_call_oi: int = int(df.iloc[i, [0]][0])
            call_oi_list.append(int_call_oi)
        call_oi_index: int = call_oi_list.index(max(call_oi_list))
        self.max_call_oi: float = round(max(call_oi_list) / 1000, 1)
        self.max_call_oi_sp: numpy.float64 = df.iloc[call_oi_index]['Strike Price']

        put_oi_list: List[int] = []
        for i in range(len(df)):
            int_put_oi: int = int(df.iloc[i, [20]][0])
            put_oi_list.append(int_put_oi)
        put_oi_index: int = put_oi_list.index(max(put_oi_list))
        self.max_put_oi: float = round(max(put_oi_list) / 1000, 1)
        self.max_put_oi_sp: numpy.float64 = df.iloc[put_oi_index]['Strike Price']

        sp_range_list: List[numpy.float64] = []
        for i in range(put_oi_index, call_oi_index + 1):
            sp_range_list.append(df.iloc[i]['Strike Price'])

        self.max_call_oi_2: float
        self.max_call_oi_sp_2: numpy.float64
        self.max_put_oi_2: float
        self.max_put_oi_sp_2: numpy.float64
        if self.max_call_oi_sp == self.max_put_oi_sp:
            self.max_call_oi_2 = self.max_call_oi
            self.max_call_oi_sp_2 = self.max_call_oi_sp
            self.max_put_oi_2 = self.max_put_oi
            self.max_put_oi_sp_2 = self.max_put_oi_sp
        elif len(sp_range_list) == 2:
            self.max_call_oi_2 = df[df['Strike Price'] == self.max_put_oi_sp].iloc[0, 0]
            self.max_call_oi_sp_2 = self.max_put_oi_sp
            self.max_put_oi_2 = df[df['Strike Price'] == self.max_call_oi_sp].iloc[0, 20]
            self.max_put_oi_sp_2 = self.max_call_oi_sp
        else:
            call_oi_list_2: List[int] = []
            for i in range(put_oi_index, call_oi_index):
                int_call_oi_2: int = int(df.iloc[i, [0]][0])
                call_oi_list_2.append(int_call_oi_2)
            call_oi_index_2: int = put_oi_index + call_oi_list_2.index(max(call_oi_list_2))
            self.max_call_oi_2 = round(max(call_oi_list_2) / 1000, 1)
            self.max_call_oi_sp_2 = df.iloc[call_oi_index_2]['Strike Price']

            put_oi_list_2: List[int] = []
            for i in range(put_oi_index + 1, call_oi_index + 1):
                int_put_oi_2: int = int(df.iloc[i, [20]][0])
                put_oi_list_2.append(int_put_oi_2)
            put_oi_index_2: int = put_oi_index + 1 + put_oi_list_2.index(max(put_oi_list_2))
            self.max_put_oi_2 = round(max(put_oi_list_2) / 1000, 1)
            self.max_put_oi_sp_2 = df.iloc[put_oi_index_2]['Strike Price']

        total_call_oi: int = sum(call_oi_list)
        total_put_oi: int = sum(put_oi_list)
        self.put_call_ratio: float
        try:
            self.put_call_ratio = round(total_put_oi / total_call_oi, 2)
        except ZeroDivisionError:
            self.put_call_ratio = 0

        try:
            index: int = int(df[df['Strike Price'] == self.sp].index.tolist()[0])
        except IndexError as err:
            print(err, "11")
            messagebox.showerror(title="Error",
                                 message="Incorrect Strike Price.\nPlease enter correct Strike Price.")
            self.root.destroy()
            return

        a: pandas.DataFrame = df[['Change in Open Interest']][df['Strike Price'] == self.sp]
        b1: pandas.Series = a.iloc[:, 0]
        c1: numpy.int64 = b1.get(index)
        b2: pandas.Series = df.iloc[:, 1]
        c2: numpy.int64 = b2.get((index + 1), 'Change in Open Interest')
        b3: pandas.Series = df.iloc[:, 1]
        c3: numpy.int64 = b3.get((index + 2), 'Change in Open Interest')
        if isinstance(c2, str):
            c2 = 0
        if isinstance(c3, str):
            c3 = 0
        self.call_sum: numpy.float64 = round((c1 + c2 + c3) / 1000, 1)
        if self.call_sum == -0:
            self.call_sum = 0.0
        self.call_boundary: numpy.float64 = round(c3 / 1000, 1)

        o1: pandas.Series = a.iloc[:, 1]
        p1: numpy.int64 = o1.get(index)
        o2: pandas.Series = df.iloc[:, 19]
        p2: numpy.int64 = o2.get((index + 1), 'Change in Open Interest')
        p3: numpy.int64 = o2.get((index + 2), 'Change in Open Interest')
        self.p4: numpy.int64 = o2.get((index + 4), 'Change in Open Interest')
        o3: pandas.Series = df.iloc[:, 1]
        self.p5: numpy.int64 = o3.get((index + 4), 'Change in Open Interest')
        self.p6: numpy.int64 = o3.get((index - 2), 'Change in Open Interest')
        self.p7: numpy.int64 = o2.get((index - 2), 'Change in Open Interest')
        if isinstance(p2, str):
            p2 = 0
        if isinstance(p3, str):
            p3 = 0
        if isinstance(self.p4, str):
            self.p4 = 0
        if isinstance(self.p5, str):
            self.p5 = 0
        self.put_sum: numpy.float64 = round((p1 + p2 + p3) / 1000, 1)
        self.put_boundary: numpy.float64 = round(p1 / 1000, 1)
        self.difference: float = float(round(self.call_sum - self.put_sum, 1))
        self.call_itm: numpy.float64
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
        self.put_itm: numpy.float64
        if self.p7 == 0:
            self.put_itm = 0.0
        else:
            self.put_itm = round(self.p6 / self.p7, 1)
            if self.put_itm == -0:
                self.put_itm = 0.0

        if self.stop:
            return

        self.set_values()

        if self.first_run:
            self.first_run = False
        if self.str_current_time == '15:30:00' and not self.stop and self.auto_stop:
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
