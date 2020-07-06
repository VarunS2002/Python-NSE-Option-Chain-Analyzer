from tkinter import *
from tkinter import messagebox
import tksheet
import bs4
import requests
import pandas
import datetime
import webbrowser
import csv

seconds = 30

headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/80.0.3987.149 Safari/537.36',
           'accept-language': 'en,gu;q=0.9,hi;q=0.8', 'accept-encoding': 'gzip, deflate, br'}

previous_time = None
old_points = None
old_call_sum = None
old_put_sum = None
old_difference = None
old_call_boundary = None
old_put_boundary = None
old_call_itm = None
old_put_itm = None
stop = 0


def login():
    global login
    global url_entry
    global sp_entry

    login = Tk()
    login.title("NSE-Option-Chain-Analyzer")
    # login.resizable(False, False)
    window_width = login.winfo_reqwidth()
    window_height = login.winfo_reqheight()
    position_right = int(login.winfo_screenwidth() / 3 - window_width / 2)
    position_down = int(login.winfo_screenheight() / 2 - window_height / 2)
    login.geometry("800x70+{}+{}".format(position_right, position_down))

    url_label = Label(login, text="URL: ", justify=LEFT)
    url_label.grid(row=0, column=0, sticky=N + S + W)
    url_entry = Entry(login, width=120)
    url_entry.grid(row=0, column=1, sticky=N + S + E)
    sp_label = Label(login, text="Strike Price: ")
    sp_label.grid(row=1, column=0, sticky=N + S + W)
    sp_entry = Entry(login, width=10)
    sp_entry.grid(row=1, column=1, sticky=N + S + W)
    start_btn = Button(login, text="Start", command=start)
    start_btn.grid(row=2, column=0, stick=N + S + E + W)
    url_entry.focus_set()

    def focus_widget(event, mode):
        if mode == 1:
            sp_entry.focus()
        elif mode == 2:

            start()

    url_entry.bind('<Return>', lambda event, a=1: focus_widget(event, a))
    sp_entry.bind('<Return>', lambda event, a=2: focus_widget(event, a))

    login.mainloop()


def start():
    global login
    global url
    global sp

    url = url_entry.get()
    sp = str(sp_entry.get())

    if "." in sp:
        pass
    else:
        sp = str(f"{sp}.00")
    login.destroy()

    main_win()


def change_state(event="empty"):
    global stop

    if stop == 0:
        stop = 1
        options.entryconfig(options.index(0), label="Start   (Ctrl+X)")
        messagebox.showinfo(title="Stopped", message="Scraping new data has been stopped.")
    else:
        stop = 0
        options.entryconfig(options.index(0), label="Stop   (Ctrl+X)")
        messagebox.showinfo(title="Started", message="Scraping new data has been started.")

        main()


def export(event="empty"):
    global sheet

    sheet_data = sheet.get_sheet_data()

    try:
        with open("NSE-Option-Chain-Analyzer.csv", "a", newline="") as row:
            data_writer = csv.writer(row)
            data_writer.writerows(sheet_data)

        messagebox.showinfo(title="Export Complete",
                            message="Data has been exported to NSE-Option-Chain-Analyzer.csv.")
    except Exception:
        messagebox.showerror(title="Export Failed",
                             message="An error occurred while exporting the data URL.")


def links(event, link):
    global info

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

    info.attributes('-topmost', False)


def about_window():
    global info

    info = Toplevel()
    info.title("About")
    # info.resizable(False, False)
    window_width = info.winfo_reqwidth()
    window_height = info.winfo_reqheight()
    position_right = int(info.winfo_screenwidth() / 2 - window_width / 2)
    position_down = int(info.winfo_screenheight() / 2 - window_height / 2)
    info.geometry("250x150+{}+{}".format(position_right, position_down))
    info.attributes('-topmost', True)
    info.grab_set()
    info.focus_force()

    return info


def about(event="empty"):
    info = about_window()
    info.rowconfigure(0, weight=1)
    info.rowconfigure(1, weight=1)
    info.rowconfigure(2, weight=1)
    info.rowconfigure(3, weight=1)
    info.rowconfigure(4, weight=1)
    info.columnconfigure(0, weight=1)
    info.columnconfigure(1, weight=1)

    heading = Label(info, text="NSE-Option-Chain-Analyzer", relief=RIDGE, font=("TkDefaultFont", 10, "bold"))
    heading.grid(row=0, column=0, columnspan=2, sticky=N + S + W + E)
    version_label = Label(info, text="Version:", relief=RIDGE)
    version_label.grid(row=1, column=0, sticky=N + S + W + E)
    version_val = Label(info, text="2.0", relief=RIDGE)
    version_val.grid(row=1, column=1, sticky=N + S + W + E)
    dev_label = Label(info, text="Developer:", relief=RIDGE)
    dev_label.grid(row=2, column=0, sticky=N + S + W + E)
    dev_val = Label(info, text="Varun Shanbhag", fg="blue", cursor="hand2", relief=RIDGE)
    dev_val.bind("<Button-1>", lambda event, link="developer": links(event, link))
    dev_val.grid(row=2, column=1, sticky=N + S + W + E)
    readme = Label(info, text="README", fg="blue", cursor="hand2", relief=RIDGE)
    readme.bind("<Button-1>", lambda event, link="readme": links(event, link))
    readme.grid(row=3, column=0, sticky=N + S + W + E)
    licenses = Label(info, text="License", fg="blue", cursor="hand2", relief=RIDGE)
    licenses.bind("<Button-1>", lambda event, link="license": links(event, link))
    licenses.grid(row=3, column=1, sticky=N + S + W + E)
    releases = Label(info, text="Releases", fg="blue", cursor="hand2", relief=RIDGE)
    releases.bind("<Button-1>", lambda event, link="releases": links(event, link))
    releases.grid(row=4, column=0, sticky=N + S + W + E)
    sources = Label(info, text="Sources", fg="blue", cursor="hand2", relief=RIDGE)
    sources.bind("<Button-1>", lambda event, link="sources": links(event, link))
    sources.grid(row=4, column=1, sticky=N + S + W + E)

    info.mainloop()


def close(event="empty"):
    ask_quit = messagebox.askyesno("Quit", "All unsaved data will be lost.\nProceed to quit?", icon='warning',
                                   default='no')
    if ask_quit:
        quit()
    elif not ask_quit:
        pass


def main_win():
    global root
    global options
    global max_call_oi_val
    global max_call_oi_sp_val
    global max_put_oi_sp_val
    global max_put_oi_val
    global sheet
    global oi_val
    global pcr_val
    global call_exits_val
    global put_exits_val
    global call_itm_val
    global put_itm_val

    root = Tk()
    root.focus_force()
    root.title("NSE-Option-Chain-Analyzer")
    root.protocol('WM_DELETE_WINDOW', close)
    # root.resizable(False, False)
    window_width = root.winfo_reqwidth()
    window_height = root.winfo_reqheight()
    position_right = int(root.winfo_screenwidth() / 3 - window_width / 2)
    position_down = int(root.winfo_screenheight() / 3 - window_height / 2)
    root.geometry("910x510+{}+{}".format(position_right, position_down))
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    menubar = Menu(root)
    options = Menu(menubar, tearoff=0)
    options.add_command(label="Stop   (Ctrl+X)", command=change_state)
    options.add_command(label="Export to CSV   (Ctrl+S)", command=export)
    options.add_separator()
    options.add_command(label="About   (Ctrl+M)", command=about)
    options.add_command(label="Quit   (Ctrl+Q)", command=close)
    menubar.add_cascade(label="Menu", menu=options)
    root.config(menu=menubar)

    root.bind('<Control-s>', export)
    root.bind('<Control-x>', change_state)
    root.bind('<Control-m>', about)
    root.bind('<Control-q>', close)

    top_frame = Frame(root)
    top_frame.rowconfigure(0, weight=1)
    top_frame.columnconfigure(0, weight=1)
    top_frame.pack(fill="both", expand=True)

    output_columns = (
        'Time', 'Points', 'Call Sum', 'Put Sum', 'Difference', 'Call Boundary', 'Put Boundary', 'Call ITM',
        'Put ITM')
    sheet = tksheet.Sheet(top_frame, column_width=95, align="center", headers=output_columns,
                          header_font=("TkDefaultFont", 9, "bold"), empty_horizontal=0, empty_vertical=20)
    sheet.enable_bindings(("toggle_select", "drag_select", "column_select", "row_select", "column_width_resize",
                           "arrowkeys", "right_click_popup_menu", "rc_select", "copy", "select_all"))
    sheet.grid(row=0, column=0, sticky=N + S + W + E)

    bottom_frame = Frame(root)
    bottom_frame.rowconfigure(0, weight=1)
    bottom_frame.rowconfigure(1, weight=1)
    bottom_frame.rowconfigure(2, weight=1)
    bottom_frame.rowconfigure(3, weight=1)
    bottom_frame.rowconfigure(4, weight=1)
    bottom_frame.columnconfigure(0, weight=1)
    bottom_frame.columnconfigure(1, weight=1)
    bottom_frame.columnconfigure(2, weight=1)
    bottom_frame.columnconfigure(3, weight=1)
    bottom_frame.columnconfigure(4, weight=1)
    bottom_frame.columnconfigure(5, weight=1)
    bottom_frame.columnconfigure(6, weight=1)
    bottom_frame.columnconfigure(7, weight=1)
    bottom_frame.pack(fill="both", expand=True)

    oi_ub_label = Label(bottom_frame, text="Open Interest Upper Boundary", relief=RIDGE,
                        font=("TkDefaultFont", 10, "bold"))
    oi_ub_label.grid(row=0, column=0, columnspan=4, sticky=N + S + W + E)
    max_call_oi_sp_label = Label(bottom_frame, text="Strike Price:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    max_call_oi_sp_label.grid(row=1, column=0, sticky=N + S + W + E)
    max_call_oi_sp_val = Label(bottom_frame, text="", relief=RIDGE)
    max_call_oi_sp_val.grid(row=1, column=1, sticky=N + S + W + E)
    max_call_oi_label = Label(bottom_frame, text="OI:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    max_call_oi_label.grid(row=1, column=2, sticky=N + S + W + E)
    max_call_oi_val = Label(bottom_frame, text="", relief=RIDGE)
    max_call_oi_val.grid(row=1, column=3, sticky=N + S + W + E)
    oi_lb_label = Label(bottom_frame, text="Open Interest Lower Boundary", relief=RIDGE,
                        font=("TkDefaultFont", 10, "bold"))
    oi_lb_label.grid(row=0, column=4, columnspan=4, sticky=N + S + W + E)
    max_put_oi_sp_label = Label(bottom_frame, text="Strike Price:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    max_put_oi_sp_label.grid(row=1, column=4, sticky=N + S + W + E)
    max_put_oi_sp_val = Label(bottom_frame, text="", relief=RIDGE)
    max_put_oi_sp_val.grid(row=1, column=5, sticky=N + S + W + E)
    max_put_oi_label = Label(bottom_frame, text="OI:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    max_put_oi_label.grid(row=1, column=6, sticky=N + S + W + E)
    max_put_oi_val = Label(bottom_frame, text="", relief=RIDGE)
    max_put_oi_val.grid(row=1, column=7, sticky=N + S + W + E)

    oi_label = Label(bottom_frame, text="Open Interest:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    oi_label.grid(row=2, column=0, columnspan=2, sticky=N + S + W + E)
    oi_val = Label(bottom_frame, text="", relief=RIDGE)
    oi_val.grid(row=2, column=2, columnspan=2, sticky=N + S + W + E)
    pcr_label = Label(bottom_frame, text="PCR:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    pcr_label.grid(row=2, column=4, columnspan=2, sticky=N + S + W + E)
    pcr_val = Label(bottom_frame, text="", relief=RIDGE)
    pcr_val.grid(row=2, column=6, columnspan=2, sticky=N + S + W + E)
    call_exits_label = Label(bottom_frame, text="Call Exits:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    call_exits_label.grid(row=3, column=0, columnspan=2, sticky=N + S + W + E)
    call_exits_val = Label(bottom_frame, text="", relief=RIDGE)
    call_exits_val.grid(row=3, column=2, columnspan=2, sticky=N + S + W + E)
    put_exits_label = Label(bottom_frame, text="Put Exits:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    put_exits_label.grid(row=3, column=4, columnspan=2, sticky=N + S + W + E)
    put_exits_val = Label(bottom_frame, text="", relief=RIDGE)
    put_exits_val.grid(row=3, column=6, columnspan=2, sticky=N + S + W + E)
    call_itm_label = Label(bottom_frame, text="Call ITM:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    call_itm_label.grid(row=4, column=0, columnspan=2, sticky=N + S + W + E)
    call_itm_val = Label(bottom_frame, text="", relief=RIDGE)
    call_itm_val.grid(row=4, column=2, columnspan=2, sticky=N + S + W + E)
    put_itm_label = Label(bottom_frame, text="Put ITM:", relief=RIDGE, font=("TkDefaultFont", 9, "bold"))
    put_itm_label.grid(row=4, column=4, columnspan=2, sticky=N + S + W + E)
    put_itm_val = Label(bottom_frame, text="", relief=RIDGE)
    put_itm_val.grid(row=4, column=6, columnspan=2, sticky=N + S + W + E)

    root.after(100, main)

    root.mainloop()


def main():
    global stop
    global login
    global previous_time
    global max_call_oi_val
    global max_call_oi_sp_val
    global max_put_oi_sp_val
    global max_put_oi_val
    global oi_val
    global pcr_val
    global call_exits_val
    global put_exits_val
    global call_itm_val
    global put_itm_val
    global old_points
    global old_call_sum
    global old_put_sum
    global old_difference
    global old_call_boundary
    global old_put_boundary
    global old_call_itm
    global old_put_itm
    global sheet

    if stop == 1:
        return

    try:
        response = requests.get(url, headers=headers, timeout=5)
    except (requests.exceptions.MissingSchema, requests.exceptions.InvalidSchema, requests.exceptions.URLRequired,
            requests.exceptions.InvalidURL):
        messagebox.showerror(title="Error", message="Incorrect URL.\nPlease enter correct URL.")
        root.destroy()
        return
    except Exception:
        if stop == 1:
            return
        root.after((seconds * 1000), main)
        return

    html_content = bs4.BeautifulSoup(response.content, "html.parser")

    try:
        str_current_time = html_content.findAll('span')[1].text.split(" ")[5]
        current_time = datetime.datetime.strptime(str_current_time, '%H:%M:%S').time()
    except IndexError:
        messagebox.showerror(title="Error", message="Incorrect URL.\nPlease enter correct URL.")
        root.destroy()
        return

    if previous_time is None:
        previous_time = current_time
    elif current_time > previous_time:
        previous_time = current_time
    else:
        if stop == 1:
            return

        root.after((seconds * 1000), main)
        return

    # print("------------------------------------------------------------------------------------------------------",
    #      end='')
    # print("-----------------------------------------------------")

    table = html_content.find('table', {'id': 'octable'})
    links = table.findAll('th')
    rows = table.find_all('tr')
    underlying_stock = (html_content.findAll('b')[0]).text.split(" ")[0]
    points = float((html_content.findAll('b')[0]).text.split(" ")[1])

    try:
        expiry_dates = html_content.find('select', {'id': 'date', 'class': 'goodTextBox'})
        expiry_date = expiry_dates.find('option', selected=True).text
    except IndexError:
        messagebox.showerror(title="Error", message="Incorrect URL.\nPlease enter correct URL.")

    data = []
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele])

    long_head_row = []
    for i in range(4, 25):
        long_head_row.append(links[i].get('title'))

    df = pandas.DataFrame()
    pandas.set_option('display.max_rows', None)
    pandas.set_option('display.max_columns', None)
    pandas.set_option('display.width', 320)

    y = 0
    for i in range(2, (len(data) - 1)):
        df[str(y)] = data[i]
        y += 1
    df = df.transpose()
    df.columns = long_head_row

    # print(df, "\n")

    # print(current_time, end=', ')
    # print(points, end=', ')

    call_oi_list = []
    for i in range(len(df)):
        if df.iloc[i, [0]][0] == "-":
            int_call_oi = 0
        else:
            int_call_oi = int(df.iloc[i, [0]][0].replace(',', ''))
        call_oi_list.append(int_call_oi)
    call_oi_index = call_oi_list.index(max(call_oi_list))
    max_call_oi = round(max(call_oi_list) / 100000, 1)

    put_oi_list = []
    for i in range(len(df)):
        if df.iloc[i, [20]][0] == "-":
            int_put_oi = 0
        else:
            int_put_oi = int(df.iloc[i, [20]][0].replace(',', ''))
        put_oi_list.append(int_put_oi)
    put_oi_index = put_oi_list.index(max(put_oi_list))
    max_put_oi = round(max(put_oi_list) / 100000, 1)

    total_call_oi = sum(call_oi_list)
    total_put_oi = sum(put_oi_list)
    try:
        put_call_ratio = round(total_put_oi / total_call_oi, 2)
    except ZeroDivisionError:
        put_call_ratio = 0

    max_call_oi_sp = df.iloc[call_oi_index]['Strike Price']
    max_put_oi_sp = df.iloc[put_oi_index]['Strike Price']

    # print(f"Max Call OI: {max_call_oi} & Strike Price: {max_call_oi_sp}, ", end='')
    # print(f"Max Put OI: {max_put_oi} & Strike Price: {max_put_oi_sp}, ", end='')
    # print(f"PCR: {put_call_ratio}, ")

    try:
        index = int(df[df['Strike Price'] == f'{sp}'].index.tolist()[0])
    except IndexError:
        messagebox.showerror(title="Error",
                             message="Incorrect Strike Price.\nPlease enter correct Strike Price.")
        root.destroy()
        return

    a = df[['Change in Open Interest']][df['Strike Price'] == f'{sp}']
    b1 = a.iloc[:, 0]
    c1 = b1.get(0, 'Change in Open Interest')
    b2 = df.iloc[:, 1]
    c2 = b2.get((index + 1), 'Change in Open Interest')
    b3 = df.iloc[:, 1]
    c3 = b3.get((index + 2), 'Change in Open Interest')
    try:
        c1 = int(c1.replace(',', ''))
    except ValueError:
        c1 = 0
    try:
        c2 = int(c2.replace(',', ''))
    except ValueError:
        c2 = 0
    try:
        c3 = int(c3.replace(',', ''))
    except ValueError:
        c3 = 0
    call_sum = round((c1 + c2 + c3) / 100000, 1)
    call_boundary = round(c3 / 100000, 1)

    o1 = a.iloc[:, 1]
    p1 = o1.get(0, 'Change in Open Interest')
    o2 = df.iloc[:, 19]
    p2 = o2.get((index + 1), 'Change in Open Interest')
    p3 = o2.get((index + 2), 'Change in Open Interest')
    p4 = o2.get((index + 4), 'Change in Open Interest')
    o3 = df.iloc[:, 1]
    p5 = o3.get((index + 4), 'Change in Open Interest')
    p6 = o3.get((index - 2), 'Change in Open Interest')
    p7 = o2.get((index - 2), 'Change in Open Interest')
    try:
        p1 = int(p1.replace(',', ''))
    except ValueError:
        p1 = 0
    try:
        p2 = int(p2.replace(',', ''))
    except ValueError:
        p2 = 0
    try:
        p3 = int(p3.replace(',', ''))
    except ValueError:
        p3 = 0
    try:
        p4 = int(p4.replace(',', ''))
    except ValueError:
        p4 = 0
    try:
        p5 = int(p5.replace(',', ''))
    except ValueError:
        p5 = 0
    try:
        p6 = int(p6.replace(',', ''))
    except ValueError:
        p6 = 0
    try:
        p7 = int(p7.replace(',', ''))
    except ValueError:
        p7 = 0

    put_sum = round((p1 + p2 + p3) / 100000, 1)
    put_boundary = round(p1 / 100000, 1)
    difference = round(call_sum - put_sum, 1)
    try:
        call_itm = round(p4 / p5, 1)
    except ZeroDivisionError:
        call_itm = 0
    try:
        put_itm = round(p6 / p7, 1)
    except ZeroDivisionError:
        put_itm = 0

    # print(f"Call Sum: {call_sum}, ", end='')
    # print(f"Put Sum: {put_sum}, ", end='')
    # print(f"Difference: {difference}, ", end='')
    # print(f"Call Boundary: {call_boundary}, ", end='')
    # print(f"Put Boundary: {put_boundary}, ", end='')
    # print(f"Call ITM: {call_itm}, ", end='')
    # print(f"Put ITM: {put_itm}, ")

    root.title(f"NSE-Option-Chain-Analyzer - {underlying_stock} - {expiry_date} - {sp}")

    max_call_oi_val.config(text=max_call_oi)
    max_call_oi_sp_val.config(text=max_call_oi_sp)
    max_put_oi_val.config(text=max_put_oi)
    max_put_oi_sp_val.config(text=max_put_oi_sp)

    red = "#e53935"
    green = "#00e676"
    default = "SystemButtonFace"

    if call_sum >= put_sum:
        oi_val.config(text="Bearish", bg=red)
    # print("OI: Bearish, ", end='')
    else:
        oi_val.config(text="Bullish", bg=green)
        # print("OI: Bullish, ", end='')
    if put_call_ratio >= 1:
        pcr_val.config(text=put_call_ratio, bg=green)
    else:
        pcr_val.config(text=put_call_ratio, bg=red)

    def set_itm_labels(call_change, put_change):
        label = "No"
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

    call = set_itm_labels(call_change=p5, put_change=p4)

    if call == "No":
        call_itm_val.config(text="No", bg=default)
        # print("Call ITM: No, ", end='')
    else:
        call_itm_val.config(text="Yes", bg=green)
        # print("Call ITM: Yes, ", end='')

    put = set_itm_labels(call_change=p7, put_change=p6)

    if put == "No":
        put_itm_val.config(text="No", bg=default)
        # print("Put ITM: No, ", end='')
    else:
        put_itm_val.config(text="Yes", bg=red)
        # print("Put ITM: Yes, ", end='')

    if call_boundary <= 0:
        call_exits_val.config(text="Yes", bg=green)
        # print("Call Exits: Yes, ", end='')
    elif call_sum <= 0:
        call_exits_val.config(text="Yes", bg=green)
        # print("Call Exits: Yes, ", end='')
    else:
        call_exits_val.config(text="No", bg=default)
        # print("Call Exits: No, ", end='')
    if put_boundary <= 0:
        put_exits_val.config(text="Yes", bg=red)
        # print("Put Exits: Yes ")
    elif put_sum <= 0:
        put_exits_val.config(text="Yes", bg=red)
        # print("Put Exits: Yes ")
    else:
        put_exits_val.config(text="No", bg=default)
        # print("Put Exits: No ")

    output_values = [str_current_time, points, call_sum, put_sum, difference, call_boundary, put_boundary, call_itm,
                     put_itm]
    sheet.insert_row(values=output_values)

    last_row = sheet.get_total_rows() - 1

    if old_points is None or points == old_points:
        old_points = points
    elif points > old_points:
        sheet.highlight_cells(row=last_row, column=1, bg=green)
        old_points = points
    else:
        sheet.highlight_cells(row=last_row, column=1, bg=red)
        old_points = points
    if old_call_sum is None or old_call_sum == call_sum:
        old_call_sum = call_sum
    elif call_sum > old_call_sum:
        sheet.highlight_cells(row=last_row, column=2, bg=red)
        old_call_sum = call_sum
    else:
        sheet.highlight_cells(row=last_row, column=2, bg=green)
        old_call_sum = call_sum
    if old_put_sum is None or old_put_sum == put_sum:
        old_put_sum = put_sum
    elif put_sum > old_put_sum:
        sheet.highlight_cells(row=last_row, column=3, bg=green)
        old_put_sum = put_sum
    else:
        sheet.highlight_cells(row=last_row, column=3, bg=red)
        old_put_sum = put_sum
    if old_difference is None or old_difference == difference:
        old_difference = difference
    elif difference > old_difference:
        sheet.highlight_cells(row=last_row, column=4, bg=red)
        old_difference = difference
    else:
        sheet.highlight_cells(row=last_row, column=4, bg=green)
        old_difference = difference
    if old_call_boundary is None or old_call_boundary == call_boundary:
        old_call_boundary = call_boundary
    elif call_boundary > old_call_boundary:
        sheet.highlight_cells(row=last_row, column=5, bg=red)
        old_call_boundary = call_boundary
    else:
        sheet.highlight_cells(row=last_row, column=5, bg=green)
        old_call_boundary = call_boundary
    if old_put_boundary is None or old_put_boundary == put_boundary:
        old_put_boundary = put_boundary
    elif put_boundary > old_put_boundary:
        sheet.highlight_cells(row=last_row, column=6, bg=green)
        old_put_boundary = put_boundary
    else:
        sheet.highlight_cells(row=last_row, column=6, bg=red)
        old_put_boundary = put_boundary
    if old_call_itm is None or old_call_itm == call_itm:
        old_call_itm = call_itm
    elif call_itm > old_call_itm:
        sheet.highlight_cells(row=last_row, column=7, bg=green)
        old_call_itm = call_itm
    else:
        sheet.highlight_cells(row=last_row, column=7, bg=red)
        old_call_itm = call_itm
    if old_put_itm is None or old_put_itm == put_itm:
        old_put_itm = put_itm
    elif put_itm > old_put_itm:
        sheet.highlight_cells(row=last_row, column=8, bg=red)
        old_put_itm = put_itm
    else:
        sheet.highlight_cells(row=last_row, column=8, bg=green)
        old_put_itm = put_itm

    if sheet.get_yview()[1] >= 0.9:
        sheet.see(last_row)
        sheet.set_yview(1)
    sheet.refresh()

    if stop == 1:
        return

    root.after((seconds * 1000), main)
    return


login()
