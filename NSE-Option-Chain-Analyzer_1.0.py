from bs4 import BeautifulSoup
import requests
import pandas as pd
import time

url = str(input("Enter URL: "))
headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/80.0.3987.149 Safari/537.36',
           'accept-language': 'en,gu;q=0.9,hi;q=0.8', 'accept-encoding': 'gzip, deflate, br'}
sp = str(input("Enter Strike Price: "))

previous_time = ""
seconds = 30

while True:
    try:
        response = requests.get(url, headers=headers)
    except:
        time.sleep(seconds)
        continue
    
    html_content = BeautifulSoup(response.content, "html.parser")
    my_time = (html_content.findAll('span')[1]).text.split(" ")[5]
    if previous_time == "":
        previous_time = my_time
    elif previous_time != my_time:
        previous_time = my_time
    else:
        time.sleep(seconds)
        continue
        
    print("------------------------------------------------------------------------------------------------------",
          end='')
    print("-----------------------------------------------------")

    my_table = html_content.find('table', {'id': 'octable'})
    links = my_table.findAll('th')
    rows = my_table.find_all('tr')
    my_points = (html_content.findAll('b')[0]).text.split(" ")[1]
    data = []
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele])

    long_head_row = []
    for i in range(4, 25):
        for link in links:
            long_head_row.append(links[i].get('title'))
            break

    df = pd.DataFrame()
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 320)
    y = 0
    for i in range(2, (len(data) - 1)):
        df[str(y)] = data[i]
        y += 1
    df = df.transpose()
    df.columns = long_head_row
    # print(df, "\n")
    print(my_time, end=' ')
    print(my_points, end=' ')

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

    print(f"Max Call OI: {max_call_oi}, Strike Price: {df.iloc[call_oi_index]['Strike Price']} ", end='')
    print(f"Max Put OI: {max_put_oi}, Strike Price: {df.iloc[put_oi_index]['Strike Price']} ", end='')
    print(f"PCR: {put_call_ratio} ")

    index = int(df[df['Strike Price'] == f'{sp}'].index.tolist()[0])
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

    print(f"Call Sum: {call_sum} ", end='')
    print(f"Put Sum: {put_sum} ", end='')
    print(f"Difference: {difference} ", end='')
    print(f"Call Boundary: {call_boundary} ", end='')
    print(f"Put Boundary: {put_boundary} ", end='')
    print(f"Call ITM: {call_itm} ", end='')
    print(f"Put ITM: {put_itm} ")

    time.sleep(seconds)
