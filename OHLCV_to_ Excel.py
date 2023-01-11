# По API Бинанс получить исторические данные OHLCV для дневного
# таймфрейма пары BTC/USDT и загрузить их в Эксель

import json
from datetime import datetime as dt
import requests
import xlsxwriter


r = requests.get("https://api.binance.com/api/v3/klines",
                 params={"symbol": "BTCUSDT", "interval": '1d', 'limit': 5})
r_load = json.loads(r.text)

for _list in r_load:
    _list[0] = dt.fromtimestamp(_list[0] / 1000).strftime('%d.%m.%Y')
    del _list[6:]
    _list.insert(0, 'BTC/USDT')
    index = 2
    for num in _list[2:]:
        _list[index] = float(num)
        index += 1

# find max length for each column
# useful if you have to create many different excel tables
max_lens = []
for _list in r_load:
    for index in range(len(_list)):
        try:
            max_lens[index].append(len(str(_list[index])))
        except:
            max_lens.insert(index, [])
max_lens = [max(lens) for lens in max_lens]

book = xlsxwriter.Workbook('file1.xlsx')
sheet = book.add_worksheet()

# make headers bold
headers = ['TICKER', 'DATE', 'OPEN', 'HIGH', 'LOW', 'CLOSE', 'VOLUME']
bold = book.add_format({'bold': True})
for index in range(len(headers)):
    sheet.write(0, index, headers[index], bold)

row_number = 2
for _list in r_load:
    for index in range(len(_list)):
        sheet.write(row_number, index, _list[index])
        sheet.set_column(row_number, index, max_lens[index] + 1)
    row_number += 1

book.close()
