'''
生成单个滑点统计报表
luolan
'''


import datetime as dt
from DataQuery import query
from Report.Report import write_report_ETH
from Report.Report import write_report_XBT


f_name = "note3.txt"
qr = query.query(f_name=f_name)

#bitmex_4a

file_path = "C:/Users/lmy/Desktop/交易记录/"
date = dt.date(2019,6,16)
date1 = date.strftime("%Y%m%d")
date2 = (date + dt.timedelta(days=1)).strftime("%Y%m%d")

# file_name = "bitmex_4a_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
# strategy = "bitmex_4a_ETH_Balance"
# write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

# file_name = "bitmex_4a_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
# strategy = "bitmex_4a_BTC_Balance"
# write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_4a_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_4a_BTC_Balance"
write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_4a_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_4a_ETH_Balance"
write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_4b_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_4b_BTC_Balance"
write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_4b_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_4b_ETH_Balance"
write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_666a_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_666a_BTC_Balance"
write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_666a_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_666a_ETH_Balance"
write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_5_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_5_BTC_Balance"
write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_5_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_5_ETH_Balance"
write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_a1_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_a1_BTC_Balance"
write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_a1_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_a1_ETH_Balance"
write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_5a_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_5a_BTC_Balance"
write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)

file_name = "bitmex_5a_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
strategy = "bitmex_5a_ETH_Balance"
write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)


