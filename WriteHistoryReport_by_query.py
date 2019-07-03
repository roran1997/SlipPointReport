'''
通过循环调用Report中函数，批量生成滑点统计报表
luolan
'''


import datetime as dt
from DataQuery import query
from Report.Report import write_report_ETH
from Report.Report import write_report_XBT


f_name = "note.txt"
qr = query.query(f_name=f_name)

#bitmex_4a

file_path = "C:/Users/lmy/Desktop/Bitmex/daily_slip_point_analysis(new)/bitmex_4a/"
date = dt.date(2019,2,20)
date_end = dt.date(2019,5,29)
while(date< date_end):
    date1 = date.strftime("%Y%m%d")
    date2 = (date + dt.timedelta(days=1)).strftime("%Y%m%d")

    # ETH
    file_name = "bitmex_4a_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_4a_ETH_Balance"
    try:
        write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:"+ strategy+" " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy+" " + date.strftime("%Y-%m-%d")+'\n')
        f.close()
    #XBT
    file_name = "bitmex_4a_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_4a_BTC_Balance"
    try:
        write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:" + strategy + " " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy + " " + date.strftime("%Y-%m-%d") + '\n')
        f.close()

    date = date+dt.timedelta(days=1)

#bitmex_4b

file_path = "C:/Users/lmy/Desktop/Bitmex/daily_slip_point_analysis(new)/bitmex_4b/"
date = dt.date(2019,3,14)
date_end = dt.date(2019,5,29)
while(date< date_end):
    date1 = date.strftime("%Y%m%d")
    date2 = (date + dt.timedelta(days=1)).strftime("%Y%m%d")

    # ETH
    file_name = "bitmex_4b_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_4b_ETH_Balance"
    try:
        write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:"+ strategy+" " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy+" " + date.strftime("%Y-%m-%d")+'\n')
        f.close()
    #XBT
    file_name = "bitmex_4b_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_4b_BTC_Balance"
    try:
        write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:" + strategy + " " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy + " " + date.strftime("%Y-%m-%d") + '\n')
        f.close()

    date = date+dt.timedelta(days=1)


#bitmex_666a

file_path = "C:/Users/lmy/Desktop/Bitmex/daily_slip_point_analysis(new)/bitmex_666a/"
date = dt.date(2019,3,14)
date_end = dt.date(2019,5,29)
while(date< date_end):
    date1 = date.strftime("%Y%m%d")
    date2 = (date + dt.timedelta(days=1)).strftime("%Y%m%d")

    # ETH
    file_name = "bitmex_666a_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_666a_ETH_Balance"
    try:
        write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:"+ strategy+" " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy+" " + date.strftime("%Y-%m-%d")+'\n')
        f.close()
    #XBT
    file_name = "bitmex_666a_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_666a_BTC_Balance"
    try:
        write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:" + strategy + " " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy + " " + date.strftime("%Y-%m-%d") + '\n')
        f.close()

    date = date+dt.timedelta(days=1)


#bitmex_5

file_path = "C:/Users/lmy/Desktop/Bitmex/daily_slip_point_analysis(new)/bitmex_4a/"
date = dt.date(2019,5,20)
date_end = dt.date(2019,5,29)
while(date< date_end):
    date1 = date.strftime("%Y%m%d")
    date2 = (date + dt.timedelta(days=1)).strftime("%Y%m%d")

    # ETH
    file_name = "bitmex_5_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_5_ETH_Balance"
    try:
        write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:"+ strategy+" " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy+" " + date.strftime("%Y-%m-%d")+'\n')
        f.close()
    #XBT
    file_name = "bitmex_5_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_5_BTC_Balance"
    try:
        write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:" + strategy + " " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy + " " + date.strftime("%Y-%m-%d") + '\n')
        f.close()

    date = date+dt.timedelta(days=1)

#bitmex_a1

file_path = "C:/Users/lmy/Desktop/Bitmex/daily_slip_point_analysis(new)/bitmex_4a/"
date = dt.date(2019,5,21)
date_end = dt.date(2019,5,29)
while(date< date_end):
    date1 = date.strftime("%Y%m%d")
    date2 = (date + dt.timedelta(days=1)).strftime("%Y%m%d")

    # ETH
    file_name = "bitmex_a1_ETH_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_a1_ETH_Balance"
    try:
        write_report_ETH(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:"+ strategy+" " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy+" " + date.strftime("%Y-%m-%d")+'\n')
        f.close()
    #XBT
    file_name = "bitmex_a1_BTC_Balance_" + date1 + "_" + date2 + ".xlsx"
    strategy = "bitmex_a1_BTC_Balance"
    try:
        write_report_XBT(file_path, file_name, date.strftime("%Y-%m-%d"), qr, strategy)
    except:
        print("report error:" + strategy + " " + date.strftime("%Y-%m-%d"))
        f = open(f_name, 'a')
        f.write("report error:" + strategy + " " + date.strftime("%Y-%m-%d") + '\n')
        f.close()

    date = date+dt.timedelta(days=1)