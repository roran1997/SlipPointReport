import psycopg2
import datetime as dt


class query():
    def __init__(self,f_name):
        """
        :param f_name: 后缀为.txt文件名称，用于记录一些异常情况
        """
        self.conn_risk = psycopg2.connect(database='',user="", password="",host="", port="")
        self.conn_history_data = psycopg2.connect(database='',user="", password="",host="", port="")
        self.f_name = f_name

    def get_price(self, datetime:str, name, label = 'close'):
        datetime = dt.datetime.strptime(datetime, "%Y-%m-%d %H:%M:%S")
        cur =self.conn_history_data.cursor()
        if name == 'eth':
            sql = "Select "+label + ", datetime from public.bitmex_eth_usd_1m " \
                  "where datetime >= '" + datetime.strftime("%Y-%m-%d %H:%M:%S") + "' " \
                  "and datetime < '" + (datetime + dt.timedelta(minutes=1)).strftime("%Y-%m-%d %H:%M:%S") + "'" \
                  ";"
        elif name == 'xbt':
            sql = "Select "+label+", datetime from public.bitmex_xbt_usd_1m " \
                  "where datetime >= '" + datetime.strftime("%Y-%m-%d %H:%M:%S") + "' " \
                  "and datetime < '" + (datetime + dt.timedelta(minutes=1)).strftime("%Y-%m-%d %H:%M:%S") + "'" \
                  ";"
        else:
            print('Error:Name is not \'xbt\' or \'eth\'')
            return None
        cur.execute(sql)
        rows = cur.fetchall()
        if len(rows)==0:
            print("query: No "+name+" price data at " + datetime.strftime("%Y-%m-%d %H:%M:%S"))
            f = open(self.f_name,'a')
            f.write("query: No "+name+" price data at " + datetime.strftime("%Y-%m-%d %H:%M:%S")+'\n')
            f.close()
            return None
        return rows[0][0]

    def get_position(self,datetime:str, strategy):
        datetime = dt.datetime.strptime(datetime, "%Y-%m-%d %H:%M:%S")
        cur = self.conn_risk.cursor()
        sql = "select position from public.position " \
              "where strategy = '" + strategy + "'" \
              "and typ = 'curr'" \
              "and datetime < '" + datetime.strftime("%Y-%m-%d %H:%M:%S") + "' " \
              "order by datetime desc limit 1;"
        cur.execute(sql)
        rows = cur.fetchall()
        if len(rows)==0:
            print("query: No "+strategy+" positon data at "+datetime.strftime("%Y-%m-%d %H:%M:%S"))
            f = open(self.f_name,'a')
            f.write("query: No "+strategy+" positon data at "+datetime.strftime("%Y-%m-%d %H:%M:%S")+'\n')
            f.close()
            return None
        return list(rows[0][0].values())[1]

    def get_trade_event(self,begin_datetime:str, end_datetime:str, strategy):
        cur = self.conn_risk.cursor()
        begin_datetime = dt.datetime.strptime(begin_datetime, "%Y-%m-%d %H:%M:%S")
        end_datetime = dt.datetime.strptime(end_datetime,"%Y-%m-%d %H:%M:%S")
        sql = "select public.trade_event.order_id, strategy, direction, qty, price, datetime,slippage_rate " \
              "from public.trade_event " \
              "left join public.trade_misc " \
              "on public.trade_event.order_id = public.trade_misc.order_id " \
              "where strategy = '" + strategy + "' " \
              "and datetime >= '" + begin_datetime.strftime("%Y-%m-%d %H:%M:%S") + "' " \
              "and datetime < '" + end_datetime.strftime("%Y-%m-%d %H:%M:%S") +"' " \
              "order by datetime asc;"

        cur.execute(sql)
        rows = cur.fetchall()
        if(len(rows))==0:
            f = open(self.f_name,'a')
            f.write("query: No trade event data between " + begin_datetime.strftime("%Y-%m-%d %H:%M:%S")
                    +" and "+ end_datetime.strftime("%Y-%m-%d %H:%M:%S")+'\n')
            f.close()
            return None

        return rows









