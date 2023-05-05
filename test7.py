#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import datetime as dt

import pandas as pd

from Database_query import connection, equipment_lot_history
from WeekTime_parser import last_week, today, next_week


def query_implant_lot_history(target_date=None, end_date=None, start_date=None, filename=3):
    connection()
    if target_date is None:
        target_date = dt.date(year=2018, month=12, day=28)
    if end_date is None or end_date >= today:
        end_date = today
    if start_date is None:
        start_date = last_week(target_date)
    concat_list = []
    while True:
        target_date_str = target_date.strftime('%Y-%m-%d')
        start_date_str = start_date.strftime('%Y-%m-%d')
        df = equipment_lot_history("L182501", target_date, start_date)
        concat_list.append(df)
        target_date = next_week(target_date)
        start_date = next_week(start_date)
        print(f"{start_date_str},{target_date_str}")
        if isinstance(target_date, dt.date):
            if target_date > end_date:
                break
    df_result = pd.concat(concat_list, ignore_index=True, sort=False)
    df_result.to_excel(r"C:\Users\ernie.chen\Desktop\ReportGen ver2\AP197 Rin\imp_time{}.xlsx".format(filename),
                       sheet_name='result')


def main():
    target_date = dt.date(year=2019, month=9, day=6)
    # mid_date = dt.date(year=2020, month=5, day=1)
    # query_implant_lot_history(target_date=target_date, end_date=mid_date)
    # query_implant_lot_history(target_date=next_week(mid_date), filename=1)
    query_implant_lot_history(target_date=target_date, filename=2)
    # import sys; print('Python %s on %s' % (sys.version, sys.platform))
    # sys.path.extend(['C:\\Users\\ernie.chen\\Desktop\\Project\\Project Week', 'C:/Users/ernie.chen/Desktop/Project/Project Week'])
    # exec(open(r"C:\Users\ernie.chen\Desktop\Project\Project Week\test7.py").read())


if __name__ == '__main__':
    main()
