#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import datetime as dt

import calendar
import pandas as pd
from cachetools import cached, LRUCache
from cachetools.keys import hashkey
from functools import partial
from threading import RLock

cache2 = LRUCache(maxsize=30)
cache_lock = RLock()
today = dt.date.today()
now = dt.datetime.now()


# custom_day = dt.date(year=, month=, day=)
def _isoweek1friday(year):
    """ Helper to calculate the day number of the Friday starting week 1"""
    FRIDAY = 4  # week 01 is the week with the year's first Friday in it
    firstday = dt.date(year, 1, 1).toordinal()
    firstweekday = (firstday + 6) % 7
    week1friday = firstday - firstweekday + 4
    if firstweekday > FRIDAY:
        week1friday += 7
    return week1friday


def friday_calendar(date):
    """ calculate the calendar which week number is based on friday is the first day of week"""
    year = date.year
    week1friday = _isoweek1friday(year)
    date_day = date.toordinal()
    week, day = divmod(date_day - week1friday, 7)
    if week < 0:
        year -= 1
        week1friday = _isoweek1friday(year)
        week, day = divmod(date_day - week1friday, 7)
    elif week >= 52:
        if date_day >= _isoweek1friday(year + 1):
            year += 1
            week = 0
    return year, week + 1, day + 5


def date_quarter(date):
    return (date.month - 1) // 3 + 1


def last_week(date):
    result = date - dt.timedelta(days=7)
    return result


def last_month(date, return_date=False):
    first = date.replace(day=1)
    lastmonth = first - dt.timedelta(days=1)
    if return_date:
        return lastmonth
    else:
        return short_month(lastmonth)


def next_week(date):
    result = date + dt.timedelta(days=7)
    return result


# @cached(cache2, key=partial(hashkey, 'week_check_day'))
# def week_check_day(date):
#     if date.weekday() < 4:  # mon, tue, wen, thursday
#         date = last_week(date)
#     this_monday = date - dt.timedelta(days=date.weekday())
#     return this_monday

@cached(cache2, key=partial(hashkey, 'week_check_day'), lock=cache_lock)
def week_check_day(date, anchor_day=4):
    """
    Determine which day to produce week number, chk cross week, cross month, cross year ... etc
    :param date: target_day, query data between target_day and last_week(target_day)
    :param anchor_day: default 4, use friday as anchor_day to produce week number... etc
    :return: this_week_anchor_day
    """
    # if date.weekday() < 4 + anchor_day:
    #     date = last_week(date)
    this_week_day = date - dt.timedelta(days=date.weekday() - anchor_day)
    return this_week_day


@cached(cache2, key=partial(hashkey, 'week_number'), lock=cache_lock)
def week_number(date):
    """
    return week number like 'W52' or 'W01'
    :param date:
    :return:
    """
    # result = date.isocalendar()[1]
    result = friday_calendar(date)[1]
    return 'W' + str(result).zfill(2)


def week_interval(end_date, start_date=None):
    """
    return: like '20181221-20181228' string
    """
    if start_date is None:
        start_date = last_week(end_date)
    result = str(start_date).replace('-', '') + '-' + str(end_date).replace('-', '')
    return result


def week_interval_to_date(interval):
    dtstr = interval.rsplit('-')[-1]
    date = dt.datetime.strptime(dtstr, '%Y%m%d').date()
    return date


def chk_cross_month(date):
    """
    chk last week and this week if or not cross month
    :return: boolean True or False
    """
    # date = date - dt.timedelta(days=1)  # a week start at friday and end at thursday, friday - 1 = thursday
    return bool(date.month - last_week(date).month)


# def chk_cross_month(date, dtstring, monday=True):
#     """
#     chk last week and this week if or not cross month
#     :return: boolean True or False
#     """
#     if dtstring is np.nan:
#         return False
#     dtstr = dtstring.split('-')[1]
#     last_date = dt.datetime.strptime(dtstr, '%Y%m%d').date()
#     if monday:
#         last_date = week_check_day(last_date)
#     return bool(date.month - last_date.month)


def chk_cross_quarter(date, last_date=None):
    # date = date - dt.timedelta(days=1)  # a week start at friday and end at thursday, friday - 1 = thursday
    if last_date is None:
        last_date = last_week(date)
    return bool(date_quarter(date) - date_quarter(last_date))


def chk_every_2_quarter(date, last_date=None):
    if chk_cross_quarter(date, last_date=last_date):
        q = date_quarter(date)
        return bool(q % 2)
    return False


def chk_cross_year(date):
    """
    chk if or not last week and this week cross year
    :return: boolean True or False
    """
    # date = date - dt.timedelta(days=1)  # a week start at friday and end at thursday, friday - 1 = thursday
    return bool(date.year - last_week(date).year)


def short_month(date):
    """
    :param date:
    :return: 'Jan', 'Fab', 'Mar' ....
    """
    return calendar.month_abbr[date.month]


# def week_numbers_in_last_month(date, only_week_number=True):
#     """
#     only work when last_week of date is cross_month
#     :param date:set as friday
#     :param only_week_number
#     :return: week numbers of last month contained in list format like ['W39', 'W38', 'W37', 'W36', 'W35']
#     only
#     """
#     # year = date.year
#     # month = date.month
#     # ending_day = calendar.monthrange(year, month)[1]  # get the last day of month
#     # initial_week = dt.datetime(year, month, 1).isocalendar()[1]
#     # ending_week = dt.datetime(year, month, ending_day).isocalendar()[1]
#     # return range(initial_week, ending_week + 1)
#     week_numbers, months, shipping_dates = ([] for l in range(3))
#     if chk_cross_month(date=date):
#         while True:
#             date = last_week(date=date)
#             week_numbers.insert(0, week_number(date=date))
#             months.insert(0, short_month(date=date))
#             shipping_dates.insert(0, week_interval(date=date))
#             # else:
#             #     df = pd.DataFrame(data={'month': short_month(date=date), 'shipping date': week_interval(date=date),
#             #                             'week': week_number(date=date)})
#             if chk_cross_month(date=date):
#                 break
#     if only_week_number:
#         return week_numbers
#     else:
#         df = pd.DataFrame(data={('SUMMARY', 'month'): months, ('SUMMARY', 'shipping_date'): shipping_dates,
#                                 ('SUMMARY', 'week'): week_numbers})
#         return df


def time_header(date, multi_index=True, start_date=None, check_day=True):
    if date is None:
        return pd.DataFrame()
    if check_day:
        date = week_check_day(date)
    if start_date is None:
        start_date = last_week(date)
    # date = date - dt.timedelta(days=1)  # a week start at friday and end at thursday, friday - 1 = thursday

    # if check_day:
    #     month = [short_month(week_check_day(date))]
    #     week = [week_number(week_check_day(date))]
    # else:
    month = [short_month(date)]
    week = [week_number(date)]
    shipping_date = [week_interval(date, start_date)]
    if multi_index:
        df = pd.DataFrame(data={('DATETIME', 'month'): month, ('DATETIME', 'shipping_date'): shipping_date,
                                ('DATETIME', 'week'): week})
    else:
        df = pd.DataFrame(data={'month': month, 'shipping_date': shipping_date, 'week': week})
    return df


def main():
    print(week_number())


if __name__ == '__main__':
    main()
