#!/usr/bin/env python
# _*_ coding:utf-8 _*_3

import calendar
import Template_generator as Tg
import pandas as pd
import os
import datetime as dt
from Monthly_parser import get_monthly_data, multi_idx_df_multiply_one_col
from df_to_excel import df_update_xlsx, cell_week_content
from excel_styler import WeeklyHeaderStyles
from WeekTime_parser import date_quarter

minimal_pcs = int(Tg.cfg['update bsl minimal pcs'])


def set_uni(s, d):
    result_set = set()
    for item in s:
        result_set |= d[item]
    return result_set


# {1: {'Feb', 'Mar', 'Jan'}, 2: {'May', 'Apr', 'Jun'}, 3: {'Aug', 'Jul', 'Sep'}, 4: {'Dec', 'Oct', 'Nov'}}
quarter_map = dict()
for m in range(1, 13):
    q = (m - 1) // 3 + 1
    qm_set = quarter_map.get(q, set())
    qm_set.add(calendar.month_abbr[m])
    quarter_map[q] = qm_set
# Q1 search last year Q3, Q4, Q3 search Q1, Q2
search_map = {1: {4, 3}, 3: {2, 1}}
q_str_map = {1: ' Q3-Q4', 3: ' Q1-Q2'}
# {1: {'Aug', 'Dec', 'Jul', 'Oct', 'Nov', 'Sep'}, 3: {'Apr', 'May', 'Jun', 'Jan', 'Mar', 'Feb'}}
quarter_search_map = {k: set_uni(s, quarter_map) for k, s in search_map.items()}


def bsl_update(date):
    quarter = date_quarter(date)
    writer = r'reference/baseline.xlsx'
    new_file = not os.path.isfile(writer)
    for part in Tg.bsl_parts:
        df = get_monthly_data(part, search_quarter=True)
        if df.empty:
            continue
        df.drop(columns=[('SUMMARY', 'Ignored_wafers')], inplace=True)
        df.reindex(quarter_search_map[quarter]).dropna()
        year = df.index.name.split("_")[0]
        multi_idx_df_multiply_one_col(df)
        df = pd.DataFrame(df.sum()).T
        if df.at[df.last_valid_index(), ('SUMMARY', "Q'ty")] >= minimal_pcs:
            multi_idx_df_multiply_one_col(df, operator='/')
            df.index = pd.Index([year + q_str_map[quarter]])
            writer = df_update_xlsx(df, writer, part, apply_axis=None, multindex_headers=2,
                                    cell_width_height=cell_week_content, fit_col_width=False, header=new_file,
                                    truncate_sheet=new_file, header_style=WeeklyHeaderStyles if new_file else None,
                                    return_writer=True)
    if not isinstance(writer, str):
        writer.save()


def main():
    bsl_update(dt.date(2019, 7, 4))


if __name__ == '__main__':
    main()
