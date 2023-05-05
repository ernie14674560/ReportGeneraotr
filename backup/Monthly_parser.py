#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import calendar
import os
from functools import reduce

import numpy as np
import pandas as pd

import Template_generator as Tg
from Weekly_parser import get_previous_data, xlsx_sheet_to_df
from df_to_excel import df_update_xlsx, cell_week_content
from excel_styler import WeeklyHeaderStyles

month_abbr_idx = pd.Index(list(calendar.month_abbr)[1:], name='month')
df_month_abbr = pd.DataFrame(index=month_abbr_idx)
df_month_abbr.columns = pd.MultiIndex.from_product([df_month_abbr.columns, ['']])
monthly_setting = Tg.cfg['parts show in monthly pptx']
monthly_parts_and_title = monthly_setting['parts and title']
monthly_merge_parts = monthly_setting['merge parts']
monthly_exclude_parts = monthly_setting['exclude parts']
monthly_show_parts = monthly_setting['parts show in slide']
monthly_use_parts = monthly_setting['parts use in report']


# for title, content in monthly_merge_parts.items():
#     merge_parts.add(tuple(content['parts']))


def chk_month_order(df):
    """

    :param df:
    :return: if order is normal return True, else False
    """
    month_order = [month_abbr_idx.get_loc(idx) for idx in df.index]
    return all(month_order[i] <= month_order[i + 1] for i in range(len(month_order) - 1))


def map_tuple_gen(func, tup):
    """
    Applies _func to each element of tup and returns a new tuple.

     >>a = (1, 2, 3, 4)
     >>_func = lambda x: x * x
     >>map_tuple(_func, a)
     >>(1, 4, 9, 16)
    """
    return tuple(func(itup) for itup in tup)


def drop_unnecessary_cols_and_set_idx(df_month, drop_year=True):
    # del_cols = [('DATETIME', 'month')]
    del_cols = []
    if drop_year:
        del_cols.append(('DATETIME', 'shipping_date'))
    for grp in ['OOC', 'OOC_GROUP']:
        if grp in df_month.columns:
            del_cols.append(grp)
    # if 'OOC' and 'OOC_GROUP' in df_month.columns:
    #     del_cols += ['OOC', 'OOC_GROUP']
    year = df_month.tail(1).at[df_month.last_valid_index(), ('DATETIME', 'shipping_date')]
    df_month.drop(columns=del_cols, inplace=True)
    if drop_year:
        df_month.set_index(('DATETIME', 'month'), inplace=True)
        df_month.index.rename(year + '_month', inplace=True)
    else:
        df_month.set_index([('DATETIME', 'shipping_date'), ('DATETIME', 'month')], inplace=True)
        df_month.index.rename(['year', 'month'], inplace=True)
    # df_month.index.name = 'month'
    # df_month.rename_axis('month')


def search_df_month_idx(df, month):
    df_slice = df[df[('DATETIME', 'month')] == month]
    if df_slice.empty:  # part not have data this month will return empty DataFrame
        i = month_abbr_idx.get_loc(month)
        i -= 1
        if i >= 0:
            pre_month = month_abbr_idx[i]
            slice_idx = search_df_month_idx(df, pre_month)
        else:
            slice_idx = 0
    else:
        slice_idx = df_slice.index[0]
    return slice_idx


def get_monthly_data(part, search_quarter=False, specific_year=None, specific_month=None):
    df = get_previous_data(part, monthly=True, bsl=False, search_quarter=search_quarter, specific_year=specific_year)
    if df.empty:  # part not have data yet will return empty DataFrame
        return df
    if specific_month is not None:
        slice_idx = search_df_month_idx(df, specific_month)
        df = df.loc[:slice_idx]
        if specific_month not in df[('DATETIME', 'month')].values:
            last_i = len(df)
            df.loc[last_i] = np.nan
            df.at[last_i, ('DATETIME', 'month')] = specific_month

    # df_slice = df[df[('DATETIME', 'month')] == specific_month]
    # if df_slice.empty:  # part not have data this month will return empty DataFrame
    #     return df_slice
    # else:
    #     slice_idx = df_slice.index[0]

    if specific_year is not None and len(df) < 2:
        # if len(df) < 2:  # search back last year
        df2 = get_previous_data(part, monthly=True, bsl=False, search_quarter=search_quarter,
                                specific_year=specific_year - 1)
        df = pd.concat([df2, df], ignore_index=True, sort=False)
    drop_unnecessary_cols_and_set_idx(df, drop_year=search_quarter)
    return df


def write_monthly_data(part, specific_year, specific_month, ignored_wafers_list, modified_df):
    search_path = '{}\\{}\\monthly_data_log.xlsx'.format(
        '{}\\Monthly_Report'.format(Tg.read_data_root_folder), '{}_data'.format(specific_year))
    df = xlsx_sheet_to_df(search_path, sheet_name=part, index_col=0, header=[0, 1])
    idx = df.index[df[('DATETIME', 'month')] == specific_month].tolist()[0]
    cols_order = list(modified_df)
    # modified_df.at[modified_df.last_valid_index(), ('SUMMARY', 'Ignored_wafers')] = ignored_wafers_list
    cols_order.insert(1, ('SUMMARY', 'Ignored_wafers'))
    # modified_df = modified_df.loc[:, cols_order]
    df.fillna('', inplace=True)
    for first, second in modified_df.columns:
        df.at[idx, (first, second)] = modified_df.at[(str(specific_year), specific_month), (first, second)]
    wafers = ''.join(e + ', ' for e in ignored_wafers_list).rstrip(', ')

    df.at[idx, ('SUMMARY', 'Ignored_wafers')] = wafers
    df_update_xlsx(df, search_path, sheetname=part, multindex_headers=2,
                   header=True, truncate_sheet=True,
                   header_style=WeeklyHeaderStyles, cell_width_height=cell_week_content,
                   fit_col_width=False)


def monthly_data_fill(df):
    """
    return a year monthly fill na data
    :param df: like df_month = get_monthly_data('AP196')
    :return: df like
              SUMMARY                       ...        CP
             Q'ty     Yield Yield_loss  ...       Rin    Offset       RB1
    month                               ...
    Jan      24.0  0.944275   0.055725  ...  0.000014  0.000165  0.003130
    Feb      24.0  0.937354   0.062646  ...  0.000006  0.000705  0.000633
    Mar       NaN       NaN        NaN  ...       NaN       NaN       NaN
    Apr      10.0  0.964734   0.035266  ...  0.000007  0.000270  0.000548
    May      67.0  0.950028   0.049972  ...  0.000052  0.000372  0.001555
    Jun       NaN       NaN        NaN  ...       NaN       NaN       NaN
    Jul       NaN       NaN        NaN  ...       NaN       NaN       NaN
    Aug       NaN       NaN        NaN  ...       NaN       NaN       NaN
    Sep       NaN       NaN        NaN  ...       NaN       NaN       NaN
    Oct       NaN       NaN        NaN  ...       NaN       NaN       NaN
    Nov       NaN       NaN        NaN  ...       NaN       NaN       NaN
    Dec       NaN       NaN        NaN  ...       NaN       NaN       NaN
    """
    df_month = df.copy()
    # if 'DATETIME' in df_month.columns:
    #     df_month.drop(columns=['DATETIME'], inplace=True)
    # drop_unnecessary_cols_and_set_idx(df_month)
    df_result = df_month_abbr.join(df_month, how='outer').reindex(month_abbr_idx)
    return df_result


def multi_idx_df_multiply_one_col(df, operator='*', col=('SUMMARY', "Q'ty"),
                                  ignored_cols=(('SUMMARY', "Ignored_wafers"),)):
    for first, second in df:
        if second == col[1]:  # Q'ty
            continue
        elif (first, second) in ignored_cols:
            continue
        # df[(first, second)] = df[(first, second)] * df[('SUMMARY', "Q'ty")]
        if operator == '*':
            df[(first, second)] *= df[col]
        elif operator == '/':
            df[(first, second)] /= df[col]


def monthly_df_add_or_sub(df1, df2, operator='+', ignore_col=('SUMMARY', "Ignored_wafers")):
    if ignore_col in df1 and ignore_col in df2:
        df_ign = df1[[ignore_col]] + df2[[ignore_col]]
        df1.drop(columns=[ignore_col], inplace=True)
        df2.drop(columns=[ignore_col], inplace=True)
        if operator == '+':
            df_result = pd.concat([df_ign, df1.add(df2, fill_value=0)], axis=1, sort=False)
        elif operator == '-':
            df_result = pd.concat([df_ign, df1.sub(df2, fill_value=0)], axis=1, sort=False)
    else:
        if operator == '+':
            df_result = df1.add(df2, fill_value=0)
        elif operator == '-':
            df_result = df1.sub(df2, fill_value=0)
    multi_idx_df_multiply_one_col(df_result, operator='/')
    #  restore idx order
    df_result = df_result.reindex(month_abbr_idx, level='month')
    cols = list(df_result.drop(('SUMMARY', "Ignored_wafers"), axis=1))
    df_result.dropna(subset=cols, inplace=True)
    return df_result


def df_monthly_merge_or_exclude(tup, operator='+'):
    add_or_sub_list = []
    for df in tup:
        df = df.copy()  # don't change the original tuple element
        # print(df.to_string())
        multi_idx_df_multiply_one_col(df)
        add_or_sub_list.append(df)
    df_result = reduce(lambda df1, df2: monthly_df_add_or_sub(df1, df2, operator=operator), add_or_sub_list)
    # print(df_result.to_string())
    return df_result


def last_two_row_difference(df, ignore_col=('SUMMARY', "Ignored_wafers")):
    """

    :param df: like df_month = get_monthly_data('AP196')
    :param ignore_col: ignore_col
    :return:df like
            SUMMARY
            Q'ty   Yield Yield_loss        CP       BSV       FSV Passivation crack   ...
    month                                                                              ...
    May     -50.0  0.0004    -0.0004 -0.002075  0.002931 -0.001256               0.0  ...
    """
    df_month = df.copy()
    # drop_unnecessary_cols_and_set_idx(df_month)
    if len(df_month) > 1:
        if ignore_col in df_month:
            df_month = df_month.drop(columns=ignore_col)
        last_two_row_diff = df_month.dropna().tail(2).diff().dropna()
    else:
        last_two_row_diff = pd.DataFrame()
    return last_two_row_diff


def df_operation_parser(idx, parts, title, specific_month, specific_year, monthly_dfs, wafer_list_dict, op,
                        key='df_multi'):
    if idx > 1:  # for UI input exclude Lots
        parts = parts.copy()
        i = parts.last_valid_index()
        parts.at[i, ('DATETIME', 'month')] = specific_month
        parts.at[i, ('DATETIME', 'shipping_date')] = str(specific_year)
        df_o = monthly_dfs[title, 'df_multi']
        df_o.fillna({('SUMMARY', 'Ignored_wafers'): ''}, inplace=True)
        for idx1, i1 in enumerate(df_o.index.values):
            y = i1[0]
            m = i1[1]
            if parts.at[0, ('DATETIME', 'month')] == m and parts.at[0, ('DATETIME', 'shipping_date')] == y:
                continue
            parts.at[idx1 + 1, ('DATETIME', 'month')] = m  # month
            parts.at[idx1 + 1, ('DATETIME', 'shipping_date')] = y  # year
            parts.at[idx1 + 1, ('SUMMARY', 'Ignored_wafers')] = ''
        drop_unnecessary_cols_and_set_idx(parts, drop_year=False)
        dfs_tup = tuple(
            [monthly_dfs[title, 'df_multi'], parts])  # parts is df for UI input exclude Lots, title is the part
    else:
        dfs_tup = tuple(
            monthly_dfs[part, key] for part in parts if not monthly_dfs[part, key].empty)
    tuplen = len(dfs_tup)
    if tuplen == 0:
        # columns = monthly_dfs[title, 'df'].columns
        # return pd.DataFrame(columns=columns)
        return pd.DataFrame()
    elif tuplen == 1:
        df = dfs_tup[0]
    else:
        df = df_monthly_merge_or_exclude(dfs_tup, operator=op)
        if idx > 1:  # for Ui input ignore lots
            ignored_wafers_list = wafer_list_dict[title]
            write_monthly_data(title, specific_year, specific_month, ignored_wafers_list, df)
            monthly_dfs[title, 'df_multi'] = df
    return df


def monthly_data_arrangement(specific_year=None, specific_month=None, sub_df_dict=None, wafer_list_dict=None,
                             compare_year=False):
    """

    :param specific_year: 2018...specific year, int
    :param specific_month: 'Jan'...specific month, str
    :param sub_df_dict: dict like {modifying_part:minus_df, ...}
    :param wafer_list_dict: dict like {modifying_part: wafer_id of ignored_wafers, ...}
    :param compare_year: the comparison reference for the monthly_report, default is comparing to the last month,
                         set True to compare with last whole year
    :return: dict like {monthly_title: monthly_df, ....}
    """
    if sub_df_dict is None:
        ignored_wafers = []
    else:
        ignored_wafers = [(sub_df_dict, '-')]  # ignored_wafers is dict like {modifying_part:sub_df, ...}
        # morphed into list
    path = Tg.read_data_root_folder + '\\Monthly_Report'
    if specific_year is None:
        year_name = sorted([name for name in os.listdir(path) if os.path.isdir(path + '\\' + name)], reverse=True)[0]
        this_year = int(year_name.split('_')[0])
    else:
        this_year = specific_year
    titles = []
    modify_list = [(monthly_merge_parts, '+'),
                   (monthly_exclude_parts, '-')] + ignored_wafers
    monthly_dfs = Tg.NestedDict()
    for part in monthly_use_parts:
        df = get_monthly_data(part, specific_year=specific_year, specific_month=specific_month)
        if compare_year:
            df_last_year = get_monthly_data(part, specific_year=this_year - 1)
            if df_last_year.empty and part == 'AP17407':
                df_last_year = get_monthly_data('AP17406', specific_year=this_year - 1)
            monthly_dfs[part, 'df_multi_last_year'] = df_last_year
            monthly_dfs[part, 'df_last_year'] = df_last_year.copy()

        # # add annual yield report
        # if specific_month == 'Dec':
        #     df_this_year = get_monthly_data(part, specific_year=this_year)
        #     monthly_dfs[part, 'df_multi_this_year'] = df_this_year
        #     monthly_dfs[part, 'df_this_year'] = df_this_year.copy()

        df2 = df if df.empty else df.reset_index(level='year', drop=True)
        month = '' if df2.empty else df2.last_valid_index()
        monthly_dfs[part, 'df'] = df2
        monthly_dfs[part, 'df_multi'] = df
        monthly_dfs[part, 'last_month'] = month
    for idx, tup in reversed(list(enumerate(
            modify_list))):  # reverse is for month ignore wafers also be account by merge_parts and exclude_parts
        op_dict, op = tup
        for title, parts in op_dict.items():
            df3 = df_operation_parser(idx, parts, title, specific_month, specific_year, monthly_dfs, wafer_list_dict,
                                      op)

            if specific_month is None and specific_year is None:
                monthly_dfs[title, 'df'] = df3
            elif df3.index.isin([(str(specific_year), specific_month)]).any():
                monthly_dfs[title, 'df'] = df3
            else:
                monthly_dfs[title, 'df'] = pd.DataFrame()

            if compare_year:
                monthly_dfs[title, 'df_last_year'] = df_operation_parser(idx, parts, title, specific_month,
                                                                         specific_year, monthly_dfs, wafer_list_dict,
                                                                         op, key='df_multi_last_year')
            # # add annual yield report
            # if specific_month == 'Dec':
            #     monthly_dfs[title, 'df_this_year'] = df_operation_parser(idx, parts, title, specific_month,
            #                                                              specific_year, monthly_dfs, wafer_list_dict,
            #                                                              op, key='df_multi_this_year')
            titles.append(title)
    for title in titles:
        df = monthly_dfs[title, 'df']
        if not df.empty:
            df3 = df.reset_index(level='year', drop=True)
            monthly_dfs[title, 'df'] = df3
            monthly_dfs[title, 'last_month'] = df3.last_valid_index()
        else:
            monthly_dfs[title, 'last_month'] = 'No data'
        if compare_year:
            df = monthly_dfs[title, 'df_last_year']
            if not df.empty:
                df = df.reset_index(level='year', drop=True)
            monthly_dfs[title, 'df_last_year'] = df

    return monthly_dfs.dict


def main():
    a = monthly_data_arrangement(compare_year=True)
    pass


if __name__ == '__main__':
    main()
