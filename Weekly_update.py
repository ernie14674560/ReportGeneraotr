#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import calendar
import datetime as dt
import os
import timeit

import numpy as np
import pandas as pd
from filelock import Timeout, FileLock
from xlrd.biffh import XLRDError

import Template_generator as Tg
from Database_query import connection, scrap_query_parts, inline_prompts_dict
from Document_udpate import report_gen
from Presentation_update import monthly_report_gen
from Reference_update import bsl_update
from WeekTime_parser import chk_cross_year, chk_cross_month, last_week, next_week, today, week_number, last_month, \
    week_check_day, chk_every_2_quarter
from Weekly_parser import df_weekly_by_part, weekly_shipping_information, query_product_group, get_previous_data, \
    df_scrap_table, verify_xlsx_and_get_sheet_details, xlsx_sheet_to_df
from chart import ChartIOGen
from df_to_excel import df_update_xlsx, cell_week_content, cell_week_summary_table
from excel_styler import WeeklyHeaderStyles, bsl_highlight
from multi_thread import concurrent_thread_run

# import warnings
#
# warnings.filterwarnings("error")
month_abbr = list(calendar.month_abbr)[1:]
month_abbr.append('Baseline')


def df_writer(df, part, filename, sheetname, new_file=True, freeze=None, write_bsl=True, **to_df_update_kwargs
              ):  # TODO: reduce writer.save usage
    """
    :param df:df_write
    :param part: part
    :param filename: target filename
    :param sheetname: target sheetname
    :param new_file: new file or not
    :param freeze: default None, the given str are cell like 'A2'.. etc,
                freeze rows above the given cell and columns to the left.
    :param write_bsl: write baseline or not
    :return: update excel or writer with df and df-BSL
    """
    sub_bsl_dict = Tg.df_sub_bsl(df, part)
    df_write = sub_bsl_dict['df_write']
    df_ref = sub_bsl_dict['df_ref']
    filename = df_update_xlsx(df, filename, sheetname, styler=bsl_highlight, apply_axis=None,
                              apply_kwargs={'part': part, 'df_ref': df_ref},
                              multindex_headers=2, cell_width_height=cell_week_content, fit_col_width=False,
                              header=new_file,
                              truncate_sheet=new_file, header_style=WeeklyHeaderStyles if new_file else None,
                              return_writer=True, freeze=freeze, **to_df_update_kwargs)
    if not df_write.empty and write_bsl:
        filename = df_update_xlsx(df_write, filename, sheetname + '-BSL', multindex_headers=2,
                                  cell_width_height=cell_week_content,
                                  fit_col_width=False, header=new_file, truncate_sheet=new_file,
                                  header_style=WeeklyHeaderStyles if new_file else None, return_writer=True,
                                  freeze=freeze, **to_df_update_kwargs)
    return filename


# plots = {'Yield Loss Box Plot': ('box_plot', None), 'Weekly Yield Trend': ('yield', None),
#          'Weekly Final Yield Loss Breakdown': ('yield_loss', None),
#          'FSV Yield Loss Trend': ('stacked_bar_chart', 'FSV'), 'BSV Yield Loss Trend': ('stacked_bar_chart', 'BSV'),
#          'CP Yield Loss Trend': ('stacked_bar_chart', 'CP'), 'FSV Yield Loss Pareto': ('pareto', 'FSV'),
#          'BSV Yield Loss Pareto': ('pareto', 'BSV')}


class NoMonthBefore(Exception):
    pass


def weekly_summary_writer(part, date, writer=None, force_month_summary=False):
    plots = {'Yield Loss Box Plot': ('box_plot', None), 'Weekly Yield Trend': ('yield', None),
             'Weekly Final Yield Loss Breakdown': ('yield_loss', None),
             'FSV Yield Loss Trend': ('stacked_bar_chart', 'FSV'), 'BSV Yield Loss Trend': ('stacked_bar_chart', 'BSV'),
             'CP Yield Loss Trend': ('stacked_bar_chart', 'CP'), 'FSV Yield Loss Pareto': ('pareto', 'FSV'),
             'BSV Yield Loss Pareto': ('pareto', 'BSV')}

    if writer is None:
        # week_check_day is to prevent some error occur in leap year (year that have W53)
        path = Tg.dump_data_root_folder + '\\Weekly_Report\\{}_data'.format(week_check_day(date).year)
        Tg.chk_dir(path)
        file = '\\{}.xlsx'.format(week_number(week_check_day(date)))
        filename = path + file
    sheet_name = part + '_Weekly'
    df = get_previous_data(part)
    # print(part, date)
    part_have_bsl = part in Tg.df_ref_dict
    # summary_cache for write to docx summary table#####################################################################
    Tg.summary_cache[part, 'DataFrame'] = df

    # for u in ["Q'ty", 'Yield']:
    #     summary_cache[part, u] = df.at[df.last_valid_index(), ('SUMMARY', u)]
    week_yield_idx = df.last_valid_index()
    # if df_ref is not None:
    # print(df)
    if part_have_bsl:
        for side in ['FSV', 'BSV', 'CP']:
            Tg.summary_cache[part, 'baseline', side] = df.at[week_yield_idx, ('SUMMARY', side)]
            ser_week = df.loc[week_yield_idx - 1, side]
            ser_bsl = df.loc[week_yield_idx, side]
            ser_week_sub_bsl = ser_week - ser_bsl
            ser_bsl_highlight = ser_week_sub_bsl.loc[lambda x: x > Tg.bsl_tolerance]
            Tg.summary_cache[part, 'greater_than_bsl_highlight', side] = ser_bsl_highlight.to_dict()
            # for i in ser_bsl_highlight.index:
            #     Tg.summary_cache[part, 'out_bsl_highlight', n, i] = ser_bsl_highlight.at[i]

        week_yield_idx -= 1
    # else:,
    # bsl_yield = np.nan
    for u in ["Q'ty", 'Yield', 'FSV', 'BSV', 'CP']:
        Tg.summary_cache[part, u] = df.at[week_yield_idx, ('SUMMARY', u)]
    # Tg.summary_cache[part, 'Baseline yield'] = bsl_yield
    lastmonth = last_month(date)
    try:
        if df.isnull().values.any():
            df_null = df[df.isnull().any(axis=1)]
            # select last index or second-last index if part_have_bsl
            if len(df_null.index) == 1 and part_have_bsl:
                raise NoMonthBefore
            i = df_null.index[df_null.count(1) > 0][-2 if part_have_bsl else -1]
            month = df.at[i, ('DATETIME', "week")]
            if month == lastmonth:
                month_yield = df.at[i, ('SUMMARY', 'Yield')]
                month_qty = df.at[i, ('SUMMARY', "Q'ty")]
            else:
                raise NoMonthBefore
        else:
            raise NoMonthBefore
    except NoMonthBefore:
        month_yield, month_qty = (np.nan for l in range(2))
        month = lastmonth
    Tg.summary_cache[part, 'Last month yield'] = month_yield
    Tg.summary_cache[part, 'Last month qty'] = month_qty
    Tg.summary_cache[part, 'Last month'] = month
    # summary_cache for write to docx summary table#####################################################################
    if force_month_summary:
        if part_have_bsl:
            df_head = df[:week_yield_idx + 1]
            df_tail = df[week_yield_idx + 1:]
            df_month_summary = Tg.df_monthly_summary(df_head)
            df = pd.concat([df_head, df_month_summary, df_tail], ignore_index=True)
        else:
            df_month_summary = Tg.df_monthly_summary(df)
            df = pd.concat([df, df_month_summary], ignore_index=True)
    writer = df_writer(df, part, writer if writer is not None else filename, sheet_name, write_bsl=False)
    startrow = writer.book[sheet_name].max_row + 1
    worksheet = writer.sheets[sheet_name]
    chart = ChartIOGen(df, part=part)
    i = 0
    for name, tup in plots.items():
        kind, site = tup
        image = chart.weekly_defect_data(kind, site)
        worksheet.add_image(image, 'A' + str(startrow))
        if i:
            startrow += 21
        else:
            startrow += 84
        i += 1
        Tg.summary_cache[part, 'chart', name] = image
    return writer


def weekly_writer(part, filename, date, product_group, shipping_information, cross_year, cross_month, weekly=True,
                  return_df=False, year=None, ignore_db_error=False):
    writer = None
    new_file = False
    input_writer = False
    sheet_name = part + '_Weekly'
    pcs_group_sheet = part + '_raw_group'
    pcs_code_sheet = part + '_raw_code'
    new_index = 0
    # pcs_index = 0
    try:
        df = xlsx_sheet_to_df(filename, sheet_name=sheet_name, index_col=0, header=[0, 1])
        # the variable new_index and read_excel arg index_col=0 is serve to avoid following bug and issue
        # https://github.com/stephenrauch/pandas/commit/9b37ff94643296d489498138c79bb0244aaa3f79
        # https://github.com/pandas-dev/pandas/pull/23703
        new_index = df.index.max() + 1
        # monthly summary writer section, need to sync between threads to protect monthly_data_log.xlsx file
        ################################################################################################################
        if cross_month:
            path = Tg.dump_data_root_folder + '\\Monthly_Report\\{}_data'.format(year)
            Tg.chk_dir(path)
            monthly_report_filename = path + '\\monthly_data_log.xlsx'
            df_summary = Tg.df_monthly_summary(df)
            # Locker for multi threads synchronization##################################################################
            try:
                lock_path = monthly_report_filename + '.lock'
                locker = FileLock(lock_path)
                with locker.acquire():
                    m_new_file = False
                    try:
                        verify_xlsx_and_get_sheet_details(monthly_report_filename, part)
                    except (FileNotFoundError, XLRDError) as e:  # File or sheet not found
                        m_new_file = True
                    df_monthly_report_log = df_summary.dropna(axis='columns')
                    df_monthly_report_log = df_monthly_report_log.rename(columns={'week': 'month'}, level=1)
                    cols_order = list(df_monthly_report_log)
                    df_monthly_report_log.at[
                        df_monthly_report_log.last_valid_index(), ('SUMMARY', 'Ignored_wafers')] = 'NA'
                    cols_order.insert(1, ('SUMMARY', 'Ignored_wafers'))
                    df_monthly_report_log = df_monthly_report_log.loc[:, cols_order]
                    df_update_xlsx(df_monthly_report_log, monthly_report_filename, sheetname=part, multindex_headers=2,
                                   header=m_new_file, truncate_sheet=m_new_file,
                                   header_style=WeeklyHeaderStyles if m_new_file else None,
                                   cell_width_height=cell_week_content,
                                   fit_col_width=False)
            except Timeout:
                print("Another instance of this application currently holds the lock.")
            ############################################################################################################
            writer = df_writer(df_summary, part, filename, sheet_name, new_file)
            new_index = df.index.max() + 2
            input_writer = True
            if cross_year:
                writer.save()
                input_writer = False
        ################################################################################################################
    except FileNotFoundError:
        new_file = True
    except Tg.NoWeekBefore:  # No week to summarize into month when cross month or year
        pass

    if shipping_information is None:
        pass
    else:
        df_write = df_weekly_by_part(part, date, ignore_db_error, shipping_information=shipping_information)
        df_summary = df_write['df_summary']
        qty = df_summary.at[0, ('SUMMARY', "Q'ty")]
        _yield = df_summary.at[0, ('SUMMARY', 'Yield')]
        print(f'{part} output {int(qty)} pcs; yield is {_yield:.1%}')
        df_pcs_group = df_write['df_pcs_group']
        # df_pcs_code = Tg.df_code_update(df_write['df_pcs_code'], part)  # add description on top of df_code
        df_pcs_code = df_write['df_pcs_code']
        # new_indexes = range(0, 0 + len(df_pcs_group))
        if cross_year or new_file:
            new_file = True
            if weekly:
                path = Tg.dump_data_root_folder + '\\Weekly_Report\\{}_data\\{}'.format(week_check_day(date).year,
                                                                                        product_group)
                Tg.chk_dir(path)
                filename = path + '\\{}.xlsx'.format(part)
            new_index = 0
            # new_indexes = range(0, len(df_pcs_group))
        df_summary.index = [new_index]
        if return_df:
            return df_summary
        # df_pcs_group.index = new_indexes
        # df_pcs_code.index = new_indexes
        writer = df_writer(df_summary, part, writer if input_writer else filename, sheet_name, new_file,
                           freeze='B3' if date is None else 'E3')
        writer = df_writer(df_pcs_group, part, writer, pcs_group_sheet, new_file,
                           freeze='D3' if date is None else 'G3', pcs_group_idx_optimize=True)
        writer = df_update_xlsx(df_pcs_code, writer, pcs_code_sheet, fit_col_width=False, header=new_file,
                                truncate_sheet=new_file, multindex_headers=2, return_writer=True,
                                freeze='D3' if date is None else 'G3', pcs_group_idx_optimize=True)
        if part in Tg.df_custom_raw_code_output_dict:
            df_custom_raw_code = Tg.df_custom_row_code_output(df_pcs_code, part)
            writer = df_update_xlsx(df_custom_raw_code, writer, 'custom_' + pcs_code_sheet, fit_col_width=False,
                                    header=new_file, header_style=WeeklyHeaderStyles if new_file else None,
                                    read_header_style_row=1, mapping_style_dict=Tg.custom_raw_code_style_dict[part],
                                    truncate_sheet=new_file, multindex_headers=2, return_writer=True,
                                    freeze='D3' if date is None else 'G3', pcs_group_idx_optimize=True)

    if writer is not None:
        try:
            writer.save()
        except PermissionError as err:
            path = writer.path
            new_path = path[:-5] + '-' + path[-5:]
            writer.path = new_path
            writer.save()
            Tg.update_all_report_until_error_log.append(
                str(err) + f"\nfile has been saved to {new_path}, please change back to origin file name.")


def weekly_updater(part, date, week_info, cross_year, cross_month):
    """
    dedicated func to generate report in given time interval
    :param part: part
    :param date: date
    :param week_info: dict contain parts, lots df
    :param cross_year: cross year
    :param cross_month: cross month
    :return: None
    """
    try:
        if week_info is None:
            shipping_information = None
        else:
            shipping_information = week_info['lots'].get(part)
        year = week_check_day(last_week(date)).year
        group = query_product_group(part)
        path = Tg.dump_data_root_folder + '\\Weekly_Report\\{}_data\\{}'.format(year, group)
        Tg.chk_dir(path)
        file = '\\{}.xlsx'.format(part)
        # print(f'updating {part} data')
        weekly_writer(part, path + file, date, group, shipping_information, cross_year, cross_month, year=year,
                      ignore_db_error=True)
    except Exception as e:
        Tg.update_all_report_until_error_log.append(str(e))


def summary_table_writer(writer=None):
    sheet_name = 'Weekly_summary_table'
    df = get_previous_data(sheet_name=sheet_name, row_limit=4)
    cols_order = list(df)  # restored col order

    df = df[~df[('Item', 'Part', 'Goal')].isin(month_abbr)]
    df_this_week = Tg.df_summary_table()

    df_bsl_and_month = df_this_week[:2]
    df_append = df_this_week[-1:]

    df = df.append(df_append, ignore_index=True, sort=False)

    df = df_bsl_and_month.append(df, ignore_index=True, sort=False)

    df.fillna('--', inplace=True)

    df = df.loc[:, cols_order]  # restored col order
    Tg.summary_cache['week_summary_table'] = df.T
    writer = df_update_xlsx(df, writer, sheet_name, cell_width_height=cell_week_summary_table, multindex_headers=3,
                            return_writer=True, fit_col_width=False, new_sheet_position=0)
    return writer


def scrap_table_writer(date):
    path = Tg.dump_data_root_folder + '\\Weekly_Report\\{}_data'.format(week_check_day(date).year)
    Tg.chk_dir(path)
    sheetname = 'Weekly_scrap_part'
    sheetname2 = 'Scrap_wafers'
    filename = path + '\\Scrap_wafers.xlsx'
    new_file = not os.path.isfile(filename)
    df, df2 = df_scrap_table(date)
    df_empty = df.empty
    parts = Tg.active_parts.copy()
    for part in scrap_query_parts:
        if part not in parts:
            parts.append(part)
    for part in parts:
        df2[part] = 0 if df_empty else df.loc[df.PARTID.str.contains(part), "Q'ty"].sum()
    filename = df_update_xlsx(df2, filename, sheetname, truncate_sheet=False, return_writer=True, fit_col_width=False,
                              header=new_file)
    if not df_empty:
        filename = df_update_xlsx(df, filename, sheetname2, truncate_sheet=False, return_writer=True, offset=1,
                                  header=new_file)
    filename.save()


def update_all_report_until(target_date=None, end_date=None, comparing_to_the_last_year=False, start_date=None,
                            force_month_summary=False, debug_parts=None):
    connection()
    if target_date is None:
        target_date = dt.date(year=2018, month=12, day=28)
    if end_date is None or end_date >= today:
        end_date = today
    if start_date is None:
        start_date = last_week(target_date)
    else:
        Tg.current_start_date = start_date
    start_time = timeit.default_timer()
    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = end_date.strftime('%Y-%m-%d')
    Tg.update_all_report_until_error_log.clear()
    while True:

        loop_start_time = timeit.default_timer()
        chk_date = week_check_day(target_date)
        ################################################################################################################
        cross_year = chk_cross_year(chk_date)
        cross_month = chk_cross_month(chk_date)
        next_week_chk_date = next_week(chk_date)
        next_week_cross_month = chk_cross_month(next_week_chk_date)
        next_week_cross_year = chk_cross_year(next_week_chk_date)
        week_info = weekly_shipping_information(target_date, start_date=start_date)
        Tg.current_date = target_date

        parts_list = week_info['parts']
        if not parts_list:
            print('This week has no parts output.')
        else:
            parts_list.sort(key=len)
            print('This week production parts passed {}: {}'.format(list(inline_prompts_dict.keys()), parts_list))
            active_parts = []
            if debug_parts is not None:
                parts_list = debug_parts
            for part in parts_list:
                try:
                    query_product_group(part)
                    active_parts.append(part)
                except Exception as e:
                    Tg.update_all_report_until_error_log.append(str(e))
            active_parts.sort(key=len)
            week_info['parts'] = active_parts
            print('This week active parts: {}'.format(active_parts))
            print('updating weekly data')
            concurrent_thread_run(weekly_updater, active_parts,
                                  tuple((target_date, week_info, cross_year, cross_month)),
                                  3)
            print('summarizing data')
            for index, part in enumerate(active_parts):
                summary_writer = weekly_summary_writer(part, date=target_date,
                                                       writer=summary_writer if index else None,
                                                       force_month_summary=force_month_summary)

            summary_writer = summary_table_writer(writer=summary_writer)
            summary_writer.save()
            scrap_table_writer(target_date)
            print('generating weekly document')
            report_gen(target_date, Tg.summary_cache['week_summary_table'])
            Tg.summary_cache.clear()
        if next_week_cross_month:
            print('updating monthly data')
            concurrent_thread_run(weekly_updater, Tg.active_parts,
                                  tuple((next_week(target_date), None, next_week_cross_year, True)), 3)
            print('generating monthly presentation')
            monthly_report_gen(chk_date, force_month_summary=force_month_summary,
                               comparing_to_the_last_year=comparing_to_the_last_year)
            # print('monthly report generate')
            if chk_every_2_quarter(next_week_chk_date):
                print('updating baseline')
                bsl_update(next_week_chk_date)
                Tg.update_ref_dict()
        ################################################################################################################
        loop_elapsed = timeit.default_timer() - loop_start_time  # try to measure the time for each loop
        finish_line = '{} to {}, elapsed time:{} seconds'.format(start_date, target_date, loop_elapsed)
        print(finish_line)
        target_date = next_week(target_date)
        start_date = next_week(start_date)
        Tg.current_start_date = None
        if isinstance(target_date, dt.date):
            if target_date > end_date:
                break
        else:
            if target_date.date > end_date:
                break
    elapsed_seconds = timeit.default_timer() - start_time
    elapsed_time = dt.timedelta(seconds=elapsed_seconds)
    msg = f'Weekly reports from {start_date_str} to {end_date_str} have been generated successfully.     \n' \
        f'Elapsed time: {elapsed_time}'
    if Tg.update_all_report_until_error_log:
        msg += '\nErrors occurred during processing:\n{}'.format('\n'.join(Tg.update_all_report_until_error_log))
    Tg.update_all_report_until_error_log.clear()
    return 'finish', msg


def main():
    # update_all_report_until(target_date=dt.date(2017, 12, 29))
    # next_week(dt.date(year=2019, month=2, day=25))
    # update_all_report_until()
    # update_all_report_until(target_date=dt.date(year=2018, month=2, day=16))
    # update_all_report_until(target_date=dt.date(2019, 7, 12), debug_parts=['AP196'])
    # update_all_report_until(target_date=dt.date(2017, 2, 3))
    # update_all_report_until(target_date=dt.date(2019, 6, 28))
    update_all_report_until(target_date=dt.date(2021, 10, 15), debug_parts=['AP1FG'])
    # update_all_report_until(target_date=dt.date(2017, 2, 3))
    # update_all_report_until(target_date=dt.date(2018, 4, 6))
    # update_all_report_until(target_date=dt.date(2015, 8, 21))
    # update_all_report_until(target_date=dt.date(2015, 12, 11))
    # update_all_report_until(target_date=dt.date(2019, 6, 28))


def main1():
    from chart import ChartIOGen
    from Weekly_parser import get_previous_data
    df = get_previous_data('AP196', bsl=False, row_limit=False)
    chart = ChartIOGen(df, part='AP196')
    image = chart.weekly_defect_data(kind='box_plot', site=None, show=True)


if __name__ == '__main__':
    main()
