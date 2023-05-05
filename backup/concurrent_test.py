#!/usr/bin/env python
# _*_ coding:utf-8 _*_
from multi_thread import ConcurrentProcessesGen, ConcurrentThreadsGen
from Database_query import con, connection
from WeekTime_parser import chk_cross_year, chk_cross_month, last_week, next_week, today, week_check_day, \
    chk_every_2_quarter, time_header
from Weekly_update import weekly_shipping_information, weekly_updater, weekly_summary_writer, summary_table_writer, \
    scrap_table_writer
from Document_udpate import report_gen
from Presentation_update import monthly_report_gen
from Reference_update import bsl_update
from df_to_excel import df_update_xlsx
import timeit
import Template_generator as Tg
import datetime as dt
import gc
import pandas as pd
import os


def test_worker_num_for_concurrent_task(max_worker_num, function, var_list, func_args, process=True):
    # print('worker_num, process time, vars_len')
    worker_list, elapsed_list, var_len_list = ([] for t in range(3))
    cache_read, cache_dump = Tg.read_data_root_folder, Tg.dump_data_root_folder
    for n in range(max_worker_num):
        print('worker number {} processing'.format(n + 1))
        if n > 0:
            Tg.read_data_root_folder, Tg.dump_data_root_folder = cache_read + '\\worker_num_{}'.format(
                n + 1), cache_dump + '\\worker_num_{}'.format(n + 1)

        start_time = timeit.default_timer()
        if process:
            multi = ConcurrentProcessesGen(function, var_list, worker_num=n + 1, func_args=func_args)
        else:
            multi = ConcurrentThreadsGen(function, var_list, worker_num=n + 1, func_args=func_args)
        result = multi.run()
        elapsed = timeit.default_timer() - start_time
        worker_list.append(n + 1)
        elapsed_list.append(elapsed)
        var_len_list.append(len(var_list))
        # print('{}, {}, {}'.format(n + 1, elapsed, len(var_list)))
    else:
        Tg.read_data_root_folder, Tg.dump_data_root_folder = cache_read, cache_dump
    return {'worker num': worker_list, 'elapsed time': elapsed_list, 'variables len': var_len_list}


def concurrent_thread_run(function, var_list, worker_num, func_args):
    multi = ConcurrentThreadsGen(function, var_list, worker_num, func_args)
    result = multi.run()
    return result

def write_test_data_to_excel(write_date, purpose, max_worker_num, function, var_list, func_args):
    search_path = '{}\\test_data\\{}\\'.format(Tg.read_data_root_folder, purpose)
    Tg.chk_dir(search_path)
    filename = search_path + 'data.xlsx'
    new_file = not os.path.isfile(filename)
    # dict_thread = test_worker_num_for_concurrent_task(max_worker_num, function, var_list, func_args, process=False)
    for p in [False]:  # cant use multiprocess because always re import
        dict_process = test_worker_num_for_concurrent_task(max_worker_num, function, var_list, func_args, process=p)
        df = pd.DataFrame(dict_process)
        tag = time_header(write_date, start_date=Tg.current_start_date, multi_index=False)
        tag = pd.concat([tag] * len(df.index), ignore_index=True)
        df_result = pd.concat([tag, df], axis=1)
        filename = df_update_xlsx(df_result, filename, sheetname='multiprocess' if p else 'multithreads',
                                  header=new_file, truncate_sheet=new_file, fit_col_width=False, return_writer=True)
    filename.save()


def test_update_all_report_until(target_date=None, end_date=None, comparing_to_the_last_year=False, start_date=None,
                                 force_month_summary=False, debug_parts=None, max_worker_num=3):
    connection()
    if target_date is None:
        target_date = dt.date(year=2018, month=12, day=28)
    if end_date is None or end_date >= today:
        end_date = today
    if start_date is None:
        start_date = last_week(target_date)
    else:
        Tg.current_start_date = start_date
    while True:
        start_time = timeit.default_timer()
        chk_date = week_check_day(target_date)
        ################################################################################################################
        cross_year = chk_cross_year(chk_date)
        cross_month = chk_cross_month(chk_date)
        next_week_chk_date = next_week(chk_date)
        next_week_cross_month = chk_cross_month(next_week_chk_date)
        next_week_cross_year = chk_cross_year(next_week_chk_date)
        week_info = weekly_shipping_information(target_date, start_date=start_date)
        Tg.current_date = target_date
        # updated monthly summary for cross month part and cross year part that have no lots when cross week and year
        # if cross_year or cross_month:
        #     parts_list = active_parts
        # else:
        parts_list = week_info['parts']
        print('This week product parts: {}'.format(parts_list))
        if debug_parts is not None:
            parts_list = debug_parts

        # for part in parts_list:  # TODO: add multi-processes function
        #     print('updating {} weekly data'.format(part))
        #     weekly_updater(part, target_date, week_info, cross_year, cross_month)
        print('updating weekly data')
        write_test_data_to_excel(target_date, 'weekly_content_update', max_worker_num, weekly_updater,
                                 parts_list,
                                 tuple((target_date, week_info, cross_year, cross_month)))

        print('summarizing data')

        for index, part in enumerate(week_info['parts']):
            summary_writer = weekly_summary_writer(part, date=target_date, writer=summary_writer if index else None,
                                                   force_month_summary=force_month_summary)
        else:
            summary_writer = summary_table_writer(writer=summary_writer)
            summary_writer.save()
            scrap_table_writer(target_date)
            print('generating weekly document')
            report_gen(target_date, Tg.summary_cache['week_summary_table'])
            Tg.summary_cache.clear()
        # import warnings
        #
        # warnings.filterwarnings("error")
        if next_week_cross_month:
            # for part in Tg.active_parts:  # TODO: add multi-processes function
            #     weekly_updater(part, next_week(target_date), None, cross_year=next_week_cross_year, cross_month=True)
            write_test_data_to_excel(next_week(target_date), 'monthly_content_update', max_worker_num,
                                     weekly_updater,
                                     Tg.active_parts,
                                     tuple((next_week(target_date), None, next_week_cross_year, True)))

            print('generating monthly presentation')
            monthly_report_gen(chk_date, force_month_summary=force_month_summary,
                               comparing_to_the_last_year=comparing_to_the_last_year)
            # print('monthly report generate')
            if chk_every_2_quarter(next_week_chk_date):
                print('updating baseline')
                bsl_update(next_week_chk_date)
                Tg.update_ref_dict()
        gc.collect()  # try to release some memory
        ################################################################################################################
        elapsed = timeit.default_timer() - start_time  # try to measure the time for each loop
        finish_line = '{} to {}, elapsed time:{} seconds'.format(start_date, target_date, elapsed)
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
    con.close()


def main():
    # test_update_all_report_until(target_date=dt.date(2015, 8, 29), max_worker_num=6)
    test_update_all_report_until(target_date=dt.date(2019, 2, 2), max_worker_num=6)


if __name__ == '__main__':
    # main()
    main()
