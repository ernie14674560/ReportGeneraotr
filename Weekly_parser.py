#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import io
import os
import re
import threading
import zipfile

import pandas as pd
import xmltodict
from cachetools import cached, LFUCache
from xlrd.biffh import XLRDError
from xlsx2csv import Xlsx2csv

import Template_generator as Tg
from Database_query import weekly_lot, fsbs_insp_weekly, cp_weekly, scrap_lot
from WeekTime_parser import time_header, week_interval, short_month, week_number, week_check_day

cache1 = LFUCache(maxsize=20)


class PartNotInGroupError(Exception):
    def __init__(self, part):
        self.value = f"part {part} doesn't belong to any group like SMP, Conti, GEN1, ....etc"

    def __str__(self):
        return self.value


def xlsx_to_string_io(filename, sheet_name, merge_cells=True):
    metadata = io.StringIO()
    sheet_details = verify_xlsx_and_get_sheet_details(filename, sheet_name)
    sheet_id = sheet_details[sheet_name]
    Xlsx2csv(filename, outputencoding="utf-8", merge_cells=merge_cells).convert(metadata, sheetid=int(sheet_id))
    metadata.seek(0)
    return metadata


def xlsx_sheet_to_df(filename, sheet_name, merge_cells=True, *read_args, **read_kwargs):
    """
    pd.read_excel will load entire workbook at beginning, which will waste enormous time and resources if workbook have
    lots of sheet. Use this func to avoid the problem
    :param filename:
    :param sheet_name:
    :param merge_cells:
    :param read_args:
    :param read_kwargs:
    :return: df
    """
    metadata = xlsx_to_string_io(filename, sheet_name, merge_cells=merge_cells)
    # metadata.seek(0)
    df = pd.read_csv(metadata, *read_args, **read_kwargs)
    return df


def sheet_details_dict(sheet_obj, at_sign=True, key='name', value='sheetId'):
    sheet_details = {}
    if at_sign:
        key = '@' + key
        value = '@' + value
    if isinstance(sheet_obj, list):
        for sheet in sheet_obj:
            sheet_details[sheet[key]] = sheet[value]
    else:
        # only one sheet
        sheet_details[sheet_obj[key]] = sheet_obj[value]
    return sheet_details


def get_sheet_details(file_path):
    """
    :param file_path:
    :return: sheet names list
    """
    with zipfile.ZipFile(file_path, 'r') as zip_xlsx:
        with zip_xlsx.open('xl/workbook.xml', 'r') as f:
            xml = f.read()
            dictionary = xmltodict.parse(xml)
            sheet_obj = dictionary['workbook']['sheets']['sheet']
            try:
                sheet_details = sheet_details_dict(sheet_obj)
            except KeyError:
                sheet_details = sheet_details_dict(sheet_obj, at_sign=False)
    return sheet_details


def verify_xlsx_and_get_sheet_details(filepath, sheet_name):
    sheet_details = get_sheet_details(filepath)
    if sheet_name not in sheet_details:
        raise XLRDError
    return sheet_details


@cached(cache1)
def query_product_group(part):
    for i in Tg.active_parts_and_product_group.items():
        s = part[:5]
        lst = i[1]
        for p in lst:
            if s in p:
                group = i[0]
                return group
        # if part[:5] in i[1]:
        #     group = i[0]
        #     break
    else:
        raise PartNotInGroupError(part)


def parts_lots_arrange(df, first_5_char=False):
    actives = set(Tg.active_parts.copy())  # for weekly update
    inputs = set(df['PARTNAME'].tolist())  # for wafer map generate/ part may not belong to active part
    inputs_5char = {part[:5] for part in inputs}
    if actives.intersection(inputs_5char):
        iter_parts = actives.union(inputs_5char)
    else:
        iter_parts = inputs
    if first_5_char:
        # iter_parts = [part for part in iter_parts if len(part) < 6]
        iter_parts = {part[:5] for part in iter_parts}
    else:
        iter_parts = {part[:9] for part in iter_parts}
    lots = {part: df[df.PARTNAME.str.contains(part)].sort_index() for part in iter_parts if
            df.PARTNAME.str.contains(part).any()}
    parts = [*lots]
    return {'lots': lots, "parts": parts}


def weekly_shipping_information(end_date, start_date):
    """
    :param end_date: datetime.date object, must be friday to produce correct weekly report
    :param start_date: datetime.date object, must be friday to produce correct weekly report
    :return: dict contain shipping_lots and shipping_parts information
    shipping_lots: dict like {part: part_shipping_lot df, ....}
     df like            COMPIDS	PARTNAME
            LOTID
        1NIJ089.1	1NIJ089.1	AP17400AH-C4N3-GB
        1NIJ089.10	1NIJ089.3	AP17400AH-C4N3-GB
        1NIJ089.11	1NIJ089.4	AP17400AH-C4N3-GB
        ...         ...         ...
    shipping_part: list contain this end_date week shipping part
    """
    df = weekly_lot(end_date, start_date)
    return parts_lots_arrange(df)


def query_data(wf, func, ignore_db_error):
    if ignore_db_error:
        try:
            data = func(wf)
        except Exception as e:
            Tg.update_all_report_until_error_log.append(str(e))
            data = pd.DataFrame()
    else:
        data = func(wf)
    return data


def query_data_by_func(part, wf, ignore_db_error, parent_part, child_part, func):
    if child_part:
        wait_event = threading.Event()
        timeout_counter = 0
        timeout_flag = Tg.summary_cache.get(('Timeout', part), False)
        err_msg = f"After 60s, still can't find wafer {wf} in parent part {part[:5]} cache, query data by part {part}."
        while True:
            if timeout_flag or timeout_counter > 1200:  # try 60 sec, if still no parent cache, query data
                print(err_msg)
                Tg.summary_cache['Timeout', part] = True
                data = query_data(wf, func, ignore_db_error)
                return data
            data = Tg.summary_cache.get(('QueryingCache', part[:5], wf, func.__name__))
            if data is None:
                timeout_counter += 1
                wait_event.wait(timeout=0.05)
            else:
                break
    else:
        data = query_data(wf, func, ignore_db_error)
        # if ignore_db_error:
        #     try:
        #         data = func(wf)
        #     except Exception as e:
        #         Tg.update_all_report_until_error_log.append(str(e))
        #         data = pd.DataFrame()
        # else:
        #     data = func(wf)
    if parent_part:
        Tg.summary_cache['QueryingCache', part, wf, func.__name__] = data
    return data


def query_wf_data(part, wf, ignore_db_error, child_part, parent_part):
    # get insp data
    insp_data = query_data_by_func(part, wf, ignore_db_error, parent_part, child_part, fsbs_insp_weekly)
    # get cp data
    cp_data = query_data_by_func(part, wf, ignore_db_error, parent_part, child_part, cp_weekly)

    return insp_data, cp_data


def df_weekly_by_part(part, end_date, ignore_db_error, shipping_information=None, reset_index=True):
    """

    :param part:
    :param end_date: datetime.date object, must be friday to produce correct weekly report
    :param ignore_db_error: ignore_db_error, set True when use in query week by week data(call by update_all_until)
    :param shipping_information: called weekly_shipping_information(end_date) object
    :param reset_index: default True, reset LotID index to prevent concat mis-align, maybe will fix in future
    :return:
                      SUMMARY                                  FSV                                    ...
                  Q'ty     Yield  Yield_loss CP  BSV FSV  Passivation crack Scratch Contamination ...
    df like    0  3         0.99   ...


    OR summary = False, update wafer by wafer data


               WAFER           SUMMARY                                  FSV                                    ...
            WAFER_ID PART_ID   Yield  Yield_loss CP  BSV FSV  Passivation crack Scratch Contamination ...
     df like   xx      xxx     0.99   ...
               yy      yyy     0.94   ...
               ...     ...     ...
    """
    result_dict, insp_wafer_dict, cp_wafer_dict = ({} for l in range(3))

    ids = ["WAFER_ID", "PART_ID"]
    active_codes = Tg.part_active_code_string[part]
    active_codes_set = Tg.part_active_code_str_set[part]
    sheet_list_header = ids + active_codes
    correction = {'COMPIDS': 'WAFER_ID', 'PARTNAME': 'PART_ID'}
    lots = shipping_information.rename(columns=correction)
    header = pd.DataFrame(columns=sheet_list_header)
    df_weekly = pd.concat([lots, header], sort=False).fillna(0)
    if reset_index:
        df_weekly.reset_index(inplace=True, drop=True)
    wafers = df_weekly.WAFER_ID

    if ignore_db_error:
        parent_part = True if part in Tg.parts_parents_to_child else False
        child_part = True if part in Tg.parts_parents_to_child.get(part[:5], []) else False
    else:
        parent_part, child_part = False, False

    for i, wf in wafers.iteritems():
        insp_data, cp_data = query_wf_data(part, wf, ignore_db_error, child_part, parent_part)
        insp_wafer_dict[wf] = insp_data
        cp_wafer_dict[wf] = cp_data

    for i, wf in wafers.iteritems():
        # get insp, cp data
        insp_data = insp_wafer_dict[wf]
        cp_data = cp_wafer_dict[wf]

        for cat, cat_count in insp_data.itertuples():
            if cat in active_codes_set:
                df_weekly.at[i, cat] = cat_count
            else:
                # if part == 'AP174':  # AP174 dont have 'Others' defect, discard code
                #     continue
                try:
                    backside_code = cat[2]
                except IndexError:  # AP197 wf 1NHE364.11 has one fail '52' defect code = =
                    backside_code = '{0:0<4}'.format(cat)  # fill '0' after missing code
                if backside_code == '9' or cat in {'6051', '7000'}:
                    cat = '8888'  # 'BS others' group
                else:
                    cat = '8887'  # 'FS others' group

                df_weekly.at[i, cat] = cat_count
                # pass  # input code do not belong in active_codes by part TODO:add fix code function
        for bin_code, bin_count in cp_data.itertuples():
            if bin_code in active_codes_set:
                if bin_code != '1':  # pass die

                    # if bin_code == '10':
                    #     df_weekly.at[i, '9'] = bin_count
                    # else:
                    #     df_weekly.at[i, bin_code] = bin_count
                    df_weekly.at[i, bin_code] = bin_count
            else:
                pass  # input code do not belong in active_codes by part TODO:add fix code function

    # result_dict['df_pcs_code'] = Tg.df_code_update(df_weekly, part)
    df_dict = Tg.df_weekly_update(df_weekly, part)
    df_dict['df_pcs_code'] = Tg.df_code_update(df_weekly, part)
    for name, summary in [('df_summary', True), ('df_pcs_group', False), ('df_pcs_code', False)]:
        df = df_dict[name]
        tag = time_header(end_date, start_date=Tg.current_start_date)
        if not summary:
            tag = pd.concat([tag] * len(df.index), ignore_index=True)
        if tag.empty:
            result = df
        else:
            result = pd.concat([tag, df], axis=1)
        result_dict[name] = result

    return result_dict


def prepend_df_break_it(concat_list, df):
    concat_list.insert(0, df)
    raise Tg.BreakIt


def get_previous_data(part="", sheet_name=None, row_limit=24, search_week=True, ignore_index=True, bsl=True,
                      monthly=False, search_quarter=False, regex_pattern=r'.*W\d\d\.xlsx', specific_year=None):
    """
    :param part:part
    :param sheet_name: concat sheet name in part_weekly.xlsx
    :param row_limit: default 24, set false to turn off limit
    :param search_week: search_week
    :param ignore_index: ignore_index when concat or not
    :param bsl: default True, determine whether add baseline at the bottom of dataframe or not
    :param monthly: default False, search part monthly data in a year, have to turn bsl off
    :param regex_pattern: regex_pattern for search file if no part input
    :param search_quarter: search this year quarter month
    :param specific_year: specify one year to search
    :return: df contain concatenated sheet
    if no part input, will search W01~ WXX xlsx file, need to specify row_limit and sheet_name
    """
    if monthly:
        path = Tg.read_data_root_folder + '\\Monthly_Report'
    else:
        path = Tg.read_data_root_folder + '\\Weekly_Report'
    if specific_year is None:
        year_folders = sorted([name for name in os.listdir(path) if os.path.isdir(path + '\\' + name)], reverse=True)
    else:
        name = '{}_data'.format(str(specific_year))
        year_folders = [name]
    df = pd.DataFrame()  # if search everything and still don't have data, return empty DataFrame
    # year = str()
    if part:
        product_group = query_product_group(part)
        df_ref = Tg.df_ref_dict.get(part)
        # add baseline data at bottom of df if possible
        if part in Tg.bsl_parts and df_ref is not None and bsl:
            df_result = df_ref.tail(1).copy()
            i = df_result.last_valid_index()  # use last updated BSL
            df_result.at[i, ('DATETIME', 'shipping_date')] = df_result.last_valid_index()
            df_result.at[i, ('DATETIME', 'week')] = 'Baseline'
        else:
            df_result = pd.DataFrame()
    else:
        df_result = pd.DataFrame()
    concat_list = [df_result]

    if sheet_name is None:
        sheet_name = part + '_Weekly'
    try:
        for year in year_folders:
            try:
                if part:
                    # if file not found
                    # raise FileNotFoundError
                    if not monthly:
                        search_path = '{}\\{}\\{}\\{}.xlsx'.format(path, year, product_group, part)
                        df = xlsx_sheet_to_df(search_path, sheet_name=sheet_name, index_col=0, header=[0, 1])
                    else:
                        #     search_path = '{}\\{}\\monthly_data_log.xlsx'.format(
                        #         Tg.read_data_root_folder + '\\Monthly_Report', year)
                        search_path = '{}\\{}\\monthly_data_log.xlsx'.format(path, year)
                        try:
                            df = xlsx_sheet_to_df(search_path, sheet_name=part, index_col=0, header=[0, 1])
                        except (FileNotFoundError, XLRDError) as e:  # File or sheet not found
                            print(e)
                            raise Tg.BreakIt
                        # df = df[df.isnull().any(1)]
                        df.loc[:, ('DATETIME', 'shipping_date')] = year.split('_')[0]
                        if search_quarter:
                            prepend_df_break_it(concat_list, df)
                        # only have one month data or no data
                        elif len(df) < 2:
                            pass  # search last year
                        else:
                            prepend_df_break_it(concat_list, df)
                    # df_result = pd.concat([df, df_result], ignore_index=True, sort=False)
                    # concat will auto sort column
                    concat_list.insert(0, df)
                    total_row = sum([len(dfs) for dfs in concat_list])
                    if row_limit and total_row > row_limit:
                        # df_result = df_result.tail(row_limit)
                        raise Tg.BreakIt
                        # break

                else:
                    weeks_path = path + '\\' + year
                    weeks = sorted([name for name in os.listdir(weeks_path) if
                                    os.path.isfile(weeks_path + '\\' + name) and re.match(regex_pattern, name)],
                                   reverse=True)
                    if not weeks:
                        continue  # to next search year loop
                    for week in weeks:
                        search_path = '{}\\{}'.format(weeks_path, week)
                        if search_week:  # search W01, W02.xlsx like file
                            df = xlsx_sheet_to_df(search_path, sheet_name=sheet_name, index_col=0,
                                                  header=[0, 1, 2], na_values='--').tail(1)
                        else:  # search scrap table and immediately return
                            df = xlsx_sheet_to_df(search_path, sheet_name=sheet_name, index_col=0, header=[0])
                        # df_result = pd.concat([df, df_result], ignore_index=ignore_index, sort=False)
                        # concat will auto sort column
                        concat_list.insert(0, df)
                        total_row = sum([len(dfs) for dfs in concat_list])
                        if row_limit and total_row > row_limit:
                            # df_result = df_result.tail(row_limit)
                            # break
                            raise Tg.BreakIt
                    else:
                        continue  # search next year
                    # break
                    # raise FileNotFoundError
            except FileNotFoundError:
                continue
        else:  # after search everything and still can't meet row_limit, jump out for-loop
            if part:
                raise Tg.BreakIt
            else:
                if search_week:
                    raise Tg.NoWeekBefore
                else:  # search scrap table and immediately return
                    raise Tg.BreakIt
        # cols_order = list(df)
        # df_result = df_result.loc[:, cols_order]  # restore column order
    except Tg.BreakIt:
        cols_order = list(df)
        df_result = pd.concat(concat_list, ignore_index=ignore_index, sort=False)
        if row_limit:
            df_result = df_result.tail(row_limit)
        df_result = df_result.loc[:, cols_order]  # restore column order
    except Tg.NoWeekBefore:
        # initialize docx summary table
        df_result = Tg.df_summary_table(init=True)
    if ignore_index:
        df_result.reset_index(drop=True, inplace=True)
    # if monthly:
    #     df_result.index.rename(year, inplace=True)
    return df_result


def get_week_ooc():
    """
    get this week ooc data
    :return: dataframe contain ooc parts information
    """
    result = pd.DataFrame()
    date = Tg.current_date
    for part in Tg.ooc_parts_show_in_doc:
        df_ = Tg.summary_cache.get((part, 'DataFrame'))
        if df_ is None:
            df = pd.DataFrame(
                {'month': [short_month(week_check_day(date))], 'shipping_date': [week_interval(date)],
                 'week': [week_number(week_check_day(date))],
                 'wafers': ['Normal'], 'groups': ['Normal'], 'counts': [0], 'pcs': [0], "Q'ty": [0]}, index=[part])
            df.index.name = 'PART'
        else:
            df = df_.dropna()
            df = df.tail(1)
            qty = df['SUMMARY']["Q'ty"]
            df = pd.concat([df['DATETIME'], df['OOC']], axis=1, sort=False)
            df["Q'ty"] = qty
            df['PART'] = pd.Series(part, index=df.index)
            df.set_index('PART', inplace=True)
        result = pd.concat([result, df], sort=False)
    return result


def df_scrap_table(target_date):
    """
    :param target_date: datetime object
    :return: scrap lot information df and week index df contain total_scrap_pcs
    """
    df = scrap_lot(end_date=target_date)
    tag = time_header(target_date, multi_index=False, start_date=Tg.current_start_date)
    if df.empty:
        result = df
        tag['total_scrap_pcs'] = 0
    else:
        total_scrap_pcs = df["Q'ty"].sum()
        multiple_tag = pd.concat([tag] * len(df.index), ignore_index=True)
        multiple_tag.set_index(df.index, inplace=True)
        result = pd.concat([multiple_tag, df], axis=1)
        tag['total_scrap_pcs'] = total_scrap_pcs
    tag.set_index('week', inplace=True)
    return result, tag


def ser_previous_scrap_table():
    """
    :return: series of scrap pcs docx table
    """
    df = get_previous_data(sheet_name='Weekly_scrap_part', row_limit=4, regex_pattern=r'Scrap_wafers\.xlsx',
                           search_week=False, ignore_index=False)
    ser_docx_scrap_pcs = df[Tg.docx_scrap_pcs_parts].sum(axis=1)
    return ser_docx_scrap_pcs


def main():
    part = 'AP196'
    # end_date = dt.date(2019, 1, 4)
    # c = df_weekly_by_part(part, end_date, summary=False)
    # d = df_weekly_by_part(part, end_date, summary=False, reset_index=True)
    a = get_previous_data('AP196')
    return a


def main2():
    pass


if __name__ == '__main__':
    main()

    #
    # def query_wf_data(part, wf, ignore_db_error, child_part, parent_part):
    #     # get insp data
    #     if child_part:
    #         insp_data = Tg.summary_cache.get(('QueryingCache', part[:5], wf, 'insp'), fsbs_insp_weekly, wf)
    #     else:
    #         if ignore_db_error:
    #             try:
    #                 insp_data = fsbs_insp_weekly(wf)
    #             except Exception as e:
    #                 Tg.update_all_report_until_error_log.append(str(e))
    #                 insp_data = pd.DataFrame()
    #         else:
    #             insp_data = fsbs_insp_weekly(wf)
    #     if parent_part:
    #         Tg.summary_cache['QueryingCache', part, wf, 'insp'] = insp_data
    #     # get cp data
    #     if child_part:
    #         cp_data = Tg.summary_cache.get(('QueryingCache', part[:5], wf, 'cp'), cp, wf)
    #     else:
    #         if ignore_db_error:
    #             try:
    #                 cp_data = cp(wf)
    #             except Exception as e:
    #                 Tg.update_all_report_until_error_log.append(str(e))
    #                 cp_data = pd.DataFrame()
    #         else:
    #             cp_data = cp(wf)
    #     if parent_part:
    #         Tg.summary_cache['QueryingCache', part, wf, 'cp'] = cp_data
    #     return insp_data, cp_data
