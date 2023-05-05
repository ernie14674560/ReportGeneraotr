#!/usr/bin/env python
# _*_ coding:utf-8 _*

import pandas as pd
import numpy as np
import warnings
import os
from string import ascii_uppercase
from WeekTime_parser import last_month, week_number, week_check_day
from cachetools import cached, LFUCache
from cachetools.keys import hashkey
from functools import partial
from Configuration import cfg
from threading import RLock


class NestedDict:
    """
    Dictionary that can use a list to get a value

    :example:
    >> nested_dict = NestedDict({'aggs': {'aggs': {'field_I_want': 'value_I_want'}, 'None': None}})
    >> path = ['aggs', 'aggs', 'field_I_want']
    >> nested_dict[path]
    'value_I_want'
    >> nested_dict[path] = 'changed'
    >> nested_dict[path]
    'changed'
    """

    def __init__(self, *args, **kwargs):
        self.dict = dict(*args, **kwargs)

    def __getitem__(self, keys):
        # Allows getting top-level branch when a single key was provided
        if not isinstance(keys, tuple):
            keys = (keys,)

        branch = self.dict
        if not isinstance(branch, dict):
            raise KeyError
        for key in keys:
            branch = branch[key]

        # If we return a branch, and not a leaf value, we wrap it into a NestedDict
        return NestedDict(branch) if isinstance(branch, dict) else branch

    def __setitem__(self, keys, value):
        # Allows setting top-level item when a single key was provided
        if not isinstance(keys, tuple):
            keys = (keys,)

        branch = self.dict
        for key in keys[:-1]:
            if key not in branch:
                branch[key] = {}
            branch = branch[key]
        branch[keys[-1]] = value

    def clear(self):
        self.dict.clear()

    def get(self, keys, default=None, *default_args, **default_kwargs):
        try:
            return self.__getitem__(keys)
        except KeyError:
            if callable(default):
                print('QQQQ, {}'.format(*default_args))
                return default(*default_args, **default_kwargs)
            else:
                return default

    def pop(self, keys, default=None):
        try:
            result = self.__getitem__(keys)
            self.dict.pop(keys)
        except KeyError:
            result = default
        return result


def nested_dict_by_excel(filename, index_col=None, usecols=None, string=False, sheet_as_key=True, transpose=True,
                         force_str=False, header=0):
    """return {excel sheet name:{index1:{index2:{.....:[values, ....], ...}}}
       need to input filename and index_col like [0, 1]"""
    # default index col
    if index_col is None:
        index_col = 0

    full_dict = {}
    df_dict = pd.read_excel(filename, sheet_name=None, index_col=index_col, usecols=usecols, header=header,
                            dtype=object)
    for sheet, df in df_dict.items():
        nest = NestedDict()
        if transpose:
            df = df.T
        # if string:
        #     for keys, values in df.to_dict('list').items():
        #         nest[keys] = values[0]
        # else:
        #     for keys, values in df.to_dict('list').items():
        #         nest[keys] = [n for n in values if pd.notnull(n)]
        for keys, values in df.to_dict('list').items():
            if string:
                nest[str(keys) if force_str else keys] = str(values[0]) if force_str else values[0]
            else:
                nest[str(keys) if force_str else keys] = [str(n) if force_str else n for n in values if pd.notnull(n)]
            # else:
            #     if string:
            #         nest[keys] = values[0]
            #     else:
            #         nest[keys] = [n for n in values if pd.notnull(n)]
        if sheet_as_key:
            full_dict[sheet] = nest.dict
        else:
            full_dict = nest.dict
    return full_dict


########################################################################################################################
dump_data_root_folder, read_data_root_folder = (os.path.dirname(os.getcwd()) for q in range(2))
warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
current_date = None
current_start_date = None
# summary_cache = NestedDict()
# current_bsl_yield = NestedDict()
# custom_raw_code_style_dict = NestedDict()
update_all_report_until_error_log = []
summary_cache, current_bsl_yield, custom_raw_code_style_dict = (NestedDict() for i in range(3))
part_group_code_dict = nested_dict_by_excel(r'code adjustment/group by part.xlsx', index_col=[0, 1])
part_code_dict = nested_dict_by_excel(r'code adjustment/Active code by part.xlsx')
ooc_reference = nested_dict_by_excel(r'reference/ooc_ref.xlsx', string=True)
des_dict = nested_dict_by_excel(r'Defect code lists.xlsx', usecols='A:B', string=True, sheet_as_key=False,
                                force_str=True)
part_action_item_dict = nested_dict_by_excel(r'reference/action_item_for_ppt.xlsx', index_col=[0, 1], string=True)
summary_table_parts = nested_dict_by_excel(r'code adjustment/docx summary table parts and goal.xlsx', string=True,
                                           sheet_as_key=False)
df_ref_dict = pd.read_excel(r'reference/baseline.xlsx', index_col=0, sheet_name=None,
                            header=[0, 1])
ooc_code_dict = nested_dict_by_excel(r'code adjustment/ooc_code.xlsx', force_str=True)

########################################################################################################################
cache1 = LFUCache(maxsize=len(des_dict) + 10)
cache_locker = RLock()
########################################################################################################################
ser_ooc = {}
ooc_part_list = cfg['OOC part']
ooc_parts_show_in_doc = cfg['part show in docx summary']['OOC']
for _part in ooc_part_list:
    ser_ooc[_part] = ooc_code_dict[_part]
active_parts_and_product_group = cfg['active part']
# summary_table_parts = cfg['summary table parts and goal']
bsl_parts = cfg['BSL part']
docx_summary_table_parts = cfg['part show in docx summary']['table']
docx_summary_charts_parts = cfg['part show in docx summary']['charts']
docx_scrap_pcs_parts = cfg['part show in docx summary']['scrap pcs']
bsl_tolerance = float(cfg['BSL tolerance'])


def chk_dir(path):
    """
    :param path:chk directory existence, if not, create path folder
    :return:None
    """
    exist = os.path.exists(path)
    if not exist:
        try:
            os.makedirs(path)
        except FileExistsError:
            pass  # do nothing in parallel multi process


########################################################################################################################
def query_active_parts():
    """
    :return: list contained active_parts to process
    """
    parts = list()
    for i in cfg['active part'].items():
        parts += i[1]
    parts.sort(key=len)
    return parts


########################################################################################################################
active_parts = query_active_parts()


def query_parent_child_parts(parts):
    """

    :param parts:active_parts
    :return: {parent:[child part, ....], ...}
        Ex:{AP196:[AP19601, Ap19602, ...], ...}
    """
    parent, child = set(), set()
    result = dict()
    for part in parts:
        if len(part) == 5:
            parent.add(part)
        else:
            child.add(part)
    for c in child:
        parent_str = c[:5]
        if parent_str in parent:
            ori_lst = result.get(parent_str, [])
            ori_lst.append(c)
            result[parent_str] = ori_lst
    return result


parts_parents_to_child = query_parent_child_parts(active_parts)


########################################################################################################################

def query_active_code(string=False, to_set=False):
    result = {}
    for part in part_code_dict.keys():
        active_codes = []
        reference = part_code_dict[part].items()
        for side, codelist in reference:
            if string:
                codelist = [str(n) for n in codelist]
            active_codes += codelist
        if to_set:
            active_codes = set(active_codes)
        result[part] = active_codes
    return result


part_active_code = query_active_code()

part_active_code_string = query_active_code(string=True)

part_active_code_str_set = query_active_code(string=True, to_set=True)


########################################################################################################################

def goals(part):
    # summery cache
    df_ref = df_ref_dict.get(part)
    if df_ref is None:
        bsl_yield = np.nan
    else:
        bsl_yield = df_ref.at[df_ref.last_valid_index(), ('SUMMARY', 'Yield')]
    current_bsl_yield[part, 'Baseline yield'] = bsl_yield
    goal = summary_table_parts[part]
    return ['{:.0%}'.format(goal), "Q'ty"]


def query_summary_table_headers(selected_parts, return_index=False):
    tuples = []
    for product_group, parts in active_parts_and_product_group.items():
        tup = [(product_group, part, goal) for part in parts for goal in goals(part) if part in selected_parts]
        tuples += tup
    # tuples += ('Scrap pcs', 'Scrap pcs', 'Scrap pcs')
    tuples = [('Item', 'Part', 'Goal')] + tuples
    index = pd.MultiIndex.from_tuples(tuples)
    if return_index:
        return index
    df = pd.DataFrame(index=index)
    return df


summary_table_header = query_summary_table_headers(summary_table_parts)

docx_summary_table_header = query_summary_table_headers(docx_summary_table_parts, return_index=True)


########################################################################################################################

def query_description_dict():
    """return {part:{code:description, ....}}
    CP description is equal to group name
    """
    result = {}
    # parts = active_parts.copy()
    # for part in parts:
    #     cp_des_dict = {}
    #     for group, codelist in part_group_code_dict[part]['CP'].items():
    #         for code in codelist:
    #             cp_des_dict[str(code)] = group
    #     result[part] = {**des_dict, **cp_des_dict}
    parts = part_group_code_dict.keys()
    for part in parts:
        cp = part_group_code_dict[part].get('CP')
        if cp is not None:
            cp_des_dict = {}
            for group, codelist in part_group_code_dict[part]['CP'].items():
                for code in codelist:
                    cp_des_dict[str(code)] = group
            result[part] = {**des_dict, **cp_des_dict}
        else:
            result[part] = {**des_dict}
    return result


description_dict = query_description_dict()


########################################################################################################################

def query_custom_raw_code_output_dict():
    try:
        custom_raw_code_output_dict = pd.read_excel(r'code adjustment/custom_row_code_output.xlsx', sheet_name=None,
                                                    header=[0, 1, 2])
        for part, df in custom_raw_code_output_dict.items():
            str_columns = [(sty, des, str(code)) for sty, des, code in df.columns]
            df.columns = pd.MultiIndex.from_tuples(str_columns)
    except FileNotFoundError:
        return dict()
    return custom_raw_code_output_dict


df_custom_raw_code_output_dict = query_custom_raw_code_output_dict()


########################################################################################################################
def query_custom_raw_code_style_dict(dicts):
    result = {}
    for part, df in dicts.items():
        df_ref = df.copy()
        df_ref.columns = df.columns.droplevel(1)
        correction_dict = {code: style for style, code in df_ref.columns}
        for n in ['month', 'shipping_date', 'week']:
            correction_dict[n] = 'DATETIME'
        result[part] = correction_dict
        return result


custom_raw_code_style_dict = query_custom_raw_code_style_dict(df_custom_raw_code_output_dict)


########################################################################################################################
def query_gross_die():
    from Database_query import gross_die_by_part, gross_die_modified_dict
    result = dict()
    # oracle_query_parts = set()
    parts = (set(active_parts) | set(gross_die_modified_dict.keys()))
    for part in parts:
        result[part] = gross_die_by_part(part)
        # oracle_query_parts.add(part[:5])  # only query first 5 character in part
    return result


gross_die = query_gross_die()


########################################################################################################################
def update_ref_dict():
    global df_ref_dict
    df_ref_dict = pd.read_excel(r'reference/baseline.xlsx', index_col=0, sheet_name=None, header=[0, 1])


def constrained_sum_sample_pos(n, total):
    """Return a randomly chosen list of n positive integers summing to total.
    Each such list is equally likely to occur."""
    import random
    dividers = sorted(random.sample(range(1, total), n - 1))
    return [a - b for a, b in zip(dividers + [total], [0] + dividers)]


def df_example_1(part):
    """

    :param part: part   0     1  2    3  4    5    6
    :return:df like  0  WAFER_ID PART_ID 1000 1100 1101 ...
                     1  1     1  1    1  1    1    1    ...

    """
    # active_codes = []
    ids = ["WAFER_ID", "PART_ID"]
    # for side, codelist in part_code_dict[part].items():
    #     active_codes += codelist
    active_codes = part_active_code[part]
    sheet_list_header = ids + active_codes
    df = pd.DataFrame({0: sheet_list_header, 1: [1 for n in sheet_list_header]}).T
    return df.T


def df_example_2(part, x, random_content=None, random_size=0.9):
    """

           WAFER_ID PART_ID 1000 1100 1101 ...
    df like  1  ........randint betwen 0 ~9
             2
                   ..
             5
    """
    # active_codes = []
    ids = ["WAFER_ID", "PART_ID"]
    # for side, codelist in part_code_dict[part].items():
    #     codelist = [str(n) for n in codelist]
    #     active_codes += codelist
    active_codes = part_active_code[part]
    sheet_list_header = ids + active_codes
    if random_content:
        # df = pd.DataFrame(
        # np.random.randint(low=0, high=10, size=(x+1, len(sheet_list_header))), columns = sheet_list_header)

        df = pd.DataFrame(
            [['a', 'b'] + constrained_sum_sample_pos(len(sheet_list_header) - 2,
                                                     round(int(gross_die[part]) * random_size))
             for l in range(x + 1)], columns=sheet_list_header)
    else:
        df = pd.DataFrame({n: [n + 1 for i in sheet_list_header] for n in range(x + 1)}, index=sheet_list_header).T
    return df


# def df_month_summary_example(part, date):
#     from random import randint
#     from WeekTime_parser import week_numbers_in_last_month
#     df = pd.DataFrame()
#     for i in week_numbers_in_last_month(date=date):
#         df1 = df_weekly_update(
#             df_example_2(part, x=randint(20, 100), random_content=True, random_size=randint(50, 100) / 1000), part,
#         )['df_week_summary']
#         df = pd.concat([df, df1])
#     df = df.reset_index(drop=True)
#     df2 = week_numbers_in_last_month(date=date, only_week_number=False)
#     df = pd.concat([df2, df], axis=1)
#     return df


# def freezeargs(_func):
#     """Transform mutable dictionnary
#     Into immutable
#     Useful to be compatible with cache
#     """
#
#     @functools.wraps(_func)
#     def wrapped(*args, **kwargs):
#         args = tuple([frozendict(arg) if isinstance(arg, dict) else arg for arg in args])
#         kwargs = {k: frozendict(v) if isinstance(v, dict) else v for k, v in kwargs.items()}
#         return _func(*args, **kwargs)
#
#     return wrapped


def code_side_dict():
    """return {part:{code:side, ....}}"""
    nest = NestedDict()
    for part, sides in part_code_dict.items():
        for side, codelist in sides.items():
            for code in codelist:
                nest[part, code] = side
    return nest.dict


# def formula_gen_dict(part, code_loc_dict, y):
#     """return {side:{group:formula string, ...}}
#     formula like =SUM(Sheet3!N2:N100,Sheet3!U2:U100,Sheet3!AD2:AD100)/14419/'Weekly data'!A2
#     """
#     nest = NestedDict()
#     num = gross_die[part]
#     work_dist = y + 2
#     for side, groups in part_group_code_dict[part].items():
#         for group, code_list in groups.items():
#             base = str()
#             for code in code_list:
#                 column = xl_col_to_name(code_loc_dict[code])
#                 base += 'list_code!{0}{1}:{0}1048576,'.format(column, work_dist)
#             base = base.rstrip(',')
#             formula = "=SUM({0})/{1}/'Weekly data'!A2".format(base, num)
#             nest[side, group] = formula
#     return nest.dict


# def template(part):
#     """
#     initial set up
#     """
#     writer = pd.ExcelWriter(r'templates/{}.xlsx'.format(part), engine='xlsxwriter')
#     workbook = writer.book
#     cell_fsv = workbook.add_format({'bg_color': '#CCFFFF', 'border': True})
#     cell_bsv = workbook.add_format({'bg_color': '#CCFFCC', 'border': True})
#     cell_cp = workbook.add_format({'bg_color': '#FFFF99', 'border': True})
#     cell_format = workbook.add_format({'border': True})
#     cell_formula = workbook.add_format({'num_format': 0x0a, 'border': True})
#     # cp_loc_list, fsv_loc_list, bsv_loc_list, store group locations groupby side
#     group, chk_duplicate, cp_loc_list, fsv_loc_list, bsv_loc_list = ([] for l in range(5))
#     # loc_dict store locations of "Q'ty", "Yield", "Yield_loss", "CP", "BSV", "FSV"
#     # code_loc_dict store locations of code
#     code_loc_dict, loc_dict = ({} for l in range(2))
#     # store group locations groupby side
#     side_loc_dict = {'FSV': fsv_loc_list, 'BSV': bsv_loc_list, 'CP': cp_loc_list}
#     ids = ["WAFER_ID", "PART_ID"]
#     # for side, codelist in part_code_dict[part].items():
#     #     active_codes += codelist
#     active_codes = part_active_code[part]
#     sheet_list_header = ids + active_codes
#     """
#     sheet list_code
#     """
#     # df = pd.DataFrame(sheet_list_header).T
#     df = pd.DataFrame({0: [description_dict[part].get(n, '') for n in sheet_list_header], 1: sheet_list_header}).T
#     df.to_excel(writer, sheet_name='list_code', header=False, index=False)
#     worksheet1 = writer.sheets['list_code']
#     adjustment_source = code_side_dict()[part]
#     # relative work distance to x axis in excel
#     y = 1
#     for idx, col in enumerate(df):
#         n = df[idx][y]
#         if n in ids:
#             worksheet1.conditional_format(y - 1, idx, y, idx, {'type': 'no_errors', 'format': cell_format})
#             worksheet1.set_column(idx, idx, 11.38)
#             continue
#         elif adjustment_source[n] == '正檢':
#             worksheet1.conditional_format(y - 1, idx, y, idx, {'type': 'no_errors', 'format': cell_fsv})
#         elif adjustment_source[n] == '背檢':
#             worksheet1.conditional_format(y - 1, idx, y, idx, {'type': 'no_errors', 'format': cell_bsv})
#         elif adjustment_source[n] == 'CP':
#             worksheet1.conditional_format(y - 1, idx, y, idx, {'type': 'no_errors', 'format': cell_cp})
#         worksheet1.set_column(idx, idx, 8.38)
#         code_loc_dict[n] = idx
#     worksheet1.freeze_panes(2, 2)  # freeze column A, B and row 1, 2
#
#     """
#     sheet Weekly data
#     and write formula in Weekly
#     """
#
#     for side, grp in part_group_code_dict[part].items():
#         group = group + list(grp)
#     df = pd.DataFrame(["Q'ty", "Yield", "Yield_loss", "CP", "BSV", "FSV"] + group).T
#     df.to_excel(writer, sheet_name='Weekly data', header=False, index=False)
#     worksheet = writer.sheets['Weekly data']  # pull worksheet object
#     adjustment_source = part_group_code_dict[part]
#     fsv_list = list(adjustment_source['正檢'])
#     bsv_list = list(adjustment_source['背檢'])
#     cp_list = list(adjustment_source['CP'])
#     fsv_bsv_duplicate = list(set(fsv_list) & set(bsv_list))
#
#     formula = formula_gen_dict(part, code_loc_dict, y=y)
#     for idx, col in enumerate(df):  # loop through all columns
#         n = df[idx][0]
#         if n in ["Q'ty", "Yield", "Yield_loss"]:
#             worksheet.conditional_format(0, idx, 0, idx, {'type': 'no_errors', 'format': cell_format})
#             loc_dict[n] = idx
#         elif n in fsv_bsv_duplicate:
#             if n in chk_duplicate:
#                 worksheet.conditional_format(0, idx, 0, idx, {'type': 'no_errors', 'format': cell_bsv})
#                 worksheet.write_formula(1, idx, formula['背檢'][n], cell_formula)
#                 bsv_loc_list.append(idx)
#             else:
#                 worksheet.conditional_format(0, idx, 0, idx, {'type': 'no_errors', 'format': cell_fsv})
#                 worksheet.write_formula(1, idx, formula['正檢'][n], cell_formula)
#                 chk_duplicate.append(n)
#                 fsv_loc_list.append(idx)
#         elif n in fsv_list or n == 'FSV':
#             worksheet.conditional_format(0, idx, 0, idx, {'type': 'no_errors', 'format': cell_fsv})
#             if n == 'FSV':
#                 loc_dict[n] = idx
#             else:
#                 worksheet.write_formula(1, idx, formula['正檢'][n], cell_formula)
#                 fsv_loc_list.append(idx)
#         elif n in bsv_list or n == 'BSV':
#             worksheet.conditional_format(0, idx, 0, idx, {'type': 'no_errors', 'format': cell_bsv})
#             if n == 'BSV':
#                 loc_dict[n] = idx
#             else:
#                 worksheet.write_formula(1, idx, formula['背檢'][n], cell_formula)
#                 bsv_loc_list.append(idx)
#         elif n in cp_list or n == 'CP':
#             worksheet.conditional_format(0, idx, 0, idx, {'type': 'no_errors', 'format': cell_cp})
#             if n == 'CP':
#                 loc_dict[n] = idx
#             else:
#                 worksheet.write_formula(1, idx, formula['CP'][n], cell_formula)
#                 cp_loc_list.append(idx)
#         worksheet.set_column(idx, idx, 8.38)  # set column width
#
#     # write formula below header FSV, BSV, CP
#     for side, locs in side_loc_dict.items():
#         base = str()
#         for loc in locs:
#             column = xl_col_to_name(loc)
#             base += '{0}2,'.format(column)
#         base = base.rstrip(',')
#         formula = "=SUM({})".format(base)
#         worksheet.write_formula(1, loc_dict[side], formula, cell_formula)
#     # write formula below header "Q'ty", "Yield", "Yield_loss"
#     for k in ["Q'ty", "Yield", "Yield_loss"]:
#         if k == "Q'ty":
#             formula = "=COUNTA(list_code!A:A)-1"
#             worksheet.write_formula(1, loc_dict[k], formula, cell_format)
#             continue
#         elif k == "Yield":
#             column = xl_col_to_name(loc_dict["Yield_loss"])
#             formula = "=1-{}2".format(column)
#         elif k == "Yield_loss":
#             base = str()
#             for l in ["CP", "FSV", "BSV"]:
#                 column = xl_col_to_name(loc_dict[l])
#                 base += '{0}2+'.format(column)
#             base = base.rstrip('+')
#             formula = "={}".format(base)
#         worksheet.write_formula(1, loc_dict[k], formula, cell_formula)
#
#     writer.save()


def df_week_contents_gen(index, part, df, summary=True):
    """
    weekly data init set up
    :param index: Multi index like
    SUMMARY                                 FSV...
    Q'ty	Yield	Yield_loss	CP	BSV	FSV	Bond ring shift	Scratch	Defect on glass	ILD hump....

    :param part:part
    :param df:
               WAFER_ID PART_ID 1000 1100 1101 ...
    df       1  ........int of defect count
             2
                   ..
             5
    :param summary: determine to generate weekly summary or wafer by wafer data
    :return: dict contain content list of defect group yield loss of each side and cp
    """
    adjustment = part_group_code_dict[part]
    fsv_list = list(adjustment['正檢'])  # fsv defect group name
    bsv_list = list(adjustment['背檢'])
    cp_list = list(adjustment['CP'])
    content_list = []
    qty = len(df.index)
    formula = df_formula_dict(part, summary=summary)
    df.rename(columns={old: int2str(old) for old in list(df)}, inplace=True)
    for first, second in index:
        if summary:
            result = 0
            if (first, second) == ('SUMMARY', "Q'ty"):
                result = qty
        else:
            result = pd.Series([0] * qty)
            qty = 1
        if first == 'FSV':
            result = pd.eval(formula['正檢'][second].format(qty))
        elif first == 'BSV':
            result = pd.eval(formula['背檢'][second].format(qty))
        elif first == 'CP':
            result = pd.eval(formula['CP'][second].format(qty))
        content_list.append(result)

    return {'content': content_list, 'FSV': fsv_list, 'BSV': bsv_list, 'CP': cp_list}


@cached(cache1, key=partial(hashkey, 'str2int'))
def str2int(s):
    """turn string back to unique int:defcode"""
    chars = ascii_uppercase
    i = 0
    if s in {'WAFER_ID', 'PART_ID'}:
        return s
    else:
        for c in reversed(s):
            i *= len(chars)
            i += chars.index(c)
        return i


@cached(cache1, key=partial(hashkey, 'int2str'), lock=cache_locker)
def int2str(i):
    """turn int:defcode to unique string base """
    chars = ascii_uppercase
    s = ""
    if isinstance(i, str):
        if i in {'WAFER_ID', 'PART_ID'}:
            return i
        else:
            i = int(i)
    while i:
        s += chars[i % len(chars)]
        i //= len(chars)
    return s


# @cached(cache1, key=partial(hashkey, 'formula_constructor'), lock=cache_locker)
# def formula_constructor(code_tup, ele, num, qty):
#     base = str()
#     for code in code_tup:
#         code = int2str(code)
#         base += ele.format(code)
#     base = base.rstrip('+')
#     formula = "({})/{}/{}".format(base, num, qty)
#     return formula

@cached(cache1, key=partial(hashkey, 'formula_constructor'), lock=cache_locker)
def formula_constructor(code_tup, ele, num):
    base = str()
    for code in code_tup:
        code = int2str(code)
        base += ele.format(code)
    base = base.rstrip('+')
    formula = "({})/{}/".format(base, num) + '{}'
    return formula


@cached(cache1, key=partial(hashkey, 'df_formula_dict'), lock=cache_locker)
def df_formula_dict(part, summary=True, ooc=False):
    """return {side:{group:formula string, ...}}
    formula like =SUM(Sheet3!N2:N100,Sheet3!U2:U100,Sheet3!AD2:AD100)/14419/'Weekly data'!A2
    """
    nest = NestedDict()
    num = gross_die[part]
    if summary:
        ele = 'df.{}.sum()+'
    else:
        ele = 'df.{}+'
    if ooc:
        ser = ser_ooc[part]
        # active_codes = []
        # reference = part_code_dict[part].items()
        # for side, codelist in reference:
        #     active_codes += codelist
        active_codes = part_active_code[part]
        active_codes = list(map(str, active_codes))
        for ooc_group, code_list in ser.items():
            function_codes = tuple(list(set(active_codes) & set(code_list)))
            nest[ooc_group] = formula_constructor(function_codes, ele, num)
    else:
        for side, groups in part_group_code_dict[part].items():
            for group, code_list in groups.items():
                code_list = tuple(code_list)
                nest[side, group] = formula_constructor(code_list, ele, num)
                # base = str()
                # for code in code_list:
                #     code = int2str(code)
                #     base += ele.format(code)
                # base = base.rstrip('+')
                # formula = "({})/{}/{}".format(base, num, qty)
                # nest[side, group] = formula
    return nest.dict


# @cached(cache1, key=partial(hashkey, 'df_formula_dict'), lock=cache_locker)
# def df_formula_dict(part, qty, summary=True, ooc=False):
#     """return {side:{group:formula string, ...}}
#     formula like =SUM(Sheet3!N2:N100,Sheet3!U2:U100,Sheet3!AD2:AD100)/14419/'Weekly data'!A2
#     """
#     nest = NestedDict()
#     num = gross_die[part]
#     if summary:
#         ele = 'df.{}.sum()+'
#     else:
#         ele = 'df.{}+'
#         qty = 1
#     if ooc:
#         ser = ser_ooc[part]
#         # active_codes = []
#         # reference = part_code_dict[part].items()
#         # for side, codelist in reference:
#         #     active_codes += codelist
#         active_codes = part_active_code[part]
#         active_codes = list(map(str, active_codes))
#         for ooc_group, code_list in ser.items():
#             function_codes = tuple(list(set(active_codes) & set(code_list)))
#             nest[ooc_group] = formula_constructor(function_codes, ele, num, qty)
#     else:
#         for side, groups in part_group_code_dict[part].items():
#             for group, code_list in groups.items():
#                 code_list = tuple(code_list)
#                 nest[side, group] = formula_constructor(code_list, ele, num, qty)
#                 # base = str()
#                 # for code in code_list:
#                 #     code = int2str(code)
#                 #     base += ele.format(code)
#                 # base = base.rstrip('+')
#                 # formula = "({})/{}/{}".format(base, num, qty)
#                 # nest[side, group] = formula
#     return nest.dict

def df_weekly_update(df, part):
    """

    input df like



           WAFER_ID PART_ID 1000 1100 1101 ...
    df       1  ........int of defect count
             2
                   ..
             5

    return df like~


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
    chinese2english = {'正檢': 'FSV', '背檢': 'BSV', 'CP': 'CP'}
    tuples1 = [('SUMMARY', n) for n in ["Q'ty", "Yield", "Yield_loss", "CP", "BSV", "FSV"]]
    tuples2 = []
    for side, grp in part_group_code_dict[part].items():
        for n in list(grp):
            tuples2.append((chinese2english[side], n))
    tuples = tuples1 + tuples2
    index = pd.MultiIndex.from_tuples(tuples)
    result_dict = dict()
    for name, summary in [('df_summary', True), ('df_pcs_group', False)]:
        df_input = df.copy()
        week_contents = df_week_contents_gen(index, part, df_input, summary=summary)
        df_week = pd.DataFrame(week_contents['content'], index=index).T
        if not summary:
            df_week.drop(('SUMMARY', "Q'ty"), axis=1, inplace=True)
            # df_week.drop(('SUMMARY', "Scrap_pcs"), axis=1, inplace=True)
            df_wafer = df.iloc[:, :2]
            df_wafer.columns = pd.MultiIndex.from_product([['WAFER'], df_wafer.columns])
            df_week = pd.concat([df_wafer, df_week], axis=1, sort=False)
        for n in df_week.index:
            for i in ["CP", "BSV", "FSV"]:
                #     df_week.ix[n, ('SUMMARY', i)] = df_week.loc[n, i][week_contents[i]].sum()
                # df_week.ix[n, ('SUMMARY', "Yield_loss")] = df_week.ix[n, ('SUMMARY', "CP")] + df_week.ix[
                #     n, ('SUMMARY', "BSV")] + df_week.ix[n, ('SUMMARY', i)]
                # df_week.ix[n, ('SUMMARY', "Yield")] = 1 - df_week.ix[n, ('SUMMARY', "Yield_loss")]
                df_week.at[n, ('SUMMARY', i)] = df_week.loc[n, i][week_contents[i]].sum()
            df_week.at[n, ('SUMMARY', "Yield_loss")] = df_week.at[n, ('SUMMARY', "CP")] + df_week.at[
                n, ('SUMMARY', "BSV")] + df_week.at[n, ('SUMMARY', i)]
            df_week.at[n, ('SUMMARY', "Yield")] = 1 - df_week.at[n, ('SUMMARY', "Yield_loss")]
        if ooc_parts(part):
            df_input = df.copy()
            df_ooc = df_weekly_ooc_update(df_input, part)
            df_ooc_drop = df_ooc.drop(columns='WAFER')
            if summary:
                df_week_ooc_summary = pd.DataFrame(df_ooc_drop.mean()).T
                df_week = pd.concat([pd.DataFrame(data=ooc_data(df_ooc, part)), df_week, df_week_ooc_summary], axis=1)
            else:
                df_week = pd.concat([df_week, df_ooc_drop], axis=1)
        result_dict[name] = df_week
    return result_dict


def df_code_update(df, part):
    """
    add description on top of code df

          WAFER         微粒Particle 污染 contamination ...  Offset RB1 SPAT DPAT
   WAFER_ID PART_ID       1000             1100 ...       3   4    8    9
0         1       1          1                1 ...       1   1    1    1
1         2       2          2                2 ...       2   2    2    2
2         3       3          3                3 ...       3   3    3    3
3         4       4          4                4 ...       4   4    4    4
4         5       5          5                5 ...       5   5    5    5

    """
    # active_codes = []
    ids = ["WAFER_ID", "PART_ID"]
    # reference = part_code_dict[part].items()
    # for side, codelist in reference:
    #     active_codes += codelist
    active_codes = part_active_code[part]
    active_codes = list(map(str, active_codes))
    sheet_list_header = ids + active_codes
    columns = [(description_dict[part].get(n, 'WAFER'), n) for n in sheet_list_header]
    df.columns = pd.MultiIndex.from_tuples(columns)
    return df


def df_custom_row_code_output(df, part):
    """
    :param df: code df with description on top
    :param part: part
    :return: df_code with custom order and description
    """
    df_ori = df.copy()
    df_ref = df_custom_raw_code_output_dict[part].copy()
    correction_dict = {code: des for sty, des, code in df_ref.columns}
    for n in ['month', 'shipping_date', 'week']:
        correction_dict[n] = 'DATETIME'
    ref_cols = set(df_ref.columns.droplevel().droplevel())
    ori_cols = set(df_ori.drop(columns=['DATETIME', 'WAFER']).columns.droplevel(0))
    del_cols = ori_cols - ref_cols
    del_cols = [(description_dict[part][n], n) for n in del_cols]
    df_ori.drop(columns=del_cols, inplace=True)
    new_columns = [(correction_dict.get(code, 'WAFER'), code) for des, code in df_ori.columns]
    df_ori.columns = pd.MultiIndex.from_tuples(new_columns)
    df_ref.columns = df_ref.columns.droplevel(0)
    cols_order = list(df_ref)
    # cols_order = list(df)[:2] + cols_order
    cols_order = list(df)[:5] + cols_order
    df_result = pd.concat([df_ref, df_ori], sort=False)
    df_result.fillna(0, inplace=True)
    df_result = df_result.loc[:, cols_order]
    return df_result


def df_weekly_ooc_update(df, part):
    """

    :param df: code df
    :param part: 'AP196' or ....
    :return:df like
      WAFER         OOC_GROUP
   WAFER_ID PART_ID BSV others Bond chuck mark Contamination on cavity Contamination/Passivation anomaly ...
0         1       1   0.003591        0.000718                0.002154                          0.003232 ...
1         2       2   0.007181        0.001436                0.004309                          0.006463 ...
2         3       3   0.010772        0.002154                0.006463                          0.009695 ...
3         4       4   0.014363        0.002873                0.008618                          0.012926 ...
4         5       5   0.017953        0.003591                0.010772                          0.016158 ...
5         6       6   0.021544        0.004309                0.012926                          0.019390 ...
6         7       7   0.025135        0.005027                0.015081                          0.022621 ...
7         8       8   0.028725        0.005745                0.017235                          0.025853 ...
8         9       9   0.032316        0.006463                0.019390                          0.029084 ...
9        10      10   0.035907        0.007181                0.021544                          0.032316 ...
10       11      11   0.039497        0.007899                0.023698                          0.035548 ...
    """
    # qty = len(df.index)
    ser = ser_ooc[part]
    tuples = [('OOC_GROUP', n) for n in ser.keys()]
    index = pd.MultiIndex.from_tuples(tuples)
    formula = df_formula_dict(part, summary=False, ooc=True)
    df.rename(columns={old: int2str(old) for old in list(df)}, inplace=True)
    contents = [pd.eval(formula[grp].format(1), local_dict={'df': df}) for grp in ser.keys()]
    # contents = []
    # for a in ser_ooc.index:
    #     result = pd.eval(formula[a])
    #     contents.append(result)
    # contents = list(map(lambda a: pd.eval(formula[a]), ser_ooc.index))
    df_data = pd.DataFrame(contents, index=index).T
    df_wafer = df.iloc[:, :2]
    df_wafer.columns = pd.MultiIndex.from_product([['WAFER'], df_wafer.columns])
    df_ooc = pd.concat([df_wafer, df_data], axis=1)

    return df_ooc


def ooc_data(df_ooc, part):
    """
    :param df_ooc:
    :param part:
    :return: dict contain ooc_indexes and ooc_stats
    ooc_indexes: {groups_name: indexes for ooc happened, ...}
    ooc_stats: {'ooc counts': ooc qtys, 'ooc wafers': wafers that have ooc group}
    """
    qty = int()
    ooc_dict, ooc_stats = ({} for l in range(2))
    ref = ooc_reference
    num = gross_die[part]
    ref_dict = {grp: count / num for grp, count in ref[part].items()}
    df_target = df_ooc['OOC_GROUP']
    df = pd.DataFrame()
    for grp, ref_line in ref_dict.items():
        df1 = df_target[(df_target[grp] > ref_line)]
        df_len = len(df1.index)
        if df_len > 0:
            ooc_dict[grp + ': {}'.format(df_len)] = ', '.join(
                e for e in [df_ooc[('WAFER', 'WAFER_ID')][i] for i in df1.index]) + ', '
            qty += df_len
            df = pd.concat([df, df1]).drop_duplicates()
    ooc_stats['ooc counts'] = qty
    ooc_stats['ooc wafers'] = len(df.index)
    if qty > 0:
        wafers = ''.join(e for e in ooc_dict.values()).rstrip(', ')
        groups = ', '.join(e for e in ooc_dict.keys())
    else:
        wafers, groups = ('Normal' for l in range(2))
    # df = pd.concat([df_target[(df_target[grp] > ref_line)] for grp, ref_line in ref_dict.items()])
    # return {'df_confirm_ooc': df, 'ooc_indexes': ooc_indexes} if need ooc df
    return {('OOC', 'wafers'): [wafers],
            ('OOC', 'groups'): [groups],
            ('OOC', 'counts'): [ooc_stats['ooc counts']],
            ('OOC', 'pcs'): [ooc_stats['ooc wafers']]}


ix_not_use = pd.MultiIndex.from_tuples(
    [('DATETIME', 'month'), ('DATETIME', 'shipping_date'), ('DATETIME', 'shipping_date'), ('DATETIME', 'week'),
     ('WAFER', 'WAFER_ID'), ('WAFER', 'PART_ID'), ('SUMMARY', "Q'ty"), ('OOC', 'wafers'),
     ('OOC', 'groups'), ('OOC', 'counts'), ('OOC', 'pcs')])


def df_sub_bsl(df, part):
    """

    :param df: origin df
    :param part: part
    :return: origin df - BSL df - 0.002
    """
    # lexsorted df is not desired result, so use some workaround = =
    df_ref = df_ref_dict.get(part)  # TODO:fix lexsort problem
    ref_dict = None
    ooc = ooc_parts(part)
    if ooc:
        num = gross_die[part]
        ref_dict = {grp: float(count / num) for grp, count in ooc_reference[part].items()}
    if df_ref is None:
        return {'df_write': pd.DataFrame(), 'df_ref': pd.DataFrame()}
    else:
        # df_ref = pd.concat([df_ref] * len(df.index), ignore_index=True)
        # df_ref.index = df.index
        idx_ref_sub = [x for x in df_ref.columns if x not in ix_not_use]
        idx_df_sub = [x for x in df.columns if x not in ix_not_use]
        idx_df_head = [x for x in df.columns if x in ix_not_use]
        df_ref1 = df_ref.ix[:, idx_ref_sub]
        df1 = df.ix[:, idx_df_sub]
        df_head = df.ix[:, idx_df_head]
        # df2 = df1.sub(df_ref1, level=1)
        # lexsorted df is not desired result, so use some workaround = =
        for idx in df1.index:
            for first, second in df1.columns:
                if first == 'OOC_GROUP':
                    df1.at[idx, (first, second)] = df1.ix[idx, (first, second)] - ref_dict[second]
                else:
                    df1.at[idx, (first, second)] = df1.ix[idx, (first, second)] - df_ref1.ix[
                        df_ref1.last_valid_index(), (first, second)]
        result1 = pd.concat([df_head, df1], axis=1)
        if ooc:
            df1['OOC_GROUP'] += bsl_tolerance  # 0.002
        df1.loc[:, ('SUMMARY', 'Yield')] *= -1  # inverse Yield value to compare with BSL
        df1 -= bsl_tolerance  # tolerance 0.2%
        df_head = df_head.applymap(lambda x: np.nan)
        result2 = pd.concat([df_head, df1], axis=1)
        return {'df_write': result1, 'df_ref': result2}


class NoWeekBefore(Exception):
    pass


class BreakIt(Exception):
    pass


def df_monthly_summary(df):
    """

    :param df: raw df read from excel data like
     DATETIME                   SUMMARY									FSV...
month	shipping_date	week	Q'ty	Yield...
Nov	20181026-20181102	W44	43	96.98%	3.02%...
Nov	20181102-20181109	W45	25	90.99%	9.01%...
Nov	20181109-20181116	W46	52	96.01%	3.99%...
Nov	20181116-20181123	W47	68	96.98%	3.02%...
Nov	20181123-20181130	W48	46	93.50%	6.50%...
        Nov	234	95.44%	4.56%	0.36%	1.65%...
Dec	20181130-20181207	W49	89	96.98%	3.02%...
Dec	20181207-20181214	W50	93	93.00%	7.00%...
Dec	20181214-20181221	W51	98	94.00%	6.00%...
Dec	20181221-20181228	W52	16	96.98%	3.02%...

    :return: result df like
    SUMMARY SUMMARY                                                                                              FSV ...
month     month shipping_date week Q'ty     Yield Yield_loss          CP        BSV        FSV Passivation crack     ...
10          NaN           NaN  Dec  296  0.947449  0.0525511  0.00414989  0.0191894  0.0292118       0.000872192     ...
    """
    if df.isnull().values.any():
        head_index = df[df.isnull().any(axis=1)].index.max() + 1
    else:
        head_index = df.index.min()
    tail_index = df.index.max() + 1
    df = df.loc[head_index:]  # select row below nan row
    if df.empty:
        raise NoWeekBefore
    total = df['SUMMARY']["Q'ty"].sum()
    result = pd.DataFrame(columns=df.columns, index=[tail_index])
    for first, second in result.columns:
        if first == 'SUMMARY':
            if second == "Q'ty":
                result.at[tail_index, (first, second)] = total
            else:
                result.at[tail_index, (first, second)] = sum(df[first][second] * df['SUMMARY']["Q'ty"]) / total
        elif first == 'DATETIME':
            if second in ['month', 'shipping_date']:
                result.at[tail_index, (first, second)] = np.nan
            elif second == 'week':
                result.at[tail_index, (first, second)] = df[('DATETIME', 'month')][head_index]
        elif first == 'OOC':
            result.at[tail_index, (first, second)] = np.nan
        else:
            result.at[tail_index, (first, second)] = sum(df[first][second] * df['SUMMARY']["Q'ty"]) / total
    return result


def df_summary_table(init=False):
    """

    :param init: determine whether return empty Weekly summary table
    :return: if not init, return this week summary table df
    """
    date = current_date
    df = summary_table_header
    if not init:
        for product_group, part, goal in df.index:
            if product_group == 'Item':
                df.at[(product_group, part, goal), 0] = 'Baseline'
                df.at[(product_group, part, goal), 1] = last_month(date)
                df.at[(product_group, part, goal), 2] = week_number(week_check_day(date))
            elif goal != "Q'ty":
                df.at[(product_group, part, goal), 0] = current_bsl_yield.get((part, 'Baseline yield'), np.nan)
                df.at[(product_group, part, goal), 1] = summary_cache.get((part, 'Last month yield'), np.nan)
                df.at[(product_group, part, goal), 2] = summary_cache.get((part, 'Yield'), np.nan)
            else:
                df.at[(product_group, part, goal), 1] = summary_cache.get((part, 'Last month qty'), np.nan)
                df.at[(product_group, part, goal), 2] = summary_cache.get((part, "Q'ty"), np.nan)
    # for part in summary_table_parts:
    #     pass

    return df.T


def ooc_parts(part):
    """
    check part need to generate ooc data or not
    :param part: part
    :return: bool value
    """
    if part in ooc_part_list and part in ooc_reference:
        return True
    else:
        return False


def test_func(p, folder):
    df1 = df_example_2(p, x=1000, random_content=True, random_size=0.1)
    chk_dir(dump_data_root_folder + "\\{}".format(folder))
    df1.to_excel(dump_data_root_folder + "\\{}\\{}_test.xlsx".format(folder, p))
    # print('dump to {}'.format(p))
    return p + 'Done'


# def main():
#     # part = 'AP196'
#     # template(part)
#
#     # df = df_weekly_update(df_example_2(part, x=10), part, summary=False)
#     # print(df.to_string())
#
#     # df = df_code_update(df_example_2(part, x=10), part)
#     # print(df.to_string())
#     # print(ooc_identifier(df, part))
#     # print(df.to_string())
#
#     # df = df_weekly_update(df_example_2(part, x=100, random_content=True, random_size=0.05), part, summary=True) # generate bsl checker
#     # print(df.to_string())
#     # print(bsl_identifier(df, part))
#
#     # import datetime as dt # generate test dummy excel
#     # from DefGroup_update import df_update_xlsx
#     # date = dt.date(year=2018, month=12, day=7)
#     # df = df_month_summary_example(part, date)
#     # print(df.to_string())
#     # df_update_xlsx(df, 'test_month', 'test_month', offset=1)
#     # return df
#
#     # filename = 'test_month.xlsx'
#     # df = df_monthly_summary(part, filename)
#     # print(df.to_string())
#     # a = df_formula_dict('AP1C0', 2)
#     from multi_thread import MultiProcessGenerator
#     import timeit
#
#     chk_dir(dump_data_root_folder + "\\normal_test_data")
#     part1 = ['AP196', 'AP174', 'AP197', 'AP1C0', 'AP1CC00', 'AP1CC01', 'AP1C5', 'AP19600', 'AP19601', 'AP19602',
#              'AP19603', 'AP19700', 'AP19701', 'AP19702', 'AP17400', 'AP17406']
#
#     parts = part1 + part1 + part1 + part1
#     start_time = timeit.default_timer()
#     for part in parts:
#         test_func(part, 'normal_test_data')
#
#     elapsed = timeit.default_timer() - start_time  # try to measure the time for each loop
#     finish_line = 'normal way: elapsed time:{} seconds'.format(elapsed)
#     print(finish_line)
#     chk_dir(dump_data_root_folder + "\\multi_test_data")
#     start_time2 = timeit.default_timer()
#
#     multi = MultiProcessGenerator(test_func, parts, worker_num=16, func_args=('multi_test_data',))
#     multi.run()
#
#     elapsed2 = timeit.default_timer() - start_time2
#     finish_line2 = 'multi way: elapsed time:{} seconds'.format(elapsed2)
#     print(finish_line2)
#
#
# from concurrent_test import ConcurrentTaskGen
# import timeit
#
#
# def main1():
#     chk_dir(dump_data_root_folder + "\\normal_test_data")
#     part1 = ['AP196', 'AP174', 'AP197', 'AP1C0', 'AP1CC00', 'AP1CC01', 'AP1C5', 'AP19600', 'AP19601', 'AP19602',
#              'AP19603', 'AP19700', 'AP19701', 'AP19702', 'AP17400', 'AP17406']
#
#     parts = part1 + part1 + part1 + part1
#
#     chk_dir(dump_data_root_folder + "\\Concurrent_test_data")
#     start_time2 = timeit.default_timer()
#
#     multi = ConcurrentTaskGen(test_func, parts, worker_num=2, func_args=('Concurrent_test_data',))
#     a = multi.run()
#
#     elapsed2 = timeit.default_timer() - start_time2
#     finish_line2 = 'multi way: elapsed time:{} seconds'.format(elapsed2)
#     print(finish_line2)
#
#     start_time = timeit.default_timer()
#     for part in parts:
#         test_func(part, 'normal_test_data')
#
#     elapsed = timeit.default_timer() - start_time  # try to measure the time for each loop
#     finish_line = 'normal way: elapsed time:{} seconds'.format(elapsed)
#     print(finish_line)


# def main2():
#     from concurrent_test import ConcurrentProcessesGen
#     import timeit
#
#     def test_worker_num_for_concurrent_task(max_worker_num, function, var_list, func_args):
#         print('worker_num, process time, vars_len')
#         for n in range(max_worker_num):
#             start_time = timeit.default_timer()
#             multi = ConcurrentProcessesGen(function, var_list, worker_num=n + 1, func_args=func_args)
#             multi.run()
#             elapsed = timeit.default_timer() - start_time
#             print('{}, {}, {}'.format(n + 1, elapsed, len(var_list)))
#
#     part1 = ['AP196', 'AP174', 'AP197', 'AP1C0', 'AP1CC00', 'AP1CC01', 'AP1C5', 'AP19600', 'AP19601', 'AP19602',
#              'AP19603', 'AP19700', 'AP19701', 'AP19702', 'AP17400', 'AP17406']
#
#     parts = part1 + part1 + part1 + part1
#     test_worker_num_for_concurrent_task(20, test_func, parts, func_args=('Concurrent_test_data',))


if __name__ == '__main__':
    pass
