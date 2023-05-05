#!/usr/bin/env python
# _*_ coding:utf-8 _*_
from __future__ import division

import datetime as dt
import logging
# import math
import operator
import os
from functools import wraps
from itertools import product
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image
from wx import Colour

# from pyparsing import Literal, CaselessLiteral, Word, Combine, Group, Optional, ZeroOrMore, Forward, nums, alphas, \
#     oneOf, ParseException
import Template_generator as Tg
import wm_constants as wm_const
from Configuration import cfg
from Database_query import cp_weekly, fsbs_insp_weekly, recp_query


class DataNotFoundError(Exception):

    def __init__(self):
        self.value = u"Can't find input lots/wafers, please confirm the input list is correct , or cancel '\u8acb\u8f38\u5165FINAL YIELD PASS DIE', or ensure input lotID has wafer in it."

    def __str__(self):
        return self.value


class OperatorWrongError(Exception):

    def __init__(self, op):
        self.value = f"Input operator {op} is not mathematical operator, please reconfirm in the setting"

    def __str__(self):
        return self.value


# class NumericStringParser(object):
#     """https://stackoverflow.com/questions/57234319/pyparsing-how-to-parse-string-with-comparison-operators"""
#
#     def __init__(self):
#         self.exprStack = []
#         point = Literal(".")
#         e = CaselessLiteral("E")
#         fnumber = Combine(Word("+-" + nums, nums) +
#                           Optional(point + Optional(Word(nums))) +
#                           Optional(e + Word("+-" + nums, nums)))
#         ident = Word(alphas, alphas + nums + "_$")
#         # diffop = Literal('<=') | Literal('>=')
#         diffop = oneOf("<= >= > <")
#         compop = oneOf("== != =")
#         plus = Literal("+")
#         minus = Literal("-")
#         mult = Literal("*")
#         floordiv = Literal("//")
#         div = Literal("/")
#         mod = Literal("%")
#         lpar = Literal("(").suppress()
#         rpar = Literal(")").suppress()
#         addop = plus | minus
#         multop = mult | floordiv | div | mod
#         expop = Literal("^")
#         pi = CaselessLiteral("PI")
#         tau = CaselessLiteral("TAU")
#         expr = Forward()
#         atom = ((Optional(oneOf("- +")) +
#                  (ident + lpar + expr + rpar | pi | e | tau | fnumber).setParseAction(self.__push_first__))
#                 | Optional(oneOf("- +")) + Group(lpar + expr + rpar)
#                 ).setParseAction(self.__push_minus__)
#
#         factor = Forward()
#         factor << atom + \
#         ZeroOrMore((expop + factor).setParseAction(self.__push_first__))
#         term = factor + \
#                ZeroOrMore((multop + factor).setParseAction(self.__push_first__))
#         arith_expr = term + \
#                      ZeroOrMore((addop + term).setParseAction(self.__push_first__))
#         relational = arith_expr + \
#                      ZeroOrMore((diffop + arith_expr).setParseAction(self.__push_first__))
#         expr <<= relational + \
#                  ZeroOrMore((compop + relational).setParseAction(self.__push_first__))
#
#         self.bnf = expr
#
#         self.opn = {
#             "+": operator.add,
#             "-": operator.sub,
#             "*": operator.mul,
#             "/": operator.truediv,
#             "//": operator.floordiv,
#             "%": operator.mod,
#             "^": operator.pow,
#             "=": operator.eq,
#             "==": operator.eq,
#             "!=": operator.ne,
#             "<=": operator.le,
#             ">=": operator.ge,
#             "<": operator.lt,
#             ">": operator.gt
#         }
#
#         self.fn = {
#             "sin": math.sin,
#             "cos": math.cos,
#             "tan": math.tan,
#             "asin": math.asin,
#             "acos": math.acos,
#             "atan": math.atan,
#             "exp": math.exp,
#             "abs": abs,
#             "sqrt": math.sqrt,
#             "floor": math.floor,
#             "ceil": math.ceil,
#             "trunc": math.trunc,
#             "round": round,
#             "fact": math.factorial,
#             "gamma": math.gamma
#         }
#
#     def __push_first__(self, strg, loc, toks):
#         self.exprStack.append(toks[0])
#
#     def __push_minus__(self, strg, loc, toks):
#         if toks and toks[0] == "-":
#             self.exprStack.append("unary -")
#
#     def __evaluate_stack__(self, s):
#         op = s.pop()
#         if op == "unary -":
#             return -self.__evaluate_stack__(s)
#         if op in ("+", "-", "*", "//", "/", "^", "%", "!=", "<=", ">=", "<", ">", "=", "=="):
#             op2 = self.__evaluate_stack__(s)
#             op1 = self.__evaluate_stack__(s)
#             return self.opn[op](op1, op2)
#         if op == "PI":
#             return math.pi
#         if op == "E":
#             return math.e
#         if op == "PHI":
#             return (1 + math.sqrt(5)) / 2
#         if op == "TAU":
#             return math.tau
#         if op in self.fn:
#             return self.fn[op](self.__evaluate_stack__(s))
#         if op[0].isalpha():
#             raise NameError(f"{op} is not defined.")
#         return float(op)
#
#
# def evaluate(expression, parse_all=True):
#     nsp = NumericStringParser()
#     nsp.exprStack = []
#     try:
#         nsp.bnf.parseString(expression, parse_all)
#     except ParseException as error:
#         raise SyntaxError(error)
#     return nsp.__evaluate_stack__(nsp.exprStack[:])




def quantileNormalize(df_input):
    df = df_input.copy()
    # compute rank
    dic = {}
    for col in df:
        dic.update({col: sorted(df[col])})
    sorted_df = pd.DataFrame(dic)
    rank = sorted_df.mean(axis=1).tolist()
    # sort
    for col in df:
        t = np.searchsorted(np.sort(df[col]), df[col])
        df[col] = [rank[i] for i in t]
    return df


def nanpercentile(a, percentile):
    """
    Perform ``numpy.percentile(a, percentile)`` while ignoring NaN values.

    Only works on a 1D array.
    """
    if type(a) != np.ndarray:
        a = np.array(a, dtype=np.float64)
    return np.percentile(a[np.logical_not(np.isnan(a))], percentile)


def ensure_content_not_empty(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        if self.empty:
            raise DataNotFoundError
        else:
            result = func(self, *args, **kwargs)
            return result

    return wrapper


def set_color(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        wm_const.wm_HIGH_COLOR.Set(kwargs['high_color'])
        wm_const.wm_LOW_COLOR.Set(kwargs['low_color'])
        wm_const.wm_OOR_HIGH_COLOR.Set(kwargs['oos_high_color'])
        wm_const.wm_OOR_LOW_COLOR.Set(kwargs['oos_low_color'])
        result = func(self, *args, **kwargs)
        return result

    return wrapper


# def set_color(default):
#     def wrap(func):
#         @wraps(func)
#         def wrapper(self, *args, **kwargs):
#             if default:
#                 wm_const.wm_HIGH_COLOR.Set(*cfg['map info']['default color(RGB)']['High color'])
#                 wm_const.wm_LOW_COLOR.Set(*cfg['map info']['default color(RGB)']['Low color'])
#                 wm_const.wm_OOR_HIGH_COLOR.Set(*cfg['map info']['default color(RGB)']['OOS high color'])
#                 wm_const.wm_OOR_LOW_COLOR.Set(*cfg['map info']['default color(RGB)']['OOS low color'])
#             else:
#                 wm_const.wm_HIGH_COLOR.Set(kwargs['high_color'])
#                 wm_const.wm_LOW_COLOR.Set(kwargs['low_color'])
#                 wm_const.wm_OOR_HIGH_COLOR.Set(kwargs['oos_high_color'])
#                 wm_const.wm_OOR_LOW_COLOR.Set(kwargs['oos_low_color'])
#             result = func(self, *args, **kwargs)
#             return result
#
#         return wrapper
#
#     return wrap


def set_to_default_color(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        wm_const.wm_HIGH_COLOR.Set(*cfg['map info']['default color(RGB)']['High color'])
        wm_const.wm_LOW_COLOR.Set(*cfg['map info']['default color(RGB)']['Low color'])
        wm_const.wm_OOR_HIGH_COLOR.Set(*cfg['map info']['default color(RGB)']['OOS high color'])
        wm_const.wm_OOR_LOW_COLOR.Set(*cfg['map info']['default color(RGB)']['OOS low color'])
        result = func(self, *args, **kwargs)
        return result

    return wrapper


def chk_jpg_and_convert_to_png(root_dir):
    path = Path(root_dir).glob('**/*')
    counter = 0
    for p in path:
        if p.is_file():
            if p.suffix == '.jpg':
                new_p = Path(f"{p.parent.as_posix()}/{p.stem}.png")
                if not new_p.is_file():
                    img = Image.open(p)
                    print(f"Convert {p} to png")
                    img.save(new_p)
                    counter += 1

    if counter:
        print(f"Convert {counter} jpg files to png")
    else:
        print('no jpg file to convert')


def df_xls_color(df_die_map, df_bin_map):
    iter_list = list(product([-1, 0, 1], [-1, 0, 1]))
    bgcol = df_bin_map.replace(df_bin_map, 0)
    ori_bgcol = bgcol.copy()
    ori_die, add_die = (set() for l in range(2))
    iter_lists = iter_list
    columns = df_bin_map.columns
    indexes = df_bin_map.index
    for idx in indexes:
        for col in columns:
            bin_code = df_bin_map.at[(idx, col)]
            if bin_code == '9':
                ori_bgcol.at[(idx, col)] = 1
                die = df_die_map.at[(idx, col)]
                ori_die.add(die)
                for r, c in iter_lists:
                    rowplus = idx + r
                    colplus = col + c
                    try:
                        origin = bgcol.at[(rowplus, colplus)]
                        bgcol.at[(rowplus, colplus)] = 1
                        die = df_die_map.at[(rowplus, colplus)]
                        if pd.notnull(die):
                            add_die.add(die)
                    except KeyError:
                        continue

    return {'add_color': bgcol,
            'origin_color': ori_bgcol, 'add_die': add_die, 'origin_die': ori_die}


def max_bin_by_wafer(lot_id, wafer_id, part):
    df_cp = cp_weekly(wafer_id)
    try:
        df_cp.drop(['1', '*'], inplace=True)
    except KeyError:
        df_cp.drop(['1'], inplace=True)
    df_insp = fsbs_insp_weekly(wafer_id)
    df_cp.columns, df_insp.columns = (['COUNT'] for l in range(2))
    df = pd.concat([df_cp, df_insp], sort=True)
    df = df.loc[df.idxmax()]
    des_dict = Tg.description_dict[part]
    df['NAME'] = df.index.map(lambda a: des_dict[a])
    df['WAFER_ID'] = wafer_id
    df['LOT_ID'] = lot_id
    df['PART'] = part
    df['COUNT'] = df['COUNT'].astype('int64')
    return df


def current_time_stamp():
    return dt.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')


def recipe_query_by_title_des_cap(title, des, cap):
    print('Querying...')
    df = recp_query(title, des, cap)
    print('Finish!')
    return (
        'open table viewer, no alternated color', {'title': 'recipe information', 'df': df})


def df_join_and_drop(df_left, df_right, drop_cols=None, on=None, index=False):
    if index:
        df = df_left.merge(df_right, how='outer', on=on, left_index=index, right_index=index)
    else:
        df = df_left.reset_index().merge(df_right, how='outer', on=on, left_index=index, right_index=index).set_index(
            df_left.index.name)
    if drop_cols is not None:
        df.dropna(inplace=True, subset=drop_cols)
    df.fillna('1', inplace=True)
    return df


def df_bin_normalize(bin_code, bin_set: set):
    if bin_set:
        if bin_code in bin_set:
            return 1
        else:
            return 0
    elif bin_code == '1':
        return 0
    else:
        return 1


def color_val_gen(n, side):
    space = int(255 / n)
    if side == 'fs_bin':
        colors = [Colour(0 + space * i, 255, 255) for i in range(n)]  # cyan blue
    elif side == 'bs_bin':
        colors = [Colour(0 + space * i, 255, 0 + space * i) for i in range(n)]  # green
    elif side == 'cp_bin':
        colors = [Colour(255, 255, 0 + space * i) for i in range(n)]  # yellow
    elif side == 'CP':  # for item pass/fail map
        colors = [Colour(255, 0 + space * i, 0 + space * i) for i in range(n)]  # red
    else:
        colors = []
    return colors


def plot_range_determination(df_map, sp_usl, sp_lsl, use_dpat, usl_perc, lsl_perc):
    if sp_usl and sp_lsl:
        plot_range = (float(sp_lsl), float(sp_usl))
    elif use_dpat == "Yes":
        q1 = nanpercentile(df_map.DATA, 25)
        q2 = nanpercentile(df_map.DATA, 50)
        q3 = nanpercentile(df_map.DATA, 75)
        q3_q1 = q3 - q1
        robust_sigma = q3_q1 / 1.35
        pat_upper = q2 + 6 * robust_sigma
        pat_lower = q2 - 6 * robust_sigma
        s = df_map.DATA
        arr = s[(s > pat_lower) & (s < pat_upper)]
        plot_range = (min(arr), max(arr))
    elif usl_perc and lsl_perc:
        p_up = float(nanpercentile(df_map.DATA, int(usl_perc)))
        p_down = float(nanpercentile(df_map.DATA, int(lsl_perc)))
        if p_down and p_up:  # add by Ernest, ensure not divide by 0
            plot_range = (p_down, p_up)
        else:
            data_min = min(df_map.DATA)
            data_max = max(df_map.DATA)
            plot_range = (data_min, data_max)

    return plot_range


def search_aoi_folder(wafer_id):
    """search folder that contain wafer_id str folder"""
    lot_id, wafer = wafer_id.split('.')
    wafer = f"{int(wafer):.0f}".zfill(2)
    wafer_id = f"{lot_id}#{wafer}"
    aoi1 = '\\\\192.168.7.27\\\\aoi1'
    aoi3 = '\\\\192.168.7.27\\\\aoi3'
    aoi4 = '\\\\192.168.7.28\\\\aoi4'
    aoi5 = '\\\\192.168.7.29\\\\aoi5'
    result_d = {}
    for root_dir in [aoi1, aoi3, aoi4, aoi5]:
        try:
            lst = [f for f in os.listdir(root_dir) if f.split('_')[0] == wafer_id if
                   os.path.isdir(os.path.join(root_dir, f))]
        except OSError as e:
            try:
                logging.exception('An exception was thrown!')
                continue
            finally:
                e = None
                del e

        if lst:
            for f in lst:
                result_d[f] = os.path.join(root_dir, f)

    print('AOI Folder:' + str(result_d))
    return result_d


def consecutive_duplicate_idx_select(idx):
    length = len(idx)
    record_range = False
    left = []
    right = []
    for i, tup in enumerate(idx):
        past = i + 1
        if past < length:
            same = bool(idx[i] == idx[past])
        else:
            same = False
        if same:
            if not record_range:
                record_range = True
                left.append(i)
            if not same:
                if record_range:
                    record_range = False
                    right.append(i)

    return zip(left, right)


def split_list(alist, wanted_parts=1):
    length = len(alist)
    return [alist[i * length // wanted_parts: (i + 1) * length // wanted_parts]
            for i in range(wanted_parts)]


def description_gen(items, des_dict):
    """for cp item pass/fail map description gen"""
    item_list = items.split(' ')
    if len(item_list) > 1:
        des_list = [des_dict.get(i, 'Others') for i in item_list]
        result = '+'.join(des_list)
    elif item_list[0] == 'Pass':
        result = 'Pass die'
    else:
        result = des_dict.get(item_list[0], 'Others')
    return result


def evaluate_func(data, op, spec):
    opn = {
        "+": operator.add,
        "-": operator.sub,
        "*": operator.mul,
        "/": operator.truediv,
        "//": operator.floordiv,
        "%": operator.mod,
        "^": operator.pow,
        "=": operator.eq,
        "==": operator.eq,
        "!=": operator.ne,
        "<=": operator.le,
        ">=": operator.ge,
        "<": operator.lt,
        ">": operator.gt
    }
    a = float(data)
    b = float(spec)
    try:
        result = opn[op](a, b)
    except KeyError:
        raise OperatorWrongError(op)
    return result
