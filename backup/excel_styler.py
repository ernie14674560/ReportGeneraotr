#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import matplotlib
import numpy as np
import pandas as pd
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill
from wx import C2S_HTML_SYNTAX

import Template_generator as Tg


# from Template_generator import bsl_identifier

class WeeklyHeaderStyles(NamedStyle):
    """
    Weekly header exclusive style
    """
    bd = Side(style="thin", color="000000")
    color_dict = {'DATETIME': 'FFA500', 'WAFER': 'FFD700', 'SUMMARY': 'FCD5B4', 'FSV': 'CCFFFF', 'BSV': 'CCFFCC',
                  'CP': 'FFFF99', 'OOC': 'FFC0CB', 'OOC_GROUP': 'FFC0CB'}

    def __init__(self,
                 name="Normal",
                 font=Font(bold=True, size=12),
                 border=Border(left=bd, top=bd, right=bd, bottom=bd),
                 ):
        super().__init__()
        self.name = name
        self.fill = PatternFill("solid", fgColor=self.color_dict.get(self.name, 'FFFFFF'))
        self.font = font
        self.border = border


def cmap_map(function, cmap):
    """ Applies function (which should operate on vectors of shape 3: [r, g, b]), on colormap cmap.
    This routine will break any discontinuous points in a colormap.
    """
    cdict = cmap._segmentdata
    step_dict = {}
    # Firt get the list of points where the segments start or end
    for key in ('red', 'green', 'blue'):
        step_dict[key] = list(map(lambda x: x[0], cdict[key]))
    step_list = sum(step_dict.values(), [])
    step_list = np.array(list(set(step_list)))
    # Then compute the LUT, and apply the function to the LUT
    reduced_cmap = lambda step: np.array(cmap(step)[0:3])
    old_LUT = np.array(list(map(reduced_cmap, step_list)))
    new_LUT = np.array(list(map(function, old_LUT)))
    # Now try to make a minimal segment definition of the new LUT
    cdict = {}
    for i, key in enumerate(['red', 'green', 'blue']):
        this_cdict = {}
        for j, step in enumerate(step_list):
            if step in step_dict[key]:
                this_cdict[step] = new_LUT[j, i]
            elif new_LUT[j, i] != old_LUT[j, i]:
                this_cdict[step] = new_LUT[j, i]
        colorvector = list(map(lambda x: x + (x[1],), this_cdict.items()))
        colorvector.sort()
        cdict[key] = colorvector

    return matplotlib.colors.LinearSegmentedColormap('colormap', cdict, 1024)


def history_colorization(cell_value, column_name, color_dict):
    style_str = 'background-color:{}'
    item = column_name[1]
    if item in {"Stage", "Equipment", "Procedure", "TrackInUser", "TrackOutUser", "LOTID"}:
        return style_str.format(color_dict[item].get(cell_value, ''))
    elif item == 'Events':
        if pd.notna(cell_value):
            if "HOLDLOT" in cell_value:
                return style_str.format("#FFC7CE")
            else:
                return style_str.format("#C6EFCE")
        else:
            return ''
    else:
        return ''


def history_highlight(df):
    idx = pd.IndexSlice
    color_dict = {}
    light_jet = cmap_map(lambda x: x / 2 + 0.7, matplotlib.cm.jet)

    for item in {"Stage", "Equipment", "Procedure", "TrackInUser", "TrackOutUser", "LOTID"}:
        unique_items = pd.unique(df.loc[:, idx[:, item]].values.ravel('K'))
        unique_items = unique_items[~pd.isnull(unique_items)].tolist()
        unique_items.sort()
        length = len(unique_items)
        norm = matplotlib.colors.Normalize(vmin=0, vmax=length)
        color_dict2 = {}
        for i, unique in enumerate(unique_items):
            rgb = light_jet(norm(i))[:3]
            hex_color = matplotlib.colors.rgb2hex(rgb)
            color_dict2[unique] = hex_color
        color_dict[item] = color_dict2
    df_ref = df.apply(lambda x: x.apply(history_colorization, args=(x.name, color_dict)))
    # data = {col: '' for col in df.columns}
    # df_ref = pd.DataFrame(data=data, columns=df.columns, index=df.index)
    # for i in df_ref.index:
    #     for wafer in df_ref.columns.levels[0]:
    #         for item in lst:
    #             string = df.at[i, (wafer, item)]
    #             df_ref.at[i, (wafer, item)] = style_str.format(color_dict[item].get(string, ''))
    #         strings = df.at[i, (wafer, 'Events')]
    #         if pd.notna(strings) and "HOLDLOT" in strings:
    #             df_ref.at[i, (wafer, 'Events')] = style_str.format("#FFC7CE")
    return df_ref


def bsl_highlight(df, part, df_ref=pd.DataFrame(), font_color='red',
                  font_style='oblique', font_weight='bold',
                  background_color='#FFFFFF'):
    style_str = 'color:{}; font-style:{}; font-weight:{}; background-color:{}'
    if ('DATETIME', 'month') in df:
        highlight_row_index = df[df[('DATETIME', 'month')].isnull()].index  # use excel blank cell as indicator
    else:
        highlight_row_index = []
    sub_bg_color = 'background-color:{}'.format(background_color)
    new_bg_color = 'background-color:Yellow'
    if part in Tg.bsl_parts and not df_ref.empty:  # chk the BSL function and baseline.xlsx is correct
        cell_style = style_str.format(font_color, font_style, font_weight, background_color)
        df_ref = df_ref.applymap(lambda x: cell_style if x >= 0 else '')
    else:
        data = {col: '' for col in df.columns}
        df_ref = pd.DataFrame(data=data, columns=df.columns, index=df.index)
    for idx in highlight_row_index:
        for col in df_ref.columns:
            strings = df_ref.at[idx, col]
            if sub_bg_color in strings:
                df_ref.at[idx, col] = strings.replace(sub_bg_color, new_bg_color)
            else:
                df_ref.at[idx, col] += new_bg_color
    return df_ref


def color_map(x, ori_style, add_style):
    if x == 2:
        return ori_style
    elif x == 1:
        return add_style
    else:
        return ''


def tco_highlight(df, df_ref_dict, add_color='Yellow', ori_color='Red'):
    style_str = 'background-color:{}'
    df_ref_add = df_ref_dict['add_color']
    df_ref_ori = df_ref_dict['origin_color']
    df = df.applymap(lambda x: True if pd.notnull(x) else False)
    add_cell_style = style_str.format(add_color)
    ori_cell_style = style_str.format(ori_color)
    df_ref = df_ref_add + df_ref_ori
    df_result = df * df_ref

    df_result = df_result.applymap(lambda x: color_map(x, ori_cell_style, add_cell_style))

    return df_result


def wafer_color_map(x, color_func):
    style_str = 'background-color:{}'
    if pd.notnull(x):
        # rgb_color = color_func(x)
        # hex_color = rgb2hex(rgb_color, force_long=True, multiply_255=False)

        wx_colour = color_func(x)
        hex_color = wx_colour.GetAsString(flags=C2S_HTML_SYNTAX)

        return style_str.format(hex_color)
    else:
        return ''


def wafer_map_styler(df, color_func):
    df_result = df.applymap(lambda x: wafer_color_map(x, color_func))
    return df_result


def fsv_bsv_cp(x, part):
    from Template_generator import part_group_code_dict
    base = 'background-color: {}'
    fsv_color = base.format('#CCFFFF')
    bsv_color = base.format('#CCFFCC')
    cp_color = base.format('#FFFF99')
    adjustment_source = part_group_code_dict[part]
    fsv_list = list(adjustment_source['正檢'])
    bsv_list = list(adjustment_source['背檢'])
    cp_list = list(adjustment_source['CP'])
    fsv_bsv_duplicate = list(set().union(fsv_list, bsv_list))
    chk_duplicate = []
    if x in ["Q'ty", "Yield", "Yield loss"]:
        x = "font-weight: bold"
    elif x in fsv_bsv_duplicate:
        if x in chk_duplicate:
            x = bsv_color
        chk_duplicate.append(x)
    elif x in fsv_list or x == 'FSV':
        x = fsv_color
    elif x in bsv_list or x == 'BSV':
        x = bsv_color
    elif x in cp_list or x == 'CP':
        x = cp_color
    return x


if __name__ == '__main__':
    pass
