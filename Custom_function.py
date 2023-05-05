# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Custom_function.py'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

# uncompyle6 version 3.5.0
# Python bytecode 3.7 (3394)
# Decompiled from: Python 2.7.5 (default, Aug  7 2019, 00:51:29)
# [GCC 4.8.5 20150623 (Red Hat 4.8.5-39)]
# Embedded file name: C:\Users\ernie.chen\Desktop\Project\Project Week\Custom_function.py
# Size of source mod 2**32: 37877 bytes


import logging
import os
from collections import Counter

import numpy as np
import pandas as pd
from wx import Colour

import Template_generator as Tg
from Configuration import cfg
from Custom_parser import ensure_content_not_empty, quantileNormalize, df_xls_color, max_bin_by_wafer, \
    current_time_stamp, df_join_and_drop, df_bin_normalize, color_val_gen, plot_range_determination, \
    consecutive_duplicate_idx_select, split_list, description_gen, evaluate_func, set_color, set_to_default_color
from Database_query import connection, lot_list_summary, map_info, insp_map, cp_map, \
    lot_to_lots, cp_value_map, wafer_history, lot_event_search, cp_summary
from Presentation_update import monthly_report_gen
from Weekly_parser import parts_lots_arrange
from Weekly_update import weekly_writer
from df_to_excel import df_update_xlsx, cell_lot_history_table
from excel_styler import tco_highlight, history_highlight

unnecessary_dies_dict = cfg['map info']['unnecessary dies in map']


def df_holder_apply_error_handler(func, args=None, kwargs=None):
    if args is None:
        args = []
    if kwargs is None:
        kwargs = {}
    try:
        return DfHolder(func(*args, **kwargs))
    except Exception as e:
        print(e)
        return np.nan


class DfHolder:

    def __init__(self, df):
        self.df = df


class CPItemSpecNotFoundError(Exception):

    def __init__(self, part):
        s = f"Can't find cp item spec of part {part}, please add it in the setting(map info->CP item pass spec).."
        self.value = s

    def __str__(self):
        return self.value


class CPItemDescriptionNotFoundError(Exception):

    def __init__(self, part):
        s = f"Can't find cp item description of part {part}, please add it in the setting(map info->CP item pass spec)."
        self.value = s

    def __str__(self):
        return self.value


class MoreThanOneWaferError(Exception):

    def __init__(self):
        self.value = 'Input lots have more than one wafer, please input one wafer to proceed.'

    def __str__(self):
        return self.value


class MoreThanOnePartError(Exception):

    def __init__(self):
        self.value = 'Input lots have more than one part, please ensure only input one part to proceed.'

    def __str__(self):
        return self.value


class OneWaferError(Exception):

    def __init__(self):
        self.value = 'Input lot only has one wafer, please input more wafer to proceed.'

    def __str__(self):
        return self.value


class WaferStackedDataZeroError(Exception):

    def __init__(self):
        self.value = 'Input lots stacked map data is zero, please input another lot to proceed.'

    def __str__(self):
        return self.value


class NotLotIdError(Exception):

    def __init__(self):
        self.value = 'Input id is not Lot id, please input lot id to continue.'

    def __str__(self):
        return self.value


class DescriptionNotFoundError(Exception):

    def __init__(self):
        self.value = "Can't find the bin code description, please check"

    def __str__(self):
        return self.value


class AOIFolder:
    def __init__(self):
        aoi_dirs = ['\\\\192.168.7.27\\\\aoi1', '\\\\192.168.7.27\\\\aoi3', '\\\\192.168.7.28\\\\aoi4',
                    '\\\\192.168.7.29\\\\aoi5']
        data = {}
        for root_dir in aoi_dirs:
            try:
                lst = os.listdir(root_dir)
                data[root_dir] = lst
            except OSError as e:
                try:
                    logging.exception('An exception was thrown!')
                    continue
                finally:
                    e = None
                    del e
        self.df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in data.items()])).fillna('')

    def search_folder(self, wafer_id):
        lot_id, wafer = wafer_id.split('.')
        wafer = f"{int(wafer):.0f}".zfill(2)
        wafer_id = f"{lot_id}#{wafer}"
        df = self.df
        result_d = {}
        for root_dir in df:
            arr = df[df[root_dir].str.contains(wafer_id)][root_dir]
            for i, f in arr.iteritems():
                path = os.path.join(root_dir, f)
                if os.path.isdir(path):
                    result_d[f] = path
        print('AOI Folder:' + str(result_d))
        return result_d


class InputLotListToGet:
    """class that parse input lot or wafer list to desire information.
    the class method that return ('purpose', data) is to inform GUI how to deal with returned result
        """
    connection()
    dump_path = Tg.dump_data_root_folder

    def __init__(self, column_str, wafer_id=False, first_5_char=True, final_yield=True, wafer_map=False):
        """

        :param column_str: input excel column like string
        :param wafer_id: input id is wafer id or lot id
        :param first_5_char: part only account first 5 character like AP1CC, AP196... etc
        """
        self._lot_list = []
        self.wafer_id = wafer_id
        if column_str:
            for s in column_str.splitlines():
                lot = s.strip()
                if lot:
                    if not lot.isupper():
                        lot = lot.upper()
                    if len(lot) == 7:
                        self._lot_list += lot_to_lots(lot)
                    else:
                        self._lot_list.append(lot)
            split_num = (len(self._lot_list) // 1000) + 1
            list_of_lists = split_list(self._lot_list, wanted_parts=split_num)
            concat_list = []
            for lst in list_of_lists:
                df_list = lot_list_summary(lst, wafer_id=wafer_id, final_yield=final_yield)
                concat_list.append(df_list)
            self._df_list = pd.concat(concat_list, sort=False)
            # self._df_list = lot_list_summary(self._lot_list, wafer_id=wafer_id, final_yield=final_yield)
            self._lot_list_info = parts_lots_arrange(self._df_list, first_5_char=first_5_char)
            if self._df_list.empty:
                self.empty = True
            else:
                self.empty = False
        else:
            self.empty = True
            self._df_list = pd.DataFrame()
            self._lot_list_info = {'lots': [], 'parts': []}
        self._parts_list = self._lot_list_info['parts']
        self._lots_dict = self._lot_list_info['lots']
        self.wafer_map = wafer_map
        self.die_size_dict = cfg['map info']['die size in mm']
        # if wafer_map:
        #     self._map_info = {part: map_info(part, cp_stage) for part in self._parts_list}
        # else:
        self._map_info = {}
        self.lots_events = {}

    @ensure_content_not_empty
    def summary_data(self, return_df_dict=False):
        print('Querying summary data')
        if return_df_dict:
            df_dict = {}
            wafer_list_dict = {}
        else:
            df_dict = None
            wafer_list_dict = None
        for part in self._parts_list:
            information = self._lots_dict.get(part)
            path = self.dump_path + '\\Wafer_lot_summary\\{}_data'.format(current_time_stamp())
            Tg.chk_dir(path)
            file = '\\{}.xlsx'.format(part)
            filepath = path + file
            if return_df_dict:
                df = weekly_writer(part, filepath, cross_month=False, cross_year=False, weekly=False, date=None,
                                   product_group=None, ignore_db_error=True,
                                   shipping_information=information,
                                   return_df=True)
                df_dict[part] = df
                wafer_list_dict[part] = information['COMPIDS'].tolist()
            else:
                weekly_writer(part, filepath, cross_month=False, cross_year=False, weekly=False, date=None,
                              product_group=None, ignore_db_error=True,
                              shipping_information=information)

        if return_df_dict:
            return df_dict, wafer_list_dict
        else:
            return 'open file', filepath

    def modify_monthly_data_and_report(self, date, comparing_to_the_last_year):
        if not self.empty:
            sub_df_dict, wafer_list_dict = self.summary_data(return_df_dict=True)
        monthly_report_gen(date=date, ignored_parts_and_wafers=(None if self.empty else (sub_df_dict, wafer_list_dict)),
                           comparing_to_the_last_year=comparing_to_the_last_year,
                           specific_date=True)

    @staticmethod
    def _bin_code_classifier(cp_bin, fs_bin, bs_bin, bin_count_dict: dict, part: str):
        active_codes_set = Tg.part_active_code_str_set[part]
        if cp_bin == '1':
            if fs_bin == '1':
                if bs_bin == '1':
                    return '1'
                if bs_bin in active_codes_set:
                    s = bs_bin
                else:
                    s = '8888'
                bin_count_dict['bs_bin'].append(s)
            else:
                if fs_bin in active_codes_set:
                    s = fs_bin
                else:
                    s = '8887'
                bin_count_dict['fs_bin'].append(s)
        else:
            s = cp_bin
            bin_count_dict['cp_bin'].append(s)
        return s

    # leaf
    def _discrete_bin_map_info_gen(self, wafer_id, part_name, fs_bins='all', bs_bins='all', cp_bins='all',
                                   cp_stage='CP3', insp_dep='MFG'):
        part = part_name[:5]
        df = self._map_info[part].copy()
        df_cp = cp_map(wafer_id, bin_codes_tup=cp_bins, cp_stage=cp_stage)
        if not df_cp.empty:
            df = df_join_and_drop(df, df_cp, ['coord'], index=True)
        else:
            df['cp_bin'] = '1'
        df_fs = insp_map(wafer_id, side='1', bin_codes_tup=fs_bins, insp_dep=insp_dep)
        if not df_fs.empty:
            df = df_join_and_drop(df, df_fs, on='coord')
        else:
            df['fs_bin'] = '1'
        df_bs = insp_map(wafer_id, side='2', bin_codes_tup=bs_bins, insp_dep=insp_dep)
        if not df_bs.empty:
            df = df_join_and_drop(df, df_bs, on='coord')
        else:
            df['bs_bin'] = '1'
        df['X'] = df['X'].astype(int)
        df['Y'] = df['Y'].astype(int)
        return DfHolder(df)

    def _cp_value_map_info_gen(self, wafer_id, part_name, item, cp_stage):
        part = part_name[:5]
        df = self._map_info[part].copy()
        df_cp = cp_value_map(wafer_id, item, cp_stage)
        df = df_join_and_drop(df, df_cp, ['coord'], index=True)
        return DfHolder(df)

    def _cp_item_pass_map_info_gen(self, wafer_id, part_name, cp_stage, spec_dict, wafer_fail_count_dict):
        """spec_dict:                                       {'1': {'>': '1'},
                                                             '2': {'>': '1'},
                                                             '3': {'>': '1'},
                                                             '4': {'>': '1'},
                                                             '5': {'>': '1'},
                                                             '6': {'>': '1'},
                                                             '7': {'>': '1'},
                                                             '8': {'>': '1'},
                                                             '9': {'>': '1'},
                                                             '10': {'>': '1'},
                                                             '11': {'>': '1'},
                                                             '12': {'>': '1'},
                                                             '13': {'>': '1'},
                                                             '14': {'>': '1'}"""
        part = part_name[:5]

        df = self._map_info[part].copy()
        df_wafer_pass = pd.DataFrame()
        count1 = 0
        item_fail_counter = Counter()
        spec_l = sorted(spec_dict.items(), key=lambda i: int(i[0]))
        for item, pass_spec in spec_l:

            df_cp = cp_value_map(wafer_id, item, cp_stage)
            count2 = 0
            df_item_pass = pd.DataFrame()
            sr_item_pass = pd.Series(dtype=object)
            for operator, spec in pass_spec.items():  # pass condition is value satisfy all spec

                if count2:  # '' is pass, nan is fail
                    sr_item_pass += df_cp.apply(lambda row:
                                                '' if evaluate_func(row["DATA"], operator, spec) else np.nan, axis=1)
                else:  # first time
                    sr_item_pass = df_cp.apply(lambda row:
                                               '' if evaluate_func(row["DATA"], operator, spec) else np.nan, axis=1)
                count2 += 1
            df_item_pass['DATA'] = sr_item_pass
            fail_count = df_item_pass['DATA'].isna().sum()
            item_fail_counter[item] = fail_count
            df_item_pass.fillna(f'{item} ', inplace=True)
            if count1:
                df_wafer_pass += df_item_pass
            else:  # first time
                gross_die = len(df.index)
                wafer_fail_count_dict['gross_die'] = gross_die
                df_wafer_pass = df_item_pass
            count1 += 1

        wafer_fail_count_dict['wafer_id'].update(
            {wafer_id: sorted(item_fail_counter.items(), key=lambda i: int(i[0]))})
        df_wafer_pass.replace('', 'Pass', inplace=True)
        df_wafer_pass['DATA'] = df_wafer_pass['DATA'].str.rstrip()
        df = df_join_and_drop(df, df_wafer_pass, ['coord'], index=True)
        return DfHolder(df)

    @ensure_content_not_empty
    def _lots_discrete_bin_map_gen(self, fs_bins='all', bs_bins='all', cp_bins='all', cp_stage='CP3', insp_dep='MFG'):
        """empty str bin code will not query map"""
        df_lot_maps = self._df_list.copy()
        self._map_info = {part: map_info(part, cp_stage) for part in self._parts_list}
        df_lot_maps['MAP'] = df_lot_maps.apply(
            (lambda row: self._discrete_bin_map_info_gen((row['COMPIDS']), (row['PARTNAME']),
                                                         fs_bins=fs_bins,
                                                         bs_bins=bs_bins,
                                                         cp_bins=cp_bins,
                                                         cp_stage=cp_stage, insp_dep=insp_dep)), axis=1)
        return df_lot_maps

    @ensure_content_not_empty
    def _lots_cp_value_map_gen(self, item, cp_stage):
        df_lot_maps = self._df_list.copy()
        self._map_info = {part: map_info(part, cp_stage) for part in self._parts_list}
        df_lot_maps['MAP'] = df_lot_maps.apply(
            (lambda row: self._cp_value_map_info_gen((row['COMPIDS']), (row['PARTNAME']),
                                                     item=item,
                                                     cp_stage=cp_stage)),
            axis=1)
        return df_lot_maps

    def _lots_cp_item_pass_map_gen(self, cp_stage, part):
        df_lot_maps = self._df_list.copy()
        d = cfg['map info']['CP item pass spec']
        self._map_info = {part: map_info(part, cp_stage) for part in self._parts_list}
        try:
            spec_dict = d[part][cp_stage]['item spec']
        except KeyError:
            raise CPItemSpecNotFoundError(part)
        try:
            des_dict = d[part][cp_stage]['item description']
        except KeyError:
            raise CPItemDescriptionNotFoundError(part)
        wafer_fail_count_dict = {'wafer_id': {}}
        df_lot_maps['MAP'] = df_lot_maps.apply(
            (lambda row: self._cp_item_pass_map_info_gen((row['COMPIDS']), (row['PARTNAME']),
                                                         cp_stage=cp_stage, spec_dict=spec_dict,
                                                         wafer_fail_count_dict=wafer_fail_count_dict)),
            axis=1)
        return df_lot_maps, wafer_fail_count_dict, des_dict

    def _lots_continuous_bin_map_gen(self, fs_bins='all', bs_bins='all', cp_bins='all', bin_set=None, cp_stage='CP3', insp_dep='MFG'):
        df_lot_maps = self._lots_discrete_bin_map_gen(fs_bins=fs_bins, bs_bins=bs_bins, cp_bins=cp_bins,
                                                      cp_stage=cp_stage,insp_dep=insp_dep)
        total = len(df_lot_maps)
        df_maps = pd.DataFrame()
        df_coord = pd.DataFrame()
        bin_count_dict = {'fs_bin': [], 'bs_bin': [], 'cp_bin': []}
        idx = 0
        for lot_id, wafer_id, part_name, df_holder in df_lot_maps.itertuples(index=True):
            df_map = df_holder.df
            part = part_name[:5]
            df_map['DATA'] = [self._bin_code_classifier(x, y, z, bin_count_dict, part) for x, y, z in
                              zip(df_map['cp_bin'], df_map['fs_bin'], df_map['bs_bin'])]
            df_add = df_map['DATA'].map(lambda x: df_bin_normalize(x, bin_set))
            if not idx:
                df_maps = df_add
                df_coord = df_map[['X', 'Y']].copy()
            else:
                df_maps += df_add
            idx += 1

        df_maps = df_maps.div(total)
        df_maps = df_join_and_drop(df_coord, df_maps, index=True)
        return df_maps

    def _events_selector(self, lot_id, start_time, end_time):
        df_lot_events = self.lots_events[lot_id]
        df = df_lot_events.loc[start_time:end_time]
        if df.empty:
            return np.nan
        else:
            return df.to_string()

    def _search_wafer_history(self, _id, search_wafer_history):
        print(f"Searching {'wafer id' if search_wafer_history else 'lot id'}: {_id} history")
        df = wafer_history(_id, search_wafer_history)
        idx = df.index.to_flat_index()
        consecutive_duplicate_idx = consecutive_duplicate_idx_select(idx)
        for start, end in consecutive_duplicate_idx:
            df.iat[(end, 4)] = df.iat[(start, 4)]
            df.iat[(end, 8)] = df.iat[(start, 8)]
            for i in range(start, end):
                df.iloc[i] = np.nan

        df.dropna(inplace=True, subset=['TrackInTime', 'TrackOutTime'])
        df_dup_idx_count = df.groupby(level=0).cumcount().astype(str) + '__'
        df_dup_idx_count = df_dup_idx_count.replace('0__', '')
        df.index = df_dup_idx_count + df.index
        lots = df.LOTID.unique()
        for lot_id in lots:
            if lot_id not in self.lots_events:
                self.lots_events[lot_id] = lot_event_search(lot_id)

        df['Events'] = df.apply(
            (lambda row: self._events_selector(row['LOTID'], row['TrackInTime'], row['TrackOutTime'])),
            axis=1)
        return DfHolder(df)

    @ensure_content_not_empty
    def open_wafer_history(self, search_wafer_history=True, return_df=False):
        """

        :param search_wafer_history:default True, False is search Lot history
        :return:
        """
        if not search_wafer_history:
            if self.wafer_id:
                raise NotLotIdError
        print('Searching history')
        path = self.dump_path + '\\Wafer_lot_history'
        Tg.chk_dir(path)
        file = '\\{}.xlsm'.format('\\{}_history'.format(current_time_stamp()))
        filename = path + file
        for part, df_lot in self._lot_list_info['lots'].items():
            df_lot_history = df_lot.copy()
            if search_wafer_history:
                df_lot_history['ID'] = df_lot_history.apply((lambda row: int(row['COMPIDS'].split('.')[1])), axis=1)
                df_lot_history['HISTORY'] = df_lot_history.apply(
                    (lambda row: self._search_wafer_history(row['COMPIDS'], True)),
                    axis=1)
            else:
                df_lot_history = df_lot_history.loc[(~df_lot_history.index.duplicated(keep='first'))]
                df_lot_history['ID'] = df_lot_history.apply((lambda row: int(row.name.split('.')[1])), axis=1)
                df_lot_history['HISTORY'] = df_lot_history.apply(
                    (lambda row: self._search_wafer_history(row.name, False)), axis=1)
                df_lot_history.sort_values(by='ID', inplace=True)
            concat_list = []
            for lot_id, wafer_id, part_name, _id, df_holder in df_lot_history.itertuples(index=True):
                df = df_holder.df
                if search_wafer_history:
                    df.columns = pd.MultiIndex.from_product([[wafer_id], df.columns])
                else:
                    print(f"Searching {part_name}, lot id: {lot_id} history ..")
                    df.columns = pd.MultiIndex.from_product([[lot_id], df.columns])
                concat_list.append(df)

            df_result = pd.concat(concat_list, axis=1, sort=False)
            df_result.rename_axis(['InstNo/Procedure'], inplace=True)
            if search_wafer_history:
                df_result.rename_axis(['WaferID', 'Item'], inplace=True, axis=1)
            else:
                df_result.rename_axis(['LotID', 'Item'], inplace=True, axis=1)
            if return_df:
                return df_result
            print(f"Updating {part} history to excel, it may take a while")
            filename = df_update_xlsx(df_result, filename, sheetname=f"{part}_History", truncate_sheet=False,
                                      startrow=0,
                                      cell_width_height=cell_lot_history_table,
                                      freeze='B3',
                                      styler=history_highlight,
                                      apply_axis=None,
                                      return_writer=True,
                                      xlsm_template='templates/history_template.xlsm',
                                      multindex_headers=2)

        filename.save()
        print('Excel file saved')
        return 'open file', path + file

    def open_lots_history(self):
        return self.open_wafer_history(search_wafer_history=False)

    def search_cp_summary(self):
        df_list = []
        df_lot_cp_summary = self._df_list.copy()
        df_lot_cp_summary['CP_SUMMARY'] = df_lot_cp_summary.apply(
            (lambda row: df_holder_apply_error_handler(cp_summary, args=[row['COMPIDS']])), axis=1)
        # df_lot_cp_summary['CP_SUMMARY'] = df_lot_cp_summary.apply((lambda row: DfHolder(cp_summary((row['COMPIDS']))))
        # , axis=1)
        df_lot_cp_summary.dropna(inplace=True)
        for lot_id, wafer_id, part_name, df_holder in df_lot_cp_summary.itertuples(index=True):
            df_cp_summary = df_holder.df
            df_cp_summary['WAFER'] = wafer_id
            df_cp_summary['LOTID'] = lot_id
            df_list.append(df_cp_summary)
        df_result = pd.concat(df_list, sort=False, ignore_index=True)
        return 'open table viewer', {'title': 'CP SUMMARY', 'df': df_result}

    def open_bin_selection_gui(self, partname, purpose, title):
        purpose_func = {'open discrete bin map selection': self.open_discrete_bin_map_gui,
                        'open continuous bin map selection': self.open_continuous_stack_bin_map_gui}
        part = partname[:7]
        try:
            des_dict = Tg.description_dict[part]
            code_dict = Tg.part_group_code_dict[part]
        except Exception as e:
            part = partname[:5]
            des_dict = Tg.description_dict[part]
            code_dict = Tg.part_group_code_dict[part]

        df_selc = pd.concat(
            {k: pd.DataFrame(dict([(key, pd.Series(val)) for key, val in v.items()])).T for k, v in code_dict.items()},
            axis=0).T
        df_selc.rename(columns={u'\u6b63\u6aa2': 'FS', u'\u80cc\u6aa2': 'BS'}, inplace=True)
        df_selc = df_selc.melt(var_name=['Side', 'Group'], value_name='Bin')
        df_selc.dropna(inplace=True)
        df_selc['Bin'] = df_selc['Bin'].astype(int).astype(str)
        try:
            df_selc['Description'] = df_selc.apply((lambda row: des_dict[row['Bin']]), axis=1)
        except KeyError as e:
            print(e)
            raise

        df_selc.reset_index(drop=True, inplace=True)
        return purpose, {'df': df_selc, 'query_func': purpose_func[purpose], 'title': title}

    @ensure_content_not_empty
    def open_discrete_bin_map_selection(self):
        if self.wafer_map:
            if len(self._parts_list) > 1:
                raise MoreThanOnePartError
            elif len(self._df_list.index) > 1:
                partname = self._parts_list[0]
                title = f"Bin selection for part: {partname}, total pcs: {len(self._df_list.index)}"
                return self.open_bin_selection_gui(partname, 'open discrete bin map selection', title=title)
            else:
                for lot_id, wafer_id, partname in self._df_list.itertuples(index=True):
                    title = f"Bin selection for lot id: {lot_id}, wafer id: {wafer_id}, part: {partname}"
                    return self.open_bin_selection_gui(partname, 'open discrete bin map selection', title=title)

    def open_discrete_bin_map_gui(self, fs_bins='all', bs_bins='all', cp_bins='all', cp_stage='CP3', insp_dep='MFG'):
        if self.wafer_map:

            side_dict = {'fs_bin': 'FS',
                         'bs_bin': 'BS', 'cp_bin': 'CP'}
            lot_maps = self._lots_discrete_bin_map_gen(fs_bins=fs_bins, bs_bins=bs_bins, cp_bins=cp_bins,
                                                       cp_stage=cp_stage, insp_dep=insp_dep)
            result_list = []
            t = f"Discrete wafer map for INSP/CP map"
            aoi_folder = AOIFolder()
            for lot_id, wafer_id, part_name, df_holder in lot_maps.itertuples(index=True):
                df_map = df_holder.df
                wafer_detail = Tg.NestedDict()
                part = part_name[:5]
                # try:
                #     gross_die = Tg.gross_die[part]
                # except Tg.PartNotFoundInSettingsError:
                #     gross_die = len(df_map.index)

                gross_die = len(df_map.index)
                des_dict = Tg.description_dict[part]
                discrete_legend_values = []
                discrete_legend_colors = []
                bin_count_dict = {'fs_bin': [], 'bs_bin': [], 'cp_bin': []}
                df_map['DATA'] = [self._bin_code_classifier(x, y, z, bin_count_dict, part) for x, y, z in
                                  zip(df_map['cp_bin'], df_map['fs_bin'], df_map['bs_bin'])]

                df_map = df_map[['X', 'Y', 'DATA']]
                try:
                    pass_die_count = df_map['DATA'].value_counts().at['1']
                except KeyError:
                    pass_die_count = 0
                if pass_die_count:
                    discrete_legend_values.append('1')  # pass die
                    discrete_legend_colors.append(Colour(192, 192, 192))  # pass die is gray
                    final_yield = '{:.2%}'.format(pass_die_count / gross_die)
                    wafer_detail[('bin', '1')] = {'side': 'Pass', 'description': 'Pass die',
                                                  'count': str(pass_die_count),
                                                  'weight': final_yield}

                for side in ('fs_bin', 'bs_bin', 'cp_bin'):
                    bin_list = bin_count_dict[side]
                    if bin_list:
                        bin_counter = Counter(bin_list)
                        bins = list(bin_counter)
                        discrete_legend_colors += color_val_gen(len(bins), side)
                        for bin_code, count in bin_counter.most_common():
                            discrete_legend_values.append(bin_code)
                            wafer_detail[('bin', bin_code)] = {'side': side_dict[side],
                                                               'description': des_dict.get(bin_code, 'Others'),
                                                               'count': str(count),
                                                               'weight': '{:.2%}'.format(count / gross_die)}

                wafer_detail['AOI folder'] = aoi_folder.search_folder(wafer_id)
                wafer_detail['df_map'] = df_map
                wafer_detail['tab_title'] = wafer_id
                die_size = self.die_size_dict.get(part, (1, 1))
                # t = f"Discrete wafer map for {part_name}, lot id {lot_id}, wafer id {wafer_id}, yield {final_yield}"

                result_dict = {'die_size': die_size,

                               'wafer_detail': wafer_detail.dict,
                               'data_type': 'discrete',

                               'discrete_legend_values': discrete_legend_values,
                               'discrete_legend_colors': discrete_legend_colors}
                result_list.append(result_dict)

            return 'open wafer map', {'title': t, 'wafer_list': result_list}

    @ensure_content_not_empty
    def open_continuous_bin_map_selection(self):
        if self.wafer_map:
            pass
        if len(self._df_list.index) == 1:
            raise OneWaferError
        elif len(self._parts_list) > 1:
            raise MoreThanOnePartError
        else:
            partname = self._parts_list[0]
            title = f"Stacked map bin selection for part: {partname}"
            return self.open_bin_selection_gui(partname, 'open continuous bin map selection', title=title)

    @set_to_default_color
    def open_continuous_stack_bin_map_gui(self, fs_bins='all', bs_bins='all', cp_bins='all', cp_stage='CP3', insp_dep='MFG'):
        part = self._parts_list[0]
        die_size = self.die_size_dict.get(part, (1, 1))
        bin_set = set()
        for s in [fs_bins, bs_bins, cp_bins]:
            if isinstance(s, tuple):
                bin_set.update(s)

        if bin_set:
            bins = str(bin_set)
        else:
            bins = 'all'
        df_lots_map = self._lots_continuous_bin_map_gen(fs_bins=fs_bins, bs_bins=bs_bins, cp_bins=cp_bins,
                                                        bin_set=bin_set, cp_stage=cp_stage, insp_dep=insp_dep)
        s = df_lots_map.loc[:, 'DATA']
        if not (s > 0).any():
            raise WaferStackedDataZeroError
        wafer_detail = {'unit': '{0:.3%}', 'df_map': df_lots_map, 'tab_title': f"Stacked bin "}
        t = f"Stacked map for selected bin{bins}, part {part}. Total {len(self._df_list.index)} pcs"
        result_dict = {'die_size': die_size,

                       'wafer_detail': wafer_detail,
                       'data_type': 'continuous',
                       }
        return 'open wafer map', {'title': t, 'wafer_list': [result_dict]}

    @ensure_content_not_empty
    def open_continuous_cp_value_stack_map_selection(self):
        if self.wafer_map:
            if len(self._parts_list) > 1:
                raise MoreThanOnePartError

            elif len(self._df_list.index) == 1:
                raise OneWaferError
            else:
                partname = self._parts_list[0]
                title = f"CP stack map for part: {partname}, total pcs: {len(self._df_list.index)}"
                return (
                    'cp item stacked map selection',
                    {'title': title, 'query_func': self.open_continuous_cp_value_stack_map_gui})

    @ensure_content_not_empty
    def open_continuous_cp_value_map_selection(self):
        if self.wafer_map:
            if len(self._parts_list) > 1:
                raise MoreThanOnePartError
            elif len(self._df_list.index) > 0:
                partname = self._parts_list[0]
                title = f"CP map for part: {partname}, total pcs: {len(self._df_list.index)}"
                return (
                    'cp item map selection',
                    {'title': title, 'query_func': self.open_continuous_cp_value_map_gui})
            # else:
            #     raise MoreThanOnePartError

    @ensure_content_not_empty
    def open_discrete_cp_item_pass_map_selection(self):
        if self.wafer_map:
            if len(self._parts_list) > 1:
                raise MoreThanOnePartError
            elif len(self._df_list.index) > 0:
                partname = self._parts_list[0]
                title = f"CP item pass map for part: {partname}, total pcs: {len(self._df_list.index)}"
                return (
                    'cp item pass map',
                    {'title': title, 'query_func': self.open_discrete_cp_item_pass_map_gui})

    def open_discrete_cp_item_pass_map_gui(self, cp_stage='', to_excel='', window='', output_type=''):
        part = self._parts_list[0][:5]
        lot_maps, wafer_fail_count_dict, des_dict = self._lots_cp_item_pass_map_gen(part=part, cp_stage=cp_stage)
        # wafer_fail_count_dict['wafer_id'][wafer_id][item] = fail_count
        # wafer_fail_count_dict['gross_die'] = gross_die
        result_list = []
        t = f"{part}, test type: {cp_stage}, item pass/fail map"
        side = 'CP'
        gross_die = wafer_fail_count_dict['gross_die']
        aoi_folder = AOIFolder()
        for lot_id, wafer_id, part_name, df_holder in lot_maps.itertuples(index=True):
            df_map = df_holder.df
            wafer_detail = Tg.NestedDict()
            discrete_legend_values = []
            discrete_legend_colors = []
            bin_counter = Counter(df_map['DATA'])
            bins = list(bin_counter)

            if "Pass" in bins:
                color_num = len(bins) - 1  # pass die
                discrete_legend_colors.append(Colour(192, 192, 192))  # pass die is gray
            else:
                color_num = len(bins)
            if not color_num == 0:
                discrete_legend_colors += color_val_gen(color_num, side)

            # sort by items len(more item in the same die the length is longer), 10000 for put pass die at first
            bin_list_sort_by_len = sorted(bin_counter.items(),
                                          key=lambda i: (10000, 10000 if i[0] == 'Pass' else len(i[0]), int(i[1])),
                                          reverse=True)
            for bin_code, count in bin_list_sort_by_len:
                discrete_legend_values.append(bin_code)
                wafer_detail[('bin', bin_code)] = {'side': side,
                                                   'description': description_gen(bin_code, des_dict),
                                                   'count': str(count),
                                                   'weight': '{:.2%}'.format(count / gross_die)}

            df_map = df_map[['X', 'Y', 'DATA']]
            # add uni fail item summary
            for code, bin_fail_count in wafer_fail_count_dict['wafer_id'][wafer_id]:
                uni_bin_code = f'*{code}'
                wafer_detail[('bin', uni_bin_code)] = {'side': 'CP item',
                                                       'description': des_dict.get(code, 'Others'),
                                                       'count': str(bin_fail_count),
                                                       'weight': '{:.2%}'.format(bin_fail_count / gross_die)}
                discrete_legend_values.append(uni_bin_code)  # uni_bin_code count
                discrete_legend_colors.append(None)  # not draw die

            wafer_detail['df_map'] = df_map
            wafer_detail['tab_title'] = wafer_id
            wafer_detail['AOI folder'] = aoi_folder.search_folder(wafer_id)
            die_size = self.die_size_dict.get(part, (1, 1))
            # t = f"Discrete wafer map for {part_name}, lot id {lot_id}, wafer id {wafer_id}, yield {final_yield}"

            result_dict = {'die_size': die_size,
                           'wafer_detail': wafer_detail.dict,
                           'data_type': 'discrete',
                           'discrete_legend_values': discrete_legend_values,
                           'discrete_legend_colors': discrete_legend_colors}
            result_list.append(result_dict)
        purpose = 'open wafer map'
        if to_excel == 'Yes':
            purpose += 'and to excel'
        if window == 'No':
            purpose += 'and no window'
        if output_type == "raw data":
            purpose += 'raw data'
        return purpose, {'title': t, 'wafer_list': result_list}

    @set_color
    def open_continuous_cp_value_stack_map_gui(self, item='', cp_stage='', to_excel='', window='', stack_method="",
                                               sp_usl='', sp_lsl='', use_dpat='', usl_perc='', lsl_perc='',
                                               output_type='',
                                               **kwargs):  # **kwargs for color pick in decorator
        df_lot_maps = self._lots_cp_value_map_gen(item=item, cp_stage=cp_stage)
        df_list = []
        t = f"Item {item}, test type: {cp_stage}, stacked map"
        total = len(df_lot_maps)
        df_coord = pd.DataFrame()
        idx = 0
        for lot_id, wafer_id, part_name, df_holder in df_lot_maps.itertuples(index=True):
            df_map = df_holder.df
            df_add = df_map['DATA']
            df_add.rename(wafer_id, inplace=True)
            if not idx:
                df_coord = df_map[['X', 'Y']].copy()
            df_list.append(df_add)
            idx += 1
        df_maps = pd.concat(df_list, axis=1)
        if stack_method == 'Quantile normalization':
            df_maps = quantileNormalize(df_maps)
        elif stack_method == "No normalization":
            pass
        if idx:
            df_maps.loc[:, 'DATA'] = df_maps.sum(numeric_only=True, axis=1)
            df_map = df_maps['DATA']
            df_map = df_map.div(total)
            df_map = df_join_and_drop(df_coord, df_map, index=True)
            plot_range = plot_range_determination(df_map, sp_usl, sp_lsl, use_dpat, usl_perc, lsl_perc)
            part = part_name[:5]
            die_size = self.die_size_dict.get(part, (1, 1))
            wafer_detail = {'df_map': df_map, 'tab_title': f'CP stack map, {stack_method}'}
            result_dict = {'die_size': die_size,
                           'wafer_detail': wafer_detail,
                           'data_type': 'continuous',
                           'plot_range': plot_range}
            purpose = 'open wafer map'
            if to_excel == 'Yes':
                purpose += 'and to excel'
            if window == 'No':
                purpose += 'and no window'
            if output_type == "raw data":
                purpose += 'raw data'
            return purpose, {'title': t, 'wafer_list': [result_dict]}

    @set_color
    def open_continuous_cp_value_map_gui(self, item='', cp_stage='', to_excel='', window='', usl_perc='', lsl_perc='',
                                         use_dpat='', sp_usl='', sp_lsl='', output_type='',
                                         **kwargs):  # **kwargs for color pick in decorator
        lot_maps = self._lots_cp_value_map_gen(item=item, cp_stage=cp_stage)
        result_list = []
        t = f"Item {item}, test type: {cp_stage}"
        for lot_id, wafer_id, part_name, df_holder in lot_maps.itertuples(index=True):
            part = part_name[:5]
            die_size = self.die_size_dict.get(part, (1, 1))
            df_map = df_holder.df

            df_map = df_map[['X', 'Y', 'DATA']]
            plot_range = plot_range_determination(df_map, sp_usl, sp_lsl, use_dpat, usl_perc, lsl_perc)

            # t = f"Item {item}, CP stage {cp_stage}, CP map for part {part}, lot {lot_id}, wafer {wafer_id}"

            wafer_detail = {'df_map': df_map, 'tab_title': wafer_id}
            result_dict = {'die_size': die_size,
                           'wafer_detail': wafer_detail,
                           'data_type': 'continuous',
                           'plot_range': plot_range}
            result_list.append(result_dict)

        purpose = 'open wafer map'
        if to_excel == 'Yes':
            purpose += 'and to excel'
        if window == 'No':
            purpose += 'and no window'
        if output_type == "raw data":
            purpose += 'raw data'
        return purpose, {'title': t, 'wafer_list': result_list}

    def _tco_map_dfs_dict(self):
        cp_map_dict = Tg.NestedDict()
        for part in self._parts_list:
            unnecessary_dies = unnecessary_dies_dict.get(part)
            df_info = self._lots_dict.get(part)
            for row in df_info.itertuples(name='MAP'):
                wafer_id = row[1]
                df_cp_raw = cp_map(wafer_id, part)
                if unnecessary_dies is not None:
                    for x, y in unnecessary_dies:
                        index = df_cp_raw[((df_cp_raw['X'] == x) & (df_cp_raw['Y'] == y))].index
                        df_cp_raw.drop(index, inplace=True)

                df_bin_map = df_cp_raw.pivot(index='Y', columns='X')['BIN']
                df_die_map = df_cp_raw[['X', 'Y']]
                df_die = df_die_map.applymap(lambda z: f"{z:.0f}".zfill(2))
                df_die_map['DIE'] = df_die['Y'] + df_die['X']
                df_die_map = df_die_map.pivot(index='Y', columns='X')['DIE']
                for name, data in [('die_map', df_die_map), ('bin_map', df_bin_map)]:
                    data = data.reindex(index=(data.index[::-1]))
                    data.reset_index(inplace=True, drop=True)
                    cp_map_dict[(part, wafer_id, name)] = data

        return cp_map_dict.dict

    @ensure_content_not_empty
    def cp_reverse_bias_updated_map(self):
        print('Querying cp reverse bias updated map..')
        map_dict = self._tco_map_dfs_dict()
        path = self.dump_path + '\\TCO\\'
        Tg.chk_dir(path)
        file = '_TCO_map_{}'.format(current_time_stamp())
        filename = path + file
        for part, wafer_id_dict in map_dict.items():
            for wafer_id, item_dict in wafer_id_dict.items():
                sheetname = '{}_{}'.format(part, wafer_id)
                df_die_map = item_dict['die_map']
                df_bin_map = item_dict['bin_map']
                df_ref_dict = df_xls_color(df_die_map, df_bin_map)
                filename = df_update_xlsx(df_die_map, filename, (sheetname + '_map_update'), styler=tco_highlight,
                                          apply_axis=None,
                                          offset=0,
                                          apply_kwargs={'df_ref_dict': df_ref_dict},
                                          index=False,
                                          header=False,
                                          truncate_sheet=True,
                                          return_writer=True)
                ori_die = df_ref_dict['origin_die']
                add_die = df_ref_dict['add_die']
                fail_die_num = len(ori_die) + len(add_die)
                gross_die = Tg.gross_die(part)
                tco_yield = 1 - fail_die_num / gross_die
                additional_die = add_die - ori_die
                s_add = pd.Series((list(additional_die)), name=('additional die: {}'.format(len(additional_die))))
                s_ori = pd.Series((list(ori_die)), name=('original die: {}'.format(len(ori_die))))
                df_die_info = pd.concat([s_ori, s_add], axis=1)
                filename = df_update_xlsx(df_die_info, filename,
                                          (sheetname + '_yield_{0:.1%}_die_info'.format(tco_yield)),
                                          index=False,
                                          offset=0,
                                          header=True,
                                          truncate_sheet=True,
                                          return_writer=True)

        filename.save()
        return (
            'open file', path + file)

    @ensure_content_not_empty
    def lot_wafer_part(self):
        df = self._df_list.reset_index()
        return 'open table viewer', {'title': 'Current lot, wafer, part', 'df': df}

    @ensure_content_not_empty
    def max_cat_count(self):
        """
        :return:df of max cat count code description
        """
        concat_list = []
        for part in self._parts_list:
            df_info = self._lots_dict.get(part)
            for row in df_info.itertuples():
                lot_id = row[0]
                wafer_id = row[1]
                df1 = max_bin_by_wafer(lot_id, wafer_id, part)
                concat_list.append(df1)

        df = pd.concat(concat_list)
        df.reset_index(inplace=True)
        df.rename(columns={'index': 'SUFFER_BIN'}, inplace=True)
        df.sort_values(by=['LOT_ID'], inplace=True)
        df = df.loc[:, ['LOT_ID', 'WAFER_ID', 'PART', 'SUFFER_BIN', 'COUNT', 'NAME']]
        df.index.name = 'Index'
        df.sort_index(inplace=True)
        return 'open table viewer', {'title': 'max defect count', 'df': df}


def main():
    a = InputLotListToGet("""1NLI117.13
""", final_yield=False, wafer_id=True, first_5_char=True)
    # b = a.open_discrete_cp_item_pass_map_gui(cp_stage='CP3', to_excel='No', window='Yes')
    # b = a.open_wafer_history(search_wafer_history=False)
    b = a.open_discrete_bin_map_gui()
    pass


if __name__ == '__main__':
    main()
