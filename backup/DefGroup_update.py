#!/usr/bin/env python
# _*_ coding:utf-8 _*_

from configparser import RawConfigParser, SectionProxy, MissingSectionHeaderError, DuplicateOptionError
from df_to_excel import df_update_xlsx
import Template_generator as Tg
import pandas as pd
import numpy as np
import os
import sys
import re


class DuplicateParser(RawConfigParser):
    def _read(self, fp, fpname):
        """Parse a sectioned configuration file.

        Each section in a configuration file contains a header, indicated by
        a name in square brackets (`[]'), plus key/value options, indicated by
        `name' and `value' delimited with a specific substring (`=' or `:' by
        default).

        Values can span multiple lines, as long as they are indented deeper
        than the first line of the value. Depending on the parser's mode, blank
        lines may be treated as parts of multiline values or ignored.

        Configuration files may include comments, prefixed by specific
        characters (`#' and `;' by default). Comments may appear on their own
        in an otherwise empty line or may be entered in lines holding values or
        section names.

        ### Change the section between 'sectname = mo.group('header') and 'if sectname == self.default_section:'
        to avoid ERROR when the section names are duplicated

        ### Change the section between '(self._strict and(sectname, optname) in elements_added)'
        and 'raise DuplicateOptionError(sectname, optname, fpname, lineno)'
        to avoid ERROR when the option names are duplicated
        """
        elements_added = set()
        cursect = None  # None, or a dictionary
        sectname = None
        optname = None
        lineno = 0
        indent_level = 0
        _unique = 0  # unique number for duplicated sections name
        _unique1 = 0  # unique number for duplicated options name
        e = None  # None, or an exception
        for lineno, line in enumerate(fp, start=1):
            comment_start = sys.maxsize
            # strip inline comments
            inline_prefixes = {p: -1 for p in self._inline_comment_prefixes}
            while comment_start == sys.maxsize and inline_prefixes:
                next_prefixes = {}
                for prefix, index in inline_prefixes.items():
                    index = line.find(prefix, index + 1)
                    if index == -1:
                        continue
                    next_prefixes[prefix] = index
                    if index == 0 or (index > 0 and line[index - 1].isspace()):
                        comment_start = min(comment_start, index)
                inline_prefixes = next_prefixes
            # strip full line comments
            for prefix in self._comment_prefixes:
                if line.strip().startswith(prefix):
                    comment_start = 0
                    break
            if comment_start == sys.maxsize:
                comment_start = None
            value = line[:comment_start].strip()
            if not value:
                if self._empty_lines_in_values:
                    # add empty line to the value, but only if there was no
                    # comment on the line
                    if (comment_start is None and
                            cursect is not None and
                            optname and
                            cursect[optname] is not None):
                        cursect[optname].append('')  # newlines added at join
                else:
                    # empty line marks end of value
                    indent_level = sys.maxsize
                continue
            # continuation line?
            first_nonspace = self.NONSPACECRE.search(line)
            cur_indent_level = first_nonspace.start() if first_nonspace else 0
            if (cursect is not None and optname and
                    cur_indent_level > indent_level):
                cursect[optname].append(value)
            # a section header or option header?
            else:
                indent_level = cur_indent_level
                # is it a section header?
                mo = self.SECTCRE.match(value)
                if mo:
                    sectname = mo.group('header')
                    # modify this section from the source code to compromise the situation
                    while sectname in self._sections:
                        if self._strict and sectname in elements_added:
                            _unique += 1
                            sectname += str(_unique)
                            continue
                            # when section names in properties file (.ini) are duplicated
                        cursect = self._sections[sectname]
                        elements_added.add(sectname)

                    if sectname == self.default_section:
                        cursect = self._defaults
                    else:
                        cursect = self._dict()
                        self._sections[sectname] = cursect
                        self._proxies[sectname] = SectionProxy(self, sectname)
                        elements_added.add(sectname)
                    # So sections can't start with a continuation line
                    optname = None
                # no section header in the file?
                elif cursect is None:
                    raise MissingSectionHeaderError(fpname, lineno, line)
                # an option line?
                else:
                    mo = self._optcre.match(value)
                    if mo:
                        optname, vi, optval = mo.group('option', 'vi', 'value')
                        if not optname:
                            e = self._handle_error(e, fpname, lineno, line)
                        optname = self.optionxform(optname.rstrip())
                        while (self._strict and
                               (sectname, optname) in elements_added):
                            _unique1 += 1
                            optname += str(_unique1)
                            continue
                            # raise DuplicateOptionError(sectname, optname, fpname, lineno)
                        elements_added.add((sectname, optname))
                        # This check is fine because the OPTCRE cannot
                        # match if it would set optval to None
                        if optval is not None:
                            optval = optval.strip()
                            cursect[optname] = [optval]
                        else:
                            # valueless option handling
                            cursect[optname] = None
                    else:
                        # a non-fatal parsing error occurred. set up the
                        # exception but keep going. the exception will be
                        # raised at the end of the file and will contain a
                        # list of all bogus lines
                        e = self._handle_error(e, fpname, lineno, line)
        self._join_multiline_values()
        # if any parsing errors occurred, raise an exception
        if e:
            raise e


# class NestedDict:
#     """
#     Dictionary that can use a list to get a value
#
#     :example:
#     >> nested_dict = NestedDict({'aggs': {'aggs': {'field_I_want': 'value_I_want'}, 'None': None}})
#     >> path = ['aggs', 'aggs', 'field_I_want']
#     >> nested_dict[path]
#     'value_I_want'
#     >> nested_dict[path] = 'changed'
#     >> nested_dict[path]
#     'changed'
#     """
#
#     def __init__(self, *args, **kwargs):
#         self.dict = dict(*args, **kwargs)
#
#     def __getitem__(self, keys):
#         # Allows getting top-level branch when a single key was provided
#         if not isinstance(keys, tuple):
#             keys = (keys,)
#
#         branch = self.dict
#         if not isinstance(branch, dict):
#             raise KeyError
#         for key in keys:
#             branch = branch[key]
#
#         # If we return a branch, and not a leaf value, we wrap it into a NestedDict
#         return Tg.NestedDict(branch) if isinstance(branch, dict) else branch
#
#     def __setitem__(self, keys, value):
#         # Allows setting top-level item when a single key was provided
#         if not isinstance(keys, tuple):
#             keys = (keys,)
#
#         branch = self.dict
#         for key in keys[:-1]:
#             if key not in branch:
#                 branch[key] = {}
#             branch = branch[key]
#         branch[keys[-1]] = value
#
#     def clear(self):
#         self.dict.clear()
#
#     def get(self, keys, default=None, *default_args, **default_kwargs):
#         try:
#             return self.__getitem__(keys)
#         except KeyError:
#             if callable(default):
#                 print('QQQQ, {}'.format(*default_args))
#                 return default(*default_args, **default_kwargs)
#             else:
#                 return default
#
#     def pop(self, keys, default=None):
#         try:
#             result = self.__getitem__(keys)
#             self.dict.pop(keys)
#         except KeyError:
#             result = default
#         return result
#
#
# def nested_dict_by_excel(filename, index_col=None, usecols=None, string=False, sheet_as_key=True, transpose=True,
#                          force_str=False, header=0):
#     """return {excel sheet name:{index1:{index2:{.....:[values, ....], ...}}}
#        need to input filename and index_col like [0, 1]"""
#     # default index col
#     if index_col is None:
#         index_col = 0
#
#     full_dict = {}
#     df_dict = pd.read_excel(filename, sheet_name=None, index_col=index_col, usecols=usecols, header=header,
#                             dtype=object)
#     for sheet, df in df_dict.items():
#         nest = Tg.NestedDict()
#         if transpose:
#             df = df.T
#         # if string:
#         #     for keys, values in df.to_dict('list').items():
#         #         nest[keys] = values[0]
#         # else:
#         #     for keys, values in df.to_dict('list').items():
#         #         nest[keys] = [n for n in values if pd.notnull(n)]
#         for keys, values in df.to_dict('list').items():
#             if string:
#                 nest[str(keys) if force_str else keys] = str(values[0]) if force_str else values[0]
#             else:
#                 nest[str(keys) if force_str else keys] = [str(n) if force_str else n for n in values if pd.notnull(n)]
#             # else:
#             #     if string:
#             #         nest[keys] = values[0]
#             #     else:
#             #         nest[keys] = [n for n in values if pd.notnull(n)]
#         if sheet_as_key:
#             full_dict[sheet] = nest.dict
#         else:
#             full_dict = nest.dict
#     return full_dict


def as_dict(config):
    """
    Converts a ConfigParser object into a dictionary.

    The resulting dictionary has sections as keys which point to a dict of the
    sections options as key => value pairs.
    """
    the_dict = {}
    for section in config.sections():
        if re.match(r'BSVDefect:|FSVDefect:|Dummy', section):  # search BSVDefect or FSVDefect or Dummy section
            the_dict[section] = {}
            for key, val in config.items(section):
                the_dict[section][key] = val
    return the_dict


def val_parser(my_dict):
    """change{BSVDefect:Bond chuck mark, {'code': '2193,8888', 'part': 'AP148 | AP150', 'sblgp': 'Bond chuck mark'}}
    to [{'code': ['2193','8888'], 'part': ['AP148,'AP150'] , 'sblgp': 'Bond chuck mark'}, ...]
    """
    my_list = []
    for key, value in my_dict.items():
        for a, b in value.items():
            if a == 'code':
                value[a] = b.split(',')
            elif a == 'part' or re.match(r'defect\d*', a):  # r'DEFECT\d+' for the Inspection ini file section
                value[a] = b.split(' | ')
        my_list.append(value)
    return my_list


def query_gp(code, part, search_list):
    """input code and part to get defect group name"""
    result = [n['sblgp'] for n in search_list if code in n['code'] and part in n['part']]
    if len(result) > 1:
        print(code, part, result)
        raise ValueError('Identical part and code but different sbl group is not allowed!')
    return ''.join(result)


def ini_prefix(filepath):
    """add dummy section name to ini which has no section and return string"""
    with open(filepath) as f:
        file_content = '[Dummy]\n' + f.read()
        return file_content


def ini_parser(filepath):
    """input ini file and output dictionary for each section"""
    if os.path.isfile(filepath):
        config = DuplicateParser(comment_prefixes=('#', ';', "'"))
        if filepath.endswith('EDAS.ini'):
            config.read(filepath)
        elif filepath.endswith('Inspection.ini'):
            config.read_string(ini_prefix('Inspection.ini'))
        else:
            raise NameError('Need EDAS or Inspection ini file to parse')
        return val_parser(as_dict(config))
    else:
        raise FileNotFoundError('Need EDAS or Inspection ini file to parse')


def explode(df, lst_cols, fill_value=''):
    # make sure `lst_cols` is a list
    if lst_cols and not isinstance(lst_cols, list):
        lst_cols = [lst_cols]
    # all columns except `lst_cols`
    idx_cols = df.columns.difference(lst_cols)

    # calculate lengths of lists
    lens = df[lst_cols[0]].str.len()

    if (lens > 0).all():
        # ALL lists in cells aren't empty
        return pd.DataFrame({
            # col: np.repeat(df[col].values, lens)
            col: np.repeat(df[col].values, lens.astype(np.int32))
            # .astype(np.int32) to running on 32 bit system, useless in 64 bit system
            for col in idx_cols
        }).assign(**{col: np.concatenate(df[col].values) for col in lst_cols}) \
                   .loc[:, df.columns]
    else:
        # at least one list in cells is empty
        return pd.DataFrame({
            # col: np.repeat(df[col].values, lens)
            col: np.repeat(df[col].values, lens.astype(np.int32))
            # .astype(np.int32) to running on 32 bit system, useless in 64 bit system
            for col in idx_cols
        }).assign(**{col: np.concatenate(df[col].values) for col in lst_cols}) \
                   .append(df.loc[lens == 0, idx_cols]).fillna(fill_value) \
                   .loc[:, df.columns]


def df_group(modify=True, ooc_group=False):
    df_code = pd.read_excel('Defect code lists.xlsx', sheet_name='DEFCODE', index_col='CODE', dtype=object)
    part = Tg.active_parts
    code = df_code.index.values
    original_working_directory = os.getcwd()
    # new_networked_directory = r'\\hc01rpt6\CIM\EDAS'
    # change to the networked directory
    # os.chdir(new_networked_directory)
    # do stuff
    search_list = ini_parser(r'\\hc01rpt6\CIM\EDAS\EDAS.ini')
    # change back to original working directory
    # os.chdir(original_working_directory)
    list_group = []
    for c in code:
        for p in part:
            c = str(c)
            if p == 'AP174':
                n = {'PART': p, 'CODE': c, 'GROUP': query_gp(c, 'AP149', search_list)}
            else:
                n = {'PART': p, 'CODE': c, 'GROUP': query_gp(c, p, search_list)}
            list_group.append(n)
    df = pd.DataFrame(list_group)
    if ooc_group:
        df = df.groupby(['PART', 'GROUP'])['CODE'].apply(list).reset_index().pivot(index='GROUP', columns='PART',
                                                                                   values='CODE')
    else:
        df = df.pivot(index='CODE', columns='PART', values='GROUP')
    # adjust sbl group name by part and code
    if modify:
        for adjust_part_name, adjust_group in Tg.nested_dict_by_excel(r'code adjustment/group by part.xlsx',
                                                                   index_col=[0, 1]).items():
            for side, group in adjust_group.items():
                # drop cp code
                if side == 'CP':
                    continue
                for grp_name, code_list in group.items():
                    for n in code_list:
                        n = str(n)
                        df.loc[n, adjust_part_name] = grp_name
                        # df.style.apply(highlight_specific_cell(df, n, adjust_part_name), axis=None)
    return df


def df_defcode():
    df = pd.read_excel('Defect code lists.xlsx', sheet_name='DEFCODE',
                       dtype=object)
    df = df.groupby(['SIDE', 'CODE']).DESCRIPTION.sum().to_frame()
    return df


# class NestedDict(dict):
#     def __getitem__(self, key):
#         if key in self:
#             return self.get(key)
#         return self.setdefault(key, NestedDict())
def df_ooc_group_update():
    df = df_group(modify=False, ooc_group=True)
    df_update_xlsx(df, 'ooc_group', 'ooc_group', offset=1)


def highlight_active_code(x):
    """ highlight active code in DefectGroup.xlsx"""
    df = pd.DataFrame('', index=x.index, columns=x.columns)
    # color = 'background-color: #FFFF00'
    color = 'background-color: yellow'
    adjustment_source = df_defcode()
    for part, group in Tg.part_code_dict.items():
        for side, codelist in group.items():
            if side == 'CP':
                continue
            for code in codelist:
                code = str(code)
                _side = side
                # adjust side name according to DefectGroup.xlsx, Defect code lists.xlsx
                if code in adjustment_source.loc['暫時性Reject code'].index:
                    _side = '暫時性Reject code'
                elif code in adjustment_source.loc['AP174 正檢'].index:
                    _side = 'AP174 正檢'
                elif code in adjustment_source.loc['AP174 背檢'].index:
                    _side = 'AP174 背檢'
                df.loc[_side, code][part] = color
    # return color df
    return df


def join_by_code(df_left, df_right):
    result = df_left.join(df_right, how='inner')
    return result


def df_group_update(modify=True):
    # df = join_by_code(df_defcode(), df_group())
    df = join_by_code(df_defcode(), df_group(modify=modify))
    if modify:
        filename = 'DefectGroup'
    else:
        filename = 'DefectGroup_origin'
    # df_update_xlsx(df, 'DefectGroup', 'DEFGROUP', styler=highlight_active_code)
    df_update_xlsx(df, filename, 'DEFGROUP', styler=highlight_active_code)


def main():
    # return nested_dict_by_excel(r'code adjustment/test.xlsx', index_col=[0, 1, 2])
    return Tg.nested_dict_by_excel(r'reference/baseline.xlsx', string=True, index_col=None, header=[0, 1], transpose=False)


if __name__ == '__main__':
    main()
