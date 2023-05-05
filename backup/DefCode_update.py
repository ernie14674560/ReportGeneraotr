#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import pandas as pd
import os
from DefGroup_update import ini_parser, explode
from Database_query import conti_defcode
from Template_generator import nested_dict_by_excel

def deduplicate(x):
    if 'AP174 背檢' in x:
        x = 'AP174 背檢'
    elif 'AP174 正檢' in x:
        x = 'AP174 正檢'
    elif '正檢' in x:
        x = '正檢'
    elif '背檢' in x:
        x = '背檢'
    return x


def df_update_xlsx(df, filename):
    # groupby() & sum() to merge duplicate index
    # add(', ') & apply(lambda x: x.str.rstrip(', ')) to add ', ' between strings after sum()
    df = df.add(', ').groupby(df.index).sum().apply(lambda x: x.str.rstrip(', '))
    # deduplicate side name
    df['SIDE'] = df['SIDE'].apply(lambda x: deduplicate(x))
    # adjust side by code
    # with open('config', 'r', encoding="utf8") as ymlfile:
    #     cfg = yaml.load(ymlfile)
    # sp_to_gen = cfg['code adjustment']['side']
    # for side, codelist in sp_to_gen.items():
    # adjustment_source = nested_dict_by_excel(r'code adjustment/side by code.xlsx')
    # for sheetname, sides in adjustment_source.items():
    #     for side, codelist in sides.items():
    #         for n in codelist:
    #             n = str(n)
    #             df.loc[n, 'SIDE'] = side
    # writer = StyleFrame.ExcelWriter('Defect code lists.xlsx')
    # sf = StyleFrame(df.sort_values(by=['CODE']))
    # sf.to_excel(writer, sheet_name='DEFCODE', index=True, best_fit=sf.columns.values.tolist())
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df = df.sort_values(by=['CODE'])
    df.to_excel(writer, sheet_name='DEFCODE', index=True)
    worksheet = writer.sheets['DEFCODE']  # pull worksheet object
    # try to adjust column width
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
        )) + 1  # adding a little extra space
        worksheet.set_column(idx + 1, idx + 1, max_len)  # set column width
    writer.save()


def update_by_ini():
    original_working_directory = os.getcwd()
    new_networked_directory = r'\\hc01rpt6\CIM\Inspection'
    # change to the networked directory
    os.chdir(new_networked_directory)
    # do stuff
    df = pd.DataFrame(ini_parser('Inspection.ini')).drop(['part'], axis=1)
    df = explode(df, df.columns.values.tolist()).rename(index={0: 'CODE', 1: 'DESCRIPTION', 2: 'SIDE'})
    # change back to original working directory
    os.chdir(original_working_directory)
    df = df.T.set_index('CODE')
    df_update_xlsx(df, 'Defect code lists.xlsx')


def update_by_db():
    df = conti_defcode()
    df_update_xlsx(df, 'Defect code lists by db.xlsx')


# def main():
#     inspcode update_by_ini() or update_by_db()

def main():
    update_by_ini()


if __name__ == '__main__':
    main()
