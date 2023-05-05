# !/usr/bin/python

from Custom_function import InputLotListToGet, split_list
from wm_app import WaferMapApp
from wm_utils import export_to_excel, data_export_to_excel


def save_function(idx, lst):  # auto snap shot
    s = ''.join(lst)
    input_lot = InputLotListToGet(s, wafer_id=True, final_yield=False, wafer_map=True)
    # args = ('2', 'CP3', 'Yes', 'Yes', '99', '1', 'No')
    kwargs = {'item': '2', 'cp_stage': 'CP3', 'window': 'Yes', 'usl_perc': '99', 'lsl_perc': '1'}
    purpose, result_dict = input_lot.open_continuous_cp_value_map_gui(**kwargs)
    filename = r"C:\Users\ernie.chen\Desktop\ReportGen ver2\AP197 Rin\AP197Rin.xlsx"
    lst1 = result_dict['wafer_list']
    print('save to excel')
    df_data_lst = []
    for i, d in enumerate(lst1):
        tab_title = d['wafer_detail']['tab_title']
        df_map = d['wafer_detail']['df_map']
        die_size = d['die_size']
        plot_range = d['plot_range']
        filename, df_data = export_to_excel(df_map=df_map, color_func=None, filename=filename, die_size=die_size,
                                            tab_title=tab_title, plot_range=plot_range, first=False if i else True)
        df_data_lst.append(df_data)
    else:
        filename = data_export_to_excel(df_data_lst, filename)
        filename.save()
    result_dict['auto_snapshot'] = True
    wafer_map = WaferMapApp(**result_dict)


# def save_function1(idx, lst, filename):
#     s = ''.join(lst)
#     input_lot = InputLotListToGet(s, wafer_id=True, final_yield=False, wafer_map=True)
#     args = ('2', 'CP3', 'Yes', 'Yes', '99', '1', 'No')
#     purpose, result_dict = input_lot.open_continuous_cp_value_map_gui(*args)
#     lst1 = result_dict['wafer_list']
#     print('save to excel')
#     for d in lst1:
#         tab_title = d['wafer_detail']['tab_title']
#         df_map = d['wafer_detail']['df_map']
#         die_size = d['die_size']
#         plot_range = d['plot_range']
#         filename = export_to_excel(df_map, None, filename, die_size, return_writer=True,
#                                    tab_title=tab_title, plot_range=plot_range, save_map=False)
#     else:
#         return filename


# def main1():
#     fp = open(r'C:\Users\ernie.chen\Desktop\ReportGen ver2\AP197 Rin\AP197.txt', "r")
#     filename = r"C:\Users\ernie.chen\Desktop\ReportGen ver2\AP197 Rin\AP197Rin.xlsx"
#     lines = fp.readlines()
#
#     # close file
#     fp.close()
#     split_num = (len(lines) // 100) + 1
#     list_of_lists = split_list(lines, wanted_parts=split_num)
#     for idx, lst in enumerate(list_of_lists):
#         filename = save_function1(idx, lst, filename)
#     else:
#         filename.save()


def main():
    fp = open(r'C:\Users\ernie.chen\Desktop\ReportGen ver2\AP197 Rin\AP197.txt', "r")

    lines = fp.readlines()

    # close file
    fp.close()
    split_num = (len(lines) // 100) + 1
    list_of_lists = split_list(lines, wanted_parts=split_num)
    for idx, lst in enumerate(list_of_lists):
        save_function(idx, lst)


if __name__ == '__main__':
    main()
    # main1()
