#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import copy

import yaml

cfg_file = 'config.yml'


def unnecessary_dies_to_tup(d):
    p = d['map info']['unnecessary dies in map']
    for part, dies in p.items():
        die_yx_tups = []
        for die in dies:
            half_len = int(len(die) / 2)
            y = die[:half_len]
            x = die[half_len:]
            die_yx_tups.append((int(y), int(x)))
        p[part] = die_yx_tups


def die_size_to_tup(d):
    p = d['map info']['die size in mm']
    for part, die_size in p.items():
        lst = die_size.split('*')
        y = float(lst[0])
        x = float(lst[1])
        p[part] = (y, x)


def color_to_tup(d):
    p = d['map info']['default color(RGB)']
    for color, s in p.items():
        rgb = s.replace(' ', '')
        lst = rgb.split(',')
        r = int(lst[0])
        g = int(lst[1])
        b = int(lst[2])
        p[color] = (r, g, b)


def gross_die_to_int(d):
    p = d['gross die by part']
    for part, gross_die in p.items():
        p[part] = int(gross_die)


def yield_to_float(d):
    p = d['summary table parts and yield goal']
    for part, goal_yield in p.items():
        i = goal_yield.rstrip('%')
        f = int(i) / 100
        p[part] = f


def spec_to_float(d):
    """{'AP9B5': {'CP3': {'item': {'1': {'>': '1'},
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
                                                             '14': {'>': '1'}}}}},"""
    p = d['map info']['CP item pass spec']
    for part, spec_condition in p.items():
        for cp_stage, items in spec_condition.items():
            d = items['item spec']
            for item, spec_operator in d.items():
                for operator, spec in spec_operator.items():
                    f = float(spec)
                    p[part][cp_stage]['item spec'][item][operator] = f


def dump_cfg(data):
    """
    :param data: dict
    :return: None
    """
    with open(cfg_file, 'w') as outfile:
        yaml.dump(data, outfile)


def set_default_cfg():
    default_config = {
        'database': {'username': 'plldba', 'password': 'password', 'driver': '{Microsoft ODBC for Oracle}',
                     'server': 'MES', 'host': '192.168.1.125', 'port': '1521'},
        'inline_prompts_dict': {
            'FSINSP_SYL': ['AP196', 'AP197', 'AP174', 'AP1C0', 'AP1CC', 'AP1C5', 'AP174', 'AP148', 'AP150', 'AP191',
                           'AP198'],
            'CP1_YIELD': ['AP1FG']},
        'active part': {'Gen1': ['AP196', 'AP197', 'AP174', 'AP1C0', 'AP1CC00', 'AP1CC01', 'AP1C5', 'AP17400',
                                 'AP17406', 'AP17407', 'AP19600', 'AP1CC',
                                 'AP19601',
                                 'AP19602', 'AP19603', 'AP19700', 'AP19701', 'AP19702'],
                        'Conti': ['AP148', 'AP150'],
                        'SMP': ['AP191', 'AP198'],
                        'Panasonic': ['AP1FG']},
        'summary table parts and yield goal': {"AP148": "88%", "AP150": "88%", "AP196": "89%", "AP174": "88%",
                                               'AP197': '89%', 'AP1CC00': '88%', 'AP1CC01': '88%', 'AP1C5': '88%',
                                               'AP1C0': '88%', 'AP19600': '89%', 'AP19601': '89%', 'AP19602': '89%',
                                               'AP19603': '89%', 'AP19700': '89%', 'AP19701': '89%', 'AP19702': '89%',
                                               'AP17406': '88%', 'AP17407': '88%', 'AP17400': '88%', 'AP191': '85%',
                                               'AP198': '85%', 'AP1FG': '99%'},
        'BSL part': ['AP191', 'AP198', 'AP174', 'AP197', 'AP196', 'AP148', 'AP150', 'AP17400'],
        'BSL tolerance': '0.002',
        'OOC part': ['AP196', 'AP197', 'AP1C0'],
        'part show in docx summary': {
            'table': ['AP148', 'AP150', 'AP1C5', 'AP17400', 'AP19600', 'AP19601', 'AP19602', 'AP19603', 'AP19700',
                      'AP19701', 'AP19702', 'AP191', 'AP198'],
            'charts': ['AP148', 'AP150', 'AP1C5', 'AP196', 'AP197', 'AP191', 'AP198'],
            'scrap pcs': ['AP148', 'AP150', 'AP1C5', 'AP196', 'AP197', 'AP191', 'AP198'],
            'OOC': ['AP196', 'AP197']},
        'parts show in monthly pptx': {
            'parts use in report': ['AP196', 'AP197', 'AP17400', 'AP17406', 'AP17407', 'AP1C0'],
            'parts and title': {'AP196': 'Gen-1 Differential', 'AP197': 'Gen-1 Absolute',
                                'AP17400': 'Embedded Tie Bar', 'AP17406': 'Embedded Tie Bar',
                                'AP17407': 'Embedded Tie Bar',
                                'AP1C0': 'Back Side Absolute'},
            'parts show in slide': {'AP196 & AP197': 'Gen-1 Diff & Abs pressure sensor',
                                    'AP17400': 'high pressure sensor',
                                    'AP17406': 'high pressure sensor',
                                    'AP17407': 'high pressure sensor',
                                    'AP1C0': 'BSA pressure sensor'},
            'merge parts': {'AP196 & AP197': ['AP196', 'AP197']},
            'exclude parts': {}},
        'query scrap parts': ['AP148', 'AP150', 'AP196', 'AP174', 'AP197', 'AP1C0', 'AP1CC', 'AP1C5', 'AP191', 'AP198',
                              'AP1BF', 'AC1B2', 'AP1FG'],
        'update bsl minimal pcs': '18',
        'gross die by part': {"AP197": '4528', 'AP19700': '4528', 'AP19701': '4528', 'AP19702': '4528', 'AP148': '1680',
                              'AP150': '4537', 'AP174': '14419', 'AP1C0': '2686', 'AP17406': '14419', 'AP1BF': '7572',
                              'AP17400': '14419', 'AP17407': '14419', 'AP1CC': '3029'},
        'map info': {
            'unnecessary dies in map': {'AP197': ['0220', '0120', '0121', '0122', '0259', '0159', '0157', '0158'],
                                        'AP174': ['6516']},
            'die size in mm': {'AP196': '2.06*2.06', 'AP197': '1.58*1.58', 'AP1C5': '2.7*2.7', 'AP1CC': '2.05*2.05',
                               'AP174': '0.788*0.788', 'AP1C0': '2.01*2.36', 'AP1CU': '2.1*2.1', 'AP9B5': '9*1',
                               'AP1BQ': '0.64*1.13'},
            'CP item pass spec': {'AP9B5': {'CP3': {'item spec': {'1': {'>': '1'},
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
                                                                  '14': {'>': '1'}},
                                                    'item description': {'1': 'short1',
                                                                         '2': 'short2',
                                                                         '3': 'short3',
                                                                         '4': 'short4',
                                                                         '5': 'short5',
                                                                         '6': 'short6',
                                                                         '7': 'short7',
                                                                         '8': 'short8',
                                                                         '9': 'short9',
                                                                         '10': 'short10',
                                                                         '11': 'short11',
                                                                         '12': 'short12',
                                                                         '13': 'short13',
                                                                         '14': 'short14'}}},
                                  'AP1FG': {'CP1': {"item spec": {'1': {'>': '-20', '<': '20'},
                                                                  '2': {'>': '4.5', '<': '5.5'},
                                                                  '3': {'>': '4.5', '<': '5.5'},
                                                                  '4': {'>': '0', '<': '1'}},
                                                    'item description': {'1': 'offset',
                                                                         '2': 'Rin',
                                                                         '3': 'Rout',
                                                                         '4': 'leakage'}}}},
            'default color(RGB)': {'OOS high color': '255, 20, 147',
                                   'OOS low color': '160, 160, 160',
                                   'High color': '255, 0, 0',
                                   'Low color': '0, 0, 255'},
            'test types': ['CP1', 'CP2', 'CP3', 'CP4', 'CP5', 'CP6', 'CPY1', 'CPY2', 'CPY3']
        },
        'Convert jpg to png': [r'I:\D00ALL\!!IROM\Fusion bond\AP1BF\2020',
                               r'I:\D00ALL\!!IROM\Fusion bond\AP1BF-ANNEAL']
    }
    dump_cfg(default_config)
    return default_config


# display_cfg = {}
# cfg = {}


def reset_cfg():
    global cfg
    global display_cfg
    # try:
    with open(cfg_file, 'r', encoding="utf8") as ymlfile:
        # global cfg
        # global display_cfg
        display_cfg = {}
        cfg = {}
        cfg = yaml.load(ymlfile, Loader=yaml.BaseLoader)
        display_cfg = copy.deepcopy(cfg)
        # display_cfg = cfg.copy()
        gross_die_to_int(cfg)
        unnecessary_dies_to_tup(cfg)
        die_size_to_tup(cfg)
        yield_to_float(cfg)
        spec_to_float(cfg)
        color_to_tup(cfg)
    # except FileNotFoundError:
    #     p = r"C:\Users\ernie.chen\Desktop\Project\Project Week\config.yml"
    #     with open(p, 'r', encoding="utf8") as ymlfile:
    #         # global cfg
    #         # global display_cfg
    #         display_cfg = {}
    #         cfg = {}
    #         cfg = yaml.load(ymlfile, Loader=yaml.BaseLoader)
    #         display_cfg = copy.deepcopy(cfg)
    #         # display_cfg = cfg.copy()
    #         gross_die_to_int(cfg)
    #         unnecessary_dies_to_tup(cfg)
    #         die_size_to_tup(cfg)


def main():
    set_default_cfg()
    # reset_cfg()


if __name__ == '__main__':
    main()
else:
    reset_cfg()
