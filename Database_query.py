#!/usr/bin/env python
# _*_ coding:utf-8 _*_

import getpass
import pyodbc
import re
import socket
import string

import pandas as pd

from Configuration import cfg
from WeekTime_parser import last_week

unnecessary_dies_dict = cfg['map info']['unnecessary dies in map']

user = cfg['database']['username']
if 'password' in cfg['database']:
    pwd = cfg['database']['password']
else:
    pwd = getpass.getpass('Database password: ')
host = cfg['database']['host']
port = cfg['database']['port']
driver = cfg['database']['driver']
server = cfg['database']['server']
inline_prompts_dict = cfg['inline_prompts_dict']  # lots that pass these DCOP item Name will count in weekly report

# inline_prompts_lst = ['FSINSP_SYL']
connstr = 'DRIVER={};SERVER={};UID={};PWD={}'.format(driver, server, user, pwd)

gross_die_modified_dict = cfg['gross die by part']
id_flag = re.compile("^1[A-Z0-9][A-Z][A-Z]\d{3}\.\d{1,2}$")
part_flag = re.compile("^[A-Z]{2}[A-Z0-9]{3}\d{0,}$")
lot_flag = re.compile("^1[A-Z0-9][A-Z][A-Z]\d{3}")
sanitized_filter = string.punctuation.replace(".", "")

# verify database server

sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
# result = sock.connect_ex(('192.168.1.125', 1521))
check_result = sock.connect_ex((host, int(port)))
if check_result:
    con = None
    print("Port is not open, can not connect to database")
else:
    try:
        con = pyodbc.connect(connstr, timeout=5)
    except pyodbc.Error as err:
        print("Couldn't connect")


# dict_event_reason = {1: 'SCHEDULE', 2: 'NEWLOTSTART', 3: 'PLOTSTART', 4: 'RECEIVELOT', 5: 'STARTINVLOT', 6: 'LOTMOVE',
#                      7: 'TRACKIN', 8: 'ACCEPTINVLOT', 9: 'ENTERTEST', 10: 'CORRECTTEST', 11: 'TRACKOUT',
#                      12: 'CONTINUEINVLOT', 13: 'HOLDLOT', 14: 'RELEASELOT', 15: 'MOVELOTTOLOC', 16: 'NEWPARTFORLOT',
#                      17: 'CANCELMERGEHOLD', 18: 'MERGELOTS', 19: 'MAKESPECPART', 20: 'ADJUSTCOUNT', 21: 'MOVEHELDLOT',
#                      22: 'REROUTELOT', 23: 'RECOVERLOT', 24: 'BACKUPLOT', 25: 'SPLITLOT', 26: 'IMMEDIATEUNSCRAP',
#                      27: 'KNOWNUNSCRAP', 28: 'UNKNOWNUNSCRAP', 29: 'SHIPINVLOT', 30: 'TERMINATELOT', 31: 'UNTERMINVLOT',
#                      32: 'RETURNFROMCUST', 33: 'CHANGELOT', 34: 'LOTTIMER', 35: 'SCRAP', 36: 'KITHYBRID',
#                      37: 'MATCONSUMPTION', 38: 'SPECBRANCHRECOV', 39: 'UNSCRAP', 40: 'ABORTSTEP', 41: 'NEWPRCDFORLOT',
#                      42: 'ORDERLOT', 43: 'SKIPSTEP', 44: 'CANCELSCHEDULE', 45: 'CANCELORDER', 46: 'PARMFEEDFWD',
#                      47: 'CHANGELOTID', 48: 'CHANGELOTPARM', 49: 'TRANSFERSHIP', 50: 'TRANSFERRECV',
#                      51: 'TRANSFERCANCEL', 52: 'WITHDRAWCTLDMAT', 53: 'DESTROYCTLDMAT', 54: 'COMMENT', 55: 'UNSHIP',
#                      56: 'RUN', 57: 'CELL', 58: 'MULTIPARTSFORLOT', 59: 'LATEDATA', 60: 'TRANSFERADHOC',
#                      61: 'ADHOCMATCON', 62: 'DELPRCDCHANGE', 63: 'REMFUTUREHOLD', 64: 'CALLPRCD', 65: 'EDTHYBACT',
#                      66: 'REVISEBOM', 67: 'REPAIR', 68: 'STOCKPOS', 69: 'ASSOCIATELOT', 70: 'DISASSOCLOT',
#                      71: 'SPECIFICUNSCRAP', 72: 'AUDIT', 73: 'PLACEMARK', 74: 'UNMARK', 75: 'ADHOCCONSUME',
#                      76: 'BACKFLUSH', 77: 'CONSUME', 78: 'MARKITEM', 79: 'UNMARKITEM', 80: 'STEPRETURN',
#                      81: 'POSTSTEPCORRECT', 82: 'POSTSTEPCONSUME', 83: 'POSTSTEPRETURN', 84: 'ADDFUTA',
#                      85: 'REMOVEFUTA', 86: 'ADDLOTTOBATCH', 87: 'REMLOTFROMBATCH', 88: 'BATCHID', 255: 'CONVERT'}


def sanitize_sql(sql):
    """
    strip all the punctuation and special character
    :param sql:sql statement
    :return: strip all the punctuation and special character except "."
    """
    sql = sql.rstrip()
    sql.translate(str.maketrans('', '', sanitized_filter))
    return sql


def connection():
    global con
    if con is not None:
        try:
            con.autocommit
        except pyodbc.ProgrammingError as e:
            if e.args[0] == 'Attempt to use a closed connection.':
                try:
                    con = pyodbc.connect(connstr, timeout=5)
                except pyodbc.Error as err:
                    print("Couldn't connect")
            else:
                raise e


def con_close():
    global con
    if con is not None:
        try:
            con.close()
        except pyodbc.ProgrammingError as e:
            if not e.args[0] == 'Attempt to use a closed connection.':
                raise e


# id_error = 'please input correct pattern of LotID or WaferID '
# part_error = 'please input correct pattern of PartID'


class IDError(Exception):
    def __init__(self, _id):
        self.value = '{} is not a valid id, please input correct pattern of LotID or WaferID '.format(_id)

    def __str__(self):
        return self.value


class LotIDNotUniqueError(Exception):
    def __init__(self, _lot_ids):
        self.value = "The relationship between LotID {} to WaferID is not one-to-one. Please use WaferID instead".format(
            _lot_ids)

    def __str__(self):
        return self.value


class PartNameError(Exception):
    def __init__(self, part):
        self.value = '{} is not a valid part name, please input correct pattern of PartID'.format(part)

    def __str__(self):
        return self.value


class InspectionDataNotFoundError(Exception):
    def __init__(self, _wf_ids):
        self.value = "Can't find the inspection data of {} (wafer_id). Please check the lot history.".format(
            _wf_ids)

    def __str__(self):
        return self.value


class InspectionAOIFolderNotFoundError(Exception):
    def __init__(self, _wf_ids):
        self.value = "Can't find the AOI folder  of {} (wafer_id). Please check the lot history.".format(
            _wf_ids)

    def __str__(self):
        return self.value


class MapDataNotFoundError(Exception):
    def __init__(self, _part_id, _cp_stage):
        self.value = "Can't find the map data of {} in test type {}.".format(_part_id, _cp_stage)

    def __str__(self):
        return self.value


class CPDataNotFoundError(Exception):
    def __init__(self, _id, _test):
        self.value = "Can't find the cp data of {} (wafer_id) in testtype {}. Please check the lot history or input parameter.".format(
            _id, _test)

    def __str__(self):
        return self.value


class CPSummaryDataNotFoundError(Exception):
    def __init__(self, _id):
        self.value = "Can't find the cp data of {} (wafer_id). Please check the lot history or input parameter.".format(
            _id)

    def __str__(self):
        return self.value


class CPItemNotFoundError(Exception):
    def __init__(self, _id, _item):
        self.value = "Can't find the cp item {} of {} (wafer_id). Please reconfirm.".format(_item, _id)

    def __str__(self):
        return self.value


class GrossDieNotFoundError(Exception):
    def __init__(self, _part):
        self.value = "Can't find the cp gross die of {}. Please check the part in setting -> gross die by part.".format(
            _part)

    def __str__(self):
        return self.value


class RecipeNotFoundError(Exception):
    def __init__(self):
        self.value = "Can't find the recipe that satisfy the input condition."

    def __str__(self):
        return self.value


class NotFoundError(Exception):
    def __init__(self, wafer_id, _id):
        if wafer_id:
            _type = 'wafer_id'
        else:
            _type = 'lot_id'

        self.value = f"Can't find the {_type}: {_id}, please make sure it is in the WIP."

    def __str__(self):
        return self.value


def id_verification(lid):
    """verify lot id and wafer id to match pattern like 1NJE025.15"""
    return bool(id_flag.match(lid))


def lot_verification(lot):
    """verify lot like 1NJK859 is valid for sql line"""
    return bool(lot_flag.match(lot))


def verify_id(lid, is_id=True):
    if is_id:
        result = id_verification(lid)
    else:
        result = lot_verification(lid)
    if not result:
        raise IDError(lid)
    return result


def part_verification(p):
    """verify lot id and wafer id to match pattern like 1NJE025.15"""
    return bool(part_flag.match(p))


def verify_part(p):
    result = part_verification(p)
    if not result:
        raise PartNameError(p)
    return result


def gross_die_by_part(part, cp_stage="CP3"):
    """Query most recent gross_die by part like AP197, AP1CC, AP198 .... use cp data"""
    if verify_part(part):
        if part in gross_die_modified_dict:
            result = gross_die_modified_dict[part]
        else:
            # con = pyodbc.connect(connstr)
            sql = f"""SELECT * FROM
                  (SELECT CAST(GROSSDIE AS INTEGER) AS Numeric FROM EDADBA.CP_DATAW
                  WHERE PART LIKE '{part}%'
                  AND TESTTYPE = '{cp_stage}'
                  ORDER BY TESTDATE DESC)
                  WHERE ROWNUM = 1"""
            if con is None:
                return 0
            else:
                crsr = con.cursor()
            crsr.execute(sql)
            sql_result = crsr.fetchone()
            if sql_result is None:
                if cp_stage == "CP3":
                    result = gross_die_by_part(part, cp_stage="CP2")
                elif cp_stage == 'CP2':
                    result = gross_die_by_part(part, cp_stage='CP1')
                else:
                    raise GrossDieNotFoundError(part)
                # raise GrossDieNotFoundError(part)
            else:
                result = sql_result[0]
            # con.close()
        return result


########################################################################################################################
# gross_die = dict()
# oracle_query_parts = set()
# for _part in active_parts:
#     gross_die[_part] = query_gross_die(_part)
#     oracle_query_parts.add(_part[:5])  # only query first 5 character in part
########################################################################################################################
scrap_query_parts = cfg['query scrap parts']


########################################################################################################################

def query(sql, *args, **kwargs):
    df = pd.read_sql(sql, con, *args, **kwargs)
    return df


def recp_query_full_info(recipes: list):
    df_list = []
    for recipe in recipes:
        df_des = recp_des_query(recipe=recipe)
        df_des.sort_values(by=['SEQNUM'], inplace=True)
        df_des.reset_index(drop=True, inplace=True)
        df_tit = recp_title_query(recipe=recipe)
        df_des.loc[-1] = df_des.columns  # adding a row
        df_des.index = df_des.index + 1  # shifting index
        df_des.sort_index(inplace=True)  # sorting by index
        df_des.rename(
            columns={"RECPID": "TITLE", "SEQNUM": "CAPABILITY", 'OPIDS': 'PRODSTATUS', 'OPTYPES': 'ACTIVEFLAG',
                     'OPDESCS': 'CREATEDATE', 'ENGDESC': 'ENGCHANGEDATE'}, inplace=True)
        df = pd.concat([df_tit, df_des], sort=False, ignore_index=True)
        df_list.append(df)
    if df_list:
        df_result = pd.concat(df_list, sort=False, ignore_index=True)
    else:
        raise RecipeNotFoundError
    return df_result


def recp_query(title='', opdescs='', capability=''):
    conditions = list()
    map_list = [('C.TITLE', title), ('A.OPDESCS', opdescs), ('C.CAPABILITY', capability)]
    conditions_map = {key: sanitize_sql(val) for key, val in map_list}
    for k, v in conditions_map.items():
        if v:
            item = "{} LIKE '%{}%'".format(k, v)
            conditions.append(item)
    filter_for_query = " AND ".join(conditions)
    sql = f"""SELECT DISTINCT A.RECPID
          FROM PLLDBA.RECP_OPERATIONS A, EDADBA.TBLS_OPER B, PLLDBA.RECP C
          WHERE A.OPIDS=B.OPERID
          AND A.RECPNAME=C.RECPNAME
          AND A.RECPVERSION=C.RECPVERSION
          AND {filter_for_query}"""
    crsr = con.cursor()
    crsr.execute(sql)
    recipes = [row[0] for row in crsr]
    df = recp_query_full_info(recipes)
    return df


def recp_des_query(opdescs='', capability='', recipe=''):
    conditions = list()
    if recipe:
        recp_ver = recipe.split('.')
        if len(recp_ver) == 1:
            recp, ver = recp_ver[0], str()
        else:
            recp, ver = recp_ver[0], recp_ver[1]
    else:
        recp, ver = str(), str()
    map_list = [('RECPNAME', recp), ('RECPVERSION', ver), ('OPDESCS', opdescs), ('CAPABILITY', capability)]
    conditions_map = {key: sanitize_sql(val) for key, val in map_list}
    for k, v in conditions_map.items():
        if v:
            item = "{} LIKE '%{}%'".format(k, v)
            conditions.append(item)
    filter_for_query = " AND ".join(conditions)
    sql = f"""SELECT A.RECPID,A.SEQNUM, A.OPIDS, A.OPTYPES,A.OPDESCS, B.ENGDESC
          FROM PLLDBA.RECP_OPERATIONS A, EDADBA.TBLS_OPER B
          WHERE A.OPIDS=B.OPERID
          AND {filter_for_query}"""
    df = pd.read_sql(sql, con)
    return df


def recp_title_query(title='', capability='', recipe=''):
    conditions = list()
    if recipe:
        recp_ver = recipe.split('.')
        if len(recp_ver) == 1:
            recp, ver = recp_ver[0], str()
        else:
            recp, ver = recp_ver[0], recp_ver[1]
    else:
        recp, ver = str(), str()
    map_list = [('RECPNAME', recp), ('RECPVERSION', ver), ('TITLE', title), ('CAPABILITY', capability)]
    conditions_map = {key: sanitize_sql(val) for key, val in map_list}
    for k, v in conditions_map.items():
        if v:
            item = "{} LIKE '%{}%'".format(k, v)
            conditions.append(item)
    filter_for_query = " AND ".join(conditions)
    sql = f"""SELECT RECPID, TITLE, CAPABILITY, PRODSTATUS, ACTIVEFLAG, CREATEDATE, ENGCHANGEDATE FROM PLLDBA.RECP
          WHERE {filter_for_query}"""
    df = pd.read_sql(sql, con)
    return df


def cp_weekly(wafer_id, cp_stage='CP3'):
    """Query CP3 data by wafer_id"""
    # con = pyodbc.connect(connstr)
    if verify_id(wafer_id):
        lot_id = wafer_id.split('.')[0] + '%'
        wafer_n = wafer_id.split('.')[1]
        sql1 = f'''SELECT BIN AS "DEFECT", count(BIN) AS "COUNT" FROM (
              Select A.BIN From EDADBA.CP_DATAD A, EDADBA.CP_DATAW B 
              Where A.Link=B.Link 
              AND B.LOTID LIKE '{lot_id}' 
              AND B.WAFER = {wafer_n} 
              AND B.TESTTYPE = '{cp_stage}') GROUP BY BIN'''
        df1 = pd.read_sql(sql1, con, index_col='DEFECT')
        sql2 = f'''SELECT MARK AS "DEFECT", COUNT(MARK) AS "COUNT" FROM (SELECT DIE, MAX(MARK) AS "MARK" FROM (SELECT A.DIE, A.MARK FROM EDADBA.CP_DATAR A, EDADBA.CP_DATAW B 
                   WHERE A.LINK=B.LINK 
                   AND B.LOTID LIKE '{lot_id}' 
                   AND B.WAFER = {wafer_n}
                   AND B.TESTTYPE = '{cp_stage}') GROUP BY DIE) GROUP BY MARK'''
        df2 = pd.read_sql(sql2, con, index_col='DEFECT')
        df = pd.concat([df1, df2]).reset_index().dropna().set_index('DEFECT')
        if df.empty:
            if cp_stage == 'CP3':
                df = cp_weekly(wafer_id, cp_stage='CP2')
            elif cp_stage == 'CP2':
                df = cp_weekly(wafer_id, cp_stage='CP1')
            else:
                raise CPDataNotFoundError(wafer_id, cp_stage)
            # raise CPDataNotFoundError(wafer_id)
            # con.close()
        return df
    else:
        return pd.DataFrame()


def coord_gen(df, drop=True):
    df['coord'] = df['Y'].map(lambda z: f"{z:.0f}".zfill(2)) + df['X'].map(lambda z: f"{z:.0f}".zfill(2))
    if drop:
        df.drop(['X', 'Y'], axis=1, inplace=True)


def map_info(part, cp_stage='CP3'):
    if verify_part(part):
        sql = f"""select b.die, b.x, b.y from 
                (select link from EDADBA.MAP_SPECH where part like '{part}%' and testtype = '{cp_stage}') a, 
                (select link, die, x, y from EDADBA.MAP_SPECR) b 
                where a.link = b.link"""
        df = pd.read_sql(sql, con, index_col='DIE').astype(int)
        # df.rename(columns={'X': 'Y', 'Y': "X"}, inplace=True)  # XY inverse = =
        unnecessary_dies = unnecessary_dies_dict.get(part)
        if unnecessary_dies is not None and not df.empty:
            for y, x in unnecessary_dies:
                index = df[(df['X'] == x) & (df['Y'] == y)].index
                df.drop(index, inplace=True)

        if df.empty:
            if cp_stage == "CP3":
                df = map_info(part, cp_stage='CP2')
            elif cp_stage == 'CP2':
                df = map_info(part, cp_stage='CP1')
            else:
                raise MapDataNotFoundError(part, cp_stage)

        coord_gen(df, drop=False)

        return df
    else:
        return pd.DataFrame()


def insp_map_folder(wafer_id, side=''):
    """Query inspection map by wafer_id
        side '1' >> FS
        side '2' >> BS
    """
    if verify_id(wafer_id):
        sql = f"""select FOLDER from EDADBA.MAP_DATAD_FOLDER
                  where link in(select link from EDADBA.MAP_DATAW where wafer = '{wafer_id}' and faildie='{side}'"""
        df = pd.read_sql(sql, con)
        if df.empty:
            raise InspectionAOIFolderNotFoundError(wafer_id)
        return df
    else:
        return pd.DataFrame()


def insp_map(wafer_id, side='1', bin_codes_tup='all', ensure_not_empty=False, insp_dep='MFG', second_loop=False):
    """Query inspection map by wafer_id
        side '1' >> FS
        side '2' >> BS
    """
    # if bin_codes_tup == 'all':
    #     bin_codes_tup = ''
    d = {'1': 'fs_bin', '2': 'bs_bin'}
    df_side = d[side]
    # if side == '1':
    #     df_side = 'fs_bin'
    # elif side == '2':
    #     df_side = 'bs_bin'
    if insp_dep == 'QC':
        db_sheet = 'GLASS'
    else:
        db_sheet = 'MAP'
    if bin_codes_tup:
        if isinstance(bin_codes_tup, tuple):
            if len(bin_codes_tup) == 1:
                bin_codes_tup = "('{}')".format(bin_codes_tup[0])
            key = f'AND cat IN {bin_codes_tup}'
        else:
            key = ''
        if verify_id(wafer_id):
            sql = f"""select x, y, cat from EDADBA.{db_sheet}_DATAD
            where link in(select link from EDADBA.{db_sheet}_DATAW where wafer = '{wafer_id}' and faildie='{side}') {key}"""
            df = pd.read_sql(sql, con)

            # df.rename(columns={'X': 'Y', 'Y': "X"}, inplace=True)  # XY inverse = =
            if df.empty:
                if insp_dep == 'MFG':
                    df = insp_map(wafer_id, side=side, bin_codes_tup=bin_codes_tup, ensure_not_empty=ensure_not_empty,
                                  insp_dep='QC', second_loop=True)
                else:
                    if ensure_not_empty:
                        raise InspectionDataNotFoundError(wafer_id)
                    else:
                        return df
                if ensure_not_empty:
                    raise InspectionDataNotFoundError(wafer_id)
            if second_loop:
                return df
            # if df.empty:
            #     if ensure_not_empty:
            #         raise InspectionDataNotFoundError(wafer_id)
            #     else:
            #         return pd.DataFrame()

            coord_gen(df)
            df.rename(columns={'CAT': df_side}, inplace=True)
            # df.set_index('coord', drop=True, inplace=True)
            return df
    else:
        return pd.DataFrame()


def cp_map(wafer_id, cp_stage='CP3', bin_codes_tup='all', ensure_not_empty=False, second_loop=False):
    """Query CP3 map by wafer_id"""
    # if part.startswith('AP197'):
    #     cp_stage = 'CP1'
    if bin_codes_tup:
        if isinstance(bin_codes_tup, tuple):
            codes_tup = bin_codes_tup + tuple('1')  # always query pass die
            key = f'AND d.bin IN {codes_tup}'
        else:
            key = ''
        if verify_id(wafer_id):
            lot_id = wafer_id.split('.')[0] + '%'
            wafer_n = wafer_id.split('.')[1]
            sql = f"""select d.die, d.bin from 
                    (select link from EDADBA.CP_DATAW 
                    where lotid like '{lot_id}' and wafer = {wafer_n} and testtype ='{cp_stage}') c, 
                    (select link, die, bin from EDADBA.CP_DATAD) d 
                    where c.link = d.link {key}"""
            df = pd.read_sql(sql, con)
            # df.rename(columns={'X': 'Y', 'Y': "X"}, inplace=True)  # XY inverse = =
            if df.empty:
                if cp_stage == "CP3":
                    df = cp_map(wafer_id, cp_stage='CP2', bin_codes_tup=bin_codes_tup, second_loop=True)
                elif cp_stage == 'CP2':
                    df = cp_map(wafer_id, cp_stage='CP1', bin_codes_tup=bin_codes_tup, second_loop=True)
                else:
                    if ensure_not_empty:
                        raise CPDataNotFoundError(wafer_id, cp_stage)
                    else:
                        return df
                if ensure_not_empty:
                    raise CPDataNotFoundError(wafer_id, cp_stage)
            if second_loop:
                return df
            df.fillna("1", inplace=True)
            df.set_index('DIE', drop=True, inplace=True)
            df.rename(columns={'BIN': 'cp_bin'}, inplace=True)
            return df
    else:
        return pd.DataFrame()


def fsbs_insp_weekly(wafer_id):
    """Query FS,BS inspection data by wafer_id"""
    # con = pyodbc.connect(connstr)
    if verify_id(wafer_id):
        sql = "SELECT CAT, count(CAT) FROM " \
              "(Select A.CAT From EDADBA.MAP_DATAD A, EDADBA.MAP_DATAW B " \
              "Where A.Link=B.Link " \
              "And B.WAFER = '{}') GROUP BY CAT".format(wafer_id)
        df = pd.read_sql(sql, con, index_col='CAT')
        # if df.empty:
        #     raise InspectionDataNotFoundError(wafer_id)
        return df
    else:
        return pd.DataFrame()


def cp_summary(wafer_id):
    if verify_id(wafer_id):
        lot_id = wafer_id.split('.')[0] + '%'
        wafer_n = wafer_id.split('.')[1]
        sql = f"""SELECT LOTID,WAFER,PART,GROSSDIE,PASSDIE,YIELD,TESTTYPE,TESTDATE,OPERATOR,PROBERCARD,PROBERID,LOADDATE
              From EDADBA.CP_DATAW WHERE
              LOTID LIKE '{lot_id}' 
              AND WAFER = '{wafer_n}'
              """
        df = pd.read_sql(sql, con)
        if df.empty:
            raise CPSummaryDataNotFoundError(wafer_id)
        return df
    else:
        return pd.DataFrame()


def cp_value_map(wafer_id, item, cp_stage='CP3', ensure_not_empty=True, second_loop=False):
    if verify_id(wafer_id):
        item = sanitize_sql(item)
        cp_stage = sanitize_sql(cp_stage)
        lot_id = wafer_id.split('.')[0] + '%'
        wafer_n = wafer_id.split('.')[1]
        sql = f"""SELECT B.DIE,B.VALUE FROM EDADBA.CP_DATAW A, EDADBA.CP_DATAR B 
        WHERE A.LOTID LIKE '{lot_id}'
        AND A.TESTTYPE = '{cp_stage}' 
        AND A.WAFER='{wafer_n}' 
        AND B.LINK= A.LINK
        AND B.ITEM = {item}"""
        df = pd.read_sql(sql, con, index_col='DIE')
        # if df.empty:
        #     raise CPItemNotFoundError(wafer_id, item)
        if df.empty:
            # if cp_stage == "CP3":
            #     df = cp_value_map(wafer_id, item, cp_stage='CP2', second_loop=True)
            # elif cp_stage == 'CP2':
            #     df = cp_value_map(wafer_id, item, cp_stage='CP1', second_loop=True)
            # else:
            #     if ensure_not_empty:
            #         raise CPDataNotFoundError(wafer_id)
            #     else:
            #         return df
            if ensure_not_empty:
                raise CPDataNotFoundError(wafer_id, cp_stage)
            else:
                return df
        if second_loop:
            return df
        df.rename(columns={'VALUE': 'DATA'}, inplace=True)
        return df


def weekly_lot(end_date, start_date=None, drop_ep_lot=True):
    """
    Query Continental weekly shipping lot
    That had passed inline_prompt(FSINSP_SYL) name stage, but not actually shipping lots
    :param end_date: datetime.date object, must be friday to produce correct weekly report
    :param start_date: datetime.date object, must be friday to produce correct weekly report
    :param drop_ep_lot: determine drop ep lot or not
    :param check_point: default count passed inline_prompt(FSINSP_SYL) lots and wafers
    :return:df like
                    COMPIDS	PARTNAME
            LOTID
        1NIJ089.1	1NIJ089.1	AP17400AH-C4N3-GB
        1NIJ089.10	1NIJ089.3	AP17400AH-C4N3-GB
        1NIJ089.11	1NIJ089.4	AP17400AH-C4N3-GB
        1NIJ089.12	1NIJ089.5	AP17400AH-C4N3-GB
        1NIJ089.13	1NIJ089.6	AP17400AH-C4N3-GB
        1NIJ089.14	1NIJ089.7	AP17400AH-C4N3-GB
        1NIJ089.15	1NIJ089.8	AP17400AH-C4N3-GB
        1NIJ089.16	1NIJ089.9	AP17400AH-C4N3-GB
        1NIJ089.17	1NIJ089.10	AP17400AH-C4N3-GB
        1NIJ089.18	1NIJ089.11	AP17400AH-C4N3-GB
            ...         ...         ...
    """
    # con = pyodbc.connect(connstr)
    df_list = []
    for prompt, parts in inline_prompts_dict.items():
        if start_date is None:
            start_date = last_week(end_date)

        sql = f"""SELECT DISTINCT B.LOTID, A.COMPIDS, B.PARTNAME
              FROM PLLDBA.ACTL_COMPONENTS A, PLLDBA.TRES_LOT B, PLLDBA.TEST_DI C
              WHERE A.COMPSTATE = 1
              AND A.LOTID = B.LOTID
              AND B.TestOpNo = C.TestOpNo
              AND B.TESTINDEX = 1
              AND B.EntTime BETWEEN DATE '{str(start_date)}' AND DATE'{str(end_date)}'
              AND C.Prompt = '{prompt}'"""
        df = pd.read_sql(sql, con, index_col=['LOTID'])  # COMPIDS = wafer_id
        df = df[df.PARTNAME.str.contains('|'.join(parts))]
        if drop_ep_lot:
            df = df[df.index.str.contains("1N")]
        df_list.append(df)
    df_result = pd.concat(df_list)
    # sql = f"""SELECT DISTINCT B.LOTID, A.COMPIDS, B.PARTNAME
    #       FROM PLLDBA.ACTL_COMPONENTS A, PLLDBA.TRES_LOT B, PLLDBA.TEST_DI C
    #       WHERE A.COMPSTATE = 1
    #       AND A.LOTID = B.LOTID
    #       AND B.TestOpNo = C.TestOpNo
    #       AND B.TESTINDEX = 1
    #       AND B.EntTime BETWEEN DATE '{str(start_date)}' AND DATE'{str(end_date)}'
    #       AND C.Prompt = {prompts_tup}"""
    # df = pd.read_sql(sql, con, index_col=['LOTID'])  # COMPIDS = wafer_id
    # # con.close()
    # if drop_ep_lot:
    #     df = df[df.index.str.contains("1N")]
    return df_result


# def actual_weekly_lot(end_date, start_date=None, drop_ep_lot=True):
#     """
#     Query Continental weekly shipping lot
#     That had passed FINISH stage, and PROMISE state is FINISH, may be the actual shipping lots
#     :param end_date: datetime.date object, must be friday to produce correct weekly report
#     :param start_date: datetime.date object, must be friday to produce correct weekly report
#     :param drop_ep_lot: determine drop ep lot or not
#     :return:df like
#                     COMPIDS	PARTNAME
#             LOTID
#         1NIJ089.1	1NIJ089.1	AP17400AH-C4N3-GB
#         1NIJ089.10	1NIJ089.3	AP17400AH-C4N3-GB
#         1NIJ089.11	1NIJ089.4	AP17400AH-C4N3-GB
#         1NIJ089.12	1NIJ089.5	AP17400AH-C4N3-GB
#         1NIJ089.13	1NIJ089.6	AP17400AH-C4N3-GB
#         1NIJ089.14	1NIJ089.7	AP17400AH-C4N3-GB
#         1NIJ089.15	1NIJ089.8	AP17400AH-C4N3-GB
#         1NIJ089.16	1NIJ089.9	AP17400AH-C4N3-GB
#         1NIJ089.17	1NIJ089.10	AP17400AH-C4N3-GB
#         1NIJ089.18	1NIJ089.11	AP17400AH-C4N3-GB
#             ...         ...         ...
#     """
#     # con = pyodbc.connect(connstr)
#     if start_date is None:
#         start_date = last_week(end_date)
#     concat_list = []
#     for part in oracle_query_parts:
#         if verify_part(part):
#             sql = "SELECT DISTINCT A.LOTID, A.COMPIDS, B.PARTID " \
#                   "FROM PLLDBA.ACTL_COMPONENTS A, PLLDBA.ACTL B " \
#                   "WHERE B.PARTID LIKE '{}%' " \
#                   "AND A.LOTID = B.LOTID " \
#                   "AND A.COMPSTATE = 1 " \
#                   "AND B.STATE = 'FINISH' " \
#                   "AND B.PREVHISTTIME " \
#                   "BETWEEN DATE '{}' AND DATE'{}'".format(part, str(start_date), str(end_date))
#             df = pd.read_sql(sql, con, index_col=['LOTID'])  # COMPIDS = wafer_id
#         else:
#             df = pd.DataFrame()
#         concat_list.append(df)
#     df = pd.concat(concat_list)
#     # con.close()
#     if drop_ep_lot:
#         df = df[df.index.str.contains("1N")]
#     df.rename(columns={'PARTID': 'PARTNAME'}, inplace=True)
#     return df


def id_verify_list(lst):
    for i in lst:
        if i:  # pass empty string
            boolean = id_verification(i)
            if not boolean:
                return boolean, i
    return True, ''


def series_zfill(s):
    df = s.str.split('.', 1, expand=True)
    s_lot = df[0]
    s_wafer = df[1]
    return s_lot + '.' + s_wafer.str.zfill(2)


# def lot_list_summary(lst, wafer_id, final_yield=True):
#     verification, issue_id = id_verify_list(lst)
#     if verification:
#         # id_tup = tuple(lst)
#         # if len(id_tup) == 1:
#         #     id_tup = "('{}')".format(id_tup[0])
#         id_tup = lst
#         if len(id_tup) == 1:
#             id_tup = "(('{}', 0))".format(id_tup[0])
#         else:
#             id_tup = tuple((n, 0) for n in lst)
#         if wafer_id:
#             key = '(A.COMPIDS,0)'
#         else:
#             key = '(A.LOTID,0)'
#         if final_yield:
#             prompt = "AND C.Prompt = 'FSINSP_SYL' "
#             testop = "AND B.TestOpNo = C.TestOpNo "
#             testdi = "TRES_LOT B, TEST_DI C "
#             part_name = 'PARTNAME'
#         else:
#             prompt, testop = ('', '')
#             testdi = "ACTL B"
#             part_name = "PARTID"
#         # sql = f"""SELECT DISTINCT A.COMPIDS, B.LOTID, B.{part_name} FROM ACTL_COMPONENTS A, {testdi}
#         #           WHERE A.COMPSTATE = 1
#         #           AND A.LOTID = B.LOTID
#         #           {testop}
#         #           {prompt}
#         #           AND {key} IN {id_tup}"""
#         sql = f"""SELECT DISTINCT A.COMPIDS, B.LOTID, B.{part_name} FROM ACTL_COMPONENTS A, {testdi}
#                           WHERE  {key} IN {id_tup}
#                           AND A.LOTID = B.LOTID
#                           {testop}
#                           {prompt}
#                           AND A.COMPSTATE IN (1, 5)"""
#         df = pd.read_sql(sql, con, index_col=['LOTID'])
#
#         df.rename(columns={'PARTID': 'PARTNAME'}, inplace=True)
#         # if not df.index.is_unique:
#         #     not_unique = df.index[df.index.duplicated()].unique()
#         #     raise LotIDNotUniqueError(not_unique.to_list())
#         df.sort_values(by='COMPIDS', key=lambda col: series_zfill(col), inplace=True)
#         return df
#     else:
#         raise IDError(issue_id)


def lot_list_summary(lst, wafer_id, final_yield=True):
    verification, issue_id = id_verify_list(lst)
    if verification:
        id_tup = tuple(lst)
        if len(id_tup) == 1:
            id_tup = "('{}')".format(id_tup[0])
        # id_tup = lst
        # if len(id_tup) == 1:
        #     id_tup = "(('{}', 0))".format(id_tup[0])
        # else:
        #     id_tup = tuple((n, 0) for n in lst)
        if wafer_id:
            key = 'A.COMPIDS'
        else:
            key = 'A.LOTID'
        if final_yield:
            prompt = "AND C.Prompt = 'FSINSP_SYL' "
            testop = "AND B.TestOpNo = C.TestOpNo "
            testdi = "TRES_LOT B, TEST_DI C "
            part_name = 'PARTNAME'
        else:
            prompt, testop = ('', '')
            testdi = "ACTL B"
            part_name = "PARTID"
        # original compstate in (1,5)
        sql = f"""SELECT DISTINCT A.COMPIDS, B.LOTID, B.{part_name} FROM ACTL_COMPONENTS A, {testdi}
                  WHERE A.COMPSTATE IN (1,4,5)
                  AND A.LOTID = B.LOTID
                  {testop}
                  {prompt}
                  AND {key} IN {id_tup}"""
        # sql = f"""SELECT DISTINCT A.COMPIDS, B.LOTID, B.{part_name} FROM ACTL_COMPONENTS A, {testdi}
        #                   WHERE  {key} IN {id_tup}
        #                   AND A.LOTID = B.LOTID
        #                   {testop}
        #                   {prompt}
        #                   AND A.COMPSTATE IN (1, 5)"""
        df = pd.read_sql(sql, con, index_col=['LOTID'])
        if df.empty:
            raise NotFoundError(wafer_id=wafer_id, _id=id_tup)
        df.rename(columns={'PARTID': 'PARTNAME'}, inplace=True)
        # if not df.index.is_unique:
        #     not_unique = df.index[df.index.duplicated()].unique()
        #     raise LotIDNotUniqueError(not_unique.to_list())
        df.sort_values(by='COMPIDS', key=lambda col: series_zfill(col), inplace=True)
        df.drop_duplicates(subset=["COMPIDS"], keep="last",inplace=True)
        return df
    else:
        raise IDError(issue_id)


# def ocap_stage(lst):
#     verification, issue_id = id_verify_list(lst)
#     if verification:
#         id_tup = tuple(lst)
#         if len(id_tup) == 1:
#             id_tup = "('{}')".format(id_tup[0])
#         sql = f"""SELECT DISTINCT LOTID, STAGE FROM EDADBA.OCAP_FILEH WHERE LOTID IN {id_tup}"""
#         df = pd.read_sql(sql, con, index_col=['LOTID'])
#         if not df.index.is_unique:
#             not_unique = df.index[df.index.duplicated()].unique()
#             raise LotIDNotUniqueError(not_unique.to_list())
#         return df
#     else:
#         raise IDError(issue_id)


def scrap_comment(lotid, evtime):
    sql = "SELECT DISTINCT VARIANTFLD1 FROM PLLDBA.HIST_LOTEVENTS " \
          "WHERE LOTID = '{}'" \
          "AND EVTIME = TO_DATE('{}', 'YYYY-MM-DD HH24:MI:SS')" \
          "AND EVTYPE ='COMM' " \
          "And ROWNUM < 2".format(lotid, evtime)
    crsr = con.cursor()
    crsr.execute(sql)
    result = crsr.fetchone()[0]
    return result


def scrap_lot(end_date, start_date=None, drop_ep_lot=True, drop_terminate_lot=True):
    if start_date is None:
        start_date = last_week(end_date)
    concat_list = []
    for part in scrap_query_parts:
        if verify_part(part):
            sql = "SELECT DISTINCT A.PARTID,A.EVTIME, A.LOTID, A.PREMAINQTY, " \
                  "B.DESCRIPTION, decode(substr(C.CURPRCDKIND,1,1),'R','R','')||C.CURPRCDCURINSTNUM, C.LOTTYPE " \
                  "FROM PLLDBA.HIST_LOTEVENTS A, PLLDBA.REJC B,  PLLDBA.HIST C " \
                  "WHERE A.PARTID LIKE '{}%' " \
                  "AND A.EVTYPE ='SCRP' " \
                  "AND A.PREVHISTTIME = C.PREVHISTTIME " \
                  "AND A.LOTID = C.LOTID " \
                  "AND B.REJCAT = A.VARIANTFLD1 " \
                  "AND A.EVTIME BETWEEN DATE '{}' AND DATE '{}'".format(part, str(start_date), str(end_date))
            df = pd.read_sql(sql, con, index_col=['LOTID'])
        else:
            df = pd.DataFrame()
        concat_list.append(df)
    df = pd.concat(concat_list)
    if not df.empty:
        if drop_ep_lot:
            df = df[df.index.str.contains("1N")]
        if drop_terminate_lot:
            df = df[~df.DESCRIPTION.str.contains("Terminate")]
    if df.empty:
        return df
    df.rename(columns={'EVTIME': 'Time', 'PREMAINQTY': "Q'ty", 'DESCRIPTION': 'Scrap Code',
                       'DECODE(SUBSTR(C.CURPRCDKIND,1,': 'InstNum', 'LOTTYPE': 'LT'}, inplace=True)
    df['Scrap Reason'] = df.apply(lambda x: scrap_comment(x.name, x['Time']), axis=1)
    return df


def lot_event_search(lot_id):
    if verify_id(lot_id):
        sql = f"""SELECT A.EVTIME, A.EVUSER, B.EV_REASON_NAME, C.EV_TYPE_NAME, A.VARIANTFLD1, A.VARIANTFLD2 
                  FROM PLLDBA.HIST_LOTEVENTS A, EDADBA.CIM_EVREASON B, EDADBA.CIM_EVTYPE C 
                  WHERE A.LOTID = '{lot_id}'
                  AND A.EVREASON = B.EV_REASON_NO
                  AND A.EVTYPE = C.EV_TYPE
                  And A.Evtype IN ('COMM','HOLD','RELS')
                  ORDER BY A.EVTIME"""
        df = pd.read_sql(sql, con, index_col=['EVTIME'])
    else:
        df = pd.DataFrame()
    if df.empty:
        return df
    df.rename(columns={'EVUSER': 'USER', 'EV_REASON_NAME': 'REASON', 'EV_TYPE_NAME': 'TYPE',
                       "VARIANTFLD1": 'COMMENT', 'VARIANTFLD2': 'HOLDCODE'},
              # index={'EVTIME': 'TIME'},
              inplace=True)
    df.rename_axis(['Time'], inplace=True)
    return df


def wafer_history(_id, search_wafer_history=True):
    if verify_id(_id):
        if search_wafer_history:
            sql_str = f"WHERE B.COMPIDS='{_id}'"
        else:
            sql_str = f"WHERE A.LOTID='{_id}'"
        sql = f"""SELECT C.CURPRCDCURINSTNUM, C.STAGE, C.EQPID, C.CURMAINQTY,
                  C.PREVHISTTIME,C.TIMEREV, C.CURPRCDID, NVL(D.TITLE,' '), 
                  C.RECPID, C.EMPIDTRACKIN, C.EMPIDTRACKOUT, A.LOTID 
                  FROM PLLDBA.HIST_COMPONENTS A, PLLDBA.ACTL_COMPONENTS B, PLLDBA.HIST C, PLLDBA.RECP D 
                  {sql_str} 
                  AND A.COMPSTATE IN (1,4,5) 
                  AND A.LOTID=B.LOTID 
                  AND A.SEQNUM=B.SEQNUM  
                  AND C.LOTID = A.LOTID   
                  AND A.TIMEREV = C.TIMEREV  
                  AND C.RECPID = D.RECPID(+)
                  ORDER BY A.TIMEREV"""
        df = pd.read_sql(sql, con)
    else:
        df = pd.DataFrame()
    if df.empty:
        return df
    if not search_wafer_history:
        df.drop_duplicates(subset=['PREVHISTTIME', 'TIMEREV'], inplace=True)
    df.rename(columns={'STAGE': 'Stage', 'CURMAINQTY': 'LotQty', 'EQPID': 'Equipment',
                       "PREVHISTTIME": 'TrackInTime', 'EMPIDTRACKIN': 'TrackInUser',
                       "TIMEREV": 'TrackOutTime', 'EMPIDTRACKOUT': 'TrackOutUser',
                       "NVL(D.TITLE,'')": 'Step Description', 'RECPID': 'Recipe',
                       'CURPRCDCURINSTNUM': 'InstNo', 'CURPRCDID': 'Procedure'},
              inplace=True)

    df.index = df.InstNo + '/' + df.Procedure
    return df


def equipment_lot_history(eqp_id, end_date, start_date=None):
    # BETWEEN DATE '{str(start_date)}' AND DATE'{str(end_date)}'
    if start_date is None:
        start_date = last_week(end_date)
    sql = f"""SELECT A.*,C.COMPIDS FROM
              (SELECT LOTID,MAX(TIMEREV) TIMEREV,MAX(PREVHISTTIME) PREVHISTTIME,LOTTYPE LT,PARTID,CURMAINQTY QTY,TRACKINTIME,EMPIDTRACKIN,TRACKOUTTIME,EMPIDTRACKOUT,CURPRCDID,CURPRCDCURINSTNUM INSTNO,RECPID,EQPID FROM PLLDBA.HIST
              WHERE TRACKINTIME BETWEEN DATE '{str(start_date)}' AND DATE'{str(
        end_date)}' AND EQPID = '{eqp_id}' AND TRACKOUTTIME IS NOT NULL
              GROUP BY LOTID,LOTTYPE,PARTID,CURMAINQTY,TRACKINTIME,EMPIDTRACKIN,TRACKOUTTIME,EMPIDTRACKOUT,CURPRCDID,CURPRCDCURINSTNUM,RECPID,EQPID) A,
              (SELECT LOTID,TIMEREV,SEQNUM,COMPSTATE FROM PLLDBA.HIST_COMPONENTS WHERE COMPSTATE=1) B,
              (SELECT LOTID,SEQNUM,COMPIDS FROM PLLDBA.ACTL_COMPONENTS) C
              WHERE A.LOTID=B.LOTID(+) AND A.TIMEREV=B.TIMEREV(+) AND B.LOTID=C.LOTID AND B.SEQNUM=C.SEQNUM
              ORDER BY A.TRACKINTIME"""
    df = pd.read_sql(sql, con)
    return df


def conti_defcode():
    # con = pyodbc.connect(connstr)
    connection()
    sql = "SELECT * FROM EDADBA.CONTI_DEFCODE"
    df = pd.read_sql(sql, con, index_col='CODE')
    # df = df[~df.index.duplicated(keep=False)]  # drop duplicated code like 0000
    # con.close()
    return df


def lot_to_wafer(lotid):
    """one to one"""
    # con = pyodbc.connect(connstr)
    if verify_id(lotid):
        sql = "SELECT COMPIDS FROM PLLDBA.ACTL_COMPONENTS " \
              "WHERE LOTID = '{}' " \
              "AND COMPSTATE IN (1,4,5)".format(lotid)
        crsr = con.cursor()
        crsr.execute(sql)
        result = crsr.fetchone()[0]
        # con.close()
        return result


def lot_to_wafers(lot):
    """one to many, lot 1NJK859 to 1NJK859.1, .2, .3 .... etc"""
    if verify_id(lot, is_id=False):
        sql = "SELECT COMPIDS FROM PLLDBA.ACTL_COMPONENTS " \
              "WHERE COMPIDS LIKE '{}.%' " \
              "AND COMPSTATE IN (1,4,5)".format(lot)
        crsr = con.cursor()
        crsr.execute(sql)
        result = [tup[0] for tup in crsr.fetchall()]
        return result


def lot_to_lots(lot):
    """one to many, lot 1NJK859 to 1NJK859.1, .2, .3 .... etc"""
    if verify_id(lot, is_id=False):
        sql = "SELECT DISTINCT LOTID FROM PLLDBA.ACTL_COMPONENTS " \
              "WHERE COMPIDS LIKE '{}.%' " \
              "AND COMPSTATE IN (1,4,5)".format(lot)
        crsr = con.cursor()
        crsr.execute(sql)
        result = [tup[0] for tup in crsr.fetchall()]
        return result


def wafer_to_lot(waferid):
    # con = pyodbc.connect(connstr)
    if verify_id(waferid):
        sql = "SELECT LOTID, COMPIDS FROM PLLDBA.ACTL_COMPONENTS " \
              "WHERE COMPIDS = '{}' " \
              "AND COMPSTATE IN (1,4,5)".format(waferid)
        crsr = con.cursor()
        crsr.execute(sql)
        result = crsr.fetchone()[0]
        # con.close()
        return result


def code_by_part(part, code):
    # con = pyodbc.connect(connstr)
    sql = "SELECT  B.WAFER, B.PARTID, B.TESTDATE, A.X, A.Y From EDADBA.MAP_DATAD A, EDADBA.MAP_DATAW B " \
          "Where A.Link=B.Link " \
          "AND B.PARTID LIKE '{}%' " \
          "AND A.CAT LIKE '{}'".format(part, code)
    df = pd.read_sql(sql, con)
    # con.close()
    return df


def main():
    # df = scrap_lot(end_date=dt.date(2018, 1, 19), start_date=dt.date(2018, 1, 12))
    # a = scrap_comment('1NIA005.2', df.loc['1NIA005.2', 'Time'])
    # print(a)
    # df = cp_map('1NJJ618.19', bin_codes_tup=('2', '3', '4'))
    df = wafer_history('1NKE531.11')
    print(df.to_string())
    return df


# def cp_map(wafer_id, part, cp_stage='CP3', bin_codes_tup='all'):
#     """Query CP3 map by wafer_id"""
#     # if part.startswith('AP197'):
#     #     cp_stage = 'CP1'
#     if isinstance(bin_codes_tup, tuple):
#         if len(bin_codes_tup) == 1:
#             bin_codes_tup = "('{}')".format(bin_codes_tup[0])
#         key = f' AND cat IN {bin_codes_tup}'
#     if verify_id(wafer_id) and verify_part(part):
#         lot_id = wafer_id.split('.')[0] + '%'
#         wafer_n = wafer_id.split('.')[1]
#         sql = f"""select n.x, n.y, m.bin from
#                 (select b.die, b.x, b.y from
#                 (select link from EDADBA.MAP_SPECH where part like '{part}' and testtype = '{cp_stage}') a,
#                 (select link, die, x, y from EDADBA.MAP_SPECR) b
#                 where a.link = b.link) n,
#                 (select d.die, d.bin from
#                 (select link from EDADBA.CP_DATAW
#                 where lotid like '{lot_id}' and wafer = {wafer_n} and testtype ='{cp_stage}') c,
#                 (select link, die, bin from EDADBA.CP_DATAD) d
#                 where c.link = d.link) m
#                 where m.die = n.die"""
#         df = pd.read_sql(sql, con)
#         # df.rename(columns={'X': 'Y', 'Y': "X"}, inplace=True)  # XY inverse = =
#         if df.empty:
#             if cp_stage == "CP3":
#                 df = cp_map(wafer_id, part, cp_stage='CP2')
#             elif cp_stage == 'CP2':
#                 df = cp_map(wafer_id, part, cp_stage='CP1')
#             else:
#                 raise CPDataNotFoundError(wafer_id)
#         return df
#     else:
#         return pd.DataFrame()

if __name__ == '__main__':
    main()
