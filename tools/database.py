# -*- coding: utf-8 -*-

import pandas as pd
import collections
import os
from docxtpl import InlineImage
from docx.shared import Cm


basedir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'database')


def chemical():
    """
    化疗疗效整理
    :return: 返回个疗效字典
    """
    path = os.path.join(basedir, 'baseinfo', 'chemical.xlsx')
    df = pd.read_excel(path, sheet_name='geneChemicalArithmetic', index_col=0)
    genechemicalarithmetic = df.to_dict('index')
    df1 = pd.read_excel(path, sheet_name='chemicalDrugEffect', index_col=1)
    chemicaldrugeffect = df1.to_dict('index')
    result = {'arithmetic': genechemicalarithmetic, 'drugeffect': chemicaldrugeffect}
    return result


def diagnose():
    """
    临床诊断列表
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'diagnosetype.xlsx')
    df = pd.read_excel(path, sheet_name='diagnosetype', index_col=0)
    result = df.to_dict()['type']
    return result


def nccn(treatid):
    """
    NCCN指南基因临床意义
    :param treatid: 适应症id
    :return: 适应症NCCN基因意义
    """
    path = os.path.join(basedir, 'baseinfo', 'nccn.xlsx')
    nccnxlsx = pd.ExcelFile(path)
    nccnlist = nccnxlsx.sheet_names
    if treatid == 'TS01':
        treatid = 'TS0101'
    for tmid in nccnlist:
        if tmid in treatid:
            reuslt = collections.OrderedDict()
            df = pd.read_excel(path, sheet_name=tmid, index_col=0)
            df = df.fillna('NA')
            for index, rows in df.iterrows():
                reuslt[index] = dict(检测范围=rows['检测范围'], 检测意义=rows['检测意义'])
            return reuslt
    else:
        return None


def nnccn():
    """
    NCCN指南基因临床意义
    :return: 适应症NCCN基因意义
    """
    path = os.path.join(basedir, 'baseinfo', 'nccn.xlsx')
    nccnxlsx = pd.read_excel(path, None)
    result = dict()
    for id, df in nccnxlsx.items():
        result[id] = [rows for _, rows in df.to_dict('index').items()]
    return result


def detectinfo():
    """
    检测基因列表
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'panel.xlsx')
    df = pd.read_excel(path, sheet_name='panel', index_col=0, header=0, skiprows=[1])
    df = df.fillna('')
    detectinfodict = df.to_dict('index')
    return detectinfodict


def treatresult():
    """
    panel 对应的适应症
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'treatresult.xlsx')
    df = pd.read_excel(path, sheet_name='treatresult', index_col=0, header=0)
    df = df.fillna('NA')
    result = df.to_dict()
    return result


def hrrcontend():
    """
    hrr 详解内容
    :return:
    """
    path = os.path.join(basedir, 'detectdetail', 'hrr', 'hrr.xlsx')
    df = pd.read_excel(path, sheet_name='hrr', header=0, index_col=0)
    hrrdict = df.to_dict('index')
    return hrrdict


def msicontend():
    """
    msi 详解内容
    :return:
    """
    path = os.path.join(basedir, 'detectdetail', 'msi', 'msicontend.xlsx')
    df = pd.read_excel(path, sheet_name='msicontend', header=0)
    contend = df['contend'].dropna().to_list()
    return contend


def immunedetailpic(tpl):
    """
    免疫详解图片补充
    :return:
    """
    result = dict()
    picpath = dict(
        immunedetaila=os.path.join(basedir, 'detectdetail', 'immune', 'immunea.png'),
        immunedetailb=os.path.join(basedir, 'detectdetail', 'immune', 'immuneb.png'),
        immunedetailc=os.path.join(basedir, 'detectdetail', 'immune', 'immunec.png'),
        immunedetaild=os.path.join(basedir, 'detectdetail', 'immune', 'immuned.png'),
    )
    picwidth = dict(
        immunedetaila=11.38,
        immunedetailb=7.94,
        immunedetailc=11.38,
        immunedetaild=9.6,
    )
    picheight = dict(
        immunedetaila=7.16,
        immunedetailb=6.7,
        immunedetailc=5.4,
        immunedetaild=6.02,
    )
    for pic in ['immunedetaila', 'immunedetailb', 'immunedetailc', 'immunedetaild']:
        result[pic] = InlineImage(tpl, picpath[pic], width=Cm(picwidth[pic]), height=Cm(picheight[pic]))
    return result


def custom():
    """
    定制版报告
    :return:
    """
    path = os.path.join(basedir, 'custom', 'custom.xlsx')
    df = pd.read_excel(path, sheet_name='custom', index_col=0, header=0, skiprows=[1])
    customdict = df.to_dict('index')
    return customdict


def source():
    """
    样品来源
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'source.xlsx')
    df = pd.read_excel(path, index_col=0, header=0)
    sourcedict = df.to_dict('index')
    return sourcedict


def tbtranslate():
    """
    变异名称翻译
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'tbtranslate.xlsx')
    df = pd.read_excel(path, sheet_name='tbtranslate', index_col=0, header=0)
    tbresult = df.to_dict('index')
    return tbresult


def treattran():
    """
    诊断ID与癌种对应表
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'treattran.xlsx')
    df = pd.read_excel(path, sheet_name='treattran', index_col=0, header=0)
    result = df.to_dict('index')
    return result


def agentname():
    """
    医疗机构名称
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'agentname.xlsx')
    df = pd.read_excel(path, sheet_name='agentname', index_col=0, header=0)
    df = df.fillna(' ')
    agresult = df.to_dict('index')
    return agresult


def thyinfo():
    """
    甲状腺基因临床意义
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'thyinfo.xlsx')
    thyresult = dict()
    thygene = pd.read_excel(path, None)
    for index, sheet in thygene.items():
        sheet.set_index('genename', inplace=True)
        tsheet = sheet.to_dict()['geneinfo']
        thyresult[index] = tsheet
    return thyresult


def thsr():
    """
    甲状腺THSR基因
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'THSR.xlsx')
    thsrresult = dict()
    thsrdata = pd.read_excel(path, None)
    for index, sheet in thsrdata.items():
        thsrresult[index] = sheet['hgsGb'].to_list()
    return thsrresult


def specialagent():
    """
    特殊送检单位不出非本癌种信息
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'specialagent.xlsx')
    df = pd.read_excel(path, sheet_name='special')
    result = df['送检单位'].to_list()
    return result


def cmsneedlist():
    """
    报告所需cms系统信息
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'cmsneedlist.xlsx')
    df = pd.read_excel(path)
    result = df['cmsneedlist'].to_list()
    return result


def sheetname():
    """
    ecxcel中sheetnanmes变化
    :return:
    """
    path = os.path.join(basedir, 'baseinfo', 'sheetname.xlsx')
    df = pd.read_excel(path, sheet_name='tbtranslate', index_col=0, header=0)
    sheetnameresult = df.to_dict()['translate']
    return sheetnameresult
