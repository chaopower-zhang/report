# -*- coding: utf-8 -*-

"""
刘帅要求
所有肿瘤检测结果为MSI-H
RET阳性患者，甲状腺髓样癌跟非小细胞肺癌
T790M阳性的非小细胞肺癌患者
ALK阳性的非小细胞肺癌患者
HER-2阳性的乳腺癌患者
"""

import datetime
import openpyxl
from . import config

config = config.config()
mreport = config['mreport']

if mreport:
    path = '/home/zc/test/specialngs.xlsx'
else:
    path = '/home/zc/specialngs/specialngs.xlsx'

geneinfo = {
    'RET': '甲状腺髓样癌,非小细胞肺癌',
    'T790M': '非小细胞肺癌',
    # 'ALK':'非小细胞肺癌',
    'HER-2': '乳腺癌',
    'FGFR': '肝癌，胆管癌，胆囊癌'  # 所有肿瘤
}


def write_excel_xlsx(value):
    """
    写入结果
    :param value:
    :return:
    """
    wb = openpyxl.load_workbook(path)
    ws = wb['Sheet1']
    ws.append(value)
    wb.save(path)


def writefile(tbgene, treatresult, msi, sampledata):
    """
    读写结果
    :param sampledata:
    :param tbgene:
    :param treatresult:
    :param msi:
    :return:
    """
    if not tbgene and msi != 'MSI-H':
        return
    datainfo = datetime.datetime.now().strftime('%Y-%m-%d')
    unitagentname = sampledata['unitAgentName']
    agentname = sampledata['agentName']
    samplesn = sampledata['sampleSn']
    name = sampledata['name']
    gender = sampledata['gender']
    age = sampledata['age']
    itemname = sampledata['itemName']
    itemtype = 'unknow'
    for gene in tbgene:
        if gene in geneinfo:
            if treatresult in geneinfo[gene].split(','):
                result = gene + '阳性'
                text = [datainfo, unitagentname, agentname, samplesn, name, gender, age,
                        treatresult, itemname, itemtype, result]
                write_excel_xlsx(text)

    if msi == 'MSI-H':
        result = 'MSI-H'
        text = [datainfo, unitagentname, agentname, samplesn, name, gender, age,
                treatresult, itemname, itemtype, result]
        write_excel_xlsx(text)
