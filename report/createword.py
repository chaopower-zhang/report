# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
"""
基本信息与结果汇总模块
"""

import os
from tools import common
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
import logging

# 生成结果路径
path = os.path.join(common.rootpath, 'result')

# 模版路径
templatepath = os.path.join(common.rootpath, 'database', 'template')  

# 日志
logger = logging.getLogger('main.sub') 

tplpath = os.path.join(templatepath, 'nreport.docx')


def thys(tpl):
    """
    甲状腺
    :param tpl:
    :return:
    """
    thydocx = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'database',
                           'template', 'summary', 'thy.docx')
    sd = tpl.new_subdoc(thydocx)
    return sd


def summary(tpl, sampleinfo):
    """
    汇总块
    :param tpl:
    :param sampleinfo:
    :return:
    """
    itemid = sampleinfo['c']['itemId']
    treatid = sampleinfo['c']['treatID']
    catecode = sampleinfo['c']['catecode']
    sigma = sampleinfo['sigma']
    content = list()
    content.append(tpl.new_subdoc(os.path.join(templatepath, '1detectinfo', f'{itemid}.docx')))
    if 'ycWcd' in catecode or 'ycB' in catecode or 'ycBplus' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '2qc', f'ycqc.docx')))
    elif 'thys' in catecode:
        pass
    else:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '2qc', f'qc.docx')))

        if 'TS05' in treatid and 'addWwx' not in catecode and 'geneWwx' in catecode:
            catecode.append('addWwx')

        content.append(tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'1title.docx')))
        if 'geneBx' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'2bx.docx')))

        # hrd
        if 'geneHrd' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'4hrd.docx')))

        # hrr
        if 'geneHrr' in catecode:
            if sigma:
                content.append(tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'4_1hrd_sigma.docx')))
            else:
                content.append(tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'4_1_1hrd_no_sigma.docx')))

        # immune
            if 'geneMy' in catecode:
                sd = tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'3my.docx'))
                content.append(sd)
            else:
                if 'geneWwx' in catecode:
                    sd = tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'3_1my.docx'))
                    content.append(sd)
                else:
                    pass

        if 'addWwx' in catecode:
            sd = tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'3_1_1my.docx'))
            content.append(sd)
        # heredity
        if 'geneYc' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'5yc.docx')))

        # chem
        if 'geneHl' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '3summaryresult', f'6hl.docx')))
    return content


def medicaltip(tpl, sampleinfo):
    """
    用药提示
    :param tpl:
    :param sampleinfo:
    :param nccn:
    :param detectgene:
    :return:
    """
    content = []
    catecode = sampleinfo['c']['catecode']
    if 'ycWcd' in catecode or 'ycB' in catecode or 'ycBplus' in catecode or 'thys' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', 'wcd.docx')))
    if 'geneBx' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', 'bx.docx')))
        if sampleinfo['nccn']:
            treadid = sampleinfo['nccn']['treadid']
            content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', 'nccn', f'{treadid}.docx')))
    if 'geneHrd' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'hrd.docx')))
    if 'geneHrr' in catecode:
        if sampleinfo['sigma']:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'hrr.docx')))
            content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'sigma.docx')))
        else:
            pass
            content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'hrr_no_sigma.docx')))
    # 肝胆肿瘤
    if sampleinfo['hepa']:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'hepa.docx')))

    if 'geneMy' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'my.docx')))

        if sampleinfo['pdl1']:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'pdl1.docx')))

    if 'geneTmb' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'tmb.docx')))

    if 'geneWwx' in catecode:
        if 'geneMy' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'msi.docx')))
        else:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'msi_simple.docx')))

    if 'geneTnb' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'tnb.docx')))
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'addtnb.docx')))

    if 'geneHl' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '4medicaltip', f'hl.docx')))
    return content


def detectresult(tpl, sampleinfo):
    content = []
    catecode = sampleinfo['c']['catecode']
    if 'geneYc' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '5detectresult', 'yc.docx')))
    if 'geneBx' in catecode:
        if 'geneTb' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '5detectresult', 'tb.docx')))
        if 'geneCp' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '5detectresult', 'cp.docx')))
        if 'geneKb' in catecode:
            content.append(tpl.new_subdoc(os.path.join(templatepath, '5detectresult', 'kb.docx')))
    content.append(tpl.new_subdoc(os.path.join(templatepath, '5detectresult', 'sign.docx')))
    return content


def detectdetail(tpl, sampleinfo):
    """
    检测结果详解
    :param tpl:
    :param sampleinfo:
    :param detectgene:
    :return:
    """
    templatepath = os.path.join(common.rootpath, 'database', 'template')
    catecode = sampleinfo['c']['catecode']
    content = []
    if 'geneBx' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '6detectdetail', 'tb.docx')))
        content.append(tpl.new_subdoc(os.path.join(templatepath, '6detectdetail', 'clinical.docx')))
    if 'geneHrr' in catecode and sampleinfo['sigma']:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '6detectdetail', 'sigma.docx')))
    if 'geneTmb' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '6detectdetail', 'tnb.docx')))
    if 'geneWwx' in catecode and sampleinfo['msi']:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '6detectdetail', 'msi.docx')))
    if 'geneYc' in catecode or 'ycWcd' in catecode:
        content.append(tpl.new_subdoc(os.path.join(templatepath, '6detectdetail', 'yc.docx')))
    return content


def reference(tpl, sampleinfo):
    """
    参考信息
    :param tpl:
    :param itemid:
    :return:
    """
    result = list()
    itemid = sampleinfo['c']['itemId']
    blankdocx = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'database',
                           'template', '7reference', 'blank.docx')
    refdocx = os.path.join(os.path.dirname((os.path.dirname(os.path.abspath(__file__)))), 'database',
                           'template', '7reference', itemid+'.docx')
    if os.path.isfile(refdocx):
        result.append(tpl.new_subdoc(blankdocx))
        result.append(tpl.new_subdoc(refdocx))
    return result


def run(sampleinfo):
    tpl = DocxTemplate(tplpath)
    sample = sampleinfo['c']['sampleSn']
    print('{}开始'.format(sample))

    contents = list()
    try:
        contents.extend(summary(tpl, sampleinfo))
        contents.extend(medicaltip(tpl, sampleinfo))
        contents.extend(detectresult(tpl, sampleinfo))
        contents.extend(detectdetail(tpl, sampleinfo))
        contents.extend(reference(tpl, sampleinfo))
    except:
        print(sample)
    sampleinfo['contents'] = contents

    pict = dict()
    for pic in sampleinfo['pic']:
        pict[pic['name']] = InlineImage(tpl, pic['path'], width=Cm(pic['width']), height=Cm(pic['height']))
    sampleinfo['pict'] = pict

    # word 中变量替换

    tpl.render(sampleinfo)
    try:
        tpl.render(sampleinfo)
    except:
        print(sample)
    common.set_updatefields_true(tpl)
    # 保存 word
    word = os.path.join(path, 'word', sample + '.docx')
    tpl.save(word)
    print('{}结束'.format(sample))
    return word

