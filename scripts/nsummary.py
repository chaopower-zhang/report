# -*- coding: utf-8 -*-


import os
import sys
from docxtpl import DocxTemplate

rootpath = os.path.dirname((os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
sys.path.append(rootpath)
from tools import database
from tools import common

basepath = os.path.join(rootpath, 'database', 'template')
detectinfopath = os.path.join(basepath, 'detectinfo', '检测信息.docx')


def detect_info(sd, detectinfo):
    """
    检测信息
    :param sd:
    :param detectinfo:
    :return:
    """
    width = [2.25, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25]
    p = sd.paragraphs[2]
    pt = '{}是一种基于高通量测序（NGS）的检测方法，'.format(detectinfo['reportItemName'])
    if detectinfo['geneBx']:
        if not detectinfo['geneKb'] and not detectinfo['geneHl']:
            tp1 = '热点'
        else:
            tp1 = ''
        pt += '可鉴定{}个与{}高度相关的基因中的{}基因变异'.format(
            len(detectinfo['geneTb'].split(',')), detectinfo['treatresult'], tp1)
        tp2 = '，和'
        bxtable = sd.add_table(rows=1, cols=8, style='1sum_detect_info1')
        common.adjust_table_width(bxtable, width)
        bxtable.cell(0, 0).merge(bxtable.cell(0, 7))
        bxtable.cell(0, 0).paragraphs[0].text = '靶向药物相关基因{}个'.format(
            len(detectinfo['geneBx'].split(',')))
        bxtable.cell(0, 0).paragraphs[0].style = '0detecinfo_table1'
        if detectinfo['geneBx']:
            genetb = detectinfo['geneBx'].split(',')
            tbtitle = sd.add_table(rows=1, cols=8, style='1sum_detect_info2')
            tbtitle.cell(0, 0).paragraphs[0].text = '检测基因单碱基变异、小片段插入缺失'
            tbtitle.cell(0, 0).paragraphs[0].style = '0detecinfo_table2'
            common.adjust_table_width(tbtitle, width)
            tbtitle.cell(0, 0).merge(tbtitle.cell(0, 7))
            tbtablenum = len(genetb)
            tbrow = tbtablenum // 8 if tbtablenum % 8 == 0 else (tbtablenum // 8) + 1
            tbtable = sd.add_table(tbrow, cols=8, style='1sum_detect_info3')
            for i in range(0, tbrow):
                tbgene = genetb[i * 8:i * 8 + 8]
                for index, gene in enumerate(tbgene):
                    tbtable.cell(i, index).paragraphs[0].text = gene
                    tbtable.cell(i, index).paragraphs[0].style = '0detecinfo_table3'
            common.adjust_table_width(tbtable, width)
        if detectinfo['geneCp']:
            cptitle = sd.add_table(rows=1, cols=8, style='1sum_detect_info2')
            cptitle.cell(0, 0).paragraphs[0].text = '检测基因重排'
            cptitle.cell(0, 0).paragraphs[0].style = '0detecinfo_table2'
            common.adjust_table_width(cptitle, width)
            cptitle.cell(0, 0).merge(cptitle.cell(0, 7))
            genecp = detectinfo['geneCp'].split(',')
            cptablenum = len(genecp)
            cprow = cptablenum // 8 if cptablenum % 8 == 0 else (cptablenum // 8) + 1
            cptable = sd.add_table(rows=cprow, cols=8, style='1sum_detect_info3')
            for i in range(0, cprow):
                fugene = genecp[i * 8:i * 8 + 8]
                for index, gene in enumerate(fugene):
                    cptable.cell(i, index).paragraphs[0].text = gene
                    cptable.cell(i, index).paragraphs[0].style = '0detecinfo_table3'
            common.adjust_table_width(cptable, width)
        if detectinfo['geneKb']:
            genekb = detectinfo['geneKb'].split(',')
            kbtitle = sd.add_table(rows=1, cols=8, style='1sum_detect_info2')
            kbtitle.cell(0, 0).paragraphs[0].text = '检测基因拷贝数变异'
            kbtitle.cell(0, 0).paragraphs[0].style = '0detecinfo_table2'
            common.adjust_table_width(kbtitle, width)
            kbtitle.cell(0, 0).merge(kbtitle.cell(0, 7))
            kbtablenum = len(genekb)
            kbrow = kbtablenum // 8 if kbtablenum % 8 == 0 else (kbtablenum // 8) + 1
            kbtable = sd.add_table(rows=kbrow, cols=8, style='1sum_detect_info3')
            for i in range(0, kbrow):
                kbgene = genekb[i * 8:i * 8 + 8]
                for index, gene in enumerate(kbgene):
                    kbtable.cell(i, index).paragraphs[0].text = gene
                    kbtable.cell(i, index).paragraphs[0].style = '0detecinfo_table3'
            common.adjust_table_width(kbtable, width)
    else:
        tp2 = '，可鉴定'

    if detectinfo['geneMy']:
        mylist = detectinfo['geneMy'].split(',')
        mygenenum = len(mylist)
        mytitle = sd.add_table(rows=1, cols=8, style='1sum_detect_info1')
        mytitle.cell(0, 0).paragraphs[0].text = '免疫治疗相关基因{}个'.format(mygenenum)
        mytitle.cell(0, 0).paragraphs[0].style = '0detecinfo_table1'
        common.adjust_table_width(mytitle, width)
        mytitle.cell(0, 0).merge(mytitle.cell(0, 7))
        myrow = mygenenum // 8 if mygenenum % 8 == 0 else (mygenenum // 8) + 1
        mygenetale = sd.add_table(rows=myrow, cols=8, style='1sum_detect_info3')
        for i in range(0, myrow):
            mygene = mylist[i * 8:i * 8 + 8]
            for index, gene in enumerate(mygene):
                mygenetale.cell(i, index).paragraphs[0].text = gene
                mygenetale.cell(i, index).paragraphs[0].style = '0detecinfo_table3'
        common.adjust_table_width(mygenetale, width)

    if detectinfo['geneHrd']:
        hrdlist = detectinfo['geneHrd'].split(',')
        hrdgenenum = len(hrdlist)
        hrdtitle = sd.add_table(rows=1, cols=8, style='1sum_detect_info1')
        hrdtitle.cell(0, 0).paragraphs[0].text = 'HRD-DDR相关基因{}个'.format(hrdgenenum)
        hrdtitle.cell(0, 0).paragraphs[0].style = '0detecinfo_table1'
        common.adjust_table_width(hrdtitle, width)
        hrdtitle.cell(0, 0).merge(hrdtitle.cell(0, 7))
        hrdrow = hrdgenenum // 8 if hrdgenenum % 8 == 0 else (hrdgenenum // 8) + 1
        hrdgenetale = sd.add_table(rows=hrdrow, cols=8, style='1sum_detect_info3')
        for i in range(0, hrdrow):
            hrdgene = hrdlist[i * 8:i * 8 + 8]
            for index, gene in enumerate(hrdgene):
                hrdgenetale.cell(i, index).paragraphs[0].text = gene
                hrdgenetale.cell(i, index).paragraphs[0].style = '0detecinfo_table3'
        common.adjust_table_width(hrdgenetale, width)

    if detectinfo['geneYc']:
        yclist = detectinfo['geneYc'].split(',')
        ycgenenum = len(yclist)
        yctitle = sd.add_table(rows=1, cols=8, style='1sum_detect_info1')
        yctitle.cell(0, 0).paragraphs[0].text = '肿瘤遗传易感相关基因{}个'.format(ycgenenum)
        yctitle.cell(0, 0).paragraphs[0].style = '0detecinfo_table1'
        common.adjust_table_width(yctitle, width)
        yctitle.cell(0, 0).merge(yctitle.cell(0, 7))
        ycrow = ycgenenum // 8 if ycgenenum % 8 == 0 else (ycgenenum // 8) + 1
        ycgenetale = sd.add_table(rows=ycrow, cols=8, style='1sum_detect_info3')
        for i in range(0, ycrow):
            ycgene = yclist[i * 8:i * 8 + 8]
            for index, gene in enumerate(ycgene):
                ycgenetale.cell(i, index).paragraphs[0].text = gene
                ycgenetale.cell(i, index).paragraphs[0].style = '0detecinfo_table3'
        common.adjust_table_width(ycgenetale, width)

    if detectinfo['geneHl']:
        pt += '{}{}个与化疗药物代谢相关基因中的基因多态性'.format(tp2, len(detectinfo['geneHl'].split(',')))
        chem = detectinfo['geneHl'].split(',')
        chemnum = len(chem)
        hltitle = sd.add_table(rows=1, cols=8, style='1sum_detect_info1')
        hltitle.cell(0, 0).paragraphs[0].text = '化疗药物相关基因{}个'.format(chemnum)
        hltitle.cell(0, 0).paragraphs[0].style = '0detecinfo_table1'
        common.adjust_table_width(hltitle, width)
        hltitle.cell(0, 0).merge(hltitle.cell(0, 7))
        chemtale = sd.add_table(rows=1, cols=8, style='1sum_detect_info2')
        chemtale.cell(0, 0).paragraphs[0].text = '检测基因多态性'
        chemtale.cell(0, 0).paragraphs[0].style = '0detecinfo_table2'
        common.adjust_table_width(chemtale, width)
        chemtale.cell(0, 0).merge(chemtale.cell(0, 7))
        hlrow = chemnum // 8 if chemnum % 8 == 0 else (chemnum // 8) + 1
        hlgenetale = sd.add_table(rows=hlrow, cols=8, style='1sum_detect_info3')
        for i in range(0, hlrow):
            chemgene = chem[i * 8:i * 8 + 8]
            for index, gene in enumerate(chemgene):
                hlgenetale.cell(i, index).paragraphs[0].text = gene
                hlgenetale.cell(i, index).paragraphs[0].style = '0detecinfo_table3'
        common.adjust_table_width(hlgenetale, width)
    catecode = detectinfo['cateCode'].split(',')
    if 'geneWwx' in catecode:
        pt += ', 及微卫星状态'
    pt += '。检测基因列表如下：'
    if 'geneHl' in catecode and len(catecode) == 1:
        pt = 'OncoDrug-Seq™化疗药物用药基因检测是一种基于MassARRAY，' \
             'qPCR和一代测序的检测方法，可鉴定 33 个与化疗/内分泌药物代谢相关基因中的基因多态性。检测基因列表如下：'
    common.add_run_self(p, pt, font_size=8, font_color=(128, 128, 128), alignment='left')
    sd.add_paragraph()
    pt2 = '检测方法 本检测基于液相探针杂交法的核酸序列靶向捕获及高通量测序技术，测序平台为Illumina NextSeq500/NovaSeq。' \
          '检测覆盖100%的碱基替换突变 (95%CI=82-100) 及95%的小片段插入缺失突变 (95%CI=98.5-100)。数据通过BWA软件' \
          '与人类参考基因组进行比对，Variant calls采用GATK软件分析。数据校准与突变注释使用的软件为TOPGEN自主知识产权分析软件' \
          '（version 20190115），突变注释的主要参考数据库包括Clinvar (version 20181225)、Intervar (version 20180118)、' \
          'COSMIC (version 83)、1000g2015aug (version 20150824)、EXAC03 (version 20160423)、' \
          'dbnsfp35a (version 20180921)、avsnp (version 20170929)、OncoKB (version 1.15)、 ' \
          'PharmGKB (version 4.0)等、TOPGEN自建中国人群数据库（version即时更新）等。'
    sd.add_paragraph(pt2, style='0方块列表')
    pt3 = '检测限 对于单碱基变异，小片段插入缺失，基因融合等变异类型，检测灵敏度≤1%'
    sd.add_paragraph(pt3, style='0方块列表')
    return sd


def bigpannel_detect_info(sd):
    """
    大panel的检测信息
    :param sd:
    :return:
    """
    p1 = sd.paragraphs[2]
    pt1 = ' OncoDrug-Seq™泛癌种基因检测是一种基于高通量测序（NGS）的检测方法，可鉴定数百个与肿瘤高度相关的基因中的基因变异。' \
          '检测基因列表及覆盖范围见后文附录。'
    common.add_run_self(p1, pt1, font_size=8, font_color=(128, 128, 128), alignment='left')
    pt2 = '检测方法 本检测基于液相探针杂交法的核酸序列靶向捕获及高通量测序技术，测序平台为Illumina NextSeq500/NovaSeq。' \
          '检测覆盖100%的碱基替换突变 (95%CI=82-100) 及95%的小片段插入缺失突变 (95%CI=98.5-100)。数据通过BWA软件' \
          '与人类参考基因组进行比对，Variant calls采用GATK软件分析。数据校准与突变注释使用的软件为TOPGEN自主知识产权分析软件' \
          '（version 20190115），突变注释的主要参考数据库包括Clinvar (version 20181225)、Intervar (version 20180118)、' \
          'COSMIC (version 83)、1000g2015aug (version 20150824)、EXAC03 (version 20160423)、' \
          'dbnsfp35a (version 20180921)、avsnp (version 20170929)、OncoKB (version 1.15)、 ' \
          'PharmGKB (version 4.0)等、TOPGEN自建中国人群数据库（version即时更新）等。'
    sd.add_paragraph(pt2, style='0方块列表')
    pt3 = '检测限 对于单碱基变异，小片段插入缺失，基因融合等变异类型，检测灵敏度≤1%'
    sd.add_paragraph(pt3, style='0方块列表')
    return sd


def ycwcd_detect_info(sd):
    """
    遗传胃肠道检测信息
    :param sd:
    :return:
    """
    p1 = sd.paragraphs[2]
    pt1 = ' 遗传性胃肠道肿瘤基因检测是一种基于高通量测序（NGS）的检测方法，可鉴定18个与胃肠道肿瘤高度相关的基因中可能引起遗传风险的基因变异。'
    common.add_run_self(p1, pt1, font_size=8, font_color=(128, 128, 128), alignment='left')
    pt2 = '检测方法 本检测基于多重聚合酶链反应法的核酸序列靶向捕获及高通量测序技术，测序平台为Illumina NextSeq500/NovaSeq。' \
          'NGS检测覆盖APC、BLM、BMPR1A、CHEK2、EPCAM、GALNT12、GREM1、MLH1、MSH2、MSH6、MUTYH、PMS2、POLD1、' \
          'POLE、PTEN、SMAD4、STK11、TP53等18个基因的外显子及其邻近±10bp 内含子区所有变异（包括点突变、小片段插入缺失），' \
          '不包括基因组结构变异（例如大片段杂合缺失、复制与倒位重排）、大片段杂合插入突变及位于基因调节区或深度内含子区的突变。' \
          '数据通过BWA-MEN软件与人类参考基因组进行比对，Variant calls采用GATK软件分析。数据校准与突变注释使用的软件为鼎晶' \
          '自主知识产权分析软件，突变注释的主要参考数据库包括Clinvar (version20191013)、Intervar (version 20180118)、' \
          'COSMIC (version 83)、1000g version 20150824、EXAC03 (version 20160423)、dbnsfp (version 20180921)、' \
          'avsnp (version 20170929)、OncoKB (version 1.15)、 PharmGKB (version 4.0)等、' \
          '鼎晶自建中国人群数据库（version即时更新）等。'
    sd.add_paragraph(pt2, style='0方块列表')
    pt3 = '检测限 对于单碱基变异，小片段插入缺失，基因融合等变异类型，检测灵敏度≤1%'
    sd.add_paragraph(pt3, style='0方块列表')
    return sd


def ycb_detect_info(sd):
    """
    遗传BRCA基因检测信息
    :param sd:
    :return:
    """
    p1 = sd.paragraphs[2]
    pt1 = ' BRCA1/2基因检测是一种基于高通量测序（NGS）的检测方法，可鉴定BRCA1和BRCA2基因中发生的基因变异。 '
    common.add_run_self(p1, pt1, font_size=8, font_color=(128, 128, 128), alignment='left')
    pt2 = '检测方法 本检测基于多重聚合酶链反应法的核酸序列靶向捕获及高通量测序技术，测序平台为Illumina NextSeq500/NovaSeq。' \
          'NGS检测覆盖APC、BLM、BMPR1A、CHEK2、EPCAM、GALNT12、GREM1、MLH1、MSH2、MSH6、MUTYH、PMS2、POLD1、' \
          'POLE、PTEN、SMAD4、STK11、TP53等18个基因的外显子及其邻近±10bp 内含子区所有变异（包括点突变、小片段插入缺失），' \
          '不包括基因组结构变异（例如大片段杂合缺失、复制与倒位重排）、大片段杂合插入突变及位于基因调节区或深度内含子区的突变。' \
          '数据通过BWA-MEN软件与人类参考基因组进行比对，Variant calls采用GATK软件分析。数据校准与突变注释使用的软件为鼎晶' \
          '自主知识产权分析软件，突变注释的主要参考数据库包括Clinvar (version20191013)、Intervar (version 20180118)、' \
          'COSMIC (version 83)、1000g version 20150824、EXAC03 (version 20160423)、dbnsfp (version 20180921)、' \
          'avsnp (version 20170929)、OncoKB (version 1.15)、 PharmGKB (version 4.0)等、' \
          '鼎晶自建中国人群数据库（version即时更新）等。'
    sd.add_paragraph(pt2, style='0方块列表')
    pt3 = '检测限 对于单碱基变异，小片段插入缺失，基因融合等变异类型，检测灵敏度≤1%'
    sd.add_paragraph(pt3, style='0方块列表')
    return sd


def ycbplus_detect_info(sd):
    """
    遗传BRCA拓展版基因检测信息
    :param sd:
    :return:
    """
    p1 = sd.paragraphs[2]
    pt1 = ' BRCA1/2基因检测拓展版是一种基于高通量测序（NGS）和多重连接探针扩增技术（MLPA）的检测方法，' \
          '可鉴定BRCA1和BRCA2基因中发生的基因变异。 '
    common.add_run_self(p1, pt1, font_size=8, font_color=(128, 128, 128), alignment='left')
    pt2 = '检测方法 本检测基于多重聚合酶链反应法的核酸序列靶向捕获及高通量测序技术，测序平台为Illumina NextSeq500/NovaSeq。' \
          'NGS检测覆盖BRCA1、BRCA2 基因的外显子及其邻近±10bp 内含子区所有变异（包括点突变、小片段插入缺失），' \
          '不包括基因组结构变异（例如大片段杂合缺失、复制与倒位重排）、大片段杂合插入突变及位于基因调节区或深度内含子区的突变。' \
          '数据通过BWA-MEN软件与人类参考基因组进行比对，Variant calls采用GATK软件分析。数据校准与突变注释使用的软件为' \
          '鼎晶自主知识产权分析软件，突变注释的主要参考数据库包括Clinvar (version20191013)、Intervar (version 20180118)、' \
          'COSMIC (version 83)、1000g version 20150824、EXAC03 (version 20160423)、dbnsfp (version 20180921)、' \
          'avsnp (version 20170929)、OncoKB (version 1.15)、 PharmGKB (version 4.0)等、' \
          '鼎晶自建中国人群数据库（version即时更新）等。'
    sd.add_paragraph(pt2, style='0方块列表')
    pt3 = '检测限 对于单碱基变异，小片段插入缺失，基因融合等变异类型，检测灵敏度≤1%'
    sd.add_paragraph(pt3, style='0方块列表')
    return sd


def summary(tpl, detectgene):
    """
    汇总块
    :param tpl:
    :param sampleinfo:
    :param detectgene:
    :return:
    """

    catecode = detectgene['cateCode'].split(',')
    sd = tpl.new_subdoc(detectinfopath)
    sd.add_paragraph()
    if 'bigPannel' in catecode:
        sd = bigpannel_detect_info(sd)
    elif 'ycWcd' in catecode:
        sd = ycwcd_detect_info(sd)
    elif 'ycB' in catecode:
        sd = ycb_detect_info(sd)
    elif 'ycBplus' in catecode:
        sd = ycbplus_detect_info(sd)
    elif 'thys' in catecode:
        pass
        return
    else:
        sd = detect_info(sd, detectgene)
    sd.add_page_break()
    return sd


if __name__ == '__main__':
    tplpath = os.path.join(basepath, 'blank.docx')
    detectinfo = database.detectinfo()
    for itemid in detectinfo:
        tpl = DocxTemplate(tplpath)
        context = summary(tpl, detectinfo[itemid])
        word = os.path.join(basepath, 'detectinfo', itemid + '.docx')
        context = {'context': context}
        tpl.render(context, autoescape=True)
        tpl.save(word)
    print(f'检测信息更新成功')
