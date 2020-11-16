# -*- coding: utf-8 -*-


import os
import sys
from docxtpl import DocxTemplate
rootpath = os.path.dirname((os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
sys.path.append(rootpath)
from tools import common
from tools import database
import pandas as pd


def bxgeneref(sd, refgene, bxgene):
    """
    主要靶向基因及药物
    :param sd:
    :param refgene:
    :param bxgene:
    :return:
    """
    sd.add_paragraph('主要靶向基因及药物', style='02级标题')
    refgeneset = set(refgene.index)
    geneset = refgeneset.intersection(set(bxgene))
    df = refgene.loc[list(geneset)]
    table = sd.add_table(rows=1, cols=6, style='reference')
    content = ['基因名称', '变异', '适应症', '药物', '检测意义', '证据等级']
    for j in range(6):
        common.add_run_self(table.cell(0, j).paragraphs[0], content[j], font_size=8, alignment='left')
    for gene, info in df.iterrows():
        rownum = len(table.rows)
        common.add_tables_self(table, 1)
        for index, sinfo in enumerate(content):
            if index == 0:
                common.add_run_self(table.cell(rownum, 0).paragraphs[0], gene, font_size=7, alignment='left')
            else:
                common.add_run_self(table.cell(rownum, index).paragraphs[0], info[sinfo], font_size=7,
                                    alignment='left')
    width = [2.12, 2.9, 3.31, 4.63, 2.28, 1.99]
    common.adjust_table_width(table, width)
    sd.add_paragraph('阅读帮助：证据等级表示该基因靶点在该适应症中对药物反应的可信度：', style='3detectresult_阅读帮助1级')
    addinfo = [
        '1：FDA认可的分子标志物，可预测本适应症中对FDA批准的药物的反应；',
        '2A：标准治疗的分子标志物，预测对该适应症中FDA批准的药物的反应；',
        '2B：在其他适应症中是标准治疗的分子标志物，预测对FDA批准的药物的反应，但在此适应症中不是标准治疗；',
        '3A：令⼈信服的临床证据⽀持⽣物标志物预测该适应症对该药物的反应；',
        '3B：令⼈信服的临床证据⽀持⽣物标志物预测其他适应症对该药物的反应；',
        '4：令⼈信服的⽣物学证据⽀持⽣物标志物预测对药物的反应；',
        'R1：标准治疗的分子标志物可预测本适应症中对FDA批准的药物的抵抗；',
        'R2：令人信服的临床证据支持分子标志物可预测对药物的抵抗。',
    ]
    for read in addinfo:
        sd.add_paragraph(read, style='0方块列表')
    sd.add_paragraph('附录内容根据本检测范围内的现有指南文件和临床研究收录。随着研究的进展，可能在未来发现新的靶标或开发新的药物。'
                     '本实验室将定期进行更新。', style='0方块列表')


def bxdocument(sd, refdocument, bxgene):
    """
    参考文献
    :param sd:
    :param refdocument:
    :param bxgene:
    :return:
    """
    sd.add_paragraph('参考文献', style='02级标题')
    refdocumentset = set(refdocument.index)
    documentset = refdocumentset.intersection(set(bxgene))
    df = refdocument.loc[list(documentset)]
    for _, info in refdocument.loc[['must']].iterrows():
        sd.add_paragraph(info['reference'], style='0reference')
    tmset = set()
    for gene, info in df.iterrows():
        tmset.add(info['reference'])
    for name in tmset:
        sd.add_paragraph(name, style='0reference')


def addgenelist(sd, detectgene):
    sd.add_paragraph('基因列表', style='02级标题')
    allnum = len(set(detectgene['geneTb'].split(',')+detectgene['geneBx'].split(',')+detectgene['geneQt'].split(',')))
    sd.add_paragraph('共检测基因{}个'.format(allnum))
    sd.add_paragraph('其中：靶向药物相关基因{}个，化疗/内分泌药物相关基因{}个，HRD-DDR相关基因{}个，肿瘤遗传易感相关基因{}个，'
                     '免疫治疗相关基因{}个，评估微卫星状态的串联重复序列240个，'
                     '用于评价肿瘤突变负荷/肿瘤新抗原的基因572个。'.format(len(detectgene['geneBx'].split(',')),
                                                       len(detectgene['geneBx'].split(',')),
                                                       len(detectgene['geneHrr'].split(',')),
                                                       len(detectgene['geneYc'].split(',')),
                                                       len(detectgene['geneMy'].split(',')),
                                                        ))
    sd.add_paragraph('其中：对{}个基因报告单碱基和小片段基因变异，对{}个基因报告基因重排，对{}个基因报告拷贝数变异，对33个基因报告基因'
                     '多态性。用于评价肿瘤突变负荷/肿瘤新抗原的基因572个。'.format(len(detectgene['geneTb'].split(',')),
                                                           len(detectgene['geneCp'].split(',')),
                                                           len(detectgene['geneKb'].split(',')),
                                                           len(detectgene['geneHl'].split(',')),))
    sd.add_paragraph('具体列表如下：')
    sd.add_paragraph()
    sd.add_paragraph('以检测目的分类')
    width = [2.25, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25, 2.25]
    descridict = dict(
        geneBx='靶向药物相关基因',
        geneMy='免疫治疗相关基因',
        geneHrr='HR-DDR相关基因',
        geneYc='肿瘤遗传易感相关基因',
        geneHl='化疗/内分泌药物相关基因',
        geneCp='报告重排（融合）的基因',
        geneKb='报告拷贝数变异的基因',
        geneHl1='报告基因多态性的基因',
        geneTb='报告单碱基变异和小片段基因变异的基因'
    )
    for tabletype in ['geneBx', 'geneMy', 'geneHrr', 'geneYc', 'geneHl',
                      'para', 'geneCp', 'geneKb', 'geneHl1', 'geneTb']:
        if tabletype == 'para':
            sd.add_paragraph('以检测范围分类')
            continue
        if tabletype == 'geneHl1':
            genebx = detectgene['geneHl'].split(',')
        else:

            genebx = detectgene[tabletype].split(',')

        bxtablenum = len(genebx)
        bxrow = bxtablenum // 8 if bxtablenum % 8 == 0 else (bxtablenum // 8) + 1
        bxtable = sd.add_table(bxrow+1, cols=8, style='reference')
        bxtable.cell(0, 0).paragraphs[0].text = '{}{}个'.format(descridict[tabletype], bxtablenum)
        for i in range(1, bxrow+1):
            tbgene = genebx[(i-1) * 8:(i-1) * 8 + 8]
            for index, gene in enumerate(tbgene):
                bxtable.cell(i, index).paragraphs[0].text = gene
            for j in range(0, 8):
                bxtable.cell(i, j).paragraphs[0].style = '0detecinfo_table3'
        common.adjust_table_width(bxtable, width)
        bxtable.cell(0, 0).merge(bxtable.cell(0, 7))
        sd.add_paragraph()


def makeref(tpl, refdocument, refgene, detectgene):
    """
    参考信息
    :param tpl:
    :param refdocument:
    :param refgene:
    :param detectgene:
    :return:
    """
    sd = tpl.new_subdoc()
    catecode = detectgene['cateCode'].split(',')
    if 'geneBx' in catecode:
        sd.add_paragraph('参考信息', style='01级标题')
        bxgene = detectgene['geneBx'].split(',')
        if 'bigPannel' in catecode:
            addgenelist(sd, detectgene)  # 大panel添加基因列表
        bxgeneref(sd, refgene, bxgene)  # 添加主要靶向基因及药物
        sd.add_page_break()
        bxdocument(sd, refdocument, bxgene)  # 添加参考文献
        return sd
    else:
        return


def refdoument():
    """
    参考文献
    :return:
    """
    path = os.path.join(rootpath, 'database', 'reference', 'refdocument.xlsx')
    df = pd.read_excel(path, sheet_name='ref', index_col=0, header=0)
    return df


def refgene():
    """
    主要靶向基因及药物
    :return:
    """
    path = os.path.join(rootpath, 'database', 'reference', 'refdocument.xlsx')
    df = pd.read_excel(path, sheet_name='refgene', index_col=0, header=0)
    return df


if __name__ == '__main__':
    refdocument = refdoument()
    refgene = refgene()
    detectinfo = database.detectinfo()
    tplpath = os.path.join(rootpath, 'database', 'template', 'blank.docx')
    path = os.path.join(rootpath, 'database', 'template', '7reference')
    content = {}
    for itmid, detectgene in detectinfo.items():
        tpl = DocxTemplate(tplpath)
        content['context'] = makeref(tpl, refdocument, refgene, detectgene)
        if content['context']:
            tpl.render(content)
            word = os.path.join(path, itmid + '.docx')
            tpl.save(word)
            print(itmid+'成功')
        else:
            print(itmid + '没内容！')
