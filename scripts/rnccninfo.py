# -*- coding: utf-8 -*-
"""
检测信息更新
"""

import os
import sys
from docxtpl import DocxTemplate
import pandas as pd

rootpath = os.path.dirname((os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
sys.path.append(rootpath)
from tools import common

basepath = os.path.join(rootpath, 'database', )
tplpath = os.path.join(basepath, 'template', 'blank.docx')
newpara = common.newpara


def nccn():
    """
    生成适应症癌种
    :return:
    """
    nccnpath = os.path.join(basepath, 'baseinfo', 'nccn.xlsx')
    tplpath = os.path.join(basepath, 'template', 'blank.docx')
    nccnxlsx = pd.read_excel(nccnpath, None)
    for treadid, df in nccnxlsx.items():
        df = df.fillna('NA')
        tpl = DocxTemplate(tplpath)
        sd = tpl.new_subdoc()
        p = sd.add_paragraph('', style='02级标题')
        sd.add_paragraph()
        table = sd.add_table(rows=1, cols=3, style='1summary_common')
        content = ['检测基因', '检测意义', '检测结果']
        table.cell(0, 0).text = content[0]
        table.cell(0, 1).text = content[1]
        table.cell(0, 2).text = content[2]
        nccndescribe = ''
        for _, rows in df.iterrows():
            gene = rows['检测基因']
            if gene == 'info':
                nccndescribe = rows['检测范围']
            elif gene == '微卫星状态':
                continue
            elif gene == 'treatresult':
                treatresult = rows['检测范围']
                p.add_run('{}NCCN指南相关基因'.format(treatresult))
            else:
                rownum = len(table.rows)
                common.add_tables_self(table, 1)
                common.add_run_self(table.cell(rownum, 0).paragraphs[0], gene, font_size=8, alignment='left')
                if rows['检测意义'] == 'NA':
                    table.cell(rownum, 1).merge(table.cell(rownum - 1, 1))
                else:
                    common.add_run_self(table.cell(rownum, 1).paragraphs[0],rows['检测意义'],
                                        font_size=8, alignment='other')
                rlist = ['{{%p if nccn.{}.f %}}'.format(gene),
                         '{{{{ nccn.{0}.r }}}}'.format(gene),
                         '{{%p else %}}',
                         '{{{{ nccn.{0}.r }}}}'.format(gene),
                         '{{%p endif %}}']
                for i in range(5):
                    if i > 0:
                        table.cell(rownum, 2).add_paragraph('')
                    if i == 1:
                        common.add_run_self(table.cell(rownum, 2).paragraphs[i], rlist[i], font_size=8,
                                            font_color=(255, 99, 28), alignment='left')
                    else:
                        common.add_run_self(table.cell(rownum, 2).paragraphs[i], rlist[i], font_size=8,
                                            alignment='left')
            width = [2.74, 11, 4.0]
            common.adjust_table_width(table, width)
        sd.add_paragraph(nccndescribe, style='0星号段落')
        word = os.path.join(basepath, 'template', '4medicaltip', 'nccn', f'{treadid}.docx')
        context = {'context': sd}
        tpl.render(context, autoescape=True)
        tpl.save(word)


if __name__ == '__main__':
    nccn()
    print(f'nccn信息更新成功')
