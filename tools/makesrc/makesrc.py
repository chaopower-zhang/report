# -*- coding: utf-8 -*-
import hashlib
from string import Template
from . import estamp
from tools.common import rootpath
import os
import re

newline = '</w:t><w:br/><w:t xml:space="preserve">'


class NewTemplate(Template):
    delimiter = "%"


def read_template(template):
    with open(template, "r", encoding='utf-8') as fh:
        return NewTemplate(fh.read())


def generate_webpage(sampleinfo, detectgene):
    htmlpath = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'template', 'tpl.html')
    derpath = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'template', '.privkey.der')
    catecode = detectgene['cateCode'].split(',')
    if 'geneBx' in catecode:
        targets = sampleinfo['target']
        target_html = make_target_string(targets)
    else:
        target_html = ''
    if 'geneCp' in catecode:
        fusions = sampleinfo['fusion']
        fusion_html = make_fusion_string(fusions)
    else:
        fusion_html = ''
    if 'geneKb' in catecode:
        cna = sampleinfo['cna']
        cna_html = make_cna_string(cna)
    else:
        cna_html = ''
    if 'geneYc' in catecode:
        hered = sampleinfo['sum_yc']
        hered_html = make_heredity_string(hered)
    else:
        hered_html = ''
    if 'geneHl' in catecode:
        chem = sampleinfo['chemical']
        chemo_html = make_chemo_string(chem)
    else:
        chemo_html = ''
    if 'thys' in catecode:
        targets = sampleinfo['target']
        target_html = make_target_string(targets)
        fusions = sampleinfo['fusion']
        fusion_html = make_fusion_string(fusions)
        cna = sampleinfo['cna']
        cna_html = make_cna_string(cna)
        hered = sampleinfo['heredity']
        hered_html = make_heredity_string(hered)

    immune = sampleinfo['immune']
    immune_html = ''
    name = sampleinfo['c']['name']
    sample = sampleinfo['c']['sampleSn']
    substitute_dict = dict([("name", "*" + name[-1]),
                            ("sampleID", sample),
                            ('target', target_html),
                            ('fusion', fusion_html),
                            ('cna', cna_html),
                            ('hered',hered_html),
                            ('immune',immune_html),
                            ('chemo', chemo_html)])
    origin_tpl = read_template(htmlpath)
    replace_html = origin_tpl.safe_substitute(substitute_dict)
    url_name = hashlib.md5(sample.encode()).hexdigest() + '.html'
    path = os.path.join(rootpath, 'result', 'html', url_name)
    with open(path, 'w', encoding='utf-8') as fw:
        fw.write(replace_html)
    src = estamp.create_digital_signature(path, derpath)
    return src


def make_target_string(parsed_result):
    target_strings = []
    for peralter in parsed_result:
        gene = peralter['geneName']
        if gene.find('野生型') > -1:
            continue
        target_strings.append("{}基因，外显子{}，{}突变，突变频率为{}".format(
            peralter['geneName'],
            peralter['wxz'],
            peralter['alteration'],
            peralter['freq']))
    if not target_strings:
        target_strings = ["未检测到可指导靶向药物的相关基因突变"]
    # generate html string for target section
    target_html = ''
    for index, each_string in enumerate(target_strings, start=1):
        target_html += '<div class="item">\n\t\t\t\t\t{}、{}\n\t\t\t\t</div>\n\t\t\t\t'.format(index, each_string)
    target_html = """
    <div>
        <h3>特异性突变列表：</h3>
        <div class="ui bulleted list">
            {}
        </div>
    </div>
    <br />
    """.format(target_html)
    return target_html


def make_fusion_string(parsed_result):
    """
    """
    fusion_strings = []
    if not parsed_result :
        fusion_strings = ["未检测到靶向药物相关的融合变异"]
    else:
        for peralter in parsed_result:
            if peralter['freq'] == '.' or peralter['freq'] == '-':
                freq = '-'
            else:
                freq = peralter['freq']
            fusion_strings.append("{0}，变异频率为{1}".format(peralter['geneName'], freq))

    # generate html string for target section
    fusion_html = ''
    for index, each_string in enumerate(fusion_strings, start=1):
        fusion_html += '<div class="item">\n\t\t\t\t\t{}、{}\n\t\t\t\t</div>\n\t\t\t\t'.format(index, each_string)
    fusion_html = """
    <div>
        <h3>基因融合列表：</h3>
        <div class="ui bulleted list">
            {0}
        </div>
    </div>
    <br />
    """.format(fusion_html)
    return fusion_html


def make_cna_string(parsed_result):
    """
    """
    cna_strings = []
    if not parsed_result:
        cna_strings = ["未检测到靶向药物相关的拷贝数变异"]
    else:
        for peralter in parsed_result:
            if peralter['geneName'] != '.' or peralter['cnaRatio'] != '.':
                cna_strings.append("{0}，变异倍数为{1}".format(peralter['geneName'], peralter['cnaRatio']))
        if not cna_strings:
            cna_strings = ["未检测到靶向药物相关的拷贝数变异"]

    # generate html string for target section
    cna_html = ''
    for index, each_string in enumerate(cna_strings, start=1):
        cna_html += '<div class="item">\n\t\t\t\t\t{}、{}\n\t\t\t\t</div>\n\t\t\t\t'.format(index, each_string)
    cna_html = """
    <div>
        <h3>基因拷贝数变异列表：</h3>
        <div class="ui bulleted list">
            {0}
        </div>
    </div>
    <br />
    """.format(cna_html)
    return cna_html


def make_immune_string(parsed_result, bed_type):
    """

    :param bed_type:
    :param parsed_result:
    :return:
    """
    if not isinstance(parsed_result['msi'], float):
        return ''

    if bed_type == "DE":
        if parsed_result['msi'] >= 0.3:
            msi_level = 'MSI-H'
        elif 0.1 <= parsed_result['msi'] < 0.3:
            msi_level = 'MSI-L'
        else:
            msi_level = 'MSS'
        immune_html = '<div class="ui items">\n\t\t\t<div class="item">\n\t\t\t\t<div class="content">\n\t\t\t\t\t<a class="header">免疫检查点抑制剂相关分子标志物：</a>\n\t\t\t\t\t<div class="description">\n\t\t\t\t\t\t<p>微卫星不稳定性（MSI）：MSI SCORE ' + str(parsed_result['msi']) + '，分析结果为 ' + msi_level + ';</p>\n\t\t\t\t\t</div>\n\t\t\t\t</div>\n\t\t\t</div>\n\t\t</div>'
    elif bed_type == "DU":
        if parsed_result['msi'] >= 0.25:
            msi_level = 'MSI-H'
        elif 0.06 <= parsed_result['msi'] < 0.25:
            msi_level = 'MSI-L'
        else:
            msi_level = 'MSS'
        if parsed_result['tmb'] <= 6:
            tmb_level = '低'
        elif 6 < parsed_result['tmb'] <= 20:
            tmb_level = '中'
        else:
            tmb_level = '高'
        immune_html = '<div class="ui items">\n\t\t\t<div class="item">\n\t\t\t\t<div class="content">\n\t\t\t\t\t<a class="header">免疫检查点抑制剂相关分子标志物：</a>\n\t\t\t\t\t<div class="description">\n\t\t\t\t\t\t<p>微卫星不稳定性（MSI）：MSI SCORE ' + str(parsed_result['msi']) + '，分析结果为 ' + msi_level + ';</p>\n\t\t\t\t\t\t<p>肿瘤突变负荷（TMB）：' + str(parsed_result['tmb']) + 'muts/Mb，评级为 ' + tmb_level + '；</p>\n\t\t\t\t\t</div>\n\t\t\t\t</div>\n\t\t\t</div>\n\t\t</div>'
    else:
        immune_html = ""
    return immune_html


def make_heredity_string(parsed_result):
    """

    :param parsed_result:
    :param panel_project:
    :return:
    """

    heredity_strings = parsed_result['describe'].split(newline)
    heredity_html = ''
    for index, each_string in enumerate(heredity_strings, start=1):
        heredity_html += '\t\t\t\t<div class="item">\n\t\t\t\t\t{}、{}\n\t\t\t\t</div>\n'.format(index, each_string)
    heredity_html = ('<div>\n\t\t\t<h3>遗传性肿瘤相关基因：</h3>\n\t\t\t<div class="ui bulleted list">\n' +
                    heredity_html + '\t\t\t</div>\n\t\t</div><br />')
    return heredity_html


def make_chemo_string(parsed_result):
    """
    :param parsed_result:
    :param panel_project:
    :return:
    """
    chemo_html = """<div class='ui basic segment'>
                        <h3>药物代谢酶SNP</h3>
                        <table class="ui celled table">
                            <thead>
                                <tr>
                                    <th>基因名</th>
                                    <th>位点</th>
                                    <th>基因型</th>
                                    <th>基因名</th>
                                    <th>位点</th>
                                    <th>基因型</th>
                                    <th>基因名</th>
                                    <th>位点</th>
                                    <th>基因型</th>
                                </tr>
                            </thead>
                            <tbody>

    """
    try:
        chemo_html += '<tr>'
        i = 1
        for content in parsed_result:
            for genetype in content['medicalgene']:
                chemo_html += """<td data-label="基因名">{}</td><td data-label="位点">{}</td><td data-label="基因型">{}</td>""".format(genetype['genename'], genetype['rsid'],genetype['type'])
                if i % 3 == 0:
                    chemo_html += '</tr>'
                    chemo_html += '<tr>'
                i += 1
        for j in range(4 - (i % 3)):
            chemo_html += """<td data-label="基因名"></td><td data-label="位点"></td><td data-label="基因型"></td>"""
    except:
        chemo_html += '<tr>'
    chemo_html += '</tr></tbody></table>'
    return chemo_html
