# -*- coding: utf-8 -*-


from collections import defaultdict
import pandas as pd
import os
from tools import common
from tools import database
import requests
import json
import re


agentnamedict = database.agentname()
sourcedict = database.source()


def get_complementary(genotype):
    genotype = genotype.upper()
    genotype.replace("G", "c")
    genotype.replace("C", "g")
    genotype.replace("A", "t")
    genotype.replace("T", "a")
    genotype.upper()
    return genotype


def read_chemo_database():
    """
    化疗报告数据库
    :return:
    """
    path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'database', 'baseinfo', 'chemical.xlsx')
    df = pd.read_excel(path, sheet_name='geneChemicalArithmetic', dtype={'7_level': str})
    df['key'] = df.apply(lambda x: "-".join([x['2_gene'], x['3_rsID'], x['4_geneType']]), axis=1)
    df['value'] = df.apply(lambda x: '{}-{}-{}-{}'.format(x['1_drug'], x['5_type'], x['6_clinical'],
                                                          x['7_level']), axis=1)
    chemo_dict = {}
    df1 = pd.DataFrame(df, columns=['key', 'value'])
    for key, tmval in df1.groupby(['key']):
        chemo_dict[key] = []
        for _, val in tmval.iterrows():
             chemo_dict[key].append(val['value'])
    return chemo_dict


def read_jiema_report(jiema_report):
    """
    :param jiema_report:
    :return:
    """
    df = pd.read_excel(jiema_report, na_values=['Null'])
    # sampleIDs = df.columns.tolist()[3:]
    sampleIDs = [i for i in df.columns.tolist() if i.startswith('NGS')]
    if len(sampleIDs) == 0:
        return None
    jiema_dict = defaultdict(set)
    for index, row in df.iterrows():
        for isample in sampleIDs:
            if pd.isnull(row[isample]):
                genotype = '纯合缺陷突变'
            elif row[isample].lower() == 'present':
                genotype = '野生型'
            elif len(row[isample]) == 2 and '/' not in row[isample]:
                genotype = row[isample][0] + '/' + row[isample][1]
                genotype2 = row[isample][1] + '/' + row[isample][0]
                genotype3 = get_complementary(genotype)
                genotype4 = get_complementary(genotype2)
                jiema_dict[isample].add((row['基因'], row['位点RS号'], genotype2))
                jiema_dict[isample].add((row['基因'], row['位点RS号'], genotype3))
                jiema_dict[isample].add((row['基因'], row['位点RS号'], genotype4))
            else:
                genotype = row[isample].replace("DEL", "del")
            jiema_dict[isample].add((row['基因'], row['位点RS号'], genotype))
    return jiema_dict


def cms(sample):
    """

    :param sample:
    :return:
    """
    useragent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6)\
     AppleWebKit/537.36 (KHTML, like Gecko)\
      Chrome/74.0.3729.169 Safari/537.36'
    header = {'User-Agent': useragent}
    posturl = 'http://cms.topgen.com.cn/user/login/auth'
    postdata = {"userName": "xiecaixia", "password": "xcx50800383"}
    # postdata = {"userName": "kaitian", "password": "top123"}
    response = requests.post(posturl, params=postdata, headers=header)
    text = response.text
    if response.status_code == 200:
        token = json.loads(text)["data"]["accessToken"]
        # print(token)
        payload = {"search[sampleSn]": sample, "accessToken": token}
        r = requests.get("http://cms.topgen.com.cn/sample/sample/search?report=true", params=payload)
        # print(r)
    else:
        raise Exception("网站不能在此账号下访问!", response.status_code)
    info_list = json.loads(r.text)['data'][0]
    custom = database.custom()
    info_list['sampleSn'] = sample
    info_list['name'] = info_list['name'].strip()
    if info_list['gender'] == '男':
        info_list['gender1'] = info_list['name'] + '先生'
    else:
        info_list['gender1'] = info_list['name'] + '女士'
    info_list['makeTime1'] = common.datatran(info_list['makeTime'])
    info_list['receiveTime1'] = common.datatran(info_list['receiveTime'])
    info_list['decTime'] = common.datatran(info_list['receiveTime'], flag=1, add=1)
    info_list['itemId'] = info_list['itemName'].split('.')[0]
    if info_list['agentName'] in custom:
        agent = info_list['agentName']
        info_list['iscustom'] = False
    else:
        agent = '鼎晶生物'
        info_list['iscustom'] = True
    info_list['custom'] = custom[agent]
    info_list['custom']['agentName'] = agent
    source = info_list['source'].split(' ')[0]
    if source in sourcedict:
        info_list['psource'] = sourcedict[source]['分类']
    else:
        info_list['psource'] = ' '
    if info_list['agentName'] in agentnamedict:
        info_list['agentName'] = agentnamedict[info_list['agentName']]['hospitalName']
    else:
        info_list['agentName'] = ' '
    for tm in info_list:
        if tm == 'decTime' or 'agentName':
            continue
        info_list[tm] = str(info_list[tm]).strip('.0')
        if info_list[tm] == 'NA':
            info_list[tm] = ' '
    return info_list


def run(path):
    """

    :param path:
    :return:
    """
    chemo_dict = read_chemo_database()
    chemicaldatabase = database.chemical()
    geneChemicalArithmetic = chemicaldatabase['arithmetic']
    chemicalDrugEffect = chemicaldatabase['drugeffect']
    jiema_dict = read_jiema_report(path)
    result = common.tree()
    for sample in jiema_dict:
        drug = common.tree()
        tmedical1, tmedical2, tmedical3 = [], [], []
        smedical1, smedical2, smedical3 = [], [], []
        for variant in jiema_dict[sample]:
            ikey1 = "-".join([variant[0], variant[1], variant[2]])
            if ikey1 not in chemo_dict:
                continue
            tmivalue = chemo_dict[ikey1]
            for value in tmivalue:
                ivaluelist = value.split('-')
                medicine = ivaluelist[0]
                if medicine not in drug:
                    drug[medicine]['content'] = []
                drug[medicine]['content'].append([variant[0], variant[1], variant[2], ivaluelist[2], ivaluelist[3]])
        for medicineName in drug:
            therapeuticEffect = []
            sideEffect = []
            for tmlist in drug[medicineName]['content']:
                ikey = "_".join([medicineName, tmlist[0], tmlist[1], tmlist[2]])
                stherapeuticEffect = geneChemicalArithmetic[ikey]['8_therapeuticEffect']
                ssideEffect = geneChemicalArithmetic[ikey]['9_sideEffect']
                if stherapeuticEffect != '/' and stherapeuticEffect != '':
                    therapeuticEffect.append(stherapeuticEffect)
                if ssideEffect != '/' and ssideEffect != '':
                    sideEffect.append(ssideEffect)
            if therapeuticEffect:
                therapeuticlevel = sum([int(i) for i in therapeuticEffect if i != '/' and i != ''])
                try:
                    tkey = '{}-疗效={}'.format(medicineName, therapeuticlevel)
                    therapeuticdescr = chemicalDrugEffect[tkey]['3_effect']
                except:
                    therapeuticdescr = ''
            else:
                therapeuticdescr = ''
            if sideEffect:
                sideEffectlevel = sum([int(j) for j in sideEffect if j != '/' and j != ''])
                try:
                    skey = '{}-毒副={}'.format(medicineName, sideEffectlevel)
                    sideEffectdescr = chemicalDrugEffect[skey]['3_effect']
                except:
                    sideEffectdescr = ''
            else:
                sideEffectdescr = ''
            if therapeuticdescr.find('疗效较好') > -1:
                tmedical1.append(medicineName)
            elif therapeuticdescr.find('疗效中等') > -1:
                tmedical2.append(medicineName)
            elif therapeuticdescr.find('疗效较差') > -1:
                tmedical3.append(medicineName)
            else:
                pass
            if sideEffectdescr.find('毒副较高') > -1:
                smedical3.append(medicineName)
            elif sideEffectdescr.find('毒副中等') > -1:
                smedical2.append(medicineName)
            elif sideEffectdescr.find('毒副较低') > -1:
                smedical1.append(medicineName)
            else:
                pass
            descr = filter(lambda x: x is not '', [therapeuticdescr, sideEffectdescr])
            medical = ','.join(descr)
            drug[medicineName]['medical'] = medical
        result[sample]['chemical']['drug'] = drug
        for n in range(1, 4):
            curtname = 'tmedical' + str(n)
            cursname = 'smedical' + str(n)
            tdescribe = '，'.join(eval(curtname))
            sdescribe = '，'.join(eval(cursname))
            if not tdescribe:
                tdescribe = '无'
            if not sdescribe:
                sdescribe = '无'
            result[sample]['chemical'][curtname] = tdescribe
            result[sample]['chemical'][cursname] = sdescribe
        sampleinfo = cms(sample)
        result[sample]['cms'] = sampleinfo
        result[sample]['sampleSn'] =sample
        result[sample]['itemid'] = '0201025'
        result[sample]['treatresult'] = '纯化疗项目'
        result[sample]['treatid'] = '纯化疗项目'
        contents = [
            'target', 'fusion', 'heredity', 'immune',
            'msi', 'tmb', 'tnb', 'hla', 'cna',
            'clinical', 'hrr', 'qc', 'projectAbbreviation', 'varnum', 'drug', 'nccngene', 'hrd'
        ]
        for t in contents:
            result[sample][t] = {}
    resultjson = json.dumps(result, sort_keys=True, indent=4, ensure_ascii=False)
    return resultjson
