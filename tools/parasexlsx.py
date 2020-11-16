# -*- coding: utf-8 -*-

from tools import readxlsx
from tools.common import tree
from tools.common import datatran
from tools import database
import json
import re
import logging
from collections import defaultdict
from tools.common import drugcmp
import functools
import glob
import os
import pandas as pd
import numpy as np
from tools.common import newline
from tools.common import rootpath
from tools.common import qrcodeself
from tools import makesrc

# 数据库相关
chemicaldatabase = database.chemical()  # 化疗
tbtranslate = database.tbtranslate()  # 突变类型转换
treattran = database.treattran()  # 诊断ID与癌种
specialagent = database.specialagent()  # 特殊送检单位不出本癌种信息
nccn = database.nnccn()  # nccn相关
detectinfo = database.detectinfo()  # 检测信息表
cmsneedlist = database.cmsneedlist()  # 报告需要的填充的cms系统信息
agentnamedict = database.agentname()  # 机构名称
sourcedict = database.source()  # 样品类型转换
custom = database.custom()  # 定制版报告


class NpEncoder(json.JSONEncoder):
    """
    json 强制
    """

    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        else:
            return super(NpEncoder, self).default(obj)


class Parase:
    """
    单个样品信息处理
    """

    def __init__(self, sampledata):
        """
        :param sampledata: 包含样品名的字典信息。
        """
        self.result = tree()  # 报告结果
        self.sampledata = sampledata
        self.nccnunknow = set()
        self.nccngene = defaultdict(set)
        self.hepagene = defaultdict(set)
        self.lowfreq = list()
        self.type1, self.type2, self.type3 = set(), set(), set()
        self.tarnum, self.funum, self.cnanum, = 0, 0, 0
        self.unknowtarnum = set()
        self.logger = logging.getLogger('main.sub')
        self.tbgene = set()
        self.hgene = set()
        self.hrdgene = tree()
        self.pic = list()
        self.trueresult = defaultdict(set)
        self.cms()  # 依据cms信息获取全局信息

    def run(self):
        self.dectgene()  # 判断产品ID是否在panellist中
        self.baseinfo()  # 报告抬头等基本信息
        self.target()
        self.fusion()
        self.cna()
        self.clinical()
        self.chemical()
        self.immune()
        self.msi()
        self.addtnb()
        self.hla()
        self.heredity()
        self.hrd()
        self.pdl1()
        self.testnccn()
        self.hepa()
        self.sign()
        self.specialtreatresult()
        self.specialagent()
        self.falseresult()
        self.collect()
        self.addpict()
        self.sqrcode()
        return self.result

    def cms(self):
        """
        cms信息处理
        :return:
        """
        data = self.sampledata['cms'][0]  # 数据结构为列表

        data['name'] = str(data['name']).strip()
        if data['gender'] == '男':
            data['gender1'] = data['name'] + '先生'
        else:
            data['gender1'] = data['name'] + '女士'
        data['makeTime1'] = datatran(data['makeTime'])
        data['receiveTime1'] = datatran(data['receiveTime'])
        data['decTime'] = datatran(data['receiveTime'], flag=1, add=1)
        data['itemId'] = data['itemName'].split('.')[0]
        if data['agentName'] in custom:
            agent = data['agentName']
            data['iscustom'] = False
        else:
            agent = '鼎晶生物'
            data['iscustom'] = True
        data['custom'] = custom[agent]
        data['custom']['agentName'] = agent
        source = data['source'].split(' ')[0]
        if source in sourcedict:
            data['psource'] = sourcedict[source]['分类']
        else:
            data['psource'] = ' '
        if data['psource'] == '血液/口腔粘膜细胞/口腔拭子':
            data['somatic'] = False
        elif data['psource'] == '组织或胸腹水细胞':
            data['somatic'] = True
        else:
            data['somatic'] = False
        if data['agentName'] in agentnamedict:
            data['agentName'] = agentnamedict[data['agentName']]['hospitalName']
        else:
            data['agentName'] = ' '
        if data['hospitalName'] == '.':
            data['hospitalName'] = ' '
        for tm in data:
            if tm == 'decTime' or 'agentName':
                continue
            data[tm] = str(data[tm]).strip('.0')
            if data[tm] == 'NA' or data[tm] == '.':
                data[tm] = ' '
        # 报告样本信息表中内容
        for need in cmsneedlist:
            self.result['c'][need] = data[need]
        self.sample = data['sampleSn']
        self.agentname = data['agentName']
        self.unitagent_name = data['unitAgentName']
        self.treatid = data['treatID']
        self.itemName = data['itemName']
        self.itemid = data['itemId']
        self.source = data['source']
        self.detecttime = data['decTime']
        self.custom = data['custom']
        self.somatic = data['somatic']
        if self.agentname in specialagent:
            self.valid = True
        else:
            self.valid = False

    def dectgene(self):
        try:
            self.detectgene = detectinfo[self.itemid]
            self.catecode = self.detectgene['cateCode'].split(',')
        except Exception as e:
            self.logger.warning('{} 没有{} 的产品报告模版'.format(e, self.itemid))
            print('{}结束, 失败！没有{} 的产品报告模版'.format(self.sample, self.itemid))
            raise UserWarning('{}结束, 失败！没有{} 的产品报告模版'.format(self.sample, self.itemid))

    def baseinfo(self):
        """
        报告模版基础内容
        :return:
        """
        self.result['c']['reportname'] = self.detectgene['reportName']
        self.result['c']['reporttype'] = self.detectgene['reportType']
        self.result['c']['reporttitle'] = self.detectgene['reportItemName']
        self.result['c']['time'] = datatran()
        self.result['c']['catecode'] = self.catecode

    def collect(self):
        type1 = self.type1.difference(self.type3)
        type2 = self.type2.difference(self.type3)
        type2 = type2.difference(type1)
        type3 = self.type3
        for i in range(1, 4):
            tmp = eval('type' + str(i))
            if '' in tmp:
                tmp.remove('')
            if '.' in tmp:
                tmp.remove('.')
        self.result['sum_drug']['maynum'] = len(type1) + len(type2)
        self.result['sum_drug']['resistantnum'] = len(type3)
        self.result['sum_drug']['allnum'] = len(type1) + len(type2) + len(type3)
        for tmtype in [type1, type2, type3]:
            if not tmtype:
                tmtype.add('无')
        self.result['sum_drug']['type1'] = '，'.join(sorted(type1, key=functools.cmp_to_key(drugcmp)))
        self.result['sum_drug']['type2'] = '，'.join(sorted(type2, key=functools.cmp_to_key(drugcmp)))
        self.result['sum_drug']['type3'] = '，'.join(sorted(type3, key=functools.cmp_to_key(drugcmp)))
        self.result['sum_varnum']['tarnum'] = self.tarnum
        self.result['sum_varnum']['funum'] = self.funum
        self.result['sum_varnum']['cnanum'] = self.cnanum
        self.result['sum_varnum']['unknowtarnum'] = len(self.unknowtarnum)
        if 'addUnkowgene' in self.catecode:
            self.result['sum_varnum']['addUnkowgene'] = True
        else:
            self.result['sum_varnum']['addUnkowgene'] = False
        if self.sampledata['qc']:
            self.result['sum_qc'] = self.sampledata['qc'][0]
        else:
            self.result['sum_qc'] = {}
        self.result['pic'] = self.pic

    def target(self):
        """
        靶向信息整理
        :return:
        """
        data = pd.DataFrame(self.sampledata['target'])
        self.result['target'] = list()
        if data.empty:
            return
        data.fillna('.', inplace=True)
        for info, data_1 in data.groupby(['geneName', 'alteration']):
            genename, alteration = info
            self.hepagene['target'].add(genename)
            alterdrugtype1, alterdrugtype2, alterdrugtype3 = set(), set(), set()
            if genename == '.':
                continue
            self.trueresult['target'].add(genename)
            self.tarnum += 1
            alterinfo = data_1.iloc[0].to_dict()
            try:
                freq = '{:.2%}'.format(float(alterinfo['freq']))
            except ValueError:
                freq = alterinfo['freq']
                self.logger.warning(f'靶向表{self.sample}样品{genename}的{alteration}变异的频率{freq}不能转换成小数')
                raise Exception(f'靶向表{self.sample}样品{genename}的{alteration}变异的频率{freq}不能转换成小数')

            if float(alterinfo['freq']) < 0.01:
                lowfreq = True
            else:
                lowfreq = False
            genealter = '{} {}{}{}'.format(genename, alterinfo['wxz'], newline, alterinfo['alteration'])
            cancertypelist = data_1['cancerType'].to_list()
            validcancer = list()
            if self.valid:
                for tcancertype in cancertypelist:
                    if tcancertype in treattran:
                        tmptreatid = treattran[tcancertype]['treatid']
                    else:
                        tmptreatid = 'TL'
                    if tmptreatid in self.treatid:
                        validcancer.append(tcancertype)
            if not validcancer:
                validcancer = cancertypelist
            alterinfo['cancerdrug'] = tree()
            for cancertype, data_2 in data_1.groupby(['cancerType']):
                if cancertype not in validcancer:
                    continue
                if cancertype in treattran:
                    tmptreatid = treattran[cancertype]['treatid']
                else:
                    tmptreatid = 'TL'
                type1, type2, type3 = set(), set(), set()
                for _, rows in data_2.iterrows():
                    druglevel = rows['drugLevel']
                    drugs = rows['drugs']
                    if druglevel != '.' and druglevel != '本次检测到的变异对临床用药的指导意义暂不明确':
                        self.tbgene.add(genename)
                        self.nccngene['target'].add(genename)
                        if druglevel.startswith('1') or druglevel.startswith('2'):
                            type1.update(set(drugs.split('，')))
                            if tmptreatid in self.treatid:
                                self.type1.update(set(drugs.split('，')))
                            else:
                                self.type2.update(set(drugs.split('，')))
                        elif druglevel.startswith('R'):
                            self.type3.update(set(drugs.split('，')))
                            type3.update(set(drugs.split('，')))
                        else:
                            self.type2.update(set(drugs.split('，')))
                            type2.update(set(drugs.split('，')))
                    else:
                        if 'geneBx' in alterinfo['geneType']:
                            self.unknowtarnum.add(alteration)
                        self.nccnunknow.add(genename)

                if any((type1, type2, type3)):
                    type1 = type1.difference(type3)
                    type2 = type2.difference(type3)
                    type2 = type2.difference(type1)
                    alterdrugtype1.update(type1)
                    alterdrugtype2.update(type2)
                    alterdrugtype3.update(type3)
                    alterinfo['cancerdrug'][cancertype]['type1'] = '，'.join(
                        sorted(type1, key=functools.cmp_to_key(drugcmp)))
                    alterinfo['cancerdrug'][cancertype]['type2'] = '，'.join(
                        sorted(type2, key=functools.cmp_to_key(drugcmp)))
                    alterinfo['cancerdrug'][cancertype]['type3'] = '，'.join(
                        sorted(type3, key=functools.cmp_to_key(drugcmp)))
                else:
                    self.tarnum -= 1
                    self.trueresult['target'].discard(genename)
            mutiontype = alterinfo['exonicfuncRefgene']
            if mutiontype in tbtranslate:
                mutiontype = tbtranslate[mutiontype]['translate']
            alterinfo['freq'] = freq
            alterinfo['mutiontype'] = mutiontype
            alterinfo['genealter'] = genealter
            alterinfo['lowfreq'] = lowfreq
            alterinfo['drugtype1'] = '，'.join(sorted(alterdrugtype1, key=functools.cmp_to_key(drugcmp)))
            alterinfo['drugtype2'] = '，'.join(sorted(alterdrugtype2, key=functools.cmp_to_key(drugcmp)))
            alterinfo['drugtype3'] = '，'.join(sorted(alterdrugtype3, key=functools.cmp_to_key(drugcmp)))
            self.result['target'].append(alterinfo)
            if 'geneHRD' in alterinfo['geneType'] and genename not in self.hgene:
                if alterinfo['level'].startswith('4') or alterinfo['level'].startswith('5'):
                    self.hgene.add(genename)
                    self.hrdgene[genename]['freq'] = freq
                    self.hrdgene[genename]['level'] = alterinfo['level']

    def fusion(self):
        """
        :return:
        """
        data = pd.DataFrame(self.sampledata['fusion'])
        self.result['fusion'] = list()
        if data.empty:
            return
        data.fillna('.', inplace=True)
        for info, data_1 in data.groupby(['geneName', 'alteration']):
            genename, alteration = info
            self.hepagene['fusion'].add(genename)
            alterdrugtype1, alterdrugtype2, alterdrugtype3 = set(), set(), set()
            if genename == '.':
                continue
            self.trueresult['fusion'].add(genename)
            self.funum += 1
            alterinfo = data_1.iloc[0].to_dict()
            try:
                freq = '{:.2%}'.format(float(alterinfo['freq']))
            except ValueError:
                freq = alterinfo['freq']
                self.logger.warning(f'融合表{self.sample}样品{genename}的{alteration}变异的频率{freq}不能转换成小数')
                raise Exception(f'融合表{self.sample}样品{genename}的{alteration}变异的频率{freq}不能转换成小数')
            match = re.match(r'(.*) exon(\d+) *-(.*) exon(\d+)', alteration)
            genealter = match.group(1) + '-' + match.group(3) + '融合'

            cancertypelist = data_1['cancerType'].to_list()
            validcancer = list()
            if self.valid:
                for tcancertype in cancertypelist:
                    if tcancertype in treattran:
                        tmptreatid = treattran[tcancertype]['treatid']
                    else:
                        tmptreatid = 'TL'
                    if tmptreatid in self.treatid:
                        validcancer.append(tcancertype)
            if not validcancer:
                validcancer = cancertypelist
            alterinfo['cancerdrug'] = tree()
            for cancertype, data_2 in data_1.groupby(['cancerType']):
                if cancertype not in validcancer:
                    continue
                if cancertype in treattran:
                    tmptreatid = treattran[cancertype]['treatid']
                else:
                    tmptreatid = 'TL'
                type1, type2, type3 = set(), set(), set()
                for _, rows in data_2.iterrows():
                    druglevel = rows['drugLevel']
                    drugs = rows['drugs']
                    if druglevel != '.' and druglevel != '本次检测到的变异对临床用药的指导意义暂不明确':
                        self.tbgene.add(genename)
                        self.nccngene['fusion'].add(genename)
                        if druglevel.startswith('1') or druglevel.startswith('2'):
                            type1.update(set(drugs.split('，')))
                            if tmptreatid in self.treatid:
                                self.type1.update(set(drugs.split('，')))
                            else:
                                self.type2.update(set(drugs.split('，')))
                        elif druglevel.startswith('R'):
                            self.type3.update(set(drugs.split('，')))
                            type3.update(set(drugs.split('，')))
                        else:
                            self.type2.update(set(drugs.split('，')))
                            type2.update(set(drugs.split('，')))
                    else:
                        self.nccnunknow.add(genename)
                if any((type1, type2, type3)):
                    type1 = type1.difference(type3)
                    type2 = type2.difference(type3)
                    type2 = type2.difference(type1)
                    alterdrugtype1.update(type1)
                    alterdrugtype2.update(type2)
                    alterdrugtype3.update(type3)
                    alterinfo['cancerdrug'][cancertype]['type1'] = '，'.join(
                        sorted(type1, key=functools.cmp_to_key(drugcmp)))
                    alterinfo['cancerdrug'][cancertype]['type2'] = '，'.join(
                        sorted(type2, key=functools.cmp_to_key(drugcmp)))
                    alterinfo['cancerdrug'][cancertype]['type3'] = '，'.join(
                        sorted(type3, key=functools.cmp_to_key(drugcmp)))
                else:
                    self.funum -= 1
            alterinfo['mutiontype'] = genealter
            alterinfo['genealter'] = genealter
            alterinfo['freq'] = freq
            alterinfo['drugtype1'] = '，'.join(sorted(alterdrugtype1, key=functools.cmp_to_key(drugcmp)))
            alterinfo['drugtype2'] = '，'.join(sorted(alterdrugtype2, key=functools.cmp_to_key(drugcmp)))
            alterinfo['drugtype3'] = '，'.join(sorted(alterdrugtype3, key=functools.cmp_to_key(drugcmp)))
            self.result['fusion'].append(alterinfo)

    def cna(self):
        """
        :return:
        """
        data = pd.DataFrame(self.sampledata['cna'])
        self.result['cna'] = list()
        if data.empty:
            return
        data.fillna('.', inplace=True)
        testcna = set()
        for info, data_1 in data.groupby(['geneName', 'alteration']):
            genename, alteration = info
            self.hepagene['cna'].add(genename)
            alterdrugtype1, alterdrugtype2, alterdrugtype3 = set(), set(), set()
            if genename == '.':
                continue
            self.trueresult['cna'].add(genename)
            testcna.add(genename)
            self.cnanum += 1
            alterinfo = data_1.iloc[0].to_dict()
            genealter = genename + alteration
            freq = '-'
            cancertypelist = data_1['cancerType'].to_list()

            validcancer = list()
            if self.valid:
                for tcancertype in cancertypelist:
                    if tcancertype in treattran:
                        tmptreatid = treattran[tcancertype]['treatid']
                    else:
                        tmptreatid = 'TL'
                    if tmptreatid in self.treatid:
                        validcancer.append(tcancertype)
            if not validcancer:
                validcancer = cancertypelist
            alterinfo['cancerdrug'] = tree()
            for cancertype, data_2 in data_1.groupby(['cancerType']):
                if cancertype not in validcancer:
                    continue
                if cancertype in treattran:
                    tmptreatid = treattran[cancertype]['treatid']
                else:
                    tmptreatid = 'TL'
                type1, type2, type3 = set(), set(), set()

                for _, rows in data_2.iterrows():
                    druglevel = rows['drugLevel']
                    drugs = rows['drugs']
                    if druglevel != '.' and druglevel != '本次检测到的变异对临床用药的指导意义暂不明确':
                        self.tbgene.add(genename)
                        self.nccngene['cna'].add(genename)
                        if druglevel.startswith('1') or druglevel.startswith('2'):
                            type1.update(set(drugs.split('，')))
                            if tmptreatid in self.treatid:
                                self.type1.update(set(drugs.split('，')))
                            else:
                                self.type2.update(set(drugs.split('，')))
                        elif druglevel.startswith('R'):
                            self.type3.update(set(drugs.split('，')))
                            type3.update(set(drugs.split('，')))
                        else:
                            self.type2.update(set(drugs.split('，')))
                            type2.update(set(drugs.split('，')))
                    else:
                        self.nccnunknow.add(genename)
                if any((type1, type2, type3)):
                    type1 = type1.difference(type3)
                    type2 = type2.difference(type3)
                    type2 = type2.difference(type1)
                    alterdrugtype1.update(type1)
                    alterdrugtype2.update(type2)
                    alterdrugtype3.update(type3)
                    alterinfo['cancerdrug'][cancertype]['type1'] = '，'.join(
                        sorted(type1, key=functools.cmp_to_key(drugcmp)))
                    alterinfo['cancerdrug'][cancertype]['type2'] = '，'.join(
                        sorted(type2, key=functools.cmp_to_key(drugcmp)))
                    alterinfo['cancerdrug'][cancertype]['type3'] = '，'.join(
                        sorted(type3, key=functools.cmp_to_key(drugcmp)))
                else:
                    self.funum -= 1
            alterinfo['mutiontype'] = genealter
            alterinfo['genealter'] = genealter
            alterinfo['freq'] = freq
            alterinfo['drugtype1'] = '，'.join(sorted(alterdrugtype1, key=functools.cmp_to_key(drugcmp)))
            alterinfo['drugtype2'] = '，'.join(sorted(alterdrugtype2, key=functools.cmp_to_key(drugcmp)))
            alterinfo['drugtype3'] = '，'.join(sorted(alterdrugtype3, key=functools.cmp_to_key(drugcmp)))
            self.result['cna'].append(alterinfo)

    def clinical(self):
        data = pd.DataFrame(self.sampledata['clinical'])
        self.result['clinical'] = list()
        if data.empty:
            return
        data.fillna('.', inplace=True)
        i = 0
        for index, row in data.iterrows():
            gene = row['geneName']
            if gene not in self.tbgene:
                continue
            if i == 10:
                break
            if len(row['condition']) < 60 and len(row['country']) < 30:
                tmp = dict(
                    nct_num=row['nct_num'],
                    condition=row['condition'],
                    study_first_posted=row['study_first_posted'],
                    phase=row['phase'],
                    country=row['country'],
                    intervention_name=row['intervention_name']
                )
                i += 1
                self.result['clinical'].append(tmp)

    def chemical(self):
        data = pd.DataFrame(self.sampledata['chemical'])
        self.result['chemical'] = list()
        if data.empty:
            return
        data.fillna('.', inplace=True)
        gene_chemical_arithmetic = chemicaldatabase['arithmetic']
        chemical_drug_effect = chemicaldatabase['drugeffect']
        tmedical1, tmedical2, tmedical3 = [], [], []
        smedical1, smedical2, smedical3 = [], [], []
        for medicine, data in data.groupby(['medicineName']):
            mgenelist = list()
            therapeutic_effect = []
            side_effect = []
            for _, row in data.iterrows():
                mgenelist.append(dict(genename=row['geneName'],
                                      rsid=row['rsId'],
                                      type=row['geneType'],
                                      guide=row['clinicalGuide'],
                                      level=row['evidenceLevel']
                                      ))
                ikey = "_".join([medicine, row['geneName'], row['rsId'], row['geneType']])
                stherapeutic_effect = gene_chemical_arithmetic[ikey]['8_therapeuticEffect']
                sside_effect = gene_chemical_arithmetic[ikey]['9_sideEffect']
                if stherapeutic_effect != '/' and stherapeutic_effect != '':
                    therapeutic_effect.append(stherapeutic_effect)
                if sside_effect != '/' and sside_effect != '':
                    side_effect.append(sside_effect)
            if therapeutic_effect:
                therapeuticlevel = sum([int(i) for i in therapeutic_effect if i != '/' and i != ''])
                try:
                    tkey = '{}-疗效={}'.format(medicine, therapeuticlevel)
                    therapeuticdescr = chemical_drug_effect[tkey]['3_effect']
                except KeyError:
                    self.logger.warning(f'{self.sample}样品中{medicine}疗效等级之和不在所在库中，请注意!')
                    therapeuticdescr = ''
            else:
                therapeuticdescr = ''
            if side_effect:
                side_effectlevel = sum([int(j) for j in side_effect if j != '/' and j != ''])
                try:
                    skey = '{}-毒副={}'.format(medicine, side_effectlevel)
                    side_effectdescr = chemical_drug_effect[skey]['3_effect']
                except KeyError:
                    self.logger.warning(f'{self.sample}样品中{medicine}毒副等级之和不在所在库中，请注意!')
                    side_effectdescr = ''
            else:
                side_effectdescr = ''
            if therapeuticdescr.find('疗效较好') > -1:
                tmedical1.append(medicine)
            elif therapeuticdescr.find('疗效中等') > -1:
                tmedical2.append(medicine)
            elif therapeuticdescr.find('疗效较差') > -1:
                tmedical3.append(medicine)
            else:
                pass
            if side_effectdescr.find('毒副较高') > -1:
                smedical3.append(medicine)
            elif side_effectdescr.find('毒副中等') > -1:
                smedical2.append(medicine)
            elif side_effectdescr.find('毒副较低') > -1:
                smedical1.append(medicine)
            else:
                pass
            descr = filter(lambda x: x is not '', [therapeuticdescr, side_effectdescr])
            medicaltype = ','.join(descr)
            medical = dict(
                name=medicine,
                medicalgene=mgenelist,
                medicaltype=medicaltype
            )
            self.result['chemical'].append(medical)
        for n in range(1, 4):
            curtname = 'tmedical' + str(n)
            cursname = 'smedical' + str(n)
            tdescribe = '，'.join(eval(curtname))
            sdescribe = '，'.join(eval(cursname))
            if not tdescribe:
                tdescribe = '无'
            if not sdescribe:
                sdescribe = '无'
            self.result['sum_hl'][curtname] = tdescribe
            self.result['sum_hl'][cursname] = sdescribe

    def immune(self):
        data = pd.DataFrame(self.sampledata['immune'])
        self.result['immune'] = list()
        if data.empty:
            return
        data.fillna('.', inplace=True)
        describe1list = []
        describe2list = []
        num = 4
        for geneinfo, data_1 in data.groupby(['clinicalGuide']):
            genecontent = list()
            for _, rows in data_1.iterrows():
                genename = rows['geneName']
                generesult = rows['testResult']
                if 'TMB' in genename:
                    cutoff, cutoffname = self.tmbcutoff()
                    try:
                        tmlist = re.split('[（）]', generesult)
                        tbcontent = '{}Muts/Mb，{}'.format(tmlist[0], tmlist[1])
                    except IndexError:
                        self.logger.error(f'{self.sample}免疫表TMB结果testResult不是常规格式')
                        raise Exception(f'{self.sample}免疫表TNB结果testResult不是常规格式')
                    newgene = f'肿瘤突变负荷{newline}（TMB）'
                    datainfo = '/'
                    self.result['sum_my']['tmb'] = '肿瘤突变负荷，{}({}Muts/Mb)</w:t><w:br/><w:t xml:space="preserve">'.format(
                        tmlist[1], tmlist[0])
                    self.result['tmb']['content'] = tmlist[0]
                    self.result['tmb']['describe'] = tmlist[1]
                    self.result['tmb']['cutoff'] = cutoff
                    self.result['tmb']['cutoffname'] = cutoffname
                    self.hepagene['tmb'] = tbcontent
                elif 'TNB' in genename:
                    try:
                        tnb = float(generesult)
                    except ValueError:
                        self.logger.error(f'{self.sample}免疫表TNB结果testResult不能转换成小数。')
                        raise Exception(f'{self.sample}免疫表TNB结果testResult不能转换成小数。')
                    if tnb < 5.33:
                        tnbvalue = '低'
                    elif 5.33 <= tnb < 12:
                        tnbvalue = '中'
                    else:
                        tnbvalue = '高'
                    newgene = f'肿瘤新生抗原负荷{newline}（TNB）'
                    tbcontent = '{} neo-peptide /Mb，{}'.format(tnb, tnbvalue)
                    datainfo = rows['dataInfo']
                    self.result['sum_my'][
                        'tnb'] = '肿瘤新生抗原负荷，{}({}neo-peptide/Mb)</w:t><w:br/><w:t xml:space="preserve">'.format(tnbvalue,
                                                                                                               tnb)
                    self.result['tnb']['tnb'] = str(tnb)
                    self.result['tnb']['tnbvalue'] = tnbvalue
                elif '微卫星分析' in genename:
                    try:
                        tmlist = re.split('[（）]', generesult)
                    except ValueError:
                        self.logger.error(f'{self.sample}免疫表MSI结果testResult不是常规格式。')
                        raise Exception(f'{self.sample}免疫表MSI结果testResult不是常规格式。')
                    self.result['sum_my']['msi'] = '微卫星状态，{}'.format(tmlist[1])
                    self.result['msi']['score'] = tmlist[0]
                    self.result['msi']['describe'] = tmlist[1]
                    newgene = f'微卫星状态{newline}（MS status）'
                    tbcontent = '{}'.format(tmlist[1])
                    datainfo = '/'
                    self.hepagene['msi'] = tbcontent
                else:
                    newgene = genename
                    tbcontent = rows['testResult']
                    datainfo = rows['dataInfo']
                    num += 1
                    if generesult.find('未检测') == -1:
                        describe = generesult.lstrip('检测到').rstrip('变异')
                        if rows['dataInfo'] == '正向':
                            describe1list.append('{}{}'.format(genename, describe))
                        elif rows['dataInfo'] == '负向':
                            describe2list.append('{}{}'.format(genename, describe))
                        else:
                            self.logger.warning('{}免疫结果表中的基因未能说明是正向还是负向！'.format(genename))
                genecontent.append(dict(newgene=newgene, tbcontent=tbcontent, datainfo=datainfo))
            self.result['immune'].append(dict(geneinfo=geneinfo, genecontent=genecontent))
        if describe1list:
            activegene = ','.join(describe1list)
        else:
            activegene = '未检出相关变异'
        if describe2list:
            passivegene = ','.join(describe2list)
        else:
            passivegene = '未检出相关变异'
        self.result['sum_my']['activegene'] = activegene
        self.result['sum_my']['passivegene'] = passivegene

    def msi(self):
        data = pd.DataFrame(self.sampledata['msi'])
        if data.empty:
            self.result['msi'] = {}
            return
        data.fillna('.', inplace=True)

        for _, rows in data.iterrows():
            self.result['msi']['describe'] = rows['level']
            self.result['msi']['score'] = '{}%'.format(rows['percent'])
        title = '微卫星状态，{}'.format(self.result['msi']['describe'])
        if self.result['msi']['describe'] == 'MSS' or self.result['msi']['describe'] == '微卫星稳定，MSS':
            describe = 'II期结直肠癌患者中，MSS亚组的患者预后较差。{}在II期结直肠癌患者中，与MSI-H的患者相比，' \
                       '5-FU辅助化疗治疗MSS的患者更能获益。'.format(newline)
        else:
            describe = 'II期结肠癌患者中，MSI-H亚组的患者预后较好。{}在II期结直肠癌患者中，' \
                       '与MSS的患者相比，5-FU的辅助化疗治疗MSI-H的患者不能获益。'.format(newline)
        self.result['sum_othermy']['describe'] = describe
        self.result['sum_othermy']['title'] = title

    def hla(self):
        data = pd.DataFrame(self.sampledata['hla'])
        if data.empty:
            self.result['hla'] = {}
            return
        data.fillna('.', inplace=True)
        result = []
        namedict = {'HLA-A': 'A', 'HLA-B': 'B', 'HLA-C': 'C'}
        for _, rows in data.iterrows():
            self.result['hla'][namedict[rows['locusName']]]['r2'] = rows['r2']
            self.result['hla'][namedict[rows['locusName']]]['r3'] = rows['r3']
            result.append(rows['r3'])
        num = result.count('杂合')
        if num == 3:
            desrcribe = '高'
        elif num == 0:
            desrcribe = '低'
        else:
            desrcribe = '中'
        self.result['sum_my']['hla'] = f'{desrcribe}{newline}'

    def addtnb(self):
        data = pd.DataFrame(self.sampledata['tnb'])
        self.result['addtnb'] = list()
        if data.empty:
            return
        data.fillna('.', inplace=True)
        for _, rows in data.iterrows():
            if rows['r3'] == '.':
                continue
            content = dict(
                r2=rows['r2'],
                locusName=rows['locusName'],
                r5=rows['r5'],
                r3='{:.2f}'.format(float(rows['r3']))
            )
            self.result['addtnb'].append(content)

    def heredity(self):
        data = pd.DataFrame(self.sampledata['heredity'])
        self.result['heredity']['sign'] = list()
        self.result['heredity']['nosign'] = list()
        if data.empty:
            self.result['sum_yc']['describe'] = '未发现相关基因致病/疑似致病变异'
            return
        data.fillna('.', inplace=True)
        describelist = []
        tmdict = tbtranslate
        for info, data_1 in data.groupby(['geneName', 'alteration']):
            genename, alteration = info
            alterinfo = data_1.iloc[0].to_dict()
            if alterinfo['geneName'] == '.':
                continue
            if alterinfo['exonicfuncRefgene'] in tmdict:
                mutation = tmdict[alterinfo['exonicfuncRefgene']]['translate']
            else:
                mutation = alterinfo['exonicfuncRefgene']
            describelist.append('{}基因 {} {} {},与{}相关，疑似胚系突变。'.format(
                genename, alterinfo['hgsGb'], alterinfo['ajsGb'], mutation, alterinfo['relativeDisease']))
            if alterinfo['mutationType'] != '杂合突变':
                alterinfo['tranmutationType'] = True
            else:
                alterinfo['tranmutationType'] = False
            alterinfo['mutation'] = mutation
            if alterinfo['level'].startswith('5'):
                alterinfo['tranlevel'] = True
                self.result['heredity']['sign'].append(alterinfo)
            elif alterinfo['level'].startswith('4'):
                alterinfo['tranlevel'] = False
                self.result['heredity']['sign'].append(alterinfo)
            elif alterinfo['level'].startswith('3'):
                alterinfo['tranlevel'] = False
                self.result['heredity']['nosign'].append(alterinfo)
            else:
                pass
        if not describelist:
            self.result['sum_yc']['describe'] = '未发现致病/疑似致病的相关基因变异'
        else:
            self.result['sum_yc']['describe'] = f'{newline}'.join(describelist)

    def tmbcutoff(self):
        """
        tmb阈值
        :return:
        """
        # treatid = self.sampledata['cms']['treatID']
        if self.source == '血液(cfDNA)':
            cutoff = f'低：TMB≤10.59；{newline}中：10.59＜TMB≤15.88；{newline}高：TMB＞15.88'
            cutoffname = '游离DNA参考阈值'
        else:
            if 'TS01' in self.treatid:
                cutoff = f'低：TMB≤6.47；{newline}中：6.47＜TMB≤12.94 ；{newline}高：TMB＞12.94 '
                cutoffname = '肺癌参考阈值'
            elif 'TS2101' in self.treatid:
                cutoff = 'f低：TMB≤8.35；{newline}中：8.35＜TMB≤17.06 ；{newline}高：TMB＞17.06 '
                cutoffname = '肝癌参考阈值'
            else:
                cutoff = f'低：TMB≤6.47；{newline}中：6.47＜TMB≤13.53 ；{newline}高：TMB＞13.53  '
                cutoffname = '实体瘤参考阈值'
        return cutoff, cutoffname

    def pdl1(self):
        """
        pdl1处理
        :return:
        """
        self.result['pdl1'] = list()
        self.result['sum_my']['pdl1'] = ''
        if self.itemName.find('PDL1蛋白表达') > -1:
            path = '/pdl1'
            samplepath = os.path.join(path, self.sample + '.xlsx')
            try:
                df = pd.read_excel(samplepath, header=0)
            except FileNotFoundError:
                self.logger.warning('{}没有上传PDL1的结果！'.format(self.sample))
                return
            df = df.fillna('.')
            pdl1describelist = list()
            for _, rows in df.iterrows():
                if rows['TPS'] == '.':
                    self.logger.warning('{}的TPS结果为空，该{}抗体将不显示！'.format(self.sample, rows['抗体名称']))
                    continue
                # 检测图片文件是否存在
                pdl1type = rows['抗体名称']
                picpath = os.path.join('/pdl1', pdl1type)
                picpathe = os.path.join(picpath, 'HE')
                picpbthe = os.path.join(picpath, 'positive')
                picpcthe = os.path.join(picpath, 'negative')
                piclist = [[picpath, '结果检测图片'],
                           [picpathe, 'HE图片'],
                           [picpbthe, 'positive图片'],
                           [picpcthe, 'negative图片']]
                pictag = True
                for pic in piclist:
                    if not os.path.exists(os.path.join(pic[0], self.sample + '.png')):
                        self.logger.warning('{} {}抗体中{}未上传！'.format(self.sample, rows['抗体名称'], pic[1]))
                        pictag = False
                        break
                if not pictag:
                    continue
                tps = rows['TPS']
                cps = rows['CPS']
                tcc = rows['肿瘤细胞含量']
                if cps != '.':
                    p3 = ', CPS {}'.format(cps)
                else:
                    p3 = ''
                if type(tps) is float:
                    rtps = '{:.0%}'.format(float(tps))
                    if tps <= 0.01:
                        p2 = '无PD-L1表达'
                    elif 0.01 < tps <= 0.49:
                        p2 = '有PD-L1表达'
                    else:
                        p2 = 'PD-L1高表达'
                elif type(tps) is str:
                    if tps.find('<') > -1:
                        p2 = '无PD-L1表达'
                    else:
                        p2 = '有PD-L1表达'
                    rtps = '{}'.format(tps)
                else:
                    logging.warning('PDL1的表格误！报告当PDL1阴性处理。')
                    continue
                if type(tcc) == float:
                    tcc = '{:.0%}'.format(tcc)
                pdl1describe = 'PD-L1免疫组化分析（Dako {}），{}（TPS {}{}）'.format(pdl1type, p2, rtps, p3)
                pdl1describelist.append(pdl1describe)
                describe1 = '({})'.format(p2)
                self.result['pdl1'].append(dict(rtps=rtps, tcc=tcc, cps=cps))
                self.hepagene['pdl1'] += 'Dako {}），{}（TPS {}{}）'.format(pdl1type, p2, rtps, p3)
            self.result['sum_my']['pdl1'] = '{}{}'.format(f'{newline}'.join(pdl1describelist), newline)

    def hrd(self):
        """
        :return:
        """

        self.result['sigma'] = tree()
        self.result['hrd'] = tree()
        testhrdgene = self.detectgene['geneHrd'].split(',')
        describe = list()
        describe1 = list()
        describe2 = list()
        for gene in testhrdgene:
            flag = False
            freq = '/'
            if self.itemid in ['0201027', '0201028', '0204026']:
                if gene in self.hrdgene:
                    flag = True
                    if gene == 'BRCA1' or gene == 'BRCA2':
                        describe.append(f'{gene}基因突变，提示PARP抑制剂敏感')
                    else:
                        describe.append(f'{gene}基因突变，提示PARP抑制剂治疗潜在敏感')
                    freq = self.hrdgene[gene]['freq']

            elif self.itemid in ['0201023', '0201024', '202028', '202027']:
                if gene in self.hrdgene:
                    flag = True
                    if gene == 'BRCA1' or gene == 'BRCA2':
                        if self.hrdgene[gene]['level'].startswith('5'):
                            describe1.append(f'{gene}基因致病性变异，PARP抑制剂可能敏感')
                        elif self.hrdgene[gene]['level'].startswith('4'):
                            describe1.append(f'{gene}疑似致病性变异，PARP抑制剂可能敏感')
                        else:
                            pass
                    else:
                        if self.hrdgene[gene]['level'].startswith('5'):
                            describe2.append(f'{gene}基因致病性变异，PARP抑制剂潜在敏感')
                        elif self.hrdgene[gene]['level'].startswith('4'):
                            describe2.append(f'{gene}疑似致病性变异，PARP抑制剂潜在敏感')
                        else:
                            pass
                    freq = self.hrdgene[gene]['freq']
            self.result['hrd'][gene]['freq'] = freq
            self.result['hrd'][gene]['f'] = flag
            if describe:
                self.result['sum_hrd']['describe'] = f'{newline}'.join(describe)
            else:
                self.result['sum_hrd']['describe'] = '未见HR-DDR相关基因有害/疑似有害变异'
            if describe1:
                self.result['sum_hrd']['describe1'] = f'{newline}'.join(describe1)
            else:
                self.result['sum_hrd']['describe1'] = '未见BRCA1和BRCA2基因有害/疑似有害变异'
            if describe2:
                self.result['sum_hrd']['describe2'] = f'{newline}'.join(describe2)
            else:
                self.result['sum_hrd']['describe2'] = '未见其他HRD-DDR相关基因有害/疑似有害变异'

        data = pd.DataFrame(self.sampledata['hrr'])
        if not data.empty:
            hrr = data.iloc[0].to_dict()
            if hrr['clinicalAdvice'] == '.' or not hrr['clinicalAdvice']:
                clinicaladvice = hrr['testResult']
            else:
                clinicaladvice = '{},{}'.format(hrr['testResult'], hrr['clinicalAdvice'])
            self.result['sum_hrd']['describe3'] = 'SNVs count {}，Sig3_Modest{}，Sig3_Stringent{}，{}'.format(
                hrr['totalSnvs'], hrr['passMva'], hrr['passMvaStrict'], clinicaladvice)
            self.result['sigma']['count'] = hrr['totalSnvs']
            self.result['sigma']['modest'] = hrr['passMva']
            self.result['sigma']['stringent'] = hrr['passMvaStrict']
            self.result['sigma']['conclusion'] = clinicaladvice
            pt = glob.glob('/home/zc/files/figures/' + self.sample + '*png')[0]
            if not os.path.exists(pt):
                self.logger.error(f'{self.sample}样品没有上传sigma图片，联系流程人员！')
                raise Exception(f'{self.sample}样品没有上传sigma图片，联系流程人员！')
            self.pic.append(dict(name='sigma', path=pt, width=17.98, height=7.06))

            path = os.path.join(rootpath, 'database', 'detectdetail', 'hrr')
            pica = [os.path.join(path, 'hrra.png'), 14.47, 4.39]
            picb = [os.path.join(path, 'hrrb.png'), 14.82, 4.84]
            picc = [os.path.join(path, 'hrrc.png'), 17.07, 6.06]
            picd = [os.path.join(path, 'hrrd.png'), 17.01, 3.86]
            pice = [os.path.join(path, 'hrre.png'), 14.74, 6.51]
            piclist = [pica, picb, picc, picd, pice]
            for index, pic in enumerate(['a', 'b', 'c', 'd', 'e']):
                self.pic.append(dict(name=f'detectdetail_sigma_{pic}',
                                     path=piclist[index][0],
                                     width=piclist[index][1],
                                     height=piclist[index][2])
                                )

    def hepa(self):
        """
        肝胆肿瘤模块
        :return:
        """
        self.result['hepa'] = tree()
        if 'TS21' not in self.treatid:
            return
        hepagene = self.detectgene['geneGand'].split(',')
        if not hepagene:
            return
        for gene in hepagene:
            flag = False
            result = ''
            if gene in self.hepagene['target']:
                result += '基因突变'
                flag = True
            if gene in self.hepagene['fusion']:
                result += '基因融合'
                flag = True
            if gene in self.hepagene['cna']:
                result += '基因扩增'
                flag = True
            if not result:
                result += '未检测到变异'
            self.result['hepa'][gene]['f'] = flag
            self.result['hepa'][gene]['r'] = result
        for typename in ['tmb', 'msi', 'pdl1']:
            if typename in self.hepagene:
                self.result['hepa'][typename]['r'] = self.hepagene[typename]
            else:
                self.result['hepa'][typename]['r'] = 'NA'

    def testnccn(self):
        """nccn基因"""
        self.result['nccn'] = tree()
        id = ''
        for nccnid in nccn.keys():
            if nccnid in self.treatid:
                id = nccnid
                break
        if not id:
            return
        treadidnccn = nccn[id]
        self.result['nccn']['treadid'] = id
        data = pd.DataFrame(treadidnccn)
        if data.empty:
            return
        for _, rows in data.iterrows():
            gene = rows['检测基因']
            flag = False
            if gene == 'info' or gene == '微卫星状态' or gene == 'treatresult':
                continue
            else:
                result = ''
                if gene in self.nccngene['unknow']:
                    result = '临床意义未明变异'
                else:
                    tresult = set()
                    tmlist = rows['检测范围'].split('，')
                    for t in tmlist:
                        if t == '突变' and gene in self.nccngene['target']:
                            tresult.add('基因突变')
                        elif t == '融合' and gene in self.nccngene['fusion']:
                            tresult.add('基因重排')
                        elif t == '扩增' and gene in self.nccngene['cna']:
                            tresult.add('基因扩增')
                    result += f'{newline}'.join(tresult)
            if not result:
                result = '未检测到变异'
            if result.startswith('基因'):
                flag = True
            self.result['nccn'][gene]['f'] = flag
            self.result['nccn'][gene]['r'] = result

    def sign(self):
        reporttime = datatran(flag=1)
        self.result['sign']['detecttime'] = self.detecttime
        self.result['sign']['reporttime'] = reporttime

        signname1 = self.custom['signname1']
        signname2 = self.custom['signname2']
        seal = self.custom['isSeal']
        if signname1 == '无':
            signname1 = ' '
        if signname2 == '无':
            signname2 = ' '
        flag = False
        if seal == '显示':
            signname2 = 'seal' + signname2
            flag = True

        sign1 = os.path.join(rootpath, 'database', 'custom', 'sign', 'signname1', f'{signname1}.png')
        if signname1 == ' ' or not os.path.exists(sign1):
            pass
        else:
            self.pic.append(dict(name='sign_det', path=sign1, width=2.29, height=1.27))

        sign2 = os.path.join(rootpath, 'database', 'custom', 'sign', 'signname2', signname2 + '.png')
        if signname2 == ' ' or not os.path.isfile(sign2):
            pass
        else:
            if flag:
                self.pic.append(dict(name='sign_view', path=sign2, width=4.22, height=3.02))
            else:
                self.pic.append(dict(name='sign_view', path=sign2, width=2.29, height=1.27))
        sign3 = os.path.join(rootpath, 'database', 'custom', 'sign', 'signname3', '开震天.png')
        sign4 = os.path.join(rootpath, 'database', 'custom', 'sign', 'signname3', '王慧玲.png')
        self.pic.append(dict(name='sign_kai', path=sign3, width=2.29, height=1.27))
        self.pic.append(dict(name='sign_huiling', path=sign4, width=2.29, height=1.27))

    def falseresult(self):
        targetgene = self.detectgene['geneBx'].split(',')
        fusiongene = self.detectgene['geneCp'].split(',')
        cnagene = self.detectgene['geneKb'].split(',')
        self.result['falseresult']['target'] = list()
        self.result['falseresult']['fusion'] = list()
        self.result['falseresult']['cna'] = list()
        if not (self.trueresult['target'] | self.trueresult['fusion'] | self.trueresult['cna']):
            self.result['falseresult']['falseall'] = True
        else:
            self.result['falseresult']['falseall'] = False
        for tgene in targetgene:
            if tgene not in self.trueresult['target']:
                self.result['falseresult']['target'].append(tgene)
        for fgene in fusiongene:
            if fgene not in self.trueresult['fusion']:
                self.result['falseresult']['fusion'].append(fgene)
        for cgene in cnagene:
            if cgene not in self.trueresult['cna']:
                self.result['falseresult']['cna'].append(cgene)

    def specialtreatresult(self):
        """
        TS05结直肠癌;TS10头颈部鳞状细胞癌添加药物
        :return:
        """
        self.result['specialtarget'] = list()
        tbgenelist = detectinfo[self.itemid]['geneTb']
        if 'KRAS' in tbgenelist and ('KRAS' not in self.tbgene) and ('NRAS' not in self.tbgene) \
                and ('HRAS' not in self.tbgene):
            if 'TS05' in self.treatid:
                self.type1.update({'西妥昔单抗', 'Panitumumab[帕尼单抗]'})
            elif 'TS10' in self.treatid:
                self.type2.update({'西妥昔单抗', 'Panitumumab[帕尼单抗]'})
            else:
                return
            if 'NRAS' in tbgenelist:
                gene = 'RAS 野生型'
                geneinfo = 'RAS 野生型基因描述'
            else:
                gene = 'KRAS 野生型'
                geneinfo = 'KRAS 野生型基因描述'
            self.tbgene.add(gene)
            alter = tree()
            alter['alteration'] = '野生型'
            alter['geneName'] = gene
            alter['freq'] = '/'
            alter['geneInfo'] = geneinfo
            alter['genealter'] = gene
            alter['cancerdrug']['结直肠癌']['type3'] = ''
            alter['cancerdrug']['结直肠癌']['type2'] = ''
            alter['cancerdrug']['结直肠癌']['type1'] = '西妥昔单抗, Panitumumab[帕尼单抗]'
            self.result['specialtarget'].append(alter)

    def specialagent(self):
        """
        特殊供应商添加特殊位点
        :return:
        """
        gene = 'RAS 野生型'
        geneinfo = 'RAS 野生型基因描述'
        if self.unitagent_name == '吴雄飞' and ('RAS 野生型' not in self.tbgene) and ('KRAS 野生型' not in self.tbgene) \
                and ('KRAS' not in self.tbgene) and ('NRAS' not in self.tbgene) and ('HRAS' not in self.tbgene):
            if 'TS05' in self.treatid:
                self.type1.update({'西妥昔单抗', 'Panitumumab[帕尼单抗]'})
            else:
                self.type2.update({'西妥昔单抗', 'Panitumumab[帕尼单抗]'})
            alter = tree()
            alter['alteration'] = '野生型'
            alter['geneName'] = gene
            alter['freq'] = '/'
            alter['geneInfo'] = geneinfo
            alter['genealter'] = gene
            alter['cancerdrug']['结直肠癌']['type3'] = ''
            alter['cancerdrug']['结直肠癌']['type2'] = ''
            alter['cancerdrug']['结直肠癌']['type1'] = '西妥昔单抗, Panitumumab[帕尼单抗]'
            self.result['specialtarget'].append(alter)

    def addpict(self):
        path = os.path.join(rootpath, 'database', 'template', '8pic')
        detectdetail_tnba = [os.path.join(path, 'detectdetail_tnb_a.png'), 11.4, 7.2]
        detectdetail_tnbb = [os.path.join(path, 'detectdetail_tnb_b.png'), 8.5, 7.2]
        detectdetail_tnbc = [os.path.join(path, 'detectdetail_tnb_c.png'), 11.85, 5.62]
        detectdetail_tnbd = [os.path.join(path, 'detectdetail_tnb_d.png'), 10.28, 6.45]
        detectdetail_tnblist = [detectdetail_tnba, detectdetail_tnbb, detectdetail_tnbc, detectdetail_tnbd]
        for index, itm in enumerate(['a', 'b', 'c', 'd']):
            self.pic.append(dict(name=f'detectdetail_tnb_{itm}',
                                 path=detectdetail_tnblist[index][0],
                                 width=detectdetail_tnblist[index][1],
                                 height=detectdetail_tnblist[index][2]))

        picmsi = [os.path.join(path, 'detectdetail_msi_a.png'), 17.17, 4.02]
        self.pic.append(dict(name='detectdetail_msi_a',
                             path=picmsi[0],
                             width=picmsi[1],
                             height=picmsi[2]))

        detectdetail_sigmaa = [os.path.join(path, 'detectdetail_sigmaa.png'), 14.47, 4.39]
        detectdetail_sigmab = [os.path.join(path, 'detectdetail_sigmab.png'), 14.82, 4.84]
        detectdetail_sigmac = [os.path.join(path, 'detectdetail_sigmac.png'), 17.07, 6.06]
        detectdetail_sigmad = [os.path.join(path, 'detectdetail_sigmad.png'), 17.01, 3.86]
        detectdetail_sigmae = [os.path.join(path, 'detectdetail_sigmae.png'), 14.74, 6.51]
        detectdetail_sigmalist = [detectdetail_sigmaa, detectdetail_sigmab, detectdetail_sigmac,
                                  detectdetail_sigmad, detectdetail_sigmae]
        for index, itm in enumerate(['a', 'b', 'c', 'd', 'e']):
            self.pic.append(dict(name=f'detectdetail_sigma_{itm}',
                                 path=detectdetail_sigmalist[index][0],
                                 width=detectdetail_sigmalist[index][1],
                                 height=detectdetail_sigmalist[index][2]))

    def sqrcode(self):
        self.result['qrcode'] = '{{ pict.qrcode }}'
        path = os.path.join(rootpath, 'result', 'qrcode')
        qrcodeself(path, self.sample)
        pic = os.path.join(path, self.sample + '.png')
        self.pic.append(dict(name='qrcode',
                             path=pic,
                             width=2.49,
                             height=2.49)
                        )
        # 生成html，数字加密符
        try:
            self.result['src'] = makesrc.makesrc.generate_webpage(self.result, self.detectgene)
        except:
            print(self.sample)


def run(path):
    """
    :param path:
    :return:
    """
    sampleresults = json.loads(readxlsx.read(path))
    total = []
    # 生成结果
    for sampleinfo in sampleresults:
        paraseresult = Parase(sampleinfo)
        result = paraseresult.run()
        total.append(result)
    resultjson = json.dumps(total, sort_keys=True, indent=4, ensure_ascii=False, cls=NpEncoder)
    # with open('para.json', 'w') as f:
    #     f.write(resultjson)
    return resultjson
