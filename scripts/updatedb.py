# -*- coding: utf-8 -*-

import django
import os
import sys
import pandas as pd

BASE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'djreport')
sys.path.append(BASE_DIR)
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "djreport.settings")
django.setup()
from exceldata.models import Batches, Samples, Target, Heredity
from django.contrib.auth.models import User


def update(path):
    """
    excle信息插入到sqlite3
    :param path:
    :return:
    """


    # 存入批次信息
    batch = os.path.basename(path)
    author = User.objects.get(username='chao')

    newbatch = Batches.objects.create(
        batch=batch,
        author=author
    )
    newbatch.save()


    cms = pd.read_excel(path, sheet_name='样本信息', header=0, skiprows=[1])
    samplelist = cms['sampleSn'].to_list()
    target = pd.read_excel(path, sheet_name='靶向', dtype={'drugLevel': str}, header=0, skiprows=[1])
    here = pd.read_excel(path, sheet_name='遗传',  header=0, skiprows=[1])
    fusion = pd.read_excel(path, sheet_name='融合', dtype={'drugLevel': str}, header=0, skiprows=[1])

    for sample in samplelist:

        # 存入样品信息
        if not Samples.objects.filter(sampleSn=sample).exists():
            newsamples = Samples.objects.create(
                sampleSn=sample,
                batch=newbatch,
            )
        else:
            oldsample = Samples.objects.get(sampleSn=sample)
            time = oldsample.time
            oldsample.delete()
            newsamples = Samples.objects.create(
                sampleSn=sample,
                batch=newbatch,
                time=time
            )
        newsamples.save()

    # 靶向
        sampletarget = target[target['sampleSn'] == sample]
        if not sampletarget.empty:
            newtargets = sampletarget.to_dict('index')
            for _, newtarget in newtargets.items():
                newtarget['sampleSn'] = newsamples
                newtarget['_GeneName'] = newtarget.pop('GeneName')
                newtarget['_1000g2015aug_all'] = newtarget.pop('1000g2015aug_all')
                Target.objects.create(**newtarget)

    # 遗传
        samplehere = here[here['sampleSn'] == sample]
        if not sampletarget.empty:
            newheres = samplehere.to_dict('index')
            for _, newhere in newheres.items():
                newhere['sampleSn'] = newsamples
                Heredity.objects.create(**newhere)
    # 融合
        samplefusion = fusion[fusion['sampleSn'] == sample]
        if not samplefusion.empty:
            newfus = samplefusion.to_dict('index')
            for _, newfu in newfus.items():
                newfu['sampleSn'] = newsamples
                Heredity.objects.create(**newfu)





if __name__ == '__main__':
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'example', '泛癌种测试数据2.xlsx')
    update(path)
