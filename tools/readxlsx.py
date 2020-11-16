# -*- coding: utf-8 -*-

import pandas as pd
import logging
import json
import sys
from tools import database


logger = logging.getLogger('main.sub')
sheetname = database.sheetname()


def read(merge):
    result = list()
    df = pd.read_excel(merge, None)
    samplelist = df['样本信息']['sampleSn'].to_list()
    if not samplelist:
        logger.error('样本信息表为空！读取excel信息失败！')
        raise Exception('样本信息表为空！读取excel信息失败！')
    for sample in samplelist:
        if sample == '样本编码':
            continue
        samdict = dict()
        for name, contents in df.items():
            sheet = sheetname[name]
            if contents.empty:
                samdict[sheet] = []
                continue
            contents.fillna('.', inplace=True)
            sampledf = contents[contents['sampleSn'] == sample]
            if sampledf.empty:
                samdict[sheet] = []
            else:
                samdict[sheet] = [rows for _, rows in sampledf.to_dict('index').items()]
        result.append(samdict)
    resultjson = json.dumps(result, sort_keys=True, indent=4, ensure_ascii=False)
    return resultjson


if __name__ == '__main__':
    myjson = read(sys.argv[1])
    print(myjson)
