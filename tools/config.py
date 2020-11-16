# -*- coding: utf-8 -*-

"""
配置文件，处理开发环境、测试环境、生产环境自动通用处理
"""

import os
import logging


logger = logging.getLogger('main.sub')


venv = os.path.basename(os.path.dirname(os.path.dirname(__file__)))


def config():
    rconfig = dict(
        mreport=True,
        treport=True

    )
    if venv == 'mreport':
        logger.info('开发环境')
        return rconfig
    elif venv == 'treport':
        logger.info('测试环境')
        rconfig['mreport'] = False
        return rconfig
    elif venv == 'report':
        rconfig['mreport'] = False
        rconfig['treport'] = False
        return rconfig
    else:
        print('不在指定路径运行！')
        return
