# -*- coding: utf-8 -*-


from tools import parasexlsx
from tools import common
from tools import parasechemxls
import report
import json
import os
from multiprocessing import Pool
import time
import logging
from tools.tcpclient import recvdata, senddata
from tools import config
import socket
import sys


basedir = os.path.dirname(os.path.realpath(__file__))
logger = logging.getLogger('main.sub')
config = config.config()


def run(resultfile):
    t1 = time.time()
    filename = os.path.basename(resultfile)
    filelist = list()
    xlsxname = filename.rstrip('.xlsx')
    log = common.reportlog(xlsxname)
    if not config:
        return
    if not filename.endswith('.xlsx'):
        logger.error('{}不是xlsx文件'.format(filename))
        return
    if filename.startswith('化疗'):
        resutltjson = parasechemxls.run(resultfile)
    else:
        resutltjson = parasexlsx.run(resultfile)
    result = json.loads(resutltjson)
    reportlist = list()
    pp = Pool(10)
    for sampledata in result:
        try:
            reportlist.append(pp.apply_async(report.createword.run, args=(sampledata,)))
        except:
            continue
    pp.close()
    pp.join()
    wordlist = [file.get() for file in reportlist]
    for word in wordlist:
        if not word:
            continue
        filelist.append(word)
        sample = os.path.basename(word).rstrip('.docx')
        sword = os.path.join(basedir, 'result', 'simpleword', 'S{}.docx'.format(sample))
        spdf = os.path.join(basedir, 'result', 'pdf', 'S{}.pdf'.format(sample))
        pdfpath = os.path.join(basedir, 'result', 'pdf', '{}.pdf'.format(sample))
        if os.path.exists(sword):
            filelist.append(sword)
    filelist.append(log)
    filezip = common.filetozip(filelist, xlsxname)
    t2 = time.time()
    print(t2 - t1)
    return filezip


def server():
    myserver = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if config['treport']:
        adrss = ("", 7797)
    else:
        adrss = ("", 7787)
    myserver.bind(adrss)
    myserver.listen(10)
    while True:
        myclient, adddr = myserver.accept()
        recv_content = recvdata(myclient, os.path.join(basedir, 'example'))
        myfilezip = run(recv_content)
        senddata(myclient, myfilezip)
        print('发送成功！')


if __name__ == '__main__':
    if len(sys.argv) > 1:
        run(sys.argv[1])
    else:
        server()
