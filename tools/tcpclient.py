# -*- coding: utf-8 -*-

import json
import os
import socket
import struct


def recvdata(conn, path):
    """
    接受文件
    :param conn:
    :param path:
    :return:
    """
    header_size = struct.unpack('i', conn.recv(4))[0]
    header_bytes = conn.recv(header_size)
    header_json = header_bytes.decode('utf-8')
    header_dic = json.loads(header_json)
    content_len = header_dic['contentlen']
    content_name = header_dic['contentname']
    recv_len = 0
    fielpath = os.path.join(path, content_name)
    file = open(fielpath, 'wb')
    while recv_len < content_len:
        correntrecv = conn.recv(1024 * 1000)
        file.write(correntrecv)
        recv_len += len(correntrecv)
    file.close()
    return fielpath


def senddata(conn, path, message=None):
    name = os.path.basename(os.path.realpath(path))
    if not message:
        with open(path, 'rb') as file:
            content = file.read()
        headerdic = dict(
            contentlen=len(content),
            contentname=name
        )
        headerjson = json.dumps(headerdic)
        headerbytes = headerjson.encode('utf-8')
        headersize = len(headerbytes)
        conn.send(struct.pack('i', headersize))
        conn.send(headerbytes)
        conn.sendall(content)
    else:
        headerdic = dict(
            contentlen=len(path),
            contentname='message'
        )
        headerjson = json.dumps(headerdic)
        headerbytes = headerjson.encode('utf-8')
        headersize = len(headerbytes)
        conn.send(struct.pack('i', headersize))
        conn.send(headerbytes)
        conn.sendall(path.encode('utf-8'))


# def client(treport):
#     """
#
#     :param treport:
#     :return:
#     """
#     myserver = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
#     if treport:
#         adrss = ("", 7799)
#     else:
#         adrss = ("", 7788)
#     myserver.bind(adrss)
#     myserver.listen(10)
#     return myserver
