__doc__ = """
create a encrypted code for a file
"""
__author__ = "Kai"
__version__ = "v1.1"


from Crypto.Signature import PKCS1_v1_5
from Crypto.PublicKey import RSA
from Crypto.Hash import SHA256
import os


def create_digital_signature(key_info, der):
    """
    create a digital signature from hashed info
    :param key_info
    :return:
    """
    private_key = _read_key(None, der)
    hashed_info = _create_hashed_info(key_info)
    signer = PKCS1_v1_5.new(private_key)
    signature = signer.sign(hashed_info)
    return SHA256.new(signature.hex().encode()).hexdigest()


def _create_hashed_info(key_info):
    """
    create hashed info using crypto hashing SHA256 (secure) algorithms
    :param key_info: a json file that contains key information on report
    :return:
    """
    hashed_obj = SHA256.new()
    with open(key_info, "rb") as fh:
        for line in fh:
            hashed_obj.update(line)
    return hashed_obj


def _read_key(passphrase, der):
    """
    read the private key
    :return:
    """
    with open(der, 'rb') as fh:
        data = fh.read()
        private_key = RSA.import_key(data, passphrase=passphrase)
    return private_key
