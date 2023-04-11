import os
import pandas
import tempfile
import re
import sys
import requests
import hashlib
import shutil
from pathlib import Path


def checkFile(path):
    """
    Checks if the input 'path' is an excel file an can be correctly read, returns a boolean
    """
    try:
        data = pandas.ExcelFile(path, engine='openpyxl')
        return True
    except:
        return False



def hash_file(filename):
   """"This function returns the SHA-1 hash
   of the file passed into it"""

   # make a hash object
   h = hashlib.sha1()

   # open file for reading in binary mode
   with open(filename,'rb') as file:

       # loop till the end of the file
       chunk = 0
       while chunk != b'':
           # read only 1024 bytes at a time
           chunk = file.read(1024)
           h.update(chunk)

   # return the hex representation of digest
   return h.hexdigest()

def gdriveToXMLX(url, path):
    temp = tempfile.NamedTemporaryFile(prefix="mapeathor-gdrive", delete=False, suffix=".xlsx") 
    localfile = path+'.xlsx'
    m = re.search(r'(?:file|spreadsheets)\/d\/(.*)\/', url)

    if not m:
        raise Exception("Malformed Google Spreadsheets URL")
        sys.exit()

    docid = m.groups()[0]

    url = 'https://docs.google.com/spreadsheets/d/'+docid+'/export?exportFormat=xlsx'

    r = requests.get(url)

    temp.write(r.content)
    temp.close()
    if not (Path(localfile).is_file()) or hash_file(temp.name) != hash_file(localfile):
        shutil.copyfile(temp.name, localfile)
        if Path(path+'.yaml').is_file():
            os.remove(path+'.yaml')
        return localfile
    else:
        return ''

