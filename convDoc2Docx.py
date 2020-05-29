#!/usr/bin/python
# coding: utf-8
#
# @descroptions:
# selected .doc files is converted into .docx format
# path should be complete, containing both "xxx.doc" and "C、D、E、:/"

import os
from win32com import client
import time

def convDocToDocx(path):
    if os.path.splitext(path)[-1] == ".doc":
        new_path = os.path.splitext(path)[0]+".docx"  # new .docx file path is set
        word = client.Dispatch('Word.Application')
        doc = word.Documents.Open(path)  # open original .doc file
        doc.SaveAs(new_path, 16)  # save converted file
        doc.Close()
        word.Quit()
        print(path + " is converted into " + new_path)
        time.sleep(3)  # ensure safety
        return(new_path)
    else:
        print('your input file is not .doc file.')

if __name__ == '__main__':
    path = os.getcwd()
    convDocToDocx(path+'/data/1.doc')