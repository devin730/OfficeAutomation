#!/usr/bin/python
# coding: utf-8
#
# @descroptions:
# all .doc files (including subdirectories) are converted into .docx format
# if param dir_path is not set，then this .py file path is used

import os
from win32com import client
import time

class convAllDoc2Docx():
    def __init__(self, dir_path=os.getcwd()):
        ext = ".doc"
        file_list = self.find_file(dir_path, ext)
        for file in file_list:
            print("file = ", file)
            self.doc_to_docx(file)

    def doc_to_docx(self, path):
        if os.path.splitext(path)[-1] == ".doc":
            word = client.Dispatch('Word.Application')
            doc = word.Documents.Open(path)
            doc.SaveAs(os.path.splitext(path)[0]+".docx", 16)
            doc.Close()
            word.Quit()
            print(path + " is converted into " + os.path.splitext(path)[0]+".docx")
            time.sleep(3)

    def find_file(self, path, ext, file_list=[]):
        dir = os.listdir(path)
        for i in dir:
            i = os.path.join(path, i)
            if os.path.isdir(i):
                self.find_file(i, ext, file_list)
            else:
                if ext == os.path.splitext(i)[1]:
                    file_list.append(i)
        return file_list
 
if __name__ == '__main__':
    convAllDoc2Docx()