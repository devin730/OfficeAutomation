#!/usr/bin/python
# coding: utf-8

# @description:
# generate numbers of .docx files based on given template .docx file.

import xlrd
from EditWordFile import WordPro as replace

# read .xlsx file to load data.
excel_file_path = "./data/info.xlsx"
data_excel = xlrd.open_workbook(excel_file_path)
table_excel = data_excel.sheet_by_index(0)
rows_cnt = table_excel.nrows


# generate a dictionary about replace strings
# key should be set in .docx file
# value is obtained in .xlsx file
dict_1 = {'name': '', 'id': '', 'grade': '', 'money': ''}
for i in range(1, rows_cnt):
    print(table_excel.row_values(i))
    dict_1['name'] = str(table_excel.row_values(i)[0])
    dict_1['id'] = str(table_excel.row_values(i)[1])
    dict_1['grade'] = str(table_excel.row_values(i)[2])
    dict_1['money'] = str(table_excel.row_values(i)[3])
    new_file_path = './data/' + dict_1['name'] + '.docx'
    replace(in_dict=dict_1, origin_word_file='./data/template.docx', new_word_file=new_file_path)

# !IMPORTANT NOTICE
# !Note: text underline will not influence the string matching
# !Note: you can not use symbols such as '.', '<', '>' in keys of applied dictionary
# !      Symbols in keys will influence the string matching