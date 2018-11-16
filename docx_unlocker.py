# -*- coding: utf-8 -*-
"""
Created on Wed Oct 24 10:41:48 2018

@author: moskalev_iv
"""

import os
import shutil


FILE_NAME = '1521.122.06ДО Ч.3 - ВМ_v00.100155186390074'
DIR_NAME = os.path.dirname(os.path.abspath(FILE_NAME + '.docx'))
f_name = DIR_NAME + '\\@' + FILE_NAME + '.docx'
d_name = DIR_NAME + '\\' + FILE_NAME        


shutil.unpack_archive(FILE_NAME + '.docx', d_name, 'zip')

inFile = open(d_name + '\\word\\settings.xml', 'r')
data = inFile.readlines()
inFile.close()
start = data[1].find('<w:documentProtection')
end = data[1].find('/>', start)
data[1] = data[1].replace(data[1][start:end+2], '')
outFile = open(d_name + '\\word\\settings.xml', 'w')
outFile.writelines(data)
outFile.close()


shutil.make_archive(f_name, 'zip', d_name)

os.rename(f_name + '.zip', f_name)      
shutil.rmtree(d_name)  
