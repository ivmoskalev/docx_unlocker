# -*- coding: utf-8 -*-
"""
Created on Wed Oct 24 10:41:48 2018

@author: moskalev_iv
"""

import os
import shutil
import argparse
import glob


def unlock(file_name : str):
    print("{} unlocked".format(file_name))
# FILE_NAME = '1521.122.06ДО Ч.3 - ВМ_v00.100155186390074'
# DIR_NAME = os.path.dirname(os.path.abspath(FILE_NAME + '.docx'))
# f_name = DIR_NAME + '\\@' + FILE_NAME + '.docx'
# d_name = DIR_NAME + '\\' + FILE_NAME        


# shutil.unpack_archive(FILE_NAME + '.docx', d_name, 'zip')

# inFile = open(d_name + '\\word\\settings.xml', 'r')
# data = inFile.readlines()
# inFile.close()
# start = data[1].find('<w:documentProtection')
# end = data[1].find('/>', start)
# data[1] = data[1].replace(data[1][start:end+2], '')
# outFile = open(d_name + '\\word\\settings.xml', 'w')
# outFile.writelines(data)
# outFile.close()


# shutil.make_archive(f_name, 'zip', d_name)

# os.rename(f_name + '.zip', f_name)      
# shutil.rmtree(d_name)  


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("path", help="path to docx")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="recursive unlocked docx in path")
    args = parser.parse_args()
    if args.recursive:
        user_answer = input("Are you want to unlock all files in \"{}\" recursively? [y/n] ".format(os.path.abspath(args.path))) 
        if str.lower(user_answer) == "y":
            files = []
            for _, _, filenames in os.walk(os.path.abspath(args.path)):
                files.extend([os.path.abspath(x) for x in filenames if ".json"in x])
            if files != []:
                for f in files:
                    unlock(f)
            else:
                print("Do not exist suitable files")
    else:
        if os.path.isdir(args.path):
            user_answer = input("Are you want to unlock all files in \"{}\"? [y/n] ".format(os.path.abspath(args.path))) 
            if str.lower(user_answer) == "y":
                files = []
                for _, _, filenames in os.walk(os.path.abspath(args.path)):
                    files.extend([os.path.abspath(x) for x in filenames if ".json"in x])
                    break
                if files != []:
                    for f in files:
                        unlock(f)
                else:
                    print("Do not exist suitable files")

        else:
            unlock(args.path)

if __name__ == "__main__":
    main()