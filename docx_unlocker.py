# -*- coding: utf-8 -*-
"""
Created on Wed Oct 24 10:41:48 2018

@author: moskalev_iv
"""

import os
import shutil
import argparse

if os.name == 'nt':
    FD = '\\'
else:
    FD = '/'

def unlock(file_name : str):
    f_name = os.path.dirname(file_name) + FD + '@' + os.path.basename(file_name)
    shutil.unpack_archive(file_name, file_name[:-5], 'zip')
    inFile = open(file_name[:-5] + FD + 'word' + FD +  'settings.xml', 'r')
    data = inFile.readlines()
    inFile.close()
    start = data[1].find('<w:documentProtection')
    if start > 0:
        end = data[1].find('/>', start)
        data[1] = data[1].replace(data[1][start:end+2], '')
        outFile = open(file_name[:-5] + FD + 'word' + FD + 'settings.xml', 'w')
        outFile.writelines(data)
        outFile.close()
        shutil.make_archive(f_name, 'zip', file_name[:-5])
        os.rename(f_name + '.zip', f_name)  
        print("{} unlocked".format(file_name))
    else:
        print("File \"{}\" isn't locked".format(file_name))
    shutil.rmtree(file_name[:-5])  

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
            print(os.path.abspath(args.path))
            for root, _, filenames in os.walk(os.path.abspath(args.path)):
                files.extend([root + FD + file for file in filenames if ".docx" in file])
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
                for root, _, filenames in os.walk(os.path.abspath(args.path)):
                    files.extend([root + FD + file for file in filenames if ".docx" in file])
                    break
                if files != []:
                    for f in files:
                        unlock(f)
                else:
                    print("Do not exist suitable files")
        else:
            unlock(os.path.abspath(args.path))

if __name__ == "__main__":
    main()
