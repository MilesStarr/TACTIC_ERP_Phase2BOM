# -*- coding: utf-8 -*-
"""
Created on Wed Oct 21 09:09:43 2020

@author: eclark
"""

import subprocess

xmlFiles = subprocess.run(["git", "ls-files", "*.xml"], capture_output=True).stdout.decode().splitlines()
fileLog = {}
for file in xmlFiles:
    fileLog[file.split("/")[-1].split(".")[0]] = subprocess.run(["git", "log", "-n 1", '--format="%H"', file], capture_output=True).stdout.decode().strip().strip('"')
del file

excelFiles = subprocess.run(["git", "ls-files", r"OpsAndMats\*.xls*"], capture_output=True).stdout.decode().splitlines()
resultLog = {}
for file in excelFiles:
    resultLog[file.split("/")[-1].split(".")[0]] = subprocess.run(["git", "log", "-n 1", '--format="%H"', file], capture_output=True).stdout.decode().strip().strip('"')
del file

commitOrder = subprocess.run(["git", "log", '--format="%H"'], capture_output=True).stdout.decode().splitlines()
for i in range(len(commitOrder)):
    commitOrder[i] = commitOrder[i].strip('"')
del i

for xml, value in fileLog.items():
    if xml in resultLog.keys():
        xmlHashIndex = commitOrder.index(value)
        excelHashIndex = commitOrder.index(resultLog[xml])
        if xmlHashIndex < excelHashIndex:
            print("{} is dirty".format(xml))
            print("TortoiseGitProc /command:diff /path {file}.xml /startrev:{excelHash} /endrev:{xmlHash}".format(file=xml, excelHash=resultLog[xml], xmlHash=value))
            print()
    elif xml=="": #catch the ".TacTicMasterBOM.xml" file as not to be displayed
        pass
    else:
        print("{} has never been processed".format(xml))
        print()
del xml, value, xmlHashIndex, excelHashIndex

input("press enter to terminate...")