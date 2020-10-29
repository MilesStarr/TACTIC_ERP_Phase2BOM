# -*- coding: utf-8 -*-
"""
Created on Thu Oct 22 13:24:55 2020

@author: eclark
"""

from lxml import etree
import re

def check(testPN, partSet):
    testPN = ".*{}.*".format(testPN)
    hits = []
    for part in partSet:
        if re.match(testPN, part):
            hits.append(part)
    return hits

#add new files here to validate that parts are not already defined in the master
files = ["M265_266.xml"]

masterTree = etree.parse(".TacTicMasterBOM.xml")
masterRoot = masterTree.getroot()

for file in files:
    thisTree = etree.parse(file)
    theseParts = thisTree.getroot().findall(".//Part")
    for thisPart in theseParts:
        if masterRoot.find(".//Part[@PartID='{0}']".format(thisPart.attrib['PartID'])) is None:
            masterRoot.append(thisPart)
        else:
            print("Crashed on part {} in {}".format(thisPart.attrib['PartID'], file))

allParts = masterRoot.findall(".//Part")
partDefs = set()
for part in allParts:
    partDefs.add(part.attrib['PartID'])

#masterTree.write(".TacTicMasterBOM.xml") #uncomment to overwrite master BOM xml file