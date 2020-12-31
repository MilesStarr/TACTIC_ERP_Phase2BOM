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
files = ["8025043-1.xml", "8025044.xml", "8025737-X.xml", "8026007_Tank Mtg Rail.xml", "8027003_HD Safety Whl Ass'y.xml", "8027098.xml", "8033334_HD Drive Unit.xml", "8034211_ValveStation.xml", "8034919.xml", "8035422-X_Control Box.xml", "C27876-X_HD Prec Angle Adj.xml", "HD_Drive_Stops.xml", "HD_Storage_Racks.xml", "Hold down ass'y_8028444-X.xml", "M105-X.xml", "M140.xml", "M188.xml", "M204.xml", "M219.xml", "M258A.xml", "M259A_260A.xml", "M265_266.xml", "M33H.xml", "M411.xml", "M45-9.xml", "M600H.xml", "M900.xml", "Model148.xml", "Model28.xml", "Model33M.xml", "Model40.xml", "Model76E.xml", "Model76E_Motor.xml", "Model76FG.xml", "Model76FG_Motor.xml", "Model83.xml", "Prewetter Ass'y_8027220.xml", "Safety Whl Ass'y_8026015-X.xml", "ShaftGuards_8026139.xml", "Tanks.xml"]

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

# Master Build Order:
# files = ["8025043-1.xml", "8025044.xml", "8025737-X.xml", "8026007_Tank Mtg Rail.xml", "8027003_HD Safety Whl Ass'y.xml", "8027098.xml", "8033334_HD Drive Unit.xml", "8034919.xml", "C27876-X_HD Prec Angle Adj.xml", "HD_Drive_Stops.xml", "Hold down ass'y_8028444-X.xml", "M105-X.xml", "M258A.xml", "M259A_260A.xml", "M265_266.xml", "M33H.xml", "M411.xml", "M45-9.xml", "M900.xml", "Model148.xml", "Model28.xml", "Model33M.xml", "Model40.xml", "Model76E.xml", "Model76E_Motor.xml", "Model76FG.xml", "Model83.xml", "Prewetter Ass'y_8027220.xml", "Safety Whl Ass'y_8026015-X.xml", "ShaftGuards_8026139.xml", "Tanks.xml", "M600H.xml", "8034211_ValveStation.xml", "HD_Storage_Racks.xml", "M204.xml"]
# files = ["M188.xml", "M140.xml"]



