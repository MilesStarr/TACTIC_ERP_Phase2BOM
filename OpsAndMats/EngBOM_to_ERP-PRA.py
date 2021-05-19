# -*- coding: utf-8 -*-
"""
Created on Fri Aug 28 12:46:12 2020

@author: eclark
"""

import os, sys

file_dir = os.path.dirname(__file__)
sys.path.append(file_dir)

from TACTIC_ERP import TACTIC_ERP as BOM
import pandas as pd
from lxml import etree

name = "8027003_HD Safety Whl Ass'y"

tree = etree.parse("../" + name + ".xml")

Operations = pd.DataFrame()
Materials = pd.DataFrame()

engBOM = BOM(tree)

buildMe = tree.getroot().findall(".//Part")
for buildItem in buildMe:
    (ops, mats) = engBOM.buildItem(buildItem.attrib["PartID"])
    Operations = Operations.append(ops, ignore_index=True)
    Materials = Materials.append(mats, ignore_index=True)
engBOM.itemRecommendations.to_excel(name + "_recs.xlsx")

OpColumns = ["Item", "Item Description", "Alternate ID", "Description", "Operation", "Shared", "WC", "WC Description", "Use Fixed Schedule", "Fixed Sched Hours", "Run-Hours Basis (Machine)", "Mach Hrs per Piece", "Run-Hours Basis (Labor)", "Labor Hr per Piece", "Sched Driver", "Run Duration", "Batch Definition", "Yield", "Seconds Per Cycle", "Formula Material Weight", "Formula Material Weight U/M", "Move Hours", "Queue Time", "Setup Hours", "Finish", "Use Offset Hours", "Offset Hours", "Effective Date", "Obsolete Date", "Control Point", "Backflush", "Setup Resource Group", "Setup Rule", "Setup Basis", "Setup Time Rule", "Setup Matrix", "Scheduler Rule", "Custom Planner Rule", "Break Rule", "Split Rule", "Split Size", "Split Group", "Efficiency", "Setup Rate", "Run Rate (Labor)", "Var Mach Ovhd Rate", "Fix Machine Ovhd Rate", "Var Ovhd Rate", "Fixed Ovhd Rate"]
MatColumns = ["Item", "Item Description", "Alternate ID", "Description", "Operation", "Shared", "WC", "WC Description", "Material", "Material Description", "Seq", "Alt Group", "Alt Group Rank", "Manufacturer", "Manufacturer Name", "Manufacturer Item", "Manufacturer Item Description", "Type", "Quantity", "Per", "U/M", "Cost", "Scrap Factor", "Effective Date", "Obsolete Date", "BOM Seq", "Ref", "Backflush", "Backflush Location", "Feature", "Option Code", "Probable", "Incremental Price", "Formula Material Weight %", "Estimated Break Date", "Date of Last Report", "Fixed Material", "Variable Material", "Material Cost", "Labor Cost", "Outside Cost", "Fixed Overhead Cost", "Variable Overhead Cost"]

with pd.ExcelWriter(name + ".xlsx") as outFile:
    Materials.to_excel(outFile, sheet_name = "Materials", columns = MatColumns, index=False)
    Operations.to_excel(outFile, sheet_name = "Operations", columns = OpColumns, index=False)