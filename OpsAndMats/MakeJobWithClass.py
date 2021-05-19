# -*- coding: utf-8 -*-
"""
Created on Fri Sep 18 08:55:29 2020

@author: eclark
"""

import pandas as pd
from lxml import etree
import copy

from TACTIC_ERP import TACTIC_ERP

pd.options.mode.chained_assignment = None  # default='warn'

def makeThisJob(job = "J000000000", suffix = "0000"):
    # Main section
    temp = TACTIC_ERP()
    JobInformation = temp.requestJobOrder(job,suffix)
    
    JobInformation.update({"Prepared": "HH", "Checked": " ", "Appr": " "})
    
    #Build dataset from ERP report querry
    #Preload first tier of materials data
    debug_requestsMade = [[JobInformation['Item']]]
    matListPD = temp.requestMats(JobInformation['Item'])
    matListWorking = copy.deepcopy(matListPD)
    
    print("fetched main")
    notDone = True
    counter = 0
    neededHistory = []
    #Fetch lower tiers of BOMs
    while notDone and (counter < 30): #Infor ERP has a tier limit of 20, so don't try too hard
        counter += 1
        matListFilter = matListWorking.loc[:,'matl_product_code'] == "TT-MfgStep"
        neededMaterials = list(set(matListWorking.loc[matListFilter]['material']))
        notDone = len(neededMaterials) > 0
        neededHistory.append(neededMaterials)
        if notDone:
            debug_requestsMade.append(neededMaterials)
            matListWorking = temp.requestMats(neededMaterials)
            matListPD = matListPD.append(matListWorking, ignore_index=True)
            print("fetched {} down, {} items".format(counter, len(neededMaterials)))
    #Fetch Operations listing for all items identified in materials data
    #This assumes it is imposible to make something from nothing (at least one
    #input material must exist in a BOM)
    neededOperations = list(set(matListPD['item']))
    opListPD = temp.requestOPs(neededOperations)
    print("fetched operations")
    
    #I ain't afraid of no ghosts
    matListPD = temp.resolvePhantoms(matListPD)
    print("I ain't afraid of no ghosts")
    
    #non-recursive build out of operations with dependancies
    (matList, refList, nextOP, nextMat) = temp.buildOps(opListPD, matListPD, Part = JobInformation['Item'])
    nextDepth = 1
    while len(refList) > 0:
        newRefList = []
        for ref in refList:
            (nextMatList, nextRefList, nextOP, nextMat) = temp.buildOps(opListPD, matListPD, 
                                                                   Part = ref['material'],
                                                                   ParentQTY = ref['ParentQTY'],
                                                                   ParentOP = ref['opID'],
                                                                   opID = nextOP,
                                                                   matID = nextMat,
                                                                   depth = nextDepth)
            matList.extend(nextMatList)
            newRefList.extend(nextRefList)
        refList = newRefList
        nextDepth += 1
    print("Made my list")
    
    # ERP Table Information for Paste-Rows-Append
    JobOpColumns = ["Job", "Suffix", "Operation", "Shared", "WC", "WC Description", "Complete", "Use Fixed Schedule", "Fixed Sched Hours", "Run-Hours Basis (Machine)", "Run Machine Hours", "Run-Hours Basis (Labor)", "Run Lbr Hours", "Sched Driver", "Run Duration", "Batch Definition", "Batch ID", "Yield", "Seconds Per Cycle", "Formula Material Weight", "Formula Material Weight U/M", "Move Hours", "Queue Hours", "Setup Hours", "Finish Hours", "Use Offset Hours", "Offset Hours", "Start Date", "End Date", "Control Point", "Backflush", "Setup Resource Group", "Setup Rule", "Setup Basis", "Setup Time Rule", "Setup Matrix", "Scheduler Rule", "Custom Planner Rule", "Break Rule", "Split Rule", "Split Size", "Split Group", "Efficiency", "Setup Rate", "Labor Run Rate", "Var Mach Ovhd Rate", "Fix Machine Ovhd Rate", "Var Ovhd Rate", "Fixed Ovhd Rate", "Received", "Received U/M", "Completed", "Completed U/M", "Moved", "Moved U/M", "Scrapped", "Scrapped U/M", "U/M", "Total Setup Hours", "Total Run Hours (Labor)", "Total Run Hours (Machine)", "Total Setup Cost", "Total Fixed Ovhd (Machine)", "Total Var Ovhd (Machine)", "Total Run Cost (Labor)", "Total Fixed Ovhd (Labor)", "Total Var Ovhd (Labor)", "Operation Type", "Operation Type Description", "Operation Type Code", "Operation Type Code Description", "Sheet Count", "Original Colors", "Special Colors", "Additional Colors", "Operation Price", "Operation Difficulty Factor", "Up", "Out", "Length", "Width", "Setup Finish Run Option 1", "Setup Finish Run Option 2", "Setup Finish Run Option 3", "Setup Finish Run Option 4", "Setup Finish Run Option 5", "Setup Finish Run Option 6", "Setup Finish Run Option 7", "Setup Finish Run Option 8", "Setup Finish Run Option 9", "Setup Finish Run Option 10", "Setup Finish Run Option 11", "Setup Finish Run Option 12", "Setup Finish Run Option 13", "Setup Finish Run Option 14", "Setup Finish Run Option 15", "Setup Finish Run Option 16", "Setup Finish Run Option 17", "Setup Finish Run Option 18", "Setup Finish Run Option 19", "Setup Finish Run Option 20", "Paper Consumption Qty", "Job Qty Per Sheet", "Number of Manual Handling Steps", "Number of Sides to Print", "Operation Rating 1", "Operation Rating 2", "Operation Rating 3", "Operation Rating 4", "Operation Rating 5", "Operation Rating 6", "Operation Rating 7", "Operation Rating 8", "Operation Rating 9", "Operation Rating 10", "Operation Rating 11", "Operation Rating 12", "Operation Rating 13", "Operation Rating 14", "Operation Rating 15", "Operation Rating 16", "Operation Rating 17", "Operation Rating 18", "Operation Rating 19", "Operation Rating 20", "Suffix", "Item", "Status", "Released", "Item Description"]
    JobMatColumns = ["Job", "Suffix", "Item", "Description", "Status", "Released", "Operation", "WC", "WC Description", "Material", "Material Description", "Seq", "Alternate", "Manufacturer", "Manufacturer Name", "Manufacturer Item", "Manufacturer Item Description", "Type", "Kit", "Qty", "Per", "U/M", "Cost", "Scrap Factor", "BOM Seq", "Backflush", "Backflush Location", "Source Type", "Source", "Source Line-Suffix", "Source Release", "Fixed Matl Ovhd Rate", "Variable Matl Ovhd Rate", "Material Cost", "Labor Cost", "Fix Ovhd Cost", "Var Ovhd Cost", "Outside Cost", "Qty Issued", "Total Fixed Overhead", "Total Variable Overhead", "Last Pick List",  "Actual Material Cost", "Actual Labor Cost", "Actual Fix Ovhd Cost", "Actual Var Ovhd Cost", "Actual Outside Cost", "Actual Cost Total", "Shared", "Formula Material Weight %", "Material is Paper"]
    
    #Filling in the gaps
    JobOpTemplate = {'Shared': "0", 'Complete': "0", 'Use Fixed Schedule': "0", 'Fixed Sched Hours': "0",
                     'Run-Hours Basis (Machine)': "Mch Hours/Pc", 'Run-Hours Basis (Labor)': "Lbr Hours/Pc",
                     'Sched Driver': "Machine", 'Batch Definition': "", 'Batch ID': "", 'Yield': "100",
                     'Seconds Per Cycle': "0", 'Formula Material Weight': "0", 'Formula Material Weight U/M': "",
                     'Start Date': "", 'End Date': "", 'Control Point': "1", 'Backflush': "Neither", 
                     'Setup Resource Group': "", 'Setup Rule': "Always", 'Setup Basis': "Item",
                     'Setup Time Rule': "Fixed Time", 'Setup Matrix': "", 'Scheduler Rule': "Per Piece",
                     'Custom Planner Rule': "0", 'Break Rule': "Shifts", 'Split Rule': "No Splitting",
                     'Split Size': "0", 'Split Group': "", 'Efficiency': "100", 'Received': "0",
                     'Total Setup Hours': "0", 'Total Run Hours (Labor)': "0", 'Total Run Hours (Machine)': "0",
                     'Total Setup Cost': "0", 'Total Fixed Ovhd (Machine)': "0", 'Total Var Ovhd (Machine)': "0",
                     'Total Run Cost (Labor)': "0", 'Total Fixed Ovhd (Labor)': "0", 'Total Var Ovhd (Labor)': "0",
                     'Operation Type': "", 'Operation Type Description': "", 'Operation Type Code': "",
                     'Operation Type Code Description': "", 'Sheet Count': "0", 'Original Colors': "",
                     'Special Colors': "", 'Additional Colors': "", 'Operation Price': "0",
                     'Operation Difficulty Factor': "0", 'Up': "", 'Out': "", 'Length': "", 'Width': "",
                     'Setup Finish Run Option 1': "", 'Setup Finish Run Option 2': "", 'Setup Finish Run Option 3': "",
                     'Setup Finish Run Option 4': "", 'Setup Finish Run Option 5': "", 'Setup Finish Run Option 6': "",
                     'Setup Finish Run Option 7': "", 'Setup Finish Run Option 8': "", 'Setup Finish Run Option 9': "",
                     'Setup Finish Run Option 10': "", 'Setup Finish Run Option 11': "", 'Setup Finish Run Option 12': "",
                     'Setup Finish Run Option 13': "", 'Setup Finish Run Option 14': "", 'Setup Finish Run Option 15': "",
                     'Setup Finish Run Option 16': "", 'Setup Finish Run Option 17': "", 'Setup Finish Run Option 18': "",
                     'Setup Finish Run Option 19': "", 'Setup Finish Run Option 20': "", 'Paper Consumption Qty': "0", 'Job Qty Per Sheet': "0",
                     'Number of Manual Handling Steps': "", 'Number of Sides to Print': "", 'Operation Rating 1': "0",
                     'Operation Rating 2': "0", 'Operation Rating 3': "0", 'Operation Rating 4': "0",
                     'Operation Rating 5': "0", 'Operation Rating 6': "0", 'Operation Rating 7': "0",
                     'Operation Rating 8': "0", 'Operation Rating 9': "0", 'Operation Rating 10': "0",
                     'Operation Rating 11': "0", 'Operation Rating 12': "0", 'Operation Rating 13': "0",
                     'Operation Rating 14': "0", 'Operation Rating 15': "0", 'Operation Rating 16': "0",
                     'Operation Rating 17': "0", 'Operation Rating 18': "0", 'Operation Rating 19': "0",
                     'Operation Rating 20': "0", 'Status': "Firm", 'Item Description': "",
                     'Setup Rate': "", 'Labor Run Rate': "", 'Var Mach Ovhd Rate': "", 'Fix Machine Ovhd Rate': "",
                     'Var Ovhd Rate': "", 'Fixed Ovhd Rate': "", 'Received U/M': "", 'Completed U/M': "", 'Moved U/M': "",
                     'Scrapped U/M': "", 'U/M': ""}
    JobMatTemplate = {'Description': "", 'Status': "Firm", 'WC Description': "",
                      'Alternate': "0", 'Manufacturer': "", 'Manufacturer Name': "",
                      'Manufacturer Item': "", 'Manufacturer Item Description': "", 'Type': "Material",
                      'Kit': "0", 'Per': "Unit", 'Cost': "", 'Scrap Factor': "0",
                      'BOM Seq': "", 'Backflush': "1", 'Backflush Location': "", 'Source': "",
                      'Source Line-Suffix': "0", 'Source Release': "0", 'Fixed Matl Ovhd Rate': "0",
                      'Variable Matl Ovhd Rate': "0", 'Material Cost': "0", 'Labor Cost': "0", 'Fix Ovhd Cost': "0",
                      'Var Ovhd Cost': "0", 'Outside Cost': "0", 'Qty Issued': "0", 'Total Fixed Overhead': "0",
                      'Total Variable Overhead': "0", 'Last Pick List': "", 'Actual Material Cost': "0", 
                      'Actual Labor Cost': "0", 'Actual Fix Ovhd Cost': "0", 'Actual Var Ovhd Cost': "0",
                      'Actual Outside Cost': "0", 'Actual Cost Total': "0", 'Shared': "0", 
                      'Formula Material Weight %': "0", 'Material is Paper': "0"}
    
    #Process matList into sequential steps
    
    order_dTypes = {'MST-run': "int", 'Major': "int", 'Milestone': "int", 'Minor': "int", 'ParentOP': "int", 'ParentQTY': "int",
                    'depth': "int", 'item': "str", 'item_description': "str", 'item_drawing_nbr': "str",
                    'item_p_m_t_code': "str", 'item_phantom_flag': "int", 'item_product_code': "str",
                    'item_revision': "str", 'item_u_m': "str", 'leaf': "int", 'matID': "int", 'material': "str",
                    'matl_description': "str", 'matl_drawing_nbr': "str", 'matl_p_m_t_code': "str",
                    'matl_phantom_flag': "int", 'matl_product_code': "str", 'matl_qty_conv': "float",
                    'matl_revision': "str", 'matl_u_m': "str", 'opID': "int", 'oper_num': "int", 'run_lbr_hrs': "float",
                    'run_mch_hrs': "float", 'sched_drv': "str", 'wc': "str", 'wc_description': "str"}
    
    ERP_JobOP = pd.DataFrame()
    ERP_JobMat = pd.DataFrame()
    matListCopy = pd.DataFrame(matList)
    jobOrderDetails = pd.DataFrame()
    
    delayBackflushList = ["TT-Raw-HV", "TT-NonSale", "TT-Spare", "TT-Assy"]
    
    loopLimit = len(matListCopy) + 5
    loopIter = 0
    majorOP = [10]
    minorOP = 10
    milestones = [10, 20]
    while len(matListCopy) > 0 and loopIter < loopLimit:
        bumpMilestone = False
        # iterate over all process listings, identifying leaf nodes
        for matIndex in matListCopy.index:
            matListCopy.loc[matIndex, 'leaf'] = matListCopy.loc[matIndex, 'opID'] not in matListCopy['ParentOP'].values
    
        # isolate the leaf nodes and sort them into priority levels 
        # currently not accurate, but something.
        #future improvement: sort by least slack time by aggregating operation duration
        leafs = matListCopy.loc[matListCopy['leaf'] == True, :].sort_values(['material', 'item', 'depth', 'opID'], ascending = [False, True, False, False])
        
        # iterate through leafs to generate shop information, removing from material list and leaf list as they are consumed
        while len(leafs) > 0:
            # start at the top of the list and pick all listings that are the same operation to process them together
            firstLeaf = leafs.iloc[0]
            thisOperation = leafs.loc[leafs['opID'] == firstLeaf['opID'], :]
            # put information from the listing into the shop order details
            for matIndex, material in thisOperation.iterrows():
                bumpMilestone = bumpMilestone or material['matl_product_code'] in delayBackflushList
                milestone = milestones[-2] if material['matl_product_code'] not in delayBackflushList else milestones[-1]
                material = material.append(pd.Series([milestone, milestones[-2], majorOP[-1], minorOP], index=['Milestone', 'MST-run', 'Major', 'Minor']))
                jobOrderDetails = jobOrderDetails.append(material, ignore_index=True)
                leafs.drop(matIndex, inplace = True)
                matListCopy.drop(matIndex, inplace = True)
            minorOP += 10
        
        majorOP.append(majorOP[-1] + 10)
        minorOP = 10
        if bumpMilestone:
            milestones.append(milestones[-1] + 10)
        loopIter += 1  # failsafe iteration cap
    #undo last milestone bump if happened on last iteration
    if bumpMilestone:
        milestones.pop()
    jobOrderDetails = jobOrderDetails.astype(order_dTypes)
    print("Checked it twice")
    
    ERP_JobOP = pd.DataFrame()
    ERP_JobMat = pd.DataFrame()
    
    #populate Job Operations
    jobWC = "TACwip"
    for milestone in milestones:
        tempOp = {'Job': JobInformation['Job'],
                  'Suffix': JobInformation['Suffix'],
                  'Operation': milestone,
                  'WC': jobWC,
                  'WC Description': "",
                  'Run Machine Hours': "0", 'Run Lbr Hours': "0",
                  'Run Duration': "0",
                  'Move Hours': "0", 'Queue Hours': "0", 'Setup Hours': "0",
                  'Finish Hours': "0", 'Use Offset Hours': "0", 'Offset Hours': "0",
                  'Completed': "0", 'Moved': "0", 'Scrapped': "0",
                  'Item': JobInformation['Item'], 'Released': JobInformation['Released'],}
        tempOp.update(JobOpTemplate)
        ERP_JobOP = ERP_JobOP.append(tempOp, ignore_index=True)
        jobWC = "TACrun"
    
    #populate Job Materials
    seq = {m: 1 for m in milestones}
    for matIndex, material in jobOrderDetails.query("material != '' & matl_product_code != 'TT-MfgStep'").iterrows():
        tempMat = {'Job': JobInformation['Job'], 'Suffix': JobInformation['Suffix'],
                   'Item': JobInformation['Item'],
                   'Released': JobInformation['Released'],
                   'Operation': material['Milestone'],
                   'WC': "TACwip" if material['Milestone'] == milestones[0] else "TACrun",
                   'Material': material['material'],
                   'Material Description': material['matl_description'],
                   'Seq': seq[material['Milestone']],
                   'U/M': material['matl_u_m'],
                   'Qty': material['matl_qty_conv'], 
                   'Source Type': "Inventory"}
        tempMat.update(JobMatTemplate)
        ERP_JobMat = ERP_JobMat.append(tempMat, ignore_index = True)
        seq[material['Milestone']] += 1
    
    with pd.ExcelWriter("{}_{}.xlsx".format(JobInformation['Job'],JobInformation['Suffix'])) as outFile:
        ERP_JobOP.to_excel(outFile, sheet_name = "Operations", columns = JobOpColumns, index=False)
        ERP_JobMat.to_excel(outFile, sheet_name = "Materials", columns = JobMatColumns, index=False)
    print("ERP dataset is finished")    
    
    #generate Printable Shop Order
    
    #Expand list for each part of size the number of times listed in released and increase the quantityies for aggragated parts
    aggrUnits = ["EA", "PT", ""]
    sizeUnits = ["FT", "IN", "LB", "ml", "P", "SF"]
    shopOrderDetails = jobOrderDetails[[x in aggrUnits for x in list(jobOrderDetails['matl_u_m'])]]
    shopOrderDetails.loc[:, 'ParentQTY'] = shopOrderDetails['ParentQTY'] * int(JobInformation["Released"])
    shopOrderDetails.loc[:, 'matl_qty_conv'] = shopOrderDetails['matl_qty_conv'] * int(JobInformation["Released"])
    shopOrderDetails = shopOrderDetails.append([jobOrderDetails[[x not in aggrUnits for x in list(jobOrderDetails['matl_u_m'])]]] * int(JobInformation["Released"]), ignore_index=True)
    shopOrderDetails = shopOrderDetails.astype(order_dTypes)
    
    # Building a shop order HTML File
    shopOrderFile = etree.Element("HTML")
    # Header
    headerHTML = etree.Element("HEAD")
    
    etree.SubElement(headerHTML, "TITLE")
    headerHTML[-1].text = "Shop Order"
    shopOrderFile.append(headerHTML)
    
    bodyHTML = etree.Element("BODY")
    etree.SubElement(bodyHTML, "H1")
    bodyHTML[-1].text = "TAC Technical Instrument Corp."
    
    etree.SubElement(bodyHTML, "H2")
    bodyHTML[-1].text = "Shop Order {}".format(JobInformation['Job'])
    
    soHeader = etree.SubElement(bodyHTML, "TABLE", border="1", width="85%")
    soHeaderRow = etree.SubElement(soHeader, "TR")
    etree.SubElement(soHeaderRow, "TH", width = "28%")
    soHeaderRow[-1].text = "Product"
    etree.SubElement(soHeaderRow, "TH", width = "12%")
    soHeaderRow[-1].text = "Start_Date"
    etree.SubElement(soHeaderRow, "TH", width = "12%")
    soHeaderRow[-1].text = "Complete by Date"
    etree.SubElement(soHeaderRow, "TH", width = "12%")
    soHeaderRow[-1].text = "Prepared By"
    etree.SubElement(soHeaderRow, "TH", width = "12%")
    soHeaderRow[-1].text = "Checked By"
    etree.SubElement(soHeaderRow, "TH", width = "12%")
    soHeaderRow[-1].text = "Approved By"
    etree.SubElement(soHeaderRow, "TH", width = "12%")
    soHeaderRow[-1].text = "Build Qty."
    
    soHeaderRow = etree.SubElement(soHeader, "TR")
    etree.SubElement(soHeaderRow, "TD")
    soHeaderRow[-1].text = JobInformation['Item']
    etree.SubElement(soHeaderRow, "TD")
    soHeaderRow[-1].text = JobInformation["Start_Date"]
    etree.SubElement(soHeaderRow, "TD")
    soHeaderRow[-1].text = JobInformation["End_Date"]
    etree.SubElement(soHeaderRow, "TD")
    soHeaderRow[-1].text = JobInformation["Prepared"]
    etree.SubElement(soHeaderRow, "TD")
    soHeaderRow[-1].text = JobInformation["Checked"]
    etree.SubElement(soHeaderRow, "TD")
    soHeaderRow[-1].text = JobInformation["Appr"]
    etree.SubElement(soHeaderRow, "TD")
    soHeaderRow[-1].text = str(JobInformation["Released"])
    
    etree.SubElement(bodyHTML, "H2")
    bodyHTML[-1].text = "Operations"
    
    
    
    for milestone in milestones:
        etree.SubElement(bodyHTML, "H3")
        bodyHTML[-1].text = "ERP Milestone {}".format(milestone)
        
        tempSOdata = shopOrderDetails.query("`MST-run` == {}".format(milestone))
        tempSOdata = tempSOdata.sort_values(['Major', 'Minor'], ascending = True)
        
        while len(tempSOdata) > 0:
            thisOP = tempSOdata.loc[(tempSOdata['Major'] == tempSOdata.iloc[0]['Major']) & (tempSOdata['Minor'] == tempSOdata.iloc[0]['Minor'])].sort_values('material')
            soOPs = etree.SubElement(bodyHTML, "TABLE", style="border-style: double; table-layout:fixed;", width="100%")
            soOPcolDef = etree.SubElement(soOPs, "COLGROUP")
            etree.SubElement(soOPcolDef, "COL", width="15%")
            etree.SubElement(soOPcolDef, "COL")
            etree.SubElement(soOPcolDef, "COL", width="10%")
            etree.SubElement(soOPcolDef, "COL", width="10%")
            etree.SubElement(soOPcolDef, "COL", width="10%")
            soOPhead = etree.SubElement(soOPs, "THEAD")
            soOProw = etree.SubElement(soOPhead, "TR")
            soOPheaderCell = etree.SubElement(soOProw, "TH", colspan="5", style="border-style: none none solid none;")
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:15%;")
            soOPheaderCell[-1].text = "{}-{}".format(thisOP.iloc[0]['Major'], thisOP.iloc[0]['Minor'])
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:45%;")
            soOPheaderCell[-1].text = "{}: {}".format(thisOP.iloc[0]['wc'],thisOP.iloc[0]['wc_description'])
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:20%;")
            soOPheaderCell[-1].text = "{} hr.".format(max(thisOP.iloc[0]['run_lbr_hrs'], thisOP.iloc[0]['run_mch_hrs'])*int(JobInformation["Released"]))
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:100%;")
            soOPheaderCell[-1].text = "{}".format(thisOP.iloc[0]['item_tt_description'])
            soOProw = etree.SubElement(soOPhead, "TR")
            etree.SubElement(soOProw, "TH")
            soOProw[-1].text = "Material #"
            etree.SubElement(soOProw, "TH")
            soOProw[-1].text = "Description"
            etree.SubElement(soOProw, "TH")
            soOProw[-1].text = "Ref."
            etree.SubElement(soOProw, "TH")
            soOProw[-1].text = "Qty."
            etree.SubElement(soOProw, "TH")
            soOProw[-1].text = "Source"
            soOPbody = etree.SubElement(soOPs, "TBODY")
            for index, row in thisOP.iterrows():
                matRow = etree.SubElement(soOPbody, "TR")
                etree.SubElement(matRow, "TD")
                if row['material'] != "":
                    matRow[-1].text = "{}".format(row["material"])
                else:
                    matRow[-1].text = "{}".format("N/A")
                etree.SubElement(matRow, "TD")
                if row['material'] != "":
                    matRow[-1].text = "{}".format(row.loc["matl_description"])
                else:
                    matRow[-1].text = "WIP - {}".format(thisOP.iloc[0]['item'])
                etree.SubElement(matRow, "TD")
                if row.loc['matl_drawing_nbr'] != "":
                    matRow[-1].text = "{} Rev. {}".format(row['matl_drawing_nbr'],row['matl_revision'])
                etree.SubElement(matRow, "TD")
                
                if row['matl_u_m'] in ["IN", "SF"]: #Units of measure that should display fractionally
                    matRow[-1].text = temp.fraction(value=row['matl_qty_conv'], maxDen = 16, unit=row['matl_u_m'])
                elif row['matl_u_m']  in ["EA", "P", ""]: #Units that should only contain integers
                    matRow[-1].text = "{:d} {}".format(int(row.loc["matl_qty_conv"]), row.loc["matl_u_m"])
                else:
                    matRow[-1].text = "{} {}".format(row.loc["matl_qty_conv"], row.loc["matl_u_m"])
                
                etree.SubElement(matRow, "TD")
                if row['matl_product_code'] == "TT-MfgStep":
    #                rowParent = shopOrderDetails.query('ParentOP == {0:d} and item == {1}'.format(row['opID'],row['material'])).iloc[0]
                    rowParent = shopOrderDetails[(shopOrderDetails['ParentOP'] == row['opID']) & (shopOrderDetails['item'] == row['material'])].iloc[0]
                    source = "{:d}-{:d}".format(rowParent['Major'],rowParent['Minor'])
                elif row['material'] == "":
                    rowParent = shopOrderDetails[(shopOrderDetails['ParentOP'] == row['opID'])].iloc[0]
                    source = "{:d}-{:d}".format(rowParent['Major'],rowParent['Minor'])
                elif row['matl_p_m_t_code'] == "P":
                    source = "Inventory"
                else:
                    source = "Job #: ________"
                matRow[-1].text = source
            lastRow = etree.SubElement(soOPbody, "TR")
            soOPheaderCell = etree.SubElement(lastRow, "TH", colspan="5")
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:25%;")
            soOPheaderCell[-1].text = "{}".format(thisOP.iloc[0]['item'])
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:10%; background-color:lightgrey; color:white")
            soOPheaderCell[-1].text = "{}".format(thisOP.iloc[0]['oper_num'])
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:30%; text-align:left;")
            soOPheaderCell[-1].text = "Run Time: "
            etree.SubElement(soOPheaderCell, "span", style="display: inline-block; width:30%; text-align:left;")
            soOPheaderCell[-1].text = "Completed: "
            tempSOdata.drop(thisOP.index, inplace=True)
    
    
    
    shopOrderFile.append(bodyHTML)
    #etree.dump(shopOrderFile)
    shopOrderHTML = etree.ElementTree(shopOrderFile)
    shopOrderHTML.write("{}_{}.html".format(JobInformation['Job'], JobInformation['Suffix']))
    
    
    
    #
    # Section for alternate shop order view
    #
    
    shopOrder_PartView = copy.deepcopy(shopOrderDetails)
    # make a unique operation identifier that is smallest to largest in operation order
    shopOrder_PartView['opOrder'] = shopOrder_PartView['Major']*max(shopOrder_PartView['Minor']) + shopOrder_PartView['Minor']
    # use the smallest operation identifier for each item as the part identifier
    shopOrder_PartView['partOrder'] = shopOrder_PartView['opOrder'].apply(lambda x: min(shopOrder_PartView.loc[shopOrder_PartView['item'] == shopOrder_PartView.loc[shopOrder_PartView['opOrder'] == x, :]['item'].iloc[0],:]['opOrder']))
    #sort them to iterate sequentially to generate shop order
    shopOrder_PartView = shopOrder_PartView.sort_values(['partOrder', 'opOrder'])





















