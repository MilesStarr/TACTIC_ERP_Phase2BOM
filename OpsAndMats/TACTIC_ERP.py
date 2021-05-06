# -*- coding: utf-8 -*-
"""
Created on Tue Apr  6 14:41:48 2021

@author: eclark
"""

import requests
from requests_negotiate_sspi import HttpNegotiateAuth
from io import StringIO
import pandas as pd
from lxml import etree
import copy
import re

class TACTIC_ERP:
    """
    Rebuild of older BOM class generated in 2020 that utilized data from 
    excel files, but now that TACTIC has crossed Go-Live to using ERP
    with Infor SyteLine, aka the bane of my existance or the software from 
    1960 that has incremental patches to keep it alive and make money, the
    excel file data is outdated and needs to get polled live from the SQL 
    server.  To this end we make requests from MSSQL via ReportServer to
    access read-only versions of what is contained by the ERP.  Any input
    into the ERP needs to be written to an excel file to then be copied 
    and pasted into the SyleLine software.
    """

    # Server connection strings
    _report_server = "http://erpsql01.labtesting.local/"
    _url_item = "ReportServer?%2FTactic%20Reports%2FTacticItems&Item="
    _url_drawing = "ReportServer?%2FTactic%20Reports%2FTacticDrawings&Dwg="
    _url_noninv = "ReportServer?%2FTactic%20Reports%2FTacticNonInvItems&NonInvItem="
    _url_work_center = "ReportServer?%2FTactic%20Reports%2FTacticWcs&WC="
    _url_current_operation = "ReportServer?%2FTactic%20Reports%2FTestTacticCurrentOps&Item="
    _url_current_material = "ReportServer?%2FTactic%20Reports%2FTestTacticCurrentMatls&Item="
    _report_suffix = "&rc:Toolbar=False&rs:Format=csv"

    # Data Parameters
    item_dtypes = {'item': str, 'description': str, 'Uf_tt_description': str, 'revision': str, 'Revision Track': int, 'ECN': int, 'drawing_nbr': str, 'Alternate Item': str, 'Buyer': str, 'stocked': int, 'Show In Drop-Down Lists': int, 'u_m': str, 'matl_type': str, 'p_m_t_code': str, 'product_code': str, 'abc_code': str, 'cost_type': str, 'cost_method': str, 'Unit Cost': str, 'Current Unit Cost': str, 'Lot Size': int, 'Unit Weight': int, 'Weight Units': str, 'Quantity On Hand': float, 'Non-Nettable Stock': float, 'Safety Stock': float, 'Quantity Ordered': float, 'Quantity WIP': float, 'Allocated To Prod': float, 'Allocated To Customer Orders': float, 'Reserved For Customer Orders': float, 'Low Level': int, 'active_for_data_integration': int, 'plan_code': str, 'Shrink Factor': float, 'phantom_flag': int, 'MPS Flag': int, 'Net Change': int, 'MPS Plan Fence': int, 'Family Code': str, 'Production Type': str, 'Rate/Day': float, 'Inventory LCL %': str, 'Inventory UCL %': str, 'Supply Site': str, 'Supply Whse': str, 'Paper Work': int, 'Fixed Lead Time': int, 'Expedited Fixed': int, 'Dock-to-Stock': int, 'Variable': int, 'Expedited Variable': int, 'Separation': str, 'Release 1': str, 'Release 2': str, 'Release 3': str, 'MRP Item': int, 'Infinite': int, 'Planned Mfg Supply Switching': int, 'Accept Requirements': int, 'Pass Requirements': int, 'Must use future POs before creating PLNs': int, 'Supply Usage Tolerance': int, 'Time Fence Rule': str, 'Time Fence Value': int, 'Pull-Up SS Rule': str, 'Pull-Up SS Value': int, 'Setup Group': str, 'Charge Item': str, 'Order Minimum': int, 'Order Multiple': int, 'Order Maximum': int, 'Days Supply': int, 'Use Reorder Point': int, 'Reorder Point': int, 'Fixed Order Qty': int, 'Earliest Planned Purchase Receipt': str, 'Targeted Safety Stock Replenishment': str, 'Lot Track': int, 'Preassign Lots': int, 'Lot Prefix': str, 'S/N Track': int, 'Preassign Serials': int, 'S/N Prefix': str, 'Shelf Life': str, 'Issue By': str, 'stat': str, 'Reason': str, 'Last Change': str, 'User': str, 'backflush': int, 'bflush_loc': str, 'Preferred Co-product Mix': str, 'Reservable': int, 'Tax-Free Imported Material': int, 'Tax Free Days': int, 'Safety Stock Percent': int, 'Tariff Classification': str, 'PO Tolerance Over': str, 'PO Tolerance Under': str, 'Kit': int, 'Print Kit Components on Customer Paperwork': int} #, 'Std Due Period': str, 'Commodity': str, 'Commodity Description': str, 'Tax Code': str, 'Tax Code Description': str, 'Origin': str, 'Country': str, 'Preference Criterion': str, 'Country Of Origin': str, 'Producer': int, 'Subject To RVC Requirements': int, 'Purchased YTD': float, 'Manufactured YTD': float, 'Used YTD': float, 'Sold YTD': float, 'Subject To Excise Tax': int, 'Excise Tax Percent': float, 'Wholesale Price': float, 'Includes Item Content': int, 'Order Configurable': int, 'Job Configurable': int, 'Auto Job Generation': str, 'Name Space': str, 'Configuration Flag': int, 'Feature String': str, 'Feature Template': str, 'Last Import Date': str, 'Save Current Revision Upon Import': int, 'Overview': str, 'Active For Customer Portal': int, 'Featured Item': int, 'Top Seller': int, 'Item Attribute Group': str, 'Item Attribute Group Description': str, 'Lot Attribute Group': str, 'Lot Attribute Group Description': str, 'Enable Pieces Inventory': int, 'Piece Dimension Group': str, 'Piece Dimension Group Description': str, 'Portal Pricing Enabled': int, 'Portal Pricing Site': str, 'Freight': str, 'Estimated Break Date': str, 'Date of Last Report': str, 'Commodity Jurisdiction': str, 'ECCN or USML CAT': str, 'Program (ITAR/EAR600 Series)': str, 'Schedule B Number': str, 'HTS Code': str, 'HTS Code Description': str, 'Country Of Origin': str, 'Length Linear Dimension': str, 'Linear Dimension UM': str, 'Width Linear Dimension': str, 'Height Linear Dimension': str, 'Density': float, 'Density UM': str, 'Area': float, 'Area UM': str, 'Bulk Mass': float, 'Bulk Mass UM': str, 'Ream Mass': float, 'Ream Mass UM': str, 'Paper Mass Basis': str, 'Grade': str, 'Abnormal Size': int}
    itemERP_dtypes = {'item': str, 'description': str, 'drawing_nbr': str, 'stocked': int, 'Uf_tt_description': str, 'u_m': str, 'matl_type': str, 'p_m_t_code': str, 'product_code': str, 'abc_code': str, 'cost_type': str, 'cost_method': str, 'active_for_data_integration': int, 'plan_code': str, 'phantom_flag': int, 'stat': str, 'backflush': int, 'bflush_loc': str}
    recommendation_item = {'item': "", 'description': "", 'Uf_tt_description': "", 'revision': "", 'Revision Track': 0, 'ECN': 0,
                           'drawing_nbr': "", 'Alternate Item': "", 'Buyer': "", 'stocked': 1, 'Show In Drop-Down Lists': 1,
                           'u_m': "EA", 'matl_type': "Material", 'p_m_t_code': "Manufactured", 'product_code': "__Unknown__", 'abc_code': "C",
                           'cost_type': "Actual", 'cost_method': "FIFO", 'Unit Cost': "", 'Current Unit Cost': "", 'Lot Size': 1,
                           'Unit Weight': 0, 'Weight Units': "", 'Quantity On Hand': 0.0, 'Non-Nettable Stock': 0.0, 'Safety Stock': 0.0,
                           'Quantity Ordered': 0.0, 'Quantity WIP': 0.0, 'Allocated To Prod': 0.0, 'Allocated To Customer Orders': 0.0,
                           'Reserved For Customer Orders': 0.0, 'Low Level': 0, 'active_for_data_integration': 1, 'plan_code': "",
                           'Shrink Factor': 0.0, 'phantom_flag': 0, 'MPS Flag': 0, 'Net Change': 0, 'MPS Plan Fence': 0, 'Family Code': "",
                           'Production Type': "Job", 'Rate/Day': 1.0, 'Inventory LCL %': "", 'Inventory UCL %': "", 'Supply Site': "",
                           'Supply Whse': "", 'Paper Work': 0, 'Fixed Lead Time': 0, 'Expedited Fixed': 0, 'Dock-to-Stock': 0, 
                           'Variable': 0, 'Expedited Variable': 0, 'Separation': "", 'Release 1': "", 'Release 2': "", 'Release 3': "",
                           'MRP Item': 0, 'Infinite': 0, 'Planned Mfg Supply Switching': 0, 'Accept Requirements': 1, 'Pass Requirements': 1,
                           'Must use future POs before creating PLNs': 0, 'Supply Usage Tolerance': 0, 'Time Fence Rule': "No Time Fence",
                           'Time Fence Value': 0, 'Pull-Up SS Rule': "No Pull-Up", 'Pull-Up SS Value': 0, 'Setup Group': "",
                           'Charge Item': "", 'Order Minimum': 0, 'Order Multiple': 0, 'Order Maximum': 0, 'Days Supply': 0, 
                           'Use Reorder Point': 0, 'Reorder Point': 0, 'Fixed Order Qty': 0, 'Earliest Planned Purchase Receipt': "",
                           'Targeted Safety Stock Replenishment': "", 'Lot Track': 0, 'Preassign Lots': 0, 'Lot Prefix': "",
                           'S/N Track': 0, 'Preassign Serials': 0, 'S/N Prefix': "", 'Shelf Life': "", 'Issue By': "LOT", 
                           'stat': "Active", 'Reason': "", 'Last Change': "", 'User': "", 'backflush': 1,
                           'bflush_loc': "", 'Preferred Co-product Mix': "", 'Reservable': 0, 'Tax-Free Imported Material': 0,
                           'Tax Free Days': 0, 'Safety Stock Percent': 0, 'Tariff Classification': "", 'PO Tolerance Over': "",
                           'PO Tolerance Under': "", 'Kit': 0, 'Print Kit Components on Customer Paperwork': 0}
#    itemERP_missingCols = {x: y for x, y in recommendation_item.items() if x not in list(itemERP_dtypes)}

    drawings_dtypes = {'drawing_nbr': str, 'revision': str, 'drawing_size': str, 'approved': int, 'title': str, 'book': str, 'drawing_date': str, 'drn': str, 'project_num': str, 'UpdatedBy': str, 'RecordDate': str}

    operation_dtypes = {'item': str, 'item_description': str, 'Alternate ID': str, 'description': str, 'oper_num': int, 'Shared': str, 'WC': str, 'WC Description': str, 'Use Fixed Schedule': int, 'Fixed Sched Hours': str, 'Run-Hours Basis (Machine)': str, 'Mach Hrs per Piece': float, 'Run-Hours Basis (Labor)': str, 'Labor Hr per Piece': int, 'Sched Driver': str, 'Run Duration': float, 'Batch Definition': str, 'Yield': int, 'Seconds Per Cycle': int, 'Formula Material Weight': int, 'Formula Material Weight U/M': str, 'Move Hours': int, 'Queue Time': int, 'Setup Hours': int, 'Finish': int, 'Use Offset Hours': str, 'Offset Hours': int, 'Effective Date': str, 'Obsolete Date': str, 'Control Point': int, 'backflush': str, 'Setup Resource Group': str, 'Setup Rule': str, 'Setup Basis': str, 'Setup Time Rule': str, 'Setup Matrix': str, 'Scheduler Rule': str, 'Custom Planner Rule': int, 'Break Rule': str, 'Split Rule': str, 'Split Size': int, 'Split Group': str, 'Efficiency': int, 'Setup Rate': str, 'Run Rate (Labor)': int, 'Var Mach Ovhd Rate': int, 'Fix Machine Ovhd Rate': int, 'Var Ovhd Rate': int, 'Fixed Ovhd Rate': int}
    operationERP_dtypes = {'item': str, 'p_m_t_code': str, 'product_code': str, 'drawing_nbr': str, 'item_description': str, 'item_tt_description': str, 'revision': str, 'oper_num': int, 'u_m': str, 'wc': str, 'phantom_flag': int, 'wc_description': str, 'run_mch_hrs': float, 'run_lbr_hrs': float, 'sched_drv': str}

    material_dtypes = {'item': str, 'item_description': str, 'Alternate ID': str, 'description': str, 'oper_num': int, 'Shared': str, 'WC': str, 'WC Description': str, 'material': str, 'matl_description': str, 'Seq': int, 'Alt Group': int, 'Alt Group Rank': int, 'Manufacturer': str, 'Manufacturer Name': str, 'Manufacturer Item': str, 'Manufacturer Item Description': str, 'matl_type': str, 'matl_qty_conv': int, 'Per': str, 'matl_u_m': str, 'Cost': int, 'Scrap Factor': int, 'Effective Date': str, 'Obsolete Date': str, 'BOM Seq': str, 'Ref': str, 'backflush': int, 'bflush_loc': str, 'Feature': str, 'Option Code': str, 'Probable': int, 'Incremental Price': int, 'Formula Material Weight %': int, 'Estimated Break Date': str, 'Date of Last Report': str, 'Fixed Material': int, 'Variable Material': int, 'Material Cost': int, 'Labor Cost': int, 'Outside Cost': int, 'Fixed Overhead Cost': int, 'Variable Overhead Cost': int}
    materialERP_dtypes = {'item': str, 'item_p_m_t_code': str, 'item_product_code': str, 'item_drawing_nbr': str, 'item_revision': str, 'item_description': str, 'item_tt_description': str, 'item_u_m': str, 'item_phantom_flag': int, 'oper_num': int, 'material': str, 'matl_qty_conv': float, 'matl_u_m': str, 'matl_p_m_t_code': str, 'matl_product_code': str, 'matl_drawing_nbr': str, 'matl_revision': str, 'matl_description': str, 'matl_phantom_flag': int, 'matl_tt_description': str}

    workCenterERP_dtypes = {'wc': str, 'description': str}

    noninvERP_dtypes = {'item': str, 'description': str, 'u_m': str, 'matl_type': str, 'product_code': str}
    

    recommendation_item = {'item': "", 'description': "", 'Uf_tt_description': "", 'revision': "", 'Revision Track': 0, 'ECN': 0,
                           'drawing_nbr': "", 'Alternate Item': "", 'Buyer': "", 'stocked': 1, 'Show In Drop-Down Lists': 1,
                           'u_m': "EA", 'matl_type': "Material", 'p_m_t_code': "Manufactured", 'product_code': "__Unknown__", 'abc_code': "C",
                           'cost_type': "Actual", 'cost_method': "FIFO", 'Unit Cost': "", 'Current Unit Cost': "", 'Lot Size': 1,
                           'Unit Weight': 0, 'Weight Units': "", 'Quantity On Hand': 0.0, 'Non-Nettable Stock': 0.0, 'Safety Stock': 0.0,
                           'Quantity Ordered': 0.0, 'Quantity WIP': 0.0, 'Allocated To Prod': 0.0, 'Allocated To Customer Orders': 0.0,
                           'Reserved For Customer Orders': 0.0, 'Low Level': 0, 'active_for_data_integration': 1, 'plan_code': "",
                           'Shrink Factor': 0.0, 'phantom_flag': 0, 'MPS Flag': 0, 'Net Change': 0, 'MPS Plan Fence': 0, 'Family Code': "",
                           'Production Type': "Job", 'Rate/Day': 1.0, 'Inventory LCL %': "", 'Inventory UCL %': "", 'Supply Site': "",
                           'Supply Whse': "", 'Paper Work': 0, 'Fixed Lead Time': 0, 'Expedited Fixed': 0, 'Dock-to-Stock': 0, 
                           'Variable': 0, 'Expedited Variable': 0, 'Separation': "", 'Release 1': "", 'Release 2': "", 'Release 3': "",
                           'MRP Item': 0, 'Infinite': 0, 'Planned Mfg Supply Switching': 0, 'Accept Requirements': 1, 'Pass Requirements': 1,
                           'Must use future POs before creating PLNs': 0, 'Supply Usage Tolerance': 0, 'Time Fence Rule': "No Time Fence",
                           'Time Fence Value': 0, 'Pull-Up SS Rule': "No Pull-Up", 'Pull-Up SS Value': 0, 'Setup Group': "",
                           'Charge Item': "", 'Order Minimum': 0, 'Order Multiple': 0, 'Order Maximum': 0, 'Days Supply': 0, 
                           'Use Reorder Point': 0, 'Reorder Point': 0, 'Fixed Order Qty': 0, 'Earliest Planned Purchase Receipt': "",
                           'Targeted Safety Stock Replenishment': "", 'Lot Track': 0, 'Preassign Lots': 0, 'Lot Prefix': "",
                           'S/N Track': 0, 'Preassign Serials': 0, 'S/N Prefix': "", 'Shelf Life': "", 'Issue By': "LOT", 
                           'stat': "Active", 'Reason': "", 'Last Change': "", 'User': "", 'backflush': 1,
                           'bflush_loc': "", 'Preferred Co-product Mix': "", 'Reservable': 0, 'Tax-Free Imported Material': 0,
                           'Tax Free Days': 0, 'Safety Stock Percent': 0, 'Tariff Classification': "", 'PO Tolerance Over': "",
                           'PO Tolerance Under': "", 'Kit': 0, 'Print Kit Components on Customer Paperwork': 0}
    recommendation_operation = {'item': "", 'item_description': "", 'Alternate ID': "", 'description': "", 'oper_num': 0,
                                'Shared': "", 'WC': "TACwip", 'WC Description': "TACTIC-Pull Materials from Inventory",
                                'Use Fixed Schedule': 0, 'Fixed Sched Hours': "", 'Run-Hours Basis (Machine)': "Mch Hours/Pc",
                                'Mach Hrs per Piece': 0.5, 'Run-Hours Basis (Labor)': "Lbr Hours/Pc", 'Labor Hr per Piece': 0,
                                'Sched Driver': "Machine", 'Run Duration': 0.5, 'Batch Definition': "", 'Yield': 100, 'Seconds Per Cycle': 0,
                                'Formula Material Weight': 0, 'Formula Material Weight U/M': "", 'Move Hours': 0, 'Queue Time': 0,
                                'Setup Hours': 0, 'Finish': 0, 'Use Offset Hours': "", 'Offset Hours': 0, 'Effective Date': "",
                                'Obsolete Date': "", 'Control Point': 1, 'backflush': "Neither", 'Setup Resource Group': "",
                                'Setup Rule': "Always", 'Setup Basis': "Item", 'Setup Time Rule': "Fixed Time", 'Setup Matrix': "",
                                'Scheduler Rule': "Per Piece", 'Custom Planner Rule': 0, 'Break Rule': "Shifts", 'Split Rule': "No Splitting",
                                'Split Size': 0, 'Split Group': "", 'Efficiency': 100, 'Setup Rate': "", 'Run Rate (Labor)': 35,
                                'Var Mach Ovhd Rate': 35, 'Fix Machine Ovhd Rate': 0, 'Var Ovhd Rate': 0, 'Fixed Ovhd Rate': 0}
    recommendation_material  = {'item': "", 'item_description': "", 'Alternate ID': "", 'description': "", 'oper_num': 0, 'Shared': "",
                                'WC': "TACwip", 'WC Description': "", 'material': "", 'matl_description': "", 'Seq': 0, 'Alt Group': 0, 
                                'Alt Group Rank': 0, 'Manufacturer': "", 'Manufacturer Name': "", 'Manufacturer Item': "",
                                'Manufacturer Item Description': "", 'matl_type': "Material", 'matl_qty_conv': 0, 'Per': "Unit", 'matl_u_m': "EA", 'Cost': 0,
                                'Scrap Factor': 0, 'Effective Date': "", 'Obsolete Date': "", 'BOM Seq': "", 'Ref': '', 'backflush': 1,
                                'bflush_loc': "", 'Feature': "", 'Option Code': "", 'Probable': 1, 'Incremental Price': 0,
                                'Formula Material Weight %': 0, 'Estimated Break Date': "", 'Date of Last Report': "", 'Fixed Material': 0,
                                'Variable Material': 0, 'Material Cost': 0, 'Labor Cost': 0, 'Outside Cost': 0, 'Fixed Overhead Cost': 0,
                                'Variable Overhead Cost': 0}

    # Item Regex Patterns
    regex_hardware = re.compile("[12](\d{2})(\d{2})(\d{2})(.*)")
    regex_drawing  = re.compile("80(\d{5})(.*)")
    regex_wire     = re.compile("787(?P<type>\d)(?P<color>\d)(?P<tracer>\d)(?P<size>[B-NP-Z])(?P<dash>-.+)?")
    regex_cable    = re.compile("788(?P<type>[0-3])(?P<cond>\d{2})(?P<size>[B-NP-Z])(?P<dash>-.+)?")


    def __init__(self, BOM = None):
        """
        Process the BOM if provided to initialize the instance.  
        BOM can be:
            an etree _ElementTree object (copy existing)
            a string representing a filepath to open an XML file with etree
            None (default) has no XML for creating ERP Jobs
        """
        
        if BOM is None:
            self.BOM = None
        elif isinstance(BOM, etree._ElementTree):
            self.BOM = copy.deepcopy(BOM.getroot())
        elif isinstance(BOM, str):
            self.BOM = etree.parse(BOM).getroot()
        
        # Retrieve items listed in BOM if a BOM was provided
        itemsInBOM = set()
        if isinstance(self.BOM, etree._Element):
            itemsInBOM.update(set(map(lambda PartXML: PartXML.attrib['PartID'], self.BOM.findall(".//Part"))))
            itemsInBOM.update(set(map(lambda MatXML: MatXML.attrib['PartID'], self.BOM.findall(".//Material"))))
            self.itemsInBOM = itemsInBOM
            itemRequest = requests.get(''.join([self._report_server, self._url_item, ",".join(itemsInBOM), self._report_suffix]), auth=HttpNegotiateAuth())
            self.items = pd.read_csv(StringIO(itemRequest.content.decode("utf-8")), na_filter=False, dtype=self.itemERP_dtypes)
        else:
            self.items = None

        # Retrieve a list of drawings associated to items
        drawingItems = [matched[1] for matched in (self.regex_drawing.fullmatch(itemSearch) for itemSearch in itemsInBOM) if matched is not None]
        if len(drawingItems) > 0:
            drawingsInBOM = requests.get(''.join([self._report_server, self._url_drawing, ",".join(drawingItems), self._report_suffix]), auth=HttpNegotiateAuth())
            self.drawings = pd.read_csv(StringIO(drawingsInBOM.content.decode("utf-8")), na_filter=False, dtype=self.drawings_dtypes)

        # Retrieve work centers
        if isinstance(self.BOM, etree._Element):
            wcInBOM = set(map(lambda MatXML: MatXML.attrib['WC'], self.BOM.findall(".//Operation")))
            wcInBOM = requests.get(''.join([self._report_server, self._url_work_center, ",".join(wcInBOM), self._report_suffix]), auth=HttpNegotiateAuth())
            self.wc = pd.read_csv(StringIO(wcInBOM.content.decode("utf-8")), na_filter=False, dtype=self.workCenterERP_dtypes)

        # Initialize recommendation list
        self._itemRecommendations = pd.DataFrame(columns = list(self.item_dtypes))

    @property
    def itemRecommendations(self):
        return self._itemRecommendations

    def fetchItem(self, **kwargs):
        """ 
        Intended to be used for fetchItem(item = "ItemID")
        but any column value can be used for the parameter variable.
        Function searches to match parameter as the column name from
        the items table.  Failure to have an exact match will return None.
        """
        # Error checking inputs
        if len(kwargs) != 1:
            raise AssertionError("Too many parameters")
        for key in kwargs.keys():
            if not(key in self.itemERP_dtypes.keys()):
                raise NameError(key, " not found in item header")
        
        key = list(kwargs.keys())[0]
        value = list(kwargs.values())[0]
        
        # Get results from items table
        partItemDef = self.items.loc[self.items[key] == value]

        if len(partItemDef)==1:
            return partItemDef.iloc[0]
        elif len(partItemDef) > 1:
            print("multiple rows match " + str(key) + " == " + str(value))
            return None
        else:
            print(str(key) + " == " + str(value) + " not found in items table.")
            return None

    def makeItemRecommendation(self, **kwargs):
        """
        It makes a guess at what the item might be if it is given 
        Item = "value" and returns a series that should be added to the
        _itemRecommendations DataFrame.  If it fails to match known patterns
        it will return a row with the item value and blank cells, or
        it will return a row with "unknown" to signify it wasn't called with the
        Item = "value" keyword and didn't try
        """

        # Error checking inputs
        if len(kwargs) != 1:
            raise AssertionError("Too many parameters")
        for key in kwargs.keys():
            if not(key in self.item_dtypes.keys()):
                raise NameError(key, " not found in item header")
        
        key = list(kwargs.keys())[0]
        value = list(kwargs.values())[0]
        
        recommendation = copy.deepcopy(self.recommendation_item)
        
        if key == "item":
            itemSearch = value
            
            if len(self._itemRecommendations.loc[self._itemRecommendations['item'] ==  value]) > 0:
                # catch existing item recommendation entry exists and return first (hopefully only) row
                return self._itemRecommendations.loc[self._itemRecommendations['item'] ==  value].iloc[0]
            
            itemIsMfg = self.BOM.find(".//Part[@PartID='{0}']".format(str(itemSearch)))
            
            itemIsDWG = self.regex_drawing.fullmatch(itemSearch)
            itemIsHardware = self.regex_hardware.fullmatch(itemSearch)
            itemIsWire = self.regex_wire.fullmatch(itemSearch)
            itemIsCable = self.regex_cable.fullmatch(itemSearch)
            
            if itemIsDWG:
                dwgInfo = self.fetchDrawing(itemIsDWG[1])
                if type(dwgInfo) == pd.Series:
                    recommendation['item'] = itemSearch
                    if len(dwgInfo['title']) > 40:
                        recommendation['description'] = dwgInfo['title'][:39] + "\u2026"
                    else:
                        recommendation['description'] = dwgInfo['title']
                    recommendation['Uf_tt_description'] = dwgInfo['title']
                    recommendation['drawing_nbr'] = dwgInfo['drawing_nbr']
                    recommendation['revision'] = dwgInfo['revision']
                else:
                    recommendation['item'] = itemSearch
                    recommendation['description'] = "Failed Drawing Lookup"
                recommendation['phantom_flag'] = 1 if itemIsDWG[2][:2] == "-X" else 0
                recommendation['product_code'] = "TT-MfgStep" if itemIsMfg is not None else "TT-NonSale"
            elif itemIsHardware:
                hardwareInfo = self.fetchHardware(itemIsHardware)
                recommendation['item'] = itemSearch
                recommendation['description'] = hardwareInfo if len(hardwareInfo) < 41 else hardwareInfo[:39] + "\u2026"
                recommendation['Uf_tt_description'] = hardwareInfo
                recommendation['p_m_t_code'] = "Purchased"
                recommendation['product_code'] = "TT-Raw"
            elif itemIsWire:
                wireInfo = self.fetchWire(itemIsWire)
                recommendation['item'] = itemSearch
                recommendation['description'] = wireInfo if len(wireInfo) < 41 else wireInfo[:39] + "\u2026"
                recommendation['Uf_tt_description'] = wireInfo
                recommendation['p_m_t_code'] = "Purchased"
                recommendation['u_m'] = "IN"
                recommendation['product_code'] = "TT-Raw"
            elif itemIsCable:
                cableInfo = self.fetchCable(itemIsCable)
                recommendation['item'] = itemSearch
                recommendation['description'] = cableInfo if len(cableInfo) < 41 else cableInfo[:39] + "\u2026"
                recommendation['Uf_tt_description'] = cableInfo
                recommendation['p_m_t_code'] = "Purchased"
                recommendation['u_m'] = "IN"
                recommendation['product_code'] = "TT-Raw"
            else:
                recommendation['item'] = itemSearch
        else:
            recommendation['item'] = "Unknown"
            recommendation['description'] = "Unknown"
            recommendation['Uf_tt_description'] = "Unknown"
            recommendation['u_m'] = "Unknown"
        self._itemRecommendations = self._itemRecommendations.append(pd.Series(recommendation), ignore_index = True)
        return pd.Series(recommendation)

    def fetchDrawing(self, DWG_NO = "12345"):
        """
        Input is the TACTIC drawing number (5 digits at this time)
        It checks it agains the class's pd.DataFrame of drawing information
        and returns the entire row back, or returns None to signal an error
        """
        dwgDefs = self.drawings.loc[self.drawings['drawing_nbr'] == DWG_NO]
        if len(dwgDefs) == 1:
            return dwgDefs.iloc[0]
        elif len(dwgDefs) == 0:
            print("drawing {} not in index".format(DWG_NO))
            return None
        else:
            print("drawing {} overdefined in index".format(DWG_NO))
            return None

    def fetchHardware(self, hardwareMatchResult):
        """
        Input is the regex match against self.regex_hardware
        Returns a string description trying to match definitions from TT-107
        """
        HardwareType = {'00': "Hex Nut", '05': "Hex Nut", '01': "Hex Nut, narrow",
                        '03': "Nut, ASF full", '04': "Nut, ASF jam",
                        '10': "Nut, Lt Flexloc", '11': "Nut, Thin Lt Flexloc",
                        '12': "Nut, Hvy Flexloc", '13': "Nut, Thin Hvy Flexloc",
                        '14': "Nut, Locking, Other", '15': "Nut, Self Clinching",
                        '16': "Nut, Acorn", '18': "Nut, Wing", '19': "Nut, Other",
                        '21': "M.Scr, Rnd Hd.", '22': "M.Scr, Rnd Hd.",
                        '23': "M.Scr, Bnd Hd.", '24': "M.Scr, Bnd Hd.",
                        '25': "M.Scr, Pan Hd.", '26': "M.Scr, Pan Hd.",
                        '27': "M.Scr, Fstr Hd.", '28': "M.Scr, Fstr Hd.",
                        '29': "M.Scr, 82°FH", '30': "M.Scr, 82°FH",
                        '31': "M.Scr, Oval Hd.", '32': "M.Scr, Oval Hd.",
                        '33': "Setscrew, Cup Pt", '34': "St.Scr, Lck, Cup Pt",
                        '35': "BHMS", '36': "M.Scr, 100°FH",
                        '38': "Capscrew, 82°FH", '40': "HHCS", '41': "HHCS, Locking",
                        '42': "SHCS", '43': "SHCS, Nylok", '44': "Carriage Bolt",
                        '45': "Drive Screw", '46': "Elevator Bolt",
                        '47': "Shoulder Bolt", '48': "Thumb Screw", '49': "Sheet Metal Screw",
                        '50': "Washer, Standard", '51': "Washer, SAE", '52': "Washer, AN",
                        '53': "Washer, MS", '54': "Washer, Fender", '56': "Washer, Split Lock, Lt",
                        '57': "Washer, Split Lock, Med", '58': "Washer, Split Lock, Hvy",
                        '59': "Washer, Lock, Int. Tooth", '60': "Washer, Lock, Ext. Tooth",
                        '70': "Threaded Rod"}
        HardwareSize = {'00': "Special", '01': "0-80", '02': "1-64", '03': "1-72",
                        '04': "2-56", '05': "2-64", '06': "3-48", '07': "3-56",
                        '08': "4-48", '09': "4-40", '10': "5-40", '11': "5-44",
                        '12': "6-32", '13': "6-40", '14': "8-32", '15': "8-36",
                        '16': "10-24", '17': "10-32", '18': "12-24", '19': "12-28",
                        '20': "¼-20", '21': "¼-28", '22': "¼-32", '23': "9/32-32",
                        '24': "5/16-18", '25': "5/16-24", '26': "5/16-32", '27': "11/32-32",
                        '28': "3/8-16", '29': "3/8-24", '30': "3/8-32", '31': "7/16-14",
                        '32': "7/16-20", '33': "7/16-27", '34': "15/32-32", '35': "½-13",
                        '36': "½-20", '37': "½-32", '38': "9/16-18", '39': "9/16-18",
                        '40': "5/8-11", '41': "5/8-18", '42': "11/16-27", '43': "¾-10",
                        '44': "¾-16", '45': "7/8-9", '46': "7/8-14", '47': "1-8",
                        '48': "1-14", '49': "1-27", '50': "#00", '51': "#0",
                        '52': "#1", '53': "#2", '54': "#3", '55': "#4", '56': "#5",
                        '57': "#6", '58': "#7", '59': "#8", '60': "#9", '61': "#10",
                        '62': "#12", '63': "#14", '64': "#16", '65': "#18", '66': "#20",
                        '67': "#24", '70': "1/8", '71': "3/16", '72': "¼", '73': "5/16",
                        '74': "3/8", '75': "7/16", '76': "½", '77': "9/16", '78': "5/8",
                        '79': "11/16", '80': "¾", '81': "13/16", '82': "7/8", '83': "15/16",
                        '84': "1 to 1 7/8‡", '85': "2 to 2 7/8‡",
                        '90': "< 3mm", '91': "M3 - M9‡", '92': "M10 - M19‡"}
        hType = HardwareType.get(hardwareMatchResult[1], "Unknown Type")
        hSize = HardwareSize.get(hardwareMatchResult[2], "Unknown Size")

        hLen = ""
        try:
            if int(hardwareMatchResult[1]) < 50 and int(hardwareMatchResult[1]) > 20 :
                # Fastener with Length
                inches = int(hardwareMatchResult[3]) // 8
                eighths = int(hardwareMatchResult[3]) % 8
                if eighths:
                    if eighths % 2 == 0:
                        if eighths % 4 == 0:
                            hLen = (", {0:d} {1:d}/2" if inches > 0 else ", {1:d}/2").format(inches, eighths // 4)
                        else:
                            hLen = (", {0:d} {1:d}/4" if inches > 0 else ", {1:d}/4").format(inches, eighths // 2)
                    else:
                        hLen = (", {0:d} {1:d}/8" if inches > 0 else ", {1:d}/8").format(inches, eighths)
                else:
                    hLen = (", {0:d}" if inches > 0 else "").format(inches, eighths)
        except:
            hLen = None
        
        stainless = ", St. Stl." if hardwareMatchResult[0][0] == "1" else ""
        special = "" if hardwareMatchResult[4] == "" else ",  something special here"

        return hType + ", " + hSize + hLen + stainless + special

    def fetchWire(self, wireMatchResult):
        """
        Input is the regex match against self.regex_wire
        Returns a string description trying to match definitions from TT-107
        """
        digit4 = {'0': "Bare", '1': "Solid", '2': "Stranded", '5': "Enamel", '6': "Litz", '9': "Misc"}
        color = {'0': "BLK", '1': "BRN", '2': "RED", '3': "ORG", '4': "YEL", 
                 '5': "GRN", '6': "BLU", '7': "PUR", '8': "GRY", '9': "WHT"}
        size = {'B': "40", 'C': "38", 'D': "36", 'E': "34", 'F': "32", 'G': "30", 
                'H': "28", 'I': "26", 'J': "24", 'K': "22", 'L': "20", 'M': "18", 
                'N': "16", 'P': "14", 'Q': "12", 'R': "10", 'S': "8", 'T': "6", 
                'U': "4", 'V': "2", 'W': "0", 'X': "00", 'Y': "000", 'Z': "0000"}
        itemDict = {'wt': digit4.get(wireMatchResult['type'], "Unknown wire type"), 
                    'color': color.get(wireMatchResult['color'], "ERR"),
                    'tracer': color.get(wireMatchResult['tracer'], "ERR"),
                    'size': size.get(wireMatchResult['size'], "ERR"),
                    'dash': ", {} __special wire__".format(wireMatchResult['dash']) if wireMatchResult['dash'] else ""}
        if wireMatchResult['color'] == wireMatchResult['tracer']:
            return "Wire, Single, {wt}, {color}, {size}{dash}".format(**itemDict)
        else:
            return "Wire, Single, {wt}, {color} w/ {tracer}, {size}{dash}".format(**itemDict)

    def fetchCable(self, cableMatchResult):
        """
        Input is the regex match against self.regex_cable
        Returns a string description trying to match definitions from TT-107
        """
        digit4 = {'0': "Shld Sld", '1': "Shld Str", '2': "Unshld Sld", '3': "Unshld Str"}
        nCond = int(cableMatchResult['cond'])
        if nCond == 0:
            nCond = "50 T.Pair"
        elif nCond > 50:
            nCond = "{0:2d} T.Pair".format(nCond-50)
        elif nCond < 51 and nCond > 0:
            nCond = "{0:2d} Cond.".format(nCond)
        else:
            nCond = "Something went wrong"
        size = {'B': "40", 'C': "38", 'D': "36", 'E': "34", 'F': "32", 'G': "30", 
                'H': "28", 'I': "26", 'J': "24", 'K': "22", 'L': "20", 'M': "18", 
                'N': "16", 'P': "14", 'Q': "12", 'R': "10", 'S': "8", 'T': "6", 
                'U': "4", 'V': "2", 'W': "0", 'X': "00", 'Y': "000", 'Z': "0000"}
        modifier = {None: "", '-1': ", Rubber Jacket", '-2': ", PUR Jacket",
                    '-3': ", PTFE Jacket", '-V': ", VFD Cable"}
        itemDict = {'ct': digit4.get(cableMatchResult['type'], "Unknown cable type"),
                    'nCond': nCond,
                    'size': size.get(cableMatchResult['size'], "Unknown AWG"),
                    'dash': modifier.get(cableMatchResult['dash'], ", Unknown dash type")}
        return "Cable, {ct}, {nCond}, {size}{dash}".format(**itemDict)

    def isErrorInBOM(self, start = "", visited = [], depth = 0):
        errorExists = False
        if start == "":
            checkList = self.BOM.findall(".//Part")
            visited = []
        else:
            checkPart = self.BOM.find(".//Part[@PartID='{0}']".format(start))
            visited.append(checkPart.attrib['PartID'])
            checkList = []
            for mat in checkPart.findall(".//Material"):
                if mat.attrib['PartID'] in visited:
                    print("Found circular reference in Part {} sourcing Material {}".format(checkPart.attrib['PartID'], mat.attrib['PartID']))
                    errorExists = True
                else:
                    getRef = self.BOM.findall(".//Part[@PartID='{0}']".format(mat.attrib['PartID']))
                    if len(getRef) == 0:
                        pass
                    elif len(getRef) > 1:
                        print("Multiple listing for Part {} in the XML file".format(mat.attrib['PartID']))
                        errorExists = True
                    else:
                        checkList.append(getRef[0])
        print("\t"*depth, len(checkList), " materials in ", start, " against ", len(visited), " previous entries")
        for checkMe in checkList:
            print("\t"*(depth+1), "Checking {}".format(checkMe.attrib['PartID']))
            newError = self.isErrorInBOM(start = checkMe.attrib['PartID'], visited = copy.copy(visited), depth = depth + 1)
            errorExists = errorExists or newError
        return errorExists

    def buildItem(self, PartID):
        """
        Returns a tuple of pd.DataFrame (operations, materials) 
        unless the BOM is unable to locate the PartID requested
        either due to the PartID not contained in the BOM or if
        the BOM was never loaded into the instance
        """
        
        #initialize variables
        opTemplate = dict(self.recommendation_operation)  #make a local copy
        OpColumns = list(opTemplate)
        matTemplate = dict(self.recommendation_material)  #make a local copy
        MatColumns = list(matTemplate)
        
        Operations = pd.DataFrame(columns = OpColumns)
        Materials = pd.DataFrame(columns = MatColumns)
        
        # It's dangerous to go alone!
        if self.BOM is None:
            print("Cannot build {} without a BOM loaded".format(str(PartID)))
            return (Operations, Materials)  # Return the empty DataFrames
        
        itemDef = self.fetchItem(item = PartID)
        if itemDef is None:
            # Handle failed lookups.
            itemDef = self.makeItemRecommendation(item = PartID)
        
        matTemplate['item'] = itemDef['item']
        opTemplate['item'] = itemDef['item']
        matTemplate['item_description'] = itemDef['description']
        opTemplate['item_description'] = itemDef['description']
        
        buildItem = self.BOM.find(".//Part[@PartID='{0}']".format(str(PartID)))
        opID = 10
        partOPs = buildItem.findall(".//Operation")
        for op in partOPs:
            wcRow = self.wc.loc[self.wc['wc']  == op.attrib['WC']].iloc[0]
            thisOp = dict(opTemplate)
            thisOp.update({'item': buildItem.attrib["PartID"], 
                           'WC': op.attrib['WC'],
                           'WC Description': wcRow['wc'],
                           'oper_num': opID})
            Operations = Operations.append(thisOp, ignore_index = True)
            matList = []
            for mat in op.findall(".//Material"):
                matList.extend(self.genMatReturns(mat))
            seqID = 1
            for matRow in matList:
                thisMat = dict(matTemplate)
                thisMat.update({'item': buildItem.attrib['PartID'],
                                'oper_num': opID,
                                'WC': op.attrib['WC'],
                                'WC Description': wcRow['wc'],
                                'material': matRow['PartID'],
                                'matl_description': matRow['matl_description'],
                                "Seq": seqID,
                                "Alt Group": seqID,
                                "Alt Group Rank": 0,
                                'matl_qty_conv': matRow['matl_qty_conv'],
                                'matl_u_m': matRow['Unit']})
                Materials = Materials.append(thisMat, ignore_index = True)
                seqID += 1
            opID += 10
        return (Operations, Materials)

    def genMatReturns(self, target = etree.fromstring("<Material PartID='12345' Unit='EA' Qty='1' />"), quantity = 1):
        """
        Input is an lxml etree._Element for a material
        Returns a list of dicts for material unique keys only
        """
        
        results = []
        reportQty = 0
        
        try:
            reportQty = int(target.attrib['Qty']) * quantity
        except ValueError:
            reportQty = float(target.attrib['Qty']) * quantity  # if someone forgot to use Size properly for an assumed quantity of 1
        itemDef = self.fetchItem(item = target.attrib['PartID'])
        if itemDef is None:
            # Handle failed lookups.
            itemDef = self.makeItemRecommendation(item = target.attrib['PartID'])
        
        if itemDef['u_m'] != target.attrib['Unit']:
            print("Warning: Units mismatch for material " + target.attrib['PartID']) #Only print a warning because some unit conversions are acceptable

        if 'Size' not in target.attrib:
            results = [{'PartID': target.attrib['PartID'],
                        'matl_description': itemDef['description'],
                        'Unit': target.attrib['Unit'],
                        'matl_qty_conv': reportQty }]
        else:
            result = {'PartID': target.attrib['PartID'],
                      'matl_description': itemDef['description'],
                      'Unit': target.attrib['Unit'],
                      'matl_qty_conv': float(target.attrib['Size']) }
            results.extend([result] * reportQty) #list the result the number of times in reportQTY
        return results

#Start of Make Job functions
    def fraction(self, value = 0, maxDen = 16, unit = "IN"):
        if maxDen not in [2,4,8,16,32,64]:
            raise AssertionError("Denominator needs to be a power of 2 up to 64")
        result = {'whole': int(value), 'num': maxDen, 'den': maxDen, 'unit': str(unit) }
        remainder = value - result['whole']
        result['num'] = int(remainder * maxDen)
        if (result['num'] / maxDen) < remainder:
            result['num'] += 1
        #catch decimals greater than 15/16
        if result['num'] == maxDen:
            result['whole'] += 1
            result['num'], result['den'] = (0,0)
        while result['num'] % 2 == 0 and result['num'] != 0:
            result['num'] //= 2
            result['den'] //= 2
        if result['whole'] > 0:
            if result['num'] > 0:
                return "{whole:d} {num:d}/{den:d} {unit}".format(**result)
            else:
                return "{whole:d} {unit}".format(**result)
        elif result['num'] > 0:
            return "{num:d}/{den:d} {unit}".format(**result)
        else:
            return "err"

    def resolvePhantoms(self, MaterialsTable = pd.DataFrame()):
        #prime the phantom list
        phantomMats = MaterialsTable.loc[MaterialsTable['matl_phantom_flag'] == 1]
        
        (loopCount, loopLimit) = (0, 50)
        while len(phantomMats) != 0 and loopCount < loopLimit:
            #take the phantoms one at a time since the index will get FuBar during append
            phMat = phantomMats.iloc[0]
            
            addMe = MaterialsTable.loc[MaterialsTable['item'] == phMat['material']]
            if len(addMe) < 1:
                #fail quietly but not silently making a phantom without a BOM not a phantom anymore
                print("unable to find phantom " + phMat['material'])
                MaterialsTable.loc[phantomMats.index[0], 'matl_phantom_flag'] = 0
            else:
                #make the phantom materials look like they belong to the item 
                updateMe = ['item', 'item_p_m_t_code', 'item_product_code', 'item_drawing_nbr',
                            'item_revision', 'item_description', 'item_u_m', 'item_phantom_flag',
                            'oper_num']
                for update in updateMe:
                    addMe.loc[:, update] = phMat[update]
                
                #kill the phantom material and bring in it's materials
                MaterialsTable = MaterialsTable.drop(index = phantomMats.index[0])
                MaterialsTable = MaterialsTable.append(addMe, ignore_index = True)
                print("Resolved " + str(loopCount +1) + " phantom steps (" + phMat['material'] + " to " + phMat['item'] + ")")
            #regenerate phantom list since the append rebuilt the index and the drop needs a current index reference
            phantomMats = MaterialsTable.loc[MaterialsTable['matl_phantom_flag'] == 1]
            loopCount += 1
        #kill all the phantom item definitions since they are not needed anymore
        MaterialsTable = MaterialsTable.drop(index = MaterialsTable.loc[MaterialsTable['item_phantom_flag'] == 1].index)
        return MaterialsTable

    def buildOps(self, OperationsTable = pd.DataFrame(),
                    MaterialsTable = pd.DataFrame(),
                    Part = "", ParentQTY = 1,
                    ParentOP = 0, opID = 1, matID = 0, depth = 0):
    
    #empty list of materials to start this recursion level    
        matList = []
        refList = []
     
    # fetch and reformat operations associated to part.  Print error if none found
        partOpDef = OperationsTable.loc[OperationsTable['item'] == Part]
        columnRename = {'p_m_t_code': "item_p_m_t_code", 
                        'product_code': "item_product_code",
                        'drawing_nbr': "item_drawing_nbr",
                        'revision': "item_revision",
                        'u_m': "item_u_m",
                        'phantom_flag': "item_phantom_flag"}
        partOpDef.rename(columns = columnRename, inplace = True)
        if len(partOpDef) < 1:
            print("part " + str(Part) + " not found in operations list")
    
    # iterate over operations to find materials
        partOpDef = partOpDef.sort_values('oper_num', ascending = False)
        for opIndex, op in partOpDef.iterrows():
            mats = MaterialsTable.loc[(MaterialsTable['item'] == Part) & (MaterialsTable['oper_num'] == op['oper_num'])]
    #special case, add a no-material entry in material list
            if len(mats) == 0:
                matList.extend(self.buildMats(op,
                                         pd.Series({'item': "",
                                                    'item_p_m_t_code': "",
                                                    'item_product_code': "",
                                                    'item_drawing_nbr': "",
                                                    'item_revision': "",
                                                    'item_description': "",
                                                    'item_tt_description': "",
                                                    'item_u_m': "",
                                                    'item_phantom_flag': 0,
                                                    'oper_num': op['oper_num'],
                                                    'material': "",
                                                    'matl_qty_conv': 1,
                                                    'matl_u_m': "",
                                                    'matl_p_m_t_code': "P",
                                                    'matl_product_code': "",
                                                    'matl_drawing_nbr': "",
                                                    'matl_revision': "",
                                                    'matl_description': "",
                                                    'matl_phantom_flag': 0}),
                                         ParentQTY = ParentQTY,
                                         ParentOP = ParentOP,
                                         opID = opID,
                                         matID = matID,
                                         depth = depth))
                matID += 1
    #list materials under the operation
            for matIndex, mat in mats.iterrows():
                thisMat = self.buildMats(op,
                                    mat,
                                    ParentQTY = ParentQTY,
                                    ParentOP = ParentOP,
                                    opID = opID,
                                    matID = matID,
                                    depth = depth)
                matList.extend(thisMat)
                if (mat['matl_product_code'] == "TT-MfgStep") and (mat['matl_phantom_flag'] == 0):
                    refList.extend(thisMat)
                matID += 1
            ParentOP = opID
            opID += 1
        return (matList, refList, opID, matID)
        
    def buildMats(self,
                  OperationInfo = pd.Series(dtype=object),
                  MaterialInfo = pd.Series(dtype=object),
                  ParentQTY = 1, ParentOP = 0, opID = 1, matID = 0, depth = 0):
    # information about the recursion state
        recursionData = pd.Series(dict(zip(('ParentQTY', 'ParentOP', 'opID', 'matID', 'depth'),
                                           (ParentQTY, ParentOP, opID, matID, depth))))
    #remove data already contained by the operation series
        MaterialInfo = MaterialInfo.drop(labels = ['item', 'item_p_m_t_code', 'item_product_code', 
                                                   'item_drawing_nbr', 'item_revision', 'item_description',
                                                   'item_tt_description', 'item_u_m',
                                                   'item_phantom_flag', 'oper_num'])
    #determine if this material quantities should be aggragated or listed by size and generate list of return value(s)
        aggrUnits = ["EA", "PT", ""]
        sizeUnits = ["FT", "IN", "LB", "ml", "P", "SF"]
    
        aggrUnits = ["EA", "PT", "", "BX", "g"] #extended for testing on TestTactic Database
    
        if MaterialInfo['matl_u_m'] in aggrUnits:
            result = pd.concat([OperationInfo, MaterialInfo, recursionData])
            result.loc['matl_qty_conv'] = result['matl_qty_conv'] * result['ParentQTY']
            result.loc['ParentQTY'] = result['matl_qty_conv']
            results = [result]
        elif MaterialInfo['matl_u_m'] in sizeUnits:
            result = pd.concat([OperationInfo, MaterialInfo, recursionData])
            if(float(int(ParentQTY)) != ParentQTY):
                print("Material {0} had non-integer ParentQTY {1}, please check results for accuracy".format(result['material'], ParentQTY))
            results = [result for i in range(int(ParentQTY))]
    #return list of compiled series data
        return results

    def requestOPs(self, item = [], opCols = dict(operationERP_dtypes)):
        """
        take a string or list of strings
        request operations from ERP server
        return a pandas dataframe of results
        """

        if type(item) == list:
            for i in range(len(item)):
                if type(item[i]) != str:
                    item[i] = str(item[i])
            querry = ','.join(item)
        else:
            querry = str(item)
        ops = requests.get("".join([self._report_server, self._url_current_operation, querry, self._report_suffix]), auth=HttpNegotiateAuth())
        return pd.read_csv(StringIO(ops.content.decode("utf-8")), na_filter=False, dtype=opCols)
    
    def requestMats(self, item = [], matCols = dict(materialERP_dtypes)):
        """
        take a string or list of strings
        request operations from ERP server
        return a pandas dataframe of results
        """

        if type(item) == list:
            for i in range(len(item)):
                if type(item[i]) != str:
    #                print("Converting string" + str(item[i]))
                    item[i] = str(item[i])
            querry = ','.join(item)
        else:
            querry = str(item)
    #    print(querry.join(requestText))
        mats = requests.get("".join([self._report_server, self._url_current_material, querry, self._report_suffix]), auth=HttpNegotiateAuth())
        return pd.read_csv(StringIO(mats.content.decode("utf-8")), na_filter=False, dtype=matCols)

if __name__ == "__main__":

    
    name = "M411"
    tree = etree.parse("../" + name + ".xml")
    
    
    
    Operations = pd.DataFrame()
    Materials = pd.DataFrame()
    
    engBOM = TACTIC_ERP(tree)
    
    buildMe = tree.getroot().findall(".//Part")
    for buildItem in buildMe:
        (ops, mats) = engBOM.buildItem(buildItem.attrib["PartID"])
        Operations = Operations.append(ops, ignore_index=True)
        Materials = Materials.append(mats, ignore_index=True)
    engBOM.itemRecommendations.to_excel(name + "_recs.xlsx")
    
    OpColumns = list(engBOM.operation_dtypes)
    MatColumns = list(engBOM.material_dtypes)
    
    with pd.ExcelWriter(name + ".xlsx") as outFile:
        Materials.to_excel(outFile, sheet_name = "Materials", columns = MatColumns, index=False)
        Operations.to_excel(outFile, sheet_name = "Operations", columns = OpColumns, index=False)






































