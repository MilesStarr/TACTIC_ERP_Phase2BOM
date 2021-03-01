# -*- coding: utf-8 -*-
"""
Created on Fri Aug 21 15:42:28 2020

@author: eclark
"""

import pandas as pd
from lxml import etree
import copy
#import datetime
import re

class BOM:
    
    def __init__(self, BOM = etree.fromstring("<Parts Group='None'><Part ITEM_KEY='Junk'></Part></Parts>")):
        self.item_dtypes = {'Item': str, 'DESCRIPTION': str, 'TacTic Description': str, 'Revision': str, 'Revision Track': int, 'ECN': int, 'Drawing Number': str, 'Alternate Item': str, 'Buyer': str, 'Stocked': int, 'Show In Drop-Down Lists': int, 'U/M': str, 'Type': str, 'Source': str, 'Product Code': str, 'ABC Code': str, 'Cost Type': str, 'Cost Method': str, 'Unit Cost': str, 'Current Unit Cost': str, 'Lot Size': int, 'Unit Weight': int, 'Weight Units': str, 'Quantity On Hand': float, 'Non-Nettable Stock': float, 'Safety Stock': float, 'Quantity Ordered': float, 'Quantity WIP': float, 'Allocated To Prod': float, 'Allocated To Customer Orders': float, 'Reserved For Customer Orders': float, 'Low Level': int, 'Active for Data Integration': int, 'Planner Code': str, 'Shrink Factor': float, 'Phantom Flag': int, 'MPS Flag': int, 'Net Change': int, 'MPS Plan Fence': int, 'Family Code': str, 'Production Type': str, 'Rate/Day': float, 'Inventory LCL %': str, 'Inventory UCL %': str, 'Supply Site': str, 'Supply Whse': str, 'Paper Work': int, 'Fixed Lead Time': int, 'Expedited Fixed': int, 'Dock-to-Stock': int, 'Variable': int, 'Expedited Variable': int, 'Separation': str, 'Release 1': str, 'Release 2': str, 'Release 3': str, 'MRP Item': int, 'Infinite': int, 'Planned Mfg Supply Switching': int, 'Accept Requirements': int, 'Pass Requirements': int, 'Must use future POs before creating PLNs': int, 'Supply Usage Tolerance': int, 'Time Fence Rule': str, 'Time Fence Value': int, 'Pull-Up SS Rule': str, 'Pull-Up SS Value': int, 'Setup Group': str, 'Charge Item': str, 'Order Minimum': int, 'Order Multiple': int, 'Order Maximum': int, 'Days Supply': int, 'Use Reorder Point': int, 'Reorder Point': int, 'Fixed Order Qty': int, 'Earliest Planned Purchase Receipt': str, 'Targeted Safety Stock Replenishment': str, 'Lot Track': int, 'Preassign Lots': int, 'Lot Prefix': str, 'S/N Track': int, 'Preassign Serials': int, 'S/N Prefix': str, 'Shelf Life': str, 'Issue By': str, 'Material Status': str, 'Reason': str, 'Last Change': str, 'User': str, 'Backflush': int, 'Backflush Location': str, 'Preferred Co-product Mix': str, 'Reservable': int, 'Tax-Free Imported Material': int, 'Tax Free Days': int, 'Safety Stock Percent': int, 'Tariff Classification': str, 'PO Tolerance Over': str, 'PO Tolerance Under': str, 'Kit': int, 'Print Kit Components on Customer Paperwork': int} #, 'Std Due Period': str, 'Commodity': str, 'Commodity Description': str, 'Tax Code': str, 'Tax Code Description': str, 'Origin': str, 'Country': str, 'Preference Criterion': str, 'Country Of Origin': str, 'Producer': int, 'Subject To RVC Requirements': int, 'Purchased YTD': float, 'Manufactured YTD': float, 'Used YTD': float, 'Sold YTD': float, 'Subject To Excise Tax': int, 'Excise Tax Percent': float, 'Wholesale Price': float, 'Includes Item Content': int, 'Order Configurable': int, 'Job Configurable': int, 'Auto Job Generation': str, 'Name Space': str, 'Configuration Flag': int, 'Feature String': str, 'Feature Template': str, 'Last Import Date': str, 'Save Current Revision Upon Import': int, 'Overview': str, 'Active For Customer Portal': int, 'Featured Item': int, 'Top Seller': int, 'Item Attribute Group': str, 'Item Attribute Group Description': str, 'Lot Attribute Group': str, 'Lot Attribute Group Description': str, 'Enable Pieces Inventory': int, 'Piece Dimension Group': str, 'Piece Dimension Group Description': str, 'Portal Pricing Enabled': int, 'Portal Pricing Site': str, 'Freight': str, 'Estimated Break Date': str, 'Date of Last Report': str, 'Commodity Jurisdiction': str, 'ECCN or USML CAT': str, 'Program (ITAR/EAR600 Series)': str, 'Schedule B Number': str, 'HTS Code': str, 'HTS Code Description': str, 'Country Of Origin': str, 'Length Linear Dimension': str, 'Linear Dimension UM': str, 'Width Linear Dimension': str, 'Height Linear Dimension': str, 'Density': float, 'Density UM': str, 'Area': float, 'Area UM': str, 'Bulk Mass': float, 'Bulk Mass UM': str, 'Ream Mass': float, 'Ream Mass UM': str, 'Paper Mass Basis': str, 'Grade': str, 'Abnormal Size': int}
        self.items = pd.read_excel(r"\\TACTICFILE\Public\Everyone\ERP\Item Database (CSI Format) Pilot2.xlsx", sheet_name="Item Database (CSI Format)", converters=self.item_dtypes).fillna("")
        
#        self.dwgindex_dtypes = {'SIZE': str, 'DWG. NO.': str, 'TITLE': str, 'MODEL': str, 'DRN': str, 'DATE': str, 'PROJECT': str,'BOOK': str}
#        self.dwgindex = pd.read_excel("Copy of Drawing Index.xlsx", converters=self.dwgindex_dtypes).fillna("")
#        self.dwgindex.drop(self.dwgindex.loc[self.dwgindex['SIZE']==""].index, inplace=True)
#        self.dwgindex.drop(self.dwgindex.loc[self.dwgindex['SIZE']=="SIZE"].index, inplace=True)
#        self.dwgindex.loc[:, "TITLE"] = self.dwgindex.loc[:, "TITLE"] + " " + self.dwgindex.loc[:, "MODEL"]
#        self.dwgindex.drop(columns = "MODEL", inplace=True)
        
        self.dwgindex_dtypes = {'BOOK': str, 'DATE': str, 'DRN': str, 'Dwg Size': str, 'Dwg. #': str, 'PROJECT': str, 'Revision': str, 'TITLE': str}
        self.dwgindex = pd.read_excel(r"\\TACTICFILE\Public\Engineering\Reference Materials\Drawing Index_ERP.xlsx", dtype=str).fillna("")
        
        self.noninv_dtypes = {'Item': str, 'Description': str, 'Revision': str, 'Drawing Number': str, 'Buyer': str, 'Show In Drop-Down Lists': int, 'U/M': str, 'Type': str, 'Product Code': str, 'Family Code': str, 'Commodity': str, 'Country Of Origin': str, 'Subject To RVC Requirements': int, 'Preference Criterion': str, 'Producer': int, 'Unit Cost': float, 'Unit Weight': float, 'Weight Units': str, 'Unit Price': float, 'Length Linear Dimension': float, 'Linear Dimension UM': str, 'Width Linear Dimension': float, 'Height Linear Dimension': float, 'Density': float, 'Density UM': str, 'Area': float, 'Area UM': str, 'Bulk Mass': float, 'Bulk Mass UM': str, 'Ream Mass': float, 'Ream Mass UM': str, 'Paper Mass Basis': str, 'Grade': int, 'Abnormal Size': int, 'Unit Price 1': float, 'Unit Price 2': float, 'Unit Price 3': float, 'Unit Price 4': float, 'Unit Price 5': float, 'Unit Price 6': float, 'Break Qty 1': int, 'Base Code 1': str, 'Amount/Percent 1': str, 'Value 1': float, 'Calculated Unit Price 1': float, 'Break Qty 2': float, 'Base Code 2': str, 'Value 2': float, 'Amount/Percent 2': str, 'Calculated Unit Price 2': float, 'Break Qty 3': float, 'Base Code 3': str, 'Amount/Percent 3': str, 'Value 3': float, 'Calculated Unit Price 3': float, 'Break Qty 4': float, 'Base Code 4': str, 'Amount/Percent 4': str, 'Value 4': float, 'Calculated Unit Price 4': float, 'Break Qty 5': float, 'Base Code 5': str, 'Amount/Percent 5': str, 'Value 5': float, 'Calculated Unit Price 5': float}
        self.noninv = pd.read_excel(r"\\TACTICFILE\Public\Everyone\ERP\Non-InventoryItems.xlsx", converters=self.noninv_dtypes).fillna("")
        
        self.wc_dtypes = {'Work Center': str, 'Name': str, 'Department': str, 'Dept Description': str, 'Alt. Work Ctr': str, 'Alt. Work Ctr Name': str, 'Shift ID': str, 'Backflush': str, 'Overhead Basis': str, 'Schedule Driver': str, 'Efficiency': float, 'Queue Hours': float, 'Finish Hours': float, 'Control Point': int, 'Outside': int, 'Setup Rate': float, 'Run Rate (Lbr)': float, 'Fix Mach Ovhd Rate': float, 'Var Mach Ovhd Rate': float, 'Avg Setup Rate': float, 'Avg Run Rate (Lbr)': float, 'Cost Code': str, 'Fix Mach Ovhd Applied Acct': str, 'Fix Mach Ovhd Appl Acct Unit1': str, 'Fix Mach Ovhd Appl Acct Unit2': str, 'Fix Mach Ovhd Appl Acct Unit3': str, 'Fix Mach Ovhd Appl Acct Unit4': str, 'Fix Mach Ovhd Appl Acct Description': str, 'Var Mach Ovhd Applied Acct': str, 'Var Mach Ovhd Appl Acct Unit1': str, 'Var Mach Ovhd Appl Acct Unit2': str, 'Var Mach Ovhd Appl Acct Unit3': str, 'Var Mach Ovhd Appl Acct Unit4': str, 'Var Mach Ovhd Appl Acct Description': str, 'Material WIP Acct': str, 'Material WIP Acct Unit1': str, 'Material WIP Acct Unit2': str, 'Material WIP Acct Unit3': str, 'Material WIP Acct Unit4': str, 'Material WIP Acct Description': str, 'Labor WIP Acct': str, 'Labor WIP Acct Unit1': str, 'Labor WIP Acct Unit2': str, 'Labor WIP Acct Unit3': str, 'Labor WIP Acct Unit4': str, 'Labor WIP Acct Description': str, 'Fixed Ovhd WIP Acct': str, 'Fix Ovhd WIP Acct Unit1': str, 'Fix Ovhd WIP Acct Unit2': str, 'Fix Ovhd WIP Acct Unit3': str, 'Fix Ovhd WIP Acct Unit4': str, 'Fix Ovhd WIP Acct Description': str, 'Var Ovhd WIP Acct': str, 'Var Ovhd WIP Acct Unit1': str, 'Var Ovhd WIP Acct Unit2': str, 'Var Ovhd WIP Acct Unit3': str, 'Var Ovhd WIP Acct Unit4': str, 'Var Ovhd WIP Acct Description': str, 'Outside WIP Acct': str, 'Outside WIP Acct Unit1': str, 'Outside WIP Acct Unit2': str, 'Outside WIP Acct Unit3': str, 'Outside WIP Acct Unit4': str, 'Outside WIP Acct Description': str, 'Material Usage Variance': str, 'Material Usage Var Acct Unit1': str, 'Material Usage Var Acct Unit2': str, 'Material Usage Var Acct Unit3': str, 'Material Usage Var Acct Unit4': str, 'Material Usage Var Acct Description': str, 'Labor Rate Variance': str, 'Labor Rate Var Acct Unit1': str, 'Labor Rate Var Acct Unit2': str, 'Labor Rate Var Acct Unit3': str, 'Labor Rate Var Acct Unit4': str, 'Labor Rate Var Acct Description': str, 'Labor Usage Variance': str, 'Labor Usage Var Acct Unit1': str, 'Labor Usage Var Acct Unit2': str, 'Labor Usage Var Acct Unit3': str, 'Labor Usage Var Acct Unit4': str, 'Labor Usage Var Acct Description': str, 'Fix Matl Ovhd Usage': str, 'Fix Mat Ovhd Usage Acct Unit1': str, 'Fix Mat Ovhd Usage Acct Unit2': str, 'Fix Mat Ovhd Usage Acct Unit3': str, 'Fix Mat Ovhd Usage Acct Unit4': str, 'Fix Mat Ovhd Usage Acct Description': str, 'Var Matl Ovhd Usage': str, 'Var Mat Ovhd Usage Acct Unit1': str, 'Var Mat Ovhd Usage Acct Unit2': str, 'Var Mat Ovhd Usage Acct Unit3': str, 'Var Mat Ovhd Usage Acct Unit4': str, 'Var Mat Ovhd Usage Acct Description': str, 'Fix Labor Ovhd Usage': str, 'Fix Lbr Ovhd Usage Acct Unit1': str, 'Fix Lbr Ovhd Usage Acct Unit2': str, 'Fix Lbr Ovhd Usage Acct Unit3': str, 'Fix Lbr Ovhd Usage Acct Unit4': str, 'Fix Labor Ovhd Usage Description': str, 'Var Labor Ovhd Usage': str, 'Var Lbr Ovhd Usage Acct Unit1': str, 'Var Lbr Ovhd Usage Acct Unit2': str, 'Var Lbr Ovhd Usage Acct Unit3': str, 'Var Lbr Ovhd Usage Acct Unit4': str, 'Var Lbr Ovhd Usage Acct Description': str, 'Fix Mach Ovhd Usage': str, 'Fix Mach Ovhd Usage Acct Unit1': str, 'Fix Mach Ovhd Usage Acct Unit2': str, 'Fix Mach Ovhd Usage Acct Unit3': str, 'Fix Mach Ovhd Usage Acct Unit4': str, 'Fix Mach Ovhd Usage Acct Description': str, 'Var Mach Ovhd Usage': str, 'Var Mach Ovhd Usage Acct Unit1': str, 'Var Mach Ovhd Usage Acct Unit2': str, 'Var Mach Ovhd Usage Acct Unit3': str, 'Var Mach Ovhd Usage Acct Unit4': str, 'Var Mach Ovhd Usage Acct Description': str, 'Total Qty Queued': float, 'Total Setup Hours': float, 'Total Run Hours (Lbr)': float, 'Total Run Hours (Mach)': float, 'Avg Queue Hours': float, 'Material WIP': float, 'Labor WIP': float, 'Fix Ovhd WIP': float, 'Var Ovhd WIP': float, 'Outside WIP': float, 'Total WIP': float}
        self.wc = pd.read_excel(r"\\TACTICFILE\Public\Everyone\ERP\WorkCenters.xlsx", converters=self.wc_dtypes).fillna("")
        
        self.BOM = copy.deepcopy(BOM.getroot())
        item_cols = list(self.item_dtypes)
        self.__itemRecommendations = pd.DataFrame(columns = item_cols)
        
    @property
    def itemRecommendations(self):
        return self.__itemRecommendations
    
    def fetchItem(self, **kwargs):
        """ 
        Intended to be used for fetchItem(Item = "ItemID")
        but any column value can be used for the parameter variable.
        Function searches to match parameter as the column name from
        the items table loaded from excel.  Failure to match will be noted
        and recommended item table additions generated.
        Recommended additions to the item table will assume that it was used
        to search for an ItemID until the time comes to upgrade this function
        """
        # Error checking inputs
        if len(kwargs) != 1:
            raise AssertionError("Too many parameters")
        for key in kwargs.keys():
            if not(key in self.item_dtypes.keys()):
                raise NameError(key, " not found in item header")
        
        key = list(kwargs.keys())[0]
        value = list(kwargs.values())[0]
        
        # Get results from items table
        partItemDef = self.items.loc[self.items[key] == value]
        partNonInvDef = self.noninv.loc[self.noninv[key] == value]
        if len(partItemDef)==1 or len(partNonInvDef) > 0:
            # resolve if inventory or non-inventory and return item definition
            if len(partNonInvDef) > 0: #if multiple hits, first one wins
                fakeInvDef = {'Item': partNonInvDef.iloc[0]["Item"], 'DESCRIPTION': partNonInvDef.iloc[0]["Description"], 'TacTic Description': partNonInvDef.iloc[0]["Description"], 'Revision': partNonInvDef.iloc[0]["Revision"], 'Revision Track': 0, 'ECN': 0, 'Drawing Number': partNonInvDef.iloc[0]["Drawing Number"], 'Alternate Item': "", 'Buyer': partNonInvDef.iloc[0]["Buyer"], 'Stocked': 1, 'Show In Drop-Down Lists': partNonInvDef.iloc[0]["Show In Drop-Down Lists"], 'U/M': partNonInvDef.iloc[0]["U/M"], 'Type': partNonInvDef.iloc[0]["Type"], 'Source': "Purchased", 'Product Code': partNonInvDef.iloc[0]["Product Code"], 'ABC Code': "C", 'Cost Type': "Standard", 'Cost Method': "Standard", 'Unit Cost': partNonInvDef.iloc[0]["Unit Cost"], 'Current Unit Cost': partNonInvDef.iloc[0]["Unit Cost"], 'Lot Size': 1, 'Unit Weight': partNonInvDef.iloc[0]["Unit Weight"], 'Weight Units': partNonInvDef.iloc[0]["Weight Units"], 'Quantity On Hand': 123456789, 'Non-Nettable Stock': 123456789, 'Safety Stock': 123456789, 'Quantity Ordered': 123456789, 'Quantity WIP': 123456789, 'Allocated To Prod': 123456789, 'Allocated To Customer Orders': 123456789, 'Reserved For Customer Orders': 123456789, 'Low Level': 20, 'Active for Data Integration': 1, 'Planner Code': "Non-Inv", 'Shrink Factor': 0, 'Phantom Flag': 0, 'MPS Flag': 0, 'Net Change': 0, 'MPS Plan Fence': 0, 'Family Code': "", 'Production Type': "Job", 'Rate/Day': 1, 'Inventory LCL %': "", 'Inventory UCL %': "", 'Supply Site': "", 'Supply Whse': "", 'Paper Work': 0, 'Fixed Lead Time': 0, 'Expedited Fixed': 0, 'Dock-to-Stock': 0, 'Variable': 0, 'Expedited Variable': 0, 'Separation': "", 'Release 1': "", 'Release 2': "", 'Release 3': "", 'MRP Item': 0, 'Infinite': 1, 'Planned Mfg Supply Switching': 0, 'Accept Requirements': 1, 'Pass Requirements': 1, 'Must use future POs before creating PLNs': 0, 'Supply Usage Tolerance': 0, 'Time Fence Rule': "", 'Time Fence Value': 0, 'Pull-Up SS Rule': "", 'Pull-Up SS Value': 0, 'Setup Group': "", 'Charge Item': "", 'Order Minimum': 0, 'Order Multiple': 0, 'Order Maximum': 0, 'Days Supply': 0, 'Use Reorder Point': 0, 'Reorder Point': 0, 'Fixed Order Qty': 0, 'Earliest Planned Purchase Receipt': "", 'Targeted Safety Stock Replenishment': "", 'Lot Track': 0, 'Preassign Lots': 0, 'Lot Prefix': "", 'S/N Track': 0, 'Preassign Serials': 0, 'S/N Prefix': "", 'Shelf Life': "", 'Issue By': "", 'Material Status': "", 'Reason': "", 'Last Change': "", 'User': "", 'Backflush': 1, 'Backflush Location': "", 'Preferred Co-product Mix': "", 'Reservable': 0, 'Tax-Free Imported Material': 0, 'Tax Free Days': 0, 'Safety Stock Percent': 0, 'Tariff Classification': "", 'PO Tolerance Over': "", 'PO Tolerance Under': "", 'Kit': 0, 'Print Kit Components on Customer Paperwork': 0, 'Std Due Period': "", 'Commodity': "", 'Commodity Description': "", 'Tax Code': "", 'Tax Code Description': "", 'Origin': "", 'Country': "", 'Preference Criterion': "", 'Country Of Origin': "", 'Producer': 0, 'Subject To RVC Requirements': 0, 'Purchased YTD': 123456789, 'Manufactured YTD': 123456789, 'Used YTD': 123456789, 'Sold YTD': 123456789, 'Subject To Excise Tax': 0, 'Excise Tax Percent': 123456789, 'Wholesale Price': 123456789, 'Includes Item Content': 0, 'Order Configurable': 0, 'Job Configurable': 0, 'Auto Job Generation': "", 'Name Space': "", 'Configuration Flag': 0, 'Feature String': "", 'Feature Template': "", 'Last Import Date': "", 'Save Current Revision Upon Import': 0, 'Overview': "", 'Active For Customer Portal': 0, 'Featured Item': 0, 'Top Seller': 0, 'Item Attribute Group': "", 'Item Attribute Group Description': "", 'Lot Attribute Group': "", 'Lot Attribute Group Description': "", 'Enable Pieces Inventory': 0, 'Piece Dimension Group': "", 'Piece Dimension Group Description': "", 'Portal Pricing Enabled': 0, 'Portal Pricing Site': "", 'Freight': "", 'Estimated Break Date': "", 'Date of Last Report': "", 'Commodity Jurisdiction': "", 'ECCN or USML CAT': "", 'Program (ITAR/EAR600 Series)': "", 'Schedule B Number': "", 'HTS Code': "", 'HTS Code Description': "", 'Country Of Origin': "", 'Length Linear Dimension': "", 'Linear Dimension UM': "", 'Width Linear Dimension': "", 'Height Linear Dimension': "", 'Density': 123456789, 'Density UM': "", 'Area': 123456789, 'Area UM': "", 'Bulk Mass': 123456789, 'Bulk Mass UM': "", 'Ream Mass': 123456789, 'Ream Mass UM': "", 'Paper Mass Basis': "", 'Grade': "", 'Abnormal Size': 0}
                return pd.Series(fakeInvDef)
            else:
                 return partItemDef.iloc[0]
        elif len(partItemDef) == 0:            
            # part not found, make an entry into the recommended additions to the item table
            recExists = self.__itemRecommendations.loc[self.__itemRecommendations[key] == value]
            if len(recExists) == 0:
                print(str(key) + " == " + str(value) + " not found in items table.  Recommended Item added to itemRecommendations attribute")
                recommendedItem = self.makeItemRecommendation(**kwargs)
                self.__itemRecommendations = self.__itemRecommendations.append(recommendedItem, ignore_index = True)
                return recommendedItem
            else:
                print(str(key) + " == " + str(value) + " already recommended as missing")
                return recExists.iloc[0]
        else:
            print("multiple rows match " + str(key) + " == " + str(value))
            return partItemDef.iloc[0]
    
    def makeItemRecommendation(self, **kwargs):
        #default items table row definition

        # Error checking inputs
        if len(kwargs) != 1:
            raise AssertionError("Too many parameters")
        for key in kwargs.keys():
            if not(key in self.item_dtypes.keys()):
                raise NameError(key, " not found in item header")
        
        key = list(kwargs.keys())[0]
        value = list(kwargs.values())[0]
        
        recommendation = {'Item': "", 'DESCRIPTION': "", 'TacTic Description': "", 'Revision': "", 'Revision Track': 0, 'ECN': 0,
                          'Drawing Number': "", 'Alternate Item': "", 'Buyer': "", 'Stocked': 1, 'Show In Drop-Down Lists': 1,
                          'U/M': "EA", 'Type': "Material", 'Source': "Manufactured", 'Product Code': "__Unknown__", 'ABC Code': "C",
                          'Cost Type': "Actual", 'Cost Method': "FIFO", 'Unit Cost': "", 'Current Unit Cost': "", 'Lot Size': 1,
                          'Unit Weight': 0, 'Weight Units': "", 'Quantity On Hand': 0.0, 'Non-Nettable Stock': 0.0, 'Safety Stock': 0.0,
                          'Quantity Ordered': 0.0, 'Quantity WIP': 0.0, 'Allocated To Prod': 0.0, 'Allocated To Customer Orders': 0.0,
                          'Reserved For Customer Orders': 0.0, 'Low Level': 0, 'Active for Data Integration': 1, 'Planner Code': "",
                          'Shrink Factor': 0.0, 'Phantom Flag': 0, 'MPS Flag': 0, 'Net Change': 0, 'MPS Plan Fence': 0, 'Family Code': "",
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
                          'Material Status': "Active", 'Reason': "", 'Last Change': "", 'User': "", 'Backflush': 1,
                          'Backflush Location': "", 'Preferred Co-product Mix': "", 'Reservable': 0, 'Tax-Free Imported Material': 0,
                          'Tax Free Days': 0, 'Safety Stock Percent': 0, 'Tariff Classification': "", 'PO Tolerance Over': "",
                          'PO Tolerance Under': "", 'Kit': 0, 'Print Kit Components on Customer Paperwork': 0}
#                           , 'Std Due Period': "",
#                          'Commodity': "", 'Commodity Description': "", 'Tax Code': "N", 'Tax Code Description': "Not Taxable",
#                          'Origin': "", 'Country': "", 'Preference Criterion': "", 'Country Of Origin': "", 'Producer': 0,
#                          'Subject To RVC Requirements': 0, 'Purchased YTD': 0.0, 'Manufactured YTD': 0.0, 'Used YTD': 0.0,
#                          'Sold YTD': 0.0, 'Subject To Excise Tax': 0, 'Excise Tax Percent': 0, 'Wholesale Price': 0,
#                          'Includes Item Content': 0, 'Order Configurable': 0, 'Job Configurable': 0, 'Auto Job Generation': "Never",
#                          'Name Space': "", 'Configuration Flag': 0, 'Feature String': "", 'Feature Template': "",
#                          'Last Import Date': "", 'Save Current Revision Upon Import': 0, 'Overview': "", 'Active For Customer Portal': 0,
#                          'Featured Item': 0, 'Top Seller': 0, 'Item Attribute Group': "", 'Item Attribute Group Description': "",
#                          'Lot Attribute Group': "", 'Lot Attribute Group Description': "", 'Enable Pieces Inventory': 0,
#                          'Piece Dimension Group': "", 'Piece Dimension Group Description': "", 'Portal Pricing Enabled': 0,
#                          'Portal Pricing Site': "", 'Freight': "", 'Estimated Break Date': "", 'Date of Last Report': "",
#                          'Commodity Jurisdiction': "", 'ECCN or USML CAT': "", 'Program (ITAR/EAR600 Series)': "", 'Schedule B Number': "",
#                          'HTS Code': "", 'HTS Code Description': "", 'Country Of Origin': "", 'Length Linear Dimension': "",
#                          'Linear Dimension UM': "", 'Width Linear Dimension': "", 'Height Linear Dimension': "", 'Density': 0.0,
#                          'Density UM': "", 'Area': 0.0, 'Area UM': "", 'Bulk Mass': 0.0, 'Bulk Mass UM': "", 'Ream Mass': 0.0,
#                          'Ream Mass UM': "", 'Paper Mass Basis': "None", 'Grade': "", 'Abnormal Size': 0}
        if key == "Item":
            itemSearch = value
            itemIsMfg = self.BOM.find(".//Part[@PartID='{0}']".format(str(itemSearch)))
            recommendation['Product Code'] = "TT-MfgStep" if itemIsMfg is not None else "TT-Raw"
            itemIsDWG = re.fullmatch("80(\d{5})(.*)", itemSearch)
            itemIsHardware = re.fullmatch("[12](\d{2})(\d{2})(\d{2})(.*)", itemSearch)
            itemIsWire = re.fullmatch("787(?P<type>\d)(?P<color>\d)(?P<tracer>\d)(?P<size>[B-NP-Z])(?P<dash>-.+)?",itemSearch)
            itemIsCable = re.fullmatch("788(?P<type>[0-3])(?P<cond>\d{2})(?P<size>[B-NP-Z])(?P<dash>-.+)?",itemSearch)
            if itemIsDWG:
                dwgInfo = self.fetchDrawing(itemIsDWG[1])
                if type(dwgInfo) == pd.Series:
                    recommendation['Item'] = itemSearch
                    if len(dwgInfo['TITLE']) > 40:
                        recommendation['DESCRIPTION'] = dwgInfo['TITLE'][:39] + "\u2026"
                    else:
                        recommendation['DESCRIPTION'] = dwgInfo['TITLE']
                    recommendation['TacTic Description'] = dwgInfo['TITLE']
                    recommendation['Drawing Number'] = dwgInfo['Dwg. #']
                    recommendation['Revision'] = dwgInfo['Revision']
                else:
                    recommendation['Item'] = itemSearch
                    recommendation['DESCRIPTION'] = "Failed Drawing Lookup"
                recommendation['Phantom Flag'] = 1 if itemIsDWG[2][:2] == "-X" else 0
                recommendation['Product Code'] = "TT-MfgStep" if itemIsMfg is not None else "TT-NonSale"
            elif itemIsHardware:
                hardwareInfo = self.fetchHardware(itemIsHardware)
                recommendation['Item'] = itemSearch
                recommendation['DESCRIPTION'] = hardwareInfo if len(hardwareInfo) < 41 else hardwareInfo[:39] + "\u2026"
                recommendation['TacTic Description'] = hardwareInfo
                recommendation['Source'] = "Purchased"
            elif itemIsWire:
                wireInfo = self.fetchWire(itemIsWire)
                recommendation['Item'] = itemSearch
                recommendation['DESCRIPTION'] = wireInfo if len(wireInfo) < 41 else wireInfo[:39] + "\u2026"
                recommendation['TacTic Description'] = wireInfo
                recommendation['Source'] = "Purchased"
                recommendation['U/M'] = "IN"
            elif itemIsCable:
                cableInfo = self.fetchCable(itemIsCable)
                recommendation['Item'] = itemSearch
                recommendation['DESCRIPTION'] = cableInfo if len(cableInfo) < 41 else cableInfo[:39] + "\u2026"
                recommendation['TacTic Description'] = cableInfo
                recommendation['Source'] = "Purchased"
                recommendation['U/M'] = "IN"
            else:
                recommendation['Item'] = itemSearch
        else:
            recommendation['Item'] = "Unknown"
            recommendation['DESCRIPTION'] = "Unknown"
            recommendation['TacTic Description'] = "Unknown"
        return pd.Series(recommendation)
    
    def fetchDrawing(self, DWG_NO = "12345"):
        dwgDefs = self.dwgindex.loc[self.dwgindex['Dwg. #'] == DWG_NO]
        if len(dwgDefs) == 1:
            return dwgDefs.iloc[0]
        elif len(dwgDefs) == 0:
            print("drawing {} not in index".format(DWG_NO))
            return None
        else:
            print("drawing {} overdefined in index".format(DWG_NO))
            return None

    def fetchHardware(self, hardwareMatchResult):
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
        if hardwareMatchResult[1] in HardwareType.keys():
            hType = HardwareType[hardwareMatchResult[1]]
        else:
            hType = "Unknown Type"

        if hardwareMatchResult[2] in HardwareSize.keys():
            hSize = HardwareSize[hardwareMatchResult[2]]
        else:
            hSize = "Unknown Size"

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
        
        if hardwareMatchResult[0][0] == "1":
            stainless = ", St. Stl."
        else:
            stainless = ""

        if hardwareMatchResult[4] == "":
            special = ""
        else:
            special = ",  something special here"

        return hType + ", " + hSize + hLen + stainless + special
    
    def buildItem(self, PartID):
        OpColumns = ["Item", "Item Description", "Alternate ID", "Description", "Operation", "Shared", "WC", "WC Description", "Use Fixed Schedule", "Fixed Sched Hours", "Run-Hours Basis (Machine)", "Mach Hrs per Piece", "Run-Hours Basis (Labor)", "Labor Hr per Piece", "Sched Driver", "Run Duration", "Batch Definition", "Yield", "Seconds Per Cycle", "Formula Material Weight", "Formula Material Weight U/M", "Move Hours", "Queue Time", "Setup Hours", "Finish", "Use Offset Hours", "Offset Hours", "Effective Date", "Obsolete Date", "Control Point", "Backflush", "Setup Resource Group", "Setup Rule", "Setup Basis", "Setup Time Rule", "Setup Matrix", "Scheduler Rule", "Custom Planner Rule", "Break Rule", "Split Rule", "Split Size", "Split Group", "Efficiency", "Setup Rate", "Run Rate (Labor)", "Var Mach Ovhd Rate", "Fix Machine Ovhd Rate", "Var Ovhd Rate", "Fixed Ovhd Rate"]
        opTemplate = {"Alternate ID": "", "Description": "",
                      "Shared": "", "Use Fixed Schedule": 0, "Fixed Sched Hours": "",
                      "Run-Hours Basis (Machine)": "Mch Hours/Pc", "Mach Hrs per Piece": 0.5,
                      "Run-Hours Basis (Labor)": "Lbr Hours/Pc", "Labor Hr per Piece": 0,
                      "Sched Driver": "Machine", "Run Duration": 0.5, "Batch Definition": "",
                      "Yield": 100, "Seconds Per Cycle": 0, "Formula Material Weight": 0,
                      "Formula Material Weight U/M": "", "Move Hours": 0,
                      "Queue Time": 0, "Setup Hours": 0, "Finish": 0, "Use Offset Hours": "",
                      "Offset Hours": 0, "Effective Date": "", "Obsolete Date": "",
                      "Control Point": 1, "Backflush": "Neither", "Setup Resource Group": "",
                      "Setup Rule": "Always", "Setup Basis": "Item", "Setup Time Rule": "Fixed Time",
                      "Setup Matrix": "", "Scheduler Rule": "Per Piece", "Custom Planner Rule": 0,
                      "Break Rule": "Shifts", "Split Rule": "No Splitting",
                      "Split Size": 0, "Split Group": "", "Efficiency": 100,
                      "Setup Rate": "", "Run Rate (Labor)": 35, "Var Mach Ovhd Rate": 35,
                      "Fix Machine Ovhd Rate": 0, "Var Ovhd Rate": 0, "Fixed Ovhd Rate": 0}
        MatColumns = ["Item", "Item Description", "Alternate ID", "Description", "Operation", "Shared", "WC", "WC Description", "Material", "Material Description", "Seq", "Alt Group", "Alt Group Rank", "Manufacturer", "Manufacturer Name", "Manufacturer Item", "Manufacturer Item Description", "Type", "Quantity", "Per", "U/M", "Cost", "Scrap Factor", "Effective Date", "Obsolete Date", "BOM Seq", "Ref", "Backflush", "Backflush Location", "Feature", "Option Code", "Probable", "Incremental Price", "Formula Material Weight %", "Estimated Break Date", "Date of Last Report", "Fixed Material", "Variable Material", "Material Cost", "Labor Cost", "Outside Cost", "Fixed Overhead Cost", "Variable Overhead Cost"]
        matTemplate = {"Alternate ID": "", "Description": "", "Shared": "",
                       "Manufacturer": "", "Manufacturer Name": "",
                       "Manufacturer Item": "", "Manufacturer Item Description": "",
                       "Alt Group Rank": 0, "Type": 'Material', "Per": 'Unit',
                       "Cost": 0, "Scrap Factor": 0, "Effective Date": "", 
                       "Obsolete Date": "", "BOM Seq": "", "Ref": '', 
                       "Backflush": 1, "Backflush Location": "", "Feature": "",
                       "Option Code": "", "Probable": 1, "Incremental Price": 0,
                       "Formula Material Weight %": 0, "Estimated Break Date": "",
                       "Date of Last Report": "", "Fixed Material": 0,
                       "Variable Material": 0, "Material Cost": 0,
                       "Labor Cost": 0, "Outside Cost": 0,
                       "Fixed Overhead Cost": 0, "Variable Overhead Cost": 0}
        
        itemDef = self.fetchItem(Item = PartID)
        matTemplate["Item"] = itemDef["Item"]
        opTemplate["Item"] = itemDef["Item"]
        matTemplate["Item Description"] = itemDef["DESCRIPTION"]
        opTemplate["Item Description"] = itemDef["DESCRIPTION"]
        
        Operations = pd.DataFrame(columns = OpColumns)
        Materials = pd.DataFrame(columns = MatColumns)
        
        buildItem = self.BOM.find(".//Part[@PartID='{0}']".format(str(PartID)))
        opID = 10
        partOPs = buildItem.findall(".//Operation")
        for op in partOPs:
            wcRow = self.wc.loc[self.wc['Work Center']  == op.attrib['WC']].iloc[0]
            thisOp = {"Item": buildItem.attrib["PartID"], 
                      "WC": op.attrib['WC'],
                      "WC Description": wcRow['Name'],
                      "Operation": opID}
            thisOp.update(opTemplate)
            Operations = Operations.append(thisOp, ignore_index = True)
            matList = []
            for mat in op.findall(".//Material"):
                matList.extend(self.genMatReturns(mat))
            seqID = 1
            for matRow in matList:
                thisMat = {"Item": buildItem.attrib['PartID'],
                           "Operation": opID,
                           "WC": op.attrib['WC'],
                           "WC Description": wcRow['Name'],
                           "Material": matRow['PartID'],
                           "Material Description": matRow['Material Description'],
                           "Seq": seqID,
                           "Alt Group": seqID,
                           "Alt Group Rank": 0,
                           "Quantity": matRow['Quantity'],
                           "U/M": matRow['Unit']}
                thisMat.update(matTemplate)
                Materials = Materials.append(thisMat, ignore_index = True)
                seqID += 1
            opID += 10
        return (Operations, Materials)

    def genMatReturns(self, target = etree.fromstring("<Material PartID='12345' Unit='EA' Qty='1' />"), quantity = 1):
        results = []
        reportQty = 0
        
        try:
            reportQty = int(target.attrib['Qty']) * quantity
        except ValueError:
            reportQty = float(target.attrib['Qty']) * quantity
        itemDef = self.fetchItem(Item = target.attrib['PartID'])
        if 'Size' not in target.attrib:
            results = [{'PartID': target.attrib['PartID'],
                        'Material Description': itemDef['DESCRIPTION'],
                        'Unit': target.attrib['Unit'],
                        'Quantity': reportQty }]
        else:
            result = {'PartID': target.attrib['PartID'],
                      'Material Description': itemDef['DESCRIPTION'],
                      'Unit': target.attrib['Unit'],
                      'Quantity': float(target.attrib['Size']) }
            for i in range(reportQty):
                results.extend([result])
        return results

    def fetchWire(self, wireMatchResult):
        digit4 = {'0': "Bare", '1': "Solid", '2': "Stranded", '5': "Enamel", '6': "Litz", '9': "Misc"}
        color = {'0': "BLK", '1': "BRN", '2': "RED", '3': "ORG", '4': "YEL", 
                 '5': "GRN", '6': "BLU", '7': "PUR", '8': "GRY", '9': "WHT"}
        size = {'B': "40", 'C': "38", 'D': "36", 'E': "34", 'F': "32", 'G': "30", 
                'H': "28", 'I': "26", 'J': "24", 'K': "22", 'L': "20", 'M': "18", 
                'N': "16", 'P': "14", 'Q': "12", 'R': "10", 'S': "8", 'T': "6", 
                'U': "4", 'V': "2", 'W': "0", 'X': "00", 'Y': "000", 'Z': "0000"}
        itemDict = {'wt': digit4[wireMatchResult['type']] if wireMatchResult['type'] in digit4.keys() else "Unknown wire type", 
                    'color': color[wireMatchResult['color']],
                    'tracer': color[wireMatchResult['tracer']],
                    'size': size[wireMatchResult['size']],
                    'dash': ", {} __special wire__".format(wireMatchResult['dash']) if wireMatchResult['dash'] else ""}
        if wireMatchResult['color'] == wireMatchResult['tracer']:
            return "Wire, Single, {wt}, {color}, {size}{dash}".format(**itemDict)
        else:
            return "Wire, Single, {wt}, {color} w/ {tracer}, {size}{dash}".format(**itemDict)

    def fetchCable(self, cableMatchResult):
        digit4 = {'0': "Shld Sld",
                  '1': "Shld Str",
                  '2': "Unshld Sld",
                  '3': "Unshld Str"}
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
        modifier = {None: "",
                    '-1': ", Rubber Jacket",
                    '-2': ", PUR Jacket",
                    '-3': ", PTFE Jacket"}
        itemDict = {'ct': digit4[cableMatchResult['type']] if cableMatchResult['type'] in digit4.keys() else "Unknown cable type",
                    'nCond': nCond,
                    'size': size[cableMatchResult['size']],
                    'dash': modifier[cableMatchResult['dash']] if cableMatchResult['dash'] in modifier.keys() else ", Unknown dash type"}
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





























