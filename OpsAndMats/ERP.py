# -*- coding: utf-8 -*-
"""
Created on Mon May 17 13:28:36 2021

@author: eclark
"""

import argparse
import re
import os

from TACTIC_ERP import TACTIC_ERP
from MakeJobWithClass import makeThisJob

if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("--xml", action="store", type=str, help="the XML file containing a current BOM")
    parser.add_argument("--job", action="store", type=str, help="run the job creation tool")
    jobValidator = re.compile("J?(\d{1,9})-(\d{1,4})")
    args = parser.parse_args()
    
    if not (args.xml or args.job):
        # for those running the program by double click instead of proper CLI
        print("Select operation mode:")
        print("1: Build BOM from XML")
        print("2: Build Job BOM from Job Number")
        mode = input("Mode: ")
        if mode not in ["1", "2"]:
            print("improper response, terminating")
        elif mode[0] == "1":
            print("include the file extension (.xml) in your input...")
            args.xml = input ("What file do you want to process? ")
        elif mode[0] == "2":
            print("Job identifiers are two integers separated by a dash.")
            print("The ERP is stupid.  We are smart. We will add the J and zeros for you!")
            args.job = input("The job to generate: ")
    
    
    if args.xml:
        if not os.path.isfile(args.xml):
            print("file {} does not exist in this directory".format(args.xml))
        else:
            engBOM = TACTIC_ERP(args.xml)
            engBOM.parseCurrentBOM(os.path.splitext(args.xml)[0])

    if args.job:
        jobRegex = jobValidator.fullmatch(args.job)
        if not jobRegex:
            print("invalid job requested")
        else:
            print("Processing J{:09d}-{:04d}".format(int(jobRegex[1]),int(jobRegex[2])))
            makeThisJob("J{:09d}".format(int(jobRegex[1])),"{:04d}".format(int(jobRegex[2])))


    input("press enter to exit...") #keep window open for the user