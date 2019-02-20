#!/usr/bin/env python
#
# requires python 2.7
#
# lib/tableprinter.py - tabular listing of a table file
#

import sys, getopt, glob, string
import tableparser

def exit(*args):
    if len(args):
        print "tableprinter.py:",
        for each in args:
            print  each,
        print 
    sys.exit(2)

class Formatter:
    def __init__(self, someArguments):
        self.options, self.arguments = getopt.getopt(someArguments, "s:hl")

    def usage(self):
        print "tabby.py: extract excel data"
        print "usage: tabby.py [-h] [-s num] file [ file...]"
        print "  -h - help"
        print "  -l - list of key value pairs"
        print "  -s num (integer)"
        print "    extract worksheet with index num,"
        print "    e.g -s 2 extracts the second workseet"

    def setWorksheetsToExtractFromOptionValue(self, optionValue):
        self.indexesToExtract = [ int(optionValue) ]

    def processFile(self, aFilename):
        self.workbook = tableparser.FileReaderInterface(aFilename).getWorkbook()
        indexes = self.indexesToExtract
        if indexes is None:
            indexes = range(1, len(self.workbook.sheets) + 1)
        #print "indexes", indexes
        for each in indexes:
            sheet = self.getSheetWithIndex(each)
            if sheet is None:
                exit("worksheet", str(each), "not found")
            else:
                self.outputSheet(sheet)
        self.counter += 1
    
    # The index is based on one and not on zero.
    def getSheetWithIndex(self, index):
        sheets = self.workbook.sheets
        if 1 <= index <= len(sheets):
            return sheets[index-1]       
        else:
            return None

    def outputSheet(self, aSheet):
        if self.listMode:
            self.listSheet(aSheet)
        else:
            self.printSheet(aSheet)
            
    def listSheet(self, aSheet):
        for each in aSheet.records:
            self.listRecord(each)
            
    def listRecord(self, aRecord):
        for each in aRecord:
            self.listCell(each)
            
    def listCell(self, aCell):
        print aCell
        
    #
    # A sheet is a list of records.
    #
    def printSheet(self, aSheet):
        maxRecordLength = 0
        for each in aSheet.records:
            maxRecordLength = max(maxRecordLength, len(each))
        widths = [0] * maxRecordLength      # array of zeros [0,0,0,0 ...]
        for each in aSheet.records:
            for index, value in enumerate(each):
                widths[index] = max(widths[index], len(value))
        for each in aSheet.records:
            for index, value in enumerate(each):
                alphabeticCharacters = filter(lambda ch: ch.isalpha(), value)
                if alphabeticCharacters == "":
                    print value.rjust(widths[index]),
                else:
                    print value.ljust(widths[index]),
            print
            
    def process(self):
        self.indexesToExtract = None  # by default extract all worksheets
        self.listMode = False

        for key, value in self.options:
            if key == "-h":
                self.usage()
                exit()
            elif key == "-s":
                self.setWorksheetsToExtractFromOptionValue(value)    
            elif key == "-l":
                self.listMode = True
            else:
                exit("unsupported option [%s]" % key)

        if self.arguments == []:
            exit("no arguments found")
            
        self.counter = 0
        for pattern in self.arguments:
            for file in glob.glob(pattern):
                self.processFile(file)
        if self.counter == 0:
            exit("no files found")

