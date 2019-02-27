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
        self.options, self.arguments = getopt.getopt(someArguments, "s:hlfg")

    def process(self):
        self.indexesToProcess = None  # by default extract all worksheets
        self.listMode = False
        self.formulaMode = False
        self.absoluteFormulaMode = False

        for key, value in self.options:
            if key == "-h":
                self.usage()
                exit()
            elif key == "-s":
                self.setWorksheetsToProcessFromOptionValue(value)    
            elif key == "-l":
                self.listMode = True
            elif key == "-f":
                self.formulaMode = True
            elif key == "-g":
                self.formulaMode = True
                self.absoluteFormulaMode = True
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
        
    def usage(self):
        print "tabby.py: extract excel data"
        print "usage: tabby.py [-h] [-s num] file [ file...]"
        print "  -h - help"
        print "  -l - list of key value pairs"
        print "  -s num (integer)"
        print "    extract worksheet with index num,"
        print "    e.g -s 2 extracts the second workseet"
        print "  -f use formulas"
        print "  -g use formulas (converted to absolute references)"
        
    def setWorksheetsToProcessFromOptionValue(self, optionValue):
        self.indexesToProcess = [ int(optionValue) ]

    def processFile(self, aFilename):
        reader = tableparser.FileReaderInterface(aFilename).getFileReader()
        reader.setFormulaMode(self.formulaMode)
        self.workbook = reader.getWorkbook()
        indexes = self.indexesToProcess
        if indexes is None:
            indexes = range(1, len(self.workbook.sheets) + 1)
        #print "indexes", indexes
        for sheetIndex in indexes:
            sheet = self.getSheetWithIndex(sheetIndex)
            if sheet is None:
                exit("worksheet", str(each), "not found")
            else:
                self.outputSheet(sheetIndex, sheet)
        self.counter += 1
    
    # The index is based on one and not on zero.
    def getSheetWithIndex(self, index):
        sheets = self.workbook.sheets
        if 1 <= index <= len(sheets):
            return sheets[index-1]       
        else:
            return None

    def outputSheet(self, sheetIndex, aSheet):
        if self.listMode:
            self.listSheet(sheetIndex, aSheet)
        else:
            self.printSheet(aSheet)
            
    def listSheet(self, sheetNumber, aSheet):
        for rowNumber, each in enumerate(aSheet.records):
            self.listRecord(sheetNumber, rowNumber, each)
            
    def listRecord(self, sheetNumber, rowNumber, aRecord):
        for columnNumber, each in enumerate(aRecord):
            self.listCell(sheetNumber, rowNumber, columnNumber, each)
            
    def listCell(self, sheetNumber, rowNumber, columnNumber, aCell):
        prefix = str(sheetNumber) + "." + self.excelReference(rowNumber+1, columnNumber+1) +  ":"
        print prefix, aCell
        
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
            
    # Convert given row and column number to an Excel-style cell name.
    def excelReference(self, row, column):
        quot, rem = divmod(column - 1, 26)
        return((chr(quot-1 + ord('A')) if quot else '') +
               (chr(rem + ord('A')) + str(row)))
            
