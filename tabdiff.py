#!/usr/bin/env python
#
# requires python 2.7
#
# tabdiff.py - tabular listing of a table file
#

import sys, getopt, glob
import string
import lib.tabbybase
import lib.tabbyparser

class Differ(lib.tabbybase.TabbyBaseTool):
    def __init__(self, someArguments):
        lib.tabbybase.TabbyBaseTool.__init__(self)
        self.options, self.arguments = getopt.getopt(someArguments, "s:hfe:")

    def process(self):
        self.maxNumberOfDifferences = None
        self.resolveFormulas = False
        self.workbook1 = None
        self.workbook2 = None
        for key, value in self.options:
            if key == "-h":
                 self.usage()
                 exit()
            elif key == "-e":
                self.maxNumberOfDifferences = int(value)
            elif key == "-s":
                self.setWorksheetsToProcessFromOptionValue(value)
            elif key == "-f":
                self.resolveFormulas = True
            else:
                exit("unsupported option [%s]" % key)
        if self.arguments == []:
            exit("no arguments found")
        for pattern in self.arguments:
            for file in glob.glob(pattern):
                self.processFile(file)
        self.compare()
    
    def usage(self):
        print "tabdiff.py"
        print "-e num  - limit the number of reported errors"
        print "-f      - resolve formulas"
        print "-h      - help"
        print "-s num  - sheet numbers to compare"
        print "-s 1,3,6-8"
        
    def processFile(self, aFilename):
        if self.workbook1 is None:
            self.workbook1 = self.readFile(aFilename)
            return
        if self.workbook2 is None:
            self.workbook2 = self.readFile(aFilename)
            return
        exit("Too many files")

    def readFile(self, aFilename):
        reader = lib.tabbyparser.FileReaderInterface(aFilename).getFileReader()
        reader.resolveFormulas = self.resolveFormulas
        return reader.getWorkbook()

    def compare(self):
        self.numberOfDifferences = 0
        if self.workbook1 is None or self.workbook2 is None:
            exit("two files required")
        numberOfSheets1 = len(self.workbook1.sheets)
        numberOfSheets2 = len(self.workbook2.sheets)
        if numberOfSheets1 != numberOfSheets2:
            print "Number of worksheets is different"
        numberOfSheets = min(numberOfSheets1, numberOfSheets2)
        indexes = self.indexesToProcess
        if indexes == []:
            indexes = range(1, numberOfSheets + 1)
        print indexes
        for each in indexes:
            sheet1 = self.workbook1.getSheetWithIndex(each)
            sheet2 = self.workbook2.getSheetWithIndex(each)
            self.compareSheets(each, sheet1, sheet2)

    def compareSheets(self, sheetIndex, sheet1, sheet2):
        numberOfRecords1 = len(sheet1.records)
        numberOfRecords2 = len(sheet2.records)
        if numberOfRecords1 != numberOfRecords2:
            print "Sheet", sheetIndex, "Number of rows is different", numberOfRecords1, numberOfRecords2
        numberOfRecords = max(numberOfRecords1, numberOfRecords2)
        for each in range(1, numberOfRecords + 1):
            record1 = sheet1.getRecordWithIndex(each)
            record2 = sheet2.getRecordWithIndex(each)
            self.compareRecords(sheetIndex, each, record1, record2)

    def compareRecords(self, sheetIndex, recordIndex, record1, record2):
        numberOfFields1 = len(record1)
        numberOfFields2 = len(record2)
        numberOfFields = max(numberOfFields1, numberOfFields2)
        for each in range(1, numberOfFields + 1):
            field1 = self.getField(record1, each)
            field2 = self.getField(record2, each)
            if field1 != field2:
                 self.printDifference(sheetIndex, recordIndex, each, field1, field2)
                 self.numberOfDifferences += 1
                 if self.maxNumberOfDifferences is not None:
                    if self.numberOfDifferences > self.maxNumberOfDifferences:
                        print "More differences skipped."
                        sys.exit(1)

    # The index is based on one and not on zero.
    def getField(self, record, index, default = ""):
        try:
            value = record[index-1]
        except IndexError:
            value = None
        if value is None:
            return default
        # Be a bit tolerant.
        value = value.strip()
        # Truncate trailing zeros after the decimal point.
        parts = value.split(".")
        if len(parts) == 2 and parts[-1].isdigit():
            a, b = parts
            value = a + "." + b.rstrip("0")
            value = value.strip(".")
        return value

    def printDifference(self, sheetIndex, recordIndex, fieldIndex, field1, field2):
        indicator = self.indicator(sheetIndex, recordIndex, fieldIndex)
        print indicator, "<" + field1 + ">", "<" + field2 + ">"

    def indicator(self, sheetIndex, recordIndex, fieldIndex):
        return '%d.%s:' % (sheetIndex, self.excelReference(recordIndex, fieldIndex))

    # Convert given row and column number to an Excel-style cell name.
    def excelReference(self, row, column):
        quot, rem = divmod(column - 1, 26)
        return((chr(quot-1 + ord('A')) if quot else '') +
               (chr(rem + ord('A')) + str(row)))

if __name__ == "__main__":
    Differ(sys.argv[1:]).process()

