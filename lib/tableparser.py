#!/usr/bin/env python
#
# requires python 2.7
#
# lib/tableparser.py - tabular listing of a table file
#

import sys, getopt, glob, string
import xml.etree.ElementTree

def exit(*args):
    if len(args):
        print "table.py:",
        for each in args:
            print  each,
        print 
    sys.exit(2)

class Workbook:
    def __init__(self):
        self.sheets = []

    def addSheet(self, aSheet):
        self.sheets.append(aSheet)
        
    # The index is based on one and not on zero.
    def getSheetWithIndex(self, index):
        if 1 <= index <= len(self.sheets):
            return self.sheets[index-1]       
        else:
            return None

class Sheet:
    def __init__(self):
        self.records = []

    def addRecord(self, aRecord):
        self.records.append(aRecord)
        
    # The index is based on one and not on zero.
    def getRecordWithIndex(self, index):
        if 1 <= index <= len(self.records):
            return self.records[index-1]       
        else:
            return []

class FileReaderInterface:
    def __init__(self, aFilename):
        self.filename = aFilename

    def getFileReaderClass(self):
        if self.isXMLFile():
            return XMLFileReader
        if self.isZipFile():
            exit("unsupported file format")
        return TextFileReader

    def isXMLFile(self):   
        stream = open(self.filename)
        # BOM should be handled
        contents = stream.read(10)
        stream.close()
        return contents.startswith("<")

    def isZipFile(self):   
        stream = open(self.filename)
        contents = stream.read(10)
        stream.close()
        return contents.startswith("PK")

    def getWorkbook(self):
        return self.getFileReaderClass()(self.filename).getWorkbook()
        
class FileReader:
    def __init__(self, aFilename):
        self.filename = aFilename

class TextFileReader(FileReader):
    def __init__(self, aFilename):
        FileReader.__init__(self, aFilename)

    #
    # Read a tab separated text file and store the records in self.sheets.
    # Text files have exactly one worksheet with one table.
    #
    def getWorkbook(self):
        self.workbook = Workbook()
        self.workbook.addSheet(Sheet())
        stream = open(self.filename)
        for each in stream.readlines():
            self.currentSheet().addRecord(map(string.strip, each.split("\t")))
        stream.close()
        return self.workbook

    def currentSheet(self):
        return self.workbook.sheets[-1]

class XMLFileReader(FileReader):
    def __init__(self, aFilename):
        FileReader.__init__(self, aFilename)

    def getWorkbook(self):
        self.workbook = Workbook()
        tree = xml.etree.ElementTree.parse(self.filename)
        root = tree.getroot()
        for each in root:
            if self.hasTag(each, "Worksheet"):
                self.processWorksheet(each)
        return self.workbook

    def processWorksheet(self, anElement):
        self.currentSheet = Sheet()
        self.workbook.addSheet(self.currentSheet)
        for each in anElement:
            if self.hasTag(each, "Table"):
                self.processTable(each)

    def processTable(self, anElement):
        for each in anElement:
            if self.hasTag(each, "Row"):
                self.processRow(each)

    def processRow(self, anElement):
        index = self.getAttribute(anElement, "Index")
        if index is not None:
            assert index.isdigit(), "numeric row index expected"
            index = int(index)
            assert index > len(self.currentSheet.records)
            for i in range(len(self.currentSheet.records), index - 1):
                self.currentSheet.addRecord([])
        self.currentRecord = []
        self.currentSheet.addRecord(self.currentRecord)
        for each in anElement:
            if self.hasTag(each, "Cell"):
                self.processCell(each)

    def processCell(self, anElement):
        index = self.getAttribute(anElement, "Index")
        if index is not None:
            assert index.isdigit(), "numeric column index expected"
            index = int(index)
            assert index > len(self.currentRecord)
            for i in range(len(self.currentRecord), index - 1):
                self.currentRecord.append("")
        formula = self.getAttribute(anElement, "Formula")
        if formula is not None:
            self.currentRecord.append(formula)
            return
        if list(anElement) == []:
            self.currentRecord.append("")
            return
        for each in anElement:
            if self.hasTag(each, "Data"):
                self.processData(each)

    def processData(self, anElement):
        self.currentRecord.append(anElement.text)

    def hasTag(self, anElement, tagName):
        return self.strippedTag(anElement.tag) == tagName

    def strippedTag(self, aString):
        tokens = aString.split("}")
        assert(len(tokens) == 2)
        return tokens[-1]

    def getAttribute(self, anElement, aString):
        for key, value in anElement.items():
            if self.strippedTag(key) == aString:
                return value
        return None
