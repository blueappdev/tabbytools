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
    # Text files have exactly one worksheet/table.
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

    def processWorksheet(self, anXMLElement):
        self.currentSheet = Sheet()
        self.workbook.addSheet(self.currentSheet)
        for each in anXMLElement:
            if self.hasTag(each, "Table"):
                self.processTable(each)

    def processTable(self, anXMLElement):
        for each in anXMLElement:
            if self.hasTag(each, "Row"):
                self.processRow(each)

    def processRow(self, anXMLElement):
        self.currentRecord = []
        self.currentSheet.addRecord(self.currentRecord)
        for each in anXMLElement:
            if self.hasTag(each, "Cell"):
                self.processCell(each)

    def processCell(self, anXMLElement):
        for each in anXMLElement:
            if self.hasTag(each, "Data"):
                self.processData(each)

    def processData(self, anElement):
        self.currentRecord.append(anElement.text)

    def hasTag(self, anElement, tagName):
        return self.extractTag(anElement) == tagName

    def extractTag(self, anElement):
        tokens = anElement.tag.split("}")
        assert(len(tokens) == 2)
        return tokens[-1]



