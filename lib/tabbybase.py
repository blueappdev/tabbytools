#!/usr/bin/env python
#
# requires python 2.7
#
# lib/tabbybase.py - tabular listing of a table file
#

class TabbyBaseTool:
    def __init__(self):
        self.indexesToProcess = None # by default all sheets
        
    def setWorksheetsToProcessFromOptionValue(self, aString):
        for each in aString.split(","):
            if self.indexesToProcess is None:
                self.indexesToProcess = []
            parts = each.split("-")
            for each in parts:
                assert each.isdigit(), "digits expected"
            if len(parts) == 1:
                self.indexesToProcess.append(int(each))
            elif len(parts) == 2:
                first, last = parts
                first = int(first)
                last = int(last)
                for each in range(first, last + 1):
                    self.indexesToProcess.append(each)
            else:
                raise Exception, "invalid command line parameter"
        #print self.indexesToProcess
