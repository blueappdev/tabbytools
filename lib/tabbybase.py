#!/usr/bin/env python
#
# requires python 2.7
#
# lib/tabbybase.py - tabular listing of a table file
#

class TabbyBaseTool:
    def setWorksheetsToProcessFromOptionValue(self, aString):
        print "setWorksheetsToProcessFromOptionValue", aString
        self.indexesToProcess = [ int(aString) ]
