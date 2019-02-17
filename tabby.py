#!/usr/bin/env python
#
# requires python 2.7
#
# tabby.py - tabular listing of a table file
#

import sys
import lib.tableparser
import lib.tableprinter

if __name__ == "__main__":
    lib.tableprinter.Formatter(sys.argv[1:])

