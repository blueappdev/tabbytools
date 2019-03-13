#!/usr/bin/env python
#
# requires python 2.7
#
# tabby.py - tabular listing of a table file
#

import sys
import lib.tabbyprinter

if __name__ == "__main__":
    lib.tabbyprinter.Formatter(sys.argv[1:]).process()

