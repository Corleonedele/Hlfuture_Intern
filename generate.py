def generateMain(VAR_LIST):
    with open("EXCEL.py",'a') as funciontFile:
        funciontFile.write("""
import openpyxl
from os import remove
from openpyxl.styles import Alignment, PatternFill
from data import *
from main import BOOK_NAME, DATE, MONITOR_POS
        """)

        for var in VAR_LIST:
            with open("example.conf", 'r') as exampleFile:
                for line in exampleFile:
                    if "EXAMPLE" in line:
                        rep_line = line.replace("EXAMPLE", var)
                    else:
                        rep_line = line
                    funciontFile.write(rep_line)
        funciontFile.write("remove(\"data.py\")")