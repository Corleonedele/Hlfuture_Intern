from time import sleep

#生成总结数据
def generateSum():
    with open("EXCEL.py",'a') as funciontFile:
        with open("dataSum.conf", 'r') as exampleFile:
                for line in exampleFile:
                    funciontFile.write(line)

#生成每一个sheet的数据
def generateMain(VAR_LIST):
    with open("EXCEL.py",'a') as funciontFile: #打开一个叫EXCEL.py的文件
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
                        rep_line = line.replace("EXAMPLE", var) #将配置文件里的内容进行替换
                    else:
                        rep_line = line
                    funciontFile.write(rep_line) #将替换完的内容写入EXCEL.py
        # funciontFile.write("remove(\"data.py\")")
    sleep(2)
    generateSum()
