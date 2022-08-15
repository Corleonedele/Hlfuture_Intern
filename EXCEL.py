#此文件可忽略
import openpyxl
from os import remove
from openpyxl.styles import Alignment, PatternFill
from data import *
from main import BOOK_NAME, DATE, MONITOR_POS
        
def writeToExcel_eb2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = eb2209_long.get("&nbsp;")
        total_short =  eb2209_short.get("&nbsp;")
        total_long_change = eb2209_long_change.get("&nbsp;")
        total_short_change = eb2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = eb2209_long.get("")
        total_short =  eb2209_short.get("")
        total_long_change = eb2209_long_change.get("")
        total_short_change = eb2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in eb2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = eb2209_long[key]
        sheet.cell(ROW, COL+2).value = eb2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in eb2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = eb2209_short[key]
        sheet.cell(ROW, COL+2).value = eb2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in eb2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = eb2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = eb2209_short[key]
            val = eb2209_long[key] - eb2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = eb2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = eb2209_short_change[key]
            val = eb2209_long_change[key] - eb2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in eb2209_short:
        if key == "&nbsp;" or key == "":continue
        if eb2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = eb2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = eb2209_long[key]
            val = eb2209_long[key] - eb2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = eb2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = eb2209_long_change[key]
            val = eb2209_long_change[key] - eb2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in eb2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += eb2209_long[key]
        long_5_change = eb2209_long_change[key]
    tem_count = 0
    for key in eb2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += eb2209_short[key]
        short_5_change = eb2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = eb2209_long_change[key]-eb2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = eb2209_long[hold]
            sheet.cell(ROW, COL+2).value = eb2209_short[hold]
            sheet.cell(ROW, COL+3).value = eb2209_long[hold]-eb2209_short[hold]
            sheet.cell(ROW, COL+4).value = eb2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = eb2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = eb2209_long_change[hold] - eb2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_eb2209(BOOK_NAME, DATE, "eb2209")

def writeToExcel_eg2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = eg2209_long.get("&nbsp;")
        total_short =  eg2209_short.get("&nbsp;")
        total_long_change = eg2209_long_change.get("&nbsp;")
        total_short_change = eg2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = eg2209_long.get("")
        total_short =  eg2209_short.get("")
        total_long_change = eg2209_long_change.get("")
        total_short_change = eg2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in eg2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = eg2209_long[key]
        sheet.cell(ROW, COL+2).value = eg2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in eg2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = eg2209_short[key]
        sheet.cell(ROW, COL+2).value = eg2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in eg2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = eg2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = eg2209_short[key]
            val = eg2209_long[key] - eg2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = eg2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = eg2209_short_change[key]
            val = eg2209_long_change[key] - eg2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in eg2209_short:
        if key == "&nbsp;" or key == "":continue
        if eg2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = eg2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = eg2209_long[key]
            val = eg2209_long[key] - eg2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = eg2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = eg2209_long_change[key]
            val = eg2209_long_change[key] - eg2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in eg2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += eg2209_long[key]
        long_5_change = eg2209_long_change[key]
    tem_count = 0
    for key in eg2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += eg2209_short[key]
        short_5_change = eg2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = eg2209_long_change[key]-eg2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = eg2209_long[hold]
            sheet.cell(ROW, COL+2).value = eg2209_short[hold]
            sheet.cell(ROW, COL+3).value = eg2209_long[hold]-eg2209_short[hold]
            sheet.cell(ROW, COL+4).value = eg2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = eg2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = eg2209_long_change[hold] - eg2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_eg2209(BOOK_NAME, DATE, "eg2209")

def writeToExcel_pg2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = pg2209_long.get("&nbsp;")
        total_short =  pg2209_short.get("&nbsp;")
        total_long_change = pg2209_long_change.get("&nbsp;")
        total_short_change = pg2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = pg2209_long.get("")
        total_short =  pg2209_short.get("")
        total_long_change = pg2209_long_change.get("")
        total_short_change = pg2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in pg2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = pg2209_long[key]
        sheet.cell(ROW, COL+2).value = pg2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in pg2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = pg2209_short[key]
        sheet.cell(ROW, COL+2).value = pg2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in pg2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = pg2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = pg2209_short[key]
            val = pg2209_long[key] - pg2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = pg2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = pg2209_short_change[key]
            val = pg2209_long_change[key] - pg2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in pg2209_short:
        if key == "&nbsp;" or key == "":continue
        if pg2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = pg2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = pg2209_long[key]
            val = pg2209_long[key] - pg2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = pg2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = pg2209_long_change[key]
            val = pg2209_long_change[key] - pg2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in pg2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += pg2209_long[key]
        long_5_change = pg2209_long_change[key]
    tem_count = 0
    for key in pg2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += pg2209_short[key]
        short_5_change = pg2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = pg2209_long_change[key]-pg2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = pg2209_long[hold]
            sheet.cell(ROW, COL+2).value = pg2209_short[hold]
            sheet.cell(ROW, COL+3).value = pg2209_long[hold]-pg2209_short[hold]
            sheet.cell(ROW, COL+4).value = pg2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = pg2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = pg2209_long_change[hold] - pg2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_pg2209(BOOK_NAME, DATE, "pg2209")

def writeToExcel_pp2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = pp2209_long.get("&nbsp;")
        total_short =  pp2209_short.get("&nbsp;")
        total_long_change = pp2209_long_change.get("&nbsp;")
        total_short_change = pp2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = pp2209_long.get("")
        total_short =  pp2209_short.get("")
        total_long_change = pp2209_long_change.get("")
        total_short_change = pp2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in pp2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = pp2209_long[key]
        sheet.cell(ROW, COL+2).value = pp2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in pp2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = pp2209_short[key]
        sheet.cell(ROW, COL+2).value = pp2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in pp2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = pp2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = pp2209_short[key]
            val = pp2209_long[key] - pp2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = pp2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = pp2209_short_change[key]
            val = pp2209_long_change[key] - pp2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in pp2209_short:
        if key == "&nbsp;" or key == "":continue
        if pp2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = pp2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = pp2209_long[key]
            val = pp2209_long[key] - pp2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = pp2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = pp2209_long_change[key]
            val = pp2209_long_change[key] - pp2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in pp2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += pp2209_long[key]
        long_5_change = pp2209_long_change[key]
    tem_count = 0
    for key in pp2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += pp2209_short[key]
        short_5_change = pp2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = pp2209_long_change[key]-pp2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = pp2209_long[hold]
            sheet.cell(ROW, COL+2).value = pp2209_short[hold]
            sheet.cell(ROW, COL+3).value = pp2209_long[hold]-pp2209_short[hold]
            sheet.cell(ROW, COL+4).value = pp2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = pp2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = pp2209_long_change[hold] - pp2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_pp2209(BOOK_NAME, DATE, "pp2209")

def writeToExcel_l2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = l2209_long.get("&nbsp;")
        total_short =  l2209_short.get("&nbsp;")
        total_long_change = l2209_long_change.get("&nbsp;")
        total_short_change = l2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = l2209_long.get("")
        total_short =  l2209_short.get("")
        total_long_change = l2209_long_change.get("")
        total_short_change = l2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in l2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = l2209_long[key]
        sheet.cell(ROW, COL+2).value = l2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in l2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = l2209_short[key]
        sheet.cell(ROW, COL+2).value = l2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in l2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = l2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = l2209_short[key]
            val = l2209_long[key] - l2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = l2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = l2209_short_change[key]
            val = l2209_long_change[key] - l2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in l2209_short:
        if key == "&nbsp;" or key == "":continue
        if l2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = l2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = l2209_long[key]
            val = l2209_long[key] - l2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = l2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = l2209_long_change[key]
            val = l2209_long_change[key] - l2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in l2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += l2209_long[key]
        long_5_change = l2209_long_change[key]
    tem_count = 0
    for key in l2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += l2209_short[key]
        short_5_change = l2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = l2209_long_change[key]-l2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = l2209_long[hold]
            sheet.cell(ROW, COL+2).value = l2209_short[hold]
            sheet.cell(ROW, COL+3).value = l2209_long[hold]-l2209_short[hold]
            sheet.cell(ROW, COL+4).value = l2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = l2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = l2209_long_change[hold] - l2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_l2209(BOOK_NAME, DATE, "l2209")

def writeToExcel_v2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = v2209_long.get("&nbsp;")
        total_short =  v2209_short.get("&nbsp;")
        total_long_change = v2209_long_change.get("&nbsp;")
        total_short_change = v2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = v2209_long.get("")
        total_short =  v2209_short.get("")
        total_long_change = v2209_long_change.get("")
        total_short_change = v2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in v2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = v2209_long[key]
        sheet.cell(ROW, COL+2).value = v2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in v2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = v2209_short[key]
        sheet.cell(ROW, COL+2).value = v2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in v2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = v2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = v2209_short[key]
            val = v2209_long[key] - v2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = v2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = v2209_short_change[key]
            val = v2209_long_change[key] - v2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in v2209_short:
        if key == "&nbsp;" or key == "":continue
        if v2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = v2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = v2209_long[key]
            val = v2209_long[key] - v2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = v2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = v2209_long_change[key]
            val = v2209_long_change[key] - v2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in v2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += v2209_long[key]
        long_5_change = v2209_long_change[key]
    tem_count = 0
    for key in v2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += v2209_short[key]
        short_5_change = v2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = v2209_long_change[key]-v2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = v2209_long[hold]
            sheet.cell(ROW, COL+2).value = v2209_short[hold]
            sheet.cell(ROW, COL+3).value = v2209_long[hold]-v2209_short[hold]
            sheet.cell(ROW, COL+4).value = v2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = v2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = v2209_long_change[hold] - v2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_v2209(BOOK_NAME, DATE, "v2209")

def writeToExcel_MA209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = MA209_long.get("&nbsp;")
        total_short =  MA209_short.get("&nbsp;")
        total_long_change = MA209_long_change.get("&nbsp;")
        total_short_change = MA209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = MA209_long.get("")
        total_short =  MA209_short.get("")
        total_long_change = MA209_long_change.get("")
        total_short_change = MA209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in MA209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = MA209_long[key]
        sheet.cell(ROW, COL+2).value = MA209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in MA209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = MA209_short[key]
        sheet.cell(ROW, COL+2).value = MA209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in MA209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = MA209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = MA209_short[key]
            val = MA209_long[key] - MA209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = MA209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = MA209_short_change[key]
            val = MA209_long_change[key] - MA209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in MA209_short:
        if key == "&nbsp;" or key == "":continue
        if MA209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = MA209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = MA209_long[key]
            val = MA209_long[key] - MA209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = MA209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = MA209_long_change[key]
            val = MA209_long_change[key] - MA209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in MA209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += MA209_long[key]
        long_5_change = MA209_long_change[key]
    tem_count = 0
    for key in MA209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += MA209_short[key]
        short_5_change = MA209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = MA209_long_change[key]-MA209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = MA209_long[hold]
            sheet.cell(ROW, COL+2).value = MA209_short[hold]
            sheet.cell(ROW, COL+3).value = MA209_long[hold]-MA209_short[hold]
            sheet.cell(ROW, COL+4).value = MA209_long_change[hold]
            sheet.cell(ROW, COL+5).value = MA209_short_change[hold]
            sheet.cell(ROW, COL+6).value = MA209_long_change[hold] - MA209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_MA209(BOOK_NAME, DATE, "MA209")

def writeToExcel_TA209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = TA209_long.get("&nbsp;")
        total_short =  TA209_short.get("&nbsp;")
        total_long_change = TA209_long_change.get("&nbsp;")
        total_short_change = TA209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = TA209_long.get("")
        total_short =  TA209_short.get("")
        total_long_change = TA209_long_change.get("")
        total_short_change = TA209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in TA209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = TA209_long[key]
        sheet.cell(ROW, COL+2).value = TA209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in TA209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = TA209_short[key]
        sheet.cell(ROW, COL+2).value = TA209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in TA209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = TA209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = TA209_short[key]
            val = TA209_long[key] - TA209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = TA209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = TA209_short_change[key]
            val = TA209_long_change[key] - TA209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in TA209_short:
        if key == "&nbsp;" or key == "":continue
        if TA209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = TA209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = TA209_long[key]
            val = TA209_long[key] - TA209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = TA209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = TA209_long_change[key]
            val = TA209_long_change[key] - TA209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in TA209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += TA209_long[key]
        long_5_change = TA209_long_change[key]
    tem_count = 0
    for key in TA209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += TA209_short[key]
        short_5_change = TA209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = TA209_long_change[key]-TA209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = TA209_long[hold]
            sheet.cell(ROW, COL+2).value = TA209_short[hold]
            sheet.cell(ROW, COL+3).value = TA209_long[hold]-TA209_short[hold]
            sheet.cell(ROW, COL+4).value = TA209_long_change[hold]
            sheet.cell(ROW, COL+5).value = TA209_short_change[hold]
            sheet.cell(ROW, COL+6).value = TA209_long_change[hold] - TA209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_TA209(BOOK_NAME, DATE, "TA209")

def writeToExcel_PF210(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = PF210_long.get("&nbsp;")
        total_short =  PF210_short.get("&nbsp;")
        total_long_change = PF210_long_change.get("&nbsp;")
        total_short_change = PF210_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = PF210_long.get("")
        total_short =  PF210_short.get("")
        total_long_change = PF210_long_change.get("")
        total_short_change = PF210_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in PF210_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = PF210_long[key]
        sheet.cell(ROW, COL+2).value = PF210_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in PF210_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = PF210_short[key]
        sheet.cell(ROW, COL+2).value = PF210_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in PF210_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = PF210_long[key]
        try:
            sheet.cell(ROW, COL+2).value = PF210_short[key]
            val = PF210_long[key] - PF210_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = PF210_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = PF210_short_change[key]
            val = PF210_long_change[key] - PF210_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in PF210_short:
        if key == "&nbsp;" or key == "":continue
        if PF210_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = PF210_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = PF210_long[key]
            val = PF210_long[key] - PF210_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = PF210_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = PF210_long_change[key]
            val = PF210_long_change[key] - PF210_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in PF210_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += PF210_long[key]
        long_5_change = PF210_long_change[key]
    tem_count = 0
    for key in PF210_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += PF210_short[key]
        short_5_change = PF210_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = PF210_long_change[key]-PF210_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = PF210_long[hold]
            sheet.cell(ROW, COL+2).value = PF210_short[hold]
            sheet.cell(ROW, COL+3).value = PF210_long[hold]-PF210_short[hold]
            sheet.cell(ROW, COL+4).value = PF210_long_change[hold]
            sheet.cell(ROW, COL+5).value = PF210_short_change[hold]
            sheet.cell(ROW, COL+6).value = PF210_long_change[hold] - PF210_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_PF210(BOOK_NAME, DATE, "PF210")

def writeToExcel_lu2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = lu2209_long.get("&nbsp;")
        total_short =  lu2209_short.get("&nbsp;")
        total_long_change = lu2209_long_change.get("&nbsp;")
        total_short_change = lu2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = lu2209_long.get("")
        total_short =  lu2209_short.get("")
        total_long_change = lu2209_long_change.get("")
        total_short_change = lu2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in lu2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = lu2209_long[key]
        sheet.cell(ROW, COL+2).value = lu2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in lu2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = lu2209_short[key]
        sheet.cell(ROW, COL+2).value = lu2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in lu2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = lu2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = lu2209_short[key]
            val = lu2209_long[key] - lu2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = lu2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = lu2209_short_change[key]
            val = lu2209_long_change[key] - lu2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in lu2209_short:
        if key == "&nbsp;" or key == "":continue
        if lu2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = lu2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = lu2209_long[key]
            val = lu2209_long[key] - lu2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = lu2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = lu2209_long_change[key]
            val = lu2209_long_change[key] - lu2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in lu2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += lu2209_long[key]
        long_5_change = lu2209_long_change[key]
    tem_count = 0
    for key in lu2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += lu2209_short[key]
        short_5_change = lu2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = lu2209_long_change[key]-lu2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = lu2209_long[hold]
            sheet.cell(ROW, COL+2).value = lu2209_short[hold]
            sheet.cell(ROW, COL+3).value = lu2209_long[hold]-lu2209_short[hold]
            sheet.cell(ROW, COL+4).value = lu2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = lu2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = lu2209_long_change[hold] - lu2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_lu2209(BOOK_NAME, DATE, "lu2209")

def writeToExcel_bu2209(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = bu2209_long.get("&nbsp;")
        total_short =  bu2209_short.get("&nbsp;")
        total_long_change = bu2209_long_change.get("&nbsp;")
        total_short_change = bu2209_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = bu2209_long.get("")
        total_short =  bu2209_short.get("")
        total_long_change = bu2209_long_change.get("")
        total_short_change = bu2209_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in bu2209_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = bu2209_long[key]
        sheet.cell(ROW, COL+2).value = bu2209_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in bu2209_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = bu2209_short[key]
        sheet.cell(ROW, COL+2).value = bu2209_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in bu2209_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = bu2209_long[key]
        try:
            sheet.cell(ROW, COL+2).value = bu2209_short[key]
            val = bu2209_long[key] - bu2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = bu2209_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = bu2209_short_change[key]
            val = bu2209_long_change[key] - bu2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in bu2209_short:
        if key == "&nbsp;" or key == "":continue
        if bu2209_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = bu2209_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = bu2209_long[key]
            val = bu2209_long[key] - bu2209_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = bu2209_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = bu2209_long_change[key]
            val = bu2209_long_change[key] - bu2209_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in bu2209_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += bu2209_long[key]
        long_5_change = bu2209_long_change[key]
    tem_count = 0
    for key in bu2209_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += bu2209_short[key]
        short_5_change = bu2209_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = bu2209_long_change[key]-bu2209_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = bu2209_long[hold]
            sheet.cell(ROW, COL+2).value = bu2209_short[hold]
            sheet.cell(ROW, COL+3).value = bu2209_long[hold]-bu2209_short[hold]
            sheet.cell(ROW, COL+4).value = bu2209_long_change[hold]
            sheet.cell(ROW, COL+5).value = bu2209_short_change[hold]
            sheet.cell(ROW, COL+6).value = bu2209_long_change[hold] - bu2209_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_bu2209(BOOK_NAME, DATE, "bu2209")

def writeToExcel_fu2301(book_name, date, var):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    try:sheet = workbook.create_sheet(var)
    except:sheet = workbook.active
    sheet.title = var
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ALERT = PatternFill('solid', fgColor="ffc7ce")
    ALERT_long = []
    ALERT_short = []
    try:
        total_long = fu2301_long.get("&nbsp;")
        total_short =  fu2301_short.get("&nbsp;")
        total_long_change = fu2301_long_change.get("&nbsp;")
        total_short_change = fu2301_short_change.get("&nbsp;")
        if total_long == None:raise TypeError
    except:
        total_long = fu2301_long.get("")
        total_short =  fu2301_short.get("")
        total_long_change = fu2301_long_change.get("")
        total_short_change = fu2301_short_change.get("")
    COL = 4
    ROW = 3
    sheet.cell(ROW, COL-1).value=var 
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="多头变量"
    sheet.cell(ROW, COL+3).value="期货公司"
    sheet.cell(ROW, COL+4).value="空头持仓"
    sheet.cell(ROW, COL+5).value="空头变量"
    ROW=4
    for key in fu2301_long:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = fu2301_long[key]
        sheet.cell(ROW, COL+2).value = fu2301_long_change[key]
        ROW+=1
    COL=7    
    ROW=4
    for key in fu2301_short:
        if key == "&nbsp;" or key == "":tem_key = "总计"
        else:tem_key=key
        sheet.cell(ROW, COL).value = tem_key
        sheet.cell(ROW, COL+1).value = fu2301_short[key]
        sheet.cell(ROW, COL+2).value = fu2301_short_change[key]
        ROW+=1
    COL = 12
    ROW = 3
    sheet.cell(ROW, COL-1).value=var
    sheet.cell(ROW, COL).value="期货公司"
    sheet.cell(ROW, COL+1).value="多头持仓"
    sheet.cell(ROW, COL+2).value="空头持仓"
    sheet.cell(ROW, COL+3).value="多头-空头"
    sheet.cell(ROW, COL+4).value="多头变量"
    sheet.cell(ROW, COL+5).value="空头变量"
    sheet.cell(ROW, COL+6).value="多头-空头"
    COL = 12
    ROW = 4
    for key in fu2301_long:
        if key == "&nbsp;" or key == "":continue
        sheet.cell(ROW, COL).value = key
        sheet.cell(ROW, COL+1).value = fu2301_long[key]
        try:
            sheet.cell(ROW, COL+2).value = fu2301_short[key]
            val = fu2301_long[key] - fu2301_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = fu2301_long_change[key]
        try:
            sheet.cell(ROW, COL+5).value = fu2301_short_change[key]
            val = fu2301_long_change[key] - fu2301_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_long.append(key)
        except:
            pass
        ROW+=1
    for key in fu2301_short:
        if key == "&nbsp;" or key == "":continue
        if fu2301_long.get(key) == None:
            sheet.cell(ROW, COL).value = key
            sheet.cell(ROW, COL+2).value = fu2301_short[key]
        else:continue
        try:
            sheet.cell(ROW, COL+1).value = fu2301_long[key]
            val = fu2301_long[key] - fu2301_short[key]
            sheet.cell(ROW, COL+3).value = val
        except:
            pass
        sheet.cell(ROW, COL+4).value = fu2301_short_change[key]
        try:
            sheet.cell(ROW, COL+5).value = fu2301_long_change[key]
            val = fu2301_long_change[key] - fu2301_short_change[key]
            sheet.cell(ROW, COL+6).value = val
            if abs(val) > 5000:
                sheet.cell(ROW, COL+6).fill = ALERT
                ALERT_short.append(key)
        except:
            pass
        ROW+=1
    COL = 4
    ROW = 60
    sheet.cell(ROW-1, COL+1).value = "多仓"
    sheet.cell(ROW-1, COL+2).value = "空仓"
    sheet.cell(ROW, COL).value = "前20仓位总量"
    sheet.cell(ROW, COL+1).value = total_long
    sheet.cell(ROW, COL+2).value = total_short
    sheet.cell(ROW+1, COL).value = "前20仓位总变量"
    sheet.cell(ROW+1, COL+1).value = total_long_change
    sheet.cell(ROW+1, COL+2).value = total_short_change
    sheet.cell(ROW+2, COL).value = "前20仓位多空倾向"
    if total_long_change >= total_short_change:sheet.cell(ROW+2, COL+1).value="做多 轧差量为"+str(total_long-total_short)
    else:sheet.cell(ROW+2, COL+1).value="做空 轧差量为"+str(total_short-total_long)
    long_5 = 0
    long_5_change = 0
    short_5 = 0
    short_5_change = 0
    tem_count = 0
    for key in fu2301_long:
        if tem_count == 5:break
        else:tem_count+=1
        long_5 += fu2301_long[key]
        long_5_change = fu2301_long_change[key]
    tem_count = 0
    for key in fu2301_short:
        if tem_count == 5:break
        else:tem_count+=1
        short_5 += fu2301_short[key]
        short_5_change = fu2301_short_change[key]
    sheet.cell(ROW+3, COL).value = "前5仓位总量"
    sheet.cell(ROW+3, COL+1).value = long_5
    sheet.cell(ROW+3, COL+2).value = short_5
    sheet.cell(ROW+4, COL).value = "前5仓位总变量"
    sheet.cell(ROW+4, COL+1).value = long_5_change
    sheet.cell(ROW+4, COL+2).value = short_5_change 
    sheet.cell(ROW+5, COL).value = "前5仓位集中度"
    sheet.cell(ROW+6, COL).value = "前5仓位多空倾向"
    sheet.cell(ROW+5, COL+1).value = str(round(long_5/total_long, 3) * 100)+"%"
    sheet.cell(ROW+5, COL+2).value = str(round(short_5/total_short, 3) * 100)+"%"
    if (long_5/total_long) >= (short_5/total_short):sheet.cell(ROW+6, COL+1).value = "做多 轧差量为"+str(long_5-short_5)
    else:sheet.cell(ROW+6, COL+1).value = "做空 轧差量为"+str(short_5-long_5)
    COL = 10
    ROW = 60
    sheet.cell(ROW-1, COL).value = "异常值"
    sheet.cell(ROW-1, COL+1).value = "方向"
    sheet.cell(ROW-1, COL+2).value = "数量"
    for key in ALERT_long+ALERT_short:
        sheet.cell(ROW, COL).value = key
        diff = fu2301_long_change[key]-fu2301_short_change[key]
        if diff >= 0:
            sheet.cell(ROW, COL+1).value = "做多"
            sheet.cell(ROW, COL+2).value = diff
        else:
            sheet.cell(ROW, COL+1).value = "做空"
            sheet.cell(ROW, COL+2).value = diff
        ROW+=1
    COL = 4
    ROW = 80
    sheet.cell(ROW-1, COL).value = "重点席位信息"
    sheet.cell(ROW-1, COL+1).value="多头持仓"
    sheet.cell(ROW-1, COL+2).value="空头持仓"
    sheet.cell(ROW-1, COL+3).value="多头-空头"
    sheet.cell(ROW-1, COL+4).value="多头变量"
    sheet.cell(ROW-1, COL+5).value="空头变量"
    sheet.cell(ROW-1, COL+6).value="多头-空头"
    for hold in MONITOR_POS:
        sheet.cell(ROW, COL).value = hold
        try:
            sheet.cell(ROW, COL+1).value = fu2301_long[hold]
            sheet.cell(ROW, COL+2).value = fu2301_short[hold]
            sheet.cell(ROW, COL+3).value = fu2301_long[hold]-fu2301_short[hold]
            sheet.cell(ROW, COL+4).value = fu2301_long_change[hold]
            sheet.cell(ROW, COL+5).value = fu2301_short_change[hold]
            sheet.cell(ROW, COL+6).value = fu2301_long_change[hold] - fu2301_short_change[hold]
        except:
            sheet.cell(ROW, COL+1).value = "暂无数据"
        ROW +=1
    workbook.save(book_name)
writeToExcel_fu2301(BOOK_NAME, DATE, "fu2301")

def dataSum(book_name, date, sheet_name):
    try:workbook = openpyxl.load_workbook(book_name)
    except: workbook = openpyxl.Workbook()
    sheets = workbook.sheetnames 
    sheet=workbook[sheets[0]] 
    sheet.title = sheet_name
    sheet.merge_cells('A1:R1')
    sheet.cell(1,1).value = '化工数据'+date+"汇总"
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ROW_W = 5
    COL_W = 4
    sheet.cell(ROW_W-2, COL_W).value = "主流化工品种的多空仓增减"
    sheet.cell(ROW_W-2, COL_W+5).value = "主流化工品种的异常持仓增减量"
    for i in range(1, len(sheets)):
        ws = workbook[sheets[i]]
        var = ws.title
        ROW_in = ROW_W
        COL_in = COL_W

        sheet.cell(ROW_in, COL_in-1).value = var
        for col in range(3, 8):
            ROW_W = ROW_in
            for row in range(59, 67):
                sheet.cell(ROW_W, COL_W).value = ws.cell(row,col).value
                ROW_W+=1
            COL_W+=1
        

        for col in range(9, 13):
            ROW_W = ROW_in
            tem_count = 0
            for row in range(59, 79):
                if ws.cell(row,col).value == None and tem_count>6:
                    break
                else:
                    sheet.cell(ROW_W, COL_W).value = ws.cell(row,col).value
                ROW_W+=1
                tem_count+=1
            COL_W+=1
        COL_W = COL_in
        ROW_W += 3
    workbook.save(BOOK_NAME)

dataSum(BOOK_NAME, DATE, "数据汇总")
