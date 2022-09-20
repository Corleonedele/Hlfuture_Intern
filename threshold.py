import datetime
import re
from time import sleep,perf_counter
from crawler import crawlerMain
from clean import clean
from analysis import ZZ_analysis, INE_analysis, SH_analysis, DL_analysis, getint

DL_LIST = ['eb2109', 'eg2109', 'pg2110', 'pp2109', 'l2109', 'v2109']  #大连商品交易所
ZZ_LIST = ["MA109", "TA109", "PF110"] #郑州商品交易所
INE_LIST = ["lu2110"]  #上海国际能源交易中心
SH_LIST = ['bu2112', 'fu2201'] #上海期货交易所

DATA_LOSS = 0
DATA_LOSS_LIST = []

VAR_LIST = DL_LIST + ZZ_LIST + INE_LIST + SH_LIST



def getTradeDay(start_date, enddate):
    holiday = ["20210101","20210211","20210212","20210215","20210216","20210217","20210405","20210503","20210504","20210505","20210614","20210920","20210921","20211001","20211004","20211005","20211006","20211007"]
    date_list =[]
    start_date = datetime.datetime.strptime(start_date, '%Y%m%d')
    enddate = datetime.datetime.strptime(enddate, '%Y%m%d')
    while start_date < enddate:
            start_date += datetime.timedelta(days=1)
            if datetime.datetime.weekday(start_date) in [5, 6]:
                continue
            
            date_list.append(start_date.strftime('%Y%m%d'))
    for day in holiday:
        try:
            date_list.pop(day)
        except:
            pass
    return date_list

def getDegree(degree, list):
    return list[int(len(list) * float(degree))-1]

def analysis(var, one_file_status=False, one_file_date=""):
    result = ''.join(re.findall(r'[A-Za-z]', var)) 
    # print(result)
    if result in ["MA", "TA", "PF"]:
        return ZZ_analysis(var,one_file_status=False, one_file_date=one_file_date)
    elif result in ["lu", "sc"]:
        return INE_analysis(var,one_file_status=False, one_file_date=one_file_date)
    elif result in ["bu", "fu"]:
        return SH_analysis(var,one_file_status=False, one_file_date=one_file_date)
    elif result in ["eb", "eg", "pg", "pp", "l", "v"]:
        return DL_analysis(var,one_file_status=False, one_file_date=one_file_date)
    return False

def getThresholdValie(var, one_file_date):
    print("品种:", var, "日期", one_file_date)
    ob_list=[]

    trading_name, trading_hold, trading_hold_change, long_name, long_hold, long_hold_change, short_name, short_hold, short_hold_change, status = analysis(var=var, one_file_status=False, one_file_date=one_file_date)

    for value in long_hold_change + short_hold_change:
        ob_list.append(abs(int(getint(value))))

    ob_list.sort(reverse=False)
    ob_list = ob_list[:len(ob_list)-2]

    return getDegree(0.99, ob_list), getDegree(0.90, ob_list), getDegree(0.85, ob_list), getDegree(0.80, ob_list), getDegree(0.70, ob_list)

def writeThresholdValue(var, date_list):
    global DATA_LOSS, DATA_LOSS_LIST
    values_max = []
    values_90 = []
    values_85 = []
    values_80 = []
    values_70 = []

    for date in date_list:
        try:
            value_max, value_90, value_85, value_80, value_70 = getThresholdValie(var, date)
            values_max.append(value_max)
            values_90.append(value_90)
            values_85.append(value_85)
            values_80.append(value_80)
            values_70.append(value_70)

        except:
            print(var, date, "数据缺失")
            DATA_LOSS += 1
            DATA_LOSS_LIST.append(var+date)
            continue


    values_max.sort()
    values_90.sort()
    values_85.sort()
    values_80.sort()
    values_70.sort()

    print(values_max)

    with open("thresholdData.py", "a") as dataFile:
        try:
            dataFile.write(var+"_max = ")
            dataFile.write(str(getDegree(0.9, values_max))+"\n")
            dataFile.write(var+"_90 = ")
            dataFile.write(str(getDegree(0.9, values_90))+"\n")
            dataFile.write(var+"_85 = ")
            dataFile.write(str(getDegree(0.9, values_85))+"\n")
            dataFile.write(var+"_80 = ")
            dataFile.write(str(getDegree(0.9, values_80))+"\n")
            dataFile.write(var+"_70 = ")
            dataFile.write(str(getDegree(0.9, values_70))+"\n")
        except:
            pass



def thresholdMain():
    start_time = perf_counter()
    date_list = getTradeDay("20210801", "20211210")

    for date in date_list:
        try:
            print("-"*30, "爬取日期:", date)
            crawlerMain(date, DL_LIST)
            print("-"*30, "爬取成功")
            sleep(2)
        except:
            print("-"*30, "爬取日期:", date, "爬取失败")

    with open("thresholdData.py", "a") as dataFile:
        dataFile.write("# -*- coding: utf-8 -*-\n")

    for var in VAR_LIST:
        writeThresholdValue(var, date_list)

    sleep(2)
    print("数据处理完毕，清理中间文件中....")
    # clean()
    end_time = perf_counter()
    print("共用时:", end_time-start_time, "s")
    print("DATA_LOSS", DATA_LOSS)
    # print("DATA_LOSS_LIST", DATA_LOSS_LIST)

if __name__ == "__main__":
    thresholdMain()