from time import sleep, perf_counter
from os import remove
from analysis import analysisMain
from crawler import crawlerMain
from generate import generateMain
from clean import clean

#补充数据，美化，数据集合

DATE = "20220802"
BOOK_NAME = "化工数据汇总"+DATE+".xlsx"

MONITOR_POS = ["永安期货", "海通期货", "中信期货", "国泰期货", "东证期货", "恒力期货", "华泰期货", "新湖期货"]
DL_LIST = ['eb2209', 'eg2209', 'pg2209', 'pp2209', 'l2209', 'v2209']  #大连商品交易所
ZZ_LIST = ["MA209", "TA209", "PF210"] #郑州商品交易所
INE_LIST = ["lu2209"]  #上海国际能源交易中心
SH_LIST = ['bu2209', 'fu2301'] #上海期货交易所


VAR_LIST = DL_LIST + ZZ_LIST + INE_LIST + SH_LIST

def main():
    start_time = perf_counter()
    try:
        remove("EXCEL.py")
    except:
        pass
    try:
        remove("data.py")
    except:
        pass
    crawler_status = crawlerMain(DATE, DL_LIST)
    if not crawler_status:return
    sleep(2)
    analysis_status = analysisMain(VAR_LIST)
    if not analysis_status:return
    sleep(2)
    generateMain(VAR_LIST)
    sleep(2)
    clean_status = clean()
    if not clean_status:return
    print("数据准备完毕，请运行EXCEL.py生成EXCEL文件")
    print(DATE, "数据整理完毕")
    end_time = perf_counter()
    print("共用时:", end_time-start_time, "s")

if __name__ == '__main__':
    main()