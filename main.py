from time import sleep
from analysis import analysis, analysisMain
from crawler import crawlerMain



DATE = "20220801"
BOOK_NAME = "化工数据汇总"+DATE+".xlsx"

ALERT_HOLD = ["永安期货", "海通期货", "中信期货", "国泰期货", "东证期货", "恒力期货", "华泰期货", "新湖期货"]
DL_LIST = ['eb2209', 'eg2209', 'pg2209', 'pp2209', 'l2209', 'v2209']  #大连商品交易所
ZZ_LIST = ["MA209", "TA209", "PF210"] #郑州商品交易所
INE_LIST = ["lu2209"]  #上海国际能源交易中心
SH_LIST = ['bu2209', 'fu2301'] #上海期货交易所
VAR_LIST = DL_LIST + ZZ_LIST + INE_LIST + SH_LIST



def main():
    crawlerMain(DATE, DL_LIST)
    sleep(2)
    analysisMain(VAR_LIST)
    sleep(2)
    # generateMain(VAR_LIST)

    print(DATE, "数据整理完毕")

if __name__ == '__main__':
    main()