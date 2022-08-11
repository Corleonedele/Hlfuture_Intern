#coding=utf-8
import os
import requests
from time import strftime, localtime, perf_counter
from urls import *

STANDARD_TIME = strftime("%Y_%m_%d", localtime())

def makeDir(folder_name=STANDARD_TIME):
    current_path = os.getcwd()
    folder = os.listdir(current_path)
    if folder_name in folder and os.path.isdir(folder_name):
        if not os.listdir(current_path+"/"+folder_name):
            return current_path + "/" + folder_name
        print("---"*8, folder_name, "folder exists", "---"*8)
        return False
    else:
        os.mkdir(folder_name)
        print("---"*8, folder_name, "floder create successfully", "---"*8,)
        return current_path + "/" + folder_name

def load_file(response, folder_path, file_name, file_type=""):
    try:
        path = folder_path+"/"+file_name
    except:
        print("---"*8, "cannot create file", "---"*8)
        return False
    with open(path, "wb") as f:
        for chunk in response.iter_content(chunk_size=1024):  # 每次加载1024字节
            f.write(chunk)

def get_res(url, name, var=""):

    if "dec" in url:
        try:
            res = requests.post(url=url)
        except:
            pass
    else:
        try:
            res = requests.get(url=url)
        except:
            pass

    if res.status_code == 200:

        if "ine" in url or "shfe" in url:
            with open(name, "wb") as f:
                for chunk in res.iter_content(chunk_size=1024):  # 每次加载1024字节
                    f.write(chunk)
        else:
            with open(name, "wb") as f:
                for chunk in res.iter_content(chunk_size=1024):  # 每次加载1024字节
                    f.write(chunk)
        print(var+"爬取成功")
    else:
        print(var+"爬取失败")

def crawlerMain(date, DL_list):
    start_time = perf_counter()

    #郑商所
    get_res(ZZ_URL_HOLD_Daily(date, "MA"), "ZZ_MA_Hold.txt", "MA")
    get_res(ZZ_URL_HOLD_Daily(date, "PF"), "ZZ_PF_Hold.txt", "PF")
    get_res(ZZ_URL_HOLD_Daily(date, "TA"), "ZZ_TA_Hold.txt", "TA")

    #大商所
    for var in DL_list:
        get_res(DL_URL_HOLD_Daily(date, var), "DL_"+var+"_Hold.txt", var)

    #上期所
    get_res(SH_URL_HOLD_Daily(date), "SH_Hold.dat", "SH")

    #上能所
    get_res(SH_URL_HOLD_Daily(date), "INE_Hold.dat", "INE")
    end_time = perf_counter()
    print("爬取共用时:", end_time-start_time, "s")

# crawlerMain(date="20220801", DL_list = ['eb2209', 'eg2209', 'pg2209', 'pp2209', 'l2209', 'v2209'])


#主流化工品种主力合约：
# 1、郑商所：甲醇是MA2209合约、PTA是TA2209合约、短纤是PF2210。
# 2、大商所：苯乙烯是EB2209、乙二醇是EG2209、液化石油气是PG2209、聚丙烯是PP2209、聚乙烯是L2209、PVC是V2209。
# 3、上期所：沥青是BU2209、高硫燃料油或者直接叫燃料油是FU2301。
# 4、上能所：原油是SC2209、低硫燃料油是LU2209