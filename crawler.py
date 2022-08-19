#coding=utf-8
import os
import requests
from time import strftime, localtime, perf_counter
from urls import *

STANDARD_TIME = strftime("%Y_%m_%d", localtime())

#预备函数未使用，自定义文件夹
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

#预备函数未使用，在特定位置写入文件
def load_file(response, folder_path, file_name, file_type=""):
    try:
        path = folder_path+"/"+file_name
    except:
        print("---"*8, "cannot create file", "---"*8)
        return False
    with open(path, "wb") as f:
        for chunk in response.iter_content(chunk_size=1024):  # 每次加载1024字节
            f.write(chunk)


#对单个网站进行怕虫，并将结果写成暂存文件
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

#爬虫主函数
def crawlerMain(date, DL_list):
    try:
        start_time = perf_counter()

        #郑商所
        try:
            get_res(ZZ_URL_HOLD_Daily(date, "MA"), "ZZ_MA_"+date+"_Hold.txt", "MA")
        except:
            print("ZZ_MA_"+date+"爬取失败")
        try:
            get_res(ZZ_URL_HOLD_Daily(date, "PF"), "ZZ_PF_"+date+"_Hold.txt", "PF")
        except:
            print("ZZ_PF_"+date+"爬取失败")
        try:
            get_res(ZZ_URL_HOLD_Daily(date, "TA"), "ZZ_TA_"+date+"_Hold.txt", "TA")
        except:
            print("ZZ_TA_"+date+"爬取失败")
        #大商所
        for var in DL_list:
            get_res(DL_URL_HOLD_Daily(date, var), "DL_"+var+"_"+date+"_Hold.txt", var)

        #上期所
        get_res(SH_URL_HOLD_Daily(date), "SH_"+date+"_Hold.dat", "SH")

        #上能所
        get_res(SH_URL_HOLD_Daily(date), "INE_"+date+"_Hold.dat", "INE")
        end_time = perf_counter()
        print("爬取共用时:", end_time-start_time, "s")
        return True
    except:
        print("爬虫失败，程序中断")
        return False
