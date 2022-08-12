
import re
import os


#主流化工品种主力合约：
# 1、郑商所：甲醇是MA2209合约、PTA是TA2209合约、短纤是PF2210。
# 2、大商所：苯乙烯是EB2209、乙二醇是EG2209、液化石油气是PG2209、聚丙烯是PP2209、聚乙烯是L2209、PVC是V2209。
# 3、上期所：沥青是BU2209、高硫燃料油或者直接叫燃料油是FU2301。
# 4、上能所：原油是SC2209、低硫燃料油是LU2209

def get_filename(var=""):
    fl = os.listdir()
    result=[]
    for file in fl:
        if var in file:
            result.append(file)
    return result


def intcal(str1, str2):
    str1 = int(re.sub(",", "", str1))
    str2 = int(re.sub(",", "", str2))
    result = str1 - str2
    if abs(result) > 18000:
        return result, True
    else:
        return result, False

def getint(str):
    return int(re.sub(",", "", str))

def analysis(var):
    result = ''.join(re.findall(r'[A-Za-z]', var)) 
    # print(result)
    if result in ["MA", "TA", "PF"]:
        return ZZ_analysis(var)
    elif result in ["lu", "sc"]:
        return INE_analysis(var)
    elif result in ["bu", "fu"]:
        return SH_analysis(var)
    elif result in ["eb", "eg", "pg", "pp", "l", "v"]:
        return DL_analysis(var)
    return False

def SH_analysis(var):
    f = open('SH_Hold.dat', 'r')


    trading_name = []
    trading_hold = []
    trading_hold_change = []
    long_name = []
    long_hold = []
    long_hold_change = []
    short_name = []
    short_hold = []
    short_hold_change = []


    SH_file = f.read()
    info = list(filter(None, re.split(r'{ |}', SH_file)))
    for li in info:
        if var in li:
            key_info = re.split(r',', li)
            trading_name.append(key_info[5].split("\"")[3])
            trading_hold.append(key_info[6].split("\"")[2][1:])
            trading_hold_change.append(key_info[7].split("\"")[2][1:])
            long_name.append(key_info[9].split("\"")[3])
            long_hold.append(key_info[10].split("\"")[2][1:])
            long_hold_change.append(key_info[11].split("\"")[2][1:])
            short_name.append(key_info[13].split("\"")[3])
            short_hold.append(key_info[14].split("\"")[2][1:])
            short_hold_change.append(key_info[14].split("\"")[2][1:])

    f.close()
    if trading_hold == []:
        print(var, "数据提取失败")
        status = False
    elif len(trading_name) == len(trading_hold):
        print(var, "数据提取成功")
        status = True
    return trading_name, trading_hold, trading_hold_change, long_name, long_hold, long_hold_change, short_name, short_hold, short_hold_change, status

def INE_analysis(var):
    f = open('INE_Hold.dat', 'r')

    trading_name = []
    trading_hold = []
    trading_hold_change = []
    long_name = []
    long_hold = []
    long_hold_change = []
    short_name = []
    short_hold = []
    short_hold_change = []


    INE_file = f.read()
    info = list(filter(None, re.split(r'{ |}', INE_file)))
    for li in info:
        if var in li:
            key_info = re.split(r',', li)
            trading_name.append(key_info[5].split("\"")[3])
            trading_hold.append(key_info[6].split("\"")[2][1:])
            trading_hold_change.append(key_info[7].split("\"")[2][1:])
            long_name.append(key_info[9].split("\"")[3])
            long_hold.append(key_info[10].split("\"")[2][1:])
            long_hold_change.append(key_info[11].split("\"")[2][1:])
            short_name.append(key_info[13].split("\"")[3])
            short_hold.append(key_info[14].split("\"")[2][1:])
            short_hold_change.append(key_info[14].split("\"")[2][1:])

    f.close()
    if trading_hold == []:
        print(var, "数据提取失败")
        status = False
    elif len(trading_name) == len(trading_hold):
        print(var, "数据提取成功")
        status = True
    return trading_name, trading_hold, trading_hold_change, long_name, long_hold, long_hold_change, short_name, short_hold, short_hold_change, status

def DL_analysis(var):    

    DL_files = get_filename("DL")

    trading_name = []
    trading_hold = []
    trading_hold_change = []
    long_name = []
    long_hold = []
    long_hold_change = []
    short_name = []
    short_hold = []
    short_hold_change = []

    for file in DL_files:
        if var not in file:
            continue

        start = False

        with open(file, 'r') as ope_file:
            file_info = ope_file.readlines()
            if len(file_info) < 440:
                print(file, "Empty info")
                continue

            p_count = 0
           
            for index in range(202, len(file_info)):
                content = file_info[index]
                if "<!-- 列表内容 -->" in content:
                    start = True
                if start == True:
                    if "td" not in content:
                        continue

                    key_info = re.split(r'<+ |>', content)
                    info = key_info[1][:-4]
                    p_count+=1

                    if p_count == 1:
                        continue
                    elif p_count == 2:
                        trading_name.append(info)
                    elif p_count == 3:
                        trading_hold.append(info)     
                    elif p_count == 4:
                        trading_hold_change.append(info)  
                    elif p_count == 5:
                        continue
                    elif p_count == 6:
                        long_name.append(info)  
                    elif p_count == 7:
                        long_hold.append(info)  
                    elif p_count == 8:
                        long_hold_change.append(info)  
                    elif p_count == 9:
                        continue
                    elif p_count == 10:
                        short_name.append(info)  
                    elif p_count == 11:
                        short_hold.append(info)  
                    elif p_count == 12:
                        short_hold_change.append(info)
                        p_count = 0      
    if trading_hold == []:
        print(var, "数据提取失败")
        status = False
    elif len(trading_name) == len(trading_hold):
        print(var, "数据提取成功")
        status = True
    return trading_name, trading_hold, trading_hold_change, long_name, long_hold, long_hold_change, short_name, short_hold, short_hold_change, status

def ZZ_analysis(var):
    ZZ_files = get_filename("ZZ")

    trading_name = []
    trading_hold = []
    trading_hold_change = []
    long_name = []
    long_hold = []
    long_hold_change = []
    short_name = []
    short_hold = []
    short_hold_change = []

    for file in ZZ_files:
        if var[:2] not in file:
            continue
        start = False
        p_start = False

        with open(file, 'r') as ope_file:
            file_info = ope_file.readlines()

            if len(file_info) < 400:
                print(file, "Empty info")
                continue

            p_count = 0
            zj_count = 0

            for index in range(46, len(file_info)):
                content = file_info[index]
                if "colspan" in content:
                    if var in content:
                        start = True
                    else:
                        start = False

                if start:
                    key_info = re.split(r'<+ |>', content)
                    if len(key_info)<=2:
                        continue
                    info = key_info[1][:-4]

                    if info == "增减量":
                        if zj_count <2:
                            zj_count+=1
                        else:
                            p_start = not p_start
                            continue


                    if p_start:
                        p_count += 1

                        if p_count == 1:
                            continue
                        elif p_count == 2:
                            trading_name.append(info)
                        elif p_count == 3:
                            trading_hold.append(info)     
                        elif p_count == 4:
                            trading_hold_change.append(info)  
                        elif p_count == 5:
                            long_name.append(info)  
                        elif p_count == 6:
                            long_hold.append(info)  
                        elif p_count == 7:
                            long_hold_change.append(info)  
                        elif p_count == 8:
                            short_name.append(info)  
                        elif p_count == 9:
                            short_hold.append(info)  
                        elif p_count == 10:
                            short_hold_change.append(info)
                            p_count = 0      

    if trading_hold == []:
        print(var, "数据提取失败")
        status = False
    elif len(trading_name) == len(trading_hold):
        print(var, "数据提取成功")
        status = True
    return trading_name, trading_hold, trading_hold_change, long_name, long_hold, long_hold_change, short_name, short_hold, short_hold_change, status





def writeDataFile(var):
    with open("data.py",'a') as datafile:
        trading_name, trading_hold, trading_hold_change, long_name, long_hold, long_hold_change, short_name, short_hold, short_hold_change, status = analysis(var)

        if status == False:
            print(var, "数据出错，请检查")
            return

        datafile.write(var+"_trading = {")
        for index in range(0, len(trading_name)):
            datafile.write("\""+trading_name[index].strip()+"\":"+str(getint(trading_hold[index]))+",\n")
        datafile.write("}\n\n")

        datafile.write(var+"_trading_change = {")
        for index in range(0, len(trading_name)):
            datafile.write("\""+trading_name[index].strip()+"\":"+str(getint(trading_hold_change[index]))+",\n")
        datafile.write("}\n\n")
    

        datafile.write(var+"_long = {")
        for index in range(0, len(long_name)):
            datafile.write("\""+long_name[index].strip()+"\":"+str(getint(long_hold[index]))+",\n")
        datafile.write("}\n\n")

        datafile.write(var+"_long_change = {")
        for index in range(0, len(long_name)):
            datafile.write("\""+long_name[index].strip()+"\":"+str(getint(long_hold_change[index]))+",\n")
        datafile.write("}\n\n")

        datafile.write(var+"_short = {")
        for index in range(0, len(short_name)):
            datafile.write("\""+short_name[index].strip()+"\":"+str(getint(short_hold[index]))+",\n")
        datafile.write("}\n\n")

        datafile.write(var+"_short_change = {")
        for index in range(0, len(short_name)):
            datafile.write("\""+short_name[index].strip()+"\":"+str(getint(short_hold_change[index]))+",\n")
        datafile.write("}\n\n")

def analysisMain(VARS):
    try:
        for var in VARS:
            writeDataFile(var)
        return True
    except:
        print("数据读取失败，程序中断")
        return False


