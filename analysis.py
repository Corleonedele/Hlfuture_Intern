
import re
import os


#主流化工品种主力合约：
# 1、郑商所：甲醇是MA2209合约、PTA是TA2209合约、短纤是PF2210。
# 2、大商所：苯乙烯是EB2209、乙二醇是EG2209、液化石油气是PG2209、聚丙烯是PP2209、聚乙烯是L2209、PVC是V2209。
# 3、上期所：沥青是BU2209、高硫燃料油或者直接叫燃料油是FU2301。
# 4、上能所：原油是SC2209、低硫燃料油是LU2209

# 获取当前文件夹下的文件
def get_filename(var=""):
    fl = os.listdir()
    result=[]
    for file in fl:
        if var in file:
            result.append(file)
    if result == []:
        print("没有找到文件")
        return False
    return result

# 对两个string类型的函数进行计算
def intcal(str1, str2):
    str1 = int(re.sub(",", "", str1))
    str2 = int(re.sub(",", "", str2))
    result = str1 - str2
    if abs(result) > 18000:
        return result, True
    else:
        return result, False

# 将string类型转变为integer类型
def getint(str):
    return int(re.sub(",", "", str))

# 分析期货品种属于哪个交易所
def analysis(var, date):
    result = ''.join(re.findall(r'[A-Za-z]', var)) 
    # print(result)
    if result in ["MA", "TA", "PF"]:
        return ZZ_analysis(var, one_file_date=date)
    elif result in ["lu", "sc"]:
        return INE_analysis(var, one_file_date=date)
    elif result in ["bu", "fu"]:
        return SH_analysis(var, one_file_date=date)
    elif result in ["eb", "eg", "pg", "pp", "l", "v"]:
        return DL_analysis(var, one_file_date=date)
    return False



# 各个交易所内容解析思路基本相同，都是从服务器返回的文件里提取相关数据，具体差异为不同交易所的数据存储位置和格式不一样，所以提取的方式不一样


def SH_analysis(var, one_file_status=True, one_file_date=""):
    f = open("SH_"+one_file_date+"_Hold.dat", 'r')

    # 定义局部变量local variable
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


    info = list(filter(None, re.split(r'{ |}', SH_file)))    #此处比较长，用了多个函数，re代表正则表达式是对字符串的解析包，然后对其结果进行筛选，剔除空的空的字符串
    
    #将具体的数值提取到对应的list里
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

def INE_analysis(var, one_file_status=True, one_file_date=""):
    f = open("INE_"+one_file_date+"_Hold.dat", 'r')



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

def DL_analysis(var, one_file_status=True, one_file_date=""):    

    DL_files_tem = get_filename("DL")

    trading_name = []
    trading_hold = []
    trading_hold_change = []
    long_name = []
    long_hold = []
    long_hold_change = []
    short_name = []
    short_hold = []
    short_hold_change = []

    DL_files = []
    for f in DL_files_tem:
        if var in f and one_file_date in f:
            DL_files.append(f)


    for file in DL_files:
        if var not in file:
            continue

        start = False

        with open(file, 'r') as ope_file:
            file_info = ope_file.readlines()
            if len(file_info) < 400:
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

def ZZ_analysis(var, one_file_status=True, one_file_date=""):
    ZZ_files_tem = get_filename("ZZ")

    trading_name = []
    trading_hold = []
    trading_hold_change = []
    long_name = []
    long_hold = []
    long_hold_change = []
    short_name = []
    short_hold = []
    short_hold_change = []

    ZZ_files = []
    for f in ZZ_files_tem:
        if var[:2] in f and one_file_date in f:
            ZZ_files.append(f)
    

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




def writeDataFile(var, date):
    with open("data.py",'a') as datafile:
        trading_name, trading_hold, trading_hold_change, long_name, long_hold, long_hold_change, short_name, short_hold, short_hold_change, status = analysis(var, date)

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


# 数据解析主函数
def analysisMain(VARS, date):
    try:
        for var in VARS:
            writeDataFile(var, date)
        return True
    except:
        print("数据读取失败，程序中断")
        return False

