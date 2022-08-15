from random import randint, choice

order_range = ["1", "2", "3", "4", "5", "6", "7", "8", "9"]

trade_office_official_url = {
    "DL":"http://www.dce.com.cn/", #大连商品交易所
    "ZJ":"http://www.cffex.com.cn/", #中国金融期货交易所
    "ZZ":"http://www.czce.com.cn/", #郑州商品交易所
    "INE":"https://www.ine.cn/", #上海国际能源交易中心
    "SH":"https://www.shfe.com.cn/", #上海期货交易所
}

user_agent = [
    "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11",
    "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; en) Opera 9.50",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.122 Safari/537.36",
    "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.16 (KHTML, like Gecko) Chrome/10.0.648.133 Safari/534.16",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_0) AppleWebKit/535.11 (KHTML, like Gecko) "
]

# 生成特定长度的随机字符串
def random_number(times):
    result = ""
    for i in range(times):
        result += choice(order_range)
    return result

def ZJ_URL_Daily(require_time="20220801"):
    return "http://www.cffex.com.cn/sj/hqsj/rtj/" + require_time[:6]+"/"+require_time[6:] + "/index.xml?id=" + str(randint(1,100))

def ZZ_URL_Daily(require_time="20220801"):
    return "http://www.czce.com.cn/cn/DFSStaticFiles/Future/2022/" + require_time + "/FutureDataDaily.htm"

def INE_URL_Daily(require_time="20220801"):
    return "https://www.ine.cn/data/dailydata/kx/kx"+require_time+".dat?temp2="+random_number(13)+"?rnd=0."+random_number(16)+""

def SH_URL_Daily(require_time="20220801"):
    return "https://www.shfe.com.cn/data/dailydata/kx/kx"+require_time+".dat"

def DL_URL_Daily(require_time="20220801"):
    return "http://www.dce.com.cn/publicweb/quotesdata/dayQuotesCh.html?"+"dayQuotes.variety=all&dayQuotes.trade_type=0&year="+require_time[:4]+"&month="+str(int(require_time[4:6])-1)+"&day="+require_time[6:]

def ZZ_URL_HOLD_Daily(require_time="20220801", require_type=""):
    """require_type are MA, PF, TA"""
    return "http://www.czce.com.cn/cn/DFSStaticFiles/Future/2022/"+require_time+"/FutureDataHolding"+require_type+".htm"

def DL_URL_HOLD_Daily(require_time="20220801", require_type="", future_or_option="0"):
    """require_type are 种类+数字 eg. eg2209"""
    if len(require_type) == 6:
        return "http://www.dce.com.cn/publicweb/quotesdata/memberDealPosiQuotes.html?memberDealPosiQuotes.variety="+require_type[:2]+"&memberDealPosiQuotes.trade_type="+future_or_option+"&year="+require_time[:4]+"&month="+str(int(require_time[4:6])-1)+"&day="+require_time[6:]+"&contract.contract_id="+require_type+"&contract.variety_id="+require_type[:2]+"&contract="
    else:
        return "http://www.dce.com.cn/publicweb/quotesdata/memberDealPosiQuotes.html?memberDealPosiQuotes.variety="+require_type[:1]+"&memberDealPosiQuotes.trade_type="+future_or_option+"&year="+require_time[:4]+"&month="+str(int(require_time[4:6])-1)+"&day="+require_time[6:]+"&contract.contract_id="+require_type+"&contract.variety_id="+require_type[:1]+"&contract="
    
def SH_URL_HOLD_Daily(require_time="20220801"):
    return "https://www.shfe.com.cn/data/dailydata/kx/pm"+require_time+".dat"

def INE_URL_HOLD_Daily(require_time="20220801"):
    return "https://www.ine.cn/data/dailydata/kx/pm"+require_time+".dat?temp2="+random_number(13)+"?rnd=0."+random_number(16)
