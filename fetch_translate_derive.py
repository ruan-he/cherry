import pdfplumber
import re
import xlsxwriter
import random
import hashlib
import urllib.parse
import requests
import time


path = './Specimen.pdf'
outpath = path.replace('.pdf','',1).replace('./','',1)+'.xlsx'
# 创建表格
workbook = xlsxwriter.Workbook(outpath)

#选择表单
worksheet = workbook.add_worksheet("生词本")

# Widen the column to make the text clearer.
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 20)

pdf = pdfplumber.open(path)#打开pdf文件
pages = pdf.pages#获取页数信息
total_pages = len(pages)#获取总页数

#遍历每页上的文字并传到列表中
outList = []
resultList = []
PureOutList  = []
pat = '[a-zA-Z]+'#正则表达式
for i in range(0,total_pages):
    text = pdf.pages[i].extract_text()
    outList = re.findall(pat,text)#用正则表达式提取字符串中的单词,有重复
    for i in outList:
        resultList.append(i)#收集每页的单词
#去除重复单词并传到PureOutList
for i in resultList:
    if not i in PureOutList:
        PureOutList.append(i)

#翻译
appid = ''  # 你的appid
secretKey = ''  # 你的密钥
fromLang = 'en'
toLang = 'zh'
salt =  random.randint(1111111111, 9999999999)  # 生成随机值，可以设置为固定的数值
# print(f"salt={salt}")
#写入数据
for i in range(0,len(PureOutList)):
    words = PureOutList[i]
    q = words
    sign = appid + q + str(salt) + secretKey    # 拼接签名
    # print(f"sign={sign}")
    # q = urllib.parse.quote(q)
    sign = hashlib.md5(sign.encode()).hexdigest()
    # print(sign)
    # 拼接url
    url = f'http://api.fanyi.baidu.com/api/trans/vip/translate?q={urllib.parse.quote(q)}&from={fromLang}&to={toLang}&appid={appid}&salt={salt}&sign={sign}'
    # print(url)
    res = requests.get(url).json()#获取返回内容
    # print(res)
    # print(type(res))
    # print(res['trans_result'])
    # print(type(res['trans_result']))
    # print(res['trans_result'][0])
    # print(type(res['trans_result'][0]))
    # print(res['trans_result'][0]['dst'])
    result = res['trans_result'][0]['dst']   # 筛选到翻译结果
    WordsPosition = 'A'+str(i)
    ResPosition = 'B'+str(i)
    worksheet.write(WordsPosition,words)#（行，列）
    worksheet.write(ResPosition,result)
    print(i)
    time.sleep(0.1)#百度翻译API限制每秒访问次数

workbook.close()
