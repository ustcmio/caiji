import requests
import demjson
import csv
import base64
from bs4 import BeautifulSoup
import xlrd


def login(s):
    url_login = 'http://10.75.125.115:7002/wxsh/j_unieap_security_check.do'

    body = {
        'j_username': 'A0581305',
        'j_password': '123456',
        'SI_User_Custom_SystemStyle': '#174CAB'
    }

    s.post(url_login, body)


def getMembers(s, pageNum):
    body_query = '{header:{"code":0,"message":{"title":"","detail":""}},body:{dataStores:{"qmcbdjryxx":{rowSet:{"primary":[],"filter":[],"delete":[]},name:"qmcbdjryxx",pageNumber:' + str(
        pageNum) + ',pageSize:20,recordCount:1131,context:{"MENUID":"1496891996303"},statementName:"dcaas.queryInfo",attributes:{"cjzt":["0",12],"rylx":["0",12],"sssq":["320205001002",12]}}},parameters:{"synCount":"true"}}}'
    url_query = 'http://10.75.125.115:7002/wxsh/ria_grid.do?method=query'
    headers = {
        'Accept': '*/*',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN',
        'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET4.0C; .NET4.0E; Tablet PC 2.0)',
        'x-requested-with': 'XMLHttpRequest',
        'ajaxrequest': 'true',
        'Content-Type': 'application/json',
        'Referer': 'http://10.75.125.115:7002/wxsh/si/pages/dcaas/dataCollection/dataCollection.jsp?menuid=1496891996303',
        'Connection': 'Keep-Alive',
        'Pragma': 'no-cache',
        'Host': '10.75.125.115:7002'
    }
    s.headers = headers
    resq = s.post(url_query, body_query).text
    body = demjson.decode(resq)
    members = body['body']['dataStores']['qmcbdjryxx']['rowSet']['primary']
    return [list(member.values()) for member in members]


def loaddownExcel(s, pids, filename):
    if isinstance(pids, list):
        strpid = ','.join(pids)
    else:
        strpid = pids
    # 先使用utf-8编码
    bytesstr = strpid.encode(encoding='utf-8')
    # base64 加密， 再decode()解码
    base64pid = base64.b64encode(bytesstr).decode()
    # print(base64pid)
    # 拼接字符串
    url_report = 'http://10.75.125.115:7002/wxsh/Report-ResultAction.do?reportId=1e77bf40-0007-4cc4-9516-94d5ee920d33' \
                 '&newReport=true&encode=true&aac002s=' + base64pid + \
                 '&xzqh=6ZSh5bGx5Yy65Lic5Lqt6KGX6YGT5p+P5bqE56S+5Yy65bGF5rCR5aeU5ZGY5Lya'
    resq_report = s.get(url_report)

    bs_report = BeautifulSoup(resq_report.text, 'html.parser')
    inputs = bs_report.findAll('input')
    body_excel = {}
    for input in inputs:
        body_excel[input.attrs['name']] = input.attrs['value']

    body_excel['linage'] = body_excel['district']
    body_excel['needProgressBar'] = 'ture'

    url_excel = 'http://10.75.125.115:7002/wxsh/Report-PdfAction.do'
    # s.headers = {
    #     'Host': '10.75.125.115:7002',
    # 'Connection': 'keep - alive',
    # 'Upgrade - Insecure - Requests': '1',
    # 'User - Agent': 'Mozilla / 5.0(Windows NT 6.1) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 55.0.2883.87 Safari / 537.36',
    # 'Content - Type': 'application / x - www - form - urlencoded',
    # 'Accept': 'text / html, application / xhtml + xml, application / xml; q = 0.9, image / webp, * / *;q = 0.8',
    # 'Accept - Encoding': 'gzip, deflate',
    # 'Accept - Language': 'zh - CN, zh;q = 0.8'
    # }
    resq_excel = s.post(url_excel,body_excel)
    with open(filename,'wb') as f:
        f.write(resq_excel.content)

s = requests.Session()
# 1、登录
login(s)
# 2、访问社保系统
url_1 = 'http://10.75.125.115:7002/wxsh/enterapp.do?method=begin&name=/si&welcome=/si/pages/index.jsp'
s.get(url_1)

# 3、获取所有人员名单,
# 注意：修改最大页数58和最大条数recordCount
# all = []
# for index in range(1,58):
#     members = getMembers(s,index)
#     all.extend(members)
#
# with open('2.csv','w') as file:
#     myWriter = csv.writer(file)
#     for member in all:
#         print(member)
#         myWriter.writerow(member)
#     print('done',len(all))

# 4、保存到Excel文件
filename = '名单.xls'
workbook = xlrd.open_workbook(filename)
sheet = workbook.sheet_by_index(0)
begin = 0
members = []
for index in range(1): # (sheet.nrows//20 + 1):
    begin = index *20
    end = begin +20
    if end > sheet.nrows:
        end = sheet.nrows - 1
    members = []
    for i in range(begin,end):
        members.append(sheet.row_values(i)[2])
    print(begin,'~',end)
    loaddownExcel(s,members,str(begin)+'~'+str(end)+'.pdf')