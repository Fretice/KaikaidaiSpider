import requests
import datetime
from pyquery.pyquery import PyQuery as pq
from openpyxl import Workbook
from openpyxl import load_workbook

#解析网页结构并返回需要的数据方法
def get_data(url, page_index, viewstate):
    param = {"__EVENTTARGET": "rpMessage", "__EVENTARGUMENT": "pager$" + str(page_index), "__VIEWSTATE": viewstate,
             "rpMessage": ''}
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'}
    r = requests.post(url, data=param, headers=headers)
    d = pq(r.text)
    tab_count = d(".hmd_ytab").length
    person_data_list=[]
    for i in range(tab_count):
        name = d(".hmd_ytab").eq(i).find("tr").eq(0).find("td").eq(2).find("a").text()
        email = d(".hmd_ytab").eq(i).find("tr").eq(0).find("td").eq(4).text().replace("\r\n", "").replace(" ", "")
        overdue_count = d(".hmd_ytab").eq(i).find("tr").eq(0).find("td").eq(6).text().replace("\r\n", "").replace(" ", "")

        person_idcard = d(".hmd_ytab").eq(i).find("tr").eq(1).find("td").eq(1).text().replace("\r\n", "").replace(" ", "")
        home_phone = d(".hmd_ytab").eq(i).find("tr").eq(1).find("td").eq(3).text().replace("\r\n", "").replace(" ", "")
        website_pay_count = d(".hmd_ytab").eq(i).find("tr").eq(1).find("td").eq(5).text().replace("\r\n", "").replace(" ", "")

        address = d(".hmd_ytab").eq(i).find("tr").eq(2).find("td").eq(1).text().replace("\r\n", "").replace(" ","")
        mobile_number = d(".hmd_ytab").eq(i).find("tr").eq(2).find("td").eq(3).text().replace("\r\n", "").replace(" ", "")
        overdue_day = d(".hmd_ytab").eq(i).find("tr").eq(2).find("td").eq(5).text().replace("\r\n", "").replace(" ", "")

        company_name = d(".hmd_ytab").eq(i).find("tr").eq(3).find("td").eq(1).text().replace("\r\n", "").replace(" ","")
        overdue_money_total = d(".hmd_ytab").eq(i).find("tr").eq(3).find("td").eq(5).text().replace("\r\n", "").replace(" ",
                                                                                                                   "")
        company_address = d(".hmd_ytab").eq(i).find("tr").eq(4).find("td").eq(2).text().replace("\r\n", "").replace(" ",
                                                                                                                   "")
        person_data_list.append([name, email, overdue_count, person_idcard, home_phone, website_pay_count, address, mobile_number, overdue_day, company_name, overdue_money_total, company_address])

    return person_data_list

#获取第一页中隐藏域值的方法
def get_url_param(url):
    r = requests.get(url)
    d = pq(r.text)
    return d("#__VIEWSTATE")

#获取非首页页面中隐藏域值的方法
def get_url_param_new(url, page_index, viewstate):
    param = {"__EVENTTARGET": "rpMessage", "__EVENTARGUMENT": "pager$" + str(page_index), "__VIEWSTATE": viewstate,
             "rpMessage": ''}
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'}
    r = requests.post(url)
    d = pq(r.text)
    return d("#__VIEWSTATE")

#获取的数据写入Excel文件方法
def write_file(file_name):
    wb = Workbook()
    ws = wb.active
    ws.append(
        ['姓名', '邮箱', '逾期未还款数目', '身份证', '电话', '网站垫付款数目', '地址', '手机', '最长逾期天数', '公司名称', '逾期待还总额', '公司地址'])
    file_name = file_name + '.xlsx'
    wb.save(file_name)
    url = "http://www.kaikaidai.com/Lend/Black.aspx"
    data_param = get_url_param(url)
    for i in range(39):
        print("正在爬取第"+str(i)+"页的内容。开始时间："+str(datetime.datetime.utcnow()))
        data_list = get_data(url, i, data_param)
        data_param = get_url_param_new(url, i, data_param)
        for data in data_list:
            ws.append([data[0],data[1],data[2],data[3],data[4],data[5],data[6],data[7],data[8],data[9],data[10],data[11]])
            wb.save(file_name)
        print("第"+str(i)+"页的内容保存成功。结束时间："+str(datetime.datetime.utcnow()))

filename = input("请输入你想保存的文件名称：")
write_file(filename)




