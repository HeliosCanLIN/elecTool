# GetBenchmark.py
# coding: utf-8

from Src import Login
from Src import ReadConfig
import traceback
from time import sleep
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import win32com
import win32com.client

import queue #队列
import re #正则表达式
import requests #发送报文请求
from bs4 import BeautifulSoup #处理回复报文html

# 定义全局变量
query_page = ''
xlsm = ''
excel = ''
list_sheet = ''
benchmark_Sheet = ''
flag = 0

def open_payment_finance():
    global query_page   # 需要修改全局变更时加global，否则当成函数内的局部变更
    try:
        js = 'window.open("http://' + ReadConfig.query_url + '")'
        Login.browser.execute_script(js)
        # 根据网页句柄跳转到查询明细页面
        all_handle = Login.browser.window_handles   # 获取全部已打开网页的句柄
        for handle in all_handle:
            if handle != Login.main_Page:
                query_page = handle
                Login.browser.switch_to.window(query_page)  # 切换至明细查询页面
    except:
        traceback.print_exc()
        Login.browser.switch_to.window(Login.main_Page)  # 切换至登录界面

def open_excel():
    global flag, xlsm, excel, list_sheet, benchmark_Sheet
    try:
        xlsm = win32com.client.Dispatch('Excel.Application')    # 引用excel应用
        xlsm.Visible = 1  # 文档可见
        excel = xlsm.Workbooks.Open(ReadConfig.benchmarkPath + ReadConfig.benchmarkName)  # excel表格打开
        list_sheet = excel.Worksheets(ReadConfig.benchmarkList)     # 清单子表
        benchmark_Sheet = excel.Worksheets(ReadConfig.benchmarkSheet)   # 筛选子表
        list_sheet.Application.Run("Sheet2.clean")
        print('清除清单...')
        flag = 1    # 表格打开标志
    except:
        flag = 0
        print('打开直供电同环比数据excel表失败....')

def getBenchmark():
    global flag, query_page, xlsm, excel, list_sheet, benchmark_Sheet
    open_payment_finance()  # 打开查询网页
    open_excel()    # 打开直供电同环比数据excel表
    # 网页成功打开时进行下一步
    if flag == 1:
        # 等待网页加载完成
        while 1:
            try:  # 检查加载的时候会有loading的class 尝试到检测不到loading再继续
                element = Login.browser.find_element_by_class_name(
                    'ui_loading_hide')  # 根据网页是否加载完成“显示第 1 到第 1 条记录，总共 1 条记录”这个元素判断是否加载完成
                print('等待加载中...')
                sleep(2)
                continue
            except:
                flag = 2
                sleep(0.5)
                break
    sleep(2)
    # 查询条件填写
    if flag == 2:
        try:
            element = Login.browser.find_element_by_id("city")  # 定位城市
            Select(element).select_by_visible_text(benchmark_Sheet.Range("B1").value)

            element = Login.browser.find_element_by_id("region")  # 定位区域
            Select(element).select_by_visible_text(benchmark_Sheet.Range("B2").value)

            element = Login.browser.find_element_by_id("auditingStateQ")  # 定位审核状态
            Select(element).select_by_visible_text(benchmark_Sheet.Range("B3").value)

            element = Login.browser.find_element_by_id("supplyMethod")  # 定位供电类型
            Select(element).select_by_visible_text(benchmark_Sheet.Range("B4").value)

            #element = Login.browser.find_element_by_id("billaccountTypeQ")  # 定位报账点类型
            #Select(element).select_by_visible_text(benchmark_Sheet.Range("B5").value)

            element = Login.browser.find_element_by_id("overflow")  # 定位是否超标
            Select(element).select_by_visible_text(benchmark_Sheet.Range("B6").value)

            #element = Login.browser.find_element_by_id("userCodeOrName")  # 定位录入人
            #element.clear()
            #element.send_keys(benchmark_Sheet.Range("B7").value)

            element = Login.browser.find_element_by_id("billamountDateOpen")  # 定位申请日期开始
            element.send_keys(Keys.CONTROL, 'a')
            element.send_keys(str(benchmark_Sheet.Range("B8").value))

            element = Login.browser.find_element_by_id("billamountDateClose")  # 定位申请日期截止
            element.send_keys(Keys.CONTROL, 'a')
            element.send_keys(str(benchmark_Sheet.Range("B9").value))

            flag = 3
        except:
            traceback.print_exc()
            pass
    # 点击查询按钮
    if flag == 3:
        flag_break = 0
        while 1:
            if flag_break == 10:
                print("哦豁，出错了，在定位查询按钮时...")
                break
            flag_break = flag_break + 1
            try:
                sleep(1)
                Login.browser.execute_script('loadTableData()') #查询按钮的js
                """
                element = Login.browser.find_element_by_xpath('//*[@id="form"]/button[2]')  # 定位查询按钮
                ActionChains(Login.browser).move_to_element(element).perform()
                element.click()
                """
                sleep(1)
                flag = 4
                break
            except:
                sleep(1)
                continue
    # 等待网页加载完成
    if flag == 4:
        flag_break = 0
        # 等待网页加载完成
        while 1:
            if flag_break == 10:
                print("哦豁，出错了，网页加载失败...")
                break
            flag_break = flag_break + 1
            try:#检查加载的时候会有loading的class 尝试到检测不到loading再继续
                element = Login.browser.find_element_by_class_name(
                    'ui_loading_hide')  # 根据网页是否加载完成“显示第 1 到第 1 条记录，总共 1 条记录”这个元素判断是否加载完成
                print('等待加载中...')
                sleep(2)
                continue
            except:
                flag = 5
                sleep(3)
                break
    # 判断是否有查询到数据
    if flag == 5:
        flag_break = 0
        while 1:
            if flag_break == 10:
                print("查询不到数据，请确认查询条件")
                break
            flag_break = flag_break + 1
            try:
                element = Login.browser.find_element_by_name('btSelectItem')  # 湛江
                flag = 6
                break
            except:
                sleep(0.5)
                continue
    # 如查询的数据超一页时，选择显示最大行数
    if flag == 6:
        try:
            element = Login.browser.find_elements_by_class_name('dropdown-menu')[3]
            #element = Login.browser.find_elements_by_xpath('//ul[contains(@class,\'dropdown-menu\') and contains(@role,\'menu\')]')
            #element = Login.browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[4]/div[1]/span[2]/span/ul')# 显示行数
            element_list = element.find_elements_by_xpath('li')
            # print("行数显示选择为：" + str(len(element_list)))
            if len(element_list) > 1:
                #element = Login.browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[4]/div[1]/span[2]/span/button')  # 点击显示行数按钮
                element = Login.browser.find_elements_by_class_name('dropdown-toggle')[1]
                ActionChains(Login.browser).move_to_element(element).perform()
                element.click()
                sleep(0.8)
                element_list[-1].click()  # 按倒数第一个显示行数
                sleep(0.5)
            flag = 7
        except:
            print('哦豁，出错了，查找li元素不成功...')
            traceback.print_exc()
            pass
    # 等待网页加载完成
    if flag == 7:
        flag_break = 0
        # 等待网页加载完成
        while 1:
            if flag_break == 10:
                print("哦豁，出错了，网页加载失败...")
                break
            flag_break = flag_break + 1
            try:
                element = Login.browser.find_element_by_class_name('ui_loading_hide')# 根据网页是否加载完成“显示第 1 到第 1 条记录，总共 1 条记录”这个元素判断是否加载完成
                print('等待加载中...')
                sleep(2)
                continue
            except:
                flag = 8
                sleep(4)
                break
    # 取超标值
    if flag == 8:
        irow = 3
        while True:
            q = queue.Queue()
            url = "http://10.217.240.219:8087/NCMS-TELE/asserts/tpl/tele/payment/showBenchmark"  # 接口地址
            """
            flag_break = 0
            # 等待网页加载完成
            while 1:
                if flag_break == 10:
                    print("哦豁，出错了，网页加载失败...")
                    break
                flag_break = flag_break + 1
                try:
                    element = Login.browser.find_element_by_xpath(
                        '/html/body/div[3]/div[2]/div[4]/div[1]/span[1]')  # 从显示行数判断是否加载完成
                    flag = 8
                    break
                except:
                    sleep(2)
                    continue
            """
            i = 0
            for link in Login.browser.find_elements_by_xpath("//tbody//*[@href]"):  # 此循环请求速度极快 有被服务器ban掉的可能
                accountLink = link.get_attribute('href')
                accountId = re.findall(r"billaccountpaymentdetailId=(.+?)&", accountLink, flags=0)
                # 报文header 消息头数据
                headers = {
                    'Connection': 'keep-alive',
                    'Content-Length': '59',
                    'Accept': '*/*',
                    'DNT': '1',
                    'X-Requested-With': 'XMLHttpRequest',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.125 Safari/537.36',
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'Origin': 'http://10.217.240.219:8087',
                    'Referer': accountLink,
                    'Accept-Encoding': 'gzip, deflate',
                    'Accept-Language': 'zh-CN,zh;q=0.9',
                    'Cookie': '/NCMS/welcomecuNum=1; SESSION=' + Login.cookie['value'] + ';sessionId=' + Login.cookie[
                        'value'],

                }
                data = {'billaccountpaymentdetailId': accountId[0]}  # 报文body 消息数据
                r = requests.post(url, headers=headers, data=data, verify=False)  # 发送POST请求标杆
                packet = r.json()  # 获取回复
                html = packet['obj']  # 提取html
                try:
                    soup = BeautifulSoup(html, 'html.parser')  # 将html实例化为BeautifulSoup对象
                except:
                    print("html实例化失败或登录失效,检查你的cookie是否正常获取")
                    print(packet)
                for item in soup.find_all("td"):  # 遍历html提取table
                    benchmark = re.sub('\s', ' ', item.get_text())  # 去除html标签
                    q.put(benchmark)  # 数值推入队列

                j = 1
                for j in range(1, q.qsize() + 1):
                    if (j == 5):
                        list_sheet.Range("Y" + str(irow)).value = q.get()  # 同比标杆
                    elif (j == 10):
                        list_sheet.Range("P" + str(irow)).value = q.get()  # 环比标杆
                    elif (j == 15):
                        list_sheet.Range("AL" + str(irow)).value = q.get()  # 额定标杆
                    elif (j == 20):
                        list_sheet.Range("AK" + str(irow)).value = q.get()  # 动环标杆
                    else:
                        q.get()
                try:
                    print('已经获取'+str(irow-2)+'条数据')
                    list_sheet.Range("A" + str(irow)).value = Login.browser.find_element_by_xpath(
                        '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[5]').text  # 区域
                    list_sheet.Range("D" + str(irow)).value = Login.browser.find_element_by_xpath(
                        '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[9]').text  # 报账点名称
                    list_sheet.Range("F" + str(irow)).value = Login.browser.find_element_by_xpath(
                        '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[6]/a').text  # 报账点编码
                    list_sheet.Range("G" + str(irow)).value = Login.browser.find_element_by_xpath(
                        '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[7]').text  # 缴费单号
                    list_sheet.Range("I" + str(irow)).NumberFormat = '@'
                    list_sheet.Range("I" + str(irow)).value = Login.browser.find_element_by_xpath(
                        '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[13]/a/span').text  # 缴费期始
                    list_sheet.Range("J" + str(irow)).NumberFormat = '@'
                    list_sheet.Range("J" + str(irow)).value = Login.browser.find_element_by_xpath(
                        '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[14]/a/span').text  # 缴费期终
                    i += 1
                except:
                    traceback.print_exc()
                    pass
                irow = irow + 1
            if len(Login.browser.find_elements_by_name('btSelectItem'))<500:
                break
            else:
                i=0
                element = Login.browser.find_element_by_xpath("//a[contains(text(),'›')]")
                element.click()
                print('翻页...')
                sleep(7)#使用睡眠等待第二页 载入检测里面有break的话 好像在多层loop里面会有问题
    # 保存并关闭excel，并退出引用
    try:
        excel.Close(SaveChanges=True)
        xlsm.Application.Quit()
    except:
        traceback.print_exc()
        pass
