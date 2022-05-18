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
            try:
                element = Login.browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/div[1]/span[1]')
                flag = 2
                break
            except:
                sleep(0.5)
                continue
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

            element = Login.browser.find_element_by_id("billaccountTypeQ")  # 定位报账点类型
            Select(element).select_by_visible_text(benchmark_Sheet.Range("B5").value)

            element = Login.browser.find_element_by_id("overflow")  # 定位是否超标
            Select(element).select_by_visible_text(benchmark_Sheet.Range("B6").value)

            element = Login.browser.find_element_by_id("userCodeOrName")  # 定位录入人
            element.clear()
            element.send_keys(benchmark_Sheet.Range("B7").value)

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
                element = Login.browser.find_element_by_xpath('//*[@id="form"]/button[2]')  # 定位查询按钮
                ActionChains(Login.browser).move_to_element(element).perform()
                element.click()
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
            try:
                element = Login.browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/div[1]/span[1]')
                flag = 5
                break
            except:
                sleep(2)
                continue
    # 判断是否有查询到数据
    if flag == 5:
        flag_break = 0
        while 1:
            if flag_break == 10:
                print("查询不到数据，请确认查询条件")
                break
            flag_break = flag_break + 1
            try:
                element = Login.browser.find_element_by_xpath('//*[@id="tb"]/tbody/tr[1]/td[4]')  # 湛江
                flag = 6
                break
            except:
                sleep(0.5)
                continue
    # 如查询的数据超一页时，选择显示最大行数
    if flag == 6:
        try:
            element = Login.browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/div[1]/span[2]/span/ul')     # 显示行数
            element_list = element.find_elements_by_xpath('li')
            # print("行数显示选择为：" + str(len(element_list)))
            if len(element_list) > 1:
                element = Login.browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/div[1]/span[2]/span/button')  # 点击显示行数按钮
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
                element = Login.browser.find_element_by_xpath('/html/body/div[4]/div[2]/div[4]/div[1]/span[1]')     # 从显示行数判断是否加载完成
                flag = 8
                break
            except:
                sleep(2)
                continue
    # 取超标值
    if flag == 8:
        irow = 3
        try:
            sleep(5)
            element = Login.browser.find_element_by_xpath('//*[@id="tb"]/tbody')    # 内容表格窗口主体
            element_list = element.find_elements_by_tag_name('tr')      # 行数
        except:
            traceback.print_exc()
            pass
        for i in range(len(element_list)):
            try:
                list_sheet.Range("A" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[5]').text  # 区域
                list_sheet.Range("D" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="tb"]/tbody/tr[' + str(i+1) + ']/td[8]').text  # 报账点名称
                list_sheet.Range("F" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[6]/a').text  # 报账点编码
                list_sheet.Range("G" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[7]').text    # 缴费单号
                list_sheet.Range("I" + str(irow)).NumberFormat = '@'
                list_sheet.Range("I" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[12]/a/span').text    # 缴费期始
                list_sheet.Range("J" + str(irow)).NumberFormat = '@'
                list_sheet.Range("J" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[13]/a/span').text    # 缴费期终
            except:
                traceback.print_exc()
                pass
            try:
                accountLink = Login.browser.find_element_by_xpath(
                    '//*[@id="tb"]/tbody/tr[' + str(i + 1) + ']/td[6]/a').get_attribute('href')  # 获取链接
                js = 'window.open("' + accountLink + '")'
                Login.browser.execute_script(js)    # 用js打开链接
                all_handle = Login.browser.window_handles
                for handle in all_handle:
                    if handle != Login.main_Page or handle != query_page:
                        tem_page = handle
                        Login.browser.switch_to.window(tem_page)  # 切换至新开网页
                flag_break = 0
                # 5秒等待网页加载完成
                while 1:
                    if flag_break == 10:
                        print("哦豁，出错了，网页加载失败...")
                        break
                    flag_break = flag_break + 1
                    try:
                        element = Login.browser.find_element_by_id('paymentDataBtn')
                        break
                    except:
                        sleep(0.5)
                        continue
                list_sheet.Range("Y" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="benchmarktb"]/tbody/tr[1]/td[5]/font').text  # 同比标杆
                list_sheet.Range("P" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="benchmarktb"]/tbody/tr[2]/td[5]/font').text  # 环比标杆
                list_sheet.Range("AK" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="benchmarktb"]/tbody/tr[4]/td[5]/font').text  # 动环标杆
                list_sheet.Range("AL" + str(irow)).value = Login.browser.find_element_by_xpath(
                    '//*[@id="benchmarktb"]/tbody/tr[3]/td[5]/font').text  # 额定标杆
                Login.browser.close()   # 关闭当前窗口
                Login.browser.switch_to.window(query_page)  # 切换至明细维护页面
            except:
                traceback.print_exc()
                print('获取链接失败...')
                pass
            irow = irow + 1
    # 保存并关闭excel，并退出引用
    try:
        excel.Close(SaveChanges=True)
        xlsm.Application.Quit()
    except:
        traceback.print_exc()
        pass
