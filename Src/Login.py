# uft-8
# Login.py

from selenium import webdriver
from selenium.webdriver import ActionChains
from Src import ReadConfig
from time import sleep
import json

browser = ''
main_Page = ''

def login():
    global browser, main_Page,cookie
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    browser = webdriver.Chrome(options=options)
    browser.maximize_window()
    # browser = webdriver.Chrome()
    url = 'http://10.217.240.219:8090/NCMS/welcome'
    # 初次建立连接，随后方可修改cookie
    browser.get(url)
    # 删除第一次建立连接时的cooki
    browser.delete_all_cookies()
    # 读取登录时存储到本地的cookie
    try:
        with open('cookies.json', 'r', encoding='utf-8') as f:
            listCookies = json.loads(f.read())
        for cookie in listCookies:
            browser.add_cookie({
                'domain': '10.217.240.219',  # 此处xxx.com前，需要带点
                'name': cookie['name'],
                'value': cookie['value'],
                'path': '/',
                'expires': None
            })
        browser.get(url)
    except:
        browser.get(url)

def enterPassword():
    userinput = input("需要登录并获取cookie?按1继续\n>>>")
    if userinput == "1":
        element = browser.find_element_by_id("area_name")
        ActionChains(browser).move_to_element(element).perform()
        element.click()
        #element = browser.find_element_by_xpath('//*[@id="areaFor"]/div[6]/div[3]/span[1]')
        element = browser.find_element_by_xpath("//span[contains(text(),'广东')]")
        ActionChains(browser).move_to_element(element).perform()
        element.click()

        main_Page = browser.current_window_handle

        # 用户账号、密码填写
        while 1:
            nameinput = input("请选择登录账号：\n1 梁湛波\n2 刘以鹏\n>>>")
            if nameinput == "1":
                element = browser.find_element_by_id("userLoginname")
                element.send_keys(ReadConfig.user1)
                element = browser.find_element_by_id("userPassword")
                element.send_keys(ReadConfig.password1)
                break
            elif nameinput == "2":
                element = browser.find_element_by_id("userLoginname")
                element.send_keys(ReadConfig.user2)
                element = browser.find_element_by_id("userPassword")
                element.send_keys(ReadConfig.password2)
                break
        # 空闲8秒钟，等待输入验证码
        sleep(8)
        # 输入验证码后点击登录按钮
        try:
            element = browser.find_element_by_id("login")
            ActionChains(browser).move_to_element(element).perform()
            element.click()
        except:
            pass
        main_Page = browser.current_window_handle
        # 获取cookies
        try:
            dictCookies = browser.get_cookies()
            jsonCookies = json.dumps(dictCookies)
            # 登录完成后，将cookie保存到本地文件
            with open('cookies.json', 'w') as f:
                f.write(jsonCookies)
        except:
            print('获取cookies失败...')
            pass