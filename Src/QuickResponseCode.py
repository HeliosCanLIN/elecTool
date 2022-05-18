#   QuickResponseCode.py
#   coding: utf-8

import qrcode
from selenium import webdriver
import pyzbar
import pyzbar.pyzbar as pyzbar
import pynput
from PIL import ImageGrab, Image, ImageEnhance
import os
from Src import ReadConfig
from time import sleep


def quickResponseCode(code_data):
    qr = qrcode.QRCode(
        version=4,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=4,
        border=4
    )
    qr.add_data(code_data)
    qr.make(fit=True)
    img = qr.make_image()
    # img.show()
    img.save(ReadConfig.savePath + code_data + ".jpg")

def decodeDisplay():
    pdf_path = ReadConfig.pdf_path
    file_path = ReadConfig.newNamePath
    driver = webdriver.Chrome()
    driver.maximize_window()  # 设置窗口最大化
    chrome = driver.get("file:///"+ReadConfig.main_pdf)     # 用谷歌浏览器打开PDF文件
    main_page = driver.current_window_handle
    for file in os.listdir(pdf_path):
        tem_pdf_file = ''
        # print(file)
        tem_pdf_file = pdf_path+'/'+file
        sleep(1)
        try:
            js = "window.open('file:///" + tem_pdf_file + "')"
            # print(js)
            driver.execute_script(js)
            all_handle = driver.window_handles
            for handle in all_handle:
                if handle != main_page:
                    tem_handle = handle
                    driver.switch_to.window(tem_handle)
        except Exception as e:
            print(e)
            pass
        else:
            # 手工选择二维码所在位置，双击自动获取二维码位置
            # print("滚动网页到二维码处，并点击二维码中心位置")
            with pynput.mouse.Events() as event:
                for i in event:
                    if isinstance(i, pynput.mouse.Events.Click):
                        # print(i.x, i.y, i.button, i.pressed)
                        img = ImageGrab.grab(bbox=(i.x-120, i.y-120, i.x+120, i.y+120))
                        # print(img_path)
                        img.save(ReadConfig.img_path)
                        img = Image.open(ReadConfig.img_path)
                        # img.show()
                        results = pyzbar.decode(img)
                        if len(results):
                            url = results[0].data.decode("utf-8")
                            # print(url)
                            driver.close()
                            autoRename(tem_pdf_file, file_path+'\\'+url)
                        else:
                            print("读取二维码失败...")
                        break
        driver.switch_to.window(main_page)
    print("重命名完成...")
    driver.close()

def autoRename(oldName, newName):
    try:
        os.rename(oldName, newName)
    except Exception as e:
        print(e)
    else:
        print("重命名成功："+newName)
