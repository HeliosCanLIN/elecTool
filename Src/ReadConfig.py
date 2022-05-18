# uft-8
# ReadConfig.py

import yaml
import os
import traceback

ulr = 'http://10.217.240.219:8090/NCMS/welcome'
import_ulr = 'http://10.217.240.219:8090/NCMS/asserts/tpl/selfelec/payment_finance/record.html'
userName1 = ''
user1 = ''
password1 = ''
userName2 = ''
user2 = ''
password2 = ''
importName = ''
importPath = ''
outPath = ''
toolName = ''
toolPath = ''
toolSheet = ''
listSheet = ''
importSheet = ''
attachmentPath = ''
submitName = ''
submitPath = ''
submitSheet = ''
submitList = ''
oversubmitName = ''
oversubmitPath = ''
oversubmitList = ''
benchmarkName = ''
benchmarkPath = ''
benchmarkList = ''
benchmarkSheet = ''
savePath = ''
docPath = ''
docNameH = ''
docNameT = ''
docNameTH = ''
StName = ''
query_url = ''
pdf_path = ''
newNamePath = ''
img_path = ''
main_pdf = ''

def readconfig():
    global query_url
    global userName1, user1, password1, userName2, user2, password2 , ulr, importName, importPath, outPath, toolName
    global toolPath, importSheet, toolSheet, listSheet, attachmentPath, submitList, submitName, submitPath, submitSheet
    global oversubmitName, oversubmitPath, oversubmitList
    global benchmarkList, benchmarkPath, benchmarkName, benchmarkSheet
    global savePath, docPath, docNameH, docNameT, docNameTH, StName
    global img_path, newNamePath, pdf_path, main_pdf
    # 打开yaml文件
    try:
        config_path = os.path.abspath('.')
        # config_path = os.path.dirname(os.path.abspath('.'))     # 获取当前工作目录的父级目录的绝对路径
        config_file = open(config_path+'/config.yaml', 'r', encoding="utf-8")
    except:
        traceback.print_exc()
        print('打开配置文档失败...')

    # 读取yaml文件内容
    try:
        config = yaml.load(config_file, Loader=yaml.FullLoader)
        userName1 = config['userName1']     # 梁湛波账号
        user1 = config['user1']
        password1 = config['password1']

        userName2 = config['userName2']     # 刘以鹏账号
        user2 = config['user2']
        password2 = config['password2']

        importPath = config['importPath']
        importName = config['importName']
        importSheet = config['importSheet']

        outPath = config['outPath']

        toolPath = config['toolPath']
        toolName = config['toolName']
        toolSheet = config['toolSheet']
        listSheet = config['listSheet']

        # 附件位置
        attachmentPath = config['attachmentPath']

        # 提单清单
        submitList = config['submitList']
        submitName = config['submitName']
        submitPath = config['submitPath']
        submitSheet = config['submitSheet']

        oversubmitName = config['oversubmitName']
        oversubmitPath = config['oversubmitPath']
        oversubmitList = config['oversubmitList']

        # 直供电同环比清单
        benchmarkList = config['benchmarkList']
        benchmarkPath = config['benchmarkPath']
        benchmarkName = config['benchmarkName']
        benchmarkSheet = config['benchmarkSheet']

        # 同环比文档保存位置
        savePath = config['savePath']
        StName = config['StName']
        docPath = config['docPath']
        docNameH = config['docNameH']
        docNameT = config['docNameT']
        docNameTH = config['docNameTH']

        # 各种链接
        query_url = config['query_url']

        # 扫描件位置
        pdf_path = config['pdf_path']
        newNamePath = config['newNamePath']
        img_path = config['img_path']
        main_pdf = config['main_pdf']
    except:
        print('读取配置文件失败...')
