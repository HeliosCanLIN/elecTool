from Src import AutoExplain
from Src import Login
from Src import ReadConfig
from Src import NewGetBenchmark
from Src import QuickResponseCode
import warnings
import traceback

# 忽略warnings
warnings.filterwarnings('ignore')


def main():
    ReadConfig.readconfig()
    #Login.login()

    while 1:
        print("(0) - 启动三费系统")
        print("(1) - 登录三费系统")
        print("(2) - 抓取超标清单")
        print("(3) - 生成超标报告")
        # print("(4) - 生成文件名称二维码")
        print("(4) - 读取二维码重命名")
        print("exit退出")
        userinput = input("输入数字选择功能\n>>>")
        if userinput == "exit":
            try:
                all_handle = Login.browser.window_handles
                for handle in all_handle:
                    Login.browser.switch_to.window(handle)
                    Login.browser.close()
            except:
                pass
            break
        if userinput == "0":
            try:
                Login.login()
            except:
                traceback.print_exc()
                pass
        if userinput == "1":
            try:
                Login.login()
                Login.enterPassword()
            except:
                traceback.print_exc()
                pass
        if userinput == "2":
            try:
                NewGetBenchmark.getBenchmark()
            except:
                pass
        if userinput == "3":
            try:
                AutoExplain.autoExplain()
            except:
                pass
        if userinput == "4":
            try:
                # QuickResponseCode.quickResponseCode("SM-001-湛江徐闻黄定一楼机房无线1超标说明（JFZDPT-GD-20201213-01412）")
                QuickResponseCode.decodeDisplay()
            except:
                traceback.print_exc()
                pass

if __name__ == '__main__':
    main()
