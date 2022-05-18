#   AutoExplain.py
#   coding: utf-8

import win32com
import win32com.client
from time import sleep
import traceback
from Src import ReadConfig
from Src import QuickResponseCode

def autoExplain():
    excel_path = ReadConfig.benchmarkPath + ReadConfig.benchmarkName
    flag = 0
    try:
        xlsm = win32com.client.Dispatch('Excel.Application')    # 引用excel
        xlsm.Visible = 1  # 文档可见
        excel = xlsm.Workbooks.Open(excel_path)     # excel表格打开
        list_sheet = excel.Worksheets(ReadConfig.benchmarkList)
        rowtotal = list_sheet.UsedRange.Rows.Count  # 清单子表的在用行数
        flag = 1
    except:
        print("打开excel文档失败...")
    # print(rowtotal)
    irow = 3
    if flag == 1:
        while irow <= rowtotal:
            # 根据超标类型组成不同文档路径
            tem_doc_path = ''
            if list_sheet.Range("E" + str(irow)).value == "环比":
                tem_doc_path = ReadConfig.docPath + ReadConfig.docNameH
            elif list_sheet.Range("E" + str(irow)).value == "同比":
                tem_doc_path = ReadConfig.docPath + ReadConfig.docNameT
            elif list_sheet.Range("E" + str(irow)).value == "同环比":
                tem_doc_path = ReadConfig.docPath + ReadConfig.docNameTH
            else:
                irow = irow + 1
                continue
            # print(tem_doc_path)
            # 打开work文档模并打开替换操作
            try:
                word = win32com.client.Dispatch('Word.Application')     # 引用word
                # word.Visible = 1  # 文档可见
                doc = word.Documents.Open(tem_doc_path)  # work文档打开
                word.Selection.Find.ClearFormatting()
                word.Selection.Find.Replacement.ClearFormatting()
            except:
                traceback.print_exc()
                print("打开work文档模板失败...")
                irow = irow + 1
                continue
            # 统一需要替换的位置
            try:
                word.Selection.Find.Execute("XX基站", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("D" + str(irow)).value, 2)    # 标题处替换
            except:
                pass
            try:
                word.Selection.Find.Execute("移动综资XX机房", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("D" + str(irow)).value, 2)  # 移动综资机房
            except:
                pass
            try:
                word.Selection.Find.Execute("JFZDPT-GD-XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("G" + str(irow)).value, 2)  # 电费缴费单编号
            except:
                pass
            try:
                word.Selection.Find.Execute("ZDBZD-GD-XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("F" + str(irow)).value, 2)  # 报账点编号
            except:
                pass
            try:
                word.Selection.Find.Execute("起始日期XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("I" + str(irow)).value, 2)  # 起始日期
            except:
                pass
            try:
                word.Selection.Find.Execute("截止日期XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("J" + str(irow)).value, 2)  # 截止日期
            except:
                pass
            try:
                word.Selection.Find.Execute("本期天数XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("K" + str(irow)).value, 2)  # 本期天数
            except:
                pass
            try:
                word.Selection.Find.Execute("本期电量XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("L" + str(irow)).value, 2)  # 本期电量
            except:
                pass
            try:
                word.Selection.Find.Execute("本期日均XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("M" + str(irow)).value, 2)  # 本期日均电量
            except:
                pass
            try:
                word.Selection.Find.Execute("本期功率XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("N" + str(irow)).value, 2)  # 本期功率
            except:
                pass
            try:
                word.Selection.Find.Execute("本期日均功率XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("O" + str(irow)).value, 2)  # 本期功率
            except:
                pass
            try:
                word.Selection.Find.Execute("超标原因XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AH" + str(irow)).value, 2)  # 超标原因
            except:
                pass
            try:
                word.Selection.Find.Execute("盖章日期", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AJ" + str(irow)).value, 2)  # 盖章日期
            except:
                pass
            # 同比需要替换的地方
            try:
                word.Selection.Find.Execute("同比期始XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("Z" + str(irow)).value, 2)  # 同比期始
                word.Selection.Find.Execute("同比期终XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AA" + str(irow)).value, 2)  # 同比期终
                word.Selection.Find.Execute("同比天数XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AB" + str(irow)).value, 2)  # 同比天数
                word.Selection.Find.Execute("同比电量XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AC" + str(irow)).value, 2)  # 同比电量
                word.Selection.Find.Execute("同比日均XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AD" + str(irow)).value, 2)  # 同比日均
                word.Selection.Find.Execute("同比功率XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AE" + str(irow)).value, 2)  # 同比功率
                word.Selection.Find.Execute("同比日均功率XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AF" + str(irow)).value, 2)  # 同比日均功率
                word.Selection.Find.Execute("同比A/BXX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("AG" + str(irow)).value, 2)  # A/B
                tem_str = "同比超标率" + str(round(list_sheet.Range("Y" + str(irow)).value*100, 2)) + "%"
                word.Selection.Find.Execute("同比超标率XX%", False, False, False, False, False, True, 1, False,
                                            tem_str, 2)  # 同比超标率
            except:
                pass
            # 环比需要替换的地方
            try:
                word.Selection.Find.Execute("环比期始XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("Q" + str(irow)).value, 2)  # 环比期始
                word.Selection.Find.Execute("环比期终XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("R" + str(irow)).value, 2)  # 环比期终
                word.Selection.Find.Execute("环比天数XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("S" + str(irow)).value, 2)  # 环比天数
                word.Selection.Find.Execute("环比电量XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("T" + str(irow)).value, 2)  # 环比电量
                word.Selection.Find.Execute("环比日均XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("U" + str(irow)).value, 2)  # 环比日均
                word.Selection.Find.Execute("环比功率XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("V" + str(irow)).value, 2)  # 环比功率
                word.Selection.Find.Execute("环比日均功率XX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("W" + str(irow)).value, 2)  # 环比日均功率
                word.Selection.Find.Execute("环比A/BXX", False, False, False, False, False, True, 1, False,
                                            list_sheet.Range("X" + str(irow)).value, 2)  # 环比A/B
                tem_str = "环比超标率" + str(round(list_sheet.Range("P" + str(irow)).value*100, 2)) + "%"
                word.Selection.Find.Execute("环比超标率XX%", False, False, False, False, False, True, 1, False,
                                            tem_str, 2)  # 环比超标率
            except:
                pass
            # 生成文件名称二维码
            try:
                # 二维码内容
                code_data = ReadConfig.StName + list_sheet.Range("D" + str(irow)).value + "超标说明（" + \
                            list_sheet.Range("G" + str(irow)).value + "）.pdf"
                QuickResponseCode.quickResponseCode(code_data)  # 生成二维码
                code_path = ReadConfig.savePath + code_data + ".jpg"
                # doc.Paragraphs.Range.InlineShapes.AddPicture(code_path)
                parag = doc.Paragraphs.Add()    # 在打开的文档中增加段落
                parag_range = parag.Range   # 定位段落区域
                parag_range.InlineShapes.AddPicture(code_path)  # 插入图片
            except:
                traceback.print_exc()
                pass
            # 报告保存
            try:
                # 组成保存路径
                tem_path = ReadConfig.savePath + ReadConfig.StName + list_sheet.Range("D" + str(irow)).value + "超标说明（" + \
                            list_sheet.Range("G" + str(irow)).value + "）.docx"
                doc.SaveAs(tem_path)    # 文档另存为
                print('已生成超标报告：'+tem_path)
                doc.Close(SaveChanges=False)    # 模板不保存关闭
                word.Application.Quit()     # 退出引用
                sleep(0.5)
            except:
                traceback.print_exc()
                pass
            irow = irow + 1
    print("生成超标报告完成...")
    # 关闭excel文档，并退出引用
    try:
        excel.Close(SaveChanges=False)
        xlsm.Application.Quit()
    except:
        pass
