#!/usr/bin/python
# -*-coding:utf-8-*-

import win32ui
import xlwings as xw

dat = "Data Files (*.dat)|*.dat||"
excel = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx||"


def openfile(filetype):
    # 1表示打开文件对话框
    dlg = win32ui.CreateFileDialog(1, None, None, 1, filetype)
    dlg.SetOFNTitle("请打开" + filetype + "文件")
    # 设置打开文件对话框中的初始显示目录
    dlg.SetOFNInitialDir('C:/')
    flag = dlg.DoModal()
    filename = dlg.GetPathName()
    # 获取选择的文件名称(已选择文件前提下)
    if flag == 1:
        return filename
    else:
        return openfile(filetype)


def readfile(filename):
    global datalist
    datalist = {}
    file = open(filename)
    for line in file:
        content = line.strip().split("|")
        datalist[content[1]] = eval(content[2])
    return datalist


def workbook(filename, datalist):
    # 打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    workbook = app.books.open(filename)
    sheet = workbook.sheets[0]
    for row in range(1, 80):
        for column in range(1, 10):
            cell = sheet.range((row, column))
            if cell.value:
                if type(cell.value) == int or type(cell.value) == float:
                    celldata = str(int(cell.value)).strip()
                else:
                    celldata = str(cell.value).strip()
                if celldata in datalist:
                    cell.offset(1, 0).value = datalist[celldata]
    workbook.save()
    app.quit()
    return


# if __name__ == '_main_':

file1 = openfile(dat)
datalist1 = readfile(file1)
file2 = openfile(dat)
datalist2 = readfile(file2)
datalists = dict(datalist1, **datalist2)
bookname = openfile(excel)
workbook(bookname, datalists)
