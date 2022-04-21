# -*- coding:utf-8 -*-
"""
Name    |   Class
--------------------------------
xgeom   |   xxgeom5100mditask9012
.son    |   xxgeom5100mdidoc9012


"""

import win32api
import win32con
import win32gui

from run import *


sonList = [["MDIClient", 0], ["xxgeom5100mdidoc9012", 0]]

saveList = [["DUIViewWndClassName", 0], ["DirectUIHWND", 0],
            ["FloatNotifySink", 0], ["ComboBox", 0], ["Edit", 0]]

sgrList = [["MDIClient", 0], ["xEmgraph5100mdidoc9012", 0]]

# 主循环

index = 1
while index <= 10:
    mainWin = mainRun()
    win32gui.PostMessage(mainWin, win32con.WM_CLOSE, 0, 0)
    win32api.Sleep(1000)

    mainWin = mainRun()
    win32api.Sleep(1000)

    mainWin = mainRun()

    checkMenu(mainWin, "Edit", [0, 2], level=1)
    xgeomWin = 0
    xgeomWin = checkWin(xgeomWin, winName="xgeom 11.54")

    # xgeom窗口最大化
    win32gui.ShowWindow(xgeomWin, win32con.SW_MAXIMIZE)
    """
    查找.son窗口句柄
    需要注意的是.son 和xgeom窗口中间还有个窗口，类型为 "MDIClient"
    """
    sonWin = findSubHandle(xgeomWin, sonList)
    sonWin = checkWin(sonWin, winClass="xxgeom5100mdidoc9012")
    # 得到.son文件title，并将数字增一，后面保存的时候会用到
    sonTitle = win32gui.GetWindowText(sonWin)
    # .son窗口最大化
    win32gui.ShowWindow(sonWin, win32con.SW_MAXIMIZE)

    # 使用快捷键 Shift+ESC将Toolbox选项调整到pointer
    shortcutKey(16, 27)
    """
    查找特定矩形区域内的窗口
    """

    # 得到窗口的矩形
    coordinate = list(win32gui.GetWindowRect(sonWin))
    coordinate[2], coordinate[0] = coordinate[0] + coordinate[2], 0

    # 鼠标中键放大选区
    pressMid(coordinate)

    sonDC = win32gui.GetDC(sonWin)

    # 矩形坐标列表 [left, top, right, bottom] （247，703） （1239，703）
    rectangleList = returnRectCoor(coordinate, sonDC)
    win32gui.ReleaseDC(sonWin, sonDC)

    reshape(xgeomWin)

    # 在鼠标坐标系统中，屏幕在水平和垂直方向上均匀分割成65535×65535个单元
    # 鼠标左键点击矩形左下角的点
    
    xPos = rectangleList[0]
    yPos = rectangleList[3] + 40
    checkLeft(xPos, yPos)

    # 按下ctrl键
    win32api.keybd_event(17, 0, 0, 0)
    win32api.Sleep(200)

    # 鼠标点击矩形右下角
    xPos = rectangleList[2]
    win32api.SetCursorPos([xPos, yPos])

    # 鼠标往上移动12像素
    moveMouse(-12)
    rectangleList[3] -= 12

    # 松开ctrl键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.Sleep(100)

    sonTitle = selfIncreasing(sonTitle)
    # xgeom窗口菜单栏 -> Save As ->

    saveWin, saveBTN = checkMenu(xgeomWin, "Save As...", [1])
    """
    定位到"文件名"窗体
    将保存的文件名设置为加一后的son_title
    """
    editWin = findSubHandle(saveWin, saveList)
    editWin = checkEdit(saveWin, editWin, saveList)

    while getEdit(editWin) != sonTitle:
        win32api.Sleep(500)
        print("Edit Not Changed.")
        setText(editWin, sonTitle)

    while win32gui.FindWindow(None,"Save As:"):
        print("Saving...")
        shortcutKey(18, 83)
        win32api.Sleep(1000)
    print("Saved " + sonTitle + " successfully.")
    """
    菜单栏Analysis -> Setup...
    """

    checkMenu(xgeomWin, "Analyze", [8])
    anaWin = 0
    anaWin = checkWin(anaWin, winClass="xemstatus5100detdoc9012")
    anaTitle = win32gui.GetWindowText(anaWin)
    while anaTitle.find("Finished") == -1:
        # 当未完成分析时，重复尝试
        win32api.Sleep(100)
        anaTitle = win32gui.GetWindowText(anaWin)
        print(anaTitle)

    # 关闭analyze_win窗口
    win32gui.PostMessage(anaWin, win32con.WM_CLOSE, 0, 0)
    print("Finished analyzing " + sonTitle)
    """
    将分析完成后的文件保存
    """
    emgWin, emgBTN = checkMenu(xgeomWin, "New Graph", [8, 3], level=1)
    emgWin = checkWin(emgWin,
                      winName="emgraph 11.54",
                      winClass="xEmgraph5100mditask9012")

    # emgraph窗口最大化
    win32gui.ShowWindow(emgWin, win32con.SW_MAXIMIZE)
    sgrWin = 0
    while sgrWin == 0:
        print("SGR not Found")
        sgrWin = findSubHandle(emgWin, sgrList)
        win32api.Sleep(100)

    win32gui.ShowWindow(sgrWin, win32con.SW_MAXIMIZE)

    # 保存文件
    saveWin, saveBTN = checkMenu(emgWin, "Save Graph As...", [1])
    sgrText = sonTitle[:-4] + ".sgr"
    # 定位到"文件名窗体"，将保存的文件名设置为sgr_text
    editWin = findSubHandle(saveWin, saveList)
    editWin = checkEdit(saveWin, editWin, saveList)

    while getText(editWin) != sgrText:
        print("Edit Not Changed.")
        setText(editWin, sgrText)


    while win32gui.FindWindow(None,"Enter filename for saved graph file:"):
        print("Saving...")
        shortcutKey(18, 83)
        win32api.Sleep(1000)
    print("Saved " + sgrText + " successfully")

    # 关闭窗口
    win32gui.PostMessage(emgWin, win32con.WM_CLOSE, 0, 0)
    win32gui.PostMessage(xgeomWin, win32con.WM_CLOSE, 0, 0)
    win32gui.PostMessage(mainWin, win32con.WM_CLOSE, 0, 0)

    win32api.Sleep(1000)

    index += 1