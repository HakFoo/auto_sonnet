# -*- coding:utf-8 -*-
import win32api
import win32con
import win32gui


def findIdSubHandle(pHandle, winClass, index=0):
    """
    已知子窗口类名
    查找第index个同类的兄弟窗口
    """

    assert type(index) == int and index >= 0
    """
    FindWindowEX(hwndParent=0,hwndChildAfter=0,lpszClass=None,lpszWindow=None)
    搜索类名和窗口名匹配的窗口，返回句柄
    hwndParent      不为0时查找句柄为hwndParent的子窗口
    hwndChildAfter  不为0时按z-indedx的顺序从hwndChildAfter向后搜索子窗口，否则从第一个开始
    lpClassName     窗口类名 
    lpWindow        窗口名，也就是标题栏
    """
    handle = win32gui.FindWindowEx(pHandle, 0, winClass, None)
    while index > 0:
        handle = win32gui.FindWindowEx(pHandle, handle, winClass, None)
        index -= 1
    return handle


def findSubHandle(pHandle, winClassList):
    """
    递归查找子窗口句柄
    pHandle         为父窗口句柄
    winClassList    各个子窗口class列表
    """
    assert type(winClassList) == list
    if len(winClassList) == 1:
        return findIdSubHandle(pHandle, winClassList[0][0],
                                winClassList[0][1])
    else:
        pHandle = findIdSubHandle(pHandle, winClassList[0][0],
                                   winClassList[0][1])
        return findSubHandle(pHandle, winClassList[1:])


def menuCommand(Mhandle, handle, command):
    """
    发送菜单命令
    """
    assert type(command) == str
    command_dict = {  # {目录编号，打开窗口名}  
        "Edit": [2, u"xgeom 11.54"],
        "Pointer": [0, None],
        "Reshape":[1,None],
        "Analyze":[1, None],
        "New Graph": [1, u"emgraph 11.54"],
        "Save Graph As...":[4, u"Enter filename for saved graph file:"],
        "Save As...":[5, u"Save As:"]

    }
    cmd_ID = win32gui.GetMenuItemID(handle, command_dict[command][0])
    win32gui.PostMessage(Mhandle, win32con.WM_COMMAND, cmd_ID, 0)
    if command != "Pointer" or "":
        index = 1
        # 当查找不到窗口的时候重复
        while win32gui.FindWindow(None, command_dict[command][1]) == 0:
            print(
                "Failed to find handle to the " + command +
                ".  Number of attempts: ", index)
            index += 1
            win32api.Sleep(500)
        """
        dig_handle      返回弹出的对话框的句柄
        confBTN_handle  返回确定按钮的句柄
        """
        dig_handle = win32gui.FindWindow(None, command_dict[command][1])
        confBTN_handle = win32gui.FindWindowEx(dig_handle, 0, "Button", None)
        return dig_handle, confBTN_handle


def findMenu(handle, index=0):
    """
    获得菜单栏的句柄
    当index=0时得到的句柄是一级菜单
    """
    assert type(index) == int and index >= 0
    return win32gui.GetSubMenu(handle, index)


def findMenuItem(handle, index):
    """
    获得某个菜单栏的内容
    """
    import win32gui_struct
    # 初始化空结构体
    mii, extra = win32gui_struct.EmptyMENUITEMINFO()
    # 获得菜单项信息
    win32gui.GetMenuItemInfo(handle, index, True, mii)
    ftype, fstate, wid, hsubmenu, hbmpchecked, hbmpunchecked, dwitemdata, text, hbmpitem = win32gui_struct.UnpackMENUITEMINFO(
        mii)
    return text


def setText(handle, text):
    """
    设置窗口文本
    """
    win32api.SendMessage(handle, win32con.WM_SETTEXT, 0, text)

def getEdit(hwnd):
    bufSize = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0) + 1
    strBuff = win32gui.PyMakeBuffer(bufSize)
    win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, bufSize, strBuff)
    add, leng = win32gui.PyGetBufferAddressAndLen(strBuff[:-1])

    return win32gui.PyGetString(add, leng)


def selfIncreasing(title: str) -> str:
    """
    将字符串末尾数字加1
    """
    return str(title[:9] + str(int(title[9:-4]) + 1) + title[-4:])


def getColor(win_dc, x, y):
    # 返回坐标位置的颜色RGB值
   
    return win32gui.GetPixel(win_dc, x, y)


def findColor(coordinate, win_dc, color):
    # 查找范围内的第一个满足本身颜色的像素坐标
    # coordinate_range = [xleft,ytop,xright,ybottom]
    # color = [R,G,B]
    # 返回坐标
    index = 1
    for x in range(coordinate[0], coordinate[2]):
        for y in range(coordinate[1], coordinate[3]):

            print("Find No.", index, "Pixel coordinates", " x_coordinates：", x, " y_coordinates：", y, "RGB：", hexToRGB(getColor(win_dc, x, y)))

            index += 1
            if hexToRGB(getColor(win_dc, x, y)) == color:
                return [x, y]


def hexToRGB(hex):
    return ((hex & 0xff), (hex >> 8) & 0xff, (hex >> 16) & 0xff)


def mousePos(Pos,wh):
    return (Pos * 65535) // wh


def mainRun():
    win32api.ShellExecute(0, 'open', 'C:/Program Files (x86)/sonnet.11.54/bin/sonnet.exe', '','',1)
    win32api.Sleep(1000)
    mainWin = 0
    mainWin = checkWin(mainWin,
                   winClass="xSonnet5100detdoc9012",
                   winName="Sonnet Task Bar 11.54")
    win32api.Sleep(1000)
    return mainWin


def checkMenu(hwnd, menuName, index, level=0):
    """
    保存文件
    """
    menu = win32gui.GetMenu(hwnd)
    subMenu = findMenu(menu, index[0])
    if level != 0:
        subMenu = findMenu(subMenu, index[1])
    win, BTN = menuCommand(hwnd, subMenu, menuName)

    return win, BTN

def checkWin(hwnd, winName=None, winClass=None):
    """
    查找窗口是否存在，如果不存在则重复查找
    查找到后返回hwnd
    """
    index = 0
    if winName != None:
        printStr = winName + " Not Found. Attempts: "
    else:
        printStr = winClass + " Not Found. Attempts: "


    while hwnd == 0:
        print(printStr, index)

        index += 1
        hwnd = win32gui.FindWindow(winClass, winName)
        win32api.Sleep(300)
    winName = win32gui.GetWindowText(hwnd)
    print("Found " + winName)
    return hwnd


def checkEdit(hwnd, Ehwnd, list):
    index = 0
    while Ehwnd == 0:
        print("Edit Not Find. Attempts: ", index)
        index += 1
        Ehwnd = findSubHandle(hwnd, list)
    return Ehwnd


def shortcutKey(first: int, second: int):
    """
    组合快捷键
    """
    win32api.keybd_event(first, 0, 0, 0)
    win32api.Sleep(100)

    win32api.keybd_event(second, 0, 0, 0)
    win32api.Sleep(100)

    win32api.keybd_event(first, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.Sleep(100)

    win32api.keybd_event(second, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.Sleep(100)



def pressMid(coorlist):
    """
    使用鼠标中键放大选区
    """
    win32api.SetCursorPos([(coorlist[0] + coorlist[2]) // 2,
                           (coorlist[1] + coorlist[3]) // 2])
    win32api.Sleep(100)

    win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
    win32api.Sleep(100)

    win32api.SetCursorPos([(coorlist[0] + coorlist[2]) // 4,
                           (coorlist[1] + coorlist[3]) // 4 * 3])
    win32api.Sleep(100)

    win32api.mouse_event(win32con.MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
    win32api.Sleep(100)


def returnRectCoor(coorlist, winDC):
    bList = [
        coorlist[2] // 13, coorlist[3] // 4, coorlist[2], (coorlist[3] // 4) + 1
    ]
    black = findColor(bList, winDC, (0, 0, 0))

    rect = [black[0], black[1] // 2, black[0] + 1, coorlist[3]]
    leftUp = findColor(rect, winDC, (176, 48, 96))

    rect = [leftUp[0], leftUp[1], coorlist[2], leftUp[1] + 1]
    rightUp = findColor(rect, winDC, (
        255,
        255,
        255,
    ))

    rect = [leftUp[0], leftUp[1], leftUp[0] + 1, coorlist[3]]
    leftDown = findColor(rect, winDC, (0, 0, 0))

    rect = [leftDown[0], leftDown[1], leftDown[0] + 1, coorlist[3]]
    leftDown = findColor(rect, winDC, (176, 48, 96))

    rect = [leftDown[0], leftDown[1], leftDown[0] + 1, coorlist[3]]
    leftDown = findColor(rect, winDC, (0, 0, 0))

    return [leftUp[0], leftUp[1], rightUp[0], leftDown[1]]


def reshape(hwnd):
    """
    xgeom窗口菜单栏 -> Tools -> Reshape
    """
    menu = win32gui.GetMenu(hwnd)
    submenu = win32gui.GetSubMenu(menu, 4)
    menuCommand(hwnd, submenu, "Reshape")

    win32api.Sleep(300)


def checkLeft(xPos, yPos):
    """
    在[xPos, yPos]点击鼠标左键
    """
    win32api.SetCursorPos([xPos, yPos])
    win32api.Sleep(100)

    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.Sleep(100)

    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    win32api.Sleep(500)


def moveMouse(index):
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.Sleep(200)
    win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 0, index, 0, 0)
    win32api.Sleep(500)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    win32api.Sleep(500)
