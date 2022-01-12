#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import ssl
import threading
import time
import tkinter as tk

import cv2
import easyocr
import numpy as np
import win32api
import win32con
import win32gui
from PIL import ImageGrab

scale = 1.25

task_items_info = [
    # 物品名称，  商店页码, 对比图形路径
    ("茶花", 1, u"items/chahua.jpg"),
    ("茸", 1, u"items/lurong.jpg"),
    ("孔雀红", 2, u"items/kongquehong.jpg"),
    ("风", 2, u"items/feng.jpg"),
]


# 测试
# move_click(300, 300)

def resolution():  # 获取屏幕分辨率
    return win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)


# screen_resolution = resolution()

# 获取梦幻西游窗口信息吗，返回一个矩形窗口四个坐标
def get_window_info():
    global rect
    wdname = u'西游2'
    handle = 0
    hWndList = []
    win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)
    for hwnd in hWndList:
        title = win32gui.GetWindowText(hwnd)
        if (title.find(wdname) >= 0):
            rect = win32gui.GetWindowRect(hwnd)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, rect[2] - rect[0], rect[3] - rect[1],
                                  win32con.SWP_SHOWWINDOW)
            handle = hwnd
            break
    # handle = win32gui.FindWindow(0, wdname)  # 获取窗口句柄
    if handle == 0:
        return None, None, None
    else:
        return win32gui.GetWindowRect(handle), rect, handle


# window_size = get_window_info()
# 返回x相对坐标
def get_posx(x, window_size):
    return int((window_size[2] - window_size[0]) * x / 804)


# 返回y相对坐标
def get_posy(y, window_size):
    return int((window_size[3] - window_size[1]) * y / 630)


def stop():
    global is_start
    is_start = False
    print("停止")


class MyThread(threading.Thread):
    def __init__(self, func, *args):
        super().__init__()

        self.func = func
        self.args = args

        self.setDaemon(True)
        self.start()  # 在这里开始

    def run(self):
        self.func(*self.args)


# 青龙任务
def qinglongTask(window_szie):
    global is_start
    is_start = True
    topx, topy = window_size[0], window_size[1]
    img_qinglong_task = ImageGrab.grab((800, 205, 1000, 200))


def mouseLeftKeyClick(hwnd, position):
    time.sleep(0.5)
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, position)
    time.sleep(0.5)
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, position)


# 鼠标点击领取任务
def getQinglongTask(rect, hwnd):
    long_position = win32api.MAKELONG(int((rect[2] - rect[0]) / 2), int((rect[3] - rect[1]) / 2))
    # 选人
    mouseLeftKeyClick(hwnd, long_position)
    getTaskPost = win32api.MAKELONG(230, 345)
    # 领取任务
    mouseLeftKeyClick(hwnd, getTaskPost)
    # 关闭弹窗
    mouseLeftKeyClick(hwnd, getTaskPost)


# 抓取整个游戏屏幕,查找按钮位置
def catchGameWindowImgInitParams(rect):
    windowImg = ImageGrab.grab([int(rect[0] * scale), int(rect[1] * scale), int(rect[2] * scale), int(rect[3] * scale)])
    windowImg.save(u"game.jpg")
    time.sleep(3)
    # 上一个按钮
    p_x, p_y = analysisItemInfo(u"game.jpg", u"pics/previous.jpg")
    if p_x is None or p_y is None:
        return
    else:
        previous = [int((rect[0] + p_x) / scale) - 6, int((rect[1] + p_y) / scale) - 29]
        # win32api.SetCursorPos([int((rect[0] + p_x) / scale), int((rect[1] + p_y) / scale)])
    # 下一页按钮
    p_x, p_y = analysisItemInfo(u"game.jpg", u"pics/next.jpg")
    if p_x is None or p_y is None:
        return
    else:
        next = [int((rect[0] + p_x) / scale) - 6, int((rect[1] + p_y) / scale) - 29]
    # 购买按钮
    p_x, p_y = analysisItemInfo(u"game.jpg", u"pics/buyit.jpg")
    if p_x is None or p_y is None:
        return
    else:
        buy = [int((rect[0] + p_x) / scale) - 6, int((rect[1] + p_y) / scale) - 29]
    return previous, next, buy


# 重置购买页面
def resetItemsPage(rect, hwnd, previous):
    if previous is None:
        print(u"没有初始化成功")
    else:
        times = 5
        while times > 0:
            times -= 1
            mouseLeftKeyClick(hwnd, win32api.MAKELONG(previous[0], previous[1]))


# 查看当前任务是啥
def checkHasQinglongTask(rect, hwnd, previous, next, buy):
    # 125%显示比例问题
    taskBox = [806, 215, 988, 258]
    # taskBox = [645, 173, 789, 207]
    img = ImageGrab.grab(taskBox)
    # img.show()
    img.save(u't.jpg')
    result = reader.readtext(u't.jpg')
    if len(result) >= 2:
        # 获取是否有任务
        if result[0][1].find("青龙堂任务") == -1:
            print(u"没有任务，领取任务")
            # getQinglongTask(rect, hwnd)
        else:
            print(u"获取任务信息：")
            # 鼠标移动到该区域内
            contentBox = result[1][0]
            targetPos = [int(taskBox[0] / scale) + int((contentBox[2][0] - contentBox[0][0]) / 2)
                , int(taskBox[1] / scale) + int((contentBox[2][1] - contentBox[0][1]))]
            print(targetPos)
            # win32api.SetCursorPos(targetPos)
            taskInfo = getTaskNeedItem(result[1][1])
            if taskInfo is None:
                print("没有找到任务")
            else:
                buyTaskItem(rect, hwnd, taskInfo, previous, next, buy)
    print(result)


# 分析整个图片上面购买商品的位置，这个坐标当要使用到windows消息时候需要映射
def analysisItemInfo(baseImgPos, targetImgPos):
    baseImgPosTrueColor = cv2.imread(baseImgPos, 1)
    baseImgPosGray = cv2.imread(baseImgPos, 0)

    # 目标图片
    targetImg = cv2.imread(targetImgPos, 0)
    w, h = targetImg.shape[::-1]
    """
        TM_SQDIFF 平方差匹配法    该方法采用平方差来进行匹配；最好的匹配值为0；匹配越差，匹配值越大
        TM_CCORR 相关匹配法  该方法采用乘法操作；数值越大表明匹配程度越好。
        TM_CCOEFF 相关系数匹配法   1表示完美的匹配；-1表示最差的匹配。
        TM_SQDIFF_NORMED    归一化平方差匹配法　　　　　　
        TM_CCORR_NORMED 归一化相关匹配法　　　　　　
        TM_CCOEFF_NORMED    归一化相关系数匹配法
    """
    res = cv2.matchTemplate(baseImgPosGray, targetImg, cv2.TM_CCOEFF_NORMED)
    L = 0
    R = 1
    count = 0
    loopTimes: int = 20
    while --loopTimes > 0:
        threshold = (L + R) / 2
        count += 1
        loc = np.where(res >= threshold)
        if len(loc[0]) > 1:
            L += (R - L) / 2
        elif len(loc[0]) == 1:
            print(loc)
            pt = loc[::-1]
            print('目标区域的左上角坐标:', pt[0], pt[1])
            print('次数:', count)
            print('阀值', threshold)
            cv2.rectangle(baseImgPosTrueColor, (pt[0][0], pt[1][0]), (pt[0][0] + w, pt[1][0] + h), (34, 139, 139), 2)
            cv2.imwrite(u"result.jpg", baseImgPosTrueColor)
            return pt[0][0] + int(w / 2), pt[1][0] + int(h / 2)
        elif len(loc[0]) < 1:
            R -= (R - L) / 2
    # 得到指定的坐标位置
    return None, None


# 跳转购买页面
def jumpToPage(page, button):
    if previous is None:
        print(u"没有初始化成功")
    else:
        times = page
        while times > 0:
            times -= 1
            mouseLeftKeyClick(hwnd, win32api.MAKELONG(button[0], button[1]))


# 自动买东西
def buyTaskItem(rect, hwnd, itemInfo, previous, next, buyit):
    times = itemInfo[1]
    jumpToPage(times, next)
    # 重新抓图
    time.sleep(2)
    findItems = ImageGrab.grab([int(rect[0] * scale), int(rect[1] * scale), int(rect[2] * scale), int(rect[3] * scale)])
    findItems.save(u"tmp.jpg")
    # 找到物品
    px, py = analysisItemInfo(u"tmp.jpg", itemInfo[2])
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(px / scale) - 6, int(py / scale) - 29))
    # 购买
    time.sleep(2)
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(buyit[0], buyit[1]))
    # 重置页面
    time.sleep(2)
    resetItemsPage(rect, hwnd, previous)


# 获取任务物品
def getTaskNeedItem(info):
    time: int = 20
    while --time > 0:
        for itemInfo in task_items_info:
            if info.find(itemInfo[0]) >= 0:
                return itemInfo
            else:
                print(u"没有找到任务物品")
    return None


# 上交物品
def uploadItem(rect, hwnd, taskInfo):
    long_position = win32api.MAKELONG(int((rect[2] - rect[0]) / 2), int((rect[3] - rect[1]) / 2))
    # 选人
    mouseLeftKeyClick(hwnd, long_position)
    time.sleep(1)
    windowImg = ImageGrab.grab([int(rect[0] * scale), int(rect[1] * scale), int(rect[2] * scale), int(rect[3] * scale)])
    windowImg.save(u"upload.jpg")
    time.sleep(3)
    px, py = analysisItemInfo(u"upload.jpg", u"action/upload.jpg")
    # 点击上交
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(px / scale) - 6, int(py / scale) - 29))
    time.sleep(1.5)
    # 提交物品
    windowImg = ImageGrab.grab([int(rect[0] * scale), int(rect[1] * scale), int(rect[2] * scale), int(rect[3] * scale)])
    windowImg.save(u"uploadBox.jpg")
    time.sleep(3)
    px, py = analysisItemInfo(u"uploadBox.jpg", u"action/uploadbox.jpg")
    windowImg = ImageGrab.grab([px - 300, py, px + 450, py + 350])
    windowImg.save(u"currentItems.jpg")
    time.sleep(3)
    # 找物品
    px, py = analysisItemInfo(u"currentItems.jpg", taskInfo[2])
    if px is None or py is None:
        print(u"没有找到物品")
        return
    else:
        # 点击确认
        mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(px / scale) - 6, int(py / scale) - 29))
    time.sleep(1.5)
    # 上交物品
    px, py = analysisItemInfo(u"currentItems.jpg", u"action/ok.jpg")
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(px / scale) - 6, int(py / scale) - 29))
    time.sleep(1.5)
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(px / scale) - 6, int(py / scale) - 29))


# 启动
if __name__ == "__main__":
    ssl._create_default_https_context = ssl._create_unverified_context
    reader = easyocr.Reader(['ch_sim', 'en'], True)
    screen_resolution = resolution()
    # print(screen_resolution)
    window_size, rect, hwnd = get_window_info()
    print(window_size)
    # 点击人物
    global is_start
    # 创建主窗口
    root = tk.Tk()
    root.title("梦幻西游辅助")
    root.minsize(300, 250)
    root.maxsize(300, 250)
    # 初始化按钮位置
    previous, next, buy = catchGameWindowImgInitParams(window_size)
    resetItemsPage(rect, hwnd, previous)
    # 创建按钮
    qinglongButton = tk.Button(root, text=u"青龙任务,人物需要在青龙室",
                               command=lambda: MyThread(getQinglongTask, window_size, hwnd),
                               width=30, height=4)
    qinglongButton.place(relx=0.2, rely=0.15, width=200)
    qinglongButton.pack()

    # 测试截图
    qinglongClip = tk.Button(root, text=u"测试上交",
                             command=lambda: MyThread(uploadItem, window_size, hwnd, task_items_info[0]),
                             width=30, height=4)
    qinglongClip.place(relx=0.2, rely=0.55, width=200)
    qinglongClip.pack()

    # 测试cv2图像识别
    imgLear = tk.Button(root, text=u"测试",
                        command=lambda: MyThread(checkHasQinglongTask, window_size, hwnd, previous, next,
                                                 buy),
                        width=30, height=4)
    imgLear.place(relx=0.2, rely=0.85, width=200)
    imgLear.pack()

    root.mainloop()
