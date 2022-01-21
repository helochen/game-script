#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import ssl
import threading
import time
import tkinter as tk
from tkinter import messagebox
from tkinter import Entry

import cv2
import easyocr
import numpy as np
import win32api
import win32con
import win32gui
from PIL import ImageGrab

scale = 1.25
rect = None
hwnd = None
windows_x_extra = None
windows_y_extra = None
previous = None
next = None
buy = None
upload_button_pos = None
find_items = None
reader = None
running = False
upload_ok = None

task_items_info = [
    # 物品名称，  商店页码, 对比图形路径
    ("臭", 0, u"items/chou.jpg"),
    ("真豆腐", 0, u"items/chou.jpg"),
    ("烤鸭", 0, u"items/kao.jpg"),
    ("肉", 0, u"items/rou.jpg"),
    ("佛", 0, u"items/fo.jpg"),
    ("跳", 0, u"items/fo.jpg"),
    ("翡", 0, u"items/feicui.jpg"),
    ("翠", 0, u"items/feicui.jpg"),
    ("桂", 0, u"items/wan.jpg"),
    # ("花", 0, u"items/wan.jpg"),
    ("丸", 0, u"items/wan.jpg"),
    ("豆", 0, u"items/dou.jpg"),
    ("果", 0, u"items/dou.jpg"),
    ("长", 0, u"items/changshou.jpg"),
    ("寿", 0, u"items/changshou.jpg"),
    ("面", 0, u"items/changshou.jpg"),
    ("面", 0, u"items/changshou.jpg"),
    ("珍", 0, u"items/zhenglu.jpg"),
    ("梅花", 0, u"items/meihua.jpg"),
    ("梅", 0, u"items/meihua.jpg"),
    ("蛇胆", 0, u"items/she.jpg"),
    ("蛇", 0, u"items/she.jpg"),
    ("胆酒", 0, u"items/she.jpg"),
    ("虎骨酒", 0, u"items/hugu.jpg"),
    ("虎骨", 0, u"items/hugu.jpg"),
    ("虎", 0, u"items/hugu.jpg"),
    ("骨", 0, u"items/hugu.jpg"),
    ("骨", 0, u"items/hugu.jpg"),
    ("百味酒", 0, u"items/baiwei.jpg"),
    ("百味", 0, u"items/baiwei.jpg"),
    ("味", 0, u"items/baiwei.jpg"),
    ("生梦死", 0, u"items/zuishengmengsi.jpg"),
    ("梦死", 0, u"items/zuishengmengsi.jpg"),
    ("梦", 0, u"items/zuishengmengsi.jpg"),
    ("死", 0, u"items/zuishengmengsi.jpg"),
    ("女儿", 0, u"items/nver.jpg"),
    ("女", 0, u"items/nver.jpg"),

    ("老", 1, u"items/lao.jpg"),
    ("石", 1, u"items/shi.jpg"),
    ("右英", 1, u"items/shi.jpg"),
    ("英", 1, u"items/shi.jpg"),
    ("茸", 1, u"items/lurong.jpg"),
    ("茶花", 1, u"items/chahua.jpg"),
    ("六", 1, u"items/liudao.jpg"),
    ("胆", 1, u"items/dan.jpg"),
    ("蕻胆", 1, u"items/dan.jpg"),
    ("熊胆", 1, u"items/dan.jpg"),
    ("熊", 1, u"items/dan.jpg"),
    ("尾", 1, u"items/wei.jpg"),
    ("草", 1, u"items/cao.jpg"),
    ("心", 1, u"items/xin.jpg"),
    ("火", 1, u"items/fengzhi.jpg"),
    ("灭", 1, u"items/fengzhi.jpg"),
    ("香水", 1, u"items/xiangshui.jpg"),
    ("月", 1, u"items/yue.jpg"),
    ("菁", 1, u"items/yue.jpg"),
    ("爻", 1, u"items/yue.jpg"),

    ("风", 2, u"items/feng.jpg"),
    ("饮露", 2, u"items/feng.jpg"),
    ("白露为霜", 2, u"items/bai.jpg"),
    ("白露", 2, u"items/bai.jpg"),
    ("为霜", 2, u"items/bai.jpg"),
    ("地", 2, u"items/di.jpg"),
    ("灵芝", 2, u"items/di.jpg"),
    ("天龙水", 2, u"items/tian.jpg"),
    ("天龙", 2, u"items/tian.jpg"),
    ("龙水", 2, u"items/tian.jpg"),
    ("天充水", 2, u"items/tian.jpg"),
    ("天无水", 2, u"items/tian.jpg"),
    ("天充", 2, u"items/tian.jpg"),
    ("充水", 2, u"items/tian.jpg"),
    ("孔雀红", 2, u"items/kongquehong.jpg"),
    ("香", 2, u"items/xiang.jpg"),
    ("血", 2, u"items/shanhu.jpg"),
    ("血", 2, u"items/shanhu.jpg"),
    ("仙", 2, u"items/xian.jpg"),
    ("狐涎", 2, u"items/xian.jpg"),
    ("狐", 2, u"items/xian.jpg"),
    ("涎", 2, u"items/xian.jpg"),
]


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
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, rect[2] - rect[0], rect[3] - rect[1],
                                  win32con.SWP_SHOWWINDOW)
            handle = hwnd
            break
    if handle == 0:
        return None, None
    else:
        return win32gui.GetWindowRect(handle), handle


class MyThread(threading.Thread):
    def __init__(self, func, *args):
        super().__init__()

        self.func = func
        self.args = args

        self.setDaemon(True)
        self.start()  # 在这里开始

    def run(self):
        self.func(*self.args)


# 鼠标左键点击
def mouseLeftKeyClick(hwnd, position):
    time.sleep(1.0)
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, position)
    time.sleep(0.5)
    win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, position)


def mouseRightKeyClick(hwnd, position):
    time.sleep(0.5)
    win32api.SendMessage(hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, position)
    time.sleep(0.5)
    win32api.SendMessage(hwnd, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, position)


# 鼠标点击领取任务
def getQinglongTask(rect, hwnd):
    long_position = win32api.MAKELONG(int((rect[2] - rect[0]) / 2), int((rect[3] - rect[1]) / 2))
    # 选人
    mouseLeftKeyClick(hwnd, long_position)
    # 需要改坐标信息
    getTaskPost = win32api.MAKELONG(int((244 - windows_x_extra) / scale), int((466 - windows_y_extra) / scale))
    # 领取任务
    mouseLeftKeyClick(hwnd, getTaskPost)
    mouseLeftKeyClick(hwnd, getTaskPost)
    # 关闭弹窗
    mouseLeftKeyClick(hwnd, getTaskPost)


# 抓取整个游戏屏幕,查找按钮位置
def catchGameWindowImgInitParams(rect):
    global windows_x_extra
    global windows_y_extra
    windowImg = ImageGrab.grab([int(rect[0] * scale), int(rect[1] * scale), int(rect[2] * scale), int(rect[3] * scale)])
    windowImg.save(u"game.jpg")
    time.sleep(3)
    # 上一个按钮
    p_x, p_y = analysisItemInfo(u"game.jpg", u"pics/previous.jpg")
    if p_x is None or p_y is None:
        return
    else:
        previous = [int((rect[0] + p_x) / scale) - windows_x_extra, int((rect[1] + p_y) / scale) - windows_y_extra]
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
        times = 3
        while times > 0:
            times -= 1
            mouseLeftKeyClick(hwnd, win32api.MAKELONG(previous[0], previous[1]))


# 查看当前任务是啥
def checkHasQinglongTask(rect, hwnd, previous, next, buy):
    # 125%显示比例问题
    taskBox = [int(640 * scale), int(170 * scale), int(790 * scale), int(206 * scale)]
    # taskBox = [645, 173, 789, 207]
    img = ImageGrab.grab(taskBox)
    # img.show()
    img.save(u't.jpg')
    global reader
    result = reader.readtext(u't.jpg')
    print("读取结果：")
    print(result)
    if len(result) >= 2:
        # 获取是否有任务
        if result[0][1].find("青") == -1 and result[0][1].find("堂") == -1 and result[0][1].find("吉才学仟务") == -1 and \
                result[0][1].find("龙裳") == -1:
            print(u"没有任务，去领取任务")
        else:
            # print(u"获取任务信息：")
            # 鼠标移动到该区域内
            # contentBox = result[1][0]
            # targetPos = [int(taskBox[0] / scale) + int((contentBox[2][0] - contentBox[0][0]) / 2)
            #    , int(taskBox[1] / scale) + int((contentBox[2][1] - contentBox[0][1]))]
            taskInfo = getTaskNeedItem(result[1][1])
            if taskInfo is None:
                print("没有找到任务所需物品")
            else:
                buyTaskItem(rect, hwnd, taskInfo, previous, next, buy)
    else:
        print("完全没有认出来...")


# 分析整个图片上面购买商品的位置，这个坐标当要使用到windows消息时候需要映射
def analysisItemInfo(baseImgPos, targetImgPos):
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
            # print(loc)
            pt = loc[::-1]
            # print('目标区域的左上角坐标:', pt[0], pt[1])
            # print('次数:', count)
            # print('阀值', threshold)
            # cv2.rectangle(baseImgPosTrueColor, (pt[0][0], pt[1][0]), (pt[0][0] + w, pt[1][0] + h), (34, 139, 139), 2)
            # cv2.imwrite(u"result.jpg", baseImgPosTrueColor)
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
    time.sleep(0.5)
    resetItemsPage(rect, hwnd, previous)
    time.sleep(0.5)
    # 上交物品
    uploadItem(rect, hwnd, itemInfo)


# 获取任务物品
def getTaskNeedItem(info):
    print("需要的物品名称：" + info)
    time: int = 20
    while time > 0:
        time -= 1
        for itemInfo in task_items_info:
            if info.find(itemInfo[0]) >= 0:
                return itemInfo
    return None


# 上交物品
def uploadItem(rect, hwnd, taskInfo):
    long_position = win32api.MAKELONG(int((rect[2] - rect[0]) / 2), int((rect[3] - rect[1]) / 2))
    # 选人
    mouseLeftKeyClick(hwnd, long_position)
    mouseLeftKeyClick(hwnd, long_position)
    time.sleep(1)
    global upload_button_pos
    if upload_button_pos is None:
        windowImg = ImageGrab.grab(
            [int(rect[0] * scale), int(rect[1] * scale), int(rect[2] * scale), int(rect[3] * scale)])
        windowImg.save(u"upload.jpg")
        time.sleep(3)
        upload_button_pos = analysisItemInfo(u"upload.jpg", u"action/upload.jpg")
    # 点击上交
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(upload_button_pos[0] / scale) - windows_x_extra,
                                              int(upload_button_pos[1] / scale) - windows_y_extra))
    time.sleep(1.5)
    # 提交物品
    windowImg = ImageGrab.grab([int(rect[0] * scale), int(rect[1] * scale), int(rect[2] * scale), int(rect[3] * scale)])
    windowImg.save(u"uploadBox.jpg")
    time.sleep(3)
    # 坐标运算
    global find_items
    if find_items is None:
        find_items = analysisItemInfo(u"uploadBox.jpg", u"action/uploadbox.jpg")
    windowImg = ImageGrab.grab([find_items[0] - 50, find_items[1], find_items[0] + 350, find_items[1] + 350])
    windowImg.save(u"currentItems.jpg")
    time.sleep(3)
    # 找物品
    px, py = analysisItemInfo(u"currentItems.jpg", taskInfo[2])
    if px is None or py is None:
        print(u"没有找到物品")
        return
    else:
        # 点击确认
        mouseLeftKeyClick(hwnd,
                          win32api.MAKELONG(int((find_items[0] + px - 50) / scale - windows_x_extra),
                                            int((find_items[1] + py) / scale - windows_y_extra / 2)))
    time.sleep(1.5)
    # 确定给予 物品
    global upload_ok
    if upload_ok is None:
        print(u"初始化上交按钮位置")
        upload_ok = analysisItemInfo(u"uploadBox.jpg", u"action/ok.jpg")
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(upload_ok[0] / scale) - windows_x_extra,
                                              int(upload_ok[1] / scale) - windows_y_extra))
    time.sleep(0.2)
    mouseLeftKeyClick(hwnd, win32api.MAKELONG(int(upload_ok[0] / scale) - windows_x_extra,
                                              int(upload_ok[1] / scale) - windows_y_extra))


# 做一次青龙任务
def finishOneQinglongTask(rect, hwnd, previous, next, buy):
    global running
    if not running:
        running = True
        loop = 500
        while loop > 0:
            loop -= 1
            print("第%d次任务" % loop)
            getQinglongTask(rect, hwnd)
            checkHasQinglongTask(rect, hwnd, previous, next, buy)
        print("安全结束。。。")
        running = False
    else:
        messagebox.showwarning(u"正在运行中")


# 初始化按钮位置
def initSysParams(sc):
    global rect
    global hwnd
    global windows_x_extra
    global windows_y_extra
    global previous
    global next
    global buy
    global upload_button_pos
    global find_items
    global scale
    scale = float(sc)

    rect, hwnd = get_window_info()
    windows_x_extra = rect[2] - 800
    windows_y_extra = rect[3] - 600
    # 初始化按钮位置
    previous, next, buy = catchGameWindowImgInitParams(rect)
    resetItemsPage(rect, hwnd, previous)
    messagebox.showinfo(u"初始化完成")


# 启动
if __name__ == "__main__":
    ssl._create_default_https_context = ssl._create_unverified_context
    reader = easyocr.Reader(['ch_sim', 'en'], True)

    # 创建主窗口
    root = tk.Tk()
    root.title("梦幻西游辅助")
    root.minsize(350, 400)
    root.maxsize(350, 400)

    # 系统放大率
    lb = tk.Label(root, text=u"系统放大率:", width=10)
    lb.grid(row=0, column=0)
    scaleText = Entry(root)
    scaleText.insert(0, scale)
    scaleText.grid(row=0, column=1)

    # 创建按钮
    qinglongButton = tk.Button(root, text=u"第1步初始化：青龙任务,人物需要在青龙室，",
                               command=lambda: MyThread(initSysParams, scaleText.get()),
                               width=35, height=4)
    qinglongButton.place(relx=0.2, rely=0.15, width=250)
    qinglongButton.grid(row=1, column=1)

    # 创建按钮
    getQinglongButton = tk.Button(root, text=u"第2步初始化：领取青龙任务",
                                  command=lambda: MyThread(getQinglongTask, rect, hwnd),
                                  width=35, height=4)
    getQinglongButton.place(relx=0.2, rely=0.15, width=250)
    getQinglongButton.grid(row=2, column=1)
    # 测试截图
    qinglongClip = tk.Button(root, text=u"第3步测试：测试上交当前任务",
                             command=lambda: MyThread(checkHasQinglongTask, rect, hwnd, previous, next,
                                                      buy),
                             width=35, height=4)
    qinglongClip.place(relx=0.2, rely=0.15, width=250)
    qinglongClip.grid(row=3, column=1)

    # 测试cv2图像识别
    imgLear = tk.Button(root, text=u"开跑：五十十次青龙任务",
                        command=lambda: MyThread(finishOneQinglongTask, rect, hwnd, previous, next,
                                                 buy),
                        width=35, height=4)
    imgLear.place(relx=0.2, rely=0.15, width=250)
    imgLear.grid(row=4, column=1)

    root.mainloop()
