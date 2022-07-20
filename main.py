# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import csv
import time

import pyautogui
import requests
import win32api
import win32com.client
import win32con
import win32gui
from PyQt5.QtCore import pyqtSignal, QThread
from PyQt5.QtWidgets import QApplication, QWidget, QMessageBox
from PyQt5.QtGui import *
import win32gui
import sys
from ui import Ui_Form
from PyQt5.QtWidgets import QApplication, QMainWindow
from paddleocr import PaddleOCR,draw_ocr
from PIL import Image
hwnd_title = dict()
hwnd= 0
win = 0
title = ["标题","x1坐标","y1坐标","x2坐标","y2坐标"]
config_path = 'config.csv'
def jietu():
    screen = QApplication.primaryScreen()
    img = screen.grabWindow(hwnd,476,102,1224,843).toImage()
    img.save("screenshot.jpg")

def jiexitupian():
    ocr = PaddleOCR(use_angle_cls=True, use_gpu=False)  # 使用CPU预加载，不用GPU
    result = ocr.ocr("screenshot.jpg",cls=True)  # 打开图片文件
    # # 打印所有文本信息
    # for t in text:
    #     print(t)

    image = Image.open("screenshot.jpg").convert('RGB')
    boxes = [line[0] for line in result]
    txts = [line[1][0] for line in result]
    scores = [line[1][1] for line in result]
    im_show = draw_ocr(image, boxes, txts, scores)
    im_show = Image.fromarray(im_show)
    im_show.show()
    print("1111")
def stoptech():
    time.sleep(1)
    pyautogui.moveTo(1752, 280, 0.25);
    x, y = pyautogui.position();
    print(x, y)
    time.sleep(2)
    pyautogui.click()
    time.sleep(1)
    # pyautogui.click(1600, 180);
    time.sleep(1);
    pyautogui.scroll(-100)
    time.sleep(0.5)
    pyautogui.click(1600,900)
    # pyautogui.click(1080, 67); 工业开启点击

def mixwindow():
    Win.showMinimized()

def loadData1(filename):
    with open(filename,"r",encoding='gb18030') as csvfile:
        reader1 = csv.reader(csvfile)
        column1 = [row[0] for row in reader1]
        return column1

def loadData2(filename,i): #i表示列
    with open(filename,"r",encoding='gb18030') as csvfile:
        reader2 = csv.reader(csvfile)
        next(reader2)
        column2 = [row[i] for row in reader2]
        return column2


class MyMainForm(QMainWindow, Ui_Form):
    def __init__(self, parent=None):
        super(MyMainForm, self).__init__(parent)
        self.setupUi(self)
        self.pushButton_4.clicked.connect(self.machine_display)
        self.pushButton_3.clicked.connect(self.tupianparse)
        self.pushButton_2.clicked.connect(self.stoptechnology)
    def stoptechnology(self):
        time.sleep(0.3)
        mixwindow()
        time.sleep(3)
        pyautogui.click(560, 171);
        stoptech()
    def machine_display(self):
        time.sleep(0.1)
        hwnd = win32gui.FindWindow("WindowsForms10.Window.8.app.0.378734a", None)
        time.sleep(0.1)
        print(hwnd, "1111")
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        isfore = win32gui.SetForegroundWindow(hwnd)
        print(isfore)
        time.sleep(1)
        pyautogui.click(560,171);
        time.sleep(0.2)
        # pyautogui.moveTo(1752,280,0.25);
        # x,y=pyautogui.position();
        # print(x,y)
        # time.sleep(2)
        # pyautogui.click(1600,180);
        # time.sleep(1);
        # pyautogui.click(1080,67);
        pyautogui.moveTo(300, 300, 0.25);
        pyautogui.click();
        time.sleep(0.5)
        pyautogui.scroll(-100)
        time.sleep(0.2)
        pyautogui.moveTo(1053, 943, 0.4);
        time.sleep(4)
        pyautogui.doubleClick(1577, 863,0.1)
        time.sleep(4)
        pyautogui.scroll(100)
        time.sleep(1)
        jietu()
        time.sleep(0.5)
        dividepic()
        time.sleep(0.2)
        pyautogui.click(560, 171);
        # Messagebox()

    def tupianparse(self):
        ocr = PaddleOCR(use_angle_cls=True, use_gpu=False);
        textlist = [];
        print("start")
        x1 = loadData2(config_path, 1)
        tx1 = loadData2(config_path, 0)
        for i in range(len(x1)):
            path = "zhijing_{}.jpg".format(i);
            print(path)
            result = ocr.ocr(path, cls=True)
            print("result",len(result))
            if len(result)==0:
                textlist.append("无法识别")
            else:
                for t in result:
                    print(t[1][0])
                    textlist.append(t[1][0])
        for i in range(len(x1)):
            print("序号->",i,tx1[i],"---")
            if i ==0:
                self.lineEdit.setText(textlist[i])
                self.label.setText(tx1[i])
            if i ==1:
                self.lineEdit_2.setText(textlist[i])
                self.label_2.setText(tx1[i])
            if i ==2:
                self.lineEdit_3.setText(textlist[i])
                self.label_3.setText(tx1[i])
            if i ==3:
                self.lineEdit_4.setText(textlist[i])
                self.label_4.setText(tx1[i])
            if i == 4:
                self.lineEdit_5.setText(textlist[i])
                self.label_5.setText(tx1[i])
            if i == 5:
                self.lineEdit_6.setText(textlist[i])
                self.label_6.setText(tx1[i])
            if i == 6:
                self.lineEdit_7.setText(textlist[i])
                self.label_7.setText(tx1[i])
            if i == 7:
                self.lineEdit_8.setText(textlist[i])
                self.label_8.setText(tx1[i])
            if i == 8:
                self.lineEdit_9.setText(textlist[i])
                self.label_9.setText(tx1[i])
            if i == 9:
                self.lineEdit_10.setText(textlist[i])
                self.label_10.setText(tx1[i])
            if i == 10:
                self.lineEdit_11.setText(textlist[i])
                self.label_11.setText(tx1[i])
            if i == 11:
                self.lineEdit_12.setText(textlist[i])
                self.label_12.setText(tx1[i])
            if i == 12:
                self.lineEdit_13.setText(textlist[i])
                self.label_13.setText(tx1[i])
            if i == 13:
                self.lineEdit_14.setText(textlist[i])
                self.label_14.setText(tx1[i])
            if i == 14:
                self.lineEdit_15.setText(textlist[i])
                self.label_15.setText(tx1[i])
            if i == 15:
                self.lineEdit_16.setText(textlist[i])
                self.label_16.setText(tx1[i])
            if i == 16:
                self.label_17.setText(tx1[i])
                self.lineEdit_17.setText(textlist[i])
            if i == 17:
                self.label_18.setText(tx1[i])
                self.lineEdit_18.setText(textlist[i])
            if i == 18:
                self.label_19.setText(tx1[i])
                self.lineEdit_19.setText(textlist[i])
            if i == 19:
                self.label_20.setText(tx1[i])
                self.lineEdit_20.setText(textlist[i])
            if i == 20:
                self.label_21.setText(tx1[i])
                self.lineEdit_21.setText(textlist[i])
            if i == 21:
                self.label_22.setText(tx1[i])
                self.lineEdit_22.setText(textlist[i])
            if i == 22:
                self.label_23.setText(tx1[i])
                self.lineEdit_23.setText(textlist[i])
            if i == 23:
                self.label_24.setText(tx1[i])
                self.lineEdit_24.setText(textlist[i])
            if i == 24:
                self.label_25.setText(tx1[i])
                self.lineEdit_25.setText(textlist[i])
            if i == 25:
                self.label_26.setText(tx1[i])
                self.lineEdit_26.setText(textlist[i])
        textlist.clear()



    def settext(self):
        print("12211111111111111111111111111111111")
        self.lineEdit.setText("112")

def Messagebox():
    box = QMessageBox(QMessageBox.Question, '信息', '确认完成')
    box.exec_()

def get_all_hwnd(hwnd, mouse):
    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

def set_current_window(hwnd):
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    else:
        win32gui.SetForegroundWindow(hwnd)


class Mythread(QThread):
    breakSignal = pyqtSignal(int)
    def __init__(self, parent=None):
        super().__init__(parent)
    def run(self):
        for i in range(1,10):
            print(i)
            self.breakSignal.emit(i)

def dividepic():
    im = Image.open("screenshot.jpg")
    w, h = im.size
    print(w, h)
    x1 = loadData2(config_path,1)
    y1 = loadData2(config_path,2)
    x2 = loadData2(config_path,3)
    y2 = loadData2(config_path,4)
    print(x1,y1,x2,y2)
    xy = [];
    if len(x1) == len(y1) and len(x2)==len(y2):
        for num in range(len(x1)):
            box = int(x1[num]),int(y1[num]),int(x2[num]),int(y2[num])
            print(box)
            im1 = im.crop(box)
            path = "zhijing_{}.jpg".format(num).replace(",","");
            im1.save(path, quality=100)
            im1.close()
    im.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    Win = MyMainForm()
    Win.show()
    time.sleep(2)
    sys.exit(app.exec_())

    win32gui.EnumWindows(get_all_hwnd, 0)
    for h, t in hwnd_title.items():
        if t != "":
            print(h, t)
    time.sleep(5)
    hwnd = win32gui.FindWindow("WindowsForms10.Window.8.app.0.378734a", "远程显示 - Floorganizer")
    time.sleep(0.1)

    print(hwnd,"1111")
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    isfore = win32gui.SetForegroundWindow(hwnd)
    print(isfore)
    time.sleep(1)



