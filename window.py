# coding: utf8

from Tkinter import *
import tkMessageBox
from tkFileDialog import *
from gridStrategy import MakeGrid

class Window:

    def __init__(self):

        self.max = 1.2
        self.initalStep = 2
        self.rate = 0.15
        self.initialBuyNum = 1000
        self.large = 0.3
        self.middle = 0.15
        self.small = 0.05

        window = Tk()
        window.title("网格策略 - wuxi")
        window.geometry('800x600')

        x1 = 10
        x2 = 120
        y = 10
        Label(window, text="目标价格").place(x=x1, y=y)
        self.targetEn = Entry(window)
        self.targetEn.insert(0, str(self.max))
        self.targetEn.place(x=x2, y=y)
        y += 40

        Label(window, text="递增起点(0开始)").place(x=x1, y=y)
        self.startEn = Entry(window)
        self.startEn.place(x=x2, y=y)
        self.startEn.insert(0, str(self.initalStep))
        y += 40

        Label(window, text="递增比率(小数)").place(x=x1, y=y)
        self.stepEn = Entry(window)
        self.stepEn.insert(0, str(self.rate))
        self.stepEn.place(x=x2, y=y)
        y += 40

        Label(window, text="单笔购买数").place(x=x1, y=y)
        self.singleBuyEn = Entry(window)
        self.singleBuyEn.place(x=x2, y=y)
        self.singleBuyEn.insert(0, str(self.initialBuyNum))
        y += 40

        Label(window, text="大网百分比(小数)").place(x=x1, y=y)
        self.largeEn = Entry(window)
        self.largeEn.place(x=x2, y=y)
        self.largeEn.insert(0, str(self.large))
        y += 40

        Label(window, text="中网百分比(小数)").place(x=x1, y=y)
        self.middleEn = Entry(window)
        self.middleEn.place(x=x2, y=y)
        self.middleEn.insert(0, str(self.middle))
        y += 40

        Label(window, text="小网百分比(小数)").place(x=x1, y=y)
        self.smallEn = Entry(window)
        self.smallEn.place(x=x2, y=y)
        self.smallEn.insert(0, str(self.small))
        y += 40

        Button(window, text="请选择输出文件夹", command=self.selectOutputPath).place(x=x1, y=y)
        self.dirName = StringVar()
        self.dirName.set('/Users/wuxi/Desktop')
        self.dirLabel = Label(window, textvariable=self.dirName)
        self.dirLabel.place(x=x1 + 150, y=y)
        y += 40

        Button(window, text="开始", command=self.start).place(x=x1, y=y)

        window.mainloop()
        self.window = window


    # 选择输出文件路径
    def selectOutputPath(self):
        path = askdirectory()
        print '路径： ', path
        self.dirName.set(path)

    # 开始生成表格
    def start(self):
        self.max = float(self.targetEn.get())
        self.initalStep = int(self.startEn.get())
        self.initalStep = float(self.stepEn.get())
        self.initialBuyNum = int(self.singleBuyEn.get())
        self.large = float(self.largeEn.get())
        self.middle = float(self.middleEn.get())
        self.small = float(self.smallEn.get())

        dir = self.dirName.get()
        build = MakeGrid(dir=dir)
        build.max = self.max
        build.min = self.max*0.4
        build.initalStep = self.initalStep
        build.stepRate = self.rate
        build.large = self.large
        build.middle = self.middle
        build.small = self.small
        build.minBuyNum = self.initialBuyNum
        build.makeGrids()
        build.generateExcel()
        tkMessageBox.showinfo(title="提示", message="成功！")




win = Window()