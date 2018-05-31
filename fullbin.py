#!/usr/bin/env python
# -*- coding: utf-8 -*-

" a csv handing program "

__author__ = "aning"

import os, datetime, csv,time
from PyQt4.QtCore import *
from PyQt4.QtGui import *
# from PyQt4.QtWidgets import *


class Mythread(QThread):
    #定义信号,定义参数为两个str type
    _signal = pyqtSignal(str, str)

    def __init__(self):
        super(Mythread, self).__init__()
    
    def run(self):
        #获取加载文件路径框里的值-路径
        Inpath = directoryComboBox.currentText()
        #获取输出文件路径框里的值
        Outpath = OutdirectoryComboBox.currentText()
        #获取备份文件路径框里的值
        Backpath = BackupdirectoryComboBox.currentText()
        self.currentDir = QDir(Inpath)
        nums = 0
        Infiles = [item for item in self.currentDir.entryList() if item.endswith(".csv") or item.endswith(".CSV")]
        if len(Infiles) > 0:
            for fn in Infiles:   
                self.fns = fn
                nums += 1
                #打开文件
                self.openFile(self.currentDir.absoluteFilePath(fn))
                #给新表一个表头
                self.newHead()
                #对第一数据部分开始操作
                self.headData()
                #对第二数据部分开始操作
                self.middleData()
                #对数据部分进行操作
                self.hanDing()
                #定义新文件存放路径与命名,取新表index01
                newListName = os.path.join(Outpath, self.newList[0][1]) + ".csv"
                #将新表写入新文件
                self.writes(newListName)
                #将旧文件备份,重命名并移动到Backup
                os.rename(self.currentDir.absoluteFilePath(fn), os.path.join(Backpath, self.nowTime(self.currentDir.absoluteFilePath(fn).split("/")[-1][:-4])))

                self._signal.emit(self.currentDir.absoluteFilePath(fn).split("/")[-1],"%d" % nums)
                time.sleep(0.1)
        else:
            filesTable.append("该文件夹下找不到符合条件的文件,请重新选择文件夹！")
    #open files
    def openFile(self, paths):
        with open(paths, "r", encoding = "utf-8") as oldFileObj:
            try:
                readers = csv.reader(oldFileObj)
                #*****POINTS****
                self.oldList = [row for row in readers]
            finally:
                oldFileObj.close()
    #获取旧表第一数据部分的第一行表头里的各个项目,以便填充第一数据部分
    def firstDataHead(self,indexs): 
        self.firstDataList = [item for item in self.oldList[indexs:indexs + 5]]
    #新表第一行的值以及新文件的命名
    def headFileValue(self):
        grade = sums = ""
        if len(self.oldList[6][1]) < 3:
            grade = "00" + self.oldList[6][1] if len(self.oldList[6][1]) < 2 else "0" + self.oldList[6][1]
        else:
            grade = self.oldList[6][1]
        #sums
        for item in self.oldList[2][1][2:].split("-"):
            if len(item) < 2:
                sums += "0"+item
            else:
                sums += item
        
        return self.oldList[0][0][12:] + sums + grade + self.currentDir.absoluteFilePath(self.fns).split("/")[-1].split("_")[0][-4:]
    #编辑新表头index3的值 && indexs
    def newHeadIndex3(self):
        nums = rows = 0
        for item in self.oldList[8:]:
            rows += 1
            if len(item) <= 0 or item[0] == "":
                break
            else:
                nums += int(item[1])
        self.firstDataHead(rows + 8)
        return nums
    
    #编辑新表的表头
    def newHead(self):
        try:
            self.newList =  [
            ["TADE Name:", self.headFileValue()],
            ["BinBlock No:", self.oldList[6][1]],
            ["BinGrade No:", self.oldList[6][1], self.oldList[6][1]],
            ["Total Qty:",  self.newHeadIndex3()],
            ["End Time:",  self.oldList[2][1] + " " +self.oldList[3][1][:-3]],
            ["Machine No:", self.oldList[0][0][12:]],
            ["Lot No:", self.oldList[4][2]],
            ["Bin Name:", self.oldList[6][1]],
            ["Barcode No:",self.currentDir.absoluteFilePath(self.fns).split("/")[-1].split("_")[0]],
            ["Sort Bin:", self.oldList[8][5]],
            [],
            ["", "Min", "Avg", "Max", "Std"]
            ]
        except IndexError:
            filesTable.append("文件表头错误,请重新选择文件！")

    # 第一数据部分
    def headData(self):
        dataList = ["CONTA", "CONTC", "POLAR", "VF1", "VF2", "VF3", "VF4", "VFMA1", "VFMC1", "DVF", "VF", "VFD", "VZ1", "VZ2", "IR1", "mcd1", "PO1", "lm1", "mcd2", "PO2", "lm2", "mcd3", "PO3", "lm3", "WLP1", "WLD1", "WLC1", "HW1", "PURITY1", "X1", "Y2", "Z1", "CCT1", "ST1", "INT1", "WLP2", "WLD2", "WLC2", "HW2", "PURITY2", "X2", "Y2", "Z2", "CCT2", "ST2", "INT2", "WLP3", "WLD3", "WLC3", "HW3", "PURITY3", "X3", "Y3", "Z3", "CCT3", "ST3", "INT3" ]
        emptyList = []
        #生成第一部分数据结构
        for item in dataList:
            emptyList.append([item, "0", "0", "0", "0"])
        for headItem in self.firstDataList[0]:
            for dataItem in dataList:
                tempValue = dataItem
                if dataItem == "IR1" or dataItem == "WLP1" or dataItem == "PO1":
                    tempValue = dataItem[:-1] if dataItem != "PO1" else "LOP1"
                
                if headItem == tempValue:
                    oldIndex = self.firstDataList[0].index(headItem)
                    newIndex = dataList.index(dataItem)
                    emptyList[newIndex][1] = self.firstDataList[1][oldIndex]
                    emptyList[newIndex][2] = self.firstDataList[3][oldIndex]
                    emptyList[newIndex][3] = self.firstDataList[2][oldIndex]
                    emptyList[newIndex][4] = self.firstDataList[4][oldIndex]

        for item in emptyList:
            self.newList.append(item)
        self.newList.append([])
    
    #第二数据部分
    def middleData(self):
        emptyList = []
        for item in self.oldList[8:]:
            if len(item) <= 0 or item[1] == "":
                break
            else:
                emptyList.append(["Customer No:", item[0]])
                emptyList.append(["Qty", item[1]])
        
        for item in emptyList:
            self.newList.append(item)
        self.newList.append([])
        self.newList.append(["index", "WaferID", "Bin Col", "Bin Row", "Grade", "Map Col", "Map Row", "CONTA", "CONTC", "POLAR", "VF1", "VF2", "VF3", "VF4", "VFMA1", "VFMC1", "DVF", "VF", "VFD", "VZ1", "VZ2", "IR1", "mcd1", "PO1", "lm1", "mcd2", "PO2", "lm2", "mcd3", "PO3", "lm3", "WLP1", "WLD1", "WLC1", "HW1", "PURITY1", "X1", "Y1", "Z1", "CCT1", "ST1", "INT1", "WLP2", "WLD2", "WLC2", "HW2", "PURITY2", "X2", "Y2", "Z2", "CCT2", "ST2", "INT2", "WLP3", "WLD3", "WLC3", "HW3", "PURITY3", "X3", "Y3", "Z3", "CCT3", "ST3", "INT3"])
    #对数据进行操作
    def hanDing(self,startIndex = 29):
        tempList = []
        indexs = 0
        inIndex = 0
        for x in range(len(self.oldList)):
            if len(self.oldList[x]) <= 0 or self.oldList[x][0] == "":
                pass
            elif self.oldList[x][0] == "WaferID" and x > 7:
                indexs = x + 1
                break

        for item in self.oldList[indexs:]:
            inIndex += 1
            tempList.append([inIndex, item[0], item[3], item[4], item[5], item[1], item[2], "", "", "", item[8], item[16], "", item[17], "", "", "", "", "", "", "", item[10], "", item[12], "", "", "", "", "", "", "", item[14], item[13], "", "", "", "", "", "", "", "", ""])
        
        for item in tempList:
            self.newList.append(item)

    #writes files
    def writes(self, paths):
        with open(paths, "w", newline="") as newFileObj:
            try:
                csvWriter = csv.writer(newFileObj, dialect='excel')
                csvWriter.writerows(self.newList)
            finally:
                newFileObj.close()
    
    #旧文件重命名
    def nowTime(self, oldName):
        return oldName + "_" + datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S').replace("-", "") + ".csv"


if __name__ == "__main__":
    import sys

    
    app = QApplication(sys.argv)
    window = QWidget()

    #函数区域
    tempAddBrowse = ""
    def browse():
        global tempAddBrowse
        if tempAddBrowse:
            directory = QFileDialog.getExistingDirectory(window, "Find Files",
                tempAddBrowse)
        else:
            directory = QFileDialog.getExistingDirectory(window, "Find Files",
                QDir.currentPath())
        tempAddBrowse = directory

        if directory:
            if directoryComboBox.findText(directory) == -1:
                directoryComboBox.addItem(directory)
            directoryComboBox.setCurrentIndex(directoryComboBox.findText(directory))

    tempAddOutBrowse = ""
    def outBrowse():
        global tempAddOutBrowse
        if tempAddOutBrowse:
            directory = QFileDialog.getExistingDirectory(window, "Find Files",
                tempAddOutBrowse)
        else:
            directory = QFileDialog.getExistingDirectory(window, "Find Files",
                QDir.currentPath())
        tempAddOutBrowse = directory

        if directory:
            if OutdirectoryComboBox.findText(directory) == -1:
                OutdirectoryComboBox.addItem(directory)
            OutdirectoryComboBox.setCurrentIndex(OutdirectoryComboBox.findText(directory))
            
    tempAddBackBrowse = ""
    def backupBrowse():
        global tempAddBackBrowse
        if tempAddBackBrowse:
            directory = QFileDialog.getExistingDirectory(window, "Find Files",
                tempAddBackBrowse)
        else:
            directory = QFileDialog.getExistingDirectory(window, "Find Files",
                QDir.currentPath())
        tempAddBackBrowse = directory

        if directory:
            if BackupdirectoryComboBox.findText(directory) == -1:
                BackupdirectoryComboBox.addItem(directory)
            BackupdirectoryComboBox.setCurrentIndex(BackupdirectoryComboBox.findText(directory))
    #fun end
    def createButton(text, member):
        button = QPushButton(text)
        button.clicked.connect(member)
        return button
    
    def createComboBox(text=""):
        comboBox = QComboBox()
        comboBox.setEditable(True)
        comboBox.addItem(text)
        comboBox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        return comboBox
    
    def reloads(fileName, nums):
        filesTable.append(fileName + "——> OK")
        filesCount.setText("%s documents have been completed" % nums)
    def run():

        thread._signal.connect(reloads)
        thread.start()

    def createFilesTable(): 
        global filesTable
        filesTable = QTextBrowser()
    #create floder title
    importfloder = QLabel("Input Folder:")
    Outfloder = QLabel("Output Folder:")
    Backupfloder = QLabel("Backup Folder:")

    #创建打开文件对话框button
    browseButton = createButton("&Input Browse", browse)
    browseButton.setStyleSheet("font-size: 14px;font-weight: 200;")
    outBrowseButton = createButton("&Output Browse", outBrowse)
    outBrowseButton.setStyleSheet("font-size: 14px;font-weight: 200;")
    backupBrowseButton = createButton("&Backup Browse", backupBrowse)
    backupBrowseButton.setStyleSheet("font-size: 14px;font-weight: 200;")

    #创建文件夹路径显示框
    directoryComboBox = createComboBox(QDir.currentPath())
    OutdirectoryComboBox = createComboBox()
    BackupdirectoryComboBox = createComboBox()

    #create status ok nums
    filesCount = QLabel()
    #create files views
    createFilesTable()


    #create running button
    RunButton = createButton("&Auto Run", run)
    
    RunButton.setStyleSheet("background-color: green; font-size: 20px; color: #ffffff;font-weight: 600;")
    RunButton.setFixedSize(100, 40)
    #create Close button
    closeButton = QPushButton("Close")
    closeButton.setStyleSheet("background-color: red; font-size: 20px; color: #ffffff; font-weight: 600; ")
    closeButton.setFixedSize(100, 40)
    #Close exe
    closeButton.clicked.connect(window.close)

    #create version 
    # versions = QLabel("Ver1.0.108")
    # versions.setAlignment(Qt.AlignBottom)
    # 将控件注册到窗口中
    mainLayout = QGridLayout()
    # 文件title
    mainLayout.addWidget(importfloder, 0, 0)
    mainLayout.addWidget(Outfloder, 2, 0)
    mainLayout.addWidget(Backupfloder, 4, 0)
    # button
    mainLayout.addWidget(browseButton, 1, 2)
    mainLayout.addWidget(outBrowseButton, 3, 2)
    mainLayout.addWidget(backupBrowseButton, 5, 2)
    # dirctorys
    mainLayout.addWidget(directoryComboBox, 1, 0)
    mainLayout.addWidget(OutdirectoryComboBox, 3, 0)
    mainLayout.addWidget(BackupdirectoryComboBox, 5, 0)
    # filesTable
    mainLayout.addWidget(filesTable, 6, 0, 1, 3)
    # filesCount
    mainLayout.addWidget(filesCount, 7, 0)
    #buttom buttons
    mainLayout.addWidget(RunButton, 8, 1)
    mainLayout.addWidget(closeButton, 8, 2)

    # mainLayout.addWidget(versions, 9, 0)

    window.setLayout(mainLayout)

    window.setWindowTitle("ASM MS899DL Full Bin")
    window.resize(700, 400)
    window.show()

    thread = Mythread()

    sys.exit(app.exec_())