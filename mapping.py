#!/usr/bin/env python
# -*- coding: utf-8 -*-

" a csv handing program "

__author__ = "aning"

import os, datetime, csv,time
from PyQt4.QtCore import *
from PyQt4.QtGui import *
# from PyQt5.QtWidgets import *


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
        currentDir = QDir(Inpath)
        nums = 0
        Infiles = [item for item in currentDir.entryList() if item.endswith(".csv") or item.endswith(".CSV")]
        if len(Infiles) > 0:
            for fn in Infiles:
                nums += 1
                #打开文件
                self.openFile(currentDir.absoluteFilePath(fn))
                #给新表一个表头
                self.newHead()
                #对数据部分进行操作
                self.hanDing()
                #定义新文件存放路径与命名,取新表index72
                newListName = os.path.join(Outpath, self.newList[7][2])
                #将新表写入新文件
                self.writes(newListName)
                #将旧文件备份,重命名并移动到Backup
                os.rename(currentDir.absoluteFilePath(fn), os.path.join(Backpath, self.oldNowName(currentDir.absoluteFilePath(fn).split("/")[-1][:-4])))

                self._signal.emit(currentDir.absoluteFilePath(fn).split("/")[-1],"%d" % nums)
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
    
    #编辑新表的表头
    def newHead(self):
        try:
            self.newList =  [
            ["DataFileName","", self.oldList[2][2] + "_" + self.oldList[0][2] + ".csv", "[W_AOI:1]"],
            ["LotNumber", "", self.oldList[2][2]],
            ["DeviceNumber", "", self.oldList[2][2]],
            [self.oldList[1][0] + "=" + self.oldList[2][2]],
            ["EPI_ID", "" , self.oldList[0][2]],
            ["Resorting_Bin", "", self.oldList[13][2]],
            ["TestTime", "",self.oldList[3][2], ""],
            ["MapFileName", "", self.oldList[2][2] + "_" + self.oldList[0][2] + ".csv"],
            ["TransferTime", "", self.oldList[4][2]],
            [],
            ["map data"]
            ]
        except IndexError:
            filesTable.append("文件表头错误,请重新选择文件！")

    #对数据进行操作
    def hanDing(self,startIndex = 29):
        for i in range(len(self.oldList)):
            if len(self.oldList[i]) <= 0 or self.oldList[i][0] == "":
                pass
            elif self.oldList[i][0] == "TEST":
                startIndex = i + 1
                break

        for x in range(startIndex, len(self.oldList)):
            tempList = [self.oldList[x][2],self.oldList[x][3],self.oldList[x][1],self.oldList[x][0],self.oldList[x][32],self.oldList[x][7],self.oldList[x][16], self.oldList[x][18], self.oldList[x][21], self.oldList[x][20], self.oldList[x][29], self.oldList[x][28], self.oldList[x][31],self.oldList[x][8],self.oldList[x][10]]
            self.newList.append(tempList)
        self.newList.insert(11, [self.newList[11][0], self.newList[11][1]])
    #writes files
    def writes(self, paths):
        with open(paths, "w", newline="") as newFileObj:
            try:
                csvWriter = csv.writer(newFileObj, dialect='excel')
                csvWriter.writerows(self.newList)
            finally:
                newFileObj.close()
    
    #旧文件重命名
    def oldNowName(self, oldName):
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

    window.setWindowTitle("ASM MS899DL  Mapping")
    window.resize(700, 400)
    window.show()

    thread = Mythread()

    sys.exit(app.exec_())