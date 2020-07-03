import DataHandler as dh
import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication,QMainWindow
import sys
from Utils import logger
from Ui_MainWindow import Ui_MainWindow

class MainWindow(QtWidgets.QMainWindow,Ui_MainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()
        self.setupUi(self)

        self.resPath = None
        self.outputPath = None
        self.keysStr = None

    def selectResFile(self):
        self.resPath = QtWidgets.QFileDialog.getExistingDirectory(self, "选择要抓取的源文件夹")
        if len(self.resPath) > 0:
            mainWindow.lineEdit_resPath.setText(self.resPath)
        else:
            mainWindow.lineEdit_resPath.clear()

    def selectOutputPath(self):
        self.outputPath = QtWidgets.QFileDialog.getExistingDirectory(self, "选择输出文件目录")
        if len(self.outputPath) > 0:
            mainWindow.lineEdit_outPutPath.setText(self.outputPath)
        else:
            mainWindow.lineEdit_outPutPath.clear()

    def keysTextChanged(self):
        self.keysStr = mainWindow.textEdit_keys.toPlainText()

    def checkInputValid(self):
        if not (self.outputPath and len(self.outputPath) > 0):
            msgBox = QtWidgets.QMessageBox(self)
            msgBox.about(self, "提示", "请选择输出文件的目录")
            return False

        if not (self.keysStr and len(self.keysStr) > 0):
            msgBox = QtWidgets.QMessageBox(self)
            msgBox.about(self, "提示", "请填写要抓取的关键字列表，每个关键字用#号隔开")
            return False

        return True

    def goToOutput(self):
        if(self.checkInputValid()):
            logger.info("开始抓取数据 ...")
            dataHandler = dh.DataHandler()
            dataHandler.ReadFileFolder(self.resPath)

            dataHandler.SearchContentFromFiles(self.keysStr)
            dataHandler.BuildDataFrame(self.outputPath)
            logger.info("生成数据结束 ...")

if __name__=="__main__":
    app=QtWidgets.QApplication(sys.argv)
    mainWindow=MainWindow()
    mainWindow.show()

    #选择源文件夹按钮
    mainWindow.pushButton_selectResFolder.clicked.connect(mainWindow.selectResFile)

    #选择目标文件夹按钮
    mainWindow.pushButton_selectOutput.clicked.connect(mainWindow.selectOutputPath)

    #填写需要生成的关键字列表
    mainWindow.textEdit_keys.textChanged.connect(mainWindow.keysTextChanged)

    # 执行按钮
    mainWindow.go.clicked.connect(mainWindow.goToOutput)
    sys.exit(app.exec_())