# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.tab)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.listWidget = QtWidgets.QListWidget(self.tab)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout.addWidget(self.listWidget)
        self.pushButton = QtWidgets.QPushButton(self.tab)
        self.pushButton.setObjectName("pushButton")
        self.verticalLayout.addWidget(self.pushButton)
        self.horizontalLayout_2.addLayout(self.verticalLayout)
        self.toolBox = QtWidgets.QToolBox(self.tab)
        self.toolBox.setObjectName("toolBox")
        self.page = QtWidgets.QWidget()
        self.page.setGeometry(QtCore.QRect(0, 0, 370, 415))
        self.page.setObjectName("page")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.page)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.label = QtWidgets.QLabel(self.page)
        font = QtGui.QFont()
        font.setFamily("楷体")
        font.setPointSize(48)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignTop)
        self.label.setObjectName("label")
        self.verticalLayout_5.addWidget(self.label)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_5.addItem(spacerItem)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.page)
        font = QtGui.QFont()
        font.setFamily("AcadEref")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_2.addWidget(self.label_2)
        spacerItem1 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem1)
        self.spinBox = QtWidgets.QSpinBox(self.page)
        self.spinBox.setObjectName("spinBox")
        self.verticalLayout_2.addWidget(self.spinBox)
        self.horizontalLayout_4.addLayout(self.verticalLayout_2)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_3 = QtWidgets.QLabel(self.page)
        font = QtGui.QFont()
        font.setFamily("AcadEref")
        font.setPointSize(14)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.verticalLayout_3.addWidget(self.label_3)
        spacerItem2 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_3.addItem(spacerItem2)
        self.spinBox_2 = QtWidgets.QSpinBox(self.page)
        self.spinBox_2.setObjectName("spinBox_2")
        self.verticalLayout_3.addWidget(self.spinBox_2)
        self.horizontalLayout_4.addLayout(self.verticalLayout_3)
        self.verticalLayout_4.addLayout(self.horizontalLayout_4)
        self.pushButton_2 = QtWidgets.QPushButton(self.page)
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout_4.addWidget(self.pushButton_2)
        self.verticalLayout_5.addLayout(self.verticalLayout_4)
        self.horizontalLayout_5.addLayout(self.verticalLayout_5)
        self.toolBox.addItem(self.page, "")
        self.page_2 = QtWidgets.QWidget()
        self.page_2.setGeometry(QtCore.QRect(0, 0, 370, 415))
        self.page_2.setObjectName("page_2")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout(self.page_2)
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.verticalLayout_7 = QtWidgets.QVBoxLayout()
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        spacerItem3 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_7.addItem(spacerItem3)
        self.verticalLayout_6 = QtWidgets.QVBoxLayout()
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_4 = QtWidgets.QLabel(self.page_2)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_6.addWidget(self.label_4)
        self.spinBox_3 = QtWidgets.QSpinBox(self.page_2)
        self.spinBox_3.setObjectName("spinBox_3")
        self.horizontalLayout_6.addWidget(self.spinBox_3)
        self.verticalLayout_6.addLayout(self.horizontalLayout_6)
        self.pushButton_3 = QtWidgets.QPushButton(self.page_2)
        self.pushButton_3.setObjectName("pushButton_3")
        self.verticalLayout_6.addWidget(self.pushButton_3)
        self.verticalLayout_7.addLayout(self.verticalLayout_6)
        spacerItem4 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_7.addItem(spacerItem4)
        self.verticalLayout_8.addLayout(self.verticalLayout_7)
        self.toolBox.addItem(self.page_2, "")
        self.horizontalLayout_2.addWidget(self.toolBox)
        self.horizontalLayout_3.addLayout(self.horizontalLayout_2)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.tab_2)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.verticalLayout_9 = QtWidgets.QVBoxLayout()
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.listWidget_2 = QtWidgets.QListWidget(self.tab_2)
        self.listWidget_2.setObjectName("listWidget_2")
        self.verticalLayout_9.addWidget(self.listWidget_2)
        self.pushButton_4 = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_4.setObjectName("pushButton_4")
        self.verticalLayout_9.addWidget(self.pushButton_4)
        self.pushButton_5 = QtWidgets.QPushButton(self.tab_2)
        self.pushButton_5.setObjectName("pushButton_5")
        self.verticalLayout_9.addWidget(self.pushButton_5)
        self.horizontalLayout_7.addLayout(self.verticalLayout_9)
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_4 = QtWidgets.QWidget()
        self.tab_4.setObjectName("tab_4")
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout(self.tab_4)
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.verticalLayout_10 = QtWidgets.QVBoxLayout()
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.listWidget_3 = QtWidgets.QListWidget(self.tab_4)
        self.listWidget_3.setObjectName("listWidget_3")
        self.verticalLayout_10.addWidget(self.listWidget_3)
        self.pushButton_6 = QtWidgets.QPushButton(self.tab_4)
        self.pushButton_6.setObjectName("pushButton_6")
        self.verticalLayout_10.addWidget(self.pushButton_6)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.label_5 = QtWidgets.QLabel(self.tab_4)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_8.addWidget(self.label_5)
        spacerItem5 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_8.addItem(spacerItem5)
        self.lineEdit = QtWidgets.QLineEdit(self.tab_4)
        self.lineEdit.setMaximumSize(QtCore.QSize(100, 16777215))
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_8.addWidget(self.lineEdit)
        spacerItem6 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_8.addItem(spacerItem6)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.tab_4)
        self.lineEdit_2.setMaximumSize(QtCore.QSize(100, 16777215))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_8.addWidget(self.lineEdit_2)
        self.verticalLayout_10.addLayout(self.horizontalLayout_8)
        self.pushButton_7 = QtWidgets.QPushButton(self.tab_4)
        self.pushButton_7.setObjectName("pushButton_7")
        self.verticalLayout_10.addWidget(self.pushButton_7)
        self.horizontalLayout_9.addLayout(self.verticalLayout_10)
        self.tabWidget.addTab(self.tab_4, "")
        self.tab_5 = QtWidgets.QWidget()
        self.tab_5.setObjectName("tab_5")
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout(self.tab_5)
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.verticalLayout_11 = QtWidgets.QVBoxLayout()
        self.verticalLayout_11.setObjectName("verticalLayout_11")
        self.listWidget_4 = QtWidgets.QListWidget(self.tab_5)
        self.listWidget_4.setObjectName("listWidget_4")
        self.verticalLayout_11.addWidget(self.listWidget_4)
        self.pushButton_8 = QtWidgets.QPushButton(self.tab_5)
        self.pushButton_8.setObjectName("pushButton_8")
        self.verticalLayout_11.addWidget(self.pushButton_8)
        self.pushButton_9 = QtWidgets.QPushButton(self.tab_5)
        self.pushButton_9.setObjectName("pushButton_9")
        self.verticalLayout_11.addWidget(self.pushButton_9)
        self.horizontalLayout_10.addLayout(self.verticalLayout_11)
        self.tabWidget.addTab(self.tab_5, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout(self.tab_3)
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.verticalLayout_12 = QtWidgets.QVBoxLayout()
        self.verticalLayout_12.setObjectName("verticalLayout_12")
        self.listWidget_5 = QtWidgets.QListWidget(self.tab_3)
        self.listWidget_5.setObjectName("listWidget_5")
        self.verticalLayout_12.addWidget(self.listWidget_5)
        self.pushButton_10 = QtWidgets.QPushButton(self.tab_3)
        self.pushButton_10.setObjectName("pushButton_10")
        self.verticalLayout_12.addWidget(self.pushButton_10)
        self.pushButton_11 = QtWidgets.QPushButton(self.tab_3)
        self.pushButton_11.setObjectName("pushButton_11")
        self.verticalLayout_12.addWidget(self.pushButton_11)
        self.horizontalLayout_11.addLayout(self.verticalLayout_12)
        self.tabWidget.addTab(self.tab_3, "")
        self.tab_6 = QtWidgets.QWidget()
        self.tab_6.setObjectName("tab_6")
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout(self.tab_6)
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.textBrowser = QtWidgets.QTextBrowser(self.tab_6)
        self.textBrowser.setObjectName("textBrowser")
        self.horizontalLayout_12.addWidget(self.textBrowser)
        self.tabWidget.addTab(self.tab_6, "")
        self.horizontalLayout.addWidget(self.tabWidget)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(5)
        self.toolBox.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "PDF工具"))
        self.pushButton.setText(_translate("MainWindow", "添加文件"))
        self.label.setText(_translate("MainWindow", "范围"))
        self.label_2.setText(_translate("MainWindow", "开始"))
        self.label_3.setText(_translate("MainWindow", "结束"))
        self.pushButton_2.setText(_translate("MainWindow", "开始"))
        self.toolBox.setItemText(self.toolBox.indexOf(self.page), _translate("MainWindow", "自定义范围"))
        self.label_4.setText(_translate("MainWindow", "页面拆分份数"))
        self.pushButton_3.setText(_translate("MainWindow", "开始拆分"))
        self.toolBox.setItemText(self.toolBox.indexOf(self.page_2), _translate("MainWindow", "固定范围"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "PDF拆分"))
        self.pushButton_4.setText(_translate("MainWindow", "添加文件"))
        self.pushButton_5.setText(_translate("MainWindow", "合并文件"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "PDF合并"))
        self.pushButton_6.setText(_translate("MainWindow", "添加文件"))
        self.label_5.setText(_translate("MainWindow", "输入密码位数(默认 最小2位 最大6位)："))
        self.pushButton_7.setText(_translate("MainWindow", "开始破解"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_4), _translate("MainWindow", "PDF暴力破解"))
        self.pushButton_8.setText(_translate("MainWindow", "添加文件"))
        self.pushButton_9.setText(_translate("MainWindow", "开始转化"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_5), _translate("MainWindow", "PDF转Word(纯文本）"))
        self.pushButton_10.setText(_translate("MainWindow", "添加文件"))
        self.pushButton_11.setText(_translate("MainWindow", "开始OCR转化"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "PDF转Word(OCR)"))
        self.textBrowser.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:18pt;\">作者：整洁的浣熊</span></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:18pt;\"><br /></p>\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:18pt;\">Github：https://github.com/zhihao2020/</span></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:18pt;\"><br /></p>\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:18pt;\">请勿用于非法途径</span></p></body></html>"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_6), _translate("MainWindow", "关于"))

