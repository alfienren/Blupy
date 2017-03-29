# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'mainwindow.ui'
#
# Created: Thu Feb 23 15:39:14 2017
#      by: pyside-uic 0.2.15 running on PySide 1.2.4
#
# WARNING! All changes made in this file will be lost!

from PySide import QtCore, QtGui

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(600, 558)
        self.centralWidget = QtGui.QWidget(MainWindow)
        self.centralWidget.setObjectName("centralWidget")
        self.tabWidget = QtGui.QTabWidget(self.centralWidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 601, 531))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtGui.QWidget()
        self.tab.setObjectName("tab")
        self.sheet_path_label = QtGui.QLabel(self.tab)
        self.sheet_path_label.setGeometry(QtCore.QRect(30, 80, 281, 21))
        self.sheet_path_label.setObjectName("sheet_path_label")
        self.chromedriver_path = QtGui.QLabel(self.tab)
        self.chromedriver_path.setGeometry(QtCore.QRect(30, 250, 281, 21))
        self.chromedriver_path.setObjectName("chromedriver_path")
        self.crawl_urls_button = QtGui.QPushButton(self.tab)
        self.crawl_urls_button.setGeometry(QtCore.QRect(200, 420, 171, 41))
        self.crawl_urls_button.setObjectName("crawl_urls_button")
        self.label = QtGui.QLabel(self.tab)
        self.label.setGeometry(QtCore.QRect(10, 0, 531, 81))
        self.label.setObjectName("label")
        self.sheet_path = QtGui.QLineEdit(self.tab)
        self.sheet_path.setGeometry(QtCore.QRect(30, 180, 501, 31))
        self.sheet_path.setObjectName("sheet_path")
        self.chromedrive_path_sublabel = QtGui.QLabel(self.tab)
        self.chromedrive_path_sublabel.setGeometry(QtCore.QRect(30, 280, 501, 61))
        self.chromedrive_path_sublabel.setObjectName("chromedrive_path_sublabel")
        self.sheet_path_sublabel = QtGui.QLabel(self.tab)
        self.sheet_path_sublabel.setGeometry(QtCore.QRect(30, 110, 501, 71))
        self.sheet_path_sublabel.setObjectName("sheet_path_sublabel")
        self.chromedriver_path_box = QtGui.QLineEdit(self.tab)
        self.chromedriver_path_box.setGeometry(QtCore.QRect(30, 340, 501, 31))
        self.chromedriver_path_box.setObjectName("chromedriver_path_box")
        self.tabWidget.addTab(self.tab, "")
        MainWindow.setCentralWidget(self.centralWidget)
        self.menuBar = QtGui.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 600, 21))
        self.menuBar.setObjectName("menuBar")
        MainWindow.setMenuBar(self.menuBar)
        self.mainToolBar = QtGui.QToolBar(MainWindow)
        self.mainToolBar.setObjectName("mainToolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        self.statusBar = QtGui.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        MainWindow.setStatusBar(self.statusBar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QtGui.QApplication.translate("MainWindow", "MainWindow", None, QtGui.QApplication.UnicodeUTF8))
        self.sheet_path_label.setText(QtGui.QApplication.translate("MainWindow", "<html><head/><body><p><span style=\" font-size:12pt;\">Enter path of sheet containing URLs:</span></p></body></html>", None, QtGui.QApplication.UnicodeUTF8))
        self.chromedriver_path.setText(QtGui.QApplication.translate("MainWindow", "<html><head/><body><p><span style=\" font-size:12pt;\">Enter path of chromedriver:</span></p></body></html>", None, QtGui.QApplication.UnicodeUTF8))
        self.crawl_urls_button.setText(QtGui.QApplication.translate("MainWindow", "Crawl URLs", None, QtGui.QApplication.UnicodeUTF8))
        self.label.setText(QtGui.QApplication.translate("MainWindow", "<html><head/><body><p><span style=\" font-size:24pt;\">Scrape URLs for DCM Floodlight Tags</span></p></body></html>", None, QtGui.QApplication.UnicodeUTF8))
        self.chromedrive_path_sublabel.setText(QtGui.QApplication.translate("MainWindow", "<html><head/><body><p><span style=\" font-size:10pt;\">Enter path of chromedriver.exe. The .exe should be located on your local drive (not the shared drives).</span></p></body></html>", None, QtGui.QApplication.UnicodeUTF8))
        self.sheet_path_sublabel.setText(QtGui.QApplication.translate("MainWindow", "<html><head/><body><p><span style=\" font-size:10pt;\">Enter the path to the sheet containing the URLs to scrape. The sheet must contain a header and the URLs to scrape should only be in column A.</span></p></body></html>", None, QtGui.QApplication.UnicodeUTF8))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), QtGui.QApplication.translate("MainWindow", "Tab 1", None, QtGui.QApplication.UnicodeUTF8))

