import sys
import os

import pandas as pd

from PySide.QtGui import QMainWindow, QApplication, QTableWidget, QKeySequence

from mainwindow import Ui_MainWindow
from http.urls import URLs


class App(QMainWindow, Ui_MainWindow, URLs):
    def __init__(self):
        super(App, self).__init__()
        self.setupUi(self)
        self.crawl_urls_button.clicked.connect(self.crawl_urls)
        self.sheet_path_sublabel.setWordWrap(True)
        self.chromedrive_path_sublabel.setWordWrap(True)

    def crawl_urls(self):
        url_list = str(self.sheet_path.text())
        url_list = url_list.encode('string-escape')

        driver_path = str(self.chromedriver_path_box.text())
        driver_path = driver_path.encode('string-escape')

        if url_list[-4:] == '.csv':
            urls = pd.read_csv(url_list, index_col=None).ix[:, 0].tolist()
        else:
            urls = pd.read_excel(url_list, index_col=None).ix[:, 0].tolist()

        floodlights = URLs().list_floodlights_from_urls(urls, driver_path)

        save_path = os.path.join(url_list[:url_list.rindex('\\')], 'url_floodlights.csv')
        floodlights.to_csv(save_path, encoding='utf-8', index=False)

    def urls_list(self):
        urls = self.list_of_urls.text()
        u = []

        for i in urls:
            u.append(i)

        print u


class TableWidgetCustom(QTableWidget):

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            self.copy()
        else:
            QTableWidget.keyPressEvent(self, event)

    def copy(self):
        ranges = self.selectedRanges()

        if len(ranges) < 1:
            return

        text = ''
        for r in ranges:
            r_top = r.topRow()
            r_bot = r.bottomRow()
            r_left = r.leftColumn()
            r_right = r.rightColumn()

            for row in range(r_top, r_bot + 1):
                for col in range(r_left, r_right + 1):
                    cell = self.item(row, col)
                    if cell:
                        text += cell.text()
                    text += '\t'
                text += '\n'

        QApplication().clipboard().setText(text)

if __name__ == '__main__':
    application = QApplication(sys.argv)
    frame = App()
    frame.show()
    application.exec_()