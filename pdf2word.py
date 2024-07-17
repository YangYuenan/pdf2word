# -*- coding:utf-8 -*-

__author__ = 'yangyuenan'
__time__ = '2024/1/4 14:26'

import sys
from PyQt5.QtWidgets import QMainWindow, QWidget, QHBoxLayout ,QLineEdit, QVBoxLayout, QPushButton, QApplication, QFileDialog, QMessageBox
from pdf2docx import Converter


class gui(QMainWindow):

    def __init__(self, parent=None):
        super(gui, self).__init__(parent)
        self.exit = False
        self.szData = None
        self.setWindowTitle(u'PDF2Word')
        self.setMinimumWidth(500)
        self.setMinimumHeight(100)
        self.mainLayout = QVBoxLayout()
        self.widgetInit()

    def widgetInit(self):
        widget = QWidget()
        fileLayout = QHBoxLayout()
        self.setCentralWidget(widget)
        self.pdf = QLineEdit()
        selectButton = QPushButton(u'选择文件')
        addButton = QPushButton(u'开始转换')
        fileLayout.addWidget(self.pdf)
        fileLayout.addWidget(selectButton)
        self.mainLayout.addLayout(fileLayout)
        self.mainLayout.addWidget(addButton)
        widget.setLayout(self.mainLayout)
        addButton.clicked.connect(self.add)
        selectButton.clicked.connect(self.selectFile)

    def add(self):
        pdf = self.pdf.text()
        if not pdf.lower().endswith('pdf'):
            QMessageBox.information(self, u'提示', u'请选择PDF文件')
            return
        docxName = pdf[:-3] + 'docx'
        cv = Converter(pdf)
        cv.convert(docxName)
        cv.close()
        QMessageBox.information(self, u'提示', u'转换完成,\n文件已保存到' + docxName)

    def selectFile(self):
        fileName, fileType = QFileDialog.getOpenFileName(self, "选取文件", "", "All Files(*);;Text Files(*.pdf)")
        self.pdf.setText(fileName)
        if not fileName.lower().endswith('pdf'):
            QMessageBox.information(self, u'提示', u'请选择PDF文件')
            return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    m = gui()
    m.show()
    app.exec_()
