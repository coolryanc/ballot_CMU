# -*- coding: utf-8 -*-

import xlrd
import xlwt
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from random import randint #randint(x,y)
import random
import time
import os
import sys
import ballotCMU
reload(sys)
sys.setdefaultencoding('utf-8')

class MainWindow(QStackedWidget):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setGeometry(0,0,720,540)
        self.setStyleSheet("background-color:white")
        self.setWindowTitle(QString(u"CMU 藥學系試場順序抽籤"))

        window = QWidget()
        window.setStyleSheet("background-color:white;")

        ballot = HoverButton(window,"BALLOT","font-size:20px;background-color:#5F5C5C;\
                          color:#E4E4E4","font-size:23px;background-color:#3c393a;color:#ffffff")
        ballot.clicked.connect(self.fillNumberFunction)
        chooseFile = HoverButton(window,"FILE","font-size:20px;background-color:#5F5C5C;\
                          color:#E4E4E4","font-size:23px;background-color:#3c393a;color:#ffffff")
        chooseFile.clicked.connect(self.choosefileFunction)

        mcuImage = QLabel()
        mcuImage.setPixmap(QPixmap("logo.jpg"))


        layout = QVBoxLayout()
        layout.addStretch(1)
        layout.addWidget(mcuImage,0,Qt.AlignCenter)
        layout.addWidget(chooseFile,0,Qt.AlignCenter)
        layout.addWidget(ballot,0,Qt.AlignCenter)
        layout.addStretch(1)

        window.setLayout(layout)

        self.addWidget(window)
        self.setCurrentWidget(window)

    def fillNumberFunction(self):
      	fillNumberWindow = QWidget()
      	fillNumberWindow.setStyleSheet("background-color:white;")
      	grid = QGridLayout()
      	grid.setColumnStretch(1, 0)
        grid.setColumnStretch(2, 0)

        self.keepLineEdit = []
        defaultText = ["6","3","4","4","4","4","4","4","4","4","4"]
        random.shuffle(defaultText)
        defaultText += ["11" for i in range(11)]
        defaultText += ["9" for i in range(11)]
        defaultText += ["10" for i in range(11)]

        # defaultText = [8,8,8,7,7,7,7,7,7,7]
        # defaultText += [9,9,9,10,9,9,9,9,9,9]
        # defaultText += [6 for i in range(10)]
        # defaultText += [9,9,9,9,10,10,10,9,9,9]
        
        defaultTextCount=0
        # _rowcolID = [(i,j) for i in range(5)for j in range(11)]
        _rowcolID = [(i,j) for i in range(5)for j in range(12)]
       	
       	for element in _rowcolID:
       		if element[0]!=0 and element[1]!=0:
       			lineEdit = Numberenter(fillNumberWindow, element[0], element[1], str(defaultText[defaultTextCount]))
       			defaultTextCount += 1
       			self.keepLineEdit.append(lineEdit)
       			grid.addWidget(lineEdit, *element)
       		elif element[1] == 0:
       			if element[0] != 0:
       				titleLabel = QLabel(unicode("第"+str(element[0])+"梯次", 'utf8'))
       				grid.addWidget(titleLabel, *element)

       	startBtn = HoverButton(fillNumberWindow,"START","font-size:20px;background-color:#5F5C5C;\
                          color:#E4E4E4","font-size:23px;background-color:#3c393a;color:#ffffff")
       	startBtn.clicked.connect(self.getEnterNumber)

       	layout = QHBoxLayout()
       	layout.addStretch(1)
       	layout.addLayout(grid)
       	layout.addStretch(1)

       	vlayout = QVBoxLayout()
       	vlayout.addStretch(1)
       	vlayout.addLayout(layout)
       	vlayout.addWidget(startBtn,0,Qt.AlignCenter)
       	vlayout.addStretch(1)

    	fillNumberWindow.setLayout(vlayout)
    	self.addWidget(fillNumberWindow)
    	self.setCurrentWidget(fillNumberWindow)

    def choosefileFunction(self):
    	fileName = QFileDialog.getOpenFileName(self, 'Open')
    	self.getFileName = str(fileName).decode('utf-8')
    	self.GUI_ballot = ballotCMU.Ballot(self.getFileName)
    	self.GUI_ballot.writeTableFirst()
    	self.GUI_ballot.getDict()

    def getEnterNumber(self):
    	passNumber = []
    	for index in range(len(self.keepLineEdit)):
    		num = int(str(self.keepLineEdit[index].text()))
    		passNumber.append(num)
    	self.GUI_ballot.getList(passNumber)
    	self.GUI_ballot.shuffleList(passNumber)
    	self.ballotFunction()

    def ballotFunction(self):
    	ballotWindow = QWidget()
        ballotWindow.setStyleSheet("background-color:white;")

        finishBtn = HoverButton(ballotWindow,"FINISH","font-size:20px;background-color:#5F5C5C;\
                          color:#E4E4E4","font-size:23px;background-color:#3c393a;color:#ffffff")
        finishBtn.clicked.connect(self.exitAPP)
        finishBtn.setVisible(False)
        textLabel = BallotLabel(ballotWindow, finishBtn)

        layout = QVBoxLayout()
        layout.addStretch(1)
        layout.addWidget(textLabel,0,Qt.AlignCenter)
        layout.addWidget(finishBtn,0,Qt.AlignCenter)
        layout.addStretch(1)

        ballotWindow.setLayout(layout)

        self.addWidget(ballotWindow)
        self.setCurrentWidget(ballotWindow)

    def exitAPP(self):
    	QCoreApplication.instance().quit()

class HoverButton(QPushButton):
    def __init__(self,  *args):
        super(HoverButton, self).__init__()
        self.setMouseTracking(True)
        self.setText(args[1])
        self.style1 = args[2]+";width:172px;height:36px; border-radius: 15px;margin-top:36px;"
        self.style2 = args[3]+";width:172px;height:36px; border-radius: 15px;margin-top:36px;"
        self.setStyleSheet(self.style1)

    def enterEvent(self,event):
        self.setStyleSheet(self.style2)

    def leaveEvent(self,event):
        self.setStyleSheet(self.style1)

class Numberenter(QLineEdit):
	def __init__(self, *args):
		super(Numberenter, self).__init__()
		self.ROWID = args[1]
		self.COLID = args[2]
		self.setPlaceholderText("Number")
		self.setAlignment(Qt.AlignCenter)
		self.setMaxLength(2)
		self.setText(args[3])

	def keyPressEvent(self, event):
		if event.key() == Qt.Key_Tab:
			event.accept()
		else:
			QLineEdit.keyPressEvent(self, event)

class BallotLabel(QLabel):
	def __init__(self, *args):
		super(BallotLabel, self).__init__()
		self.setText(QString(u"抽籤中"))
		self.setStyleSheet("font-size:30px")
		self.btn = args[1]
		self.timer = QTimer()
		self.showResult()
	
	def showResult(self):
		self.data = xlrd.open_workbook('Result.xls')
		self.table = self.data.sheets()[ 0 ] 
		self.nrows = self.table.nrows 

		self.rowIndex = 1  
		self.timer.setInterval(100)
		self.timer.timeout.connect(lambda: self.setResultText(self.rowIndex))
		self.timer.start()

	def setResultText(self,row):
		if self.rowIndex == (self.nrows):
			self.setText(QString(u"抽籤完成 檔案儲存完畢:Result.xls"))
			self.btn.setVisible(True)		
		else:
			_admissionID = str(self.table.cell(row,5).value)
			_name = self.table.cell(row,6).value.encode('utf-8')
			_identity = self.table.cell(row,9).value.encode('utf-8')
			_session = self.table.cell(row,15).value.encode('utf-8').replace(" ","")
			_sequence = self.table.cell(row,16).value.encode('utf-8').replace("-"," ")
			_sequence = [str(s) for s in _sequence.split() if s.isdigit()]
			_allInfo = _admissionID+" "+_name+" "+_identity+" "+_session+" "+"第"+_sequence[0]+"試場第"+_sequence[1]+"順位"
			self.setText(QString(unicode(_allInfo, 'utf8')))
			self.rowIndex += 1


if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()