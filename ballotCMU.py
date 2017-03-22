# -*- coding: utf-8 -*-

import xlrd
import xlwt
from random import randint #randint(x,y)
import random
import time
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class Ballot:
	def __init__(self, fileName):
		self.FILENAME = fileName
		self.data = xlrd.open_workbook(self.FILENAME)
		self.table = self.data.sheets()[ 0 ] 
		self.nrows = self.table.nrows 
		self.ncols = self.table.ncols
		self.SAVEFILE = xlwt.Workbook()
		self.SAVETABLE = self.SAVEFILE.add_sheet('sheet name',cell_overwrite_ok=False)  

	def writeTableFirst(self):
		writwCol_Index = [0,1,3,11,12]
		for col in range(self.ncols):
			self.SAVETABLE.write(0, col, self.table.cell(0, col).value)
		for row in range(1,self.nrows):
			for col in writwCol_Index:
				self.SAVETABLE.write(row,col,self.table.cell(row,col).value)	

	def getDict(self):
		self.studentDict = []
		for row in range(1,self.nrows):
			person={}
			person['school'] = self.table.cell(row,2).value.replace(" ","")
			person['water'] = self.table.cell(row,4).value
			person['admissionID'] = self.table.cell(row,5).value
			person['name'] = self.table.cell(row,6).value.encode('utf-8').replace(" ","")
			person['cellphone'] = self.table.cell(row,7).value.replace(" ","")
			person['homenumber'] = self.table.cell(row,8).value.replace(" ","")
			person['identity'] = self.table.cell(row,9).value.encode('utf-8').replace(" ","")
			person['sequence'] = self.table.cell(row,10).value.encode('utf-8').replace(" ","")
			person['writeTime'] = self.table.cell(row,13).value
			self.studentDict.append(person)

	def getList(self, numberList): 
		rowCount = 1
		for index, each in enumerate(numberList):
			for i in range(each):
				text = str(index%11+1)+"-"+str(i+1) #!!!
				self.SAVETABLE.write(rowCount,14,str(index%11+1)) #!!!
				self.SAVETABLE.write(rowCount,15,self.table.cell(rowCount,10).value)
				self.SAVETABLE.write(rowCount,16,text)
				rowCount += 1

	def shuffleList(self, numberList):
		isWrite = [0 for i in range(sum(numberList))]
		# randomclass = randint(0, 10) #!!!
		randomclass = numberList.index(6)
		sessionList = ["第一梯次","第二梯次","第三梯次","第四梯次"]
		print "離島考生在第"+str(randomclass+1)+"試場"

		randomclassOrigin = randint(0, 10)
		while randomclassOrigin == randomclass:
			randomclassOrigin = randint(0, 10)
		print "原住民生在第"+str(randomclassOrigin+1)+"試場"

		for person in self.studentDict:
			if person['identity'] in ["離島考生", "原住民生"]:
				if person['sequence'] in sessionList:
					row = self.reBallot(sessionList, person, numberList, randomclass, randomclassOrigin)
					while isWrite[row-1] == 1:
						print "Collision!"
						row = self.reBallot(sessionList, person, numberList, randomclass, randomclassOrigin)
					isWrite[row-1] = 1
					self.writeData(row, person)
		
		fillIndex = 0
		randomAllList = [] # 一般考生
		for sessionIndex,session in enumerate(sessionList):
			randomList = []
			for person in self.studentDict:
				if person['identity'] == "一般考生":
					if person['sequence'] == session:
						randomList.append(person)
			random.shuffle(randomList)
			for person in randomList:
				randomAllList.append(person)
		while randomAllList:
			if isWrite[fillIndex] == 0:
				self.writeData(fillIndex+1,randomAllList[0])
				randomAllList.pop(0)
				fillIndex += 1
			else:
				fillIndex +=1 
		self.SAVEFILE.save('Result.xls')

	def reBallot(self, sessionList, person, numberList, randomclass, randomclassOrigin):
		if person['identity'] == "離島考生":
			randomclassInloop = randomclass
		elif person['identity'] == "原住民生":
			randomclassInloop = randomclassOrigin
		row = sessionList.index(person['sequence'])
		randomOrder = randint(1, numberList[row*11+randomclassInloop])#!!! # in class's order
		row = sum(numberList[:row*11+randomclassInloop])+randomOrder # add pervious class's sum plus class's order
		return row

	def writeData(self, row, person):
		self.SAVETABLE.write(row, 2, person['school'].decode('utf-8'))
		self.SAVETABLE.write(row, 4, person['water'])
		self.SAVETABLE.write(row, 5, person['admissionID'])
		self.SAVETABLE.write(row, 6, person['name'].decode('utf-8'))
		self.SAVETABLE.write(row, 7, person['cellphone'].decode('utf-8'))
		self.SAVETABLE.write(row, 8, person['homenumber'].decode('utf-8'))
		self.SAVETABLE.write(row, 9, person['identity'].decode('utf-8'))
		self.SAVETABLE.write(row, 10, person['sequence'].decode('utf-8'))
		self.SAVETABLE.write(row, 13, person['writeTime'])
