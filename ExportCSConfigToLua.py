#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import xlrd
import sys
import os
import datetime
reload(sys)
sys.setdefaultencoding('utf-8')
# fileOutPut = open('')
CURDIR = os.path.dirname(os.path.realpath(__file__)) + '/'
workbook = xlrd.open_workbook(CURDIR+'point.xlsx')
print ("There are %d sheets in the workbook."%(workbook.nsheets))

def getValueOfCell(booksheet, row, col):
		vType = booksheet.cell(2, col).value
		value = booksheet.cell(row, col).value
		if vType == 'STR': #string
			value = str(value)
		elif vType == 'INT': #number
			value = int(value)
		elif vType == 'BOOL': #boolean
			value = str(bool(value)).lower()
		elif vType == 'TABLE':
			value = str(value)
		return vType, value


for booksheet in workbook.sheets():
	print("The name of sheet is %s"%(booksheet.name))
	filePath = ''.join([CURDIR,booksheet.name,".lua"])
	fileOutPut = open(filePath, 'w')
	today = datetime.date.today()
	writeData = "-- @author: wangbing\n-- " + str(today) + "\n\n"

	writeData = ''.join([writeData, 'local ', booksheet.name, '={\n'])

	for row in range(booksheet.nrows):
		if (row != 0 and row != 1 and row != 2):
			itemType , itemId = getValueOfCell(booksheet,row, 0)
			print(itemId)
			writeData = ''.join([writeData, '\t', str(itemId), ' = {["'])
			for col in range(1,booksheet.ncols):
				tile = booksheet.cell(0, col).value
				valueType, value = getValueOfCell(booksheet, row, col)
				if valueType == 'TABLE':
					if str(value) == "nil":
						writeData = ''.join([writeData, tile, '"] = nil'])
					else:
						writeData = ''.join([writeData, tile, '"] = {', str(value), '}'])
				elif valueType == "STR":
					writeData = ''.join([writeData, tile, '"] = "', str(value), '"'])
				else:
					writeData = ''.join([writeData, tile, '"] = ', str(value)])
				
				if col != booksheet.ncols-1:
					writeData = ''.join([writeData, ', ["'])
			else:
				writeData = ''.join([writeData, '},\n'])


	writeData = ''.join([writeData, '}\n\nreturn ', booksheet.name])

	fileOutPut.write(writeData)
	fileOutPut.close()