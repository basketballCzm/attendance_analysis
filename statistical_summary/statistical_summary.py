# -*- coding: utf-8 -*-
import xlwt
import xlrd
import types
from ConfigParser import ConfigParser

# 学号行数
stuID = u'学号'

# 姓名行数
stuname = u'姓名'

# 班级行数
stuclass = u'班级'

def loadconfig(filename, split_ch, classnameSet=[]):
	cfg = ConfigParser()
	cfg.read(filename)
	classname = cfg.get('runconfig','classname').decode('utf-8')
	classnameSet = classname.split(split_ch)

def loadsingleexcel(filename, singleexcelSet=[]):
	xlrd.Book.encoding('utf-8')
	try:
		data = xlrd.open_workbook(filename)
	except Exception as e:
		print (str(e))

	table = data.sheet_by_index(0)
	nrows = table.nrows
	ncols = table.ncols
	excel_list = []
	for row in range(nrows):
		excel_rows = []
		for col in range(ncols):
			cell_value = table.cell(row,col).value
			if(type(cell_value) == types.StringType):
				excel_rows.append(cell_value.encode('utf-8'))
			else:
				excel_rows.append(cell_value)
		excel_list.append(excel_rows)

	
	stuIDcol = -1
	stunamecol = -1
	stuclasscol = -1
	# 校验每个excel中学号,学生姓名,学生班级的行数
	for x in range(len(excel_list[0])):
		if cell_value == stuID:
			stuIDcol = col
		elif cell_value == stuname:
			stunamecol = col
		elif cell_value == stuclass:
			stuclasscol = col

	singleexcelSet.append(stuIDcol)
	singleexcelSet.append(stunamecol)
	singleexcelSet.append(stuclasscol)
	singleexcelSet.append(excel_list)


def loadexcel(classnameSet, alldataSet=[]):
	for x in range(len(classnameSet)):
		singleexcelSet = []
		loadsingleexcel(classnameSet[x], singleexcelSet)
		alldataSet.append(singleexcelSet)

def filterdata(alldataSet, filterdataSet):
	pass

def writedatasettoexcel(filterdataSet):
	pass


def main():
	classnameSet=[]
	alldataSet=[]
	loadconfig('config.ini', ';', classnameSet)


main()