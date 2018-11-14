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

# 专业方向
stumajor = u'专业方向'

#-1表示学生未选择该门课程

def loadconfig(filename, split_ch, classnameSet=[]):
	cfg = ConfigParser()
	cfg.read(filename)
	classname = cfg.get('runconfig','classname').decode('utf-8')
	classnameL = classname.split(split_ch)
	for item in classnameL:
		classnameSet.append(item)

def loadsingleexcel(filename, singleexcelSet=[]):
	xlrd.Book.encoding = 'utf-8'
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
	stumajorcol = -1

	# 校验每个excel中学号,学生姓名,学生班级的行数
	# debug
	for x in range(len(excel_list[0])):
		if excel_list[0][x] == stuID:
			stuIDcol = x
		elif excel_list[0][x] == stuname:
			stunamecol = x
		elif excel_list[0][x] == stuclass:
			stuclasscol = x
		elif excel_list[0][x] == stumajor:
			stumajorcol = x

	singleexcelSet.append(stuIDcol)
	singleexcelSet.append(stunamecol)
	singleexcelSet.append(stuclasscol)
	singleexcelSet.append(stumajorcol)
	singleexcelSet.append(excel_list)


def loadexcel(classnameSet, alldataSet=[]):
	for x in range(len(classnameSet)):
		singleexcelSet = []
		loadsingleexcel(classnameSet[x], singleexcelSet)
		alldataSet.append(singleexcelSet)

def filterdata(alldataSet, filterdataSet=[]):
	L = []
	L.append(stuID)
	L.append(stuname)
	L.append(stuclass)
	L.append(stumajor)
	filterdataSet.append(L)

	for x in range(len(alldataSet)):
		# 假设学号，姓名，班级，专业方向是乱序
		tmp_stuIDcol = alldataSet[x][0]
		tmp_stunamecol = alldataSet[x][1]
		tmp_stuclasscol = alldataSet[x][2]
		tmp_stumajorcol = alldataSet[x][3]
		for y in range(1,len(alldataSet[x][4])):
			#print (len(alldataSet[x][4]))
			L = []
			if (1 == len(filterdataSet)):
				L.append(alldataSet[x][4][y][tmp_stuIDcol])
				L.append(alldataSet[x][4][y][tmp_stunamecol])
				L.append(alldataSet[x][4][y][tmp_stuclasscol])
				L.append(alldataSet[x][4][y][tmp_stumajorcol])
				print (L)
				filterdataSet.append(L)
			else:
				for k in range(1,len(filterdataSet)):
					# filterdataSet中有相同的数据集
					if (alldataSet[x][4][y][tmp_stuIDcol] == filterdataSet[k][0]):
						break
					# filterdataSet中没有相同的数据集
					elif (k == len(filterdataSet)-1):
						L.append(alldataSet[x][4][y][tmp_stuIDcol])
						L.append(alldataSet[x][4][y][tmp_stunamecol])
						L.append(alldataSet[x][4][y][tmp_stuclasscol])
						L.append(alldataSet[x][4][y][tmp_stumajorcol])
						filterdataSet.append(L)
						break

def addextradata(alldataSet, filterdataSet, classnameSet):
	nrows = len(filterdataSet)
	ncols = len(filterdataSet[0])
	nrowsclassname = len(classnameSet)
	
	for x in range(0,nrows):
		for y in range(nrowsclassname):
			if(x == 0):
				filterdataSet[x].append(classnameSet[y])
				continue
			filterdataSet[x].append(-1)

	
	for x in range(len(filterdataSet)):
		for y in range(len(alldataSet)):
			tmp_stuIDcol = alldataSet[y][0]
			tmp_stunamecol = alldataSet[y][1]
			tmp_stuclasscol = alldataSet[y][2]
			tmp_stumajorcol = alldataSet[y][3]
			for k in range(1,len(alldataSet[y][4])):
				if (alldataSet[y][4][k][tmp_stuIDcol] == filterdataSet[x][0]):
					sum = 0
					for z in range(4,len(alldataSet[y][4][k])):
						if (alldataSet[y][4][k][z] == 0):
							sum = sum + 1
					filterdataSet[x][ncols+y] = sum
	


def writedatasettoexcel(filterdataSet):
	wbk = xlwt.Workbook()
	sheet = wbk.add_sheet('sheet1')
	for row in range(len(filterdataSet)):
		for col in range(len(filterdataSet[row])):
			if(type(filterdataSet[row][col]) is types.StringType):
				filterdataSet[row][col] = filterdataSet[row][col].decode('utf-8')
			sheet.write(row, col, filterdataSet[row][col])

	wbk.save('student 3.0.xls')



def main():
	classnameSet=[]
	alldataSet=[]
	filterdataSet=[]
	loadconfig('config.ini', ';', classnameSet)
	print (classnameSet)
	loadexcel(classnameSet, alldataSet)
	filterdata(alldataSet, filterdataSet)
	# print (filterdataSet)
	addextradata(alldataSet, filterdataSet, classnameSet)
	writedatasettoexcel(filterdataSet)
	raw_input("success...")
main()