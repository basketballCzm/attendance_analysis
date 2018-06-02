# -*- coding: utf-8 -*- 
import operator
import xlwt
import xlrd
import types

# 数据都在第一页名字为Students
example_sheet = 'Students'
# 修改.xls名称
example_name = 'example.xls'

def loadDataset(filename, split_ch, trainingSet=[]):
	dataset = []
	for line in open(filename,'r').readlines():
		m = line.split(split_ch)
		for item in m:
			if item.replace('.','').replace('-','').replace('\n','').isdigit():
				if item.count('.')>0:
					dataset.append(float(item))
				else:
					dataset.append(item)
					#dataset.append(int(item))
			else:
				dataset.append(item)
		trainingSet.append(dataset)
		dataset = []

def checkDataset(trainingSet, checkSet, result=[]):
	for x in range(len(checkSet)):
		for y in range(len(trainingSet)):
			if((checkSet[x][0]==trainingSet[y][1][0:10]) & (trainingSet[y][1][11:19]>checkSet[x][1]) & (trainingSet[y][1][11:19]<checkSet[x][2])):
				result.append(trainingSet[y])

def writeDataset(result,checkSet):
	'''
	wbk = xlwt.Workbook()
	sheet = wbk.add_sheet("sheet1")
	for row in range(len(result)):
		for col in range(len(result[row])):
			if(type(result[row][col]) is types.StringType):
				result[row][col] = result[row][col].decode('utf-8')
			sheet.write(row,col,result[row][col])
	wbk.save("student.xls")
	'''
	#读取example.xls
	xlrd.Book.encoding = "gbk"
	try:
		data = xlrd.open_workbook(example_name)
	except Exception,e:
		print str(e)

	table = data.sheet_by_name(example_sheet)
	nrows = table.nrows
	ncols = table.ncols
	excel_list = []
	for row in range(nrows):
		excel_rows = []
		for col in range(ncols):
			cell_value = table.cell(row,col).value
			#UTF-8?
			excel_rows.append(cell_value.encode('gbk'))
		excel_list.append(excel_rows)


	for x in range(nrows):
		for y in range(len(checkSet)):
			if(0 == x):
				excel_list[x].append(checkSet[y][0])
			else:
				excel_list[x].append(0)

	#添加后面的行
	for x in range(len(result)):
		for y in range(len(excel_list)):
			if(result[x][0][-3:] == excel_list[y][0][-3:]):
				for k in range(len(checkSet)):
					#print (result[x][1][0:10],'  =  ',checkSet[k])
					if(result[x][1][0:10] == checkSet[k][0]):
						excel_list[y][ncols+k] = 1


	wbk = xlwt.Workbook()
	sheet = wbk.add_sheet("sheet1")
	for row in range(len(excel_list)):
		for col in range(len(excel_list[row])):
			if(type(excel_list[row][col]) is types.StringType):
				excel_list[row][col] = excel_list[row][col].decode('gbk')
			sheet.write(row,col,excel_list[row][col])
	wbk.save("student 2.0.xls")

	#print str(excel_list).decode("string_escape")



def main():
	trainingSet = []
	checkSet = []
	result = []
	loadDataset('data.dat','\t',trainingSet)
	loadDataset('config.txt',' ',checkSet)

	checkDataset(trainingSet,checkSet,result)
	result.sort(key=operator.itemgetter(0))
	#for x in range(len(result)):
	#	print(result[x])
	writeDataset(result,checkSet)
	print ('success...')

main()
