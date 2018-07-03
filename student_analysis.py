# -*- coding: utf-8 -*- 
import operator
import xlwt
import xlrd
import types

# 数据都在第一页名字为Students
example_sheet = 'Students'
# 修改.xls名称
example_name = 'example.xls'
# 查找的主键是学号
stuID = u'学号'
# 学号的默认行数为1
stuIDcol = 0
# 打卡机第一项的空格数
numofspace = 6
# 特殊的空格数
spnumofspace = 4

def loadDataset(filename, split_ch, trainingSet=[]):
	dataset = []
	for line in open(filename,'r').readlines():
		m = line.split(split_ch)
		for item in m:
			if item.replace('.','').replace('-','').replace('\n','').isdigit():
				if item.count('.')>0:
					#dataset.append(float(item))
					dataset.append(item)
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
	except Exception as e:
		print (str(e))

	table = data.sheet_by_name(example_sheet)
	#原始表格的长和宽
	nrows = table.nrows
	ncols = table.ncols
	excel_list = []
	for row in range(nrows):
		excel_rows = []
		for col in range(ncols):
			cell_value = table.cell(row,col).value
			#UTF-8?
			if (type(cell_value) is types.StringType):
				excel_rows.append(cell_value.encode('gbk'))
			else:
				excel_rows.append(cell_value)
		excel_list.append(excel_rows)

	for x in range(len(excel_list[0])):
		# '学号' str                   u'学号' unicode 
		# type(excel_list[0][x].encode('gbk')) str
		# excel_list[0][x]                     unicode
		# print type(stuID) 
		#print type(excel_list[0][x].encode('gbk')) #unicode
		if stuID == excel_list[0][x]:
			stuIDcol = x
			break

	for x in range(nrows):
		for y in range(len(checkSet)):
			if(0 == x):
				excel_list[x].append(checkSet[y][0])
			else:
				excel_list[x].append(0)

	# SA114 17114 100 15225086  16011066 S17225324 S17225324 B15011041
	#添加后面的行
	for x in range(len(result)):
		for y in range(len(excel_list)):
			# 卡机打出SA114和17114之类的信息
			if (result[x][0].count(' ') == spnumofspace):

				# 打卡机学号为'    17423'
				if ((result[x][0][-4]>='0') & (result[x][0][-4]<='9') & (result[x][0][-5]>='0') & (result[x][0][-5]<='9')):
					# 组合excle中的学号行
					tmp_str = excel_list[y][stuIDcol][2:4] + excel_list[y][stuIDcol][-3:]
					if(result[x][0][-5:] == tmp_str):
						# 查看是哪一天
						for k in range(len(checkSet)):
							#print (result[x][1][0:10],'  =  ',checkSet[k])
							if(result[x][1][0:10] == checkSet[k][0]):
								excel_list[y][ncols+k] = 1

				elif ((result[x][0][-4]>='A') & (result[x][0][-4]<='Z') & (result[x][0][-5]>='A') & (result[x][0][-5]<='Z')):
					# 组合excle中的学号行
					tmp_str = excel_list[y][stuIDcol][0:2] + excel_list[y][stuIDcol][-3:]
					if(result[x][0][-5:] == tmp_str):
						# 查看是哪一天
						for k in range(len(checkSet)):
							#print (result[x][1][0:10],'  =  ',checkSet[k])
							if(result[x][1][0:10] == checkSet[k][0]):
								excel_list[y][ncols+k] = 1
			# 匹配学号
			#elif (result[x][0].count(' ') == numofspace):
			#出现这样的学号 S17225324 还是只匹配最后3位
			else:
				if(result[x][0][-3:] == excel_list[y][stuIDcol][-3:]):
					# 查看是哪一天
					for k in range(len(checkSet)):
						#print (result[x][1][0:10],'  =  ',checkSet[k])
						if(result[x][1][0:10] == checkSet[k][0]):
							excel_list[y][ncols+k] = 1

	excel_list.append([])
	# 这里用'gbk'乱码
	excel_list[-1].append('应到'.decode('utf-8'))

	excel_list.append([])
	excel_list[-1].append('实到'.decode('utf-8'))

	excel_list.append([])
	excel_list[-1].append('缺勤'.decode('utf-8'))

	# 下面3行全部为0
	for x in range(3):
		for y in range(1,ncols+len(checkSet)):
			excel_list[-(x+1)].append('')

	for x in range(len(checkSet)):
		cnt = 0
		for y in range(nrows):
			if 1 == excel_list[y][ncols+x]:
				cnt = cnt + 1
		excel_list[nrows][ncols+x] = nrows-1
		excel_list[nrows+1][ncols+x] = cnt
 		excel_list[nrows+2][ncols+x] = nrows-cnt-1
		

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
	'''
	for x in range(len(result)):
		print (result[x][0].count(' '))
		print (len(result[x][0]))

	for x in range(9):
		print (result[10][0][x])
 	#	print result[x]
 	'''

	for x in range(0,len(result)):
		print(result[x])
	writeDataset(result,checkSet)
	print ('success...')

main()
