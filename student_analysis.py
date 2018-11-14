# -*- coding: utf-8 -*- 
import operator
import xlwt
import xlrd
import types
# python3  from configparser import ConfigParser
from ConfigParser import ConfigParser

# 数据集的名称
example_data = u'data.dat'
# 数据都在第一页名字为Students
example_sheet = u'Students'
# 修改.xls名称
example_name = u'example.xls'
# 查找的主键是学号
stuID = u'学号'
# 学生姓名
stuName = u'姓名'
# 学生班级
stuClass = u'班级'
# 学生专业方向
stuMajor = u'专业方向'
# 学号的默认行数为1
stuIDcol = 0

# 打卡机导出项 S17225001 B15225001
zeronumofspace = 0

# 打卡机导出项 17225001
onenumofspace = 1

# 打卡机导出项 SB17001
twonumofspace = 2

# 打卡机导出项 B17001
threenumofspace = 3

# 打卡机导出项 SA001 17001
fournumofspace = 4

# 打卡机导出项 001
sixnumofspace = 6

def loadConfig(filename, split_ch, checkSet=[]):
	cfg = ConfigParser()
	cfg.read(filename)
	# 如果要改变全局变量的值就需要声明global
	# windows 下面的汉字是用'utf-8'编码
	global example_name,example_sheet,example_data
	example_data = cfg.get('runconfig', 'data_name').decode('utf-8')
	example_name = cfg.get('runconfig', 'excel_name').decode('utf-8')
	example_sheet = cfg.get('runconfig', 'sheet_name').decode('utf-8')
	classtime = cfg.get('time', 'classtime')
	m = classtime.split(split_ch)
	for item in m:
		L = []
		tmp_m = item.split(' ')
		for tmp_item in tmp_m:
			L.append(tmp_item)
		checkSet.append(L)

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
	#读取example.xls
	xlrd.Book.encoding = "utf-8"
	try:
		data = xlrd.open_workbook(example_name)
	except Exception as e:
		print (str(e))

	# table = data.sheet_by_index(0)
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
				excel_rows.append(cell_value.encode('utf-8'))
			else:
				excel_rows.append(cell_value)
		excel_list.append(excel_rows)

	# 查看学号在哪一行
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
				# 添加第一项时间
				excel_list[x].append(checkSet[y][0])
			else:
				# 初始化后面的项全部为0
				excel_list[x].append(0)

	# RY001 17525 525 100 15225086  16011066 S17225324 S17225324 B15011041 S16011150 15225086 16437
	# 100 RY001 16437 17525 525 15225086 16011066 S17225324 B15011041
	#添加后面的行
	for x in range(len(result)):
		for y in range(len(excel_list)):
			# 卡机打出SA114和17114之类的信息
			tmp_excel_list = ''
			tmp_result = ''
			if (result[x][0].count(' ') == zeronumofspace):
				# 打卡机出现B15011041 和 S17225324
				if(result[x][0][0] == 'S'):
					tmp_excel_list = excel_list[y][stuIDcol][0] + excel_list[y][stuIDcol][2:]
				elif(result[x][0][0] == 'B'):
					tmp_excel_list = excel_list[y][stuIDcol][1:]
				tmp_result = result[x][0]

			elif (result[x][0].count(' ') == onenumofspace):
				# 打卡机打出' 15225086' ' 16011066'
				tmp_excel_list = excel_list[y][stuIDcol][2:]
				tmp_result = result[x][0][1:]

			elif (result[x][0].count(' ') == twonumofspace):
				# 打卡机打出 '  SB17001'
				tmp_excel_list = excel_list[y][stuIDcol][0:4] + excel_list[y][stuIDcol][-3:]
				tmp_result = result[x][0][2:]

			elif (result[x][0].count(' ') == threenumofspace):
				# 打卡机打出 '   B17001'
				tmp_excel_list = excel_list[y][stuIDcol][1:4] + excel_list[y][stuIDcol][-3:]
				tmp_result = result[x][0][3:]

			elif (result[x][0].count(' ') == fournumofspace):
				# 打卡机打出为'    17423' '    RY001'
				if ((result[x][0][-4]>='0') & (result[x][0][-4]<='9') & (result[x][0][-5]>='0') & (result[x][0][-5]<='9')):
					# 组合excle中的学号行
					tmp_excel_list = excel_list[y][stuIDcol][2:4] + excel_list[y][stuIDcol][-3:]
					tmp_result = result[x][0][-5:]

				elif ((result[x][0][-4]>='A') & (result[x][0][-4]<='Z') & (result[x][0][-5]>='A') & (result[x][0][-5]<='Z')):
					# 组合excle中的学号行
					tmp_excel_list = excel_list[y][stuIDcol][0:2] + excel_list[y][stuIDcol][-3:]
					tmp_result = result[x][0][-5:]

			elif (result[x][0].count(' ') == sixnumofspace):
				#打卡机打出 '      001'
				tmp_excel_list = excel_list[y][stuIDcol][-3:]
				tmp_result = result[x][0][-3:]
				
			else:
				#出现上面没有包含的情况还是只匹配最后3位
				tmp_excel_list = excel_list[y][stuIDcol][-3:]
				tmp_result = result[x][0][-3:]

			if(tmp_result == tmp_excel_list):
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
				excel_list[row][col] = excel_list[row][col].decode('utf-8')
			sheet.write(row,col,excel_list[row][col])
	wbk.save("student 3.0.xls")

	#print str(excel_list).decode("string_escape")



def main():
	trainingSet = []
	checkSet = []
	result = []
	loadConfig('config.ini',';',checkSet)
	loadDataset(example_data,'\t',trainingSet)

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
	#for x in range(0,len(result)):
	#	print(result[x])
	writeDataset(result,checkSet)
	raw_input("success...")

main()
