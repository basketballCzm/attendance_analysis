# -*- coding: utf-8 -*- 
import operator
import xlwt
import types


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

def writeDataset(result):
	wbk = xlwt.Workbook()
	sheet = wbk.add_sheet("sheet1")
	for row in range(len(result)):
		for col in range(len(result[row])):
			if(type(result[row][col]) is types.StringType):
				result[row][col] = result[row][col].decode('utf-8')
			sheet.write(row,col,result[row][col])
	wbk.save("student.xls")

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
	writeDataset(result)

main()
