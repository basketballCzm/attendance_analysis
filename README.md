# attendance_analysis
分析打卡机导出文件的打卡记录，将相应课程的考勤导出成.xls文件

## 使用说明
1. python的版本是2.7
2. 运行该程序python要安装pip工具，安装xlrd和xlwt包。安装参考步骤[pip安装xlrd和xlwt](https://www.jianshu.com/p/28d45b71f15f)

## v1.0
这个版本是生成的例如student 1.0形式的excel,对于这个数据还要根据example.xls来进行数据填写

## v2.0
这个版本是根据example.xls自动生成的student 2.0.xls文件，但是存在一个问题，有特殊的学号就需要手动校验。比如一个班出现SA16225219和SA17225219这两个学号，打卡机会记录16219和219的记录，但是我的程序中只是对后面三位数字进行校验，因此会出现两个人的打卡记录相与(&)的情况。因此如果你的班级中出现同学的后三位学号相同的情况，请务必自己校验这两个人的打卡情况。出现这种情况，也可以先生成1.0版本.xls再生成2.0版本.xls。避免查看.dat文件。

## v3.0
解决上述问题，在修改config.ini的时候注意不要使用windows自带的记事本进行修改配置文件，它会自动在文件都加BOM头，可以使用notepad++或sublime进行修改。


## 注意事项
1. config.txt中的样例时间段的选取必须按照给定的格式填写。例如：
2018-05-15 07:00:00 11:30:00
2. 每次删除以前生成的.xls文件后再生成新文件。
3. 从教务系统导出的上课名单必须满足examle.xls的格式
注意在代码中修改下列数据： 
数据都在第一页名字为Students
example_sheet = 'Students'
修改.xls名称
example_name = 'example.xls'