# attendance_analysis
分析打卡机导出文件的打卡记录，将相应课程的考勤导出成.xls文件

## 使用说明
1. python的版本是2.7
2. 运行该程序python要安装pip工具，安装xlrd和xlwt包。安装参考步骤[pip安装xlrd和xlwt](https://www.jianshu.com/p/28d45b71f15f)

##注意事项
1. config.txt中的样例时间段的选取必须按照给定的格式填写。例如：
2018-05-15 07:00:00 11:30:00
2. 每次删除以前生成的.xls文件后再生成新文件。