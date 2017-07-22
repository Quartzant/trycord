# -*- coding: utf-8 -*-
import xlrd
import xlwt


#设计表头
file = xlwt.Workbook()
table = file.add_sheet('info', cell_overwrite_ok=True)

file123 = xlrd.open_workbook('test.xlsx')
table1 = file123.sheets()[0]
table2 = file123.sheets()[1]

#收集行列数据
nrows1 = table1.nrows
nrows2 = table2.nrows
ncols1 = table1.ncols
ncols2 = table2.ncols

#填入新建工作表
for i in range(ncols2):
    table.write(0,i,table2.row_values(0)[i])

for i in range(ncols1):
   if i == 0 and 1 and 2:
      continue
   table.write(0,i+ncols2-1,table1.row_values(0)[i])



#写入所有表二数据
for i in range(nrows2):
    if i == 0:
        continue
    for x in range(ncols2):
       table.write(i, x, table2.row_values(i)[x])



#判断IP是否相同，同则写入，不同略过
for i in range(nrows2):
    for x in range(nrows1):
        if x == 0 :
           continue
        elif table2.row_values(i)[0] == table1.row_values(x)[0]:
            #table.write(nrows2,i,table2.row_values(x)[i])
            table.write(i,nrows1,  table1.row_values(x)[3])
            table.write(i,nrows1+1,table1.row_values(x)[4])
            table.write(i,nrows1+2, table1.row_values(x)[5])
            table.write(i,nrows1+3, table1.row_values(x)[6])
            table.write(i,nrows1+4, table1.row_values(x)[7])
            table.write(i,nrows1+5, table1.row_values(x)[8])
        else:
           continue



#保存文件
file.save('file.xls')

