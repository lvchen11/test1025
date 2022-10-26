import xlrd


#使用xlrd模块的open_workbook函数打开指定excel文件，并获得book对象
wb = xlrd.open_workbook("阿里巴巴2020年股票数据.xls")
#通过book对象的sheet_names方法获取所有表单名称
sheetnames = wb.sheet_names()
print(sheetnames)
#通过指定的表单名称获取sheet对象
sheet = wb.sheet_by_name(sheetnames[0])
#通过sheet对象的nrow和ncols属性获取表单的行数和列数
print(sheet.nrows,sheet.ncols)
for row in range(sheet.nrows):
    for col in range(sheet.ncols):
        #通过sheet对象的cell方法获取指定cell对象（单元格）
        #通过cell对象的value属性获取单元格的值
        value = sheet.cell(row,col).value
        #对除首行外的其他行进行数据格式化处理
        if row > 0:
            #第一列的xldate类型先转换成元组在格式化为“年月日”的格式
            if col == 0:
                # xldate_as_tuple函数的第二个参数只有0和1两个取值
                #其中0代表以1900-01-01为基准的日期，1代表以1904-01-01为基准的日期
                value = xlrd.xldate_as_tuple(value,0)
                value = f"{value[0]}年{value[1]:>02d}月{value[2]:>02d}日"
            else:
                value = f"{value:.2f}"
        print(value,end='\t')
    print()

#获取最后一个单元格的数据类型
#0 - 空值 1-字符串 2-字符串，3-日期，4-布尔，5-错误
last_cell_type = sheet.cell_type(sheet.nrows - 1,sheet.ncols - 1)
print(last_cell_type)

# 获取第一行的值(列表)
print(sheet.row_values(0))
#获取指定行指定列范围的数据（列表）
#第一个参数代表行索引，第二个和第三个参数代表列的开始（含）和结束（不含）索引
print(sheet.row_slice(3,0,5))


print("wwww","baidu","com",sep='.')
a=[]
for i in range(0,5):
    a.append(i)
    print(i,end=' ')

print(a)