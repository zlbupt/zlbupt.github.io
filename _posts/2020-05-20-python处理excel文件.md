---
layout:     post
title:      python 处理excel 文件
subtitle:   
date:       2020-05-20
author:     zl
header-img: img/post-bg-ios9-web.jpg
catalog: true
tags:
    - python
    - excel
---

> 总结的一部分使用python 处理excel 的工具 



# xlwings

对已经打开的excel进行操作，如果excel没有打开，则打开该excel，也可以使用VBA宏不显示excle的客户端。底层使用微软的VBA进行处理，所以在excel中可以看到其对excel进行的操作。其中大部分功能依然要靠VBA宏进行处理。

[xlwings参考文档](https://docs.xlwings.org/en/stable/)

由于xlwings大部分功能依然使用微软VBA的宏，所以列出VBA的[参考文档](https://docs.microsoft.com/en-us/office/vba/api/excel.range(object))

```python
import xlwings as xw

# 设置程序不可见运行
# 第一种打开excel的方法
app = xw.App(visible=False, add_book=False)
wb = app.books.add()
ws = wb.sheets.active

# 第二种打开excel的方法
wb = xw.Book()
sht = wb.sheets['Sheet1']

# 两种读取或者设置值的方式
sht.range('A1').value = 'Foo 1'
sht.range('A1').value
sht.range('A1:B2').value
sht.range((1, 2)).value = 'Foo 2'
sht.range((1, 2)).value 
sht.range((1, 2), (1, 6)).value 

# 使用VBA获取已经使用的行数和列数
info = sht.used_range
nrows = info.last_cell.row
ncolumns = info.last_cell.column

# 或者使用VBA
rows = sht.api.UsedRange.Rows.count
cols = sht.api.UsedRange.Columns.count

# 使用VBA合并单元格
sht.range("A1:B2").api.merge()
sht.range((1, 1), (1,3)).api.merge()

# 获取和设置当前格子的高度和宽度
sht.range("A1").width
sht.range("A1").height
sht.range("A1").row_height = value
sht.range("A1").column_width = value

# 设置单元格格式
sht.range("A1").api.HorizontalAlignment = -4152
# -4131 ：left， -4152 ：right， -4108 ：中
sht.range("A1").api.VerticalAlignment = -4107
# -4160：top , -4108 : center, -4107 : bottom

### 其他具体操作可以参考官方文档或者VBA宏
```



# openpyxl

处理未被打开的excel，功能很齐全。把已经打开的excel视为只读文件，只能对其进行读取，不能对此进行写入操作



# 

