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

# 保存或者另存为

wb.save()
wb.save(new_path)

# 两种读取或者设置值的方式

sht.range('A1').value = 'Foo 1'
sht.range('A1').value
sht.range('A1:B2').value # 使用多个单元格
sht.range((1, 2)).value = 'Foo 2'
sht.range((1, 2)).value 
sht.range((1, 2), (1, 6)).value 

# 使用VBA获取已经使用的行数和列数

info = sht.used_range
nrows = info.last_cell.row
ncolumns = info.last_cell.column

# 设置或者读取单元格的背景颜色, 但是只能使用rgb元组

sht.range('A1').color = (0,0,0)

# 清除内容和格式

sht.range('A1').clear()

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

# 设置单元格居中格式

sht.range("A1").api.HorizontalAlignment = -4152
# -4131 ：left， -4152 ：right， -4108 ：中

sht.range("A1").api.VerticalAlignment = -4107
# -4160：top , -4108 : center, -4107 : bottom

# 设置边框

# Borders(9) 底部边框，LineStyle = 1 直线。
sht.range('C2').api.Borders(9).LineStyle = 1
sht.range('C2').api.Borders(9).Weight = 3  # 设置边框粗细。


# Borders(7) 左边框，LineStyle = 2 虚线。

sht.range('C2').api.Borders(7).LineStyle = 2
sht.range('C2').api.Borders(7).Weight = 3

# Borders(8) 顶部框，LineStyle = 5 双点划线。

sht.range('C2').api.Borders(8).LineStyle = 5
sht.range('C2').api.Borders(8).Weight = 3

# Borders(10) 右边框，LineStyle = 4 点划线。

sht.range('C2').api.Borders(10).LineStyle = 4
sht.range('C2').api.Borders(10).Weight = 3

### 其他具体操作可以参考官方文档或者VBA宏
```



# openpyxl

处理未被打开的excel，功能很齐全。把已经打开的excel视为只读文件，只能对其进行读取，不能对此进行写入操作。

[参考文档](https://openpyxl.readthedocs.io/en/stable/)

```python
# 新建一个文件
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
# 创建工作表

ws1 = wb.create_sheet()
# 创建工作表并插入position

ws2 = wb.create_sheet(position)
# 给工作表重新命名

ws.title = 'Sheet1'
# 得到某个工作表

ws = wb[name]
ws = wb.get_sheet_by_name(name)
# 获得所有的表

wb.get_sheet_names()

# 使用已经存在的表

wb.load_workbook(path)
'''operator'''

#保存与另存为
# 注意：当一个工作表被创建是，其中不包含单元格。只有当单元格被获取是才被创建。这种方式我们不会创建我们从不会使用的单元格，从而减少了内存消耗。而且对单元格只有保存后才会改变excel中。

wb.save(path)

## 操作单元格的几种方式

ws['A1'] = 1 #赋值
a = ws['A2'] # 取值
ws.cell('A1') = 1 .value
b = ws.cell('A1')
ws.cell(row=1, coloum=2).value

# 使用多个单元格

ws['A1':'A3']
# 按照行列操作

for row in ws.iter_rows(min_row=1, max_row=3,
                        min_col=1, max_col=2):
    for cell in row:
        print(cell)
# 合并单元格

ws.merge_cells('F1:G1')
# 或者

ws.merge_cells(start_row=2, start_column=6, end_row=3, end_column=8)

# 可以不使用VBA，直接更改单元格的样式

from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
font = Font(name='Calibri',
                 size=11,
                 bold=False,
                 italic=False,
                 vertAlign=None,
                 underline='none',
                 strike=False,
                 color='FF000000')
fill = PatternFill(fill_type=None,
                 start_color='FFFFFFFF',
                 end_color='FF000000')
border = Border(left=Side(border_style=None,
                           color='FF000000'),
                 right=Side(border_style=None,
                            color='FF000000'),
                 top=Side(border_style=None,
                          color='FF000000'),
                 bottom=Side(border_style=None,
                             color='FF000000'),
                 diagonal=Side(border_style=None,
                               color='FF000000'),
                 diagonal_direction=0,
                 outline=Side(border_style=None,
                              color='FF000000'),
                 vertical=Side(border_style=None,
                               color='FF000000'),
                 horizontal=Side(border_style=None,
                                color='FF000000')
                )
alignment=Alignment(horizontal='general',
                     vertical='bottom',
                     text_rotation=0,
                     wrap_text=False,
                     shrink_to_fit=False,
                     indent=0)
number_format = 'General'
protection = Protection(locked=True,
                         hidden=False)

# 设置单元格格式

side = Side(border_style='thin', color="fff0f0f0")
border = Border(
    left=excel_cell.border.left,
    right=excel_cell.border.right,
    top=excel_cell.border.top,
    bottom=excel_cell.border.bottom
)

border.left = side
border.right = side
border.top = side
border.bottom = side
wb['A1'].border = border
```

暂时先写这么多

# XLRD,XLWT,XLUTLS

•xlrd － 读取 Excel 文件

•xlwt － 写入 Excel 文件

•xlutils － 操作 Excel 文件的实用工具，如复制、分割、筛选等

上述三个一般配合使用。只能处理.XLS文件，处理.xlsx文件会报文件错误,这里不详细讲解

