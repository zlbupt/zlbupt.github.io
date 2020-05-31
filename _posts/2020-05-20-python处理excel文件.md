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

对已经打开的excel进行操作，如果excel没有打开，则打开该excel。底层使用微软的VBA进行处理，所以在excel中可以看到其对excel进行的操作。其中大部分功能依然要靠VBA宏进行处理。

# openpyxl

处理未被打开的excel，功能很齐全。把已经打开的excel视为只读文件，不能对此进行操作

# xlwriter

与openpyxl类似，只能处理未被打开的excel，相当于不能可视化。

# xls

