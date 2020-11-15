# OOXML

## 简介

`OOXML`, 全称 Office Open XML. 是一种基于 xml 的格式标准. 该格式主要用于 office 文档,包括文字处理,电子表格,演示文稿以及图标,形状等其他图形材料.

该标准由微软开发,并且被 ECMA 采用, ECMA-376. 目前 OOXML 是微软文档的默认格式(.docx, .xlsx, .pptx).

## 意义

`OOXML` 是一种开放的文档标准，任何个人和组织都可以基于此开发，微软的Office套件，WPS的套件等生成的文件可以互相兼容。也使用JAVA、Go等也可以操作这些文档，生成Office文件，套用模板文件，清理修改痕迹等操作.

## 标准概述

`OOXML` 主要由两部标准组成:

1. markup 标准
2. 文件打包标准

### markup标准

主要有三种的文档类型:

- WordprocesingML (docx)
- SpreadsheetML(xlsx)
- PresentationML(pptx)

### 文件打包标准

将多个 xml 文件使用 zip 打包成一个文件. xml 文件主要的 encoding 为 UTF-8 或者 UTF-16.