# md2docx

基于markdown模板生成符合中文论文要求格式的docx文档  
generate thesis paper with certain format in docx from markdown

## 主要解决的问题

原有word模板没有正确使用word提供的编号、格式以及ref机制，且压缩保存的word文档不易版本控制。此外，latex难以生成完美符合格式要求的终产物。

## 技术路线

基于“毕业设计论文模板.docx”中内置的样式，使用python-docx库直接对文档内容进行操作，生成新文档。  
富文本支持：内联&独立latex公式、图片、表格。

## 模块功能说明

### 前端 `mdloader.py`

将来自markdown的数据转为html后使用`bs4`解析，提取为中间表示

### 后端 `md2paper.py`

将中间表示填充至word模板中

## 待解决的问题

- 支持英文论文翻译
