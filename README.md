# md2paper

基于markdown模板生成符合中文论文要求格式的docx文档  
generate thesis paper with certain format in docx from markdown

## 使用方法

依据 `requirements.txt` 安装 python 依赖。

参考 `example/` 编写 Markdown 文档，参考执行 `example.sh` 转换为 Word。

**注意**
本项目无法保证生成文档严格符合模板全部格式要求，
请及时检查生成的 Word 格式是否符合预期。

## 主要解决的问题

原有word模板没有正确使用word提供的编号、格式以及ref机制，且压缩保存的word文档不易版本控制。此外，latex难以生成完美符合格式要求的终产物。

## 技术路线

基于“毕业设计论文模板.docx”中内置的样式，使用python-docx库直接对文档内容进行操作，生成新文档。word线性（数组）存放文档中每个段落，md2paper使用该数组下标进行内容的增删查改操作，参考`DocManager`类以及代码中相关用法。  
富文本支持：内联&独立latex公式、图片、表格。

## 模块功能说明

### 前端 `mdloader.py`

将来自markdown的数据转为html后使用`bs4`解析，提取为中间表示

### 后端 `md2paper.py`

将中间表示填充至word模板中

## 待解决的问题

- 论文首页修改日期多一个空格（中文和其他字符中间会空半个空格，故无法得到整数长度的日期字符串）
