# md2paper

基于markdown模板生成符合中文论文要求格式的docx文档  
generate thesis paper with certain format in docx from markdown

## 使用方法

依据 `requirements.txt` 安装 python 依赖。

参考 `example/*.md` 编写符合扩展语法的 Markdown 文档，参考执行 `example.sh` 转换为 Word。

md2paper使用的静态资源（`word-template/*`, `md2paper/mml2omml.xsl`）理论上都支持**任意**执行路径，事实上可以在任意path执行 `xx/xx/md2paper/main.py [CMDLINE ARGUMENTS]`，最终产物会保存至`cwd`

项目根目录下`/libs`文件夹，用于支持实验性的wasm静态页面【WIP】

**注意**
本项目无法保证生成文档严格符合模板全部格式要求，比如两个字符的缩进事实上是用0.82cm实现的。请*务必*及时检查生成的 Word 格式是否符合预期。

### 编码格式问题

我们推荐使用`UTF-8`的OS使用该工具（如：Ubuntu, Debian, WSL, MacOS...）
对于windows，建议在系统设置中启用“实验性的utf8编码”或直接使用WSL以绕开此问题。

**警告** 启用windows的utf8编码会对你的系统中已有的部分软件带来乱码的问题。

## 主要解决的问题

考虑到以下问题：

- 原有word模板没有正确使用word提供的编号、格式以及ref机制
- 难以不出差错地写出符合格式要求的内容
- 压缩保存的word文档不易版本控制
- latex难以生成完美符合格式要求的终产物，也不具备与docx格式之间的互可操作性

创建一种包含格式context的中间表示，并使用脚本把实际内容渲染到docx文档中是一种可能的缓解难以避免的格式问题的办法。

## 技术路线

基于“毕业设计论文模板.docx”中内置的样式，使用python-docx库直接对文档内容进行操作，生成新文档。由于docx格式事实上线性地存放文档中每个段落，md2paper使用该数组下标进行内容的增删查改操作，参考`DocManager`类以及`md2paper.py`代码中相关用法。

### 可能更合适的方案

完善原有模板后完全使用模板的格式，不再螺蛳壳里做道场（在原有模板中强行插入、删除内容很大程度上造成了现有代码的混乱）

完全可以在读取模板样式表后抹除全部内容，从头到尾自行渲染，代码可以简洁很多。

### 富文本支持

- 内联&独立latex公式

- 独立图片

- 独立表格

## 核心模块功能说明

### 前端 `md_paper.py`

将来自markdown的数据转为html后使用`bs4`解析，提取为中间表示

### 后端 `md2paper.py`

将中间表示填充至word模板（论文、英文翻译）中

## 待解决的问题

- 论文首页修改日期多一个空格（中文和其他字符中间会空半个空格，故无法得到整数长度的日期字符串）

- 不支持bullet point / (un)ordered lists (上游依赖python-docx不支持)

- 缺乏文档
