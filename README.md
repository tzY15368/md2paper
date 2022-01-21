# md2docx

generate thesis paper with certain format in docx from markdown

主要解决的问题：无法版本控制，原有格式写起来剧痛

暂定采用markdown，其中公式使用latex语法，在文档生成时转word格式（how？待调研）
插图可解决（在毕设论文模板中不存在环绕文本，必定占据一个word paragraph）

## modules

### 提取数据

来自markdown或其他结构化数据源

### 套格式

格式直接来自毕设模板对应位置的paragraph.style

- 选项1：代码中硬编码序列化后的python对象(pickle)

- 选项2：用的时候必须提供template.docx

## roadmap

- 针对大工模板做大量硬编码（不可避免，考虑到各种cover，承诺书，xxx）优先解决问题

- generic how?
