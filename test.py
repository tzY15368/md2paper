from mdloader import *

if __name__ == "__main__":
    doc = docx.Document("毕业设计（论文）模板-docx.docx")
    DM.set_doc(doc)

    paper = load_md("test/测试论文.md", "论文")
    paper[0].render_template()  # metadata
    paper[1].render_template()  # 摘要 Abstract
    paper[2].render_template()  # 引言
    paper[3].render_template()  # 正文
    paper[4].render_template()  # 结论
    # paper[5].render_template()  # 参考文献
    # paper[6].render_template()  # 附录
    # paper[7].render_template()  # 修改记录
    paper[8].render_template()  # 致谢

    DM.update_toc()
    doc.save("test/out.docx")
