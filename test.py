from mdloader import *

if __name__ == "__main__":
    paper = GraduationPaper()
    paper.load_md("test/测试论文.md")

    paper.get_contents()
    paper.compile()
    paper.render("毕业设计（论文）模板-docx.docx", "test/out.docx")
