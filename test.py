from md2paper import *

if __name__ == "__main__":
    # GraduationPaper
    paper = GraduationPaper()
    paper.load_md("example/测试论文.md")
    paper.load_contents()
    paper.compile()

    paper.render("word-template/毕业设计（论文）模板-docx.docx", "example/out.docx")

    # TranslationPaper
    paper = TranslationPaper()
    paper.load_md("example/测试翻译.md")
    paper.load_contents()
    paper.compile()

    paper.render("word-template/外文翻译模板-docx.docx", "example/out_trans.docx")
