
import logging
import os
import sys
sys.path.append('..')
from md2paper import *
logging.getLogger().setLevel(logging.INFO)

if __name__ == "__main__":
    # GraduationPaper
    os.chdir("..")
    paper = GraduationPaper()
    paper.load_md("example/论文.md")
    paper.load_contents()
    paper.compile()

    paper.render("word-template/毕业设计（论文）模板-docx.docx",
                 "example/论文.docx", update_toc=False)

    # TranslationPaper
    paper = TranslationPaper()
    paper.load_md("example/外文翻译.md")
    paper.load_contents()
    paper.compile()

    paper.render("word-template/外文翻译模板-docx.docx",
                 "example/外文翻译.docx", update_toc=False)
