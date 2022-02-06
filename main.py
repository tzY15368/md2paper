from md2paper import GraduationPaper,TranslationPaper
import argparse, logging

"""
usage: 
python md2paper.py -g paper.md -t trans.md
"""

options = {
    'grad': {
        "paper_class":GraduationPaper,
        "paper_template_path":"毕业设计（论文）模板-docx.docx"
    },
    'trans':{
        "paper_class":TranslationPaper,
        "paper_template_path":"外文翻译模板-docx.docx"
    }
}

logging.getLogger().setLevel(logging.DEBUG)

parser = argparse.ArgumentParser(description="usage: python md2paper.py -g paper.md -t trans.md")
parser.add_argument('-g','--grad', type=str, help='指定生成毕设论文的md文件名',required=False)
parser.add_argument('-t','--trans', type=str, help='指定生成英文论文翻译的md文件名',required=False)
args = vars(parser.parse_args())
if sum([1 if not args[i] else 0 for i in args])==len(args): logging.warning(parser.description)
for arg in args:
    md_fname = args[arg]
    if not md_fname: continue
    if not (len(md_fname) > 3 and md_fname[-3:] == ".md"): raise ValueError(f"invalid md filename:{md_fname}")
    logging.info(f"generating {arg} content in docx: {md_fname[:-3]}.docx")
    paper = options[arg]['paper_class']()
    paper.load_md(md_fname)
    paper.load_contents()
    paper.compile()
    paper.render(options[arg]['paper_template_path'], f"{md_fname[:-3]}.docx")
