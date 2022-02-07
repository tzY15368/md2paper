from md2paper import GraduationPaper,TranslationPaper
from md2paper.md2paper import SRC_ROOT
import argparse, logging
import os

"""
usage: 
python main.py [-g <paper.md>] [-t <trans.md>]
"""

options = {
    'grad': {
        "paper_class":GraduationPaper,
        "paper_template_path": os.path.join(SRC_ROOT,"word-template", "毕业设计（论文）模板-docx.docx")
    },
    'trans':{
        "paper_class":TranslationPaper,
        "paper_template_path": os.path.join(SRC_ROOT,"word-template", "外文翻译模板-docx.docx")
    }
}

logging_options = {
    'debug':logging.DEBUG,
    'info':logging.INFO,
    'warning':logging.WARNING
}

parser = argparse.ArgumentParser(description="usage: python main.py [-g <paper.md>] [-t <trans.md>]")
parser.add_argument('-g','--grad', type=str, help='指定生成毕设论文的md文件名',required=False)
parser.add_argument('-t','--trans', type=str, help='指定生成英文论文翻译的md文件名',required=False)
parser.add_argument('-l','--level',type=str,choices=['info','debug','warning'],required=False,help='指定logging level')
args = vars(parser.parse_args())
if sum([1 if not args[i] else 0 for i in args])==len(args): logging.warning(parser.description)

if args['level'] != None:
    logging.getLogger().setLevel(level=logging_options[args['level']])
    args.pop('level')
else:
    logging.getLogger().setLevel(logging.INFO)

for arg in args:
    md_fname = args[arg]
    if not md_fname: continue
    if not (len(md_fname) > 3 and md_fname[-3:] == ".md"): raise ValueError(f"invalid md filename:{md_fname}")
    logging.info(f"generating {arg} content in docx: {os.path.join(os.getcwd(),md_fname[:-3])}.docx")
    paper = options[arg]['paper_class']()
    paper.load_md(md_fname)
    paper.load_contents()
    paper.compile()
    paper.render(options[arg]['paper_template_path'], f"{md_fname[:-3]}.docx")
