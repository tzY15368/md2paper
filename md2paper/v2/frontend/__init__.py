from __future__ import annotations
import logging
from typing import List, Type
import markdown
import bs4
import backend
from .preprocessor import *
from .mdext import MDExt


class Paper():
    """
    __init__ only reads the markdown-file and parses it
    """

    def __init__(self, md_file_path: str, preprocessor: Type[BasePreprocessor]) -> None:
        # TODO:  是否还需要paperpart？应当作为一个有机整体在__init__时读入
        #self.paper_parts: List[PaperPart] = []
        self.block: backend.Block = backend.Block()
        self.preprocessor: BasePreprocessor = preprocessor(self.block)
        with open(md_file_path) as f:
            data = f.read()

        md_html = markdown.markdown(
            data,
            tab_length=3,
            extensions=[
                'markdown.extensions.tables',
                MDExt()
            ]
        )
        soup = bs4.BeautifulSoup(md_html, 'html.parser')

        # 删除html注释
        for i in soup(text=lambda text: isinstance(text, bs4.Comment)):
            i.extract()

        current_blk = self.block
        for node in soup.recursiveChildGenerator():
            name = getattr(node, 'name', None)
            text = getattr(node, 'text', None)
            print(name, text)
            if not name: continue
            if len(name) ==2 and name[0]=='h':
                level = int(name[1])
            elif name == 'p':
                pass
            elif name == 'ol':
                pass
            elif name == 'li':
                pass
            elif name == 'br':
                pass
            elif name == 'img':
                pass
            elif name == 'strong':
                pass
            elif name == 'i':
                pass

        # 检查md内容是否满足preprocessor给出的要求
        self.__preprocess()
        pass

    """
    __Check checks if the sequence of h1 titles of blocks match that of 
    the given preprocessor
    """

    def __preprocess(self):
        parts = self.preprocessor.preprocess()
        j = 0
        # TODO: Check if blocks' titles match `parts``
        # ...
        if len(parts) != 0:
            logging.warning(
                f"check: part {parts[0]} or more are missing, check preprocessor for details")
        pass
    """
    render reads from `template_file_path`, 
    performs pre-processing with provided Preprocessor,
    renders content based on template styles,
    and dumps output to given path.
    """

    def render(self, template_file_path: str, out_path: str, update_toc=False):
        backend.DM.set_doc(template_file_path)
        p = self.preprocessor.initialize_template()
        self.block.render_template(p)

        if update_toc:
            backend.DM.update_toc()
        backend.DM.save(out_path)
