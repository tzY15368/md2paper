from __future__ import annotations
from typing import List
from md2paper.v2 import backend
from .preprocessor import *


# class PaperPart():
#     def __init__(self):
#         self.block = backend.Block()
#     pass


class Paper():
    """
    __init__ only reads the markdown-file and parses it
    """

    def __init__(self, template_Base:str, preprocessor:BasePreprocessor) -> None:
        # TODO:  是否还需要paperpart？应当作为一个有机整体在__init__时读入
        #self.paper_parts: List[PaperPart] = []
        self.preprocessor:BasePreprocessor = preprocessor

        self.__check()
        pass

    """
    __Check checks if the sequence of h1 titles of blocks match that of 
    the given preprocessor
    """
    def __check(self):
        parts = self.preprocessor.get_parts()
        j = 0
        # TODO: Check if blocks' titles match `parts``
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
        b = backend.Block()
        for part in self.paper_parts:
            b.add_sub_block(part.block)

        b.render_template(p)

        if update_toc:
            backend.DM.update_toc()
        backend.DM.save(out_path)
