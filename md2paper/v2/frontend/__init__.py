from __future__ import annotations
from typing import List
from md2paper.v2 import backend
from .preprocessor import *


class PaperPart():
    def __init__(self):
        self.block = backend.Block()
    pass


class Paper():
    """
    __init__ only reads the markdown-file and parses it
    """

    def __init__(self, template_Base) -> None:
        self.paper_parts: List[PaperPart] = []
        pass

    """
    render reads from `template_file_path`, 
    performs pre-processing with provided Preprocessor,
    renders content based on template styles,
    and dumps output to given path.
    """

    def render(self, preprocessor: BasePreprocessor,
               template_file_path: str, out_path: str, update_toc=False):
        backend.DM.set_doc(template_file_path)
        p = preprocessor.initialize_template()
        b = backend.Block()
        for part in self.paper_parts:
            b.add_sub_block(part.block)

        b.render_template(p)

        if update_toc:
            backend.DM.update_toc()
        backend.DM.save(out_path)
