import copy
import logging
from typing import List, Callable
from docx.text.paragraph import Paragraph
import re

from md2paper.v2 import backend


class PaperPartHandler():
    def __init__(self, block: backend.Block, functions: List[Callable]) -> None:
        self.block = block
        self.functions = functions

    """
    apply all functions to one backend.Block or subclass of backend.BaseContent
    """

    def apply_functions(self, boc):
        for f in self.functions:
            f(boc)

    def handle(self):
        self.handle_block(self.block)

    def handle_block(self, block: backend.Block):
        self.apply_functions(block)

        for content in self.block.get_content_list():
            self.apply_functions(content)

        for blk in block.sub_blocks:
            self.handle_block(blk)


class BasePreprocessor():
    MATCH_ANY = '.*'

    def __init__(self, root_block: backend.Block) -> None:
        self.root_block = root_block
        self.parts: List[str] = []

        self.handlers: List[PaperPartHandler] = []

        # 如果parts之一是*，代表任意多个level1 block
        # 如果part中含*，如“附录* 附录标题”，代表以正则表达式匹配的-
        #   -任意多个以附录开头的lv1 block
        # 否则对part名进行完整匹配

        pass

    """
    initialize_template returns the exact paragraph
    where block render will begin.
    May return None, in which case render will begin
    at the last paragraph.
    """

    def initialize_template(self) -> Paragraph:
        return None

    """
    preprocess 将原始block中数据与预定义的模板，如论文或英文文献翻译进行比对，
    检查缺少的内容，同时读取填充metadata用于在initialize_template的时候填充到
    文档头(如果需要）
    """

    def __compare_parts(self, incoming: List[str]):
        i = 0
        parts = copy.deepcopy(self.parts)
        while len(parts) != 0:
            if i >= len(incoming):
                if len(parts) != 0:
                    logging.warning('preprocess: unmatched parts:', parts)
                return
            offset = 1
            part = parts[0]
            if part == self.MATCH_ANY and len(parts) >= 2:
                part = parts[1]
                offset = 2
            while i < len(incoming) and not re.match(f"^{part}$", incoming[i]):
                if offset != 2:
                    logging.warning(
                        "preprocess: unexpected part {}".format(incoming[i]))
                i = i + 1
                if i == len(incoming):
                    logging.warning("preprocess: unmatched parts:", parts)
                    return
            parts = parts[offset:]
            i = i + 1

    @classmethod
    def register_label(cls, alt_name: str, content: backend.BaseContent, index: int):
        pass

    @classmethod
    def register_ref(cls, alt_name: str, content: backend.Text):
        pass

    def register_handler(self, block: backend.Block, functions: List[Callable]):
        self.handlers.append(PaperPartHandler(block, functions))

    def if_match_register_handler(self, block: backend.Block, title: str, functions: List[Callable]) -> bool:
        if block.title_match(title):
            self.register_handler(block, functions)

    def config_preprocess(self):
        pass

    def do_preprocess(self):
        for part in self.handlers:
            part.handle()

    def preprocess(self):
        self.config_preprocess()
        self.do_preprocess()
