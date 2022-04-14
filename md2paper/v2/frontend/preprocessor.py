import collections
import copy
import logging
from typing import List, Callable, Union, Dict
from docx.text.paragraph import Paragraph
import re

from md2paper.v2 import backend


# 处理文本


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
        if len(self.functions):
            self.handle_block(self.block)

    def handle_block(self, block: backend.Block):
        self.apply_functions(block)

        for content in block.get_content_list():
            self.apply_functions(content)

        for blk in block.sub_blocks:
            self.handle_block(blk)


class BasePreprocessor():
    MATCH_ANY = '.*'

    def __init__(self, root_block: backend.Block) -> None:
        self.root_block = root_block
        self.parts: List[str] = []

        self.handlers: List[PaperPartHandler] = []

        self.metadata: Dict[str, str] = {}
        self.reference_map: collections.OrderedDict[str,
                                                    backend.BaseContent] = collections.OrderedDict()

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

    def rbk(self, text: str):  # remove_blank
        # 删除换行符
        text = text.replace("\n", " ")
        text = text.replace("\r", "")
        text = text.strip(' ')

        cn_char = u'[\u4e00-\u9fa5。，：《》、（）“”‘’\u00a0]'
        # 中文字符后空格
        should_replace_list = re.compile(
            cn_char + u' +').findall(text)
        # 中文字符前空格
        should_replace_list += re.compile(
            u' +' + cn_char).findall(text)
        # 删除空格
        for i in should_replace_list:
            if i == u' ':
                continue
            new_i = i.strip(" ")
            text = text.replace(i, new_i)
        text = text.replace("\u00a0", " ")  # 替换为普通空格
        return text

    def f_rbk_text(self):
        def rbk_text(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.Text):
                for run in boc.runs:
                    run.text = self.rbk(run.text)
        return rbk_text

    def f_get_metadata(self):
        def get_metadata(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.Table):
                for row in boc.table[1:]:
                    self.metadata[row.row[0].raw_text()
                                  ] = row.row[1].raw_text()
            elif isinstance(boc, backend.Block) and boc.level == 1:
                self.metadata['title_zh_CN'] = self.rbk(boc.title)
                self.metadata['title_en'] = self.rbk(boc.sub_blocks[0].title)
        return get_metadata

    def handler(self, block: backend.Block, functions: List[Callable]):
        pph = PaperPartHandler(block, functions)
        pph.handle()

    def match_then_handler(self, block: backend.Block, title: str, functions: List[Callable]) -> bool:
        if block.title_match(title):
            self.handler(block, functions)
        else:
            logging.warning(title + " 匹配失败")

    def preprocess(self):
        pass
