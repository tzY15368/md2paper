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

    def f_process_table(self):
        def analyse_title(s: str):
            # ali, title, columns_width
            # 别名: 表名; 宽度表
            # 宽度表 = 10% 20% 30% ...
            if not ':' in s:
                logging.error("错误表格标题格式：需要别名或标题")
            sp = s.split(':')
            sp = [sp[0]] + sp[1].split(';')
            if len(sp) == 2:
                ali = self.rbk(sp[0])
                title = self.rbk(sp[1])
                columns_width = []
            elif len(sp) == 3:
                ali = sp[0]
                title = self.rbk(sp[1])
                widths = sp[2].split('%')

                width_sum = 0
                columns_width = []
                for width in widths:
                    w = width.strip()
                    if w:
                        width_sum += int(w)
                        columns_width.append(w/100)
                if width_sum > 100:
                    raise ValueError('表格每列宽度和不得超过 100%: ' + s)
            else:
                raise ValueError('表的标题格式错误: ' + s)

            return ali, title, columns_width

        def is_border(row: backend.Row):
            for text in row.row:
                if text == None:
                    return False
                cnt = 0
                for i in text.raw_text():
                    if i != '-':
                        return False
                    else:
                        cnt += 1
                if cnt < 3:
                    return False
            return True

        def make_borders(rows: List[backend.Row]):
            rows[0].has_top_border = True
            rows[1].has_top_border = True
            delete_list: List[backend.Row] = []
            has_top_border = False
            for i in range(2, len(rows)):
                if is_border(rows[i]):
                    delete_list.append(rows[i])
                    has_top_border = True
                else:
                    if has_top_border:
                        rows[i].has_top_border = True
                        has_top_border = False
            for row in delete_list:
                rows.remove(row)

        def process_table(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.Block):
                # get table title
                last_content: backend.Text = None
                table_title_contents: List[backend.Text] = []
                for content in boc.get_content_list():
                    if isinstance(content, backend.Table):
                        table_title_contents.append(last_content)
                        ali, title, columns_width = analyse_title(
                            last_content.raw_text())
                        content.ali = ali
                        content.title = title
                        if columns_width:
                            content.set_columns_width(columns_width)
                    else:
                        last_content = content
                pass  # TODO: delete table_title_contents in boc
            elif isinstance(boc, backend.Table):
                make_borders(boc.table)
                pass  # TODO: register table in refs
        return process_table

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
