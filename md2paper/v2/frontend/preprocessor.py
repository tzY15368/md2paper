import collections
import copy
import logging
from typing import List, Callable, Tuple, Type, Union, Dict
from docx.text.paragraph import Paragraph
from docx.shared import Cm
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
        self.content_reference_map: collections.OrderedDict[str,
                                                            backend.BaseContent] = collections.OrderedDict()
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

    def register_references(self, boc: Union[backend.BaseContent, backend.Block]):

        pass

    def register_multimedia_labels(self, boc: Union[backend.BaseContent, backend.Block]):
        # parse表名、公式名、图名
        if isinstance(boc, backend.BaseContent) or not boc:
            #logging.debug("preprocess: unexpected basecontent type")
            return

        content_count: Dict[Type, int] = {
            backend.Image: 0,
            backend.Table: 0,
            backend.Formula: 0
        }
        block_id = ""
        if boc.title and boc.title[0].isdigit():
            block_id = boc.title[0]
        content_all = boc.get_content_list(recursive=True)

        for i, content in enumerate(content_all):
            
            if isinstance(content, backend.Image):
                if not content.image:
                    raise RuntimeError("register labels MUST happen after f_process_image")
                base = ""
                if block_id:
                    content_count[backend.Image] += 1
                    base = "图{}.{} ".format(
                        block_id if block_id else "未定义", content_count[backend.Image])
                content.image.img_alt = base + content.image.img_alt

                if content.alias:
                    if content.alias in self.content_reference_map and \
                        self.content_reference_map[content.alias] != content:
                        raise ValueError(
                            "duplicate ref name:{}\ntraceback: {}\nobj:".format(content.alias, content.title, content))
                    self.content_reference_map[content.alias] = content

            elif isinstance(content, backend.Table):

                if block_id:
                    content_count[backend.Table] += 1
                    _title = "表{}.{} {}".format(
                        block_id, content_count[backend.Table], content.title)
                content.title = _title
            elif isinstance(content, backend.Formula):
                if block_id:
                    content_count[backend.Formula] += 1
                content.title = "（{}.{}）".format(
                    block_id, content_count[backend.Formula])
            else:
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

    def f_process_img(self) -> Callable:
        # alias, title, width
        def analyze_img_title(img: backend.Image) -> Tuple[str, str, float]:

            initial_alt = img.title
            img_alt = initial_alt
            real_width = 0
            ref_name = ''
            real_alt = img_alt
            if ';' in img_alt:
                fields = str(img_alt).split(';')
                if len(fields) != 2:
                    return ("", img_alt.strip(), 0)
                img_alt = fields[0]
                width_field = fields[1].strip()
                if width_field:
                    if '%' not in width_field:
                        raise ValueError(
                            "image: invalid width:" + width_field)
                    real_width = float(width_field[:-1])/100

            if ':' in img_alt:
                fields = str(img_alt).split(':')
                ref_name = fields[0]
                real_alt = fields[1]

            return (ref_name, real_alt, real_width)

        def process_image(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.BaseContent) or not boc:
                logging.debug(
                    "unexpected {} in process_image".format(type(boc)))
                return
            content_all = boc.get_content_list(recursive=True)
            for i, content in enumerate(content_all):
                if isinstance(content, backend.Image):
                    ref_name, real_alt, real_width = analyze_img_title(
                        content)
                    content.alias = ref_name
                    img_data = backend.ImageData(content.src, alt=real_alt,
                                                 width_ratio=real_width)
                    content.set_image_data(img_data)
        return process_image

    def f_process_formula(self) -> Callable:
        def process_formula(*args):
            pass
        return process_formula

    def f_process_table(self) -> Callable:

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

        def analyze_title(s: str):
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
                all_content = boc.get_content_list()
                for i, content in enumerate(all_content):
                    if isinstance(content, backend.Table):
                        if i-1 < 0 or not isinstance(all_content[i-1], backend.Text):
                            logging.warning(
                                "table header missing, expecting text before table")
                            continue
                        alias, title, columns_width = analyze_title(
                            all_content[i-1].get_text())
                        all_content[i-1].kill()
                        content.title = title
                        content.alias = alias
                        if columns_width:
                            content.set_columns_width(columns_width)

            elif isinstance(boc, backend.Table):
                make_borders(boc.table)
                for content in boc.get_content_list():
                    if isinstance(content, backend.Text):
                        content.force_style = "图名中文"
                        content.first_line_indent = Cm(0)

        return process_table

    def handler(self, block: backend.Block, functions: List[Callable]):
        pph = PaperPartHandler(block, functions)
        pph.handle()

    def match_then_handler(self, block: backend.Block, title: str, functions: List[Callable]) -> bool:
        if block.title_match(title):
            self.handler(block, functions)
        else:
            logging.warning(title + " 匹配失败")

    """
    preprocess 将原始block中数据与预定义的模板，如论文或英文文献翻译进行比对，
    检查缺少的内容，同时读取填充metadata用于在initialize_template的时候填充到
    文档头(如果需要）
    """

    def preprocess(self):
        pass
