import collections
import copy
import logging
from typing import List, Callable, Tuple, Type, Union, Dict
from docx.text.paragraph import Paragraph
from docx.shared import Cm
import re
from functools import reduce


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
                                                            Union[backend.Image,
                                                                  backend.Table,
                                                                  backend.Formula]] = collections.OrderedDict()
        self.reference_map: collections.OrderedDict[str,
                                                    int] = collections.OrderedDict()
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
        if not isinstance(boc, backend.Text):
            return
        for run in boc.runs:
            if not run.reference:
                continue
            ali = run.text
            if ali.find(",") == -1:
                if ali not in self.content_reference_map:
                    if ali not in self.reference_map:
                        self.reference_map[ali] = len(self.reference_map) + 1
            else:
                alis = ali.split(",")
                for ali in alis:
                    if ali not in self.reference_map:
                        self.reference_map[ali] = len(self.reference_map) + 1

    def register_multimedia_labels(self, boc: Union[backend.BaseContent, backend.Block]):
        # parse表名、公式名、图名
        if not (isinstance(boc, backend.Block) and boc.level == backend.Block.Heading_1):
            return

        content_count: Dict[Type, int] = {
            backend.Image: 0,
            backend.Table: 0,
            backend.Formula: 0
        }
        block_id = "未支持"
        if boc.title and boc.title[0].isdigit():
            block_id = boc.title[0]
        # TODO: 添加附录编号支持

        def make_index(id: int):
            if block_id[0].isdigit():
                return "{}.{}".format(block_id, id)
            else:
                return "{}{}".format(block_id, id)

        content_all = boc.get_content_list(recursive=True)

        for i, content in enumerate(content_all):
            if isinstance(content, backend.Image) or isinstance(content, backend.Table) or isinstance(content, backend.Formula):
                if isinstance(content, backend.Image):
                    if not content.image:
                        raise RuntimeError(
                            "register labels MUST happen after f_process_image")
                    base = ""
                    content_count[backend.Image] += 1
                    content.image.img_alt = "图{}  {}".format(
                        make_index(content_count[backend.Image]), content.image.img_alt)
                    content.refname = "图" + \
                        make_index(content_count[backend.Image])
                elif isinstance(content, backend.Table):
                    content_count[backend.Table] += 1
                    content.title = "表{}  {}".format(
                        make_index(content_count[backend.Table]), content.title)
                    content.refname = "表" + \
                        make_index(content_count[backend.Table])
                elif isinstance(content, backend.Formula):
                    content_count[backend.Formula] += 1
                    content.title = "（{}）".format(
                        make_index(content_count[backend.Formula]))
                    content.refname = "式" + \
                        make_index(content_count[backend.Formula])

                if content.alias:
                    if content.alias in self.content_reference_map and \
                            self.content_reference_map[content.alias] != content:
                        raise ValueError(
                            "duplicate ref name:{}\ntraceback: {}\nobj:".format(content.alias, content.title, content))
                    self.content_reference_map[content.alias] = content

    def replace_references_text(self, boc: Union[backend.BaseContent, backend.Block]):
        if not isinstance(boc, backend.Text):
            return
        is_text = False
        for run in boc.runs:
            if not run.reference:
                if run.text.endswith('文献'):
                    is_text = True
                else:
                    is_text = False
                continue
            ali = run.text
            if ali.find(",") == -1:
                if ali in self.content_reference_map:
                    run.text = self.content_reference_map[ali].refname
                    is_text = True
                elif ali in self.reference_map:
                    run.text = "[{}]".format(self.reference_map[ali])
                else:
                    raise ValueError('引用没有注册，请联系维护人员: ' + ali)
            else:
                alis = ali.split(",")
                index_list = []
                for ali in alis:
                    if ali not in self.reference_map:
                        raise ValueError('引用没有注册，请联系维护人员: ' + ali)
                    index_list.append(self.reference_map[ali])
                index_list.sort()
                index_list = [[x, x] for x in index_list]
                short_list = index_list[:1]
                for index_pair in index_list[1:]:
                    if short_list[-1][1]+1 == index_pair[0]:
                        short_list[-1][1] = index_pair[1]
                    else:
                        short_list.append(index_pair)
                short_list = [str(x[0]) if x[0] == x[1]
                              else "{}-{}".format(x[0], x[1])
                              for x in short_list]
                run.text = "[{}]".format(
                    reduce(lambda x, y: x+','+y, short_list))
            if is_text:
                run.reference = False
            else:
                run.reference = False
                run.superscript = True
            is_text = False

    def filt_references_part(self, boc: Union[backend.BaseContent, backend.Block]):
        if not (isinstance(boc, backend.Block) and boc.level == backend.Block.Heading_1):
            return

        def filt_references(code: backend.Code) -> List[Tuple[int, backend.Text]]:
            if code.language != 'literature':
                return []

            results: List[Tuple[int, backend.Text]] = []
            txts = code.txt.split('\n')
            for ref_item in txts:
                if not ref_item:
                    continue
                pos = ref_item.find("]")
                if not (ref_item[0] == "[" and pos != -1):
                    raise ValueError("参考文献条目应该以 `[索引]` 开头: " + ref_item)
                ali = ref_item[1: pos]
                item = ref_item[pos+1:].strip()
                if ali in self.reference_map:
                    id = self.reference_map[ali]
                    item = '[{}] {}'.format(id, item)
                    text = backend.Text(item)
                    text.first_line_indent = -backend.Text.two_chars_length
                    text.force_style = "参考文献正文"
                    results.append((id, text))

            return results

        code_list: List[backend.Code] = boc.get_content_list(
            content_type=backend.Code, recursive=True)
        tuple_list_list = [filt_references(code) for code in code_list]
        tuple_list = reduce(lambda x, y: x+y, tuple_list_list)
        tuple_list.sort(key=lambda k: k[0])

        for i, (id, text) in enumerate(tuple_list):
            if not i + 1 == id:
                logging.error('参考文献编号不递增，可能缺少或有重复')
                break
        if len(tuple_list) != len(self.reference_map):
            logging.error('参考文献可能缺少或有重复')

        boc.content_list = [x[1] for x in tuple_list]
        boc.sub_blocks.clear()

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
            if not (isinstance(boc, backend.Block) and boc.level == backend.Block.Heading_1):
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
        def process_formula(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.Block) and boc.level == backend.Block.Heading_1:
                # get formula title or alias
                all_content = boc.get_content_list(recursive=True)
                for i, content in enumerate(all_content):
                    if isinstance(content, backend.Formula):
                        if i-1 < 0 or not isinstance(all_content[i-1], backend.Text):
                            logging.warning(
                                "formula header missing, expecting text before table")
                            continue
                        alias = all_content[i-1].get_text()
                        all_content[i-1].kill()
                        content.title = alias
                        content.alias = alias
            elif isinstance(boc, backend.Formula):
                pass  # TODO: transform if need
            elif isinstance(boc, backend.Text):
                pass  # TODO: transform if need

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
            if isinstance(boc, backend.Block) and boc.level == backend.Block.Heading_1:
                # get table title
                all_content = boc.get_content_list(recursive=True)
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
