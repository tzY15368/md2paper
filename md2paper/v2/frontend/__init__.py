from __future__ import annotations
from functools import reduce
import logging
from typing import List, Type
import markdown
import bs4
from md2paper.v2.backend.docx_render import *
from .preprocessor import *
from .mdext import MDExt
from .dut_paper import DUTPaperPreprocessor
from .dut_translation import DUTTranslationPreprocessor


class Paper():
    """
    __init__ only reads the markdown-file and parses it
    """

    def __init__(self, md_file_path: str, preprocessor: Type[BasePreprocessor]) -> None:
        self.block: Block = Block()
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

        block_stack = [self.block]
        last_block = block_stack[-1]

        cur = soup.findChild()
        for cur in soup.children:
            logging.debug(cur)
            name = cur.name
            if name == None:
                pass
            elif name[0] == 'h':  # h1 h2 h3
                level = int(name[1])
                # new block
                while len(block_stack)-1 >= level:
                    block_stack.pop()
                last_block = Block()
                block_stack[-1].add_sub_block(last_block)
                block_stack.append(last_block)
                # get content
                last_block.set_title(cur.text, level)
            else:
                content_list = self.__get_contents(cur)
                last_block.add_content(*content_list)

        # 检查md内容是否满足preprocessor给出的要求
        self.preprocessor.preprocess()

    def __get_contents(self, cur):
        name = cur.name
        if name == None:
            return []
        elif name == 'p':
            return self.__get_super_texts(cur)
        elif name == 'table':
            return [self.__get_table(cur)]
        elif name == 'math':
            return [self.__get_math(cur)]
        elif name == 'ol':
            return [self.__get_ordered_list(cur)]
        elif name == 'ul':
            logging.info("暂不支持无序列表")
            return []
        else:
            logging.error("暂不支持，请反馈" + str(cur))
            return []

    def __get_super_texts(self, cur):
        content_list = []
        text = Text()
        for i in cur.children:
            name = i.name
            if name == None:
                if not hasattr(i, "text"):
                    setattr(i, "text", str(i))
                if i.text == "\n":
                    continue
                text.add_run(Run(i.text, style=Run.Normal))
            elif name == "strong":
                if not len(i.contents) == 1:
                    print("只允许粗斜体，不允许复杂嵌套")
                if i.contents[0].name == 'em':
                    text.add_run(Run(i.text), Run.Bold | Run.Italics)
                else:
                    text.add_run(Run(i.text, Run.Bold))
            elif name == "em":
                text.add_run(Run(i.text, Run.Italics))
            elif name == "math-inline":
                text.add_run(Run(i.text, Run.Formula))
            elif name == "ref":
                text.add_run(Run(i.text, Run.Reference))
            else:  # 需要分段
                if not text.empty():
                    content_list.append(text)
                    text = Text()
                if name == "br":  # 分段
                    pass
                elif name == "img":  # 图片
                    content_list.append(self.__get_image(i))
                elif name == "ol":
                    content_list.append(self.__get_ordered_list(i))
                elif name == 'code':
                    content_list.append(self.__get_code(cur))
                else:
                    logging.error("缺了什么？" + str(i))
        if not text.empty():
            content_list.append(text)
        return content_list

    def __get_table(self, table):
        def get_table_row_item(t_):
            content_list = self.__get_super_texts(t_)
            if len(content_list) == 0:
                return None
            elif len(content_list) == 1:
                return content_list[0]
            else:
                logging.error("表格元素中不能换行")
                return None

        # 表头
        row = [get_table_row_item(th)
               for th in table.find("thead").find_all("th")]
        row_list = [Row(row, False)]

        # 表身
        for tr in table.find("tbody").find_all("tr"):
            row = [get_table_row_item(td)for td in tr.find_all("td")]
            row_list.append(Row(row, False))

        return Table(None, row_list)  # with no title

    def __get_math(self, math):
        return Formula(None, math.text)  # with no title

    def __get_image(self, img):
        return Image(img["alt"], img["src"])

    def __get_ordered_list(self, ol):
        def get_list_item(li):
            if (li.contents[0].text == "\n"):  # <p>
                content_list_list = [self.__get_contents(cur)
                                     for cur in li.contents]
                content_list = reduce(
                    lambda x, y: x+y, content_list_list)  # flatten
            else:  # text
                content_list = self.__get_super_texts(li)
            return ListItem(content_list)

        ordered_list = [get_list_item(li)
                        for li in ol.find_all("li", recursive=False)]
        return OrderedList(ordered_list)

    def __get_code(self, code):
        codes: str = code.text
        first_line = codes.splitlines()[0]
        return Code(language=first_line.strip(), txt=codes[len(first_line):])

    """
    render reads from `template_file_path`, 
    performs pre-processing with provided Preprocessor,
    renders content based on template styles,
    and dumps output to given path.
    """

    def render(self, template_file_path: str, out_path: str, update_toc=False):
        DM.set_doc(template_file_path)
        p = self.preprocessor.initialize_template()
        self.block.render_template(p)

        if update_toc:
            DM.update_toc()
        DM.save(out_path)
