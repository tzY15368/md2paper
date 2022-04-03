from __future__ import annotations
import logging
from typing import List, Type
import markdown
import bs4
from md2paper.v2.backend.docx_render import *
from .preprocessor import *
from .mdext import MDExt


class Paper():
    """
    __init__ only reads the markdown-file and parses it
    """

    def __init__(self, md_file_path: str, preprocessor: Type[BasePreprocessor]) -> None:
        self.block: Block = Block()
        self.preprocessor: BasePreprocessor = preprocessor

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
            elif name == 'p':
                content_list = self.__get_contents(cur)
                last_block.add_content(*content_list)
            elif name == 'table':
                table = self.__get_table(cur)
                last_block.add_content(table)
            elif name == 'math':
                math = self.__get_math(cur)
                last_block.add_content(math)
            elif name == 'ol':
                ol = self.__get_ordered_list(cur)
                last_block.add_content(ol)
            else:
                logging.error("这是啥？" + str(cur))

        # 检查md内容是否满足preprocessor给出的要求
        # self.__preprocess()

    def __get_contents(self, cur):
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
                    text.add_run(Run(i.text), Run.Bold)
            elif name == "em":
                text.add_run(Run(i.text), Run.Italics)
            elif name == "math-inline":
                text.add_run(Run(i.text), Run.Formula, transform_required=True)
            elif name == "ref":
                text.add_run(Run(i.text), Run.Reference)
            else:  # 需要分段
                if not text.empty():
                    content_list.append(text)
                    text = Text()
                if name == "br":  # 分段
                    pass
                elif name == "img":  # 图片
                    content_list.append(self.__get_image(i))
                elif i.name == "ol":
                    content_list.append(self.__get_ordered_list(i))
                else:
                    logging.error("缺了什么？" + str(i))
        if not text.empty():
            content_list.append(text)
        return content_list

    def __get_table(self, cur):
        logging.warning('暂时未实现')
        return None

    def __get_math(self, cur):
        logging.warning('暂时未实现')
        return None

    def __get_ordered_list(self, cur):
        logging.warning('暂时未实现')
        return None

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
        DM.set_doc(template_file_path)
        p = self.preprocessor.initialize_template()
        self.block.render_template(p)

        if update_toc:
            DM.update_toc()
        DM.save(out_path)
