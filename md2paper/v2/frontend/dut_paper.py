import logging
import os
from md2paper.md2paper import ImageData
from md2paper.v2 import backend
from docx.text.paragraph import Paragraph
from typing import Callable, List, Dict, Tuple, Union
from .metadata import BaseMetadata
from .preprocessor import BasePreprocessor, PaperPartHandler
from docx.shared import Cm


class DUTPaperMetaData(BaseMetadata):
    def __init__(self) -> None:
        super().__init__()
        self.school: str = None
        self.major: str = None
        self.name: str = None
        self.number: str = None
        self.teacher: str = None
        self.auditor: str = None
        self.__finish_date: str = None
        self.title_zh_CN: str = None
        self.title_en: str = None

    # override
    def get_line_mapping(self) -> Dict[str, Tuple[str, str]]:
        data = {
            "学 院（系）：": (self.school, 'school'),
            "专       业：": (self.major, 'major'),
            "学 生 姓 名：": (self.name, 'name'),
            "学       号：": (self.number, 'number'),
            "指 导 教 师：": (self.teacher, 'teacher'),
            "评 阅 教 师：": (self.auditor, 'auditor'),
            "完 成 日 期：": (self.finish_date, 'finish_date')
        }
        return data

    # override
    def get_title_mapping(self) -> Dict[str, str]:
        data = {
            "大连理工大学本科毕业设计（论文）题目": {
                "text": self.title_zh_CN,
                "max_len": 38
            },
            "The Subject of Undergraduate Graduation Project (Thesis) of DUT": {
                "text": self.title_en,
                "max_len": 66
            }
        }
        return data


class DUTPaperPreprocessor(BasePreprocessor):
    def __init__(self, root_block: backend.Block) -> None:
        super().__init__(root_block)
        self.parts: List[str] = [
            "摘要", "Abstract", "引言",
            "正文", self.MATCH_ANY, "结论", "参考文献",
            "附录.*", "修改记录", "致谢"
        ]

    def initialize_template(self) -> Paragraph:
        meta = DUTPaperMetaData()
        # ... fill metadata
        # TODO: REMOVE ME
        meta.set_fields({
            "学院（系）": "abc",
            "专业": "def",
            "学生姓名": "sdfsdf",
            "学号": "234234",
            "指导教师": "nnnnnnnnnnn",
        })
        print(meta.__dict__)
        meta.render_template()

        anc = "摘    要"
        pos = backend.DM.get_anchor_position(anc, "Heading 1") - 1
        for i in range(pos, len(backend.DM.get_doc().paragraphs)):
            backend.DM.delete_paragraph_by_index(pos)
        return None

    def f_set_abstract_format(self) -> Callable:
        def set_abstract_format(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.Block):
                if boc.title in ['摘要', 'Abstract']:
                    keyword_text: backend.Text = boc.content_list[-1]
                    keyword_text.force_style = '关键词'
                    if len(keyword_text.runs) != 1:
                        logging.error(boc.title + ' 的关键词格式错误')
                    keyword_run = keyword_text.runs[0]
                    for run in keyword_text.runs:
                        run.bold = True
                    if boc.title == '摘要':
                        boc.set_title('摘    要', backend.Block.Heading_1, True)
                        if keyword_run.text.find('关键词：') != 0:
                            logging.error(boc.title + ' 的关键词要以 "关键词：" 开头')
                    else:
                        boc.set_title(
                            'Abstract', backend.Block.Heading_1, True)
                        if keyword_run.text.find('Key Words:') != 0:
                            logging.error(
                                boc.title + ' 的关键词要以 "Key Words:" 开头')
                        keyword_run.text.replace(':', '：')
                        keyword_run.text.replace(';', '；')
                        keyword_run.text = self.rbk(keyword_run.text)
                    # do more format check here
                else:
                    logging.error('错误的摘要标题: ' + boc.title)
        return set_abstract_format

    def f_set_intro_format(self) -> Callable:
        def set_intro_format(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.Block):
                boc.set_title(
                    '引    言', level=backend.Block.Heading_1, centered=True)
                # TODO: more formatting...
        return set_intro_format

    def set_ref_format(self, boc: Union[backend.BaseContent, backend.Block]):
        if isinstance(boc, backend.Block):
            boc.set_title(
                '参 考 文 献', level=backend.Block.Heading_1, centered=True)
            for content in boc.get_content_list():
                if isinstance(content, backend.Text):
                    content.force_style = "参考文献正文"
                    content.first_line_indent = Cm(0)

    def preprocess(self):
        blocks = self.root_block.sub_blocks

        index = 0
        # first pass:
        self.handler(blocks[index], [self.f_rbk_text(), self.f_get_metadata()])
        blocks.remove(blocks[index])

        self.match_then_handler(
            blocks[index], '摘要', [self.f_rbk_text(), self.f_set_abstract_format()])
        index += 1

        self.match_then_handler(
            blocks[index], 'Abstract', [self.f_rbk_text(), self.f_set_abstract_format()])
        index += 1

        if blocks[index].title_match('目录'):
            blocks.remove(blocks[index])

        self.match_then_handler(
            blocks[index], '引言', [self.f_rbk_text(), self.f_set_intro_format()])
        index += 1

        if blocks[index].title_match('正文'):
            blocks.remove(blocks[index])

        main_start = index
        cnt = 0
        while index < len(blocks):
            if (blocks[index].title == '结论'):
                break
            cnt += 1
            self.match_then_handler(
                blocks[index], '*', [
                    self.f_rbk_text(),
                    self.f_process_table(),
                    self.f_process_img(),
                    self.f_process_formula(),
                    self.register_multimedia_labels])
            index += 1
        main_end = index - 1

        self.match_then_handler(
            blocks[main_end+1], '结论', [])
        self.match_then_handler(
            blocks[main_end+2], '参考文献', [
                self.f_rbk_text(),
                self.set_ref_format,
                self.register_references
            ])

        index = main_end+3
        append_start = index
        cnt = 0
        while index < len(blocks):
            if (blocks[index].title == '修改记录'):
                break
            cnt += 1
            self.match_then_handler(
                blocks[index], '*', [])
            index += 1
        append_end = index-1
        self.match_then_handler(
            blocks[append_end+1], '修改记录', [])
        self.match_then_handler(
            blocks[append_end+2], '致谢', [])

        # secend pass:

        for i in range(main_start, main_end+1):
            self.handler(blocks[i], [self.register_references,
                                     self.replace_references_text])
        for i in range(append_start, append_end+1):
            self.handler(blocks[i], [self.register_references,
                                     self.replace_references_text])
        self.handler(blocks[main_end+2],
                     [self.filt_references_part])  # TODO filt
