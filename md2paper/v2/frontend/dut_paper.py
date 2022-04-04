from asyncio.log import logger
import logging
from md2paper.v2 import backend
from docx.text.paragraph import Paragraph
from typing import List, Dict, Union
from .metadata import BaseMetadata
from .preprocessor import BasePreprocessor, PaperPartHandler


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
    def get_line_mapping(self) -> Dict[str, str]:
        data = {
            "学 院（系）：": self.school,
            "专       业：": self.major,
            "学 生 姓 名：": self.name,
            "学       号：": self.number,
            "指 导 教 师：": self.teacher,
            "评 阅 教 师：": self.auditor,
            "完 成 日 期：": self.finish_date
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
        meta.render_template()

        anc = "摘    要"
        pos = backend.DM.get_anchor_position(anc, "Heading 1") - 1
        for i in range(pos, len(backend.DM.get_doc().paragraphs)):
            backend.DM.delete_paragraph_by_index(pos)
        return None

    def f_set_abstract_format(self):
        def set_abstract_format(boc: Union[backend.BaseContent, backend.Block]):
            if isinstance(boc, backend.Block):
                if boc.title in ['摘要', 'Abstract']:
                    keyword_text: backend.Text = boc.content_list[-1]
                    keyword_text.force_style = '关键词'
                    if len(keyword_text.runs) != 1:
                        logging.error(boc.title + ' 的关键词格式错误')
                    keyword_run = keyword_text.runs[0]
                    if boc.title == '摘要':
                        if keyword_run.text.find('关键词：') != 0:
                            logging.error(boc.title + ' 的关键词要以 "关键词：" 开头')
                    else:
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

    def preprocess(self):
        blocks = self.root_block.sub_blocks

        # first pass:
        self.match_then_handler(
            blocks[0], '*', [self.f_rbk_text(), self.f_get_metadata()])
        self.match_then_handler(
            blocks[1], '摘要', [self.f_rbk_text(), self.f_set_abstract_format()])
        self.match_then_handler(
            blocks[2], 'Abstract', [self.f_rbk_text(), self.f_set_abstract_format()])
        self.match_then_handler(
            blocks[3], '目录', [])
        self.match_then_handler(
            blocks[4], '引言', [])
        self.match_then_handler(
            blocks[5], '正文', [])

        index = 6
        main_start = index
        cnt = 0
        while index < len(blocks):
            if (blocks[index].title == '结论'):
                break
            cnt += 1
            self.match_then_handler(
                blocks[index], '*', [])
            index += 1
        main_end = index - 1

        self.match_then_handler(
            blocks[main_end+1], '结论', [])
        self.match_then_handler(
            blocks[main_end+2], '参考文献', [])

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
            blocks[append_end+1], '致谢', [])
        self.match_then_handler(
            blocks[append_end+2], '致谢', [])

        # secend pass:

        for i in range(main_start, main_end+1):
            self.handler(blocks[i], [])
