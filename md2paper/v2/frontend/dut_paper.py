import logging
import os
from md2paper.md2paper import ImageData
from md2paper.v2 import backend
from docx.text.paragraph import Paragraph
from typing import Callable, List, Dict, Tuple, Union
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
        # TODO: END OF TODO
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
                if boc.title != "引言":
                    logging.error('错误的引言标题: ' + boc.title)
                    return
                boc.set_title(
                    '引    言', level=backend.Block.Heading_1, centered=True)
                # TODO: more formatting...
        return set_intro_format

    def preprocess(self):
        blocks = self.root_block.sub_blocks

        # parse表名、公式名、图名
        content_all = self.root_block.get_content_list(recursive=True)
        for i, content in enumerate(content_all):
            if isinstance(content, backend.Image):
                img = content
                initial_alt = img.title
                img_alt = initial_alt
                real_width = 0
                ref_name = ''
                real_alt = img_alt
                if ';' in img_alt:
                    fields = str(img_alt).split(';')
                    if len(fields) != 2:
                        continue
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

                img_data = ImageData(img.src, alt=real_alt,
                                     width_ratio=real_width)
                img.set_image_data(img_data)

                if ref_name:
                    if ref_name in self.reference_map:
                        raise ValueError(
                            "duplicate ref name:{}\n traceback: {}".format(ref_name, initial_alt))
                    self.reference_map[ref_name] = content

            elif isinstance(content, backend.Table):
                _title = ""
                if not (i-1 > 0 and isinstance(content_all[i-1],backend.Text)):
                    logging.warning("title of table went missing, content offset at {}".format(i))
                else:
                    _title = content_all[i-1].get_text()
                    content_all[i-1].kill()
                if _title:
                    if ':' in _title:
                        fields = _title.split(':')
                        ref_name = fields[0].strip()
                        _title = fields[1].strip()
                        self.reference_map[ref_name] = content
                content.title = _title
            else:
                pass
            

        # first pass:
        self.match_then_handler(
            blocks[0], '*', [self.f_rbk_text(), self.f_get_metadata()])
        self.match_then_handler(
            blocks[1], '摘要', [self.f_rbk_text(), self.f_set_abstract_format()])
        self.match_then_handler(
            blocks[2], 'Abstract', [self.f_rbk_text(), self.f_set_abstract_format()])
        if blocks[3].title_match('目录'):
            blocks.remove(blocks[3])
        self.match_then_handler(
            blocks[3], '引言', [self.f_rbk_text(), self.f_set_intro_format()])
        if blocks[4].title_match('正文'):
            blocks.remove(blocks[4])

        index = 4
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
            blocks[append_end+1], '修改记录', [])
        self.match_then_handler(
            blocks[append_end+2], '致谢', [])

        # secend pass:

        for i in range(main_start, main_end+1):
            self.handler(blocks[i], [])
