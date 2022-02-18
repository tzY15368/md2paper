from typing import Dict
from md2paper.md2paper import *
from md2paper.dut_paper import MainContent, Metadata
import logging
"""
外文翻译要求：
1．毕业设计（论文）外文翻译的译文不得少于5千汉字。选取的内容如超出要求字数上限，应翻译至原文完整的段落结束为宜。（国际班、外语专业学生翻译8000印刷字符的专业外文文献或写出10000字符的外文文献的中文读书报告）
2．译文内容必须与毕业设计（论文）题目（或专业内容）有关，且是正式出版日期为近5年内的外文期刊，由指导教师在下达任务书时指定。
3．外文原文、译文应用标准A4纸，双面打字成文。
4．译文的基本结构与外文结构相同，页边距：上3.5cm，下2.5cm，左2.5cm、右2.5cm；页眉：2.5cm，页眉：译文的中文题目，页脚：2cm。文中标题为宋体，小四号，字体加粗。
5．原文中的图、表等的名称必须翻译，参考文献内容不翻译。
6．外文翻译装订顺序：
（1）封面；
（2）外文原文；
（3）中文译文。
7．特殊情况应在译文后附件说明。 
"""


class TranslationMetadata(Metadata):

    def get_title_mapping(self) -> Dict[str, str]:
        data = {
            "外文的中文题目": {
                "text": self.title_zh_CN,
                "max_len": 38
            },
            "The title of foreign language": {
                "text": self.title_en,
                "max_len": 66
            },
        }
        return data

    def get_line_mapping(self) -> Dict[str, str]:
        data = {
            "学 部（院）：": self.school,
            "专       业：": self.major,
            "学 生 姓 名：": self.name,
            "学       号：": self.number,
            "指 导 教 师：": self.teacher,
            "完 成 日 期：": self.finish_date
        }
        return data


class TranslationAbstract(Component):
    def __init__(self, title_zh_CN: str, author_en: str, work_place: str) -> None:
        super().__init__()
        self.author_en = author_en
        self.work_place = work_place
        self.title_zh_CN = title_zh_CN
        self.keywords: List[str] = []

    def add_keywords(self, keywords: List[str]):
        self.keywords += keywords

    def render_template(self) -> int:
        new_offset = DM.get_anchor_position(anchor_text="翻译外文的中文题目") - 1
        DM.get_paragraph(new_offset).runs[0].text = self.title_zh_CN
        for run in DM.get_paragraph(new_offset).runs[1:]:
            run.text = ""
        new_offset += 1

        DM.get_paragraph(new_offset).runs[0].text = self.author_en
        for run in DM.get_paragraph(new_offset).runs[1:]:
            run.text = ""
        new_offset += 1

        DM.get_paragraph(new_offset).runs[0].text = self.work_place
        for run in DM.get_paragraph(new_offset).runs[1:]:
            run.text = ""
        new_offset += 1

        abstract_start = "摘要："
        DM.get_paragraph(new_offset).runs[0].text = abstract_start
        new_offset += 1

        incr_kw = "关键词：(黑体、小四、加粗)"
        new_offset = super().render_template(
            anchor_text=abstract_start, incr_kw=incr_kw, incr_next=0)
        DM.get_paragraph(
            new_offset).runs[0].text = f"关键词：{'；'.join(self.keywords)}"
        new_offset += 1

        # hack: 最后给正文预留锚点
        anchor_text = "1  正文格式说明"
        p = DM.get_paragraph(new_offset).insert_paragraph_before()
        p.text = anchor_text
        new_offset += 1
        p = DM.get_doc().add_paragraph()
        p = DM.get_doc().add_paragraph()
        p = DM.get_doc().add_paragraph()
        p.text = "结    论（设计类为设计总结"

        new_offset += 4
        # 后面删完
        while new_offset != len(DM.get_doc().paragraphs):
            DM.delete_paragraph_by_index(new_offset)
        return new_offset


class TranslationMainContent(MainContent):
    def render_template(self) -> int:
        new_offset = super().render_template()
        while new_offset != len(DM.get_doc().paragraphs):
            DM.delete_paragraph_by_index(new_offset)
        return new_offset
