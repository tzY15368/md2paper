from typing import Dict
import datetime
import logging
from md2paper.v2 import backend


class BaseMetadata():
    school: str = None
    major: str = None
    name: str = None
    number: str = None
    teacher: str = None
    auditor: str = None
    __finish_date: str = None
    title_zh_CN: str = None
    title_en: str = None

    # 封面个人信息留空长度
    BLANK_LENGTH = 23

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

    @property
    def finish_date(self):
        if self.__finish_date:
            return self.__finish_date
        else:
            now = datetime.datetime.now()
            return now.strftime("%Y年%m月%d日")

    @finish_date.setter
    def finish_date(self, value):
        self.__finish_date = value

    def __get_data_len(self, data: str) -> int:
        # 判断是否中文, 一个中文算两个字符
        def is_zh_CN(char) -> bool:
            return '\u4e00' <= char <= '\u9fa5'
        len = 0
        prev_char_type = False
        prev_char = None
        for char in data:
            current_char_type = is_zh_CN(char)
            len += 1
            if current_char_type:
                len += 1
            if current_char_type != prev_char_type and prev_char != None:
                len += 0.5
            prev_char_type = current_char_type
            prev_char = char
        return int(len)

    def __fill_blank(self, blank_length: int, data: str) -> str:
        """
        填充诸如 "学 生 姓 名：______________"的域
        **一个中文算两个字符
        **中文和数字之间会多一个空格
        """
        head_length = int((blank_length - self.__get_data_len(data)) / 2)
        if head_length < 0:
            raise ValueError("值过长")
        content = " " * head_length + data + " " * \
            (blank_length-self.__get_data_len(data)-head_length)
        return content

    def render_template(self):
        # 首先设置header
        for section in backend.DM.get_doc().sections:
            p = section.header.paragraphs[0]
            if len(p.runs) == 0:
                continue
            p.runs[0].text = f"\t{self.title_zh_CN}\t"
            # 清空剩下的run
            for run in p.runs[1:]:
                run.text = ""

        mapping = self.get_title_mapping()
        for field in mapping:
            text = mapping[field].get('text')
            if not text:
                continue
            offset = backend.DM.get_anchor_position(field) - 1
            backend.DM.get_doc().paragraphs[offset].runs[0].text = text

            # 这里如果标题太长导致折行，则额外删去一行，以防止封面溢出到第二页
            logging.debug(
                f"metadata:text len = {self.__get_data_len(text)}, max len ={mapping[field]['max_len']}")
            if self.__get_data_len(text) >= mapping[field]['max_len']:
                backend.DM.delete_paragraph_by_index(offset + 5)

        mapping = self.get_line_mapping()
        for field in mapping:
            if not mapping[field]:
                continue
            offset = backend.DM.get_anchor_position(field) - 1
            data = self.__fill_blank(self.BLANK_LENGTH, mapping[field])
            backend.DM.get_doc().paragraphs[offset].runs[-1].text = data
