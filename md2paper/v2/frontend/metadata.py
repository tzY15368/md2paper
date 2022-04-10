from typing import Dict, Tuple
import datetime
import logging
from md2paper.v2 import backend


class BaseMetadata():
    __finish_date: str = None

    # 封面个人信息留空长度
    BLANK_LENGTH = 23

    # 返回的tuple[0]是field的值，[1]是field的attr名，
    def get_line_mapping(self) -> Dict[str, Tuple[str, str]]:
        raise NotImplementedError

    def get_title_mapping(self) -> Dict[str, str]:
        raise NotImplementedError

    def set_fields(self, data: Dict[str, str]):
        line_mapping = self.get_line_mapping()
        for field_cn_name in line_mapping:
            trimmed_name = field_cn_name.strip().replace(' ', '')
            if trimmed_name[-1] != "：":
                raise ValueError(
                    "metadata: unexpected field_cn_name: "+trimmed_name)
            trimmed_name = trimmed_name[:-1]
            if trimmed_name not in data:
                logging.warning(f"metadata: field {trimmed_name} went missing")
            else:
                setattr(self, line_mapping[field_cn_name]
                        [1], data[trimmed_name])

    @property
    def finish_date(self):
        if self.__finish_date:
            return self.__finish_date
        else:
            now = datetime.datetime.now()
            return now.strftime("%Y年%m月%d日".encode('unicode_escape')
                                .decode('utf8'))\
                .encode('utf-8')\
                .decode('unicode_escape')

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
            if not mapping[field][0]:
                continue
            offset = backend.DM.get_anchor_position(field) - 1
            data = self.__fill_blank(self.BLANK_LENGTH, mapping[field][0])
            backend.DM.get_doc().paragraphs[offset].runs[-1].text = data
