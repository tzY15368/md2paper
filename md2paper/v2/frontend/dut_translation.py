from md2paper.v2 import backend
from docx.text.paragraph import Paragraph
from typing import List, Dict, Tuple
from .metadata import BaseMetadata
from .preprocessor import BasePreprocessor


class DUTTranslationMetaData(BaseMetadata):
    def __init__(self) -> None:
        super().__init__()
        self.school: str = None
        self.major: str = None
        self.name: str = None
        self.number: str = None
        self.teacher: str = None
        self.__finish_date: str = None
        self.title_zh_CN: str = None
        self.title_en: str = None

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

    def get_line_mapping(self) -> Dict[str, Tuple[str, str]]:
        data = {
            "学 部（院）：": (self.school, 'school'),
            "专       业：": (self.major, 'major'),
            "学 生 姓 名：": (self.name, 'name'),
            "学       号：": (self.number, 'number'),
            "指 导 教 师：": (self.teacher, 'teacher'),
            "完 成 日 期：": (self.finish_date, 'finish_date')
        }
        return data


class DUTTranslationPreprocessor(BasePreprocessor):
    def __init__(self, root_block: backend.Block) -> None:
        super().__init__(root_block)

        self.parts: List[str] = [
            '摘要', '正文', '*'
        ]

    def initialize_template(self) -> Paragraph:
        meta = DUTTranslationMetaData()
        # ... FILL METADATA
        meta.render_template()

        anc = "翻译外文的中文题目（宋体、三号、加粗）"
        pos = backend.DM.get_anchor_position(anc) - 1
        for i in range(pos, len(backend.DM.get_doc().paragraphs)):
            backend.DM.delete_paragraph_by_index(pos)
        return None

    def preprocess(self):
        # compile blocks, and stuff

        # ===============================
        # check parts
        return super().preprocess()
