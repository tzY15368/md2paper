from md2paper.v2 import backend
from docx.text.paragraph import Paragraph
from typing import List, Dict
from .metadata import BaseMetadata
from .preprocessor import BasePreprocessor, PaperPartPreprocessor


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
        self.parts_handler = map(PaperPartPreprocessor, self.parts)

    def initialize_template(self) -> Paragraph:
        meta = DUTPaperMetaData()
        # ... fill metadata
        meta.render_template()

        anc = "摘    要"
        pos = backend.DM.get_anchor_position(anc, "Heading 1") - 1
        for i in range(pos, len(backend.DM.get_doc().paragraphs)):
            backend.DM.delete_paragraph_by_index(pos)
        return None

    def preprocess(self):
        # compile blocks, and stuff
        
        # TODO: THIS CODE WON'T WORK, parts_handler needs filtering
        for i, part in enumerate(self.parts_handler):
            part.call_methods(i)
        # 此此时已经注册了所有引用和标签，再次遍历text进行替换
        for content in self.root_block.iter_content(backend.Text):
            # 替换
            pass

        # ===============================
        # check parts
        return super().preprocess()
