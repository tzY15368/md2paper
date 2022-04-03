from typing import List
from docx.text.paragraph import Paragraph

import backend


class BasePreprocessor():

    def __init__(self, root_block:backend.Block) -> None:
        self.root_block = root_block
        pass

    """
    initialize_template returns the exact paragraph
    where block render will begin.
    May return None, in which case render will begin
    at the last paragraph.
    """
    def initialize_template(self) -> Paragraph:
        return None


    """
    preprocess 将原始block中数据与预定义的模板，如论文或英文文献翻译进行比对，
    检查缺少的内容，同时读取填充metadata用于在initialize_template的时候填充到
    文档头(如果需要）
    """
    def preprocess(self):
        pass


class DUTPaperPreprocessor(BasePreprocessor):
    def __init__(self, root_block: backend.Block) -> None:
        super().__init__(root_block)
        