from typing import List
from docx.text.paragraph import Paragraph


class BasePreprocessor():

    """
    initialize_template returns the exact paragraph
    where block render will begin.
    May return None, in which case render will begin
    at the last paragraph.
    """
    def initialize_template(self) -> Paragraph:
        return None

    """
    parts returns ALL of the MANDATORY heading-
    -titles IN ORDER
    """
    def get_parts(self)->List[str]:
        return []