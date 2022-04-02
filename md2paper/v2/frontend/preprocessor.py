from docx.text.paragraph import Paragraph

class BasePreprocessor():

    """
    initialize_template returns the exact paragraph
    where block render will begin.
    May return None, in which case render will begin
    at the last paragraph.
    """
    def initialize_template() -> Paragraph:
        return None