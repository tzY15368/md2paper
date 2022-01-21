from __future__ import annotations
from typing import Union,List
import docx

class BaseComponent():
    doc_target: docx.Document

    def __init__(self,doc_target:docx.Document) -> None:
        self.doc_target = doc_target

    def render_template(self):
        raise NotImplementedError("not implemented")

class Text():
    # 换行会被该类内部自动处理
    raw_text = ""

    def __init__(self, raw_text:str="") -> None:
        self.raw_text = raw_text

class Metadata(BaseComponent):
    school: str = None
    major: str = None
    name: str = None
    number: str = None
    teacher: str = None
    auditor: str = None
    finish_date: str = None
    title_zh_CN: str = None
    title_en: str = None

    def __fill_blank(self, blank_length:int, data:str)->str:
        """
        填充诸如 "学 生 姓 名：______________"的域
        **一个中文算两个字符
        """
        def get_data_len(data:str)->int:
            # 判断是否中文
            len = 0
            for char in data:
                if '\u4e00' <= char <= '\u9fa5':
                    len += 2
                else:
                    len += 1
            return len

        head_length = int((blank_length - get_data_len(data)) /2)
        if head_length <0:
            raise ValueError("值过长")
        content = " " * head_length + data + " " * (blank_length-get_data_len(data)-head_length)
        print(data,get_data_len(data))
        return content

    def render_template(self):
        title_mapping = {
            'zh_CN': 4,
            'en': 5,
        }
        self.doc_target.paragraphs[title_mapping['zh_CN']].runs[0].text = self.title_zh_CN
        self.doc_target.paragraphs[title_mapping['en']].runs[0].text = self.title_en

        line_mapping = {
            15: self.school,
            16: self.major,
            17: self.name,
            18: self.number,
            19: self.teacher,
            20: self.auditor,
            21: self.finish_date
        }
        BLANK_LENGTH = 23
        for line_no in line_mapping:
            if line_mapping[line_no] == None:
                continue
            print(len(self.doc_target.paragraphs[line_no].runs[-1].text))
            self.doc_target.paragraphs[line_no].runs[-1].text = self.__fill_blank(BLANK_LENGTH,line_mapping[line_no])

class Abstract():   
    keyword_zh_CN: Text = None
    keyword_en: Text = None
    text_zh_CN: Text = None
    text_en: Text = None

class Conclusion():
    text_zh_CN: Text = None

class Image():
    img_url = ""
    img_alt = ""

class Formula():
    pass

class Chapter(): #4
    pass

class Section(): #4.1
    pass

class SubSection(): # 4.1.1
    pass

class References(): #参考文献
    pass

class Appendixes(): #附录abcdefg
    pass

class ChangeRecord(): #修改记录
    pass

class Acknowledgments(): #致谢
    pass

class Block(): #content
    # 每个block是多个image，formula，text的组合，内部有序
    title:str = None
    __content_list__ = List[Union[Text,Image,Formula]]
    def __init__(self,title:str = "") -> None:
        self.title = title
    
    def add_content(self,content:Union[Text,Image,Formula]) -> Block:
        self.__content_list__.append(content)
        return self

class DUTThesisPaper():
    """
    每一章是一个chapter，
    每个chapter内标题后可以直接跟block或section，
    每个section内标题后可以直接跟block或subsection，
    每个subsection内标题后可以跟block
    """

    metadata:Metadata = None
    def setMetadata(self,data:Metadata)->None:
        self.metadata = data
    

    def toDocx():
        pass


class MD2Paper():
    def __init__(self) -> None:
        pass

if __name__ == "__main__":
    doc = docx.Document("毕业设计（论文）模板-docx.docx")
    meta = Metadata(doc_target=doc)
    meta.school = "电子信息与电气工程"
    meta.number = "201800000"
    meta.auditor = "张三"
    meta.finish_date = "1234年5月6号"
    meta.teacher = "里斯"
    meta.major = "计算机科学与技术"
    meta.title_en = "what is this world"
    meta.title_zh_CN = "这是个怎样的世界"
    meta.render_template()
    doc.save("out.docx")

