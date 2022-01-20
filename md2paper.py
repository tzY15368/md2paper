from typing import Union,List
from __future__ import annotations
from fastapi import Form

class MD2Paper():
    def __init__(self) -> None:
        pass

class Text():
    # 换行会被该类内部自动处理
    text = ""

class Metadata():
    school: Text = None
    major: Text = None
    name: Text = None
    number: Text = None
    teacher: Text = None
    auditor: Text = None
    finish_date: Text = None
    title_zh_CN: Text = None
    title_en: Text = None

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
    pass
