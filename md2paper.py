from __future__ import annotations
from ast import Str
from curses import raw
from typing import Union,List
import docx
from docx.shared import Inches
class BaseComponent():
    doc_target: docx.Document

    def __init__(self,doc_target:docx.Document) -> None:
        self.doc_target = doc_target

    def render_template(self):
        raise NotImplementedError("not implemented")

class BaseContent():
    def to_paragraph():
        raise NotImplementedError

class Text(BaseContent):
    # 换行会被该类内部自动处理
    raw_text = ""

    def __init__(self, raw_text:str="") -> None:
        self.raw_text = raw_text

    def to_paragraph(self)->str:
        return self.raw_text

    @classmethod
    def read(cls, txt:str)->List[Text]:
        return [Text(i) for i in txt.split('\n')]

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
        # 只支持论文，不支持翻译！！
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

class Abstract(BaseComponent):   
    __keyword_zh_CN: Text = None
    __keyword_en: Text = None
    __text_zh_CN: Block = None
    __text_en: Block = None

    def __init__(self, doc_target: docx.Document) -> None:
        super().__init__(doc_target)
        self.__text_en = Block(doc_target)
        self.__text_zh_CN = Block(doc_target)

    def set_keyword(self, zh_CN:str,en:str):
        self.__keyword_en = Text(en)
        self.__keyword_zh_CN = Text(zh_CN)

    def set_text(self, zh_CN:str,en:str):
        self.__text_en.add_content(content_list=Text.read(en))
        self.__text_zh_CN.add_content(content_list=Text.read(zh_CN))

    def render_template(self, en_title='this is english title')->None:
        # 64开始是摘要正文
        abs_cn_start = 64
        abs_cn_end = self.__text_zh_CN.render_block(abs_cn_start,title=False)
        #p = self.doc_target.paragraphs[ABSTRACT_ZH_CN_START].insert_paragraph_before(text=self.text_zh_CN.to_paragraph())
        
        # https://stackoverflow.com/questions/30584681/how-to-properly-indent-with-python-docx
        #p.paragraph_format.first_line_indent = Inches(0.25)
        
        while not self.doc_target.paragraphs[abs_cn_end+2].text.startswith("关键词："):
            p = self.doc_target.paragraphs[abs_cn_end+1]._element
            p.getparent().remove(p)
            p._p = p._element = None
        
        # cn kw
        kw_cn_start = abs_cn_end + 2
        self.doc_target.paragraphs[kw_cn_start].runs[1].text = self.__keyword_zh_CN.to_paragraph()
        
        # en start

        en_title_start = kw_cn_start+4
        self.doc_target.paragraphs[en_title_start].runs[1].text = en_title

        en_abs_start = en_title_start + 3
        en_abs_end = self.__text_en.render_block(en_abs_start,title=False)-1
        #self.doc_target.paragraphs[en_abs_start].insert_paragraph_before(text=self.__text_en.to_paragraph())

        # https://stackoverflow.com/questions/61335992/how-can-i-use-python-to-delete-certain-paragraphs-in-docx-document
        while not self.doc_target.paragraphs[en_abs_end+2].text.startswith("Key Words："):
            p = self.doc_target.paragraphs[en_abs_end+1]._element
            p.getparent().remove(p)
            p._p = p._element = None

        # en kw
        kw_en_start = en_abs_end +2

        # https://github.com/python-openxml/python-docx/issues/740
        delete_num = len(self.doc_target.paragraphs[kw_en_start].runs) - 4
        for run in reversed(list(self.doc_target.paragraphs[kw_en_start].runs)):
            self.doc_target.paragraphs[kw_en_start]._p.remove(run._r)
            delete_num -= 1
            if delete_num < 1:
                break
        
        
        self.doc_target.paragraphs[kw_en_start].runs[3].text = self.__keyword_en.to_paragraph()
class Conclusion():
    text_zh_CN: Text = None

class Image(BaseContent):
    img_url = ""
    img_alt = ""

class Formula(BaseContent):
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

class Block(BaseComponent): #content
    # 每个block是多个image，formula，text的组合，内部有序
    
    def __init__(self, doc_target: docx.Document) -> None:
        super().__init__(doc_target)
        self.title:str = None
        self.__content_list:List[Union[Text,Image,Formula]] = []
    def set_title(self,title:str) -> None:
        self.title = title
    
    def add_content(self,content:Union[Text,Image,Formula]=None,
            content_list:Union[List[Text],List[Image],List[Formula]]=[]) -> Block:
        if content:
            self.__content_list.append(content)
        for i in content_list:
            self.__content_list.append(i)
        #print('added content with len',len(content_list),content_list[0].raw_text,id(self))
        return self


    # render_block returns the final paragraph's offset
    def render_block(self, offset:int, title:bool=True)->int:
        if not self.__content_list:
            return
        if not title:
            # generate necessary paragraphs
            internal_content_list = []
            p = self.doc_target.paragraphs[offset].insert_paragraph_before()
            internal_content_list.insert(0,p)
            for i in range(len(self.__content_list)-1):
                p = internal_content_list[0].insert_paragraph_before()
                internal_content_list.insert(0,p)
            
            #print('got content list',len(internal_content_list),id(self))
            assert len(self.__content_list)==len(internal_content_list)
            for i in range(len(self.__content_list)):
                #p = self.doc_target.paragraphs[offset].insert_paragraph_before()
                if type(self.__content_list[i]) == Text:
                    p = internal_content_list[i]
                    p.text = self.__content_list[i].to_paragraph()
                    p.paragraph_format.first_line_indent = Inches(0.25)
                else:
                    raise NotImplementedError
            return offset + len(self.__content_list)
        else:
            raise NotImplementedError("title unsupported")

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
    
    abstract:Abstract = None
    def setAbstract(self,abstract:Abstract)->None:
        self.abstract = abstract

    def toDocx():
        pass


class MD2Paper():
    def __init__(self) -> None:
        pass

if __name__ == "__main__":
    doc = docx.Document("毕业设计（论文）模板-docx.docx")
    # meta = Metadata(doc_target=doc)
    # meta.school = "电子信息与电气工程"
    # meta.number = "201800000"
    # meta.auditor = "张三"
    # meta.finish_date = "1234年5月6号"
    # meta.teacher = "里斯"
    # meta.major = "计算机科学与技术"
    # meta.title_en = "what is this world"
    # meta.title_zh_CN = "这是个怎样的世界"
    # meta.render_template()

    abs = Abstract(doc)
    a = """CommonMark中并未定义普通文本高亮。
CSDN中支持通过==文本高亮==，实现文本高亮

如果你使用的markdown编辑器不支持该便捷设置，可通过HTML标签<mark>实现：
语法：<mark>文本高亮<mark>
效果：文本高亮"""
    b = """代码块中，也支持出现`。与行内代码类似，代码中连续的`，不能超过代码块最外侧的`数量。例如，要打印```（3个），代码块需使用````（4个）声明"""
    
    c = """Any subsequent access to the "deleted" paragraph object will raise AttributeError, so you should be careful not to keep the reference hanging around, including as a member of a stored value of Document.paragraphs.
The reason it's not in the library yet is because the general case is much trickier, in particular needing to detect and handle the variety of linked items that can be present in a paragraph; things like a picture, a hyperlink, or chart etc.
But if you know for sure none of those are present, these few lines should get the job done."""
    
    d = """you should cite the location you copied this from, as acknowledgement of the original author's work."""

    abs.set_text(a,c)
    abs.set_keyword(b,d)
    abs.render_template()

    doc.save("out.docx")