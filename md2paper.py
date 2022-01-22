from __future__ import annotations
from typing import Union,List
import docx
from docx.shared import Inches

class DocNotSetException(Exception):
    pass
class DocManager():
    __doc_target = None

    @classmethod
    def set_doc(cls,doc_target:docx.Document):
        cls.__doc_target = doc_target
        cls.__clear_tables()

    @classmethod
    def get_doc(cls)->docx.Document:
        if not cls.__doc_target:
            raise DocNotSetException("doc target is not set, call DM.set_doc")
        return cls.__doc_target

    @classmethod
    def __clear_tables(cls):
        # delete all tables on startup as we don't need them
        for i in range(len(cls.get_doc().tables)):
            t = cls.get_doc().tables[0]._element
            t.getparent().remove(t)
            t._t = t._element = None

    @classmethod
    def delete_paragraph_by_index(cls, index):
        p = cls.get_doc().paragraphs[index]._element
        p.getparent().remove(p)
        p._p = p._element = None    

    @classmethod
    def get_anchor_position(cls,anchor_text:str,anchor_style_name="")->int:
        # FIXME: 需要优化
        # 目前被设计成无状态的，只依赖template.docx文件以便测试，增加了性能开销
        # USE-WITH-CARE
        # 只靠标题的anchor-text找paragraph很容易找错，用的时候注意
        i = -1
        for _i,paragraph in enumerate(cls.get_doc().paragraphs):
            if anchor_text in paragraph.text:
                if (not anchor_style_name) or (paragraph.style.name == anchor_style_name):
                    i = _i
                    break
                        
        if i==-1: raise ValueError(f"anchor `{anchor_text}` not found") 
        return i + 1

DM = DocManager

class Component():
    def __init__(self) -> None:
        self.__internal_text = Block()
    
    def set_text(self, text:str):
        self.__internal_text.add_content(content_list=Text.read(text))

    # anchor_text: 用于找到插入段落位置
    # incr_next: 用于在插入新内容后往后删老模板当前段内容，
    # 直到删除到incr_kw往前incr_next个paragraph
    # incr_kw：见上面incr_next
    def render_template(self, anchor_text:str,  incr_next:int, incr_kw, anchor_style_name="")->int:
        offset = DM.get_anchor_position(anchor_text=anchor_text,anchor_style_name=anchor_style_name)
        while not incr_kw in DM.get_doc().paragraphs[offset+incr_next].text\
                 and (offset+incr_next)!=(len(DM.get_doc().paragraphs)-1):
            DM.delete_paragraph_by_index(offset)
            #print('deleted 1 line for anchor',anchor_text)
        return self.__internal_text.render_block(offset)

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

class Image(BaseContent):
    img_url = ""
    img_alt = ""

class Formula(BaseContent):
    pass

class Table(BaseContent):
    pass

class Metadata(Component):
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
        fixme!!
        """
        def get_data_len(data:str)->int:
            # 判断是否中文, 一个中文算两个字符
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
        DM.get_doc().paragraphs[title_mapping['zh_CN']].runs[0].text = self.title_zh_CN
        DM.get_doc().paragraphs[title_mapping['en']].runs[0].text = self.title_en

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
            print(len(DM.get_doc().paragraphs[line_no].runs[-1].text))
            DM.get_doc().paragraphs[line_no].runs[-1].text = self.__fill_blank(BLANK_LENGTH,line_mapping[line_no])

class Abstract(Component):   
    __keyword_zh_CN: Text = None
    __keyword_en: Text = None
    __text_zh_CN: Block = None
    __text_en: Block = None

    def __init__(self) -> None:
        self.__text_en = Block()
        self.__text_zh_CN = Block()

    def set_keyword(self, zh_CN:List[str],en:List[str]):
        SEPARATOR = "；"
        self.__keyword_en = Text(SEPARATOR.join(en))
        self.__keyword_zh_CN = Text(SEPARATOR.join(zh_CN))

    def set_text(self, zh_CN:str,en:str):
        self.__text_en.add_content(content_list=Text.read(en))
        self.__text_zh_CN.add_content(content_list=Text.read(zh_CN))

    def render_template(self, en_title='this is english title')->int:
        # 64开始是摘要正文
        abs_cn_start = 64
        abs_cn_end = self.__text_zh_CN.render_block(abs_cn_start)
        #p = self.doc_target.paragraphs[ABSTRACT_ZH_CN_START].insert_paragraph_before(text=self.text_zh_CN.to_paragraph())
        
        # https://stackoverflow.com/questions/30584681/how-to-properly-indent-with-python-docx
        #p.paragraph_format.first_line_indent = Inches(0.25)
        
        while not DM.get_doc().paragraphs[abs_cn_end+2].text.startswith("关键词："):
            DM.delete_paragraph_by_index(abs_cn_end+1)
        
        # cn kw
        kw_cn_start = abs_cn_end + 2
        DM.get_doc().paragraphs[kw_cn_start].runs[1].text = self.__keyword_zh_CN.to_paragraph()
        
        # en start

        en_title_start = kw_cn_start+4
        DM.get_doc().paragraphs[en_title_start].runs[1].text = en_title

        en_abs_start = en_title_start + 3
        en_abs_end = self.__text_en.render_block(en_abs_start)-1
        #self.doc_target.paragraphs[en_abs_start].insert_paragraph_before(text=self.__text_en.to_paragraph())

        # https://stackoverflow.com/questions/61335992/how-can-i-use-python-to-delete-certain-paragraphs-in-docx-document
        while not DM.get_doc().paragraphs[en_abs_end+2].text.startswith("Key Words："):
            DM.delete_paragraph_by_index(en_abs_end+1)

        # en kw
        kw_en_start = en_abs_end +2

        # https://github.com/python-openxml/python-docx/issues/740
        delete_num = len(DM.get_doc().paragraphs[kw_en_start].runs) - 4
        for run in reversed(list(DM.get_doc().paragraphs[kw_en_start].runs)):
            DM.get_doc().paragraphs[kw_en_start]._p.remove(run._r)
            delete_num -= 1
            if delete_num < 1:
                break
        
        
        DM.get_doc().paragraphs[kw_en_start].runs[3].text = self.__keyword_en.to_paragraph()
        return kw_en_start+1
class Conclusion(Component):
    def render_template(self) -> int:
        ANCHOR = "结    论（设计类为设计总结）"
        incr_next = 3
        incr_kw = "参 考 文 献"
        return super().render_template(ANCHOR, incr_next, incr_kw)



class Introduction(Component): #引言 由于正文定位依赖引言，如果没写引言，依旧会生成引言，最后删掉
    def render_template(self) -> int:
        anchor_text = "引    言"
        incr_next = 2
        incr_kw = "正文格式说明"
        anchor_style_name = "Heading 1"
        return super().render_template(anchor_text, incr_next, incr_kw, anchor_style_name=anchor_style_name)


class References(Component): #参考文献
    def render_template(self) -> int:
        ANCHOR = "参 考 文 献"
        incr_next = 1
        incr_kw = "附录A"
        offset_start = DM.get_anchor_position(ANCHOR)
        offset_end = super().render_template(ANCHOR, incr_next, incr_kw) -incr_next+1
        _style = DM.get_doc().styles['参考文献正文']
        for i in range(offset_start,offset_end):
            DM.get_doc().paragraphs[i].style = _style
        return offset_end
        

class Appendixes(Component): #附录abcdefg, 是一种特殊的正文
    pass

class ChangeRecord(Component): #修改记录
    def render_template(self) -> int:
        # fixme: this anchor doesn't work, need to traverse backwards.
        # add API in render_template?
        ANCHOR = "修改记录"
        ANCHOR_STYLE = "Heading 1"
        incr_next = 0
        incr_kw = "致    谢"
        return super().render_template(ANCHOR,incr_next,incr_kw,anchor_style_name=ANCHOR_STYLE)

class Acknowledgments(Component): #致谢
    def render_template(self) -> int:
        ANCHOR = "致    谢"
        incr_next = 0

        #hack: 致谢已经到论文末尾，因此用无法匹配上的字符串直接让他删到最后一行
        incr_kw = "/\,.;'" 
        return super().render_template(ANCHOR,incr_next,incr_kw)
    

class Block(): #content
    # 每个block是多个image，formula，text的组合，内部有序
    
    def __init__(self) -> None:
        self.__title:str = None
        self.__content_list:List[Union[Text,Image,Formula]] = []
        self.__sub_blocks:List[Block] = []
        self.__id:int = None

    def set_id(self, id:int):
        self.__id = id

    def set_title(self,title:str) -> None:
        self.__title = title
    
    def add_sub_block(self,block:Block)->Block:
        self.__sub_blocks.append(block)
        return block

    def add_content(self,content:Union[Text,Image,Formula]=None,
            content_list:Union[List[Text],List[Image],List[Formula]]=[]) -> Block:
        if content:
            self.__content_list.append(content)
        for i in content_list:
            self.__content_list.append(i)
        #print('added content with len',len(content_list),content_list[0].raw_text,id(self))
        return self

    # render_template是基于render_block的api，增加了嵌套blocks的渲染 以支持递归生成章节/段落，
    # 同时增加了对段落标题和段落号的支持
    # 顺序：先title，再自己的content-list，再自己的sub-block
    def render_template(self, offset:int)->int:
        new_offset = offset
        if self.__title:
            p_title = DM.get_doc().paragraphs[offset].insert_paragraph_before()
            p_title.style = DM.get_doc().styles['Heading 1']
            p_title.add_run()
            p_title.runs[0].text= str(self.__id) if self.__id else "" +"  "+self.__title
            new_offset = new_offset + 1

        new_offset = self.render_block(new_offset)

        for i,block in enumerate(self.__sub_blocks):
            new_offset = block.render_template(new_offset) 

        return new_offset


    # render_block是最底层的api，只将自己的content-list加到已有文档给定位置
    # render_block takes the desired paragraph position's offset,
    # renders the block with native elements: text, image and formulas,
    # and returns the final paragraph's offset
    def render_block(self, offset:int)->int:
        if not self.__content_list:
            return offset
        # generate necessary paragraphs
        internal_content_list = []
        p = DM.get_doc().paragraphs[offset].insert_paragraph_before()
        internal_content_list.insert(0,p)
        for i in range(len(self.__content_list)-1):
            p = internal_content_list[0].insert_paragraph_before()
            internal_content_list.insert(0,p)
            
        #print('got content list',len(internal_content_list),id(self))
        assert len(self.__content_list)==len(internal_content_list)
        for i in range(len(self.__content_list)):
            if type(self.__content_list[i]) == Text:
                p = internal_content_list[i]
                p.style = DM.get_doc().styles['Normal']
                p.text = self.__content_list[i].to_paragraph()
                p.paragraph_format.first_line_indent = Inches(0.25)
            else:
                raise NotImplementedError
        return offset + len(self.__content_list)

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
    DM.set_doc(doc)
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

    abs = Abstract()
    a = """CommonMark中并未定义普通文本高亮。
CSDN中支持通过==文本高亮==，实现文本高亮

如果你使用的markdown编辑器不支持该便捷设置，可通过HTML标签<mark>实现：
语法：<mark>文本高亮<mark>
效果：文本高亮"""
    b = ['您配吗','那匹马','进来看是否']
    
    c = """Any subsequent access to the "deleted" paragraph object will raise AttributeError, so you should be careful not to keep the reference hanging around, including as a member of a stored value of Document.paragraphs.
The reason it's not in the library yet is because the general case is much trickier, in particular needing to detect and handle the variety of linked items that can be present in a paragraph; things like a picture, a hyperlink, or chart etc.
But if you know for sure none of those are present, these few lines should get the job done."""
    
    d = ['abc','def','gh']

    abs.set_text(a,c)
    abs.set_keyword(b,d)
    abs.render_template()

    intro = Introduction()
    t = """这样做违反了Liskov替代原则。换句话说，这是一个可怕的想法，B不应该是A的子类型。我只是感兴趣：您为什么要这样做？@delnan出于某种原因每当有人提到我总是想到Who Doctor的Blinovitch限制效应时。
现在就称其为好奇心。我感谢警告，但我仍然感到好奇。
一个用例是，如果您要使用Django库公开的Form类，但不包含其字段之一。在Django中，表单字段是由某些类属性定义的。例如，请参阅此SO问题。"""
    intro.set_text(t)
    intro.render_template()

    conc = Conclusion()
    e = """如果代码中出现太多的条件判断语句的话，代码就会变得难以维护和阅读。 这里的解决方案是将每个状态抽取出来定义成一个类。
这里看上去有点奇怪，每个状态对象都只有静态方法，并没有存储任何的实例属性数据。 实际上，所有状态信息都只存储在 Connection 实例中。 
在基类中定义的 NotImplementedError 是为了确保子类实现了相应的方法。 这里你或许还想使用8.12小节讲解的抽象基类方式。
设计模式中有一种模式叫状态模式，这一小节算是一个初步入门！"""
    #conc.set_conclusion(e)
    conc.set_text(e)
    conc.render_template()

    ref = References()
    h = """[1] 国家标准局信息分类编码研究所.GB/T 2659-1986 世界各国和地区名称代码[S]//全国文献工作标准化技术委员会.文献工作国家标准汇编:3.北京:中国标准出版社,1988:59-92. 
[2] 韩吉人.论职工教育的特点[G]//中国职工教育研究会.职工教育研究论文集.北京:人民教育出版社,1985:90-99. """
    ref.set_text(h)
    ref.render_template()

    ack = Acknowledgments()
    f = """肾衰竭（Kidney failure）是一种终末期的肾脏疾病，此时肾脏的功能会低于其正常水平的15%。由于透析会严重影响患者的生活质量，肾移植一直是治疗肾衰竭的理想方式。但肾脏供体一直处于短缺状态，移植需等待时间约为5-10年。近日，据一篇发表于《美国移植杂志》的文章，阿拉巴马大学伯明翰分校的科学家首次成功将基因编辑猪的肾脏成功移植给一名脑死亡的人类接受者。
研究使用的供体猪的10个关键基因经过了基因编辑，其肾脏的功能更适合人体，且植入人体后引发的免疫排斥反应更轻微。研究人员首先对异种供体和接受者进行交叉配型测试。经配对后，研究人员将肾脏移植入脑死亡接受者的肾脏对应的解剖位置，与肾动脉、肾静脉和输尿管相连接。移植后，他们为患者进行了常规的免疫抑制治疗。目前，肾脏已在患者体内正常工作77小时。该研究按照1期临床试验标准进行，完全按人类供体器官的移植标准实施。研究显示，异种移植的发展在未来或可缓解世界范围器官供应压力。"""
    ack.set_text(f)
    ack.render_template()

    cha = ChangeRecord()
    g = """在无线传能技术中，非辐射无线传能（即电磁感应充电）可以高效传输能量，但有效传输距离被限制在收发器尺寸的几倍之内；而辐射无线传能（如无线电、激光）虽然可以远距离传输能量，但需要复杂的控制机制来跟踪移动的能量接收器。
近日，同济大学电子与信息工程学院的研究团队通过理论和实验证明，"""
    cha.set_text(g)
    cha.render_template()

    doc.save("out.docx")

    