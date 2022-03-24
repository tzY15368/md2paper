import os
if os.path.split(os.getcwd())[1] != 'test':
    os.chdir('test')
import sys
sys.path.append('..')
import logging
logging.getLogger().setLevel(logging.DEBUG)
from md2paper.dut_paper_translation import *

if __name__ == "__main__":
    DM.set_doc("word-template/外文翻译模板-docx.docx")
    meta = TranslationMetadata()
    meta.school = "电子信息与电气工程"
    meta.number = "201800000"
    meta.auditor = "张三"
    meta.finish_date = "1234年5月6号"
    meta.teacher = "里斯"
    meta.major = "计算机科学与技术"
    meta.title_en = "what is this world"
    meta.title_zh_CN = "这是个怎样的世界"
    meta.render_template()

    abs = TranslationAbstract("塔塔开塔塔开","eren yaeger","帕拉迪岛")
    abs.add_keywords(["塔塔开","都得死"])
    
    images = Image([
        ImageData("classes.png","图1：these are the classes")
    ])
    formula = Formula("公式3.444",r"\sum_{i=1}^{10}{\frac{\sigma_{zp,i}}{E_i} kN")
    more_text = Text().add_run(Run("only italics",Run.Italics))
    abs.add_text("如果代码中出现太多的条件判断语句的话，代码就会变得难以维护和阅读。 这里的解决方案是将每个状态抽取出来定义成一个类。这里看上去有点奇怪，每个状态对象都只有静态方法")
    abs.add_text([images,images,formula,more_text])

    abs.render_template()

    mc = TranslationMainContent()
    c1 = mc.add_chapter("第一章 刘姥姥")
    s1 = mc.add_section("1.1 asdfasdf",chapter=c1)
    s2 = mc.add_section("1.2 bbbb",chapter=c1)
    h = """目前的娛樂型電腦螢幕市場，依照玩家的需求大致可以分為兩大勢力：一派是主打對戰類型
    的電競玩家、另一派則主打追劇的多媒體影音玩家。前者需要需要高更新率的螢幕，在分秒必爭的對戰中搶得先機；後者則需要較高的解析度以及HDR的顯示內容，好用來欣賞畫面的每一個細節。"""
    mc.add_text(h,location=c1)
    mc.add_text(h,location=s2)
    c2 = mc.add_chapter("第二章 菜花")
    s3 = mc.add_section("2.1 aaa",chapter=c2)
    ss1 = mc.add_subsection("2.1.1 asdf",section=s3)
    txt = mc.add_text(h,location=ss1)
    txt.add_run(Run("this should be bold",Run.Bold))
    txt.add_run(Run("italic and bold",Run.Italics|Run.Bold))
    images = Image([
        ImageData("classes.png","图1：these are the classes"),
        ImageData("classes.png","图2:asldkfja;sldkf")
    ])
    formula = Formula("公式3.4",r"\sum_{i=1}^{10}{\frac{\sigma_{zp,i}}{E_i} kN")
    more_text = Text().add_run(Run("only italics",Run.Italics))
    mc.add_text([images,formula,more_text])

    c3 = mc.add_chapter("第三章 大观园")
    t = """Any subsequent access to the "deleted" paragraph object will raise AttributeError, so you should be careful not to keep the reference hanging around, including as a member of a stored value of Document.paragraphs.
The reason it's not in the library yet is because the general case is much trickier, in particular needing to detect and handle the variety of linked items that can be present in a paragraph; things like a picture, a hyperlink, or chart etc.
But if you know for sure none of those are present, these few lines should get the job done."""
    mc.add_text(t,location=c3)
    data = [
        Row(['第一章','第二章','第三章'],top_border=True),
        Row(['刘姥姥初试钢铁侠','刘姥姥初试大不净者','刘姥姥倒拔绿巨人'],top_border=True),
        Row(['刘姥姥初试惊奇队长',None,'刘姥姥菜花染诸神']),
        Row(['菜花反噬！','天地乖离菜花之星','重启刘姥姥菜花宇宙'],top_border=True)
    ]
    table = Table("表1 刘姥姥背叛斯大林",data)
    mc.add_text([table, Text("wtf is this?")],location=c3)
    mc.render_template()


    DM.save("out_translation.docx")