import sys
sys.path.append('..')
from md2paper.v2.backend.docx_render import *

DM.set_doc("../../word-template/毕业设计（论文）模板-docx.docx")
# for par in DM.get_doc().paragraphs:
#     print(par.style.name)
anc = "摘    要"
pos = DM.get_anchor_position(anc, "Heading 1") - 1
for i in range(pos, len(DM.get_doc().paragraphs)):
    DM.delete_paragraph_by_index(pos)

blk_pkg = Block()
blk_abs = Block()
blk_main = Block()
blk_ref = Block()
blk_c1 = Block()
blk_c2 = Block()

blk_abs.set_title("摘    要", level=Block.Heading_1)
blk_c1.set_title("1 稍有常识的人", level=Block.Heading_1)
blk_c2.set_title("2 没有尝试的人", level=Block.Heading_1)
blk_ref.set_title("参 考 文 献", level=Block.Heading_1)

c = """Any subsequent access to the "deleted" paragraph object will raise AttributeError, so you should be careful not to keep the reference hanging around, including as a member of a stored value of Document.paragraphs.
The reason it's not in the library yet is because the general case is much trickier, in particular needing to detect and handle the variety of linked items that can be present in a paragraph; things like a picture, a hyperlink, or chart etc.
But if you know for sure none of those are present, these few lines should get the job done."""
kw = Text("关键字：那匹马；您配吗；你怕吗",force_style='关键词',style=Run.Bold)
blk_abs.add_content(Text(c),Text(),kw)

e = """如果代码中出现太多的条件判断语句的话，代码就会变得难以维护和阅读。 这里的解决方案是将每个状态抽取出来定义成一个类。
这里看上去有点奇怪，每个状态对象都只有静态方法，并没有存储任何的实例属性数据。 实际上，所有状态信息都只存储在 Connection 实例中。 
在基类中定义的 NotImplementedError 是为了确保子类实现了相应的方法。 这里你或许还想使用8.12小节讲解的抽象基类方式。
设计模式中有一种模式叫状态模式，这一小节算是一个初步入门！"""
img = Image("图1 classes","../test/classes.png")
img.set_image_data(ImageData(img.src,img.title,width_ratio=0.2))
blk_c1.add_content(Text(e[:60]),Text(e[:60]),img,Text(e[60:]))
blk_c1.add_content(img,Formula("公式1 真的",r"\sum^{n}_{i=0}{i}"),Text("this is eof"))

blk_c21 = Block()
blk_c21.set_title(title='2.1 最大限度',level=Block.Heading_2)
blk_c21.add_content(Text("表格测试"))

data = [
        Row([Text('第一章'),Text('第二章'),Text('第三章')],top_border=True),
        Row([Text('刘姥姥初试钢铁侠'),Text('刘姥姥初试大不净者'),Text('刘姥姥倒拔绿巨人')],top_border=True),
        Row([Text('刘姥姥初试惊奇队长'),None,Text('刘姥姥菜花染诸神')]),
        Row([Text('菜花反噬！'),Text().add_run(Run(r"\sum^{n}_{i=0}{i}",style=Run.Formula)),Text('重启刘姥姥菜花宇宙')],top_border=True)
    ]
table = Table("表1 真的是表",data)

ol = OrderedList([
    Text("helo, thank you, thank you very much"),
    Text("how you doing?"),
    img
])

blk_c21.add_content(table,Text("boom!"),ol)
blk_c2.add_sub_block(blk_c21)

h = """[1] 国家标准局信息分类编码研究所.GB/T 2659-1986 世界各国和地区名称代码[S]//全国文献工作标准化技术委员会.文献工作国家标准汇编:3.北京:中国标准出版社,1988:59-92. 
[2] 韩吉人.论职工教育的特点[G]//中国职工教育研究会.职工教育研究论文集.北京:人民教育出版社,1985:90-99. """
txts = Text.read(h)
for txt in txts:
    txt.first_line_indent = -Text.two_chars_length
    txt.force_style = "参考文献正文"
blk_ref.add_content(*txts)

blk_main.add_sub_block(blk_c1).add_sub_block(blk_c2)
blk_pkg.add_sub_block(blk_abs).add_sub_block(blk_main).add_sub_block(blk_ref)
blk_pkg.render_template()

DM.save('out.docx')
