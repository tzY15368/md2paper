import markdown
from bs4 import BeautifulSoup, Comment
import logging
import re
from functools import reduce

from md2paper import *


# 检查

def assert_warning(e: bool, s: str):
    if not e:
        logging.warning(s)
    return e


def assert_error(e: bool, s: str):
    if not e:
        logging.error(s)
        exit(-1)
    return e


def log_error(s: str):
    return assert_error(False, s)


main_h_level = [0]


def process_headline(h_label: str, headline: str):
    global main_h_level
    level = int(h_label[1:])
    assert_warning(1 <= level and level <= len(main_h_level)+1,
                   "标题层级应该递进")
    if level == len(main_h_level) + 1:  # new sub section
        main_h_level.append(1)
    elif 1 <= level and level <= len(main_h_level):  # new section
        main_h_level[level-1] += 1
        main_h_level = main_h_level[:level]
    else:
        log_error("错误的标题编号")

    index = str(main_h_level[0])
    for i in range(1, len(main_h_level)):
        index += "." + str(main_h_level[i])

    headline = headline.strip()
    assert_warning(headline[:len(index)] == index, "没有编号或者编号错误")
    assert_warning(headline[len(index):len(index)+2] == "  " and
                   headline[len(index)+2] != " ",
                   "编号后应该有两个空格: " + headline)
    headline = headline[:len(index)+2] + rbk(headline[len(index)+2:])

    return (h_label, headline)


# 处理文本

def rbk(text: str):  # remove_blank
    # 删除换行符
    text = text.replace("\n", "")
    text = text.replace("\r", "")

    cn_char = u'[\u4e00-\u9fa5。，：《》、（）“”‘’]'
    # 中文字符后空格
    should_replace_list = re.compile(
        cn_char + u' +').findall(text)
    # 中文字符前空格
    should_replace_list += re.compile(
        u' +' + cn_char).findall(text)
    # 删除空格
    for i in should_replace_list:
        if i == u' ':
            continue
        new_i = i.strip()
        text = text.replace(i, new_i)
    return text


def assemble(texts: list[str]):
    return reduce(lambda x, y: x+"\n"+y, texts)


# 提取内容

def get_ps(h1):
    ps = []
    cur = h1.next_sibling
    while cur != None and cur.name != "h1":
        if cur.name == "p":
            ps.append(rbk(cur.text))
        cur = cur.next_sibling
    return ps


def get_content(h1, until_h1):
    conts = []
    cur = h1.next_sibling
    while cur != until_h1:
        if cur.name != None:
            if cur.name == "table":
                conts.append(("table", "something"))  # FIXME
            elif cur.name[0] == "h":  # h1 h2 ...
                headline_pair = process_headline(cur.name, cur.text)
                conts.append(headline_pair)
            else:
                conts.append((cur.name, rbk(cur.text)))
        cur = cur.next_sibling
    return conts


# 获得每个论文模块

def get_metadata(soup: BeautifulSoup):
    mete_h1 = soup.find("h1")

    data_table = mete_h1.find_next_sibling("table").find("tbody")
    data_lines = data_table.find_all("tr")
    data_pairs = [list(map(lambda x: rbk(x.text), i.find_all("td")))
                  for i in data_lines]
    data_dict = dict(data_pairs)

    meta = Metadata()
    meta.title_zh_CN = rbk(mete_h1.text)
    meta.title_en = rbk(mete_h1.find_next_sibling("h2").text)
    meta.school = data_dict["学院（系）"]
    meta.major = data_dict["专业"]
    meta.name = data_dict["学生姓名"]
    meta.number = data_dict["学号"]
    meta.teacher = data_dict["指导教师"]
    meta.auditor = data_dict["评阅教师"]
    meta.finish_date = data_dict["完成日期"]

    return meta


def get_abs(soup: BeautifulSoup):
    # 摘要
    abs_cn_h1 = soup.find("h1", string=re.compile("摘要"))
    ps_key = get_ps(abs_cn_h1)
    assert_warning(ps_key[-1] == "关键词：", '摘要应该以"关键词："后接关键词列表结尾')
    ps_cn = ps_key[:-1]
    keywords_cn = [rbk(i.text)
                   for i in abs_cn_h1.find_next_sibling("ul").find_all("li")]

    # Abstract
    abs_h1 = soup.find("h1", string=re.compile("Abstract"))
    ps_key = get_ps(abs_h1)
    assert_warning(ps_key[-1] == "Key Words:",
                   'Abstract应该以"Key Words:"后接关键词列表结尾')
    ps_en = ps_key[:-1]
    keywords_en = [rbk(i.text)
                   for i in abs_h1.find_next_sibling("ul").find_all("li")]

    abs = Abstract()
    abs.set_text(assemble(ps_cn), assemble(ps_en))
    abs.set_keyword(keywords_cn, keywords_en)

    return abs


def get_intro(soup: BeautifulSoup):
    intro_h1 = soup.find("h1", string=re.compile("引言"))
    ps = get_ps(intro_h1)  # TODO

    intro = Introduction()
    intro.set_text(assemble(ps))

    return intro


def get_body(soup: BeautifulSoup):
    body_h1 = soup.find("h1", string=re.compile("正文"))
    content = get_content(body_h1,
                          soup.find("h1", string=re.compile("结论")))

    mc = MainContent()
    for (name, text) in content:
        if name == "h1":
            chapter = mc.add_chapter(text)
        elif name == "h2":
            section = mc.add_section(chapter, text)
        elif name == "h3":
            subsection = mc.add_subsection(section, text)
        elif name == "p":
            mc.append_paragraph(text)
        else:
            print("还没实现now", name)

    return mc


def get_conclusion(soup: BeautifulSoup):
    conclusion_h1 = soup.find("h1", string=re.compile("结论"))
    ps = get_ps(conclusion_h1)

    conclusion = Conclusion()
    conclusion.set_text(assemble(ps))

    return conclusion


def get_reference(soup: BeautifulSoup):
    reference_h1 = soup.find("h1", string=re.compile("参考文献"))
    # 需要一个专门的处理方式
    print(reference_h1)  # FIXME
    return "tmp"  # TODO


def get_appendix(soup: BeautifulSoup):
    appendix_h1s = soup.find_all("h1", string=re.compile("附录"))
    for i in range(0, len(appendix_h1s)-1):
        content = get_content(appendix_h1s[i], appendix_h1s[i+1])
    content = get_content(appendix_h1s[-1],
                          soup.find("h1", string=re.compile("修改记录")))
    print(appendix_h1s)  # FIXME
    return "tmp"  # TODO


def get_record(soup: BeautifulSoup):
    mod_record_h1 = soup.find("h1", string=re.compile("修改记录"))
    content = get_content(mod_record_h1,
                          soup.find("h1", string=re.compile("致谢")))
    print(mod_record_h1)  # FIXME
    return "tmp"  # TODO


def get_thanks(soup: BeautifulSoup):
    thanks_h1 = soup.find("h1", string=re.compile("致谢"))
    ps = get_ps(thanks_h1)

    ack = Acknowledgments()
    ack.set_text(assemble(ps))
    return ack


# 处理文章

def handle_paper(soup: BeautifulSoup):
    # test
    with open("out.html", "w") as f:
        f.write(soup.prettify())

    data_ls = [
        get_metadata(soup),    # metadata
        get_abs(soup),         # 摘要 Abstract
                               # 目录 pass
        get_intro(soup),       # 引言
        get_body(soup),        # 正文
        get_conclusion(soup),  # 结论
        get_reference(soup),   # 参考文献
        get_appendix(soup),    # 附录
        get_record(soup),      # 修改记录
        get_thanks(soup)       # 致谢
    ]
    return data_ls


def handle_translation(soup: BeautifulSoup):
    pass


def load_md(file_name: str, file_type: str):
    with open(file_name, "r") as f:
        md_file = f.read()
    md_html = markdown.markdown(md_file,
                                extensions=['markdown.extensions.tables'])
    soup = BeautifulSoup(md_html, 'html.parser')
    for i in soup(text=lambda text: isinstance(text, Comment)):
        i.extract()  # 删除 html 注释

    if file_type == "论文":
        return handle_paper(soup)
    elif file_type == "外文翻译":
        return handle_translation(soup)
    else:
        log_error('错误的文件类型，应该选择 "论文" / "外文翻译"')


if __name__ == "__main__":
    doc = docx.Document("毕业设计（论文）模板-docx.docx")
    DM.set_doc(doc)

    paper = load_md("论文模板.md", "论文")
    paper[0].render_template()  # metadata
    paper[1].render_template()  # 摘要 Abstract
    main_start = paper[2].render_template()  # 引言
    paper[3].render_template(main_start)  # 正文
    paper[4].render_template()  # 结论
    # paper[5].render_template()  # 参考文献
    # paper[6].render_template()  # 附录
    # paper[7].render_template()  # 修改记录
    paper[8].render_template()  # 致谢

    DM.update_toc()
    doc.save("out.docx")
