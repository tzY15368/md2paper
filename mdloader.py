import markdown
from bs4 import BeautifulSoup, Comment
import logging
import re
from functools import reduce
import os
import docx
from docx import Document

from mdext import MDExt
import md2paper as word

file_dir = ""
debug = True


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


# 处理文本

def rbk(text: str):  # remove_blank
    # 删除换行符
    text = text.replace("\n", " ")
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


def raw_text(runs):
    strs = [i["text"] for i in runs]
    return reduce(lambda x, y: x+y, strs)


def assemble_ps(ps):
    strs = []
    for (_, runs) in ps:
        strs.append(raw_text(runs))
    return reduce(lambda x, y: x+"\n"+y, strs)


def split_title(title):
    assert_error(len(title.split(':')) >= 2, "应该有别名或者标题: " + title)
    ali = title.split(':')[0]
    title = rbk(title[len(ali)+1:].strip())
    return ali, title


# 处理标签

def process_headline(head_counter: list[int], h_label: str, headline: str):
    level = int(h_label[1:])
    assert_warning(1 <= level and level <= len(head_counter)+1,
                   "标题层级应该递进" + headline)
    if level == len(head_counter) + 1:  # new sub part
        head_counter.append(1)
    elif 1 <= level and level <= len(head_counter):  # new part
        head_counter[level-1] += 1
        head_counter = head_counter[:level]
    else:
        log_error("错误的标题编号")

    index = str(head_counter[0])
    for i in range(1, len(head_counter)):
        index += "." + str(head_counter[i])

    headline = headline.strip()
    assert_warning(headline[:len(index)] == index,
                   "没有编号或者编号错误: {} {}".format(h_label, headline))
    assert_warning(headline[len(index):len(index)+2] == "  " and
                   headline[len(index)+2] != " ",
                   "编号后应该有两个空格: {} {}".format(h_label, headline))
    headline = headline[:len(index)+2] + rbk(headline[len(index)+2:])

    return head_counter, (h_label, headline)


def process_ps(p, ollevel=4):
    ps = []
    data = []
    for i in p.children:
        if i.name == None:
            if i.text == "\n":
                continue
            data.append({"type": "text", "text": rbk(i.text)})
        elif i.name == "strong":
            assert_warning(len(i.contents) == 1, "只允许粗斜体，不允许复杂嵌套")
            if i.contents[0].name == "em":
                data.append({"type": "strong-em", "text": rbk(i.text)})
            else:
                data.append({"type": "strong", "text": rbk(i.text)})
        elif i.name == "em":
            data.append({"type": "em", "text": rbk(i.text)})
        elif i.name == "math-inline":
            data.append({"type": "math-inline", "text": i.text})
        elif i.name == "ref":
            data.append({"type": "ref", "text": rbk(i.text.strip())})
        else:  # 需要分段
            if data:
                ps.append(("p", data))
                data = []
            if i.name == "br":  # 分段
                pass
            elif i.name == "img":  # 图片
                ps.append(process_img(i))
            elif i.name == "ol":
                ps += process_ol(i, ollevel)
            else:
                log_error("缺了什么？" + str(i))
    if data:
        ps.append(("p", data))
    return ps


def process_img(img):
    global file_dir
    img_path = os.path.join(file_dir, img["src"])
    ali, title = split_title(img["alt"])
    return ("img", {"alias": ali,
                    "title": title,
                    "src": img_path})


def process_table(title, table):
    data = []
    # 表头，有上实线
    data.append(word.Row([rbk(i.text) for i in table.find("thead").find_all("th")],
                         top_border=True))
    has_border = True  # 表身第一行有上实线
    for tr in table.find("tbody").find_all("tr"):
        row = [rbk(i.text) for i in tr.find_all("td")]  # get all text
        row = list(map(lambda x: None if x == '' else x,
                       row))  # replace '' with None
        if has_border:
            data.append(word.Row(row, top_border=True))
            has_border = False
        else:
            is_border = True
            for i in row:
                if i == None:
                    is_border = False
                    break
                for j in i:
                    if j != '-':
                        is_border = False
                        break
            if is_border:
                has_border = True  # 自定义的实线，下一行数据有上实线
            else:
                data.append(word.Row(row))

    ali, title = split_title(title)
    return ("table", {"alias": ali,
                      "title": title,
                      "data": data})


def process_lis(li, level):
    if (li.contents[0].text == "\n"):  # <p>
        conts = get_content_from(li.contents[0], level+1)
    else:  # text
        conts = process_ps(li, level+1)
    conts[0] = ("fh" + str(level), conts[0][1])
    return conts


def process_ol(ol, level):
    assert_error(level <= 5, "层次至多两层")
    datas = [process_lis(i, level)
             for i in ol.find_all("li", recursive=False)]
    # make index
    for i in range(len(datas)):
        li_data = datas[i][0]
        if level == 4:
            li_data[1].insert(
                0, {"type": "text", "text": "（{}） ".format(i+1)})
        else:
            assert_warning(i < 20, "层次二不能超过 20 项")
            li_data[1].insert(
                0, {"type": "text", "text": "{} ".format(chr(i+0x2460))})  # get ①②..⑳
    data = reduce(lambda x, y: x + y, datas)
    return data


def process_math(title, math):
    return ("math", {"alias": title,
                     "title": "",
                     "text": math.text})


# 提取内容

def get_content_until(cur, until, ollevel=4):
    conts = []
    head_counter = [0]
    while cur != until:
        if cur.name == None:
            cur = cur.next_sibling
            continue
        if cur.name[0] == "h":  # h1 h2 h3
            head_counter, pair = process_headline(head_counter,
                                                  cur.name, cur.text)
            conts.append(pair)
        elif cur.name == "p":
            conts += process_ps(cur)
        elif cur.name == "table":
            table_name = raw_text(conts[-1][1])
            conts = conts[:-1]
            conts.append(process_table(table_name, cur))
        elif cur.name == "ol":
            conts += process_ol(cur, ollevel)
        elif cur.name == "math":
            math_title = raw_text(conts[-1][1])
            conts = conts[:-1]
            conts.append(process_math(math_title, cur))
        else:
            log_error("这是啥？" + cur.prettify())
        cur = cur.next_sibling
    return conts


def get_content_from(cur, ollevel=4):
    return get_content_until(cur, None, ollevel)


def set_content(cont_block, conts):
    for (name, cont) in conts:
        if name == "h1":
            cont_block.add_chapter(cont)
        elif name == "h2":
            cont_block.add_section(cont)
        elif name == "h3":
            cont_block.add_subsection(cont)
        elif name in ["p", "fh4", "fh5"]:
            if not debug:
                para = cont_block.add_text("")
            else:
                para = cont_block.add_text(name)
            for run in cont:
                if run["type"] == "text":
                    para.add_run(word.Run(run["text"], word.Run.normal))
                elif run["type"] == "strong":
                    para.add_run(word.Run(run["text"], word.Run.bold))
                elif run["type"] == "em":
                    para.add_run(word.Run(run["text"], word.Run.italics))
                elif run["type"] == "strong-em":
                    para.add_run(word.Run(run["text"],
                                          word.Run.italics | word.Run.bold))
                elif run["type"] == "math-inline":
                    para.add_run(word.Run(run["text"], word.Run.formula))
                elif run["type"] == "ref":
                    para.add_run(word.Run(run["text"], word.Run.normal))
                else:
                    print("还没实现now", name)
        elif name == "img":
            cont_block.add_image([word.ImageData(cont["src"], cont["title"])])
        elif name == "table":
            cont_block.add_table(cont['title'], cont['data'])
        elif name == "math":
            cont_block.add_formula(cont['title'], cont['text'])
        else:
            print("还没实现now", name)


# 索引处理

def get_index(conts):
    index_table = {}
    text_table = {}
    chapter_cnt = 0

    for name, cont in conts:
        if name == "h1":
            chapter_cnt += 1
            img_cnt = 0
            table_cnt = 0
            formula_cnt = 0
        elif name in ["img", "table", "math"]:
            ali = cont['alias']
            assert_warning(ali not in index_table, "有重复别名" + ali)
            if name == "img":
                img_cnt += 1
                index_table[ali] = "{}.{}".format(chapter_cnt, img_cnt)
                text_table[ali] = "图{}.{}".format(chapter_cnt, img_cnt)
                cont['title'] = "图{}  {}".format(
                    index_table[ali], cont['title'])
            elif name == "table":
                table_cnt += 1
                index_table[ali] = "{}.{}".format(chapter_cnt, table_cnt)
                text_table[ali] = "表{}.{}".format(chapter_cnt, table_cnt)
                cont['title'] = "表{}  {}".format(
                    index_table[ali], cont['title'])
            elif name == "math":
                formula_cnt += 1
                index_table[ali] = "{}.{}".format(chapter_cnt, formula_cnt)
                text_table[ali] = "式{}.{}".format(chapter_cnt, formula_cnt)
                cont['title'] = "（{}）".format(index_table[ali])

    for name, cont in conts:
        if name not in ["p", "fh4", "fh5"]:
            continue
        for run in cont:
            if run["type"] != "ref":
                continue
            if run["text"] in text_table:
                run["text"] = text_table[run["text"]]
            else:
                print("未知ref: " + run["text"])

    return conts


# 获得每个论文模块

def get_metadata(soup: BeautifulSoup):
    mete_h1 = soup.find("h1")

    data_table = mete_h1.find_next_sibling("table").find("tbody")
    data_lines = data_table.find_all("tr")
    data_pairs = [list(map(lambda x: rbk(x.text), i.find_all("td")))
                  for i in data_lines]
    data_dict = dict(data_pairs)

    meta = word.Metadata()
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
    abs_cn_ul = abs_cn_h1.find_next_sibling("ul")
    conts_cn = get_content_until(abs_cn_h1.next_sibling, abs_cn_ul)
    assert_warning(conts_cn[-1] == ("p", [{"type": "text", "text": "关键词："}]),
                   '摘要应该以"关键词："后接关键词列表结尾')
    conts_cn = conts_cn[:-1]
    keywords_cn = [rbk(i.text)
                   for i in abs_cn_h1.find_next_sibling("ul").find_all("li")]

    # Abstract
    abs_en_h1 = soup.find("h1", string=re.compile("Abstract"))
    abs_en_ul = abs_en_h1.find_next_sibling("ul")
    conts_en = get_content_until(abs_en_h1.next_sibling, abs_en_ul)
    assert_warning(conts_en[-1] == ("p", [{"type": "text", "text": "Key Words:"}]),
                   'Abstract应该以"Key Words:"后接关键词列表结尾')
    conts_en = conts_en[:-1]
    keywords_en = [rbk(i.text)
                   for i in abs_en_h1.find_next_sibling("ul").find_all("li")]

    # TODO
    # abs sp check

    abs = word.Abstract()
    abs.add_text(assemble_ps(conts_cn), assemble_ps(conts_en))
    abs.set_keyword(keywords_cn, keywords_en)

    return abs


def get_intro(soup: BeautifulSoup):
    intro_h1 = soup.find("h1", string=re.compile("引言"))
    conts = get_content_until(intro_h1.next_sibling,
                              soup.find("h1", string=re.compile("正文")))

    # TODO
    # intro sp check

    intro = word.Introduction()
    intro.add_text(assemble_ps(conts))  # FIXME

    return intro


def get_body(soup: BeautifulSoup):
    body_h1 = soup.find("h1", string=re.compile("正文"))
    conts = get_content_until(body_h1.next_sibling,
                              soup.find("h1", string=re.compile("结论")))
    conts = get_index(conts)

    mc = word.MainContent()
    set_content(mc, conts)

    return mc


def get_conclusion(soup: BeautifulSoup):
    conclusion_h1 = soup.find("h1", string=re.compile("结论"))
    conts = get_content_until(conclusion_h1.next_sibling,
                              soup.find("h1", string=re.compile("参考文献")))
    # TODO
    # conclusion sp check

    conclusion = word.Conclusion()
    conclusion.add_text(assemble_ps(conts))  # FIXME

    return conclusion


def get_reference(soup: BeautifulSoup):
    reference_h1 = soup.find("h1", string=re.compile("参考文献"))
    # 需要一个专门的处理方式
    print(reference_h1)  # FIXME
    return "tmp"  # TODO


def get_appendix(soup: BeautifulSoup):
    appendix_h1s = soup.find_all("h1", string=re.compile("附录"))
    appendix_h1s.append(soup.find("h1", string=re.compile("修改记录")))
    for i in range(0, len(appendix_h1s)-1):
        conts = get_content_until(appendix_h1s[i].next_sibling,
                                  appendix_h1s[i+1])
    # TODO
    # appendix sp check
    # some thing add_content
    # if no appendix
    print(appendix_h1s)  # FIXME
    return "tmp"  # TODO


def get_record(soup: BeautifulSoup):
    mod_record_h1 = soup.find("h1", string=re.compile("修改记录"))
    conts = get_content_until(mod_record_h1.next_sibling,
                              soup.find("h1", string=re.compile("致谢")))
    # TODO
    # record sp check
    # some thing add_content
    print(mod_record_h1)  # FIXME
    return "tmp"  # TODO


def get_thanks(soup: BeautifulSoup):
    thanks_h1 = soup.find("h1", string=re.compile("致谢"))
    conts = get_content_from(thanks_h1.next_sibling)

    # TODO
    # thanks sp check

    ack = word.Acknowledgments()
    ack.add_text(assemble_ps(conts))
    return ack


# 处理文章

def handle_paper(soup: BeautifulSoup):
    if debug:
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
                                tab_length=3,
                                extensions=['markdown.extensions.tables',
                                            MDExt()])
    soup = BeautifulSoup(md_html, 'html.parser')
    for i in soup(text=lambda text: isinstance(text, Comment)):
        i.extract()  # 删除 html 注释

    global file_dir
    file_dir = os.path.dirname(file_name)
    if file_type == "论文":
        return handle_paper(soup)
    elif file_type == "外文翻译":
        return handle_translation(soup)
    else:
        log_error('错误的文件类型，应该选择 "论文" / "外文翻译"')


class Paper:
    def load_md(self, md_path: str):
        with open(md_path, "r") as f:
            md_file = f.read()
        md_html = markdown.markdown(md_file,
                                    tab_length=3,
                                    extensions=['markdown.extensions.tables',
                                                MDExt()])
        self.soup = BeautifulSoup(md_html, 'html.parser')
        for i in self.soup(text=lambda text: isinstance(text, Comment)):
            i.extract()  # 删除 html 注释

        if debug:
            with open("out.html", "w") as f:
                f.write(self.soup.prettify())


class PaperPart:
    def __init__(self):
        self.contents = []
        self.block: word.Component = None

    def _set_body(self):
        for (name, cont) in self.contents:
            if name == "h1":
                self.block.add_chapter(cont)
            elif name == "h2":
                self.block.add_section(cont)
            elif name == "h3":
                self.block.add_subsection(cont)
            elif name in ["p", "fh4", "fh5"]:
                if not debug:
                    para = self.block.add_text("")
                else:
                    para = self.block.add_text(name)
                for run in cont:
                    if run["type"] == "text":
                        para.add_run(word.Run(run["text"], word.Run.normal))
                    elif run["type"] == "strong":
                        para.add_run(word.Run(run["text"], word.Run.bold))
                    elif run["type"] == "em":
                        para.add_run(word.Run(run["text"], word.Run.italics))
                    elif run["type"] == "strong-em":
                        para.add_run(word.Run(run["text"],
                                              word.Run.italics | word.Run.bold))
                    elif run["type"] == "math-inline":
                        para.add_run(word.Run(run["text"], word.Run.formula))
                    elif run["type"] == "ref":
                        para.add_run(word.Run(run["text"], word.Run.normal))
                    else:
                        print("还没实现now", name)
            elif name == "img":
                self.block.add_image(
                    [word.ImageData(cont["src"], cont["title"])])
            elif name == "table":
                self.block.add_table(cont['title'], cont['data'])
            elif name == "math":
                self.block.add_formula(cont['title'], cont['text'])
            else:
                print("还没实现now", name)

    def _set_contents(self):
        self._set_body()

    def render(self):
        self._set_contents()
        self.block.render_template()


class MetaPart(PaperPart):
    def get_contents(self, soup: BeautifulSoup):
        mete_h1 = soup.find("h1")

        data_table = mete_h1.find_next_sibling("table").find("tbody")
        data_lines = data_table.find_all("tr")
        data_pairs = [list(map(lambda x: rbk(x.text), i.find_all("td")))
                      for i in data_lines]
        data_dict = dict(data_pairs)

        self.title_zh_CN = rbk(mete_h1.text)
        self.title_en = rbk(mete_h1.find_next_sibling("h2").text)
        self.school = data_dict["学院（系）"]
        self.major = data_dict["专业"]
        self.name = data_dict["学生姓名"]
        self.number = data_dict["学号"]
        self.teacher = data_dict["指导教师"]
        self.auditor = data_dict["评阅教师"]
        self.finish_date = data_dict["完成日期"]

    def _set_contents(self):
        self.block = word.Metadata()

        self.block.title_zh_CN = self.title_zh_CN
        self.block.title_en = self.title_en
        self.block.school = self.school
        self.block.major = self.major
        self.block.name = self.name
        self.block.number = self.number
        self.block.teacher = self.teacher
        self.block.auditor = self.auditor
        self.block.finish_date = self.finish_date


class MainPart(PaperPart):
    def get_contents(self, soup: BeautifulSoup):
        main_h1 = soup.find("h1", string=re.compile("正文"))
        conts = get_content_until(main_h1.next_sibling,
                                  soup.find("h1", string=re.compile("结论")))
        self.contents = conts

    def _set_contents(self):
        self.block = word.MainContent()
        self._set_body()


class GraduationPaper(Paper):
    def __init__(self):
        self.meta = MetaPart()
        self.main = MainPart()

    def get_contents(self):
        self.meta.get_contents(self.soup)    # metadata
        self.abs = get_abs(self.soup)          # 摘要 Abstract
        # 目录 pass
        self.intro = get_intro(self.soup)      # 引言
        self.main.get_contents(self.soup)      # 正文
        self.conc = get_conclusion(self.soup)  # 结论
        self.ref = get_reference(self.soup)    # 参考文献
        self.appen = get_appendix(self.soup)   # 附录
        self.record = get_record(self.soup)    # 修改记录
        self.thanks = get_thanks(self.soup)    # 致谢

    def compile(self):
        self.main.contents = get_index(self.main.contents)

    def render(self, doc_path: str, out_path: str):
        doc = docx.Document(doc_path)
        word.DM.set_doc(doc)

        self.meta.render()  # metadata
        self.abs.render_template()   # 摘要 Abstract
        self.intro.render_template()  # 引言
        self.main.render()  # 正文
        self.conc.render_template()  # 结论
        # self.ref.render_template()  # 参考文献
        # self.appen.render_template()  # 附录
        # self.record.render_template()  # 修改记录
        self.thanks.render_template()  # 致谢

        word.DM.update_toc()
        doc.save(out_path)


if __name__ == "__main__":
    paper = GraduationPaper()
    paper.load_md("论文模板.md")
    paper.get_contents()
    paper.compile()

    paper.render("毕业设计（论文）模板-docx.docx", "out.docx")

'''
("h1", "something")
("h2", "something")
("h3", "something")

("fh4", like p)
("fh5", like p)

("p", [("text",      "something"),
       ("strong",    "something"),
       ("strong-em", "something"),
       ("em",        "something")])
("img",     (title, src))
("table",   (title, [Row]))
("formula", (title, "somthing"))
'''
