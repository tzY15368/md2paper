import markdown
from bs4 import BeautifulSoup, Comment
import logging
import re


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
        if cur.name == "table":
            conts.append(("table", "something"))  # FIXME
        elif cur.name != None:
            conts.append((cur.name, rbk(cur.text)))
        cur = cur.next_sibling
    print(conts[0])
    return conts


# 获得每个论文模块

def get_metadata(soup: BeautifulSoup):
    mete_h1 = soup.find("h1")
    title_cn = rbk(mete_h1.text)  # TODO
    title_en = rbk(mete_h1.find_next_sibling("h2").text)  # TODO
    data_table = mete_h1.find_next_sibling("table").find("tbody")
    data_lines = data_table.find_all("tr")
    data_pairs = [list(map(lambda x: rbk(x.text), i.find_all("td")))
                  for i in data_lines]
    data_dict = dict(data_pairs)  # TODO
    return "tmp"  # TODO


def get_abs(soup: BeautifulSoup):
    # 摘要
    abs_cn_h1 = soup.find("h1", string=re.compile("摘要"))
    ps_key = get_ps(abs_cn_h1)
    assert_warning(ps_key[-1] == "关键词：", '摘要应该以"关键词："后接关键词列表结尾')
    ps = ps_key[:-1]  # TODO
    keywords = [rbk(i.text)
                for i in abs_cn_h1.find_next_sibling("ul").find_all("li")]  # TODO

    # Abstract
    abs_h1 = soup.find("h1", string=re.compile("Abstract"))
    ps_key = get_ps(abs_h1)
    assert_warning(ps_key[-1] == "Key Words:",
                   'Abstract应该以"Key Words:"后接关键词列表结尾')
    ps = ps_key[:-1]  # TODO
    keywords = [rbk(i.text)
                for i in abs_h1.find_next_sibling("ul").find_all("li")]  # TODO
    return "tmp"  # TODO


def get_intro(soup: BeautifulSoup):
    intro_h1 = soup.find("h1", string=re.compile("引言"))
    ps = get_ps(intro_h1)  # TODO
    return "tmp"  # TODO


def get_body(soup: BeautifulSoup):
    body_h1 = soup.find("h1", string=re.compile("正文"))
    content = get_content(body_h1,
                          soup.find("h1", string=re.compile("结论")))
    print(body_h1)  # FIXME
    return "tmp"  # TODO


def get_conclusion(soup: BeautifulSoup):
    conclusion_h1 = soup.find("h1", string=re.compile("结论"))
    ps = get_ps(conclusion_h1)  # TODO
    return "tmp"  # TODO


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
    ps = get_ps(thanks_h1)  # TODO
    return "tmp"  # TODO


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
    md = load_md("论文模板.md", "论文")
