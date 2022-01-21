import markdown
from bs4 import BeautifulSoup, Comment
import logging
import re


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


def rbk(s: str):  # remove_blank
    return s


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

        cur = cur.next_sibling
    return conts


def handle_paper(soup: BeautifulSoup):
    # test
    with open("out.html", "w") as f:
        f.write(soup.prettify())

    # metadata
    mete_h1 = soup.find("h1")
    title_cn = rbk(mete_h1.text)  # TODO
    title_en = rbk(mete_h1.find_next_sibling("h2").text)  # TODO
    data_table = mete_h1.find_next_sibling("table").find("tbody")
    data_lines = data_table.find_all("tr")
    data_pairs = [list(map(lambda x: rbk(x.text), i.find_all("td")))
                  for i in data_lines
                  ]
    data_dict = dict(data_pairs)  # TODO

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

    # 目录
    pass

    # 引言
    intro_h1 = soup.find("h1", string=re.compile("引言"))
    ps = get_ps(intro_h1)  # TODO

    # 正文
    body_h1 = soup.find("h1", string=re.compile("正文"))
    content = get_content(body_h1,
                          soup.find("h1", string=re.compile("结论")))
    print(body_h1)  # FIXME

    # 结论
    conclusion_h1 = soup.find("h1", string=re.compile("结论"))
    ps = get_ps(conclusion_h1)  # TODO

    # 参考文献
    reference_h1 = soup.find("h1", string=re.compile("参考文献"))
    # 需要一个专门的处理方式
    print(reference_h1)  # FIXME

    # 附录
    appendix_h1s = soup.find_all("h1", string=re.compile("附录"))
    content = get_content(appendix_h1s,
                          soup.find("h1", string=re.compile("修改记录")))
    print(appendix_h1s)  # FIXME

    # 修改记录
    mod_record_h1 = soup.find("h1", string=re.compile("修改记录"))
    content = get_content(mod_record_h1,
                          soup.find("h1", string=re.compile("致谢")))
    print(mod_record_h1)  # FIXME

    # 致谢
    thanks_h1 = soup.find("h1", string=re.compile("致谢"))
    ps = get_ps(thanks_h1)  # TODO


def handle_translation(soup: BeautifulSoup):
    pass


def load_md(file_name: str, file_type: str):
    with open(file_name, "r") as f:
        md_file = f.read()
    md_html = markdown.markdown(md_file,
                                extensions=[
                                    'markdown.extensions.tables'
                                ]
                                )
    soup = BeautifulSoup(md_html, 'html.parser')
    for i in soup(text=lambda text: isinstance(text, Comment)):
        i.extract()

    if file_type == "论文":
        return handle_paper(soup)
    elif file_type == "外文翻译":
        return handle_translation(soup)
    else:
        log_error('错误的文件类型，应该选择 "论文" / "外文翻译"')


if __name__ == "__main__":
    md = load_md("论文模板.md", "论文")
