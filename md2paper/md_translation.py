from md2paper.md_paper import *

import md2paper.dut_paper_translation as transword


# 论文模块

class TranslationPart(PaperPart):
    def _split_title(self, title: str):
        sp = title.split(';')
        if len(sp) == 1:
            ali = ""
            title = rbk(sp[0])
            ratio = 0
        elif len(sp) == 2:
            ali = ""
            title = rbk(sp[0])
            ratio_s = rbk(sp[1])
            if ratio_s != "":
                ratio = int(ratio_s[:-1])
            else:
                ratio = 0
            assert_warning(0 <= ratio <= 100,
                           "图片占页面宽度应该在 [0%, 100%] 间: " + title)
        else:
            log_error("图、表的标题格式错误: " + title)

        return ali, title, ratio/100


class TransMetaPart(TranslationPart):
    def load_contents(self, soup: BeautifulSoup):
        mete_h1 = soup.find("h1")

        # 个人信息
        data_table = mete_h1.find_next_sibling("table").find("tbody")
        data_lines = data_table.find_all("tr")
        data_pairs = [list(map(lambda x: rbk(x.text), i.find_all("td")))
                      for i in data_lines]
        data_dict = dict(data_pairs)

        self.title_zh_CN = rbk(mete_h1.text)
        self.title_en = rbk(mete_h1.find_next_sibling("h2").text)
        self.school = data_dict["学部（院）"]
        self.major = data_dict["专业"]
        self.name = data_dict["学生姓名"]
        self.number = data_dict["学号"]
        self.teacher = data_dict["指导教师"]
        self.finish_date = data_dict["完成日期"]

        # 外文作者信息
        data_table = mete_h1.find_next_siblings("table")[1].find("tbody")
        data_lines = data_table.find_all("tr")
        data_pairs = [list(map(lambda x: rbk(x.text), i.find_all("td")))
                      for i in data_lines]
        data_dict = dict(data_pairs)

        self.author = data_dict["author"]
        self.organization = data_dict["工作单位"]

    def _block_load_contents(self):
        self.block = transword.TranslationMetadata()

        self.block.title_zh_CN = self.title_zh_CN
        self.block.title_en = self.title_en
        self.block.school = self.school
        self.block.major = self.major
        self.block.name = self.name
        self.block.number = self.number
        self.block.teacher = self.teacher
        self.block.finish_date = self.finish_date


class TransAbsPart(TranslationPart):
    def load_contents(self, soup: BeautifulSoup):
        # 摘要
        abs_cn_h1 = soup.find("h1", string=re_space("摘要"))
        if abs_cn_h1 == None:
            self.conts_zh_CN = None
            return
        abs_cn_ul = abs_cn_h1.find_next_sibling("ul")
        conts_cn = self._get_content_until(abs_cn_h1.next_sibling, abs_cn_ul)
        assert_warning(conts_cn[-1] == ("p", [{"type": "text", "text": "关键词："}]),
                       '摘要应该以"关键词："后接关键词列表结尾')
        self.conts_zh_CN = conts_cn[:-1]
        self.keywords_zh_CN = [rbk(i.text)
                               for i in abs_cn_h1.find_next_sibling("ul").find_all("li")]
        self.title_zh_CN = ""
        self.author = ""
        self.organization = ""

    def _block_load_contents(self):
        self.block = transword.TranslationAbstract(self.title_zh_CN,
                                                   self.author,
                                                   self.organization)
        self._block_load_body(self.conts_zh_CN)
        self.block.add_keywords(self.keywords_zh_CN)


class TransMainPart(TranslationPart):
    def load_contents(self, soup: BeautifulSoup):
        main_h1 = soup.find("h1", string=re_space("正文"))
        conts = self._get_content_from(main_h1.next_sibling)
        self.contents = conts

    def _link_ref(self) -> int:
        for name, cont in self.contents:
            if name not in ["p", "fh4", "fh5"]:
                continue
            is_text = False
            for run in cont:
                if run["type"] != "ref":
                    if run["type"] == "text" and run["text"].endswith("文献"):
                        is_text = True
                    continue

                if is_text:
                    run["type"] = "text"
                else:
                    run["type"] = "ref"
                run["text"] = "[{}]".format(run["text"])

                is_text = False

    def compile(self):
        self._link_ref()

        self.contents.insert(-2, ("p", []))

    def _block_load_contents(self):
        self.block = transword.TranslationMainContent()
        self._block_load_body()


# 论文

class TranslationPaper(Paper):
    def __init__(self):
        super().__init__()
        self.meta = TransMetaPart()
        self.abs = TransAbsPart()
        self.main = TransMainPart()

        self.parts: List[PaperPart] = [
            self.meta,
            self.abs,
            self.main
        ]

    def compile(self):
        super().compile()

        self.abs.author = self.meta.author
        self.abs.organization = self.meta.organization
        self.abs.title_zh_CN = self.meta.title_zh_CN
