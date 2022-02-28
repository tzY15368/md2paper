from md2paper.md_paper import *


# 论文模块

class MetaPart(PaperPart):
    def load_contents(self, soup: BeautifulSoup):
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

    def _block_load_contents(self):
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


class AbsPart(PaperPart):
    def load_contents(self, soup: BeautifulSoup):
        # 摘要
        abs_cn_h1 = soup.find("h1", string=re_space("摘要"))
        abs_cn_ul = abs_cn_h1.find_next_sibling("ul")
        conts_cn = self._get_content_until(abs_cn_h1.next_sibling, abs_cn_ul)
        assert_warning(conts_cn[-1] == ("p", [{"type": "text", "text": "关键词："}]),
                       '摘要应该以"关键词："后接关键词列表结尾')
        self.conts_zh_CN = conts_cn[:-1]
        self.keywords_zh_CN = [rbk(i.text)
                               for i in abs_cn_h1.find_next_sibling("ul").find_all("li")]
        self.title_zh_CN = ""

        # Abstract
        abs_en_h1 = soup.find("h1", string=re_space("Abstract"))
        abs_en_ul = abs_en_h1.find_next_sibling("ul")
        conts_en = self._get_content_until(abs_en_h1.next_sibling, abs_en_ul)
        assert_warning(conts_en[-1] == ("p", [{"type": "text", "text": "Key Words:"}]),
                       'Abstract应该以"Key Words:"后接关键词列表结尾')
        self.conts_en = conts_en[:-1]
        self.keywords_en = [rbk(i.text)
                            for i in abs_en_h1.find_next_sibling("ul").find_all("li")]
        self.title_en = ""

    def _block_load_contents(self):
        self.block = word.Abstract()
        self.block.set_title(self.title_zh_CN,
                             self.title_en)
        self.block.add_text(assemble_ps(self.conts_zh_CN),
                            assemble_ps(self.conts_en))
        self.block.set_keyword(self.keywords_zh_CN,
                               self.keywords_en)


class IntroPart(PaperPart):
    def load_contents(self, soup: BeautifulSoup):
        intro_h1 = soup.find("h1", string=re_space("引言"))
        conts = self._get_content_until(intro_h1.next_sibling,
                                        soup.find("h1", string=re_space("正文")))
        self.contents = conts

    def _block_load_contents(self):
        self.block = word.Introduction()
        self.block.add_text(assemble_ps(self.contents))


class MainPart(PaperPart):
    def load_contents(self, soup: BeautifulSoup):
        main_h1 = soup.find("h1", string=re_space("正文"))
        conts = self._get_content_until(main_h1.next_sibling,
                                        soup.find("h1", string=re_space("结论")))
        self.contents = conts

    def _block_load_contents(self):
        self.block = word.MainContent()
        self._block_load_body()


class ConcPart(PaperPart):
    def load_contents(self, soup: BeautifulSoup):
        conclusion_h1 = soup.find("h1", string=re_space("结论"))
        if conclusion_h1 == None:
            conclusion_h1 = soup.find("h1", string=re_space("设计总结"))
        assert_error(conclusion_h1 != None, "应该有结论或设计总结")
        conts = self._get_content_until(conclusion_h1.next_sibling,
                                        soup.find("h1", string=re_space("参考文献")))
        self.contents = conts
        headline = rbk(conclusion_h1.text)
        assert_warning(headline in ["结论", "设计总结"],
                       "结论部分的标题应该是结论/设计总结: "+headline)
        if headline == "结论":
            headline = "结    论"
        self.headline = headline

    def _block_load_contents(self):
        self.block = word.Conclusion()
        self.block.add_text(assemble_ps(self.contents))

    def render(self):
        self._block_load_contents()
        self.block.render_template(self.headline)


class RefPart(PaperPart):
    def __init__(self):
        super().__init__()
        self.ref_map: Dict[str, str] = {}
        self.ref_list: List[str] = []

    def load_contents(self, soup: BeautifulSoup):
        reference_h1 = soup.find("h1", string=re_space("参考文献"))
        until_h1 = until_h1 = soup.find("h1", string=re.compile("^ *附录"))
        if until_h1 == None:
            until_h1 = soup.find("h1", string=re_space("修改记录"))

        self.bib_path = ""
        refs: List[str] = []

        cur = reference_h1.next_sibling
        while cur != until_h1:
            if cur.name != "p":
                cur = cur.next_sibling
                continue
            for i in cur.children:
                if i.name == "code":
                    text = i.text.split("\n")
                    if text[0] == "literature":
                        refs += text[1:]
                    elif text[0] == "bib":
                        bib_path = os.path.join(self.file_dir, text[1])
                        self.bib_path = bib_path
                    else:
                        log_error("这啥? " + i)
            cur = cur.next_sibling

        for ref_item in refs:
            pos = ref_item.find("]")
            assert_error(ref_item[0] == "[" and pos != -1,
                         "参考文献条目应该以 `[索引]` 开头: " + ref_item)
            ref = ref_item[1: pos]
            item = ref_item[pos+1:].strip()
            assert_warning(ref not in self.ref_map,
                           "参考文献索引不能重复: " + ref_item)
            self.ref_map[ref] = item

    def _block_load_contents(self):
        self.block = word.References()
        self._block_load_body()

    def _ref_get_author(self, data: Dict[str, str]) -> List[str]:
        if data["langid"] == "english":
            names = data["author"].split("and")
            authors = []
            for full_name in names:
                full_name = full_name.split(',')
                last_name = full_name[0].strip()
                name = full_name[1].strip().split(" ")
                name = [x[0] for x in name]
                name = reduce(lambda x, y: x+" "+y, name)
                sort_name = "{} {}".format(last_name, name)
                authors.append(sort_name)
            if len(authors) > 3:
                authors = authors[:3]
                authors.append("et al")
            author = reduce(lambda x, y: x+", "+y, authors)
        elif data["langid"] == "chinese":
            names = data["author"].split("and")
            authors = []
            for full_name in names:
                full_name = full_name.split(',')
                last_name = full_name[0].strip()
                name = full_name[1].strip()
                sort_name = "{}{}".format(last_name, name)
                authors.append(sort_name)
            if len(authors) > 3:
                authors = authors[:3]
                authors.append("等")
            author = reduce(lambda x, y: x+", "+y, authors)
        else:
            log_error("没做"+str(data))
        return author

    def _ref_get_entrytype(self, data: Dict[str, str]) -> str:
        type_map = {"book": "M",
                    "inbook": "M",
                    "inproceedings": "C",
                    "": "G",
                    "": "N",
                    "article": "J",
                    "phdthesis": "D",
                    "techreport": "R",
                    "misc": "S",
                    "patent": "P",
                    "": "DB",
                    "": "CP",
                    "": "EB",
                    }
        return type_map[data["ENTRYTYPE"]]

    def _ref_get_back(self, data: Dict[str, str]) -> str:
        back = ""
        if "address" in data and "publisher" in data:
            address = data["address"].replace("{", "").replace("}", "")
            publisher = data["publisher"].replace("{", "").replace("}", "")
            back = "{}: {}, ".format(address, publisher)
        return back

    def _ref_GB_T_7714_2005(self, data: Dict[str, str]) -> str:
        if not 'langid' in data:
            logging.warning(f"参考文献应该有语言信息: {str(data['title'])}，此处默认英文")
            data["langid"] = 'english'
        
        langid = data["langid"]
        author = self._ref_get_author(data)
        title = data["title"].replace("{", "").replace("}", "")
        entrytype = self._ref_get_entrytype(data)
        year = data["year"].replace("{", "").replace("}", "")
        back = self._ref_get_back(data)

        if langid == "english":
            ref_item = "{}. {} [{}]. {}{}.".format(
                author, title, entrytype, back, year)
        elif langid == "chinese":
            ref_item = "{}. {}[{}]. {}{}.".format(
                author, title, entrytype, back, year)
        else:
            log_error("没做"+str(data))
        return ref_item

    def _load_bib(self) -> Dict[str, str]:
        if self.bib_path == "":
            return {}
        with open(self.bib_path) as bibtex_file:
            parser = BibTexParser(common_strings=True)
            bib_database = bibtexparser.load(bibtex_file, parser=parser)
        ref_map = {}
        for item in bib_database.entries:
            ref_map["@"+item["ID"]] = self._ref_GB_T_7714_2005(item)
        return ref_map

    def compile(self):
        ref_map = self._load_bib()
        for ref in ref_map:
            assert_warning(ref not in self.ref_map,
                           "参考文献索引不能重复: " + ref)
            self.ref_map[ref] = ref_map[ref]

    def filt_ref(self, ref_items: Dict[str, RefItem]):
        ali_list = [(int(ref_items[ali].index), ali)
                    for ali in ref_items
                    if ref_items[ali].type == RefItem.LITER]
        ali_list.sort()
        self.ref_list = []
        for index, ali in ali_list:
            assert_warning(ali in self.ref_map,
                           "引用的文献应该在参考文献中出现: " + ali +
                           " BibTeX_path: '" + self.bib_path + "'")
            if ali in self.ref_map:
                self.ref_list.append(
                    "[{}] {}".format(index, self.ref_map[ali]))
        self.contents = [("p", [{"type": "text", "text": text}])
                         for text in self.ref_list]


class AppenPart(PaperPart):
    class AppenOne:
        def __init__(self, title: str, conts):
            self.title = title
            self.contents = conts

    def __init__(self):
        super().__init__()
        self.appens: List[self.AppenOne] = []

    def load_contents(self, soup: BeautifulSoup):
        appendix_h1s = soup.find_all("h1", string=re.compile("^ *附录"))
        appendix_h1s.append(soup.find("h1", string=re_space("修改记录")))
        appens = []
        for i in range(0, len(appendix_h1s)-1):
            conts = self._get_content_until(appendix_h1s[i].next_sibling,
                                            appendix_h1s[i+1])
            title = self._process_title(appendix_h1s[i].text, i)
            appens.append(self.AppenOne(title, conts))
        self.appens = appens

    def _block_load_contents(self):
        self.block = word.Appendixes()
        for appen in self.appens:
            self.block.add_appendix(appen.title)
            self._block_load_body(appen.contents)

    def _process_title(self, title: str, index: int):
        assert_warning(title[:2] == "附录", "附录应该以附录和编号开头: " + title)
        if title[2] == ' ':
            title = title[:2] + title[3:]
        assert_warning(title[2] == chr(ord("A") + index),
                       "附录应该以大写字母顺序编号: " + title)
        assert_warning(title[3] == " " and title[4] != " ",
                       "MD 中附录编号后应该有一个空格: " + title)
        title = title[:3] + "  " + rbk(title[4:])
        return title

    def get_ref_items(self):
        ref_items_list = [self._get_ref_items(appen.contents, appen.title[2])
                          for appen in self.appens]
        return ref_items_list_unfold(ref_items_list)


class RecordPart(PaperPart):
    def load_contents(self, soup: BeautifulSoup):
        mod_record_h1 = soup.find("h1", string=re_space("修改记录"))
        conts = self._get_content_until(mod_record_h1.next_sibling,
                                        soup.find("h1", string=re_space("致谢")))
        self.contents = conts

    def _block_load_contents(self):
        self.block = word.ChangeRecord()
        self._block_load_body()


class ThanksPart(PaperPart):
    def load_contents(self, soup: BeautifulSoup):
        thanks_h1 = soup.find("h1", string=re_space("致谢"))
        self.contents = self._get_content_from(thanks_h1.next_sibling)

    def _block_load_contents(self):
        self.block = word.Acknowledgments()
        self.block.add_text(assemble_ps(self.contents))


# 论文

class GraduationPaper(Paper):
    def __init__(self):
        super().__init__()
        self.meta = MetaPart()
        self.abs = AbsPart()
        self.intro = IntroPart()
        self.main = MainPart()
        self.conc = ConcPart()
        self.ref = RefPart()
        self.appen = AppenPart()
        self.record = RecordPart()
        self.thanks = ThanksPart()

        self.parts: List[PaperPart] = [
            self.meta,
            self.abs,
            self.intro,
            self.main,
            self.conc,
            self.ref,
            self.appen,
            self.record,
            self.thanks
        ]

    def compile(self):
        super().compile()

        self.abs.title_zh_CN = self.meta.title_zh_CN
        self.abs.title_en = self.meta.title_en

        ref_items_list = [
            self.main.get_ref_items(),
            self.appen.get_ref_items()
        ]
        self.ref_items = ref_items_list_unfold(ref_items_list)
        liter_cnt = 0
        for part in self.parts:
            liter_cnt = part.link_ref(self.ref_items, liter_cnt)
        self.ref.filt_ref(self.ref_items)
