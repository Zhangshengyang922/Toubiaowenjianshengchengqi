"""
响应文件模板提取器 v4 —— 直接从招标文件DOCX中提取"响应文件格式"章节
使用XML级别替换占位符，完全保留原始格式不变

v4 改进：消除所有硬编码位置假设，适应任意招标文件结构
- 章节定位：全文搜索，取最后匹配（避免目录干扰），不做段落索引跳过
- 封面检测：自动感知格式分界（第二个独立格式标记 = 封面结束）
- 目录生成：通用正则匹配格式/附件/表格等任意格式标识
- 目录分组：自动按前缀分组，不硬编码"格式1-/格式2-"
"""
import os
import re
import copy
from lxml import etree
from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# Word文档XML命名空间
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# ── 通用正则：匹配章节标题 + 格式编号，不硬编码章节号/名称 ──
# 要求段落以"第X章"开头（真正的章节标题），限制间距 .{0,30}? 防止误匹配
_CHAPTER_PATTERN = re.compile(
    r'^\s*第[一二三四五六七八九十\d]+章.{0,30}?(?:响应文件格式|投标文件格式|格式要求|响应文件)'
)

# 格式编号：格式X-X / 附件X / 附表X / 表格X / 附录X（支持全角半角连字符）
_FORMAT_MARKER_PATTERN = re.compile(
    r'(?:格式|附件|附表|表格|附录)\s*[零一二三四五六七八九十\d]+[-–—.\u2013\u2014\u2015]?[零一二三四五六七八九十\d]*'
)

# ── 用于封面截止检测：按顺序记录所有出现的独立格式标记 ──
_UNIQUE_FORMAT_RE = re.compile(
    r'((?:格式|附件|附表|表格|附录)\s*[零一二三四五六七八九十\d]+[-–—.\u2013\u2014\u2015]?[零一二三四五六七八九十\d]*)'
)


class ResponseTemplateExtractor:
    """从招标文件DOCX提取响应文件格式章节，填充后生成响应文件"""

    def __init__(self):
        self.paragraph_start = None
        self.paragraph_end = None
        self.chapter_keyword = None  # 实际匹配到的章节关键词（如"第七章 响应文件格式"）

    def generate(self, docx_path: str, fill_data: dict) -> str:
        """
        docx_path: 招标文件DOCX路径
        fill_data: 填充数据字典
        """
        doc = Document(docx_path)
        self._find_format_chapter(doc)
        if self.paragraph_start is None:
            raise ValueError(
                "未找到包含'响应文件格式'的章节。"
                "请确认招标文件中存在类似'第X章 响应文件格式'的标题。"
            )

        new_doc = self._build_new_document(doc, fill_data)

        output_dir = os.path.join(
            os.path.dirname(os.path.dirname(__file__)), 'output'
        )
        os.makedirs(output_dir, exist_ok=True)
        safe_name = re.sub(r'[\\/:*?"<>|]', '_', fill_data.get('project_name', 'response'))
        output_path = os.path.join(output_dir, f'{safe_name}_响应文件.docx')
        new_doc.save(output_path)
        return output_path

    # ==================== 定位格式章节（无硬编码位置假设） ====================

    def _find_format_chapter(self, doc):
        """
        全文搜索找到"响应文件格式"对应的章节范围。

        策略（无段落索引硬编码）：
        1. 扫描所有段落，匹配 "第X章" + "格式/响应文件" 模式（支持任意章节号）
        2. 取最后匹配的作为真实章节标题（目录中的引用总是出现在前面）
        3. 自动查找下一章作为结束边界
        """
        all_matches = []  # [(paragraph_index, chapter_number_text, full_text)]

        for i, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            # 匹配 "第X章" + 格式相关描述
            m = _CHAPTER_PATTERN.search(text)
            if m:
                # 提取章节数字用于寻找下一章
                chap_num_match = re.search(r'第([一二三四五六七八九十\d]+)章', text)
                chap_num = chap_num_match.group(1) if chap_num_match else ''
                all_matches.append((i, chap_num, text))

        if not all_matches:
            return

        # ── 取最后匹配的作为真实章节标题 ──
        # 原因：招标文件的目录中会先出现章节引用，真实章节内容在后面
        #      取最后一个匹配是最安全的策略
        last_idx, chap_num, full_text = all_matches[-1]
        self.paragraph_start = last_idx
        self.chapter_keyword = full_text[:60]

        # ── 查找下一章作为结束标记 ──
        next_chap_num = self._next_chapter_number(chap_num)
        if next_chap_num:
            for i, p in enumerate(doc.paragraphs):
                if i <= self.paragraph_start:
                    continue
                text = p.text.strip()
                if f'第{next_chap_num}章' in text:
                    self.paragraph_end = i
                    break

        # 如果找不到下一章，取文档末尾
        if self.paragraph_end is None:
            self.paragraph_end = len(doc.paragraphs)

        print(f"    [定位] 章节: {self.chapter_keyword} (段落{self.paragraph_start}~{self.paragraph_end})")

    @staticmethod
    def _next_chapter_number(chap_num: str) -> str | None:
        """计算下一章的数字（支持中文数字和阿拉伯数字）"""
        cn_map = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10,
                  '十一':11,'十二':12,'十三':13,'十四':14,'十五':15}
        num = None
        if chap_num.isdigit():
            num = int(chap_num)
        elif chap_num in cn_map:
            num = cn_map[chap_num]
        if num is not None:
            next_num = num + 1
            # 优先用中文数字
            rev_map = {v:k for k,v in cn_map.items()}
            return rev_map.get(next_num, str(next_num))
        return None

    # ==================== 构建新文档 ====================

    def _build_new_document(self, src_doc, fill_data):
        """按 body 元素顺序遍历，同时处理段落和表格"""
        new_doc = Document()

        for sec in new_doc.sections:
            sec.top_margin = Cm(2.54)
            sec.bottom_margin = Cm(2.54)
            sec.left_margin = Cm(3.17)
            sec.right_margin = Cm(3.17)

        body = new_doc.element.body
        src_body = src_doc.element.body

        # 第一步：按 body 元素顺序收集第七章的所有元素
        para_count = 0
        table_count = 0
        ch7_elements = []
        cover_end_index = 0  # 0 = 未检测到独立封面

        for child in src_body:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'p':
                if para_count < self.paragraph_start:
                    para_count += 1
                    continue
                if para_count >= self.paragraph_end:
                    break

                cloned = copy.deepcopy(child)
                self._replace_in_xml(cloned, fill_data)
                ch7_elements.append(cloned)
                para_count += 1

            elif tag == 'tbl':
                # 表格位于段落之间，判断是否在章节范围内
                if self.paragraph_start <= para_count < self.paragraph_end:
                    cloned = copy.deepcopy(child)
                    self._replace_in_xml(cloned, fill_data)
                    ch7_elements.append(cloned)
                    table_count += 1

        # ── 封面截止检测：找到第二个独立格式标记即为封面结束 ──
        cover_end_index = self._detect_cover_end(ch7_elements)

        print(f"    [调试] 共 {len(ch7_elements)} 个元素（含 {table_count} 个表格），封面截止索引: {cover_end_index}")

        # 第二步：封面 → 分页 → 目录 → 分页 → 正文
        # 2.1 添加封面
        for i in range(cover_end_index):
            body.append(ch7_elements[i])

        # 2.2 分页符 + 目录（仅当存在独立封面时插入目录）
        if cover_end_index > 0:
            self._add_section_break(body)
            self._add_toc(new_doc, fill_data, ch7_elements, cover_end_index)

        # 2.3 分页符 + 正文
        self._add_section_break(body)
        for i in range(cover_end_index, len(ch7_elements)):
            body.append(ch7_elements[i])

        return new_doc

    def _detect_cover_end(self, ch7_elements: list) -> int:
        """
        自动检测封面截止位置。

        策略：按顺序扫描元素，记录出现的独立格式标记。
        当找到第2个不同的格式标记时，之前的元素即为封面。
        这样无论格式编号是"格式1-1"、"附件一"还是其他命名，都能正确识别。

        返回封面元素数量；若未检测到独立封面则返回0。
        """
        seen_formats = []
        for i, elem in enumerate(ch7_elements):
            text = self._get_element_text(elem)
            m = _UNIQUE_FORMAT_RE.search(text)
            if m:
                fmt_key = m.group(1)
                # 归一化处理（去空格、统一连字符）
                normalized = re.sub(r'\s+', '', fmt_key).replace('\u2013','-').replace('\u2014','-').replace('\u2015','-')
                if normalized not in seen_formats:
                    seen_formats.append(normalized)
                    if len(seen_formats) >= 2:
                        return i  # 第2个格式标记之前 = 封面
        return 0  # 未检测到独立封面，不做封面/目录分离

    def _get_element_text(self, element):
        """获取XML元素中所有文本"""
        texts = []
        for t in element.findall(f'.//{{{W}}}t'):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)

    # ==================== 目录生成（无硬编码格式假设） ====================

    def _add_toc(self, new_doc, fill_data, ch7_elements, cover_end_index):
        """在封面后生成目录页"""
        body = new_doc.element.body

        # 收集格式标题及描述
        format_titles = []
        for i, elem in enumerate(ch7_elements):
            text = self._get_element_text(elem)
            # 匹配 "格式X-X" / "附件X" / "附表X" 等任意格式标记
            m = _FORMAT_MARKER_PATTERN.search(text)
            if m:
                title_key = m.group()
                if title_key not in [t[0] for t in format_titles]:
                    rest = text[m.end():].strip()
                    rest = re.sub(r'^[：:、，,\s]+', '', rest)
                    if not rest and i + 1 < len(ch7_elements):
                        next_text = self._get_element_text(ch7_elements[i + 1]).strip()
                        if next_text and not any(w in next_text for w in ['正本', '副本', '格式', '附件', '附表', '表格']):
                            rest = next_text
                    full_title = f'{title_key}  {rest}' if rest else title_key
                    format_titles.append((title_key, full_title))

        if not format_titles:
            return

        # 去重保持顺序
        seen = set()
        unique_titles = []
        for key, title in format_titles:
            if key not in seen:
                seen.add(key)
                unique_titles.append((key, title))

        # 写入目录标题
        toc_heading = self._make_paragraph(
            '目    录',
            font_name='宋体', font_size=Pt(16), bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER
        )
        body.append(toc_heading)
        body.append(self._make_paragraph(''))

        # ── 自动按前缀分组 ──
        # 提取每个标题的前缀（如 "格式1"、"格式2"、"附件一" 等）
        groups = self._group_titles(unique_titles)

        if len(groups) > 1:
            # 多组 → 分组显示
            group_names = {
                0: '第一部分',
                1: '第二部分',
                2: '第三部分',
                3: '第四部分',
                4: '第五部分',
            }
            for gi, group in enumerate(groups):
                group_label = group_names.get(gi, f'第{gi+1}部分')
                body.append(self._make_paragraph(
                    f'{group_label}',
                    font_name='黑体', font_size=Pt(14), bold=True,
                    left_indent=Cm(0.5)
                ))
                for key, title in group:
                    body.append(self._make_paragraph(
                        title,
                        font_name='仿宋', font_size=Pt(14),
                        left_indent=Cm(1.5)
                    ))
                body.append(self._make_paragraph(''))
        else:
            # 单组 → 直接列表
            for key, title in unique_titles:
                body.append(self._make_paragraph(
                    title,
                    font_name='仿宋', font_size=Pt(14),
                    left_indent=Cm(1.5)
                ))

    @staticmethod
    def _group_titles(titles: list) -> list:
        """
        按前缀自动分组。
        返回 [[(key, title), ...], ...]
        """
        if not titles:
            return [titles]

        prefix_re = re.compile(r'^((?:格式|附件|附表|表格|附录)\s*\d+)')
        groups = []
        current_prefix = None
        current_group = []

        for key, title in titles:
            pm = prefix_re.match(key)
            prefix = pm.group(1) if pm else key
            if current_prefix is not None and prefix != current_prefix:
                groups.append(current_group)
                current_group = []
            current_prefix = prefix
            current_group.append((key, title))

        if current_group:
            groups.append(current_group)

        return groups if groups else [titles]

    def _make_paragraph(self, text, font_name='仿宋', font_size=None, bold=False,
                        alignment=None, left_indent=None):
        """创建格式化的段落"""
        from docx.oxml import OxmlElement

        p = OxmlElement('w:p')

        # 段落属性
        pPr = OxmlElement('w:pPr')
        if alignment is not None:
            jc = OxmlElement('w:jc')
            align_map = {
                WD_ALIGN_PARAGRAPH.CENTER: 'center',
                WD_ALIGN_PARAGRAPH.LEFT: 'left',
                WD_ALIGN_PARAGRAPH.RIGHT: 'right',
            }
            jc.set(qn('w:val'), align_map.get(alignment, 'left'))
            pPr.append(jc)
        if left_indent:
            ind = OxmlElement('w:ind')
            ind.set(qn('w:left'), str(int(left_indent / Cm(1) * 567)))  # EMU
            pPr.append(ind)
        p.append(pPr)

        # 文本run
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        if font_name:
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:eastAsia'), font_name)
            rFonts.set(qn('w:ascii'), font_name)
            rPr.append(rFonts)
        if font_size:
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), str(int(font_size.pt * 2)))  # half-points
            rPr.append(sz)
        if bold:
            b = OxmlElement('w:b')
            rPr.append(b)
        r.append(rPr)

        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = text
        r.append(t)
        p.append(r)

        return p

    def _add_section_break(self, body):
        """添加分页符"""
        from docx.oxml import OxmlElement
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        p.append(r)
        body.append(p)

    # ==================== XML级别占位符替换 ====================

    def _replace_in_xml(self, element, fill_data: dict):
        """在XML元素中递归替换所有文本节点中的占位符"""
        text_elements = element.findall(f'.//{{{W}}}t')
        full_text = ''
        text_elems = []
        for t_elem in text_elements:
            text = t_elem.text or ''
            full_text += text
            text_elems.append((t_elem, text))

        if not full_text.strip():
            return

        new_full = self._apply_replacements(full_text, fill_data)
        if new_full == full_text:
            return

        for idx, (t_elem, _) in enumerate(text_elems):
            if idx == 0:
                t_elem.text = new_full
            else:
                t_elem.text = ''

    def _apply_replacements(self, text: str, fill: dict) -> str:
        """对所有占位符进行替换"""
        bidder = fill.get('bidder_name', '') or ''
        legal = fill.get('legal_representative', '') or ''
        authorized = fill.get('authorized_person', '') or ''
        agency = fill.get('agency_name', '') or ''
        proj_name = fill.get('project_name', '') or ''
        proj_id = fill.get('project_id', '') or ''
        date_val = fill.get('bid_opening_date', '') or ''
        address = fill.get('address', '') or ''
        phone = fill.get('phone', '') or ''
        zip_code = fill.get('zip_code', '') or ''
        fax = fill.get('fax', '') or ''

        t = text

        # ==== 项目名称: xxxx项目 / XX项目 ====
        t = re.sub(r'[xX]{2,4}\s*项目', proj_name, t)
        t = re.sub(r'[xX]{2}\s*项目', proj_name, t)

        # ==== 引号内项目名称 ====
        t = re.sub(r'"X{3,10}"', f'"{proj_name}"', t)
        t = re.sub(r"'X{3,10}'", f"'{proj_name}'", t)
        t = re.sub(r'"[xX]{3,10}"', f'"{proj_name}"', t)
        t = re.sub(r'\u201c[xX]{3,10}\u201d', f'\u201c{proj_name}\u201d', t)

        # ==== 项目编号 ====
        if proj_id:
            t = t.replace('（项目编号：XXXX）', f'（项目编号：{proj_id}）')
            t = t.replace('(项目编号：XXXX)', f'(项目编号：{proj_id})')
            t = t.replace('项目编号：XXXX', f'项目编号：{proj_id}')

        # ==== 日期 ====
        if date_val:
            t = re.sub(r'XXX+年XXX+月XXX+日', date_val, t)
            t = re.sub(r'XX年XX月XX日', date_val, t)
            t = re.sub(r'日\s*期[：:]\s*XXX+年XXX+月XXX+日', f'日    期：{date_val}', t)
            t = re.sub(r'时间[：:]\s*XX年XX月XX日', f'时间：{date_val}', t)
            t = re.sub(r'日\s*期[：:]\s*[xX]{3,4}\s*[。]?', f'日    期：{date_val}', t)
            t = re.sub(r'日期[：:]\s*[xX]{3,4}\s*[。]?', f'日期：{date_val}', t)
            t = re.sub(r'日期[：:]\s*[xX]{3,4}\s*$', f'日期：{date_val}', t)

        # ==== 采购代理机构 ====
        if agency:
            t = re.sub(r'[xX]{3,8}\s*[（\(]采购代理机构名称[）\)]\s*[：:]', f'{agency}：', t)
            t = re.sub(r'[xX]{3,4}\s*[（\(]采购代理机构名称[）\)]\s*[：:]', f'{agency}：', t)

        # ==== 供应商名称：XXXX/XXX ====
        if bidder:
            t = re.sub(
                r'供应商名称[：:]\s*[xX]{3,4}\s*[（(]\s*单位[公盖]\s*[章]*\s*[）)][。.]?',
                f'供应商名称：{bidder}（盖章）。',
                t
            )
            t = re.sub(
                r'供应商名称[：:]\s*[xX]{3,4}\s*[（(]\s*盖\s*[章]*\s*[）)]',
                f'供应商名称：{bidder}（盖章）',
                t
            )
            t = re.sub(
                r'供应商名称[：:]\s*[xX]{3,4}\s*[（(]盖单位公章[）)]',
                f'供应商名称：{bidder}（盖章）',
                t
            )
            t = re.sub(
                r'(供\s*应\s*商名称[：:])\s*$',
                f'\\1{bidder}',
                t
            )

        # ==== 法定代表人/单位负责人或授权代表 ====
        if legal:
            t = re.sub(
                r'(法定代表人/单位负责人或授权代表)\s*[（\(]签字或加盖个人印章[）\)]\s*[：:]\s*[xX]{3,4}',
                f'\\1：{legal}',
                t
            )
            t = re.sub(
                r'(法定代表人/单位负责人或授权代表)\s*[（\(]签字或加盖个人印章[）\)]\s*[：:]\s*XXX+',
                f'\\1：{legal}',
                t
            )

        # ==== 授权书内容替换 ====
        if bidder and legal:
            t = re.sub(
                r'\s[Xx]{4}\s*[（\(]供应商名称[）\)]\s*[Xx]{4}\s*[（\(]法定代表人',
                f' {bidder}（供应商名称）{legal}（法定代表人',
                t
            )

        if authorized:
            t = re.sub(
                r'授权\s*[Xx]{4}\s*[（\(]被授权人',
                f'授权{authorized}（被授权人',
                t
            )

        # ==== 法定代表人/单位负责人签字 ====
        if legal:
            t = re.sub(
                r'(法定代表人/单位负责人（委托人）\s*签字或加盖个人印章[：:]\s*)[Xx]{3,4}',
                f'\\1{legal}',
                t
            )

        if authorized:
            t = re.sub(
                r'(授权代表（被授权人）\s*签字[：:]\s*)[Xx]{3,4}',
                f'\\1{authorized}',
                t
            )

        # ==== 通讯地址、邮编、电话、传真 ====
        if address:
            t = re.sub(r'通讯地址[：:]\s*[xX]{3,4}\s*$', f'通讯地址：{address}', t)
            t = re.sub(r'通讯地址[：:]\s*[xX]{3,4}(?:\s|$)', f'通讯地址：{address}', t)
        if zip_code:
            t = re.sub(r'邮政编码[：:]\s*[xX]{3,4}', f'邮政编码：{zip_code}', t)
        if phone:
            t = re.sub(r'联系电话[：:]\s*[xX]{3,4}[xX\-]*\s*$', f'联系电话：{phone}', t)
            t = re.sub(r'联系电话[：:]\s*[xX]{3,4}(?:\s|$)', f'联系电话：{phone}', t)
        if fax:
            t = re.sub(r'传\s*真[：:]\s*[xX]{3,4}', f'传    真：{fax}', t)

        # ==== 包号 ====
        t = t.replace('包        号：', f'包        号：{fill.get("package_no", "1")}')

        # ==== 采购项目编号：空值补填 ====
        if proj_id:
            t = re.sub(r'(采购项目编号[：:])\s*$', f'\\1{proj_id}', t)

        # ==== 最终兜底 ====
        t = re.sub(r'"[xX]{4,10}"', f'"{proj_name}"', t)
        t = re.sub(r'\u201c[xX]{4,10}\u201d', f'\u201c{proj_name}\u201d', t)

        return t


def generate_response_docx(bidding_docx_path: str, fill_data: dict) -> str:
    """便捷函数"""
    extractor = ResponseTemplateExtractor()
    return extractor.generate(bidding_docx_path, fill_data)
