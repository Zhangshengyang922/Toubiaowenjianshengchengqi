"""
投标文件生成器 v2 —— 严格按照招标文件要求生成投标文件
- 封面：按招标文件封面要求，自动填入项目名称、编号、开标日期
- 顺序：按招标文件要求的文档组成顺序排列
- 日期：所有日期使用开标日期
"""
import json
import os
from typing import Dict, Optional, List


class BidResponseGenerator:

    def __init__(self, company_profile_path: str = None):
        if company_profile_path is None:
            company_profile_path = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                'data', 'company_profile.json'
            )
        with open(company_profile_path, 'r', encoding='utf-8') as f:
            company_data = json.load(f)
        self.company = company_data.get('company', {})
        self.qualifications = company_data.get('qualifications', [])
        self.past_projects = company_data.get('past_projects', [])
        self.team_members = company_data.get('team_members', [])

        self.llm_client = None
        self._llm_enabled = False

    def enable_llm(self, api_key: str, base_url: str = None):
        try:
            import openai
            openai.api_key = api_key
            if base_url:
                openai.api_base = base_url
            self.llm_client = openai
            self._llm_enabled = True
        except ImportError:
            print("警告: openai 未安装")

    # ==================== 主入口 ====================

    def generate(self, bidding_doc_info: Dict, extra_input: Dict = None) -> Dict:
        """
        生成投标文件，返回有序章节结构
        """
        project = dict(bidding_doc_info.get('project_info', {}))
        requirements = bidding_doc_info.get('requirements', {})
        raw_text = bidding_doc_info.get('raw_text', '')
        required_docs = bidding_doc_info.get('required_documents', [])
        cover_reqs = bidding_doc_info.get('cover_requirements', {})
        format_reqs = bidding_doc_info.get('format_requirements', {})

        if extra_input:
            project.update(extra_input)

        # ----- 确定日期：开标日期优先 -----
        bid_date = (
            project.get('bid_opening_date') or
            project.get('submission_date') or
            project.get('bid_opening_time_raw') or
            '_____年___月___日'
        )

        # 存储核心变量
        ctx = {
            'project_name': project.get('project_name', '________'),
            'project_id': project.get('project_id', '________'),
            'bidder_name': project.get('bidder_name', self.company.get('bidder_name', '________')),
            'tenderee_name': project.get('tenderee_name', '贵单位'),
            'bid_date': bid_date,
            'bid_amount': project.get('bid_amount', '________'),
            'implementation_period': project.get('implementation_period', '________'),
            'warranty_period': project.get('warranty_period', '________'),
            'validity_days': requirements.get('validity_days', '90'),
            'legal_representative': project.get('legal_representative', self.company.get('legal_representative', '________')),
            'authorized_person': project.get('authorized_person', self.company.get('authorized_person', '________')),
        }

        # ----- 按招标文件要求的顺序生成章节 -----
        document = {}
        chapter_funcs = self._get_chapter_funcs()

        for doc_desc in required_docs:
            chapter_key, content = self._match_chapter(doc_desc, chapter_funcs, ctx, project, requirements, raw_text, cover_reqs, format_reqs)
            if content:
                document[chapter_key] = {
                    'title': doc_desc.strip(),
                    'content': content,
                }

        # 确保核心章节都存在
        self._ensure_core_chapters(document, chapter_funcs, ctx, project, requirements, raw_text, cover_reqs, format_reqs)

        return document

    def _get_chapter_funcs(self) -> Dict:
        """章节关键词 -> (key, 生成函数) 映射"""
        return {
            '封面':     ('cover', self._ch_cover),
            '投标函':   ('bid_letter', self._ch_bid_letter),
            '一览表':   ('bid_summary', self._ch_bid_summary),
            '开标':     ('bid_summary', self._ch_bid_summary),
            '授权':     ('authorization', self._ch_authorization),
            '法定代表人': ('authorization', self._ch_authorization),
            '资格':     ('qualifications', self._ch_qualifications),
            '资质':     ('qualifications', self._ch_qualifications),
            '技术方案': ('technical', self._ch_technical),
            '技术':     ('technical', self._ch_technical),
            '实施方案': ('implementation', self._ch_implementation),
            '实施':     ('implementation', self._ch_implementation),
            '业绩':     ('performance', self._ch_performance),
            '案例':     ('performance', self._ch_performance),
            '报价':     ('bid_summary', self._ch_bid_summary),
            '其他':     ('appendices', self._ch_appendices),
            '附件':     ('appendices', self._ch_appendices),
            '附录':     ('appendices', self._ch_appendices),
        }

    def _match_chapter(self, desc: str, funcs, ctx, project, reqs, raw, cover, fmt):
        """根据文档描述匹配章节"""
        desc_lower = desc.lower()
        for keyword, (key, func) in funcs.items():
            if keyword in desc_lower:
                return (key, func(ctx, project, reqs, raw, cover, fmt))
        # 未匹配 → 生成通用章节
        return (f'custom_{hash(desc) % 10000}', self._ch_generic(desc, ctx))

    def _ensure_core_chapters(self, doc, funcs, ctx, project, reqs, raw, cover, fmt):
        """确保核心章节不缺失"""
        core = [
            ('cover', '封面', funcs['封面'][1]),
            ('bid_letter', '投标函', funcs['投标函'][1]),
            ('bid_summary', '开标一览表', funcs['开标'][1]),
            ('authorization', '法定代表人授权委托书', funcs['法定代表人'][1]),
        ]
        for key, title, func in core:
            if key not in doc:
                doc[key] = {
                    'title': title,
                    'content': func(ctx, project, reqs, raw, cover, fmt),
                }

    # ==================== 各章节生成 ====================

    def _ch_cover(self, ctx, project, reqs, raw, cover_reqs, fmt):
        """封面 —— 严格按招标文件封面要求"""

        project_name = ctx['project_name']
        project_id = ctx['project_id']
        bidder = ctx['bidder_name']
        bid_date = ctx['bid_date']
        tenderer = ctx['tenderee_name']

        # 招标文件要求的封面内容
        cover_items = cover_reqs.get('required_items', [])
        copies = cover_reqs.get('copies', '')

        lines = []

        # 正本/副本标注
        if copies:
            lines.append(copies)
        elif '正本' in str(cover_items):
            lines.append("正本／副本")
        else:
            lines.append("正本／副本")

        lines.append("")
        lines.append("=" * 52)
        lines.append("")

        # 项目全称
        lines.append(f"        {project_name}")
        if project_id and project_id != '________':
            lines.append(f"        （项目编号：{project_id}）")

        lines.append("")
        lines.append("            投  标  文  件")
        lines.append("")
        lines.append("=" * 52)
        lines.append("")

        # 招标人（如果招标文件要求封面标注）
        if tenderer and tenderer != '贵单位':
            lines.append(f"  招标人：{tenderer}")

        # 投标人
        lines.append(f"  投标人：{bidder}（公章）")

        # 法定代表人或授权代表
        lines.append(f"  法定代表人或授权代表：______________（签字）")

        # 日期：开标日期
        lines.append(f"  日期：{bid_date}")

        # 如果招标文件要求其他封面信息
        if cover_items:
            lines.append("")
            lines.append("  ---封面标注信息---")
            for item in cover_items[:5]:
                lines.append(f"  · {item.strip()}")

        return "\n".join(lines)

    def _ch_bid_letter(self, ctx, project, reqs, raw, cover, fmt):
        project_name = ctx['project_name']
        project_id = ctx['project_id']
        bidder = ctx['bidder_name']
        tenderer = ctx['tenderee_name']
        bid_amount = ctx['bid_amount']
        validity = ctx['validity_days']
        impl = ctx['implementation_period']
        bid_date = ctx['bid_date']

        content = f"""致：{tenderer}

我方（投标人名称：{bidder}）在仔细研究了贵方 {project_name}（项目编号：{project_id}）招标文件的全部内容（包括澄清、修改及补充通知）后，决定参加该项目的投标活动。

我方在此郑重承诺：

  1. 我方愿意按照招标文件的要求，以人民币 {bid_amount} 元的总报价承担本项目的全部工作。

  2. 如果我方中标，我方保证在合同签订后 {impl} 内完成本项目，并通过验收交付使用。

  3. 我方提交的投标文件在投标截止日后 {validity} 天内有效。在此期间，我方受本投标文件的约束。

  4. 我方承诺不将中标项目转包，不违法分包。

  5. 我方承诺，与招标人不存在任何可能影响招标公正性的利害关系。

  6. 我方理解并同意，贵方不一定要接受最低报价的投标，也不对未中标原因作任何解释。

  7. 我方声明，本投标文件中所有资料内容完整、真实、准确、有效。

  8. 如我方中标，我方将按招标文件规定的时间、金额提交履约保证金，并按时签订合同。

投标人（公章）：{bidder}
法定代表人或授权代表（签字）：______________
日  期：{bid_date}
地  址：{self.company.get('bidder_address', '')}
电  话：{self.company.get('bidder_phone', '')}
传  真：{self.company.get('bidder_fax', '')}
"""
        return content

    def _ch_bid_summary(self, ctx, project, reqs, raw, cover, fmt):
        project_name = ctx['project_name']
        project_id = ctx['project_id']
        bidder = ctx['bidder_name']
        bid_amount = ctx['bid_amount']
        validity = ctx['validity_days']
        impl = ctx['implementation_period']
        warranty = ctx['warranty_period']

        header = f"开标一览表\n\n"
        header += f"项目名称：{project_name}\n"
        header += f"项目编号：{project_id}\n"
        header += f"投标人名称：{bidder}（公章）\n"
        header += f"日期：{ctx['bid_date']}\n\n"

        table = """┌──────────────────────────┬──────────────────────────────┐
│         项  目           │           内  容             │
├──────────────────────────┼──────────────────────────────┤"""
        rows = [
            ("项目名称", project_name),
            ("项目编号", project_id),
            ("投标总报价", f"人民币 {bid_amount} 元"),
            ("实施周期", impl),
            ("质保期", warranty),
            ("投标有效期", f"{validity} 天"),
            ("备注", "完全响应招标文件全部要求"),
        ]
        for label, val in rows:
            # 截断过长内容
            label_display = label[:14]
            val_display = str(val)[:28]
            table += f"\n│ {label_display:<24} │ {val_display:<28} │"

        table += """
└──────────────────────────┴──────────────────────────────┘"""

        return header + table + "\n\n注：本表中报价与投标函不一致时，以本表为准。"

    def _ch_authorization(self, ctx, project, reqs, raw, cover, fmt):
        bidder = ctx['bidder_name']
        legal = ctx['legal_representative']
        authorized = ctx['authorized_person']
        project_name = ctx['project_name']
        project_id = ctx['project_id']
        bid_date = ctx['bid_date']
        tenderer = ctx['tenderee_name']

        return f"""致：{tenderer}

本授权委托书声明：我 {legal} 系 {bidder} 的法定代表人，现授权委托 {authorized} 为我方代理人，以我方的名义参加 {project_name}（项目编号：{project_id}）的投标活动。

代理人在投标、开标、评标、合同谈判、合同签订过程中所签署的一切文件和处理与之有关的一切事务，我方均予以承认。

代理人无转委托权。特此委托。

投标人（公章）：{bidder}
法定代表人（签字）：______________
授权委托人（签字）：______________
日  期：{bid_date}

附：
  1. 法定代表人身份证复印件（加盖公章）
  2. 授权代表身份证复印件（加盖公章）
"""

    def _ch_qualifications(self, ctx, project, reqs, raw, cover, fmt):
        qual_list = "\n".join(f"  [{i+1}] {q}" for i, q in enumerate(self.qualifications))

        return f"""资格证明文件

一、公司基本信息

  公司名称：{self.company.get('bidder_name', '')}
  统一社会信用代码：{self.company.get('business_license_no', '')}
  注册资本：{self.company.get('registered_capital', '')}
  成立日期：{self.company.get('established_date', '')}
  注册地址：{self.company.get('bidder_address', '')}
  经营范围：{self.company.get('scope_of_business', '')}

二、资质证书

{qual_list}

三、财务状况

  提供近三年经审计的财务报告（见附件）。

四、信誉声明

  我方郑重声明：近三年内在经营活动中没有重大违法记录，未被列入失信被执行人名单、重大税收违法案件当事人名单和政府采购严重违法失信行为记录名单。

注：以上证明材料复印件见附件。
"""

    def _ch_technical(self, ctx, project, reqs, raw, cover, fmt):
        project_name = ctx['project_name']
        scope = project.get('project_scope', '')
        tech_reqs = reqs.get('technical', [])

        if self._llm_enabled and tech_reqs:
            try:
                return self._llm_technical(project, reqs, raw)
            except:
                pass

        sections = []
        sections.append(f"技术方案\n")
        sections.append(f"项目名称：{project_name}")
        sections.append(f"项目编号：{ctx['project_id']}\n")

        sections.append("一、项目理解")
        sections.append(f"我方已认真研读 {project_name} 招标文件中的全部技术要求。")
        if scope:
            sections.append(f"招标范围：{scope}")
        sections.append("我方对该项目的关键技术点和实施难点有充分的认识和准备。\n")

        sections.append("二、技术需求逐条响应")
        if tech_reqs:
            for i, req in enumerate(tech_reqs[:15], 1):
                sections.append(f"  需求[{i}]：{req}")
                sections.append(f"  我方响应：完全理解并满足该要求。我方将采用成熟、先进的技术方案来保障该需求的实现，确保交付质量达到或超过招标文件规定的标准。\n")
        else:
            sections.append("  我方承诺完全响应并满足招标文件中的全部技术要求和功能需求。\n")

        sections.append("三、技术路线与架构设计")
        sections.append("  遵循先进性、可靠性、安全性、可扩展性的设计原则，采用业界成熟的技术框架。")
        sections.append("  系统架构采用分层设计，包括：表现层、业务逻辑层、数据访问层、基础设施层。\n")

        sections.append("四、质量保障措施")
        sections.append("  1. 严格执行ISO9001质量管理体系标准")
        sections.append("  2. 制定详细的质量控制计划，覆盖开发全过程")
        sections.append("  3. 引入自动化测试，确保代码质量和系统稳定性")
        sections.append("  4. 设置里程碑评审节点，关键节点须经招标方确认后继续推进\n")

        sections.append(f"投标人（公章）：{ctx['bidder_name']}")
        sections.append(f"日期：{ctx['bid_date']}")

        return "\n".join(sections)

    def _llm_technical(self, project, reqs, raw):
        from .bid_document_parser import BiddingDocumentParser as P
        prompt = f"""你是专业投标方案专家。请根据以下招标要求撰写技术方案。

项目：{project.get('project_name', '')}
编号：{project.get('project_id', '')}
范围：{project.get('project_scope', '')}

技术要求：
{json.dumps(reqs.get('technical', []), ensure_ascii=False)}

请按以下结构输出（每部分200-300字）：

一、项目理解
二、技术需求逐条响应（逐条列明我方满足方案）
三、总体技术方案
四、质量保障措施

要求：专业、严谨、突出优势，使用正式书面语。"""
        resp = self.llm_client.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7, max_tokens=2500,
        )
        text = resp['choices'][0]['message']['content']
        return f"技术方案\n\n项目名称：{project.get('project_name', '')}\n项目编号：{project.get('project_id', '')}\n\n{text}"

    def _ch_implementation(self, ctx, project, reqs, raw, cover, fmt):
        project_name = ctx['project_name']
        bidder = ctx['bidder_name']
        bid_date = ctx['bid_date']

        team_rows = []
        for m in self.team_members[:5]:
            team_rows.append(f"| {m.get('role', '')} | {m.get('name', '')} | {m.get('experience', '')} |")

        team_table = "\n".join([
            "| 角色 | 姓名 | 经验/资质 |",
            "|------|------|----------|",
        ] + team_rows)

        return f"""项目实施方案

项目名称：{project_name}
项目编号：{ctx['project_id']}

一、项目实施方案概述
我方根据 {project_name} 的招标要求，结合丰富的项目实施经验，制定了科学、详尽的实施方案。

二、实施计划

| 阶段 | 时间安排 | 主要工作内容 | 交付成果 |
|------|---------|-------------|---------|
| 启动阶段 | 第1周 | 组建团队、需求调研、制定详细计划 | 项目启动报告 |
| 设计阶段 | 第2-3周 | 方案设计、评审、确认 | 详细设计文档 |
| 实施阶段 | 第4-10周 | 开发、测试、联调 | 系统代码及文档 |
| 验收阶段 | 第11-12周 | 系统测试、用户验收 | 测试报告 |
| 上线阶段 | 第13周 | 部署上线、数据迁移 | 上线报告 |
| 收尾阶段 | 第14周 | 培训、文档移交 | 运维手册 |

三、项目团队配置

{team_table}

四、质量保障措施
1. 严格执行ISO9001质量管理体系
2. 实行周报、月报制度，定期向招标方汇报进度
3. 配置专职质量管理人员
4. 建立变更管理和风险预警机制
5. 所有交付物须经内部评审通过后方可提交

五、沟通协调机制
1. 项目经理为项目联系人，保持7×24小时响应
2. 每周召开项目例会，形成会议纪要
3. 重大问题2小时内上报，24小时内给出解决方案

投标人（公章）：{bidder}
日期：{bid_date}
"""

    def _ch_performance(self, ctx, project, reqs, raw, cover, fmt):
        project_name = ctx['project_name']
        bidder = ctx['bidder_name']

        proj_lines = []
        for i, proj in enumerate(self.past_projects[:5], 1):
            proj_lines.append(
                f"  [{i}] {proj.get('project_name', '')}\n"
                f"      委托方：{proj.get('client', '')}\n"
                f"      合同金额：{proj.get('amount', '')}\n"
                f"      完成时间：{proj.get('completion_date', '')}\n"
                f"      项目概况：{proj.get('description', '')}\n"
            )

        return f"""类似项目业绩

项目名称：{project_name}
项目编号：{ctx['project_id']}

我方在近三年内成功完成了多项与本项目类似的实施案例，积累了丰富的经验：

{''.join(proj_lines) if proj_lines else '  （业绩证明材料详见附件）'}

以上项目均通过了业主单位的验收并获得一致好评，充分证明我方具备完成 {project_name} 的技术实力和实施能力。

投标人（公章）：{bidder}
日期：{ctx['bid_date']}

注：合同关键页、验收报告等证明材料复印件见附件。
"""

    def _ch_appendices(self, ctx, project, reqs, raw, cover, fmt):
        bidder = ctx['bidder_name']
        bid_date = ctx['bid_date']

        return f"""其他材料 / 附录

投标人（公章）：{bidder}
日期：{bid_date}

附录清单：
  [1] 营业执照副本复印件（加盖公章）
  [2] 资质证书复印件（加盖公章）
  [3] 法定代表人身份证复印件（加盖公章）
  [4] 授权代表身份证复印件（加盖公章）
  [5] 类似项目合同与验收报告复印件（加盖公章）
  [6] 近三年审计报告复印件（加盖公章）
  [7] 近三个月社保缴纳证明复印件（加盖公章）
  [8] 近三年无重大违法记录声明函
  [9] 信用中国查询截图（加盖公章）
"""

    def _ch_generic(self, title: str, ctx):
        """未匹配章节的通用模板"""
        return f"""{title}

项目名称：{ctx['project_name']}
项目编号：{ctx['project_id']}
投标人：{ctx['bidder_name']}（公章）
日期：{ctx['bid_date']}

我方郑重承诺，针对"{title}"相关要求，我方完全响应并满足招标文件中的全部规定。
"""


def generate_bid_response(bidding_doc_path: str, extra_input: Dict = None) -> Dict:
    from .bid_document_parser import BiddingDocumentParser
    parser = BiddingDocumentParser()
    info = parser.parse(bidding_doc_path)
    gen = BidResponseGenerator()
    return gen.generate(info, extra_input)
