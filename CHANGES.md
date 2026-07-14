# 投标响应文件生成器 — 修改痕迹总结

> 本文档记录所有算法和识别逻辑的修改，以便换上新的招标文件后按相同的逻辑运行。
> 最后更新：2026-07-14（v6.1 标题识别鲁棒性增强+章节边界修复）

---

## v6.1 关键改进：标题识别鲁棒性增强 & 章节边界修复

| 改动 | 文件 | 说明 |
|------|------|------|
| 去掉标题加粗强制要求 | `src/response_template_extractor.py` | 不同招标文件格式不一致，用长度+标点排除代替加粗判断 |
| 二级标题长度收紧 | `src/response_template_extractor.py` | 20→18字，防止响应函内编号列表被误识别 |
| 句号排除规则 | `src/response_template_extractor.py` | 以"。"结尾的段落不算标题（内容项而非标题） |
| 章节边界修复 | `src/response_template_extractor.py` | `in` 改为 `re.match(rf'^\s*第{next}章')`，防止正文引用被误判为下一章 |

**触发场景**：新加入的"昆明市呈贡区"招标文件，其:
1. 标题不加粗（如"一、资格证明材料"）→ 去掉加粗强制要求
2. 正文中引用"第三章"（如"详见第三章服务要求"）→ 章节边界误判

**测试结果**：
- 昆明市呈贡区文档：10 一级 + 10 二级（之前仅 1+1）
- 云南现代职业技术学院文档：8 一级 + 4 二级（章节边界修复后正确提取）

---

## v6 关键改进：自动识别一级/二级标题

| 改动 | 文件 | 说明 |
|------|------|------|
| 一级标题自动识别 | `src/response_template_extractor.py` | 识别 "一、xxx" 格式（加粗 ≤25字）、"格式X-Y/附件X" 格式 |
| 二级标题自动识别 | `src/response_template_extractor.py` | 识别 "（一）xxx" 格式（加粗 ≤20字） |
| 清除旧 outlineLvl | `src/response_template_extractor.py` | 统一清除源文档残留层级标记，避免误识别 |
| 封面标题标记 | `src/response_template_extractor.py` | 封面项目名自动设为文档一级标题 |
| 排除规则 | `src/response_template_extractor.py` | 注/说明/致/项目编号/供应商盖章/日期/纯序号 ≠ 标题 |
| Word导航窗格支持 | 输出 DOCX | 所有标题含 `outlineLvl` 属性，打开 Word 即可在导航窗格看到文档结构 |

**识别策略（纯文本+格式特征，完全通用化）：**
- 一级标题：`一、～十、` 中文序号开头 + 加粗 + ≤25字（区分 "一、商务条款偏离表"[标题] vs "一、我公司承诺..."[内容]）
- 二级标题：`（一）～（十）/（1）～（N）` 括号序号开头 + 加粗 + ≤20字
- 自动跳过内容性编号列表、签名行、日期行、注释行等

**本次测试结果（云南现代职业技术学院项目）：**
- 一级标题：19个（1封面 + 18章节）
- 二级标题：5个
- 误识别：0

---

## v5 关键改进：格式与解析细节优化

| 改动 | 文件 | 说明 |
|------|------|------|
| 页边距统一 2.5cm | `src/response_template_extractor.py` | 上下左右全部改为 `Cm(2.5)` |
| 标题统一宋体、四号、1.5 倍行距 | `src/response_template_extractor.py` | 封面标题/项目名/"响应文件"/"项目编号"、目录标题与条目统一使用；签名区保留招标文件原始字体 |
| 封面自动检测与跳页 | `src/response_template_extractor.py` | 以"封面"关键字定位封面，跳过"第X章"和编制须知；封面后分页 → 目录 → 分页 → 正文 |
| 日期统一填开标日期 | `src/response_template_extractor.py` + `src/bid_document_parser.py` | 新增空格/空值型日期占位符替换；解析器支持竞争性磋商"开启"关键字 |
| 项目编号支持中文编号 | `src/bid_document_parser.py` | 正则支持中文、中文括号 `〔〕`、全角括号等，如 `云咏招〔2026〕0702号` |
| 项目名/编号空格占位符 | `src/response_template_extractor.py` | 支持 `项目编号：             `、`（项目名称：           、项目编号：              ）` 等格式 |

---

## v4 关键改进：消除所有硬编码位置假设

v3 版本中有以下隐式假设，在换用不同招标文件时可能失效。v4 全部消除：

| v3 硬编码 | 问题 | v4 解决方案 |
|-----------|------|------------|
| `if i < 30: continue` | 假设第七章一定在第 30 段之后 | 全文搜索，取**以"第X章"开头**的段落中最后一个匹配 |
| `'格式1-2' in text or '授权书'` | 假设封面格式一定是这种命名 | 自动收集独立格式标记，**第二个唯一标记之前 = 封面** |
| `cover_end_index = min(15, ...)` | 兜底值完全随机 | 未检测到第二个格式标记时返回 0，不插目录不分离封面 |
| `r'(格式\d+[-–—]\d+)'` TOC 正则 | 只匹配"格式X-X" | 扩展为 `(格式\|附件\|附表\|表格\|附录)任意编号` |
| `startswith('格式1-')` / `startswith('格式2-')` | 假设格式永远分两组 | 自动按前缀分组，支持任意数量的分组 |
| `第[一二三...]章.*(?:格式\|响应文件)` | `.*` 贪婪跨大段误匹配 | `. {0,30}?` 限制间距 + `^\s*` 锚定段落开头 |

---

## 修改概览

| 文件 | 状态 | 修改类型 |
|------|------|----------|
| `process_all.py` | 已修改 | 核心流程重写 |
| `src/bid_document_parser.py` | 已修改 | 新增代理机构提取 |
| `src/bid_recognizer.py` | 已修改 | 扩展识别关键词 |
| `src/response_template_extractor.py` | **新增** | 核心算法模块（v4 消除全部硬编码） |

---

## 1. `process_all.py` — 主流程（v1 → v2）

### 删除内容（不再使用的旧逻辑）
- 移除了 `BidResponseGenerator`（AI 生成技术方案）
- 移除了 `generate_docx`（通用 DOCX 生成）
- 移除了 `--api-key`、`--base-url`、`--bid-info` 命令行参数

### 新增内容
- **引入 `ResponseTemplateExtractor`**：替代旧生成器，直接从招标文件第七章复制格式
- **自动加载公司信息**：`_load_company()` → 默认读取 `data/company_profile.json`，返回 `company` 字段
- **跳过非 DOCX 文件**：PDF/TXT 等不再处理
- **输出命名**：`{项目名称}_响应文件.docx`

### 核心更新后的流程（三条流水线）

```
步骤 1/3 → BidFileRecognizer.scan_directory()
           扫描 input/ 目录，识别招标文件

步骤 2/3 → _load_company()
           加载 data/company_profile.json

步骤 3/3 → BiddingDocumentParser.parse()
           解析第一章"磋商邀请" → 提取项目信息
        → ResponseTemplateExtractor.generate()
           提取第七章格式 → XML 替换占位符 → 生成 DOCX
```

### 填充数据字段（传给提取器的字典）

```python
fill_data = {
    'project_name':         # 项目名称（来自第一章解析）
    'project_id':           # 项目编号（如 SCIT-GN-2026060281）
    'bid_opening_date':     # 开标日期（格式：2026年7月10日）
    'bidder_name':          # 供应商名称（来自 company_profile.json）
    'legal_representative': # 法定代表人
    'authorized_person':    # 授权代表
    'agency_name':          # 采购代理机构（来自第一章解析）
    'address':              # 公司地址
    'phone':                # 联系电话
    'zip_code':             # 邮政编码
    'fax':                  # 传真
    'package_no':           # 包号（默认 "1"）
}
```

---

## 2. `src/bid_document_parser.py` — 解析器（v1 → v2）

### 新增方法：`_extract_agency()`

在 `_extract_dates()` 之后新增一步，提取采购代理机构名称。

**匹配策略（4 条正则，依次尝试）**：

| 优先级 | 正则模式 | 示例原文 |
|--------|----------|----------|
| 1 | `(?:采购代理机构\|招标代理机构\|代理机构)[：:\s]*([^\n]{4,40}?)` | `采购代理机构：四川国际招标有限责任公司` |
| 2 | `([^\n]{4,30}(?:招标\|代理\|咨询\|管理)(?:有限\|股份)?公司)\s*受\s*[^\n]{2,20}\s*(?:委托)` | `四川国际招标有限责任公司受成都市第四人民医院委托` |
| 3 | `(?:我司\|我公司)\s*[：:在]?\s*([^\n]{4,30}...)` | `我司：四川国际招标有限责任公司` |
| 4 | 仅检测存在性 | `代理机构：地址` |

结果写入 `project_info['agency_name']`，最终用于填充响应文件的代理机构占位符。

---

## 3. `src/bid_recognizer.py` — 识别器

### 新增关键词

**竞争性磋商专用**（识别 `竞争性磋商` 类型的招标文件）：
- `竞争性磋商`、`磋商邀请`、`磋商文件`、`磋商须知`
- `响应文件格式`、`响应文件`
- `评审方法`、`评审标准`

**辅助识别**（提高识别准确率）：
- `采购代理机构`
- `采购项目编号`、`采购编号`、`项目编号`

### 其他改动
- 从 `SUPPORTED_EXTENSIONS` 中移除 `.md`（不扫描 Markdown）
- 新增 `EXCLUDE_FILENAMES = {'readme.md', 'readme.txt'}` 排除规则
- 输出标记从 emoji（✓✗?）替换为纯文本（[OK][??]）

---

## 4. `src/response_template_extractor.py` — 核心算法（新增全体）

> **这是最重要的模块**，包含了所有文档生成的核心逻辑。

### 4.1 工作流程

```
招标文件 .docx
      │
      ▼
_find_chapter7()        ←─ 定位"第七章 响应文件格式"
      │
      ▼
_build_new_document()   ←─ 按 body 元素顺序克隆所有段落和表格
      │                     ├─ w:p（段落） → deepcopy + 占位符替换
      │                     └─ w:tbl（表格） → deepcopy + 占位符替换
      │
      ├─► 封面（前 N 个元素）
      ├─► 分页符
      ├─► _add_toc()       ←─ 自动生成目录
      ├─► 分页符
      └─► 正文（余下全部元素）
```

### 4.2 章节定位算法（`_find_format_chapter`）—— v4 零硬编码

**策略**：
1. 全文扫描所有段落，匹配正则：`^\s*第[\d一二三...]+章.{0,30}?(响应文件格式|投标文件格式|格式要求|响应文件)`
   - `^\s*` 锚定段落开头（真正的章节标题以"第X章"开始，非正文引用）
   - `.{0,30}?` 限制间距防止跨大段误匹配
2. 取**最后一个**匹配（目录引用在前，真实标题在后）
3. 自动计算下一章序号（`_next_chapter_number()` 支持中文/阿拉伯数字）
4. 搜索下一章标题确定结束边界，找不到则取文档末尾

### 4.3 关键创新：按 Body 元素遍历（解决表格丢失问题）

**旧版 BUG**：只遍历 `doc.paragraphs`，表格 `w:tbl` 被跳过（段落与表格在 XML body 中是同级交替元素）。

**新版修复**：遍历 `doc.element.body` 的所有子元素，检测 tag 类型：
```python
for child in src_body:
    tag = child.tag.split('}')[-1]
    if tag == 'p':       # 段落 → deepcopy + 替换占位符
    elif tag == 'tbl':   # 表格 → deepcopy + 替换占位符
```

这样可以保留所有表格（供应商基本情况表、技术/服务要求应答表、商务应答表等）。

### 4.4 封面截止检测（`_detect_cover_end`）—— v4 零硬编码

**策略**：按顺序扫描元素，用 `_UNIQUE_FORMAT_RE` 收集独立格式标记。
**第二个独立标记出现的位置 = 封面结束**。

```python
seen_formats = []
for i, elem in enumerate(ch7_elements):
    m = _UNIQUE_FORMAT_RE.search(text)
    if m:
        normalized = normalize(m.group(1))
        if normalized not in seen_formats:
            seen_formats.append(normalized)
            if len(seen_formats) >= 2:
                return i  # 第2个格式前 = 封面
return 0  # 未检测到独立封面
```

- 通用匹配 `(格式|附件|附表|表格|附录)+编号`，不硬编码格式命名
- 未检测到独立封面时返回 0（不插入目录，不分离封面）

### 4.5 目录生成算法（`_add_toc`）—— v4 零硬编码

```
遍历第七章元素
  │
  ├─ 正则匹配 (格式|附件|附表|表格|附录)任意编号
  │   例：格式1-2、附件一、附表3-1
  │
  ├─ 获取描述文字
  │   ├─ 优先：当前段落中冒号/逗号后的文字
  │   └─ 备用：下一段落文字（排除格式/正本/副本等干扰）
  │
  ├─ 去重保持顺序
  │
  └─ _group_titles() 自动按前缀分组
       ├─ 提取前缀（格式1、格式2、附件一...）
       ├─ 相邻同前缀归为一组
       └─ >1 组时显示"第一部分/第二部分/..."，否则直接列表

### 4.6 占位符替换算法（`_replace_in_xml` + `_apply_replacements`）

在 XML 级别操作，不经过 python-docx 高层 API，**完全保留原始格式**（字体、字号、颜色、对齐等）。

**替换流程**：
1. 收集元素内所有 `w:t` 节点的文本
2. 合并为一个完整字符串
3. 用正则匹配并替换
4. 将结果写回第一个 `w:t` 节点，其余清空

**替换规则表**（按执行顺序）：

| 序号 | 替换目标 | 正则模式 | 填充来源 |
|------|----------|----------|----------|
| 1 | 项目名称 | `XX项目` / `XXXX项目` | `fill['project_name']` |
| 2 | 引号内项目名 | `"XXXXXX"` / `"XXXXXX"` | `fill['project_name']` |
| 3 | 项目编号 | `项目编号：XXXX` | `fill['project_id']` |
| 4 | 日期 | `XXX年XXX月XXX日` | `fill['bid_opening_date']` |
| 5 | 代理机构 | `XXX（采购代理机构名称）：` | `fill['agency_name']` |
| 6 | 供应商名称 | `供应商名称：XXXX（盖章）` | `fill['bidder_name']` |
| 7 | 法人签字 | `法定代表人...：XXXX` | `fill['legal_representative']` |
| 8 | 授权书内容 | `XXXX（供应商名称）XXXX（法定代表人...` | `bidder_name` + `legal_representative` |
| 9 | 被授权人 | `授权XXXX（被授权人` | `fill['authorized_person']` |
| 10 | 法人/授权代表签字 | `法定代表人...签字...：XXXX` | `legal_representative` / `authorized_person` |
| 11 | 通讯地址 | `通讯地址：XXXX` | `fill['address']` |
| 12 | 邮政编码 | `邮政编码：XXXX` | `fill['zip_code']` |
| 13 | 联系电话 | `联系电话：XXXX` | `fill['phone']` |
| 14 | 传真 | `传    真：XXXX` | `fill['fax']` |
| 15 | 包号 | `包        号：` | `fill['package_no']` |
| 16 | 兜底项目名 | `"XXXXXX"` 残余 | `fill['project_name']` |

### 4.7 分页符与段落构建

- `_add_section_break()`: 直接构建 `w:br` XML 元素实现分页
- `_make_paragraph()`: 用 `OxmlElement` 构建格式化段落，支持字体、字号、加粗、对齐、缩进

### 4.8 便捷函数

```python
def generate_response_docx(bidding_docx_path: str, fill_data: dict) -> str:
    """一行调用：返回生成的 DOCX 路径"""
```

---

## 5. 使用方法（换新招标文件后）

### 5.1 准备工作

1. 将新的招标文件（`.docx`）放入 `input/` 文件夹
2. 编辑 `data/company_profile.json`，更新公司信息：
   - `bidder_name`：公司名称
   - `legal_representative`：法定代表人
   - `authorized_person`：授权代表
   - `bidder_address`：通讯地址
   - `bidder_phone`：联系电话
   - `bidder_fax`：传真
   - etc.

### 5.2 运行

```bash
# 一键生成
python process_all.py

# 仅预览（不生成）
python process_all.py --scan-only

# 指定公司配置
python process_all.py --company ./my_company.json
```

### 5.3 输出

- 响应文件：`output/{项目名称}_响应文件.docx`
- 处理报告：`output/生成报告_YYYYMMDD_HHMMSS.json`

---

## 6. 对新招标文件的适配要求

生成器**不硬编码**章节号或编号格式，适应绝大多数政府采购文件的标准结构：

| 条件 | v4 匹配策略 |
|------|------------|
| 包含"第X章" + "响应文件格式/投标文件格式/格式要求" | 正则：`^\s*第[\d一二三...]+章.{0,30}?(响应文件格式\|投标文件格式\|格式要求\|响应文件)` |
| 格式用"(格式\|附件\|附表\|表格\|附录)+编号"标识 | 正则：`(格式\|附件\|附表\|表格\|附录)任意编号`，支持中文/阿拉伯数字 |
| 下一章自动检测 | `_next_chapter_number()` 递推，支持中文和阿拉伯数字 |
| 占位符用 X 或 x | `XXXX项目`、`供应商名称：XXXX`、`XXX年XXX月XXX日` 等多模式覆盖 |
| 第一章含项目信息 | 项目名称、编号、开标日期、代理机构名称 |

### 如果不满足以上条件，需要调整的地方：

1. **章节标题不在段落开头**（如前面有空格）→ 调整 `_CHAPTER_PATTERN` 的 `^\s*` 前缀
2. **格式标记命名完全不同**（如"材料一"、"文件A"）→ 扩展 `_FORMAT_MARKER_PATTERN` 正则
3. **无下一章**（招标文件很短）→ `_find_format_chapter` 自动取文档末尾兜底
4. **占位符不同**（如用下划线 `____`）→ 在 `_apply_replacements` 中添加对应正则规则

---

## 7. 故障排查

| 现象 | 可能原因 | 解决 |
|------|----------|------|
| 未找到格式章节 | 章节标题不在段落开头或缺少"响应文件格式"关键词 | 检查 `_find_format_chapter` 日志，调整 `_CHAPTER_PATTERN` |
| 表格仍然丢失 | 表格不在检测到的章节范围内 | 检查章节定位日志的 paragraph_start/end |
| 封面/目录未分离 | 格式标记少于 2 个（只有一种格式编号） | 正常行为，`_detect_cover_end` 返回 0；手动设置 `cover_end_index` |
| 目录分组不对 | 格式前缀提取不正确 | 调整 `_group_titles` 的 `prefix_re` 正则 |
| 占位符未被替换 | 正则不匹配 | 在 `_apply_replacements` 添加对应规则 |
| 项目名称显示"未知项目" | 第一章解析失败 | 检查 `_extract_project_info` 正则 |
| 代理机构为空 | 原文描述方式不同 | 在 `_extract_agency` 添加匹配模式 |
