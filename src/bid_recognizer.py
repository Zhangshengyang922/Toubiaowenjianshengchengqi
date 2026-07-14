"""
文件识别器 v3 —— 智能扫描目录，识别招投标相关文件

设计原则：
  - .docx/.doc 文件一律作为可处理候选，不再按关键字硬过滤
  - 关键字评分仅作辅助信息展示（置信度），不做阻挡
  - 排除名单只过滤明确的非招标文件（如已生成的投标文件本身）
"""
import os
import re
from typing import List, Dict, Tuple
from datetime import datetime


class BidFileRecognizer:
    """招投标文件识别器"""

    # ── 采购/招标关键词（用于置信度评估，不用于过滤）──
    BIDDING_KEYWORDS = [
        # 招标类
        '招标公告', '招标文件', '投标邀请', '投标人须知',
        '招标项目', '公开招标', '邀请招标', '竞争性谈判',
        '投标截止', '开标时间', '投标有效期', '投标保证金',
        '评标办法', '技术规格', '商务条款', '合同条款',
        '资格要求', '资质要求', '资格审查',
        '投标人资格', '投标人资质', '联合体投标',
        '招标范围', '项目需求', '技术需求',
        '中标', '流标', '废标', '唱标',
        # 磋商/询价/比选类
        '竞争性磋商', '磋商邀请', '磋商文件', '磋商须知',
        '询价', '比选', '单一来源', '框架协议',
        # 采购通用类
        '采购项目', '采购公告', '采购人', '采购代理机构',
        '采购项目编号', '采购编号', '项目编号',
        '采购文件', '供应商', '供应商须知',
        # 响应文件类
        '响应文件格式', '响应文件', '评审方法', '评审标准',
        # 英文
        'bidding', 'tender', 'RFP', 'RFQ', 'procurement',
    ]

    # ── 真正要排除的：已生成的投标文件自身 ──
    EXCLUDE_PATTERNS = [
        # 以这些词开头的文件很可能是已生成的投标响应文件
        r'^(投标文件|投标函|投标响应|响应文件|技术方案)',
        # 生成报告
        r'^生成报告',
    ]

    SUPPORTED_EXTENSIONS = {
        '.pdf', '.docx', '.doc', '.txt',
        '.xlsx', '.xls', '.csv',
    }

    # .docx/.doc 无条件可处理
    ALWAYS_PROCESSABLE = {'.docx', '.doc'}

    EXCLUDE_FILENAMES = {'readme.md', 'readme.txt', 'desktop.ini', 'thumbs.db'}

    def __init__(self, input_dir: str = None):
        if input_dir is None:
            input_dir = os.path.join(
                os.path.dirname(os.path.dirname(__file__)), 'input'
            )
        self.input_dir = input_dir
        os.makedirs(self.input_dir, exist_ok=True)

    # ===== 核心识别逻辑 =====

    def scan_directory(self, progress_callback=None) -> Dict[str, List[Dict]]:
        """
        扫描目录，返回分类结果

        Args:
            progress_callback: 可选的进度回调，签名为 (filename, info) -> None

        Returns:
            {
                'bidding_docs': [...],    # 可处理的招标/采购文件（全部 .docx）
                'other_files': [...],     # 其他支持的文件（PDF/TXT等，供参考）
                'unrecognized': [...],    # 无法读取/文本过短
            }
        """
        result = {
            'bidding_docs': [],
            'other_files': [],
            'unrecognized': [],
        }

        if not os.path.exists(self.input_dir):
            return result

        files = []
        for f in os.listdir(self.input_dir):
            full_path = os.path.join(self.input_dir, f)
            if os.path.isfile(full_path) and not f.startswith('~'):
                if f.lower() in self.EXCLUDE_FILENAMES:
                    continue
                ext = os.path.splitext(f)[1].lower()
                if ext in self.SUPPORTED_EXTENSIONS or ext == '':
                    files.append(full_path)

        for filepath in sorted(files):
            info = self._analyze_file(filepath)
            category = info['category']
            result[category].append(info)
            if progress_callback is not None:
                try:
                    progress_callback(os.path.basename(filepath), info)
                except Exception:
                    pass

        return result

    def get_all_processable(self) -> List[Dict]:
        """
        ⭐ GUI 专用：获取所有可处理的文件列表
        
        规则：
          - .docx/.doc 文件 → 直接纳入（不管关键字匹配度）
          - 其他格式 → 关键字评分 >= 1 才纳入
          - 排除明确是已生成投标文件的
        
        Returns:
            每个文件包含: path, filename, confidence, matched_keywords, summary 等
        """
        result = self.scan_directory()
        processable = []

        for f in result['bidding_docs']:
            f['can_process'] = True
            processable.append(f)

        # 未被分类为 bidding_docs 的 .docx 也加进来
        for f in result['unrecognized']:
            ext = f.get('extension', '').lower()
            if ext in self.ALWAYS_PROCESSABLE and f.get('text_length', 0) >= 50:
                f['can_process'] = True
                f['confidence'] = '低'
                f['note'] = '文件内容与常见招标关键词匹配度低，但可以尝试处理'
                processable.append(f)
            elif ext in self.ALWAYS_PROCESSABLE:
                f['can_process'] = True
                f['confidence'] = '极低'
                f['note'] = '无法读取文件内容，但仍可尝试处理'
                processable.append(f)

        # 其他格式中有得分的也纳入
        for f in result['other_files']:
            if f.get('bidding_score', 0) >= 1:
                f['can_process'] = True
                f['note'] = '非 DOCX 格式，语义匹配度有限'
                processable.append(f)

        # 标记哪些文件明确被排除（看起来像已生成的投标文件）
        for f in processable:
            f['is_excluded'] = self._looks_like_bid_response(f['filename'])

        return processable

    def _looks_like_bid_response(self, filename: str) -> bool:
        """判断文件名是否像已生成的投标响应文件"""
        basename = os.path.splitext(filename)[0]
        for pattern in self.EXCLUDE_PATTERNS:
            if re.search(pattern, basename):
                return True
        return False

    def _analyze_file(self, filepath: str) -> Dict:
        """分析单个文件"""
        filename = os.path.basename(filepath)
        ext = os.path.splitext(filepath)[1].lower()
        size_kb = round(os.path.getsize(filepath) / 1024, 1)

        info = {
            'path': filepath,
            'filename': filename,
            'extension': ext,
            'size_kb': size_kb,
            'mtime': datetime.fromtimestamp(os.path.getmtime(filepath)).strftime('%Y-%m-%d %H:%M'),
        }

        # 读取文本
        text = self._read_text(filepath)
        info['text_preview'] = text[:200] if text else ''
        info['text_length'] = len(text) if text else 0

        if not text or info['text_length'] < 50:
            info['category'] = 'unrecognized'
            info['reason'] = '文本内容过短，无法分析'
            return info

        # 计算关键词得分
        bidding_score = self._calculate_bidding_score(text)
        matched = self._get_matched_keywords(text)
        info['bidding_score'] = bidding_score
        info['matched_keywords'] = matched

        # 排除检查（已生成的投标文件）
        is_excluded = self._looks_like_bid_response(filename)

        # ⭐ .docx/.doc 文件一律归为可处理
        if ext in self.ALWAYS_PROCESSABLE:
            info['category'] = 'bidding_docs'
            if bidding_score >= 6:
                info['confidence'] = '高'
            elif bidding_score >= 3:
                info['confidence'] = '中'
            else:
                info['confidence'] = '低'
                info['note'] = '关键词匹配度较低，但文件格式支持处理'
            info['matched_keywords'] = matched
            if is_excluded:
                info['note'] = (info.get('note', '') + ' | 文件名疑似已生成的投标文件').strip(' |')
            return info

        # 非 DOCX 文件：需要关键字匹配
        if bidding_score >= 3:
            info['category'] = 'bidding_docs'
            info['confidence'] = '高' if bidding_score >= 6 else '中'
            info['matched_keywords'] = matched
        elif bidding_score >= 1:
            info['category'] = 'other_files'
            info['reason'] = f'招标特征不明显（得分{bidding_score}），建议确认是否处理'
        else:
            info['category'] = 'other_files'
            info['reason'] = '未检测到招标/采购相关关键词'

        return info

    def _read_text(self, filepath: str) -> str:
        """读取文件文本"""
        ext = os.path.splitext(filepath)[1].lower()
        text = ''

        try:
            if ext == '.pdf':
                try:
                    import PyPDF2
                    with open(filepath, 'rb') as f:
                        reader = PyPDF2.PdfReader(f)
                        pages = min(5, len(reader.pages))
                        for i in range(pages):
                            page_text = reader.pages[i].extract_text()
                            if page_text:
                                text += page_text
                except:
                    text = self._read_raw(filepath)

            elif ext in ('.docx', '.doc'):
                try:
                    from docx import Document
                    doc = Document(filepath)
                    for para in doc.paragraphs[:100]:
                        if para.text.strip():
                            text += para.text + '\n'
                except:
                    text = self._read_raw(filepath)

            elif ext in ('.txt', '.md', ''):
                text = self._read_raw(filepath)

            elif ext in ('.xlsx', '.xls'):
                try:
                    import pandas as pd
                    df = pd.read_excel(filepath, nrows=20)
                    text = df.to_string()
                except:
                    text = self._read_raw(filepath)

            elif ext == '.csv':
                try:
                    import pandas as pd
                    df = pd.read_csv(filepath, nrows=20)
                    text = df.to_string()
                except:
                    text = self._read_raw(filepath)

        except Exception as e:
            text = self._read_raw(filepath)

        return text[:5000]

    def _read_raw(self, filepath: str) -> str:
        """原始文本读取（兜底）"""
        for encoding in ['utf-8', 'gbk', 'gb2312', 'latin-1']:
            try:
                with open(filepath, 'r', encoding=encoding, errors='ignore') as f:
                    return f.read(5000)
            except:
                continue
        return ''

    def _calculate_bidding_score(self, text: str) -> int:
        """计算招标文件匹配分"""
        score = 0
        text_lower = text.lower()
        for kw in self.BIDDING_KEYWORDS:
            if kw.lower() in text_lower:
                if kw in ['招标文件', '招标公告', '投标人须知']:
                    score += 3
                elif kw in ['公开招标', '投标截止', '开标时间', '评标办法']:
                    score += 2
                else:
                    score += 1
        return score

    def _get_matched_keywords(self, text: str) -> List[str]:
        """获取匹配到的关键词"""
        matched = []
        for kw in self.BIDDING_KEYWORDS:
            if kw in text:
                matched.append(kw)
        return matched[:10]

    # ===== 对外接口 =====

    def identify(self, filepath: str) -> Dict:
        """识别单个文件"""
        return self._analyze_file(filepath)

    def get_bidding_files(self) -> List[str]:
        """获取所有可处理文件的路径"""
        processable = self.get_all_processable()
        return [f['path'] for f in processable if not f.get('is_excluded')]

    def print_report(self) -> Dict:
        """打印识别报告"""
        result = self.scan_directory()

        bidding_count = len(result['bidding_docs'])
        other_count = len(result['other_files'])
        unrecognized_count = len(result['unrecognized'])
        total = bidding_count + other_count + unrecognized_count

        print("\n" + "=" * 65)
        print("  招投标文件识别报告")
        print(f"  扫描目录: {self.input_dir}")
        print(f"  扫描时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 65)
        print(f"  总文件数: {total}")
        print(f"  ├─ 可处理文件: {bidding_count} 个  ← 将自动生成响应文件")
        print(f"  ├─ 其他文件: {other_count} 个")
        print(f"  └─ 无法识别: {unrecognized_count} 个")

        if bidding_count > 0:
            print(f"\n  【可处理的文件】")
            for f in result['bidding_docs']:
                conf = f.get('confidence', '-')
                note = f.get('note', '')
                extra = f' ⚠️ {note}' if note else ''
                print(f"  [OK] {f['filename']} ({f['size_kb']}KB, 置信度:{conf}){extra}")

        if other_count > 0:
            print(f"\n  【其他文件（非 DOCX 或关键字不匹配）】")
            for f in result['other_files']:
                reason = f.get('reason', '')
                print(f"  [..] {f['filename']} - {reason}")

        if unrecognized_count > 0:
            print(f"\n  【无法读取的文件】")
            for f in result['unrecognized']:
                print(f"  [?] {f['filename']} - {f.get('reason', '')}")
                if f.get('text_preview'):
                    print(f"      预览: {f['text_preview'][:80]}...")

        print("=" * 65 + "\n")
        return result


if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else None
    recognizer = BidFileRecognizer(target)
    recognizer.print_report()
