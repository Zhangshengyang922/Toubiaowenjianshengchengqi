"""
文件识别器 —— 自动扫描目录，识别并分类招投标相关文件

工作流程：
1. 扫描指定目录下的所有文件
2. 读取内容，通过关键词 / 模式匹配判断是否为招标文件
3. 按识别结果分类：招标文件 / 非招标文件 / 无法识别
4. 返回可处理的招标文件列表
"""
import os
import re
from typing import List, Dict, Tuple
from datetime import datetime


class BidFileRecognizer:
    """招投标文件识别器"""

    # 招标文件关键词（出现越多，越可能是招标文件）
    BIDDING_KEYWORDS = [
        # 高频确定性关键词
        '招标公告', '招标文件', '投标邀请', '投标人须知',
        '招标项目', '公开招标', '邀请招标', '竞争性谈判',
        '投标截止', '开标时间', '投标有效期', '投标保证金',
        '评标办法', '技术规格', '商务条款', '合同条款',
        '资格要求', '资质要求', '资格审查',
        # 辅助识别
        '采购项目', '采购公告', '采购人',
        '投标人资格', '投标人资质', '联合体投标',
        '招标范围', '项目需求', '技术需求',
        'bidding', 'tender', 'RFP', 'RFQ',
    ]

    # 非招标文件标识（避免误识别）
    EXCLUDE_KEYWORDS = [
        '投标文件', '投标函', '中标通知书',
        '开标一览表', '法定代表人', '授权委托书',
    ]

    SUPPORTED_EXTENSIONS = {
        '.pdf', '.docx', '.doc', '.txt', '.md',
        '.xlsx', '.xls', '.csv',
    }

    def __init__(self, input_dir: str = None):
        if input_dir is None:
            input_dir = os.path.join(
                os.path.dirname(os.path.dirname(__file__)), 'input'
            )
        self.input_dir = input_dir
        os.makedirs(self.input_dir, exist_ok=True)

    # ===== 核心识别逻辑 =====

    def scan_directory(self) -> Dict[str, List[Dict]]:
        """
        扫描目录，返回分类结果
        
        Returns:
            {
                'bidding_docs': [...],   # 确认为招标文件
                'other_files': [...],    # 非招标文件
                'unrecognized': [...],   # 无法识别
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
                ext = os.path.splitext(f)[1].lower()
                if ext in self.SUPPORTED_EXTENSIONS or ext == '':
                    files.append(full_path)

        for filepath in sorted(files):
            info = self._analyze_file(filepath)
            category = info['category']
            result[category].append(info)

        return result

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

        # 计分识别
        bidding_score = self._calculate_bidding_score(text)
        exclude_score = self._calculate_exclude_score(text)
        
        info['bidding_score'] = bidding_score
        info['exclude_score'] = exclude_score

        if bidding_score >= 3 and bidding_score > exclude_score:
            info['category'] = 'bidding_docs'
            info['confidence'] = '高' if bidding_score >= 6 else '中'
            info['matched_keywords'] = self._get_matched_keywords(text)
        elif bidding_score >= 1:
            info['category'] = 'unrecognized'
            info['reason'] = f'招标特征不明显（得分{bidding_score}），建议确认'
        else:
            info['category'] = 'other_files'
            info['reason'] = '非招标文件'

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
                        pages = min(5, len(reader.pages))  # 只读前5页做识别
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
                    for para in doc.paragraphs[:100]:  # 前100段
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

        return text[:5000]  # 截断，识别不需要全文

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
                # 高频关键词加分更多
                if kw in ['招标文件', '招标公告', '投标人须知']:
                    score += 3
                elif kw in ['公开招标', '投标截止', '开标时间', '评标办法']:
                    score += 2
                else:
                    score += 1
        return score

    def _calculate_exclude_score(self, text: str) -> int:
        """计算排除分（越像投标文件越要排除）"""
        score = 0
        for kw in self.EXCLUDE_KEYWORDS:
            if kw in text:
                score += 2
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
        """获取所有确认为招标文件的路径"""
        result = self.scan_directory()
        return [f['path'] for f in result['bidding_docs']]

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
        print(f"  ├─ 招标文件: {bidding_count} 个  ← 将自动生成投标文件")
        print(f"  ├─ 非招标文件: {other_count} 个")
        print(f"  └─ 无法识别: {unrecognized_count} 个")

        if bidding_count > 0:
            print(f"\n  【待处理的招标文件】")
            for f in result['bidding_docs']:
                print(f"  ✓ {f['filename']} ({f['size_kb']}KB, 置信度:{f.get('confidence','-')})")
        
        if other_count > 0:
            print(f"\n  【已跳过的非招标文件】")
            for f in result['other_files']:
                print(f"  ✗ {f['filename']} - {f.get('reason', '')}")

        if unrecognized_count > 0:
            print(f"\n  【无法识别的文件（请手动确认）】")
            for f in result['unrecognized']:
                print(f"  ? {f['filename']} - {f.get('reason', '')}")
                if f.get('text_preview'):
                    print(f"    预览: {f['text_preview'][:80]}...")

        print("=" * 65 + "\n")
        return result


if __name__ == "__main__":
    import sys
    target = sys.argv[1] if len(sys.argv) > 1 else None
    recognizer = BidFileRecognizer(target)
    recognizer.print_report()
