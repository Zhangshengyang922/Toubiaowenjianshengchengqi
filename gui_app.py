"""
招投标文件自动生成系统 —— GUI 桌面应用

功能:
  - 选择招标文件所在文件夹
  - 编辑/加载公司信息
  - 一键扫描 & 生成投标响应文件
  - 启动时自动检查更新
"""

import os
import sys
import json
import threading
import queue
import shutil
import traceback
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

# ── 路径处理（支持 PyInstaller 打包后的路径） ──
def _get_base_dir():
    """获取应用根目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = _get_base_dir()

# 添加 src 到路径
sys.path.insert(0, os.path.join(BASE_DIR, 'src'))
sys.path.insert(0, BASE_DIR)

from version import __version__, __app_name__, __repo_url__
from src.auto_updater import AutoUpdater
from src.bid_recognizer import BidFileRecognizer
from src.bid_document_parser import BiddingDocumentParser
from src.response_template_extractor import ResponseTemplateExtractor


class BiddingApp:
    """主应用窗口"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{__app_name__} v{__version__}")
        self.root.geometry("900x700")
        self.root.minsize(700, 550)

        # 设置窗口图标（如果有的话）
        try:
            icon_path = os.path.join(BASE_DIR, 'data', 'icon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass

        # 消息队列（工作线程 → GUI 线程）
        self.msg_queue = queue.Queue()
        self.is_processing = False
        self.cancel_requested = False

        # 默认路径（exe 模式下使用用户文档目录，避免 Program Files 权限问题）
        if getattr(sys, 'frozen', False):
            user_docs = os.path.join(os.path.expanduser('~'), 'Documents', '招投标文件')
            self._input_dir = tk.StringVar(value=os.path.join(user_docs, 'input'))
            self._output_dir = tk.StringVar(value=os.path.join(user_docs, 'output'))
        else:
            self._input_dir = tk.StringVar(value=os.path.join(BASE_DIR, 'input'))
            self._output_dir = tk.StringVar(value=os.path.join(BASE_DIR, 'output'))
        self._company_name = tk.StringVar()
        self._legal_rep = tk.StringVar()
        self._authorized_person = tk.StringVar()
        self._company_address = tk.StringVar()
        self._company_phone = tk.StringVar()
        self._company_fax = tk.StringVar()
        self._company_zip = tk.StringVar()

        # 扫描结果缓存
        self._scan_result = None
        self._bidding_files = []

        # 构建界面
        self._build_ui()

        # 启动后自动检查更新
        self.root.after(1000, self._check_updates_silent)

        # 定期处理消息队列
        self.root.after(100, self._process_queue)

    # ══════════════════════════════════════════════
    #  UI 构建
    # ══════════════════════════════════════════════

    def _build_ui(self):
        """构建完整界面"""
        # ── 菜单栏 ──
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="检查更新", command=self._check_updates_manual)
        help_menu.add_separator()
        help_menu.add_command(label="关于", command=self._show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        # ── 主容器（带滚动） ──
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── 标题区 ──
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(title_frame, text=__app_name__,
                  font=('Microsoft YaHei', 16, 'bold')).pack(side=tk.LEFT)
        self._version_label = ttk.Label(title_frame,
                                        text=f'v{__version__}',
                                        foreground='gray')
        self._version_label.pack(side=tk.LEFT, padx=10)

        # ── 上半部分：设置区 ──
        settings_frame = ttk.LabelFrame(main_frame, text='设置', padding=10)
        settings_frame.pack(fill=tk.X, pady=(0, 5))

        self._build_settings(settings_frame)

        # ── 中间：操作按钮 ──
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        self._scan_btn = ttk.Button(btn_frame, text='🔍 扫描文件',
                                    command=self._on_scan, width=14)
        self._scan_btn.pack(side=tk.LEFT, padx=(0, 10))

        self._generate_btn = ttk.Button(btn_frame, text='📄 一键生成响应文件',
                                        command=self._on_generate, width=20)
        self._generate_btn.pack(side=tk.LEFT, padx=(0, 10))

        self._cancel_btn = ttk.Button(btn_frame, text='取消',
                                      command=self._on_cancel, width=10,
                                      state=tk.DISABLED)
        self._cancel_btn.pack(side=tk.LEFT)

        self._progress = ttk.Progressbar(btn_frame, mode='indeterminate', length=150)
        self._progress.pack(side=tk.RIGHT, padx=10)

        # ── 下半部分：日志区 ──
        log_frame = ttk.LabelFrame(main_frame, text='处理日志', padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self._log = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=('Consolas', 10),
            bg='#1e1e1e', fg='#d4d4d4', insertbackground='white'
        )
        self._log.pack(fill=tk.BOTH, expand=True)

        # 配置日志颜色标签
        self._log.tag_config('INFO', foreground='#4ec9b0')
        self._log.tag_config('WARN', foreground='#dcdcaa')
        self._log.tag_config('ERROR', foreground='#f44747')
        self._log.tag_config('SUCCESS', foreground='#6a9955')
        self._log.tag_config('DIM', foreground='#808080')

        # ── 状态栏 ──
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))
        self._status_var = tk.StringVar(value='就绪')
        ttk.Label(status_frame, textvariable=self._status_var,
                  foreground='gray').pack(side=tk.LEFT)
        self._file_count_var = tk.StringVar()
        ttk.Label(status_frame, textvariable=self._file_count_var,
                  foreground='gray').pack(side=tk.RIGHT)

    def _build_settings(self, parent):
        """构建设置区"""
        # ── 输入目录 ──
        row1 = ttk.Frame(parent)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text='招标文件目录:', width=14, anchor=tk.E).pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self._input_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(row1, text='浏览...', command=self._browse_input, width=8).pack(side=tk.LEFT)

        # ── 输出目录 ──
        row2 = ttk.Frame(parent)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text='输出目录:', width=14, anchor=tk.E).pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self._output_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(row2, text='浏览...', command=self._browse_output, width=8).pack(side=tk.LEFT)

        # ── 分隔线 ──
        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)

        # ── 公司信息 ──
        info_label = ttk.Frame(parent)
        info_label.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(info_label, text='公司信息', font=('Microsoft YaHei', 10, 'bold')).pack(side=tk.LEFT)
        ttk.Button(info_label, text='加载配置...', command=self._load_company_profile, width=10).pack(side=tk.RIGHT, padx=2)
        ttk.Button(info_label, text='保存配置...', command=self._save_company_profile, width=10).pack(side=tk.RIGHT, padx=2)

        fields = [
            ('公司名称:', self._company_name),
            ('法定代表人:', self._legal_rep),
            ('授权代表:', self._authorized_person),
            ('公司地址:', self._company_address),
            ('联系电话:', self._company_phone),
            ('传真:', self._company_fax),
            ('邮编:', self._company_zip),
        ]

        for label, var in fields:
            row = ttk.Frame(parent)
            row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=label, width=14, anchor=tk.E).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

    # ══════════════════════════════════════════════
    #  事件处理
    # ══════════════════════════════════════════════

    def _browse_input(self):
        path = filedialog.askdirectory(title='选择招标文件所在文件夹')
        if path:
            self._input_dir.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title='选择输出文件夹')
        if path:
            self._output_dir.set(path)

    def _load_company_profile(self):
        path = filedialog.askopenfilename(
            title='选择公司信息配置文件',
            filetypes=[('JSON 文件', '*.json'), ('所有文件', '*.*')]
        )
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            company = data.get('company', data)
            self._company_name.set(company.get('bidder_name', ''))
            self._legal_rep.set(company.get('legal_representative', ''))
            self._authorized_person.set(company.get('authorized_person', ''))
            self._company_address.set(company.get('bidder_address', ''))
            self._company_phone.set(company.get('bidder_phone', ''))
            self._company_fax.set(company.get('bidder_fax', ''))
            self._company_zip.set(company.get('bidder_zip_code', ''))
            self._log_msg(f'已加载公司配置: {os.path.basename(path)}', 'SUCCESS')
        except Exception as e:
            self._log_msg(f'加载配置失败: {e}', 'ERROR')

    def _save_company_profile(self):
        path = filedialog.asksaveasfilename(
            title='保存公司信息配置',
            defaultextension='.json',
            filetypes=[('JSON 文件', '*.json')]
        )
        if not path:
            return
        try:
            data = {
                'company': {
                    'bidder_name': self._company_name.get(),
                    'legal_representative': self._legal_rep.get(),
                    'authorized_person': self._authorized_person.get(),
                    'bidder_address': self._company_address.get(),
                    'bidder_phone': self._company_phone.get(),
                    'bidder_fax': self._company_fax.get(),
                    'bidder_zip_code': self._company_zip.get(),
                }
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self._log_msg(f'公司配置已保存: {os.path.basename(path)}', 'SUCCESS')
        except Exception as e:
            self._log_msg(f'保存失败: {e}', 'ERROR')

    def _on_scan(self):
        """扫描文件按钮"""
        if self.is_processing:
            return
        self._start_task('扫描招标文件...', self._do_scan)

    def _on_generate(self):
        """生成按钮"""
        if self.is_processing:
            return

        # 如果没有扫描过，先扫描
        if self._scan_result is None:
            self._start_task('扫描并生成...', self._do_scan_then_generate)
        else:
            self._start_task('生成响应文件...', self._do_generate)

    def _on_cancel(self):
        """取消按钮"""
        self.cancel_requested = True
        self._log_msg('正在取消...', 'WARN')

    # ══════════════════════════════════════════════
    #  业务逻辑（工作线程）
    # ══════════════════════════════════════════════

    def _start_task(self, status_msg: str, task_func):
        """启动后台任务"""
        self.is_processing = True
        self.cancel_requested = False
        self._set_ui_state(tk.DISABLED)
        self._progress.start()
        self._status_var.set(status_msg)
        self._log.clear()
        self._log_msg(f'═══ {__app_name__} v{__version__} ═══', 'INFO')
        self._log_msg(f'{status_msg}\n')

        t = threading.Thread(target=task_func, daemon=True)
        t.start()

    def _finish_task(self, msg: str = '完成', tag: str = 'SUCCESS'):
        """结束任务"""
        self.is_processing = False
        self._set_ui_state(tk.NORMAL)
        self._progress.stop()
        self._status_var.set(msg)
        self._log_msg(f'\n═══ {msg} ═══', tag)

    def _set_ui_state(self, state):
        """设置控件状态"""
        self._scan_btn.config(state=state)
        self._generate_btn.config(state=state)
        self._cancel_btn.config(state=tk.NORMAL if state == tk.DISABLED else tk.DISABLED)

    def _do_scan(self):
        """后台：扫描文件"""
        input_dir = self._input_dir.get()
        self._put_msg(f'扫描目录: {input_dir}', 'INFO')

        if not os.path.isdir(input_dir):
            self._put_msg('目录不存在，请检查路径', 'ERROR')
            self._finish_task('扫描失败', 'ERROR')
            return

        recognizer = BidFileRecognizer(input_dir)
        self._scan_result = recognizer.print_report()
        self._bidding_files = self._scan_result.get('bidding_docs', [])

        self._put_msg(f'\n发现 {len(self._bidding_files)} 个招标文件:')
        for f in self._bidding_files:
            self._put_msg(f'  📄 {f["filename"]}', 'SUCCESS')

        if not self._bidding_files:
            self._put_msg('\n未发现招标文件！请确认目录中有 .docx 招标文件', 'WARN')

        # 更新状态栏
        self._queue.put(('file_count', len(self._bidding_files)))
        self._finish_task(f'扫描完成 — 发现 {len(self._bidding_files)} 个招标文件')

    def _do_scan_then_generate(self):
        """后台：扫描 + 生成"""
        self._do_scan()
        if self._bidding_files:
            self._do_generate()

    def _do_generate(self):
        """后台：生成响应文件"""
        bidding_files = self._bidding_files
        if not bidding_files:
            self._put_msg('没有可处理的招标文件，请先扫描', 'WARN')
            self._finish_task('无文件可处理', 'WARN')
            return

        # 获取公司信息
        company = self._get_company_data()

        # 输出目录
        output_dir = self._output_dir.get()
        os.makedirs(output_dir, exist_ok=True)

        # 保存公司信息到默认配置（供 extractor 内部使用）
        self._save_default_config(company)

        extractor = ResponseTemplateExtractor()
        results = []
        total = len(bidding_files)

        for i, file_info in enumerate(bidding_files):
            if self.cancel_requested:
                self._put_msg('\n用户取消操作', 'WARN')
                break

            filepath = file_info['path']
            filename = file_info['filename']

            self._put_msg(f'\n[{i+1}/{total}] {filename}', 'INFO')
            self._put_msg('─' * 50, 'DIM')

            try:
                # 只处理 DOCX
                if not filename.lower().endswith('.docx'):
                    self._put_msg('  跳过: 非 DOCX 文件', 'WARN')
                    results.append({'status': 'skipped', 'file': filename, 'reason': '非DOCX'})
                    continue

                # 解析项目信息
                self._put_msg('  [解析] 提取项目信息...')
                doc_parser = BiddingDocumentParser()
                bidding_info = doc_parser.parse(filepath)
                project = bidding_info.get('project_info', {})

                proj_name = project.get('project_name', project.get('project_name_original', '未知项目'))
                proj_id = project.get('project_id', '未识别')
                bid_date = project.get('bid_opening_date', '未识别')
                agency = project.get('agency_name', '')
                tenderer = project.get('tenderee_name', '')

                self._put_msg(f'  项目名称: {proj_name}')
                self._put_msg(f'  项目编号: {proj_id}')
                self._put_msg(f'  开标日期: {bid_date}')

                # 准备填充数据
                fill_data = {
                    'project_name': proj_name,
                    'project_id': proj_id,
                    'bid_opening_date': bid_date,
                    'bidder_name': company.get('bidder_name', ''),
                    'legal_representative': company.get('legal_representative', ''),
                    'authorized_person': company.get('authorized_person', ''),
                    'agency_name': agency or tenderer or '贵单位',
                    'address': company.get('bidder_address', ''),
                    'phone': company.get('bidder_phone', ''),
                    'zip_code': company.get('bidder_zip_code', ''),
                    'fax': company.get('bidder_fax', ''),
                    'package_no': '1',
                }

                # 生成
                self._put_msg('  [生成] 按招标文件格式生成响应文件...')
                output_path = extractor.generate(filepath, fill_data)

                # 移动到输出目录
                final_path = os.path.join(output_dir, os.path.basename(output_path))
                if os.path.abspath(output_path) != os.path.abspath(final_path):
                    shutil.move(output_path, final_path)

                self._put_msg(f'  ✓ 完成 -> {os.path.basename(final_path)}', 'SUCCESS')
                results.append({
                    'status': 'success',
                    'file': filename,
                    'project_name': proj_name,
                    'project_id': proj_id,
                    'output': final_path,
                })

            except Exception as e:
                self._put_msg(f'  ✗ 失败: {e}', 'ERROR')
                self._put_msg(traceback.format_exc(), 'DIM')
                results.append({'status': 'failed', 'file': filename, 'error': str(e)})

        # ── 汇总 ──
        success = sum(1 for r in results if r['status'] == 'success')
        failed = sum(1 for r in results if r['status'] == 'failed')
        skipped = sum(1 for r in results if r['status'] == 'skipped')

        self._put_msg(f'\n{"=" * 50}', 'DIM')
        self._put_msg(f'处理完成 — 成功: {success}  失败: {failed}  跳过: {skipped}', 'INFO')

        if success > 0:
            self._put_msg(f'\n生成的文件:')
            for r in results:
                if r['status'] == 'success':
                    self._put_msg(f'  ✓ {os.path.basename(r["output"])}', 'SUCCESS')

        # 保存 JSON 报告
        report_path = os.path.join(output_dir,
                                   f'生成报告_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump({
                'time': datetime.now().isoformat(),
                'version': __version__,
                'total': total,
                'success': success,
                'failed': failed,
                'skipped': skipped,
                'details': results,
            }, f, ensure_ascii=False, indent=2)
        self._put_msg(f'\n详细报告: {report_path}', 'DIM')

        # 尝试打开输出文件夹
        try:
            os.startfile(output_dir)
        except Exception:
            pass

        self._finish_task(f'完成 — 成功 {success} 个')

    # ══════════════════════════════════════════════
    #  辅助方法
    # ══════════════════════════════════════════════

    def _get_company_data(self) -> dict:
        """从 UI 字段获取公司数据"""
        return {
            'bidder_name': self._company_name.get(),
            'legal_representative': self._legal_rep.get(),
            'authorized_person': self._authorized_person.get(),
            'bidder_address': self._company_address.get(),
            'bidder_phone': self._company_phone.get(),
            'bidder_fax': self._company_fax.get(),
            'bidder_zip_code': self._company_zip.get(),
        }

    def _save_default_config(self, company: dict):
        """保存公司信息到默认配置文件"""
        config_dir = os.path.join(BASE_DIR, 'data')
        os.makedirs(config_dir, exist_ok=True)
        config_path = os.path.join(config_dir, 'company_profile.json')
        try:
            # 保留原有完整数据，只更新 company 字段
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    old_data = json.load(f)
            else:
                old_data = {}
            old_data['company'] = {**old_data.get('company', {}), **company}
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(old_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _log_msg(self, msg: str, tag: str = 'INFO'):
        """线程安全地记录日志"""
        self._queue.put(('log', (msg, tag)))

    def _put_msg(self, msg: str, tag: str = 'INFO'):
        """直接入队消息（工作线程调用）"""
        self._queue.put(('log', (msg, tag)))

    # ══════════════════════════════════════════════
    #  消息队列处理（GUI 线程）
    # ══════════════════════════════════════════════

    def _process_queue(self):
        """定期处理消息队列中的消息"""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                msg_type = msg[0]

                if msg_type == 'log':
                    text, tag = msg[1]
                    self._log.insert(tk.END, text + '\n', tag)
                    self._log.see(tk.END)

                elif msg_type == 'file_count':
                    count = msg[1]
                    if count > 0:
                        self._file_count_var.set(f'已发现 {count} 个招标文件')
                    else:
                        self._file_count_var.set('')

        except queue.Empty:
            pass

        self.root.after(100, self._process_queue)

    # ══════════════════════════════════════════════
    #  更新检查
    # ══════════════════════════════════════════════

    def _check_updates_silent(self):
        """启动时静默检查更新"""
        updater = AutoUpdater(__version__, __repo_url__)
        updater.check(self._on_update_result, timeout=5)

    def _check_updates_manual(self):
        """手动检查更新"""
        self._status_var.set('正在检查更新...')
        updater = AutoUpdater(__version__, __repo_url__)
        updater.check(self._on_manual_update_result, timeout=8)

    def _on_update_result(self, has_update: bool, latest_version: str = '',
                          download_url: str = '', changelog: str = '',
                          error: str = None):
        """静默更新结果回调"""
        if has_update:
            self._status_var.set(f'发现新版本 v{latest_version}')
            self._version_label.config(text=f'v{__version__} → v{latest_version}',
                                       foreground='#e07000')

            # 弹窗提示
            msg = f'当前版本: v{__version__}\n最新版本: v{latest_version}\n\n更新内容:\n{changelog[:500]}'
            if messagebox.askyesno('发现新版本',
                                   f'{msg}\n\n是否打开下载页面？'):
                self._open_url(download_url)

        elif error:
            self._status_var.set(f'更新检查失败: {error}')
        else:
            pass  # 最新版本，无需提示

    def _on_manual_update_result(self, has_update: bool, latest_version: str = '',
                                  download_url: str = '', changelog: str = '',
                                  error: str = None):
        """手动检查更新结果"""
        if error:
            messagebox.showinfo('检查更新', f'检查失败: {error}')
            self._status_var.set('就绪')
        elif has_update:
            msg = f'当前版本: v{__version__}\n最新版本: v{latest_version}\n\n更新内容:\n{changelog[:500]}'
            if messagebox.askyesno('发现新版本',
                                   f'{msg}\n\n是否打开下载页面？'):
                self._open_url(download_url)
            self._status_var.set(f'发现新版本 v{latest_version}')
        else:
            messagebox.showinfo('检查更新', f'当前已是最新版本 v{__version__}')
            self._status_var.set('就绪')

    def _show_about(self):
        """关于窗口"""
        messagebox.showinfo(
            '关于',
            f'{__app_name__}\n\n'
            f'版本: {__version__}\n'
            f'功能: 自动解析招标文件，生成投标响应文档\n\n'
            f'使用方法:\n'
            f'  1. 将招标文件 (.docx) 放入输入文件夹\n'
            f'  2. 填写公司信息\n'
            f'  3. 点击"一键生成响应文件"\n\n'
            f'输出: 自动生成的 DOCX 格式投标响应文件'
        )

    @staticmethod
    def _open_url(url: str):
        """在浏览器中打开 URL"""
        try:
            if sys.platform == 'win32':
                os.startfile(url)
            elif sys.platform == 'darwin':
                subprocess.run(['open', url])
            else:
                subprocess.run(['xdg-open', url])
        except Exception:
            pass  # 静默失败


def main():
    """启动 GUI 应用"""
    # 确保输入输出目录存在
    for d in ['input', 'output']:
        p = os.path.join(BASE_DIR, d)
        os.makedirs(p, exist_ok=True)

    root = tk.Tk()

    # 设置 DPI 感知（Windows 高分屏）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = BiddingApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
