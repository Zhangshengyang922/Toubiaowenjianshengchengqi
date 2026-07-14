"""
自动更新模块 —— 启动时检查 GitHub Releases 是否有新版本

工作原理:
1. 查询 GitHub Releases API 获取最新版本号
2. 与当前 version.py 中的 __version__ 对比
3. 如有新版本，弹窗提示用户下载
"""
import json
import urllib.request
import urllib.error
import ssl
import threading
from packaging import version as pkg_version


class AutoUpdater:
    """GitHub Releases 自动更新检查"""

    def __init__(self, current_version: str, repo_url: str):
        """
        Args:
            current_version: 当前版本号，如 "1.0.0"
            repo_url: GitHub 仓库地址，如 "https://github.com/user/repo"
        """
        self.current_version = current_version
        self.repo_url = repo_url.rstrip('/')
        # 从 repo_url 提取 owner/repo
        parts = self.repo_url.rstrip('/').split('/')
        if len(parts) >= 2:
            self.repo_slug = '/'.join(parts[-2:])
        else:
            self.repo_slug = None

    def check(self, callback, timeout: int = 5) -> None:
        """
        异步检查更新（非阻塞）
        
        Args:
            callback: 回调函数 callback(has_update, latest_version, download_url, changelog)
            timeout: 网络超时秒数
        """
        def _worker():
            try:
                result = self._check_sync(timeout)
            except Exception as e:
                result = {
                    'has_update': False,
                    'error': str(e)
                }
            callback(**result)

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _check_sync(self, timeout: int = 5) -> dict:
        """同步检查更新（在工作线程中调用）"""
        if not self.repo_slug:
            return {'has_update': False, 'error': '无效的仓库地址'}

        api_url = f"https://api.github.com/repos/{self.repo_slug}/releases/latest"

        # 创建忽略 SSL 验证的上下文（某些企业网络可能需要）
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        req = urllib.request.Request(api_url)
        req.add_header('Accept', 'application/vnd.github.v3+json')
        req.add_header('User-Agent', 'BiddingDocGen-Updater/1.0')

        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                data = json.loads(resp.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return {'has_update': False, 'error': '暂无发布版本'}
            return {'has_update': False, 'error': f'HTTP {e.code}'}
        except urllib.error.URLError as e:
            return {'has_update': False, 'error': f'网络错误: {e.reason}'}
        except Exception as e:
            return {'has_update': False, 'error': str(e)}

        # 提取信息
        latest_tag = data.get('tag_name', '').lstrip('v')
        download_url = data.get('html_url', self.repo_url + '/releases/latest')
        changelog = data.get('body', '暂无更新说明')

        # 解析资产下载链接
        assets = data.get('assets', [])
        installer_url = None
        for asset in assets:
            name = asset.get('name', '')
            if name.endswith('.exe') and 'setup' in name.lower():
                installer_url = asset.get('browser_download_url')
                break
        if not installer_url and assets:
            installer_url = assets[0].get('browser_download_url')

        # 版本比较
        try:
            has_update = pkg_version.parse(latest_tag) > pkg_version.parse(self.current_version)
        except Exception:
            # 简单字符串比较 fallback
            has_update = latest_tag != self.current_version

        return {
            'has_update': has_update,
            'latest_version': latest_tag,
            'download_url': installer_url or download_url,
            'changelog': changelog,
            'error': None,
        }
