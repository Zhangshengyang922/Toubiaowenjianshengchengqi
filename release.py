"""
版本发布脚本 —— 自动构建 + 发布到 GitHub

使用方法:
    python release.py [--no-build]

发布流程:
    1. 确认当前版本号
    2. 一键打包 (PyInstaller)
    3. 构建安装包 (Inno Setup)
    4. (可选) 创建 GitHub Release

注意事项:
    - 发布前请修改 version.py 中的 __version__
    - 请确保已将修改提交到 GitHub
    - 需要先安装 PyInstaller: pip install pyinstaller
"""
import os
import sys
import json
import argparse
import subprocess
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
from version import __version__, __app_name__, __repo_url__


def main():
    parser = argparse.ArgumentParser(description=f'{__app_name__} 发布脚本')
    parser.add_argument('--no-build', action='store_true', help='跳过构建，仅生成发布信息')
    parser.add_argument('--skip-installer', action='store_true', help='跳过安装包构建')
    args = parser.parse_args()

    print('=' * 60)
    print(f'  发布 {__app_name__} v{__version__}')
    print(f'  时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print('=' * 60)
    print()

    # ── 确认 ──
    resp = input(f'确认发布 v{__version__}? (y/n): ').strip().lower()
    if resp != 'y':
        print('已取消')
        return

    if not args.no_build:
        # ── 步骤1: PyInstaller 打包 ──
        print('\n[1/3] PyInstaller 打包...')
        result = subprocess.run([sys.executable, 'build.py'], cwd=BASE_DIR)
        if result.returncode != 0:
            print('打包失败，退出')
            sys.exit(1)

        # ── 步骤2: 构建安装包 ──
        if not args.skip_installer:
            print('\n[2/3] 构建安装包...')
            result = subprocess.run(
                [sys.executable, '-c', '''
import os, re, subprocess, sys
sys.path.insert(0, ".")
from version import __version__

iss_path = "installer.iss"
with open(iss_path, "r", encoding="utf-8") as f:
    content = f.read()

exe_name = f"BiddingDocGen_v{__version__}.exe"
content = re.sub(r'#define MyAppVersion ".*"', f'#define MyAppVersion "{__version__}"', content)
content = re.sub(r'#define MyAppExeName ".*"', f'#define MyAppExeName \"{exe_name}\"', content)

with open(iss_path, "w", encoding="utf-8") as f:
    f.write(content)

# 尝试调用 Inno Setup
iscc_paths = [
    r"C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe",
    r"C:\\Program Files\\Inno Setup 6\\ISCC.exe",
]
iscc = None
for p in iscc_paths:
    if os.path.exists(p):
        iscc = p
        break

if iscc:
    subprocess.run([iscc, iss_path], cwd=".")
    print(f"\\n安装包已构建: dist/BiddingDocGen_Setup_v{__version__}.exe")
else:
    print("\\nInno Setup 未安装，跳过安装包构建")
    print("请手动用 Inno Setup Compiler 打开 installer.iss 编译")
'''],
                cwd=BASE_DIR
            )
    else:
        print('\n跳过构建')

    # ── 步骤3: 生成发布信息 ──
    print('\n[3/3] 生成发布信息...')

    dist_dir = os.path.join(BASE_DIR, 'dist')
    assets = []
    for f in os.listdir(dist_dir) if os.path.exists(dist_dir) else []:
        if f'v{__version__}' in f:
            size = os.path.getsize(os.path.join(dist_dir, f))
            assets.append({'name': f, 'size_mb': round(size / 1024 / 1024, 1)})

    # 读取最近的 CHANGES 内容作为发布说明
    changelog = ''
    changes_path = os.path.join(BASE_DIR, 'CHANGES.md')
    if os.path.exists(changes_path):
        with open(changes_path, 'r', encoding='utf-8') as f:
            changelog_text = f.read()
        # 提取当前版本对应的更新内容
        import re
        match = re.search(r'## v[\d.]+.*?(?=## v|$)', changelog_text, re.DOTALL)
        if match:
            changelog = match.group(0).strip()

    print(f'\n{"=" * 60}')
    print('  GitHub Release 信息')
    print(f'{"=" * 60}')
    print()
    print(f'  Tag:     v{__version__}')
    print(f'  标题:    {__app_name__} v{__version__}')
    print()
    print(f'  资产文件:')
    for a in assets:
        print(f'    - {a["name"]} ({a["size_mb"]} MB)')

    # 生成发布用的 JSON
    release_info = {
        'tag_name': f'v{__version__}',
        'name': f'{__app_name__} v{__version__}',
        'body': changelog,
        'draft': False,
        'prerelease': False,
        'assets': [os.path.join(dist_dir, a['name']) for a in assets],
    }

    release_json_path = os.path.join(dist_dir, 'release_info.json')
    with open(release_json_path, 'w', encoding='utf-8') as f:
        json.dump(release_info, f, ensure_ascii=False, indent=2)

    print(f'\n  发布信息已保存: {release_json_path}')
    print()
    print(f'  接下来请:')
    print(f'    1. git push origin master')
    print(f'    2. 在 GitHub 上创建 Release: {__repo_url__}/releases/new')
    print(f'    3. Tag 填写: v{__version__}')
    print(f'    4. 上传 dist/ 中的文件作为资产')
    print()


if __name__ == '__main__':
    main()
