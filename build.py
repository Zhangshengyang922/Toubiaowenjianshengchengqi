"""
PyInstaller 打包脚本 —— 将 GUI 应用打包为单个 .exe 文件

使用方法:
    python build.py

输出文件在 dist/ 目录

打包前请确保:
    pip install pyinstaller
    pip install -r requirements-gui.txt
"""
import os
import sys
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
from version import __version__, __app_name__


def build():
    """执行 PyInstaller 打包"""
    # 输出名称
    app_pinyin = 'BiddingDocGen'
    output_name = f'{app_pinyin}_v{__version__}'

    # 需要打包的数据文件
    data_files = [
        ('data', 'data'),
        ('templates', 'templates'),
    ]

    # 排除不需要的模块（减小体积）
    exclude_modules = [
        'neo4j',
        'neo4j_driver',
        'openai',
        'langchain',
        'sklearn',
        'scipy',
        'numpy.core._dotblas',
        'matplotlib',
        'PIL',
        'tkinter.test',
    ]

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',           # 单文件 exe
        '--windowed',          # 不显示命令行窗口
        '--clean',             # 清理临时文件
        f'--name={output_name}',
        '--add-data', f'data{os.pathsep}data',
        '--add-data', f'templates{os.pathsep}templates',
    ]

    # 排除模块
    for mod in exclude_modules:
        cmd.extend(['--exclude-module', mod])

    # 隐藏导入（确保这些被打包）
    cmd.extend([
        '--hidden-import', 'lxml.etree',
        '--hidden-import', 'packaging.version',
        '--hidden-import', 'jinja2.ext',
    ])

    # 入口文件
    cmd.append('gui_app.py')

    print('=' * 60)
    print(f'  打包 {__app_name__} v{__version__}')
    print('=' * 60)
    print()
    print(f'  命令: {" ".join(cmd)}')
    print()

    # 执行打包
    result = subprocess.run(cmd, cwd=BASE_DIR)

    if result.returncode == 0:
        exe_path = os.path.join(BASE_DIR, 'dist', f'{output_name}.exe')
        print()
        print('=' * 60)
        print(f'  ✓ 打包成功!')
        print(f'  输出: {exe_path}')
        print('=' * 60)
    else:
        print()
        print('=' * 60)
        print(f'  ✗ 打包失败，退出码: {result.returncode}')
        print('=' * 60)
        sys.exit(result.returncode)


if __name__ == '__main__':
    build()
