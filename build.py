# -*- coding: utf-8 -*-
"""
内网文档处理工作台 — 打包脚本

将源码打包为单文件可执行程序（Windows 下生成 .exe，macOS / Linux 生成同名可执行文件）。

用法:
    python build.py              # 构建单文件可执行程序
    python build.py --clean      # 先清理 build/ dist/ 再构建
    python build.py --onedir     # 生成目录版（启动更快，便于排查问题）

依赖:
    pip install pyinstaller
"""

import os
import shutil
import subprocess
import sys

APP_NAME = '内网文档处理工作台'
ENTRY = 'app.py'
ROOT = os.path.dirname(os.path.abspath(__file__))
SEP = ';' if os.name == 'nt' else ':'


def log(msg):
    print(f'[build] {msg}')


def clean():
    for d in ('build', 'dist', '__pycache__'):
        p = os.path.join(ROOT, d)
        if os.path.isdir(p):
            shutil.rmtree(p, ignore_errors=True)
            log(f'已清理 {d}/')
    spec = os.path.join(ROOT, f'{APP_NAME}.spec')
    if os.path.isfile(spec):
        try:
            os.remove(spec)
            log('已清理 spec 文件')
        except OSError as e:
            log(f'跳过 spec 清理（{e.__class__.__name__}），PyInstaller 将直接覆盖')


def build(onefile=True):
    args = [
        sys.executable, '-m', 'PyInstaller',
        '--noconfirm',
        '--clean',
        '--name', APP_NAME,
        '--onefile' if onefile else '--onedir',
        '--console',
        '--add-data', f'templates{SEP}templates',
        '--collect-data', 'docx',
        '--collect-data', 'pptx',
        '--exclude-module', 'tkinter',
        '--exclude-module', 'matplotlib',
        '--exclude-module', 'numpy',
        '--exclude-module', 'scipy',
        '--exclude-module', 'pandas',
        '--exclude-module', 'pytest',
        '--exclude-module', 'setuptools',
        '--exclude-module', 'pip',
        ENTRY,
    ]
    log('执行: ' + ' '.join(args))
    print('-' * 64)
    result = subprocess.run(args, cwd=ROOT)
    print('-' * 64)
    if result.returncode != 0:
        log('构建失败，请查看上方 PyInstaller 输出。')
        return result.returncode

    out_dir = os.path.join(ROOT, 'dist')
    if onefile:
        exe = os.path.join(out_dir, APP_NAME + ('.exe' if os.name == 'nt' else ''))
    else:
        exe = os.path.join(out_dir, APP_NAME)

    if os.path.exists(exe):
        if onefile and os.path.isfile(exe):
            size = os.path.getsize(exe) / 1024 / 1024
            log(f'构建完成: {exe}  ({size:.1f} MB)')
        else:
            log(f'构建完成: {exe}/')
        log('首次运行请放在有写入权限的目录，程序会在同级生成 config.json、uploads/、downloads/。')
        return 0

    log('构建结束但未找到产物，请检查 dist/ 目录。')
    return 1


def main():
    if '--clean' in sys.argv:
        clean()
    onefile = '--onedir' not in sys.argv
    sys.exit(build(onefile=onefile))


if __name__ == '__main__':
    main()
