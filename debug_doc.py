# -*- coding: utf-8 -*-
"""
运行方式：在项目目录下执行 python debug_doc.py
"""
import os, sys, platform, shutil, subprocess, tempfile, glob

# Windows 控制台编码修复
if platform.system() == 'Windows':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

print("=" * 50)
print("DocFlow .doc 诊断工具")
print("=" * 50)

print(f"\n[1] 系统信息")
print(f"  OS: {platform.system()} {platform.release()}")
print(f"  Python: {platform.python_version()}")
print(f"  CWD: {os.getcwd()}")

print(f"\n[2] soffice 查找")
candidates = [
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    '/usr/lib/libreoffice/program/soffice',
    r'C:\Program Files\LibreOffice\program\soffice.exe',
    r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    '/Applications/LibreOffice.app/Contents/MacOS/soffice',
]
found = []
for c in candidates:
    exists = os.path.isfile(c)
    exe = os.access(c, os.X_OK) if exists else False
    if exists:
        print(f"  {'✓' if exe else '!'} {c}  (executable={exe})")
        if exe:
            found.append(c)
    
which = shutil.which('soffice') or shutil.which('libreoffice')
print(f"  shutil.which: {which}")
print(f"  PATH: {os.environ.get('PATH','(empty)')}")

print(f"\n[3] Lock 文件检查")
lock_dirs = [
    os.path.expanduser('~/.config/libreoffice'),
    os.path.expanduser('~/Library/Application Support/LibreOffice'),
    os.path.expandvars(r'%APPDATA%\LibreOffice'),
]
for d in lock_dirs:
    if os.path.exists(d):
        locks = glob.glob(os.path.join(d, '**', '.~lock*'), recursive=True)
        print(f"  {d}: {locks or 'no locks'}")

if not found and not which:
    print("\n[!] 未找到 soffice，无法测试转换")
    print("    请安装 LibreOffice: https://www.libreoffice.org/download/")
    sys.exit(1)

soffice = found[0] if found else which

print(f"\n[4] soffice 版本")
try:
    r = subprocess.run([soffice, '--version'], capture_output=True, text=True, timeout=10)
    print(f"  {r.stdout.strip() or r.stderr.strip()}")
except Exception as e:
    print(f"  错误: {e}")

print(f"\n[5] 实际转换测试")
# 找一个 .doc 文件测试
doc_file = None
for root, _, files in os.walk('.'):
    for f in files:
        if f.lower().endswith('.doc') and not f.lower().endswith('.docx'):
            doc_file = os.path.join(root, f)
            break
    if doc_file:
        break

if not doc_file:
    print("  当前目录没有 .doc 文件，跳过转换测试")
else:
    print(f"  测试文件: {doc_file}")
    tmp_dir = tempfile.mkdtemp()
    user_profile = tempfile.mkdtemp()
    abs_file = os.path.abspath(doc_file)
    env = os.environ.copy()
    env['PATH'] = '/usr/bin:/usr/local/bin:' + env.get('PATH', '')
    cmd = [soffice,
           f'-env:UserInstallation=file://{user_profile}',
           '--headless', '--norestore', '--nofirststartwizard',
           '--convert-to', 'docx', '--outdir', tmp_dir, abs_file]
    print(f"  命令: {' '.join(cmd)}")
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)
        print(f"  返回码: {r.returncode}")
        print(f"  stdout: {r.stdout.strip()}")
        print(f"  stderr: {r.stderr[:300]}")
        print(f"  输出文件: {os.listdir(tmp_dir)}")
    except subprocess.TimeoutExpired:
        print("  超时（90秒）")
    except Exception as e:
        print(f"  异常: {e}")

print("\n" + "=" * 50)
