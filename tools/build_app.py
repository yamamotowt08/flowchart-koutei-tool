#!/usr/bin/env python3
"""Windows 用のアプリ（EXE 1 つ）と、配る ZIP を作る。

    python3 tools/build_app.py

1. tools/build_web.py でツール本体（HTML 版と app/tool.html）を作り直す
2. app/ の起動アプリ（.NET Framework 4.8 + WebView2）をビルドする。
   ツール本体と WebView2 の DLL は EXE の中に入る
3. dist/ に EXE と、手順書（docs/ にあれば）を入れた ZIP（英数字の名前）を作る。
   HTML 版はそのまま残す（EXE のダウンロードが止められる会社のパソコンでも使えるように）
4. フローチャートと工程表だけの版（履行報告書のタブなし）の EXE と ZIP も作る。
   app/ を app_fp/ に写し、名前・取り出し先（FLOWPLAN）と中のツール本体を変えてビルドする
"""
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EXE = ROOT / 'app' / 'bin' / 'Release' / 'net48' / 'FlowKouteiTool.exe'
APP_NAME = 'フローチャート・工程表作成ツール.exe'
DIST = ROOT / 'dist'
ZIP_NAME = 'FlowKouteiTool.zip'
FP_DIR = ROOT / 'app_fp'
FP_EXE = FP_DIR / 'bin' / 'Release' / 'net48' / 'FlowKouteiToolFP.exe'
FP_NAME = 'フローチャート・工程表作成ツール（フロー・工程表のみ）.exe'
FP_HTML = ROOT / 'フローチャート・工程表作成ツール（フロー・工程表のみ）.html'
FP_ZIP = 'FlowKouteiTool_FlowPlanOnly.zip'
MANUALS = ['フローチャート・工程表作成ツール_操作説明書.docx']


class ShiftJisName(zipfile.ZipInfo):
    """ZIP の中の日本語のファイル名を Shift_JIS で書く（Windows のエクスプローラーで文字化けしないように）。"""
    def _encodeFilenameFlags(self):
        return self.filename.encode('cp932'), self.flag_bits & ~0x800


def run(cmd, cwd):
    print('>', ' '.join(str(c) for c in cmd))
    subprocess.run(cmd, cwd=cwd, check=True)


run([sys.executable, 'tools/build_web.py'] + sys.argv[1:], ROOT)
run(['dotnet', 'build', '-c', 'Release', '--nologo'], ROOT / 'app')

# フローチャートと工程表だけの版：app/ の中身を写して、名前と取り出し先とツール本体を変える
FP_DIR.mkdir(exist_ok=True)
for name in ('Program.cs', 'app.manifest'):
    shutil.copy2(ROOT / 'app' / name, FP_DIR / name)
shutil.copy2(FP_HTML, FP_DIR / 'tool.html')
proj = (ROOT / 'app' / 'FlowKouteiTool.csproj').read_text(encoding='utf-8')
proj = proj.replace('<AssemblyName>FlowKouteiTool</AssemblyName>', '<AssemblyName>FlowKouteiToolFP</AssemblyName>\n    <DefineConstants>$(DefineConstants);FLOWPLAN</DefineConstants>')
proj = re.sub(r'<(Product|AssemblyTitle)>フローチャート・工程表作成ツール</', r'<\1>フローチャート・工程表作成ツール（フロー・工程表のみ）</', proj)
assert 'FLOWPLAN' in proj and proj.count('（フロー・工程表のみ）') == 2
(FP_DIR / 'FlowKouteiToolFP.csproj').write_text(proj, encoding='utf-8')
run(['dotnet', 'build', '-c', 'Release', '--nologo'], FP_DIR)

if DIST.exists():
    shutil.rmtree(DIST)
DIST.mkdir()
shutil.copy2(EXE, DIST / APP_NAME)
shutil.copy2(FP_EXE, DIST / FP_NAME)
manuals = [(ROOT / 'docs' / m, m) for m in MANUALS if (ROOT / 'docs' / m).exists()]
for zip_name, exe, exe_name in ((ZIP_NAME, EXE, APP_NAME), (FP_ZIP, FP_EXE, FP_NAME)):
    files = [(exe, exe_name)] + manuals
    with zipfile.ZipFile(DIST / zip_name, 'w') as z:
        for src, name in files:
            info = ShiftJisName.from_file(src, name)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, src.read_bytes())
    print(f'{exe_name} を作りました（{exe.stat().st_size // 1024} KB）')
    print(f'dist/{zip_name} を作りました（{(DIST / zip_name).stat().st_size // 1024} KB、{len(files)} ファイル）')
