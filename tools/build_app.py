#!/usr/bin/env python3
"""Windows 用のアプリ（EXE 1 つ）と、配る ZIP を作る。

    python3 tools/build_app.py

1. tools/build_web.py でツール本体（HTML 版と app/tool.html）を作り直す
2. app/ の起動アプリ（.NET Framework 4.8 + WebView2）をビルドする。
   ツール本体と WebView2 の DLL は EXE の中に入る
3. dist/ に EXE と、手順書（docs/ にあれば）を入れた ZIP（英数字の名前）を作る。
   HTML 版はそのまま残す（EXE のダウンロードが止められる会社のパソコンでも使えるように）
"""
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
MANUALS = ['フローチャート・工程表作成ツール_つかいかた.docx', 'フローチャート・工程表作成ツール_つかいかた.pdf']


class ShiftJisName(zipfile.ZipInfo):
    """ZIP の中の日本語のファイル名を Shift_JIS で書く（Windows のエクスプローラーで文字化けしないように）。"""
    def _encodeFilenameFlags(self):
        return self.filename.encode('cp932'), self.flag_bits & ~0x800


def run(cmd, cwd):
    print('>', ' '.join(str(c) for c in cmd))
    subprocess.run(cmd, cwd=cwd, check=True)


run([sys.executable, 'tools/build_web.py'] + sys.argv[1:], ROOT)
run(['dotnet', 'build', '-c', 'Release', '--nologo'], ROOT / 'app')

if DIST.exists():
    shutil.rmtree(DIST)
DIST.mkdir()
shutil.copy2(EXE, DIST / APP_NAME)
files = [(EXE, APP_NAME)] + [(ROOT / 'docs' / m, m) for m in MANUALS if (ROOT / 'docs' / m).exists()]
with zipfile.ZipFile(DIST / ZIP_NAME, 'w') as z:
    for src, name in files:
        info = ShiftJisName.from_file(src, name)
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, src.read_bytes())
print(f'{APP_NAME} を作りました（{EXE.stat().st_size // 1024} KB）')
print(f'dist/{ZIP_NAME} を作りました（{(DIST / ZIP_NAME).stat().st_size // 1024} KB、{len(files)} ファイル）')
