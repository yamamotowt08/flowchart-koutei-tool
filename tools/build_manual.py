#!/usr/bin/env python3
"""操作説明書を Word で作る。絵は tools/make_shots.py で docs/img/ に作っておく。

    python3 tools/make_shots.py
    python3 tools/build_manual.py              # docs/フローチャート・工程表作成ツール_操作説明書.docx
    python3 tools/build_manual.py --pdf DIR    # 確認用に PDF も DIR に作る（LibreOffice が必要）

python-docx が必要（pip install python-docx）。
本文の [[ボタン]] は、ボタンのような灰色の印にする。
"""
import os
import re
import subprocess
import sys
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(ROOT, 'docs', 'img')
OUT = os.path.join(ROOT, 'docs', 'フローチャート・工程表作成ツール_操作説明書.docx')
FONT = 'BIZ UDPゴシック'
ACCENT = RGBColor(0x1D, 0x6A, 0x84)
BOX = {'point': ('FFF7D6', '補足'), 'warn': ('FDE4DA', '注意')}

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2)
sec.top_margin, sec.bottom_margin = Cm(1.8), Cm(1.8)


def set_font(target, size=None, bold=None, color=None):
    target.font.name = FONT
    rpr = target._element.get_or_add_rPr()
    fonts = rpr.find(qn('w:rFonts'))
    if fonts is None:
        fonts = OxmlElement('w:rFonts')
        rpr.append(fonts)
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        fonts.set(qn(a), FONT)
    if size: target.font.size = Pt(size)
    if bold is not None: target.font.bold = bold
    if color is not None: target.font.color.rgb = color


st = doc.styles['Normal']
set_font(st, 10.5)
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.35


def shade(el, fill):
    pr = el.get_or_add_tcPr() if el.tag.endswith('}tc') else el.get_or_add_rPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), fill)
    pr.append(shd)


def runs(p, text, size=None, bold=None, color=None):
    """[[ボタン]] はボタンの印、**太字** は太字にする。"""
    for part in re.split(r'(\[\[.+?\]\]|\*\*.+?\*\*)', text):
        if not part: continue
        if part.startswith('[['):
            r = p.add_run(f' {part[2:-2]} ')
            set_font(r, size, True, RGBColor(0x1E, 0x2A, 0x2D)); shade(r._element, 'DDE3E1')
        elif part.startswith('**'):
            r = p.add_run(part[2:-2]); set_font(r, size, True, color or ACCENT)
        else:
            r = p.add_run(part); set_font(r, size, bold, color)
    return p


def para(text, size=None, bold=None, color=None, align=None, after=None):
    p = runs(doc.add_paragraph(), text, size, bold, color)
    if align: p.alignment = align
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p


def bottom_line(p, color='1D6A84', sz=12):
    ppr = p._p.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr'); b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), str(sz)); b.set(qn('w:space'), '4'); b.set(qn('w:color'), color)
    bdr.append(b); ppr.append(bdr)


def h1(num, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14); p.paragraph_format.space_after = Pt(8); p.paragraph_format.keep_with_next = True
    r = p.add_run(f'{num}　' if num else ''); set_font(r, 18, True, ACCENT)
    r = p.add_run(text); set_font(r, 18, True)
    bottom_line(p)
    return p


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(4); p.paragraph_format.keep_with_next = True
    r = p.add_run('■ '); set_font(r, 13.5, True, ACCENT)
    r = p.add_run(text); set_font(r, 13.5, True)
    return p


def steps(items):
    for i, t in enumerate(items):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.9); p.paragraph_format.first_line_indent = Cm(-0.9)
        r = p.add_run(chr(0x2460 + i) + ' '); set_font(r, 13, True, ACCENT)
        runs(p, t)


def bullets(items):
    for t in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.7); p.paragraph_format.first_line_indent = Cm(-0.5)
        r = p.add_run('・'); set_font(r, None, True, ACCENT)
        runs(p, t)


def box(kind, lines):
    fill, label = BOX[kind]
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = t.cell(0, 0); c.width = Cm(17); shade(c._tc, fill)
    t.rows[0]._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))   # 見出しと本文をページで分けない
    p = c.paragraphs[0]; r = p.add_run(label); set_font(r, 11.5, True, ACCENT)
    for ln in lines:
        q = c.add_paragraph(); q.paragraph_format.space_after = Pt(2)
        runs(q, ln)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def fix_widths(t, widths):
    t.autofit = False
    for row in t.rows:   # 表の中は行間をつめる
        for c in row.cells:
            for q in c.paragraphs:
                q.paragraph_format.space_after = Pt(1); q.paragraph_format.line_spacing = 1.15
    for col, w in zip(t.columns, widths): col.width = w
    for row in t.rows:
        for c, w in zip(row.cells, widths): c.width = w


def img(name, width=17, caption=None):
    path = os.path.join(IMG, name + '.png')
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = bool(caption)
    p.add_run().add_picture(path, width=Cm(width))
    if caption: para(caption, 9.5, color=RGBColor(0x5B, 0x69, 0x6D), align=WD_ALIGN_PARAGRAPH.CENTER, after=8)


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def page_number_footer():
    p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(); set_font(r, 9)
    for kind, text in (('begin', None), (None, 'PAGE'), ('end', None)):
        if kind:
            f = OxmlElement('w:fldChar'); f.set(qn('w:fldCharType'), kind); r._r.append(f)
        else:
            t = OxmlElement('w:instrText'); t.set(qn('xml:space'), 'preserve'); t.text = text; r._r.append(t)


page_number_footer()

# ---------------- 表紙 ----------------
for _ in range(3): doc.add_paragraph()
para('フローチャート・工程表作成ツール', 24, True, ACCENT, WD_ALIGN_PARAGRAPH.CENTER, 8)
para('操作説明書', 22, True, None, WD_ALIGN_PARAGRAPH.CENTER, 36)
img('f_paper', 5.5)
para('バージョン 1.1.0', 10.5, None, None, WD_ALIGN_PARAGRAPH.CENTER, 18)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
for i, t in enumerate(['概要', '動作環境と起動', '画面構成', 'フローチャートの作成', '図の編集', 'PowerPoint への出力',
                       '工程表の作成', 'Excel への出力', '工程表からフローチャートへの反映', '入力の保存と読み込み', 'よくある質問']):
    q = para(f'{i + 1}．{t}', 10.5, after=1); q.paragraph_format.left_indent = Cm(5.5)
page_break()

# ---------------- 1 概要 ----------------
h1('1', '概要')
para('本ツールは、業務の手順を文章で入力するだけでフローチャート（業務フロー図）を作成し、同じ作業項目から工程表を作成するためのツールです。'
     'マクロを使用しないため、マクロの実行が制限されたパソコンでも利用できます。')
bullets([
    '**フローチャート**：手順の欄に 1 行ずつ入力すると図を自動で作成し、PowerPoint 形式（.pptx）で保存します。四角・矢印は PowerPoint の図形として編集できます。',
    '**工程表**：作業項目に開始日・終了日などを入力すると、月を上旬・中旬・下旬に分けた工程表（第２号様式）を作成し、Excel 形式（.xlsx）で保存します。バーは Excel の図形として編集できます。',
    '**相互の反映**：フローチャートの手順を工程表に取り込むこと、工程表の作業項目をフローチャートに反映することができます。',
])

# ---------------- 2 動作環境と起動 ----------------
h1('2', '動作環境と起動')
h2('動作環境')
bullets([
    'Windows 10 / 11',
    'EXE 版：.NET Framework 4.8 および Microsoft Edge WebView2 ランタイム（いずれも通常は Windows に標準で導入されています）',
    'HTML 版：Microsoft Edge または Google Chrome',
])
h2('EXE 版の起動')
steps([
    '配布された「FlowKouteiTool.zip」を右クリックし、[[すべて展開]] を選択します。',
    '展開したフォルダ内の「フローチャート・工程表作成ツール.exe」をダブルクリックします。',
])
h2('HTML 版の起動')
para('「フローチャート・工程表作成ツール.html」をダブルクリックし、Edge または Chrome で開きます。機能は EXE 版と同じです。')
box('warn', [
    '起動時に「Microsoft Edge WebView2 ランタイム」が入っていない旨のメッセージが表示された場合は、システム管理者に Microsoft Edge WebView2 ランタイムの導入を依頼してください。',
    'EXE ファイルのダウンロードや実行が制限されている環境では、HTML 版を使用してください。',
])

# ---------------- 3 画面構成 ----------------
h1('3', '画面構成')
img('f_header', 17, '画面上部のボタン')
TAB = [
    ('元に戻す／やり直す', '直前の操作を取り消す、または取り消した操作をやり直します。'),
    ('入力を保存／入力を開く', '手順と工程表の入力内容を 1 つのファイル（.json）に保存し、後から読み込みます。'),
    ('画像をコピー', 'フローチャートを画像としてクリップボードにコピーします（フローチャート画面のみ）。'),
    ('PowerPoint で保存', 'フローチャートを .pptx で保存します（フローチャート画面）。'),
    ('Excel で保存', '工程表を .xlsx で保存します（工程表画面）。'),
]
t = doc.add_table(rows=1, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, txt in zip(t.rows[0].cells, ['ボタン', '機能']):
    shade(c._tc, 'DDEFF4'); r = c.paragraphs[0].add_run(txt); set_font(r, 10, True)
for k, v in TAB:
    row = t.add_row().cells
    runs(row[0].paragraphs[0], k, 10, True); runs(row[1].paragraphs[0], v, 10)
fix_widths(t, [Cm(4.5), Cm(12.5)])
doc.add_paragraph()
para('タブの [[フローチャート]] と [[工程表]] で画面を切り替えます。フローチャート画面では、左側が手順の入力欄、右側が作成された図です。入力すると図は即座に更新されます。')
page_break()

# ---------------- 4 フローチャートの作成 ----------------
h1('4', 'フローチャートの作成')
img('e_bench', 17, '手順の入力欄（左）と作成される図（右）')
para('手順の欄に、作業を 1 行に 1 つずつ入力します。入力した順に上から矢印でつながります。入力欄の上のボタン（[[作業]] [[小項目]] [[照査・打合せ]] [[打合せ]] [[並行（左）]] [[吹き出し]]）を押すと、カーソル位置の下に該当する行を追加します。')
h2('入力の書き方')
NOT = [
    ('作業名', '作業（黄色の箱）', '1 行に 1 つ'),
    ('・作業名', '小項目（直前の作業の箱の中）', '「・」で始める'),
    ('@照査、打合せ', '照査・打合せ（右に分かれる箱）', '「、」で区切って複数可'),
    ('< 作業 > 合流先', '並行作業（左に分かれ、合流先の手順の前で合流）', '合流先は作業名または #番号'),
    ('判断？ > 処理 > #2', '判断（「いいえ」の場合の処理と戻り先）', '#番号は左端の青い数字'),
    ('| 作業A → 作業B', '同時進行（1 行が 1 列）', '「→」で同じ列の次の作業'),
    ('「名前」', '吹き出し（番号なし、名前は自由）', 'すぐ上の作業・小項目・照査・打合せに付く'),
    ('「名前」＞照査', '「@」の行の特定の箱に付ける吹き出し', '書かなければその行の最後の箱'),
    ('※ メモ', '図に表示しないメモ', ''),
]
t = doc.add_table(rows=1, cols=3); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, txt in zip(t.rows[0].cells, ['入力例', '意味', '補足']):
    shade(c._tc, 'DDEFF4'); r = c.paragraphs[0].add_run(txt); set_font(r, 10, True)
for a_, b_, c_ in NOT:
    row = t.add_row().cells
    for c, txt, bold in zip(row, [a_, b_, c_], [True, False, False]): runs(c.paragraphs[0], txt, 10, bold)
fix_widths(t, [Cm(4.5), Cm(7.5), Cm(5)])
doc.add_paragraph()
box('point', [
    '作業名に「照査」「打合せ」が含まれる行は、別の行へカーソルを移した時点で自動的に「@」の行（右に分かれる箱）に変換されます。小項目として「第1回」「第2回」のように回数を入力した場合は「打合せ①」「打合せ②」となります。',
    '箱の追加・削除・並べ替えによって手順の番号がずれた場合、「> #4」などの戻り先の番号は自動で付け直されます。',
])
h2('サンプルと白紙')
para('[[サンプルを入れる…]] から記入例を読み込めます。[[白紙にする]] を選択すると、手順を空にして作成を始められます。')

# ---------------- 5 図の編集 ----------------
h1('5', '図の編集')
para('右側の図はマウスで直接編集できます。図の編集内容は、左側の手順の欄にも反映されます。')
TAB = [
    ('箱を上下にドラッグ', '作業の順番を入れ替えます。'),
    ('箱を左右に少しドラッグ', '位置を調整します。矢印は常に直線のまま追従します。'),
    ('小項目を別の箱へドラッグ', 'その作業の小項目に移動します。何もない場所で離すと通常の作業になります。'),
    ('照査・打合せの箱や黒い点をドラッグ', '付く位置（見出しの横・小項目の間・作業の後）を変更します。'),
    ('吹き出しを箱の上へドラッグ', 'その箱（作業・小項目・照査・打合せ）に付け替えます。何もない場所で離すと位置だけ調整します。'),
    ('並行の分岐点・合流点をドラッグ', '並行作業の分岐元・合流先を変更します。'),
    ('矢印の端を別の箱へドラッグ', '接続先を変更します。'),
    ('線をクリック', '実線・点線の切り替え、または削除を選択します。'),
    ('箱をクリックして Delete キー', 'その作業を削除します。'),
]
t = doc.add_table(rows=1, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, txt in zip(t.rows[0].cells, ['操作', '結果']):
    shade(c._tc, 'DDEFF4'); r = c.paragraphs[0].add_run(txt); set_font(r, 10, True)
for k, v in TAB:
    row = t.add_row().cells
    runs(row[0].paragraphs[0], k, 10, True); runs(row[1].paragraphs[0], v, 10)
fix_widths(t, [Cm(6), Cm(11)])
doc.add_paragraph()
box('point', ['操作を誤った場合は [[元に戻す]]（Ctrl＋Z）で取り消せます。'])

# ---------------- 6 PowerPoint ----------------
h1('6', 'PowerPoint への出力')
img('f_opts', 15, 'スライドの大きさ・色・文字・タイトルの設定')
steps([
    '必要に応じて、スライドの大きさ（16:9／4:3／A4 縦／A4 横）、色、文字（フォント）、タイトルを設定します。',
    '[[PowerPoint で保存]]（Ctrl＋S）を押し、保存先とファイル名を指定して保存します。',
])
para('保存したファイルの四角・矢印は PowerPoint の図形として編集できます。矢印は図形に接続されているため、PowerPoint 上で箱を移動しても追従します。[[画像をコピー]] では、図を画像としてクリップボードにコピーできます。')
page_break()

# ---------------- 7 工程表の作成 ----------------
h1('7', '工程表の作成')
h2('作業項目の入力')
img('n_tasks', 17, '作業と日程の表')
steps([
    'タブの [[工程表]] を選択します。',
    '[[フローチャートの手順を取り込む]] を押すと、フローチャートの手順が作業項目として入ります。手入力する場合は [[＋ 作業を追加]] [[＋ 小項目を追加]] で行を追加します。',
    '各行の左端で「作業」（フローチャートの黄色の箱）か「小項目」（箱の中の作業）かを選択します。番号（1. ／1.1.）を使用している表では、番号が自動で付け直されます。',
])
h2('日程の入力')
bullets([
    '開始日・終了日・日数のうち 2 つを入力すると、残りの 1 つを自動計算して入力します（開始日を変更した場合は日数を保ったまま終了日を移動します）。',
    '日数のみを入力した行は、直前の作業の翌日から開始します。',
    '同じ作業に複数の期間がある場合（2 回目の照査など）は [[＋期間]] で行を追加します。',
    '日程を変更した場合は「変更後」の欄に入力します。工程表では上段に変更前、下段に変更後のバーを表示します。',
])
h2('行の並べ替え')
para('行の左端（⠿ と行番号の部分）をドラッグすると、青い線の位置に行を移動できます。作業を移動すると、その小項目と「＋期間」の行も一緒に移動します。[[↑]] [[↓]] ボタンでも同様に移動できます。')
h2('見出しと設定')
img('p_cols', 17, '工程表の見出しと設定')
TAB = [
    ('日数の数え方', '暦日（カレンダーどおり）または土日祝除く。祝日は自動で計算します。年末年始などは「ほかの休み」に入力します。'),
    ('バーの標準色・太さ', '工程表のバーの色と太さ。作業ごとの色は表の「バー色」で指定します。'),
    ('変更後バーの色・変更前バーの表示', '日程を変更した作業のバーの表示方法。'),
    ('期間が重なる作業', 'フローチャートへの反映時、直前の作業と期間が重なる作業を並行（左）に分けるかどうか。'),
]
t = doc.add_table(rows=1, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, txt in zip(t.rows[0].cells, ['設定', '内容']):
    shade(c._tc, 'DDEFF4'); r = c.paragraphs[0].add_run(txt); set_font(r, 10, True)
for k, v in TAB:
    row = t.add_row().cells
    runs(row[0].paragraphs[0], k, 10, True); runs(row[1].paragraphs[0], v, 10)
fix_widths(t, [Cm(5), Cm(12)])
page_break()

# ---------------- 8 Excel ----------------
h1('8', 'Excel への出力')
img('p_preview', 17, '画面下部のプレビュー（保存される工程表と同じ内容）')
para('[[Excel で保存]]（Ctrl＋S）を押し、保存先とファイル名を指定して保存します。業務名・作成者などの見出しを含む第２号様式の工程表が作成されます。'
     'バーは Excel の図形として編集でき、A4 横・1 ページ幅で印刷できるよう設定されています。')

# ---------------- 9 反映 ----------------
page_break()
h1('9', '工程表からフローチャートへの反映')
para('工程表で作業項目を変更した後、[[フローチャートに反映]] を押すと、フローチャートの手順に反映します。')
img('r_paper', 7.5, '工程表（記入例）から反映したフローチャート')
bullets([
    '作業項目の名称変更・追加・削除、および作業と小項目の区分の変更を反映します。',
    '手順を工程表の開始日の順に並べ替えます（開始・終了・判断および日付のない手順は移動しません）。',
    '直前の作業と期間が重なる作業は、並行（左）として分岐させます（設定で無効にできます）。',
    '照査・打合せは、工程表で配置した行の位置に対応する作業に付けます。',
    '白紙のフローチャートに反映した場合は、A4 縦・業務フロー（黄色）の書式とし、最後に「業務納品」を付けます。',
])
box('point', ['フローチャートと工程表の内容が異なる場合は、工程表画面の上部にお知らせが表示されます。どちらに合わせるかを選択できます。'])
page_break()

# ---------------- 10 保存 ----------------
h1('10', '入力の保存と読み込み')
bullets([
    '[[入力を保存]]：手順と工程表の入力内容を 1 つのファイル（.json）に保存します。',
    '[[入力を開く]]：保存したファイルを読み込み、続きから作業できます。',
    '最後に入力した内容は自動で保持され、次回起動時に表示されます。',
])
box('warn', ['自動で保持される内容はパソコンごと（EXE 版と HTML 版でも別）です。他のパソコンで使用する場合や重要なデータは、[[入力を保存]] でファイルに保存してください。'])

# ---------------- 11 よくある質問 ----------------
h1('11', 'よくある質問')
QA = [
    ('操作を取り消したい', '[[元に戻す]]（Ctrl＋Z）で取り消せます。'),
    ('以前の入力内容が表示される', '[[サンプルを入れる…]] から [[白紙にする]] を選択します（工程表も同様）。'),
    ('EXE がダウンロード・実行できない', 'HTML 版を使用してください。機能は同じです。'),
    ('WebView2 ランタイムがないと表示される', 'システム管理者に Microsoft Edge WebView2 ランタイムの導入を依頼してください。'),
    ('照査・打合せが右に分かれない', '別の行へカーソルを移すと変換されます。先頭の行、並行作業の中の行は変換されません。'),
    ('意図しない位置で並行（左）に分かれる', '工程表で直前の作業と期間が 1 日でも重なると分岐します。日程を確認するか、設定の「期間が重なる作業」を「分けない」にしてください。'),
]
t = doc.add_table(rows=1, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, txt in zip(t.rows[0].cells, ['症状', '対処']):
    shade(c._tc, 'DDEFF4'); r = c.paragraphs[0].add_run(txt); set_font(r, 10, True)
for q, a in QA:
    row = t.add_row().cells
    runs(row[0].paragraphs[0], q, 10, True); runs(row[1].paragraphs[0], a, 10)
fix_widths(t, [Cm(6), Cm(11)])

doc.save(OUT)
print(OUT)

# 確認用の PDF（--pdf DIR のときだけ。配るのは Word だけ）
if '--pdf' in sys.argv:
    out_dir = sys.argv[sys.argv.index('--pdf') + 1]
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', out_dir, OUT],
                   check=True, capture_output=True, timeout=180, env=dict(os.environ, HOME=out_dir))
    print(os.path.join(out_dir, os.path.basename(OUT)[:-5] + '.pdf'))
