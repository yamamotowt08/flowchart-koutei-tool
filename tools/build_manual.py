#!/usr/bin/env python3
"""手順書（つかいかたブック）を Word で作る。絵は tools/make_shots.py で docs/img/ に作っておく。

    python3 tools/make_shots.py
    python3 tools/build_manual.py      # docs/フローチャート・工程表作成ツール_つかいかた.docx（と PDF）

python-docx が必要（pip install python-docx）。PDF は LibreOffice（soffice）があれば作る。
文は小学生でも読めるように、短く、むずかしいことばには説明を付ける。
本文の [[ボタン]] は、ボタンのような灰色の印にする。
"""
import os
import re
import subprocess
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(ROOT, 'docs', 'img')
OUT = os.path.join(ROOT, 'docs', 'フローチャート・工程表作成ツール_つかいかた.docx')
FONT = 'BIZ UDPゴシック'
ACCENT = RGBColor(0x1D, 0x6A, 0x84)
BOX = {'point': ('FFF7D6', 'ポイント'), 'warn': ('FDE4DA', 'ちゅうい'), 'word': ('E2F0F5', 'ことばのいみ')}

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
set_font(st, 11.5)
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
for _ in range(5): doc.add_paragraph()
para('フローチャート・工程表 作成ツール', 26, True, ACCENT, WD_ALIGN_PARAGRAPH.CENTER, 6)
para('つかいかたブック', 30, True, None, WD_ALIGN_PARAGRAPH.CENTER, 18)
para('できることリスト と てじゅんしょ', 15, None, None, WD_ALIGN_PARAGRAPH.CENTER, 30)
img('f_paper', 7.5)
para('手順を書くと、フローチャート（流れの図）ができます。\n日付を入れると、工程表（予定の表）ができます。'.replace('\n', '　'), 12, None, None, WD_ALIGN_PARAGRAPH.CENTER)
page_break()

# ---------------- できることリスト ----------------
h1('', 'できることリスト')
para('このツールで できることです。やってみたら □ に ✓ を つけてみましょう。')
LIST = [
    ('フローチャート', '作業を 1 行に 1 つ書くと、黄色い箱の図ができる', '2'),
    ('フローチャート', '箱の中に「小項目」を入れる', '2'),
    ('フローチャート', '「照査」「打合せ」を、横に分かれた箱にする', '2'),
    ('フローチャート', '左に分かれて進む「並行」の作業を書く', '2'),
    ('フローチャート', '箱や黒い点をドラッグして、順番や場所をなおす', '3'),
    ('フローチャート', 'まちがえたら「元に戻す」', '3'),
    ('フローチャート', 'PowerPoint のファイルにする・画像をコピーする', '4'),
    ('工程表', 'フローチャートの手順を、工程表に取り込む', '5'),
    ('工程表', '「作業」か「小項目」かを選ぶ', '5'),
    ('工程表', '開始日・終了日を入れると、日数が自動で入る', '5'),
    ('工程表', '行をドラッグして、間に差し込む（番号も自動）', '5'),
    ('工程表', '日付が変わったときの「変更後」を書く', '5'),
    ('工程表', 'Excel のファイル（第２号様式）にする', '6'),
    ('両方', '工程表の内容を、フローチャートにもどす（反映）', '7'),
    ('両方', '入力をファイルに保存して、あとで続きから使う', '8'),
]
t = doc.add_table(rows=1, cols=4); t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.style = 'Table Grid'
widths = [Cm(1.0), Cm(3.4), Cm(10.8), Cm(1.8)]
for c, w, txt in zip(t.rows[0].cells, widths, ['✓', 'どこで', 'できること', '章']):
    c.width = w; shade(c._tc, 'DDEFF4')
    r = c.paragraphs[0].add_run(txt); set_font(r, 11, True); c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
for where, what, sec_no in LIST:
    row = t.add_row().cells
    for c, w in zip(row, widths): c.width = w
    for c, txt, al in zip(row, ['□', where, what, f'{sec_no} 章'], [1, 1, 0, 1]):
        r = c.paragraphs[0].add_run(txt); set_font(r, 11, txt == '□')
        if al: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
fix_widths(t, widths)
doc.add_paragraph()
box('word', [
    '**クリック**：マウスの左ボタンを 1 回「カチッ」と おすこと。',
    '**ダブルクリック**：左ボタンを すばやく 2 回 おすこと。',
    '**ドラッグ**：左ボタンを おしたまま マウスを うごかし、すきな所で はなすこと。',
    '**入力（にゅうりょく）**：キーボードで 文字や 数字を 打つこと。',
    '**保存（ほぞん）**：作ったものを ファイルにして、パソコンに のこすこと。',
])
page_break()

# ---------------- 1 はじめる ----------------
h1('1', 'はじめる')
h2('アプリ（EXE）で ひらく')
steps([
    'もらった「FlowKouteiTool.zip」を 右クリックして [[すべて展開]] を えらびます。',
    'できたフォルダの中の「フローチャート・工程表作成ツール.exe」を ダブルクリックします。',
    'まどが ひらいて、ツールが はじまります。',
])
box('warn', [
    '「WebView2 が入っていません」と出たら、パソコンの係の人に「Microsoft Edge WebView2 ランタイム」を入れてもらってください。Windows 10・11 には、たいてい はじめから入っています。',
    '会社のパソコンで EXE が ダウンロードできないときは、つぎの「HTML でひらく」を つかってください。できることは 同じです。',
])
h2('HTML で ひらく（EXE が つかえないとき）')
steps([
    '「フローチャート・工程表作成ツール.html」を ダブルクリックします。',
    'Edge（エッジ）か Chrome（クローム）で ひらきます。',
])
h2('がめんの みかた')
img('f_header', 17, '上のボタン：元に戻す・やり直す・入力を保存・入力を開く・保存')
img('f_tabs', 9, 'タブ：「フローチャート」と「工程表」を クリックで きりかえます')
bullets([
    '**左がわ**：「手順」を 書くところ。',
    '**右がわ**：できた図。書くと すぐに かわります。',
])
page_break()

# ---------------- 2 フローチャートを作る ----------------
h1('2', 'フローチャートを 作る')
h2('作業（黄色い箱）を 書く')
steps([
    'タブの [[フローチャート]] を クリックします。',
    '左の「手順」に、作業を **1 行に 1 つ** 書きます。上から じゅんばんに 矢印で つながります。',
    'ボタンの [[作業]] を おすと、カーソルの行の下に 1 行 ふえます。',
])
img('e_bench', 17, '左に書くと、右に図が できる')
h2('箱の中に「小項目」を 入れる')
para('作業の行の すぐ下に、**「・」**から 書きます。または ボタンの [[小項目]] を おします。')
para('れい：「基本検討」の下に「・条件の整理」「・比較検討」と書くと、黄色い箱の中に 小さい箱が 入ります。')
h2('「照査」「打合せ」を 横に分ける')
para('**「@」**から 書きます。れい：「@照査、打合せ」。ボタンの [[照査・打合せ]] [[打合せ]] でも 入ります。')
box('point', [
    '名前に「照査」「打合せ」が 入っていれば、ふつうの行に 書いても、ほかの行に うつったときに 自動で「@」の行に なおります。',
    '「第1回」「第2回」のように 回数を 小項目に 書くと、「打合せ①」「打合せ②」に なります。',
])
h2('左に分かれて すすむ「並行」')
para('**「< 作業 > 合流先」** と 書きます。れい：「< 対策設計 > 図面作成」は、すぐ上の作業と いっしょに 左で すすみ、「図面作成」の前で もどります。ボタンの [[並行（左）]] でも 入ります。')
h2('サンプル・白紙')
para('[[サンプルを入れる…]] から、れいを 入れられます。[[白紙にする]] を えらぶと、からっぽから はじめられます。')
box('point', ['左の「書き方」を クリックすると、書き方の 早見表が ひらきます。'])
page_break()

# ---------------- 3 図をなおす ----------------
h1('3', '図を なおす（ドラッグ）')
para('右の図は、マウスで うごかせます。うごかすと、左の「手順」も いっしょに かわります。')
bullets([
    '**箱を 上や下へ ドラッグ** → 作業の じゅんばんが かわります。',
    '**箱を 少しだけ よこへ ドラッグ** → 位置だけ かわります。線は まっすぐの ままです。',
    '**小項目を ほかの箱の上へ ドラッグ** → その箱の 小項目に なります。',
    '**照査・打合せの箱や 黒い点を ドラッグ** → 付く場所が かわります。',
    '**矢印の はしを ほかの箱まで ドラッグ** → つながる先が かわります。',
    '**線を クリック** → 「実線」「点線」「消す」を えらべます。',
    '**箱を クリックして Delete キー** → その作業を 消します。',
])
box('point', ['まちがえたら 上の [[元に戻す]] を おします。もう一度 すすめたいときは [[やり直す]] です。'])
page_break()

# ---------------- 4 PowerPoint ----------------
h1('4', 'PowerPoint に する')
img('f_opts', 15, 'スライドの大きさ・色・文字・タイトルを えらべます')
steps([
    '上の [[PowerPoint で保存]] を おします。',
    '保存する場所と 名前を きめて [[保存]] を おします。',
    'できたファイルを PowerPoint で ひらくと、四角や矢印を そのまま なおせます。',
])
para('[[画像をコピー]] を おすと、図が 画像として コピーされます。Word や メールに Ctrl + V で はりつけられます。')
page_break()

# ---------------- 5 工程表を作る ----------------
h1('5', '工程表を 作る')
h2('フローチャートの 手順を 取り込む')
steps([
    'タブの [[工程表]] を クリックします。',
    '[[フローチャートの手順を取り込む]] を おすと、手順が 作業項目として 入ります。',
])
h2('「作業」か「小項目」かを えらぶ')
img('n_tasks', 17, '左の「作業／小項目」で えらぶ。小項目は 少し右に さがる')
bullets([
    '**作業**（黄色）：フローチャートの 黄色い箱に なります。',
    '**小項目**：すぐ上の 作業の 箱の中に 入ります。',
    '下の [[＋ 作業を追加]] [[＋ 小項目を追加]] で 行を ふやせます。',
])
h2('日付を 入れる（日数は 自動）')
bullets([
    '**開始日と 終了日を 入れる** → 日数が 自動で 入ります。',
    '**開始日と 日数を 入れる** → 終了日が 自動で 入ります。',
    '日数を なおすと 終了日が、終了日を なおすと 日数が かわります。',
    '同じ作業が 2 回目・3 回目と あるときは、[[＋期間]] で 行を ふやします。',
    '日付が かわったときは、オレンジ色の **「変更後」** に 入れます。工程表に 2 本の バーが 出ます。',
])
h2('間に 差し込む・じゅんばんを かえる')
steps([
    '行を ふやしたら、その行の いちばん左（⠿ と 番号の所）を ドラッグします。',
    '青い線が 出た所で はなすと、そこに 入ります。',
])
box('point', [
    '作業を うごかすと、その下の 小項目も いっしょに うごきます。',
    '「1.」「1.1.」の 番号が ついているときは、番号も 自動で つけなおします。[[↑]] [[↓]] でも うごかせます。',
])
h2('見出しと 設定')
img('p_cols', 17, '業務名・作成者などの 見出しと、日数の数え方などの 設定')
para('「日数の数え方」で **暦日**（カレンダーどおり）か **土日祝除く** を えらべます。祝日は 自動で 入ります。')
page_break()

# ---------------- 6 Excel ----------------
h1('6', 'Excel に する')
img('p_preview', 17, '画面の下に、できあがりの 工程表が 出ます')
steps([
    '上の [[Excel で保存]] を おします。',
    '保存する場所と 名前を きめて [[保存]] を おします。',
    'Excel で ひらくと、表も バーも そのまま なおせます。A4 横で 印刷できます。',
])
page_break()

# ---------------- 7 反映 ----------------
h1('7', '工程表から フローチャートへ（反映）')
para('工程表で 作業を なおしたら、[[フローチャートに反映]] を おすと、フローチャートにも もどせます。')
img('r_paper', 8.5, '工程表（Excel のれい）から 反映した フローチャート')
bullets([
    '**名前を かえた行**：同じ箱の 名前が かわります。',
    '**ふやした行・けした行**：箱が ふえたり へったり します。',
    '**開始日の じゅん**に、手順が ならびかわります。',
    '**期間が かさなる作業**は、左に 分かれる「並行」に なります（設定で「分けない」にも できます）。',
    '照査・打合せは、工程表で 置いた所の 横に 付きます。',
])
box('point', ['フローチャートと 工程表が ちがうときは、工程表の上に お知らせが 出ます。どちらに あわせるか、ボタンで えらべます。'])
page_break()

# ---------------- 8 保存 ----------------
h1('8', '入力を 保存する・ひらく')
bullets([
    '[[入力を保存]]：手順と 工程表の 入力を、1 つの ファイルに 保存します。',
    '[[入力を開く]]：保存した ファイルを ひらいて、続きから つかえます。',
    'ツールを とじても、さいごに 書いていた 内容は のこっています。',
])
box('warn', ['べつの パソコンで つかうときや、だいじな 内容は、[[入力を保存]] で ファイルに しておきましょう。'])

# ---------------- 9 こまったとき ----------------
h1('9', 'こまったときは')
QA = [
    ('まちがえて 消してしまった', '上の [[元に戻す]] を おします。何回でも もどれます。'),
    ('前の 内容が 出てくる', '[[サンプルを入れる…]] から [[白紙にする]] を えらびます（工程表も 同じです）。'),
    ('EXE が ダウンロードできない', '「フローチャート・工程表作成ツール.html」を つかってください。できることは 同じです。'),
    ('「WebView2 が入っていません」と出る', 'パソコンの係の人に「Microsoft Edge WebView2 ランタイム」を 入れてもらってください。'),
    ('照査・打合せが 横に 分かれない', 'その行から べつの行へ うつると なおります。いちばん上の行や、小項目つきの「並行」の中では 分かれません。'),
    ('思っていない 所で 左に 分かれる', '工程表の 日付が 1 日でも かさなっていると 分かれます。日付を 見なおすか、設定で「分けない」に します。'),
]
t = doc.add_table(rows=1, cols=2); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, w, txt in zip(t.rows[0].cells, [Cm(6), Cm(11)], ['こまったこと', 'こうする']):
    c.width = w; shade(c._tc, 'DDEFF4'); r = c.paragraphs[0].add_run(txt); set_font(r, 11, True)
for q, a in QA:
    row = t.add_row().cells
    row[0].width, row[1].width = Cm(6), Cm(11)
    runs(row[0].paragraphs[0], q, 11, True)
    runs(row[1].paragraphs[0], a, 11)
fix_widths(t, [Cm(6), Cm(11)])

doc.save(OUT)
print(OUT)

# PDF も作る（LibreOffice があれば）
try:
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(OUT), OUT],
                   check=True, capture_output=True, timeout=180, env=dict(os.environ, HOME=os.environ.get('SOFFICE_HOME', os.environ.get('HOME', '/tmp'))))
    print(OUT[:-5] + '.pdf')
except Exception as e:
    print('PDF は作れませんでした:', e)
