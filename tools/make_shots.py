"""手順書（docs/）の絵を作る。ツールを headless Chrome で場面ごとに開き、決めた部分だけを切り出して docs/img/ に保存する。

    python3 tools/make_shots.py          # すべての場面
    python3 tools/make_shots.py plan     # 場面を選んで
"""
import os, subprocess, sys, tempfile, shutil, json, re, html as H
from PIL import Image
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'docs', 'img')
WORK = tempfile.mkdtemp()
page0 = open(os.path.join(ROOT, 'フローチャート・工程表作成ツール.html'), encoding='utf8').read()
COMMON = r"""
const pickSample = k => { const s = document.getElementById('sample'); s.value = k; s.dispatchEvent(new Event('change')); };
const pickPlan = k => { const s = document.getElementById('planSample'); s.value = k; s.dispatchEvent(new Event('change')); };
const hideToast = () => { const t = document.getElementById('toast'); if (t) t.hidden = true; document.querySelectorAll('.zoombar').forEach(z => z.style.display = 'none'); };
const rect = sel => { const el = typeof sel === 'string' ? document.querySelector(sel) : sel; const r = el.getBoundingClientRect(); return [r.left + scrollX, r.top + scrollY, r.width, r.height]; };
"""
# 場面：[操作, {切り出す名前: 要素}]
SCENES = {
  'flow': ("setView('flow'); pickSample('word');", {
      'f_whole': "document.querySelector('.app')", 'f_header': "'.bar'", 'f_tabs': "'.tabs'", 'f_steps': "'#flowView > section'",
      'f_opts': "'.opts'", 'f_paper': "'#paper'", 'f_inserts': "'.inserts'"}),
  'flow_edit': ("setView('flow'); pickSample('blank'); ta.value='設計計画\\n@照査、打合せ\\n現地踏査\\n基本検討\\n・条件の整理\\n・比較検討\\n< 対策設計 > 図面作成\\n施工計画\\n図面作成\\n報告書作成'; jumpRefs=null; ta.blur(); update();", {
      'e_steps': "'#flowView > section'", 'e_paper': "'#paper'", 'e_bench': "'#flowView'"}),
  'blank': ("setView('flow'); pickSample('blank');", {'b_steps': "'#flowView > section'", 'b_sample': "'#sample'"}),
  'plan': ("setView('plan'); pickPlan('koutei');", {
      'p_tasks': "'#planView > section'", 'p_preview': "'#planView .stage'", 'p_cols': "document.querySelector('#planView .plan-cols')", 'p_header': "'.bar'", 'p_tabs': "'.tabs'"}),
  'plan_new': (r"""setView('flow'); pickSample('blank'); setView('plan'); pickPlan('blank');
    sched.rows = [newRow({name:'設計計画', lv:'1', s:'2026-04-01', e:'2026-04-30', d:'30'}), newRow({name:'条件の整理', lv:'2', s:'2026-04-01', e:'2026-04-10', d:'10'}),
      newRow({name:'比較検討', lv:'2', s:'2026-04-11', e:'2026-04-30', d:'20'}), newRow({name:'図面作成', lv:'1', s:'2026-05-01', e:'2026-05-31', d:'31'}), newRow({name:'', lv:'2'})];
    renderTasks(); planUpdate();""", {'n_tasks': "'#planView > section'", 'n_preview': "'#planView .stage'"}),
  'report': ("setView('plan'); pickPlan('koutei'); sched.report.form = 'rep'; sched.report.month = '2026-06'; sched.report.actual['2026-05'] = '35'; setView('report');", {
      'h_fields': "document.querySelector('#repView .plan-cols')", 'h_chart': "'#repView .stage'", 'h_tabs': "'.tabs'"}),
  'report_g': ("setView('plan'); pickPlan('koutei'); sched.report.form = 'gyomu'; sched.report.month = '2026-06'; setView('report');", {
      'h_items': "'#repItemsSec'"}),
  'reflect': ("setView('flow'); pickSample('blank'); setView('plan'); pickPlan('koutei'); exportToFlow(); setView('flow');", {'r_paper': "'#paper'", 'r_bench': "'#flowView'"}),
}
W, HGT = 1280, 3000   # 位置を測るときと写すときは、同じ大きさの画面にする（図の大きさが画面の高さで変わるため）
def chrome(args, prof):
    return subprocess.run(['google-chrome', '--headless=new', '--no-sandbox', '--disable-gpu', f'--user-data-dir={prof}', '--hide-scrollbars',
                           '--virtual-time-budget=6000', '--allow-file-access-from-files'] + args, capture_output=True, timeout=180, text=True)
for name, (act, parts) in SCENES.items():
    if len(sys.argv) > 1 and name not in sys.argv[1:]: continue
    grab = ','.join(f'"{k}": rect({v})' for k, v in parts.items())
    js = ('window.addEventListener("load", () => setTimeout(() => { try {' + COMMON + act + '; hideToast(); window.scrollTo(0,0);'
          'document.body.setAttribute("data-rects", JSON.stringify({' + grab + ', "_h": document.documentElement.scrollHeight}));'
          '} catch (e) { document.body.setAttribute("data-rects", JSON.stringify({error: e.message})); } }, 600));')
    src = os.path.join(WORK, f'_{name}.html')
    open(src, 'w', encoding='utf8').write(page0.replace('</body>', f'<script>{js}</script>\n</body>', 1))
    prof = tempfile.mkdtemp(dir=WORK)
    dom = chrome([f'--window-size={W},{HGT}', '--dump-dom', 'file://' + src], prof).stdout
    m = re.search(r'data-rects="([^"]*)"', dom)
    rects = json.loads(H.unescape(m.group(1)))
    if 'error' in rects: print(name, 'ERROR', rects['error']); continue
    shutil.rmtree(prof, ignore_errors=True); prof = tempfile.mkdtemp(dir=WORK)
    full = os.path.join(WORK, f'{name}_full.png')
    chrome([f'--window-size={W},{HGT}', '--force-device-scale-factor=2', f'--screenshot={full}', 'file://' + src], prof)
    shutil.rmtree(prof, ignore_errors=True)
    im = Image.open(full)
    for k, r in rects.items():
        if k.startswith('_') or not r: continue
        x, y, w, h = [v * 2 for v in r]
        pad = 12
        box = (max(0, int(x - pad)), max(0, int(y - pad)), min(im.width, int(x + w + pad)), min(im.height, int(y + h + pad)))
        im.crop(box).save(os.path.join(OUT, f'{k}.png'), optimize=True)
        print(name, k, box[2] - box[0], 'x', box[3] - box[1])

shutil.rmtree(WORK, ignore_errors=True)
