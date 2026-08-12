#!/usr/bin/env python3
"""扫描版 PDF -> 保版式可编辑 .docx  (全本地, 无需联网/API)

流水线:  PyMuPDF 渲染 -> RapidOCR 识别(带坐标) -> 版面重建 -> python-docx

用法:
    python3 convert.py 文件.pdf [更多.pdf ...]     # 端到端
    python3 convert.py --rebuild                   # 仅用已有 OCR 缓存重建 docx(调版式时用)

依赖:  pip install pymupdf rapidocr onnxruntime python-docx
"""
import json, glob, os, re, sys
import argparse, collections
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CACHE = '.pdfrec_cache'          # 页图与 OCR 结果缓存(运行时按 PDF 所在目录重设)

# ---------- 可调参数 (GUI 参数面板直接映射到这里) ----------
DEFAULTS = {
    'long_edge':   2560,    # 渲染长边像素: 越大越准也越慢
    'header_y':    0.125,   # 此线以上视为页眉区
    'footer_y':    0.915,   # 此线以下视为页脚区
    'col_split':   0.50,    # 两列分界(占页宽): 右侧内容视为"值"列
    'line_tol':    0.45,    # 视觉行聚类: y 重叠比例阈值
    'x_tol':       0.025,   # 同列判定的 x 容差(占页宽)
    'full_line':   0.78,    # 行右端超过此比例才算"排满"(才可能有续行)
    'bullet_ind':  0.030,   # 缩进超出基准多少算项目符号
    'stamp_conf':  0.88,    # 低于此置信度的短串视为印章/手写, 剔除
    'drop_header': True,    # 剔除页眉
    'drop_footer': True,    # 剔除页脚
    'drop_stamp':  True,    # 剔除印章/签名噪声
    'page_marker': True,    # 插入"—— 原第 N 页 ——"分隔标记
    'two_col':     True,    # 启用两列->无边框表格还原
    'zh_font':     '宋体',
    'en_font':     'Times New Roman',
    'font_size':   10.5,
}
CFG = dict(DEFAULTS)

def G(k):
    return CFG.get(k, DEFAULTS[k])

FROZEN = getattr(sys, 'frozen', False)     # PyInstaller 打包后为 True

_last_tag = ''     # 最近一次转换的文件标签(GUI 预览据此定位页图)

class Cancelled(Exception):
    """用户中断转换"""

def _tick(progress, stage, cur, total, msg=''):
    if progress:
        progress(stage, cur, total, msg)

def _check(stop):
    if stop and stop():
        raise Cancelled()

def res_dir(*parts):
    """资源目录: 打包后指向解包目录 sys._MEIPASS, 否则指向脚本所在目录"""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)

def model_root():
    """离线模型目录; 打包时随 exe 分发, 避免首次运行联网下载"""
    for cand in (res_dir('models'), res_dir('rapidocr', 'models')):
        if os.path.isdir(cand) and glob.glob(os.path.join(cand, '*.onnx')):
            return cand
    return None                  # 未随包分发 -> 交给 rapidocr 自行下载

# ---------- 0. 渲染 + OCR ----------
def render_pdf(pdf, tag, progress=None, stop=None):
    """按页面实际旋转渲染为 PNG; 自动选 dpi 使长边 ~G('long_edge')"""
    import pymupdf
    d = pymupdf.open(pdf)
    outs = []
    n = d.page_count
    for i, p in enumerate(d):
        _check(stop)
        _tick(progress, 'render', i + 1, n, f'渲染第 {i+1}/{n} 页')
        out = f'{CACHE}/pages/{tag}_{i+1:03d}.png'
        outs.append(out)
        if os.path.exists(out):
            continue
        long_pt = max(p.rect.width, p.rect.height)
        dpi = max(150, min(300, round(G('long_edge') / long_pt * 72)))
        p.get_pixmap(dpi=dpi).save(out)     # get_pixmap 已应用 page.rotation
    d.close()
    return outs

_ENGINE = None

def get_engine():
    """OCR 引擎全局复用: 加载模型约 1-2 秒, 每次新建会拖慢批量转换"""
    global _ENGINE
    if _ENGINE is None:
        from rapidocr import RapidOCR
        root = model_root()
        _ENGINE = RapidOCR(params={'Global.model_root_dir': root} if root else None)
    return _ENGINE

def run_ocr(pngs, tag, progress=None, stop=None):
    todo = [f for f in pngs
            if not os.path.exists(f'{CACHE}/ocr/{os.path.basename(f)[:-4]}.json')]
    if todo:
        _tick(progress, 'ocr', 0, len(todo), '加载识别模型...')
        engine = get_engine()
        for n, f in enumerate(todo, 1):
            _check(stop)
            stem = os.path.basename(f)[:-4]
            _tick(progress, 'ocr', n, len(todo), f'识别 {stem} ({n}/{len(todo)})')
            r = engine(f)
            items = []
            if r is not None and r.txts:
                for txt, box, sc in zip(r.txts, r.boxes, r.scores):
                    xs = [float(q[0]) for q in box]; ys = [float(q[1]) for q in box]
                    items.append({'t': txt, 'x0': min(xs), 'y0': min(ys),
                                  'x1': max(xs), 'y1': max(ys), 's': float(sc)})
            items.sort(key=lambda d: (d['y0'], d['x0']))
            json.dump({'page': stem, 'items': items},
                      open(f'{CACHE}/ocr/{stem}.json', 'w'), ensure_ascii=False)
            if not progress:
                print(f'  OCR [{n}/{len(todo)}] {stem}  {len(items)} 行', flush=True)
    return sorted(glob.glob(f'{CACHE}/ocr/{tag}_*.json'))

HEADER_PAT = re.compile(
    r'^(TROESTER|EXCELLENCE IN EXTRUSION|Contract No\.|January \d+|Attachment \d|'
    r'De?a?tailed Technical Description|Page\s*\d+\s*/\s*\d+|正直诚信|山东泰开电缆有限公司$)')
FOOTER_PAT = re.compile(r'^(\d{1,3}\s*/\s*\d{1,3}$|\d{1,3}$|邮箱|地址|邮编|网址|http|tkdlsc)')

NUM_START = re.compile(
    r'^(\d{1,2}\.\d{1,2}(\.\d{1,2})?[\s、．.]|'      # 1.1  6.1.2
    r'\d{1,2}[\.、]\s*|'                             # 1.  2、 3、设备(无空格)
    r'[一二三四五六七八九十]+[、．.]|'                 # 一、
    r'[（(]\d{1,2}[）)]|'                            # (1)
    r'[A-Z][\.、]\s|'                                # A.
    r'[•·●○\*]\s*|'                                  # 项目符号
    r'[-–—]\s+)')                                    # 短横线子项
END_PUNCT = re.compile(r'[。．.：:；;!?！？）)】\]"”]$')
BULLET    = re.compile(r'^[•·●○\*]\s*|^[-–—]\s+')

def is_zh(s):
    return sum(1 for c in s if '一' <= c <= '鿿') >= 2

def clean(t):
    t = t.replace('　', ' ').strip()
    t = re.sub(r'\s{2,}', ' ', t)
    return t

# ---------- 1. 噪声过滤 ----------
def drop_noise(items, W, H):
    """剔除签名/印章/装订线噪声, 并把页眉页脚分离出来"""
    body, header, footer = [], [], []
    for it in items:
        t = clean(it['t'])
        if not t:
            continue
        ry0, ry1 = it['y0'] / H, it['y1'] / H
        # 手写签名/印章: 边缘区 + 低置信度 + 极短
        if G('drop_stamp') and (ry1 > 0.90 or ry0 < 0.05) and it['s'] < 0.88 and len(t) <= 3:
            continue
        if G('drop_stamp') and it['s'] < G('stamp_conf') and len(t) <= 4 \
                and not re.match(r'^[\dA-Za-z][\.、)）]?$', t):
            continue                           # 骑缝章/手写签名: 短串+低置信, 印刷体正文均 >0.95
        rec = dict(it, t=t, rx0=it['x0'] / W, rx1=it['x1'] / W, ry0=ry0, ry1=ry1)
        if G('drop_header') and ry1 < G('header_y') and HEADER_PAT.match(t):
            header.append(rec)
        elif G('drop_footer') and (ry0 > G('footer_y')
                                   or (ry0 > 0.85 and FOOTER_PAT.match(t))):
            footer.append(rec)
        else:
            body.append(rec)
    return body, header, footer

# ---------- 2. 视觉行聚类 ----------
def group_lines(items):
    """y 区间重叠 > G('line_tol') 的 item 合并为一个视觉行"""
    items = sorted(items, key=lambda d: (d['y0'], d['x0']))
    lines = []
    for it in items:
        placed = False
        for ln in lines:
            ov = min(ln['y1'], it['y1']) - max(ln['y0'], it['y0'])
            h = min(ln['y1'] - ln['y0'], it['y1'] - it['y0'])
            if h > 0 and ov / h > G('line_tol'):
                ln['items'].append(it)
                ln['y0'] = min(ln['y0'], it['y0']); ln['y1'] = max(ln['y1'], it['y1'])
                placed = True
                break
        if not placed:
            lines.append({'y0': it['y0'], 'y1': it['y1'], 'items': [it]})
    for ln in lines:
        ln['items'].sort(key=lambda d: d['x0'])
        ln['rx0'] = ln['items'][0]['rx0']
        ln['rx1'] = max(i['rx1'] for i in ln['items'])
        ln['ry0'] = ln['items'][0]['ry0']
        ln['h'] = ln['y1'] - ln['y0']
        # 内容起点: 首 item 若是窄编号(如 "1." "1.1"), 基准取下一个 item
        first = ln['items'][0]
        ln['cx0'] = (ln['items'][1]['rx0']
                     if len(ln['items']) > 1 and (first['rx1'] - first['rx0']) < 0.08
                     else first['rx0'])
    return sorted(lines, key=lambda l: l['y0'])

# ---------- 3. 块划分: 两列 vs 单列 ----------
def split_blocks(lines):
    """连续出现右列内容(x>G('col_split'))的区域 -> 两列块; 其余 -> 单列块"""
    blocks, cur, cur_kind = [], [], None
    for ln in lines:
        has_r = any(i['rx0'] >= G('col_split') for i in ln['items'])
        has_l = any(i['rx0'] < G('col_split') for i in ln['items'])
        kind = 'two' if has_r else 'one'
        # 只有右列没左列的行, 归属于前一个两列块(值的续行)
        if kind == 'one' and cur_kind == 'two' and not has_l:
            kind = 'two'
        if kind != cur_kind and cur:
            blocks.append((cur_kind, cur)); cur = []
        cur_kind = kind
        cur.append(ln)
    if cur:
        blocks.append((cur_kind, cur))
    return blocks

# ---------- 4. 单列块: 合并续行 ----------
def merge_paras(lines):
    """把被 OCR 拆散的续行合并回段落; 未排满的行绝不吸收下一行"""
    # 本页行距中位数 -> 空行阈值(A卷 ~1.5%, B卷 ~3.2%, 写死阈值必误伤其一)
    gaps = sorted(b['ry0'] - a['ry0'] for a, b in zip(lines, lines[1:])
                  if 0 < b['ry0'] - a['ry0'] < 0.08)
    lead = gaps[len(gaps) // 2] if gaps else 0.02
    gap_max = lead * 1.75
    paras = []
    for ln in lines:
        text = clean(' '.join(i['t'] for i in ln['items']))
        if not text:
            continue
        w = ln['rx1'] - ln['rx0']
        cur = {'text': text, 'rx0': ln['rx0'], 'cx0': ln['cx0'],
               'rx1': ln['rx1'], 'h': ln['h'], 'ry0': ln['ry0'],
               'center': abs((ln['rx0'] + ln['rx1']) / 2 - 0.5) < 0.07 and w < 0.5 and ln['rx0'] > 0.25}
        if not paras:
            paras.append(cur); continue
        prev = paras[-1]
        new_block = (
            not (prev['rx1'] > G('full_line')) or              # ★上一行没排满 -> 它已结束
            NUM_START.match(text) or                      # 新编号/新项目符号
            END_PUNCT.search(prev['text']) or             # 上一段已收尾
            is_zh(text) != is_zh(prev['text']) or         # 中英切换 -> 对照的另一半
            abs(cur['cx0'] - prev['cx0']) > G('x_tol') or      # 缩进层级变了
            (cur['ry0'] - prev['ry0']) > gap_max          # 行距明显变大 = 空行
        )
        if new_block:
            paras.append(cur)
        else:
            sep = '' if is_zh(prev['text']) else ' '
            prev['text'] = clean(prev['text'] + sep + text)
            prev['rx1'] = cur['rx1']
            prev['ry0'] = cur['ry0']          # ★ 推进基准行, 否则下一续行会被当空行断开
    return paras

def mark_bullets(paras):
    """OCR 会丢弃 • 图形符号: 用"缩进大于块基准"把列表项找回来

    基准按块(而非整页)取众数: 技术规格那种整块统一缩进的"标签—值"清单,
    块内众数就是它自己, 于是不会被误判成列表; 换成页级基准反而会把它们
    全标成 •。
    """
    if not paras:
        return paras
    base = collections.Counter(round(p['cx0'] * 200) / 200 for p in paras).most_common(1)[0][0]
    for p in paras:
        p['bullet'] = ((p['cx0'] - base) > G('bullet_ind')
                       and not NUM_START.match(p['text'])
                       and not p.get('center'))   # 居中大标题左缩进同样很大, 但不是列表项
    return paras

# ---------- 5. 两列块: 按左列标签分组 ----------
def build_rows(lines):
    """左列 item 开启一行, 其后到下一个左列标签之间的右列内容归入该行"""
    rows = []
    for ln in lines:
        left  = [i for i in ln['items'] if i['rx0'] <  G('col_split')]
        right = [i for i in ln['items'] if i['rx0'] >= G('col_split')]
        lt = clean(' '.join(i['t'] for i in left))
        rt = clean(' '.join(i['t'] for i in right))
        if lt or not rows:
            rows.append({'l': lt, 'r': [rt] if rt else [], 'indent': right[0]['rx0'] if right else 0})
        else:
            if rt:
                rows[-1]['r'].append(rt)
    return [r for r in rows if r['l'] or r['r']]

# ---------- 6. docx 输出 ----------
def set_font(run, zh_font=None, en_font=None, size=None, bold=False):
    zh_font = zh_font or G('zh_font'); en_font = en_font or G('en_font')
    size = G('font_size') if size is None else size
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = en_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), zh_font)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)

def add_para(doc, text, level=0, bold=False, size=None, style=None):
    p = doc.add_paragraph(style=style)
    pf = p.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(3)
    pf.line_spacing = 1.15
    if style is None and level:
        pf.left_indent = Cm(0.74 * level)
    r = p.add_run(text)
    set_font(r, size=size, bold=bold)
    return p

def borderless(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement(f'w:{edge}')
        e.set(qn('w:val'), 'none'); e.set(qn('w:sz'), '0')
        borders.append(e)
    tblPr.append(borders)

def heading_level(text):
    """按编号形态判定标题层级; 返回 None 表示正文"""
    t = text.strip()
    if re.search(r'[。；;]$', t):      # 完整句子 -> 正文, 不是标题
        return None
    if re.match(r'^[A-Z]\.\s+[A-Z]', t) and len(t) < 70:
        return 1
    if re.match(r'^\d{1,2}[\.、]\s*\S', t) and len(t) < 90 and not re.match(r'^\d{1,2}\.\d', t):
        return 2
    if re.match(r'^[一二三四五六七八九十]+[、．.]', t) and len(t) < 60:
        return 2
    if re.match(r'^\d{1,2}\.\d{1,2}[\s、．.]', t) and len(t) < 90:
        return 3
    if re.match(r'^\d{1,2}\s+[A-Z][a-z]', t) and len(t) < 70:
        return 3
    return None

def render_page(doc, blocks, page_label):
    for kind, lines in blocks:
        if kind == 'two' and G('two_col'):
            rows = build_rows(lines)
            if not rows:
                continue
            # 只有一行且左列为空 = 上页续行漂过来的孤立文字, 不是真两列
            if len(rows) == 1 and not rows[0]['l']:
                add_para(doc, ' '.join(rows[0]['r']))
                continue
            table = doc.add_table(rows=0, cols=2)
            table.alignment = WD_TABLE_ALIGNMENT.LEFT
            borderless(table)
            for row in rows:
                cells = table.add_row().cells
                for cell, txt, bold in ((cells[0], row['l'], True),
                                        (cells[1], '\n'.join(row['r']), False)):
                    cell.text = ''
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_after = Pt(1)
                    p.paragraph_format.line_spacing = 1.1
                    if txt:
                        r = p.add_run(txt)
                        set_font(r, size=10, bold=bold)
            table.columns[0].width = Cm(6.2)
            table.columns[1].width = Cm(10.3)
            doc.add_paragraph()
        else:
            paras = mark_bullets(merge_paras(lines))
            for i, para in enumerate(paras):
                t = para['text']
                if BULLET.match(t):
                    add_para(doc, BULLET.sub('', t), style='List Bullet'); continue
                if para.get('bullet'):
                    add_para(doc, t, style='List Bullet'); continue
                lv = heading_level(t)
                if lv:
                    add_para(doc, t, bold=True, size=12 - lv, level=0); continue
                # 短行 + 无收尾标点 + 下一行是长正文 => 无编号小标题
                nxt = paras[i + 1] if i + 1 < len(paras) else None
                if (para['rx1'] < 0.58 and not END_PUNCT.search(t) and len(t) < 60
                        and not NUM_START.match(t)
                        and nxt and nxt['rx1'] > G('full_line')):
                    add_para(doc, t, bold=True); continue
                if para.get('center'):
                    p = add_para(doc, t, bold=True, size=13)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    continue
                add_para(doc, t, level=1 if para['cx0'] > 0.16 else 0)

def make_doc(tag, title, pages, progress=None, stop=None):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for m in ('top_margin', 'bottom_margin'):
        setattr(sec, m, Cm(2.2))
    for m in ('left_margin', 'right_margin'):
        setattr(sec, m, Cm(2.5))
    st = doc.styles['Normal']
    st.font.name = G('en_font'); st.font.size = Pt(G('font_size'))
    st.element.rPr.rFonts.set(qn('w:eastAsia'), G('zh_font'))

    h = doc.add_paragraph(); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(h.add_run(title), size=16, bold=True)

    for idx, f in enumerate(pages):
        _check(stop)
        _tick(progress, 'layout', idx + 1, len(pages), f'重建版面 {idx+1}/{len(pages)}')
        d = json.load(open(f))
        from PIL import Image
        png = f'{CACHE}/pages/{os.path.basename(f)[:-5]}.png'
        W, H = Image.open(png).size
        body, header, footer = drop_noise(d['items'], W, H)
        lines = group_lines(body)
        blocks = split_blocks(lines)
        if G('page_marker'):
            marker = doc.add_paragraph()
            marker.paragraph_format.space_before = Pt(8)
            mr = marker.add_run(f"—— 原第 {idx+1} 页 ——")
            set_font(mr, size=8); mr.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            marker.alignment = WD_ALIGN_PARAGRAPH.CENTER
        render_page(doc, blocks, idx + 1)
    return doc

def set_cache(base_dir):
    """缓存放在 PDF 所在目录, 便于 --rebuild 与手工清理 (exe 可能在任意位置运行)"""
    global CACHE
    CACHE = os.path.join(base_dir or '.', '.pdfrec_cache')
    for sub in ('pages', 'ocr'):
        os.makedirs(os.path.join(CACHE, sub), exist_ok=True)
    return CACHE

def convert(pdf, progress=None, stop=None, out_dir=None, cfg=None):
    """转换单个 PDF -> docx, 返回输出路径

    progress(stage, cur, total, msg)  stage in {render, ocr, layout, done}
    stop() -> True 时中断并抛 Cancelled
    """
    if cfg:
        CFG.update(cfg)
    if not os.path.isfile(pdf):
        raise FileNotFoundError(pdf)
    pdf = os.path.abspath(pdf)
    set_cache(os.path.dirname(pdf))
    name = os.path.splitext(os.path.basename(pdf))[0]
    tag = re.sub(r'[^0-9A-Za-z\u4e00-\u9fff]', '', name)[:16] or 'DOC'
    global _last_tag
    _last_tag = tag
    if not progress:
        print(f'[{name}] 渲染中...')
    pngs = render_pdf(pdf, tag, progress, stop)
    if not progress:
        print(f'[{name}] {len(pngs)} 页, OCR 识别中(首次启动稍慢)...')
    jsons = run_ocr(pngs, tag, progress, stop)
    if not progress:
        print(f'[{name}] 重建版面...')
    doc = make_doc(tag, name, jsons, progress, stop)
    out = os.path.join(out_dir or os.path.dirname(pdf), name + '.docx')
    os.makedirs(os.path.dirname(out) or '.', exist_ok=True)
    doc.save(out)
    if not progress:
        print(f'  ✓ {out}  ({len(jsons)} 页, {os.path.getsize(out)/1024:.0f} KB)\n')
    _tick(progress, 'done', len(jsons), len(jsons), out)
    return out

def main():
    ap = argparse.ArgumentParser(
        description='扫描版 PDF -> 保版式可编辑 Word (全本地, 无需联网)')
    ap.add_argument('pdfs', nargs='*', help='待转换的扫描版 PDF (可拖拽到程序图标上)')
    ap.add_argument('--rebuild', metavar='DIR', nargs='?', const='.',
                    help='跳过渲染/OCR, 仅用该目录下的缓存重建 docx')
    a = ap.parse_args()

    if a.rebuild:
        set_cache(a.rebuild)
        tags = sorted({os.path.basename(f).rsplit('_', 1)[0]
                       for f in glob.glob(f'{CACHE}/ocr/*.json')})
        for tag in tags:
            js = sorted(glob.glob(f'{CACHE}/ocr/{tag}_*.json'))
            out = os.path.join(a.rebuild, f'{tag}.docx')
            make_doc(tag, tag, js).save(out)
            print(f'  ✓ {out}  <- {len(js)} 页 (缓存重建)')
        return

    pdfs = a.pdfs
    if not pdfs and FROZEN:
        # 双击运行(未拖文件): 给提示而不是一闪而过
        print('=' * 58)
        print('  扫描版 PDF -> 可编辑 Word    (全本地运行, 不联网)')
        print('=' * 58)
        print('\n用法: 把 PDF 文件直接拖到本程序图标上即可。\n')
        try:
            s = input('或在此粘贴 PDF 路径(多个用空格分隔, 直接回车退出): ').strip()
        except EOFError:
            s = ''
        pdfs = [p.strip(' "\'') for p in re.findall(r'"[^"]+"|\S+', s)] if s else []

    if not pdfs:
        ap.print_help()
    else:
        for p in pdfs:
            try:
                convert(p)
            except Exception as e:
                print(f'  ✗ 转换失败 {p}: {type(e).__name__}: {e}')
    if FROZEN:
        try:
            input('\n全部完成。按回车键关闭...')
        except EOFError:
            pass          # 管道/无 stdin 环境(如 CI)下不阻塞

if __name__ == '__main__':
    main()
