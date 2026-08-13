#!/usr/bin/env python3
"""合成一份"扫描件"示例 PDF —— 内容全部虚构, 不含任何真实合同信息

生成的 sample_scanned.pdf 与真实扫描件同构: 纯图像、无文字层、带轻微旋转
与噪点, 用来演示 convert.py 的版面重建能力, 覆盖这几种典型版式:

    中英对照 / 编号层级(1. 1.1 三、) / 两列"标签—值" / 项目符号列表 /
    带框线的表格(含跨行/跨列合并单元格) / 页眉页脚 / 印章与手写签名噪声

用法:  python3 make_sample.py        # 就地生成 sample_scanned.pdf
依赖:  pillow pymupdf   (字体路径按 macOS 写, 其他平台改 ZH/EN 即可)
"""
import os
import random

import pymupdf
from PIL import Image, ImageDraw, ImageFilter, ImageFont

random.seed(7)                      # 固定随机种子, 保证可复现
W, H = 1654, 2339                   # A4 @200dpi
L = 150                             # 正文左边距
HERE = os.path.dirname(os.path.abspath(__file__))

ZH = '/System/Library/Fonts/Supplemental/Songti.ttc'
EN = '/System/Library/Fonts/Supplemental/Arial.ttf'


def _f(path, size, index=0):
    try:
        return ImageFont.truetype(path, size, index=index)
    except Exception:
        return ImageFont.truetype(path, size)


Fz  = lambda s: _f(ZH, s, 2)        # 宋体
Fzb = lambda s: _f(ZH, s, 3)        # 宋体粗
Fe  = lambda s: _f(EN, s)


def draw_all(d, items):
    for it in items:
        if it[0] == 't':
            _, x, y, txt, fnt = it
            d.text((x, y), txt, font=fnt, fill=(20, 20, 20))
        else:
            _, x0, y0, x1, y1 = it
            d.line((x0, y0, x1, y1), fill=(30, 30, 30), width=2)


def page1():
    im = Image.new('RGB', (W, H), (252, 252, 250))
    d = ImageDraw.Draw(im)
    b = [('t', 1180, 90, 'EXAMPLE CORP', Fe(40)),
         ('t', 1180, 140, 'SAMPLE  DOCUMENT', Fe(18)),
         ('t', 700, 230, '合    同', Fzb(38)),
         ('t', 600, 290, 'PURCHASING CONTRACT', Fe(34))]
    y = 400
    b += [('t', L, y, '合同号:', Fz(26)), ('t', L + 330, y, 'DEMO-20250001', Fe(26)),
          ('t', L, y + 40, 'CONTRACT NUMBER', Fe(24))]

    y = 530
    # 中英混排必须分两次画: Arial 没有汉字字形, 一次画完会出一排豆腐块
    b += [('t', L, y, '法定地址', Fz(26)), ('t', L + 116, y, '/ LEGAL ADDRESS BUYER', Fe(26)),
          ('line', L, y + 52, L + 600, y + 52)]   # 下划线离字底留够间距, 否则 OCR 会把汉字读糊
    y += 84
    for zh, en, v1, v2 in [('买方：', 'The Buyer', '示例科技有限公司', 'Example Technology Co., Ltd.'),
                           ('地址：', 'Address', '中国某省某市示例路 1 号', 'No.1 Sample Road, Demo City, China'),
                           ('电话：', 'Tel.', '0000-0000000', '')]:
        b += [('t', L, y, zh, Fz(26)), ('t', L + 330, y, v1, Fz(26)), ('t', L, y + 38, en, Fe(24))]
        if v2:
            b += [('t', L + 330, y + 38, v2, Fe(24))]
        y += 96

    y += 44
    b += [('t', L, y, '法定地址', Fz(26)), ('t', L + 116, y, '/ LEGAL ADDRESS SELLER', Fe(26)),
          ('line', L, y + 52, L + 610, y + 52)]
    y += 84
    for zh, en, v1, v2 in [('卖方：', 'The Seller', '示范设备制造股份公司', 'Demo Equipment Manufacturing AG'),
                           ('地址：', 'Address', '示例国示范市工业大道 8 号', '8 Industrial Ave, Sampleton, Demoland')]:
        b += [('t', L, y, zh, Fz(26)), ('t', L + 330, y, v1, Fz(26)),
              ('t', L, y + 38, en, Fe(24)), ('t', L + 330, y + 38, v2, Fe(24))]
        y += 96

    y += 100
    b += [('t', L, y, '签约地点及日期：', Fz(26)), ('t', L + 380, y, '示范市，2025 年 01 月 01 日', Fz(26)),
          ('t', L, y + 40, 'SIGNING PLACE AND DATE', Fe(24)),
          ('t', L + 380, y + 40, 'Demo City, January 1st, 2025', Fe(24)),
          ('t', 780, 2230, '1/3', Fe(24))]
    draw_all(d, b)

    # 手写签名 + 印章: 用来演示"低置信度短串"噪声剔除
    d.line((200, 2150, 260, 2100), fill=(40, 40, 90), width=3)
    d.line((230, 2100, 280, 2160), fill=(40, 40, 90), width=3)
    d.ellipse((1180, 180, 1260, 260), outline=(190, 40, 40), width=3)
    d.text((1196, 205), '样章', font=Fz(22), fill=(190, 40, 40))
    return im


def page2():
    im = Image.new('RGB', (W, H), (252, 252, 250))
    d = ImageDraw.Draw(im)
    b = [('t', 1180, 90, 'EXAMPLE CORP', Fe(40)), ('t', 1180, 140, 'SAMPLE  DOCUMENT', Fe(18))]
    y = 240
    b += [('t', L, y, '买卖双方通过协商同意根据如下条款签订本合同:', Fz(26)),
          ('t', L, y + 40, 'The Buyer and the Seller agree to sign the present contract under the following terms:', Fe(23))]

    y += 120
    b += [('t', L, y, '1.', Fe(26)), ('t', L + 90, y, '供货范围和价格 / SCOPE OF SUPPLY AND PRICE', Fzb(26)),
          ('line', L + 90, y + 40, L + 700, y + 40)]
    y += 64
    b += [('t', L, y, '1.1', Fe(26)), ('t', L + 90, y, '描述：', Fz(26)),
          ('t', L + 90, y + 40, '示例生产线核心部分 (00-000 kV, 样品), 1 套', Fz(26)),
          ('t', L + 90, y + 80, 'Description: The Kernel Parts of Demo-Line, 1 set', Fe(23))]
    y += 150
    b += [('t', L, y, '1.2', Fe(26)), ('t', L + 90, y, '合同总价为：￥ 000,000.00 元（大写：示例金额整）。', Fz(26)),
          ('t', L + 90, y + 40, 'The total contract value is CNY 000,000.00 (say: SAMPLE AMOUNT only).', Fe(23))]

    y += 140
    b += [('t', L, y, '2.', Fe(26)), ('t', L + 90, y, '技术数据 / TECHNICAL DATA', Fzb(26)),
          ('line', L + 90, y + 40, L + 480, y + 40)]
    y += 74
    # 两列"标签—值": 还原成无边框表格
    for k, v in [('Function', 'to guide the sample material'),
                 ('Place of installation', 'in front of the demo stand'),
                 ('Design', 'adjustable in height'),
                 ('Consisting of', 'rigid steel frame with linear guide'),
                 ('', '2 vertical rollers, adjustable'),
                 ('Display language', 'English/Chinese switchable'),
                 ('Codes and Standards', 'SAMPLE / DEMO, no further standards')]:
        if k:
            b += [('t', L, y, k, Fe(25))]
        b += [('t', 860, y, v, Fe(24))]
        y += 48

    y += 54
    # 正文排到右边界 -> 演示"续行合并"; 正文行数多于列表项 -> 块基准落在正文上,
    # 列表项才认得出来(真实文档的版式也是这样)
    for t in ['The scope of supply of the demonstration line is hereby defined in the present clause and covers all of the main',
              'components which are to be delivered by the Seller under this sample purchase contract, as listed in the following:']:
        b += [('t', L, y, t, Fe(24))]
        y += 44
    y += 6
    for it in ['Main distribution', 'Drive cabinets', 'Temperature control unit']:
        d.ellipse((L + 34, y + 11, L + 48, y + 25), fill=(20, 20, 20))   # OCR 会丢弃这个圆点
        b += [('t', L + 90, y, it, Fe(24))]
        y += 46
    y += 10
    for t in ['Every one of the components listed above is subject to the technical data given in clause 2 above and shall be',
              'delivered as one complete set. Any deviation from the agreed scope of supply shall be recorded in writing and',
              'then signed by both parties before the shipment takes place.']:
        b += [('t', L, y, t, Fe(24))]
        y += 44

    y += 66
    b += [('t', L, y, '三、付款方式', Fzb(26))]
    y += 52
    for t in ['1、合同签订支付 30% 预付款，金额为 000000.00 元（示例金额）。',
              '2、发货前支付 60% 提货款，金额为 000000.00 元（示例金额）。',
              '3、验收合格后支付 10% 质保金，金额为 00000.00 元（示例金额）。']:
        b += [('t', L + 20, y, t, Fz(25))]
        y += 48
    b += [('t', 780, 2230, '2/3', Fe(24))]
    draw_all(d, b)
    d.line((200, 2150, 255, 2105), fill=(40, 40, 90), width=3)
    return im


# 第 3 页的框线表: 列边界与行边界(像素)。演示三件事 ——
#   格内多行文字不被拆成多行、跨行合并(备注)、跨列合并(合计行)
GX = [150, 270, 660, 1010, 1150, 1360, 1504]
GY = [430, 500, 620, 690, 760, 830]


def page3():
    im = Image.new('RGB', (W, H), (252, 252, 250))
    d = ImageDraw.Draw(im)
    b = [('t', 1180, 90, 'EXAMPLE CORP', Fe(40)), ('t', 1180, 140, 'SAMPLE  DOCUMENT', Fe(18)),
         ('t', L, 250, '四、供货清单 / LIST OF SUPPLY', Fzb(26)),
         ('t', L, 330, '下表为本合同项下的供货明细，最终以附件一为准。', Fz(25))]
    # 竖线: 中间几条到"合计"行就停(合计行左半边是一个跨 4 列的合并格)
    for i, x in enumerate(GX):
        b += [('line', x, GY[0], x, GY[-1] if i in (0, 4, 6) else GY[-2])]
    # 横线: 中间两条不画进"备注"列 —— 那里是一个跨 3 行的合并格
    for i, y in enumerate(GY):
        b += [('line', GX[0], y, GX[5] if i in (2, 3) else GX[-1], y)]
    cells = [
        (0, 0, '序号', Fz(24)), (0, 1, '品名 / Item', Fz(24)), (0, 2, '规格型号', Fz(24)),
        (0, 3, '数量', Fz(24)), (0, 4, '金额（元）', Fz(24)), (0, 5, '备注', Fz(24)),
        (1, 0, '1', Fe(24)), (1, 2, 'DEMO-100', Fe(24)), (1, 3, '1', Fe(24)),
        (1, 4, '000.00', Fe(24)),
        (2, 0, '2', Fe(24)), (2, 1, '示例附件', Fz(24)), (2, 2, 'DEMO-200', Fe(24)),
        (2, 3, '2', Fe(24)), (2, 4, '000.00', Fe(24)),
        (3, 0, '3', Fe(24)), (3, 1, '示例备件包', Fz(24)), (3, 2, 'DEMO-300', Fe(24)),
        (3, 3, '1', Fe(24)), (3, 4, '000.00', Fe(24)),
    ]
    for r, c, t, f in cells:
        b += [('t', GX[c] + 16, GY[r] + 18, t, f)]
    # 一格两行: 老办法会把它拆成两行表格行, 照框线还原才留在同一格里
    b += [('t', GX[1] + 16, GY[1] + 12, '示例主机', Fz(24)),
          ('t', GX[1] + 16, GY[1] + 58, 'Demo Main Unit', Fe(23))]
    b += [('t', GX[5] + 16, GY[1] + 96, '按附件一', Fz(24))]        # 跨 3 行的合并格
    b += [('t', GX[0] + 16, GY[4] + 18, '合计 / TOTAL（含税）', Fz(24)),   # 跨 4 列的合并格
          ('t', GX[4] + 16, GY[4] + 18, '0000.00', Fe(24))]
    b += [('t', L, 900, '注：上表金额均为示例数据，不代表任何真实报价。', Fz(24)),
          ('t', 780, 2230, '3/3', Fe(24))]
    draw_all(d, b)
    return im


def scanify(im, angle):
    """模拟扫描: 轻微歪斜 + 噪点 + 轻微失焦"""
    im = im.rotate(angle, expand=False, fillcolor=(252, 252, 250), resample=Image.BICUBIC)
    px = im.load()
    for _ in range(9000):
        x, y = random.randrange(W), random.randrange(H)
        v = random.randrange(150, 235)
        px[x, y] = (v, v, v)
    return im.filter(ImageFilter.GaussianBlur(0.4))


def main():
    pages = [scanify(page1(), -0.35), scanify(page2(), 0.25), scanify(page3(), -0.2)]
    doc = pymupdf.open()
    tmp = []
    for i, im in enumerate(pages):
        p = os.path.join(HERE, f'_tmp_{i}.jpg')
        im.save(p, quality=88)
        tmp.append(p)
        page = doc.new_page(width=595, height=842)          # A4, 单位 pt
        page.insert_image(pymupdf.Rect(0, 0, 595, 842), filename=p)
    out = os.path.join(HERE, 'sample_scanned.pdf')
    doc.save(out)
    doc.close()
    for p in tmp:
        os.remove(p)
    print(f'已生成 {out}  ({os.path.getsize(out) // 1024} KB, {len(pages)} 页, 无文字层)')


if __name__ == '__main__':
    main()
