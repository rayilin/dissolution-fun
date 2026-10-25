# -*- coding: utf-8 -*-
"""建立「奇妙的溶解」教學簡報（24 頁，16:9，水色系）"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from PIL import Image

ROOT = r'C:\Users\rayi\Documents\自然科學教學專欄\115-4_三上自然_奇妙的溶解'
IMG = os.path.join(ROOT, '教學網頁', 'media', 'img')
OUT = os.path.join(ROOT, '簡報', '奇妙的溶解_教學簡報.pptx')

BG = RGBColor(0xFA, 0xF7, 0xF0)        # 米白
INK = RGBColor(0x1D, 0x4E, 0x5F)      # 深水墨藍
INK_SOFT = RGBColor(0x6E, 0x8A, 0x94)
ORANGE = RGBColor(0xC8, 0x6A, 0x2B)   # 強調橘
WATER = RGBColor(0x2C, 0x7F, 0xB8)    # 水藍
SUGAR = RGBColor(0xD9, 0xA4, 0x41)    # 糖黃
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = '微軟正黑體'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def slide():
    s = prs.slides.add_slide(BLANK)
    bg = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = BG
    bg.line.fill.background(); bg.shadow.inherit = False
    return s


def tb(s, x, y, w, h, text, size=24, color=INK, bold=False, align=PP_ALIGN.LEFT,
       spacing=1.35, anchor=MSO_ANCHOR.TOP):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    lines = text.split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = color; r.font.name = FONT
        rPr = r.font._rPr
        rPr.set('lang', 'zh-TW')
        latin = rPr.find(qn('a:latin'))
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = OxmlElement('a:ea')
            (latin.addnext(ea) if latin is not None else rPr.append(ea))
        ea.set('typeface', FONT)
        cs = rPr.find(qn('a:cs'))
        if cs is None:
            cs = OxmlElement('a:cs')
            (ea.addnext(cs) if ea is not None else rPr.append(cs))
        cs.set('typeface', FONT)
    return box


def rule(s, y, w=2.2, color=ORANGE, x=0.9):
    bar = s.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Pt(4))
    bar.fill.solid(); bar.fill.fore_color.rgb = color
    bar.line.fill.background(); bar.shadow.inherit = False


def _find(code):
    for ext in ('jpg', 'gif'):
        p = os.path.join(IMG, 'm%s.%s' % (code, ext))
        if os.path.exists(p):
            return p
    return None


def pic(s, code, x, y, w, h, border=True):
    p = _find(code)
    if not p:
        return
    im = Image.open(p); iw, ih = im.size
    sc = min(w / iw, h / ih)
    tw, th = iw * sc, ih * sc
    shape = s.shapes.add_picture(p, Inches(x + (w - tw) / 2), Inches(y + (h - th) / 2),
                                 Inches(tw), Inches(th))
    if border:
        shape.line.color.rgb = RGBColor(0xE2, 0xDA, 0xC8)
        shape.line.width = Pt(1)


def pic_grid(s, codes, x, y, w, h, cols=None, gap=0.22, label=None):
    n = len(codes)
    cols = cols or min(n, 4)
    rows = (n + cols - 1) // cols
    cw = (w - gap * (cols - 1)) / cols
    ch = (h - gap * (rows - 1)) / rows
    for i, c in enumerate(codes):
        r, col = divmod(i, cols)
        pic(s, c, x + col * (cw + gap), y + r * (ch + gap), cw, ch)
    if label:
        tb(s, x, y + h + 0.12, w, 0.4, label, 14, INK_SOFT, align=PP_ALIGN.CENTER)


def head(s, no, title, sub=None):
    tb(s, 0.9, 0.62, 1.6, 0.7, no, 20, ORANGE, True)
    tb(s, 1.9, 0.5, 10.5, 0.9, title, 34, INK, True)
    rule(s, 1.42)
    if sub:
        tb(s, 1.9, 1.6, 10.5, 0.6, sub, 18, INK_SOFT)


def bullets(s, items, x=1.0, y=1.9, w=10.8, size=24, gap=1.9, color=INK):
    for i, t in enumerate(items):
        dot = s.shapes.add_shape(9, Inches(x), Inches(y + i * gap + 0.09), Inches(0.13), Inches(0.13))
        dot.fill.solid(); dot.fill.fore_color.rgb = ORANGE
        dot.line.fill.background(); dot.shadow.inherit = False
        tb(s, x + 0.35, y + i * gap, w, 0.7, t, size, color)


def footer(s, n):
    tb(s, 0.9, 6.95, 8, 0.4, '奇妙的溶解｜翰林三上自然 第4單元', 11, RGBColor(0xA9, 0xA0, 0x8E))
    tb(s, 11.4, 6.95, 1.0, 0.4, str(n), 11, RGBColor(0xA9, 0xA0, 0x8E), align=PP_ALIGN.RIGHT)


# ── 1 封面 ──────────────────────────────────────────────
s = slide()
tb(s, 0.9, 1.5, 6, 0.5, '翰林版｜三年級上學期｜自然科學', 18, INK_SOFT)
tb(s, 0.9, 2.05, 7.2, 1.5, '奇妙的溶解', 66, INK, True)
rule(s, 3.55, 2.6)
tb(s, 0.9, 3.85, 7, 0.6, '第 4 單元｜4 節課｜溶解・快慢與濃淡・生活應用', 24, INK_SOFT)
pic(s, '007', 8.4, 1.6, 3.6, 3.4)
tb(s, 0.9, 6.3, 8, 0.5, '115 學年度　11/30、12/3、12/7、12/10', 16, RGBColor(0xA9, 0xA0, 0x8E))

# ── 2 單元地圖 ──────────────────────────────────────────
s = slide(); head(s, 'MAP', '這 4 節課，我們要一起做這些事')
rows = [('1', '11/30 一', '什麼是溶解？', '觀察五招 → 溶解與不溶解'),
        ('2', '12/3 四', '溶解的快慢', '攪拌・水溫・顆粒大小'),
        ('3', '12/7 一', '溶液的濃淡', '量筒讀數・濃度・飽和'),
        ('4', '12/10 四', '溶解與生活', '廚房與Cleaning、測驗 10 題、線上對戰')]
rows[3] = ('4', '12/10 四', '溶解與生活', '廚房與清潔、測驗 10 題、線上對戰')
y = 1.75
for i, (n, d, t, sub) in enumerate(rows):
    box = s.shapes.add_shape(5, Inches(0.9), Inches(y), Inches(11.5), Inches(1.12))
    box.fill.solid(); box.fill.fore_color.rgb = WHITE if i % 2 == 0 else RGBColor(0xF3, 0xEE, 0xE1)
    box.line.color.rgb = RGBColor(0xE6, 0xDE, 0xCC); box.shadow.inherit = False
    tb(s, 1.05, y + 0.26, 0.7, 0.6, n, 28, ORANGE, True)
    tb(s, 1.7, y + 0.12, 2.1, 0.5, d, 15, INK_SOFT)
    tb(s, 1.7, y + 0.52, 4.6, 0.5, t, 22, INK, True)
    tb(s, 6.6, y + 0.36, 5.6, 0.5, sub, 17, INK_SOFT)
    y += 1.28
footer(s, 2)

# ── 3 第1節 目標 ───────────────────────────────────────
s = slide(); head(s, '4-1', '第 1 節｜什麼是溶解？', '今天的三個任務')
bullets(s, ['① 用五種感官觀察物質：摸、看、聞、放大鏡、嚐（老師確認才安全）',
            '② 動手實測，把結果記錄在學習單',
            '③ 找出規律：什麼東西會溶解？溶解是什麼意思？'])
footer(s, 3)

# ── 4 觀察五招 ──────────────────────────────────────────
s = slide(); head(s, '4-1', '觀察物質的五種方法', '先學會安全地觀察')
pic_grid(s, ['001', '002', '003', '004', '011'], 0.9, 1.85, 11.5, 3.3, cols=5)
tb(s, 0.9, 5.35, 11.5, 0.5, '用手摸　｜　用眼睛看　｜　用放大鏡　｜　搗聞氣味　｜　用嘴巴嘗', 18, INK_SOFT, align=PP_ALIGN.CENTER)
tb(s, 0.9, 5.95, 11.5, 0.7, '安全守則：不明的物質不摸、不聞、不嘗；\n聞味道要用鼻子聞、手輕扇，不是湊近吸。', 19, ORANGE, True)
footer(s, 4)

# ── 5 預測 ──────────────────────────────────────────────
s = slide(); head(s, '4-1', '先猜一猜，再動手', '你覺得哪幾個會溶解？')
pic_grid(s, ['007', '017', '015', '016', '012', '014', '018', '042'], 0.9, 1.8, 11.5, 3.9, cols=4)
tb(s, 0.9, 5.95, 11.5, 0.6, '食鹽・蜂蜜・麵粉・小石子・黑胡椒粒・胡椒粒・黑糖・芝麻', 17, INK_SOFT)
footer(s, 5)

# ── 6 公平測試 ──────────────────────────────────────────
s = slide(); head(s, '4-1', '怎麼測才公平？', '每個人都用同樣的方法，結果才能比較')
bullets(s, ['用同樣的水量（都用 100 毫升）',
            '用同樣的杯子，同樣的水溫',
            '每一樣都攪拌 10 秒，再看結果',
            '溶解了 → 打「✓」；還看得見顆粒 → 打「✗」'], gap=1.15)
footer(s, 6)

# ── 7 實驗結果 ──────────────────────────────────────────
s = slide(); head(s, '4-1', '實驗結果：分成兩堆')
tb(s, 0.9, 1.75, 5.4, 0.5, '✓ 會溶解', 22, WATER, True)
pic_grid(s, ['007', '018', '017', '014'], 0.9, 2.3, 5.4, 2.4, cols=4, gap=0.14)
tb(s, 7.0, 1.75, 5.4, 0.5, '✗ 不會溶解', 22, ORANGE, True)
pic_grid(s, ['006', '016', '042', '015'], 7.0, 2.3, 5.4, 2.4, cols=4, gap=0.14)
tb(s, 0.9, 5.05, 11.5, 0.6, '食鹽・黑糖・蜂蜜・胡椒粒　｜　細沙・小石子・芝麻・麵粉', 17, INK_SOFT)
tb(s, 0.9, 5.75, 11.5, 0.6, '你的那一組，結果也一樣嗎？', 19, INK, True)
footer(s, 7)

# ── 8 結論 ──────────────────────────────────────────────
s = slide(); head(s, '4-1', '結論：什麼叫「溶解」？')
tb(s, 1.0, 2.0, 11.0, 1.0, '溶解＝物質均勻地分散在水裡，看不見顆粒。', 32, INK, True)
bullets(s, ['溶解後，整杯水到處都一樣（溶液是均一的）',
            '物質沒有不見，總質量不變（可以秤重量證明）',
            '不溶解的東西會沉在杯底，水還是水'], y=3.3, size=22, gap=1.0)
tb(s, 1.0, 5.4, 11.0, 0.8, '所以：「看不見」不等於「消失」——這是最容易搞錯的地方！', 19, ORANGE)
footer(s, 8)

# ── 9 迷思 消失 ────────────────────────────────────────
s = slide(); head(s, '4-1', '常見迷思：溶解＝看不見＝消失？', '不對！物質只是散開了')
pic_grid(s, ['005', '007', '010'], 0.9, 1.85, 6.2, 3.4, cols=3)
tb(s, 7.4, 2.0, 5.2, 3.4,
   '把糖放進水裡，\n用眼睛看不見顆粒，\n但重量沒有減少。\n\n溶到水裡的糖，\n還是那些糖。', 21, INK, spacing=1.5)
tb(s, 0.9, 5.6, 11.5, 0.6, '小提醒：把鹽水倒掉，鹽會跟著水一起走——它沒有不見。', 18, INK_SOFT)
footer(s, 9)

# ── 10 第2節 目標 ──────────────────────────────────────
s = slide(); head(s, '4-2', '第 2 節｜溶解的快慢', '今天要找出三個祕訣')
bullets(s, ['同樣的糖，怎麼有的溶解快、有的慢？',
            '水溫、攪拌、顆粒大小，哪一個最有用？',
            '做一個公平的小實驗來證明'], gap=1.4)
footer(s, 10)

# ── 11 三個加速法 ───────────────────────────────────────
s = slide(); head(s, '4-2', '讓溶解變快的三個方法')
pic_grid(s, ['019', '021', '030'], 0.9, 1.85, 11.5, 3.3, cols=3)
tb(s, 0.9, 5.35, 11.5, 0.5, '① 攪拌　　② 提高水溫　　③ 顆粒變小（敲碎）', 20, INK, True, align=PP_ALIGN.CENTER)
tb(s, 0.9, 5.95, 11.5, 0.6, '想一想：哪一個方法對「溶解的速度」影響最大？我們來比一比。', 18, INK_SOFT)
footer(s, 11)

# ── 12 攪拌 ────────────────────────────────────────────
s = slide(); head(s, '4-2', '祕訣一：攪拌')
pic_grid(s, ['040', '041'], 0.9, 1.85, 7.4, 3.3, cols=2)
tb(s, 8.7, 1.95, 4.0, 3.6, '大力攪拌：\n水一直流動，\n糖很快散開。\n\n輕輕攪拌：\n溶解得比較慢。', 21, INK, spacing=1.45)
tb(s, 0.9, 5.45, 11.5, 0.7, '為什麼？水一直在動，糖才能到處「被水沖刷」。', 22, ORANGE, True)
footer(s, 12)

# ── 13 水溫 ────────────────────────────────────────────
s = slide(); head(s, '4-2', '祕訣二：提高水溫')
pic_grid(s, ['021', '022', '047', '048'], 0.9, 1.85, 11.5, 3.3, cols=4)
tb(s, 0.9, 5.35, 11.5, 0.5, '溫水　｜　加冰塊（冷水）　｜　提高水溫　｜　放進冰箱', 17, INK_SOFT, align=PP_ALIGN.CENTER)
tb(s, 0.9, 5.95, 11.5, 0.6, '結論：水越熱，溶解得越快。（放進冰箱反而最慢）', 22, ORANGE, True, align=PP_ALIGN.CENTER)
footer(s, 13)

# ── 14 顆粒大小 ────────────────────────────────────────
s = slide(); head(s, '4-2', '祕訣三：把顆粒變小')
pic_grid(s, ['028', '029', '030'], 0.9, 1.85, 7.0, 3.3, cols=3)
tb(s, 8.3, 2.0, 4.4, 3.2, '半匙　｜　一平匙　｜　尖尖一匙\n\n顆粒越小、\n接觸面積越大，\n溶解越快。', 21, INK, spacing=1.45)
tb(s, 0.9, 5.45, 11.5, 0.7, '所以把糖敲碎、把鹽磨細，會溶解得比較快。', 22, ORANGE, True)
footer(s, 14)

# ── 15 公平比較 ────────────────────────────────────────
s = slide(); head(s, '4-2', '要比較快慢，條件要一樣')
bullets(s, ['同樣的糖、同樣的量',
            '同樣的水量、同一個杯子',
            '同一個水溫',
            '只改「攪拌方式」或「水溫」其中一個'], gap=1.15)
tb(s, 0.9, 5.6, 11.5, 0.6, '一次只改一個條件，這樣才知道是誰造成的差別。', 20, ORANGE, True)
footer(s, 15)

# ── 16 第3節 目標 ──────────────────────────────────────
s = slide(); head(s, '4-2', '第 3 節｜溶液的濃淡', '今天要用「量」來說話')
bullets(s, ['怎麼用眼睛量水的體積？（量筒）',
            '讀量筒的時候，眼睛要放哪裡？',
            '多溶一點、少溶一點，差在哪裡？'], gap=1.4)
footer(s, 16)

# ── 17 量筒讀法 ────────────────────────────────────────
s = slide(); head(s, '4-2', '量筒要這樣看才對')
pic_grid(s, ['031', '033', '032', '044'], 0.9, 1.85, 11.5, 3.3, cols=4)
tb(s, 0.9, 5.35, 11.5, 0.5, '俯視（偏大）　｜　平視（正確）　｜　仰視（偏小）　｜　讀刻度示意', 17, INK_SOFT, align=PP_ALIGN.CENTER)
tb(s, 0.9, 5.95, 11.5, 0.6, '口訣：眼睛和凹面最低處一樣高。', 22, ORANGE, True, align=PP_ALIGN.CENTER)
footer(s, 17)

# ── 18 濃與稀 ──────────────────────────────────────────
s = slide(); head(s, '4-2', '有多濃？水多還是糖多？')
pic_grid(s, ['024', '025', '028', '029', '037', '038'], 0.9, 1.85, 11.5, 3.3, cols=6, gap=0.14)
tb(s, 0.9, 5.35, 11.5, 0.5, '水 50 毫升　｜　水 100 毫升　｜　半匙　｜　一平匙　｜　尖尖一匙（一）　｜　尖尖一匙（二）', 16, INK_SOFT, align=PP_ALIGN.CENTER)
tb(s, 0.9, 5.95, 11.5, 0.6, '同樣多的水，溶的糖越多 → 溶液越濃。', 22, ORANGE, True, align=PP_ALIGN.CENTER)
footer(s, 18)

# ── 19 飽和溶液 ────────────────────────────────────────
s = slide(); head(s, '4-2', '溶到不能再溶：飽和溶液')
pic_grid(s, ['043', '023', '020'], 0.9, 1.85, 7.4, 3.3, cols=3)
tb(s, 8.7, 1.95, 4.0, 3.6, '水裡的糖\n已經到極限了。\n\n再加油，\n糖就沉在杯底，\n不會再溶解。', 21, INK, spacing=1.45)
tb(s, 0.9, 5.45, 11.5, 0.7, '飽和＝「溶到最多」，不是「溶得最快」。', 22, ORANGE, True)
footer(s, 19)

# ── 20 溫度與溶解量 ────────────────────────────────────
s = slide(); head(s, '4-2', '溫度也會改變「能溶多少」')
pic_grid(s, ['026', '027', '036', '022'], 0.9, 1.85, 11.5, 3.3, cols=4)
tb(s, 0.9, 5.35, 11.5, 0.5, '水 20℃　｜　水 50℃　｜　降低水溫　｜　加冰塊', 17, INK_SOFT, align=PP_ALIGN.CENTER)
tb(s, 0.9, 5.95, 11.5, 0.6, '把飽和糖水放進冰箱，會析出結晶——因為溶得變少了。', 20, ORANGE, True, align=PP_ALIGN.CENTER)
footer(s, 20)

# ── 21 第4節 目標 ──────────────────────────────────────
s = slide(); head(s, '4-3', '第 4 節｜溶解與生活', '找一找：溶解藏在哪裡？')
bullets(s, ['廚房裡的哪些事，是溶解？',
            '清潔、洗衣為什麼有效？',
            '曬衣服、水溝不通，又是什麼原理？'], gap=1.4)
footer(s, 21)

# ── 22 廚房裡的溶解 ────────────────────────────────────
s = slide(); head(s, '4-3', '廚房裡的溶解')
pic_grid(s, ['054', '058', '059', '063', '064', '061', '076', '066'], 0.9, 1.8, 11.5, 3.4, cols=8, gap=0.14)
tb(s, 0.9, 5.45, 11.5, 0.6, '煮湯加鹽・咖啡加糖・開水加維他命・紅茶加糖・奶茶加珍珠・鮮奶加麥片・豆漿加糖・蛋花湯', 15, INK_SOFT, align=PP_ALIGN.CENTER)
tb(s, 0.9, 6.05, 11.5, 0.5, '哪些真的溶解了？哪些只是「泡著」？', 20, ORANGE, True, align=PP_ALIGN.CENTER)
footer(s, 22)

# ── 23 清潔與溶解 ──────────────────────────────────────
s = slide(); head(s, '4-3', '清潔也是溶解')
pic_grid(s, ['049', '052', '055', '071', '070', '079'], 0.9, 1.8, 11.5, 3.3, cols=6, gap=0.16)
tb(s, 0.9, 5.35, 11.5, 0.6, '洗衣膠囊・清潔劑除垢・小蘇打洗餐具・肥皂洗手・沖掉泥沙・通樂清水管', 16, INK_SOFT, align=PP_ALIGN.CENTER)
tb(s, 0.9, 6.0, 11.5, 0.6, '清潔劑溶解並帶走油汙，再用清水沖掉——這就是溶解的應用。', 19, ORANGE)
footer(s, 23)

# ── 24 統整 ────────────────────────────────────────────
s = slide(); head(s, '統整', '溶解三件事，你記得了嗎？')
bullets(s, ['① 溶解＝物質均勻分散在水裡，物質沒有消失',
            '② 攪拌、水溫高、顆粒小，溶解會變快',
            '③ 水少糖多就濃；溶到不能再溶就是飽和'], y=1.85, size=24, gap=1.15)
tb(s, 0.9, 5.35, 11.5, 0.5, '挑戰：打開教學網頁，做 10 題測驗，再找同學兩人對戰！', 20, ORANGE, True)
footer(s, 24)

prs.save(OUT)
print('已儲存', OUT, '頁數', len(prs.slides._sldIdLst))
