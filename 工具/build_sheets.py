# -*- coding: utf-8 -*-
"""產生「奇妙的溶解」學習單 2 份 ＋ 測驗卷（10 題＋答案解析）"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r'C:\Users\rayi\Documents\自然科學教學專欄\115-4_三上自然_奇妙的溶解'
OUT = os.path.join(ROOT, '教案')
os.makedirs(OUT, exist_ok=True)
INK = RGBColor(0x1D, 0x4E, 0x5F)
ORANGE = RGBColor(0xC8, 0x6A, 0x2B)
GREY = RGBColor(0x6E, 0x8A, 0x94)
FONT = '微軟正黑體'


def setup(doc, margin=1.8):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for k in ('top_margin', 'bottom_margin', 'left_margin', 'right_margin'):
        setattr(sec, k, Cm(margin))
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(11.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.3


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    el = OxmlElement('w:shd')
    el.set(qn('w:val'), 'clear'); el.set(qn('w:fill'), hexcolor)
    tcPr.append(el)


def para(doc, text='', size=11.5, bold=False, color=None, align=None,
         space_before=0, space_after=4, indent=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    r = p.add_run(text)
    r.font.size = Pt(size); r.font.bold = bold
    r.font.name = FONT
    r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    if color:
        r.font.color.rgb = color
    return p


def h(doc, text, size=14, space_before=10, color=None):
    return para(doc, text, size=size, bold=True, color=color or ORANGE,
                space_before=space_before, space_after=5)


def table(doc, rows, widths=None, header=True, size=10.5, aligns=None, row_h=None):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.text = ''
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            if aligns and j < len(aligns) and aligns[j] == 'c':
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(val))
            r.font.size = Pt(size); r.font.name = FONT
            r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
            if header and i == 0:
                r.font.bold = True
                r.font.color.rgb = INK
                shade(cell, 'F3EEE1')
    if widths:
        for j, w in enumerate(widths):
            for row in t.rows:
                row.cells[j].width = Cm(w)
    if row_h:
        for row in t.rows:
            row.height = Cm(row_h)
    return t


def header_row(doc, left='姓名：', mid='座號：', right='日期：　　　月　　　日'):
    table(doc, [[left, mid, right, '第 1 節　什麼是溶解？']],
          widths=[3.0, 2.6, 5.4, 6.4], header=False, size=10.5, aligns=['c', 'c', 'c', 'c'])


# ════════════════════════════════════════════════════════
# 學習單①　什麼是溶解？
# ════════════════════════════════════════════════════════
doc = Document(); setup(doc)
header_row(doc)
para(doc, '奇妙的溶解｜學習單①　觀察與溶解', 17, bold=True, color=INK,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

h(doc, '任務 1　用五種感官觀察物質（先看老師示範，注意安全）', space_before=2)
para(doc, '安全守則：不明的粉末不摸、不聞、不嘗；聞味道要用鼻子聞、手輕輕扇。', 10.5, color=GREY)
table(doc, [
    ['物質', '顏色', '摸起來', '聞起來', '放大鏡下看到', '放進水裡會怎樣？（先猜）'],
    ['食　鹽', '', '', '', '', ''],
    ['砂　糖', '', '', '', '', ''],
    ['黑　糖', '', '', '', '', ''],
    ['蜂　蜜', '', '', '', '', ''],
    ['麵　粉', '', '', '', '', ''],
    ['細　沙', '', '', '', '', ''],
    ['小石子', '', '', '', '', ''],
    ['芝　麻', '', '', '', '', ''],
    ['胡椒粉', '', '', '', '', ''],
], widths=[2.2, 1.8, 2.4, 2.4, 3.2, 5.4], size=9.5, row_h=1.05)

h(doc, '任務 2　實際測試（100 毫升水，攪拌 10 秒）')
table(doc, [
    ['物質', '溶解了（✓）', '沒溶解、沉在下面（✗）', '攪成糊狀', '我的理由'],
    ['食鹽', '', '', '', ''],
    ['砂糖', '', '', '', ''],
    ['蜂蜜', '', '', '', ''],
    ['細沙', '', '', '', ''],
    ['小石子', '', '', '', ''],
    ['芝麻', '', '', '', ''],
], widths=[2.6, 3.2, 4.4, 2.8, 4.4], size=9.5, row_h=1.0)

h(doc, '任務 3　我學到的「溶解」是什麼？')
para(doc, '我的句子：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿',
     11.5, space_before=4)
para(doc, '（提示：溶解之後，水看起來有什麼改變？物質還在嗎？）', 10.5, color=GREY)

h(doc, '任務 4　挑戰題（寫出理由）')
para(doc, '1. 糖放進水裡看不見了，糖是不是就不見了？為什麼？', 11.5)
para(doc, '　答：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', 11.5)
para(doc, '2. 細沙和小石子放進水裡，攪拌之後水會變混濁嗎？和糖水有什麼不同？', 11.5)
para(doc, '　答：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', 11.5)

h(doc, '任務 5　回家任務')
para(doc, '找 3 樣家裡「會溶解」的東西，寫下來並說明溶解前後的變化：', 11.5)
for i in range(3):
    para(doc, '　%d. ＿＿＿＿＿＿＿＿＿＿＿　溶解前後：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿' % (i + 1), 11.5)

doc.save(os.path.join(OUT, '學習單1_什麼是溶解.docx'))
print('已儲存 學習單1')

# ════════════════════════════════════════════════════════
# 學習單②　溶解的快慢與濃淡
# ════════════════════════════════════════════════════════
doc = Document(); setup(doc)
header_row(doc, right='日期：　　　月　　　日')
para(doc, '奇妙的溶解｜學習單②　溶解的快慢與濃淡', 17, bold=True, color=INK,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

h(doc, '任務 1　溶解速度大挑戰（同樣的糖、同樣的水量，只改一個條件）', space_before=2)
table(doc, [
    ['我做的條件', '糖溶解完大約幾秒', '看起來快還慢？', '我的發現'],
    ['不攪拌', '', '', ''],
    ['輕輕攪拌', '', '', ''],
    ['大力攪拌', '', '', ''],
    ['用冷水', '', '', ''],
    ['用溫水', '', '', ''],
    ['糖整塊放', '', '', ''],
    ['糖敲碎再放', '', '', ''],
], widths=[3.6, 3.6, 3.0, 7.2], size=9.5, row_h=0.95)

h(doc, '任務 2　量筒讀數練習（眼睛要和凹面最低處一樣高）')
table(doc, [
    ['要量多少', '我的讀數（毫升）', '檢查：平視嗎？'],
    ['20 毫升', '', ''],
    ['50 毫升', '', ''],
    ['100 毫升', '', ''],
], widths=[4.0, 4.0, 9.4], size=10, row_h=0.9, aligns=['c', 'c', ''])

h(doc, '任務 3　溶液的濃淡')
para(doc, '同樣 100 毫升的水，分別加入不同的糖量：', 11.5)
table(doc, [
    ['加入的糖', '半匙', '一平匙', '尖尖一匙'],
    ['甜的程度（淡／普通／甜）', '', '', ''],
    ['哪一杯最濃？為什麼？', '', '', ''],
], widths=[5.4, 4.0, 4.0, 4.0], size=9.5, row_h=0.95, aligns=['', 'c', 'c', 'c'])

h(doc, '任務 4　飽和溶液與溫度')
para(doc, '1. 一直加糖，直到有糖沉在杯底 → 這叫做＿＿＿＿＿＿。', 11.5)
para(doc, '2. 把這杯糖水放進冰箱，等 5 分鐘後看到什麼？＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', 11.5)
para(doc, '3. 為什麼會這樣？＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', 11.5)

h(doc, '任務 5　溶解設計師')
para(doc, '請設計一杯糖水，寫下你的條件與理由：', 11.5)
para(doc, '　我想要：＿＿＿＿（淡／普通／甜）　用 ＿＿＿ 毫升水 ＋ ＿＿＿ 匙糖', 11.5)
para(doc, '　理由：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', 11.5)
para(doc, '　我預測溶解完大約要 ＿＿＿ 秒，因為：＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿', 11.5)

doc.save(os.path.join(OUT, '學習單2_溶解快慢與濃淡.docx'))
print('已儲存 學習單2')

# ════════════════════════════════════════════════════════
# 測驗卷（10 題＋答案解析）
# ════════════════════════════════════════════════════════
QUIZ = [
    ('把食鹽放進水裡攪拌，鹽看不見了，這是什麼現象？', ['蒸發', '溶解', '凝固', '燃燒'], 1,
     '食鹽均勻地分散在水裡、看不見顆粒，就是「溶解」。'),
    ('鹽溶解在水裡之後，鹽的質量會怎樣？', ['變少', '不變', '變多', '變成水的質量'], 1,
     '溶解只是均勻分散，物質本身沒有消失，總質量不變。'),
    ('下列哪一種物質放到水裡不會溶解？', ['食鹽', '砂糖', '小石子', '黑糖'], 2,
     '小石子不溶於水、會沉在杯底；食鹽、砂糖、黑糖都會溶解。'),
    ('想聞一包粉末的味道，正確的作法是什麼？', ['鼻子湊近深吸一口', '用手輕輕扇，讓氣味飄到鼻子',
                                     '打開蓋子湊過去', '用嘴舔一下'], 1,
     '手輕輕扇動讓氣味飄到鼻子，既安全又不會吸進太多。'),
    ('用量的時候如果眼睛是往下俯視，讀出來的數字會怎樣？', ['偏小', '偏大', '剛好', '讀不出來'], 1,
     '俯視會看到比實際更高的刻度，所以讀數偏大；仰視則偏小。'),
    ('同樣多的水，一杯加一平匙糖、一杯加半匙糖，哪一杯比較濃？', ['加一平匙的', '加半匙的',
                                                       '一樣濃', '看杯子大小'], 0,
     '同樣多的水裡，溶解的糖越多，溶液越濃。'),
    ('一杯糖水已經溶到不能再溶了，這叫做什麼？', ['稀溶液', '飽和溶液', '冰水', '混合液'], 1,
     '溶到某一個程度就不能再溶，這就是「飽和溶液」。'),
    ('把一杯飽和的糖水放進冰箱冷藏，會有什麼現象？', ['糖溶解得更多', '溶在水裡的糖變少、會析出結晶',
                                          '水變成冰', '什麼都不變'], 1,
     '溫度降低時溶解的量變少，多出來的糖就析出結晶。'),
    ('下列哪一組做法可以讓糖溶解得更快？', ['加冷水、不攪拌', '加熱水、攪拌、敲碎糖',
                                 '加很多糖、不動', '把糖放進冰箱'], 1,
     '提高水溫、攪拌、敲碎增加接觸面積，三個方法都會讓溶解變快。'),
    ('咖啡裡加了糖，咖啡喝完質量會變少嗎？', ['會少很多', '不會少，是溶在咖啡裡', '會變甜', '看加多少糖'], 1,
     '溶解不改變總質量，糖只是均勻分散在咖啡裡。'),
]

doc = Document(); setup(doc)
para(doc, '翰林版 三年級上學期 自然科學　第 4 單元　奇妙的溶解', 11, color=ORANGE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6)
para(doc, '隨堂測驗卷（10 題）', 19, bold=True, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
table(doc, [['班級：　　　　　', '座號：　　　', '姓名：　　　　　　　　', '得分：　　　/ 10']],
      widths=[4.2, 2.6, 6.4, 4.2], header=False, size=11, aligns=['c', 'c', 'c', 'c'])
para(doc, '', 6)
for i, (q, o, a, e) in enumerate(QUIZ, 1):
    para(doc, '%d. %s' % (i, q), 11.5, bold=True, space_before=6, space_after=2)
    para(doc, '　（%s）%s　　　（%s）%s' % ('A', o[0], 'B', o[1]), 11, space_after=1)
    para(doc, '　（%s）%s　　　（%s）%s' % ('C', o[2], 'D', o[3]), 11, space_after=2)

doc.add_page_break()
para(doc, '答案與解析（老師版）', 16, bold=True, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
table(doc, [['題號', '答案', '解析']] +
          [[str(i), 'ABCD'[a], e] for i, (q, o, a, e) in enumerate(QUIZ, 1)],
      widths=[1.6, 1.8, 13.6], size=10, aligns=['c', 'c', ''])
para(doc, '', 6)
para(doc, '提醒：答案分布 A 3 題、B 4 題、C 2 題、D 1 題。', 10.5, color=GREY)

doc.save(os.path.join(OUT, '測驗卷_奇妙的溶解.docx'))
print('已儲存 測驗卷')
