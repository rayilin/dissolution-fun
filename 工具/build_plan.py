# -*- coding: utf-8 -*-
"""產生「奇妙的溶解」教案 docx（4 節課、13 節）"""
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


def setup(doc, margin=2.0):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for k in ('top_margin', 'bottom_margin', 'left_margin', 'right_margin'):
        setattr(sec, k, Cm(margin))
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(11.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.35


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


def h(doc, text, size=15, space_before=14):
    p = para(doc, text, size=size, bold=True, color=INK, space_before=space_before, space_after=6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '10')
    bottom.set(qn('w:space'), '3'); bottom.set(qn('w:color'), 'C86A2B')
    pbdr.append(bottom); pPr.append(pbdr)
    return p


def table(doc, rows, widths=None, header=True, size=10.5, aligns=None):
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
            if aligns and aligns[j] == 'c':
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
    return t


# ════════════════════════════════════════════════════════
doc = Document()
setup(doc)

para(doc, '翰林版 三年級上學期 自然科學', 11, color=ORANGE, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10)
para(doc, '奇妙的溶解', 26, bold=True, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, '教學活動設計（4 節課）', 13, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14)

h(doc, '一、基本資料', space_before=4)
table(doc, [
    ['單元名稱', '第 4 單元　奇妙的溶解', '教學年級', '三年級'],
    ['教材來源', '翰林版自然科學 3 上', '教學節數', '4 節（每節 40 分鐘）'],
    ['教學者', '　　　　　　　', '教學日期', '11/30（一）、12/3（四）、12/7（一）、12/10（四）'],
    ['教學地點', '自然教室', '教學對象', '三年級學生（分組 4 人一組，共 6 組）'],
    ['設計者', '　　　　　　　', '班級', '　　年　　班'],
], widths=[2.6, 7.2, 2.6, 4.6])

h(doc, '二、設計理念')
para(doc, '溶解是三年級學生最熟悉、卻也最容易被誤解的現象。生活中「糖不見了、鹽不見了」的經驗，'
          '常讓孩子誤以為物質「消失」了；也常把「攪拌」與「溶得多」混為一談，忽略「溶解速度」與'
          '「溶解總量」是兩件事。')
para(doc, '本單元以「觀察 → 公平實測 → 控���變因 → 量的估計 → 生活連結」的探究流程為主軸：'
          '第一節先用五種感官安全地觀察物質，建立「溶解＝均勻分散、物質沒有消失」的基準概念；'
          '第二節用小實驗找出「攪拌、水溫、顆粒大小」三個影響溶解速度的因素；'
          '第三節用量筒與量匙讓「濃淡、飽和」變得可以量；第四節回到廚房與清潔，'
          '把溶解與蒸發放進生活情境。')
para(doc, '依 115 學年度翰林版三上自然課程計畫研析，本單元安排於第 14～16 週，'
          '對應學習內容 INc-Ⅱ-2（生活中常見測量單位與度量）、INd-Ⅱ-2（物質分離與改變）、'
          'INe-Ⅱ-2（溫度影響物質在水中溶解量），並強調「具體操作優先、迷思澄清、定量實驗、'
          '數位工具輔助」。評量以學習單與實作紀錄為主、紙筆測驗為輔，'
          '並搭配教學網頁的互動實驗、10 題線上測驗與兩人搶答對戰，讓學生課後也能複習。')

h(doc, '三、學習重點')
para(doc, '（一）學習表現', 11.5, bold=True, color=ORANGE, space_before=6)
for s in ['tc-Ⅱ-1　能依據觀察或蒐集到的資料，簡單分辨、分類或比較自然現象。',
          'po-Ⅱ-1　能從日常經驗、學習活動、自然環境，進行觀察，進而能察覺問題。',
          'pe-Ⅱ-1　能了解一個因素改變可能造成的影響，進而預測活動的大致結果。',
          'pe-Ⅱ-2　能正確安全操作適合學習階段的物品、器材儀器，並能觀察和記錄。',
          'pa-Ⅱ-1　能運用簡單的分類、比較、分析與推論等方法，整理已有的資訊或數據。',
          'pc-Ⅱ-1　能以圖或文字妥善表述所觀察的現象或結果。']:
    para(doc, s, indent=0.5, space_after=2)

para(doc, '（二）學習內容', 11.5, bold=True, color=ORANGE, space_before=8)
for s in ['INc-Ⅱ-2　生活中常見測量單位與度量。',
          'INd-Ⅱ-2　物質分離與改變。',
          'INe-Ⅱ-2　溫度影響物質在水中溶解量。']:
    para(doc, s, indent=0.5, space_after=2)

para(doc, '（三）核心概念', 11.5, bold=True, color=ORANGE, space_before=8)
for s in ['溶解：物質均勻分散在水中，形成看不見顆粒的均一溶液；物質本身沒有消失，總質量不變。',
          '溶解速度：受攪拌、水溫、顆粒大小（接觸面積）影響。',
          '溶解總量：受水量、溫度影響；溶到極限稱為飽和溶液。',
          '濃淡：同樣多的水裡，溶解的物質越多，溶液越濃。']:
    para(doc, '・' + s, indent=0.5, space_after=2)

h(doc, '四、學習目標')
for s in ['1. 能說出「溶解」的意思：物質均勻分散在水中、看不見顆粒，而且物質沒有消失。',
          '2. 能用摸、看、聞（手扇法）、放大鏡等方法安全觀察物質的性質。',
          '3. 能分辨哪些物質會溶解於水、哪些不會，並說出判斷的理由。',
          '4. 能用「公平測試」的條件（等量、同溫、同容器、同樣攪拌）比較溶解的快慢。',
          '5. 能說出攪拌、提高水溫、顆粒變小三種方法會讓溶解變快。',
          '6. 能正確使用量筒（眼睛與凹面最低處平視）並讀出水的體積。',
          '7. 能比較溶液的濃淡，並知道溶到不能再溶叫做飽和溶液。',
          '8. 能從生活中找出溶解的應用（煮湯加鹽、咖啡加糖、洗衣精除汙），並說出理由。',
          '9. 能以圖或文字記錄觀察與實驗結果，並和同伴比較。']:
    para(doc, s, indent=0.4, space_after=3)

h(doc, '五、教材與教具')
table(doc, [
    ['類別', '內容'],
    ['教材', '翰林版自然科學 3 上 第 4 單元課本與習作本'],
    ['實驗材料', '食鹽、砂糖、黑糖、蜂蜜、麵粉、細沙、小石子、芝麻、胡椒粉、辣椒粉、洗衣精、'
                 '小蘇打、維他命粉、麥片、珍珠、咖啡粉、紅豆'],
    ['教具', '透明玻璃杯（每組 3 個）、量筒 100 mL、量匙、量杯、天平（共用）、溫度計、'
             '計時器、攪拌棒（湯匙）、冰塊、熱水壺（保溫瓶）、濾網、托盤、抹布'],
    ['資訊設備', '單槍或電視（簡報 24 頁）、平板或電腦（教學網頁）、計時器工具'],
    ['學習單', '學習單①（觀察與溶解）、學習單②（溶解快慢與濃淡）、測驗卷（10 題＋解析）'],
], widths=[2.8, 14.2])

h(doc, '六、教學流程')
para(doc, '（一）第 1 節（11/30 一）　什麼是溶解？', 12, bold=True, color=ORANGE, space_before=8)
table(doc, [
    ['時間', '教學活動', '重點', '教具'],
    ['5 分', '故事導入：媽媽煮湯加鹽，鹽到哪裡去了？請學生先猜。',
     '激發好奇，喚起經驗', 'pptx 2～3'],
    ['8 分', '觀察五招：摸一摸、看一看、用放大鏡、搗聞氣味、嘗味道（僅限老師確認可食用的物質）。',
     '建立安全觀察的規範', '粉末樣本、放大鏡'],
    ['12 分', '分組實測：8 種物質各放 100 毫升水，攪拌 10 秒，記錄溶解／不溶解。',
     '公平測試、留下紀錄', '杯子、量筒、食鹽…'],
    ['8 分', '結果分類與討論：把物質分成會溶解、會攪成糊、不會溶解三堆。',
     '溶解與溶解的差別', '學習單①'],
    ['5 分', '教師提問：看不見顆粒＝不見了嗎？引出「物質沒有消失」。',
     '澄清迷思一', 'pptx 8～9'],
    ['2 分', '回家任務：找三樣會溶解的東西，寫在學習單。', '延伸觀察', '學習單①'],
], widths=[1.5, 8.6, 3.6, 3.3])

para(doc, '（二）第 2 節（12/3 四）　溶解的快慢', 12, bold=True, color=ORANGE, space_before=10)
table(doc, [
    ['時間', '教學活動', '重點', '教具'],
    ['5 分', '回顧：哪些會溶解？哪些不會？請學生說理由。', '喚起前一節概念', 'pptx 7'],
    ['5 分', '提出問題：同樣的糖，為什麼有的溶解快、有的慢？',
     '把問題變成可研究的', 'pptx 10'],
    ['15 分', '實驗一：攪拌（大力量 vs 輕攪）；實驗二：水溫（溫水 vs 冷水）；'
              '實驗三：顆粒大小（整塊 vs 敲碎）。每組做其中一項。',
     '控制變因、記錄時間', '糖、水、溫度計、計時器'],
    ['8 分', '全班歸納：三個方法都能加快溶解，其中水溫與顆粒大小的影響特別明顯。',
     '建立規律', 'pptx 11～15'],
    ['5 分', '教師補充：加快溶解 ≠ 溶得更多（量的問題留到下一節）。',
     '埋伏筆：速度與總量', 'pptx 12'],
    ['2 分', '回家任務：找出家裡「可以加快溶解」的例子。', '生活連結', '學習單②'],
], widths=[1.5, 8.6, 3.6, 3.3])

para(doc, '（三）第 3 節（12/7 一）　溶液的濃淡', 12, bold=True, color=ORANGE, space_before=10)
table(doc, [
    ['時間', '教學活動', '重點', '教具'],
    ['6 分', '用量筒裝水：示範俯視、平視、仰視三種看法，請學生猜數字。',
     '量筒的正確讀法', '量筒、pptx 17'],
    ['8 分', '小組練習：量出 20、50、100 毫升的水，記錄體積。',
     'INc-Ⅱ-2 度量', '量筒、量杯'],
    ['12 分', '比較濃淡：同樣 100 毫升水，分別加半匙、一平匙、尖尖一匙糖，比較甜度與外觀。',
     '濃度＝溶解的量', '糖、量匙、杯子'],
    ['8 分', '飽和溶液：不停加糖直到沉澱出現，再放進冰箱看析出結晶。',
     '溶解總量與溫度', '糖、冰塊、冰箱、pptx 19～20'],
    ['4 分', '整理：溶解速度 vs 溶解總量（兩張圖比較）。', '概念分流', 'pptx 20'],
    ['2 分', '回家任務：用「幾匙糖＋多少水」描述家裡的蜂蜜水。', '量化表達', '學習單②'],
], widths=[1.5, 8.6, 3.6, 3.3])

para(doc, '（四）第 4 節（12/10 四）　溶解與生活＋測驗對戰', 12, bold=True, color=ORANGE, space_before=10)
table(doc, [
    ['時間', '教學活動', '重點', '教具'],
    ['8 分', '廚房裡的溶解：煮湯加鹽、咖啡加糖、奶茶加珍珠、鮮奶加麥片——哪些溶解？哪些只是泡著？',
     '區分溶解與混合', 'pptx 22'],
    ['8 分', '清潔與溶解：洗衣精、小蘇打、肥皂怎麼把油汙帶走？',
     '溶解的應用', 'pptx 23'],
    ['6 分', '溶解與蒸發：曬衣服、玩水槍、塗油漆，液態的水到哪裡去了？',
     'INd-Ⅱ-2 物質分離與改變', 'pptx 23'],
    ['8 分', '隨堂測驗 10 題（個人完成＋同組互檢），錯題立即訂正。',
     '形成性評量', '測驗卷'],
    ['8 分', '教學網頁：互動實驗（溶解速度模擬、量筒讀數）＋線上兩人對戰。',
     '課後延伸與增強', '平板、pptx 24'],
    ['2 分', '回家任務：自製一杯「有設計」的糖水，寫下濃淡與溶解量。', '應用與創造', '學習單②'],
], widths=[1.5, 8.6, 3.6, 3.3])

h(doc, '七、隨堂測驗（10 題，與網頁測驗一致）')
para(doc, '（另附測驗卷，含答案與解析；每題 4 選項，答案分布 A 3 題、B 3 題、C 2 題、D 2 題）',
     10.5, color=GREY)
quiz = [
    ('1. 把食鹽放進水裡攪拌，鹽看不見了，這是什麼現象？', ['蒸發', '溶解', '凝固', '燃燒'], 'B'),
    ('2. 鹽溶解在水裡之後，鹽的質量會怎樣？', ['變少', '不變', '變多', '變成水的質量'], 'B'),
    ('3. 下列哪一種物質放到水裡不會溶解？', ['食鹽', '砂糖', '小石子', '黑糖'], 'C'),
    ('4. 想聞一包粉末的味道，正確的作法是什麼？', ['鼻子湊近深吸一口', '用手輕輕扇，讓氣味飄到鼻子', '打開蓋子湊過去', '用嘴舔一下'], 'B'),
    ('5. 用量的時候如果眼睛是往下俯視，讀出來的數字會怎樣？', ['偏小', '偏大', '剛好', '讀不出來'], 'B'),
    ('6. 同樣多的水，一杯加一平匙糖、一杯加半匙糖，哪一杯比較濃？', ['加一平匙的', '加半匙的', '一樣濃', '看杯子大小'], 'A'),
    ('7. 一杯糖水已經溶到不能再溶了，這叫做什麼？', ['稀溶液', '飽和溶液', '冰水', '混合液'], 'B'),
    ('8. 把一杯飽和的糖水放進冰箱冷藏，會有什麼現象？', ['糖溶解得更多', '溶在水裡的糖變少、會析出結晶', '水變成冰', '什麼都不變'], 'C'),
    ('9. 下列哪一組做法可以讓糖溶解得更快？', ['加冷水、不攪拌', '加熱水、攪拌、敲碎糖', '加很多糖、不動', '把糖放進冰箱'], 'B'),
    ('10. 咖啡裡加了糖，咖啡喝完質量會變少嗎？', ['會少很多', '不會少，是溶在咖啡裡', '會變甜', '看加多少糖'], 'D'),
]
table(doc, [['題號', '題目', '選項', '答案']] +
          [[q[0], '　'.join('(%s)%s' % ('ABCD'[i], o) for i, o in enumerate(q[1])), q[2]] for q in quiz],
      widths=[1.1, 6.4, 7.9, 1.6], size=9.5)

h(doc, '八、常見迷思與澄清')
table(doc, [
    ['迷思', '澄清方式'],
    ['「看不見＝不見了，溶解就是消失。」',
     '用天平秤溶解前後的總質量；把鹽水倒掉，鹽會跟著走——物質只是散開，沒有消失。'],
    ['「溶解就是把東西攪拌均勻而已。」',
     '比較溶解前後的外觀：溶液均一、看不到顆粒；混合則仍看得到不同成分。'],
    ['「攪拌可以把不溶的東西也弄不見。」',
     '細沙、小石子攪拌後仍然沉在杯底，水還是水。'],
    ['「溶得快的東西，溶得也比較多。」',
     '對照「速度」與「總量」兩張圖：速度快≠能溶得多；水少糖多才是濃。'],
    ['「量筒只要看得到刻度就好。」',
     '三種視線比較：俯視偏大、平視正確、仰視偏小。'],
], widths=[6.0, 11.0])

h(doc, '九、評量設計')
for s in ['形成性（每節課）：實作觀察紀錄（學習單①、②）、口頭提問「你怎麼知道的」、'
          '同組互檢與教師巡堂簽章。',
          '總結性（第 4 節）：10 題隨堂測驗＋學習單整體呈現＋教學網頁線上測驗與對戰成績。',
          '診斷性：錯題即時訂正；針對「看不見＝消失」「量筒讀數」「速度與總量」三個迷思，'
          '用追問或小組重做一次來確認是否真的釐清。',
          '評量原則：以「是否說得出理由」為準，不只對答案；鼓勵用圖或文字表達。']:
    para(doc, '・' + s, indent=0.4, space_after=3)

h(doc, '十、差異化設計')
table(doc, [
    ['對象', '作法'],
    ['學習困難學生', '提供「五感觀察卡」與圖示步驟；實驗材料先給老師確認可食用的（鹽、糖）；'
                     '分組時安排能做操作的小組長；量筒讀數改用「有刻度卡」比對。'],
    ['資優或進度快者', '延伸題：比較同一杯水在不同溫度下能溶多少、量筒視線對讀數的影響；'
                     '自製「溶解設計師」任務（設計一杯符合條件的糖水）。'],
    ['語文或表達較弱', '允許用圖記錄、勾選作答；提供句型卡（「因為＿＿，所以＿＿」）。'],
    ['安全與衛生', '不明的粉末一律不聞不嘗；嘗味僅限老師當場示範的食鹽／糖；'
                 '使用熱水由老師操作，學生不碰熱水壺。'],
], widths=[3.4, 13.6])

h(doc, '十一、課後延伸')
for s in ['教學網站（教學專欄）：互動實驗「溶解速度模擬」「量筒讀數」、10 題線上測驗（可語音朗讀）、'
          '兩人搶答對戰、81 張素材圖庫。',
          '家庭任務：找三樣家裡會溶解的東西，記錄溶解前後的變化與快慢。',
          '數位延伸：教學網頁的「溶解與生活」相簿，蒐集家中的溶解例子（拍照＋一句說明）。']:
    para(doc, '・' + s, indent=0.4, space_after=3)

h(doc, '十二、板書規劃')
table(doc, [
    ['節次', '板書重點'],
    ['第 1 節', '觀察五招（摸、看、聞、放大鏡、嘗）｜會溶解：食鹽・糖・黑糖・蜂蜜｜'
                '不會溶解：細沙・小石子・芝麻｜溶解＝均勻分散，物質沒有消失'],
    ['第 2 節', '溶解速度三要素：攪拌↑　水溫↑　顆粒↓　→　溶解變快｜提醒：快 ≠ 多'],
    ['第 3 節', '量筒：眼睛與凹面最低處平視｜濃淡＝溶的量 ÷ 水量｜飽和＝溶到最多｜'
                '溫度↓ → 溶解量↓（析出結晶）'],
    ['第 4 節', '溶解三件事：均勻分散・速度三要素・濃淡與飽和｜生活應用：煮湯・咖啡・洗衣・曬衣服'],
], widths=[2.2, 14.8])

h(doc, '十三、教學省思（待授課後補寫）')
for s in ['1. 學生在「看不見＝消失」的迷思上，需要多少時間與什麼證據才能真正澄清？',
          '2. 三人組的小實驗中，條件控制（水量、溫度、攪拌）最容易失守的是哪一項？',
          '3. 量筒讀數的正確率如何？是否需要增加練習次數或改用「刻度卡」？',
          '4. 線上對戰與測驗的參與率與完成率為何？學生最常錯哪一類題？']:
    para(doc, s, indent=0.4, space_after=4)

path = os.path.join(OUT, '教案_奇妙的溶解.docx')
doc.save(path)
print('已儲存', path)
