# -*- coding: utf-8 -*-
"""把「04奇妙的溶解」素材庫整理成網頁圖庫（media/img＋media/thumb＋images.js）

來源：D:\\opencode\\...\\素材庫\\1_試題\\04奇妙的溶解
檔名格式：04-01-001-用手摸.jpg / 04-02-043-胡椒粉.gif
輸出：media/img/mNNN.(jpg|gif)、media/thumb/mNNN.jpg、media/images.js
"""
import os, json, glob, re
from PIL import Image
from collections import Counter

SRC = r'D:\opencode\115上翰林國小3年級題庫\115上翰林國小3年級題庫\檔案資源\國小自然科學\素材庫\1_試題\04奇妙的溶解'
ROOT = r'C:\Users\rayi\Documents\自然科學教學專欄\115-4_三上自然_奇妙的溶解'
SITE = os.path.join(ROOT, '教學網頁')
IMG = os.path.join(SITE, 'media', 'img')
THUMB = os.path.join(SITE, 'media', 'thumb')
os.makedirs(IMG, exist_ok=True)
os.makedirs(THUMB, exist_ok=True)

files = sorted(glob.glob(os.path.join(SRC, '04-*.jpg')) + glob.glob(os.path.join(SRC, '04-*.gif')))
rows = []
for i, f in enumerate(files, 1):
    base = os.path.splitext(os.path.basename(f))[0]
    m = re.match(r'^(04-\d\d)-(\d{3})-(.+)$', base)
    if not m:
        print('略過', base)
        continue
    group, num, name = m.group(1), m.group(2), m.group(3)
    is_gif = f.lower().endswith('.gif')

    im = Image.open(f)
    if is_gif:
        im.seek(0)
        out_name = f'm{i:03d}.gif'
        im.convert('P', palette=Image.ADAPTIVE).save(os.path.join(IMG, out_name), 'GIF')
    else:
        out_name = f'm{i:03d}.jpg'
        im.convert('RGB').save(os.path.join(IMG, out_name), 'JPEG', quality=88, optimize=True)

    # 縮圖一律 jpg（gif 取第一格）
    still = Image.open(f)
    still.seek(0)
    still = still.convert('RGB')
    still.thumbnail((360, 360))
    still.save(os.path.join(THUMB, f'm{i:03d}.jpg'), 'JPEG', quality=80, optimize=True)

    rows.append({
        'code': num,                      # '004' → media/img/m004.jpg
        'full': f'{group}-{num}',         # '04-01-004'
        'name': name,
        'g': group[-2:],
        'f': 'media/img/' + out_name,
        'h': 'media/thumb/' + f'm{i:03d}.jpg',
    })

print('總數', len(rows), Counter(r['g'] for r in rows))

js = ('/* 素材圖庫清單（自動產生，勿手改） */\nconst IMAGES = '
      + json.dumps(rows, ensure_ascii=False, indent=1) + ';\n')
with open(os.path.join(SITE, 'media', 'images.js'), 'w', encoding='utf-8') as fp:
    fp.write(js)
print('images.js 已寫入', len(js), 'bytes')
