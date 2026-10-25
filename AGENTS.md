# 奇妙的溶解（專案藍圖）

> 本檔為跨 Agent 通用的專案藍圖（AGENTS.md 開放標準）。任何 Agent 的每個 session 都應先讀本檔＋`handoff.md`。

## 專案簡介
翰林三上 第4單元「奇妙的溶解」的完整教學專案：教案、教學簡報、學習單、測驗卷、
互動教學網頁（課程地圖＋互動實驗＋簡報播放器＋10 題測驗＋語音朗讀＋參考影片＋抽籤計時器＋素材圖庫）、
兩人即時配對搶答遊戲，並部署到 GitHub Pages（＋ Render 跑對戰後端）。

## 關鍵時程
- 第 1 節：（日期）　主題
- 第 2 節：（日期）　主題
- 第 3 節：（日期）　主題
- 第 4 節：（日期）　主題

## 目標與路線圖
- [x] 階段 0：專案初始化（藍圖、交接檔、repo、Obsidian 筆記）
- [x] 階段 1：確認教學範圍與節數（INc-Ⅱ-2／INd-Ⅱ-2／INe-Ⅱ-2）
- [x] 階段 2：內容與素材（81 張圖庫、5 部影片、迷思清單）
- [x] 階段 3：教案 docx＋學習單＋測驗卷（docx／PDF）
- [x] 階段 4：教學簡報 pptx＋圖檔（PowerPoint COM 匯出 JPG）
- [x] 階段 5：教學網頁與對戰（battle.html／server.js／30 題題庫）
- [x] 階段 6：驗收（`verify_site.js` 36/36、瀏覽器 0 console error）
- [x] 階段 7：上線 GitHub Pages（workflow 已準備）
- [ ] 階段 8：對戰後端上 Render（藍圖已準備，待套用）
- [ ] 階段 9：（選）旁白動畫 MP4

## 資料夾結構
```
115-4_三上自然_奇妙的溶解/                  ← 獨立 git repo（公開）
├── AGENTS.md                     專案藍圖（本檔）
├── handoff.md                    交接檔
├── .github/workflows/pages.yml   GitHub Pages 部署
├── render.yaml                   Render 藍圖（對戰後端）
├── battle-server.js              Render 啟動入口
├── 工具/                         產生程式（build_*.py）
├── 教案/                         教案、學習單、測驗卷
├── 簡報/                         教學簡報 pptx
├── 影片/                         影片素材
└── 教學網頁/
    ├── index.html                主網頁（單檔、file:// 可開）
    ├── battle.html               對戰頁
    ├── server.js                 對戰伺服器（區網／Render 共用）
    ├── questions.json            30 題題庫（部署時排除）
    ├── battle-api.js             雲端後端位址（空字串＝區網模式）
    ├── start-battle.bat          老師一鍵啟動區網對戰
    ├── slides/slideNN.jpg        簡報圖（1600×900）
    ├── media/img/ media/thumb/   素材圖庫
    └── downloads/                可下載教材
```

## 重建方式
```powershell
python 工具\build_plan.py       # 教案
python 工具\build_sheets.py     # 學習單＋測驗卷
python 工具\build_deck.py       # 簡報 pptx
python 工具\build_questions.py  # 對戰題庫
python 工具\build_images.py     # 素材圖庫

# 簡報 → JPG（必須用 PowerPoint COM）
powershell -File <skill>\scripts\pptx_to_jpg.ps1 -Pptx "簡報\X.pptx" -OutDir "教學網頁\slides"
# docx → PDF（必須用 Word COM）
powershell -File <skill>\scripts\docx_to_pdf.ps1 -Docx "教案\X.docx"
```
- **絕對不要用 LibreOffice 轉中文 PDF**：字型子集損壞，整頁變成重複的單元名二字。
- **python-pptx 要補 `a:ea` 字型**，否則中文掉字型。
- 驗收：`node <skill>\scripts\verify_site.js "115-4_三上自然_奇妙的溶解"`

## 同步層級
| 層級 | 平台 | 位置 | 讀取時機 |
|------|------|------|---------|
| L1 | 本地 | `AGENTS.md`＋`handoff.md` | 每個 session |
| L2 | GitHub | repo（公開） | 指定時 |
| L3 | Obsidian | `奇妙的溶解/專案工作流程.md` | 有需要時 |

## 已上線網址
- 教學網頁：https://rayilin.github.io/dissolution-fun/
- 對戰頁：https://rayilin.github.io/dissolution-fun/battle.html
- 雲端對戰後端：https://dissolution-battle.onrender.com（待 Render 建立服務）

## 工作約定
- **開工先讀 `handoff.md`，收工必更新 `handoff.md`**
- 修改共用檔案前先讀最新內容，避免覆蓋其他 Agent 的變更
- 所有回應與文件使用繁體中文
- 修改前先確認計畫，優先保留原有資料結構
- 不刪使用者既有資料夾、舊版教材、素材原圖

## 安全與隱私（不可違反）
- **不把 API key、密碼、憑證寫進 repo**；一律放 `.env` 並列入 `.gitignore`
- **學生資料只用座號**，不出現姓名、學號、班級以外的個資、照片或聯絡方式
- `questions.json`（對戰答案）**不得**部署到公開網站（Pages 與 Render 都要 404）
- 要公開分享前，先確認檔案裡沒有上述內容
