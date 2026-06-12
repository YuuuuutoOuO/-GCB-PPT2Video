# PPT 自動配音轉影片工具

將 PowerPoint 簡報自動轉換為 MP4 影片。
每頁投影片依照設定規則組合講稿，透過 Edge-TTS 生成 AI 語音，最後合成為完整影片。

---

## 環境需求

- Windows 作業系統（截圖功能需要 PowerPoint COM 介面）
- Python 3.9 以上
- Microsoft PowerPoint（已安裝）
- ffmpeg（需在 settings.yaml 中設定路徑，或加入系統 PATH）

### 安裝 ffmpeg

1. 至 [https://ffmpeg.org/download.html](https://ffmpeg.org/download.html) 下載 Windows 版本
2. 解壓縮後記下 `ffmpeg.exe` 的完整路徑（例如 `C:\ffmpeg\bin\ffmpeg.exe`）
3. 填入 `settings.yaml` 的 `ffmpeg_path`

### 安裝 Python 套件

```bash
pip install -r requirements.txt
```

> ⚠️ **注意**：執行程式前請以**系統管理員身份**開啟命令提示字元，
> DPI 設定需要寫入 Registry 才能提升截圖畫質。

---

## 專案結構

```
├── phase1_generate.py   # Phase 1 主程式入口
├── phase2_compose.py    # Phase 2 主程式入口
├── settings.yaml        # 設定檔（副標題、語音、畫質等）
├── requirements.txt     # Python 套件需求
├── README.md            # 本文件
└── core/                # 核心模組
    ├── __init__.py
    ├── parser.py        # PPT 結構解析
    ├── script.py        # 講稿組合（含 SSML 停頓）
    ├── tts.py           # TTS 配音（Edge-TTS）
    ├── video.py         # 影片合成（截圖、clip、ffmpeg）
    └── utils.py         # 共用工具（文字清理、範圍解析、DPI 設定）
```

---

## 快速開始

### Step 1：設定 settings.yaml

至少確認以下三項：

```yaml
# ffmpeg 路徑
ffmpeg_path: "C:\\ffmpeg\\bin\\ffmpeg.exe"

# 前導頁範圍（唸備註欄，其餘頁唸投影片內容）
notes_slides: [1]

# 截圖畫質（DPI）
export_dpi: 200
```

### Step 2：執行 Phase 1（生成 clips）

```bash
# 以系統管理員身份開啟命令提示字元，再執行：

# 全部頁面
python phase1_generate.py --pptx your_file.pptx

# 只處理前導頁（頁碼指定）
python phase1_generate.py --pptx your_file.pptx --pages 1-3

# 只處理特定標題編號
python phase1_generate.py --pptx your_file.pptx --range 58,60-65

# 同時指定頁碼與標題編號（會顯示警告並要求確認）
python phase1_generate.py --pptx your_file.pptx --pages 1-3 --range 58-60
```

執行完成後，確認 `clips/` 資料夾結構與 `clips/scripts.txt` 講稿內容是否正確。

### Step 3：執行 Phase 2（合成最終影片）

```bash
# 基本用法
python phase2_compose.py

# 指定 clips 資料夾與輸出檔名
python phase2_compose.py --clips clips --output final.mp4

# 合成完後自動刪除 clips 資料夾（節省空間）
python phase2_compose.py --output final.mp4 --clean
```

---

## 參數說明

### phase1_generate.py

| 參數 | 預設值 | 說明 |
|------|--------|------|
| `--pptx` | （必填） | PPT 檔案路徑 |
| `--settings` | `settings.yaml` | 設定檔路徑 |
| `--clips` | `clips` | clips 輸出資料夾 |
| `--pages` | （空，全部） | 指定頁碼，如 `1-3,5,10-15` |
| `--range` | （空，全部） | 指定標題編號，如 `58,60-65` |

**`--pages` 與 `--range` 同時使用時**，程式會顯示警告並列出各自涵蓋的頁面，
輸入 `y` 確認後才繼續執行，結果取兩者的聯集。

### phase2_compose.py

| 參數 | 預設值 | 說明 |
|------|--------|------|
| `--clips` | `clips` | clips 資料夾路徑 |
| `--output` | `final_output.mp4` | 輸出影片檔名 |
| `--settings` | `settings.yaml` | 設定檔路徑 |
| `--clean` | `false` | 合成完後刪除 clips 資料夾 |

---

## clips 資料夾結構

```
clips/
├── scripts.txt          # 本次處理頁的講稿預覽
├── intro/               # 無標題編號的前導頁
│   ├── page_0001.mp4
│   └── page_0002.mp4
├── 001/                 # 標題編號第 1 條的所有頁
│   ├── page_0004.mp4
│   └── page_0005.mp4
└── 058/                 # 標題編號第 58 條的所有頁
    ├── page_0268.mp4
    └── page_0270.mp4
```

最終影片串接順序：`intro/` → `001/` → `002/` → ... 按編號升序。

---

## settings.yaml 完整說明

### ffmpeg 路徑

```yaml
ffmpeg_path: "C:\\ffmpeg\\bin\\ffmpeg.exe"
```

### 前導頁設定

```yaml
# 三種寫法任選一種
notes_slides: [1]          # 只有第 1 頁唸備註欄
notes_slides: "1-5"        # 第 1 到第 5 頁唸備註欄
notes_slides: [1, 2, 5]    # 第 1、2、5 頁唸備註欄
notes_slides:              # 空白 = 全部不唸備註欄
```

### 五種副標題設定

每個副標題可設定以下四個欄位：

```yaml
subtitle_config:
  "說明":
    opener: "以下說明此原則的內容。"   # 開場白
    closer: "以上為本原則的說明。"      # 收尾語
    read_shapes: false                  # 是否唸圖形內文字
    read_table: true                    # 是否唸表格
```

| 副標題 | read_shapes | read_table |
|--------|-------------|------------|
| 說明 | ✗ | ✓ |
| 設定方法 | ✗ | ✗ |
| Azure 入口網站操作示意圖 | ✓ | ✗ |
| Azure CLI設定方法 | ✗ | ✗ |
| PowerShell設定方法 | ✗ | ✗ |

### 語音與品質設定

```yaml
# 段落間停頓時間（毫秒），每個 PPT 段落之間自動插入停頓
pause_between_paragraphs_ms: 300

# TTS 語音（台灣繁體中文）
tts_voice: "zh-TW-HsiaoChenNeural"
tts_batch_size: 20        # 每批次 TTS 請求數，過高可能被限流

# 截圖畫質 DPI（預設 96，建議 150~300）
export_dpi: 200

# 影片幀率（靜態投影片 12fps 即可）
video_fps: 12
```

### 可用語音清單

| 語音名稱 | 性別 | 風格 |
|---------|------|------|
| `zh-TW-HsiaoChenNeural` | 女 | 自然（預設） |
| `zh-TW-HsiaoYuNeural` | 女 | 活潑 |
| `zh-TW-YunJheNeural` | 男 | 自然 |

---

## 核心模組說明

| 模組 | 功能 |
|------|------|
| `core/parser.py` | 解析投影片的標題、副標題、段落、圖形文字、表格、備註，程式碼行（`#` 開頭）自動跳過不唸 |
| `core/script.py` | 依 settings 規則組合講稿，輸出 SSML 格式（含 `<break>` 停頓標籤） |
| `core/tts.py` | 呼叫 Edge-TTS 生成音檔，自動重試 3 次並驗證檔案大小 |
| `core/video.py` | PowerPoint COM 截圖、單頁 MP4 合成、ffmpeg 串接 |
| `core/utils.py` | 文字清理、`--pages`/`--range` 範圍解析、Registry DPI 寫入 |

---

## 符號清理規則

講稿組合後會自動清理以下符號：

| 原始符號 | 處理方式 |
|---------|---------|
| `【` `】` | 移除 |
| `/` | 替換為 `，` |
| 換行符 `\n` | 替換為 `，` |
| 連續標點 | 合併為單一 `，` |
| 多餘空白 | 壓縮為單一空格 |
| `#` 開頭的行 | 跳過不唸（程式碼行）|

---

## 斷點續跑

Phase 1 支援斷點續跑：
- 已合成的 clip（`page_XXXX.mp4`）不會重新生成
- 重新執行相同指令即可從中斷處繼續
- 若需強制重新生成某頁，手動刪除對應的 `page_XXXX.mp4` 再執行

---

## 常見問題

**Q：執行時出現 PowerPoint 視窗**
A：截圖過程中 PowerPoint 會短暫開啟，屬正常現象，請勿手動關閉。

**Q：DPI 設定沒有效果**
A：需以系統管理員身份執行命令提示字元，Registry 才有寫入權限。

**Q：TTS 生成失敗或音質異常**
A：Edge-TTS 為爬取微軟 Edge 瀏覽器的語音服務，偶發網路問題屬正常現象。
程式支援斷點續跑，重新執行即可補生成失敗的頁面。

**Q：某頁投影片沒有聲音**
A：該頁備註欄與投影片文字皆為空，程式會插入約 1 秒靜音。
建議補充備註欄內容，或在 `settings.yaml` 的 `notes_slides` 中指定該頁手動撰寫備註。

**Q：Phase 2 提示缺漏 clip**
A：執行 `phase1_generate.py --pages <缺漏頁碼>` 補生成後再執行 Phase 2。

**Q：第1條的說明、設定方法頁沒有聲音**
A：確認 `settings.yaml` 的 `notes_slides` 設定，確保這些頁不在前導頁範圍內。
例如若只有封面是前導頁，應設定為 `notes_slides: [1]`。