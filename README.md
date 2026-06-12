# PPT 自動配音轉影片工具

將 PowerPoint 簡報自動轉換為 MP4 影片。
每頁投影片依照設定規則組合講稿，透過 Edge-TTS 生成 AI 語音，最後合成為完整影片。

---

## 環境需求

- Windows 作業系統（截圖功能需要 PowerPoint COM 介面）
- Python 3.9 以上
- Microsoft PowerPoint（已安裝）
- ffmpeg（需加入系統 PATH）

### 安裝 ffmpeg

1. 至 [https://ffmpeg.org/download.html](https://ffmpeg.org/download.html) 下載 Windows 版本
2. 解壓縮後將 `bin/` 資料夾路徑加入系統環境變數 `PATH`
3. 開啟命令提示字元輸入 `ffmpeg -version` 確認安裝成功

### 安裝 Python 套件

```bash
pip install -r requirements.txt
```

---

## 檔案說明

```
├── phase1_generate.py   # Phase 1：解析 PPT → TTS 配音 → 合成單頁 clip
├── phase2_compose.py    # Phase 2：掃描 clips → 完整性檢查 → 串接最終 MP4
├── settings.yaml        # 設定檔（副標題開場白、收尾、語音等）
├── requirements.txt     # Python 套件需求
└── README.md            # 本文件
```

---

## 快速開始

### Step 1：設定 settings.yaml

編輯 `settings.yaml`，設定各副標題的開場白與收尾，以及前導頁範圍：

```yaml
notes_slides: "1-3"   # 第 1~3 頁為前導頁，改唸備註欄內容
```

### Step 2：執行 Phase 1（生成 clips）

```bash
# 全部頁面
python phase1_generate.py --pptx your_file.pptx

# 只處理前導頁（第 1~3 頁）
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

## settings.yaml 說明

### 副標題設定

每個副標題可設定以下四個欄位：

```yaml
subtitle_config:
  "說明":
    opener: "以下說明此原則的內容。"   # 開場白
    closer: "以上為本原則的說明。"      # 收尾語
    read_shapes: false                  # 是否唸圖形內文字（橢圓圖說文字）
    read_table: true                    # 是否唸表格
```

目前支援的五種副標題：

| 副標題 | read_shapes | read_table |
|--------|-------------|------------|
| 說明 | ✗ | ✓ |
| 設定方法 | ✗ | ✗ |
| Azure 入口網站操作示意圖 | ✓ | ✗ |
| Azure CLI設定方法 | ✗ | ✗ |
| PowerShell設定方法 | ✗ | ✗ |

### 前導頁設定

```yaml
# 三種寫法任選一種
notes_slides: "1-3"        # 第 1 到第 3 頁改唸備註欄
notes_slides: [1, 2, 5]    # 第 1、2、5 頁改唸備註欄
notes_slides:              # 空白 = 全部不唸備註欄
```

### 其他設定

```yaml
table_prefix: "此原則支援以下設定方法："   # 表格唸法前綴
continuation_opener: "接續上頁，"           # 同副標題第二頁的開場白
tts_voice: "zh-TW-HsiaoChenNeural"         # TTS 語音（台灣女聲）
tts_batch_size: 20                          # TTS 每批次請求數
video_fps: 12                               # 影片幀率
```

### 更換語音

Edge-TTS 支援多種繁體中文語音：

| 語音名稱 | 性別 | 風格 |
|---------|------|------|
| `zh-TW-HsiaoChenNeural` | 女 | 自然（預設） |
| `zh-TW-HsiaoYuNeural` | 女 | 活潑 |
| `zh-TW-YunJheNeural` | 男 | 自然 |

---

## 符號清理規則

講稿組合後會自動清理以下符號，避免 TTS 唸出奇怪的內容：

| 原始符號 | 處理方式 |
|---------|---------|
| `【` `】` | 移除 |
| `/` | 替換為 `，` |
| 多餘空白 | 壓縮為單一空格 |

---

## 斷點續跑

Phase 1 已支援斷點續跑，若執行中途失敗：

- 已合成的 clip 不會重新生成
- 重新執行相同指令即可從中斷處繼續
- 若需強制重新生成某頁，手動刪除對應的 `page_XXXX.mp4` 再執行即可

---

## 常見問題

**Q：執行時出現 PowerPoint 視窗**
A：截圖過程中 PowerPoint 會短暫開啟，屬正常現象，請勿手動關閉。

**Q：TTS 生成失敗或音質異常**
A：Edge-TTS 為爬取微軟 Edge 瀏覽器的語音服務，偶發網路問題屬正常現象。
程式已支援斷點續跑，重新執行即可補生成失敗的頁面。

**Q：某頁投影片沒有聲音**
A：該頁備註欄與投影片文字皆為空，程式會插入約 1 秒靜音。
建議補充備註欄內容或在 `settings.yaml` 的 `notes_slides` 中指定該頁。

**Q：Phase 2 提示缺漏 clip**
A：執行 `phase1_generate.py --pages <缺漏頁碼>` 補生成後再執行 Phase 2。
