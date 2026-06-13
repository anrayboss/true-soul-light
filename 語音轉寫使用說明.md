# 音影片自動轉寫與大綱生成技能 使用說明書

本專案已建置本地端高效率音影片轉寫（Whisper）與自動大綱生成技能。此說明書將引導你如何在目前專案或未來的新專案中調用與使用此技能。

---

## 📂 已建置之工具與檔案

- **語音轉寫主腳本**：[tools/transcribe.py](file:///c:/Users/anray/Desktop/%E9%99%B3%E4%BB%B2%E8%B1%AA%E5%90%88%E4%BD%9C/%E5%BF%83%E6%99%BA%E5%9C%96/tools/transcribe.py)
- **Agent 技能說明檔**：[transcribe_skill.md](file:///c:/Users/anray/Desktop/%E9%99%B3%E4%BB%B2%E8%B1%AA%E5%90%88%E4%BD%9C/%E5%BF%83%E6%99%BA%E5%9C%96/transcribe_skill.md)

---

## 🚀 在本專案中的使用方式

### 方法一：讓 Agent 自動執行（最推薦）
直接在對話中對 Agent 輸入：
> 「幫我轉寫這個音檔/影片，並生成字幕和大綱：`C:\路徑\到\你的\檔案.mp3`」

**Agent 會自動執行以下步驟：**
1. 呼叫 `tools/transcribe.py` 完成轉寫。
2. 產出 `*_逐字稿.txt`、`*_時間軸逐字稿.txt` 及 `*_字幕.srt`。
3. 讀取逐字稿內容，並使用 LLM 智能整理出 `*_大綱.md` 檔案。

### 方法二：手動使用終端機執行
你也可以在 PowerShell 終端機中手動執行轉寫腳本：
```powershell
python tools/transcribe.py "<你的音影片檔案路徑>" --model base
```

**參數設定選項：**
- `--model`：預設為 `base`（約 140MB 記憶體佔用）。
  - 可選：`tiny` (最快)、`base` (推薦)、`small` (中文準確度更好，約 460MB)、`medium` 或 `large-v3` (最精準)。
- `--device`：預設為 `cpu`。
  - **GPU 加速**：若你的電腦有 NVIDIA 顯示卡且安裝了 CUDA，請加上 `--device cuda`，轉寫速度可提升數倍。
- `--language`：預設為 `zh` (中文)。如果轉寫英文音檔，請指定 `--language en`。

---

## 🆕 開啟新專案時，如何套用此技能？

因為 `ffmpeg` 與 Python 依賴環境已經全域安裝在你的電腦上，在新專案中**不需要重新安裝環境**，只需讓新專案的 Agent 知道去哪裡讀取技能即可。

有以下兩種方式可以套用：

### 方法一：複製技能檔（推薦，保持專案整潔）
1. 將現有的 [transcribe_skill.md](file:///c:/Users/anray/Desktop/%E9%99%B3%E4%BB%B2%E8%B1%AA%E5%90%88%E4%BD%9C/%E5%BF%83%E6%99%BA%E5%9C%96/transcribe_skill.md) 檔案複製到你**新專案的根目錄**下。
2. 因為技能檔內部的 Python 腳本路徑為絕對路徑（指向此專案的 `tools/transcribe.py`），新專案的 Agent 讀取該技能檔後，即可**直接執行**轉寫，不需複製程式碼。

### 方法二：直接在對話中指引絕對路徑（最快速）
在新專案的對話中，直接對 Agent 指示：
> 「請讀取並載入位於 `c:\Users\anray\Desktop\陳仲豪合作\心智圖\transcribe_skill.md` 的技能，幫我將這個檔案轉寫成逐字稿與大綱：`<新專案音檔路徑>`」

Agent 就會直接讀取該技能檔，並調用背後的轉寫工具與大綱生成邏輯。
