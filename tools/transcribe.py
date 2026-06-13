import os
import sys
import argparse
from tqdm import tqdm
from faster_whisper import WhisperModel

def format_time(seconds):
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    milliseconds = int((seconds - int(seconds)) * 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d},{milliseconds:03d}"

def format_timestamp_text(seconds):
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    if hours > 0:
        return f"[{hours:02d}:{minutes:02d}:{secs:02d}]"
    else:
        return f"[{minutes:02d}:{secs:02d}]"

def main():
    parser = argparse.ArgumentParser(description="音影片本地用語音轉文字工具 (faster-whisper)")
    parser.add_argument("file_path", help="要轉寫的音訊或影片檔案路徑")
    parser.add_argument("--model", default="small", choices=["tiny", "base", "small", "medium", "large-v3"], help="Whisper 模型大小 (預設: small)")
    parser.add_argument("--device", default="cpu", choices=["cpu", "cuda", "auto"], help="運行硬體 (預設: cpu)")
    parser.add_argument("--language", default=None, help="語言 (預設: None 自動偵測，亦可手動指定如 zh, en)")
    parser.add_argument("--beam-size", type=int, default=1, help="Beam size (預設: 1，速度較快且穩定；較大數值如 5 可提高部分精準度但較慢且可能在短片中被過濾)")
    parser.add_argument("--temperature", type=float, default=0.0, help="Temperature (預設: 0.0，使用貪婪解碼以保證結果一致且穩定；若設為大於 0 的值，在低品質音訊中可能會隨機回傳空內容)")
    
    args = parser.parse_args()
    
    file_path = args.file_path
    if not os.path.exists(file_path):
        print(f"錯誤：找不到指定的檔案 '{file_path}'")
        sys.exit(1)
        
    model_size = args.model
    device = args.device
    language = args.language
    beam_size = args.beam_size
    temperature = args.temperature
    
    # CPU 預設使用 int8 加速，GPU 預設使用 float16
    compute_type = "int8" if device == "cpu" else "float16"
    if device == "auto":
        compute_type = "default"
        
    print(f"正在載入 Whisper 機器學習模型 '{model_size}' (使用 {device})...")
    try:
        model = WhisperModel(model_size, device=device, compute_type=compute_type)
    except Exception as e:
        print(f"模型載入失敗：{e}")
        print("如果是第一次執行，請確認你的網路能連線至 Hugging Face。")
        sys.exit(1)
        
    print(f"開始轉寫檔案: {os.path.basename(file_path)}")
    print(f"參數偵測: file_path={file_path!r}, beam_size={beam_size!r}, language={language!r}, temperature={temperature!r}, device={device!r}, compute_type={compute_type!r}")
    print("這可能需要幾分鐘，請稍候...")
    
    try:
        segments, info = model.transcribe(file_path, beam_size=beam_size, language=language, temperature=temperature)
    except Exception as e:
        print(f"語音轉寫初始化失敗：{e}")
        print("請確認 ffmpeg 是否已正確安裝並加入環境變數系統路徑中。")
        sys.exit(1)
        
    print(f"偵測到語言: {info.language} (信心度: {info.language_probability:.2f})")
    print(f"音訊總長度: {info.duration:.2f} 秒")
    
    # 建立輸出檔名
    base_name, _ = os.path.splitext(file_path)
    txt_path = f"{base_name}_逐字稿.txt"
    srt_path = f"{base_name}_字幕.srt"
    time_txt_path = f"{base_name}_時間軸逐字稿.txt"
    
    # 立即將 generator 轉換為 list，確保資料完整載入並避免生命週期問題
    print("正在進行語音解碼與辨識...")
    segments_list = list(segments)
    
    # 更新進度條
    with tqdm(total=info.duration, unit="sec", desc="轉寫進度") as pbar:
        last_pos = 0.0
        for segment in segments_list:
            pbar.update(segment.end - last_pos)
            last_pos = segment.end
            
    # 寫入檔案
    # 1. 寫入純文字逐字稿 (TXT)
    with open(txt_path, "w", encoding="utf-8") as f_txt:
        for segment in segments_list:
            f_txt.write(f"{segment.text}\n")
            
    # 2. 寫入時間軸逐字稿 (TXT)
    with open(time_txt_path, "w", encoding="utf-8") as f_time:
        for segment in segments_list:
            time_str = format_timestamp_text(segment.start)
            f_time.write(f"{time_str} {segment.text}\n")
            
    # 3. 寫入字幕檔 (SRT)
    with open(srt_path, "w", encoding="utf-8") as f_srt:
        for i, segment in enumerate(segments_list, start=1):
            start_str = format_time(segment.start)
            end_str = format_time(segment.end)
            f_srt.write(f"{i}\n")
            f_srt.write(f"{start_str} --> {end_str}\n")
            f_srt.write(f"{segment.text.strip()}\n\n")
            
    print("\n轉寫完成！已成功產出以下檔案：")
    print(f"1. 純文字逐字稿：{os.path.basename(txt_path)}")
    print(f"2. 時間軸逐字稿：{os.path.basename(time_txt_path)}")
    print(f"3. SRT 字幕檔：{os.path.basename(srt_path)}")

if __name__ == "__main__":
    main()
