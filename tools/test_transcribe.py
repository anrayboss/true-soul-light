from faster_whisper import WhisperModel
import os

file_path = r"c:\Users\anray\Desktop\陳仲豪合作\心智圖\30秒測試影片.mp4"
model_size = "base"
device = "cpu"
compute_type = "int8"
beam_size = 1
language = None

print(f"Testing local execution...")
print(f"File exists: {os.path.exists(file_path)}")

model = WhisperModel(model_size, device=device, compute_type=compute_type)
segments, info = model.transcribe(file_path, beam_size=beam_size, language=language)

print(f"Detected language: {info.language}")
segments_list = list(segments)
print(f"Segments found: {len(segments_list)}")
for s in segments_list:
    print(f"[{s.start:.2f} -> {s.end:.2f}]: {s.text}")
