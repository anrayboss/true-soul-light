import os
import shutil

src_dir = r"C:\Users\anray\AppData\Local\Google\Chrome\User Data" # wait, the path is C:\Users\anray\.gemini\antigravity-ide\brain\c842dec3-a88a-4def-aad7-8403fe8550d1
src_dir = r"C:\Users\anray\.gemini\antigravity-ide\brain\c842dec3-a88a-4def-aad7-8403fe8550d1"
dest_dir = r"d:\Git\true-soul-light\tools\portraits"

if not os.path.exists(dest_dir):
    os.makedirs(dest_dir)

files = os.listdir(src_dir)

mapping = {
    "maslow_portrait": "maslow.png",
    "inamori_portrait": "inamori.png",
    "porter_portrait": "porter.png",
    "suntzu_portrait": "sun_tzu.png",
    "trout_portrait": "trout.png",
    "guiguzi_portrait": "guiguzi.png",
    "kahneman_portrait": "kahneman.png",
    "confucius_portrait": "confucius.png",
    "cialdini_portrait": "cialdini.png",
    "berger_portrait": "berger.png",
    "ogilvy_portrait": "ogilvy.png",
    "guanzhong_portrait": "guan_zhong.png",
    "ford_portrait": "ford.png",
    "wang_portrait": "wang.png",
    "wattles_portrait": "wattles.png",
    "fanli_portrait": "fan_li.png",
    "musk_portrait": "musk.png",
    "chang_portrait": "chang.png"
}

for f in files:
    for prefix, target_name in mapping.items():
        if f.startswith(prefix) and f.endswith(".png"):
            src_path = os.path.join(src_dir, f)
            dest_path = os.path.join(dest_dir, target_name)
            shutil.copy2(src_path, dest_path)
            print(f"Copied {f} to {target_name}")

# Fallback for chang if not copied (since quota ran out)
chang_path = os.path.join(dest_dir, "chang.png")
if not os.path.exists(chang_path):
    # Use wang as temporary placeholder for Morris Chang so it doesn't crash or look blank
    wang_path = os.path.join(dest_dir, "wang.png")
    if os.path.exists(wang_path):
        shutil.copy2(wang_path, chang_path)
        print("Copied wang as placeholder for chang")
