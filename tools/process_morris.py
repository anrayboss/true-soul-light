from PIL import Image
import os

src = r"C:\Users\anray\.gemini\antigravity-ide\brain\c842dec3-a88a-4def-aad7-8403fe8550d1\morris_chang_looking_left_1782277710792.png"
dst = r"d:\Git\true-soul-light\tools\portraits\chang.png"

if not os.path.exists(src):
    print(f"Source file not found: {src}")
    exit(1)

img = Image.open(src).convert("RGBA")
pixels = img.load()
w, h = img.size

for x in range(w):
    for y in range(h):
        r, g, b, a = pixels[x, y]
        L = (r + g + b) // 3
        if L > 225:
            alpha = 0
        elif L < 100:
            alpha = 255
        else:
            alpha = int(255 * (225 - L) / (225 - 100))
        pixels[x, y] = (r, g, b, alpha)

os.makedirs(os.path.dirname(dst), exist_ok=True)
img.save(dst, "PNG")
print("Successfully processed and saved Morris Chang portrait with transparency!")
