import os
from PIL import Image

portrait_dir = r"d:\Git\true-soul-light\tools\portraits"
for filename in os.listdir(portrait_dir):
    if filename.endswith(".png"):
        path = os.path.join(portrait_dir, filename)
        img = Image.open(path).convert("RGBA")
        pixels = img.load()
        w, h = img.size
        
        for x in range(w):
            for y in range(h):
                r, g, b, a = pixels[x, y]
                # Calculate brightness
                L = (r + g + b) // 3
                
                # Luminance to Alpha with contrast enhancement
                if L > 225:
                    alpha = 0
                elif L < 100:
                    alpha = 255
                else:
                    # Linearly map 100-225 to 255-0
                    alpha = int(255 * (225 - L) / (225 - 100))
                    
                # To keep the drawings pure, we can make the color black or keep original dark tones
                pixels[x, y] = (r, g, b, alpha)
                
        img.save(path, "PNG")
        print(f"Processed transparency for {filename}")
