from PIL import Image
import os
root = r'd:\\webProjects\\opaGroup\\assets'
rows = []
for name in os.listdir(root):
    path = os.path.join(root, name)
    if not os.path.isfile(path):
        continue
    try:
        with Image.open(path) as im:
            w,h = im.size
        size = os.path.getsize(path)
        rows.append((w*h, w, h, size, name))
    except Exception:
        continue
rows.sort(reverse=True)
for r in rows[:20]:
    print(f"{r[4]} {r[1]}x{r[2]} {r[3]//1024}KB")
