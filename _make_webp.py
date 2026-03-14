from PIL import Image
import os
src_dir = r'd:\\webProjects\\opaGroup\\assets'
out_dir = r'd:\\webProjects\\opaGroup\\assets'
images = {
    'hero.webp': ('page5_img2.jpg', 1800),
    'distribution.webp': ('page2_img1.jpg', 1400),
    'logistics.webp': ('page4_img2.jpg', 1400),
    'sourcing.webp': ('page6_img2.jpg', 1400),
    'network.webp': ('page8_img2.jpg', 1400),
}
for out_name, (src_name, max_w) in images.items():
    src_path = os.path.join(src_dir, src_name)
    out_path = os.path.join(out_dir, out_name)
    with Image.open(src_path) as im:
        im = im.convert('RGB')
        w, h = im.size
        if w > max_w:
            new_h = int(h * (max_w / w))
            im = im.resize((max_w, new_h), Image.LANCZOS)
        im.save(out_path, 'WEBP', quality=82, method=6)
print('done')
