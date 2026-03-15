from PIL import Image
import os
src_dir = r'd:\\webProjects\\opaGroup\\assets'
imgs = {
    'extended.webp': ('page6_img3.jpg', 1400),
    'products.webp': ('page10_img1.jpg', 1400),
    'logistics2.webp': ('page9_img1.jpg', 1400)
}
for out_name, (src_name, max_w) in imgs.items():
    src_path = os.path.join(src_dir, src_name)
    out_path = os.path.join(src_dir, out_name)
    with Image.open(src_path) as im:
        im = im.convert('RGB')
        w, h = im.size
        if w > max_w:
            new_h = int(h * (max_w / w))
            im = im.resize((max_w, new_h), Image.LANCZOS)
        im.save(out_path, 'WEBP', quality=82, method=6)
print('done')
