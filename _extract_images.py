from pypdf import PdfReader
import os
pdf = r"C:\Users\Nekoula\Desktop\OPA-GROUP-Company-Profile-.pdf"
out_dir = r"d:\webProjects\opaGroup\assets"
os.makedirs(out_dir, exist_ok=True)
reader = PdfReader(pdf)
count = 0
for i, page in enumerate(reader.pages):
    for j, img in enumerate(page.images):
        count += 1
        ext = "bin"
        if img.name and "." in img.name:
            ext = img.name.split(".")[-1]
        filename = f"page{i+1}_img{j+1}.{ext}"
        path = os.path.join(out_dir, filename)
        with open(path, "wb") as f:
            f.write(img.data)
print("images", count)
