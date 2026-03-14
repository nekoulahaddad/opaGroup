import textwrap
p=r'd:\\webProjects\\opaGroup\\_pdf_text_clean.txt'
t=open(p,encoding='utf-8').read()
pages=t.split('--- Page ')
print('pages', len(pages)-1)
for p in pages[1:]:
    num=p.split(' ---',1)[0].strip()
    body=p.split(' ---',1)[1] if ' ---' in p else p
    print('\n===== Page', num, '=====\n')
    print(body[:3000])
