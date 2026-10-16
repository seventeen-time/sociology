# -*- coding: utf-8 -*-
"""从 docx 提取纯文本（只用标准库 zipfile/re）"""
import zipfile, re, sys, os, io

sys.stdout.reconfigure(encoding='utf-8')

SRC = r'C:\Users\a1772\Desktop\哲学的贫困'
OUT = r'C:\Users\a1772\Desktop\哲学的贫困\_记录'

files = [
    '第一章 科学的发现.docx',
    '第二章 政治经济学的形而上学 正文1.docx',
    '第二章 政治经济学的形而上学 正文2.docx',
    '第二章 政治经济学的形而上学 正文3.docx',
]

os.makedirs(OUT, exist_ok=True)

def unescape(t):
    t = t.replace('&amp;', '&')
    t = t.replace('&lt;', '<')
    t = t.replace('&gt;', '>')
    t = t.replace('&apos;', "'")
    t = t.replace('&quot;', '"')
    return t

all_lines = []
for fn in files:
    p = os.path.join(SRC, fn)
    if not os.path.exists(p):
        print('MISSING', fn)
        continue
    z = zipfile.ZipFile(p)
    x = z.read('word/document.xml').decode('utf-8')
    paras = re.findall(r'<w:p[ >].*?</w:p>', x, re.S)
    out = []
    for para in paras:
        t = ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', para, re.S))
        t = unescape(t).strip()
        if t:
            out.append(t)
    print('=== %s === paras=%d chars=%d' % (fn, len(out), sum(len(s) for s in out)))
    all_lines.append('\n<<<FILE: %s>>>\n' % fn)
    all_lines.extend(out)

with io.open(os.path.join(OUT, '源文本_提取自4个docx.txt'), 'w', encoding='utf-8') as f:
    f.write('\n'.join(all_lines))

total = sum(len(s) for s in all_lines)
print('TOTAL lines=%d chars=%d' % (len(all_lines), total))
