# -*- coding: utf-8 -*-
"""把 skill 页签各英雄的技能 tag 填到 hero 页签 S 列 skillList（zip 直改）。
口径：以 shieldBearer 行（已填）为示例，逗号分隔、skillN 升序；排除 _sub 行。"""
import zipfile, re, shutil, sys

SRC = 'config/config_v1.xlsx'
TMP = 'config/config_v1.tmp.xlsx'
BAK = 'config_v1_backup_20261027_skillList.xlsx'

shutil.copyfile(SRC, BAK)

z = zipfile.ZipFile(SRC)
ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')
sst_head = re.match(r'.*?<sst[^>]*>', ss_xml, re.S).group(0)
raw_si = re.findall(r'<si>.*?</si>', ss_xml, re.S)

def plain(s):
    m = re.match(r'^<t[^>]*>(.*)</t>$', s, re.S)
    return m.group(1) if m else s

existing = [plain(re.sub(r'^<si>|</si>$', '', it)) for it in raw_si]
idx_of = {}
for i, t in enumerate(existing):
    idx_of.setdefault(t, i)
new_items = []

def S(text):
    if text not in idx_of:
        idx_of[text] = len(existing) + len(new_items)
        new_items.append('<si><t>%s</t></si>' % text)
    return idx_of[text]

def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

# ---------- 1. 从 skill 页签(sheet3)收集 tag -> hero 的技能列表 ----------
s3 = z.read('xl/worksheets/sheet3.xml').decode('utf-8')
hero_skills = {}  # heroTag -> [skillTag, ...] 按 A 列出现顺序
for rm in re.finditer(r'<row r="\d+"[^>]*>(.*?)</row>', s3, re.S):
    m = re.search(r'<c r="A\d+" t="s"><v>(\d+)</v></c>', rm.group(1))
    if not m:
        continue
    tag = existing[int(m.group(1))]
    if tag == 'tag':
        continue
    hero = re.match(r'(.+)_skill\d+(_sub)?$', tag)
    assert hero, 'unexpected skill tag: ' + tag
    if not hero.group(2):  # 排除 _sub
        hero_skills.setdefault(hero.group(1), []).append(tag)

# ---------- 2. 改写 sheet1 ----------
s1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
added_cells = 0
filled = []
missing = []

def patch_row(m):
    global added_cells
    r = int(m.group(1))
    body = m.group(0)
    a = re.search(r'<c r="A%d" t="s"><v>(\d+)</v></c>' % r, body)
    if not a or r == 1:  # 表头行
        return body
    tag = existing[int(a.group(1))]
    if tag not in hero_skills:
        missing.append((r, tag))
        return body
    want = ','.join(hero_skills[tag])
    s_cell = re.search(r'<c r="S%d"[^>]*>(?:<v>(\d+)</v>)?</c>' % r, body)
    if s_cell:
        cur = existing[int(s_cell.group(1))] if s_cell.group(1) else ''
        if cur != want:
            print('row %d (%s) already filled: %r != %r, SKIP' % (r, tag, cur, want))
        return body
    idx = S(esc(want))
    filled.append((r, tag, want))
    added_cells += 1
    return body[:-len('</row>')] + '<c r="S%d" t="s"><v>%d</v></c></row>' % (r, idx)

s1_new = re.sub(r'<row r="(\d+)"[^>]*>.*?</row>', patch_row, s1, flags=re.S)
assert not missing, 'hero without skills: %s' % missing
assert len(filled) == 56, 'expected 56 filled, got %d' % len(filled)
assert s1_new != s1

# ---------- 3. 更新 sharedStrings 计数 ----------
m = re.match(r'(.*?<sst[^>]*?count=")(\d+)(" uniqueCount=")(\d+)(">)', ss_xml, re.S)
new_count = int(m.group(2)) + added_cells
new_unique = int(m.group(4)) + len(new_items)
header = m.group(1) + str(new_count) + m.group(3) + str(new_unique) + m.group(5)
new_ss = header + ''.join(raw_si) + ''.join(new_items) + '</sst>'

# ---------- 4. 重写 zip ----------
zout = zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED)
replace = {'xl/sharedStrings.xml': new_ss.encode('utf-8'),
           'xl/worksheets/sheet1.xml': s1_new.encode('utf-8')}
for item in z.infolist():
    data = replace.get(item.filename, None)
    if data is None:
        data = z.read(item.filename)
    zout.writestr(item, data)
zout.close(); z.close()
shutil.move(TMP, SRC)
print('filled %d hero rows, new shared strings %d, sst count +%d' % (len(filled), len(new_items), added_cells))
