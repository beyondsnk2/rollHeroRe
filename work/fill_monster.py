# -*- coding: utf-8 -*-
"""向 config_v1.xlsx 写入 pic.xlsx/monster 页签 4 张怪卡的识别结果：
- monster 页签（sheet7.xml）重写 2-5 行：wolf/blackWolf/bandit/banditArcher
  class=monster, tier=1, baseHp=4/5/6/6, baseShield=0, skillList=<tag>_skill1
- skill 页签（sheet2.xml）追加 114-117 行：
  技能名=撕咬/撕咬/劈砍/射击，消耗=null，生效=自动，K=enemy
zip 直改 sharedStrings + sheet2 + sheet7，其余条目（含 drawing）原样透传。
"""
import zipfile, re, shutil, os

SRC = 'config/config_v1.xlsx'
TMP = 'config/config_v1.tmp.xlsx'
BACKUP = 'config_v1_backup_20261026.xlsx'

# ---------- sharedStrings 复用/新增 ----------
z = zipfile.ZipFile(SRC)
ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')
old_items = re.findall(r'<si>.*?</si>', ss_xml, re.S)
def plain(s):
    m = re.match(r'^<t[^>]*>(.*)</t>$', s, re.S)
    return m.group(1) if m else s
idx_of = {}
for i, it in enumerate(old_items):
    idx_of.setdefault(plain(it), i)
new_items = []
def S(text):
    if text not in idx_of:
        idx_of[text] = len(old_items) + len(new_items)
        new_items.append('<si><t>%s</t></si>' % text)
    return idx_of[text]
def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
def Sesc(text):
    return S(esc(text))

def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def build_row(r, values, spans):
    cells = []
    for ci, v in enumerate(values, start=1):
        if v[0] is None:
            continue
        ref = '%s%d' % (col_letter(ci), r)
        if v[0] == 's':
            cells.append('<c r="%s" t="s"><v>%d</v></c>' % (ref, Sesc(v[1])))
        else:
            cells.append('<c r="%s"><v>%s</v></c>' % (ref, v[1]))
    return '<row r="%d" spans="%s">%s</row>' % (r, spans, ''.join(cells))

NULL = 'null'

# ---------- monster 数据（A tag, B 名, C class, D tier, E hp, F shield, G skillList）----------
MONSTERS = [
    ('wolf',         '狼',       4),
    ('blackWolf',    '黑狼',     5),
    ('bandit',       '山贼',     6),
    ('banditArcher', '山贼弓手', 6),
]
monster_rows = []
for i, (tag, name, hp) in enumerate(MONSTERS):
    r = 2 + i
    vals = [('s', tag), ('s', name), ('s', 'monster'), ('n', 1),
            ('n', hp), ('n', 0), ('s', '%s_skill1' % tag)]
    monster_rows.append(build_row(r, vals, '1:7'))

# ---------- skill 数据（11 列）----------
# tag, 角色, 技能名, 消耗, 目标, 生效, 伤害, 护甲, buff, 额外, 伤害目标
SKILL_MONSTER = [
    ('wolf_skill1',         '狼',       '撕咬', NULL, '近战目标', '自动', 2, 0, NULL, NULL, 'enemy'),
    ('blackWolf_skill1',    '黑狼',     '撕咬', NULL, '近战目标', '自动', 3, 0, NULL, NULL, 'enemy'),
    ('bandit_skill1',       '山贼',     '劈砍', NULL, '近战目标', '自动', 3, 0, NULL, NULL, 'enemy'),
    ('banditArcher_skill1', '山贼弓手', '射击', NULL, '任意目标', '自动', 3, 0, NULL, NULL, 'enemy'),
]
def conv(v):
    if v is None:
        return (None,)
    if isinstance(v, int):
        return ('n', v)
    if v == NULL:
        return ('s', 'null')
    return ('s', v)
skill_rows = []
for i, row in enumerate(SKILL_MONSTER):
    r = 114 + i
    skill_rows.append(build_row(r, [conv(v) for v in row], '1:11'))

# ---------- 改写 sheet7（monster）----------
s7 = z.read('xl/worksheets/sheet7.xml').decode('utf-8')
old_row2 = re.search(r'<row r="2" spans="1:7">.*?</row>', s7, re.S).group(0)
assert 'B2' in old_row2 and 'E2' in old_row2
s7_new = s7.replace(old_row2, ''.join(monster_rows), 1)
s7_new = s7_new.replace('<dimension ref="A1:G2"/>', '<dimension ref="A1:G5"/>')
assert '<row r="5"' in s7_new and 'banditArcher' not in old_row2
# 新增前校验：row2 原只有 2 个单元格，现 4 行各 7 格
assert s7_new.count('<row ') == 5

# ---------- 改写 sheet2（skill）----------
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
assert s2.count('<row ') == 113 and '<row r="114"' not in s2
s2_new = s2.replace('</sheetData>', ''.join(skill_rows) + '</sheetData>')
s2_new = s2_new.replace('<dimension ref="A1:K113"/>', '<dimension ref="A1:K117"/>')
assert '<row r="117"' in s2_new

# ---------- 改写 sharedStrings ----------
m = re.match(r'(.*?<sst[^>]*?count=")(\d+)(" uniqueCount=")(\d+)(">)', ss_xml, re.S)
def count_s(row_xml):
    return len(re.findall(r't="s"', row_xml))
added_refs = sum(count_s(x) for x in monster_rows) + sum(count_s(x) for x in skill_rows)
new_count = int(m.group(2)) + added_refs
new_unique = int(m.group(4)) + len(new_items)
header = m.group(1) + str(new_count) + m.group(3) + str(new_unique) + m.group(5)
new_ss = header + ''.join(old_items) + ''.join(new_items) + '</sst>'
print('new shared strings: %d, new s-cells: %d, count %s->%d, uniqueCount %s->%d'
      % (len(new_items), added_refs, m.group(2), new_count, m.group(4), new_unique))

# ---------- 备份 + 重写 zip ----------
if not os.path.exists(BACKUP):
    shutil.copy2(SRC, BACKUP)
    print('backup ->', BACKUP)

zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED)
replace = {
    'xl/sharedStrings.xml': new_ss.encode('utf-8'),
    'xl/worksheets/sheet2.xml': s2_new.encode('utf-8'),
    'xl/worksheets/sheet7.xml': s7_new.encode('utf-8'),
}
for item in zin.infolist():
    data = replace.get(item.filename, None)
    if data is None:
        data = zin.read(item.filename)
    zout.writestr(item, data)
zout.close(); zin.close(); z.close()
shutil.move(TMP, SRC)
print('written OK')
