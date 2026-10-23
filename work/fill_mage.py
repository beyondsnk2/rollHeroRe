# -*- coding: utf-8 -*-
"""写入 mage 卡片识别结果：hero 24-36 行、skill 40-66 行；
并把方阵士兵击杀回甲改为 _sub 口径（skill 行 27 改 + 插入 row28，后续行顺延）。
zip 直改，保护内嵌图片。"""
import zipfile, re, shutil

SRC = 'config/config_v1.xlsx'
TMP = 'config/config_v1.tmp.xlsx'

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

NULL = 'null'

def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def build_row(r, values, spans):
    """values: ('s',text) / ('n',num) / (None,) 跳过"""
    cells = []
    for ci, v in enumerate(values, start=1):
        if v[0] is None:
            continue
        ref = '%s%d' % (col_letter(ci), r)
        if v[0] == 's':
            cells.append('<c r="%s" t="s"><v>%d</v></c>' % (ref, S(esc(v[1]))))
        else:
            cells.append('<c r="%s"><v>%s</v></c>' % (ref, v[1]))
    return '<row r="%d" spans="%s">%s</row>' % (r, spans, ''.join(cells))

# ---------- hero 行（24-36）----------
HEROS = [
    # tag, 中文名, [(type,num)x6], baseHp
    ('apprentice',       '学徒',     [('red',1),('red',1),('green',1),('green',1),('green',1),('green',1)], 3),
    ('flameApprentice',  '烈焰学徒', [('red',1),('red',1),('green',2),('green',2),('green',1),('green',1)], 4),
    ('thunderApprentice','雷语学徒', [('red',1),('green',2),('green',2),('green',1),('green',1),('green',1)], 4),
    ('frostApprentice',  '寒冰学徒', [('red',1),('blue',1),('blue',1),('green',2),('green',2),('green',1)], 6),
    ('flameSwordsman',   '炎剑士',   [('red',2),('red',1),('red',1),('green',2),('green',1),('green',1)], 6),
    ('thunderAcolyte',   '雷霆侍僧', [('red',1),('blue',1),('green',2),('green',2),('green',1),('green',1)], 4),
    ('frostAcolyte',     '寒霜侍僧', [('blue',1),('blue',1),('green',2),('green',2),('green',1),('green',1)], 5),
    ('blazeMage',        '爆炎术士', [('red',1),('red',1),('green',3),('green',3),('green',3),('green',2)], 6),
    ('thunderMage',      '雷语法师', [('red',2),('green',3),('green',3),('green',3),('green',2),('green',1)], 6),
    ('frostMage',        '寒冰法师', [('blue',2),('blue',2),('green',3),('green',3),('green',3),('green',2)], 8),
    ('blazeSwordsman',   '爆炎剑士', [('red',2),('red',2),('red',2),('green',3),('green',3),('green',2)], 8),
    ('thunderPriest',    '雷霆祭司', [('red',2),('blue',2),('green',3),('green',3),('green',2),('green',1)], 7),
    ('frostPriest',      '寒霜祭司', [('red',1),('blue',3),('blue',2),('green',3),('green',3),('green',3)], 8),
]
hero_rows = []
for i, (tag, name, slots, hp) in enumerate(HEROS):
    r = 24 + i
    vals = [('s', tag), ('s', name), ('s', 'mage'), (None,)]
    for st, sn in slots:
        vals.append(('s', st))
    for st, sn in slots:
        vals.append(('n', sn))
    vals += [('n', hp), ('n', 0), (None,)]
    hero_rows.append(build_row(r, vals, '1:19'))

# ---------- skill 新增行（mage：40-66）----------
SKILLS = [
    # tag, 角色, 技能名, 消耗, 目标, 生效, 伤害, 恢复护甲, buff, 额外
    ('apprentice_skill1',        '学徒',     '火球',     'red,1;blue,1',        '任意目标',            '选择目标', 6, 0, NULL, NULL),
    ('flameApprentice_skill1',   '烈焰学徒', '大火球',   'red,1;blue,2',        '任意目标',            '选择目标', 15, 0, NULL, NULL),
    ('thunderApprentice_skill1', '雷语学徒', '雷电术',   'green,2',             '任意目标',            '选择目标', 8, 0, NULL, NULL),
    ('thunderApprentice_skill2', '雷语学徒', '静电',     'red,2;blue,1',        '自身',                '自动',     0, 0, '雷电术_充能+1', NULL),
    ('frostApprentice_skill1',   '寒冰学徒', '冰箭术',   'red,1;blue,1;green,1','任意目标',            '选择目标', 6, 3, NULL, NULL),
    ('frostApprentice_skill2',   '寒冰学徒', '冷却',     'green,3',             '自身',                '自动',     0, 3, '攻击力+4', NULL),
    ('flameSwordsman_skill1',    '炎剑士',   '烈焰横扫', 'green,3',             '近战目标及同行',      '选择目标', 5, 0, NULL, NULL),
    ('flameSwordsman_skill2',    '炎剑士',   '火球',     'red,2',               '任意目标',            '选择目标', 6, 0, NULL, NULL),
    ('thunderAcolyte_skill1',    '雷霆侍僧', '闪电链',   'red,1;blue,2',        '近战目标和1个随机目标','选择目标', 5, 0, NULL, NULL),
    ('thunderAcolyte_skill2',    '雷霆侍僧', '超载',     'blue,2',              '自身',                '自动',     0, 0, '闪电链_随机目标+1', NULL),
    ('frostAcolyte_skill1',      '寒霜侍僧', '冰锥术',   'blue,1;green,2',      '任意目标',            '选择目标', 6, 0, NULL, 'frostAcolyte_skill1_sub'),
    ('frostAcolyte_skill1_sub',  '寒霜侍僧', NULL,       NULL,                  '相邻单位',            NULL,       5, 0, NULL, NULL),
    ('frostAcolyte_skill2',      '寒霜侍僧', '霜盾',     'red,2;blue,1',        '自身所在列',          '自动',     0, 4, NULL, NULL),
    ('blazeMage_skill1',         '爆炎术士', '大火球',   'red,1;blue,3',        '任意目标',            '选择目标', 22, 0, NULL, NULL),
    ('blazeMage_skill2',         '爆炎术士', '爆炎火球', 'green,5',             '任意目标及相邻单位',  '选择目标', 4, 0, NULL, NULL),
    ('thunderMage_skill1',       '雷语法师', '轰雷术',   'green,5',             '任意目标',            '选择目标', 8, 0, NULL, 'thunderMage_skill1_sub'),
    ('thunderMage_skill1_sub',   '雷语法师', NULL,       NULL,                  '相邻单位',            NULL,       4, 0, NULL, NULL),
    ('thunderMage_skill2',       '雷语法师', '静电',     'red,3;blue,2',        '自身',                '自动',     0, 0, '轰雷术_充能+1', NULL),
    ('frostMage_skill1',         '寒冰法师', '冰枪术',   'blue,1;green,3',      '任意目标',            '选择目标', 9, 3, NULL, 'frostMage_skill1_sub'),
    ('frostMage_skill1_sub',     '寒冰法师', NULL,       NULL,                  '自身',                NULL,       0, 0, '冰枪术_充能+1', NULL),
    ('frostMage_skill2',         '寒冰法师', '冷却',     'blue,1;green,3',      '自身',                '自动',     0, 5, '攻击力+6', NULL),
    ('blazeSwordsman_skill1',    '爆炎剑士', '烈焰魔剑', 'red,2;green,4',       '近战目标及同行',      '选择目标', 8, 0, NULL, NULL),
    ('blazeSwordsman_skill2',    '爆炎剑士', '怒焰魔剑', 'blue,4;green,2',      '任意列',              '选择目标', 8, 0, NULL, NULL),
    ('thunderPriest_skill1',     '雷霆祭司', '大闪电链', 'green,4',             '近战目标和2个随机目标','选择目标', 6, 0, NULL, NULL),
    ('thunderPriest_skill2',     '雷霆祭司', '超载',     'red,2;blue,2',        '自身',                '自动',     0, 0, '攻击力+2;大闪电链_随机目标+2', NULL),
    ('frostPriest_skill1',       '寒霜祭司', '冰暴术',   'red,2;blue,2',        '任意目标及相邻单位',  '选择目标', 8, 0, NULL, NULL),
    ('frostPriest_skill2',       '寒霜祭司', '霜盾',     'blue,3;green,3',      '自身及相邻单位',      '自动',     0, 5, NULL, NULL),
]
def conv(v):
    if v is None:
        return (None,)
    if isinstance(v, int):
        return ('n', v)
    if v == NULL:
        return ('s', 'null')
    return ('s', v)

mage_rows = []
for i, row in enumerate(SKILLS):
    r = 40 + i
    mage_rows.append(build_row(r, [conv(v) for v in row], '1:10'))

# 方阵士兵击杀回甲 _sub 行（插入为 row28，28+ 顺延）
sub_phalanx = build_row(28, [conv(v) for v in
    ('phalanxSoldier_skill1_sub', '方阵士兵', None, None, '自身', None, 0, 3, None, None)], '1:10')

# ---------- 改写 sheet2（skill）----------
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
assert '<row r="39"' not in s2 and '<row r="40"' not in s2
rows = re.findall(r'<row r="(\d+)"[^>]*>.*?</row>', s2, re.S)
assert [int(r) for r in rows] == list(range(1, 39)), [int(r) for r in rows]
rowmap = {}
for mm in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>', s2, re.S):
    rowmap[int(mm.group(1))] = mm.group(0)

# 修改 row27：I27 → null，J27 → phalanxSoldier_skill1_sub
r27 = rowmap[27]
i27_idx = S('null')
j27_idx = S('phalanxSoldier_skill1_sub')
r27_new = re.sub(r'<c r="I27"[^/]*>.*?</c>', '<c r="I27" t="s"><v>%d</v></c>' % i27_idx, r27)
r27_new = re.sub(r'<c r="J27"[^/]*>.*?</c>', '<c r="J27" t="s"><v>%d</v></c>' % j27_idx, r27_new)
assert r27_new != r27

# 28-38 顺延为 29-39
shifted = {}
for old_r in range(28, 39):
    rx = rowmap[old_r]
    new_r = old_r + 1
    rx2 = rx.replace('<row r="%d"' % old_r, '<row r="%d"' % new_r, 1)
    rx2 = re.sub(r'r="([A-J])%d"' % old_r, r'r="\g<1>%d"' % new_r, rx2)
    shifted[new_r] = rx2

ordered = []
for r in range(1, 28):
    ordered.append(rowmap[r])
ordered.append(r27_new)
ordered.append(sub_phalanx)
for r in range(29, 40):
    ordered.append(shifted[r])
ordered += mage_rows

s2_new = re.sub(r'(<sheetData>).*?(</sheetData>)', lambda m: m.group(1) + ''.join(ordered) + m.group(2), s2, flags=re.S)
s2_new = s2_new.replace('<dimension ref="A1:J38"/>', '<dimension ref="A1:J66"/>')
assert '<row r="28"' in s2_new and '<row r="66"' in s2_new

# ---------- 改写 sheet1（hero）----------
s1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
assert '<row r="24"' not in s1
s1_new = s1.replace('</sheetData>', ''.join(hero_rows) + '</sheetData>')
s1_new = s1_new.replace('<dimension ref="A1:S23"/>', '<dimension ref="A1:S36"/>')
assert '<row r="36"' in s1_new

# ---------- 改写 sharedStrings ----------
m = re.match(r'(.*?<sst[^>]*?count=")(\d+)(" uniqueCount=")(\d+)(">)', ss_xml, re.S)
def count_s(row_xml):
    return len(re.findall(r't="s"', row_xml))
added_refs = sum(count_s(x) for x in hero_rows) + sum(count_s(x) for x in mage_rows) + count_s(sub_phalanx)
new_count = int(m.group(2)) + added_refs
new_unique = int(m.group(4)) + len(new_items)
header = m.group(1) + str(new_count) + m.group(3) + str(new_unique) + m.group(5)
new_ss = header + ''.join(old_items) + ''.join(new_items) + '</sst>'

# ---------- 重写 zip ----------
zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED)
replace = {
    'xl/sharedStrings.xml': new_ss.encode('utf-8'),
    'xl/worksheets/sheet1.xml': s1_new.encode('utf-8'),
    'xl/worksheets/sheet2.xml': s2_new.encode('utf-8'),
}
for item in zin.infolist():
    data = replace.get(item.filename, None)
    if data is None:
        data = zin.read(item.filename)
    zout.writestr(item, data)
zout.close(); zin.close(); z.close()
shutil.move(TMP, SRC)
print('written. new strings: %d, new s-cells: %d' % (len(new_items), added_refs))
