# -*- coding: utf-8 -*-
"""向 config_v1.xlsx 写入 ranger/priest 卡片识别结果：
- hero 页签追加 37-58 行（ranger 37-47，priest 48-58）
- skill 页签追加 67-113 行（ranger 67-89，priest 90-113）
- skill 页签新增 K 列"伤害目标"（self/enemy），现有行 技能伤害>0 补 enemy
zip 直改 sharedStrings + sheet1/sheet2，其余条目原样透传。
"""
import zipfile, re, shutil, os
from xml.etree import ElementTree as ET

SRC = 'config/config_v1.xlsx'
TMP = 'config/config_v1.tmp.xlsx'
BACKUP = 'config_v1_backup_20261024.xlsx'

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
    """text 为原文（未转义）；不存在则追加。"""
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
    """values: ('s',text) / ('n',num) / (None,) 跳过"""
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

# ---------- hero 数据 ----------
HERO_RANGER = [
    # tag, 中文名, slot类型x6, slot数量x6, baseHp
    ('hunter',           '猎人',       ['red','red','red','blue','blue','green'],  [1,1,1,1,1,1], 4),
    ('crossbowman',      '连弩射手',   ['red','red','green','green','green','green'],[1,1,2,2,1,1], 6),
    ('shortbowHunter',   '短弓猎手',   ['red','red','blue','blue','green','green'],  [2,1,2,1,2,1], 5),
    ('archer',           '弓箭手',     ['red','red','red','blue','blue','green'],    [2,2,1,2,1,2], 5),
    ('lightCrossbowman', '轻弩射手',   ['red','red','red','red','blue','blue'],      [2,2,1,1,1,1], 5),
    ('ranger',           '游侠',       ['red','red','red','blue','blue','blue'],    [2,1,1,2,1,1], 6),
    ('eliteCrossbowman', '连弩精锐',   ['red','red','green','green','green','green'],[2,2,3,3,2,1], 8),
    ('swiftHunter',      '急速猎手',   ['red','red','blue','blue','green','green'],  [3,1,3,1,3,1], 7),
    ('longbowman',       '长弓手',     ['red','red','red','blue','blue','green'],    [3,3,2,3,2,3], 7),
    ('heavyCrossbowman', '重弩射手',   ['red','red','red','red','blue','blue'],      [3,3,2,2,2,2], 7),
    ('eliteRanger',      '精锐游侠',   ['red','red','red','blue','blue','blue'],    [3,3,2,2,2,1], 8),
]
HERO_PRIEST = [
    ('priest',        '牧师',       ['blue','blue','blue','green','green','green'],  [1,1,1,1,1,1], 4),
    ('paladin',       '圣武士',     ['red','red','blue','blue','green','green'],      [1,1,2,1,2,1], 8),
    ('prayer',        '祈祷士',     ['red','blue','green','green','green','green'],   [1,1,2,2,1,1], 6),
    ('whitePriest',   '白牧师',     ['blue','blue','blue','green','green','green'],   [2,1,1,2,1,1], 6),
    ('bloodPriest',   '鲜血祭司',   ['blue','green','green','green','green','green'], [1,2,2,1,1,1], 8),
    ('shadowPriest',  '影牧师',     ['red','red','blue','green','green','green'],     [2,1,1,1,1,1], 7),
    ('templeWarrior', '神殿武士',   ['red','red','blue','blue','green','green'],      [4,4,4,1,4,1], 12),
    ('prophet',       '祈祷先知',   ['red','blue','blue','green','green','green'],    [1,3,1,3,3,2], 8),
    ('archbishop',    '牧首',       ['blue','blue','blue','green','green','green'],   [3,3,2,3,2,1], 7),
    ('bloodBishop',   '鲜血主教',   ['blue','green','green','green','green','green'], [3,1,3,2,1,1], 11),
    ('darkPriest',    '暗影牧师',   ['red','red','blue','blue','green','green'],      [2,1,2,1,3,3], 9),
]

hero_rows = []
all_heros = [(t, n, 'ranger', st, sn, hp) for t, n, st, sn, hp in HERO_RANGER] + \
            [(t, n, 'priest', st, sn, hp) for t, n, st, sn, hp in HERO_PRIEST]
for i, (tag, name, cls, types, nums, hp) in enumerate(all_heros):
    r = 37 + i
    vals = [('s', tag), ('s', name), ('s', cls), (None,)]          # A tag, B 中文名, C class, D tier 留空
    for t in types:
        vals.append(('s', t))                                       # E-J slotType
    for n in nums:
        vals.append(('n', n))                                       # K-P slotNum
    vals += [('n', hp), ('n', 0), (None,)]                          # Q baseHp, R baseShield, S skillList 留空
    hero_rows.append(build_row(r, vals, '1:19'))

# ---------- skill 数据（11 列：A-J 原有 + K 伤害目标）----------
NULL = 'null'
SKILL_RANGER = [
    # tag, 角色, 技能名, 消耗, 目标, 生效, 伤害, 护甲, buff, 额外, 伤害目标
    ('hunter_skill1',           '猎人',     '射击',       'red,1',          '任意目标',                 '选择目标', 4,  0, NULL, NULL,                    'enemy'),
    ('crossbowman_skill1',      '连弩射手', '连段射击',   'red,1;green,2',  '随机近战目标',             '选择目标', 4,  0, '连击+2', NULL,                 'enemy'),
    ('shortbowHunter_skill1',   '短弓猎手', '速射',       'red,2;blue,2',   '任意目标',                 '选择目标', 5,  0, NULL, 'shortbowHunter_skill1_sub', 'enemy'),
    ('shortbowHunter_skill1_sub','短弓猎手',None,          None,             '自身',                     'null',     0,  0, '速射_充能+1', NULL,          None),
    ('shortbowHunter_skill2',   '短弓猎手', '弹药装填',   'green,2',        '自身',                     '自动',     0,  0, '速射_充能+1', NULL,          None),
    ('archer_skill1',           '弓箭手',   '射击',       'red,1;green,1',  '任意目标',                 '选择目标', 10, 0, NULL, NULL,                    'enemy'),
    ('archer_skill2',           '弓箭手',   '急速射击',   'red,1;blue,1',   '随机近战目标',             '自动',     6,  0, NULL, NULL,                    'enemy'),
    ('lightCrossbowman_skill1', '轻弩射手', '贯通射击',   'red,1;green,2',  '任意目标及背后相邻单位',   '选择目标', 6,  0, NULL, NULL,                    'enemy'),
    ('lightCrossbowman_skill2', '轻弩射手', '强弩蓄力',   'green,2',        '自身',                     '自动',     0,  0, '攻击力+2', NULL,              None),
    ('ranger_skill1',           '游侠',     '刺杀',       'red,1;green,2',  '任意目标',                 '选择目标', 8,  0, NULL, 'ranger_skill1_sub',       'enemy'),
    ('ranger_skill1_sub',       '游侠',     None,          None,             '自身',                     'null',     0,  0, '刺杀_充能+1', NULL,          None),
    ('eliteCrossbowman_skill1', '连弩精锐', '连弩速射',   'red,1;green,4',  '随机近战目标',             '选择目标', 5,  0, '连击+3', NULL,                 'enemy'),
    ('swiftHunter_skill1',      '急速猎手', '急速散射',   'red,3;blue,2',   '近战目标和1个随机目标',     '选择目标', 5,  0, '连击+1', 'swiftHunter_skill1_sub','enemy'),
    ('swiftHunter_skill1_sub',  '急速猎手', None,          None,             '自身',                     'null',     0,  0, '急速散射_充能+1', NULL,      None),
    ('swiftHunter_skill2',      '急速猎手', '弹药装填',   'blue,3;green,2', '自身',                     '自动',     0,  0, '急速散射_连击+1', NULL,      None),
    ('longbowman_skill1',       '长弓手',   '射击',       'red,3;green,2',  '任意目标',                 '选择目标', 18, 0, NULL, NULL,                    'enemy'),
    ('longbowman_skill2',       '长弓手',   '匕首斩击',   'red,1;blue,2',   '近战目标',                 '选择目标', 8,  0, NULL, NULL,                    'enemy'),
    ('heavyCrossbowman_skill1', '重弩射手', '贯通连击',   'red,3;green,2',  '任意列',                   '选择目标', 6,  0, '连击+1', NULL,                 'enemy'),
    ('heavyCrossbowman_skill2', '重弩射手', '弹药装填',   'blue,4',         '自身',                     '自动',     0,  0, '贯通连击_连击+1', NULL,      None),
    ('eliteRanger_skill1',      '精锐游侠', '刺杀',       'red,3;green,2',  '任意目标',                 '选择目标', 10, 0, NULL, 'eliteRanger_skill1_sub',  'enemy'),
    ('eliteRanger_skill1_sub',  '精锐游侠', None,          None,             '自身',                     'null',     0,  0, '刺杀_充能+1', NULL,          None),
    ('eliteRanger_skill2',      '精锐游侠', '预谋',       'blue,4',         '近战目标',                 '选择目标', 0,  0, NULL, 'eliteRanger_skill2_sub',  None),
    ('eliteRanger_skill2_sub',  '精锐游侠', None,          None,             '近战目标',                 'null',     0,  0, '破碎+1', NULL,                None),
]
SKILL_PRIEST = [
    ('priest_skill1',           '牧师',       '保护',       'green,1',        '我方任意单位',             '自动',     0,  4, NULL, NULL,                        None),
    ('paladin_skill1',          '圣武士',     '辉光斩',     'green,3',        '近战目标',                 '选择目标', 10, 0, NULL, NULL,                        'enemy'),
    ('paladin_skill2',          '圣武士',     '圣光之力',   'red,1;blue,2',   '自身',                     '自动',     0,  3, '攻击力+3', NULL,                  None),
    ('prayer_skill1',           '祈祷士',     '灌注',       'green,2',        '我方任意单位',             '自动',     0,  0, '攻击力+3', NULL,                  None),
    ('prayer_skill2',           '祈祷士',     '圣光术',     'red,1;blue,1',   '任意目标',                 '选择目标', 6,  0, NULL, NULL,                        'enemy'),
    ('whitePriest_skill1',      '白牧师',     '圣盾术',     'blue,1;green,1', '我方任意单位和自身',       '自动',     0,  5, NULL, NULL,                        None),
    ('whitePriest_skill2',      '白牧师',     '闪光术',     'red,1;blue,2',   '任意目标',                 '选择目标', 5,  0, NULL, NULL,                        'enemy'),
    ('bloodPriest_skill1',      '鲜血祭司',   '鲜血灌注',   'green,3',        '我方任意单位',             '自动',     2,  3, '攻击力+3', NULL,                  'self'),
    ('bloodPriest_skill2',      '鲜血祭司',   '鲜血护盾',   'green,3',        '我方前排',                 '自动',     0,  3, NULL, NULL,                        None),
    ('shadowPriest_skill1',     '影牧师',     '重伤',       'green,3',        '任意目标',                 '选择目标', 0,  0, NULL, 'shadowPriest_skill1_sub',   None),
    ('shadowPriest_skill1_sub', '影牧师',     None,          None,             '任意目标',                 'null',     0,  0, '重伤+4', NULL,                    None),
    ('shadowPriest_skill2',     '影牧师',     '暗影之力',   'red,2',          '自身所在行',               '自动',     0,  0, '攻击力+3', NULL,                  None),
    ('templeWarrior_skill1',    '神殿武士',   '圣光斩',     'red,1;green,4',  '近战目标',                 '选择目标', 16, 0, NULL, NULL,                        'enemy'),
    ('templeWarrior_skill2',    '神殿武士',   '圣光帷幕',   'blue,4',         '自身及我方前排',           '自动',     0,  4, '攻击力+5', NULL,                  None),
    ('prophet_skill1',          '祈祷先知',   '灌注',       'green,4',        '我方任意单位',             '自动',     0,  0, '攻击力+5', NULL,                  None),
    ('prophet_skill2',          '祈祷先知',   '虔诚',       'red,2;blue,2',   '自身',                     '自动',     0,  0, '灌注_充能+1', NULL,                None),
    ('archbishop_skill1',       '牧首',       '圣盾术',     'blue,3;green,2', '我方任意单位和自身',       '自动',     0,  7, NULL, NULL,                        None),
    ('archbishop_skill2',       '牧首',       '闪光术',     'red,1;green,4',  '任意目标',                 '选择目标', 5,  0, NULL, 'archbishop_skill2_sub',     'enemy'),
    ('archbishop_skill2_sub',   '牧首',       None,          None,             '自身',                     'null',     0,  0, '闪光术_充能+1', NULL,            None),
    ('bloodBishop_skill1',      '鲜血主教',   '鲜血灌注',   'red,1;green,3',  '我方前排',                 '自动',     3,  5, '攻击力+5', NULL,                  'self'),
    ('bloodBishop_skill2',      '鲜血主教',   '鲜血浪涌',   'red,3;blue,2',   '近战目标及同行',           '选择目标', 7,  0, NULL, NULL,                        'enemy'),
    ('darkPriest_skill1',       '暗影牧师',   '群体重伤',   'red,1;green,4',  '任意目标及相邻单位',       '选择目标', 0,  0, NULL, 'darkPriest_skill1_sub',     None),
    ('darkPriest_skill1_sub',   '暗影牧师',   None,          None,             '任意目标及相邻单位',       'null',     0,  0, '重伤+4', NULL,                    None),
    ('darkPriest_skill2',       '暗影牧师',   '暗影飞弹',   'red,5',          '任意目标及相邻单位',       '选择目标', 1,  0, '连击+1', NULL,                    'enemy'),
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
for i, row in enumerate(SKILL_RANGER + SKILL_PRIEST):
    r = 67 + i
    skill_rows.append(build_row(r, [conv(v) for v in row], '1:11'))

# ---------- 改写 sheet2（skill）----------
s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
assert '<row r="67"' not in s2 and '<row r="37"' not in z.read('xl/worksheets/sheet1.xml').decode('utf-8')
assert s2.count('<row ') == 66

# 1) 表头行加 K1"伤害目标"，spans 1:10 -> 1:11
hdr = re.search(r'<row r="1" spans="1:10">.*?</row>', s2, re.S).group(0)
k1_idx = S('伤害目标')
hdr_new = hdr.replace('spans="1:10"', 'spans="1:11"').replace('</row>', '<c r="K1" t="s"><v>%d</v></c></row>' % k1_idx)
assert hdr_new != hdr
s2 = s2.replace(hdr, hdr_new, 1)

# 2) 现有行 2-66：技能伤害>0（数值）或为伤害表达式字符串 -> K=enemy
enemy_idx = S('enemy')
rowmap = {}
for mm in re.finditer(r'<row r="(\d+)" spans="1:10">.*?</row>', s2, re.S):
    rowmap[int(mm.group(1))] = mm.group(0)
assert sorted(rowmap) == list(range(2, 67)), sorted(rowmap)
existing_enemy_rows = []
for r in range(2, 67):
    rx = rowmap[r]
    mg = re.search(r'<c r="G%d"( t="s")?><v>([^<]*)</v></c>' % r, rx)
    assert mg, 'row %d 无 G 单元格' % r
    if mg.group(1):  # 字符串：伤害表达式（表内无 G=null 的现有行）
        is_dmg = True
    else:
        is_dmg = int(mg.group(2)) > 0
    if is_dmg:
        rx_new = rx.replace('spans="1:10"', 'spans="1:11"').replace(
            '</row>', '<c r="K%d" t="s"><v>%d</v></c></row>' % (r, enemy_idx))
        existing_enemy_rows.append((r, rx, rx_new))
for r, old, new in existing_enemy_rows:
    s2 = s2.replace(old, new, 1)
print('existing rows patched with K=enemy: %d 行 -> %s' % (len(existing_enemy_rows), [r for r, _, _ in existing_enemy_rows]))

# 3) 追加新行 + dimension
s2_new = s2.replace('</sheetData>', ''.join(skill_rows) + '</sheetData>')
s2_new = s2_new.replace('<dimension ref="A1:J66"/>', '<dimension ref="A1:K113"/>')
assert '<row r="113"' in s2_new and '<dimension ref="A1:K113"/>' in s2_new

# ---------- 改写 sheet1（hero）----------
s1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
assert '<row r="37"' not in s1 and s1.count('<row ') == 36
s1_new = s1.replace('</sheetData>', ''.join(hero_rows) + '</sheetData>')
s1_new = s1_new.replace('<dimension ref="A1:S36"/>', '<dimension ref="A1:S58"/>')
assert '<row r="58"' in s1_new and '<row r="37"' in s1_new

# ---------- 改写 sharedStrings ----------
m = re.match(r'(.*?<sst[^>]*?count=")(\d+)(" uniqueCount=")(\d+)(">)', ss_xml, re.S)
def count_s(row_xml):
    return len(re.findall(r't="s"', row_xml))
added_refs = (sum(count_s(hdr_new) for _ in [0])
              + sum(count_s(n) for _, _, n in existing_enemy_rows)
              + sum(count_s(x) for x in skill_rows)
              + sum(count_s(x) for x in hero_rows))
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
print('written OK')
