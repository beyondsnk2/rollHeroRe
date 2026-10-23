# -*- coding: utf-8 -*-
"""向 config_v1.xlsx 的 hero/skill 页签写入 tank 卡片识别结果（zip 直改，保护内嵌图片）。"""
import zipfile, re, shutil, os
from xml.etree import ElementTree as ET

SRC = 'config/config_v1.xlsx'
TMP = 'config/config_v1.tmp.xlsx'

# ---------- 1. 读取 sharedStrings，建立 复用/新增 机制 ----------
z = zipfile.ZipFile(SRC)
ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8')
sst_head = re.match(r'.*?<sst[^>]*>', ss_xml, re.S).group(0)
existing = [re.sub(r'^<si>|</si>$', '', it) for it in re.findall(r'<si>.*?</si>', ss_xml, re.S)]
# 归一化：<si><t>xx</t></si> -> xx
def plain(s):
    m = re.match(r'^<t[^>]*>(.*)</t>$', s, re.S)
    return m.group(1) if m else s
idx_of = {}
for i, it in enumerate(existing):
    idx_of.setdefault(plain(it), i)
new_items = []
def S(text):
    """返回文本对应的 shared string idx，不存在则追加。"""
    if text not in idx_of:
        idx_of[text] = len(existing) + len(new_items)
        new_items.append('<si><t>%s</t></si>' % text)
    return idx_of[text]

def esc(t):
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

# 字符串统一先转义再入 S（复用表里的原文不含转义字符，安全）
def Sesc(text):
    return S(esc(text))

# ---------- 2. hero 行数据（rows 13-23）----------
HERO_HEADER = ['tag', None, 'class', 'tier', 'slot1Type', 'slot2Type', 'slot3Type', 'slot4Type',
               'slot5Type', 'slot6Type', 'slot1Num', 'slot2Num', 'slot3Num', 'slot4Num',
               'slot5Num', 'slot6Num', 'baseHp', 'baseShield', 'skillList']
HEROS = [
    # tag, 中文名, slots(type,num)x6, baseHp, baseShield
    ('guard',             '守卫',     [('red',1),('red',1),('blue',1),('blue',1),('blue',1),('blue',1)], 7),
    ('shieldBearer',      '持盾侍从', [('red',2),('red',2),('blue',2),('blue',2),('blue',2),('blue',2)], 9),
    ('shieldFighter',     '盾斗士',   [('red',1),('red',1),('red',1),('blue',2),('blue',2),('blue',1)], 8),
    ('shieldGuard',       '盾卫',     [('red',1),('red',1),('blue',2),('blue',2),('blue',1),('blue',1)], 8),
    ('apprenticeKnight',  '见习骑士', [('red',1),('red',1),('blue',2),('blue',2),('blue',2),('blue',1)], 8),
    ('phalanxSoldier',    '方阵士兵', [('red',1),('blue',2),('blue',2),('blue',1),('blue',1),('blue',1)], 9),
    ('shieldOfficer',     '盾卫士官', [('red',3),('blue',3),('blue',3),('blue',3),('blue',3),('yellow',1)], 12),
    ('dualShieldFighter', '双盾斗士', [('red',3),('red',3),('red',2),('blue',2),('blue',2),('blue',2)], 11),
    ('heavyShieldKnight', '重盾骑士', [('red',2),('blue',3),('blue',3),('blue',2),('blue',2),('blue',2)], 10),
    ('whiteKnight',       '白骑士',   [('red',1),('blue',3),('blue',3),('blue',2),('blue',2),('green',2)], 10),
    ('bulwarkKnight',     '壁垒骑士', [('red',1),('blue',4),('blue',3),('blue',3),('blue',2),('green',1)], 11),
]

def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def build_row(r, values, spans):
    """values: list of (kind, val)；kind: 's'=字符串, 'n'=数字, None=空格子跳过"""
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

hero_rows = []
for i, (tag, name, slots, hp) in enumerate(HEROS):
    r = 13 + i
    vals = [('s', tag), ('s', name), ('s', 'tank'), (None, None)]
    for st, sn in slots:
        vals.append(('s', st))
    for st, sn in slots:
        vals.append(('n', sn))
    vals += [('n', hp), ('n', 0), (None, None)]  # baseHp, baseShield, skillList 留空
    hero_rows.append(build_row(r, vals, '1:19'))

# ---------- 3. skill 行数据（rows 19+，10 列）----------
NULL = 'null'
SKILLS = [
    # tag, 角色, 技能名, 消耗, 目标, 生效, 伤害, 恢复护甲, buff, 额外
    ('guard_skill1',             '守卫',     '持盾攻击', 'red,1;blue,1',  '近战目标',       '选择目标', 3, 3, NULL, NULL),
    ('shieldBearer_skill1',      '持盾侍从', '劈砍',     'red,1',         '近战目标',       '选择目标', 7, 0, NULL, NULL),
    ('shieldBearer_skill2',      '持盾侍从', '保护',     'blue,1',        '我方任意单位',   '自动',     0, 4, NULL, NULL),
    ('shieldFighter_skill1',     '盾斗士',   '持盾突击', 'red,2',         '近战目标',       '选择目标', '3+当前护甲值', 0, NULL, NULL),
    ('shieldFighter_skill2',     '盾斗士',   '援护',     'blue,2',        '我方任意单位和自身', '自动', 0, 3, NULL, NULL),
    ('shieldGuard_skill1',       '盾卫',     '持盾攻击', 'red,1;blue,2',  '近战目标',       '选择目标', 9, 5, NULL, NULL),
    ('apprenticeKnight_skill1',  '见习骑士', '劈砍',     'red,1',         '近战目标',       '选择目标', 6, 0, NULL, NULL),
    ('apprenticeKnight_skill2',  '见习骑士', '援护',     'blue,2',        '我方任意单位和自身', '自动', 0, 3, NULL, NULL),
    ('phalanxSoldier_skill1',    '方阵士兵', '蓄能冲击', 'red,1;blue,1',  '近战目标',       '选择目标', 7, 0, '击杀敌人时自身护甲+3', NULL),
    ('phalanxSoldier_skill2',    '方阵士兵', '竖守',     'blue,3',        '我方前排',       '自动',     0, 4, NULL, NULL),
    ('shieldOfficer_skill1',     '盾卫士官', '劈砍',     'red,3',         '近战目标',       '选择目标', 8, 0, NULL, NULL),
    ('shieldOfficer_skill2',     '盾卫士官', '援护',     'blue,1;green,2','我方任意单位',   '自动',     0, 5, NULL, NULL),
    ('dualShieldFighter_skill1', '双盾斗士', '二连盾击', 'red,3;blue,2',  '近战目标',       '选择目标', '4+当前护甲值', 0, '连击+1', NULL),
    ('dualShieldFighter_skill2', '双盾斗士', '蓄力',     'blue,3;green,2','自身所在列',     '自动',     0, 7, '二连盾击_连击+1', NULL),
    ('heavyShieldKnight_skill1', '重盾骑士', '重盾猛击', 'blue,4',        '近战目标',       '选择目标', '自身当前护甲值', 6, NULL, NULL),
    ('heavyShieldKnight_skill2', '重盾骑士', '阻挡',     'red,1;blue,2',  '自身所在列',     '自动',     0, 5, NULL, NULL),
    ('whiteKnight_skill1',       '白骑士',   '二连斩击', 'red,2;green,2', '近战目标',       '选择目标', 6, 0, '连击+1', NULL),
    ('whiteKnight_skill2',       '白骑士',   '祈福',     'red,2;blue,2',  '我方任意单位和自身', '自动', 0, 4, '攻击力+2', NULL),
    ('bulwarkKnight_skill1',     '壁垒骑士', '持盾突击', 'red,2',         '近战目标',       '选择目标', '5+当前护甲值', 0, NULL, NULL),
    ('bulwarkKnight_skill2',     '壁垒骑士', '壁垒',     'blue,4;green,1','自身所在行',     '自动',     0, 6, '攻击力+2', NULL),
]

skill_rows = []
for i, row in enumerate(SKILLS):
    r = 19 + i
    vals = []
    for v in row:
        if isinstance(v, (int, float)):
            vals.append(('n', int(v)))
        elif v == NULL:
            vals.append(('s', 'null'))
        else:
            vals.append(('s', v))
    skill_rows.append(build_row(r, vals, '1:10'))

# ---------- 4. 改写 sharedStrings ----------
old_si = re.findall(r'<si>.*?</si>', ss_xml, re.S)
new_si = old_si + new_items
m = re.match(r'(.*?<sst[^>]*?count=")(\d+)(" uniqueCount=")(\d+)(">)', ss_xml, re.S)
# 统计新增 t="s" 单元格数：hero 行 + skill 行里的字符串格数
def count_s(row_xml):
    return len(re.findall(r't="s"', row_xml))
added_refs = sum(count_s(x) for x in hero_rows) + sum(count_s(x) for x in skill_rows)
new_count = int(m.group(2)) + added_refs
new_unique = int(m.group(4)) + len(new_items)
header = m.group(1) + str(new_count) + m.group(3) + str(new_unique) + m.group(5)
new_ss = header + ''.join(old_si) + ''.join(new_items) + '</sst>'

# ---------- 5. 改写 sheet1/sheet2 ----------
s1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
assert '<row r="13"' not in s1
s1_new = s1.replace('</sheetData>', ''.join(hero_rows) + '</sheetData>')
s1_new = s1_new.replace('<dimension ref="A1:S12"/>', '<dimension ref="A1:S23"/>')
assert s1_new != s1 and '<row r="23"' in s1_new

s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
assert '<row r="19"' not in s2
s2_new = s2.replace('</sheetData>', ''.join(skill_rows) + '</sheetData>')
s2_new = s2_new.replace('<dimension ref="A1:J18"/>', '<dimension ref="A1:J38"/>')
assert s2_new != s2 and '<row r="38"' in s2_new

# ---------- 6. 重写 zip ----------
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
print('written. new shared strings: %d, new s-cells: %d' % (len(new_items), added_refs))
