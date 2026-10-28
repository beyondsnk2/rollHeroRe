# -*- coding: utf-8 -*-
"""Fill skillDesc (col D) of skill sheet (sheet3.xml) in config_v1.xlsx from card transcriptions.
Zip XML direct edit (project convention, openpyxl would drop the embedded image in $type sheet).
Rows whose tag ends with _sub are skipped (user confirmed 2026-10-27).
"""
import re
import shutil
import zipfile

SRC = 'config/config_v1.xlsx'
SHEET = 'xl/worksheets/sheet3.xml'
SST = 'xl/sharedStrings.xml'

DESC = {
    2: '对近战目标造成5点伤害',
    3: '对近战目标及背后相邻单位造成9点伤害',
    4: '对近战目标造成12点伤害',
    5: '对近战目标及左右相邻单位造成7点伤害',
    6: '对近战目标造成4点伤害，连击1',
    7: '对近战目标造成8点伤害',
    8: '对近战目标及相邻单位造成12点伤害',
    9: '自动：自身攻击力+4',
    10: '对近战目标造成18点伤害',
    11: '自动：自身获得6点护甲',
    12: '对近战目标及左右相邻单位造成11点伤害',
    13: '自动：为自身的横扫增加连击1',
    14: '对近战目标造成7点伤害，随后我方前排攻击力+3',
    15: '自动：自身及相邻单位获得3点护甲',
    17: '对近战目标造成4点伤害，连击3',
    18: '自动：为自身的乱拳连打增加连击1',
    19: '对近战目标造成5点伤害，自身获得4点护甲',
    20: '对近战目标造成6点伤害',
    21: '自动：我方任意单位获得6点护甲',
    22: '对近战目标造成6点伤害，自身获得3点护甲',
    23: '自动：我方前排获得4点护甲',
    24: '对近战目标造成6点伤害，自身获得6点护甲',
    25: '对近战目标造成8点伤害',
    26: '自动：我方前排获得6点护甲',
    27: '对近战目标造成8点伤害，充能1',
    29: '自动：自身所在列获得8点护甲',
    30: '对近战目标造成10点伤害',
    31: '自动：我方前排获得8点护甲',
    32: '对近战目标造成4点伤害，自身获得4点护甲，连击1',
    33: '自动：自身获得4点护甲，攻击力+2',
    34: '对近战目标造成8点伤害，自身获得8点护甲',
    35: '自动：自身所在列获得5点护甲',
    36: '对近战目标造成5点伤害，连击1',
    37: '自动：为自身的二连斩击增加连击1',
    38: '对近战目标造成8点伤害，自身获得5点护甲',
    39: '自动：自身及相邻单位获得6点护甲',
    40: '对任意目标造成6点伤害',
    41: '对任意目标造成15点伤害',
    42: '对任意目标造成8点伤害',
    43: '自动：为自身的雷电术增加充能1',
    44: '对任意目标造成6点伤害，自身获得3点护甲',
    45: '自动：自身获得3点护甲，攻击力+4',
    46: '对近战目标及同行造成5点伤害',
    47: '对任意目标造成6点伤害',
    48: '对近战目标和1个随机目标造成5点伤害',
    49: '自动：为自身的闪电链增加1个随机目标',
    50: '对任意目标造成6点伤害，击杀敌人时对其相邻单位造成5点伤害',
    52: '自动：自身所在列获得4点护甲',
    53: '对任意目标造成22点伤害',
    54: '对任意目标及相邻单位造成4点伤害',
    55: '对任意目标造成8点伤害，对其相邻单位造成4点伤害',
    57: '自动：为自身的轰雷术增加充能1',
    58: '对任意目标造成9点伤害，自身获得3点护甲，回合中首次攻击击杀敌人时充能1',
    60: '自动：自身获得5点护甲，攻击力+6',
    61: '对近战目标及同行造成8点伤害',
    62: '对任意列造成8点伤害',
    63: '对近战目标和2个随机目标造成6点伤害',
    64: '自动：自身攻击力+2，为自身的大闪电链增加2个随机目标',
    65: '对任意目标及相邻单位造成8点伤害',
    66: '自动：自身及相邻单位获得5点护甲',
    67: '对任意目标造成4点伤害',
    68: '对随机近战目标造成4点伤害，连击2',
    69: '对任意目标造成5点伤害，充能1',
    71: '自动：为自身的速射增加充能1',
    72: '对任意目标造成10点伤害',
    73: '自动：对随机近战目标造成6点伤害',
    74: '对任意目标及背后相邻单位造成6点伤害',
    75: '自动：自身攻击力+2',
    76: '对任意目标造成8点伤害，回合中首次击杀敌人时充能1',
    78: '对随机近战目标造成5点伤害，连击3',
    79: '对近战目标和1个随机目标造成5点伤害，连击1，充能1',
    81: '自动：为自身的急速散射增加连击1',
    82: '对任意目标造成18点伤害',
    83: '对近战目标造成8点伤害',
    84: '对任意列造成6点伤害，连击1',
    85: '自动：为自身的贯通连击增加连击1',
    86: '对任意目标造成10点伤害，回合中首次击杀敌人时充能1',
    88: '对近战目标施加破碎1',
    90: '我方任意单位获得4点护甲',
    91: '对近战目标造成10点伤害',
    92: '自动：自身获得3点护甲，攻击力+3',
    93: '我方任意单位攻击力+3',
    94: '对任意目标造成6点伤害',
    95: '我方任意单位和自身获得5点护甲',
    96: '对任意目标造成5点伤害',
    97: '我方任意单位获得3点护甲，攻击力+3，对自身造成2点伤害',
    98: '自动：我方前排获得3点护甲',
    99: '对任意目标施加重伤4',
    101: '自动：自身所在行攻击力+3',
    102: '对近战目标造成16点伤害',
    103: '自动：自身攻击力+5，我方前排获得4点护甲',
    104: '我方任意单位攻击力+5',
    105: '自动：为灌注增加充能1',
    106: '我方任意单位和自身获得7点护甲',
    107: '对任意目标造成5点伤害，充能1',
    109: '我方前排获得5点护甲，攻击力+5，对自身造成3点伤害',
    110: '对近战目标及同行造成7点伤害',
    111: '对任意目标及相邻单位施加重伤4',
    113: '对任意目标及相邻单位造成1点伤害，连击1',
    114: '对近战目标造成2点伤害',
    115: '对近战目标造成3点伤害',
    116: '对近战目标造成3点伤害',
    117: '对任意目标造成3点伤害',
}


def xml_escape(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def main():
    zin = zipfile.ZipFile(SRC)
    sheet_xml = zin.read(SHEET).decode('utf-8')
    sst_xml = zin.read(SST).decode('utf-8')

    # --- parse existing sst entries, build text->index map ---
    m_sst_open = re.search(r'<sst\b[^>]*>', sst_xml)
    sst_open_tag = m_sst_open.group(0)
    count = int(re.search(r'count="(\d+)"', sst_open_tag).group(1))
    unique = int(re.search(r'uniqueCount="(\d+)"', sst_open_tag).group(1))

    body = sst_xml[m_sst_open.end():sst_xml.rindex('</sst>')]
    si_blocks = re.findall(r'<si>.*?</si>|<si/>', body, re.S)
    assert len(si_blocks) == unique, f'sst blocks {len(si_blocks)} != uniqueCount {unique}'

    text_to_idx = {}
    for i, blk in enumerate(si_blocks):
        txt = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', blk, re.S))
        # unescape for comparison key
        txt = txt.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        text_to_idx.setdefault(txt, i)

    # --- assign indices for descriptions (reuse existing strings when identical) ---
    idx_for_row = {}
    new_sis = []
    for row, desc in DESC.items():
        if desc in text_to_idx:
            idx_for_row[row] = text_to_idx[desc]
        else:
            text_to_idx[desc] = unique + len(new_sis)
            idx_for_row[row] = text_to_idx[desc]
            new_sis.append('<si><t>%s</t></si>' % xml_escape(desc))

    n_new_unique = len(new_sis)
    n_refs = len(DESC)

    # --- edit sheet3.xml: insert <c r="Dn" t="s"> after Cn cell ---
    # verify no existing D cells in target rows
    for row in DESC:
        assert not re.search(r'<c r="D%d"[^>]*>' % row, sheet_xml), f'D{row} already exists'

    missing_c = []
    for row in DESC:
        cm = re.search(r'(<c r="C%d"[^>]*?(?:/>|>.*?</c>))' % row, sheet_xml, re.S)
        if not cm:
            missing_c.append(row)
            continue
        new_cell = '<c r="D%d" t="s"><v>%d</v></c>' % (row, idx_for_row[row])
        sheet_xml = sheet_xml[:cm.end(1)] + new_cell + sheet_xml[cm.end(1):]
    assert not missing_c, f'rows without C cell: {missing_c}'

    # --- update sst header attrs and append new si entries ---
    new_count = count + n_refs
    new_unique = unique + n_new_unique
    new_open = re.sub(r'count="\d+"', 'count="%d"' % new_count, sst_open_tag)
    new_open = re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % new_unique, new_open)
    sst_xml = sst_xml[:m_sst_open.start()] + new_open + body + ''.join(new_sis) + '</sst>' + sst_xml[sst_xml.rindex('</sst>') + len('</sst>'):]

    # --- rewrite zip ---
    tmp = SRC + '.tmp'
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == SHEET:
            data = sheet_xml.encode('utf-8')
        elif item.filename == SST:
            data = sst_xml.encode('utf-8')
        zout.writestr(item, data)
    zout.close()
    zin.close()
    shutil.move(tmp, SRC)
    print(f'done: refs added={n_refs}, new unique strings={n_new_unique}, sst count {count}->{new_count}, uniqueCount {unique}->{new_unique}')


if __name__ == '__main__':
    main()
