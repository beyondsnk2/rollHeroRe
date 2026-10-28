# -*- coding: utf-8 -*-
"""Fix skillDesc (col D) mismatches found in 2x re-verification against card images.
Zip XML direct edit; replaces existing t="s" refs in D cells, reuses sst entries when identical.
count attr (total refs) unchanged; uniqueCount grows only for genuinely new strings.
"""
import re
import shutil
import zipfile

SRC = 'config/config_v1.xlsx'
SHEET = 'xl/worksheets/sheet3.xml'
SST = 'xl/sharedStrings.xml'

FIX = {
    19: '对近战目标造成3点伤害，自身获得3点护甲',      # guard_skill1
    20: '对近战目标造成7点伤害',                       # shieldBearer_skill1
    21: '我方任意单位获得4点护甲',                     # shieldBearer_skill2 (卡面无"自动：")
    22: '对近战目标造成3+当前护甲值的伤害',           # shieldFighter_skill1
    23: '自动：我方任意单位和自身获得3点护甲',         # shieldFighter_skill2
    24: '对近战目标造成9点伤害，自身获得5点护甲',      # shieldGuard_skill1
    25: '对近战目标造成6点伤害',                       # apprenticeKnight_skill1
    26: '自动：我方任意单位和自身获得3点护甲',         # apprenticeKnight_skill2
    27: '对近战目标造成7点伤害，击杀敌人时自身获得3点护甲',  # phalanxSoldier_skill1
    29: '自动：我方前排获得4点护甲',                   # phalanxSoldier_skill2
    30: '对近战目标造成8点伤害',                       # shieldOfficer_skill1
    31: '自动：我方任意单位获得5点护甲',               # shieldOfficer_skill2
    32: '对近战目标造成4+当前护甲值的伤害，连击1',    # dualShieldFighter_skill1
    33: '自动：自身所在列获得7点护甲，为自身的二连盾击增加连击1',  # dualShieldFighter_skill2
    34: '自身获得6点护甲，之后对近战目标造成等同于自身当前护甲值的伤害',  # heavyShieldKnight_skill1
    36: '对近战目标造成6点伤害，连击1',                # whiteKnight_skill1
    37: '自动：我方任意单位和自身获得4点护甲，且攻击力+2',  # whiteKnight_skill2
    38: '对近战目标造成5+当前护甲值的伤害',           # bulwarkKnight_skill1
    39: '自动：自身所在行获得6点护甲，攻击力+2',       # bulwarkKnight_skill2
    58: '对任意目标造成9点伤害，自身获得3点护甲，回合中首次击杀敌人时充能1',  # frostMage_skill1
}


def xml_escape(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def main():
    zin = zipfile.ZipFile(SRC)
    sheet_xml = zin.read(SHEET).decode('utf-8')
    sst_xml = zin.read(SST).decode('utf-8')

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
        txt = txt.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        text_to_idx.setdefault(txt, i)

    new_sis = []
    for row, desc in FIX.items():
        if desc in text_to_idx:
            new_idx = text_to_idx[desc]
        else:
            text_to_idx[desc] = unique + len(new_sis)
            new_idx = text_to_idx[desc]
            new_sis.append('<si><t>%s</t></si>' % xml_escape(desc))

    # replace existing D cell values
    changed = 0
    for row, desc in FIX.items():
        pat = re.compile(r'(<c r="D%d" t="s"><v>)(\d+)(</v></c>)' % row)
        m = pat.search(sheet_xml)
        assert m, f'D{row} cell not found'
        old_idx = int(m.group(2))
        new_idx = text_to_idx[desc]
        if old_idx != new_idx:
            sheet_xml = pat.sub(r'\g<1>%d\g<3>' % new_idx, sheet_xml, count=1)
            changed += 1

    new_unique = unique + len(new_sis)
    new_open = re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % new_unique, sst_open_tag)
    sst_xml = sst_xml[:m_sst_open.start()] + new_open + body + ''.join(new_sis) + '</sst>' + sst_xml[sst_xml.rindex('</sst>') + len('</sst>'):]

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
    print(f'fixed cells: {changed}/{len(FIX)}, new unique strings: {len(new_sis)}, uniqueCount {unique}->{new_unique}, count attr unchanged ({count})')


if __name__ == '__main__':
    main()
