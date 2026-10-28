# -*- coding: utf-8 -*-
"""Repair skillDesc by TAG (row numbers shifted after user deleted _sub rows).
1. Build expected desc by tag (final verified card transcription).
2. Read current tag->row from sheet via openpyxl.
3. XML-direct replace D cell where current != expected; reuse sst entries.
"""
import re
import shutil
import zipfile

import openpyxl

SRC = 'config/config_v1.xlsx'
SHEET = 'xl/worksheets/sheet3.xml'
SST = 'xl/sharedStrings.xml'

EXPECTED = {
    'militia_skill1': '对近战目标造成5点伤害',
    'axeWarrior_skill1': '对近战目标及背后相邻单位造成9点伤害',
    'swordsman_skill1': '对近战目标造成12点伤害',
    'greatSwordsman_skill1': '对近战目标及左右相邻单位造成7点伤害',
    'fistMaster_skill1': '对近战目标造成4点伤害，连击1',
    'militiaCaptain_skill1': '对近战目标造成8点伤害',
    'giantAxeKnight_skill1': '对近战目标及相邻单位造成12点伤害',
    'giantAxeKnight_skill2': '自动：自身攻击力+4',
    'longswordKnight_skill1': '对近战目标造成18点伤害',
    'longswordKnight_skill2': '自动：自身获得6点护甲',
    'greatswordKnight_skill1': '对近战目标及左右相邻单位造成11点伤害',
    'greatswordKnight_skill2': '自动：为自身的横扫增加连击1',
    'swordKnightOfficer_skill1': '对近战目标造成7点伤害，随后我方前排攻击力+3',
    'swordKnightOfficer_skill2': '自动：自身及相邻单位获得3点护甲',
    'martialArtist_skill1': '对近战目标造成4点伤害，连击3',
    'martialArtist_skill2': '自动：为自身的乱拳连打增加连击1',
    'guard_skill1': '对近战目标造成3点伤害，自身获得3点护甲',
    'shieldBearer_skill1': '对近战目标造成7点伤害',
    'shieldBearer_skill2': '我方任意单位获得4点护甲',
    'shieldFighter_skill1': '对近战目标造成3+当前护甲值的伤害',
    'shieldFighter_skill2': '自动：我方任意单位和自身获得3点护甲',
    'shieldGuard_skill1': '对近战目标造成9点伤害，自身获得5点护甲',
    'apprenticeKnight_skill1': '对近战目标造成6点伤害',
    'apprenticeKnight_skill2': '自动：我方任意单位和自身获得3点护甲',
    'phalanxSoldier_skill1': '对近战目标造成7点伤害，击杀敌人时自身获得3点护甲',
    'phalanxSoldier_skill2': '自动：我方前排获得4点护甲',
    'shieldOfficer_skill1': '对近战目标造成8点伤害',
    'shieldOfficer_skill2': '自动：我方任意单位获得5点护甲',
    'dualShieldFighter_skill1': '对近战目标造成4+当前护甲值的伤害，连击1',
    'dualShieldFighter_skill2': '自动：自身所在列获得7点护甲，为自身的二连盾击增加连击1',
    'heavyShieldKnight_skill1': '自身获得6点护甲，之后对近战目标造成等同于自身当前护甲值的伤害',
    'heavyShieldKnight_skill2': '自动：自身所在列获得5点护甲',
    'whiteKnight_skill1': '对近战目标造成6点伤害，连击1',
    'whiteKnight_skill2': '自动：我方任意单位和自身获得4点护甲，且攻击力+2',
    'bulwarkKnight_skill1': '对近战目标造成5+当前护甲值的伤害',
    'bulwarkKnight_skill2': '自动：自身所在行获得6点护甲，攻击力+2',
    'apprentice_skill1': '对任意目标造成6点伤害',
    'flameApprentice_skill1': '对任意目标造成15点伤害',
    'thunderApprentice_skill1': '对任意目标造成8点伤害',
    'thunderApprentice_skill2': '自动：为自身的雷电术增加充能1',
    'frostApprentice_skill1': '对任意目标造成6点伤害，自身获得3点护甲',
    'frostApprentice_skill2': '自动：自身获得3点护甲，攻击力+4',
    'flameSwordsman_skill1': '对近战目标及同行造成5点伤害',
    'flameSwordsman_skill2': '对任意目标造成6点伤害',
    'thunderAcolyte_skill1': '对近战目标和1个随机目标造成5点伤害',
    'thunderAcolyte_skill2': '自动：为自身的闪电链增加1个随机目标',
    'frostAcolyte_skill1': '对任意目标造成6点伤害，击杀敌人时对其相邻单位造成5点伤害',
    'frostAcolyte_skill2': '自动：自身所在列获得4点护甲',
    'blazeMage_skill1': '对任意目标造成22点伤害',
    'blazeMage_skill2': '对任意目标及相邻单位造成4点伤害',
    'thunderMage_skill1': '对任意目标造成8点伤害，对其相邻单位造成4点伤害',
    'thunderMage_skill2': '自动：为自身的轰雷术增加充能1',
    'frostMage_skill1': '对任意目标造成9点伤害，自身获得3点护甲，回合中首次击杀敌人时充能1',
    'frostMage_skill2': '自动：自身获得5点护甲，攻击力+6',
    'blazeSwordsman_skill1': '对近战目标及同行造成8点伤害',
    'blazeSwordsman_skill2': '对任意列造成8点伤害',
    'thunderPriest_skill1': '对近战目标和2个随机目标造成6点伤害',
    'thunderPriest_skill2': '自动：自身攻击力+2，为自身的大闪电链增加2个随机目标',
    'frostPriest_skill1': '对任意目标及相邻单位造成8点伤害',
    'frostPriest_skill2': '自动：自身及相邻单位获得5点护甲',
    'hunter_skill1': '对任意目标造成4点伤害',
    'crossbowman_skill1': '对随机近战目标造成4点伤害，连击2',
    'shortbowHunter_skill1': '对任意目标造成5点伤害，充能1',
    'shortbowHunter_skill2': '自动：为自身的速射增加充能1',
    'archer_skill1': '对任意目标造成10点伤害',
    'archer_skill2': '自动：对随机近战目标造成6点伤害',
    'lightCrossbowman_skill1': '对任意目标及背后相邻单位造成6点伤害',
    'lightCrossbowman_skill2': '自动：自身攻击力+2',
    'ranger_skill1': '对任意目标造成8点伤害，回合中首次击杀敌人时充能1',
    'eliteCrossbowman_skill1': '对随机近战目标造成5点伤害，连击3',
    'swiftHunter_skill1': '对近战目标和1个随机目标造成5点伤害，连击1，充能1',
    'swiftHunter_skill2': '自动：为自身的急速散射增加连击1',
    'longbowman_skill1': '对任意目标造成18点伤害',
    'longbowman_skill2': '对近战目标造成8点伤害',
    'heavyCrossbowman_skill1': '对任意列造成6点伤害，连击1',
    'heavyCrossbowman_skill2': '自动：为自身的贯通连击增加连击1',
    'eliteRanger_skill1': '对任意目标造成10点伤害，回合中首次击杀敌人时充能1',
    'eliteRanger_skill2': '对近战目标施加破碎1',
    'priest_skill1': '我方任意单位获得4点护甲',
    'paladin_skill1': '对近战目标造成10点伤害',
    'paladin_skill2': '自动：自身获得3点护甲，攻击力+3',
    'prayer_skill1': '我方任意单位攻击力+3',
    'prayer_skill2': '对任意目标造成6点伤害',
    'whitePriest_skill1': '我方任意单位和自身获得5点护甲',
    'whitePriest_skill2': '对任意目标造成5点伤害',
    'bloodPriest_skill1': '我方任意单位获得3点护甲，攻击力+3，对自身造成2点伤害',
    'bloodPriest_skill2': '自动：我方前排获得3点护甲',
    'shadowPriest_skill1': '对任意目标施加重伤4',
    'shadowPriest_skill2': '自动：自身所在行攻击力+3',
    'templeWarrior_skill1': '对近战目标造成16点伤害',
    'templeWarrior_skill2': '自动：自身攻击力+5，我方前排获得4点护甲',
    'prophet_skill1': '我方任意单位攻击力+5',
    'prophet_skill2': '自动：为灌注增加充能1',
    'archbishop_skill1': '我方任意单位和自身获得7点护甲',
    'archbishop_skill2': '对任意目标造成5点伤害，充能1',
    'bloodBishop_skill1': '我方前排获得5点护甲，攻击力+5，对自身造成3点伤害',
    'bloodBishop_skill2': '对近战目标及同行造成7点伤害',
    'darkPriest_skill1': '对任意目标及相邻单位施加重伤4',
    'darkPriest_skill2': '对任意目标及相邻单位造成1点伤害，连击1',
    'wolf_skill1': '对近战目标造成2点伤害',
    'blackWolf_skill1': '对近战目标造成3点伤害',
    'bandit_skill1': '对近战目标造成3点伤害',
    'banditArcher_skill1': '对任意目标造成3点伤害',
}


def xml_escape(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def main():
    # tag -> current desc
    wb = openpyxl.load_workbook(SRC, read_only=True)
    ws = wb['skill']
    tag_row = {}
    tag_cur = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        tag, desc = r[0], r[3]
        if tag and tag != 'tag':
            tag_row[tag] = i
            tag_cur[tag] = desc
    wb.close()

    diffs = {t: d for t, d in EXPECTED.items() if tag_cur.get(t) != d}
    unknown = [t for t in tag_cur if t not in EXPECTED and not t.endswith('_sub')]
    print('tags mismatched:', len(diffs), '| unexpected tags:', unknown)

    zin = zipfile.ZipFile(SRC)
    sheet_xml = zin.read(SHEET).decode('utf-8')
    sst_xml = zin.read(SST).decode('utf-8')

    m_sst_open = re.search(r'<sst\b[^>]*>', sst_xml)
    sst_open_tag = m_sst_open.group(0)
    count = int(re.search(r'count="(\d+)"', sst_open_tag).group(1))
    unique = int(re.search(r'uniqueCount="(\d+)"', sst_open_tag).group(1))
    body = sst_xml[m_sst_open.end():sst_xml.rindex('</sst>')]
    si_blocks = re.findall(r'<si>.*?</si>|<si/>', body, re.S)
    assert len(si_blocks) == unique

    text_to_idx = {}
    for i, blk in enumerate(si_blocks):
        txt = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', blk, re.S))
        txt = txt.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        text_to_idx.setdefault(txt, i)

    new_sis = []
    for tag, desc in diffs.items():
        if desc not in text_to_idx:
            text_to_idx[desc] = unique + len(new_sis)
            new_sis.append('<si><t>%s</t></si>' % xml_escape(desc))

    for tag, desc in diffs.items():
        row = tag_row[tag]
        pat = re.compile(r'(<c r="D%d" t="s"><v>)(\d+)(</v></c>)' % row)
        m = pat.search(sheet_xml)
        if not m:
            # D cell missing -> insert after C cell
            cm = re.search(r'(<c r="C%d"[^>]*?(?:/>|>.*?</c>))' % row, sheet_xml, re.S)
            assert cm, f'no C cell for {tag} row {row}'
            cell = '<c r="D%d" t="s"><v>%d</v></c>' % (row, text_to_idx[desc])
            sheet_xml = sheet_xml[:cm.end(1)] + cell + sheet_xml[cm.end(1):]
        else:
            sheet_xml = pat.sub(r'\g<1>%d\g<3>' % text_to_idx[desc], sheet_xml, count=1)

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
    print(f'repaired: {len(diffs)} cells, new unique strings: {len(new_sis)}, uniqueCount {unique}->{new_unique}, count attr unchanged ({count})')
    for t in sorted(diffs, key=lambda x: tag_row[x]):
        print(' ', t, '->', diffs[t])


if __name__ == '__main__':
    main()
