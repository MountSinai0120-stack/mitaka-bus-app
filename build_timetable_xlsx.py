# -*- coding: utf-8 -*-
"""
三鷹駅南口発 平日 時刻表 Excel 生成スクリプト（構造のみ版）

このスクリプトは「凡例・読み方」「のりば一覧」と、
のりば(2/3/5/6/7/8番)ごとの時刻記入用シート(6〜24時)を作る。
1分単位の発車時刻と前運用番号は写真からは正確に読めないため、
ここでは枠だけ用意し、あとから正確に埋められるようにしている。
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ------- 共通スタイル -------
TITLE = Font(name="Yu Gothic", size=14, bold=True)
HEAD = Font(name="Yu Gothic", size=11, bold=True, color="FFFFFF")
BOLD = Font(name="Yu Gothic", size=11, bold=True)
NORM = Font(name="Yu Gothic", size=11)
SMALL = Font(name="Yu Gothic", size=9, color="6B7280")
RED = Font(name="Yu Gothic", size=11, bold=True, color="C00000")     # 出庫
BLUE = Font(name="Yu Gothic", size=11, bold=True, color="1F4E79")    # 終点から回送

HEAD_FILL = PatternFill("solid", fgColor="2563EB")
HOUR_FILL = PatternFill("solid", fgColor="E8EEF7")
YELLOW = PatternFill("solid", fgColor="FFF2A8")   # 前運用 鷹XX
ORANGE = PatternFill("solid", fgColor="FCD9A8")   # 出庫系(要確認)
GREY = PatternFill("solid", fgColor="F2F4F7")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

thin = Side(style="thin", color="BFC6CF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ------- のりば定義（写真ヘッダ＋本人説明から） -------
# routes: (系統番号, 行き先まとめ)
# marks : (記号, 意味)  ← そののりばの行き先略号凡例
PLATFORMS = [
    {
        "no": "2",
        "header": "55 調布 ・ 56 神代植物公園 ・ 新小金井 ・ 営業所",
        "routes": [
            ("56", "調布駅(北口) / 神代植物公園(植物公園駅)"),
            ("53", "新小金井駅"),
            ("57", "営業所(武蔵境営業所)"),
        ],
        "marks": [
            ("無印", "調布駅(北口)行き"),
            ("営", "営業所(武蔵境営業所)行き"),
            ("新", "新小金井駅行き"),
            ("植", "神代植物公園前行き"),
        ],
    },
    {
        "no": "3",
        "header": "武蔵境 ・ 営業所 ・ 神代植物公園 ・ 深大寺 ・ 60 調布",
        "routes": [
            ("65", "神代植物公園 ・ 深大寺"),
            ("57", "武蔵境駅 ・ 武蔵境営業所"),
            ("66", "(60 調布 ほか) ※要確認"),
        ],
        "marks": [
            ("無印", "武蔵境駅行き"),
            ("営", "武蔵境営業所行き"),
            ("植", "神代植物公園行き"),
            ("深", "深大寺行き"),
            ("K", "京王(調布方面)"),
        ],
    },
    {
        "no": "5",
        "header": "キリスト教大学 ・ 51 調布 ・ 51 飛行場 ・ 大沢 ・ 十字路 ・ 西野",
        "routes": [
            ("51", "キリスト教大学 / 調布駅北口 / 調布飛行場 / 西野"),
            ("52", "大沢 / 大沢十字路"),
        ],
        "marks": [
            ("キ", "キリスト教大学行き"),
            ("調", "調布駅北口行き"),
            ("飛", "調布飛行場行き"),
            ("大", "大沢行き"),
            ("十", "大沢十字路行き"),
            ("西", "西野行き"),
        ],
    },
    {
        "no": "6",
        "header": "榊原病院前 ・ 朝日町三丁目 ・ 車返団地 ・ 多磨駅 ・ 60 循環 ・ 吉祥寺駅",
        "routes": [
            ("52", "榊原病院前 / 朝日町三丁目 / 車返団地 / 多磨駅"),
            ("60", "三鷹池循環(59と同じ循環) / 吉祥寺駅"),
        ],
        "marks": [
            ("榊", "榊原病院前行き"),
            ("朝三", "朝日町三丁目行き"),
            ("車返", "車返団地行き"),
            ("多", "多磨駅行き"),
            ("吉", "吉祥寺(駅)行き"),
            ("循環", "三鷹池循環(60/59)"),
        ],
    },
    {
        "no": "7",
        "header": "仙川 ・ 新川団地 ・ 61 調布 ・ 晃華学園東 ・ 67 三鷹駅 ・ 68 三鷹駅",
        "routes": [
            ("54", "仙川 / 新川団地"),
            ("61", "調布駅北口"),
            ("62", "晃華学園東"),
            ("67", "六中先経由 三鷹駅"),
            ("68", "五中先経由 三鷹駅 ※本人説明の「65」はこれの可能性"),
            ("07", "吉祥寺(6番系統の吉祥寺)"),
        ],
        "marks": [
            ("無印", "仙川行き"),
            ("新", "(南浦経由)新川団地行き"),
            ("調", "(南浦経由)調布行き"),
            ("晃", "晃華学園東行き"),
            ("吉07", "吉祥寺行き"),
            ("六", "六中先経由 三鷹駅行き"),
            ("五", "五中先経由 三鷹駅行き"),
        ],
    },
    {
        "no": "8",
        "header": "野ヶ谷 ・ 杏林大学 井の頭キャンパス ・ 牟礼団地 ・ 59 循環",
        "routes": [
            ("55", "野ヶ谷"),
            ("63", "杏林大学 井の頭キャンパス"),
            ("59", "三鷹駅循環"),
            ("58", "(牟礼団地 ほか) ※要確認"),
        ],
        "marks": [
            ("無印", "野ヶ谷行き"),
            ("牟", "牟礼団地行き"),
            ("杏", "杏林大学 井の頭キャンパス行き"),
            ("循環", "59 循環(三鷹駅)"),
        ],
    },
]

# 全のりば共通の色・記号凡例
COLOR_LEGEND = [
    ("黒字（無印）", "通常の発車", None, NORM),
    ("黄色セル / 鷹XX", "前運用あり（直前に走っていた路線＝三鷹+系統番号）", YELLOW, BOLD),
    ("赤字", "出庫（前運用なし。営業所などから出てくる便）", None, RED),
    ("青字", "終点から回送", None, BLUE),
    ("オレンジ", "出庫系 ※色の意味は要確認", ORANGE, BOLD),
]

HOURS = list(range(6, 25))   # 6〜24時
DEP_COLS = 14                # 発車記入欄の数（1行あたり最大の便数の目安）

wb = openpyxl.Workbook()

# ============ シート1: 凡例・読み方 ============
ws = wb.active
ws.title = "凡例・読み方"
ws.sheet_view.showGridLines = False
ws["A1"] = "三鷹駅南口発　時刻表（平日）　凡例・読み方"
ws["A1"].font = TITLE
ws["A2"] = "26.4.1 改正 / 構造版（時刻は別途記入）"
ws["A2"].font = SMALL

r = 4
ws.cell(r, 1, "■ 色・記号の意味").font = BOLD
r += 1
for h, c in [("表示", 1), ("意味", 2)]:
    cell = ws.cell(r, c, h)
    cell.font = HEAD; cell.fill = HEAD_FILL; cell.alignment = CENTER; cell.border = BORDER
r += 1
for label, meaning, fill, font in COLOR_LEGEND:
    c1 = ws.cell(r, 1, label); c1.font = font; c1.alignment = LEFT; c1.border = BORDER
    if fill: c1.fill = fill
    c2 = ws.cell(r, 2, meaning); c2.font = NORM; c2.alignment = LEFT; c2.border = BORDER
    r += 1

r += 1
ws.cell(r, 1, "■ 行き先の略号（共通）").font = BOLD
r += 1
common_marks = [
    ("無印", "そののりばの主行き先（のりば一覧シート参照）"),
    ("営", "営業所（武蔵境営業所）行き"),
    ("新", "新小金井駅行き"),
    ("植", "神代植物公園(前)行き"),
    ("循環", "循環便（59 / 60）"),
]
for label, meaning in common_marks:
    c1 = ws.cell(r, 1, label); c1.font = BOLD; c1.alignment = LEFT; c1.border = BORDER
    c2 = ws.cell(r, 2, meaning); c2.font = NORM; c2.alignment = LEFT; c2.border = BORDER
    r += 1

r += 1
ws.cell(r, 1, "■ 記入のしかた（各のりばシート）").font = BOLD
r += 1
notes = [
    "左端の列が「時」（6〜24時）。右へ向かって、その時間帯の発車を早い順に書く。",
    "1セル＝1便。書き方の目安：「分 鷹56」「分 鷹53新」「分 鷹56植」など（分＋系統＋行き先記号）。",
    "前運用がある便（写真で黄色）は、セルを黄色にして「(前運用 例: 鷹52)」を添える。",
    "出庫便（写真で赤字）は文字色を赤に。終点から回送（青字）は文字色を青に。",
    "便の上に小さく書いてある数字（前運用便番号など）は、必要なら同じセルに ( ) で添える。",
]
for n in notes:
    ws.cell(r, 1, "・" + n).font = NORM
    ws.cell(r, 1).alignment = LEFT
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r += 1

ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 60

# ============ シート2: のりば一覧 ============
ws2 = wb.create_sheet("のりば一覧")
ws2.sheet_view.showGridLines = False
ws2["A1"] = "のりば → 系統 → 行き先 対応表（三鷹駅南口発・平日）"
ws2["A1"].font = TITLE
r = 3
headers = ["のりば", "系統", "行き先", "備考"]
for c, h in enumerate(headers, start=1):
    cell = ws2.cell(r, c, h); cell.font = HEAD; cell.fill = HEAD_FILL
    cell.alignment = CENTER; cell.border = BORDER
r += 1
for p in PLATFORMS:
    start = r
    for sys, dest in p["routes"]:
        ws2.cell(r, 1, p["no"] + "番").font = BOLD
        ws2.cell(r, 1).alignment = CENTER; ws2.cell(r, 1).border = BORDER
        ws2.cell(r, 1).fill = HOUR_FILL
        ws2.cell(r, 2, sys).font = BOLD; ws2.cell(r, 2).alignment = CENTER; ws2.cell(r, 2).border = BORDER
        ws2.cell(r, 3, dest).font = NORM; ws2.cell(r, 3).alignment = LEFT; ws2.cell(r, 3).border = BORDER
        ws2.cell(r, 4, "").border = BORDER
        r += 1
    if r - start > 1:
        ws2.merge_cells(start_row=start, start_column=1, end_row=r-1, end_column=1)
ws2.column_dimensions["A"].width = 8
ws2.column_dimensions["B"].width = 8
ws2.column_dimensions["C"].width = 50
ws2.column_dimensions["D"].width = 30

# ============ シート3〜: のりば別 時刻表（枠） ============
for p in PLATFORMS:
    ws = wb.create_sheet(f"{p['no']}番のりば")
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"三鷹駅南口発　平日　{p['no']}番のりば"
    ws["A1"].font = TITLE
    ws["A2"] = f"〈 {p['header']} 〉"
    ws["A2"].font = BOLD
    ws["A3"] = "26.4.1 改正 / ★時刻は写真を見ながら記入する枠です（未入力）"
    ws["A3"].font = SMALL

    # 行き先凡例（そののりば）
    r = 5
    ws.cell(r, 1, "■ このりばの行き先記号").font = BOLD
    r += 1
    for mark, meaning in p["marks"]:
        ws.cell(r, 1, mark).font = BOLD; ws.cell(r, 1).alignment = CENTER; ws.cell(r, 1).border = BORDER
        ws.cell(r, 1).fill = GREY
        ws.cell(r, 2, meaning).font = NORM; ws.cell(r, 2).alignment = LEFT; ws.cell(r, 2).border = BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        r += 1

    # 時刻表ヘッダ
    r += 1
    table_top = r
    ws.cell(r, 1, "時").font = HEAD; ws.cell(r, 1).fill = HEAD_FILL
    ws.cell(r, 1).alignment = CENTER; ws.cell(r, 1).border = BORDER
    for c in range(2, 2 + DEP_COLS):
        cell = ws.cell(r, c, f"{c-1}")
        cell.font = HEAD; cell.fill = HEAD_FILL; cell.alignment = CENTER; cell.border = BORDER
    r += 1

    # 時の行
    for h in HOURS:
        hc = ws.cell(r, 1, h)
        hc.font = BOLD; hc.fill = HOUR_FILL; hc.alignment = CENTER; hc.border = BORDER
        for c in range(2, 2 + DEP_COLS):
            cell = ws.cell(r, c)
            cell.alignment = CENTER; cell.border = BORDER
            cell.font = NORM
        r += 1

    ws.column_dimensions["A"].width = 6
    for c in range(2, 2 + DEP_COLS):
        ws.column_dimensions[get_column_letter(c)].width = 9
    ws.freeze_panes = ws.cell(table_top + 1, 2)

out = "三鷹駅南口_平日_時刻表.xlsx"
wb.save(out)
print("saved:", out)
