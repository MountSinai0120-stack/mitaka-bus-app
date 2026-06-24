# -*- coding: utf-8 -*-
"""
2番のりば の時刻を記入(改訂版)。
鮮明な写真 IMG_1799/1800/1801/1802 で読み直し、6〜10/12/14時の確信度を上げた。
各タプル: (分, 前運用route, uncertain[, dest])
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

PATH = "三鷹駅南口_平日_時刻表.xlsx"
NORM = Font(name="Yu Gothic", size=10)
YELLOW = PatternFill("solid", fgColor="FFF2A8")
ORANGE = PatternFill("solid", fgColor="FCD9A8")
NOFILL = PatternFill(fill_type=None)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

DATA = {
 6: [(15,"",False),(28,"52",False),(38,"63",False),(48,"51",False),(56,"51",False)],
 7: [(4,"",False),(8,"63",False),(13,"55",False),(22,"65",False),(32,"51",False),(38,"51",False),(50,"60",False)],
 8: [(3,"51",False),(10,"60",False),(15,"65",False),(34,"55",False),(45,"54",False),(55,"62",True)],
 9: [(0,"63",False),(15,"59",False),(30,"59",True),(45,"63",False),(50,"63",False)],
 10:[(0,"",False),(15,"51",False),(30,"57",False),(45,"51",True)],
 11:[(0,"",True),(8,"54",False),(23,"51",True),(30,"",True),(45,"52",True),(52,"63",True),(59,"63",True)],
 12:[(8,"51",False),(13,"",False),(20,"51",False),(27,"",False),(35,"55",False),(42,"",False),(50,"61",False),(57,"",False)],
 13:[(0,"55",True),(12,"65",False),(20,"",True),(30,"63",False),(42,"",True),(50,"52",False),(57,"62",True)],
 14:[(9,"55",False),(12,"",False),(20,"52",False),(27,"62",False),(35,"",False),(42,"62",False),(57,"",True)],
 15:[(5,"51",True),(12,"",True),(20,"54",True),(27,"",True),(35,"",True),(42,"54",True),(50,"",True)],
 16:[(0,"",True),(12,"51",True),(22,"56",True),(30,"",True),(38,"",True),(48,"",True),(55,"62",True)],
 17:[(0,"56",True),(14,"",True),(28,"",True),(40,"",True),(46,"",True),(52,"",True),(57,"62",True)],
 18:[(0,"",True),(6,"53",True),(14,"",True),(20,"",True),(36,"52",True),(50,"52",True)],
 19:[(5,"",True),(11,"63",True),(23,"55",True),(30,"",True),(43,"",True),(48,"62",True),(55,"63",True)],
 20:[(0,"",True),(9,"51",True),(20,"",True),(37,"",True),(48,"54",True)],
 21:[(4,"",True),(14,"54",True),(23,"",True),(31,"56",True),(39,"",True),(53,"54",True)],
 22:[(10,"55",True),(20,"",True),(32,"",True)],
 24:[(9,"54",True,"植")],
}

wb = openpyxl.load_workbook(PATH)
ws = wb["2番のりば"]
HOUR6 = 12  # 2番: マーク4個 → 6時=行12

# まず時刻欄(列2〜15, 行12〜30)をクリア
for r in range(HOUR6, HOUR6 + 19):
    for c in range(2, 16):
        cell = ws.cell(r, c)
        cell.value = None; cell.fill = NOFILL

for h, deps in DATA.items():
    r = HOUR6 + (h - 6)
    for i, rec in enumerate(deps):
        mm, route, unc = rec[0], rec[1], rec[2]
        dest = rec[3] if len(rec) > 3 else ""
        text = f"{mm:02d}"
        if route: text += f"\n鷹{route}"
        if dest: text += dest
        cell = ws.cell(r, 2 + i, text)
        cell.font = NORM; cell.alignment = CENTER
        if unc: cell.fill = ORANGE
        elif route: cell.fill = YELLOW
    ws.row_dimensions[r].height = 30

wb.save(PATH)
print("2番のりば 改訂完了")
