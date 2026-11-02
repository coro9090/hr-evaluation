# -*- coding: utf-8 -*-
"""
公式Excelテンプレート出力モジュール
"""
import io, os
import openpyxl
from constants import GRADE_EVAL_ITEMS, TEMPLATE_ITEM_ROWS, PG_POINTS_BY_GRADE, calc_final_grade

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "templates", "人事考課フォーム_template.xlsx")

def export_excel(d: dict) -> bytes:
    """
    評価データdictからExcelバイト列を生成して返す。
    d keys: year, period, name, dept, grade_level, pg_grade, pg_points,
            pg_goal, item_grades, item_points, adj_points, total_score,
            final_grade, evaluator1, evaluator2,
            evaluatee_comment, evaluator_comment
    """
    grade = d.get("grade_level", "A2").upper()
    if grade not in TEMPLATE_ITEM_ROWS:
        raise ValueError(f"等級 {grade} はテンプレートに対応していません")

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    sheet_name = f"人事考課フォーム（{grade}）"
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"シート「{sheet_name}」が見つかりません")

    ws = wb[sheet_name]
    cfg = TEMPLATE_ITEM_ROWS[grade]
    items = GRADE_EVAL_ITEMS[grade]
    pg_pts = PG_POINTS_BY_GRADE[grade]

    # ヘッダー情報
    ws.cell(row=2, column=1).value  = d.get("year", "")         # A2: 年度
    ws.cell(row=2, column=3).value  = d.get("period", "上期")    # C2: 期
    grade_num = {"A5":5,"A4":4,"A3":3,"A2":2,"A1":1,"M3":3,"M2":2,"M1":1}.get(grade, "")
    ws.cell(row=2, column=8).value  = grade_num                   # H2: 等級番号
    ws.cell(row=3, column=10).value = d.get("name", "")           # J3: 氏名
    ws.cell(row=3, column=4).value  = d.get("dept", "製造第1課")  # D3: 部署
    ws.cell(row=2, column=17).value = d.get("evaluator1", "")     # Q2: 一次評価者
    ws.cell(row=2, column=19).value = d.get("evaluator2", "")     # S2: 二次評価者

    # 個人目標
    pg_row = cfg["pg"]
    pg_grade = d.get("pg_grade", "")
    pg_score = d.get("pg_points", 0)
    grade_cols = {"A": 13, "B": 14, "C": 15, "D": 16, "E": 17}
    if pg_grade in grade_cols:
        ws.cell(row=pg_row, column=grade_cols[pg_grade]).value = "○"
    ws.cell(row=pg_row, column=21).value = pg_score

    # 各評価項目
    item_grades = d.get("item_grades", {})
    item_points = d.get("item_points", {})
    for iid, cat, iname, max_pts, gpts, obs in items:
        row = cfg["items"].get(iid)
        if not row:
            continue
        ig = item_grades.get(iid, "")
        ip = item_points.get(iid, 0)
        if ig in grade_cols:
            ws.cell(row=row, column=grade_cols[ig]).value = "○"
        ws.cell(row=row, column=21).value = ip

    # 合計・最終評価
    total_row = cfg["total_row"]
    ws.cell(row=total_row - 1, column=20).value = d.get("adj_points", 0)
    ws.cell(row=total_row, column=21).value = d.get("total_score", 0)
    ws.cell(row=cfg["grade_row"], column=21).value = d.get("final_grade", "")

    # コメント
    comment_row = cfg["comment_row"]
    ws.cell(row=comment_row, column=1).value  = d.get("evaluatee_comment", "")
    ws.cell(row=comment_row, column=10).value = d.get("evaluator_comment", "")

    # 対象シート以外を削除
    for sn in list(wb.sheetnames):
        if sn != sheet_name:
            del wb[sn]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
