# -*- coding: utf-8 -*-
"""
AI人事評価システム Web版
株式会社栄工舎 製造第1課
"""
import streamlit as st
import json, ssl, urllib.request, urllib.error
from datetime import datetime

from constants import (GRADE_LEVELS, GRADE_EVAL_ITEMS, PG_POINTS_BY_GRADE,
                       GRADE_THRESHOLDS, calc_final_grade)
from database import (get_all_employees, get_employee, upsert_employee,
                      delete_employee, get_evaluations, get_evaluation,
                      upsert_evaluation, delete_evaluation, parse_json_cols,
                      get_daily_records, insert_daily_record, delete_daily_record,
                      get_import_history, delete_import_history)
from excel_export import export_excel

# ============================================================
# ページ設定
# ============================================================
st.set_page_config(
    page_title="AI人事評価システム | 栄工舎 製造第1課",
    page_icon="🏭",
    layout="wide",
)

# ============================================================
# パスワード認証
# ============================================================
def check_password():
    if st.session_state.get("authenticated"):
        return True
    st.title("🔒 AI人事評価システム")
    st.subheader("株式会社栄工舎 製造第1課")
    pw = st.text_input("パスワードを入力してください", type="password")
    if st.button("ログイン"):
        correct = st.secrets.get("app_password", "eikosya2024")
        if pw == correct:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False

if not check_password():
    st.stop()

# ============================================================
# スタイル
# ============================================================
st.markdown("""
<style>
.grade-badge {
    display:inline-block;padding:4px 12px;border-radius:4px;
    font-weight:bold;font-size:1.1em;
}
.grade-A{background:#EAF3DE;color:#27500A;}
.grade-B{background:#E6F1FB;color:#0C447C;}
.grade-C{background:#FAEEDA;color:#633806;}
.grade-D{background:#FCEBEB;color:#791F1F;}
.grade-E{background:#D3D1C7;color:#444441;}
.cat-header{background:#F0F0F0;padding:4px 8px;border-left:4px solid #666;
            font-weight:bold;margin:8px 0 4px 0;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# サイドバー
# ============================================================
st.sidebar.title("🏭 人事評価システム")
st.sidebar.markdown("**栄工舎 製造第1課**")
st.sidebar.markdown("---")
page = st.sidebar.radio(
    "メニュー",
    ["📝 評価入力", "📋 評価一覧", "👥 社員管理", "📊 統計", "📅 日次記録", "📂 取込履歴"],
    label_visibility="collapsed"
)
st.sidebar.markdown("---")
if st.sidebar.button("🚪 ログアウト"):
    st.session_state.authenticated = False
    st.rerun()

# ============================================================
# ヘルパー
# ============================================================
def grade_badge(g):
    return f'<span class="grade-badge grade-{g}">{g}</span>'

def ai_score(emp_name, grade_level, pg_goal, items, obs_text):
    """Claude APIで採点"""
    api_key = st.secrets.get("anthropic_api_key", "")
    if not api_key:
        return None, "APIキーが設定されていません"

    item_list = "\n".join([f"- {iname}（最大{max_pts}点）: {obs}"
                           for iid,cat,iname,max_pts,gpts,obs in items])
    pg_pts_dict = PG_POINTS_BY_GRADE[grade_level]
    item_grades_json = "{" + ",".join([f'"{iid}":"A/B/C/D/E"' for iid,*_ in items]) + "}"

    prompt = f"""あなたは製造業の人事評価の専門家です。
以下の従業員を評価してください。

従業員名: {emp_name}
等級: {grade_level}
個人目標: {pg_goal or '（未記入）'}
観察・備考: {obs_text or '（なし）'}

【評価項目】
{item_list}

以下のJSON形式のみで回答してください（説明不要）:
{{
  "pg_grade": "A/B/C/D/E のいずれか",
  "item_grades": {item_grades_json},
  "comment": "総合コメント（200字以内）"
}}"""

    body = json.dumps({
        "model": "claude-sonnet-4-6",
        "max_tokens": 1024,
        "messages": [{"role": "user", "content": prompt}]
    }, ensure_ascii=False).encode()

    ctx = ssl.create_default_context()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
    )
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
            resp = json.loads(r.read())
        text = resp["content"][0]["text"].strip()
        # JSONブロック抽出
        if "```" in text:
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        data = json.loads(text)
        return data, None
    except Exception as e:
        return None, str(e)

def calc_total(grade_level, pg_grade, item_grades_dict, adj=0):
    pg_pts = PG_POINTS_BY_GRADE[grade_level].get(pg_grade, 0)
    items = GRADE_EVAL_ITEMS[grade_level]
    item_total = sum(item[4].get(item_grades_dict.get(item[0], "E"), 0) for item in items)
    total = pg_pts + item_total + adj
    return pg_pts, item_total, total

# ============================================================
# ページ: 評価入力
# ============================================================
def page_eval_input():
    st.title("📝 人事評価入力")

    employees = get_all_employees()
    if not employees:
        st.warning("社員が登録されていません。先に「社員管理」から登録してください。")
        return

    emp_options = {f"{e['emp_id']} {e['name']}": e for e in employees}
    sel = st.selectbox("社員を選択", list(emp_options.keys()))
    emp = emp_options[sel]

    col1, col2, col3 = st.columns(3)
    with col1:
        year = st.number_input("年度", value=datetime.now().year, min_value=2000, max_value=2100)
    with col2:
        period = st.selectbox("期", ["上期", "下期"])
    with col3:
        grade_level = st.selectbox("等級", GRADE_LEVELS,
                                   index=GRADE_LEVELS.index(emp.get("grade_level","A2")))

    items = GRADE_EVAL_ITEMS[grade_level]
    pg_pts_map = PG_POINTS_BY_GRADE[grade_level]
    grade_opts = [f"{g}（{v}点）" for g, v in pg_pts_map.items()]
    grade_keys = list(pg_pts_map.keys())

    st.markdown("---")
    st.subheader("個人目標")
    pg_goal = st.text_area("個人目標の内容", height=80, key="pg_goal")
    pg_grade_label = st.selectbox("個人目標 評価グレード", grade_opts, key="pg_grade")
    pg_grade = grade_keys[grade_opts.index(pg_grade_label)]

    st.markdown("---")
    st.subheader("評価項目")

    obs_lines = st.text_area(
        "観察・備考（AI採点の参考にします）",
        height=100,
        placeholder="例：作業ミスが○回あった。改善提案を△件出した。欠勤○日。",
        key="obs_text"
    )

    # グレード選択
    item_grade_selections = {}
    cur_cat = ""
    for iid, cat, iname, max_pts, gpts, obs in items:
        if cat != cur_cat:
            st.markdown(f'<div class="cat-header">📌 {cat}</div>', unsafe_allow_html=True)
            cur_cat = cat
        c1, c2, c3 = st.columns([3, 2, 2])
        with c1:
            st.markdown(f"**{iname}**（最大{max_pts}点）")
            st.caption(obs)
        with c2:
            opts = [f"{g}（{v}点）" for g, v in gpts.items()]
            chosen = st.selectbox("", opts, key=f"ig_{iid}", label_visibility="collapsed")
            item_grade_selections[iid] = list(gpts.keys())[opts.index(chosen)]
        with c3:
            pts = gpts[item_grade_selections[iid]]
            st.metric("得点", f"{pts}点")

    # 調整点数
    st.markdown("---")
    adj = st.number_input("調整点数（±）", value=0, min_value=-20, max_value=20)

    # 合計計算
    pg_pts_val, item_total, total = calc_total(grade_level, pg_grade, item_grade_selections, adj)
    final_g = calc_final_grade(total)

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("個人目標", f"{pg_pts_val}点")
    col_b.metric("項目合計", f"{item_total}点")
    col_c.metric("総合計", f"{total}点")
    col_d.markdown(f"**最終評価**<br>{grade_badge(final_g)}", unsafe_allow_html=True)

    st.markdown("---")
    col_e, col_f = st.columns(2)
    evaluatee_comment = col_e.text_area("被評価者コメント", height=80)
    evaluator_comment = col_f.text_area("評価者コメント", height=80)
    evaluator1 = col_e.text_input("一次評価者")
    evaluator2 = col_f.text_input("二次評価者")

    # AI採点ボタン
    st.markdown("---")
    col_ai1, col_ai2, col_save = st.columns(3)
    with col_ai1:
        if st.button("🤖 AI採点", use_container_width=True):
            with st.spinner("AIが採点中..."):
                result, err = ai_score(
                    emp["name"], grade_level, pg_goal, items, obs_lines
                )
            if err:
                st.error(f"AI採点エラー: {err}")
            else:
                pg_g = result.get("pg_grade", "C")
                ig_map = result.get("item_grades", {})
                comment = result.get("comment", "")
                st.session_state["ai_result"] = {
                    "pg_grade": pg_g, "item_grades": ig_map, "comment": comment
                }
                st.success(f"AI採点完了！コメント: {comment}")
                st.info("「AI採点を反映して保存」ボタンで保存できます。")

    with col_ai2:
        if "ai_result" in st.session_state and st.button("💾 AI採点を反映して保存", use_container_width=True):
            ai = st.session_state["ai_result"]
            ai_pg = ai["pg_grade"]
            ai_ig = ai["item_grades"]
            ai_pg_pts = pg_pts_map.get(ai_pg, 0)
            ai_item_pts = {iid: GRADE_EVAL_ITEMS[grade_level][i][4].get(ai_ig.get(iid,"E"),0)
                           for i, (iid,*_) in enumerate(items)}
            ai_total = ai_pg_pts + sum(ai_item_pts.values()) + adj
            ai_final = calc_final_grade(ai_total)
            ev = {
                "emp_id": emp["emp_id"],
                "year": int(year), "period": period,
                "grade_level": grade_level,
                "pg_grade": ai_pg, "pg_points": ai_pg_pts, "pg_goal": pg_goal,
                "item_grades": ai_ig, "item_points": ai_item_pts,
                "adj_points": adj, "total_score": ai_total, "final_grade": ai_final,
                "evaluator1": evaluator1, "evaluator2": evaluator2,
                "evaluatee_comment": evaluatee_comment,
                "evaluator_comment": evaluator_comment,
                "ai_comment": ai["comment"],
            }
            upsert_evaluation(ev)
            st.success("AI採点結果を保存しました！")
            del st.session_state["ai_result"]

    with col_save:
        if st.button("💾 手動評価を保存", use_container_width=True):
            item_pts = {iid: GRADE_EVAL_ITEMS[grade_level][i][4].get(item_grade_selections.get(iid,"E"),0)
                        for i, (iid,*_) in enumerate(items)}
            ev = {
                "emp_id": emp["emp_id"],
                "year": int(year), "period": period,
                "grade_level": grade_level,
                "pg_grade": pg_grade, "pg_points": pg_pts_val, "pg_goal": pg_goal,
                "item_grades": item_grade_selections, "item_points": item_pts,
                "adj_points": adj, "total_score": total, "final_grade": final_g,
                "evaluator1": evaluator1, "evaluator2": evaluator2,
                "evaluatee_comment": evaluatee_comment,
                "evaluator_comment": evaluator_comment,
                "ai_comment": "",
            }
            upsert_evaluation(ev)
            st.success("評価を保存しました！")

# ============================================================
# ページ: 評価一覧
# ============================================================
def page_eval_list():
    st.title("📋 評価一覧")

    employees = get_all_employees()
    options = ["全員"] + [f"{e['emp_id']} {e['name']}" for e in employees]
    filter_emp = st.selectbox("絞り込み", options)
    emp_id_filter = None
    if filter_emp != "全員":
        emp_id_filter = filter_emp.split()[0]

    evals = get_evaluations(emp_id_filter)
    if not evals:
        st.info("評価データがありません。")
        return

    emp_map = {e["emp_id"]: e["name"] for e in employees}

    for ev in evals:
        ev = parse_json_cols(ev)
        name = emp_map.get(ev["emp_id"], ev["emp_id"])
        with st.expander(
            f"**{name}** | {ev['year']}年度 {ev['period']} | "
            f"等級{ev['grade_level']} | 総合計{ev['total_score']}点 | "
            f"最終評価：{ev['final_grade']}"
        ):
            c1, c2 = st.columns(2)
            c1.markdown(f"**個人目標グレード:** {ev['pg_grade']}（{ev['pg_points']}点）")
            c1.markdown(f"**個人目標:** {ev.get('pg_goal','')}")
            c2.markdown(f"**調整点数:** {ev.get('adj_points',0)}点")
            c2.markdown(f"**AIコメント:** {ev.get('ai_comment','')}")

            # 評価項目一覧
            items = GRADE_EVAL_ITEMS.get(ev["grade_level"], [])
            rows = []
            for iid,cat,iname,max_pts,gpts,obs in items:
                ig = ev["item_grades"].get(iid, "-")
                ip = ev["item_points"].get(iid, 0)
                rows.append({"カテゴリ":cat,"項目":iname,"グレード":ig,"得点":ip,"最大":max_pts})
            if rows:
                import pandas as pd
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            # 削除 / Excel出力
            col_del, col_xl = st.columns(2)
            with col_del:
                if st.button("🗑️ 削除", key=f"del_{ev['id']}"):
                    delete_evaluation(ev["id"])
                    st.success("削除しました")
                    st.rerun()
            with col_xl:
                try:
                    emp_data = get_employee(ev["emp_id"]) or {}
                    xl_data = {**ev, "dept": emp_data.get("dept","製造第1課"),
                               "name": emp_data.get("name", ev["emp_id"])}
                    xl_bytes = export_excel(xl_data)
                    filename = f"人事考課_{name}_{ev['year']}_{ev['period']}.xlsx"
                    st.download_button(
                        "📥 Excelダウンロード", xl_bytes, filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"xl_{ev['id']}"
                    )
                except Exception as e:
                    st.error(f"Excel出力エラー: {e}")

# ============================================================
# ページ: 社員管理
# ============================================================
def page_employees():
    st.title("👥 社員管理")

    tab_list, tab_add = st.tabs(["社員一覧", "社員登録・編集"])

    with tab_list:
        employees = get_all_employees()
        if not employees:
            st.info("社員が登録されていません。")
        for emp in employees:
            with st.expander(f"**{emp['emp_id']}** {emp['name']}（等級：{emp['grade_level']}）"):
                st.write(f"部署: {emp.get('dept','')}")
                st.write(f"入社日: {emp.get('hire_date','')}")
                st.write(f"備考: {emp.get('notes','')}")
                if st.button("🗑️ 削除", key=f"demp_{emp['emp_id']}"):
                    delete_employee(emp["emp_id"])
                    st.success("削除しました（評価データも削除）")
                    st.rerun()

    with tab_add:
        st.subheader("社員登録・編集")
        employees = get_all_employees()
        edit_opts = ["新規登録"] + [f"{e['emp_id']} {e['name']}" for e in employees]
        edit_sel = st.selectbox("登録 or 編集する社員", edit_opts)

        if edit_sel == "新規登録":
            defaults = {"emp_id":"","name":"","dept":"製造第1課","grade_level":"A2","hire_date":"","notes":""}
        else:
            eid = edit_sel.split()[0]
            defaults = get_employee(eid) or {}

        with st.form("emp_form"):
            emp_id   = st.text_input("社員番号 *", defaults.get("emp_id",""))
            name     = st.text_input("氏名 *", defaults.get("name",""))
            dept     = st.text_input("部署", defaults.get("dept","製造第1課"))
            gl       = st.selectbox("等級", GRADE_LEVELS,
                                    index=GRADE_LEVELS.index(defaults.get("grade_level","A2")))
            hire_dt  = st.text_input("入社日（例：2018-04-01）", defaults.get("hire_date",""))
            notes    = st.text_area("備考", defaults.get("notes",""))
            submitted = st.form_submit_button("💾 保存")
            if submitted:
                if not emp_id or not name:
                    st.error("社員番号と氏名は必須です")
                else:
                    upsert_employee({
                        "emp_id": emp_id, "name": name, "dept": dept,
                        "grade_level": gl, "hire_date": hire_dt, "notes": notes
                    })
                    st.success(f"「{name}」を保存しました")
                    st.rerun()

# ============================================================
# ページ: 統計
# ============================================================
def page_stats():
    st.title("📊 統計・集計")
    evals = get_evaluations()
    if not evals:
        st.info("評価データがありません。")
        return

    import pandas as pd
    employees = get_all_employees()
    emp_map = {e["emp_id"]: e["name"] for e in employees}

    rows = []
    for ev in evals:
        rows.append({
            "氏名": emp_map.get(ev["emp_id"], ev["emp_id"]),
            "年度": ev["year"],
            "期": ev["period"],
            "等級": ev["grade_level"],
            "総合計": ev["total_score"],
            "最終評価": ev["final_grade"],
        })
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("最終評価の分布")
    grade_counts = df["最終評価"].value_counts().reindex(["A","B","C","D","E"], fill_value=0)
    st.bar_chart(grade_counts)

# ============================================================
# ページ: 日次記録
# ============================================================
def page_daily_records():
    st.title("📅 日次記録")

    employees = get_all_employees()
    emp_map = {e["emp_id"]: e["name"] for e in employees}

    tab_list, tab_add = st.tabs(["記録一覧", "新規登録"])

    with tab_list:
        # 絞り込み
        options = ["全員"] + [f"{e['emp_id']} {e['name']}" for e in employees]
        filter_sel = st.selectbox("社員で絞り込み", options, key="dr_filter")
        emp_id_filter = None
        if filter_sel != "全員":
            emp_id_filter = filter_sel.split()[0]

        records = get_daily_records(emp_id_filter)
        if not records:
            st.info("日次記録がありません。")
        else:
            st.write(f"**{len(records)} 件**")
            import pandas as pd
            rows = []
            for r in records:
                rows.append({
                    "ID": r["id"],
                    "日付": r.get("record_date", ""),
                    "社員名": emp_map.get(str(r.get("emp_id","")), r.get("emp_name","")),
                    "カテゴリ": r.get("category", ""),
                    "記録": r.get("record_text", "")[:60] + ("..." if len(r.get("record_text","")) > 60 else ""),
                })
            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("詳細・削除")
            for r in records:
                name = emp_map.get(str(r.get("emp_id","")), r.get("emp_name",""))
                with st.expander(f"{r.get('record_date','')} | {name} | {r.get('category','')}"):
                    st.write(r.get("record_text",""))
                    if st.button("🗑️ 削除", key=f"dr_del_{r['id']}"):
                        delete_daily_record(r["id"])
                        st.success("削除しました")
                        st.rerun()

    with tab_add:
        st.subheader("日次記録 新規登録")
        with st.form("dr_form"):
            emp_opts = {f"{e['emp_id']} {e['name']}": e["emp_id"] for e in employees}
            sel_emp = st.selectbox("社員", list(emp_opts.keys()))
            record_date = st.date_input("日付")
            category = st.text_input("カテゴリ（例：作業記録、改善提案）")
            record_text = st.text_area("記録内容", height=120)
            submitted = st.form_submit_button("💾 保存")
            if submitted:
                eid = emp_opts[sel_emp]
                ename = emp_map.get(str(eid), "")
                insert_daily_record({
                    "emp_id": str(eid),
                    "emp_name": ename,
                    "record_date": str(record_date),
                    "category": category,
                    "record_text": record_text,
                })
                st.success("記録を保存しました！")
                st.rerun()

# ============================================================
# ページ: 取込履歴
# ============================================================
def page_import_history():
    st.title("📂 取込履歴")

    records = get_import_history()
    if not records:
        st.info("取込履歴がありません。")
        return

    st.write(f"**{len(records)} 件**")

    import pandas as pd
    rows = []
    for r in records:
        rows.append({
            "ID": r["id"],
            "取込日時": r.get("import_date", ""),
            "ファイル名": r.get("file_name", ""),
            "種別": r.get("file_type", ""),
            "件数": r.get("affected_count", 0),
            "状態": r.get("status", ""),
            "詳細": r.get("detail", "")[:50] + ("..." if len(r.get("detail","")) > 50 else ""),
        })
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("詳細")
    for r in records:
        label = f"{r.get('import_date','')} | {r.get('file_name','')} | {r.get('status','')}"
        with st.expander(label):
            col1, col2 = st.columns(2)
            col1.write(f"**ファイル種別:** {r.get('file_type','')}")
            col1.write(f"**件数:** {r.get('affected_count',0)}")
            col2.write(f"**詳細:** {r.get('detail','')}")
            if st.button("🗑️ 削除", key=f"ih_del_{r['id']}"):
                delete_import_history(r["id"])
                st.success("削除しました")
                st.rerun()

# ============================================================
# ルーティング
# ============================================================
if "📝 評価入力" in page:
    page_eval_input()
elif "📋 評価一覧" in page:
    page_eval_list()
elif "👥 社員管理" in page:
    page_employees()
elif "📊 統計" in page:
    page_stats()
elif "📅 日次記録" in page:
    page_daily_records()
elif "📂 取込履歴" in page:
    page_import_history()
