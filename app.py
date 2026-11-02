# -*- coding: utf-8 -*-
"""
AI人事評価システム Web版
株式会社栄工舎 製造第1課
"""
import streamlit as st
import json, ssl, urllib.request, urllib.error
from datetime import datetime
import io

from constants import (GRADE_LEVELS, GRADE_EVAL_ITEMS, PG_POINTS_BY_GRADE,
                       GRADE_THRESHOLDS, calc_final_grade)
from database import (get_all_employees, get_employee, upsert_employee,
                      delete_employee, get_evaluations, get_evaluation,
                      upsert_evaluation, delete_evaluation, parse_json_cols,
                      get_daily_records, insert_daily_record, delete_daily_record,
                      get_import_history, delete_import_history, insert_import_history)
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
    ["📝 評価入力", "📋 評価一覧", "👥 社員管理", "📊 統計",
     "📅 日次記録", "📂 取込履歴", "📥 一括取込"],
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
# ページ: 評価入力（既存データ自動ロード対応）
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
        # 読み込んだ等級があればそれを使う
        gl_default = st.session_state.get("_ev_grade_level") or emp.get("grade_level", "A2")
        if gl_default not in GRADE_LEVELS:
            gl_default = "A2"
        grade_level = st.selectbox("等級", GRADE_LEVELS,
                                   index=GRADE_LEVELS.index(gl_default))

    items = GRADE_EVAL_ITEMS[grade_level]
    pg_pts_map = PG_POINTS_BY_GRADE[grade_level]
    grade_opts = [f"{g}（{v}点）" for g, v in pg_pts_map.items()]
    grade_keys = list(pg_pts_map.keys())

    # ★ 既存評価データの自動ロード ★
    load_key = f"{emp['emp_id']}_{int(year)}_{period}"
    if st.session_state.get("_eval_load_key") != load_key:
        evals = get_evaluations(emp["emp_id"])
        existing = next(
            (parse_json_cols(e) for e in evals
             if e["year"] == int(year) and e["period"] == period),
            None
        )
        if existing:
            gl = existing.get("grade_level") or "A2"
            if gl not in GRADE_LEVELS:
                gl = "A2"
            loaded_pg_map = PG_POINTS_BY_GRADE.get(gl, pg_pts_map)
            loaded_grade_opts = [f"{g}（{v}点）" for g, v in loaded_pg_map.items()]
            loaded_items = GRADE_EVAL_ITEMS.get(gl, items)

            pg_g = existing.get("pg_grade") or "C"
            pg_label = f"{pg_g}（{loaded_pg_map.get(pg_g, 0)}点）"
            if pg_label not in loaded_grade_opts:
                pg_label = loaded_grade_opts[2] if len(loaded_grade_opts) > 2 else loaded_grade_opts[0]

            st.session_state["_ev_grade_level"] = gl
            st.session_state["pg_goal"] = existing.get("pg_goal") or ""
            st.session_state["pg_grade"] = pg_label
            for iid, cat, iname, max_pts, gpts, obs in loaded_items:
                ig = (existing.get("item_grades") or {}).get(iid, "E")
                opts_i = [f"{g}（{v}点）" for g, v in gpts.items()]
                ig_label = f"{ig}（{gpts.get(ig, 0)}点）"
                st.session_state[f"ig_{iid}"] = ig_label if ig_label in opts_i else opts_i[-1]
            st.session_state["adj_points"] = int(existing.get("adj_points") or 0)
            st.session_state["ev_comment"] = existing.get("evaluatee_comment") or ""
            st.session_state["er_comment"] = existing.get("evaluator_comment") or ""
            st.session_state["ev1"] = existing.get("evaluator1") or ""
            st.session_state["ev2"] = existing.get("evaluator2") or ""
            st.session_state["obs_text"] = ""
            st.session_state["_existing_eval_id"] = existing.get("id")
        else:
            st.session_state["_ev_grade_level"] = None
            st.session_state["pg_goal"] = ""
            st.session_state["pg_grade"] = grade_opts[2] if len(grade_opts) > 2 else grade_opts[0]
            for iid, cat, iname, max_pts, gpts, obs in items:
                opts_i = [f"{g}（{v}点）" for g, v in gpts.items()]
                st.session_state[f"ig_{iid}"] = opts_i[-1]
            st.session_state["adj_points"] = 0
            st.session_state["ev_comment"] = ""
            st.session_state["er_comment"] = ""
            st.session_state["ev1"] = ""
            st.session_state["ev2"] = ""
            st.session_state["obs_text"] = ""
            st.session_state["_existing_eval_id"] = None

        st.session_state["_eval_load_key"] = load_key
        st.rerun()

    # 読み込み状態の表示
    existing_id = st.session_state.get("_existing_eval_id")
    if existing_id:
        st.success(f"✅ 既存の評価データを読み込みました（更新モード）")
    else:
        st.info("📝 新規評価入力")

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
    adj = st.number_input("調整点数（±）", value=0, min_value=-20, max_value=20, key="adj_points")

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
    evaluatee_comment = col_e.text_area("被評価者コメント", height=80, key="ev_comment")
    evaluator_comment = col_f.text_area("評価者コメント", height=80, key="er_comment")
    evaluator1 = col_e.text_input("一次評価者", key="ev1")
    evaluator2 = col_f.text_input("二次評価者", key="ev2")

    def build_ev(ai_data=None):
        if ai_data:
            ai_pg = ai_data["pg_grade"]
            ai_ig = ai_data["item_grades"]
            ai_pg_pts = pg_pts_map.get(ai_pg, 0)
            ai_item_pts = {iid: GRADE_EVAL_ITEMS[grade_level][i][4].get(ai_ig.get(iid,"E"),0)
                           for i, (iid,*_) in enumerate(items)}
            ai_total = ai_pg_pts + sum(ai_item_pts.values()) + adj
            ev = {
                "emp_id": emp["emp_id"], "year": int(year), "period": period,
                "grade_level": grade_level,
                "pg_grade": ai_pg, "pg_points": ai_pg_pts, "pg_goal": pg_goal,
                "item_grades": ai_ig, "item_points": ai_item_pts,
                "adj_points": adj, "total_score": ai_total,
                "final_grade": calc_final_grade(ai_total),
                "evaluator1": evaluator1, "evaluator2": evaluator2,
                "evaluatee_comment": evaluatee_comment,
                "evaluator_comment": evaluator_comment,
                "ai_comment": ai_data.get("comment", ""),
            }
        else:
            item_pts = {iid: GRADE_EVAL_ITEMS[grade_level][i][4].get(item_grade_selections.get(iid,"E"),0)
                        for i, (iid,*_) in enumerate(items)}
            ev = {
                "emp_id": emp["emp_id"], "year": int(year), "period": period,
                "grade_level": grade_level,
                "pg_grade": pg_grade, "pg_points": pg_pts_val, "pg_goal": pg_goal,
                "item_grades": item_grade_selections, "item_points": item_pts,
                "adj_points": adj, "total_score": total, "final_grade": final_g,
                "evaluator1": evaluator1, "evaluator2": evaluator2,
                "evaluatee_comment": evaluatee_comment,
                "evaluator_comment": evaluator_comment,
                "ai_comment": "",
            }
        if st.session_state.get("_existing_eval_id"):
            ev["id"] = st.session_state["_existing_eval_id"]
        return ev

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
                st.session_state["ai_result"] = result
                st.success(f"AI採点完了！コメント: {result.get('comment','')}")
                st.info("「AI採点を反映して保存」ボタンで保存できます。")

    with col_ai2:
        if "ai_result" in st.session_state and st.button("💾 AI採点を反映して保存", use_container_width=True):
            ev = build_ev(ai_data=st.session_state["ai_result"])
            result = upsert_evaluation(ev)
            if result and not st.session_state.get("_existing_eval_id"):
                st.session_state["_existing_eval_id"] = result[0].get("id")
            st.success("AI採点結果を保存しました！")
            del st.session_state["ai_result"]

    with col_save:
        if st.button("💾 手動評価を保存", use_container_width=True):
            ev = build_ev()
            result = upsert_evaluation(ev)
            if result and not st.session_state.get("_existing_eval_id"):
                st.session_state["_existing_eval_id"] = result[0].get("id")
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

            items = GRADE_EVAL_ITEMS.get(ev["grade_level"], [])
            rows = []
            for iid,cat,iname,max_pts,gpts,obs in items:
                ig = ev["item_grades"].get(iid, "-")
                ip = ev["item_points"].get(iid, 0)
                rows.append({"カテゴリ":cat,"項目":iname,"グレード":ig,"得点":ip,"最大":max_pts})
            if rows:
                import pandas as pd
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

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
# ページ: 一括取込
# ============================================================
def page_bulk_import():
    st.title("📥 一括取込")

    import pandas as pd

    st.markdown("""
    CSVまたはExcelファイルから評価データを一括インポートします。

    **取込対象:** 評価データ（基本情報・個人目標・コメント）
    """)

    # テンプレートダウンロード
    with st.expander("📄 CSVテンプレートをダウンロード"):
        template_data = {
            "社員番号": ["1", "2"],
            "年度": [2026, 2026],
            "期": ["上期", "下期"],
            "等級": ["A2", "A2"],
            "個人目標": ["目標テキストをここに入力", ""],
            "個人目標グレード": ["B", "C"],
            "調整点数": [0, 0],
            "被評価者コメント": ["", ""],
            "評価者コメント": ["", ""],
        }
        df_template = pd.DataFrame(template_data)
        csv_template = df_template.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            "⬇️ テンプレートCSVをダウンロード",
            csv_template.encode("utf-8-sig"),
            "評価データ取込テンプレート.csv",
            mime="text/csv",
        )
        st.dataframe(df_template, use_container_width=True, hide_index=True)
        st.caption("※ 個人目標グレードは A/B/C/D/E のいずれかを入力してください")
        st.caption("※ 等級は A1/A2/B1/B2/C1/C2 のいずれかを入力してください")

    st.markdown("---")

    # ファイルアップロード
    uploaded = st.file_uploader(
        "ファイルをアップロード（CSV / Excel / PDF / Word）",
        type=["csv", "xlsx", "xls", "xlsm", "pdf", "docx", "doc"],
        help="CSV・Excel（xlsx/xls/xlsm）・PDF・Word（docx）に対応"
    )

    if not uploaded:
        return

    fname = uploaded.name.lower()

    # ファイル読み込み
    try:
        if fname.endswith(".csv"):
            # UTF-8 → Shift-JIS の順で試行
            raw = uploaded.read()
            df = None
            for enc in ["utf-8-sig", "utf-8", "shift_jis", "cp932"]:
                try:
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc, dtype=str)
                    break
                except Exception:
                    continue
            if df is None:
                st.error("CSVの文字コードを判別できませんでした")
                return

        elif fname.endswith(".pdf"):
            import pdfplumber
            raw = uploaded.read()
            tables = []
            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                for page in pdf.pages:
                    tbl = page.extract_table()
                    if tbl:
                        tables.append(tbl)
            if not tables:
                # テーブルなし → テキスト抽出してプレビュー表示
                with pdfplumber.open(io.BytesIO(raw)) as pdf:
                    text = "\n".join(p.extract_text() or "" for p in pdf.pages)
                st.warning("PDFにテーブルが見つかりませんでした。テキストを表示します（データ取込にはCSV/Excelを推奨）。")
                st.text_area("PDFテキスト", text, height=400)
                return
            # 最初のテーブルを使用（複数ある場合はユーザーが選択）
            if len(tables) > 1:
                tbl_idx = st.selectbox(
                    "取込するテーブルを選択",
                    list(range(len(tables))),
                    format_func=lambda i: f"テーブル {i+1}（{len(tables[i])}行）"
                )
                tbl = tables[tbl_idx]
            else:
                tbl = tables[0]
            header = tbl[0]
            rows = tbl[1:]
            df = pd.DataFrame(rows, columns=header).fillna("").astype(str)

        elif fname.endswith(".docx"):
            from docx import Document as DocxDoc
            raw = uploaded.read()
            doc = DocxDoc(io.BytesIO(raw))
            if not doc.tables:
                # テーブルなし → パラグラフをテキスト表示
                text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
                st.warning("Wordファイルにテーブルが見つかりませんでした（データ取込にはCSV/Excelを推奨）。")
                st.text_area("Wordテキスト", text, height=400)
                return
            if len(doc.tables) > 1:
                tbl_idx = st.selectbox(
                    "取込するテーブルを選択",
                    list(range(len(doc.tables))),
                    format_func=lambda i: f"テーブル {i+1}（{len(doc.tables[i].rows)}行）"
                )
                tbl = doc.tables[tbl_idx]
            else:
                tbl = doc.tables[0]
            rows = [[cell.text for cell in row.cells] for row in tbl.rows]
            header = rows[0]
            df = pd.DataFrame(rows[1:], columns=header).fillna("").astype(str)

        else:
            # Excel系（xlsx / xls / xlsm）
            import openpyxl
            raw_bytes = uploaded.read()
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            uploaded.seek(0)

            if len(sheet_names) > 1:
                selected_sheet = st.selectbox(
                    "読み込むシートを選択",
                    sheet_names,
                    help="取込対象のシートを選んでください"
                )
            else:
                selected_sheet = sheet_names[0]

            df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=selected_sheet,
                               dtype=str, engine="openpyxl")

    except Exception as e:
        st.error(f"ファイル読み込みエラー: {e}")
        return

    df = df.fillna("")

    st.subheader(f"プレビュー（{len(df)}行）")
    st.dataframe(df.head(20), use_container_width=True, hide_index=True)
    st.caption(f"列名: {', '.join(df.columns.tolist())}")

    # カラムマッピング（自動マッチング）
    col_map = {
        "社員番号": ["社員番号", "emp_id", "社員ID", "ID", "No", "NO", "no", "番号", "社員No"],
        "年度":   ["年度", "year", "評価年度", "年"],
        "期":     ["期", "period", "評価期", "上下期"],
        "等級":   ["等級", "grade_level", "等級レベル", "グレード", "grade"],
        "個人目標": ["個人目標", "pg_goal", "個人目標の内容", "目標", "目標内容"],
        "個人目標グレード": ["個人目標グレード", "pg_grade", "目標グレード", "個人目標評価"],
        "調整点数": ["調整点数", "adj_points", "調整", "加減点"],
        "被評価者コメント": ["被評価者コメント", "evaluatee_comment", "本人コメント", "自己コメント"],
        "評価者コメント": ["評価者コメント", "evaluator_comment", "上長コメント", "上司コメント"],
    }

    def find_col(df_cols, candidates):
        for c in candidates:
            if c in df_cols:
                return c
        # 部分一致も試みる
        for c in candidates:
            for dc in df_cols:
                if c in dc or dc in c:
                    return dc
        return None

    cols = df.columns.tolist()
    mapped = {k: find_col(cols, v) for k, v in col_map.items()}

    # 自動マッピング結果を表示
    with st.expander("🔧 列マッピング確認（クリックして展開）"):
        for k, v in mapped.items():
            req = "✅" if v else ("⚠️ 必須" if k in ["社員番号","年度","期"] else "（任意）")
            st.write(f"{req} **{k}** → `{v or '未マッピング'}`")

        # 必須列が未マッピングの場合は手動選択
        missing_required = [k for k in ["社員番号", "年度", "期"] if not mapped[k]]
        if missing_required:
            st.warning(f"必須列を手動で選択してください: {', '.join(missing_required)}")
            none_opt = ["（選択してください）"] + cols
            for req_k in missing_required:
                sel = st.selectbox(f"「{req_k}」に対応する列", none_opt, key=f"manual_{req_k}")
                if sel != "（選択してください）":
                    mapped[req_k] = sel

    missing_required = [k for k in ["社員番号", "年度", "期"] if not mapped[k]]
    if missing_required:
        st.error(f"必須列が選択されていません: {', '.join(missing_required)}")
        return

    employees = get_all_employees()
    emp_set = {e["emp_id"] for e in employees}

    # バリデーション
    errors = []
    valid_rows = []
    for i, row in df.iterrows():
        emp_id = str(row.get(mapped["社員番号"], "")).strip()
        year_raw = str(row.get(mapped["年度"], "")).strip()
        period = str(row.get(mapped["期"], "")).strip()

        if not emp_id:
            errors.append(f"行{i+2}: 社員番号が空")
            continue
        try:
            year = int(float(year_raw))
        except:
            errors.append(f"行{i+2}: 年度が不正 ({year_raw})")
            continue
        if period not in ["上期", "下期"]:
            errors.append(f"行{i+2}: 期が不正 ({period}) ← 上期 or 下期")
            continue

        gl_col = mapped.get("等級")
        gl = str(row.get(gl_col, "A2")).strip() if gl_col else "A2"
        if gl not in GRADE_LEVELS:
            gl = "A2"

        pg_goal_col = mapped.get("個人目標")
        pg_goal = str(row.get(pg_goal_col, "")).strip() if pg_goal_col else ""

        pg_grade_col = mapped.get("個人目標グレード")
        pg_grade = str(row.get(pg_grade_col, "C")).strip().upper() if pg_grade_col else "C"
        if pg_grade not in ["A", "B", "C", "D", "E"]:
            pg_grade = "C"

        adj_col = mapped.get("調整点数")
        try:
            adj = int(float(row.get(adj_col, 0))) if adj_col else 0
        except:
            adj = 0

        ev_comment_col = mapped.get("被評価者コメント")
        ev_comment = str(row.get(ev_comment_col, "")).strip() if ev_comment_col else ""

        er_comment_col = mapped.get("評価者コメント")
        er_comment = str(row.get(er_comment_col, "")).strip() if er_comment_col else ""

        pg_pts_map = PG_POINTS_BY_GRADE.get(gl, {})
        pg_pts = pg_pts_map.get(pg_grade, 0)
        total = pg_pts + adj
        final_g = calc_final_grade(total)

        valid_rows.append({
            "emp_id": emp_id, "year": year, "period": period,
            "grade_level": gl, "pg_goal": pg_goal,
            "pg_grade": pg_grade, "pg_points": pg_pts,
            "item_grades": {}, "item_points": {},
            "adj_points": adj, "total_score": total, "final_grade": final_g,
            "evaluatee_comment": ev_comment, "evaluator_comment": er_comment,
            "evaluator1": "", "evaluator2": "", "ai_comment": "",
        })

    if errors:
        st.warning(f"以下の行にエラーがあります（スキップします）:")
        for e in errors[:10]:
            st.write(f"  - {e}")

    if not valid_rows:
        st.error("取込可能なデータがありません")
        return

    st.success(f"✅ {len(valid_rows)} 件が取込可能です（エラー: {len(errors)} 件）")

    # 重複チェック
    existing_evals = get_evaluations()
    existing_keys = {(str(e["emp_id"]), e["year"], e["period"]) for e in existing_evals}
    new_rows = [r for r in valid_rows if (r["emp_id"], r["year"], r["period"]) not in existing_keys]
    update_rows = [r for r in valid_rows if (r["emp_id"], r["year"], r["period"]) in existing_keys]

    col_a, col_b = st.columns(2)
    col_a.metric("新規追加", f"{len(new_rows)} 件")
    col_b.metric("上書き更新", f"{len(update_rows)} 件")

    overwrite = st.checkbox("既存データを上書きする（チェックなし=新規のみ取込）")

    rows_to_import = new_rows + (update_rows if overwrite else [])

    if not rows_to_import:
        st.info("取込対象がありません（全て既存データです）")
        return

    if st.button(f"📥 {len(rows_to_import)} 件を取込実行", type="primary", use_container_width=True):
        ok = 0
        ng = 0
        ng_details = []
        progress = st.progress(0)
        for idx, rec in enumerate(rows_to_import):
            # 既存データのIDを取得して上書き
            if overwrite and (rec["emp_id"], rec["year"], rec["period"]) in existing_keys:
                matched = next(
                    (e for e in existing_evals
                     if str(e["emp_id"]) == rec["emp_id"]
                     and e["year"] == rec["year"]
                     and e["period"] == rec["period"]),
                    None
                )
                if matched:
                    rec["id"] = matched["id"]
            try:
                upsert_evaluation(rec)
                ok += 1
            except Exception as ex:
                ng += 1
                ng_details.append(f"  社員{rec['emp_id']} {rec['year']}{rec['period']}: {ex}")
            progress.progress((idx + 1) / len(rows_to_import))

        # 取込履歴に記録
        detail = f"新規:{len(new_rows)}件, 更新:{len(update_rows) if overwrite else 0}件, エラー:{ng}件"
        insert_import_history({
            "file_name": uploaded.name,
            "file_type": "CSV" if uploaded.name.endswith(".csv") else "Excel",
            "affected_count": ok,
            "detail": detail,
            "status": "完了" if ng == 0 else "一部エラー",
        })

        if ok:
            st.success(f"✅ {ok} 件の取込が完了しました！")
        if ng:
            st.error(f"❌ {ng} 件でエラーが発生しました")
            for d in ng_details[:5]:
                st.write(d)

        st.info("「📋 評価一覧」で取込結果を確認できます")
        st.info("取込履歴は「📂 取込履歴」に記録されました")

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
elif "📥 一括取込" in page:
    page_bulk_import()
