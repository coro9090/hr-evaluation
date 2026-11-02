# -*- coding: utf-8 -*-
"""
データベース操作モジュール（Supabase）
"""
import streamlit as st
from supabase import create_client, Client
import json
from datetime import datetime

@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["anon_key"]
    return create_client(url, key)

# ============================================================
# 社員
# ============================================================
def get_all_employees():
    sb = get_supabase()
    res = sb.table("employees").select("*").order("emp_id").execute()
    return res.data or []

def get_employee(emp_id: str):
    sb = get_supabase()
    res = sb.table("employees").select("*").eq("emp_id", emp_id).execute()
    return res.data[0] if res.data else None

def upsert_employee(data: dict):
    sb = get_supabase()
    res = sb.table("employees").upsert(data, on_conflict="emp_id").execute()
    return res.data

def delete_employee(emp_id: str):
    sb = get_supabase()
    sb.table("evaluations").delete().eq("emp_id", emp_id).execute()
    sb.table("employees").delete().eq("emp_id", emp_id).execute()

# ============================================================
# 評価
# ============================================================
def get_evaluations(emp_id: str = None):
    sb = get_supabase()
    q = sb.table("evaluations").select("*").order("created_at", desc=True)
    if emp_id:
        q = q.eq("emp_id", emp_id)
    res = q.execute()
    return res.data or []

def get_evaluation(eval_id: str):
    sb = get_supabase()
    res = sb.table("evaluations").select("*").eq("id", eval_id).execute()
    return res.data[0] if res.data else None

def upsert_evaluation(data: dict):
    sb = get_supabase()
    # JSON列をシリアライズ
    for col in ("item_grades", "item_points"):
        if col in data and isinstance(data[col], dict):
            data[col] = json.dumps(data[col], ensure_ascii=False)
    data["updated_at"] = datetime.utcnow().isoformat()
    res = sb.table("evaluations").upsert(data).execute()
    return res.data

def delete_evaluation(eval_id: str):
    sb = get_supabase()
    sb.table("evaluations").delete().eq("id", eval_id).execute()

def parse_json_cols(ev: dict) -> dict:
    """item_grades / item_points が文字列ならdictに戻す"""
    for col in ("item_grades", "item_points"):
        v = ev.get(col)
        if isinstance(v, str):
            try:
                ev[col] = json.loads(v)
            except Exception:
                ev[col] = {}
        elif v is None:
            ev[col] = {}
    return ev

# ============================================================
# 日次記録
# ============================================================
def get_daily_records(emp_id: str = None, limit: int = 200):
    sb = get_supabase()
    q = sb.table("daily_records").select("*").order("record_date", desc=True).limit(limit)
    if emp_id:
        q = q.eq("emp_id", emp_id)
    res = q.execute()
    return res.data or []

def insert_daily_record(data: dict):
    sb = get_supabase()
    data["created_at"] = datetime.utcnow().isoformat()
    data["updated_at"] = datetime.utcnow().isoformat()
    res = sb.table("daily_records").insert(data).execute()
    return res.data

def delete_daily_record(record_id: int):
    sb = get_supabase()
    sb.table("daily_records").delete().eq("id", record_id).execute()

# ============================================================
# 取込履歴
# ============================================================
def get_import_history(limit: int = 300):
    sb = get_supabase()
    res = sb.table("import_history").select("*").order("import_date", desc=True).limit(limit).execute()
    return res.data or []

def delete_import_history(record_id: int):
    sb = get_supabase()
    sb.table("import_history").delete().eq("id", record_id).execute()

def insert_import_history(data: dict):
    sb = get_supabase()
    data["import_date"] = datetime.utcnow().isoformat()
    res = sb.table("import_history").insert(data).execute()
    return res.data
