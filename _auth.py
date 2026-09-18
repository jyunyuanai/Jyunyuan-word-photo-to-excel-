# -*- coding: utf-8 -*-
"""簡單的密碼保護：在 Streamlit Cloud 的 App settings → Secrets 裡設定
APP_PASSWORD = "你的密碼"
本機執行、或沒有設定 APP_PASSWORD 時，不會要求密碼。"""
import streamlit as st


def require_password():
    try:
        pw = st.secrets.get("APP_PASSWORD")
    except Exception:
        pw = None
    if not pw:
        return

    if st.session_state.get("_authed"):
        return

    st.title("🔒 請輸入密碼")
    entered = st.text_input("密碼", type="password", key="_pw_input")
    if st.button("登入", key="_pw_submit"):
        if entered == pw:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("密碼錯誤")
    st.stop()
