"""
auth.py — パスワード認証モジュール

Streamlit Cloud の secrets.toml に以下を設定してください:
    APP_PASSWORD = "任意のパスワード"
"""

import streamlit as st


def check_password() -> bool:
    """パスワード認証。未認証なら入力画面を表示して False を返す。"""
    try:
        correct = st.secrets["APP_PASSWORD"]
    except Exception:
        return True  # secrets 未設定（ローカル開発）はそのまま通す

    if st.session_state.get("_authenticated"):
        return True

    # 認証画面
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { display: none; }
    </style>
    <div style="max-width:380px;margin:80px auto;text-align:center">
        <h1 style="color:#1F4E79;font-size:2rem">🏪 POS分析ポータル</h1>
        <p style="color:#666;margin-bottom:32px">アクセスにはパスワードが必要です</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pwd = st.text_input("パスワード", type="password", key="_pwd_input",
                            label_visibility="collapsed",
                            placeholder="パスワードを入力")
        if st.button("ログイン", use_container_width=True, type="primary", key="_login_btn"):
            if pwd == correct:
                st.session_state["_authenticated"] = True
                st.rerun()
            else:
                st.error("パスワードが違います")

    return False
