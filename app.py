"""
app.py — POS売上分析ポータル ホーム画面

起動方法:
    streamlit run app.py
"""

import streamlit as st

from pos_db import get_summary, load_filtered
from pos_analysis import RETAILER_CONFIG

st.set_page_config(
    page_title="POS売上分析ポータル",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
/* カード全体 */
.retailer-card {
    border-radius: 14px;
    padding: 28px 20px 20px;
    text-align: center;
    border: 2px solid #E0E0E0;
    background: white;
    transition: box-shadow 0.2s, border-color 0.2s;
    height: 100%;
}
.retailer-card:hover {
    box-shadow: 0 6px 20px rgba(0,0,0,0.12);
}
/* メトリクス */
[data-testid="metric-container"] {
    background: #F8FAFC;
    border: 1px solid #E2ECF6;
    border-radius: 10px;
    padding: 12px 16px;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 700 !important;
    color: #1F4E79;
}
/* サイドバー非表示にしてすっきりさせる */
[data-testid="collapsedControl"] { display: none; }
</style>
""", unsafe_allow_html=True)

# ─── ページタイトル ──────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;padding:32px 0 8px">
    <h1 style="font-size:2.4rem;color:#1F4E79;margin:0">🏪 POS売上分析ポータル</h1>
    <p style="color:#666;font-size:1.05rem;margin-top:8px">
        小売店ごとのPOSデータを蓄積・分析するダッシュボード
    </p>
</div>
""", unsafe_allow_html=True)

# ─── 全体サマリー ────────────────────────────────────────────────────
summary = get_summary()

st.markdown("---")
c1, c2, c3, c4 = st.columns(4)
c1.metric("総レコード数",    f"{summary['total_records']:,} 件")
c2.metric("蓄積月数",        f"{len(summary['year_months'])} ヶ月")
c3.metric("取り扱いブランド", f"{len(summary['brands'])} 件")
c4.metric("登録小売店",      "4 社")
st.markdown("---")

# ─── 小売店カード ────────────────────────────────────────────────────
RETAILERS = list(RETAILER_CONFIG.keys())   # ["PLAZA", "ハンズ", "ロフト", "アインズ"]
PAGE_MAP   = {
    "PLAZA":    "pages/1_PLAZA.py",
    "ハンズ":   "pages/2_Hands.py",
    "ロフト":   "pages/3_Loft.py",
    "アインズ": "pages/4_Ainz.py",
}

cols = st.columns(4, gap="large")

for col, retailer in zip(cols, RETAILERS):
    cfg   = RETAILER_CONFIG[retailer]
    color = cfg["color"]
    icon  = cfg["icon"]
    bg    = cfg["bg"]

    # 小売店ごとの統計
    df_r    = load_filtered(retailer=retailer)
    total   = len(df_r)
    months  = df_r["年月"].nunique() if not df_r.empty else 0
    brands  = df_r["ブランド名"].nunique() if not df_r.empty else 0
    last_ym = df_r["年月"].max().replace("-", "年") + "月" if not df_r.empty else "—"

    with col:
        st.markdown(f"""
        <div class="retailer-card" style="border-color:{color}20;
             background:linear-gradient(160deg, white 60%, {bg} 100%)">
            <div style="font-size:3.2rem;line-height:1">{icon}</div>
            <h2 style="color:{color};margin:10px 0 6px;font-size:1.5rem">{retailer}</h2>
            <hr style="border-color:{color}30;margin:10px 0">
            <p style="color:#555;margin:4px 0;font-size:0.9rem">
                📦 <b>{total:,}</b> 件蓄積
            </p>
            <p style="color:#555;margin:4px 0;font-size:0.9rem">
                📅 <b>{months}</b> ヶ月分
            </p>
            <p style="color:#555;margin:4px 0;font-size:0.9rem">
                🏷 ブランド <b>{brands}</b> 件
            </p>
            <p style="color:#888;margin:8px 0 0;font-size:0.8rem">
                最終更新: {last_ym}
            </p>
        </div>
        """, unsafe_allow_html=True)

        # ページ遷移ボタン
        st.markdown("<div style='margin-top:12px'>", unsafe_allow_html=True)
        if st.button(
            f"📊 {retailer}の分析を開く",
            key=f"nav_{retailer}",
            use_container_width=True,
            type="primary" if total > 0 else "secondary",
        ):
            st.switch_page(PAGE_MAP[retailer])
        st.markdown("</div>", unsafe_allow_html=True)

# ─── 使い方ガイド ────────────────────────────────────────────────────
st.markdown("---")
with st.expander("📖 使い方ガイド"):
    st.markdown("""
    #### 基本的な流れ

    1. **上のカードから小売店を選んで**「分析を開く」ボタンをクリック
    2. 左サイドバーの **ファイルアップロード** からCSV/Excelを選択
    3. **「DBに保存する」** を押すとデータが蓄積されます
    4. 毎月アップロードするだけで自動的に履歴が蓄積されます

    #### 各タブの説明

    | タブ | 内容 |
    |------|------|
    | 📈 ダッシュボード | KPI・日別推移・ヒートマップ・店舗ランキング |
    | 📅 月別比較 | 複数月を選んで売上推移を重ね合わせ比較 |
    | 📋 商品別マトリックス | 日×店舗のクロス集計表（カラーグラデーション付き） |
    | 📥 Excelダウンロード | メーカー提出用のExcelレポートを自動生成 |
    | 🗄 データ管理 | 蓄積データ一覧・削除・CSVエクスポート |

    #### CSVフォーマット（必須列）

    | 列名 | 例 |
    |------|----|
    | 日付 | 2026-05-01 |
    | 小売店名 | PLAZA |
    | 店舗名 | 渋谷店 |
    | ブランド名 | ブランドA |
    | 商品名 | 商品X |
    | 売上数量 | 5 |
    | 売上金額 | 2500 |
    """)
