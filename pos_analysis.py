"""
pos_analysis.py — 各小売店ページ共通の分析レンダリングモジュール

使い方（各 pages/*.py から呼び出す）:
    from pos_analysis import render_retailer_page
    render_retailer_page("PLAZA")
"""

import calendar
import os
import tempfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from pos_db import (
    delete_by_month_brand,
    export_csv_bytes,
    get_summary,
    list_month_brand_pairs,
    load_filtered,
    save_records,
)
from pos_report import create_sample_data, generate_reports, load_pos_data

# ─── 小売店ブランド設定 ──────────────────────────────────────────────
RETAILER_CONFIG: dict[str, dict] = {
    "PLAZA":        {"color": "#C0392B", "icon": "🛍️",  "bg": "#FDF2F2"},
    "ハンズ":       {"color": "#1A6BA0", "icon": "🔧",  "bg": "#EBF5FB"},
    "ロフト":       {"color": "#D68910", "icon": "✏️",  "bg": "#FEF9E7"},
    "アインズ":     {"color": "#1E8449", "icon": "💄",  "bg": "#EAFAF1"},
    "アットコスメ": {"color": "#C2185B", "icon": "🌸",  "bg": "#FCE4EC"},
}


# ─── ユーティリティ ──────────────────────────────────────────────────

def _load_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    import io as _io
    from pos_report import (
        _is_ainz_format, _convert_ainz_format,
        _is_loft_raw_format, _convert_loft_format,
        _is_plaza_raw_format, _convert_plaza_format,
        _is_cosme_format, _convert_cosme_format,
        REQUIRED_COLS,
    )

    if not file_bytes:
        raise ValueError(f"ファイルが空です（0バイト）: {filename}")

    suffix = Path(filename).suffix.lower()

    if suffix in (".xlsx", ".xls"):
        df = pd.read_excel(_io.BytesIO(file_bytes), dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        if _is_ainz_format(df):
            df = _convert_ainz_format(df, Path(filename))
        elif _is_cosme_format(df):
            df = _convert_cosme_format(df)
        elif _is_plaza_raw_format(df):
            df = _convert_plaza_format(df)

    elif suffix == ".csv":
        # csv.reader で直接パース（pandas の EmptyDataError を完全回避）
        import csv as _csv
        text = None
        for enc in ("cp932", "utf-8-sig", "utf-8", "euc-jp"):
            try:
                text = file_bytes.decode(enc)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        if text is None:
            text = file_bytes.decode("cp932", errors="replace")
        # CR のみ / CRLF を LF に正規化
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        rows = list(_csv.reader(_io.StringIO(text)))
        rows = [r for r in rows if any(c.strip() for c in r)]  # 空行除去
        if not rows:
            raise ValueError(f"ファイルにデータがありません: {filename}")
        df_raw = pd.DataFrame(rows)

        if df_raw.empty:
            raise ValueError(f"ファイルにデータがありません: {filename}")

        # 先頭セルが YYYYMMDD（8桁数字）ならロフト形式（ヘッダーなし）
        first_cell = str(df_raw.iloc[0, 0]).strip().strip('"')
        if first_cell.isdigit() and len(first_cell) == 8 and len(df_raw.columns) >= 9:
            df = _convert_loft_format(df_raw)
        else:
            # 通常形式：1行目がヘッダー
            headers = [str(c).strip() for c in df_raw.iloc[0]]
            df = df_raw.iloc[1:].reset_index(drop=True)
            df.columns = headers
            if _is_ainz_format(df):
                df = _convert_ainz_format(df, Path(filename))
            elif _is_plaza_raw_format(df):
                df = _convert_plaza_format(df)

    else:
        raise ValueError(f"未対応のファイル形式: {suffix}（.xlsx / .xls / .csv のみ対応）")

    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        raise ValueError(
            f"必須カラムが見つかりません: {sorted(missing)}\n実際のカラム: {list(df.columns)}"
        )

    df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
    df = df.dropna(subset=["日付"])
    df["売上数量"] = pd.to_numeric(df["売上数量"], errors="coerce").fillna(0).astype(int)
    if "売上金額" not in df.columns:
        df["売上金額"] = 0
    df["売上金額"] = pd.to_numeric(df["売上金額"], errors="coerce").fillna(0).astype(int)
    for col in ("小売店名", "店舗名", "ブランド名", "商品名"):
        df[col] = df[col].astype(str).str.strip().replace("nan", "（不明）").fillna("（不明）")

    return df


@st.cache_data(ttl=5, show_spinner=False)
def _cached_load(retailer: str) -> pd.DataFrame:
    try:
        return load_filtered(retailer=retailer)
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=5, show_spinner=False)
def _cached_summary(retailer: str) -> dict:
    # _cached_load を呼ばず直接 load_filtered を呼ぶ（ネストキャッシュ回避）
    try:
        df = load_filtered(retailer=retailer)
    except Exception:
        df = pd.DataFrame()
    if df.empty:
        return {"total_records": 0, "year_months": [], "brands": [],
                "date_range": (None, None)}
    return {
        "total_records": len(df),
        "year_months":   sorted(df["年月"].unique().tolist()),
        "brands":        sorted(df["ブランド名"].unique().tolist()),
        "date_range":    (df["日付"].min().strftime("%Y-%m-%d"),
                          df["日付"].max().strftime("%Y-%m-%d")),
    }


def _clear_cache(retailer: str) -> None:
    _cached_load.clear()
    _cached_summary.clear()


# ═══════════════════════════════════════════════════════════════════════
# メインレンダリング関数
# ═══════════════════════════════════════════════════════════════════════

def render_retailer_page(retailer_name: str) -> None:
    """小売店専用ダッシュボードページ全体をレンダリングする"""

    cfg   = RETAILER_CONFIG.get(retailer_name, {"color": "#1F4E79", "icon": "🏪", "bg": "#EBF5FB"})
    color = cfg["color"]
    icon  = cfg["icon"]
    bg    = cfg["bg"]

    # ── ページヘッダー ───────────────────────────────────────────────
    st.markdown(f"""
    <div style="background:{bg};border-left:6px solid {color};
                padding:14px 20px;border-radius:0 10px 10px 0;margin-bottom:4px">
        <h1 style="color:{color};margin:0;font-size:1.8rem">{icon} {retailer_name}</h1>
        <p style="color:#666;margin:4px 0 0;font-size:0.9rem">POSデータ分析ダッシュボード</p>
    </div>
    """, unsafe_allow_html=True)

    summary = _cached_summary(retailer_name)

    # ═══════════════════════════════════════════════════════════════
    # サイドバー
    # ═══════════════════════════════════════════════════════════════
    with st.sidebar:
        st.markdown(f"## {icon} {retailer_name}")
        st.markdown("---")

        # DB ステータス
        if summary["total_records"] > 0:
            st.success(
                f"🗄 **{summary['total_records']:,}** 件蓄積済み  \n"
                f"{summary['date_range'][0]} 〜 {summary['date_range'][1]}"
            )
        else:
            st.info("🗄 まだデータがありません")

        st.markdown("---")

        # ファイルアップロード
        st.markdown("### データを追加")
        uploaded = st.file_uploader(
            "CSVまたはExcelをアップロード",
            type=["csv", "xlsx", "xls"],
            label_visibility="collapsed",
            key=f"upload_{retailer_name}",
        )

        # 保存結果をセッション状態で保持（st.rerun後も表示できるよう）
        _save_key = f"_save_result_{retailer_name}"
        if _save_key in st.session_state:
            res = st.session_state.pop(_save_key)
            if res.get("error"):
                st.error(res["error"])
            else:
                st.success(f"✅ **{res['inserted']:,}** 件保存しました（うち {retailer_name}: {res['this']:,} 件）")
                if res.get("replaced"):
                    st.warning("上書きしたデータ:\n" + "\n".join(f"• {r}" for r in res["replaced"]))

        if uploaded is not None:
            st.caption(f"📎 {uploaded.name}  ({uploaded.size:,} bytes)")
            if st.button("💾 DBに保存する", type="primary",
                         use_container_width=True, key=f"save_{retailer_name}"):
                with st.spinner("保存中…"):
                    file_bytes = b""
                    try:
                        uploaded.seek(0)
                        file_bytes = uploaded.read()
                        df_up = _load_file(file_bytes, uploaded.name)
                        df_this = df_up[df_up["小売店名"] == retailer_name]
                        result  = save_records(df_up)
                        _clear_cache(retailer_name)
                        st.session_state[_save_key] = {
                            "inserted": result["inserted"],
                            "this":     len(df_this),
                            "replaced": result.get("replaced", []),
                        }
                    except Exception as e:
                        head_hex = file_bytes[:16].hex() if file_bytes else "empty"
                        st.session_state[_save_key] = {
                            "error": f"エラー [{len(file_bytes):,}bytes, head={head_hex}]: {e}"
                        }
                    st.rerun()

        # サンプルデータ
        with st.expander("🔽 サンプルデータで試す"):
            if st.button("サンプルCSVを生成", key=f"sample_{retailer_name}"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as f:
                    path = f.name
                create_sample_data(path)
                data = Path(path).read_bytes()
                os.unlink(path)
                st.download_button("⬇ ダウンロード", data, "sample_pos.csv",
                                   "text/csv", key=f"dl_sample_{retailer_name}")

        st.markdown("---")

        # フィルター
        st.markdown("### フィルター")
        if summary["total_records"] == 0:
            st.caption("データをアップロードするとフィルターが表示されます")
            filter_ym    = None
            filter_brand = None
        else:
            ym_list   = summary["year_months"]
            filter_ym = st.selectbox(
                "年月", ym_list, index=len(ym_list) - 1,
                format_func=lambda s: s.replace("-", "年") + "月",
                key=f"ym_{retailer_name}",
            )
            brands_in_ym = sorted(
                load_filtered(year_month=filter_ym, retailer=retailer_name)
                ["ブランド名"].dropna().unique()
            )
            filter_brand = (
                st.selectbox("ブランド", brands_in_ym, key=f"brand_{retailer_name}")
                if brands_in_ym else None
            )

    # ─── データなし ──────────────────────────────────────────────────
    if summary["total_records"] == 0:
        st.info(
            f"左サイドバーから {retailer_name} のPOSデータをアップロードしてください。\n\n"
            "「サンプルデータで試す」からサンプルCSVを使って動作確認もできます。"
        )
        return

    # ─── 分析データ取得 ───────────────────────────────────────────────
    df_all = _cached_load(retailer_name)
    df = load_filtered(year_month=filter_ym,
                       brand=filter_brand,
                       retailer=retailer_name)

    if filter_ym:
        _year  = int(filter_ym.split("-")[0])
        _month = int(filter_ym.split("-")[1])
    else:
        _year  = df_all["日付"].dt.year.max()
        _month = df_all["日付"].dt.month.max()

    # ═══════════════════════════════════════════════════════════════
    # タブ
    # ═══════════════════════════════════════════════════════════════
    tab_dash, tab_compare, tab_matrix, tab_excel, tab_manage = st.tabs([
        "📈 ダッシュボード",
        "📅 月別比較",
        "📋 商品別マトリックス",
        "📥 Excelダウンロード",
        "🗄 データ管理",
    ])

    # ── Tab1: ダッシュボード ─────────────────────────────────────────
    with tab_dash:
        ym_label    = filter_ym.replace("-", "年") + "月" if filter_ym else "全期間"
        brand_label = filter_brand or "全ブランド"
        st.markdown(f"### {ym_label}　{brand_label}")

        if df.empty:
            st.info("条件に合うデータがありません"); return

        total_qty   = int(df["売上数量"].sum())
        n_products  = df["商品名"].nunique()
        n_stores    = df["店舗名"].nunique()
        top_product = df.groupby("商品名")["売上数量"].sum().idxmax()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("総売上数量",    f"{total_qty:,} 個")
        c2.metric("取り扱い商品数", f"{n_products} 品目")
        c3.metric("取り扱い店舗数", f"{n_stores} 店舗")
        c4.metric("最多売上商品",   top_product)

        st.markdown("---")
        col_l, col_r = st.columns([3, 2])

        with col_l:
            st.subheader("日別売上推移")
            trend = df.groupby(["日付", "商品名"])["売上数量"].sum().reset_index()
            fig = px.line(trend, x="日付", y="売上数量", color="商品名",
                          markers=True, template="plotly_white",
                          color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(height=320, margin=dict(t=10, b=10),
                              legend=dict(orientation="h", y=-0.3),
                              xaxis_tickformat="%m/%d")
            st.plotly_chart(fig, use_container_width=True)

        with col_r:
            st.subheader("商品別売上ランキング")
            rank = (df.groupby("商品名")["売上数量"].sum()
                    .sort_values(ascending=True).reset_index())
            fig2 = px.bar(rank, x="売上数量", y="商品名", orientation="h",
                          text="売上数量", template="plotly_white",
                          color="売上数量", color_continuous_scale="Blues")
            fig2.update_traces(textposition="outside")
            fig2.update_layout(height=320, margin=dict(t=10, b=10),
                               yaxis_title=None, coloraxis_showscale=False)
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("---")
        st.subheader("店舗 × 商品　売上ヒートマップ")
        hm_pivot = (df.groupby(["店舗名", "商品名"])["売上数量"].sum()
                    .reset_index()
                    .pivot_table(index="店舗名", columns="商品名",
                                 values="売上数量", fill_value=0))
        fig3 = px.imshow(hm_pivot, text_auto=True, aspect="auto",
                         color_continuous_scale="Blues", template="plotly_white")
        fig3.update_layout(height=max(250, len(hm_pivot) * 45 + 100),
                           margin=dict(t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(fig3, use_container_width=True)

        st.markdown("---")
        st.subheader("店舗別売上ランキング TOP10")
        store_rank = (df.groupby("店舗名")["売上数量"].sum().reset_index()
                      .sort_values("売上数量", ascending=False)
                      .head(10).sort_values("売上数量", ascending=True))
        fig4 = px.bar(store_rank, x="売上数量", y="店舗名", orientation="h",
                      template="plotly_white",
                      color_discrete_sequence=[color])
        fig4.update_layout(height=max(260, len(store_rank) * 38 + 80),
                           margin=dict(t=10, b=10), yaxis_title=None)
        st.plotly_chart(fig4, use_container_width=True)

    # ── Tab2: 月別比較 ───────────────────────────────────────────────
    with tab_compare:
        st.markdown("## 月別比較")
        ym_all = summary["year_months"]

        if len(ym_all) < 2:
            st.info("月別比較には2ヶ月以上のデータが必要です");
        else:
            col1, col2 = st.columns(2)
            with col1:
                cmp_brand = st.selectbox("ブランド", summary["brands"],
                                         key=f"cmp_b_{retailer_name}")
            with col2:
                cmp_months = st.multiselect(
                    "比較する月（複数可）", ym_all,
                    default=ym_all[-min(3, len(ym_all)):],
                    format_func=lambda s: s.replace("-", "年") + "月",
                    key=f"cmp_m_{retailer_name}",
                )

            if cmp_months:
                df_cmp = df_all[
                    (df_all["ブランド名"] == cmp_brand) &
                    (df_all["年月"].isin(cmp_months))
                ].copy()
                df_cmp["年月ラベル"] = df_cmp["年月"].str.replace("-", "年") + "月"

                if not df_cmp.empty:
                    st.markdown("---")
                    st.subheader("月別 × 商品別　売上合計")
                    mp = (df_cmp.groupby(["年月ラベル", "商品名"])["売上数量"]
                          .sum().reset_index())
                    fig_c1 = px.bar(mp, x="年月ラベル", y="売上数量", color="商品名",
                                    barmode="group", template="plotly_white",
                                    text_auto=True,
                                    color_discrete_sequence=px.colors.qualitative.Set2)
                    fig_c1.update_layout(height=360, margin=dict(t=10, b=10),
                                         xaxis_title=None,
                                         legend=dict(orientation="h", y=-0.2))
                    st.plotly_chart(fig_c1, use_container_width=True)

                    st.markdown("---")
                    st.subheader("日次推移の重ね合わせ")
                    prods_cmp = sorted(df_cmp["商品名"].unique())
                    sel_p = st.selectbox("商品", prods_cmp,
                                         key=f"cmp_p_{retailer_name}")
                    df_pc = df_cmp[df_cmp["商品名"] == sel_p].copy()
                    df_pc["日"] = df_pc["日付"].dt.day
                    daily_cmp = (df_pc.groupby(["年月ラベル", "日"])["売上数量"]
                                 .sum().reset_index())
                    fig_c2 = px.line(daily_cmp, x="日", y="売上数量",
                                     color="年月ラベル", markers=True,
                                     template="plotly_white",
                                     color_discrete_sequence=px.colors.qualitative.Bold)
                    fig_c2.update_layout(height=340, margin=dict(t=10, b=10),
                                          legend=dict(orientation="h", y=-0.2),
                                          xaxis=dict(dtick=1))
                    st.plotly_chart(fig_c2, use_container_width=True)

                    st.markdown("---")
                    st.subheader("月別サマリー表")
                    tbl = (df_cmp.groupby(["年月ラベル", "商品名"])["売上数量"]
                           .sum().unstack(fill_value=0))
                    tbl["合計"] = tbl.sum(axis=1)
                    st.dataframe(
                        tbl.style.background_gradient(cmap="Blues", axis=None)
                        .format("{:,.0f}"),
                        use_container_width=True,
                    )

    # ── Tab3: 商品別マトリックス ─────────────────────────────────────
    with tab_matrix:
        st.markdown("## 商品別　日別売上マトリックス")
        if df.empty:
            st.info("データがありません")
        else:
            products = sorted(df["商品名"].dropna().unique())
            sel_prod = st.selectbox("商品を選択", products,
                                    key=f"mat_p_{retailer_name}")
            df_p = df[df["商品名"] == sel_prod].copy()

            if not df_p.empty:
                _, last_day = calendar.monthrange(_year, _month)
                df_p["_day"] = df_p["日付"].dt.day
                pivot = (df_p.pivot_table(
                    index="店舗名", columns="_day",
                    values="売上数量", aggfunc="sum", fill_value=0)
                    .reindex(columns=list(range(1, last_day + 1)), fill_value=0))
                pivot.columns = [f"{c}日" for c in pivot.columns]
                pivot["合計"] = pivot.sum(axis=1)
                total_row = pivot.sum(axis=0).to_frame().T
                total_row.index = pd.Index(["合計"], name="店舗名")
                pivot_disp = pd.concat([pivot, total_row])

                day_cols = [f"{d}日" for d in range(1, last_day + 1)]
                try:
                    st.dataframe(
                        pivot_disp.style
                        .background_gradient(cmap="Blues", subset=day_cols)
                        .format("{:.0f}"),
                        use_container_width=True,
                        height=min(600, (len(pivot_disp) + 1) * 36 + 60),
                    )
                except Exception:
                    st.dataframe(pivot_disp, use_container_width=True)

                st.markdown("---")
                st.subheader(f"{sel_prod}　店舗別日次推移")
                store_trend = (df_p.groupby(["日付", "店舗名"])["売上数量"]
                               .sum().reset_index())
                fig5 = px.line(store_trend, x="日付", y="売上数量", color="店舗名",
                               markers=True, template="plotly_white",
                               color_discrete_sequence=px.colors.qualitative.Set1)
                fig5.update_layout(height=340, margin=dict(t=10, b=10),
                                    legend=dict(orientation="h", y=-0.2),
                                    xaxis_tickformat="%m/%d")
                st.plotly_chart(fig5, use_container_width=True)

    # ── Tab4: Excelダウンロード ──────────────────────────────────────
    with tab_excel:
        st.markdown("## Excelレポート生成")
        if filter_ym and filter_brand:
            st.markdown(
                f"**対象:** {filter_ym.replace('-','年')}月　{filter_brand}　{retailer_name}"
            )
        split_choice = st.radio(
            "出力形式",
            ["ブランドごとにファイル（商品ごとにシート）", "商品ごとに別ファイル"],
            horizontal=True, key=f"split_{retailer_name}",
        )
        split_mode = "brand" if "ブランド" in split_choice else "product"

        if st.button("📊 Excelを生成する", type="primary",
                     key=f"gen_{retailer_name}"):
            if df.empty:
                st.warning("対象データがありません")
            else:
                with st.spinner("生成中…"):
                    try:
                        with tempfile.TemporaryDirectory() as d:
                            files = generate_reports(df, Path(d), split_mode)
                            file_data = [(p.name, p.read_bytes()) for p in files]
                        st.success(f"{len(file_data)} ファイルを生成しました")
                        for fname, fbytes in file_data:
                            st.download_button(
                                f"⬇ {fname}", fbytes, fname,
                                "application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet",
                                key=f"dl_{retailer_name}_{fname}",
                            )
                    except Exception as e:
                        st.error(f"生成エラー: {e}")

    # ── Tab5: データ管理 ─────────────────────────────────────────────
    with tab_manage:
        st.markdown(f"## {retailer_name} データ管理")

        pairs = list_month_brand_pairs(retailer=retailer_name)
        if pairs:
            rows_s = []
            for ym, brand in pairs:
                cnt = len(load_filtered(year_month=ym,
                                        brand=brand, retailer=retailer_name))
                rows_s.append({"年月": ym.replace("-", "年") + "月",
                               "ブランド名": brand, "件数": cnt})
            st.dataframe(pd.DataFrame(rows_s), use_container_width=True,
                         hide_index=True)
        else:
            st.info("データがありません")

        st.markdown("---")
        st.markdown("### CSVエクスポート")
        if summary["total_records"] > 0:
            df_exp = _cached_load(retailer_name).copy()
            df_exp["日付"] = df_exp["日付"].dt.strftime("%Y-%m-%d")
            csv_bytes = df_exp.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                f"⬇ {retailer_name}の全データをCSVでダウンロード",
                csv_bytes, f"{retailer_name}_pos_data.csv", "text/csv",
                key=f"csv_{retailer_name}",
            )

        st.markdown("---")
        st.markdown("### 月×ブランド単位で削除")
        if pairs:
            pair_labels = [f"{ym.replace('-','年')}月 / {b}" for ym, b in pairs]
            sel = st.selectbox("削除するデータ", pair_labels,
                               key=f"del_sel_{retailer_name}")
            idx       = pair_labels.index(sel)
            del_ym, del_brand = pairs[idx]
            if st.button("🗑 削除する", type="secondary",
                         key=f"del_{retailer_name}"):
                n = delete_by_month_brand(del_ym, del_brand)
                _clear_cache(retailer_name)
                st.success(f"✅ {n:,} 件を削除しました")
                st.rerun()

        st.markdown("---")
        st.markdown("### 一括削除")
        if summary["total_records"] > 0:
            if st.button(f"🗑 {retailer_name}の全データを削除する",
                         type="secondary", key=f"del_all_{retailer_name}"):
                from pos_db import delete_retailer_all
                n = delete_retailer_all(retailer_name)
                _clear_cache(retailer_name)
                st.success(f"✅ {n:,} 件をすべて削除しました")
                st.rerun()
        else:
            st.caption("削除するデータがありません")
