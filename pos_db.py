"""
pos_db.py — POS データ永続化モジュール

ローカル実行時  : SQLite  (pos_data.db)
Streamlit Cloud : GitHub リポジトリ内の CSV ファイル（小売店ごとに分割）
                  data/pos_PLAZA.csv, data/pos_ロフト.csv, ...
                  secrets.toml に以下を設定:
                    GITHUB_TOKEN = "ghp_xxxx"
                    GITHUB_OWNER = "あなたのGitHubユーザー名"
                    GITHUB_REPO  = "pos-portal"
"""

import base64
import io
import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pandas as pd

# ─────────────────────────────────────────────────
# モード判定
# ─────────────────────────────────────────────────

def _get_github_config() -> dict | None:
    """GitHub 設定を secrets / 環境変数から取得。未設定なら None。"""
    try:
        import streamlit as st
        token = st.secrets.get("GITHUB_TOKEN", None)
        owner = st.secrets.get("GITHUB_OWNER", None)
        repo  = st.secrets.get("GITHUB_REPO",  None)
        if token and owner and repo:
            return {"token": token, "owner": owner, "repo": repo}
    except Exception:
        pass
    token = os.environ.get("GITHUB_TOKEN")
    owner = os.environ.get("GITHUB_OWNER")
    repo  = os.environ.get("GITHUB_REPO")
    if token and owner and repo:
        return {"token": token, "owner": owner, "repo": repo}
    return None


_KNOWN_RETAILERS = ["PLAZA", "ハンズ", "ロフト", "アインズ", "アットコスメ"]
_LEGACY_PATH = "data/pos_data.csv"


def _gh_data_path(retailer: str) -> str:
    """小売店名から GitHub 上のファイルパスを返す"""
    safe = retailer.replace("/", "_").replace("\\", "_").replace(" ", "_")
    return f"data/pos_{safe}.csv"


# ─────────────────────────────────────────────────
# GitHub ストレージ（クラウド用）
# ─────────────────────────────────────────────────

def _gh_headers(cfg: dict) -> dict:
    return {
        "Authorization": f"token {cfg['token']}",
        "Accept": "application/vnd.github.v3+json",
    }


def _parse_gh_csv(content: str) -> pd.DataFrame:
    """GitHub から取得した CSV 文字列を DataFrame に変換する"""
    df = pd.read_csv(io.StringIO(content))
    if not df.empty and "日付" in df.columns:
        df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
        df = df.dropna(subset=["日付"])
        df["売上数量"] = pd.to_numeric(df["売上数量"], errors="coerce").fillna(0).astype(int)
        df["売上金額"] = pd.to_numeric(df["売上金額"], errors="coerce").fillna(0).astype(int)
    return df


def _gh_read_path(path: str) -> tuple[pd.DataFrame, str | None]:
    """指定パスの CSV を GitHub から読み込む。(DataFrame, sha) を返す。"""
    import requests
    cfg = _get_github_config()
    url = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/{path}"
    resp = requests.get(url, headers=_gh_headers(cfg), timeout=15)
    if resp.status_code == 404:
        return pd.DataFrame(), None
    resp.raise_for_status()
    data = resp.json()
    content = base64.b64decode(data["content"]).decode("utf-8-sig")
    return _parse_gh_csv(content), data["sha"]


def _gh_write_path(df: pd.DataFrame, sha: str | None, path: str, message: str) -> None:
    """DataFrame を GitHub の指定パスに CSV として書き込む"""
    import requests
    cfg = _get_github_config()
    df_out = df.copy()
    if "日付" in df_out.columns and pd.api.types.is_datetime64_any_dtype(df_out["日付"]):
        df_out["日付"] = df_out["日付"].dt.strftime("%Y-%m-%d")
    csv_bytes = df_out.to_csv(index=False).encode("utf-8-sig")
    content_b64 = base64.b64encode(csv_bytes).decode()
    payload: dict = {"message": message, "content": content_b64}
    if sha:
        payload["sha"] = sha
    url = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/{path}"
    resp = requests.put(url, headers=_gh_headers(cfg), json=payload, timeout=30)
    resp.raise_for_status()


def _gh_read_retailer(retailer: str) -> tuple[pd.DataFrame, str | None]:
    """小売店専用 CSV を GitHub から読み込む"""
    return _gh_read_path(_gh_data_path(retailer))


def _gh_write_retailer(df: pd.DataFrame, sha: str | None, retailer: str) -> None:
    """小売店専用 CSV を GitHub に書き込む"""
    msg = f"POS data update {retailer} {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    _gh_write_path(df, sha, _gh_data_path(retailer), msg)


# ─────────────────────────────────────────────────
# SQLite ストレージ（ローカル用）
# ─────────────────────────────────────────────────

DEFAULT_DB = Path(__file__).parent / "pos_data.db"

_CREATE_TABLE = """
CREATE TABLE IF NOT EXISTS pos_records (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    年月        TEXT    NOT NULL,
    日付        TEXT    NOT NULL,
    小売店名    TEXT    NOT NULL,
    店舗名      TEXT    NOT NULL,
    ブランド名  TEXT    NOT NULL,
    商品名      TEXT    NOT NULL,
    売上数量    INTEGER NOT NULL DEFAULT 0,
    売上金額    INTEGER NOT NULL DEFAULT 0,
    登録日時    TEXT    NOT NULL
)
"""

_CREATE_INDEX = """
CREATE INDEX IF NOT EXISTS idx_pos_ym_brand_retailer
    ON pos_records (年月, ブランド名, 小売店名)
"""


@contextmanager
def _conn(db_path: Path):
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


def _sqlite_init(db_path: Path = DEFAULT_DB) -> None:
    with _conn(db_path) as con:
        con.execute(_CREATE_TABLE)
        con.execute(_CREATE_INDEX)


# ─────────────────────────────────────────────────
# 公開 API（GitHub / SQLite を自動切替）
# ─────────────────────────────────────────────────

def init_db(db_path: Path = DEFAULT_DB) -> None:
    """初期化（SQLite のみ必要。GitHub は自動作成）"""
    if _get_github_config() is None:
        _sqlite_init(db_path)


def save_records(df: pd.DataFrame, db_path: Path = DEFAULT_DB) -> dict:
    """DataFrame を DB に保存。同じ（年月×ブランド名×小売店名）は上書き。"""
    df = df.copy()
    df["年月"]    = df["日付"].dt.strftime("%Y-%m")
    df["登録日時"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    pairs = df[["年月", "ブランド名", "小売店名"]].drop_duplicates().values.tolist()
    replaced: list[str] = []

    if _get_github_config():
        # ── GitHub モード：小売店ごとに別ファイルに書き込む ──
        retailers_in_df = df["小売店名"].unique().tolist()
        for retailer in retailers_in_df:
            retailer_new = df[df["小売店名"] == retailer].copy()
            existing_df, sha = _gh_read_retailer(retailer)

            # 同一（年月×ブランド×小売店）は削除して上書き
            for ym, brand, ret in pairs:
                if ret != retailer:
                    continue
                if not existing_df.empty:
                    mask = (
                        (existing_df["年月"] == ym) &
                        (existing_df["ブランド名"] == brand) &
                        (existing_df["小売店名"] == retailer)
                    )
                    if mask.any():
                        existing_df = existing_df[~mask]
                        replaced.append(f"{ym} / {brand} / {retailer}")

            retailer_new["日付"] = retailer_new["日付"].dt.strftime("%Y-%m-%d")
            merged = (
                pd.concat([existing_df, retailer_new], ignore_index=True)
                if not existing_df.empty else retailer_new
            )
            _gh_write_retailer(merged, sha, retailer)

        return {"inserted": len(df), "replaced": replaced}

    else:
        # ── SQLite モード ──
        _sqlite_init(db_path)
        df_save = df.copy()
        df_save["日付"] = df_save["日付"].dt.strftime("%Y-%m-%d")

        with _conn(db_path) as con:
            for ym, brand, retailer in pairs:
                existing = con.execute(
                    "SELECT COUNT(*) FROM pos_records WHERE 年月=? AND ブランド名=? AND 小売店名=?",
                    (ym, brand, retailer),
                ).fetchone()[0]
                if existing > 0:
                    con.execute(
                        "DELETE FROM pos_records WHERE 年月=? AND ブランド名=? AND 小売店名=?",
                        (ym, brand, retailer),
                    )
                    replaced.append(f"{ym} / {brand} / {retailer}")

            rows = df_save[["年月", "日付", "小売店名", "店舗名", "ブランド名",
                            "商品名", "売上数量", "売上金額", "登録日時"]].values.tolist()
            con.executemany(
                """INSERT INTO pos_records
                   (年月, 日付, 小売店名, 店舗名, ブランド名, 商品名, 売上数量, 売上金額, 登録日時)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                rows,
            )
        return {"inserted": len(rows), "replaced": replaced}


def load_all(db_path: Path = DEFAULT_DB) -> pd.DataFrame:
    """全レコードを DataFrame で返す"""
    if _get_github_config():
        dfs = []
        for retailer in _KNOWN_RETAILERS:
            df, _ = _gh_read_retailer(retailer)
            if not df.empty:
                dfs.append(df)
        if dfs:
            return pd.concat(dfs, ignore_index=True)
        # レガシーファイル（移行前の統合CSV）にフォールバック
        df, _ = _gh_read_path(_LEGACY_PATH)
        return df

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        df = pd.read_sql(
            "SELECT 年月, 日付, 小売店名, 店舗名, ブランド名, 商品名, 売上数量, 売上金額"
            " FROM pos_records ORDER BY 日付, 小売店名, 店舗名", con,
        )
    if df.empty:
        return df
    df["日付"]    = pd.to_datetime(df["日付"])
    df["売上数量"] = df["売上数量"].astype(int)
    df["売上金額"] = df["売上金額"].astype(int)
    return df


def load_filtered(
    db_path: Path = DEFAULT_DB,
    year_month: str | None = None,
    brand: str | None = None,
    retailer: str | None = None,
) -> pd.DataFrame:
    """条件を指定してレコードを取得する"""
    if _get_github_config() and retailer:
        # 小売店が指定されている場合はその専用ファイルだけ読む（高速）
        df, _ = _gh_read_retailer(retailer)
        if df.empty:
            return df
    else:
        df = load_all(db_path)
        if df.empty:
            return df
        if retailer:
            df = df[df["小売店名"] == retailer]

    if year_month:
        df = df[df["年月"] == year_month]
    if brand:
        df = df[df["ブランド名"] == brand]
    return df.reset_index(drop=True)


def get_summary(db_path: Path = DEFAULT_DB) -> dict:
    """DB の概要情報を返す"""
    df = load_all(db_path)
    if df.empty:
        return {"total_records": 0, "year_months": [], "brands": [], "retailers": [], "date_range": (None, None)}
    return {
        "total_records": len(df),
        "year_months":   sorted(df["年月"].unique().tolist()),
        "brands":        sorted(df["ブランド名"].unique().tolist()),
        "retailers":     sorted(df["小売店名"].unique().tolist()),
        "date_range":    (df["日付"].min().strftime("%Y-%m-%d"), df["日付"].max().strftime("%Y-%m-%d")),
    }


def list_month_brand_pairs(
    db_path: Path = DEFAULT_DB,
    retailer: str | None = None,
) -> list[tuple[str, str]]:
    """（年月, ブランド名）ペアのリストを返す"""
    df = load_filtered(db_path=db_path, retailer=retailer)
    if df.empty:
        return []
    pairs = df[["年月", "ブランド名"]].drop_duplicates().sort_values(["年月", "ブランド名"])
    return list(pairs.itertuples(index=False, name=None))


def delete_by_month_brand(
    year_month: str, brand: str, db_path: Path = DEFAULT_DB,
    retailer: str | None = None,
) -> int:
    """指定した（年月 × ブランド）のレコードを削除。削除件数を返す。"""
    if _get_github_config():
        targets = [retailer] if retailer else _KNOWN_RETAILERS
        total = 0
        for ret in targets:
            df, sha = _gh_read_retailer(ret)
            if df.empty:
                continue
            mask = (df["年月"] == year_month) & (df["ブランド名"] == brand)
            n = int(mask.sum())
            if n > 0:
                _gh_write_retailer(df[~mask], sha, ret)
                total += n
        return total

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        cur = con.execute(
            "DELETE FROM pos_records WHERE 年月=? AND ブランド名=?",
            (year_month, brand),
        )
    return cur.rowcount


def delete_retailer_all(retailer: str, db_path: Path = DEFAULT_DB) -> int:
    """指定した小売店の全レコードを削除。削除件数を返す。"""
    if _get_github_config():
        df, sha = _gh_read_retailer(retailer)
        n = len(df)
        if n > 0 and sha:
            _gh_write_retailer(pd.DataFrame(), sha, retailer)
        return n

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        cur = con.execute("DELETE FROM pos_records WHERE 小売店名=?", (retailer,))
    return cur.rowcount


def delete_all(db_path: Path = DEFAULT_DB) -> int:
    """全レコードを削除。削除件数を返す。"""
    if _get_github_config():
        total = 0
        for retailer in _KNOWN_RETAILERS:
            df, sha = _gh_read_retailer(retailer)
            if not df.empty and sha:
                total += len(df)
                _gh_write_retailer(pd.DataFrame(), sha, retailer)
        return total

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        cur = con.execute("DELETE FROM pos_records")
    return cur.rowcount


def export_csv_bytes(db_path: Path = DEFAULT_DB) -> bytes:
    """全データを UTF-8 BOM 付き CSV のバイト列で返す"""
    df = load_all(db_path)
    if df.empty:
        return "データがありません\n".encode("utf-8-sig")
    df["日付"] = df["日付"].dt.strftime("%Y-%m-%d")
    return df.to_csv(index=False).encode("utf-8-sig")
