"""
pos_db.py — POS データ永続化モジュール

ローカル実行時  : SQLite  (pos_data.db)
Streamlit Cloud : GitHub リポジトリ内の CSV ファイル
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


_GITHUB_DATA_PATH = "data/pos_data.csv"


# ─────────────────────────────────────────────────
# GitHub ストレージ（クラウド用）
# ─────────────────────────────────────────────────

def _gh_headers(cfg: dict) -> dict:
    return {
        "Authorization": f"token {cfg['token']}",
        "Accept": "application/vnd.github.v3+json",
    }


def _gh_read() -> tuple[pd.DataFrame, str | None]:
    """GitHub から CSV を読み込む。(DataFrame, sha) を返す。ファイルなし→空DataFrame"""
    import requests
    cfg = _get_github_config()
    url = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/{_GITHUB_DATA_PATH}"
    resp = requests.get(url, headers=_gh_headers(cfg), timeout=15)
    if resp.status_code == 404:
        return pd.DataFrame(), None
    resp.raise_for_status()
    data = resp.json()
    content = base64.b64decode(data["content"]).decode("utf-8-sig")
    sha = data["sha"]
    df = pd.read_csv(io.StringIO(content))
    if not df.empty and "日付" in df.columns:
        df["日付"] = pd.to_datetime(df["日付"])
        df["売上数量"] = df["売上数量"].astype(int)
        df["売上金額"] = df["売上金額"].astype(int)
    return df, sha


def _gh_write(df: pd.DataFrame, sha: str | None) -> None:
    """DataFrame を GitHub の CSV ファイルに書き込む"""
    import requests
    cfg = _get_github_config()
    df_out = df.copy()
    if "日付" in df_out.columns and pd.api.types.is_datetime64_any_dtype(df_out["日付"]):
        df_out["日付"] = df_out["日付"].dt.strftime("%Y-%m-%d")
    csv_bytes = df_out.to_csv(index=False).encode("utf-8-sig")
    content_b64 = base64.b64encode(csv_bytes).decode()
    payload: dict = {
        "message": f"POS data update {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": content_b64,
    }
    if sha:
        payload["sha"] = sha
    url = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/{_GITHUB_DATA_PATH}"
    resp = requests.put(url, headers=_gh_headers(cfg), json=payload, timeout=20)
    resp.raise_for_status()


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
    df["年月"]   = df["日付"].dt.strftime("%Y-%m")
    df["登録日時"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    pairs = df[["年月", "ブランド名", "小売店名"]].drop_duplicates().values.tolist()
    replaced: list[str] = []

    if _get_github_config():
        # ── GitHub モード ──
        existing_df, sha = _gh_read()
        for ym, brand, retailer in pairs:
            if not existing_df.empty:
                mask = (
                    (existing_df["年月"] == ym) &
                    (existing_df["ブランド名"] == brand) &
                    (existing_df["小売店名"] == retailer)
                )
                if mask.any():
                    existing_df = existing_df[~mask]
                    replaced.append(f"{ym} / {brand} / {retailer}")

        df["日付"] = df["日付"].dt.strftime("%Y-%m-%d")
        merged = pd.concat([existing_df, df], ignore_index=True) if not existing_df.empty else df
        _gh_write(merged, sha)
        return {"inserted": len(df), "replaced": replaced}

    else:
        # ── SQLite モード ──
        _sqlite_init(db_path)
        df_save = df.copy()
        df_save["日付"] = df_save["日付"].dt.strftime("%Y-%m-%d")
        df_save["年月"] = df_save["年月"]

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
        df, _ = _gh_read()
        return df

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        df = pd.read_sql(
            "SELECT 年月, 日付, 小売店名, 店舗名, ブランド名, 商品名, 売上数量, 売上金額"
            " FROM pos_records ORDER BY 日付, 小売店名, 店舗名", con,
        )
    if df.empty:
        return df
    df["日付"]   = pd.to_datetime(df["日付"])
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
    df = load_all(db_path)
    if df.empty:
        return df
    if year_month:
        df = df[df["年月"] == year_month]
    if brand:
        df = df[df["ブランド名"] == brand]
    if retailer:
        df = df[df["小売店名"] == retailer]
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
    df = load_all(db_path)
    if df.empty:
        return []
    if retailer:
        df = df[df["小売店名"] == retailer]
    pairs = df[["年月", "ブランド名"]].drop_duplicates().sort_values(["年月", "ブランド名"])
    return list(pairs.itertuples(index=False, name=None))


def delete_by_month_brand(
    year_month: str, brand: str, db_path: Path = DEFAULT_DB
) -> int:
    """指定した（年月 × ブランド）のレコードを削除。削除件数を返す。"""
    if _get_github_config():
        df, sha = _gh_read()
        if df.empty:
            return 0
        mask = (df["年月"] == year_month) & (df["ブランド名"] == brand)
        n = int(mask.sum())
        if n > 0:
            _gh_write(df[~mask], sha)
        return n

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        cur = con.execute(
            "DELETE FROM pos_records WHERE 年月=? AND ブランド名=?",
            (year_month, brand),
        )
    return cur.rowcount


def delete_all(db_path: Path = DEFAULT_DB) -> int:
    """全レコードを削除。削除件数を返す。"""
    if _get_github_config():
        df, sha = _gh_read()
        n = len(df)
        if n > 0:
            _gh_write(pd.DataFrame(), sha)
        return n

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
