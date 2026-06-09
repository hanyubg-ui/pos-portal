"""
pos_db.py — POS データ永続化モジュール

ローカル実行時  : SQLite  (pos_data.db)
Streamlit Cloud : GitHub リポジトリ内の CSV ファイル（小売店×月ごとに分割）
                  data/pos_PLAZA_2026-04.csv
                  data/pos_ロフト_2026-04.csv
                  data/pos_ロフト_2026-05.csv  ...
                  secrets.toml に以下を設定:
                    GITHUB_TOKEN = "ghp_xxxx"
                    GITHUB_OWNER = "あなたのGitHubユーザー名"
                    GITHUB_REPO  = "pos-portal"
"""

import base64
import io
import os
import re
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

# 月別ファイル名パターン: data/pos_{retailer}_{YYYY-MM}.csv
_YM_PATTERN = re.compile(r"^pos_(.+)_(\d{4}-\d{2})\.csv$")


def _gh_ym_path(retailer: str, ym: str) -> str:
    """小売店名＋年月から GitHub 上のファイルパスを返す"""
    safe = retailer.replace("/", "_").replace("\\", "_").replace(" ", "_")
    return f"data/pos_{safe}_{ym}.csv"


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
    if not content.strip():
        return pd.DataFrame()
    try:
        df = pd.read_csv(io.StringIO(content))
    except Exception:
        return pd.DataFrame()
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
    if "日付" in df_out.columns:
        df_out["日付"] = df_out["日付"].apply(
            lambda x: x.strftime("%Y-%m-%d") if hasattr(x, "strftime") else str(x)[:10]
        )
    csv_bytes = df_out.to_csv(index=False).encode("utf-8-sig")
    content_b64 = base64.b64encode(csv_bytes).decode()
    payload: dict = {"message": message, "content": content_b64}
    if sha:
        payload["sha"] = sha
    url = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/{path}"
    resp = requests.put(url, headers=_gh_headers(cfg), json=payload, timeout=30)
    resp.raise_for_status()


def _gh_list_data_files() -> list[dict]:
    """data/ ディレクトリのファイル一覧を返す（name, path, sha を含む）"""
    import requests
    cfg = _get_github_config()
    url = f"https://api.github.com/repos/{cfg['owner']}/{cfg['repo']}/contents/data"
    resp = requests.get(url, headers=_gh_headers(cfg), timeout=15)
    if resp.status_code == 404:
        return []
    resp.raise_for_status()
    return [f for f in resp.json() if f.get("type") == "file" and f["name"].endswith(".csv")]


def _gh_read_ym(retailer: str, ym: str) -> tuple[pd.DataFrame, str | None]:
    """小売店×年月の CSV を読み込む"""
    return _gh_read_path(_gh_ym_path(retailer, ym))


def _gh_write_ym(df: pd.DataFrame, sha: str | None, retailer: str, ym: str) -> None:
    """小売店×年月の CSV を書き込む"""
    msg = f"POS update {retailer} {ym} {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    _gh_write_path(df, sha, _gh_ym_path(retailer, ym), msg)


def _gh_load_retailer_all(retailer: str) -> pd.DataFrame:
    """小売店の全月データを読み込んで結合する"""
    files = _gh_list_data_files()
    safe = retailer.replace("/", "_").replace("\\", "_").replace(" ", "_")
    prefix = f"pos_{safe}_"
    dfs = []
    for f in files:
        if f["name"].startswith(prefix) and _YM_PATTERN.match(f["name"]):
            df, _ = _gh_read_path(f["path"])
            if not df.empty:
                dfs.append(df)
    if not dfs:
        # 旧形式フォールバック: data/pos_{retailer}.csv
        df, _ = _gh_read_path(f"data/pos_{safe}.csv")
        return df
    return pd.concat(dfs, ignore_index=True)


def _gh_migrate_if_needed(retailer: str) -> None:
    """旧形式（月混在）ファイルを月別ファイルに分割する"""
    import requests
    cfg = _get_github_config()
    safe = retailer.replace("/", "_").replace("\\", "_").replace(" ", "_")
    old_path = f"data/pos_{safe}.csv"

    old_df, old_sha = _gh_read_path(old_path)
    if old_df.empty or old_sha is None:
        return

    if "年月" not in old_df.columns:
        old_df["年月"] = old_df["日付"].dt.strftime("%Y-%m")

    for ym, group in old_df.groupby("年月"):
        new_path = _gh_ym_path(retailer, ym)
        existing, sha = _gh_read_path(new_path)
        if existing.empty:  # 新ファイルがなければ移行
            _gh_write_ym(group.reset_index(drop=True), sha, retailer, ym)

    # 旧ファイルを空にして事実上無効化
    _gh_write_path(pd.DataFrame(), old_sha, old_path,
                   f"Migrated {retailer} to per-month files")


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


def _is_monthly_summary(group_df: pd.DataFrame) -> bool:
    """全行の日付が同一かつ月末日なら月次集計データと判定する（LFPOSMON等）"""
    import calendar as _cal
    if group_df.empty:
        return False
    try:
        dates = pd.to_datetime(group_df["日付"]).dt.normalize()
        if dates.nunique() != 1:
            return False
        d = dates.iloc[0]
        return d.day == _cal.monthrange(d.year, d.month)[1]
    except Exception:
        return False


def save_records(df: pd.DataFrame, db_path: Path = DEFAULT_DB) -> dict:
    """DataFrame を DB に保存。月ごとに別ファイル。
    ・日別データ（LFPO等）: 同一ブランドのみ上書き
    ・月次集計データ（LFPOSMON等）: その月の全データを完全上書き（日別と重複しない）
    """
    df = df.copy()
    df["年月"]    = df["日付"].dt.strftime("%Y-%m")
    df["登録日時"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    replaced: list[str] = []

    if _get_github_config():
        # ── GitHub モード：小売店×年月ごとに別ファイルに書き込む ──
        for (retailer, ym), group in df.groupby(["小売店名", "年月"]):
            group = group.copy()
            existing_df, sha = _gh_read_ym(retailer, ym)

            if _is_monthly_summary(group):
                # 月次集計データ → その月の全データを完全上書き（日別データも削除）
                if not existing_df.empty:
                    replaced.append(f"{ym} / 全データ / {retailer}（月次集計で上書き）")
                merged = group
            elif not existing_df.empty:
                # 日別データ → 同一ブランドのみ上書き
                brands_in_new = group["ブランド名"].unique().tolist()
                for brand in brands_in_new:
                    b_mask = existing_df["ブランド名"] == brand
                    if b_mask.any():
                        existing_df = existing_df[~b_mask]
                        replaced.append(f"{ym} / {brand} / {retailer}")
                merged = pd.concat([existing_df, group], ignore_index=True)
            else:
                merged = group

            _gh_write_ym(merged, sha, retailer, ym)

        return {"inserted": len(df), "replaced": replaced}

    else:
        # ── SQLite モード ──
        _sqlite_init(db_path)
        pairs = df[["年月", "ブランド名", "小売店名"]].drop_duplicates().values.tolist()
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
        files = _gh_list_data_files()
        dfs = []
        for f in files:
            if _YM_PATTERN.match(f["name"]):
                df, _ = _gh_read_path(f["path"])
                if not df.empty:
                    dfs.append(df)
        return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

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
    if _get_github_config():
        if retailer and year_month:
            # 月別ファイルを優先、なければ全月ファイルを読んでフィルター
            df, _ = _gh_read_ym(retailer, year_month)
            if df.empty:
                df = _gh_load_retailer_all(retailer)
        elif retailer:
            df = _gh_load_retailer_all(retailer)
        else:
            df = load_all(db_path)

        if df.empty:
            return df
        if "日付" in df.columns and not pd.api.types.is_datetime64_any_dtype(df["日付"]):
            df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
        if retailer and "小売店名" in df.columns:
            df = df[df["小売店名"] == retailer]
        if year_month:
            if "年月" not in df.columns and "日付" in df.columns:
                df["年月"] = df["日付"].dt.strftime("%Y-%m")
            if "年月" in df.columns:
                df = df[df["年月"] == year_month]
        if brand and "ブランド名" in df.columns:
            df = df[df["ブランド名"] == brand]
        return df.reset_index(drop=True)

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        conditions = []
        params = []
        if year_month:
            conditions.append("年月=?"); params.append(year_month)
        if brand:
            conditions.append("ブランド名=?"); params.append(brand)
        if retailer:
            conditions.append("小売店名=?"); params.append(retailer)
        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        df = pd.read_sql(
            f"SELECT 年月, 日付, 小売店名, 店舗名, ブランド名, 商品名, 売上数量, 売上金額"
            f" FROM pos_records {where} ORDER BY 日付, 小売店名, 店舗名",
            con, params=params,
        )
    if df.empty:
        return df
    df["日付"]    = pd.to_datetime(df["日付"])
    df["売上数量"] = df["売上数量"].astype(int)
    df["売上金額"] = df["売上金額"].astype(int)
    return df


def get_summary(db_path: Path = DEFAULT_DB) -> dict:
    """DB の概要情報を返す"""
    df = load_all(db_path)
    if df.empty:
        return {"total_records": 0, "year_months": [], "brands": [], "retailers": [], "date_range": (None, None)}
    if "年月" not in df.columns:
        df["年月"] = df["日付"].dt.strftime("%Y-%m")
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
    if "年月" not in df.columns:
        df["年月"] = df["日付"].dt.strftime("%Y-%m")
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
            df, sha = _gh_read_ym(ret, year_month)
            if df.empty:
                continue
            mask = (df["ブランド名"] == brand)
            n = int(mask.sum())
            if n > 0:
                _gh_write_ym(df[~mask].reset_index(drop=True), sha, ret, year_month)
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
        files = _gh_list_data_files()
        safe = retailer.replace("/", "_").replace("\\", "_").replace(" ", "_")
        prefix = f"pos_{safe}_"
        total = 0
        for f in files:
            if f["name"].startswith(prefix) and _YM_PATTERN.match(f["name"]):
                df, sha = _gh_read_path(f["path"])
                total += len(df)
                if sha:
                    # ファイルを空にする（GitHub は空ファイルでも削除はAPIが別）
                    _gh_write_path(pd.DataFrame(), sha, f["path"],
                                   f"Delete all {retailer} data")
        return total

    _sqlite_init(db_path)
    with _conn(db_path) as con:
        cur = con.execute("DELETE FROM pos_records WHERE 小売店名=?", (retailer,))
    return cur.rowcount


def delete_all(db_path: Path = DEFAULT_DB) -> int:
    """全レコードを削除。削除件数を返す。"""
    if _get_github_config():
        files = _gh_list_data_files()
        total = 0
        for f in files:
            if _YM_PATTERN.match(f["name"]):
                df, sha = _gh_read_path(f["path"])
                total += len(df)
                if sha:
                    _gh_write_path(pd.DataFrame(), sha, f["path"], "Delete all POS data")
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
    df_out = df.copy()
    df_out["日付"] = df_out["日付"].apply(
        lambda x: x.strftime("%Y-%m-%d") if hasattr(x, "strftime") else str(x)[:10]
    )
    return df_out.to_csv(index=False).encode("utf-8-sig")
