# -*- coding: utf-8 -*-
"""create_db.py
EkiSpert Alignment - committees.db 作成/更新
VBA(func_GetDb.bas) から中間TSVを受け取り、SQLite でDBを構築する。

使い方:
    python create_db.py <tsv_path>

TSV仕様(Unicode/UTF-16LE, TAB区切り, 改行CRLF):
    1行目: comiteename (委員会名)
    2行目: startYear   (年度の開始年, 例: 2026)
    3行目: ヘッダ(18列)
    4行目以降: データ(各メンバー)

DB構成(共通DB1つ):
    fiscal_years      : 年度マスタ(fiscal_year_id PK, start_year, end_year, label UQ)
    committees        : 委員会マスタ(committee_id PK, fiscal_year_id FK, committee_name)
    committee_members : メンバーデータ(committee_id FK)

年度の生成:
    C1(開始年) から label = "{start}-{start+1}" を生成(例: 2026 -> "2026-2027")

年度の重複チェック(新規作成時のみ):
    既存の全年度と期間が重複する場合、新規作成を拒否(終了コード2)。
    例: 2026-2027 既存 -> 2027-2028 は 2027年が重複なので拒否。
        2026-2027 既存 -> 2028-2029 は重複なしので作成OK。

再保存時挙動: 当該 committee_id の committee_members を全削除->全挿入(全置換)
終了コード: 0=成功 / 1=失敗 / 2=年度重複で拒否
"""
import os
import sys
import sqlite3


def get_db_path():
    """DBファイルのパスを解決。環境変数SHARED_ROOT基準。"""
    shared_root = os.environ.get("SHARED_ROOT", "")
    if not shared_root:
        print("[create_db] 環境変数 SHARED_ROOT が未設定", file=sys.stderr)
        return None
    db_dir = os.path.join(shared_root, "05総務", "99 Database")
    if not os.path.isdir(db_dir):
        try:
            os.makedirs(db_dir, exist_ok=True)
            print(f"[create_db] DBフォルダ作成: {db_dir}")
        except OSError as e:
            print(f"[create_db] DBフォルダ作成失敗: {e}", file=sys.stderr)
            return None
    return os.path.join(db_dir, "committees.db")


def read_tsv(tsv_path):
    """TSVを読み込む。戻り値: (comiteename, startYear, header[], data_rows[])。
    1行目=comiteename, 2行目=startYear, 3行目=ヘッダ, 4行目以降=データ。
    """
    with open(tsv_path, "r", encoding="utf-16", newline="") as f:
        lines = f.read().splitlines()
    if len(lines) < 3:
        print("[create_db] TSV行数不足(3行未満)", file=sys.stderr)
        return None, None, None, None
    comiteename = lines[0].strip()
    start_year_str = lines[1].strip()
    header = lines[2].split("\t")
    data_rows = [ln.split("\t") for ln in lines[3:] if ln.strip() != ""]
    return comiteename, start_year_str, header, data_rows


CREATE_FISCAL_YEARS = """
CREATE TABLE IF NOT EXISTS fiscal_years (
    fiscal_year_id INTEGER PRIMARY KEY AUTOINCREMENT,
    start_year     INTEGER NOT NULL,
    end_year       INTEGER NOT NULL,
    label          TEXT    NOT NULL UNIQUE,
    created_at     TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at     TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
)
"""

CREATE_COMMITTEES = """
CREATE TABLE IF NOT EXISTS committees (
    committee_id   INTEGER PRIMARY KEY AUTOINCREMENT,
    fiscal_year_id INTEGER NOT NULL,
    committee_name TEXT    NOT NULL,
    created_at     TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at     TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    UNIQUE (fiscal_year_id, committee_name),
    FOREIGN KEY (fiscal_year_id) REFERENCES fiscal_years(fiscal_year_id)
)
"""

CREATE_COMMITTEE_MEMBERS = """
CREATE TABLE IF NOT EXISTS committee_members (
    member_id         INTEGER PRIMARY KEY AUTOINCREMENT,
    committee_id      INTEGER NOT NULL,
    position          TEXT,
    name              TEXT,
    company           TEXT,
    office            TEXT,
    home_address      TEXT,
    company_phone     TEXT,
    fax               TEXT,
    mobile_phone      TEXT,
    email             TEXT,
    option            TEXT,
    round_trip_fare   TEXT,
    departure_station TEXT,
    via_station_1     TEXT,
    via_station_2     TEXT,
    via_station_3     TEXT,
    via_station_4     TEXT,
    via_station_5     TEXT,
    arrival_station   TEXT,
    created_at        TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    updated_at        TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (committee_id) REFERENCES committees(committee_id)
)
"""

MEMBER_COLUMNS = [
    "position", "name", "company", "office", "home_address",
    "company_phone", "fax", "mobile_phone", "email", "option",
    "round_trip_fare", "departure_station",
    "via_station_1", "via_station_2", "via_station_3", "via_station_4", "via_station_5",
    "arrival_station",
]


def resolve_fiscal_year(cur, start_year):
    """年度を解決。戻り値: (fiscal_year_id, label) または (None, None)。
    既存年度があれば再利用。未存在なら重複チェック→OKなら新規作成。
    重複ありなら (None, None) を返し、呼出元で終了コード2を返す。
    """
    start = int(start_year)
    end = start + 1
    label = f"{start}-{end}"

    # 既存年度を label で検索
    cur.execute("SELECT fiscal_year_id FROM fiscal_years WHERE label = ?", (label,))
    row = cur.fetchone()
    if row:
        fiscal_year_id = row[0]
        cur.execute(
            "UPDATE fiscal_years SET updated_at = datetime('now','localtime') WHERE fiscal_year_id = ?",
            (fiscal_year_id,),
        )
        print(f"[create_db] 既存年度を再利用: label={label} / fiscal_year_id={fiscal_year_id}")
        return fiscal_year_id, label

    # 新規作成: 既存の全年度と期間重複チェック
    cur.execute("SELECT start_year, end_year, label FROM fiscal_years")
    for exist_start, exist_end, exist_label in cur.fetchall():
        # 重複条件: new_start <= exist_end AND new_end >= exist_start
        if start <= exist_end and end >= exist_start:
            print(
                f"[create_db] 年度重複で拒否: 新規{label} は 既存{exist_label}({exist_start}-{exist_end}) と期間重複",
                file=sys.stderr,
            )
            return None, None

    # 重複なし -> 新規挿入
    cur.execute(
        "INSERT INTO fiscal_years (start_year, end_year, label) VALUES (?, ?, ?)",
        (start, end, label),
    )
    fiscal_year_id = cur.lastrowid
    print(f"[create_db] 新規年度を挿入: label={label} / fiscal_year_id={fiscal_year_id}")
    return fiscal_year_id, label


def main():
    if len(sys.argv) < 2:
        print("[create_db] 引数不足: python create_db.py <tsv_path>", file=sys.stderr)
        return 1

    tsv_path = sys.argv[1]
    if not os.path.isfile(tsv_path):
        print(f"[create_db] TSVが見つかりません: {tsv_path}", file=sys.stderr)
        return 1

    db_path = get_db_path()
    if db_path is None:
        return 1

    comiteename, start_year_str, header, data_rows = read_tsv(tsv_path)
    if comiteename is None or not comiteename:
        print("[create_db] 委員会名が空", file=sys.stderr)
        return 1
    if not start_year_str or not start_year_str.isdigit():
        print(f"[create_db] 年度の開始年が不正: {start_year_str}", file=sys.stderr)
        return 1

    print(f"[create_db] 委員会名: {comiteename}")
    print(f"[create_db] 年度開始年: {start_year_str}")
    print(f"[create_db] データ行数: {len(data_rows) if data_rows else 0}")
    print(f"[create_db] DBパス: {db_path}")

    try:
        conn = sqlite3.connect(db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.cursor()

        # テーブル作成(存在しない場合のみ)
        cur.execute(CREATE_FISCAL_YEARS)
        cur.execute(CREATE_COMMITTEES)
        cur.execute(CREATE_COMMITTEE_MEMBERS)

        # 年度の解決(既存再利用 or 新規作成 or 重複拒否)
        fiscal_year_id, label = resolve_fiscal_year(cur, start_year_str)
        if fiscal_year_id is None:
            # 重複で拒否
            conn.rollback()
            return 2

        # committees: (fiscal_year_id, committee_name) で検索->なければ挿入->committee_id 取得
        cur.execute(
            "SELECT committee_id FROM committees WHERE fiscal_year_id = ? AND committee_name = ?",
            (fiscal_year_id, comiteename),
        )
        row = cur.fetchone()
        if row:
            committee_id = row[0]
            cur.execute(
                "UPDATE committees SET updated_at = datetime('now','localtime') WHERE committee_id = ?",
                (committee_id,),
            )
            print(f"[create_db] 既存委員会を更新: committee_id={committee_id}")
        else:
            cur.execute(
                "INSERT INTO committees (fiscal_year_id, committee_name) VALUES (?, ?)",
                (fiscal_year_id, comiteename),
            )
            committee_id = cur.lastrowid
            print(f"[create_db] 新規委員会を挿入: committee_id={committee_id}")

        # committee_members: 当該 committee_id の既存レコードを全削除(全置換)
        cur.execute("DELETE FROM committee_members WHERE committee_id = ?", (committee_id,))
        deleted = cur.rowcount
        print(f"[create_db] 既存メンバー削除: {deleted}件")

        # committee_members: 全行挿入
        if data_rows:
            placeholders = ", ".join(["?"] * len(MEMBER_COLUMNS))
            cols = ", ".join(MEMBER_COLUMNS)
            sql = f"INSERT INTO committee_members (committee_id, {cols}) VALUES (?, {placeholders})"
            inserted = 0
            for r in data_rows:
                vals = list(r) + [""] * (len(MEMBER_COLUMNS) - len(r))
                vals = vals[: len(MEMBER_COLUMNS)]
                vals = ["" if v is None else v for v in vals]
                cur.execute(sql, [committee_id] + vals)
                inserted += 1
            print(f"[create_db] メンバー挿入: {inserted}件")
        else:
            print("[create_db] メンバー挿入: 0件(データ行なし)")

        conn.commit()
        print("[create_db] 完了")
        return 0
    except Exception as e:
        print(f"[create_db] エラー: {e}", file=sys.stderr)
        return 1
    finally:
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())