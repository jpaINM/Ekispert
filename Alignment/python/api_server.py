# -*- coding: utf-8 -*-
"""api_server.py
EkiSpert Alignment - DB参照APIサーバ
committees.db を読み込み、フロントエンドにJSONを提供する。

起動:
    python api_server.py
    (ブラウザで http://localhost:8001/ を開く)
"""
from __future__ import annotations

import os
import sqlite3
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
FRONTEND_DIR = PROJECT_DIR / "frontend"
INDEX_HTML = FRONTEND_DIR / "index.html"
ENV_FILE = BASE_DIR / ".env"

# .env を読み込む（既に環境変数が設定されている場合は上書きしない）
load_dotenv(ENV_FILE, override=False)


def get_ekispert_access_key() -> str:
    """駅すぱあと API のアクセスキーを返す（サーバー側でのみ使用）。

    ブラウザ（フロントエンド）にはこのキーを渡さないこと。
    プロキシエンドポイント経由で駅すぱあと API を呼ぶ場合のみ使用する。
    未設定時は空文字を返す。
    """
    return os.environ.get("EKISPERT_ACCESS_KEY", "").strip()


def get_db_path() -> Path:
    shared_root = os.environ.get("SHARED_ROOT", "")
    if not shared_root:
        raise RuntimeError("環境変数 SHARED_ROOT が未設定")
    return Path(shared_root) / "05総務" / "99 Database" / "committees.db"


app = FastAPI(title="EkiSpert Alignment Server", version="1.0.0")

if FRONTEND_DIR.is_dir():
    app.mount("/frontend", StaticFiles(directory=FRONTEND_DIR), name="frontend")


from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request


class NoCacheMiddleware(BaseHTTPMiddleware):
    """開発中のキャッシュ問題を防止するため /frontend 配下に no-cache を付与。"""

    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        if request.url.path.startswith("/frontend"):
            response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        return response


app.add_middleware(NoCacheMiddleware)


@app.get("/", response_class=FileResponse)
def index() -> FileResponse:
    return FileResponse(INDEX_HTML, headers={
        "Cache-Control": "no-cache, no-store, must-revalidate",
    })


@app.get("/api/committees", response_class=JSONResponse)
def get_committees() -> dict[str, Any]:
    """committees.db を読み込み、fiscal_years -> committees -> members 構造のJSONを返す。"""
    try:
        db_path = get_db_path()
    except RuntimeError as e:
        return JSONResponse(status_code=500, content={"ok": False, "error": {"code": "DB_PATH_ERROR", "message": str(e)}})

    if not db_path.is_file():
        return JSONResponse(status_code=404, content={"ok": False, "error": {"code": "DB_NOT_FOUND", "message": f"DBが見つかりません: {db_path}"}})

    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    try:
        fiscal_years = []
        for fy in cur.execute("SELECT * FROM fiscal_years ORDER BY start_year").fetchall():
            committees = []
            for cm in cur.execute(
                "SELECT * FROM committees WHERE fiscal_year_id = ? ORDER BY committee_name",
                (fy["fiscal_year_id"],),
            ).fetchall():
                members = []
                for m in cur.execute(
                    "SELECT * FROM committee_members WHERE committee_id = ? ORDER BY member_id",
                    (cm["committee_id"],),
                ).fetchall():
                    via_stations = []
                    for i in range(1, 6):
                        val = m[f"via_station_{i}"]
                        if val:
                            via_stations.append(val)

                    members.append({
                        "position": m["position"] or "",
                        "name": m["name"] or "",
                        "departure_station": m["departure_station"] or "",
                        "via_stations": via_stations,
                        "arrival_station": m["arrival_station"] or "",
                    })

                committees.append({
                    "name": cm["committee_name"] or "",
                    "members": members,
                })

            fiscal_years.append({
                "label": fy["label"] or "",
                "committees": committees,
            })

        return {"ok": True, "fiscal_years": fiscal_years}
    except Exception as e:
        return JSONResponse(status_code=500, content={"ok": False, "error": {"code": "DB_QUERY_ERROR", "message": str(e)}})
    finally:
        conn.close()


if __name__ == "__main__":
    import uvicorn

    # 起動時にアクセスキーの設定状況を確認（キー自体は表示しない）
    if get_ekispert_access_key():
        print("[EkiSpert] 駅すぱあとAPI連携: 有効（アクセスキー設定済み）")
    else:
        print("[EkiSpert] 駅すぱあとAPI連携: 無効（.env の EKISPERT_ACCESS_KEY が未設定）")

    uvicorn.run(app, host="127.0.0.1", port=8001)