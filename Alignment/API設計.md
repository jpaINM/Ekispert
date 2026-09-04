# EkiSpert Alignment - 駅すぱあとAPI連携 設計書

## 1. 目的

`EkiSpert Alignment` から「駅すぱあと API」を安全に呼び出すための、
アクセスキー管理方針とサーバー側プロキシの基本設計を定める。

本フェーズでは **アクセスキーの管理枠（場所＋読込機能）** のみを用意する。
実際のAPI呼び出しロジック・フロントエンドの本実装は後続フェーズで行う。

## 2. アクセスキー管理方針（方式A：サーバー側管理）

### 2.1 採用方式

**サーバー側（FastAPI）でアクセスキーを管理し、ブラウザにはキーを一切渡さない。**

### 2.2 採用理由

| 観点 | サーバー側管理（方式A） | フロントエンド管理（方式B） |
| --- | --- | --- |
| ブラウザでの露出 | なし | 開発者ツールで参照可能 |
| git への混入 | .gitignore で防止 | 同左 |
| 秘密度 | 高 | 中 |

「秘密度高く」という要件を満たすため、ブラウザにキーが見えない方式Aを採用する。

### 2.3 キーの保存場所

```
EkiSpert/Alignment/python/.env
```

形式（dotenv形式）:

```
EKISPERT_ACCESS_KEY=実際のアクセスキー
```

- `.env.example` はコミット用テンプレート（空値）
- `.env` は実キーを記入する実ファイルで **.gitignore により管理対象外**

### 2.4 キーの読込

`api_server.py` 起動時に `python-dotenv` で `.env` を読み込み、
環境変数 `EKISPERT_ACCESS_KEY` から取得する。

```python
from dotenv import load_dotenv
load_dotenv(BASE_DIR / ".env", override=False)

def get_ekispert_access_key() -> str:
    return os.environ.get("EKISPERT_ACCESS_KEY", "").strip()
```

- `override=False`: 既に環境変数が設定されている場合は `.env` で上書きしない
  （CI/本番では環境変数で直接指定することを想定）

### 2.5 セキュリティ要件

1. **アクセスキーをフロントエンドに返すエンドポイントは作らない**
   - `/api/ekispert/key` のような露出エンドポイントは禁止
2. **`.env` はコミットしない**
   - `.gitignore` で `python/.env` を除外
   - コミットされるのは `.env.example`（空値）のみ
3. **ログにキーを出力しない**
   - 起動時メッセージは「設定済み/未設定」の判定のみ（キー値は非表示）

## 3. セットアップ手順（ユーザー向け）

1. `python/.env.example` を `python/.env` にコピーする
   （本リポジトリでは既に空の `.env` も同梱済み）
2. `python/.env` を開き、`EKISPERT_ACCESS_KEY=` の右に実際のアクセスキーを記入する
   ```
   EKISPERT_ACCESS_KEY=XXXXXXXXXXXXXXXXXXXXXXXX
   ```
3. 依存ライブラリをインストールする
   ```
   pip install -r python/requirements.txt
   ```
4. APIサーバを起動する
   ```
   python python/api_server.py
   ```
   - キー設定済み: `[EkiSpert] 駅すぱあとAPI連携: 有効（アクセスキー設定済み）`
   - キー未設定: `[EkiSpert] 驻すぱあとAPI連携: 無効（.env の EKISPERT_ACCESS_KEY が未設定）`

## 4. 将来構想：サーバー側プロキシ（後続フェーズ）

ブラウザが直接「駅すぱあと API」を叩くのではなく、
FastAPI が **プロキシエンドポイント** として中継する。

### 4.1 想定エンドポイント群

| エンドポイント | 役割 |
| --- | --- |
| `GET /api/ekispert/station` | 駅名検索（駅すぱあとAPIの station を中継） |
| `GET /api/ekispert/search` | 経路探索（course/search を中継） |
| `GET /api/ekispert/timetable/station` | 駅時刻表 |
| その他 | 必要に応じて追加 |

### 4.2 呼出フロー

```
ブラウザ --[検索条件]--> FastAPI プロキシ --[条件+アクセスキー]--> 駅すぱあとAPI
                                                                      |
ブラウザ <--[結果JSON]-- FastAPI プロキシ <--[結果JSON]--------------+
```

- ブラウザ → FastAPI の通信にアクセスキーは含まれない
- FastAPI → 駅すぱあとAPI の通信でのみ `get_ekispert_access_key()` を使用
- 結果JSONはGUIサンプル（`EkiSpert/GUI`）の描画形式を参考に、
  Alignment フロントエンドで自前描画する

### 4.3 フロントエンド

- `frontend/app.js` の「駅すぱあと連携」ボタンは現状モック
- プロキシエンドポイント実装後に、ボタン押下で `/api/ekispert/*` を呼ぶよう差し替え
- GUIサンプルコンポーネント（`expGuiXxx.js`）はフロントエンド用のため、
  方式Aではこれらを直接は使わず、結果JSONを自前で描画する方針

## 5. 現フェーズのスコープ

### 実施済み

- [x] `python/.env.example` 作成（コミット用テンプレート）
- [x] `python/.env` 作成（空テンプレート、ユーザーが後でキー記入）
- [x] `.gitignore` 新設（`python/.env` と `__pycache__/` を除外）
- [x] `requirements.txt` に `python-dotenv` 追加
- [x] `api_server.py` に `load_dotenv` と `get_ekispert_access_key()` を追加
- [x] 起動時のキー有無チェックメッセージ追加
- [x] 本設計書 `API設計.md` 作成

### 今フェーズでは実施しない（後続フェーズ）

- [ ] プロキシエンドポイント（`/api/ekispert/*`）の実装
- [ ] 駅すぱあとAPI呼出ロジック
- [ ] フロントエンド「駅すぱあと連携」ボタンの本実装
- [ ] 経路・時刻表の描画
