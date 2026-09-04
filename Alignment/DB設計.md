# EkiSpert Alignment - DB設計（年度・委員会・メンバーデータ）

## 1. 概要

本資料は、Excel「DB」シートの表データをRDBに保存する際の設計をまとめる。
本エクセルマクロは **Excel と SQLite DB のインターフェイス** として位置づける。
EkiSpert は実質的にDBとして扱い、本マクロで委員会情報をDBに保存する。

- 1つのワークブックには複数の「DB」シートが存在し、各シートが1つの委員会を表す
- 各「DB」シートの **A1セル** に委員会名、**C1セル** に年度の開始年が格納されている
- **A3:R{lastRow}** が表データ（3行目=ヘッダ, 4行目以降=データ）で、1行=1メンバー

### データの紐づきイメージ

```
[年度 2026-2027] ──┬── [委員会A] ──┬── [メンバー1: ...]
                    │               ├── [メンバー2: ...]
                    │               └── [メンバーN: ...]
                    │
                    └── [委員会B] ──┬── [メンバー1: ...]
                                    └── [メンバー2: ...]

[年度 2028-2029] ──┬── [委員会A] ──┬── [メンバー1: ...]
                    :               :
```

これは **3階層（1年度 : N委員会 : Nメンバー）** のリレーションであり、
`fiscal_years`（年度）→ `committees`（委員会）→ `committee_members`（メンバー）の3テーブルで表現する。

### 採用技術

- **DBMS**: SQLite
- **連携方式**: VBA ⇄ Python ⇄ SQLite（VBAから直接SQLite操作する手段がないため）
- **DBファイル**: 共通DB1つに全年度・全委員会を集約
- **配置先**: `sharedRoot\05総務\99 Database\committees.db`
- **Python実行環境**: Python 3.x（標準 `sqlite3` モジュール使用）

---

## 2. ヘッダと英語カラム名の対応表

DBシートの3行目（ヘッダ）と、DBカラム名（英語）の対応を以下に示す。
A1の委員会名は `committees.committee_name`、C1の開始年は `fiscal_years` に格納する。

| #   | 列   | 日本語ヘッダ | 英語カラム名              | 格納先テーブル           | 備考            |
| --- | --- | ------ | ------------------- | ----------------- | ------------- |
| 0   | A1  | 委員会名   | `committee_name`    | committees        | A1セル。NN |
| 0'  | C1  | 年度開始年  | `start_year`        | fiscal_years      | C1セル。数値(例:2026) |
| 1   | A3  | 役職     | `position`          | committee_members |               |
| 2   | B   | 氏名     | `name`              | committee_members |               |
| 3   | C   | 所属会社   | `company`           | committee_members |               |
| 4   | D   | 事業所    | `office`            | committee_members | Excelエラー時は空白 |
| 5   | E   | 自宅住所   | `home_address`      | committee_members | Excelエラー時は空白 |
| 6   | F   | 会社電話番号 | `company_phone`     | committee_members |               |
| 7   | G   | FAX番号  | `fax`               | committee_members |               |
| 8   | H   | 携帯番号   | `mobile_phone`      | committee_members |               |
| 9   | I   | E-mail | `email`             | committee_members |               |
| 10  | J   | オプション  | `option`            | committee_members |               |
| 11  | K   | 旅費（往復） | `round_trip_fare`   | committee_members | 往復運賃          |
| 12  | L   | 始発     | `departure_station` | committee_members | 駅（EkiSpert入力） |
| 13  | M   | 経由1    | `via_station_1`     | committee_members | 経由駅1          |
| 14  | N   | 経由2    | `via_station_2`     | committee_members | 経由駅2          |
| 15  | O   | 経由3    | `via_station_3`     | committee_members | 経由駅3          |
| 16  | P   | 経由4    | `via_station_4`     | committee_members | 経由駅4          |
| 17  | Q   | 経由5    | `via_station_5`     | committee_members | 経由駅5          |
| 18  | R   | 到着     | `arrival_station`   | committee_members | 駅（EkiSpert入力） |

> **office / home_address のエラー処理**: Excel上でセルがエラー値（`#N/A` 等）の場合、
> VBA 側（`func_GetDb.bas`）で `IsError()` 判定し **空白に置換** してから DB に保存する。

---

## 3. テーブル構成

| テーブル | 役割 | レコードの粒度 |
|---|---|---|
| `fiscal_years` | 年度マスタ | 年度1件につき1レコード（2年区切り） |
| `committees` | 委員会マスタ | 委員会1件につき1レコード（年度内） |
| `committee_members` | 委員会別メンバーデータ | メンバー1人につき1レコード |

---

## 4. ER図（テキストベース）

```
┌─────────────────┐
│  fiscal_years   │ 1
│ (年度マスタ)     │
│ fiscal_year_id  │────────┐
│ start_year      │        │ N
│ end_year        │        ▼
│ label          │   ┌─────────────────┐ 1
└─────────────────┘   │   committees    │
                       │ (委員会マスタ)   │
                       │ committee_id    │────────┐
                       │ fiscal_year_id  │        │ N
                       │ committee_name  │        ▼
                       └─────────────────┘   ┌─────────────────────┐
                                             │ committee_members   │
                                             │ (メンバーデータ)     │
                                             │ member_id           │
                                             │ committee_id(FK)    │
                                             │ position 〜 arrival  │
                                             └─────────────────────┘
```

- `committees.fiscal_year_id` → `fiscal_years.fiscal_year_id`（N:1）
- `committee_members.committee_id` → `committees.committee_id`（N:1）
- 1年度に複数委員会、1委員会に複数メンバー

---

## 5. テーブル定義（SQLite型）

### 5-1. `fiscal_years`（年度マスタ）

C1の開始年から生成される年度（2年区切り）。同じ label の重複を許さない。

| カラム名 | SQLite型 | 制約 | 説明 |
|---|---|---|---|
| fiscal_year_id | INTEGER | PK, AUTOINCREMENT, NN | 年度ID（DB内部キー） |
| start_year | INTEGER | NN | 開始年（← C1, 例: 2026） |
| end_year | INTEGER | NN | 終了年（= start_year + 1, 例: 2027） |
| label | TEXT | NN, UNIQUE | 表示用（例: "2026-2027"） |
| created_at | TEXT | NN, DEFAULT datetime | レコード作成日時 |
| updated_at | TEXT | NN, DEFAULT datetime | レコード更新日時 |

> label は `{start_year}-{end_year}` 形式（例: 2026 → "2026-2027"）。
> 同一 label での重複登録は `UNIQUE` 制約で防止。2回目以降は既存の `fiscal_year_id` を再利用する。

### 5-2. `committees`（委員会マスタ）

A1の委員会名を格納するテーブル。`fiscal_year_id` で所属年度を紐付ける。
同一年度内で委員会名の重複を許さない。

| カラム名 | SQLite型 | 制約 | 説明 |
|---|---|---|---|
| committee_id | INTEGER | PK, AUTOINCREMENT, NN | 委員会ID（DB内部キー） |
| fiscal_year_id | INTEGER | FK→fiscal_years.fiscal_year_id, NN | 所属年度ID（外部キー） |
| committee_name | TEXT | NN | 委員会名（← A1） |
| created_at | TEXT | NN, DEFAULT datetime | レコード作成日時 |
| updated_at | TEXT | NN, DEFAULT datetime | レコード更新日時 |
| | | UNIQUE(fiscal_year_id, committee_name) | 年度内で委員会名一意 |

> 同一委員会名でも年度が異なれば別レコード（例: 2026-2027の委員会A と 2028-2029の委員会A は別）。

### 5-3. `committee_members`（メンバーデータ）

A3:R{lastRow} の各行（1メンバー）を1レコードとして格納する子テーブル。
`committee_id` で所属委員会を紐付ける。

| カラム名 | SQLite型 | 制約 | 説明 |
|---|---|---|---|
| member_id | INTEGER | PK, AUTOINCREMENT, NN | メンバーID（DB内部キー） |
| committee_id | INTEGER | FK→committees.committee_id, NN | 所属委員会ID（外部キー） |
| position 〜 arrival_station | TEXT | | 上記18カラム（§2対応表） |
| created_at | TEXT | NN, DEFAULT datetime | レコード作成日時 |
| updated_at | TEXT | NN, DEFAULT datetime | レコード更新日時 |

> ※ SQLite は動的型付けのため TEXT で統一（数値も TEXT として保持）。

---

## 6. 保存フロー（VBA ⇄ Python 連携）

1つの「DB」シート（1委員会分）を保存する流れ:

```
1. [VBA] CollectCellData で comiteename(A1), startYear(C1), rawData(A3:R{lastRow}) を取得
       （氏名空行はカット済）
2. [VBA] 2段階 MsgBox 確認
       ├ "DBへ同期しますか？" → No で中止
       └ "本当に同期しますか？" → No で中止
3. [VBA] rawData + comiteename + startYear を中間TSV(%TEMP%) に書き出し
       ├ 1行目: comiteename
       ├ 2行目: startYear(年度の開始年)
       ├ 3行目: ヘッダ(18列)
       └ 4行目以降: データ(office/home_address は Excelエラー時 空白化済)
4. [VBA] python.exe create_db.py <tsv_path> を同期待ち実行
5. [Python] TSV 読込 → SQLite で committees.db 作成/更新
       ├ fiscal_years テーブル: label で検索
       │    ├ 存在 → fiscal_year_id 再利用 + updated_at 更新（既存年度）
       │    └ なし → 重複チェック（既存の全年度と期間重複しないか）
       │         ├ 重複あり → 拒否(終了コード2)
       │         └ 重複なし → 新規挿入して fiscal_year_id 取得
       ├ committees テーブル: (fiscal_year_id, committee_name) で検索
       │    ├ 存在 → committee_id 再利用 + updated_at 更新（既存委員会）
       │    └ なし → 新規挿入して committee_id 取得
       └ committee_members テーブル:
            ├ 当該 committee_id の既存レコード全削除(全置換)
            └ 全行 INSERT
6. [VBA] 終了コード(0=成功, 2=年度重複拒否, 1=失敗) で結果表示
```

### 年度の重複チェック（新規作成時のみ）

新規年度 [new_start, new_end] と既存年度 [exist_start, exist_end] が重複する条件:

```
new_start <= exist_end AND new_end >= exist_start
```

1つでも重複する既存年度があれば、新規作成を拒否（終了コード2）。

#### 具体例

| 既存年度 | C1入力 | 新規label | 判定 |
|---|---|---|---|
| 2026-2027 | 2026 | 2026-2027 | 既存再利用（OK） |
| 2026-2027 | 2027 | 2027-2028 | 2027年重複 → 拒否 |
| 2026-2027 | 2028 | 2028-2029 | 重複なし → 新規作成（OK） |

> この仕組みで「2026-2027 既存の状態で間違えて 2027 と入力しても 2027-2028 を作らない」を実現する。

### 同一委員会の再保存時の挙動

**A. 全置換**（採用）: 当該 `committee_id` の `committee_members` を全削除→全挿入。
シンプルだが member_id が毎回変わる。現段階はこの方式で運用する。

> **全置換 vs 差分適用**: 現状は**全置換**で運用。
> 差分適用方式（既存レコードを特定し追加/更新/削除を個別適用）は**検討事項**。

---

## 7. 正規化の検討（メモ）

### 7-1. 経由駅の繰返しグループ

`via_station_1..5` は第1正規形の観点で繰返しグループに該当する。厳密な正規化を行う場合は、
`member_via_stations` のように別テーブルに分離する。ただし EkiSpert が最大5経由を固定想定しており、
現状の列保持のほうが実用的。必要になれば分離テーブルへ移行する。

### 7-2. 人物情報と経路情報の分離

1人につき「人物情報（役職/氏名/連絡先）」と「経路情報（始発/経由/到着）」が1:1で対応する。
別テーブルに分離することも可能だが、今回は1人1経路を想定し `committee_members` に集約する方針。

---

## 8. 構成ファイル

| ファイル | 役割 | 配置 |
|---|---|---|
| `func_CollectCellData.bas` | Excel データ収集（A1/C1取得＋lastRow カット＋氏名空行カット） | `EkiSpert\Alignment\vba\` |
| `func_GetDb.bas` | TSV書出し＋Python呼出し＋結果受取 | `EkiSpert\Alignment\vba\` |
| `main_SynchDb.bas` | DBインターフェイス主導（SynchDb/Step1/Step2） | `EkiSpert\Alignment\vba\` |
| `create_db.py` | TSV読込→SQLite DB作成/更新（年度重複チェック付き） | `EkiSpert\Alignment\python\` |
| `committees.db` | SQLite DB本体 | `sharedRoot\05総務\99 Database\` |

> `python\` フォルダは `main_SynchModules.bas` の同期対象外（.bas のみ同期のため）。

---

## 9. スコープ外（後続フェーズ）

- 差分適用方式の検討・実装（現状は全置換で運用）
- DB読込機能（API連携で経路探索結果をDBに書き戻す等）
- EkiSpert API 連携（現状はDBインターフェイスとして完結）