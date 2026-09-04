# EkiSpert Alignment - Excelデータ収集 設計書

## 1. 目的

Excel ワークブック内の「DB」シートに配置された表データを走査し、
lastRow で末尾の空行をカットした上で **2次元配列** として取得する基礎機能を提供する。
併せて A1セル（委員会名）と C1セル（年度の開始年）を取得して返す。
さらに氏名列（列2）が空/null/エラー値の不要行をカットする。

本設計書は Phase 1b（Excelデータ収集）に対応する。
DB保存時のテーブル設計は `DB設計.md` に分離して記載する。
DB同期処理（`GetDb` + `create_db.py`）は Phase 2 で実装済み。

---

## 2. 対象シートレイアウト

- シート名: **`DB`**
- A1: **委員会名**（`comiteename` として取得）
- C1: **年度の開始年**（`startYear` として取得。数値、例: 2026）
- ヘッダ行: **3行目**
- データ行: **4行目以降**（実データ末尾まで、lastRow でカット）
- 取得範囲: **`A3:R{lastRow}`**（lastRow は動的）
- 列構成: AからR（18列）

```
行(シート)   役割
──────────────────────────
1            委員会名（A1） / 年度開始年（C1）
3            ヘッダ（列名）
4 〜 lastRow  データレコード（1行＝1レコード）
lastRow+1〜   空行（取得対象外）
```

> 取得範囲の下端は lastRow で動的に決定し、24行目等で固定しない。
> さらに、データ行のうち**氏名列（列2=B列）が空/null/エラー値の行は不要行としてカット**する。

---

## 3. 戻り値のデータ構造

```
Variant (2次元配列、1ベース)
  ├ 行1   = シート3行目(ヘッダ)
  └ 行2〜 = 不要行カット済のデータ(末尾空行・氏名空行除去済)
```

- 列数は固定18（AからR）
- 行数は「ヘッダ1行＋有効データ行数」
- 下流からは `rawData(行, 列)` でインデックスアクセスする

### 例

```vba
Dim errorfrag As Boolean, comiteename As String, startYear As String
Dim rawData As Variant
rawData = CollectCellData(errorfrag, comiteename, startYear)

Debug.Print comiteename          ' 委員会名(A1)
Debug.Print startYear            ' 年度開始年(C1)
Debug.Print rawData(1, 1)    ' ヘッダ1列目
```

---

## 4. 関数仕様

### `CollectCellData`（`func_CollectCellData.bas`）

```vba
Function CollectCellData(ByRef errorfrag As Boolean, ByRef comiteename As String, ByRef startYear As String) As Variant
```

| 項目 | 内容 |
|---|---|
| 引数1 | `errorfrag As Boolean`（ByRef）。エラー時に `True` が設定される |
| 引数2 | `comiteename As String`（ByRef）。A1の委員会名が設定される（正常時） |
| 引数3 | `startYear As String`（ByRef）。C1の年度開始年が設定される（正常時） |
| 戻り値 | `Variant`（2次元配列）。エラー時・データ0件時は `Empty` |
| 副作用 | なし（読込のみ。シートは変更しない） |
| ログ | `Debug.Print` で comiteename・startYear・lastRow・取得範囲・フィルタ件数を出力 |

### エラー判定・挙動

| 状況 | 挙動 |
|---|---|
| シート「DB」が存在しない | `errorfrag=True`、`Empty` を返して Exit |
| A1（委員会名）が空 | `errorfrag=True`、`Empty` を返して Exit |
| C1（年度開始年）が空 | `errorfrag=True`、`Empty` を返して Exit |
| C1（年度開始年）が数値でない | `errorfrag=True`、`Empty` を返して Exit |
| lastRow < 4（データ0件） | `Empty` を返して Exit |
| フィルタ後に有効データ0件 | `Empty` を返して Exit |
| 正常時 | ヘッダ1行＋有効データ行の2次元配列を返す |

### 不要行フィルタの判定条件

データ行の**氏名列（列2=B列）**の値が以下のいずれかなら、その行は不要としてカット:

- `IsError(v)` — Excelエラー値（`#N/A` 等）
- `IsNull(v)` — Null
- `IsEmpty(v)` — Empty
- `Len(Trim$(CStr(v))) = 0` — 空文字・空白のみ

> ヘッダ行（配列行1）は判定対象外・必ず残す。
> office/home_address のエラー値はフィルタ対象外（下流 GetDb で空白置換で処理）。

### 前提条件

- ワークブックに「DB」シートが存在すること
- A1に委員会名、C1に年度の開始年（数値）が入力されていること
- ヘッダは3行目、データは4行目以降

---

## 5. lastRow カットの確認

本実装では lastRow を2通りで取得し、Debug.Print で比較確認する:

- **lastRowByA**: A列基準（`ws.Cells(Rows.Count,"A").End(xlUp).Row`）
- **lastRowMax**: AからR 全列の最大（A列が空で後続列にデータがあるケースの検出用）

標準は **A列基準** を採用。両者が一致すれば迷いなし。
乖離があれば Debug.Print 出力を見て基準を再検討する。

---

## 6. 処理フロー

```
1. シート「DB」を取得（ThisWorkbook.Sheets("DB")）
       └ 存在しない → errorfrag=True / Empty 返却 / Exit
2. A1（委員会名）を取得 → comiteename
       └ 空なら errorfrag=True / Empty 返却 / Exit
2-2. C1（年度の開始年）を取得 → startYear
       └ 空 or 数値でない → errorfrag=True / Empty 返却 / Exit
3. lastRow を取得
       ├ A列基準: ws.Cells(Rows.Count,"A").End(xlUp).Row
       └ 全列最大: AからR 各列の End(xlUp).Row の最大値
4. 採用 lastRow < 4（データ0件）なら Empty 返却 / Exit
5. 範囲 A3:R{lastRow} を2次元配列で取得
6. 不要行フィルタ: 氏名列(列2)が空/null/エラー値のデータ行をカット
       ├ 残す行数をカウント（ヘッダ1行は固定で残す）
       └ 全行カット(有効0件)なら Empty 返却 / Exit
7. 新しい2次元配列 filteredData を組み立て（ヘッダ+有効データ行）
8. filteredData を戻り値として返す
```

---

## 7. 検証方法

`func_CollectCellData.bas` の `TestCollectCellData` を実行し、
イミディエイトウィンドウで以下を確認:

- A1（委員会名）の取得結果（comiteename）
- C1（年度開始年）の取得結果（startYear）
- lastRow（A列基準 vs 全列最大）の値と妥当性
- 取得範囲（A3:R{lastRow}）と行数
- フィルタ結果（残件数 / カット件数）
- 氏名空行がカットされていること

---

## 8. DBインターフェイス主導処理 `SynchDb`（Phase 2 実装済み）

本エクセルマクロは **Excel と SQLite DB のインターフェイス** として位置づける。
各ステップを独立Subとし、`SynchDb` が連続実行する。

```vba
Sub SynchDb()              ' main: 連続実行
    Step1_CollectAndShow    ' 収集＋debug出力
    Step2_SynchDb          ' DBへ同期(2段階確認)
End Sub
```

| Sub | 役割 |
|---|---|
| `Step1_CollectAndShow` | `CollectCellData` を呼出し、取得結果をdebug出力 |
| `Step2_SynchDb` | 2段階 MsgBox 確認 → `GetDb` でDB同期 |

> モジュールファイル: `main_SynchDb.bas`

---

## 9. DB同期 `GetDb`（Phase 2 実装済み）

```vba
Function GetDb(ByRef errorfrag As Boolean, ByVal comiteename As String, ByVal startYear As String, ByRef rawData As Variant) As Boolean
```

| 項目 | 内容 |
|---|---|
| 引数1 | `errorfrag As Boolean`（ByRef）。エラー時に `True` |
| 引数2 | `comiteename As String`（ByVal）。委員会名 |
| 引数3 | `startYear As String`（ByVal）。年度の開始年 |
| 引数4 | `rawData As Variant`（ByRef）。CollectCellData の戻り値配列 |
| 戻り値 | `Boolean`。True=同期成功 / False=中止または失敗 |

### 処理フロー

```
1. rawData を走査し、office(4列)/home_address(5列) が Excelエラー値なら空白に置換
2. rawData + comiteename + startYear を中間TSV(%TEMP%) に書き出し
       ├ 1行目: comiteename
       ├ 2行目: startYear
       ├ 3行目: ヘッダ(18列)
       └ 4行目以降: データ
3. python.exe create_db.py <tsv_path> を同期待ち実行(WScript.Shell.Run)
4. 終了コード(0=成功, 2=年度重複拒否, 1=失敗) で結果受取 → Debug.Print + 戻り値
5. 中間TSVを削除
```

> DB同期の詳細（テーブル定義・年度重複チェック・保存フロー）は `DB設計.md` 参照。
> Python スクリプト `create_db.py` は `EkiSpert\Alignment\python\` に配置。

### 同期方式（新規/既存）

- **新規年度**: `fiscal_years` に INSERT（重複チェックで既存と重複なら拒否）
- **既存年度**: `fiscal_year_id` 再利用
- **新規委員会**: `committees` に INSERT
- **既存委員会**: `committee_id` 再利用、`committee_members` を**全置換**

### 前提条件

- Python 3.x がインストールされ、PATH が通っていること
- 環境変数 `SHARED_ROOT` が設定されていること
- `create_db.py` が `Alignment\python\` に配置されていること

---

## 10. 後続フェーズ

| 機能 | 状態 | 概要 |
|---|---|---|
| `func_PostDb` | 不要（GetDb がDB同期を兼ねる） | 当初予定のPostDb機能は GetDb に統合 |
| EkiSpert API 連携 | 未実装 | 将来の拡張（現状はDBインターフェイスとして完結） |
| 差分適用方式 | 検討事項 | 現状は全置換。必要に応じて差分適用へ移行 |