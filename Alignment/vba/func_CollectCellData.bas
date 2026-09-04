'===========================================
' 機能: エクセルデータ情報を収集する
'          DBシートの実データ範囲を lastRow でカットして2次元配列で取得
'          併せて A1(委員会名) と C1(年度の開始年) を取得して返す
'          氏名列(列2)が空/null/エラー値の行は不要行としてカットする
'          このステップでは取得のみ。データ削除は行わない。
' 戻り値: Variant (2次元配列)
'          行1 = シート3行目(ヘッダ) / 行2以降 = データ(末尾空行・氏名空行カット済)
'          エラー時は errorfrag=True、戻り値は Empty
' 引数  : errorfrag   (ByRef Boolean)  エラー時にTrueが設定される
'          comiteename (ByRef String)    A1の委員会名が設定される
'          startYear   (ByRef String)    C1の年度開始年が設定される(例: 2026)
' 前提  : ・ワークブックに「DB」シートが存在すること
'          ・A1に委員会名、C1に年度の開始年(数値)が入力されていること
'          ・データ先頭行は3行目(3行目=ヘッダ, 4行目以降=データ)
' 備考  : DB形式化(1行=1レコードのDictionary化)は後続フェーズで追加予定
'===========================================
Function CollectCellData(ByRef errorfrag As Boolean, ByRef comiteename As String, ByRef startYear As String) As Variant

    Const SHEET_NAME As String = "DB"
    Const HEADER_ROW As Long = 3        ' シート上のヘッダ行
    Const FIRST_DATA_ROW As Long = 4    ' シート上のデータ先頭行
    Const LAST_COL As String = "R"      ' 取得範囲の右端列
    Const COL_NAME As Long = 2          ' 氏名列(B列)。この列が空/null/エラーの行は不要

    ' 1. シート「DB」の取得（存在確認）
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        errorfrag = True
        Debug.Print "[CollectCellData] シート[" & SHEET_NAME & "]が見つかりません"
        Exit Function
    End If

    ' 2. A1(委員会名) を取得
    comiteename = Trim$(CStr(ws.Range("A1").Value))
    Debug.Print "[CollectCellData] comiteename(A1) = " & comiteename
    If Len(comiteename) = 0 Then
        errorfrag = True
        Debug.Print "[CollectCellData] A1(委員会名)が空のため処理中止"
        Exit Function
    End If

    ' 2-2. C1(年度の開始年) を取得
    startYear = Trim$(CStr(ws.Range("C1").Value))
    Debug.Print "[CollectCellData] startYear(C1) = " & startYear
    If Len(startYear) = 0 Then
        errorfrag = True
        Debug.Print "[CollectCellData] C1(年度の開始年)が空のため処理中止"
        Exit Function
    End If
    ' 数値かどうか簡易チェック(数値でなければエラー)
    If Not IsNumeric(startYear) Then
        errorfrag = True
        Debug.Print "[CollectCellData] C1(年度の開始年)が数値でないため処理中止: " & startYear
        Exit Function
    End If

    ' 3. lastRow を取得してDebug.Printで確認
    '    基本は A列基準(既存コードと同じ End(xlUp) 方式)
    Dim lastRowByA As Long
    lastRowByA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    '    念のため A:R 全列の最大 lastRow も併記
    '    (A列が空で後続列にデータがあるケースの検出用)
    Dim col As Long, lastRowMax As Long, tmpLast As Long
    lastRowMax = 0
    For col = 1 To ws.Range(LAST_COL & "1").Column   ' 1から18 (AからR)
        tmpLast = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If tmpLast > lastRowMax Then lastRowMax = tmpLast
    Next col

    ' 4. 採用 lastRow の決定
    '    A列基準と全列最大が一致すれば迷いなし。
    '    乖離があれば Debug.Print で両方を出して結果を見て調整する。
    Dim lastRow As Long
    lastRow = lastRowByA    ' 標準は A列基準

    Debug.Print "[CollectCellData] lastRow(A列基準) = " & lastRowByA
    Debug.Print "[CollectCellData] lastRow(A:R最大)  = " & lastRowMax
    Debug.Print "[CollectCellData] 採用 lastRow     = " & lastRow

    ' 5. lastRow がデータ先頭行未満(=データ0件)なら空で返す
    If lastRow < FIRST_DATA_ROW Then
        Debug.Print "[CollectCellData] データ行なし(取得範囲空)"
        Exit Function
    End If

    ' 6. 範囲 A3:R{lastRow} を2次元配列で取得
    Dim rangeAddr As String
    rangeAddr = "A" & HEADER_ROW & ":" & LAST_COL & lastRow
    Debug.Print "[CollectCellData] 取得範囲 = " & rangeAddr & _
                " / 行数 = " & (lastRow - HEADER_ROW + 1)

    Dim rawData As Variant
    rawData = ws.Range(rangeAddr).Value

    ' 7. 不要行フィルタ: 氏名列(列2)が空/null/エラー値のデータ行をカット
    '    ヘッダ行(行1)は必ず残す。データ行(行2以降)のみ判定。
    Dim rawLastRow As Long, rawLastCol As Long
    rawLastRow = UBound(rawData, 1)
    rawLastCol = UBound(rawData, 2)

    Dim r As Long, c As Long, v As Variant
    Dim keepCount As Long, cutCount As Long
    keepCount = 1    ' ヘッダ行1件は固定で残す
    cutCount = 0

    ' 7-1. まず残す行数をカウント(氏名列で判定)
    For r = 2 To rawLastRow
        v = rawData(r, COL_NAME)
        If IsRowNeeded(v) Then
            keepCount = keepCount + 1
        Else
            cutCount = cutCount + 1
        End If
    Next r

    Debug.Print "[CollectCellData] フィルタ: 残=" & (keepCount - 1) & _
                " / カット=" & cutCount

    ' 7-2. 全行カットされた(=有効データ0件)なら空で返す
    If keepCount = 1 Then
        Debug.Print "[CollectCellData] フィルタ後に有効データなし"
        Exit Function
    End If

    ' 7-3. 新しい2次元配列を組み立て(ヘッダ+有効データ行)
    Dim filteredData() As Variant
    ReDim filteredData(1 To keepCount, 1 To rawLastCol)

    Dim destRow As Long
    destRow = 1
    ' ヘッダ行をコピー
    For c = 1 To rawLastCol
        filteredData(1, c) = rawData(1, c)
    Next c

    ' データ行をコピー(有効行のみ)
    For r = 2 To rawLastRow
        v = rawData(r, COL_NAME)
        If IsRowNeeded(v) Then
            destRow = destRow + 1
            For c = 1 To rawLastCol
                filteredData(destRow, c) = rawData(r, c)
            Next c
        End If
    Next r

    CollectCellData = filteredData

End Function

'-------------------------------------------
' 機能: 氏名列の値から「その行が必要か」を判定
' 戻り値: True=必要 / False=不要(空/null/エラー値)
'-------------------------------------------
Private Function IsRowNeeded(ByVal v As Variant) As Boolean
    ' Excelエラー値(#N/A等)は不要
    If IsError(v) Then
        IsRowNeeded = False
        Exit Function
    End If
    ' Null は不要(CStr(Null)はエラーになるため先にチェック)
    If IsNull(v) Then
        IsRowNeeded = False
        Exit Function
    End If
    ' Empty は不要
    If IsEmpty(v) Then
        IsRowNeeded = False
        Exit Function
    End If
    ' 空文字・空白のみは不要
    If Len(Trim$(CStr(v))) = 0 Then
        IsRowNeeded = False
        Exit Function
    End If
    ' それ以外は必要
    IsRowNeeded = True
End Function
'===========================================
' 検証用: CollectCellData の取得結果をイミディエイトへダンプ
'===========================================
Sub TestCollectCellData()

    Dim errorfrag As Boolean: errorfrag = False
    Dim comiteename As String: comiteename = ""
    Dim startYear As String: startYear = ""
    Dim rawData As Variant
    rawData = CollectCellData(errorfrag, comiteename, startYear)

    If errorfrag Then
        Debug.Print "[Test] エラー発生のため終了"
        Exit Sub
    End If

    Debug.Print "===== CollectCellData ダンプ ====="
    Debug.Print "委員会名(A1): " & comiteename
    Debug.Print "年度開始年(C1): " & startYear

    If IsEmpty(rawData) Then
        Debug.Print "データなし"
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = UBound(rawData, 1)
    lastCol = UBound(rawData, 2)
    Debug.Print "配列サイズ: " & lastRow & "行 x " & lastCol & "列"

    Dim r As Long, c As Long, line As String
    For r = 1 To lastRow
        line = "行" & r & " (シート" & (r + 2) & "行): "
        For c = 1 To lastCol
            line = line & "[" & CStr(rawData(r, c)) & "]"
        Next c
        Debug.Print line
    Next r
    Debug.Print "===== ダンプ終了 ====="

End Sub