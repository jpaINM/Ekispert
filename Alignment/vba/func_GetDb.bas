'===========================================
' 機能: CollectCellData のデータを Python(create_db.py) に渡して
'          SQLite DB(committees.db) を作成・更新する
' 戻り値: Boolean  True=作成成功 / False=中止または失敗
' 引数  : errorfrag   (ByRef Boolean)  エラー時にTrueが設定される
'          comiteename (ByVal String)   委員会名(A1)
'          startYear   (ByVal String)   年度の開始年(C1, 例: 2026)
'          rawData     (ByRef Variant)  CollectCellData の戻り値(2次元配列)
' 前提  : ・Python 3.x がインストールされていること
'          ・create_db.py が Alignment\python\ に配置されていること
' 仕様  : ・office(4列), home_address(5列) が Excelエラー値なら空白に置換
'          ・中間TSVを%TEMP%に書き出し、Python に引数で渡す
'          ・TSV構成: 1行目=comiteename, 2行目=startYear, 3行目=ヘッダ, 4行目以降=データ
'===========================================
Function GetDb(ByRef errorfrag As Boolean, ByVal comiteename As String, ByVal startYear As String, ByRef rawData As Variant) As Boolean

    Const COL_OFFICE As Long = 4        ' D列: 事業所
    Const COL_HOME_ADDR As Long = 5      ' E列: 自宅住所
    Const PYTHON_EXE As String = "python.exe"
    ' create_db.py はこのプロジェクトの python フォルダに配置(環境変数SHARED_ROOT基準)
    Const SCRIPT_REL As String = "\93 md\EkiSpert\Alignment\python\create_db.py"

    GetDb = False    ' 既定は失敗

    ' 1. データの有無確認
    If IsEmpty(rawData) Then
        errorfrag = True
        Debug.Print "[GetDb] rawData が空のため処理中止"
        Exit Function
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = UBound(rawData, 1)
    lastCol = UBound(rawData, 2)
    Debug.Print "[GetDb] 対象行数(ヘッダ含)=" & lastRow & " / 列数=" & lastCol

    ' 2. 中間TSVを%TEMP%に書き出し
    '    1行目: comiteename
    '    2行目: startYear(年度の開始年)
    '    3行目: ヘッダ(rawData 行1)
    '    4行目以降: データ(rawData 行2以降)
    '      office(4列), home_address(5列) は Excelエラー値なら空白に置換
    Dim tsvPath As String
    tsvPath = Environ("TEMP") & "\ekispert_collect.tsv"
    Debug.Print "[GetDb] 中間TSV: " & tsvPath

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' TSV書込(上書き、TristateTrue=Unicode で日本語対応)
    Set ts = fso.CreateTextFile(tsvPath, True, True)

    Dim r As Long, c As Long, v As Variant
    Dim line As String

    ' 1行目: comiteename
    ts.WriteLine comiteename

    ' 2行目: startYear(年度の開始年)
    ts.WriteLine startYear

    ' 3行目以降: rawData(ヘッダ+データ)
    For r = 1 To lastRow
        line = ""
        For c = 1 To lastCol
            v = rawData(r, c)
            ' office(4列), home_address(5列) の Excelエラー値を空白化
            If (c = COL_OFFICE Or c = COL_HOME_ADDR) Then
                If IsError(v) Then v = ""
            End If
            If c > 1 Then line = line & vbTab
            ' Empty/Null は空文字扱い
            If IsEmpty(v) Or IsNull(v) Then
                line = line & ""
            Else
                line = line & CStr(v)
            End If
        Next c
        ts.WriteLine line
    Next r

    ts.Close
    Set ts = Nothing

    ' 3. Python スクリプトのパス解決
    Dim sharedRoot As String, scriptPath As String
    sharedRoot = Environ("SHARED_ROOT")
    If Len(sharedRoot) = 0 Then
        errorfrag = True
        Debug.Print "[GetDb] 環境変数 SHARED_ROOT が未設定"
        GoTo Cleanup
    End If
    scriptPath = sharedRoot & SCRIPT_REL
    If Not fso.FileExists(scriptPath) Then
        errorfrag = True
        Debug.Print "[GetDb] Pythonスクリプトが見つかりません: " & scriptPath
        GoTo Cleanup
    End If

    ' 4. Python を実行(同期待ち)
    '    WScript.Shell.Run の第3引数 True で完了まで待機
    Dim wsh As Object, cmd As String, exitCode As Long
    Set wsh = CreateObject("WScript.Shell")
    ' パスにスペースが含まれる可能性があるため引用符で囲む
    cmd = PYTHON_EXE & " " & Chr(34) & scriptPath & Chr(34) & " " & Chr(34) & tsvPath & Chr(34)
    Debug.Print "[GetDb] 実行コマンド: " & cmd

    exitCode = wsh.Run(cmd, 0, True)    ' 0=非表示ウィンドウ, True=待機
    Debug.Print "[GetDb] Python終了コード: " & exitCode

    If exitCode = 0 Then
        GetDb = True
        Debug.Print "[GetDb] DB作成/更新 成功"
    Else
        errorfrag = True
        Debug.Print "[GetDb] DB作成/更新 失敗(終了コード=" & exitCode & ")"
    End If

Cleanup:
    ' 中間TSVを削除
    On Error Resume Next
    If fso.FileExists(tsvPath) Then fso.DeleteFile tsvPath, True
    On Error GoTo 0
    Set wsh = Nothing
    Set fso = Nothing

End Function