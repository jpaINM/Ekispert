'===========================================
' 機能: DBインターフェイスの主導処理
'          このエクセルマクロは Excel と SQLite DB のインターフェイス。
'          EkiSpert は実質的にDBとして扱い、本マクロで委員会情報をDBに保存する。
'          各ステップを独立Subとし、SynchDb が連続実行する。
'          Step1: CollectCellData で収集＋debug出力
'          Step2: DBへ同期(2段階確認付き)
' 備考  : 同期方式は現状「全置換」(当該委員会の committee_members を全削除→全挿入)。
'          差分適用方式は検討事項(後続フェーズで検討)。
'          年度は C1 から取得(開始年)。既存年度と期間重複する新規年度は作成しない。
'===========================================

' --- main: 連続実行 ---
Sub SynchDb()
    Step1_CollectAndShow
    Step2_SynchDb
End Sub

' --- Step1: 収集＋debug出力 ---
Sub Step1_CollectAndShow()
    Dim errorfrag As Boolean: errorfrag = False
    Dim comiteename As String: comiteename = ""
    Dim startYear As String: startYear = ""
    Dim rawData As Variant

    rawData = CollectCellData(errorfrag, comiteename, startYear)

    If errorfrag Then
        Debug.Print "[Step1] CollectCellData でエラー発生のため終了"
        Exit Sub
    End If

    Debug.Print "===== Step1: CollectCellData 結果 ====="
    Debug.Print "委員会名(A1): " & comiteename
    Debug.Print "年度開始年(C1): " & startYear

    If IsEmpty(rawData) Then
        Debug.Print "データなし"
        Debug.Print "================================="
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long, r As Long, c As Long, line As String
    lastRow = UBound(rawData, 1)
    lastCol = UBound(rawData, 2)
    Debug.Print "配列サイズ: " & lastRow & "行 x " & lastCol & "列"

    For r = 1 To lastRow
        line = "行" & r & " (シート" & (r + 2) & "行): "
        For c = 1 To lastCol
            line = line & "[" & CStr(rawData(r, c)) & "]"
        Next c
        Debug.Print line
    Next r
    Debug.Print "================================="

End Sub

' --- Step2: DBへ同期(2段階確認付き) ---
' 新規委員会: committees にINSERT、committee_members に全INSERT
' 既存委員会: committee_id を再利用、committee_members を全置換(全削除→全挿入)
Sub Step2_SynchDb()
    Dim errorfrag As Boolean: errorfrag = False
    Dim comiteename As String: comiteename = ""
    Dim startYear As String: startYear = ""
    Dim rawData As Variant

    ' データ再取得(GetDb は rawData を引数で受け取るため)
    rawData = CollectCellData(errorfrag, comiteename, startYear)
    If errorfrag Then
        Debug.Print "[Step2] CollectCellData でエラー発生のため終了"
        Exit Sub
    End If
    If IsEmpty(rawData) Then
        Debug.Print "[Step2] データなしのため終了"
        Exit Sub
    End If

    ' 2段階確認
    If MsgBox("DBへ同期しますか？", vbYesNo + vbQuestion, "DB同期確認(1/2)") = vbNo Then
        Debug.Print "[Step2] 1回目の確認でキャンセルされました"
        Exit Sub
    End If

    If MsgBox("本当に同期しますか？", vbYesNo + vbExclamation, "DB同期確認(2/2)") = vbNo Then
        Debug.Print "[Step2] 2回目の確認でキャンセルされました"
        Exit Sub
    End If

    ' DB同期実行
    Dim ok As Boolean
    ok = GetDb(errorfrag, comiteename, startYear, rawData)

    If ok Then
        MsgBox "DBの同期が完了しました。", vbInformation, "完了"
    Else
        MsgBox "DBの同期に失敗しました。" & vbCrLf & _
               "イミディエイトウィンドウで詳細を確認してください。", vbCritical, "エラー"
    End If

End Sub