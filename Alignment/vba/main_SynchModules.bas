Sub SyncModulesFromFolder_FullSync()

    Dim TARGET_FOLDER_PATH As String
    Dim sharedRoot As String

    sharedRoot = getCurrentFolderPath()
    
    TARGET_FOLDER_PATH = sharedRoot & "\93 md\EkiSpert\Alignment\vba"
   
    Const THIS_PROC_NAME As String = "main_SynchModules"
    
    Dim fso As Object       ' FileSystemObject
    Dim vbProj As Object    ' VBA Project
    Dim targetFolder As Object
    Dim targetFile As Object
    Dim moduleName As String
    Dim existingModule As Object
    Dim syncCount As Long
    Const MAX_MODULE_NAME_LEN As Long = 30
    
    ' FSOの初期化とフォルダの存在チェック
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(TARGET_FOLDER_PATH) Then
        MsgBox "指定されたフォルダが見つかりません: " & TARGET_FOLDER_PATH, vbCritical
        Exit Sub
    End If
    
    ' 現在のブックのVBAプロジェクトを取得
    Set vbProj = ThisWorkbook.VBProject
    Set targetFolder = fso.GetFolder(TARGET_FOLDER_PATH)

    Application.ScreenUpdating = False ' 画面更新を停止
    syncCount = 0

    ' --- 事前チェック: ファイル名(拡張子なし)＝モジュール名が長すぎると同期に失敗する ---
    ' VBA標準モジュール名は31文字制限があるため、運用上30文字以内に統一する
    If Not ValidateModuleNameLengths(targetFolder, fso, THIS_PROC_NAME, MAX_MODULE_NAME_LEN) Then
        Application.ScreenUpdating = True
        Set fso = Nothing
        Exit Sub
    End If

    ' 実行中のプロシージャ名を取得

    ' フォルダ内のファイルをループ処理
    For Each targetFile In targetFolder.Files
        
        ' .bas ファイルのみを対象
        If LCase(fso.GetExtensionName(targetFile.Name)) = "bas" Then

            ' --- ガード: 空ファイルはスキップ(退避中モジュール等) ---
            If targetFile.Size = 0 Then
                Debug.Print targetFile.Name & " は空ファイルのためスキップ"
                GoTo NextFile
            End If
        
            ' ファイル名（拡張子なし）をモジュール名として取得
            moduleName = fso.GetBaseName(targetFile.Name)
            
            ' ----------------------------------------------------
            ' ★★★ 追加されたスキップ処理 ★★★
            ' 実行中のモジュール名（ファイル名）と同じ場合はスキップ
            If StrComp(moduleName, THIS_PROC_NAME, vbTextCompare) = 0 Then
                Debug.Print "【スキップ】" & moduleName & " は実行中のモジュールのため処理しません。"
                GoTo NextFile ' 次のファイルへ
            End If
            ' ----------------------------------------------------
            
            ' ----------------------------------------------------
            ' 1. 既存モジュールの確認と削除（＝更新準備）
            ' ----------------------------------------------------
            Set existingModule = Nothing
            On Error Resume Next
            ' 名前が一致する既存モジュールを探す
            Set existingModule = vbProj.VBComponents(moduleName)
            On Error GoTo 0
            
            If Not existingModule Is Nothing Then
                ' 既にモジュールが存在する場合（更新対象）は、一旦削除
                vbProj.VBComponents.Remove existingModule
                Debug.Print moduleName & " を削除しました (更新準備)"
            End If
            
            ' ----------------------------------------------------
            ' 2. モジュールのインポートと名前の強制変更
            ' ----------------------------------------------------

            ' ImportはVBComponentを返す（参照設定が無くてもObjectで受けられる）
            Set existingModule = vbProj.VBComponents.Import(targetFile.Path)
            If existingModule Is Nothing Then
                Debug.Print targetFile.Name & " のインポートに失敗しました。"
                GoTo NextFile
            End If

            ' モジュール名 (Name プロパティ) をファイル名に強制設定
            On Error Resume Next
            existingModule.Name = moduleName
            If Err.Number <> 0 Then
                Debug.Print "【警告】" & existingModule.Name & " の名前を " & moduleName & " に変更できませんでした。既に別のコンポーネントで使用されている可能性があります。"
                Err.Clear
            Else
                Debug.Print moduleName & " をインポートし、名前を確定しました。"
            End If
            On Error GoTo 0

            syncCount = syncCount + 1
            
NextFile:
        End If ' .basファイルのチェック終了
        
    Next targetFile ' ファイルのループ終了
    Application.ScreenUpdating = True
    Set fso = Nothing
    MsgBox "モジュールの同期（新規作成・既存更新）が完了しました。処理数: " & syncCount, vbInformation
    
End Sub

Private Function ValidateModuleNameLengths(ByVal targetFolder As Object, ByVal fso As Object, ByVal thisProcName As String, ByVal maxLen As Long) As Boolean
    Dim targetFile As Object
    Dim moduleName As String
    Dim badList As String
    Dim countBad As Long

    badList = ""
    countBad = 0

    For Each targetFile In targetFolder.Files
        If LCase(fso.GetExtensionName(targetFile.Name)) = "bas" Then
            moduleName = fso.GetBaseName(targetFile.Name)
            If StrComp(moduleName, thisProcName, vbTextCompare) <> 0 Then
                If Len(moduleName) > maxLen Then
                    countBad = countBad + 1
                    badList = badList & vbCrLf & "- " & moduleName & "（" & Len(moduleName) & "文字）"
                End If
            End If
        End If
    Next targetFile

    If countBad > 0 Then
        MsgBox "同期できないファイル名があります（.basの拡張子を除いた名前＝モジュール名）。" & vbCrLf & _
               "以下は " & maxLen & "文字以内にしてください。" & vbCrLf & _
               badList, vbExclamation
        ValidateModuleNameLengths = False
        Exit Function
    End If

    ValidateModuleNameLengths = True
End Function

Private Function GetCurrentFolderPath() As String
    ' Outlookにはブックの概念がないため、固定パスまたは環境変数を使う
    GetCurrentFolderPath = Environ("SHARED_ROOT")
End Function
