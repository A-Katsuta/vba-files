Attribute VB_Name = "MainModule"
Option Explicit

Private Const APP_TITLE As String = "Phase1 自動ビルドシステム"

Public Sub Main()
    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Logger.Initialize
    WorkbookInitializer.InitializeThisWorkbook

    Dim mgr As DataManager
    Set mgr = New DataManager
    mgr.Initialize

    mgr.ValidateInput True
    mgr.RefreshAggregations

    Logger.Success "初期化と集計更新が完了しました。"
    MsgBox "ビルド済みワークブックの準備が完了しました。", vbInformation, APP_TITLE

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

EH:
    Logger.Error "Main でエラーが発生: " & Err.Description
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, APP_TITLE
    Resume CleanExit
End Sub

Public Sub RefreshData()
    On Error GoTo EH

    Logger.Info "入力データ検証と集計更新を開始します。"

    Dim mgr As DataManager
    Set mgr = New DataManager
    mgr.Initialize

    mgr.ValidateInput True
    mgr.RefreshAggregations

    Logger.Success "入力データ検証と集計更新が正常に完了しました。"
    Exit Sub

EH:
    Logger.Error "RefreshData でエラーが発生: " & Err.Description
    MsgBox "データ更新中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub ExportLogSnapshot()
    On Error GoTo EH

    Dim targetPath As String
    targetPath = ThisWorkbook.Path & "\build-log-" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    Logger.ExportToFile targetPath
    MsgBox "ログを出力しました: " & targetPath, vbInformation, APP_TITLE
    Exit Sub

EH:
    Logger.Error "ExportLogSnapshot でエラーが発生: " & Err.Description
    MsgBox "ログの書き出しに失敗しました。" & vbCrLf & Err.Description, vbCritical, APP_TITLE
End Sub
