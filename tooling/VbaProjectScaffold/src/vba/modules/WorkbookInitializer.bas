Attribute VB_Name = "WorkbookInitializer"
Option Explicit

Private Const XL_BUTTON_CONTROL As Long = 7

Public Sub InitializeThisWorkbook()
    Logger.Info "ワークブック初期化を開始します。"

    CreateOrUpdateSheets
    ApplySheetFormatting
    EnsureActionButtons

    Logger.Success "ワークブック初期化が完了しました。"
End Sub

Private Sub CreateOrUpdateSheets()
    Dim mgr As DataManager
    Set mgr = New DataManager
    mgr.Initialize
End Sub

Private Sub ApplySheetFormatting()
    Dim wsInput As Worksheet
    Dim wsSummary As Worksheet

    Set wsInput = Utils.SafeSheet("入力")
    Set wsSummary = Utils.SafeSheet("集計")

    With wsInput.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(219, 229, 241)
    End With
    wsInput.Columns("A:G").AutoFit

    wsSummary.Columns("A:B").AutoFit
End Sub

Private Sub EnsureActionButtons()
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = Utils.SafeSheet("集計")

    Dim shapeName As String
    shapeName = "btnRefreshData"

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo EH

    Dim btn As Shape
    Set btn = ws.Shapes.AddFormControl(XL_BUTTON_CONTROL, 20, 60, 140, 30)
    btn.Name = shapeName
    btn.TextFrame.Characters.Text = "集計を更新"
    btn.OnAction = "RefreshData"

    Logger.Info "集計シートに更新ボタンを配置しました。"
    Exit Sub
EH:
    Logger.Warn "ボタンの配置に失敗しました: " & Err.Description
End Sub
