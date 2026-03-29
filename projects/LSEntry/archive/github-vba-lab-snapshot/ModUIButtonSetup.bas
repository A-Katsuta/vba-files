Option Explicit

Private Const BTN_W As Double = 110
Private Const BTN_H As Double = 28

' 配置と基本設定
Public Sub InstallActionButtons()
    On Error GoTo EH
    Dim ws As Worksheet

    ' データ取得
    Set ws = GetSheet(Sheet_DataAcquire)
    AddOrUpdateButton ws, "btnGetOutlook", "予定取得", "Run_予定取得", "B2"

    ' データ登録
    Set ws = GetSheet(Sheet_DataEntry)
    AddOrUpdateButton ws, "btnExecuteTransfer", "転記実行", "Run_転記実行", "B2"
    AddOrUpdateButton ws, "btnClearInput", "入力クリア", "Run_入力クリア", "D2"
    ApplyDataEntryValidation ws

    ' 月次データ
    Set ws = GetSheet(Sheet_Monthly)
    AddOrUpdateButton ws, "btnRefreshMonthly", "月次更新", "Run_月次更新", "B2"
    AddOrUpdateButton ws, "btnCopyClipboard", "コピー", "Run_コピー", "D2"

    MsgBox "ボタンの配置と基本設定を完了しました。", vbInformation
    Exit Sub
EH:
    MsgBox "InstallActionButtons: " & Err.Description, vbExclamation
End Sub

Private Sub AddOrUpdateButton(ws As Worksheet, btnName As String, caption As String, macroName As String, anchorCell As String)
    On Error Resume Next
    Dim btn As Button
    Set btn = ws.Buttons(btnName)
    On Error GoTo 0

    Dim leftPos As Double, topPos As Double
    leftPos = ws.Range(anchorCell).Left
    topPos = ws.Range(anchorCell).Top

    If btn Is Nothing Then
        Set btn = ws.Buttons.Add(leftPos, topPos, BTN_W, BTN_H)
        btn.Name = btnName
    Else
        btn.Left = leftPos
        btn.Top = topPos
        btn.Width = BTN_W
        btn.Height = BTN_H
    End If

    btn.OnAction = macroName
    btn.Characters.Text = caption
End Sub

' ラッパー（ステータスバー付き）
Public Sub Run_予定取得()
    Dim sb As Variant: sb = Application.StatusBar
    On Error GoTo EH
    Application.StatusBar = "Outlook予定を取得中..."
    ExecuteOutlookSchedule
    Application.StatusBar = False: Exit Sub
EH:
    Application.StatusBar = False
    MsgBox "予定取得でエラー: " & Err.Description, vbExclamation
End Sub

Public Sub Run_転記実行()
    Dim sb As Variant: sb = Application.StatusBar
    On Error GoTo EH
    Application.StatusBar = "月次へ転記中..."
    TransferDataToMonthlySheet
    Application.StatusBar = False: Exit Sub
EH:
    Application.StatusBar = False
    MsgBox "転記でエラー: " & Err.Description, vbExclamation
End Sub

Public Sub Run_入力クリア()
    On Error GoTo EH
    ClearInputData
    Exit Sub
EH:
    MsgBox "入力クリアでエラー: " & Err.Description, vbExclamation
End Sub

Public Sub Run_月次更新(Optional ByVal showConfirm As Boolean = True)
    On Error GoTo EH
    ClearMonthlyDataAndRefreshCalendar showConfirm
    Exit Sub
EH:
    MsgBox "月次更新でエラー: " & Err.Description, vbExclamation
End Sub

Public Sub Run_コピー()
    On Error GoTo EH
    CopyLatestDataToClipboard
    Exit Sub
EH:
    MsgBox "コピーでエラー: " & Err.Description, vbExclamation
End Sub

Private Sub ApplyDataEntryValidation(ws As Worksheet)
    On Error GoTo EH
    Const START_ROW As Long = 8

    ' 時間列: HHMM または Excel 時刻
    Dim colTime As Long: colTime = DataCol_Time
    Dim rngTime As Range
    Set rngTime = ws.Range(ws.Cells(START_ROW, colTime), ws.Cells(ws.Rows.Count, colTime))
    With rngTime.Validation
        .Delete
    End With

    Dim f As String
    f = "=OR(AND(LEN(" & ws.Cells(START_ROW, colTime).Address(False, False) & ")>=3,LEN(" & ws.Cells(START_ROW, colTime).Address(False, False) & ")<=4,ISNUMBER(--" & ws.Cells(START_ROW, colTime).Address(False, False) & ")),ISNUMBER(" & ws.Cells(START_ROW, colTime).Address(False, False) & "))"
    With rngTime.Validation
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=f
        .IgnoreBlank = True
        .InputTitle = "時間入力"
        .InputMessage = "HHMM または Excel 時刻([hh]:mm)"
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "HHMM（例: 930/0930）または [hh]:mm を入力してください。"
        .ShowError = True
        .ShowInput = True
    End With

    ' 作番列: マスタの作番をプルダウン提供
    Dim colWorkNo As Long: colWorkNo = DataCol_WorkNo
    Dim rngWorkNo As Range
    Set rngWorkNo = ws.Range(ws.Cells(START_ROW, colWorkNo), ws.Cells(ws.Rows.Count, colWorkNo))
    With rngWorkNo.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=Master_WorkNumber"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' 作業コード列: マスタの作業コードをプルダウン提供
    Dim colWorkCode As Long: colWorkCode = DataCol_Category
    Dim rngWorkCode As Range
    Set rngWorkCode = ws.Range(ws.Cells(START_ROW, colWorkCode), ws.Cells(ws.Rows.Count, colWorkCode))
    With rngWorkCode.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=Master_WorkCode"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    Exit Sub
EH:
    ' 失敗時は無視（他環境でも動くように）
End Sub
