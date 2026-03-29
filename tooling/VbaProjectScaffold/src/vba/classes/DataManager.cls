VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataManager"
Option Explicit

Private Const INPUT_SHEET As String = "入力"
Private Const SUMMARY_SHEET As String = "集計"
Private Const MASTER_SHEET As String = "マスタ"

Private mInitialized As Boolean

Public Sub Initialize()
    EnsureEssentialSheets
    EnsureNamedRanges
    mInitialized = True
End Sub

Public Function ValidateInput(Optional ByVal showMessage As Boolean = False) As Boolean
    Dim ws As Worksheet
    Set ws = EnsureSheet(INPUT_SHEET)

    Dim lastRow As Long
    lastRow = WorksheetLastRow(ws, 1)

    Dim issues As Collection
    Set issues = New Collection

    Dim rowIndex As Long
    Dim qty As Double
    Dim price As Double

    If lastRow < 2 Then
        Logger.Warn "入力シートにデータがありません。"
        ValidateInput = False
        Exit Function
    End If

    For rowIndex = 2 To lastRow
        If LenB(Trim$(CStr(ws.Cells(rowIndex, 1).Value))) = 0 Then
            issues.Add "行" & rowIndex & " の ID が未入力です。"
        End If

        qty = Utils.TryParseNumber(ws.Cells(rowIndex, 4).Value, 0)
        price = Utils.TryParseNumber(ws.Cells(rowIndex, 5).Value, 0)
        ws.Cells(rowIndex, 6).Value = qty * price

        If Not IsDate(ws.Cells(rowIndex, 7).Value) Then
            ws.Cells(rowIndex, 7).Value = Date
        End If
    Next rowIndex

    If issues.Count = 0 Then
        ValidateInput = True
        Logger.Info "入力データの検証が完了しました。"
    Else
        ValidateInput = False
        Dim message As String
        message = JoinIssues(issues)
        Logger.Warn "入力データに問題があります: " & message
        If showMessage Then
            MsgBox "入力データに問題があります:" & vbCrLf & message, vbExclamation, "入力検証"
        End If
    End If
End Function

Public Sub RefreshAggregations()
    Dim wsInput As Worksheet
    Dim wsSummary As Worksheet
    Set wsInput = EnsureSheet(INPUT_SHEET)
    Set wsSummary = EnsureSheet(SUMMARY_SHEET)

    Dim lastRow As Long
    lastRow = WorksheetLastRow(wsInput, 1)

    Dim totalAmount As Double
    Dim totalRecords As Long

    If lastRow >= 2 Then
        totalAmount = Application.WorksheetFunction.Sum(wsInput.Range("F2:F" & lastRow))
        totalRecords = lastRow - 1
    Else
        totalAmount = 0
        totalRecords = 0
    End If

    With wsSummary
        .Cells(1, 1).Value = "集計結果"
        .Cells(2, 1).Value = "更新日時"
        .Cells(2, 2).Value = Now
        .Cells(3, 1).Value = "レコード数"
        .Cells(3, 2).Value = totalRecords
        .Cells(4, 1).Value = "合計金額"
        .Cells(4, 2).Value = totalAmount
        .Cells(4, 2).NumberFormatLocal = "\¥#,##0"
        .Range("A1:B4").Columns.AutoFit
    End With

    Logger.Info "集計情報を更新しました。"
End Sub

Public Sub EnsureNamedRanges()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    On Error Resume Next
    wb.Names.Item("rngInputTable").Delete
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = EnsureSheet(INPUT_SHEET)

    Dim lastRow As Long
    lastRow = WorksheetLastRow(ws, 1)
    If lastRow < 2 Then lastRow = 2

    Dim address As String
    address = ws.Range("A1:G" & lastRow).Address(True, True, xlA1, True)
    wb.Names.Add Name:="rngInputTable", RefersTo:="=" & address, Visible:=True

    Logger.Info "名前定義 rngInputTable を更新しました。"
End Sub

Private Sub EnsureEssentialSheets()
    EnsureSheet INPUT_SHEET
    EnsureSheet SUMMARY_SHEET
    EnsureSheet MASTER_SHEET
End Sub

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        Logger.Info "シートを新規作成しました: " & sheetName
    End If

    Set EnsureSheet = ws
End Function

Private Function WorksheetLastRow(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    WorksheetLastRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
End Function

Private Function JoinIssues(ByVal issues As Collection) As String
    Dim arr() As String
    Dim idx As Long
    ReDim arr(0 To issues.Count - 1)

    For idx = 1 To issues.Count
        arr(idx - 1) = CStr(issues(idx))
    Next idx

    JoinIssues = Join(arr, vbCrLf)
End Function
