Attribute VB_Name = "Logger"
Option Explicit

Private Const LOG_SHEET_NAME As String = "Logs"
Private Const LOG_DATE_FORMAT As String = "yyyy/mm/dd hh:nn:ss"
Private Const XL_SHEET_VERY_HIDDEN As Long = 2

Private mIsInitialized As Boolean

Public Sub Initialize()
    EnsureLogSheet
    mIsInitialized = True
    Info "Logger が初期化されました。"
End Sub

Public Sub Info(ByVal message As String)
    WriteEntry "INFO", message
End Sub

Public Sub Warn(ByVal message As String)
    WriteEntry "WARN", message
End Sub

Public Sub [Error](ByVal message As String)
    WriteEntry "ERROR", message
End Sub

Public Sub Success(ByVal message As String)
    WriteEntry "SUCCESS", message
End Sub

Public Sub ExportToFile(ByVal targetPath As String)
    Dim fs As Integer
    Dim ws As Worksheet
    Set ws = EnsureLogSheet

    fs = FreeFile
    Open targetPath For Output As #fs

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim rowIndex As Long
    For rowIndex = 2 To lastRow
        Print #fs, ws.Cells(rowIndex, 1).Value & Chr$(9) & ws.Cells(rowIndex, 2).Value & Chr$(9) & ws.Cells(rowIndex, 3).Value
    Next rowIndex

    Close #fs
    Info "ログをファイルに書き出しました: " & targetPath
End Sub

Private Sub WriteEntry(ByVal level As String, ByVal message As String)
    Dim ws As Worksheet
    Set ws = EnsureLogSheet

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Format$(Now, LOG_DATE_FORMAT)
    ws.Cells(nextRow, 2).Value = level
    ws.Cells(nextRow, 3).Value = message
    ws.Cells(nextRow, 4).Value = Environ$("USERNAME")

    Debug.Print "[" & level & "] " & message
End Sub

Private Function EnsureLogSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
        ws.Visible = XL_SHEET_VERY_HIDDEN

        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Level"
        ws.Cells(1, 3).Value = "Message"
        ws.Cells(1, 4).Value = "User"
        ws.Range("A1:D1").Font.Bold = True
    End If

    Set EnsureLogSheet = ws
End Function
