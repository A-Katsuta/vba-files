Option Explicit
'==============================================================================='
' モジュール名: ModInitialSetup
'
' 【概要】マスタの内容を基に「月次データ」シートのヘッダーを整備し、
'         推奨作番・作業コードの一覧を初期化します。
'==============================================================================='

Public Sub RunInitialSetup()
    Dim prevState As ApplicationState
    Dim wsMonthly As Worksheet
    Dim wsMaster As Worksheet
    Dim protMonthly As SheetProtectionInfo
    Dim monthlyEditable As Boolean
    Dim masterCheckMessage As String
    Dim masterHasError As Boolean
    Dim masterRecords() As MasterRecord
    Dim masterRecordCount As Long
    Dim masterLoadWarnings As String
    Dim workNos As Object
    Dim workCodes As Object
    Dim pairDict As Object
    Dim workPairs As Collection
    Dim warnings As String
    Dim finalMessage As String
    Dim finalTitle As String
    Dim setupCompleted As Boolean
    Dim promptResult As VbMsgBoxResult
    Dim i As Long
    Dim rec As MasterRecord
    Dim key As Variant
    Dim pairKey As String

    On Error GoTo ErrHandler

    finalTitle = "初期設定"
    SaveAndSetApplicationState prevState

    Set wsMonthly = GetSheet(Sheet_Monthly)

    masterCheckMessage = EnsureMasterReady(wsMaster, masterHasError)
    If masterHasError Then
        finalMessage = "マスタの確認に失敗しました。" & vbCrLf & masterCheckMessage
        GoTo CleanUp
    End If
    If Len(masterCheckMessage) > 0 Then
        warnings = AppendLine(warnings, masterCheckMessage)
    End If

    masterRecordCount = LoadMasterWithCount(masterRecords, wsMaster)
    masterLoadWarnings = GetMasterLoadWarnings()
    If Len(masterLoadWarnings) > 0 Then
        warnings = AppendLine(warnings, masterLoadWarnings)
    End If
    If masterRecordCount = 0 Then
        finalMessage = "マスタに有効な行が登録されていません。作番と作業コードを入力してから再実行してください。"
        GoTo CleanUp
    End If

    Set workNos = CreateObject("Scripting.Dictionary")
    workNos.CompareMode = vbTextCompare
    Set workCodes = CreateObject("Scripting.Dictionary")
    workCodes.CompareMode = vbTextCompare
    Set pairDict = CreateObject("Scripting.Dictionary")
    pairDict.CompareMode = vbTextCompare

    For i = 1 To masterRecordCount
        rec = masterRecords(i)
        If Len(rec.WorkNumber) > 0 Then
            If Not workNos.Exists(rec.WorkNumber) Then workNos.Add rec.WorkNumber, True
        End If
        If Len(rec.WorkCode) > 0 Then
            If Not workCodes.Exists(rec.WorkCode) Then workCodes.Add rec.WorkCode, True
        End If
        If Len(rec.WorkNumber) > 0 And Len(rec.WorkCode) > 0 Then
            pairKey = rec.WorkNumber & "|" & rec.WorkCode
            If Not pairDict.Exists(pairKey) Then
                pairDict.Add pairKey, Array(rec.WorkNumber, rec.WorkCode)
            End If
        End If
    Next i

    If workNos.Count = 0 Or workCodes.Count = 0 Then
        warnings = AppendLine(warnings, "マスタに作番または作業コードが登録されていません。")
    End If

    Set workPairs = New Collection
    For Each key In pairDict.Keys
        workPairs.Add pairDict(key)
    Next key

    monthlyEditable = UnprotectSheetIfNeeded(wsMonthly, protMonthly)
    If Not monthlyEditable Then
        finalMessage = "月次データシートの保護解除がキャンセルされたため、処理を中断しました。"
        GoTo CleanUp
    End If

    UpdateMonthlyHeader wsMonthly, workPairs

    setupCompleted = True
    finalTitle = "初期設定完了"
    finalMessage = "マスタを基に初期設定を完了しました。" & vbCrLf & _
                   "・作番: " & workNos.Count & " 件" & vbCrLf & _
                   "・作業コード: " & workCodes.Count & " 件" & vbCrLf & _
                   "・月次マスタ列: " & workPairs.Count & " 列"
    If Len(warnings) > 0 Then
        finalMessage = finalMessage & vbCrLf & vbCrLf & "【注意】" & vbCrLf & warnings
    End If
    finalMessage = finalMessage & vbCrLf & vbCrLf & _
                   "マスタを表示しますか？"

CleanUp:
    On Error Resume Next
    If monthlyEditable Then RestoreSheetProtection wsMonthly, protMonthly
    RestoreApplicationState prevState
    On Error GoTo 0

    If setupCompleted Then
        promptResult = MsgBox(finalMessage, vbYesNo + vbQuestion, finalTitle)
        If promptResult = vbYes Then
            On Error Resume Next
            wsMaster.Activate
            Application.Goto ThisWorkbook.Names("Master_Table").RefersToRange.Cells(1, 1), True
            On Error GoTo 0
        End If
    ElseIf Len(finalMessage) > 0 Then
        MsgBox finalMessage, vbExclamation, finalTitle
    End If
    Exit Sub

ErrHandler:
    finalTitle = "初期設定エラー"
    finalMessage = "初期設定でエラーが発生しました: " & Err.Description
    Resume CleanUp
End Sub

Private Function AppendLine(ByVal baseText As String, ByVal addText As String) As String
    If Len(addText) = 0 Then
        AppendLine = baseText
    ElseIf Len(baseText) = 0 Then
        AppendLine = addText
    Else
        AppendLine = baseText & vbCrLf & addText
    End If
End Function

Private Sub UpdateMonthlyHeader(ByRef wsMonthly As Worksheet, ByRef pairs As Collection)
    Dim startCol As Long
    Dim headerLastCol As Long
    Dim colIndex As Long
    Dim pair As Variant
    Dim dataLastRow As Long

    startCol = MonthlyCol_Min + 2 ' C列(合計)とD列(見出し)を除外してE列から開始
    headerLastCol = wsMonthly.Cells(MonthlyRow_Header, wsMonthly.Columns.Count).End(xlToLeft).Column
    If headerLastCol < startCol - 1 Then
        headerLastCol = startCol - 1
    End If

    dataLastRow = wsMonthly.Cells(wsMonthly.Rows.Count, MonthlyCol_Date).End(xlUp).Row
    If dataLastRow < MonthlyRow_DataStart Then
        dataLastRow = MonthlyRow_DataStart
    End If

    colIndex = startCol
    For Each pair In pairs
        wsMonthly.Cells(MonthlyRow_WorkNo, colIndex).Value = pair(0)
        wsMonthly.Cells(MonthlyRow_Header, colIndex).Value = pair(1)
        colIndex = colIndex + 1
    Next pair

    If colIndex > startCol Then
        wsMonthly.Range(wsMonthly.Cells(MonthlyRow_DataStart, startCol), _
                        wsMonthly.Cells(dataLastRow, colIndex - 1)).ClearContents
    End If

    If headerLastCol >= colIndex Then
        wsMonthly.Range(wsMonthly.Cells(MonthlyRow_WorkNo, colIndex), _
                        wsMonthly.Cells(MonthlyRow_WorkNo, headerLastCol)).ClearContents
        wsMonthly.Range(wsMonthly.Cells(MonthlyRow_Header, colIndex), _
                        wsMonthly.Cells(MonthlyRow_Header, headerLastCol)).ClearContents
        wsMonthly.Range(wsMonthly.Cells(MonthlyRow_DataStart, colIndex), _
                        wsMonthly.Cells(dataLastRow, headerLastCol)).ClearContents
    End If
End Sub
