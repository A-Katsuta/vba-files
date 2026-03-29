Option Explicit

'=================================================================================
' モジュール名: ModMaster
'=================================================================================
'
'=================================================================================
' 【モジュール概要】
' マスタシートの雛形生成・読み込み・照合処理を提供するユーティリティ
'=================================================================================

'=================================================================================
' 【定数・構造体定義】
' マスタで参照する名前定義や列番号、照合結果の型をまとめる
'=================================================================================

Private Const NAME_MASTER_TABLE As String = "Master_Table"
Private Const NAME_MASTER_WORKNUMBER As String = "Master_WorkNumber"
Private Const NAME_MASTER_WORKCODE As String = "Master_WorkCode"
Private Const NAME_MASTER_WORKNUMBER_LOOKUP As String = "Master_WorkNumberLookup"
Private Const NAME_MASTER_WORKCODE_LOOKUP As String = "Master_WorkCodeLookup"
Private Const LISTOBJECT_MASTER As String = "tblMaster"
Private Const SHEET_NAME_MASTER As String = "マスタ"
Private Const MASTER_WORD_COLUMN_COUNT As Long = 20

Public Type MasterRecord
    ItemNumber As Long
    WorkNumber As String
    WorkNumberName As String
    WorkCode As String
    WorkCodeName As String
    Words() As String
    SourceRow As Long
End Type

Public Type MasterMatchResult
    Record As MasterRecord
    MatchedWord As String
    IsMatch As Boolean
    WarningMessage As String
End Type

Private Enum MasterColumn
    MasterColumn_ItemNo = 1
    MasterColumn_WorkNumber = 2
    MasterColumn_WorkNumberName = 3
    MasterColumn_WorkCode = 4
    MasterColumn_WorkCodeName = 5
    MasterColumn_FirstWord = 6
End Enum

Private mLastLoadWarnings As String

'=================================================================================
' 【初期化・雛形生成】
' マスタシートの作成と検証
'=================================================================================

Public Sub GenerateMasterSheet()
    Const PROC_TITLE As String = "マスタ生成"

    Dim appState As ApplicationState
    Dim wb As Workbook
    Dim wsMaster As Worksheet
    Dim ws As Worksheet
    Dim headers As Variant
    Dim headerRange As Range
    Dim lo As ListObject
    Dim nameDefinitions As Variant
    Dim i As Long
    Dim promptResult As VbMsgBoxResult
    Dim isCompleted As Boolean
    Dim tempRow As ListRow
    Dim addedTempRow As Boolean

    On Error GoTo ErrHandler

    headers = BuildMasterHeaders()
    Set wb = ThisWorkbook
    SaveAndSetApplicationState appState

    Set wsMaster = Nothing
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, SHEET_NAME_MASTER, vbBinaryCompare) = 0 Then
            Set wsMaster = ws
            Exit For
        End If
    Next ws

    If Not wsMaster Is Nothing Then
        promptResult = MsgBox("既存の「マスタ」シートを初期化して再作成します。すべてのデータが削除されますがよろしいですか？", _
                              vbYesNo + vbExclamation, PROC_TITLE)
        If promptResult <> vbYes Then
            GoTo CleanUp
        End If
        ResetMasterSheet wsMaster
    Else
        Set wsMaster = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsMaster.Name = SHEET_NAME_MASTER
    End If

    For i = LBound(headers) To UBound(headers)
        wsMaster.Cells(1, i + 1).Value = headers(i)
    Next i

    Set headerRange = wsMaster.Range(wsMaster.Cells(1, 1), wsMaster.Cells(2, UBound(headers) + 1))
    Set lo = wsMaster.ListObjects.Add(SourceType:=xlSrcRange, Source:=headerRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = LISTOBJECT_MASTER
    lo.DisplayName = LISTOBJECT_MASTER
    lo.ShowTotals = False
    If lo.ListRows.Count > 0 Then
        lo.ListRows(1).Delete
    End If

    nameDefinitions = Array( _
        Array(NAME_MASTER_TABLE, "=" & LISTOBJECT_MASTER & "[#All]"), _
        Array(NAME_MASTER_WORKNUMBER, "=" & LISTOBJECT_MASTER & "[作番]"), _
        Array(NAME_MASTER_WORKCODE, "=" & LISTOBJECT_MASTER & "[作業ｺｰﾄﾞ]") _
    )

    For i = LBound(nameDefinitions) To UBound(nameDefinitions)
        AssignWorkbookName CStr(nameDefinitions(i)(0)), CStr(nameDefinitions(i)(1))
    Next i

    EnsureLookupPlaceholders wsMaster

    If lo.ListRows.Count = 0 Then
        Set tempRow = lo.ListRows.Add
        addedTempRow = True
    End If

    If Not lo.DataBodyRange Is Nothing Then
        lo.ListColumns(MasterColumn_WorkNumberName).DataBodyRange.Formula = _
            "=IF([@作番]=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & _
            ",IFERROR(VLOOKUP([@作番]," & NAME_MASTER_WORKNUMBER_LOOKUP & ",2,FALSE)," & Chr(34) & Chr(34) & "))"
        lo.ListColumns(MasterColumn_WorkCodeName).DataBodyRange.Formula = _
            "=IF([@作業ｺｰﾄﾞ]=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & _
            ",IFERROR(VLOOKUP([@作業ｺｰﾄﾞ]," & NAME_MASTER_WORKCODE_LOOKUP & ",2,FALSE)," & Chr(34) & Chr(34) & "))"
    End If

    If addedTempRow Then
        tempRow.Delete
    End If

    wsMaster.Columns("A:Z").AutoFit
    On Error Resume Next
    wsMaster.Activate
    wsMaster.Cells(1, 1).Select
    On Error GoTo 0

    isCompleted = True

CleanUp:
    On Error Resume Next
    RestoreApplicationState appState
    On Error GoTo 0
    If isCompleted Then
        MsgBox "マスタシートの雛形を作成しました。必要なデータと照合ワードを登録してください。", vbInformation, PROC_TITLE
    End If
    Exit Sub

ErrHandler:
    MsgBox "マスタシートの生成中にエラーが発生しました: " & Err.Description, vbCritical, PROC_TITLE
    Resume CleanUp
End Sub

Private Sub EnsureLookupPlaceholders(ByVal wsMaster As Worksheet)
    Dim workNoRange As String
    Dim workCodeRange As String

    workNoRange = wsMaster.Columns("AA:AB").Address(True, True, xlA1, True)
    workCodeRange = wsMaster.Columns("AC:AD").Address(True, True, xlA1, True)

    If Len(GetCellString(wsMaster.Range("AA1").Value)) = 0 Then wsMaster.Range("AA1").Value = "作番キー"
    If Len(GetCellString(wsMaster.Range("AB1").Value)) = 0 Then wsMaster.Range("AB1").Value = "作番名称"
    If Len(GetCellString(wsMaster.Range("AC1").Value)) = 0 Then wsMaster.Range("AC1").Value = "作業ｺｰﾄﾞキー"
    If Len(GetCellString(wsMaster.Range("AD1").Value)) = 0 Then wsMaster.Range("AD1").Value = "作業ｺｰﾄﾞ名称"

    AssignWorkbookName NAME_MASTER_WORKNUMBER_LOOKUP, "=" & workNoRange
    AssignWorkbookName NAME_MASTER_WORKCODE_LOOKUP, "=" & workCodeRange
End Sub

Private Function BuildMasterHeaders() As Variant
    Dim headers() As String
    Dim i As Long

    ReDim headers(0 To MASTER_WORD_COLUMN_COUNT + MasterColumn_FirstWord - 2)
    headers(0) = "項番"
    headers(1) = "作番"
    headers(2) = "作番名称"
    headers(3) = "作業ｺｰﾄﾞ"
    headers(4) = "作業ｺｰﾄﾞ名称"

    For i = 1 To MASTER_WORD_COLUMN_COUNT
        headers(MasterColumn_FirstWord - 1 + i - 1) = "照合ﾜｰﾄﾞ" & CStr(i)
    Next i

    BuildMasterHeaders = headers
End Function

'=================================================================================
' 【マスタ準備状況の確認】
' 名前定義・テーブルの存在を検証し、致命的エラーや警告文をまとめて返す
'=================================================================================

Public Function EnsureMasterReady(ByRef wsMaster As Worksheet, ByRef hasError As Boolean) As String
    Dim masterRange As Range
    Dim lo As ListObject
    Dim messages As String
    Dim fatalError As Boolean
    Dim requiredNames As Variant
    Dim i As Long
    Dim nameRange As Range
    Dim nameMessage As String

    hasError = False
    Set wsMaster = Nothing
    messages = ""
    fatalError = False

    On Error Resume Next
    Set masterRange = ThisWorkbook.Names(NAME_MASTER_TABLE).RefersToRange
    If Err.Number <> 0 Then
        messages = AppendLine(messages, "名前定義「" & NAME_MASTER_TABLE & "」が見つかりません。マスタの設定を確認してください。")
        fatalError = True
        Err.Clear
    End If
    On Error GoTo 0

    If Not masterRange Is Nothing Then
        Set wsMaster = masterRange.Worksheet
        On Error Resume Next
        Set lo = wsMaster.ListObjects(LISTOBJECT_MASTER)
        On Error GoTo 0
        If lo Is Nothing Then
            messages = AppendLine(messages, "マスタのテーブル """ & LISTOBJECT_MASTER & """ が見つかりません。")
            fatalError = True
        ElseIf lo.ListRows.Count = 0 Then
            messages = AppendLine(messages, "マスタが空です。最低1行以上のデータを登録してください。")
            fatalError = True
        End If
    Else
        fatalError = True
    End If

    requiredNames = Array(NAME_MASTER_WORKNUMBER, NAME_MASTER_WORKCODE)

    For i = LBound(requiredNames) To UBound(requiredNames)
        If Not masterRange Is Nothing Then
            If Not TryGetMasterNameRange(CStr(requiredNames(i)), wsMaster, nameRange, nameMessage) Then
                messages = AppendLine(messages, nameMessage)
                fatalError = True
            End If
        Else
            Exit For
        End If
    Next i

    hasError = fatalError
    EnsureMasterReady = messages
End Function

Private Function TryGetMasterNameRange(ByVal nameText As String, ByVal wsMaster As Worksheet, _
                                       ByRef nameRange As Range, ByRef message As String) As Boolean
    On Error Resume Next
    Set nameRange = ThisWorkbook.Names(nameText).RefersToRange
    If Err.Number <> 0 Then
        message = "名前定義「" & nameText & "」が見つかりません。"
        Err.Clear
        TryGetMasterNameRange = False
        Exit Function
    End If
    On Error GoTo 0

    If nameRange Is Nothing Then
        message = "名前定義「" & nameText & "」の参照先が不正です。"
        TryGetMasterNameRange = False
    ElseIf Not wsMaster Is Nothing And Not nameRange.Worksheet Is wsMaster Then
        message = "名前定義「" & nameText & "」はマスタと同じシートを参照していません。"
        TryGetMasterNameRange = False
    Else
        message = ""
        TryGetMasterNameRange = True
    End If
End Function

'=================================================================================
' 【データ読み込み処理】
' マスタ行を配列へ読み込み、照合で利用できる形に整える
'=================================================================================

Public Sub LoadMaster(ByRef records() As MasterRecord, ByRef recordCount As Long, ByRef wsMaster As Worksheet)
    recordCount = LoadMasterCore(records, wsMaster)
End Sub

Public Function LoadMasterWithCount(ByRef records() As MasterRecord, ByRef wsMaster As Worksheet) As Long
    LoadMasterWithCount = LoadMasterCore(records, wsMaster)
End Function

Private Function LoadMasterCore(ByRef records() As MasterRecord, ByRef wsMaster As Worksheet) As Long
    Dim lo As ListObject
    Dim dataRange As Range
    Dim values As Variant
    Dim r As Long
    Dim record As MasterRecord
    Dim words() As String
    Dim warnings As String
    Dim recordCount As Long
    Dim lastWordColumn As Long
    Dim hasContent As Boolean

    recordCount = 0
    Erase records
    mLastLoadWarnings = ""

    Set lo = wsMaster.ListObjects(LISTOBJECT_MASTER)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set dataRange = lo.DataBodyRange
    values = dataRange.Value
    lastWordColumn = MasterColumn_FirstWord + MASTER_WORD_COLUMN_COUNT - 1

    ReDim records(1 To UBound(values, 1))

    For r = 1 To UBound(values, 1)
        record = DefaultMasterRecord()
        record.SourceRow = dataRange.Row + r - 1

        If ColumnExists(values, MasterColumn_ItemNo) Then
            If IsNumeric(values(r, MasterColumn_ItemNo)) Then
                record.ItemNumber = CLng(values(r, MasterColumn_ItemNo))
            ElseIf Len(GetCellString(values(r, MasterColumn_ItemNo))) > 0 Then
                warnings = AppendLine(warnings, "マスタ行" & record.SourceRow & " の項番が数値ではありません。0として扱います。")
            End If
        End If

        record.WorkNumber = GetCellString(values(r, MasterColumn_WorkNumber))
        record.WorkNumberName = GetCellString(values(r, MasterColumn_WorkNumberName))
        record.WorkCode = GetCellString(values(r, MasterColumn_WorkCode))
        record.WorkCodeName = GetCellString(values(r, MasterColumn_WorkCodeName))

        words = CollectMasterWords(values, r, lastWordColumn)
        If IsArrayInitialized(words) Then
            record.Words = words
        End If

        hasContent = (Len(record.WorkNumber) > 0) Or (Len(record.WorkCode) > 0) Or IsArrayInitialized(record.Words)
        If hasContent Then
            recordCount = recordCount + 1
            records(recordCount) = record
        End If
    Next r

    If recordCount > 0 Then
        ReDim Preserve records(1 To recordCount)
    Else
        Erase records
    End If

    mLastLoadWarnings = Trim$(warnings)
    LoadMasterCore = recordCount
End Function

Private Function ColumnExists(ByRef values As Variant, ByVal columnIndex As Long) As Boolean
    On Error GoTo ErrHandler
    Dim dummy As Variant
    dummy = values(1, columnIndex)
    ColumnExists = True
    Exit Function
ErrHandler:
    ColumnExists = False
End Function

Public Function GetMasterLoadWarnings() As String
    GetMasterLoadWarnings = mLastLoadWarnings
End Function

'=================================================================================
' 【照合処理】
' 件名文字列とマスタを突き合わせて最適な行を抽出
'=================================================================================

Public Sub MatchSubjectWithMaster(ByVal subject As String, ByRef records() As MasterRecord, ByVal recordCount As Long, _
                                  ByRef result As MasterMatchResult)
    Dim normalizedSubject As String
    Dim i As Long
    Dim wordIndex As Long
    Dim candidate As String

    result = DefaultMatchResult()
    normalizedSubject = Trim$(subject)
    If Len(normalizedSubject) = 0 Then Exit Sub
    If recordCount = 0 Then Exit Sub

    For i = 1 To recordCount
        If IsArrayInitialized(records(i).Words) Then
            For wordIndex = LBound(records(i).Words) To UBound(records(i).Words)
                candidate = records(i).Words(wordIndex)
                If Len(candidate) = 0 Then GoTo NextWord
                If InStr(1, normalizedSubject, candidate, vbTextCompare) > 0 Then
                    result.Record = records(i)
                    result.MatchedWord = candidate
                    result.IsMatch = True
                    Exit Sub
                End If
NextWord:
            Next wordIndex
        End If
    Next i
End Sub

'=================================================================================
' 【内部ユーティリティ】
'=================================================================================

Private Sub ResetMasterSheet(ByRef wsMaster As Worksheet)
    Dim lo As ListObject

    If wsMaster Is Nothing Then Exit Sub
    If wsMaster.ProtectContents Then
        Err.Raise vbObjectError + 750, "ResetMasterSheet", "マスタのシート保護を解除してから再作成してください。"
    End If

    For Each lo In wsMaster.ListObjects
        lo.Delete
    Next lo
    wsMaster.Cells.Clear
    wsMaster.Cells.ClearFormats
End Sub

Private Sub AssignWorkbookName(ByVal nameText As String, ByVal refersTo As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersTo
End Sub

Private Function CollectMasterWords(ByRef values As Variant, ByVal rowIndex As Long, ByVal lastWordColumn As Long) As String()
    Dim temp() As String
    Dim count As Long
    Dim c As Long
    Dim candidate As String
    Dim upperBound As Long

    On Error GoTo ErrHandler
    upperBound = Application.WorksheetFunction.Min(UBound(values, 2), lastWordColumn)

    ReDim temp(0 To MASTER_WORD_COLUMN_COUNT - 1)
    count = -1

    For c = MasterColumn_FirstWord To upperBound
        candidate = GetCellString(values(rowIndex, c))
        If Len(candidate) > 0 Then
            count = count + 1
            temp(count) = candidate
        End If
    Next c

    If count >= 0 Then
        ReDim Preserve temp(0 To count)
        CollectMasterWords = temp
    End If
    Exit Function

ErrHandler:
    Erase temp
End Function

Private Function GetCellString(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    GetCellString = Trim$(CStr(value))
End Function

Private Function AppendLine(ByVal baseText As String, ByVal addText As String) As String
    If Len(addText) = 0 Then
        AppendLine = baseText
    ElseIf Len(baseText) = 0 Then
        AppendLine = addText
    Else
        AppendLine = baseText & vbCrLf & addText
    End If
End Function

Private Function DefaultMasterRecord() As MasterRecord
    Dim record As MasterRecord
    record.ItemNumber = 0
    record.WorkNumber = ""
    record.WorkNumberName = ""
    record.WorkCode = ""
    record.WorkCodeName = ""
    record.SourceRow = 0
    DefaultMasterRecord = record
End Function

Private Function DefaultMatchResult() As MasterMatchResult
    Dim result As MasterMatchResult
    result.IsMatch = False
    result.MatchedWord = ""
    result.WarningMessage = ""
    DefaultMatchResult = result
End Function

Private Function IsArrayInitialized(ByRef arr() As String) As Boolean
    On Error GoTo ErrHandler
    If (LBound(arr) <= UBound(arr)) Then
        IsArrayInitialized = True
    End If
    Exit Function
ErrHandler:
    IsArrayInitialized = False
End Function

