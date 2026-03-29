Attribute VB_Name = "Utils"
Option Explicit

Public Function Nz(ByVal v As Variant, Optional ByVal def As String = "") As String
    On Error GoTo EH
    If IsError(v) Then Nz = def: Exit Function
    If IsNull(v) Then Nz = def: Exit Function
    If VarType(v) = vbString Then
        If LenB(v) = 0 Then Nz = def Else Nz = CStr(v)
    Else
        Nz = CStr(v)
    End If
    Exit Function
EH:
    Nz = def
End Function

Public Function TryParseNumber(ByVal value As Variant, Optional ByVal defaultValue As Double = 0) As Double
    On Error GoTo Fallback
    If IsNumeric(value) Then
        TryParseNumber = CDbl(value)
    Else
        TryParseNumber = defaultValue
    End If
    Exit Function
Fallback:
    TryParseNumber = defaultValue
End Function

Public Function FormatIsoDate(ByVal value As Variant) As String
    On Error GoTo EH
    If IsDate(value) Then
        FormatIsoDate = Format$(CDate(value), "yyyy-mm-dd")
    Else
        FormatIsoDate = ""
    End If
    Exit Function
EH:
    FormatIsoDate = ""
End Function

Public Function JoinRange(ByVal target As Range, Optional ByVal delimiter As String = ", ") As String
    Dim cell As Range
    Dim parts As Collection
    Set parts = New Collection

    If target Is Nothing Then Exit Function

    For Each cell In target.Cells
        If LenB(Trim$(CStr(cell.Value))) > 0 Then
            parts.Add Trim$(CStr(cell.Value))
        End If
    Next cell

    JoinRange = Join(CollectionToArray(parts), delimiter)
End Function

Private Function CollectionToArray(ByVal col As Collection) As Variant
    Dim arr() As String
    Dim idx As Long
    ReDim arr(0 To IIf(col Is Nothing, 0, col.Count) - 1)

    If col Is Nothing Then
        CollectionToArray = arr
        Exit Function
    End If

    For idx = 1 To col.Count
        arr(idx - 1) = CStr(col.Item(idx))
    Next idx

    CollectionToArray = arr
End Function

Public Function SafeSheet(ByVal sheetName As String) As Worksheet
    On Error GoTo NotFound
    Set SafeSheet = ThisWorkbook.Worksheets(sheetName)
    Exit Function
NotFound:
    Err.Raise vbObjectError + 1024, "Utils.SafeSheet", "シート '" & sheetName & "' が見つかりません。"
End Function
