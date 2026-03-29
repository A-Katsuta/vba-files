Option Explicit
' モジュール名: ModGetOutlookSch
'
' 【概要】  Outlookから指定した日付の予定を取得し、Excelシートに出力します。
'          件名に含まれるキーワードを基に、推奨作番・作業コードを自動判定する
'          機能も持ちます。
' 【作成】「JJ-07」2025/08
' 【対象環境】Excel 2016+ / Windows
' 【前提条件】
' ・参照設定「Microsoft Outlook XX.X Object Library」が有効であること。
' ・マスタ（`tblMaster`）と `Master_*` 系名前定義
'   （例: `Master_Table`、`Master_WorkNumber`、
'    `Master_WorkCode` など）が設定されていること。

' 【メイン処理】Outlook予定取得＆Excel出力
Public Sub GetOutlookSchedule()

    '============================================================
    ' ■ 1. 初期設定と定数宣言
    '============================================================
    
    ' --- ユーザー設定項目 ---
    Const DATE_INPUT_CELL   As String = "C3"              ' 日付が入力されているセル番地
    Const OUTPUT_HEADER_ROW As Long = 7                   ' ヘッダー行の行番号
    Const OUTPUT_START_COLUMN As String = "C"             ' 出力先の開始列

    ' --- データ転記に関する設定項目 ---
    Const SOURCE_CELL     As String = "C4"                ' 転記元の日付セル（データ取得シート）

    ' --- 出力列の列番号定数（C列=3） ---
    Const COL_TIME As Long = 3      ' C列: 時間
    Const COL_SUBJECT As Long = 4     ' D列: 件名
    Const COL_DURATION As Long = 5    ' E列: 会議時間（"HHMM"形式）
    Const COL_WORK_NO As Long = 6         ' F列: 推奨作番（マスタ照合結果）
    Const COL_WORK_NO_NAME As Long = 7    ' G列: 推奨作番名称（マスタ照合結果）
    Const COL_WORK_CODE As Long = 8       ' H列: 推奨作業コード（マスタ照合結果）
    Const COL_WORK_CODE_NAME As Long = 9  ' I列: 推奨作業コード名称（マスタ照合結果）
    Const COL_MATCH_WORD As Long = 10     ' J列: 照合ワード

    ' --- 変数宣言 ---
    ' --- Excelオブジェクト関連 ---
    Dim ws As Worksheet, wsDest As Worksheet        ' 操作対象のワークシートオブジェクト
    Dim protInfo As SheetProtectionInfo             ' シート保護情報
    Dim cellValue As Variant                        ' 日付セルの値を一時的に格納

    ' --- Outlookオブジェクト関連 ---
    Dim olApp As Object, olNs As Object, olFolder As Object ' Outlook基本オブジェクト
    Dim olItems As Object, olRestrictedItems As Object      ' 予定アイテムコレクション
    Dim olApt As Object                                     ' 個別の予定アイテム

    ' --- 処理制御用 ---
    Dim targetDate As Date                          ' 取得対象の日付
    Dim outputRow As Long, lastOutputRow As Long    ' 出力先の行番号管理
    Dim actualCount As Long                         ' 実際に取得した予定の件数
    Dim excludedCount As Long                       ' 表への入力を除外した予定の件数
    Dim warnMsg As String                           ' 処理中の警告メッセージを格納

    ' --- マスタ関連 ---
    Dim wsMaster As Worksheet
    Dim masterRecords() As MasterRecord
    Dim masterRecordCount As Long
    Dim masterCheckMessage As String
    Dim masterHasError As Boolean
    Dim masterLoadWarnings As String
    Dim masterResult As MasterMatchResult
    Dim rngExclude As Range                        ' 除外キーワードの名前付き範囲
    Dim arrExclude As Variant                      ' 除外キーワード配列

    ' --- ステップ1：実行前設定 ---
    On Error GoTo ErrorHandler          ' エラーハンドラを有効化
    Application.ScreenUpdating = False  ' 処理中の画面描画を停止し、高速化

    '============================================================
    ' ■ 2. 実行前チェックと準備
    '============================================================

    ' --- ステップ2：ワークシートオブジェクトの取得 ---
    Set ws = GetSheet(Sheet_DataAcquire)

    ' --- ステップ3：シート保護の確認と一時解除 ---
    If Not UnprotectSheetIfNeeded(ws, protInfo) Then GoTo CleanUp

    ' --- ステップ4：日付入力のチェック ---
    cellValue = ws.Range(DATE_INPUT_CELL).Value
    If IsEmpty(cellValue) Or cellValue = "" Then
        MsgBox "セル " & DATE_INPUT_CELL & " が空欄です。日付を入力してください。", vbExclamation, "入力エラー"
        GoTo CleanUp
    End If
    If Not IsDate(cellValue) Then
        MsgBox "セル " & DATE_INPUT_CELL & " には正しい日付を入力してください。", vbExclamation, "入力エラー"
        GoTo CleanUp
    End If
    targetDate = CDate(cellValue)

    '============================================================
    ' ■ 3. 出力範囲クリア & ヘッダー設定
    '============================================================
    
    ' --- ステップ5：前回の出力データをクリア ---
    outputRow = OUTPUT_HEADER_ROW + 1
    lastOutputRow = ws.Cells(ws.rows.Count, COL_MATCH_WORD).End(xlUp).Row
    If lastOutputRow < OUTPUT_HEADER_ROW Then lastOutputRow = OUTPUT_HEADER_ROW

    ' --- データ存在チェック：ヘッダー行より下に出力があればクリア実行 ---
    If lastOutputRow >= outputRow Then
        ws.Range(ws.Cells(outputRow, COL_TIME), ws.Cells(lastOutputRow, COL_MATCH_WORD)).ClearContents
    End If

    ' --- ステップ6：ヘッダー行の再設定 ---
    ws.Cells(OUTPUT_HEADER_ROW, COL_TIME).Value = "時間"
    ws.Cells(OUTPUT_HEADER_ROW, COL_SUBJECT).Value = "件名"
    ws.Cells(OUTPUT_HEADER_ROW, COL_DURATION).Value = "会議時間"
    ws.Cells(OUTPUT_HEADER_ROW, COL_WORK_NO).Value = "推奨作番"
    ws.Cells(OUTPUT_HEADER_ROW, COL_WORK_NO_NAME).Value = "推奨作番名称"
    ws.Cells(OUTPUT_HEADER_ROW, COL_WORK_CODE).Value = "推奨作業コード"
    ws.Cells(OUTPUT_HEADER_ROW, COL_WORK_CODE_NAME).Value = "推奨作業コード名称"
    ws.Cells(OUTPUT_HEADER_ROW, COL_MATCH_WORD).Value = "照合ﾜｰﾄﾞ"
    ws.Range(ws.Cells(OUTPUT_HEADER_ROW, COL_TIME), ws.Cells(OUTPUT_HEADER_ROW, COL_MATCH_WORD)).Font.Bold = True

'============================================================
' ■ 4. マスタの読み込み
'============================================================

    ' --- ステップ7：マスタの存在確認 ---
    masterCheckMessage = EnsureMasterReady(wsMaster, masterHasError)
    If masterHasError Then
        MsgBox masterCheckMessage, vbCritical, "マスタ確認エラー"
        GoTo CleanUp
    End If
    If Len(masterCheckMessage) > 0 Then
        warnMsg = AppendLine(warnMsg, masterCheckMessage)
    End If

    ' --- ステップ8：マスタの読込 ---
    masterRecordCount = LoadMasterWithCount(masterRecords, wsMaster)
    masterLoadWarnings = GetMasterLoadWarnings()
    If Len(masterLoadWarnings) > 0 Then
        warnMsg = AppendLine(warnMsg, masterLoadWarnings)
    End If
    If masterRecordCount = 0 Then
        MsgBox "マスタに有効な行が登録されていません。", vbExclamation, "マスタ"
        GoTo CleanUp
    End If

    ' --- ステップ9：除外キーワードの名前付き範囲を取得 ---
    If TryGetNamedRange("ExcludeKeywords", rngExclude, warnMsg) Then
        arrExclude = To2DArray(rngExclude)
    Else
        arrExclude = Empty
    End If
    
    '============================================================
    ' ■ 5. Outlook 接続
    '============================================================

    ' --- ステップ11：Outlookアプリケーションへの接続 ---
    On Error Resume Next ' ※Outlookが起動していない場合に備え、エラーを一時的に無視
    Set olApp = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set olApp = CreateObject("Outlook.Application")
    End If
    If Err.Number <> 0 Then
        MsgBox "Outlookアプリケーションに接続できません。", vbCritical, "Outlook接続エラー"
        Err.Clear
        GoTo CleanUp
    End If
    Err.Clear
    On Error GoTo ErrorHandler ' エラーハンドリングを通常に戻す
    If olApp Is Nothing Then
        MsgBox "Outlookに接続できません。", vbCritical, "Outlook接続エラー"
        GoTo CleanUp
    End If

    '============================================================
    ' ■ 6. 指定日の予定を取得
    '============================================================

    ' --- ステップ12：予定表フォルダへのアクセスと予定の絞り込み ---
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.GetDefaultFolder(9) ' 9 = olFolderCalendar (予定表フォルダ)
    Set olItems = olFolder.items
    
    ' --- 予定を時系列にソートし、定期的な予定も対象に含める ---
    olItems.Sort "[Start]"
    olItems.IncludeRecurrences = True

    ' --- 指定日に少しでもかかる予定をすべて抽出するフィルタ文字列を作成 ---
    Dim filterString As String
    filterString = "[Start] <= '" & Format(targetDate, "yyyy/MM/dd 23:59") & "' AND [End] >= '" & Format(targetDate, "yyyy/MM/dd 00:00") & "'"
    Set olRestrictedItems = olItems.Restrict(filterString)

    '============================================================
    ' ■ 7. Excelへ出力（分類・区分の判定を追加）
    '============================================================
    
    ' --- ステップ13：取得した予定をExcelシートに出力 ---
    If olRestrictedItems.Count = 0 Then
        ' --- 予定がない場合の処理 ---
        ws.Cells(outputRow, COL_TIME).Value = "予定はありません"
        MsgBox Format(targetDate, "yyyy年mm月dd日") & " の予定はありませんでした。" & _
               vbCrLf & "除外された予定: 0 件", vbInformation, "処理完了"
    Else
        ' --- 予定がある場合のループ処理 ---
        actualCount = 0
        excludedCount = 0
        Dim subj As String
        For Each olApt In olRestrictedItems
            subj = NzString(olApt.Subject)
            If ShouldExclude(subj, arrExclude) Then
                excludedCount = excludedCount + 1
                GoTo NextItem
            End If

            actualCount = actualCount + 1

            ' --- 予定情報の書き込み ---
            ws.Cells(outputRow, COL_TIME).Value = Format(olApt.Start, "hhmm") & "-" & Format(olApt.End, "hhmm") ' 時間
            ws.Cells(outputRow, COL_SUBJECT).Value = subj ' 件名

              ' --- 会議時間を "HHMM" 形式で計算・書式設定 ---
              Dim totalMinutes As Long, hours As Long, minutes As Long
              totalMinutes = DateDiff("n", olApt.Start, olApt.End)
              hours = totalMinutes \ 60
              minutes = totalMinutes Mod 60
              With ws.Cells(outputRow, COL_DURATION)
                  .NumberFormat = "@" ' 文字列として設定
                  .Value = Format(hours, "00") & Format(minutes, "00")
              End With

              ' --- マスタ照合結果の反映 ---
              MatchSubjectWithMaster subj, masterRecords, masterRecordCount, masterResult

              If masterResult.IsMatch Then
                  ws.Cells(outputRow, COL_WORK_NO).Value = masterResult.Record.WorkNumber
                  ws.Cells(outputRow, COL_WORK_CODE).Value = masterResult.Record.WorkCode
                  ws.Cells(outputRow, COL_WORK_NO_NAME).Value = masterResult.Record.WorkNumberName
                  ws.Cells(outputRow, COL_WORK_CODE_NAME).Value = masterResult.Record.WorkCodeName
                  ws.Cells(outputRow, COL_MATCH_WORD).Value = masterResult.MatchedWord
              End If

              If Not masterResult.IsMatch Then
                  warnMsg = AppendLine(warnMsg, "行" & outputRow & " の件名に一致する照合ワードが見つかりません。")
              End If

              outputRow = outputRow + 1
NextItem:
          Next olApt

        ' --- 完了メッセージの表示 ---
        Dim doneMsg As String
        doneMsg = Format(targetDate, "yyyy年mm月dd日") & " の予定を " & actualCount & " 件取得しました。"
        doneMsg = doneMsg & vbCrLf & "除外された予定: " & excludedCount & " 件"
        ' ※名前付き範囲に関する警告があれば、完了メッセージに追記
        If Len(warnMsg) > 0 Then
            doneMsg = doneMsg & vbCrLf & "（注意）" & vbCrLf & Trim$(warnMsg)
        End If
        MsgBox doneMsg, vbInformation, "処理完了"
    End If

    '============================================================
    ' ■ 8. データ転記処理（データ取得!C4 → データ登録!D4）
    '============================================================

    ' --- ステップ14：取得日を「データ登録」シートへ転記 ---
    On Error Resume Next ' ※転記先シートが存在しない場合もエラーにしない
    Set wsDest = GetSheet(Sheet_DataEntry)
    If Err.Number <> 0 Then
        MsgBox "データ転記先シートの取得に失敗しました。", vbExclamation, "シート取得エラー"
        Err.Clear
        GoTo CleanUp
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    If Not wsDest Is Nothing Then
        ' --- 転記元に値がある場合のみ実行 ---
        If NzString(ws.Range(SOURCE_CELL).Value) <> "" Then
            ' ※重要：wsDestのシート保護は考慮していないため、必要に応じて解除/再保護処理を追加すること
            wsDest.Range(DATA_ENTRY_DATE_CELL).Value = ws.Range(SOURCE_CELL).Value
        End If
    End If

    GoTo CleanUp

' 【エラーハンドラ・終了処理セクション】
ErrorHandler:
    ' --- エラーハンドリング ---
    Dim errorTitle As String, errorMsg As String
    errorTitle = "エラーが発生しました"
    Select Case Err.Number
        Case 9 ' シートが見つからない
            errorMsg = "シート「" & GetSheetName(Sheet_DataAcquire) & "」または「" & GetSheetName(Sheet_DataEntry) & "」が見つかりませんでした。"
        Case 13 ' 型が一致しない（日付変換エラーなど）
            errorMsg = "セル「" & DATE_INPUT_CELL & "」の値を日付として認識できません。"
        Case 287, -2147467259, -2147221233 ' Outlook関連のエラー
            errorMsg = "Outlookへのアクセスで問題が発生しました。"
        Case vbObjectError ' Outlook起動失敗
            errorMsg = "Outlookアプリケーションの起動に失敗しました。"
        Case Else ' その他の予期せぬエラー
            errorMsg = "予期しないエラーが発生しました。" & vbCrLf & _
                       "エラー番号: " & Err.Number & vbCrLf & _
                       "エラー内容: " & Err.description
    End Select
    MsgBox errorMsg, vbCritical, errorTitle
    '（エラー発生時も必ず後始末処理へ）

CleanUp:
    ' --- 終了処理 ---
    ' --- シート保護状態を元に戻す ---
    If Not ws Is Nothing Then
        RestoreSheetProtection ws, protInfo
    End If

    ' --- オブジェクト変数の解放 ---
    On Error Resume Next ' 解放時のエラーは無視
    Set olApt = Nothing
    Set olRestrictedItems = Nothing
    Set olItems = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    Set wsDest = Nothing
    Set ws = Nothing
    If Err.Number <> 0 Then
        MsgBox "オブジェクト解放中にエラーが発生しました: " & Err.Description, _
               vbExclamation, "解放エラー"
        Err.Clear
    End If
    Err.Clear
    
    ' --- Excelの設定を元に戻す（画面描画を最後に有効化） ---
    Application.ScreenUpdating = True
End Sub


' 【機能名】実行用サブルーチン
' 【概要】  GetOutlookScheduleプロシージャを呼び出す。
'           Excelのボタンなどに登録することを想定したエントリーポイント。
Public Sub ExecuteOutlookSchedule()
    Call GetOutlookSchedule()
End Sub

' 【内部ヘルパー関数セクション】
' メイン処理から呼び出される補助的な機能

' 【機能名】名前付き範囲の安全な取得
' 【概要】  指定された名前付き範囲を取得する。取得に失敗した場合はFalseを返し、
'           引数の警告メッセージ用変数(warn)に情報を追記する。
' 【引数】  nameStr (String): 取得対象の名前付き範囲の名前
'           rng (Range)     : [出力]取得したRangeオブジェクトを格納する変数
'           warn (String)   : [入出力]警告メッセージを追記する変数
' 【戻り値】Boolean: 範囲の取得に成功した場合はTrue、失敗した場合はFalse
Private Function TryGetNamedRange(ByVal nameStr As String, ByRef rng As Range, ByRef warn As String) As Boolean
    On Error Resume Next
    Set rng = Nothing
    Set rng = ThisWorkbook.Names(nameStr).RefersToRange
    If Err.Number <> 0 Then
        MsgBox "名前付き範囲 """ & nameStr & """ の取得でエラーが発生しました。", _
               vbExclamation, "範囲取得エラー"
        warn = warn & IIf(Len(warn) > 0, vbCrLf, "") & _
               "名前付き範囲 """ & nameStr & """ の取得に失敗しました。"
        Err.Clear
        On Error GoTo 0
        TryGetNamedRange = False
        Exit Function
    End If
    On Error GoTo 0
    Err.Clear

    If rng Is Nothing Then
        TryGetNamedRange = False
        warn = warn & IIf(Len(warn) > 0, vbCrLf, "") & "名前付き範囲 """ & nameStr & """ が見つかりません。"
    Else
        TryGetNamedRange = True
    End If
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

' 【機能名】Rangeオブジェクトから2次元配列への変換
' 【概要】  RangeオブジェクトをVariant型の2次元配列に変換する。
'           対象が単一セルの場合でも、(1 To 1, 1 To 1)の2次元配列として返す。
' 【引数】  rng (Range): 変換元のRangeオブジェクト
' 【戻り値】Variant: 変換後の2次元配列
Private Function To2DArray(ByVal rng As Range) As Variant
    Dim v As Variant
    If rng.Cells.Count = 1 Then
        ' --- 単一セルの場合、強制的に2次元配列を作成 ---
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = rng.Value
        To2DArray = v
    Else
        ' --- 複数セルの場合、Valueプロパティで一括取得 ---
        To2DArray = rng.Value
    End If
End Function

' 【機能名】Null安全な文字列変換
' 【概要】  Variant型の値を文字列に変換する。Null、Empty、Errorを空文字("")に変換する。
' 【引数】  v (Variant): 変換対象の値
' 【戻り値】String: 変換後の文字列
Private Function NzString(ByVal v As Variant) As String
    If IsError(v) Then
        NzString = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzString = ""
    Else
        NzString = CStr(v)
    End If
End Function

' 【機能名】除外キーワード判定
' 【概要】  件名(subject)が除外キーワード配列(excludeArr)のいずれかに
'           部分一致するかを判定する。
' 【引数】  subject (String)  : 判定対象の件名
'           excludeArr (Variant): 除外キーワードの2次元配列 (行,1)
' 【戻り値】Boolean: 部分一致した場合はTrue、しない場合はFalse
Private Function ShouldExclude(ByVal subject As String, ByRef excludeArr As Variant) As Boolean
    Dim subjectNormalized As String
    Dim r As Long
    Dim c As Long
    Dim kw As String
    Dim rowLBound As Long
    Dim rowUBound As Long
    Dim colLBound As Long
    Dim colUBound As Long

    subjectNormalized = NzString(subject)
    If Len(subjectNormalized) = 0 Then Exit Function
    If IsEmpty(excludeArr) Then Exit Function

    rowLBound = LBound(excludeArr, 1)
    rowUBound = UBound(excludeArr, 1)
    colLBound = LBound(excludeArr, 2)
    colUBound = UBound(excludeArr, 2)

    For r = rowLBound To rowUBound
        For c = colLBound To colUBound
            kw = Trim$(NzString(excludeArr(r, c)))
            If Len(kw) > 0 Then
                If InStr(1, subjectNormalized, kw, vbTextCompare) > 0 Then
                    ShouldExclude = True
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

