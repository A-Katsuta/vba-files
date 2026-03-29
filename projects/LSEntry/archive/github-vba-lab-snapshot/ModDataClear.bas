Option Explicit

'=================================================================================
' 【モジュール概要】
' データ取得シートおよびデータ登録シートの入力値を一括クリアする機能群
'=================================================================================

'=================================================================================
' 【定数定義】
' クリア対象セル・範囲の設定値（変更時はここを更新する）
'=================================================================================

' --- シート名は ModAppConfig.bas の SheetName Enum を使用 ---

' --- クリア対象セル/範囲定数 ---
' データ取得シート
Private Const DATE_CELL_GETOUT    As String = "C4"      ' 任意日付セル
Private Const CLEAR_RANGE_ACQ     As String = "C8:J22"  ' データ範囲（時間～照合ﾜｰﾄﾞ）
' データ登録シート
Private Const DATE_CELL_WORKTIME  As String = "E24"     ' 勤務時間セル
Private Const CLEAR_RANGE_DATA    As String = "F8:F22"  ' データ範囲

'=================================================================================
' 【メイン処理】
' ユーザー操作により入力欄を初期化する手順
'=================================================================================

'=================================================================================
' 【機能名】入力データ一括クリア
' データ取得・データ登録両シートの入力範囲を安全に初期化する
' ※重要：シート保護状態は一時解除後に必ず復元する
'=================================================================================
Public Sub ClearInputData()
    ' --- 変数宣言 ---
    Dim wsAcq As Worksheet                  ' 「データ取得」シートオブジェクト
    Dim wsData As Worksheet                 ' 「データ登録」シートオブジェクト
    Dim protInfoAcq As SheetProtectionInfo  ' 「データ取得」シートの保護情報
    Dim protInfoData As SheetProtectionInfo ' 「データ登録」シートの保護情報
    Dim prevState As ApplicationState       ' Excel の実行前状態

    ' --- ステップ1：Excel 状態の保存と高速化設定 ---
    SaveAndSetApplicationState prevState
    On Error GoTo ErrorHandler

    ' --- ステップ2：ユーザーへの最終確認 ---
    If MsgBox( _
        "「" & GetSheetName(Sheet_DataAcquire) & "」「" & GetSheetName(Sheet_DataEntry) & "」の入力値をクリアします。" & vbCrLf & _
        "よろしいですか？", _
        vbYesNo + vbQuestion + vbDefaultButton2, "クリアの確認") = vbNo Then
        GoTo CleanUp
    End If

    ' --- ステップ3：ワークシートオブジェクトの取得 ---
    Set wsAcq = GetSheet(Sheet_DataAcquire)
    Set wsData = GetSheet(Sheet_DataEntry)

    ' --- ステップ4：シート保護の一時解除（失敗時は中断） ---
    If Not UnprotectSheetIfNeeded(wsAcq, protInfoAcq) Then GoTo CleanUp
    If Not UnprotectSheetIfNeeded(wsData, protInfoData) Then GoTo CleanUp

    ' --- ステップ5：指定範囲のデータクリア ---
    wsAcq.Range(CLEAR_RANGE_ACQ).ClearContents
    wsAcq.Range(DATE_CELL_GETOUT).ClearContents
    wsData.Range(DATA_ENTRY_DATE_CELL).ClearContents
    wsData.Range(CLEAR_RANGE_DATA).ClearContents
    wsData.Range(DATE_CELL_WORKTIME).ClearContents

    ' --- ステップ6：完了メッセージの表示 ---
    MsgBox "クリア完了", vbInformation, "完了"

CleanUp:
    ' --- 最終処理：シート保護と Excel 状態を元に戻す ---
    If Not wsAcq Is Nothing Then
        RestoreSheetProtection wsAcq, protInfoAcq
    End If
    If Not wsData Is Nothing Then
        RestoreSheetProtection wsData, protInfoData
    End If
    RestoreApplicationState prevState
    Exit Sub

ErrorHandler:
    ' --- エラー発生時の処理 ---
    MsgBox "クリア処理でエラーが発生しました: " & Err.Description, vbCritical, "エラー"
    Resume CleanUp
End Sub
