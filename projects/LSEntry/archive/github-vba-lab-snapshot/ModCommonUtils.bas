Option Explicit

'=================================================================================
' 【モジュール概要】
' 共通ユーティリティ: アプリケーション状態・シート保護・エラー表示
'=================================================================================

'=================================================================================
' 【構造体定義】
' Excel の状態とシート保護情報を一時保持するための Type
'=================================================================================

Public Type ApplicationState
    ScreenUpdating As Boolean    ' 画面更新の状態
    EnableEvents   As Boolean    ' イベント発生の状態
    Calculation    As Long       ' 計算モード（XlCalculation 列挙体の値）
End Type

Public Type SheetProtectionInfo
    IsProtected As Boolean       ' 元の保護状態（True=保護中）
    Password    As String        ' 保護解除に使用したパスワード
End Type

'=================================================================================
' 【アプリケーション状態管理】
' ScreenUpdating / EnableEvents / Calculation の保存と復元
'=================================================================================

'=================================================================================
' 【機能名】Excel 実行状態の保存と高速化設定
' Excel の現在状態を保存し、画面更新・イベントを停止して高速化する
' 引数説明: prevState - 復元用に状態を格納する構造体（呼び出し側で保持）
'=================================================================================
Public Sub SaveAndSetApplicationState(ByRef prevState As ApplicationState)
    ' --- 変数宣言 ---
    With prevState
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .Calculation = Application.Calculation
    End With

    ' --- ステップ1：画面更新とイベントを停止し計算モードを手動へ変更 ---
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub

'=================================================================================
' 【機能名】Excel 実行状態の復元
' 保存しておいた状態を Application に戻し、呼び出し前と同じ環境にする
' 引数説明: prevState - SaveAndSetApplicationState で取得した状態
'=================================================================================
Public Sub RestoreApplicationState(ByRef prevState As ApplicationState)
    With Application
        .Calculation = prevState.Calculation
        .EnableEvents = prevState.EnableEvents
        .ScreenUpdating = prevState.ScreenUpdating
    End With
End Sub

'=================================================================================
' 【シート保護管理】
' 保護解除の自動化と復元処理
'=================================================================================

'=================================================================================
' 【機能名】必要に応じてシート保護を解除
' シートが保護済みならパスワード入力を促し、解除結果を Boolean で返す
' 引数説明: ws - 対象シート / protInfo - 復元用の保護情報を格納
' 戻り値説明: True=処理継続可能、False=解除できず
'=================================================================================
Public Function UnprotectSheetIfNeeded(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo) As Boolean
    protInfo.IsProtected = ws.ProtectContents
    protInfo.Password = ""

    ' --- ステップ1：既に保護されていない場合は即時成功 ---
    If Not protInfo.IsProtected Then
        UnprotectSheetIfNeeded = True
        Exit Function
    End If

    ' --- ステップ2：空パスワードで解除できるか確認 ---
    On Error Resume Next
    ws.Unprotect ""
    If Err.Number = 0 Then
        UnprotectSheetIfNeeded = True
        protInfo.Password = ""
        On Error GoTo 0
        Exit Function
    End If

    ' --- ステップ3：ユーザーにパスワード入力を要求 ---
    Err.Clear
    protInfo.Password = InputBox("シート【" & ws.Name & "】の保護パスワードを入力してください。", "保護解除")
    If protInfo.Password = "" Then
        UnprotectSheetIfNeeded = False
        On Error GoTo 0
        Exit Function
    End If

    ' --- ステップ4：入力パスワードで解除を試行 ---
    ws.Unprotect protInfo.Password
    UnprotectSheetIfNeeded = (Err.Number = 0)
    On Error GoTo 0
End Function

'=================================================================================
' 【機能名】シート保護の復元
' 保護解除時に取得した情報を用いて元の状態に戻す
' 引数説明: ws - 対象シート / protInfo - UnprotectSheetIfNeeded で得た情報
'=================================================================================
Public Sub RestoreSheetProtection(ByRef ws As Worksheet, ByRef protInfo As SheetProtectionInfo)
    If ws Is Nothing Then Exit Sub
    If protInfo.IsProtected Then
        On Error Resume Next
        ws.Protect Password:=protInfo.Password, UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

'=================================================================================
' 【日付およびエラー表示ユーティリティ】
' 月次シート連携における共通処理
'=================================================================================

'=================================================================================
' 【機能名】対象日（D4 優先→D3）の取得
' データ登録シートから優先順位付きで日付セルを確認し、取得に成功したか返す
' 引数説明: wsData - データ登録シート / targetDate - 取得した日付（ByRef）
' 戻り値説明: True=日付取得成功、False=日付未設定
'=================================================================================
Public Function DetermineTargetDate(ByRef wsData As Worksheet, ByRef targetDate As Date) As Boolean
    Const DATE_CELL_NORMAL As String = "D3"

    DetermineTargetDate = False

    ' --- ステップ1：優先セル（D4）を確認 ---
    If IsDate(wsData.Range(DATA_ENTRY_DATE_CELL).Value) Then
        targetDate = CDate(wsData.Range(DATA_ENTRY_DATE_CELL).Value)
        DetermineTargetDate = True
    ' --- ステップ2：通常セル（D3）で代替取得 ---
    ElseIf IsDate(wsData.Range(DATE_CELL_NORMAL).Value) Then
        targetDate = CDate(wsData.Range(DATE_CELL_NORMAL).Value)
        DetermineTargetDate = True
    End If
End Function

'=================================================================================
' 【機能名】月次シートへのエラーメッセージ出力
' 月次シートのエラー表示セルにメッセージを設定し、必要に応じて追記する
' 引数説明: message - 表示する内容 / append - True で既存内容に追記
'=================================================================================
Public Sub ReportErrorToMonthlySheet(ByVal message As String, Optional ByVal append As Boolean = False)
    ' --- 変数宣言 ---
    Dim wsMonthly As Worksheet      ' 月次シートオブジェクト

    ' --- ステップ1：月次シート取得（存在しなければ終了） ---
    On Error Resume Next
    Set wsMonthly = GetSheet(Sheet_Monthly)
    If wsMonthly Is Nothing Then Exit Sub

    ' --- ステップ2：エラーセルへ設定（追記指定時は改行して追加） ---
    With wsMonthly.Range(ERR_CELL_ADDR)
        If append And Len(.Value) > 0 Then
            .Value = CStr(.Value) & vbLf & message
        Else
            .Value = message
        End If
        .WrapText = True
    End With
    On Error GoTo 0
End Sub

'=================================================================================
' 【機能名】月次シートのエラーセル初期化
' エラー表示セルの値をクリアし、後続処理で再利用できる状態にする
'=================================================================================
Public Sub ClearErrorCellOnMonthlySheet()
    ' --- 変数宣言 ---
    Dim wsMonthly As Worksheet      ' 月次シートオブジェクト

    On Error Resume Next
    Set wsMonthly = GetSheet(Sheet_Monthly)
    If wsMonthly Is Nothing Then Exit Sub

    wsMonthly.Range(ERR_CELL_ADDR).ClearContents
    On Error GoTo 0
End Sub

'=================================================================================
' 【機能名】エラー詳細メッセージ生成
' エラー番号と説明を整形し、ユーザー向けの表示文を作成する
' 引数説明: errNo - 発生したエラー番号 / errDesc - エラー内容
' 戻り値説明: 整形済みの文字列
'=================================================================================
Public Function GetErrorDetails(ByVal errNo As Long, ByVal errDesc As String) As String
    ' --- 変数宣言 ---
    Dim displayNo As Long           ' 表示用に調整したエラー番号
    Dim msg As String               ' 出力用メッセージ

    ' --- ステップ1：VBA 固有エラー番号を調整 ---
    If errNo >= vbObjectError And errNo < 0 Then
        displayNo = errNo - vbObjectError
    Else
        displayNo = errNo
    End If

    ' --- ステップ2：番号と内容を連結 ---
    msg = "エラー番号: " & CStr(displayNo)
    If Len(errDesc) > 0 Then
        msg = msg & vbCrLf & "内容: " & errDesc
    End If

    ' --- 結果返却 ---
    GetErrorDetails = msg
End Function
