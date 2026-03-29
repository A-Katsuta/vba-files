# フェーズ1 詳細設計書 - VBA開発自動化MVP

## 1. システム構成

### 1.1 全体アーキテクチャ
```
┌─────────────────────────────────────────────┐
│  開発者                                      │
└────────────┬────────────────────────────────┘
             │ 実行（ダブルクリック）
             ▼
┌─────────────────────────────────────────────┐
│  build.bat                                  │
│  - PowerShell実行ポリシー回避               │
│  - Build-VbaProject.ps1呼び出し            │
└────────────┬────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────────┐
│  Build-VbaProject.ps1                      │
│  - テンプレート複製                         │
│  - VBAインポート処理                        │
│  - 参照設定追加                             │
│  - ファイル保存                             │
└─────────────────────────────────────────────┘
             │
    ┌────────┼────────┬─────────┐
    ▼        ▼        ▼         ▼
[template] [src/vba] [config]  [build]
```

### 1.2 ディレクトリ構造
```
project-root/
├── src/
│   └── vba/
│       ├── modules/         # 標準モジュール(.bas)
│       │   ├── Main.bas
│       │   └── Utils.bas
│       └── classes/         # クラスモジュール(.cls)
│           └── DataManager.cls
├── template/
│   └── BaseTemplate.xltm   # 基本テンプレート
├── tools/
│   └── Build-VbaProject.ps1 # ビルドスクリプト
├── build/                   # 出力ディレクトリ
│   └── (生成物)
├── config/
│   └── build-config.json   # ビルド設定（将来拡張用）
├── logs/                    # ログ出力
│   └── build-YYYYMMDD.log
├── .gitignore
├── build.bat               # 実行用バッチ
└── README.md              # ドキュメント
```

## 2. モジュール設計

### 2.1 build.bat
```batch
@echo off
setlocal enabledelayedexpansion

:: 設定
set "PS_SCRIPT=tools\Build-VbaProject.ps1"
set "LOG_DIR=logs"
set "LOG_FILE=%LOG_DIR%\build-%date:~0,4%%date:~5,2%%date:~8,2%.log"

:: ログディレクトリ作成
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

:: タイトル表示
echo ========================================
echo   VBA Project Build Tool v1.0
echo   %date% %time%
echo ========================================
echo.

:: PowerShellスクリプト実行
echo [INFO] ビルド処理を開始します...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" 2>&1 | tee -a "%LOG_FILE%"

:: 終了コード確認
if %ERRORLEVEL% equ 0 (
    echo.
    echo [SUCCESS] ビルドが正常に完了しました。
    echo 成果物: build\フォルダを確認してください。
) else (
    echo.
    echo [ERROR] ビルド中にエラーが発生しました。
    echo 詳細: %LOG_FILE% を確認してください。
)

echo.
pause
endlocal
```

### 2.2 Build-VbaProject.ps1

#### 2.2.1 メイン処理フロー
```powershell
# Build-VbaProject.ps1
[CmdletBinding()]
param(
    [string]$TemplatePath = ".\template\BaseTemplate.xltm",
    [string]$SourcePath = ".\src\vba",
    [string]$OutputPath = ".\build",
    [string]$OutputFileName = "Output_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsm"
)

# 初期設定
$ErrorActionPreference = "Stop"
$script:Excel = $null

# ログ関数
function Write-BuildLog {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $(
        switch($Level) {
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
    )
}

# クリーンアップ関数
function Clear-ExcelProcess {
    if ($script:Excel) {
        try {
            $script:Excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Excel) | Out-Null
            Remove-Variable -Name Excel -Scope Script
        } catch {
            Write-BuildLog "Excelプロセスの終了に失敗: $_" "WARNING"
        }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# エラーハンドラー設定
trap {
    Write-BuildLog "予期しないエラーが発生しました: $_" "ERROR"
    Clear-ExcelProcess
    exit 1
}

try {
    # 前提条件チェック
    if (-not (Test-Path $TemplatePath)) {
        throw "テンプレートファイルが見つかりません: $TemplatePath"
    }
    
    if (-not (Test-Path $SourcePath)) {
        throw "ソースフォルダが見つかりません: $SourcePath"
    }
    
    # 出力フォルダ作成
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-BuildLog "出力フォルダを作成しました: $OutputPath"
    }
    
    # Excel起動
    Write-BuildLog "Excelアプリケーションを起動しています..."
    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false
    
    # テンプレートからワークブック作成
    Write-BuildLog "テンプレートからワークブックを作成しています..."
    $Workbook = $script:Excel.Workbooks.Add($TemplatePath)
    
    # VBAプロジェクトへのアクセス
    $VBProject = $Workbook.VBProject
    
    # VBAモジュールインポート
    Import-VBAModules -VBProject $VBProject -SourcePath $SourcePath
    
    # 参照設定追加
    Add-References -VBProject $VBProject
    
    # ファイル保存
    $OutputFullPath = Join-Path (Resolve-Path $OutputPath) $OutputFileName
    Write-BuildLog "ファイルを保存しています: $OutputFullPath"
    $Workbook.SaveAs($OutputFullPath, 52) # 52 = xlOpenXMLWorkbookMacroEnabled
    
    # クローズ
    $Workbook.Close($false)
    
    Write-BuildLog "ビルドが正常に完了しました！" "SUCCESS"
    Write-BuildLog "出力ファイル: $OutputFullPath" "SUCCESS"
    
} finally {
    Clear-ExcelProcess
}
```

#### 2.2.2 VBAインポート関数
```powershell
function Import-VBAModules {
    param(
        [Parameter(Mandatory=$true)]$VBProject,
        [Parameter(Mandatory=$true)][string]$SourcePath
    )
    
    Write-BuildLog "VBAモジュールをインポートしています..."
    
    # 標準モジュール（.bas）のインポート
    $basFiles = Get-ChildItem -Path $SourcePath -Filter "*.bas" -Recurse
    foreach ($file in $basFiles) {
        try {
            $VBProject.VBComponents.Import($file.FullName)
            Write-BuildLog "  [OK] $($file.Name)"
        } catch {
            Write-BuildLog "  [FAIL] $($file.Name): $_" "ERROR"
            throw
        }
    }
    
    # クラスモジュール（.cls）のインポート
    $clsFiles = Get-ChildItem -Path $SourcePath -Filter "*.cls" -Recurse
    foreach ($file in $clsFiles) {
        try {
            $VBProject.VBComponents.Import($file.FullName)
            Write-BuildLog "  [OK] $($file.Name)"
        } catch {
            Write-BuildLog "  [FAIL] $($file.Name): $_" "ERROR"
            throw
        }
    }
    
    Write-BuildLog "インポート完了: 標準モジュール×$($basFiles.Count), クラス×$($clsFiles.Count)"
}
```

#### 2.2.3 参照設定追加関数
```powershell
function Add-References {
    param(
        [Parameter(Mandatory=$true)]$VBProject
    )
    
    Write-BuildLog "参照設定を追加しています..."
    
    # 必要な参照設定のGUID定義
    $references = @(
        @{
            Name = "Microsoft Scripting Runtime"
            GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"
            Major = 1
            Minor = 0
        },
        @{
            Name = "Microsoft ActiveX Data Objects 6.1 Library"
            GUID = "{B691E011-1797-432E-907A-4D8C69339129}"
            Major = 6
            Minor = 1
        }
    )
    
    foreach ($ref in $references) {
        try {
            # 既存チェック
            $exists = $false
            foreach ($existingRef in $VBProject.References) {
                if ($existingRef.GUID -eq $ref.GUID) {
                    $exists = $true
                    break
                }
            }
            
            if (-not $exists) {
                $VBProject.References.AddFromGuid($ref.GUID, $ref.Major, $ref.Minor)
                Write-BuildLog "  [追加] $($ref.Name)"
            } else {
                Write-BuildLog "  [スキップ] $($ref.Name) (既に追加済み)"
            }
        } catch {
            Write-BuildLog "  [警告] $($ref.Name): $_" "WARNING"
        }
    }
}
```

### 2.3 BaseTemplate.xltm の構成

#### 2.3.1 事前設定内容
- VBAプロジェクト名: "VBAProject"
- 基本シート構成:
  - Sheet1 (Main): メイン処理用
  - Sheet2 (Data): データ格納用
  - Sheet3 (Config): 設定値格納用

#### 2.3.2 プロジェクト設定
```vba
' ThisWorkbook モジュール
Option Explicit

Private Sub Workbook_Open()
    ' 初期化処理用のフック
    On Error Resume Next
    Application.Run "Initialize"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' クリーンアップ処理用のフック
    On Error Resume Next
    Application.Run "Cleanup"
    On Error GoTo 0
End Sub
```

## 3. エラー処理設計

### 3.1 エラー分類と対処

| エラー種別 | 発生箇所 | 対処方法 | 終了コード |
|-----------|---------|---------|-----------|
| テンプレート不在 | 初期チェック | 即座に終了、明確なエラーメッセージ | 1 |
| Excel起動失敗 | COM操作 | リトライ後、失敗時は終了 | 2 |
| VBAアクセス拒否 | VBProject操作 | セキュリティ設定の案内を表示 | 3 |
| インポート失敗 | モジュール読込 | 該当ファイル名を表示して終了 | 4 |
| 保存失敗 | ファイル出力 | 別名での保存を試行 | 5 |

### 3.2 ログ設計

#### ログレベル
- **ERROR**: 処理を継続できない重大なエラー
- **WARNING**: 処理は継続するが注意が必要な事象
- **INFO**: 正常な処理の進行状況
- **SUCCESS**: 処理の正常完了

#### ログフォーマット
```
[YYYY-MM-DD HH:MM:SS] [LEVEL] メッセージ
```

## 4. セキュリティ設計

### 4.1 必要な権限設定

#### Excel側の設定（初回のみ）
1. Excelのオプション → セキュリティセンター
2. 「マクロの設定」→「すべてのマクロを有効にする」または「デジタル署名されたマクロを除き、すべてのマクロを無効にする」
3. 「開発者向けのマクロ設定」→「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック

#### PowerShell実行ポリシー
- build.bat内で `-ExecutionPolicy Bypass` を指定することで回避

### 4.2 セキュリティリスクと対策

| リスク | 対策 |
|-------|------|
| 悪意のあるVBAコードの実行 | ソースコードのGit管理とレビュー体制 |
| 不正なテンプレートの使用 | テンプレートのハッシュ値チェック（将来実装） |
| 権限昇格 | 最小権限での実行、管理者権限不要 |

## 5. パフォーマンス設計

### 5.1 処理時間目標

| 処理 | 目標時間 | 備考 |
|------|---------|------|
| Excel起動 | 3秒以内 | |
| テンプレート読込 | 2秒以内 | |
| VBAインポート（10ファイル） | 5秒以内 | ファイル数に比例 |
| 参照設定追加 | 3秒以内 | |
| ファイル保存 | 2秒以内 | |
| **合計** | **15秒以内** | 標準的なプロジェクト |

### 5.2 最適化方針
- Excel.Visibleを$falseに設定して画面更新を抑制
- DisplayAlertsを$falseに設定して確認ダイアログを抑制
- 一括処理できる操作はまとめて実行

## 6. テスト設計

### 6.1 単体テスト

| テストID | テスト項目 | 確認内容 | 期待結果 |
|---------|-----------|---------|---------|
| UT-001 | テンプレート存在確認 | 存在しないパスを指定 | エラーメッセージ表示 |
| UT-002 | VBAファイル検索 | 空のソースフォルダ | 警告メッセージ表示 |
| UT-003 | Excel起動 | COM操作 | Excelプロセス起動 |
| UT-004 | モジュールインポート | .basファイル | 正常インポート |
| UT-005 | クラスインポート | .clsファイル | 正常インポート |

### 6.2 結合テスト

| テストID | テストシナリオ | 確認内容 | 期待結果 |
|---------|--------------|---------|---------|
| IT-001 | 正常系フロー | build.bat実行 | .xlsmファイル生成 |
| IT-002 | VBA実行確認 | 生成ファイルのマクロ実行 | 正常動作 |
| IT-003 | 参照設定確認 | VBEで参照設定を確認 | 指定ライブラリ追加済 |
| IT-004 | プロセス終了確認 | タスクマネージャー確認 | Excelプロセス残留なし |

## 7. 移行・導入計画

### 7.1 導入ステップ

| Step | 作業内容 | 所要時間 | 担当 |
|------|---------|---------|------|
| 1 | Git環境構築 | 30分 | 開発者 |
| 2 | フォルダ構造作成 | 15分 | 開発者 |
| 3 | Excelセキュリティ設定 | 10分 | 開発者 |
| 4 | テンプレート作成 | 1時間 | 開発者 |
| 5 | スクリプト配置 | 10分 | 開発者 |
| 6 | 動作確認 | 30分 | 開発者 |

### 7.2 既存プロジェクトからの移行

1. 既存.xlsmファイルからVBAコードをエクスポート（手動）
2. src/vbaフォルダに配置
3. テンプレートファイルの調整
4. build.bat実行による動作確認

## 8. 運用・保守設計

### 8.1 定期メンテナンス項目

| 項目 | 頻度 | 内容 |
|------|------|------|
| ログファイル削除 | 月次 | 30日以上前のログを削除 |
| buildフォルダ整理 | 週次 | 不要な生成物を削除 |
| テンプレート更新 | 必要時 | 基本構造の変更時 |

### 8.2 トラブルシューティング

| 症状 | 原因 | 対処法 |
|------|------|-------|
| "VBAプロジェクトへのアクセスが拒否されました" | セキュリティ設定未実施 | Excelのセキュリティセンターで設定 |
| "ファイルが見つかりません" | パス指定誤り | 相対パスを確認 |
| Excelプロセスが残る | 異常終了 | タスクマネージャーで手動終了 |
| 参照設定エラー | Office版数違い | GUIDを環境に合わせて修正 |

## 9. 拡張性の考慮

### 9.1 将来の拡張ポイント

- **設定ファイル対応**: build-config.jsonによる動的設定
- **複数テンプレート対応**: プロジェクトタイプ別テンプレート
- **増分ビルド**: 変更されたモジュールのみインポート
- **並列処理**: 複数ファイルの同時ビルド

### 9.2 フェーズ2への接続点

- テスト実行フック: RunAllTestsマクロの呼び出し準備
- テスト結果取得: 結果シートからのデータ読み取り
- ビルド結果判定: テスト失敗時の処理中断

---
*本詳細設計書は、フェーズ1（MVP）の技術的実装詳細を定義したものである。*