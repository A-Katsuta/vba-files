# Repository Guidelines

## プロジェクト構成
- LS入力/: Excel マクロ本体（.bas, .cls, .xlsm）。例: ModDataTransfer.bas, ModGetOutlookSch.bas。
- LS入力/tests/: 手動テスト仕様。例: 	est_patterns.md。
- excel-vba/shortcut-mail-tool/: ���[����^���c�[���BShortcutMailTool.bas, sample_data.csv�B
- JJ-27: 傾斜角センサー用ラベル仕様メモ
- documentation/markdown/docs/: 仕様・手順。例: documentation/markdown/docs/LS-input-r1.md。
- その他: 設計文書・スクリーンショット・サンプル。

## 開発・実行・テスト
- 実行: LS入力/LS入力r1.xlsm を開き、[開発]→[マクロ]→RunAllTests または個別 Test_* を実行。
- 参照設定: Outlook 連携は「Microsoft Outlook xx.x Object Library」。クリップボード操作は「Microsoft Forms 2.0 Object Library」。
- 手動テスト: LS入力/tests/test_patterns.md に準拠。変更時はケースを更新。
- Git 例: git status（変更確認）, git diff（差分確認）。

## コーディング規約・命名
- Option Explicit 必須。インデントは 4 スペース。
- モジュール: ModXxx.bas（例: ModMonthlyMaintenance.bas）、クラスは *.cls。
- 手続き: PascalCase（例: TransferDataToMonthlySheet）。テストは Test_*、一括は RunAllTests。
- エラー処理: On Error GoTo ErrHandler + Debug.Print で記録し、必要時に MsgBox 通知。
- 名前付き範囲や設定は ModAppConfig に集約し、セル参照の直書きを避ける。

## ビルド・整形
- ビルド不要。.bas/.cls がソース・オブ・トゥルース。VBE からエクスポート/インポートで管理。
- 文字コードは UTF-8 を推奨。CSV も UTF-8（例: excel-vba/shortcut-mail-tool/sample_data.csv）。

## テスト指針
- 新規機能には対応する Test_* を追加し、RunAllTests の完走を確認。
- 正常/異常/境界のケースを 	ests/test_patterns.md に追記。イミディエイトウィンドウでログ確認。

## コミット・PR
- コミット: 要約 1 行 + 必要なら本文。ix: ... / eat: ... 接頭辞を推奨（既存履歴に準拠）。
- PR: 目的・影響範囲・手動テスト結果・関連 Issue・必要に応じスクリーンショット（Excel 画面）を記載。
- .xlsm 更新時は、対応する .bas/.cls の差分も含める。

## セキュリティ・運用
- 機密情報はコード/CSVに含めない。マクロの署名/信頼設定を確認。
- 依存環境: Windows + Excel 2016 以降、Outlook（Outlook 連携機能使用時）。

