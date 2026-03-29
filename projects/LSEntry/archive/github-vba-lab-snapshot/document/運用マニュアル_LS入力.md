# LS入力 運用マニュアル

## 1. 本書の目的
本書は Excel ブック「LS入力r1.xlsm」に収録されたマクロの操作手順と運用上の注意点を、日々の入力担当者と管理者向けに整理したものです。Outlook の予定から作業実績を取り込み、月次シートへ転記して活用する一連の流れを解説します。

## 2. シート構成と主要機能
| シート名 | 役割 | 主な内容 |
| --- | --- | --- |
| データ取得 | Outlook 予定の取り込みと分類候補の確認 | `C3` で対象日を指定し、`C7:K` に時間・件名・分類・推奨作番・区分・推奨作業コード・照合キー・備考を出力します。取得日 (`C4`) はデータ登録シートにコピーされます。【F:LS入力/ModGetOutlookSch.bas†L34-L307】 |
| データ登録 | 月次転記前の入力・確認 | `D4` が転記対象日（空欄時は `D3` を使用）。列 C=作番、列 D=作業コード、列 E=作業時間を入力すると、`TransferDataToMonthlySheet` が日付×作業コード×作番の組み合わせで集計します。【F:LS入力/ModCommonUtils.bas†L80-L95】【F:LS入力/ModDataTransfer.bas†L505-L518】 |
| 月次データ | 月次集計とエラー表示 | B 列にカレンダーを生成し、10 行目=作番ヘッダー、11 行目=作業コードヘッダー、12 行目以降に時間を転記します。エラーセル `J3` にメッセージが追記されます。【F:LS入力/ModInitialSetup.bas†L70-L126】【F:LS入力/ModDataTransfer.bas†L870-L918】【F:LS入力/ModCommonUtils.bas†L40-L78】 |
| 統合マスタ | 件名と推奨値のテーブル管理 | `tblIntegratedMaster` と `IntegratedMaster_*` の名前定義を保持し、予定取得時の照合および初期設定時のヘッダー更新に利用されます。【F:LS入力/ModIntegratedMaster.bas†L101-L200】【F:LS入力/ModInitialSetup.bas†L1-L214】 |
| 要望・提案 | 改善メモ | マクロからの参照はありません。 |

## 3. ボタンとマクロの対応
| シート | ボタン | 実行マクロ | 主な処理 |
| --- | --- | --- | --- |
| データ取得 | 予定取得 | `ExecuteOutlookSchedule` | 指定日の Outlook 予定を取得し、統合マスタと照合して一覧を出力します。除外キーワード `ExcludeKeywords` に一致した予定はスキップされます。【F:LS入力/ModGetOutlookSch.bas†L122-L284】【F:LS入力/ModGetOutlookSch.bas†L434-L453】 |
| データ登録 | 転記実行 | `TransferDataToMonthlySheet` | 作番・作業コード・時間を集計し、月次シートへ書き込みます。重複セルはハイライトされ `J3` に記録されます。【F:LS入力/ModDataTransfer.bas†L229-L398】【F:LS入力/ModDataTransfer.bas†L870-L918】 |
| データ登録 | 入力クリア | `ClearInputData` | 「データ取得」「データ登録」シートの入力値・推奨値列・日付セルを一括で初期化します。【F:LS入力/ModDataClear.bas†L36-L108】 |
| 月次データ | 月次更新 | `ClearMonthlyDataAndRefreshCalendar` | 月次転記領域をクリアし、対象月のカレンダーを再生成します。【F:LS入力/ModMonthlyMaintenance.bas†L1-L152】 |
| 月次データ | コピー | `CopyLatestDataToClipboard` | データ登録シートの最新入力をタブ区切りでクリップボードにコピーします。【F:LS入力/ModDataTransfer.bas†L909-L1099】 |

ボタンは `ModUIButtonSetup` の `InstallActionButtons` を実行すると再配置および入力規則の再設定が行われます。【F:LS入力/ModUIButtonSetup.bas†L1-L148】

## 4. 動作環境と事前準備
1. **Excel/Outlook**: Windows 版 Excel 2016 以降および Outlook (MAPI) を使用します。参照設定「Microsoft Outlook XX.X Object Library」を有効化してください。【F:LS入力/ModGetOutlookSch.bas†L5-L17】
2. **マクロ有効化**: ブックを開いた際に表示される警告で「コンテンツの有効化」を選択します。
3. **統合マスタの整備**:
   - `GenerateIntegratedMasterSheet` で雛形を作成し、必要な列に値を登録します。
   - 名前定義 `IntegratedMaster_*` がテーブル列を指していることを確認します。【F:LS入力/ModIntegratedMaster.bas†L62-L200】
4. **初期設定の実行**: 統合マスタの準備後に `RunInitialSetup` を実行し、月次ヘッダーとデータ登録シートのプルダウンを更新します。【F:LS入力/ModInitialSetup.bas†L1-L214】
5. **シート保護**: 保護を有効にする場合はパスワードを記録のうえ、マクロが解除・復元できることを確認します。【F:LS入力/ModCommonUtils.bas†L40-L78】

## 5. 日常業務の操作手順
1. **対象日の入力**: 「データ取得」シート `C3` に対象日を入力します。空欄のまま実行すると日付入力を促すメッセージが表示されます。【F:LS入力/ModGetOutlookSch.bas†L52-L86】
2. **予定の取り込み**: 「予定取得」ボタンを押し、一覧に表示された分類・推奨値・備考を確認します。処理完了後、取得日が自動的に「データ登録」シート `D4` へコピーされます。【F:LS入力/ModGetOutlookSch.bas†L286-L307】
3. **入力・修正**: 推奨値を参考に作番（列 C）・作業コード（列 D）・作業時間（列 E）を入力します。時間は `0930` などの `HHMM` 形式または `[hh]:mm` 形式に対応します。【F:LS入力/ModDataTransfer.bas†L505-L518】
4. **月次転記**: 「転記実行」ボタンを押すと、対象日のデータが集計され月次シートに書き込まれます。ダイアログで処理件数・重複件数・新規列追加数・エラー数を確認してください。【F:LS入力/ModDataTransfer.bas†L229-L398】
5. **結果の共有（必要時）**: 「月次データ」シートの「コピー」ボタンで最新の入力一覧をクリップボードへコピーし、メールやチャットに貼り付けられます。【F:LS入力/ModDataTransfer.bas†L909-L1099】
6. **入力のリセット**: 翌日の準備として「入力クリア」ボタンで「データ取得」「データ登録」シートの入力欄を初期化します。【F:LS入力/ModDataClear.bas†L36-L108】
7. **月次カレンダーの更新**: 月初や統合マスタ更新後にヘッダーが変わった場合は「月次更新」ボタンを押してカレンダーを再生成します。対象月の日付が B 列に再配置され、転記領域がクリアされます。【F:LS入力/ModMonthlyMaintenance.bas†L40-L152】

## 6. 統合マスタと名前定義の運用
- `tblIntegratedMaster` の列は「優先度」「照合タイプ」「照合キー」「分類」「区分」「標準作業コード」「標準作番」「登録方針」「備考」を順に配置します。優先度が小さい行が優先されます。【F:LS入力/ModIntegratedMaster.bas†L101-L200】
- `IntegratedMaster_WorkNumber` と `IntegratedMaster_WorkCode` はデータ登録シートの入力規則で使用されるため、列幅や末尾の空行に注意してください。【F:LS入力/ModUIButtonSetup.bas†L107-L148】
- マスタを更新したら `RunInitialSetup` を再実行し、月次ヘッダーとプルダウンを同期します。【F:LS入力/ModInitialSetup.bas†L70-L214】
- 除外キーワードを追加したい場合は名前定義 `ExcludeKeywords` の範囲に語句を追記します。【F:LS入力/ModGetOutlookSch.bas†L145-L150】

## 7. トラブルシューティング
| 症状 | 想定原因 | 対処方法 |
| --- | --- | --- |
| Outlook 予定が取得できない | Outlook ライブラリ参照が外れている、または Outlook が未起動 | 参照設定を確認し、Outlook を起動したうえで再実行します。【F:LS入力/ModGetOutlookSch.bas†L10-L82】 |
| 統合マスタの警告が表示される | `IntegratedMaster_*` 名前定義の不足やテーブルが空 | 統合マスタを開いて名前定義とデータを確認し、必要に応じて `RunInitialSetup` を再実行します。【F:LS入力/ModIntegratedMaster.bas†L154-L239】【F:LS入力/ModInitialSetup.bas†L20-L126】 |
| 転記時に対象日が見つからない | 月次カレンダーの更新漏れ | ダイアログの案内に従い「はい」を選択すると自動でカレンダーを再生成します。【F:LS入力/ModDataTransfer.bas†L297-L341】【F:LS入力/ModMonthlyMaintenance.bas†L75-L136】 |
| 重複セルが黄色になる | 同じ日付・作業コード・作番のセルに既存値がある | `J3` のメッセージと該当セルを確認し、必要に応じて値を調整後に再転記します。【F:LS入力/ModDataTransfer.bas†L870-L918】 |
| 入力クリアが実行できない | シート保護解除をキャンセルした | シート保護のパスワードを入力して解除し、再度ボタンを押します。【F:LS入力/ModDataClear.bas†L60-L108】 |

## 8. 運用上の注意
- マクロ実行中は画面更新や計算モードが一時的に変更されます。処理完了まで操作を控えてください。【F:LS入力/ModCommonUtils.bas†L1-L108】【F:LS入力/ModDataTransfer.bas†L229-L276】
- シート保護はマクロが自動で解除・復元しますが、パスワード管理はチームの運用ルールに従ってください。【F:LS入力/ModCommonUtils.bas†L40-L78】
- 統合マスタを更新した場合は、月次ヘッダーと入力規則を最新化するために `RunInitialSetup` を必ず実行します。【F:LS入力/ModInitialSetup.bas†L70-L214】
- クリップボードコピー機能は Forms.DataObject もしくは WinAPI を利用します。社内端末でセキュリティポリシーにより制限されている場合は IT 管理者へ相談してください。【F:LS入力/ModDataTransfer.bas†L957-L1099】

## 9. 連絡先
運用に関する問い合わせや統合マスタの更新依頼は、ツール管理者またはチーム内で定めた窓口まで連絡してください。バージョン更新時は本マニュアルを最新化し、利用者へ周知します。
