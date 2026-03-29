# LS入力システム 詳細設計書

最終更新: 2026-02-24 / 作成: JJ-07

## 1. システム構成
### 1.1 モジュール一覧
| モジュール / 要素 | 主なプロシージャ | 役割 |
| --- | --- | --- |
| ModAppConfig | GetSheetName, GetSheet | シート名や列番号、列追加ポリシーを列挙体で提供します。月次シートやデータ登録シートの列定義を共通化します。【F:LS入力/ModAppConfig.bas†L1-L56】 |
| ModCommonUtils | SaveAndSetApplicationState, RestoreApplicationState, DetermineTargetDate, ReportErrorToMonthlySheet | 画面更新停止・計算モード切替、シート保護の一時解除、対象日取得、月次シート `J3` へのメッセージ出力など共通処理を提供します。【F:LS入力/ModCommonUtils.bas†L1-L108】 |
| ModIntegratedMaster | GenerateIntegratedMasterSheet, EnsureIntegratedMasterReady, LoadIntegratedMasterWithCount | 統合マスタ `tblIntegratedMaster` の雛形生成、名前定義の検証、テーブル読込と警告管理を担います。【F:LS入力/ModIntegratedMaster.bas†L62-L239】 |
| ModInitialSetup | RunInitialSetup | 統合マスタを読み込み、作番・作業コード一覧を月次シートのヘッダーに反映するとともに、統合マスタへの遷移ダイアログを表示します。【F:LS入力/ModInitialSetup.bas†L1-L214】 |
| ModGetOutlookSch | ExecuteOutlookSchedule, GetOutlookSchedule | Outlook 予定の取得・分類・一覧出力を行います。`ExcludeKeywords` 名称範囲による除外や、統合マスタとの照合結果出力を実装しています。【F:LS入力/ModGetOutlookSch.bas†L1-L307】【F:LS入力/ModGetOutlookSch.bas†L434-L453】 |
| ModDataTransfer | TransferDataToMonthlySheet, CopyLatestDataToClipboard | 作番×作業コード単位で時間を集計し月次シートへ転記、重複ハイライト、結果ダイアログ、クリップボード出力を提供します。【F:LS入力/ModDataTransfer.bas†L229-L398】【F:LS入力/ModDataTransfer.bas†L909-L1099】 |
| ModDataClear | ClearInputData | 「データ取得」「データ登録」シートの入力値・推奨値列・日付セルを一括でクリアします。【F:LS入力/ModDataClear.bas†L36-L108】 |
| ModMonthlyMaintenance | ClearMonthlyDataAndRefreshCalendar, ClearAllMonthlyTransferArea, RefreshMonthlyCalendar | 月次シートの転記領域クリアとカレンダー再生成、保護解除・復元を行います。【F:LS入力/ModMonthlyMaintenance.bas†L1-L152】 |
| ModUIButtonSetup | InstallActionButtons, ApplyDataEntryValidation | ボタン設置と入力規則の再設定（統合マスタ連携のプルダウン）を自動化します。【F:LS入力/ModUIButtonSetup.bas†L1-L148】 |
| ダブルクリック削除r1 | Worksheet_BeforeDoubleClick | データ登録シート B 列ダブルクリックで該当行の入力値をクリアします。【F:LS入力/ダブルクリック削除r1.cls†L22-L146】 |

### 1.2 データフロー
```mermaid
flowchart LR
  Outlook -->|ModGetOutlookSch| DataAcquire[データ取得]
  IntegratedMaster[統合マスタ
(tblIntegratedMaster)] -->|照合| ModGetOutlookSch
  DataAcquire -->|手動入力| DataEntry[データ登録]
  DataEntry -->|ModDataTransfer| Monthly[月次データ]
  DataEntry -->|ModDataClear| DataEntry
  Monthly -->|ModMonthlyMaintenance| Monthly
  DataEntry -->|ダブルクリック削除r1| DataEntry
  IntegratedMaster -->|RunInitialSetup| Monthly
```

## 2. モジュール詳細
### 2.1 ModAppConfig
- `SheetName` 列挙体で「データ登録」「月次データ」「データ取得」を定義し、`GetSheetName` / `GetSheet` で参照します。
- `DataSheetColumn`（列 C=作番、D=作業コード、E=作業時間）、`MonthlySheetColumn`、`MonthlySheetRow` を列挙体で提供し、モジュール間で共通化しています。【F:LS入力/ModAppConfig.bas†L25-L54】
- 列追加ポリシー `ColumnAddPolicy` と既定値 `DEFAULT_COLUMN_ADD_POLICY`、ドライラン設定 `DEFAULT_DRY_RUN_MODE` を保持します。【F:LS入力/ModAppConfig.bas†L10-L24】

### 2.2 ModCommonUtils
- `ApplicationState` 構造体に画面更新・イベント・計算モードの状態を保持し、`SaveAndSetApplicationState` と `RestoreApplicationState` で高速化設定を往復させます。【F:LS入力/ModCommonUtils.bas†L1-L55】
- `UnprotectSheetIfNeeded` / `RestoreSheetProtection` でシート保護の解除・復元を実装し、保護されている場合は空パスワード試行→InputBox の順で処理します。【F:LS入力/ModCommonUtils.bas†L40-L78】
- `DetermineTargetDate` は `D4` 優先で日付を取得し、空の場合は `D3` を参照します。【F:LS入力/ModCommonUtils.bas†L80-L95】
- `ReportErrorToMonthlySheet` / `ClearErrorCellOnMonthlySheet` が月次シート `J3` の表示と折り返し設定を管理します。【F:LS入力/ModCommonUtils.bas†L96-L108】

### 2.3 ModIntegratedMaster
- `GenerateIntegratedMasterSheet` で統合マスタの雛形（列ヘッダー、ListObject、`IntegratedMaster_*` 名前定義）を一括生成します。【F:LS入力/ModIntegratedMaster.bas†L62-L137】
- `EnsureIntegratedMasterReady` は `tblIntegratedMaster` の存在と必要な名前定義を検証し、致命的な不足があればメッセージを返します。【F:LS入力/ModIntegratedMaster.bas†L154-L239】
- `LoadIntegratedMasterWithCount` は統合マスタを配列に読み込み、優先度昇順で返しつつ警告をキャプチャします。照合キーは `;` 区切りで分割し空白をトリムします。【F:LS入力/ModIntegratedMaster.bas†L240-L421】
- `MatchSubjectWithMaster` と `NormalizeRegistrationPolicy` は予定取得時に使用され、照合タイプ（完全一致・部分一致・前方一致・正規表現）と登録方針を解釈します。【F:LS入力/ModIntegratedMaster.bas†L323-L421】

### 2.4 ModInitialSetup
- 統合マスタの存在確認・読込後、作番と作業コードの辞書・組み合わせコレクションを生成し、月次シート 10 行目・11 行目にヘッダーを再配置します。【F:LS入力/ModInitialSetup.bas†L20-L142】
- 月次シートの保護を一時解除 (`UnprotectSheetIfNeeded`) し、実行後に復元します。処理結果（作番件数・作業コード件数・列数）をダイアログ表示し、統合マスタへ移動するオプションを提示します。【F:LS入力/ModInitialSetup.bas†L126-L214】

### 2.5 ModGetOutlookSch
- `GetOutlookSchedule` は日付入力検証→出力範囲クリア→ヘッダー設定→統合マスタ検証→除外キーワード読込→Outlook 接続→予定抽出→照合結果出力→取得日の転記を順に実行します。【F:LS入力/ModGetOutlookSch.bas†L34-L307】
- 出力列は「時間」「件名」「会議時間」「分類」「推奨作番」「区分」「推奨作業コード」「照合キー」「備考」で固定し、統合マスタと照合した情報を表示します。【F:LS入力/ModGetOutlookSch.bas†L89-L220】
- 取得完了後、メッセージボックスで件数や警告を通知し、統合マスタ関連の注意は `warnMsg` として蓄積して `J3` に追記します。【F:LS入力/ModGetOutlookSch.bas†L221-L307】

### 2.6 ModDataTransfer
- `TransferDataToMonthlySheet` は高速化設定・エラーセル初期化後、対象シート取得、シート保護解除、対象日の検出、月次行の確定を行います。対象日が見つからない場合は月次カレンダー再生成の確認ダイアログを表示します。【F:LS入力/ModDataTransfer.bas†L229-L341】
- `CollectTimeDataFromSheet` で列 C・D・E の入力値を走査し、作番・作業コード・分数（`ConvertToMinutesEx` で正規化）を収集します。空欄や 0 分は除外します。【F:LS入力/ModDataTransfer.bas†L495-L580】
- `AggregateTimeData` で `作業コード|作番` 単位に合算し、`BuildColumnMapping` で月次シートの列とキーの対応を辞書化します。列が不足する場合は列追加ポリシーに従い追加・キャンセルを判断します。【F:LS入力/ModDataTransfer.bas†L581-L838】
- `WriteTimeDataToCell` で `[hh]:mm` 書式の上書きを行い、既存値がある場合は黄色ハイライトと `J3` への重複ログ追記を実施します。【F:LS入力/ModDataTransfer.bas†L870-L918】
- 結果ダイアログには処理件数・重複件数・新規列追加数・エラー件数を表示し、成功時は `CopyDataToClipboard` でタブ区切り形式をクリップボードへ出力するボタンに誘導できます。【F:LS入力/ModDataTransfer.bas†L342-L398】【F:LS入力/ModDataTransfer.bas†L909-L1099】

### 2.7 ModDataClear
- 実行前に確認ダイアログを表示し、両シートの保護を一時解除します。クリア対象は「データ取得」`C8:K22` と `C4`、「データ登録」`D4`、推奨値列 `F8:F22`、勤務時間セル `E24` です。処理後に保護と Excel 状態を復元します。【F:LS入力/ModDataClear.bas†L36-L108】

### 2.8 ModMonthlyMaintenance
- 対象日（`DetermineTargetDate`）が取得できない場合は `J3` にエラーを追記し処理を中断します。【F:LS入力/ModMonthlyMaintenance.bas†L52-L90】
- 対象月の日数から末日行を算出し、`ClearAllMonthlyTransferArea` で 12 行目以降の値と塗りつぶしをクリアします。`RefreshMonthlyCalendar` で B 列に 1 日〜末日を再生成します。【F:LS入力/ModMonthlyMaintenance.bas†L90-L152】

### 2.9 ModUIButtonSetup
- 各シートにフォームボタンを配置または更新し、`OnAction` にラッパー (`Run_予定取得` など) を割り当てます。ラッパーはステータスバー表示とエラー処理を担当します。【F:LS入力/ModUIButtonSetup.bas†L1-L112】
- データ登録シートの入力規則として、作番・作業コードのリストを `IntegratedMaster_WorkNumber` と `IntegratedMaster_WorkCode` から取得し、時間列にはカスタム検証式を設定します。【F:LS入力/ModUIButtonSetup.bas†L113-L148】

### 2.10 ダブルクリック削除r1
- 「データ登録」シートの B 列をダブルクリックすると、対象行の C〜I 列をクリアします。イベント処理中はシート保護を一時解除し、復元時に再設定します。【F:LS入力/ダブルクリック削除r1.cls†L22-L146】

## 3. データ構造・名前定義
- ListObject `tblIntegratedMaster` の列順は `Priority` / `MatchType` / `MatchKeys` / `Class` / `Category` / `WorkCode` / `WorkNumber` / `RegistrationPolicy` / `Notes`。優先度は数値で昇順に評価されます。【F:LS入力/ModIntegratedMaster.bas†L101-L200】
- 名前定義 `IntegratedMaster_Table`、`IntegratedMaster_Priority`、`IntegratedMaster_Keys`、`IntegratedMaster_WorkCode`、`IntegratedMaster_WorkNumber` などを利用します。【F:LS入力/ModIntegratedMaster.bas†L115-L200】
- エラー表示セルは「月次データ」`J3` で統一し、`ReportErrorToMonthlySheet` を通じて改行追記と折り返し設定を行います。【F:LS入力/ModCommonUtils.bas†L96-L108】
- 除外キーワードは名前定義 `ExcludeKeywords` に設定した範囲を配列化して照合します。【F:LS入力/ModGetOutlookSch.bas†L145-L180】

## 4. 処理シーケンス概要
1. 利用者が「予定取得」を実行すると `ExecuteOutlookSchedule` が統合マスタを検証し、予定一覧を更新します。
2. 利用者が作番・作業コード・時間を入力し「転記実行」を押すと、`TransferDataToMonthlySheet` が日付行と列を確定して時間を集計・書き込みます。
3. 重複がある場合は `J3` にログが追記され、セルが黄色で強調されます。
4. 必要に応じて「月次更新」でカレンダーを再生成、「入力クリア」で入力欄を初期化します。
5. 統合マスタ更新後は「初期設定」で月次ヘッダーと入力規則を再同期します。

## 5. 変更履歴
- 2025-08-30: 初版作成（担当: JJ-07）
- 2026-02-24: 統合マスタ方式を導入。Outlook 照合・月次転記・ボタン配置処理を統合マスタ連携に対応させた（担当: JJ-07）。
