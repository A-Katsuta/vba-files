# VBA Files Repository

更新日: 2026-03-29

このリポジトリは、VBA マクロ、関連ドキュメント、補助ツール、旧資産を一元管理するための統合リポジトリです。  
2026-03-29 時点で `github-vba-lab` の主要資産を取り込み、現役資産と旧資産の置き場を分離しました。

## ディレクトリ構成

### `projects/`
業務で利用する現役の VBA プロジェクトを格納します。

- `ExcelMergeSystem`
- `LSEntry`
- `ShortcutMailTool`
- `TranscriptionSystem`
- `UsefulItems`

### `tooling/`
VBA 開発を支援する基盤、生成ツール、補助スクリプトを格納します。

- `VbaProjectScaffold`: `github-vba-lab/01_VBA開発自動化` から取り込んだ雛形生成基盤

### `docs/`
一般ドキュメント、プロンプト集、移行メモを格納します。

- `General`
- `Prompts`
- `Migration`

### `archive/`
すぐに現役採用しないが、参照価値のある旧資産を格納します。

- `github-vba-lab/documentation`
- `github-vba-lab/python-tools`
- `github-vba-lab/web-assets`
- `projects/LSEntry/archive/github-vba-lab-snapshot`
- `projects/ShortcutMailTool/archive/github-vba-lab-snapshot`

### `trash/`
削除ではなく退避したファイルを格納します。  
ビルド成果物、ロックファイル、一時的に不要と判断したものはここへ移します。

### `utils/`
バッチや軽量ユーティリティを格納します。

## 運用ルール

- 現役資産は `projects/` と `tooling/` に置く
- 旧版や比較用スナップショットは `archive/` に置く
- 削除候補や中間成果物は `trash/` に置く
- Excel のロックファイルやビルド生成物は Git 管理しない

## 補足

- 元の `C:\\github-vba-lab` は `C:\\github-vba-trash\\2026-03-29\\github-vba-lab` へ退避済みです
- GitHub 側での統合手順は `docs/Migration/GitHub統合手順.md` を参照してください
