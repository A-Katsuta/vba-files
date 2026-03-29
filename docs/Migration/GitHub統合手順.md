# GitHub 統合手順

作成日: 2026-03-29 10:08 JST
作成者: Codex (GPT-5)

更新日: 2026-03-29

## 前提

- ローカル統合先は `C:\github-vba\vba-files`
- `github-vba-lab` の内容は、現役採用分を `tooling/` と `projects/` に、旧資産を `archive/` に取り込み済み
- 元のローカル `github-vba-lab` は `C:\github-vba-trash\2026-03-29\github-vba-lab` に退避済み

## 推奨方針

GitHub 上では `vba-files` を統合先の正本リポジトリにします。  
`github-vba-lab` は最終的に README だけの退避リポジトリ、または Archive / Private のどれでも構いません。

## 手順

1. `vba-files` 側で統合作業用ブランチを push する
2. GitHub 上で Pull Request を作成し、統合内容をレビューする
3. PR には `github-vba-lab` のどの資産を `tooling` `archive` `projects/*/archive` に寄せたかを明記する
4. Merge 後、必要であれば `github-vba-lab` は README のみ残す構成へ簡略化する
5. `github-vba-lab` は Archive するか、README のみを残して更新停止にする

## 履歴を残したい場合

`github-vba-lab` の Git 履歴も `vba-files` に取り込みたい場合は、別ブランチで履歴統合作業を行います。  
安全なのは `git subtree` または履歴変換ツールを使って `archive/github-vba-lab-history` のようなサブディレクトリに寄せる方法です。

### 例

```powershell
git remote add github-vba-lab https://github.com/A-Katsuta/github-vba-lab.git
git fetch github-vba-lab
git subtree add --prefix=archive/github-vba-lab-history github-vba-lab main
```

この方法は履歴を残せますが、既存構成と今回の整理後構成が二重化しやすいため、通常は「ファイルを整理済みの現在構成を正本にし、旧リポジトリは Archive」にする方が管理しやすいです。

## GitHub 上で確認すべきこと

- `vba-files` の `main` に保護ルールを設定する
- `github-vba-lab` の README に移行案内を入れる
- 必要なら `github-vba-lab` の Topics や説明文に「migrated」を明記する
- 旧 Issue や PR を残す必要がある場合は、Archive 前に URL を控える

## ローカルで次にやること

- `git status` で今回の差分を確認する
- 問題なければ `vba-files` でコミットする
- PR 作成後に `github-vba-lab` 側の案内更新へ進む
