# Pythonスクリプトメモ

- `create_reference_docx.py` : `python-docx` を使い、和文フォント設定済みの Word テンプレート `reference.docx` を自動生成する。
- `insert_label.py` : `Rrecision_tilt_sensor.html` に欠けているセンサーラベル `<div class="sensor-label">JJ-07</div>` を一度だけ挿入する。
- `inspect_section.py` : 同 HTML の対象位置付近を `repr` で出力し、改行やタグ構造をデバッグ確認する。
- `show_tail.py` : `Rrecision_tilt_sensor.html` の末尾 200 文字を表示してファイル終端を素早く点検する。
- `refresh_rrecision_html.py` : 文字化け修正、`.sensor-label` のスタイル・要素追加、`<script>` 部分を最新の Web Audio／デバイス姿勢処理ロジックへ全置換する。
- `update_rrecision_html.py` : 上記 HTML の個別関数（`createTone` など）を書き換え、ステレオパンや連続モード対応を差分適用する。
- `presentations/FactoryTour/build_factory_tour_pptx.py` : BeautifulSoup で HTML を解析し、`python-pptx` で工場見学プレゼン `FactoryTour_Presentation.pptx` を生成する。
- `shortcut-mail-tool/apply_patch.py` : `ShortcutMailTool.bas` を丸ごと書き出し、CSV 読み込みやクリップボード操作を備えた VBA メニュー機能を自動投入する。
- `shortcut-mail-tool/check.py` / `shortcut-mail-tool/tmp_read.py` : 生成済みモジュールのエンコード断片を表示し、バイト列・テキスト状態を簡易診断する。
