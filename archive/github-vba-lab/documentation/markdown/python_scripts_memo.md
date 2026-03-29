# Pythonスクリプトメモ

- `python-tools/doc_templates/create_reference_docx.py` : `python-docx` で游ゴシック設定済みの Word テンプレート `reference.docx` を生成します。
- `python-tools/html_rrecision/insert_label.py` : `web-assets/rrecision/Rrecision_tilt_sensor.html` に欠けているセンサーラベル `<div class="sensor-label">JJ-07</div>` を一度だけ追記します。
- `python-tools/html_rrecision/inspect_section.py` : 同 HTML の対象位置付近を `repr` 出力し、改行やタグ構造を簡易確認します。
- `python-tools/html_rrecision/show_tail.py` : 上記 HTML の末尾 200 文字を表示し、ファイル終端の状態を素早く点検します。
- `python-tools/html_rrecision/refresh_rrecision_html.py` : 文字化け修正、`.sensor-label` のスタイル・要素追加、`<script>` ブロックの最新 Web Audio／デバイス姿勢処理ロジックへの置換を行います。
- `python-tools/html_rrecision/update_rrecision_html.py` : HTML 内の `createTone` など主要関数を差し替え、ステレオパンや連続モード対応を強化します。
- `python-tools/ppt_factory_tour/build_factory_tour_pptx.py` : BeautifulSoup で工場見学 HTML を解析し、`python-pptx` でプレゼン `FactoryTour_Presentation.pptx` を生成します。
- `excel-vba/shortcut-mail-tool/apply_patch.py` : `ShortcutMailTool.bas` を丸ごと書き出し、CSV 読み込みやクリップボード操作を備えた VBA メニュー機能を自動投入します。
- `excel-vba/shortcut-mail-tool/check.py` / `excel-vba/shortcut-mail-tool/tmp_read.py` : 生成済みモジュールのエンコード断片を表示し、バイト列／テキスト状態を簡易診断します。
