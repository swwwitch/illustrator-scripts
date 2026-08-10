# AiFontConverter

[![Direct](https://img.shields.io/badge/Direct%20Link-AiFontConverter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/AiFontConverter.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiFontConverter.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

フォントの種別（文字セット・P・UD・N・NT・ウエイト）をまとめて変更するスクリプト。

- 対象は「選択オブジェクト / ドキュメント全体 / アクティブアートボード」から選択
- 変換設定は 3 カラム（左=文字セット／中央=N・NT／右=UD・P）。各項目は「現状維持 / なし / あり」で切り替え
- 文字セットは Std / Pro / Pr5 / Pr6（N は別軸でトグル）
- 新ゴ ⇄ 新ゴNT を NT 設定で切り替え
- プリセット Max（収録最多の N なし＋UD＋P）/ MaxN（収録最多の N あり込み＋UD＋P）
- G-OTF 学参書体（常改 / 学参 / K書体）を A-OTF に統合（チェックボックス）
- A1明朝など特殊シリーズは太さ等価で対応（A-OTF A1明朝 Std B ＝ A P-OTF A1明朝 StdN R）
- CID フォント・実行前確認は先頭のスイッチ（UI 非表示）で制御
- 段落スタイル・文字スタイル、ロック / 非表示オブジェクトも対象に含められる
- 実行前に変更内容（旧 → 新）を和文フォント名でプレビュー確認、カンバス上の位置順（上→下）に表示、未インストールフォントは事前に警告
- 同名ウエイトが見つからない場合は近いウエイトへ置換
- 適用は textRange 単位でまとめて処理（高速）
- 変換対象ファミリーはフォントデータベースから生成（FONT_FAMILIES）。モリサワ基幹書体に加え、筑紫書体シリーズ・UD書体・Fontworks 由来デザイン（セザンヌ / マティス / ロダン など）までカバー

### 参考

https://sttk3.com/blog/tips/illustrator/unify-character-set.html

### 紹介記事

https://note.com/dtp_tranist/n/n261c771b4b41

### 更新履歴

- v1.0.0 : 初版
- v1.0.1 : 確認ダイアログを和文フォント名で表示・カンバス上から順（上→下）に並べ替え・→ の位置をそろえる（変更前列を固定幅）・タイトルのバージョン表記を削除・左右マージンを調整
- v1.1.0 : 変換対象ファミリーを拡充（筑紫書体シリーズ・UD書体・Fontworks 由来デザインなど）
- v1.1.1 : 変換対象に AXIS（タイププロジェクト書体）を追加。専用処理で幅（Basic/Cond/Comp）と Joyo を保持し、N と Std⇄Pro のみ切り替え、Max/MaxN プリセット時は設定に関わらず ProN へ寄せる

### スクリプト情報

- バージョン: v1.1.1
