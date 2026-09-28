# 各種環境設定をダイアログから変更

[![Direct](https://img.shields.io/badge/Direct%20Link-PreferenceManager--unit.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PreferenceManager-unit.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PreferenceManager-unit.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- Illustrator の各種環境設定をダイアログボックスから変更可能にします。
- 単位、文字設定、変形／整列設定などをワンパネルで調整できます。
- Allows changing various Illustrator preferences via a dialog box.
- Units, text settings, transform/align settings can be adjusted in one panel.

### 主な機能

- 単位（一般・線・文字・東アジア言語）の設定 / Set units (general, stroke, text, East Asian)
- キー入力の移動量と角丸ツールの既定半径 / Keyboard increment and the default corner radius
- 文字サイズ・行送り・ベースラインシフトの増減量 / Step sizes for type size, leading and baseline shift
- 最近使用したフォント数とフォント表記切替 / Recent fonts count and font name localization
- 変形と整列の各種オプション / Various transform and align options
- 字形の境界に整列の設定 / Align to glyph bounds setting
- モード切替（プリント pt／プリント Q／オンスクリーン） / Mode switch (Print pt / Print Q / Onscreen)

### 参考

https://judicious-night-bca.notion.site/app-getIntegerPreference-e4088e6caef64b3ba5b801d86fba7877

### 更新履歴

- v1.0 (20250804): 初期バージョン / Initial version
- v1.1 (20250804): ダイアログを2カラムに改修、単位とフォント設定を追加 / Dialog changed to two columns; added units and font settings
- v1.2 (20250804): 角の拡大のロジックを修正 / Fixed logic for corner scaling
- v1.2.2 (2026-09-19): 機能が重複していた `PreferenceManager.jsx` を統合。全項目にツールチップを追加 / Merged the overlapping `PreferenceManager.jsx`; added tooltips to every item
- v1.2.3 (2026-09-27): 単位のドロップダウンにツールチップを追加し、「プリント（Q）」のツールチップを東アジア言語=H に訂正。英語UIの項目名のコロンを半角に統一（「Keyboard Increment::」の重複も修正）。コードを整理 / Added tooltips to the unit dropdowns and corrected the "Print (Q)" tooltip to East Asian=H; English field labels now use a single half-width colon (fixed the doubled "Keyboard Increment::"); code cleanup
- v1.3.0 (2026-09-27): 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ） / Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.3.1 (2026-09-28): ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた / The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%

### スクリプト情報

- バージョン: v1.3.1
