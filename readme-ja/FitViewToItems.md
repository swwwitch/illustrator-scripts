# オブジェクトが指定の割合で収まるよう表示を合わせる（再利用テンプレート）

[![Direct](https://img.shields.io/badge/Direct%20Link-FitViewToItems.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/_templates/FitViewToItems.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitViewToItems.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

オブジェクトの中心へ表示を移し、ウィンドウに対して指定した割合（％）で収まるよう表示倍率を合わせる再利用テンプレートです。

ダイアログに「□画面にフィット［65］%」の行を追加でき、キャンセル時に元の表示へ戻すための関数もそろっています。

### 主な機能

- 対象の外接範囲（効果を含まない）の中心に表示を移し、指定の割合で収まるよう拡大・縮小します
- 表示倍率はIllustratorが受け付ける範囲（3.125〜6400%）に収めます
- 日本語／英語のラベルとツールチップを内蔵した「□画面にフィット［％］」の行を追加できます
- 表示位置と倍率を控えて戻す `captureView` / `restoreView`

### 使い方

1. `var FitViewToItems = (function () { ... })();` のブロックまるごとを、対象スクリプトのIIFE内へコピーします。
2. ダイアログを開く前に表示を控え、行を追加します。

        var initialView = FitViewToItems.captureView(doc);
        var fitViewControls = FitViewToItems.addControls(optionPanel, { value: false, percent: 65 });

3. プレビューを作り直したところで呼びます。

        if (fitViewControls.checkbox.value) {
            FitViewToItems.fit([previewItem], { doc: doc, fillRatio: fitViewControls.getFillRatio() });
            app.redraw();
        }

4. チェックの切り替えで入力欄の有効状態をそろえ、キャンセル時は表示を戻します。

        fitViewControls.checkbox.onClick = function () { fitViewControls.updateEnabled(); refitView(); };
        /* キャンセル時 */
        FitViewToItems.restoreView(initialView, doc);

### 注意点

- 表示倍率の変更はUndo履歴に残らないため、キャンセル時の `restoreView` は呼び出し側で行います。
- 設定を変えるたびに呼ぶと倍率が頻繁に動いて落ち着かないため、大きさが変わる操作（幅の変更など）とダイアログを開いたときだけ呼ぶのがおすすめです。
- すでに見えているときは動かしたくない場合は [KeepInView](KeepInView.md) を使います。
- `#include` は使わず、各スクリプトを1ファイルで完結させる方針のためコピーして使います。
- 使用例: `jsx/shape/SmartShapeMaker.jsx`（「画面にフィット」）

### 更新履歴

- v1.0.0: 初版（テンプレート。SmartShapeMaker v2.3.0 から抜き出し）
