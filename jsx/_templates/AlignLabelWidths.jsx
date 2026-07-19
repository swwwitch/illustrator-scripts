#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

AlignLabelWidths.jsx（テンプレート）

### 概要：

- ScriptUI の複数ラベル（statictext）の幅を最長のものへ自動で揃え、右揃えにするための再利用テンプレート
- 「ラベル：値」を縦に並べる情報パネル／設定パネルで、コロン位置と値の開始位置を揃える用途
- ロケールや文言の長さに依存せず、実際の描画幅を測って揃える（`characters` の固定値より堅牢）

### 使い方：

1. `alignLabelWidths()`（または簡易版 `setLabelsFixedWidth()`）と、必要なら `addLabelValueRow()` を対象スクリプトのトップレベルにコピー
2. 各行を作りながらラベルを配列に集める
3. 集めたラベル配列を揃え関数に渡す
   - 方式A（推奨・実測して自動）：`alignLabelWidths(labels, 'right')`
   - 方式B（簡易・固定文字数。PathCleanupTool と同方式）：`setLabelsFixedWidth(labels, 13, 'right')`

### 更新履歴：

- v1.0.0 : 初版（テンプレート）

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AlignLabelWidths";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// 再利用パーツ / Reusable parts
//   ↓↓↓ ここから下の2関数を対象スクリプトへコピーして使う ↓↓↓
// =========================================

/**
 * 複数のラベル（statictext）の幅を最長のものへ揃え、指定方向に揃えます。
 * 実際の描画幅（preferredSize.width）を測るため、文言長やロケールに依存しません。
 * @param {Array<StaticText>} labels - 幅を揃える statictext の配列。
 * @param {string} [justify] - 揃え方向 "left" | "center" | "right"（省略時は "right"）。
 * @returns {number} 揃えた後の共通ラベル幅（px）。
 */
function alignLabelWidths(labels, justify) {
    if (!labels || !labels.length) return 0;
    justify = justify || 'right';

    /* 1パス目：各ラベルの自然幅を測り、最大値を求める / Pass 1: find the widest natural width */
    var maxWidth = 0;
    for (var i = 0; i < labels.length; i++) {
        var naturalWidth = labels[i].preferredSize.width;
        /* 取得できない環境向けのフォールバック / Fallback when preferredSize is unavailable */
        if (!naturalWidth || naturalWidth < 0) naturalWidth = labels[i].size.width;
        if (naturalWidth > maxWidth) maxWidth = naturalWidth;
    }

    /* 2パス目：全ラベルを最大幅に固定し、揃え方向を適用 / Pass 2: apply common width and justification */
    for (var j = 0; j < labels.length; j++) {
        labels[j].preferredSize.width = maxWidth;
        labels[j].justify = justify;
    }

    return maxWidth;
}

/**
 * 複数ラベルを固定文字数の幅にして指定方向に揃えます（実測しない簡易版）。
 * PathCleanupTool 方式（`label.characters = N; label.justify = 'right';`）と同じ考え方。
 * 実測しないので手軽だが、和文の実描画幅が chars を超える場合は実幅が優先される点に注意。
 * @param {Array<StaticText>} labels - 幅を揃える statictext の配列。
 * @param {number} chars - 予約する文字数（例: 13）。
 * @param {string} [justify] - 揃え方向（省略時は "right"）。
 * @returns {void}
 */
function setLabelsFixedWidth(labels, chars, justify) {
    if (!labels || !labels.length) return;
    justify = justify || 'right';
    for (var i = 0; i < labels.length; i++) {
        labels[i].characters = chars;
        labels[i].justify = justify;
    }
}

/**
 * 「ラベル：値」形式の1行を親コンテナに追加します（ラベルは後で alignLabelWidths に渡す）。
 * @param {Group|Panel} parent - 行を追加する親コンテナ。
 * @param {string} labelString - ラベル文字列（コロン込みで渡す）。
 * @param {string} [valueString] - 値の初期文字列（省略時は空）。
 * @param {number} [valueChars] - 値欄の予約文字数（省略時は 8）。
 * @returns {{row: Group, label: StaticText, value: StaticText}} 生成した行・ラベル・値。
 */
function addLabelValueRow(parent, labelString, valueString, valueChars) {
    var row = parent.add('group');
    row.orientation = 'row';
    row.alignChildren = ['left', 'center'];

    var label = row.add('statictext', undefined, labelString);

    var value = row.add('statictext', undefined, valueString || '');
    value.characters = (typeof valueChars === 'number') ? valueChars : 8;

    return { row: row, label: label, value: value };
}

// =========================================
//   ↑↑↑ ここまでが再利用パーツ ↑↑↑
// =========================================


// =========================================
// デモ（動作確認用。コピー時は不要） / Demo (for testing; omit when copying)
// =========================================

(function () {
    var dlg = new Window('dialog', 'AlignLabelWidths ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = 16;
    dlg.spacing = 12;

    var infoPanel = dlg.add('panel', undefined, '情報');
    infoPanel.orientation = 'column';
    infoPanel.alignChildren = ['fill', 'top'];
    infoPanel.alignment = 'left'; /* 内容幅に縮めて右余白をなくす / shrink to content */
    infoPanel.margins = [16, 20, 16, 12];
    infoPanel.spacing = 8;

    /* 各行を作りながらラベルを集める / Build rows and collect labels */
    var rows = [
        addLabelValueRow(infoPanel, 'パスの数：', '0'),
        addLabelValueRow(infoPanel, 'アンカーポイント数：', '0 → 0'),
        addLabelValueRow(infoPanel, 'ハンドル数：', '0 → 0')
    ];

    var labels = [];
    for (var i = 0; i < rows.length; i++) labels.push(rows[i].label);

    /* 方式A（推奨）：最長ラベルへ自動で揃えて右揃え / align to the widest and right-justify */
    alignLabelWidths(labels, 'right');

    /* 方式B（簡易・PathCleanupTool と同方式）：固定文字数で右揃え。使う場合は上を消してこちらに差し替え
       setLabelsFixedWidth(labels, 13, 'right'); */

    var btnRow = dlg.add('group');
    btnRow.alignment = ['center', 'top'];
    btnRow.add('button', undefined, 'OK', { name: 'ok' });

    dlg.show();
})();
