#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ScriptUI の複数ラベル（statictext）の幅を、実際の描画幅を測って最長のものへ揃える再利用テンプレートです。
「ラベル：値」を縦に並べるパネルで、コロンの位置と値の開始位置をそろえる用途に使います。

詳細は README を参照してください。

### Overview

A reusable template that measures the rendered width of ScriptUI statictext labels and aligns them all to the widest one.
It keeps the colon and the value column lined up in stacked "label: value" panels.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AlignLabelWidths";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AlignLabelWidths.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AlignLabelWidths.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// レイアウト / Layout
// =========================================
var DEFAULT_VALUE_CHAR_COUNT = 8;  /* 値欄の既定予約文字数 / default reserved width of the value column */

/**
 * 複数のラベル（statictext）の幅を最長のものへ揃え、指定方向に揃えます。
 * 実際の描画幅（preferredSize.width）を測るため、文言長やロケールに依存しません。
 * @param {StaticText[]} labelControls - 幅を揃える statictext の配列。
 * @param {string} [justify] - 揃え方向 "left" | "center" | "right"（省略時は "right"）。
 * @returns {number} 揃えた後の共通ラベル幅（px）。
 */
function alignLabelWidths(labelControls, justify) {
    if (!labelControls || !labelControls.length) return 0;
    justify = justify || 'right';

    /* 1パス目：各ラベルの自然幅を測り、最大値を求める / Pass 1: find the widest natural width */
    var maxLabelWidth = 0;
    for (var i = 0; i < labelControls.length; i++) {
        /* preferredSize を返さない環境では実サイズで代用する / fall back to the laid-out size */
        var naturalWidth = labelControls[i].preferredSize.width || labelControls[i].size.width;
        if (naturalWidth > maxLabelWidth) maxLabelWidth = naturalWidth;
    }

    /* 2パス目：全ラベルを最大幅に固定し、揃え方向を適用 / Pass 2: apply common width and justification */
    for (var j = 0; j < labelControls.length; j++) {
        labelControls[j].preferredSize.width = maxLabelWidth;
        labelControls[j].justify = justify;
    }

    return maxLabelWidth;
}

/**
 * 複数ラベルを固定文字数の幅にして指定方向に揃えます（実測しない簡易版）。
 * PathCleanupTool 方式（`label.characters = N; label.justify = 'right';`）と同じ考え方。
 * 実測しないので手軽だが、和文の実描画幅が charCount を超える場合は実幅が優先される点に注意。
 * @param {StaticText[]} labelControls - 幅を揃える statictext の配列。
 * @param {number} charCount - 予約する文字数（例: 13）。
 * @param {string} [justify] - 揃え方向（省略時は "right"）。
 * @returns {void}
 */
function setLabelsFixedWidth(labelControls, charCount, justify) {
    if (!labelControls || !labelControls.length) return;
    justify = justify || 'right';
    for (var i = 0; i < labelControls.length; i++) {
        labelControls[i].characters = charCount;
        labelControls[i].justify = justify;
    }
}

/**
 * 「ラベル：値」形式の1行を親コンテナに追加します（ラベルは後で alignLabelWidths に渡す）。
 * @param {Group|Panel} parentContainer - 行を追加する親コンテナ。
 * @param {string} labelString - ラベル文字列（コロン込みで渡す）。
 * @param {string} [valueString] - 値の初期文字列（省略時は空）。
 * @param {number} [valueCharCount] - 値欄の予約文字数（省略時は DEFAULT_VALUE_CHAR_COUNT）。
 * @returns {{row: Group, label: StaticText, value: StaticText}} 生成した行・ラベル・値。
 */
function addLabelValueRow(parentContainer, labelString, valueString, valueCharCount) {
    var rowGroup = parentContainer.add('group');
    rowGroup.orientation = 'row';
    rowGroup.alignChildren = ['left', 'center'];

    var labelControl = rowGroup.add('statictext', undefined, labelString);

    var valueControl = rowGroup.add('statictext', undefined, valueString || '');
    valueControl.characters = (typeof valueCharCount === 'number') ? valueCharCount : DEFAULT_VALUE_CHAR_COUNT;

    return { row: rowGroup, label: labelControl, value: valueControl };
}

// =========================================
//   ↑↑↑ ここまでが再利用パーツ ↑↑↑
// =========================================

// =========================================
// デモ（動作確認用。コピー時は不要） / Demo (for testing; omit when copying)
// =========================================

(function () {
    var demoDialog = new Window('dialog', 'AlignLabelWidths ' + SCRIPT_VERSION);
    demoDialog.orientation = 'column';
    demoDialog.alignChildren = ['fill', 'top'];
    demoDialog.margins = 16;
    demoDialog.spacing = 12;

    var pathInfoPanel = demoDialog.add('panel', undefined, 'パス情報');
    pathInfoPanel.orientation = 'column';
    pathInfoPanel.alignChildren = ['fill', 'top'];
    pathInfoPanel.alignment = 'left'; /* 内容幅に縮めて右余白をなくす / shrink to content */
    pathInfoPanel.margins = [16, 20, 16, 12];
    pathInfoPanel.spacing = 8;

    /* 行を作りながら、揃える対象のラベルを集める / Build the rows and collect the labels to align */
    var infoRows = [
        addLabelValueRow(pathInfoPanel, 'パスの数：', '0'),
        addLabelValueRow(pathInfoPanel, 'アンカーポイント数：', '0 → 0'),
        addLabelValueRow(pathInfoPanel, 'ハンドル数：', '0 → 0')
    ];

    var labelControls = [];
    for (var i = 0; i < infoRows.length; i++) labelControls.push(infoRows[i].label);

    /* 方式A（推奨）：最長ラベルへ自動で揃えて右揃え / align to the widest and right-justify */
    alignLabelWidths(labelControls, 'right');

    /* 方式B（簡易・PathCleanupTool と同方式）：固定文字数で右揃え。使う場合は上を消してこちらに差し替え
       setLabelsFixedWidth(labelControls, 13, 'right'); */

    var btnRowGroup = demoDialog.add('group');
    btnRowGroup.alignment = ['center', 'top'];
    btnRowGroup.add('button', undefined, 'OK', { name: 'ok' });

    demoDialog.show();
})();
