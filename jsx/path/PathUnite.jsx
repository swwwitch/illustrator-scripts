#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトに対して、複合パスの解除→パスの合体→アピアランスの拡張→グループ解除を一括で実行します。

詳細は README を参照してください。

### Overview

Runs Release Compound Path, Unite, Expand Appearance and Ungroup on the selection in one pass.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathUnite";                    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-10";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathUnite.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathUnite.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    errorNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    errorNoSelection: {
        ja: "オブジェクトが選択されていません。",
        en: "No objects are selected."
    }
};

function getLabel(key) {
    if (LABELS[key] && LABELS[key][lang]) {
        return LABELS[key][lang];
    }
    if (LABELS[key] && LABELS[key].en) {
        return LABELS[key].en;
    }
    return key;
}

// =========================================
// パス合体処理 / Path unite processing
// =========================================

function runUniteWorkflow() {
    app.executeMenuCommand('group');
    app.executeMenuCommand('noCompoundPath');
    app.executeMenuCommand('Live Pathfinder Add');
    app.executeMenuCommand('expandStyle');

    /* グループでない場合は失敗することがあるため無視 / Ignore when the selection is not grouped. */
    try {
        app.executeMenuCommand('ungroup');
    } catch (ungroupError) {
    }
}

// =========================================
// メイン / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(getLabel('errorNoDocument'));
        return;
    }

    var selection = app.activeDocument.selection;
    if (!selection || selection.length === 0) {
        alert(getLabel('errorNoSelection'));
        return;
    }

    runUniteWorkflow();
})();
