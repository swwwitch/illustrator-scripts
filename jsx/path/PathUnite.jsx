#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  概要 / Overview

  選択中のオブジェクトに対して以下の処理を順に適用します。
  1. 複合パスの解除
  2. パスの合体（ライブパスファインダ）
  3. アピアランスの拡張
  4. グループ解除

  Apply the following operations to the selected objects in order:
  1. Release compound paths
  2. Unite paths (Live Pathfinder Add)
  3. Expand appearance
  4. Ungroup
*/

// =========================================
// ローカライズ / Localization
// =========================================

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
