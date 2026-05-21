#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

PathUnite.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathUnite.md

### 概要：

- 更新日：2026-05-10
- 選択中のオブジェクトに対して、複合パスの解除 → パスの合体 → アピアランスの拡張 → グループ解除を一括実行する Illustrator スクリプト

### 主な機能：

- 複合パスの解除
- パスの合体（ライブパスファインダ）
- アピアランスの拡張
- グループ解除

### 処理の流れ：

1) 複合パスの解除
2) パスの合体（ライブパスファインダ）
3) アピアランスの拡張
4) グループ解除

### 更新履歴：

- v1.0.0 (2026-05-10) : 初期バージョン

*/

/*

### Script Name:

PathUnite.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathUnite.md

### Description:

- Last Updated: 2026-05-10
- Applies Release Compound Path → Unite (Live Pathfinder Add) → Expand Appearance → Ungroup in sequence to the current selection

### Main Features:

- Release compound paths
- Unite paths (Live Pathfinder Add)
- Expand appearance
- Ungroup

### Process Flow:

1) Release compound paths
2) Unite paths (Live Pathfinder Add)
3) Expand appearance
4) Ungroup

### Changelog:

- v1.0.0 (2026-05-10) : Initial version

*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

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
