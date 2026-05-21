#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

PathUniteOffsetTool.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathUniteOffsetTool.md

### 概要：

- 更新日：2026-05-10
- 選択中のオブジェクトに対して複合パス解除 → パスの合体 → アピアランス拡張 → グループ解除 → オフセットパスを一括実行
- ダイアログでプレビューを有効化すると、閉じずに結果を確認可能

### 主な機能：

- 複合パスの解除
- パスの合体（ライブパスファインダ）
- アピアランスの拡張
- グループ解除
- 指定値（mm）でオフセットパス効果を適用
- プレビュー機能

### 処理の流れ：

1) 複合パスの解除
2) パスの合体（ライブパスファインダ）
3) アピアランスの拡張
4) グループ解除
5) 指定値（mm）でオフセットパス効果を適用

### 更新履歴：

- v1.0.0 (2026-05-10) : 初期バージョン

*/

/*

### Script Name:

PathUniteOffsetTool.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathUniteOffsetTool.md

### Description:

- Last Updated: 2026-05-10
- Applies Release Compound Path → Unite → Expand Appearance → Ungroup → Offset Path in sequence to the current selection
- When the Preview checkbox is enabled, the result is shown live without closing the dialog

### Main Features:

- Release compound paths
- Unite paths (Live Pathfinder Add)
- Expand appearance
- Ungroup
- Apply Offset Path effect with the specified value (mm)
- Live preview

### Process Flow:

1) Release compound paths
2) Unite paths (Live Pathfinder Add)
3) Expand appearance
4) Ungroup
5) Apply Offset Path effect with the specified value (mm)

### Changelog:

- v1.0.0 (2026-05-10) : Initial version

*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "マド埋めとオフセット",
        en: "Fill Gaps & Offset"
    },
    offsetLabel: {
        ja: "オフセット",
        en: "Offset"
    },
    previewLabel: {
        ja: "プレビュー",
        en: "Preview"
    },
    joinTypeLabel: {
        ja: "角の形状",
        en: "Joins"
    },
    joinRound: {
        ja: "ラウンド",
        en: "Round"
    },
    joinBevel: {
        ja: "ベベル",
        en: "Bevel"
    },
    joinMiter: {
        ja: "マイター",
        en: "Miter"
    },
    okButton: {
        ja: "OK",
        en: "OK"
    },
    cancelButton: {
        ja: "キャンセル",
        en: "Cancel"
    },
    errorNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    errorNoSelection: {
        ja: "オブジェクトが選択されていません。",
        en: "No objects are selected."
    },
    errorPreviewDuplicateFailed: {
        ja: "一部のオブジェクトをプレビュー用に複製できませんでした。ロックされたレイヤーや複製できないオブジェクトが含まれていないか確認してください。",
        en: "Some objects could not be duplicated for preview. Check for locked layers or objects that cannot be duplicated."
    },
    errorPreviewEmpty: {
        ja: "プレビュー用のオブジェクトを作成できませんでした。",
        en: "Could not create preview objects."
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

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLabel(key) + (lang === 'ja' ? '：' : ':');
}

var DIALOG_MARGINS = 15;
var PANEL_MARGINS = [15, 20, 15, 10];

/* パネル共通設定。panel は必ずこの関数を通す / Shared panel setup. Always pass panels through this function. */
function setupPanel(panel, spacing, orientation) {
    panel.orientation = orientation || "column";
    panel.alignChildren = "left";
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    if (typeof spacing === "number") {
        panel.spacing = spacing;
    }
}

// =========================================
// ユーティリティ / Utilities
// =========================================

/* 定規の単位設定を取得 / Get the current ruler unit setting
   rulerType: 0=inch, 1=mm, 2=pt, 3=pica, 4=cm, 5=Q, 6=px */
function getRulerUnit() {
    var rulerType = app.preferences.getIntegerPreference("rulerType");
    var label = "pt";
    var factor = 1.0;

    switch (rulerType) {
        case 0: label = "inch"; factor = 72.0; break;
        case 1: label = "mm";   factor = 72.0 / 25.4; break;
        case 2: label = "pt";   factor = 1.0; break;
        case 3: label = "pica"; factor = 12.0; break;
        case 4: label = "cm";   factor = 72.0 / 2.54; break;
        case 5: label = "Q";    factor = 72.0 / 25.4 * 0.25; break;
        case 6: label = "px";   factor = 1.0; break;
        default: label = "pt";  factor = 1.0;
    }
    return { label: label, factor: factor };
}

/* joinTypes: 0 = Round, 1 = Bevel, 2 = Miter */
function buildOffsetEffectXml(offsetInPoints, joinType) {
    return '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetInPoints + ' I jntp ' + joinType + ' "/></LiveEffect>';
}

/* ↑↓キーでの値増減 / Arrow key value adjustment
   - ↑↓: ±1 (整数丸め / round to integer)
   - shift + ↑↓: ±10 (10の倍数にスナップ / snap to multiples of 10)
   - option + ↑↓: ±0.1 (小数第1位 / 1 decimal place)
*/
function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            if (event.keyName == "Up") {
                value += 1;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= 1;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
        } else {
            value = Math.round(value); /* 整数に丸め / Round to integer */
        }

        if (!allowNegative && value < 0) value = 0;

        editText.text = value;
        /* プレビュー更新 / Update preview */
        if (typeof onValueChanged === "function") {
            onValueChanged();
        }
    });
}

// =========================================
// オフセット処理 / Offset processing
// =========================================

function applyOffsetEffectToSelection(offsetValue, unitFactor, joinType) {
    var doc = app.activeDocument;
    var selection = doc.selection;
    if (!selection || selection.length === 0) return 0;

    var offsetInPoints = offsetValue * unitFactor;
    var effectXml = buildOffsetEffectXml(offsetInPoints, joinType);
    var appliedCount = 0;
    for (var i = 0; i < selection.length; i++) {
        selection[i].applyEffect(effectXml);
        appliedCount++;
    }
    return appliedCount;
}

function removePreviewLayer(previewLayer) {
    if (!previewLayer) return;
    try {
        previewLayer.locked = false;
        previewLayer.visible = true;
        previewLayer.remove();
    } catch (removePreviewLayerError) {
    }
}

function createPreviewLayerFromSelection(initialSelection) {
    var doc = app.activeDocument;
    var previewLayer = doc.layers.add();
    previewLayer.name = "__preview_path_unite_offset__";

    var duplicatedCount = 0;
    var failedCount = 0;

    for (var itemIndex = 0; itemIndex < initialSelection.length; itemIndex++) {
        try {
            initialSelection[itemIndex].duplicate(previewLayer, ElementPlacement.PLACEATEND);
            duplicatedCount++;
        } catch (duplicateError) {
            failedCount++;
        }
    }

    return {
        layer: previewLayer,
        duplicatedCount: duplicatedCount,
        failedCount: failedCount
    };
}

function selectLayerItems(layer) {
    var doc = app.activeDocument;
    doc.selection = null;

    for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
        try {
            layer.pageItems[itemIndex].selected = true;
        } catch (selectItemError) {
        }
    }
}

/* 一連のオフセット処理を実行 / Run the full offset workflow */
function runOffsetWorkflow(offsetValue, unitFactor, joinType) {
    app.executeMenuCommand('noCompoundPath');
    app.executeMenuCommand('Live Pathfinder Add');
    app.executeMenuCommand('expandStyle');

    /* グループでない場合は失敗することがあるため無視 / Ignore when the selection is not grouped. */
    try {
        app.executeMenuCommand('ungroup');
    } catch (ungroupError) {
    }

    applyOffsetEffectToSelection(offsetValue, unitFactor, joinType);
}

// =========================================
// ダイアログ / Dialog
// =========================================

function showOffsetDialog(rulerUnit) {
    var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
    dialog.orientation = 'column';
    dialog.alignChildren = 'left';
    dialog.margins = DIALOG_MARGINS;

    /* 入力欄 / Input row */
    var inputGroup = dialog.add('group');
    inputGroup.add('statictext', undefined, labelText('offsetLabel'));
    var offsetInput = inputGroup.add('edittext', undefined, '1');
    offsetInput.characters = 3;
    inputGroup.add('statictext', undefined, rulerUnit.label);

    /* 角の形状（ラジオボタン）/ Join type (radio buttons) */
    var joinPanel = dialog.add('panel', undefined, getLabel('joinTypeLabel'));
    setupPanel(joinPanel, 6, 'row');
    var joinRoundRadio = joinPanel.add('radiobutton', undefined, getLabel('joinRound'));
    var joinBevelRadio = joinPanel.add('radiobutton', undefined, getLabel('joinBevel'));
    var joinMiterRadio = joinPanel.add('radiobutton', undefined, getLabel('joinMiter'));
    joinRoundRadio.value = true;

    function getSelectedJoinType() {
        if (joinBevelRadio.value) return 1;
        if (joinMiterRadio.value) return 2;
        return 0; /* Round */
    }

    /* プレビュー / Preview */
    var previewCheckbox = dialog.add('checkbox', undefined, getLabel('previewLabel'));
    previewCheckbox.value = false;

    /* 初期選択を保持 / Preserve initial selection */
    var initialSelection = app.activeDocument.selection;
    var previewLayer = null;

    /* ボタン / Buttons */
    var buttonGroup = dialog.add('group');
    buttonGroup.alignment = ['center', 'center'];
    buttonGroup.margins = [0, 10, 0, 0];
    buttonGroup.add('button', undefined, getLabel('cancelButton'));
    buttonGroup.add('button', undefined, getLabel('okButton'), { name: 'ok' });

    /* プレビュー状態 / Preview state */
    var previewedValue = null;
    var previewedJoinType = null;

    function clearPreview() {
        removePreviewLayer(previewLayer);
        previewLayer = null;
        previewedValue = null;
        previewedJoinType = null;

        try {
            app.activeDocument.selection = initialSelection;
        } catch (restoreSelectionError) {
        }

        app.redraw();
    }

    function refreshPreview() {
        clearPreview();
        if (!previewCheckbox.value) return;

        var currentValue = parseFloat(offsetInput.text);
        if (isNaN(currentValue)) return;

        var currentJoinType = getSelectedJoinType();
        var previewResult = createPreviewLayerFromSelection(initialSelection);
        previewLayer = previewResult.layer;

        if (previewResult.failedCount > 0) {
            clearPreview();
            alert(getLabel('errorPreviewDuplicateFailed'));
            previewCheckbox.value = false;
            return;
        }

        if (previewResult.duplicatedCount === 0) {
            clearPreview();
            alert(getLabel('errorPreviewEmpty'));
            previewCheckbox.value = false;
            return;
        }

        selectLayerItems(previewLayer);
        runOffsetWorkflow(currentValue, rulerUnit.factor, currentJoinType);

        previewedValue = currentValue;
        previewedJoinType = currentJoinType;

        try {
            app.activeDocument.selection = initialSelection;
        } catch (restoreSelectionError) {
        }

        app.redraw();
    }

    previewCheckbox.onClick = refreshPreview;
    offsetInput.onChange = function () {
        if (previewCheckbox.value) refreshPreview();
    };
    joinRoundRadio.onClick = function () { if (previewCheckbox.value) refreshPreview(); };
    joinBevelRadio.onClick = function () { if (previewCheckbox.value) refreshPreview(); };
    joinMiterRadio.onClick = function () { if (previewCheckbox.value) refreshPreview(); };

    /* ↑↓キーでの増減（プレビュー連動）/ Arrow keys with preview sync */
    changeValueByArrowKey(offsetInput, true, function () {
        if (previewCheckbox.value) refreshPreview();
    });

    offsetInput.active = true;

    var dialogResult = dialog.show();

    if (dialogResult != 1) {
        clearPreview();
        return null;
    }

    var finalValue = parseFloat(offsetInput.text);
    if (isNaN(finalValue)) {
        clearPreview();
        return null;
    }
    var finalJoinType = getSelectedJoinType();

    clearPreview();

    return {
        value: finalValue,
        joinType: finalJoinType,
        alreadyApplied: false
    };
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

    var rulerUnit = getRulerUnit();
    var dialogResult = showOffsetDialog(rulerUnit);
    if (dialogResult === null) return;

    runOffsetWorkflow(dialogResult.value, rulerUnit.factor, dialogResult.joinType);
})();
