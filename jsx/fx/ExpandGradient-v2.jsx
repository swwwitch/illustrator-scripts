#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
選択オブジェクトに［オブジェクト］>［分割・拡張］を実行し、グラデーションを指定ステップ数で分割する Illustrator スクリプトです。
内部で一時的な .aia アクションを生成・ロード・再生し、実行後にアンロードして一時ファイルを削除します。
アクションのセット名／アクション名は定数で管理し、.aia 内の名前も同じ値から生成します。

ダイアログ：
・ステップ数（既定値 5、2 以上の整数）。↑↓で±1、Shift+↑↓で10の倍数にスナップ。
・「実行後に拡張」チェックON時に、後処理として次を順に実行します：
  ① Pathfinder ＞ クロップ（Live Pathfinder Crop）
  ② アピアランスを分割（expandStyle）
  ③ Pathfinder Merge ライブエフェクト（command 8）を適用 → アピアランスを分割

安全対策：
・後処理前に対象ドキュメントをアクティブ化します。
・Pathfinder Merge を適用できる選択オブジェクトを確認し、対象がない場合は中断します。
・一時アクションファイルは finally で close/remove を試みます。

スクリプト冒頭の設定スイッチ：
・DEFAULT_GRADIENT_STEPS：ステップ数の既定値
・DEFAULT_FURTHER_EXPAND：後処理ON/OFFの既定値
・SHOW_DIALOG：false でダイアログを出さず既定値のまま即実行
・ACTION_SET_NAME / ACTION_NAME：一時アクションのセット名／アクション名

This Illustrator script runs Object > Expand on the current selection and converts gradients into the specified number of steps.
It creates, loads, plays, unloads, and removes a temporary .aia action file.
The action set name and action name are defined as constants, and the names inside the .aia source are generated from the same values.

Dialog:
- Step count (default 5, integer ≥ 2). Arrow keys ±1, Shift+Arrow snaps to multiples of 10.
- "Further expand after run" runs the following post-process in order:
  (1) Pathfinder > Crop (Live Pathfinder Crop)
  (2) Expand appearance (expandStyle)
  (3) Apply Pathfinder Merge live effect (command 8) → expand appearance

Safety:
- The target document is activated before post-processing.
- A valid selection item for Pathfinder Merge is checked before applying the effect.
- The temporary action file is closed/removed in finally whenever possible.

Config switches at the top:
- DEFAULT_GRADIENT_STEPS: default step count
- DEFAULT_FURTHER_EXPAND: default for the post-process toggle
- SHOW_DIALOG: when false, runs silently using the defaults
- ACTION_SET_NAME / ACTION_NAME: temporary action set/action names

Updated: 2026-05-25
*/

// =========================================
// 設定スイッチ / Config switches
// =========================================

/* 分割数（既定値、2 以上の整数） / Default gradient step count (integer ≥ 2) */
var DEFAULT_GRADIENT_STEPS = 5;

/* 後処理の既定値：Pathfinder クロップ → アピアランス分割 → Pathfinder Merge → アピアランス分割 / Default post-process toggle: Pathfinder Crop → expand → Pathfinder Merge → expand */
var DEFAULT_FURTHER_EXPAND = true;

/* ダイアログ表示の有無（false で既定値のまま即実行） / Show dialog (false: run silently with default) */
var SHOW_DIALOG = true;

/* 一時アクションのセット名 / Temporary action set name */
var ACTION_SET_NAME = "Expand";

/* 一時アクションのアクション名 / Temporary action name */
var ACTION_NAME = "Expand-gradient";

(function () {

    var SCRIPT_VERSION = "v1.0.1";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {

        /* === 共通 / Common === */
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },

        /* === ダイアログ / Dialog === */
        dialogTitle: {
            ja: "グラデーションを分割・拡張",
            en: "Expand Gradient"
        },
        steps: {
            ja: "ステップ数",
            en: "Steps"
        },
        furtherExpand: {
            ja: "実行後に拡張",
            en: "Further expand after run"
        },

        /* === アラート / Alerts === */
        alertNoDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        alertNoSelection: {
            ja: "オブジェクトを選択してください。",
            en: "Please select an object."
        },
        alertInvalidSteps: {
            ja: "ステップ数は 2 以上の整数で指定してください。",
            en: "Steps must be an integer of 2 or more."
        },
        alertMergeTargetNotFound: {
            ja: "Pathfinder Merge を適用できる対象を取得できませんでした。処理を中断します。",
            en: "Could not find a valid target for applying Pathfinder Merge. The process will stop."
        }
    };

    function L(key) {
        return (LABELS[key] && LABELS[key][currentLanguage]) ? LABELS[key][currentLanguage] : key;
    }

    function labelText(key) {
        return L(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 入口チェック / Entry checks
    // =========================================

    if (app.documents.length === 0) {
        alert(L("alertNoDocument"));
        return;
    }

    var activeDoc = app.activeDocument;
    if (activeDoc.selection.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    var gradientSteps = DEFAULT_GRADIENT_STEPS;
    var furtherExpand = DEFAULT_FURTHER_EXPAND;

    if (SHOW_DIALOG) {
        var dialogResult = showStepsDialog(SCRIPT_VERSION, DEFAULT_GRADIENT_STEPS, DEFAULT_FURTHER_EXPAND, L, labelText);
        if (dialogResult === null) return;
        gradientSteps = dialogResult.steps;
        furtherExpand = dialogResult.furtherExpand;
    }

    playEmbeddedAction(buildExpandActionSource(gradientSteps, ACTION_SET_NAME, ACTION_NAME), ACTION_SET_NAME, ACTION_NAME);

    if (furtherExpand) {
        if (!furtherExpandSelection(activeDoc)) {
            alert(L("alertMergeTargetNotFound"));
            return;
        }
    }

})();

// =========================================
// 後処理 / Post-processing
// =========================================

function furtherExpandSelection(targetDoc) {
    targetDoc.activate();
    app.executeMenuCommand('Live Pathfinder Crop');
    app.executeMenuCommand('expandStyle');

    /* Pathfinder Merge ライブエフェクト（command 8）を適用し、アピアランスを分割
       Apply the Pathfinder Merge live effect (command 8) and expand appearance */
    var pathfinderMergeXml = '<LiveEffect name="Adobe Pathfinder" isPre="1">'
        + '<Dict data="I Command 8 B ConvertCustom 1 B ExtractUnpainted 1 R Mix 0.5 R Precision 10 B RemovePoints 1 R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0 R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">'
        + '<Entry name="DisplayString" value="Merge" valueType="S"/>'
        + '</Dict></LiveEffect>';

    var mergeTarget = getFirstEffectApplicableSelectionItem(targetDoc);
    if (mergeTarget === null) return false;

    mergeTarget.applyEffect(pathfinderMergeXml);
    app.redraw();
    app.executeMenuCommand("deselectall");
    mergeTarget.selected = true;
    app.executeMenuCommand('expandStyle');

    return true;
}

function getFirstEffectApplicableSelectionItem(targetDoc) {
    if (!targetDoc || !targetDoc.selection || targetDoc.selection.length === 0) return null;

    for (var i = 0; i < targetDoc.selection.length; i++) {
        var selectedItem = targetDoc.selection[i];
        if (selectedItem && typeof selectedItem.applyEffect === "function") {
            return selectedItem;
        }
    }

    return null;
}

// =========================================
// ダイアログUI / Dialog UI
// =========================================

function showStepsDialog(scriptVersion, defaultSteps, defaultFurtherExpand, L, labelText) {

    var stepsDialog = new Window("dialog", L("dialogTitle") + " " + scriptVersion);
    stepsDialog.orientation = "column";
    stepsDialog.alignChildren = ["fill", "top"];
    stepsDialog.margins = 16;

    var stepsRow = stepsDialog.add("group");
    stepsRow.orientation = "row";
    stepsRow.alignChildren = ["left", "center"];
    stepsRow.add("statictext", undefined, labelText("steps"));
    var stepsInput = stepsRow.add("edittext", undefined, String(defaultSteps));
    stepsInput.characters = 5;
    stepsInput.active = true;
    changeValueByArrowKey(stepsInput);

    var furtherExpandCb = stepsDialog.add("checkbox", undefined, L("furtherExpand"));
    furtherExpandCb.value = defaultFurtherExpand;

    var okCancelGroup = stepsDialog.add("group");
    okCancelGroup.alignment = ["right", "center"];
    okCancelGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    okCancelGroup.add("button", undefined, L("ok"), { name: "ok" });

    if (stepsDialog.show() !== 1) return null;

    var parsedSteps = parseInt(stepsInput.text, 10);
    if (isNaN(parsedSteps) || parsedSteps < 2) {
        alert(L("alertInvalidSteps"));
        return null;
    }
    return { steps: parsedSteps, furtherExpand: furtherExpandCb.value };

}

function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        var value = parseInt(editText.text, 10);
        if (isNaN(value)) return;

        var sign = (event.keyName === "Up") ? 1 : -1;
        if (ScriptUI.environment.keyboardState.shiftKey) {
            /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
            value = (sign > 0) ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
        } else {
            value += sign;
        }

        if (value < 2) value = 2;
        editText.text = String(Math.round(value));
        event.preventDefault();
    });
}

// =========================================
// 一時アクション生成 / Temporary action generation
// =========================================

function buildExpandActionSource(gradientSteps, setName, actionName) {
    return ''
        + '/version 3'
        + buildActionNameLine(setName)
        + '/isOpen 1'
        + '/actionCount 1'
        + '/action-1 {'
        + ' ' + buildActionNameLine(actionName)
        + ' /keyIndex 0'
        + ' /colorIndex 0'
        + ' /isOpen 1'
        + ' /eventCount 1'
        + ' /event-1 {'
        + ' /useRulersIn1stQuadrant 0'
        + ' /internalName (ai_plugin_expand)'
        + ' /localizedName [ 15 e58886e589b2e383bbe68ba1e5bcb5 ]'
        + ' /isOpen 1'
        + ' /isOn 1'
        + ' /hasDialog 1'
        + ' /showDialog 0'
        + ' /parameterCount 5'
        + ' /parameter-1 { /key 1868720756 /showInPalette 4294967295 /type (boolean) /value 0 }'
        + ' /parameter-2 { /key 1718185068 /showInPalette 4294967295 /type (boolean) /value 1 }'
        + ' /parameter-3 { /key 1937011307 /showInPalette 4294967295 /type (boolean) /value 0 }'
        + ' /parameter-4 { /key 1936553064 /showInPalette 4294967295 /type (boolean) /value 0 }'
        + ' /parameter-5 { /key 1937007984 /showInPalette 4294967295 /type (integer) /value ' + gradientSteps + ' }'
        + ' }'
        + '}';
}

function buildActionNameLine(actionName) {
    return '/name [ ' + actionName.length + ' ' + stringToHex(actionName) + ' ]\n';
}

function stringToHex(sourceText) {
    var hexText = "";
    for (var i = 0; i < sourceText.length; i++) {
        var hexValue = sourceText.charCodeAt(i).toString(16);
        if (hexValue.length < 2) hexValue = "0" + hexValue;
        hexText += hexValue;
    }
    return hexText;
}

// =========================================
// 一時アクション実行 / Temporary action playback
// =========================================

function playEmbeddedAction(actionSource, setName, actionName) {

    var actionFile = new File('~/ExpandGradientAction.aia');
    var isActionLoaded = false;
    var isActionFileOpen = false;

    /* 既存の同名セットが残っているとロードが効かないので、先に外す / Unload first in case the set is still loaded */
    try { app.unloadAction(setName, ""); } catch (e) { }

    try {
        if (!actionFile.open('w')) {
            throw new Error('Failed to open temporary action file for writing.');
        }
        isActionFileOpen = true;

        actionFile.write(actionSource);
        actionFile.close();
        isActionFileOpen = false;

        app.loadAction(actionFile);
        isActionLoaded = true;

        app.doScript(actionName, setName, false);
    } finally {
        if (isActionFileOpen) {
            try { actionFile.close(); } catch (e) { }
        }

        if (actionFile.exists) {
            try { actionFile.remove(); } catch (e) { }
        }

        if (isActionLoaded) {
            try { app.unloadAction(setName, ""); } catch (e) { }
        }
    }

}