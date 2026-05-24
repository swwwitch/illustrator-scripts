#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
選択オブジェクトに［オブジェクト］>［分割・拡張］を実行し、グラデーションを指定ステップ数で分割する Illustrator スクリプトです。
内部で一時的な .aia アクション（セット名 "Expand" / アクション名 "Expand-gradient"）を生成・ロードして再生し、実行後にアンロードします。

ダイアログ：
・ステップ数（既定値 5、2 以上の整数）。↑↓で±1、Shift+↑↓で10の倍数にスナップ。
・「実行後に拡張」チェックON時に、後処理として次を順に実行します：
  ① グループ解除（ungroup）
  ② クリッピングマスク解除（releaseMask）
  ③ 選択内の「塗りなし／線なし」のパスが1つだけの場合のみ削除（解除されたマスクパス候補を除去）
  ④ 選択内の「塗りあり／線なし／アンカー 2 点」のパスを削除

スクリプト冒頭の設定スイッチ：
・DEFAULT_GRADIENT_STEPS：ステップ数の既定値
・DEFAULT_FURTHER_EXPAND：後処理ON/OFFの既定値
・SHOW_DIALOG：false でダイアログを出さず既定値のまま即実行

This Illustrator script runs Object > Expand on the current selection and converts gradients into the specified number of steps.
A temporary .aia action (set "Expand", action "Expand-gradient") is generated, loaded, played, and then unloaded.

Dialog:
- Step count (default 5, integer ≥ 2). Arrow keys ±1, Shift+Arrow snaps to multiples of 10.
- "Further expand after run" checkbox runs the following post-process in order:
  (1) ungroup
  (2) releaseMask
  (3) Delete only when there is exactly one fill-less, stroke-less path in the selection (former clipping mask path candidate)
  (4) Delete fill-only, stroke-less paths with exactly 2 anchor points

Config switches at the top:
- DEFAULT_GRADIENT_STEPS: default step count
- DEFAULT_FURTHER_EXPAND: default for the post-process toggle
- SHOW_DIALOG: when false, runs silently using the defaults

Updated: 2026-05-25
*/

// =========================================
// 設定スイッチ / Config switches
// =========================================

/* 分割数（既定値、2 以上の整数） / Default gradient step count (integer ≥ 2) */
var DEFAULT_GRADIENT_STEPS = 5;

/* 実行後にグループ解除→クリッピングマスク解除→マスクパス候補削除→塗りのみ・アンカー2点のパス除去を行うかの既定値 / Default for post-process (ungroup → release mask → delete mask path candidate → remove fill-only 2-point paths) */
var DEFAULT_FURTHER_EXPAND = true;

/* ダイアログ表示の有無（false で既定値のまま即実行） / Show dialog (false: run silently with default) */
var SHOW_DIALOG = true;

/* Pathfinder Merge ライブエフェクトの XML（command 8 = Merge 固定）
   Live effect XML for Pathfinder Merge (command 8, all other params at defaults) */
var PATHFINDER_MERGE_XML = '<LiveEffect name="Adobe Pathfinder" isPre="1">'
    + '<Dict data="I Command 8 B ConvertCustom 1 B ExtractUnpainted 1 R Mix 0.5 R Precision 10 B RemovePoints 1 R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0 R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">'
    + '<Entry name="DisplayString" value="Merge" valueType="S"/>'
    + '</Dict></LiveEffect>';

(function () {

    var SCRIPT_VERSION = "v1.0.0";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

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
        }
    };

    function L(key) {
        return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
    }

    function labelText(key) {
        return L(key) + (lang === "ja" ? "：" : ":");
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
        var dialogChoices = showStepsDialog(SCRIPT_VERSION, DEFAULT_GRADIENT_STEPS, DEFAULT_FURTHER_EXPAND, L, labelText);
        if (dialogChoices === null) return;
        gradientSteps = dialogChoices.steps;
        furtherExpand = dialogChoices.furtherExpand;
    }

    playEmbeddedAction(buildExpandActionSource(gradientSteps), "Expand", "Expand-gradient");

    if (furtherExpand) {
        app.executeMenuCommand('ungroup');
        app.executeMenuCommand('releaseMask');
        deleteReleasedMaskPathInSelection();
        deleteFilledTwoPointPathsInSelection();
        applyLiveEffectAndExpand(activeDoc, PATHFINDER_MERGE_XML);
    }

})();

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

function deleteReleasedMaskPathInSelection() {
    var sel = app.activeDocument.selection;
    if (!sel || sel.length === 0) return;

    var candidates = [];
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        /* PathItem かつ塗りも線もないもの＝解除済みマスクパス候補 / PathItem with no fill and no stroke */
        if (isPathItem(item) && !item.filled && !item.stroked) {
            candidates.push(item);
        }
    }

    /* 候補が複数ある場合は誤削除を避ける / Avoid accidental deletion when multiple candidates exist */
    if (candidates.length !== 1) return;
    candidates[0].remove();
}

function isPathItem(item) {
    return item && item.typename === "PathItem";
}

function deleteFilledTwoPointPathsInSelection() {
    var sel = app.activeDocument.selection;
    if (!sel || sel.length === 0) return;

    /* 削除でselectionが変化するためスナップショットを取る / Snapshot since removal mutates the selection */
    var items = [];
    for (var i = 0; i < sel.length; i++) items.push(sel[i]);

    for (var j = items.length - 1; j >= 0; j--) {
        var item = items[j];
        if (!isPathItem(item)) continue;
        if (!item.filled) continue;                 /* 塗りなしはスキップ / Skip paths without fill */
        if (item.stroked) continue;                 /* 線ありはスキップ / Skip paths with stroke */
        if (item.pathPoints.length !== 2) continue; /* アンカー2点以外はスキップ / Skip unless exactly 2 anchor points */
        item.remove();
    }
}

function buildExpandActionSource(gradientSteps) {
    return ''
        + '/version 3'
        + '/name [ 6 457870616e64]'
        + '/isOpen 1'
        + '/actionCount 1'
        + '/action-1 {'
        +   ' /name [ 15 457870616e642d6772616469656e74 ]'
        +   ' /keyIndex 0'
        +   ' /colorIndex 0'
        +   ' /isOpen 1'
        +   ' /eventCount 1'
        +   ' /event-1 {'
        +     ' /useRulersIn1stQuadrant 0'
        +     ' /internalName (ai_plugin_expand)'
        +     ' /localizedName [ 15 e58886e589b2e383bbe68ba1e5bcb5 ]'
        +     ' /isOpen 1'
        +     ' /isOn 1'
        +     ' /hasDialog 1'
        +     ' /showDialog 0'
        +     ' /parameterCount 5'
        +     ' /parameter-1 { /key 1868720756 /showInPalette 4294967295 /type (boolean) /value 0 }'
        +     ' /parameter-2 { /key 1718185068 /showInPalette 4294967295 /type (boolean) /value 1 }'
        +     ' /parameter-3 { /key 1937011307 /showInPalette 4294967295 /type (boolean) /value 0 }'
        +     ' /parameter-4 { /key 1936553064 /showInPalette 4294967295 /type (boolean) /value 0 }'
        +     ' /parameter-5 { /key 1937007984 /showInPalette 4294967295 /type (integer) /value ' + gradientSteps + ' }'
        +   ' }'
        + '}';
}

/* ====== ライブエフェクト適用＋アピアランス分割 / Apply live effect and expand ======
   選択をグループ化し、線を塗りに変換してから指定 XML のライブエフェクトを適用、アピアランスを分割してグループを解除
   Group selection, convert strokes to fills, apply the given live effect XML, expand appearance, then ungroup */
function applyLiveEffectAndExpand(doc, liveEffectXml) {
    app.executeMenuCommand('group');
    /* 線を塗りに変換 / Convert strokes to fills */
    app.executeMenuCommand('OffsetPath v22');
    var group = doc.selection[0];
    group.applyEffect(liveEffectXml);
    app.redraw();
    doc.selection = null;
    group.selected = true;
    app.executeMenuCommand('expandStyle');
    app.executeMenuCommand('ungroup');
}

function playEmbeddedAction(actionSource, setName, actionName) {

    var actionFile = new File('~/ScriptAction.aia');
    var isActionLoaded = false;

    /* 既存の同名セットが残っているとロードが効かないので、先に外す / Unload first in case the set is still loaded */
    try { app.unloadAction(setName, ""); } catch (_) { }

    try {
        if (!actionFile.open('w')) {
            throw new Error('Failed to open temporary action file for writing.');
        }

        actionFile.write(actionSource);
        actionFile.close();

        app.loadAction(actionFile);
        isActionLoaded = true;

        app.doScript(actionName, setName, false);
    } finally {
        if (actionFile.exists) {
            try { actionFile.remove(); } catch (_) { }
        }

        if (isActionLoaded) {
            try { app.unloadAction(setName, ""); } catch (_) { }
        }
    }

}