#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
選択オブジェクトに［オブジェクト］>［分割・拡張］を実行し、グラデーションを指定ステップ数で分割する Illustrator スクリプトです。
内部で一時的な .aia アクションを生成・ロード・再生し、実行後にアンロードして一時ファイルを削除します。
アクションのセット名／アクション名は定数で管理し、.aia 内の名前も同じ値から生成します。

ダイアログ：
・ステップ数（既定値 5、2 以上の整数）。↑↓で±1、Shift+↑↓で10の倍数にスナップ。
・「実行後の処理」パネルで次の3モードから選択：
  - なし：分割・拡張のみ
  - 単純に拡張：分割・拡張後に
      ① Pathfinder ＞ クロップ（Live Pathfinder Crop）
      ② アピアランスを分割（expandStyle）
      ③ Pathfinder Merge ライブエフェクト（command 8）を適用 → アピアランスを分割
  - ブレンドに変換：「単純に拡張」と同じ前処理を実行した後に
      ④ 配下の塗りパスを再帰収集（コンパウンドは subpath まで展開）し、最前面・最背面以外を削除
      ⑤ 残った2点をアクティブレイヤー直下へ移動し、包んでいたグループ／コンパウンドを除去
      ⑥ 2点でブレンド作成（Path Blend Make）
      ⑦ ブレンドオプションを Specified Steps = (gradientSteps − 2) に設定（一時アクション ai_plugin_liveblend）

安全対策：
・後処理前に対象ドキュメントをアクティブ化します。
・Pathfinder Merge を適用できる選択オブジェクトを確認し、対象がない場合は中断します。
・ブレンド変換時、塗り（または線）のあるパスが2点未満の場合は中断します。
・一時アクションファイルは finally で close/remove を試みます。

スクリプト冒頭の設定スイッチ：
・DEFAULT_GRADIENT_STEPS：ステップ数の既定値
・DEFAULT_POST_PROCESS_MODE：後処理モードの既定値（"none" / "simple" / "blend"）
・SHOW_DIALOG：false でダイアログを出さず既定値のまま即実行
・ACTION_SET_NAME / ACTION_NAME：一時アクションのセット名／アクション名

This Illustrator script runs Object > Expand on the current selection and converts gradients into the specified number of steps.
It creates, loads, plays, unloads, and removes a temporary .aia action file.
The action set name and action name are defined as constants, and the names inside the .aia source are generated from the same values.

Dialog:
- Step count (default 5, integer ≥ 2). Arrow keys ±1, Shift+Arrow snaps to multiples of 10.
- "Post-Processing" panel offers three modes:
  - None: expand only
  - Simple expand: after the expand,
      (1) Pathfinder > Crop (Live Pathfinder Crop)
      (2) Expand appearance (expandStyle)
      (3) Apply Pathfinder Merge live effect (command 8) → expand appearance
  - Convert to blend: run the same pre-processing as Simple expand, then
      (4) Recursively collect painted leaf paths (compound paths are expanded down to their sub-paths) and remove all except the frontmost and backmost
      (5) Lift the two remaining paths up to the active layer and drop the empty wrappers
      (6) Run Path Blend Make on the two items
      (7) Set the blend's Specified Steps to (gradientSteps − 2) via a temporary action (ai_plugin_liveblend)

Safety:
- The target document is activated before post-processing.
- A valid selection item for Pathfinder Merge is checked before applying the effect.
- Blend conversion aborts when fewer than 2 painted paths remain.
- The temporary action file is closed/removed in finally whenever possible.

Config switches at the top:
- DEFAULT_GRADIENT_STEPS: default step count
- DEFAULT_POST_PROCESS_MODE: default post-process mode ("none" / "simple" / "blend")
- SHOW_DIALOG: when false, runs silently using the defaults
- ACTION_SET_NAME / ACTION_NAME: temporary action set/action names

Updated: 2026-05-25
*/

// =========================================
// 設定スイッチ / Config switches
// =========================================

/* 分割数（既定値、2 以上の整数） / Default gradient step count (integer ≥ 2) */
var DEFAULT_GRADIENT_STEPS = 5;

/* 後処理モードの既定値（"none" / "simple" / "blend"） / Default post-process mode ("none" / "simple" / "blend") */
var DEFAULT_POST_PROCESS_MODE = "simple";

/* ダイアログ表示の有無（false で既定値のまま即実行） / Show dialog (false: run silently with default) */
var SHOW_DIALOG = true;

/* 一時アクションのセット名（衝突回避のためユニーク名） / Temporary action set name (unique to avoid collisions) */
var ACTION_SET_NAME = "ExpandGradient_tmp";

/* 一時アクションのアクション名 / Temporary action name */
var ACTION_NAME = "Expand-gradient";

/* ブレンドオプション用の一時アクション名 / Temporary action names for blend options */
var BLEND_ACTION_SET_NAME = "ExpandGradient_blend_tmp";
var BLEND_ACTION_NAME = "setBlendStep";

(function () {

    var SCRIPT_VERSION = "v1.1.0";

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
        postProcessPanel: {
            ja: "実行後の処理",
            en: "Post-Processing"
        },
        postProcessNone: {
            ja: "なし",
            en: "None"
        },
        postProcessSimple: {
            ja: "単純に拡張",
            en: "Simple expand"
        },
        postProcessBlend: {
            ja: "ブレンドに変換",
            en: "Convert to blend"
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
        },
        alertBlendTargetNotFound: {
            ja: "ブレンドに必要なオブジェクトを取得できませんでした。処理を中断します。",
            en: "Could not get the objects needed for the blend. The process will stop."
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
    var postProcessMode = DEFAULT_POST_PROCESS_MODE;

    if (SHOW_DIALOG) {
        var dialogResult = showStepsDialog(SCRIPT_VERSION, DEFAULT_GRADIENT_STEPS, DEFAULT_POST_PROCESS_MODE, L, labelText);
        if (dialogResult === null) return;
        gradientSteps = dialogResult.steps;
        postProcessMode = dialogResult.postProcessMode;
    }

    playEmbeddedAction(buildExpandActionSource(gradientSteps, ACTION_SET_NAME, ACTION_NAME), ACTION_SET_NAME, ACTION_NAME);

    if (postProcessMode === "simple") {
        if (!furtherExpandSelection(activeDoc)) {
            alert(L("alertMergeTargetNotFound"));
            return;
        }
    } else if (postProcessMode === "blend") {
        if (!convertToBlend(activeDoc, gradientSteps)) {
            alert(L("alertBlendTargetNotFound"));
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

function convertToBlend(targetDoc, gradientSteps) {
    /* furtherExpandSelection と同じ前処理（Crop → expandStyle → Merge → expandStyle）
       Same pre-processing as furtherExpandSelection */
    if (!furtherExpandSelection(targetDoc)) return false;

    /* 包んでいるグループ／コンパウンドを後で片付けるため記録
       Remember the wrapping containers so we can drop them later */
    var originalContainers = [];
    for (var i = 0; i < targetDoc.selection.length; i++) {
        originalContainers.push(targetDoc.selection[i]);
    }

    /* 選択配下から塗り（または線）のある葉パスを再帰収集
       Recursively collect painted leaf paths under the selection */
    var paintedPaths = [];
    for (var i = 0; i < targetDoc.selection.length; i++) {
        collectPaintedDescendantPaths(targetDoc.selection[i], paintedPaths);
    }
    if (paintedPaths.length < 2) return false;

    /* 親共通なら親内 z 順（前面=0）で並べ替え / Sort by parent's z-order when parents are shared */
    paintedPaths = sortByZOrderWhenShared(paintedPaths);

    var frontItem = paintedPaths[0];
    var backItem = paintedPaths[paintedPaths.length - 1];
    for (var i = 1; i < paintedPaths.length - 1; i++) {
        paintedPaths[i].remove();
    }

    /* front/back をレイヤー直下へ移動し、空になった包みを除去
       Lift front/back to the layer and drop the now-empty wrappers */
    var hostLayer = targetDoc.activeLayer;
    frontItem.move(hostLayer, ElementPlacement.PLACEATEND);
    backItem.move(hostLayer, ElementPlacement.PLACEATEND);
    for (var i = 0; i < originalContainers.length; i++) {
        try { originalContainers[i].remove(); } catch (e) { }
    }

    /* 2点を選択してブレンド作成 / Select the two items and run Blend Make */
    app.executeMenuCommand("deselectall");
    backItem.selected = true;
    frontItem.selected = true;
    app.executeMenuCommand('Path Blend Make');

    /* ブレンドのステップ数を gradientSteps - 2 で設定 / Set blend specified steps */
    var blendSteps = (gradientSteps - 2 > 0) ? gradientSteps - 2 : 0;
    playEmbeddedAction(
        buildBlendOptionsActionSource(blendSteps, BLEND_ACTION_SET_NAME, BLEND_ACTION_NAME),
        BLEND_ACTION_SET_NAME,
        BLEND_ACTION_NAME
    );

    return true;
}

function collectPaintedDescendantPaths(item, result) {
    if (!item) return;
    if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            collectPaintedDescendantPaths(item.pageItems[i], result);
        }
    } else if (item.typename === "CompoundPathItem") {
        for (var i = 0; i < item.pathItems.length; i++) {
            collectPaintedDescendantPaths(item.pathItems[i], result);
        }
    } else if (item.typename === "PathItem") {
        if (item.clipping) return;
        if (!item.filled && !item.stroked) return;
        if (item.pathPoints && item.pathPoints.length === 2) return;
        result.push(item);
    }
}

function sortByZOrderWhenShared(paths) {
    if (paths.length < 2) return paths;
    var sharedParent = paths[0].parent;
    for (var i = 1; i < paths.length; i++) {
        if (paths[i].parent !== sharedParent) return paths;
    }
    var indexed = [];
    for (var i = 0; i < paths.length; i++) {
        for (var j = 0; j < sharedParent.pageItems.length; j++) {
            if (sharedParent.pageItems[j] === paths[i]) {
                indexed.push({ item: paths[i], zIndex: j });
                break;
            }
        }
    }
    indexed.sort(function (a, b) { return a.zIndex - b.zIndex; });
    var sorted = [];
    for (var i = 0; i < indexed.length; i++) sorted.push(indexed[i].item);
    return sorted;
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

function showStepsDialog(scriptVersion, defaultSteps, defaultPostProcessMode, L, labelText) {

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

    var postProcessPanel = stepsDialog.add("panel", undefined, L("postProcessPanel"));
    postProcessPanel.orientation = "column";
    postProcessPanel.alignChildren = ["left", "top"];
    postProcessPanel.margins = [15, 20, 15, 10];

    var postProcessNoneRb = postProcessPanel.add("radiobutton", undefined, L("postProcessNone"));
    var postProcessSimpleRb = postProcessPanel.add("radiobutton", undefined, L("postProcessSimple"));
    var postProcessBlendRb = postProcessPanel.add("radiobutton", undefined, L("postProcessBlend"));

    if (defaultPostProcessMode === "blend") {
        postProcessBlendRb.value = true;
    } else if (defaultPostProcessMode === "simple") {
        postProcessSimpleRb.value = true;
    } else {
        postProcessNoneRb.value = true;
    }

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

    var selectedMode = "none";
    if (postProcessSimpleRb.value) selectedMode = "simple";
    else if (postProcessBlendRb.value) selectedMode = "blend";

    return { steps: parsedSteps, postProcessMode: selectedMode };

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

function buildBlendOptionsActionSource(blendSteps, setName, actionName) {
    /* ai_plugin_liveblend に Specified Steps / step 数 / 方向(Page) を渡す
       Pass Specified Steps mode, step count, and orientation (Page) to ai_plugin_liveblend */
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
        + ' /internalName (ai_plugin_liveblend)'
        + ' /localizedName [ 12 e38396e383ace383b3e38389 ]'
        + ' /isOpen 0'
        + ' /isOn 1'
        + ' /hasDialog 1'
        + ' /showDialog 0'
        + ' /parameterCount 3'
        + ' /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 15 e382aae38397e382b7e383a7e383b3 ] /value 5 }'
        + ' /parameter-2 { /key 1937007984 /showInPalette 4294967295 /type (integer) /value ' + blendSteps + ' }'
        + ' /parameter-3 { /key 1919906913 /showInPalette 4294967295 /type (enumerated) /name [ 12 e59e82e79bb4e696b9e59091 ] /value 0 }'
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