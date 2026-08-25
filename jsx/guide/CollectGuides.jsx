#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のレイヤーやサブレイヤーに散在するガイドを、1つのレイヤー（既定は「// guide」）へ集約します。
非表示・ロックされたレイヤーのガイドも一時解除して対象にし、処理後に元の状態へ戻します。

詳細は README を参照してください。

### Overview

Collects guides scattered across layers and sublayers into a single layer ("// guide" by default).
Hidden and locked layers are unlocked temporarily so their guides are included, then restored afterwards.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CollectGuides";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-08-16";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CollectGuides.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CollectGuides.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

var RESTORE_GUIDES_VISIBILITY /* 'show' | 'hide' | null */ = null;
var RESTORE_GUIDES_LOCK /* 'lock' | 'unlock' | null */ = null;
// ============================

// ターゲットレイヤー名 / Target layer name
var TARGET_GUIDE_LAYER_NAME = "// guide"; // プロジェクト毎に変更可

// Redraw抑制（任意）/ Redraw suppression (optional)
// true: 処理の前後で View > Preview をトグル（2回呼びで元の表示状態に戻す）
var USE_PREVIEW_TOGGLE_WRAPPER = true;

function _safeTogglePreview() {
    try {
        app.executeMenuCommand('preview');
    } catch (e) {}
}

// このレイヤー配下にガイドが存在するかを事前判定 / Check if a layer (including sublayers) contains any guides
function layerHasGuides(lyr) {
    if (!lyr || lyr.name === TARGET_GUIDE_LAYER_NAME) return false;
    // Check direct items
    for (var j = 0; j < lyr.pageItems.length; j++) {
        var it;
        try {
            it = lyr.pageItems[j];
        } catch (e) {
            continue;
        }
        try {
            if (it && it.guides === true) return true;
        } catch (e2) {}
    }
    // Recurse into sublayers
    for (var k = 0; k < lyr.layers.length; k++) {
        try {
            if (layerHasGuides(lyr.layers[k])) return true;
        } catch (e3) {}
    }
    return false;
}

// 全レイヤー/サブレイヤーを再帰的に走査し、ガイドを移動 / Iterate through all layers (recursive) and move guides
function moveGuidesInLayer(lyr, guideLayer) {
    if (lyr.name === TARGET_GUIDE_LAYER_NAME) return; // skip target layer itself

    // レイヤー状態の退避（locked/visible） / Save original layer states
    var layerWasLocked = lyr.locked;
    var layerWasVisible = lyr.visible;

    // 一時的にロック解除＆表示 / Temporarily unlock and show the layer
    if (lyr.locked) lyr.locked = false;
    if (!lyr.visible) lyr.visible = true;

    // このレイヤー直下のガイドを移動 / Move guides from this layer
    for (var j = lyr.pageItems.length - 1; j >= 0; j--) {
        var it = lyr.pageItems[j];
        if (it.guides === true) {
            // ガイドアイテムの状態退避（locked/hidden） / Save item states (locked/hidden)
            var itemWasLocked = it.locked;
            var itemWasHidden = it.hidden;

            // 移動のため一時的に解除 / Temporarily unlock/show the item to allow moving
            if (it.locked) it.locked = false;
            if (it.hidden) it.hidden = false;

            it.move(guideLayer, ElementPlacement.PLACEATBEGINNING);

            // アイテム状態を復元 / Restore item states
            it.locked = itemWasLocked;
            it.hidden = itemWasHidden;
        }
    }

    // サブレイヤーを再帰処理 / Recurse into sublayers
    for (var k = 0; k < lyr.layers.length; k++) {
        moveGuidesInLayer(lyr.layers[k], guideLayer);
    }

    // レイヤー状態を復元 / Restore original layer states
    lyr.locked = layerWasLocked;
    lyr.visible = layerWasVisible;
}

function main() {
    var doc = app.activeDocument;
    var guideLayer;

    // ガイドを一時的に表示＆ロック解除（非表示/ロック中でも移動可能に） / Ensure global guides are visible and unlocked
    try {
        app.executeMenuCommand('showGuides');
    } catch (e) {}
    try {
        app.executeMenuCommand('unlockGuides');
    } catch (e) {}

    // --- Redraw抑制（任意）開始 / Begin redraw suppression (optional)
    if (USE_PREVIEW_TOGGLE_WRAPPER) _safeTogglePreview();

    // Find existing "// guide" layer or create a new one
    try {
        guideLayer = doc.layers.getByName(TARGET_GUIDE_LAYER_NAME);
    } catch (e) {
        guideLayer = doc.layers.add();
        guideLayer.name = TARGET_GUIDE_LAYER_NAME;
    }

    // Kick off from top-level layers (対象限定: ガイドを含むレイヤーのみ処理) / Process only layers that contain guides
    for (var i = 0; i < doc.layers.length; i++) {
        var root = doc.layers[i];
        if (layerHasGuides(root)) {
            moveGuidesInLayer(root, guideLayer);
        }
    }

    // --- グローバル設定の復帰（オプション） / Restore global settings (optional)
    try {
        if (RESTORE_GUIDES_VISIBILITY === 'show') app.executeMenuCommand('showGuides');
        else if (RESTORE_GUIDES_VISIBILITY === 'hide') app.executeMenuCommand('hideGuides');
    } catch (eVis) {}

    try {
        if (RESTORE_GUIDES_LOCK === 'lock') app.executeMenuCommand('lockGuides');
        else if (RESTORE_GUIDES_LOCK === 'unlock') app.executeMenuCommand('unlockGuides');
    } catch (eLock) {}

    // --- Redraw抑制（任意）終了 / End redraw suppression (optional)
    if (USE_PREVIEW_TOGGLE_WRAPPER) _safeTogglePreview(); // 2回目で元の表示状態へ戻す

    // 最後に1回だけ再描画 / Single final redraw
    try {
        app.redraw();
    } catch (eRedraw) {}
}

main();
