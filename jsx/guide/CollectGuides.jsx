#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
Script: CollectGuides.jsx

目的 / Purpose:
- 複数レイヤー/サブレイヤーに散在するガイドを新規レイヤー（デフォルト: "// guide"）に集約します。
- 非表示/ロックされたレイヤーやガイドも対象。処理中は一時解除し、終了後に元の状態へ復元します。

概要 / Overview:
- 全レイヤー（サブレイヤー含む）を再帰的に走査し、ガイド（pageItem.guides === true）をターゲットレイヤーへ移動。
- レイヤー/アイテムの locked・visible・hidden を退避→復元。
- 大規模ドキュメント向けの最適化（対象限定スキャン / Redraw抑制）を搭載。

主な機能 / Features:
- ガイドの一括収集 / Collect all guides into a single layer
- 非表示/ロック状態のガイド・レイヤーにも対応 / Handles hidden/locked guides and layers
- ターゲットレイヤー名を定数で変更可能 / Target layer name configurable
- 実行後のガイド表示/ロック状態の復帰オプション / Optional restore of global guide show/lock
- Redraw抑制オプション / Optional redraw suppression

使い方 / How to use:
- ドキュメントを開いて実行（対象はアクティブドキュメント）。
- 既にターゲットレイヤーが存在する場合は再利用。

仕様 / Specs:
- グローバルの「表示 > ガイドを表示」「表示 > ガイドをロック解除」を一時的に実行（showGuides / unlockGuides）。
- すべてのレイヤーとサブレイヤーを再帰的に走査し、ガイド（pageItem.guides === true）を移動。
- レイヤーとアイテムの locked / visible / hidden 状態は退避→復元。

更新履歴 / Changelog:
- v1.0 (20250816): 初版。ガイド集約、再帰走査、ロック/可視状態の復元を実装。
- v1.1 (20250816): 対象限定スキャンを追加。事前に「ガイドを含むレイヤーのみ」絞込んで処理。
- v1.2 (20250816): 実行後に「ガイド表示/ロック」を任意状態へ復帰できる設定を追加（自動検出はAPI制約により非対応）。
- v1.3 (20250816): ターゲットレイヤー名を定数 `TARGET_GUIDE_LAYER_NAME` で切替可能に。
- v1.4 (20250816): Redraw抑制オプションを追加（処理前後で Preview トグル／最後に app.redraw() 1回）。
*/

// ===== 設定 / Settings =====
// 実行後にグローバルの「ガイド表示/ロック」状態を所望の状態へ復帰するか
// Restore global guide visibility/lock after processing.
// null の場合は復帰しません（現状維持）。
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