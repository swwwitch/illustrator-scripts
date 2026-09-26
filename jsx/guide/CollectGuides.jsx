#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のレイヤーやサブレイヤーに散在するガイドを、1つのレイヤー（既定は「// guide」）へ集約します。
非表示・ロックされたレイヤーのガイドも一時解除して対象にし、処理後に元の状態へ戻します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CollectGuides.md

### Overview

Collects guides scattered across layers and sublayers into a single layer ("// guide" by default).
Hidden and locked layers are unlocked temporarily so their guides are included, then restored afterwards.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CollectGuides.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CollectGuides";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CollectGuides.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CollectGuides.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ガイドを集めるレイヤー名（プロジェクトごとに変更可） / Layer that receives the guides; change per project */
    var TARGET_GUIDE_LAYER_NAME = "// guide";

    /* 処理後のガイドの表示（"show" / "hide" / null = そのまま） / Guide visibility afterwards (null keeps it) */
    var RESTORE_GUIDES_VISIBILITY = null;

    /* 処理後のガイドのロック（"lock" / "unlock" / null = そのまま） / Guide lock afterwards (null keeps it) */
    var RESTORE_GUIDES_LOCK = null;

    /* true: 処理の前後で［表示］＞［プレビュー］を切り替えて再描画を抑える（2回呼んで元の表示に戻す）
       Toggle View > Preview before and after to suppress redraws; the second call restores the view */
    var USE_PREVIEW_TOGGLE_WRAPPER = true;

    // =========================================
    // メニューコマンド / Menu commands
    // =========================================

    /**
     * メニューコマンドを実行する（コマンド名が無い環境でも止めない）
     * @param {string} commandName - executeMenuCommand に渡すコマンド名
     * @returns {void}
     */
    function runMenuCommandSafely(commandName) {
        /* 存在しないコマンド名は例外になるので握る / Unknown command names throw */
        try {
            app.executeMenuCommand(commandName);
        } catch (e) {}
    }

    /**
     * 処理後のガイドの表示・ロックを、ユーザー設定に従って切り替える
     * @returns {void}
     */
    function restoreGuideSettings() {
        if (RESTORE_GUIDES_VISIBILITY === "show") runMenuCommandSafely("showGuides");
        else if (RESTORE_GUIDES_VISIBILITY === "hide") runMenuCommandSafely("hideGuides");

        if (RESTORE_GUIDES_LOCK === "lock") runMenuCommandSafely("lockGuides");
        else if (RESTORE_GUIDES_LOCK === "unlock") runMenuCommandSafely("unlockGuides");
    }

    // =========================================
    // ガイドの移動 / Moving guides
    // =========================================

    /**
     * レイヤー（サブレイヤーを含む）にガイドがあるかを返す。集約先のレイヤーは対象外
     * @param {Layer} sourceLayer - 調べるレイヤー
     * @returns {boolean} ガイドがあれば true
     */
    function layerHasGuides(sourceLayer) {
        if (!sourceLayer || sourceLayer.name === TARGET_GUIDE_LAYER_NAME) return false;
        for (var i = 0; i < sourceLayer.pageItems.length; i++) {
            if (sourceLayer.pageItems[i].guides === true) return true;
        }
        for (var j = 0; j < sourceLayer.layers.length; j++) {
            if (layerHasGuides(sourceLayer.layers[j])) return true;
        }
        return false;
    }

    /**
     * ガイドを集約先のレイヤーへ移す（ロック・非表示は一時的に外して元に戻す）
     * @param {PageItem} guideItem - 移すガイド
     * @param {Layer} guideLayer - 集約先のレイヤー
     * @returns {void}
     */
    function moveGuideItem(guideItem, guideLayer) {
        var itemWasLocked = guideItem.locked;
        var itemWasHidden = guideItem.hidden;

        /* 移動のため一時的に解除 / Temporarily unlock and show the item to allow moving */
        if (itemWasLocked) guideItem.locked = false;
        if (itemWasHidden) guideItem.hidden = false;

        guideItem.move(guideLayer, ElementPlacement.PLACEATBEGINNING);

        guideItem.locked = itemWasLocked;
        guideItem.hidden = itemWasHidden;
    }

    /**
     * レイヤーとサブレイヤーを再帰的に走査し、ガイドを集約先のレイヤーへ移す
     * @param {Layer} sourceLayer - 走査するレイヤー
     * @param {Layer} guideLayer - 集約先のレイヤー
     * @returns {void}
     */
    function moveGuidesInLayer(sourceLayer, guideLayer) {
        if (sourceLayer.name === TARGET_GUIDE_LAYER_NAME) return;

        var layerWasLocked = sourceLayer.locked;
        var layerWasVisible = sourceLayer.visible;

        /* 一時的にロック解除＆表示 / Temporarily unlock and show the layer */
        if (layerWasLocked) sourceLayer.locked = false;
        if (!layerWasVisible) sourceLayer.visible = true;

        /* 移すと添字がずれるので末尾から / Walk backwards; moving shifts later indexes */
        for (var i = sourceLayer.pageItems.length - 1; i >= 0; i--) {
            var pageItem = sourceLayer.pageItems[i];
            if (pageItem.guides === true) moveGuideItem(pageItem, guideLayer);
        }

        for (var j = 0; j < sourceLayer.layers.length; j++) {
            moveGuidesInLayer(sourceLayer.layers[j], guideLayer);
        }

        sourceLayer.locked = layerWasLocked;
        sourceLayer.visible = layerWasVisible;
    }

    /**
     * 集約先のレイヤーを返す（無ければ作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 集約先のレイヤー
     */
    function getOrCreateGuideLayer(doc) {
        /* getByName は見つからないと例外 / getByName throws when the layer is missing */
        try {
            return doc.layers.getByName(TARGET_GUIDE_LAYER_NAME);
        } catch (e) {
            var guideLayer = doc.layers.add();
            guideLayer.name = TARGET_GUIDE_LAYER_NAME;
            return guideLayer;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * レイヤーに散在するガイドを、1つのレイヤーへ集める
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        /* ガイドを一時的に表示＆ロック解除（非表示・ロック中でも移動できるように） / Make guides visible and unlocked */
        runMenuCommandSafely("showGuides");
        runMenuCommandSafely("unlockGuides");

        if (USE_PREVIEW_TOGGLE_WRAPPER) runMenuCommandSafely("preview");

        var guideLayer = getOrCreateGuideLayer(doc);

        /* ガイドを含むトップレベルのレイヤーだけ処理 / Only process top-level layers that contain guides */
        for (var i = 0; i < doc.layers.length; i++) {
            if (layerHasGuides(doc.layers[i])) moveGuidesInLayer(doc.layers[i], guideLayer);
        }

        restoreGuideSettings();

        /* 2回目の切り替えで元の表示に戻す / The second toggle restores the view */
        if (USE_PREVIEW_TOGGLE_WRAPPER) runMenuCommandSafely("preview");

        app.redraw();
    }

    main();

})();
