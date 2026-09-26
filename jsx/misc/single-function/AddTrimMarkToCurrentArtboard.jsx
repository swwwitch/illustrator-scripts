#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

現在のアートボードを対象に、トンボを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMarkToCurrentArtboard.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n40e3e39cf9f2

### Overview

Creates trim marks for the current artboard.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMarkToCurrentArtboard.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddTrimMarkToCurrentArtboard"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-02-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMarkToCurrentArtboard.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMarkToCurrentArtboard.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n40e3e39cf9f2"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* トンボを作るレイヤー名（無ければ作る） / Layer that receives the trim marks; created when missing */
    var TRIM_LAYER_NAME = "トンボ";

    // =========================================
    // トンボ / Trim marks
    // =========================================

    /**
     * トンボ用のレイヤーを返す（無ければ作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} トンボ用のレイヤー
     */
    function getOrCreateTrimLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === TRIM_LAYER_NAME) return doc.layers[i];
        }
        var trimLayer = doc.layers.add();
        trimLayer.name = TRIM_LAYER_NAME;
        return trimLayer;
    }

    /**
     * アートボードと同じ大きさの矩形を作り、そこからトンボを作成して矩形をガイドにする
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} trimLayer - トンボ用のレイヤー（ロック解除済み）
     * @param {Artboard} artboard - 対象のアートボード
     * @returns {void}
     */
    function createTrimMarks(doc, trimLayer, artboard) {
        /* 日本式トンボをONに設定 / Enable Japanese-style trim marks */
        app.preferences.setBooleanPreference("cropMarkStyle", 1);

        /* 「トンボ」レイヤー上にアートボード矩形を作成 / Create the artboard rectangle on the trim layer */
        var artboardRect = artboard.artboardRect; /* [左, 上, 右, 下] / [left, top, right, bottom] */
        var artboardFrame = trimLayer.pathItems.rectangle(artboardRect[1], artboardRect[0], artboardRect[2] - artboardRect[0], artboardRect[1] - artboardRect[3]);
        artboardFrame.filled = false;
        artboardFrame.stroked = false;

        /* 矩形を選択してトリムマークを作成 / Select the rectangle and create trim marks */
        doc.selection = [artboardFrame];
        app.executeMenuCommand("TrimMark v25");

        /* 矩形はガイドにする / Turn the rectangle into a guide */
        artboardFrame.guides = true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * アクティブなアートボードにトンボを作成する（「トンボ」レイヤーのロックは一時的に外して戻す）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var trimLayer = getOrCreateTrimLayer(doc);
        var layerWasLocked = trimLayer.locked;
        if (layerWasLocked) trimLayer.locked = false;

        /* 途中で失敗しても選択解除とロックの復元は行う / Always clear the selection and restore the lock */
        try {
            createTrimMarks(doc, trimLayer, doc.artboards[doc.artboards.getActiveArtboardIndex()]);
        } finally {
            doc.selection = null;
            trimLayer.locked = layerWasLocked;
        }
    }

    main();

})();
