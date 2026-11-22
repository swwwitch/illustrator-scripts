#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SendToBgLayer
Meta-Default-language=ja
Name=選択しているオブジェクトを新規レイヤーに移動し、そのレイヤーを最背面に移動してロック
Name[en]=Move the selection to a "bg" layer
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1
Release-Date=2024-06-24
Tags=layer,レイヤー
X-com.swwwitch-Script-Name=SendToBgLayer
X-com.swwwitch-Updated=2024-06-25
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SendToBgLayer.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SendToBgLayer.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/nf7c1e8a0f0c7
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/SendToBgLayer.jsx
Description-BEGIN
[ja]
選択したオブジェクトを、重ね順を維持したまま「bg」レイヤーへ移動して最背面に配置します。
「bg」レイヤーは自動的に作成され、処理後にロックされます。
[en]
Moves the selected objects to a "bg" layer, preserving their stacking order, and sends that layer to the back.
The "bg" layer is created automatically and locked once the move is done.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var TARGET_LAYER_NAME = "bg";  /* 移動先のレイヤー名 / name of the destination layer */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 名前でレイヤーを探し、無ければ作る
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 見つけた、または作ったレイヤー
     */
    function getOrCreateLayer(doc, layerName) {
        try {
            /* 見つからないと例外になる / getByName throws when the layer does not exist */
            return doc.layers.getByName(layerName);
        } catch (e) {
            var newLayer = doc.layers.add();
            newLayer.name = layerName;
            return newLayer;
        }
    }

    /**
     * 選択オブジェクトを重ね順を保ったまま背景レイヤーへ移し、レイヤーを最背面にしてロックする
     * @returns {void}
     */
    function main() {
        var activeDoc = app.activeDocument;
        var originalLayer = activeDoc.activeLayer; /* 元のアクティブレイヤー / Remember the active layer */

        /* 表示状態を控えてから、表示・ロック解除 / Note the visibility, then show and unlock */
        var bgLayer = getOrCreateLayer(activeDoc, TARGET_LAYER_NAME);
        var wasHidden = !bgLayer.visible;
        if (wasHidden) bgLayer.visible = true;
        bgLayer.locked = false;

        /* 選択オブジェクトを重ね順を保ったまま移動 / Move the selection, keeping its stacking order */
        var selectedItems = activeDoc.selection;
        for (var i = 0; i < selectedItems.length; i++) {
            try {
                selectedItems[i].move(bgLayer, ElementPlacement.PLACEATEND);
            } catch (e) {
                /* 移動できないもの（ロックされた親の中など）は飛ばす / Skip items that cannot move */
            }
        }

        /* レイヤーを最背面へ、表示状態を戻してロック / Send to back, restore visibility and lock */
        bgLayer.zOrder(ZOrderMethod.SENDTOBACK);
        if (wasHidden) bgLayer.visible = false;
        bgLayer.locked = true;

        activeDoc.activeLayer = originalLayer;
    }

    main();

})();
