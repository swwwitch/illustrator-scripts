#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ClipWithSquare
Meta-Default-language=ja
Name=画像を中心基準の正方形でクリップ
Name[en]=Clip an image with a centered square
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.1
Release-Date=2023-11-26
Tags=mask,マスク
X-com.swwwitch-Script-Name=ClipWithSquare
X-com.swwwitch-Updated=2026-09-30
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipWithSquare.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipWithSquare.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/single-function/ClipWithSquare.jsx
Description-BEGIN
[ja]
選択したオブジェクトを、正方形のクリッピングマスクで切り抜きます。
[en]
Clips the selected objects with a square clipping mask.
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
    var WORK_LAYER_NAME = "_clip_work";  /* 画像のレイヤーが編集できないときに使う作業レイヤー / work layer used when the image's layer is not editable */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 画像（配置画像／埋め込み画像）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} 画像なら true
     */
    function isImageItem(item) {
        return item.typename === "PlacedItem" || item.typename === "RasterItem";
    }

    /**
     * 作業レイヤーを取得、無ければ作成する（既存がロック・テンプレートでも常に編集可能にする）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 作業レイヤー
     */
    function getOrCreateWorkLayer(doc) {
        var workLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === WORK_LAYER_NAME) {
                workLayer = doc.layers[i];
                break;
            }
        }
        if (!workLayer) {
            workLayer = doc.layers.add();
            workLayer.name = WORK_LAYER_NAME;
        }
        workLayer.locked = false;
        workLayer.visible = true;
        workLayer.isTemplate = false;
        return workLayer;
    }

    /**
     * 画像の短辺に合わせた中央の正方形でクリッピングマスクを作り、選択する
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacedItem|RasterItem} image - 対象の画像
     * @returns {void}
     */
    function clipImageWithSquare(doc, image) {
        var bounds = image.visibleBounds;
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        var sideLength = Math.min(width, height);
        var centerX = bounds[0] + width / 2;
        var centerY = bounds[1] - height / 2;

        /* 画像のレイヤーが編集できなければ作業レイヤーを使う / Use the work layer when the image's layer is not editable */
        var parentLayer = image.layer;
        var targetLayer = (parentLayer.locked || parentLayer.isTemplate) ? getOrCreateWorkLayer(doc) : parentLayer;

        var square = targetLayer.pathItems.rectangle(centerY + sideLength / 2, centerX - sideLength / 2, sideLength, sideLength);
        var clipGroup = targetLayer.groupItems.add();

        image.moveToBeginning(clipGroup);
        square.moveToBeginning(clipGroup);
        clipGroup.clipped = true;
        clipGroup.selected = true;
    }

    /**
     * クリップグループなら中の画像だけを取り出す（マスクのパスなど画像以外は削除）
     * @param {GroupItem} clipGroup - クリップグループ
     * @returns {PageItem[]} 取り出した画像
     */
    function extractImagesFromClipGroup(clipGroup) {
        clipGroup.clipped = false;
        var images = [];
        for (var i = clipGroup.pageItems.length - 1; i >= 0; i--) {
            var pageItem = clipGroup.pageItems[i];
            if (isImageItem(pageItem)) {
                images.push(pageItem);
            } else {
                pageItem.remove();
            }
        }
        return images;
    }

    /**
     * 選択した画像（クリップグループ内の画像を含む）を正方形で切り抜く
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、length は文字数になる / With characters selected by the Type tool, selection is a TextRange whose length is the character count */
        if (selectedItems.typename === "TextRange" || !selectedItems.length) return;

        /* 作ったグループだけが選択に残るよう、先に選択を解除 / Clear the selection first so only new groups end up selected */
        doc.selection = null;

        for (var i = 0; i < selectedItems.length; i++) {
            var item = selectedItems[i];
            if (item.typename === "GroupItem" && item.clipped) {
                var images = extractImagesFromClipGroup(item);
                for (var j = 0; j < images.length; j++) {
                    clipImageWithSquare(doc, images[j]);
                }
            } else if (isImageItem(item)) {
                clipImageWithSquare(doc, item);
            }
        }
    }

    main();

})();
