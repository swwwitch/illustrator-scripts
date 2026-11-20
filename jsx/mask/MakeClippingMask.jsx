#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.MakeClippingMask
Meta-Default-language=ja
Name=クリッピングマスクの作成と解除を切り替える
Name[en]=Create or release a clipping mask
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.1
Release-Date=2024-11-18
Tags=mask,マスク
X-com.swwwitch-Script-Name=MakeClippingMask
X-com.swwwitch-Updated=2026-09-30
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeClippingMask.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeClippingMask.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/MakeClippingMask.jsx
Description-BEGIN
[ja]
選択オブジェクトの状態に応じて、クリッピングマスクの作成と解除を切り替えて実行します。
すでにマスクが設定されていれば解除し、配置画像やパスの選択からは新たにマスクを作成します。
[en]
Creates or releases a clipping mask, depending on what is selected.
An existing mask is released, while a selection of placed images or paths produces a new one.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

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
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 判定 / Checks
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
     * 配列の要素がすべてパスか
     * @param {PageItem[]} items - 判定する配列
     * @returns {boolean} すべて PathItem なら true
     */
    function isAllPathItems(items) {
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename !== "PathItem") return false;
        }
        return true;
    }

    /**
     * 最前面のパス（zOrderPosition が最大）を返す
     * @param {PageItem[]} items - 候補
     * @returns {PathItem|null} 最前面のパス。無ければ null
     */
    function getFrontmostPath(items) {
        var frontmostPath = null;
        var highestZ = -1;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "PathItem" && item.zOrderPosition > highestZ) {
                highestZ = item.zOrderPosition;
                frontmostPath = item;
            }
        }
        return frontmostPath;
    }

    // =========================================
    // マスクの作成・解除 / Make and release masks
    // =========================================

    /**
     * レイヤーにグループを作り、中身とマスクのパスを入れてクリップする（パスが最前面）
     * @param {Layer} targetLayer - グループを作るレイヤー
     * @param {PathItem} maskPath - マスクになるパス
     * @param {PageItem} contentItem - マスクされるアイテム
     * @returns {GroupItem} クリップグループ
     */
    function buildClipGroup(targetLayer, maskPath, contentItem) {
        var clipGroup = targetLayer.groupItems.add();
        contentItem.moveToBeginning(clipGroup);
        maskPath.moveToBeginning(clipGroup);
        clipGroup.clipped = true;
        return clipGroup;
    }

    /**
     * 画像と同じ大きさの長方形でマスクする（ロック・非表示・テンプレートのレイヤーは一時的に解除）
     * @param {PlacedItem|RasterItem} imageItem - 対象の画像
     * @returns {PathItem} マスクの長方形
     */
    function createClippingMask(imageItem) {
        var targetLayer = imageItem.layer;
        var wasLocked = targetLayer.locked;
        var wasVisible = targetLayer.visible;
        var wasTemplate = targetLayer.isTemplate;

        if (wasLocked) targetLayer.locked = false;
        if (!wasVisible) targetLayer.visible = true;
        if (wasTemplate) targetLayer.isTemplate = false;

        var clippingRect = targetLayer.pathItems.rectangle(
            imageItem.top,
            imageItem.left,
            imageItem.width,
            imageItem.height
        );
        clippingRect.stroked = false;
        clippingRect.filled = false;
        buildClipGroup(targetLayer, clippingRect, imageItem);

        if (wasLocked) targetLayer.locked = true;
        if (!wasVisible) targetLayer.visible = false;
        if (wasTemplate) targetLayer.isTemplate = true;

        return clippingRect;
    }

    /**
     * 選択したパスで画像をマスクする（パスは画像のレイヤーへ移す）
     * @param {PlacedItem|RasterItem} imageItem - 対象の画像
     * @param {PathItem} pathItem - マスクになるパス
     * @returns {PathItem} マスクのパス
     */
    function createMaskWithPath(imageItem, pathItem) {
        var targetLayer = imageItem.layer;
        if (pathItem.layer != targetLayer) {
            pathItem.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
        }
        buildClipGroup(targetLayer, pathItem, imageItem);
        return pathItem;
    }

    /**
     * 選択中のパスのうち最前面のものをマスクにして、ほかのパスをクリップする
     * @param {Document} doc - 対象ドキュメント
     * @returns {PathItem|null} マスクのパス。無ければ null
     */
    function clipPathsWithFrontmost(doc) {
        var frontmostPath = getFrontmostPath(doc.selection);
        if (frontmostPath === null) return null;

        var clipGroup = doc.groupItems.add();
        frontmostPath.moveToBeginning(clipGroup);
        frontmostPath.clipping = true;

        for (var i = 0; i < doc.selection.length; i++) {
            var pathItem = doc.selection[i];
            if (pathItem !== frontmostPath) {
                pathItem.moveToEnd(clipGroup);
            }
        }

        clipGroup.clipped = true;
        return frontmostPath;
    }

    /**
     * グループを解除する（中身を親へ出してグループを削除）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {void}
     */
    function ungroupGroupItem(groupItem) {
        var parentContainer = groupItem.parent;
        while (groupItem.pageItems.length > 0) {
            groupItem.pageItems[0].moveToBeginning(parentContainer);
        }
        groupItem.remove();
    }

    /**
     * クリッピングマスクを解除し、マスクのパスを削除してグループを解く
     * @param {GroupItem} groupItem - クリップグループ
     * @returns {void}
     */
    function releaseClippingMask(groupItem) {
        var clippingPath = null;
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            if (groupItem.pageItems[i].clipping) {
                clippingPath = groupItem.pageItems[i];
                break;
            }
        }
        if (clippingPath !== null) {
            groupItem.clipped = false;
            clippingPath.remove();
        }
        ungroupGroupItem(groupItem);
    }

    /**
     * 作成したマスクのクリップグループだけを選択する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} clippingMasks - マスクのパス
     * @returns {void}
     */
    function selectClipGroups(doc, clippingMasks) {
        doc.selection = null;
        for (var i = 0; i < clippingMasks.length; i++) {
            clippingMasks[i].parent.selected = true;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択に応じてクリッピングマスクを解除・作成する
     * （処理中に選択が変わるため、選択は毎回 doc.selection から読み直す）
     * @returns {void}
     */
    function main() {
        var doc = app.activeDocument;
        /* 文字ツールで文字を選択しているときは TextRange が返り、length は文字数になる / With characters selected by the Type tool, selection is a TextRange whose length is the character count */
        if (doc.selection.typename === "TextRange") return;
        var createdMasks = [];
        var i;

        /* 選択中のクリップグループを解除 / Release selected clipping groups */
        for (i = 0; i < doc.selection.length; i++) {
            var selectedObject = doc.selection[i];
            if (selectedObject.typename === "GroupItem" && selectedObject.clipped) {
                releaseClippingMask(selectedObject);
            }
        }

        if (doc.selection.length === 2) {
            /* 画像＋パス：パスでマスク / Image + path: mask with the path */
            var imageItem = null;
            var pathItem = null;
            for (i = 0; i < 2; i++) {
                if (isImageItem(doc.selection[i])) {
                    imageItem = doc.selection[i];
                } else if (doc.selection[i].typename === "PathItem") {
                    pathItem = doc.selection[i];
                }
            }
            if (imageItem !== null && pathItem !== null) {
                createdMasks.push(createMaskWithPath(imageItem, pathItem));
            }

        } else if (isAllPathItems(doc.selection) && doc.selection.length >= 2) {
            /* パスのみ：最前面のパスでマスク / Paths only: mask with the frontmost path */
            var frontmostPath = clipPathsWithFrontmost(doc);
            if (frontmostPath !== null) createdMasks.push(frontmostPath);

        } else {
            /* 画像単体：画像と同じ大きさの長方形でマスク / Single images: mask with a same-size rectangle */
            for (i = 0; i < doc.selection.length; i++) {
                if (isImageItem(doc.selection[i])) {
                    createdMasks.push(createClippingMask(doc.selection[i]));
                }
            }
        }

        selectClipGroups(doc, createdMasks);
    }

    main();

})();
