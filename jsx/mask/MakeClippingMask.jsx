#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの状態に応じて、クリッピングマスクの作成と解除を切り替えて実行します。
すでにマスクが設定されていれば解除し、配置画像やパスの選択からは新たにマスクを作成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeClippingMask.md

### Overview

Creates or releases a clipping mask, depending on what is selected.
An existing mask is released, while a selection of placed images or paths produces a new one.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeClippingMask.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MakeClippingMask";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-11-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeClippingMask.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeClippingMask.md"; /* README (English) */

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
