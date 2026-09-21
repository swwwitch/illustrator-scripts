#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、一緒に選んだ既存のグループへ加えます。グループの効果やクリッピングマスクはそのまま残ります。
グループが無いか複数あるときは、グループを1段解除してから1つのグループにまとめ直します（クリップグループは解除しません）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddToGroup.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n36fbd4162721

### Overview

Adds the selected objects to the existing group selected with them, keeping the group's effects and clipping mask.
With no group or several groups, the groups are released one level and everything is regrouped as one (clipping groups are kept intact).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddToGroup.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddToGroup";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddToGroup.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddToGroup.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n36fbd4162721"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 選択 / Selection
    // =========================================

    /**
     * 処理対象の選択オブジェクトを返す
     * @returns {PageItem[]|null} 選択オブジェクト（前面→背面の順）。ドキュメントがない・未選択・文字の選択中は null
     */
    function getValidSelection() {
        if (!app.documents.length) return null;
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る
           A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || !selectedItems.length || selectedItems.typename === "TextRange") return null;
        return selectedItems;
    }

    /**
     * 選択の中でグループが並んでいる位置を集める
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @returns {number[]} GroupItem の添字
     */
    function findGroupIndexes(selectedItems) {
        var groupIndexes = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem") groupIndexes.push(i);
        }
        return groupIndexes;
    }

    // =========================================
    // グループ操作 / Group operations
    // =========================================

    /**
     * クリップグループのマスクを探す
     * パスは clipping、複合パスは先頭のサブパスの clipping で見分ける。テキストのマスクにはフラグが無いので、最前面の項目をマスクとみなす
     * @param {GroupItem} clipGroup - クリップグループ（オブジェクトを入れる前の状態）
     * @returns {PageItem|null} マスク（中身が空なら null）
     */
    function findMaskItem(clipGroup) {
        for (var i = 0; i < clipGroup.pageItems.length; i++) {
            var childItem = clipGroup.pageItems[i];
            if (childItem.clipping) return childItem;
            if (childItem.typename === "CompoundPathItem" && childItem.pathItems.length > 0 && childItem.pathItems[0].clipping) return childItem;
        }
        /* マスクはクリップグループの最前面にある / The mask sits at the top of the clipping group */
        return (clipGroup.pageItems.length > 0) ? clipGroup.pageItems[0] : null;
    }

    /**
     * グループ以外の選択オブジェクトを、重ね順を保ったままグループへ移す
     * グループより前面のものはグループ内の最前面へ、背面のものは最背面へ入れる
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @param {number} groupIndex - selectedItems の中でのグループの添字
     * @returns {void}
     */
    function moveItemsIntoGroup(selectedItems, groupIndex) {
        var targetGroup = selectedItems[groupIndex];
        /* 前面側はグループに近いものから最前面へ入れる / Front side: nearest first, each to the top */
        for (var i = groupIndex - 1; i >= 0; i--) {
            selectedItems[i].move(targetGroup, ElementPlacement.PLACEATBEGINNING);
        }
        /* 背面側はグループに近いものから最背面へ入れる / Back side: nearest first, each to the bottom */
        for (var j = groupIndex + 1; j < selectedItems.length; j++) {
            selectedItems[j].move(targetGroup, ElementPlacement.PLACEATEND);
        }
    }

    /**
     * 既存のグループへ、ほかの選択オブジェクトを加える
     * グループは解除しないので、効果・名前・クリッピングマスク・ロックや非表示の中身はそのまま残る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @param {number} groupIndex - selectedItems の中でのグループの添字
     * @returns {void}
     */
    function addItemsToGroup(doc, selectedItems, groupIndex) {
        var targetGroup = selectedItems[groupIndex];
        /* 前面側のオブジェクトはマスクより上に入るので、入れる前にマスクを控えて最前面へ戻す
           Front-side objects land above the mask; remember the mask first and bring it back to the top */
        var maskItem = targetGroup.clipped ? findMaskItem(targetGroup) : null;
        moveItemsIntoGroup(selectedItems, groupIndex);
        if (maskItem) maskItem.zOrder(ZOrderMethod.BRINGTOFRONT);
        doc.selection = [targetGroup];
    }

    /**
     * グループを1段解除してから、選択全体を1つのグループにまとめ直す
     * クリップグループは解除するとマスクが外れるので、解除せずにそのまま入れる
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @returns {void}
     */
    function regroupSelection(doc, selectedItems) {
        var releasedItems = [];
        var keptClipGroups = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem" && selectedItems[i].clipped) {
                keptClipGroups.push(selectedItems[i]);
            } else {
                releasedItems.push(selectedItems[i]);
            }
        }

        if (releasedItems.length > 0) {
            doc.selection = releasedItems;
            app.executeMenuCommand("ungroup");
            /* 解除後の選択は、グループの中身とグループ以外のオブジェクト / Afterwards the selection holds the released contents and the other objects */
            releasedItems = doc.selection;
        }

        var itemsToGroup = [];
        for (var j = 0; j < releasedItems.length; j++) itemsToGroup.push(releasedItems[j]);
        for (var k = 0; k < keptClipGroups.length; k++) itemsToGroup.push(keptClipGroups[k]);
        doc.selection = itemsToGroup;
        app.executeMenuCommand("group");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * グループが1つならほかのオブジェクトをそこへ加え、それ以外は1つのグループにまとめ直す
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        if (!selectedItems || selectedItems.length < 2) return;

        var doc = app.activeDocument;
        var groupIndexes = findGroupIndexes(selectedItems);

        /* グループが1つ：ほかのオブジェクトをそのグループへ入れる / One group: move the other objects into it */
        if (groupIndexes.length === 1) {
            addItemsToGroup(doc, selectedItems, groupIndexes[0]);
            return;
        }

        /* グループが無いか複数：1つのグループにまとめ直す / No group or several groups: regroup everything as one */
        regroupSelection(doc, selectedItems);
    }

    main();

})();
