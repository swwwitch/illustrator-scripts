#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.SimplifyGroups
Meta-Default-language=ja
Name=選択したグループ内のサブグループを解除して、グループ構造を簡素化
Name[en]=Ungroup nested subgroups recursively
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.2
Release-Date=2025-07-07
Tags=group,グループ
X-com.swwwitch-Script-Name=SimplifyGroups
X-com.swwwitch-Updated=2026-11-23
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SimplifyGroups.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SimplifyGroups.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/n45797beb72bb
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SimplifyGroups.jsx
Description-BEGIN
[ja]
選択したグループの中にあるサブグループを再帰的に解除し、最外層のグループとクリップグループだけを残します。
グループ以外のオブジェクトも選んでいるときは、それらをグループに取り込みます（グループが複数あるときは、全体を新しいグループにまとめてから処理します）。
[en]
Recursively ungroups the subgroups inside the selection, leaving only the outermost group and clipping groups.
Other selected objects are moved into the group; with several groups, everything is first combined into a new group.
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
     * グループ内のサブグループを再帰的に解除する
     * parentGroup 自身と、ロック・非表示のサブグループは残す。クリップグループは中だけ解除して残す
     * @param {Document} doc - 対象ドキュメント
     * @param {GroupItem} parentGroup - 解除せずに残すグループ
     * @returns {void}
     */
    function ungroupSubGroups(doc, parentGroup) {
        /* 解除すると後ろの添字がずれるので末尾から処理 / Walk backwards; ungrouping shifts later indexes */
        for (var i = parentGroup.pageItems.length - 1; i >= 0; i--) {
            var childItem = parentGroup.pageItems[i];
            if (childItem.typename !== "GroupItem") continue;
            /* ロック・非表示のグループは選択できないので残す / Locked or hidden groups cannot be selected, so leave them */
            if (childItem.locked || childItem.hidden) continue;
            ungroupSubGroups(doc, childItem);
            /* クリップグループは解除するとマスクが外れるので残す / Ungrouping a clipping group releases its mask, so keep it */
            if (childItem.clipped) continue;
            doc.selection = [childItem];
            app.executeMenuCommand("ungroup");
        }
    }

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

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択の組み合わせに応じて、グループへの取り込みとサブグループの解除を行う
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        if (!selectedItems) return;

        var doc = app.activeDocument;
        var groupIndexes = findGroupIndexes(selectedItems);
        if (groupIndexes.length === 0) return;

        /* グループだけ：それぞれのサブグループを解除 / Groups only: flatten each group */
        if (groupIndexes.length === selectedItems.length) {
            for (var i = 0; i < selectedItems.length; i++) {
                ungroupSubGroups(doc, selectedItems[i]);
            }
            doc.selection = selectedItems;
            return;
        }

        /* グループ1つと非グループ：既存のグループへ取り込む / One group plus other objects: move them into it */
        if (groupIndexes.length === 1) {
            var targetGroup = selectedItems[groupIndexes[0]];
            /* 前面側のオブジェクトはマスクより上に入るので、入れる前にマスクを控えて最前面へ戻す
               Front-side objects land above the mask; remember the mask first and bring it back to the top */
            var maskItem = targetGroup.clipped ? findMaskItem(targetGroup) : null;
            moveItemsIntoGroup(selectedItems, groupIndexes[0]);
            if (maskItem) maskItem.zOrder(ZOrderMethod.BRINGTOFRONT);
            ungroupSubGroups(doc, targetGroup);
            doc.selection = [targetGroup];
            return;
        }

        /* 複数のグループと非グループ：まとめてグループ化。選んだグループはサブグループとして解除される（クリップグループは残る）
           Several groups plus other objects: group them all first; the selected groups become subgroups and are ungrouped (clipping groups stay) */
        app.executeMenuCommand("group");
        var groupedItems = doc.selection;
        if (groupedItems.length === 1 && groupedItems[0].typename === "GroupItem") {
            ungroupSubGroups(doc, groupedItems[0]);
            doc.selection = [groupedItems[0]];
        }
    }

    main();

})();
