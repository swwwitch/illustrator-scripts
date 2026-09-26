#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトがグループ内にある場合、親グループを辿って所属レイヤーの直下へ移動します。
重ね順が反転しないよう、選択オブジェクトは逆順に処理します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseFromGroup.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n36fbd4162721

### Overview

Moves the selected objects out of their groups and directly onto the layer that owns them.
They are processed in reverse order so that the stacking order is preserved.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseFromGroup.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReleaseFromGroup";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseFromGroup.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseFromGroup.md"; /* README (English) */
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

    // =========================================
    // グループから出す / Release from group
    // =========================================

    /**
     * いちばん外側の親グループ（レイヤー直下のグループ）を返す
     * @param {PageItem} targetItem - 対象オブジェクト
     * @returns {GroupItem|null} 親グループ（グループの中に無ければ null）
     */
    function findOutermostGroup(targetItem) {
        var outermostGroup = null;
        var currentContainer = targetItem.parent;
        while (currentContainer.typename === "GroupItem") {
            outermostGroup = currentContainer;
            currentContainer = currentContainer.parent;
        }
        return outermostGroup;
    }

    /**
     * グループの中で選んだオブジェクトを、所属レイヤーの最前面へ出す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @returns {void}
     */
    function releaseToLayerTop(doc, selectedItems) {
        /* レイヤーの最前面へは背面側から出す（最後に出した最前面のものが一番上になる）
           Back to front: the frontmost one is moved last and ends up on top */
        for (var i = selectedItems.length - 1; i >= 0; i--) {
            var outermostGroup = findOutermostGroup(selectedItems[i]);
            if (!outermostGroup) continue;
            selectedItems[i].move(outermostGroup.parent, ElementPlacement.PLACEATBEGINNING);
        }
        /* 移動で外れた選択を元に戻す / Restore the selection the moves dropped */
        doc.selection = selectedItems;
        /* グループの中を選ぶのに使ったダイレクト選択ツールから戻す / Switch back from the Direct Selection tool */
        app.selectTool("Adobe Select Tool");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * グループの中で選んだオブジェクトを、グループの外（レイヤーの最前面）へ出す
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        if (!selectedItems) return;
        releaseToLayerTop(app.activeDocument, selectedItems);
    }

    main();

})();
