#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ReleaseFromGroup
Meta-Default-language=ja
Name=グループから出してレイヤー直下に移動
Name[en]=Move objects out of their group onto the layer
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.2
Tags=group,グループ
X-com.swwwitch-Script-Name=ReleaseFromGroup
X-com.swwwitch-Updated=2026-09-27
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseFromGroup.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseFromGroup.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/n36fbd4162721
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/single-function/ReleaseFromGroup.jsx
Description-BEGIN
[ja]
選択中のオブジェクトがグループ内にある場合、親グループを辿って所属レイヤーの直下へ移動します。
重ね順が反転しないよう、選択オブジェクトは逆順に処理します。
[en]
Moves the selected objects out of their groups and directly onto the layer that owns them.
They are processed in reverse order so that the stacking order is preserved.
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
