#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ClipEditMask
Meta-Default-language=ja
Name=マスク編集モードへの切り替えとクリッピングマスク作成
Name[en]=Toggle mask-edit mode or create a clipping mask
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.1
Tags=mask,マスク
X-com.swwwitch-Script-Name=ClipEditMask
X-com.swwwitch-Updated=2026-09-30
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClipEditMask.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipEditMask.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipEditMask.jsx
Description-BEGIN
[ja]
選択がクリップグループのときはマスク編集モードへ切り替え、そうでないときは重なり合うオブジェクトごとにクリッピングマスクを作成します。
[en]
Enters mask-edit mode when the selection is a clipping group, and otherwise builds a clipping mask for each cluster of overlapping objects.
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
    // 重なりの判定 / Overlap detection
    // =========================================

    /**
     * 2つの矩形（geometricBounds）が重なっているかを判定する
     * @param {number[]} boundsA - [left, top, right, bottom]（top > bottom）
     * @param {number[]} boundsB - [left, top, right, bottom]（top > bottom）
     * @returns {boolean} 重なっていれば true
     */
    function isBoundsOverlapping(boundsA, boundsB) {
        var horizontalOverlap = (boundsA[0] < boundsB[2]) && (boundsB[0] < boundsA[2]);
        var verticalOverlap = (boundsA[3] < boundsB[1]) && (boundsB[3] < boundsA[1]);
        return horizontalOverlap && verticalOverlap;
    }

    /**
     * オブジェクトを重なりで連結成分（クラスタ）に分割する
     * 直接重ならなくても、間のオブジェクトを介して繋がっていれば同じクラスタになる
     * @param {PageItem[]} items - 対象のオブジェクト
     * @returns {PageItem[][]} クラスタごとのオブジェクト
     */
    function clusterItemsByOverlap(items) {
        var bounds = [];
        var visited = [];
        for (var i = 0; i < items.length; i++) {
            bounds.push(items[i].geometricBounds);
            visited.push(false);
        }

        var clusters = [];
        for (var start = 0; start < items.length; start++) {
            if (visited[start]) continue;

            /* start を起点に、重なりで繋がるものを幅優先で集める / Breadth-first from start over overlaps */
            var cluster = [];
            var queue = [start];
            visited[start] = true;

            while (queue.length > 0) {
                var current = queue.shift();
                cluster.push(items[current]);

                for (var other = 0; other < items.length; other++) {
                    if (!visited[other] && isBoundsOverlapping(bounds[current], bounds[other])) {
                        visited[other] = true;
                        queue.push(other);
                    }
                }
            }

            clusters.push(cluster);
        }

        return clusters;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 指定したオブジェクトだけを選択状態にする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 選択するオブジェクト
     * @returns {void}
     */
    function selectOnly(doc, items) {
        doc.selection = null;
        for (var i = 0; i < items.length; i++) {
            items[i].selected = true;
        }
    }

    /**
     * 1つ選択ならマスク編集／作成、複数なら重なりの塊ごとにクリッピングマスクを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、length は文字数になる / With characters selected by the Type tool, selection is a TextRange whose length is the character count */
        if (!currentSelection || currentSelection.typename === "TextRange" || currentSelection.length === 0) return;

        if (currentSelection.length === 1) {
            var selectedItem = currentSelection[0];
            if (selectedItem.typename === "GroupItem" && selectedItem.clipped) {
                /* ［マスクを編集］（2回で編集モードに入る）/ Edit Mask, run twice to enter the mode */
                app.executeMenuCommand("editMask");
                app.executeMenuCommand("editMask");
            } else {
                app.executeMenuCommand("makeMask");
            }
            return;
        }

        /* 2つ以上：重なりで塊に分けてから、塊ごとにマスク / Two or more: cluster by overlap, then mask each cluster */
        var selectedItems = [];
        for (var i = 0; i < currentSelection.length; i++) {
            selectedItems.push(currentSelection[i]);
        }

        var clusters = clusterItemsByOverlap(selectedItems);
        if (clusters.length === 1) {
            /* すべてが1つの塊 → そのまま1回だけマスク / One cluster: mask once as is */
            app.executeMenuCommand("makeMask");
            return;
        }

        for (var c = 0; c < clusters.length; c++) {
            selectOnly(doc, clusters[c]);
            app.executeMenuCommand("makeMask");
        }
    }

    main();

})();
