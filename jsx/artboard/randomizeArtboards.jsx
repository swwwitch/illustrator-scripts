#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

すべてのアートボードを、指定した列数のグリッドへランダムな順序で並べ替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/randomizeArtboards.md

### Overview

Shuffles all artboards into a grid with a fixed number of columns.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/randomizeArtboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "randomizeArtboards";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/randomizeArtboards.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/randomizeArtboards.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var COLUMNS = 4;   /* グリッドの列数 / number of grid columns */
    var GAP = 100;     /* アートボード間の余白（pt）/ gap between artboards (pt) */

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * アートボードをランダムな順序でグリッドに並べ替える
     * @returns {void}
     */
    function main() {
        /* ドキュメントが開かれているかチェック / Make sure a document is open */
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var doc = app.activeDocument;
        var artboards = doc.artboards;
        var artboardEntries = [];

        /* 1. 現在のアートボードの情報を控える（rect / name / 所属アイテム）/ Record each artboard's rect, name and member items */
        for (var i = 0; i < artboards.length; i++) {
            artboardEntries.push({
                rect: artboards[i].artboardRect,
                name: artboards[i].name,
                items: []
            });
        }

        /* シャッフル前に基準位置（元のアートボード 0 の左上）を記憶 / Remember the top-left of the original artboard 0 */
        var anchorLeft = artboardEntries[0].rect[0];
        var anchorTop = artboardEntries[0].rect[1];

        /* 2. 各ページアイテムを中心点で所属アートボードに割り当て（最初に一致したもの）/ Assign items by their center point */
        assignItemsToArtboards(doc, artboardEntries);

        /* 3. 配列をランダムにシャッフル / Shuffle the entries */
        shuffleArtboardOrder(artboardEntries);

        /* 4. グリッドに再配置（中身のアイテムも一緒に移動）/ Lay out on the grid, moving member items along */
        layoutArtboardsInGrid(artboardEntries, COLUMNS, GAP, anchorLeft, anchorTop);

        /* 5. シャッフル＋再配置した情報を元のアートボードに上書き / Write the shuffled rects and names back */
        for (var j = 0; j < artboards.length; j++) {
            artboards[j].artboardRect = artboardEntries[j].rect;
            artboards[j].name = artboardEntries[j].name;
        }

        app.redraw();

        /* 画面を全体表示に更新 / Fit all in view */
        app.executeMenuCommand("fitall");
    }

    /**
     * 配列をランダムにシャッフルする（フィッシャー–イェーツのシャッフル）
     * @param {Object[]} artboardEntries - アートボードの控え（その場で並べ替える）
     * @returns {void}
     */
    function shuffleArtboardOrder(artboardEntries) {
        for (var i = artboardEntries.length - 1; i > 0; i--) {
            var swapIndex = Math.floor(Math.random() * (i + 1));
            var swappedEntry = artboardEntries[i];
            artboardEntries[i] = artboardEntries[swapIndex];
            artboardEntries[swapIndex] = swappedEntry;
        }
    }

    /**
     * 各ページアイテムを、中心点が含まれるアートボードに割り当てる
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} artboardEntries - アートボードの控え（items に追加する）
     * @returns {void}
     */
    function assignItemsToArtboards(doc, artboardEntries) {
        for (var itemIndex = 0; itemIndex < doc.pageItems.length; itemIndex++) {
            var pageItem = doc.pageItems[itemIndex];
            /* 入れ子のアイテムは親（グループ等）の移動で連動するためスキップ / Nested items move with their parent */
            if (pageItem.parent.typename !== "Layer") continue;
            if (pageItem.locked || pageItem.hidden) continue;
            if (pageItem.parent.locked || pageItem.parent.visible === false) continue;

            var itemBounds = pageItem.geometricBounds;
            var centerX = (itemBounds[0] + itemBounds[2]) / 2;
            var centerY = (itemBounds[1] + itemBounds[3]) / 2;

            for (var artboardIndex = 0; artboardIndex < artboardEntries.length; artboardIndex++) {
                var artboardRect = artboardEntries[artboardIndex].rect;
                var minX = Math.min(artboardRect[0], artboardRect[2]);
                var maxX = Math.max(artboardRect[0], artboardRect[2]);
                var minY = Math.min(artboardRect[1], artboardRect[3]);
                var maxY = Math.max(artboardRect[1], artboardRect[3]);
                if (centerX >= minX && centerX <= maxX && centerY >= minY && centerY <= maxY) {
                    artboardEntries[artboardIndex].items.push(pageItem);
                    break;
                }
            }
        }
    }

    /**
     * アートボードを N 列のグリッドに配置する（所属アイテムも同じ移動量で平行移動）
     * @param {Object[]} artboardEntries - アートボードの控え（rect を書き換える）
     * @param {number} columns - 列数
     * @param {number} gap - アートボード間の余白（pt）
     * @param {number} anchorLeft - 配置の基準にする左端
     * @param {number} anchorTop - 配置の基準にする上端
     * @returns {void}
     */
    function layoutArtboardsInGrid(artboardEntries, columns, gap, anchorLeft, anchorTop) {
        if (artboardEntries.length === 0) return;

        /* 各アートボードの最大幅・高さからセルサイズを決める / Cell size from the largest artboard */
        var maxWidth = 0, maxHeight = 0;
        for (var i = 0; i < artboardEntries.length; i++) {
            var entryRect = artboardEntries[i].rect;
            var entryWidth = entryRect[2] - entryRect[0];
            var entryHeight = Math.abs(entryRect[3] - entryRect[1]);
            if (entryWidth > maxWidth) maxWidth = entryWidth;
            if (entryHeight > maxHeight) maxHeight = entryHeight;
        }

        var cellWidth = maxWidth + gap;
        var cellHeight = maxHeight + gap;

        /* y 軸の向き（top と bottom の大小関係から判定）/ Direction of the y axis */
        var referenceRect = artboardEntries[0].rect;
        var verticalDirection = (referenceRect[3] > referenceRect[1]) ? 1 : -1;

        for (var j = 0; j < artboardEntries.length; j++) {
            var columnIndex = j % columns;
            var rowIndex = Math.floor(j / columns);
            var oldRect = artboardEntries[j].rect;
            var artboardWidth = oldRect[2] - oldRect[0];
            var artboardHeight = oldRect[3] - oldRect[1];

            var newLeft = anchorLeft + columnIndex * cellWidth;
            var newTop = anchorTop + rowIndex * cellHeight * verticalDirection;

            var deltaX = newLeft - oldRect[0];
            var deltaY = newTop - oldRect[1];

            /* 所属アイテムを同じ移動量で平行移動 / Move member items by the same amount */
            var memberItems = artboardEntries[j].items;
            for (var itemIndex = 0; itemIndex < memberItems.length; itemIndex++) {
                /* 動かせないアイテムは飛ばして続行 / Skip items that cannot be moved */
                try {
                    memberItems[itemIndex].translate(deltaX, deltaY);
                } catch (e) { }
            }

            artboardEntries[j].rect = [newLeft, newTop, newLeft + artboardWidth, newTop + artboardHeight];
        }
    }

    main();
})();
