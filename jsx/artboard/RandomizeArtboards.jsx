#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

すべてのアートボードを、指定した列数のグリッドへランダムな順序で並べ替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RandomizeArtboards.md

### Overview

Shuffles all artboards into a grid with a fixed number of columns.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RandomizeArtboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RandomizeArtboards";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RandomizeArtboards.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RandomizeArtboards.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var GRID_COLUMNS = 4;    /* グリッドの列数 / number of grid columns */
    var ARTBOARD_GAP = 100;  /* アートボード間の余白（pt）/ gap between artboards (pt) */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

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
            alert(getLabel("alert.noDocument"));
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
        layoutArtboardsInGrid(artboardEntries, GRID_COLUMNS, ARTBOARD_GAP, anchorLeft, anchorTop);

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
        var pageItems = doc.pageItems;
        var pageItemCount = pageItems.length;
        for (var itemIndex = 0; itemIndex < pageItemCount; itemIndex++) {
            var pageItem = pageItems[itemIndex];
            /* 入れ子のアイテムは親（グループ等）の移動で連動するためスキップ / Nested items move with their parent */
            if (pageItem.parent.typename !== "Layer") continue;
            if (pageItem.locked || pageItem.hidden) continue;
            if (pageItem.parent.locked || !pageItem.parent.visible) continue;

            var itemBounds = pageItem.geometricBounds;
            var centerX = (itemBounds[0] + itemBounds[2]) / 2;
            var centerY = (itemBounds[1] + itemBounds[3]) / 2;

            for (var artboardIndex = 0; artboardIndex < artboardEntries.length; artboardIndex++) {
                /* artboardRect は [左, 上, 右, 下]（上 > 下）/ artboardRect is [left, top, right, bottom] with top > bottom */
                var artboardRect = artboardEntries[artboardIndex].rect;
                if (centerX >= artboardRect[0] && centerX <= artboardRect[2] &&
                    centerY <= artboardRect[1] && centerY >= artboardRect[3]) {
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
            maxWidth = Math.max(maxWidth, entryRect[2] - entryRect[0]);
            maxHeight = Math.max(maxHeight, entryRect[1] - entryRect[3]);
        }

        var cellWidth = maxWidth + gap;
        var cellHeight = maxHeight + gap;

        for (var j = 0; j < artboardEntries.length; j++) {
            var columnIndex = j % columns;
            var rowIndex = Math.floor(j / columns);
            var oldRect = artboardEntries[j].rect;
            var artboardWidth = oldRect[2] - oldRect[0];
            var artboardHeight = oldRect[1] - oldRect[3];

            /* 行は下へ進む（Illustrator の y は上が大きい）/ Rows go downward (y grows upward in Illustrator) */
            var newLeft = anchorLeft + columnIndex * cellWidth;
            var newTop = anchorTop - rowIndex * cellHeight;

            var deltaX = newLeft - oldRect[0];
            var deltaY = newTop - oldRect[1];

            /* 所属アイテムを同じ移動量で平行移動 / Move member items by the same amount */
            var memberItems = artboardEntries[j].items;
            for (var itemIndex = 0; itemIndex < memberItems.length; itemIndex++) {
                /* サブレイヤーは親レイヤーのロックが自分の locked に出ないため、動かせないものは飛ばす / A sublayer's locked does not reflect its parent layer's lock, so skip items that cannot be moved */
                try {
                    memberItems[itemIndex].translate(deltaX, deltaY);
                } catch (e) { }
            }

            artboardEntries[j].rect = [newLeft, newTop, newLeft + artboardWidth, newTop - artboardHeight];
        }
    }

    main();
})();
