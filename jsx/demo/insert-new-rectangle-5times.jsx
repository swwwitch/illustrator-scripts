#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

現在の表示領域の中心付近に、黒く塗ってランダムな不透明度を与えた正方形を5つ作成し、重ならないように配置してから「Convert to Shape」「Make Pixel Perfect」コマンドを適用します。

詳細は README を参照してください。

### Overview

Creates five black squares with random opacity near the center of the current view, places them so that
they do not overlap, and applies the "Convert to Shape" and "Make Pixel Perfect" commands.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InsertNewRectangle5Times";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-25";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InsertNewRectangle5Times.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InsertNewRectangle5Times.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n509eb6aa0a19"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var RECT_SIZE     = 100; /* 作成する正方形の一辺 / side length of the square */
    var RECT_COUNT    = 5;   /* 作成する個数 / number of squares */
    var SCATTER_RANGE = 40;  /* 初期配置のばらつき（±px）/ initial scatter range (±px) */
    var OPACITY_MIN   = 30;  /* 不透明度の下限（%）/ minimum opacity (%) */
    var OPACITY_MAX   = 100; /* 不透明度の上限（%）/ maximum opacity (%) */

    /* 重なり回避の探索条件（100px角×5個なので少し広めに）/ Overlap avoidance search settings */
    var PLACEMENT_OPTIONS = {
        padding: 6,           /* オブジェクト間の最小間隔 / minimum gap between items */
        baseX: 160,           /* 水平方向の探索幅 / horizontal search range */
        baseY: 160,           /* 垂直方向の探索幅 / vertical search range */
        maxScaleFactor: 20,   /* 探索幅を広げる最大倍率 / maximum search range multiplier */
        attemptsPerItem: 300  /* 1個あたりの試行回数 / attempts per item */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントがありません。", en: "No document is open." },
            noEditableLayer: {
                ja: "ロック解除かつ表示されているレイヤーがありません。",
                en: "No unlocked and visible layers found."
            }
        }
    };

    /**
     * 現在のUI言語に合わせた文言を返す
     * @param {object} labelEntry - ja / en を持つラベル定義
     * @returns {string} 表示する文言
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * ドキュメントのカラーモードに応じた黒を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {CMYKColor|RGBColor} 黒のカラーオブジェクト
     */
    function getBlackFillColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 100;
            return cmyk;
        }
        var rgb = new RGBColor();
        rgb.red = 0;
        rgb.green = 0;
        rgb.blue = 0;
        return rgb;
    }

    /**
     * ロックされておらず表示されている最初のレイヤーを返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 編集可能なレイヤー。見つからない場合は null
     */
    function getUnlockedVisibleLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.locked && layer.visible) return layer;
        }
        return null;
    }

    /**
     * 作成先レイヤーを決め、必要ならアクティブレイヤーを切り替える
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 作成先レイヤー。見つからない場合は null
     */
    function resolveTargetLayer(doc) {
        var targetLayer = doc.activeLayer;
        if (!targetLayer.locked && targetLayer.visible) return targetLayer;

        /* ロック／非表示なら、順にロック解除かつ表示のレイヤーを探す / Fall back to the first editable layer */
        var editableLayer = getUnlockedVisibleLayer(doc);
        if (!editableLayer) return null;

        doc.activeLayer = editableLayer;
        return editableLayer;
    }

    /**
     * 2つのバウンディングボックスが指定間隔以内で重なっているか判定する
     * @param {number[]} bounds - 判定するバウンズ [left, top, right, bottom]
     * @param {number[]} otherBounds - 比較先のバウンズ [left, top, right, bottom]
     * @param {number} padding - 重なりとみなす余白
     * @returns {boolean} 重なっていれば true
     */
    function isOverlapping(bounds, otherBounds, padding) {
        return !(bounds[2] + padding < otherBounds[0] ||
                 bounds[0] - padding > otherBounds[2] ||
                 bounds[1] + padding < otherBounds[3] ||
                 bounds[3] - padding > otherBounds[1]);
    }

    /**
     * 元の位置を起点にランダム移動を繰り返し、互いに重ならない位置へ配置する
     * @param {Array<{item: PageItem, position: number[]}>} states - 対象と元の位置の組
     * @param {object} placementOptions - PLACEMENT_OPTIONS と同じ形の探索条件
     * @returns {boolean} 全件を配置できたら true
     */
    function placeItemsAvoidOverlap(states, placementOptions) {
        if (!states || states.length === 0) return false;

        var scaleFactor = 1;

        while (scaleFactor <= placementOptions.maxScaleFactor) {
            var placedBounds = [];
            var allPlaced = true;

            for (var i = 0; i < states.length; i++) {
                var state = states[i];
                var placed = false;

                for (var attempt = 0; attempt < placementOptions.attemptsPerItem; attempt++) {
                    var randX = (Math.random() * 2 - 1) * placementOptions.baseX * scaleFactor;
                    var randY = (Math.random() * 2 - 1) * placementOptions.baseY * scaleFactor;
                    state.item.position = [state.position[0] + randX, state.position[1] + randY];

                    var bounds = state.item.visibleBounds;
                    var overlap = false;
                    for (var j = 0; j < placedBounds.length; j++) {
                        if (isOverlapping(bounds, placedBounds[j], placementOptions.padding)) {
                            overlap = true;
                            break;
                        }
                    }

                    if (!overlap) {
                        placedBounds.push(bounds);
                        placed = true;
                        break;
                    }
                }

                if (!placed) {
                    allPlaced = false;
                    break;
                }
            }

            if (allPlaced) return true;

            /* 収まらなければ探索幅を広げて再試行 / Widen the search range and retry */
            scaleFactor += 1;
        }

        return false;
    }

    /**
     * 中央に残っている同じサイズ・同じ位置の長方形を削除する
     * @param {Layer} layer - 対象レイヤー
     * @param {number} rectTop - 判定する上端座標
     * @param {number} rectLeft - 判定する左端座標
     * @param {number} rectSize - 判定する一辺の長さ
     * @returns {void}
     */
    function removeExistingCenterRect(layer, rectTop, rectLeft, rectSize) {
        for (var i = layer.pathItems.length - 1; i >= 0; i--) {
            var pathItem = layer.pathItems[i];
            if (Math.abs(pathItem.top - rectTop) < 1 && Math.abs(pathItem.left - rectLeft) < 1 &&
                Math.abs(pathItem.width - rectSize) < 1 && Math.abs(pathItem.height - rectSize) < 1) {
                try {
                    pathItem.remove();
                } catch (e) {
                    /* 削除できないものは飛ばす / Skip items that cannot be removed */
                }
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 表示中心付近に正方形を複数作成し、重ならないように配置してコマンドを適用する
     * @returns {void}
     */
    function main() {
        /* ドキュメント確認 / Ensure a document is open */
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;

        var targetLayer = resolveTargetLayer(doc);
        if (!targetLayer) {
            alert(getLabel(LABELS.alert.noEditableLayer));
            return;
        }

        /* 表示領域の中心座標を取得 / Get view center */
        var viewCenterX = doc.activeView.centerPoint[0];
        var viewCenterY = doc.activeView.centerPoint[1];

        /* 中央に残っている既存の長方形を削除 / Remove the existing center rectangle if any */
        removeExistingCenterRect(targetLayer, viewCenterY + RECT_SIZE / 2, viewCenterX - RECT_SIZE / 2, RECT_SIZE);

        /* いったん表示中心付近に作成し、後で重なり回避で再配置する / Create near the center, reposition later */
        var rects = [];
        for (var i = 0; i < RECT_COUNT; i++) {
            var offsetX = Math.random() * SCATTER_RANGE * 2 - SCATTER_RANGE;
            var offsetY = Math.random() * SCATTER_RANGE * 2 - SCATTER_RANGE;

            var rect = targetLayer.pathItems.rectangle(
                viewCenterY + offsetY + RECT_SIZE / 2,
                viewCenterX + offsetX - RECT_SIZE / 2,
                RECT_SIZE,
                RECT_SIZE
            );
            rect.fillColor = getBlackFillColor(doc);
            rect.stroked = false;
            rect.opacity = Math.random() * (OPACITY_MAX - OPACITY_MIN) + OPACITY_MIN;
            rects.push(rect);
        }

        /* 重ならないように配置 / Place without overlapping */
        var states = [];
        for (var s = 0; s < rects.length; s++) {
            states.push({ item: rects[s], position: rects[s].position }); /* position: [left, top] */
        }
        placeItemsAvoidOverlap(states, PLACEMENT_OPTIONS);

        /* 作成した正方形だけを選択 / Select the created squares only */
        doc.selection = null;
        for (var k = 0; k < rects.length; k++) {
            rects[k].selected = true;
        }

        /* 選択オブジェクトにコマンドを適用 / Apply commands to the selection */
        app.executeMenuCommand("Convert to Shape");
        app.executeMenuCommand("Make Pixel Perfect");
    }

    main();

})();
