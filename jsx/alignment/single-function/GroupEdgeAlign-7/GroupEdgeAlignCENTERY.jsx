#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択されているオブジェクト群の端または中心を取得し、上下中央に揃えます。
揃え先は、アクティブアートボードの端、または条件に合うガイドです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignCENTERY.md

### Overview

Takes the edges or the center of the selected objects and aligns them to the vertical center.
The target is the edge of the active artboard, or a matching guide.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlignCENTERY.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEdgeAlignCENTERY";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlignCENTERY.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlignCENTERY.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* アートボード端の手前にガイドがあれば、そのガイドへ揃える / snap to a guide when one is in the way */
    var USE_GUIDES = true;

    /* 境界の取り方 "preference"（環境設定に従う）| "preview"（線を含む）| "geometric"（線を含まない）*/
    var BOUNDS_MODE = "preference";

    /* ガイドの探し方 "inside"（揃える向きで最も近いもの）| "nearest"（距離が最も近いもの）*/
    var GUIDE_SEARCH_MODE = "inside";

    /* ガイドの水平・垂直判定に使う許容値（pt）/ tolerance for classifying a guide as H or V */
    var GUIDE_ORIENTATION_TOLERANCE = 0.01;

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
            invalidBoundsMode: { ja: "BOUNDS_MODE の指定が不正です", en: "Invalid BOUNDS_MODE" },
            invalidGuideSearchMode: { ja: "GUIDE_SEARCH_MODE の指定が不正です", en: "Invalid GUIDE_SEARCH_MODE" },
            invalidAlignmentSide: { ja: "揃える方向を判定できませんでした", en: "Could not determine the alignment direction" }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - LABELS のリーフ（{ ja, en }）
     * @returns {string} 現在の言語の文字列（無ければ空文字）
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * ラベルに言語別のコロンと値を続けた文字列を返す（日本語は全角、英語は半角＋スペース）
     * @param {Object} labelSet - LABELS のリーフ（{ ja, en }）
     * @param {string} value - コロンの後ろに続ける値
     * @returns {string} 「ラベル：値」の文字列
     */
    function labelWithValue(labelSet, value) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ": ") + value;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクト群の端または中心を、ファイル名から判定した方向へ揃える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var documentRef = app.activeDocument;
        var selectedItems = documentRef.selection;
        if (selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var alignmentSide = resolveAlignmentSideFromFileName();

        var includeStrokeInBounds = resolveIncludeStrokeInBounds(BOUNDS_MODE);
        if (includeStrokeInBounds === null) {
            alert(labelWithValue(LABELS.alert.invalidBoundsMode, BOUNDS_MODE));
            return;
        }

        var artboards = documentRef.artboards;
        var activeArtboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
        var selectionBounds = getSelectionBounds(selectedItems, includeStrokeInBounds);

        var moveOffset = getAlignmentOffset(documentRef, selectionBounds, activeArtboardRect, alignmentSide);
        if (moveOffset === null) return; /* 方向の指定が不正。getAlignmentOffset が通知済み */

        for (var i = 0; i < selectedItems.length; i++) {
            selectedItems[i].translate(moveOffset.x, moveOffset.y);
        }
    }

    /**
     * スクリプトのファイル名から揃える方向を判定する（例: GroupEdgeAlignRIGHT.jsx → "right"）
     * CENTERX / CENTERY は CENTER より先に判定する必要がある。
     * @returns {string} 揃える方向。判定できなければ "left"
     */
    function resolveAlignmentSideFromFileName() {
        var fileNameUpper = File($.fileName).name.toUpperCase();

        if (fileNameUpper.indexOf("CENTERX") !== -1) return "CENTER_X";
        if (fileNameUpper.indexOf("CENTERY") !== -1) return "CENTER_Y";
        if (fileNameUpper.indexOf("CENTER") !== -1) return "CENTER";
        if (fileNameUpper.indexOf("LEFT") !== -1) return "left";
        if (fileNameUpper.indexOf("RIGHT") !== -1) return "right";
        if (fileNameUpper.indexOf("TOP") !== -1) return "top";
        if (fileNameUpper.indexOf("BOTTOM") !== -1) return "bottom";
        return "left";
    }

    /**
     * BOUNDS_MODE から「線を境界に含めるか」を決める
     * @param {string} boundsMode - "preference" | "preview" | "geometric"
     * @returns {boolean|null} 線を含めるなら true。指定が不正なら null
     */
    function resolveIncludeStrokeInBounds(boundsMode) {
        if (boundsMode === "preference") return app.preferences.getBooleanPreference("includeStrokeInBounds");
        if (boundsMode === "preview") return true;
        if (boundsMode === "geometric") return false;
        return null;
    }

    /**
     * 選択オブジェクト全体を囲む境界を返す
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @param {boolean} includeStrokeInBounds - 線を境界に含めるか
     * @returns {number[]} [left, top, right, bottom]
     */
    function getSelectionBounds(selectedItems, includeStrokeInBounds) {
        var selectionBounds = getItemBounds(selectedItems[0], includeStrokeInBounds).slice(0);

        for (var i = 1; i < selectedItems.length; i++) {
            var itemBounds = getItemBounds(selectedItems[i], includeStrokeInBounds);
            if (itemBounds[0] < selectionBounds[0]) selectionBounds[0] = itemBounds[0];
            if (itemBounds[1] > selectionBounds[1]) selectionBounds[1] = itemBounds[1];
            if (itemBounds[2] > selectionBounds[2]) selectionBounds[2] = itemBounds[2];
            if (itemBounds[3] < selectionBounds[3]) selectionBounds[3] = itemBounds[3];
        }

        return selectionBounds;
    }

    /**
     * 指定オブジェクトの境界を返す
     * クリッピンググループはマスクパス（先頭アイテム）の geometricBounds を使う。
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {boolean} includeStrokeInBounds - 線を境界に含めるか
     * @returns {number[]} 境界 [L, T, R, B]
     */
    function getItemBounds(pageItem, includeStrokeInBounds) {
        if (pageItem.typename === "GroupItem" && pageItem.clipped === true) {
            return pageItem.pageItems[0].geometricBounds;
        }
        return includeStrokeInBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
    }

    /**
     * 揃える方向から移動量を求める
     * @param {Document} documentRef - 対象のドキュメント
     * @param {number[]} selectionBounds - 選択範囲の境界 [L, T, R, B]
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @returns {{x: number, y: number}|null} 移動量。方向の指定が不正なら null
     */
    function getAlignmentOffset(documentRef, selectionBounds, artboardRect, alignmentSide) {
        var centerOffsetX = getHorizontalCenterValue(artboardRect) - getHorizontalCenterValue(selectionBounds);
        var centerOffsetY = getVerticalCenterValue(artboardRect) - getVerticalCenterValue(selectionBounds);

        if (alignmentSide === "CENTER_X") return { x: centerOffsetX, y: 0 };
        if (alignmentSide === "CENTER_Y") return { x: 0, y: centerOffsetY };
        if (alignmentSide === "CENTER") return { x: centerOffsetX, y: centerOffsetY };

        var selectionEdge = getEdgeValueForAlignmentSide(selectionBounds, alignmentSide);
        var artboardEdge = getEdgeValueForAlignmentSide(artboardRect, alignmentSide);
        if (selectionEdge === null || artboardEdge === null) {
            alert(labelWithValue(LABELS.alert.invalidAlignmentSide, alignmentSide));
            return null;
        }

        var targetEdge = artboardEdge;
        if (USE_GUIDES) {
            var snappedGuideValue = findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide);
            if (snappedGuideValue !== null) targetEdge = snappedGuideValue;
        }

        var edgeOffset = targetEdge - selectionEdge;
        return isHorizontalSide(alignmentSide) ? { x: edgeOffset, y: 0 } : { x: 0, y: edgeOffset };
    }

    /**
     * 揃える方向が左右（X方向の移動）かを返す
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} "left" / "right" なら true
     */
    function isHorizontalSide(alignmentSide) {
        return alignmentSide === "left" || alignmentSide === "right";
    }

    /**
     * 指定した方向に対応する境界値を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向
     * @returns {number|null} 境界値。方向が不正なら null
     */
    function getEdgeValueForAlignmentSide(boundsRect, alignmentSide) {
        if (alignmentSide === "left") return boundsRect[0];
        if (alignmentSide === "top") return boundsRect[1];
        if (alignmentSide === "right") return boundsRect[2];
        if (alignmentSide === "bottom") return boundsRect[3];
        return null;
    }

    /**
     * 左右中央の座標を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @returns {number} 左右中央のX
     */
    function getHorizontalCenterValue(boundsRect) {
        return (boundsRect[0] + boundsRect[2]) / 2;
    }

    /**
     * 上下中央の座標を返す
     * @param {number[]} boundsRect - 境界 [L, T, R, B]
     * @returns {number} 上下中央のY
     */
    function getVerticalCenterValue(boundsRect) {
        return (boundsRect[1] + boundsRect[3]) / 2;
    }

    /**
     * アートボード内側にあるガイドのうち、揃える方向と GUIDE_SEARCH_MODE に合う吸着先座標を返す
     * @param {Document} documentRef - 対象のドキュメント
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {number} selectionEdge - 選択範囲の該当する端の座標
     * @param {string} alignmentSide - 揃える方向（"left" | "right" | "top" | "bottom"）
     * @returns {number|null} 吸着先の座標。該当するガイドがなければ null
     */
    function findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide) {
        if (GUIDE_SEARCH_MODE !== "inside" && GUIDE_SEARCH_MODE !== "nearest") {
            alert(labelWithValue(LABELS.alert.invalidGuideSearchMode, GUIDE_SEARCH_MODE));
            return null;
        }

        /* アートボード端へ向かう向きの符号。left / bottom は座標が減る向き / sign toward the artboard edge */
        var outwardSign = (alignmentSide === "left" || alignmentSide === "bottom") ? -1 : 1;
        var nearestGuideValue = null;
        var nearestGuideDistance = null;
        var documentPathItems = documentRef.pathItems;

        for (var i = 0; i < documentPathItems.length; i++) {
            var guidePathItem = documentPathItems[i];
            if (guidePathItem.guides !== true) continue;

            var guideValue = getGuideValueForAlignmentSide(guidePathItem.geometricBounds, alignmentSide);
            if (guideValue === null) continue;
            if (!isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide)) continue;

            if (GUIDE_SEARCH_MODE === "inside") {
                /* 選択範囲より揃える向きの側にあり、その中で選択範囲に最も近いもの / nearest guide on the outward side */
                if ((guideValue - selectionEdge) * outwardSign <= 0) continue;
                if (nearestGuideValue === null || (guideValue - nearestGuideValue) * outwardSign < 0) {
                    nearestGuideValue = guideValue;
                }
            } else {
                var guideDistance = Math.abs(guideValue - selectionEdge);
                if (nearestGuideDistance === null || guideDistance < nearestGuideDistance) {
                    nearestGuideValue = guideValue;
                    nearestGuideDistance = guideDistance;
                }
            }
        }

        return nearestGuideValue;
    }

    /**
     * 揃える方向に対応するガイド座標を返す（左右なら垂直ガイドのX、上下なら水平ガイドのY）
     * @param {number[]} guideBounds - ガイドの境界 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向（"left" | "right" | "top" | "bottom"）
     * @returns {number|null} ガイドの座標。向きが対応しなければ null
     */
    function getGuideValueForAlignmentSide(guideBounds, alignmentSide) {
        if (isHorizontalSide(alignmentSide)) {
            var isVerticalGuide = Math.abs(guideBounds[2] - guideBounds[0]) <= GUIDE_ORIENTATION_TOLERANCE;
            return isVerticalGuide ? guideBounds[0] : null;
        }
        var isHorizontalGuide = Math.abs(guideBounds[1] - guideBounds[3]) <= GUIDE_ORIENTATION_TOLERANCE;
        return isHorizontalGuide ? guideBounds[1] : null;
    }

    /**
     * ガイド座標がアクティブアートボード内にあるかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [L, T, R, B]
     * @param {string} alignmentSide - 揃える方向（"left" | "right" | "top" | "bottom"）
     * @returns {boolean} 内側にあれば true
     */
    function isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide) {
        if (isHorizontalSide(alignmentSide)) {
            return guideValue >= artboardRect[0] && guideValue <= artboardRect[2];
        }
        return guideValue <= artboardRect[1] && guideValue >= artboardRect[3];
    }

    main();

})();
