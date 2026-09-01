#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、アクティブアートボードの端・中央、または条件に合うガイドへ移動します。
テキストは字形（アウトライン）の境界で測るため、「字形の境界に整列」をオンにして整列したときと同じ位置になります。

詳細は README を参照してください。

### Overview

Moves the selected objects to an edge or the center of the active artboard, or to a matching guide.
Type is measured from its outlined glyphs, so the result matches aligning with Align to Glyph Bounds turned on.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "アートボードの端またはガイドに移動"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-31";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/アートボードの端またはガイドに移動.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/アートボードの端またはガイドに移動.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4ae0e1e70481"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion オリジナル / Original: Gorolib Design
 * https://gorolib.blog.jp/archives/63149753.html
 */

(function (scriptArguments) {

    // =========================================================
    // ユーザー設定 / User Settings
    // =========================================================

    /* 揃える方向。Keyboard Maestro などから引数で渡されたときはそちらを優先
       "left" | "right" | "top" | "bottom" | "CENTER_X" | "CENTER_Y" | "CENTER" */
    var DEFAULT_ALIGNMENT_SIDE = "right";

    var USE_GUIDES        = true;         /* ガイドへスナップするか（端揃えのみ）/ snap to guides (edges only) */
    var GUIDE_SEARCH_MODE = "inside";     /* "inside"（揃える方向側のみ）| "nearest"（最も近いガイド）*/
    var BOUNDS_MODE       = "preference"; /* "preference"（環境設定）| "preview"（線幅込み）| "geometric"（線幅なし）*/
    var GLYPH_BOUNDS_MODE = "on";         /* 字形の境界に整列 "on" | "off" | "preference"（環境設定）*/

    var GUIDE_ORIENTATION_TOLERANCE = 0.01; /* ガイドの水平・垂直判定に使う許容値 / tolerance for guide orientation */

    // =========================================================
    // ローカライズ / Localization
    // =========================================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noSelection:            { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
            invalidBoundsMode:      { ja: "BOUNDS_MODE の指定が不正です：", en: "Invalid BOUNDS_MODE: " },
            invalidGlyphBoundsMode: { ja: "GLYPH_BOUNDS_MODE の指定が不正です：", en: "Invalid GLYPH_BOUNDS_MODE: " },
            invalidGuideSearchMode: { ja: "GUIDE_SEARCH_MODE の指定が不正です：", en: "Invalid GUIDE_SEARCH_MODE: " },
            invalidAlignmentSide:   { ja: "ALIGNMENT_SIDE の指定が不正です：", en: "Invalid ALIGNMENT_SIDE: " },
            unexpectedError:        { ja: "エラーが発生しました：", en: "An error occurred: " }
        }
    };

    /**
     * 現在のUI言語に合わせたラベルを返す
     * @param {string} categoryKey - LABELS のカテゴリ名
     * @param {string} labelKey - カテゴリ内のキー名
     * @returns {string} ラベル文字列（見つからない場合はキー名）
     */
    function getLabel(categoryKey, labelKey) {
        var category = LABELS[categoryKey];
        if (!category || !category[labelKey]) return labelKey;
        return category[labelKey][uiLang] || category[labelKey].en || labelKey;
    }

    // =========================================================
    // メイン処理 / Main
    // =========================================================

    /**
     * 選択オブジェクトをアートボードの端・中央、または条件に合うガイドへ移動する
     * @param {string} alignmentSideArg - 揃える方向。未指定時は DEFAULT_ALIGNMENT_SIDE を使う
     * @returns {void}
     */
    function main(alignmentSideArg) {
        try {
            var documentRef = app.activeDocument;
            var selectedItems = documentRef.selection;
            var alignmentSide = alignmentSideArg || DEFAULT_ALIGNMENT_SIDE;

            /* 選択されていない場合は処理中止 / Stop when nothing is selected */
            if (selectedItems.length === 0) {
                alert(getLabel("alert", "noSelection"));
                return;
            }

            var includeStrokeInBounds = resolveIncludeStrokeInBounds(BOUNDS_MODE);
            if (includeStrokeInBounds === null) {
                alert(getLabel("alert", "invalidBoundsMode") + BOUNDS_MODE);
                return;
            }

            var useGlyphBounds = resolveUseGlyphBounds(GLYPH_BOUNDS_MODE);
            if (useGlyphBounds === null) {
                alert(getLabel("alert", "invalidGlyphBoundsMode") + GLYPH_BOUNDS_MODE);
                return;
            }

            if (GUIDE_SEARCH_MODE !== "inside" && GUIDE_SEARCH_MODE !== "nearest") {
                alert(getLabel("alert", "invalidGuideSearchMode") + GUIDE_SEARCH_MODE);
                return;
            }

            var artboards = documentRef.artboards;
            var activeArtboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
            var selectionBounds = getSelectionBounds(selectedItems, includeStrokeInBounds, useGlyphBounds);

            var offset = getAlignmentOffset(documentRef, selectionBounds, activeArtboardRect, alignmentSide);
            if (offset === null) {
                alert(getLabel("alert", "invalidAlignmentSide") + alignmentSide);
                return;
            }

            /* オブジェクトを移動 / Move the objects */
            for (var moveItemIndex = 0; moveItemIndex < selectedItems.length; moveItemIndex++) {
                selectedItems[moveItemIndex].translate(offset.x, offset.y);
            }

        } catch (e) {
            alert(getLabel("alert", "unexpectedError") + e.message);
        }
    }

    // =========================================================
    // 境界の取得 / Bounds
    // =========================================================

    /**
     * BOUNDS_MODE から線幅を含めるかどうかを決める
     * @param {string} boundsMode - "preference" | "preview" | "geometric"
     * @returns {boolean|null} 線幅を含めるならtrue、不正な指定の場合はnull
     */
    function resolveIncludeStrokeInBounds(boundsMode) {
        if (boundsMode === "preference") return app.preferences.getBooleanPreference("includeStrokeInBounds") === true;
        if (boundsMode === "preview") return true;
        if (boundsMode === "geometric") return false;
        return null;
    }

    /**
     * GLYPH_BOUNDS_MODE から字形の境界で測るかどうかを決める
     * @param {string} glyphBoundsMode - "on" | "off" | "preference"
     * @returns {boolean|null} 字形の境界で測るならtrue、不正な指定の場合はnull
     */
    function resolveUseGlyphBounds(glyphBoundsMode) {
        if (glyphBoundsMode === "on") return true;
        if (glyphBoundsMode === "off") return false;
        if (glyphBoundsMode === "preference") {
            /* ポイント文字・エリア内文字のどちらかがオンなら字形の境界で測る */
            return app.preferences.getBooleanPreference("EnableActualPointTextSpaceAlign") === true ||
                app.preferences.getBooleanPreference("EnableActualAreaTextSpaceAlign") === true;
        }
        return null;
    }

    /**
     * 選択オブジェクト全体を囲む境界を返す
     * @param {Array<PageItem>} selectedItems - 選択オブジェクト
     * @param {boolean} includeStrokeInBounds - 線幅を含めるか
     * @param {boolean} useGlyphBounds - テキストを字形の境界で測るか
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getSelectionBounds(selectedItems, includeStrokeInBounds, useGlyphBounds) {
        var selectionBounds = null;

        for (var selectedItemIndex = 0; selectedItemIndex < selectedItems.length; selectedItemIndex++) {
            var itemBounds = getItemBounds(selectedItems[selectedItemIndex], includeStrokeInBounds, useGlyphBounds);
            if (!itemBounds) continue;

            if (selectionBounds === null) {
                selectionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
            } else {
                if (itemBounds[0] < selectionBounds[0]) selectionBounds[0] = itemBounds[0];
                if (itemBounds[1] > selectionBounds[1]) selectionBounds[1] = itemBounds[1];
                if (itemBounds[2] > selectionBounds[2]) selectionBounds[2] = itemBounds[2];
                if (itemBounds[3] < selectionBounds[3]) selectionBounds[3] = itemBounds[3];
            }
        }

        return selectionBounds;
    }

    /**
     * 指定オブジェクトの境界を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {boolean} includeStrokeInBounds - 線幅を含めるか
     * @param {boolean} useGlyphBounds - テキストを字形の境界で測るか
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(pageItem, includeStrokeInBounds, useGlyphBounds) {
        /* クリッピンググループは内部要素の境界を参照 / Clipping groups use the clipping path */
        if (pageItem.typename === "GroupItem" && pageItem.clipped === true) {
            return pageItem.pageItems[0].geometricBounds;
        }

        /* テキストは複製をアウトライン化して字形の境界を実測 / Measure type from a duplicated outline */
        if (useGlyphBounds === true && pageItem.typename === "TextFrame") {
            var outlineBounds = getOutlineBounds(pageItem, includeStrokeInBounds);
            if (outlineBounds !== null) return outlineBounds;
        }

        return includeStrokeInBounds === true ? pageItem.visibleBounds : pageItem.geometricBounds;
    }

    /**
     * テキストを複製・アウトライン化して字形の境界を取得する（複製は削除する）
     * @param {TextFrame} textFrame - 対象テキスト
     * @param {boolean} includeStrokeInBounds - 線幅を含めるか
     * @returns {number[]|null} [左, 上, 右, 下]。取得できない場合はnull
     */
    function getOutlineBounds(textFrame, includeStrokeInBounds) {
        var duplicated = null;
        var outlined = null;
        var outlineBounds = null;

        try {
            if (textFrame.contents === "") return null;
            duplicated = textFrame.duplicate();
            outlined = duplicated.createOutline();
            outlineBounds = includeStrokeInBounds === true ? outlined.visibleBounds : outlined.geometricBounds;
        } catch (e) {
            outlineBounds = null;
        } finally {
            safeRemove(outlined);
            safeRemove(duplicated);
        }

        return outlineBounds;
    }

    /**
     * オブジェクトを削除する（すでに無い場合は何もしない）
     * @param {PageItem} pageItem - 削除するオブジェクト
     * @returns {void}
     */
    function safeRemove(pageItem) {
        try {
            if (pageItem) pageItem.remove();
        } catch (e) {}
    }

    // =========================================================
    // 移動量の算出 / Offset
    // =========================================================

    /**
     * 揃える方向に応じた移動量を返す
     * @param {Document} documentRef - 対象ドキュメント
     * @param {number[]} selectionBounds - 選択範囲の境界 [左, 上, 右, 下]
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} alignmentSide - 揃える方向
     * @returns {{x: number, y: number}|null} 移動量。方向の指定が不正な場合はnull
     */
    function getAlignmentOffset(documentRef, selectionBounds, artboardRect, alignmentSide) {
        if (alignmentSide === "CENTER_X") {
            return { x: getHorizontalCenterValue(artboardRect) - getHorizontalCenterValue(selectionBounds), y: 0 };
        }

        if (alignmentSide === "CENTER_Y") {
            return { x: 0, y: getVerticalCenterValue(artboardRect) - getVerticalCenterValue(selectionBounds) };
        }

        if (alignmentSide === "CENTER") {
            return {
                x: getHorizontalCenterValue(artboardRect) - getHorizontalCenterValue(selectionBounds),
                y: getVerticalCenterValue(artboardRect) - getVerticalCenterValue(selectionBounds)
            };
        }

        var selectionEdge = getEdgeValueForAlignmentSide(selectionBounds, alignmentSide);
        var artboardEdge = getEdgeValueForAlignmentSide(artboardRect, alignmentSide);
        if (selectionEdge === null || artboardEdge === null) return null;

        /* ガイドが見つかればガイドへ、無ければアートボード端へ / Snap to a guide when found, otherwise the artboard edge */
        var targetEdge = artboardEdge;
        if (USE_GUIDES) {
            var snappedGuideValue = findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide);
            if (snappedGuideValue !== null) targetEdge = snappedGuideValue;
        }

        var edgeOffset = targetEdge - selectionEdge;
        var isHorizontal = (alignmentSide === "left" || alignmentSide === "right");
        return { x: isHorizontal ? edgeOffset : 0, y: isHorizontal ? 0 : edgeOffset };
    }

    /**
     * 指定した方向に対応する境界値を返す
     * @param {number[]} bounds - 境界 [左, 上, 右, 下]
     * @param {string} alignmentSide - 揃える方向
     * @returns {number|null} 境界値。方向の指定が不正な場合はnull
     */
    function getEdgeValueForAlignmentSide(bounds, alignmentSide) {
        if (alignmentSide === "left") return bounds[0];
        if (alignmentSide === "top") return bounds[1];
        if (alignmentSide === "right") return bounds[2];
        if (alignmentSide === "bottom") return bounds[3];
        return null;
    }

    /**
     * 左右中央の座標を返す
     * @param {number[]} bounds - 境界 [左, 上, 右, 下]
     * @returns {number} 左右中央のX座標
     */
    function getHorizontalCenterValue(bounds) {
        return (bounds[0] + bounds[2]) / 2;
    }

    /**
     * 上下中央の座標を返す
     * @param {number[]} bounds - 境界 [左, 上, 右, 下]
     * @returns {number} 上下中央のY座標
     */
    function getVerticalCenterValue(bounds) {
        return (bounds[1] + bounds[3]) / 2;
    }

    // =========================================================
    // ガイドの探索 / Guides
    // =========================================================

    /**
     * アートボード内のガイドから、指定した方向と探索条件に合う吸着先座標を返す
     * @param {Document} documentRef - 対象ドキュメント
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {number} selectionEdge - 選択範囲の端の座標
     * @param {string} alignmentSide - 揃える方向
     * @returns {number|null} 吸着先の座標。該当するガイドが無い場合はnull
     */
    function findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide) {
        var nearestGuideValue = null;
        var nearestGuideDistance = null;
        var guidePathItems = documentRef.pathItems;

        for (var guidePathIndex = 0; guidePathIndex < guidePathItems.length; guidePathIndex++) {
            var guidePathItem = guidePathItems[guidePathIndex];
            if (guidePathItem.guides !== true) continue;

            var guideValue = getGuideValueForAlignmentSide(guidePathItem.geometricBounds, alignmentSide);
            if (guideValue === null) continue;
            if (!isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide)) continue;

            if (GUIDE_SEARCH_MODE === "inside") {
                if (!isGuideOnAlignmentSide(guideValue, selectionEdge, alignmentSide)) continue;

                if (nearestGuideValue === null || isGuideCloserFromInside(guideValue, nearestGuideValue, alignmentSide)) {
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
     * 揃える方向に対応するガイドの座標を返す
     * @param {number[]} guideBounds - ガイドの境界 [左, 上, 右, 下]
     * @param {string} alignmentSide - 揃える方向
     * @returns {number|null} ガイドの座標。方向が対応しない場合はnull
     */
    function getGuideValueForAlignmentSide(guideBounds, alignmentSide) {
        var isVerticalGuide = Math.abs(guideBounds[2] - guideBounds[0]) <= GUIDE_ORIENTATION_TOLERANCE;
        var isHorizontalGuide = Math.abs(guideBounds[1] - guideBounds[3]) <= GUIDE_ORIENTATION_TOLERANCE;

        if ((alignmentSide === "left" || alignmentSide === "right") && isVerticalGuide) return guideBounds[0];
        if ((alignmentSide === "top" || alignmentSide === "bottom") && isHorizontalGuide) return guideBounds[1];
        return null;
    }

    /**
     * ガイドの座標がアクティブアートボードの内側にあるかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} 内側にあればtrue
     */
    function isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide) {
        if (alignmentSide === "left" || alignmentSide === "right") {
            return guideValue >= artboardRect[0] && guideValue <= artboardRect[2];
        }

        if (alignmentSide === "top" || alignmentSide === "bottom") {
            return guideValue <= artboardRect[1] && guideValue >= artboardRect[3];
        }

        return false;
    }

    /**
     * GUIDE_SEARCH_MODE が "inside" のとき、揃える方向側にあるガイドかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number} selectionEdge - 選択範囲の端の座標
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} 揃える方向側にあればtrue
     */
    function isGuideOnAlignmentSide(guideValue, selectionEdge, alignmentSide) {
        if (alignmentSide === "left") return guideValue < selectionEdge;
        if (alignmentSide === "right") return guideValue > selectionEdge;
        if (alignmentSide === "top") return guideValue > selectionEdge;
        if (alignmentSide === "bottom") return guideValue < selectionEdge;
        return false;
    }

    /**
     * "inside" 側にある複数のガイドのうち、選択範囲に近い方かを返す
     * @param {number} guideValue - 比較するガイドの座標
     * @param {number} currentBestGuideValue - 現時点で最も近いガイドの座標
     * @param {string} alignmentSide - 揃える方向
     * @returns {boolean} 近い方であればtrue
     */
    function isGuideCloserFromInside(guideValue, currentBestGuideValue, alignmentSide) {
        if (alignmentSide === "left") return guideValue > currentBestGuideValue;
        if (alignmentSide === "right") return guideValue < currentBestGuideValue;
        if (alignmentSide === "top") return guideValue < currentBestGuideValue;
        if (alignmentSide === "bottom") return guideValue > currentBestGuideValue;
        return false;
    }

    main(scriptArguments[0]);

})(typeof arguments !== "undefined" ? arguments : []);
