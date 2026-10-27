#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.PatternSwatchSquares
Name=PatternSwatchSquares
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-10-27
Description-BEGIN
200 px 四方の長方形を描き、パターンスウォッチを適用します。
スウォッチパネルで選択したパターン（選択が無ければドキュメントのすべてのパターン）ごとに1つずつ、横に並べて作ります。

Draws 200 px squares and fills them with pattern swatches.
One square is created for each pattern selected in the Swatches panel (or every pattern in the document if none is selected), laid out in a row.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

200 px 四方の長方形を描き、パターンスウォッチを適用します。
スウォッチパネルで選択したパターン（選択が無ければドキュメントのすべてのパターン）ごとに1つずつ、横に並べて作ります。

### Overview

Draws 200 px squares and fills them with pattern swatches.
One square is created for each pattern selected in the Swatches panel (or every pattern in the document if none is selected), laid out in a row.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PatternSwatchSquares";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-27";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 設定 / Settings
    // =========================================
    var SQUARE_SIZE = 200;   /* 長方形の幅・高さ（pt = px） / Square width and height */
    var SQUARE_GAP = 40;     /* 長方形どうしの間隔 / Gap between squares */
    var STROKE_WIDTH = 2;    /* 線幅 / Stroke width */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noPattern: { ja: "パターンスウォッチがありません", en: "There are no pattern swatches in the document." }
        }
    };

    /**
     * LABELS から今の UI 言語の文言を取り出す
     * @param {string} labelPath - "alert.noDocument" のようなパス
     * @returns {string} 文言
     */
    function getLabel(labelPath) {
        var labelEntry = LABELS;
        var labelPathKeys = labelPath.split(".");
        for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
            labelEntry = labelEntry[labelPathKeys[i]];
        }
        return (labelEntry != null) ? (labelEntry[uiLang] || labelEntry.en) : labelPath;
    }

    // =========================================
    // パターンの取得 / Collecting patterns
    // =========================================

    /**
     * スウォッチの配列から、パターンスウォッチだけを抜き出す
     * @param {Swatches|Swatch[]} swatchList - スウォッチの一覧
     * @returns {Swatch[]} パターンスウォッチ
     */
    function filterPatternSwatches(swatchList) {
        var patternSwatches = [];
        for (var i = 0; i < swatchList.length; i++) {
            if (swatchList[i].color.typename === "PatternColor") {
                patternSwatches.push(swatchList[i]);
            }
        }
        return patternSwatches;
    }

    /**
     * 対象のパターンスウォッチを返す（スウォッチパネルの選択を優先し、無ければすべて）
     * @param {Document} doc - 対象のドキュメント
     * @returns {Swatch[]} パターンスウォッチ
     */
    function getTargetPatternSwatches(doc) {
        var selectedPatterns = filterPatternSwatches(doc.swatches.getSelected());
        if (selectedPatterns.length > 0) return selectedPatterns;
        return filterPatternSwatches(doc.swatches);
    }

    // =========================================
    // 長方形の作成 / Drawing squares
    // =========================================

    /**
     * ドキュメントのカラーモードに合わせた黒を返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {CMYKColor|RGBColor} 黒
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykBlack = new CMYKColor();
            cmykBlack.black = 100;
            return cmykBlack;
        }
        return new RGBColor(); /* 既定値が 0,0,0 / Defaults to 0,0,0 */
    }

    /**
     * パターンを塗りにした正方形を描く
     * @param {Document} doc - 対象のドキュメント
     * @param {number} left - 左端の X 座標
     * @param {number} top - 上端の Y 座標
     * @param {Swatch} patternSwatch - 適用するパターンスウォッチ
     * @returns {PathItem} 描いた正方形
     */
    function drawPatternSquare(doc, left, top, patternSwatch) {
        var square = doc.activeLayer.pathItems.rectangle(top, left, SQUARE_SIZE, SQUARE_SIZE);
        square.filled = true;
        square.fillColor = patternSwatch.color;
        square.stroked = true;
        square.strokeColor = createBlackColor(doc);
        square.strokeWidth = STROKE_WIDTH;
        return square;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var patternSwatches = getTargetPatternSwatches(doc);
        if (patternSwatches.length === 0) {
            alert(getLabel("alert.noPattern"));
            return;
        }

        /* 並び全体を表示中の画面の中央に置く / Center the whole row in the current view */
        var rowWidth = patternSwatches.length * SQUARE_SIZE + (patternSwatches.length - 1) * SQUARE_GAP;
        var viewCenter = doc.activeView.centerPoint;
        var rowLeft = viewCenter[0] - rowWidth / 2;
        var rowTop = viewCenter[1] + SQUARE_SIZE / 2;

        var createdSquares = [];
        for (var i = 0; i < patternSwatches.length; i++) {
            var squareLeft = rowLeft + i * (SQUARE_SIZE + SQUARE_GAP);
            createdSquares.push(drawPatternSquare(doc, squareLeft, rowTop, patternSwatches[i]));
        }

        doc.selection = null;
        doc.selection = createdSquares;
    }

    main();
})();
