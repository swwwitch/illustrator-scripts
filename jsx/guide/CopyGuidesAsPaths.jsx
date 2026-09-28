#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブアートボードのガイドを線付きのパスに変換してクリップボードへ送ります。
元のガイドは残し、同じ座標に重なった余分なガイドだけを削除します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyGuidesAsPaths.md

### Overview

Converts the guides on the active artboard into stroked paths and sends them to the clipboard.
The original guides are kept, while duplicated guides stacked at the same position are removed.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyGuidesAsPaths.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyGuidesAsPaths";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-07";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyGuidesAsPaths.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyGuidesAsPaths.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================
var STROKE_WIDTH    = 0.5;                       /* 変換後のパスに付ける線幅（pt） */
var MATCH_TOLERANCE = 0.001;                     /* 同じ座標とみなす誤差（pt） */
var TEMP_LAYER_NAME = "__guide_to_path__";       /* 作業用レイヤー名 */

// =========================================
// ラベル定義 / Labels
// =========================================
var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noGuides: { ja: "対象となるガイドがありません。", en: "No guides found on the active artboard." },
        copied: { ja: "%1本のガイドをコピーしました。", en: "Copied %1 guide(s)." }
    }
};

// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
// ローカライズ（再利用パーツ） / Localization (reusable)
//
// 【移植手順 / How to port】
// 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
//    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
// 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
//    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
// 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
//      getLabel("dialog.title")        … パス
//      getLabel(LABELS.dialog.title)   … { ja, en } を直接
//      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
//      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
//      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
//      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
// 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
// ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
// ローカライズ（再利用パーツ）ここまで / End of the reusable localization
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

(function () {
    /**
     * ドキュメントのカラーモードに合わせた黒の線色を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} RGBColor または CMYKColor
     */
    function createStrokeColor(doc) {
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
     * アートボード内にあるガイドを集める（ロックされたオブジェクト・レイヤーも対象）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @returns {PathItem[]} 対象のガイド
     */
    function collectGuidesOnArtboard(doc, artboardRect) {
        var left = artboardRect[0];
        var top = artboardRect[1];
        var right = artboardRect[2];
        var bottom = artboardRect[3];
        var pathItems = doc.pathItems;
        var itemCount = pathItems.length;
        var result = [];

        for (var i = 0; i < itemCount; i++) {
            var item = pathItems[i];

            if (!item.guides) {
                continue;
            }

            /* ロックは無視し、非表示のものだけ除外 / Ignore lock, skip hidden only */
            if (item.hidden || !item.layer.visible) {
                continue;
            }

            var bounds = item.geometricBounds;
            var centerX = (bounds[0] + bounds[2]) / 2;
            var centerY = (bounds[1] + bounds[3]) / 2;

            if (
                centerX >= left &&
                centerX <= right &&
                centerY <= top &&
                centerY >= bottom
            ) {
                result.push(item);
            }
        }

        return result;
    }

    /**
     * 2本のパスが同じ座標かどうかを判定する（描画方向の違いも同一とみなす）
     * @param {PathItem} pathA - 比較するパス
     * @param {PathItem} pathB - 比較するパス
     * @param {number} tolerance - 同一とみなす誤差（pt）
     * @returns {boolean} 同じ座標なら true
     */
    function isSameGeometry(pathA, pathB, tolerance) {
        var pointsA = pathA.pathPoints;
        var pointsB = pathB.pathPoints;

        if (pathA.closed !== pathB.closed || pointsA.length !== pointsB.length) {
            return false;
        }

        var pointCount = pointsA.length;
        var forwardMatched = true;
        var reversedMatched = true;

        for (var i = 0; i < pointCount; i++) {
            var anchorA = pointsA[i].anchor;

            if (forwardMatched) {
                var anchorForward = pointsB[i].anchor;
                if (
                    Math.abs(anchorA[0] - anchorForward[0]) > tolerance ||
                    Math.abs(anchorA[1] - anchorForward[1]) > tolerance
                ) {
                    forwardMatched = false;
                }
            }

            if (reversedMatched) {
                var anchorReversed = pointsB[pointCount - 1 - i].anchor;
                if (
                    Math.abs(anchorA[0] - anchorReversed[0]) > tolerance ||
                    Math.abs(anchorA[1] - anchorReversed[1]) > tolerance
                ) {
                    reversedMatched = false;
                }
            }

            if (!forwardMatched && !reversedMatched) {
                return false;
            }
        }

        return true;
    }

    /**
     * ロックを一時的に解除してページアイテムを削除する
     * @param {PageItem} item - 削除するアイテム
     * @returns {void}
     */
    function removeLockedItem(item) {
        var parentLayer = item.layer;
        var layerWasLocked = parentLayer.locked;

        if (layerWasLocked) {
            parentLayer.locked = false;
        }

        if (item.locked) {
            item.locked = false;
        }

        item.remove();

        if (layerWasLocked) {
            parentLayer.locked = true;
        }
    }

    /**
     * 同じ座標に重なったガイドを1本だけ残し、残りをドキュメントから削除する
     * @param {PathItem[]} items - 対象のガイド
     * @param {number} tolerance - 同一とみなす誤差（pt）
     * @returns {PathItem[]} 重複を取り除いたガイド
     */
    function removeDuplicatedGuides(items, tolerance) {
        var kept = [];

        for (var i = 0; i < items.length; i++) {
            var isDuplicated = false;

            for (var j = 0; j < kept.length; j++) {
                if (isSameGeometry(items[i], kept[j], tolerance)) {
                    isDuplicated = true;
                    break;
                }
            }

            if (isDuplicated) {
                removeLockedItem(items[i]);
            } else {
                kept.push(items[i]);
            }
        }

        return kept;
    }

    /**
     * ガイドの形状を写し取り、線付きの通常パスとして作業用レイヤーに作る
     * @param {PathItem} source - 元のガイド
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @param {Object} strokeColor - 適用する線色
     * @param {number} strokeWidth - 適用する線幅（pt）
     * @returns {PathItem} 作成した通常パス
     */
    function createStrokedCopy(source, targetLayer, strokeColor, strokeWidth) {
        var newPath = targetLayer.pathItems.add();
        var sourcePoints = source.pathPoints;
        var pointCount = sourcePoints.length;

        for (var i = 0; i < pointCount; i++) {
            var sourcePoint = sourcePoints[i];
            var newPoint = newPath.pathPoints.add();
            newPoint.anchor = sourcePoint.anchor;
            newPoint.leftDirection = sourcePoint.leftDirection;
            newPoint.rightDirection = sourcePoint.rightDirection;
            newPoint.pointType = sourcePoint.pointType;
        }

        newPath.closed = source.closed;
        newPath.filled = false;
        newPath.stroked = true;
        newPath.strokeWidth = strokeWidth;
        newPath.strokeColor = strokeColor;

        return newPath;
    }

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var targetGuides = collectGuidesOnArtboard(doc, artboard.artboardRect);

    if (targetGuides.length === 0) {
        alert(getLabel(LABELS.alert.noGuides));
        return;
    }

    /* 同じ座標に重なったガイドを削除 / Remove guides stacked at the same position */
    targetGuides = removeDuplicatedGuides(targetGuides, MATCH_TOLERANCE);

    /* 作業用レイヤーを作り、そこに線付きパスを作成 / Build stroked paths on a temp layer */
    var previousActiveLayer = doc.activeLayer;
    var tempLayer = doc.layers.add();
    tempLayer.name = TEMP_LAYER_NAME;

    var strokeColor = createStrokeColor(doc);
    var createdPaths = [];

    for (var i = 0; i < targetGuides.length; i++) {
        createdPaths.push(
            createStrokedCopy(targetGuides[i], tempLayer, strokeColor, STROKE_WIDTH)
        );
    }

    /* 作成したパスだけを選択 / Select only the new paths */
    app.executeMenuCommand("deselectall");

    for (var j = 0; j < createdPaths.length; j++) {
        createdPaths[j].selected = true;
    }

    /* クリップボードへカット / Cut to clipboard */
    app.redraw();                   /* 作成直後は再描画しないとカット対象にならない */
    app.executeMenuCommand("cut");  /* app.cut() は黙って無視されることがある */
    app.redraw();

    /* 作業用レイヤーを片付け、元の状態に戻す / Clean up the temp layer */
    tempLayer.remove();
    doc.activeLayer = previousActiveLayer;

    alert(getLabel(LABELS.alert.copied, [targetGuides.length]));
})();
