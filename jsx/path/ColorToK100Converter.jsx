#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

RGBまたはCMYKで構成された黒を、安定したK100の黒に変換します。
テキスト、パス、スウォッチの塗りおよび線カラーが一括の対象です。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorToK100Converter.md

### Overview

Converts blacks built from RGB or CMYK into a stable K100 black.
Text, paths and swatches are all covered, for both fill and stroke colors.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorToK100Converter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ColorToK100Converter";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorToK100Converter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorToK100Converter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 置き換え先にするスウォッチ名（無ければ K100 を新規作成） / Swatch used as K100; a new K100 is built when missing */
    var K100_SWATCH_NAME = "ブラック";

    /* 黒と判定する閾値 / Thresholds for treating a color as black */
    var BLACK_THRESHOLDS = {
        rgb: { red: 39, green: 39, blue: 39 },       /* RGB がすべてこれ未満 / All RGB below these */
        cmyk: { cyan: 70, magenta: 70, yellow: 70 }, /* CMY がすべてこれ以上 / All CMY at or above these */
        cmykAll: 60,                                 /* CMYK がすべてこれ以上 / All CMYK at or above this */
        cmykTotal: 310,                              /* CMYK の合計がこれ以上 / CMYK total at or above this */
        solidBlackMaxCmy: 10                         /* 墨ベタとして許容する CMY の上限 / Max CMY kept as solid black */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
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

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            rgbDocument: {
                ja: "このスクリプトはCMYKカラーモードのドキュメントでのみ使用できます。",
                en: "This script works only in CMYK color mode documents."
            }
        }
    };

    // =========================================
    // 色の変換 / Color conversion
    // =========================================

    /**
     * 置き換え先の K100 カラーを返す（スウォッチが CMYK ならその色、無ければ K100 を新規作成）
     * @param {Swatch|null} k100Swatch - 置き換え先のスウォッチ
     * @returns {CMYKColor} K100 のカラー
     */
    function getK100Color(k100Swatch) {
        if (k100Swatch && k100Swatch.color.typename === "CMYKColor") {
            return k100Swatch.color;
        }
        var k100Color = new CMYKColor();
        k100Color.cyan = 0;
        k100Color.magenta = 0;
        k100Color.yellow = 0;
        k100Color.black = 100;
        return k100Color;
    }

    /**
     * RGB の黒か判定する（すべての値が閾値未満）
     * @param {RGBColor} color - 判定するカラー
     * @returns {boolean} 黒なら true
     */
    function isRgbBlack(color) {
        var limits = BLACK_THRESHOLDS.rgb;
        return color.red < limits.red && color.green < limits.green && color.blue < limits.blue;
    }

    /**
     * CMYK のリッチブラックか判定する（墨ベタは除く）
     * @param {CMYKColor} color - 判定するカラー
     * @returns {boolean} K100 に変換する対象なら true
     */
    function isCmykRichBlack(color) {
        var c = color.cyan;
        var m = color.magenta;
        var y = color.yellow;
        var k = color.black;
        var maxCmy = BLACK_THRESHOLDS.solidBlackMaxCmy;

        /* 墨ベタ（K=100 かつ CMY が小さい）はそのまま許容 / Keep solid black as is */
        if (k === 100 && c <= maxCmy && m <= maxCmy && y <= maxCmy) {
            return false;
        }

        /* CMY がすべて閾値以上／CMYK がすべて閾値以上／合計が閾値以上 / CMY, all-CMYK or total over the thresholds */
        var cmyLimits = BLACK_THRESHOLDS.cmyk;
        var allLimit = BLACK_THRESHOLDS.cmykAll;
        return (c >= cmyLimits.cyan && m >= cmyLimits.magenta && y >= cmyLimits.yellow) ||
            (c >= allLimit && m >= allLimit && y >= allLimit && k >= allLimit) ||
            (c + m + y + k) >= BLACK_THRESHOLDS.cmykTotal;
    }

    /**
     * RGB／CMYK の黒なら K100 に置き換えたカラーを返す（それ以外はそのまま）
     * @param {Color} color - 元のカラー
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {Color} 変換後のカラー
     */
    function convertBlackToK100(color, k100Color) {
        if (color.typename === "RGBColor" && isRgbBlack(color)) {
            return k100Color;
        }
        if (color.typename === "CMYKColor" && isCmykRichBlack(color)) {
            return k100Color;
        }
        return color;
    }

    /**
     * パス（複合パス）の塗りと線を変換する。ロック・非表示は対象外
     * @param {PathItem|CompoundPathItem} pathItem - 対象のパス
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {void}
     */
    function convertPathColors(pathItem, k100Color) {
        if (pathItem.locked || pathItem.hidden) return;
        if (pathItem.filled) pathItem.fillColor = convertBlackToK100(pathItem.fillColor, k100Color);
        if (pathItem.stroked) pathItem.strokeColor = convertBlackToK100(pathItem.strokeColor, k100Color);
    }

    /**
     * テキストの文字ごとに塗りと線を変換する。ロック・非表示は対象外
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {void}
     */
    function convertTextColors(textFrame, k100Color) {
        if (textFrame.locked || textFrame.hidden) return;
        var chars = textFrame.characters;
        for (var i = 0; i < chars.length; i++) {
            chars[i].fillColor = convertBlackToK100(chars[i].fillColor, k100Color);
            chars[i].strokeColor = convertBlackToK100(chars[i].strokeColor, k100Color);
        }
    }

    /**
     * グループ内のパス・複合パス・テキストを再帰的に変換する
     * @param {GroupItem} groupItem - 対象のグループ
     * @param {CMYKColor} k100Color - 置き換え先の K100
     * @returns {void}
     */
    function convertGroupColors(groupItem, k100Color) {
        var i;
        for (i = 0; i < groupItem.pathItems.length; i++) {
            convertPathColors(groupItem.pathItems[i], k100Color);
        }
        for (i = 0; i < groupItem.compoundPathItems.length; i++) {
            convertPathColors(groupItem.compoundPathItems[i], k100Color);
        }
        for (i = 0; i < groupItem.textFrames.length; i++) {
            convertTextColors(groupItem.textFrames[i], k100Color);
        }
        for (i = 0; i < groupItem.groupItems.length; i++) {
            convertGroupColors(groupItem.groupItems[i], k100Color);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメント内のテキスト・パス・複合パス・グループ・スウォッチの黒を K100 に変換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        /* RGB ドキュメントでは中断 / Stop on RGB documents */
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            alert(getLabel("alert.rgbDocument"));
            return;
        }

        /* getByName は見つからないと例外 / getByName throws when the swatch is missing */
        var k100Swatch = null;
        try {
            k100Swatch = doc.swatches.getByName(K100_SWATCH_NAME);
        } catch (e) { }
        var k100Color = getK100Color(k100Swatch);

        var i;
        for (i = 0; i < doc.textFrames.length; i++) {
            convertTextColors(doc.textFrames[i], k100Color);
        }
        for (i = 0; i < doc.pathItems.length; i++) {
            convertPathColors(doc.pathItems[i], k100Color);
        }
        for (i = 0; i < doc.compoundPathItems.length; i++) {
            convertPathColors(doc.compoundPathItems[i], k100Color);
        }
        for (i = 0; i < doc.groupItems.length; i++) {
            convertGroupColors(doc.groupItems[i], k100Color);
        }

        /* スウォッチのカラー定義 / Swatch color definitions */
        for (i = 0; i < doc.swatches.length; i++) {
            var swatchColor = doc.swatches[i].color;
            if (swatchColor.typename === "RGBColor" || swatchColor.typename === "CMYKColor") {
                doc.swatches[i].color = convertBlackToK100(swatchColor, k100Color);
            }
        }
    }

    main();

})();
