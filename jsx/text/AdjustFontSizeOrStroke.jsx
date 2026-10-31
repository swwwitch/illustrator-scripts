#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "AdjustFontSizeOrStrokeEngine"

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.AdjustFontSizeOrStroke
Name=テキストは文字サイズ、それ以外は線幅を一定量ずつ増減
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-11-01
Description-BEGIN
選択したテキストは文字サイズを、それ以外のオブジェクトは線幅を、環境設定の単位で一定量ずつ増減します（ダイアログボックスは出しません）。

Steps the font size of selected text, or the stroke weight of other selected objects, by a fixed amount in the units set in Preferences (no dialog box).
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したテキストは文字サイズを、それ以外のオブジェクトは線幅を、環境設定の単位で一定量ずつ増減します（ダイアログボックスは出しません）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSizeOrStroke.md

### Overview

Steps the font size of selected text, or the stroke weight of other selected objects, by a fixed amount in the units set in Preferences (no dialog box).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSizeOrStroke.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustFontSizeOrStroke";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-01";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "テキストは文字サイズ、それ以外は線幅を一定量ずつ増減"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSizeOrStroke.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSizeOrStroke.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AdjustFontSizeOrStroke.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var STEP_SIGN = 1;              /* 正負の切り替え。1 で大きく・太く、-1 で小さく・細く / 1 to increase, -1 to decrease */
    var STEP_MULTIPLIER = 1;        /* 倍率の切り替え。1 / 5 / 10 / step multiplier: 1, 5 or 10 */

    var FONT_SIZE_STEP = 1;         /* 文字サイズの1回分の量（環境設定の［文字］の単位） / font size step, in the type unit set in Preferences */
    var STROKE_STEP = 0.1;          /* 線幅の1回分の量（環境設定の［線］の単位） / stroke weight step, in the stroke unit set in Preferences */
    var VALUE_DECIMALS = 4;         /* 足したあとの値を丸める桁数（単位の値で） / decimals kept after adding, in the unit */
    var FONT_SIZE_MIN_PT = 0.1;     /* 文字サイズの下限（pt） / minimum font size (pt) */
    var FONT_SIZE_MAX_PT = 1296;    /* 文字サイズの上限（pt） / maximum font size (pt) */
    var STROKE_WIDTH_MIN_PT = 0;    /* 線幅の下限（pt） / minimum stroke weight (pt) */
    var STROKE_WIDTH_MAX_PT = 1000; /* 線幅の上限（pt） / maximum stroke weight (pt) */

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

    /**
     * pt の値を単位の値に換算する
     * @param {number} ptValue - pt の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} 単位の値
     */
    function ptToUnit(ptValue, unit) {
        return ptValue / unit.pointsPerUnit;
    }

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してから実行してください。", en: "Select one or more objects first." },
            nothingChanged: { ja: "選択範囲にテキストも線が設定されたオブジェクトもありません。", en: "The selection contains no text or stroked objects." }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * @typedef {Object} SizeAdjustment
     * @property {{pointsPerUnit: number}} textUnit - 環境設定の［文字］の単位
     * @property {{pointsPerUnit: number}} strokeUnit - 環境設定の［線］の単位
     * @property {number} fontSizeStep - 文字サイズに足す量（textUnit の値。負なら小さく）
     * @property {number} strokeStep - 線幅に足す量（strokeUnit の値。負なら細く）
     */

    /**
     * 選択範囲のテキストは文字サイズを、それ以外のオブジェクトは線幅を、ユーザー設定の量だけ変える
     * （ダイアログボックスは出さない）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返る（カーソルだけなら長さ0） / a TextRange while selecting characters with the Type tool */
        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }
        var stepFactor = STEP_SIGN * STEP_MULTIPLIER;
        var sizeAdjustment = {
            textUnit: getUnitInfo("text/units"),
            strokeUnit: getUnitInfo("strokeUnits"),
            fontSizeStep: FONT_SIZE_STEP * stepFactor,
            strokeStep: STROKE_STEP * stepFactor
        };
        var changedCount = (selectedItems.typename === "TextRange")
            ? adjustFontSizes(selectedItems, sizeAdjustment)
            : adjustItems(selectedItems, sizeAdjustment);
        if (changedCount === 0) alert(getLabel(LABELS.alert.nothingChanged));
    }

    /**
     * 複数のオブジェクトの文字サイズ・線幅を変える
     * @param {Array|PageItems|PathItems} items - 対象のオブジェクト
     * @param {SizeAdjustment} sizeAdjustment - 単位と足す量
     * @returns {number} 変更したパス・文字の数
     */
    function adjustItems(items, sizeAdjustment) {
        var changedCount = 0;
        for (var i = 0; i < items.length; i++) {
            changedCount += adjustItem(items[i], sizeAdjustment);
        }
        return changedCount;
    }

    /**
     * 1つのオブジェクトを変える。テキストは文字サイズ、パスは線幅を変え、グループ・複合パスは中までたどる。線の無いパスは変えない
     * @param {PageItem} item - 対象のオブジェクト
     * @param {SizeAdjustment} sizeAdjustment - 単位と足す量
     * @returns {number} 変更したパス・文字の数
     */
    function adjustItem(item, sizeAdjustment) {
        /* グループの中のロック・非表示のオブジェクトは触らない / leave locked or hidden items inside groups alone */
        if (item.locked || item.hidden) return 0;
        switch (item.typename) {
            case "TextFrame":
                return adjustFontSizes(item.textRange, sizeAdjustment);
            case "PathItem":
                if (!item.stroked) return 0;
                item.strokeWidth = addUnitStep(item.strokeWidth, sizeAdjustment.strokeUnit, sizeAdjustment.strokeStep, STROKE_WIDTH_MIN_PT, STROKE_WIDTH_MAX_PT);
                return 1;
            case "CompoundPathItem":
                return adjustItems(item.pathItems, sizeAdjustment);
            case "GroupItem":
                return adjustItems(item.pageItems, sizeAdjustment);
            default:
                return 0;
        }
    }

    /**
     * テキストの文字サイズを変える。文字ごとにサイズが異なることがあるので1文字ずつ足す
     * @param {TextRange} textRange - 対象のテキスト
     * @param {SizeAdjustment} sizeAdjustment - 単位と足す量
     * @returns {number} サイズを変えた文字の数
     */
    function adjustFontSizes(textRange, sizeAdjustment) {
        var textCharacters = textRange.characters;
        var characterCount = textCharacters.length;
        for (var i = 0; i < characterCount; i++) {
            var characterAttrs = textCharacters[i].characterAttributes;
            characterAttrs.size = addUnitStep(characterAttrs.size, sizeAdjustment.textUnit, sizeAdjustment.fontSizeStep, FONT_SIZE_MIN_PT, FONT_SIZE_MAX_PT);
        }
        return characterCount;
    }

    /**
     * pt の値に、単位の値で step を足した値を返す。単位の値で丸めて、繰り返し足しても端数が残らないようにする
     * @param {number} valuePt - 今の値（pt）
     * @param {{pointsPerUnit: number}} unit - 足す量の単位
     * @param {number} step - 足す量（unit の値。負なら減らす）
     * @param {number} minPt - 下限（pt）
     * @param {number} maxPt - 上限（pt）
     * @returns {number} 足した値（pt。下限・上限の範囲に収める）
     */
    function addUnitStep(valuePt, unit, step, minPt, maxPt) {
        var roundingFactor = Math.pow(10, VALUE_DECIMALS);
        var steppedValue = Math.round((ptToUnit(valuePt, unit) + step) * roundingFactor) / roundingFactor;
        return Math.max(minPt, Math.min(unitToPt(steppedValue, unit), maxPt));
    }

    main();

})();
