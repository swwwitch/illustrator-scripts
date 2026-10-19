#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.CurrentUnitsState
Name=CurrentUnitsState
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-10-18
Description-BEGIN
環境設定の［単位］→［一般］［線］［文字］［東アジア言語のオプション］の現在の単位を読み出して表示します。

Shows the current units in Units > General, Stroke, Type and East Asian Type.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

環境設定の［単位］→［一般］［線］［文字］［東アジア言語のオプション］の現在の単位を読み出して表示します。

### Overview

Shows the current units in Units > General, Stroke, Type and East Asian Type.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CurrentUnitsState";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-18";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 読み出す項目 / Items to read
    // =========================================

    /* 環境設定キーと表示名のキー / Preference keys and their label keys */
    var UNIT_ITEMS = [
        { prefKey: "rulerType", labelKey: "general" },
        { prefKey: "strokeUnits", labelKey: "stroke" },
        { prefKey: "text/units", labelKey: "type" },
        { prefKey: "text/asianunits", labelKey: "asian" }
    ];

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

    var LABELS = {
        item: {
            general: { ja: "一般", en: "General" },
            stroke: { ja: "線", en: "Stroke" },
            type: { ja: "文字", en: "Type" },
            asian: { ja: "東アジア言語のオプション", en: "East Asian Type" }
        },
        alert: {
            title: { ja: "単位", en: "Units" }
        }
    };

    /**
     * { ja, en } から現在の言語の文言を返す
     * @param {Object} entry - { ja, en }
     * @returns {string} 文言
     */
    function getLabel(entry) {
        return entry ? (entry[uiLang] || entry.en || "") : "";
    }

    // =========================================
    // 処理 / Logic
    // =========================================

    /**
     * 単位を読み出して alert で表示する
     */
    function main() {
        var colon = (uiLang === "ja") ? " : " : ": ";
        var lines = [];
        for (var i = 0; i < UNIT_ITEMS.length; i++) {
            var item = UNIT_ITEMS[i];
            var unitInfo = getUnitInfo(item.prefKey);
            var codeText = (uiLang === "ja") ? "（" + unitInfo.code + "）" : " (" + unitInfo.code + ")";
            lines.push(getLabel(LABELS.item[item.labelKey]) + colon + unitInfo.label + codeText);
        }
        alert(lines.join("\n"), getLabel(LABELS.alert.title));
    }

    main();

})();
