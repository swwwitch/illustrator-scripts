#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトに［効果］＞［スタイライズ］＞［ぼかし］のライブエフェクトを適用します。
半径は現在の定規単位で入力し、内部で pt に換算します。

詳細は README を参照してください。

### Overview

Applies the Effect > Stylize > Feather live effect to the current selection.
The radius is entered in the current ruler unit and converted to points internally.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LEFeather";                    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LEFeather.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LEFeather.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/*
選択オブジェクトに「効果 > スタイライズ > ぼかし」ライブエフェクトを適用する。
半径はダイアログから入力し、現在の定規単位（rulerType）に合わせて表示・換算する。
*/

/*
Apply the Effect > Stylize > Feather live effect to the current selection.
The radius input matches the current ruler unit (rulerType) and is converted to points internally.
*/

(function () {

    // ================================================================================
    // バージョンとローカライズ / Version and localization
    // ================================================================================

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    var LABELS = {
        dialogTitle: { ja: "ぼかし（スタイライズ）", en: "Feather" },
        radius:      { ja: "半径",            en: "Radius" },
        cancel:      { ja: "キャンセル",      en: "Cancel" },
        noDoc:       { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSel:       { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." },
        invalid:     { ja: "半径には正の数値を入力してください。", en: "Please enter a positive number for the radius." }
    };

    function L(key) {
        return LABELS[key][lang];
    }

    function labelText(key) {
        return L(key) + (lang === "ja" ? "：" : ":");
    }

    // ================================================================================
    // 設定 / Settings
    // ================================================================================

    var RULER_UNITS = [
        { label: "inch", factor: 72.0 },
        { label: "mm",   factor: 72.0 / 25.4 },
        { label: "pt",   factor: 1.0 },
        { label: "pica", factor: 12.0 },
        { label: "cm",   factor: 72.0 / 2.54 },
        { label: "Q",    factor: 72.0 / 25.4 * 0.25 },
        { label: "px",   factor: 1.0 }
    ];

    /* 現在の定規単位 / Current ruler unit */
    function getRulerUnit() {
        var rulerType = app.preferences.getIntegerPreference("rulerType");
        return RULER_UNITS[rulerType] || RULER_UNITS[2];
    }

    // ================================================================================
    // ヘルパー関数 / Helper functions
    // ================================================================================

    /* ぼかし（スタイライズ）を適用 / Apply Stylize > Feather */
    function applyFeather(item, radiusPt) {
        var xml = '<LiveEffect name="Adobe Fuzzy Mask"><Dict data="R Radius ' + radiusPt + ' "/></LiveEffect>';
        item.applyEffect(xml);
    }

    /* 文字列を正の数値に変換、不正なら null / Parse to positive number or null */
    function parsePositive(text) {
        var value = parseFloat(text);
        if (isNaN(value) || value <= 0) return null;
        return value;
    }

    // ================================================================================
    // ダイアログ / Dialog
    // ================================================================================

    function promptRadius(rulerUnit, defaultValueInUnit) {
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;

        var radiusRow = dialog.add("group");
        radiusRow.orientation = "row";
        radiusRow.alignChildren = "center";
        radiusRow.add("statictext", undefined, labelText("radius"));
        var radiusInput = radiusRow.add("edittext", undefined, String(defaultValueInUnit));
        radiusInput.characters = 8;
        radiusRow.add("statictext", undefined, rulerUnit.label);

        var buttonRow = dialog.add("group");
        buttonRow.alignment = "right";
        buttonRow.add("button", undefined, L("cancel"), { name: "cancel" });
        buttonRow.add("button", undefined, "OK", { name: "ok" });

        radiusInput.active = true;

        if (dialog.show() !== 1) return null;

        var value = parsePositive(radiusInput.text);
        if (value === null) {
            alert(L("invalid"));
            return null;
        }
        return value;
    }

    // ================================================================================
    // メイン処理 / Main flow
    // ================================================================================

    if (app.documents.length === 0) {
        alert(L("noDoc"));
        return;
    }

    if (app.selection.length === 0) {
        alert(L("noSel"));
        return;
    }

    var rulerUnit = getRulerUnit();
    var defaultPt = 10;
    var defaultInUnit = Math.round((defaultPt / rulerUnit.factor) * 100) / 100;

    var radiusInUnit = promptRadius(rulerUnit, defaultInUnit);
    if (radiusInUnit === null) return;

    var radiusPt = radiusInUnit * rulerUnit.factor;

    for (var i = 0; i < app.selection.length; i++) {
        applyFeather(app.selection[i], radiusPt);
    }

    app.redraw();

})();
