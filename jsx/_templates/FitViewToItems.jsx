#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

オブジェクトが指定した割合でウィンドウに収まるよう、表示位置と倍率を合わせる再利用テンプレートです。
「□画面にフィット［65］%」の行を追加でき、キャンセル時に元の表示へ戻す関数もそろっています。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitViewToItems.md

### Overview

A reusable template that centers the view on objects and zooms so they fill a given share of the window.
It can add a "Fit to Window [65] %" row, and includes helpers to restore the original view on Cancel.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitViewToItems.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitViewToItems";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitViewToItems.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitViewToItems.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

var FitViewToItems = (function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ウィンドウに対して対象が占める割合の既定値（%）と範囲。呼び出し側で上書きできる
       Default share of the window the items fill, in percent, and its range; callers can override the default */
    var DEFAULT_FIT_PERCENT = 65;
    var FIT_PERCENT_RANGE = [10, 100];

    /* Illustratorが受け付ける表示倍率の範囲（3.125%〜6400%） / Zoom range Illustrator accepts */
    var VIEW_ZOOM_RANGE = [0.03125, 64];

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var LABELS = {
        checkbox: {
            fitView: { ja: "画面にフィット", en: "Fit to Window" }
        },
        tooltip: {
            fitView: {
                ja: "作成するオブジェクトが収まるよう表示倍率を合わせます。",
                en: "Refits the view to the objects being created."
            },
            fitViewPercent: {
                ja: "ウィンドウに対するオブジェクトの大きさ（100%でいっぱい）",
                en: "Size of the objects relative to the window; 100% fills it"
            }
        }
    };

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    /**
     * LABELS の組から指定言語の文言を返す
     * @param {object} labelSet - { ja: string, en: string } の組
     * @param {string} uiLang - "ja" または "en"
     * @returns {string} 文言（無ければ英語）
     */
    function getLabel(labelSet, uiLang) {
        return (labelSet[uiLang] != null) ? labelSet[uiLang] : labelSet.en;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 数値を範囲に収める。数値として読めないときは既定値を返す
     * @param {string|number} value - 入力値
     * @param {number[]} range - [下限, 上限]
     * @param {number} fallbackValue - 読めないときの既定値
     * @returns {number} 範囲内の数値
     */
    function clampNumber(value, range, fallbackValue) {
        var numberValue = Number(value);
        if (isNaN(numberValue) || (typeof value === "string" && !/\S/.test(value))) numberValue = fallbackValue;
        return Math.min(range[1], Math.max(range[0], numberValue));
    }

    /**
     * 「□画面にフィット［65］%」の行を追加する（ラベルとツールチップは内蔵）
     * @param {Group|Panel|Window} parentContainer - 追加先のコンテナ
     * @param {object} [rowOptions] - value: チェックの初期値（既定 false）／percent: 割合の初期値／lang: 表示言語
     * @returns {{row: Group, checkbox: Checkbox, percentInput: EditText, getFillRatio: function, updateEnabled: function}} 作成したコントロール一式
     */
    function addControls(parentContainer, rowOptions) {
        if (!rowOptions) rowOptions = {};
        var uiLang = rowOptions.lang || getCurrentLang();

        var fitViewRow = parentContainer.add("group");
        fitViewRow.orientation = "row";
        fitViewRow.alignChildren = ["left", "center"];
        fitViewRow.spacing = 6;

        var fitViewCheck = fitViewRow.add("checkbox", undefined, getLabel(LABELS.checkbox.fitView, uiLang));
        fitViewCheck.helpTip = getLabel(LABELS.tooltip.fitView, uiLang);
        /* 明示的に true を渡したときだけONで始める / only an explicit true starts it checked */
        fitViewCheck.value = (rowOptions.value === true);

        var fitPercent = (rowOptions.percent > 0) ? rowOptions.percent : DEFAULT_FIT_PERCENT;
        var percentInput = fitViewRow.add("edittext", undefined, String(fitPercent));
        percentInput.characters = 3;
        percentInput.helpTip = getLabel(LABELS.tooltip.fitViewPercent, uiLang);
        var percentUnitLabel = fitViewRow.add("statictext", undefined, "%");

        var controls = {
            row: fitViewRow,
            checkbox: fitViewCheck,
            percentInput: percentInput,

            /**
             * 入力欄の割合を 0〜1 の比率で返す
             * @returns {number} ウィンドウに対して占める割合（1でいっぱい）
             */
            getFillRatio: function () {
                return clampNumber(percentInput.text, FIT_PERCENT_RANGE, fitPercent) / 100;
            },

            /**
             * 割合の入力欄をチェックの状態に合わせて有効・無効にする
             * @returns {void}
             */
            updateEnabled: function () {
                percentInput.enabled = fitViewCheck.value;
                percentUnitLabel.enabled = fitViewCheck.value;
            }
        };
        controls.updateEnabled();
        return controls;
    }

    /**
     * 複数アイテムを囲む外接範囲を求める（効果を含まない geometricBounds）
     * @param {PageItem[]} targetItems - 対象アイテム
     * @returns {number[]|null} [left, top, right, bottom]（求められない場合は null）
     */
    function getItemsBounds(targetItems) {
        var unionBounds = null;
        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds = targetItems[i].geometricBounds;
            if (unionBounds === null) {
                unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 対象が指定の割合でウィンドウに収まるよう、中心を合わせて表示倍率を変える
     * 拡大・縮小のどちらも行う（KeepInView と違い、常に同じ大きさに見せる）
     * @param {PageItem[]} targetItems - 対象アイテム
     * @param {object} [fitOptions] - doc: 対象ドキュメント（省略時は最前面）／fillRatio: 占める割合（1でいっぱい）
     * @returns {boolean} 表示を動かしたら true
     */
    function fit(targetItems, fitOptions) {
        if (!targetItems || targetItems.length === 0) return false;
        if (!fitOptions) fitOptions = {};

        var targetDoc = fitOptions.doc || app.activeDocument;
        var bounds = getItemsBounds(targetItems);
        if (bounds === null) return false;

        var itemWidth = bounds[2] - bounds[0];
        var itemHeight = bounds[1] - bounds[3];
        var activeView = targetDoc.activeView;
        activeView.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        if (itemWidth <= 0 || itemHeight <= 0) return true;

        /* 中心をそろえたあとの表示範囲を基準に倍率を求める / Scale from the view bounds after the center has moved */
        var fillRatio = (fitOptions.fillRatio > 0) ? fitOptions.fillRatio : DEFAULT_FIT_PERCENT / 100;
        var viewBounds = activeView.bounds;
        var scale = Math.min(
            (viewBounds[2] - viewBounds[0]) / itemWidth,
            (viewBounds[1] - viewBounds[3]) / itemHeight
        ) * fillRatio;
        activeView.zoom = clampNumber(activeView.zoom * scale, VIEW_ZOOM_RANGE, 1);
        return true;
    }

    /**
     * 現在の表示位置と倍率を控える（キャンセル時に restoreView で戻す）
     * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
     * @returns {{centerPoint: number[], zoom: number}} 控えた表示状態
     */
    function captureView(targetDoc) {
        var activeView = (targetDoc || app.activeDocument).activeView;
        return { centerPoint: activeView.centerPoint, zoom: activeView.zoom };
    }

    /**
     * captureView で控えた表示位置と倍率に戻す
     * @param {{centerPoint: number[], zoom: number}} viewState - 控えた表示状態
     * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
     * @returns {void}
     */
    function restoreView(viewState, targetDoc) {
        if (!viewState) return;
        var activeView = (targetDoc || app.activeDocument).activeView;
        activeView.centerPoint = viewState.centerPoint;
        activeView.zoom = viewState.zoom;
    }

    return {
        addControls: addControls,
        fit: fit,
        captureView: captureView,
        restoreView: restoreView,
        getItemsBounds: getItemsBounds
    };

})();
