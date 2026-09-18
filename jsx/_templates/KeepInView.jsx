#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

生成・変更したオブジェクトが画面から外れたときだけ、見える位置へ表示を移す再利用テンプレートです。
すでに見えているときは動かさず、画面に収まらないときはズームアウトのみ行います。

詳細は README を参照してください。

### Overview

A reusable template that scrolls the view only when the objects you created or changed have fallen outside it.
Objects already in view are left alone, and it only zooms out — never in — when they do not fit.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "KeepInView";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-14";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KeepInView.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KeepInView.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

var KeepInView = (function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 結果が可視領域からはみ出したときに合わせる倍率（1で余白なし。小さいほど余白が増える）
       呼び出し側から viewOptions.fitRatio で上書きできる */
    var DEFAULT_FIT_RATIO = 0.9;

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var LABELS = {
        checkbox: {
            keepInView: { ja: "結果を画面内に表示", en: "Keep the result in view" }
        },
        tooltip: {
            keepInView: {
                ja: "処理した結果が画面から外れたときに、見える位置へ表示を移します。すでに見えているときは動かしません。",
                en: "Moves the view so the result stays visible. The view is left alone when it is already in sight."
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
     * LABELS からラベルを取り出す
     * @param {string} labelPath - "checkbox.keepInView" のようなドット区切りのキー
     * @param {string} [uiLang] - "ja" または "en"（省略時はUI言語）
     * @returns {string} ラベル文字列（見つからない場合は空文字）
     */
    function getLabel(labelPath, uiLang) {
        var pathParts = labelPath.split(".");
        var entry = LABELS;
        for (var i = 0; i < pathParts.length; i++) {
            entry = entry[pathParts[i]];
            if (!entry) return "";
        }
        if (!uiLang) uiLang = getCurrentLang();
        return (entry[uiLang] != null) ? entry[uiLang] : entry.en;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 「結果を画面内に表示」チェックボックスを作る（ラベルとツールチップは内蔵）
     * @param {Group|Panel|Window} parentContainer - 追加先のコンテナ
     * @param {object} [checkboxOptions] - value: 初期値（既定 true）／lang: 表示言語／text: ラベルの差し替え
     * @returns {Checkbox} 作成したチェックボックス
     */
    function addCheckbox(parentContainer, checkboxOptions) {
        if (!checkboxOptions) checkboxOptions = {};

        var uiLang = checkboxOptions.lang || getCurrentLang();
        var checkbox = parentContainer.add('checkbox', undefined,
            checkboxOptions.text || getLabel('checkbox.keepInView', uiLang));
        checkbox.helpTip = getLabel('tooltip.keepInView', uiLang);
        /* 明示的に false を渡したときだけOFFで始める / only an explicit false starts it unchecked */
        checkbox.value = (checkboxOptions.value !== false);
        return checkbox;
    }

    /**
     * 複数アイテムを囲む外接範囲を求める
     * @param {PageItem[]} targetItems - 対象アイテム
     * @returns {{left: number, top: number, right: number, bottom: number}|null} 外接範囲（求められない場合は null）
     */
    function getItemsBounds(targetItems) {
        var unionBounds = null;

        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds;
            /* visibleBounds を持たないアイテムが混ざることがある / some items do not expose visibleBounds */
            try {
                itemBounds = targetItems[i].visibleBounds; // [left, top, right, bottom]
            } catch (e) {
                continue;
            }
            if (unionBounds === null) {
                unionBounds = { left: itemBounds[0], top: itemBounds[1], right: itemBounds[2], bottom: itemBounds[3] };
                continue;
            }
            if (itemBounds[0] < unionBounds.left) unionBounds.left = itemBounds[0];
            if (itemBounds[1] > unionBounds.top) unionBounds.top = itemBounds[1];
            if (itemBounds[2] > unionBounds.right) unionBounds.right = itemBounds[2];
            if (itemBounds[3] < unionBounds.bottom) unionBounds.bottom = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 外接範囲がすでに可視領域に収まっているか調べる
     * @param {{left: number, top: number, right: number, bottom: number}} targetBounds - 対象の外接範囲
     * @param {number[]} viewBounds - ビューの範囲 [left, top, right, bottom]
     * @returns {boolean} 全体が見えていれば true
     */
    function isFullyVisible(targetBounds, viewBounds) {
        return targetBounds.left >= viewBounds[0] && targetBounds.right <= viewBounds[2] &&
            targetBounds.top <= viewBounds[1] && targetBounds.bottom >= viewBounds[3];
    }

    /**
     * 結果が可視領域に収まっていなければ、見えるように表示位置とズームを合わせる
     * すでに見えているときは何もしないので、操作のたびに画面が動くことはない
     * @param {PageItem[]} targetItems - 見えるようにしたいアイテム
     * @param {object} [viewOptions] - doc: 対象ドキュメント（省略時は最前面）／fitRatio: 収めるときの倍率
     * @returns {boolean} 表示を動かしたら true
     */
    function ensureVisible(targetItems, viewOptions) {
        if (!targetItems || targetItems.length === 0) return false;
        if (!viewOptions) viewOptions = {};

        var targetBounds = getItemsBounds(targetItems);
        if (targetBounds === null) return false;

        var targetDoc = viewOptions.doc || (app.documents.length > 0 ? app.activeDocument : null);
        if (!targetDoc || targetDoc.views.length === 0) return false;
        var activeView = targetDoc.views[0];

        var viewBounds = activeView.bounds; // [left, top, right, bottom]
        /* すでに全体が見えているなら動かさない / leave the view alone when everything is already visible */
        if (isFullyVisible(targetBounds, viewBounds)) return false;

        var fitRatio = (viewOptions.fitRatio > 0) ? viewOptions.fitRatio : DEFAULT_FIT_RATIO;

        var targetWidth = targetBounds.right - targetBounds.left;
        var targetHeight = targetBounds.top - targetBounds.bottom;
        var visibleWidth = viewBounds[2] - viewBounds[0];
        var visibleHeight = viewBounds[1] - viewBounds[3];

        /* 収まらないときだけズームアウトする。拡大はしない（操作のたびに倍率が変わると落ち着かない）
           Only zoom out when it does not fit; never zoom in, so the magnification stays predictable */
        var targetZoom = activeView.zoom;
        if (targetWidth > visibleWidth || targetHeight > visibleHeight) {
            var zoomByWidth = (targetWidth > 0) ? targetZoom * visibleWidth / targetWidth : targetZoom;
            var zoomByHeight = (targetHeight > 0) ? targetZoom * visibleHeight / targetHeight : targetZoom;
            targetZoom = Math.min(zoomByWidth, zoomByHeight) * fitRatio;
        }

        /* 中心を合わせてからズームする（ズームは中心を保つ）/ center first, then zoom about that center */
        activeView.centerPoint = [targetBounds.left + targetWidth / 2, targetBounds.top - targetHeight / 2];
        activeView.zoom = targetZoom;
        return true;
    }

    return {
        addCheckbox: addCheckbox,
        ensureVisible: ensureVisible,
        getItemsBounds: getItemsBounds,
        getLabel: getLabel
    };

})();
