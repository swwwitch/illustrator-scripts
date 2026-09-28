#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトに合わせて、アクティブビューをズーム＆センタリングします。
複数選択時は全体の外接矩形にフィットし、選択がないときは100%表示に戻します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZoomToSelection.md

### Overview

Zooms and centers the active view on the selected objects.
With several objects selected it fits their overall bounding box, and with nothing selected it returns to 100%.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZoomToSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ZoomToSelection";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZoomToSelection.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZoomToSelection.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author John Wundes (www.wundes.com)
 * @discussion http://www.wundes.com/js4ai/copyright.txt
 */

(function () {

// =========================================
// ユーザー設定 / User Settings
// =========================================

/* フィット時に残す余白の係数（1.0 でぴったり、0.9 で 10% 余白） / Fit margin ratio (1.0 = exact fit, 0.9 = 10% margin) */
var ZOOM_FIT_RATIO = 0.9;

/* 補間ステップ数の上限（大きな移動・ズーム時） / Maximum interpolation steps (for large moves or zooms) */
var MAX_ANIMATION_STEP_COUNT = 32;
/* 各フレームの待機ミリ秒（redraw 自体も時間を食う） / Delay per frame in ms (redraw itself also takes time) */
var FRAME_DELAY_MS = 3;

// =========================================
// ローカライズ / Localization
// =========================================

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

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
    }
};

// =========================================
// ビューの計算 / View calculation
// =========================================

/**
 * 選択オブジェクト全体の外接矩形を求める（visibleBounds 基準、top > bottom）
 * @param {PageItem[]} selectedItems - 選択オブジェクト
 * @returns {{left: number, top: number, right: number, bottom: number}} 外接矩形
 */
function getSelectionBounds(selectedItems) {
    var firstBounds = selectedItems[0].visibleBounds;
    var boundsLeft = firstBounds[0];
    var boundsTop = firstBounds[1];
    var boundsRight = firstBounds[2];
    var boundsBottom = firstBounds[3];

    for (var i = 1; i < selectedItems.length; i++) {
        var itemBounds = selectedItems[i].visibleBounds;
        if (itemBounds[0] < boundsLeft) boundsLeft = itemBounds[0];
        if (itemBounds[1] > boundsTop) boundsTop = itemBounds[1];
        if (itemBounds[2] > boundsRight) boundsRight = itemBounds[2];
        if (itemBounds[3] < boundsBottom) boundsBottom = itemBounds[3];
    }

    return {
        left: boundsLeft,
        top: boundsTop,
        right: boundsRight,
        bottom: boundsBottom
    };
}

/**
 * 対象を画面に収めるズーム倍率を、現在のズームと表示領域から求める
 * （現在の表示を基準にするので、開始位置からそのまま動かせる）
 * @param {number} startZoom - 現在のズーム倍率
 * @param {number} visibleWidth - 現在の表示領域の幅（ドキュメント座標）
 * @param {number} visibleHeight - 現在の表示領域の高さ（ドキュメント座標）
 * @param {number} targetWidth - 収めたい範囲の幅
 * @param {number} targetHeight - 収めたい範囲の高さ
 * @returns {number} 幅・高さの両方が収まるズーム倍率（求められないときは startZoom）
 */
function computeFitZoom(startZoom, visibleWidth, visibleHeight, targetWidth, targetHeight) {
    var fitZoom = null;

    if (targetWidth > 0 && visibleWidth > 0) {
        fitZoom = startZoom * visibleWidth / targetWidth;
    }
    if (targetHeight > 0 && visibleHeight > 0) {
        var fitZoomByHeight = startZoom * visibleHeight / targetHeight;
        if (fitZoom === null || fitZoomByHeight < fitZoom) {
            fitZoom = fitZoomByHeight;
        }
    }

    return (fitZoom === null) ? startZoom : fitZoom;
}

// =========================================
// アニメーション / Animation
// =========================================

/**
 * 指定ミリ秒だけ待機する（アニメーションのフレーム間隔）
 * @param {number} durationMs - 待機するミリ秒
 * @returns {void}
 */
function sleep(durationMs) {
    var startTime = new Date().getTime();
    while (new Date().getTime() - startTime < durationMs) { }
}

/**
 * 移動量に応じて補間ステップ数を決める（小さな移動はステップを減らしてもたつきを防ぐ）
 * 移動距離（現在の表示領域に対する割合）とズームの変化率の大きい方を「移動量」とする
 * @param {View} activeView - 対象のビュー
 * @param {number} startCenterX - 開始時の中心 X
 * @param {number} startCenterY - 開始時の中心 Y
 * @param {number} startZoom - 開始時のズーム倍率
 * @param {number} targetCenterX - 目標の中心 X
 * @param {number} targetCenterY - 目標の中心 Y
 * @param {number} targetZoom - 目標のズーム倍率
 * @returns {number} 補間ステップ数
 */
function resolveStepCount(activeView, startCenterX, startCenterY, startZoom, targetCenterX, targetCenterY, targetZoom) {
    var viewBounds = activeView.bounds; /* [left, top, right, bottom] */
    var visibleExtent = Math.max(
        Math.abs(viewBounds[2] - viewBounds[0]),
        Math.abs(viewBounds[1] - viewBounds[3])
    );

    var dx = targetCenterX - startCenterX;
    var dy = targetCenterY - startCenterY;
    var moveDistance = Math.sqrt(dx * dx + dy * dy);
    var moveFraction = (visibleExtent > 0) ? (moveDistance / visibleExtent) : 1;

    var maxZoom = Math.max(startZoom, targetZoom);
    var zoomFraction = (maxZoom > 0) ? (Math.abs(targetZoom - startZoom) / maxZoom) : 0;

    var magnitude = Math.min(Math.max(moveFraction, zoomFraction), 1);

    var minSteps = Math.max(2, Math.round(MAX_ANIMATION_STEP_COUNT * 0.2));
    var resolvedSteps = Math.round(MAX_ANIMATION_STEP_COUNT * magnitude);
    if (resolvedSteps < minSteps) resolvedSteps = minSteps;
    if (resolvedSteps > MAX_ANIMATION_STEP_COUNT) resolvedSteps = MAX_ANIMATION_STEP_COUNT;
    return resolvedSteps;
}

/**
 * 現在のビューから目標の中心・ズームへ少しずつ補間して動かす
 * @param {View} activeView - 対象のビュー
 * @param {number} targetCenterX - 目標の中心 X
 * @param {number} targetCenterY - 目標の中心 Y
 * @param {number} targetZoom - 目標のズーム倍率
 * @returns {void}
 */
function animateView(activeView, targetCenterX, targetCenterY, targetZoom) {
    var startCenter = activeView.centerPoint;
    var startCenterX = startCenter[0];
    var startCenterY = startCenter[1];
    var startZoom = activeView.zoom;

    /* 移動量に応じてステップ数を決める / Decide the step count from the amount of movement */
    var stepCount = resolveStepCount(
        activeView, startCenterX, startCenterY, startZoom,
        targetCenterX, targetCenterY, targetZoom
    );

    /* ズームは倍率を一定割合ずつ動かすと滑らかに見えるので、線形ではなく幾何補間にする
       Zoom is interpolated geometrically, which looks smoother than linear */
    var zoomRatio = (startZoom > 0) ? (targetZoom / startZoom) : 1;

    for (var i = 1; i <= stepCount; i++) {
        var progress = i / stepCount;
        /* easeOutQuad：終わりに向かって減速 / easeOutQuad: slow down toward the end */
        var easedProgress = 1 - (1 - progress) * (1 - progress);

        activeView.centerPoint = [
            startCenterX + (targetCenterX - startCenterX) * easedProgress,
            startCenterY + (targetCenterY - startCenterY) * easedProgress
        ];
        /* 幾何補間：startZoom × (targetZoom / startZoom)^t / Geometric interpolation */
        activeView.zoom = startZoom * Math.pow(zoomRatio, easedProgress);

        app.redraw();
        sleep(FRAME_DELAY_MS);
    }

    /* 最後に正確な値へ合わせる / Snap to the exact target at the end */
    activeView.centerPoint = [targetCenterX, targetCenterY];
    activeView.zoom = targetZoom;
    app.redraw();
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * 選択範囲にズームする。選択が無いときは現在位置のまま 100% に戻す
 * @returns {void}
 */
function main() {
    if (app.documents.length < 1) {
        alert(getLabel("alert.noDocument"));
        return;
    }
    var doc = app.activeDocument;
    var selectedItems = doc.selection;
    var activeView = doc.views[0];

    if (!(selectedItems && selectedItems.length > 0)) {
        /* 選択が無いときは現在位置のまま 100% へ（こちらも滑らかに） / No selection: animate back to 100% in place */
        var currentCenter = activeView.centerPoint;
        animateView(activeView, currentCenter[0], currentCenter[1], 1);
        return;
    }

    var selectionBounds = getSelectionBounds(selectedItems);
    var selectionWidth = selectionBounds.right - selectionBounds.left;
    var selectionHeight = selectionBounds.top - selectionBounds.bottom;

    /* 現在の表示領域（ドキュメント座標系）を基準にフィット倍率を求める / Fit zoom from the current visible area */
    var viewBounds = activeView.bounds; /* [left, top, right, bottom] */
    var targetZoom = computeFitZoom(
        activeView.zoom,
        viewBounds[2] - viewBounds[0],
        viewBounds[1] - viewBounds[3],
        selectionWidth,
        selectionHeight
    ) * ZOOM_FIT_RATIO;

    animateView(
        activeView,
        selectionBounds.left + (selectionWidth / 2),
        selectionBounds.top - (selectionHeight / 2),
        targetZoom
    );
}

main();

})();
