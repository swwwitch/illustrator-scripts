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

/**
 * 表示言語を判定する
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var uiLang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
    }
};

/**
 * LABELS からドット区切りのパスで表示言語のテキストを取り出す
 * @param {string} labelPath - "alert.noDocument" のようなドット区切りのキー
 * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
 */
function getLabel(labelPath) {
    var labelPathKeys = labelPath.split(".");
    var labelNode = LABELS;
    for (var i = 0; i < labelPathKeys.length; i++) {
        labelNode = labelNode[labelPathKeys[i]];
        if (!labelNode) {
            return labelPath;
        }
    }
    return labelNode[uiLang] || labelNode.en || labelPath;
}

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
