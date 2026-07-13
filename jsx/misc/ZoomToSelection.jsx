#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

/*
    概要：
    選択オブジェクトに合わせてアクティブビューをズーム＆センタリングする。
    複数選択時は全体のバウンディングボックスにフィット（周囲に少し余白を残す）。
    選択が無い場合は現在位置のまま 100% 表示に戻す。
    ズームは一気に切り替えず、少しずつ補間してアニメーション風に動かす。

    アルゴリズム原典：
    John Wundes ( john@wundes.com ) www.wundes.com  "Zoom and Center to Selection v2."
    http://www.wundes.com/js4ai/copyright.txt
    アニメーション補間は ArtboardNavigator.jsx（古島佑起さん）を参考にした。
*/

// ============================================================
// スクリプトバージョン / Script version
// ============================================================
var SCRIPT_VERSION = "v2.1.0";

// フィット時に残す余白の係数（1.0 でぴったり、0.9 で 10% 余白）
var ZOOM_FIT_RATIO = 0.9;

// アニメーション設定 / Animation settings
var MAX_ANIMATION_STEP_COUNT = 32; // 補間ステップ数の上限（大きな移動・ズーム時）
var FRAME_DELAY_MS = 3;            // 各フレームの待機ミリ秒（redraw 自体も時間を食う）

// ============================================================
// ローカライズ / Localization
// ============================================================
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    alertNoDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
};

function L(key) {
    if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
    if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
    return key;
}

// ============================================================
// 選択オブジェクト全体のバウンディングボックスを求める
// visibleBounds = [left, top, right, bottom]（top > bottom）
// ============================================================
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

// ============================================================
// 対象を画面に収めるズーム倍率を view.bounds から算出
// 現在のズーム・表示領域を基準にするので、開始位置からそのまま動かせる
// ============================================================
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

// ============================================================
// 指定ミリ秒だけ待機（アニメーションのフレーム間隔）
// ============================================================
function sleep(durationMs) {
    var startTime = new Date().getTime();
    while (new Date().getTime() - startTime < durationMs) { }
}

// ============================================================
// 移動量に応じて補間ステップ数を決める（小さな移動はステップを減らしてモタつき防止）
// 移動距離（現在の表示領域に対する割合）とズーム変化率の大きい方を「移動量」とする
// ============================================================
function resolveStepCount(activeView, startCenterX, startCenterY, startZoom, targetCenterX, targetCenterY, targetZoom) {
    var magnitude;

    var viewBounds = activeView.bounds; // [left, top, right, bottom]
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

    magnitude = Math.max(moveFraction, zoomFraction);
    if (magnitude > 1) magnitude = 1;

    var minSteps = Math.max(2, Math.round(MAX_ANIMATION_STEP_COUNT * 0.2));
    var resolvedSteps = Math.round(MAX_ANIMATION_STEP_COUNT * magnitude);
    if (resolvedSteps < minSteps) resolvedSteps = minSteps;
    if (resolvedSteps > MAX_ANIMATION_STEP_COUNT) resolvedSteps = MAX_ANIMATION_STEP_COUNT;
    return resolvedSteps;
}

// ============================================================
// 現在のビューから目標の中心・ズームへ少しずつ補間して動かす
// ============================================================
function animateView(activeView, targetCenterX, targetCenterY, targetZoom) {
    var startCenter = activeView.centerPoint;
    var startCenterX = startCenter[0];
    var startCenterY = startCenter[1];
    var startZoom = activeView.zoom;

    // 移動量に応じてステップ数を決める（bounds 取得前に呼ぶ）
    var stepCount = resolveStepCount(
        activeView, startCenterX, startCenterY, startZoom,
        targetCenterX, targetCenterY, targetZoom
    );

    // ズームは「倍率」を一定割合ずつ動かすと知覚的に滑らかになるため、
    // 線形補間ではなく幾何補間（指数補間）を使う。
    var zoomRatio = (startZoom > 0) ? (targetZoom / startZoom) : 1;

    for (var i = 1; i <= stepCount; i++) {
        var progress = i / stepCount;
        // easeOutQuad：終わりに向かってふわっと減速
        var easedProgress = 1 - (1 - progress) * (1 - progress);

        activeView.centerPoint = [
            startCenterX + (targetCenterX - startCenterX) * easedProgress,
            startCenterY + (targetCenterY - startCenterY) * easedProgress
        ];
        // 幾何補間：startZoom × (targetZoom/startZoom)^t
        activeView.zoom = startZoom * Math.pow(zoomRatio, easedProgress);

        app.redraw();
        sleep(FRAME_DELAY_MS);
    }

    // 最後に正確な値へ合わせる
    activeView.centerPoint = [targetCenterX, targetCenterY];
    activeView.zoom = targetZoom;
    app.redraw();
}

// ============================================================
// エントリポイント / Entry point
// ============================================================
if (app.documents.length < 1) {
    alert(L("alertNoDocument"));
} else {
    var doc = app.activeDocument;
    var selectedItems = doc.selection;
    var activeView = doc.views[0];

    if (selectedItems && selectedItems.length > 0) {
        var selectionBounds = getSelectionBounds(selectedItems);
        var selectionWidth = selectionBounds.right - selectionBounds.left;
        var selectionHeight = selectionBounds.top - selectionBounds.bottom;

        // 現在の表示領域（ドキュメント座標系）を基準にフィット倍率を求める
        var startZoom = activeView.zoom;
        var viewBounds = activeView.bounds; // [left, top, right, bottom]
        var visibleWidth = viewBounds[2] - viewBounds[0];
        var visibleHeight = viewBounds[1] - viewBounds[3];

        var targetZoom = computeFitZoom(
            startZoom, visibleWidth, visibleHeight, selectionWidth, selectionHeight
        ) * ZOOM_FIT_RATIO;

        animateView(
            activeView,
            selectionBounds.left + (selectionWidth / 2),
            selectionBounds.top - (selectionHeight / 2),
            targetZoom
        );
    } else {
        // 選択が無いときは現在位置のまま 100% へ（こちらも滑らかに）
        var currentCenter = activeView.centerPoint;
        animateView(activeView, currentCenter[0], currentCenter[1], 1);
    }
}

})();
