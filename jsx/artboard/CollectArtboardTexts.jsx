#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name

CollectArtboardTexts.jsx

### 概要 / Overview

- 全アートボード上にあるテキストフレームを収集し、最後のアートボードの右側に縦に並べて配置する
- Collect text frames on every artboard and place them as new text frames, stacked vertically to the right of the last artboard.

### 主な機能 / Main Features

- アートボードごとに重なるテキストフレームを抽出 / Pick text frames overlapping each artboard
- ロック・非表示のテキスト／レイヤー／グループは除外（親階層も走査）/ Skip locked or hidden frames and ancestor layers/groups
- 改行で結合せず、1 テキスト＝1 フレームで縦に並べる（スイッチで一括結合モードも可）/ One frame per source text stacked vertically (switch to combined mode available)
- 書き出した全フレームを選択してズーム表示 / Select the placed frames and zoom to fit

### 処理の流れ / Processing Flow

1. アクティブドキュメントと出力先レイヤーを検証 / Validate the active document and target layer
2. アートボードごとに含まれるテキストを収集 / Collect contained texts per artboard
3. テキストごとに新規フレームを縦に並べて配置 / Place each text as a separate frame, stacked vertically
4. 書き出した全フレームを選択しズーム表示 / Select created frames and zoom in

### 更新履歴 / Update History

- v1.0.0 (20260513) : 初版 / Initial release
*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CollectArtboardTexts";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// 設定 / Settings
// =========================================

/*
 * 配置スタイル / Placement style
 * true  : テキスト 1 つに対して新規フレームを 1 つ作り、縦に並べる（既定）/ One frame per source text, stacked vertically (default)
 * false : 全テキストを 1 つのテキストフレームに改行で結合 / Merge all texts into a single frame with newlines
 */
var SEPARATE_FRAMES = false;

/* 最後のアートボードの右端から書き出しフレームまでの距離（pt）/ Gap from the right edge of the last artboard (pt) */
var OUTPUT_GAP_PT = 50;

/* 縦に並べる際のフレーム間隔（pt）/ Vertical gap between stacked frames (pt) */
var INTER_FRAME_GAP_PT = 20;

/* ズーム時の余白率（1 で隙間なし）/ Padding factor for zoom (1 = no margin) */
var ZOOM_PADDING = 0.7;

/* Illustrator のズーム上限 / Illustrator zoom max */
var ZOOM_MAX = 64;

// =========================================
// ユーティリティ関数 / Utility Functions
// =========================================

/*
 * 2つの矩形（bounds）が重なっているかを判定する
 * Return true if two rectangles (bounds) overlap.
 * bounds: [left, top, right, bottom] — Illustrator の Y 座標は上が大きい / Y increases upward in Illustrator
 */
function boundsIntersect(boundsA, boundsB) {
    return !(boundsA[2] < boundsB[0] || boundsA[0] > boundsB[2] ||
        boundsA[1] < boundsB[3] || boundsA[3] > boundsB[1]);
}

/*
 * テキストフレームと、その全ての親コンテナ（グループ／レイヤー）が可視かつ未ロックかを判定
 * Return true if the text frame and all its ancestor containers (groups/layers) are visible and unlocked.
 */
function isFrameAccessible(textFrame) {
    if (textFrame.hidden || textFrame.locked) return false;
    var node = textFrame.parent;
    while (node && node.typename !== "Document") {
        if (node.typename === "Layer") {
            if (node.locked || !node.visible) return false;
        } else {
            /* GroupItem などのページアイテム / GroupItem and other page items */
            if (node.locked || node.hidden) return false;
        }
        node = node.parent;
    }
    return true;
}

/*
 * 指定アートボード範囲と重なる可視テキストフレームの文字列を集める
 * Collect contents of visible/unlocked text frames overlapping the given artboard bounds.
 */
function collectTextsInArtboard(allTextFrames, artboardBounds) {
    var collected = [];
    for (var j = 0; j < allTextFrames.length; j++) {
        var textFrame = allTextFrames[j];
        if (!isFrameAccessible(textFrame)) continue;
        if (textFrame.contents === "") continue;
        if (!boundsIntersect(textFrame.geometricBounds, artboardBounds)) continue;
        collected.push(textFrame.contents);
    }
    return collected;
}

/*
 * 複数の bounds を内包する最小の包含 bounds を返す
 * Compute the union (enclosing) bounds from a list of bounds.
 */
function unionOfBounds(boundsList) {
    var u = [boundsList[0][0], boundsList[0][1], boundsList[0][2], boundsList[0][3]];
    for (var k = 1; k < boundsList.length; k++) {
        var b = boundsList[k];
        if (b[0] < u[0]) u[0] = b[0]; /* left:   min */
        if (b[1] > u[1]) u[1] = b[1]; /* top:    max (Y up) */
        if (b[2] > u[2]) u[2] = b[2]; /* right:  max */
        if (b[3] < u[3]) u[3] = b[3]; /* bottom: min */
    }
    return u;
}

/*
 * 指定 bounds が画面に収まるようビューをズーム＆センタリングする
 * Zoom and center the active view so the given bounds fit on screen.
 */
function zoomToBounds(bounds) {
    var view = app.activeDocument.activeView;
    var targetWidth = bounds[2] - bounds[0];
    var targetHeight = bounds[1] - bounds[3];
    if (targetWidth <= 0 || targetHeight <= 0) return;

    var viewBounds = view.bounds;
    var viewWidth = viewBounds[2] - viewBounds[0];
    var viewHeight = viewBounds[1] - viewBounds[3];

    /* 過剰ズームを防ぐため Illustrator の上限でクランプ / Clamp to Illustrator's zoom max to avoid overshoot */
    view.zoom = Math.min(view.zoom * Math.min(viewWidth / targetWidth, viewHeight / targetHeight) * ZOOM_PADDING, ZOOM_MAX);
    view.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
}

// =========================================
// メイン処理 / Main
// =========================================

/*
 * エントリポイント / Entry point
 * 各アートボード上のテキストを収集し、最後のアートボードの右側に縦並びで配置してズーム表示する
 * Collect texts from each artboard, place them stacked vertically to the right of the last artboard, and zoom to fit.
 */
function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var activeDoc = app.activeDocument;

    /* 出力先レイヤーがロック・非表示なら処理不可 / Abort if the active layer is locked or hidden */
    var targetLayer = activeDoc.activeLayer;
    if (targetLayer.locked || !targetLayer.visible) {
        alert("出力先（アクティブレイヤー）がロックまたは非表示のため作成できません。");
        return;
    }

    var artboardList = activeDoc.artboards;
    var allTextFrames = activeDoc.textFrames;

    /* アートボード順を保ったままテキストを平坦化 / Flatten texts while keeping artboard order */
    var flatTexts = [];
    for (var i = 0; i < artboardList.length; i++) {
        var collected = collectTextsInArtboard(allTextFrames, artboardList[i].artboardRect);
        for (var k = 0; k < collected.length; k++) {
            flatTexts.push(collected[k]);
        }
    }

    if (flatTexts.length === 0) {
        alert("対象となるテキストが見つかりませんでした。");
        return;
    }

    var lastArtboardBounds = artboardList[artboardList.length - 1].artboardRect;
    var outputLeft = lastArtboardBounds[2] + OUTPUT_GAP_PT;
    var createdFrames = [];

    if (SEPARATE_FRAMES) {
        /* 1 テキスト＝1 フレームで縦に並べる / One frame per text, stacked vertically */
        var currentTop = lastArtboardBounds[1];
        for (var m = 0; m < flatTexts.length; m++) {
            var newFrame = activeDoc.textFrames.add();
            newFrame.contents = flatTexts[m];
            newFrame.position = [outputLeft, currentTop];
            app.redraw();
            currentTop = newFrame.geometricBounds[3] - INTER_FRAME_GAP_PT;
            createdFrames.push(newFrame);
        }
    } else {
        /* 全テキストを 1 フレームに結合 / Merge all into a single frame */
        var combinedFrame = activeDoc.textFrames.add();
        combinedFrame.contents = flatTexts.join("\r");
        combinedFrame.position = [outputLeft, lastArtboardBounds[1]];
        createdFrames.push(combinedFrame);
    }

    /* 書き出した全フレームを選択してズーム / Select created frames and zoom to fit */
    app.redraw();
    activeDoc.selection = null;
    var allBounds = [];
    for (var n = 0; n < createdFrames.length; n++) {
        createdFrames[n].selected = true;
        allBounds.push(createdFrames[n].geometricBounds);
    }
    zoomToBounds(unionOfBounds(allBounds));
    app.redraw();
}

main();