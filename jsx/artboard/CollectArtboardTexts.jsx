#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

全アートボード上にあるテキストの内容を集め、最後のアートボードの右側に新しいテキストフレームとして配置します（既定では1つにまとめ、設定で1つずつ縦に並べることもできます）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CollectArtboardTexts.md

### Overview

Collects the contents of the text frames on every artboard and places them as new text to the right of the last artboard (merged into one frame by default, or stacked one frame per text via a setting).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CollectArtboardTexts.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CollectArtboardTexts";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CollectArtboardTexts.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CollectArtboardTexts.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /*
     * 配置スタイル / Placement style
     * true  : テキスト 1 つに対して新規フレームを 1 つ作り、縦に並べる / One frame per source text, stacked vertically
     * false : 全テキストを 1 つのテキストフレームに改行で結合（既定）/ Merge all texts into a single frame with newlines (default)
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
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            layerUnavailable: {
                ja: "出力先（アクティブレイヤー）がロックまたは非表示のため作成できません。",
                en: "Cannot create the text: the active layer is locked or hidden."
            },
            noTexts: { ja: "対象となるテキストが見つかりませんでした。", en: "No text was found on the artboards." }
        }
    };

    /**
     * ラベルを取得する
     * @param {string} labelPath - "alert.noDocument" のようなドット区切りのキー
     * @returns {string} 現在のUI言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split('.');
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // ユーティリティ関数 / Utility Functions
    // =========================================

    /**
     * 2つの矩形が重なっているか判定する（Illustrator の Y 座標は上が大きい）
     * @param {number[]} boundsA - 矩形 [left, top, right, bottom]
     * @param {number[]} boundsB - 矩形 [left, top, right, bottom]
     * @returns {boolean} 重なっていれば true
     */
    function boundsIntersect(boundsA, boundsB) {
        return !(boundsA[2] < boundsB[0] || boundsA[0] > boundsB[2] ||
            boundsA[1] < boundsB[3] || boundsA[3] > boundsB[1]);
    }

    /**
     * テキストフレームと、その全ての親コンテナ（グループ／レイヤー）が可視かつ未ロックか判定する
     * @param {TextFrame} textFrame - 判定するテキストフレーム
     * @returns {boolean} 可視かつ未ロックなら true
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

    /**
     * 指定アートボード範囲と重なる可視テキストフレームの文字列を集める
     * @param {TextFrames} allTextFrames - ドキュメントの全テキストフレーム
     * @param {number[]} artboardBounds - アートボードの矩形 [left, top, right, bottom]
     * @param {string[]} result - 文字列を追加する配列
     * @returns {void}
     */
    function collectTextsInArtboard(allTextFrames, artboardBounds, result) {
        for (var j = 0; j < allTextFrames.length; j++) {
            var textFrame = allTextFrames[j];
            if (!isFrameAccessible(textFrame)) continue;
            if (textFrame.contents === "") continue;
            if (!boundsIntersect(textFrame.geometricBounds, artboardBounds)) continue;
            result.push(textFrame.contents);
        }
    }

    /**
     * 複数の矩形を内包する最小の矩形を返す
     * @param {Array<number[]>} boundsList - 矩形の配列
     * @returns {number[]} 包含矩形 [left, top, right, bottom]
     */
    function unionOfBounds(boundsList) {
        var unionBounds = [boundsList[0][0], boundsList[0][1], boundsList[0][2], boundsList[0][3]];
        for (var k = 1; k < boundsList.length; k++) {
            var bounds = boundsList[k];
            if (bounds[0] < unionBounds[0]) unionBounds[0] = bounds[0]; /* left:   min */
            if (bounds[1] > unionBounds[1]) unionBounds[1] = bounds[1]; /* top:    max (Y up) */
            if (bounds[2] > unionBounds[2]) unionBounds[2] = bounds[2]; /* right:  max */
            if (bounds[3] < unionBounds[3]) unionBounds[3] = bounds[3]; /* bottom: min */
        }
        return unionBounds;
    }

    /**
     * 指定の矩形が画面に収まるようビューをズーム＆センタリングする
     * @param {number[]} bounds - 矩形 [left, top, right, bottom]
     * @returns {void}
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

    /**
     * テキストを新しいフレームとして配置する
     * @param {Document} activeDoc - 対象ドキュメント
     * @param {string[]} flatTexts - 配置する文字列
     * @param {number[]} lastArtboardBounds - 最後のアートボードの矩形
     * @returns {TextFrame[]} 作成したテキストフレーム
     */
    function placeTexts(activeDoc, flatTexts, lastArtboardBounds) {
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
        return createdFrames;
    }

    /**
     * 各アートボード上のテキストを収集し、最後のアートボードの右側に配置してズーム表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var activeDoc = app.activeDocument;

        /* 出力先レイヤーがロック・非表示なら処理不可 / Abort if the active layer is locked or hidden */
        var targetLayer = activeDoc.activeLayer;
        if (targetLayer.locked || !targetLayer.visible) {
            alert(getLabel("alert.layerUnavailable"));
            return;
        }

        var artboardList = activeDoc.artboards;
        var allTextFrames = activeDoc.textFrames;

        /* アートボード順を保ったままテキストを平坦化 / Flatten texts while keeping artboard order */
        var flatTexts = [];
        for (var i = 0; i < artboardList.length; i++) {
            collectTextsInArtboard(allTextFrames, artboardList[i].artboardRect, flatTexts);
        }

        if (flatTexts.length === 0) {
            alert(getLabel("alert.noTexts"));
            return;
        }

        var lastArtboardBounds = artboardList[artboardList.length - 1].artboardRect;
        var createdFrames = placeTexts(activeDoc, flatTexts, lastArtboardBounds);

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

})();
