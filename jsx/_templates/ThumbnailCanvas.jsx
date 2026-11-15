#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログボックスの中に、設定に合わせた縮小図（並び・配置の模式図）を自作描画する再利用テンプレートです。
描画域の追加・描き直し・矩形の塗りと枠・文字の計測をまとめています。Illustrator・InDesign のどちらでも使えます。

### Overview

A reusable template that draws a thumbnail (a schematic of the layout) inside a dialog and redraws it as the settings change.
It covers adding the canvas, redrawing it, filling and framing rectangles, and measuring text. Works in both Illustrator and InDesign.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ThumbnailCanvas";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-11-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-15";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内に貼る
    // 2. 描く内容はスクリプト側の関数にし、addThumbnailCanvas() に渡す。描画関数は値を引数で受けず、ダイアログの変数を読む
    //      var thumbnailCanvas = addThumbnailCanvas(panel, [330, 160], function (canvasGraphics, canvasWidth, canvasHeight) { ... });
    // 3. 設定が変わったら redrawThumbnailCanvas(thumbnailCanvas) で描き直す（group には notify() が無いので hide()→show()）
    // 4. 配色は isDarkUI()（UITheme）で明暗を分ける。InDesign は文字サイズの指定を無視する（13pt 前後で描かれる）ので、
    //    文字は measureThumbnailText() で測って、収まるときだけ描く

    // 縮小図（再利用パーツ） / Thumbnail canvas (reusable)

    var thumbnailFontCache = {}; /* 文字サイズごとの書体 / fonts per size */

    /**
     * 縮小図を自作描画する描画域（group）を追加する
     * @param {Window|Panel|Group} parent - 追加先
     * @param {number[]} canvasSize - [幅, 高さ]（px）。fillWidth のときの幅は最小幅
     * @param {Function} drawFn - 描くたびに drawFn(canvasGraphics, canvasWidth, canvasHeight) で呼ぶ関数
     * @param {boolean} [fillWidth] - true なら親の幅いっぱいに伸ばす（ダイアログの幅は広げない）
     * @returns {Group} 描画域
     */
    function addThumbnailCanvas(parent, canvasSize, drawFn, fillWidth) {
        var thumbnailCanvas = parent.add("group");
        thumbnailCanvas.preferredSize = [canvasSize[0], canvasSize[1]];
        thumbnailCanvas.minimumSize = [canvasSize[0], canvasSize[1]];
        thumbnailCanvas.maximumSize = [fillWidth ? 10000 : canvasSize[0], canvasSize[1]];
        if (fillWidth) thumbnailCanvas.alignment = ["fill", "top"];
        thumbnailCanvas.onDraw = function () {
            var drawnSize = thumbnailCanvas.size;
            drawFn(thumbnailCanvas.graphics, drawnSize.width, drawnSize.height);
        };
        return thumbnailCanvas;
    }

    /**
     * 縮小図を描き直す（group には notify() が無いので、隠して再表示する）
     * @param {Group} thumbnailCanvas - addThumbnailCanvas() の戻り値
     * @returns {void}
     */
    function redrawThumbnailCanvas(thumbnailCanvas) {
        thumbnailCanvas.hide();
        thumbnailCanvas.show();
    }

    /**
     * 縮小図に矩形を塗る（rectPath の前に newPath を呼ばないとパスが重なって線画になる）
     * @param {ScriptUIGraphics} canvasGraphics - 描画先
     * @param {number} left - 左（px）
     * @param {number} top - 上（px）
     * @param {number} width - 幅（px）
     * @param {number} height - 高さ（px）
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function fillThumbnailRect(canvasGraphics, left, top, width, height, color) {
        canvasGraphics.newPath();
        canvasGraphics.rectPath(left, top, width, height);
        canvasGraphics.fillPath(canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, color));
    }

    /**
     * 縮小図に矩形の枠を描く（1px の線が内側に収まるよう 0.5px 内に寄せる）
     * @param {ScriptUIGraphics} canvasGraphics - 描画先
     * @param {number} left - 左（px）
     * @param {number} top - 上（px）
     * @param {number} width - 幅（px）
     * @param {number} height - 高さ（px）
     * @param {number[]} color - RGBA
     * @returns {void}
     */
    function strokeThumbnailRect(canvasGraphics, left, top, width, height, color) {
        canvasGraphics.newPath();
        canvasGraphics.rectPath(left + 0.5, top + 0.5, Math.max(width - 1, 0), Math.max(height - 1, 0));
        canvasGraphics.strokePath(canvasGraphics.newPen(canvasGraphics.PenType.SOLID_COLOR, color, 1));
    }

    /**
     * 縮小図の文字の書体を返す（作れない環境ではコントロールの書体を使う）
     * @param {ScriptUIGraphics} canvasGraphics - 描画先
     * @param {number} fontSize - 文字サイズ（px。InDesign では無視される）
     * @returns {ScriptUIFont} 書体
     */
    function getThumbnailFont(canvasGraphics, fontSize) {
        if (thumbnailFontCache[fontSize]) return thumbnailFontCache[fontSize];
        try {
            thumbnailFontCache[fontSize] = ScriptUI.newFont(canvasGraphics.font.name, "REGULAR", fontSize);
        } catch (e) {
            thumbnailFontCache[fontSize] = canvasGraphics.font;
        }
        return thumbnailFontCache[fontSize];
    }

    /**
     * 文字列の幅と高さを測る（measureString は配列か {width, height} で返る）
     * @param {ScriptUIGraphics} canvasGraphics - 描画先
     * @param {string} text - 文字列
     * @param {ScriptUIFont} font - 書体
     * @returns {number[]} [幅, 高さ]（px）
     */
    function measureThumbnailText(canvasGraphics, text, font) {
        var measured = canvasGraphics.measureString(text, font);
        if (measured.width !== undefined) return [measured.width, measured.height];
        return [measured[0], measured[1]];
    }

    // 縮小図（再利用パーツ）ここまで / End of the reusable thumbnail canvas

    // =========================================
    // デモ / Demo
    // =========================================

    /* 列数を変えると、縮小図に並びが描き直される / Changing the column count redraws the thumbnail */
    var demoDialog = new Window("dialog", "ThumbnailCanvas");
    var columnRow = demoDialog.add("group");
    columnRow.add("statictext", undefined, "列数 :");
    var columnInput = columnRow.add("edittext", undefined, "3");
    columnInput.characters = 4;
    var demoCanvas = addThumbnailCanvas(demoDialog, [240, 120], function (canvasGraphics, canvasWidth, canvasHeight) {
        fillThumbnailRect(canvasGraphics, 0, 0, canvasWidth, canvasHeight, [0.9, 0.9, 0.9, 1]);
        var columnCount = Math.max(1, parseInt(columnInput.text, 10) || 1);
        var cellCount = 7, cellSize = 20, gap = 6;
        var labelFont = getThumbnailFont(canvasGraphics, 9);
        var labelPen = canvasGraphics.newPen(canvasGraphics.PenType.SOLID_COLOR, [0.4, 0.4, 0.4, 1], 1);
        for (var i = 0; i < cellCount; i++) {
            var cellLeft = 10 + (i % columnCount) * (cellSize + gap);
            var cellTop = 10 + Math.floor(i / columnCount) * (cellSize + gap);
            fillThumbnailRect(canvasGraphics, cellLeft, cellTop, cellSize, cellSize, [1, 1, 1, 1]);
            strokeThumbnailRect(canvasGraphics, cellLeft, cellTop, cellSize, cellSize, [0.5, 0.5, 0.5, 1]);
            var labelSize = measureThumbnailText(canvasGraphics, String(i + 1), labelFont);
            canvasGraphics.drawString(String(i + 1), labelPen,
                cellLeft + (cellSize - labelSize[0]) / 2, cellTop + (cellSize - labelSize[1]) / 2, labelFont);
        }
    }, true);
    columnInput.onChanging = function () { redrawThumbnailCanvas(demoCanvas); };
    demoDialog.add("button", undefined, "OK", { name: "ok" });
    demoDialog.show();

})();
