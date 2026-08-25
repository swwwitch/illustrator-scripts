#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

現在のアートボードと同じ大きさの長方形を作成し、塗り・線を「なし」に設定します。

詳細は README を参照してください。

### Overview

Creates a rectangle exactly the size of the current artboard, with no fill and no stroke.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DrawArtboardRectangle";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-08-20";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawArtboardRectangle.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawArtboardRectangle.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function createBlackColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) {
            var blackColor = new RGBColor();
            blackColor.red = 0;
            blackColor.green = 0;
            blackColor.blue = 0;
            return blackColor;
        } else {
            var blackColor = new CMYKColor();
            blackColor.black = 100;
            blackColor.cyan = 0;
            blackColor.magenta = 0;
            blackColor.yellow = 0;
            return blackColor;
        }
    }

    function drawArtboardRectangle() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var abRect = ab.artboardRect; // [left, top, right, bottom]

        var abWidth = abRect[2] - abRect[0];
        var abHeight = abRect[1] - abRect[3];

        app.executeMenuCommand('deselectall'); // 既存選択を解除

        var rect = doc.pathItems.rectangle(abRect[1], abRect[0], abWidth, abHeight);
        rect.filled = true;
        rect.fillColor = createBlackColor(doc);
        rect.stroked = false;
        rect.opacity = 15;
        rect.name = "Artboard Bounds";
        rect.selected = true;
        rect.zOrder(ZOrderMethod.SENDTOBACK);
    }

    drawArtboardRectangle();

})();
