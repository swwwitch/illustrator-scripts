#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
DrawArtboardRectangle.jsx
 * 概要 / Summary
 * 現在のアートボードと同じ大きさの長方形を作成し、塗り・線を「なし」に設定します。
 * Creates a rectangle exactly the size of the current artboard with no fill and no stroke.
 *
 * 使い方 / Usage
 * 対象アートボードをアクティブにして実行してください。長方形はアクティブレイヤー最前面に追加されます。
 *
 * 更新履歴 / Changelog
 * 2025-08-20: 初版 / Initial version
 */


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