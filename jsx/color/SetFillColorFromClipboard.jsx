#target illustrator

/*
 * クリップボードの「204,225,179」のようなRGB文字列を塗り色に設定。
 * ReplaceTextWithPaste.jsx の2回ペースト方式を参考にしています。
 * 1回目は削除、2回目は値を読み取ってからカットします。
 * カットによってクリップボードはIllustratorのテキストオブジェクトになります。
 */
(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。");
        return;
    }

    var doc = app.activeDocument;
    var originalSelection = [];
    var pastedItems = [];
    var rgb = null;
    var failure = null;
    var selection = doc.selection;

    // TextRangeは編集終了後に無効になる可能性があるため保持しない。
    if (selection && selection.typename === "TextRange") {
        alert("文字の編集を終了してから実行してください。");
        return;
    }
    if (selection) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i]) originalSelection.push(selection[i]);
        }
    }

    function capturePastedItems() {
        var current = doc.selection;
        pastedItems = [];
        if (!current) return;
        if (current.typename === "TextRange") {
            throw new Error("テキスト編集モードを終了してから実行してください。");
        }
        for (var j = 0; j < current.length; j++) {
            if (current[j]) pastedItems.push(current[j]);
        }
    }

    function removePastedItems() {
        for (var j = pastedItems.length - 1; j >= 0; j--) {
            // カット済みの参照は無効なので無視する。
            try { pastedItems[j].remove(); } catch (e) {}
        }
        pastedItems = [];
    }

    function collectText(item, texts) {
        if (item.typename === "TextFrame") {
            texts.push(item.contents);
        } else if (item.typename === "GroupItem") {
            for (var j = 0; j < item.pageItems.length; j++) {
                collectText(item.pageItems[j], texts);
            }
        }
    }

    try {
        app.selectTool("Adobe Select Tool");
        // 1回目は内部クリップボード更新用。カットせず削除する。
        doc.selection = null;
        try {
            app.paste();
            app.redraw();
        } finally {
            capturePastedItems();
            removePastedItems();
        }

        doc.selection = null;
        try {
            app.paste();
            app.redraw();
        } finally {
            capturePastedItems();
        }

        var texts = [];
        for (var k = 0; k < pastedItems.length; k++) {
            collectText(pastedItems[k], texts);
        }
        if (texts.length !== 1) {
            throw new Error("RGB値を含む1つのテキストをコピーしてください。");
        }
        var match = /^\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*$/.exec(texts[0]);
        if (!match) {
            throw new Error("「204,225,179」の形式でRGB値をコピーしてください。");
        }
        rgb = [Number(match[1]), Number(match[2]), Number(match[3])];
        for (var n = 0; n < 3; n++) {
            if (rgb[n] > 255) throw new Error("RGB値は0〜255で指定してください。");
        }

        // 値を変数へ退避した後、貼り付けたオブジェクトだけをカット。
        doc.selection = pastedItems;
        app.cut();
    } catch (e) {
        failure = e;
    } finally {
        // 不正な値やカット失敗の場合も、一時オブジェクトを残さない。
        removePastedItems();
        doc.selection = originalSelection.length ? originalSelection : null;
    }

    if (failure) {
        alert("塗り色を設定できませんでした。\n" + failure);
        return;
    }

    var color = new RGBColor();
    color.red = rgb[0];
    color.green = rgb[1];
    color.blue = rgb[2];
    doc.defaultFillColor = color;
    app.redraw();
})();
