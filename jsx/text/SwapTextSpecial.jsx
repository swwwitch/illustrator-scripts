#target illustrator
#targetengine "SwapTextSpecialEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中の2つのテキストオブジェクトの内容を入れ替えます。
入れ替える対象（文字列／スタイル／座標）はダイアログで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapTextSpecial.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n071e09af28a7

### Overview

Swaps the contents of two selected text objects.
A dialog picks what is swapped: the contents, the style, or the position.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapTextSpecial.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapTextSpecial";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapTextSpecial.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapTextSpecial.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n071e09af28a7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 15];  /* パネル余白 / panel margins */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 日本語 / English */
    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "テキストの入れ替え", en: "Swap Text" }
        },
        panel: {
            target: { ja: "入れ替える対象", en: "Swap target" }
        },
        radio: {
            contents: { ja: "文字列", en: "String" },
            format: { ja: "書式", en: "Format" },
            position: { ja: "座標", en: "Position" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            needTwo: { ja: "テキストオブジェクトを2つ選択してください。", en: "Please select two text objects." },
            needText: {
                ja: "選択した2つは両方ともテキストオブジェクトである必要があります。",
                en: "Both selected objects must be text objects."
            }
        },
        tooltip: {
            contents: {
                ja: "2つのテキストの文字列だけを入れ替えます。書式と位置はそのままです。",
                en: "Swaps only the strings. The formatting and positions stay put."
            },
            format: {
                ja: "フォント・サイズ・色などの書式だけを入れ替えます。文字列と位置はそのままです。",
                en: "Swaps only the formatting, such as font, size, and colour. The strings and positions stay put."
            },
            position: {
                ja: "2つのテキストの位置だけを入れ替えます。中身はそのままです。",
                en: "Swaps only the positions. The contents stay put."
            }
        }
    };

    /**
     * 表示言語の文言を返す（無ければ英語）
     * @param {Object} labelSet - { ja, en } の文言オブジェクト
     * @returns {string} 表示言語の文言
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲


    // =============================================================
    // ダイアログ / Dialog
    // =============================================================

    /**
     * 入れ替え対象のラジオボタンを tooltip 付きで追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} modeKey - LABELS.radio と LABELS.tooltip のキー（"contents" / "format" / "position"）
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addModeRadio(parentPanel, modeKey) {
        var modeRadio = parentPanel.add("radiobutton", undefined, getLabel(LABELS.radio[modeKey]));
        modeRadio.helpTip = getLabel(LABELS.tooltip[modeKey]);
        return modeRadio;
    }

    /**
     * 入れ替える対象を選ぶダイアログを表示する
     * @returns {string|null} "contents" / "format" / "position"（キャンセル時は null）
     */
    function showSwapDialog() {
        var swapDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        swapDialog.alignChildren = "fill";

        var targetPanel = swapDialog.add("panel", undefined, getLabel(LABELS.panel.target));
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.margins = PANEL_MARGINS;

        var radioContents = addModeRadio(targetPanel, "contents");
        var radioFormat = addModeRadio(targetPanel, "format");
        var radioPosition = addModeRadio(targetPanel, "position");
        radioContents.value = true;

        var btnRowGroup = swapDialog.add("group");
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        prepareDialogWindow(swapDialog, SCRIPT_NAME);
        if (swapDialog.show() !== 1) {
            return null;
        }

        if (radioFormat.value) return "format";
        if (radioPosition.value) return "position";
        return "contents";
    }

    // =============================================================
    // 入れ替え処理 / Swap operations
    // =============================================================

    /**
     * 2つのテキストの文字列を入れ替える
     * @param {TextFrame} firstTextFrame - 1つ目のテキスト
     * @param {TextFrame} secondTextFrame - 2つ目のテキスト
     * @returns {void}
     */
    function swapContents(firstTextFrame, secondTextFrame) {
        var firstContents = firstTextFrame.contents;
        firstTextFrame.contents = secondTextFrame.contents;
        secondTextFrame.contents = firstContents;
    }

    /*
       入れ替える書式（characterAttributes）の一覧。
       List of character attributes to swap. textFont はフォント＋スタイルを兼ねる。
    */
    var FORMAT_ATTRIBUTE_KEYS = [
        "textFont",        /* フォント＋スタイル / font family + style */
        "size",            /* サイズ / size */
        "fillColor",       /* 文字カラー / text color */
        "strokeColor",     /* 線カラー / stroke color */
        "strokeWeight",    /* 線幅 / stroke weight */
        "tracking",        /* トラッキング / tracking */
        "leading",         /* 行送り / leading */
        "autoLeading",     /* 自動行送り / auto leading */
        "horizontalScale", /* 水平比率 / horizontal scale */
        "verticalScale",   /* 垂直比率 / vertical scale */
        "baselineShift",   /* ベースラインシフト / baseline shift */
        "capitalization"   /* 大文字小文字 / capitalization */
    ];

    /**
     * テキスト全体の書式を控える
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {Object} 属性名と値の組（読み取れなかった属性は含まない）
     */
    function captureFormatAttributes(textFrame) {
        var charAttributes = textFrame.textRange.characterAttributes;
        var capturedAttributes = {};
        for (var i = 0; i < FORMAT_ATTRIBUTE_KEYS.length; i++) {
            var attributeKey = FORMAT_ATTRIBUTE_KEYS[i];
            /* 属性によっては読み取りで例外になる / some attributes throw on read */
            try {
                capturedAttributes[attributeKey] = charAttributes[attributeKey];
            } catch (e) {}
        }
        return capturedAttributes;
    }

    /**
     * 控えた書式をテキスト全体に適用する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Object} capturedAttributes - captureFormatAttributes() の戻り値
     * @returns {void}
     */
    function applyFormatAttributes(textFrame, capturedAttributes) {
        var charAttributes = textFrame.textRange.characterAttributes;
        for (var i = 0; i < FORMAT_ATTRIBUTE_KEYS.length; i++) {
            var attributeKey = FORMAT_ATTRIBUTE_KEYS[i];
            if (!capturedAttributes.hasOwnProperty(attributeKey)) continue;
            /* 書き込めない値（未定義の色など）は飛ばす / skip values that cannot be written */
            try {
                charAttributes[attributeKey] = capturedAttributes[attributeKey];
            } catch (e) {}
        }
    }

    /**
     * 2つのテキストの書式を入れ替える
     * @param {TextFrame} firstTextFrame - 1つ目のテキスト
     * @param {TextFrame} secondTextFrame - 2つ目のテキスト
     * @returns {void}
     */
    function swapFormat(firstTextFrame, secondTextFrame) {
        /* 両方の書式を先に取得してから入れ替える / Capture both before applying */
        var firstAttributes = captureFormatAttributes(firstTextFrame);
        var secondAttributes = captureFormatAttributes(secondTextFrame);
        applyFormatAttributes(firstTextFrame, secondAttributes);
        applyFormatAttributes(secondTextFrame, firstAttributes);
    }

    /**
     * 2つのテキストの位置（左上）を入れ替える
     * @param {TextFrame} firstTextFrame - 1つ目のテキスト
     * @param {TextFrame} secondTextFrame - 2つ目のテキスト
     * @returns {void}
     */
    function swapPosition(firstTextFrame, secondTextFrame) {
        /*
           position はベースライン基準で上端/左端が崩れるため geometricBounds を使う。
           Use geometricBounds (not position) because TextFrame.position is baseline-based.
           geometricBounds = [left, top, right, bottom]
        */
        app.redraw(); // bounds が更新されない環境対策 / refresh stale bounds

        var firstBounds = firstTextFrame.geometricBounds;
        var secondBounds = secondTextFrame.geometricBounds;

        var deltaX = secondBounds[0] - firstBounds[0];
        var deltaY = secondBounds[1] - firstBounds[1];

        firstTextFrame.translate(deltaX, deltaY);
        secondTextFrame.translate(-deltaX, -deltaY);
    }

    // =============================================================
    // メイン / Main
    // =============================================================

    /**
     * 選択を確かめ、ダイアログで選んだ対象を入れ替える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (selectedItems.length !== 2) {
            alert(getLabel(LABELS.alert.needTwo));
            return;
        }
        if (selectedItems[0].typename !== "TextFrame" || selectedItems[1].typename !== "TextFrame") {
            alert(getLabel(LABELS.alert.needText));
            return;
        }

        var swapMode = showSwapDialog();
        if (swapMode === null) {
            return;
        }

        var firstTextFrame = selectedItems[0];
        var secondTextFrame = selectedItems[1];

        if (swapMode === "format") {
            swapFormat(firstTextFrame, secondTextFrame);
        } else if (swapMode === "position") {
            swapPosition(firstTextFrame, secondTextFrame);
        } else {
            swapContents(firstTextFrame, secondTextFrame);
        }
    }

    main();

})();
