#target illustrator
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
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

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
