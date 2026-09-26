#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内で使用されているフォントを集計し、名前順に一覧表示します（括弧内は使用数）。
一覧から選んだフォントを選択テキスト（グループ内を含む）へ即座に適用でき、テキストファイルとして書き出すこともできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyDocumentFonts.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n01d6ef7e9b5f

### Overview

Collects the fonts used in the document and lists them by name with their usage counts.
A font picked from the list can be applied to the selected text (including text in groups) right away, and the list can be exported as a text file.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyDocumentFonts.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyDocumentFonts";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-02-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyDocumentFonts.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyDocumentFonts.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n01d6ef7e9b5f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS        = 20;            /* ダイアログの余白 */
    var CONTROL_WIDTH         = 400;           /* 検索欄・一覧の幅 */
    var FILTER_HEIGHT         = 24;            /* 検索欄の高さ */
    var LIST_ROW_HEIGHT       = 20;            /* 一覧1行あたりの高さ（一覧の高さの計算用） */
    var LIST_MIN_HEIGHT       = 100;           /* 一覧の最小の高さ */
    var LIST_MAX_HEIGHT       = 300;           /* 一覧の最大の高さ */
    var BUTTON_ROW_TOP_MARGIN = 10;            /* ボタン行の上余白 */
    var BUTTON_SPACING        = 10;            /* 右側ボタンの間隔 */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ドキュメントフォントを適用", en: "Apply Document Fonts" }
        },
        fieldLabel: {
            filter:   { ja: "検索フィルター", en: "Search filter" },
            fontList: { ja: "ドキュメントフォント（名前順）", en: "Document fonts (by name)" }
        },
        tooltip: {
            filter: {
                ja: "フォント名・スタイル名・PostScript名に含まれる文字で一覧を絞り込みます。",
                en: "Filters the list by text in the font name, style name or PostScript name."
            },
            fontList: {
                ja: "ドキュメントで使われているフォントの一覧です（括弧内は使っているテキストの数）。選ぶと選択中のテキストにすぐ適用し、キャンセルで元に戻します。",
                en: "The fonts used in the document (the number of text objects in parentheses). Picking one applies it to the selected text right away; Cancel reverts it."
            },
            exportList: {
                ja: "表示中のフォント一覧を、デスクトップにテキストファイルとして書き出します。",
                en: "Saves the listed fonts to a text file on the desktop."
            }
        },
        button: {
            exportList: { ja: "書き出し", en: "Export" },
            cancel:     { ja: "キャンセル", en: "Cancel" },
            ok:         { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument:     { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            applyFontError: { ja: "フォントの適用に失敗しました。", en: "Failed to apply the font." },
            exportSuccess:  { ja: "書き出しました：", en: "Font list saved:\n" },
            exportError:    { ja: "ファイルの書き出しに失敗しました。", en: "Failed to save the file." }
        },
        exportText: {
            header:    { ja: "Adobe Illustrator ドキュメント情報", en: "Adobe Illustrator Document Info" },
            document:  { ja: "ドキュメント : ", en: "Document: " },
            fontCount: { ja: "このドキュメントに使用されているフォント : ", en: "Fonts used in this document: " },
            fileSuffix: { ja: "-ドキュメントフォント一覧.txt", en: "-Document-Font-List.txt" }
        }
    };

    /**
     * ラベルを取得する
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
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

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ドット区切りのキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === 'ja' ? '：' : ':');
    }

    // =========================================
    // フォントの収集 / Font collection
    // =========================================

    /**
     * フォントの表示名（ファミリー - スタイル）を返す
     * @param {TextFont} textFont - フォント
     * @returns {string} 表示名
     */
    function getFontDisplayName(textFont) {
        return textFont.style ? (textFont.family + " - " + textFont.style) : textFont.family;
    }

    /**
     * ドキュメントのテキストフレームからフォントの使用数を集計する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} 表示名をキーに {count, postScriptName} を持つ表
     */
    function collectDocumentFonts(doc) {
        var fontMap = {};
        for (var i = 0; i < doc.textFrames.length; i++) {
            var textFrame = doc.textFrames[i];
            if (!textFrame.contents || textFrame.contents.length === 0) continue;
            /* 未導入フォントなどで textFont が読めないことがある / textFont can throw (e.g. missing fonts) */
            try {
                var textFont = textFrame.textRange.characterAttributes.textFont;
                var displayName = getFontDisplayName(textFont);
                if (!fontMap[displayName]) {
                    fontMap[displayName] = { count: 1, postScriptName: textFont.name };
                } else {
                    fontMap[displayName].count++;
                }
            } catch (e) {}
        }
        return fontMap;
    }

    /**
     * 集計表を表示名の順に並べた配列にする
     * @param {Object} fontMap - collectDocumentFonts() の結果
     * @returns {Array<{displayName: string, postScriptName: string, count: number}>} 並べ替えたフォント
     */
    function sortFontsByName(fontMap) {
        var sortedFonts = [];
        for (var displayName in fontMap) {
            if (fontMap.hasOwnProperty(displayName)) {
                sortedFonts.push({
                    displayName: displayName,
                    postScriptName: fontMap[displayName].postScriptName,
                    count: fontMap[displayName].count
                });
            }
        }
        sortedFonts.sort(function(a, b) {
            return a.displayName.localeCompare(b.displayName);
        });
        return sortedFonts;
    }

    // =========================================
    // フォントの適用 / Font application
    // =========================================

    /**
     * 選択オブジェクトからテキストフレームを集める（グループの中も含む）
     * @param {PageItem} item - 対象オブジェクト
     * @param {TextFrame[]} result - テキストフレームを追加する配列
     * @returns {void}
     */
    function collectTextFrames(item, result) {
        if (item.typename === "TextFrame") {
            result.push(item);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectTextFrames(item.pageItems[i], result);
            }
        }
    }

    /**
     * テキストフレームと、その現在のフォントを控える
     * @param {Array|TextRange} selection - 選択オブジェクト
     * @returns {Array<{item: TextFrame, font: TextFont}>} 控え（キャンセル時の復元用）
     */
    function recordOriginalFonts(selection) {
        /* 文字を編集中の選択は TextRange で、配列ではない / While editing text, the selection is a TextRange */
        if (!selection || selection.typename) return [];
        var textFrames = [];
        for (var i = 0; i < selection.length; i++) {
            collectTextFrames(selection[i], textFrames);
        }
        var originalFonts = [];
        for (var j = 0; j < textFrames.length; j++) {
            /* 未導入フォントなどで textFont が読めないことがある / textFont can throw (e.g. missing fonts) */
            try {
                originalFonts.push({ item: textFrames[j], font: textFrames[j].textRange.characterAttributes.textFont });
            } catch (e) {}
        }
        return originalFonts;
    }

    /**
     * 控えたテキストフレームにフォントを適用する
     * @param {Array<{item: TextFrame, font: TextFont}>} originalFonts - 対象の控え
     * @param {TextFont} targetFont - 適用するフォント
     * @returns {void}
     */
    function applyFontToFrames(originalFonts, targetFont) {
        for (var i = 0; i < originalFonts.length; i++) {
            originalFonts[i].item.textRange.characterAttributes.textFont = targetFont;
        }
        app.redraw();
    }

    /**
     * 控えたフォントに戻す
     * @param {Array<{item: TextFrame, font: TextFont}>} originalFonts - 控え
     * @returns {void}
     */
    function restoreOriginalFonts(originalFonts) {
        for (var i = 0; i < originalFonts.length; i++) {
            /* 途中で削除されたテキストなどは飛ばして続ける / Skip frames that can no longer be written */
            try {
                originalFonts[i].item.textRange.characterAttributes.textFont = originalFonts[i].font;
            } catch (e) {}
        }
        app.redraw();
    }

    // =========================================
    // 書き出し / Export
    // =========================================

    /**
     * フォント一覧をデスクトップのテキストファイルに書き出す
     * @param {Document} doc - 対象ドキュメント
     * @param {string[]} fontNames - 書き出すフォントの表示名
     * @returns {void}
     */
    function exportFontList(doc, fontNames) {
        var docName = doc.name.replace(/\.[^\.]+$/, "");
        var saveFile = new File(Folder.desktop.fsName + "/" + docName + getLabel("exportText.fileSuffix"));

        var output = "";
        output += getLabel("exportText.header") + "\n\n";
        output += getLabel("exportText.document") + doc.fullName.fsName + "\n\n";
        output += getLabel("exportText.fontCount") + fontNames.length + "\n\n";
        output += fontNames.join("\n") + "\n";

        /* ファイルの書き込みは権限などで失敗しうる / File I/O can fail (permissions etc.) */
        try {
            saveFile.encoding = "UTF-8";
            saveFile.open("w");
            saveFile.write(output);
            saveFile.close();
            alert(getLabel("alert.exportSuccess") + saveFile.fsName);
        } catch (e) {
            alert(getLabel("alert.exportError"));
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 絞り込み文字列に合うフォントで一覧を作り直す
     * @param {ListBox} fontListBox - フォント一覧
     * @param {Array<{displayName: string, postScriptName: string, count: number}>} sortedFonts - 全フォント
     * @param {string} filterText - 絞り込み文字列（空ならすべて）
     * @returns {void}
     */
    function updateFontList(fontListBox, sortedFonts, filterText) {
        fontListBox.removeAll();
        var filterLower = filterText.toLowerCase();
        for (var i = 0; i < sortedFonts.length; i++) {
            var font = sortedFonts[i];
            if (filterLower === "" ||
                font.displayName.toLowerCase().indexOf(filterLower) !== -1 ||
                font.postScriptName.toLowerCase().indexOf(filterLower) !== -1) {
                var listItem = fontListBox.add("item", font.displayName + " (" + font.count + ")");
                listItem.fontDisplayName = font.displayName;
                listItem.postScriptName = font.postScriptName;
            }
        }
        if (fontListBox.items.length > 0) fontListBox.selection = 0;
    }

    /**
     * ダイアログを組み立てて表示する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<{displayName: string, postScriptName: string, count: number}>} sortedFonts - 全フォント
     * @param {Array<{item: TextFrame, font: TextFont}>} originalFonts - 選択中のテキストフレームと元のフォント
     * @returns {void}
     */
    function showDialog(doc, sortedFonts, originalFonts) {
        var dialog = new Window("dialog", getLabel("dialog.title"));
        dialog.orientation = "column";
        dialog.alignChildren = ["left", "top"];
        dialog.margins = DIALOG_MARGINS;

        dialog.add("statictext", undefined, labelText("fieldLabel.filter"));
        var filterInput = dialog.add("edittext", undefined, "");
        filterInput.helpTip = getLabel("tooltip.filter");
        filterInput.preferredSize = [CONTROL_WIDTH, FILTER_HEIGHT];

        dialog.add("statictext", undefined, labelText("fieldLabel.fontList"));
        var fontListBox = dialog.add("listbox", undefined, [], { multiselect: false });
        fontListBox.helpTip = getLabel("tooltip.fontList");
        fontListBox.preferredSize = [CONTROL_WIDTH, Math.min(LIST_MAX_HEIGHT, Math.max(LIST_MIN_HEIGHT, sortedFonts.length * LIST_ROW_HEIGHT))];

        updateFontList(fontListBox, sortedFonts, "");

        filterInput.onChanging = function() {
            updateFontList(fontListBox, sortedFonts, filterInput.text);
        };

        fontListBox.onChange = function() {
            if (!fontListBox.selection) return;
            /* getByName は見つからないと例外 / getByName throws when the font is missing */
            try {
                applyFontToFrames(originalFonts, app.textFonts.getByName(fontListBox.selection.postScriptName));
            } catch (e) {
                alert(getLabel("alert.applyFontError"));
            }
        };

        /* ボタン行 / Button row */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = "fill";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = 0;

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = "left";
        btnLeftGroup.alignment = ["left", "center"];
        var btnExport = btnLeftGroup.add("button", undefined, getLabel("button.exportList"));
        btnExport.helpTip = getLabel("tooltip.exportList");

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 10;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.spacing = BUTTON_SPACING;
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnOK.onClick = function() {
            dialog.close();
        };

        btnCancel.onClick = function() {
            restoreOriginalFonts(originalFonts);
            dialog.close();
        };

        btnExport.onClick = function() {
            var fontNames = [];
            for (var i = 0; i < fontListBox.items.length; i++) {
                fontNames.push(fontListBox.items[i].fontDisplayName);
            }
            exportFontList(doc, fontNames);
        };

        dialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var sortedFonts = sortFontsByName(collectDocumentFonts(doc));
        var originalFonts = recordOriginalFonts(doc.selection);
        showDialog(doc, sortedFonts, originalFonts);
    }

    main();
})();
