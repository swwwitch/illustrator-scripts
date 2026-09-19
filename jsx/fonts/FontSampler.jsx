#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

入力したテキストを、インストールされているフォントで並べて表示します。
30個ぶん／アートボードいっぱい／すべてのフォント、の3モードから選べます。

詳細は README を参照してください。

### Overview

Lays out the text you type in each of the installed fonts.
Three modes are available: 30 fonts, as many as fit the artboard, or every font.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontSampler";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontSampler.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontSampler.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        fieldLabel: {
            sampleText: { ja: "テキスト", en: "Text" }
        },
        panel: {
            fontCount: { ja: "フォント数の設定", en: "How many fonts" }
        },
        radio: {
            limit30: { ja: "30個まで", en: "Up to 30" },
            fitBoard: { ja: "アートボードいっぱい", en: "Fill the artboard" },
            all: { ja: "すべて", en: "All" }
        },
        tooltip: {
            sampleText: { ja: "各フォントの見本として並べる文字です。", en: "The text shown as the specimen for each font." },
            limit30: { ja: "先頭から30書体までを並べます。", en: "Lays out the first 30 typefaces." },
            fitBoard: { ja: "アートボードに収まる数だけ並べます。", en: "Lays out as many as fit on the artboard." },
            all: { ja: "環境にあるすべての書体を並べます。数が多いと時間がかかります。", en: "Lays out every typeface on the machine. This can take a while." }
        },
        defaultSampleText: { ja: "山路を登りながら", en: "Handgloves" }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.fontCount" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのドット区切りキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ": ");
    }

    function main() {
        try {
            if (app.documents.length === 0) {
                alert("ドキュメントを開いてください。");
                return;
            }

            var doc = app.activeDocument;

            // ダイアログ作成 / Create dialog
            var dlg = new Window("dialog", "テキストを入力");
            dlg.orientation = "column";
            dlg.alignChildren = ["fill", "top"];

            var inputGroup = dlg.add("group");
            inputGroup.add("statictext", undefined, labelText("fieldLabel.sampleText"));
            var inputText = inputGroup.add("edittext", undefined, getLabel("defaultSampleText"), {
                multiline: false
            });
            inputText.helpTip = getLabel("tooltip.sampleText");
            inputText.characters = 30;

            // 新しいパネルにラジオボタンを追加 / Add radio buttons in new panel
            var optionPanel = dlg.add("panel", undefined, getLabel("panel.fontCount"));
            optionPanel.orientation = "column";
            optionPanel.alignChildren = ["left", "top"];
            optionPanel.margins = [15, 20, 15, 10];
            var rb30 = optionPanel.add("radiobutton", undefined, getLabel("radio.limit30"));
            rb30.helpTip = getLabel("tooltip.limit30");
            var rbBoard = optionPanel.add("radiobutton", undefined, getLabel("radio.fitBoard"));
            rbBoard.helpTip = getLabel("tooltip.fitBoard");
            var rbAll = optionPanel.add("radiobutton", undefined, getLabel("radio.all"));
            rbAll.helpTip = getLabel("tooltip.all");
            rb30.value = true; // デフォルト / Default

            dlg.onShow = function() {
                inputText.active = true;
            };

            var btnGroup = dlg.add("group");
            btnGroup.alignment = "center";
            var cancelBtn = btnGroup.add("button", undefined, "キャンセル");
            var okBtn = btnGroup.add("button", undefined, "OK");

            if (dlg.show() != 1) return; // キャンセル時 / Cancelled

            var textContent = inputText.text;
            if (!textContent || textContent === "") {
                alert("テキストを入力してください。");
                return;
            }

            // 使用可能フォントを取得 / Get available fonts
            var fonts = app.textFonts;

            // アートボードのサイズを取得 / Get artboard size
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            var abLeft = ab[0];
            var abTop = ab[1];
            var abRight = ab[2];
            var abBottom = ab[3];

            // レイアウト設定 / Layout settings
            var startX = abLeft + 50;
            var startY = abTop - 50;
            var lineHeight = 30; // 行間 / line height

            // 1つ仮のテキストフレームを作成して幅を測る / Create temp text frame to measure width
            var tempTF = doc.textFrames.add();
            tempTF.contents = textContent;
            tempTF.textRange.characterAttributes.size = 12; // 標準サイズ / standard size
            var sampleWidth = tempTF.width;
            tempTF.remove();

            // 列幅を文字幅に余白を加えて算出 / Calculate column width with margin
            var colWidth = sampleWidth + 50; // 40ptを余白として加算 / add 40pt margin

            var rowsPerCol = Math.floor((abTop - abBottom - 100) / lineHeight);
            var colsPerBoard = 2; // 固定で2カラム
            var maxCells = rowsPerCol * colsPerBoard;

            var mode = rb30.value ? "30" : rbBoard.value ? "board" : "all";

            var maxCount;
            if (mode === "30") {
                maxCount = Math.min(fonts.length, 30);
            } else if (mode === "board") {
                maxCount = Math.min(fonts.length, maxCells);
            } else {
                maxCount = fonts.length;
            }

            var progressWin;
            var progressBar;
            if (mode === "all") {
                progressWin = new Window("palette", "進行状況");
                progressWin.alignChildren = "fill";
                progressBar = progressWin.add("progressbar", undefined, 0, maxCount);
                progressBar.preferredSize = [300, 20];
                progressWin.show();
            }

            var col = 0;
            var row = 0;
            var artIndex = doc.artboards.getActiveArtboardIndex();

            for (var i = 0; i < maxCount; i++) {
                // 新規アートボード作成判定（「すべて」モード） / Check if new artboard needed (mode: All)
                if (mode === "all" && i > 0 && i % maxCells === 0) {
                    artIndex++;
                    // 新規アートボードを追加し位置をずらす / Add new artboard and offset its position
                    var abWidth = abRight - abLeft;
                    var abHeight = abTop - abBottom;
                    var offsetX = (artIndex) * (abWidth + 100); // 100pt間隔で横にずらす / offset horizontally by 100pt
                    var newRect = [abLeft + offsetX, abTop, abRight + offsetX, abBottom];
                    var newAB = doc.artboards.add(newRect);
                    doc.artboards.setActiveArtboardIndex(artIndex);
                    ab = doc.artboards[artIndex].artboardRect;
                    abLeft = ab[0];
                    abTop = ab[1];
                    abRight = ab[2];
                    abBottom = ab[3];
                    startX = abLeft + 50;
                    startY = abTop - 50;
                    col = 0;
                    row = 0;
                }

                var tf = doc.textFrames.add();
                tf.contents = textContent;
                var x = startX + (col * colWidth);
                var y = startY - (row * lineHeight);
                tf.position = [x, y];

                try {
                    tf.textRange.characterAttributes.textFont = fonts[i];
                } catch (e) {}

                row++;
                if (row >= rowsPerCol) {
                    row = 0;
                    col++;
                }
                if (mode === "all" && progressBar) {
                    progressBar.value = i + 1;
                    progressWin.update();
                }
            }

            if (mode === "all" && progressWin) {
                progressWin.close();
            }

        } catch (e) {
            alert("エラー: " + e);
        }
    }

    main();

})();
