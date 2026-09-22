#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のテキストフレームを、指定した文字（、。〜など）の直後で改行します。
改行の対象にする文字はチェックボックスで選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/titlemaker.md

### Overview

Breaks the selected text frame onto a new line right after the characters you choose, such as 、 。 or 〜.
Which characters trigger a break is set with checkboxes.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/titlemaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "titlemaker";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/titlemaker.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/titlemaker.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 改行対象文字と、開いたときのチェック状態 / Line-break characters and whether each starts checked */
    var TARGET_CHARS = [
        { mark: "、", defaultOn: true },
        { mark: "。", defaultOn: true },
        { mark: "〜", defaultOn: false }
    ];

    /* サイズ調整のデフォルト倍率（%）/ Default scale for size adjustment (%) */
    var DEFAULT_SIZE_PERCENT = 80;

    // =========================================
    // レイアウト / Layout
    // =========================================
    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;
    var PANEL_ITEM_SPACING = 6;   /* チェックボックスが並ぶパネルの間隔 / spacing inside the checkbox panels */
    var SIZE_ROW_SPACING = 4;     /* サイズ入力行の間隔 / spacing of the size row */
    var SIZE_FIELD_CHARS = 4;     /* サイズ欄の文字数 / size field width */

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループに共通のレイアウト設定を適用する（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // 文字種の判定 / Character Matching
    // =========================================
    /* 格助詞本体の正規表現ソース（長い候補を先に）/ Regex source for the case particle itself (longer alternatives first) */
    var CASE_PARTICLE_SOURCE = "(?:から|より|が|を|に|へ|と|で|の)";

    /* 助詞の直前に許される文字クラス（漢字・ひらがな・カタカナ・英数字）。ExtendScript は lookbehind 非対応のため手動判定に使う / Allowed preceding-char class for a particle (kanji / hiragana / katakana / alphanumeric); used manually because ExtendScript lacks lookbehind */
    var PARTICLE_PRECURSOR_SOURCE = "[一-龯ぁ-んァ-ヶA-Za-z0-9]";

    /**
     * ひらがなかどうか判定する（U+3041〜U+309F）
     * @param {string} oneChar - 判定する1文字
     * @returns {boolean} ひらがななら true
     */
    function isHiragana(oneChar) {
        var charCode = oneChar.charCodeAt(0);
        return charCode >= 0x3041 && charCode <= 0x309F;
    }

    /**
     * 格助詞として縮小する文字のインデックスを集める（直前が漢字・かな・英数字のものだけ）
     * @param {string} sourceText - 対象の文字列
     * @returns {Object} インデックスをキーにした集合（値は true）
     */
    function findCaseParticleIndices(sourceText) {
        var markedIndices = {};
        var particlePattern = new RegExp(CASE_PARTICLE_SOURCE, "g");
        var precursorPattern = new RegExp(PARTICLE_PRECURSOR_SOURCE);
        var particleMatch;

        while ((particleMatch = particlePattern.exec(sourceText)) !== null) {
            var matchStart = particleMatch.index;
            /* lookbehind の代わりに直前の1文字を判定 / Emulate lookbehind by testing the single preceding char */
            var prevChar = (matchStart > 0) ? sourceText.charAt(matchStart - 1) : "";
            if (prevChar !== "" && precursorPattern.test(prevChar)) {
                for (var k = 0; k < particleMatch[0].length; k++) markedIndices[matchStart + k] = true;
            }
            /* 空マッチによる無限ループを防ぐ / Guard against infinite loops on zero-length matches */
            if (particleMatch.index === particlePattern.lastIndex) particlePattern.lastIndex++;
        }
        return markedIndices;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する（ja で始まれば日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "タイトルメーカー", en: "Title Maker" }
        },
        panel: {
            targets: { ja: "改行対象文字", en: "Target Characters" },
            sizeAdjust: { ja: "フォントサイズのサイズ調整", en: "Font Size Adjustment" }
        },
        checkbox: {
            caseParticle: { ja: "格助詞", en: "Case particles" },
            hiragana: { ja: "ひらがな", en: "Hiragana" }
        },
        fieldLabel: {
            size: { ja: "サイズ", en: "Size" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            targetChar: { ja: "この文字の後ろで改行します。", en: "Inserts a line break after this character." },
            caseParticle: {
                ja: "「が」「を」「に」などの格助詞を小さくします。",
                en: "Shrinks case particles such as \u0022\u304c\u0022, \u0022\u3092\u0022, and \u0022\u306b\u0022."
            },
            hiragana: { ja: "ひらがなをまとめて小さくします。", en: "Shrinks all hiragana." },
            size: {
                ja: "小さくするときの大きさです。元のフォントサイズに対する割合で指定します。",
                en: "Size applied when shrinking, as a percentage of the original font size."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            noTarget: {
                ja: "改行対象文字またはサイズ調整の対象を1つ以上選択してください。",
                en: "Please select at least one line-break character or size-adjustment target."
            },
            invalidSize: {
                ja: "サイズには正の数値を入力してください。",
                en: "Please enter a positive number for the size."
            }
        }
    };

    /**
     * ドット区切りのキーで表示言語の文言を返す
     * @param {string} keyPath - "dialog.title" のようなキー
     * @returns {string} 表示言語の文言
     */
    function getLabel(keyPath) {
        var keyParts = keyPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            labelNode = labelNode[keyParts[i]];
        }
        return labelNode[uiLang];
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} keyPath - ラベルのキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(keyPath) {
        return getLabel(keyPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // テキストの加工 / Text processing
    // =========================================

    /**
     * 対象文字の後ろで改行し、欧文ベースラインを適用する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string[]} breakMarks - 改行する文字
     * @returns {void}
     */
    function insertLineBreaksAfterMarks(textFrame, breakMarks) {
        var frameContents = textFrame.contents;

        /* 各対象文字の後ろに改行を挿入 / Insert a line break after each target character */
        for (var i = 0; i < breakMarks.length; i++) {
            var breakMark = breakMarks[i];
            frameContents = frameContents.split(breakMark).join(breakMark + "\r");
        }

        /* 連続改行を整理 / Collapse consecutive line breaks */
        frameContents = frameContents.replace(/\r\r+/g, "\r");

        textFrame.contents = frameContents;

        /* 文字揃え：欧文ベースライン / Character alignment: Roman baseline */
        textFrame.textRange.characterAttributes.baselinePosition =
            FontBaselineOption.NORMALBASELINE;
    }

    /**
     * 格助詞・ひらがなのフォントサイズを指定の割合に縮小する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {{adjustCaseParticle: boolean, adjustHiragana: boolean, sizePercent: number}} sizeOptions - サイズ調整の設定
     * @returns {void}
     */
    function adjustCharacterSizes(textFrame, sizeOptions) {
        var sizeScale = sizeOptions.sizePercent / 100;
        var frameText = textFrame.contents;
        var frameCharacters = textFrame.textRange.characters;

        /* 縮小対象の文字のインデックス（文字列のインデックスと characters は 1:1 対応）/ Indices to scale (string index maps 1:1 to characters) */
        var markedIndices = sizeOptions.adjustCaseParticle ? findCaseParticleIndices(frameText) : {};

        /* ひらがなは1文字ずつ判定して追加 / Add hiragana per character */
        if (sizeOptions.adjustHiragana) {
            for (var i = 0; i < frameText.length; i++) {
                if (isHiragana(frameText.charAt(i))) markedIndices[i] = true;
            }
        }

        /* マーク済みの文字を縮小 / Scale the marked characters */
        for (var j = 0; j < frameCharacters.length; j++) {
            if (markedIndices[j]) {
                var charAttributes = frameCharacters[j].characterAttributes;
                charAttributes.size = charAttributes.size * sizeScale;
            }
        }
    }

    /**
     * 選択中の各テキストフレームに改行挿入とサイズ調整を適用する
     * @param {string[]} breakMarks - 改行する文字（空なら改行しない）
     * @param {Object|null} sizeOptions - サイズ調整の設定（null なら調整しない）
     * @returns {void}
     */
    function processSelection(breakMarks, sizeOptions) {
        var docSelection = app.activeDocument.selection;
        for (var i = 0; i < docSelection.length; i++) {
            var selectedItem = docSelection[i];

            if (selectedItem.typename !== "TextFrame") continue;

            if (breakMarks.length > 0) insertLineBreaksAfterMarks(selectedItem, breakMarks);
            if (sizeOptions) adjustCharacterSizes(selectedItem, sizeOptions);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる（イベントの配線は main() で行う）
     * @returns {Object} 作成したダイアログとコントロール
     */
    function buildTitleMakerDialog() {
        /* タイトルバーにバージョンを表示 / Version shown in the title bar */
        var lineBreakDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        lineBreakDialog.orientation = "column";
        lineBreakDialog.alignChildren = "fill";

        /* 改行対象文字のチェックボックスパネル / Checkbox panel for target characters */
        var targetPanel = lineBreakDialog.add("panel", undefined, getLabel("panel.targets"));
        setupPanel(targetPanel, PANEL_ITEM_SPACING);

        var targetCheckboxes = [];
        for (var i = 0; i < TARGET_CHARS.length; i++) {
            var targetCheckbox = targetPanel.add("checkbox", undefined, TARGET_CHARS[i].mark);
            targetCheckbox.helpTip = getLabel("tooltip.targetChar");
            targetCheckbox.value = TARGET_CHARS[i].defaultOn;
            targetCheckboxes.push(targetCheckbox);
        }

        /* フォントサイズ調整のパネル / Panel for font size adjustment */
        var sizePanel = lineBreakDialog.add("panel", undefined, getLabel("panel.sizeAdjust"));
        setupPanel(sizePanel, PANEL_ITEM_SPACING);

        var caseParticleCheckbox = sizePanel.add("checkbox", undefined, getLabel("checkbox.caseParticle"));
        caseParticleCheckbox.helpTip = getLabel("tooltip.caseParticle");
        var hiraganaCheckbox = sizePanel.add("checkbox", undefined, getLabel("checkbox.hiragana"));
        hiraganaCheckbox.helpTip = getLabel("tooltip.hiragana");

        /* サイズ：［　］% の入力行 / Size: [ ] % input row */
        var sizeRow = sizePanel.add("group");
        setupGroup(sizeRow, "row", SIZE_ROW_SPACING);
        sizeRow.add("statictext", undefined, labelText("fieldLabel.size"));
        var sizeField = sizeRow.add("edittext", undefined, String(DEFAULT_SIZE_PERCENT));
        sizeField.characters = SIZE_FIELD_CHARS;
        sizeField.helpTip = getLabel("tooltip.size");
        sizeRow.add("statictext", undefined, "%");

        /* ボタン（Mac 規約: キャンセル → OK、OK は右）/ Buttons (Mac convention: Cancel → OK, OK on the right) */
        var btnRowGroup = lineBreakDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));

        return {
            lineBreakDialog: lineBreakDialog,
            targetCheckboxes: targetCheckboxes,
            caseParticleCheckbox: caseParticleCheckbox,
            hiraganaCheckbox: hiraganaCheckbox,
            sizeField: sizeField,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを表示し、OK で改行とサイズ調整を行う
     * @returns {void}
     */
    function main() {
        /* 事前チェック / Pre-flight checks */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        if (app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var dialogControls = buildTitleMakerDialog();
        var lineBreakDialog = dialogControls.lineBreakDialog;

        dialogControls.btnCancel.onClick = function () {
            lineBreakDialog.close();
        };

        dialogControls.btnOK.onClick = function () {
            /* チェックされた改行対象文字を収集 / Collect the checked line-break characters */
            var selectedMarks = [];
            for (var i = 0; i < dialogControls.targetCheckboxes.length; i++) {
                if (dialogControls.targetCheckboxes[i].value) selectedMarks.push(TARGET_CHARS[i].mark);
            }

            /* サイズ調整の設定を収集 / Collect the size-adjustment settings */
            var sizeOptions = {
                adjustCaseParticle: dialogControls.caseParticleCheckbox.value,
                adjustHiragana: dialogControls.hiraganaCheckbox.value,
                sizePercent: parseFloat(dialogControls.sizeField.text)
            };
            var wantsSizeAdjust = sizeOptions.adjustCaseParticle || sizeOptions.adjustHiragana;

            if (selectedMarks.length === 0 && !wantsSizeAdjust) {
                alert(getLabel("alert.noTarget"));
                return;
            }
            if (wantsSizeAdjust && (isNaN(sizeOptions.sizePercent) || sizeOptions.sizePercent <= 0)) {
                alert(getLabel("alert.invalidSize"));
                return;
            }

            processSelection(selectedMarks, wantsSizeAdjust ? sizeOptions : null);
            lineBreakDialog.close();
        };

        lineBreakDialog.center();
        lineBreakDialog.show();
    }

    main();

})();
