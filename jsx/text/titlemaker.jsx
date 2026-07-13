#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択中のテキストフレームを、指定した文字の直後で改行する
- 改行対象の文字（、。〜）をチェックボックスで選択
- 改行後、連続する改行を1つにまとめる
- 文字揃えを欧文ベースラインに設定

### 処理の流れ

1. ダイアログで改行対象文字を選択
2. 選択中の各テキストフレームの内容を走査
3. 対象文字の後ろに改行を挿入し、重複改行を整理
4. 欧文ベースラインを適用

*/

/*

### Overview

- Insert a line break after specified characters in selected text frames
- Choose target characters (、。〜) via checkboxes
- Collapse consecutive line breaks into one
- Set character alignment to the Roman baseline

### Flow

1. Choose target characters in the dialog
2. Walk through each selected text frame's contents
3. Insert a line break after each target character and clean up duplicates
4. Apply the Roman baseline

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ユーザー設定 / User Settings
// =========================================
/* 改行対象文字とデフォルトの ON/OFF / Target characters and default ON/OFF */
var TARGET_CHARS = [
    { mark: "、", on: true },
    { mark: "。", on: true },
    { mark: "〜", on: false }
];

/* 格助詞本体の正規表現ソース（長い候補を先に）/ Regex source for the case particle itself (longer alternatives first) */
var CASE_PARTICLE_SOURCE = "(?:から|より|が|を|に|へ|と|で|の)";

/* 助詞の直前に許される文字クラス（漢字・ひらがな・カタカナ・英数字）。ExtendScript は lookbehind 非対応のため手動判定に使う / Allowed preceding-char class for a particle (kanji / hiragana / katakana / alphanumeric); used manually because ExtendScript lacks lookbehind */
var PARTICLE_PRECURSOR_SOURCE = "[一-龯ぁ-んァ-ヶA-Za-z0-9]";

/* サイズ調整のデフォルト倍率（%）/ Default scale for size adjustment (%) */
var DEFAULT_SIZE_PERCENT = 80;

// =========================================
// ローカライズ / Localization
// =========================================
/* 現在の言語を取得（ja で始まれば日本語、それ以外は英語）/ Get current language (ja-prefixed = Japanese, otherwise English) */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

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
    field: {
        size: { ja: "サイズ", en: "Size" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
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

/* ドット区切りキーでラベルを取得 / Resolve a label by dot-separated key */
function L(keyPath) {
    var parts = keyPath.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        node = node[parts[i]];
    }
    return node[currentLanguage];
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(keyPath) {
    return L(keyPath) + (currentLanguage === "ja" ? "：" : ":");
}

/* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
function labelWithCount(keyPath, count) {
    if (currentLanguage === "ja") {
        return L(keyPath) + "（" + count + "）";
    }
    return L(keyPath) + " (" + count + ")";
}

// =========================================
// UI レイアウト設定 / UI Layout
// =========================================
/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
function setupGroup(group, orientation, spacing) {
    var groupOrientation = orientation || "column";
    group.orientation = groupOrientation;
    /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
    group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// 文字種の判定 / Character Matching
// =========================================
/* ひらがなかどうか（U+3041〜U+309F）/ Whether the glyph is hiragana (U+3041–U+309F) */
function isHiragana(glyph) {
    var code = glyph.charCodeAt(0);
    return code >= 0x3041 && code <= 0x309F;
}

/* 格助詞として縮小するグリフのインデックス集合を返す / Return the set of glyph indices to scale as case particles */
function findCaseParticleIndices(text) {
    var marked = {};
    var pattern = new RegExp(CASE_PARTICLE_SOURCE, "g");
    var precursor = new RegExp(PARTICLE_PRECURSOR_SOURCE);
    var match;

    while ((match = pattern.exec(text)) !== null) {
        var start = match.index;
        /* lookbehind の代わりに直前の1文字を判定 / Emulate lookbehind by testing the single preceding char */
        var prevChar = (start > 0) ? text.charAt(start - 1) : "";
        if (prevChar !== "" && precursor.test(prevChar)) {
            for (var p = 0; p < match[0].length; p++) marked[start + p] = true;
        }
        /* 空マッチによる無限ループを防ぐ / Guard against infinite loops on zero-length matches */
        if (match.index === pattern.lastIndex) pattern.lastIndex++;
    }
    return marked;
}

(function () {

    /* 事前チェック / Pre-flight checks */
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }
    if (app.selection.length === 0) {
        alert(L("alert.noSelection"));
        return;
    }

    /* ダイアログを構築（タイトルバーにバージョンを表示）/ Build the dialog (version shown in the title bar) */
    var lineBreakDialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    lineBreakDialog.orientation = "column";
    lineBreakDialog.alignChildren = "fill";

    /* 改行対象文字のチェックボックスパネル / Checkbox panel for target characters */
    var targetPanel = lineBreakDialog.add("panel", undefined, L("panel.targets"));
    setupPanel(targetPanel, 6);

    var targetCheckboxes = [];
    for (var i = 0; i < TARGET_CHARS.length; i++) {
        var targetCheckbox = targetPanel.add("checkbox", undefined, TARGET_CHARS[i].mark);
        targetCheckbox.value = TARGET_CHARS[i].on;
        targetCheckboxes.push(targetCheckbox);
    }

    /* フォントサイズ調整のパネル / Panel for font size adjustment */
    var sizePanel = lineBreakDialog.add("panel", undefined, L("panel.sizeAdjust"));
    setupPanel(sizePanel, 6);

    var caseParticleCheckbox = sizePanel.add("checkbox", undefined, L("checkbox.caseParticle"));
    var hiraganaCheckbox = sizePanel.add("checkbox", undefined, L("checkbox.hiragana"));

    /* サイズ：［　］% の入力行 / Size: [ ] % input row */
    var sizeGroup = sizePanel.add("group");
    setupGroup(sizeGroup, "row", 4);
    sizeGroup.add("statictext", undefined, labelText("field.size"));
    var sizeField = sizeGroup.add("edittext", undefined, String(DEFAULT_SIZE_PERCENT));
    sizeField.characters = 4;
    sizeGroup.add("statictext", undefined, "%");

    /* ボタン（Mac 規約: キャンセル → OK、OK は右）/ Buttons (Mac convention: Cancel → OK, OK on the right) */
    var buttonGroup = lineBreakDialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "right";

    var cancelButton = buttonGroup.add("button", undefined, L("button.cancel"));
    var okButton = buttonGroup.add("button", undefined, "OK");

    cancelButton.onClick = function () {
        lineBreakDialog.close();
    };

    okButton.onClick = function () {
        /* チェックされた改行対象文字を収集 / Collect the checked line-break characters */
        var selectedMarks = [];
        for (var j = 0; j < targetCheckboxes.length; j++) {
            if (targetCheckboxes[j].value) selectedMarks.push(TARGET_CHARS[j].mark);
        }

        /* サイズ調整の設定を収集 / Collect the size-adjustment settings */
        var sizeOptions = {
            adjustCaseParticle: caseParticleCheckbox.value,
            adjustHiragana: hiraganaCheckbox.value,
            sizePercent: parseFloat(sizeField.text)
        };
        var wantsSizeAdjust = sizeOptions.adjustCaseParticle || sizeOptions.adjustHiragana;

        if (selectedMarks.length === 0 && !wantsSizeAdjust) {
            alert(L("alert.noTarget"));
            return;
        }
        if (wantsSizeAdjust && (isNaN(sizeOptions.sizePercent) || sizeOptions.sizePercent <= 0)) {
            alert(L("alert.invalidSize"));
            return;
        }

        processSelection(selectedMarks, wantsSizeAdjust ? sizeOptions : null);
        lineBreakDialog.close();
    };

    lineBreakDialog.center();
    lineBreakDialog.show();

    /* 選択中の各テキストフレームに改行挿入とサイズ調整を適用 / Apply line breaks and size adjustment to each selected text frame */
    function processSelection(marks, sizeOptions) {
        for (var i = 0; i < app.selection.length; i++) {
            var selectedItem = app.selection[i];

            if (selectedItem.typename !== "TextFrame") continue;

            if (marks.length > 0) insertLineBreaksAfterMarks(selectedItem, marks);
            if (sizeOptions) adjustCharacterSizes(selectedItem, sizeOptions);
        }
    }

    /* 対象文字の後ろで改行し、欧文ベースラインを適用 / Insert line breaks after target characters and apply the Roman baseline */
    function insertLineBreaksAfterMarks(textFrame, marks) {
        var frameContents = textFrame.contents;

        /* 各対象文字の後ろに改行を挿入 / Insert a line break after each target character */
        for (var j = 0; j < marks.length; j++) {
            var mark = marks[j];
            frameContents = frameContents.split(mark).join(mark + "\r");
        }

        /* 連続改行を整理 / Collapse consecutive line breaks */
        frameContents = frameContents.replace(/\r\r+/g, "\r");

        textFrame.contents = frameContents;

        /* 文字揃え：欧文ベースライン / Character alignment: Roman baseline */
        textFrame.textRange.characterAttributes.baselinePosition =
            FontBaselineOption.NORMALBASELINE;
    }

    /* 格助詞・ひらがなのフォントサイズを指定倍率に縮小 / Scale font size of case particles / hiragana to the given percentage */
    function adjustCharacterSizes(textFrame, sizeOptions) {
        var scale = sizeOptions.sizePercent / 100;
        var text = textFrame.contents;
        var glyphs = textFrame.textRange.characters;

        /* 縮小対象グリフのインデックス集合（文字列インデックスと glyphs は 1:1 対応）/ Indices of glyphs to scale (string index maps 1:1 to glyphs) */
        var marked = sizeOptions.adjustCaseParticle ? findCaseParticleIndices(text) : {};

        /* ひらがなは1文字ずつ判定して追加 / Add hiragana per character */
        if (sizeOptions.adjustHiragana) {
            for (var h = 0; h < text.length; h++) {
                if (isHiragana(text.charAt(h))) marked[h] = true;
            }
        }

        /* マーク済みのグリフを縮小 / Scale the marked glyphs */
        for (var c = 0; c < glyphs.length; c++) {
            if (marked[c]) {
                var attributes = glyphs[c].characterAttributes;
                attributes.size = attributes.size * scale;
            }
        }
    }
})();
