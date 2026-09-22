#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定した文字を、基準文字の中心に合わせてベースラインシフトで上下に動かします。
複数のテキストフレームへまとめて適用できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustBaselineVerticalCenter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na7a8c907c68c

### Overview

Shifts the specified characters up or down with a baseline shift so they line up with the center of a reference character.
It can be applied to several text frames at once.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustBaselineVerticalCenter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustBaselineVerticalCenter"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustBaselineVerticalCenter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustBaselineVerticalCenter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na7a8c907c68c"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考、謝辞 / Reference and acknowledgements
 * Egor Chistyakov (@tchegr)
 * https://x.com/tchegr
 */

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_REFERENCE_CHAR = "0"; /* 基準文字の初期値 / Initial reference character */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var INPUT_GROUP_MARGINS = [15, 5, 15, 5]; /* 入力欄グループの余白 [左,上,右,下] / Input group margins */
    var CHAR_INPUT_CHARACTERS = 5;            /* 文字入力欄の幅（文字数）/ Width of the character fields (in characters) */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "ベースライン調整", en: "Adjust Baseline" },
            description: { ja: "対象文字を縦方向に揃えます。", en: "This will align selected symbol vertically." }
        },
        fieldLabel: {
            targetChar: { ja: "対象文字", en: "Target Character" },
            referenceChar: { ja: "基準文字", en: "Reference Character" }
        },
        tooltip: {
            targetChar: {
                ja: "入力した文字をベースラインシフトで上下に動かします（複数可）。初期値は、選択中のテキストで英数字と空白を除いていちばん多い文字です。",
                en: "Moves these characters up or down with baseline shift (you can enter several). Defaults to the most frequent character in the selection, excluding letters, digits and spaces."
            },
            referenceChar: {
                ja: "対象文字の天地中央を、この文字の天地中央にそろえます（1文字）。",
                en: "Aligns the vertical center of the target characters with the vertical center of this character (one character)."
            }
        },
        button: {
            adjust: { ja: "調整", en: "Adjust" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            selectTextFrame: { ja: "テキストフレームを選択してください。", en: "Select one or more text frames." },
            invalidChars: {
                ja: "対象文字は1文字以上、基準文字は1文字を入力してください。",
                en: "Enter at least one target character and exactly one reference character."
            },
            errorPrefix: { ja: "エラー: ", en: "Error: " }
        }
    };

    /**
     * 現在の言語のラベルを取得する
     * @param {object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} コロン付きのラベル文字列
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 文字の中心の実測 / Measuring character centers
    // =========================================

    /**
     * アイテムの天地中央のY座標を求める
     * @param {PageItem} pageItem - 対象のアイテム
     * @returns {number} 天地中央のY座標
     */
    function getCenterY(pageItem) {
        var itemBounds = pageItem.geometricBounds;
        return (itemBounds[1] + itemBounds[3]) / 2;
    }

    /**
     * テキストフレームを複製して1文字だけにし、アウトライン化した字形の天地中央を測る
     * @param {TextFrame} textFrame - 書式の元になるテキストフレーム
     * @param {string} character - 測る文字
     * @returns {number} 字形の天地中央のY座標
     */
    function measureCharCenterY(textFrame, character) {
        var tempFrame = textFrame.duplicate();
        var outlineGroup = null;
        try {
            tempFrame.contents = character;
            outlineGroup = tempFrame.createOutline(); /* 複製はここで消費される / The duplicate is consumed here */
            return getCenterY(outlineGroup);
        } finally {
            /* 途中で失敗しても一時オブジェクトを残さない / Never leave the temporary objects behind, even on failure */
            if (outlineGroup) outlineGroup.remove();
            else tempFrame.remove();
        }
    }

    // =========================================
    // 対象の収集 / Collecting targets
    // =========================================

    /**
     * 選択からテキストフレームを取り出す（文字の編集中は対象なし）
     * @param {Array<PageItem>|TextRange} selection - ドキュメントの選択
     * @returns {TextFrame[]} 選択中のテキストフレーム
     */
    function collectTextFrames(selection) {
        var textFrames = [];
        /* 文字の編集中は選択が TextRange になる / While editing text, the selection is a TextRange */
        if (!selection || selection.typename === "TextRange") return textFrames;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") textFrames.push(selection[i]);
        }
        return textFrames;
    }

    /**
     * 英数字と空白を除いた文字のうち、いちばん多く出てくる文字を求める（対象文字の初期値）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {string} 最も多い文字（見つからなければ空文字）
     */
    function findDefaultTargetChar(textFrames) {
        var charCounts = {};
        for (var i = 0; i < textFrames.length; i++) {
            var frameText = textFrames[i].contents;
            for (var j = 0; j < frameText.length; j++) {
                var character = frameText.charAt(j);
                if (!/^[A-Za-z0-9\s]$/.test(character)) charCounts[character] = (charCounts[character] || 0) + 1;
            }
        }

        var mostFrequentChar = "";
        var maxCount = 0;
        for (var countedChar in charCounts) {
            if (charCounts[countedChar] > maxCount) {
                maxCount = charCounts[countedChar];
                mostFrequentChar = countedChar;
            }
        }
        return mostFrequentChar;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名＋文字入力欄の1行を追加する
     * @param {Group} parentGroup - 追加先
     * @param {object} labelSet - 項目名のラベル
     * @param {object} tooltipLabelSet - 入力欄の tooltip のラベル
     * @param {string} initialText - 入力欄の初期値
     * @returns {EditText} 追加した入力欄
     */
    function addCharField(parentGroup, labelSet, tooltipLabelSet, initialText) {
        var fieldRow = parentGroup.add("group");
        fieldRow.add("statictext", undefined, labelText(labelSet));
        var charInput = fieldRow.add("edittext", undefined, initialText);
        charInput.characters = CHAR_INPUT_CHARACTERS;
        charInput.helpTip = getLabel(tooltipLabelSet);
        return charInput;
    }

    /**
     * 対象文字と基準文字を入力するダイアログを表示する
     * @param {string} defaultTargetChar - 対象文字の初期値
     * @returns {{targetChars: string, referenceChar: string}|null} 入力された文字（キャンセル時は null）
     */
    function showCharDialog(defaultTargetChar) {
        var baselineDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        baselineDialog.orientation = "column";
        baselineDialog.alignChildren = "left";

        baselineDialog.add("statictext", undefined, getLabel(LABELS.dialog.description));

        var inputGroup = baselineDialog.add("group");
        inputGroup.orientation = "column";
        inputGroup.alignChildren = "left";
        inputGroup.margins = INPUT_GROUP_MARGINS;

        var targetCharInput = addCharField(inputGroup, LABELS.fieldLabel.targetChar, LABELS.tooltip.targetChar, defaultTargetChar);
        targetCharInput.active = true;
        var referenceCharInput = addCharField(inputGroup, LABELS.fieldLabel.referenceChar, LABELS.tooltip.referenceChar, DEFAULT_REFERENCE_CHAR);

        /* ボタンエリア（左右中央）/ Button area (centered) */
        var btnRowGroup = baselineDialog.add("group");
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.adjust), { name: "ok" });

        btnOK.onClick = function () {
            if (targetCharInput.text.length === 0 || referenceCharInput.text.length !== 1) {
                alert(getLabel(LABELS.alert.invalidChars));
                return;
            }
            baselineDialog.close(1);
        };
        btnCancel.onClick = function () { baselineDialog.close(0); };

        if (baselineDialog.show() !== 1) return null;
        return { targetChars: targetCharInput.text, referenceChar: referenceCharInput.text };
    }

    // =========================================
    // ベースラインの調整 / Baseline adjustment
    // =========================================

    /**
     * フレーム内の指定文字すべてにベースラインシフトを設定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} targetChar - 対象の文字
     * @param {number} shiftAmount - ベースラインシフトの値（pt）
     * @returns {void}
     */
    function applyBaselineShift(textFrame, targetChar, shiftAmount) {
        var frameChars = textFrame.textRange.characters;
        for (var i = 0; i < frameChars.length; i++) {
            if (frameChars[i].contents === targetChar) frameChars[i].characterAttributes.baselineShift = shiftAmount;
        }
    }

    /**
     * 1つのフレームで、対象文字それぞれの天地中央を基準文字の天地中央にそろえる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（複数可）
     * @param {string} referenceChar - 基準文字（1文字）
     * @returns {void}
     */
    function adjustTextFrame(textFrame, targetChars, referenceChar) {
        var frameText = textFrame.contents;

        /* 先にすべて測ってから適用する。測定用の複製は先頭文字の書式を引き継ぐため、
           先頭文字をずらしたあとに測ると基準文字の測定値と食い違う
           Measure everything first: the measuring duplicate inherits the first character's formatting,
           so measuring after that character has been shifted would disagree with the reference measurement */
        var referenceCenterY = null; /* フレームごとに1回だけ測る / Measured once per frame */
        var charShifts = [];
        for (var i = 0; i < targetChars.length; i++) {
            var targetChar = targetChars.charAt(i);
            if (frameText.indexOf(targetChar) === -1) continue;

            if (referenceCenterY === null) referenceCenterY = measureCharCenterY(textFrame, referenceChar);
            charShifts.push({ targetChar: targetChar, shiftAmount: referenceCenterY - measureCharCenterY(textFrame, targetChar) });
        }

        for (var j = 0; j < charShifts.length; j++) {
            applyBaselineShift(textFrame, charShifts[j].targetChar, charShifts[j].shiftAmount);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前提を確かめてダイアログを表示し、選択中のテキストフレームを調整する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var textFrames = collectTextFrames(app.activeDocument.selection);
        if (textFrames.length === 0) {
            alert(getLabel(LABELS.alert.selectTextFrame));
            return;
        }

        var charSettings = showCharDialog(findDefaultTargetChar(textFrames));
        if (!charSettings) return;

        /* 字形を持たない文字（スペースなど）はアウトライン化で例外になるため、ここで知らせる
           Characters with no glyph (spaces etc.) throw during outlining, so report it here */
        try {
            for (var i = 0; i < textFrames.length; i++) {
                adjustTextFrame(textFrames[i], charSettings.targetChars, charSettings.referenceChar);
            }
        } catch (e) {
            alert(getLabel(LABELS.alert.errorPrefix) + e);
        }
    }

    main();

})();
