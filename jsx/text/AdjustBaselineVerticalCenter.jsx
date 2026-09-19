#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定した文字を、基準文字の中心に合わせてベースラインシフトで上下に動かします。
複数のテキストフレームへまとめて適用できます。

詳細は README を参照してください。

### Overview

Shifts the specified characters up or down with a baseline shift so they line up with the center of a reference character.
It can be applied to several text frames at once.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustBaselineVerticalCenter"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-05";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustBaselineVerticalCenter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustBaselineVerticalCenter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考、謝辞 / Reference and acknowledgements
 * Egor Chistyakov (@tchegr)
 * https://x.com/tchegr
 */

/* ロケール判定 / Locale detection */
function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日本語 / English */
var LABELS = {
    dialogTitle: { ja: "ベースライン調整", en: "Adjust Baseline" },
    infoTextMsg: { ja: "対象文字を縦方向に揃えます。", en: "This will align selected symbol vertically." },
    targetCharLabel: { ja: "対象文字:", en: "Target Character:" },
    baseCharLabel: { ja: "基準文字:", en: "Reference Character:" },
    okBtnLabel: { ja: "調整", en: "Adjust" },
    cancelBtnLabel: { ja: "キャンセル", en: "Cancel" },
    selectFrameMsg: { ja: "テキストフレームを選択してください。", en: "Select one or more text frames." },
    docOpenMsg: { ja: "ドキュメントが開かれていません。", en: "No document open." },
    invalidCharMsg: { ja: "対象文字は1文字以上、基準文字は1文字を入力してください。", en: "Enter at least one target character and exactly one reference character." },
    notFoundMsg: { ja: "対象文字が含まれていません。", en: "Target character not found." },
    errorMsg: { ja: "エラー: ", en: "Error: " }
};

/* アイテムのジオメトリック境界から中心のY座標を計算して返す / Calculate and return the center Y coordinate from item's geometric bounds */
function getCenterY(item) {
    /*
    アイテムのジオメトリック境界から中心のY座標を計算して返す。
    */
    var bounds = item.geometricBounds;
    return bounds[1] - (bounds[1] - bounds[3]) / 2;
}

/* 指定文字を含むテキストフレームを複製し、アウトライン化後に中心Y座標を取得 / Duplicate text frame with specified character, outline it, then get center Y */
function createOutlineAndGetCenterY(textFrame, character) {
    /*
    指定文字を含むテキストフレームを複製し、アウトライン化後に中心Y座標を取得する。
    */
    var tempFrame = textFrame.duplicate();
    tempFrame.contents = character;
    tempFrame.filled = false;
    tempFrame.stroked = false;
    var outline = tempFrame.createOutline();
    var centerY = getCenterY(outline);
    outline.remove();
    return centerY;
}

/* 選択テキスト内の記号・非英数字の出現頻度を集計して返す / Count frequency of symbols and non-alphanumeric chars in selection */
function getSymbolFrequency(sel) {
    /*
    選択テキスト内の記号・非英数字の出現頻度を集計して返す。
    */
    var charCount = {};
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename == "TextFrame") {
            var selText = item.contents;
            for (var j = 0; j < selText.length; j++) {
                var c = selText.charAt(j);
                if (!c.match(/^[A-Za-z0-9\s]$/)) {
                    charCount[c] = (charCount[c] || 0) + 1;
                }
            }
        }
    }
    return charCount;
}

/* ユーザーに対象文字と基準文字を入力させるダイアログを表示し、入力結果を返す / Show dialog to input target and reference characters, return input */
function showDialog() {
    /*
    ユーザーに対象文字と基準文字を入力させるダイアログを表示し、入力結果を返す。
    */
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    var infoText = dialog.add("statictext", undefined, LABELS.infoTextMsg[lang]);
    infoText.alignment = "left";

    /* デフォルト対象文字（複数選択でも最頻出記号を抽出） / Default target character (most frequent symbol even in multiple selection) */
    var defaultTarget = "";
    var sel = app.activeDocument.selection;
    if (sel && sel.length > 0) {
        var charCount = getSymbolFrequency(sel);
        var maxCount = 0;
        for (var key in charCount) {
            if (charCount[key] > maxCount) {
                maxCount = charCount[key];
                defaultTarget = key;
            }
        }
    }

    var inputGroup = dialog.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";
    inputGroup.margins = [15, 5, 15, 5];

    var targetGroup = inputGroup.add("group");
    targetGroup.add("statictext", undefined, LABELS.targetCharLabel[lang]);
    var targetInput = targetGroup.add("edittext", undefined, defaultTarget);
    targetInput.characters = 5;
    targetInput.active = true;

    var refGroup = inputGroup.add("group");
    refGroup.add("statictext", undefined, LABELS.baseCharLabel[lang]);
    var refInput = refGroup.add("edittext", undefined, "0");
    refInput.characters = 5;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancelBtnLabel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.okBtnLabel[lang], { name: "ok" });

    okBtn.onClick = function () {
        if (targetInput.text.length == 0) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        if (refInput.text.length != 1) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        dialog.close(1);
    };
    cancelBtn.onClick = function () { dialog.close(0); };

    if (dialog.show() == 1) {
        return { target: targetInput.text, reference: refInput.text };
    }
    return null;
}

/* メイン処理。ドキュメントと選択状態をチェックし、ダイアログを表示して調整を実行 / Main process: check document and selection, show dialog and apply adjustments */
function main() {
    /*
    メイン処理。ドキュメントと選択状態をチェックし、ダイアログを表示して調整を実行。
    */
    try {
        if (app.documents.length == 0) {
            alert(LABELS.docOpenMsg[lang]);
            return;
        }

        var selection = app.activeDocument.selection;
        if (!selection || selection.length == 0) {
            alert(LABELS.selectFrameMsg[lang]);
            return;
        }

        var input = showDialog();
        if (!input) return;

        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename == "TextFrame") {
                var contents = item.contents;
                for (var c = 0; c < input.target.length; c++) {
                    var targetChar = input.target.charAt(c);
                    if (contents.indexOf(targetChar) == -1) {
                        continue;
                    }

                    var refCenterY = createOutlineAndGetCenterY(item, input.reference);
                    var targetCenterY = createOutlineAndGetCenterY(item, targetChar);
                    var yOffset = targetCenterY - refCenterY;

                    var chars = item.textRange.characters;
                    for (var j = 0; j < chars.length; j++) {
                        var ch = chars[j].contents;
                        if (ch && ch === targetChar) {
                            chars[j].characterAttributes.baselineShift = -yOffset;
                        }
                    }
                }
            }
        }

    } catch (e) {
        alert(LABELS.errorMsg[lang] + e);
    }
}

/* メイン処理の実行 / Execute main process */
main();