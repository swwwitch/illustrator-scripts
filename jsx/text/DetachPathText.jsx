#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

パス上文字を、テキストとパスに分離します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DetachPathText.md

### Overview

Detaches text on a path back into a separate text object and path.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DetachPathText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DetachPathText";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DetachPathText.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DetachPathText.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 18;                   /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 12];      /* パネルの余白 / panel margins */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "パス上文字の解除", en: "Detach Path Text" }
        },
        panel: {
            textFormat: { ja: "テキストの書式", en: "Text Formatting" },
            path: { ja: "パス", en: "Path" }
        },
        radio: {
            formatFull: { ja: "完全に保持", en: "Preserve Completely" },
            formatFast: { ja: "高速に保持", en: "Preserve Quickly" },
            formatNone: { ja: "削除", en: "Remove" },
            pathBlackStroke: { ja: "1pt黒に設定", en: "Set 1pt Black Stroke" },
            pathNoStroke: { ja: "線なし", en: "No Stroke" },
            pathDelete: { ja: "削除", en: "Delete" }
        },
        tooltip: {
            formatFull: {
                ja: "1文字ずつ書式を写します。正確ですが、文字数が多いと時間がかかります。",
                en: "Copies the formatting character by character. Accurate, but slow for long text."
            },
            formatFast: {
                ja: "テキスト全体の書式をまとめて写します。速いかわりに、文字ごとの違いは失われます。",
                en: "Copies the formatting for the whole text at once. Faster, but per-character differences are lost."
            },
            formatNone: { ja: "書式を写さず、文字だけを取り出します。", en: "Takes only the characters, without the formatting." },
            pathBlackStroke: { ja: "残したパスに1ptの黒い線を設定します。", en: "Gives the remaining path a 1 pt black stroke." },
            pathNoStroke: { ja: "残したパスを線なしにします。", en: "Leaves the remaining path without a stroke." },
            pathDelete: { ja: "パスを残さず削除します。", en: "Deletes the path instead of keeping it." }
        },
        button: {
            cancel: { ja: "閉じる", en: "Close" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "パス上文字を選択してください。", en: "Please select path text." },
            noPathText: { ja: "選択範囲にパス上文字が含まれていません。", en: "No path text found in selection." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // ==========================================
    // UI: ダイアログ / Dialog
    // ==========================================

    /**
     * ラジオボタンを縦に並べるパネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} titlePath - パネル見出しのラベルパス
     * @returns {Panel} 追加したパネル
     */
    function addRadioPanel(parentWindow, titlePath) {
        var radioPanel = parentWindow.add("panel", undefined, getLabel(titlePath));
        radioPanel.orientation = "column";
        radioPanel.alignChildren = ["left", "top"];
        radioPanel.margins = PANEL_MARGINS;
        return radioPanel;
    }

    /**
     * ツールチップ付きのラジオボタンを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルパス
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadioWithTip(parentPanel, labelPath, tooltipPath) {
        var radioButton = parentPanel.add("radiobutton", undefined, getLabel(labelPath));
        radioButton.helpTip = getLabel(tooltipPath);
        return radioButton;
    }

    /**
     * 処理方法を選ぶダイアログを表示する
     * @returns {Object|null} 処理設定（textFormatMode / pathBlackStroke / pathNoStroke / pathDelete）。キャンセル時は null
     */
    function showOptionsDialog() {
        var optionsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        optionsDialog.orientation = "column";
        optionsDialog.alignChildren = ["fill", "top"];
        optionsDialog.margins = DIALOG_MARGINS;

        /* テキストパネル / Text panel */
        var textFormatPanel = addRadioPanel(optionsDialog, "panel.textFormat");
        var rbFormatFull = addRadioWithTip(textFormatPanel, "radio.formatFull", "tooltip.formatFull");
        var rbFormatFast = addRadioWithTip(textFormatPanel, "radio.formatFast", "tooltip.formatFast");
        var rbFormatNone = addRadioWithTip(textFormatPanel, "radio.formatNone", "tooltip.formatNone");
        rbFormatFull.value = true;

        /* パスパネル / Path panel */
        var pathPanel = addRadioPanel(optionsDialog, "panel.path");
        var rbPathBlackStroke = addRadioWithTip(pathPanel, "radio.pathBlackStroke", "tooltip.pathBlackStroke");
        var rbPathNoStroke = addRadioWithTip(pathPanel, "radio.pathNoStroke", "tooltip.pathNoStroke");
        var rbPathDelete = addRadioWithTip(pathPanel, "radio.pathDelete", "tooltip.pathDelete");
        rbPathBlackStroke.value = true; /* デフォルト / Default */

        /* ボタングループ / Button group */
        var btnRowGroup = optionsDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["right", "center"];
        btnRowGroup.alignment = ["fill", "top"];

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOk = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var detachOptions = null;
        btnOk.onClick = function () {
            detachOptions = {
                textFormatMode: rbFormatNone.value ? "none" : (rbFormatFull.value ? "full" : "fast"),
                pathBlackStroke: rbPathBlackStroke.value,
                pathNoStroke: rbPathNoStroke.value,
                pathDelete: rbPathDelete.value
            };
            optionsDialog.close(1);
        };
        btnCancel.onClick = function () {
            optionsDialog.close(0);
        };

        return (optionsDialog.show() === 1) ? detachOptions : null;
    }

    // ==========================================
    // 書式の控えと復元 / Formatting snapshot and restore
    // ==========================================

    /* 控える文字属性（読み取り順） / Character attributes to snapshot, in reading order */
    var SNAPSHOT_ATTRIBUTE_NAMES = ["textFont", "size", "fillColor", "strokeColor", "strokeWeight", "tracking",
        "baselineShift", "horizontalScale", "verticalScale", "autoLeading", "leading"];

    /* 線より前・後に書き戻す属性（書き戻し順） / Attributes written back before / after the stroke, in writing order */
    var ATTRIBUTES_BEFORE_STROKE = ["textFont", "size", "fillColor"];
    var ATTRIBUTES_AFTER_STROKE = ["tracking", "baselineShift", "horizontalScale", "verticalScale", "autoLeading"];

    /**
     * パス上文字の1文字ずつの文字属性を控える
     * @param {TextFrame} pathText - 元のパス上文字
     * @returns {Object[]} 文字ごとの属性の控え（キーは characterAttributes のプロパティ名）
     */
    function snapshotCharacterAttributes(pathText) {
        var attributeSnapshots = [];
        for (var charIndex = 0; charIndex < pathText.characters.length; charIndex++) {
            var sourceAttributes = pathText.characters[charIndex].characterAttributes;
            var attributeSnapshot = {};
            for (var attrIndex = 0; attrIndex < SNAPSHOT_ATTRIBUTE_NAMES.length; attrIndex++) {
                attributeSnapshot[SNAPSHOT_ATTRIBUTE_NAMES[attrIndex]] = sourceAttributes[SNAPSHOT_ATTRIBUTE_NAMES[attrIndex]];
            }
            attributeSnapshots.push(attributeSnapshot);
        }
        return attributeSnapshots;
    }

    /**
     * 文字属性を1つずつ書き戻す。書けない属性は飛ばして続ける
     * @param {CharacterAttributes} targetAttributes - 書き戻し先の文字属性
     * @param {Object} attributeSnapshot - 控えた属性
     * @param {string[]} attributeNames - 書き戻す属性名（この順で書く）
     * @returns {void}
     */
    function writeAttributesSafely(targetAttributes, attributeSnapshot, attributeNames) {
        for (var attrIndex = 0; attrIndex < attributeNames.length; attrIndex++) {
            /* 値によっては DOM が代入を拒む / the DOM may reject some values */
            try { targetAttributes[attributeNames[attrIndex]] = attributeSnapshot[attributeNames[attrIndex]]; } catch (e) { }
        }
    }

    /**
     * 控えた文字属性を新しいテキストの1文字ずつに書き戻す
     * @param {TextFrame} newText - 書き戻し先のポイント文字
     * @param {Object[]} attributeSnapshots - snapshotCharacterAttributes() の結果
     * @returns {void}
     */
    function restoreCharacterAttributes(newText, attributeSnapshots) {
        for (var charIndex = 0; charIndex < newText.characters.length; charIndex++) {
            var targetAttributes = newText.characters[charIndex].characterAttributes;
            var attributeSnapshot = attributeSnapshots[charIndex];
            if (!attributeSnapshot) break;

            writeAttributesSafely(targetAttributes, attributeSnapshot, ATTRIBUTES_BEFORE_STROKE);
            /* stroke: 元が「なし」の場合は strokeWeight を 0 に / stroke: set weight to 0 if original had no stroke */
            try {
                var sourceStrokeColor = attributeSnapshot.strokeColor;
                targetAttributes.strokeColor = sourceStrokeColor;
                if (sourceStrokeColor && sourceStrokeColor.typename === "NoColor") {
                    targetAttributes.strokeWeight = 0;
                } else {
                    targetAttributes.strokeWeight = attributeSnapshot.strokeWeight;
                }
            } catch (e) { }
            writeAttributesSafely(targetAttributes, attributeSnapshot, ATTRIBUTES_AFTER_STROKE);

            /* autoLeading が false の場合のみ leading を設定 / Set leading only when autoLeading is false */
            if (!attributeSnapshot.autoLeading) {
                writeAttributesSafely(targetAttributes, attributeSnapshot, ["leading"]);
            }
        }
    }

    // ==========================================
    // 分離処理 / Detach
    // ==========================================

    /**
     * 選択からパス上文字だけを取り出す
     * @param {Array} selectedItems - ドキュメントの選択
     * @returns {TextFrame[]} パス上文字
     */
    function collectPathTexts(selectedItems) {
        var pathTexts = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame" && selectedItems[i].kind === TextType.PATHTEXT) {
                pathTexts.push(selectedItems[i]);
            }
        }
        return pathTexts;
    }

    /**
     * パス上文字のパスを、アンカーポイントを写して新しいパスとして作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} originalPath - パス上文字の textPath
     * @param {Object} detachOptions - showOptionsDialog() の結果
     * @returns {void}
     */
    function duplicateTextPath(doc, originalPath, detachOptions) {
        var newPath = doc.pathItems.add();

        for (var k = 0; k < originalPath.pathPoints.length; k++) {
            var originalPoint = originalPath.pathPoints[k];
            var newPoint = newPath.pathPoints.add();
            newPoint.anchor = originalPoint.anchor;
            newPoint.leftDirection = originalPoint.leftDirection;
            newPoint.rightDirection = originalPoint.rightDirection;
            newPoint.pointType = originalPoint.pointType;
        }
        newPath.closed = originalPath.closed;
        newPath.filled = false;

        if (detachOptions.pathNoStroke) {
            newPath.stroked = false;
            return;
        }
        newPath.stroked = true;

        if (detachOptions.pathBlackStroke) {
            var blackColor = new CMYKColor();
            blackColor.cyan = 0;
            blackColor.magenta = 0;
            blackColor.yellow = 0;
            blackColor.black = 100;

            newPath.strokeColor = blackColor;
            newPath.strokeWidth = 1;
        }
    }

    /**
     * 1つのパス上文字をポイント文字とパスに分け、元のパス上文字を削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} pathText - 元のパス上文字
     * @param {Object} detachOptions - showOptionsDialog() の結果
     * @returns {void}
     */
    function detachPathText(doc, pathText, detachOptions) {
        var originalPath = pathText.textPath;
        var textFormatMode = detachOptions.textFormatMode;

        /* 1. テキスト属性のスナップショットを取得（full: 文字ごとの属性まで保持） */
        /* 1. Snapshot text attributes (full: preserve per-character attributes) */
        var attributeSnapshots = (textFormatMode === "full") ? snapshotCharacterAttributes(pathText) : null;

        /* 段落属性（先頭段落から取得）/ Paragraph attributes (from first paragraph) */
        var justification = null;
        if (pathText.paragraphs.length > 0) {
            justification = pathText.paragraphs[0].paragraphAttributes.justification;
        }

        var textContents = pathText.contents; /* 文字列を保存 / Save string contents */

        /* 2. パスを複製 / Duplicate path */
        if (!detachOptions.pathDelete) {
            duplicateTextPath(doc, originalPath, detachOptions);
        }

        /* 3. 新しいポイントテキストを作成し属性を復元 / Create new point text and restore attributes */
        var newText = doc.textFrames.add();

        /* 位置をパス開始点に合わせる / Align position to path start point */
        var anchorPoint = originalPath.pathPoints[0].anchor;
        newText.position = [anchorPoint[0], anchorPoint[1]];

        if (textFormatMode === "fast") {
            /* fast: textRange を丸ごと複製（内容・書式をまとめて複製）/ fast: duplicate the entire textRange with its formatting */
            pathText.textRange.duplicate(newText);

            /* 文字色を確実に引き継ぐ / Ensure character colors are carried over */
            for (var rangeIndex = 0; rangeIndex < pathText.textRanges.length; rangeIndex++) {
                /* 複製先に同じ番号の範囲が無いことがある / the copy may lack a range with the same index */
                try {
                    var sourceRangeAttributes = pathText.textRanges[rangeIndex].characterAttributes;
                    var targetRangeAttributes = newText.textRanges[rangeIndex].characterAttributes;
                    targetRangeAttributes.fillColor = sourceRangeAttributes.fillColor;
                    targetRangeAttributes.strokeColor = sourceRangeAttributes.strokeColor;
                } catch (eRR) { }
            }
        } else if (textFormatMode === "none") {
            /* none: 内容だけ流し込み / none: insert content only, use default formatting */
            newText.contents = textContents;
        } else {
            /* full: 内容を流し込み、文字ごとの属性を復元 / full: insert content and restore per-character attributes */
            newText.contents = textContents;

            /* 段落属性の復元 / Restore paragraph attributes */
            if (justification !== null && newText.paragraphs.length > 0) {
                newText.paragraphs[0].paragraphAttributes.justification = justification;
            }

            /* 重要：新規テキストの既定の線を先に消す。元に線があるときだけ、後の文字ごとの復元で戻る */
            /* Important: clear the default stroke on the new text first; it comes back only where the original had one */
            try {
                newText.textRange.characterAttributes.strokeColor = new NoColor();
                newText.textRange.characterAttributes.strokeWeight = 0;
            } catch (eClr) { }

            /* 文字属性の復元 / Restore character attributes */
            restoreCharacterAttributes(newText, attributeSnapshots);
        }

        /* 4. 元のパス上文字を削除 / Remove original path text */
        pathText.remove();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで処理方法を選び、選択中のパス上文字をすべて分離する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var detachOptions = showOptionsDialog();
        if (!detachOptions) return; /* キャンセル / Cancelled */

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 選択からパス上文字を抽出 / Extract path text from selection */
        var pathTexts = collectPathTexts(selectedItems);
        if (pathTexts.length === 0) {
            alert(getLabel("alert.noPathText"));
            return;
        }

        /* パス上文字の変換処理 / Convert path text items */
        for (var j = pathTexts.length - 1; j >= 0; j--) {
            detachPathText(doc, pathTexts[j], detachOptions);
        }
    }

    main();

})();
