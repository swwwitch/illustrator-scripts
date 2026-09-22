#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの文字比率とベースラインシフトを調整します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustTextScaleBaseline.md

### Overview

Adjusts the character scale and the baseline shift of the selected text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustTextScaleBaseline.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustTextScaleBaseline";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustTextScaleBaseline.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustTextScaleBaseline.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OFFSET_X = 300;               /* ダイアログを右へずらす量 / Horizontal dialog offset */
    var DIALOG_OPACITY = 0.97;               /* ダイアログの不透明度 / Dialog opacity */
    var COLUMN_SPACING = 20;                 /* 左右の列と左列内の間隔 / Spacing of the columns and inside the left column */
    var ADJUST_PANEL_MARGINS = [15, 10, 15, 10]; /* ［調整］パネルの余白 / Margins of the Adjust panel */
    var TARGET_CHAR_INPUT_CHARACTERS = 10;   /* 対象文字欄の幅（文字数）/ Width of the target character field */
    var NUMBER_INPUT_CHARACTERS = 4;         /* 数値欄の幅（文字数）/ Width of the number fields */
    var APPARENT_SIZE_CHARACTERS = 5;        /* 見かけのサイズ表示の幅（文字数）/ Width of the apparent size display */
    var ROW_LABEL_WIDTH = 120;               /* 行ラベルの幅 / Width of the row labels */
    var BUTTON_WIDTH = 90;                   /* ボタンの幅 / Button width */
    var CANCEL_RESET_GAP = 50;               /* ［キャンセル］と［リセット］の間隔 / Gap between Cancel and Reset */

    // =========================================
    // プレビュー / Preview
    // =========================================

    var PREVIEW_INTERVAL_MS = 80; /* 入力中のプレビューを間引く間隔（ミリ秒）/ Throttle interval for previews while typing (ms) */

    /* ↑↓キーで増減する量 / Arrow key steps */
    var ARROW_STEPS = { step: 1, shiftStep: 10, altStep: 0.1 };
    var SCALE_ARROW_STEPS = { step: 1, shiftStep: 10, altStep: 5 };

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
            title: {
                ja: "フォントサイズとベースライン調整 " + SCRIPT_VERSION,
                en: "Font Size & Baseline Adjuster " + SCRIPT_VERSION
            }
        },
        panel: {
            adjust: { ja: "調整", en: "Adjust" }
        },
        fieldLabel: {
            targetChar: { ja: "対象文字", en: "Target Char" },
            fontSize: { ja: "フォントサイズ", en: "Font Size" },
            scale: { ja: "水平比率/垂直比率", en: "Scale" },
            apparent: { ja: "見かけ", en: "Apparent" },
            baselineShift: { ja: "ベースラインシフト", en: "Baseline Shift" },
            kerning: { ja: "カーニング", en: "Kerning" },
            tracking: { ja: "トラッキング", en: "Tracking" }
        },
        tooltip: {
            targetChar: {
                ja: "ここに書いた文字だけを調整します。空欄にすると選択範囲すべてが対象です。",
                en: "Adjusts only the characters listed here. Leave blank to affect the whole selection."
            },
            fontSize: { ja: "対象文字のフォントサイズを増減します。", en: "Changes the font size of the target characters." },
            scale: {
                ja: "対象文字の長体・平体です。100%で変形なしになります。",
                en: "Horizontal and vertical scale of the target characters. 100% means no distortion."
            },
            apparentSize: {
                ja: "フォントサイズに比率を掛けた、見かけの文字サイズを表示します。比率が100%のときは淡色表示になります。",
                en: "Shows the apparent type size: the font size multiplied by the scale. Dimmed when the scale is 100%."
            },
            baselineShift: {
                ja: "対象文字を上下にずらす量です。負の値で下がります。",
                en: "How far the target characters move up. Negative values move them down."
            },
            kerning: { ja: "対象文字の前後の詰めです。単位は1/1000em。", en: "Spacing around the target characters, in 1/1000 em." },
            tracking: { ja: "選択範囲全体の字間です。単位は1/1000em。", en: "Letter spacing across the whole selection, in 1/1000 em." },
            reset: {
                ja: "対象文字とフォントサイズを開いたときの値に戻し、比率を100%、ベースラインシフト・カーニング・トラッキングを0にします。",
                en: "Restores the target characters and font size to their initial values, sets the scale to 100% and baseline shift, kerning and tracking to 0."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" }
        },
        alert: {
            selectText: { ja: "テキストを選択してください", en: "Please select text" }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "fieldLabel.fontSize" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 対象の収集 / Collecting targets
    // =========================================

    /**
     * 選択から調整対象の TextRange を集める（テキストフレームは全文、文字の選択はその範囲）
     * @param {Array|TextRange} currentSelection - 現在の選択
     * @returns {TextRange[]} 調整対象の TextRange
     */
    function collectSelectedTextRanges(currentSelection) {
        var textRanges = [];
        if (!currentSelection || currentSelection.length === 0) return textRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem instanceof TextFrame) {
                textRanges.push(selectedItem.textRange);
            } else if (selectedItem instanceof TextRange) {
                textRanges.push(selectedItem);
            }
        }
        /* 文字の編集中は選択そのものが TextRange / While editing text, the selection itself is a TextRange */
        if (!(currentSelection instanceof Array) && currentSelection instanceof TextRange) {
            textRanges.push(currentSelection);
        }
        return textRanges;
    }

    /**
     * すべての文字列に共通して含まれる文字を、最初の文字列での出現順に重複なく返す
     * @param {string[]} texts - 文字列の配列
     * @returns {string} 共通の文字
     */
    function getCommonCharacters(texts) {
        if (texts.length === 0) return "";
        var commonChars = texts[0];
        for (var i = 1; i < texts.length; i++) {
            var otherText = texts[i];
            var sharedChars = "";
            for (var j = 0; j < commonChars.length; j++) {
                var currentChar = commonChars.charAt(j);
                if (otherText.indexOf(currentChar) !== -1 && sharedChars.indexOf(currentChar) === -1) {
                    sharedChars += currentChar;
                }
            }
            commonChars = sharedChars;
            if (commonChars.length === 0) break;
        }
        return commonChars;
    }

    /**
     * 英数字を除いた文字を、重複なく出現順に返す
     * @param {string} text - 元の文字列
     * @returns {string} 英数字以外の文字（重複なし）
     */
    function getUniqueNonAlphanumerics(text) {
        var strippedText = text.replace(/[0-9A-Za-z]/g, "");
        var uniqueChars = "";
        for (var i = 0; i < strippedText.length; i++) {
            var currentChar = strippedText.charAt(i);
            if (uniqueChars.indexOf(currentChar) === -1) {
                uniqueChars += currentChar;
            }
        }
        return uniqueChars;
    }

    /**
     * 対象文字の初期値を求める（すべての選択範囲に共通する、英数字以外の文字）
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @returns {string} 対象文字の初期値
     */
    function findDefaultTargetChars(targetRanges) {
        var rangeTexts = [];
        for (var i = 0; i < targetRanges.length; i++) {
            rangeTexts.push(targetRanges[i].contents);
        }
        return getUniqueNonAlphanumerics(getCommonCharacters(rangeTexts));
    }

    // =========================================
    // 文字の調整 / Character adjustment
    // =========================================

    /**
     * 文字が対象文字に当たるかを判定する（対象文字が空ならすべて対象）
     * @param {TextRange} textChar - 1文字の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @returns {boolean} 対象なら true
     */
    function isTargetChar(textChar, charFilter) {
        return !charFilter.hasTargetChars || charFilter.targetChars.indexOf(textChar.contents) !== -1;
    }

    /**
     * 対象文字の入力欄から、対象文字の指定を読み取る
     * @param {EditText} targetCharInput - 対象文字の入力欄
     * @returns {{targetChars: string, hasTargetChars: boolean}} 対象文字の指定
     */
    function readCharFilter(targetCharInput) {
        var targetChars = targetCharInput.text;
        return { targetChars: targetChars, hasTargetChars: targetChars.length > 0 };
    }

    /**
     * すべての TextRange で、対象文字に当たる文字それぞれに処理を行う
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @param {Function} charCallback - 1文字の TextRange を受け取る処理
     * @returns {void}
     */
    function forEachTargetChar(targetRanges, charFilter, charCallback) {
        for (var i = 0; i < targetRanges.length; i++) {
            var rangeChars = targetRanges[i].characters;
            for (var j = 0; j < rangeChars.length; j++) {
                if (isTargetChar(rangeChars[j], charFilter)) charCallback(rangeChars[j]);
            }
        }
    }

    /**
     * 対象文字に当たる最初の文字を返す
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @returns {TextRange|null} 最初の対象文字（無ければ null）
     */
    function findFirstTargetChar(targetRanges, charFilter) {
        for (var i = 0; i < targetRanges.length; i++) {
            var rangeChars = targetRanges[i].characters;
            for (var j = 0; j < rangeChars.length; j++) {
                if (isTargetChar(rangeChars[j], charFilter)) return rangeChars[j];
            }
        }
        return null;
    }

    /**
     * 対象文字にサイズ・比率・ベースラインシフト・カーニング・トラッキングを適用する（null の項目は触らない）
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {{targetChars: string, hasTargetChars: boolean}} charFilter - 対象文字の指定
     * @param {{size: ?number, scale: ?number, baseline: ?number, kerning: ?number, tracking: ?number}} adjustParams - 適用する値
     * @returns {void}
     */
    function applyTextAdjustments(targetRanges, charFilter, adjustParams) {
        if (adjustParams.size !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.size = adjustParams.size;
            });
        }
        if (adjustParams.scale !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.horizontalScale = adjustParams.scale;
                textChar.characterAttributes.verticalScale = adjustParams.scale;
            });
        }
        if (adjustParams.baseline !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                textChar.characterAttributes.baselineShift = adjustParams.baseline;
            });
        }
        if (adjustParams.kerning !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                textChar.kerning = adjustParams.kerning;
            });
        }
        if (adjustParams.tracking !== null) {
            forEachTargetChar(targetRanges, charFilter, function (textChar) {
                textChar.characterAttributes.tracking = adjustParams.tracking;
            });
        }
    }

    /**
     * 見かけの文字サイズ（フォントサイズ×比率）を求める
     * @param {number} fontSize - フォントサイズ
     * @param {number} scalePercent - 比率（%）
     * @returns {number} 見かけの文字サイズ（小数第2位まで）
     */
    function calculateApparentSize(fontSize, scalePercent) {
        return Math.round(fontSize * scalePercent) / 100;
    }

    // =========================================
    // プレビューの取り消し管理 / Preview undo management
    // =========================================

    /**
     * プレビュー時に Undo 履歴を汚さないための小さな管理クラス
     * - addStep(): 変更処理を実行して undoDepth を数える
     * - rollback(): プレビューで行った変更をすべて取り消す
     * - confirm(finalAction): 一度戻してから本番処理を1回だけ実行し、Undo を1回にまとめる
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;

        /**
         * 変更処理を1ステップとして実行し、Undo の深さを数える
         * @param {Function} stepAction - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (stepAction) {
            /* 文字属性の書き込みは DOM が例外を投げうる / Writing character attributes can throw */
            try {
                stepAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert("Preview Error: " + e);
            }
        };

        /**
         * プレビューで行った変更をすべて取り消す
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * プレビューを戻してから本番処理を1回だけ実行する
         * @param {Function} finalAction - 確定時に実行する処理
         * @returns {void}
         */
        this.confirm = function (finalAction) {
            this.rollback();
            finalAction();
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ↑↓キーで数値を増減できるようにする（Shift: 大きく、Option/Alt: 細かく）
     * @param {EditText} editText - 対象の入力欄
     * @param {{step: number, shiftStep: number, altStep: number}} stepOptions - 増減量
     * @returns {void}
     */
    function changeValueByArrowKey(editText, stepOptions) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;
            var keyboardState = ScriptUI.environment.keyboardState;
            var delta = 1;
            var precision = 0;

            if (keyboardState.shiftKey) {
                delta = stepOptions.shiftStep || 10;
            } else if (keyboardState.altKey) {
                delta = stepOptions.altStep || 0.1;
                precision = 1;
            } else {
                delta = stepOptions.step || 1;
            }

            if (event.keyName == "Up") {
                if (keyboardState.shiftKey) {
                    value = Math.floor(value / delta) * delta + delta;
                } else {
                    value += delta;
                }
            } else if (event.keyName == "Down") {
                if (keyboardState.shiftKey) {
                    value = Math.ceil(value / delta) * delta - delta;
                } else {
                    value -= delta;
                }
            } else {
                return;
            }

            if (precision > 0) {
                var factor = Math.pow(10, precision);
                value = Math.round(value * factor) / factor;
            } else {
                value = Math.round(value);
            }

            event.preventDefault();
            editText.text = value;
            editText.notify("onChange");
        });
    }

    /**
     * 右揃えの行ラベル＋数値欄＋単位の1行を追加する
     * @param {Group|Panel} parentGroup - 追加先
     * @param {string} labelPath - 行ラベルのパス
     * @param {string} tooltipPath - 数値欄の tooltip のパス
     * @param {string} initialText - 数値欄の初期値
     * @param {string} unitText - 単位の表示
     * @returns {EditText} 追加した数値欄（行は .parent で取れる）
     */
    function addNumberRow(parentGroup, labelPath, tooltipPath, initialText, unitText) {
        var numberRow = parentGroup.add("group");
        numberRow.orientation = "row";
        var rowLabel = numberRow.add("statictext", undefined, getLabel(labelPath));
        rowLabel.justify = "right";
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
        var numberInput = numberRow.add("edittext", undefined, initialText);
        numberInput.characters = NUMBER_INPUT_CHARACTERS;
        numberInput.helpTip = getLabel(tooltipPath);
        numberRow.add("statictext", undefined, unitText);
        return numberInput;
    }

    /**
     * ベースラインシフト・カーニング・トラッキングの行（値を右揃え）を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - 行ラベルのパス
     * @param {string} tooltipPath - 数値欄の tooltip のパス
     * @param {string} unitText - 単位の表示
     * @returns {EditText} 追加した数値欄
     */
    function addShiftRow(parentPanel, labelPath, tooltipPath, unitText) {
        var shiftInput = addNumberRow(parentPanel, labelPath, tooltipPath, "0", unitText);
        shiftInput.parent.alignChildren = ["right", "center"];
        shiftInput.justify = "right";
        return shiftInput;
    }

    /**
     * ダイアログを組み立てる（イベントはまだ付けない）
     * @param {string} defaultTargetChars - 対象文字の初期値
     * @param {string} unitLabel - 文字サイズの単位
     * @returns {Object} ダイアログと各コントロール
     */
    function buildDialog(defaultTargetChars, unitLabel) {
        var adjustDialog = new Window("dialog", getLabel("dialog.title"));
        adjustDialog.alignChildren = "left";
        adjustDialog.opacity = DIALOG_OPACITY;

        var contentGroup = adjustDialog.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = "top";
        contentGroup.spacing = COLUMN_SPACING;

        var leftColumn = contentGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "left";
        leftColumn.spacing = COLUMN_SPACING;

        var rightColumn = contentGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "right";

        /* 対象文字 / Target characters */
        var targetCharGroup = leftColumn.add("group");
        targetCharGroup.orientation = "row";
        targetCharGroup.add("statictext", undefined, labelText("fieldLabel.targetChar"));
        var targetCharInput = targetCharGroup.add("edittext", undefined, defaultTargetChars);
        targetCharInput.characters = TARGET_CHAR_INPUT_CHARACTERS;
        targetCharInput.helpTip = getLabel("tooltip.targetChar");

        var adjustPanel = leftColumn.add("panel", undefined, getLabel("panel.adjust"));
        adjustPanel.orientation = "column";
        adjustPanel.alignChildren = ["left", "top"];
        adjustPanel.margins = ADJUST_PANEL_MARGINS;

        /* フォントサイズと比率 / Font size and scale */
        var sizeScaleGroup = adjustPanel.add("group");
        sizeScaleGroup.orientation = "column";
        sizeScaleGroup.alignChildren = "left";

        var sizeInput = addNumberRow(sizeScaleGroup, "fieldLabel.fontSize", "tooltip.fontSize", "0", unitLabel);
        sizeInput.parent.margins = [0, 10, 0, 0];
        sizeInput.parent.alignChildren = "left";

        var hScaleInput = addNumberRow(sizeScaleGroup, "fieldLabel.scale", "tooltip.scale", "100", "%");
        hScaleInput.parent.margins = [0, 0, 0, 6];

        /* 見かけのサイズ（表示のみ）/ Apparent size (display only) */
        var apparentGroup = adjustPanel.add("group");
        apparentGroup.orientation = "row";
        var apparentLabel = apparentGroup.add("statictext", undefined, getLabel("fieldLabel.apparent"));
        apparentLabel.justify = "right";
        apparentLabel.preferredSize.width = ROW_LABEL_WIDTH;
        var apparentSizeText = apparentGroup.add("statictext", undefined, "--");
        apparentSizeText.characters = APPARENT_SIZE_CHARACTERS;
        apparentSizeText.helpTip = getLabel("tooltip.apparentSize");
        var apparentUnitLabel = apparentGroup.add("statictext", undefined, unitLabel);

        var baselineInput = addShiftRow(adjustPanel, "fieldLabel.baselineShift", "tooltip.baselineShift", unitLabel);
        var kerningInput = addShiftRow(adjustPanel, "fieldLabel.kerning", "tooltip.kerning", "/1000");
        var trackingInput = addShiftRow(adjustPanel, "fieldLabel.tracking", "tooltip.tracking", "/1000");

        /* ボタン（右列）/ Buttons (right column) */
        var btnColumnGroup = rightColumn.add("group");
        btnColumnGroup.alignment = "right";
        btnColumnGroup.orientation = "column";
        var btnOK = btnColumnGroup.add("button", undefined, getLabel("button.ok"));
        var btnCancel = btnColumnGroup.add("button", undefined, getLabel("button.cancel"));
        var cancelResetSpacer = btnColumnGroup.add("statictext", undefined, "");
        cancelResetSpacer.preferredSize.height = CANCEL_RESET_GAP;
        var btnReset = btnColumnGroup.add("button", undefined, getLabel("button.reset"));
        btnReset.helpTip = getLabel("tooltip.reset");
        btnReset.preferredSize.width = BUTTON_WIDTH;
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;

        return {
            dialog: adjustDialog,
            targetCharInput: targetCharInput,
            sizeInput: sizeInput,
            hScaleInput: hScaleInput,
            apparentLabel: apparentLabel,
            apparentSizeText: apparentSizeText,
            apparentUnitLabel: apparentUnitLabel,
            baselineInput: baselineInput,
            kerningInput: kerningInput,
            trackingInput: trackingInput,
            btnOK: btnOK,
            btnCancel: btnCancel,
            btnReset: btnReset
        };
    }

    /**
     * 入力欄の数値をまとめて読み取る（数値でない欄は null）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {{size: ?number, scale: ?number, baseline: ?number, kerning: ?number, tracking: ?number}} 入力値
     */
    function readAdjustParams(dialogControls) {
        var sizeValue = parseFloat(dialogControls.sizeInput.text);
        var scaleValue = parseFloat(dialogControls.hScaleInput.text);
        var baselineValue = parseFloat(dialogControls.baselineInput.text);
        var kerningValue = parseFloat(dialogControls.kerningInput.text);
        var trackingValue = parseFloat(dialogControls.trackingInput.text);
        return {
            size: isNaN(sizeValue) ? null : sizeValue,
            scale: isNaN(scaleValue) ? null : scaleValue,
            baseline: isNaN(baselineValue) ? null : baselineValue,
            kerning: isNaN(kerningValue) ? null : kerningValue,
            tracking: isNaN(trackingValue) ? null : trackingValue
        };
    }

    /**
     * 見かけのサイズの表示を更新する（比率 100% のときは淡色表示）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {void}
     */
    function updateApparentSizeDisplay(dialogControls) {
        var fontSize = parseFloat(dialogControls.sizeInput.text);
        var scalePercent = parseFloat(dialogControls.hScaleInput.text);
        if (isNaN(fontSize) || isNaN(scalePercent)) {
            dialogControls.apparentSizeText.text = "--";
        } else {
            dialogControls.apparentSizeText.text = calculateApparentSize(fontSize, scalePercent) + "";
        }

        var isDimmed = (scalePercent === 100);
        dialogControls.apparentLabel.enabled = !isDimmed;
        dialogControls.apparentSizeText.enabled = !isDimmed;
        dialogControls.apparentUnitLabel.enabled = !isDimmed;
    }

    /**
     * ダイアログにイベントを付け、初期値の読み込みと最初のプレビューを行う
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {TextRange[]} targetRanges - 調整対象の TextRange
     * @param {string} defaultTargetChars - 対象文字の初期値
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, targetRanges, defaultTargetChars) {
        var adjustDialog = dialogControls.dialog;
        var targetCharInput = dialogControls.targetCharInput;
        var sizeInput = dialogControls.sizeInput;
        var hScaleInput = dialogControls.hScaleInput;
        var previewManager = new PreviewManager();
        var initialFontSize = null; /* 最初に読み込んだフォントサイズ（リセット用）/ First font size read, for Reset */
        var lastPreviewTime = 0;

        /**
         * 入力欄の値を対象文字に適用する
         * @returns {void}
         */
        function applyCurrentValues() {
            applyTextAdjustments(targetRanges, readCharFilter(targetCharInput), readAdjustParams(dialogControls));
        }

        /**
         * プレビューを取り消してから掛け直す
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(applyCurrentValues);
        }

        /**
         * プレビューを間引いて更新する（onChanging の連打で Undo と再適用が過剰にならないように）
         * @returns {void}
         */
        function updatePreviewThrottled() {
            var now = (new Date()).getTime();
            if (now - lastPreviewTime < PREVIEW_INTERVAL_MS) return;
            lastPreviewTime = now;
            updatePreview();
        }

        /**
         * 最初の対象文字のサイズと比率を入力欄に読み込む
         * @returns {void}
         */
        function loadFirstTargetCharValues() {
            var firstChar = findFirstTargetChar(targetRanges, readCharFilter(targetCharInput));
            if (!firstChar) {
                sizeInput.text = "";
                hScaleInput.text = "";
                dialogControls.apparentSizeText.text = "--";
                return;
            }
            var fontSize = firstChar.size;
            if (initialFontSize === null) {
                initialFontSize = fontSize;
            }
            sizeInput.text = (Math.round(fontSize * 10) / 10) + "";
            hScaleInput.text = (Math.round(firstChar.characterAttributes.horizontalScale * 10) / 10) + "";
            updateApparentSizeDisplay(dialogControls);
        }

        /**
         * サイズ・比率の確定を受けて、プレビューしてから値を読み直す
         * @returns {void}
         */
        function onSizeOrScaleChange() {
            updatePreview();
            loadFirstTargetCharValues();
            updateApparentSizeDisplay(dialogControls);
        }

        /* ベースラインシフト・カーニング・トラッキング / Baseline shift, kerning and tracking */
        var shiftInputs = [dialogControls.baselineInput, dialogControls.kerningInput, dialogControls.trackingInput];
        for (var i = 0; i < shiftInputs.length; i++) {
            shiftInputs[i].onChange = updatePreview;
            shiftInputs[i].onChanging = updatePreviewThrottled;
            changeValueByArrowKey(shiftInputs[i], ARROW_STEPS);
        }
        /* 既存の動作：ベースラインシフトだけ↑↓が2重に登録されている（1回の押下で2段進む）
           Existing behavior: baseline shift has the arrow handler twice, so one keypress moves two steps */
        changeValueByArrowKey(dialogControls.baselineInput, ARROW_STEPS);

        sizeInput.onChange = onSizeOrScaleChange;
        sizeInput.onChanging = function () { updateApparentSizeDisplay(dialogControls); };
        changeValueByArrowKey(sizeInput, ARROW_STEPS);

        hScaleInput.onChange = onSizeOrScaleChange;
        hScaleInput.onChanging = function () { updateApparentSizeDisplay(dialogControls); };
        changeValueByArrowKey(hScaleInput, SCALE_ARROW_STEPS);

        targetCharInput.onChange = function () {
            loadFirstTargetCharValues();
            updatePreview();
        };

        dialogControls.btnReset.onClick = function () {
            targetCharInput.text = defaultTargetChars;
            if (initialFontSize !== null) {
                sizeInput.text = initialFontSize;
            }
            hScaleInput.text = "100";
            dialogControls.baselineInput.text = "0";
            dialogControls.kerningInput.text = "0";
            dialogControls.trackingInput.text = "0";

            updatePreview();
            loadFirstTargetCharValues();
        };

        dialogControls.btnCancel.onClick = function () {
            previewManager.rollback();
            adjustDialog.close(2);
        };

        dialogControls.btnOK.onClick = function () {
            /* Undo を1回にまとめて確定 / Confirm as a single undo step */
            previewManager.confirm(applyCurrentValues);
            targetCharInput.onChange = null;
            adjustDialog.close();
        };

        adjustDialog.onShow = function () {
            var dialogLocation = adjustDialog.location;
            adjustDialog.location = [dialogLocation[0] + DIALOG_OFFSET_X, dialogLocation[1]];
            hScaleInput.active = true;
        };

        loadFirstTargetCharValues();
        updateApparentSizeDisplay(dialogControls);

        /* 初期状態も管理下で1回プレビュー適用（この時点で undoDepth=1 になる）
           Apply the initial state once under management (undoDepth becomes 1 here) */
        updatePreview();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のテキストを集めてダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length <= 0) {
            return;
        }

        var targetRanges = collectSelectedTextRanges(app.activeDocument.selection);
        if (targetRanges.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var defaultTargetChars = findDefaultTargetChars(targetRanges);
        var dialogControls = buildDialog(defaultTargetChars, getUnitInfo("text/units").label);
        bindDialogEvents(dialogControls, targetRanges, defaultTargetChars);
        dialogControls.dialog.show();
    }

    main();

})();
