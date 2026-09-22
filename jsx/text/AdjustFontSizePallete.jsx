#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択している文字を対象に、フォントサイズと水平比率／垂直比率を調整する常駐パレットです。
ライブプレビューで結果を確認しながら調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSizePallete.md

note記事も参照してください。
https://note.com/dtp_tranist/n/xxxxxxxx

### Overview

A persistent palette for adjusting the font size and the horizontal and vertical scale of the selected characters.
A live preview shows the result as you work.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSizePallete.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustFontSizePallete";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSizePallete.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSizePallete.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DIALOG_OPACITY  = 0.98;  /* ダイアログ透明度 / dialog opacity */
    var DIALOG_OFFSET_X = 0;     /* 表示位置の横オフセット / horizontal offset on show */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [16, 20, 16, 12];  /* パネル余白 / panel margins */
    var PANEL_SPACING = 8;                 /* パネル内の標準間隔 / default spacing inside panels */
    var FIELD_SPACING = 6;                 /* 入力行どうしの間隔 / spacing between field rows */
    var LABEL_WIDTH   = 118;               /* ラベル幅（揃える）/ unified label width */
    var BUTTON_WIDTH  = 90;                /* OK・キャンセルの幅 / width of OK and Cancel */
    var CONVERT_BUTTON_WIDTH = 150;        /* 実サイズ↔見かけボタンの幅 / width of the actual↔apparent button */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールからUIの表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "フォントサイズの調整", en: "Font Size Adjuster" }
        },
        panel: {
            fontSize: { ja: "フォントサイズの調整", en: "Font Size Adjustment" }
        },
        fieldLabel: {
            fontSize: { ja: "フォントサイズ", en: "Font Size" },
            scale: { ja: "水平比率/垂直比率", en: "Scale" },
            apparent: { ja: "見かけ", en: "Apparent" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" },
            toApparent: { ja: "実サイズ↔見かけ", en: "Actual ↔ Apparent" }
        },
        alert: {
            selectText: { ja: "テキストを選択してください", en: "Please select text" },
            previewError: { ja: "プレビュー更新エラー / Preview update error: ", en: "Preview update error: " }
        },
        tooltip: {
            scale: { ja: "水平比率・垂直比率を同じ値でまとめて設定します。", en: "Sets the horizontal and vertical scale together to the same value." },
            apparent: { ja: "フォントサイズ×比率で計算した、実際の見た目のサイズです。", en: "The actual visual size, computed as font size × scale." },
            toApparent: {
                ja: "サイズ×比率を見かけサイズとして実フォントサイズに焼き込み、比率を100%にします。もう一度押すと焼き込み前の比率付き状態に戻ります",
                en: "Bakes size × scale into the actual font size at 100%. Press again to restore the previous scaled state."
            },
            reset: {
                ja: "選択している文字すべてを先頭文字のフォントサイズに統一し、水平比率・垂直比率を100%に揃えて適用します。",
                en: "Unifies every selected character to the first character's font size and sets the horizontal / vertical scale to 100%, then applies."
            }
        }
    };

    /**
     * キーからラベルを現在の言語で取得する（"panel.fontSize" のようにドット区切り）
     * @param {string} labelPath - カテゴリ名とキー名をドットでつないだラベルキー
     * @returns {string} 現在の言語のラベル文字列（未定義の場合は英語にフォールバック）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) {
                return labelPath;
            }
        }
        return labelNode[uiLang] || labelNode.en;
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

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // UI部品 / UI helpers
    // =========================================

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は既定値）
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
     * グループに共通のレイアウト設定を適用する
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は既定値）
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

    /**
     * ラベル＋入力欄＋単位の行を追加する
     * @param {Panel|Group} parentContainer - 追加先のコンテナ
     * @param {string} labelKey - ラベルキー
     * @param {string} defaultValue - 入力欄の初期値
     * @param {string} unitText - 単位表示の文字列
     * @returns {{row: Group, label: StaticText, input: EditText, unit: StaticText}} 生成した行の各コントロール
     */
    function addFieldRow(parentContainer, labelKey, defaultValue, unitText) {
        var fieldRowGroup = parentContainer.add("group");
        setupGroup(fieldRowGroup, "row");
        var fieldLabelText = fieldRowGroup.add("statictext", undefined, labelText(labelKey));
        fieldLabelText.justify = "right";
        var valueInput = fieldRowGroup.add("edittext", undefined, defaultValue);
        valueInput.characters = 4;
        valueInput.justify = "right";
        var unitLabelText = fieldRowGroup.add("statictext", undefined, unitText);
        return { row: fieldRowGroup, label: fieldLabelText, input: valueInput, unit: unitLabelText };
    }

    /**
     * ラベル＋表示専用テキスト＋単位の行を追加する
     * @param {Panel|Group} parentContainer - 追加先のコンテナ
     * @param {string} labelKey - ラベルキー
     * @param {string} unitText - 単位表示の文字列
     * @returns {{row: Group, label: StaticText, value: StaticText, unit: StaticText}} 生成した行の各コントロール
     */
    function addReadoutRow(parentContainer, labelKey, unitText) {
        var readoutRowGroup = parentContainer.add("group");
        setupGroup(readoutRowGroup, "row");
        var fieldLabelText = readoutRowGroup.add("statictext", undefined, labelText(labelKey));
        fieldLabelText.justify = "right";
        var valueText = readoutRowGroup.add("statictext", undefined, "--");
        valueText.characters = 5;
        var unitLabelText = readoutRowGroup.add("statictext", undefined, unitText);
        return { row: readoutRowGroup, label: fieldLabelText, value: valueText, unit: unitLabelText };
    }

    /**
     * 行（ラベル＋入力／表示＋単位）にまとめてヘルプチップを設定する
     * @param {object} fieldRow - addFieldRow / addReadoutRow が返した行オブジェクト
     * @param {string} tooltipText - 設定するヘルプチップ文字列
     * @returns {void}
     */
    function setFieldTooltip(fieldRow, tooltipText) {
        fieldRow.label.helpTip = tooltipText;
        fieldRow.unit.helpTip = tooltipText;
        if (fieldRow.input) fieldRow.input.helpTip = tooltipText;
        if (fieldRow.value) fieldRow.value.helpTip = tooltipText;
    }

    /**
     * 複数ラベルの幅を揃える
     * @param {number} labelWidth - 設定する幅（px）
     * @param {StaticText[]} labelControls - 幅を揃えるラベルの配列
     * @returns {void}
     */
    function alignLabelWidths(labelWidth, labelControls) {
        for (var i = 0; i < labelControls.length; i++) {
            labelControls[i].preferredSize.width = labelWidth;
        }
    }

    /**
     * 入力欄で上下キーによる数値の増減を有効にする
     * @param {EditText} editText - 対象の入力欄
     * @param {{step: number, shiftStep: number, altStep: number}} stepOptions - 通常／shift／option時の増減幅
     * @returns {void}
     */
    function changeValueByArrowKey(editText, stepOptions) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;
            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;
            var precision = 0;

            if (keyboard.shiftKey) {
                delta = stepOptions.shiftStep || 10;
            } else if (keyboard.altKey) {
                delta = stepOptions.altStep || 0.1;
                precision = 1;
            } else {
                delta = stepOptions.step || 1;
            }

            if (event.keyName == "Up") {
                if (keyboard.shiftKey) {
                    value = Math.floor(value / delta) * delta + delta;
                } else {
                    value += delta;
                }
            } else if (event.keyName == "Down") {
                if (keyboard.shiftKey) {
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

    // =========================================
    // 選択取得 / Selection
    // =========================================

    /**
     * 選択中のテキスト範囲を取得する
     * @returns {TextRange[]} 選択されているテキスト範囲の配列（なければ空配列）
     */
    function getTextSelection() {
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        var textRanges = [];
        if (!currentSelection) return textRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (currentSelection.constructor.name === "TextRange") {
            textRanges.push(currentSelection);
            return textRanges;
        }
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem.constructor.name === "TextFrame") {
                textRanges.push(selectedItem.textRange);
            } else if (selectedItem.constructor.name === "TextRange") {
                textRanges.push(selectedItem);
            }
        }
        return textRanges;
    }

    // =========================================
    // 文字の読み取り・適用 / Read & apply character values
    // =========================================

    /**
     * 選択している文字の先頭を取得する
     * @param {TextRange[]} targetRanges - 選択しているテキスト範囲
     * @returns {TextRange|null} 最初の文字（選択している文字がなければ null）
     */
    function findFirstChar(targetRanges) {
        for (var i = 0; i < targetRanges.length; i++) {
            var rangeCharacters = targetRanges[i].characters;
            if (rangeCharacters.length > 0) return rangeCharacters[0];
        }
        return null;
    }

    /**
     * 選択している文字すべてにコールバックを適用する
     * @param {TextRange[]} targetRanges - 選択しているテキスト範囲
     * @param {function} characterAction - 各文字に対して実行する処理
     * @returns {void}
     */
    function forEachSelectedChar(targetRanges, characterAction) {
        for (var i = 0; i < targetRanges.length; i++) {
            var rangeCharacters = targetRanges[i].characters;
            for (var j = 0; j < rangeCharacters.length; j++) {
                characterAction(rangeCharacters[j]);
            }
        }
    }

    /**
     * 選択している文字にフォントサイズと比率を適用する（null の値は変えない）
     * @param {TextRange[]} targetRanges - 選択しているテキスト範囲
     * @param {number|null} sizeInPt - フォントサイズ（pt）
     * @param {number|null} scale - 水平比率・垂直比率（%）
     * @returns {void}
     */
    function applySizeAndScale(targetRanges, sizeInPt, scale) {
        if (sizeInPt !== null) {
            forEachSelectedChar(targetRanges, function (character) {
                character.size = sizeInPt;
            });
        }
        if (scale !== null) {
            forEachSelectedChar(targetRanges, function (character) {
                character.characterAttributes.horizontalScale = scale;
                character.characterAttributes.verticalScale = scale;
            });
        }
    }

    /**
     * フォントサイズと比率から見かけのサイズを求める
     * @param {number} size - フォントサイズ
     * @param {number} scale - 比率（%）
     * @returns {number} 見かけのサイズ（小数第2位まで）
     */
    function calculateApparentSize(size, scale) {
        return Math.round(size * scale) / 100;
    }

    /**
     * 入力欄の数値を読む
     * @param {EditText} editText - 対象の入力欄
     * @returns {number|null} 数値（空欄や数値でないときは null）
     */
    function readFieldNumber(editText) {
        var fieldValue = parseFloat(editText.text);
        return isNaN(fieldValue) ? null : fieldValue;
    }

    // =========================================
    // エラー処理 / Error reporting
    // =========================================

    /* 直近に表示したエラーメッセージ（連続する同一エラーの抑制用）/ last shown error message (to suppress repeats) */
    var lastReportedError = "";

    /**
     * 同一メッセージの連続表示を抑制してアラートを出す（プレビュー連打によるアラート氾濫を防ぐ）
     * @param {string} messagePrefix - メッセージの先頭に付ける説明
     * @param {Error} caughtError - 捕捉した例外
     * @returns {void}
     */
    function reportError(messagePrefix, caughtError) {
        var errorMessage = messagePrefix + String(caughtError);
        if (errorMessage === lastReportedError) return;
        lastReportedError = errorMessage;
        alert(errorMessage);
    }

    /**
     * 正常に処理できたときにエラー抑制状態をリセットする
     * @returns {void}
     */
    function clearReportedError() {
        lastReportedError = "";
    }

    // =========================================
    // PreviewManager
    // プレビュー時にUndo履歴を汚さないための小さな管理クラス。
    // - addStep(): 変更処理を実行してundoDepthをカウント
    // - rollback(): 適用済みのプレビューをすべて取り消し
    // - confirm(finalAction): OK時にプレビューを一度戻し、本適用を1回だけ実行して確定する
    // undoDepth = 適用してまだ取り消していないステップ数 / steps applied and not yet undone
    // =========================================

    /**
     * プレビュー適用とUndoの深さを管理するクラス
     * @returns {void}
     */
    function PreviewManager() {
        this.undoDepth = 0;

        /**
         * 変更処理を1ステップとして実行し、Undoの深さを数える
         * @param {function} stepAction - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (stepAction) {
            /* 文字属性の書き込みは DOM が例外を投げうる / Writing character attributes can throw */
            try {
                stepAction();
                clearReportedError();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                reportError(getLabel("alert.previewError"), e);
            }
        };

        /**
         * 適用済みのプレビューをすべて取り消す
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
         * プレビューを戻してから本適用を1回実行し、確定する
         * @param {function} finalAction - 確定時に実行する適用処理
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
     * ダイアログを組み立てる（イベントは main() で結び付ける）
     * @param {string} unitLabel - フォントサイズの単位表示
     * @returns {object} ダイアログと各コントロール（fontSizeDialog / sizeInput / scaleInput / apparentRow / convertButton / btnReset / btnOK / btnCancel）
     */
    function buildDialog(unitLabel) {
        var fontSizeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        fontSizeDialog.alignChildren = "fill";
        fontSizeDialog.opacity = DIALOG_OPACITY;

        /* フォントサイズの調整パネル / font-size adjustment panel */
        var fontSizePanel = fontSizeDialog.add("panel", undefined, getLabel("panel.fontSize"));
        setupPanel(fontSizePanel, FIELD_SPACING);

        var sizeField = addFieldRow(fontSizePanel, "fieldLabel.fontSize", "0", unitLabel);

        var scaleField = addFieldRow(fontSizePanel, "fieldLabel.scale", "100", "%");
        setFieldTooltip(scaleField, getLabel("tooltip.scale"));

        var apparentRow = addReadoutRow(fontSizePanel, "fieldLabel.apparent", unitLabel);
        setFieldTooltip(apparentRow, getLabel("tooltip.apparent"));

        /* 実サイズ↔見かけのトグルボタン / toggle between actual size and apparent (baked) size */
        var convertButton = fontSizePanel.add("button", undefined, getLabel("button.toApparent"));
        convertButton.helpTip = getLabel("tooltip.toApparent");
        convertButton.alignment = "right";
        convertButton.preferredSize.width = CONVERT_BUTTON_WIDTH;

        alignLabelWidths(LABEL_WIDTH, [sizeField.label, scaleField.label, apparentRow.label]);

        /* ボタン（下部・3カラム: 左=リセット / 中央=スペーサー / 右=OK・キャンセル）/ buttons (bottom, 3 columns) */
        var btnRowGroup = fontSizeDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "fill";
        btnRowGroup.alignChildren = ["fill", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignment = ["left", "center"];
        btnLeftGroup.spacing = FIELD_SPACING;
        var btnReset = btnLeftGroup.add("button", undefined, getLabel("button.reset"));
        btnReset.helpTip = getLabel("tooltip.reset");

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "center"];

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnOK.preferredSize.width = BUTTON_WIDTH;
        btnCancel.preferredSize.width = BUTTON_WIDTH;

        return {
            fontSizeDialog: fontSizeDialog,
            sizeInput: sizeField.input,
            scaleInput: scaleField.input,
            apparentRow: apparentRow,
            convertButton: convertButton,
            btnReset: btnReset,
            btnOK: btnOK,
            btnCancel: btnCancel
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択している文字のフォントサイズと比率を調整するダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length <= 0) {
            return;
        }

        var targetRanges = getTextSelection();
        if (targetRanges.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var previewManager = new PreviewManager();
        var textUnit = getUnitInfo("text/units");
        var unitFactor = textUnit.pointsPerUnit;
        var dialogControls = buildDialog(textUnit.label);
        var sizeInput = dialogControls.sizeInput;
        var scaleInput = dialogControls.scaleInput;
        var apparentRow = dialogControls.apparentRow;

        /* 焼き込み前の状態（順方向で保存→逆方向で復元）。手動でサイズ/比率を変えたら無効化
           pre-bake state (saved on forward, restored on back); cleared when size/scale is edited by hand */
        var apparentToggleState = null;

        /**
         * ポイント値を文字の単位に換算する（小数第1位まで）
         * @param {number} sizeInPt - ポイント値
         * @returns {number} 文字の単位での値
         */
        function roundedSizeInUnit(sizeInPt) {
            return Math.round(sizeInPt / unitFactor * 10) / 10;
        }

        /**
         * 現在のUIの値を選択している文字にまとめて適用する
         * @returns {void}
         */
        function applyAllCurrentValues() {
            var sizeValue = readFieldNumber(sizeInput);
            applySizeAndScale(targetRanges, (sizeValue !== null) ? sizeValue * unitFactor : null, readFieldNumber(scaleInput));
        }

        /**
         * Undo履歴を汚さずにプレビューを更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(applyAllCurrentValues);
        }

        /**
         * 見かけサイズの表示を更新する（比率100%のときはディム表示）
         * @returns {void}
         */
        function updateApparentSizeDisplay() {
            var sizeValue = readFieldNumber(sizeInput);
            var scaleValue = readFieldNumber(scaleInput);
            if (sizeValue === null || scaleValue === null) {
                apparentRow.value.text = "--";
            } else {
                apparentRow.value.text = calculateApparentSize(sizeValue, scaleValue) + "";
            }

            var isDimmed = (scaleValue === 100);
            apparentRow.label.enabled = !isDimmed;
            apparentRow.value.enabled = !isDimmed;
            apparentRow.unit.enabled = !isDimmed;
        }

        /**
         * 選択している文字の先頭の現在値を読み取って入力欄に反映する
         * @returns {void}
         */
        function loadValuesFromSelection() {
            /* 実際の値を読み直すので、焼き込み前の保存状態（トグル）は破棄する
               Reloading the actual values invalidates the saved pre-bake (toggle) state */
            apparentToggleState = null;
            var firstChar = findFirstChar(targetRanges);
            if (!firstChar) {
                sizeInput.text = "";
                scaleInput.text = "";
                apparentRow.value.text = "--";
                return;
            }
            var horizontalScale = Math.round(firstChar.characterAttributes.horizontalScale * 10) / 10;
            sizeInput.text = roundedSizeInUnit(firstChar.size) + "";
            scaleInput.text = horizontalScale + "";
            updateApparentSizeDisplay();
        }

        /**
         * 選択している文字すべてを先頭文字のフォントサイズに統一し、比率を100%に揃えて適用する
         * @returns {void}
         */
        function resetToUniformSize() {
            apparentToggleState = null;
            var firstChar = findFirstChar(targetRanges);
            if (firstChar) {
                sizeInput.text = roundedSizeInUnit(firstChar.size);
            }
            scaleInput.text = "100";
            updatePreview();
            updateApparentSizeDisplay();
        }

        /**
         * 実サイズと見かけサイズを切り替える
         * 順方向：サイズ×比率を実サイズに焼き込み比率100%へ。逆方向：直前の比率付き状態へ戻す
         * @returns {void}
         */
        function toggleApparentSize() {
            var sizeValue = readFieldNumber(sizeInput);
            var scaleValue = readFieldNumber(scaleInput);
            if (sizeValue === null || scaleValue === null) return;
            if (apparentToggleState !== null) {
                /* 見かけ→実サイズ：焼き込み前の比率付き状態に戻す / restore the pre-bake scaled state */
                sizeInput.text = apparentToggleState.size + "";
                scaleInput.text = apparentToggleState.scale + "";
                apparentToggleState = null;
            } else {
                /* 実サイズ→見かけ：比率をサイズへ焼き込み100%に / bake scale into size and reset to 100% */
                apparentToggleState = { size: sizeValue, scale: scaleValue };
                sizeInput.text = calculateApparentSize(sizeValue, scaleValue) + "";
                scaleInput.text = "100";
            }
            updatePreview();
            updateApparentSizeDisplay();
        }

        /**
         * サイズ・比率を手で編集したときの処理（入力値をそのまま適用する）
         * loadValuesFromSelection() で入力欄を読み直すと入力値が丸めで戻る恐れがあるため呼ばない（見かけ表示だけ更新する）
         * @returns {void}
         */
        function handleValueEdited() {
            apparentToggleState = null; /* 手動編集でトグル復元を無効化 / manual edit invalidates the toggle */
            updatePreview();
            updateApparentSizeDisplay();
        }

        /* 入力欄のイベント / wire input events */
        sizeInput.onChange = handleValueEdited;
        sizeInput.onChanging = updateApparentSizeDisplay;
        changeValueByArrowKey(sizeInput, { step: 1, shiftStep: 10, altStep: 0.1 });

        scaleInput.onChange = handleValueEdited;
        scaleInput.onChanging = updateApparentSizeDisplay;
        changeValueByArrowKey(scaleInput, { step: 1, shiftStep: 10, altStep: 5 });

        dialogControls.convertButton.onClick = toggleApparentSize;
        dialogControls.btnReset.onClick = resetToUniformSize;

        var fontSizeDialog = dialogControls.fontSizeDialog;
        dialogControls.btnOK.onClick = function () {
            /* プレビューを1回の適用として確定 / commit the preview as a single apply */
            previewManager.confirm(applyAllCurrentValues);
            fontSizeDialog.close();
        };
        dialogControls.btnCancel.onClick = function () {
            /* 開いてから適用した分をすべて取り消してから閉じる / undo everything applied since open, then close */
            previewManager.rollback();
            fontSizeDialog.close(2);
        };

        /* 初期表示（loadValuesFromSelection が先頭文字の実サイズを読み取って各欄を設定）/ initial state (loadValuesFromSelection reads the actual size of the first char) */
        loadValuesFromSelection();
        updateApparentSizeDisplay();

        fontSizeDialog.onShow = function () {
            fontSizeDialog.location = [fontSizeDialog.location[0] + DIALOG_OFFSET_X, fontSizeDialog.location[1]];
            scaleInput.active = true;
        };

        /* 開いた時点では何も適用しない（現在の状態をそのまま保持）。値を変更したときだけプレビュー適用
           apply nothing on open (keep the current state as-is); preview only kicks in once a value changes */
        fontSizeDialog.show();
    }

    main();

})();
