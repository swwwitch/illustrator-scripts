#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択している文字を対象に、フォントサイズと水平比率／垂直比率を調整します。
ライブプレビューで結果を確認しながら調整でき、キャンセルすると開く前の状態に戻ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSize.md

note記事も参照してください。
https://note.com/dtp_tranist/n/xxxxxxxx

### Overview

Adjusts the font size and the horizontal and vertical scale of the selected characters.
A live preview shows the result, and cancelling restores the state from before the dialog opened.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSize.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustFontSize";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSize.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSize.md"; /* README (English) */
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
    var EDIT_CHARACTERS    = 4;            /* 入力欄の文字数 / width of an edittext in characters */
    var READOUT_CHARACTERS = 5;            /* 表示専用テキストの文字数 / width of a readout in characters */

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
                ja: "調整を取り消して、開いた直後の状態に戻します。optionキーを押しながらクリックすると、選択している文字すべてを先頭文字のフォントサイズに統一し、水平比率・垂直比率を100%に揃えて適用します。",
                en: "Discards the adjustments and returns to the just-opened state. Option-click to unify every selected character to the first character's font size and set the horizontal / vertical scale to 100%."
            }
        }
    };

    /**
     * キーからラベルを現在の言語で取得する（"panel.fontSize" のようにドット区切り）
     * @param {string} key - カテゴリ名とキー名をドットでつないだラベルキー
     * @returns {string} 現在の言語のラベル文字列（未定義の場合は英語にフォールバック）
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelEntry = LABELS[keyParts[0]][keyParts[1]];
        return labelEntry[uiLang] || labelEntry.en;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} key - ラベルキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ":");
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
    // 数値ユーティリティ / Number helpers
    // =========================================

    /**
     * コントロールの文字列を数値として読み取る
     * @param {EditText|StaticText} control - 対象のコントロール
     * @returns {number|null} 数値（空欄や数値でない場合は null）
     */
    function readNumber(control) {
        var parsedValue = parseFloat(control.text);
        return isNaN(parsedValue) ? null : parsedValue;
    }

    /**
     * 小数第1位に丸める
     * @param {number} value - 丸める値
     * @returns {number} 小数第1位までの値
     */
    function roundToTenth(value) {
        return Math.round(value * 10) / 10;
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

    // =========================================
    // レイアウト補助 / Layout helpers
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
     * ラベル＋コントロール＋単位の行を追加する
     * @param {Panel|Group} parent - 追加先のコンテナ
     * @param {string} labelKey - ラベルキー
     * @param {string} controlType - "edittext"（入力欄）または "statictext"（表示専用）
     * @param {string} initialText - コントロールの初期表示文字列
     * @param {string} unitText - 単位表示の文字列
     * @returns {{label: StaticText, control: EditText|StaticText, unit: StaticText}} 生成した行の各コントロール
     */
    function addRow(parent, labelKey, controlType, initialText, unitText) {
        var rowGroup = parent.add("group");
        setupGroup(rowGroup, "row");
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelKey));
        rowLabel.justify = "right";
        var isEditable = (controlType === "edittext");
        var valueControl = rowGroup.add(controlType, undefined, initialText);
        valueControl.characters = isEditable ? EDIT_CHARACTERS : READOUT_CHARACTERS;
        if (isEditable) valueControl.justify = "right";
        var unitLabel = rowGroup.add("statictext", undefined, unitText);
        return { label: rowLabel, control: valueControl, unit: unitLabel };
    }

    /**
     * 行（ラベル＋コントロール＋単位）にまとめてヘルプチップを設定する
     * @param {{label: StaticText, control: EditText|StaticText, unit: StaticText}} rowControls - addRow が返した行オブジェクト
     * @param {string} tooltipText - 設定するヘルプチップ文字列
     * @returns {void}
     */
    function setRowTooltip(rowControls, tooltipText) {
        rowControls.label.helpTip = tooltipText;
        rowControls.control.helpTip = tooltipText;
        rowControls.unit.helpTip = tooltipText;
    }

    /**
     * 行のラベル・コントロール・単位をまとめて有効／無効にする
     * @param {{label: StaticText, control: EditText|StaticText, unit: StaticText}} rowControls - addRow が返した行オブジェクト
     * @param {boolean} enabled - 有効にするなら true
     * @returns {void}
     */
    function setRowEnabled(rowControls, enabled) {
        rowControls.label.enabled = enabled;
        rowControls.control.enabled = enabled;
        rowControls.unit.enabled = enabled;
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
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            var isFine = (!keyboard.shiftKey && keyboard.altKey);
            var delta;
            if (keyboard.shiftKey) {
                /* shiftキーは刻み幅の倍数に丸めてから増減 / snap to a multiple of the step, then move */
                delta = stepOptions.shiftStep;
                value = isUp ? Math.floor(value / delta) * delta : Math.ceil(value / delta) * delta;
            } else {
                delta = isFine ? stepOptions.altStep : stepOptions.step;
            }
            value = isUp ? value + delta : value - delta;
            /* optionキーは小数第1位まで、それ以外は整数に丸める / option keeps one decimal, otherwise round to an integer */
            value = isFine ? roundToTenth(value) : Math.round(value);

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
        var docSelection = app.activeDocument.selection;
        var textRanges = [];
        if (!docSelection) return textRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (docSelection.constructor.name === "TextRange") {
            textRanges.push(docSelection);
            return textRanges;
        }
        for (var i = 0; i < docSelection.length; i++) {
            var selectedItem = docSelection[i];
            if (selectedItem.constructor.name === "TextFrame") {
                textRanges.push(selectedItem.textRange);
            } else if (selectedItem.constructor.name === "TextRange") {
                textRanges.push(selectedItem);
            }
        }
        return textRanges;
    }

    /**
     * テキスト範囲の先頭文字を取得する
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @returns {TextRange|null} 最初の文字（文字がなければ null）
     */
    function findFirstChar(textRanges) {
        for (var i = 0; i < textRanges.length; i++) {
            if (textRanges[i].characters.length > 0) return textRanges[i].characters[0];
        }
        return null;
    }

    /**
     * テキスト範囲のすべての文字にコールバックを適用する
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @param {function} charCallback - 各文字に対して実行する処理
     * @returns {void}
     */
    function forEachChar(textRanges, charCallback) {
        for (var i = 0; i < textRanges.length; i++) {
            var characters = textRanges[i].characters;
            for (var j = 0; j < characters.length; j++) {
                charCallback(characters[j]);
            }
        }
    }

    // =========================================
    // PreviewManager
    // プレビュー時にUndo履歴を汚さないための小さな管理クラス。
    // - addStep(): 変更処理を実行してundoDepthをカウント
    // - rollback(): 適用済みのプレビューをすべて取り消し
    // undoDepth = 適用してまだ取り消していないステップ数 / steps applied and not yet undone
    // 適用中の例外でダイアログごと落ちないよう addStep だけ try で囲み、
    // 上下キー連打で同じアラートが溢れないよう同一メッセージは1度だけ表示する。
    // =========================================

    /**
     * プレビュー適用とUndoの深さを管理するクラス
     * @returns {void}
     */
    function PreviewManager() {
        this.undoDepth = 0;
        var lastReportedError = "";

        /**
         * 変更処理を1ステップとして実行し、Undoの深さを数える
         * @param {function} applyChange - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (applyChange) {
            /* 文字属性の書き込みが失敗しうる / writing character attributes can fail */
            try {
                applyChange();
                lastReportedError = "";
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                var message = getLabel("alert.previewError") + e;
                if (message === lastReportedError) return;
                lastReportedError = message;
                alert(message);
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
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    /**
     * フォントサイズの調整パネルを構築する（イベントの配線は呼び出し側で行う）
     * @param {Window} parentDialog - 追加先のダイアログ
     * @param {string} unitLabel - サイズ欄に表示する単位ラベル
     * @returns {{sizeRow: object, scaleRow: object, apparentRow: object, convertButton: Button}} 構築したコントロール
     */
    function buildFontSizePanel(parentDialog, unitLabel) {
        var fontSizePanel = parentDialog.add("panel", undefined, getLabel("panel.fontSize"));
        setupPanel(fontSizePanel, FIELD_SPACING);

        var sizeRow = addRow(fontSizePanel, "fieldLabel.fontSize", "edittext", "0", unitLabel);
        var scaleRow = addRow(fontSizePanel, "fieldLabel.scale", "edittext", "100", "%");
        setRowTooltip(scaleRow, getLabel("tooltip.scale"));
        var apparentRow = addRow(fontSizePanel, "fieldLabel.apparent", "statictext", "--", unitLabel);
        setRowTooltip(apparentRow, getLabel("tooltip.apparent"));

        var convertButton = fontSizePanel.add("button", undefined, getLabel("button.toApparent"));
        convertButton.helpTip = getLabel("tooltip.toApparent");
        convertButton.alignment = "right";
        convertButton.preferredSize.width = CONVERT_BUTTON_WIDTH;

        alignLabelWidths(LABEL_WIDTH, [sizeRow.label, scaleRow.label, apparentRow.label]);
        return { sizeRow: sizeRow, scaleRow: scaleRow, apparentRow: apparentRow, convertButton: convertButton };
    }

    /**
     * 下部のボタン行を構築する（左＝リセット／右＝キャンセル・OK）
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{btnReset: Button, btnCancel: Button, btnOK: Button}} 構築したボタン
     */
    function buildButtonRow(parentDialog) {
        var btnRowGroup = parentDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "fill";
        btnRowGroup.alignChildren = ["fill", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignment = ["left", "center"];
        var btnReset = btnLeftGroup.add("button", undefined, getLabel("button.reset"));
        btnReset.helpTip = getLabel("tooltip.reset");

        /* 左右のボタンを両端に押し広げるスペーサー / spacer that pushes both sides apart */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "center"];

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;

        return { btnReset: btnReset, btnCancel: btnCancel, btnOK: btnOK };
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
        var unitLabel = textUnit.label;
        var unitFactor = textUnit.pointsPerUnit;

        var adjustDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        adjustDialog.alignChildren = "fill";
        adjustDialog.opacity = DIALOG_OPACITY;

        var fontSizeUI = buildFontSizePanel(adjustDialog, unitLabel);
        var buttonUI = buildButtonRow(adjustDialog);
        var sizeInput = fontSizeUI.sizeRow.control;
        var scaleInput = fontSizeUI.scaleRow.control;
        var apparentRow = fontSizeUI.apparentRow;

        /* 焼き込み前の状態（順方向で保存→逆方向で復元）。手動でサイズ/比率を変えたら無効化
           pre-bake state (saved on forward, restored on back); cleared when size/scale is edited by hand */
        var apparentToggleState = null;

        // ---- 値の適用・プレビュー / Apply values & preview ----

        /**
         * 現在の入力値を選択している文字にまとめて適用する（空欄の項目は適用しない）
         * @returns {void}
         */
        function applyCurrentValues() {
            var size = readNumber(sizeInput);
            var scale = readNumber(scaleInput);
            if (size === null && scale === null) return;
            var sizeInPt = (size === null) ? null : size * unitFactor;
            forEachChar(targetRanges, function (character) {
                if (sizeInPt !== null) character.size = sizeInPt;
                if (scale !== null) {
                    character.characterAttributes.horizontalScale = scale;
                    character.characterAttributes.verticalScale = scale;
                }
            });
        }

        /**
         * Undo履歴を汚さずにプレビューを更新し、見かけサイズの表示も更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(applyCurrentValues);
            updateApparentSizeDisplay();
        }

        // ---- 表示更新 / Display updates ----

        /**
         * 見かけサイズの表示を更新する（比率100%のときはディム表示）
         * @returns {void}
         */
        function updateApparentSizeDisplay() {
            var size = readNumber(sizeInput);
            var scale = readNumber(scaleInput);
            var hasValue = (size !== null && scale !== null);
            apparentRow.control.text = hasValue ? calculateApparentSize(size, scale) + "" : "--";
            setRowEnabled(apparentRow, scale !== 100);
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
            sizeInput.text = firstChar ? roundToTenth(firstChar.size / unitFactor) + "" : "";
            scaleInput.text = firstChar ? roundToTenth(firstChar.characterAttributes.horizontalScale) + "" : "";
            updateApparentSizeDisplay();
        }

        // ---- イベント / Events ----

        /**
         * サイズ・比率の確定入力を受けてプレビューと表示を更新する
         * @returns {void}
         */
        function onValueChanged() {
            apparentToggleState = null; /* 手動編集でトグル復元を無効化 / manual edit invalidates the toggle */
            updatePreview();
        }

        /* サイズ・比率は「入力値をそのまま適用」。loadValuesFromSelection() で入力欄を読み直すと
           入力値が丸めで戻る恐れがあるため onChange では呼ばない（見かけ表示だけ更新する）
           apply the typed value as-is; do NOT reload the fields on change (re-reading them
           could snap the typed value back via rounding). Only refresh the apparent readout */
        sizeInput.onChange = onValueChanged;
        scaleInput.onChange = onValueChanged;
        sizeInput.onChanging = updateApparentSizeDisplay;
        scaleInput.onChanging = updateApparentSizeDisplay;
        changeValueByArrowKey(sizeInput, { step: 1, shiftStep: 10, altStep: 0.1 });
        changeValueByArrowKey(scaleInput, { step: 1, shiftStep: 10, altStep: 5 });

        /* 実サイズ↔見かけのトグル / toggle between actual size and apparent (baked) size
           順方向：サイズ×比率を実サイズに焼き込み比率100%へ。逆方向：直前の比率付き状態へ戻す
           forward: bake size × scale into the actual size at 100%; back: restore the previous scaled state */
        fontSizeUI.convertButton.onClick = function () {
            var size = readNumber(sizeInput);
            var scale = readNumber(scaleInput);
            if (size === null || scale === null) return;
            if (apparentToggleState !== null) {
                sizeInput.text = apparentToggleState.size + "";
                scaleInput.text = apparentToggleState.scale + "";
                apparentToggleState = null;
            } else {
                apparentToggleState = { size: size, scale: scale };
                sizeInput.text = calculateApparentSize(size, scale) + "";
                scaleInput.text = "100";
            }
            updatePreview();
        };

        /* リセット：プレビューを取り消して開いた直後の状態に戻す。optionキー併用のときは
           そこからさらに、選択している文字すべてを先頭文字のサイズ・比率100%に統一して適用する
           Reset: undo the preview and return to the just-opened state; with the option key,
           also unify every selected character to the first character's size at 100% scale */
        buttonUI.btnReset.onClick = function () {
            previewManager.rollback();
            loadValuesFromSelection(); /* 調整前の先頭文字の値を読み直す / re-read the pre-adjustment values */
            if (!ScriptUI.environment.keyboardState.altKey) return;
            scaleInput.text = "100";
            updatePreview();
        };

        buttonUI.btnOK.onClick = function () {
            /* プレビュー分を戻し、本適用を1回だけ実行して確定（Undo履歴は1つ）
               undo the preview, then apply once so it lands as a single undo entry */
            previewManager.rollback();
            applyCurrentValues();
            adjustDialog.close();
        };

        buttonUI.btnCancel.onClick = function () {
            /* 開いてから適用した分をすべて取り消してから閉じる / undo everything applied since open, then close */
            previewManager.rollback();
            adjustDialog.close(2);
        };

        adjustDialog.onShow = function () {
            adjustDialog.location = [adjustDialog.location[0] + DIALOG_OFFSET_X, adjustDialog.location[1]];
            scaleInput.active = true;
        };

        /* 開いた時点では何も適用しない（現在の状態をそのまま保持）。値を変更したときだけプレビュー適用
           apply nothing on open (keep the current state as-is); preview only kicks in once a value changes */
        loadValuesFromSelection();
        adjustDialog.show();
    }

    main();

})();
