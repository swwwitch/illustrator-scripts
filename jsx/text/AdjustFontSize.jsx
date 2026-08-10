#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択している文字を対象に、フォントサイズと水平比率／垂直比率を調整する
- ライブプレビューで結果を確認しながら調整でき、キャンセルで開く前の状態に戻る
- 「実サイズ↔見かけ」でサイズ×比率の焼き込みと復元を切り替えられる

### 注意

- 詳しい機能・使い方は README を参照

### Overview

- Adjusts the font size and horizontal / vertical scale of the selected characters
- Live preview; Cancel restores the state from before the dialog opened
- “Actual ↔ Apparent” toggles between baking size × scale into the size and restoring it

### Notes

- See the README for the full feature list and usage

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustFontSize";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustFontSize.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustFontSize.md
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
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

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
        var parts = key.split(".");
        var label = LABELS[parts[0]][parts[1]];
        return label[currentLanguage] || label.en;
    }

    /**
     * コロン付きのラベル文字列を取得する（日本語は全角、英語は半角）
     * @param {string} key - ラベルキー
     * @returns {string} コロンを付けたラベル文字列
     */
    function getLabelWithColon(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * 単位コードとプリファレンスキーに応じて単位ラベルを返す
     * @param {number} code - Illustratorの単位コード
     * @param {string} prefKey - 単位を取得したプリファレンスキー
     * @returns {string} 単位ラベル（"mm" / "pt" / "Q" / "H" など）
     */
    function getUnitLabel(code, prefKey) {
        var unitMap = {
            0: "in",
            1: "mm",
            2: "pt",
            3: "pica",
            4: "cm",
            5: "Q/H",
            6: "px",
            7: "ft/in",
            8: "m",
            9: "yd",
            10: "ft"
        };
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return unitMap[code] || "pt";
    }

    /**
     * 単位コード1つ分のポイント数（pt換算係数）を返す
     * @param {number} code - Illustratorの単位コード
     * @returns {number} 単位1に相当するポイント数
     */
    function getPointsPerUnit(code) {
        var MM = 72 / 25.4; /* 1mm = 72/25.4 pt */
        var ptPerUnit = {
            0: 72,        /* in */
            1: MM,        /* mm */
            2: 1,         /* pt */
            3: 12,        /* pica */
            4: MM * 10,   /* cm */
            5: MM * 0.25, /* Q/H（0.25mm）/ Q/H (0.25mm) */
            6: 1,         /* px（72ppi）/ px (72ppi) */
            7: 72,        /* ft/in */
            8: MM * 1000, /* m */
            9: 72 * 36,   /* yd */
            10: 72 * 12   /* ft */
        };
        return ptPerUnit[code] || 1;
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
        var value = parseFloat(control.text);
        return isNaN(value) ? null : value;
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
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は既定値）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループに共通のレイアウト設定を適用する
     * @param {Group} group - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は既定値）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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
        var row = parent.add("group");
        setupGroup(row, "row");
        var label = row.add("statictext", undefined, getLabelWithColon(labelKey));
        label.justify = "right";
        var isEditable = (controlType === "edittext");
        var control = row.add(controlType, undefined, initialText);
        control.characters = isEditable ? EDIT_CHARACTERS : READOUT_CHARACTERS;
        if (isEditable) control.justify = "right";
        var unit = row.add("statictext", undefined, unitText);
        return { label: label, control: control, unit: unit };
    }

    /**
     * 行（ラベル＋コントロール＋単位）にまとめてヘルプチップを設定する
     * @param {{label: StaticText, control: EditText|StaticText, unit: StaticText}} row - addRow が返した行オブジェクト
     * @param {string} tooltip - 設定するヘルプチップ文字列
     * @returns {void}
     */
    function setRowTooltip(row, tooltip) {
        row.label.helpTip = tooltip;
        row.control.helpTip = tooltip;
        row.unit.helpTip = tooltip;
    }

    /**
     * 行のラベル・コントロール・単位をまとめて有効／無効にする
     * @param {{label: StaticText, control: EditText|StaticText, unit: StaticText}} row - addRow が返した行オブジェクト
     * @param {boolean} enabled - 有効にするなら true
     * @returns {void}
     */
    function setRowEnabled(row, enabled) {
        row.label.enabled = enabled;
        row.control.enabled = enabled;
        row.unit.enabled = enabled;
    }

    /**
     * 複数ラベルの幅を揃える
     * @param {number} width - 設定する幅（px）
     * @param {StaticText[]} labels - 幅を揃えるラベルの配列
     * @returns {void}
     */
    function alignLabelWidths(width, labels) {
        for (var i = 0; i < labels.length; i++) {
            labels[i].preferredSize.width = width;
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
        var selection = app.activeDocument.selection;
        var ranges = [];
        if (!selection) return ranges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (selection.constructor.name === "TextRange") {
            ranges.push(selection);
            return ranges;
        }
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.constructor.name === "TextFrame") {
                ranges.push(item.textRange);
            } else if (item.constructor.name === "TextRange") {
                ranges.push(item);
            }
        }
        return ranges;
    }

    /**
     * テキスト範囲の先頭文字を取得する
     * @param {TextRange[]} ranges - 対象のテキスト範囲
     * @returns {Characters|null} 最初の文字（文字がなければ null）
     */
    function findFirstChar(ranges) {
        for (var i = 0; i < ranges.length; i++) {
            if (ranges[i].characters.length > 0) return ranges[i].characters[0];
        }
        return null;
    }

    /**
     * テキスト範囲のすべての文字にコールバックを適用する
     * @param {TextRange[]} ranges - 対象のテキスト範囲
     * @param {function} action - 各文字に対して実行する処理
     * @returns {void}
     */
    function forEachChar(ranges, action) {
        for (var i = 0; i < ranges.length; i++) {
            var characters = ranges[i].characters;
            for (var j = 0; j < characters.length; j++) {
                action(characters[j]);
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
         * @param {function} func - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (func) {
            try {
                func();
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
     * @param {Window} dialog - 追加先のダイアログ
     * @param {string} unitLabel - サイズ欄に表示する単位ラベル
     * @returns {{sizeRow: object, scaleRow: object, apparentRow: object, convertButton: Button}} 構築したコントロール
     */
    function buildFontSizePanel(dialog, unitLabel) {
        var fontSizePanel = dialog.add("panel", undefined, getLabel("panel.fontSize"));
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
     * @param {Window} dialog - 追加先のダイアログ
     * @returns {{resetButton: Button, cancelButton: Button, okButton: Button}} 構築したボタン
     */
    function buildButtonRow(dialog) {
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "fill";
        buttonGroup.alignChildren = ["fill", "center"];

        var leftGroup = buttonGroup.add("group");
        leftGroup.alignment = ["left", "center"];
        var resetButton = leftGroup.add("button", undefined, getLabel("button.reset"));
        resetButton.helpTip = getLabel("tooltip.reset");

        /* 左右のボタンを両端に押し広げるスペーサー / spacer that pushes both sides apart */
        var spacerGroup = buttonGroup.add("group");
        spacerGroup.alignment = ["fill", "center"];

        var rightGroup = buttonGroup.add("group");
        rightGroup.alignment = ["right", "center"];
        var cancelButton = rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        cancelButton.preferredSize.width = BUTTON_WIDTH;
        okButton.preferredSize.width = BUTTON_WIDTH;

        return { resetButton: resetButton, cancelButton: cancelButton, okButton: okButton };
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
        var unitCode = app.preferences.getIntegerPreference("text/units");
        var unitLabel = getUnitLabel(unitCode, "text/units");
        var unitFactor = getPointsPerUnit(unitCode);

        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.opacity = DIALOG_OPACITY;

        var fontSizeUI = buildFontSizePanel(dialog, unitLabel);
        var buttonUI = buildButtonRow(dialog);
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
         * Undo履歴を汚さずにプレビューを更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(applyCurrentValues);
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
            updateApparentSizeDisplay();
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
            updateApparentSizeDisplay();
        };

        /* リセット：選択している文字すべてを先頭文字のサイズに統一し、比率100%で適用
           Reset: unify every selected character to the first character's size and apply at 100% scale */
        /* リセット：プレビューを取り消して開いた直後の状態に戻す。optionキー併用のときは
           そこからさらに、選択している文字すべてを先頭文字のサイズ・比率100%に統一して適用する
           Reset: undo the preview and return to the just-opened state; with the option key,
           also unify every selected character to the first character's size at 100% scale */
        buttonUI.resetButton.onClick = function () {
            previewManager.rollback();
            loadValuesFromSelection(); /* 調整前の先頭文字の値を読み直す / re-read the pre-adjustment values */
            if (!ScriptUI.environment.keyboardState.altKey) return;
            scaleInput.text = "100";
            updatePreview();
            updateApparentSizeDisplay();
        };

        buttonUI.okButton.onClick = function () {
            /* プレビュー分を戻し、本適用を1回だけ実行して確定（Undo履歴は1つ）
               undo the preview, then apply once so it lands as a single undo entry */
            previewManager.rollback();
            applyCurrentValues();
            dialog.close();
        };

        buttonUI.cancelButton.onClick = function () {
            /* 開いてから適用した分をすべて取り消してから閉じる / undo everything applied since open, then close */
            previewManager.rollback();
            dialog.close(2);
        };

        dialog.onShow = function () {
            dialog.location = [dialog.location[0] + DIALOG_OFFSET_X, dialog.location[1]];
            scaleInput.active = true;
        };

        /* 開いた時点では何も適用しない（現在の状態をそのまま保持）。値を変更したときだけプレビュー適用
           apply nothing on open (keep the current state as-is); preview only kicks in once a value changes */
        loadValuesFromSelection();
        dialog.show();
    }

    main();

})();
