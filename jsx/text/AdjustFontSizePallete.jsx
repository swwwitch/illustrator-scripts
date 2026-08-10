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
                ja: "選択している文字すべてを先頭文字のフォントサイズに統一し、水平比率・垂直比率を100%に揃えて適用します。",
                en: "Unifies every selected character to the first character's font size and sets the horizontal / vertical scale to 100%, then applies."
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
     * ラベル＋入力欄＋単位の行を追加する
     * @param {Panel|Group} parent - 追加先のコンテナ
     * @param {string} labelKey - ラベルキー
     * @param {string} defaultValue - 入力欄の初期値
     * @param {string} unitText - 単位表示の文字列
     * @returns {{row: Group, label: StaticText, input: EditText, unit: StaticText}} 生成した行の各コントロール
     */
    function addFieldRow(parent, labelKey, defaultValue, unitText) {
        var row = parent.add("group");
        setupGroup(row, "row");
        var label = row.add("statictext", undefined, getLabelWithColon(labelKey));
        label.justify = "right";
        var input = row.add("edittext", undefined, defaultValue);
        input.characters = 4;
        input.justify = "right";
        var unit = row.add("statictext", undefined, unitText);
        return { row: row, label: label, input: input, unit: unit };
    }

    /**
     * ラベル＋表示専用テキスト＋単位の行を追加する
     * @param {Panel|Group} parent - 追加先のコンテナ
     * @param {string} labelKey - ラベルキー
     * @param {string} unitText - 単位表示の文字列
     * @returns {{row: Group, label: StaticText, value: StaticText, unit: StaticText}} 生成した行の各コントロール
     */
    function addReadoutRow(parent, labelKey, unitText) {
        var row = parent.add("group");
        setupGroup(row, "row");
        var label = row.add("statictext", undefined, getLabelWithColon(labelKey));
        label.justify = "right";
        var value = row.add("statictext", undefined, "--");
        value.characters = 5;
        var unit = row.add("statictext", undefined, unitText);
        return { row: row, label: label, value: value, unit: unit };
    }

    /**
     * 行（ラベル＋入力／表示＋単位）にまとめてヘルプチップを設定する
     * @param {object} fieldRow - addFieldRow / addReadoutRow が返した行オブジェクト
     * @param {string} tooltip - 設定するヘルプチップ文字列
     * @returns {void}
     */
    function setFieldTooltip(fieldRow, tooltip) {
        fieldRow.label.helpTip = tooltip;
        fieldRow.unit.helpTip = tooltip;
        if (fieldRow.input) fieldRow.input.helpTip = tooltip;
        if (fieldRow.value) fieldRow.value.helpTip = tooltip;
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
        var selection = doc.selection;
        var ranges = [];
        if (!selection) return ranges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (selection.constructor.name === "TextRange") {
            ranges.push(selection);
            return ranges;
        }
        if (selection.length === 0) return ranges;
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

    // =========================================
    // エラー処理 / Error reporting
    // =========================================

    /* 直近に表示したエラーメッセージ（連続する同一エラーの抑制用）/ last shown error message (to suppress repeats) */
    var lastReportedError = "";

    /**
     * 同一メッセージの連続表示を抑制してアラートを出す（プレビュー連打によるアラート氾濫を防ぐ）
     * @param {string} prefix - メッセージの先頭に付ける説明
     * @param {Error} e - 捕捉した例外
     * @returns {void}
     */
    function reportError(prefix, e) {
        var message = prefix + String(e);
        if (message === lastReportedError) return;
        lastReportedError = message;
        alert(message);
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
         * @param {function} func - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (func) {
            try {
                func();
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
            if (typeof finalAction === "function") {
                this.rollback();
                finalAction();
            } else {
                this.undoDepth = 0;
            }
        };
    }

    /* UI入力欄の参照をまとめるオブジェクト / references to UI input controls */
    var uiElements = {
        sizeInput: null,
        scaleInput: null,
        apparentSizeText: null
    };

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

        /**
         * ポイント値をルーラー単位に換算する
         * @param {number} pt - ポイント値
         * @returns {number} ルーラー単位の値
         */
        function ptToUnit(pt) {
            return pt / unitFactor;
        }

        /**
         * ルーラー単位の値をポイントに換算する
         * @param {number} value - ルーラー単位の値
         * @returns {number} ポイント値
         */
        function unitToPt(value) {
            return value * unitFactor;
        }

        /**
         * 選択している文字の先頭を取得する
         * @returns {Characters|null} 最初の文字（選択している文字がなければ null）
         */
        function findFirstChar() {
            for (var i = 0; i < targetRanges.length; i++) {
                var characters = targetRanges[i].characters;
                if (characters.length > 0) return characters[0];
            }
            return null;
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

        // ---- 値の取得・適用 / Read & apply values ----

        /**
         * 選択している文字すべてにコールバックを適用する
         * @param {function} action - 各文字に対して実行する処理
         * @returns {void}
         */
        function forEachSelectedChar(action) {
            for (var i = 0; i < targetRanges.length; i++) {
                var characters = targetRanges[i].characters;
                for (var j = 0; j < characters.length; j++) {
                    action(characters[j]);
                }
            }
        }

        /**
         * 現在のUIの入力値をまとめて取得する（空欄やNaNは null）
         * @returns {{size: number|null, scale: number|null}} フォントサイズと比率
         */
        function getCurrentParams() {
            var sizeValue = parseFloat(uiElements.sizeInput.text);
            var scaleValue = parseFloat(uiElements.scaleInput.text);
            return {
                size: isNaN(sizeValue) ? null : sizeValue,
                scale: isNaN(scaleValue) ? null : scaleValue
            };
        }

        /**
         * 現在のUIの値を選択している文字にまとめて適用する
         * @returns {void}
         */
        function applyAllCurrentValues() {
            var params = getCurrentParams();

            if (params.size !== null) {
                var sizeInPt = unitToPt(params.size);
                forEachSelectedChar(function (character) {
                    character.size = sizeInPt;
                });
            }
            if (params.scale !== null) {
                forEachSelectedChar(function (character) {
                    character.characterAttributes.horizontalScale = params.scale;
                    character.characterAttributes.verticalScale = params.scale;
                });
            }
        }

        // ---- プレビュー / Preview ----

        /**
         * Undo履歴を汚さずにプレビューを更新する
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(function () {
                applyAllCurrentValues();
            });
        }

        // ---- 表示更新 / Display updates ----

        /**
         * 見かけサイズの表示を更新する（比率100%のときはディム表示）
         * @returns {void}
         */
        function updateApparentSizeDisplay() {
            var size = parseFloat(uiElements.sizeInput.text);
            var scale = parseFloat(uiElements.scaleInput.text);
            if (isNaN(size) || isNaN(scale)) {
                uiElements.apparentSizeText.text = "--";
            } else {
                uiElements.apparentSizeText.text = calculateApparentSize(size, scale) + "";
            }

            var isDimmed = (scale === 100);
            apparentRow.label.enabled = !isDimmed;
            apparentRow.value.enabled = !isDimmed;
            apparentRow.unit.enabled = !isDimmed;
        }

        /**
         * 選択している文字の先頭の現在値を読み取って入力欄に反映する
         * @returns {void}
         */
        function updateInfoText() {
            /* 実際の値を読み直すので、焼き込み前の保存状態（トグル）は破棄する
               Reloading the actual values invalidates the saved pre-bake (toggle) state */
            apparentToggleState = null;
            var firstChar = findFirstChar();
            if (!firstChar) {
                uiElements.sizeInput.text = "";
                uiElements.scaleInput.text = "";
                uiElements.apparentSizeText.text = "--";
                return;
            }
            var sizeInUnit = Math.round(ptToUnit(firstChar.size) * 10) / 10;
            var horizontalScale = Math.round(firstChar.characterAttributes.horizontalScale * 10) / 10;
            uiElements.sizeInput.text = sizeInUnit + "";
            uiElements.scaleInput.text = horizontalScale + "";
            updateApparentSizeDisplay();
        }

        // ---- 値のリセット / Reset values ----

        /**
         * 選択している文字すべてを先頭文字のフォントサイズに統一し、比率を100%に揃えて適用する
         * @returns {void}
         */
        function resetToUniformSize() {
            apparentToggleState = null;
            var firstChar = findFirstChar();
            if (firstChar) {
                uiElements.sizeInput.text = Math.round(ptToUnit(firstChar.size) * 10) / 10;
            }
            uiElements.scaleInput.text = "100";
            updatePreview();
            updateApparentSizeDisplay();
        }

        // =========================================
        // UI構築 / Build UI
        // =========================================
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.opacity = DIALOG_OPACITY;

        /* フォントサイズの調整パネル / font-size adjustment panel */
        var fontSizePanel = dialog.add("panel", undefined, getLabel("panel.fontSize"));
        setupPanel(fontSizePanel, FIELD_SPACING);

        var sizeField = addFieldRow(fontSizePanel, "fieldLabel.fontSize", "0", unitLabel);
        uiElements.sizeInput = sizeField.input;

        var scaleField = addFieldRow(fontSizePanel, "fieldLabel.scale", "100", "%");
        uiElements.scaleInput = scaleField.input;
        setFieldTooltip(scaleField, getLabel("tooltip.scale"));

        var apparentRow = addReadoutRow(fontSizePanel, "fieldLabel.apparent", unitLabel);
        uiElements.apparentSizeText = apparentRow.value;
        setFieldTooltip(apparentRow, getLabel("tooltip.apparent"));

        /* 焼き込み前の状態（順方向で保存→逆方向で復元）。手動でサイズ/比率を変えたら無効化
           pre-bake state (saved on forward, restored on back); cleared when size/scale is edited by hand */
        var apparentToggleState = null;

        /* 実サイズ↔見かけのトグルボタン / toggle between actual size and apparent (baked) size
           順方向：サイズ×比率を実サイズに焼き込み比率100%へ。逆方向：直前の比率付き状態へ戻す
           forward: bake size × scale into the actual size at 100%; back: restore the previous scaled state */
        var convertButton = fontSizePanel.add("button", undefined, getLabel("button.toApparent"));
        convertButton.helpTip = getLabel("tooltip.toApparent");
        convertButton.alignment = "right";
        convertButton.preferredSize.width = CONVERT_BUTTON_WIDTH;
        convertButton.onClick = function () {
            var size = parseFloat(uiElements.sizeInput.text);
            var scale = parseFloat(uiElements.scaleInput.text);
            if (isNaN(size) || isNaN(scale)) return;
            if (apparentToggleState !== null) {
                /* 見かけ→実サイズ：焼き込み前の比率付き状態に戻す / restore the pre-bake scaled state */
                uiElements.sizeInput.text = apparentToggleState.size + "";
                uiElements.scaleInput.text = apparentToggleState.scale + "";
                apparentToggleState = null;
            } else {
                /* 実サイズ→見かけ：比率をサイズへ焼き込み100%に / bake scale into size and reset to 100% */
                apparentToggleState = { size: size, scale: scale };
                uiElements.sizeInput.text = calculateApparentSize(size, scale) + "";
                uiElements.scaleInput.text = "100";
            }
            updatePreview();
            updateApparentSizeDisplay();
        };

        alignLabelWidths(LABEL_WIDTH, [sizeField.label, scaleField.label, apparentRow.label]);

        /* 入力欄のイベント / wire input events
           サイズ・比率は「入力値をそのまま適用」。updateInfoText() で入力欄を読み直すと
           入力値が丸めで戻る恐れがあるため呼ばない（見かけ表示だけ更新する）
           apply the typed value as-is; do NOT call updateInfoText() here (re-reading the field
           could snap the typed value back via rounding). Only refresh the apparent readout */
        uiElements.sizeInput.onChange = function () {
            apparentToggleState = null; /* 手動編集でトグル復元を無効化 / manual edit invalidates the toggle */
            updatePreview();
            updateApparentSizeDisplay();
        };
        uiElements.sizeInput.onChanging = function () {
            updateApparentSizeDisplay();
        };
        changeValueByArrowKey(uiElements.sizeInput, { step: 1, shiftStep: 10, altStep: 0.1 });

        uiElements.scaleInput.onChange = function () {
            apparentToggleState = null; /* 手動編集でトグル復元を無効化 / manual edit invalidates the toggle */
            updatePreview();
            updateApparentSizeDisplay();
        };
        uiElements.scaleInput.onChanging = function () {
            updateApparentSizeDisplay();
        };
        changeValueByArrowKey(uiElements.scaleInput, { step: 1, shiftStep: 10, altStep: 5 });

        /* ボタン（下部・3カラム: 左=リセット / 中央=スペーサー / 右=OK・キャンセル）/ buttons (bottom, 3 columns) */
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "fill";
        buttonGroup.alignChildren = ["fill", "center"];

        var buttonLeft = buttonGroup.add("group");
        buttonLeft.alignment = ["left", "center"];
        buttonLeft.spacing = FIELD_SPACING;
        var resetButton = buttonLeft.add("button", undefined, getLabel("button.reset"));
        resetButton.helpTip = getLabel("tooltip.reset");

        var buttonCenter = buttonGroup.add("group");
        buttonCenter.alignment = ["fill", "center"];

        var buttonRight = buttonGroup.add("group");
        buttonRight.alignment = ["right", "center"];
        var cancelButton = buttonRight.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = buttonRight.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        okButton.preferredSize.width = BUTTON_WIDTH;
        cancelButton.preferredSize.width = BUTTON_WIDTH;

        okButton.onClick = function () {
            /* プレビューを1回の適用として確定 / commit the preview as a single apply */
            previewManager.confirm(function () {
                applyAllCurrentValues();
            });
            dialog.close();
        };
        cancelButton.onClick = function () {
            /* 開いてから適用した分をすべて取り消してから閉じる / undo everything applied since open, then close */
            previewManager.rollback();
            dialog.close(2);
        };
        resetButton.onClick = function () {
            resetToUniformSize();
        };

        /* 初期表示（updateInfoText が先頭文字の実サイズを読み取って各欄を設定）/ initial state (updateInfoText reads the actual size of the first char) */
        updateInfoText();
        updateApparentSizeDisplay();

        dialog.onShow = function () {
            dialog.location = [dialog.location[0] + DIALOG_OFFSET_X, dialog.location[1]];
            uiElements.scaleInput.active = true;
        };

        /* 開いた時点では何も適用しない（現在の状態をそのまま保持）。値を変更したときだけプレビュー適用
           apply nothing on open (keep the current state as-is); preview only kicks in once a value changes */
        dialog.show();
    }

    main();

})();
