#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストのベースラインシフトを調整します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBaselineShifter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n5e41727cf265

### Overview

Adjusts the baseline shift of the selected text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBaselineShifter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartBaselineShifter";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-21";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartBaselineShifter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartBaselineShifter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n5e41727cf265"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_REFERENCE_CHAR = "0"; /* 基準文字の初期値 / Initial reference character */
    var SHIFT_DECIMAL_PLACES   = 4;   /* 計算したシフト量の小数桁数 / Decimal places of the calculated shift amount */

    /* 対象文字の初期値から外す文字（空白・改行・英数字・ひらがな・カタカナ・漢字）
       Characters left out of the initial target (whitespace, line breaks, alphanumerics, kana, kanji) */
    var NON_SYMBOL_CHAR_PATTERN = /^[\x00-\x20 　A-Za-z0-9぀-ゟ゠-ヿ一-鿿]$/;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var INPUT_COLUMN_MARGINS       = [15, 5, 15, 5];  /* 入力欄の列の余白 [左,上,右,下] / Input column margins */
    var SHIFT_ROW_MARGINS          = [0, 0, 0, 10];   /* シフト量の行の余白 [左,上,右,下] / Shift amount row margins */
    var AUTO_ADJUST_PANEL_MARGINS  = [15, 20, 15, 5]; /* 自動調整パネルの余白 [左,上,右,下] / Auto adjust panel margins */
    var TEXT_INPUT_CHARACTERS      = 6;               /* 対象文字・シフト量の欄の幅（文字数）/ Width of the target and shift fields */
    var REFERENCE_INPUT_CHARACTERS = 3;               /* 基準文字の欄の幅（文字数）/ Width of the reference field */
    var CALCULATE_BUTTON_BOUNDS    = [0, 0, 60, 25];  /* 計算ボタンの大きさ / Calculate button bounds */
    var BUTTON_SPACER_BOUNDS       = [0, 0, 0, 30];   /* キャンセルとリセットの間隔 / Gap between Cancel and Reset */
    var DIALOG_OFFSET_X            = 300;             /* ダイアログを右へずらす量 / Horizontal dialog offset */
    var DIALOG_OPACITY             = 0.97;            /* ダイアログの不透明度 / Dialog opacity */

    /**
     * ↑↓キーで数値を増減できるようにする（Shift で10刻み、Option で0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {function} [onChanged] - 値が変わったときに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onChanged) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            event.preventDefault();

            if (keyboard.shiftKey) {
                /* Shift：10 単位にスナップ / Shift snaps to multiples of 10 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                /* Option：0.1 刻み / Option steps by 0.1 */
                value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                value = Math.round(value + (isUp ? 1 : -1));
            }

            editText.text = value;
            if (typeof onChanged === "function") onChanged();
        });
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
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "ベースライン調整", en: "Adjust Baseline" }
        },
        panel: {
            autoAdjust: { ja: "自動調整（天地）", en: "Auto Adjust (Vertical)" }
        },
        fieldLabel: {
            targetChars: { ja: "対象文字", en: "Target Character" },
            shiftAmount: { ja: "シフト量", en: "Shift Amount" },
            referenceChar: { ja: "基準文字", en: "Reference Character" }
        },
        button: {
            adjust: { ja: "調整", en: "Adjust" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" },
            calculate: { ja: "計算", en: "Calculate" }
        },
        tooltip: {
            targetChars: {
                ja: "ベースラインをシフトする対象文字を入力します。",
                en: "Enter the character(s) to shift."
            },
            shiftAmount: {
                ja: "手動で指定するベースラインシフト量（数値）です。",
                en: "Specify the baseline shift amount manually."
            },
            referenceChar: {
                ja: "基準となる文字を1文字入力します。",
                en: "Enter the reference character (1 character)."
            },
            calculate: {
                ja: "対象文字と基準文字からシフト量を自動計算します。",
                en: "Calculate shift amount automatically."
            },
            adjust: {
                ja: "指定したシフト量を確定して適用します。",
                en: "Apply the specified shift amount."
            },
            reset: {
                ja: "選択しているテキストのベースラインシフトを全リセットします。",
                en: "Reset baseline shifts in all text frames."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            selectTextFrame: { ja: "テキストフレームを選択してください。", en: "Select one or more text frames." },
            invalidChars: {
                ja: "対象文字は1文字以上、基準文字は1文字を入力してください。",
                en: "Enter at least one target character and exactly one reference character."
            },
            targetNotFound: { ja: "対象文字が含まれていません。", en: "Target character not found." },
            shiftNotNumber: { ja: "シフト量は数値で入力してください。", en: "Shift amount must be a number." },
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
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ": ");
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューで加えた変更を数えておき、app.undo() でまとめて取り消す
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;
    }

    /**
     * 変更を実行し、取り消す段数を1つ増やす
     * @param {function} changeFunc - 実行する変更
     * @returns {void}
     */
    PreviewManager.prototype.addStep = function (changeFunc) {
        /* プレビュー中の失敗はアラートを出さずに見送る / Preview failures are skipped without an alert */
        try {
            changeFunc();
            this.undoDepth++;
            app.redraw();
        } catch (e) { }
    };

    /**
     * プレビューで加えた変更をすべて取り消す
     * @returns {void}
     */
    PreviewManager.prototype.rollback = function () {
        while (this.undoDepth > 0) {
            /* 取り消せなくなったら段数を捨てて数え違いを残さない / If undo fails, drop the count so it cannot drift */
            try {
                app.undo();
            } catch (e) {
                this.undoDepth = 0;
                break;
            }
            this.undoDepth--;
        }
        app.redraw();
    };

    /**
     * プレビューを取り消してから本番の処理を1回だけ実行する（取り消しを1段にまとめる）
     * @param {function} finalAction - 本番の処理
     * @returns {void}
     */
    PreviewManager.prototype.confirm = function (finalAction) {
        this.rollback();
        finalAction();
    };

    // =========================================
    // 対象の収集 / Collecting targets
    // =========================================

    /**
     * 文字の編集中で、そのストーリーがテキストフレーム1つだけなら、そのフレームを選択し直す
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function selectFrameOfEditedText(doc) {
        var selection = doc.selection;
        if (!selection || selection.typename !== "TextRange") return;

        var storyFrames = selection.story.textFrames;
        if (storyFrames.length !== 1) return;

        app.executeMenuCommand("deselectall");
        doc.selection = [storyFrames[0]];
        app.selectTool("Adobe Select Tool");
    }

    /**
     * アイテムがテキストフレームなら追加し、グループなら中を再帰的にたどる
     * @param {PageItem} item - 調べるアイテム
     * @param {TextFrame[]} textFrames - 見つけたフレームを追加する配列
     * @returns {void}
     */
    function appendTextFrames(item, textFrames) {
        if (item.typename === "TextFrame") {
            textFrames.push(item);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                appendTextFrames(item.pageItems[i], textFrames);
            }
        }
    }

    /**
     * 選択（グループの中を含む）からテキストフレームを集める
     * @param {Array<PageItem>|TextRange} selection - ドキュメントの選択
     * @returns {TextFrame[]} テキストフレーム（文字の編集中は空）
     */
    function collectTextFrames(selection) {
        var textFrames = [];
        if (!selection || selection.typename === "TextRange") return textFrames;
        for (var i = 0; i < selection.length; i++) {
            appendTextFrames(selection[i], textFrames);
        }
        return textFrames;
    }

    /**
     * 英数字・かな・漢字以外の文字を、出てきた順に重複なく集める（対象文字の初期値）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {string} 集めた文字
     */
    function collectSymbolChars(textFrames) {
        var symbolChars = "";
        for (var i = 0; i < textFrames.length; i++) {
            var frameText = textFrames[i].contents;
            for (var j = 0; j < frameText.length; j++) {
                var character = frameText.charAt(j);
                if (!NON_SYMBOL_CHAR_PATTERN.test(character) && symbolChars.indexOf(character) === -1) symbolChars += character;
            }
        }
        return symbolChars;
    }

    // =========================================
    // ベースラインシフト / Baseline shift
    // =========================================

    /**
     * テキストフレームのベースラインシフトをすべて0に戻す
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function resetBaselineShift(textFrames) {
        for (var i = 0; i < textFrames.length; i++) {
            /* 書き換えられないフレームは飛ばして残りを続ける / Skip frames that cannot be modified and carry on */
            try {
                textFrames[i].textRange.characterAttributes.baselineShift = 0;
            } catch (e) { }
        }
    }

    /**
     * フレーム内の対象文字すべてにベースラインシフトを設定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（複数可）
     * @param {number} shiftAmount - ベースラインシフトの値（pt）
     * @returns {void}
     */
    function applyBaselineShift(textFrame, targetChars, shiftAmount) {
        var characters = textFrame.textRange.characters;
        for (var i = 0; i < characters.length; i++) {
            var character = characters[i].contents;
            if (character && targetChars.indexOf(character) !== -1) characters[i].characterAttributes.baselineShift = shiftAmount;
        }
    }

    /**
     * ベースラインシフトをいったん全部戻し、対象文字だけにシフト量を設定する
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（複数可）
     * @param {number} shiftAmount - ベースラインシフトの値（pt）
     * @returns {void}
     */
    function applyShiftToAll(textFrames, targetChars, shiftAmount) {
        resetBaselineShift(textFrames);
        for (var i = 0; i < textFrames.length; i++) {
            applyBaselineShift(textFrames[i], targetChars, shiftAmount);
        }
    }

    // =========================================
    // 文字の中心の実測 / Measuring character centers
    // =========================================

    /**
     * アイテムの天地中央のY座標を求める
     * @param {PageItem} item - 対象のアイテム
     * @returns {number} 天地中央のY座標
     */
    function getCenterY(item) {
        var bounds = item.geometricBounds;
        return (bounds[1] + bounds[3]) / 2;
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

    /**
     * 対象文字を含む最初のフレームで、基準文字と対象文字の天地中央の差（シフト量）を求める
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（最初に見つかった1文字で測る）
     * @param {string} referenceChar - 基準文字（1文字）
     * @returns {number|null} シフト量（pt）。対象文字が見つからなければ null
     */
    function calculateShiftAmount(textFrames, targetChars, referenceChar) {
        for (var i = 0; i < textFrames.length; i++) {
            var frameText = textFrames[i].contents;
            for (var j = 0; j < targetChars.length; j++) {
                var targetChar = targetChars.charAt(j);
                if (frameText.indexOf(targetChar) === -1) continue;

                var referenceCenterY = measureCharCenterY(textFrames[i], referenceChar);
                return referenceCenterY - measureCharCenterY(textFrames[i], targetChar);
            }
        }
        return null;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名＋入力欄の1行を追加する
     * @param {Group|Panel} parent - 追加先
     * @param {object} labelSet - 項目名のラベル
     * @param {string} initialText - 入力欄の初期値
     * @param {number} characters - 入力欄の幅（文字数）
     * @returns {EditText} 追加した入力欄（行のグループは parent で取れる）
     */
    function addFieldRow(parent, labelSet, initialText, characters) {
        var fieldRow = parent.add("group");
        fieldRow.add("statictext", undefined, labelText(labelSet));
        var fieldInput = fieldRow.add("edittext", undefined, initialText);
        fieldInput.characters = characters;
        return fieldInput;
    }

    /**
     * 左の列（対象文字・シフト量・自動調整パネル）を組む
     * @param {Group} parent - 追加先
     * @param {string} defaultTargetChars - 対象文字の初期値
     * @param {string} shiftUnitLabel - シフト量の単位の表示
     * @returns {{targetInput: EditText, shiftInput: EditText, referenceInput: EditText, btnCalculate: Button}} 入力欄と計算ボタン
     */
    function buildInputColumn(parent, defaultTargetChars, shiftUnitLabel) {
        var inputColumn = parent.add("group");
        inputColumn.orientation = "column";
        inputColumn.alignChildren = "left";
        inputColumn.margins = INPUT_COLUMN_MARGINS;

        var targetInput = addFieldRow(inputColumn, LABELS.fieldLabel.targetChars, defaultTargetChars, TEXT_INPUT_CHARACTERS);
        targetInput.helpTip = getLabel(LABELS.tooltip.targetChars);

        var shiftInput = addFieldRow(inputColumn, LABELS.fieldLabel.shiftAmount, "0", TEXT_INPUT_CHARACTERS);
        shiftInput.helpTip = getLabel(LABELS.tooltip.shiftAmount);
        shiftInput.parent.add("statictext", undefined, shiftUnitLabel);
        shiftInput.parent.margins = SHIFT_ROW_MARGINS;
        shiftInput.active = true;

        var autoAdjustPanel = inputColumn.add("panel", undefined, getLabel(LABELS.panel.autoAdjust));
        autoAdjustPanel.orientation = "column";
        autoAdjustPanel.alignChildren = "left";
        autoAdjustPanel.margins = AUTO_ADJUST_PANEL_MARGINS;

        var referenceInput = addFieldRow(autoAdjustPanel, LABELS.fieldLabel.referenceChar, DEFAULT_REFERENCE_CHAR, REFERENCE_INPUT_CHARACTERS);
        referenceInput.helpTip = getLabel(LABELS.tooltip.referenceChar);
        var btnCalculate = referenceInput.parent.add("button", CALCULATE_BUTTON_BOUNDS, getLabel(LABELS.button.calculate));
        btnCalculate.helpTip = getLabel(LABELS.tooltip.calculate);

        return { targetInput: targetInput, shiftInput: shiftInput, referenceInput: referenceInput, btnCalculate: btnCalculate };
    }

    /**
     * 右の列（調整・キャンセル・リセット）を組む
     * @param {Group} parent - 追加先
     * @returns {{btnOK: Button, btnCancel: Button, btnReset: Button}} ボタン
     */
    function buildButtonColumn(parent) {
        var buttonColumn = parent.add("group");
        buttonColumn.orientation = "column";
        buttonColumn.alignChildren = "fill";

        var btnOK = buttonColumn.add("button", undefined, getLabel(LABELS.button.adjust), { name: "ok" });
        btnOK.helpTip = getLabel(LABELS.tooltip.adjust);
        var btnCancel = buttonColumn.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });

        buttonColumn.add("statictext", BUTTON_SPACER_BOUNDS, " "); /* スペーサー / Spacer */
        var btnReset = buttonColumn.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);

        return { btnOK: btnOK, btnCancel: btnCancel, btnReset: btnReset };
    }

    /**
     * 対象文字とシフト量を指定するダイアログを表示する（入力のたびにプレビュー）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {PreviewManager} previewManager - プレビューの取り消し管理
     * @returns {{targetChars: string, shiftAmount: number}|null} 対象文字とシフト量（pt）。キャンセル時は null
     */
    function showShiftDialog(textFrames, previewManager) {
        /* シフト量の欄は環境設定の「東アジア言語」の単位で表示し、適用するときに pt へ換算する
           The shift field uses the East Asian type unit; values are converted to points when applied */
        var shiftUnit = getUnitInfo("text/asianunits");
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "left";
        dialog.opacity = DIALOG_OPACITY;
        dialog.onShow = function () {
            dialog.location = [dialog.location[0] + DIALOG_OFFSET_X, dialog.location[1]];
        };

        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];

        var inputControls = buildInputColumn(columnsGroup, collectSymbolChars(textFrames), shiftUnit.label);
        var dialogButtons = buildButtonColumn(columnsGroup);
        var targetInput = inputControls.targetInput;
        var shiftInput = inputControls.shiftInput;
        var referenceInput = inputControls.referenceInput;

        /* 直前のプレビューを取り消してから、いまの入力で掛け直す / Undo the previous preview, then apply the current input */
        function updatePreview() {
            previewManager.rollback();
            if (!targetInput.text) return;

            var shiftAmount = parseFloat(shiftInput.text);
            if (isNaN(shiftAmount)) shiftAmount = 0;
            previewManager.addStep(function () {
                applyShiftToAll(textFrames, targetInput.text, shiftAmount * shiftUnit.pointsPerUnit);
            });
        }

        targetInput.onChanging = updatePreview;
        shiftInput.onChanging = updatePreview;
        changeValueByArrowKey(shiftInput, updatePreview);

        /* 基準文字との天地中央の差をシフト量に入れる / Put the center difference from the reference character into the shift field */
        inputControls.btnCalculate.onClick = function () {
            if (targetInput.text.length === 0 || referenceInput.text.length !== 1) {
                alert(getLabel(LABELS.alert.invalidChars));
                return;
            }
            var shiftAmount = calculateShiftAmount(textFrames, targetInput.text, referenceInput.text);
            if (shiftAmount === null) {
                alert(getLabel(LABELS.alert.targetNotFound));
                return;
            }

            shiftInput.text = (shiftAmount / shiftUnit.pointsPerUnit).toFixed(SHIFT_DECIMAL_PLACES);
            updatePreview();
        };

        /* ベースラインシフトを全部0に戻すのも、プレビューの1段として扱う / Resetting everything is also a single preview step */
        dialogButtons.btnReset.onClick = function () {
            previewManager.rollback();
            previewManager.addStep(function () {
                resetBaselineShift(textFrames);
            });
            shiftInput.text = "0";
        };

        /* 確定はダイアログを閉じてから main() で行う / The final apply happens in main() after the dialog closes */
        dialogButtons.btnOK.onClick = function () {
            if (targetInput.text.length === 0) {
                alert(getLabel(LABELS.alert.invalidChars));
                return;
            }
            if (isNaN(Number(shiftInput.text))) {
                alert(getLabel(LABELS.alert.shiftNotNumber));
                return;
            }
            dialog.close(1);
        };

        if (dialog.show() !== 1) return null;
        return { targetChars: targetInput.text, shiftAmount: Number(shiftInput.text) * shiftUnit.pointsPerUnit };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前提を確かめてダイアログを表示し、確定したシフト量を適用する
     * @returns {void}
     */
    function main() {
        try {
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }

            var doc = app.activeDocument;
            selectFrameOfEditedText(doc);

            var textFrames = collectTextFrames(doc.selection);
            if (textFrames.length === 0) {
                alert(getLabel(LABELS.alert.selectTextFrame));
                return;
            }

            var previewManager = new PreviewManager();
            var dialogResult = showShiftDialog(textFrames, previewManager);
            if (!dialogResult) {
                /* キャンセルでも閉じるボタンでもプレビューを戻す / Undo the preview on Cancel and on the close box */
                previewManager.rollback();
                return;
            }

            previewManager.confirm(function () {
                applyShiftToAll(textFrames, dialogResult.targetChars, dialogResult.shiftAmount);
            });
        } catch (e) {
            alert(getLabel(LABELS.alert.errorPrefix) + e);
        }
    }

    main();

})();
