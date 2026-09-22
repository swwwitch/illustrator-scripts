#targetengine "RulesBetweenObjects"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクト（図形／テキスト）を上から順に並べ、その間に水平の罫線を描画します。
入力単位は環境設定の「線」に追従し、［延長］で罫線を左右方向に伸縮できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawLinesBetween.md

### Overview

Sorts the selected objects (shapes or text) from top to bottom and draws a horizontal rule between each pair.
Input units follow the stroke-units preference, and Extend stretches or shrinks the rules horizontally.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawLinesBetween.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DrawLinesBetween";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DrawLinesBetween.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawLinesBetween.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /*
      DrawLinesBetween.jsx

      選択したオブジェクト（図形／テキスト）を上から順に並べ、間に水平の罫線（ケイ線）を描画します。
      / Sort selected objects (shapes/text) from top to bottom and draw horizontal rules between them.

      主なポイント / Key features:
      - 入力単位は「線（strokeUnits）」の環境設定に追従（表示ラベル＆内部pt換算）
        / Follows "strokeUnits" preference (UI label + internal pt conversion)
      - 「延長」で左右方向に罫線を伸縮（+で延長、-で短縮）
        / "Extension" expands/contracts the rule length horizontally (+ extends, - shrinks)
      - 線端（なし／丸型／突出）を選択可能
        / Select line cap (Butt / Round / Projecting)
      - ケイ線の長さ：共通（全体幅）／オブジェクトに合わせる（行ペア幅）
        / Rule length: Common (overall width) / Match Objects (per-row-pair width)
      - テキストは複製→アウトライン化した境界で計算し、終了時に一時オブジェクトを削除
        / Text is measured via duplicate→outline bounds, then temporary items are removed
      - 左右に並ぶ要素は同一行として束ねて扱う（行グルーピング）
        / Horizontally aligned items are treated as one row (row grouping)
      - 実行後、作成した罫線を選択状態に
        / Select created rules after execution
      - 最後に使った設定を保存して次回起動時に復元（線幅／延長／線端など）
        / Persist last-used settings and restore on next run (stroke/extension/cap etc.)
    */

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_LINE_WEIGHT = 0.25; /* 線幅の既定値（pt）/ default stroke weight in pt */
    var DEFAULT_EXTENSION   = 0;    /* 延長の既定値（pt）。左右にどれだけ伸ばすか / default extension on each side in pt */
    var ROW_OVERLAP_RATIO   = 0.5;  /* 同じ行とみなす縦方向の重なり率（0.0–1.0）/ vertical overlap ratio for one row */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OFFSET_X    = 300;              /* ダイアログの表示位置のずらし量 / dialog position offset */
    var DIALOG_OFFSET_Y    = 0;
    var DIALOG_OPACITY     = 0.98;             /* ダイアログの不透明度 / dialog opacity */
    var PANEL_MARGINS      = [15, 20, 15, 10]; /* パネル余白 [左,上,右,下] / panel margins */
    var NUMBER_FIELD_CHARS = 6;                /* 数値入力欄の幅（文字数）/ number field width in characters */

    /**
     * ダイアログを表示するときに位置をずらす
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    /**
     * ダイアログの不透明度を設定する
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} opacityValue - 不透明度（0〜1）
     * @returns {void}
     */
    function setDialogOpacity(targetDialog, opacityValue) {
        try {
            targetDialog.opacity = opacityValue;
        } catch (e) { /* 環境によっては opacity を持たない / opacity is not supported in some environments */ }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクト間に罫線", en: "Rules Between Objects" }
        },
        panel: {
            lineCap:    { ja: "線端", en: "Line Cap" },
            ruleLength: { ja: "ケイ線の長さ", en: "Rule Length" }
        },
        fieldLabel: {
            lineWeight: { ja: "線幅", en: "Stroke" },
            extension:  { ja: "延長", en: "Extension" }
        },
        radio: {
            capButt:          { ja: "なし", en: "Butt" },
            capRound:         { ja: "丸型線端", en: "Round" },
            capProjecting:    { ja: "突出線端", en: "Projecting" },
            ruleLengthObject: { ja: "オブジェクトに合わせる", en: "Match Objects" },
            ruleLengthCommon: { ja: "共通", en: "Common" }
        },
        tooltip: {
            lineWeight:       { ja: "ケイ線の太さです。", en: "Weight of the rules." },
            extension:        { ja: "オブジェクトの端からケイ線を離す距離です。", en: "How far the rules sit from the edge of the objects." },
            capButt:          { ja: "線の端を切りっぱなしにします。", en: "Leaves the line ends flat." },
            capRound:         { ja: "線の端を丸くします。", en: "Rounds the line ends." },
            capProjecting:    { ja: "線の端を太さの半分だけ延ばします。", en: "Extends the line ends by half the weight." },
            ruleLengthObject: {
                ja: "ケイ線の長さを、上下のオブジェクトの幅に合わせます。",
                en: "Matches each rule to the width of the objects it sits between."
            },
            ruleLengthCommon: { ja: "すべてのケイ線を同じ長さにそろえます。", en: "Gives every rule the same length." }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            needPositiveStroke:  { ja: "線幅は正の数値を入力してください。", en: "Please enter a positive number for stroke." },
            needNumberExtension: { ja: "延長は数値を入力してください。", en: "Please enter a number for extension." },
            noDocument:          { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文言を取り出す
     * @param {string} labelPath - "category.key" 形式のパス
     * @returns {string} 表示用の文言（見つからないときはパスそのもの）
     */
    function getLabel(labelPath) {
        var labelNode = LABELS;
        var pathKeys = labelPath.split(".");
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
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
    // 前回値の記憶 / Session settings
    // =========================================

    var SETTINGS_KEY = "RulesBetweenObjectsSettings";

    /**
     * 前回の設定を読み込む
     * @returns {{lineWeightPt: number, extensionPt: number, capIndex: number}} 保存値（無いときは null）
     */
    function loadSettings() {
        try {
            /* 保存値が無いと例外になる / throws when nothing is saved */
            var settingsDescriptor = app.getCustomOptions(SETTINGS_KEY);
            return {
                lineWeightPt: settingsDescriptor.getReal(stringIDToTypeID("lineWeightPt")),
                extensionPt: settingsDescriptor.getReal(stringIDToTypeID("marginPt")),
                capIndex: settingsDescriptor.getInteger(stringIDToTypeID("capIndex"))
            };
        } catch (e) {
            return null;
        }
    }

    /**
     * 最後に使った設定を保存する（長さは pt で持つ）
     * @param {number} lineWeightPt - 線幅（pt）
     * @param {number} extensionPt - 延長（pt）
     * @param {number} capIndex - 線端（0:なし / 1:丸型 / 2:突出）
     * @returns {void}
     */
    function saveSettings(lineWeightPt, extensionPt, capIndex) {
        try {
            var settingsDescriptor = new ActionDescriptor();
            settingsDescriptor.putReal(stringIDToTypeID("lineWeightPt"), lineWeightPt);
            settingsDescriptor.putReal(stringIDToTypeID("marginPt"), extensionPt);
            settingsDescriptor.putInteger(stringIDToTypeID("capIndex"), capIndex);
            app.putCustomOptions(SETTINGS_KEY, settingsDescriptor, true);
        } catch (e) { }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値を入力欄向けに丸める（小数第2位まで）
     * @param {number} numberValue - 表示したい数値
     * @returns {string} 整形後の文字列
     */
    function formatNumberForUI(numberValue) {
        var roundedValue = Math.round(numberValue * 100) / 100;
        /* -0 を防ぐ / avoid -0 */
        if (Math.abs(roundedValue) < 0.000001) roundedValue = 0;
        return String(roundedValue);
    }

    /**
     * 入力欄に↑↓キーでの数値増減を設定する（Shift で 10 刻み、Option で 0.1 刻み）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

    /**
     * 「項目名＋数値欄＋（単位）」の行を追加する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {number} initialValue - 初期値（現在の線の単位）
     * @param {string} unitLabel - 単位の表示
     * @param {string} tooltipPath - tooltip のラベルのパス
     * @returns {EditText} 追加した数値欄
     */
    function addNumberRow(parentContainer, labelPath, initialValue, unitLabel, tooltipPath) {
        var numberRow = parentContainer.add("group");
        numberRow.add("statictext", undefined, getLabel(labelPath));
        var numberInput = numberRow.add("edittext", undefined, formatNumberForUI(initialValue));
        numberInput.helpTip = getLabel(tooltipPath);
        numberInput.characters = NUMBER_FIELD_CHARS;
        changeValueByArrowKey(numberInput);
        numberRow.add("statictext", undefined, "(" + unitLabel + ")");
        return numberInput;
    }

    /**
     * ラジオボタンを縦に並べるパネルを追加する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {string} titlePath - パネル見出しのラベルのパス
     * @returns {Group} ラジオボタンを入れるグループ
     */
    function addRadioPanel(parentContainer, titlePath) {
        var radioPanel = parentContainer.add("panel", undefined, getLabel(titlePath));
        radioPanel.orientation = "column";
        radioPanel.alignChildren = ["left", "top"];
        radioPanel.margins = PANEL_MARGINS;

        var radioGroup = radioPanel.add("group");
        radioGroup.orientation = "column";
        radioGroup.alignChildren = ["left", "center"];
        return radioGroup;
    }

    /**
     * tooltip 付きのラジオボタンを追加する
     * @param {Group} parentContainer - 追加先のグループ
     * @param {string} labelKey - radio と tooltip に共通のキー
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadioButton(parentContainer, labelKey) {
        var radioButton = parentContainer.add("radiobutton", undefined, getLabel("radio." + labelKey));
        radioButton.helpTip = getLabel("tooltip." + labelKey);
        return radioButton;
    }

    /**
     * ダイアログを組み立てる
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @param {{lineWeight: number, extension: number, capIndex: number}} initialValues - 初期値（長さは現在の線の単位）
     * @returns {Object} ダイアログと入力の読み取りに使うコントロール
     */
    function buildRulesDialog(strokeUnit, initialValues) {
        var rulesDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setDialogOpacity(rulesDialog, DIALOG_OPACITY);
        shiftDialogPosition(rulesDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        rulesDialog.orientation = "column";
        rulesDialog.alignChildren = ["fill", "top"];

        var lineWeightInput = addNumberRow(rulesDialog, "fieldLabel.lineWeight", initialValues.lineWeight,
            strokeUnit.label, "tooltip.lineWeight");
        var extensionInput = addNumberRow(rulesDialog, "fieldLabel.extension", initialValues.extension,
            strokeUnit.label, "tooltip.extension");

        /* 線端（既定は「なし」、前回値があれば反映）/ Line cap: Butt by default, or the saved one */
        var capRadioGroup = addRadioPanel(rulesDialog, "panel.lineCap");
        var capButtRadio = addRadioButton(capRadioGroup, "capButt");
        var capRoundRadio = addRadioButton(capRadioGroup, "capRound");
        var capProjectingRadio = addRadioButton(capRadioGroup, "capProjecting");
        capButtRadio.value = true;
        if (initialValues.capIndex === 1) capRoundRadio.value = true;
        else if (initialValues.capIndex === 2) capProjectingRadio.value = true;

        /* ケイ線の長さ（既定は「共通」）/ Rule length: Common by default */
        var lengthRadioGroup = addRadioPanel(rulesDialog, "panel.ruleLength");
        var lengthObjectRadio = addRadioButton(lengthRadioGroup, "ruleLengthObject");
        var lengthCommonRadio = addRadioButton(lengthRadioGroup, "ruleLengthCommon");
        lengthCommonRadio.value = true;

        var btnRowGroup = rulesDialog.add("group");
        btnRowGroup.alignment = ["right", "center"];
        btnRowGroup.orientation = "row";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnOK.active = true;              /* Enterキー＝OK / Enter = OK */
        rulesDialog.defaultElement = btnOK; /* 環境依存対策 / for environments that ignore active */

        return {
            dialog: rulesDialog,
            lineWeightInput: lineWeightInput,
            extensionInput: extensionInput,
            capRoundRadio: capRoundRadio,
            capProjectingRadio: capProjectingRadio,
            lengthObjectRadio: lengthObjectRadio
        };
    }

    /**
     * ダイアログの入力を検証して設定にまとめる
     * @param {Object} dialogControls - buildRulesDialog() の戻り値
     * @param {Object} strokeUnit - 線の単位（getUnitInfo() の戻り値）
     * @returns {Object} 線幅・延長（pt）・線端・長さの基準（入力が不正なときは null）
     */
    function readRuleSettings(dialogControls, strokeUnit) {
        var lineWeightValue = Number(dialogControls.lineWeightInput.text);
        if (isNaN(lineWeightValue) || lineWeightValue <= 0) {
            alert(getLabel("alert.needPositiveStroke"));
            return null;
        }

        var extensionValue = Number(dialogControls.extensionInput.text);
        if (isNaN(extensionValue)) {
            alert(getLabel("alert.needNumberExtension"));
            return null;
        }

        var capIndex = 0;
        if (dialogControls.capRoundRadio.value) capIndex = 1;
        else if (dialogControls.capProjectingRadio.value) capIndex = 2;

        return {
            lineWeight: lineWeightValue * strokeUnit.pointsPerUnit,
            extension: extensionValue * strokeUnit.pointsPerUnit,
            capIndex: capIndex,
            lineCap: [StrokeCap.BUTTENDCAP, StrokeCap.ROUNDENDCAP, StrokeCap.PROJECTINGENDCAP][capIndex],
            ruleLengthMode: dialogControls.lengthObjectRadio.value ? "object" : "common"
        };
    }

    /**
     * ダイアログを表示し、確定した設定を保存して返す
     * @returns {Object} readRuleSettings() の戻り値（キャンセル・入力エラーのときは null）
     */
    function showRulesDialog() {
        var savedSettings = loadSettings();
        var strokeUnit = getUnitInfo("strokeUnits");

        /* 保存値（pt）があれば優先し、現在の線の単位で表示する / Show saved pt values in the current stroke unit */
        var initialValues = {
            lineWeight: (savedSettings ? savedSettings.lineWeightPt : DEFAULT_LINE_WEIGHT) / strokeUnit.pointsPerUnit,
            extension: (savedSettings ? savedSettings.extensionPt : DEFAULT_EXTENSION) / strokeUnit.pointsPerUnit,
            capIndex: savedSettings ? savedSettings.capIndex : 0
        };

        var dialogControls = buildRulesDialog(strokeUnit, initialValues);
        if (dialogControls.dialog.show() !== 1) return null;

        var ruleSettings = readRuleSettings(dialogControls, strokeUnit);
        if (ruleSettings) saveSettings(ruleSettings.lineWeight, ruleSettings.extension, ruleSettings.capIndex);
        return ruleSettings;
    }

    // =========================================
    // 行の検出 / Row detection
    // =========================================

    /**
     * グループ内のテキストをアウトライン化する（境界を安定させるため。複製に対して使う）
     * @param {GroupItem} targetGroup - 対象のグループ
     * @returns {void}
     */
    function outlineTextFramesInContainer(targetGroup) {
        if (!targetGroup || !targetGroup.textFrames || targetGroup.textFrames.length === 0) return;

        /* textFrames はライブコレクションになり得るので、いったん配列化 / snapshot the live collection */
        var textFrameList = [];
        for (var i = 0; i < targetGroup.textFrames.length; i++) {
            textFrameList.push(targetGroup.textFrames[i]);
        }

        for (var j = 0; j < textFrameList.length; j++) {
            try {
                /* createOutline() は元のテキストを消費してアウトラインのグループを返す / consumes the text frame */
                var outlineGroup = textFrameList[j].createOutline();
                outlineGroup.hidden = true;
            } catch (e) { /* 変換できないテキストは無視 / skip frames that cannot be outlined */ }
        }
    }

    /**
     * 境界の計測に使うオブジェクトを返す
     * テキストとグループは複製してアウトライン化したもの（非表示）を返し、tempItems に控える。
     * @param {PageItem} sourceItem - 選択中のオブジェクト
     * @param {PageItem[]} tempItems - 一時オブジェクトの控え（追加される）
     * @returns {PageItem} 計測用のオブジェクト（複製できないときは元のオブジェクト）
     */
    function makeBoundsProxy(sourceItem, tempItems) {
        if (!sourceItem) return sourceItem;

        if (sourceItem.typename === "TextFrame") {
            try {
                /* createOutline() は複製を消費してアウトラインのグループを返す / consumes the duplicate */
                var outlineGroup = sourceItem.duplicate().createOutline();
                try { outlineGroup.hidden = true; } catch (e) { }
                tempItems.push(outlineGroup);
                return outlineGroup;
            } catch (e) {
                /* 変換できない場合は元のテキストを使う / fall back to the original text */
                return sourceItem;
            }
        }

        if (sourceItem.typename === "GroupItem") {
            try {
                var groupDuplicate = sourceItem.duplicate();
                /* グループ内のテキストもアウトライン化してから境界を見る / outline nested text too */
                outlineTextFramesInContainer(groupDuplicate);
                try { groupDuplicate.hidden = true; } catch (e) { }
                tempItems.push(groupDuplicate);
                return groupDuplicate;
            } catch (e) {
                return sourceItem;
            }
        }

        return sourceItem;
    }

    /**
     * 2つのオブジェクトの縦方向の重なり率を返す（低いほうの高さに対する比率）
     * @param {Object} itemA - visibleBounds を持つオブジェクト
     * @param {Object} itemB - visibleBounds を持つオブジェクト
     * @returns {number} 重なり率（重ならないときは 0）
     */
    function getVerticalOverlapRatio(itemA, itemB) {
        /* visibleBounds: [left, top, right, bottom] */
        var boundsA = itemA.visibleBounds;
        var boundsB = itemB.visibleBounds;
        var overlap = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
        if (overlap <= 0) return 0;

        var minHeight = Math.min(boundsA[1] - boundsA[3], boundsB[1] - boundsB[3]);
        return overlap / minHeight;
    }

    /**
     * 複数のオブジェクトを囲む境界を返す
     * @param {Object[]} boundsItems - visibleBounds を持つオブジェクト
     * @returns {{visibleBounds: number[]}} まとめた境界
     */
    function mergeBounds(boundsItems) {
        var firstBounds = boundsItems[0].visibleBounds;
        var left = firstBounds[0];
        var top = firstBounds[1];
        var right = firstBounds[2];
        var bottom = firstBounds[3];

        for (var i = 1; i < boundsItems.length; i++) {
            var itemBounds = boundsItems[i].visibleBounds;
            left = Math.min(left, itemBounds[0]);
            top = Math.max(top, itemBounds[1]);
            right = Math.max(right, itemBounds[2]);
            bottom = Math.min(bottom, itemBounds[3]);
        }

        return { visibleBounds: [left, top, right, bottom] };
    }

    /**
     * 左右に並ぶオブジェクトを同じ行として束ね、行ごとの境界を返す
     * @param {Object[]} boundsItems - visibleBounds を持つオブジェクト
     * @returns {Object[]} 行ごとの境界（{visibleBounds}）
     */
    function groupItemsByRow(boundsItems) {
        var rowMembers = [];

        for (var i = 0; i < boundsItems.length; i++) {
            var isPlaced = false;
            for (var r = 0; r < rowMembers.length; r++) {
                /* 行の代表要素（最初の1つ）と縦方向の重なりを比較 / compare with the first item of the row */
                if (getVerticalOverlapRatio(boundsItems[i], rowMembers[r][0]) >= ROW_OVERLAP_RATIO) {
                    rowMembers[r].push(boundsItems[i]);
                    isPlaced = true;
                    break;
                }
            }
            if (!isPlaced) rowMembers.push([boundsItems[i]]);
        }

        var rowBoundsList = [];
        for (var j = 0; j < rowMembers.length; j++) {
            rowBoundsList.push(mergeBounds(rowMembers[j]));
        }
        return rowBoundsList;
    }

    /**
     * 選択を行にまとめ、上から順に並べた行の境界を返す
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @param {PageItem[]} tempItems - 計測用に作った一時オブジェクトの控え（追加される）
     * @returns {Object[]} 上から順の行の境界（{visibleBounds}）
     */
    function collectSortedRows(selectedItems, tempItems) {
        var proxyItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            proxyItems.push(makeBoundsProxy(selectedItems[i], tempItems));
        }

        var rowBoundsList = groupItemsByRow(proxyItems);
        rowBoundsList.sort(function (rowA, rowB) {
            /* top が大きいほうが上（一般的な定規の設定を想定）/ larger top = higher */
            return rowB.visibleBounds[1] - rowA.visibleBounds[1];
        });
        return rowBoundsList;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 複数の境界の左右の端を返す
     * @param {Object[]} boundsItems - visibleBounds を持つオブジェクト
     * @returns {{left: number, right: number}} 左端と右端
     */
    function getHorizontalSpan(boundsItems) {
        var left = boundsItems[0].visibleBounds[0];
        var right = boundsItems[0].visibleBounds[2];
        for (var i = 1; i < boundsItems.length; i++) {
            left = Math.min(left, boundsItems[i].visibleBounds[0]);
            right = Math.max(right, boundsItems[i].visibleBounds[2]);
        }
        return { left: left, right: right };
    }

    /**
     * 上下に並ぶ行の間にケイ線を描く
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} rowBoundsList - 上から順の行の境界
     * @param {Object} ruleSettings - showRulesDialog() の戻り値
     * @returns {PathItem[]} 描いたケイ線
     */
    function drawRulesBetweenRows(doc, rowBoundsList, ruleSettings) {
        /* 線の色（K=100）/ Line color: black K=100 */
        var lineColor = new CMYKColor();
        lineColor.cyan = 0;
        lineColor.magenta = 0;
        lineColor.yellow = 0;
        lineColor.black = 100;

        /* 延長：0 ならオブジェクトと同じ幅、正の数なら長く、負の数なら短く / 0 = same width, + longer, - shorter */
        var extension = ruleSettings.extension;

        /* 「共通」は全体の左右幅、「オブジェクトに合わせる」は上下の行の左右幅 / common span vs per-pair span */
        var commonSpan = (ruleSettings.ruleLengthMode === "common" && rowBoundsList.length > 1)
            ? getHorizontalSpan(rowBoundsList)
            : null;

        var createdLines = [];
        for (var i = 0; i < rowBoundsList.length - 1; i++) {
            var upperRow = rowBoundsList[i];
            var lowerRow = rowBoundsList[i + 1];

            /* 上の行の下端と下の行の上端の中間 / midway between the upper bottom and the lower top */
            var midY = (upperRow.visibleBounds[3] + lowerRow.visibleBounds[1]) / 2;
            var ruleSpan = commonSpan || getHorizontalSpan([upperRow, lowerRow]);

            var ruleLine = doc.pathItems.add();
            ruleLine.setEntirePath([[ruleSpan.left - extension, midY], [ruleSpan.right + extension, midY]]);
            ruleLine.filled = false;
            ruleLine.stroked = true;
            ruleLine.strokeWidth = ruleSettings.lineWeight;
            ruleLine.strokeColor = lineColor;
            ruleLine.strokeCap = ruleSettings.lineCap;
            createdLines.push(ruleLine);
        }
        return createdLines;
    }

    /**
     * 計測用に作った一時オブジェクトを削除する
     * @param {PageItem[]} tempItems - 一時オブジェクト
     * @returns {void}
     */
    function removeTemporaryItems(tempItems) {
        for (var i = 0; i < tempItems.length; i++) {
            try {
                if (tempItems[i] && tempItems[i].typename) tempItems[i].remove();
            } catch (e) { }
        }
    }

    /**
     * 描いたケイ線を選択する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} createdLines - 描いたケイ線
     * @returns {void}
     */
    function selectCreatedLines(doc, createdLines) {
        if (createdLines.length === 0) return;
        doc.selection = null;
        for (var i = 0; i < createdLines.length; i++) {
            try {
                createdLines[i].selected = true;
            } catch (e) { }
        }
    }

    /**
     * ダイアログで設定を受け取り、選択したオブジェクトの間にケイ線を描く
     * @returns {void}
     */
    function main() {
        var ruleSettings = showRulesDialog();
        if (ruleSettings === null) return;

        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        /* テキストとグループは複製→アウトライン化した境界で計算し、最後に削除する / measured on outlined duplicates */
        var tempItems = [];
        var rowBoundsList = collectSortedRows(doc.selection, tempItems);
        var createdLines = drawRulesBetweenRows(doc, rowBoundsList, ruleSettings);

        removeTemporaryItems(tempItems);
        selectCreatedLines(doc, createdLines);
    }

    main();

})();
