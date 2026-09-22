#targetengine "TypeScalerEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストのフォントサイズを、基準サイズと比率から算出して適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypeScaler.md

### Overview

Sets the font size of the selected text from a base size and a ratio.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypeScaler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TypeScaler";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypeScaler.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypeScaler.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_BASE_SIZE = "12";          /* 基準サイズの初期値（前回値が無いとき）/ default base size */
    var DEFAULT_RATIO_INDEX = 3;           /* 倍率の初期値（Major Third）/ default ratio */
    var FALLBACK_RATIO = 1.25;             /* 倍率が選ばれていないときの倍率 / ratio used when none is selected */
    var SAMPLE_START_POSITION = [20, -20]; /* 見本の1行目の位置 [left, top] / position of the first sample line */
    var SAMPLE_LINE_GAP = 20;              /* 見本の行ごとにサイズへ足す間隔 / gap added below each sample line */

    /* 倍率の候補 / Ratio choices */
    var TYPE_SCALE_RATIOS = [
        { label: "Minor Second 1.067", value: 1.067 },
        { label: "Major Second 1.125", value: 1.125 },
        { label: "Minor Third 1.2", value: 1.2 },
        { label: "Major Third 1.25", value: 1.25 },
        { label: "Golden Ratio: ½ 1.309", value: 1.309 },
        { label: "Perfect Fourth 1.333", value: 1.333 },
        { label: "Augmented Fourth 1.414", value: 1.414 },
        { label: "Golden Ratio 1.618", value: 1.618 }
    ];

    // =========================================
    // レイアウト / Layout
    // =========================================
    var SIZE_ROW_MARGINS = [0, 0, 0, 15];  /* 基準サイズの行の余白 / margins of the base size row */
    var PANEL_MARGINS = [15, 20, 15, 10];  /* パネル余白 / panel margins */
    var SIZE_FIELD_CHARS = 4;              /* 基準サイズ欄の文字数 / base size field width */
    var SAMPLE_FIELD_CHARS = 20;           /* 見本の文字列欄の文字数 / sample text field width */
    var SIZE_LIST_SIZE = [85, 136];        /* サイズ一覧の大きさ / size list size */
    var DIALOG_OPACITY = 0.97;             /* ダイアログの不透明度 / dialog opacity */

    // =========================================
    // セッション / Session
    // =========================================
    /* 見本作成で使った基準サイズ / Base size used by the last sample */
    var persistentBaseSize = null;
    $.global.__sizeValue = $.global.__sizeValue || DEFAULT_BASE_SIZE;
    $.global.__ratioIndex = ($.global.__ratioIndex !== undefined) ? $.global.__ratioIndex : DEFAULT_RATIO_INDEX;

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
            title: { ja: "タイプスケール", en: "Type Scale" }
        },
        panel: {
            sample: { ja: "見本作成", en: "Create Sample" }
        },
        fieldLabel: {
            base: { ja: "基準", en: "Base" },
            unitSuffix: { ja: "（{unit}）", en: " ({unit})" }
        },
        defaultValue: {
            sampleText: { ja: "山路を登りながら", en: "Sample Text" }
        },
        checkbox: {
            showSize: { ja: "サイズ表示", en: "Show Size" }
        },
        button: {
            createSample: { ja: "見本作成", en: "Create" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            baseSize: { ja: "タイプスケールの基準になるフォントサイズです。", en: "Font size the type scale is built from." },
            ratio: {
                ja: "1段ごとに掛ける倍率です。大きいほどサイズの差が開きます。",
                en: "Multiplier applied at each step. A larger ratio spreads the sizes further apart."
            },
            sizeList: {
                ja: "計算したサイズの一覧です。選んで OK すると、選択中のテキストに適用します。",
                en: "The calculated sizes. Pick one and press OK to apply it to the selected text."
            },
            sampleText: { ja: "見本に使う文字列です。", en: "Text used for the sample." },
            showSize: { ja: "見本の各行にサイズの数値を添えます。", en: "Adds the size value to each line of the sample." },
            createSample: {
                ja: "一覧のすべてのサイズで見本を作り、ドキュメントに配置します。",
                en: "Creates a sample in every size on the list and places it in the document."
            }
        },
        alert: {
            selectSize: { ja: "リストからサイズを選択してください。", en: "Please select a size from the list." },
            invalidSize: { ja: "正しいサイズを選択してください。", en: "Please select a valid size." },
            invalidBase: { ja: "基準フォントサイズが不正です。", en: "Invalid base font size." },
            applyError: { ja: "フォントサイズの適用に失敗しました：", en: "Failed to apply font size: " },
            fontError: { ja: "フォントの適用に失敗しました：", en: "Failed to apply font: " }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文言を取り出す
     * @param {string} labelPath - "dialog.title" のようなパス
     * @returns {string} 表示言語の文言
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
        }
        return labelNode[uiLang];
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
    // タイプスケール / Type scale
    // =========================================

    /**
     * 基準サイズの2段下から4段上まで、7段のサイズを求める
     * @param {number} baseSize - 基準サイズ
     * @param {number} ratio - 1段ごとの倍率
     * @returns {number[]} 小数第1位に丸めたサイズ（小さい順）
     */
    function generateTypeScaleSizes(baseSize, ratio) {
        var typeSizes = [];
        var stepSize = baseSize;
        for (var i = 0; i < 2; i++) {
            stepSize /= ratio;
        }
        for (var j = 0; j < 7; j++) {
            typeSizes.push(Math.round(stepSize * 10) / 10);
            stepSize *= ratio;
        }
        return typeSizes;
    }

    /**
     * 倍率ドロップダウンの表示名の一覧を返す
     * @returns {string[]} 表示名
     */
    function getRatioLabels() {
        var ratioLabels = [];
        for (var i = 0; i < TYPE_SCALE_RATIOS.length; i++) {
            ratioLabels.push(TYPE_SCALE_RATIOS[i].label);
        }
        return ratioLabels;
    }

    /**
     * 選ばれている倍率を返す
     * @param {DropDownList} ratioPopup - 倍率のドロップダウン
     * @returns {number} 倍率（未選択なら FALLBACK_RATIO）
     */
    function getSelectedRatio(ratioPopup) {
        return ratioPopup.selection ? TYPE_SCALE_RATIOS[ratioPopup.selection.index].value : FALLBACK_RATIO;
    }

    /**
     * サイズ一覧を計算し直す
     * @param {ListBox} sizeList - サイズ一覧
     * @param {DropDownList} ratioPopup - 倍率のドロップダウン
     * @param {number} baseSize - 基準サイズ
     * @param {string} textUnitLabel - 文字サイズの単位
     * @returns {void}
     */
    function updateSizeList(sizeList, ratioPopup, baseSize, textUnitLabel) {
        sizeList.removeAll();
        if (isNaN(baseSize) || baseSize <= 0) return;
        var typeSizes = generateTypeScaleSizes(baseSize, getSelectedRatio(ratioPopup));
        for (var i = 0; i < typeSizes.length; i++) {
            sizeList.add("item", typeSizes[i] + " " + textUnitLabel);
        }
    }

    /**
     * 一覧の項目（"12 pt" など）からサイズの数値を読む
     * @param {string} listItemText - 一覧の項目の文字列
     * @returns {number} サイズ（読めなければ NaN）
     */
    function parseSizeFromListText(listItemText) {
        var textParts = listItemText.split(" ");
        for (var i = 0; i < textParts.length; i++) {
            if (textParts[i].indexOf("pt") !== -1 || !isNaN(parseFloat(textParts[i]))) {
                return parseFloat(textParts[i]);
            }
        }
        return NaN;
    }

    /**
     * 基準サイズと倍率の選択をセッションに残す
     * @param {EditText} sizeInput - 基準サイズの入力欄
     * @param {DropDownList} ratioPopup - 倍率のドロップダウン
     * @returns {void}
     */
    function storeSessionValues(sizeInput, ratioPopup) {
        $.global.__sizeValue = sizeInput.text;
        $.global.__ratioIndex = ratioPopup.selection ? ratioPopup.selection.index : 0;
    }

    // =========================================
    // 選択テキスト / Selected text
    // =========================================

    /**
     * 選択にテキストフレームがあるか判定する
     * @returns {boolean} 1つでもあれば true
     */
    function hasSelectedTextFrame() {
        var docSelection = app.activeDocument.selection;
        for (var i = 0; i < docSelection.length; i++) {
            if (docSelection[i].typename === "TextFrame") return true;
        }
        return false;
    }

    /**
     * 文字編集中で、そのストーリーが1つのテキストフレームだけなら、フレームを選択して選択ツールにする
     * @returns {void}
     */
    function selectFrameOfEditedText() {
        if (app.selection.constructor.name !== "TextRange") return;
        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length !== 1) return;
        app.executeMenuCommand("deselectall");
        app.selection = [textFramesInStory[0]];
        /* ツールの切り替えは失敗しても続行 / switching tools is best effort */
        try {
            app.selectTool("Adobe Select Tool");
        } catch (e) {}
    }

    /**
     * 選択中のテキストにフォントサイズを適用する
     * @param {number} sizeValue - フォントサイズ
     * @returns {void}
     */
    function applySizeToSelection(sizeValue) {
        var docSelection = app.activeDocument.selection;
        for (var i = 0; i < docSelection.length; i++) {
            var selectedItem = docSelection[i];
            /* 文字属性を書き込めないテキストがある / some text rejects the size */
            try {
                if (selectedItem.typename === "TextRange") {
                    selectedItem.characterAttributes.size = sizeValue;
                } else if (selectedItem.typename === "TextFrame") {
                    if (selectedItem.textRange && selectedItem.textRange.characters.length > 0) {
                        selectedItem.textRange.characterAttributes.size = sizeValue;
                    }
                }
            } catch (e) {
                alert(getLabel("alert.applyError") + e.message);
            }
        }
    }

    /**
     * 選択中の最初のテキストフレームのフォントを返す
     * @returns {TextFont|null} フォント（見つからなければ null）
     */
    function getFirstSelectedFont() {
        var docSelection = app.activeDocument.selection;
        for (var i = 0; i < docSelection.length; i++) {
            if (docSelection[i].typename === "TextFrame") {
                /* 読めないものは飛ばして次を見る / skip frames whose font cannot be read */
                try {
                    return docSelection[i].textRange.characterAttributes.textFont;
                } catch (e) {}
            }
        }
        return null;
    }

    /**
     * タイプスケールの各サイズで見本のテキストを縦に並べて作る
     * @param {number} baseSize - 基準サイズ
     * @param {number} ratio - 1段ごとの倍率
     * @param {string} sampleText - 見本の文字列
     * @param {boolean} showSize - 各行にサイズを添えるなら true
     * @param {string} textUnitLabel - 文字サイズの単位
     * @param {TextFont|null} sampleFont - 見本に使うフォント（null ならそのまま）
     * @returns {void}
     */
    function createSampleTexts(baseSize, ratio, sampleText, showSize, textUnitLabel, sampleFont) {
        var doc = app.activeDocument;
        var sampleTop = SAMPLE_START_POSITION[1];
        var typeSizes = generateTypeScaleSizes(baseSize, ratio);
        for (var i = 0; i < typeSizes.length; i++) {
            var fontSize = typeSizes[i];
            var sampleFrame = doc.textFrames.add();
            var contentText = sampleText;
            if (showSize) {
                contentText += "（" + fontSize + textUnitLabel + "）";
            }
            sampleFrame.contents = contentText;
            sampleFrame.left = SAMPLE_START_POSITION[0];
            sampleFrame.top = sampleTop;

            if (sampleFont) {
                /* 環境にないフォントは適用できない / a missing font cannot be applied */
                try {
                    sampleFrame.textRange.characterAttributes.textFont = sampleFont;
                } catch (e) {
                    alert(getLabel("alert.fontError") + e.message);
                }
            }

            /* 範囲外のサイズは受け付けない / out-of-range sizes are rejected */
            try {
                sampleFrame.textRange.characterAttributes.size = fontSize;
            } catch (e) {
                alert(getLabel("alert.applyError") + e.message);
            }

            sampleTop -= fontSize + SAMPLE_LINE_GAP;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 上下キーで値を変更する（Shift で10刻み、Option で0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboardState.shiftKey) {
                delta = 10;
                if (event.keyName == "Up") {
                    currentValue = Math.ceil((currentValue + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    currentValue = Math.floor((currentValue - 1) / delta) * delta;
                    if (currentValue < 0) currentValue = 0;
                    event.preventDefault();
                }
            } else if (keyboardState.altKey) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    currentValue += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    currentValue -= delta;
                    event.preventDefault();
                }
            } else {
                if (event.keyName == "Up") {
                    currentValue += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    currentValue -= delta;
                    if (currentValue < 0) currentValue = 0;
                    event.preventDefault();
                }
            }

            if (keyboardState.altKey) {
                currentValue = Math.round(currentValue * 10) / 10;
            } else {
                currentValue = Math.round(currentValue);
            }

            editText.text = currentValue;
        });
    }

    /**
     * ダイアログを組み立てる（イベントの配線は main() で行う）
     * @param {string} textUnitLabel - 文字サイズの単位
     * @returns {Object} 作成したダイアログとコントロール
     */
    function buildTypeScaleDialog(textUnitLabel) {
        var typeScaleDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        typeScaleDialog.orientation = "column";
        typeScaleDialog.alignChildren = "left";

        /* 基準サイズと倍率 / Base size and ratio */
        var sizeRow = typeScaleDialog.add("group");
        sizeRow.orientation = "row";
        sizeRow.margins = SIZE_ROW_MARGINS;
        sizeRow.spacing = 5;
        sizeRow.add("statictext", undefined, labelText("fieldLabel.base"));
        var sizeInput = sizeRow.add("edittext", undefined, $.global.__sizeValue);
        sizeInput.characters = SIZE_FIELD_CHARS;
        sizeInput.helpTip = getLabel("tooltip.baseSize");
        sizeRow.add("statictext", undefined, getLabel("fieldLabel.unitSuffix").replace("{unit}", textUnitLabel));
        changeValueByArrowKey(sizeInput);

        var ratioPopup = sizeRow.add("dropdownlist", undefined, getRatioLabels());
        ratioPopup.selection = $.global.__ratioIndex;
        ratioPopup.helpTip = getLabel("tooltip.ratio");

        sizeRow.alignment = "center";

        var columnsGroup = typeScaleDialog.add("group");
        columnsGroup.orientation = "row";

        /* 左カラム：サイズ一覧 / Left column: size list */
        var listColumn = columnsGroup.add("group");
        listColumn.orientation = "column";
        listColumn.alignChildren = "left";

        var sizeList = listColumn.add("listbox", undefined, [], { multiselect: false });
        sizeList.preferredSize = SIZE_LIST_SIZE;
        sizeList.helpTip = getLabel("tooltip.sizeList");

        /* 右カラム：見本作成 / Right column: sample */
        var sampleColumn = columnsGroup.add("group");
        sampleColumn.orientation = "column";
        sampleColumn.alignment = "top";
        sampleColumn.alignChildren = "left";

        var samplePanel = sampleColumn.add("panel", undefined, getLabel("panel.sample"));
        samplePanel.orientation = "column";
        samplePanel.alignChildren = "left";
        samplePanel.margins = PANEL_MARGINS;

        var sampleInput = samplePanel.add("edittext", undefined, getLabel("defaultValue.sampleText"));
        sampleInput.characters = SAMPLE_FIELD_CHARS;
        sampleInput.helpTip = getLabel("tooltip.sampleText");
        var showSizeCheckbox = samplePanel.add("checkbox", undefined, getLabel("checkbox.showSize"));
        showSizeCheckbox.value = true;
        showSizeCheckbox.helpTip = getLabel("tooltip.showSize");
        var btnCreateSample = samplePanel.add("button", undefined, getLabel("button.createSample"));
        btnCreateSample.alignment = "right";
        btnCreateSample.helpTip = getLabel("tooltip.createSample");

        /* ボタンエリア / Button area */
        var btnRowGroup = typeScaleDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));

        return {
            typeScaleDialog: typeScaleDialog,
            sizeInput: sizeInput,
            ratioPopup: ratioPopup,
            sizeList: sizeList,
            sampleInput: sampleInput,
            showSizeCheckbox: showSizeCheckbox,
            btnCreateSample: btnCreateSample,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、選んだサイズの適用または見本の作成を行う
     * @returns {void}
     */
    function main() {
        var textUnitLabel = getUnitInfo("text/units").label;
        var dialogControls = buildTypeScaleDialog(textUnitLabel);
        var typeScaleDialog = dialogControls.typeScaleDialog;
        var sizeInput = dialogControls.sizeInput;
        var ratioPopup = dialogControls.ratioPopup;
        var sizeList = dialogControls.sizeList;

        /* 基準サイズ入力変更時 / Base size edited */
        sizeInput.onChanging = function () {
            var inputValue = parseFloat(sizeInput.text);
            if (!isNaN(inputValue) && inputValue > 0) {
                persistentBaseSize = inputValue;
                $.global.__sizeInput = sizeInput;
            }
            $.global.__sizeValue = sizeInput.text;
            updateSizeList(sizeList, ratioPopup, inputValue, textUnitLabel);
        };

        /* 倍率変更時 / Ratio changed */
        ratioPopup.onChange = function () {
            updateSizeList(sizeList, ratioPopup, parseFloat(sizeInput.text), textUnitLabel);
            $.global.__ratioPopup = ratioPopup;
            $.global.__ratioIndex = ratioPopup.selection ? ratioPopup.selection.index : 0;
        };

        /* OK：選択中のテキストにサイズを適用 / OK: apply the size to the selected text */
        dialogControls.btnOK.onClick = function () {
            if (!sizeList.selection) {
                alert(getLabel("alert.selectSize"));
                return;
            }
            var sizeValue = parseSizeFromListText(sizeList.selection.text);
            if (isNaN(sizeValue) || sizeValue <= 0) {
                alert(getLabel("alert.invalidSize"));
                return;
            }

            applySizeToSelection(sizeValue);
            app.redraw();
            storeSessionValues(sizeInput, ratioPopup);
            typeScaleDialog.close();
        };

        dialogControls.btnCancel.onClick = function () {
            typeScaleDialog.close();
        };

        /* 見本作成：一覧で選んだサイズ（無ければ基準サイズ）を基準に見本を作る
           Create sample: build the samples around the selected size, or the base size */
        dialogControls.btnCreateSample.onClick = function () {
            var baseSize;
            if (sizeList.selection) {
                baseSize = parseFloat(sizeList.selection.text.split(" ")[0]);
            } else if (persistentBaseSize !== null) {
                baseSize = persistentBaseSize;
            } else {
                baseSize = parseFloat(sizeInput.text);
            }
            persistentBaseSize = baseSize;
            storeSessionValues(sizeInput, ratioPopup);
            if (isNaN(baseSize) || baseSize <= 0) {
                alert(getLabel("alert.invalidBase"));
                return;
            }

            createSampleTexts(baseSize, getSelectedRatio(ratioPopup), dialogControls.sampleInput.text,
                dialogControls.showSizeCheckbox.value, textUnitLabel, getFirstSelectedFont());
            typeScaleDialog.close();
        };

        /* 文字編集中ならテキストフレームの選択に切り替える / Switch from text editing to the frame */
        selectFrameOfEditedText();

        updateSizeList(sizeList, ratioPopup, parseFloat(sizeInput.text), textUnitLabel);
        /* 選択テキストがない場合は一覧を無効化 / Disable the list when no text is selected */
        if (!hasSelectedTextFrame()) {
            sizeList.enabled = false;
        }

        typeScaleDialog.opacity = DIALOG_OPACITY;
        typeScaleDialog.center();
        typeScaleDialog.onShow = function () {
            sizeInput.active = true;
        };
        typeScaleDialog.show();
    }

    if (app.documents.length > 0) main();

})();
