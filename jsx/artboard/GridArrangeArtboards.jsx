#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード名を解析し、行列グリッドとして再配置します。
名前全体が「行番号＋区切り文字（-, _, x）＋列番号」に一致するアートボードを、行列指定として扱います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridArrangeArtboards.md

### Overview

Parses the artboard names and re-lays the artboards out as a row-column grid.
An artboard counts as positioned when its whole name matches "row number + separator (-, _, x) + column number".

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridArrangeArtboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GridArrangeArtboards";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridArrangeArtboards.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridArrangeArtboards.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値（間隔はアクティブなアートボード幅の1/8） / Dialog defaults (the gap is 1/8 of the active artboard width) */
    var DIALOG_DEFAULTS = {
        exceptionMode: 'rowEnd',            // 'rowEnd'（直前の行の末尾）| 'lastRow'（最後の行にまとめる）
        linkSpacing: true,
        excludeLockedHiddenLayers: false,
        excludeLockedHiddenItems: false,
        changeArtboardOrder: false
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;                    /* ダイアログ外周の余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネルの余白 / panel margins */
    var PANEL_SPACING = 6;                      /* パネル内の間隔 / spacing inside panels */
    var SPACING_COLUMN_GAP = 20;                /* 間隔欄と「連動」の列の間 / gap between the inputs and the link column */
    var SPACING_FIELD_CHARACTERS = 5;           /* 間隔欄の桁数 / width of the gap fields */

    /**
     * パネルを共通スタイルで初期化する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は変更しない）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードを再配置（行列）", en: "Rearrange Artboards (Grid)" }
        },
        panel: {
            spacing: { ja: "間隔", en: "Spacing" },
            exclusion: { ja: "ロック／非表示を除外", en: "Exclude Locked / Hidden" },
            exception: { ja: "未指定／重複", en: "Unmatched / Duplicate" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            spacingX: { ja: "列間", en: "Column gap" },
            spacingY: { ja: "行間", en: "Row gap" }
        },
        checkbox: {
            linkSpacing: { ja: "連動", en: "Link" },
            excludeLockedHiddenLayers: { ja: "レイヤー", en: "Layers" },
            excludeLockedHiddenItems: { ja: "オブジェクト", en: "Objects" },
            changeArtboardOrder: { ja: "パネル上の並び順を変更", en: "Reorder in Artboards panel" }
        },
        radio: {
            exceptionRowEnd: { ja: "直前の行の末尾に配置", en: "Append to current row" },
            exceptionLastRow: { ja: "最後の行にまとめる", en: "Collect in final row" }
        },
        tooltip: {
            spacingX: { ja: "左右に隣り合うアートボードのあいだにあける間隔です。", en: "Gap left between horizontally adjacent artboards." },
            spacingY: { ja: "上下に隣り合うアートボードのあいだにあける間隔です。", en: "Gap left between vertically adjacent artboards." },
            linkSpacing: { ja: "列間と同じ値を行間にも使います。", en: "Uses the column gap for the row gap too." },
            excludeLayers: {
                ja: "ロックまたは非表示のレイヤーに載っているオブジェクトは動かしません。",
                en: "Leaves objects on locked or hidden layers where they are."
            },
            excludeItems: {
                ja: "ロックまたは非表示のオブジェクトは動かしません。",
                en: "Leaves locked or hidden objects where they are."
            },
            exceptionRowEnd: {
                ja: "アートボード名から行列を読み取れなかったものを、それぞれの行の末尾に置きます。",
                en: "Puts artboards with no readable row-column at the end of each row."
            },
            exceptionLastRow: {
                ja: "アートボード名から行列を読み取れなかったものを、最終行の次の行にまとめます。",
                en: "Collects artboards with no readable row-column into a row after the last one."
            },
            changeArtboardOrder: {
                ja: "アートボードパネルの並び順も、配置後の順序に合わせて入れ替えます。",
                en: "Also reorders the Artboards panel to match the new layout."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noMatch: {
                ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）の形式のアートボード名が見つかりませんでした。",
                en: "No artboard names in '<row><sep><column>' format (e.g., 1-2) or '<prefix><sep><number>' format (e.g., banner-1) were found."
            }
        }
    };

    /**
     * ドット区切りのパスから現在言語のラベルを取得する
     * @param {string} labelPath - "panel.spacing" のようなパス
     * @returns {string} 表示言語の文字列（見つからなければ labelPath）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
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
    // 単位と数値入力 / Units and numeric input
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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * pt値を表示単位に変換する
     * @param {number} points - pt値
     * @param {Object} unitInfo - getUnitInfo() の結果
     * @returns {number} 表示単位の値
     */
    function pointsToDisplayUnit(points, unitInfo) {
        return points / unitInfo.pointsPerUnit;
    }

    /**
     * 表示単位をpt値に変換する
     * @param {number} value - 表示単位の値
     * @param {Object} unitInfo - getUnitInfo() の結果
     * @returns {number} pt値
     */
    function displayUnitToPoints(value, unitInfo) {
        return value * unitInfo.pointsPerUnit;
    }

    /**
     * 表示用の数値を整える（小数点以下3桁まで）
     * @param {number} value - 値
     * @returns {string} 表示用の文字列
     */
    function formatDisplayNumber(value) {
        var rounded = Math.round(value * 1000) / 1000;
        return String(rounded);
    }

    /**
     * ↑↓キーで数値を増減する（0 未満にはしない）
     * ↑↓で±1、Shift+↑↓で10の倍数にスナップ、Option+↑↓で±0.1。
     * @param {EditText} numericEditText - 対象の入力欄
     * @param {function} [afterChange] - 値を変えたあとに呼ぶ関数（入力欄を受け取る）
     * @returns {void}
     */
    function changeValueByArrowKey(numericEditText, afterChange) {
        numericEditText.addEventListener("keydown", function (event) {
            var currentValue = Number(numericEditText.text);
            if (isNaN(currentValue)) return;

            var isUp = (event.keyName === "Up");
            if (!isUp && event.keyName !== "Down") return;

            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.shiftKey) {
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift */
                currentValue = isUp ? Math.ceil((currentValue + 1) / 10) * 10 : Math.floor((currentValue - 1) / 10) * 10;
            } else {
                /* Optionキー押下時は0.1単位で増減 / Change by 0.1 with Option */
                var step = keyboard.altKey ? 0.1 : 1;
                currentValue += isUp ? step : -step;
            }
            event.preventDefault();

            currentValue = keyboard.altKey ? Math.round(currentValue * 10) / 10 : Math.round(currentValue);
            if (currentValue < 0) currentValue = 0;
            numericEditText.text = currentValue;
            if (typeof afterChange === "function") afterChange(numericEditText);
        });
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 間隔の入力行（項目名＋入力欄）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルのパス
     * @param {number} initialPoints - 初期値（pt）
     * @param {Object} unitInfo - 定規単位の情報
     * @returns {EditText} 入力欄
     */
    function addSpacingField(parentGroup, labelPath, tooltipPath, initialPoints, unitInfo) {
        var fieldGroup = parentGroup.add('group');
        fieldGroup.add('statictext', undefined, labelText(labelPath));
        var fieldInput = fieldGroup.add('edittext', undefined, formatDisplayNumber(pointsToDisplayUnit(initialPoints, unitInfo)));
        fieldInput.helpTip = getLabel(tooltipPath);
        fieldInput.characters = SPACING_FIELD_CHARACTERS;
        return fieldInput;
    }

    /**
     * 間隔欄の値を pt で読み取る。数値でないか負の値なら初期値を使う
     * @param {EditText} spacingInput - 間隔欄
     * @param {number} fallbackPoints - 読み取れないときの値（pt）
     * @param {Object} unitInfo - 定規単位の情報
     * @returns {number} 間隔（pt）
     */
    function readSpacingPoints(spacingInput, fallbackPoints, unitInfo) {
        var parsedValue = parseFloat(spacingInput.text);
        if (isNaN(parsedValue) || parsedValue < 0) return fallbackPoints;
        return displayUnitToPoints(parsedValue, unitInfo);
    }

    /**
     * チェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルのパス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentPanel.add('checkbox', undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = !!initialValue;
        return optionCheckbox;
    }

    /**
     * 設定ダイアログを表示し、ユーザーが入力した値を返す
     * @param {Object} defaultSettings - 初期値（間隔は pt）
     * @returns {Object|null} 設定。キャンセル時は null
     */
    function showSettingsDialog(defaultSettings) {
        var settingsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        settingsDialog.orientation = 'column';
        settingsDialog.alignChildren = 'fill';
        settingsDialog.margins = DIALOG_MARGINS;

        var rulerUnitInfo = getUnitInfo();

        /* 間隔パネル（列間・行間と連動チェックボックス）/ Spacing panel (column/row gaps with link option) */
        var spacingPanel = settingsDialog.add('panel', undefined, getLabel('panel.spacing') + '（' + rulerUnitInfo.label + '）');
        setupPanel(spacingPanel, PANEL_SPACING);

        var spacingContentGroup = spacingPanel.add('group');
        spacingContentGroup.orientation = 'row';
        spacingContentGroup.alignChildren = ['left', 'center'];
        spacingContentGroup.spacing = SPACING_COLUMN_GAP;

        var spacingInputColumn = spacingContentGroup.add('group');
        spacingInputColumn.orientation = 'column';
        spacingInputColumn.alignChildren = 'left';
        spacingInputColumn.spacing = PANEL_SPACING;

        var spacingXInput = addSpacingField(spacingInputColumn, 'fieldLabel.spacingX', 'tooltip.spacingX', defaultSettings.spacingX, rulerUnitInfo);
        spacingXInput.active = true;
        var spacingYInput = addSpacingField(spacingInputColumn, 'fieldLabel.spacingY', 'tooltip.spacingY', defaultSettings.spacingY, rulerUnitInfo);

        var linkColumn = spacingContentGroup.add('group');
        linkColumn.orientation = 'column';
        linkColumn.alignChildren = ['left', 'center'];
        linkColumn.alignment = ['left', 'center'];
        var linkCheckbox = addOptionCheckbox(linkColumn, 'checkbox.linkSpacing', 'tooltip.linkSpacing', defaultSettings.linkSpacing);

        /* 連動時は列間を行間にミラーし、行間欄をディムする / When linked, mirror the column gap to the row gap and dim the row field */
        function syncLinkedSpacingValue() {
            if (linkCheckbox.value) spacingYInput.text = spacingXInput.text;
        }
        function syncLinkedSpacingInputState() {
            spacingYInput.enabled = !linkCheckbox.value;
        }
        syncLinkedSpacingInputState();

        changeValueByArrowKey(spacingXInput, syncLinkedSpacingValue);
        changeValueByArrowKey(spacingYInput);
        spacingXInput.onChanging = syncLinkedSpacingValue;
        linkCheckbox.onClick = function () {
            syncLinkedSpacingValue();
            syncLinkedSpacingInputState();
        };

        /* ロック／非表示の除外パネル / Locked / hidden exclusion panel */
        var exclusionPanel = settingsDialog.add('panel', undefined, getLabel('panel.exclusion'));
        setupPanel(exclusionPanel, PANEL_SPACING);
        var excludeLayersCheckbox = addOptionCheckbox(exclusionPanel, 'checkbox.excludeLockedHiddenLayers', 'tooltip.excludeLayers', defaultSettings.excludeLockedHiddenLayers);
        var excludeItemsCheckbox = addOptionCheckbox(exclusionPanel, 'checkbox.excludeLockedHiddenItems', 'tooltip.excludeItems', defaultSettings.excludeLockedHiddenItems);

        /* 未指定／重複の扱いパネル / Unmatched / duplicate handling panel */
        var exceptionPanel = settingsDialog.add('panel', undefined, getLabel('panel.exception'));
        setupPanel(exceptionPanel, PANEL_SPACING);
        var rowEndRadio = exceptionPanel.add('radiobutton', undefined, getLabel('radio.exceptionRowEnd'));
        rowEndRadio.helpTip = getLabel('tooltip.exceptionRowEnd');
        var lastRowRadio = exceptionPanel.add('radiobutton', undefined, getLabel('radio.exceptionLastRow'));
        lastRowRadio.helpTip = getLabel('tooltip.exceptionLastRow');
        if (defaultSettings.exceptionMode === 'lastRow') {
            lastRowRadio.value = true;
        } else {
            rowEndRadio.value = true;
        }

        /* オプションパネル / Options panel */
        var optionsPanel = settingsDialog.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel, PANEL_SPACING);
        var changeArtboardOrderCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.changeArtboardOrder', 'tooltip.changeArtboardOrder', defaultSettings.changeArtboardOrder);

        /* OK / キャンセルボタン / OK and Cancel buttons */
        var btnRowGroup = settingsDialog.add('group');
        btnRowGroup.alignment = 'right';
        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        if (settingsDialog.show() !== 1) return null;

        return {
            spacingX: readSpacingPoints(spacingXInput, defaultSettings.spacingX, rulerUnitInfo),
            spacingY: readSpacingPoints(spacingYInput, defaultSettings.spacingY, rulerUnitInfo),
            exceptionMode: lastRowRadio.value ? 'lastRow' : 'rowEnd',
            linkSpacing: linkCheckbox.value,
            excludeLockedHiddenLayers: excludeLayersCheckbox.value,
            excludeLockedHiddenItems: excludeItemsCheckbox.value,
            changeArtboardOrder: changeArtboardOrderCheckbox.value
        };
    }

    // =========================================
    // 配置処理 / Placement
    // =========================================

    /**
     * アクティブアートボード幅の1/8を初期間隔として取得する
     * @param {Document} documentRef - 対象のドキュメント
     * @returns {number} 間隔（pt）
     */
    function getDefaultSpacingFromActiveArtboard(documentRef) {
        var activeArtboardIndex = documentRef.artboards.getActiveArtboardIndex();
        var activeArtboardRect = documentRef.artboards[activeArtboardIndex].artboardRect;
        var activeArtboardWidth = activeArtboardRect[2] - activeArtboardRect[0];
        return Math.round(activeArtboardWidth / 8);
    }

    /**
     * アートボード名を解析する
     * 「数字 + 区切り文字 + 数字」（区切り文字 -, _, x。大文字小文字無視）を行列指定とし、
     * 行列指定の形でない場合のみ「接頭辞 + 区切り文字 + 数字」を接頭辞指定として判定する。
     * 0行・0列・0番、数字だけの接頭辞（行列指定と紛らわしい）は無効
     * @param {string} artboardName - アートボード名
     * @returns {Object|null} { matchType: 'rowColumn' | 'prefixNumber', rowNumber, columnNumber, prefixName }。どちらでもなければ null
     */
    function parseArtboardName(artboardName) {
        var rowColumnMatch = artboardName.match(/^(\d+)[-_x](\d+)$/i);
        if (rowColumnMatch) {
            var rowNumber = parseInt(rowColumnMatch[1], 10);
            var columnNumber = parseInt(rowColumnMatch[2], 10);
            if (rowNumber < 1 || columnNumber < 1) return null;
            return { matchType: 'rowColumn', rowNumber: rowNumber, columnNumber: columnNumber, prefixName: '' };
        }

        var prefixNumberMatch = artboardName.match(/^(.+?)[-_x](\d+)$/i);
        if (!prefixNumberMatch || /^\d+$/.test(prefixNumberMatch[1])) return null;
        var prefixColumnNumber = parseInt(prefixNumberMatch[2], 10);
        if (prefixColumnNumber < 1) return null;
        return { matchType: 'prefixNumber', rowNumber: 0, columnNumber: prefixColumnNumber, prefixName: prefixNumberMatch[1] };
    }

    /**
     * 第1パス：アートボードの情報を収集し、行列指定と接頭辞指定を判定する
     * @param {Artboards} artboards - アートボードのコレクション
     * @returns {Object} placementItems と、最大列番号・最大行番号・接頭辞の出現順などをまとめたグリッド情報
     */
    function collectArtboardPlacements(artboards) {
        var gridInfo = {
            placementItems: [],
            maxColumnNumber: 0,
            maxRowNumber: 0,
            prefixOrder: [],
            prefixRowOffsetByName: {},
            hasMatchedArtboard: false
        };

        for (var artboardIndex = 0; artboardIndex < artboards.length; artboardIndex++) {
            var artboard = artboards[artboardIndex];
            var artboardRect = artboard.artboardRect; // [left, top, right, bottom]
            var placementItem = {
                artboard: artboard,
                width: artboardRect[2] - artboardRect[0],
                height: artboardRect[1] - artboardRect[3],
                matched: false,
                matchType: '',
                prefixName: '',
                rowNumber: 0,
                columnNumber: 0,
                assignedRow: 0,
                assignedColumn: 0
            };

            var nameMatch = parseArtboardName(artboard.name);
            if (nameMatch) {
                placementItem.matched = true;
                placementItem.matchType = nameMatch.matchType;
                placementItem.prefixName = nameMatch.prefixName;
                placementItem.rowNumber = nameMatch.rowNumber;
                placementItem.columnNumber = nameMatch.columnNumber;

                if (nameMatch.matchType === 'rowColumn') {
                    if (nameMatch.rowNumber > gridInfo.maxRowNumber) gridInfo.maxRowNumber = nameMatch.rowNumber;
                } else if (gridInfo.prefixRowOffsetByName[nameMatch.prefixName] === undefined) {
                    /* 接頭辞は出現順に1行ずつ割り当てる / one row per prefix, in order of appearance */
                    gridInfo.prefixRowOffsetByName[nameMatch.prefixName] = gridInfo.prefixOrder.length + 1;
                    gridInfo.prefixOrder.push(nameMatch.prefixName);
                }
                if (nameMatch.columnNumber > gridInfo.maxColumnNumber) gridInfo.maxColumnNumber = nameMatch.columnNumber;
                gridInfo.hasMatchedArtboard = true;
            }
            gridInfo.placementItems.push(placementItem);
        }

        /* 接頭辞の行は通常の行列指定の下に続く / prefix rows follow the numeric rows */
        gridInfo.maxMatchedRowNumber = gridInfo.maxRowNumber + gridInfo.prefixOrder.length;
        return gridInfo;
    }

    /**
     * 第2パス：順番に走査して行・列番号を割り当てる
     * 例外処理モード / Exception handling modes:
     *   'rowEnd'  : 非一致名は直前にマッチした行に追従し、列は (maxColumnNumber + 1) から右へ並べる
     *               先頭の非一致は行1に配置 / Leading unmatched goes to row 1
     *   'lastRow' : 非一致名は (maxMatchedRowNumber + 1) 行目に、出現順に列 1, 2, 3, ... で並べる
     * 接頭辞指定は、通常の行列指定の下に接頭辞ごとの行として配置する。
     * 重複名は元の (row, col) と衝突した時点で、その行の末尾へ押し出す。
     * @param {Object} gridInfo - collectArtboardPlacements() の結果（assignedRow / assignedColumn を書き込む）
     * @param {string} exceptionMode - 'rowEnd' | 'lastRow'
     * @returns {void}
     */
    function assignGridSlots(gridInfo, exceptionMode) {
        var occupiedSlots = {};                                     // "row,col" → true
        var nextFreeColumnByRow = {};                               // row → next column to try
        var rowEndFirstColumn = gridInfo.maxColumnNumber + 1;       // 行末に押し出すときの開始列 / first column at a row end
        var exceptionRowNumber = gridInfo.maxMatchedRowNumber + 1;  // 'lastRow' 用の行番号 / row for 'lastRow'
        var currentRowNumber = 1;

        for (var assignIndex = 0; assignIndex < gridInfo.placementItems.length; assignIndex++) {
            var assignTarget = gridInfo.placementItems[assignIndex];

            if (assignTarget.matched) {
                var targetRowNumber = assignTarget.rowNumber;
                if (assignTarget.matchType === 'prefixNumber') {
                    targetRowNumber = gridInfo.maxRowNumber + gridInfo.prefixRowOffsetByName[assignTarget.prefixName];
                }

                currentRowNumber = targetRowNumber;
                assignTarget.assignedRow = targetRowNumber;
                var requestedKey = targetRowNumber + ',' + assignTarget.columnNumber;
                if (!occupiedSlots[requestedKey]) {
                    /* 通常配置 / Normal placement */
                    occupiedSlots[requestedKey] = true;
                    assignTarget.assignedColumn = assignTarget.columnNumber;
                } else {
                    /* 重複 → その行の末尾へ / Duplicate → push to row end */
                    assignTarget.assignedColumn = reserveNextFreeColumn(targetRowNumber, rowEndFirstColumn, occupiedSlots, nextFreeColumnByRow);
                }
            } else if (exceptionMode === 'lastRow') {
                assignTarget.assignedRow = exceptionRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(exceptionRowNumber, 1, occupiedSlots, nextFreeColumnByRow);
            } else {
                /* 'rowEnd' */
                assignTarget.assignedRow = currentRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(currentRowNumber, rowEndFirstColumn, occupiedSlots, nextFreeColumnByRow);
            }
        }
    }

    /**
     * 第3〜5パス：列ごとの最大幅・行ごとの最大高さから、各アートボードの新しい位置と移動量を求める
     * 例外領域に入る境界（未指定・重複の列、未指定の行）では間隔を3倍にする。
     * IllustratorのY座標は上がプラス、下がマイナス。
     * @param {Object[]} placementItems - 行列を割り当てたアートボード情報（oldRect / newLeft / newTop / deltaX / deltaY を書き込む）
     * @param {number} spacingX - 列間（pt）
     * @param {number} spacingY - 行間（pt）
     * @param {number} columnBoundary - 間隔を3倍にする列番号
     * @param {number} rowBoundary - 間隔を3倍にする行番号
     * @returns {void}
     */
    function computeArtboardDestinations(placementItems, spacingX, spacingY, columnBoundary, rowBoundary) {
        var originX = 0;    // 配置の開始X座標 / X origin
        var originY = 0;    // 配置の開始Y座標 / Y origin

        /* 列ごとの最大幅・行ごとの最大高さを集計する / Per-column max width and per-row max height */
        var columnWidthByNumber = {};
        var rowHeightByNumber = {};
        for (var aggregateIndex = 0; aggregateIndex < placementItems.length; aggregateIndex++) {
            var aggregateItem = placementItems[aggregateIndex];
            var columnKey = aggregateItem.assignedColumn;
            var rowKey = aggregateItem.assignedRow;
            if (columnWidthByNumber[columnKey] === undefined || aggregateItem.width > columnWidthByNumber[columnKey]) {
                columnWidthByNumber[columnKey] = aggregateItem.width;
            }
            if (rowHeightByNumber[rowKey] === undefined || aggregateItem.height > rowHeightByNumber[rowKey]) {
                rowHeightByNumber[rowKey] = aggregateItem.height;
            }
        }

        /* 使われている列番号・行番号の昇順で累積オフセットを計算する / Cumulative offsets over the used numbers */
        var columnOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(columnWidthByNumber), columnWidthByNumber, spacingX, columnBoundary
        );
        var rowOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(rowHeightByNumber), rowHeightByNumber, spacingY, rowBoundary
        );

        for (var placeIndex = 0; placeIndex < placementItems.length; placeIndex++) {
            var placeItem = placementItems[placeIndex];
            var oldRect = placeItem.artboard.artboardRect;
            placeItem.oldRect = oldRect;
            placeItem.newLeft = originX + columnOffsetByNumber[placeItem.assignedColumn];
            placeItem.newTop = originY - rowOffsetByNumber[placeItem.assignedRow];
            placeItem.deltaX = placeItem.newLeft - oldRect[0];
            placeItem.deltaY = placeItem.newTop - oldRect[1];
        }
    }

    /**
     * 第6パス：アートボード上のオブジェクトを、アートボードと同じだけ移動する
     * オブジェクトの中心が元のアートボード矩形内にあるものを対象にする。
     * @param {Layers} layerCollection - ドキュメントのレイヤー
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {boolean} excludeLockedHiddenLayers - ロック／非表示レイヤーのオブジェクトを動かさないか
     * @param {boolean} excludeLockedHiddenItems - ロック／非表示のオブジェクトを動かさないか
     * @returns {void}
     */
    function translateItemsWithArtboards(layerCollection, placementItems, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        var itemTranslations = [];
        appendItemTranslations(layerCollection, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems);
        for (var translateIndex = 0; translateIndex < itemTranslations.length; translateIndex++) {
            var pendingMove = itemTranslations[translateIndex];
            if (pendingMove.deltaX === 0 && pendingMove.deltaY === 0) continue;
            try {
                pendingMove.item.translate(pendingMove.deltaX, pendingMove.deltaY);
            } catch (translateError) {
                /* ロック・非表示などで失敗したオブジェクトはスキップ
                   Skip items that can't be translated (locked, hidden layer, etc.) */
            }
        }
    }

    /**
     * 第7パス：アートボードの矩形を更新する
     * @param {Object[]} placementItems - 新しい位置を求めたアートボード情報
     * @returns {void}
     */
    function applyArtboardRects(placementItems) {
        for (var applyIndex = 0; applyIndex < placementItems.length; applyIndex++) {
            var applyItem = placementItems[applyIndex];
            var newRight = applyItem.newLeft + applyItem.width;
            var newBottom = applyItem.newTop - applyItem.height;
            applyItem.artboard.artboardRect = [applyItem.newLeft, applyItem.newTop, newRight, newBottom];
        }
    }

    /**
     * アートボード名を解析してグリッド状に再配置する
     * 「行-列」は行列指定、「接頭辞-番号」は接頭辞ごとの行として扱い、
     * 列幅・行高は「列ごと・行ごとの最大サイズ」で計算する。
     * @returns {void}
     */
    function arrangeArtboardsByRowColumnName() {
        if (app.documents.length === 0) {
            return;
        }

        var documentRef = app.activeDocument;
        var defaultSpacing = getDefaultSpacingFromActiveArtboard(documentRef);

        var arrangeSettings = showSettingsDialog({
            spacingX: defaultSpacing,
            spacingY: defaultSpacing,
            exceptionMode: DIALOG_DEFAULTS.exceptionMode,
            linkSpacing: DIALOG_DEFAULTS.linkSpacing,
            excludeLockedHiddenLayers: DIALOG_DEFAULTS.excludeLockedHiddenLayers,
            excludeLockedHiddenItems: DIALOG_DEFAULTS.excludeLockedHiddenItems,
            changeArtboardOrder: DIALOG_DEFAULTS.changeArtboardOrder
        });
        if (!arrangeSettings) return;

        var gridInfo = collectArtboardPlacements(documentRef.artboards);
        if (!gridInfo.hasMatchedArtboard) {
            alert(getLabel('alert.noMatch'));
            return;
        }

        assignGridSlots(gridInfo, arrangeSettings.exceptionMode);
        computeArtboardDestinations(gridInfo.placementItems, arrangeSettings.spacingX, arrangeSettings.spacingY,
            gridInfo.maxColumnNumber + 1, gridInfo.maxMatchedRowNumber + 1);

        /* オブジェクトを先に動かしてからアートボードを動かす / Move the objects first, then the artboards */
        translateItemsWithArtboards(documentRef.layers, gridInfo.placementItems,
            arrangeSettings.excludeLockedHiddenLayers, arrangeSettings.excludeLockedHiddenItems);
        applyArtboardRects(gridInfo.placementItems);

        /* 第8パス: アートボードパネル上の順番を、行→列の順に並べ替える / Reorder the Artboards panel row-then-column */
        if (arrangeSettings.changeArtboardOrder) {
            reorderArtboardsByGridOrder(documentRef, gridInfo.placementItems);
        }

        /* 画面を全体表示に更新 / Fit all in view */
        app.executeMenuCommand('fitall');
    }

    // =========================================
    // 配置ヘルパー / Placement helpers
    // =========================================

    /**
     * 指定行の firstColumn 以降で空きスロットを予約する
     * @param {number} rowNumber - 行番号
     * @param {number} firstColumn - その行で最初に試す列番号
     * @param {Object} occupiedSlots - "row,col" → true の表
     * @param {Object} nextFreeColumnByRow - 行番号 → 次に試す列番号の表
     * @returns {number} 予約した列番号
     */
    function reserveNextFreeColumn(rowNumber, firstColumn, occupiedSlots, nextFreeColumnByRow) {
        if (nextFreeColumnByRow[rowNumber] === undefined) {
            nextFreeColumnByRow[rowNumber] = firstColumn;
        }
        while (true) {
            var candidateColumn = nextFreeColumnByRow[rowNumber]++;
            var candidateSlotKey = rowNumber + ',' + candidateColumn;
            if (!occupiedSlots[candidateSlotKey]) {
                occupiedSlots[candidateSlotKey] = true;
                return candidateColumn;
            }
        }
    }

    /**
     * オブジェクトの数値キーを昇順で取り出す
     * @param {Object} numericKeyedObject - 数値キーの表
     * @returns {number[]} 昇順のキー
     */
    function collectSortedNumericKeys(numericKeyedObject) {
        var sortedKeys = [];
        for (var numericKey in numericKeyedObject) {
            if (numericKeyedObject.hasOwnProperty(numericKey)) {
                sortedKeys.push(parseInt(numericKey, 10));
            }
        }
        sortedKeys.sort(function (a, b) { return a - b; });
        return sortedKeys;
    }

    /**
     * レイヤーを再帰的に走査し、直下の page item の移動量を集める
     * 必要に応じてロック／非表示レイヤー・オブジェクトを除外する。
     * @param {Layers} layerCollection - 走査するレイヤー
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {Object[]} itemTranslations - { item, deltaX, deltaY } の格納先
     * @param {boolean} excludeLockedHiddenLayers - ロック／非表示レイヤーを除外するか
     * @param {boolean} excludeLockedHiddenItems - ロック／非表示オブジェクトを除外するか
     * @returns {void}
     */
    function appendItemTranslations(layerCollection, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            if (excludeLockedHiddenLayers && shouldSkipLayerForItemMove(layer)) continue;

            for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
                var pageItem = layer.pageItems[itemIndex];
                if (pageItem.parent !== layer) continue;
                if (excludeLockedHiddenItems && shouldSkipPageItemForMove(pageItem)) continue;
                appendTranslationForItem(pageItem, placementItems, itemTranslations);
            }

            if (layer.layers && layer.layers.length > 0) {
                appendItemTranslations(layer.layers, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems);
            }
        }
    }

    /**
     * 移動対象から除外するレイヤーか判定する
     * @param {Layer} layer - 対象のレイヤー
     * @returns {boolean} ロックまたは非表示なら true
     */
    function shouldSkipLayerForItemMove(layer) {
        try {
            return layer.locked || !layer.visible;
        } catch (layerStateError) {
            return false;
        }
    }

    /**
     * 移動対象から除外するオブジェクトか判定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} ロックまたは非表示なら true
     */
    function shouldSkipPageItemForMove(pageItem) {
        try {
            return pageItem.locked || pageItem.hidden;
        } catch (pageItemStateError) {
            return false;
        }
    }

    /**
     * 単一オブジェクトの移動量を追加する（中心を含む最初の元アートボードの移動量）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {Object[]} itemTranslations - { item, deltaX, deltaY } の格納先
     * @returns {void}
     */
    function appendTranslationForItem(pageItem, placementItems, itemTranslations) {
        var center = getItemGeometricCenter(pageItem);
        if (!center) return;

        for (var placementIndex = 0; placementIndex < placementItems.length; placementIndex++) {
            var placement = placementItems[placementIndex];
            if (isCenterInsideRect(center, placement.oldRect)) {
                itemTranslations.push({ item: pageItem, deltaX: placement.deltaX, deltaY: placement.deltaY });
                break;
            }
        }
    }

    /**
     * オブジェクトの幾何中心を取得する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number[]|null} [x, y]。取得できない場合は null
     */
    function getItemGeometricCenter(pageItem) {
        try {
            var bounds = pageItem.geometricBounds; // [left, top, right, bottom]
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (geometricBoundsError) {
            return null;
        }
    }

    /**
     * 中心座標が矩形内にあるか（artboardRect の Y は上がプラス、下がマイナス）
     * @param {number[]} center - [x, y]
     * @param {number[]} rect - [left, top, right, bottom]
     * @returns {boolean} 矩形内なら true
     */
    function isCenterInsideRect(center, rect) {
        return center[0] >= rect[0] && center[0] <= rect[2] &&
            center[1] <= rect[1] && center[1] >= rect[3];
    }

    /**
     * アートボードパネル上の順序を、割り当てた (assignedRow, assignedColumn) の昇順に再構築する
     * 一時名にリネームしてから新規作成→旧アートボードを削除することで名称重複を回避する。
     * @param {Document} documentRef - 対象のドキュメント
     * @param {Object[]} placementItems - 行列を割り当てたアートボード情報
     * @returns {void}
     */
    function reorderArtboardsByGridOrder(documentRef, placementItems) {
        var sortedPlacements = placementItems.slice();
        sortedPlacements.sort(function (placementA, placementB) {
            if (placementA.assignedRow !== placementB.assignedRow) return placementA.assignedRow - placementB.assignedRow;
            return placementA.assignedColumn - placementB.assignedColumn;
        });

        var artboards = documentRef.artboards;
        var originalCount = artboards.length;

        var orderedArtboardSpecs = [];
        for (var collectIndex = 0; collectIndex < sortedPlacements.length; collectIndex++) {
            var sourceArtboard = sortedPlacements[collectIndex].artboard;
            var sourceRect = sourceArtboard.artboardRect;
            orderedArtboardSpecs.push({
                name: sourceArtboard.name,
                rect: [sourceRect[0], sourceRect[1], sourceRect[2], sourceRect[3]]
            });
        }

        for (var renameIndex = 0; renameIndex < originalCount; renameIndex++) {
            artboards[renameIndex].name = '__reorder_tmp__' + renameIndex;
        }

        for (var addIndex = 0; addIndex < orderedArtboardSpecs.length; addIndex++) {
            var newArtboard = artboards.add(orderedArtboardSpecs[addIndex].rect);
            newArtboard.name = orderedArtboardSpecs[addIndex].name;
        }

        for (var removeIndex = originalCount - 1; removeIndex >= 0; removeIndex--) {
            artboards[removeIndex].remove();
        }

        documentRef.artboards.setActiveArtboardIndex(0);
    }

    /**
     * 各番号（列または行）の累積オフセットを計算する
     * boundaryNumber と一致するキーの直前のギャップは3倍に広げる。
     * @param {number[]} sortedNumbers - 昇順の番号
     * @param {Object} sizeByNumber - 番号 → 幅または高さ
     * @param {number} gap - 間隔（pt）
     * @param {number} boundaryNumber - 間隔を3倍にする番号
     * @returns {Object} 番号 → オフセット
     */
    function computeCumulativeOffsets(sortedNumbers, sizeByNumber, gap, boundaryNumber) {
        var offsetByNumber = {};
        var cumulative = 0;
        for (var orderIndex = 0; orderIndex < sortedNumbers.length; orderIndex++) {
            var currentNumber = sortedNumbers[orderIndex];
            if (orderIndex > 0) {
                var currentGap = gap;
                if (currentNumber === boundaryNumber) currentGap *= 3;
                cumulative += currentGap;
            }
            offsetByNumber[currentNumber] = cumulative;
            cumulative += sizeByNumber[currentNumber];
        }
        return offsetByNumber;
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    arrangeArtboardsByRowColumnName();

})();
