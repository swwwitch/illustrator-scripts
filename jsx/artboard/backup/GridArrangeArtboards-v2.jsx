#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要

Illustrator のアートボード名を解析し、行列グリッドとして再配置するスクリプト。
列間・行間は、環境設定の定規単位（rulerType）に合わせて入力できる。
初期値は、アクティブアートボード幅の1/8を整数値に丸めた値を使用する。

名前全体が「行番号 + 区切り文字（-, _, x）+ 列番号」に一致するアートボードは、行列指定として扱う。
例: 1-1, 1_2, 2x1

行列指定に一致しない場合のみ、「接頭辞 + 区切り文字（-, _, x）+ 番号」を接頭辞指定として扱う。
例: banner-1, banner_2, iconx3
接頭辞指定は、通常の行列指定の下に、接頭辞ごとの行として配置する。

0行・0列、および0番は無効として扱う。
数字だけの接頭辞は、行列指定との混同を避けるため接頭辞指定として扱わない。

未指定名や重複名は、指定した例外処理モードに従って末尾側へ自動配置する。
例外領域に入る境界では、列間・行間を広めに取って区別しやすくする。
アートボード上のオブジェクトも、元アートボード内にあるものは同じ移動量で移動する。
グループはグループ単位で移動し、グループ内の個別オブジェクトは分解しない。
ロック／非表示レイヤー、およびロック／非表示オブジェクトを除外できる。
必要に応じて、アートボードパネル上の並び順も行→列の順に再構築できる。

### Overview

Parses Illustrator artboard names and rearranges artboards as a row/column grid.
Column and row gaps can be entered using the document ruler unit (rulerType).
The default gap is the active artboard width divided by 8, rounded to an integer.

Artboard names that fully match "row number + separator (-, _, or x) + column number" are treated as row/column targets.
Examples: 1-1, 1_2, 2x1

Only when a name does not match the row/column pattern, names that fully match "prefix + separator (-, _, or x) + number" are treated as prefix-number targets.
Examples: banner-1, banner_2, iconx3
Prefix-number targets are placed below numeric row/column targets, grouped into separate rows by prefix.

Row 0, column 0, and number 0 are treated as invalid.
Numeric-only prefixes are not treated as prefix-number targets to avoid ambiguity with row/column names.

Unmatched or duplicate names are automatically placed at the end according to the selected exception-handling mode.
The gap is widened at the boundary into the exception area for better visual separation.
Items inside the original artboard are moved by the same delta as the artboard.
Groups are moved as groups; individual items inside groups are not separated.
Locked/hidden layers and locked/hidden objects can optionally be excluded.
The Artboards panel order can also be rebuilt in row-then-column order when needed.
*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version and localization
    // =========================================

    var SCRIPT_VERSION = "v1.1.0";

    /* 現在の UI 言語を判定する / Detect the current UI language */
    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "アートボードを再配置（行列）",
            en: "Rearrange Artboards (Grid)"
        },
        spacingPanel: {
            ja: "間隔",
            en: "Spacing"
        },
        spacingX: {
            ja: "列間",
            en: "Column gap"
        },
        spacingY: {
            ja: "行間",
            en: "Row gap"
        },
        linkSpacing: {
            ja: "連動",
            en: "Link"
        },
        optionPanel: {
            ja: "ロック／非表示を除外",
            en: "Exclude Locked / Hidden"
        },
        excludeLockedHiddenLayers: {
            ja: "レイヤー",
            en: "Layers"
        },
        excludeLockedHiddenItems: {
            ja: "オブジェクト",
            en: "Objects"
        },
        exceptionPanel: {
            ja: "未指定／重複",
            en: "Unmatched / Duplicate"
        },
        exceptionRowEnd: {
            ja: "直前の行の末尾に配置",
            en: "Append to current row"
        },
        exceptionLastRow: {
            ja: "最後の行にまとめる",
            en: "Collect in final row"
        },
        optionsPanel: {
            ja: "オプション",
            en: "Options"
        },
        changeArtboardOrder: {
            ja: "パネル上の並び順を変更",
            en: "Reorder in Artboards panel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        noMatchAlert: {
            ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）の形式のアートボード名が見つかりませんでした。",
            en: "No artboard names in '<row><sep><column>' format (e.g., 1-2) or '<prefix><sep><number>' format (e.g., banner-1) were found."
        }
    };

    /* キーから現在言語のラベルを取得する / Resolve a label for the current language */
    function getLabel(labelKey) {
        var labelEntry = LABELS[labelKey];
        if (!labelEntry) return labelKey;
        return labelEntry[currentLanguage] || labelEntry.en;
    }

    // =========================================
    // 共通ヘルパー / Common helpers
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];

    /* パネルを共通スタイルで初期化 / Initialize a panel with common styling */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    // =========================================
    // 単位変換と数値入力 / Unit conversion and numeric input
    // =========================================

    /* 環境設定の定規単位を取得する / Get the ruler unit from preferences */
    function getRulerUnitInfo() {
        var rulerUnit = app.preferences.getIntegerPreference("rulerType");
        var unitLabel = "pt";
        var unitFactor = 1.0;

        switch (rulerUnit) {
            case 0: // inch
                unitLabel = "inch";
                unitFactor = 72.0;
                break;
            case 1: // mm
                unitLabel = "mm";
                unitFactor = 72.0 / 25.4;
                break;
            case 2: // pt
                unitLabel = "pt";
                unitFactor = 1.0;
                break;
            case 3: // pica
                unitLabel = "pica";
                unitFactor = 12.0;
                break;
            case 4: // cm
                unitLabel = "cm";
                unitFactor = 72.0 / 2.54;
                break;
            case 5: // Q
                unitLabel = "Q";
                unitFactor = 72.0 / 25.4 * 0.25;
                break;
            case 6: // px
                unitLabel = "px";
                unitFactor = 1.0;
                break;
            default:
                unitLabel = "pt";
                unitFactor = 1.0;
        }

        return {
            label: unitLabel,
            factor: unitFactor
        };
    }

    /* pt値を表示単位に変換する / Convert points to display units */
    function pointsToDisplayUnit(points, unitInfo) {
        return points / unitInfo.factor;
    }

    /* 表示単位をpt値に変換する / Convert display units to points */
    function displayUnitToPoints(value, unitInfo) {
        return value * unitInfo.factor;
    }

    /* 表示用の数値を整える / Format a number for display */
    function formatDisplayNumber(value) {
        var rounded = Math.round(value * 1000) / 1000;
        return String(rounded);
    }

    /* ↑↓キーで数値を増減する / Change numeric value with Up/Down arrow keys */
    function changeValueByArrowKey(numericEditText, afterChange) {
        numericEditText.addEventListener("keydown", function (event) {
            var currentValue = Number(numericEditText.text);
            if (isNaN(currentValue)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;
            var handled = false;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift */
                if (event.keyName === "Up") {
                    currentValue = Math.ceil((currentValue + 1) / delta) * delta;
                    event.preventDefault();
                    handled = true;
                } else if (event.keyName === "Down") {
                    currentValue = Math.floor((currentValue - 1) / delta) * delta;
                    if (currentValue < 0) currentValue = 0;
                    event.preventDefault();
                    handled = true;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                /* Optionキー押下時は0.1単位で増減 / Change by 0.1 with Option */
                if (event.keyName === "Up") {
                    currentValue += delta;
                    event.preventDefault();
                    handled = true;
                } else if (event.keyName === "Down") {
                    currentValue -= delta;
                    event.preventDefault();
                    handled = true;
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    currentValue += delta;
                    event.preventDefault();
                    handled = true;
                } else if (event.keyName === "Down") {
                    currentValue -= delta;
                    if (currentValue < 0) currentValue = 0;
                    event.preventDefault();
                    handled = true;
                }
            }

            if (!handled) return;

            if (keyboard.altKey) {
                currentValue = Math.round(currentValue * 10) / 10;
            } else {
                currentValue = Math.round(currentValue);
            }

            if (currentValue < 0) currentValue = 0;
            numericEditText.text = currentValue;
            if (typeof afterChange === "function") afterChange(numericEditText);
        });
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 設定ダイアログを表示し、ユーザーが入力した値を返す / Show the settings dialog and return the user input */
    function showSettingsDialog(defaultSettings) {
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = 'fill';
        dialog.margins = 16;

        var rulerUnitInfo = getRulerUnitInfo();
        var displaySpacingX = pointsToDisplayUnit(defaultSettings.spacingX, rulerUnitInfo);
        var displaySpacingY = pointsToDisplayUnit(defaultSettings.spacingY, rulerUnitInfo);

        /* 間隔パネル（列間・行間と連動チェックボックス）/ Spacing panel (column/row gaps with link option) */
        var spacingPanel = dialog.add('panel', undefined, getLabel('spacingPanel') + '（' + rulerUnitInfo.label + '）');
        setupPanel(spacingPanel, 6);

        var spacingContentGroup = spacingPanel.add('group');
        spacingContentGroup.orientation = 'row';
        spacingContentGroup.alignChildren = ['left', 'center'];
        spacingContentGroup.spacing = 20;

        var spacingInputColumn = spacingContentGroup.add('group');
        spacingInputColumn.orientation = 'column';
        spacingInputColumn.alignChildren = 'left';
        spacingInputColumn.spacing = 6;

        var spacingXGroup = spacingInputColumn.add('group');
        spacingXGroup.add('statictext', undefined, getLabel('spacingX') + (currentLanguage === 'ja' ? '：' : ':'));
        var spacingXInput = spacingXGroup.add('edittext', undefined, formatDisplayNumber(displaySpacingX));
        spacingXInput.characters = 5;
        spacingXInput.active = true;

        var spacingYGroup = spacingInputColumn.add('group');
        spacingYGroup.add('statictext', undefined, getLabel('spacingY') + (currentLanguage === 'ja' ? '：' : ':'));
        var spacingYInput = spacingYGroup.add('edittext', undefined, formatDisplayNumber(displaySpacingY));
        spacingYInput.characters = 5;

        var linkColumn = spacingContentGroup.add('group');
        linkColumn.orientation = 'column';
        linkColumn.alignChildren = ['left', 'center'];
        linkColumn.alignment = ['left', 'center'];
        var linkCheckbox = linkColumn.add('checkbox', undefined, getLabel('linkSpacing'));
        linkCheckbox.value = !!defaultSettings.linkSpacing;

        function syncLinkedSpacingInputState() {
            spacingYInput.enabled = !linkCheckbox.value;
        }
        syncLinkedSpacingInputState();

        function syncLinkedSpacingValue() {
            if (linkCheckbox.value) spacingYInput.text = spacingXInput.text;
        }

        changeValueByArrowKey(spacingXInput, syncLinkedSpacingValue);
        changeValueByArrowKey(spacingYInput);

        /* 連動時は片方の入力をもう片方にミラーする / When linked, mirror one input to the other */
        spacingXInput.onChanging = function () {
            syncLinkedSpacingValue();
        };
        linkCheckbox.onClick = function () {
            if (linkCheckbox.value) spacingYInput.text = spacingXInput.text;
            syncLinkedSpacingInputState();
        };

        /* オプションパネル / Options panel */
        var optionPanel = dialog.add('panel', undefined, getLabel('optionPanel'));
        setupPanel(optionPanel, 6);
        var excludeLockedHiddenLayersCheckbox = optionPanel.add('checkbox', undefined, getLabel('excludeLockedHiddenLayers'));
        excludeLockedHiddenLayersCheckbox.value = !!defaultSettings.excludeLockedHiddenLayers;

        var excludeLockedHiddenItemsCheckbox = optionPanel.add('checkbox', undefined, getLabel('excludeLockedHiddenItems'));
        excludeLockedHiddenItemsCheckbox.value = !!defaultSettings.excludeLockedHiddenItems;

        /* 例外処理パネル / Exception handling panel */
        var exceptionPanel = dialog.add('panel', undefined, getLabel('exceptionPanel'));
        setupPanel(exceptionPanel, 6);
        var rowEndRadio = exceptionPanel.add('radiobutton', undefined, getLabel('exceptionRowEnd'));
        var lastRowRadio = exceptionPanel.add('radiobutton', undefined, getLabel('exceptionLastRow'));
        if (defaultSettings.exceptionMode === 'lastRow') {
            lastRowRadio.value = true;
        } else {
            rowEndRadio.value = true;
        }

        /* オプションパネル / Options panel */
        var optionsPanel = dialog.add('panel', undefined, getLabel('optionsPanel'));
        setupPanel(optionsPanel, 6);
        var changeArtboardOrderCheckbox = optionsPanel.add('checkbox', undefined, getLabel('changeArtboardOrder'));
        changeArtboardOrderCheckbox.value = !!defaultSettings.changeArtboardOrder;

        /* OK / キャンセルボタン / OK and Cancel buttons */
        var buttonGroup = dialog.add('group');
        buttonGroup.alignment = 'right';
        buttonGroup.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
        buttonGroup.add('button', undefined, getLabel('ok'), { name: 'ok' });

        if (dialog.show() !== 1) return null;

        var parsedSpacingX = parseFloat(spacingXInput.text);
        if (isNaN(parsedSpacingX) || parsedSpacingX < 0) {
            parsedSpacingX = defaultSettings.spacingX;
        } else {
            parsedSpacingX = displayUnitToPoints(parsedSpacingX, rulerUnitInfo);
        }
        var parsedSpacingY = parseFloat(spacingYInput.text);
        if (isNaN(parsedSpacingY) || parsedSpacingY < 0) {
            parsedSpacingY = defaultSettings.spacingY;
        } else {
            parsedSpacingY = displayUnitToPoints(parsedSpacingY, rulerUnitInfo);
        }
        var selectedExceptionMode = lastRowRadio.value ? 'lastRow' : 'rowEnd';
        return {
            spacingX: parsedSpacingX,
            spacingY: parsedSpacingY,
            exceptionMode: selectedExceptionMode,
            linkSpacing: linkCheckbox.value,
            excludeLockedHiddenLayers: excludeLockedHiddenLayersCheckbox.value,
            excludeLockedHiddenItems: excludeLockedHiddenItemsCheckbox.value,
            changeArtboardOrder: changeArtboardOrderCheckbox.value
        };
    }

    // =========================================
    // 配置処理 / Placement
    // =========================================

    /* アクティブアートボード幅の1/8を初期間隔として取得する / Get 1/8 of the active artboard width as the default gap */
    function getDefaultSpacingFromActiveArtboard(activeDocument) {
        var activeArtboardIndex = activeDocument.artboards.getActiveArtboardIndex();
        var activeArtboardRect = activeDocument.artboards[activeArtboardIndex].artboardRect;
        var activeArtboardWidth = activeArtboardRect[2] - activeArtboardRect[0];
        return Math.round(activeArtboardWidth / 8);
    }

    /* アートボード名を解析してグリッド状に再配置する
       Parse artboard names and rearrange them as a grid.
       「行-列」は行列指定、「接頭辞-番号」は接頭辞ごとの行として扱う
       '<row><sep><column>' is treated as row-column placement;
       '<prefix><sep><number>' is placed in rows grouped by prefix.
       列幅・行高は「列ごと・行ごとの最大サイズ」で計算する
       Column widths and row heights are computed per-column / per-row max size. */
    function arrangeArtboardsByRowColumnName() {
        if (app.documents.length === 0) {
            return;
        }

        var activeDocument = app.activeDocument;
        var artboards = activeDocument.artboards;
        var defaultSpacing = getDefaultSpacingFromActiveArtboard(activeDocument);

        var settings = showSettingsDialog({
            spacingX: defaultSpacing,
            spacingY: defaultSpacing,
            exceptionMode: 'rowEnd',
            linkSpacing: true,
            excludeLockedHiddenLayers: false,
            excludeLockedHiddenItems: false,
            changeArtboardOrder: false
        });
        if (!settings) return;

        /* 設定 / Settings */
        var spacingX = settings.spacingX;           // 列間の間隔 / Column gap
        var spacingY = settings.spacingY;           // 行間の間隔 / Row gap
        var exceptionMode = settings.exceptionMode; // 'rowEnd' | 'lastRow'
        var excludeLockedHiddenLayers = settings.excludeLockedHiddenLayers; // ロック／非表示レイヤーを除外 / Exclude locked/hidden layers
        var excludeLockedHiddenItems = settings.excludeLockedHiddenItems; // ロック／非表示オブジェクトを除外 / Exclude locked/hidden items
        var changeArtboardOrder = settings.changeArtboardOrder; // アートボードパネルの順を再構築 / Reorder artboards in panel
        var originX = 0;                            // 配置の開始X座標 / X origin
        var originY = 0;                            // 配置の開始Y座標 / Y origin

        var maxColumnNumber = 0;
        var maxRowNumber = 0;
        var placementItems = [];
        var hasMatchedArtboard = false;
        var prefixOrder = [];
        var prefixRowOffsetByName = {};

        /* 第1パス: アートボードの情報を収集し、行列指定と接頭辞指定を判定する
           First pass: collect artboard info and detect row-column and prefix-number names */
        for (var artboardIndex = 0; artboardIndex < artboards.length; artboardIndex++) {
            var artboard = artboards[artboardIndex];
            var artboardName = artboard.name;

            var artboardRect = artboard.artboardRect; // [left, top, right, bottom]
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];

            /* 「数字 + 区切り文字 + 数字」を抽出（区切り文字 -, _, x。大文字小文字無視）
               Extract '<digits><separator><digits>' (separator: -, _, x; case-insensitive).
               行列指定に一致しない場合のみ「接頭辞 + 区切り文字 + 数字」を判定する
               Prefix-number names are checked only when the row-column pattern does not match. */
            var rowColumnMatch = artboardName.match(/^(\d+)[-_x](\d+)$/i);
            var placementItem = {
                artboard: artboard,
                width: artboardWidth,
                height: artboardHeight,
                matched: false,
                matchType: '',
                prefixName: '',
                rowNumber: 0,
                columnNumber: 0,
                assignedRow: 0,
                assignedColumn: 0
            };
            if (rowColumnMatch) {
                placementItem.rowNumber = parseInt(rowColumnMatch[1], 10);
                placementItem.columnNumber = parseInt(rowColumnMatch[2], 10);

                /* 0行・0列は無効として扱う / Treat row 0 or column 0 as invalid */
                if (placementItem.rowNumber >= 1 && placementItem.columnNumber >= 1) {
                    placementItem.matched = true;
                    placementItem.matchType = 'rowColumn';

                    if (placementItem.columnNumber > maxColumnNumber) {
                        maxColumnNumber = placementItem.columnNumber;
                    }

                    if (placementItem.rowNumber > maxRowNumber) {
                        maxRowNumber = placementItem.rowNumber;
                    }

                    hasMatchedArtboard = true;
                }
            } else {
                var prefixNumberMatch = artboardName.match(/^(.+?)[-_x](\d+)$/i);

                /* 数字だけの接頭辞は行列指定と紛らわしいため除外する / Exclude numeric-only prefixes to avoid ambiguity with row-column names */
                if (prefixNumberMatch && !/^\d+$/.test(prefixNumberMatch[1])) {
                    placementItem.prefixName = prefixNumberMatch[1];
                    placementItem.columnNumber = parseInt(prefixNumberMatch[2], 10);

                    /* 0番は無効として扱う / Treat number 0 as invalid */
                    if (placementItem.columnNumber >= 1) {
                        placementItem.matched = true;
                        placementItem.matchType = 'prefixNumber';

                        if (prefixRowOffsetByName[placementItem.prefixName] === undefined) {
                            prefixRowOffsetByName[placementItem.prefixName] = prefixOrder.length + 1;
                            prefixOrder.push(placementItem.prefixName);
                        }

                        if (placementItem.columnNumber > maxColumnNumber) {
                            maxColumnNumber = placementItem.columnNumber;
                        }

                        hasMatchedArtboard = true;
                    }
                }
            }
            placementItems.push(placementItem);
        }

        if (!hasMatchedArtboard) {
            alert(getLabel('noMatchAlert'));
            return;
        }

        var maxMatchedRowNumber = maxRowNumber + prefixOrder.length;

        /* 第2パス: 順番に走査して行・列番号を割り当てる
           Second pass: traverse in order and assign row/column numbers
           例外処理モード / Exception handling modes:
             'rowEnd'  : 非一致名は直前にマッチした行に追従し、列は (maxColumnNumber + 1) から右へ並べる
                         先頭の非一致は行1に配置 / Leading unmatched goes to row 1
             'lastRow' : 非一致名は (maxMatchedRowNumber + 1) 行目に、出現順に列 1, 2, 3, ... で並べる
           接頭辞指定は、通常の行列指定の下に接頭辞ごとの行として配置する
           Prefix-number names are placed below numeric row-column names, grouped by prefix.
           重複名は元の (row, col) と衝突した時点で、その行の末尾へ押し出す
           Duplicate names are pushed to the end of their row when they collide. */
        var occupiedSlots = {};                                // "row,col" → true
        var nextExceptionColumnByRow = {};                     // row → next column to try (maxColumnNumber + 1 起点)
        var currentRowNumber = 1;
        var lastRowExceptionColumn = 1;                        // 'lastRow' 用カウンタ / counter for 'lastRow'
        var exceptionRowNumber = maxMatchedRowNumber + 1;      // 'lastRow' 用の行番号 / row for 'lastRow'

        for (var assignIndex = 0; assignIndex < placementItems.length; assignIndex++) {
            var assignTarget = placementItems[assignIndex];

            if (assignTarget.matched) {
                var targetRowNumber = assignTarget.rowNumber;
                if (assignTarget.matchType === 'prefixNumber') {
                    targetRowNumber = maxRowNumber + prefixRowOffsetByName[assignTarget.prefixName];
                }

                currentRowNumber = targetRowNumber;
                var requestedKey = targetRowNumber + ',' + assignTarget.columnNumber;
                if (!occupiedSlots[requestedKey]) {
                    /* 通常配置 / Normal placement */
                    occupiedSlots[requestedKey] = true;
                    assignTarget.assignedRow = targetRowNumber;
                    assignTarget.assignedColumn = assignTarget.columnNumber;
                } else {
                    /* 重複 → その行の末尾へ / Duplicate → push to row end */
                    assignTarget.assignedRow = targetRowNumber;
                    assignTarget.assignedColumn = reserveNextColumnAtRowEnd(
                        targetRowNumber, occupiedSlots, nextExceptionColumnByRow, maxColumnNumber
                    );
                }
            } else if (exceptionMode === 'lastRow') {
                assignTarget.assignedRow = exceptionRowNumber;
                while (true) {
                    var lastRowColCandidate = lastRowExceptionColumn++;
                    var lastRowKey = exceptionRowNumber + ',' + lastRowColCandidate;
                    if (!occupiedSlots[lastRowKey]) {
                        occupiedSlots[lastRowKey] = true;
                        assignTarget.assignedColumn = lastRowColCandidate;
                        break;
                    }
                }
            } else {
                /* 'rowEnd' */
                assignTarget.assignedRow = currentRowNumber;
                assignTarget.assignedColumn = reserveNextColumnAtRowEnd(
                    currentRowNumber, occupiedSlots, nextExceptionColumnByRow, maxColumnNumber
                );
            }
        }

        /* 第3パス: 列ごとの最大幅・行ごとの最大高さを集計する
           Third pass: aggregate per-column max width and per-row max height */
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

        /* 使用されている列番号・行番号を昇順で取得 / Collect used column/row numbers in ascending order */
        var usedColumnNumbers = collectSortedNumericKeys(columnWidthByNumber);
        var usedRowNumbers = collectSortedNumericKeys(rowHeightByNumber);

        /* 第4パス: 累積オフセットを計算する
           Fourth pass: compute cumulative offsets.
           例外領域に入る境界（maxColumnNumber+1 / maxRowNumber+1）では間隔を3倍にする
           Triple the gap at the boundary into the exception area. */
        var columnOffsetByNumber = computeCumulativeOffsets(
            usedColumnNumbers, columnWidthByNumber, spacingX, maxColumnNumber + 1
        );
        var rowOffsetByNumber = computeCumulativeOffsets(
            usedRowNumbers, rowHeightByNumber, spacingY, maxMatchedRowNumber + 1
        );

        /* 第5パス: 各アートボードの新位置と移動量を算出
           Fifth pass: compute new position and translation delta for each artboard.
           (IllustratorのY座標は上がプラス、下がマイナス / Y is positive upward) */
        for (var placeIndex = 0; placeIndex < placementItems.length; placeIndex++) {
            var placeItem = placementItems[placeIndex];
            var oldRect = placeItem.artboard.artboardRect;
            var newLeft = originX + columnOffsetByNumber[placeItem.assignedColumn];
            var newTop = originY - rowOffsetByNumber[placeItem.assignedRow];
            placeItem.oldRect = oldRect;
            placeItem.newLeft = newLeft;
            placeItem.newTop = newTop;
            placeItem.deltaX = newLeft - oldRect[0];
            placeItem.deltaY = newTop - oldRect[1];
        }

        /* 第6パス: アートボード上のオブジェクトを先に移動する
           Sixth pass: translate the objects sitting on each artboard before moving the artboard rect.
           オブジェクトの中心が元のアートボード矩形内にあるものを対象にする
           Items whose geometric center is within the original artboard rect are translated. */
        var itemTranslations = collectItemTranslations(
            activeDocument.layers,
            placementItems,
            excludeLockedHiddenLayers,
            excludeLockedHiddenItems
        );
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

        /* 第7パス: アートボードの矩形を更新する / Seventh pass: apply the new artboard rects */
        for (var applyIndex = 0; applyIndex < placementItems.length; applyIndex++) {
            var applyItem = placementItems[applyIndex];
            var newRight = applyItem.newLeft + applyItem.width;
            var newBottom = applyItem.newTop - applyItem.height;
            applyItem.artboard.artboardRect = [applyItem.newLeft, applyItem.newTop, newRight, newBottom];
        }

        /* 第8パス: アートボードパネル上の順番を、行→列の順に並べ替える
           Eighth pass: reorder artboards in the Artboards panel as row-then-column. */
        if (changeArtboardOrder) {
            reorderArtboardsByGridOrder(activeDocument, placementItems);
        }

        /* 画面を全体表示に更新 / Fit all in view */
        app.executeMenuCommand('fitall');
    }

    // =========================================
    // 配置ヘルパー / Placement helpers
    // =========================================

    /* 指定行の末尾（maxColumnNumber + n）で空きスロットを予約する
       Reserve the next empty slot at the end of the given row. */
    function reserveNextColumnAtRowEnd(rowNumber, occupiedSlots, nextExceptionColumnByRow, maxColumnNumber) {
        if (nextExceptionColumnByRow[rowNumber] === undefined) {
            nextExceptionColumnByRow[rowNumber] = maxColumnNumber + 1;
        }
        while (true) {
            var candidateColumn = nextExceptionColumnByRow[rowNumber]++;
            var candidateSlotKey = rowNumber + ',' + candidateColumn;
            if (!occupiedSlots[candidateSlotKey]) {
                occupiedSlots[candidateSlotKey] = true;
                return candidateColumn;
            }
        }
    }

    /* オブジェクトの数値キーを昇順で取り出す / Collect numeric keys of an object in ascending order */
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

    /* レイヤーを再帰的に走査し、直下の page item の移動量を集める
       Walk layers recursively and collect a translation delta for each direct child page item.
       必要に応じてロック／非表示レイヤーを除外する
       Optionally exclude locked/hidden layers.
       オブジェクトの中心がどの元アートボード内にあるか判定して、対応する deltaX/deltaY を割り当てる
       Each item is matched against placementItems by geometric-center inclusion in the original rect. */
    function collectItemTranslations(layerCollection, placementItems, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        var collected = [];
        appendItemTranslations(
            layerCollection,
            placementItems,
            collected,
            excludeLockedHiddenLayers,
            excludeLockedHiddenItems
        );
        return collected;
    }

    function appendItemTranslations(layerCollection, placementItems, output, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            if (excludeLockedHiddenLayers && shouldSkipLayerForItemMove(layer)) continue;

            for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
                var pageItem = layer.pageItems[itemIndex];
                if (pageItem.parent !== layer) continue;
                if (excludeLockedHiddenItems && shouldSkipPageItemForMove(pageItem)) continue;
                appendTranslationForItem(pageItem, placementItems, output);
            }

            if (layer.layers && layer.layers.length > 0) {
                appendItemTranslations(
                    layer.layers,
                    placementItems,
                    output,
                    excludeLockedHiddenLayers,
                    excludeLockedHiddenItems
                );
            }
        }
    }

    /* 移動対象から除外するレイヤーか判定する / Check whether a layer should be skipped for item translation */
    function shouldSkipLayerForItemMove(layer) {
        try {
            return layer.locked || !layer.visible;
        } catch (layerStateError) {
            return false;
        }
    }

    /* 移動対象から除外するオブジェクトか判定する / Check whether a page item should be skipped for translation */
    function shouldSkipPageItemForMove(pageItem) {
        try {
            return pageItem.locked || pageItem.hidden;
        } catch (pageItemStateError) {
            return false;
        }
    }

    /* 単一オブジェクトの移動量を追加する / Append a translation delta for a single item */
    function appendTranslationForItem(pageItem, placementItems, output) {
        var center = getItemGeometricCenter(pageItem);
        if (!center) return;

        for (var placementIndex = 0; placementIndex < placementItems.length; placementIndex++) {
            var placement = placementItems[placementIndex];
            if (isCenterInsideRect(center, placement.oldRect)) {
                output.push({ item: pageItem, deltaX: placement.deltaX, deltaY: placement.deltaY });
                break;
            }
        }
    }

    /* オブジェクトの幾何中心を取得 / Get an item's geometric center. 取得できない場合は null を返す */
    function getItemGeometricCenter(pageItem) {
        try {
            var bounds = pageItem.geometricBounds; // [left, top, right, bottom]
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (geometricBoundsError) {
            return null;
        }
    }

    /* 中心座標が矩形内にあるか / Whether a center point falls inside an artboard rect.
       artboardRect の Y は上がプラス、下がマイナス / Illustrator rects use Y positive upward */
    function isCenterInsideRect(center, rect) {
        return center[0] >= rect[0] && center[0] <= rect[2] &&
            center[1] <= rect[1] && center[1] >= rect[3];
    }

    /* アートボードパネル上の順序を、割り当てた (assignedRow, assignedColumn) の昇順に再構築する
       Rebuild the artboards in (assignedRow, assignedColumn) ascending order so the Artboards panel
       reflects the visual grid order (row 1 left→right, row 2 left→right, ...).
       一時名にリネームしてから新規作成→旧アートボードを削除することで名称重複を回避する
       Rename originals to temp names, add new artboards in the desired order, then remove the originals
       to avoid name collisions. */
    function reorderArtboardsByGridOrder(activeDocument, placementItems) {
        var sortedPlacements = placementItems.slice();
        sortedPlacements.sort(function (a, b) {
            if (a.assignedRow !== b.assignedRow) return a.assignedRow - b.assignedRow;
            return a.assignedColumn - b.assignedColumn;
        });

        var artboards = activeDocument.artboards;
        var originalCount = artboards.length;

        var targetData = [];
        for (var collectIndex = 0; collectIndex < sortedPlacements.length; collectIndex++) {
            var sourceArtboard = sortedPlacements[collectIndex].artboard;
            var sourceRect = sourceArtboard.artboardRect;
            targetData.push({
                name: sourceArtboard.name,
                rect: [sourceRect[0], sourceRect[1], sourceRect[2], sourceRect[3]]
            });
        }

        for (var renameIndex = 0; renameIndex < originalCount; renameIndex++) {
            artboards[renameIndex].name = '__reorder_tmp__' + renameIndex;
        }

        for (var addIndex = 0; addIndex < targetData.length; addIndex++) {
            var newArtboard = artboards.add(targetData[addIndex].rect);
            newArtboard.name = targetData[addIndex].name;
        }

        for (var removeIndex = originalCount - 1; removeIndex >= 0; removeIndex--) {
            artboards[removeIndex].remove();
        }

        activeDocument.artboards.setActiveArtboardIndex(0);
    }

    /* 各番号（列または行）の累積オフセットを計算する / Compute cumulative offset for each number (column or row).
       boundaryNumber と一致するキーの直前のギャップは3倍に広げる / Triple the gap right before boundaryNumber. */
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
