#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ReorderArtboardsByPosition.jsx

### 概要

- 見た目の配置を基準に、アートボードを左上（Y優先）で並べ替えます
- 許容差を調整しながら並び順を確認し、必要に応じて再配置できます
- アートボード名を「行-列」形式に自動更新できます

### 主な機能

- 左上基準（Y優先）でのアートボード順並べ替え
- 許容差の自動計算とスライダーによる調整
- 行単位の並びをリストで確認（例：1 | 2 | 3）
- パネル上の並び替えの ON / OFF を切り替え
- カンバス上の再配置モードを 2 つのチェックボックスで切り替え（排他的に選択）
  - 列数を指定して再配置
  - アートボード名（「行-列」または「接頭辞-番号」）から行列に再配置
- 再配置時は列間・行間（現在の定規単位）を指定可能、［連動］で同値入力
- 間隔入力欄に現在の定規単位を表示
- 列数・間隔は ↑↓キーで ±1、Shift + ↑↓キーで 10 の倍数にスナップ
- 名前から再配置時の未指定／重複アートボードの扱いを選択可能
  - 各行の末尾に配置
  - 最終行の次の行にまとめる
- アートボード名パネルでアートボード名を「行-列」形式に更新
  - 「配置位置から作成」または「既存名を整形」を選択可能
  - 区切り文字（- / _ / x）を選択可能
  - 桁数（0 / 00 / 000）を選択可能

### オリジナル、謝辞

- m1b 氏: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### 更新履歴

- v1.0 (20231115) : 初期バージョン（Andrew_BJ による UI 改良と上限拡張）
- v1.2.0 (20260415) : 左上基準専用ツールとしてUIと構成を整理、再配置設定とプレビュー表示を調整
- v1.3.0 (20260508) : アートボード名パネル追加（「行-列」形式の自動命名、既存名の整形、区切り文字／桁数の選択）、再配置ロジックとUI構築を責務ごとの関数に分割、L() 経由のローカライズに統一、列間／行間の連動処理を整理
- v1.3.1 (20260508) : 再配置とパネル並び順の実行順を入れ替え、再配置後の見た目に合わせてパネル順が更新されるよう修正
- v1.4.0 (20260513) : パネル上の並び順を「名前順／カンバス上の並び順に／変更しない」のラジオボタンに変更、名前順は数字列を10桁ゼロ埋めした自然順ソート

---

### Script Name:

ReorderArtboardsByPosition.jsx

### Overview

- Reorders artboards by visual layout using a Top Left (Y priority) order
- Preview and adjust row grouping tolerance before reordering
- Optionally update artboard names using a row-column format

### Main Features

- Reorder artboards using Top Left base order (Y priority)
- Auto-calculate tolerance and fine-tune it with a slider
- Preview row-based order in a list (e.g. 1 | 2 | 3)
- Toggle reorder execution in the Artboards panel on or off
- Choose a canvas rearrange mode with two mutually exclusive checkboxes
  - Rearrange by column count
  - Rearrange by row-column / prefix-number from artboard names
- When rearranging, set column and row gaps (current ruler unit), with optional linked values
- Show the current ruler unit next to the spacing input
- Up/Down keys change column count and spacing by 1; Shift + Up/Down snaps to multiples of 10
- Choose how unspecified / duplicate names are handled when rearranging by name
  - Append to each row end
  - Group in row after last row
- Artboard Names panel updates artboard names to a row-column format
  - Choose source: "Create from position" or "Reformat existing names"
  - Choose separator: - / _ / x
  - Choose digit width: 0 / 00 / 000

### Original / Credit

- m1b: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### Changelog

- v1.0 (20231115): Initial version (UI improvements and limit extension by Andrew_BJ)
- v1.2.0 (20260415): Refined the UI and structure as a dedicated Top Left reorder tool, and adjusted rearrange settings and preview display
- v1.3.0 (20260508): Added Artboard Names panel (auto row-column naming, reformat existing names, separator and digit width selection), split rearrange logic and UI construction into responsibility-based helper functions, unified localization via L(), and refined linked column/row gap behavior
- v1.3.1 (20260508): Swapped the execution order of rearrange and panel reorder so the panel order matches the post-rearrange visual layout
*/

(function () {

    // =========================================
    // バージョンとローカライズ、ユーティリティ / Version, localization, and utility functions  
    // =========================================

    var SCRIPT_VERSION = "v1.3.1";

    var COORDINATE_PRECISION_DIGITS = 3;
    var PANEL_MARGINS = [15, 20, 15, 10];

    /* byName 再配置で例外領域（未指定／重複）の境界に適用するギャップ倍率
     * Gap multiplier applied at the boundary into the unspecified/duplicate exception area in byName mode. */
    var EXCEPTION_BOUNDARY_GAP_MULTIPLIER = 3;

    /* 複製対象のアートボードプロパティ / Artboard properties to copy when duplicating */
    var ARTBOARD_COPYABLE_PROPS = ["name", "rulerOrigin", "rulerPAR", "showCenter", "showCrossHairs", "showSafeAreas"];

    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ローカライズ済みラベル取得 / Get localized label by key */
    function L(key) {
        var entry = LABELS[key];
        return entry ? entry[currentLanguage] : key;
    }

    /* -------------------------------------------------- */
    /* 日英ラベル定義 / Japanese-English label definitions */
    /* -------------------------------------------------- */

    var LABELS = {
        dialogTitle: {
            ja: "アートボードの並び順を整理",
            en: "Organize Artboard Order"
        },
        noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noArtboards: { ja: "アートボードが存在しません。", en: "No artboards found." },
        preview: { ja: "パネル上の並び順", en: "Artboards Panel Order" },
        execute: { ja: "OK", en: "OK" },
        close: { ja: "キャンセル", en: "Cancel" },
        rearrange: { ja: "カンバス上のアートボードを再配置", en: "Rearrange Canvas Artboards" },
        rearrangeByColumns: { ja: "列数を指定して再配置", en: "Rearrange by column count" },
        rearrangeByName: { ja: "アートボード名から行列に再配置", en: "Rearrange by row-column from names" },
        duplicateHandling: { ja: "未指定／重複", en: "Unspecified / Duplicate" },
        duplicateAppendToPrevRow: { ja: "各行の末尾に配置", en: "Append to each row end" },
        duplicateGroupInLastRow: { ja: "最終行の次の行にまとめる", en: "Group in row after last row" },
        columns: { ja: "列数：", en: "Columns:" },
        columnGap: { ja: "列間：", en: "Column gap:" },
        rowGap: { ja: "行間：", en: "Row gap:" },
        gapLink: { ja: "連動", en: "Link" },
        errorPrefix: { ja: "エラー: ", en: "Error: " },
        namingPanel: { ja: "アートボード名の更新", en: "Update names as row-column" },
        namingEnable: { ja: "「行-列」形式に更新", en: "Update" },
        namingFromPosition: { ja: "配置位置から作成", en: "Create from position" },
        namingFromExisting: { ja: "既存名を整形", en: "Reformat existing names" },
        namingSeparator: { ja: "区切り文字", en: "Separator" },
        namingPadWidth: { ja: "桁数", en: "Digits" },
        sortByPosition: { ja: "カンバス上の並び順に", en: "Match canvas order" },
        sortByName: { ja: "名前順", en: "By name" },
        sortKeepAsIs: { ja: "変更しない", en: "Keep as is" },
        noMatchAlert: {
            ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）形式のアートボード名が見つかりませんでした。",
            en: "No artboard names in '<row><sep><column>' (e.g., 1-2) or '<prefix><sep><number>' (e.g., banner-1) format were found."
        }
    };

    // =========================================
    // 単位コードとラベルのマッピング / Mapping of unit codes to labels
    // =========================================

    var unitLabelMap = {
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

    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return unitLabelMap[unitCode] || "pt";
    }
    var currentUnitLabel = getCurrentUnitLabel();

    /* 表示単位の数値をポイントに変換 / Convert display-unit value to points */
    function convertDisplayUnitToPoints(value) {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        switch (unitCode) {
            case 0: return value * 72;                  /* inch */
            case 1: return value * 72 / 25.4;           /* mm */
            case 2: return value;                       /* pt */
            case 3: return value * 12;                  /* pica */
            case 4: return value * 72 / 2.54;           /* cm */
            case 5: return value * 0.25 * 72 / 25.4;    /* Q/H */
            case 6: return value;                       /* px: Illustrator scripting commonly treats px as pt */
            case 7: return value * 72 * 12;             /* ft/in */
            case 8: return value * 72 / 0.0254;         /* m */
            case 9: return value * 72 * 36;             /* yd */
            case 10: return value * 72 * 12;             /* ft */
            default: return value;
        }
    }

    /* 入力欄の表示単位値を読み取り、pt に変換 / Read display-unit input and convert it to points */
    function readDisplayUnitInputAsPoints(input, defaultValue) {
        var value = parseFloat(input.text);
        if (isNaN(value)) value = defaultValue;
        return convertDisplayUnitToPoints(value);
    }

    // =========================================
    // メイン処理
    // =========================================

    function main() {
        if (app.documents.length === 0) {
            alert(L("noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (doc.artboards.length === 0) {
            alert(L("noArtboards"));
            return;
        }

        var previewContext = buildPreviewContext(doc);
        var ui = buildDialogUI(previewContext.defaultTolerance, previewContext.sliderMax);

        bindEvents(doc, previewContext, ui);

        ui.dialog.center();
        ui.dialog.show();
    }

    function buildPreviewContext(doc) {
        /* アートボード情報を配列に格納（プレビュー・自動計算用） */
        var artboardEntries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            artboardEntries.push({
                name: doc.artboards[artboardIndex].name,
                artboardRect: doc.artboards[artboardIndex].artboardRect
            });
        }

        /* スライダー最大値をアートボードの最大高さに設定 */
        var maxHeight = 0;
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            var artboardRect = artboardEntries[entryIndex].artboardRect;
            var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);
            if (artboardHeight > maxHeight) maxHeight = artboardHeight;
        }
        var sliderMax = Math.round(maxHeight);

        /* スライダー初期値を自動計算値×1.1に設定 */
        var autoTolerance = calculateAutoTolerance(artboardEntries);
        var defaultTolerance = Math.round(autoTolerance * 1.1);
        defaultTolerance = Math.min(defaultTolerance, sliderMax);

        return {
            artboardEntries: artboardEntries,
            sliderMax: sliderMax,
            defaultTolerance: defaultTolerance
        };
    }

    function bindEvents(doc, previewContext, ui) {
        var sortRadios = ui.preview.sortModeRadios;

        /* ラジオは親グループが異なるため、ScriptUI の自動排他が効かない。
         * 手動で他のラジオを OFF にして相互排他を成立させる。
         * Radios live in separate parent groups, so ScriptUI's built-in exclusion
         * does not apply. Toggle the others off manually to enforce exclusivity. */
        function selectSortMode(targetRadio) {
            sortRadios.byName.value = (targetRadio === sortRadios.byName);
            sortRadios.byPosition.value = (targetRadio === sortRadios.byPosition);
            sortRadios.keepAsIs.value = (targetRadio === sortRadios.keepAsIs);
            syncSortMode();
        }

        function syncSortMode() {
            var byPositionSelected = sortRadios.byPosition.value;
            ui.preview.toleranceSlider.enabled = byPositionSelected;
            var currentTolerance = Math.round(ui.preview.toleranceSlider.value);
            updateReorderPreview(previewContext, ui, currentTolerance);
            /* リスト更新後に enabled を当て直す（add の影響でディムが効かないケースを回避）
             * Re-apply enabled after add() to make dimming take effect reliably */
            ui.preview.reorderList.enabled = byPositionSelected;
        }

        syncSortMode();

        ui.preview.toleranceSlider.onChanging = function () {
            var tolerance = Math.round(ui.preview.toleranceSlider.value);
            updateReorderPreview(previewContext, ui, tolerance);
        };

        sortRadios.byPosition.onClick = function () { selectSortMode(sortRadios.byPosition); };
        sortRadios.byName.onClick = function () { selectSortMode(sortRadios.byName); };
        sortRadios.keepAsIs.onClick = function () { selectSortMode(sortRadios.keepAsIs); };

        ui.buttons.executeBtn.onClick = function () {
            executeReorder(doc, ui);
        };

        ui.buttons.closeBtn.onClick = function () {
            ui.dialog.close(-1);
        };
    }

    function executeReorder(doc, ui) {
        var tolerance = Math.round(ui.preview.toleranceSlider.value) || 0;

        /* 再配置を先に実行（入力値は表示単位として受け取り、ptに変換して渡す）
         * パネル並び順より先に再配置すると、パネル順が「再配置後の見た目」と揃う
         * Run rearrange first so the panel order reflects post-rearrange positions */
        if (ui.rearrange.modeChecks.byColumns.value) {
            var columns = parseInt(ui.rearrange.columnsInput.text, 10) || 4;
            var columnGapPoints = readDisplayUnitInputAsPoints(ui.rearrange.columnGapInput, 100);
            var rowGapPoints = ui.rearrange.gapLinkCheckbox.value
                ? columnGapPoints
                : readDisplayUnitInputAsPoints(ui.rearrange.rowGapInput, 100);
            try {
                rearrangeArtboardsWithGaps(doc, columns, columnGapPoints, rowGapPoints);
            } catch (e) {
                alert(L("errorPrefix") + e.message);
            }
        } else if (ui.rearrange.modeChecks.byName.value) {
            var byNameColumnGapPoints = readDisplayUnitInputAsPoints(ui.rearrange.columnGapInput, 100);
            var byNameRowGapPoints = ui.rearrange.gapLinkCheckbox.value
                ? byNameColumnGapPoints
                : readDisplayUnitInputAsPoints(ui.rearrange.rowGapInput, 100);
            var exceptionMode = ui.rearrange.duplicateRadios.groupLast.value ? 'lastRow' : 'rowEnd';
            try {
                rearrangeArtboardsByRowColumnName(doc, byNameColumnGapPoints, byNameRowGapPoints, exceptionMode);
            } catch (e) {
                alert(L("errorPrefix") + e.message);
            }
        }

        /* 並び順モードに応じてソーター切替（変更しない場合はスキップ）
         * Pick a sorter based on the selected sort mode; skip when "Keep as is" */
        if (!ui.preview.sortModeRadios.keepAsIs.value) {
            var sorter;
            if (ui.preview.sortModeRadios.byName.value) {
                sorter = function (_artboards) {
                    sortArtboardsByName(_artboards);
                };
            } else {
                sorter = function (_artboards, dp) {
                    sortArtboardsTopLeftWithTolerance(_artboards, dp, tolerance);
                };
            }
            rebuildArtboardsInSortedOrder(doc, sorter, COORDINATE_PRECISION_DIGITS);
        }

        /* 命名（配置位置から作成 / 既存名を整形） / Update or reformat artboard names as row-column */
        if (ui.naming.enableCheckbox.value) {
            var separators = ui.naming.separators;
            var separator = separators.underscore.value ? "_" : (separators.x.value ? "x" : "-");
            var padRadios = ui.naming.padRadios;
            var padWidth = padRadios.w3.value ? 3 : (padRadios.w2.value ? 2 : 1);
            if (ui.naming.source.fromPosition.value) {
                renameArtboardsFromPositions(doc, separator, padWidth, tolerance);
            } else {
                renameArtboardsFromExistingNames(doc, separator, padWidth);
            }
        }

        ui.dialog.close(1);
    }

    function updateReorderPreview(previewContext, ui, tolerance) {
        var reorderList = ui.preview.reorderList;
        reorderList.removeAll();

        /* 変更しないモード: 現在の並びをそのまま表示 / Keep-as-is mode: show current order untouched */
        if (ui.preview.sortModeRadios.keepAsIs.value) {
            for (var keepIndex = 0; keepIndex < previewContext.artboardEntries.length; keepIndex++) {
                reorderList.add("item", previewContext.artboardEntries[keepIndex].name);
            }
            return;
        }

        /* 名前順モード: アートボード名で並べてリスト表示 / By-name mode: list names in name-sort order */
        if (ui.preview.sortModeRadios.byName.value) {
            var byNameEntries = previewContext.artboardEntries.slice();
            sortArtboardsByName(byNameEntries);
            for (var nameIndex = 0; nameIndex < byNameEntries.length; nameIndex++) {
                reorderList.add("item", byNameEntries[nameIndex].name);
            }
            return;
        }

        var decimalPlaces = Math.pow(10, COORDINATE_PRECISION_DIGITS);
        var sortedEntries = previewContext.artboardEntries.slice();
        sortArtboardsTopLeftWithTolerance(sortedEntries, decimalPlaces, tolerance);

        var rowGroups = groupSortedIntoRows(sortedEntries, tolerance, decimalPlaces);

        for (var rowIndex = 0; rowIndex < rowGroups.length; rowIndex++) {
            var rowNames = [];
            for (var columnIndex = 0; columnIndex < rowGroups[rowIndex].length; columnIndex++) {
                rowNames.push(rowGroups[rowIndex][columnIndex].name);
            }
            reorderList.add("item", rowNames.join(" | "));
        }
    }

    // =========================================
    // ダイアログUI
    // =========================================

    function buildDialogUI(defaultTolerance, sliderMax) {
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ['fill', 'top'];
        dialog.spacing = 15;

        var mainColumn = dialog.add("group");
        mainColumn.orientation = "column";
        mainColumn.alignChildren = ['fill', 'top'];
        mainColumn.alignment = ['fill', 'top'];

        /* 並び順は上段にフル幅 / Order panel spans full width on top */
        var preview = buildPreviewPanel(mainColumn, defaultTolerance, sliderMax);

        /* 下段は 2 カラム（左: 再配置 / 右: 命名） / Two-column row: rearrange (left) / naming (right) */
        var twoColumnsRow = mainColumn.add("group");
        twoColumnsRow.orientation = "row";
        twoColumnsRow.alignChildren = ['fill', 'top'];
        twoColumnsRow.alignment = ['fill', 'top'];
        twoColumnsRow.spacing = 10;

        var leftColumn = twoColumnsRow.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ['fill', 'top'];
        leftColumn.alignment = ['fill', 'top'];

        var rightColumn = twoColumnsRow.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ['fill', 'top'];
        rightColumn.alignment = ['fill', 'top'];

        return {
            dialog: dialog,
            preview: preview,
            rearrange: buildRearrangePanel(leftColumn),
            naming: buildNamingPanel(rightColumn),
            buttons: buildDialogButtons(dialog)
        };
    }

    /* プレビューパネル（並び順モード＋許容差スライダー＋並び順リスト） / Preview panel */
    function buildPreviewPanel(parent, defaultTolerance, sliderMax) {
        var reorderPanel = parent.add("panel", undefined, L("preview"));
        setupPanel(reorderPanel);

        /* 名前順（ラジオ） / Radio "By name" */
        var byNameRow = reorderPanel.add("group");
        byNameRow.orientation = "row";
        byNameRow.alignChildren = ['left', 'center'];
        var sortByNameRadio = byNameRow.add("radiobutton", undefined, L("sortByName"));

        /* カンバス上の並び順に（ラジオ）＋ 許容差スライダーを同じ行に配置
         * Radio "Match canvas order" and tolerance slider on the same row */
        var byPositionRow = reorderPanel.add("group");
        byPositionRow.orientation = "row";
        byPositionRow.alignChildren = ['left', 'center'];

        var sortByPositionRadio = byPositionRow.add("radiobutton", undefined, L("sortByPosition"));
        sortByPositionRadio.value = true;

        var toleranceSlider = byPositionRow.add("slider", undefined, defaultTolerance, 0, sliderMax);
        toleranceSlider.preferredSize = [250, 20];

        /* 変更しない（ラジオ） / Radio "Keep as is" */
        var keepAsIsRow = reorderPanel.add("group");
        keepAsIsRow.orientation = "row";
        keepAsIsRow.alignChildren = ['left', 'center'];
        var sortKeepAsIsRadio = keepAsIsRow.add("radiobutton", undefined, L("sortKeepAsIs"));

        var reorderList = reorderPanel.add("listbox", undefined, [], { multiselect: false });
        reorderList.preferredSize.width = 160;
        reorderList.minimumSize.height = 120;
        reorderList.alignment = ['fill', 'fill'];

        return {
            sortModeRadios: {
                byPosition: sortByPositionRadio,
                byName: sortByNameRadio,
                keepAsIs: sortKeepAsIsRadio
            },
            toleranceSlider: toleranceSlider,
            reorderList: reorderList
        };
    }

    /* 再配置パネル / Rearrange panel */
    function buildRearrangePanel(parent) {
        var rearrangeGroup = parent.add("panel", undefined, L("rearrange"));
        setupPanel(rearrangeGroup);

        var modeChecks = buildRearrangeModeCheckboxes(rearrangeGroup);
        var rearrangeSettingsGroup = rearrangeGroup.add("group");
        rearrangeSettingsGroup.orientation = "column";
        rearrangeSettingsGroup.alignChildren = ['fill', 'top'];

        var columnsRow = addLabeledInput(rearrangeSettingsGroup, L("columns"), "4", 50);
        var spacingControls = buildRearrangeSpacingControls(rearrangeSettingsGroup);
        var duplicateRadios = buildDuplicateHandlingPanel(rearrangeGroup);

        function syncEnabled() {
            syncRearrangePanelState(modeChecks, rearrangeSettingsGroup, columnsRow, duplicateRadios.panel);
        }

        modeChecks.byColumns.onClick = function () {
            if (modeChecks.byColumns.value) modeChecks.byName.value = false;
            syncEnabled();
        };
        modeChecks.byName.onClick = function () {
            if (modeChecks.byName.value) modeChecks.byColumns.value = false;
            syncEnabled();
        };
        syncEnabled();

        return {
            modeChecks: modeChecks,
            columnsInput: columnsRow.input,
            columnGapInput: spacingControls.columnGapInput,
            rowGapInput: spacingControls.rowGapInput,
            gapLinkCheckbox: spacingControls.gapLinkCheckbox,
            duplicateRadios: {
                append: duplicateRadios.append,
                groupLast: duplicateRadios.groupLast
            }
        };
    }

    /* 再配置モードのチェックボックスを作成 / Build rearrange mode checkboxes */
    function buildRearrangeModeCheckboxes(parent) {
        /* 再配置モード（列数指定 / アートボード名から、排他的に選択） / Rearrange mode checkboxes (mutually exclusive) */
        var modeGroup = parent.add("group");
        modeGroup.orientation = "column";
        modeGroup.alignChildren = ['left', 'center'];

        var byColumnsCheckbox = modeGroup.add("checkbox", undefined, L("rearrangeByColumns"));
        var byNameCheckbox = modeGroup.add("checkbox", undefined, L("rearrangeByName"));
        byColumnsCheckbox.value = false;
        byNameCheckbox.value = false;

        return {
            byColumns: byColumnsCheckbox,
            byName: byNameCheckbox
        };
    }

    /* 再配置の列間／行間入力と連動チェックを作成 / Build gap inputs and link checkbox */
    function buildRearrangeSpacingControls(parent) {
        /* 列間／行間 + 連動チェックの2カラム / Two-column row: gap inputs (left) + link checkbox (right) */
        var gapsRow = parent.add("group");
        gapsRow.orientation = "row";
        gapsRow.alignChildren = ['left', 'center'];
        gapsRow.spacing = 20;

        var gapsLeft = gapsRow.add("group");
        gapsLeft.orientation = "column";
        gapsLeft.alignChildren = ['left', 'top'];

        var columnGapRow = addLabeledInput(gapsLeft, L("columnGap"), "100", 50);
        var columnGapInput = columnGapRow.input;
        columnGapRow.row.add("statictext", undefined, currentUnitLabel);

        var rowGapRow = addLabeledInput(gapsLeft, L("rowGap"), "100", 50);
        var rowGapInput = rowGapRow.input;
        rowGapRow.row.add("statictext", undefined, currentUnitLabel);

        var gapLinkCheckbox = gapsRow.add("checkbox", undefined, L("gapLink"));
        gapLinkCheckbox.value = true;

        bindGapLinkControls(columnGapInput, rowGapInput, rowGapRow.row, gapLinkCheckbox);

        return {
            columnGapInput: columnGapInput,
            rowGapInput: rowGapInput,
            gapLinkCheckbox: gapLinkCheckbox
        };
    }

    /* 列間／行間の連動挙動を設定 / Bind linked gap input behavior */
    function bindGapLinkControls(columnGapInput, rowGapInput, rowGapControlRow, gapLinkCheckbox) {
        function syncRowGapEnabled() {
            rowGapControlRow.enabled = !gapLinkCheckbox.value;
        }

        columnGapInput.onChange = function () {
            if (gapLinkCheckbox.value) rowGapInput.text = columnGapInput.text;
        };
        rowGapInput.onChange = function () {
            if (gapLinkCheckbox.value) columnGapInput.text = rowGapInput.text;
        };
        gapLinkCheckbox.onClick = function () {
            if (gapLinkCheckbox.value) rowGapInput.text = columnGapInput.text;
            syncRowGapEnabled();
        };

        syncRowGapEnabled();
    }

    /* 未指定／重複の扱いパネルを作成 / Build duplicate handling panel */
    function buildDuplicateHandlingPanel(parent) {
        var duplicatePanel = parent.add("panel", undefined, L("duplicateHandling"));
        setupPanel(duplicatePanel);

        var duplicateAppendRadio = duplicatePanel.add("radiobutton", undefined, L("duplicateAppendToPrevRow"));
        var duplicateGroupRadio = duplicatePanel.add("radiobutton", undefined, L("duplicateGroupInLastRow"));
        duplicateAppendRadio.value = true;

        return {
            panel: duplicatePanel,
            append: duplicateAppendRadio,
            groupLast: duplicateGroupRadio
        };
    }

    /* 再配置パネルの有効／無効状態を同期 / Sync enabled state for rearrange panel controls */
    function syncRearrangePanelState(modeChecks, rearrangeSettingsGroup, columnsRow, duplicatePanel) {
        var hasRearrangeMode = modeChecks.byColumns.value || modeChecks.byName.value;
        rearrangeSettingsGroup.enabled = hasRearrangeMode;
        /* 列数は「列数を指定して再配置」のときだけアクティブ
         * (byName では名前から行列を取るので列数は不要) / Columns input is enabled only for byColumns */
        columnsRow.row.enabled = modeChecks.byColumns.value;
        /* 未指定／重複は byName のときだけアクティブ / Duplicate panel is active only for byName */
        duplicatePanel.enabled = modeChecks.byName.value;
    }

    /* 命名パネル / Naming panel */
    function buildNamingPanel(parent) {
        var namingGroup = parent.add("panel", undefined, L("namingPanel"));
        setupPanel(namingGroup);

        var namingEnableCheckbox = namingGroup.add("checkbox", undefined, L("namingEnable"));
        namingEnableCheckbox.value = false;

        var namingSettingsGroup = namingGroup.add("group");
        namingSettingsGroup.orientation = "column";
        namingSettingsGroup.alignChildren = ['fill', 'top'];

        /* 命名ソース / Naming source */
        var sourceRow = namingSettingsGroup.add("group");
        sourceRow.orientation = "column";
        sourceRow.alignChildren = ['left', 'center'];
        var fromPositionRadio = sourceRow.add("radiobutton", undefined, L("namingFromPosition"));
        var fromExistingRadio = sourceRow.add("radiobutton", undefined, L("namingFromExisting"));
        fromPositionRadio.value = true;

        /* 区切り文字パネルと桁数パネルを横並び / Separator and digits panels side by side */
        var sepPadRow = namingSettingsGroup.add("group");
        sepPadRow.orientation = "row";
        sepPadRow.alignChildren = ['fill', 'top'];
        sepPadRow.spacing = 10;

        /* 区切り文字 / Separator */
        var sepPanel = sepPadRow.add("panel", undefined, L("namingSeparator"));
        setupPanel(sepPanel);
        var sepRow = sepPanel.add("group");
        sepRow.orientation = "column";
        sepRow.alignChildren = ['left', 'center'];
        var sepHyphen = sepRow.add("radiobutton", undefined, "-");
        var sepUnderscore = sepRow.add("radiobutton", undefined, "_");
        var sepX = sepRow.add("radiobutton", undefined, "x");
        sepHyphen.value = true;

        /* ゼロ埋め桁数（0 / 00 / 000） / Zero-pad width options (0 / 00 / 000) */
        var padPanel = sepPadRow.add("panel", undefined, L("namingPadWidth"));
        setupPanel(padPanel);
        var padRow = padPanel.add("group");
        padRow.orientation = "column";
        padRow.alignChildren = ['left', 'center'];
        var pad1Radio = padRow.add("radiobutton", undefined, "0");
        var pad2Radio = padRow.add("radiobutton", undefined, "00");
        var pad3Radio = padRow.add("radiobutton", undefined, "000");
        pad1Radio.value = true;

        namingEnableCheckbox.onClick = function () {
            namingSettingsGroup.enabled = namingEnableCheckbox.value;
        };
        namingSettingsGroup.enabled = namingEnableCheckbox.value;

        return {
            enableCheckbox: namingEnableCheckbox,
            source: {
                fromPosition: fromPositionRadio,
                fromExisting: fromExistingRadio
            },
            separators: {
                hyphen: sepHyphen,
                underscore: sepUnderscore,
                x: sepX
            },
            padRadios: {
                w1: pad1Radio,
                w2: pad2Radio,
                w3: pad3Radio
            }
        };
    }

    /* キャンセル / OK ボタン / Dialog buttons */
    function buildDialogButtons(dialog) {
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = ['right', 'bottom'];

        var closeBtn = buttonGroup.add('button', undefined, L("close"));
        var executeBtn = buttonGroup.add('button', undefined, L("execute"));

        var btnWidth = Math.max(executeBtn.preferredSize.width, closeBtn.preferredSize.width);
        executeBtn.preferredSize.width = btnWidth;
        closeBtn.preferredSize.width = btnWidth;

        return {
            closeBtn: closeBtn,
            executeBtn: executeBtn
        };
    }

    /* ラベル付き edittext 行を追加 / Add a labeled edittext row */
    function addLabeledInput(parent, labelText, defaultValue, labelWidth) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignChildren = ['left', 'center'];
        var label = row.add("statictext", undefined, labelText);
        label.preferredSize.width = labelWidth;
        label.justify = 'right';
        var input = row.add("edittext", undefined, defaultValue);
        input.characters = 5;
        changeValueByArrowKey(input);
        return { row: row, input: input };
    }

    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var isShift = ScriptUI.environment.keyboardState.shiftKey;
            var direction = (event.keyName === "Up") ? 1 : -1;

            if (isShift) {
                // Shift+矢印キーで 10 の倍数にスナップ / Snap to multiples of 10
                value = (direction > 0)
                    ? Math.ceil((value + 1) / 10) * 10
                    : Math.floor((value - 1) / 10) * 10;
            } else {
                value += direction;
            }
            if (value < 0) value = 0;

            editText.text = Math.round(value);
            editText.notify("onChange");
            event.preventDefault();
        });
    }

    // =========================================
    // ソート処理
    // =========================================

    /* src のプロパティを dst にコピー / Copy listed properties from src to dst */
    function copyArtboardProps(sourceArtboard, targetArtboard) {
        for (var propertyIndex = 0; propertyIndex < ARTBOARD_COPYABLE_PROPS.length; propertyIndex++) {
            var propertyName = ARTBOARD_COPYABLE_PROPS[propertyIndex];
            targetArtboard[propertyName] = sourceArtboard[propertyName];
        }
    }

    /* アートボードを並べ替えて［アートボード］パネル順を再構築 / Rebuild Artboards panel order based on sorted result */
    function rebuildArtboardsInSortedOrder(doc, sorterFunction, precisionDigits) {
        var decimalPlaces = Math.pow(10, precisionDigits || 3);

        /* 元アートボードのスナップショット（プロパティ＋rect）を取得 / Snapshot original artboards */
        var artboardSnapshots = [];
        var sortableEntries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var sourceArtboard = doc.artboards[artboardIndex];
            var snapshot = { artboardRect: sourceArtboard.artboardRect };
            copyArtboardProps(sourceArtboard, snapshot);
            artboardSnapshots.push(snapshot);
            sortableEntries.push({
                artboardRect: sourceArtboard.artboardRect,
                sourceIndex: artboardIndex,
                name: sourceArtboard.name
            });
        }

        sorterFunction(sortableEntries, decimalPlaces);

        /* ソート結果順に末尾へ複製を追加 / Append duplicates in sorted order */
        for (var sortedIndex = 0; sortedIndex < sortableEntries.length; sortedIndex++) {
            var sourceSnapshot = artboardSnapshots[sortableEntries[sortedIndex].sourceIndex];
            var newArtboard = doc.artboards.add(sourceSnapshot.artboardRect);
            copyArtboardProps(sourceSnapshot, newArtboard);
        }

        /* 元のアートボードを後ろから削除 / Remove original artboards from back to front */
        for (var removeIndex = artboardSnapshots.length - 1; removeIndex >= 0; removeIndex--) {
            doc.artboards[removeIndex].remove();
        }
    }

    /* 許容差付き 左上基準ソート（上辺降順 → tolerance 内なら左辺昇順）
     * Top Left ordering with tolerance (top edge desc, then left edge asc) */
    function sortArtboardsTopLeftWithTolerance(artboardEntries, decimalPlaces, tolerance) {
        decimalPlaces = decimalPlaces || 1000;
        artboardEntries.sort(function (firstEntry, secondEntry) {
            var firstTop = Math.round(firstEntry.artboardRect[1] * decimalPlaces) / decimalPlaces;
            var secondTop = Math.round(secondEntry.artboardRect[1] * decimalPlaces) / decimalPlaces;
            if (Math.abs(firstTop - secondTop) <= tolerance) {
                var firstLeft = Math.round(firstEntry.artboardRect[0] * decimalPlaces) / decimalPlaces;
                var secondLeft = Math.round(secondEntry.artboardRect[0] * decimalPlaces) / decimalPlaces;
                return firstLeft - secondLeft;
            }
            return secondTop - firstTop;
        });
    }

    /* 名前順（自然順）ソート: 数字を10桁ゼロ埋めして比較、大文字小文字は無視
     * Natural sort by artboard name: pad digits to 10 zeros, compare case-insensitively. */
    function sortArtboardsByName(artboardEntries) {
        artboardEntries.sort(function (firstEntry, secondEntry) {
            var firstName = padNumbersForNaturalSort((firstEntry.name || "").toLowerCase());
            var secondName = padNumbersForNaturalSort((secondEntry.name || "").toLowerCase());
            if (firstName < secondName) return -1;
            if (firstName > secondName) return 1;
            return 0;
        });
    }

    /* 文字列中の数字列を10桁ゼロ埋め / Zero-pad each run of digits in the string to 10 chars */
    function padNumbersForNaturalSort(text) {
        return text.replace(/\d+/g, function (digitRun) {
            var pad = "0000000000";
            return (pad + digitRun).slice(-pad.length);
        });
    }

    /* ソート済みエントリを上辺の差で行に分割 / Group sorted entries into rows by top-edge tolerance */
    function groupSortedIntoRows(sortedEntries, tolerance, decimalPlaces) {
        if (sortedEntries.length === 0) return [];
        var rows = [[sortedEntries[0]]];
        for (var entryIndex = 1; entryIndex < sortedEntries.length; entryIndex++) {
            var previousTop = Math.round(sortedEntries[entryIndex - 1].artboardRect[1] * decimalPlaces) / decimalPlaces;
            var currentTop = Math.round(sortedEntries[entryIndex].artboardRect[1] * decimalPlaces) / decimalPlaces;
            if (Math.abs(previousTop - currentTop) > tolerance) {
                rows.push([sortedEntries[entryIndex]]);
            } else {
                rows[rows.length - 1].push(sortedEntries[entryIndex]);
            }
        }
        return rows;
    }

    // =========================================
    // 再配置処理（列間／行間を独立指定） / Rearrange with separate column / row gaps
    // =========================================

    /* 列間を spacing として GridByRow で初回レイアウトし、行ごとに追加オフセットを適用
     * Lay out with columnGap as base spacing, then apply extra row offsets per row */
    function rearrangeArtboardsWithGaps(doc, columns, columnGapPoints, rowGapPoints) {
        doc.rearrangeArtboards(DocumentArtboardLayout.GridByRow, columns, columnGapPoints, true);

        var artboardCount = doc.artboards.length;
        var rowCount = Math.ceil(artboardCount / columns);
        if (rowCount <= 1) return;

        var extraRowGap = rowGapPoints - columnGapPoints;
        if (extraRowGap === 0) return;

        /* 各アートボードの追加シフト量を placements として組み立てる
         * Build placements: each artboard gets deltaY = -row * extraRowGap (Y up). */
        var placements = [];
        for (var artboardIndex = 0; artboardIndex < artboardCount; artboardIndex++) {
            var rect = doc.artboards[artboardIndex].artboardRect;
            var rowShift = Math.floor(artboardIndex / columns) * extraRowGap;
            placements.push({
                artboard: doc.artboards[artboardIndex],
                oldRect: [rect[0], rect[1], rect[2], rect[3]],
                deltaX: 0,
                deltaY: -rowShift
            });
        }

        /* アートボード上のアイテムを共通 helper で移動 / Translate items via shared helper */
        applyItemTranslations(doc, placements);

        /* アートボードの rect 自体を下方向にシフト / Shift artboard rects down */
        for (var placementIndex = 0; placementIndex < placements.length; placementIndex++) {
            var deltaY = placements[placementIndex].deltaY;
            if (deltaY === 0) continue;
            var oldRect = placements[placementIndex].oldRect;
            placements[placementIndex].artboard.artboardRect = [oldRect[0], oldRect[1] + deltaY, oldRect[2], oldRect[3] + deltaY];
        }
    }

    // =========================================
    // 再配置処理（行-列名から） / Rearrange by row-column / prefix-number names
    // =========================================

    /* アートボード名（「行-列」または「接頭辞-番号」）に従ってグリッド状に再配置
     * Rearrange artboards as a grid based on '<row><sep><col>' or '<prefix><sep><num>' names.
     * 区切り文字は -, _, x（大小文字無視）/ Separators: -, _, x (case-insensitive)
     * exceptionMode: 'rowEnd' = 各行の末尾に配置 / Append to each row end
     *                'lastRow' = 最終行の次の行にまとめる / Collect in row after last matched row */
    function rearrangeArtboardsByRowColumnName(doc, columnGapPoints, rowGapPoints, exceptionMode) {
        var placementContext = parseArtboardNamePlacements(doc.artboards);
        if (!placementContext.hasMatchedArtboard) {
            alert(L("noMatchAlert"));
            return;
        }

        assignArtboardPlacementSlots(placementContext, exceptionMode);
        var placementMetrics = collectPlacementMetrics(placementContext.placementItems);
        computePlacementOffsets(placementContext, placementMetrics, columnGapPoints, rowGapPoints);
        applyComputedArtboardPositions(doc, placementContext.placementItems);

        app.executeMenuCommand('fitall');
    }

    /* アートボード名を解析して配置候補を作成 / Parse artboard names and build placement candidates */
    function parseArtboardNamePlacements(artboards) {
        var maxColumnNumber = 0;
        var maxRowNumber = 0;
        var placementItems = [];
        var hasMatchedArtboard = false;
        var prefixOrder = [];
        var prefixRowOffsetByName = {};

        for (var artboardIndex = 0; artboardIndex < artboards.length; artboardIndex++) {
            var artboard = artboards[artboardIndex];
            var artboardName = artboard.name;
            var artboardRect = artboard.artboardRect;
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];
            var placementItem = createPlacementItem(artboard, artboardWidth, artboardHeight);

            if (applyRowColumnNameMatch(placementItem, artboardName)) {
                if (placementItem.columnNumber > maxColumnNumber) maxColumnNumber = placementItem.columnNumber;
                if (placementItem.rowNumber > maxRowNumber) maxRowNumber = placementItem.rowNumber;
                hasMatchedArtboard = true;
            } else if (applyPrefixNumberNameMatch(placementItem, artboardName, prefixOrder, prefixRowOffsetByName)) {
                if (placementItem.columnNumber > maxColumnNumber) maxColumnNumber = placementItem.columnNumber;
                hasMatchedArtboard = true;
            }
            placementItems.push(placementItem);
        }

        return {
            maxColumnNumber: maxColumnNumber,
            maxRowNumber: maxRowNumber,
            maxMatchedRowNumber: maxRowNumber + prefixOrder.length,
            placementItems: placementItems,
            hasMatchedArtboard: hasMatchedArtboard,
            prefixOrder: prefixOrder,
            prefixRowOffsetByName: prefixRowOffsetByName
        };
    }

    /* 配置候補の初期オブジェクトを作成 / Create an initial placement candidate object */
    function createPlacementItem(artboard, artboardWidth, artboardHeight) {
        return {
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
    }

    /* 「行-列」形式の名前を配置候補に反映 / Apply a row-column name match to a placement candidate */
    function applyRowColumnNameMatch(placementItem, artboardName) {
        var rowColumnMatch = artboardName.match(/^(\d+)[-_x](\d+)$/i);
        if (!rowColumnMatch) return false;

        placementItem.rowNumber = parseInt(rowColumnMatch[1], 10);
        placementItem.columnNumber = parseInt(rowColumnMatch[2], 10);
        if (placementItem.rowNumber < 1 || placementItem.columnNumber < 1) return false;

        placementItem.matched = true;
        placementItem.matchType = 'rowColumn';
        return true;
    }

    /* 「接頭辞-番号」形式の名前を配置候補に反映 / Apply a prefix-number name match to a placement candidate */
    function applyPrefixNumberNameMatch(placementItem, artboardName, prefixOrder, prefixRowOffsetByName) {
        var prefixNumberMatch = artboardName.match(/^(.+?)[-_x](\d+)$/i);
        if (!prefixNumberMatch || /^\d+$/.test(prefixNumberMatch[1])) return false;

        placementItem.prefixName = prefixNumberMatch[1].toLowerCase();
        placementItem.columnNumber = parseInt(prefixNumberMatch[2], 10);
        if (placementItem.columnNumber < 1) return false;

        placementItem.matched = true;
        placementItem.matchType = 'prefixNumber';
        if (prefixRowOffsetByName[placementItem.prefixName] === undefined) {
            prefixRowOffsetByName[placementItem.prefixName] = prefixOrder.length + 1;
            prefixOrder.push(placementItem.prefixName);
        }
        return true;
    }

    /* 行・列番号を割り当てる / Assign row and column numbers */
    function assignArtboardPlacementSlots(placementContext, exceptionMode) {
        var occupiedSlots = {};
        var nextExceptionColumnByRow = {};
        var currentRowNumber = 1;
        var lastRowExceptionColumn = 1;
        var exceptionRowNumber = placementContext.maxMatchedRowNumber + 1;
        var placementItems = placementContext.placementItems;

        for (var assignIndex = 0; assignIndex < placementItems.length; assignIndex++) {
            var assignTarget = placementItems[assignIndex];

            if (assignTarget.matched) {
                currentRowNumber = assignMatchedPlacementSlot(assignTarget, placementContext, occupiedSlots, nextExceptionColumnByRow);
            } else if (exceptionMode === 'lastRow') {
                lastRowExceptionColumn = assignLastRowExceptionSlot(assignTarget, exceptionRowNumber, lastRowExceptionColumn, occupiedSlots);
            } else {
                assignRowEndExceptionSlot(assignTarget, currentRowNumber, placementContext.maxColumnNumber, occupiedSlots, nextExceptionColumnByRow);
            }
        }
    }

    /* 一致した名前の配置スロットを割り当てる / Assign a placement slot for a matched name */
    function assignMatchedPlacementSlot(assignTarget, placementContext, occupiedSlots, nextExceptionColumnByRow) {
        var targetRowNumber = assignTarget.rowNumber;
        if (assignTarget.matchType === 'prefixNumber') {
            targetRowNumber = placementContext.maxRowNumber + placementContext.prefixRowOffsetByName[assignTarget.prefixName];
        }

        var requestedKey = targetRowNumber + ',' + assignTarget.columnNumber;
        if (!occupiedSlots[requestedKey]) {
            occupiedSlots[requestedKey] = true;
            assignTarget.assignedRow = targetRowNumber;
            assignTarget.assignedColumn = assignTarget.columnNumber;
        } else {
            assignTarget.assignedRow = targetRowNumber;
            assignTarget.assignedColumn = reserveNextColumnAtRowEnd(
                targetRowNumber, occupiedSlots, nextExceptionColumnByRow, placementContext.maxColumnNumber
            );
        }
        return targetRowNumber;
    }

    /* 最終行の次の行にまとめる例外スロットを割り当てる / Assign an exception slot in the row after the last matched row */
    function assignLastRowExceptionSlot(assignTarget, exceptionRowNumber, lastRowExceptionColumn, occupiedSlots) {
        assignTarget.assignedRow = exceptionRowNumber;
        while (true) {
            var lastRowColumnCandidate = lastRowExceptionColumn++;
            var lastRowKey = exceptionRowNumber + ',' + lastRowColumnCandidate;
            if (!occupiedSlots[lastRowKey]) {
                occupiedSlots[lastRowKey] = true;
                assignTarget.assignedColumn = lastRowColumnCandidate;
                break;
            }
        }
        return lastRowExceptionColumn;
    }

    /* 各行の末尾に置く例外スロットを割り当てる / Assign an exception slot at each row end */
    function assignRowEndExceptionSlot(assignTarget, currentRowNumber, maxColumnNumber, occupiedSlots, nextExceptionColumnByRow) {
        assignTarget.assignedRow = currentRowNumber;
        assignTarget.assignedColumn = reserveNextColumnAtRowEnd(
            currentRowNumber, occupiedSlots, nextExceptionColumnByRow, maxColumnNumber
        );
    }

    /* 列ごとの最大幅・行ごとの最大高さを集計 / Collect per-column max widths and per-row max heights */
    function collectPlacementMetrics(placementItems) {
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

        return {
            columnWidthByNumber: columnWidthByNumber,
            rowHeightByNumber: rowHeightByNumber,
            usedColumnNumbers: collectSortedNumericKeys(columnWidthByNumber),
            usedRowNumbers: collectSortedNumericKeys(rowHeightByNumber)
        };
    }

    /* 累積オフセットと移動量を計算 / Compute cumulative offsets and movement deltas */
    function computePlacementOffsets(placementContext, placementMetrics, columnGapPoints, rowGapPoints) {
        var columnOffsetByNumber = computeCumulativeOffsets(
            placementMetrics.usedColumnNumbers,
            placementMetrics.columnWidthByNumber,
            columnGapPoints,
            placementContext.maxColumnNumber + 1
        );
        var rowOffsetByNumber = computeCumulativeOffsets(
            placementMetrics.usedRowNumbers,
            placementMetrics.rowHeightByNumber,
            rowGapPoints,
            placementContext.maxMatchedRowNumber + 1
        );

        var originX = 0;
        var originY = 0;
        var placementItems = placementContext.placementItems;
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
    }

    /* オブジェクト移動とアートボード矩形更新を適用 / Apply item translations and update artboard rectangles */
    function applyComputedArtboardPositions(doc, placementItems) {
        applyItemTranslations(doc, placementItems);

        for (var applyIndex = 0; applyIndex < placementItems.length; applyIndex++) {
            var applyItem = placementItems[applyIndex];
            var newRight = applyItem.newLeft + applyItem.width;
            var newBottom = applyItem.newTop - applyItem.height;
            applyItem.artboard.artboardRect = [applyItem.newLeft, applyItem.newTop, newRight, newBottom];
        }
    }

    /* 指定行の末尾（maxColumnNumber + n）で空きスロットを予約
     * Reserve the next empty slot at the end of the given row */
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

    /* 数値キーを昇順で取り出す / Collect numeric keys in ascending order */
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

    /* 累積オフセットを計算（boundaryNumber と一致するキーの直前の間隔は EXCEPTION_BOUNDARY_GAP_MULTIPLIER 倍）
     * Compute cumulative offsets; the gap right before boundaryNumber is multiplied by EXCEPTION_BOUNDARY_GAP_MULTIPLIER. */
    function computeCumulativeOffsets(sortedNumbers, sizeByNumber, gap, boundaryNumber) {
        var offsetByNumber = {};
        var cumulative = 0;
        for (var orderIndex = 0; orderIndex < sortedNumbers.length; orderIndex++) {
            var currentNumber = sortedNumbers[orderIndex];
            if (orderIndex > 0) {
                var currentGap = gap;
                if (currentNumber === boundaryNumber) currentGap *= EXCEPTION_BOUNDARY_GAP_MULTIPLIER;
                cumulative += currentGap;
            }
            offsetByNumber[currentNumber] = cumulative;
            cumulative += sizeByNumber[currentNumber];
        }
        return offsetByNumber;
    }

    /* placements（{artboard, oldRect, deltaX, deltaY}）に従ってアートボード上のアイテムを移動
     * Translate items sitting on each artboard by the corresponding delta in placements. */
    function applyItemTranslations(doc, placements) {
        var translations = collectItemTranslations(doc.layers, placements);
        for (var translateIndex = 0; translateIndex < translations.length; translateIndex++) {
            var pendingMove = translations[translateIndex];
            if (pendingMove.deltaX === 0 && pendingMove.deltaY === 0) continue;
            try {
                pendingMove.item.translate(pendingMove.deltaX, pendingMove.deltaY);
            } catch (translateError) { /* 念のためのフォールバック / Defensive fallback */ }
        }
    }

    /* レイヤーを再帰的に走査してアートボード上のオブジェクトの移動量を集める
     * Walk layers recursively and collect translation deltas for items on artboards */
    function collectItemTranslations(layerCollection, placementItems) {
        var collected = [];
        appendItemTranslations(layerCollection, placementItems, collected);
        return collected;
    }

    function appendItemTranslations(layerCollection, placementItems, output) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
                var pageItem = layer.pageItems[itemIndex];
                if (pageItem.parent !== layer) continue;
                /* ロックされたアイテムはスキップ / Skip locked items */
                try { if (pageItem.locked) continue; } catch (lockedReadError) { /* プロパティが取得不能なら通常通り扱う */ }
                appendTranslationForItem(pageItem, placementItems, output);
            }
            if (layer.layers && layer.layers.length > 0) {
                appendItemTranslations(layer.layers, placementItems, output);
            }
        }
    }

    /* オブジェクトの幾何中心が含まれる元アートボードを特定し、移動量を1件追加
     * Determine the original artboard that contains the item center, then queue a delta */
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

    /* オブジェクトの幾何中心。取得できない場合は null
     * Item geometric center, or null if unavailable */
    function getItemGeometricCenter(pageItem) {
        try {
            var bounds = pageItem.geometricBounds; /* [left, top, right, bottom] */
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (geometricBoundsError) {
            return null;
        }
    }

    /* 中心座標が矩形内か（Illustrator の Y は上方向が正）
     * Whether the center is inside the rect (Y positive upward) */
    function isCenterInsideRect(center, rect) {
        return center[0] >= rect[0] && center[0] <= rect[2] &&
            center[1] <= rect[1] && center[1] >= rect[3];
    }

    // =========================================
    // リネーム処理
    // =========================================

    /* 数値をゼロ埋め / Zero-pad a number */
    function padNumber(num, width) {
        var s = String(num);
        while (s.length < width) s = "0" + s;
        return s;
    }

    /* 行-列形式の名前を組み立てる / Format a row-column style name */
    function formatRowColName(rowNum, colNum, separator, padWidth) {
        return padNumber(rowNum, padWidth) + separator + padNumber(colNum, padWidth);
    }

    /* 物理的な配置から 行-列 名を割り当てる / Assign row-column names based on physical positions */
    function renameArtboardsFromPositions(doc, separator, padWidth, tolerance) {
        var dp = Math.pow(10, COORDINATE_PRECISION_DIGITS);

        var entries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            entries.push({
                index: artboardIndex,
                artboardRect: doc.artboards[artboardIndex].artboardRect
            });
        }
        sortArtboardsTopLeftWithTolerance(entries, dp, tolerance);

        var rows = groupSortedIntoRows(entries, tolerance, dp);
        for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            for (var columnIndex = 0; columnIndex < rows[rowIndex].length; columnIndex++) {
                doc.artboards[rows[rowIndex][columnIndex].index].name = formatRowColName(rowIndex + 1, columnIndex + 1, separator, padWidth);
            }
        }
    }

    /* 既存の 行-列 名を再フォーマット / Reformat existing row-column names */
    function renameArtboardsFromExistingNames(doc, separator, padWidth) {
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboard = doc.artboards[artboardIndex];
            var match = artboard.name.match(/^(\d+)[-_x](\d+)(.*)$/i);
            if (!match) continue;
            var rowNumber = parseInt(match[1], 10);
            var columnNumber = parseInt(match[2], 10);
            var rest = match[3] || "";
            artboard.name = formatRowColName(rowNumber, columnNumber, separator, padWidth) + rest;
        }
    }

    /* 許容差自動計算関数 / Auto calculate tolerance function */
    function calculateAutoTolerance(artboardEntries) {
        var tops = [];
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            tops.push(artboardEntries[entryIndex].artboardRect[1]);
        }
        tops.sort(function (a, b) {
            return b - a;
        }); /* 上から下に並べる / Sort top to bottom */

        var diffs = [];
        for (var entryIndex = 1; entryIndex < tops.length; entryIndex++) {
            var diff = Math.abs(tops[entryIndex] - tops[entryIndex - 1]);
            if (diff > 0) {
                diffs.push(diff);
            }
        }

        if (diffs.length === 0) {
            return 5; /* 差が無ければデフォルト / Default if no difference */
        }

        var minDiff = Math.min.apply(null, diffs);
        return minDiff + 2; /* 少しマージンを加える / Add some margin */
    }
    main();
})();