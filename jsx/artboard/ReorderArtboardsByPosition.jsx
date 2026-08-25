#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

カンバス上の見た目の並び（左上基準・Y優先）をもとに、［アートボード］パネルの並び順を整理します。
列数指定やアートボード名にもとづくカンバス上の再配置、「行-列」形式へのリネームも実行できます。

詳細は README を参照してください。

### Overview

Reorders the Artboards panel to match the visual arrangement on the canvas, working from the top-left with Y taking priority.
It can also rearrange the artboards on the canvas by column count or by name, and rename them into a "row-column" form.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReorderArtboardsByPosition";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2023-11-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-07";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReorderArtboardsByPosition.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReorderArtboardsByPosition.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nb416cb01728a"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 座標比較に使う丸め桁数 / Rounding digits used when comparing coordinates */
    var COORDINATE_PRECISION_DIGITS = 3;

    /* 許容差の自動計算に使う係数と下限 / Factors used to auto-calculate the row tolerance */
    var TOLERANCE_AUTO_MARGIN_RATIO  = 1.1; /* 自動計算値に掛ける倍率 / multiplier applied to the auto value */
    var TOLERANCE_AUTO_MARGIN_POINTS = 2;   /* 最小差に加えるマージン（pt） / margin added to the smallest gap (pt) */
    var TOLERANCE_FALLBACK_POINTS    = 5;   /* 上辺に差がないときの既定値（pt） / fallback when all top edges match (pt) */

    /* 再配置ダイアログの初期値と下限 / Initial values and minimums for the rearrange settings */
    var DEFAULT_COLUMN_COUNT = 4;   /* 列数 / column count */
    var MIN_COLUMN_COUNT     = 1;   /* 列数の下限 / minimum column count */
    var DEFAULT_GAP_VALUE    = 100; /* 列間・行間（表示単位） / column and row gap in the display unit */
    var MIN_GAP_VALUE        = 0;   /* 列間・行間の下限 / minimum gap value */

    /* byName 再配置で例外領域（未指定／重複）の境界に適用するギャップ倍率
     * Gap multiplier applied at the boundary into the unspecified/duplicate exception area in byName mode. */
    var EXCEPTION_BOUNDARY_GAP_MULTIPLIER = 3;

    /* 複製対象のアートボードプロパティ / Artboard properties to copy when duplicating */
    var ARTBOARD_COPYABLE_PROPS = ["name", "rulerOrigin", "rulerPAR", "showCenter", "showCrossHairs", "showSafeAreas"];

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var DENSE_SPACING  = 6;                  /* ラジオを並べる密なパネルの間隔 / spacing for dense radio panels */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* コントロールの寸法 / Control metrics */
    var FIELD_LABEL_WIDTH   = 50;  /* 入力欄ラベルの幅 / labeled field label width */
    var FIELD_INPUT_CHARS   = 5;   /* 入力欄の文字数 / edittext width in characters */
    var SLIDER_WIDTH        = 250; /* 許容差スライダーの幅 / tolerance slider width */
    var SLIDER_HEIGHT       = 20;  /* 許容差スライダーの高さ / tolerance slider height */
    var PREVIEW_LIST_WIDTH  = 160; /* 並び順リストの幅 / reorder list width */
    var PREVIEW_LIST_HEIGHT = 120; /* 並び順リストの最小高さ / reorder list minimum height */

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
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
     * 行グループの共通設定を適用する
     * @param {Group} group - 対象グループ
     * @param {string|Array} [alignment] - 配置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は ScriptUI 既定値）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        group.alignment = alignment || "left";
        if (typeof spacing === "number") group.spacing = spacing;
    }

    /**
     * 列グループの共通設定を適用する
     * @param {Group} group - 対象グループ
     * @param {string|Array} [alignChildren] - 子要素の配置（省略時は ["fill", "top"]）
     * @param {string|Array} [alignment] - グループ自身の配置（省略時は ["fill", "top"]）
     * @returns {void}
     */
    function setupColumn(group, alignChildren, alignment) {
        group.orientation = "column";
        group.alignChildren = alignChildren || ["fill", "top"];
        group.alignment = alignment || ["fill", "top"];
    }

    /**
     * 2カラムを横に並べる行グループの共通設定を適用する
     * @param {Group} group - 対象グループ
     * @returns {void}
     */
    function setupColumnsRow(group) {
        group.orientation = "row";
        group.alignChildren = ["fill", "top"];
        group.alignment = ["fill", "top"];
        group.spacing = COLUMN_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードの並び順を整理", en: "Organize Artboard Order" }
        },
        panel: {
            reorder:          { ja: "パネル上の並び順", en: "Artboards Panel Order" },
            rearrange:        { ja: "カンバス上のアートボードを再配置", en: "Rearrange Canvas Artboards" },
            duplicateHandling: { ja: "未指定／重複", en: "Unspecified / Duplicate" },
            naming:           { ja: "アートボード名の更新", en: "Update names as row-column" },
            namingSeparator:  { ja: "区切り文字", en: "Separator" },
            namingPadWidth:   { ja: "桁数", en: "Digits" }
        },
        radio: {
            sortByName:               { ja: "名前順", en: "By name" },
            sortByPosition:           { ja: "カンバス上の並び順に", en: "Match canvas order" },
            sortKeepAsIs:             { ja: "変更しない", en: "Keep as is" },
            duplicateAppendToRowEnd:  { ja: "各行の末尾に配置", en: "Append to each row end" },
            duplicateGroupInLastRow:  { ja: "最終行の次の行にまとめる", en: "Group in row after last row" },
            namingFromPosition:       { ja: "配置位置から作成", en: "Create from position" },
            namingFromExisting:       { ja: "既存名を整形", en: "Reformat existing names" }
        },
        checkbox: {
            rearrangeByColumns: { ja: "列数を指定して再配置", en: "Rearrange by column count" },
            rearrangeByName:    { ja: "アートボード名から行列に再配置", en: "Rearrange by row-column from names" },
            gapLink:            { ja: "連動", en: "Link" },
            namingEnable:       { ja: "「行-列」形式に更新", en: "Update" }
        },
        fieldLabel: {
            columns:   { ja: "列数：", en: "Columns:" },
            columnGap: { ja: "列間：", en: "Column gap:" },
            rowGap:    { ja: "行間：", en: "Row gap:" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noArtboards: { ja: "アートボードが存在しません。", en: "No artboards found." },
            errorPrefix: { ja: "エラー: ", en: "Error: " },
            noMatch: {
                ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）形式のアートボード名が見つかりませんでした。",
                en: "No artboard names in '<row><sep><column>' (e.g., 1-2) or '<prefix><sep><number>' (e.g., banner-1) format were found."
            }
        }
    };

    /**
     * ラベルを現在のUI言語で取得する
     * @param {...string} labelPath - たどるキー（例: getLabel("dialog", "title")）
     * @returns {string} ローカライズ済みラベル
     */
    function getLabel() {
        var node = LABELS;
        for (var pathIndex = 0; pathIndex < arguments.length; pathIndex++) {
            if (node == null) break;
            node = node[arguments[pathIndex]];
        }
        return (node && node[uiLang] != null) ? node[uiLang] : "";
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードとラベルのマッピング / Mapping of unit codes to labels */
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

    /**
     * 現在の定規単位のラベルを取得する
     * @returns {string} 単位ラベル（取得できない場合は "pt"）
     */
    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return unitLabelMap[unitCode] || "pt";
    }
    var currentUnitLabel = getCurrentUnitLabel();

    /**
     * 表示単位の数値をポイントに変換する
     * @param {number} value - 表示単位の値
     * @returns {number} ポイント値
     */
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
            case 10: return value * 72 * 12;            /* ft */
            default: return value;
        }
    }

    /**
     * 入力欄の表示単位値を読み取り、ポイントに変換する
     * @param {EditText} input - 対象の入力欄
     * @param {number} defaultValue - 数値として読めない場合の既定値（表示単位）
     * @returns {number} ポイント値
     */
    function readDisplayUnitInputAsPoints(input, defaultValue) {
        var value = parseFloat(input.text);
        if (isNaN(value)) value = defaultValue;
        return convertDisplayUnitToPoints(value);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * @typedef {Object} ArtboardEntry
     * @property {string} name - アートボード名
     * @property {number[]} artboardRect - [左, 上, 右, 下]
     * @property {number} [sourceIndex] - 元のアートボードインデックス
     * @property {number} [rowBand] - 上辺の近さで割り当てた 0 始まりの行番号
     */

    /**
     * @typedef {Object} PreviewContext
     * @property {ArtboardEntry[]} artboardEntries - プレビュー用のアートボード情報
     * @property {number} sliderMax - 許容差スライダーの最大値
     * @property {number} defaultTolerance - 許容差スライダーの初期値
     */

    /**
     * ドキュメントを検証してダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (doc.artboards.length === 0) {
            alert(getLabel("alert", "noArtboards"));
            return;
        }

        var previewContext = buildPreviewContext(doc);
        var ui = buildDialogUI(previewContext.defaultTolerance, previewContext.sliderMax);

        bindEvents(doc, previewContext, ui);

        ui.dialog.center();
        ui.dialog.show();
    }

    /**
     * プレビューと許容差の初期値を組み立てる
     * @param {Document} doc - 対象ドキュメント
     * @returns {PreviewContext} プレビュー用のコンテキスト
     */
    function buildPreviewContext(doc) {
        /* アートボード情報を配列に格納（プレビュー・自動計算用） / Snapshot artboards for preview and auto tolerance */
        var artboardEntries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            artboardEntries.push({
                name: doc.artboards[artboardIndex].name,
                artboardRect: doc.artboards[artboardIndex].artboardRect
            });
        }

        /* スライダー最大値をアートボードの最大高さに設定 / Use the tallest artboard as the slider maximum */
        var maxHeight = 0;
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            var artboardRect = artboardEntries[entryIndex].artboardRect;
            var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);
            if (artboardHeight > maxHeight) maxHeight = artboardHeight;
        }
        var sliderMax = Math.round(maxHeight);

        /* スライダー初期値を自動計算値にマージンを掛けて設定 / Seed the slider with the auto value plus a margin */
        var autoTolerance = calculateAutoTolerance(artboardEntries);
        var defaultTolerance = Math.round(autoTolerance * TOLERANCE_AUTO_MARGIN_RATIO);
        defaultTolerance = Math.min(defaultTolerance, sliderMax);

        return {
            artboardEntries: artboardEntries,
            sliderMax: sliderMax,
            defaultTolerance: defaultTolerance
        };
    }

    /**
     * 許容差を自動計算する（上辺の最小差 + マージン）
     * @param {ArtboardEntry[]} artboardEntries - アートボード情報
     * @returns {number} 許容差（pt）
     */
    function calculateAutoTolerance(artboardEntries) {
        var tops = [];
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            tops.push(artboardEntries[entryIndex].artboardRect[1]);
        }
        /* 上から下に並べる / Sort top to bottom */
        tops.sort(function (firstTop, secondTop) {
            return secondTop - firstTop;
        });

        var diffs = [];
        for (var topIndex = 1; topIndex < tops.length; topIndex++) {
            var diff = Math.abs(tops[topIndex] - tops[topIndex - 1]);
            if (diff > 0) diffs.push(diff);
        }

        /* 差が無ければデフォルト / Default if no difference */
        if (diffs.length === 0) return TOLERANCE_FALLBACK_POINTS;

        /* 少しマージンを加える / Add some margin */
        return Math.min.apply(null, diffs) + TOLERANCE_AUTO_MARGIN_POINTS;
    }

    /**
     * ダイアログのイベントを設定する
     * @param {Document} doc - 対象ドキュメント
     * @param {PreviewContext} previewContext - プレビュー用のコンテキスト
     * @param {Object} ui - buildDialogUI() が返すUI参照
     * @returns {void}
     */
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

        /* 許容差はパネル並び順と「配置位置から作成」の両方で使うため、どちらかが有効なら操作できるようにする
         * The tolerance drives both the panel order and position-based naming, so enable it for either */
        function syncToleranceEnabled() {
            var usedByPanelOrder = sortRadios.byPosition.value;
            var usedByNaming = ui.naming.enableCheckbox.value && ui.naming.source.fromPosition.value;
            ui.preview.toleranceSlider.enabled = usedByPanelOrder || usedByNaming;
        }

        function syncSortMode() {
            syncToleranceEnabled();
            updateReorderPreview(previewContext, ui, Math.round(ui.preview.toleranceSlider.value));
        }

        function syncNamingState() {
            ui.naming.settingsGroup.enabled = ui.naming.enableCheckbox.value;
            syncToleranceEnabled();
        }

        syncNamingState();
        syncSortMode();

        ui.preview.toleranceSlider.onChanging = function () {
            updateReorderPreview(previewContext, ui, Math.round(ui.preview.toleranceSlider.value));
        };

        sortRadios.byPosition.onClick = function () { selectSortMode(sortRadios.byPosition); };
        sortRadios.byName.onClick = function () { selectSortMode(sortRadios.byName); };
        sortRadios.keepAsIs.onClick = function () { selectSortMode(sortRadios.keepAsIs); };

        ui.naming.enableCheckbox.onClick = syncNamingState;
        ui.naming.source.fromPosition.onClick = syncToleranceEnabled;
        ui.naming.source.fromExisting.onClick = syncToleranceEnabled;

        ui.buttons.okBtn.onClick = function () {
            executeReorder(doc, ui);
        };

        ui.buttons.cancelBtn.onClick = function () {
            ui.dialog.close(-1);
        };
    }

    /**
     * 並び順リストのプレビューを更新する
     * @param {PreviewContext} previewContext - プレビュー用のコンテキスト
     * @param {Object} ui - buildDialogUI() が返すUI参照
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
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

        /* 位置順モード: 行ごとにまとめて表示 / Position mode: show one row per line */
        var decimalPlaces = Math.pow(10, COORDINATE_PRECISION_DIGITS);
        var sortedEntries = previewContext.artboardEntries.slice();
        sortArtboardsTopLeftWithTolerance(sortedEntries, decimalPlaces, tolerance);

        var rowGroups = groupSortedIntoRows(sortedEntries);
        for (var rowIndex = 0; rowIndex < rowGroups.length; rowIndex++) {
            var rowNames = [];
            for (var columnIndex = 0; columnIndex < rowGroups[rowIndex].length; columnIndex++) {
                rowNames.push(rowGroups[rowIndex][columnIndex].name);
            }
            reorderList.add("item", rowNames.join(" | "));
        }
    }

    /**
     * ダイアログの設定を読み取り、再配置・並べ替え・リネームを実行する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - buildDialogUI() が返すUI参照
     * @returns {void}
     */
    function executeReorder(doc, ui) {
        var tolerance = Math.round(ui.preview.toleranceSlider.value) || 0;

        /* 再配置に失敗したらアラート済みなので、ダイアログを開いたまま中断する
         * Abort with the dialog still open when the rearrange failed (already alerted) */
        var rearrangeResult = applyCanvasRearrange(doc, ui);
        if (!rearrangeResult.ok) return;

        applyPanelReorder(doc, ui, tolerance);
        applyArtboardRenaming(doc, ui, tolerance);

        ui.dialog.close(1);

        /* ダイアログを閉じてからビューを合わせる（モーダル表示中のメニュー実行を避ける）
         * Fit the view after closing, so no menu command runs while the modal dialog is up */
        if (rearrangeResult.rearranged) app.executeMenuCommand('fitall');
    }

    /**
     * カンバス上の再配置を実行する
     * 再配置を先に実行すると、パネル順が「再配置後の見た目」と揃う
     * Run rearrange first so the panel order reflects post-rearrange positions
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - buildDialogUI() が返すUI参照
     * @returns {{ok: boolean, rearranged: boolean}} ok は後続処理を続けてよいか、rearranged は実際に動かしたか
     */
    function applyCanvasRearrange(doc, ui) {
        var modeChecks = ui.rearrange.modeChecks;
        if (!modeChecks.byColumns.value && !modeChecks.byName.value) {
            return { ok: true, rearranged: false };
        }

        /* 入力値は表示単位として受け取り、ptに変換して渡す / Read inputs in display units and convert to points */
        var columnGapPoints = readDisplayUnitInputAsPoints(ui.rearrange.columnGapInput, DEFAULT_GAP_VALUE);
        var rowGapPoints = ui.rearrange.gapLinkCheckbox.value
            ? columnGapPoints
            : readDisplayUnitInputAsPoints(ui.rearrange.rowGapInput, DEFAULT_GAP_VALUE);

        try {
            if (modeChecks.byColumns.value) {
                rearrangeArtboardsWithGaps(doc, readColumnCount(ui.rearrange.columnsInput), columnGapPoints, rowGapPoints);
                return { ok: true, rearranged: true };
            }
            var exceptionMode = ui.rearrange.duplicateRadios.groupLast.value ? 'lastRow' : 'rowEnd';
            var matched = rearrangeArtboardsByRowColumnName(doc, columnGapPoints, rowGapPoints, exceptionMode);
            return { ok: matched, rearranged: matched };
        } catch (rearrangeError) {
            alert(getLabel("alert", "errorPrefix") + rearrangeError.message);
            return { ok: false, rearranged: false };
        }
    }

    /**
     * 列数入力を読み取り、下限でクランプする
     * @param {EditText} columnsInput - 列数の入力欄
     * @returns {number} 1 以上の列数
     */
    function readColumnCount(columnsInput) {
        var columns = parseInt(columnsInput.text, 10);
        if (isNaN(columns)) columns = DEFAULT_COLUMN_COUNT;
        if (columns < MIN_COLUMN_COUNT) columns = MIN_COLUMN_COUNT;
        columnsInput.text = columns;
        return columns;
    }

    /**
     * ［アートボード］パネルの並び順を更新する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - buildDialogUI() が返すUI参照
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
    function applyPanelReorder(doc, ui, tolerance) {
        /* 並び順モードに応じてソーター切替（変更しない場合はスキップ）
         * Pick a sorter based on the selected sort mode; skip when "Keep as is" */
        if (ui.preview.sortModeRadios.keepAsIs.value) return;

        var sorter;
        if (ui.preview.sortModeRadios.byName.value) {
            sorter = function (entries) {
                sortArtboardsByName(entries);
            };
        } else {
            sorter = function (entries, decimalPlaces) {
                sortArtboardsTopLeftWithTolerance(entries, decimalPlaces, tolerance);
            };
        }
        rebuildArtboardsInSortedOrder(doc, sorter, COORDINATE_PRECISION_DIGITS);
    }

    /**
     * アートボード名を「行-列」形式に更新する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - buildDialogUI() が返すUI参照
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
    function applyArtboardRenaming(doc, ui, tolerance) {
        if (!ui.naming.enableCheckbox.value) return;

        var separators = ui.naming.separators;
        var separator = separators.underscore.value ? "_" : (separators.x.value ? "x" : "-");
        var padRadios = ui.naming.padRadios;
        var padWidth = padRadios.w3.value ? 3 : (padRadios.w2.value ? 2 : 1);

        /* 配置位置から作成 / 既存名を整形 / Create from position, or reformat existing names */
        if (ui.naming.source.fromPosition.value) {
            renameArtboardsFromPositions(doc, separator, padWidth, tolerance);
        } else {
            renameArtboardsFromExistingNames(doc, separator, padWidth);
        }
    }

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * ダイアログ全体を組み立てる
     * @param {number} defaultTolerance - 許容差スライダーの初期値
     * @param {number} sliderMax - 許容差スライダーの最大値
     * @returns {Object} ダイアログとUI参照をまとめたオブジェクト
     */
    function buildDialogUI(defaultTolerance, sliderMax) {
        var dialog = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 並び順は上段にフル幅 / Order panel spans full width on top */
        var preview = buildPreviewPanel(dialog, defaultTolerance, sliderMax);

        /* 下段は 2 カラム（左: 再配置 / 右: 命名） / Two-column row: rearrange (left) / naming (right) */
        var twoColumnsRow = dialog.add("group");
        setupColumnsRow(twoColumnsRow);

        var leftColumn = twoColumnsRow.add("group");
        setupColumn(leftColumn);

        var rightColumn = twoColumnsRow.add("group");
        setupColumn(rightColumn);

        var rearrange = buildRearrangePanel(leftColumn);
        var naming = buildNamingPanel(rightColumn);
        var buttons = buildDialogButtons(dialog);

        return {
            dialog: dialog,
            preview: preview,
            rearrange: rearrange,
            naming: naming,
            buttons: buttons
        };
    }

    /**
     * プレビューパネル（並び順モード＋許容差スライダー＋並び順リスト）を作成する
     * @param {Group|Window} parent - 追加先のコンテナ
     * @param {number} defaultTolerance - 許容差スライダーの初期値
     * @param {number} sliderMax - 許容差スライダーの最大値
     * @returns {Object} 並び順ラジオ・スライダー・リストの参照
     */
    function buildPreviewPanel(parent, defaultTolerance, sliderMax) {
        var reorderPanel = parent.add("panel", undefined, getLabel("panel", "reorder"));
        setupPanel(reorderPanel, DENSE_SPACING);

        /* 名前順（ラジオ） / Radio "By name" */
        var byNameRow = reorderPanel.add("group");
        setupRow(byNameRow);
        var sortByNameRadio = byNameRow.add("radiobutton", undefined, getLabel("radio", "sortByName"));

        /* カンバス上の並び順に（ラジオ）＋ 許容差スライダーを同じ行に配置
         * Radio "Match canvas order" and tolerance slider on the same row */
        var byPositionRow = reorderPanel.add("group");
        setupRow(byPositionRow);

        var sortByPositionRadio = byPositionRow.add("radiobutton", undefined, getLabel("radio", "sortByPosition"));
        sortByPositionRadio.value = true;

        var toleranceSlider = byPositionRow.add("slider", undefined, defaultTolerance, 0, sliderMax);
        toleranceSlider.preferredSize = [SLIDER_WIDTH, SLIDER_HEIGHT];

        /* 変更しない（ラジオ） / Radio "Keep as is" */
        var keepAsIsRow = reorderPanel.add("group");
        setupRow(keepAsIsRow);
        var sortKeepAsIsRadio = keepAsIsRow.add("radiobutton", undefined, getLabel("radio", "sortKeepAsIs"));

        var reorderList = reorderPanel.add("listbox", undefined, [], { multiselect: false });
        reorderList.preferredSize.width = PREVIEW_LIST_WIDTH;
        reorderList.minimumSize.height = PREVIEW_LIST_HEIGHT;
        reorderList.alignment = ["fill", "fill"];

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

    /**
     * 再配置パネルを作成する
     * @param {Group} parent - 追加先のコンテナ
     * @returns {Object} 再配置設定コントロールの参照
     */
    function buildRearrangePanel(parent) {
        var rearrangePanel = parent.add("panel", undefined, getLabel("panel", "rearrange"));
        setupPanel(rearrangePanel);

        var modeChecks = buildRearrangeModeCheckboxes(rearrangePanel);

        var rearrangeSettingsGroup = rearrangePanel.add("group");
        setupColumn(rearrangeSettingsGroup);

        var columnsRow = addLabeledInput(rearrangeSettingsGroup, getLabel("fieldLabel", "columns"), String(DEFAULT_COLUMN_COUNT), MIN_COLUMN_COUNT);
        var spacingControls = buildRearrangeSpacingControls(rearrangeSettingsGroup);
        var duplicateRadios = buildDuplicateHandlingPanel(rearrangePanel);

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

    /**
     * 再配置モードのチェックボックスを作成する（列数指定 / アートボード名、排他的に選択）
     * @param {Panel} parent - 追加先のパネル
     * @returns {Object} 各モードのチェックボックス参照
     */
    function buildRearrangeModeCheckboxes(parent) {
        var modeGroup = parent.add("group");
        setupColumn(modeGroup, ["left", "top"]);

        var byColumnsCheckbox = modeGroup.add("checkbox", undefined, getLabel("checkbox", "rearrangeByColumns"));
        var byNameCheckbox = modeGroup.add("checkbox", undefined, getLabel("checkbox", "rearrangeByName"));
        byColumnsCheckbox.value = false;
        byNameCheckbox.value = false;

        return {
            byColumns: byColumnsCheckbox,
            byName: byNameCheckbox
        };
    }

    /**
     * 再配置の列間／行間入力と連動チェックを作成する
     * @param {Group} parent - 追加先のコンテナ
     * @returns {Object} 列間・行間入力欄と連動チェックの参照
     */
    function buildRearrangeSpacingControls(parent) {
        /* 列間／行間 + 連動チェックの2カラム / Two-column row: gap inputs (left) + link checkbox (right) */
        var gapsRow = parent.add("group");
        setupRow(gapsRow, "left", COLUMN_SPACING);

        var gapsLeft = gapsRow.add("group");
        setupColumn(gapsLeft, ["left", "top"], ["left", "top"]);

        var columnGapRow = addLabeledInput(gapsLeft, getLabel("fieldLabel", "columnGap"), String(DEFAULT_GAP_VALUE), MIN_GAP_VALUE);
        var columnGapInput = columnGapRow.input;
        columnGapRow.row.add("statictext", undefined, currentUnitLabel);

        var rowGapRow = addLabeledInput(gapsLeft, getLabel("fieldLabel", "rowGap"), String(DEFAULT_GAP_VALUE), MIN_GAP_VALUE);
        var rowGapInput = rowGapRow.input;
        rowGapRow.row.add("statictext", undefined, currentUnitLabel);

        var gapLinkCheckbox = gapsRow.add("checkbox", undefined, getLabel("checkbox", "gapLink"));
        gapLinkCheckbox.value = true;

        bindGapLinkControls(columnGapInput, rowGapInput, rowGapRow.row, gapLinkCheckbox);

        return {
            columnGapInput: columnGapInput,
            rowGapInput: rowGapInput,
            gapLinkCheckbox: gapLinkCheckbox
        };
    }

    /**
     * 列間／行間の連動挙動を設定する
     * @param {EditText} columnGapInput - 列間の入力欄
     * @param {EditText} rowGapInput - 行間の入力欄
     * @param {Group} rowGapControlRow - 行間の入力行（連動時にディムする）
     * @param {Checkbox} gapLinkCheckbox - 連動チェックボックス
     * @returns {void}
     */
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

    /**
     * 未指定／重複の扱いパネルを作成する
     * @param {Panel} parent - 追加先のパネル
     * @returns {Object} パネルと各ラジオの参照
     */
    function buildDuplicateHandlingPanel(parent) {
        var duplicatePanel = parent.add("panel", undefined, getLabel("panel", "duplicateHandling"));
        setupPanel(duplicatePanel, DENSE_SPACING);

        var duplicateAppendRadio = duplicatePanel.add("radiobutton", undefined, getLabel("radio", "duplicateAppendToRowEnd"));
        var duplicateGroupRadio = duplicatePanel.add("radiobutton", undefined, getLabel("radio", "duplicateGroupInLastRow"));
        duplicateAppendRadio.value = true;

        return {
            panel: duplicatePanel,
            append: duplicateAppendRadio,
            groupLast: duplicateGroupRadio
        };
    }

    /**
     * 再配置パネルの有効／無効状態を同期する
     * @param {Object} modeChecks - 再配置モードのチェックボックス参照
     * @param {Group} rearrangeSettingsGroup - 再配置設定のコンテナ
     * @param {Object} columnsRow - 列数入力行（addLabeledInput の戻り値）
     * @param {Panel} duplicatePanel - 未指定／重複パネル
     * @returns {void}
     */
    function syncRearrangePanelState(modeChecks, rearrangeSettingsGroup, columnsRow, duplicatePanel) {
        var hasRearrangeMode = modeChecks.byColumns.value || modeChecks.byName.value;
        rearrangeSettingsGroup.enabled = hasRearrangeMode;
        /* 列数は「列数を指定して再配置」のときだけアクティブ
         * (byName では名前から行列を取るので列数は不要) / Columns input is enabled only for byColumns */
        columnsRow.row.enabled = modeChecks.byColumns.value;
        /* 未指定／重複は byName のときだけアクティブ / Duplicate panel is active only for byName */
        duplicatePanel.enabled = modeChecks.byName.value;
    }

    /**
     * 命名パネルを作成する
     * @param {Group} parent - 追加先のコンテナ
     * @returns {Object} 命名設定コントロールの参照
     */
    function buildNamingPanel(parent) {
        var namingPanel = parent.add("panel", undefined, getLabel("panel", "naming"));
        setupPanel(namingPanel);

        var namingEnableCheckbox = namingPanel.add("checkbox", undefined, getLabel("checkbox", "namingEnable"));
        namingEnableCheckbox.value = false;

        var namingSettingsGroup = namingPanel.add("group");
        setupColumn(namingSettingsGroup);

        /* 命名ソース / Naming source */
        var sourceGroup = namingSettingsGroup.add("group");
        setupColumn(sourceGroup, ["left", "top"]);
        var fromPositionRadio = sourceGroup.add("radiobutton", undefined, getLabel("radio", "namingFromPosition"));
        var fromExistingRadio = sourceGroup.add("radiobutton", undefined, getLabel("radio", "namingFromExisting"));
        fromPositionRadio.value = true;

        /* 区切り文字パネルと桁数パネルを横並び / Separator and digits panels side by side */
        var separatorPadRow = namingSettingsGroup.add("group");
        setupColumnsRow(separatorPadRow);

        var separatorRadios = buildOptionRadioPanel(separatorPadRow, getLabel("panel", "namingSeparator"), ["-", "_", "x"]);
        var padRadios = buildOptionRadioPanel(separatorPadRow, getLabel("panel", "namingPadWidth"), ["0", "00", "000"]);

        /* 有効／無効の同期は bindEvents() が担当（許容差スライダーの状態と連動するため）
         * bindEvents() owns the enable sync, since it also drives the tolerance slider */
        return {
            enableCheckbox: namingEnableCheckbox,
            settingsGroup: namingSettingsGroup,
            source: {
                fromPosition: fromPositionRadio,
                fromExisting: fromExistingRadio
            },
            separators: {
                hyphen: separatorRadios[0],
                underscore: separatorRadios[1],
                x: separatorRadios[2]
            },
            padRadios: {
                w1: padRadios[0],
                w2: padRadios[1],
                w3: padRadios[2]
            }
        };
    }

    /**
     * 選択肢ラジオを縦に並べたパネルを作成する（先頭を選択状態にする）
     * @param {Group} parent - 追加先のコンテナ
     * @param {string} panelLabel - パネルのラベル
     * @param {string[]} optionLabels - ラジオのラベル
     * @returns {RadioButton[]} 作成したラジオボタン
     */
    function buildOptionRadioPanel(parent, panelLabel, optionLabels) {
        var optionPanel = parent.add("panel", undefined, panelLabel);
        setupPanel(optionPanel, DENSE_SPACING);

        var optionGroup = optionPanel.add("group");
        setupColumn(optionGroup, ["left", "top"]);

        var radios = [];
        for (var optionIndex = 0; optionIndex < optionLabels.length; optionIndex++) {
            radios.push(optionGroup.add("radiobutton", undefined, optionLabels[optionIndex]));
        }
        if (radios.length > 0) radios[0].value = true;
        return radios;
    }

    /**
     * キャンセル / OK ボタンを作成する
     * @param {Window} dialog - 対象ダイアログ
     * @returns {Object} 各ボタンの参照
     */
    function buildDialogButtons(dialog) {
        var buttonGroup = dialog.add("group");
        setupRow(buttonGroup, ["right", "bottom"]);

        var cancelBtn = buttonGroup.add("button", undefined, getLabel("button", "cancel"));
        var okBtn = buttonGroup.add("button", undefined, getLabel("button", "ok"));

        var buttonWidth = Math.max(okBtn.preferredSize.width, cancelBtn.preferredSize.width);
        okBtn.preferredSize.width = buttonWidth;
        cancelBtn.preferredSize.width = buttonWidth;

        return {
            cancelBtn: cancelBtn,
            okBtn: okBtn
        };
    }

    /**
     * ラベル付き edittext 行を追加する
     * @param {Group} parent - 追加先のコンテナ
     * @param {string} labelText - ラベル文字列
     * @param {string} defaultValue - 入力欄の初期値
     * @param {number} minValue - ↑↓キーで下回らせない下限値
     * @returns {{row: Group, input: EditText}} 行グループと入力欄
     */
    function addLabeledInput(parent, labelText, defaultValue, minValue) {
        var row = parent.add("group");
        setupRow(row);
        var label = row.add("statictext", undefined, labelText);
        label.preferredSize.width = FIELD_LABEL_WIDTH;
        label.justify = "right";
        var input = row.add("edittext", undefined, defaultValue);
        input.characters = FIELD_INPUT_CHARS;
        changeValueByArrowKey(input, minValue);
        return { row: row, input: input };
    }

    /**
     * 入力欄で ↑↓キーによる数値の増減を有効にする（Shift で 10 の倍数にスナップ）
     * @param {EditText} editText - 対象の入力欄
     * @param {number} minValue - 下回らせない下限値
     * @returns {void}
     */
    function changeValueByArrowKey(editText, minValue) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var isShift = ScriptUI.environment.keyboardState.shiftKey;
            var direction = (event.keyName === "Up") ? 1 : -1;

            if (isShift) {
                /* Shift+矢印キーで 10 の倍数にスナップ / Snap to multiples of 10 */
                value = (direction > 0)
                    ? Math.ceil((value + 1) / 10) * 10
                    : Math.floor((value - 1) / 10) * 10;
            } else {
                value += direction;
            }
            if (value < minValue) value = minValue;

            editText.text = Math.round(value);
            editText.notify("onChange");
            event.preventDefault();
        });
    }

    // =========================================
    // ソート処理 / Sorting
    // =========================================

    /**
     * アートボードのプロパティをコピーする
     * @param {Artboard|Object} sourceArtboard - コピー元
     * @param {Artboard|Object} targetArtboard - コピー先
     * @returns {void}
     */
    function copyArtboardProps(sourceArtboard, targetArtboard) {
        for (var propertyIndex = 0; propertyIndex < ARTBOARD_COPYABLE_PROPS.length; propertyIndex++) {
            var propertyName = ARTBOARD_COPYABLE_PROPS[propertyIndex];
            targetArtboard[propertyName] = sourceArtboard[propertyName];
        }
    }

    /**
     * アートボードを並べ替えて［アートボード］パネル順を再構築する
     * @param {Document} doc - 対象ドキュメント
     * @param {function(ArtboardEntry[], number): void} sorterFunction - 並べ替え関数
     * @param {number} precisionDigits - 座標比較の丸め桁数
     * @returns {void}
     */
    function rebuildArtboardsInSortedOrder(doc, sorterFunction, precisionDigits) {
        var decimalPlaces = Math.pow(10, precisionDigits || COORDINATE_PRECISION_DIGITS);

        /* 元アートボードのスナップショット（プロパティ＋rect）を取得 / Snapshot original artboards */
        var artboardSnapshots = [];
        var sortableEntries = [];
        var liveOriginalIndexes = [];
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
            liveOriginalIndexes.push(artboardIndex);
        }

        sorterFunction(sortableEntries, decimalPlaces);

        /* 1件複製するたびに元を削除する。まとめて複製すると一時的にアートボード数が倍になり、
         * Illustrator のアートボード上限に当たるため。未削除の元は常に先頭側に残る。
         * Copy-then-delete one at a time: duplicating them all at once would temporarily
         * double the artboard count and hit Illustrator's limit. Pending originals stay at the front. */
        for (var sortedIndex = 0; sortedIndex < sortableEntries.length; sortedIndex++) {
            var sourceIndex = sortableEntries[sortedIndex].sourceIndex;
            var sourceSnapshot = artboardSnapshots[sourceIndex];
            var newArtboard = doc.artboards.add(sourceSnapshot.artboardRect);
            copyArtboardProps(sourceSnapshot, newArtboard);

            var livePosition = indexOfValue(liveOriginalIndexes, sourceIndex);
            doc.artboards[livePosition].remove();
            liveOriginalIndexes.splice(livePosition, 1);
        }
    }

    /**
     * 配列から値の位置を探す（ES3 に Array#indexOf がないため）
     * @param {Array} list - 検索対象の配列
     * @param {*} value - 探す値
     * @returns {number} 見つかった位置。見つからなければ -1
     */
    function indexOfValue(list, value) {
        for (var searchIndex = 0; searchIndex < list.length; searchIndex++) {
            if (list[searchIndex] === value) return searchIndex;
        }
        return -1;
    }

    /**
     * 座標を指定の係数で丸める
     * @param {number} value - 座標値
     * @param {number} decimalPlaces - 丸め係数（10 の冪）
     * @returns {number} 丸めた座標値
     */
    function roundCoordinate(value, decimalPlaces) {
        return Math.round(value * decimalPlaces) / decimalPlaces;
    }

    /**
     * 上辺の近さで行バンド（0 始まりの行番号）を割り当てる
     * 行の基準は「その行でいちばん上のアートボードの上辺」なので、
     * 許容差ずつ階段状にずれた並びが1行に連鎖して吸収されることはない
     * Each row is anchored to its own topmost edge, so a staircase of
     * within-tolerance steps no longer collapses into a single row.
     * @param {ArtboardEntry[]} artboardEntries - 対象の配列（上辺降順に並べ替え、rowBand を書き込む）
     * @param {number} decimalPlaces - 座標の丸め係数
     * @param {number} tolerance - 同一行とみなす上辺の差（pt）
     * @returns {void}
     */
    function assignRowBands(artboardEntries, decimalPlaces, tolerance) {
        artboardEntries.sort(function (firstEntry, secondEntry) {
            return roundCoordinate(secondEntry.artboardRect[1], decimalPlaces) -
                roundCoordinate(firstEntry.artboardRect[1], decimalPlaces);
        });

        var bandIndex = 0;
        var bandTop = null;
        for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
            var currentTop = roundCoordinate(artboardEntries[entryIndex].artboardRect[1], decimalPlaces);
            if (bandTop === null) {
                bandTop = currentTop;
            } else if (bandTop - currentTop > tolerance) {
                bandIndex++;
                bandTop = currentTop;
            }
            artboardEntries[entryIndex].rowBand = bandIndex;
        }
    }

    /**
     * 許容差付きの左上基準ソート（行バンド昇順 → 左辺昇順）
     * 行バンドを整数化してから比較するため、比較関数は推移律を満たす
     * @param {ArtboardEntry[]} artboardEntries - 並べ替える配列（破壊的に並べ替える）
     * @param {number} decimalPlaces - 座標の丸め係数
     * @param {number} tolerance - 同一行とみなす上辺の差（pt）
     * @returns {void}
     */
    function sortArtboardsTopLeftWithTolerance(artboardEntries, decimalPlaces, tolerance) {
        decimalPlaces = decimalPlaces || Math.pow(10, COORDINATE_PRECISION_DIGITS);
        assignRowBands(artboardEntries, decimalPlaces, tolerance);
        artboardEntries.sort(function (firstEntry, secondEntry) {
            if (firstEntry.rowBand !== secondEntry.rowBand) {
                return firstEntry.rowBand - secondEntry.rowBand;
            }
            return roundCoordinate(firstEntry.artboardRect[0], decimalPlaces) -
                roundCoordinate(secondEntry.artboardRect[0], decimalPlaces);
        });
    }

    /**
     * 名前順（自然順）にソートする。数字列を10桁ゼロ埋めして比較し、大文字小文字は無視する
     * @param {ArtboardEntry[]} artboardEntries - 並べ替える配列（破壊的に並べ替える）
     * @returns {void}
     */
    function sortArtboardsByName(artboardEntries) {
        artboardEntries.sort(function (firstEntry, secondEntry) {
            var firstName = padNumbersForNaturalSort((firstEntry.name || "").toLowerCase());
            var secondName = padNumbersForNaturalSort((secondEntry.name || "").toLowerCase());
            if (firstName < secondName) return -1;
            if (firstName > secondName) return 1;
            return 0;
        });
    }

    /**
     * 文字列中の数字列を10桁ゼロ埋めする
     * @param {string} text - 対象文字列
     * @returns {string} ゼロ埋め後の文字列
     */
    function padNumbersForNaturalSort(text) {
        return text.replace(/\d+/g, function (digitRun) {
            var pad = "0000000000";
            return (pad + digitRun).slice(-pad.length);
        });
    }

    /**
     * 行バンドごとにエントリをまとめる
     * 事前に sortArtboardsTopLeftWithTolerance() を通し、rowBand が入っている必要がある
     * @param {ArtboardEntry[]} sortedEntries - rowBand 付きでソート済みのアートボード情報
     * @returns {Array<ArtboardEntry[]>} 行ごとにまとめた配列
     */
    function groupSortedIntoRows(sortedEntries) {
        var rows = [];
        for (var entryIndex = 0; entryIndex < sortedEntries.length; entryIndex++) {
            if (entryIndex === 0 || sortedEntries[entryIndex].rowBand !== sortedEntries[entryIndex - 1].rowBand) {
                rows.push([]);
            }
            rows[rows.length - 1].push(sortedEntries[entryIndex]);
        }
        return rows;
    }

    // =========================================
    // 再配置処理（列数指定） / Rearrange by column count
    // =========================================

    /**
     * 列間を spacing として GridByRow で初回レイアウトし、行ごとに追加オフセットを適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {number} columns - 列数
     * @param {number} columnGapPoints - 列間（pt）
     * @param {number} rowGapPoints - 行間（pt）
     * @returns {void}
     */
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
    // 再配置処理（アートボード名） / Rearrange by row-column names
    // =========================================

    /**
     * @typedef {Object} PlacementItem
     * @property {Artboard} artboard - 対象アートボード
     * @property {number} width - 幅（pt）
     * @property {number} height - 高さ（pt）
     * @property {boolean} matched - 名前が解析できたか
     * @property {string} matchType - "rowColumn" または "prefixNumber"
     * @property {string} prefixName - 接頭辞（小文字）
     * @property {number} rowNumber - 名前から得た行番号
     * @property {number} columnNumber - 名前から得た列番号
     * @property {number} assignedRow - 実際に割り当てた行番号
     * @property {number} assignedColumn - 実際に割り当てた列番号
     */

    /**
     * アートボード名（「行-列」または「接頭辞-番号」）に従ってグリッド状に再配置する
     * 区切り文字は -, _, x（大小文字無視）/ Separators: -, _, x (case-insensitive)
     * @param {Document} doc - 対象ドキュメント
     * @param {number} columnGapPoints - 列間（pt）
     * @param {number} rowGapPoints - 行間（pt）
     * @param {string} exceptionMode - "rowEnd"（各行の末尾に配置）または "lastRow"（最終行の次の行にまとめる）
     * @returns {boolean} 解析できる名前が1件以上あり、再配置した場合は true
     */
    function rearrangeArtboardsByRowColumnName(doc, columnGapPoints, rowGapPoints, exceptionMode) {
        var placementContext = parseArtboardNamePlacements(doc.artboards);
        if (!placementContext.hasMatchedArtboard) {
            alert(getLabel("alert", "noMatch"));
            return false;
        }

        assignArtboardPlacementSlots(placementContext, exceptionMode);
        var placementMetrics = collectPlacementMetrics(placementContext.placementItems);
        computePlacementOffsets(placementContext, placementMetrics, columnGapPoints, rowGapPoints);
        applyComputedArtboardPositions(doc, placementContext.placementItems);
        return true;
    }

    /**
     * アートボード名を解析して配置候補を作成する
     * @param {Artboards} artboards - 対象のアートボードコレクション
     * @returns {Object} 配置候補と集計値をまとめたコンテキスト
     */
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

    /**
     * 配置候補の初期オブジェクトを作成する
     * @param {Artboard} artboard - 対象アートボード
     * @param {number} artboardWidth - 幅（pt）
     * @param {number} artboardHeight - 高さ（pt）
     * @returns {PlacementItem} 配置候補
     */
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

    /**
     * 「行-列」形式の名前を配置候補に反映する
     * @param {PlacementItem} placementItem - 配置候補
     * @param {string} artboardName - アートボード名
     * @returns {boolean} 反映できたら true
     */
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

    /**
     * 「接頭辞-番号」形式の名前を配置候補に反映する
     * @param {PlacementItem} placementItem - 配置候補
     * @param {string} artboardName - アートボード名
     * @param {string[]} prefixOrder - 出現順の接頭辞リスト（破壊的に追加する）
     * @param {Object} prefixRowOffsetByName - 接頭辞ごとの行オフセット（破壊的に追加する）
     * @returns {boolean} 反映できたら true
     */
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

    /**
     * 各配置候補に行・列番号を割り当てる
     * @param {Object} placementContext - parseArtboardNamePlacements() の戻り値
     * @param {string} exceptionMode - "rowEnd" または "lastRow"
     * @returns {void}
     */
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

    /**
     * 一致した名前の配置スロットを割り当てる
     * @param {PlacementItem} assignTarget - 配置候補
     * @param {Object} placementContext - parseArtboardNamePlacements() の戻り値
     * @param {Object} occupiedSlots - 使用済みスロット（"行,列" をキーにする）
     * @param {Object} nextExceptionColumnByRow - 行ごとの次の例外列
     * @returns {number} 割り当てた行番号
     */
    function assignMatchedPlacementSlot(assignTarget, placementContext, occupiedSlots, nextExceptionColumnByRow) {
        var targetRowNumber = assignTarget.rowNumber;
        if (assignTarget.matchType === 'prefixNumber') {
            targetRowNumber = placementContext.maxRowNumber + placementContext.prefixRowOffsetByName[assignTarget.prefixName];
        }

        assignTarget.assignedRow = targetRowNumber;
        var requestedKey = targetRowNumber + ',' + assignTarget.columnNumber;
        if (!occupiedSlots[requestedKey]) {
            occupiedSlots[requestedKey] = true;
            assignTarget.assignedColumn = assignTarget.columnNumber;
        } else {
            /* 同じスロットが埋まっている場合は行末に逃がす / Fall back to the row end when the slot is taken */
            assignTarget.assignedColumn = reserveNextColumnAtRowEnd(
                targetRowNumber, occupiedSlots, nextExceptionColumnByRow, placementContext.maxColumnNumber
            );
        }
        return targetRowNumber;
    }

    /**
     * 最終行の次の行にまとめる例外スロットを割り当てる
     * @param {PlacementItem} assignTarget - 配置候補
     * @param {number} exceptionRowNumber - 例外行の行番号
     * @param {number} lastRowExceptionColumn - 次に試す列番号
     * @param {Object} occupiedSlots - 使用済みスロット
     * @returns {number} 更新後の次に試す列番号
     */
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

    /**
     * 各行の末尾に置く例外スロットを割り当てる
     * @param {PlacementItem} assignTarget - 配置候補
     * @param {number} currentRowNumber - 直前に割り当てた行番号
     * @param {number} maxColumnNumber - 名前から得た最大列番号
     * @param {Object} occupiedSlots - 使用済みスロット
     * @param {Object} nextExceptionColumnByRow - 行ごとの次の例外列
     * @returns {void}
     */
    function assignRowEndExceptionSlot(assignTarget, currentRowNumber, maxColumnNumber, occupiedSlots, nextExceptionColumnByRow) {
        assignTarget.assignedRow = currentRowNumber;
        assignTarget.assignedColumn = reserveNextColumnAtRowEnd(
            currentRowNumber, occupiedSlots, nextExceptionColumnByRow, maxColumnNumber
        );
    }

    /**
     * 指定行の末尾（maxColumnNumber + n）で空きスロットを予約する
     * @param {number} rowNumber - 対象の行番号
     * @param {Object} occupiedSlots - 使用済みスロット
     * @param {Object} nextExceptionColumnByRow - 行ごとの次の例外列
     * @param {number} maxColumnNumber - 名前から得た最大列番号
     * @returns {number} 予約した列番号
     */
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

    /**
     * 列ごとの最大幅・行ごとの最大高さを集計する
     * @param {PlacementItem[]} placementItems - 配置候補
     * @returns {Object} 列幅・行高さと、使用されている列番号・行番号
     */
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

    /**
     * 数値キーを昇順で取り出す
     * @param {Object} numericKeyedObject - 数値をキーにしたオブジェクト
     * @returns {number[]} 昇順の数値キー
     */
    function collectSortedNumericKeys(numericKeyedObject) {
        var sortedKeys = [];
        for (var numericKey in numericKeyedObject) {
            if (numericKeyedObject.hasOwnProperty(numericKey)) {
                sortedKeys.push(parseInt(numericKey, 10));
            }
        }
        sortedKeys.sort(function (firstKey, secondKey) { return firstKey - secondKey; });
        return sortedKeys;
    }

    /**
     * 累積オフセットと移動量を計算する
     * @param {Object} placementContext - parseArtboardNamePlacements() の戻り値
     * @param {Object} placementMetrics - collectPlacementMetrics() の戻り値
     * @param {number} columnGapPoints - 列間（pt）
     * @param {number} rowGapPoints - 行間（pt）
     * @returns {void}
     */
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

    /**
     * 累積オフセットを計算する（boundaryNumber と一致するキーの直前の間隔は EXCEPTION_BOUNDARY_GAP_MULTIPLIER 倍）
     * @param {number[]} sortedNumbers - 昇順の行番号または列番号
     * @param {Object} sizeByNumber - 行高さまたは列幅
     * @param {number} gap - 基本の間隔（pt）
     * @param {number} boundaryNumber - 例外領域が始まる番号
     * @returns {Object} 番号ごとの累積オフセット
     */
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

    /**
     * オブジェクト移動とアートボード矩形更新を適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacementItem[]} placementItems - 移動量を計算済みの配置候補
     * @returns {void}
     */
    function applyComputedArtboardPositions(doc, placementItems) {
        applyItemTranslations(doc, placementItems);

        for (var applyIndex = 0; applyIndex < placementItems.length; applyIndex++) {
            var applyItem = placementItems[applyIndex];
            var newRight = applyItem.newLeft + applyItem.width;
            var newBottom = applyItem.newTop - applyItem.height;
            applyItem.artboard.artboardRect = [applyItem.newLeft, applyItem.newTop, newRight, newBottom];
        }
    }

    // =========================================
    // アイテム移動 / Item translation
    // =========================================

    /**
     * placements に従ってアートボード上のアイテムを移動する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<{artboard: Artboard, oldRect: number[], deltaX: number, deltaY: number}>} placements - 移動量
     * @returns {void}
     */
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

    /**
     * レイヤーを再帰的に走査してアートボード上のオブジェクトの移動量を集める
     * @param {Layers} layerCollection - 走査するレイヤーコレクション
     * @param {Array<Object>} placementItems - oldRect と移動量を持つ配置情報
     * @returns {Array<{item: PageItem, deltaX: number, deltaY: number}>} 移動予定リスト
     */
    function collectItemTranslations(layerCollection, placementItems) {
        var collected = [];
        appendItemTranslations(layerCollection, placementItems, collected);
        return collected;
    }

    /**
     * レイヤーとサブレイヤーを走査して移動予定を追加する
     * @param {Layers} layerCollection - 走査するレイヤーコレクション
     * @param {Array<Object>} placementItems - oldRect と移動量を持つ配置情報
     * @param {Array<Object>} output - 追加先の配列
     * @returns {void}
     */
    function appendItemTranslations(layerCollection, placementItems, output) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
                var pageItem = layer.pageItems[itemIndex];
                /* サブレイヤーやグループの中身は親側で移動するのでスキップ / Sublayer and group children move with their parent */
                if (pageItem.parent !== layer) continue;
                /* ロックされたアイテムはスキップ / Skip locked items */
                try {
                    if (pageItem.locked) continue;
                } catch (lockedReadError) { /* プロパティが取得不能なら通常通り扱う / Treat as unlocked when unreadable */ }
                appendTranslationForItem(pageItem, placementItems, output);
            }
            if (layer.layers && layer.layers.length > 0) {
                appendItemTranslations(layer.layers, placementItems, output);
            }
        }
    }

    /**
     * オブジェクトの幾何中心が含まれる元アートボードを特定し、移動量を1件追加する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {Array<Object>} placementItems - oldRect と移動量を持つ配置情報
     * @param {Array<Object>} output - 追加先の配列
     * @returns {void}
     */
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

    /**
     * オブジェクトの幾何中心を取得する
     * @param {PageItem} pageItem - 対象オブジェクト
     * @returns {number[]|null} [x, y]、取得できない場合は null
     */
    function getItemGeometricCenter(pageItem) {
        try {
            var bounds = pageItem.geometricBounds; /* [左, 上, 右, 下] / [left, top, right, bottom] */
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (geometricBoundsError) {
            return null;
        }
    }

    /**
     * 中心座標が矩形内かを判定する（Illustrator の Y は上方向が正）
     * @param {number[]} center - [x, y]
     * @param {number[]} rect - [左, 上, 右, 下]
     * @returns {boolean} 矩形内なら true
     */
    function isCenterInsideRect(center, rect) {
        return center[0] >= rect[0] && center[0] <= rect[2] &&
            center[1] <= rect[1] && center[1] >= rect[3];
    }

    // =========================================
    // リネーム処理 / Renaming
    // =========================================

    /**
     * 数値をゼロ埋めする
     * @param {number} value - 対象の数値
     * @param {number} width - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function padNumber(value, width) {
        var padded = String(value);
        while (padded.length < width) padded = "0" + padded;
        return padded;
    }

    /**
     * 「行-列」形式の名前を組み立てる
     * @param {number} rowNumber - 行番号
     * @param {number} columnNumber - 列番号
     * @param {string} separator - 区切り文字
     * @param {number} padWidth - ゼロ埋め桁数
     * @returns {string} 組み立てた名前
     */
    function formatRowColumnName(rowNumber, columnNumber, separator, padWidth) {
        return padNumber(rowNumber, padWidth) + separator + padNumber(columnNumber, padWidth);
    }

    /**
     * 物理的な配置から「行-列」名を割り当てる
     * @param {Document} doc - 対象ドキュメント
     * @param {string} separator - 区切り文字
     * @param {number} padWidth - ゼロ埋め桁数
     * @param {number} tolerance - 行判定の許容差（pt）
     * @returns {void}
     */
    function renameArtboardsFromPositions(doc, separator, padWidth, tolerance) {
        var decimalPlaces = Math.pow(10, COORDINATE_PRECISION_DIGITS);

        var entries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            entries.push({
                sourceIndex: artboardIndex,
                artboardRect: doc.artboards[artboardIndex].artboardRect
            });
        }
        sortArtboardsTopLeftWithTolerance(entries, decimalPlaces, tolerance);

        var rows = groupSortedIntoRows(entries);
        for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            for (var columnIndex = 0; columnIndex < rows[rowIndex].length; columnIndex++) {
                var targetArtboard = doc.artboards[rows[rowIndex][columnIndex].sourceIndex];
                targetArtboard.name = formatRowColumnName(rowIndex + 1, columnIndex + 1, separator, padWidth);
            }
        }
    }

    /**
     * 既存の「行-列」名を再フォーマットする
     * @param {Document} doc - 対象ドキュメント
     * @param {string} separator - 区切り文字
     * @param {number} padWidth - ゼロ埋め桁数
     * @returns {void}
     */
    function renameArtboardsFromExistingNames(doc, separator, padWidth) {
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboard = doc.artboards[artboardIndex];
            var nameMatch = artboard.name.match(/^(\d+)[-_x](\d+)(.*)$/i);
            if (!nameMatch) continue;
            var rowNumber = parseInt(nameMatch[1], 10);
            var columnNumber = parseInt(nameMatch[2], 10);
            var rest = nameMatch[3] || "";
            artboard.name = formatRowColumnName(rowNumber, columnNumber, separator, padWidth) + rest;
        }
    }

    main();
})();
