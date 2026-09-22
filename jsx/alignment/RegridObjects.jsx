#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

だいたいグリッド状に並んでいる選択オブジェクトを、指定した左右・上下の間隔で再配置します。
間隔は現在の定規単位で入力でき、常時プレビューで結果を確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegridObjects.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n08861d0e40c3

### Overview

Re-lays out a roughly grid-shaped selection at the horizontal and vertical spacing you specify.
The spacing is entered in the current ruler units, with a continuous preview of the result.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegridObjects.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RegridObjects";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegridObjects.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegridObjects.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n08861d0e40c3"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 行列入れ替えで「同じ列」とみなす左端Xの許容差（pt）/ Tolerance for treating lefts as one column when transposing (pt) */
    var TRANSPOSE_SNAP_X_TOLERANCE = 8.0;

    /* 行列入れ替えで「同じ行」とみなす上端Yの許容差（pt）/ Tolerance for treating tops as one row when transposing (pt) */
    var TRANSPOSE_SNAP_Y_TOLERANCE = 8.0;

    /* ハニカムの行送りに掛ける係数（通常・レンガ状は 1.0）/ Row-step factor for the honeycomb layout (normal and brick use 1.0) */
    var HONEYCOMB_ROW_STEP_FACTOR = 0.75;

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* 強制グリッド（ダイアログから切替）/ Force Grid mode (toggled from dialog) */
    $.global.__regridForceGrid = $.global.__regridForceGrid || false;

    /* 中央揃え：各セルの天地左右中央に整列（強制グリッドのサブオプション）/ Center each object in its cell (sub-option of Force Grid) */
    $.global.__regridCenterInCell = $.global.__regridCenterInCell || false;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS     = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;                 /* 2カラムの間隔 / gap between columns */
    var SUB_OPTION_MARGINS = [15, 0, 0, 0];      /* サブオプションの字下げ / indent of sub-options */
    var GAP_INPUT_CHARS    = 4;                  /* 間隔の入力欄の文字数 / width of the gap fields, in characters */

    /**
     * ウィンドウに共通のレイアウト設定（縦並び・外周余白・要素間隔）を適用する
     * @param {Window} targetWindow - 対象のダイアログウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウト設定（縦並び・横いっぱい・余白・要素間隔）を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の要素間隔（省略時は PANEL_SPACING）
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
     * 行グループ（ボタン列など）に共通の横並び設定を適用する
     * @param {Group} rowGroup - 対象のグループ
     * @param {string} [alignment] - グループの配置（"left" / "center" / "right" など。省略時は "left"）
     * @param {number} [spacing] - グループ内の要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 入力欄で↑↓キーによる増減を行えるようにする
     * （↑↓：±1、Shift＋↑↓：±10（10の倍数にスナップ）、Option(Alt)＋↑↓：±0.1）。
     * 修飾キーは event を優先して読む（keyboardState は macOS で altKey を誤報するため）
     * @param {EditText} editText - 対象の入力欄
     * @param {Function} [onUpdate] - 値を変えたあとに呼ぶ関数（プレビュー更新用）
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            // 矢印キー（↑↓）以外は何もしない（手入力中のカーソル位置や通常入力を壊さない）
            // Only handle Up/Down; leave manual typing and caret behavior untouched
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            // 修飾キーは event を優先して読む（keyboardState は macOS で altKey を誤報するため）
            // Read modifiers from event first (keyboardState misreports altKey on macOS)
            var isShiftDown = event.shiftKey || (keyboard && keyboard.shiftKey);
            var isAltDown = event.altKey || (keyboard && keyboard.altKey);
            var delta = 1;

            if (isShiftDown) {
                // 10単位で増減 / change by 10
                delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (isAltDown) {
                // 0.1単位で増減 / change by 0.1
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                // 1単位 / change by 1
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            }

            // 丸め / rounding
            if (isAltDown) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = value;

            // 値変更後にプレビュー / update preview after change
            if (typeof onUpdate === "function") onUpdate();
        });
    }

    // =========================================
    // プレビュー履歴ユーティリティ / Preview history util
    // =========================================

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 他スクリプトでもこのブロックをコピペすれば再利用できます。
     * $.global に載せて共有するので、このスクリプトで使わない cancelTask も残しておく。
     * 使い方:
     *   PreviewHistory.start();      // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();       // プレビュー描画ごと（または操作ごと）にカウント(+1)
     *   PreviewHistory.undo();       // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */

    (function (g) {
        if (!g.PreviewHistory) {
            g.PreviewHistory = {
                start: function () { g.__previewUndoCount = 0; },
                bump: function () { g.__previewUndoCount = (g.__previewUndoCount | 0) + 1; },
                undo: function () {
                    var n = g.__previewUndoCount | 0;
                    try { for (var i = 0; i < n; i++) app.executeMenuCommand('undo'); } catch (e) { }
                    g.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

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
     * 現在のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* ラベル定義（カテゴリ分け）/ Label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "グリッドの間隔を再定義", en: "Redefine Grid Spacing" }
        },
        panel: {
            spacing: { ja: "間隔", en: "Spacing" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            horizontal: { ja: "左右", en: "H" },
            vertical: { ja: "上下", en: "V" }
        },
        checkbox: {
            link: { ja: "連動", en: "Link" },
            brick: { ja: "レンガ状", en: "Brick" },
            honeycomb: { ja: "ハニカム", en: "Honeycomb" },
            forceGrid: { ja: "強制グリッド", en: "Force Grid" },
            centerInCell: { ja: "中央揃え", en: "Center in cell" },
            transpose: { ja: "行列入れ替え", en: "Swap Rows{slash}Columns" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            horizontal: {
                ja: "左右に隣り合うオブジェクトのあいだにあける間隔です。↑↓キーで増減できます。",
                en: "Gap left between horizontally adjacent objects. The arrow keys step the value."
            },
            vertical: {
                ja: "上下に隣り合うオブジェクトのあいだにあける間隔です。↑↓キーで増減できます。",
                en: "Gap left between vertically adjacent objects. The arrow keys step the value."
            },
            link: {
                ja: "左右の間隔と同じ値を上下にも使います。オフにすると上下を個別に指定できます。",
                en: "Uses the horizontal gap for the vertical one too. Turn it off to set them separately."
            },
            brick: {
                ja: "1行ごとに半ピッチずらして、レンガ積みのように配置します。",
                en: "Offsets every other row by half a pitch, like a brick wall."
            },
            honeycomb: {
                ja: "レンガ状に加えて行送りを詰め、六角形に近い並びにします。",
                en: "Adds to the brick offset a tighter row step, giving a honeycomb-like arrangement."
            },
            forceGrid: {
                ja: "歯抜けや行ごとの個数違いがあっても、行数・列数をそろえた格子として並べ直します。",
                en: "Rebuilds the layout as an even grid even when rows have gaps or different counts."
            },
            centerInCell: {
                ja: "各セルの中でオブジェクトを天地左右中央にそろえます（強制グリッドのときだけ使えます）。",
                en: "Centers each object inside its cell. Available only with Force Grid."
            },
            transpose: {
                ja: "行と列を入れ替えて並べ直します。歯抜けのある配置にも対応します。",
                en: "Swaps rows and columns. Layouts with gaps are handled too."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            noSelection: { ja: "グリッド状に並んだオブジェクトを選択してください。", en: "Please select grid-like objects first." },
            needTwo: { ja: "2つ以上のオブジェクトを選択してください。", en: "Please select at least two objects." },
            cellConflict: {
                ja: "同一セルに複数オブジェクトが割り当てられました。\n許容値を下げるか、整列状態を確認してください。\n衝突セル: ",
                en: "Multiple objects were assigned to the same cell.\nReduce tolerances or check alignment.\nConflict cell: "
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す。"{slash}" は "/" に置き換える
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (labelNode == null) return labelPath;
        }
        var labelString = labelNode[uiLang] || labelNode.ja || labelNode.en || labelPath;
        return String(labelString).replace(/\{slash\}/g, "/");
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object|string} labelSet - ラベル、またはラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名と間隔の入力欄を1行追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {string} initialText - 入力欄の初期値
     * @param {string} tooltipPath - 入力欄の tooltip のラベルのパス
     * @returns {EditText} 追加した入力欄
     */
    function addGapField(parentGroup, labelPath, initialText, tooltipPath) {
        var gapRow = parentGroup.add('group');
        gapRow.add('statictext', undefined, labelText(labelPath));
        var gapInput = gapRow.add('edittext', undefined, initialText);
        gapInput.helpTip = getLabel(tooltipPath);
        gapInput.characters = GAP_INPUT_CHARS;
        return gapInput;
    }

    /**
     * オプションのチェックボックスを1行追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelPath - チェックボックスのラベルのパス
     * @param {string} tooltipPath - tooltip のラベルのパス
     * @param {boolean} isSubOption - サブオプションとして字下げするか
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, isSubOption) {
        var optionRow = parentPanel.add('group');
        setupRow(optionRow, 'left');
        if (isSubOption) optionRow.margins = SUB_OPTION_MARGINS;
        var optionCheckbox = optionRow.add('checkbox', undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        return optionCheckbox;
    }

    /**
     * 間隔設定ダイアログを組み立てる（イベント結線は呼び出し側で行う）
     * @param {number} initialGapX - 左右間隔の初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {object} ダイアログと各コントロールの参照
     */
    function buildGridSpacingDialog(initialGapX, rulerUnit) {
        // タイトルとバージョンを合成 / combine title and version
        var spacingDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(spacingDialog);

        // パネル名に単位を出す / show unit in panel title
        var spacingPanel = spacingDialog.add('panel', undefined, getLabel('panel.spacing') + ' (' + rulerUnit.label + ')');
        // 2カラム構成なので row のまま余白のみ共通化 / two-column panel: keep row, share margins
        spacingPanel.orientation = 'row';
        spacingPanel.alignChildren = 'top';
        spacingPanel.alignment = 'fill';
        spacingPanel.margins = PANEL_MARGINS;
        spacingPanel.spacing = COLUMN_SPACING;

        // 左カラム / left column
        var gapInputColumn = spacingPanel.add('group');
        gapInputColumn.orientation = 'column';
        gapInputColumn.alignChildren = 'left';

        // 初期値は pt を表示単位に換算して表示 / show initial value converted from pt to the display unit
        var horizontalGapInput = addGapField(gapInputColumn, 'fieldLabel.horizontal',
            (initialGapX / rulerUnit.pointsPerUnit).toFixed(1), 'tooltip.horizontal');
        // 初期表示では負の値を使わない / no negative value at first
        var verticalGapInput = addGapField(gapInputColumn, 'fieldLabel.vertical', '0', 'tooltip.vertical');

        // 右カラム（連動）/ right column (link)
        var linkColumn = spacingPanel.add('group');
        linkColumn.orientation = 'column';
        linkColumn.alignChildren = 'center';
        linkColumn.alignment = ['fill', 'fill'];
        var linkColumnSpacer = linkColumn.add('statictext', undefined, '');
        linkColumnSpacer.alignment = ['fill', 'fill'];
        var linkCheckbox = linkColumn.add('checkbox', undefined, getLabel('checkbox.link'));
        linkCheckbox.helpTip = getLabel('tooltip.link');

        // オプション（チェックボックスをまとめる）/ Options panel
        var optionsPanel = spacingDialog.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel, 6);

        // レンガ状／ハニカム（レンガ状のサブオプション）/ Brick and Honeycomb (sub-option of Brick)
        var brickCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.brick', 'tooltip.brick', false);
        brickCheckbox.value = false;
        var honeycombCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.honeycomb', 'tooltip.honeycomb', true);
        honeycombCheckbox.value = false;
        honeycombCheckbox.enabled = false;

        // 強制グリッド／中央揃え（強制グリッドのサブオプション）/ Force Grid and Center in cell (sub-option of Force Grid)
        var forceGridCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.forceGrid', 'tooltip.forceGrid', false);
        forceGridCheckbox.value = !!$.global.__regridForceGrid;
        var centerInCellCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.centerInCell', 'tooltip.centerInCell', true);
        centerInCellCheckbox.value = !!$.global.__regridCenterInCell;
        centerInCellCheckbox.enabled = forceGridCheckbox.value;

        // 行列入れ替え / Swap rows/columns
        var transposeCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.transpose', 'tooltip.transpose', false);
        transposeCheckbox.value = false;

        // 初期状態 / initial state
        linkCheckbox.value = true;
        horizontalGapInput.active = true;
        verticalGapInput.enabled = false;
        verticalGapInput.text = horizontalGapInput.text;

        honeycombCheckbox.enabled = brickCheckbox.value;

        // ボタン行（パネル外・中央寄せ、いっぱいに広げない）/ buttons (outside panels, centered, not stretched)
        var btnRowGroup = spacingDialog.add('group');
        setupRow(btnRowGroup, 'center');
        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        return {
            spacingDialog: spacingDialog,
            horizontalGapInput: horizontalGapInput,
            verticalGapInput: verticalGapInput,
            linkCheckbox: linkCheckbox,
            brickCheckbox: brickCheckbox,
            honeycombCheckbox: honeycombCheckbox,
            forceGridCheckbox: forceGridCheckbox,
            centerInCellCheckbox: centerInCellCheckbox,
            transposeCheckbox: transposeCheckbox
        };
    }

    /**
     * 左右・上下の入力値を数値で読む（数値でなければ 0）。単位は表示単位のまま
     * @param {object} dialogControls - buildGridSpacingDialog() の戻り値
     * @returns {{gapX: number, gapY: number}} 左右・上下の間隔
     */
    function readGapInputs(dialogControls) {
        var gapX = parseFloat(dialogControls.horizontalGapInput.text);
        var gapY = parseFloat(dialogControls.verticalGapInput.text);
        if (isNaN(gapX)) gapX = 0;
        if (isNaN(gapY)) gapY = 0;
        return { gapX: gapX, gapY: gapY };
    }

    /**
     * 間隔設定ダイアログを表示し、プレビューと確定適用を行う
     * @param {object} layoutActions - main が用意する配置処理一式
     *   applySpacing / applySpacingBrick / applySpacingHexagon（(gapX, gapY) => void）、
     *   transpose（行列入れ替え）、restoreInitialPositions（ダイアログ開始時点へ戻す）、
     *   resetBaselineToCurrent（現在位置を新しい基準にする）
     * @param {number} initialGapX - 左右間隔の初期値（pt）
     * @returns {void}
     */
    function showGridSpacingDialog(layoutActions, initialGapX) {
        // 表示単位↔pt の換算係数（入力値は表示単位、内部処理は pt）
        // Points per display unit (inputs are in display units; internal geometry is in pt)
        var rulerUnit = getUnitInfo();
        var unitFactor = rulerUnit.pointsPerUnit;

        var dialogControls = buildGridSpacingDialog(initialGapX, rulerUnit);
        var horizontalGapInput = dialogControls.horizontalGapInput;
        var verticalGapInput = dialogControls.verticalGapInput;
        var linkCheckbox = dialogControls.linkCheckbox;
        var brickCheckbox = dialogControls.brickCheckbox;
        var honeycombCheckbox = dialogControls.honeycombCheckbox;
        var forceGridCheckbox = dialogControls.forceGridCheckbox;
        var centerInCellCheckbox = dialogControls.centerInCellCheckbox;
        var transposeCheckbox = dialogControls.transposeCheckbox;

        // プレビュー用ヒストリー管理を開始 / Start preview history counter
        PreviewHistory.start();

        // 行列入れ替えが実行済みか（プレビュー状態）/ whether transpose has been triggered (preview state)
        var didTranspose = false;

        /**
         * 現在のオプション（レンガ／ハニカム／通常）に応じて間隔適用関数を選んで実行する
         * @param {number} gapX - 左右間隔
         * @param {number} gapY - 上下間隔
         * @returns {void}
         */
        function applySelectedSpacing(gapX, gapY) {
            if (brickCheckbox.value) {
                if (honeycombCheckbox.value) layoutActions.applySpacingHexagon(gapX, gapY);
                else layoutActions.applySpacingBrick(gapX, gapY);
            } else {
                layoutActions.applySpacing(gapX, gapY);
            }
        }

        /**
         * 転置の有無を考慮して間隔を適用する。
         * 転置ONのときはマージン0で並べ替え→転置→基準を取り直してから間隔を適用する
         * @param {number} gapX - 左右間隔
         * @param {number} gapY - 上下間隔
         * @param {boolean} bumpHistory - プレビュー時は各ステップで PreviewHistory.bump() する（確定時は false）
         * @returns {void}
         */
        function applyLayoutWithTranspose(gapX, gapY, bumpHistory) {
            if (didTranspose) {
                // 転置はマージン0で実行し、その後マージンを適用 / Transpose with 0 margins then apply margins
                layoutActions.applySpacing(0, 0);
                if (bumpHistory) PreviewHistory.bump();

                layoutActions.transpose();
                if (bumpHistory) PreviewHistory.bump();

                // 転置後の配置を新しい基準に / adopt transposed layout as baseline for spacing
                layoutActions.resetBaselineToCurrent();
                applySelectedSpacing(gapX, gapY);
                if (bumpHistory) PreviewHistory.bump();
            } else {
                applySelectedSpacing(gapX, gapY);
                if (bumpHistory) PreviewHistory.bump();
            }
        }

        /**
         * 現在のUI状態（間隔・各オプション）に基づいてプレビューを再描画する。
         * 直前のプレビューを一括Undoしてから再適用するため、ヒストリーを汚さない
         * @returns {void}
         */
        function updatePreview() {
            // 直前のプレビューを一括Undo（ヒストリーを汚さない）/ Undo previous preview
            PreviewHistory.undo();
            // 強制グリッド／中央揃えチェックボックスの状態をグローバルに反映
            $.global.__regridForceGrid = !!forceGridCheckbox.value;
            $.global.__regridCenterInCell = !!centerInCellCheckbox.value;
            // Undo後の現在位置を基準として originalPositions/layoutInfo を作り直す
            layoutActions.resetBaselineToCurrent();

            if (linkCheckbox.value) {
                verticalGapInput.enabled = false;
                verticalGapInput.text = horizontalGapInput.text;
            } else {
                verticalGapInput.enabled = true;
            }

            // 値は連動OFF時に UI 側（transposeCheckboxのonClick）で入れ替え済みのため、ここでは追加の入れ替えをしない
            // H/V are already swapped in the UI (transpose onClick) when Link is OFF, so do NOT swap again here
            var gapValues = readGapInputs(dialogControls);

            // 表示単位 → pt に換算して適用 / convert display units to pt before applying
            applyLayoutWithTranspose(gapValues.gapX * unitFactor, gapValues.gapY * unitFactor, true);
        }

        // イベント / events
        changeValueByArrowKey(horizontalGapInput, updatePreview);
        changeValueByArrowKey(verticalGapInput, updatePreview);
        horizontalGapInput.onChanging = function () { updatePreview(); };
        verticalGapInput.onChanging = function () {
            if (!linkCheckbox.value) {
                updatePreview();
            }
        };
        linkCheckbox.onClick = function () { updatePreview(); };
        forceGridCheckbox.onClick = function () {
            // 強制グリッドOFF時は中央揃えも無効化 / disable center when Force Grid is off
            centerInCellCheckbox.enabled = forceGridCheckbox.value;
            if (!forceGridCheckbox.value) centerInCellCheckbox.value = false;
            updatePreview();
        };
        centerInCellCheckbox.onClick = function () { updatePreview(); };
        // 行列入れ替え / swap rows & columns
        transposeCheckbox.onClick = function () {
            // トグル：ONで転置、OFFで直前（転置前）の状態に戻す
            // Toggle: ON = transpose, OFF = revert to the pre-transpose state
            didTranspose = transposeCheckbox.value;

            // 連動OFFなら左右/上下の値をUI上でも入れ替える（OFFでは再度入れ替えて元へ戻す）
            // Swap H/V UI values when Link is OFF (swapping again on OFF restores them)
            if (!linkCheckbox.value) {
                var swapHorizontalText = horizontalGapInput.text;
                horizontalGapInput.text = verticalGapInput.text;
                verticalGapInput.text = swapHorizontalText;
            }

            updatePreview();
        };

        // レンガ状 / Brick
        brickCheckbox.onClick = function () {
            honeycombCheckbox.enabled = brickCheckbox.value;
            if (!brickCheckbox.value) honeycombCheckbox.value = false;
            updatePreview();
        };

        // 六角形 / Hexagon
        honeycombCheckbox.onClick = function () {
            updatePreview();
        };

        // 開いたときに一度プレビュー / first preview when opened
        updatePreview();

        var dialogResult = dialogControls.spacingDialog.show();

        if (dialogResult == 1) {
            // OK時は最終値で適用 / apply with final values
            var finalGaps = readGapInputs(dialogControls);
            if (linkCheckbox.value) finalGaps.gapY = finalGaps.gapX;

            // 値は連動OFF時に UI 側で入れ替え済みのため、確定時も追加の入れ替えはしない
            // H/V are already swapped in the UI when Link is OFF, so no extra swap on final apply

            // プレビュー分を一括Undoしてから確定適用 / Clear preview history before final apply
            PreviewHistory.undo();
            layoutActions.resetBaselineToCurrent();
            // 表示単位 → pt に換算して適用 / convert display units to pt before applying
            applyLayoutWithTranspose(finalGaps.gapX * unitFactor, finalGaps.gapY * unitFactor, false);
        } else {
            // キャンセル時：プレビュー分を一括Undoしてから初期状態へ / Undo preview then restore
            PreviewHistory.undo();
            layoutActions.restoreInitialPositions();
        }

        // 念のためカウンタ初期化 / reset counter
        PreviewHistory.start();
    }

    // =========================================
    // レイアウト計算の補助 / Layout helpers
    // =========================================

    /**
     * レイアウト計算に使う外接矩形を返す。
     * すでにグループになっているものは中身を分解せず「グループ＝1つのオブジェクト」として扱う。
     * クリップグループはクリップパスの geometricBounds（＝可視領域）を優先し、
     * それ以外は item.geometricBounds を使う
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]} [left, top, right, bottom]
     */
    function getLayoutBounds(item) {
        /* クリップグループの中を読めないときは、グループ自体の外接矩形に戻す / Fall back to the group's own bounds */
        try {
            if (item.typename === 'GroupItem' && item.clipped) {
                // GroupItem の中から clipping パスを探す / look for the clipping path
                if (item.pathItems) {
                    for (var i = 0; i < item.pathItems.length; i++) {
                        if (item.pathItems[i].clipping) return item.pathItems[i].geometricBounds;
                    }
                }
                // CompoundPath が clipping のケース / compound path used as the mask
                if (item.compoundPathItems) {
                    for (var j = 0; j < item.compoundPathItems.length; j++) {
                        var compoundPath = item.compoundPathItems[j];
                        if (compoundPath.pathItems && compoundPath.pathItems.length > 0 && compoundPath.pathItems[0].clipping) {
                            return compoundPath.pathItems[0].geometricBounds;
                        }
                    }
                }
            }
        } catch (e) { }

        return item.geometricBounds;
    }

    /**
     * 各オブジェクトの左上（getLayoutBounds 基準）を控える
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Array<Object>} item / left / top を持つ位置情報の配列
     */
    function snapshotPositions(selectedItems) {
        var positionSnapshot = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = getLayoutBounds(selectedItems[i]);
            positionSnapshot.push({ item: selectedItems[i], left: itemBounds[0], top: itemBounds[1] });
        }
        return positionSnapshot;
    }

    /**
     * 控えた位置情報を複製する（元の控えを書き換えても影響しないように）
     * @param {Array<Object>} positionSnapshot - item / left / top を持つ位置情報の配列
     * @returns {Array<Object>} 複製した位置情報の配列
     */
    function clonePositions(positionSnapshot) {
        var clonedSnapshot = [];
        for (var i = 0; i < positionSnapshot.length; i++) {
            clonedSnapshot.push({ item: positionSnapshot[i].item, left: positionSnapshot[i].left, top: positionSnapshot[i].top });
        }
        return clonedSnapshot;
    }

    /**
     * 控えておいた位置へオブジェクトを戻す
     * @param {Array<Object>} positionSnapshot - item / left / top を持つ位置情報の配列
     * @returns {void}
     */
    function restorePositions(positionSnapshot) {
        for (var i = 0; i < positionSnapshot.length; i++) {
            var snapshotEntry = positionSnapshot[i];
            var currentBounds = getLayoutBounds(snapshotEntry.item);
            snapshotEntry.item.translate(snapshotEntry.left - currentBounds[0], snapshotEntry.top - currentBounds[1]);
        }
    }

    /**
     * 位置情報の配列から、boundsList の要素の元座標（左上）を引く。
     * 見つからなければ現在の bbox（currentEntry.bounds）を使う
     * @param {Array<Object>} originalPositions - item / left / top を持つ位置情報の配列
     * @param {object} currentEntry - layoutInfo.boundsList の要素（{item, bounds}）
     * @returns {{left: number, top: number}} 元の左端X・上端Y
     */
    function findOriginalLeftTop(originalPositions, currentEntry) {
        for (var i = 0; i < originalPositions.length; i++) {
            if (originalPositions[i].item === currentEntry.item) {
                return { left: originalPositions[i].left, top: originalPositions[i].top };
            }
        }
        return { left: currentEntry.bounds[0], top: currentEntry.bounds[1] };
    }

    /**
     * 選択の並びから左右間隔の初期値を測る（上→下、同じ行は左→右で並べた先頭2つの間隔）
     * @param {PageItem[]} selectedItems - 対象のオブジェクト（2つ以上）
     * @returns {number} 左右間隔（pt）。負の値は 0
     */
    function measureInitialGapX(selectedItems) {
        // 上端の降順、同じ行（差1pt未満）は左端の昇順 / Sort by top descending and left ascending within same row
        var sortedByPosition = selectedItems.slice().sort(function (itemA, itemB) {
            var boundsA = getLayoutBounds(itemA);
            var boundsB = getLayoutBounds(itemB);
            if (Math.abs(boundsA[1] - boundsB[1]) < 1) {
                return boundsA[0] - boundsB[0]; // same row → compare left
            }
            return boundsB[1] - boundsA[1]; // sort by top descending
        });

        var firstBounds = getLayoutBounds(sortedByPosition[0]);
        var secondBounds = getLayoutBounds(sortedByPosition[1]);
        // ダイアログ初期値では負の値を使わない / no negative value in the dialog
        var initialGapX = secondBounds[0] - firstBounds[2];
        return (initialGapX < 0) ? 0 : initialGapX;
    }

    /**
     * 昇順に並んだ数値配列の、隣接する差分の中央値を返す
     * @param {number[]} sortedAsc - 昇順に並んだ数値配列
     * @returns {number} 隣接差分の中央値。要素が2未満なら0
     */
    function medianAdjacentDiff(sortedAsc) {
        if (!sortedAsc || sortedAsc.length < 2) return 0;
        var diffs = [];
        for (var i = 1; i < sortedAsc.length; i++) {
            diffs.push(Math.abs(sortedAsc[i] - sortedAsc[i - 1]));
        }
        diffs.sort(function (a, b) { return a - b; });
        return diffs[Math.floor(diffs.length / 2)];
    }

    /**
     * 選択オブジェクトの外接境界一覧と、最小の幅・高さを収集する
     * （buildLayoutInfo / buildLayoutInfoForceGrid の共通前処理）
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {{boundsList: Array, minWidth: number, minHeight: number}} {item, bounds} の配列と、最小の幅・高さ
     */
    function collectBoundsList(selectedItems) {
        var boundsList = [];
        var minWidth = Number.MAX_VALUE;
        var minHeight = Number.MAX_VALUE;

        for (var j = 0; j < selectedItems.length; j++) {
            var itemBounds = getLayoutBounds(selectedItems[j]);
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];
            if (itemWidth < minWidth) minWidth = itemWidth;
            if (itemHeight < minHeight) minHeight = itemHeight;
            boundsList.push({ item: selectedItems[j], bounds: itemBounds });
        }
        return { boundsList: boundsList, minWidth: minWidth, minHeight: minHeight };
    }

    /**
     * 行・列をまとめる許容差を返す（最小の幅・高さの半分。1pt 未満にはしない）
     * @param {number} minSize - 選択中の最小の幅または高さ
     * @returns {number} 許容差（pt）
     */
    function getClusterTolerance(minSize) {
        var tolerance = minSize * 0.5;
        return (tolerance < 1) ? 1 : tolerance;
    }

    /**
     * 並べ済みの境界レコードを、座標の近いものどうしでまとめる。
     * 既存のまとまりに入るたびに、まとまりの座標を平均へ更新する
     * @param {Array<Object>} sortedRecords - {item, bounds} の配列（まとめる順に並べたもの）
     * @param {number} boundsIndex - 比べる境界の要素（0 = 左端、1 = 上端）
     * @param {string} centerKey - まとまりの座標を入れるキー（"x" / "y"）
     * @param {number} tolerance - 同じまとまりとみなす許容差
     * @returns {Array<Object>} まとまり（centerKey の座標と members）の配列
     */
    function clusterBoundsRecords(sortedRecords, boundsIndex, centerKey, tolerance) {
        var clusters = [];
        for (var i = 0; i < sortedRecords.length; i++) {
            var boundsRecord = sortedRecords[i];
            var coordinate = boundsRecord.bounds[boundsIndex];
            var isMerged = false;
            for (var c = 0; c < clusters.length; c++) {
                if (Math.abs(clusters[c][centerKey] - coordinate) <= tolerance) {
                    clusters[c].members.push(boundsRecord);
                    clusters[c][centerKey] = (clusters[c][centerKey] * (clusters[c].members.length - 1) + coordinate) / clusters[c].members.length;
                    isMerged = true;
                    break;
                }
            }
            if (!isMerged) {
                var newCluster = { members: [boundsRecord] };
                newCluster[centerKey] = coordinate;
                clusters.push(newCluster);
            }
        }
        return clusters;
    }

    /**
     * 境界レコードの中で最大の幅または高さを返す
     * @param {Array<Object>} boundsRecords - {item, bounds} の配列
     * @param {boolean} measureWidth - true なら幅、false なら高さ
     * @returns {number} 最大の幅または高さ
     */
    function getMaxExtent(boundsRecords, measureWidth) {
        var maxExtent = 0;
        for (var i = 0; i < boundsRecords.length; i++) {
            var recordBounds = boundsRecords[i].bounds;
            var extent = measureWidth ? (recordBounds[2] - recordBounds[0]) : (recordBounds[1] - recordBounds[3]);
            if (extent > maxExtent) maxExtent = extent;
        }
        return maxExtent;
    }

    /**
     * 選択オブジェクトの現在位置から、行と列の構成を推定する
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Object} 行・列の中心座標と、各オブジェクトの行列位置を持つレイアウト情報
     */
    function buildLayoutInfo(selectedItems) {
        var collected = collectBoundsList(selectedItems);
        var boundsList = collected.boundsList;

        // 列まとめ（左→右）/ group columns (left -> right)
        boundsList.sort(function (recordA, recordB) { return recordA.bounds[0] - recordB.bounds[0]; });
        var colCenters = clusterBoundsRecords(boundsList, 0, "x", getClusterTolerance(collected.minWidth));

        // 行まとめ（上→下）/ group rows (top -> bottom)
        var boundsSortedByTop = boundsList.slice().sort(function (recordA, recordB) { return recordB.bounds[1] - recordA.bounds[1]; });
        var rowCenters = clusterBoundsRecords(boundsSortedByTop, 1, "y", getClusterTolerance(collected.minHeight));

        // 並び順の確定 / sort
        colCenters.sort(function (a, b) { return a.x - b.x; });
        rowCenters.sort(function (a, b) { return b.y - a.y; });

        // 各列の最大幅・各行の最大高さ / max width per column, max height per row
        var colWidths = [];
        for (var ci = 0; ci < colCenters.length; ci++) colWidths.push(getMaxExtent(colCenters[ci].members, true));
        var rowHeights = [];
        for (var ri = 0; ri < rowCenters.length; ri++) rowHeights.push(getMaxExtent(rowCenters[ri].members, false));

        return {
            boundsList: boundsList,
            colCenters: colCenters,
            rowCenters: rowCenters,
            colWidths: colWidths,
            rowHeights: rowHeights,
            baseX: colCenters[0].x,
            baseY: rowCenters[0].y
        };
    }

    /**
     * 選択オブジェクトを強制的に格子とみなして、行と列の構成を組み立てる。
     * 行ごと（上→下）に左→右で(行,列)を割り当てる：行は上端Yの近さでまとめ、各行の中は左端Xで並べる。
     * 欠け（歯抜け）は許容する（行ごとに列数が異なってよい）
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Object} 行・列の中心座標と、各オブジェクトの行列位置を持つレイアウト情報
     */
    function buildLayoutInfoForceGrid(selectedItems) {
        var collected = collectBoundsList(selectedItems);
        var boundsList = collected.boundsList;

        // 行まとめ（上→下）/ group rows (top -> bottom)
        var boundsSortedByTop = boundsList.slice().sort(function (recordA, recordB) { return recordB.bounds[1] - recordA.bounds[1]; });
        var rows = clusterBoundsRecords(boundsSortedByTop, 1, "y", getClusterTolerance(collected.minHeight));
        rows.sort(function (a, b) { return b.y - a.y; });

        // 各行の中を左→右で確定し、rowIndex/colIndexを付与 / sort within row and assign indices
        var maxCols = 0;
        for (var r = 0; r < rows.length; r++) {
            rows[r].members.sort(function (recordA, recordB) { return recordA.bounds[0] - recordB.bounds[0]; });
            if (rows[r].members.length > maxCols) maxCols = rows[r].members.length;
            for (var c = 0; c < rows[r].members.length; c++) {
                rows[r].members[c].rowIndex = r;
                rows[r].members[c].colIndex = c;
            }
        }

        // colWidths（列番号ごとの最大幅）/ max width per column index
        var colWidths = [];
        for (var colIndex = 0; colIndex < maxCols; colIndex++) {
            var maxWidth = 0;
            for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
                if (rows[rowIndex].members.length > colIndex) {
                    var memberBounds = rows[rowIndex].members[colIndex].bounds;
                    var memberWidth = memberBounds[2] - memberBounds[0];
                    if (memberWidth > maxWidth) maxWidth = memberWidth;
                }
            }
            colWidths.push(maxWidth);
        }

        // rowHeights（行ごとの最大高さ）/ max height per row
        var rowHeights = [];
        for (var ri = 0; ri < rows.length; ri++) rowHeights.push(getMaxExtent(rows[ri].members, false));

        // baseX/baseY は最左/最上 / baseX/baseY = top-left
        var baseX = Number.MAX_VALUE;
        var baseY = Number.MIN_VALUE;
        for (var q = 0; q < boundsList.length; q++) {
            if (boundsList[q].bounds[0] < baseX) baseX = boundsList[q].bounds[0];
            if (boundsList[q].bounds[1] > baseY) baseY = boundsList[q].bounds[1];
        }

        // ダミーのcolCenters/rowCenters（互換のため）/ dummy centers (compat)
        var colCenters = [];
        for (var cc = 0; cc < maxCols; cc++) colCenters.push({ x: baseX, members: [] });
        var rowCenters = [];
        for (var rr = 0; rr < rows.length; rr++) rowCenters.push({ y: rows[rr].y, members: rows[rr].members });

        return {
            boundsList: boundsList,
            colCenters: colCenters,
            rowCenters: rowCenters,
            colWidths: colWidths,
            rowHeights: rowHeights,
            baseX: baseX,
            baseY: baseY
        };
    }

    /**
     * ［強制グリッド］の状態に合わせてレイアウト情報を作る
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {Object} レイアウト情報
     */
    function buildCurrentLayoutInfo(selectedItems) {
        return ($.global.__regridForceGrid) ? buildLayoutInfoForceGrid(selectedItems) : buildLayoutInfo(selectedItems);
    }

    /**
     * 対象の列・行インデックスを解決する。
     * Force Grid で事前割り当て済み（colIndex/rowIndex）ならそれを優先し、
     * なければ列／行センターへの最近傍で推定する
     * @param {object} currentEntry - layoutInfo.boundsList の要素
     * @param {Array} colCenters - 列センター配列
     * @param {Array} rowCenters - 行センター配列
     * @returns {{col: number, row: number}} 列・行インデックス
     */
    function resolveColRow(currentEntry, colCenters, rowCenters) {
        var colIndex = (typeof currentEntry.colIndex === 'number') ? currentEntry.colIndex : null;
        var rowIndex = (typeof currentEntry.rowIndex === 'number') ? currentEntry.rowIndex : null;

        if (colIndex === null) {
            colIndex = 0;
            var minDistanceX = Number.MAX_VALUE;
            for (var c = 0; c < colCenters.length; c++) {
                var dx = Math.abs(colCenters[c].x - currentEntry.bounds[0]);
                if (dx < minDistanceX) { minDistanceX = dx; colIndex = c; }
            }
        }

        if (rowIndex === null) {
            rowIndex = 0;
            var minDistanceY = Number.MAX_VALUE;
            for (var r = 0; r < rowCenters.length; r++) {
                var dy = Math.abs(rowCenters[r].y - currentEntry.bounds[1]);
                if (dy < minDistanceY) { minDistanceY = dy; rowIndex = r; }
            }
        }

        return { col: colIndex, row: rowIndex };
    }

    /**
     * レンガ／ハニカムで使う半ピッチを算出する。
     * 目標列ピッチ＝中央値幅 + gapX。推定できない場合は列センター間隔の中央値を使う
     * @param {number[]} colWidths - 列ごとの最大幅
     * @param {Array} colCenters - 列センター配列
     * @param {number} gapX - 左右間隔
     * @returns {number} 半ピッチ（pitch / 2）
     */
    function computeHalfPitch(colWidths, colCenters, gapX) {
        var sortedColWidths = [];
        if (colWidths && colWidths.length > 0) {
            for (var i = 0; i < colWidths.length; i++) sortedColWidths.push(colWidths[i]);
            sortedColWidths.sort(function (a, b) { return a - b; });
        }
        var medianColWidth = (sortedColWidths.length > 0) ? sortedColWidths[Math.floor(sortedColWidths.length / 2)] : 0; // median width
        var pitch = (medianColWidth > 0) ? (medianColWidth + gapX) : 0;

        // fallback：列が1つ等で推定できない場合は現状の列位置差 / fallback to current centers diff
        if (pitch === 0) {
            var colLefts = [];
            for (var j = 0; j < colCenters.length; j++) colLefts.push(colCenters[j].x);
            colLefts.sort(function (a, b) { return a - b; });
            pitch = medianAdjacentDiff(colLefts);
        }
        return pitch / 2.0;
    }

    /**
     * グリッド配置を適用する共通処理。通常／レンガ／ハニカムを引数で切り替える。
     * 直前にプレビュー基準位置へ戻してから、列幅・行高さの累積で再配置する
     * @param {Object} layoutInfo - レイアウト情報
     * @param {Array<Object>} originalPositions - プレビュー基準の位置情報
     * @param {number} gapX - 左右間隔
     * @param {number} gapY - 上下間隔
     * @param {boolean} isBrick - 奇数行を半ピッチ横にずらすか（レンガ／ハニカム）
     * @param {number} rowStepFactor - 行送りに掛ける係数（通常・レンガ=1.0、ハニカム=HONEYCOMB_ROW_STEP_FACTOR）
     * @returns {void}
     */
    function applyGridLayout(layoutInfo, originalPositions, gapX, gapY, isBrick, rowStepFactor) {
        // いったん元に戻す / restore first
        restorePositions(originalPositions);

        var boundsList = layoutInfo.boundsList;
        var colWidths = layoutInfo.colWidths;
        var rowHeights = layoutInfo.rowHeights;

        var halfPitch = isBrick ? computeHalfPitch(colWidths, layoutInfo.colCenters, gapX) : 0;

        // 中央揃え：各セルの天地左右中央に整列（強制グリッドのサブオプションなので Force Grid 時のみ有効）
        // Center each object within its cell (sub-option of Force Grid, so only when Force Grid is on)
        var centerInCell = !!$.global.__regridCenterInCell && !!$.global.__regridForceGrid;

        for (var i = 0; i < boundsList.length; i++) {
            var currentEntry = boundsList[i];

            var originalPos = findOriginalLeftTop(originalPositions, currentEntry);
            var cellIndex = resolveColRow(currentEntry, layoutInfo.colCenters, layoutInfo.rowCenters);
            var colIndex = cellIndex.col;
            var rowIndex = cellIndex.row;

            // 新しいX（セル左端）/ new X (cell left)
            var newX = layoutInfo.baseX;
            for (var c = 0; c < colIndex; c++) {
                newX += colWidths[c] + gapX;
            }

            // 新しいY（セル上端）/ new Y (cell top)
            var newY = layoutInfo.baseY;
            for (var r = 0; r < rowIndex; r++) {
                newY -= (rowHeights[r] * rowStepFactor) + gapY;
            }

            // 中央揃え：セル内でオブジェクトを左右・天地中央へ寄せる
            // Center within the cell (cell size = column width × row height)
            if (centerInCell) {
                var itemWidth = currentEntry.bounds[2] - currentEntry.bounds[0];
                var itemHeight = currentEntry.bounds[1] - currentEntry.bounds[3];
                newX += (colWidths[colIndex] - itemWidth) / 2;
                newY -= (rowHeights[rowIndex] - itemHeight) / 2;
            }

            // レンガ状：奇数行を半ピッチずらす / Brick: shift odd rows by half pitch
            if (isBrick && halfPitch !== 0 && (rowIndex % 2 === 1)) {
                newX += halfPitch;
            }

            currentEntry.item.translate(newX - originalPos.left, newY - originalPos.top);
        }

        // 再描画 / redraw
        app.redraw();
    }

    // =========================================
    // 行列入れ替え / Transpose
    // getLayoutBounds() の left/top を基準に、行・列をクラスタリングして推定し、
    // 推定したピッチ（隣接差の中央値）で左上基準に再配置する（歯抜け対応）。
    // グループは中身を見ず、グループ全体の外接 bbox（クリップグループはクリップパス）で1オブジェクトとして扱う。
    // 1行→1列 / 1列→1行 も対応（ピッチ流用）
    // =========================================

    /**
     * 近い値どうしをまとめて、クラスタの中心値の配列を作る（coordinateValues は昇順に並べ替わる）
     * @param {number[]} coordinateValues - まとめる対象の値
     * @param {number} tolerance - 同じクラスタとみなす許容差
     * @returns {number[]} 昇順に並べたクラスタ中心値の配列
     */
    function clusterValues(coordinateValues, tolerance) {
        coordinateValues.sort(function (a, b) { return a - b; });
        var clusterCenters = [];
        for (var i = 0; i < coordinateValues.length; i++) {
            var coordinate = coordinateValues[i];
            var foundIndex = -1;
            for (var c = 0; c < clusterCenters.length; c++) {
                if (Math.abs(coordinate - clusterCenters[c]) <= tolerance) { foundIndex = c; break; }
            }
            if (foundIndex < 0) clusterCenters.push(coordinate);
            else clusterCenters[foundIndex] = (clusterCenters[foundIndex] + coordinate) / 2.0;
        }
        clusterCenters.sort(function (a, b) { return a - b; });
        return clusterCenters;
    }

    /**
     * 中心値の配列から、指定した値に最も近い要素の位置を返す
     * @param {number[]} sortedCenters - 並べた中心値の配列
     * @param {number} coordinate - 探す値
     * @returns {number} 最も近い要素のインデックス
     */
    function findNearestIndex(sortedCenters, coordinate) {
        var bestIndex = 0;
        var bestDistance = Math.abs(coordinate - sortedCenters[0]);
        for (var i = 1; i < sortedCenters.length; i++) {
            var distance = Math.abs(coordinate - sortedCenters[i]);
            if (distance < bestDistance) { bestDistance = distance; bestIndex = i; }
        }
        return bestIndex;
    }

    /**
     * 各オブジェクトを(行,列)に割り当てる。同じセルに2つ入ったらアラートを出して中止する
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @param {number[]} rowClusters - 行の中心値（上→下）
     * @param {number[]} colClusters - 列の中心値（左→右）
     * @returns {Array<Object>|null} {item, row, col} の配列。衝突したら null
     */
    function mapItemsToCells(selectedItems, rowClusters, colClusters) {
        var occupiedCells = {}; // cellKey "row,col" -> item
        var cellMapping = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var targetItem = selectedItems[i];
            var itemBounds = getLayoutBounds(targetItem);
            var rowIndex = findNearestIndex(rowClusters, itemBounds[1]);
            var colIndex = findNearestIndex(colClusters, itemBounds[0]);
            var cellKey = rowIndex + "," + colIndex;
            if (occupiedCells[cellKey]) {
                alert(getLabel('alert.cellConflict') + "(" + rowIndex + "," + colIndex + ")");
                return null;
            }
            occupiedCells[cellKey] = targetItem;
            cellMapping.push({ item: targetItem, row: rowIndex, col: colIndex });
        }
        return cellMapping;
    }

    /**
     * 転置後に使う横・縦のピッチを求める（元の列・行の隣接差の中央値）。
     * 1行しかないときは横ピッチを縦にも、1列しかないときは縦ピッチを横にも流用する
     * @param {number[]} rowClusters - 行の中心値
     * @param {number[]} colClusters - 列の中心値
     * @returns {{x: number, y: number}|null} ピッチ。推定できなければ null
     */
    function getTransposePitch(rowClusters, colClusters) {
        var rowCount = rowClusters.length;
        var colCount = colClusters.length;
        // ※行や列が1つしかない場合は0になる / 0 when there is only one row or column
        var pitchX = (colCount >= 2) ? medianAdjacentDiff(colClusters) : 0;
        var pitchY = (rowCount >= 2) ? medianAdjacentDiff(rowClusters) : 0;

        // ほぼ重なり等で1セル扱いになったケース / everything collapsed into one cell
        if (rowCount === 1 && colCount === 1) return null;

        if (rowCount === 1 && colCount > 1) {
            // 1行 → 1列：横ピッチを縦へ流用 / one row -> one column
            return (pitchX === 0) ? null : { x: pitchX, y: pitchX };
        }
        if (colCount === 1 && rowCount > 1) {
            // 1列 → 1行：縦ピッチを横へ流用 / one column -> one row
            return (pitchY === 0) ? null : { x: pitchY, y: pitchY };
        }
        // 通常（2行以上 かつ 2列以上）/ two or more rows and columns
        return (pitchX === 0 || pitchY === 0) ? null : { x: pitchX, y: pitchY };
    }

    /**
     * 歯抜けを許容したままグリッドを転置（行⇄列）する
     * @param {PageItem[]} selectedItems - 対象のオブジェクト
     * @returns {void}
     */
    function transposeGridWithHoles(selectedItems) {
        if (!selectedItems || selectedItems.length < 1) return;

        // left/top を集める / collect left/top
        var leftValues = [], topValues = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = getLayoutBounds(selectedItems[i]);
            leftValues.push(itemBounds[0]);
            topValues.push(itemBounds[1]);
        }

        var colClusters = clusterValues(leftValues, TRANSPOSE_SNAP_X_TOLERANCE); // left -> right
        var rowClusters = clusterValues(topValues, TRANSPOSE_SNAP_Y_TOLERANCE); // will sort top -> bottom next

        // Illustrator座標では上ほどYが大きいことが多いので「上→下」/ sort top -> bottom
        rowClusters.sort(function (a, b) { return b - a; });

        var cellMapping = mapItemsToCells(selectedItems, rowClusters, colClusters);
        if (!cellMapping) return;

        var pitch = getTransposePitch(rowClusters, colClusters);
        if (!pitch) return;

        // 転置後グリッドの基準（左上固定）/ origin at top-left
        var originLeft = colClusters[0];
        var originTop = rowClusters[0];

        // 転置: 新しい列 = 元の行、新しい行 = 元の列 / newCol = oldRow, newRow = oldCol
        for (var t = 0; t < cellMapping.length; t++) {
            var mappedItem = cellMapping[t].item;
            var targetLeft = originLeft + cellMapping[t].row * pitch.x;
            var targetTop = originTop - cellMapping[t].col * pitch.y;

            var currentBounds = getLayoutBounds(mappedItem);
            mappedItem.translate(targetLeft - currentBounds[0], targetTop - currentBounds[1]);
        }

        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを収集し、間隔設定ダイアログを起動する
     * @returns {void}
     */
    function main() {
        // ドキュメントチェック / document check
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel('alert.noSelection'));
            return;
        }

        // 選択を拾う / collect selection
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }
        if (selectedItems.length < 2) {
            alert(getLabel('alert.needTwo'));
            return;
        }

        var initialGapX = measureInitialGapX(selectedItems);

        // プレビューの基準位置と、ダイアログ開始時点の位置（キャンセルで必ずここへ戻す）
        // Preview baseline, and the snapshot Cancel always returns to
        var originalPositions = snapshotPositions(selectedItems);
        var initialPositions = clonePositions(originalPositions);

        var layoutInfo = buildCurrentLayoutInfo(selectedItems);

        /**
         * 現在の位置を新しい基準として控え直し、レイアウト情報を作り直す
         * @returns {void}
         */
        function resetBaselineToCurrent() {
            originalPositions = snapshotPositions(selectedItems);
            layoutInfo = buildCurrentLayoutInfo(selectedItems);
        }

        // ダイアログ表示 / show dialog
        showGridSpacingDialog({
            /* 間隔を適用（通常グリッド）/ apply spacing (normal grid) */
            applySpacing: function (gapX, gapY) {
                applyGridLayout(layoutInfo, originalPositions, gapX, gapY, false, 1.0);
            },
            /* 間隔を適用（レンガ状）/ apply spacing (brick layout) */
            applySpacingBrick: function (gapX, gapY) {
                applyGridLayout(layoutInfo, originalPositions, gapX, gapY, true, 1.0);
            },
            /* 間隔を適用（六角形/ハニカム）/ apply spacing (hexagon/honeycomb layout) */
            applySpacingHexagon: function (gapX, gapY) {
                applyGridLayout(layoutInfo, originalPositions, gapX, gapY, true, HONEYCOMB_ROW_STEP_FACTOR);
            },
            transpose: function () {
                transposeGridWithHoles(selectedItems);
            },
            /* キャンセル時に戻す（ダイアログ開始時点）/ restore initial positions on Cancel */
            restoreInitialPositions: function () {
                restorePositions(initialPositions);
            },
            resetBaselineToCurrent: resetBaselineToCurrent
        }, initialGapX);
    }

    // 実行 / run
    main();

})();
