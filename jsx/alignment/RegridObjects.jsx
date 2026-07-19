#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択中のオブジェクトが「だいたいグリッド状」に並んでいることを前提に、左右・上下の間隔で再配置します。
- 常時プレビュー対応。値は EditText で直接入力でき、↑↓ キー（Shift で ×10、Option で ×0.1）でも増減できます。
- 間隔の値は現在の定規単位（mm／pt／px など）で入力でき、内部では pt に換算して処理します（パネル名に単位を表示）。
- すでにグループになっているもの（クリップグループ含む）は、中身を分解せず「1 つのオブジェクト＝1 つの外接 bbox」として扱います。
- 実行後に自動でグループ化はしません（選択状態のまま）。
- ダイアログは日本語／英語の自動切り替えに対応（$.locale）。
- ［連動］で左右の値を上下に反映できます。
- ［レンガ状］：行ごとに左右位置を半ピッチずらして再配置します。
- ［ハニカム］：［レンガ状］と併用し、奇数行を（幅＋左右の値）の半分だけ横にずらし、行の高さだけ 0.75 倍にしてハニカム状にします（上下の値はそのまま反映）。
- ［強制グリッド］：位置の近さで列／行を推定せず、上→下の行ごとに左→右の順で (行, 列) を割り当てて再配置します。
- ［中央揃え］（［強制グリッド］のサブオプション）：各セル（列幅×行高さ）の天地左右中央にオブジェクトを整列します。
- ［行列入れ替え］はトグルです。ON で歯抜け（欠け）を許容しつつ行⇄列を転置し、OFF で転置前の状態に戻します。
- 1 行だけ→1 列、1 列だけ→1 行の転置にも対応します。

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RegridObjects.jsx

### 紹介記事（note）

https://note.com/dtp_tranist/n/n08861d0e40c3

### 更新履歴

- 2026-07-08: v1.6.0 中央揃え（強制グリッドのサブオプション）を追加、行列入れ替えをトグル化（OFFで転置前に戻る）、定規単位（mm/pt/px 等）での入力に対応（内部で pt 換算）、apply系関数の共通化・命名整理などコード整理
- 2025-10-31: v1.0 プレビューの常時 ON、連動（上下ディム）、↑↓ キーでの増減

*/

/*

### Overview

- Assumes the selected objects are arranged in an "approximately grid-like" layout.
- Always-on preview. Values can be entered directly in the EditText fields, and adjusted via ↑↓ keys (×10 with Shift, ×0.1 with Option).
- Gap values are entered in the current ruler unit (mm / pt / px, etc.) and converted to points internally (the unit is shown in the panel title).
- Existing groups (including clipped groups) are treated as a single object — one outer bounding box — and never broken apart.
- Does not auto-group results (keeps items selected).
- Repositions them using the horizontal/vertical gaps entered in the dialog.
- The dialog supports automatic Japanese/English switching (based on $.locale).
- The "Link" option mirrors the horizontal value to the vertical value.
- "Brick": shifts every other row by a half pitch.
- "Honeycomb": used with "Brick"; shifts odd rows by half of the target cell pitch (width + H gap) and compresses only row height (0.75×) while keeping the V gap as-is.
- "Force Grid": assigns (row, col) by row-major order instead of inferring nearest columns/rows by position.
- "Center in cell" (sub-option of "Force Grid"): centers each object within its cell (column width × row height).
- "Swap Rows/Columns": a toggle — ON transposes rows/columns while tolerating missing cells (including 1 row ⇄ 1 column); OFF reverts to the pre-transpose state.

### Update history

- 2026-07-08: v1.6.0 Add "Center in cell" (Force Grid sub-option), make Swap Rows/Columns a toggle (OFF reverts to pre-transpose), support ruler-unit input (mm/pt/px, converted to points internally), unify apply functions, rename/cleanup.
- 2025-10-31: v1.0 Always-on preview, link (dim V), arrow-key increment.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RegridObjects";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 強制グリッド（ダイアログから切替）/ Force Grid mode (toggled from dialog) */
$.global.__regridForceGrid = $.global.__regridForceGrid || false;

/* 中央揃え：各セルの天地左右中央に整列（強制グリッドのサブオプション）/ Center each object in its cell (sub-option of Force Grid) */
$.global.__regridCenterInCell = $.global.__regridCenterInCell || false;

// =========================================
// プレビュー履歴ユーティリティ / Preview history util
// =========================================

/* =========================================
 * PreviewHistory util (extractable)
 * ヒストリーを残さないプレビューのための小さなユーティリティ。
 * 他スクリプトでもこのブロックをコピペすれば再利用できます。
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
// ローカライズ / Localization
// =========================================

/**
 * 現在のロケールから表示言語（"ja" / "en"）を判定します。
 *
 * @returns {string} "ja" または "en"。
 */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* ラベル定義（カテゴリ分け）/ Label definitions (categorized) */
var LABELS = {
    dialog: {
        title: { ja: "グリッドの間隔を再定義", en: "Redefine Grid Spacing" }
    },
    panel: {
        spacing: { ja: "間隔", en: "Spacing" },
        options: { ja: "オプション", en: "Options" }
    },
    field: {
        horizontal: { ja: "左右:", en: "H:" },
        vertical: { ja: "上下:", en: "V:" }
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
        cancel: { ja: "キャンセル", en: "Cancel" }
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
 * ラベルを現在の言語で取得します。ドット区切りで LABELS のカテゴリを辿ります。
 * "{slash}" プレースホルダは "/" に置換します。
 *
 * @param {string} path - ラベルのパス（例: "dialog.title", "checkbox.brick"）。
 * @returns {string} 現在の言語のラベル文字列。見つからない場合は path をそのまま返す。
 */
function L(path) {
    var node = LABELS;
    var parts = path.split('.');
    for (var i = 0; i < parts.length; i++) {
        if (node == null) break;
        node = node[parts[i]];
    }
    if (node == null) return path;
    var text = node[currentLanguage] || node.ja || node.en || path;
    return String(text).replace(/\{slash\}/g, '/');
}

// =========================================
// 単位 / Units
// =========================================

/* 単位コードとラベルのマップ / Unit code → label map */
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
 * 現在の定規単位に対応するラベル文字列を取得します。
 *
 * @returns {string} 単位ラベル（例: "mm", "pt", "px"）。未知の単位コードは "pt"。
 */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

/* 単位コード → 1単位あたりの pt 数 / Unit code → points per unit */
var unitToPointFactor = {
    0: 72.0,                  // in
    1: 72.0 / 25.4,           // mm
    2: 1.0,                   // pt
    3: 12.0,                  // pica
    4: 72.0 / 2.54,           // cm
    5: 72.0 / 25.4 * 0.25,    // Q/H（1Q = 0.25mm）
    6: 1.0,                   // px（72dpi）
    7: 72.0,                  // ft/in（inch 基準で近似）
    8: 72.0 / 25.4 * 1000.0,  // m
    9: 72.0 * 36.0,           // yd
    10: 72.0 * 12.0           // ft
};

/**
 * 現在の定規単位における「1単位 = 何 pt か」を返します。
 * geometricBounds / translate() は pt 基準なので、表示単位↔pt の換算に使います。
 *
 * @returns {number} 1単位あたりの pt 数（未知単位や不正値は 1.0）。
 */
function getUnitToPointFactor() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var factor = unitToPointFactor[unitCode];
    return (typeof factor === "number" && factor > 0) ? factor : 1.0;
}

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/**
 * ウィンドウに共通のレイアウト設定（縦並び・外周余白・要素間隔）を適用します。
 *
 * @param {Window} win - 対象のダイアログウィンドウ。
 * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING を使用。
 * @returns {void}
 */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルに共通のレイアウト設定（縦並び・横いっぱい・余白・要素間隔）を適用します。
 *
 * @param {Panel} panel - 対象のパネル。
 * @param {number} [spacing] - パネル内の要素間隔。省略時は PANEL_SPACING を使用。
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
 * 行グループ（ボタン列など）に共通の横並び設定を適用します。
 *
 * @param {Group} group - 対象のグループ。
 * @param {string} [alignment] - グループの配置（"left" / "center" / "right" など）。省略時は "left"。
 * @param {number} [spacing] - グループ内の要素間隔。省略時は PANEL_SPACING を使用。
 * @returns {void}
 */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/*
EditTextで↑↓キーによる値の増減を可能にする
Enable arrow-key increment on EditText
- ↑ / ↓ : ±1
- Shift + ↑ / ↓ : ±10 (snap to 10)
- Option(Alt) + ↑ / ↓ : ±0.1
*/
function changeValueByArrowKey(editText) {
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
        if (typeof $.global.updatePreview === "function") {
            $.global.updatePreview();
        }
    });
}

/**
 * 間隔設定ダイアログを生成・表示し、プレビューと確定適用を行います。
 *
 * @param {Function} applySpacing - 通常グリッドで間隔を適用する関数 (gapX, gapY) => void。
 * @param {Function} applySpacingBrick - レンガ状で間隔を適用する関数 (gapX, gapY) => void。
 * @param {Function} applySpacingHexagon - ハニカム状で間隔を適用する関数 (gapX, gapY) => void。
 * @param {Function} runTransposeIfNeeded - 行列入れ替えを実行する関数 () => void。
 * @param {Function} restoreInitialPositions - ダイアログ開始時点の位置へ戻す関数 () => void。
 * @param {Function} resetBaselineToCurrent - 現在位置を新しい基準として再設定する関数 () => void。
 * @param {number} initialGapX - 左右間隔の初期値。
 * @returns {void}
 *
 * 注: getCurrentUnitLabel / getUnitToPointFactor / changeValueByArrowKey は
 *     トップレベル関数のため引数で受け取らず、内部で直接呼び出す。
 */
function showGridSpacingDialog(applySpacing, applySpacingBrick, applySpacingHexagon, runTransposeIfNeeded, restoreInitialPositions, resetBaselineToCurrent, initialGapX) {
    // タイトルとバージョンを合成 / combine title and version
    var dialogTitle = L('dialog.title') + ' ' + SCRIPT_VERSION;

    var dialog = new Window('dialog', dialogTitle);
    setupWindow(dialog);

    // プレビュー用ヒストリー管理を開始 / Start preview history counter
    PreviewHistory.start();

    // 行列入れ替えが実行済みか（プレビュー状態）/ whether transpose has been triggered (preview state)
    var didTranspose = false;

    // 表示単位↔pt の換算係数（入力値は表示単位、内部処理は pt）
    // Points per display unit (inputs are in display units; internal geometry is in pt)
    var unitFactor = getUnitToPointFactor();

    // パネル名に単位を出す / show unit in panel title
    var spacingPanel = dialog.add('panel', undefined, L('panel.spacing') + ' (' + getCurrentUnitLabel() + ')');
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

    var horizontalGapGroup = gapInputColumn.add('group');
    horizontalGapGroup.add('statictext', undefined, L('field.horizontal'));
    // 初期値は pt を表示単位に換算して表示 / show initial value converted from pt to the display unit
    var horizontalGapInput = horizontalGapGroup.add('edittext', undefined, (initialGapX / unitFactor).toFixed(1));
    horizontalGapInput.characters = 4;
    changeValueByArrowKey(horizontalGapInput);

    var verticalGapGroup = gapInputColumn.add('group');
    verticalGapGroup.add('statictext', undefined, L('field.vertical'));
    // 初期表示では負の値を使わない
    var verticalGapInput = verticalGapGroup.add('edittext', undefined, '0');
    verticalGapInput.characters = 4;
    changeValueByArrowKey(verticalGapInput);

    // 右カラム（連動）/ right column (link)
    var linkColumn = spacingPanel.add('group');
    linkColumn.orientation = 'column';
    linkColumn.alignChildren = 'center';
    linkColumn.alignment = ['fill', 'fill'];
    var linkColumnSpacer = linkColumn.add('statictext', undefined, '');
    linkColumnSpacer.alignment = ['fill', 'fill'];
    var linkCheckbox = linkColumn.add('checkbox', undefined, L('checkbox.link'));

    // オプション（チェックボックスをまとめる）/ Options panel
    var optionsPanel = dialog.add('panel', undefined, L('panel.options'));
    setupPanel(optionsPanel, 6);

    // レンガ状 / Brick
    var brickGroup = optionsPanel.add('group');
    setupRow(brickGroup, 'left');
    var brickCheckbox = brickGroup.add('checkbox', undefined, L('checkbox.brick'));
    brickCheckbox.value = false;

    // ハニカム（レンガ状のサブオプション）/ Honeycomb (sub-option of Brick)
    var honeycombGroup = optionsPanel.add('group');
    setupRow(honeycombGroup, 'left');
    honeycombGroup.margins = [15, 0, 0, 0];
    var honeycombCheckbox = honeycombGroup.add('checkbox', undefined, L('checkbox.honeycomb'));
    honeycombCheckbox.value = false;
    honeycombCheckbox.enabled = false;

    // 強制グリッド / Force Grid
    var forceGridGroup = optionsPanel.add('group');
    setupRow(forceGridGroup, 'left');
    var forceGridCheckbox = forceGridGroup.add('checkbox', undefined, L('checkbox.forceGrid'));
    forceGridCheckbox.value = !!$.global.__regridForceGrid;

    // 中央揃え（強制グリッドのサブオプション）/ Center in cell (sub-option of Force Grid)
    var centerInCellGroup = optionsPanel.add('group');
    setupRow(centerInCellGroup, 'left');
    centerInCellGroup.margins = [15, 0, 0, 0];
    var centerInCellCheckbox = centerInCellGroup.add('checkbox', undefined, L('checkbox.centerInCell'));
    centerInCellCheckbox.value = !!$.global.__regridCenterInCell;
    centerInCellCheckbox.enabled = forceGridCheckbox.value;

    // 行列入れ替え / Swap rows/columns
    var transposeGroup = optionsPanel.add('group');
    setupRow(transposeGroup, 'left');
    var transposeCheckbox = transposeGroup.add('checkbox', undefined, L('checkbox.transpose'));
    transposeCheckbox.value = false;

    // 初期状態 / initial state
    linkCheckbox.value = true;
    horizontalGapInput.active = true;
    verticalGapInput.enabled = false;
    verticalGapInput.text = horizontalGapInput.text;

    honeycombCheckbox.enabled = brickCheckbox.value;

    // ボタン行（パネル外・中央寄せ、いっぱいに広げない）/ buttons (outside panels, centered, not stretched)
    var buttonGroup = dialog.add('group');
    setupRow(buttonGroup, 'center');
    var cancelButton = buttonGroup.add('button', undefined, L('button.cancel'), { name: 'cancel' });
    var okButton = buttonGroup.add('button', undefined, 'OK', { name: 'ok' });

    /**
     * 現在のオプション（レンガ／ハニカム／通常）に応じて間隔適用関数を選んで実行する。
     *
     * @param {number} gapX - 左右間隔。
     * @param {number} gapY - 上下間隔。
     * @returns {void}
     */
    function applySelectedSpacing(gapX, gapY) {
        if (brickCheckbox.value) {
            if (honeycombCheckbox.value) applySpacingHexagon(gapX, gapY);
            else applySpacingBrick(gapX, gapY);
        } else {
            applySpacing(gapX, gapY);
        }
    }

    /**
     * 転置の有無を考慮して間隔を適用する。
     * 転置ONのときはマージン0で並べ替え→転置→基準を取り直してから間隔を適用する。
     *
     * @param {number} gapX - 左右間隔。
     * @param {number} gapY - 上下間隔。
     * @param {boolean} bumpHistory - プレビュー時は各ステップで PreviewHistory.bump() する（確定時は false）。
     * @returns {void}
     */
    function applyLayoutWithTranspose(gapX, gapY, bumpHistory) {
        if (didTranspose) {
            // 転置はマージン0で実行し、その後マージンを適用 / Transpose with 0 margins then apply margins
            applySpacing(0, 0);
            if (bumpHistory) PreviewHistory.bump();

            runTransposeIfNeeded();
            if (bumpHistory) PreviewHistory.bump();

            // 転置後の配置を新しい基準に / adopt transposed layout as baseline for spacing
            resetBaselineToCurrent();
            applySelectedSpacing(gapX, gapY);
            if (bumpHistory) PreviewHistory.bump();
        } else {
            applySelectedSpacing(gapX, gapY);
            if (bumpHistory) PreviewHistory.bump();
        }
    }

    /**
     * 現在のUI状態（間隔・各オプション）に基づいてプレビューを再描画します。
     * 直前のプレビューを一括Undoしてから再適用するため、ヒストリーを汚しません。
     *
     * @returns {void}
     */
    function updatePreviewImpl() {
        // 直前のプレビューを一括Undo（ヒストリーを汚さない）/ Undo previous preview
        PreviewHistory.undo();
        // 強制グリッド／中央揃えチェックボックスの状態をグローバルに反映
        $.global.__regridForceGrid = !!forceGridCheckbox.value;
        $.global.__regridCenterInCell = !!centerInCellCheckbox.value;
        // Undo後の現在位置を基準として originalPositions/layoutInfo を作り直す
        resetBaselineToCurrent();

        // 通常時 / normal

        if (linkCheckbox.value) {
            verticalGapInput.enabled = false;
            verticalGapInput.text = horizontalGapInput.text;
        } else {
            verticalGapInput.enabled = true;
        }

        // 値は連動OFF時に UI 側（transposeCheckboxのonClick）で入れ替え済みのため、ここでは追加の入れ替えをしない
        // H/V are already swapped in the UI (transpose onClick) when Link is OFF, so do NOT swap again here
        var gapX = parseFloat(horizontalGapInput.text);
        var gapY = parseFloat(verticalGapInput.text);
        if (isNaN(gapX)) gapX = 0;
        if (isNaN(gapY)) gapY = 0;

        // 表示単位 → pt に換算して適用 / convert display units to pt before applying
        applyLayoutWithTranspose(gapX * unitFactor, gapY * unitFactor, true);
    }

    // changeValueByArrowKey() から呼べるようにグローバルへ / expose for arrow-key handler
    $.global.updatePreview = updatePreviewImpl;

    // イベント / events
    horizontalGapInput.onChanging = function () { updatePreviewImpl(); };
    verticalGapInput.onChanging = function () {
        if (!linkCheckbox.value) {
            updatePreviewImpl();
        }
    };
    linkCheckbox.onClick = function () { updatePreviewImpl(); };
    forceGridCheckbox.onClick = function () {
        // 強制グリッドOFF時は中央揃えも無効化 / disable center when Force Grid is off
        centerInCellCheckbox.enabled = forceGridCheckbox.value;
        if (!forceGridCheckbox.value) centerInCellCheckbox.value = false;
        updatePreviewImpl();
    };
    centerInCellCheckbox.onClick = function () { updatePreviewImpl(); };
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

        updatePreviewImpl();
    };

    // レンガ状 / Brick
    brickCheckbox.onClick = function () {
        honeycombCheckbox.enabled = brickCheckbox.value;
        if (!brickCheckbox.value) honeycombCheckbox.value = false;
        updatePreviewImpl();
    };

    // 六角形 / Hexagon
    honeycombCheckbox.onClick = function () {
        updatePreviewImpl();
    };

    // 開いたときに一度プレビュー / first preview when opened
    updatePreviewImpl();

    var dialogResult = dialog.show();

    if (dialogResult == 1) {
        // OK時は最終値で適用 / apply with final values
        var finalGapX = parseFloat(horizontalGapInput.text);
        var finalGapY = parseFloat(verticalGapInput.text);
        if (isNaN(finalGapX)) finalGapX = 0;
        if (isNaN(finalGapY)) finalGapY = 0;
        if (linkCheckbox.value) finalGapY = finalGapX;

        // 値は連動OFF時に UI 側で入れ替え済みのため、確定時も追加の入れ替えはしない
        // H/V are already swapped in the UI when Link is OFF, so no extra swap on final apply

        // プレビュー分を一括Undoしてから確定適用 / Clear preview history before final apply
        PreviewHistory.undo();
        resetBaselineToCurrent();
        // 表示単位 → pt に換算して適用 / convert display units to pt before applying
        applyLayoutWithTranspose(finalGapX * unitFactor, finalGapY * unitFactor, false);
    } else {
        // キャンセル時：プレビュー分を一括Undoしてから初期状態へ / Undo preview then restore
        PreviewHistory.undo();
        restoreInitialPositions();
    }

    // 念のためカウンタ初期化 / reset counter
    PreviewHistory.start();

    // グローバルに公開したプレビュー関数を破棄（古いクロージャの残留・別スクリプトとの衝突を防ぐ）
    // Drop the exposed preview closure so a stale dialog does not linger / collide with other scripts
    try { $.global.updatePreview = null; } catch (e) { }

    // (removed global focus cleanup)
}

/**
 * メイン処理。選択オブジェクトを収集し、間隔設定ダイアログを起動します。
 *
 * @returns {void}
 */
function main() {
    // ドキュメントチェック / document check
    if (app.documents.length === 0) {
        alert(L('alert.noDocument'));
        return;
    }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert(L('alert.noSelection'));
        return;
    }

    // 選択を拾う / collect selection
    var selectedItems = [];
    for (var i = 0; i < doc.selection.length; i++) {
        var selectionItem = doc.selection[i];
        selectedItems.push(selectionItem);
    }
    if (selectedItems.length < 2) {
        alert(L('alert.needTwo'));
        return;
    }

    // Sort selected items by top (Y) descending and left (X) ascending within same row
    var sortedByPosition = selectedItems.slice().sort(function (a, b) {
        var boundsA = getLayoutBounds(a);
        var boundsB = getLayoutBounds(b);
        if (Math.abs(boundsA[1] - boundsB[1]) < 1) {
            return boundsA[0] - boundsB[0]; // same row → compare left
        }
        return boundsB[1] - boundsA[1]; // sort by top descending
    });

    var firstBounds = getLayoutBounds(sortedByPosition[0]);
    var secondBounds = getLayoutBounds(sortedByPosition[1]);
    // ダイアログ初期値では負の値を使わない
    var initialGapX = secondBounds[0] - firstBounds[2];
    if (initialGapX < 0) initialGapX = 0;

    // 元位置を保存 / save original positions
    var originalPositions = [];
    for (var k = 0; k < selectedItems.length; k++) {
        var itemBounds = getLayoutBounds(selectedItems[k]);
        originalPositions.push({
            item: selectedItems[k],
            left: itemBounds[0],
            top: itemBounds[1]
        });
    }

    // ダイアログ開始時点の位置を保持（キャンセルで必ずここへ戻す）/ Snapshot for Cancel
    var initialPositions = [];
    for (var k0 = 0; k0 < originalPositions.length; k0++) {
        initialPositions.push({
            item: originalPositions[k0].item,
            left: originalPositions[k0].left,
            top: originalPositions[k0].top
        });
    }

    // 任意のスナップショットへ復元 / Restore from a given snapshot
    function restoreFrom(snapshot) {
        for (var i = 0; i < snapshot.length; i++) {
            var snapshotEntry = snapshot[i];
            var currentBounds = getLayoutBounds(snapshotEntry.item);
            var currentLeft = currentBounds[0];
            var currentTop = currentBounds[1];
            snapshotEntry.item.translate(snapshotEntry.left - currentLeft, snapshotEntry.top - currentTop);
        }
    }

    /*
レイアウト計算用の境界を取得 / Get bounds for layout calculation
- すでにグループになっているものは、中身を分解せず「グループ＝1つのオブジェクト」として扱う
- クリップグループの場合はクリップパスの geometricBounds を優先（=可視領域）
- それ以外は item.geometricBounds（GroupItem なら子全体の外接 bbox）
*/
    function getLayoutBounds(item) {
        try {
            if (item.typename === 'GroupItem' && item.clipped) {
                // GroupItem の中から clipping パスを探す
                if (item.pathItems && item.pathItems.length > 0) {
                    for (var i = 0; i < item.pathItems.length; i++) {
                        if (item.pathItems[i].clipping) return item.pathItems[i].geometricBounds;
                    }
                }
                // CompoundPath が clipping のケース
                if (item.compoundPathItems && item.compoundPathItems.length > 0) {
                    for (var j = 0; j < item.compoundPathItems.length; j++) {
                        var compoundPath = item.compoundPathItems[j];
                        if (compoundPath.pathItems && compoundPath.pathItems.length > 0 && compoundPath.pathItems[0].clipping) {
                            return compoundPath.pathItems[0].geometricBounds;
                        }
                    }
                }
                // fallback
                return item.geometricBounds;
            }
        } catch (e) { }

        return item.geometricBounds;
    }

    /**
     * 選択オブジェクトの外接境界一覧と、最小の幅・高さを収集する。
     * buildLayoutInfo / buildLayoutInfoForceGrid の共通前処理。
     *
     * @returns {{boundsList: Array, minWidth: number, minHeight: number}}
     *   boundsList: {item, bounds} の配列。minWidth/minHeight: 選択中の最小幅・最小高さ。
     */
    function collectBoundsList() {
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

    /*
    行・列の推定 / build layout info
    */
    function buildLayoutInfo() {
        var collected = collectBoundsList();
        var boundsList = collected.boundsList;
        var minWidth = collected.minWidth;
        var minHeight = collected.minHeight;

        // 列まとめ / group columns
        var colCenters = [];
        var colTolerance = minWidth * 0.5;
        if (colTolerance < 1) colTolerance = 1;

        boundsList.sort(function (a, b) { return a.bounds[0] - b.bounds[0]; });

        for (var c = 0; c < boundsList.length; c++) {
            var boundsRec = boundsList[c];
            var merged = false;
            var itemLeft = boundsRec.bounds[0];
            for (var cc = 0; cc < colCenters.length; cc++) {
                if (Math.abs(colCenters[cc].x - itemLeft) <= colTolerance) {
                    colCenters[cc].members.push(boundsRec);
                    colCenters[cc].x = (colCenters[cc].x * (colCenters[cc].members.length - 1) + itemLeft) / colCenters[cc].members.length;
                    merged = true;
                    break;
                }
            }
            if (!merged) {
                colCenters.push({ x: itemLeft, members: [boundsRec] });
            }
        }

        // 行まとめ / group rows
        var rowCenters = [];
        var rowTolerance = minHeight * 0.5;
        if (rowTolerance < 1) rowTolerance = 1;

        var boundsSortedByTop = boundsList.slice().sort(function (a, b) { return b.bounds[1] - a.bounds[1]; });

        for (var r = 0; r < boundsSortedByTop.length; r++) {
            var rowBoundsRec = boundsSortedByTop[r];
            var merged2 = false;
            var itemTop = rowBoundsRec.bounds[1];
            for (var rr = 0; rr < rowCenters.length; rr++) {
                if (Math.abs(rowCenters[rr].y - itemTop) <= rowTolerance) {
                    rowCenters[rr].members.push(rowBoundsRec);
                    rowCenters[rr].y = (rowCenters[rr].y * (rowCenters[rr].members.length - 1) + itemTop) / rowCenters[rr].members.length;
                    merged2 = true;
                    break;
                }
            }
            if (!merged2) {
                rowCenters.push({ y: itemTop, members: [rowBoundsRec] });
            }
        }

        // 並び順の確定 / sort
        colCenters.sort(function (a, b) { return a.x - b.x; });
        rowCenters.sort(function (a, b) { return b.y - a.y; });

        // 各列の最大幅 / max width per column
        var colWidths = [];
        for (var ci = 0; ci < colCenters.length; ci++) {
            var maxWidth = 0;
            for (var ci2 = 0; ci2 < colCenters[ci].members.length; ci2++) {
                var memberBounds = colCenters[ci].members[ci2].bounds;
                var memberWidth = memberBounds[2] - memberBounds[0];
                if (memberWidth > maxWidth) maxWidth = memberWidth;
            }
            colWidths.push(maxWidth);
        }

        // 各行の最大高さ / max height per row
        var rowHeights = [];
        for (var ri = 0; ri < rowCenters.length; ri++) {
            var maxHeight = 0;
            for (var ri2 = 0; ri2 < rowCenters[ri].members.length; ri2++) {
                var memberBounds = rowCenters[ri].members[ri2].bounds;
                var memberHeight = memberBounds[1] - memberBounds[3];
                if (memberHeight > maxHeight) maxHeight = memberHeight;
            }
            rowHeights.push(maxHeight);
        }

        var baseX = colCenters[0].x;
        var baseY = rowCenters[0].y;

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

    var layoutInfo = ($.global.__regridForceGrid) ? buildLayoutInfoForceGrid() : buildLayoutInfo();

    /*
    強制グリッド：行ごと（上→下）に左→右で(行,列)を割り当て / Force Grid: row-major assignment
    - 行はtop(Y)の近さでクラスタリング
    - 各行の中はleft(X)でソート
    - 欠け（歯抜け）は許容（行ごとに列数が異なってよい）
    */
    function buildLayoutInfoForceGrid() {
        var collected = collectBoundsList();
        var boundsList = collected.boundsList;
        var minHeight = collected.minHeight;

        // 行まとめ（上→下）/ group rows (top -> bottom)
        var rows = [];
        var rowTolerance = minHeight * 0.5;
        if (rowTolerance < 1) rowTolerance = 1;

        var boundsSortedByTop = boundsList.slice().sort(function (a, b) { return b.bounds[1] - a.bounds[1]; });
        for (var i = 0; i < boundsSortedByTop.length; i++) {
            var boundsRec = boundsSortedByTop[i];
            var y = boundsRec.bounds[1];
            var placed = false;
            for (var r = 0; r < rows.length; r++) {
                if (Math.abs(rows[r].y - y) <= rowTolerance) {
                    rows[r].members.push(boundsRec);
                    rows[r].y = (rows[r].y * (rows[r].members.length - 1) + y) / rows[r].members.length;
                    placed = true;
                    break;
                }
            }
            if (!placed) rows.push({ y: y, members: [boundsRec] });
        }
        rows.sort(function (a, b) { return b.y - a.y; });

        // 各行の中を左→右で確定し、rowIndex/colIndexを付与 / sort within row and assign indices
        var maxCols = 0;
        for (var r2 = 0; r2 < rows.length; r2++) {
            rows[r2].members.sort(function (a, b) { return a.bounds[0] - b.bounds[0]; });
            if (rows[r2].members.length > maxCols) maxCols = rows[r2].members.length;
            for (var c2 = 0; c2 < rows[r2].members.length; c2++) {
                rows[r2].members[c2].rowIndex = r2;
                rows[r2].members[c2].colIndex = c2;
            }
        }

        // colWidths（列ごとの最大幅）/ max width per column index
        var colWidths = [];
        for (var c3 = 0; c3 < maxCols; c3++) {
            var maxWidth = 0;
            for (var r3 = 0; r3 < rows.length; r3++) {
                if (rows[r3].members.length > c3) {
                    var memberBounds = rows[r3].members[c3].bounds;
                    var memberWidth = memberBounds[2] - memberBounds[0];
                    if (memberWidth > maxWidth) maxWidth = memberWidth;
                }
            }
            colWidths.push(maxWidth);
        }

        // rowHeights（行ごとの最大高さ）/ max height per row
        var rowHeights = [];
        for (var r4 = 0; r4 < rows.length; r4++) {
            var maxHeight = 0;
            for (var k = 0; k < rows[r4].members.length; k++) {
                var memberBounds = rows[r4].members[k].bounds;
                var memberHeight = memberBounds[1] - memberBounds[3];
                if (memberHeight > maxHeight) maxHeight = memberHeight;
            }
            rowHeights.push(maxHeight);
        }

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

    // 現在位置を新しい基準（originalPositions + layoutInfo）として再設定 / Reset baseline to current layout
    function resetBaselineToCurrent() {
        originalPositions = [];
        for (var kB = 0; kB < selectedItems.length; kB++) {
            var currentItemBounds = getLayoutBounds(selectedItems[kB]);
            originalPositions.push({ item: selectedItems[kB], left: currentItemBounds[0], top: currentItemBounds[1] });
        }
        layoutInfo = ($.global.__regridForceGrid) ? buildLayoutInfoForceGrid() : buildLayoutInfo();
    }

    /*
    行列入れ替え（歯抜け対応）/ Transpose rows & columns (tolerate missing cells)
    - getLayoutBounds() の left/top を基準に、行・列をクラスタリングして推定
    - グループは中身を見ず、グループ全体の外接 bbox（クリップグループはクリップパス）で1オブジェクトとして扱う
    - 推定したピッチ（隣接差の中央値）で左上基準に再配置
    - 1行→1列 / 1列→1行 も対応（ピッチ流用）
    */
    function transposeGridWithHoles() {
        if (!selectedItems || selectedItems.length < 1) return;

        // ---- 調整パラメータ / tuning params
        var SNAP_X_TOL = 8.0;   // X方向の「同じ列」とみなす許容（pt）
        var SNAP_Y_TOL = 8.0;   // Y方向の「同じ行」とみなす許容（pt）

        // getLayoutBounds: [left, top, right, bottom]（グループは1つのbboxとして扱う）
        function leftX(targetItem) { return getLayoutBounds(targetItem)[0]; }
        function topY(targetItem) { return getLayoutBounds(targetItem)[1]; }

        // 近い値をクラスタリングして中心値配列を作る / cluster near values into centers
        function clusterValues(values, tolerance) {
            values.sort(function (a, b) { return a - b; });
            var centers = [];
            for (var i = 0; i < values.length; i++) {
                var value = values[i];
                var foundIndex = -1;
                for (var c = 0; c < centers.length; c++) {
                    if (Math.abs(value - centers[c]) <= tolerance) { foundIndex = c; break; }
                }
                if (foundIndex < 0) centers.push(value);
                else centers[foundIndex] = (centers[foundIndex] + value) / 2.0;
            }
            centers.sort(function (a, b) { return a - b; });
            return centers;
        }

        function nearestIndex(sortedCenters, value) {
            var bestIndex = 0;
            var bestDistance = Math.abs(value - sortedCenters[0]);
            for (var i = 1; i < sortedCenters.length; i++) {
                var distance = Math.abs(value - sortedCenters[i]);
                if (distance < bestDistance) { bestDistance = distance; bestIndex = i; }
            }
            return bestIndex;
        }

        // left/top を集める / collect left/top
        var leftValues = [], topValues = [];
        for (var k = 0; k < selectedItems.length; k++) {
            leftValues.push(leftX(selectedItems[k]));
            topValues.push(topY(selectedItems[k]));
        }

        var colClusters = clusterValues(leftValues, SNAP_X_TOL); // left -> right
        var rowClusters = clusterValues(topValues, SNAP_Y_TOL); // will sort top -> bottom next

        // Illustrator座標では上ほどYが大きいことが多いので「上→下」/ sort top -> bottom
        rowClusters.sort(function (a, b) { return b - a; });

        var rowCount = rowClusters.length;
        var colCount = colClusters.length;

        // 各オブジェクトを(行,列)に割り当て / assign each item to (row, col)
        var occupancy = {}; // cellKey "row,col" -> item
        var mapping = [];   // {item, row, col}
        for (var m = 0; m < selectedItems.length; m++) {
            var targetItem = selectedItems[m];
            var rowIndex = nearestIndex(rowClusters, topY(targetItem));
            var colIndex = nearestIndex(colClusters, leftX(targetItem));
            var cellKey = rowIndex + "," + colIndex;
            if (occupancy[cellKey]) {
                alert(L('alert.cellConflict') + "(" + rowIndex + "," + colIndex + ")");
                return;
            }
            occupancy[cellKey] = targetItem;
            mapping.push({ item: targetItem, row: rowIndex, col: colIndex });
        }

        // 元の列/行のピッチを推定（隣接差の中央値）
        // ※行や列が1つしかない場合は0になる
        var pitchX = (colClusters.length >= 2) ? medianAdjacentDiff(colClusters) : 0;
        var pitchY = (rowClusters.length >= 2) ? medianAdjacentDiff(rowClusters) : 0;

        // 1行→1列、1列→1行にも対応
        // - 1行しかない場合: 横方向ピッチ(pitchX)を縦方向の並び間隔として流用
        // - 1列しかない場合: 縦方向ピッチ(pitchY)を横方向の並び間隔として流用
        var effectivePitchX = pitchX;
        var effectivePitchY = pitchY;

        if (rowCount === 1 && colCount === 1) {
            // ほぼ重なり等で1セル扱いになったケース
            return;
        }

        if (rowCount === 1 && colCount > 1) {
            // 1行 → 1列
            if (pitchX === 0) return;
            effectivePitchX = pitchX;
            effectivePitchY = pitchX; // 横ピッチを縦へ流用
        } else if (colCount === 1 && rowCount > 1) {
            // 1列 → 1行
            if (pitchY === 0) return;
            effectivePitchX = pitchY; // 縦ピッチを横へ流用
            effectivePitchY = pitchY;
        } else {
            // 通常（2行以上 かつ 2列以上）
            if (pitchX === 0 || pitchY === 0) return;
        }

        // 転置後グリッドの基準（左上固定）/ origin at top-left
        var originLeft = colClusters[0];
        var originTop = rowClusters[0];

        // 転置: 新しい列 = 元の行、新しい行 = 元の列 / newCol = oldRow, newRow = oldCol
        for (var t = 0; t < mapping.length; t++) {
            var mappedItem = mapping[t].item;
            var oldRow = mapping[t].row;
            var oldCol = mapping[t].col;

            var newCol = oldRow;
            var newRow = oldCol;

            var targetLeft = originLeft + newCol * effectivePitchX;
            var targetTop = originTop - newRow * effectivePitchY;

            var itemBounds = getLayoutBounds(mappedItem);
            var currentLeft = itemBounds[0];
            var currentTop = itemBounds[1];

            mappedItem.translate(targetLeft - currentLeft, targetTop - currentTop);
        }

        app.redraw();
    }

    /* 元位置に戻す（プレビュー基準）/ restore baseline positions for preview */
    function restoreOriginalPositions() {
        restoreFrom(originalPositions);
    }

    /* キャンセル時に戻す（ダイアログ開始時点）/ restore initial positions on Cancel */
    function restoreInitialPositions() {
        restoreFrom(initialPositions);
    }

    // 数値配列の隣接差分の中央値 / Median of adjacent diffs (ascending)
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
     * boundsList の要素の元座標（左上）を originalPositions から引く。
     * 見つからなければ現在の bbox（currentEntry.bounds）を使う。
     *
     * @param {object} currentEntry - layoutInfo.boundsList の要素（{item, bounds}）。
     * @returns {{left:number, top:number}} 元の左端X・上端Y。
     */
    function findOriginalLeftTop(currentEntry) {
        for (var oo = 0; oo < originalPositions.length; oo++) {
            if (originalPositions[oo].item === currentEntry.item) {
                return { left: originalPositions[oo].left, top: originalPositions[oo].top };
            }
        }
        return { left: currentEntry.bounds[0], top: currentEntry.bounds[1] };
    }

    /**
     * 対象の列・行インデックスを解決する。
     * Force Grid で事前割り当て済み（colIndex/rowIndex）ならそれを優先し、
     * なければ列／行センターへの最近傍で推定する。
     *
     * @param {object} currentEntry - layoutInfo.boundsList の要素。
     * @param {Array} colCenters - 列センター配列。
     * @param {Array} rowCenters - 行センター配列。
     * @returns {{col:number, row:number}} 列・行インデックス。
     */
    function resolveColRow(currentEntry, colCenters, rowCenters) {
        var colIndex = (typeof currentEntry.colIndex === 'number') ? currentEntry.colIndex : null;
        var rowIndex = (typeof currentEntry.rowIndex === 'number') ? currentEntry.rowIndex : null;

        if (colIndex === null) {
            colIndex = 0;
            var minDX = Number.MAX_VALUE;
            for (var c2 = 0; c2 < colCenters.length; c2++) {
                var dx = Math.abs(colCenters[c2].x - currentEntry.bounds[0]);
                if (dx < minDX) { minDX = dx; colIndex = c2; }
            }
        }

        if (rowIndex === null) {
            rowIndex = 0;
            var minDY = Number.MAX_VALUE;
            for (var r2 = 0; r2 < rowCenters.length; r2++) {
                var dy = Math.abs(rowCenters[r2].y - currentEntry.bounds[1]);
                if (dy < minDY) { minDY = dy; rowIndex = r2; }
            }
        }

        return { col: colIndex, row: rowIndex };
    }

    /**
     * レンガ／ハニカムで使う半ピッチを算出する。
     * 目標列ピッチ＝中央値幅 + gapX。推定できない場合は列センター間隔の中央値を使う。
     *
     * @param {Array} colWidths - 列ごとの最大幅。
     * @param {Array} colCenters - 列センター配列。
     * @param {number} gapX - 左右間隔。
     * @returns {number} 半ピッチ（pitch / 2）。
     */
    function computeHalfPitch(colWidths, colCenters, gapX) {
        var sortedColWidths = [];
        if (colWidths && colWidths.length > 0) {
            for (var wi = 0; wi < colWidths.length; wi++) sortedColWidths.push(colWidths[wi]);
            sortedColWidths.sort(function (a, b) { return a - b; });
        }
        var medianColWidth = (sortedColWidths.length > 0) ? sortedColWidths[Math.floor(sortedColWidths.length / 2)] : 0; // median width
        var pitch = (medianColWidth > 0) ? (medianColWidth + gapX) : 0;

        // fallback：列が1つ等で推定できない場合は現状の列位置差 / fallback to current centers diff
        if (pitch === 0) {
            var colLefts = [];
            for (var iC = 0; iC < colCenters.length; iC++) colLefts.push(colCenters[iC].x);
            colLefts.sort(function (a, b) { return a - b; });
            pitch = medianAdjacentDiff(colLefts);
        }
        return pitch / 2.0;
    }

    /**
     * グリッド配置を適用する共通処理。通常／レンガ／ハニカムを引数で切り替える。
     * 直前にプレビュー基準位置へ戻してから、列幅・行高さの累積で再配置する。
     *
     * @param {number} gapX - 左右間隔。
     * @param {number} gapY - 上下間隔。
     * @param {boolean} isBrick - 奇数行を半ピッチ横にずらすか（レンガ／ハニカム）。
     * @param {number} rowStepFactor - 行送りに掛ける係数（通常・レンガ=1.0、ハニカム=0.75）。
     * @returns {void}
     */
    function applyGridLayout(gapX, gapY, isBrick, rowStepFactor) {
        // いったん元に戻す / restore first
        restoreOriginalPositions();

        var boundsList = layoutInfo.boundsList;
        var colCenters = layoutInfo.colCenters;
        var rowCenters = layoutInfo.rowCenters;
        var colWidths = layoutInfo.colWidths;
        var rowHeights = layoutInfo.rowHeights;
        var baseX = layoutInfo.baseX;
        var baseY = layoutInfo.baseY;

        var halfPitch = isBrick ? computeHalfPitch(colWidths, colCenters, gapX) : 0;

        // 中央揃え：各セルの天地左右中央に整列（強制グリッドのサブオプションなので Force Grid 時のみ有効）
        // Center each object within its cell (sub-option of Force Grid, so only when Force Grid is on)
        var centerInCell = !!$.global.__regridCenterInCell && !!$.global.__regridForceGrid;

        for (var k2 = 0; k2 < boundsList.length; k2++) {
            var currentEntry = boundsList[k2];

            var originalPos = findOriginalLeftTop(currentEntry);
            var cellIndex = resolveColRow(currentEntry, colCenters, rowCenters);
            var colIndex = cellIndex.col;
            var rowIndex = cellIndex.row;

            // 新しいX（セル左端）/ new X (cell left)
            var newX = baseX;
            for (var cc2 = 0; cc2 < colIndex; cc2++) {
                newX += colWidths[cc2] + gapX;
            }

            // 新しいY（セル上端）/ new Y (cell top)
            var newY = baseY;
            for (var rr2 = 0; rr2 < rowIndex; rr2++) {
                newY -= (rowHeights[rr2] * rowStepFactor) + gapY;
            }

            // 中央揃え：セル内でオブジェクトを左右・天地中央へ寄せる
            // Center within the cell (cell size = column width × row height)
            if (centerInCell) {
                var cellWidth = colWidths[colIndex];
                var cellHeight = rowHeights[rowIndex];
                var itemWidth = currentEntry.bounds[2] - currentEntry.bounds[0];
                var itemHeight = currentEntry.bounds[1] - currentEntry.bounds[3];
                newX += (cellWidth - itemWidth) / 2;
                newY -= (cellHeight - itemHeight) / 2;
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

    /* 間隔を適用（通常グリッド）/ apply spacing (normal grid) */
    function applySpacing(gapX, gapY) {
        applyGridLayout(gapX, gapY, false, 1.0);
    }

    /* 間隔を適用（レンガ状）/ apply spacing (brick layout) */
    function applySpacingBrick(gapX, gapY) {
        applyGridLayout(gapX, gapY, true, 1.0);
    }

    /* 間隔を適用（六角形/ハニカム）/ apply spacing (hexagon/honeycomb layout) */
    function applySpacingHexagon(gapX, gapY) {
        applyGridLayout(gapX, gapY, true, 0.75);
    }

    // ダイアログ表示 / show dialog
    showGridSpacingDialog(
        applySpacing,
        applySpacingBrick,
        applySpacingHexagon,
        transposeGridWithHoles,
        restoreInitialPositions,
        resetBaselineToCurrent,
        initialGapX
    );
}

// 実行 / run
main();

})();