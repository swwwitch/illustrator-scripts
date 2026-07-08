#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択中のオブジェクトが「だいたいグリッド状」に並んでいることを前提に、左右・上下の間隔で再配置します。
- 常時プレビュー対応。値は EditText で直接入力でき、↑↓ キー（Shift で ×10、Option で ×0.1）でも増減できます。
- すでにグループになっているもの（クリップグループ含む）は、中身を分解せず「1 つのオブジェクト＝1 つの外接 bbox」として扱います。
- 実行後に自動でグループ化はしません（選択状態のまま）。
- ダイアログは日本語／英語の自動切り替えに対応（$.locale）。
- ［連動］で左右の値を上下に反映できます。
- ［レンガ状］：行ごとに左右位置を半ピッチずらして再配置します。
- ［ハニカム］：［レンガ状］と併用し、奇数行を（幅＋左右の値）の半分だけ横にずらし、行の高さだけ 0.75 倍にしてハニカム状にします（上下の値はそのまま反映）。
- ［強制グリッド］：位置の近さで列／行を推定せず、上→下の行ごとに左→右の順で (行, 列) を割り当てて再配置します。
- ［行列入れ替え］を ON にすると、歯抜け（欠け）を許容しつつ行⇄列を転置します。
- 1 行だけ→1 列、1 列だけ→1 行の転置にも対応します。

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RegridObjects.jsx

### 紹介記事（note）

https://note.com/dtp_tranist/n/n08861d0e40c3

### 更新履歴

- 2026-05-06: v1.5.9 グループ／クリップグループは中身を分解せず 1 オブジェクトとして扱うよう統一（行列入れ替えも getLayoutBounds 基準に変更）
- 2025-10-31: v1.0 プレビューの常時 ON、連動（上下ディム）、↑↓ キーでの増減

*/

/*

### Overview

- Assumes the selected objects are arranged in an "approximately grid-like" layout.
- Always-on preview. Values can be entered directly in the EditText fields, and adjusted via ↑↓ keys (×10 with Shift, ×0.1 with Option).
- Existing groups (including clipped groups) are treated as a single object — one outer bounding box — and never broken apart.
- Does not auto-group results (keeps items selected).
- Repositions them using the horizontal/vertical gaps entered in the dialog.
- The dialog supports automatic Japanese/English switching (based on $.locale).
- The "Link" option mirrors the horizontal value to the vertical value.
- "Brick": shifts every other row by a half pitch.
- "Honeycomb": used with "Brick"; shifts odd rows by half of the target cell pitch (width + H gap) and compresses only row height (0.75×) while keeping the V gap as-is.
- "Force Grid": assigns (row, col) by row-major order instead of inferring nearest columns/rows by position.
- "Swap Rows/Columns": transposes rows/columns while tolerating missing cells (including 1 row ⇄ 1 column).

### Update history

- 2026-05-06: v1.5.9 Treat groups / clipped groups as a single object (transpose also uses getLayoutBounds).
- 2025-10-31: v1.0 Always-on preview, link (dim V), arrow-key increment.

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.5.9";

(function () {

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 強制グリッド（ダイアログから切替）/ Force Grid mode (toggled from dialog) */
$.global.__regridForceGrid = $.global.__regridForceGrid || false;

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
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            // 10単位で増減 / change by 10
            delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
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
        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10;
        } else {
            value = Math.round(value);
        }

        editText.text = value;

        // 値変更後にプレビュー / update preview after change
        if (typeof updatePreview === "function") {
            updatePreview();
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
 * @param {Function} restoreOriginalPositions - プレビュー基準位置へ戻す関数 () => void。
 * @param {Function} restoreInitialPositions - ダイアログ開始時点の位置へ戻す関数 () => void。
 * @param {Function} resetBaselineToCurrent - 現在位置を新しい基準として再設定する関数 () => void。
 * @param {Function} getCurrentUnitLabel - 現在の単位ラベルを返す関数 () => string。
 * @param {Function} changeValueByArrowKey - EditText に↑↓キー増減を付与する関数 (EditText) => void。
 * @param {number} initialGapX - 左右間隔の初期値。
 * @returns {void}
 */
function showGridSpacingDialog(applySpacing, applySpacingBrick, applySpacingHexagon, runTransposeIfNeeded, restoreOriginalPositions, restoreInitialPositions, resetBaselineToCurrent, getCurrentUnitLabel, changeValueByArrowKey, initialGapX) {
    // タイトルとバージョンを合成 / combine title and version
    var dialogTitle = L('dialog.title') + ' ' + SCRIPT_VERSION;

    var dialog = new Window('dialog', dialogTitle);
    setupWindow(dialog);

    // プレビュー用ヒストリー管理を開始 / Start preview history counter
    PreviewHistory.start();

    // 行列入れ替えが実行済みか（プレビュー状態）/ whether transpose has been triggered (preview state)
    var didTranspose = false;

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
    var horizontalGapInput = horizontalGapGroup.add('edittext', undefined, initialGapX.toFixed(1));
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

    // 行列入れ替え / Swap rows/columns
    var transposeGroup = optionsPanel.add('group');
    setupRow(transposeGroup, 'left');
    var transposeCheckbox = transposeGroup.add('checkbox', undefined, L('checkbox.transpose'));
    transposeCheckbox.value = false;

    // 転置ON時はマージンを0に固定（UIも0表示）/ When transposing, force margins to 0
    var savedHorizontalGapText = null;
    var savedVerticalGapText = null;



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
     * 現在のUI状態（間隔・各オプション）に基づいてプレビューを再描画します。
     * 直前のプレビューを一括Undoしてから再適用するため、ヒストリーを汚しません。
     *
     * @returns {void}
     */
    function updatePreviewImpl() {
        // 直前のプレビューを一括Undo（ヒストリーを汚さない）/ Undo previous preview
        PreviewHistory.undo();
        // 強制グリッドチェックボックスの状態をグローバルに反映
        $.global.__regridForceGrid = !!forceGridCheckbox.value;
        // Undo後の現在位置を基準として originalPositions/layoutInfo を作り直す
        resetBaselineToCurrent();

        if (transposeCheckbox.value) {
            // チェックON中は表示上0にする（実処理はクリック時に1回だけ実行）
            if (horizontalGapInput.text !== "0") horizontalGapInput.text = "0";
            if (verticalGapInput.text !== "0") verticalGapInput.text = "0";
        }

        // 通常時 / normal

        if (linkCheckbox.value) {
            verticalGapInput.enabled = false;
            verticalGapInput.text = horizontalGapInput.text;
        } else {
            verticalGapInput.enabled = true;
        }

        var gapX = parseFloat(horizontalGapInput.text);
        var gapY = parseFloat(verticalGapInput.text);
        if (isNaN(gapX)) gapX = 0;
        if (isNaN(gapY)) gapY = 0;

        // 転置時は左右/上下のマージンを入れ替える（連動OFFのときのみ）/ Swap H/V margins when transposing (only if not linked)
        if (didTranspose && !linkCheckbox.value) {
            var swapTemp = gapX;
            gapX = gapY;
            gapY = swapTemp;
        }

        if (didTranspose) {
            // 転置はマージン0で実行し、その後マージンを適用 / Transpose with 0 margins then apply margins
            applySpacing(0, 0);
            PreviewHistory.bump();

            runTransposeIfNeeded();
            PreviewHistory.bump();

            // 転置後の配置を新しい基準に / adopt transposed layout as baseline for spacing
            resetBaselineToCurrent();
            if (brickCheckbox.value) {
                if (honeycombCheckbox.value) applySpacingHexagon(gapX, gapY);
                else applySpacingBrick(gapX, gapY);
            } else {
                applySpacing(gapX, gapY);
            }
            PreviewHistory.bump();
        } else {
            if (brickCheckbox.value) {
                if (honeycombCheckbox.value) applySpacingHexagon(gapX, gapY);
                else applySpacingBrick(gapX, gapY);
            } else {
                applySpacing(gapX, gapY);
            }
            PreviewHistory.bump();
        }
    }

    // changeValueByArrowKey() から呼べるようにグローバルへ / expose for arrow-key handler
    updatePreview = updatePreviewImpl;
    $.global.updatePreview = updatePreviewImpl;

    // イベント / events
    horizontalGapInput.onChanging = function () { updatePreviewImpl(); };
    verticalGapInput.onChanging = function () {
        if (!linkCheckbox.value) {
            updatePreviewImpl();
        }
    };
    linkCheckbox.onClick = function () { updatePreviewImpl(); };
    forceGridCheckbox.onClick = function () { updatePreviewImpl(); };
    // 行列入れ替え / swap rows & columns
    transposeCheckbox.onClick = function () {
        if (transposeCheckbox.value) {
            // 転置フラグON（ワンショット扱い）/ set transpose flag (one-shot)
            didTranspose = true;
            transposeCheckbox.value = false;

            // 直前のマージンを保持（UIはそのまま）
            savedHorizontalGapText = horizontalGapInput.text;
            savedVerticalGapText = verticalGapInput.text;

            // 連動OFFなら左右/上下の値をUI上でも入れ替える / Swap UI values when Link is OFF
            if (!linkCheckbox.value) {
                var swapHorizontalText = horizontalGapInput.text;
                var swapVerticalText = verticalGapInput.text;
                horizontalGapInput.text = swapVerticalText;
                verticalGapInput.text = swapHorizontalText;
            }

            updatePreviewImpl();
            return;
        }
        // OFFは特に何もしない（ワンショットなので）
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

        // 転置確定時は左右/上下のマージンを入れ替える（連動OFFのときのみ）/ Swap margins on final apply when transposing (only if not linked)
        if (didTranspose && !linkCheckbox.value) {
            var swapTemp = finalGapX;
            finalGapX = finalGapY;
            finalGapY = swapTemp;
        }

        // プレビュー分を一括Undoしてから確定適用 / Clear preview history before final apply
        PreviewHistory.undo();
        resetBaselineToCurrent();
        if (didTranspose) {
            applySpacing(0, 0);
            runTransposeIfNeeded();
            resetBaselineToCurrent();
            if (brickCheckbox.value) {
                if (honeycombCheckbox.value) applySpacingHexagon(finalGapX, finalGapY);
                else applySpacingBrick(finalGapX, finalGapY);
            } else {
                applySpacing(finalGapX, finalGapY);
            }
        } else {
            if (brickCheckbox.value) {
                if (honeycombCheckbox.value) applySpacingHexagon(finalGapX, finalGapY);
                else applySpacingBrick(finalGapX, finalGapY);
            } else {
                applySpacing(finalGapX, finalGapY);
            }
        }
    } else {
        // キャンセル時：プレビュー分を一括Undoしてから初期状態へ / Undo preview then restore
        PreviewHistory.undo();
        restoreInitialPositions();
    }

    // 念のためカウンタ初期化 / reset counter
    PreviewHistory.start();

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
    var items = [];
    for (var i = 0; i < doc.selection.length; i++) {
        var selectionItem = doc.selection[i];
        items.push(selectionItem);
    }
    if (items.length < 2) {
        alert(L('alert.needTwo'));
        return;
    }

    // Sort selected items by top (Y) descending and left (X) ascending within same row
    var sortedByPosition = items.slice().sort(function (a, b) {
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
    for (var k = 0; k < items.length; k++) {
        var itemBounds = getLayoutBounds(items[k]);
        originalPositions.push({
            item: items[k],
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
            if (!item) return item.geometricBounds;

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
                        var cpi = item.compoundPathItems[j];
                        if (cpi.pathItems && cpi.pathItems.length > 0 && cpi.pathItems[0].clipping) {
                            return cpi.pathItems[0].geometricBounds;
                        }
                    }
                }
                // fallback
                return item.geometricBounds;
            }
        } catch (e) { }

        return item.geometricBounds;
    }

    /*
    行・列の推定 / build layout info
    */
    function buildLayoutInfo() {
        var bbs = [];
        var minW = Number.MAX_VALUE;
        var minH = Number.MAX_VALUE;

        for (var j = 0; j < items.length; j++) {
            var b = getLayoutBounds(items[j]);
            var w = b[2] - b[0];
            var h = b[1] - b[3];
            if (w < minW) minW = w;
            if (h < minH) minH = h;
            bbs.push({ item: items[j], gb: b });
        }

        // 列まとめ / group columns
        var colCenters = [];
        var colTolerance = minW * 0.5;
        if (colTolerance < 1) colTolerance = 1;

        bbs.sort(function (a, b) { return a.gb[0] - b.gb[0]; });

        for (var c = 0; c < bbs.length; c++) {
            var bb = bbs[c];
            var merged = false;
            var cx = bb.gb[0];
            for (var cc = 0; cc < colCenters.length; cc++) {
                if (Math.abs(colCenters[cc].x - cx) <= colTolerance) {
                    colCenters[cc].items.push(bb);
                    colCenters[cc].x = (colCenters[cc].x * (colCenters[cc].items.length - 1) + cx) / colCenters[cc].items.length;
                    merged = true;
                    break;
                }
            }
            if (!merged) {
                colCenters.push({ x: cx, items: [bb] });
            }
        }

        // 行まとめ / group rows
        var rowCenters = [];
        var rowTolerance = minH * 0.5;
        if (rowTolerance < 1) rowTolerance = 1;

        var bbsY = bbs.slice().sort(function (a, b) { return b.gb[1] - a.gb[1]; });

        for (var r = 0; r < bbsY.length; r++) {
            var bb2 = bbsY[r];
            var merged2 = false;
            var cy = bb2.gb[1];
            for (var rr = 0; rr < rowCenters.length; rr++) {
                if (Math.abs(rowCenters[rr].y - cy) <= rowTolerance) {
                    rowCenters[rr].items.push(bb2);
                    rowCenters[rr].y = (rowCenters[rr].y * (rowCenters[rr].items.length - 1) + cy) / rowCenters[rr].items.length;
                    merged2 = true;
                    break;
                }
            }
            if (!merged2) {
                rowCenters.push({ y: cy, items: [bb2] });
            }
        }

        // 並び順の確定 / sort
        colCenters.sort(function (a, b) { return a.x - b.x; });
        rowCenters.sort(function (a, b) { return b.y - a.y; });

        // 各列の最大幅 / max width per column
        var colWidths = [];
        for (var ci = 0; ci < colCenters.length; ci++) {
            var maxW = 0;
            for (var ci2 = 0; ci2 < colCenters[ci].items.length; ci2++) {
                var gb2 = colCenters[ci].items[ci2].gb;
                var w2 = gb2[2] - gb2[0];
                if (w2 > maxW) maxW = w2;
            }
            colWidths.push(maxW);
        }

        // 各行の最大高さ / max height per row
        var rowHeights = [];
        for (var ri = 0; ri < rowCenters.length; ri++) {
            var maxH = 0;
            for (var ri2 = 0; ri2 < rowCenters[ri].items.length; ri2++) {
                var gb3 = rowCenters[ri].items[ri2].gb;
                var h3 = gb3[1] - gb3[3];
                if (h3 > maxH) maxH = h3;
            }
            rowHeights.push(maxH);
        }

        var baseX = colCenters[0].x;
        var baseY = rowCenters[0].y;

        return {
            bbs: bbs,
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
        var bbs = [];
        var minW = Number.MAX_VALUE;
        var minH = Number.MAX_VALUE;

        for (var j = 0; j < items.length; j++) {
            var b = getLayoutBounds(items[j]);
            var w = b[2] - b[0];
            var h = b[1] - b[3];
            if (w < minW) minW = w;
            if (h < minH) minH = h;
            bbs.push({ item: items[j], gb: b });
        }

        // 行まとめ（上→下）/ group rows (top -> bottom)
        var rows = [];
        var rowTolerance = minH * 0.5;
        if (rowTolerance < 1) rowTolerance = 1;

        var bbsY = bbs.slice().sort(function (a, b) { return b.gb[1] - a.gb[1]; });
        for (var i = 0; i < bbsY.length; i++) {
            var bb = bbsY[i];
            var y = bb.gb[1];
            var placed = false;
            for (var r = 0; r < rows.length; r++) {
                if (Math.abs(rows[r].y - y) <= rowTolerance) {
                    rows[r].items.push(bb);
                    rows[r].y = (rows[r].y * (rows[r].items.length - 1) + y) / rows[r].items.length;
                    placed = true;
                    break;
                }
            }
            if (!placed) rows.push({ y: y, items: [bb] });
        }
        rows.sort(function (a, b) { return b.y - a.y; });

        // 各行の中を左→右で確定し、rowIndex/colIndexを付与 / sort within row and assign indices
        var maxCols = 0;
        for (var r2 = 0; r2 < rows.length; r2++) {
            rows[r2].items.sort(function (a, b) { return a.gb[0] - b.gb[0]; });
            if (rows[r2].items.length > maxCols) maxCols = rows[r2].items.length;
            for (var c2 = 0; c2 < rows[r2].items.length; c2++) {
                rows[r2].items[c2].rowIndex = r2;
                rows[r2].items[c2].colIndex = c2;
            }
        }

        // colWidths（列ごとの最大幅）/ max width per column index
        var colWidths = [];
        for (var c3 = 0; c3 < maxCols; c3++) {
            var maxW = 0;
            for (var r3 = 0; r3 < rows.length; r3++) {
                if (rows[r3].items.length > c3) {
                    var gb2 = rows[r3].items[c3].gb;
                    var w2 = gb2[2] - gb2[0];
                    if (w2 > maxW) maxW = w2;
                }
            }
            colWidths.push(maxW);
        }

        // rowHeights（行ごとの最大高さ）/ max height per row
        var rowHeights = [];
        for (var r4 = 0; r4 < rows.length; r4++) {
            var maxH = 0;
            for (var k = 0; k < rows[r4].items.length; k++) {
                var gb3 = rows[r4].items[k].gb;
                var h3 = gb3[1] - gb3[3];
                if (h3 > maxH) maxH = h3;
            }
            rowHeights.push(maxH);
        }

        // baseX/baseY は最左/最上 / baseX/baseY = top-left
        var baseX = Number.MAX_VALUE;
        var baseY = Number.MIN_VALUE;
        for (var q = 0; q < bbs.length; q++) {
            if (bbs[q].gb[0] < baseX) baseX = bbs[q].gb[0];
            if (bbs[q].gb[1] > baseY) baseY = bbs[q].gb[1];
        }

        // ダミーのcolCenters/rowCenters（互換のため）/ dummy centers (compat)
        var colCenters = [];
        for (var cc = 0; cc < maxCols; cc++) colCenters.push({ x: baseX, items: [] });
        var rowCenters = [];
        for (var rr = 0; rr < rows.length; rr++) rowCenters.push({ y: rows[rr].y, items: rows[rr].items });

        return {
            bbs: bbs,
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
        for (var kB = 0; kB < items.length; kB++) {
            var gbB = getLayoutBounds(items[kB]);
            originalPositions.push({ item: items[kB], left: gbB[0], top: gbB[1] });
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
        if (!items || items.length < 1) return;

        // ---- 調整パラメータ / tuning params
        var SNAP_X_TOL = 8.0;   // X方向の「同じ列」とみなす許容（pt）
        var SNAP_Y_TOL = 8.0;   // Y方向の「同じ行」とみなす許容（pt）

        // getLayoutBounds: [left, top, right, bottom]（グループは1つのbboxとして扱う）
        function leftX(it) { return getLayoutBounds(it)[0]; }
        function topY(it) { return getLayoutBounds(it)[1]; }

        // 近い値をクラスタリングして中心値配列を作る / cluster near values into centers
        function clusterValues(values, tol) {
            values.sort(function (a, b) { return a - b; });
            var centers = [];
            for (var i = 0; i < values.length; i++) {
                var v = values[i];
                var found = -1;
                for (var c = 0; c < centers.length; c++) {
                    if (Math.abs(v - centers[c]) <= tol) { found = c; break; }
                }
                if (found < 0) centers.push(v);
                else centers[found] = (centers[found] + v) / 2.0;
            }
            centers.sort(function (a, b) { return a - b; });
            return centers;
        }

        function nearestIndex(arr, v) {
            var best = 0;
            var bestD = Math.abs(v - arr[0]);
            for (var i = 1; i < arr.length; i++) {
                var d = Math.abs(v - arr[i]);
                if (d < bestD) { bestD = d; best = i; }
            }
            return best;
        }

        // 値の差分（隣接）の中央値 / median of adjacent diffs
        function medianDiff(sortedDescOrAsc) {
            if (sortedDescOrAsc.length < 2) return 0;
            var diffs = [];
            for (var i = 1; i < sortedDescOrAsc.length; i++) {
                diffs.push(Math.abs(sortedDescOrAsc[i] - sortedDescOrAsc[i - 1]));
            }
            diffs.sort(function (a, b) { return a - b; });
            return diffs[Math.floor(diffs.length / 2)];
        }

        // left/top を集める / collect left/top
        var xs = [], ys = [];
        for (var k = 0; k < items.length; k++) {
            xs.push(leftX(items[k]));
            ys.push(topY(items[k]));
        }

        var colXs = clusterValues(xs, SNAP_X_TOL); // left -> right
        var rowYs = clusterValues(ys, SNAP_Y_TOL); // will sort top -> bottom next

        // Illustrator座標では上ほどYが大きいことが多いので「上→下」/ sort top -> bottom
        rowYs.sort(function (a, b) { return b - a; });

        var rowCount = rowYs.length;
        var colCount = colXs.length;

        // 各オブジェクトを(行,列)に割り当て / assign each item to (row, col)
        var occupancy = {}; // key "r,c" -> item
        var mapping = [];   // {item, r, c}
        for (var m = 0; m < items.length; m++) {
            var it = items[m];
            var r = nearestIndex(rowYs, topY(it));
            var c = nearestIndex(colXs, leftX(it));
            var key = r + "," + c;
            if (occupancy[key]) {
                alert(L('alert.cellConflict') + "(" + r + "," + c + ")");
                return;
            }
            occupancy[key] = it;
            mapping.push({ item: it, r: r, c: c });
        }

        // 元の列/行のピッチを推定（隣接差の中央値）
        // ※行や列が1つしかない場合は0になる
        var pitchX = (colXs.length >= 2) ? medianDiff(colXs) : 0;
        var pitchY = (rowYs.length >= 2) ? medianDiff(rowYs) : 0;

        // 1行→1列、1列→1行にも対応
        // - 1行しかない場合: 横方向ピッチ(pitchX)を縦方向の並び間隔として流用
        // - 1列しかない場合: 縦方向ピッチ(pitchY)を横方向の並び間隔として流用
        var pitchXEff = pitchX;
        var pitchYEff = pitchY;

        if (rowCount === 1 && colCount === 1) {
            // ほぼ重なり等で1セル扱いになったケース
            return;
        }

        if (rowCount === 1 && colCount > 1) {
            // 1行 → 1列
            if (pitchX === 0) return;
            pitchXEff = pitchX;
            pitchYEff = pitchX; // 横ピッチを縦へ流用
        } else if (colCount === 1 && rowCount > 1) {
            // 1列 → 1行
            if (pitchY === 0) return;
            pitchXEff = pitchY; // 縦ピッチを横へ流用
            pitchYEff = pitchY;
        } else {
            // 通常（2行以上 かつ 2列以上）
            if (pitchX === 0 || pitchY === 0) return;
        }

        // 転置後グリッドの基準（左上固定）/ origin at top-left
        var originLeft = colXs[0];
        var originTop = rowYs[0];

        // 転置: newCol = oldRow, newRow = oldCol
        for (var t = 0; t < mapping.length; t++) {
            var obj = mapping[t].item;
            var r0 = mapping[t].r;
            var c0 = mapping[t].c;

            var newCol = r0;
            var newRow = c0;

            var targetLeft = originLeft + newCol * pitchXEff;
            var targetTop = originTop - newRow * pitchYEff;

            var b = getLayoutBounds(obj);
            var curLeft = b[0];
            var curTop = b[1];

            obj.translate(targetLeft - curLeft, targetTop - curTop);
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

    /* 間隔を適用 / apply spacing */
    function applySpacing(gapX, gapY) {
        // いったん元に戻す / restore first
        restoreOriginalPositions();

        var bbs = layoutInfo.bbs;
        var colCenters = layoutInfo.colCenters;
        var rowCenters = layoutInfo.rowCenters;
        var colWidths = layoutInfo.colWidths;
        var rowHeights = layoutInfo.rowHeights;
        var baseX = layoutInfo.baseX;
        var baseY = layoutInfo.baseY;

        for (var k2 = 0; k2 < bbs.length; k2++) {
            var cur = bbs[k2];

            // 元座標を探す / find original
            var origLeft = null, origTop = null;
            for (var oo = 0; oo < originalPositions.length; oo++) {
                if (originalPositions[oo].item === cur.item) {
                    origLeft = originalPositions[oo].left;
                    origTop = originalPositions[oo].top;
                    break;
                }
            }
            if (origLeft === null) {
                origLeft = cur.gb[0];
                origTop = cur.gb[1];
            }

            // Force Gridでは事前に割り当てた(row,col)を優先 / In Force Grid, prefer pre-assigned (row,col)
            var colIndex = (typeof cur.colIndex === 'number') ? cur.colIndex : null;
            var rowIndex = (typeof cur.rowIndex === 'number') ? cur.rowIndex : null;

            if (colIndex === null) {
                colIndex = 0;
                var minDX = Number.MAX_VALUE;
                for (var c2 = 0; c2 < colCenters.length; c2++) {
                    var dx = Math.abs(colCenters[c2].x - cur.gb[0]);
                    if (dx < minDX) { minDX = dx; colIndex = c2; }
                }
            }

            if (rowIndex === null) {
                rowIndex = 0;
                var minDY = Number.MAX_VALUE;
                for (var r2 = 0; r2 < rowCenters.length; r2++) {
                    var dy = Math.abs(rowCenters[r2].y - cur.gb[1]);
                    if (dy < minDY) { minDY = dy; rowIndex = r2; }
                }
            }

            // 新しいX / new X
            var newX = baseX;
            for (var cc2 = 0; cc2 < colIndex; cc2++) {
                newX += colWidths[cc2] + gapX;
            }

            // 新しいY / new Y
            var newY = baseY;
            for (var rr2 = 0; rr2 < rowIndex; rr2++) {
                newY -= (rowHeights[rr2] + gapY);
            }

            var dxMove = newX - origLeft;
            var dyMove = newY - origTop;
            cur.item.translate(dxMove, dyMove);
        }

        // 再描画 / redraw
        app.redraw();
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

    // 六角形っぽいか判定（PathItem 6点）/ Detect hexagon-like paths (6 points)
    function isHexagonLike(item) {
        try {
            if (!item) return false;
            if (item.typename === 'PathItem') {
                return !!item.closed && item.pathPoints && item.pathPoints.length === 6;
            }
            if (item.typename === 'CompoundPathItem') {
                if (item.pathItems && item.pathItems.length > 0) {
                    var p = item.pathItems[0];
                    return !!p.closed && p.pathPoints && p.pathPoints.length === 6;
                }
            }
        } catch (e) { }
        return false;
    }

    // 選択が六角形中心か（過半数）/ Is the selection mostly hexagons?
    function isMostlyHexagons(itemsArr) {
        if (!itemsArr || itemsArr.length === 0) return false;
        var hexCount = 0;
        for (var i = 0; i < itemsArr.length; i++) {
            if (isHexagonLike(itemsArr[i])) hexCount++;
        }
        return hexCount >= Math.ceil(itemsArr.length * 0.5);
    }

    /* 間隔を適用（レンガ状）/ apply spacing (brick layout) */
    function applySpacingBrick(gapX, gapY) {
        // いったん元に戻す / restore first
        restoreOriginalPositions();

        var bbs = layoutInfo.bbs;
        var colCenters = layoutInfo.colCenters;
        var rowCenters = layoutInfo.rowCenters;
        var colWidths = layoutInfo.colWidths;
        var rowHeights = layoutInfo.rowHeights;
        var baseX = layoutInfo.baseX;
        var baseY = layoutInfo.baseY;

        // 列ピッチ（目標）：幅 + gapX を基準にする / Target column pitch: width + gapX
        var wArr = [];
        if (colWidths && colWidths.length > 0) {
            for (var wi = 0; wi < colWidths.length; wi++) wArr.push(colWidths[wi]);
            wArr.sort(function (a, b) { return a - b; });
        }
        var baseW = (wArr.length > 0) ? wArr[Math.floor(wArr.length / 2)] : 0; // median width
        var pitch = (baseW > 0) ? (baseW + gapX) : 0;

        // fallback：列が1つ等で推定できない場合は現状の列位置差 / fallback to current centers diff
        if (pitch === 0) {
            var colX = [];
            for (var iC = 0; iC < colCenters.length; iC++) colX.push(colCenters[iC].x);
            colX.sort(function (a, b) { return a - b; });
            pitch = medianAdjacentDiff(colX);
        }
        var halfPitch = pitch / 2.0;

        for (var k2 = 0; k2 < bbs.length; k2++) {
            var cur = bbs[k2];

            // 元座標を探す / find original
            var origLeft = null, origTop = null;
            for (var oo = 0; oo < originalPositions.length; oo++) {
                if (originalPositions[oo].item === cur.item) {
                    origLeft = originalPositions[oo].left;
                    origTop = originalPositions[oo].top;
                    break;
                }
            }
            if (origLeft === null) {
                origLeft = cur.gb[0];
                origTop = cur.gb[1];
            }

            // Force Gridでは事前に割り当てた(row,col)を優先 / In Force Grid, prefer pre-assigned (row,col)
            var colIndex = (typeof cur.colIndex === 'number') ? cur.colIndex : null;
            var rowIndex = (typeof cur.rowIndex === 'number') ? cur.rowIndex : null;

            if (colIndex === null) {
                colIndex = 0;
                var minDX = Number.MAX_VALUE;
                for (var c2 = 0; c2 < colCenters.length; c2++) {
                    var dx = Math.abs(colCenters[c2].x - cur.gb[0]);
                    if (dx < minDX) { minDX = dx; colIndex = c2; }
                }
            }

            if (rowIndex === null) {
                rowIndex = 0;
                var minDY = Number.MAX_VALUE;
                for (var r2 = 0; r2 < rowCenters.length; r2++) {
                    var dy = Math.abs(rowCenters[r2].y - cur.gb[1]);
                    if (dy < minDY) { minDY = dy; rowIndex = r2; }
                }
            }

            // 新しいX / new X
            var newX = baseX;
            for (var cc2 = 0; cc2 < colIndex; cc2++) {
                newX += colWidths[cc2] + gapX;
            }

            // 新しいY / new Y
            var newY = baseY;
            for (var rr2 = 0; rr2 < rowIndex; rr2++) {
                newY -= (rowHeights[rr2] + gapY);
            }

            // レンガ状：奇数行を半ピッチずらす / Brick: shift odd rows by half pitch
            if (halfPitch !== 0 && (rowIndex % 2 === 1)) {
                newX += halfPitch;
            }

            var dxMove = newX - origLeft;
            var dyMove = newY - origTop;
            cur.item.translate(dxMove, dyMove);
        }

        app.redraw();
    }
    /* 間隔を適用（六角形/ハニカム）/ apply spacing (hexagon/honeycomb layout) */
    function applySpacingHexagon(gapX, gapY) {
        restoreOriginalPositions();

        var bbs = layoutInfo.bbs;
        var colCenters = layoutInfo.colCenters;
        var rowCenters = layoutInfo.rowCenters;
        var colWidths = layoutInfo.colWidths;
        var rowHeights = layoutInfo.rowHeights;
        var baseX = layoutInfo.baseX;
        var baseY = layoutInfo.baseY;

        // 列ピッチ（目標）：幅 + gapX を基準にする / Target column pitch: width + gapX
        var wArr = [];
        if (colWidths && colWidths.length > 0) {
            for (var wi = 0; wi < colWidths.length; wi++) wArr.push(colWidths[wi]);
            wArr.sort(function (a, b) { return a - b; });
        }
        var baseW = (wArr.length > 0) ? wArr[Math.floor(wArr.length / 2)] : 0; // median width
        var pitch = (baseW > 0) ? (baseW + gapX) : 0;
        // fallback：列が1つ等で推定できない場合は現状の列位置差 / fallback to current centers diff
        if (pitch === 0) {
            var colX = [];
            for (var iC = 0; iC < colCenters.length; iC++) colX.push(colCenters[iC].x);
            colX.sort(function (a, b) { return a - b; });
            pitch = medianAdjacentDiff(colX);
        }
        var halfPitch = pitch / 2.0;

        var rowStepFactor = 0.75;

        for (var k2 = 0; k2 < bbs.length; k2++) {
            var cur = bbs[k2];

            var origLeft = null, origTop = null;
            for (var oo = 0; oo < originalPositions.length; oo++) {
                if (originalPositions[oo].item === cur.item) {
                    origLeft = originalPositions[oo].left;
                    origTop = originalPositions[oo].top;
                    break;
                }
            }
            if (origLeft === null) { origLeft = cur.gb[0]; origTop = cur.gb[1]; }

            // Force Gridでは事前に割り当てた(row,col)を優先 / In Force Grid, prefer pre-assigned (row,col)
            var colIndex = (typeof cur.colIndex === 'number') ? cur.colIndex : null;
            var rowIndex = (typeof cur.rowIndex === 'number') ? cur.rowIndex : null;

            if (colIndex === null) {
                colIndex = 0;
                var minDX = Number.MAX_VALUE;
                for (var c2 = 0; c2 < colCenters.length; c2++) {
                    var dx = Math.abs(colCenters[c2].x - cur.gb[0]);
                    if (dx < minDX) { minDX = dx; colIndex = c2; }
                }
            }

            if (rowIndex === null) {
                rowIndex = 0;
                var minDY = Number.MAX_VALUE;
                for (var r2 = 0; r2 < rowCenters.length; r2++) {
                    var dy = Math.abs(rowCenters[r2].y - cur.gb[1]);
                    if (dy < minDY) { minDY = dy; rowIndex = r2; }
                }
            }

            var newX = baseX;
            for (var cc2 = 0; cc2 < colIndex; cc2++) newX += colWidths[cc2] + gapX;

            var newY = baseY;
            for (var rr2 = 0; rr2 < rowIndex; rr2++) {
                newY -= (rowHeights[rr2] * rowStepFactor) + gapY;
            }

            if (halfPitch !== 0 && (rowIndex % 2 === 1)) newX += halfPitch;

            cur.item.translate(newX - origLeft, newY - origTop);
        }

        app.redraw();
    }

    // ダイアログ表示 / show dialog
    showGridSpacingDialog(
        applySpacing,
        applySpacingBrick,
        applySpacingHexagon,
        transposeGridWithHoles,
        restoreOriginalPositions,
        restoreInitialPositions,
        resetBaselineToCurrent,
        getCurrentUnitLabel,
        changeValueByArrowKey,
        initialGapX
    );
}

// 実行 / run
main();

})();