#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
PathTextToolkit.jsx
 
### 更新日：
20260518
 
### 概要：
ポイント文字／パス上文字を、用途に応じて「作成」「分離」「調整」できるツールです。

- パス上文字にする：選択したパス（または既存のパス上文字のパス）に沿ってテキストを配置
- アーチ状のパスを生成：テキストからアーチのパスを自動生成してパス上文字を作成
- 正円を生成：テキスト幅から正円パスを自動生成してパス上文字を作成

- テキストを分離：パス上文字を「テキスト」と「パス」に分離（書式保持／保持しない、パス削除オプションあり）

- 行揃え：左／中央／右／両端揃え
- 効果：虹／歪み／3Dリボン／階段／引力
- 位置：開始位置／終了位置（必要なときだけチェックONで適用）
- テキスト調整：ベースライン／トラッキング／文字サイズ（現在値に対して増減）
- フィット：パス上文字が端まで収まるように文字サイズを自動調整（開いたパスのみ）

- プレビュー：ダイアログ操作中に結果を確認（OFFで元に戻せます）
- テキスト編集：内容をまとめて置換（複数選択にも対応）
*/

(function () {

    /* バージョン / Version */
    var SCRIPT_VERSION = "v1.3.3";

    /* 表示言語を判定（ja / en）/ Detect the UI language (ja / en) */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "パス上文字（変換と調整）", en: "Text on a Path (Convert & Adjust)" },

        /* Panels / パネル */
        panelProcess: { ja: "パス上文字にする", en: "Create Path Text" },
        panelSplit: { ja: "テキストを分離", en: "Split Text" },
        panelAlign: { ja: "行揃え", en: "Alignment" },
        panelOption: { ja: "オプション", en: "Options" },
        panelEffect: { ja: "効果", en: "Effect" },
        panelPosition: { ja: "位置", en: "Position" },
        panelTextAdjust: { ja: "テキスト調整", en: "Text Adjust" },
        panelFitWidth: { ja: "パス幅に合わせる", en: "Fit to Path Width" },

        /* Process radios / 処理ラジオ */
        toPathText: { ja: "パス上文字にする", en: "Convert to Path Text" },
        genArcPath: { ja: "アーチ状のパスを生成", en: "Generate Arc Path" },
        genCircle: { ja: "正円を生成", en: "Generate Circle" },
        btnTextEdit: { ja: "テキスト編集", en: "Edit Text" },
        dlgTextEditTitle: { ja: "テキスト編集", en: "Edit Text" },
        dlgTextEditHint: { ja: "内容を編集してOKで反映します。複数選択の場合は全て置換します。", en: "Edit the content and press OK to apply. If multiple items are selected, it replaces all." },
        splitTextAndPath: { ja: "書式を保持", en: "Keep Formatting" },
        splitTextAndPathNoFormat: { ja: "書式を保持しない", en: "Do Not Keep Formatting" },
        splitDeletePath: { ja: "パスを削除", en: "Remove Path" },

        /* Alignment radios / 行揃え */
        alignLeft: { ja: "左揃え", en: "Left" },
        alignCenter: { ja: "中央", en: "Center" },
        alignRight: { ja: "右揃え", en: "Right" },
        alignFullJustify: { ja: "両端揃え", en: "Justify" },

        /* Options / オプション */
        optReverse: { ja: "内側配置", en: "Inside" },

        arcDirection: { ja: "アーチ方向", en: "Arc Direction" },
        arcUp: { ja: "上", en: "Up" },
        arcDown: { ja: "下", en: "Down" },
        arcRoundness: { ja: "まるみ", en: "Roundness" },

        /* Fit to path width / パス幅に合わせる */
        fitWidthNone: { ja: "しない", en: "None" },
        fitWidthFontSize: { ja: "文字サイズを調整", en: "Adjust Font Size" },
        fitWidthTracking: { ja: "トラッキングを調整", en: "Adjust Tracking" },

        /* Effects / 効果 */
        effectRainbow: { ja: "虹", en: "Rainbow" },
        effectDistort: { ja: "歪み", en: "Skew" },
        effectRibbon: { ja: "3D リボン", en: "3D Ribbon" },
        effectStep: { ja: "階段", en: "Stair Step" },
        effectGravity: { ja: "引力", en: "Gravity" },

        /* Position / 位置 */
        startPos: { ja: "開始位置", en: "Start" },
        endPos: { ja: "終了位置", en: "End" },

        /* Text Adjust / テキスト調整 */
        baseShift: { ja: "ベースライン", en: "Baseline Shift" },
        tracking: { ja: "トラッキング", en: "Tracking" },
        fontSize: { ja: "文字サイズ", en: "Font Size" },

        /* Footer / フッター */
        preview: { ja: "プレビュー", en: "Preview" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },

        /* Alerts / アラート */
        alertNoDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
        alertNoText: { ja: "対象のテキストが見つかりません", en: "No target text found." },
        alertNeedPath: { ja: "パスを一緒に選択してください", en: "Select a path together." },
        alertDupPathFail: { ja: "パスの複製に失敗しました", en: "Failed to duplicate the path." },
        alertArcFail: { ja: "アーチ状のパス生成に失敗しました", en: "Failed to generate an arc path." },
        alertNeedPathText: { ja: "パス上文字を選択してください", en: "Select path text." }
    };

    /* キーからラベル文字列を取得 / Resolve a label string from its key */
    function L(key) {
        // LABELS は固定のオブジェクトリテラル。&& ガード済みで例外は起きないため try 不要
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].ja) return LABELS[key].ja;
        return key;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (lang === 'ja' ? '：' : ':');
    }

    /* edittext の値を矢印キーで増減（Up/Down=±1, Shift=±10, Option=±0.1）/ Change an edittext value with arrow keys */
    // Up/Down: ±1, Shift+Up/Down: ±10, Option(Alt)+Up/Down: ±0.1
    // allowNegative: false -> clamp to >= 0
    // onChanged: optional callback invoked after value change
    function changeValueByArrowKey(editText, allowNegative, onChanged) {
        if (!editText) return;
        try {
            editText.addEventListener('keydown', function (event) {
                try {
                    if (!event || (event.keyName !== 'Up' && event.keyName !== 'Down')) return;

                    var value = Number(editText.text);
                    if (isNaN(value)) return;

                    var keyboard = ScriptUI.environment.keyboardState;
                    var delta = 1;

                    if (keyboard.shiftKey) {
                        delta = 10;
                        // Snap to multiples of 10 for Shift
                        if (event.keyName === 'Up') {
                            value = Math.ceil((value + 1) / delta) * delta;
                        } else {
                            value = Math.floor((value - 1) / delta) * delta;
                        }
                    } else if (keyboard.altKey) {
                        delta = 0.1;
                        if (event.keyName === 'Up') value += delta;
                        else value -= delta;
                    } else {
                        delta = 1;
                        if (event.keyName === 'Up') value += delta;
                        else value -= delta;
                    }

                    // Round
                    if (keyboard.altKey) value = Math.round(value * 10) / 10;
                    else value = Math.round(value);

                    // Clamp
                    if (!allowNegative && value < 0) value = 0;

                    editText.text = String(value);

                    // Prevent default arrow key behavior (cursor move)
                    try { event.preventDefault(); } catch (_) { }

                    if (onChanged && typeof onChanged === 'function') {
                        try { onChanged(); } catch (_) { }
                    }
                } catch (_) { }
            });
        } catch (_) { }
    }

    // Get items
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return false;
    }
    var doc = app.activeDocument;
    var selection = doc.selection;
    // Base selection snapshot (used for stable preview while dialog is open)
    var __baseSelection = [];
    try { __baseSelection = selection.slice(0); } catch (_) { __baseSelection = []; }

    var targetItems = getTargetTextItems(selection);
    var selectedPaths = getSelectedPathItems(selection);


    /* 現在の選択からターゲット・パスと UI 有効状態を取り直す / Re-read targets, paths and UI enabled-state from the current selection */
    function __refreshInputsFromSelection() {
        var curSel = [];
        try { curSel = doc.selection; } catch (_) { curSel = []; }

        // During preview, selection may become empty; fall back to the base selection.
        if (!curSel || curSel.length === 0) {
            curSel = __baseSelection;
        }

        targetItems = getTargetTextItems(curSel);
        selectedPaths = getSelectedPathItems(curSel);

        // Detect availability
        var hasSelectedPath = (selectedPaths && selectedPaths.length > 0);
        var hasPathTextTarget = false;
        try {
            for (var ti = 0; ti < targetItems.length; ti++) {
                var t = targetItems[ti];
                if (t && t.typename === 'TextFrame') {
                    try {
                        if (t.kind === TextType.PATHTEXT && t.textPath) {
                            hasPathTextTarget = true;
                            break;
                        }
                    } catch (_) { }
                }
            }
        } catch (_) { }

        // "パス上文字に" is available if a path is selected OR a PathText provides its own path
        var canToPathText = hasSelectedPath || hasPathTextTarget;
        // "テキストとパスに分離" is available only when PathText is selected
        var canSplit = hasPathTextTarget;

        // If UI is not yet constructed, stop here (selection analysis only)
        if (!rbToPathText || !rbGenArcPath || !rbGenCircle || !rbSplitTextAndPath || !rbSplitTextAndPathNoFormat) {
            return;
        }

        try { rbToPathText.enabled = canToPathText; } catch (_) { }
        try { rbSplitTextAndPath.enabled = canSplit; } catch (_) { }
        try { cbSplitDeletePath.enabled = canSplit; } catch (_) { }
        try { rbSplitTextAndPathNoFormat.enabled = canSplit; } catch (_) { }
        try { rbGenCircle.enabled = true; } catch (_) { }

        // Auto-switch only if the currently selected mode is not available
        try {
            if (rbToPathText.value && !canToPathText) {
                rbToPathText.value = false;
                rbGenArcPath.value = true;
                applyArcModePresets();
            }
            if ((rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value) && !canSplit) {
                rbSplitTextAndPath.value = false;
                rbSplitTextAndPathNoFormat.value = false;
                rbGenArcPath.value = true;
                applyArcModePresets();
            }
        } catch (_) { }

        // Update panel enabled state only after UI is ready
        try { __updatePanelsByMode(); } catch (_) { }
    }

    // Validation
    if (targetItems.length === 0) {
        alert(L('alertNoText'));
        return false;
    }

    /* パネル共通設定 / Shared panel setup */
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* パネルを縦並び・共通マージン/間隔で初期化 / Initialize a panel as a column with shared margins/spacing */
    function setupPanel(panel, spacing) {
        panel.orientation = 'column';
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = 'fill';
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
    }

    /* ダイアログ表示 / Show dialog */
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = [15, 20, 15, 15];

    /* 2カラムレイアウト / Two-column layout */
    var mainRow = dlg.add('group');
    mainRow.orientation = 'row';
    mainRow.alignChildren = ['fill', 'top'];
    mainRow.alignment = ['fill', 'top'];

    var leftCol = mainRow.add('group');
    leftCol.orientation = 'column';
    leftCol.alignChildren = ['fill', 'top'];
    leftCol.alignment = ['fill', 'top'];

    var rightCol = mainRow.add('group');
    rightCol.orientation = 'column';
    rightCol.alignChildren = ['fill', 'top'];
    rightCol.alignment = ['fill', 'top'];

    // 処理パネル（パス上文字にする / アーチ生成 / 正円生成）/ Process panel
    var pnlProcess = leftCol.add('panel', undefined, L('panelProcess'));
    setupPanel(pnlProcess);

    var rbToPathText = pnlProcess.add('radiobutton', undefined, L('toPathText'));
    var rbGenArcPath = pnlProcess.add('radiobutton', undefined, L('genArcPath'));
    var rbGenCircle = pnlProcess.add('radiobutton', undefined, L('genCircle'));
    rbToPathText.value = true;

    // Text edit button (logic improved)
    var btnTextEdit = pnlProcess.add('button', undefined, L('btnTextEdit'));
    btnTextEdit.alignment = ['left', 'center'];
    btnTextEdit.onClick = function () {
        // --- Edit dialog logic ---
        var editDlg = new Window('dialog', L('dlgTextEditTitle'));
        editDlg.orientation = 'column';
        editDlg.alignChildren = ['fill', 'top'];
        editDlg.margins = [15, 20, 15, 15];

        // Hint
        var st = editDlg.add('statictext', undefined, L('dlgTextEditHint'), { multiline: true });
        try { st.preferredSize.width = 360; } catch (_) { }

        // Resolve targets at click-time (do not rely on stale snapshot)
        var __editTargets = [];
        try {
            var __curSel = [];
            try { __curSel = doc.selection; } catch (_) { __curSel = []; }
            if (!__curSel || __curSel.length === 0) __curSel = __baseSelection;
            __editTargets = getTargetTextItems(__curSel);
        } catch (_) { __editTargets = []; }

        // Initial text
        var initText = '';
        try {
            if (__editTargets && __editTargets.length > 0 && __editTargets[0] && __editTargets[0].typename === 'TextFrame') {
                initText = String(__editTargets[0].contents);
            }
        } catch (_) { initText = ''; }

        var et = editDlg.add('edittext', undefined, initText, { multiline: true });
        et.preferredSize = [320, 80];

        var btnRow = editDlg.add('group');
        btnRow.orientation = 'row';
        btnRow.alignChildren = ['center', 'center'];
        var btnCancel2 = btnRow.add('button', undefined, L('cancel'));
        var btnOk2 = btnRow.add('button', undefined, L('ok'), { name: 'ok' });

        btnCancel2.onClick = function () {
            editDlg.close(0);
        };
        btnOk2.onClick = function () {
            var newText = et.text;
            // Apply to all current targets at click-time
            for (var i = 0; i < __editTargets.length; i++) {
                var textFrame = __editTargets[i];
                if (!textFrame || textFrame.typename !== 'TextFrame') continue;
                try { textFrame.contents = newText; } catch (_) { }
            }
            // Refresh input states and preview
            __refreshInputsFromSelection();
            refreshPreviewIfNeeded();
            try { app.redraw(); } catch (_) { }
            editDlg.close(1);
        };
        editDlg.show();
    };
    btnTextEdit.height = 16;

    // Split panel (under Process)
    var pnlSplit = leftCol.add('panel', undefined, L('panelSplit'));
    setupPanel(pnlSplit);

    var rbSplitTextAndPath = pnlSplit.add('radiobutton', undefined, L('splitTextAndPath'));
    var rbSplitTextAndPathNoFormat = pnlSplit.add('radiobutton', undefined, L('splitTextAndPathNoFormat'));
    rbSplitTextAndPath.value = false;
    rbSplitTextAndPathNoFormat.value = false;

    var cbSplitDeletePath = pnlSplit.add('checkbox', undefined, L('splitDeletePath'));
    cbSplitDeletePath.value = false; // デフォルトOFF

    // --- Mutually exclusive helpers for Process/Split radios ---
    /* 「処理」ラジオをすべて OFF（排他制御用）/ Turn off all "Process" radios (mutual-exclusion helper) */
    function __clearProcessRadios() {
        try { rbToPathText.value = false; } catch (_) { }
        try { rbGenArcPath.value = false; } catch (_) { }
        try { rbGenCircle.value = false; } catch (_) { }
    }

    /* 「分離」ラジオをすべて OFF（排他制御用）/ Turn off all "Split" radios (mutual-exclusion helper) */
    function __clearSplitRadios() {
        try { rbSplitTextAndPath.value = false; } catch (_) { }
        try { rbSplitTextAndPathNoFormat.value = false; } catch (_) { }
    }

    // Auto-enable/disable based on selection (selected path OR PathText can provide a path)
    var hasSelectedPath = (selectedPaths && selectedPaths.length > 0);
    var hasPathTextTarget = false;
    try {
        for (var t0 = 0; t0 < targetItems.length; t0++) {
            if (targetItems[t0] && targetItems[t0].typename === 'TextFrame') {
                if (targetItems[t0].kind === TextType.PATHTEXT && targetItems[t0].textPath) {
                    hasPathTextTarget = true;
                    break;
                }
            }
        }
    } catch (_) { }

    // If PathText is selected, default effect = Rainbow (apply after effect UI is created)
    var __defaultRainbow = hasPathTextTarget;

    var canToPathText = hasSelectedPath || hasPathTextTarget;
    var canSplit = hasPathTextTarget;

    try { rbToPathText.enabled = canToPathText; } catch (_) { }
    try { rbSplitTextAndPath.enabled = canSplit; } catch (_) { }
    try { rbSplitTextAndPathNoFormat.enabled = canSplit; } catch (_) { }
    try { cbSplitDeletePath.enabled = canSplit; } catch (_) { }
    try { rbGenCircle.enabled = true; } catch (_) { }

    if (!canToPathText && rbToPathText.value) {
        rbToPathText.value = false;
        rbGenArcPath.value = true;
        // Arc presets are applied after the whole UI is built (see applyArcModePresets call).
    }
    if (!canSplit && (rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value)) {
        rbSplitTextAndPath.value = false;
        rbSplitTextAndPathNoFormat.value = false;
        rbGenArcPath.value = true;
        // Arc presets are applied after the whole UI is built (see applyArcModePresets call).
    }

    // Option panel (above Alignment in RIGHT column)
    var pnlOption = rightCol.add('panel', undefined, L('panelOption'));
    setupPanel(pnlOption);

    var cbReverse = pnlOption.add('checkbox', undefined, L('optReverse'));
    cbReverse.value = false;

    // Arc direction UI
    var grpArcDir = pnlOption.add('group');
    grpArcDir.orientation = 'row';
    grpArcDir.alignChildren = ['left', 'center'];

    var stArcDir = grpArcDir.add('statictext', undefined, labelText('arcDirection'));
    stArcDir.preferredSize.width = 70;

    var rbArcUp = grpArcDir.add('radiobutton', undefined, L('arcUp'));
    var rbArcDown = grpArcDir.add('radiobutton', undefined, L('arcDown'));
    rbArcUp.value = true;

    // Arc roundness UI (controls how round/curved the generated arch is)
    var grpArcRoundness = pnlOption.add('group');
    grpArcRoundness.orientation = 'row';
    grpArcRoundness.alignChildren = ['left', 'center'];

    var stArcRoundness = grpArcRoundness.add('statictext', undefined, L('arcRoundness'));

    // Slider: 0 = flat, 50 = default arch, 100 = roundest
    var slArcRoundness = grpArcRoundness.add('slider', undefined, 50, 0, 100);
    slArcRoundness.preferredSize.width = 130;

    // パス幅パネルの上に余白（10px）/ Spacer (10px) above the Fit-to-path-width panel
    var fitWidthSpacer = pnlOption.add('group');
    fitWidthSpacer.preferredSize.height = 5;

    /* パス幅に合わせるパネル（オプションパネル内の最下部）/ Fit-to-path-width panel (bottom of the Options panel) */
    var pnlFitWidth = pnlOption.add('panel', undefined, L('panelFitWidth'));
    setupPanel(pnlFitWidth);

    var rbFitWidthNone = pnlFitWidth.add('radiobutton', undefined, L('fitWidthNone'));
    var rbFitWidthFontSize = pnlFitWidth.add('radiobutton', undefined, L('fitWidthFontSize'));
    var rbFitWidthTracking = pnlFitWidth.add('radiobutton', undefined, L('fitWidthTracking'));

    // Default: しない / Default: None
    rbFitWidthNone.value = true;

    /* パネル選択に応じてパス幅フィットを適用 / Apply path-width fit according to the panel selection */
    function applyFitToPathWidth(frames) {
        try {
            if (rbFitWidthFontSize && rbFitWidthFontSize.value) {
                fitTextToOpenPath(frames);            // 文字サイズで合わせる / fit by font size
            } else if (rbFitWidthTracking && rbFitWidthTracking.value) {
                fitTextToOpenPathByTracking(frames);  // トラッキングで合わせる / fit by tracking
            }
            // 「しない」: 何もしない / "None": do nothing
        } catch (_) { }
    }

    // -------------------------------------------------
    // Full-width (span across columns): 効果 -> 位置 -> ベースラインシフト
    // -------------------------------------------------
    var wideCol = dlg.add('group');
    wideCol.orientation = 'column';
    wideCol.alignChildren = ['fill', 'top'];
    wideCol.alignment = ['fill', 'top'];

    // Effect panel (full-width)
    var pnlEffect = wideCol.add('panel', undefined, L('panelEffect'));
    // 効果は5項目を横並びにするため setupPanel（縦並び）は使わず個別設定
    pnlEffect.orientation = 'row';
    pnlEffect.alignChildren = ['center', 'center'];
    pnlEffect.alignment = 'fill';
    pnlEffect.margins = PANEL_MARGINS;
    pnlEffect.spacing = PANEL_SPACING;

    var rbEffectRainbow = pnlEffect.add('radiobutton', undefined, L('effectRainbow'));
    var rbEffectDistort = pnlEffect.add('radiobutton', undefined, L('effectDistort'));
    var rbEffectRibbon = pnlEffect.add('radiobutton', undefined, L('effectRibbon'));
    var rbEffectStep = pnlEffect.add('radiobutton', undefined, L('effectStep'));
    var rbEffectGravity = pnlEffect.add('radiobutton', undefined, L('effectGravity'));

    // No effect selected by default
    rbEffectRainbow.value = false;
    rbEffectDistort.value = false;
    rbEffectRibbon.value = false;
    rbEffectStep.value = false;
    rbEffectGravity.value = false;

    // Apply default rainbow AFTER effect UI exists
    if (__defaultRainbow) {
        try { rbEffectRainbow.value = true; } catch (_) { }
    }

    // Position (start/end) panel  (moved to FULL WIDTH)
    var pnlPosition = wideCol.add('panel', undefined, L('panelPosition'));
    setupPanel(pnlPosition);

    var rowStart = pnlPosition.add('group');
    rowStart.orientation = 'row';
    rowStart.alignChildren = ['left', 'center'];
    var cbStartT = rowStart.add('checkbox', undefined, '');
    cbStartT.value = false;
    var stStart = rowStart.add('statictext', undefined, L('startPos'));
    stStart.preferredSize.width = 60;
    var etStartT = rowStart.add('edittext', undefined, '0.0');
    etStartT.characters = 5;
    var slStartT = rowStart.add('slider', undefined, 0, 0, 400); // 0.0 - 4.0
    slStartT.preferredSize.width = 180;
    // Arrow-key support for startT
    changeValueByArrowKey(etStartT, false, function () { __syncTFromEdits(); refreshPreviewIfNeeded(); });

    var rowEnd = pnlPosition.add('group');
    rowEnd.orientation = 'row';
    rowEnd.alignChildren = ['left', 'center'];
    var cbEndT = rowEnd.add('checkbox', undefined, '');
    cbEndT.value = false;
    var stEnd = rowEnd.add('statictext', undefined, L('endPos'));
    stEnd.preferredSize.width = 60;
    var etEndT = rowEnd.add('edittext', undefined, '1.0');
    etEndT.characters = 5;
    var slEndT = rowEnd.add('slider', undefined, 100, 0, 500); // 0.0 - 5.0
    slEndT.preferredSize.width = 180;
    // Arrow-key support for endT
    changeValueByArrowKey(etEndT, false, function () { __syncTFromEdits(); refreshPreviewIfNeeded(); });

    // Baseline shift panel  (moved to FULL WIDTH)
    var pnlTextAdjust = wideCol.add('panel', undefined, L('panelTextAdjust'));
    setupPanel(pnlTextAdjust);

    // 行揃え（テキスト調整パネル先頭の行）/ Alignment (top row of Text Adjust panel)
    var rowAlign = pnlTextAdjust.add('group');
    rowAlign.orientation = 'row';
    rowAlign.alignChildren = ['left', 'center'];

    var stAlign = rowAlign.add('statictext', undefined, labelText('panelAlign'));
    stAlign.preferredSize.width = 95;

    var rbAlignLeft = rowAlign.add('radiobutton', undefined, L('alignLeft'));
    var rbAlignCenter = rowAlign.add('radiobutton', undefined, L('alignCenter'));
    var rbAlignRight = rowAlign.add('radiobutton', undefined, L('alignRight'));
    var rbAlignFullJustify = rowAlign.add('radiobutton', undefined, L('alignFullJustify'));

    // Default: 中央
    rbAlignCenter.value = true;

    var rowBaseShift = pnlTextAdjust.add('group');
    rowBaseShift.orientation = 'row';
    rowBaseShift.alignChildren = ['left', 'center'];

    var stBaseShift = rowBaseShift.add('statictext', undefined, labelText('baseShift'));
    stBaseShift.preferredSize.width = 95;

    var etBaseShift = rowBaseShift.add('edittext', undefined, '0.0');
    etBaseShift.characters = 6;
    // Slider uses 0.1pt steps, range will be updated dynamically to ±fontSize (in 0.1pt units)
    var slBaseShift = rowBaseShift.add('slider', undefined, 0, -1000, 1000);
    slBaseShift.preferredSize.width = 180;
    // Arrow-key support for baseline shift
    changeValueByArrowKey(etBaseShift, true, function () { __syncBSFromEdit(); refreshPreviewIfNeeded(); });

    // Tracking panel row
    var rowTracking = pnlTextAdjust.add('group');
    rowTracking.orientation = 'row';
    rowTracking.alignChildren = ['left', 'center'];

    var stTracking = rowTracking.add('statictext', undefined, labelText('tracking'));
    stTracking.preferredSize.width = 95;

    var etTracking = rowTracking.add('edittext', undefined, '0');
    etTracking.characters = 6;
    var slTracking = rowTracking.add('slider', undefined, 0, -100, 500);
    slTracking.preferredSize.width = 180;
    // Arrow-key support for tracking
    changeValueByArrowKey(etTracking, true, function () { __syncTrkFromEdit(); refreshPreviewIfNeeded(); });

    // Font size panel row
    var rowFontSize = pnlTextAdjust.add('group');
    rowFontSize.orientation = 'row';
    rowFontSize.alignChildren = ['left', 'center'];

    var stFontSize = rowFontSize.add('statictext', undefined, labelText('fontSize'));
    stFontSize.preferredSize.width = 95;

    var etFontSize = rowFontSize.add('edittext', undefined, '0.0');
    etFontSize.characters = 6;
    // Slider uses 0.1pt steps, range will be updated dynamically to ±fontSize (in 0.1pt units)
    var slFontSize = rowFontSize.add('slider', undefined, 0, -1000, 1000);
    slFontSize.preferredSize.width = 180;
    // Arrow-key support for font size
    changeValueByArrowKey(etFontSize, true, function () { __syncFSFromEdit(); refreshPreviewIfNeeded(); });

    // Footer (Preview independent) + Buttons (OK on the right)
    var footer = dlg.add('group');
    footer.orientation = 'row';
    footer.alignChildren = ['fill', 'center'];
    footer.alignment = ['fill', 'top'];

    var leftFooter = footer.add('group');
    leftFooter.orientation = 'row';
    leftFooter.alignChildren = ['left', 'center'];
    leftFooter.alignment = ['left', 'center'];

    var cbPreview = leftFooter.add('checkbox', undefined, L('preview'));
    cbPreview.value = true;

    var rightFooter = footer.add('group');
    rightFooter.orientation = 'row';
    rightFooter.alignChildren = ['right', 'center'];
    rightFooter.alignment = ['right', 'center'];


    var btnCancel = rightFooter.add('button', undefined, L('cancel'));
    var btnOk = rightFooter.add('button', undefined, L('ok'), { name: 'ok' });

    /* プレビュー（Undoなし） / Preview (no undo) */
    var __previewTempItems = [];        // items created during preview (TextFrame etc.)
    var __previewHiddenOriginals = [];  // original items hidden during preview

    var __previewPathStates = [];       // { item: PathItem, stroked: Boolean, filled: Boolean, strokeWidth, strokeColor, opacity }

    /* プレビューで作った一時オブジェクトを削除し元の状態へ戻す / Remove preview temp items and restore originals */
    function __preview_clear() {
        // Remove temp items
        for (var i = __previewTempItems.length - 1; i >= 0; i--) {
            try { __previewTempItems[i].remove(); } catch (_) { }
        }
        __previewTempItems = [];

        // Restore originals visibility
        for (var j = __previewHiddenOriginals.length - 1; j >= 0; j--) {
            try { __previewHiddenOriginals[j].hidden = false; } catch (_) { }
        }
        // Restore path appearance changed during preview
        for (var p = __previewPathStates.length - 1; p >= 0; p--) {
            try {
                var st = __previewPathStates[p];
                st.item.stroked = st.stroked;
                st.item.filled = st.filled;
                st.item.strokeWidth = st.strokeWidth;
                st.item.strokeColor = st.strokeColor;
                st.item.opacity = st.opacity;
            } catch (_) { }
        }
        __previewPathStates = [];

        __previewHiddenOriginals = [];
        // Removed redundant redraw here to reduce flicker.
    }

    /* プレビュー復元用にパスの元の見た目を記録 / Record a path's original appearance for preview restore */
    function __preview_recordPathState(pathItem) {
        try {
            if (!pathItem) return;
            for (var i = 0; i < __previewPathStates.length; i++) {
                if (__previewPathStates[i].item === pathItem) return;
            }
            __previewPathStates.push({
                item: pathItem,
                stroked: pathItem.stroked,
                filled: pathItem.filled,
                strokeWidth: pathItem.strokeWidth,
                strokeColor: pathItem.strokeColor,
                opacity: pathItem.opacity
            });
        } catch (_) { }
    }

    /* CompoundPath も展開して各 PathItem に処理を適用 / Run a function on each underlying PathItem (handles CompoundPathItem) */
    function __forEachPathItem(pathItem, styleFn) {
        try {
            if (!pathItem || !styleFn) return;
            if (pathItem.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < pathItem.pathItems.length; cp++) {
                    try { styleFn(pathItem.pathItems[cp]); } catch (_) { }
                }
            } else {
                try { styleFn(pathItem); } catch (_) { }
            }
        } catch (_) { }
    }

    /* スミ100%の CMYKColor を生成 / Create a 100%-black CMYKColor */
    function makeBlackCMYK() {
        var black = new CMYKColor();
        black.cyan = 0;
        black.magenta = 0;
        black.yellow = 0;
        black.black = 100;
        return black;
    }

    /* パスを見えるガイド（塗りなし・黒1pt・不透明度50%）に整形 / Style a path as a visible guide (no fill, black 1pt stroke, 50% opacity) */
    function __styleVisiblePath(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeWidth = 1;
        pathItem.strokeColor = makeBlackCMYK();
        pathItem.opacity = 50;
    }

    /* パスを非表示（塗り・線なし）に整形 / Style a path as invisible (no fill, no stroke) */
    function __styleInvisiblePath(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = false;
        pathItem.strokeWidth = 0;
    }

    /* プレビュー：元の見た目を記録してから見えるガイドにする / Preview: record original style, then make the path a visible guide */
    function __preview_applyPathStyle(basePath) {
        __forEachPathItem(basePath, function (pathItem) {
            __preview_recordPathState(pathItem);
            __styleVisiblePath(pathItem);
        });
    }

    /* 実行：パスを見えるガイドのまま残す（復元不要）/ Execute: keep the path visible as a guide (no restore needed) */
    function __applyExecutePathStyle(pathItem) {
        __forEachPathItem(pathItem, __styleVisiblePath);
    }

    /* 生成パスを非表示ガイド（線幅0）にする / Make a generated path an invisible guide (zero stroke) */
    function __applyInvisiblePathStyle(pathItem) {
        __forEachPathItem(pathItem, __styleInvisiblePath);
    }

    /* 元オブジェクトをプレビュー中だけ非表示にする / Hide an original item during preview */
    function __preview_hideOriginal(item) {
        try {
            if (!item) return;
            // Avoid duplicates
            for (var k = 0; k < __previewHiddenOriginals.length; k++) {
                if (__previewHiddenOriginals[k] === item) return;
            }
            item.hidden = true;
            __previewHiddenOriginals.push(item);
        } catch (_) { }
    }

    /* 現在の設定でプレビューを再生成（Undo なし）/ Rebuild the preview with the current settings (no undo) */
    function __preview_apply() {
        __preview_clear();
        // Restore base selection so preview stays stable even after selection changes
        try { doc.selection = __baseSelection; } catch (_) { }
        __refreshInputsFromSelection();

        if (!targetItems || targetItems.length === 0) {
            cbPreview.value = false;
            return;
        }

        // Apply preview conversion (does NOT permanently modify paths)
        if (rbSplitTextAndPath.value) {
            splitPathTextAndPath(false, true, true);
        } else if (rbSplitTextAndPathNoFormat.value) {
            splitPathTextAndPath(false, true, false);
        } else if (rbToPathText.value) {
            mainProcess(false, true);
        } else if (rbGenArcPath.value) {
            mainProcessGenerateArcPath(false, true);
        } else if (rbGenCircle.value) {
            mainProcessGenerateCircle(false, true);
        }
        try { app.redraw(); } catch (_) { }
    }

    /* プレビューを適用 / Apply the preview */
    function applyPreview() {
        __preview_apply();
    }

    // Live preview handlers
    cbPreview.onClick = function () {
        if (cbPreview.value) {
            applyPreview();
        } else {
            __preview_clear();
            try { app.redraw(); } catch (_) { }
        }
    };

    /* プレビュー ON のときだけ再描画 / Refresh the preview only when it is ON */
    function refreshPreviewIfNeeded() {
        if (cbPreview.value) {
            applyPreview();
        }
    }

    /* アーチモード時の既定値（中央揃え・虹・文字サイズでフィット）を適用 / Apply arc-mode presets (center align, rainbow effect, fit by font size) */
    function applyArcModePresets() {
        try {
            rbAlignCenter.value = true;
            rbEffectRainbow.value = true;
            rbFitWidthFontSize.value = true;
        } catch (_) { }
    }

    rbToPathText.onClick = function () {
        __clearSplitRadios();
        try { rbFitWidthNone.value = true; } catch (_) { }
        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    rbGenCircle.onClick = function () {
        __clearSplitRadios();
        try { rbFitWidthNone.value = true; } catch (_) { }
        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    rbGenArcPath.onClick = function () {
        __clearSplitRadios();
        __updatePanelsByMode();
        applyArcModePresets();
        refreshPreviewIfNeeded();
    };
    rbSplitTextAndPath.onClick = function () {
        __clearProcessRadios();
        try { rbFitWidthNone.value = true; } catch (_) { }
        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    rbSplitTextAndPathNoFormat.onClick = function () {
        __clearProcessRadios();
        try { rbFitWidthNone.value = true; } catch (_) { }
        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    /* モードに応じてパネルの有効/無効を切り替え / Enable or disable panels depending on the current mode */
    function __updatePanelsByMode() {
        var isSplitMode = false;
        var isCircleMode = false;
        try {
            isSplitMode = (rbSplitTextAndPath && rbSplitTextAndPath.value) ||
                (rbSplitTextAndPathNoFormat && rbSplitTextAndPathNoFormat.value);
            isCircleMode = (rbGenCircle && rbGenCircle.value);
        } catch (_) {
            isSplitMode = false;
            isCircleMode = false;
        }

        // 各パネルの有効/無効をまとめて設定（UI 構築後のみ呼ばれるため try は1つで足りる）
        var en = !isSplitMode;
        try {
            // 「処理」「分離」は常に有効
            pnlProcess.enabled = true;
            pnlSplit.enabled = true;
            // Splitモード時はそれ以外をディム
            pnlOption.enabled = en;
            pnlEffect.enabled = en;
            pnlTextAdjust.enabled = en;
            pnlPosition.enabled = en;
            // パス幅・アーチ方向・まるみは正円モードでも無効（オープンパスのみ対象）
            pnlFitWidth.enabled = en && !isCircleMode;
            grpArcDir.enabled = en && !isCircleMode;
            grpArcRoundness.enabled = en && !isCircleMode;
        } catch (_) { }
    }

    cbReverse.onClick = refreshPreviewIfNeeded;
    rbEffectRainbow.onClick = refreshPreviewIfNeeded;
    rbEffectDistort.onClick = refreshPreviewIfNeeded;
    rbEffectRibbon.onClick = refreshPreviewIfNeeded;
    rbEffectStep.onClick = refreshPreviewIfNeeded;
    rbEffectGravity.onClick = refreshPreviewIfNeeded;

    rbAlignLeft.onClick = refreshPreviewIfNeeded;
    rbAlignCenter.onClick = refreshPreviewIfNeeded;
    rbAlignRight.onClick = refreshPreviewIfNeeded;
    rbAlignFullJustify.onClick = refreshPreviewIfNeeded;
    // Arc direction radios: hook preview refresh
    rbArcUp.onClick = refreshPreviewIfNeeded;
    rbArcDown.onClick = refreshPreviewIfNeeded;
    // Fit-to-path-width radios: hook preview refresh
    rbFitWidthNone.onClick = refreshPreviewIfNeeded;
    rbFitWidthFontSize.onClick = refreshPreviewIfNeeded;
    rbFitWidthTracking.onClick = refreshPreviewIfNeeded;

    // Arc roundness slider: refresh preview on release
    slArcRoundness.onChange = refreshPreviewIfNeeded;
    // Sync baseline shift UI (edittext <-> slider)
    // - slider unit: 0.1pt (value = delta * 10)
    // - slider range: ±fontSize (pt) based on current selection
    var __bsSyncLock = false;

    /* 数値を小数1桁の文字列へ整形（0.1pt 単位）/ Format a number to a one-decimal string (0.1pt steps) */
    function formatOneDecimal(value) {
        return (Math.round(value * 10) / 10).toFixed(1);
    }

    /* 基準フォントサイズ（先頭ターゲット）を pt で取得 / Get the reference font size in pt (first target) */
    function getReferenceFontSizePt() {
        try {
            var firstTarget = (targetItems && targetItems.length > 0) ? targetItems[0] : null;
            if (firstTarget && firstTarget.typename === 'TextFrame') {
                // Prefer first textRange, fall back to whole range
                if (firstTarget.textRanges && firstTarget.textRanges.length > 0) {
                    var size = firstTarget.textRanges[0].characterAttributes.size;
                    if (size && !isNaN(size)) return size;
                }
                var fallbackSize = firstTarget.textRange.characterAttributes.size;
                if (fallbackSize && !isNaN(fallbackSize)) return fallbackSize;
            }
        } catch (_) { }
        return 100; // fallback
    }

    /* 増減スライダーの範囲を ±フォントサイズ（0.1pt 単位）に更新 / Update a delta slider range to ±font size (0.1pt steps) */
    function updateDeltaSliderRange(slider) {
        try {
            var limit = Math.max(1, Math.round(getReferenceFontSizePt() * 10));
            slider.minvalue = -limit;
            slider.maxvalue = limit;
        } catch (_) { }
    }

    /* 入力欄からベースラインのスライダーへ同期 / Sync the baseline-shift slider from the edit field */
    function __syncBSFromEdit() {
        if (__bsSyncLock) return;
        __bsSyncLock = true;
        try {
            updateDeltaSliderRange(slBaseShift);

            var v = __parseNumber(etBaseShift.text, 0);
            if (isNaN(v)) v = 0;

            // clamp to slider range
            var minPt = slBaseShift.minvalue / 10.0;
            var maxPt = slBaseShift.maxvalue / 10.0;
            v = __clamp(v, minPt, maxPt);

            etBaseShift.text = formatOneDecimal(v);
            try { slBaseShift.value = Math.round(v * 10); } catch (_) { }
        } catch (_) { }
        __bsSyncLock = false;
    }

    /* スライダーからベースラインの入力欄へ同期 / Sync the baseline-shift edit field from the slider */
    function __syncBSFromSlider() {
        if (__bsSyncLock) return;
        __bsSyncLock = true;
        try {
            updateDeltaSliderRange(slBaseShift);
            var v = (slBaseShift.value / 10.0);
            etBaseShift.text = formatOneDecimal(v);
        } catch (_) { }
        __bsSyncLock = false;
    }

    etBaseShift.onChanging = function () { __syncBSFromEdit(); refreshPreviewIfNeeded(); };
    slBaseShift.onChanging = function () { __syncBSFromSlider(); };
    slBaseShift.onChange = function () { __syncBSFromSlider(); refreshPreviewIfNeeded(); };
    __syncBSFromEdit();

    // Sync tracking UI (edittext <-> slider)
    var __trkSyncLock = false;

    /* 入力欄からトラッキングのスライダーへ同期 / Sync the tracking slider from the edit field */
    function __syncTrkFromEdit() {
        if (__trkSyncLock) return;
        __trkSyncLock = true;
        try {
            var v = Math.round(__parseNumber(etTracking.text, 0));
            if (v > 500) v = 500;
            if (v < -100) v = -100;
            etTracking.text = String(v);
            try { slTracking.value = v; } catch (_) { }
        } catch (_) { }
        __trkSyncLock = false;
    }

    /* スライダーからトラッキングの入力欄へ同期 / Sync the tracking edit field from the slider */
    function __syncTrkFromSlider() {
        if (__trkSyncLock) return;
        __trkSyncLock = true;
        try {
            var v = Math.round(slTracking.value);
            etTracking.text = String(v);
        } catch (_) { }
        __trkSyncLock = false;
    }

    etTracking.onChanging = function () { __syncTrkFromEdit(); refreshPreviewIfNeeded(); };
    slTracking.onChanging = function () { __syncTrkFromSlider(); };
    slTracking.onChange = function () { __syncTrkFromSlider(); refreshPreviewIfNeeded(); };
    __syncTrkFromEdit();

    // Sync font size delta UI (edittext <-> slider)
    // - slider unit: 0.1pt (value = delta * 10)
    // - slider range: ±fontSize (pt) based on current selection
    var __fsSyncLock = false;

    /* 入力欄から文字サイズのスライダーへ同期 / Sync the font-size slider from the edit field */
    function __syncFSFromEdit() {
        if (__fsSyncLock) return;
        __fsSyncLock = true;
        try {
            updateDeltaSliderRange(slFontSize);

            var v = __parseNumber(etFontSize.text, 0);
            if (isNaN(v)) v = 0;

            var minPt = slFontSize.minvalue / 10.0;
            var maxPt = slFontSize.maxvalue / 10.0;
            v = __clamp(v, minPt, maxPt);

            etFontSize.text = formatOneDecimal(v);
            try { slFontSize.value = Math.round(v * 10); } catch (_) { }
        } catch (_) { }
        __fsSyncLock = false;
    }

    /* スライダーから文字サイズの入力欄へ同期 / Sync the font-size edit field from the slider */
    function __syncFSFromSlider() {
        if (__fsSyncLock) return;
        __fsSyncLock = true;
        try {
            updateDeltaSliderRange(slFontSize);
            var v = (slFontSize.value / 10.0);
            etFontSize.text = formatOneDecimal(v);
        } catch (_) { }
        __fsSyncLock = false;
    }

    etFontSize.onChanging = function () { __syncFSFromEdit(); refreshPreviewIfNeeded(); };
    slFontSize.onChanging = function () { __syncFSFromSlider(); };
    slFontSize.onChange = function () { __syncFSFromSlider(); refreshPreviewIfNeeded(); };
    __syncFSFromEdit();
    // Sync t-value UI (edittext <-> slider)
    var __tSyncLock = false;

    /* チェックボックスに応じて開始/終了位置の入力を有効化 / Enable start/end position inputs per their checkboxes */
    function __updateTEnableUI() {
        try {
            var sOn = (cbStartT && cbStartT.value);
            var eOn = (cbEndT && cbEndT.value);
            etStartT.enabled = sOn;
            slStartT.enabled = sOn;
            etEndT.enabled = eOn;
            slEndT.enabled = eOn;
        } catch (_) { }
    }

    /* 入力欄から開始/終了位置のスライダーへ同期 / Sync start/end position sliders from the edit fields */
    function __syncTFromEdits() {
        if (__tSyncLock) return;
        __tSyncLock = true;
        try {
            __updateTEnableUI();

            var sOn = (cbStartT && cbStartT.value);
            var eOn = (cbEndT && cbEndT.value);

            var s = sOn ? __parseTValue(etStartT.text, 0.0) : 0.0;
            var e = eOn ? __parseTValue(etEndT.text, 1.0) : 1.0;

            // clamp to UI ranges
            s = __clamp(s, 0.0, 4.0);
            e = __clamp(e, 0.0, 5.0);

            if (!sOn) s = 0.0;
            if (!eOn) e = 1.0;

            etStartT.text = formatOneDecimal(s);
            etEndT.text = formatOneDecimal(e);

            try { slStartT.value = Math.round(s * 100); } catch (_) { }
            try { slEndT.value = Math.round(e * 100); } catch (_) { }
        } catch (_) { }
        __tSyncLock = false;
    }

    /* スライダーから開始/終了位置の入力欄へ同期 / Sync start/end position edit fields from the sliders */
    function __syncTFromSliders() {
        if (__tSyncLock) return;
        __tSyncLock = true;
        try {
            __updateTEnableUI();

            var sOn = (cbStartT && cbStartT.value);
            var eOn = (cbEndT && cbEndT.value);

            var s = sOn ? (slStartT.value / 100.0) : 0.0;
            var e = eOn ? (slEndT.value / 100.0) : 1.0;

            // clamp to UI ranges
            s = __clamp(s, 0.0, 4.0);
            e = __clamp(e, 0.0, 5.0);

            if (!sOn) s = 0.0;
            if (!eOn) e = 1.0;

            etStartT.text = formatOneDecimal(s);
            etEndT.text = formatOneDecimal(e);
        } catch (_) { }
        __tSyncLock = false;
    }

    etStartT.onChanging = function () { __syncTFromEdits(); refreshPreviewIfNeeded(); };
    etEndT.onChanging = function () { __syncTFromEdits(); refreshPreviewIfNeeded(); };
    slStartT.onChanging = function () { __syncTFromSliders(); };
    slStartT.onChange = function () { __syncTFromSliders(); refreshPreviewIfNeeded(); };
    slEndT.onChanging = function () { __syncTFromSliders(); };
    slEndT.onChange = function () { __syncTFromSliders(); refreshPreviewIfNeeded(); };

    cbStartT.onClick = function () { __syncTFromEdits(); refreshPreviewIfNeeded(); };
    cbEndT.onClick = function () { __syncTFromEdits(); refreshPreviewIfNeeded(); };

    __syncTFromEdits();

    // Buttons
    btnCancel.onClick = function () {
        __preview_clear();
        dlg.close(0);
    };

    btnOk.onClick = function () {
        // If preview is ON, clear it first so we don't stack temporary objects.
        if (cbPreview.value) {
            __preview_clear();
        }

        if (rbSplitTextAndPath.value) {
            splitPathTextAndPath(true, false, true);
        } else if (rbSplitTextAndPathNoFormat.value) {
            splitPathTextAndPath(true, false, false);
        } else if (rbToPathText.value) {
            mainProcess(true, false);
        } else if (rbGenArcPath.value) {
            mainProcessGenerateArcPath(true, false);
        } else if (rbGenCircle.value) {
            mainProcessGenerateCircle(true, false);
        }

        dlg.close(1);
    };

    // If arc mode ended up active at startup (auto-switched), apply its presets
    // now that the whole UI exists.
    if (rbGenArcPath.value) applyArcModePresets();

    // Set initial panel enabled state
    __updatePanelsByMode();

    // Auto-apply preview once on open (no undo)
    try {
        if (cbPreview.value) {
            __preview_apply();
            try { app.redraw(); } catch (_) { }
        }
    } catch (_) { }

    var res = dlg.show();
    if (res !== 1) return false;
    return true;

    // Remove old preview hook for the checkbox
    // cbFixOverset.onClick = refreshPreviewIfNeeded;

    /* メニューコマンドでパス上文字に効果（虹・歪み等）を適用 / Apply a path-text effect (rainbow, skew, etc.) via a menu command */
    function applyPathTextEffect(textFrame, previewMode) {
        var effectCommand = getSelectedEffectCommand();
        if (!effectCommand) return;

        var previousSelection = null;
        try { previousSelection = doc.selection; } catch (_) { previousSelection = null; }

        try {
            try { doc.selection = []; } catch (_) { }
            try { textFrame.selected = true; } catch (_) { }
            app.executeMenuCommand(effectCommand);
        } catch (_) {
            // ignore
        }

        // Always restore selection (safer)
        try { doc.selection = previousSelection; } catch (_) { }
    }

    /* 選択中の効果に対応するメニューコマンド名を返す / Return the menu command name for the selected effect */
    function getSelectedEffectCommand() {
        if (rbEffectRainbow.value) return 'Rainbow';
        if (rbEffectDistort.value) return 'Skew';
        if (rbEffectRibbon.value) return '3D ribbon';
        if (rbEffectStep.value) return 'Stair Step';
        if (rbEffectGravity.value) return 'Gravity';
        return null;
    }

    /* 行揃え（段落揃え）をテキストフレームへ適用（複製後に呼ぶ）/ Apply paragraph justification to a text frame (call after content is duplicated) */
    function applyJustificationToTextFrame(textFrame) {
        try {
            if (!textFrame) return;

            var justification = Justification.FULLJUSTIFYLASTLINELEFT;
            if (rbAlignFullJustify && rbAlignFullJustify.value) {
                justification = Justification.FULLJUSTIFY;
            } else if (rbAlignRight && rbAlignRight.value) {
                justification = Justification.RIGHT;
            } else if (rbAlignCenter && rbAlignCenter.value) {
                justification = Justification.CENTER;
            } else {
                justification = Justification.FULLJUSTIFYLASTLINELEFT;
            }

            // Apply to all paragraphs if available
            try {
                if (textFrame.paragraphs && textFrame.paragraphs.length > 0) {
                    for (var i = 0; i < textFrame.paragraphs.length; i++) {
                        try { textFrame.paragraphs[i].paragraphAttributes.justification = justification; } catch (_) { }
                    }
                }
            } catch (_) { }

            // Fallback apply to whole range
            try { textFrame.textRange.paragraphAttributes.justification = justification; } catch (_) { }
        } catch (_) { }
    }

    /* 値を min〜max の範囲に収める / Clamp a number to the min–max range */
    function __clamp(n, min, max) {
        if (n < min) return min;
        if (n > max) return max;
        return n;
    }

    /* 文字列を t 値（0〜5）として解釈 / Parse a string as a t-value (0–5) */
    function __parseTValue(str, fallback) {
        var n = Number(str);
        if (isNaN(n)) return fallback;
        return __clamp(n, 0, 5);
    }

    /* UI の開始/終了位置をパス上文字へ適用 / Apply the start/end position from the UI to path text */
    function applyStartEndTValue(textFrame) {
        try {
            if (!textFrame) return;

            if (cbStartT && cbStartT.value) {
                var startValue = __parseTValue(etStartT.text, 0.0);
                textFrame.startTValue = startValue;
            }

            if (cbEndT && cbEndT.value) {
                var endValue = __parseTValue(etEndT.text, 1.0);
                textFrame.endTValue = endValue;
            }

        } catch (_) { }
    }

    /* 文字列を数値へ変換（失敗時は既定値）/ Parse a string as a number (fallback when invalid) */
    function __parseNumber(str, fallback) {
        var n = Number(str);
        if (isNaN(n)) return fallback;
        return n;
    }

    /* テキストフレームの textRange を配列で取得（無ければ textRange 単体）/ Collect a frame's textRanges as an array (falls back to its single textRange) */
    function collectTextRanges(textFrame) {
        var ranges = [];
        try {
            if (textFrame.textRanges && textFrame.textRanges.length > 0) {
                for (var i = 0; i < textFrame.textRanges.length; i++) ranges.push(textFrame.textRanges[i]);
            }
        } catch (_) { }
        if (ranges.length === 0) {
            try { if (textFrame.textRange) ranges = [textFrame.textRange]; } catch (_) { ranges = []; }
        }
        return ranges;
    }

    /* 各 textRange の文字属性に delta を加算（clampMin 指定時は下限を適用）/ Add a delta to a character attribute of every textRange (optional lower clamp) */
    function addDeltaToTextRanges(textFrame, attributeName, delta, clampMin) {
        if (!textFrame || !delta) return;
        var ranges = collectTextRanges(textFrame);
        for (var r = 0; r < ranges.length; r++) {
            try {
                var attributes = ranges[r].characterAttributes;
                var nextValue = attributes[attributeName] + delta;
                if (typeof clampMin === 'number' && nextValue < clampMin) nextValue = clampMin;
                attributes[attributeName] = nextValue;
            } catch (_) { }
        }
    }

    /* 各 textRange の既存ベースラインに UI 指定値を加算 / Add the UI delta to each textRange's baseline shift */
    function applyBaselineShiftDelta(textFrame) {
        addDeltaToTextRanges(textFrame, 'baselineShift', __parseNumber(etBaseShift.text, 0));
    }

    /* 各 textRange の既存トラッキングに UI 指定値を加算 / Add the UI delta to each textRange's tracking */
    function applyTrackingDelta(textFrame) {
        addDeltaToTextRanges(textFrame, 'tracking', Math.round(__parseNumber(etTracking.text, 0)));
    }

    /* 各 textRange の既存文字サイズに UI 指定値を加算（最小 0.1pt）/ Add the UI delta to each textRange's font size (min 0.1pt) */
    function applyFontSizeDelta(textFrame) {
        addDeltaToTextRanges(textFrame, 'size', __parseNumber(etFontSize.text, 0), 0.1);
    }

    /* 指定パス上にパス上文字を作成し、元テキストの直前へ配置 / Create path text on the given path, placed just before the original */
    function createPathTextFrame(textPath, originalText, currentLayer, previewMode) {
        var textOnAPath = currentLayer.textFrames.pathText(textPath);
        // 重ね順を保持（他オブジェクトの背面に回り込まないように）/ Keep stacking order
        try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (_) { }
        if (previewMode) __previewTempItems.push(textOnAPath);
        return textOnAPath;
    }

    /* 段落をすべて中央揃えにする / Set every paragraph to center justification */
    function applyCenterJustification(textFrame) {
        try {
            if (textFrame.paragraphs && textFrame.paragraphs.length > 0) {
                for (var p = 0; p < textFrame.paragraphs.length; p++) {
                    try { textFrame.paragraphs[p].paragraphAttributes.justification = Justification.CENTER; } catch (_) { }
                }
            }
        } catch (_) { }
        try { textFrame.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (_) { }
    }

    /* パス上文字の共通仕上げ：内側配置・効果・内容複製・各種調整を適用し元テキストを除去 / Shared finishing of path text (inside placement, effect, content copy, adjustments, remove original) */
    function decoratePathText(textOnAPath, originalText, previewMode, options) {
        options = options || {};

        // 内側配置 / Inside placement
        if (cbReverse.value) {
            try { textOnAPath.textPath.polarity = PolarityValues.NEGATIVE; } catch (_) { }
        }

        applyPathTextEffect(textOnAPath, previewMode);
        applyStartEndTValue(textOnAPath); // 位置（start/end）

        // 正円＋内側：開始/終了位置を +0.5 ずらして見かけの開始位置を合わせる
        if (options.circleInsideShift) {
            try {
                var shiftAmount = 0.5;
                // Clamp only to the global end cap (0..5); start may exceed 1 for circles.
                textOnAPath.startTValue = __clamp(textOnAPath.startTValue + shiftAmount, 0.0, 5.0);
                textOnAPath.endTValue = __clamp(textOnAPath.endTValue + shiftAmount, 0.0, 5.0);
            } catch (_) { }
        }

        // テキスト内容を元から複製 / Copy text content from the original
        for (var i = 0; i < originalText.textRanges.length; i++) {
            originalText.textRanges[i].duplicate(textOnAPath);
        }

        applyFontSizeDelta(textOnAPath); // 文字サイズ（既存値 + 指定値）

        // 行揃え（duplicate 後に適用しないと上書きされる）/ Justification (apply after content is copied)
        if (options.forceCenter) {
            applyCenterJustification(textOnAPath);
        } else {
            applyJustificationToTextFrame(textOnAPath);
        }

        applyBaselineShiftDelta(textOnAPath); // ベースラインシフト（既存値 + 指定値）
        applyTrackingDelta(textOnAPath);      // トラッキング（既存値 + 指定値）

        // 元テキストを隠す/削除 / Hide or remove the original text
        if (previewMode) {
            __preview_hideOriginal(originalText);
        } else {
            try { originalText.remove(); } catch (_) { }
        }

        // 実行時は生成したパス上文字を選択 / Select the created path text on execute
        if (!previewMode) {
            try { textOnAPath.selected = true; } catch (_) { }
        }
    }

    /* メイン処理：選択パスに沿ってパス上文字を作成 / Main process: lay text along the selected path */
    function mainProcess(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var __createdTexts = [];

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentLayer = originalText.layer;

            // この文字に使う元パスを決定 / Resolve the base path for this text
            var basePath = resolveBasePathForText(selectedPaths, targetItems.length, j, originalText);
            if (!basePath) {
                // Validation 済みのため通常は来ないが、防御的に
                if (showAlerts) alert(L('alertNeedPath'));
                return;
            }

            if (!previewMode) {
                makeOriginalPathInvisible(basePath);
            } else {
                // Preview: show the original selected path with no fill/stroke (restored on preview clear)
                __preview_applyPathStyle(basePath);
            }

            var textPath = duplicatePathForText(basePath, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(L('alertDupPathFail'));
                continue;
            }
            if (previewMode) {
                __previewTempItems.push(textPath);
            } else {
                __applyExecutePathStyle(textPath); // Execute: keep the duplicated path visible
            }

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);
            decoratePathText(textOnAPath, originalText, previewMode);
            __createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(__createdTexts);
    }

    /* ハンドル・ポイント種別を保持してパスを複製 / Duplicate a path preserving handles and point types */
    function duplicatePathWithHandles(originalPath, targetLayer) {
        try {
            if (!originalPath) return null;
            var container = (targetLayer && targetLayer.pathItems) ? targetLayer.pathItems : doc.pathItems;
            var newPath = container.add();

            for (var k = 0; k < originalPath.pathPoints.length; k++) {
                var origPt = originalPath.pathPoints[k];
                var newPt = newPath.pathPoints.add();
                newPt.anchor = origPt.anchor;
                newPt.leftDirection = origPt.leftDirection;
                newPt.rightDirection = origPt.rightDirection;
                newPt.pointType = origPt.pointType;
            }
            newPath.closed = originalPath.closed;
            return newPath;
        } catch (_) {
            return null;
        }
    }

    /* 1件のパス上文字をパス＋ポイント文字に分離（処理したら true）/ Split one path text into a path + point text (returns true if processed) */
    function splitOnePathText(pathText, previewMode, keepFormatting) {
        if (!pathText || pathText.typename !== 'TextFrame') return false;

        var isPathText = false;
        try { isPathText = (pathText.kind === TextType.PATHTEXT); } catch (_) { isPathText = false; }
        if (!isPathText) return false;

        var originalPath = null;
        try { originalPath = pathText.textPath; } catch (_) { originalPath = null; }
        if (!originalPath) return false;

        var curLayer = null;
        try { curLayer = pathText.layer; } catch (_) { curLayer = null; }

        // 1) 「パスを削除」OFF ならハンドルを保持してパスを複製
        var deletePath = false;
        try { deletePath = (cbSplitDeletePath && cbSplitDeletePath.value); } catch (_) { deletePath = false; }

        if (!deletePath) {
            var newPath = duplicatePathWithHandles(originalPath, curLayer);
            if (newPath) {
                // 塗りなし・スミ1pt / no fill, black 1pt stroke
                try {
                    newPath.filled = false;
                    newPath.stroked = true;
                    newPath.strokeColor = makeBlackCMYK();
                    newPath.strokeWidth = 1;
                } catch (_) { }
                if (previewMode) __previewTempItems.push(newPath);
            }
        }

        // 2) 元の開始アンカー付近にポイント文字を作成
        var newText = null;
        try {
            newText = (curLayer ? curLayer.textFrames.add() : doc.textFrames.add());
        } catch (_) {
            newText = null;
        }

        if (newText) {
            try {
                var anchorPoint = originalPath.pathPoints[0].anchor;
                newText.position = [anchorPoint[0], anchorPoint[1]];
            } catch (_) { }

            // 内容コピー（書式を保持する/しない）/ Copy content (keep formatting or not)
            if (keepFormatting) {
                try { pathText.textRange.duplicate(newText); } catch (_) { }
            } else {
                try { newText.contents = pathText.contents; } catch (_) { }
            }

            applyFontSizeDelta(newText);      // 文字サイズ（既存値 + 指定値）
            applyBaselineShiftDelta(newText); // ベースラインシフト（既存値 + 指定値）
            applyTrackingDelta(newText);      // トラッキング（既存値 + 指定値）

            // 書式保持時のみ塗り/線カラーを引き継ぐ
            if (keepFormatting) {
                for (var c = 0; c < pathText.textRanges.length; c++) {
                    try {
                        var srcAttr = pathText.textRanges[c].characterAttributes;
                        var dstAttr = newText.textRanges[c].characterAttributes;
                        dstAttr.fillColor = srcAttr.fillColor;
                        dstAttr.strokeColor = srcAttr.strokeColor;
                    } catch (_) { }
                }
            }

            // 行揃え（duplicate 後に適用）と重ね順の保持
            try { applyJustificationToTextFrame(newText); } catch (_) { }
            try { newText.move(pathText, ElementPlacement.PLACEBEFORE); } catch (_) { }

            if (previewMode) __previewTempItems.push(newText);
        }

        // 3) ポイント文字が作れたときだけ元のパス上文字を隠す/削除（失敗時は元を残す）
        if (newText) {
            if (previewMode) {
                __preview_hideOriginal(pathText);
            } else {
                try { pathText.remove(); } catch (_) { }
            }
        }
        return true;
    }

    /* パス上文字をパス（可視）とポイント文字に分離 / Split path text into a visible path and point text */
    function splitPathTextAndPath(showAlerts, previewMode, keepFormatting) {
        if (typeof keepFormatting === 'undefined') keepFormatting = true;
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;

        var didAny = false;
        for (var j = targetItems.length - 1; j >= 0; j--) {
            if (splitOnePathText(targetItems[j], previewMode, keepFormatting)) didAny = true;
        }

        if (!didAny && showAlerts) {
            alert(L('alertNeedPathText'));
        }
    }

    /* アーチ/正円モードで一緒に選択されたパスを非表示（プレビュー）/削除（実行）/ Hide (preview) or remove (execute) paths selected together in arc/circle mode */
    function handleArcModeSelectedPaths(previewMode) {
        try {
            if (!selectedPaths || selectedPaths.length === 0) return;
            for (var i = selectedPaths.length - 1; i >= 0; i--) {
                var p = selectedPaths[i];
                if (!p) continue;
                if (previewMode) {
                    __preview_hideOriginal(p);
                } else {
                    try { p.remove(); } catch (_) { }
                }
            }
        } catch (_) { }
    }

    /* メイン処理：アーチパスを生成してパス上文字を作成 / Main process: generate an arc path and create path text */
    function mainProcessGenerateArcPath(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var __createdTexts = [];

        // アーチモードでは一緒に選択されたパスを隠す/削除
        handleArcModeSelectedPaths(previewMode);

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentLayer = originalText.layer;

            // テキスト範囲からアーチ状パスを生成
            var textPath = createArcPathFromText(originalText, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(L('alertArcFail'));
                continue;
            }

            // 生成アーチパスは非表示ガイド / Generated arc path is an invisible guide
            __applyInvisiblePathStyle(textPath);
            if (previewMode) __previewTempItems.push(textPath);

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);

            // 変換時に AI がスタイルを上書きするため、パス上文字のパスも非表示に
            try {
                if (textOnAPath.textPath) __applyInvisiblePathStyle(textOnAPath.textPath);
            } catch (_) { }

            decoratePathText(textOnAPath, originalText, previewMode);
            __createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(__createdTexts);
    }

    /* メイン処理：テキスト幅を直径とする正円を生成してパス上文字を作成 / Main process: generate a circle (diameter = text width) and create path text */
    function mainProcessGenerateCircle(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var __createdTexts = [];

        // 正円モードでも一緒に選択されたパスを隠す/削除（アーチと同じ）
        handleArcModeSelectedPaths(previewMode);

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentLayer = originalText.layer;

            // テキスト幅を直径とする正円パスを生成
            var textPath = createCirclePathFromText(originalText, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(L('alertArcFail'));
                continue;
            }

            if (previewMode) {
                __preview_applyPathStyle(textPath);
                __previewTempItems.push(textPath);
            } else {
                __applyExecutePathStyle(textPath); // Execute: keep the generated circle visible
            }

            // 生成した正円を中心まわりに時計回り90°回転
            try { textPath.rotate(90, true, true, true, true, Transformation.CENTER); } catch (_) { }

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);

            // 変換時に AI がスタイルを上書きするため、パス上文字のパスを再整形
            try {
                if (textOnAPath.textPath) {
                    if (previewMode) {
                        __preview_applyPathStyle(textOnAPath.textPath);
                    } else {
                        __applyExecutePathStyle(textOnAPath.textPath);
                    }
                }
            } catch (_) { }

            // 正円は常に中央揃え。内側配置時は開始/終了位置を +0.5 補正
            decoratePathText(textOnAPath, originalText, previewMode, {
                forceCenter: true,
                circleInsideShift: cbReverse.value
            });
            __createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(__createdTexts);
    }

    /* アウトライン化でテキストの範囲を計測 / Measure the text bounds via outlines ([L, T, R, B]) */
    function measureArcTextBounds(originalText) {
        // ライブテキストの geometricBounds より安定するためアウトライン化して計測
        var measureTexts = [originalText.duplicate(), originalText.duplicate()];
        measureTexts[0].contents = '';

        // 先頭行の textRange を measureTexts[0] に複製（レガシー手法、ポイント文字＋1行以上を前提）
        for (var i = 0; i < originalText.lines[0].length; i++) {
            originalText.textRanges[i].duplicate(measureTexts[0]);
        }
        for (var k = 0; k < measureTexts.length; k++) {
            measureTexts[k] = measureTexts[k].createOutline();
        }

        var textBounds = measureTexts[1].geometricBounds; // [L, T, R, B]
        textBounds[3] = measureTexts[0].geometricBounds[3];

        for (var r = 0; r < measureTexts.length; r++) {
            try { measureTexts[r].remove(); } catch (_) { }
        }
        return textBounds;
    }

    /* まるみスライダー（0-100、50=既定）を読み取る / Read the roundness slider (0-100, 50 = default) */
    function readArcRoundnessPercent() {
        var percent = 50;
        try {
            if (typeof slArcRoundness !== 'undefined' && slArcRoundness) {
                percent = Number(slArcRoundness.value);
            }
        } catch (_) { percent = 50; }
        if (isNaN(percent)) percent = 50;
        if (percent < 0) percent = 0;
        if (percent > 100) percent = 100;
        return percent;
    }

    /* アーチ方向の符号を返す（上=+1 / 下=-1）/ Return the arc direction sign (Up=+1, Down=-1) */
    function readArcDirectionSign() {
        try {
            if (typeof rbArcDown !== 'undefined' && rbArcDown && rbArcDown.value) return -1;
        } catch (_) { }
        return 1;
    }

    /* 直線パスのハンドルを調整してアーチ状に曲げる / Bend a straight path into an arc by adjusting handles */
    function applyArcHandles(arcPath, roundnessPercent, directionSign) {
        var handleOffsetDivisors = [3.5, 4];
        var pathLength = arcPath.length;

        // handleOffsets[0]: 水平ハンドル（固定）/ handleOffsets[1]: 垂直ハンドル＝カーブの深さ
        // 50% で垂直ハンドルは pathLength / handleOffsetDivisors[1]（レガシー既定）
        var handleOffsets = [
            pathLength / handleOffsetDivisors[0],
            pathLength * (roundnessPercent / 100) * (2 / handleOffsetDivisors[1])
        ];
        var pathPoints = arcPath.pathPoints;

        pathPoints[0].rightDirection = [
            pathPoints[0].rightDirection[0] + handleOffsets[0],
            pathPoints[0].rightDirection[1] + (directionSign * handleOffsets[1])
        ];
        pathPoints[1].leftDirection = [
            pathPoints[1].leftDirection[0] - handleOffsets[0],
            pathPoints[1].leftDirection[1] + (directionSign * handleOffsets[1])
        ];
    }

    /* テキストのアウトライン範囲からアーチ状のパスを生成 / Build an arc-like path from the text outline bounds */
    function createArcPathFromText(originalText, layer) {
        try {
            // Guard: empty / invalid text
            if (!originalText || originalText.typename !== 'TextFrame') return null;
            if (!originalText.lines || originalText.lines.length === 0) return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;

            var textBounds = measureArcTextBounds(originalText);

            // ベースラインに沿った直線パスを作成（baselineYMultiplier=1.02 はレガシー既定）
            var baselineY = textBounds[3] * 1.02;
            var arcPath = layer.pathItems.add();
            arcPath.setEntirePath([
                [textBounds[0], baselineY],
                [textBounds[2], baselineY]
            ]);

            // 生成パスは非表示（塗り・線なし）/ Make the generated path invisible
            try {
                arcPath.stroked = false;
                arcPath.filled = false;
            } catch (_) { }

            // ハンドルを調整してアーチ状に曲げる / Bend into an arc
            applyArcHandles(arcPath, readArcRoundnessPercent(), readArcDirectionSign());
            return arcPath;
        } catch (e) {
            return null;
        }
    }

    /* テキスト幅を直径とする正円パスを生成 / Build a circle path whose diameter equals the full text width */
    function createCirclePathFromText(originalText, layer) {
        try {
            // Guard
            if (!originalText || originalText.typename !== 'TextFrame') return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;

            // Use outline bounds for stability
            var measureText = originalText.duplicate();
            var measureOutline = null;
            try { measureOutline = measureText.createOutline(); } catch (_) { measureOutline = null; }
            try { measureText.remove(); } catch (_) { }
            if (!measureOutline) return null;

            var outlineBounds = measureOutline.geometricBounds; // [L, T, R, B]
            try { measureOutline.remove(); } catch (_) { }

            var textWidth = outlineBounds[2] - outlineBounds[0];
            if (!textWidth || isNaN(textWidth) || textWidth <= 0) return null;

            // Center of text outline
            var centerX = (outlineBounds[0] + outlineBounds[2]) / 2;
            var centerY = (outlineBounds[1] + outlineBounds[3]) / 2;

            var diameter = textWidth; // diameter = text width
            var left = centerX - (diameter / 2);
            var top = centerY + (diameter / 2);

            var circlePath = layer.pathItems.ellipse(top, left, diameter, diameter);

            // Default invisible (caller will set preview style if needed)
            try {
                circlePath.stroked = false;
                circlePath.filled = false;
            } catch (_) { }

            return circlePath;
        } catch (e) {
            return null;
        }
    }

    /* 各テキストに使う元パスを決定（選択パス優先、無ければ自身の textPath）/ Decide the base path for each text (selected paths first, else its own textPath) */
    function resolveBasePathForText(selectedPaths, textCount, index, originalText) {
        if (selectedPaths && selectedPaths.length > 0) {
            if (selectedPaths.length === textCount) return selectedPaths[index];
            return selectedPaths[0];
        }
        // No selected path: try PathText's own path
        try {
            if (originalText && originalText.typename === 'TextFrame') {
                if (originalText.kind === TextType.PATHTEXT && originalText.textPath) {
                    return originalText.textPath;
                }
            }
        } catch (_) { }
        return null;
    }

    /* 選択された元パスを非表示（塗り・線なし）にする / Make the original selected path invisible (no fill, no stroke) */
    function makeOriginalPathInvisible(basePath) {
        try {
            if (basePath.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < basePath.pathItems.length; cp++) {
                    basePath.pathItems[cp].stroked = false;
                    basePath.pathItems[cp].filled = false;
                }
            } else {
                basePath.stroked = false;
                basePath.filled = false;
            }
        } catch (_) { }
    }

    /* 複製対象の PathItem を取得（CompoundPath は先頭を使用）/ Resolve the PathItem to duplicate (first item for a CompoundPathItem) */
    function resolveSourcePath(basePath) {
        try {
            return (basePath.typename === 'CompoundPathItem') ? basePath.pathItems[0] : basePath;
        } catch (_) {
            return null;
        }
    }

    /* テキスト用にパスを複製し同じレイヤーへ配置 / Duplicate a path for the text and place it on the same layer */
    function duplicatePathForText(basePath, currentLayer) {
        var srcPath = resolveSourcePath(basePath);
        if (!srcPath) return null;

        // 1) Try native duplicate first (fast)
        try {
            var dup = srcPath.duplicate();
            try { dup.move(currentLayer, ElementPlacement.PLACEATBEGINNING); } catch (_) { }
            return dup;
        } catch (_) {
            // 2) Fallback: duplicate by copying anchors/handles (more robust)
            try {
                var dup2 = duplicatePathWithHandles(srcPath, currentLayer);
                return dup2;
            } catch (_) {
                return null;
            }
        }
    }

    /* 選択内のポイント文字/パス上文字を再帰的に収集 / Collect point/path text items recursively from the selection */
    function getTargetTextItems(items) {
        var textItems = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === 'TextFrame') {
                try {
                    if (item.kind === TextType.POINTTEXT || item.kind === TextType.PATHTEXT) {
                        textItems.push(item);
                    }
                } catch (_) { }
            } else if (item.typename === 'GroupItem') {
                textItems = textItems.concat(getTargetTextItems(item.pageItems));
            }
        }
        return textItems;
    }

    /* 選択内のパス（PathItem/CompoundPathItem）を再帰的に収集 / Collect path items recursively from the selection (groups included) */
    function getSelectedPathItems(items) {
        var pathItems = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === 'PathItem') {
                pathItems.push(item);
            } else if (item.typename === 'CompoundPathItem') {
                pathItems.push(item);
            } else if (item.typename === 'GroupItem') {
                pathItems = pathItems.concat(getSelectedPathItems(item.pageItems));
            }
        }
        return pathItems;
    }

    /* パスが閉じているか判定 / Tell whether a path item is closed */
    function isClosedPathItem(item) {
        try {
            if (!item) return false;
            if (item.typename === 'CompoundPathItem') {
                if (item.pathItems && item.pathItems.length > 0) return !!item.pathItems[0].closed;
                return false;
            }
            if (item.typename === 'PathItem') return !!item.closed;
        } catch (_) { }
        return false;
    }

    /* フィット対象（編集可能なオープンパス上文字）か判定 / Tell whether a frame is editable open path text to fit */
    function isTargetPathText(textFrame) {
        try {
            if (!textFrame || textFrame.typename !== 'TextFrame') return false;
            if (textFrame.kind !== TextType.PATHTEXT) return false;
            if (!textFrame.editable || textFrame.locked || textFrame.hidden) return false;
            var textPath = textFrame.textPath;
            if (!textPath) return false;
            // Only OPEN paths
            if (isClosedPathItem(textPath)) return false;
            return true;
        } catch (_) { }
        return false;
    }

    /* 表示行数を取得（最低1）/ Get the visible line count (at least 1) */
    function getLineAmount(textFrame) {
        try {
            if (textFrame.lines && textFrame.lines.length > 0) return textFrame.lines.length;
        } catch (_) { }
        return 1;
    }

    /* テキストがあふれている（オーバーセット）か判定 / Detect whether the text is overset */
    function isOverset(textFrame, lineAmount) {
        try {
            if (!textFrame) return false;
            if (textFrame.lines.length > 0) {
                var charactersOnVisibleLines = 0;
                if (typeof (lineAmount) === 'undefined' || lineAmount === null) {
                    lineAmount = 1;
                } else {
                    lineAmount = Math.floor(lineAmount);
                    if (lineAmount < 1) lineAmount = 1;
                    if (lineAmount > textFrame.lines.length) lineAmount = textFrame.lines.length;
                }
                for (var i = 0; i < lineAmount; i++) {
                    charactersOnVisibleLines += textFrame.lines[i].characters.length;
                }
                return (charactersOnVisibleLines < textFrame.characters.length);
            } else if (textFrame.characters.length > 0) {
                return true;
            }
        } catch (_) { }
        return false;
    }

    /* オープンパスの端まで文字サイズだけでパス上文字をフィット / Fit path text to the open-path endpoints by adjusting font size only */
    // - If overset: shrink by small steps until it fits.
    // - If NOT overset: grow to intentionally create overset, then shrink to fit (near-maximum).
    function fitTextToOpenPath(frames) {
        if (!frames || frames.length === 0) return false;

        var opt = {
            increment: 0.1,
            minFontSize: 0.1,
            maxShrinkIter: 2000,
            maxGrowIter: 10,
            maxFontSize: 2000,
            alertOnMaxIter: false
        };

        /* オーバーセットになるまで拡大→収まるまで縮小 / Grow until overset, then shrink until it fits */
        function shrinkFont(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);

                // If it is NOT overset, grow first (doubling) until it becomes overset (or hits safety limit)
                if (!isOverset(textFrame, lineAmount)) {
                    var growIterations = 0;
                    while (!isOverset(textFrame, lineAmount) && growIterations < opt.maxGrowIter) {
                        var growSize = textFrame.textRange.characterAttributes.size;
                        if (growSize >= opt.maxFontSize) break;
                        textFrame.textRange.characterAttributes.size = Math.min(opt.maxFontSize, growSize * 2);
                        growIterations++;
                    }
                }

                // Then shrink in small steps until it fits
                var shrinkIterations = 0;
                while (isOverset(textFrame, lineAmount)) {
                    var currentSize = textFrame.textRange.characterAttributes.size;
                    if (currentSize <= opt.minFontSize) break;

                    textFrame.textRange.characterAttributes.size = Math.max(opt.minFontSize, currentSize - opt.increment);

                    shrinkIterations++;
                    if (shrinkIterations >= opt.maxShrinkIter) {
                        if (opt.alertOnMaxIter) {
                            try { alert('フィット処理（縮小）が上限回数に達しました'); } catch (_) { }
                        }
                        break;
                    }
                }
            } catch (_) { }
        }

        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            if (!isTargetPathText(textFrame)) continue;
            shrinkFont(textFrame);
        }

        return true;
    }

    /* オープンパスの端までトラッキングだけで合わせる（文字サイズは維持）/ Fit path text to the open-path endpoints by adjusting tracking only (font size kept) */
    function fitTextToOpenPathByTracking(frames) {
        if (!frames || frames.length === 0) return false;

        var opt = {
            coarseStep: 50,     // 粗調整1ステップのトラッキング量 / tracking units per coarse step
            fineStep: 1,        // 微調整1ステップのトラッキング量 / tracking units per fine step
            minTracking: -1000, // 累積トラッキングの下限 / tightest allowed cumulative delta
            maxTracking: 20000, // 累積トラッキングの上限 / loosest allowed cumulative delta
            maxIter: 4000
        };

        /* 1フレームをトラッキングでパス幅にフィット / Fit a single frame to the path width by tracking */
        function fitByTracking(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);
                var applied = 0; // これまでに適用した累積トラッキング / cumulative tracking delta applied so far
                var iterations;

                if (isOverset(textFrame, lineAmount)) {
                    // あふれている：収まるまで粗く詰める / too wide: tighten (coarse) until it fits
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied - opt.coarseStep < opt.minTracking) break;
                        addDeltaToTextRanges(textFrame, 'tracking', -opt.coarseStep);
                        applied -= opt.coarseStep;
                        iterations++;
                    }
                    // 再びあふれるまで微調整で広げる / loosen back (fine) until it overflows again
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied + opt.fineStep > opt.maxTracking) break;
                        addDeltaToTextRanges(textFrame, 'tracking', opt.fineStep);
                        applied += opt.fineStep;
                        iterations++;
                    }
                    // 1ステップ行き過ぎたら戻して収める / stepped one step too far: pull back once so it fits
                    if (isOverset(textFrame, lineAmount)) {
                        addDeltaToTextRanges(textFrame, 'tracking', -opt.fineStep);
                        applied -= opt.fineStep;
                    }
                } else {
                    // 余裕あり：あふれるまで粗く広げる / fits with room: loosen (coarse) until it overflows
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied + opt.coarseStep > opt.maxTracking) break;
                        addDeltaToTextRanges(textFrame, 'tracking', opt.coarseStep);
                        applied += opt.coarseStep;
                        iterations++;
                    }
                    // 収まるまで微調整で詰める / tighten back (fine) until it fits
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied - opt.fineStep < opt.minTracking) break;
                        addDeltaToTextRanges(textFrame, 'tracking', -opt.fineStep);
                        applied -= opt.fineStep;
                        iterations++;
                    }
                }
            } catch (_) { }
        }

        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            if (!isTargetPathText(textFrame)) continue;
            fitByTracking(textFrame);
        }

        return true;
    }
}());