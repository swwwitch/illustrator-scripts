#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
テキストとパスの変換
 
### 更新日：
20260302
 
### 概要：
ポイント文字／パス上文字を、目的に応じて変換します。
- パスが一緒に選択されている場合：そのパスを複製して「パス上文字」を生成（複数テキストにも対応）
- パスが選択されていない場合：アーチ状のパスを自動生成して「パス上文字」を生成（テキストのみ選択でも実行可能）
- 処理：パス上文字を「テキスト」と「パス」に分離（パスを複製し、ポイント文字を生成）
- オプション：行揃え（左揃え / 中央 / 右揃え / 両端揃え）を指定可能
- 位置：開始位置(startTValue) / 終了位置(endTValue) を指定可能
- テキスト調整：ベースラインシフト / トラッキング / 文字サイズ を既存値に対して加算/減算可能
*/

(function () {

    /* バージョン / Version */
    // Version
    var SCRIPT_VERSION = "v1.1";

    // Language
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "パス上テキスト（変換と調整）", en: "Path Text (Convert & Adjust)" },

        /* Panels / パネル */
        panelProcess: { ja: "パス上文字にする", en: "Create Path Text" },
        panelSplit: { ja: "分離", en: "Split" },
        panelAlign: { ja: "行揃え", en: "Alignment" },
        panelOption: { ja: "オプション", en: "Options" },
        panelEffect: { ja: "効果", en: "Effect" },
        panelPosition: { ja: "位置", en: "Position" },
        panelTextAdjust: { ja: "テキスト調整", en: "Text Adjust" },

        /* Process radios / 処理ラジオ */
        toPathText: { ja: "パス上文字", en: "Convert to Path Text" },
        genArcPath: { ja: "アーチ状のパスを生成", en: "Generate Arc Path" },
        genCircle: { ja: "正円を生成", en: "Generate Circle" },
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
        optFixOverset: { ja: "文字あふれを解消", en: "Fix Overset" },

        arcDirection: { ja: "アーチ方向", en: "Arc Direction" },
        arcUp: { ja: "上", en: "Up" },
        arcDown: { ja: "下", en: "Down" },

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

    function L(key) {
        try {
            if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
            if (LABELS[key] && LABELS[key].ja) return LABELS[key].ja;
        } catch (_) { }
        return key;
    }

    // Get items
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return false;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;
    // Base selection snapshot (used for stable preview while dialog is open)
    var __baseSelection = [];
    try { __baseSelection = sel.slice(0); } catch (_) { __baseSelection = []; }

    var targetItems = getTargetTextItems(sel);
    var selectedPaths = getSelectedPathItems(sel);


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

        // "パス上テキストに" is available if a path is selected OR a PathText provides its own path
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
                try { rbAlignCenter.value = true; } catch (_) { }
                try { rbEffectRainbow.value = true; } catch (_) { }
            }
            if ((rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value) && !canSplit) {
                rbSplitTextAndPath.value = false;
                rbSplitTextAndPathNoFormat.value = false;
                rbGenArcPath.value = true;
                try { rbAlignCenter.value = true; } catch (_) { }
                try { rbEffectRainbow.value = true; } catch (_) { }
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

    // Option (radio)
    var pnlOpt = leftCol.add('panel', undefined, L('panelProcess'));
    pnlOpt.orientation = 'column';
    pnlOpt.alignChildren = ['left', 'top'];
    pnlOpt.margins = [15, 20, 15, 10];

    var rbToPathText = pnlOpt.add('radiobutton', undefined, L('toPathText'));
    var rbGenArcPath = pnlOpt.add('radiobutton', undefined, L('genArcPath'));
    var rbGenCircle = pnlOpt.add('radiobutton', undefined, L('genCircle'));
    rbToPathText.value = true;

    // Split panel (under Process)
    var pnlSplit = leftCol.add('panel', undefined, L('panelSplit'));
    pnlSplit.orientation = 'column';
    pnlSplit.alignChildren = ['left', 'top'];
    pnlSplit.margins = [15, 20, 15, 10];

    var rbSplitTextAndPath = pnlSplit.add('radiobutton', undefined, L('splitTextAndPath'));
    var rbSplitTextAndPathNoFormat = pnlSplit.add('radiobutton', undefined, L('splitTextAndPathNoFormat'));
    rbSplitTextAndPath.value = false;
    rbSplitTextAndPathNoFormat.value = false;

    var cbSplitDeletePath = pnlSplit.add('checkbox', undefined, L('splitDeletePath'));
    cbSplitDeletePath.value = false; // デフォルトOFF

    // --- Mutually exclusive helpers for Process/Split radios ---
    function __clearProcessRadios() {
        try { rbToPathText.value = false; } catch (_) { }
        try { rbGenArcPath.value = false; } catch (_) { }
        try { rbGenCircle.value = false; } catch (_) { }
    }

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
        try { rbAlignCenter.value = true; } catch (_) { }
        // 虹はあとで作ってから適用
    }
    if (!canSplit && (rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value)) {
        rbSplitTextAndPath.value = false;
        rbSplitTextAndPathNoFormat.value = false;
        rbGenArcPath.value = true;
        try { rbAlignCenter.value = true; } catch (_) { }
        // 虹はあとで作ってから適用
    }

    // Option panel (above Alignment in RIGHT column)
    var pnlOption = rightCol.add('panel', undefined, L('panelOption'));
    pnlOption.orientation = 'column';
    pnlOption.alignChildren = ['left', 'top'];
    pnlOption.margins = [15, 20, 15, 10];

    var cbReverse = pnlOption.add('checkbox', undefined, L('optReverse'));
    cbReverse.value = false;

    // Arc direction UI
    var grpArcDir = pnlOption.add('group');
    grpArcDir.orientation = 'row';
    grpArcDir.alignChildren = ['left', 'center'];

    var stArcDir = grpArcDir.add('statictext', undefined, L('arcDirection'));
    stArcDir.preferredSize.width = 70;

    var rbArcUp = grpArcDir.add('radiobutton', undefined, L('arcUp'));
    var rbArcDown = grpArcDir.add('radiobutton', undefined, L('arcDown'));
    rbArcUp.value = true;

    var cbFixOverset = pnlOption.add('checkbox', undefined, L('optFixOverset'));
    cbFixOverset.value = false;

    // Alignment / Indent panel  (moved to RIGHT)
    var pnlAlignIndent = rightCol.add('panel', undefined, L('panelAlign'));
    pnlAlignIndent.orientation = 'column';
    pnlAlignIndent.alignChildren = ['left', 'top'];
    pnlAlignIndent.margins = [15, 20, 15, 10];

    var rbAlignLeft = pnlAlignIndent.add('radiobutton', undefined, L('alignLeft'));
    var rbAlignCenter = pnlAlignIndent.add('radiobutton', undefined, L('alignCenter'));
    var rbAlignRight = pnlAlignIndent.add('radiobutton', undefined, L('alignRight'));
    var rbAlignFullJustify = pnlAlignIndent.add('radiobutton', undefined, L('alignFullJustify'));

    // Default: 左揃え
    rbAlignLeft.value = true;

    // -------------------------------------------------
    // Full-width (span across columns): 効果 -> 位置 -> ベースラインシフト
    // -------------------------------------------------
    var wideCol = dlg.add('group');
    wideCol.orientation = 'column';
    wideCol.alignChildren = ['fill', 'top'];
    wideCol.alignment = ['fill', 'top'];

    // Effect panel (full-width)
    var pnlEffect = wideCol.add('panel', undefined, L('panelEffect'));
    pnlEffect.orientation = 'row';
    pnlEffect.alignChildren = ['center', 'center'];
    pnlEffect.margins = [15, 20, 15, 10];

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
    var pnlTValue = wideCol.add('panel', undefined, L('panelPosition'));
    pnlTValue.orientation = 'column';
    pnlTValue.alignChildren = ['left', 'top'];
    pnlTValue.margins = [15, 20, 15, 10];

    var rowStart = pnlTValue.add('group');
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

    var rowEnd = pnlTValue.add('group');
    rowEnd.orientation = 'row';
    rowEnd.alignChildren = ['left', 'center'];
    var cbEndT = rowEnd.add('checkbox', undefined, '');
    cbEndT.value = false;
    var stEnd = rowEnd.add('statictext', undefined, L('endPos'));
    stEnd.preferredSize.width = 60;
    var etEndT = rowEnd.add('edittext', undefined, '1.0');
    etEndT.characters = 5;
    var slEndT = rowEnd.add('slider', undefined, 100, 0, 100); // 0.0 - 1.0
    slEndT.preferredSize.width = 180;

    // Baseline shift panel  (moved to FULL WIDTH)
    var pnlBaselineShift = wideCol.add('panel', undefined, L('panelTextAdjust'));
    pnlBaselineShift.orientation = 'column';
    pnlBaselineShift.alignChildren = ['left', 'top'];
    pnlBaselineShift.margins = [15, 20, 15, 10];

    var rowBaseShift = pnlBaselineShift.add('group');
    rowBaseShift.orientation = 'row';
    rowBaseShift.alignChildren = ['left', 'center'];

    var stBaseShift = rowBaseShift.add('statictext', undefined, L('baseShift'));
    stBaseShift.preferredSize.width = 80;

    var etBaseShift = rowBaseShift.add('edittext', undefined, '0.0');
    etBaseShift.characters = 6;
    // Slider uses 0.1pt steps, range will be updated dynamically to ±fontSize (in 0.1pt units)
    var slBaseShift = rowBaseShift.add('slider', undefined, 0, -1000, 1000);
    slBaseShift.preferredSize.width = 180;

    // Tracking panel row
    var rowTracking = pnlBaselineShift.add('group');
    rowTracking.orientation = 'row';
    rowTracking.alignChildren = ['left', 'center'];

    var stTracking = rowTracking.add('statictext', undefined, L('tracking'));
    stTracking.preferredSize.width = 80;

    var etTracking = rowTracking.add('edittext', undefined, '0');
    etTracking.characters = 6;

    var slTracking = rowTracking.add('slider', undefined, 0, -100, 500);
    slTracking.preferredSize.width = 180;

    // Font size panel row
    var rowFontSize = pnlBaselineShift.add('group');
    rowFontSize.orientation = 'row';
    rowFontSize.alignChildren = ['left', 'center'];

    var stFontSize = rowFontSize.add('statictext', undefined, L('fontSize'));
    stFontSize.preferredSize.width = 80;

    var etFontSize = rowFontSize.add('edittext', undefined, '0.0');
    etFontSize.characters = 6;

    // Slider uses 0.1pt steps, range will be updated dynamically to ±fontSize (in 0.1pt units)
    var slFontSize = rowFontSize.add('slider', undefined, 0, -1000, 1000);
    slFontSize.preferredSize.width = 180;

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

    function __preview_applyPathStyle(basePath) {
        try {
            if (!basePath) return;

            function applyPreviewStyle(pi) {
                __preview_recordPathState(pi);
                try {
                    pi.filled = false;
                    pi.stroked = true;
                    pi.strokeWidth = 1;

                    var black = new CMYKColor();
                    black.cyan = 0;
                    black.magenta = 0;
                    black.yellow = 0;
                    black.black = 100;
                    pi.strokeColor = black;

                    pi.opacity = 50;
                } catch (_) { }
            }

            if (basePath.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < basePath.pathItems.length; cp++) {
                    applyPreviewStyle(basePath.pathItems[cp]);
                }
            } else {
                applyPreviewStyle(basePath);
            }
        } catch (_) { }
    }

    // Apply path style for execute mode (keep path visible)
    function __applyExecutePathStyle(pathItem) {
        try {
            if (!pathItem) return;

            function applyStyle(pi) {
                try {
                    pi.filled = false;
                    pi.stroked = true;
                    pi.strokeWidth = 1;

                    var black = new CMYKColor();
                    black.cyan = 0;
                    black.magenta = 0;
                    black.yellow = 0;
                    black.black = 100;
                    pi.strokeColor = black;

                    pi.opacity = 50;
                } catch (_) { }
            }

            if (pathItem.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < pathItem.pathItems.length; cp++) {
                    applyStyle(pathItem.pathItems[cp]);
                }
            } else {
                applyStyle(pathItem);
            }
        } catch (_) { }
    }

    function __preview_hideOriginal(it) {
        try {
            if (!it) return;
            // Avoid duplicates
            for (var k = 0; k < __previewHiddenOriginals.length; k++) {
                if (__previewHiddenOriginals[k] === it) return;
            }
            it.hidden = true;
            __previewHiddenOriginals.push(it);
        } catch (_) { }
    }

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

    function refreshPreviewIfNeeded() {
        if (cbPreview.value) {
            applyPreview();
        }
    }

    rbToPathText.onClick = function () {
        __clearSplitRadios();
        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    rbGenCircle.onClick = function () {
        __clearSplitRadios();

        // 正円モードのデフォルト開始/終了位置
        try {
            cbStartT.value = true;
            cbEndT.value = true;
            etStartT.text = "0.5";
            etEndT.text = "3.5";
            __syncTFromEdits();
        } catch (_) { }

        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    rbGenArcPath.onClick = function () {
        __clearSplitRadios();
        __updatePanelsByMode();
        // Auto-preset for arc path mode
        try { rbAlignCenter.value = true; } catch (_) { }
        try { rbEffectRainbow.value = true; } catch (_) { }
        refreshPreviewIfNeeded();
    };
    rbSplitTextAndPath.onClick = function () {
        __clearProcessRadios();
        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    rbSplitTextAndPathNoFormat.onClick = function () {
        __clearProcessRadios();
        __updatePanelsByMode();
        refreshPreviewIfNeeded();
    };
    // Disable/enable panels depending on mode
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

        // Keep "処理" and "分離" always enabled
        try { pnlOpt.enabled = true; } catch (_) { }
        try { pnlSplit.enabled = true; } catch (_) { }

        // Splitモード時はそれ以外をすべてディム
        var en = !isSplitMode;
        try { pnlOption.enabled = en; } catch (_) { }
        try { pnlAlignIndent.enabled = en; } catch (_) { }
        try { pnlEffect.enabled = en; } catch (_) { }
        try { pnlBaselineShift.enabled = en; } catch (_) { }

        // 位置パネルは常に有効（Splitモードでない限り）
        try { pnlTValue.enabled = en; } catch (_) { }
    }

    cbReverse.onClick = refreshPreviewIfNeeded;
    cbFixOverset.onClick = refreshPreviewIfNeeded;
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
    // Sync baseline shift UI (edittext <-> slider)
    // - slider unit: 0.1pt (value = delta * 10)
    // - slider range: ±fontSize (pt) based on current selection
    var __bsSyncLock = false;

    function __formatBS(n) {
        // 0.1pt
        var v = Math.round(n * 10) / 10;
        return v.toFixed(1);
    }

    function __getBaselineShiftRefFontSizePt() {
        // Use first available target's font size as reference
        try {
            var s = null;
            if (targetItems && targetItems.length > 0) {
                s = targetItems[0];
            }
            if (s && s.typename === 'TextFrame') {
                // Prefer first textRange
                if (s.textRanges && s.textRanges.length > 0) {
                    var sz = s.textRanges[0].characterAttributes.size;
                    if (sz && !isNaN(sz)) return sz;
                }
                // Fallback
                var sz2 = s.textRange.characterAttributes.size;
                if (sz2 && !isNaN(sz2)) return sz2;
            }
        } catch (_) { }
        return 100; // fallback
    }

    function __updateBaselineShiftSliderRange() {
        try {
            var fs = __getBaselineShiftRefFontSizePt();
            // ±fontSize
            var minV = -Math.max(1, Math.round(fs * 10));
            var maxV = Math.max(1, Math.round(fs * 10));
            slBaseShift.minvalue = minV;
            slBaseShift.maxvalue = maxV;
        } catch (_) { }
    }

    function __syncBSFromEdit() {
        if (__bsSyncLock) return;
        __bsSyncLock = true;
        try {
            __updateBaselineShiftSliderRange();

            var v = __parseNumber(etBaseShift.text, 0);
            if (isNaN(v)) v = 0;

            // clamp to slider range
            var minPt = slBaseShift.minvalue / 10.0;
            var maxPt = slBaseShift.maxvalue / 10.0;
            v = __clamp(v, minPt, maxPt);

            etBaseShift.text = __formatBS(v);
            try { slBaseShift.value = Math.round(v * 10); } catch (_) { }
        } catch (_) { }
        __bsSyncLock = false;
    }

    function __syncBSFromSlider() {
        if (__bsSyncLock) return;
        __bsSyncLock = true;
        try {
            __updateBaselineShiftSliderRange();
            var v = (slBaseShift.value / 10.0);
            etBaseShift.text = __formatBS(v);
        } catch (_) { }
        __bsSyncLock = false;
    }

    etBaseShift.onChanging = function () { __syncBSFromEdit(); refreshPreviewIfNeeded(); };
    slBaseShift.onChanging = function () { __syncBSFromSlider(); refreshPreviewIfNeeded(); };
    __syncBSFromEdit();

    // Sync tracking UI (edittext <-> slider)
    var __trkSyncLock = false;

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
    slTracking.onChanging = function () { __syncTrkFromSlider(); refreshPreviewIfNeeded(); };
    __syncTrkFromEdit();

    // Sync font size delta UI (edittext <-> slider)
    // - slider unit: 0.1pt (value = delta * 10)
    // - slider range: ±fontSize (pt) based on current selection
    var __fsSyncLock = false;

    function __formatFS(n) {
        var v = Math.round(n * 10) / 10;
        return v.toFixed(1);
    }

    function __getFontSizeRefPt() {
        // Use first available target's font size as reference
        try {
            var s = null;
            if (targetItems && targetItems.length > 0) s = targetItems[0];
            if (s && s.typename === 'TextFrame') {
                if (s.textRanges && s.textRanges.length > 0) {
                    var sz = s.textRanges[0].characterAttributes.size;
                    if (sz && !isNaN(sz)) return sz;
                }
                var sz2 = s.textRange.characterAttributes.size;
                if (sz2 && !isNaN(sz2)) return sz2;
            }
        } catch (_) { }
        return 100;
    }

    function __updateFontSizeSliderRange() {
        try {
            var fs = __getFontSizeRefPt();
            var minV = -Math.max(1, Math.round(fs * 10));
            var maxV = Math.max(1, Math.round(fs * 10));
            slFontSize.minvalue = minV;
            slFontSize.maxvalue = maxV;
        } catch (_) { }
    }

    function __syncFSFromEdit() {
        if (__fsSyncLock) return;
        __fsSyncLock = true;
        try {
            __updateFontSizeSliderRange();

            var v = __parseNumber(etFontSize.text, 0);
            if (isNaN(v)) v = 0;

            var minPt = slFontSize.minvalue / 10.0;
            var maxPt = slFontSize.maxvalue / 10.0;
            v = __clamp(v, minPt, maxPt);

            etFontSize.text = __formatFS(v);
            try { slFontSize.value = Math.round(v * 10); } catch (_) { }
        } catch (_) { }
        __fsSyncLock = false;
    }

    function __syncFSFromSlider() {
        if (__fsSyncLock) return;
        __fsSyncLock = true;
        try {
            __updateFontSizeSliderRange();
            var v = (slFontSize.value / 10.0);
            etFontSize.text = __formatFS(v);
        } catch (_) { }
        __fsSyncLock = false;
    }

    etFontSize.onChanging = function () { __syncFSFromEdit(); refreshPreviewIfNeeded(); };
    slFontSize.onChanging = function () { __syncFSFromSlider(); refreshPreviewIfNeeded(); };
    __syncFSFromEdit();
    // Sync t-value UI (edittext <-> slider)
    var __tSyncLock = false;

    function __formatT(n) {
        // keep 1 decimal
        var v = Math.round(n * 10) / 10;
        // force one decimal place
        return v.toFixed(1);
    }

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

            etStartT.text = __formatT(s);
            etEndT.text = __formatT(e);

            try { slStartT.value = Math.round(s * 100); } catch (_) { }
            try { slEndT.value = Math.round(e * 100); } catch (_) { }
        } catch (_) { }
        __tSyncLock = false;
    }

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

            etStartT.text = __formatT(s);
            etEndT.text = __formatT(e);
        } catch (_) { }
        __tSyncLock = false;
    }

    etStartT.onChanging = function () { __syncTFromEdits(); refreshPreviewIfNeeded(); };
    etEndT.onChanging = function () { __syncTFromEdits(); refreshPreviewIfNeeded(); };
    slStartT.onChanging = function () { __syncTFromSliders(); refreshPreviewIfNeeded(); };
    slEndT.onChanging = function () { __syncTFromSliders(); refreshPreviewIfNeeded(); };

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

    // Apply path text effect via menu command (requires selection)
    function applyPathTextEffect(textFrame, previewMode) {
        var cmd = getSelectedEffectCommand();
        if (!cmd) return;

        var prevSel = null;
        try { prevSel = doc.selection; } catch (_) { prevSel = null; }

        try {
            try { doc.selection = []; } catch (_) { }
            try { textFrame.selected = true; } catch (_) { }
            app.executeMenuCommand(cmd);
        } catch (_) {
            // ignore
        }

        // Always restore selection (safer)
        try { doc.selection = prevSel; } catch (_) { }
    }

    function getSelectedEffectCommand() {
        if (rbEffectRainbow.value) return 'Rainbow';
        if (rbEffectDistort.value) return 'Skew';
        if (rbEffectRibbon.value) return '3D ribbon';
        if (rbEffectStep.value) return 'Stair Step';
        if (rbEffectGravity.value) return 'Gravity';
        return null;
    }

    // Apply paragraph justification AFTER text content is duplicated (so it won't be overwritten)
    function applyJustificationToTextFrame(tf) {
        try {
            if (!tf) return;

            var just = Justification.FULLJUSTIFYLASTLINELEFT;
            if (rbAlignFullJustify && rbAlignFullJustify.value) {
                just = Justification.FULLJUSTIFY;
            } else if (rbAlignRight && rbAlignRight.value) {
                just = Justification.RIGHT;
            } else if (rbAlignCenter && rbAlignCenter.value) {
                just = Justification.CENTER;
            } else {
                just = Justification.FULLJUSTIFYLASTLINELEFT;
            }

            // Apply to all paragraphs if available
            try {
                if (tf.paragraphs && tf.paragraphs.length > 0) {
                    for (var i = 0; i < tf.paragraphs.length; i++) {
                        try { tf.paragraphs[i].paragraphAttributes.justification = just; } catch (_) { }
                    }
                }
            } catch (_) { }

            // Fallback apply to whole range
            try { tf.textRange.paragraphAttributes.justification = just; } catch (_) { }
        } catch (_) { }
    }

    function __clamp(n, min, max) {
        if (n < min) return min;
        if (n > max) return max;
        return n;
    }

    function __parseTValue(str, fallback) {
        var n = Number(str);
        if (isNaN(n)) return fallback;
        return __clamp(n, 0, 5);
    }

    function applyStartEndTValue(tf) {
        try {
            if (!tf) return;
            var s = (cbStartT && cbStartT.value) ? __parseTValue(etStartT.text, 0.0) : 0.0;
            var e = (cbEndT && cbEndT.value) ? __parseTValue(etEndT.text, 1.0) : 1.0;

            // Ensure start <= end
            if (s > e) {
                var tmp = s;
                s = e;
                e = tmp;
            }

            tf.startTValue = s;
            tf.endTValue = e;
        } catch (_) { }
    }

    function __parseNumber(str, fallback) {
        var n = Number(str);
        if (isNaN(n)) return fallback;
        return n;
    }

    // Add delta to existing baselineShift for each textRange (keep original values + delta)
    function applyBaselineShiftDelta(tf) {
        try {
            if (!tf) return;
            var delta = __parseNumber(etBaseShift.text, 0);
            if (!delta) return; // 0 -> do nothing

            // Apply to each textRange so existing baselineShift is preserved and shifted
            for (var i = 0; i < tf.textRanges.length; i++) {
                var tr = tf.textRanges[i];
                try {
                    var cur = tr.characterAttributes.baselineShift;
                    tr.characterAttributes.baselineShift = cur + delta;
                } catch (_) { }
            }
        } catch (_) { }
    }

    // Add delta to existing tracking for each textRange (keep original values + delta)
    function applyTrackingDelta(tf) {
        try {
            if (!tf) return;
            var delta = Math.round(__parseNumber(etTracking.text, 0));
            if (!delta) return; // 0 -> do nothing

            for (var i = 0; i < tf.textRanges.length; i++) {
                var tr = tf.textRanges[i];
                try {
                    var cur = tr.characterAttributes.tracking;
                    tr.characterAttributes.tracking = cur + delta;
                } catch (_) { }
            }
        } catch (_) { }
    }

    // Add delta to existing font size for each textRange (keep original values + delta)
    function applyFontSizeDelta(tf) {
        try {
            if (!tf) return;
            var delta = __parseNumber(etFontSize.text, 0);
            if (!delta) return; // 0 -> do nothing

            for (var i = 0; i < tf.textRanges.length; i++) {
                var tr = tf.textRanges[i];
                try {
                    var cur = tr.characterAttributes.size;
                    var next = cur + delta;
                    if (next < 0.1) next = 0.1;
                    tr.characterAttributes.size = next;
                } catch (_) { }
            }
        } catch (_) { }
    }

    /* メイン処理 / Main process */
    function mainProcess(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var __createdTexts = [];

        for (var j = 0; j < targetItems.length; j++) {
            // Get original text frame item & current layer
            var originalText = targetItems[j];
            var currentlayer = originalText.layer;

            // Resolve base path + duplicate path for this text
            var basePath = resolveBasePathForText(selectedPaths, targetItems.length, j, originalText);
            if (!basePath) {
                // Should not happen because of Validation, but keep defensive
                if (showAlerts) alert(L('alertNeedPath'));
                return;
            }

            if (!previewMode) {
                makeOriginalPathInvisible(basePath);
            } else {
                // Preview: show original selected path as no fill / no stroke (restore on preview clear)
                __preview_applyPathStyle(basePath);
            }
            var textPath = duplicatePathForText(basePath, currentlayer);
            if (!textPath) {
                if (showAlerts) alert(L('alertDupPathFail'));
                continue;
            }

            if (previewMode) {
                __previewTempItems.push(textPath);
            } else {
                // Execute: keep the duplicated path visible
                __applyExecutePathStyle(textPath);
            }

            // Create Text on a path
            var textOnAPath = currentlayer.textFrames.pathText(textPath);
            // Keep stacking position (avoid appearing to disappear behind other objects)
            try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (_) { }
            if (previewMode) {
                __previewTempItems.push(textOnAPath);
            }

            // Inside placement option
            if (cbReverse.value) {
                try {
                    textOnAPath.textPath.polarity = PolarityValues.NEGATIVE;
                } catch (_) { }
            }

            applyPathTextEffect(textOnAPath, previewMode);
            // 位置（start/end）
            applyStartEndTValue(textOnAPath);

            // Duplicate textrange from original text frame item
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

            // 文字サイズ（既存値 + 指定値）
            applyFontSizeDelta(textOnAPath);

            // 行揃え（段落揃え）※ duplicate 後に適用しないと上書きされる
            applyJustificationToTextFrame(textOnAPath);

            if (!previewMode) __createdTexts.push(textOnAPath);

            // ベースラインシフト（既存値 + 指定値）
            applyBaselineShiftDelta(textOnAPath);

            // トラッキング（既存値 + 指定値）
            applyTrackingDelta(textOnAPath);

            // Remove or hide original text frame item
            if (previewMode) {
                __preview_hideOriginal(originalText);
            } else {
                originalText.remove();
            }

            // Select text on a path
            if (!previewMode) {
                try { textOnAPath.selected = true; } catch (_) { }
            }
            // (overset fix is applied once after the loop)
        }
        if (!previewMode && cbFixOverset.value) {
            try { fixOversetTextFrames(__createdTexts); } catch (_) { }
        }
    }

    // Duplicate a path preserving handles/pointTypes (based on the provided reference logic)
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

    // Split PathText into (1) duplicated path (visible) + (2) point text with copied content/format
    function splitPathTextAndPath(showAlerts, previewMode, keepFormatting) {
        if (typeof keepFormatting === 'undefined') keepFormatting = true;
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;

        var didAny = false;

        for (var j = targetItems.length - 1; j >= 0; j--) {
            var pathText = targetItems[j];
            if (!pathText || pathText.typename !== 'TextFrame') continue;

            var isPathText = false;
            try { isPathText = (pathText.kind === TextType.PATHTEXT); } catch (_) { isPathText = false; }
            if (!isPathText) continue;

            var originalPath = null;
            try { originalPath = pathText.textPath; } catch (_) { originalPath = null; }
            if (!originalPath) continue;

            didAny = true;

            var curLayer = null;
            try { curLayer = pathText.layer; } catch (_) { curLayer = null; }

            // 1) Duplicate path (preserve handles) unless "パスを削除" is ON
            var deletePath = false;
            try { deletePath = (cbSplitDeletePath && cbSplitDeletePath.value); } catch (_) { deletePath = false; }

            var newPath = null;
            if (!deletePath) {
                newPath = duplicatePathWithHandles(originalPath, curLayer);
                if (newPath) {
                    // Path appearance: no fill, black 1pt stroke
                    try {
                        newPath.filled = false;
                        newPath.stroked = true;

                        var blackColor = new CMYKColor();
                        blackColor.cyan = 0;
                        blackColor.magenta = 0;
                        blackColor.yellow = 0;
                        blackColor.black = 100;

                        newPath.strokeColor = blackColor;
                        newPath.strokeWidth = 1;
                    } catch (_) { }

                    if (previewMode) __previewTempItems.push(newPath);
                }
            }

            // 2) Create new point text near the original start anchor
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

                // Copy content
                if (keepFormatting) {
                    // Keep formatting
                    try { pathText.textRange.duplicate(newText); } catch (_) { }
                } else {
                    // Do not keep formatting (use the new text frame defaults)
                    try { newText.contents = pathText.contents; } catch (_) { }
                }

                // 文字サイズ（既存値 + 指定値）
                applyFontSizeDelta(newText);

                // ベースラインシフト（既存値 + 指定値）
                applyBaselineShiftDelta(newText);

                // トラッキング（既存値 + 指定値）
                applyTrackingDelta(newText);

                // Ensure fill/stroke colors are inherited only when keeping formatting
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

                // Apply current justification option (after duplicate)
                try { applyJustificationToTextFrame(newText); } catch (_) { }

                // Keep stacking position similar to original
                try { newText.move(pathText, ElementPlacement.PLACEBEFORE); } catch (_) { }

                if (previewMode) __previewTempItems.push(newText);
            }

            // 3) Remove or hide original PathText
            if (previewMode) {
                __preview_hideOriginal(pathText);
            } else {
                try { pathText.remove(); } catch (_) { }
            }
        }

        if (!didAny && showAlerts) {
            alert(L('alertNeedPathText'));
        }
    }

    // In Arc mode, if a path is selected together with text, it should not remain.
    // Preview: hide it (restored by __preview_clear)
    // Execute: delete it
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

    // Main process (generate arc path)
    function mainProcessGenerateArcPath(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var __createdTexts = [];

        // If a path is selected together, hide/remove it in arc mode
        handleArcModeSelectedPaths(previewMode);

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentlayer = originalText.layer;

            // Create an arc-like path from the text bounds (reference logic)
            var textPath = createArcPathFromText(originalText, currentlayer);
            if (!textPath) {
                if (showAlerts) alert(L('alertArcFail'));
                continue;
            }

            if (previewMode) {
                // Apply preview style to generated arc path
                __preview_applyPathStyle(textPath);
                __previewTempItems.push(textPath);
            } else {
                // Execute: keep the generated arc path visible
                __applyExecutePathStyle(textPath);
            }

            // Create Text on a path
            var textOnAPath = currentlayer.textFrames.pathText(textPath);
            // Keep stacking position (avoid appearing to disappear behind other objects)
            try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (_) { }
            if (previewMode) {
                __previewTempItems.push(textOnAPath);
            }

            // Ensure the path used by the PathText is visible (AI may override style on conversion)
            try {
                var tp = textOnAPath.textPath;
                if (tp) {
                    if (previewMode) {
                        __preview_applyPathStyle(tp);
                    } else {
                        __applyExecutePathStyle(tp);
                    }
                }
            } catch (_) { }

            // Inside placement option
            if (cbReverse.value) {
                try {
                    textOnAPath.textPath.polarity = PolarityValues.NEGATIVE;
                } catch (_) { }
            }

            applyPathTextEffect(textOnAPath, previewMode);
            // 位置（start/end）
            applyStartEndTValue(textOnAPath);

            // Duplicate textrange from original text frame item
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

            // 文字サイズ（既存値 + 指定値）
            applyFontSizeDelta(textOnAPath);

            // 行揃え（段落揃え）※ duplicate 後に適用しないと上書きされる
            applyJustificationToTextFrame(textOnAPath);

            if (!previewMode) __createdTexts.push(textOnAPath);

            // ベースラインシフト（既存値 + 指定値）
            applyBaselineShiftDelta(textOnAPath);

            // トラッキング（既存値 + 指定値）
            applyTrackingDelta(textOnAPath);

            // Remove or hide original text frame item
            if (previewMode) {
                __preview_hideOriginal(originalText);
            } else {
                originalText.remove();
            }

            // Select text on a path
            if (!previewMode) {
                try { textOnAPath.selected = true; } catch (_) { }
            }
            // (overset fix is applied once after the loop)
        }
        if (!previewMode && cbFixOverset.value) {
            try { fixOversetTextFrames(__createdTexts); } catch (_) { }
        }
    }

    // Main process (generate circle path with diameter = text width)
    function mainProcessGenerateCircle(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var __createdTexts = [];

        // If a path is selected together, hide/remove it in circle mode (same rule as arc mode)
        handleArcModeSelectedPaths(previewMode);

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentlayer = originalText.layer;

            // Create a circle path from the text bounds
            var textPath = createCirclePathFromText(originalText, currentlayer);
            if (!textPath) {
                if (showAlerts) alert(L('alertArcFail'));
                continue;
            }

            if (previewMode) {
                // Apply preview style to generated circle path
                __preview_applyPathStyle(textPath);
                __previewTempItems.push(textPath);
            } else {
                // Execute: keep the generated circle path visible
                __applyExecutePathStyle(textPath);
            }

            // Rotate the generated circle 90° clockwise (about its center)
            try {
                textPath.rotate(90, true, true, true, true, Transformation.CENTER);
            } catch (_) { }

            // Create Text on a path
            var textOnAPath = currentlayer.textFrames.pathText(textPath);
            try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (_) { }
            if (previewMode) __previewTempItems.push(textOnAPath);

            // Ensure the path used by the PathText is visible (AI may override style on conversion)
            try {
                var tp2 = textOnAPath.textPath;
                if (tp2) {
                    if (previewMode) {
                        __preview_applyPathStyle(tp2);
                    } else {
                        __applyExecutePathStyle(tp2);
                    }
                }
            } catch (_) { }

            // Inside placement option
            var __circleInside = false;
            if (cbReverse.value) {
                __circleInside = true;
                try { textOnAPath.textPath.polarity = PolarityValues.NEGATIVE; } catch (_) { }
            }

            applyPathTextEffect(textOnAPath, previewMode);

            // 正円モードでもUIの開始/終了位置を適用
            applyStartEndTValue(textOnAPath);

            // Circle + Inside compensation:
            // keep the apparent start position consistent by shifting start/end by +0.5
            if (__circleInside) {
                try {
                    var s0 = textOnAPath.startTValue;
                    var e0 = textOnAPath.endTValue;
                    var shift = 0.5;
                    var s1 = s0 + shift;
                    var e1 = e0 + shift;
                    // Clamp only to the global end cap (0..5). start may exceed 1 for circles.
                    s1 = __clamp(s1, 0.0, 5.0);
                    e1 = __clamp(e1, 0.0, 5.0);
                    textOnAPath.startTValue = s1;
                    textOnAPath.endTValue = e1;
                } catch (_) { }
            }

            // Duplicate textrange from original text frame item
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

            applyFontSizeDelta(textOnAPath);
            // 正円モードでは常に中央揃え
            try {
                if (textOnAPath.paragraphs && textOnAPath.paragraphs.length > 0) {
                    for (var pj = 0; pj < textOnAPath.paragraphs.length; pj++) {
                        try { textOnAPath.paragraphs[pj].paragraphAttributes.justification = Justification.CENTER; } catch (_) { }
                    }
                }
            } catch (_) { }
            try { textOnAPath.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (_) { }

            if (!previewMode) __createdTexts.push(textOnAPath);

            applyBaselineShiftDelta(textOnAPath);
            applyTrackingDelta(textOnAPath);

            // Remove or hide original text frame item
            if (previewMode) {
                __preview_hideOriginal(originalText);
            } else {
                originalText.remove();
            }

            if (!previewMode) {
                try { textOnAPath.selected = true; } catch (_) { }
            }
        }

        if (!previewMode && cbFixOverset.value) {
            try { fixOversetTextFrames(__createdTexts); } catch (_) { }
        }
    }

    // Create an arc-like path from point text using outline bounds (reference logic)
    function createArcPathFromText(originalText, layer) {
        // Reference defaults from the attached legacy script
        var baseLineOffset = 1.02;
        var handleOffsetDiv = [3.5, 4];

        // Guard: empty / invalid text
        try {
            if (!originalText || originalText.typename !== 'TextFrame') return null;
            if (!originalText.lines || originalText.lines.length === 0) return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;
        } catch (_) {
            return null;
        }

        try {
            // Get bounds using outline (more stable than geometricBounds of live text)
            var dummyTexts = [originalText.duplicate(), originalText.duplicate()];
            dummyTexts[0].contents = '';

            // Copy first line ranges into dummyTexts[0]
            // (This mirrors the legacy approach; it assumes point text with at least one line)
            for (var i = 0; i < originalText.lines[0].length; i++) {
                originalText.textRanges[i].duplicate(dummyTexts[0]);
            }

            for (var k = 0; k < dummyTexts.length; k++) {
                dummyTexts[k] = dummyTexts[k].createOutline();
            }

            var bounds = dummyTexts[1].geometricBounds; // [L, T, R, B]
            bounds[3] = dummyTexts[0].geometricBounds[3];

            for (var r = 0; r < dummyTexts.length; r++) {
                try { dummyTexts[r].remove(); } catch (_) { }
            }

            // Create base straight path
            var y = bounds[3] * baseLineOffset;
            var p = layer.pathItems.add();
            p.setEntirePath([
                [bounds[0], y],
                [bounds[2], y]
            ]);

            // Make the generated path invisible (no fill / no stroke)
            try {
                p.stroked = false;
                p.filled = false;
            } catch (_) { }

            // Arc-like handle adjustment
            var len = p.length;
            var h = [len / handleOffsetDiv[0], len / handleOffsetDiv[1]];
            var pts = p.pathPoints;

            // Arc direction: Up (default) or Down
            var dirY = 1;
            try {
                // Only available after UI exists; default to Up.
                if (typeof rbArcDown !== 'undefined' && rbArcDown && rbArcDown.value) dirY = -1;
            } catch (_) { dirY = 1; }

            pts[0].rightDirection = [pts[0].rightDirection[0] + h[0], pts[0].rightDirection[1] + (dirY * h[1])];
            pts[1].leftDirection = [pts[1].leftDirection[0] - h[0], pts[1].leftDirection[1] + (dirY * h[1])];

            return p;
        } catch (e) {
            return null;
        }
    }

    // Create a circle path whose diameter equals the full text width (using outline bounds)
    function createCirclePathFromText(originalText, layer) {
        // Guard
        try {
            if (!originalText || originalText.typename !== 'TextFrame') return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;
        } catch (_) {
            return null;
        }

        try {
            // Use outline bounds for stability
            var dummy = originalText.duplicate();
            var ol = null;
            try { ol = dummy.createOutline(); } catch (_) { ol = null; }
            try { dummy.remove(); } catch (_) { }
            if (!ol) return null;

            var b = ol.geometricBounds; // [L, T, R, B]
            try { ol.remove(); } catch (_) { }

            var w = b[2] - b[0];
            if (!w || isNaN(w) || w <= 0) return null;

            // Center of text outline
            var cx = (b[0] + b[2]) / 2;
            var cy = (b[1] + b[3]) / 2;

            var d = w; // diameter = text width
            var left = cx - (d / 2);
            var top = cy + (d / 2);

            var p = layer.pathItems.ellipse(top, left, d, d);

            // Default invisible (caller will set preview style if needed)
            try {
                p.stroked = false;
                p.filled = false;
            } catch (_) { }

            return p;
        } catch (e) {
            return null;
        }
    }

    // Resolve which base path to use for each text
    // - If paths are selected: use them (1-to-1 if counts match, otherwise first)
    // - If no path is selected but target is PathText: use its own textPath
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

    // Make ORIGINAL selected path invisible (no fill / no stroke)
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

    // Resolve actual PathItem to duplicate (CompoundPathItem -> first pathItem)
    function resolveSourcePath(basePath) {
        try {
            return (basePath.typename === 'CompoundPathItem') ? basePath.pathItems[0] : basePath;
        } catch (_) {
            return null;
        }
    }

    // Duplicate a path for the text and place it on the same layer
    function duplicatePathForText(basePath, currentLayer) {
        var srcPath = resolveSourcePath(basePath);
        if (!srcPath) return null;
        try {
            var dup = srcPath.duplicate();
            try { dup.move(currentLayer, ElementPlacement.PLACEATBEGINNING); } catch (_) { }
            return dup;
        } catch (eDup) {
            return null;
        }
    }

    // Get target point-text items (recursive)
    // Get target text items (recursive)
    function getTargetTextItems(items) {
        var ti = [];
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.typename === 'TextFrame') {
                try {
                    if (it.kind === TextType.POINTTEXT || it.kind === TextType.PATHTEXT) {
                        ti.push(it);
                    }
                } catch (_) { }
            } else if (it.typename === 'GroupItem') {
                ti = ti.concat(getTargetTextItems(it.pageItems));
            }
        }
        return ti;
    }

    // Get selected path items (PathItem). If GroupItem contains paths, collect them too.
    function getSelectedPathItems(items) {
        var pi = [];
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.typename === 'PathItem') {
                pi.push(it);
            } else if (it.typename === 'CompoundPathItem') {
                pi.push(it);
            } else if (it.typename === 'GroupItem') {
                pi = pi.concat(getSelectedPathItems(it.pageItems));
            }
        }
        return pi;
    }

    // Fix overset text by shrinking font size (targets only)
    // NOTE: Illustrator TextFrame has no reliable `editable` flag; avoid filtering by it.
    function fixOversetTextFrames(frames) {
        if (!frames || frames.length === 0) return false;

        var step = 0.1;
        var maxIter = 2000;

        function isOversetTF(tf) {
            try {
                if (!tf) return false;
                if (!tf.lines || tf.lines.length === 0) {
                    // If there are characters but no lines, treat as overset
                    return (tf.characters && tf.characters.length > 0);
                }
                var visible = 0;
                for (var i = 0; i < tf.lines.length; i++) {
                    try { visible += tf.lines[i].characters.length; } catch (_) { }
                }
                var total = 0;
                try { total = tf.characters.length; } catch (_) { total = 0; }
                return (visible < total);
            } catch (_) {
                return false;
            }
        }

        function shrinkUntilFit(tf) {
            try {
                if (!tf) return;
                // Only AreaText / PathText can overset in Illustrator
                try {
                    if (!(tf.kind === TextType.AREATEXT || tf.kind === TextType.PATHTEXT)) return;
                } catch (_) { return; }

                // Skip locked/hidden objects
                try { if (tf.locked) return; } catch (_) { }
                try { if (tf.hidden) return; } catch (_) { }
                try { if (tf.layer && tf.layer.locked) return; } catch (_) { }
                try { if (tf.layer && !tf.layer.visible) return; } catch (_) { }

                var iter = 0;
                while (isOversetTF(tf) && iter < maxIter) {
                    iter++;
                    try {
                        var cur = tf.textRange.characterAttributes.size;
                        if (cur <= step) break;
                        tf.textRange.characterAttributes.size = cur - step;
                    } catch (_) {
                        break;
                    }
                }
            } catch (_) { }
        }

        for (var i = 0; i < frames.length; i++) {
            shrinkUntilFit(frames[i]);
        }
        return true;
    }

}());