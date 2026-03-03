/*
### スクリプト名：
ポイント文字をパス上文字に変換

### 更新日：
20260302

### 概要：
選択中の「ポイント文字」を、パス上文字に変換します。
- パスが一緒に選択されている場合：そのパスを複製して使用（複数テキストにも対応）
- パスが選択されていない場合：アーチ状のパスを自動生成して使用（テキストのみ選択でも実行可能）
*/

(function () {

    // Constant
    const SCRIPT_TITLE = 'パス上文字に変換';
    const SCRIPT_VERSION = '0.8.0';

    // Get items
    if (app.documents.length === 0) {
        alert('ドキュメントが開かれていません');
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

        __hasSelectedPath = (selectedPaths && selectedPaths.length > 0);
        if (!__hasSelectedPath) {
            try { rbToPathText.enabled = false; } catch (_) { }
            rbToPathText.value = false;
            rbGenArcPath.value = true;
        } else {
            try { rbToPathText.enabled = true; } catch (_) { }
        }
    }

    // Validation
    if (targetItems.length === 0) {
        alert('対象のポイント文字が見つかりません');
        return false;
    }

    // Show dialog
    var dlg = new Window('dialog', SCRIPT_TITLE + '  ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = [15, 20, 15, 10];

    // 2-column layout
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
    var pnlOpt = leftCol.add('panel', undefined, '処理');
    pnlOpt.orientation = 'column';
    pnlOpt.alignChildren = ['left', 'top'];
    pnlOpt.margins = [15, 20, 15, 10];

    var rbToPathText = pnlOpt.add('radiobutton', undefined, 'パス上テキストに');
    var rbGenArcPath = pnlOpt.add('radiobutton', undefined, 'アーチ状のパスを生成');
    rbToPathText.value = true;

    // Auto-select arc generation when no path is selected
    var __hasSelectedPath = (selectedPaths && selectedPaths.length > 0);
    if (!__hasSelectedPath) {
        rbToPathText.value = false;
        rbGenArcPath.value = true;
        // Prevent accidental selection of path mode when no path is selected
        try { rbToPathText.enabled = false; } catch (_) { }
    }

    // Option panel
    var pnlOption = leftCol.add('panel', undefined, 'オプション');
    pnlOption.orientation = 'column';
    pnlOption.alignChildren = ['left', 'top'];
    pnlOption.margins = [15, 20, 15, 10];

    var cbReverse = pnlOption.add('checkbox', undefined, 'パスの反対側に');
    cbReverse.value = false;

    var cbFixOverset = pnlOption.add('checkbox', undefined, '文字あふれを解消');
    cbFixOverset.value = false;

    // Effect panel
    var pnlEffect = rightCol.add('panel', undefined, '効果');
    pnlEffect.orientation = 'column';
    pnlEffect.alignChildren = ['left', 'top'];
    pnlEffect.margins = [15, 20, 15, 10];

    var rbEffectRainbow = pnlEffect.add('radiobutton', undefined, '虹');
    var rbEffectDistort = pnlEffect.add('radiobutton', undefined, '歪み');
    var rbEffectRibbon = pnlEffect.add('radiobutton', undefined, '3D リボン');
    var rbEffectStep = pnlEffect.add('radiobutton', undefined, '階段');
    var rbEffectGravity = pnlEffect.add('radiobutton', undefined, '引力');

    // No effect selected by default
    rbEffectRainbow.value = false;
    rbEffectDistort.value = false;
    rbEffectRibbon.value = false;
    rbEffectStep.value = false;
    rbEffectGravity.value = false;

    // Position panel
    var pnlPosition = rightCol.add('panel', undefined, 'パス上の位置');
    pnlPosition.orientation = 'column';
    pnlPosition.alignChildren = ['left', 'top'];
    pnlPosition.margins = [15, 20, 15, 10];

    var rbPosAscender = pnlPosition.add('radiobutton', undefined, 'アセンダ');
    var rbPosDescender = pnlPosition.add('radiobutton', undefined, 'ディセンダ');
    var rbPosCenter = pnlPosition.add('radiobutton', undefined, '中央');
    var rbPosBaseline = pnlPosition.add('radiobutton', undefined, '欧文ベースライン');

    // Default: 欧文ベースライン
    rbPosBaseline.value = true;

    // Footer (Preview independent) + Buttons (OK on the right)
    var footer = dlg.add('group');
    footer.orientation = 'row';
    footer.alignChildren = ['fill', 'center'];
    footer.alignment = ['fill', 'top'];

    var leftFooter = footer.add('group');
    leftFooter.orientation = 'row';
    leftFooter.alignChildren = ['left', 'center'];
    leftFooter.alignment = ['left', 'center'];

    var cbPreview = leftFooter.add('checkbox', undefined, 'プレビュー');
    cbPreview.value = true;

    var rightFooter = footer.add('group');
    rightFooter.orientation = 'row';
    rightFooter.alignChildren = ['right', 'center'];
    rightFooter.alignment = ['right', 'center'];


    var btnCancel = rightFooter.add('button', undefined, 'キャンセル');
    var btnOk = rightFooter.add('button', undefined, 'OK', { name: 'ok' });

    // -------------------------------------------------
    // Preview (no undo): create temporary objects + hide originals
    // -------------------------------------------------
    var __previewTempItems = [];        // items created during preview (TextFrame etc.)
    var __previewHiddenOriginals = [];  // original items hidden during preview

    var __previewPathStates = [];       // { item: PathItem, stroked: Boolean, filled: Boolean }

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
                __previewPathStates[p].item.stroked = __previewPathStates[p].stroked;
                __previewPathStates[p].item.filled = __previewPathStates[p].filled;
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
                filled: pathItem.filled
            });
        } catch (_) { }
    }

    function __preview_makePathNoFillNoStroke(basePath) {
        try {
            if (!basePath) return;

            if (basePath.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < basePath.pathItems.length; cp++) {
                    var pi = basePath.pathItems[cp];
                    __preview_recordPathState(pi);
                    try { pi.stroked = false; } catch (_) { }
                    try { pi.filled = false; } catch (_) { }
                }
            } else {
                __preview_recordPathState(basePath);
                try { basePath.stroked = false; } catch (_) { }
                try { basePath.filled = false; } catch (_) { }
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
        if (rbToPathText.value) {
            if (!selectedPaths || selectedPaths.length === 0) {
                // Auto-fallback to arc generation when no path is selected
                rbToPathText.value = false;
                rbGenArcPath.value = true;
                mainProcessGenerateArcPath(false, true);
                return;
            }
            mainProcess(false, true);
        } else if (rbGenArcPath.value) {
            mainProcessGenerateArcPath(false, true);
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

    rbToPathText.onClick = refreshPreviewIfNeeded;
    rbGenArcPath.onClick = refreshPreviewIfNeeded;
    cbReverse.onClick = refreshPreviewIfNeeded;
    cbFixOverset.onClick = refreshPreviewIfNeeded;
    rbEffectRainbow.onClick = refreshPreviewIfNeeded;
    rbEffectDistort.onClick = refreshPreviewIfNeeded;
    rbEffectRibbon.onClick = refreshPreviewIfNeeded;
    rbEffectStep.onClick = refreshPreviewIfNeeded;
    rbEffectGravity.onClick = refreshPreviewIfNeeded;

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

        // Apply once before closing
        if (rbToPathText.value) {
            if (!selectedPaths || selectedPaths.length === 0) {
                // Auto-fallback to arc generation when no path is selected
                rbToPathText.value = false;
                rbGenArcPath.value = true;
                mainProcessGenerateArcPath(true, false);
            } else {
                mainProcess(true, false);
            }
        } else if (rbGenArcPath.value) {
            mainProcessGenerateArcPath(true, false);
        }

        dlg.close(1);
    };

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

    // Main process
    function mainProcess(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;

        for (var j = 0; j < targetItems.length; j++) {
            // Get original text frame item & current layer
            var originalText = targetItems[j];
            var currentlayer = originalText.layer;

            // Resolve base path + duplicate path for this text
            var basePath = resolveBasePathForText(selectedPaths, targetItems.length, j);
            if (!basePath) {
                // Should not happen because of Validation, but keep defensive
                if (showAlerts) alert('パスを一緒に選択してください');
                return;
            }

            if (!previewMode) {
                makeOriginalPathInvisible(basePath);
            } else {
                // Preview: show original selected path as no fill / no stroke (restore on preview clear)
                __preview_makePathNoFillNoStroke(basePath);
            }
            var textPath = duplicatePathForText(basePath, currentlayer);
            if (!textPath) {
                if (showAlerts) alert('パスの複製に失敗しました');
                continue;
            }

            if (previewMode) {
                __previewTempItems.push(textPath);
            }

            // Create Text on a path
            var textOnAPath = currentlayer.textFrames.pathText(textPath);
            // Keep stacking position (avoid appearing to disappear behind other objects)
            try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (_) { }
            if (previewMode) {
                __previewTempItems.push(textOnAPath);
            }

            // Reverse side option
            if (cbReverse.value) {
                try {
                    textOnAPath.textPath.polarity = PolarityValues.NEGATIVE;
                } catch (_) { }
            }

            applyPathTextEffect(textOnAPath, previewMode);

            // Duplicate textrange from original text frame item
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

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
            if (!previewMode && cbFixOverset.value) {
                try {
                    DealWithOversetText_SingleLine();
                } catch (_) { }
            }
        }
    }

    // Main process (generate arc path)
    function mainProcessGenerateArcPath(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentlayer = originalText.layer;

            // Create an arc-like path from the text bounds (reference logic)
            var textPath = createArcPathFromText(originalText, currentlayer);
            if (!textPath) {
                if (showAlerts) alert('アーチ状のパス生成に失敗しました');
                continue;
            }

            if (previewMode) {
                __previewTempItems.push(textPath);
            }

            // Create Text on a path
            var textOnAPath = currentlayer.textFrames.pathText(textPath);
            // Keep stacking position (avoid appearing to disappear behind other objects)
            try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (_) { }
            if (previewMode) {
                __previewTempItems.push(textOnAPath);
            }

            // Reverse side option
            if (cbReverse.value) {
                try {
                    textOnAPath.textPath.polarity = PolarityValues.NEGATIVE;
                } catch (_) { }
            }

            applyPathTextEffect(textOnAPath, previewMode);

            // Duplicate textrange from original text frame item
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

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
            if (!previewMode && cbFixOverset.value) {
                try {
                    DealWithOversetText_SingleLine();
                } catch (_) { }
            }
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
            pts[0].rightDirection = [pts[0].rightDirection[0] + h[0], pts[0].rightDirection[1] + h[1]];
            pts[1].leftDirection = [pts[1].leftDirection[0] - h[0], pts[1].leftDirection[1] + h[1]];

            return p;
        } catch (e) {
            return null;
        }
    }

    // Resolve which selected path to use for each text
    function resolveBasePathForText(selectedPaths, textCount, index) {
        if (!selectedPaths || selectedPaths.length === 0) return null;
        if (selectedPaths.length === textCount) return selectedPaths[index];
        return selectedPaths[0];
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
    function getTargetTextItems(items) {
        var ti = [];
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === 'TextFrame' && items[i].kind === TextType.POINTTEXT) {
                ti.push(items[i]);
            } else if (items[i].typename === 'GroupItem') {
                ti = ti.concat(getTargetTextItems(items[i].pageItems));
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

    function DealWithOversetText_SingleLine() {
        var defaultSizeTagName = "overset_text_default_size";
        var defaultIncrement = 0.1;

        function recordFontSizeInTag(size, art) {
            var tag;
            var tags = art.tags;
            try {
                tag = tags.getByName(defaultSizeTagName);
                tag.value = size;
            } catch (e) {
                tag = tags.add();
                tag.name = defaultSizeTagName;
                tag.value = size;
            }
        };

        function readFontSizeFromTag(art) {
            var tag;
            var tags = art.tags;
            try {
                tag = tags.getByName(defaultSizeTagName);
                return tag.value * 1;
            } catch (e) {
                return null;
            }
        };

        function recordFontSize() {
            var doc = app.activeDocument;
            for (var i = 0; i < doc.textFrames.length; i++) {
                var t = doc.textFrames[i];
                if ((t.kind == TextType.AREATEXT || t.kind == TextType.PATHTEXT) && t.editable && !t.locked && !t.hidden) {
                    recordFontSizeInTag(t.textRange.characterAttributes.size, t);
                }
            };
        };

        function isOverset(textBox, lineAmt) {
            if (textBox.lines.length > 0) {
                var charactersOnVisibleLines = 0;
                if (typeof (lineAmt) != "undefined") {
                    lineAmt = 1;
                }
                for (var i = 0; i < lineAmt; i++) {
                    charactersOnVisibleLines += textBox.lines[i].characters.length;
                }
                if (charactersOnVisibleLines < textBox.characters.length) {
                    return true;
                } else {
                    return false;
                }
            } else if (textBox.characters.length > 0) {
                return true;
            }
        };

        function shrinkFont(textBox) {
            var totalCharCount = textBox.characters.length;
            var lineCount = textBox.lines.length;
            if (textBox.lines.length > 0) {
                if (isOverset(textBox, lineCount)) {
                    var inc = defaultIncrement;
                    while (isOverset(textBox, lineCount)) {
                        textBox.textRange.characterAttributes.size -= inc;
                    }
                }
            } else if (textBox.characters.length > 0) {
                var inc = defaultIncrement;
                while (isOverset(textBox, lineCount)) {
                    textBox.textRange.characterAttributes.size -= inc;
                }
            }
        };

        function resetSize(textAreaBox) {
            var t = textAreaBox;
            if (t.contents != "") {
                var size = readFontSizeFromTag(t);
                if (size != null) {
                    t.textRange.characterAttributes.size = size;
                }
            }
        };

        function removeTagsOnText() {
            var doc = app.activeDocument;
            for (var i = 0; i < doc.textFrames.length; i++) {
                try {
                    doc.textFrames[i].tags.getByName(defaultSizeTagName).remove();
                } catch (e) { }
            }
        };

        function resetAllTextBoxes() {
            var doc = app.activeDocument;
            for (var i = 0; i < doc.textFrames.length; i++) {
                var t = doc.textFrames[i];
                if ((t.kind == TextType.AREATEXT || t.kind == TextType.PATHTEXT) && t.editable && !t.locked && !t.hidden) {
                    resetSize(t);
                }
            };
        };

        function shrinkAllTextBoxes() {
            var doc = app.activeDocument;
            for (var i = 0; i < doc.textFrames.length; i++) {
                var t = doc.textFrames[i];
                if ((t.kind == TextType.AREATEXT || t.kind == TextType.PATHTEXT) && t.editable && !t.locked && !t.hidden) {
                    shrinkFont(t);
                }
            };
        };

        if (app.documents.length > 0) {
            var doc = app.activeDocument;
            if (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[0]) {
                recordFontSize();
            }
            resetAllTextBoxes();
            shrinkAllTextBoxes();
            if (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[doc.dataSets.length - 1]) {
                removeTagsOnText();
            }
        }
        return true;
    }
}());