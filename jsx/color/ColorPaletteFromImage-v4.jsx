#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/**********************************************************

ColorPaletteFromImage.jsx

DESCRIPTION

選択した配置画像／ラスター画像（またはベクター）からカラーパレットを作成します。

- ラスター/配置画像: 複製→画像トレース→拡張で色を抽出
- ベクター: 選択オブジェクトから直接カラーを抽出（ラスタライズ不要）
- 16色／11色／8色／5色のパレットを、元画像の下に正方形で出力
  - 16色は画像幅にフィット、他の行も左右端を画像に揃えて出力
- 各行（16/11/8/5色）ごとに全色から面積重み付き最大距離法で代表色を選択
  - 面積が大きい色ほど選ばれやすい（pow 0.75で重み付け）
- 色の並びは最近傍法でグラデーション風にソート
- 5色はHEX、5色（CMYK補正）はCMYK表示（任意）
  - CMYK補正は各C/M/Y/Kを+5%し、10刻みで丸め

初回に色が取得できたタイミングでダイアログを表示し、
出力する行（16/11/8/5/5補正）と、カラー情報（HEX/CMYK）を選択できます。
ダイアログ表示中はプレビューが更新されます。
OK確定後にスウォッチグループを作成します（キャンセル時は作成しません）。

更新日: 2026-03-05

**********************************************************/


var SCRIPT_VERSION = "v1.5";

var __DIALOG_BOUNDS_OUTPUT__ = null; // session-only dialog position memory

/* ロケール判定 / Detect locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    noDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    noSelection: {
        ja: "対象のオブジェクトが選択されていません。",
        en: "No applicable objects are selected."
    },
    group16Name: {
        ja: "16色",
        en: "16 Colors"
    },
    colorsSuffix: {
        ja: "色",
        en: "Colors"
    },
    itemPrefix: {
        ja: "項目_",
        en: "item_"
    },
    unknownColorName: {
        ja: "色",
        en: "Color"
    },
    progressTitle: {
        ja: "処理中",
        en: "Processing"
    },
    progressPreparing: {
        ja: "準備中…",
        en: "Preparing…"
    },
    progressItem: {
        ja: "{0}/{1} を処理中",
        en: "Processing {0}/{1}"
    },
    progressDone: {
        ja: "完了",
        en: "Done"
    },
    // --- Output Options Dialog ---
    dialogTitle: {
        ja: "カラーパレット作成",
        en: "Create Color Palette"
    },
    btnOK: {
        ja: "OK",
        en: "OK"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    opt16: {
        ja: "16色",
        en: "16"
    },
    opt11: {
        ja: "11色",
        en: "11"
    },
    opt8: {
        ja: "8色",
        en: "8"
    },
    opt5: {
        ja: "5色",
        en: "5"
    },
    opt5Adj: {
        ja: "5色（CMYK補正）",
        en: "5 (CMYK Adjust)"
    },
    optHEX: {
        ja: "HEX",
        en: "HEX"
    },
    optCMYK: {
        ja: "CMYK",
        en: "CMYK"
    },
    optPreview: {
        ja: "プレビュー",
        en: "Preview"
    },
    optInfoAll: {
        ja: "すべての情報",
        en: "All info"
    },
    optInfo5Only: {
        ja: "5色のみ",
        en: "Only 5"
    },
    cmykPrefix: {
        ja: "CMYK: ",
        en: "CMYK: "
    },
    panelCountsTitle: {
        ja: "色数",
        en: "Counts"
    },
    panelColorInfoTitle: {
        ja: "カラー情報",
        en: "Color Info"
    },
    msgNoRow: {
        ja: "出力する行が選ばれていません。",
        en: "No rows selected to output."
    }
};

/* ラベル取得関数 / Get localized label */
function L(key) {
    return LABELS[key][lang];
}

function LF(key, a, b) {
    var s = L(key);
    if (a !== undefined) s = s.replace("{0}", a);
    if (b !== undefined) s = s.replace("{1}", b);
    return s;
}

// --- Output Options Dialog ---
function showOutputOptionsDialog(onPreviewChange) {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.alignChildren = 'fill';
    dlg.margins = 15;

    // Restore dialog position during this session
    try {
        if (__DIALOG_BOUNDS_OUTPUT__) dlg.bounds = __DIALOG_BOUNDS_OUTPUT__;
    } catch (eB) { }

    // Update stored bounds when moved/resized
    dlg.onMove = dlg.onResize = function () {
        try { __DIALOG_BOUNDS_OUTPUT__ = dlg.bounds; } catch (eM) { }
    };

    // Info scope (UI)
    var infoModeRow = dlg.add('group');
    infoModeRow.alignment = 'center';
    infoModeRow.orientation = 'row';
    infoModeRow.alignChildren = 'center';

    var rbAllInfo = infoModeRow.add('radiobutton', undefined, L('optInfoAll'));
    var rb5Only = infoModeRow.add('radiobutton', undefined, L('optInfo5Only'));

    // Two-column layout
    var cols = dlg.add('group');
    cols.orientation = 'row';
    cols.alignChildren = 'fill';
    cols.alignment = 'fill';
    cols.spacing = 12;

    var colL = cols.add('group');
    colL.orientation = 'column';
    colL.alignChildren = 'fill';

    var colR = cols.add('group');
    colR.orientation = 'column';
    colR.alignChildren = 'fill';

    var pnl = colL.add('panel', undefined, L('panelCountsTitle'));
    pnl.alignChildren = 'left';
    pnl.margins = [15, 20, 15, 10];

    var cb16 = pnl.add('checkbox', undefined, L('opt16'));
    var cb11 = pnl.add('checkbox', undefined, L('opt11'));
    var cb8 = pnl.add('checkbox', undefined, L('opt8'));
    var cb5 = pnl.add('checkbox', undefined, L('opt5'));
    var cb5a = pnl.add('checkbox', undefined, L('opt5Adj'));

    var pnlInfo = colR.add('panel', undefined, L('panelColorInfoTitle'));
    pnlInfo.alignment = 'fill';
    pnlInfo.alignChildren = 'left';
    pnlInfo.margins = [15, 20, 15, 10];

    var cbHEX = pnlInfo.add('checkbox', undefined, L('optHEX'));
    var cbCMYK = pnlInfo.add('checkbox', undefined, L('optCMYK'));

    // Preview (UI)
    var previewRow = dlg.add('group');
    previewRow.alignment = 'center';
    var cbPreview = previewRow.add('checkbox', undefined, L('optPreview'));
    cbPreview.value = true;

    // Defaults: all ON
    rbAllInfo.value = true;
    rb5Only.value = false;

    cb16.value = true;
    cb11.value = true;
    cb8.value = true;
    cb5.value = true;
    cb5a.value = true;
    cbHEX.value = true;
    cbCMYK.value = true;
    cbPreview.value = true;

    // Apply dependent dimming rules without triggering preview twice
    updateInfoDims(false);

    function getCurrentOptions() {
        return {
            out16: cb16.value,
            out11: cb11.value,
            out8: cb8.value,
            out5: cb5.value,
            out5Adj: cb5a.value,
            infoMode: rb5Only.value ? '5only' : 'all',
            preview: cbPreview.value,
            showHEX: cbHEX.value,
            showCMYK: cbCMYK.value
        };
    }

    function notifyPreviewChange() {
        try {
            if (onPreviewChange) onPreviewChange(getCurrentOptions());
        } catch (eNP) { }
    }

    function updateInfoDims(doNotify) {
        cbHEX.enabled = cb5.value;
        if (!cb5.value) cbHEX.value = false;

        cbCMYK.enabled = cb5a.value;
        if (!cb5a.value) cbCMYK.value = false;

        if (doNotify !== false) notifyPreviewChange();
    }

    function updateInfoMode() {
        if (rb5Only.value) {
            // 5色のみ → 16/11/8 をOFFにするだけ（ディムしない）
            cb16.value = false;
            cb11.value = false;
            cb8.value = false;
        } else {
            // "All info" -> turn everything ON
            cb16.enabled = true;
            cb11.enabled = true;
            cb8.enabled = true;

            cb16.value = true;
            cb11.value = true;
            cb8.value = true;
            cb5.value = true;
            cb5a.value = true;
            cbHEX.value = true;
            cbCMYK.value = true;
            cbPreview.value = true;

            updateInfoDims(false);
        }

        notifyPreviewChange();
    }

    // Wire events
    rbAllInfo.onClick = updateInfoMode;
    rb5Only.onClick = updateInfoMode;

    cb16.onClick = notifyPreviewChange;
    cb11.onClick = notifyPreviewChange;
    cb8.onClick = notifyPreviewChange;

    cb5.onClick = function () {
        updateInfoDims(false);
        // If 5色 is turned ON and CMYK option is available, auto-enable CMYK
        if (cb5.value && cbCMYK.enabled && !cbCMYK.value) {
            cbCMYK.value = true;
        }
        notifyPreviewChange();
    };
    cb5a.onClick = function () { updateInfoDims(); };

    cbHEX.onClick = notifyPreviewChange;
    cbCMYK.onClick = notifyPreviewChange;
    cbPreview.onClick = notifyPreviewChange;

    updateInfoDims();
    updateInfoMode();
    notifyPreviewChange();

    // Buttons (center-aligned)
    var btnRow = dlg.add('group');
    btnRow.alignment = 'center';
    btnRow.margins = [0, 10, 0, 0];

    var btns = btnRow.add('group');
    btns.alignment = 'center';
    var ng = btns.add('button', undefined, L('btnCancel'), { name: 'cancel' });
    var ok = btns.add('button', undefined, L('btnOK'), { name: 'ok' });

    ok.onClick = function () {
        if (!cb16.value && !cb11.value && !cb8.value && !cb5.value && !cb5a.value) {
            alert(L('msgNoRow'));
            return;
        }
        try { __DIALOG_BOUNDS_OUTPUT__ = dlg.bounds; } catch (eM2) { }
        dlg.close(1);
    };
    ng.onClick = function () {
        try { __DIALOG_BOUNDS_OUTPUT__ = dlg.bounds; } catch (eM3) { }
        dlg.close(0);
    };

    // Center only when no stored bounds (so session position memory works)
    try {
        if (!__DIALOG_BOUNDS_OUTPUT__) dlg.center();
    } catch (e) { }

    var res = dlg.show();
    try { __DIALOG_BOUNDS_OUTPUT__ = dlg.bounds; } catch (eM4) { }
    if (res !== 1) return null;

    return getCurrentOptions();
}

/* メインコード / Main code */

if (app.documents.length === 0) {
    alert(L('noDocument'));
} else {
    var doc = app.activeDocument;
    var sel = doc.selection;

    /* 選択中のオブジェクトを収集（配置画像＋ベクター） / Collect selected items (placed + vector) */
    var rasterItems = [];
    var vectorItems = [];
    for (var s = 0; s < sel.length; s++) {
        if (sel[s].typename === "PlacedItem" || sel[s].typename === "RasterItem") {
            rasterItems.push(sel[s]);
        } else if (sel[s].typename === "PathItem" || sel[s].typename === "CompoundPathItem" || sel[s].typename === "GroupItem") {
            vectorItems.push(sel[s]);
        }
    }

    /* 処理対象リストを構築 / Build process list */
    // NOTE: Separate "what we process" from "where we place palettes".
    // Raster/Placed items are processed one-by-one.
    // Vector selection is processed as a single grouped job, while palette placement uses the first vector as an anchor.
    var jobs = []; // { type: "raster"|"vector", originalItem: PageItem }

    /* ラスター/配置画像は個別に処理 / Process raster/placed items individually */
    for (var s = 0; s < rasterItems.length; s++) {
        jobs.push({ type: "raster", originalItem: rasterItems[s] });
    }

    /* ベクターは「まとめて1回」処理（現状挙動のまま）/ Vectors are processed once as a group (keep current behavior) */
    if (vectorItems.length > 0) {
        jobs.push({ type: "vector", originalItem: vectorItems[0] }); /* パレット配置の基準 / Anchor for palette position */
    }

    if (jobs.length === 0) {
        alert(L('noSelection'));
    } else {
        var tracingPresets = app.tracingPresetsList;
        // Safe preset selection: prefer index 1 (current behavior), fallback to index 0, or skip if none.
        var tracingPresetName = null;
        if (tracingPresets && tracingPresets.length) {
            tracingPresetName = tracingPresets[(tracingPresets.length > 1) ? 1 : 0];
        }

        // --- Progress UI (ScriptUI) ---
        function createProgress(maxValue) {
            var w = new Window('palette', L('progressTitle'));
            w.alignChildren = 'fill';
            w.margins = 15;
            var txt = w.add('statictext', undefined, L('progressPreparing'));
            txt.characters = 28;
            var bar = w.add('progressbar', undefined, 0, Math.max(1, maxValue));
            bar.preferredSize = [260, 14];
            w.show();
            w.update();
            return {
                win: w,
                text: txt,
                bar: bar,
                set: function (value, label) {
                    try {
                        if (label) txt.text = label;
                        bar.value = Math.max(0, Math.min(bar.maxvalue, value));
                        w.update();
                        app.redraw();
                    } catch (e) { }
                },
                close: function () {
                    try { w.close(); } catch (e) { }
                }
            };
        }

        var progress = createProgress(jobs.length);
        progress.set(0, L('progressPreparing'));

        /* 作業用レイヤーを作成 / Create work layer */
        var workLayer = doc.layers.add();
        workLayer.name = "__workLayer__"; // workLayer

        try {

            /* 選択したオブジェクトを複製し、作業用レイヤーに移動 / Duplicate selected items to work layer */
            // itemsToProcess holds the work item plus its original reference for naming and palette placement.
            var itemsToProcess = []; // { type: "raster"|"vector", originalItem: PageItem, workItem: PageItem }

            for (var i = 0; i < jobs.length; i++) {
                if (jobs[i].type === "vector") {
                    /* ベクターは全アイテムを複製し、作業用レイヤー上でグループ化 / Duplicate all vectors and group on work layer */
                    var vecGroup = workLayer.groupItems.add();
                    for (var v = vectorItems.length - 1; v >= 0; v--) {
                        var vdup = vectorItems[v].duplicate();
                        vdup.move(vecGroup, ElementPlacement.PLACEATEND);
                    }
                    itemsToProcess.push({ type: "vector", originalItem: jobs[i].originalItem, workItem: vecGroup });
                } else {
                    var dup = jobs[i].originalItem.duplicate();
                    dup.move(workLayer, ElementPlacement.PLACEATEND);
                    itemsToProcess.push({ type: "raster", originalItem: jobs[i].originalItem, workItem: dup });
                }
            }

            // --- Helper: Trace and expand (raster/placed only) ---
            function traceAndExpand(itemToTrace) {
                var t = itemToTrace.trace();
                if (tracingPresetName) {
                    try {
                        t.tracing.tracingOptions.loadFromPreset(tracingPresetName);
                    } catch (e) {
                        // If preset load fails (e.g., missing/unsupported), continue with current tracing options.
                    }
                }
                app.redraw();
                var expanded = t.tracing.expandTracing();
                try {
                    if (expanded && expanded.typename && workLayer) {
                        expanded.move(workLayer, ElementPlacement.PLACEATEND);
                    }
                } catch (e) { }
                return expanded;
            }

            /* 各複製に対してトレース→拡張→スウォッチ登録 / Trace, expand, and register swatches */
            var outOpt = null; // decided after first successful extraction
            var canceled = false;
            for (var j = 0; j < itemsToProcess.length; j++) {
                try {
                    var job = itemsToProcess[j];
                    progress.set(j, LF('progressItem', j + 1, itemsToProcess.length));
                    var p = job.workItem;
                    var colors;
                    if (job.type === "vector") {
                        // ベクター: オブジェクトから直接カラーを抽出（ラスタライズ不要）
                        colors = collectFillColors(p, []);
                        colors = deduplicateColors(colors);
                    } else {
                        var grp = traceAndExpand(p);
                        colors = collectFillColors(grp, []);
                        colors = deduplicateColors(colors);
                    }

                    // 2. On first extraction, show dialog, draw preview from colors (not swatchGroup)
                    if (outOpt === null) {
                        if (colors && colors.length > 0) {
                            var outAll = { out16: true, out11: true, out8: true, out5: true, out5Adj: true, showHEX: true, showCMYK: true };
                            var previewGroup = null;
                            try {
                                previewGroup = job.originalItem.layer.groupItems.add();
                                previewGroup.name = "__ColorPalettePreview__";
                            } catch (ePg) { }

                            try {
                                if (previewGroup) {
                                    drawSwatchSquares(doc, job.originalItem, colors, outAll, previewGroup);
                                }
                            } catch (ePrev) { }

                            // Ensure preview is rendered before opening the modal dialog
                            try { app.redraw(); } catch (eRd) { }
                            try { $.sleep(80); } catch (eSl) { }

                            // Hide progress UI while dialog shown
                            try { if (progress && progress.win) progress.win.hide(); } catch (eHide) { }
                            outOpt = showOutputOptionsDialog(function (optNow) {
                                try {
                                    if (!previewGroup) return;
                                    // Remove all preview items
                                    while (previewGroup.pageItems.length) {
                                        try { previewGroup.pageItems[0].remove(); } catch (eRm) { break; }
                                    }
                                    // Draw preview again if enabled
                                    if (optNow && optNow.preview) {
                                        drawSwatchSquares(doc, job.originalItem, colors, optNow, previewGroup);
                                    }
                                    try { app.redraw(); } catch (eRd2) { }
                                } catch (eCb) { }
                            });
                            try { if (progress && progress.win) progress.win.show(); } catch (eShow) { }

                            if (!outOpt) {
                                // Canceled: remove preview and stop
                                try { if (previewGroup) previewGroup.remove(); } catch (eRmPrev) { }
                                canceled = true;
                            } else {
                                // OK: keep preview visible until final output is ready

                                // Now create swatch group and register colors
                                var groupName;
                                if (job.originalItem.typename === "PlacedItem" && job.originalItem.file) {
                                    groupName = job.originalItem.file.name;
                                } else {
                                    groupName = L('itemPrefix') + j;
                                }
                                var swatchGroup = createSwatchGroupFromColors(doc, groupName, colors);
                                // Draw final output for this item
                                var finalGroup = null;
                                try {
                                    finalGroup = job.originalItem.layer.groupItems.add();
                                    finalGroup.name = "__ColorPalette__";
                                } catch (eFg) { }
                                try {
                                    if (finalGroup) {
                                        drawSwatchSquares(doc, job.originalItem, colors, outOpt, finalGroup);
                                    } else {
                                        drawSwatchSquares(doc, job.originalItem, colors, outOpt);
                                    }
                                } catch (eFinal) { }
                                // Remove preview after final output is drawn
                                try { if (previewGroup) previewGroup.remove(); } catch (eRmPrev2) { }
                            }
                        }
                    } else {
                        // For subsequent items, create swatch group immediately and proceed
                        if (!canceled && outOpt && colors && colors.length > 0) {
                            var groupName;
                            if (job.originalItem.typename === "PlacedItem" && job.originalItem.file) {
                                groupName = job.originalItem.file.name;
                            } else {
                                groupName = L('itemPrefix') + j;
                            }
                            var swatchGroup = createSwatchGroupFromColors(doc, groupName, colors);
                            var finalGroup2 = null;
                            try {
                                finalGroup2 = job.originalItem.layer.groupItems.add();
                                finalGroup2.name = "__ColorPalette__";
                            } catch (eFg2) { }
                            try {
                                if (finalGroup2) {
                                    drawSwatchSquares(doc, job.originalItem, colors, outOpt, finalGroup2);
                                } else {
                                    drawSwatchSquares(doc, job.originalItem, colors, outOpt);
                                }
                            } catch (eFinal2) { }
                        }
                    }

                    // NOTE: ここにあった旧 drawSwatchSquares 呼び出し（if (!canceled && outOpt) ...）は削除してください。

                    progress.set(j + 1, LF('progressItem', j + 1, itemsToProcess.length));
                } catch (err) {
                    // ignore per-item failures to keep processing others
                }
                if (canceled) break;
            }
            progress.set(itemsToProcess.length, L('progressDone'));

        } finally {
            /* 作業用レイヤーを削除 / Remove work layer */
            try {
                if (workLayer) workLayer.remove();
            } catch (e) { }
            try { if (progress) progress.close(); } catch (e) { }
        }

        app.redraw();
    }
}

/* グループ内のパスの塗り色をスウォッチグループに登録 / Register fill colors from group to swatch group */
function addColorsToSwatchGroup(doc, swatchGroup, item) {
    if (item.typename === "PathItem") {
        if (item.filled && item.fillColor) addSwatchToGroup(doc, swatchGroup, item.fillColor);
        return;
    }
    if (item.typename !== "GroupItem" && item.typename !== "CompoundPathItem") return;
    var children = (item.typename === "GroupItem") ? item.pageItems : item.pathItems;
    for (var k = 0; k < children.length; k++) {
        addColorsToSwatchGroup(doc, swatchGroup, children[k]);
    }
}

// Helper: Collect fill colors with area from expanded result into array
// Returns [{color: Color, area: Number}, ...]
function collectFillColors(item, out) {
    if (!out) out = [];
    if (!item) return out;
    if (item.typename === "PathItem") {
        if (item.filled && item.fillColor) {
            var c = item.fillColor;
            if (c.typename !== "NoColor" && c.typename !== "PatternColor" &&
                c.typename !== "GradientColor" && c.typename !== "SpotColor") {
                var a = 1;
                try { a = Math.abs(item.area); } catch (e) { }
                if (a < 1) a = 1;
                out.push({ color: c, area: a });
            }
        }
        return out;
    }
    if (item.typename === "GroupItem" || item.typename === "CompoundPathItem") {
        var children = (item.typename === "GroupItem") ? item.pageItems : item.pathItems;
        for (var k = 0; k < children.length; k++) {
            collectFillColors(children[k], out);
        }
    }
    return out;
}

/* スウォッチをグループに追加 / Add swatch to group */
function addSwatchToGroup(doc, swatchGroup, color) {
    try {
        if (color.typename === "SpotColor") return;
        if (color.typename === "PatternColor") return;
        if (color.typename === "GradientColor") return;
        if (color.typename === "NoColor") return;

        var swatchName = colorToName(color);

        /* 同名スウォッチが既に存在すればそれをグループに追加、なければ新規作成 / Reuse existing swatch or create new */
        var swatch;
        try {
            swatch = doc.swatches.getByName(swatchName);
        } catch (e) {
            swatch = doc.swatches.add();
            swatch.name = swatchName;
            swatch.color = color;
        }
        swatchGroup.addSwatch(swatch);
    } catch (err) {
        /* スウォッチ追加エラーは無視 / Ignore swatch add errors */
    }
}

// Helper: Create swatch group from colors array
// colors: [{color, area}, ...] or [Color, ...]
function createSwatchGroupFromColors(doc, groupName, colors) {
    var swatchGroup = doc.swatchGroups.add();
    swatchGroup.name = groupName;
    for (var i = 0; i < colors.length; i++) {
        var c = colors[i].color ? colors[i].color : colors[i];
        addSwatchToGroup(doc, swatchGroup, c);
    }
    return swatchGroup;
}

/* 元画像の下にカラーパレットを描画 / Draw color palette squares below original image */
// swatchSource: SwatchGroup or Array of {color, area}
function drawSwatchSquares(doc, originalItem, swatchSource, outOpt, containerGroup) {
    var swatches;
    // Accept SwatchGroup or Array of {color, area}
    if (swatchSource && typeof swatchSource.getAllSwatches === "function") {
        swatches = swatchSource.getAllSwatches();
    } else if (swatchSource && swatchSource.length !== undefined) {
        // treat as array of {color, area}
        swatches = [];
        for (var i = 0; i < swatchSource.length; i++) {
            swatches.push({ color: swatchSource[i].color, area: swatchSource[i].area || 1 });
        }
    } else {
        swatches = [];
    }
    var numSquares = 16;

    outOpt = outOpt || { out16: true, out11: true, out8: true, out5: true, out5Adj: true, showHEX: true, showCMYK: true };

    /* 最近傍法でソート（最も暗い色から、色距離が近い順に並べる） / Sort by nearest neighbor from darkest */
    var remaining = [];
    for (var m = 0; m < swatches.length; m++) {
        var rgb = colorToRGB(swatches[m].color);
        remaining.push({ swatch: swatches[m], r: rgb[0], g: rgb[1], b: rgb[2], area: swatches[m].area || 1 });
    }

    var colorList = [];
    if (remaining.length > 0) {
        /* 最も暗い色（輝度が低い）をスタート地点にする / Start from darkest color */
        var startIdx = 0;
        var minLum = Infinity;
        for (var q = 0; q < remaining.length; q++) {
            var lum = remaining[q].r * 0.299 + remaining[q].g * 0.587 + remaining[q].b * 0.114;
            if (lum < minLum) { minLum = lum; startIdx = q; }
        }
        colorList.push(remaining.splice(startIdx, 1)[0]);

        /* 残りから最も近い色を順に選択 / Select nearest color iteratively */
        while (remaining.length > 0) {
            var last = colorList[colorList.length - 1];
            var nearestIdx = 0;
            var nearestDist = Infinity;
            for (var q = 0; q < remaining.length; q++) {
                var dr = last.r - remaining[q].r;
                var dg = last.g - remaining[q].g;
                var db = last.b - remaining[q].b;
                var dist = dr * dr + dg * dg + db * db;
                if (dist < nearestDist) { nearestDist = dist; nearestIdx = q; }
            }
            colorList.push(remaining.splice(nearestIdx, 1)[0]);
        }
    }

    /* 元画像の位置・サイズを取得 / Get original image position and size */
    var imgLeft = originalItem.left;
    var imgTop = originalItem.top;
    var imgWidth = originalItem.width;
    var imgBottom = imgTop - originalItem.height;

    /* 正方形のサイズを計算 / Calculate square size */
    var cols16 = numSquares;      // number of columns (16)
    var cell16 = imgWidth / cols16;

    // Keep a ratio-based gap, but compute squareSize so that the row fits the image width exactly:
    // cols16*squareSize + (cols16-1)*gap = imgWidth
    var gapRatio = 0.10;          // gap is based on 16-column cell width
    var gap = cell16 * gapRatio;
    var squareSize = (imgWidth - (cols16 - 1) * gap) / cols16;

    // Vertical gap should match the column gap of a 12-column layout.
    // (Same gap rule as this script: gap = (imgWidth / N) / 10)
    var gap12 = (imgWidth / 12) / 10;
    var rowGap = gap12;

    // Gap between image bottom and the first row equals the height of a 16-color swatch.
    var startY = imgBottom - squareSize;

    /* 元画像のレイヤーに作成する / Create on original image's layer */
    var targetLayer = originalItem.layer;
    var container = containerGroup || targetLayer;

    function getLabelBlackColor() {
        try {
            if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK) {
                var c = new CMYKColor();
                c.cyan = 0; c.magenta = 0; c.yellow = 0; c.black = 100;
                return c;
            }
        } catch (e) { }
        var r = new RGBColor();
        r.red = 0; r.green = 0; r.blue = 0;
        return r;
    }

    // --- Swatch label helpers ---
    function toHex2(n) {
        var s = Math.round(n).toString(16).toUpperCase();
        return (s.length === 1) ? ("0" + s) : s;
    }
    function rgbToHex(r, g, b) {
        return "#" + toHex2(r) + toHex2(g) + toHex2(b);
    }
    function round10(v) {
        return Math.round(v / 10) * 10;
    }
    function colorToCMYKVals(color) {
        // Return [C,M,Y,K] in 0..100 (numbers). Best-effort.
        if (color && color.typename === "CMYKColor") {
            return [color.cyan, color.magenta, color.yellow, color.black];
        }
        try {
            if (app.convertSampleColor && typeof ImageColorSpace !== "undefined" && typeof ColorConvertPurpose !== "undefined") {
                if (color && color.typename === "RGBColor") {
                    var dst = app.convertSampleColor(
                        ImageColorSpace.RGB,
                        [color.red, color.green, color.blue],
                        ImageColorSpace.CMYK,
                        ColorConvertPurpose.defaultpurpose,
                        false,
                        false
                    );
                    if (dst && dst.length >= 4) return [dst[0], dst[1], dst[2], dst[3]];
                }
                if (color && color.typename === "LabColor") {
                    var dst2 = app.convertSampleColor(
                        ImageColorSpace.LAB,
                        [color.l, color.a, color.b],
                        ImageColorSpace.CMYK,
                        ColorConvertPurpose.defaultpurpose,
                        false,
                        false
                    );
                    if (dst2 && dst2.length >= 4) return [dst2[0], dst2[1], dst2[2], dst2[3]];
                }
            }
        } catch (e) { }
        return null;
    }
    function buildColorLabel(color, mode) {
        // mode: "both" | "hex" | "cmyk"
        if (!mode) mode = "both";

        var rgb = colorToRGB(color);
        var hex = rgbToHex(rgb[0], rgb[1], rgb[2]);
        var cmyk = colorToCMYKVals(color);

        if (mode === "hex") {
            return { text: hex, rgb: rgb };
        }

        if (mode === "cmyk") {
            if (cmyk) {
                var Cc = round10(cmyk[0]);
                var Mm = round10(cmyk[1]);
                var Yy = round10(cmyk[2]);
                var Kk = round10(cmyk[3]);
                return { text: L('cmykPrefix') + Cc + ", " + Mm + ", " + Yy + ", " + Kk, rgb: rgb };
            }
            // If CMYK is unavailable, fall back to HEX only
            return { text: hex, rgb: rgb };
        }

        // mode === "both"
        if (cmyk) {
            var C = round10(cmyk[0]);
            var M = round10(cmyk[1]);
            var Y = round10(cmyk[2]);
            var K = round10(cmyk[3]);
            return { text: L('cmykPrefix') + C + ", " + M + ", " + Y + ", " + K + "\r" + hex, rgb: rgb };
        }
        return { text: hex, rgb: rgb };
    }
    function addSwatchLabelToGroup(rect, size, group, mode) {
        try {
            if (!rect || !rect.filled) return;
            var info = buildColorLabel(rect.fillColor, mode);

            // Font size scales with swatch size
            var fs = size / 10;

            // Left-aligned label, aligned to the swatch width
            // Place label under the swatch: keep a gap about half the font size.
            var x = rect.left;
            var y = rect.top - rect.height - (fs / 2);

            var tf = targetLayer.textFrames.add();
            tf.contents = info.text;
            tf.textRange.justification = Justification.LEFT;
            tf.textRange.characterAttributes.size = fs;
            // Label color: always black (match document color space)
            tf.textRange.fillColor = getLabelBlackColor();

            // Font: Myriad (best-effort; depends on installed font name)
            try {
                tf.textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Regular");
            } catch (eFont) {
                try {
                    tf.textRange.characterAttributes.textFont = app.textFonts.getByName("Myriad Pro");
                } catch (eFont2) { }
            }

            tf.position = [x, y];

            // Move into group so it stays with the swatches
            try { tf.move(group, ElementPlacement.PLACEATEND); } catch (e) { }
        } catch (e) { }
    }

    /* 16色の正方形を作成・グループ化 / Create and group 16-color squares */
    if (outOpt.out16) {
        var sel16 = selectByMaxDistance(colorList, 16);
        var sorted16 = sortByNearest(sel16);
        var group16 = container.groupItems.add();
        group16.name = L('group16Name');
        for (var n = 0; n < numSquares; n++) {
            var rect = group16.pathItems.rectangle(
                startY,
                imgLeft + n * (squareSize + gap),
                squareSize,
                squareSize
            );
            rect.stroked = false;
            if (sorted16.length > 0) {
                rect.filled = true;
                rect.fillColor = sorted16[n % sorted16.length].swatch.color;
            } else {
                rect.filled = false;
            }
        }
    }

    /* 11色・8色・5色バージョン / 11, 8, 5 color rows */
    var rows = [11, 8, 5];
    var prevBottom;
    if (outOpt.out16) {
        // previous row is the 16-color row
        prevBottom = startY - squareSize;
    } else {
        // even if 16-color row is hidden, keep the same top gap (squareSize) from the image
        // so the first visible row starts at startY (imgBottom - squareSize)
        prevBottom = startY + rowGap;
    }

    for (var r = 0; r < rows.length; r++) {
        var num = rows[r];
        // Conditional output: skip if not enabled for this row
        if ((num === 11 && !outOpt.out11) || (num === 8 && !outOpt.out8) || (num === 5 && !outOpt.out5)) {
            // Do not consume vertical space if not output
            continue;
        }
        // Make each row a true square grid: size is computed to fit within imgWidth with the same horizontal gap.
        var rowSize = (imgWidth - gap * (num - 1)) / num;
        var yPos = prevBottom - rowGap;

        // Align each row to the image edges.
        var rowLeft = imgLeft;

        /* 全色から最大距離法で代表色を選択 / Select representative colors from full set */
        var selected = selectByMaxDistance(colorList, num);

        /* 選択した色を最近傍法で並べ替え / Sort selected colors by nearest neighbor */
        var sorted = sortByNearest(selected);

        /* グループ化 / Group the row */
        var rowGroup = container.groupItems.add();
        rowGroup.name = (lang === 'ja') ? (num + L('colorsSuffix')) : (num + " " + L('colorsSuffix'));

        for (var n = 0; n < num; n++) {
            var rect2 = rowGroup.pathItems.rectangle(
                yPos,
                rowLeft + n * (rowSize + gap),
                rowSize,
                rowSize
            );
            rect2.stroked = false;

            if (sorted.length > 0) {
                rect2.filled = true;
                rect2.fillColor = sorted[n % sorted.length].swatch.color;
            } else {
                rect2.filled = false;
            }
            if (num === 5 && outOpt.showHEX) addSwatchLabelToGroup(rect2, rowSize, rowGroup, "hex");
        }

        // Duplicate the 5-color row below itself (CMYK adjust) only if enabled
        if (num === 5 && outOpt.out5Adj) {
            try {
                var rowGroup2 = rowGroup.duplicate();
                // Move down by (row height + quarter height)
                rowGroup2.translate(0, - (rowSize + (rowSize / 4)));

                // Adjust colors of duplicated swatches
                for (var p = 0; p < rowGroup2.pathItems.length; p++) {
                    var item = rowGroup2.pathItems[p];
                    if (!item.filled) continue;

                    var c = item.fillColor;
                    var cmykVals = null;

                    // Best-effort: get CMYK values from any color type
                    if (c && c.typename === "CMYKColor") {
                        cmykVals = [c.cyan, c.magenta, c.yellow, c.black];
                    } else {
                        cmykVals = colorToCMYKVals(c);
                    }

                    if (cmykVals && cmykVals.length >= 4) {
                        var cc = Math.min(100, cmykVals[0] * 1.05);
                        var mm = Math.min(100, cmykVals[1] * 1.05);
                        var yy = Math.min(100, cmykVals[2] * 1.05);
                        var kk = Math.min(100, cmykVals[3] * 1.05);

                        // round at ones place (nearest 10)
                        var cNew = new CMYKColor();
                        cNew.cyan = Math.round(cc / 10) * 10;
                        cNew.magenta = Math.round(mm / 10) * 10;
                        cNew.yellow = Math.round(yy / 10) * 10;
                        cNew.black = Math.round(kk / 10) * 10;

                        item.fillColor = cNew;
                    }
                }

                // Rebuild duplicated row labels to match adjusted colors (or remove if disabled)
                // NOTE: Duplicating a group does not guarantee textFrames order matches pathItems.
                // To avoid mismatched labels, we remove all labels and recreate them from pathItems.
                try {
                    // Remove all existing labels first
                    try {
                        if (rowGroup2.textFrames && rowGroup2.textFrames.length) {
                            for (var tfi = rowGroup2.textFrames.length - 1; tfi >= 0; tfi--) {
                                try { rowGroup2.textFrames[tfi].remove(); } catch (eRm) { }
                            }
                        }
                    } catch (eRmAll) { }

                    if (outOpt.showCMYK) {
                        // Create CMYK-only labels for the adjusted row
                        for (var u = 0; u < rowGroup2.pathItems.length; u++) {
                            var pi = rowGroup2.pathItems[u];
                            if (!pi.filled) continue;
                            addSwatchLabelToGroup(pi, rowSize, rowGroup2, "cmyk");
                        }
                    }
                } catch (e2) { }

            } catch (e) { }
        }

        prevBottom = yPos - rowSize;
    }
}

/* 最大距離法でN色を選択（面積で重み付け） / Select N colors by max-distance method (area-weighted) */
// colorList entries: {swatch, r, g, b, area}
// score = minDist × sqrt(area / avgArea) — 面積が大きい色は選ばれやすく、小さい色は選ばれにくい
function selectByMaxDistance(colorList, num) {
    if (colorList.length <= num) return colorList.slice();

    var selected = [];
    var used = [];
    for (var i = 0; i < colorList.length; i++) used.push(false);

    /* 面積の重みを事前計算: pow(area / avgArea, 0.75)
       面積が平均の4倍 → 重み約2.83倍、1/4 → 重み約0.35倍
       sqrt(0.5乗)より強く面積を反映する */
    var totalArea = 0;
    for (var i = 0; i < colorList.length; i++) {
        totalArea += (colorList[i].area || 1);
    }
    var avgArea = totalArea / colorList.length;
    var areaWeights = [];
    for (var i = 0; i < colorList.length; i++) {
        areaWeights.push(Math.pow((colorList[i].area || 1) / avgArea, 0.75));
    }

    /* 最初の色: 最も暗い色 / First color: darkest */
    var firstIdx = 0;
    var minLum = Infinity;
    for (var i = 0; i < colorList.length; i++) {
        var lum = colorList[i].r * 0.299 + colorList[i].g * 0.587 + colorList[i].b * 0.114;
        if (lum < minLum) { minLum = lum; firstIdx = i; }
    }
    selected.push(colorList[firstIdx]);
    used[firstIdx] = true;

    /* 既に選ばれた色群から最も遠い色を順に選択（面積重み付き） / Select farthest color with area weight */
    while (selected.length < num) {
        var bestIdx = -1;
        var bestScore = -1;

        for (var i = 0; i < colorList.length; i++) {
            if (used[i]) continue;

            /* この色と既選択色との最小距離を求める / Find min distance to already selected */
            var minDist = Infinity;
            for (var s = 0; s < selected.length; s++) {
                var dr = colorList[i].r - selected[s].r;
                var dg = colorList[i].g - selected[s].g;
                var db = colorList[i].b - selected[s].b;
                var dist = dr * dr + dg * dg + db * db;
                if (dist < minDist) minDist = dist;
            }

            /* 距離 × 面積重み = スコア（大きいほど優先） */
            var score = minDist * areaWeights[i];
            if (score > bestScore) {
                bestScore = score;
                bestIdx = i;
            }
        }

        selected.push(colorList[bestIdx]);
        used[bestIdx] = true;
    }

    return selected;
}

/* 最近傍法で色を並べ替え / Sort colors by nearest neighbor */
function sortByNearest(colors) {
    if (colors.length <= 1) return colors.slice();

    var remaining = colors.slice();
    var sorted = [];

    /* 最も暗い色からスタート / Start from darkest */
    var startIdx = 0;
    var minLum = Infinity;
    for (var i = 0; i < remaining.length; i++) {
        var lum = remaining[i].r * 0.299 + remaining[i].g * 0.587 + remaining[i].b * 0.114;
        if (lum < minLum) { minLum = lum; startIdx = i; }
    }
    sorted.push(remaining.splice(startIdx, 1)[0]);

    while (remaining.length > 0) {
        var last = sorted[sorted.length - 1];
        var nearestIdx = 0;
        var nearestDist = Infinity;
        for (var i = 0; i < remaining.length; i++) {
            var dr = last.r - remaining[i].r;
            var dg = last.g - remaining[i].g;
            var db = last.b - remaining[i].b;
            var dist = dr * dr + dg * dg + db * db;
            if (dist < nearestDist) { nearestDist = dist; nearestIdx = i; }
        }
        sorted.push(remaining.splice(nearestIdx, 1)[0]);
    }

    return sorted;
}

/* カラーの一致判定用キー / Color identity key for deduplication */
function colorKey(color) {
    var rgb = colorToRGB(color);
    return Math.round(rgb[0]) + "," + Math.round(rgb[1]) + "," + Math.round(rgb[2]);
}

/* 重複除去（面積を合算） / Deduplicate colors (summing areas) */
// Input/Output: [{color, area}, ...]
function deduplicateColors(colors) {
    var seen = {};
    var result = [];
    for (var i = 0; i < colors.length; i++) {
        var key = colorKey(colors[i].color);
        if (seen[key] === undefined) {
            seen[key] = result.length;
            result.push({ color: colors[i].color, area: colors[i].area || 1 });
        } else {
            result[seen[key]].area += (colors[i].area || 1);
        }
    }
    return result;
}

/* カラーをRGB値に変換 / Convert color to RGB values */
function colorToRGB(color) {
    // Fast paths for common types
    if (color.typename === "RGBColor") {
        return [color.red, color.green, color.blue];
    } else if (color.typename === "CMYKColor") {
        var r = 255 * (1 - color.cyan / 100) * (1 - color.black / 100);
        var g = 255 * (1 - color.magenta / 100) * (1 - color.black / 100);
        var b = 255 * (1 - color.yellow / 100) * (1 - color.black / 100);
        return [r, g, b];
    } else if (color.typename === "GrayColor") {
        var v = 255 * (1 - color.gray / 100);
        return [v, v, v];
    }

    // Fallback: try to convert via app.convertSampleColor (e.g., LabColor, pattern/gradient internals, etc.)
    // If conversion is unavailable or fails, fall back to black.
    try {
        if (app.convertSampleColor && typeof ImageColorSpace !== "undefined" && typeof ColorConvertPurpose !== "undefined") {
            var srcSpace = null;
            var src = null;

            if (color.typename === "LabColor") {
                srcSpace = ImageColorSpace.LAB;
                src = [color.l, color.a, color.b];
            } else if (color.typename === "NoColor") {
                return [0, 0, 0];
            } else {
                // Unsupported object type for direct conversion
                srcSpace = null;
            }

            if (srcSpace && src) {
                var dst = app.convertSampleColor(
                    srcSpace,
                    src,
                    ImageColorSpace.RGB,
                    ColorConvertPurpose.defaultpurpose,
                    false,
                    false
                );

                if (dst && dst.length >= 3) {
                    // Clamp to [0..255]
                    var rr = Math.max(0, Math.min(255, dst[0]));
                    var gg = Math.max(0, Math.min(255, dst[1]));
                    var bb = Math.max(0, Math.min(255, dst[2]));
                    return [rr, gg, bb];
                }
            }
        }
    } catch (e) {
        // ignore and fall back
    }

    return [0, 0, 0];
}

/* カラー値から名前を生成 / Generate name from color values */
function colorToName(color) {
    if (color.typename === "RGBColor") {
        return "R=" + Math.round(color.red) + " G=" + Math.round(color.green) + " B=" + Math.round(color.blue);
    } else if (color.typename === "CMYKColor") {
        return "C=" + Math.round(color.cyan) + " M=" + Math.round(color.magenta) + " Y=" + Math.round(color.yellow) + " K=" + Math.round(color.black);
    } else if (color.typename === "GrayColor") {
        return "Gray=" + Math.round(color.gray);
    }
    return L('unknownColorName');
}
