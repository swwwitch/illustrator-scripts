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
- 11色の選出では、ほぼ白／ほぼ黒を除外して中間色を拾いやすく調整
  - CMYKが取得できない場合はRGBフォールバックで判定
- 5色はHEX、5色（CMYK補正）はCMYK表示（任意）
- CMYK→RGB変換は Illustrator の convertSampleColor() を優先し、色距離計算やラベル表示の精度を改善
- CMYK補正は5%刻みで丸め
  - 0/5 は 2.5、5/10 は 7.5 を境界に判定

初回に色が取得できたタイミングでダイアログを表示し、
出力する行（16/11/8/5/5補正）と、カラー情報（HEX/CMYK）を選択できます。
色の選出は段階的減色で固定しています。
処理状況は「準備中 → 画像トレース → カラーパレット生成」のフェーズで表示されます。
ダイアログ表示中はプレビューが更新されます。
OK確定後にスウォッチグループを作成します（キャンセル時は作成しません）。

更新日: 2026-03-06

**********************************************************/


var SCRIPT_VERSION = "v1.6.5";

var __DIALOG_BOUNDS_OUTPUT__ = null; // session-only dialog position memory

/*
ほぼ白の除外しきい値 / Near‑white exclusion thresholds

目的:
11色の代表色選出時に「背景の白」などを除外し、中間色を拾いやすくする。

判定ロジック:
1. まず RGB の見た目の明るさで判定
   - R,G,B がすべて NEAR_WHITE_RGB_MIN 以上なら「白候補」
2. CMYK が取得できる場合
   - C+M+Y+K が NEAR_WHITE_CMYK_TOTAL_MAX 以下なら「ほぼ白」
3. CMYK が取得できない場合
   - RGB 条件のみで「ほぼ白」と判定

※ 淡い色（薄いピンク・水色など）を誤って白扱いしないため、
   RGB 判定を先に行う保守的な条件になっている。
*/
var NEAR_WHITE_CMYK_TOTAL_MAX = 20;   // CMYK合成値（C+M+Y+K）がこの値以下で「ほぼ白」
var NEAR_WHITE_RGB_MIN = 235;         // RGBの各チャンネルがこの値以上で「ほぼ白」

/*
ほぼ黒の除外しきい値 / Near‑black exclusion thresholds

目的:
11色の代表色選出時に「背景の黒」などの極端に暗い色を除外する。

判定ロジック:
1. まず RGB の暗さで判定
   - R,G,B がすべて NEAR_BLACK_RGB_MAX 以下なら「黒候補」
2. CMYK が取得できる場合
   - (C+M+Y+K が NEAR_BLACK_CMYK_TOTAL_MIN 以上) または
   - (K が NEAR_BLACK_K_MIN 以上)
   のどちらかを満たす場合「ほぼ黒」
3. CMYK が取得できない場合
   - RGB 条件のみで「ほぼ黒」と判定

※ 濃い色（濃紺・濃茶など）を黒として誤除外しないよう、
   RGB と CMYK の両方で確認する保守的な条件になっている。
*/
var NEAR_BLACK_CMYK_TOTAL_MIN = 280;  // CMYK合成値（C+M+Y+K）がこの値以上で「ほぼ黒」
var NEAR_BLACK_K_MIN = 85;            // K単独がこの値(%)以上で「ほぼ黒」
var NEAR_BLACK_RGB_MAX = 40;          // RGBの各チャンネルがこの値以下で「ほぼ黒」

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
    progressTracing: {
        ja: "画像トレースを実行中…",
        en: "Running Image Trace…"
    },
    progressPalette: {
        ja: "カラーパレット生成中…",
        en: "Generating Color Palette…"
    },
    progressItem: {
        ja: "処理中…",
        en: "Processing…"
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
    helpTipCMYKDocOnly: {
        ja: "CMYKラベルはCMYKドキュメントでのみ利用できます。",
        en: "CMYK labels are available only in a CMYK document."
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
    },
    fitView: {
        ja: "フィット",
        en: "Fit"
    },
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
function showOutputOptionsDialog(onPreviewChange, onFitView) {
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


    // CMYK label availability: only meaningful/expected in a CMYK document
    var __isDocCMYK = false;
    try {
        __isDocCMYK = (app.activeDocument && app.activeDocument.documentColorSpace === DocumentColorSpace.CMYK);
    } catch (eCS) { __isDocCMYK = false; }

    try {
        if (!__isDocCMYK) {
            cbCMYK.helpTip = L('helpTipCMYKDocOnly');
        }
    } catch (eTip) { }

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
            preview: true,
            showHEX: cbHEX.value,
            showCMYK: cbCMYK.value,
            cascade: true
        };
    }

    var __isInitializing = true;

    function notifyPreviewChange() {
        // Avoid redundant preview redraws during dialog initialization.
        if (__isInitializing) return;
        try {
            if (onPreviewChange) onPreviewChange(getCurrentOptions());
        } catch (eNP) { }
    }

    function updateInfoDims(doNotify) {
        cbHEX.enabled = cb5.value;
        if (!cb5.value) cbHEX.value = false;

        cbCMYK.enabled = (cb5a.value && __isDocCMYK);
        if (!cb5a.value || !__isDocCMYK) cbCMYK.value = false;

        if (doNotify !== false) notifyPreviewChange();
    }

    function updateInfoMode(doNotify) {
        if (rb5Only.value) {
            // Only 5: lock non-5 rows OFF and disable them to avoid confusion.
            cb16.value = false;
            cb11.value = false;
            cb8.value = false;
            cb16.enabled = false;
            cb11.enabled = false;
            cb8.enabled = false;
        } else {
            // All info: enable rows and turn everything ON (default behavior)
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
            updateInfoDims(false);
        }

        if (doNotify !== false) notifyPreviewChange();
    }

    // Wire events
    rbAllInfo.onClick = function () { updateInfoMode(true); };
    rb5Only.onClick = function () { updateInfoMode(true); };

    cb16.onClick = notifyPreviewChange;
    cb11.onClick = notifyPreviewChange;
    cb8.onClick = notifyPreviewChange;

    cb5.onClick = function () {
        updateInfoDims(false);
        notifyPreviewChange();
    };
    cb5a.onClick = function () {
        updateInfoDims(false);
        // If 5色（CMYK補正） is turned ON and CMYK labels are available, auto-enable CMYK labels.
        if (cb5a.value && cbCMYK.enabled && !cbCMYK.value) {
            cbCMYK.value = true;
        }
        notifyPreviewChange();
    };

    cbHEX.onClick = notifyPreviewChange;
    cbCMYK.onClick = notifyPreviewChange;

    // Bottom bar: Fit (left) + spacer + Cancel/OK (right)
    var btnRow = dlg.add('group');
    btnRow.alignment = 'fill';
    btnRow.orientation = 'row';
    btnRow.margins = [0, 10, 0, 0];

    var btnFit = btnRow.add('button', undefined, L('fitView'));
    btnFit.onClick = function () {
        try { if (onFitView) onFitView(); } catch (_) { }
    };

    var spacer = btnRow.add('group');
    spacer.alignment = ['fill', 'center'];
    spacer.preferredSize = [-1, 1];

    var ng = btnRow.add('button', undefined, L('btnCancel'), { name: 'cancel' });
    var ok = btnRow.add('button', undefined, L('btnOK'), { name: 'ok' });

    // Initialize UI state (must be after all controls are created)
    updateInfoDims(false);
    updateInfoMode(false);
    __isInitializing = false;
    notifyPreviewChange();

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

        var progress = createProgress(jobs.length * 2);
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
            // Important: Trace/Expand can leave behind intermediate objects depending on AI version.
            // We try to explicitly clean up the tracing object and the original traced item to avoid residue.
            function traceAndExpand(itemToTrace) {
                var t = null;
                var expanded = null;

                try {
                    t = itemToTrace.trace();
                } catch (eTrace) {
                    // If trace fails, return null (caller should handle)
                    return null;
                }

                if (tracingPresetName && t && t.tracing && t.tracing.tracingOptions) {
                    try {
                        t.tracing.tracingOptions.loadFromPreset(tracingPresetName);
                    } catch (ePreset) {
                        // If preset load fails (e.g., missing/unsupported), continue with current tracing options.
                    }
                }

                try { app.redraw(); } catch (eRd0) { }

                try {
                    // expandTracing() typically creates a GroupItem/CompoundPathItem result.
                    expanded = (t && t.tracing) ? t.tracing.expandTracing() : null;
                } catch (eExpand) {
                    expanded = null;
                }

                // Best-effort: ensure expanded result is on the work layer
                try {
                    if (expanded && expanded.typename && workLayer) {
                        expanded.move(workLayer, ElementPlacement.PLACEATEND);
                    }
                } catch (eMove) { }

                // --- Cleanup residue ---
                // Depending on Illustrator, after expand the tracing object (t) and/or the traced source may remain.
                // We remove them if they still exist and are not the expanded result.
                try {
                    // Remove tracing object if it still exists (and isn't the same as expanded)
                    if (t && t.typename) {
                        try {
                            if (!expanded || t !== expanded) {
                                t.remove();
                            }
                        } catch (eTRm) { }
                    }
                } catch (eTOuter) { }

                try {
                    // Remove the original traced item if it still exists and is different from expanded.
                    // (Some versions keep the original duplicate item around after tracing.)
                    if (itemToTrace && itemToTrace.typename) {
                        try {
                            if (!expanded || itemToTrace !== expanded) {
                                itemToTrace.remove();
                            }
                        } catch (eIRm) { }
                    }
                } catch (eIOuter) { }

                return expanded;
            }

            /* 各複製に対してトレース→拡張→スウォッチ登録 / Trace, expand, and register swatches */
            var outOpt = null; // decided after first successful extraction
            var canceled = false;
            for (var j = 0; j < itemsToProcess.length; j++) {
                try {
                    var job = itemsToProcess[j];
                    progress.set(j * 2 + 1, L('progressTracing'));
                    var p = job.workItem;
                    var colors;
                    if (job.type === "vector") {
                        // ベクター: オブジェクトから直接カラーを抽出（ラスタライズ不要）
                        colors = collectFillColors(p, []);
                        colors = deduplicateColors(colors);
                    } else {
                        var grp = traceAndExpand(p);
                        if (!grp) {
                            // Trace/expand failed; skip this item without stopping the whole batch.
                            continue;
                        }
                        colors = collectFillColors(grp, []);
                        colors = deduplicateColors(colors);
                    }

                    progress.set(j * 2 + 2, L('progressPalette'));
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
                                    // Robust preview refresh:
                                    // - Recreate the preview group each time to avoid lingering items/residue.
                                    // - This also prevents partial-delete failures from leaving behind artifacts.

                                    // Remove existing preview group (best-effort)
                                    try {
                                        if (previewGroup) {
                                            previewGroup.remove();
                                        }
                                    } catch (eRmPrev) { }
                                    previewGroup = null;

                                    // Draw preview only if enabled
                                    if (optNow && optNow.preview) {
                                        try {
                                            previewGroup = job.originalItem.layer.groupItems.add();
                                            previewGroup.name = "__ColorPalettePreview__";
                                        } catch (eNewPg) {
                                            previewGroup = null;
                                        }
                                        if (previewGroup) {
                                            drawSwatchSquares(doc, job.originalItem, colors, optNow, previewGroup);
                                        }
                                    }

                                    try { app.redraw(); } catch (eRd2) { }
                                } catch (eCb) { }
                            }, function () {
                                // Fit view to originalItem + previewGroup
                                var items = [];
                                try { items.push(job.originalItem); } catch (_) { }
                                try { if (previewGroup) items.push(previewGroup); } catch (_) { }
                                if (items.length === 0) return;
                                var l = Infinity, t = -Infinity, r = -Infinity, b = Infinity;
                                for (var fi = 0; fi < items.length; fi++) {
                                    var gb = items[fi].geometricBounds;
                                    if (gb[0] < l) l = gb[0];
                                    if (gb[1] > t) t = gb[1];
                                    if (gb[2] > r) r = gb[2];
                                    if (gb[3] < b) b = gb[3];
                                }
                                var v = doc.activeView;
                                var margin = 1.05;
                                v.centerPoint = [(l + r) / 2, (t + b) / 2];
                                var vb = v.bounds;
                                var vw = vb[2] - vb[0];
                                var vh = vb[1] - vb[3];
                                var ow = r - l;
                                var oh = t - b;
                                if (ow > 0 && oh > 0) {
                                    var scale = Math.min(vw / ow, vh / oh) / margin;
                                    v.zoom = v.zoom * scale;
                                }
                                try { app.redraw(); } catch (_) { }
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
                                // Register swatches ONLY for 5 and 5 (CMYK Adjust), each as its own group.
                                createSwatchGroupsFor5Only(doc, groupName, colors, outOpt);
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
                            // Register swatches ONLY for 5 and 5 (CMYK Adjust), each as its own group.
                            createSwatchGroupsFor5Only(doc, groupName, colors, outOpt);
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

                    progress.set(j * 2 + 2, L('progressPalette'));
                } catch (err) {
                    // ignore per-item failures to keep processing others
                }
                if (canceled) break;
            }
            progress.set(jobs.length * 2, L('progressDone'));

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

// -----------------------------------------------------------------------------
// LEGACY / UNUSED
// The following helper is currently not used by this script.
// Kept intentionally for potential reuse / backwards compatibility.
// -----------------------------------------------------------------------------
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
// --- end of LEGACY / UNUSED ---

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

// Helper: build representative 5-color list from full colorList (max-distance + nearest sort)
// Returns an array of entries: [{swatch:{color,area}, r,g,b, area}, ...]
function buildRepresentativeList(colorList, n) {
    try {
        var sel = selectByMaxDistance(colorList, n);
        return sortByNearest(sel);
    } catch (e) {
        return [];
    }
}

// Helper: CMYK adjust used for the duplicated 5-color row (same policy as drawSwatchSquares)
function buildCmykAdjustedColor(color) {
    var cmykVals = null;
    try {
        if (color && color.typename === "CMYKColor") {
            cmykVals = [color.cyan, color.magenta, color.yellow, color.black];
        } else {
            cmykVals = colorToCMYKVals(color);
        }
    } catch (e) { cmykVals = null; }

    if (!cmykVals || cmykVals.length < 4) return null;

    function roundToNearest5(v) {
        // Midpoint rule: 0–2.49→0, 2.5–7.49→5, 7.5–12.49→10 ...
        return Math.floor((v + 2.5) / 5) * 5;
    }
    function clampPercent(v) {
        return Math.max(0, Math.min(100, v));
    }

    var cNew = new CMYKColor();
    cNew.cyan = clampPercent(roundToNearest5(cmykVals[0]));
    cNew.magenta = clampPercent(roundToNearest5(cmykVals[1]));
    cNew.yellow = clampPercent(roundToNearest5(cmykVals[2]));
    cNew.black = clampPercent(roundToNearest5(cmykVals[3]));

    return cNew;
}

// Helper: create the two swatch groups (5 and 5Adj) from the full extracted colors.
function createSwatchGroupsFor5Only(doc, baseName, fullColors, outOpt) {
    if (!outOpt) return;

    // Build the same working list used for palette selection: [{swatch,color,r,g,b,area}, ...]
    var remaining = [];
    for (var m = 0; m < fullColors.length; m++) {
        var cObj = fullColors[m].color ? fullColors[m].color : fullColors[m];
        var aObj = (fullColors[m] && fullColors[m].area !== undefined) ? fullColors[m].area : 1;
        var rgb = colorToRGB(cObj);
        remaining.push({ swatch: { color: cObj, area: aObj }, r: rgb[0], g: rgb[1], b: rgb[2], area: aObj });
    }

    if (!remaining.length) return;

    // Representative 5 colors
    var rep5 = buildRepresentativeList(remaining, 5);

    // Group naming
    var name5 = baseName + " - " + L('opt5');
    var name5a = baseName + " - " + L('opt5Adj');

    if (outOpt.out5) {
        // Pass as [{color,area}, ...] to reuse existing swatch creation logic.
        var colors5 = [];
        for (var i = 0; i < rep5.length; i++) {
            colors5.push({ color: rep5[i].swatch.color, area: rep5[i].swatch.area || 1 });
        }
        createSwatchGroupFromColors(doc, name5, colors5);
    }

    if (outOpt.out5Adj) {
        var colors5a = [];
        for (var j = 0; j < rep5.length; j++) {
            var baseC = rep5[j].swatch.color;
            var adj = buildCmykAdjustedColor(baseC);
            colors5a.push({ color: adj ? adj : baseC, area: rep5[j].swatch.area || 1 });
        }
        createSwatchGroupFromColors(doc, name5a, colors5a);
    }
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

    outOpt = outOpt || { out16: true, out11: true, out8: true, out5: true, out5Adj: true, showHEX: true, showCMYK: true, cascade: true };

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
    function round5(v) {
        // Midpoint rule: 0–2.49→0, 2.5–7.49→5, 7.5–12.49→10 ...
        return Math.floor((v + 2.5) / 5) * 5;
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
                var Cc = round5(cmyk[0]);
                var Mm = round5(cmyk[1]);
                var Yy = round5(cmyk[2]);
                var Kk = round5(cmyk[3]);
                return { text: L('cmykPrefix') + Cc + ", " + Mm + ", " + Yy + ", " + Kk, rgb: rgb };
            }
            // If CMYK is unavailable, fall back to HEX only
            return { text: hex, rgb: rgb };
        }

        // mode === "both"
        if (cmyk) {
            var C = round5(cmyk[0]);
            var M = round5(cmyk[1]);
            var Y = round5(cmyk[2]);
            var K = round5(cmyk[3]);
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
    /* cascade モード用: 前段の選出結果を保持 / Keep previous selection for cascade mode */
    var cascadePrev = null;
    if (outOpt.out16 || outOpt.cascade) {
        var sel16 = selectByMaxDistance(colorList, 16);
        cascadePrev = sel16;
    }
    if (outOpt.out16) {
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
        var isRowEnabled = !((num === 11 && !outOpt.out11) || (num === 8 && !outOpt.out8) || (num === 5 && !outOpt.out5));

        /* 段階的減色: 出力しない行でも内部計算を行い、次段へ渡す */
        /* Cascade: compute selection even for skipped rows to feed the next stage */
        var sourceColors;
        if (outOpt.cascade && cascadePrev) {
            sourceColors = cascadePrev;
        } else {
            sourceColors = colorList;
            /* 11色のみ: ほぼ白・ほぼ黒を除外 / For 11 colors: exclude near-white/black */
            if (num === 11) {
                sourceColors = [];
                for (var f = 0; f < colorList.length; f++) {
                    if (!isNearlyWhite(colorList[f]) && !isNearlyBlack(colorList[f])) sourceColors.push(colorList[f]);
                }
                if (sourceColors.length === 0) sourceColors = colorList; // fallback
            }
        }

        var selected = selectByMaxDistance(sourceColors, num);
        if (outOpt.cascade) cascadePrev = selected;

        if (!isRowEnabled) continue;

        // Make each row a true square grid: size is computed to fit within imgWidth with the same horizontal gap.
        var rowSize = (imgWidth - gap * (num - 1)) / num;
        var yPos = prevBottom - rowGap;

        // Align each row to the image edges.
        var rowLeft = imgLeft;

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
                        var cNew = buildCmykAdjustedColor(c);
                        if (cNew) item.fillColor = cNew;
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

/* エントリからCMYK値を取得 / Get CMYK values from a color entry
   Return: [C, M, Y, K] or null */
function entryToCMYK(entry) {
    var cmyk = null;
    try {
        var c = entry && entry.swatch ? entry.swatch.color : null;
        if (c && c.typename === "CMYKColor") {
            cmyk = [c.cyan, c.magenta, c.yellow, c.black];
        } else if (c && c.typename === "RGBColor" && app.convertSampleColor) {
            var dst = app.convertSampleColor(
                ImageColorSpace.RGB,
                [c.red, c.green, c.blue],
                ImageColorSpace.CMYK,
                ColorConvertPurpose.defaultpurpose,
                false,
                false
            );
            if (dst && dst.length >= 4) {
                cmyk = [dst[0], dst[1], dst[2], dst[3]];
            }
        }
    } catch (e) { }
    return cmyk;
}

/* 最大距離法でN色を選択（面積で重み付け） / Select N colors by max-distance method (area-weighted) */
// colorList entries: {swatch, r, g, b, area}
// score = minDist × sqrt(area / avgArea) — 面積が大きい色は選ばれやすく、小さい色は選ばれにくい
/* ほぼ白の判定 / Determine if a color entry is nearly white
   RGB の明るさで白候補を絞り、可能なら CMYK 合計量で確認する。
   CMYK が取得できない場合は RGB 条件のみで判定する。 */
function isNearlyWhite(entry) {
    // RGB条件
    if (entry.r < NEAR_WHITE_RGB_MIN || entry.g < NEAR_WHITE_RGB_MIN || entry.b < NEAR_WHITE_RGB_MIN) {
        return false;
    }
    // CMYK条件
    var cmyk = entryToCMYK(entry);
    if (cmyk && cmyk.length >= 4) {
        var total = cmyk[0] + cmyk[1] + cmyk[2] + cmyk[3];
        return total <= NEAR_WHITE_CMYK_TOTAL_MAX;
    }
    // CMYKが取得できない場合はRGB条件のみで「ほぼ白」と判定
    return true;
}

/* ほぼ黒の判定 / Determine if a color entry is nearly black
   RGB の暗さで黒候補を絞り、可能なら CMYK 合計量または K 値で確認する。
   CMYK が取得できない場合は RGB 条件のみで判定する。 */
function isNearlyBlack(entry) {
    // RGB条件
    if (entry.r > NEAR_BLACK_RGB_MAX || entry.g > NEAR_BLACK_RGB_MAX || entry.b > NEAR_BLACK_RGB_MAX) {
        return false;
    }
    // CMYK条件
    var cmyk = entryToCMYK(entry);
    if (cmyk && cmyk.length >= 4) {
        var total = cmyk[0] + cmyk[1] + cmyk[2] + cmyk[3];
        if (total >= NEAR_BLACK_CMYK_TOTAL_MIN || cmyk[3] >= NEAR_BLACK_K_MIN) {
            return true;
        }
        return false;
    }
    // CMYKが取得できない場合はRGB条件のみで「ほぼ黒」と判定
    return true;
}

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

/* カラーをCMYK値に変換 / Convert color to CMYK values */
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

/* カラーをRGB値に変換 / Convert color to RGB values */
function colorToRGB(color) {
    if (!color) return [0, 0, 0];

    function clamp255(v) {
        return Math.max(0, Math.min(255, v));
    }

    function convertViaSample(srcSpace, srcValues) {
        try {
            if (app.convertSampleColor && typeof ImageColorSpace !== "undefined" && typeof ColorConvertPurpose !== "undefined") {
                var dst = app.convertSampleColor(
                    srcSpace,
                    srcValues,
                    ImageColorSpace.RGB,
                    ColorConvertPurpose.defaultpurpose,
                    false,
                    false
                );
                if (dst && dst.length >= 3) {
                    return [clamp255(dst[0]), clamp255(dst[1]), clamp255(dst[2])];
                }
            }
        } catch (e) { }
        return null;
    }

    // Fast path: RGB is already the target space.
    if (color.typename === "RGBColor") {
        return [clamp255(color.red), clamp255(color.green), clamp255(color.blue)];
    }

    // Prefer Illustrator's own conversion for CMYK to better match document appearance.
    if (color.typename === "CMYKColor") {
        var rgbFromCmyk = convertViaSample(
            ImageColorSpace.CMYK,
            [color.cyan, color.magenta, color.yellow, color.black]
        );
        if (rgbFromCmyk) return rgbFromCmyk;

        // Fallback: simple approximation if sample conversion is unavailable.
        var r = 255 * (1 - color.cyan / 100) * (1 - color.black / 100);
        var g = 255 * (1 - color.magenta / 100) * (1 - color.black / 100);
        var b = 255 * (1 - color.yellow / 100) * (1 - color.black / 100);
        return [clamp255(r), clamp255(g), clamp255(b)];
    }

    if (color.typename === "GrayColor") {
        var rgbFromGray = convertViaSample(ImageColorSpace.GrayScale, [color.gray]);
        if (rgbFromGray) return rgbFromGray;

        var v = 255 * (1 - color.gray / 100);
        return [clamp255(v), clamp255(v), clamp255(v)];
    }

    if (color.typename === "LabColor") {
        var rgbFromLab = convertViaSample(ImageColorSpace.LAB, [color.l, color.a, color.b]);
        if (rgbFromLab) return rgbFromLab;
    }

    if (color.typename === "NoColor") {
        return [0, 0, 0];
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
