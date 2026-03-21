#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/**********************************************************

ColorPaletteFromImage.jsx

DESCRIPTION

選択したオブジェクトの塗り色からカラーパレットを作成します。

- アートボード上でオブジェクトを選択して実行
- 塗り色を抽出し、重複を除去（同色の面積を合算）
- 全色 → 16色 → 11色 → 8色 → 5色 と段階的に絞り込み
- 各行を選択オブジェクトの下に正方形で出力
- 色の絞り込みは最大距離法（面積重み付き）、並びは最近傍法でソート
  - 面積が大きい色ほど代表色として選ばれやすくなる
- 5色はHEX、5色（CMYK補正）はCMYK表示（任意）
  - CMYK補正は各C/M/Y/Kを+5%し、10刻みで丸め

初回に色が取得できたタイミングでダイアログを表示し、
出力する行（16/11/8/5/5補正）と、カラー情報（HEX/CMYK）を選択できます。
ダイアログ表示中はプレビューが更新されます。
OK確定後にスウォッチグループを作成します（キャンセル時は作成しません）。

更新日: 2026-03-05

**********************************************************/


var SCRIPT_VERSION = "v1.3";

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
    noColors: {
        ja: "塗り色が見つかりませんでした。",
        en: "No fill colors found."
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
            cb16.value = false;
            cb11.value = false;
            cb8.value = false;
        } else {
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

    if (!sel || sel.length === 0) {
        alert(L('noSelection'));
    } else {
        /* 選択オブジェクトから塗り色を収集 / Collect fill colors from selection */
        var rawColors = [];
        for (var s = 0; s < sel.length; s++) {
            collectFillColors(sel[s], rawColors);
        }

        /* 重複除去 / Deduplicate */
        var uniqueColors = deduplicateColors(rawColors);

        if (uniqueColors.length === 0) {
            alert(L('noColors'));
        } else {
            /* 選択オブジェクトのバウンディングボックス / Selection bounding box */
            var bounds = sel[0].geometricBounds;
            var bLeft = bounds[0], bTop = bounds[1], bRight = bounds[2], bBottom = bounds[3];
            for (var s = 1; s < sel.length; s++) {
                var b = sel[s].geometricBounds;
                if (b[0] < bLeft) bLeft = b[0];
                if (b[1] > bTop) bTop = b[1];
                if (b[2] > bRight) bRight = b[2];
                if (b[3] < bBottom) bBottom = b[3];
            }
            /* パレット配置の基準となる仮想アイテム / Virtual anchor for palette placement */
            var anchorItem = {
                left: bLeft,
                top: bTop,
                width: bRight - bLeft,
                height: bTop - bBottom,
                layer: sel[0].layer
            };

            /* プレビュー＆ダイアログ / Preview and dialog */
            var outAll = { out16: true, out11: true, out8: true, out5: true, out5Adj: true, showHEX: true, showCMYK: true };
            var previewGroup = null;
            try {
                previewGroup = anchorItem.layer.groupItems.add();
                previewGroup.name = "__ColorPalettePreview__";
            } catch (ePg) { }

            try {
                if (previewGroup) {
                    drawSwatchSquares(doc, anchorItem, uniqueColors, outAll, previewGroup);
                }
            } catch (ePrev) { }

            try { app.redraw(); } catch (eRd) { }
            try { $.sleep(80); } catch (eSl) { }

            var outOpt = showOutputOptionsDialog(function (optNow) {
                try {
                    if (!previewGroup) return;
                    while (previewGroup.pageItems.length) {
                        try { previewGroup.pageItems[0].remove(); } catch (eRm) { break; }
                    }
                    if (optNow && optNow.preview) {
                        drawSwatchSquares(doc, anchorItem, uniqueColors, optNow, previewGroup);
                    }
                    try { app.redraw(); } catch (eRd2) { }
                } catch (eCb) { }
            });

            if (!outOpt) {
                /* キャンセル / Canceled */
                try { if (previewGroup) previewGroup.remove(); } catch (eRmPrev) { }
            } else {
                /* スウォッチグループ作成 / Create swatch group */
                var swatchGroup = createSwatchGroupFromColors(doc, L('itemPrefix') + "0", uniqueColors);

                /* 最終出力 / Final output */
                var finalGroup = null;
                try {
                    finalGroup = anchorItem.layer.groupItems.add();
                    finalGroup.name = "__ColorPalette__";
                } catch (eFg) { }
                try {
                    if (finalGroup) {
                        drawSwatchSquares(doc, anchorItem, uniqueColors, outOpt, finalGroup);
                    } else {
                        drawSwatchSquares(doc, anchorItem, uniqueColors, outOpt);
                    }
                } catch (eFinal) { }

                try { if (previewGroup) previewGroup.remove(); } catch (eRmPrev2) { }
            }

            app.redraw();
        }
    }
}

/* 塗り色を再帰的に収集（面積付き） / Collect fill colors recursively (with area) */
function collectFillColors(item, out) {
    if (!out) out = [];
    if (!item) return out;
    if (item.typename === "PathItem") {
        if (item.filled && item.fillColor) {
            var c = item.fillColor;
            if (c.typename !== "NoColor" && c.typename !== "PatternColor" &&
                c.typename !== "GradientColor" && c.typename !== "SpotColor") {
                var a = 1;
                try { a = Math.abs(item.area); } catch (eA) { }
                if (a < 1) a = 1;
                out.push({ color: c, area: a });
            }
        }
        return out;
    }
    if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            collectFillColors(item.pageItems[i], out);
        }
    } else if (item.typename === "CompoundPathItem") {
        for (var i = 0; i < item.pathItems.length; i++) {
            collectFillColors(item.pathItems[i], out);
        }
    }
    return out;
}

/* カラーの一致判定用キー / Color identity key for deduplication */
function colorKey(color) {
    var rgb = colorToRGB(color);
    return Math.round(rgb[0]) + "," + Math.round(rgb[1]) + "," + Math.round(rgb[2]);
}

/* 重複除去（面積を合算） / Deduplicate colors (summing areas) */
function deduplicateColors(colors) {
    var seen = {};
    var result = [];
    for (var i = 0; i < colors.length; i++) {
        var key = colorKey(colors[i].color);
        if (seen[key] === undefined) {
            seen[key] = result.length;
            result.push({ color: colors[i].color, area: colors[i].area });
        } else {
            result[seen[key]].area += colors[i].area;
        }
    }
    return result;
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

/* スウォッチをグループに追加 / Add swatch to group */
function addSwatchToGroup(doc, swatchGroup, color) {
    try {
        if (color.typename === "SpotColor") return;
        if (color.typename === "PatternColor") return;
        if (color.typename === "GradientColor") return;
        if (color.typename === "NoColor") return;

        var swatchName = colorToName(color);

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

// Helper: Create swatch group from colors array (accepts Color or {color, area})
function createSwatchGroupFromColors(doc, groupName, colors) {
    var swatchGroup = doc.swatchGroups.add();
    swatchGroup.name = groupName;
    for (var i = 0; i < colors.length; i++) {
        var c = colors[i].color || colors[i];
        addSwatchToGroup(doc, swatchGroup, c);
    }
    return swatchGroup;
}

/* 元画像の下にカラーパレットを描画 / Draw color palette squares below original item */
// swatchSource: SwatchGroup or Array of Color
function drawSwatchSquares(doc, originalItem, swatchSource, outOpt, containerGroup) {
    var swatches;
    // Accept SwatchGroup or Array of Color
    if (swatchSource && typeof swatchSource.getAllSwatches === "function") {
        swatches = swatchSource.getAllSwatches();
    } else if (swatchSource && swatchSource.length !== undefined) {
        // treat as array of Color or {color, area}
        swatches = [];
        for (var i = 0; i < swatchSource.length; i++) {
            var src = swatchSource[i];
            swatches.push({ color: src.color || src, area: src.area || 1 });
        }
    } else {
        swatches = [];
    }
    var numSquares = 16;

    outOpt = outOpt || { out16: true, out11: true, out8: true, out5: true, out5Adj: true, showHEX: true, showCMYK: true };

    /* 最近傍法でソート / Sort by nearest neighbor from darkest */
    var remaining = [];
    for (var m = 0; m < swatches.length; m++) {
        var rgb = colorToRGB(swatches[m].color);
        remaining.push({ swatch: swatches[m], r: rgb[0], g: rgb[1], b: rgb[2], area: swatches[m].area || 1 });
    }

    var colorList = [];
    if (remaining.length > 0) {
        var startIdx = 0;
        var minLum = Infinity;
        for (var q = 0; q < remaining.length; q++) {
            var lum = remaining[q].r * 0.299 + remaining[q].g * 0.587 + remaining[q].b * 0.114;
            if (lum < minLum) { minLum = lum; startIdx = q; }
        }
        colorList.push(remaining.splice(startIdx, 1)[0]);

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

    /* 元アイテムの位置・サイズを取得 / Get original item position and size */
    var imgLeft = originalItem.left;
    var imgTop = originalItem.top;
    var imgWidth = originalItem.width;
    var imgBottom = imgTop - originalItem.height;

    /* 正方形のサイズを計算 / Calculate square size */
    var cols16 = numSquares;
    var cell16 = imgWidth / cols16;

    var gapRatio = 0.10;
    var gap = cell16 * gapRatio;
    var squareSize = (imgWidth - (cols16 - 1) * gap) / cols16;

    var gap12 = (imgWidth / 12) / 10;
    var rowGap = gap12;

    var startY = imgBottom - squareSize;

    /* 元アイテムのレイヤーに作成する / Create on original item's layer */
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

            var fs = size / 10;

            var x = rect.left;
            var y = rect.top - rect.height - (fs / 2);

            var tf = targetLayer.textFrames.add();
            tf.contents = info.text;
            tf.textRange.justification = Justification.LEFT;
            tf.textRange.characterAttributes.size = fs;
            tf.textRange.fillColor = getLabelBlackColor();

            try {
                tf.textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Regular");
            } catch (eFont) {
                try {
                    tf.textRange.characterAttributes.textFont = app.textFonts.getByName("Myriad Pro");
                } catch (eFont2) { }
            }

            tf.position = [x, y];

            try { tf.move(group, ElementPlacement.PLACEATEND); } catch (e) { }
        } catch (e) { }
    }

    /* 16色の正方形を作成・グループ化 / Create and group 16-color squares */
    if (outOpt.out16) {
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

            if (colorList.length > 0) {
                rect.filled = true;
                rect.fillColor = colorList[n % colorList.length].swatch.color;
            } else {
                rect.filled = false;
            }
        }
    }

    /* 11色・8色・5色バージョン（最大距離法で代表色を選択） / 11, 8, 5 color rows via max-distance selection */
    var rows = [11, 8, 5];
    var prevBottom;
    if (outOpt.out16) {
        prevBottom = startY - squareSize;
    } else {
        prevBottom = startY + rowGap;
    }

    for (var r = 0; r < rows.length; r++) {
        var num = rows[r];
        if ((num === 11 && !outOpt.out11) || (num === 8 && !outOpt.out8) || (num === 5 && !outOpt.out5)) {
            continue;
        }
        var rowSize = (imgWidth - gap * (num - 1)) / num;
        var yPos = prevBottom - rowGap;

        var rowLeft = imgLeft;

        var selected = selectByMaxDistance(colorList, num);
        var sorted = sortByNearest(selected);

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
                rowGroup2.translate(0, - (rowSize + (rowSize / 4)));

                for (var p = 0; p < rowGroup2.pathItems.length; p++) {
                    var item = rowGroup2.pathItems[p];
                    if (!item.filled) continue;

                    var c = item.fillColor;
                    if (c && c.typename === "CMYKColor") {
                        var cc = Math.min(100, c.cyan * 1.05);
                        var mm = Math.min(100, c.magenta * 1.05);
                        var yy = Math.min(100, c.yellow * 1.05);
                        var kk = Math.min(100, c.black * 1.05);

                        c.cyan = Math.round(cc / 10) * 10;
                        c.magenta = Math.round(mm / 10) * 10;
                        c.yellow = Math.round(yy / 10) * 10;
                        c.black = Math.round(kk / 10) * 10;

                        item.fillColor = c;
                    }
                }

                try {
                    if (!outOpt.showCMYK) {
                        if (rowGroup2.textFrames && rowGroup2.textFrames.length) {
                            for (var tfi = rowGroup2.textFrames.length - 1; tfi >= 0; tfi--) {
                                try { rowGroup2.textFrames[tfi].remove(); } catch (eRm) { }
                            }
                        }
                    } else {
                        if (!rowGroup2.textFrames || rowGroup2.textFrames.length === 0) {
                            for (var u0 = 0; u0 < rowGroup2.pathItems.length; u0++) {
                                var pi0 = rowGroup2.pathItems[u0];
                                if (!pi0.filled) continue;
                                addSwatchLabelToGroup(pi0, rowSize, rowGroup2, "cmyk");
                            }
                        } else {
                            var nMin = Math.min(rowGroup2.textFrames.length, rowGroup2.pathItems.length);
                            for (var u = 0; u < nMin; u++) {
                                var pi = rowGroup2.pathItems[u];
                                if (!pi.filled) continue;
                                var info2 = buildColorLabel(pi.fillColor, "cmyk");
                                var tf2 = rowGroup2.textFrames[u];
                                tf2.contents = info2.text;
                                tf2.textRange.fillColor = getLabelBlackColor();
                                try {
                                    tf2.textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-Regular");
                                } catch (eFont3) {
                                    try {
                                        tf2.textRange.characterAttributes.textFont = app.textFonts.getByName("Myriad Pro");
                                    } catch (eFont4) { }
                                }
                            }
                        }
                    }
                } catch (e2) { }

            } catch (e) { }
        }

        prevBottom = yPos - rowSize;
    }
}

/* 最大距離法でN色を選択（面積重み付き） / Select N colors by max-distance method (area-weighted) */
function selectByMaxDistance(colorList, num) {
    if (colorList.length <= num) return colorList.slice();

    var selected = [];
    var used = [];
    for (var i = 0; i < colorList.length; i++) used.push(false);

    /* 面積の重みを事前計算: sqrt(area / avgArea)
       面積が平均の4倍 → 重み2倍、1/4 → 重み0.5倍 */
    var totalArea = 0;
    for (var i = 0; i < colorList.length; i++) {
        totalArea += (colorList[i].area || 1);
    }
    var avgArea = totalArea / colorList.length;
    var areaWeights = [];
    for (var i = 0; i < colorList.length; i++) {
        areaWeights.push(Math.sqrt((colorList[i].area || 1) / avgArea));
    }

    var firstIdx = 0;
    var minLum = Infinity;
    for (var i = 0; i < colorList.length; i++) {
        var lum = colorList[i].r * 0.299 + colorList[i].g * 0.587 + colorList[i].b * 0.114;
        if (lum < minLum) { minLum = lum; firstIdx = i; }
    }
    selected.push(colorList[firstIdx]);
    used[firstIdx] = true;

    while (selected.length < num) {
        var bestIdx = -1;
        var bestScore = -1;

        for (var i = 0; i < colorList.length; i++) {
            if (used[i]) continue;

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

/* カラーをRGB値に変換 / Convert color to RGB values */
function colorToRGB(color) {
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