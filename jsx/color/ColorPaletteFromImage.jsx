#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/**********************************************************

ColorPaletteFromImage.jsx 

DESCRIPTION

選択している配置画像を複製・トレース・拡張し、
抽出した色をスウォッチグループに登録。
元画像の下に16色・11色・8色・5色の
カラーパレットを正方形で表示するスクリプト。
色は最近傍法でグラデーション風にソート。
11色・8色・5色は最大距離法で代表色を選択。

**********************************************************/

var SCRIPT_VERSION = "v1.2";

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
    progressWorking: {
        ja: "処理中…",
        en: "Working…"
    },
    progressItem: {
        ja: "{0}/{1} を処理中",
        en: "Processing {0}/{1}"
    },
    progressDone: {
        ja: "完了",
        en: "Done"
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
                    } catch (e) {}
                },
                close: function () {
                    try { w.close(); } catch (e) {}
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

            // --- Helper: Trace and expand, for vector/raster ---
            function traceAndExpand(itemToTrace, isVector) {
                var t;
                if (isVector) {
                    var rasterOpts = new RasterizeOptions();
                    rasterOpts.resolution = 72;
                    rasterOpts.antiAliasingMethod = AntiAliasingMethod.ARTOPTIMIZED;
                    var rasterItem = doc.rasterize(itemToTrace, itemToTrace.geometricBounds, rasterOpts);
                    rasterItem.move(workLayer, ElementPlacement.PLACEATEND);
                    t = rasterItem.trace();
                } else {
                    t = itemToTrace.trace();
                }
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
                } catch (e) {}
                return expanded;
            }

            /* 各複製に対してトレース→拡張→スウォッチ登録 / Trace, expand, and register swatches */
            for (var j = 0; j < itemsToProcess.length; j++) {
                try {
                    var job = itemsToProcess[j];
                    progress.set(j, LF('progressItem', j + 1, itemsToProcess.length));
                    var p = job.workItem;
                    var grp = traceAndExpand(p, job.type === "vector");

                    /* スウォッチグループを作成 / Create swatch group */
                    var groupName;
                    if (job.originalItem.typename === "PlacedItem" && job.originalItem.file) {
                        groupName = job.originalItem.file.name;
                    } else {
                        groupName = L('itemPrefix') + j;
                    }
                    var swatchGroup = doc.swatchGroups.add();
                    swatchGroup.name = groupName;

                    /* 拡張結果からスウォッチグループに登録 / Register colors to swatch group */
                    addColorsToSwatchGroup(doc, swatchGroup, grp);

                    /* 元画像の下に正方形を作成してスウォッチの色を適用 / Draw color palette squares below image */
                    drawSwatchSquares(doc, job.originalItem, swatchGroup);
                    progress.set(j + 1, LF('progressItem', j + 1, itemsToProcess.length));
                } catch (err) {
                    // ignore per-item failures to keep processing others
                }
            }
            progress.set(itemsToProcess.length, L('progressDone'));

        } finally {
            /* 作業用レイヤーを削除 / Remove work layer */
            try {
                if (workLayer) workLayer.remove();
            } catch (e) { }
            try { if (progress) progress.close(); } catch (e) {}
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

/* 元画像の下にカラーパレットを描画 / Draw color palette squares below original image */
function drawSwatchSquares(doc, originalItem, swatchGroup) {
    var swatches = swatchGroup.getAllSwatches();
    var numSquares = 16;

    /* 最近傍法でソート（最も暗い色から、色距離が近い順に並べる） / Sort by nearest neighbor from darkest */
    var remaining = [];
    for (var m = 0; m < swatches.length; m++) {
        var rgb = colorToRGB(swatches[m].color);
        remaining.push({ swatch: swatches[m], r: rgb[0], g: rgb[1], b: rgb[2] });
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

    var gapRatio = 0.10;          // gap = 10% of cell width
    var gap = cell16 * gapRatio;
    var squareSize = cell16 - gap; // square = remaining 90%

    // Vertical gap should match the column gap of a 12-column layout.
    // (Same gap rule as this script: gap = (imgWidth / N) / 10)
    var gap12 = (imgWidth / 12) / 10;
    var rowGap = gap12;

    // Gap between image bottom and the first row is half of the swatch height.
    var startY = imgBottom - (squareSize / 2);

    /* 元画像のレイヤーに作成する / Create on original image's layer */
    var targetLayer = originalItem.layer;

    /* 16色の正方形を作成・グループ化 / Create and group 16-color squares */
    var group16 = targetLayer.groupItems.add();
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

    /* 11色・8色・5色バージョン（最大距離法で代表色を選択） / 11, 8, 5 color rows via max-distance selection */
    var rows = [11, 8, 5];
    var prevBottom = startY - squareSize;

    for (var r = 0; r < rows.length; r++) {
        var num = rows[r];
        // Make each row a true square grid: size is computed to fit within imgWidth with the same horizontal gap.
        var rowSize = (imgWidth - gap * (num - 1)) / num;
        var yPos = prevBottom - rowGap;

        // Right-align each row to the right edge of the 16-color row.
        var rightEdge16 = imgLeft + (numSquares - 1) * (squareSize + gap) + squareSize;
        var rowTotalWidth = num * rowSize + (num - 1) * gap;
        var rowLeft = rightEdge16 - rowTotalWidth;

        /* 最大距離法でN色を選択 / Select N colors by max-distance */
        var selected = selectByMaxDistance(colorList, num);

        /* 選択した色を最近傍法で並べ替え / Sort selected colors by nearest neighbor */
        var sorted = sortByNearest(selected);

        /* グループ化 / Group the row */
        var rowGroup = targetLayer.groupItems.add();
        rowGroup.name = num + " " + L('colorsSuffix');

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
        }

        prevBottom = yPos - rowSize;
    }
}

/* 最大距離法でN色を選択 / Select N colors by max-distance method */
function selectByMaxDistance(colorList, num) {
    if (colorList.length <= num) return colorList.slice();

    var selected = [];
    var used = [];
    for (var i = 0; i < colorList.length; i++) used.push(false);

    /* 最初の色: 最も暗い色 / First color: darkest */
    var firstIdx = 0;
    var minLum = Infinity;
    for (var i = 0; i < colorList.length; i++) {
        var lum = colorList[i].r * 0.299 + colorList[i].g * 0.587 + colorList[i].b * 0.114;
        if (lum < minLum) { minLum = lum; firstIdx = i; }
    }
    selected.push(colorList[firstIdx]);
    used[firstIdx] = true;

    /* 既に選ばれた色群から最も遠い色を順に選択 / Select farthest color from already selected */
    while (selected.length < num) {
        var bestIdx = -1;
        var bestDist = -1;

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

            /* 最小距離が最大のものを選択 / Pick the one with largest min distance */
            if (minDist > bestDist) {
                bestDist = minDist;
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
