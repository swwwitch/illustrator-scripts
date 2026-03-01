#target illustrator

// --- Geometry tolerance (for grouping "same shape") ---
// Quantize values to reduce tiny float jitter after Pathfinder/Expand.
var GEOM_TOL_PT = 0.1;        // position/size tolerance in points
var AREA_TOL = 0.01;          // area tolerance in document units (PathItem.area is in square points)

// --- Blending mode ---
// If true: convert to RGB and blend in linear light (can better match on-screen opacity compositing,
// but may shift colors due to color profile conversions).
// If false: blend directly in the source color space (previous behavior; preserves original hues better).
var USE_GAMMA_CORRECT_BLEND = false;
// If true: when flattening a single item (opacity over white), composite in RGB (linear light) to better match on-screen appearance.
// This avoids the common "too dark" result from naive CMYK ink scaling.
var USE_RGB_WHITE_COMPOSITE = false;

// Convert a value to an integer "tick" based on step.
// Using ticks avoids float-string instability in geometry keys.
function toTick(v, step) {
    if (!step || step <= 0) return v;
    return Math.round(v / step);
}

// (Optional) Convert a tick back to a value. Useful for debugging.
function tickToValue(tick, step) {
    if (!step || step <= 0) return tick;
    return tick * step;
}

function main() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;

    if (doc.selection.length < 1) {
        alert("オブジェクトを選択してください。");
        return;
    }

    // 1. 線が含まれる場合は OffsetPath を実行
    if (selectionHasStroke(doc.selection)) {
        app.executeMenuCommand('OffsetPath v22');
        // 再選択（念のため）
        try {
            var currentSel = doc.selection;
            doc.selection = null;
            doc.selection = currentSel;
        } catch (_) {}
    }

    // If only one object is selected, there is no overlap to merge.
    // Bake its opacity directly to avoid Divide/Expand collapsing transparency (e.g., CMYK K100 @ 50% becoming K100).
    if (doc.selection && doc.selection.length === 1) {
        try {
            bakeOpacityIntoFillRecursive(doc.selection[0], 1.0, doc);
        } catch (_) { }
        alert('処理が完了しました。');
        return;
    }

    // 2. 分割コマンドの実行
    app.executeMenuCommand('group');
    app.executeMenuCommand('Live Pathfinder Divide');
    app.executeMenuCommand('expandStyle');

    // --- 安定化のための処理 ---
    app.redraw(); // 描画を強制更新
    var workGroup = doc.selection[0];
    if (!workGroup || workGroup.typename !== "GroupItem") return;

    // パスアイテムをすべて抽出（この時点ではまだプロパティに深くアクセスしない）
    var rawItems = [];
    getAllPathItems(workGroup, rawItems);

    // 2. 情報を「安全な形式」で抽出する
    // Illustratorのアイテムは [0]が一番上、[last]が一番下
    var dataStack = [];
    for (var i = 0; i < rawItems.length; i++) {
        var item = rawItems[i];
        try {
            var gb = item.geometricBounds; // [left, top, right, bottom]
            var zpos = null;
            try { zpos = item.zOrderPosition; } catch (_) { zpos = null; }
            var stackIdx = getStackIndexInParent(item);

            dataStack.push({
                obj: item,

                // Use absolute area to avoid sign flips (CW/CCW). Store as integer ticks for stable keys.
                areaTick: toTick(Math.abs(item.area), AREA_TOL),

                // Use geometricBounds (more stable than left/top alone). Store as integer ticks for stable keys.
                gbLTick: toTick(gb[0], GEOM_TOL_PT),
                gbTTick: toTick(gb[1], GEOM_TOL_PT),
                gbRTick: toTick(gb[2], GEOM_TOL_PT),
                gbBTick: toTick(gb[3], GEOM_TOL_PT),

                points: item.pathPoints.length,
                opacity: item.opacity,
                fillColor: item.filled ? item.fillColor : null,

                // Depth: prefer actual parent stacking index (0=frontmost). Fallback to zOrderPosition, then extraction order.
                // We will sort by this so that larger value means "more back".
                depth: (stackIdx !== null) ? stackIdx : ((zpos !== null) ? zpos : i)
            });
        } catch (e) {
            // エラーが出るアイテムはスキップ
        }
    }

    // 3. 幾何学キーでグループ化
    var geometryMap = {};
    for (var j = 0; j < dataStack.length; j++) {
        var d = dataStack[j];
        var key = d.areaTick + "_" + d.gbLTick + "_" + d.gbTTick + "_" + d.gbRTick + "_" + d.gbBTick + "_" + d.points;
        if (!geometryMap[key]) geometryMap[key] = [];
        geometryMap[key].push(d);
    }

    // Composite a color over white using the same RGB linear-light path (used for overlap stacks).
    // This is used ONLY for the bottom-most item in an overlap group so its own opacity is respected.
    function blendOverWhiteForOverlap(col, alpha, doc) {
        if (!col) return col;
        if (!doc) return col;

        // If fully opaque, keep as-is to avoid unnecessary color conversions.
        if (alpha >= 0.999) return col;

        // Fast-path for CMYK neutral (pure K / gray axis): avoid RGB roundtrip which can distort K.
        // Example: K100 @ 50% should become roughly K50 over white.
        if (doc.documentColorSpace === DocumentColorSpace.CMYK && col.typename === 'CMYKColor') {
            var tol = 1e-6;
            if (Math.abs(col.cyan) < tol && Math.abs(col.magenta) < tol && Math.abs(col.yellow) < tol) {
                var cmyk = new CMYKColor();
                cmyk.cyan = 0;
                cmyk.magenta = 0;
                cmyk.yellow = 0;
                cmyk.black = col.black * alpha;
                return cmyk;
            }
        }

        var rgb = colorToRGB8(col);
        if (!rgb) return col;

        var rgb255 = [
            Math.max(0, Math.min(255, rgb[0])),
            Math.max(0, Math.min(255, rgb[1])),
            Math.max(0, Math.min(255, rgb[2]))
        ];
        var white = [255, 255, 255];
        var out = blendRGB8_linear(rgb255, white, alpha); // col over white
        return rgb8ToDocColor(doc, out);
    }

    // 4. 合成処理
    for (var k in geometryMap) {
        var group = geometryMap[k];

        // 深度（depth）が大きい順（＝背面→前面の順で処理できるように）にソート
        group.sort(function (a, b) { return b.depth - a.depth; });

        if (group.length > 1) {
            // 【重なりあり】
            // NOTE: Z-order preservation
            // ここでは「最前面（frontmost）」のオブジェクトを残すことで、周囲のオブジェクトとの前後関係が変わらないようにします。
            // 色の合成自体は背面→前面の順で行い、最終色を survivor（最前面）へ適用します。

            var backData = group[0];                 // 背面（backmost）
            var survivor = group[group.length - 1]; // 最前面（frontmost）

            // Respect the backmost object's own opacity by first compositing it over white.
            var baseColor = backData.fillColor;
            if (baseColor) {
                baseColor = blendOverWhiteForOverlap(baseColor, backData.opacity / 100, doc);
            }

            // Composite back→front
            for (var n = 1; n < group.length; n++) {
                var topData = group[n];
                if (baseColor && topData.fillColor) {
                    baseColor = blendColors(topData.fillColor, baseColor, topData.opacity / 100, doc);
                }
            }

            // Apply final color to the survivor (frontmost), set opacity to 100
            try {
                survivor.obj.fillColor = baseColor;
            } catch (_) {}
            try {
                survivor.obj.opacity = 100;
            } catch (_) {}

            // Remove all others (keep survivor to preserve stacking order)
            for (var m = 0; m < group.length; m++) {
                if (group[m] === survivor) continue;
                try { group[m].obj.remove(); } catch (_) {}
            }

        } else {
            // 【重なりなし（単品）】
            var single = group[0];
            if (single.fillColor) {
                single.obj.fillColor = blendWithWhite(single.fillColor, single.opacity / 100, doc);
            }
            single.obj.opacity = 100;
        }
    }

    alert("処理が完了しました。");
}
// Bake opacity into fills recursively for a single selection (no overlap case).
// This avoids menu operations (Divide/Expand) that can collapse transparency and lose opacity values.
function bakeOpacityIntoFillRecursive(item, parentAlpha, doc) {
    if (!item) return;
    if (parentAlpha === undefined || parentAlpha === null) parentAlpha = 1.0;

    var t = item.typename;

    if (t === 'GroupItem') {
        var a = parentAlpha;
        try { a = a * (item.opacity / 100); } catch (_) { }

        // Recurse into children
        try {
            for (var i = 0; i < item.pageItems.length; i++) {
                bakeOpacityIntoFillRecursive(item.pageItems[i], a, doc);
            }
        } catch (_) { }

        // Normalize group opacity
        try { item.opacity = 100; } catch (_) { }
        return;
    }

    if (t === 'CompoundPathItem') {
        // CompoundPathItem contains pathItems
        var a2 = parentAlpha;
        try { a2 = a2 * (item.opacity / 100); } catch (_) { }
        try {
            for (var j = 0; j < item.pathItems.length; j++) {
                bakeOpacityIntoFillRecursive(item.pathItems[j], a2, doc);
            }
        } catch (_) { }
        try { item.opacity = 100; } catch (_) { }
        return;
    }

    if (t === 'PathItem') {
        var a3 = parentAlpha;
        try { a3 = a3 * (item.opacity / 100); } catch (_) { }

        // Only bake fill; keep stroke as-is (this script focuses on fill flattening)
        try {
            if (item.filled && item.fillColor) {
                item.fillColor = blendWithWhite(item.fillColor, a3, doc);
            }
        } catch (_) { }

        // Normalize opacity
        try { item.opacity = 100; } catch (_) { }
        return;
    }

    // Other types (TextFrame, PlacedItem, etc.) are ignored.
}

// Get stacking index within the immediate parent (GroupItem or Layer).
// Illustrator’s pageItems are ordered with [0] as frontmost and [last] as backmost.
// Using this index is more reliable than zOrderPosition across versions.
function getStackIndexInParent(it) {
    if (!it) return null;
    var p = null;
    try { p = it.parent; } catch (_) { p = null; }
    if (!p) return null;

    var items = null;
    try { items = p.pageItems; } catch (_) { items = null; }
    if (!items) return null;

    try {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === it) return i;
        }
    } catch (_) {
        // ignore
    }
    return null;
}

// 再帰的にパスを取得
function getAllPathItems(container, resultAry) {
    var items = container.pageItems;
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (item.typename === "PathItem") {
            resultAry.push(item);
        } else if (item.typename === "GroupItem") {
            getAllPathItems(item, resultAry);
        }
    }
}

// --- Color conversion & gamma-correct compositing ---
// Illustrator opacity compositing is effectively done in RGB display space.
// If we blend channel values directly (especially in CMYK, or sRGB without gamma correction),
// results tend to look darker than Illustrator’s visual result.
// We therefore convert to RGB, blend in linear light, then convert back to document space.

// --- Neutral CMYK (K-only) blending helpers ---
function _isNeutralCMYK(col) {
    if (!col || col.typename !== 'CMYKColor') return false;
    var tol = 1e-6;
    return (Math.abs(col.cyan) < tol && Math.abs(col.magenta) < tol && Math.abs(col.yellow) < tol);
}

function _neutralCMYK_to_gray255(col) {
    // Map K-only CMYK to sRGB gray (no profile roundtrip): K=0 -> 255 (white), K=100 -> 0 (black)
    var k = Math.max(0, Math.min(100, col.black));
    return 255 * (1 - k / 100);
}

function _gray255_to_neutralCMYK(g255) {
    var g = Math.max(0, Math.min(255, g255));
    var k = 100 * (1 - g / 255);
    var cmyk = new CMYKColor();
    cmyk.cyan = 0;
    cmyk.magenta = 0;
    cmyk.yellow = 0;
    cmyk.black = Math.max(0, Math.min(100, k));
    return cmyk;
}

function _srgb8_to_linear01(v8) {
    var v = v8 / 255;
    // Simple gamma 2.2 approximation (good enough for this use case)
    return Math.pow(v, 2.2);
}
function _linear01_to_srgb8(v01) {
    var v = Math.max(0, Math.min(1, v01));
    return Math.round(Math.pow(v, 1 / 2.2) * 255);
}

function colorToRGB8(col) {
    // Returns [r,g,b] in 0..255
    if (!col) return null;

    if (col.typename === "RGBColor") {
        return [col.red, col.green, col.blue];
    }
    if (col.typename === "CMYKColor") {
        try {
            // convertSampleColor returns numbers in target space ranges (RGB: 0..255)
            var rgb = app.convertSampleColor(
                ImageColorSpace.CMYK,
                [col.cyan, col.magenta, col.yellow, col.black],
                ImageColorSpace.RGB,
                ColorConvertPurpose.defaultpurpose
            );
            return [rgb[0], rgb[1], rgb[2]];
        } catch (_) {
            // Fallback: naive conversion (rarely used)
            var r = 255 * (1 - col.cyan / 100) * (1 - col.black / 100);
            var g = 255 * (1 - col.magenta / 100) * (1 - col.black / 100);
            var b = 255 * (1 - col.yellow / 100) * (1 - col.black / 100);
            return [r, g, b];
        }
    }
    // Unsupported color types (Spot/Pattern/Gradient etc.)
    return null;
}

function rgb8ToDocColor(doc, rgb8) {
    // Returns a Color object matching the document color space (CMYK doc => CMYKColor, RGB doc => RGBColor)
    if (!rgb8) return null;

    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
        try {
            var cmyk = app.convertSampleColor(
                ImageColorSpace.RGB,
                [rgb8[0], rgb8[1], rgb8[2]],
                ImageColorSpace.CMYK,
                ColorConvertPurpose.defaultpurpose
            );
            var c = new CMYKColor();
            c.cyan = cmyk[0];
            c.magenta = cmyk[1];
            c.yellow = cmyk[2];
            c.black = cmyk[3];
            return c;
        } catch (_) {
            // Fallback: assign RGB (Illustrator will convert internally)
            var r = new RGBColor();
            r.red = rgb8[0]; r.green = rgb8[1]; r.blue = rgb8[2];
            return r;
        }
    } else {
        var r2 = new RGBColor();
        r2.red = rgb8[0]; r2.green = rgb8[1]; r2.blue = rgb8[2];
        return r2;
    }
}

function blendRGB8_linear(topRGB8, bottomRGB8, alpha) {
    // Alpha compositing in linear light: out = top*alpha + bottom*(1-alpha)
    var inv = 1.0 - alpha;

    var tr = _srgb8_to_linear01(topRGB8[0]);
    var tg = _srgb8_to_linear01(topRGB8[1]);
    var tb = _srgb8_to_linear01(topRGB8[2]);

    var br = _srgb8_to_linear01(bottomRGB8[0]);
    var bg = _srgb8_to_linear01(bottomRGB8[1]);
    var bb = _srgb8_to_linear01(bottomRGB8[2]);

    var or = tr * alpha + br * inv;
    var og = tg * alpha + bg * inv;
    var ob = tb * alpha + bb * inv;

    return [_linear01_to_srgb8(or), _linear01_to_srgb8(og), _linear01_to_srgb8(ob)];
}

// 色の合成（通常）
function blendColors(topCol, bottomCol, alpha, doc) {
    // Default: direct blending in the same color space (previous behavior).
    // Optional: gamma-correct RGB blending if USE_GAMMA_CORRECT_BLEND is true.
    var inv = 1.0 - alpha;

    if (!topCol) return bottomCol;
    if (!bottomCol) return topCol;

    // CMYK neutral (K-only) fast path: blend in gray RGB space to avoid overly-dark stacking.
    if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK && _isNeutralCMYK(topCol) && _isNeutralCMYK(bottomCol)) {
        var topGray = _neutralCMYK_to_gray255(topCol);
        var bottomGray = _neutralCMYK_to_gray255(bottomCol);
        // Blend gray as RGB triplet in linear light
        var out = blendRGB8_linear([topGray, topGray, topGray], [bottomGray, bottomGray, bottomGray], alpha);
        // out is [r,g,b] (float-ish). Use the red channel as gray.
        return _gray255_to_neutralCMYK(out[0]);
    }

    if (!USE_GAMMA_CORRECT_BLEND) {
        if (topCol.typename === "CMYKColor" && bottomCol.typename === "CMYKColor") {
            var c = new CMYKColor();
            c.cyan = topCol.cyan * alpha + bottomCol.cyan * inv;
            c.magenta = topCol.magenta * alpha + bottomCol.magenta * inv;
            c.yellow = topCol.yellow * alpha + bottomCol.yellow * inv;
            c.black = topCol.black * alpha + bottomCol.black * inv;
            return c;
        } else if (topCol.typename === "RGBColor" && bottomCol.typename === "RGBColor") {
            var r = new RGBColor();
            r.red = Math.round(topCol.red * alpha + bottomCol.red * inv);
            r.green = Math.round(topCol.green * alpha + bottomCol.green * inv);
            r.blue = Math.round(topCol.blue * alpha + bottomCol.blue * inv);
            return r;
        }
        // Unsupported/other types -> keep top
        return topCol;
    }

    // Gamma-correct blend in RGB, then convert back to doc space.
    // `doc` is required for proper output color space.
    if (!doc) return topCol;

    var topRGB8 = colorToRGB8(topCol);
    var bottomRGB8 = colorToRGB8(bottomCol);
    if (!topRGB8 || !bottomRGB8) {
        // Unsupported color types -> fall back to keeping the top color
        return topCol;
    }

    var outRGB8 = blendRGB8_linear(topRGB8, bottomRGB8, alpha);
    return rgb8ToDocColor(doc, outRGB8);
}

// 白との合成
function blendWithWhite(col, alpha, doc) {
    var inv = 1.0 - alpha;
    if (!col) return col;

    if (USE_RGB_WHITE_COMPOSITE && doc) {
        // Composite in RGB linear light: out = col*alpha + white*(1-alpha)
        var rgb = colorToRGB8(col);
        if (rgb) {
            // Ensure numbers are within 0..255
            var rgb8 = [
                Math.max(0, Math.min(255, rgb[0])),
                Math.max(0, Math.min(255, rgb[1])),
                Math.max(0, Math.min(255, rgb[2]))
            ];
            var white = [255, 255, 255];
            var out = blendRGB8_linear(rgb8, white, alpha);
            return rgb8ToDocColor(doc, out);
        }
        // If unsupported color type, fall through to direct method below.
    }

    // Fallback: previous direct method (keeps hues stable, but can be visually darker for CMYK)
    if (col.typename === "CMYKColor") {
        var c = new CMYKColor();
        c.cyan = col.cyan * alpha;
        c.magenta = col.magenta * alpha;
        c.yellow = col.yellow * alpha;
        c.black = col.black * alpha;
        return c;
    } else if (col.typename === "RGBColor") {
        var r = new RGBColor();
        r.red = Math.round(col.red * alpha + 255 * inv);
        r.green = Math.round(col.green * alpha + 255 * inv);
        r.blue = Math.round(col.blue * alpha + 255 * inv);
        return r;
    }
    return col;
}

// Helper function to check if selection has any stroke
function selectionHasStroke(selection) {
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.stroked && item.strokeWidth > 0) {
            return true;
        }
    }
    return false;
}

// Helper function to run menu command safely
function runMenu(commandName) {
    try {
        app.executeMenuCommand(commandName);
    } catch (e) {
        // ignore errors
    }
}

main();
