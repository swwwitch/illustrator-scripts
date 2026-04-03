/*
  RandomFonts.jsx
  選択したテキストの1文字ごとにランダムなフォントを適用するスクリプト

  - ダイアログボックスを表示
  - ［再実行］ボタンでランダム適用（プレビュー更新）
*/
(function () {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    // インストールされている全フォントを取得（数が多いと処理に時間がかかる場合があります）
    var allFonts = app.textFonts;

var fontCount = allFonts.length;

// --- Background rectangles (per-character via outline) ---
var BG_LAYER_NAME = '__FontShuffle_BG__';
var BG_GROUP_NAME = '__BGRects__';
var RANSOM_LAYER_NAME = '__FontShuffle_Ransom__';

function ensureRansomLayer() {
    var lyr = null;
    for (var i = 0; i < doc.layers.length; i++) {
        try {
            if (doc.layers[i].name === RANSOM_LAYER_NAME) { lyr = doc.layers[i]; break; }
        } catch (_) { }
    }
    if (!lyr) {
        lyr = doc.layers.add();
        lyr.name = RANSOM_LAYER_NAME;
    }
    return lyr;
}
function createRansomGroups(frames) {
    if (!frames || frames.length === 0) return;

    // clear existing background rects (not used in ransom output)
    clearBackgroundRectsIfAny();

    var ransomLayer = ensureRansomLayer();

    function isSkippableChar(ch) {
        // outline側に出てこない文字は除外してインデックスずれを防ぐ
        return (ch === '\r' || ch === '\n' || ch === ' ' || ch === '\t');
    }

    for (var i = 0; i < frames.length; i++) {
        var tf = frames[i];
        if (!tf || tf.typename !== 'TextFrame') continue;

        // Outline once to obtain per-character bounds units
        var dup = null;
        var outlined = null;

        try {
            dup = tf.duplicate();
            try { dup.move(tf, ElementPlacement.PLACEAFTER); } catch (_) { }
        } catch (_) {
            continue;
        }

        try {
            outlined = dup.createOutline();
        } catch (_) {
            try { if (dup) dup.remove(); } catch (_) { }
            continue;
        }

        try { if (dup) dup.remove(); } catch (_) { }

        // Collect character outline units (sub-level groupItems)
        var units = [];
        collectCharGroupsFromOutlined(outlined, units);

        // Collect visible characters (skip returns; keep spaces as characters if they exist in units)
        var chars = [];
        try {
            var trChars = tf.textRange.characters;
            for (var ci = 0; ci < trChars.length; ci++) {
                try {
                    var cc = trChars[ci].contents;
                    if (isSkippableChar(cc)) continue;
                    chars.push(trChars[ci]);
                } catch (_) { }
            }
        } catch (_) { }

        var count = Math.min(units.length, chars.length);

        for (var u = 0; u < count; u++) {
            var unit = units[u];
            var chObj = chars[u];
            if (!unit || !chObj) continue;

            // Use outline unit bounds as placement reference
            var b;
            try { b = unit.geometricBounds; } catch (_) { continue; }
            // b: [L,T,R,B]
            var L = b[0];
            var T = b[1];

            // A) Create per-character live text (point text)
            // Place at a neutral position first, then translate so bounds match the outlined glyph bounds.
            var pt = null;
            try {
                pt = doc.textFrames.pointText([0, 0]);
                pt.contents = chObj.contents;
                pt.move(ransomLayer, ElementPlacement.PLACEATEND);

                // Copy basic character attributes (use preview state)
                try {
                    var src = chObj.characterAttributes;
                    var dst = pt.textRange.characterAttributes;
                    try { dst.textFont = src.textFont; } catch (_) { }
                    try { dst.size = src.size; } catch (_) { }
                    try { dst.tracking = src.tracking; } catch (_) { }
                    try { dst.fillColor = src.fillColor; } catch (_) { }
                    try { dst.strokeColor = src.strokeColor; } catch (_) { }
                    try { dst.strokeWeight = src.strokeWeight; } catch (_) { }
                } catch (_) { }

                // Align by geometric bounds (top-left)
                try {
                    var pb = pt.geometricBounds; // [L,T,R,B]
                    var dx = L - pb[0];
                    var dy = T - pb[1];
                    pt.translate(dx, dy);
                } catch (_) { }
            } catch (_) {
                try { if (pt) pt.remove(); } catch (_) { }
                continue;
            }

            // Group container
            var g = null;
            try {
                g = ransomLayer.groupItems.add();
                g.name = '__RansomChar__';

                // Put rect first (behind)
                addRectFromBounds(g, b); // C+D+E (PAD=1 + jitter)

                // Move live text into the group above the rect
                try { pt.move(g, ElementPlacement.PLACEATEND); } catch (_) { }

                // Ensure rect behind
                try { g.pathItems[0].zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
            } catch (_) {
                try { if (g) g.remove(); } catch (_) { }
            }
        }

        // B outline container cleanup (F)
        try { if (outlined) outlined.remove(); } catch (_) { }

        // Remove original text frame (since we split into per-character text)
        try { tf.remove(); } catch (_) { }
    }
}

function ensureBgLayer() {
    var lyr = null;
    for (var i = 0; i < doc.layers.length; i++) {
        try {
            if (doc.layers[i].name === BG_LAYER_NAME) { lyr = doc.layers[i]; break; }
        } catch (_) { }
    }
    if (!lyr) {
        lyr = doc.layers.add();
        lyr.name = BG_LAYER_NAME;
    }
    return lyr;
}

function removeBgGroupIfExists(bgLayer) {
    try {
        for (var i = bgLayer.groupItems.length - 1; i >= 0; i--) {
            if (bgLayer.groupItems[i].name === BG_GROUP_NAME) {
                bgLayer.groupItems[i].remove();
            }
        }
    } catch (_) { }
}

function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function makeRandomFill() {
    var g = new GrayColor();
    // GrayColor.gray: 0 = white, 100 = black
    // K10〜K50 の間でランダム
    g.gray = randInt(10, 50);
    return g;
}

function setRectStyle(rect) {
    try {
        rect.stroked = false;
        rect.filled = true;
        rect.fillColor = makeRandomFill();
    } catch (_) { }
}

function jitterRectOutward(rect, maxPt) {
    if (!rect || rect.typename !== 'PathItem') return;
    if (!rect.pathPoints || rect.pathPoints.length < 4) return;

    var b;
    try { b = rect.geometricBounds; } catch (_) { return; }
    // b: [L, T, R, B]
    var cx = (b[0] + b[2]) / 2;
    var cy = (b[1] + b[3]) / 2;

    for (var i = 0; i < rect.pathPoints.length; i++) {
        try {
            var pp = rect.pathPoints[i];
            var ax = pp.anchor[0], ay = pp.anchor[1];
            var dx = ax - cx;
            var dy = ay - cy;
            var len = Math.sqrt(dx * dx + dy * dy);
            if (!len || len === 0) continue;

            // random outward distance
            var d = Math.random() * maxPt;
            var ux = dx / len;
            var uy = dy / len;
            var ox = ux * d;
            var oy = uy * d;

            // move anchor + handles together to keep corner shape
            pp.anchor = [ax + ox, ay + oy];
            pp.leftDirection = [pp.leftDirection[0] + ox, pp.leftDirection[1] + oy];
            pp.rightDirection = [pp.rightDirection[0] + ox, pp.rightDirection[1] + oy];
        } catch (_) { }
    }
}

function addRectFromBounds(bgGroup, b) {
    // b: [L, T, R, B]
    var PAD = 1; // 1pt expansion on each side

    var L = b[0] - PAD;
    var T = b[1] + PAD;
    var R = b[2] + PAD;
    var B = b[3] - PAD;

    var w = R - L;
    var h = T - B;
    if (w <= 0 || h <= 0) return null;

    var rect = bgGroup.pathItems.rectangle(T, L, w, h);
    setRectStyle(rect);
    // Randomly expand anchors outward (rough, organic feel)
    jitterRectOutward(rect, 1);
    return rect;
}

function collectCharGroupsFromOutlined(outlined, out) {
    // C: アウトライン結果の「サブレベル groupItems」を 1文字単位として扱う。
    // 典型例: outlined(GroupItem)
    //   ├─ groupItems[0] (行/ブロック)
    //   │    ├─ groupItems[0] (1文字)
    //   │    ├─ groupItems[1] (1文字)
    //   │    ...
    //   ├─ groupItems[1] (行/ブロック)
    //   ...
    // 例外ケースに備えてフォールバックも用意。
    if (!outlined) return;

    function pushLeafUnits(container) {
        try {
            if (!container || !container.pageItems) return;
            for (var k = 0; k < container.pageItems.length; k++) {
                var it = container.pageItems[k];
                if (!it) continue;
                if (it.typename === 'GroupItem') {
                    pushLeafUnits(it);
                } else if (it.typename === 'PathItem' || it.typename === 'CompoundPathItem') {
                    out.push(it);
                }
            }
        } catch (_) { }
    }

    try {
        if (outlined.typename === 'GroupItem') {
            var hasPushed = false;

            // Prefer: sub-level groupItems (child group -> its groupItems)
            if (outlined.groupItems && outlined.groupItems.length > 0) {
                for (var i = 0; i < outlined.groupItems.length; i++) {
                    var g1 = outlined.groupItems[i];
                    if (!g1 || g1.typename !== 'GroupItem') continue;

                    if (g1.groupItems && g1.groupItems.length > 0) {
                        for (var j = 0; j < g1.groupItems.length; j++) {
                            out.push(g1.groupItems[j]);
                            hasPushed = true;
                        }
                    } else {
                        // If the child has no further groups, use it as a unit
                        out.push(g1);
                        hasPushed = true;
                    }
                }
            }

            if (hasPushed) return;

            // Fallback: leaf units so it won't collapse to a single unit
            pushLeafUnits(outlined);
            if (out.length > 0) return;

            // Last resort: whole outlined
            out.push(outlined);
            return;
        }

        out.push(outlined);
    } catch (_) { }
}

function createBackgroundRectsByOutlining(frames) {
    if (!frames || frames.length === 0) return;

    var bgLayer = ensureBgLayer();
    removeBgGroupIfExists(bgLayer);
    var bgGroup = bgLayer.groupItems.add();
    bgGroup.name = BG_GROUP_NAME;

    for (var i = 0; i < frames.length; i++) {
        var tf = frames[i];
        if (!tf || tf.typename !== 'TextFrame') continue;

        var dup = null;
        var outlined = null;

        // A) duplicate behind original
        try {
            dup = tf.duplicate();
            try { dup.move(tf, ElementPlacement.PLACEAFTER); } catch (_) { }
        } catch (_) {
            continue;
        }

        // B) outline the duplicate
        try {
            outlined = dup.createOutline();
        } catch (_) {
            try { if (dup) dup.remove(); } catch (_) { }
            continue;
        }

        // remove duplicate text
        try { if (dup) dup.remove(); } catch (_) { }

        // C) treat each top-level group as a character group
        var charGroups = [];
        collectCharGroupsFromOutlined(outlined, charGroups);

        // D) rectangle per character group, random fill
        for (var j = 0; j < charGroups.length; j++) {
            try {
                var b = charGroups[j].geometricBounds;
                addRectFromBounds(bgGroup, b);
            } catch (_) { }
        }

        // E) delete outlined artwork (C)
        try { if (outlined) outlined.remove(); } catch (_) { }
    }

    // Keep backgrounds behind content
    try { bgGroup.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
    try { bgLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
}

function setTrackingForFrames(frames, trackingVal) {
    if (!frames || frames.length === 0) return;
    for (var i = 0; i < frames.length; i++) {
        var item = frames[i];
        if (!item || item.typename !== 'TextFrame') continue;
        try {
            // tracking is in 1/1000 em
            item.textRange.characterAttributes.tracking = trackingVal;
        } catch (_) { }
    }
}

function getBgLayerIfExists() {
    try {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i] && doc.layers[i].name === BG_LAYER_NAME) return doc.layers[i];
        }
    } catch (_) { }
    return null;
}

function clearBackgroundRectsIfAny() {
    var lyr = getBgLayerIfExists();
    if (!lyr) return;
    removeBgGroupIfExists(lyr);
}

    function getMorisawaFonts() {
        var list = [];
        var i, f;

        function matchJAKeyword(s) {
            if (!s) return false;
            s = String(s);
            // 和文フォント判定キーワード
            // Pr6 / Pr6N / AB- / FOT を含むものを対象
            return (
                s.indexOf('Pr6N') !== -1 ||
                s.indexOf('Pr6') !== -1 ||
                s.indexOf('AB-') !== -1 ||
                s.indexOf('FOT') !== -1 ||
                s.indexOf('-OTF') !== -1
            );
        }

        for (i = 0; i < allFonts.length; i++) {
            try {
                f = allFonts[i];
                if (!f) continue;

                var n1 = f.fullName ? String(f.fullName) : '';
                var n2 = f.name ? String(f.name) : '';
                var n3 = (f.postScriptName) ? String(f.postScriptName) : '';

                if (matchJAKeyword(n1) || matchJAKeyword(n2) || matchJAKeyword(n3)) {
                    list.push(f);
                }
            } catch (_) { }
        }

        return list;
    }

    function getSelectionTextFrames() {
        var sel = doc.selection;
        if (!sel || sel.length === 0) return [];
        var frames = [];
        for (var i = 0; i < sel.length; i++) {
            try {
                if (sel[i] && sel[i].typename === "TextFrame") frames.push(sel[i]);
            } catch (_) { }
        }
        return frames;
    }

    function applyRandomFontsToFrames(frames, fontList) {
        if (!frames || frames.length === 0) return;
        if (!fontList || fontList.length === 0) return;

        for (var i = 0; i < frames.length; i++) {
            var item = frames[i];
            if (!item || item.typename !== "TextFrame") continue;

            var chars = item.textRange.characters;
            for (var j = 0; j < chars.length; j++) {
                // 改行や空白文字はスキップ（エラー回避と見た目のため）
                if (chars[j].contents === "\r" || chars[j].contents === " ") continue;

                try {
                    var randomFontIndex = Math.floor(Math.random() * fontList.length);
                    chars[j].characterAttributes.textFont = fontList[randomFontIndex];
                } catch (_) {
                    // 特定のフォントが適用できない場合は無視
                }
            }
        }

        app.redraw();
    }

    // --- Dialog ---
    var dlg = new Window('dialog', 'Random Fonts');
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    var info = dlg.add('statictext', undefined, '選択したテキストに対して、1文字ごとにランダムなフォントを適用します。');
    info.characters = 52;

    var chkMorisawa = dlg.add('checkbox', undefined, '和文フォントに限る');
    chkMorisawa.value = false;
    var chkManifesto = dlg.add('checkbox', undefined, '犯行声明文風');
    chkManifesto.value = false;

    // --- Bottom bar (3 columns) ---
    var bottom = dlg.add('group');
    bottom.orientation = 'row';
    bottom.alignChildren = ['fill', 'center'];

    // Left
    var colL = bottom.add('group');
    colL.orientation = 'row';
    colL.alignChildren = ['left', 'center'];
    var btnRerun = colL.add('button', undefined, '再実行');

    // Center spacer
    var colC = bottom.add('group');
    colC.orientation = 'row';
    colC.alignment = ['fill', 'center'];
    colC.add('statictext', undefined, '');

    // Right
    var colR = bottom.add('group');
    colR.orientation = 'row';
    colR.alignChildren = ['right', 'center'];
    var btnCancel = colR.add('button', undefined, 'キャンセル', { name: 'cancel' });
    var btnOK = colR.add('button', undefined, 'OK', { name: 'ok' });

    try { colC.preferredSize.width = 20; } catch (_) { }
    try { bottom.alignment = ['fill', 'bottom']; } catch (_) { }
    try { colR.alignment = ['right', 'center']; } catch (_) { }

    // Preview state: 直前のプレビューを undo で戻してから再適用する
    var _previewApplied = false;

    function runPreview(makeRects) {
        function _do() {
            var frames = getSelectionTextFrames();
            if (frames.length === 0) {
                alert('テキストオブジェクトを選択してください。');
                return;
            }

            if (_previewApplied) {
                try { app.undo(); } catch (_) { }
            }

            var fontList = null;
            if (chkMorisawa && chkMorisawa.value) {
                fontList = getMorisawaFonts();
                if (!fontList || fontList.length === 0) {
                    alert('対象の和文フォント（Pr6 / Pr6N）が見つかりません。');
                    return;
                }
            } else {
                fontList = allFonts;
            }

            // apply random fonts
            applyRandomFontsToFrames(frames, fontList);

            // 犯行声明文風: tracking=200 + text color white
            if (chkManifesto && chkManifesto.value) {
                setTrackingForFrames(frames, 200);
                try {
                    var white = new RGBColor();
                    white.red = 255;
                    white.green = 255;
                    white.blue = 255;
                    for (var ti = 0; ti < frames.length; ti++) {
                        try { frames[ti].textRange.characterAttributes.fillColor = white; } catch (_) { }
                    }
                } catch (_) { }
            }

            // Background rectangles: only when requested AND 犯行声明文風 ON
            if (makeRects && chkManifesto && chkManifesto.value) {
                createBackgroundRectsByOutlining(frames);
            } else {
                // OFF の場合は背景を付けない（既存があれば消す）
                clearBackgroundRectsIfAny();
            }

            _previewApplied = true;
        }

        // Try to keep the operation to a single history step (so undo once works)
        try {
            if (doc && doc.suspendHistory) {
                doc.suspendHistory('FontShuffle Preview', '_do()');
                return;
            }
        } catch (_) { }

        _do();
    }


    btnRerun.onClick = function () {
        // プレビュー更新（同じ扱い）
        runPreview(false);
    };

    btnOK.onClick = function () {
        if (chkManifesto && chkManifesto.value) {
            // ON：createRansomGroups(...) を必ず実行
            // まずプレビュー未実行ならフォント適用＋トラッキング等だけ反映（長方形は付けない）
            if (!_previewApplied) {
                runPreview(false);
            }

            try {
                if (doc && doc.suspendHistory) {
                    doc.suspendHistory('FontShuffle Ransom Output', 'createRansomGroups(getSelectionTextFrames())');
                } else {
                    createRansomGroups(getSelectionTextFrames());
                }
            } catch (_) {
                try { createRansomGroups(getSelectionTextFrames()); } catch (_) { }
            }
        } else {
            // OFF：通常の runPreview(true)（背景だけ）
            runPreview(true);
        }

        dlg.close(1);
    };

    btnCancel.onClick = function () {
        // プレビューを戻す
        if (_previewApplied) {
            try { app.undo(); } catch (_) { }
        }
        dlg.close(0);
    };

    dlg.center();
    dlg.show();
})();
