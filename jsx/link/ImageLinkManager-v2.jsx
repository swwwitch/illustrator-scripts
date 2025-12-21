#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


// ===== Script info =====
var SCRIPT_NAME = 'ImageLinkManager';
var SCRIPT_VERSION = '1.0.0';


// 【概要 / Summary】
// 配置画像（PlacedItem）を埋め込み（embed）に変換する／解除（埋め込み解除：RasterItem をリンク配置に戻す）するスクリプトです。
// ダイアログ上部のモードで「埋め込み / 解除」を切り替え、対応パネルのみ有効化（それ以外はディム表示）します。
// それぞれのパネルで処理対象を選択できます：
// ・選択した画像のみ（選択オブジェクト配下を再帰探索）
// ・すべての画像（ドキュメント内）
//
// 【埋め込み】
// PlacedItem を embed() します。PSD（.psd）の場合は embed() 後に ungroup を試行します。
// ただし親グループがクリッピンググループ（clipped === true）の場合は ungroup を行いません。
//
// 【解除】
// 埋め込み画像（RasterItem）を PSD に書き出し、リンク配置（PlacedItem）へ置き換えます。
// 可能な範囲で位置・回転・拡大率を合わせます（効果が適用されている RasterItem では差異が出る場合があります）。
//
// 実行後、選択を解除します。
// 更新日 / Updated: 2025-12-20

function showDialog() {
    var dlg = new Window('dialog', SCRIPT_NAME + ' v' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = [15, 20, 15, 10];

    // ===== モード選択（上部）/ Mode (top) =====
    var modeWrap = dlg.add('group');
    modeWrap.orientation = 'row';
    modeWrap.alignment = 'center';

    var modeGroup = modeWrap.add('group');
    modeGroup.orientation = 'row';
    modeGroup.alignChildren = ['left', 'center'];

    var rbModeEmbed = modeGroup.add('radiobutton', undefined, '埋め込み');
    var rbModeRelease = modeGroup.add('radiobutton', undefined, '解除');
    var rbModeReset = modeGroup.add('radiobutton', undefined, 'リセット');
    rbModeEmbed.value = true; // デフォルト

    // ===== 2 columns wrapper =====
    var columns = dlg.add('group');
    columns.orientation = 'row';
    columns.alignChildren = ['fill', 'top'];

    var leftCol = columns.add('group');
    leftCol.orientation = 'column';
    leftCol.alignChildren = ['fill', 'top'];

    var rightCol = columns.add('group');
    rightCol.orientation = 'column';
    rightCol.alignChildren = ['fill', 'top'];

    // ===== 埋め込み / Embed =====
    var panelEmbed = leftCol.add('panel', undefined, '埋め込み');
    panelEmbed.orientation = 'column';
    panelEmbed.alignChildren = 'left';
    panelEmbed.margins = [15, 20, 15, 10];

    var rbEmbedSelection = panelEmbed.add('radiobutton', undefined, '選択した配置画像のみ');
    var rbEmbedAll = panelEmbed.add('radiobutton', undefined, 'すべての配置画像');
    rbEmbedSelection.value = true; // デフォルト

    // 選択がない場合は「すべての配置画像」を自動選択
    var hasSelection = false;
    try {
        hasSelection = (app.documents.length > 0 && app.activeDocument.selection && app.activeDocument.selection.length > 0);
    } catch (e) {
        hasSelection = false;
    }
    if (!hasSelection) {
        rbEmbedSelection.value = false;
        rbEmbedAll.value = true;
    }

    // ===== 解除 / Release =====
    var panelRelease = leftCol.add('panel', undefined, '解除');
    panelRelease.orientation = 'column';
    panelRelease.alignChildren = 'left';
    panelRelease.margins = [15, 20, 15, 10];

    var rbReleaseSelection = panelRelease.add('radiobutton', undefined, '選択した画像のみ');
    var rbReleaseAll = panelRelease.add('radiobutton', undefined, 'すべての画像');
    rbReleaseSelection.value = true; // デフォルト

    // 選択がない場合は「すべての画像」を自動選択（解除）
    if (!hasSelection) {
        rbReleaseSelection.value = false;
        rbReleaseAll.value = true;
    }

    // ===== リセット / Reset (right column) =====
    var panelReset = rightCol.add('panel', undefined, 'リセット');
    panelReset.orientation = 'column';
    panelReset.alignChildren = ['left', 'top'];
    panelReset.margins = [15, 20, 15, 10];

    // （ResetTransform のUIを右カラムに移植：ロジックは後で接続）
    var cbReplaceReset = panelReset.add('checkbox', undefined, '再配置');
    cbReplaceReset.value = false;

    var cbRotate = panelReset.add('checkbox', undefined, '回転');
    var cbSkew = panelReset.add('checkbox', undefined, 'シアー');
    var cbRatio = panelReset.add('checkbox', undefined, '縦横比');
    var cbFlip = panelReset.add('checkbox', undefined, '反転');
    var cbScale = panelReset.add('checkbox', undefined, 'スケール');

    cbRotate.value = true;
    cbSkew.value = true;
    cbRatio.value = true;
    cbFlip.value = true;
    cbScale.value = false;

    // スケール（%）
    var gScale = panelReset.add('group');
    gScale.orientation = 'row';
    gScale.alignChildren = 'center';

    var etScale = gScale.add('edittext', undefined, '100');
    etScale.characters = 5;
    var stPercent = gScale.add('statictext', undefined, '%');

    // Arrow key increment helper（ResetTransform 互換）
    function changeValueByArrowKey(editText) {
        editText.addEventListener('keydown', function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName == 'Up') {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value = Math.round((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName == 'Up') {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == 'Up') {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

    // スケールUIの有効/無効
    etScale.enabled = cbScale.value;
    stPercent.enabled = cbScale.value;
    cbScale.onClick = function () {
        var on = cbScale.value;
        etScale.enabled = on;
        stPercent.enabled = on;
        if (on) {
            try {
                etScale.active = true;
                if (typeof etScale.select === 'function') {
                    etScale.select();
                } else if (typeof etScale.textselection !== 'undefined') {
                    etScale.textselection = etScale.text;
                }
            } catch (e) { }
        }
    };
    changeValueByArrowKey(etScale);

    // 再配置ONのときは他をディム
    function updateResetPanelEnabled() {
        var othersOn = !cbReplaceReset.value;
        cbRotate.enabled = othersOn;
        cbSkew.enabled = othersOn;
        cbRatio.enabled = othersOn;
        cbFlip.enabled = othersOn;
        cbScale.enabled = othersOn;

        etScale.enabled = othersOn && cbScale.value;
        stPercent.enabled = othersOn && cbScale.value;
    }
    cbReplaceReset.onClick = function () { updateResetPanelEnabled(); };
    updateResetPanelEnabled();

    // UI：パネルのディム切替
    function updatePanelState() {
        panelEmbed.enabled = rbModeEmbed.value;
        panelRelease.enabled = rbModeRelease.value;
        panelReset.enabled = rbModeReset.value;
    }
    rbModeEmbed.onClick = updatePanelState;
    rbModeRelease.onClick = updatePanelState;
    rbModeReset.onClick = updatePanelState;
    updatePanelState();

    // ===== Buttons (center) =====
    var btnGroup = dlg.add('group');
    btnGroup.orientation = 'row';
    btnGroup.alignment = 'center';
    btnGroup.alignChildren = ['center', 'center'];

    var cancelBtn = btnGroup.add('button', undefined, 'キャンセル', { name: 'cancel' });
    var okBtn = btnGroup.add('button', undefined, 'OK', { name: 'ok' });

    okBtn.onClick = function () {
        dlg.close(1);
    };
    cancelBtn.onClick = function () {
        dlg.close(0);
    };

    var result = dlg.show();
    if (result !== 1) return null;

    return {
        mode: rbModeEmbed.value ? 'embed' : (rbModeRelease.value ? 'release' : 'reset'),
        embedMode: rbEmbedAll.value ? 'all' : 'selection',
        releaseMode: rbReleaseAll.value ? 'all' : 'selection',
        resetOptions: {
            rotate: cbRotate.value,
            skew: cbSkew.value,
            ratio: cbRatio.value,
            flip: cbFlip.value,
            replaceReset: cbReplaceReset.value,
            scale: cbScale.value,
            scalePercent: (function () {
                var n = parseFloat(etScale.text, 10);
                if (isNaN(n)) n = 100;
                if (n <= 0) n = 100;
                n = Math.round(n);
                return n;
            })()
        }
    };
}

// ===== Collectors =====
function collectPlacedItemsFromSelection(selection) {
    var results = [];

    function walk(item) {
        if (!item) return;
        if (item.typename === 'PlacedItem') {
            results.push(item);
            return;
        }
        if (item.typename === 'GroupItem') {
            for (var i = 0; i < item.pageItems.length; i++) {
                walk(item.pageItems[i]);
            }
        }
    }

    for (var i = 0; i < selection.length; i++) {
        walk(selection[i]);
    }

    return results;
}

function collectAllPlacedItems(doc) {
    var results = [];
    var items = doc.placedItems;
    for (var i = 0; i < items.length; i++) {
        results.push(items[i]);
    }
    return results;
}

function collectRasterItemsFromSelection(selection) {
    var results = [];

    function isEmbeddedPlacedItem(it) {
        if (!it || it.typename !== 'PlacedItem') return false;
        try {
            return (it.file == null);
        } catch (e) {
            return true;
        }
    }

    function walk(item) {
        if (!item) return;

        if (item.typename === 'RasterItem') {
            results.push(item);
            return;
        }

        if (isEmbeddedPlacedItem(item)) {
            results.push(item);
            return;
        }

        if (item.typename === 'GroupItem') {
            for (var i = 0; i < item.pageItems.length; i++) {
                walk(item.pageItems[i]);
            }
        }
    }

    for (var i = 0; i < selection.length; i++) {
        walk(selection[i]);
    }

    return results;
}

function collectAllRasterItems(doc) {
    var results = [];

    // RasterItem（埋め込み画像）
    try {
        var r = doc.rasterItems;
        for (var i = 0; i < r.length; i++) {
            results.push(r[i]);
        }
    } catch (e) { }

    // 念のため：埋め込み後も PlacedItem のまま残るケース
    try {
        var p = doc.placedItems;
        for (var j = 0; j < p.length; j++) {
            try {
                if (p[j].file == null) {
                    results.push(p[j]);
                }
            } catch (e2) {
                // file 参照で例外の場合も embedded 扱い
                results.push(p[j]);
            }
        }
    } catch (e3) { }

    return results;
}

function selectionHasLinkedPlacedItem(selection) {
    function isLinkedPlacedItem(it) {
        if (!it || it.typename !== 'PlacedItem') return false;
        try {
            return (it.file != null && it.file.exists);
        } catch (e) {
            // file 参照で例外の場合は「リンク」とは断定しない
            return false;
        }
    }

    function walk(item) {
        if (!item) return false;
        if (isLinkedPlacedItem(item)) return true;
        if (item.typename === 'GroupItem') {
            for (var i = 0; i < item.pageItems.length; i++) {
                if (walk(item.pageItems[i])) return true;
            }
        }
        return false;
    }

    for (var i = 0; i < selection.length; i++) {
        if (walk(selection[i])) return true;
    }
    return false;
}

// ===== Reset Transform (placed/raster) =====
function collectPlacedOrRasterFromSelection(selection) {
    var results = [];
    function walk(item) {
        if (!item) return;
        var t = item.typename;
        if (t === 'PlacedItem' || t === 'RasterItem') {
            results.push(item);
            return;
        }
        if (t === 'GroupItem') {
            for (var i = 0; i < item.pageItems.length; i++) {
                walk(item.pageItems[i]);
            }
        }
    }
    for (var i = 0; i < selection.length; i++) {
        walk(selection[i]);
    }
    return results;
}

function applyResetTransformToItems(opts, items) {
    if (!items || items.length === 0) return;

    var previousInteractionLevel = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    try {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it || !it.typename) continue;
            if (it.typename !== 'PlacedItem' && it.typename !== 'RasterItem') continue;
            resetTransformOne(it, it.typename, opts || {});
        }
    } finally {
        app.userInteractionLevel = previousInteractionLevel;
        try { app.redraw(); } catch (e) { }
    }
}

function resetTransformOne(item, objectType, opts) {
    var sign = (objectType === 'RasterItem') ? -1 : 1;

    if (!opts) opts = {};
    if (!opts.rotate && !opts.skew && !opts.scale && !opts.ratio && !opts.flip && !opts.replaceReset) return;

    // Re-place reset (PlacedItem only)
    if (opts.replaceReset && objectType === 'PlacedItem') {
        replacePlacedItemToReset(item);
        return;
    }

    withBBoxResetAndRecenter(item, function () {
        if (opts.rotate && hasMatrix(item)) {
            cancelRotation(item, sign);
        }

        if (opts.flip && hasMatrix(item)) {
            try {
                var mNow = item.matrix;
                var fh = isFlippedHorizontal(mNow);
                var fv = isFlippedVertical(mNow);
                if (fh && fv) unflipBoth(item);
                else if (fh) unflipHorizontal(item);
                else if (fv) unflipVertical(item);
            } catch (eFlip) { }
        }

        if (opts.ratio) {
            equalizeScaleToMax(item);
        }

        if (opts.scale) {
            normalizeScaleOnly(item);
            applyUniformScalePercent(item, opts.scalePercent);
        }

        if (opts.skew) {
            removeSkewOnly(item);
            // retry
            try {
                if (hasMatrix(item)) {
                    var m = item.matrix;
                    var d = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
                    if (Math.abs(d.shear) > 1e-6) {
                        removeSkewOnly(item);
                    }
                }
            } catch (eSkew2) { }
        }
    });

    try {
        app.selection = null;
        app.selection = [item];
        app.executeMenuCommand('AI Reset Bounding Box');
    } catch (eBB) { }
}

function withBBoxResetAndRecenter(item, opFn) {
    var tl = item.position;
    var w1 = item.width;
    var h1 = item.height;
    if (typeof opFn === 'function') opFn();
    try {
        app.selection = null;
        app.selection = [item];
        app.executeMenuCommand('AI Reset Bounding Box');
    } catch (e) { }
    recenterToTopLeft(item, tl, w1, h1);
}

function recenterToTopLeft(item, tl, w1, h1) {
    var w2 = item.width;
    var h2 = item.height;
    item.position = [tl[0] + w1 / 2 - w2 / 2, tl[1] - h1 / 2 + h2 / 2];
}

function hasMatrix(obj) {
    try {
        return !!(obj && obj.matrix && typeof obj.matrix.mValueA !== 'undefined');
    } catch (e) { return false; }
}

function normalizeZero(n) { return (n === 0) ? 0 : n; }

function getRotationAngleDeg(a, b, sign) {
    var ang = Math.atan2(b, a) * 180 / Math.PI;
    ang = (sign < 0) ? -ang : ang;
    return normalizeZero(ang);
}

function getRotationMatrixSafe(deg) {
    try { if (app && typeof app.getRotationMatrix === 'function') return app.getRotationMatrix(deg); } catch (e) { }
    var rad = deg * Math.PI / 180.0;
    var cosv = Math.cos(rad), sinv = Math.sin(rad);
    var M = new Matrix();
    M.mValueA = cosv; M.mValueB = sinv; M.mValueC = -sinv; M.mValueD = cosv;
    M.mValueTX = 0; M.mValueTY = 0;
    return M;
}
function rotateBy(item, deg) { item.transform(getRotationMatrixSafe(deg)); }

function cancelRotation(item, sign) {
    if (!hasMatrix(item)) return 0;
    var a = item.matrix.mValueA, b = item.matrix.mValueB;
    var rot = getRotationAngleDeg(a, b, sign);
    rotateBy(item, rot);
    return rot;
}

function multiply2x2(a1, b1, c1, d1, a2, b2, c2, d2) { return { a: a1 * a2 + c1 * b2, b: b1 * a2 + d1 * b2, c: a1 * c2 + c1 * d2, d: b1 * c2 + d1 * d2 }; }
function invert2x2(a, b, c, d) {
    var det = a * d - b * c; if (Math.abs(det) < 1e-8) det = (det < 0 ? -1 : 1) * 1e-8;
    var invDet = 1.0 / det; return { a: d * invDet, b: -b * invDet, c: -c * invDet, d: a * invDet };
}
function toMatrix(obj) { var M = new Matrix(); M.mValueA = obj.a; M.mValueB = obj.b; M.mValueC = obj.c; M.mValueD = obj.d; M.mValueTX = 0; M.mValueTY = 0; return M; }

function decomposeQR(a, b, c, d) {
    var sx = Math.sqrt(a * a + b * b); if (sx === 0) sx = 1e-8;
    var q1x = a / sx, q1y = b / sx;
    var r12 = q1x * c + q1y * d;
    var u2x = c - r12 * q1x, u2y = d - r12 * q1y;
    var sy = Math.sqrt(u2x * u2x + u2y * u2y);
    if (sy === 0) { sy = 1e-8; u2x = -q1y; u2y = q1x; }
    var q2x = u2x / sy, q2y = u2y / sy;
    var shear = r12 / sx;
    return { sx: sx, sy: sy, shear: shear, q1x: q1x, q1y: q1y, q2x: q2x, q2y: q2y };
}
function buildFromQR(q1x, q1y, q2x, q2y, sx, sy, shear) {
    var r11 = sx, r12 = shear * sx, r21 = 0, r22 = sy;
    return { a: q1x * r11 + q2x * r21, b: q1y * r11 + q2y * r21, c: q1x * r12 + q2x * r22, d: q1y * r12 + q2y * r22 };
}
function applyDeltaToMatch(item, target2x2) {
    var m = item.matrix;
    var cur = { a: m.mValueA, b: m.mValueB, c: m.mValueC, d: m.mValueD };
    var inv = invert2x2(cur.a, cur.b, cur.c, cur.d);
    var delta = multiply2x2(inv.a, inv.b, inv.c, inv.d, target2x2.a, target2x2.b, target2x2.c, target2x2.d);
    item.transform(toMatrix(delta));
}

function removeSkewOnly(item) {
    if (!item || !hasMatrix(item)) return;
    var m = item.matrix;
    var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
    var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, dec.sx, dec.sy, 0);
    applyDeltaToMatch(item, target);
}
function normalizeScaleOnly(item) {
    if (!item || !hasMatrix(item)) return;
    var m = item.matrix;
    var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
    var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, 1, 1, dec.shear);
    applyDeltaToMatch(item, target);
}
function equalizeScaleToMax(item) {
    if (!item || !hasMatrix(item)) return;
    var m = item.matrix;
    var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
    var u = Math.max(dec.sx, dec.sy);
    var uPercent = Math.round(u * 100); u = uPercent / 100;
    var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, u, u, dec.shear);
    applyDeltaToMatch(item, target);
}
function applyUniformScalePercent(item, percent) {
    var p = Number(percent); if (!(p > 0)) return;
    try { item.resize(p, p, true, true, true, true, true); } catch (e) { }
}

function isFlippedHorizontal(mat) { return mat && mat.mValueA < 0; }
// Placed/Raster は上下判定が特殊
function isFlippedVertical(mat) { return mat && mat.mValueD > 0; }
function unflipHorizontal(item) { item.transform(app.getScaleMatrix(-100, 100), true, true, true, true, true, Transformation.CENTER); }
function unflipVertical(item) { item.transform(app.getScaleMatrix(100, -100), true, true, true, true, true, Transformation.CENTER); }
function unflipBoth(item) { item.transform(app.getScaleMatrix(-100, -100), true, true, true, true, true, Transformation.CENTER); }

function replacePlacedItemToReset(placedItem) {
    if (!placedItem || !placedItem.file) return null;
    try {
        var srcPos = placedItem.position;
        var srcW = placedItem.width;
        var srcH = placedItem.height;
        var srcCX = srcPos[0] + srcW / 2;
        var srcCY = srcPos[1] - srcH / 2;

        var doc = app.activeDocument;
        var newItem = doc.placedItems.add();
        newItem.file = placedItem.file;
        newItem.move(placedItem, ElementPlacement.PLACEBEFORE);

        try { normalizeScaleOnly(newItem); } catch (e) { }

        try {
            app.selection = null;
            app.selection = [newItem];
            app.executeMenuCommand('AI Reset Bounding Box');
        } catch (e2) { }

        try { if (placedItem.name) newItem.name = placedItem.name; } catch (e3) { }
        try { if (typeof placedItem.opacity !== 'undefined') newItem.opacity = placedItem.opacity; } catch (e4) { }
        try { if (typeof placedItem.blendingMode !== 'undefined') newItem.blendingMode = placedItem.blendingMode; } catch (e5) { }

        var newW = newItem.width;
        var newH = newItem.height;
        newItem.position = [srcCX - newW / 2, srcCY + newH / 2];

        try { placedItem.remove(); } catch (e6) { }
        return newItem;
    } catch (e7) { }
    return null;
}

// ===== Embed =====
function embedPlacedItems(placedItems) {
    var count = 0;

    for (var i = 0; i < placedItems.length; i++) {
        var item = placedItems[i];
        if (!item) continue;

        // 親がクリップグループかどうか
        var parent = item.parent;
        var isClipGroup = false;
        if (parent && parent.typename === 'GroupItem') {
            isClipGroup = parent.clipped;
        }

        // PSDファイルかどうか
        if (item.file && item.file.name.match(/\.psd$/i)) {
            try {
                item.embed();
                count++;
            } catch (e) {
                continue;
            }

            if (!isClipGroup) {
                try {
                    app.executeMenuCommand('ungroup');
                } catch (e2) { }
            }
        } else {
            try {
                item.embed();
                count++;
            } catch (e3) {
                continue;
            }
        }
    }

    return count;
}

function sanitizeFileBaseName(name) {
    if (!name) return '';
    // strip any path
    name = String(name).replace(/^.*[\\\/]/, '');
    // strip extension
    name = name.replace(/\.[^.]+$/, '');
    // replace invalid characters for mac/windows
    name = name.replace(/[\/:*?"<>|\r\n]+/g, '_');
    name = name.replace(/^\s+|\s+$/g, '');
    if (name.length > 120) name = name.substring(0, 120);
    return name;
}

function getFileNameFromXMPForEmbeddedItem(item) {
    try {
        var xmpStr = item.XMPString;
        if (!xmpStr) return '';

        // dc:title
        var titleMatch = xmpStr.match(/<dc:title>\s*<rdf:Alt>\s*<rdf:li[^>]*>(.*?)<\/rdf:li>/);
        if (titleMatch && titleMatch[1]) return titleMatch[1];

        // tiff:ImageDescription
        var descMatch = xmpStr.match(/tiff:ImageDescription>(.*?)<\/tiff:ImageDescription>/);
        if (descMatch && descMatch[1]) return descMatch[1];

        // photoshop:Source (sometimes contains a path)
        var srcMatch = xmpStr.match(/photoshop:Source="(.*?)"/);
        if (srcMatch && srcMatch[1]) return srcMatch[1];

        return '';
    } catch (e) {
        return '';
    }
}

function getPreferredBaseNameForUnembed(item, fallbackBaseName) {
    var name = '';

    // 1) item.name (レイヤーパネルの名称にファイル名が残っていることが多い)
    try {
        if (item.name && item.name !== '') name = item.name;
    } catch (e1) { }

    // 2) item.file.name（埋め込み後は null が多いが念のため）
    if (!name) {
        try {
            if (item.file && item.file.name) name = item.file.name;
        } catch (e2) { }
    }

    // 3) XMP
    if (!name) {
        name = getFileNameFromXMPForEmbeddedItem(item);
    }

    name = sanitizeFileBaseName(name);
    if (!name) name = fallbackBaseName;

    return name;
}

// ===== Release (Unembed) =====
function releaseRasterItemsToLinks(doc, rasterItems) {
    var previousInteractionLevel = app.userInteractionLevel;
    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    var exportErrors = [];
    var counter = 0;
    var linkedCount = 0;

    // export folder
    var exportFolder;
    try {
        exportFolder = Folder(doc.fullName.parent.fsName + '/Links/');
    } catch (e) {
        exportFolder = Folder(Folder.desktop.fsName + '/Links/');
    }
    if (!exportFolder.exists) {
        try { exportFolder.create(); } catch (e2) { }
    }

    try {
        // 後ろから処理（順序変化に強く）
        for (var i = rasterItems.length - 1; i >= 0; i--) {
            var oldImage = rasterItems[i];
            if (!oldImage) continue;

            var fallbackTitle = 'image' + (i + 1);
            var imageTitle = getPreferredBaseNameForUnembed(oldImage, fallbackTitle);

            // colorSpace
            var colorSpace;
            try {
                // RasterItem
                colorSpace = oldImage.imageColorSpace;
            } catch (eCS) {
                // embedded PlacedItem など
                try {
                    colorSpace = (doc.documentColorSpace === DocumentColorSpace.CMYK) ? ImageColorSpace.CMYK : ImageColorSpace.RGB;
                } catch (eCS2) {
                    colorSpace = ImageColorSpace.RGB;
                }
            }

            // script can't handle other colorSpaces
            if (
                colorSpace != ImageColorSpace.CMYK &&
                colorSpace != ImageColorSpace.RGB &&
                colorSpace != ImageColorSpace.GrayScale
            ) {
                exportErrors.push(imageTitle + ' has unsupported color space. (' + colorSpace + ')');
                continue;
            }

            var position;
            try { position = oldImage.position; } catch (ePos) { position = [0, 0]; }

            // get current scale and rotation
            var sr;
            try {
                sr = getLinkScaleAndRotation(oldImage);
            } catch (eSR) {
                sr = [100, 100, 0];
            }
            var scale = [sr[0], sr[1]];
            var rotation = sr[2];

            // move to new document for exporting
            var temp;
            try {
                temp = newDocument(imageTitle, colorSpace, 1000, 1000);
            } catch (eDoc) {
                exportErrors.push('"' + imageTitle + '" failed to create temp doc. (' + eDoc.message + ')');
                continue;
            }

            // duplicate to new document
            var workingImage;
            try {
                workingImage = oldImage.duplicate(temp.layers[0], ElementPlacement.PLACEATBEGINNING);
            } catch (eDup) {
                exportErrors.push('"' + imageTitle + '" failed to duplicate. (' + eDup.message + ')');
                try { temp.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose1) { }
                continue;
            }

            // set image to 100% scale 0° rotation and position
            try {
                var tm = app.getRotationMatrix(-rotation);
                tm = app.concatenateScaleMatrix(tm, 100 / scale[0] * 100, 100 / scale[1] * 100);
                workingImage.transform(tm, true, true, true, true, true);
                workingImage.position = [0, workingImage.height];
                temp.artboards[0].artboardRect = [0, workingImage.height, workingImage.width, 0];
            } catch (eTx) { }

            // export
            var file;
            try {
                var path = getFilePathWithOverwriteProtectionSuffix(exportFolder.fsName + '/' + imageTitle + '.psd');
                file = exportAsPSD(temp, path, colorSpace, 72);
            } catch (error) {
                exportErrors.push('"' + imageTitle + '" failed to export. (' + error.message + ')');
            }

            // close temp doc
            try { temp.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose2) { }

            if (!file || !file.exists) {
                continue;
            }

            // place link in active layer then move to oldImage position in stack
            var newImage;
            try {
                var targetLayer = doc.activeLayer || oldImage.layer;
                newImage = targetLayer.placedItems.add();

                // まず file を設定（基本）
                newImage.file = file;

                // 環境によっては file 代入だけでは反映されないことがあるため、relink/update を試行
                try {
                    if (newImage.relink) newImage.relink(file);
                } catch (eRelink) { }
                try {
                    if (newImage.update) newImage.update();
                } catch (eUpdate) { }

                // リンクできているか簡易検証
                try {
                    if (newImage.file != null && newImage.file.exists) {
                        linkedCount++;
                    }
                } catch (eCheck) { }

            } catch (eLink) {
                exportErrors.push('"' + imageTitle + '" failed to place link. (' + eLink.message + ')');
                try { if (file && file.exists) file.remove(); } catch (eRm) { }
                continue;
            }

            // scale, rotate, position to match original
            try {
                var tm2 = app.getScaleMatrix(scale[0], scale[1]);
                tm2 = app.concatenateRotationMatrix(tm2, rotation);
                newImage.transform(tm2, true, true, true, true, true);
            } catch (eFit1) { }

            // stacking + position
            try {
                newImage.move(oldImage, ElementPlacement.PLACEAFTER);
            } catch (eMove) { }
            try {
                newImage.position = position;
            } catch (eFit2) { }

            // remove old embedded image
            try {
                oldImage.remove();
            } catch (eDel) {
                exportErrors.push('Warning: Could not remove embedded image for "' + imageTitle + '".');
            }

            counter++;
        }
    } finally {
        app.userInteractionLevel = previousInteractionLevel;
        try { app.redraw(); } catch (eR) { }
    }

    return {
        count: counter,
        linkedCount: linkedCount,
        errors: exportErrors
    };
}

// ===== Helpers (from reference, simplified) =====
function exportAsPSD(doc, path, colorSpace, resolution) {
    var file = File(path);
    var options = new ExportOptionsPhotoshop();

    options.antiAliasing = false;
    options.artBoardClipping = true;
    options.imageColorSpace = colorSpace;
    options.editableText = false;
    options.flatten = true;
    options.maximumEditability = false;
    options.resolution = (resolution || 72);
    options.warnings = false;
    options.writeLayers = false;

    doc.exportFile(file, ExportType.PHOTOSHOP, options);
    return file;
}

function getLinkScaleAndRotation(item) {
    if (item == undefined) return;

    var m = item.matrix;
    var rotatedAmount;
    var unrotatedMatrix;
    var scaledAmount;

    var flipPlacedItem = (item.typename == 'PlacedItem') ? 1 : -1;

    try {
        rotatedAmount = item.tags.getByName('BBAccumRotation').value * 180 / Math.PI;
    } catch (error) {
        rotatedAmount = 0;
    }

    unrotatedMatrix = app.concatenateRotationMatrix(m, rotatedAmount * flipPlacedItem);
    scaledAmount = [unrotatedMatrix.mValueA * 100, unrotatedMatrix.mValueD * -100 * flipPlacedItem];

    return [scaledAmount[0], scaledAmount[1], rotatedAmount];
}

function newDocument(name, colorSpace, width, height) {
    var myDocPreset = new DocumentPreset();
    var myDocPresetType;

    myDocPreset.title = name;
    myDocPreset.width = width || 1000;
    myDocPreset.height = height || 1000;

    if (
        colorSpace == ImageColorSpace.CMYK ||
        colorSpace == ImageColorSpace.GrayScale ||
        colorSpace == ImageColorSpace.DeviceN
    ) {
        myDocPresetType = DocumentPresetType.BasicCMYK;
        myDocPreset.colorMode = DocumentColorSpace.CMYK;
    } else {
        myDocPresetType = DocumentPresetType.BasicRGB;
        myDocPreset.colorMode = DocumentColorSpace.RGB;
    }

    return app.documents.addDocument(myDocPresetType, myDocPreset);
}

function getFilePathWithOverwriteProtectionSuffix(path) {
    var index = 1;
    var parts = path.split(/(\.[^\.]+)$/);

    while (File(path).exists) {
        path = parts[0] + '(' + (++index) + ')' + parts[1];
    }

    return path;
}

// ===== Main =====
function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;

    var dialogResult = showDialog();
    if (!dialogResult) return;

    // リセット
    if (dialogResult.mode === 'reset') {
        if (!doc.selection || doc.selection.length === 0) {
            doc.selection = null;
            return;
        }
        var targets = collectPlacedOrRasterFromSelection(doc.selection);
        if (targets.length === 0) {
            doc.selection = null;
            return;
        }
        applyResetTransformToItems(dialogResult.resetOptions || {}, targets);
        doc.selection = null;
        return;
    }

    // モード分岐
    if (dialogResult.mode === 'embed') {
        var placedItems = [];

        if (dialogResult.embedMode === 'selection') {
            if (!doc.selection || doc.selection.length === 0) {
                doc.selection = null;
                return;
            }
            placedItems = collectPlacedItemsFromSelection(doc.selection);
        } else {
            placedItems = collectAllPlacedItems(doc);
        }

        if (placedItems.length === 0) {
            doc.selection = null;
            return;
        }

        var embeddedCount = embedPlacedItems(placedItems);
        if (embeddedCount > 0) {
            alert(embeddedCount + '個の画像を埋め込みました');
        }

        doc.selection = null;
        return;
    }

    // 解除（埋め込み解除）
    var rasterItems = [];

    if (dialogResult.releaseMode === 'selection') {
        if (!doc.selection || doc.selection.length === 0) {
            doc.selection = null;
            return;
        }
        // 選択に「リンク」画像（PlacedItem）が含まれている場合は解除できないため通知
        if (selectionHasLinkedPlacedItem(doc.selection)) {
            alert('リンク画像は解除できません。\n埋め込み画像を選択して実行してください。');
            doc.selection = null;
            return;
        }
        rasterItems = collectRasterItemsFromSelection(doc.selection);
    } else {
        rasterItems = collectAllRasterItems(doc);
    }

    if (rasterItems.length === 0) {
        doc.selection = null;
        return;
    }

    var result = releaseRasterItemsToLinks(doc, rasterItems);
    if (result) {
        if (result.count > 0) {
            var msg = result.count + '個の画像を解除しました';
            if (typeof result.linkedCount === 'number') {
                msg += '\n（リンク作成: ' + result.linkedCount + '）';
                if (result.linkedCount === 0) {
                    msg += '\n\n※リンク作成が確認できませんでした。書き出し先フォルダや権限、ドキュメントの保存状態をご確認ください。';
                }
            }
            if (result.errors && result.errors.length > 0) {
                msg += '\n\n' + result.errors.join('\n');
            }
            alert(msg);
        } else if (result.errors && result.errors.length > 0) {
            alert('解除できませんでした。\n\n' + result.errors.join('\n'));
        }
    }

    doc.selection = null;
}

main();