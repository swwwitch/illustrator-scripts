#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


// ===== Script info =====
var SCRIPT_NAME = 'ImageLinkManager';
var SCRIPT_VERSION = '1.1.0';


// 【概要 / Summary】
// 配置画像（PlacedItem）を中心に、画像の「埋め込み」「解除（埋め込み解除）」「リセット」「ケイ線」をまとめて扱うスクリプトです。
// ダイアログ上部のモードで処理を切り替え、対応パネルのみ有効化（それ以外はディム表示）します。
//
// モード切替ショートカット：
// ・埋め込み：E / 解除：U / リセット：R / ケイ線：S
//
// 【埋め込み】
// PlacedItem を embed() します。PSD（.psd）の場合は embed() 後に ungroup を試行します。
// ただし親グループがクリッピンググループ（clipped === true）の場合は ungroup を行いません。
//
// 【解除】
// 埋め込み画像（RasterItem、または file を参照できない PlacedItem）を PSD に書き出し、リンク配置（PlacedItem）へ置き換えます。
// 可能な範囲で位置・回転・拡大率を合わせます（効果が適用されている場合などは差異が出ることがあります）。
//
// 【リセット】
// 選択した PlacedItem / RasterItem に対して、回転・シアー・縦横比・反転・スケール（任意）を正規化します。
// 「再配置」を ON にすると、リンクを置き直すことで初期状態に近い形へ戻します（PlacedItem のみ）。
// スケール（%）は ↑↓ で±1、Shift+↑↓ で±10（10の倍数にスナップ）で増減できます。
//
// 【ケイ線】
// 選択オブジェクトに「新規線」を追加します。
// ・ケイ線のみ：新規線 → アウトライン
// ・クリップグループ：必要に応じてクリップ化し、オプションで角丸（LiveEffect）を適用してから
//   新規線 → パスファインダー（除外）を実行します。
// 角丸の半径は ↑↓ で±1、Shift+↑↓ で±10（10の倍数にスナップ）で増減できます。
//
// 実行後、選択を解除します。
// 更新日 / Updated: 2025-12-21

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
    var rbModeKeisen = modeGroup.add('radiobutton', undefined, 'ケイ線');
    rbModeEmbed.value = true; // デフォルト

    // ===== Keyboard shortcuts for mode (E/U/R/S) =====
    function addModeKeyHandler(dialog, rbEmbed, rbRelease, rbReset, rbKeisen, onChanged) {
        dialog.addEventListener('keydown', function (event) {
            // Avoid stealing keys while typing in an edittext
            try {
                if (dialog.active && dialog.active.type === 'edittext') return;
            } catch (e) { }

            if (event.keyName === 'E') {
                rbEmbed.value = true;
                rbRelease.value = false;
                rbReset.value = false;
                rbKeisen.value = false;
                if (typeof onChanged === 'function') onChanged();
                event.preventDefault();
            } else if (event.keyName === 'U') {
                rbEmbed.value = false;
                rbRelease.value = true;
                rbReset.value = false;
                rbKeisen.value = false;
                if (typeof onChanged === 'function') onChanged();
                event.preventDefault();
            } else if (event.keyName === 'R') {
                rbEmbed.value = false;
                rbRelease.value = false;
                rbReset.value = true;
                rbKeisen.value = false;
                if (typeof onChanged === 'function') onChanged();
                event.preventDefault();
            } else if (event.keyName === 'S') {
                rbEmbed.value = false;
                rbRelease.value = false;
                rbReset.value = false;
                rbKeisen.value = true;
                if (typeof onChanged === 'function') onChanged();
                event.preventDefault();
            }
        });
    }

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

    // ===== ケイ線 / Rules =====
    var panelKeisen = leftCol.add('panel', undefined, 'ケイ線');
    panelKeisen.orientation = 'column';
    panelKeisen.alignChildren = ['left', 'top'];
    panelKeisen.margins = [15, 20, 15, 10];

    var rbKeisenStrokeOnly = panelKeisen.add('radiobutton', undefined, 'ケイ線のみを追加');
    var rbKeisenClipGroup  = panelKeisen.add('radiobutton', undefined, 'クリップグループ');

    // 選択オブジェクト全体の外接矩形から角丸のデフォルト値を算出
    function calcDefaultRoundRadiusFromSelection(sel) {
        try {
            if (!sel || sel.length === 0) return 10;

            var top = -Infinity, left = Infinity, bottom = Infinity, right = -Infinity;
            for (var i = 0, n = sel.length; i < n; i++) {
                var b = sel[i].geometricBounds; // [top, left, bottom, right]
                if (!b || b.length !== 4) continue;
                if (b[0] > top) top = b[0];
                if (b[1] < left) left = b[1];
                if (b[2] < bottom) bottom = b[2];
                if (b[3] > right) right = b[3];
            }

            if (!isFinite(top) || !isFinite(left) || !isFinite(bottom) || !isFinite(right)) return 10;

            var height = Math.abs(top - bottom);
            var width  = Math.abs(right - left);
            var A = height + width;
            var B = A / 2;
            var r = Math.max(0, B / 20);
            return r;
        } catch (e) {
            return 10;
        }
    }

    var defaultRoundRadius = 10;
    try {
        if (app.documents.length > 0) {
            defaultRoundRadius = calcDefaultRoundRadiusFromSelection(app.activeDocument.selection);
        }
    } catch (eSel) {
        defaultRoundRadius = 10;
    }
    defaultRoundRadius = Math.round(defaultRoundRadius * 100) / 100;

    // 角丸（クリップグループ用）
    var roundRow = panelKeisen.add('group');
    roundRow.orientation = 'row';
    roundRow.alignChildren = ['left', 'center'];
    roundRow.margins = [20, 0, 0, 0]; // インデント

    var cbKeisenRoundCorners = roundRow.add('checkbox', undefined, '角丸');
    cbKeisenRoundCorners.value = true;

    var etKeisenRoundRadius = roundRow.add('edittext', undefined, String(defaultRoundRadius));
    etKeisenRoundRadius.characters = 6;
    var stKeisenRoundUnit = roundRow.add('statictext', undefined, 'pt');

    // 初期状態
    cbKeisenRoundCorners.enabled = false;
    etKeisenRoundRadius.enabled = false;
    stKeisenRoundUnit.enabled = false;

    rbKeisenStrokeOnly.value = true;

    function updateKeisenRoundUI() {
        var clipEnabled = (rbKeisenClipGroup.value === true);
        cbKeisenRoundCorners.enabled = clipEnabled;

        var radiusEnabled = clipEnabled && (cbKeisenRoundCorners.value === true);
        etKeisenRoundRadius.enabled = radiusEnabled;
        stKeisenRoundUnit.enabled = radiusEnabled;
    }

    rbKeisenStrokeOnly.onClick = updateKeisenRoundUI;
    rbKeisenClipGroup.onClick  = updateKeisenRoundUI;
    cbKeisenRoundCorners.onClick = updateKeisenRoundUI;
    updateKeisenRoundUI();

    etKeisenRoundRadius.onChange = function () {
        var v = Number(etKeisenRoundRadius.text);
        if (isNaN(v)) return;
        etKeisenRoundRadius.text = String(v);
    };

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

    // Arrow key increment helper (±1, shift±10 with snapping)
    function changeValueByArrowKey(editText) {
        editText.addEventListener('keydown', function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == 'Up') {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == 'Down') {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
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
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
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
    // ケイ線：角丸の値も↑↓で変更
    changeValueByArrowKey(etKeisenRoundRadius);

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
        panelKeisen.enabled = rbModeKeisen.value;
    }
    rbModeEmbed.onClick = updatePanelState;
    rbModeRelease.onClick = updatePanelState;
    rbModeReset.onClick = updatePanelState;
    rbModeKeisen.onClick = updatePanelState;
    updatePanelState();

    // Keyboard shortcuts: 埋め込み(E) / 解除(U) / リセット(R) / ケイ線(S)
    addModeKeyHandler(dlg, rbModeEmbed, rbModeRelease, rbModeReset, rbModeKeisen, updatePanelState);

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
        mode: rbModeEmbed.value ? 'embed' : (rbModeRelease.value ? 'release' : (rbModeReset.value ? 'reset' : 'keisen')),
        embedMode: rbEmbedAll.value ? 'all' : 'selection',
        releaseMode: rbReleaseAll.value ? 'all' : 'selection',
        keisenOptions: {
            mode: rbKeisenStrokeOnly.value ? 'strokeOnly' : 'clipGroup',
            roundCorners: (cbKeisenRoundCorners.value === true),
            roundRadius: (function () {
                var n = parseFloat(etKeisenRoundRadius.text, 10);
                if (isNaN(n)) n = 0;
                if (n < 0) n = 0;
                return n;
            })()
        },
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

// ===== Keisen (Rules) =====
function applyKeisenToSelection(doc, opts) {
    if (!doc || !doc.selection || doc.selection.length === 0) return;

    var mode = (opts && opts.mode) ? opts.mode : 'strokeOnly';
    var doClipGroup = (mode === 'clipGroup');
    var doRound = (opts && opts.roundCorners === true);
    var roundRadius = (opts && typeof opts.roundRadius !== 'undefined') ? Number(opts.roundRadius) : 0;

    if (doRound) {
        if (isNaN(roundRadius) || roundRadius < 0) {
            throw new Error('角丸の値が不正です。0以上の数値を入力してください。');
        }
    }

    // クリップグループ化（必要な場合）
    if (doClipGroup) {
        var g = makeClippingFromSelection(doc);
        if (!g) {
            alert('クリップグループ化できる選択ではありません。\n（例：画像1つ、または 画像1つ+パス1つ、または 複数オブジェクト+パス）');
            return;
        }
    }

    // ケイ線のみ
    if (!doClipGroup) {
        app.executeMenuCommand('Adobe New Stroke Shortcut');
        app.executeMenuCommand('Live Outline Object');
        return;
    }

    // クリップグループ後：各グループに適用
    var targets = doc.selection;
    if (targets && targets.length) {
        for (var i = 0; i < targets.length; i++) {
            var gi = targets[i];
            if (!gi || gi.typename !== 'GroupItem' || !gi.clipped) continue;

            if (doRound) {
                applyRoundCornersLiveEffect(gi, roundRadius);
            }

            // メニューコマンドは単体選択で実行
            doc.selection = [gi];
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            app.executeMenuCommand('Live Pathfinder Exclude');
        }
        // 選択を戻す
        doc.selection = targets;
    }
}

function isImageItem(item) {
    return item && (item.typename === 'PlacedItem' || item.typename === 'RasterItem');
}

function getFrontmostPath(pathArray) {
    var topPath = null;
    var highestZ = -1;
    for (var i = 0; i < pathArray.length; i++) {
        var item = pathArray[i];
        if (item && item.typename === 'PathItem' && item.zOrderPosition > highestZ) {
            highestZ = item.zOrderPosition;
            topPath = item;
        }
    }
    return topPath;
}

// 画像1つ → 画像外接の矩形でクリップ
function createClippingMaskGroup(imageItem) {
    var targetLayer = imageItem.layer;
    var wasLocked = targetLayer.locked;
    var wasVisible = targetLayer.visible;
    var wasTemplate = targetLayer.isTemplate;

    if (wasLocked) targetLayer.locked = false;
    if (!wasVisible) targetLayer.visible = true;
    if (wasTemplate) targetLayer.isTemplate = false;

    var rect = targetLayer.pathItems.rectangle(
        imageItem.top,
        imageItem.left,
        imageItem.width,
        imageItem.height
    );
    rect.stroked = false;
    rect.filled = false;

    var groupItem = targetLayer.groupItems.add();
    imageItem.moveToBeginning(groupItem);
    rect.moveToBeginning(groupItem);
    groupItem.clipped = true;

    if (wasLocked) targetLayer.locked = true;
    if (!wasVisible) targetLayer.visible = false;
    if (wasTemplate) targetLayer.isTemplate = true;

    return groupItem;
}

// 画像1つ + パス1つ → パスでクリップ
function createMaskWithPath(imageItem, pathItem) {
    var targetLayer = imageItem.layer;
    if (pathItem.layer != targetLayer) {
        pathItem.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
    }

    var groupItem = targetLayer.groupItems.add();
    imageItem.moveToBeginning(groupItem);
    pathItem.moveToBeginning(groupItem);
    groupItem.clipped = true;

    return groupItem;
}

// 現在の選択から「クリップグループ」を作成し、選択を更新
// 戻り値: 作成した groupItem（複数作成時は先頭を返す） / 作成できない場合は null
function makeClippingFromSelection(doc) {
    var sel = doc.selection;
    if (!sel || sel.length === 0) return null;

    // すでにクリップグループが1つ選択されているなら、そのまま
    if (sel.length === 1 && sel[0].typename === 'GroupItem' && sel[0].clipped) {
        return sel[0];
    }

    var images = [];
    var paths = [];
    var i;

    for (i = 0; i < sel.length; i++) {
        if (!sel[i]) continue;
        if (isImageItem(sel[i])) {
            images.push(sel[i]);
        } else if (sel[i].typename === 'PathItem') {
            paths.push(sel[i]);
        }
    }

    // 画像1 + パス1（選択がちょうど2つ）
    if (sel.length === 2 && images.length === 1 && paths.length === 1) {
        var g1 = createMaskWithPath(images[0], paths[0]);
        doc.selection = [g1];
        return g1;
    }

    // 画像1のみ
    if (sel.length === 1 && images.length === 1) {
        var g2 = createClippingMaskGroup(images[0]);
        doc.selection = [g2];
        return g2;
    }

    // パスが含まれている場合は、最前面パスをマスクとして全体をクリップ
    if (paths.length > 0) {
        var maskPath = getFrontmostPath(paths);
        if (!maskPath) return null;

        var targetLayer = maskPath.layer;
        var wasLocked = targetLayer.locked;
        var wasVisible = targetLayer.visible;
        var wasTemplate = targetLayer.isTemplate;

        if (wasLocked) targetLayer.locked = false;
        if (!wasVisible) targetLayer.visible = true;
        if (wasTemplate) targetLayer.isTemplate = false;

        var groupItem = targetLayer.groupItems.add();

        // マスク以外を先に移動（末尾から）
        for (i = sel.length - 1; i >= 0; i--) {
            if (!sel[i]) continue;
            if (sel[i] === maskPath) continue;
            sel[i].moveToBeginning(groupItem);
        }
        maskPath.moveToBeginning(groupItem);
        groupItem.clipped = true;

        if (wasLocked) targetLayer.locked = true;
        if (!wasVisible) targetLayer.visible = false;
        if (wasTemplate) targetLayer.isTemplate = true;

        doc.selection = [groupItem];
        return groupItem;
    }

    // 複数画像のみ → それぞれ外接矩形で個別にクリップ
    if (images.length > 0) {
        var groups = [];
        for (i = 0; i < images.length; i++) {
            try {
                groups.push(createClippingMaskGroup(images[i]));
            } catch (e) { }
        }
        if (groups.length > 0) {
            doc.selection = groups;
            return groups[0];
        }
    }

    return null;
}

// 角丸 LiveEffect XML を生成
function createRoundCornersEffectXML(radius) {
    var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
    return xml.replace('#value#', radius);
}

// 角丸（LiveEffect: Adobe Round Corners）を適用
function applyRoundCornersLiveEffect(targetItem, radius) {
    if (!targetItem) return false;
    var r = Number(radius);
    if (isNaN(r) || r < 0) return false;

    var xml = createRoundCornersEffectXML(r);

    try {
        targetItem.applyEffect(xml);
        return true;
    } catch (e) {
        return false;
    }
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

    // ケイ線
    if (dialogResult.mode === 'keisen') {
        if (!doc.selection || doc.selection.length === 0) {
            alert('オブジェクトを選択してください。');
            doc.selection = null;
            return;
        }

        try {
            applyKeisenToSelection(doc, dialogResult.keisenOptions || {});
        } catch (eKeisen) {
            alert('ケイ線の処理中にエラーが発生しました。\n' + eKeisen);
        }

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