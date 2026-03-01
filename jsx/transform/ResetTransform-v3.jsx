#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.10";

var CONFIG = {
    rectSnapMin: 0.5, // degrees
    rectSnapMax: 10, // degrees
    eps: 1e-8 // numerical epsilon for matrix ops
};

/*
Language helper / 言語判定ヘルパー
- 日本語環境なら "ja"、それ以外は "en"
*/

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* UI labels / UIラベル（ローカライズ） */
var LABELS = {
    dialogTitle: {
        ja: "リセット（回転・比率） " + SCRIPT_VERSION,
        en: "Reset (Rotate / Scale) " + SCRIPT_VERSION
    },
    panels: {
        placed: {
            ja: "リセット",
            en: "Reset"
        }
    },
    checks: {
        rotate: {
            ja: "回転",
            en: "Rotate"
        },
        skew: {
            ja: "シアー",
            en: "Shear"
        },
        ratio: {
            ja: "縦横比",
            en: "Aspect Ratio"
        },
        flip: {
            ja: "反転",
            en: "Flip"
        },
        replaceReset: {
            ja: "再配置",
            en: "Replacing"
        },
        scale100: {
            ja: "スケール",
            en: "Scale"
        }
    },
    buttons: {
        ok: {
            ja: "リセット",
            en: "Reset"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        }
    },
    alerts: {
        select: {
            ja: "オブジェクトを選択してください。",
            en: "Please select an object."
        },
        none: {
            ja: "該当するリンク/ラスター画像が選択されていません。",
            en: "No placed/raster images were selected."
        }
    }
};

/* ガイドの設定 / Set guide properties */
/* Dialog helpers / ダイアログ用ヘルパー（位置・透明度） */
var DIALOG_OFFSET_X = 300; // shift in pixels (right = positive)
var DIALOG_OFFSET_Y = 0; // shift in pixels (down = positive)
var DIALOG_OPACITY = 0.98; // 0.0 – 1.0

function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function() {
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        } catch (e) {}
    };
}


function setDialogOpacity(dlg, opacityValue) {
    try {
        dlg.opacity = opacityValue;
    } catch (e) {}
}

/* Analyze selection → capabilities / 選択状態から対応可否を判定 */
function getSelectionCapabilities(sel) {
    var caps = {
        placedOrRaster: false
    };
    if (!sel || !sel.length) return caps;
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (!it || !it.typename) continue;
        var t = it.typename;
        if (t === 'PlacedItem' || t === 'RasterItem') caps.placedOrRaster = true;
    }
    return caps;
}

function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.round((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
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

/* Hotkey → toggle checkbox / ホットキーでチェックボックスON/OFF */
function addHotkeyToggle(dialog, keyChar, checkbox, onToggle) {
    dialog.addEventListener('keydown', function(event) {
        var k = (event.keyName || '').toUpperCase();
        if (k === String(keyChar).toUpperCase()) {
            if (!checkbox.enabled) {
                event.preventDefault();
                return;
            }
            checkbox.value = !checkbox.value;
            if (typeof onToggle === 'function') onToggle(checkbox.value);
            event.preventDefault();
        }
    });
}

function showOptionsDialog() {
    var prefs = {
        rotate: true,
        skew: true,
        scale: false
    };

    var selNow = (app && app.activeDocument) ? app.activeDocument.selection : null;
    var caps = getSelectionCapabilities(selNow);

    var dlg = new Window('dialog', LABELS.dialogTitle[lang]);
    dlg.alignChildren = 'left';
    setDialogOpacity(dlg, DIALOG_OPACITY);
    shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);



    /* Main group: two-column layout / 2カラムレイアウト */
    var gMain = dlg.add('group');
    gMain.orientation = 'row';
    gMain.alignChildren = 'top';
    gMain.alignment = 'fill';

    /* Left column / 左カラム */
    var colLeft = gMain.add('group');
    colLeft.orientation = 'column';
    colLeft.alignChildren = 'left';
    colLeft.alignment = 'fill';

    /* Right column / 右カラム */
    var colRight = gMain.add('group');
    colRight.orientation = 'column';
    colRight.alignChildren = 'left';
    colRight.alignment = 'fill';

    /* Panel: Placed Images / 配置画像 */
    var pnlReset = colLeft.add('panel', undefined, LABELS.panels.placed[lang]);
    pnlReset.alignment = 'fill';
    pnlReset.orientation = 'column';
    pnlReset.alignChildren = 'left';
    pnlReset.margins = [15, 20, 15, 10];

    // Option: Reset by re-placing / ［再配置でリセット］（配置画像のみ）
    var cbReplaceReset = pnlReset.add('checkbox', undefined, LABELS.checks.replaceReset[lang]);
    cbReplaceReset.value = false; // default OFF

    var cbRotate = pnlReset.add('checkbox', undefined, LABELS.checks.rotate[lang]);
    var cbSkew = pnlReset.add('checkbox', undefined, LABELS.checks.skew[lang]);
    var cbRatio = pnlReset.add('checkbox', undefined, LABELS.checks.ratio[lang]);
    var cbFlip = pnlReset.add('checkbox', undefined, LABELS.checks.flip[lang]);
    var cbScale = pnlReset.add('checkbox', undefined, LABELS.checks.scale100[lang]);
    cbRotate.value = (prefs.rotate !== false);
    cbSkew.value = (prefs.skew !== false);
    cbRatio.value = true;
    cbFlip.value = true;
    cbScale.value = (prefs.scale !== false);
    // Add scale percent field + % label after cbScale
    var gScale = pnlReset.add('group');
    gScale.orientation = 'row';
    gScale.alignChildren = 'center';

    var etScale = gScale.add('edittext', undefined, '100');
    etScale.characters = 5; // default width

    var stPercent = gScale.add('statictext', undefined, '%');

    // Enable/disable with the checkbox
    etScale.enabled = cbScale.value;
    stPercent.enabled = cbScale.value;
    cbScale.onClick = function() {
        var on = cbScale.value;
        etScale.enabled = on;
        stPercent.enabled = on;
        if (on) {
            try {
                etScale.active = true; // focus
                // select all text (best-effort across ScriptUI versions)
                if (typeof etScale.select === 'function') {
                    etScale.select();
                } else if (typeof etScale.textselection !== 'undefined') {
                    etScale.textselection = etScale.text; // some builds accept string
                }
            } catch (e) {}
        }
    };
    // Arrow-key increment/decrement support
    changeValueByArrowKey(etScale);

    // Helper: enable/disable entire Placed Images panel based on re-place reset and selection
    function updatePlacedPanelEnabled() {
        // If selection doesn't support placed/raster, keep everything disabled.
        var supported = !!caps.placedOrRaster;

        // Panel itself stays enabled when supported so cbReplaceReset remains clickable.
        pnlReset.enabled = supported;
        cbReplaceReset.enabled = supported;

        // When re-place reset is ON, disable other options in the panel
        var othersOn = supported && !cbReplaceReset.value;

        cbRotate.enabled = othersOn;
        cbSkew.enabled  = othersOn;
        cbRatio.enabled = othersOn;
        cbFlip.enabled  = othersOn;
        cbScale.enabled = othersOn;

        // Scale field follows cbScale when others are enabled
        etScale.enabled = othersOn && cbScale.value;
        stPercent.enabled = othersOn && cbScale.value;
    }

    cbReplaceReset.onClick = function() {
        updatePlacedPanelEnabled();
    };

    updatePlacedPanelEnabled();

    // Hotkey: 'S' toggles Scale / SキーでスケールON/OFF
    addHotkeyToggle(dlg, 'S', cbScale, function(enabled) {
        if (!pnlReset.enabled) return;
        etScale.enabled = enabled;
        stPercent.enabled = enabled;
        if (enabled) {
            try {
                etScale.active = true;
                if (typeof etScale.select === 'function') {
                    etScale.select();
                } else if (typeof etScale.textselection !== 'undefined') {
                    etScale.textselection = etScale.text;
                }
            } catch (e) {}
        }
    });
    // Hotkey: 'F' toggles Flip / Fキーで反転ON/OFF
    addHotkeyToggle(dlg, 'F', cbFlip);

    /* Dialog buttons / ダイアログボタン */
    var gBtns = dlg.add('group');
    gBtns.alignment = 'center';
    gBtns.add('button', undefined, LABELS.buttons.cancel[lang], {
        name: 'cancel'
    });
    gBtns.add('button', undefined, LABELS.buttons.ok[lang], {
        name: 'ok'
    });


    if (dlg.show() !== 1) return null; // cancelled

    var out = {
        rotate: cbRotate.value,
        skew: cbSkew.value,
        ratio: cbRatio.value,
        flip: cbFlip.value,
        replaceReset: cbReplaceReset.value,
        scale: cbScale.value,
        scalePercent: (function() {
            var n = parseFloat(etScale.text, 10);
            if (isNaN(n)) n = 100;
            if (n <= 0) n = 100;
            n = Math.round(n); // 整数％に統一（例：16.3 → 16）
            return n;
        })()
    };
    return out;
}

/* Dispatch table / ディスパッチテーブル：typename → handler */
function makeHandlers(opts) {
    function res(p, s, pl, ra) {
        return {
            processed: p | 0,
            skipped: s | 0,
            countPlaced: pl | 0,
            countRaster: ra | 0
        };
    }

    function handleImage(item, typename) {
        setScaleTo100(item, typename, opts);
        // Placed/Raster は常に処理成功としてカウント（従来仕様）
        if (typename === 'PlacedItem') return res(1, 0, 1, 0);
        if (typename === 'RasterItem') return res(1, 0, 0, 1);
        return res(1, 0, 0, 0);
    }

    return {
        PathItem: function(item) {
            return res(0, 1, 0, 0); // Path (Line) support removed
        },
        PlacedItem: function(item) {
            return handleImage(item, 'PlacedItem');
        },
        RasterItem: function(item) {
            return handleImage(item, 'RasterItem');
        }
    };
}

function main() {
    var sel = activeDocument.selection;
    if (!sel || sel.length === 0) {
        alert(LABELS.alerts.select[lang]);
        return;
    }

    var opts = showOptionsDialog();
    if (!opts) {
        return;
    }

    var processed = 0;
    var skipped = 0;
    var countPlaced = 0,
        countRaster = 0;
    var handlers = makeHandlers(opts);

    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (!item || !item.typename) {
            skipped++;
            continue;
        }
        var t = item.typename;
        var handler = handlers[t];
        if (handler) {
            var r = handler(item);
            processed += r.processed || 0;
            skipped += r.skipped || 0;
            countPlaced += r.countPlaced || 0;
            countRaster += r.countRaster || 0;
            continue;
        }
        // 未対応タイプはスキップ
        skipped++;
        continue;
    }

    if (processed === 0) {
        alert(LABELS.alerts.none[lang]);
    }
}

/* Helpers / ヘルパー群 */

function withBBoxResetAndRecenter(item, opFn) {
    /* Run transform → reset BBox → recenter to keep top-left / 変形→BBoxリセット→左上基準で再配置 */
    var tl = item.position;
    var w1 = item.width,
        h1 = item.height;
    if (typeof opFn === 'function') opFn();
    try {
        app.selection = null;
        app.selection = [item];
        app.executeMenuCommand("AI Reset Bounding Box");
    } catch (e) {}
    recenterToTopLeft(item, tl, w1, h1);
}


function resetPlacedOrRasterRotationAndShear(item, doRotate, doShear) {
    if (!hasMatrix(item)) return;
    if (!doRotate && !doShear) return;
    var sign = (item.typename === 'RasterItem') ? -1 : 1;
    withBBoxResetAndRecenter(item, function() {
        if (doRotate) cancelRotation(item, sign);
        if (doShear) removeSkewOnly(item);
    });
}

function resetPathRotationAndShear(item, doRotate, doShear) {
    if (!hasMatrix(item)) return;
    if (!doRotate && !doShear) return;
    withBBoxResetAndRecenter(item, function() {
        if (doRotate) cancelRotation(item, 1);
        if (doShear) removeSkewOnly(item);
    });
}

function isWindows() {
    return Folder.fs === "Windows";
}

function isEPS(placedItem) {
    if (!placedItem || !placedItem.file) return false;
    return isWindows() ?
        (String(placedItem.file.displayName).match(/[^.]+$/i) == 'eps') :
        (placedItem.file.type == 'EPSF');
}


function getRotationAngleDeg(a, b, sign) {
    var ang = Math.atan2(b, a) * 180 / Math.PI;
    return (sign < 0) ? -ang : ang;
}

/* Rotation matrix helper (uses Illustrator API if available) / 回転行列ヘルパー（可能ならIllustrator標準APIを使用） */
function getRotationMatrixSafe(deg) {
    // Prefer Illustrator's API for clarity
    try {
        if (app && typeof app.getRotationMatrix === 'function') {
            return app.getRotationMatrix(deg);
        }
    } catch (e) {}
    // Fallback: build a Matrix manually
    var rad = deg * Math.PI / 180.0;
    var cosv = Math.cos(rad),
        sinv = Math.sin(rad);
    var M = new Matrix();
    M.mValueA = cosv; // a
    M.mValueB = sinv; // b
    M.mValueC = -sinv; // c
    M.mValueD = cosv; // d
    M.mValueTX = 0;
    M.mValueTY = 0;
    return M;
}

/* Rotate item by degrees using Illustrator matrix API / Illustratorの回転行列APIを用いて回転 */
function rotateBy(item, deg) {
    var mat = getRotationMatrixSafe(deg);
    item.transform(mat);
}

// Rotate TextFrame by deg, prefer Illustrator's native rotate with full flags and center origin

function multiply2x2(a1, b1, c1, d1, a2, b2, c2, d2) {
    return {
        a: a1 * a2 + c1 * b2,
        b: b1 * a2 + d1 * b2,
        c: a1 * c2 + c1 * d2,
        d: b1 * c2 + d1 * d2
    };
}

function invert2x2(a, b, c, d) {
    var det = a * d - b * c;
    if (Math.abs(det) < CONFIG.eps) det = (det < 0 ? -1 : 1) * CONFIG.eps;
    var invDet = 1.0 / det;
    return {
        a: d * invDet,
        b: -b * invDet,
        c: -c * invDet,
        d: a * invDet
    };
}

function toMatrix(obj) {
    var M = new Matrix();
    M.mValueA = obj.a;
    M.mValueB = obj.b;
    M.mValueC = obj.c;
    M.mValueD = obj.d;
    M.mValueTX = 0;
    M.mValueTY = 0;
    return M;
}

function decomposeQR(a, b, c, d) {
    var sx = Math.sqrt(a * a + b * b);
    if (sx === 0) sx = CONFIG.eps;
    var q1x = a / sx,
        q1y = b / sx;
    var r12 = q1x * c + q1y * d;
    var u2x = c - r12 * q1x;
    var u2y = d - r12 * q1y;
    var sy = Math.sqrt(u2x * u2x + u2y * u2y);
    if (sy === 0) {
        sy = CONFIG.eps;
        u2x = -q1y;
        u2y = q1x;
    }
    var q2x = u2x / sy,
        q2y = u2y / sy;
    var shear = r12 / sx;
    return {
        sx: sx,
        sy: sy,
        shear: shear,
        q1x: q1x,
        q1y: q1y,
        q2x: q2x,
        q2y: q2y
    };
}

function buildFromQR(q1x, q1y, q2x, q2y, sx, sy, shear) {
    var r11 = sx,
        r12 = shear * sx,
        r21 = 0,
        r22 = sy;
    return {
        a: q1x * r11 + q2x * r21,
        b: q1y * r11 + q2y * r21,
        c: q1x * r12 + q2x * r22,
        d: q1y * r12 + q2y * r22
    };
}

function applyDeltaToMatch(item, target2x2) {
    var m = item.matrix;
    var cur = {
        a: m.mValueA,
        b: m.mValueB,
        c: m.mValueC,
        d: m.mValueD
    };
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

// Make horizontal/vertical scales equal by lifting the smaller to the larger (preserving orientation/shear)
function equalizeScaleToMax(item) {
    if (!item || !hasMatrix(item)) return;
    var m = item.matrix;
    var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
    var u = Math.max(dec.sx, dec.sy); // choose the larger of current scales

    // ▼追加：整数％にスナップ（例：16.321%→16%、16.3%→16%）
    var uPercent = Math.round(u * 100); // 四捨五入（例：16.321%→16%、16.5%→17%）
    u = uPercent / 100;

    var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, u, u, dec.shear);
    applyDeltaToMatch(item, target);
}

// Apply uniform scaling by percent (100 = 100%)
function applyUniformScalePercent(item, percent) {
    var p = Number(percent);
    if (!(p > 0)) return;
    try {
        // Illustrator's resize: percent values (100 = 100%)
        item.resize(p, p, true, true, true, true, true);
    } catch (e) {}
}

function hasMatrix(obj) {
    try {
        return !!(obj && obj.matrix && typeof obj.matrix.mValueA !== 'undefined');
    } catch (e) {
        return false;
    }
}

// --- Flip detection helpers (no rotation/shear assumed) / 反転検出ヘルパー（回転・シアーなし前提） ---
function isFlippedHorizontal(mat) {
    // 左右反転: mValueA が負
    return mat && mat.mValueA < 0;
}
// NOTE: Placed/Raster はデフォルトで mValueD が負になり得るため、上下反転は mValueD &gt; 0 を基準に判定
function isFlippedVertical(mat) {
    // 上下反転: mValueD が正（Placed/Raster の基準に合わせる）
    return mat && mat.mValueD > 0;
}

// --- Flip correction operations / 反転解除処理 ---
function unflipHorizontal(item) {
    if (!item) return;
    item.transform(
        app.getScaleMatrix(-100, 100), // 左右のみ反転を打ち消す
        true, // changePositions
        true, // changeFillPatterns
        true, // changeFillGradients
        true, // changeStrokePattern
        true, // changeLineWidths
        Transformation.CENTER
    );
}

function unflipVertical(item) {
    if (!item) return;
    item.transform(
        app.getScaleMatrix(100, -100), // 上下のみ反転を打ち消す
        true, // changePositions
        true, // changeFillPatterns
        true, // changeFillGradients
        true, // changeStrokePattern
        true, // changeLineWidths
        Transformation.CENTER
    );
}

function unflipBoth(item) {
    if (!item) return;
    item.transform(
        app.getScaleMatrix(-100, -100), // 左右・上下の両反転を同時に打ち消す
        true, // changePositions
        true, // changeFillPatterns
        true, // changeFillGradients
        true, // changeStrokePattern
        true, // changeLineWidths
        Transformation.CENTER
    );
}


function cancelRotation(item, sign) {
    if (!hasMatrix(item)) return 0; // safety: some items may not expose matrix
    var a = item.matrix.mValueA;
    var b = item.matrix.mValueB;
    var rot = getRotationAngleDeg(a, b, sign);
    rotateBy(item, rot);
    return rot;
}

// Sign-agnostic cancel: always rotate by the negative of the current angle to zero out rotation

/* Cancel skew & scale in one shot / スキューと拡大縮小を一括で打ち消す */
function cancelSkew(item, sign) {
    // 1) 一旦、対象アイテムの行列を純粋な単位行列に近づける
    var m = item.matrix;
    m.mValueA = 1;
    m.mValueB = 0;
    m.mValueC = 0;
    m.mValueD = 1;
    m.mValueTX = 0.0;
    m.mValueTY = 0.0;

    // 2) 現在行列の逆行列を適用（スケール/スキューの打ち消し）
    var inv1 = invertMatrix(item.matrix);
    item.transform(inv1);

    // 3) もう一度、現在行列から逆行列を作成し、Y成分の最小補正だけ適用
    var inv2 = invertMatrix(item.matrix);
    inv2.mValueD *= sign * (-1);
    item.transform(inv2);
}

/* Scale-only handler (internally same as cancelSkew) / スケール単独用（内部は一括反転） */
function cancelScale(item, sign) {
    var inv = invertMatrix(item.matrix);
    inv.mValueD *= sign * (-1);
    item.transform(inv);
}

function recenterToTopLeft(item, tl, w1, h1) {
    var w2 = item.width,
        h2 = item.height;
    item.position = [tl[0] + w1 / 2 - w2 / 2, tl[1] - h1 / 2 + h2 / 2];
}



// --- PlacedItem re-place helper ---
function replacePlacedItemToReset(placedItem) {
    if (!placedItem || !placedItem.file) return null;
    try {
        // Note: size will revert to native 100% (no width/height matching)
        var targetPos = placedItem.position;
        var doc = app.activeDocument;
        var newItem = doc.placedItems.add();
        newItem.file = placedItem.file;
        // Move to same stacking context
        newItem.move(placedItem, ElementPlacement.PLACEBEFORE);
        // Ensure the placed item is at native 100% scale (defensive)
        try { normalizeScaleOnly(newItem); } catch (e) {}
        // Copy name, opacity, blendingMode where possible
        try {
            if (placedItem.name) newItem.name = placedItem.name;
        } catch (e) {}
        try {
            if (typeof placedItem.opacity !== "undefined") newItem.opacity = placedItem.opacity;
        } catch (e) {}
        try {
            if (typeof placedItem.blendingMode !== "undefined") newItem.blendingMode = placedItem.blendingMode;
        } catch (e) {}
        // Set position to match
        newItem.position = targetPos;
        // Reset bounding box for the new item after re-placing
        try {
            app.selection = null;
            app.selection = [newItem];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (e) {}
        // Remove original
        try { placedItem.remove(); } catch (e) {}
        return newItem;
    } catch (e) {}
    return null;
}

function setScaleTo100(item, objectType, opts) {
    var sign = (objectType === "RasterItem") ? -1 : 1;
    if (!opts) opts = {};
    // Fast path: do nothing if no operations are enabled
    if (!opts.rotate && !opts.skew && !opts.scale && !opts.ratio && !opts.flip && !opts.replaceReset) return;

    // Re-place: if enabled and PlacedItem, do it and return
    if (opts.replaceReset && objectType === "PlacedItem") {
        replacePlacedItemToReset(item);
        return;
    }

    withBBoxResetAndRecenter(item, function() {
        // 1) Rotate first (stabilize orientation)
        if (opts.rotate && hasMatrix(item)) {
            cancelRotation(item, sign);
        }

        // 1.5) Flip reset (assumes no rotation/shear for detection; run **after** rotation cancel)
        if (opts.flip && hasMatrix(item)) {
            try {
                var mNow = item.matrix;
                var fh = isFlippedHorizontal(mNow);
                var fv = isFlippedVertical(mNow);
                if (fh && fv) {
                    unflipBoth(item);
                } else if (fh) {
                    unflipHorizontal(item);
                } else if (fv) {
                    unflipVertical(item);
                }
            } catch (e) {}
        }

        // 2) Equalize aspect ratio (if requested)
        //    Do this BEFORE absolute scaling so that uniformity survives normalization
        if (opts.ratio) {
            equalizeScaleToMax(item);
        }

        // 3) Absolute scale handling (normalize to 100% → apply desired %)
        if (opts.scale) {
            normalizeScaleOnly(item);
            applyUniformScalePercent(item, opts.scalePercent);
        }

        // 4) Shear removal LAST to eliminate any tiny shear reintroduced by the above
        if (opts.skew) {
            removeSkewOnly(item);
            // Safety: if a minute shear remains due to numeric noise, remove once more
            try {
                if (hasMatrix(item)) {
                    var m = item.matrix;
                    var d = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
                    if (Math.abs(d.shear) > 1e-6) {
                        removeSkewOnly(item);
                    }
                }
            } catch (e) {}
        }
    });
    // Run bounding box reset after reset operations (requested)
    try {
        app.selection = null;
        app.selection = [item];
        app.executeMenuCommand("AI Reset Bounding Box");
    } catch (e) {}
    // When rotate is ON, we exit with rotation at 0°. If OFF, original rotation is preserved.
}



main();