#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/**
 * ResetTransform (Rotate / Shear / Aspect / Flip / Scale / Re-place) for Illustrator
 * Version: v1.10
 * Updated: 2025-12-21
 *
 * Notes for embedding:
 * - This script is structured as a module exposed on $.global.ResetTransform.
 * - To prevent auto-run when embedding, set $.global.__RESET_TRANSFORM_EMBEDDED__ = true before loading.
 * - Call $.global.ResetTransform.run() (UI) or $.global.ResetTransform.applyToSelection(opts, selection).
 */

(function() {
    'use strict';

    // ------------------------------
    // Public module (single global)
    // ------------------------------
    var MODULE_NAME = 'ResetTransform';
    var SCRIPT_VERSION = 'v1.10';

    // Keep ONE exported object on $.global for easy reuse.
    var api = $.global[MODULE_NAME] || {};
    $.global[MODULE_NAME] = api;

    // ------------------------------
    // Config
    // ------------------------------
    var CONFIG = {
        rectSnapMin: 0.5, // degrees
        rectSnapMax: 10, // degrees
        eps: 1e-8 // numerical epsilon for matrix ops
    };

    // ------------------------------
    // Localization
    // ------------------------------
    function getCurrentLang() {
        return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var LABELS = {
        dialogTitle: {
            ja: 'リセット（回転・比率） ' + SCRIPT_VERSION,
            en: 'Reset (Rotate / Scale) ' + SCRIPT_VERSION
        },
        panels: {
            placed: { ja: 'リセット', en: 'Reset' }
        },
        checks: {
            rotate: { ja: '回転', en: 'Rotate' },
            skew: { ja: 'シアー', en: 'Shear' },
            ratio: { ja: '縦横比', en: 'Aspect Ratio' },
            flip: { ja: '反転', en: 'Flip' },
            replaceReset: { ja: '再配置', en: 'Replacing' },
            scale100: { ja: 'スケール', en: 'Scale' }
        },
        buttons: {
            ok: { ja: 'リセット', en: 'Reset' },
            cancel: { ja: 'キャンセル', en: 'Cancel' }
        },
        alerts: {
            select: { ja: 'オブジェクトを選択してください。', en: 'Please select an object.' },
            none: { ja: '該当するリンク/ラスター画像が選択されていません。', en: 'No placed/raster images were selected.' }
        }
    };

    // ------------------------------
    // Dialog helpers (position/opacity)
    // ------------------------------
    var DIALOG_OFFSET_X = 300;
    var DIALOG_OFFSET_Y = 0;
    var DIALOG_OPACITY = 0.98;

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
        try { dlg.opacity = opacityValue; } catch (e) {}
    }

    // ------------------------------
    // Selection capabilities
    // ------------------------------
    function getSelectionCapabilities(sel) {
        var caps = { placedOrRaster: false };
        if (!sel || !sel.length) return caps;
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            if (!it || !it.typename) continue;
            var t = it.typename;
            if (t === 'PlacedItem' || t === 'RasterItem') caps.placedOrRaster = true;
        }
        return caps;
    }

    // ------------------------------
    // UI small utilities
    // ------------------------------
    function changeValueByArrowKey(editText) {
        editText.addEventListener('keydown', function(event) {
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
                    value = Math.round((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
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
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

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

    // ------------------------------
    // Options dialog (returns opts or null)
    // ------------------------------
    function showOptionsDialog(lang, selection) {
        var prefs = { rotate: true, skew: true, scale: false };

        var selNow = selection || ((app && app.activeDocument) ? app.activeDocument.selection : null);
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
        var panelReset = gMain.add('group');
        panelReset.orientation = 'column';
        panelReset.alignChildren = 'left';
        panelReset.alignment = 'fill';

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
        etScale.characters = 5;
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
                    etScale.active = true;
                    if (typeof etScale.select === 'function') {
                        etScale.select();
                    } else if (typeof etScale.textselection !== 'undefined') {
                        etScale.textselection = etScale.text;
                    }
                } catch (e) {}
            }
        };
        changeValueByArrowKey(etScale);

        // Helper: enable/disable panel items based on re-place reset and selection
        function updatePlacedPanelEnabled() {
            var supported = !!caps.placedOrRaster;
            pnlReset.enabled = supported;
            cbReplaceReset.enabled = supported;

            var othersOn = supported && !cbReplaceReset.value;
            cbRotate.enabled = othersOn;
            cbSkew.enabled  = othersOn;
            cbRatio.enabled = othersOn;
            cbFlip.enabled  = othersOn;
            cbScale.enabled = othersOn;

            etScale.enabled = othersOn && cbScale.value;
            stPercent.enabled = othersOn && cbScale.value;
        }

        cbReplaceReset.onClick = function() { updatePlacedPanelEnabled(); };
        updatePlacedPanelEnabled();

        // Hotkey: 'S' toggles Scale
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
        // Hotkey: 'F' toggles Flip
        addHotkeyToggle(dlg, 'F', cbFlip);

        /* Dialog buttons */
        var gBtns = dlg.add('group');
        gBtns.alignment = 'center';
        gBtns.add('button', undefined, LABELS.buttons.cancel[lang], { name: 'cancel' });
        gBtns.add('button', undefined, LABELS.buttons.ok[lang], { name: 'ok' });

        if (dlg.show() !== 1) return null;

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
                n = Math.round(n);
                return n;
            })()
        };
        return out;
    }

    // ------------------------------
    // Dispatch / handlers
    // ------------------------------
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
            if (typename === 'PlacedItem') return res(1, 0, 1, 0);
            if (typename === 'RasterItem') return res(1, 0, 0, 1);
            return res(1, 0, 0, 0);
        }

        return {
            PathItem: function(item) { return res(0, 1, 0, 0); },
            PlacedItem: function(item) { return handleImage(item, 'PlacedItem'); },
            RasterItem: function(item) { return handleImage(item, 'RasterItem'); }
        };
    }

    // ------------------------------
    // Public entry points
    // ------------------------------
    function runUI() {
        var lang = getCurrentLang();
        var doc = app.activeDocument;
        var sel = doc ? doc.selection : null;
        if (!sel || sel.length === 0) {
            alert(LABELS.alerts.select[lang]);
            return;
        }

        var opts = showOptionsDialog(lang, sel);
        if (!opts) return;

        var result = applyToSelection(opts, sel);
        if (result.processed === 0) {
            alert(LABELS.alerts.none[lang]);
        }
    }

    function applyToSelection(opts, selection) {
        var sel = selection || (app.activeDocument ? app.activeDocument.selection : null);
        var processed = 0;
        var skipped = 0;
        var countPlaced = 0;
        var countRaster = 0;

        if (!sel || !sel.length) {
            return { processed: 0, skipped: 0, countPlaced: 0, countRaster: 0 };
        }

        var handlers = makeHandlers(opts || {});

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
            skipped++;
        }

        return { processed: processed, skipped: skipped, countPlaced: countPlaced, countRaster: countRaster };
    }

    // Export API
    api.version = SCRIPT_VERSION;
    api.getCurrentLang = getCurrentLang;
    api.showOptionsDialog = showOptionsDialog;
    api.applyToSelection = applyToSelection;
    api.run = runUI;

    // ------------------------------
    // Helpers / transform logic
    // ------------------------------
    function withBBoxResetAndRecenter(item, opFn) {
        var tl = item.position;
        var w1 = item.width;
        var h1 = item.height;
        if (typeof opFn === 'function') opFn();
        try {
            app.selection = null;
            app.selection = [item];
            app.executeMenuCommand('AI Reset Bounding Box');
        } catch (e) {}
        recenterToTopLeft(item, tl, w1, h1);
    }

    function isWindows() {
        return Folder.fs === 'Windows';
    }

    function isEPS(placedItem) {
        if (!placedItem || !placedItem.file) return false;
        return isWindows() ?
            (String(placedItem.file.displayName).match(/[^.]+$/i) == 'eps') :
            (placedItem.file.type == 'EPSF');
    }

    function normalizeZero(n) {
        // Convert -0 to 0 (JS has signed zero)
        return (n === 0) ? 0 : n;
    }

    function getRotationAngleDeg(a, b, sign) {
        var ang = Math.atan2(b, a) * 180 / Math.PI;
        ang = (sign < 0) ? -ang : ang;
        return normalizeZero(ang);
    }

    function getRotationMatrixSafe(deg) {
        try {
            if (app && typeof app.getRotationMatrix === 'function') {
                return app.getRotationMatrix(deg);
            }
        } catch (e) {}
        var rad = deg * Math.PI / 180.0;
        var cosv = Math.cos(rad);
        var sinv = Math.sin(rad);
        var M = new Matrix();
        M.mValueA = cosv;
        M.mValueB = sinv;
        M.mValueC = -sinv;
        M.mValueD = cosv;
        M.mValueTX = 0;
        M.mValueTY = 0;
        return M;
    }

    function rotateBy(item, deg) {
        var mat = getRotationMatrixSafe(deg);
        item.transform(mat);
    }

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
        var q1x = a / sx;
        var q1y = b / sx;
        var r12 = q1x * c + q1y * d;
        var u2x = c - r12 * q1x;
        var u2y = d - r12 * q1y;
        var sy = Math.sqrt(u2x * u2x + u2y * u2y);
        if (sy === 0) {
            sy = CONFIG.eps;
            u2x = -q1y;
            u2y = q1x;
        }
        var q2x = u2x / sy;
        var q2y = u2y / sy;
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
        var r11 = sx;
        var r12 = shear * sx;
        var r21 = 0;
        var r22 = sy;
        return {
            a: q1x * r11 + q2x * r21,
            b: q1y * r11 + q2y * r21,
            c: q1x * r12 + q2x * r22,
            d: q1y * r12 + q2y * r22
        };
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
        var uPercent = Math.round(u * 100);
        u = uPercent / 100;
        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, u, u, dec.shear);
        applyDeltaToMatch(item, target);
    }

    function applyUniformScalePercent(item, percent) {
        var p = Number(percent);
        if (!(p > 0)) return;
        try {
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

    function isFlippedHorizontal(mat) {
        return mat && mat.mValueA < 0;
    }

    // Placed/Raster はデフォルトで mValueD が負になり得るため、上下反転は mValueD > 0 を基準に判定
    function isFlippedVertical(mat) {
        return mat && mat.mValueD > 0;
    }

    function unflipHorizontal(item) {
        if (!item) return;
        item.transform(app.getScaleMatrix(-100, 100), true, true, true, true, true, Transformation.CENTER);
    }

    function unflipVertical(item) {
        if (!item) return;
        item.transform(app.getScaleMatrix(100, -100), true, true, true, true, true, Transformation.CENTER);
    }

    function unflipBoth(item) {
        if (!item) return;
        item.transform(app.getScaleMatrix(-100, -100), true, true, true, true, true, Transformation.CENTER);
    }

    function cancelRotation(item, sign) {
        if (!hasMatrix(item)) return 0;
        var a = item.matrix.mValueA;
        var b = item.matrix.mValueB;
        var rot = getRotationAngleDeg(a, b, sign);
        rotateBy(item, rot);
        return rot;
    }

    function recenterToTopLeft(item, tl, w1, h1) {
        var w2 = item.width;
        var h2 = item.height;
        item.position = [tl[0] + w1 / 2 - w2 / 2, tl[1] - h1 / 2 + h2 / 2];
    }


    // --- PlacedItem re-place helper ---
    function replacePlacedItemToReset(placedItem) {
        if (!placedItem || !placedItem.file) return null;
        try {
            // Use ONLY position/width/height and match by CENTER.
            // Illustrator's position is top-left of the (current) bounding box.
            var srcPos = placedItem.position; // [left, top]
            var srcW = placedItem.width;
            var srcH = placedItem.height;
            var srcCX = srcPos[0] + srcW / 2;
            var srcCY = srcPos[1] - srcH / 2;

            var doc = app.activeDocument;
            var newItem = doc.placedItems.add();
            newItem.file = placedItem.file;
            newItem.move(placedItem, ElementPlacement.PLACEBEFORE);

            // Normalize scale to 100% (keeps the script's existing intent)
            try { normalizeScaleOnly(newItem); } catch (e) {}

            // Reset bounding box BEFORE positioning (otherwise the later reset can shift the center)
            try {
                app.selection = null;
                app.selection = [newItem];
                app.executeMenuCommand('AI Reset Bounding Box');
            } catch (e) {}

            // Carry over a few attributes (best-effort)
            try { if (placedItem.name) newItem.name = placedItem.name; } catch (e) {}
            try { if (typeof placedItem.opacity !== 'undefined') newItem.opacity = placedItem.opacity; } catch (e) {}
            try { if (typeof placedItem.blendingMode !== 'undefined') newItem.blendingMode = placedItem.blendingMode; } catch (e) {}

            // Place by center (position is top-left)
            var newW = newItem.width;
            var newH = newItem.height;
            newItem.position = [srcCX - newW / 2, srcCY + newH / 2];

            try { placedItem.remove(); } catch (e) {}
            return newItem;
        } catch (e) {}
        return null;
    }

    function setScaleTo100(item, objectType, opts) {
        var sign = (objectType === 'RasterItem') ? -1 : 1;
        if (!opts) opts = {};
        if (!opts.rotate && !opts.skew && !opts.scale && !opts.ratio && !opts.flip && !opts.replaceReset) return;

        if (opts.replaceReset && objectType === 'PlacedItem') {
            replacePlacedItemToReset(item);
            return;
        }

        withBBoxResetAndRecenter(item, function() {
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
                } catch (e) {}
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

        try {
            app.selection = null;
            app.selection = [item];
            app.executeMenuCommand('AI Reset Bounding Box');
        } catch (e) {}
    }

    // ------------------------------
    // Auto-run guard
    // ------------------------------
    if (!$.global.__RESET_TRANSFORM_EMBEDDED__) {
        try { api.run(); } catch (e) {}
    }

})();