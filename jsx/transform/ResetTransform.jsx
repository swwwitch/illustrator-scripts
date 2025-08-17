#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### スクリプト名：

ResetTransform.jsx（回転／シアー／スケール／縦横比のリセット）

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/ResetTransform.jsx

### 概要：

- 更新日: 2025-08-17
- 配置画像・テキスト・長方形（パス）・クリップグループ・直線パスに対して、回転／シアー（せん断）／スケール／縦横比を安全にリセットします。
- BBox リセットと左上基準の再配置により、見た目の位置を安定して維持します。

### 主な機能：

- **配置画像／ラスタ**：回転・シアー・縦横比（小さい軸を大きい軸へ揃え、％は四捨五入の整数）・スケール（指定％）の個別／同時リセット。
- **テキスト**：回転・シアー・水平/垂直比率（100%）のリセット。
- **長方形（4点パス）**：近傍角度のみスナップして回転0°/90°へ補正。
- **直線（2点パス）**：近傍軸（0°/90°）へスナップして補正。
- **クリップグループ**：子要素（配置/ラスタ **＋ マスクパス**）の回転リセット＋縦横比（配置画像から導出した等比化Δを、配置画像とマスクパスに同時適用）。
- **UI**：2カラム、対象別パネル、ホットキー、数値入力（↑↓/Shift+↑↓）、ダイアログ位置/透明度。

### 処理の流れ：

1) 選択内容を解析し、対応パネルを自動ディム表示。
2) パネルで操作項目を選択 → OK。
3) 各ハンドラが変形を適用 → BBox リセット → 左上基準で再配置。

### 謝辞：

フジタノリアキさん

### 更新履歴：

- v1.5 (20250818) : 反転（上下／左右）の対応を追加
- v1.4 (20250817) : クリップグループの縦横比（配置画像＋マスクパスへ同時適用）を実装。複合パス・ネストに対応し、変形の確実性を向上。
- v1.3 (20250817) : 配置/ラスタの［縦横比］オプションを追加（100%/100% 基準の等比化、整数％丸め）。
- v1.2 (20250816) : スケール％入力とホットキー、選択内容によるパネル自動ディム、安定化（atan2・EPS 集約）。
- v1.1 (20250810) : クリップグループ回転の子要素処理、2カラムUI、最大面積代表子選定。
- v1.0 (20250805) : 初期バージョン（基本機能、ダイアログ、配置/ラスタ/テキスト/長方形対応）。

---

### Script Name:

ResetTransform.jsx (Reset Rotate / Shear / Scale / Aspect Ratio)

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/ResetTransform.jsx

### Overview:

- Updated: 2025-08-17
- Safely resets rotation / shear / scale / aspect ratio for Placed Images, Text, Rectangles (paths), Clip Groups, and straight Paths.
- Keeps visual position stable via BBox reset and top-left recentering.

### Key Features:

- **Placed/Raster**: Reset rotation, shear, aspect ratio (lift smaller axis to the larger; integer-rounded percent), and absolute scale, individually or together.
- **Text**: Reset rotation, shear, and horizontal/vertical scaling ratios (100%).
- **Rectangle (4-point path)**: Snap near angles to 0°/90°.
- **Line (2-point path)**: Snap to nearest axis (0°/90°).
- **Clip Group**: For children (placed/raster **and** clipping path), reset rotation and equalize aspect ratio; the uniformizing delta derived from the placed image is applied **to both** the placed image and the clipping path simultaneously. Supports nested groups and compound paths.
- **UI**: Two-column panels per target, hotkeys, numeric nudge (↑↓/Shift+↑↓), dialog position/opacity.

### Flow:

1) Analyze selection → dim unsupported panels.
2) Choose operations → OK.
3) Apply transforms per handler → reset BBox → recenter to top-left.

### Notes:

- Fine scale tweak with Alt+Arrow (±0.1%) is removed.
- Angle snap tolerance configurable via `CONFIG.rectSnapMin` / `CONFIG.rectSnapMax`.

### Changelog:

- v1.5 (20250818): Added Flip (horizontal/vertical) support.
- v1.4 (20250817): Implemented Clip Group aspect-ratio equalization (applies same delta to placed image + clipping path). Added nested/compound support and improved robustness.
- v1.3 (20250817): Added Placed/Raster "Aspect Ratio" option (equalization with integer rounding).
- v1.2 (20250816): Added Scale % input & hotkeys, selection-aware panel dimming, stability improvements (atan2 & EPS consolidation).
- v1.1 (20250810): Clip group child-wise rotation, two-column UI, largest-area child picking.
- v1.0 (20250805): Initial release (core features, dialog, placed/raster/text/rect).

*/

var SCRIPT_VERSION = "v1.5";

var CONFIG = {
    rectSnapMin: 0.5, // degrees
    rectSnapMax: 10, // degrees
    eps: 1e-8 // numerical epsilon for matrix ops
};

// --- Dialog position memory helpers (shared across scripts via targetengine) ---
function _getSavedLoc(key) {
    return $.global[key] && $.global[key].length === 2 ? $.global[key] : null;
}

function _setSavedLoc(key, loc) {
    $.global[key] = [loc[0], loc[1]];
}

function _clampToScreen(loc) {
    try {
        var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
        var x = Math.max(vb[0] + 10, Math.min(loc[0], vb[2] - 10));
        var y = Math.max(vb[1] + 10, Math.min(loc[1], vb[3] - 10));
        return [x, y];
    } catch (e) {
        return loc;
    }
}
// A unique storage key for this script's main dialog
var DLG_STORE_KEY = "__ResetTransform_OptionsDialog";


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
            ja: "配置画像",
            en: "Placed Images"
        },
        clip: {
            ja: "クリップグループ",
            en: "Clip Group"
        },
        text: {
            ja: "テキスト",
            en: "Text"
        },
        rect: {
            ja: "長方形（パス）",
            en: "Rectangle (Path)"
        },
        line: {
            ja: "パス（直線）",
            en: "Path (Line)"
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
        scale100: {
            ja: "スケール",
            en: "Scale"
        },
        textRatio: {
            ja: "垂直比率／水平比率",
            en: "Horizontal & Vertical Scale"
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
        placedOrRaster: false,
        clipGroup: false,
        text: false,
        rect: false,
        line: false
    };
    if (!sel || !sel.length) return caps;
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (!it || !it.typename) continue;
        var t = it.typename;
        if (t === 'PlacedItem' || t === 'RasterItem') caps.placedOrRaster = true;
        else if (t === 'GroupItem' && it.clipped === true) caps.clipGroup = true;
        else if (t === 'TextFrame') caps.text = true;
        else if (t === 'PathItem') {
            if (it.closed && it.pathPoints && it.pathPoints.length === 4) caps.rect = true;
            if (!it.closed && it.pathPoints && it.pathPoints.length === 2) caps.line = true;
        }
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

    // Override onShow to restore last position (or apply first-run offset)
    (function() {
        var saved = _getSavedLoc(DLG_STORE_KEY);
        dlg.onShow = function() {
            try {
                if (saved) {
                    dlg.location = _clampToScreen(saved);
                } else {
                    dlg.location = [dlg.location[0] + DIALOG_OFFSET_X, dlg.location[1] + DIALOG_OFFSET_Y];
                }
            } catch (e) {}
        };
        // Save whenever the dialog is moved
        dlg.onMove = function() {
            try {
                _setSavedLoc(DLG_STORE_KEY, [dlg.location[0], dlg.location[1]]);
            } catch (e) {}
        };
    })();

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

    // Enable state based on selection capabilities
    pnlReset.enabled = caps.placedOrRaster;
    cbRotate.enabled = caps.placedOrRaster;
    cbSkew.enabled = caps.placedOrRaster;
    cbRatio.enabled = caps.placedOrRaster;
    cbFlip.enabled = caps.placedOrRaster;
    cbScale.enabled = caps.placedOrRaster;
    etScale.enabled = caps.placedOrRaster && cbScale.value;
    stPercent.enabled = caps.placedOrRaster && cbScale.value;

    // Hotkey: 'S' toggles Scale / SキーでスケールON/OFF
    addHotkeyToggle(dlg, 'S', cbScale, function(enabled) {
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

    /* Panel: Clip Group / クリップグループ */
    var pnlClipGroup = colLeft.add('panel', undefined, LABELS.panels.clip[lang]);
    pnlClipGroup.alignment = 'fill';
    pnlClipGroup.orientation = 'column';
    pnlClipGroup.alignChildren = 'left';
    pnlClipGroup.margins = [15, 20, 15, 10];
    var cbClipGroupRotate = pnlClipGroup.add('checkbox', undefined, LABELS.checks.rotate[lang]);
    var cbClipGroupRatio = pnlClipGroup.add('checkbox', undefined, LABELS.checks.ratio[lang]);
    var cbClipGroupFlip = pnlClipGroup.add('checkbox', undefined, LABELS.checks.flip[lang]);
    cbClipGroupRotate.value = true; // default ON
    cbClipGroupRatio.value = true; // default ON
    cbClipGroupFlip.value = true; // default ON

    pnlClipGroup.enabled = caps.clipGroup;
    cbClipGroupRotate.enabled = caps.clipGroup;
    cbClipGroupRatio.enabled = caps.clipGroup;
    cbClipGroupFlip.enabled = caps.clipGroup;

    /* Panel: Text / テキスト */
    var pnlText = colRight.add('panel', undefined, LABELS.panels.text[lang]);
    pnlText.alignment = 'fill';
    pnlText.orientation = 'column';
    pnlText.alignChildren = 'left';
    pnlText.margins = [15, 20, 15, 10];
    var cbTextRotate = pnlText.add('checkbox', undefined, LABELS.checks.rotate[lang]);
    var cbTextSkew = pnlText.add('checkbox', undefined, LABELS.checks.skew[lang]);
    var cbTextScaleRatio = pnlText.add('checkbox', undefined, LABELS.checks.textRatio[lang]);
    // defaults: ON unless explicitly set to false
    cbTextRotate.value = true;
    cbTextSkew.value = true;
    cbTextScaleRatio.value = true;
    pnlText.enabled = caps.text;
    cbTextRotate.enabled = caps.text;
    cbTextSkew.enabled = caps.text;
    cbTextScaleRatio.enabled = caps.text;

    /* Panel: Rectangle (Path) / 長方形（パス） */
    var pnlRect = colRight.add('panel', undefined, LABELS.panels.rect[lang]);
    pnlRect.alignment = 'fill';
    pnlRect.orientation = 'column';
    pnlRect.alignChildren = 'left';
    pnlRect.margins = [15, 20, 15, 10];
    var cbRectRotate = pnlRect.add('checkbox', undefined, LABELS.checks.rotate[lang]);
    // default ON
    cbRectRotate.value = true;
    pnlRect.enabled = caps.rect;
    cbRectRotate.enabled = caps.rect;

    /* Panel: Path (Line) / パス（直線） */
    var pnlLine = colRight.add('panel', undefined, LABELS.panels.line[lang]);
    pnlLine.alignment = 'fill';
    pnlLine.orientation = 'column';
    pnlLine.alignChildren = 'left';
    pnlLine.margins = [15, 20, 15, 10];
    var cbLineRotate = pnlLine.add('checkbox', undefined, LABELS.checks.rotate[lang]);
    cbLineRotate.value = true; // default ON
    pnlLine.enabled = caps.line;
    cbLineRotate.enabled = caps.line;

    /* Dialog buttons / ダイアログボタン */
    var gBtns = dlg.add('group');
    gBtns.alignment = 'center';
    gBtns.add('button', undefined, LABELS.buttons.cancel[lang], {
        name: 'cancel'
    });
    gBtns.add('button', undefined, LABELS.buttons.ok[lang], {
        name: 'ok'
    });

    // Persist position on button clicks (OK / Cancel)
    try {
        var btnCancel = gBtns.children[0];
        var btnOk = gBtns.children[1];
        btnOk.onClick = function() {
            try {
                _setSavedLoc(DLG_STORE_KEY, [dlg.location[0], dlg.location[1]]);
            } catch (e) {}
            dlg.close(1);
        };
        btnCancel.onClick = function() {
            try {
                _setSavedLoc(DLG_STORE_KEY, [dlg.location[0], dlg.location[1]]);
            } catch (e) {}
            dlg.close(0);
        };
    } catch (e) {}

    if (dlg.show() !== 1) return null; // cancelled

    var out = {
        rotate: cbRotate.value,
        skew: cbSkew.value,
        ratio: cbRatio.value,
        flip: cbFlip.value,
        scale: cbScale.value,
        textRotate: cbTextRotate.value,
        textSkew: cbTextSkew.value,
        textScaleRatio: cbTextScaleRatio.value,
        rectRotate: cbRectRotate.value,
        lineRotate: cbLineRotate.value,
        clipGroupRotate: cbClipGroupRotate.value,
        clipGroupRatio: cbClipGroupRatio.value,
        clipGroupFlip: cbClipGroupFlip.value,
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

    function handleText(item) {
        var did = false;
        var needBBox = !!(opts.textRotate || opts.textSkew);
        if (needBBox) {
            resetTextOps(item, !!opts.textRotate, !!opts.textSkew, !!opts.textScaleRatio);
            did = true;
        } else if (opts.textScaleRatio) {
            // 比率のみならBBox不要で軽量処理
            resetTextScaleRatio(item);
            did = true;
        }
        return did ? res(1, 0, 0, 0) : res(0, 1, 0, 0);
    }

    function handleRect(item) {
        // PathItem（4点の長方形）の近傍角度のみ回転リセット
        if (opts.rectRotate && item.closed && item.pathPoints.length === 4) {
            var ok = applyRotationCorrection(item);
            return ok ? res(1, 0, 0, 0) : res(0, 1, 0, 0);
        }
        return res(0, 1, 0, 0);
    }

    function handleLine(item) {
        if (opts.lineRotate && !item.closed && item.pathPoints.length === 2) {
            var ok = applyRotationCorrectionLine(item);
            return ok ? res(1, 0, 0, 0) : res(0, 1, 0, 0);
        }
        return res(0, 1, 0, 0);
    }

    function handleClipGroup(item) {
        var did = false;
        if (item.clipped === true) {
            // 1) 回転を正す / Fix rotation first
            if (opts.clipGroupRotate) {
                var r1 = processClippedGroupChildren(item, opts);
                if (r1) did = true;
            }
            // 1.5) BBoxリセット / Reset BBox before flip detection
            if (opts.clipGroupRotate && opts.clipGroupFlip) {
                resetBBoxOnly(item);
            }
            // 2) 反転を調整 / Then handle flip
            if (opts.clipGroupFlip) {
                var rF = processClippedGroupFlip(item);
                if (rF) did = true;
            }
            // 3) 縦横比の等比化 / Aspect ratio
            if (opts.clipGroupRatio) {
                var r2 = processClippedGroupAspect(item);
                if (r2) did = true;
            }
        }
        return did ? res(1, 0, 0, 0) : res(0, 1, 0, 0);
    }

    function handleImage(item, typename) {
        setScaleTo100(item, typename, opts);
        // Placed/Raster は常に処理成功としてカウント（従来仕様）
        if (typename === 'PlacedItem') return res(1, 0, 1, 0);
        if (typename === 'RasterItem') return res(1, 0, 0, 1);
        return res(1, 0, 0, 0);
    }

    return {
        TextFrame: handleText,
        PathItem: function(item) {
            if (!item.closed && item.pathPoints && item.pathPoints.length === 2) {
                return handleLine(item); // 直線（2点）なら直線用ハンドラ
            }
            return handleRect(item); // それ以外は長方形（4点）など既存処理
        },
        GroupItem: handleClipGroup,
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

/* Helper: calculate "safe" area (width * height), or -1 if not available */
function getAreaSafe(it) {
    try {
        var w = it.width;
        var h = it.height;
        if (typeof w !== 'number' || typeof h !== 'number') return -1;
        return Math.abs(w * h);
    } catch (e) {
        return -1;
    }
}

/* Clip group: child-wise reset (rotate only) / クリップグループ：子要素単位の回転リセット */
function processClippedGroupChildren(groupItem, opts) {
    // Ensure BBox is up-to-date before reading matrices
    resetBBoxOnly(groupItem);

    // Recursively find representative placed/raster and clipping path
    var found = findPlacedAndClipRecursive(groupItem);
    var placedOrRaster = found.placed;
    var clipPath = found.clip;
    var clipTarget = resolveClipTransformTarget(clipPath);

    // Compute a shared rotation delta from the representative child
    var rotDelta = 0;
    var haveDelta = false;
    if (placedOrRaster && hasMatrix(placedOrRaster)) {
        var pm = placedOrRaster.matrix;
        var psign = (placedOrRaster.typename === 'RasterItem') ? -1 : 1;
        rotDelta = getRotationAngleDeg(pm.mValueA, pm.mValueB, psign);
        haveDelta = Math.abs(rotDelta) > 0.0001;
    } else if (clipTarget && hasMatrix(clipTarget)) {
        var cm = clipTarget.matrix;
        rotDelta = getRotationAngleDeg(cm.mValueA, cm.mValueB, +1);
        haveDelta = Math.abs(rotDelta) > 0.0001;
    }

    var did = false;
    if (opts.clipGroupRotate && haveDelta) {
        // Rotate the GROUP once so children keep their relative alignment
        withBBoxResetAndRecenter(groupItem, function () {
            rotateBy(groupItem, rotDelta);
        });
        did = true;
    }
    return did;
}

// Recursively find the largest placed/raster and the clipping path (PathItem or CompoundPathItem) within a clipped group
function findPlacedAndClipRecursive(container) {
    var bestPlaced = null,
        bestPlacedArea = -1;
    var bestClip = null,
        bestClipArea = -1;
    var items = (container.pageItems || []);
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;
        var t = it.typename;
        if (t === 'PlacedItem' || t === 'RasterItem') {
            var ap = getAreaSafe(it);
            if (ap > bestPlacedArea) {
                bestPlaced = it;
                bestPlacedArea = ap;
            }
        } else if (t === 'PathItem') {
            if (it.clipping) {
                var ac1 = getAreaSafe(it);
                if (ac1 > bestClipArea) {
                    bestClip = it;
                    bestClipArea = ac1;
                }
            }
        } else if (t === 'CompoundPathItem') {
            // CompoundPathItem itself may not expose `.clipping`; check its children
            var subPaths = it.pathItems || [];
            for (var k = 0; k < subPaths.length; k++) {
                var sp = subPaths[k];
                try {
                    if (sp && sp.clipping) {
                        var ac2 = getAreaSafe(it); // use compound's area as proxy
                        if (ac2 > bestClipArea) {
                            bestClip = it;
                            bestClipArea = ac2;
                        }
                        break;
                    }
                } catch (e) {}
            }
        } else if (t === 'GroupItem') {
            var found = findPlacedAndClipRecursive(it);
            if (found.placed && getAreaSafe(found.placed) > bestPlacedArea) {
                bestPlaced = found.placed;
                bestPlacedArea = getAreaSafe(found.placed);
            }
            if (found.clip && getAreaSafe(found.clip) > bestClipArea) {
                bestClip = found.clip;
                bestClipArea = getAreaSafe(found.clip);
            }
        }
    }
    return {
        placed: bestPlaced,
        clip: bestClip
    };
}

// Apply transform with full flags and center origin; fallback safely if enum not available
function transformAll(item, mat) {
    try {
        item.transform(mat, true, true, true, true, true, Transformation.CENTER);
    } catch (e) {
        try {
            item.transform(mat, true, true, true, true, true);
        } catch (e2) {
            item.transform(mat);
        }
    }
}
// Temporarily unlock/unhide while applying a function, then restore
function withUnlockedVisible(item, fn) {
    var wasLocked = false,
        wasHidden = false;
    try {
        wasLocked = !!item.locked;
    } catch (e) {}
    try {
        wasHidden = !!item.hidden;
    } catch (e) {}
    try {
        try {
            item.locked = false;
        } catch (e) {}
        try {
            item.hidden = false;
        } catch (e) {}
        fn();
    } finally {
        try {
            if (wasLocked) item.locked = true;
        } catch (e) {}
        try {
            if (wasHidden) item.hidden = true;
        } catch (e) {}
    }
}

// Reset only the bounding box for a given item (no recenter) / バウンディングボックスのみをリセット（位置は維持）
function resetBBoxOnly(item) {
    try {
        app.selection = null;
        app.selection = [item];
        app.executeMenuCommand("AI Reset Bounding Box");
    } catch (e) {}
}

// Resolve actual transform target for a clipping path: if CompoundPathItem, use its clipping PathItem
function resolveClipTransformTarget(clipCandidate) {
    if (!clipCandidate) return null;
    try {
        if (clipCandidate.typename === 'CompoundPathItem') {
            var subs = clipCandidate.pathItems || [];
            for (var i = 0; i < subs.length; i++) {
                var sp = subs[i];
                try {
                    if (sp && sp.clipping) return sp; // prefer the actual clipping child
                } catch (e) {}
            }
            // fallback: transform the compound itself if no child flagged
            return clipCandidate;
        }
        return clipCandidate; // PathItem or others
    } catch (e) {
        return clipCandidate;
    }
}

// Clip group: aspect ratio equalization for placed/raster & clipping path
// クリップグループ：配置画像＋マスクパス（PathItem/CompoundPathItem）に再帰的に同じ等比化を適用
function processClippedGroupAspect(groupItem) {
    // Find placed/raster and (PathItem or CompoundPathItem) clipping path recursively
    var found = findPlacedAndClipRecursive(groupItem);
    var placedOrRaster = found.placed;
    var clipPath = found.clip;
    var clipTarget = resolveClipTransformTarget(clipPath);
    // Proceed if we have at least the placed/raster; clipPath is optional
    if (!placedOrRaster) return false;
    if (!hasMatrix(placedOrRaster)) return false; // needed to compute delta from placed image
    var haveClip = !!clipTarget; // we can still transform a clip even if it lacks `.matrix`
    var clipHasMatrix = !!(haveClip && hasMatrix(clipTarget));

    // Build target uniform scale from placed image: lift smaller axis to larger, snap to integer percent
    var pm = placedOrRaster.matrix;
    var dec = decomposeQR(pm.mValueA, pm.mValueB, pm.mValueC, pm.mValueD);
    var u = Math.max(dec.sx, dec.sy);
    var uPercent = Math.round(u * 100); // round to integer percent
    u = uPercent / 100;

    // Delta matrix from current to (u,u) with same orientation/shear
    var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, u, u, dec.shear);
    var cur = {
        a: pm.mValueA,
        b: pm.mValueB,
        c: pm.mValueC,
        d: pm.mValueD
    };
    var inv = invert2x2(cur.a, cur.b, cur.c, cur.d);
    var delta = multiply2x2(inv.a, inv.b, inv.c, inv.d, target.a, target.b, target.c, target.d);
    var deltaMat = toMatrix(delta);

    // Always apply to placed/raster; also apply to clipTarget if present (even without `.matrix`)
    withUnlockedVisible(placedOrRaster, function() {
        transformAll(placedOrRaster, deltaMat);
    });
    if (haveClip) {
        withUnlockedVisible(clipTarget, function() {
            transformAll(clipTarget, deltaMat);
        });
    }

    // Verify and fallback per-item to guarantee uniform scale
    try {
        var mPl = placedOrRaster.matrix,
            dPl = decomposeQR(mPl.mValueA, mPl.mValueB, mPl.mValueC, mPl.mValueD);
        if (Math.abs(dPl.sx - dPl.sy) > 1e-6) {
            withUnlockedVisible(placedOrRaster, function() {
                equalizeScaleToMax(placedOrRaster);
            });
        }
        if (clipHasMatrix) {
            var mCp = clipTarget.matrix,
                dCp = decomposeQR(mCp.mValueA, mCp.mValueB, mCp.mValueC, mCp.mValueD);
            if (Math.abs(dCp.sx - dCp.sy) > 1e-6) {
                withUnlockedVisible(clipTarget, function() {
                    equalizeScaleToMax(clipTarget);
                });
            }
        }
    } catch (e) {}
    return true;
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

function resetTextRotation(item) {
    withBBoxResetAndRecenter(item, function() {
        // Text rotates with Raster-like sign (逆方向補正)
        cancelRotation(item, -1);
    });
}

function resetTextScaleRatio(item) {
    if (!item || !item.textRange) return; // safe guard
    try {
        // Pass 1: primary API
        item.textRange.scaling = [1, 1];
        // Also set legacy attributes for robustness across builds
        var ca = item.textRange.characterAttributes;
        if (ca) {
            try {
                ca.horizontalScale = 100;
            } catch (e1) {}
            try {
                ca.verticalScale = 100;
            } catch (e2) {}
        }
        // Pass 2: re-apply to kill residual rounding noise
        item.textRange.scaling = [1, 1];
    } catch (e) {}
}

function resetTextOps(item, doRot, doShear, doRatio) {
    // One-shot BBox + recenter for text ops / テキスト処理を1回のBBoxでまとめて実行
    withBBoxResetAndRecenter(item, function() {
        // 1) Rotate first (stabilize orientation for text) / まず回転を0°へ
        if (doRot) cancelRotation(item, -1); // Text rotation (Raster-like sign)

        // 2) Shear removal (before ratio) / シアー除去を先に
        if (doShear) {
            removeSkewOnly(item);
            // Safety: minute residual shear → remove once more
            try {
                if (hasMatrix(item)) {
                    var m = item.matrix;
                    var d = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
                    if (Math.abs(d.shear) > 1e-6) removeSkewOnly(item);
                }
            } catch (e) {}
        }

        // 3) Ratio equalization LAST / 比率（水平・垂直）を最後に
        if (doRatio) {
            resetTextScaleRatio(item); // pass 1
            // Best-effort verification and second pass
            try {
                var sc = item.textRange && item.textRange.scaling;
                if (!sc || Math.abs(sc[0] - 1) > 1e-6 || Math.abs(sc[1] - 1) > 1e-6) {
                    resetTextScaleRatio(item); // pass 2
                }
            } catch (e) {
                // Fallback: second pass anyway
                resetTextScaleRatio(item);
            }
        }
    });
}


function applyRotationCorrection(rect) {
    if (!rect || !rect.closed || rect.pathPoints.length !== 4) return false;

    /* Rotation from first two anchors / 最初の2点から角度を算出 */
    var ptA = rect.pathPoints[0].anchor;
    var ptB = rect.pathPoints[1].anchor;

    var dx = ptB[0] - ptA[0];
    var dy = ptB[1] - ptA[1];
    var angleRad = Math.atan2(dy, dx);
    var angleDeg = angleRad * 180 / Math.PI;
    if (angleDeg < 0) angleDeg += 360;

    var rotationAmount = angleDeg; // 水平からのズレ / offset from horizontal
    var normalized = rotationAmount % 90;
    if (normalized > 45) normalized = 90 - normalized; // 最近傍の 0/90 に正規化 / distance to nearest axis

    // 0.5〜10° の範囲でのみ補正 / only snap if near axis to avoid false positives
    if (normalized >= CONFIG.rectSnapMin && normalized <= CONFIG.rectSnapMax) {
        rect.rotate(-rotationAmount);
        return true;
    }
    return false;
}

function applyRotationCorrectionLine(line) {
    if (!line || line.closed || line.pathPoints.length !== 2) return false;

    var ptA = line.pathPoints[0].anchor;
    var ptB = line.pathPoints[1].anchor;

    var dx = ptB[0] - ptA[0];
    var dy = ptB[1] - ptA[1];
    if (dx === 0 && dy === 0) return false;

    // Current angle from horizontal in degrees (-180..180)
    var angleRad = Math.atan2(dy, dx);
    var a = angleRad * 180 / Math.PI;

    // Compute minimal deltas to 0° (horizontal) and 90° (vertical)
    function norm180(x) {
        while (x > 180) x -= 360;
        while (x < -180) x += 360;
        return x;
    }

    function clamp90(x) {
        x = norm180(x);
        if (x > 90) x -= 180;
        if (x < -90) x += 180;
        return x;
    }

    var d0 = clamp90(-a); // rotate by this to reach 0°
    var d90 = clamp90(90 - a); // rotate by this to reach 90°

    var abs0 = Math.abs(d0);
    var abs90 = Math.abs(d90);
    var dist = Math.min(abs0, abs90);

    // Optional: only snap when reasonably close to axis, mirroring rectangle rule
    if (dist < CONFIG.rectSnapMin || dist > CONFIG.rectSnapMax) return false;

    var rot = (abs0 <= abs90) ? d0 : d90;
    withBBoxResetAndRecenter(line, function() {
        rotateBy(line, rot);
    });
    return true;
}

function setScaleTo100(item, objectType, opts) {
    var sign = (objectType === "RasterItem") ? -1 : 1;
    if (!opts) opts = {};
    // Fast path: do nothing if no operations are enabled
    if (!opts.rotate && !opts.skew && !opts.scale && !opts.ratio && !opts.flip) return;

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
    // When rotate is ON, we exit with rotation at 0°. If OFF, original rotation is preserved.
}


// Clip group: flip reset for placed/raster & clipping path
// クリップグループ：配置画像＋マスクパスの反転を同時に解除
function processClippedGroupFlip(groupItem) {
    // Ensure matrices reflect post-rotation state / 回転補正後の行列で正しく判定するためにBBoxを更新
    resetBBoxOnly(groupItem);
    var found = findPlacedAndClipRecursive(groupItem);
    var placedOrRaster = found.placed;
    var clipPath = found.clip;
    var clipTarget = resolveClipTransformTarget(clipPath);
    var haveAny = !!(placedOrRaster || clipTarget);
    if (!haveAny) return false;

    // Prefer detection from placed/raster when available; otherwise fall back to clip
    var mat = null,
        src = null;
    if (placedOrRaster && hasMatrix(placedOrRaster)) {
        mat = placedOrRaster.matrix;
        src = 'placed';
    } else if (clipTarget && hasMatrix(clipTarget)) {
        mat = clipTarget.matrix;
        src = 'clip';
    }
    if (!mat) return false;

    var fh = isFlippedHorizontal(mat);
    var fv = isFlippedVertical(mat);
    if (!fh && !fv) return false;

    // Apply identical unflip to both placed and clip (if present)
    function applyUnflip(it) {
        if (!it) return;
        withUnlockedVisible(it, function() {
            if (fh && fv) {
                unflipBoth(it);
            } else if (fh) {
                unflipHorizontal(it);
            } else if (fv) {
                unflipVertical(it);
            }
        });
    }
    applyUnflip(placedOrRaster);
    applyUnflip(clipTarget);
    return true;
}

main();