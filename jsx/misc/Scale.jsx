#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

Scale.jsx（配置画像のスケール）

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択した配置画像／ラスタ画像のスケールを、安全に指定パーセントへ一括適用します。
- BBox をリセットし、左上基準で再配置することで、見た目の位置ズレを抑えます。

### 主な機能：

- 配置画像／ラスタのみ対象（テキスト・長方形・直線・クリップグループの処理は削除）
- スケールは常に有効（チェックなし）。数値フィールドに％を入力
- ダイアログ表示時にテキストフィールドへ自動フォーカス＆全選択
- 選択内容に応じてUIを有効／無効化

### 処理の流れ：

1) 選択内容を解析（配置画像／ラスタが含まれるか）
2) ダイアログでスケール％を入力 → OK
3) 各アイテムに対し、スケール正規化（100%）→ 指定％を適用 → BBoxリセット → 左上基準で再配置

### オリジナル、謝辞：

（該当なし）

### note：

- スケール微調整用ホットキー（Alt+↑↓ 0.1%）は廃止済み
- 角度スナップ関連の設定（CONFIG.rectSnapMin / rectSnapMax）は現行機能では未使用

### 動画：

（該当なし）

### 更新履歴：

- v1.0 (20250805) : 初期バージョン（回転・シアー・スケール、配置/ラスタ/テキスト/長方形対応）
- v1.1 (20250810) : クリップグループ回転処理、2カラムUI、代表子選定の安定化
- v1.2 (20250816) : スケール％入力とホットキー、選択内容によるパネル自動ディム
- v1.3 (20250816) : 機能を「配置画像／ラスタのスケール」に集約。回転/シアーUIとロジック、テキスト/長方形/直線/クリップグループ処理、パネル枠を削除。スケール常時ON、起動時フォーカスと「スケール」ラベルを追加

---

### Script Name:

Scale.jsx (Scale Placed/Raster Images)

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Scales selected Placed and Raster items to a specified percent safely.
- Resets the bounding box and recenters by the top-left corner to minimize visual shift.

### Key Features:

- Targets Placed/Raster only (Text/Rectangle/Line/Clip Group removed)
- Scale is always enabled (no checkbox). Enter percent in the input field
- Auto-focus & select the percent field on dialog open
- UI enables/disables based on selection

### Flow:

1) Analyze selection (detect Placed/Raster)
2) Enter scale percent → OK
3) For each item: normalize to 100% → apply percent → reset BBox → recenter using top-left

### Original / Credits:

(n/a)

### Notes:

- Fine-tune hotkeys (Alt+Arrow ±0.1%) removed
- Angle snap settings (CONFIG.rectSnapMin / rectSnapMax) are unused in the current scope

### Changelog:

- v1.0 (20250805): Initial release (rotate/shear/scale; placed/raster/text/rect)
- v1.1 (20250810): Clip group rotation handling; two-column UI; stable representative child selection
- v1.2 (20250816): Scale % input & hotkey; selection-aware panel dimming
- v1.3 (20250816): Feature set narrowed to scaling Placed/Raster only; removed rotate/shear UI & logic, text/rect/line/clip-group handling, and panel frame; scale always-on; added on-open focus and "Scale" label
*/

var SCRIPT_VERSION = "v1.2";

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
        ja: "配置画像のスケール " + SCRIPT_VERSION,
        en: "Scale Placed Images " + SCRIPT_VERSION
    },
    panels: {
        placed: {
            ja: "配置画像",
            en: "Placed Images"
        }
    },
    checks: {
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

/* Analyze selection → capabilities / 選択状態から対応可否を判定 */
function getSelectionCapabilities(sel) {
    var caps = { placedOrRaster: false };
    if (!sel || !sel.length) return caps;
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (!it || !it.typename) continue;
        var t = it.typename;
        if (t === 'PlacedItem' || t === 'RasterItem') {
            caps.placedOrRaster = true;
            break; // found one, no need to continue
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
                value = Math.floor((value - 1) / delta) * delta;
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


function showOptionsDialog() {

/* Main group: one-column layout / 1カラムレイアウト */
var gMain = dlg.add('group');
gMain.orientation = 'column';
gMain.alignChildren = 'left';
gMain.alignment = 'fill';

/* Placed Images contents */
var pnlReset = gMain.add('group');
    pnlReset.alignment = 'fill';
    pnlReset.orientation = 'column';
    pnlReset.alignChildren = 'left';
    pnlReset.margins = [15, 20, 15, 10];

    // Scale percent field (always enabled when selection is supported)
    var gScale = pnlReset.add('group');
    gScale.orientation = 'row';
    gScale.alignChildren = 'center';

    var lbScale = gScale.add('statictext', undefined, LABELS.checks.scale100[lang]);
    var etScale = gScale.add('edittext', undefined, '100');
    etScale.characters = 5; // default width
    var stPercent = gScale.add('statictext', undefined, '%');

    // Focus the scale field when dialog opens (without clobbering existing onShow)
    (function(){
        var prevOnShow = dlg.onShow;
        dlg.onShow = function() {
            try { if (typeof prevOnShow === 'function') prevOnShow(); } catch (e) {}
            try {
                etScale.active = true;
                if (typeof etScale.select === 'function') {
                    etScale.select();
                } else if (typeof etScale.textselection !== 'undefined') {
                    etScale.textselection = etScale.text;
                }
            } catch (e) {}
        };
    })();

    // Arrow-key increment/decrement support
    changeValueByArrowKey(etScale);

    // Always enabled; non-target items are skipped at processing time
    pnlReset.enabled = true;
    etScale.enabled  = true;
    stPercent.enabled= true;
    lbScale.enabled  = true;

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
        scale: true,
        scalePercent: (function() {
            var n = parseFloat(etScale.text, 10);
            if (isNaN(n)) n = 100;
            if (n <= 0) n = 100;
            return n;
        })()
    };
    return out;
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

    var processed = 0, skipped = 0, countPlaced = 0, countRaster = 0;
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (!item || !item.typename) { skipped++; continue; }
        var t = item.typename;
        if (t === 'PlacedItem' || t === 'RasterItem') {
            setScaleTo100(item, t, opts);
            processed++;
            if (t === 'PlacedItem') countPlaced++; else countRaster++;
        } else {
            skipped++;
        }
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


function normalizeScaleOnly(item) {
    if (!item || !hasMatrix(item)) return;
    var m = item.matrix;
    var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
    var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, 1, 1, dec.shear);
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


function recenterToTopLeft(item, tl, w1, h1) {
    var w2 = item.width,
        h2 = item.height;
    item.position = [tl[0] + w1 / 2 - w2 / 2, tl[1] - h1 / 2 + h2 / 2];
}


function setScaleTo100(item, objectType, opts) {
    if (!opts) opts = {};
    if (!opts.scale) return;

    withBBoxResetAndRecenter(item, function() {
        normalizeScaleOnly(item);
        applyUniformScalePercent(item, opts.scalePercent);
    });
}

main();