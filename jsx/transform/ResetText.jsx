#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ResetText.jsx（回転／シアー／スケールのリセット）

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- テキストに対して、回転／シアー（せん断）／比率を安全にリセット
- BBoxリセットと左上基準の位置復元により、見た目の位置を安定して維持

### 主な機能：

- テキスト：回転・シアー・水平/垂直比率（100%）のリセット

### 処理の流れ：

1) 選択されたテキストに対して処理を即時適用（回転／シアー／比率のリセット）

### note：

- スケールの微調整（Alt+↑↓による0.1%増減）は廃止

### 更新履歴：

- v1.0 (20250805) : 初期バージョン

---

### Script Name:

ResetText.jsx (Reset Rotate / Shear / Scale)

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Safely reset rotation/shear/ratio for Text
- Keeps visual position stable via BBox reset and top-left recentering

### Key Features:

- Text: reset rotation, shear, and horizontal/vertical scaling ratios (100%)

### Flow:

1) Apply operations immediately to selected text (reset rotation / shear / ratios)

### Notes:

- Fine scale tweak with Alt+Arrow (±0.1%) has been removed

### Changelog:

- v1.0 (20250805): Initial release
*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.0";

/* Language helper / 言語判定ヘルパー
   日本語環境なら "ja"、それ以外は "en" */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* UIラベル定義 / UI Label Definitions */
var LABELS = {
    alerts: {
        select: {
            ja: "オブジェクトを選択してください。",
            en: "Please select an object."
        },
        none: {
            ja: "該当するオブジェクトが選択されていません。",
            en: "No supported objects were selected."
        }
    }
};

/* 数値安定化用イプシロン（逆行列・QR分解でゼロ割や精度劣化を避ける） */
var CONFIG = {
    eps: 1e-8
};


function main() {
    if (!app.documents.length) {
        alert(LABELS.alerts.select[lang]);
        return;
    }

    var sel = activeDocument.selection;
    if (!sel || sel.length === 0) {
        alert(LABELS.alerts.select[lang]);
        return;
    }

    var processed = 0;
    var skipped = 0;

    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (!it || it.typename !== 'TextFrame') { skipped++; continue; }
        resetTextOps(it, true, true, true);
        processed += 1;
    }
    if (processed === 0) {
        alert(LABELS.alerts.none[lang]);
    }
}

/* 変形→BBoxリセット→左上基準で再配置 / Transform → reset BBox → recenter to top-left */
function withBBoxResetAndRecenter(item, opFn) {
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

function getRotationAngleDeg(a, b, sign) {
    var ang = Math.atan2(b, a) * 180 / Math.PI;
    return (sign < 0) ? -ang : ang;
}

/* 回転行列ヘルパー（可能なら Illustrator API を使用） / Rotation matrix helper (prefers Illustrator API) */
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

/* 指定角度で回転 / Rotate by degrees */
function rotateBy(item, deg) {
    var mat = getRotationMatrixSafe(deg);
    item.transform(mat);
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

/* シアーのみ除去 / Remove shear only */
function removeSkewOnly(item) {
    if (!item) return;

    // --- Local helpers (scoped to this function) ---
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
        return { a: d * invDet, b: -b * invDet, c: -c * invDet, d: a * invDet };
    }
    function toMatrix(obj) {
        var M = new Matrix();
        M.mValueA = obj.a; M.mValueB = obj.b; M.mValueC = obj.c; M.mValueD = obj.d;
        M.mValueTX = 0; M.mValueTY = 0;
        return M;
    }
    function buildFromQR(q1x, q1y, q2x, q2y, sx, sy, shear) {
        var r11 = sx, r12 = shear * sx, r21 = 0, r22 = sy;
        return {
            a: q1x * r11 + q2x * r21,
            b: q1y * r11 + q2y * r21,
            c: q1x * r12 + q2x * r22,
            d: q1y * r12 + q2y * r22
        };
    }
    function applyDeltaToMatch(target2x2) {
        var m = item.matrix;
        var cur = { a: m.mValueA, b: m.mValueB, c: m.mValueC, d: m.mValueD };
        var inv = invert2x2(cur.a, cur.b, cur.c, cur.d);
        var delta = multiply2x2(inv.a, inv.b, inv.c, inv.d, target2x2.a, target2x2.b, target2x2.c, target2x2.d);
        item.transform(toMatrix(delta));
    }
    // --- Compute target (shear=0) and apply ---
    var m = item.matrix;
    var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
    var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, dec.sx, dec.sy, 0);
    applyDeltaToMatch(target);
}

function hasMatrix(obj) {
    try {
        return !!(obj && obj.matrix && typeof obj.matrix.mValueA !== 'undefined');
    } catch (e) {
        return false;
    }
}

function cancelRotation(item, sign) {
    var a = item.matrix.mValueA;
    var b = item.matrix.mValueB;
    var rot = getRotationAngleDeg(a, b, sign);
    rotateBy(item, rot);
    return rot;
}

function recenterToTopLeft(item, tl, w1, h1) {
    var w2 = item.width,
        h2 = item.height;
    item.position = [tl[0] + w1 / 2 - w2 / 2, tl[1] - h1 / 2 + h2 / 2];
}

function resetTextScaleRatio(item) {
    if (!item) return; // guard unified at caller (resetTextOps ensures TextFrame)
    // 水平・垂直スケーリングを 100% に（= [1, 1]）
    item.textRange.scaling = [1, 1];
}

function resetTextOps(item, doRot, doShear, doRatio) {
    if (!item || !hasMatrix(item)) return;
    // One-shot BBox + recenter for text ops / テキスト処理を1回のBBoxでまとめて実行
    withBBoxResetAndRecenter(item, function() {
        if (doRot) cancelRotation(item, -1); // Text rotation (Raster-like sign)
        if (doShear) removeSkewOnly(item); // Remove shear
        if (doRatio) resetTextScaleRatio(item); // Reset H/V ratio
    });
}

main();