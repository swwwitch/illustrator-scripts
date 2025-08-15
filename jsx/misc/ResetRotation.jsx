#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：
ResetRotation.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/ResetRotation.jsx

### 概要：

- 選択オブジェクトの回転を水平(0°)に補正。
- テキスト、配置画像（リンク/埋め込み）、長方形（パス）を対象に、見かけ上の回転を0°に補正。
- 画像がクリッピンググループ内にある場合は、回転が乗っている“親（ホスト）”を自動で検出して補正。
- 縦書きテキストはスキップ。ほぼ水平（許容角度内）は無視。

### 主な機能：

- 対象の選択：テキスト / 配置画像 / 長方形（チェックボックス）
- 補正条件：水平とみなす範囲（°）を指定。↑/↓=±1、Shift+↑/↓=±10、Option+↑/↓=±0.1 のキー操作に対応。

### 処理の流れ：

1) 選択オブジェクトを再帰走査し、タイプ別に収集。  
2) 変換行列（matrix）から回転角を推定（パスは頂点からの推定をフォールバック）。  
3) 鏡像（負スケール）を考慮して必要角度だけ回転。  
4) 回転直後に各オブジェクトへ個別に「Reset Bounding Box」を実行。

### オリジナル、謝辞：

### note：

- 許容角度（しきい値）は CONFIG.epsilonDeg で管理し、ダイアログから変更可能。  
- 回転はオブジェクト中心（Transformation.CENTER）。  
- ロック/非表示等で回転不可の場合はスキップ（例外は握りつぶし）。

### 更新履歴：

- v1.1 (20250815) : しきい値UI、矢印キー増減、クリップグループ対応、画像の鏡像補正、回転直後の Reset Bounding Box を実装
- v1.0 (20250815) : 初期バージョン  

---

### Script name:
ResetRotation.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/ResetRotation.jsx

### Overview:

- Level selected objects to horizontal (0°).  
- Targets text, placed/embedded images, and rectangles (paths).  
- Automatically detects and corrects rotation on parent (host) when images are inside clipping groups.  
- Skips vertical text and ignores nearly horizontal objects (within tolerance).

### Main features:

- Selection options: Text / Placed Image / Rectangle (checkboxes).  
- Correction tolerance range (degrees) adjustable via dialog with keyboard shortcuts (↑/↓=±1, Shift+↑/↓=±10, Option+↑/↓=±0.1).

### Processing flow:

1) Recursively traverse selected objects and collect by type.  
2) Estimate rotation angle from transformation matrix (fallback to path vertices for paths).  
3) Rotate by required angle considering mirrored (negative scale) transforms.  
4) Execute "Reset Bounding Box" individually on each object immediately after rotation.

### Original, acknowledgments:

### note:

- Tolerance threshold is managed by CONFIG.epsilonDeg and adjustable from dialog.  
- Rotation is performed about object center (Transformation.CENTER).  
- Skips locked/hidden objects that cannot be rotated (exceptions are suppressed).

### Changelog:

- v1.1 (20250815) : Added tolerance UI and arrow-key increments, clipping-group rotation, mirrored image handling, and per-item Reset Bounding Box after rotation
- v1.0 (20250815) : Initial release  

*/

var SCRIPT_VERSION = "v1.2";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    dialogTitle: {
        ja: "水平補正 " + SCRIPT_VERSION,
        en: "Level Objects " + SCRIPT_VERSION
    },
    panelTargets: {
        ja: "対象",
        en: "Objects to Level"
    },
    panelOptions: {
        ja: "補正条件",
        en: "Correction Options"
    },
    text: {
        ja: "テキスト",
        en: "Text"
    },
    image: {
        ja: "配置画像（埋め込み・リンク）",
        en: "Placed/Embedded Image"
    },
    rect: {
        ja: "長方形（パス）",
        en: "Rectangle (Path)"
    },
    clipGroup: {
        ja: "クリップグループ",
        en: "Clipping Group"
    },
    clipScopePanel: {
        ja: "クリップ範囲",
        en: "Clip Scope"
    },
    clipScopeNearest: {
        ja: "直近のクリップ",
        en: "Nearest Clipped"
    },
    clipScopeTopmost: {
        ja: "最上位のクリップ",
        en: "Topmost Clipped"
    },
    epsilon: {
        ja: "水平とみなす範囲(°)",
        en: "Level Tolerance (°)"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

// デフォルト選択（全部オン）およびしきい値
var CONFIG = {
    defaultTargets: {
        text: true,
        image: true,
        rect: true
    },
    epsilonDeg: 0.1, // ほぼ水平とみなす閾値（度）
    clipGroup: true, // クリップグループをデフォルトON
    clipScope: 'nearest' // 'nearest' | 'topmost'
};

function isNearlyLevel(deg) {
    var d = Math.abs(deg);
    var eps = CONFIG.epsilonDeg;
    return (d < eps) || (Math.abs(d - 360) < eps);
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

function showTargetDialog(defaults) {
    var w = new Window("dialog", LABELS.dialogTitle[lang]);

    var pTargets = w.add("panel", undefined, LABELS.panelTargets[lang]);
    pTargets.orientation = "column";
    pTargets.alignChildren = "left";
    pTargets.margins = [15, 20, 15, 10];
    pTargets.alignment = "left";

    var cbText = pTargets.add("checkbox", undefined, LABELS.text[lang]);
    var cbImage = pTargets.add("checkbox", undefined, LABELS.image[lang]);
    var cbRect = pTargets.add("checkbox", undefined, LABELS.rect[lang]);

    cbText.value = !!defaults.text;
    cbImage.value = !!defaults.image;
    cbRect.value = !!defaults.rect;

    // 対象パネル内にクリップグループのチェックボックスを配置
    var cbClip = pTargets.add("checkbox", undefined, LABELS.clipGroup[lang]);
    cbClip.value = !!(typeof defaults.clipGroup !== 'undefined' ? defaults.clipGroup : CONFIG.clipGroup);

    // クリップ範囲（直近 / 最上位）
    var pScope = pTargets.add("panel", undefined, LABELS.clipScopePanel[lang]);
    pScope.orientation = "row";
    pScope.alignChildren = "left";
    pScope.margins = [12, 10, 12, 10];
    var rbNearest = pScope.add("radiobutton", undefined, LABELS.clipScopeNearest[lang]);
    var rbTopmost = pScope.add("radiobutton", undefined, LABELS.clipScopeTopmost[lang]);

    var scopeDefault = (typeof defaults.clipScope !== 'undefined') ? defaults.clipScope : CONFIG.clipScope;
    rbNearest.value = (scopeDefault === 'nearest');
    rbTopmost.value = (scopeDefault === 'topmost');

    // オプション：しきい値
    var pOpts = w.add("panel", undefined, LABELS.panelOptions[lang]);
    pOpts.orientation = "column";
    pOpts.alignChildren = "left";
    pOpts.margins = [15, 20, 15, 10];
    pOpts.alignment = "left";
    var gEps = pOpts.add("group");
    gEps.orientation = "row";
    gEps.alignChildren = "left";
    gEps.add("statictext", undefined, LABELS.epsilon[lang]);
    var etEps = gEps.add("edittext", undefined, String((typeof defaults.epsilonDeg !== 'undefined') ? defaults.epsilonDeg : CONFIG.epsilonDeg));
    etEps.characters = 6;
    changeValueByArrowKey(etEps);

    // (Mode panel removed, rbClip now in Targets panel)

    var gBtns = w.add("group");
    gBtns.alignment = "center";
    var btnCancel = gBtns.add("button", undefined, LABELS.cancel[lang], {
        name: "cancel"
    });
    var btnOK = gBtns.add("button", undefined, LABELS.ok[lang], {
        name: "ok"
    });

    var ok = false;
    btnOK.onClick = function() {
        ok = true;
        w.close();
    };
    btnCancel.onClick = function() {
        w.close();
    };

    w.layout.layout(true);
    w.center();
    w.show();

    var epsVal = CONFIG.epsilonDeg;
    try {
        var v = parseFloat(etEps.text);
        if (!isNaN(v)) epsVal = Math.max(0.01, Math.min(10, v)); // clamp 0.01–10°
    } catch (eParse) {}
    return {
        ok: ok,
        targets: {
            text: cbText.value,
            image: cbImage.value,
            rect: cbRect.value
        },
        epsilonDeg: epsVal,
        clipGroup: cbClip.value
        ,clipScope: (rbTopmost.value ? 'topmost' : 'nearest')
    };
}

function isPlacedImage(obj) {
    // リンク配置画像
    return obj && obj.typename === "PlacedItem";
}

function isEmbeddedImage(obj) {
    // 埋め込みラスタ
    return obj && obj.typename === "RasterItem";
}

function isRectangle(obj) {
    // 単純な長方形（閉じた4点）を想定
    if (!obj || obj.typename !== "PathItem") return false;
    if (!obj.closed) return false;
    if (obj.pathPoints.length !== 4) return false;
    return true;
}

function getImageTransformHost(item) {
    // 画像がクリッピンググループ内にある場合、回転はグループ側に乗っていることが多い
    var host = item;
    try {
        var p = item.parent;
        while (p && p.typename === 'GroupItem' && p.clipped) {
            host = p; // 直近のクリップグループに回転が載る
            p = p.parent;
        }
    } catch (e) {}
    return host;
}

function getNearestClippingGroupHost(item) {
    // 親を辿って最初に見つかった clipped=true の GroupItem を返す
    try {
        var p = item.parent;
        while (p) {
            if (p.typename === 'GroupItem' && p.clipped) return p;
            p = p.parent;
        }
    } catch (e) {}
    return item; // 見つからなければ自身
}

function getTopmostClippingGroupHost(item) {
    // 親を辿って最上位の clipped=true の GroupItem を返す
    var host = null;
    try {
        var p = item.parent;
        while (p) {
            if (p.typename === 'GroupItem' && p.clipped) host = p;
            p = p.parent;
        }
    } catch (e) {}
    return host || item; // 見つからなければ自身
}

function getClippingGroupHost(item, scope) {
    return (scope === 'topmost') ? getTopmostClippingGroupHost(item) : getNearestClippingGroupHost(item);
}

function resolveHost(item) {
    // クリップグループモードならグループをホストに、そうでなければ自身
    return CONFIG.clipGroup ? getClippingGroupHost(item, CONFIG.clipScope) : item;
    // ※ clipGroup=true: グループを回転 / clipGroup=false: 個別オブジェクトを回転
}

function isVertical(tf) {
    try {
        var attr = tf.textRanges[0].characterAttributes;
        return attr.orientation === TextOrientation.VERTICAL;
    } catch (e) {
        return false;
    }
}

function normalizeAngle(deg) {
    // → (-180, 180]
    var a = deg % 360;
    if (a <= -180) a += 360;
    if (a > 180) a -= 360;
    return a;
}

function processTextFrames(textFrames) {
    var fixed = 0,
        skippedVertical = 0,
        alreadyLevel = 0;
    var rotated = [];
    var seen = {};
    for (var i = 0; i < textFrames.length; i++) {
        var tf = textFrames[i];
        var host = resolveHost(tf);

        // クリップグループ外（= 自身が TextFrame）のときのみ縦書きスキップ判定
        if (host === tf) {
            if (isVertical(tf)) {
                skippedVertical++;
                continue;
            }
        }

        // 同一ホストは一度だけ処理
        var id = host.uuid || (host.toString() + '|' + host.index);
        if (seen[id]) continue;
        seen[id] = true;

        var ang = getRotationDegrees(host);
        var norm = normalizeAngle(ang);
        if (isNearlyLevel(norm)) {
            alreadyLevel++;
            continue;
        }
        // 画像と同様、ホストが鏡像の可能性を考慮（Group にも反転が載ることがある）
        var mirrored = isMirroredTransform(host);
        var delta = mirrored ? norm : -norm;
        rotateBy(host, delta);
        rotated.push(host);
        fixed++;
    }
    return {
        fixed: fixed,
        skippedVertical: skippedVertical,
        alreadyLevel: alreadyLevel,
        rotated: rotated
    };
}

function processImages(items) {
    // 画像ごとに“回転のホスト”を解決し、重複を避けて処理
    var fixed = 0,
        alreadyLevel = 0;
    var seen = {};
    var rotated = [];
    for (var i = 0; i < items.length; i++) {
        // クリップグループONならグループ、OFFなら画像単体を回転
        var host = resolveHost(items[i]);
        var id = host.uuid || (host.toString() + '|' + host.index);
        if (seen[id]) continue; // 同一ホストの重複回避
        seen[id] = true;

        var ang = getRotationDegrees(host);
        var norm = normalizeAngle(ang);
        if (isNearlyLevel(norm)) {
            alreadyLevel++;
            continue;
        }

        // 鏡像（負のスケール・反転）を含む場合は回転の符号が逆になることがあるため補正
        var mirrored = isMirroredTransform(host);
        var delta = mirrored ? norm : -norm;
        rotateBy(host, delta);
        rotated.push(host);
        fixed++;
    }
    return {
        fixed: fixed,
        alreadyLevel: alreadyLevel,
        rotated: rotated
    };
}

function processItems(items) {
    var fixed = 0,
        alreadyLevel = 0;
    var rotated = [];
    var seen = {};
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        var host = resolveHost(it);
        var id = host.uuid || (host.toString() + '|' + host.index);
        if (seen[id]) continue;
        seen[id] = true;

        var ang = getRotationDegrees(host);
        var norm = normalizeAngle(ang);
        if (isNearlyLevel(norm)) {
            alreadyLevel++;
            continue;
        }
        var mirrored = isMirroredTransform(host);
        var delta = mirrored ? norm : -norm;
        rotateBy(host, delta);
        rotated.push(host);
        fixed++;
    }
    return {
        fixed: fixed,
        alreadyLevel: alreadyLevel,
        rotated: rotated
    };
}


function getGroupProxyAngle(grp) {
    // クリップグループの見かけ角度を子要素から推定
    // 優先順：クリッピングパス → 配置/埋め込み画像 → 任意のパス → ネストグループ
    try {
        // 1) クリッピングパスの角度
        for (var i = 0; i < grp.pageItems.length; i++) {
            var it = grp.pageItems[i];
            if (it.typename === 'PathItem' && it.clipping) {
                return getRotationDegrees(it);
            }
        }
        // 2) 配置/埋め込み画像
        for (var j = 0; j < grp.pageItems.length; j++) {
            var it2 = grp.pageItems[j];
            if (it2.typename === 'PlacedItem' || it2.typename === 'RasterItem') {
                return getRotationDegrees(it2);
            }
        }
        // 3) 任意のパス（非クリッピング）
        for (var k = 0; k < grp.pageItems.length; k++) {
            var it3 = grp.pageItems[k];
            if (it3.typename === 'PathItem' && !it3.clipping) {
                return getRotationDegrees(it3);
            }
        }
        // 4) ネストされたグループ
        for (var m = 0; m < grp.pageItems.length; m++) {
            var it4 = grp.pageItems[m];
            if (it4.typename === 'GroupItem') {
                var a = getGroupProxyAngle(it4);
                if (!isNearlyLevel(a)) return a;
            }
        }
    } catch (e) {}
    return 0;
}

function getRotationDegrees(item) {
    // Prefer the transformation matrix if available
    var m = null;
    try { m = item.matrix; } catch (e) { m = null; }

    if (m && m.mValueA !== undefined && m.mValueB !== undefined) {
        var rad = Math.atan2(m.mValueB, m.mValueA);
        var deg = rad * 180 / Math.PI;
        // GroupItem の場合、見かけは子の回転に依存することがある
        if (item && item.typename === 'GroupItem') {
            if (isNearlyLevel(deg)) {
                var proxy = getGroupProxyAngle(item);
                if (!isNearlyLevel(proxy)) return proxy;
            }
        }
        return deg;
    }

    // Fallback for PathItem: infer angle from the first segment (anchor0 -> anchor1)
    if (item && item.typename === 'PathItem' && item.pathPoints && item.pathPoints.length >= 2) {
        try {
            var p0 = item.pathPoints[0].anchor;
            var p1 = item.pathPoints[1].anchor;
            var rad2 = Math.atan2(p1[1] - p0[1], p1[0] - p0[0]);
            return rad2 * 180 / Math.PI;
        } catch (e2) {}
    }

    // GroupItem で matrix が無い/読めない場合のフォールバック
    if (item && item.typename === 'GroupItem') {
        var proxy2 = getGroupProxyAngle(item);
        if (!isNearlyLevel(proxy2)) return proxy2;
    }

    // As a last resort, assume 0° (prevents crash on exotic items)
    return 0;
}

function isMirroredTransform(item) {
    // 判定: 行列の行列式 < 0 なら反転（鏡像）を含む
    try {
        var m = item.matrix;
        if (m && m.mValueA !== undefined) {
            var det = (m.mValueA * m.mValueD) - (m.mValueB * m.mValueC);
            return det < 0;
        }
    } catch (e) {}
    return false;
}

function rotateBy(item, deg) {
    var prevSel = null;
    try {
        prevSel = app.selection;
    } catch (e) {
        prevSel = null;
    }
    try {
        // 回転（changePositions=true, patterns/gradients/strokes rotate=true, center about）
        item.rotate(deg, true, true, true, true, Transformation.CENTER);

        // 回転直後に対象のみ選択して Reset Bounding Box を実行
        try {
            app.selection = null;
            app.selection = [item];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (eInner) {}
    } catch (e) {
        // ignore rotation errors (e.g., locked/hidden)
    }
    // 元の選択へ復帰
    try {
        if (prevSel) app.selection = prevSel;
    } catch (e2) {}
}


function main() {
    if (!app.documents.length) {
        alert('ドキュメントが開いていません / No document open');
        return;
    }
    if (!app.selection || app.selection.length === 0) {
        alert('オブジェクトを選択してください / Please select objects');
        return;
    }

    // ダイアログで対象タイプを選択
    var dlg = showTargetDialog({
        text: CONFIG.defaultTargets.text,
        image: CONFIG.defaultTargets.image,
        rect: CONFIG.defaultTargets.rect,
        epsilonDeg: CONFIG.epsilonDeg,
        clipGroup: CONFIG.clipGroup,
        clipScope: CONFIG.clipScope
    });
    if (!dlg.ok) {
        return;
    }
    var targets = dlg.targets;
    CONFIG.epsilonDeg = dlg.epsilonDeg;
    CONFIG.clipGroup = dlg.clipGroup;
    CONFIG.clipScope = dlg.clipScope;

    // 選択を走査してタイプ別に収集
    var sel = app.selection;
    var textArr = [];
    var imageArr = [];
    var rectArr = [];

    // 再帰探索ユーティリティ
    function collect(items) {
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.typename === 'GroupItem') {
                collect(it.pageItems);
            } else if (it.typename === 'CompoundPathItem') {
                collect(it.pathItems);
            } else if (it.pageItems) {
                // クリップグループ等
                try {
                    collect(it.pageItems);
                } catch (e) {}
            }

            if (targets.text && it.typename === 'TextFrame') textArr.push(it);
            if (targets.image && (isPlacedImage(it) || isEmbeddedImage(it))) imageArr.push(it);
            if (targets.rect && isRectangle(it)) {
                if (CONFIG.clipGroup && it.clipping) {
                    // クリップグループモードではマスクパスは個別対象にしない
                } else {
                    rectArr.push(it);
                }
            }
        }
    }
    collect(sel);

    var changed = 0;
    var r1 = null,
        r2 = null,
        r3 = null;
    if (textArr.length) {
        r1 = processTextFrames(textArr);
        changed += r1.fixed;
    }
    if (imageArr.length) {
        r2 = processImages(imageArr);
        changed += r2.fixed;
    }
    if (rectArr.length) {
        r3 = processItems(rectArr);
        changed += r3.fixed;
    }

    if (changed > 0) app.redraw();
}

// 実行 / run
main();