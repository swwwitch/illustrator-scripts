#target illustrator
// Use a script-specific engine for session-persistent values (not across restarts)
#targetengine "CreateGradientFromSelectionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.8";

/*
  CreateGradientFromSelection.jsx

  選択オブジェクトの塗り／線カラーを、配置順（左→右、上→下）で抽出し、
  スウォッチグループに登録してグラデーションを自動生成します。

  ・グループ／複合パス／テキストは再帰的に処理
  ・塗り（フィル）と線（ストローク）の両方を対象
  ・抽出色をスウォッチ化（必要に応じてグローバルカラー（プロセス）に変換）
  ・抽出色数に合わせて線形グラデーションを作成（各ストップにスウォッチの色を適用）
  ・オプションで「グローバルカラー化／グラデーション作成／長方形作成」を切り替え可能
  ・長方形を作成する場合、サイズは「固定(100)」または「選択オブジェクトに合わせる」を選択可能
  ・長方形の配置は、選択が横並びなら下方向へ、縦並びなら右方向へ“長方形1個分”ずらして配置
  ・縦並び判定時は、アクション（gradient/90degree）で角度調整を実行（アクションが無い場合は無言でスキップ）
  ・（任意）長方形の見た目をグラフィックスタイルに登録
  ・ドキュメントが無い／選択が無い／色が1色以下の場合やエラー発生時は無言で終了
  ・オプションで「セパレートグラデーション」（ラジオ）を切り替え可能（2〜4色: 100÷色数で自動分割）

  Version: v1.8
  更新日: 2026-02-18
*/

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "グラデーション作成",
        en: "Create Gradient"
    },
    globalColor: {
        ja: "グローバルカラー",
        en: "Global colors"
    },
    createGradient: {
        ja: "グラデーションを作成",
        en: "Create gradient"
    },
    separateGradient: {
        ja: "セパレートグラデーション",
        en: "Separate gradients"
    },
    normalGradient: {
        ja: "通常",
        en: "Normal"
    },
    createRect: {
        ja: "長方形を作成し、グラデーションを適用",
        en: "Create rectangle and apply gradient"
    },
    useSelectionSize: {
        ja: "選択オブジェクトに合わせてサイズ指定",
        en: "Match rectangle size to selection"
    },
    registerGraphicStyle: {
        ja: "グラフィックスタイルに登録",
        en: "Save as Graphic Style"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    panelColor: {
        ja: "カラー",
        en: "Colors"
    },
    panelRect: {
        ja: "長方形（適用）",
        en: "Rectangle"
    }
};

function L(key) {
    try { return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key; } catch (e) { return key; }
}

// 縦並び時: グラデーション角度を90度にするアクションを実行 / If vertical: run action to set gradient angle to 90 degrees
function runGradientAngle90Action() {
    var actionSetName = "gradient";
    var actionName = "90degree";

    // アクション定義テキスト（改行は CR を使用）
    var CR = String.fromCharCode(13);
    var actionCode = [
        " /version 3",
        "/name [ 8",
        "\t6772616469656e74",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        "\t/name [ 8",
        "\t\t3930646567726565",
        "\t]",
        "\t/keyIndex 0",
        "\t/colorIndex 0",
        "\t/isOpen 1",
        "\t/eventCount 1",
        "\t/event-1 {",
        "\t\t/useRulersIn1stQuadrant 0",
        "\t\t/internalName (ai_plugin_setGradient)",
        "\t\t/localizedName [ 30",
        "\t\t\te382b0e383a9e38387e383bce382b7e383a7e383b3e38292e8a8ade5ae9a",
        "\t\t]",
        "\t\t/isOpen 1",
        "\t\t/isOn 1",
        "\t\t/hasDialog 0",
        "\t\t/parameterCount 1",
        "\t\t/parameter-1 {",
        "\t\t\t/key 1634625388",
        "\t\t\t/showInPalette 4294967295",
        "\t\t\t/type (unit real)",
        "\t\t\t/value -90.0",
        "\t\t\t/unit 591490663",
        "\t\t}",
        "\t}",
        "}",
        ""
    ].join(CR);

    try {
        var tempFile = new File(Folder.temp + "/temp_action.aia");
        tempFile.open("w");
        tempFile.write(actionCode);
        tempFile.close();

        app.loadAction(tempFile);
        app.doScript(actionName, actionSetName);
        app.unloadAction(actionSetName, "");

        try { tempFile.remove(); } catch (eDel) { }
    } catch (e) {
        // 無言
        try { app.unloadAction(actionSetName, ""); } catch (e2) { }
    }
}

function main() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;

    // 選択があるか確認
    if (!doc.selection || doc.selection.length === 0) {
        return;
    }

    // 選択の並びを推定（この結果で後続の処理を分岐するため保持）
    var selOri = detectSelectionOrientation(doc.selection);

    var selBounds = getSelectionBounds(doc.selection); // 選択の外接（サイズ算出用）

    /* =========================================
     * Options dialog / オプションダイアログ
     * ========================================= */

    // Persist dialog values in this targetengine (session-only) / ダイアログ値をエンジン内に保持（Illustrator再起動では消える）
    function getSessionSettings() {
        try {
            if (!$.global.__CGFS_SETTINGS) {
                $.global.__CGFS_SETTINGS = {};
            }
            return $.global.__CGFS_SETTINGS;
        } catch (e) {
            return {};
        }
    }
    function loadBool(key, defVal) {
        try {
            var s = getSessionSettings();
            if (typeof s[key] === 'boolean') return s[key];
        } catch (e) { }
        return defVal;
    }
    function saveBool(key, val) {
        try {
            var s = getSessionSettings();
            s[key] = !!val;
        } catch (e) { }
    }

    var opts = {
        makeGlobal: loadBool('makeGlobal', true),
        makeGradient: loadBool('makeGradient', true),
        makeRect: loadBool('makeRect', true),
        useSelectionSize: loadBool('useSelectionSize', true),
        registerGraphicStyle: loadBool('registerGraphicStyle', true),
        separateGradient: loadBool('separateGradient', false)
    };

    try {
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];

        var pColor = dlg.add('panel', undefined, L('panelColor'));
        pColor.orientation = 'column';
        pColor.alignChildren = ['fill', 'top'];
        pColor.margins = [15, 20, 15, 10];

        var cbGlobal = pColor.add('checkbox', undefined, L('globalColor'));
        cbGlobal.value = opts.makeGlobal;

        var cbGradient = pColor.add('checkbox', undefined, L('createGradient'));
        cbGradient.value = opts.makeGradient;

        // Separate gradient option (radio)
        var gSep = pColor.add('group');
        gSep.orientation = 'row';
        gSep.alignChildren = ['left', 'center'];

        var rbGradNormal = gSep.add('radiobutton', undefined, L('normalGradient'));
        var rbGradSeparate = gSep.add('radiobutton', undefined, L('separateGradient'));

        rbGradSeparate.value = !!opts.separateGradient;
        rbGradNormal.value = !rbGradSeparate.value;

        var pRect = dlg.add('panel', undefined, L('panelRect'));
        pRect.orientation = 'column';
        pRect.alignChildren = ['fill', 'top'];
        pRect.margins = [15, 20, 15, 10];

        var cbRect = pRect.add('checkbox', undefined, L('createRect'));
        cbRect.value = opts.makeRect;

        var cbSelSize = pRect.add('checkbox', undefined, L('useSelectionSize'));
        cbSelSize.value = opts.useSelectionSize;

        var cbGStyle = pRect.add('checkbox', undefined, L('registerGraphicStyle'));
        cbGStyle.value = opts.registerGraphicStyle;

        function syncEnable() {
            cbRect.enabled = cbGradient.value;
            cbSelSize.enabled = cbGradient.value && cbRect.value;
            // Graphic style can be created even if rectangle output is OFF (temporary rectangle)
            cbGStyle.enabled = cbGradient.value;
            gSep.enabled = cbGradient.value;

            if (!cbGradient.value) cbRect.value = false;
            if (!cbGradient.value) cbSelSize.value = false;
            if (!cbGradient.value) cbGStyle.value = false;
            if (!cbGradient.value) { rbGradSeparate.value = false; rbGradNormal.value = true; }

            cbSelSize.enabled = cbGradient.value && cbRect.value;
            cbGStyle.enabled = cbGradient.value;
            gSep.enabled = cbGradient.value;
        }
        cbGradient.onClick = syncEnable;
        cbRect.onClick = syncEnable;
        syncEnable();

        var btns = dlg.add('group');
        btns.alignment = 'right';
        var cancelBtn = btns.add('button', undefined, L('cancel'), { name: 'cancel' });
        var okBtn = btns.add('button', undefined, L('ok'), { name: 'ok' });

        function persistFromUI() {
            saveBool('makeGlobal', cbGlobal.value);
            saveBool('makeGradient', cbGradient.value);
            saveBool('makeRect', cbRect.value);
            saveBool('useSelectionSize', cbSelSize.value);
            saveBool('registerGraphicStyle', cbGStyle.value);
            saveBool('separateGradient', rbGradSeparate.value);
        }
        dlg.onClose = function () {
            try { persistFromUI(); } catch (e) { }
        };

        if (dlg.show() !== 1) {
            return; // キャンセル時は無言で終了
        }

        opts.makeGlobal = !!cbGlobal.value;
        opts.makeGradient = !!cbGradient.value;
        opts.makeRect = !!cbRect.value;
        opts.useSelectionSize = !!cbSelSize.value;
        opts.registerGraphicStyle = !!cbGStyle.value;
        opts.separateGradient = !!rbGradSeparate.value;
        try { persistFromUI(); } catch (ePersist) { }
    } catch (eDlg) {
        // ダイアログ生成に失敗しても無言で既定値のまま続行
    }

    /* =========================================
     * Color utilities / カラー取得ユーティリティ
     * ========================================= */

    function isNoColor(c) {
        // NoColor は typename が "NoColor" になる
        try {
            return (c == null) || (c.typename === "NoColor");
        } catch (e) {
            return true;
        }
    }

    function colorKey(c) {
        // 重複除去用の簡易キー
        if (!c) return "null";
        var t = c.typename;
        try {
            if (t === "RGBColor") {
                return "RGB:" + [c.red, c.green, c.blue].join(",");
            }
            if (t === "CMYKColor") {
                return "CMYK:" + [c.cyan, c.magenta, c.yellow, c.black].join(",");
            }
            if (t === "GrayColor") {
                return "Gray:" + c.gray;
            }
            if (t === "SpotColor") {
                // スポットはスポット名＋濃度
                var spotName = (c.spot && c.spot.name) ? c.spot.name : "(spot)";
                return "Spot:" + spotName + ":" + c.tint;
            }
            if (t === "PatternColor") {
                var patName = (c.pattern && c.pattern.name) ? c.pattern.name : "(pattern)";
                return "Pattern:" + patName;
            }
            if (t === "GradientColor") {
                var gName = (c.gradient && c.gradient.name) ? c.gradient.name : "(gradient)";
                return "Gradient:" + gName;
            }
        } catch (e) { }
        return "Other:" + t;
    }

    function pushUniqueColor(list, seenMap, c) {
        if (isNoColor(c)) return;
        var k = colorKey(c);
        if (seenMap[k]) return;
        seenMap[k] = true;
        list.push(c);
    }


    // 位置情報（左上）を取得 / Get top-left position
    function getItemTopLeft(item) {
        // geometricBounds: [left, top, right, bottom]
        try {
            var b = item.geometricBounds;
            return { left: b[0], top: b[1] };
        } catch (e) {
            return { left: 0, top: 0 };
        }
    }

    // 選択範囲の外接バウンディングを取得 / Get union bounds of selection
    // 戻り値: { left:Number, top:Number, right:Number, bottom:Number } または null
    function getSelectionBounds(selection) {
        try {
            if (!selection || selection.length === 0) return null;

            var left = 1e12, top = -1e12, right = -1e12, bottom = 1e12;
            var got = false;

            for (var i = 0; i < selection.length; i++) {
                var it = selection[i];
                if (!it) continue;
                try {
                    var b = it.geometricBounds; // [left, top, right, bottom]
                    if (b[0] < left) left = b[0];
                    if (b[1] > top) top = b[1];
                    if (b[2] > right) right = b[2];
                    if (b[3] < bottom) bottom = b[3];
                    got = true;
                } catch (eB) {
                    // 無視
                }
            }

            if (!got) return null;
            if (left > right || bottom > top) return null;

            return { left: left, top: top, right: right, bottom: bottom };
        } catch (e) {
            return null;
        }
    }

    // 選択オブジェクトが横並びか縦並びかを推定 / Detect whether selection is horizontal or vertical
    // 戻り値: { orientation: "horizontal"|"vertical"|"mixed"|"unknown", dx: Number, dy: Number, ratio: Number }
    function detectSelectionOrientation(selection) {
        try {
            if (!selection || selection.length < 2) {
                return { orientation: "unknown", dx: 0, dy: 0, ratio: 0 };
            }

            var minX = 1e12, maxX = -1e12;
            var minY = 1e12, maxY = -1e12;

            // 各アイテムの中心点を使って分布を測る
            for (var i = 0; i < selection.length; i++) {
                var it = selection[i];
                if (!it) continue;

                try {
                    var b = it.geometricBounds; // [left, top, right, bottom]
                    var cx = (b[0] + b[2]) / 2;
                    var cy = (b[1] + b[3]) / 2;

                    if (cx < minX) minX = cx;
                    if (cx > maxX) maxX = cx;
                    if (cy < minY) minY = cy;
                    if (cy > maxY) maxY = cy;
                } catch (eB) {
                    // bounds 取得できないものは無視
                }
            }

            if (minX > maxX || minY > maxY) {
                return { orientation: "unknown", dx: 0, dy: 0, ratio: 0 };
            }

            var dx = Math.abs(maxX - minX);
            var dy = Math.abs(maxY - minY);

            // ratio = 大きい方 / 小さい方（0除算回避）
            var ratio = 0;
            if (dx === 0 && dy === 0) {
                ratio = 0;
            } else if (dx === 0) {
                ratio = 1e12;
            } else if (dy === 0) {
                ratio = 1e12;
            } else {
                ratio = (dx > dy) ? (dx / dy) : (dy / dx);
            }

            // ここでは「推定」だけ。判定閾値・例外処理などの最終ロジックは後で詰める。
            var orientation = "mixed";
            if (dx > dy) orientation = "horizontal";
            else if (dy > dx) orientation = "vertical";
            else orientation = "mixed";

            return { orientation: orientation, dx: dx, dy: dy, ratio: ratio };
        } catch (e) {
            return { orientation: "unknown", dx: 0, dy: 0, ratio: 0 };
        }
    }

    // Separate gradient boundary prompt (0-100). Returns Number or null if canceled/invalid.
    function promptSeparateBoundary(defaultVal) {
        try {
            var defStr = (defaultVal != null) ? String(defaultVal) : "50";
            var input = prompt((lang === 'ja') ? "境界位置（0〜100）を入力してください" : "Enter boundary position (0-100)", defStr);
            if (input === null) return null;
            var p = parseFloat(input);
            if (isNaN(p)) return null;
            // clamp 0..100
            if (p < 0) p = 0;
            if (p > 100) p = 100;
            return p;
        } catch (e) {
            return null;
        }
    }

    // 色＋位置のエントリを収集 / Collect color entries with position
    function collectFillColorEntries(item, outEntries) {
        if (!item) return;

        try {
            // グループなどは再帰
            if (item.typename === "GroupItem") {
                for (var i = 0; i < item.pageItems.length; i++) {
                    collectFillColorEntries(item.pageItems[i], outEntries);
                }
                return;
            }

            // compoundPath は pathItems を辿る
            if (item.typename === "CompoundPathItem") {
                for (var j = 0; j < item.pathItems.length; j++) {
                    collectFillColorEntries(item.pathItems[j], outEntries);
                }
                return;
            }

            // テキスト
            if (item.typename === "TextFrame") {
                var pT = getItemTopLeft(item);

                // Fill
                var tf = item.textRange.characterAttributes.fillColor;
                if (!isNoColor(tf)) {
                    outEntries.push({ left: pT.left, top: pT.top, color: tf });
                }

                // Stroke
                try {
                    var ts = item.textRange.characterAttributes.strokeColor;
                    if (!isNoColor(ts)) {
                        outEntries.push({ left: pT.left, top: pT.top, color: ts });
                    }
                } catch (eTS) { }

                return;
            }

            // PathItem など
            if (typeof item.filled !== "undefined" || typeof item.stroked !== "undefined") {
                var p = getItemTopLeft(item);

                // Fill
                if (typeof item.filled !== "undefined" && item.filled) {
                    var fc = item.fillColor;
                    if (!isNoColor(fc)) {
                        outEntries.push({ left: p.left, top: p.top, color: fc });
                    }
                }

                // Stroke
                if (typeof item.stroked !== "undefined" && item.stroked) {
                    var sc = item.strokeColor;
                    if (!isNoColor(sc)) {
                        outEntries.push({ left: p.left, top: p.top, color: sc });
                    }
                }

                return;
            }
        } catch (e) {
            // 取得できないアイテムは無視
        }
    }

    // 選択から色を「左→右、上→下」順で収集（重複は除外）
    function collectColorsFromSelection(selection) {
        var entries = [];
        for (var i = 0; i < selection.length; i++) {
            collectFillColorEntries(selection[i], entries);
        }

        // 左→右（left 昇順）、上→下（top 降順）でソート
        entries.sort(function (a, b) {
            if (a.left < b.left) return -1;
            if (a.left > b.left) return 1;
            // top は上ほど値が大きい（座標系の都合）ため降順
            if (a.top > b.top) return -1;
            if (a.top < b.top) return 1;
            return 0;
        });

        // 重複除外（同一色は最初の1つだけ）
        var colors = [];
        var seen = {};
        for (var k = 0; k < entries.length; k++) {
            var c = entries[k].color;
            if (isNoColor(c)) continue;
            var key = colorKey(c);
            if (seen[key]) continue;
            seen[key] = true;
            colors.push(c);
        }

        return colors;
    }


    // Run Graphic Style creation via Action (ai_plugin_styles) / アクションでグラフィックスタイル作成
    function runGraphicStyleAction() {
        try {
            var CR = String.fromCharCode(13);
            var actionCode = [
                " /version 3",
                "/name [ 12",
                "\t477261706869635374796c65",
                "]",
                "/isOpen 1",
                "/actionCount 1",
                "/action-1 {",
                "\t/name [ 3",
                "\t\t6e6577",
                "\t]",
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (ai_plugin_styles)",
                "\t\t/localizedName [ 30",
                "\t\t\te382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab",
                "\t\t]",
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 1",
                "\t\t/showDialog 0",
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key 1835363957",
                "\t\t\t/showInPalette 4294967295",
                "\t\t\t/type (enumerated)",
                "\t\t\t/name [ 36",
                "\t\t\t\te696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382",
                "\t\t\t\ta4e383ab",
                "\t\t\t]",
                "\t\t\t/value 1",
                "\t\t}",
                "\t}",
                "}",
                ""
            ].join(CR);

            var actionSetName = "GraphicStyle";
            var actionName = "new";
            var tempFile = new File(Folder.temp + "/temp_graphicstyle.aia");
            tempFile.open("w");
            tempFile.write(actionCode);
            tempFile.close();

            app.loadAction(tempFile);
            app.doScript(actionName, actionSetName);
            app.unloadAction(actionSetName, "");
            try { tempFile.remove(); } catch (eDel) { }
        } catch (e) {
            try { app.unloadAction("GraphicStyle", ""); } catch (e2) { }
        }
    }

    // Register the selected item's appearance as a new Graphic Style / 選択アイテムの見た目をグラフィックスタイルとして登録
    function registerGraphicStyleFromSelected(baseName) {
        try {
            if (!doc.graphicStyles) return;

            // 参照ロジック: 1つだけ選択 → アクション経由で新規スタイル作成
            var selectedItems = [];
            try {
                if (doc.selection && doc.selection.length) {
                    for (var i = 0; i < doc.selection.length; i++) selectedItems.push(doc.selection[i]);
                }
            } catch (eSel) { }

            if (!selectedItems.length) return;

            for (var i = 0; i < selectedItems.length; i++) {
                try {
                    var beforeLen = 0;
                    try { beforeLen = doc.graphicStyles.length; } catch (eLen) { beforeLen = 0; }

                    // 1つだけ選択
                    try {
                        doc.selection = null;
                    } catch (eClr) { }
                    try {
                        doc.selection = [selectedItems[i]];
                    } catch (eSetSel) {
                        // フォールバック
                        try {
                            selectedItems[i].selected = true;
                        } catch (eSel2) { }
                    }

                    // 新規グラフィックスタイル（アクション経由）
                    runGraphicStyleAction();

                    var afterLen = 0;
                    try { afterLen = doc.graphicStyles.length; } catch (eLen2) { afterLen = beforeLen; }
                    if (afterLen <= 0) continue;

                    // 追加された（はずの）末尾を取得
                    var gs = null;
                    try {
                        gs = doc.graphicStyles[afterLen - 1];
                    } catch (eGs) {
                        gs = null;
                    }
                    if (!gs) continue;

                    // グラフィックスタイル名の設定は不要（既定名のまま）
                } catch (eEach) {
                    // 1件失敗しても続行
                }
            }
        } catch (e) {
            // silent
        }
    }

    function uniqueName(baseName, existsFunc) {
        var name = baseName;
        var n = 1;
        while (true) {
            try {
                if (existsFunc(name)) {
                    name = baseName + " " + n;
                    n++;
                    continue;
                }
                break;
            } catch (e) {
                break;
            }
        }
        return name;
    }

    function swatchExists(name) {
        try {
            doc.swatches.getByName(name);
            return true;
        } catch (e) {
            return false;
        }
    }

    function swatchGroupExists(name) {
        try {
            doc.swatchGroups.getByName(name);
            return true;
        } catch (e) {
            return false;
        }
    }

    // カラーを「グローバルカラー（プロセス）」に変換して返す
    function toGlobalProcessColor(doc, baseColor, baseName) {
        try {
            var spot = doc.spots.add();
            spot.name = baseName;
            spot.colorType = ColorModel.PROCESS; // グローバル（プロセス）
            spot.color = baseColor;

            var sc = new SpotColor();
            sc.spot = spot;
            sc.tint = 100;
            return sc;
        } catch (e) {
            // 失敗時は元のカラーを返す（無言）
            return baseColor;
        }
    }

    function addSwatchForColor(colorObj, baseName, makeGlobal) {
        var s = doc.swatches.add();
        var nm = uniqueName(baseName, swatchExists);
        s.name = nm;

        // グローバルカラー（プロセス）に変換して登録（オプション）
        if (makeGlobal) {
            var globalColor = toGlobalProcessColor(doc, colorObj, nm);
            s.color = globalColor;
        } else {
            s.color = colorObj;
        }

        try { s.selected = false; } catch (e) { }
        return s;
    }

    function getUnlockedVisibleLayer(doc) {
        try {
            if (doc.activeLayer && !doc.activeLayer.locked && doc.activeLayer.visible) return doc.activeLayer;
        } catch (e) { }
        for (var i = 0; i < doc.layers.length; i++) {
            try {
                if (!doc.layers[i].locked && doc.layers[i].visible) return doc.layers[i];
            } catch (e2) { }
        }
        return null;
    }

    // 選択オブジェクトから色を抽出
    var colors = collectColorsFromSelection(doc.selection);

    if (colors.length < 2) {
        return;
    }

    try {
        // 新規スウォッチグループを作成（重複回避）
        var baseGroupName = "AutoGradient";
        var groupName = uniqueName(baseGroupName, swatchGroupExists);
        var swGroup = doc.swatchGroups.add();
        swGroup.name = groupName;

        // 抽出色をスウォッチに登録（順番は選択の走査順）
        var createdSwatches = [];
        for (var i = 0; i < colors.length; i++) {
            var cs = addSwatchForColor(colors[i], "AutoColor", opts.makeGlobal);
            createdSwatches.push(cs);
            try { swGroup.addSwatch(cs); } catch (eAdd1) { }
        }

        // オブジェクトの選択解除（以降の処理は選択に依存しない）
        try {
            doc.selection = null;
        } catch (eSelClear) { }

        var newGradient = null;
        if (opts.makeGradient) {
            // Separate gradient (auto split, 2..4 colors)
            if (opts.separateGradient && (colors.length === 2 || colors.length === 3 || colors.length === 4)) {
                var n = colors.length;
                var d = 0.01;

                // step = 100 / n. For n=3, round to 1 decimal to match 33.3 style.
                var step = 100 / n;
                if (n === 3) step = Math.round(step * 10) / 10; // 33.3

                // Build stop points/colors for hard edges: total stops = 2n
                var stopPoints = [];
                var stopColors = [];

                function pickColor(idx) {
                    try {
                        if (createdSwatches && createdSwatches[idx] && createdSwatches[idx].color) return createdSwatches[idx].color;
                    } catch (_) { }
                    return colors[idx];
                }

                // Start
                stopPoints.push(0);
                stopColors.push(pickColor(0));

                // Boundaries: p = step*k
                for (var k = 1; k <= n - 1; k++) {
                    var p = step * k;
                    if (n === 3) p = Math.round(p * 10) / 10; // 33.3, 66.7

                    var p1 = p - d;
                    var p2 = p + d;

                    // Clamp safety
                    if (p1 < 0) p1 = 0;
                    if (p2 > 100) p2 = 100;

                    // left side keeps previous color, right side switches to next color
                    stopPoints.push(p1);
                    stopColors.push(pickColor(k - 1));

                    stopPoints.push(p2);
                    stopColors.push(pickColor(k));
                }

                // End
                stopPoints.push(100);
                stopColors.push(pickColor(n - 1));

                // Ensure exact stop count
                newGradient = doc.gradients.add();
                newGradient.type = GradientType.LINEAR;
                while (newGradient.gradientStops.length < stopPoints.length) newGradient.gradientStops.add();
                while (newGradient.gradientStops.length > stopPoints.length) newGradient.gradientStops[newGradient.gradientStops.length - 1].remove();

                for (var j = 0; j < stopPoints.length; j++) {
                    var stop = newGradient.gradientStops[j];
                    stop.rampPoint = stopPoints[j];
                    try {
                        stop.color = stopColors[j];
                    } catch (eStopColor) {
                        // Fallback: best-effort
                        try { stop.color = colors[Math.min(colors.length - 1, Math.max(0, Math.floor(j / 2)))]; } catch (e2) { }
                    }
                    stop.midPoint = 50;
                    stop.opacity = 100;
                }

                // Do not assign a unique name for separate gradients (keep default)
            } else {
                // Normal gradient logic
                newGradient = doc.gradients.add();
                newGradient.type = GradientType.LINEAR; // 線形グラデーション（必要に応じてRADIALに変更可）

                // ストップ数を抽出色数に合わせる
                while (newGradient.gradientStops.length < colors.length) {
                    newGradient.gradientStops.add();
                }
                while (newGradient.gradientStops.length > colors.length) {
                    newGradient.gradientStops[newGradient.gradientStops.length - 1].remove();
                }

                // 抽出色をグラデーションストップに適用
                for (var j = 0; j < colors.length; j++) {
                    var stop = newGradient.gradientStops[j];

                    // 位置（RampPoint）を計算 (0 〜 100)
                    var location = (j / (colors.length - 1)) * 100;
                    stop.rampPoint = location;

                    // 色を適用（スウォッチ登録時に作成したグローバルカラーを優先）
                    try {
                        if (createdSwatches && createdSwatches[j] && createdSwatches[j].color) {
                            stop.color = createdSwatches[j].color; // SpotColor（グローバル）を渡す
                        } else {
                            stop.color = colors[j];
                        }
                    } catch (eStopColor) {
                        // 無言フォールバック
                        try { stop.color = colors[j]; } catch (e2) { }
                    }

                    // 中間点（MidPoint）をデフォルトの50に設定
                    stop.midPoint = 50;

                    // 不透明度
                    stop.opacity = 100;
                }
                // グラデーション名（重複回避）: only for normal gradients or non-2color separate
                if (!(opts.separateGradient && colors.length === 2)) {
                    var baseGradientName = "New Gradient";
                    var gradientName = uniqueName(baseGradientName, function (nm) {
                        try {
                            doc.gradients.getByName(nm);
                            return true;
                        } catch (e) {
                            return false;
                        }
                    });
                    newGradient.name = gradientName;
                }
            }
        }

        // 参考: ビュー中心に長方形を作成して、作成したグラデーションを適用（またはスタイル登録用の一時長方形）
        if ((opts.makeRect || opts.registerGraphicStyle) && newGradient) {
            try {
                var targetLayer = getUnlockedVisibleLayer(doc);
                if (targetLayer) {
                    // Save state for temporary rectangle flow / 一時長方形用に状態を退避
                    var tempOnly = (!opts.makeRect && opts.registerGraphicStyle);
                    var prevSelection = null;
                    var prevActiveLayer = null;
                    var tempLayer = null;
                    try { prevActiveLayer = doc.activeLayer; } catch (eAL) { prevActiveLayer = null; }
                    try {
                        if (doc.selection && doc.selection.length) {
                            prevSelection = [];
                            for (var ps = 0; ps < doc.selection.length; ps++) prevSelection.push(doc.selection[ps]);
                        }
                    } catch (ePS) { prevSelection = null; }

                    // If tempOnly, create a dedicated temp layer and draw there / 一時用レイヤーを作成してそこに作る
                    if (tempOnly) {
                        try {
                            tempLayer = doc.layers.add();
                            tempLayer.name = "__TempGraphicStyle";
                            tempLayer.visible = true;
                            tempLayer.locked = false;
                            try { doc.activeLayer = tempLayer; } catch (eSetAL) { }
                        } catch (eTL) {
                            tempLayer = null;
                        }
                    }

                    // Decide drawing layer: normal uses targetLayer, tempOnly uses tempLayer if available
                    var drawLayer = (tempOnly && tempLayer) ? tempLayer : targetLayer;

                    var viewCenterX = doc.activeView.centerPoint[0];
                    var viewCenterY = doc.activeView.centerPoint[1];

                    // 長方形サイズ：選択の幅・高さを流用（取得できない/無効なら 100）。オプションOFF時は 100。
                    var rectWidth = 100;
                    var rectHeight = 100;
                    if (opts.useSelectionSize) {
                        try {
                            if (selBounds) {
                                rectWidth = Math.abs(selBounds.right - selBounds.left);
                                rectHeight = Math.abs(selBounds.top - selBounds.bottom);
                            }
                        } catch (eSize) { }
                    }

                    // 念のため最小サイズ
                    if (!rectWidth || rectWidth <= 0) rectWidth = 100;
                    if (!rectHeight || rectHeight <= 0) rectHeight = 100;

                    // 位置：基本はビュー中心。選択が横/縦並びなら選択に沿って配置。
                    var rectTop = viewCenterY + rectHeight / 2;
                    var rectLeft = viewCenterX - rectWidth / 2;

                    try {
                        if (selBounds && selOri && (selOri.orientation === "horizontal" || selOri.orientation === "vertical")) {
                            var selW = Math.abs(selBounds.right - selBounds.left);
                            var selH = Math.abs(selBounds.top - selBounds.bottom);

                            if (selOri.orientation === "horizontal") {
                                // 左：選択の左に合わせる／上：選択の下に「長方形1個分」離して配置
                                rectLeft = selBounds.left;
                                rectTop = selBounds.bottom - rectHeight;
                            } else if (selOri.orientation === "vertical") {
                                // 上：選択の上に合わせる／左：選択の右に「長方形1個分」離して配置
                                rectLeft = selBounds.right + rectWidth;
                                rectTop = selBounds.top;
                            }
                        }
                    } catch (ePos) { }

                    var rect = drawLayer.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);

                    // 作成した長方形を選択状態にする
                    try {
                        doc.selection = null; // 念のためクリア
                        rect.selected = true;
                        // 元の選択が縦並びなら、グラデーション角度を90度に
                        try {
                            if (selOri && selOri.orientation === "vertical") {
                                runGradientAngle90Action();
                            }
                        } catch (eAct) { }
                    } catch (eSelRect) { }

                    rect.stroked = false;
                    rect.filled = true;

                    var gc = new GradientColor();
                    gc.gradient = newGradient;
                    rect.fillColor = gc;

                    // Optionally register as Graphic Style / （任意）グラフィックスタイルに登録
                    if (opts.registerGraphicStyle) {
                        try {
                            // Ensure the rectangle is selected for the action
                            doc.selection = null;
                            rect.selected = true;
                            registerGraphicStyleFromSelected('AutoStyle');
                        } catch (eGS) { }
                    }

                    // If rectangle output is OFF but style registration is ON, delete the temporary rectangle and clean up
                    if (tempOnly) {
                        try { rect.remove(); } catch (eRm) { }

                        // Remove temp layer if we created it
                        if (tempLayer) {
                            try { tempLayer.remove(); } catch (eLayRm) { }
                        }

                        // Restore active layer
                        try {
                            if (prevActiveLayer) doc.activeLayer = prevActiveLayer;
                        } catch (eRestoreAL) { }

                        // Restore selection (best-effort)
                        try {
                            doc.selection = null;
                            if (prevSelection && prevSelection.length) {
                                doc.selection = prevSelection;
                            }
                        } catch (eRestoreSel) { }
                    }
                }
            } catch (eRect) { }
        }

        // 最後に追加されたスウォッチ（= 作成したグラデーション）を選択（グラデーション作成時のみ）
        if (newGradient) {
            try {
                var idx = doc.swatches.length - 1;
                if (idx >= 0) {
                    doc.swatches[idx].selected = true;
                }
            } catch (e) { }
        }


    } catch (e) {
        // 無言（エラー通知しない）
    }
}

main();
