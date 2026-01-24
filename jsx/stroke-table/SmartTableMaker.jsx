#target illustrator
#targetengine "SmartTableMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
SmartTableMaker

### GitHub：
https://github.com/swwwitch/illustrator-scripts

### 概要：
- 選択したテキストフレーム群から行・列を推定し、表の罫線を生成します。
- ダイアログで線（線幅／横ケイのみ／外枠は長方形／1行目を見出し行に）と、塗り（なし／全体／行ごと／セル）を指定できます。

### 主な機能：
- 行（Y）と列（X）のクラスタリングにより、表構造を推定
- 罫線（外枠＋内側の縦罫・横罫）をパスとして生成
- 生成物を任意のグループ名で新規グループにまとめる

### 処理の流れ：
1) 選択から対応オブジェクトを抽出（必要に応じてテキストを一時アウトライン化して計測）
2) 行クラスタリング（Y）→ 列推定（X）
3) visibleBounds を優先してセル外接矩形（padding込み）を求め、行・列の境界を決定
4) 罫線（外枠・内側）を作成

### Script Name:
SmartTableMaker

### GitHub:
https://github.com/swwwitch/illustrator-scripts

### Summary:
- Estimates a table grid from selected text frames and draws table lines.
- Provides a dialog for line options (stroke width / horizontal-only / outer frame / header row) and fill options (None / All / By row / By cell).

### Key Features:
- Row (Y) clustering and column (X) estimation
- Draws table lines (outer frame + inner vertical/horizontal lines)
- Optionally groups created paths under a named group
*/

(function () {
    var SCRIPT_VERSION = "v1.0";

    /* ローカライズ / Localization */

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "表組み作成", en: "Smart Table Maker" },

        panelLine: { ja: "線", en: "Lines" },
        cbLines: { ja: "線", en: "Lines" },
        stLineWidth: { ja: "線幅", en: "Stroke" },
        cbHLineOnly: { ja: "横ケイのみ", en: "Horizontal only" },
        cbOuterRect: { ja: "外枠は長方形", en: "Outer frame" },
        cbHeaderRow: { ja: "1行目を見出し行に", en: "Header row (1st)" },

        panelFill: { ja: "塗り", en: "Fill" },
        rbFillNone: { ja: "なし", en: "None" },
        rbFillAll: { ja: "全体", en: "All" },
        rbFillByRow: { ja: "行ごと", en: "By row" },
        rbFillByCell: { ja: "セル", en: "By cell" },
        cbFillHeader: { ja: "見出し対応", en: "Header-aware" },

        cbOutlineMeasure: { ja: "テキストをアウトライン化して計算", en: "Measure text by outlining" },

        btnCancel: { ja: "キャンセル", en: "Cancel" },
        btnOK: { ja: "OK", en: "OK" }
    };

    function L(key) {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].ja) return LABELS[key].ja;
        return String(key);
    }

    if (app.documents.length === 0) { alert("ドキュメントがありません"); return; }
    var doc = app.activeDocument;

    var sel = doc.selection;
    if (!sel || sel.length === 0) { alert("オブジェクトを選択してください"); return; }

    // --- Unit utilities (strokeUnits) ---
    // 単位コード → ラベル（Q/H切替対応）
    var unitMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    function getUnitLabel(code, prefKey) {
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return unitMap[code] || "pt";
    }

    function getStrokeUnitLabel() {
        var code = 2;
        try { code = app.preferences.getIntegerPreference("strokeUnits"); } catch (e) { code = 2; }
        return getUnitLabel(code, "strokeUnits");
    }

    // 設定（必要なら調整）
    var opts = {
        rowTolerancePt: 8,    // 同じ行とみなすY差（pt）
        colTolerancePt: 8,    // 同じ列とみなすX差（pt）
        paddingPt: 6,         // セル矩形に足す余白
        strokeWidthPt: 0.25,
        makeInNewGroup: true,
        groupName: "TABLE_LINES"
    };

    // --- Session Settings (Illustrator再起動で消える：targetengine + $.global) ---
    var SESSION_KEY = "__SmartTableMaker_SessionSettings__";

    function getSessionSettings() {
        try {
            if ($.global && $.global[SESSION_KEY]) return $.global[SESSION_KEY];
        } catch (e) { }
        return null;
    }

    function setSessionSettings(obj) {
        try {
            if (!$.global) $.global = {};
            $.global[SESSION_KEY] = obj;
        } catch (e) { }
    }

    // --- Dialog ---
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);

    // ダイアログ位置／透明度の調整
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        var currentX = dlg.location[0];
        var currentY = dlg.location[1];
        dlg.location = [currentX + offsetX, currentY + offsetY];
    }

    function setDialogOpacity(dlg, opacityValue) {
        try { dlg.opacity = opacityValue; } catch (e) { }
    }

    setDialogOpacity(dlg, dialogOpacity);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    // 2-column layout: left = Line, right = Fill
    var cols = dlg.add('group');
    cols.orientation = 'row';
    cols.alignChildren = ['fill', 'top'];
    cols.alignment = ['fill', 'top'];

    var panelMode = cols.add('panel', undefined, L('panelLine'));
    panelMode.orientation = 'column';
    panelMode.alignChildren = ['left', 'top'];
    panelMode.margins = [15, 20, 15, 10];

    var cbLines = panelMode.add('checkbox', undefined, L('cbLines'));

    var gLineWidth = panelMode.add('group');
    gLineWidth.orientation = 'row';
    gLineWidth.alignChildren = ['left', 'center'];

    var stLineWidth = gLineWidth.add('statictext', undefined, L('stLineWidth'));
    var etLineWidth = gLineWidth.add('edittext', undefined, String(opts.strokeWidthPt));
    etLineWidth.characters = 6;

    changeValueByArrowKey(etLineWidth);

    etLineWidth.onChange = function () {
        var v = Number(etLineWidth.text);
        if (!isFinite(v) || v < 0) v = 0;
        etLineWidth.text = formatOneDecimal(v);
    };

    // ダイアログ表示時：位置調整 + 線幅へフォーカス
    dlg.onShow = function () {
        try {
            shiftDialogPosition(dlg, offsetX, offsetY);
        } catch (e0) { }

        try {
            stLineUnit.text = getStrokeUnitLabel();
        } catch (eUnit) { }

        try {
            etLineWidth.active = true;
            etLineWidth.selection = [0, etLineWidth.text.length]; // 全選択
        } catch (e) { }
    };

    var stLineUnit = gLineWidth.add('statictext', undefined, getStrokeUnitLabel());

    var cbHLineOnly = panelMode.add('checkbox', undefined, L('cbHLineOnly'));

    var gLineOptions = panelMode.add('group');
    gLineOptions.orientation = 'column';
    gLineOptions.alignChildren = ['left', 'top'];

    var cbOuterRect = gLineOptions.add('checkbox', undefined, L('cbOuterRect'));
    var cbHeaderRow = gLineOptions.add('checkbox', undefined, L('cbHeaderRow'));


    // --- Fill Panel (UI only; logic TBD) ---
    var panelFill = cols.add('panel', undefined, L('panelFill'));
    panelFill.orientation = 'column';
    panelFill.alignChildren = ['left', 'top'];
    panelFill.margins = [15, 20, 15, 10];

    var rbFillNone = panelFill.add('radiobutton', undefined, L('rbFillNone'));
    var rbFillAll = panelFill.add('radiobutton', undefined, L('rbFillAll'));
    var rbFillByRow = panelFill.add('radiobutton', undefined, L('rbFillByRow'));
    var rbFillByCell = panelFill.add('radiobutton', undefined, L('rbFillByCell'));

    var cbFillHeader = panelFill.add('checkbox', undefined, L('cbFillHeader'));

    // Spanning options (below 2 columns) - centered
    var gSpanning = dlg.add('group');
    gSpanning.orientation = 'row';
    gSpanning.alignChildren = ['fill', 'center'];
    gSpanning.alignment = ['fill', 'top'];

    // left / right spacers to center the checkbox
    var spL = gSpanning.add('group');
    spL.alignment = ['fill', 'center'];

    var cbOutlineMeasure = gSpanning.add('checkbox', undefined, L('cbOutlineMeasure'));
    cbOutlineMeasure.alignment = ['center', 'center'];

    var spR = gSpanning.add('group');
    spR.alignment = ['fill', 'center'];

    // default (restore previous values within the same Illustrator session)
    var ss = getSessionSettings() || {};

    // line
    cbLines.value = (ss.cbLines !== undefined) ? ss.cbLines : true;
    cbHLineOnly.value = (ss.cbHLineOnly !== undefined) ? ss.cbHLineOnly : false;
    cbOuterRect.value = (ss.cbOuterRect !== undefined) ? ss.cbOuterRect : false;
    cbHeaderRow.value = (ss.cbHeaderRow !== undefined) ? ss.cbHeaderRow : false;

    // stroke width
    var sw = (ss.strokeWidthPt !== undefined) ? ss.strokeWidthPt : opts.strokeWidthPt;
    if (!isFinite(sw) || sw <= 0) sw = opts.strokeWidthPt;
    etLineWidth.text = formatOneDecimal(sw);

    // outline-measure
    cbOutlineMeasure.value = (ss.cbOutlineMeasure !== undefined) ? ss.cbOutlineMeasure : false;

    // fill radio
    var fm = ss.fillMode || 'none';
    rbFillNone.value = (fm === 'none');
    rbFillAll.value = (fm === 'all');
    rbFillByRow.value = (fm === 'row');
    rbFillByCell.value = (fm === 'cell');

    // fill header-aware
    cbFillHeader.value = (ss.cbFillHeader !== undefined) ? ss.cbFillHeader : false;

    function updateLineOptionEnables() {
        // 「線」OFF のときは、それ以外すべてをディム
        if (!cbLines.value) {
            gLineWidth.enabled = false;
            cbHLineOnly.enabled = false;
            gLineOptions.enabled = false;
            return;
        }

        // 「線」ON のときは、線幅・横ケイ・見出しは有効
        gLineWidth.enabled = true;
        cbHLineOnly.enabled = true;
        gLineOptions.enabled = true;

        // 「横ケイのみ」ON のときは「外枠は長方形」をディム
        cbOuterRect.enabled = !cbHLineOnly.value;

        // 無効時にチェックが残らないように（取り残し防止）
        if (!cbOuterRect.enabled) cbOuterRect.value = false;
    }

    function updateFillOptionEnables() {
        // 「見出し対応」は「行ごと」「セル」のときだけ有効
        var enableHeader = rbFillByRow.value || rbFillByCell.value;
        cbFillHeader.enabled = enableHeader;
        if (!enableHeader) cbFillHeader.value = false;
    }

    function formatOneDecimal(n) {
        if (!isFinite(n)) return '';
        // 常に小数第1位まで表示（例：1 -> "1.0"）
        return (Math.round(n * 10) / 10).toFixed(1);
    }

    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
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
                    if (editText === etLineWidth) {
                        // 小数が入っていても 1pt 刻みでスナップ
                        value = Math.floor(value) + 1;
                    } else {
                        value += delta;
                    }
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    if (editText === etLineWidth) {
                        // 小数が入っていても 1pt 刻みでスナップ
                        value = Math.ceil(value) - 1;
                    } else {
                        value -= delta;
                    }
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め（ただし線幅は小数を保持）
                if (editText === etLineWidth) {
                    value = Math.round(value * 100) / 100;
                } else {
                    value = Math.round(value);
                }
            }

            // 0未満防止（念のため）
            if (value < 0) value = 0;

            if (editText === etLineWidth) {
                editText.text = formatOneDecimal(value);
            } else {
                editText.text = value;
            }
        });
    }

    cbLines.onClick = updateLineOptionEnables;
    cbHLineOnly.onClick = updateLineOptionEnables;

    rbFillNone.onClick = updateFillOptionEnables;
    rbFillAll.onClick = updateFillOptionEnables;
    rbFillByRow.onClick = updateFillOptionEnables;
    rbFillByCell.onClick = updateFillOptionEnables;

    // initial state
    updateLineOptionEnables();
    updateFillOptionEnables();

    var btns = dlg.add('group');
    btns.alignment = 'right';
    var btnCancel = btns.add('button', undefined, L('btnCancel'), { name: 'cancel' });
    var btnOK = btns.add('button', undefined, L('btnOK'), { name: 'ok' });

    var r = dlg.show();

    // Save last-closed values (session only; restored on next run)
    (function saveDialogState() {
        var sw2 = Number(etLineWidth.text);
        if (!isFinite(sw2) || sw2 < 0) sw2 = opts.strokeWidthPt;

        var fm2 = 'none';
        if (rbFillAll.value) fm2 = 'all';
        else if (rbFillByRow.value) fm2 = 'row';
        else if (rbFillByCell.value) fm2 = 'cell';

        setSessionSettings({
            strokeWidthPt: sw2,
            cbLines: !!cbLines.value,
            cbHLineOnly: !!cbHLineOnly.value,
            cbOuterRect: !!cbOuterRect.value,
            cbHeaderRow: !!cbHeaderRow.value,
            cbOutlineMeasure: !!cbOutlineMeasure.value,
            fillMode: fm2,
            cbFillHeader: !!cbFillHeader.value
        });
    })();

    if (r !== 1) return;

    var mode = 'linesOnly';

    // Fill mode (UI only; logic TBD)
    var fillMode = 'none';
    if (rbFillAll.value) fillMode = 'all';
    else if (rbFillByRow.value) fillMode = 'row';
    else if (rbFillByCell.value) fillMode = 'cell';

    // Fill header-aware (UI only; logic TBD)
    var fillHeaderAware = !!cbFillHeader.value;

    // Outline-measure (UI only; logic TBD)
    var outlineMeasure = !!cbOutlineMeasure.value;

    // base mode by controls
    if (cbHLineOnly.value) {
        mode = 'hLineOnly';
    } else {
        // 線
        mode = cbOuterRect.value ? 'outerRect' : 'linesOnly';
    }

    // --- Main ---

    function makeKColor(k) {
        var col = new CMYKColor();
        col.cyan = 0;
        col.magenta = 0;
        col.yellow = 0;
        col.black = k;
        return col;
    }

    // 対象オブジェクトを抽出（テキスト／パス／グループ等）
    // ※ geometricBounds を持つ PageItem を対象とする
    function isSupportedItem(it) {
        if (!it) return false;
        var t = it.typename;
        return (
            t === "TextFrame" ||
            t === "PathItem" ||
            t === "GroupItem" ||
            t === "CompoundPathItem" ||
            t === "PlacedItem" ||
            t === "RasterItem" ||
            t === "MeshItem" ||
            t === "PluginItem"
        );
    }


    // テキストは一時的にアウトライン化して外接を計測（元オブジェクトは変更しない）
    function getTextBoundsByOutlining(tf) {
        // 失敗時は null を返してフォールバックする
        var tmpDup = null;
        var outlined = null;
        try {
            // 複製→アウトライン化→bounds取得→削除
            tmpDup = tf.duplicate(doc.activeLayer, ElementPlacement.PLACEATEND);
            outlined = tmpDup.createOutline();
            // createOutline後に複製が残る場合があるので消す
            try { tmpDup.remove(); } catch (e1) { }

            var b = null;
            try {
                // visibleBounds の方が見た目に一致しやすい（geometricBounds のズレ対策）
                b = outlined.visibleBounds;
            } catch (e2) {
                try { b = outlined.geometricBounds; } catch (e3) { b = null; }
            }

            // 一時アウトラインを削除
            try { outlined.remove(); } catch (e4) { }

            return b;
        } catch (e) {
            // cleanup
            try { if (outlined) outlined.remove(); } catch (e5) { }
            try { if (tmpDup) tmpDup.remove(); } catch (e6) { }
            return null;
        }
    }

    // bounds: [left, top, right, bottom]
    function gb(it) {
        try {
            if (it && it.typename === "TextFrame") {
                // チェックONのときだけ一時アウトライン化して計測
                if (outlineMeasure) {
                    var bOutline = getTextBoundsByOutlining(it);
                    if (bOutline) return bOutline;
                }
                // OFF時、またはアウトライン計測失敗時は bounds で計測
                try { return it.visibleBounds; } catch (e1) { }
                return it.geometricBounds;
            }

            // TextFrame 以外も、見た目に近い visibleBounds を優先
            try { return it.visibleBounds; } catch (e2) { }
            return it.geometricBounds;
        } catch (e) {
            return null;
        }
    }

    var items = [];
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (!isSupportedItem(it)) continue;

        var b = gb(it);
        if (!b) continue;

        // 念のため異常値を弾く
        if (!isFinite(b[0]) || !isFinite(b[1]) || !isFinite(b[2]) || !isFinite(b[3])) continue;

        items.push(it);
    }

    if (items.length === 0) {
        alert("選択内に対応オブジェクトがありません（テキスト／パス／グループ等を選択してください）");
        return;
    }

    // 行クラスタリング（Yはtop基準。Illustratorは上が大きい）
    function clusterByY(frames, tol) {
        // topでソート（降順：上から下へ）
        frames.sort(function (a, b) { return gb(b)[1] - gb(a)[1]; });

        var rows = [];
        for (var i = 0; i < frames.length; i++) {
            var f = frames[i];
            var t = gb(f)[1];
            var placed = false;
            for (var r = 0; r < rows.length; r++) {
                if (Math.abs(t - rows[r].y) <= tol) {
                    rows[r].items.push(f);
                    // 平均に寄せる（多少ズレてても安定）
                    rows[r].y = (rows[r].y * (rows[r].items.length - 1) + t) / rows[r].items.length;
                    placed = true;
                    break;
                }
            }
            if (!placed) rows.push({ y: t, items: [f] });
        }

        // 行内はleftで並べる
        for (var r = 0; r < rows.length; r++) {
            rows[r].items.sort(function (a, b) { return gb(a)[0] - gb(b)[0]; });
        }
        return rows;
    }

    // 列位置推定（各行の左端Xを集めてクラスタ）
    function estimateColumns(rows, tol) {
        var xs = [];
        for (var r = 0; r < rows.length; r++) {
            for (var c = 0; c < rows[r].items.length; c++) {
                xs.push(gb(rows[r].items[c])[0]);
            }
        }
        xs.sort(function (a, b) { return a - b; });

        var cols = [];
        for (var i = 0; i < xs.length; i++) {
            var x = xs[i];
            var placed = false;
            for (var k = 0; k < cols.length; k++) {
                if (Math.abs(x - cols[k].x) <= tol) {
                    cols[k].x = (cols[k].x * cols[k].n + x) / (cols[k].n + 1);
                    cols[k].n++;
                    placed = true;
                    break;
                }
            }
            if (!placed) cols.push({ x: x, n: 1 });
        }
        cols.sort(function (a, b) { return a.x - b.x; });
        return cols;
    }

    // 指定行のアイテムを列に割り当て（最も近い列x）
    function assignToColumns(rowItems, cols) {
        var cells = new Array(cols.length);
        for (var i = 0; i < cells.length; i++) cells[i] = null;

        for (var i = 0; i < rowItems.length; i++) {
            var f = rowItems[i];
            var x = gb(f)[0];
            var best = 0;
            var bestD = Math.abs(x - cols[0].x);
            for (var k = 1; k < cols.length; k++) {
                var d = Math.abs(x - cols[k].x);
                if (d < bestD) { bestD = d; best = k; }
            }
            // すでに埋まってたら（イレギュラー）右へずらして空きを探す
            var kk = best;
            while (kk < cells.length && cells[kk] !== null) kk++;
            if (kk >= cells.length) {
                kk = best;
                while (kk >= 0 && cells[kk] !== null) kk--;
            }
            if (kk >= 0) cells[kk] = f;
        }
        return cells;
    }

    // 行・列推定
    var rows = clusterByY(items, opts.rowTolerancePt);
    var cols = estimateColumns(rows, opts.colTolerancePt);

    if (rows.length < 1 || cols.length < 1) { alert("行/列を推定できません"); return; }

    // 各セルの矩形（外接＋padding）を作って列幅・行高を揃えるため集計
    // まずセルごとの bounds を集め、列ごとの左右、行ごとの上下を決める
    var colLeft = new Array(cols.length);
    var colRight = new Array(cols.length);
    var rowTop = new Array(rows.length);
    var rowBottom = new Array(rows.length);

    for (var c = 0; c < cols.length; c++) { colLeft[c] = 1e9; colRight[c] = -1e9; }
    for (var r = 0; r < rows.length; r++) { rowTop[r] = -1e9; rowBottom[r] = 1e9; }

    for (var r = 0; r < rows.length; r++) {
        var cellFrames = assignToColumns(rows[r].items, cols);
        for (var c = 0; c < cols.length; c++) {
            var f = cellFrames[c];
            if (!f) continue;
            var b = gb(f);
            var L = b[0] - opts.paddingPt;
            var T = b[1] + opts.paddingPt;
            var R = b[2] + opts.paddingPt;
            var B = b[3] - opts.paddingPt;

            if (L < colLeft[c]) colLeft[c] = L;
            if (R > colRight[c]) colRight[c] = R;
            if (T > rowTop[r]) rowTop[r] = T;
            if (B < rowBottom[r]) rowBottom[r] = B;
        }
    }

    // 欠損列/行がある場合、近傍から補完（最低限）
    for (var c = 0; c < cols.length; c++) {
        if (colLeft[c] === 1e9 || colRight[c] === -1e9) {
            // 前後の列から推定（雑）
            var prev = c - 1, next = c + 1;
            while (prev >= 0 && (colLeft[prev] === 1e9)) prev--;
            while (next < cols.length && (colLeft[next] === 1e9)) next++;
            if (prev >= 0 && next < cols.length) {
                colLeft[c] = (colLeft[prev] + colLeft[next]) / 2;
                colRight[c] = (colRight[prev] + colRight[next]) / 2;
            } else if (prev >= 0) {
                var w = colRight[prev] - colLeft[prev];
                colLeft[c] = colRight[prev];
                colRight[c] = colLeft[c] + w;
            } else if (next < cols.length) {
                var w2 = colRight[next] - colLeft[next];
                colRight[c] = colLeft[next];
                colLeft[c] = colRight[c] - w2;
            } else {
                colLeft[c] = 0; colRight[c] = 100;
            }
        }
    }
    for (var r = 0; r < rows.length; r++) {
        if (rowTop[r] === -1e9 || rowBottom[r] === 1e9) {
            var prevR = r - 1, nextR = r + 1;
            while (prevR >= 0 && (rowTop[prevR] === -1e9)) prevR--;
            while (nextR < rows.length && (rowTop[nextR] === -1e9)) nextR++;
            if (prevR >= 0 && nextR < rows.length) {
                rowTop[r] = (rowTop[prevR] + rowTop[nextR]) / 2;
                rowBottom[r] = (rowBottom[prevR] + rowBottom[nextR]) / 2;
            } else if (prevR >= 0) {
                var h = rowTop[prevR] - rowBottom[prevR];
                rowTop[r] = rowBottom[prevR];
                rowBottom[r] = rowTop[r] - h;
            } else if (nextR < rows.length) {
                var h2 = rowTop[nextR] - rowBottom[nextR];
                rowBottom[r] = rowTop[nextR];
                rowTop[r] = rowBottom[r] + h2;
            } else {
                rowTop[r] = 0; rowBottom[r] = -50;
            }
        }
    }

    // 罫線作成用の座標：列境界X、行境界Y
    var xEdges = [];
    xEdges.push(colLeft[0]);
    for (var c = 0; c < cols.length; c++) xEdges.push(colRight[c]);

    var yEdges = [];
    yEdges.push(rowTop[0]);
    for (var r = 0; r < rows.length; r++) yEdges.push(rowBottom[r]);

    // 最上段の横罫（yEdges[0]）のみ、少し下にずれるケースへの対策
    // 行高さは概ね同じと仮定し、代表的な行高（中央値）を使って yEdges[0] だけ補正する
    (function adjustTopEdgeOnly() {
        if (!yEdges || yEdges.length < 2) return;
        if (!rows || rows.length < 1) return;

        // 各行の高さ（top-bottom）から代表値（中央値）を取る
        var heights = [];
        for (var i = 0; i < rows.length; i++) {
            var h = rowTop[i] - rowBottom[i];
            if (isFinite(h) && h > 0) heights.push(h);
        }
        if (heights.length === 0) return;
        heights.sort(function (a, b) { return a - b; });
        var mid = Math.floor(heights.length / 2);
        var typicalH = (heights.length % 2 === 1) ? heights[mid] : (heights[mid - 1] + heights[mid]) / 2;
        if (!isFinite(typicalH) || typicalH <= 0) return;

        // 最上段は「2本目-3本目の間隔」を基準に決める
        // yEdges[1] と yEdges[2] の差（= 行高相当）を使って yEdges[0] を算出
        if (yEdges.length >= 3) {
            var hFrom12 = yEdges[1] - yEdges[2];
            if (isFinite(hFrom12) && hFrom12 > 0) {
                yEdges[0] = yEdges[1] + hFrom12;
                return;
            }
        }

        // フォールバック：行高さは概ね同じと仮定し、代表的な行高（中央値）で補正
        yEdges[0] = yEdges[1] + typicalH;
    })();

    // グループ作成
    var parent = doc.activeLayer;
    var grp = parent;
    if (opts.makeInNewGroup) {
        grp = parent.groupItems.add();
        grp.name = opts.groupName;
    }

    // 生成した罫線をあとで選択するために保持
    var createdItems = [];
    // 生成した塗り（背面長方形）をあとで選択するために保持
    var createdFills = [];

    function addLine(x1, y1, x2, y2, strokeWidthPt) {
        var p = grp.pathItems.add();
        p.setEntirePath([[x1, y1], [x2, y2]]);
        p.stroked = true;
        p.filled = false;
        p.strokeWidth = (strokeWidthPt != null) ? strokeWidthPt : opts.strokeWidthPt;
        createdItems.push(p);
        return p;
    }

    function addRect(left, top, right, bottom) {
        // Illustrator: rectangle(top, left, width, height)
        var w = right - left;
        var h = top - bottom;
        var r = grp.pathItems.rectangle(top, left, w, h);
        r.stroked = true;
        r.filled = false;
        r.strokeWidth = (cbHeaderRow.value ? (opts.strokeWidthPt * 2.5) : opts.strokeWidthPt);
        createdItems.push(r);
        return r;
    }

    // 描画領域（最外）
    var left = xEdges[0], right = xEdges[xEdges.length - 1];
    var top = yEdges[0], bottom = yEdges[yEdges.length - 1];

    // 塗り：全体（UI: "全体"）
    // 罫線の背面に、表全体を覆う長方形を描画（K20）
    if (fillMode === 'all') {
        var wFill = right - left;
        var hFill = top - bottom;
        if (wFill > 0 && hFill > 0) {
            // ★作成先を grp → parent（レイヤー）にする
            var fillRect = parent.pathItems.rectangle(top, left, wFill, hFill);
            fillRect.stroked = false;
            fillRect.filled = true;
            fillRect.fillColor = makeKColor(20);

            // テキストの背面へ（レイヤー背面へ送る）
            try {
                fillRect.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (e) { }
            createdFills.push(fillRect);
        }
    }
    else if (fillMode === 'row') {
        // 塗り：行ごと（奇数行 K25 / 偶数行 K10、見出し対応時1行目K40）
        var wFillRow = right - left;
        if (wFillRow > 0 && yEdges && yEdges.length >= 2) {
            for (var rr = 0; rr < yEdges.length - 1; rr++) {
                var tRow = yEdges[rr];
                var bRow = yEdges[rr + 1];
                var hRow = tRow - bRow;
                if (!(hRow > 0)) continue;

                var kVal = (fillHeaderAware && rr === 0) ? 40 : ((rr % 2 === 0) ? 25 : 10); // rr=0 => 1行目
                var rowRect = parent.pathItems.rectangle(tRow, left, wFillRow, hRow);
                rowRect.stroked = false;
                rowRect.filled = true;
                rowRect.fillColor = makeKColor(kVal);

                // テキストの背面へ（レイヤー背面へ送る）
                try {
                    rowRect.zOrder(ZOrderMethod.SENDTOBACK);
                } catch (e) { }
                createdFills.push(rowRect);
            }
        }
    }
    else if (fillMode === 'cell') {
        // 塗り：セル（奇数行 K25 / 偶数行 K10、見出し対応時1行目K40）
        if (xEdges && xEdges.length >= 2 && yEdges && yEdges.length >= 2) {
            for (var rr2 = 0; rr2 < yEdges.length - 1; rr2++) {
                var tCellRow = yEdges[rr2];
                var bCellRow = yEdges[rr2 + 1];
                var hCellRow = tCellRow - bCellRow;
                if (!(hCellRow > 0)) continue;

                var kCell = (fillHeaderAware && rr2 === 0) ? 40 : ((rr2 % 2 === 0) ? 25 : 10); // rr2=0 => 1行目

                for (var cc2 = 0; cc2 < xEdges.length - 1; cc2++) {
                    var lCell = xEdges[cc2];
                    var rCell = xEdges[cc2 + 1];
                    var wCell = rCell - lCell;
                    if (!(wCell > 0)) continue;

                    // ★テキストの背面：parent（アクティブレイヤー）に作成
                    var cellRect = parent.pathItems.rectangle(tCellRow, lCell, wCell, hCellRow);
                    cellRect.stroked = false;
                    cellRect.filled = true;
                    cellRect.fillColor = makeKColor(kCell);

                    // テキストの背面へ（レイヤー背面へ送る）
                    try {
                        cellRect.zOrder(ZOrderMethod.SENDTOBACK);
                    } catch (e) { }
                    createdFills.push(cellRect);
                }
            }
        }
    }

    // 「線」OFF の場合は、線（罫線）を描画しない
    if (cbLines.value) {
        if (mode === 'hLineOnly') {
            // 横ケイのみ：縦線は引かない（外枠の左右も作らない）
            var normalW = opts.strokeWidthPt;
            var thickW = opts.strokeWidthPt * 2.5;
            var lastIdx = yEdges.length - 1;

            for (var j = 0; j < yEdges.length; j++) {
                if (cbHeaderRow.value && (j === 0 || j === 1 || j === lastIdx)) {
                    // 見出しON：上から1本目・2本目・最終行を太く
                    addLine(left, yEdges[j], right, yEdges[j], thickW);
                } else {
                    addLine(left, yEdges[j], right, yEdges[j], normalW);
                }
            }
        } else {
            // 外枠
            if (mode === 'outerRect') {
                // 外枠を1つの長方形パスで作成
                addRect(left, top, right, bottom);
            } else {
                // 現状どおり（外枠も線で作成）
                addLine(left, top, right, top);
                addLine(left, bottom, right, bottom);
                addLine(left, top, left, bottom);
                addLine(right, top, right, bottom);
            }

            // 内側 縦罫
            for (var i = 1; i < xEdges.length - 1; i++) {
                addLine(xEdges[i], top, xEdges[i], bottom);
            }
            // 内側 横罫
            for (var j2 = 1; j2 < yEdges.length - 1; j2++) {
                if (mode === 'outerRect' && cbHeaderRow.value && j2 === 1) {
                    addLine(left, yEdges[j2], right, yEdges[j2], opts.strokeWidthPt * 2.5);
                } else {
                    addLine(left, yEdges[j2], right, yEdges[j2]);
                }
            }
        }
    }
    // 実行後、生成した罫線＋塗りを選択状態に
    try {
        doc.selection = null;
        for (var si = 0; si < createdItems.length; si++) {
            createdItems[si].selected = true;
        }
        for (var fi = 0; fi < createdFills.length; fi++) {
            createdFills[fi].selected = true;
        }
    } catch (e) {
        // ignore
    }
})();
