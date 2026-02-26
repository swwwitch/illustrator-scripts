#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  グリッド複製ツール / Duplicate in Grid Plus
  Version: v2.0（SCRIPT_VERSION で管理）

  概要 / Overview
  - 選択中のオブジェクトを、指定した繰り返し数と間隔で複製・配置します（ライブプレビュー対応）。
  - 繰り返し方式 / Repeat Method
    - グリッド（G）：行×列で配置（繰り返し数は連動ONが基本）。
    - 行（R）：横の数で1行に配置（縦は常に1、上下方向は無効）。
    - 列（C）：縦の数で1列に配置（横は常に1、左右方向は無効）。
    - ランダム配置（A）：元の中心を基準にランダムに配置（OK後もプレビューと同一の配置になるようオフセットを保持）。
  - 間隔は現在の定規単位（rulerType）で入力し、内部では pt に変換して処理します。
    - 行／列では不要側の間隔入力をディム表示し、［連動］はOFF。
    - ランダムでは間隔［連動］はON（上下=左右）で、値は散らばり範囲に反映されます。
    - ランダムで左右または上下を0にすると、その軸のランダムを完全にOFFできます。
  - 方向（右R／左L／上T／下B）を指定できます（方式により無効化されます）。
  - Fill パネル
    - ［アートボードの端まで］：選択オブジェクト基準で行列を自動計算。
    - ［アートボードいっぱいに］：行列を自動計算し、OK時に複製を含む全体をアートボード中央へ移動（行／列にも対応）。
  - クリッピングマスクを優先して境界を取得（なければ可視境界）。
    - プレビュー描画は一時オブジェクトのみ削除（noteタグで管理）し、既存の「_preview」レイヤー上の他要素は削除しません。
  - UIは2カラム構成：左（繰り返し数／方式）、右（間隔／方向／敷き詰め）。
    - 繰り返し数はスライダー（1〜20）でも操作可能で、ドラッグ中はプレビュー更新をスロットルしてチラつきを軽減。
  - 複数選択時は自動でグループ化してから処理（以降は単一オブジェクトと同様）。/ When multiple objects are selected, they are grouped first and processed like a single object.

  更新日 / Last Updated: 2026-02-26
*/

/* バージョン / Version */
var SCRIPT_VERSION = "v2.0";

/* 言語判定 / Locale detection */
function getCurrentLang() {
    return ($.locale && $.locale.toLowerCase().indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();
var PREVIEW_TAG = "__grid_preview__";

/* ラベル定義 / Label definitions (JA/EN) */
var LABELS = {
    dialogTitle: { ja: "複製配置（グリッド／行／列／ランダム）", en: "Duplicate & Arrange" },
    repeatCount: { ja: "繰り返し数", en: "Count" },
    repeatCountH: { ja: "横", en: "Horizontal" },
    repeatCountV: { ja: "縦", en: "Vertical" },
    gap: { ja: "間隔", en: "Gap" },
    unitFmt: { ja: "間隔（{unit}）:", en: "Gap ({unit}):" },
    fillTitle: { ja: "敷き詰め", en: "Fill" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoSel: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
    alertCountInvalid: { ja: "繰り返し数は1以上の整数を入力してください。", en: "Enter an integer count of 1 or more." },
    alertGapInvalid: { ja: "間隔は数値で入力してください。", en: "Enter a numeric gap value." },
    directionTitle: { ja: "方向", en: "Direction" },
    dirRight: { ja: "右", en: "Right" },
    dirLeft: { ja: "左", en: "Left" },
    verticalTitle: { ja: "縦方向", en: "Vertical" },
    dirUp: { ja: "上", en: "Up" },
    dirDown: { ja: "下", en: "Down" },
    repeatMethodTitle: { ja: "繰り返し方式", en: "Repeat Method" },
    repeatMethodGrid: { ja: "グリッド", en: "Grid" },
    repeatMethodRow: { ja: "行", en: "Row" },
    repeatMethodCol: { ja: "列", en: "Column" },
    repeatMethodRandom: { ja: "ランダム配置", en: "Random" }
};

/* ラベル取得関数 / Label resolver */
function L(key, params) {
    var table = LABELS[key];
    var text = (table && table[lang]) ? table[lang] : key;
    if (params) {
        for (var k in params) if (params.hasOwnProperty(k)) {
            text = text.replace(new RegExp("\\{" + k + "\\}", "g"), params[k]);
        }
    }
    return text;
}

/* テキストフィールドの数値操作 / Arrow key increment–decrement for numeric fields */
function changeValueByArrowKey(edittext) {
    edittext.addEventListener("keydown", function (e) {
        var v = Number(edittext.text);
        if (isNaN(v)) return;
        var kb = ScriptUI.environment.keyboardState, d = 1;
        if (kb.shiftKey) {
            d = 10;
            if (e.keyName == "Up") { v = Math.ceil((v + 1) / d) * d; e.preventDefault(); }
            else if (e.keyName == "Down") { v = Math.floor((v - 1) / d) * d; e.preventDefault(); }
        } else if (kb.altKey) {
            d = 0.1;
            if (e.keyName == "Up") { v += d; e.preventDefault(); }
            else if (e.keyName == "Down") { v -= d; e.preventDefault(); }
        } else {
            d = 1;
            if (e.keyName == "Up") { v += d; e.preventDefault(); }
            else if (e.keyName == "Down") { v -= d; e.preventDefault(); }
        }
        v = kb.altKey ? Math.round(v * 10) / 10 : Math.round(v);
        if (edittext.isInteger) v = Math.max(0, Math.round(v));
        edittext.text = v;
        if (typeof edittext.onChanging === "function") { try { edittext.onChanging(); } catch (_) { } }
    });
}

/* 単位ユーティリティ / Units */
var unitLabelMap = { 0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm", 5: "Q/H", 6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft" };
function getCurrentUnitLabel() { var c = app.preferences.getIntegerPreference("rulerType"); return unitLabelMap[c] || "pt"; }
function getCurrentUnitCode() { return app.preferences.getIntegerPreference("rulerType"); }
function unitToPoints(unitCode, value) {
    var PT_IN = 72, MM_PT = PT_IN / 25.4;
    switch (unitCode) {
        case 0: return value * PT_IN;           // in
        case 1: return value * MM_PT;           // mm
        case 2: return value;                 // pt
        case 3: return value * 12;              // pica
        case 4: return value * (MM_PT * 10);      // cm
        case 5: return value * (MM_PT * 0.25);    // Q/H
        case 6: return value;                 // px ≒ pt
        case 7: return value * PT_IN;           // ft/in → in
        case 8: return value * (MM_PT * 1000);    // m
        case 9: return value * (PT_IN * 36);      // yd
        case 10: return value * (PT_IN * 12);      // ft
        default: return value;
    }
}

/* クリッピングマスク優先の境界取得 / Bounds preferring clipping mask */
function getMaskedBounds(item) {
    try {
        if (item.typename === "GroupItem" && item.clipped) {
            for (var i = 0; i < item.pageItems.length; i++) {
                var pi = item.pageItems[i];
                if (pi.typename === "PathItem" && pi.clipping) return pi.geometricBounds;
                if (pi.typename === "CompoundPathItem" && pi.pathItems.length > 0 && pi.pathItems[0].clipping)
                    return pi.pathItems[0].geometricBounds;
            }
        }
        if (item.typename === "PathItem" && item.clipping) return item.geometricBounds;
        if (item.typename === "CompoundPathItem" && item.pathItems.length > 0 && item.pathItems[0].clipping)
            return item.pathItems[0].geometricBounds;
    } catch (_) { }
    return item.visibleBounds;
}

/* プレビュー用ユーティリティ / Preview utilities */
function getPreviewLayer(doc) {
    var name = "_preview", lyr;
    try { lyr = doc.layers.getByName(name); }
    catch (e) { lyr = doc.layers.add(); lyr.name = name; }
    lyr.visible = true; lyr.locked = false;
    try { lyr.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }
    return lyr;
}
function clearPreview(doc) {
    try {
        var lyr = doc.layers.getByName("_preview");
        // Remove only items we created (note tagged)
        for (var i = lyr.pageItems.length - 1; i >= 0; i--) {
            try {
                var it = lyr.pageItems[i];
                if (it.note === PREVIEW_TAG) it.remove();
            } catch (_) { }
        }
        app.redraw();
    } catch (e) { }
}
function buildPreview(doc, sourceItem, rows, cols, gapX, gapY, w, h, hDir, vDir, baseL, baseT) {
    var lyr = getPreviewLayer(doc);
    clearPreview(doc);
    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            if (r === 0 && c === 0) continue;
            var dup = sourceItem.duplicate(lyr, ElementPlacement.PLACEATBEGINNING);
            dup.note = PREVIEW_TAG;
            var offX = (w + gapX) * c; if (hDir === "left") offX = -offX;
            var offY = (h + gapY) * r;
            var desiredL = baseL + offX;
            var desiredT = (vDir === "up") ? (baseT + offY) : (baseT - offY);
            var mb = getMaskedBounds(dup); // [L,T,R,B]
            var dx = desiredL - mb[0], dy = desiredT - mb[1];
            dup.left += dx; dup.top += dy;
        }
    }
    app.redraw();
}

// ランダム配置プレビュー / Random placement preview
function buildPreviewRandom(doc, sourceItem, offsets) {
    var lyr = getPreviewLayer(doc);
    clearPreview(doc);

    var bb = getMaskedBounds(sourceItem);
    var baseCX = (bb[0] + bb[2]) / 2.0;
    var baseCY = (bb[1] + bb[3]) / 2.0;

    for (var i = 0; i < offsets.length; i++) {
        var dup = sourceItem.duplicate(lyr, ElementPlacement.PLACEATBEGINNING);
        dup.note = PREVIEW_TAG;

        var dx = offsets[i][0];
        var dy = offsets[i][1];

        var db = getMaskedBounds(dup);
        var dupCX = (db[0] + db[2]) / 2.0;
        var dupCY = (db[1] + db[3]) / 2.0;

        dup.left += (baseCX + dx) - dupCX;
        dup.top += (baseCY + dy) - dupCY;
    }
    app.redraw();
}

/* ダイアログ生成＆プレビュー結線 / Build dialog and preview wiring */
function showDialog(doc, sourceItem, w, h) {
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    var offsetX = 300, dialogOpacity = 0.98;
    dlg.onShow = function () { var loc = dlg.location; dlg.location = [loc[0] + offsetX, loc[1]]; };
    dlg.opacity = dialogOpacity;
    dlg.orientation = "column"; dlg.alignChildren = "fill";

    // 2カラムレイアウト / Two-column layout
    var colsGroup = dlg.add("group");
    colsGroup.orientation = "row";
    colsGroup.alignChildren = "fill";
    colsGroup.spacing = 20;

    var leftCol = colsGroup.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "fill";
    leftCol.spacing = 10;

    var rightCol = colsGroup.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = "fill";
    rightCol.spacing = 10;

    var unitCode = getCurrentUnitCode(), unitLabel = getCurrentUnitLabel();

    // Random layout cache (OKでプレビューと一致させる)
    var _randomCache = { key: null, offsets: [] };

    function _randLCG(seedObj) {
        seedObj.v = (seedObj.v * 1664525 + 1013904223) % 4294967296;
        return seedObj.v / 4294967296;
    }

    function _makeRandomOffsets(dupCount, rangeX, rangeY, seed) {
        var seedObj = { v: seed >>> 0 };
        var arr = [];
        for (var i = 0; i < dupCount; i++) {
            var rx = (rangeX > 0) ? ((-rangeX) + (2 * rangeX) * _randLCG(seedObj)) : 0;
            var ry = (rangeY > 0) ? ((-rangeY) + (2 * rangeY) * _randLCG(seedObj)) : 0;
            arr.push([rx, ry]);
        }
        return arr;
    }

    function _getRandomOffsets(cx, cy, gptX, gptY) {
        var dupCount = Math.max(cx, cy) - 1;
        if (dupCount < 0) dupCount = 0;

        // When gap is 0, allow disabling randomness on that axis
        var eps = 1e-9;
        var rangeX = (Math.abs(gptX) <= eps) ? 0 : ((w + gptX) * (cx - 1) / 2.0);
        var rangeY = (Math.abs(gptY) <= eps) ? 0 : ((h + gptY) * (cy - 1) / 2.0);
        if (rangeX < 0) rangeX = 0;
        if (rangeY < 0) rangeY = 0;

        var key = [cx, cy, gptX.toFixed(4), gptY.toFixed(4), w.toFixed(4), h.toFixed(4)].join("|");
        if (_randomCache.key === key && _randomCache.offsets && _randomCache.offsets.length === dupCount) {
            return _randomCache.offsets;
        }

        var seed = (new Date().getTime() & 0xFFFFFFFF) ^ (cx << 16) ^ (cy << 8);
        _randomCache.key = key;
        _randomCache.offsets = _makeRandomOffsets(dupCount, rangeX, rangeY, seed);
        return _randomCache.offsets;
    }

    /* 繰り返し数 / Repeat Count */
    var repeatPanel = leftCol.add("panel", undefined, L("repeatCount"));
    repeatPanel.orientation = "column";
    repeatPanel.alignChildren = "fill";
    repeatPanel.margins = [15, 20, 15, 10];
    repeatPanel.spacing = 10;

    var repeatRow = repeatPanel.add("group");
    repeatRow.orientation = "row";
    repeatRow.alignChildren = "top";
    repeatRow.spacing = 20;

    var repeatLeftCol = repeatRow.add("group"); repeatLeftCol.orientation = "column"; repeatLeftCol.alignChildren = "left";
    var repeatRightCol = repeatRow.add("group"); repeatRightCol.orientation = "column"; repeatRightCol.alignChildren = "left"; repeatRightCol.alignment = ["left", "center"];
    var repeatXGroup = repeatLeftCol.add("group");
    repeatXGroup.add("statictext", undefined, L("repeatCountH") + ":");
    var countXInput = repeatXGroup.add("edittext", undefined, "2"); countXInput.characters = 4; countXInput.isInteger = true; changeValueByArrowKey(countXInput);

    var repeatYGroup = repeatLeftCol.add("group");
    repeatYGroup.add("statictext", undefined, L("repeatCountV") + ":");
    var countYInput = repeatYGroup.add("edittext", undefined, "2"); countYInput.characters = 4; countYInput.isInteger = true; changeValueByArrowKey(countYInput);

    var linkGroup = repeatRightCol.add("group"); linkGroup.alignment = ["left", "center"];
    var linkCheck = linkGroup.add("checkbox", undefined, (lang === "ja" ? "連動" : "Link Horizontal & Vertical"));
    linkCheck.value = true;
    function syncCounts() {
        if (linkCheck.value) {
            countYInput.enabled = false;
            countYInput.text = countXInput.text;
            if (typeof countYInput.onChanging === "function") { try { countYInput.onChanging(); } catch (_) { } }
        } else {
            countYInput.enabled = true;
        }
        if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs();
    }
    linkCheck.onClick = function () { syncCounts(); applyPreview(); };
    syncCounts();

    /* 繰り返し数スライダー（連動時のみ有効） / Repeat count slider (enabled only when linked) */
    var repeatSliderGroup = repeatPanel.add("group");
    repeatSliderGroup.orientation = "row";
    repeatSliderGroup.alignChildren = ["fill", "center"];

    var countSlider = repeatSliderGroup.add("slider", undefined, parseInt(countXInput.text, 10) || 2, 1, 20);
    countSlider.alignment = ["fill", "center"];

    function clampInt(n, min, max) {
        n = parseInt(n, 10);
        if (isNaN(n)) n = min;
        if (n < min) n = min;
        if (n > max) n = max;
        return n;
    }

    function updateCountSliderFromInputs() {
        if (!countSlider) return;
        var isRowMode = !!(methodRow && methodRow.value);
        var isColMode = !!(methodCol && methodCol.value);
        var isRandomMode = !!(methodRandom && methodRandom.value);

        var src = isColMode ? countYInput : countXInput;
        var v = clampInt(src.text, 1, 20);
        countSlider.value = v;

        // Enabled when Row/Column/Random mode, or when Link is ON (and available)
        countSlider.enabled = isRowMode || isColMode || isRandomMode || !!(linkCheck.enabled && linkCheck.value);
    }

    function setCountsFromSlider() {
        var v = Math.round(countSlider.value);
        if (v < 1) v = 1;
        if (v > 20) v = 20;
        countSlider.value = v;

        var isRowMode = !!(methodRow && methodRow.value);
        var isColMode = !!(methodCol && methodCol.value);
        var isRandomMode = !!(methodRandom && methodRandom.value);

        if (isColMode) {
            // Column mode: Horizontal is always 1, slider controls Vertical
            countXInput.text = "1";
            countYInput.text = String(v);
        } else {
            // Grid / Row / Random: slider controls Horizontal
            countXInput.text = String(v);
            if (isRowMode || isRandomMode) {
                countYInput.text = "1";
            } else {
                if (linkCheck.value) countYInput.text = String(v);
            }
        }
    }

    // プレビュー更新をスロットル（スライダードラッグ中のチラつき軽減） / Throttle preview updates while dragging
    var PREVIEW_THROTTLE_MS = 120;
    var _lastPreviewTick = 0;
    function _nowMs() {
        // $.hiresTimer is microseconds
        try { return $.hiresTimer / 1000.0; } catch (_) { return new Date().getTime(); }
    }
    function applyPreviewThrottled(force) {
        var t = _nowMs();
        if (force || (t - _lastPreviewTick) >= PREVIEW_THROTTLE_MS) {
            _lastPreviewTick = t;
            applyPreview();
        }
    }

    countSlider.onChanging = function () {
        if (!countSlider.enabled) return;
        setCountsFromSlider();
        applyPreviewThrottled(false);
    };

    countSlider.onChange = function () {
        if (!countSlider.enabled) return;
        setCountsFromSlider();
        applyPreviewThrottled(true);
    };

    // init slider state
    updateCountSliderFromInputs();

    /* 繰り返し方式 / Repeat method */
    var methodPanel = leftCol.add("panel", undefined, L("repeatMethodTitle"));
    methodPanel.orientation = "column";
    methodPanel.alignChildren = "left";
    methodPanel.margins = [15, 20, 15, 10];
    methodPanel.spacing = 8;

    var methodGrid = methodPanel.add("radiobutton", undefined, L("repeatMethodGrid"));
    var methodRow = methodPanel.add("radiobutton", undefined, L("repeatMethodRow"));
    var methodCol = methodPanel.add("radiobutton", undefined, L("repeatMethodCol"));
    var methodRandom = methodPanel.add("radiobutton", undefined, L("repeatMethodRandom"));
    methodGrid.value = true;


    /* 間隔（現在単位表記） / Gap (show in current ruler units) */
    var gapPanel = rightCol.add("panel", undefined, (lang === "ja" ? ("間隔（" + unitLabel + "）") : ("Gap (" + unitLabel + ")")));
    gapPanel.orientation = "row"; gapPanel.alignChildren = "top";
    gapPanel.margins = [15, 20, 15, 10]; gapPanel.spacing = 20;

    var gapLeftCol = gapPanel.add("group"); gapLeftCol.orientation = "column"; gapLeftCol.alignChildren = "left";
    var gapRightCol = gapPanel.add("group"); gapRightCol.orientation = "column"; gapRightCol.alignChildren = "left"; gapRightCol.alignment = ["left", "center"];

    var gapXGroup = gapLeftCol.add("group");
    gapXGroup.add("statictext", undefined, (lang === "ja" ? "左右:" : "Horizontal:"));
    var gapXInput = gapXGroup.add("edittext", undefined, "10"); gapXInput.characters = 4; changeValueByArrowKey(gapXInput);

    var gapYGroup = gapLeftCol.add("group");
    gapYGroup.add("statictext", undefined, (lang === "ja" ? "上下:" : "Vertical:"));
    var gapYInput = gapYGroup.add("edittext", undefined, "10"); gapYInput.characters = 4; changeValueByArrowKey(gapYInput);

    var gapLinkGroup = gapRightCol.add("group"); gapLinkGroup.alignment = ["left", "center"];
    var gapLink = gapLinkGroup.add("checkbox", undefined, (lang === "ja" ? "連動" : "Link Horizontal & Vertical")); gapLink.value = true;
    function syncGaps() {
        if (gapLink.value) {
            gapYInput.enabled = false; gapYInput.text = gapXInput.text;
            if (typeof gapYInput.onChanging === "function") { try { gapYInput.onChanging(); } catch (_) { } }
        } else { gapYInput.enabled = true; }
    }
    gapLink.onClick = function () { syncGaps(); applyPreview(); };
    syncGaps();

    if (typeof fillToArtboardCheck !== "undefined" && fillToArtboardCheck.value) { recalcCountsForArtboard(); }

    /* 方向パネル（横・縦の展開方向を指定） / Direction panel (horizontal & vertical placement) */
    var dirPanel = rightCol.add("panel", undefined, L("directionTitle"));
    dirPanel.orientation = "column"; dirPanel.alignChildren = "left"; dirPanel.margins = [15, 20, 15, 10];

    var hGroup = dirPanel.add("group"); hGroup.orientation = "row"; hGroup.alignChildren = "left";
    hGroup.add("statictext", undefined, (lang === "ja" ? "横方向" : "Horizontal"));
    var dirRight = hGroup.add("radiobutton", undefined, L("dirRight"));
    var dirLeft = hGroup.add("radiobutton", undefined, L("dirLeft"));
    dirRight.value = true;
    dirRight.onClick = function () {
        // 右選択時はそのままプレビュー
        applyPreview();
    };
    dirLeft.onClick = function () {
        // 左を選んだら「アートボードの端まで」を自動OFF
        try {
            if (typeof fillToArtboardCheck !== "undefined" && fillToArtboardCheck.value) {
                fillToArtboardCheck.value = false;
            }
        } catch (_) { }
        applyPreview();
    };

    var vGroup = dirPanel.add("group"); vGroup.orientation = "row"; vGroup.alignChildren = "left";
    vGroup.add("statictext", undefined, (lang === "ja" ? "縦方向" : "Vertical"));
    var dirUp = vGroup.add("radiobutton", undefined, L("dirUp"));
    var dirDown = vGroup.add("radiobutton", undefined, L("dirDown"));
    dirDown.value = true;
    dirUp.onClick = function () {
        // 上を選んだら「アートボードの端まで」を自動OFF
        try {
            if (typeof fillToArtboardCheck !== "undefined" && fillToArtboardCheck.value) {
                fillToArtboardCheck.value = false;
            }
        } catch (_) { }
        applyPreview();
    };
    dirDown.onClick = function () {
        // 下選択時はそのままプレビュー
        applyPreview();
    };

    /* キー操作でラジオを選択 / Keyboard shortcuts for radios */
    function _selectRadio(radio) {
        if (!radio || radio.enabled === false) return;
        radio.value = true;
        if (typeof radio.onClick === "function") {
            try { radio.onClick(); } catch (_) { }
        }
    }

    dlg.addEventListener("keydown", function (event) {
        // Normalize single-letter key names
        var k = event.keyName;
        if (!k) return;

        // Repeat Method
        if (k === "G") {
            _selectRadio(methodGrid);
            event.preventDefault();
            return;
        }
        if (k === "C") {
            _selectRadio(methodCol);
            event.preventDefault();
            return;
        }
        if (k === "A") {
            _selectRadio(methodRandom);
            event.preventDefault();
            return;
        }
        if (k === "R" && !event.shiftKey) {
            // R = Row
            _selectRadio(methodRow);
            event.preventDefault();
            return;
        }

        // Direction
        if (k === "R" && event.shiftKey) {
            // Shift+R = Right
            _selectRadio(dirRight);
            event.preventDefault();
            return;
        }
        if (k === "L") {
            _selectRadio(dirLeft);
            event.preventDefault();
            return;
        }
        if (k === "T") {
            _selectRadio(dirUp);
            event.preventDefault();
            return;
        }
        if (k === "B") {
            _selectRadio(dirDown);
            event.preventDefault();
            return;
        }
    });

    /* パネル：Fill（アートボードの端まで／アートボードいっぱいに） / Panel: Fill */
    var fillPanel = rightCol.add("panel", undefined, L("fillTitle"));
    fillPanel.orientation = "column"; fillPanel.alignChildren = "left";
    fillPanel.margins = [15, 20, 15, 10]; fillPanel.spacing = 8;

    var fillABGroup = fillPanel.add("group"); fillABGroup.alignment = ["left", "center"];
    var fillToArtboardCheck = fillABGroup.add("checkbox", undefined, (lang === "ja" ? "アートボードの端まで" : "Fill to Artboard Edge"));
    fillToArtboardCheck.value = false;
    fillToArtboardCheck.onClick = function () {
        if (fillToArtboardCheck.value) {
            // 「アートボードの端まで」選択時は方向を［右・下］に固定
            if (typeof dirRight !== "undefined") { dirRight.value = true; }
            if (typeof dirLeft !== "undefined") { dirLeft.value = false; }
            if (typeof dirDown !== "undefined") { dirDown.value = true; }
            if (typeof dirUp !== "undefined") { dirUp.value = false; }
            // 方向コントロールは有効化
            if (typeof dirRight !== "undefined") dirRight.enabled = true;
            if (typeof dirLeft !== "undefined") dirLeft.enabled = true;
            if (typeof dirUp !== "undefined") dirUp.enabled = true;
            if (typeof dirDown !== "undefined") dirDown.enabled = true;
            linkCheck.value = false;
            syncCounts();
            recalcCountsForArtboard();
        }
        if (typeof fillToArtboardFullCheck !== "undefined" && fillToArtboardCheck.value) {
            fillToArtboardFullCheck.value = false;
        }
        applyPreview();
    };

    // 追加：アートボードいっぱいに
    var fillABFullGroup = fillPanel.add("group"); fillABFullGroup.alignment = ["left", "center"];
    var fillToArtboardFullCheck = fillABFullGroup.add("checkbox", undefined, (lang === "ja" ? "アートボードいっぱいに" : "Fill Full Artboard"));
    fillToArtboardFullCheck.value = false;
    fillToArtboardFullCheck.onClick = function () {
        if (fillToArtboardFullCheck.value) {
            // 端までと排他
            if (typeof fillToArtboardCheck !== "undefined") fillToArtboardCheck.value = false;
            // 方向は任意。既定を右・下へ
            if (typeof dirRight !== "undefined") { dirRight.value = true; }
            if (typeof dirLeft !== "undefined") { dirLeft.value = false; }
            if (typeof dirDown !== "undefined") { dirDown.value = true; }
            if (typeof dirUp !== "undefined") { dirUp.value = false; }
            // 方向コントロールをディム（無効化）
            if (typeof dirRight !== "undefined") dirRight.enabled = false;
            if (typeof dirLeft !== "undefined") dirLeft.enabled = false;
            if (typeof dirUp !== "undefined") dirUp.enabled = false;
            if (typeof dirDown !== "undefined") dirDown.enabled = false;
            // 行列をアートボード内に収まる最大値で自動計算
            recalcCountsForArtboardFull();
        } else {
            // チェック解除時は方向コントロールを再有効化
            if (typeof dirRight !== "undefined") dirRight.enabled = true;
            if (typeof dirLeft !== "undefined") dirLeft.enabled = true;
            if (typeof dirUp !== "undefined") dirUp.enabled = true;
            if (typeof dirDown !== "undefined") dirDown.enabled = true;
        }
        applyPreview();
    };

    // 初期状態の有効/無効を同期
    if (fillToArtboardFullCheck.value) {
        if (typeof dirRight !== "undefined") dirRight.enabled = false;
        if (typeof dirLeft !== "undefined") dirLeft.enabled = false;
        if (typeof dirUp !== "undefined") dirUp.enabled = false;
        if (typeof dirDown !== "undefined") dirDown.enabled = false;
    } else {
        if (typeof dirRight !== "undefined") dirRight.enabled = true;
        if (typeof dirLeft !== "undefined") dirLeft.enabled = true;
        if (typeof dirUp !== "undefined") dirUp.enabled = true;
        if (typeof dirDown !== "undefined") dirDown.enabled = true;
    }

    function updateRepeatModeUI() {
        var isRow = (methodRow && methodRow.value);
        var isCol = (methodCol && methodCol.value);
        var isRandom = (methodRandom && methodRandom.value);

        // In Row mode, Vertical count is always 1 and is dimmed.
        if (isRow) {
            // Row mode: Vertical fixed to 1
            countYInput.text = "1";
            countYInput.enabled = false;

            // Horizontal is editable
            countXInput.enabled = true;

            // Link irrelevant
            linkCheck.value = false;
            linkCheck.enabled = false;

            // Gap link OFF and disable
            if (gapLink) {
                gapLink.value = false;
                gapLink.enabled = false;
            }
            // Vertical gap irrelevant in Row mode
            if (gapYInput) gapYInput.enabled = false;
            if (gapXInput) gapXInput.enabled = true;

            // Vertical direction irrelevant
            if (dirUp) dirUp.enabled = false;
            if (dirDown) dirDown.enabled = false;

            // Horizontal direction relevant
            if (dirRight) dirRight.enabled = true;
            if (dirLeft) dirLeft.enabled = true;

        } else if (isCol) {
            // Column mode: Horizontal fixed to 1
            countXInput.text = "1";
            countXInput.enabled = false;

            // Vertical is editable
            countYInput.enabled = true;

            // Link irrelevant
            linkCheck.value = false;
            linkCheck.enabled = false;

            // Gap link OFF and disable
            if (gapLink) {
                gapLink.value = false;
                gapLink.enabled = false;
            }
            // Horizontal gap irrelevant in Column mode
            if (gapXInput) gapXInput.enabled = false;
            if (gapYInput) gapYInput.enabled = true;

            // Horizontal direction irrelevant
            if (dirRight) dirRight.enabled = false;
            if (dirLeft) dirLeft.enabled = false;

            // Vertical direction relevant
            if (dirUp) dirUp.enabled = true;
            if (dirDown) dirDown.enabled = true;
        } else if (isRandom) {
            // Random mode: treat counts like Row (single horizontal count)
            countYInput.text = "1";
            countYInput.enabled = false;
            countXInput.enabled = true;

            // Link OFF
            linkCheck.value = false;
            linkCheck.enabled = false;

            // Gap link ON in Random mode
            if (gapLink) {
                gapLink.value = true;
                gapLink.enabled = true;
            }
            if (typeof syncGaps === "function") syncGaps();

            // Ensure Horizontal gap is editable in Random mode
            if (gapXInput) gapXInput.enabled = true;

            // Directions are irrelevant in Random mode
            if (dirRight) dirRight.enabled = false;
            if (dirLeft) dirLeft.enabled = false;
            if (dirUp) dirUp.enabled = false;
            if (dirDown) dirDown.enabled = false;

            // Random mode: disable Fill options by turning them off
            try { if (fillToArtboardCheck) fillToArtboardCheck.value = false; } catch (_) { }
            try { if (fillToArtboardFullCheck) fillToArtboardFullCheck.value = false; } catch (_) { }
        } else {
            // Grid mode
            linkCheck.enabled = true;
            syncCounts();

            // Gap link and inputs restored
            if (gapLink) gapLink.enabled = true;
            if (gapXInput) gapXInput.enabled = true;
            if (gapYInput) gapYInput.enabled = gapLink ? !gapLink.value : true;
            if (gapLink) gapLink.value = true;
            if (typeof syncGaps === "function") syncGaps();

            if (dirRight) dirRight.enabled = true;
            if (dirLeft) dirLeft.enabled = true;
            if (dirUp) dirUp.enabled = true;
            if (dirDown) dirDown.enabled = true;
        }

        if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs();
    }

    methodGrid.onClick = function () {
        // Grid mode defaults to linked repeat counts
        if (linkCheck) linkCheck.value = true;
        if (gapLink) gapLink.value = true;
        if (typeof syncGaps === "function") syncGaps();
        updateRepeatModeUI();
        // If Fill Full Artboard is ON, direction controls stay dimmed.
        if (fillToArtboardFullCheck && fillToArtboardFullCheck.value) {
            if (dirRight) dirRight.enabled = false;
            if (dirLeft) dirLeft.enabled = false;
            if (dirUp) dirUp.enabled = false;
            if (dirDown) dirDown.enabled = false;
        }
        applyPreview();
    };
    methodRow.onClick = function () {
        if (linkCheck) linkCheck.value = false;
        // Carry over count when switching Column -> Row (use previous vertical count as horizontal count)
        try {
            var prev = clampInt(countYInput.text, 1, 20);
            countXInput.text = String(prev);
        } catch (_) { }
        // Row mode: allow Fill Full Artboard; rows are fixed to 1
        updateRepeatModeUI();
        if (fillToArtboardFullCheck && fillToArtboardFullCheck.value) {
            // Keep direction controls dimmed when Fill Full is ON
            if (dirRight) dirRight.enabled = false;
            if (dirLeft) dirLeft.enabled = false;
            if (dirUp) dirUp.enabled = false;
            if (dirDown) dirDown.enabled = false;
            recalcCountsForArtboardFull();
        }
        applyPreview();
    };

    methodCol.onClick = function () {
        if (linkCheck) linkCheck.value = false;
        // Carry over count when switching Row -> Column (use previous horizontal count as vertical count)
        try {
            var prev = clampInt(countXInput.text, 1, 20);
            countYInput.text = String(prev);
        } catch (_) { }
        // Column mode: allow Fill Full Artboard; cols are fixed to 1
        updateRepeatModeUI();
        if (fillToArtboardFullCheck && fillToArtboardFullCheck.value) {
            if (dirRight) dirRight.enabled = false;
            if (dirLeft) dirLeft.enabled = false;
            if (dirUp) dirUp.enabled = false;
            if (dirDown) dirDown.enabled = false;
            recalcCountsForArtboardFull();
        }
        applyPreview();
    };

    methodRandom.onClick = function () {
        updateRepeatModeUI();
        applyPreview();
    };

    // Initialize
    updateRepeatModeUI();

    function applyPreview() {
        var cx = parseInt(countXInput.text, 10);
        var cy = parseInt(countYInput.text, 10);
        if (methodRow && methodRow.value) cy = 1;
        if (methodRandom && methodRandom.value) cy = 1;
        if (methodCol && methodCol.value) cx = 1;
        var gx = parseFloat(gapXInput.text), gy = parseFloat(gapYInput.text);
        if (isNaN(cx) || cx < 1) return; if (isNaN(cy) || isNaN(cy) || cy < 1) return; if (isNaN(gx) || isNaN(gy)) return;
        var gptX = unitToPoints(unitCode, gx), gptY = unitToPoints(unitCode, gy);
        var hDir = dirRight.value ? "right" : "left";
        var vDir = (dirUp && dirUp.value) ? "up" : "down";
        var isRandomMode = !!(methodRandom && methodRandom.value);
        if (isRandomMode) {
            // Random: spread in both X and Y using the same count so it doesn't collapse vertically
            var n = cx;
            var offsets = _getRandomOffsets(n, n, gptX, gptY);
            buildPreviewRandom(doc, sourceItem, offsets);
        } else {
            var baseMask = getMaskedBounds(sourceItem), baseLeft = baseMask[0], baseTop = baseMask[1];
            buildPreview(doc, sourceItem, cy, cx, gptX, gptY, w, h, hDir, vDir, baseLeft, baseTop);
        }
    }

    // アートボードの端まで：列・行の自動計算
    function recalcCountsForArtboard() {
        try {
            var tgt = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
            var tgtL = tgt[0], tgtT = tgt[1], tgtR = tgt[2], tgtB = tgt[3];
            var baseMask = getMaskedBounds(sourceItem), baseLeft = baseMask[0], baseTop = baseMask[1];

            var gx = parseFloat(gapXInput.text), gy = parseFloat(gapYInput.text);
            if (isNaN(gx) || isNaN(gy)) return;
            var gptX = unitToPoints(unitCode, gx), gptY = unitToPoints(unitCode, gy);
            var stepX = w + gptX, stepY = h + gptY;

            var hDir = dirRight.value ? "right" : "left";
            var vDir = (typeof dirUp !== "undefined" && dirUp.value) ? "up" : "down";

            var cols = 1, rows = 1;
            var isColMode = !!(methodCol && methodCol.value);
            if (methodRow && methodRow.value) {
                rows = 1;
            }
            if (!isColMode && stepX > 0) {
                if (hDir === "right") { var availW = tgtR - baseLeft; cols = Math.floor((availW + gptX) / stepX); }
                else { var availWL = baseLeft - tgtL; cols = Math.floor((availWL + gptX) / stepX); }
                if (cols < 1) cols = 1;
            } else {
                cols = 1;
            }
            if (!(methodRow && methodRow.value) && stepY > 0) {
                if (vDir === "down") { var availH = baseTop - tgtB; rows = Math.floor((availH + gptY) / stepY); }
                else { var availHU = tgtT - baseTop; rows = Math.floor((availHU + gptY) / stepY); }
                if (rows < 1) rows = 1;
            }
            countXInput.text = String(cols);
            countYInput.text = String(rows);
            if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs();
        } catch (e) { }
    }

    // アートボードいっぱいに：アートボード内に収まる最大の列・行を計算（方向は無視）
    function recalcCountsForArtboardFull() {
        // アートボードいっぱいに：アートボード内に収まる最大の列・行を計算（方向は無視）
        try {
            var isRowMode = !!(methodRow && methodRow.value);
            var isColMode = !!(methodCol && methodCol.value);
            var tgt = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
            var tgtW = Math.abs(tgt[2] - tgt[0]);
            var tgtH = Math.abs(tgt[1] - tgt[3]);
            var gx = parseFloat(gapXInput.text), gy = parseFloat(gapYInput.text);
            if (isNaN(gx) || isNaN(gy)) return;
            var gptX = unitToPoints(unitCode, gx), gptY = unitToPoints(unitCode, gy);
            var stepX = w + gptX, stepY = h + gptY;
            var cols = isColMode ? 1 : ((stepX > 0) ? Math.floor((tgtW + gptX) / stepX) : 1);
            var rows = isRowMode ? 1 : ((stepY > 0) ? Math.floor((tgtH + gptY) / stepY) : 1);
            if (cols < 1) cols = 1;
            if (rows < 1) rows = 1;
            countXInput.text = String(cols);
            countYInput.text = String(rows);
            if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs();
        } catch (e) { }
    }

    // 値変更時のプレビュー更新と再計算
    countXInput.onChanging = function () { if (linkCheck.value) countYInput.text = countXInput.text; if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs(); applyPreview(); };
    countXInput.onChange = function () { if (linkCheck.value) countYInput.text = countXInput.text; if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs(); applyPreview(); };
    countYInput.onChanging = function () { if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs(); applyPreview(); };
    countYInput.onChange = function () { if (typeof updateCountSliderFromInputs === "function") updateCountSliderFromInputs(); applyPreview(); };

    gapXInput.onChanging = function () {
        if (gapLink.value) gapYInput.text = gapXInput.text;
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };
    gapXInput.onChange = function () {
        if (gapLink.value) gapYInput.text = gapXInput.text;
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };
    gapYInput.onChanging = function () {
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };
    gapYInput.onChange = function () {
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };

    var btnGroup = dlg.add("group"); btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, L("ok"));

    var result = null;
    okBtn.onClick = function () {
        var cx = parseInt(countXInput.text, 10), cy = parseInt(countYInput.text, 10);
        var gx = parseFloat(gapXInput.text), gy = parseFloat(gapYInput.text);
        if (isNaN(cx) || cx < 1) { alert(L("alertCountInvalid")); return; }
        if (isNaN(cy) || cy < 1) { alert(L("alertCountInvalid")); return; }
        if (isNaN(gx) || isNaN(gy)) { alert(L("alertGapInvalid")); return; }
        if (methodRow && methodRow.value) cy = 1;
        if (methodCol && methodCol.value) cx = 1;
        result = {
            cols: cx, rows: cy,
            gapX: unitToPoints(unitCode, gx),
            gapY: unitToPoints(unitCode, gy),
            randomOffsets: (methodRandom && methodRandom.value) ? (_randomCache.offsets || []) : [],
            direction: dirRight.value ? "right" : "left",
            vDirection: (dirUp && dirUp.value) ? "up" : "down",
            fillToArtboard: fillToArtboardCheck.value,
            fillFullArtboard: (typeof fillToArtboardFullCheck !== "undefined" ? fillToArtboardFullCheck.value : false),
            repeatMethod: (methodRow && methodRow.value) ? "row" :
                ((methodCol && methodCol.value) ? "col" :
                    ((methodRandom && methodRandom.value) ? "random" : "grid"))
        };
        clearPreview(doc); dlg.close();
    };
    cancelBtn.onClick = function () { clearPreview(doc); dlg.close(); };

    var origOnShow = dlg.onShow;
    dlg.onShow = function () {
        if (typeof origOnShow === "function") { try { origOnShow(); } catch (_) { } }
        $.sleep(0); applyPreview(); countXInput.active = true;
    };

    dlg.show();
    return result;
}

/* メイン：検証→ダイアログ→複製 / Main: validate → dialog → duplicate */
function main() {
    if (app.documents.length === 0) { alert(L("alertNoDoc")); return; }
    var doc = app.activeDocument;
    if (doc.selection.length === 0) { alert(L("alertNoSel")); return; }

    var sel = doc.selection;
    // 複数選択時は自動でグループ化してから処理 / Auto-group when multiple items are selected
    if (sel.length > 1) {
        try {
            var grp = doc.groupItems.add();
            // 既存の選択配列はライブで変化し得るため、コピーを回して移動
            var toMove = [];
            for (var i = 0; i < sel.length; i++) toMove.push(sel[i]);
            for (var j = 0; j < toMove.length; j++) {
                try { toMove[j].move(grp, ElementPlacement.PLACEATEND); } catch (_) { }
            }
            // グループを選択対象にして以降の処理を単一オブジェクトと同様に
            doc.selection = null;
            grp.selected = true;
            sel = [grp];
        } catch (e) { }
    }

    var bounds = getMaskedBounds(sel[0]);
    var w = bounds[2] - bounds[0], h = bounds[1] - bounds[3];

    var settings = showDialog(doc, sel[0], w, h);
    if (!settings) return;

    var rows = settings.rows, cols = settings.cols;
    // Row/Column mode constraints
    if (settings.repeatMethod === "row") rows = 1;
    if (settings.repeatMethod === "col") cols = 1;
    var gapX = settings.gapX, gapY = settings.gapY;
    var direction = settings.direction || "right";
    var vDirection = settings.vDirection || "down";

    var isRandomMode = (settings.repeatMethod === "random");

    var baseMask0 = getMaskedBounds(sel[0]);
    var baseLeftMain = baseMask0[0], baseTopMain = baseMask0[1];

    var dupItems = [];

    if (isRandomMode) {
        // Use cached offsets from preview so OK does not change the layout
        var offsets = (settings.randomOffsets && settings.randomOffsets.length) ? settings.randomOffsets : [];

        var bb = getMaskedBounds(sel[0]);
        var baseCX = (bb[0] + bb[2]) / 2.0;
        var baseCY = (bb[1] + bb[3]) / 2.0;

        for (var i = 0; i < offsets.length; i++) {
            var dup = sel[0].duplicate();
            var dxr = offsets[i][0];
            var dyr = offsets[i][1];

            var db = getMaskedBounds(dup);
            var dupCX = (db[0] + db[2]) / 2.0;
            var dupCY = (db[1] + db[3]) / 2.0;

            dup.left += (baseCX + dxr) - dupCX;
            dup.top += (baseCY + dyr) - dupCY;
            dupItems.push(dup);
        }
    } else {
        for (var row = 0; row < rows; row++) {
            for (var col = 0; col < cols; col++) {
                if (row === 0 && col === 0) continue;
                var dup = sel[0].duplicate();
                var offX = (w + gapX) * col; if (direction === "left") offX = -offX;
                var offY = (h + gapY) * row;
                var desiredL = baseLeftMain + offX;
                var desiredT = (vDirection === "up") ? (baseTopMain + offY) : (baseTopMain - offY);
                var mb = getMaskedBounds(dup);
                var dx = desiredL - mb[0], dy = desiredT - mb[1];
                dup.left += dx; dup.top += dy;
                dupItems.push(dup);
            }
        }
    }

    if (settings.fillFullArtboard) {
        try {
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
            var abCX = (ab[0] + ab[2]) / 2.0;
            var abCY = (ab[1] + ab[3]) / 2.0;

            // ユニオン境界を計算（選択元 + 複製）
            var unionL = +Infinity, unionT = -Infinity, unionR = -Infinity, unionB = +Infinity;
            function expandBy(bounds) {
                if (bounds[0] < unionL) unionL = bounds[0];
                if (bounds[1] > unionT) unionT = bounds[1];
                if (bounds[2] > unionR) unionR = bounds[2];
                if (bounds[3] < unionB) unionB = bounds[3];
            }
            expandBy(getMaskedBounds(sel[0]));
            for (var i = 0; i < dupItems.length; i++) {
                expandBy(getMaskedBounds(dupItems[i]));
            }

            var grpCX = (unionL + unionR) / 2.0;
            var grpCY = (unionT + unionB) / 2.0;

            var dx = abCX - grpCX;
            var dy = abCY - grpCY;

            // 全アイテムを同量移動（選択元＋複製）
            sel[0].left += dx; sel[0].top += dy;
            for (var j = 0; j < dupItems.length; j++) {
                dupItems[j].left += dx; dupItems[j].top += dy;
            }
        } catch (e) { }
    }
}
main();