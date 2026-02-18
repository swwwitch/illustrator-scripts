#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * カラージェネレーター / Color Generator
 * 生成したカラーパレットをアートボードに描画し、スウォッチグループにも登録します。 / Draw palettes on the artboard and register them as a swatch group.
 * 対応アルゴリズム: Tailwind / Lightness / Saturation / Complementary / LCH / すべて / Algorithms: Tailwind / Lightness / Saturation / Complementary / LCH / All
 */

var SCRIPT_VERSION = "v1.0.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: { ja: "カラージェネレーター", en: "Color Generator" },
    panelSettings: { ja: "ベースカラー", en: "Base Color" },
    labelHex: { ja: "HEX:", en: "HEX:" },
    labelSteps: { ja: "ステップ数:", en: "Steps:" },
    panelAlgorithm: { ja: "アルゴリズム", en: "Algorithm" },
    panelPreview: { ja: "プレビュー", en: "Preview" },
    panelSwatch: { ja: "スウォッチ", en: "Swatch" },
    chkRegisterSwatchGroup: { ja: "スウォッチグループに登録", en: "Register as Swatch Group" },
    chkConvertToGlobal: { ja: "グローバルに変換", en: "Convert to Global" },
    panelOutput: { ja: "出力", en: "Output" },
    chkOutputHex: { ja: "HEX", en: "HEX" },
    chkOutputRgb: { ja: "RGB", en: "RGB" },
    btnRedraw: { ja: "再描画", en: "Redraw" },
    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnGenerate: { ja: "生成", en: "Generate" },
    alertNoDoc: { ja: "ドキュメントを開いてください。", en: "Please open a document." }
};


function L(key) {
    try {
        var o = LABELS[key];
        if (!o) return key;
        return o[lang] || o.ja || o.en || key;
    } catch (_) {
        return key;
    }
}

/* 選択オブジェクトの塗りからHEXを取得 / Get HEX from selection fill */
function rgbToHexString(r, g, b) {
    function to2(n) {
        var s = Math.round(n).toString(16);
        return (s.length === 1) ? ("0" + s) : s;
    }
    return "#" + to2(r) + to2(g) + to2(b);
}

/* CMYK → RGB 変換 / Convert CMYK to RGB (using Illustrator if possible) */
function cmykToRgbApprox(c, m, y, k) {
    // c,m,y,k are 0..100
    try {
        // Illustrator's converter (more accurate than naive formula)
        var out = app.convertSampleColor(
            ImageColorSpace.CMYK,
            [c, m, y, k],
            ImageColorSpace.RGB,
            ColorConvertPurpose.defaultpurpose,
            true
        );
        if (out && out.length >= 3) {
            return { r: out[0], g: out[1], b: out[2] };
        }
    } catch (_) { }

    // Fallback (simple approximation)
    var C = c / 100, M = m / 100, Y = y / 100, K = k / 100;
    var r = 255 * (1 - C) * (1 - K);
    var g = 255 * (1 - M) * (1 - K);
    var b = 255 * (1 - Y) * (1 - K);
    return { r: r, g: g, b: b };
}

function normalizeHexString(hex) {
    if (!hex) return null;
    hex = String(hex).replace(/\s+/g, "");
    if (hex.charAt(0) !== "#") hex = "#" + hex;
    // 3桁は6桁へ
    if (/^#[0-9A-Fa-f]{3}$/.test(hex)) {
        hex = "#" + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2) + hex.charAt(3) + hex.charAt(3);
    }
    if (!/^#[0-9A-Fa-f]{6}$/.test(hex)) return null;
    return "#" + hex.substring(1).toLowerCase();
}

function tryGetSelectionFillHex() {
    try {
        if (app.documents.length === 0) return null;
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length < 1) return null;

        var it = doc.selection[0];

        // グループの場合は最初のパスを探す / If group, find first path
        if (it.typename === "GroupItem") {
            for (var i = 0; i < it.pageItems.length; i++) {
                if (it.pageItems[i].typename === "PathItem") { it = it.pageItems[i]; break; }
            }
        }

        // テキストの場合は文字の塗り / If text, use character fill
        if (it.typename === "TextFrame") {
            try {
                var c = it.textRange.characterAttributes.fillColor;
                if (c && c.typename === "RGBColor") {
                    return normalizeHexString(rgbToHexString(c.red, c.green, c.blue));
                }
                if (c && c.typename === "CMYKColor") {
                    var rgb = cmykToRgbApprox(c.cyan, c.magenta, c.yellow, c.black);
                    return normalizeHexString(rgbToHexString(rgb.r, rgb.g, rgb.b));
                }
            } catch (_) { }
            return null;
        }

        // パスの場合 / PathItem fill
        if (it.typename === "PathItem") {
            if (!it.filled) return null;
            var fc = it.fillColor;
            if (fc && fc.typename === "RGBColor") {
                return normalizeHexString(rgbToHexString(fc.red, fc.green, fc.blue));
            }
            if (fc && fc.typename === "CMYKColor") {
                var rgb2 = cmykToRgbApprox(fc.cyan, fc.magenta, fc.yellow, fc.black);
                return normalizeHexString(rgbToHexString(rgb2.r, rgb2.g, rgb2.b));
            }
        }

        return null;
    } catch (_) {
        return null;
    }
}

/* コントラストシフト / Contrast shift (-1.0 .. 1.0)
 * shift = -1 => 全て中間(128)へ寄せる（コントラスト0）
 * shift =  0 => 変化なし
 * shift =  1 => コントラスト2倍（クリップあり）
 */
function clamp(v, minV, maxV) {
    if (v < minV) return minV;
    if (v > maxV) return maxV;
    return v;
}

function applyContrastShiftToChannel(v, shift) {
    var factor = 1 + shift; // 0..2
    var out = (v - 128) * factor + 128;
    return clamp(Math.round(out), 0, 255);
}

function applyContrastShiftToRgb(rgb, shift) {
    if (!rgb) return rgb;
    shift = clamp(Number(shift) || 0, -1, 1);
    if (Math.abs(shift) < 1e-9) return rgb;
    return {
        r: applyContrastShiftToChannel(rgb.r, shift),
        g: applyContrastShiftToChannel(rgb.g, shift),
        b: applyContrastShiftToChannel(rgb.b, shift)
    };
}

/* =========================================
 * changeValueByArrowKey(editText)
 * ↑↓ : ±1 / Shift+↑↓ : ±10（整数のみ）
 * - Option(Alt) の 0.1 増減は不要（未実装）
 * - 範囲は editText.__min / editText.__max があればそれを使用
 * - 値反映後、editText.onChanging / onChange を呼んで同期・プレビュー更新
 * ========================================= */

function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.shiftKey ? 10 : 1;

        if (event.keyName === "Up") value += delta;
        if (event.keyName === "Down") value -= delta;

        // 整数のみ
        value = Math.round(value);

        // 範囲クランプ（指定があれば）
        var minV = (typeof editText.__min === "number") ? editText.__min : null;
        var maxV = (typeof editText.__max === "number") ? editText.__max : null;

        if (minV !== null && value < minV) value = minV;
        if (maxV !== null && value > maxV) value = maxV;

        editText.text = String(value);

        // デフォルト挙動抑止
        event.preventDefault();

        // 既存の同期・プレビュー更新を流用
        try { if (typeof editText.onChanging === "function") editText.onChanging(); } catch (_) { }
        try { if (typeof editText.onChange === "function") editText.onChange(); } catch (_) { }
    });
}

function main() {
    if (app.documents.length === 0) {
        alert(L("alertNoDoc"));
        return;
    }

    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 15;

    // 上部：2カラム（左=設定 / 右=スウォッチ） / Top: 2 columns (Left=Settings / Right=Swatch)
    var gTop = win.add("group");
    gTop.orientation = "row";
    gTop.alignChildren = ["fill", "top"];
    gTop.alignment = "fill";

    // 左カラム（ベースカラー / ステップ数） / Left column (Base color / Steps)
    var leftCol = gTop.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["fill", "top"];
    leftCol.alignment = "fill";

    // --- 1. 入力エリア ---
    var inputPanel = leftCol.add("panel", undefined, L("panelSettings"));
    inputPanel.margins = [15, 20, 15, 10];
    inputPanel.orientation = "column";
    inputPanel.alignChildren = ["left", "top"];

    // 右カラム（スウォッチ / 出力） / Right column (Swatch / Output)
    var rightCol = gTop.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];
    rightCol.alignment = "fill";

    // 右カラム：スウォッチ / Right column: Swatch
    var swatchPanel = rightCol.add("panel", undefined, L("panelSwatch"));
    swatchPanel.margins = [15, 20, 15, 10];
    swatchPanel.orientation = "column";
    swatchPanel.alignChildren = ["left", "top"];
    // まだロジック未接続のためディム表示 / Disabled until wired
    swatchPanel.enabled = false;

    var chkRegisterSwatchGroup = swatchPanel.add("checkbox", undefined, L("chkRegisterSwatchGroup"));
    chkRegisterSwatchGroup.value = true;
    var chkConvertToGlobal = swatchPanel.add("checkbox", undefined, L("chkConvertToGlobal"));
    chkConvertToGlobal.value = false;

    // 右カラム：出力 / Right column: Output
    var outputPanel = rightCol.add("panel", undefined, L("panelOutput"));
    outputPanel.margins = [15, 20, 15, 10];
    outputPanel.orientation = "column";
    outputPanel.alignChildren = ["left", "top"];

    var chkOutputHex = outputPanel.add("checkbox", undefined, L("chkOutputHex"));
    chkOutputHex.value = true;

    var chkOutputRgb = outputPanel.add("checkbox", undefined, L("chkOutputRgb"));
    chkOutputRgb.value = false;

    var g1 = inputPanel.add("group");
    g1.orientation = "row";
    g1.alignChildren = ["left", "center"];

    // 現在のHEXカラー表示（カラーチップ） / Current HEX color swatch
    var colorSwatch = g1.add("panel");
    try { colorSwatch.margins = 0; } catch (_) { }
    colorSwatch.preferredSize = [46, 46];

    g1.add("statictext", undefined, L("labelHex"));

    var __initHex = tryGetSelectionFillHex() || "#3b82f6";
    var inputHex = g1.add("edittext", undefined, __initHex);
    inputHex.characters = 8;

    function getPreviewRgb() {
        try {
            var nh = normalizeHexString(inputHex.text);
            if (!nh) return null;
            var rgb = hexToRgb(nh);
            return rgb;
        } catch (_) { }
        return null;
    }

    function updateColorSwatch() {
        try { colorSwatch.update(); } catch (_) { }
    }

    colorSwatch.onDraw = function () {
        var g = this.graphics;
        var rgb = getPreviewRgb();
        var w = this.size[0];
        var h = this.size[1];

        // 背景（無効時） / Background when invalid
        var fill;
        if (!rgb) {
            fill = g.newBrush(g.BrushType.SOLID_COLOR, [0.85, 0.85, 0.85, 1]);
        } else {
            fill = g.newBrush(g.BrushType.SOLID_COLOR, [rgb.r / 255, rgb.g / 255, rgb.b / 255, 1]);
        }

        g.newPath();
        g.rectPath(0, 0, w, h);
        g.fillPath(fill);

        // 枠線 / Border
        try {
            var stroke = g.newPen(g.PenType.SOLID_COLOR, [0.6, 0.6, 0.6, 1], 1);
            g.strokePath(stroke);
        } catch (_) { }
    };
    // 初期値を正規化 / Normalize initial value
    try {
        var nh = normalizeHexString(inputHex.text);
        if (nh) inputHex.text = nh;
    } catch (_) { }
    updateColorSwatch();


    // ステップ数パネル / Steps panel
    var stepsPanel = leftCol.add("panel", undefined, L("labelSteps"));
    stepsPanel.margins = [15, 20, 15, 10];
    stepsPanel.orientation = "column";
    stepsPanel.alignChildren = ["left", "top"];

    // 上段：ラベル + 入力 / Top: label + input
    var gStepsTop = stepsPanel.add("group");
    gStepsTop.orientation = "row";
    gStepsTop.alignChildren = ["left", "center"];
    gStepsTop.alignment = "left";

    // gStepsTop.add("statictext", undefined, L("labelSteps"));
    var inputCount = gStepsTop.add("edittext", undefined, "5");
    inputCount.characters = 3;
    // 矢印キーで増減（整数のみ） / Change by arrow keys (integers)
    inputCount.__min = 1;
    inputCount.__max = 20;
    changeValueByArrowKey(inputCount);

    // 下段：スライダー / Bottom: slider
    var gStepsBottom = stepsPanel.add("group");
    gStepsBottom.orientation = "row";
    gStepsBottom.alignChildren = ["left", "center"];
    gStepsBottom.alignment = "left";

    var sldCount = gStepsBottom.add("slider", undefined, 5, 1, 20);
    sldCount.preferredSize.width = 200;

    // CONTRAST SHIFT パネル / Contrast Shift panel
    var contrastPanel = leftCol.add("panel", undefined, "コントラストシフト");
    contrastPanel.margins = [15, 20, 15, 10];
    contrastPanel.orientation = "column";
    contrastPanel.alignChildren = ["left", "top"];

    var gCtrTop = contrastPanel.add("group");
    gCtrTop.orientation = "row";
    gCtrTop.alignChildren = ["left", "center"];
    gCtrTop.alignment = "left";

    var inputContrast = gCtrTop.add("edittext", undefined, "0.0");
    inputContrast.characters = 4;

    var gCtrBottom = contrastPanel.add("group");
    gCtrBottom.orientation = "row";
    gCtrBottom.alignChildren = ["left", "center"];
    gCtrBottom.alignment = "left";

    var sldContrast = gCtrBottom.add("slider", undefined, 0, -1, 1);
    sldContrast.preferredSize.width = 200;

    function getContrastShiftValue() {
        var t = String(inputContrast.text);
        t = t.replace(/^\s+|\s+$/g, "");
        t = t.replace(",", ".");
        var v = parseFloat(t);
        if (!isFinite(v) || isNaN(v)) v = Number(sldContrast.value) || 0;
        v = clamp(v, -1, 1);
        v = Math.round(v * 10) / 10; // 表示は小数1桁
        return v;
    }

    function setContrastShiftValue(v) {
        v = clamp(Number(v) || 0, -1, 1);
        v = Math.round(v * 10) / 10;
        inputContrast.text = v.toFixed(1);
        try { sldContrast.value = v; } catch (_) { }
    }
    setContrastShiftValue(0);

    // CONTRAST SHIFT UI 有効/無効 / Enable/disable contrast UI
    function setContrastEnabled(v) {
        try { contrastPanel.enabled = !!v; } catch (_) { }
        try { inputContrast.enabled = !!v; } catch (_) { }
        try { sldContrast.enabled = !!v; } catch (_) { }
    }

    // --- 2. アルゴリズム選択 (ラジオボタン) ---
    var algoPanel = win.add("panel", undefined, L("panelAlgorithm"));
    algoPanel.margins = [15, 20, 15, 10];
    algoPanel.orientation = "column";
    algoPanel.alignChildren = ["left", "top"];

    var rbTailwind = algoPanel.add("radiobutton", undefined, "Tailwind CSS");
    var rbLight = algoPanel.add("radiobutton", undefined, "Lightness Scale");
    var rbLightGeo = algoPanel.add("radiobutton", undefined, "Lightness Scale (Geometric)");
    var rbLch = algoPanel.add("radiobutton", undefined, "LCH");
    var rbSatur = algoPanel.add("radiobutton", undefined, "Saturation Scale");
    var rbComple = algoPanel.add("radiobutton", undefined, "Complementary");
    var rbAll = algoPanel.add("radiobutton", undefined, "すべて");
    rbAll.value = true; // デフォルト

    // --- 3. プレビューエリア ---
    var previewPanel = win.add("panel", undefined, L("panelPreview"));
    previewPanel.margins = [15, 20, 15, 10];
    previewPanel.size = [420, 70];

    previewPanel.onDraw = function () {
        var g = this.graphics;
        var data = getActivePalette();
        if (!data) return;
        var w = this.size[0] / data.length;
        for (var i = 0; i < data.length; i++) {
            var c0 = data[i];
            var shift = getContrastShiftValue();
            var c = applyContrastShiftToRgb(c0, shift);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, [c.r / 255, c.g / 255, c.b / 255, 1]);
            g.newPath();
            g.rectPath(i * w, 10, w, 40);
            g.fillPath(brush);
        }
    };

    // プレビュー更新（notifyより update の方が確実）
    var __previewEnabled = true;
    var updatePreview = function () {
        try {
            if (!__previewEnabled) return;
            previewPanel.update(); // onDraw を発火させる
        } catch (_) {
            try { previewPanel.notify("onDraw"); } catch (__) { }
        }
    };
    inputHex.onChanging = function () {
        updateColorSwatch();
        updatePreview();
    };

    // CONTRAST SHIFT のUI連動（updatePreview定義後に接続）
    inputContrast.onChanging = function () {
        var v = getContrastShiftValue();
        try { sldContrast.value = v; } catch (_) { }
        updatePreview();
    };
    inputContrast.onChange = function () {
        var v = getContrastShiftValue();
        setContrastShiftValue(v);
        updatePreview();
    };
    sldContrast.onChanging = function () {
        var v = Math.round(Number(sldContrast.value) * 10) / 10;
        setContrastShiftValue(v);
        updatePreview();
    };
    sldContrast.onChange = function () {
        var v = Math.round(Number(sldContrast.value) * 10) / 10;
        setContrastShiftValue(v);
        updatePreview();
    };

    // 入力欄 ⇄ スライダー同期
    function clampInt(v, minV, maxV) {
        v = Math.round(Number(v));
        if (!isFinite(v)) v = minV;
        if (v < minV) v = minV;
        if (v > maxV) v = maxV;
        return v;
    }

    inputCount.onChanging = function () {
        var v = clampInt(inputCount.text, 1, 20);
        sldCount.value = v;
        updatePreview();
    };

    // Tab移動など確定時にも反映
    inputCount.onChange = function () {
        var v = clampInt(inputCount.text, 1, 20);
        sldCount.value = v;
        inputCount.text = String(v);
        updatePreview();
    };

    sldCount.onChanging = function () {
        var v = clampInt(sldCount.value, 1, 20);
        inputCount.text = String(v);
        updatePreview();
    };

    // 念のため onChange でも反映
    sldCount.onChange = function () {
        var v = clampInt(sldCount.value, 1, 20);
        inputCount.text = String(v);
        updatePreview();
    };

    function setPreviewEnabled(v) {
        __previewEnabled = !!v;
        // visible は変えない（レイアウトが上下にガタつくため）
        try { previewPanel.enabled = __previewEnabled; } catch (_) { }
        // 見た目を即時反映
        try { previewPanel.update(); } catch (_) { }
    }

    // ステップUI有効/無効
    function setStepsEnabled(v) {
        try { stepsPanel.enabled = !!v; } catch (_) { }
        try { inputCount.enabled = !!v; } catch (_) { }
        try { sldCount.enabled = !!v; } catch (_) { }
    }

    // 初期状態：デフォルトは「すべて」なのでプレビュー無効 / Initial: default is "All" => disable preview
    setPreviewEnabled(false);
    setStepsEnabled(false);
    setContrastEnabled(false);

    rbTailwind.onClick = function () {
        // Tailwind は常に11・ステップ数UIはディム / Tailwind: fixed 11, dim steps UI
        inputCount.text = "11";
        try { sldCount.value = 11; } catch (_) { }
        setStepsEnabled(false);
        setContrastEnabled(true);
        setPreviewEnabled(true);
        updatePreview();
    };
    rbLight.onClick = function () {
        setStepsEnabled(true);
        setContrastEnabled(true);
        setPreviewEnabled(true);
        updatePreview();
    };
    rbLightGeo.onClick = function () {
        setStepsEnabled(true);
        setContrastEnabled(true);
        setPreviewEnabled(true);
        updatePreview();
    };
    rbSatur.onClick = function () {
        setStepsEnabled(true);
        setContrastEnabled(true);
        setPreviewEnabled(true);
        updatePreview();
    };
    rbComple.onClick = function () {
        setStepsEnabled(true);
        setContrastEnabled(true);
        setPreviewEnabled(true);
        updatePreview();
    };
    rbLch.onClick = function () {
        // LCH は常に11・ステップ数UIはディム / LCH: fixed 11, dim steps UI
        inputCount.text = "11";
        try { sldCount.value = 11; } catch (_) { }
        setStepsEnabled(false);
        setContrastEnabled(true);
        setPreviewEnabled(true);
        updatePreview();
    };
    rbAll.onClick = function () {
        // 「すべて」はステップ数=11で計算（UIも合わせる）
        inputCount.text = "11";
        try { sldCount.value = 11; } catch (_) { }
        // 「すべて」選択時はプレビューを無効化（レイアウトは固定）
        setPreviewEnabled(false);
        setStepsEnabled(false);
        setContrastEnabled(false);
    };

    // --- ボタン ---
    var btnRow = win.add("group");
    btnRow.orientation = "row";
    btnRow.alignChildren = ["fill", "center"];
    btnRow.alignment = "fill";

    // 左：再描画 / Left: Redraw
    var gBtnLeft = btnRow.add("group");
    gBtnLeft.orientation = "row";
    gBtnLeft.alignChildren = ["left", "center"];
    gBtnLeft.alignment = "left";

    var btnRedraw = gBtnLeft.add("button", undefined, L("btnRedraw"));

    // 中央：スペーサー / Center: spacer
    var spacer = btnRow.add("group");
    spacer.alignment = "fill";
    spacer.minimumSize.width = 160;

    // 右：キャンセル / 生成 / Right: Cancel / Generate
    var gBtnRight = btnRow.add("group");
    gBtnRight.orientation = "row";
    gBtnRight.alignChildren = ["right", "center"];
    gBtnRight.alignment = "right";

    var btnCancel = gBtnRight.add("button", undefined, L("btnCancel"));
    var btnOk = gBtnRight.add("button", undefined, L("btnGenerate"), { name: "ok" });

    // 再描画 / Redraw (preview refresh)
    btnRedraw.onClick = function () {
        try { updateColorSwatch(); } catch (_) { }

        // ダイアログ内プレビューを強制再描画（プレビュー無効時も） / Force dialog preview repaint
        try {
            var __prev = __previewEnabled;
            __previewEnabled = true;
            try { previewPanel.update(); } catch (_) { }
            try { updatePreview(); } catch (_) { }
            __previewEnabled = __prev;
        } catch (_) { }

        // ダイアログ全体の再描画 / Refresh dialog
        try { win.update(); } catch (_) { }

        // Illustrator画面の再描画 / Force Illustrator UI redraw
        try { app.redraw(); } catch (_) { }
    };

    // キャンセル / Cancel
    btnCancel.onClick = function () {
        try { win.close(); } catch (_) { }
    };

    // ステップ数取得（全角数字→半角、NaN時はスライダー値を採用、1〜20にクランプ）
    function getStepCount() {
        var t = String(inputCount.text);

        // 全角数字を半角へ
        t = t.replace(/[０-９]/g, function (ch) {
            return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
        });

        // 余計な空白を除去
        t = t.replace(/^\s+|\s+$/g, "");

        var v = parseInt(t, 10);

        // 文字入力が不正な場合はスライダーにフォールバック
        if (!isFinite(v) || isNaN(v)) {
            try { v = Math.round(Number(sldCount.value)); } catch (_) { v = 5; }
        }

        // クランプ（UI仕様：1〜20）
        if (v < 1) v = 1;
        if (v > 20) v = 20;

        return v;
    }

    function getActivePalette() {
        var hex = inputHex.text;
        var count = getStepCount();
        if (!/^#?([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/.test(hex)) return null;

        var rgb = hexToRgb(hex);
        var hsl = rgbToHsl(rgb.r, rgb.g, rgb.b);

        if (rbAll.value) return null;
        if (rbTailwind.value) return algoTailwind(hsl);
        if (rbLight.value) return algoLightness(hsl, count);
        if (rbLightGeo.value) return algoLightnessGeometric(hsl, count);
        if (rbSatur.value) return algoSaturation(hsl, count);
        if (rbComple.value) return algoComplementary(hsl, count);
        if (rbLch.value) return algoLchFromRgb(rgb);
        return null;
    }

    function getActiveAlgorithmName() {
        if (rbTailwind.value) return rbTailwind.text;
        if (rbLight.value) return rbLight.text;
        if (rbLightGeo.value) return rbLightGeo.text;
        if (rbSatur.value) return rbSatur.text;
        if (rbComple.value) return rbComple.text;
        if (rbLch.value) return rbLch.text;
        if (rbAll.value) return rbAll.text;
        return "";
    }

    btnOk.onClick = function () {
        var data = getActivePalette();
        // OK押下時に値を正規化してUIへ反映（全角/空欄対策）
        var vFixed = getStepCount();
        inputCount.text = String(vFixed);
        try { sldCount.value = vFixed; } catch (_) { }
        // 出力オプション / Output options
        var outHex = !!chkOutputHex.value;
        var outRgb = !!chkOutputRgb.value;
        var contrastShift = getContrastShiftValue();
        if (rbAll.value) {
            var hex = inputHex.text;
            var rgb0 = hexToRgb(hex);
            var hsl0 = rgbToHsl(rgb0.r, rgb0.g, rgb0.b);

            // 「すべて」生成（固定順）
            win.close();
            drawAllToIllustrator(hsl0, rgb0, outHex, outRgb, contrastShift);
            return;
        }

        if (data) {
            var algoName = getActiveAlgorithmName();
            win.close();
            drawToIllustrator(data, algoName, outHex, outRgb, contrastShift);
        }
    };

    // Escキーでも閉じる + EnterでOK / Close with Esc, trigger OK with Enter
    win.addEventListener("keydown", function (e) {
        try {
            if (!e) return;
            if (e.keyName === "Escape") {
                e.preventDefault();
                win.close();
                return;
            }
            if (e.keyName === "Enter" || e.keyName === "Return") {
                e.preventDefault();
                try { btnOk.notify("onClick"); } catch (_) {
                    try { btnOk.onClick(); } catch (__) { }
                }
                return;
            }
        } catch (_) { }
    });
    win.show();
}

// --- アルゴリズムの実装 ---

// 1. Tailwind風 (固定の輝度ステップ)
function algoTailwind(hsl) {
    var weights = [0.95, 0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3, 0.2, 0.1, 0.05]; // 50-950
    var res = [];
    // 左→右を「濃 → 薄」に
    for (var i = 0; i < weights.length; i++) {
        res.push(hslToRgb(hsl.h, hsl.s, 1 - weights[i]));
    }
    return res;
}

// 2. Lightness Scale (輝度変化) - 出力は常に count 色 / Always output exactly "count" colors
function algoLightness(hsl, count) {
    var n = Math.round(Number(count));
    if (!isFinite(n) || n < 1) n = 1;

    // n=1 のときは元色のみ / If n=1, only the base color
    if (n === 1) return [hslToRgb(hsl.h, hsl.s, hsl.l)];

    // 範囲を決めて等間隔にサンプリング / Sample evenly in a lightness range
    // - 暗側は元Lの30%〜最低0.06
    // - 明側は 0.94 まで（白飛び回避）
    var lMin = Math.max(0.06, hsl.l * 0.30);
    var lMax = Math.min(0.94, hsl.l + (1 - hsl.l) * 0.85);

    // もし範囲が狭すぎる場合は安全に広げる / Safety expand if needed
    if (lMax - lMin < 0.10) {
        lMin = Math.max(0.06, hsl.l - 0.15);
        lMax = Math.min(0.94, hsl.l + 0.15);
    }

    var res = [];
    for (var i = 0; i < n; i++) {
        var t = (n === 1) ? 0 : (i / (n - 1));
        var L = lMin + (lMax - lMin) * t;
        res.push(hslToRgb(hsl.h, hsl.s, L));
    }

    // 左→右を「濃 → 薄」にする / Dark -> Light left-to-right
    // （lMinが暗、lMaxが明なのでそのままでOK）
    return res;
}

// 2b. Lightness Scale (Geometric) - 出力は常に count 色 / Always output exactly "count" colors
// 画像の「等比（幾何）スケール」的に、暗部側の変化を細かくする。
// lMin..lMax の範囲を指数補間でサンプリング（濃→薄）
function algoLightnessGeometric(hsl, count) {
    var n = Math.round(Number(count));
    if (!isFinite(n) || n < 1) n = 1;

    if (n === 1) return [hslToRgb(hsl.h, hsl.s, hsl.l)];

    // Linear版と同じ範囲をベースにする
    var lMin = Math.max(0.06, hsl.l * 0.30);
    var lMax = Math.min(0.94, hsl.l + (1 - hsl.l) * 0.85);

    if (lMax - lMin < 0.10) {
        lMin = Math.max(0.06, hsl.l - 0.15);
        lMax = Math.min(0.94, hsl.l + 0.15);
    }

    // 比率（lMin>0 を保証） / ratio
    var ratio = lMax / Math.max(0.0001, lMin);

    var res = [];
    for (var i = 0; i < n; i++) {
        var t = (i / (n - 1)); // 0..1
        // 幾何補間: L = lMin * ratio^t
        var L = lMin * Math.pow(ratio, t);
        // 念のためクランプ
        if (L < 0) L = 0;
        if (L > 1) L = 1;
        res.push(hslToRgb(hsl.h, hsl.s, L));
    }
    return res;
}

// 3. Saturation Scale (彩度変化) - 出力は常に count 色 / Always output exactly "count" colors
function algoSaturation(hsl, count) {
    var n = Math.round(Number(count));
    if (!isFinite(n) || n < 1) n = 1;

    if (n === 1) return [hslToRgb(hsl.h, hsl.s, hsl.l)];

    // 彩度は低→高（元S）で等間隔 / Low saturation -> base saturation
    var sMin = 0.0;
    var sMax = Math.min(1.0, Math.max(0.0, hsl.s));

    var res = [];
    for (var i = 0; i < n; i++) {
        var t = (i / (n - 1));
        var S = sMin + (sMax - sMin) * t;
        res.push(hslToRgb(hsl.h, S, hsl.l));
    }
    return res;
}

// 4. Complementary (補色 + 輝度変化)
function algoComplementary(hsl, count) {
    var n = Math.round(Number(count));
    if (!isFinite(n) || n < 1) n = 1;

    var compH = (hsl.h + 0.5) % 1.0;

    // 常に合計 n 色になるように分配（n=11 なら 6+5 など）
    var n1 = Math.ceil(n / 2);
    var n2 = n - n1;

    var res = algoLightness(hsl, n1);
    var compRes = (n2 > 0) ? algoLightness({ h: compH, s: hsl.s, l: hsl.l }, n2) : [];
    return res.concat(compRes);
}

/* 5. LCH (固定11ステップ)
 * - Lightness: 8刻みの10ステップ + 99% を追加して合計11
 * - Chroma: 58 を基準に、sRGB 色域内に収まるまで下げる
 * - Hue: 入力色の Hue を使用（LCHに変換して保持）
 */
function algoLchFromRgb(baseRgb) {
    // base RGB -> LCH
    var lab = rgbToLab(baseRgb.r, baseRgb.g, baseRgb.b);
    var lch = labToLch(lab);

    var H = lch.h;

    // 8刻みの10ステップ（暗→明）＋ 99 を先頭に追加（合計11）
    // 90,82,74,66,58,50,42,34,26,18 に 99 を足す
    var Ls = [99, 90, 82, 74, 66, 58, 50, 42, 34, 26, 18];

    var res = [];
    for (var i = 0; i < Ls.length; i++) {
        var L = Ls[i];
        var C = 58;

        // gamut-fit: Cを下げてsRGB内に収める
        var rgb = null;
        for (var c = C; c >= 0; c -= 1) {
            var lab2 = lchToLab({ l: L, c: c, h: H });
            var rgb2 = labToRgb(lab2.l, lab2.a, lab2.b);
            if (rgb2 && rgb2.inGamut) {
                rgb = { r: rgb2.r, g: rgb2.g, b: rgb2.b };
                break;
            }
        }
        // それでも取れない場合はグレーにフォールバック
        if (!rgb) {
            rgb = { r: Math.round(L / 100 * 255), g: Math.round(L / 100 * 255), b: Math.round(L / 100 * 255) };
        }
        res.push(rgb);
    }
    res.reverse(); // 左→右を「濃 → 薄」に
    return res;
}

// --- ユーティリティ ---

/* =========================================
 * Color space utils: sRGB ↔ XYZ(D65) ↔ Lab ↔ LCH
 * - 参照白色点: D65 (Xn=95.047, Yn=100.000, Zn=108.883)
 * - Lab: CIE 1976
 * ========================================= */

function srgbToLinear(u) {
    u = u / 255;
    return (u <= 0.04045) ? (u / 12.92) : Math.pow((u + 0.055) / 1.055, 2.4);
}

function linearToSrgb(u) {
    var v = (u <= 0.0031308) ? (12.92 * u) : (1.055 * Math.pow(u, 1 / 2.4) - 0.055);
    return Math.round(Math.min(1, Math.max(0, v)) * 255);
}

function rgbToXyz(r, g, b) {
    var R = srgbToLinear(r);
    var G = srgbToLinear(g);
    var B = srgbToLinear(b);

    // sRGB D65
    var X = (R * 0.4124564 + G * 0.3575761 + B * 0.1804375) * 100;
    var Y = (R * 0.2126729 + G * 0.7151522 + B * 0.0721750) * 100;
    var Z = (R * 0.0193339 + G * 0.1191920 + B * 0.9503041) * 100;

    return { x: X, y: Y, z: Z };
}

function xyzToRgb(X, Y, Z) {
    X /= 100; Y /= 100; Z /= 100;

    var R = 3.2404542 * X + -1.5371385 * Y + -0.4985314 * Z;
    var G = -0.9692660 * X + 1.8760108 * Y + 0.0415560 * Z;
    var B = 0.0556434 * X + -0.2040259 * Y + 1.0572252 * Z;

    // gamut判定（linearの段階で 0..1 に入っているか）
    var inGamut = (R >= 0 && R <= 1 && G >= 0 && G <= 1 && B >= 0 && B <= 1);

    // clampしてsRGBへ
    R = Math.min(1, Math.max(0, R));
    G = Math.min(1, Math.max(0, G));
    B = Math.min(1, Math.max(0, B));

    return { r: linearToSrgb(R), g: linearToSrgb(G), b: linearToSrgb(B), inGamut: inGamut };
}

function fLab(t) {
    return (t > 0.008856451679) ? Math.pow(t, 1 / 3) : (7.787037037 * t + 16 / 116);
}

function finvLab(t) {
    var t3 = t * t * t;
    return (t3 > 0.008856451679) ? t3 : ((t - 16 / 116) / 7.787037037);
}

function xyzToLab(xyz) {
    // D65 white
    var Xn = 95.047, Yn = 100.000, Zn = 108.883;

    var x = fLab(xyz.x / Xn);
    var y = fLab(xyz.y / Yn);
    var z = fLab(xyz.z / Zn);

    var L = 116 * y - 16;
    var a = 500 * (x - y);
    var b = 200 * (y - z);

    return { l: L, a: a, b: b };
}

function labToXyz(lab) {
    var Xn = 95.047, Yn = 100.000, Zn = 108.883;

    var fy = (lab.l + 16) / 116;
    var fx = fy + (lab.a / 500);
    var fz = fy - (lab.b / 200);

    var xr = finvLab(fx);
    var yr = finvLab(fy);
    var zr = finvLab(fz);

    return { x: xr * Xn, y: yr * Yn, z: zr * Zn };
}

function rgbToLab(r, g, b) {
    return xyzToLab(rgbToXyz(r, g, b));
}

function labToRgb(L, a, b) {
    var xyz = labToXyz({ l: L, a: a, b: b });
    return xyzToRgb(xyz.x, xyz.y, xyz.z);
}

function labToLch(lab) {
    var C = Math.sqrt(lab.a * lab.a + lab.b * lab.b);
    var H = Math.atan2(lab.b, lab.a) * 180 / Math.PI;
    if (H < 0) H += 360;
    return { l: lab.l, c: C, h: H };
}

function lchToLab(lch) {
    var hr = lch.h * Math.PI / 180;
    var a = lch.c * Math.cos(hr);
    var b = lch.c * Math.sin(hr);
    return { l: lch.l, a: a, b: b };
}

function hexToRgb(hex) {
    if (hex.charAt(0) === '#') hex = hex.substring(1);
    if (hex.length === 3) hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    var num = parseInt(hex, 16);
    return { r: num >> 16, g: (num >> 8) & 0x00FF, b: num & 0x0000FF };
}

function rgbToHsl(r, g, b) {
    r /= 255, g /= 255, b /= 255;
    var max = Math.max(r, g, b), min = Math.min(r, g, b);
    var h, s, l = (max + min) / 2;
    if (max == min) { h = s = 0; }
    else {
        var d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case r: h = (g - b) / d + (g < b ? 6 : 0); break;
            case g: h = (b - r) / d + 2; break;
            case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
    }
    return { h: h, s: s, l: l };
}

function hslToRgb(h, s, l) {
    var r, g, b;
    if (s == 0) { r = g = b = l; }
    else {
        function hue2rgb(p, q, t) {
            if (t < 0) t += 1; if (t > 1) t -= 1;
            if (t < 1 / 6) return p + (q - p) * 6 * t;
            if (t < 1 / 2) return q;
            if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        }
        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        r = hue2rgb(p, q, h + 1 / 3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1 / 3);
    }
    return { r: Math.round(r * 255), g: Math.round(g * 255), b: Math.round(b * 255) };
}

function drawToIllustrator(data, algoName, outHex, outRgb, contrastShift) {
    var doc = app.activeDocument;

    // Active artboard bounds: [left, top, right, bottom]
    var abIndex = doc.artboards.getActiveArtboardIndex();
    var abRect = doc.artboards[abIndex].artboardRect;
    var abLeft = abRect[0], abTop = abRect[1], abRight = abRect[2], abBottom = abRect[3];

    // 5mm マージン（pt換算）
    var m = 5 * 72 / 25.4;

    // inset rect
    var left = abLeft + m;
    var right = abRight - m;
    var top = abTop - m;
    var bottom = abBottom + m;

    var abW = right - left;
    var abH = top - bottom;

    var n = data.length;
    if (!n) return;

    // 見出し高さ（タイトル） / Header height
    var headerH = 24;
    if (abH < headerH + 5) headerH = 0;

    // スウォッチ高さ（従来どおり最大50pt） / Swatch height (cap 50pt)
    var swH = Math.min(50, abH - headerH);
    if (swH < 10) swH = 10;

    // 共通描画へ委譲 / Delegate to shared renderer
    drawPaletteRow(doc, left, top, abW, headerH, swH, data, algoName, outHex, outRgb, contrastShift);
}

// 「すべて」用：全アルゴリズムを縦に積んで描画
function drawAllToIllustrator(hsl0, rgb0, outHex, outRgb, contrastShift) {
    var doc = app.activeDocument;

    // Active artboard bounds: [left, top, right, bottom]
    var abIndex = doc.artboards.getActiveArtboardIndex();
    var abRect = doc.artboards[abIndex].artboardRect;
    var abLeft = abRect[0], abTop = abRect[1], abRight = abRect[2], abBottom = abRect[3];

    // 5mm マージン（pt換算）
    var m = 5 * 72 / 25.4;

    // inset rect
    var left = abLeft + m;
    var right = abRight - m;
    var top = abTop - m;
    var bottom = abBottom + m;

    var abW = right - left;
    var abH = top - bottom;

    // 描画順（Tailwind, Lightness, Lightness Geometric, LCH, Saturation, Complementary）
    // 「すべて」選択時はステップ数=11固定
    var ALL_STEPS = 11;

    var blocks = [
        { name: "Tailwind CSS", data: algoTailwind(hsl0) },
        { name: "Lightness Scale", data: algoLightness(hsl0, ALL_STEPS) },
        { name: "Lightness Scale (Geometric)", data: algoLightnessGeometric(hsl0, ALL_STEPS) },
        { name: "LCH", data: algoLchFromRgb(rgb0) },
        { name: "Saturation Scale", data: algoSaturation(hsl0, ALL_STEPS) },
        { name: "Complementary", data: algoComplementary(hsl0, ALL_STEPS) }
    ];

    var nBlocks = blocks.length;

    var headerH = 24;
    var gap = 10;

    // ラベル行数に応じた領域（実際に出る行数ベース）
    var labelLines = 0;
    if (outHex) labelLines += 1;
    if (outRgb) labelLines += 1;

    // 5pt文字 + 行送り7pt（上詰めの実装に合わせる）
    var lineGap = 7;
    var labelAreaH = 0;
    if (labelLines > 0) {
        // 最上行の分 + (残り行数-1)*lineGap + 下余白
        labelAreaH = 5 + ((labelLines - 1) * lineGap) + 6;
    }

    // 各ブロックのスウォッチ高さをアートボード内に収める
    var maxSwH = 50;
    var swH = Math.floor((abH - (nBlocks * headerH) - (nBlocks * labelAreaH) - ((nBlocks - 1) * gap)) / nBlocks);
    if (swH > maxSwH) swH = maxSwH;
    if (swH < 10) swH = 10;

    var yTop = top;

    for (var i = 0; i < nBlocks; i++) {
        var b = blocks[i];
        drawPaletteRow(doc, left, yTop, abW, headerH, swH, b.data, b.name, outHex, outRgb, contrastShift);
        yTop = yTop - (headerH + swH + labelAreaH + gap);
    }
}

// 共通描画：タイトル + スウォッチ矩形 + HEX/RGB ラベル（単体/すべて共通）
// Shared renderer for both single output and "All"
function drawPaletteRow(doc, left, top, width, headerH, swH, data, algoName, outHex, outRgb, contrastShift) {
    if (!data || !data.length) return;

    // --- グループ作成（矩形/テキスト） ---
    var rectGroup = doc.groupItems.add();
    var titleGroup = doc.groupItems.add();
    var hexGroup = doc.groupItems.add();
    var rgbGroup = doc.groupItems.add();

    // 前面順（RGB/HEX/Title を前に）
    try { titleGroup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }
    try { rgbGroup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }
    try { hexGroup.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) { }

    // タイトル / Title
    if (algoName && algoName.length && headerH > 0) {
        try {
            var tf = doc.textFrames.add();
            tf.contents = algoName;

            // position: [x, y] (yはベースライン)
            tf.position = [left, top - 8];

            // 左揃え
            try { tf.textRange.paragraphAttributes.justification = Justification.LEFT; } catch (_) { }

            // フォント設定：Avenir-Book / 11pt
            try {
                tf.textRange.characterAttributes.textFont = app.textFonts.getByName("Avenir-Book");
                tf.textRange.size = 11;
            } catch (_) { }

            // 念のため属性を正規化（最初の1個だけ崩れる対策）
            try {
                tf.textRange.characterAttributes.horizontalScale = 100;
                tf.textRange.characterAttributes.verticalScale = 100;
                tf.textRange.characterAttributes.tracking = 0;
                tf.textRange.characterAttributes.baselineShift = 0;
            } catch (_) { }

            // 塗り：黒
            try {
                var tcol = new RGBColor();
                tcol.red = 0; tcol.green = 0; tcol.blue = 0;
                tf.textRange.fillColor = tcol;
            } catch (_) { }

            // 見た目の左端を left に合わせる（フォント/サイズ確定後に補正）
            try {
                var gbT = tf.geometricBounds; // [left, top, right, bottom]
                var dxT = left - gbT[0];
                tf.position = [tf.position[0] + dxT, tf.position[1]];
            } catch (_) { }

            try { tf.move(titleGroup, ElementPlacement.PLACEATEND); } catch (_) { }
        } catch (_) { }
    }

    // テキスト共通ヘルパー / Text helpers
    function __setTextStyle(tf, sizePt, fontName) {
        try {
            tf.textRange.size = sizePt;
            var f = null;
            try { f = app.textFonts.getByName(fontName); } catch (_) { }
            if (!f) {
                try { f = app.textFonts.getByName("Avenir-Book"); } catch (__) { }
            }
            if (f) tf.textRange.characterAttributes.textFont = f;
        } catch (_) { }
        try {
            var tcol = new RGBColor();
            tcol.red = 0; tcol.green = 0; tcol.blue = 0;
            tf.textRange.fillColor = tcol;
        } catch (_) { }
    }

    function __addLabel(text, x, y, sizePt, parentGroup) {
        try {
            var tf = doc.textFrames.add();
            tf.contents = text;
            __setTextStyle(tf, sizePt, "Avenir-Book");
            tf.position = [x, y];
            try { tf.move(parentGroup, ElementPlacement.PLACEATEND); } catch (_) { }
            return tf;
        } catch (_) { }
        return null;
    }

    function __alignLeftToRect(tf, targetX) {
        try {
            var gb = tf.geometricBounds; // [left, top, right, bottom]
            var dx = targetX - gb[0];
            tf.position = [tf.position[0] + dx, tf.position[1]];
        } catch (_) { }
    }

    // 矩形位置 / Rectangles
    var n = data.length;
    var swW = width / n;
    var rectTop = top - headerH;

    for (var i = 0; i < n; i++) {
        var c0 = data[i];
        var c = applyContrastShiftToRgb(c0, contrastShift);
        var col = new RGBColor();
        col.red = c.r; col.green = c.g; col.blue = c.b;

        var rect = rectGroup.pathItems.rectangle(
            rectTop,
            left + (i * swW),
            swW,
            swH
        );
        rect.fillColor = col;
        rect.stroked = false;

        // ラベル（HEX / RGB） / Labels
        if (outHex || outRgb) {
            var x0 = left + (i * swW);
            var yBot2 = rectTop - swH; // bottom edge

            // 2行（HEX値 / RGB値）
            // ※HEXがOFFの場合はRGBを上に詰める
            var lineGap = 7; // 5pt想定の行送り
            var yTopLine = yBot2 - 2;
            var yHexVal = yTopLine;
            var yRgbVal = (outHex ? (yTopLine - lineGap) : yTopLine);

            if (outHex) {
                var hx = rgbToHexString(c.r, c.g, c.b);
                var tHex = __addLabel(hx, x0, yHexVal, 5, hexGroup);
                if (tHex) {
                    __setTextStyle(tHex, 5, "Avenir");
                    __alignLeftToRect(tHex, x0);
                }
            }

            if (outRgb) {
                var rgbText = "R" + Math.round(c.r) + " G" + Math.round(c.g) + " B" + Math.round(c.b);
                var tRgb = __addLabel(rgbText, x0, yRgbVal, 5, rgbGroup);
                if (tRgb) {
                    __setTextStyle(tRgb, 5, "Avenir-Book");
                    __alignLeftToRect(tRgb, x0);
                }
            }
        }
    }
}

function drawPaletteBlock(doc, left, top, width, headerH, swH, data, algoName, outHex, outRgb, contrastShift) {
    // 旧API互換：共通描画へ委譲 / Backward-compatible wrapper
    drawPaletteRow(doc, left, top, width, headerH, swH, data, algoName, outHex, outRgb, contrastShift);
}

main();
