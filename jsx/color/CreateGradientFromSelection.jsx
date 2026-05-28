#target illustrator
#targetengine "CreateGradientFromSelectionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# CreateGradientFromSelection.jsx

選択オブジェクトの塗り／線カラーを、配置順（左→右、上→下）で抽出し、
スウォッチグループに登録してグラデーションを自動生成するスクリプトです。

## 主な機能

- グループ／複合パス／テキストを再帰的に走査
- 塗り（フィル）と線（ストローク）の両方を対象
- 抽出色をスウォッチ化（オプションでグローバルカラー化）
- 線形グラデーションを生成し、各ストップにスウォッチ色を割り当て
- 長方形を作成してグラデーションを適用（任意）
- 長方形の見た目をグラフィックスタイルに登録（任意・長方形 OFF でも一時長方形で登録可）
- 選択の並びを判定し、縦並びならアクション（gradient/90degree）で角度を 90° に
- 「セパレートグラデーション」モード（2〜6 色、100÷色数で自動分割）
- 選択が 7 つ以上の場合はセパレートを無効化

## 既定の挙動

- ドキュメント無し／選択無し／色 1 色以下／例外発生時は無言で終了
- ダイアログ値は targetengine 内でセッション保持（再起動では消える）

Version: v1.9.2
更新日: 2026-05-28

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.9.2";

// =========================================
// ユーザー設定 / User Settings
// =========================================

/* セパレートグラデーションで許可する最大色数 / Max colors allowed for Separate gradients */
var SEPARATE_MAX_COLORS = 6;

/* 長方形サイズの既定値（pt）/ Default rectangle size in points */
var DEFAULT_RECT_SIZE = 100;

/* スウォッチ由来時の固定サイズ（pt）/ Fixed size when invoked from swatches */
var SWATCH_RECT_WIDTH = 200;
var SWATCH_RECT_HEIGHT = 100;

/* 新規スウォッチグループの既定名 / Default name for the new swatch group */
var SWATCH_GROUP_BASE_NAME = "AutoGradient";
var SWATCH_BASE_NAME = "AutoColor";
var GRADIENT_BASE_NAME = "New Gradient";

/* ダイアログのデフォルトオプション / Default dialog values */
var DEFAULT_OPTIONS = {
    makeGlobal: true,
    makeGradient: true,
    makeRect: true,
    useSelectionSize: true,
    registerGraphicStyle: true,
    separateGradient: false
};

// =========================================
// ローカライズ / Localization
// =========================================

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "グラデーション作成", en: "Create Gradient" }
    },
    panel: {
        color: { ja: "カラー", en: "Colors" },
        rect: { ja: "長方形", en: "Rectangle" }
    },
    checkbox: {
        globalColor: { ja: "グローバルカラー化", en: "Make Global colors" },
        createGradient: { ja: "グラデーションを作成", en: "Create gradient" },
        createRect: { ja: "長方形を作成してグラデーションを適用", en: "Create rectangle and apply gradient" },
        useSelectionSize: { ja: "選択オブジェクトのサイズに合わせる", en: "Match selection size" },
        registerGraphicStyle: { ja: "グラフィックスタイルとして登録", en: "Save as Graphic Style" }
    },
    radio: {
        normal: { ja: "通常", en: "Smooth" },
        separate: { ja: "セパレート", en: "Segmented" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        globalColor: {
            ja: "スウォッチをグローバルカラー（プロセス）として登録し、後から一括で色を変更可能にします。",
            en: "Register swatches as Global Process colors so they can be edited together later."
        },
        separate: {
            ja: "色の境界をくっきり分割するグラデーション（最大 6 色）。選択が 7 つ以上のときは使えません。",
            en: "Hard-edged segmented gradient (up to 6 colors). Disabled when 7 or more items are selected."
        },
        useSelectionSize: {
            ja: "選択オブジェクトの外接サイズに合わせて長方形を作成します。",
            en: "Use the bounding size of the selection for the rectangle."
        },
        registerGraphicStyle: {
            ja: "作成した長方形の見た目をグラフィックスタイルに登録します。長方形作成 OFF のときは一時長方形で登録します。",
            en: "Register the rectangle's appearance as a Graphic Style. When rectangle output is off, a temporary rectangle is used."
        }
    }
};

/* ドット区切りパスで多言語ラベルを取得 / Resolve a localized label by dot-path */
function L(path) {
    var parts = path.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        node = node && node[parts[i]];
    }
    if (node && node[lang]) return node[lang];
    if (node && node.en) return node.en;
    return path;
}

// =========================================
// セッション設定 / Session Settings
// =========================================

/* targetengine 内でダイアログ値を保持 / Persist dialog values within the targetengine */
function getSessionSettings() {
    if (!$.global.__CGFS_SETTINGS) $.global.__CGFS_SETTINGS = {};
    return $.global.__CGFS_SETTINGS;
}
function loadBool(key, defaultValue) {
    var settings = getSessionSettings();
    if (typeof settings[key] === 'boolean') return settings[key];
    return defaultValue;
}
function saveBool(key, value) {
    getSessionSettings()[key] = !!value;
}

// =========================================
// 選択範囲の解析 / Selection Analysis
// =========================================

/* アイテムの左上座標を取得 / Get an item's top-left position */
function getItemTopLeft(item) {
    try {
        var bounds = item.geometricBounds; // [left, top, right, bottom]
        return { left: bounds[0], top: bounds[1] };
    } catch (e) {
        return { left: 0, top: 0 };
    }
}

/* 選択全体の外接バウンディングを取得 / Get the union bounds of the selection */
function getSelectionBounds(selection) {
    if (!selection || selection.length === 0) return null;

    var left = 1e12, top = -1e12, right = -1e12, bottom = 1e12;
    var got = false;

    for (var i = 0; i < selection.length; i++) {
        try {
            var bounds = selection[i].geometricBounds;
            if (bounds[0] < left) left = bounds[0];
            if (bounds[1] > top) top = bounds[1];
            if (bounds[2] > right) right = bounds[2];
            if (bounds[3] < bottom) bottom = bounds[3];
            got = true;
        } catch (e) { /* 無視 / ignore */ }
    }

    if (!got || left > right || bottom > top) return null;
    return { left: left, top: top, right: right, bottom: bottom };
}

/* 選択が横並びか縦並びかを判定 / Detect whether the selection is horizontal or vertical */
function detectSelectionOrientation(selection) {
    if (!selection || selection.length < 2) {
        return { orientation: "unknown", dx: 0, dy: 0, ratio: 0 };
    }

    var minX = 1e12, maxX = -1e12;
    var minY = 1e12, maxY = -1e12;

    for (var i = 0; i < selection.length; i++) {
        try {
            var bounds = selection[i].geometricBounds;
            var centerX = (bounds[0] + bounds[2]) / 2;
            var centerY = (bounds[1] + bounds[3]) / 2;
            if (centerX < minX) minX = centerX;
            if (centerX > maxX) maxX = centerX;
            if (centerY < minY) minY = centerY;
            if (centerY > maxY) maxY = centerY;
        } catch (e) { /* 無視 / ignore */ }
    }

    if (minX > maxX || minY > maxY) {
        return { orientation: "unknown", dx: 0, dy: 0, ratio: 0 };
    }

    var dx = Math.abs(maxX - minX);
    var dy = Math.abs(maxY - minY);

    var ratio = 0;
    if (dx === 0 && dy === 0) ratio = 0;
    else if (dx === 0 || dy === 0) ratio = 1e12;
    else ratio = (dx > dy) ? (dx / dy) : (dy / dx);

    var orientation = "mixed";
    if (dx > dy) orientation = "horizontal";
    else if (dy > dx) orientation = "vertical";

    return { orientation: orientation, dx: dx, dy: dy, ratio: ratio };
}

// =========================================
// カラーユーティリティ / Color Utilities
// =========================================

/* NoColor 判定 / Detect NoColor values */
function isNoColor(color) {
    try {
        return (color == null) || (color.typename === "NoColor");
    } catch (e) {
        return true;
    }
}

/* 重複除去用のカラーキーを生成 / Build a dedup key for a color */
function colorKey(color) {
    if (!color) return "null";
    var typeName = color.typename;
    try {
        if (typeName === "RGBColor") return "RGB:" + [color.red, color.green, color.blue].join(",");
        if (typeName === "CMYKColor") return "CMYK:" + [color.cyan, color.magenta, color.yellow, color.black].join(",");
        if (typeName === "GrayColor") return "Gray:" + color.gray;
        if (typeName === "SpotColor") {
            var spotName = (color.spot && color.spot.name) ? color.spot.name : "(spot)";
            return "Spot:" + spotName + ":" + color.tint;
        }
        if (typeName === "PatternColor") {
            var patternName = (color.pattern && color.pattern.name) ? color.pattern.name : "(pattern)";
            return "Pattern:" + patternName;
        }
        if (typeName === "GradientColor") {
            var gradientName = (color.gradient && color.gradient.name) ? color.gradient.name : "(gradient)";
            return "Gradient:" + gradientName;
        }
    } catch (e) { /* 無視 / ignore */ }
    return "Other:" + typeName;
}

/* アイテムから塗り・線の色＋位置エントリを収集（再帰） / Collect fill and stroke color entries with positions (recursive) */
function collectFillColorEntries(item, outEntries) {
    if (!item) return;

    try {
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectFillColorEntries(item.pageItems[i], outEntries);
            }
            return;
        }

        if (item.typename === "CompoundPathItem") {
            for (var j = 0; j < item.pathItems.length; j++) {
                collectFillColorEntries(item.pathItems[j], outEntries);
            }
            return;
        }

        if (item.typename === "TextFrame") {
            var textPos = getItemTopLeft(item);
            var textFill = item.textRange.characterAttributes.fillColor;
            if (!isNoColor(textFill)) outEntries.push({ left: textPos.left, top: textPos.top, color: textFill });
            try {
                var textStroke = item.textRange.characterAttributes.strokeColor;
                if (!isNoColor(textStroke)) outEntries.push({ left: textPos.left, top: textPos.top, color: textStroke });
            } catch (e) { /* 無視 / ignore */ }
            return;
        }

        if (typeof item.filled !== "undefined" || typeof item.stroked !== "undefined") {
            var pathPos = getItemTopLeft(item);
            if (item.filled) {
                var pathFill = item.fillColor;
                if (!isNoColor(pathFill)) outEntries.push({ left: pathPos.left, top: pathPos.top, color: pathFill });
            }
            if (item.stroked) {
                var pathStroke = item.strokeColor;
                if (!isNoColor(pathStroke)) outEntries.push({ left: pathPos.left, top: pathPos.top, color: pathStroke });
            }
        }
    } catch (e) { /* 取得できないアイテムは無視 / skip unreadable items */ }
}

/* 選択範囲から重複除外したカラー配列を、長辺方向（横なら左→右／縦なら上→下）順で返す / Collect unique colors sorted along the longer axis of the selection */
function collectColorsFromSelection(selection, orientation) {
    var entries = [];
    for (var i = 0; i < selection.length; i++) {
        collectFillColorEntries(selection[i], entries);
    }

    /* 縦並び（dy > dx）なら上→下を優先キーに / Use top→bottom as primary key when vertical */
    var verticalPrimary = !!(orientation && orientation.orientation === "vertical");

    entries.sort(function (a, b) {
        if (verticalPrimary) {
            if (a.top > b.top) return -1;
            if (a.top < b.top) return 1;
            if (a.left < b.left) return -1;
            if (a.left > b.left) return 1;
            return 0;
        }
        if (a.left < b.left) return -1;
        if (a.left > b.left) return 1;
        if (a.top > b.top) return -1;
        if (a.top < b.top) return 1;
        return 0;
    });

    var colors = [];
    var seen = {};
    for (var k = 0; k < entries.length; k++) {
        var color = entries[k].color;
        if (isNoColor(color)) continue;
        var key = colorKey(color);
        if (seen[key]) continue;
        seen[key] = true;
        colors.push(color);
    }
    return colors;
}

// =========================================
// アクション定義 / Action Definitions
// =========================================

/* 一時アクションをロードして実行し、後始末を行う共通処理 / Load a temp action, run it, then clean up */
function runTempAction(actionCode, actionSetName, actionName) {
    var tempFile = new File(Folder.temp + "/temp_action_" + actionSetName + ".aia");
    try {
        tempFile.open("w");
        tempFile.write(actionCode);
        tempFile.close();

        app.loadAction(tempFile);
        app.doScript(actionName, actionSetName);
        app.unloadAction(actionSetName, "");

        try { tempFile.remove(); } catch (e) { /* 無視 / ignore */ }
    } catch (e) {
        try { app.unloadAction(actionSetName, ""); } catch (e2) { /* 無視 / ignore */ }
    }
}

/* グラデーション角度を 90° に設定するアクションを実行 / Run an action that sets the gradient angle to 90° */
function runGradientAngle90Action() {
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
    runTempAction(actionCode, "gradient", "90degree");
}

/* 「新規グラフィックスタイル」を呼び出すアクション / Trigger the "New Graphic Style" command via action */
function runGraphicStyleAction() {
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
    runTempAction(actionCode, "GraphicStyle", "new");
}

// =========================================
// スウォッチ操作 / Swatch Operations
// =========================================

/* 命名衝突を避けた一意な名前を作る / Build a unique name avoiding collisions */
function uniqueName(baseName, existsFunc) {
    var name = baseName;
    var suffix = 1;
    while (existsFunc(name)) {
        name = baseName + " " + suffix;
        suffix++;
    }
    return name;
}

function swatchExists(doc, name) {
    try { doc.swatches.getByName(name); return true; } catch (e) { return false; }
}

function swatchGroupExists(doc, name) {
    try { doc.swatchGroups.getByName(name); return true; } catch (e) { return false; }
}

function gradientExists(doc, name) {
    try { doc.gradients.getByName(name); return true; } catch (e) { return false; }
}

/* ベースカラーをグローバル（プロセス）スポットに変換 / Convert a base color into a Global Process spot */
function toGlobalProcessColor(doc, baseColor, baseName) {
    try {
        var spot = doc.spots.add();
        spot.name = baseName;
        spot.colorType = ColorModel.PROCESS;
        spot.color = baseColor;

        var spotColor = new SpotColor();
        spotColor.spot = spot;
        spotColor.tint = 100;
        return spotColor;
    } catch (e) {
        return baseColor;
    }
}

/* 1 色をスウォッチに登録（必要に応じてグローバル化） / Register a color as a swatch (optionally as Global Process) */
function addSwatchForColor(doc, colorObj, baseName, makeGlobal) {
    var swatch = doc.swatches.add();
    var name = uniqueName(baseName, function (n) { return swatchExists(doc, n); });
    swatch.name = name;
    swatch.color = makeGlobal ? toGlobalProcessColor(doc, colorObj, name) : colorObj;
    try { swatch.selected = false; } catch (e) { /* 無視 / ignore */ }
    return swatch;
}

// =========================================
// レイヤー操作 / Layer Helpers
// =========================================

/* 描画可能なレイヤーを返す（activeLayer 優先） / Return a drawable layer, preferring activeLayer */
function getUnlockedVisibleLayer(doc) {
    var activeLayer = doc.activeLayer;
    if (activeLayer && !activeLayer.locked && activeLayer.visible) return activeLayer;
    for (var i = 0; i < doc.layers.length; i++) {
        var layer = doc.layers[i];
        if (!layer.locked && layer.visible) return layer;
    }
    return null;
}

// =========================================
// グラフィックスタイル登録 / Graphic Style Registration
// =========================================

/* 選択中アイテムをグラフィックスタイルとして登録 / Register the selected item's appearance as a Graphic Style */
function registerGraphicStyleFromSelected(doc) {
    if (!doc.graphicStyles) return;

    var selectedItems = [];
    if (doc.selection && doc.selection.length) {
        for (var i = 0; i < doc.selection.length; i++) selectedItems.push(doc.selection[i]);
    }
    if (!selectedItems.length) return;

    for (var k = 0; k < selectedItems.length; k++) {
        try {
            var beforeLen = doc.graphicStyles.length;
            try { doc.selection = null; } catch (e) { /* 無視 / ignore */ }
            try { doc.selection = [selectedItems[k]]; }
            catch (e) { try { selectedItems[k].selected = true; } catch (e2) { /* 無視 / ignore */ } }

            runGraphicStyleAction();

            var afterLen = doc.graphicStyles.length;
            if (afterLen <= beforeLen) continue;
            // 既定名のまま使用 / leave the default name
        } catch (eEach) { /* 1 件失敗しても続行 / keep going on per-item failure */ }
    }
}

// =========================================
// 入力収集 / Input Collection
// =========================================

/* 選択オブジェクトかスウォッチ選択から、色配列と関連情報を集める / Gather colors plus context from object or swatch selection */
function collectInputColors(doc) {
    var result = {
        colors: [],
        fromSwatches: false,
        selectionBounds: null,
        selectionOrientation: { orientation: "unknown", dx: 0, dy: 0, ratio: 0 },
        itemCount: 0
    };

    if (doc.selection && doc.selection.length > 0) {
        result.selectionOrientation = detectSelectionOrientation(doc.selection);
        result.colors = collectColorsFromSelection(doc.selection, result.selectionOrientation);
        result.selectionBounds = getSelectionBounds(doc.selection);
        result.itemCount = doc.selection.length;
        return result;
    }

    var selectedSwatches = null;
    try { selectedSwatches = doc.swatches.getSelected(); } catch (e) { selectedSwatches = null; }
    if (!selectedSwatches || selectedSwatches.length < 2) return result;

    result.fromSwatches = true;
    for (var i = 0; i < selectedSwatches.length; i++) {
        try {
            var swatchColor = selectedSwatches[i].color;
            if (swatchColor && swatchColor.typename !== "NoColor") result.colors.push(swatchColor);
        } catch (e) { /* 無視 / ignore */ }
    }
    result.itemCount = result.colors.length;
    return result;
}

// =========================================
// ダイアログ / Options Dialog
// =========================================

/* オプションダイアログを表示し、確定値を返す（キャンセル時は null） / Show options dialog; return resolved values or null on cancel */
function showOptionsDialog(disallowSeparate, fromSwatches) {
    var opts = {
        makeGlobal: loadBool('makeGlobal', DEFAULT_OPTIONS.makeGlobal),
        makeGradient: loadBool('makeGradient', DEFAULT_OPTIONS.makeGradient),
        makeRect: loadBool('makeRect', DEFAULT_OPTIONS.makeRect),
        useSelectionSize: loadBool('useSelectionSize', DEFAULT_OPTIONS.useSelectionSize),
        registerGraphicStyle: loadBool('registerGraphicStyle', DEFAULT_OPTIONS.registerGraphicStyle),
        separateGradient: (disallowSeparate ? false : loadBool('separateGradient', DEFAULT_OPTIONS.separateGradient))
    };

    var dlg = new Window('dialog', L('dialog.title') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    /* カラー関連パネル / Color-related panel */
    var colorPanel = dlg.add('panel', undefined, L('panel.color'));
    colorPanel.orientation = 'column';
    colorPanel.alignChildren = ['fill', 'top'];
    colorPanel.margins = [15, 20, 15, 10];

    var cbGlobal = colorPanel.add('checkbox', undefined, L('checkbox.globalColor'));
    cbGlobal.value = opts.makeGlobal;
    cbGlobal.helpTip = L('tooltip.globalColor');

    var cbGradient = colorPanel.add('checkbox', undefined, L('checkbox.createGradient'));
    cbGradient.value = opts.makeGradient;

    var radioGroup = colorPanel.add('group');
    radioGroup.orientation = 'row';
    radioGroup.alignChildren = ['left', 'center'];

    var rbNormal = radioGroup.add('radiobutton', undefined, L('radio.normal'));
    var rbSeparate = radioGroup.add('radiobutton', undefined, L('radio.separate'));
    rbSeparate.helpTip = L('tooltip.separate');
    rbSeparate.value = !!opts.separateGradient;
    rbNormal.value = !rbSeparate.value;

    /* 長方形パネル / Rectangle panel */
    var rectPanel = dlg.add('panel', undefined, L('panel.rect'));
    rectPanel.orientation = 'column';
    rectPanel.alignChildren = ['fill', 'top'];
    rectPanel.margins = [15, 20, 15, 10];

    var cbRect = rectPanel.add('checkbox', undefined, L('checkbox.createRect'));
    cbRect.value = opts.makeRect;

    var cbSelSize = rectPanel.add('checkbox', undefined, L('checkbox.useSelectionSize'));
    cbSelSize.value = opts.useSelectionSize;
    cbSelSize.helpTip = L('tooltip.useSelectionSize');

    var cbGStyle = rectPanel.add('checkbox', undefined, L('checkbox.registerGraphicStyle'));
    cbGStyle.value = opts.registerGraphicStyle;
    cbGStyle.helpTip = L('tooltip.registerGraphicStyle');

    /* チェック状態の連動 / Sync enabled state across controls */
    function syncEnable() {
        cbRect.enabled = cbGradient.value;
        cbSelSize.enabled = cbGradient.value && cbRect.value && !fromSwatches;
        cbGStyle.enabled = cbGradient.value;

        rbNormal.enabled = cbGradient.value;
        rbSeparate.enabled = cbGradient.value && !disallowSeparate;
        radioGroup.enabled = cbGradient.value;

        if (!cbGradient.value) {
            cbRect.value = false;
            cbSelSize.value = false;
            cbGStyle.value = false;
        }
        if (fromSwatches) cbSelSize.value = false;
        if (!cbGradient.value || disallowSeparate) {
            rbSeparate.value = false;
            rbNormal.value = true;
        }
    }
    cbGradient.onClick = syncEnable;
    cbRect.onClick = syncEnable;
    syncEnable();

    /* OK／キャンセル / OK and Cancel */
    var buttonGroup = dlg.add('group');
    buttonGroup.alignment = 'right';
    buttonGroup.add('button', undefined, L('button.cancel'), { name: 'cancel' });
    buttonGroup.add('button', undefined, 'OK', { name: 'ok' });

    function persistFromUI() {
        saveBool('makeGlobal', cbGlobal.value);
        saveBool('makeGradient', cbGradient.value);
        saveBool('makeRect', cbRect.value);
        saveBool('useSelectionSize', cbSelSize.value);
        saveBool('registerGraphicStyle', cbGStyle.value);
        saveBool('separateGradient', (disallowSeparate ? false : rbSeparate.value));
    }
    dlg.onClose = function () { try { persistFromUI(); } catch (e) { /* 無視 / ignore */ } };

    if (dlg.show() !== 1) return null;

    return {
        makeGlobal: !!cbGlobal.value,
        makeGradient: !!cbGradient.value,
        makeRect: !!cbRect.value,
        useSelectionSize: !!cbSelSize.value,
        registerGraphicStyle: !!cbGStyle.value,
        separateGradient: (disallowSeparate ? false : !!rbSeparate.value)
    };
}

// =========================================
// グラデーション生成 / Gradient Construction
// =========================================

/* 作成済みスウォッチがあればそれを、無ければ元色を返す / Prefer the created swatch's color over the raw input color */
function pickStopColor(createdSwatches, colors, index) {
    try {
        if (createdSwatches && createdSwatches[index] && createdSwatches[index].color) {
            return createdSwatches[index].color;
        }
    } catch (e) { /* 無視 / ignore */ }
    return colors[index];
}

/* グラデーションのストップ数を target に合わせる / Resize the gradient stop count to match target */
function resizeGradientStops(gradient, target) {
    while (gradient.gradientStops.length < target) gradient.gradientStops.add();
    while (gradient.gradientStops.length > target) {
        gradient.gradientStops[gradient.gradientStops.length - 1].remove();
    }
}

/* 通常（スムーズ）グラデーションを作成 / Build a smooth gradient evenly spaced across stops */
function buildNormalGradient(doc, colors, createdSwatches) {
    var gradient = doc.gradients.add();
    gradient.type = GradientType.LINEAR;
    resizeGradientStops(gradient, colors.length);

    for (var i = 0; i < colors.length; i++) {
        var stop = gradient.gradientStops[i];
        stop.rampPoint = (i / (colors.length - 1)) * 100;
        try { stop.color = pickStopColor(createdSwatches, colors, i); }
        catch (e) { try { stop.color = colors[i]; } catch (e2) { /* 無視 / ignore */ } }
        stop.midPoint = 50;
        stop.opacity = 100;
    }

    gradient.name = uniqueName(GRADIENT_BASE_NAME, function (n) { return gradientExists(doc, n); });
    return gradient;
}

/* セパレート（境界がくっきり）グラデーションを作成（2〜SEPARATE_MAX_COLORS 色） / Build a segmented gradient with hard edges */
function buildSeparateGradient(doc, colors, createdSwatches) {
    var n = colors.length;
    var epsilon = 0.01;
    var step = 100 / n;
    if (n === 3 || n === 6) step = Math.round(step * 10) / 10;

    var stopPoints = [0];
    var stopColors = [pickStopColor(createdSwatches, colors, 0)];

    for (var k = 1; k <= n - 1; k++) {
        var boundary = step * k;
        if (n === 3 || n === 6) boundary = Math.round(boundary * 10) / 10;
        var leftPoint = Math.max(0, boundary - epsilon);
        var rightPoint = Math.min(100, boundary + epsilon);

        stopPoints.push(leftPoint);
        stopColors.push(pickStopColor(createdSwatches, colors, k - 1));
        stopPoints.push(rightPoint);
        stopColors.push(pickStopColor(createdSwatches, colors, k));
    }

    stopPoints.push(100);
    stopColors.push(pickStopColor(createdSwatches, colors, n - 1));

    var gradient = doc.gradients.add();
    gradient.type = GradientType.LINEAR;
    resizeGradientStops(gradient, stopPoints.length);

    for (var i = 0; i < stopPoints.length; i++) {
        var stop = gradient.gradientStops[i];
        stop.rampPoint = stopPoints[i];
        try { stop.color = stopColors[i]; }
        catch (e) {
            try { stop.color = colors[Math.min(colors.length - 1, Math.max(0, Math.floor(i / 2)))]; } catch (e2) { /* 無視 / ignore */ }
        }
        stop.midPoint = 50;
        stop.opacity = 100;
    }
    return gradient;
}

// =========================================
// 長方形配置 / Rectangle Placement
// =========================================

/* 長方形サイズを決定 / Decide the rectangle size */
function computeRectSize(opts, input) {
    if (input.fromSwatches) return { width: SWATCH_RECT_WIDTH, height: SWATCH_RECT_HEIGHT };
    if (opts.useSelectionSize && input.selectionBounds) {
        var w = Math.abs(input.selectionBounds.right - input.selectionBounds.left);
        var h = Math.abs(input.selectionBounds.top - input.selectionBounds.bottom);
        if (w > 0 && h > 0) return { width: w, height: h };
    }
    return { width: DEFAULT_RECT_SIZE, height: DEFAULT_RECT_SIZE };
}

/* 長方形の配置（左上座標）を決定 / Decide the rectangle anchor (top-left) */
function computeRectPosition(doc, input, size) {
    var viewCenterX = doc.activeView.centerPoint[0];
    var viewCenterY = doc.activeView.centerPoint[1];
    var left = viewCenterX - size.width / 2;
    var top = viewCenterY + size.height / 2;

    if (!input.fromSwatches && input.selectionBounds) {
        if (input.selectionOrientation.orientation === "horizontal") {
            /* 横並び: 選択の左端揃え／真下に 1 個分離す / Horizontal: align to left edge, offset below */
            left = input.selectionBounds.left;
            top = input.selectionBounds.bottom - size.height;
        } else if (input.selectionOrientation.orientation === "vertical") {
            /* 縦並び: 選択の上端揃え／右に 1 個分離す / Vertical: align to top edge, offset to right */
            left = input.selectionBounds.right + size.width;
            top = input.selectionBounds.top;
        }
    }
    return { left: left, top: top };
}

/* 長方形を作成し、グラデーション適用（必要に応じてスタイル登録）を行う / Create a rectangle, apply gradient, optionally register a style */
function createGradientRect(doc, gradient, opts, input) {
    var targetLayer = getUnlockedVisibleLayer(doc);
    if (!targetLayer) return;

    var tempRectForStyle = (!opts.makeRect && opts.registerGraphicStyle);
    var prevSelection = null;
    var prevActiveLayer = null;
    var tempLayer = null;

    try { prevActiveLayer = doc.activeLayer; } catch (e) { /* 無視 / ignore */ }
    if (doc.selection && doc.selection.length) {
        prevSelection = [];
        for (var i = 0; i < doc.selection.length; i++) prevSelection.push(doc.selection[i]);
    }

    if (tempRectForStyle) {
        try {
            tempLayer = doc.layers.add();
            tempLayer.name = "__TempGraphicStyle";
            doc.activeLayer = tempLayer;
        } catch (e) { tempLayer = null; }
    }

    var drawLayer = (tempRectForStyle && tempLayer) ? tempLayer : targetLayer;
    var size = computeRectSize(opts, input);
    var pos = computeRectPosition(doc, input, size);

    var rect = drawLayer.pathItems.rectangle(pos.top, pos.left, size.width, size.height);
    doc.selection = null;
    rect.selected = true;

    /* 縦並びならグラデーション角度を 90° に / If vertical, rotate gradient by action */
    if (input.selectionOrientation.orientation === "vertical") {
        try { runGradientAngle90Action(); } catch (e) { /* 無視 / ignore */ }
    }

    rect.stroked = false;
    rect.filled = true;
    var gradientFill = new GradientColor();
    gradientFill.gradient = gradient;
    rect.fillColor = gradientFill;

    if (opts.registerGraphicStyle) {
        doc.selection = null;
        rect.selected = true;
        try { registerGraphicStyleFromSelected(doc); } catch (e) { /* 無視 / ignore */ }
    }

    /* 一時長方形だった場合の後始末 / Clean up the temporary rectangle */
    if (tempRectForStyle) {
        try { rect.remove(); } catch (e) { /* 無視 / ignore */ }
        if (tempLayer) { try { tempLayer.remove(); } catch (e) { /* 無視 / ignore */ } }
        try { if (prevActiveLayer) doc.activeLayer = prevActiveLayer; } catch (e) { /* 無視 / ignore */ }
        try {
            doc.selection = null;
            if (prevSelection && prevSelection.length) doc.selection = prevSelection;
        } catch (e) { /* 無視 / ignore */ }
    }
}

// =========================================
// メイン処理 / Main
// =========================================

/* 全体フロー: 入力 → ダイアログ → スウォッチ登録 → グラデーション → 長方形・スタイル / Top-level flow */
function main() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;

    var input = collectInputColors(doc);
    if (input.colors.length < 2) return;

    var disallowSeparate = (input.itemCount >= 7);
    var opts = showOptionsDialog(disallowSeparate, input.fromSwatches);
    if (!opts) return;

    try {
        /* 新規スウォッチグループ / Create a new swatch group */
        var groupName = uniqueName(SWATCH_GROUP_BASE_NAME, function (n) { return swatchGroupExists(doc, n); });
        var swatchGroup = doc.swatchGroups.add();
        swatchGroup.name = groupName;

        /* 抽出色をスウォッチに登録 / Register extracted colors as swatches */
        var createdSwatches = [];
        for (var i = 0; i < input.colors.length; i++) {
            var swatch = addSwatchForColor(doc, input.colors[i], SWATCH_BASE_NAME, opts.makeGlobal);
            createdSwatches.push(swatch);
            try { swatchGroup.addSwatch(swatch); } catch (e) { /* 無視 / ignore */ }
        }

        try { doc.selection = null; } catch (e) { /* 無視 / ignore */ }

        /* グラデーション作成 / Build the gradient */
        var gradient = null;
        if (opts.makeGradient) {
            var canSeparate = opts.separateGradient
                && input.colors.length >= 2
                && input.colors.length <= SEPARATE_MAX_COLORS;
            gradient = canSeparate
                ? buildSeparateGradient(doc, input.colors, createdSwatches)
                : buildNormalGradient(doc, input.colors, createdSwatches);
        }

        /* 長方形・グラフィックスタイル / Rectangle and Graphic Style */
        if ((opts.makeRect || opts.registerGraphicStyle) && gradient) {
            try { createGradientRect(doc, gradient, opts, input); } catch (e) { /* 無視 / ignore */ }
        }

        /* 作成した最後のスウォッチ（= グラデーション）を選択 / Select the last created swatch */
        if (gradient) {
            var idx = doc.swatches.length - 1;
            if (idx >= 0) {
                try { doc.swatches[idx].selected = true; } catch (e) { /* 無視 / ignore */ }
            }
        }
    } catch (e) {
        /* 無言で終了 / silent */
    }
}

main();