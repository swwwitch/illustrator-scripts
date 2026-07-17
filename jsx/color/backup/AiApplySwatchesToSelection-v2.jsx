#target illustrator

/*

### 概要

- 選択したオブジェクトやテキストに、スウォッチや定義済みカラーを適用するモーダルダイアログです。
- 適用単位（オブジェクト／1文字／単語／行／段落）と適用順（そのまま／逆順／ランダム）をラジオボタンで選択します。
- ラジオを変えるたびにライブプレビュー。［OK］で確定、［キャンセル］で元に戻します（元の塗り/線/不透明度をスナップショットから復元）。
- 取り込み元は「選択しているスウォッチ／選択しているオブジェクト／自動生成」のラジオで切り替えます（塗り色の黒・白は除外）。
- 実行時に選択していたスウォッチを■で自動取り込み、それ（スウォッチ名で参照）を最優先で使用。1色でも選択があれば優先します。
- 「自動生成」は CMYK で CM/CY/MY のみのランダム配色、RGB では既定色を使用します。
- 「単語ごと」は各行の先頭色が互い違いになるよう配色。
- モーダルダイアログなので DOM はメインエンジンで直接操作します（BridgeTalk 委譲なし）。

### Overview

- A modal dialog that applies swatches or predefined colors to selected objects or text.
- Choose apply unit (object / character / word / line / paragraph) and order (as-is / reverse / random) with radio buttons.
- Live preview on every change. [OK] commits, [Cancel] reverts (fills/strokes/opacity are restored from a snapshot).
- Switch the source between "Selected swatches", "Selected objects" and "Auto-generate" with radios (black/white fills are excluded).
- The swatches selected at launch are auto-captured as chips and used (by swatch name) with top priority. Even a single selected swatch is used.
- "Auto-generate" uses CM/CY/MY-only random colors for CMYK, or defaults for RGB.
- "Per word" staggers colors so each line starts on a different color.
- Being a modal dialog, DOM work runs directly in the main engine (no BridgeTalk delegation).

*/

(function () {

app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.7.1";

// =========================================
// ユーザー設定 / User settings
// =========================================
/* CMYK フォールバック生成の設定 / CMYK fallback generation settings */
var TMK_CMYK_FALLBACK_MAX_TOTAL = 200;    /* 合計上限 C+M / C+Y / M+Y / total channel limit */
var TMK_CMYK_FALLBACK_MIN_DISTANCE = 35;  /* 近似色回避の最小距離 / min distance to avoid similar colors */
var COLOR_CHIPS_PER_ROW = 12;             /* 1行あたりのチップ数 / chips per row */
var CHIP_SIZE = 18;                       /* チップの一辺(px) / chip size */
var CHIP_GAP = 4;                         /* チップ間の間隔(px) / gap between chips */

// =========================================
// ローカライズ / Localization
// =========================================
/* 現在の UI 言語を判定 / Detect current UI language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* ラベル定義（カテゴリ分け）/ Label definitions (categorized) */
var LABELS = {
    dialog: {
        title: { ja: "カラーを配色", en: "Distribute Colors" }
    },
    panel: {
        unit: { ja: "配色単位", en: "Distribute By" },
        colors: { ja: "使用するカラー", en: "Colors to Use" },
        option: { ja: "配色順", en: "Order" }
    },
    unit: {
        object:    { ja: "オブジェクトごと", en: "Per object" },
        character: { ja: "1文字ごと", en: "Per character" },
        word:      { ja: "単語ごと", en: "Per word" },
        line:      { ja: "行ごと", en: "Per line" },
        paragraph: { ja: "段落ごと", en: "Per paragraph" }
    },
    order: {
        asis:    { ja: "取り込み順", en: "As captured" },
        reverse: { ja: "逆順", en: "Reverse" },
        random:  { ja: "ランダム", en: "Random" }
    },
    source: {
        swatches: { ja: "選択しているスウォッチ", en: "Selected swatches" },
        objects:  { ja: "選択しているオブジェクト", en: "Selected objects" },
        auto:     { ja: "自動生成", en: "Auto-generate" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        sourceSwatches: { ja: "実行時に選択していたスウォッチの色を使用します。", en: "Use the colors of the swatches selected at launch." },
        sourceObjects:  { ja: "選択しているオブジェクトの塗り色を使用します（黒・白は除外）。", en: "Use the fill colors of the selected objects (black and white excluded)." },
        sourceAuto:     { ja: "自動でカラーを生成します（CMYK は CM/CY/MY のランダム、RGB は既定色）。", en: "Generate colors automatically (random CM/CY/MY for CMYK, defaults for RGB)." },
        colors:        { ja: "適用に使う配色。下のラジオで取り込み元を切り替えます。", en: "Colors used for applying; switch the source with the radios below." },
        unitObject:    { ja: "選択オブジェクト単位でカラーを順番に適用します。", en: "Applies colors in order, one per selected object." },
        unitCharacter: { ja: "テキストを1文字ずつ分けてカラーを適用します。", en: "Applies a color to each character of the text." },
        unitWord:      { ja: "英文テキストを単語ごとに分けてカラーを適用します。英文以外では単語が正しく分割されないことがあります。", en: "Applies a color to each word of English text. Words may not split correctly for non-English text." },
        unitLine:      { ja: "テキストの行ごとにカラーを適用します。", en: "Applies a color to each line of the text." },
        unitParagraph: { ja: "テキストの段落ごとにカラーを適用します。", en: "Applies a color to each paragraph of the text." },
        orderAsis:     { ja: "取り込んだカラーの並び順で適用します（適用先は位置順・文字順）。", en: "Applies colors in the captured order (targets follow position / reading order)." },
        orderReverse:  { ja: "取り込んだカラーの並びを逆にして適用します。", en: "Applies the captured colors in reverse order." },
        orderRandom:   { ja: "取り込んだカラーの並びをランダムにして適用します。", en: "Applies the captured colors in a random order." }
    },
    status: {
        noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
    },
    note: {
        autoColor: {
            ja: "自動生成のカラーを使用します",
            en: "Using auto-generated colors"
        }
    }
};

/* ドット区切りキーでローカライズ文字列を取得 / Get a localized string by dotted key path */
function L(path) {
    var parts = path.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        node = node[parts[i]];
        if (node == null) return path;
    }
    return node[currentLanguage];
}

// =========================================
// UIレイアウト共通設定 / Shared UI layout
// =========================================
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/* ウィンドウの共通設定 / Apply shared window layout */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* 行グループの共通設定（ボタン列など）/ Apply a horizontal row group */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// モーダルダイアログ / Modal dialog
// =========================================
/* ダイアログを表示（ライブプレビュー＋OK/キャンセル）/ Show the modal dialog (live preview + OK/Cancel) */
function showDialog() {
    if (app.documents.length === 0) { alert(L("status.noDoc")); return; }

    /* 選択情報を取得（初期単位・チップ・単語ディム判定）/ Read selection info (initial unit, chips, dims) */
    var info = readSelectionInfo();

    var win = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(win);

    /* 取り込んだ配色（適用に使用）。スウォッチ名 or シリアライズ済みカラー値のどちらか一方を保持
       Captured colors for applying: swatch names OR serialized values (either one) */
    var loadedSwatchNames = info.swatchNames;
    var loadedColorValues = [];

    /* 直前のプレビューで書き換えた塗り/線/不透明度の元の状態（キャンセル・再適用時にここから復元）
       Originals changed by the previous preview; restored from here on cancel / re-apply (null = none applied) */
    var previewSnapshot = null;
    /* 直前に適用した設定の署名（ラジオの二重発火で同一設定が連続適用されるのを防ぐ）
       Signature of the last applied settings (guards against duplicate applies from radio double-fire) */
    var lastPreviewSig = null;
    /* 直前のカラー取り込み元（取り込み元ラジオの二重発火対策）/ Last color source (guards its radio double-fire) */
    var lastSource = "swatches";
    /* OK 確定フラグ（onClose で戻すかどうか）/ Commit flag (whether onClose reverts) */
    var committed = false;

    /* 適用するカラー（配色チップ）＋取り込み元ラジオ / Color chips + source radios */
    var colorPanel = win.add("panel", undefined, L("panel.colors"));
    setupPanel(colorPanel);
    colorPanel.helpTip = L("tooltip.colors");
    var chipHost = colorPanel.add("group");
    chipHost.orientation = "column";
    chipHost.alignChildren = ["left", "top"];
    buildColorChips(chipHost, info.chips);
    /* カラーの取り込み元をラジオで切替（縦並び）。スウォッチは実行時に自動読込
       Color source radios (vertical); swatches auto-load on launch */
    var sourceGroup = colorPanel.add("group");
    sourceGroup.orientation = "column";
    sourceGroup.alignChildren = ["left", "top"];
    sourceGroup.margins = [0, 6, 0, 0];
    var sourceSwatchesRadio = sourceGroup.add("radiobutton", undefined, L("source.swatches"));
    sourceSwatchesRadio.value = true;
    sourceSwatchesRadio.helpTip = L("tooltip.sourceSwatches");
    var sourceObjectsRadio = sourceGroup.add("radiobutton", undefined, L("source.objects"));
    sourceObjectsRadio.helpTip = L("tooltip.sourceObjects");
    var sourceAutoRadio = sourceGroup.add("radiobutton", undefined, L("source.auto"));
    sourceAutoRadio.helpTip = L("tooltip.sourceAuto");

    /* 配色単位・配色順を2カラムで配置 / Coloring unit and order in two columns */
    var unitsRow = win.add("group");
    setupRow(unitsRow, "fill", COLUMN_SPACING);
    unitsRow.alignChildren = ["fill", "top"];

    /* 配色単位 / Coloring unit */
    var unitRadioButtons = addRadioPanel(unitsRow, L("panel.unit"), buildUnitOptions(), info.defaultUnit, onOptionChange);
    setOptionTooltips(unitRadioButtons, {
        object: L("tooltip.unitObject"),
        character: L("tooltip.unitCharacter"),
        word: L("tooltip.unitWord"),
        line: L("tooltip.unitLine"),
        paragraph: L("tooltip.unitParagraph")
    });
    updateUnitAvailability(unitRadioButtons, info);

    /* 配色順（カラム内は縦並び）/ Coloring order (vertical within its column) */
    var orderRadioButtons = addRadioPanel(unitsRow, L("panel.option"), buildOptionList("order", ["asis", "reverse", "random"]), "asis", onOptionChange);
    setOptionTooltips(orderRadioButtons, {
        asis: L("tooltip.orderAsis"),
        reverse: L("tooltip.orderReverse"),
        random: L("tooltip.orderRandom")
    });

    /* OK / キャンセル（Mac 規約：Cancel → OK）/ OK / Cancel row (Mac convention: Cancel then OK) */
    var buttonRow = win.add("group");
    setupRow(buttonRow, "right");
    var cancelButton = buttonRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
    var okButton = buttonRow.add("button", undefined, "OK", { name: "ok" });

    /* 現在の設定を取得 / Read current settings */
    function currentOptions() {
        return {
            unit: getSelectedOption(unitRadioButtons, "object"),
            order: getSelectedOption(orderRadioButtons, "asis")
        };
    }

    /* 直前のプレビューを取り消す（スナップショットがあれば元の塗り/線/不透明度へ復元）
       Revert the previous preview (restore fills/strokes/opacity from the snapshot if any) */
    function undoPreview() {
        if (!previewSnapshot) return;
        restorePreview(previewSnapshot);
        previewSnapshot = null;
        app.redraw();
    }

    /* プレビューを更新（直前を復元してから現在の設定で適用）。
       設定が前回と同じなら適用しない（ラジオの二重発火対策）。force=true で色の取り込み直し時に強制再適用
       Refresh the preview; skip when settings are unchanged (force=true to re-apply after re-capturing colors) */
    function runPreview(force) {
        var opts = currentOptions();
        var sig = opts.unit + "|" + opts.order;
        if (!force && sig === lastPreviewSig) return;
        lastPreviewSig = sig;
        undoPreview();
        previewSnapshot = applyColors(opts, loadedSwatchNames, loadedColorValues);
    }

    /* 単位・オプション変更時にプレビュー更新 / Refresh preview on any change */
    function onOptionChange() {
        runPreview();
    }

    /* チップと配色単位ディムを更新して再レイアウト（適用はしない）/ Update chips and unit dims, relayout (no apply) */
    function refreshColorUI(freshInfo) {
        clearChildren(chipHost);
        buildColorChips(chipHost, freshInfo.chips);
        updateUnitAvailability(unitRadioButtons, freshInfo);
        win.layout.layout(true);
    }

    /* 現在のカラー取り込み元 / Current color source */
    function currentSource() {
        if (sourceObjectsRadio.value) return "objects";
        if (sourceAutoRadio.value) return "auto";
        return "swatches";
    }

    /* 取り込み元を切り替えて再取り込み＋再プレビュー（同一元の再発火は無視）
       Switch source, re-capture and re-preview (ignore re-fires of the same source) */
    function onSourceChange() {
        var src = currentSource();
        if (src === lastSource) return;
        lastSource = src;
        undoPreview();
        var fresh;
        if (src === "objects") {
            fresh = readObjectColorsInfo();
            loadedColorValues = fresh.colorValues;
            loadedSwatchNames = [];
        } else if (src === "auto") {
            /* 自動生成：取り込み色は空にして resolveAppliedColors のフォールバックに任せる
               Auto: clear captured colors and let resolveAppliedColors generate them */
            fresh = readSelectionInfo();
            fresh.chips = [];
            loadedSwatchNames = [];
            loadedColorValues = [];
        } else {
            fresh = readSelectionInfo();
            loadedSwatchNames = fresh.swatchNames;
            loadedColorValues = [];
        }
        refreshColorUI(fresh);
        runPreview(true);
    }
    sourceSwatchesRadio.onClick = onSourceChange;
    sourceObjectsRadio.onClick = onSourceChange;
    sourceAutoRadio.onClick = onSourceChange;

    /* OK：現在のプレビューを確定（そのまま残す）/ OK: keep the current preview */
    okButton.onClick = function () { committed = true; win.close(); };
    /* キャンセル：プレビューを破棄 / Cancel: discard the preview */
    cancelButton.onClick = function () { committed = false; win.close(); };

    /* 未確定で閉じたら（キャンセル・Esc・閉じるボタン）プレビューを元に戻す
       Revert the preview if closed without committing (Cancel / Esc / close box) */
    win.onClose = function () {
        if (!committed) undoPreview();
        return true;
    };

    /* 表示直後に初回プレビュー / First preview right after show */
    win.onShow = function () {
        win.layout.layout(true);
        runPreview(true);
    };

    win.show();
}

/* ラジオボタンのパネルを追加して配列を返す / Add a radio-button panel and return the buttons */
function addRadioPanel(parent, title, options, defaultKey, onChange, horizontal) {
    var panel = parent.add("panel", undefined, title);
    setupPanel(panel);
    if (horizontal) {
        /* ラジオを横並びに / Lay radios out in a row */
        panel.orientation = "row";
        panel.alignChildren = ["left", "center"];
    }
    var radioButtons = [];
    for (var i = 0; i < options.length; i++) {
        var radioButton = panel.add("radiobutton", undefined, options[i].label);
        radioButton.optionKey = options[i].key;
        if (options[i].key === defaultKey) radioButton.value = true;
        radioButton.onClick = onChange;
        radioButtons.push(radioButton);
    }
    radioButtons.panel = panel; /* パネル本体も保持 / keep the panel element too */
    return radioButtons;
}

/* 選択中のラジオボタンのキーを取得 / Get the key of the checked radio button */
function getSelectedOption(radioButtons, fallbackKey) {
    for (var i = 0; i < radioButtons.length; i++) {
        if (radioButtons[i].value) return radioButtons[i].optionKey;
    }
    return fallbackKey;
}

/* キーごとの helpTip をラジオボタンに設定 / Set per-key helpTips on radio buttons */
function setOptionTooltips(radioButtons, tipsByKey) {
    for (var i = 0; i < radioButtons.length; i++) {
        var tip = tipsByKey[radioButtons[i].optionKey];
        if (tip) radioButtons[i].helpTip = tip;
    }
}

/* 選択内容に応じて配色単位ラジオの有効／無効を更新
   - オブジェクトが1個ならオブジェクトを無効
   - テキストでない場合は文字/単語/行/段落を無効
   - 段落が1つだけなら段落も無効
   - すべて無効になる場合（単一図形など）はオブジェクトを残す
   Update unit availability based on the selection */
function updateUnitAvailability(unitRadioButtons, info) {
    var enabled = {
        object: (info.itemCount !== 1),
        character: info.isText,
        word: info.isText,
        line: info.isText,
        /* 段落は、テキストが複数 or 単一テキストの段落が2つ以上のとき有効
           （＝単一テキストで段落1つのときだけディム）
           Paragraph enabled when multiple text objects, or a single text object with 2+ paragraphs */
        paragraph: (info.isText && (info.textObjectCount !== 1 || info.paragraphCount > 1))
    };

    if (!(enabled.object || enabled.character || enabled.word || enabled.line || enabled.paragraph)) {
        enabled.object = true;
    }

    for (var i = 0; i < unitRadioButtons.length; i++) {
        unitRadioButtons[i].enabled = enabled[unitRadioButtons[i].optionKey];
    }

    /* 選択中が無効になったら有効な先頭の単位へ寄せる / Move selection to the first enabled unit if the current one is disabled */
    var currentOk = false;
    for (var j = 0; j < unitRadioButtons.length; j++) {
        if (unitRadioButtons[j].value && unitRadioButtons[j].enabled) { currentOk = true; break; }
    }
    if (!currentOk) {
        for (var m = 0; m < unitRadioButtons.length; m++) {
            if (unitRadioButtons[m].enabled) {
                for (var n = 0; n < unitRadioButtons.length; n++) {
                    unitRadioButtons[n].value = (n === m);
                }
                break;
            }
        }
    }
}

/* 適用単位の選択肢 / Apply-unit options */
function buildUnitOptions() {
    return buildOptionList("unit", ["object", "character", "word", "line", "paragraph"]);
}

/* LABELS のカテゴリとキー配列から {key,label} 配列を生成 / Build a {key,label} list from a LABELS category */
function buildOptionList(category, keys) {
    var options = [];
    for (var i = 0; i < keys.length; i++) {
        options.push({ key: keys[i], label: L(category + "." + keys[i]) });
    }
    return options;
}

/* コンテナの子要素をすべて削除 / Remove all children of a container */
function clearChildren(container) {
    while (container.children.length > 0) {
        container.remove(container.children[container.children.length - 1]);
    }
}

// =========================================
// カラーチップ表示 / Color chips
// =========================================
/* 選択スウォッチを■で並べる（未選択時は注記）/ Lay out chips (note if none) */
function buildColorChips(container, chips) {
    if (!chips || chips.length === 0) {
        container.add("statictext", undefined, L("note.autoColor"));
        return;
    }
    /* 1行を1つの描画キャンバスにまとめて描く（要素ごとのレイアウトずれを避ける）
       Draw each row on a single canvas to avoid per-element vertical misalignment */
    for (var start = 0; start < chips.length; start += COLOR_CHIPS_PER_ROW) {
        addChipRowCanvas(container, chips.slice(start, start + COLOR_CHIPS_PER_ROW));
    }
}

/* 1行ぶんのチップを1つの group に手描き / Draw one row of chips onto a single group */
function addChipRowCanvas(container, rowChips) {
    var width = rowChips.length * CHIP_SIZE + (rowChips.length - 1) * CHIP_GAP;
    var canvas = container.add("group");
    canvas.preferredSize = [width, CHIP_SIZE];
    canvas.minimumSize = [width, CHIP_SIZE];
    canvas.maximumSize = [width, CHIP_SIZE];
    canvas.chipColors = rowChips;
    canvas.onDraw = function () {
        var g = this.graphics;
        for (var i = 0; i < this.chipColors.length; i++) {
            var rgb = this.chipColors[i];
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, [rgb[0] / 255, rgb[1] / 255, rgb[2] / 255, 1]);
            g.newPath();
            g.rectPath(i * (CHIP_SIZE + CHIP_GAP), 0, CHIP_SIZE, CHIP_SIZE);
            g.fillPath(brush);
        }
    };
}

// =========================================
// 選択情報の取得 / Reading the selection
// =========================================
/* 選択情報を取得（初期単位・チップRGB・スウォッチ名・選択統計）
   Read selection info (default unit, chip RGB, swatch names, selection stats) */
function readSelectionInfo() {
    var info = { ok: false, defaultUnit: "object", chips: [], swatchNames: [], itemCount: 0, isText: false, paragraphCount: 0, textObjectCount: 0 };
    if (app.documents.length === 0) return info;
    var doc = app.activeDocument;
    var items = flattenSelection(app.selection);
    var textRange = getSingleSelectedTextRange(app.selection);
    info.ok = true;
    if (textRange) { info.defaultUnit = "character"; }
    else if (items.length === 1 && items[0].typename === "TextFrame") { info.defaultUnit = "character"; }

    var swatches = doc.swatches.getSelected();
    var usable = swatches && swatches.length >= 1 && !allWhiteSwatches(swatches);
    if (usable) {
        for (var i = 0; i < swatches.length; i++) {
            info.chips.push(colorToRGB255Array(swatches[i].color));
            info.swatchNames.push(swatches[i].name);
        }
    }

    var stats = selectionTextStats(items, textRange);
    info.itemCount = items.length;
    info.isText = stats.isText;
    info.paragraphCount = stats.paragraphCount;
    info.textObjectCount = stats.textObjectCount;
    return info;
}

/* 選択オブジェクトの塗り色を取得（チップRGB＋シリアライズ済みカラー値）
   Read fill colors of selected objects (chip RGB + serialized color values) */
function readObjectColorsInfo() {
    var info = { ok: false, chips: [], colorValues: [], itemCount: 0, isText: false, paragraphCount: 0, textObjectCount: 0 };
    if (app.documents.length === 0) return info;
    var items = flattenSelection(app.selection);
    var textRange = getSingleSelectedTextRange(app.selection);
    info.ok = true;

    var fillColors = collectFillColors(items, textRange);
    for (var i = 0; i < fillColors.length; i++) {
        info.chips.push(colorToRGB255Array(fillColors[i]));
        info.colorValues.push(serializeColor(fillColors[i]));
    }

    var stats = selectionTextStats(items, textRange);
    info.itemCount = items.length;
    info.isText = stats.isText;
    info.paragraphCount = stats.paragraphCount;
    info.textObjectCount = stats.textObjectCount;
    return info;
}

/* 選択のテキスト統計（テキスト有無・段落数・テキストオブジェクト数）/ Text stats (has-text, paragraphs, text-object count) */
function selectionTextStats(items, textRange) {
    var isText = false;
    var paragraphCount = 0;
    var textObjectCount = 0;
    if (textRange) { isText = true; textObjectCount++; paragraphCount += textRange.paragraphs.length; }
    for (var i = 0; i < items.length; i++) {
        if (items[i].typename === "TextFrame") { isText = true; textObjectCount++; paragraphCount += items[i].textRange.paragraphs.length; }
    }
    return { isText: isText, paragraphCount: paragraphCount, textObjectCount: textObjectCount };
}

// =========================================
// カラー適用 / Applying colors (DOM)
// =========================================
/* カラーを適用（swatchNames=スウォッチ名 / colorValues=シリアライズ済みカラー値、両方空なら自動生成）。
   着色前に各対象の元の状態をスナップショットへ保存して返す（キャンセル・再適用時の復元に使う）。
   対象なし・エラー時は null を返す
   Apply colors and return a snapshot of the originals (used to revert); returns null when nothing was applied */
function applyColors(options, swatchNames, colorValues) {
    if (app.documents.length === 0) return null;
    var doc = app.activeDocument;
    var selection = doc.selection;
    var items = flattenSelection(selection);
    var textRange = getSingleSelectedTextRange(selection);
    /* オブジェクト選択は着色後に復元用として控える（テキスト編集中は復元しない）
       Save the object selection for restore (skip while editing text) */
    var savedSelection = textRange ? null : selection;
    var targets = (items.length > 0 || textRange) ? collectColorTargetsByUnit(items, textRange, options.unit) : [];
    if (targets.length === 0) { app.redraw(); return null; }
    var colors = orderColors(resolveAppliedColors(doc, targets.length, swatchNames, colorValues), options.order);
    /* 各対象の元の状態を保存してから着色（プレビュー取り消しで元に戻せるように）
       Snapshot each target's originals before coloring, so the preview can be reverted */
    var snapshot = [];
    try {
        for (var i = 0; i < targets.length; i++) {
            snapshotTarget(targets[i], snapshot);
            var colorIndex = (typeof targets[i].colorIndex === "number") ? targets[i].colorIndex : i;
            applyColorToTarget(targets[i], pickSwatchColor(colorIndex, colors));
        }
    } catch (e) {
        /* 途中失敗でも、ここまでのスナップショットを返せば次回 undoPreview で戻せる
           On partial failure, return the partial snapshot so it can still be reverted */
    }
    /* 適用で変わった選択（フォーカス）を元に戻す / Restore the selection changed by applying */
    if (savedSelection) { try { app.selection = savedSelection; } catch (e2) { } }
    app.redraw();
    return snapshot.length ? snapshot : null;
}

// =========================================
// プレビューのスナップショット / Preview snapshot
// =========================================
/* 着色前の対象の元の状態を snapshot 配列へ保存（テキストは文字単位で忠実に保存）
   Snapshot a target's originals into the array (text is snapshotted per character for fidelity) */
function snapshotTarget(target, out) {
    var node = target.node;
    if (target.kind === "textrange") { snapshotCharacters(node, out); }
    else if (target.kind === "textframe") { snapshotCharacters(node.textRange, out); }
    else if (target.kind === "path") { out.push({ kind: "path", node: node, fill: node.fillColor, stroked: node.stroked, opacity: node.opacity }); }
    else if (target.kind === "compound") {
        var subs = node.pathItems;
        var arr = [];
        for (var i = 0; i < subs.length; i++) { arr.push({ sub: subs[i], fill: subs[i].fillColor, stroked: subs[i].stroked, opacity: subs[i].opacity }); }
        out.push({ kind: "compound", node: node, subs: arr });
    }
}

/* テキスト範囲を文字単位でスナップショット（元が混色でも忠実に戻せる）
   Snapshot a text range per character (so mixed original colors restore faithfully) */
function snapshotCharacters(range, out) {
    var chars = range.characters;
    for (var i = 0; i < chars.length; i++) {
        var ch = chars[i];
        out.push({ kind: "char", node: ch, fill: ch.fillColor, stroke: ch.strokeColor, opacity: ch.opacity });
    }
}

/* スナップショットから元の塗り/線/不透明度を復元（対象削除時などは個別に無視）
   Restore fills/strokes/opacity from the snapshot (skip individually if an object was deleted) */
function restorePreview(snapshot) {
    for (var i = 0; i < snapshot.length; i++) {
        try { restoreOneSnapshot(snapshot[i]); } catch (e) { }
    }
}

/* スナップショット1件を復元 / Restore one snapshot entry */
function restoreOneSnapshot(entry) {
    var node = entry.node;
    if (entry.kind === "char") { node.fillColor = entry.fill; node.strokeColor = entry.stroke; node.opacity = entry.opacity; }
    else if (entry.kind === "path") { node.fillColor = entry.fill; node.stroked = entry.stroked; node.opacity = entry.opacity; }
    else if (entry.kind === "compound") {
        for (var i = 0; i < entry.subs.length; i++) {
            var s = entry.subs[i];
            try { s.sub.fillColor = s.fill; s.sub.stroked = s.stroked; s.sub.opacity = s.opacity; } catch (e) { }
        }
    }
}

/* 選択をフラット化して色付け対象のみ収集 / Flatten selection to colorable items */
function flattenSelection(selection) {
    var collected = [];
    if (!selection || selection.length === 0) return collected;
    for (var i = 0; i < selection.length; i++) { collectColorableItems(selection[i], collected); }
    return collected;
}

/* 色付け可能な要素を再帰収集（グループ内も）/ Recursively collect colorable items */
function collectColorableItems(item, out) {
    if (!item) return;
    if (item.typename === "PathItem" || item.typename === "CompoundPathItem" || item.typename === "TextFrame") { out.push(item); return; }
    if (item.typename === "GroupItem") {
        var children = item.pageItems;
        for (var i = 0; i < children.length; i++) { collectColorableItems(children[i], out); }
    }
}

/* テキスト編集中の選択範囲を取得 / Get the active text range while editing */
function getSingleSelectedTextRange(selection) {
    if (!selection || selection.length !== 1) return null;
    if (selection[0] && selection[0].typename === "TextRange") return selection[0];
    return null;
}

/* 適用単位に応じた対象を収集 / Collect targets by apply unit */
function collectColorTargetsByUnit(items, textRange, unitMode) {
    var targets = [];
    if (textRange) { pushTextRangeTargets(textRange, unitMode, targets); return targets; }
    if (items.length === 1 && items[0].typename === "TextFrame") {
        if (unitMode === "object") { targets.push({ kind: "textframe", node: items[0] }); }
        else { pushTextRangeTargets(items[0].textRange, unitMode, targets); }
        return targets;
    }
    var sorted = items.slice();
    sortByPosition(sorted);
    for (var i = 0; i < sorted.length; i++) {
        var it = sorted[i];
        if (it.typename === "TextFrame" && unitMode !== "object") { pushTextRangeTargets(it.textRange, unitMode, targets); }
        else if (it.typename === "PathItem") { targets.push({ kind: "path", node: it }); }
        else if (it.typename === "CompoundPathItem") { targets.push({ kind: "compound", node: it }); }
        else if (it.typename === "TextFrame") { targets.push({ kind: "textframe", node: it }); }
    }
    return targets;
}

/* テキスト範囲を単位ごとに分割してターゲット化 / Split a text range into per-unit targets */
function pushTextRangeTargets(textRange, unitMode, out) {
    if (unitMode === "word") { pushStaggeredWordTargets(textRange, out); return; }
    var ranges = getTextUnitRanges(textRange, unitMode);
    if (!ranges) { out.push({ kind: "textrange", node: textRange }); return; }
    for (var i = 0; i < ranges.length; i++) { out.push({ kind: "textrange", node: ranges[i] }); }
}

/* 単語ごと：Illustrator の単語境界で文字を単語にマッピングし、句読点・スペースは直前の単語色にする
   （略語・小数は1単語のまま。各行先頭は互い違い）
   Per word: map characters to Illustrator's word boundaries; punctuation/spaces share the preceding word's color */
function pushStaggeredWordTargets(textRange, out) {
    var lines = textRange.lines;
    for (var li = 0; li < lines.length; li++) {
        var line = lines[li];
        var words = line.words;
        var wordStarts = {};
        for (var wi = 0; wi < words.length; wi++) { wordStarts[words[wi].start] = true; }
        var chars = line.characters;
        var wordIndex = -1;
        for (var ci = 0; ci < chars.length; ci++) {
            if (wordStarts[chars[ci].start]) { wordIndex++; }
            out.push({ kind: "textrange", node: chars[ci], colorIndex: (wordIndex < 0 ? 0 : wordIndex) + li });
        }
    }
}

/* 単位に対応するテキスト範囲コレクション（word は pushTextRangeTargets で先に処理）
   Text range collection for a unit (word is handled earlier in pushTextRangeTargets) */
function getTextUnitRanges(textRange, unitMode) {
    if (unitMode === "character") return textRange.characters;
    if (unitMode === "line") return textRange.lines;
    if (unitMode === "paragraph") return textRange.paragraphs;
    return null;
}

/* ターゲットにカラーを設定 / Set the color on a target */
function applyColorToTarget(target, color) {
    var node = target.node;
    if (target.kind === "textrange") { node.fillColor = color; node.strokeColor = new NoColor(); node.opacity = 100; }
    else if (target.kind === "textframe") { node.textRange.fillColor = color; node.textRange.strokeColor = new NoColor(); node.textRange.opacity = 100; }
    else if (target.kind === "path") { node.fillColor = color; node.stroked = false; node.opacity = 100; }
    else if (target.kind === "compound") {
        var subs = node.pathItems;
        for (var i = 0; i < subs.length; i++) { subs[i].fillColor = color; subs[i].stroked = false; subs[i].opacity = 100; }
    }
}

/* 位置順にソート / Sort by position */
function sortByPosition(items) {
    var minLeft = Infinity, maxLeft = -Infinity, minTop = Infinity, maxTop = -Infinity;
    for (var i = 0; i < items.length; i++) {
        var left = items[i].left;
        var top = items[i].top;
        if (left < minLeft) minLeft = left;
        if (left > maxLeft) maxLeft = left;
        if (top < minTop) minTop = top;
        if (top > maxTop) maxTop = top;
    }
    if (maxLeft - minLeft > maxTop - minTop) { items.sort(function (a, b) { return comparePositionKeys(a.left, b.left, b.top, a.top); }); }
    else { items.sort(function (a, b) { return comparePositionKeys(b.top, a.top, a.left, b.left); }); }
}

/* ソート用比較（主キー→副キー）/ Compare by primary then secondary key */
function comparePositionKeys(primaryA, primaryB, secondaryA, secondaryB) {
    return primaryA == primaryB ? secondaryA - secondaryB : primaryA - primaryB;
}

/* ランダムシャッフルして返す / Return a shuffled copy */
function shuffleArray(source) {
    var shuffled = source.slice();
    for (var i = shuffled.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = shuffled[i];
        shuffled[i] = shuffled[j];
        shuffled[j] = temp;
    }
    return shuffled;
}

/* 整数乱数 / Random integer in [min, max] */
function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

/* 使用するカラーを決定（取り込んだスウォッチ名 or カラー値を優先、無ければ自動生成）
   Resolve colors (captured swatch names or serialized values take priority; otherwise auto-generate) */
function resolveAppliedColors(doc, targetCount, swatchNames, colorValues) {
    var colors = [];
    if (swatchNames && swatchNames.length > 0) {
        for (var i = 0; i < swatchNames.length; i++) {
            var swatchColor = findSwatchColorByName(doc, swatchNames[i]);
            if (swatchColor) { colors.push(swatchColor); }
        }
    } else if (colorValues && colorValues.length > 0) {
        for (var j = 0; j < colorValues.length; j++) {
            colors.push(deserializeColor(colorValues[j]));
        }
    }
    if (colors.length >= 1) { return colors; }
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) { return generateRandomCMYPaletteUnique(targetCount, TMK_CMYK_FALLBACK_MAX_TOTAL); }
    return getDefaultRGBColors();
}

/* スウォッチ名から色を取得（選択状態に依存しない）/ Find a swatch color by name (independent of selection state) */
function findSwatchColorByName(doc, name) {
    var swatches = doc.swatches;
    for (var i = 0; i < swatches.length; i++) {
        if (swatches[i].name === name) { return swatches[i].color; }
    }
    return null;
}

/* 選択オブジェクトの塗り色を重複なく収集（黒・白は除外）/ Collect distinct fill colors (excluding black/white) */
function collectFillColors(items, textRange) {
    var candidates = [];
    if (textRange) { candidates.push(textRange.characterAttributes.fillColor); }
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it.typename === "PathItem") { candidates.push(it.fillColor); }
        else if (it.typename === "CompoundPathItem" && it.pathItems.length > 0) { candidates.push(it.pathItems[0].fillColor); }
        else if (it.typename === "TextFrame") { candidates.push(it.textRange.characterAttributes.fillColor); }
    }
    var colors = [];
    var seen = {};
    for (var j = 0; j < candidates.length; j++) {
        var c = candidates[j];
        if (!isApplicableFillColor(c)) { continue; }
        /* 黒・白は取り込み対象から除外 / Exclude pure black and white */
        if (isBlackColor(c) || isWhiteColor(c)) { continue; }
        var key = serializeColor(c);
        if (seen[key]) { continue; }
        seen[key] = true;
        colors.push(c);
    }
    return colors;
}

/* 単色として取り込める塗り色か（NoColor・グラデ・パターンは除外）
   Whether a fill is a solid color usable as a swatch (skip NoColor / gradient / pattern) */
function isApplicableFillColor(color) {
    if (!color) { return false; }
    var t = color.typename;
    return (t === "CMYKColor" || t === "RGBColor" || t === "GrayColor" || t === "SpotColor");
}

/* カラーを文字列にシリアライズ / Serialize a color to a string */
function serializeColor(color) {
    var t = color.typename;
    if (t === "CMYKColor") { return "C," + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black; }
    if (t === "RGBColor") { return "R," + color.red + "," + color.green + "," + color.blue; }
    if (t === "GrayColor") { return "G," + color.gray; }
    if (t === "SpotColor") { return serializeColor(color.spot.color); }
    return "R,128,128,128";
}

/* 文字列からカラーを復元 / Reconstruct a color from a string */
function deserializeColor(str) {
    var p = str.split(",");
    if (p[0] === "C") { return buildCMYKColor(Number(p[1]), Number(p[2]), Number(p[3]), Number(p[4])); }
    if (p[0] === "R") { return buildRGBColor(Number(p[1]), Number(p[2]), Number(p[3])); }
    if (p[0] === "G") { return buildGrayColor(Number(p[1])); }
    return buildRGBColor(128, 128, 128);
}

/* Gray カラーを生成 / Build a GrayColor */
function buildGrayColor(gray) {
    var color = new GrayColor();
    color.gray = gray;
    return color;
}

/* 適用順に並べ替え / Reorder colors */
function orderColors(colors, orderMode) {
    if (orderMode === "reverse") return colors.slice().reverse();
    if (orderMode === "random") return shuffleArray(colors);
    return colors;
}

/* インデックスに応じて色を取得（ループ・不透明度100%）/ Pick a color by index (wraps, opacity 100%) */
function pickSwatchColor(index, colors) {
    var color = colors[index % colors.length];
    color.opacity = 100;
    return color;
}

/* CMYK カラーを生成 / Build a CMYKColor */
function buildCMYKColor(cyan, magenta, yellow, black) {
    var color = new CMYKColor();
    color.cyan = cyan; color.magenta = magenta; color.yellow = yellow; color.black = black;
    return color;
}

/* RGB カラーを生成 / Build an RGBColor */
function buildRGBColor(red, green, blue) {
    var color = new RGBColor();
    color.red = red; color.green = green; color.blue = blue;
    return color;
}

/* RGB ドキュメント用の定義済みカラー / Predefined colors for RGB documents */
function getDefaultRGBColors() {
    var defs = [[222, 84, 25], [245, 233, 40], [41, 163, 57], [53, 157, 209], [173, 127, 71], [238, 176, 51]];
    var colors = [];
    for (var i = 0; i < defs.length; i++) { colors.push(buildRGBColor(defs[i][0], defs[i][1], defs[i][2])); }
    return colors;
}

/* CM/CY/MY の2チャンネル（K=0）をランダムに1色ぶん生成 / Pick a random 2-channel CMY color */
function pickTwoChannelCMY(channelPair, maxTotal) {
    var amountA = randInt(1, Math.min(100, maxTotal - 1));
    var amountBMax = Math.min(100, maxTotal - amountA);
    if (amountBMax < 1) return null;
    var amountB = randInt(1, amountBMax);
    if (channelPair === "CM") return { cyan: amountA, magenta: amountB, yellow: 0 };
    if (channelPair === "CY") return { cyan: amountA, magenta: 0, yellow: amountB };
    return { cyan: 0, magenta: amountA, yellow: amountB };
}

/* CMYK 用：CM/CY/MY のみで可能な限り重複しない色を生成 / Generate CM/CY/MY colors, avoid duplicates */
function generateRandomCMYPaletteUnique(count, maxTotal) {
    var palette = [];
    var accepted = [];
    var seen = {};
    var channelPairs = ["CM", "CY", "MY"];
    var minDistance = TMK_CMYK_FALLBACK_MIN_DISTANCE;
    var maxAttempts = Math.max(3000, count * 120);
    var attempts = 0;
    while (palette.length < count) {
        attempts++;
        if (minDistance > 0 && (attempts % 500) === 0) { minDistance = Math.max(0, minDistance - 5); }
        var allowDuplicate = attempts > maxAttempts;
        if (allowDuplicate) { minDistance = 0; }
        if ((attempts % 37) === 0) { channelPairs = shuffleArray(channelPairs); }
        var channelPair = channelPairs[palette.length % channelPairs.length];
        var channels = pickTwoChannelCMY(channelPair, maxTotal);
        if (!channels) continue;
        var key = channels.cyan + "," + channels.magenta + "," + channels.yellow;
        if (!allowDuplicate && seen[key]) continue;
        if (minDistance > 0 && !isFarEnoughCMY(channels.cyan, channels.magenta, channels.yellow, accepted, minDistance)) continue;
        seen[key] = true;
        var color = buildCMYKColor(channels.cyan, channels.magenta, channels.yellow, 0);
        palette.push(color);
        accepted.push(color);
    }
    return palette;
}

/* 既存の採用色から十分離れているか / Whether far enough from accepted colors */
function isFarEnoughCMY(cyan, magenta, yellow, existing, minDistance) {
    for (var i = 0; i < existing.length; i++) {
        if (cmyDistance(cyan, magenta, yellow, existing[i].cyan, existing[i].magenta, existing[i].yellow) < minDistance) return false;
    }
    return true;
}

/* CMY のマンハッタン距離 / Manhattan distance in CMY */
function cmyDistance(c1, m1, y1, c2, m2, y2) {
    return Math.abs(c1 - c2) + Math.abs(m1 - m2) + Math.abs(y1 - y2);
}

/* 色が白か判定（RGB換算で判定するので Gray/Spot も対応）/ Whether a color is white (via RGB, covers Gray/Spot) */
function isWhiteColor(color) {
    var rgb = colorToRGB255Array(color);
    return rgb[0] === 255 && rgb[1] === 255 && rgb[2] === 255;
}

/* 色が黒か判定（RGB換算で判定するので Gray/Spot も対応）/ Whether a color is black (via RGB, covers Gray/Spot) */
function isBlackColor(color) {
    var rgb = colorToRGB255Array(color);
    return rgb[0] === 0 && rgb[1] === 0 && rgb[2] === 0;
}

/* 全スウォッチが白のみか判定 / Whether all swatches are white */
function allWhiteSwatches(swatches) {
    for (var i = 0; i < swatches.length; i++) {
        if (!isWhiteColor(swatches[i].color)) return false;
    }
    return true;
}

/* Illustrator カラーを [r,g,b](0..255) 配列へ / Convert an Illustrator color to an [r,g,b] array */
function colorToRGB255Array(color) {
    var t = color.typename;
    if (t === "RGBColor") return [Math.round(color.red), Math.round(color.green), Math.round(color.blue)];
    if (t === "CMYKColor") {
        var k = 1 - color.black / 100;
        return [Math.round(255 * (1 - color.cyan / 100) * k), Math.round(255 * (1 - color.magenta / 100) * k), Math.round(255 * (1 - color.yellow / 100) * k)];
    }
    if (t === "GrayColor") {
        var v = Math.round(255 * (1 - color.gray / 100));
        return [v, v, v];
    }
    if (t === "SpotColor") {
        var base = colorToRGB255Array(color.spot.color);
        var tint = (typeof color.tint === "number") ? color.tint / 100 : 1;
        return [Math.round(255 - (255 - base[0]) * tint), Math.round(255 - (255 - base[1]) * tint), Math.round(255 - (255 - base[2]) * tint)];
    }
    return [128, 128, 128];
}

showDialog();

})();
