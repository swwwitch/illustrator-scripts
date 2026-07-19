#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択したオブジェクトやテキストに、スウォッチや定義済みカラーを適用するモーダルダイアログです。
- 適用単位（オブジェクト／1文字／単語／行／段落）と適用順（そのまま／逆順／ランダム／完全ランダム）をラジオボタンで選択します。
- 「ランダム」はカラーの並びをシャッフルして繰り返し適用、「完全ランダム」は適用先ごとに毎回抽選するため繰り返しがありません。
- ラジオを変えるたびにライブプレビュー。［OK］で確定、［キャンセル］（Esc）で元に戻します。
- 開いた時点の選択スウォッチを■で取り込み、それ（スウォッチ名で参照）を適用に使用。1色でも選択があれば優先します。
- スウォッチ未選択時は自動カラー（CMYK は CM/CY/MY 生成、RGB は既定色）を使用。
- 「単語」は各行の先頭色が互い違いになるよう配色。
- モーダルダイアログはメインエンジンで動作するため、DOM 操作を直接実行します（BridgeTalk 委譲なし）。

### 紹介記事（note）

https://note.com/dtp_tranist/n/n5602f3084d2b

### Overview

- A modal dialog that applies swatches or predefined colors to selected objects or text.
- Choose apply unit (object / character / word / line / paragraph) and order (as-is / reverse / random / fully random) with radio buttons.
- "Random" shuffles the color list once and cycles through it; "Fully random" draws a color per target, so no sequence repeats.
- Live preview on every change. [OK] commits the result; [Cancel] (Esc) reverts it.
- The swatches selected when the dialog opens are captured as chips and used (by swatch name) for applying. Even a single selected swatch is used.
- When no swatches are selected, auto colors are used (CMYK CM/CY/MY generation, RGB defaults).
- "Per word" staggers colors so each line starts on a different color.
- The dialog runs in the main engine, so DOM work is executed directly (no BridgeTalk delegation).

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiApplySwatchesToSelection-dialog";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.0";                             /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";        /* 作者 / author */
var SCRIPT_RELEASED = "";                                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================
/* CMYK フォールバック生成の設定 / CMYK fallback generation settings */
var TMK_CMYK_FALLBACK_MAX_TOTAL = 200;    /* 合計上限 C+M / C+Y / M+Y / total channel limit */
var TMK_CMYK_FALLBACK_MIN_DISTANCE = 35;  /* 近似色回避の最小距離 / min distance to avoid similar colors */
var COLOR_CHIPS_PER_ROW = 12;             /* 1行あたりのチップ数 / chips per row */
var CHIP_SIZE = 18;                       /* チップの一辺(px) / chip size */
var CHIP_GAP = 4;                         /* チップ間の間隔(px) / gap between chips */
var PREVIEW_CHAR_CAP = 500;               /* 1文字単位プレビューで着色する最大文字数（超過分は確定時に着色）/ max chars colored in per-character preview */

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
        object:    { ja: "オブジェクト", en: "Per object" },
        character: { ja: "1文字", en: "Per character" },
        word:      { ja: "単語", en: "Per word" },
        line:      { ja: "行", en: "Per line" },
        paragraph: { ja: "段落", en: "Per paragraph" }
    },
    order: {
        asis:       { ja: "取り込み順", en: "As captured" },
        reverse:    { ja: "逆順", en: "Reverse" },
        random:     { ja: "ランダム", en: "Random" },
        fullrandom: { ja: "完全ランダム", en: "Fully random" }
    },
    button: {
        ok:     { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    source: {
        selected: { ja: "選択しているスウォッチ", en: "Selected swatches" },
        group:    { ja: "スウォッチグループ：", en: "Swatch group:" },
        noGroup:  { ja: "（グループなし）", en: "(no groups)" }
    },
    tooltip: {
        colors:        { ja: "開いた時点で選択していたスウォッチを適用に使います。", en: "The swatches selected when the dialog opened are used for applying." },
        unitObject:    { ja: "選択オブジェクト単位でカラーを順番に適用します。", en: "Applies colors in order, one per selected object." },
        unitCharacter: { ja: "テキストを1文字ずつ分けてカラーを適用します。", en: "Applies a color to each character of the text." },
        unitWord:      { ja: "英文テキストを単語ごとに分けてカラーを適用します。英文以外では単語が正しく分割されないことがあります。", en: "Applies a color to each word of English text. Words may not split correctly for non-English text." },
        unitLine:      { ja: "テキストの行ごとにカラーを適用します。", en: "Applies a color to each line of the text." },
        unitParagraph: { ja: "テキストの段落ごとにカラーを適用します。", en: "Applies a color to each paragraph of the text." },
        orderAsis:     { ja: "取り込んだカラーの並び順で適用します（適用先は位置順・文字順）。", en: "Applies colors in the captured order (targets follow position / reading order)." },
        orderReverse:  { ja: "取り込んだカラーの並びを逆にして適用します。", en: "Applies the captured colors in reverse order." },
        orderRandom:     { ja: "取り込んだカラーの並びをランダムにして適用します（並びは繰り返します）。", en: "Applies the captured colors in a random order (the sequence repeats)." },
        orderFullRandom: { ja: "適用先ごとにカラーを毎回ランダムに選びます（並びは繰り返しません）。", en: "Picks a color at random for each target (no repeating sequence)." }
    },
    alert: {
        noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSel: { ja: "オブジェクトを選択してください。", en: "Please select objects." }
    },
    note: {
        autoColor: {
            ja: "カラー未選択：自動カラーを使用します",
            en: "No colors selected: using auto colors"
        }
    },
    progress: {
        title:    { ja: "準備しています…", en: "Preparing…" },
        analyze:  { ja: "対象を解析しています…", en: "Analyzing selection…" },
        snapshot: { ja: "元の状態を保存しています…", en: "Saving original state…" },
        apply:    { ja: "カラーを適用しています…", en: "Applying colors…" }
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
// プログレスバー / Progress bar
// =========================================
/* 進捗表示用のパレットウィンドウを生成して表示（非モーダルなので処理中に更新できる）。
   戻り値の set(fraction 0..1, message) で進捗を更新、close() で閉じる
   Create and show a modeless palette progress window (updatable while work runs).
   Returned set(fraction 0..1, message) advances it; close() closes it */
function createProgressW(title) {
    var win = new Window("palette", title);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.margins = 16;
    win.spacing = 8;
    var bar = win.add("progressbar", undefined, 0, 100);
    bar.preferredSize = [320, 8];
    var label = win.add("statictext", undefined, "");
    label.preferredSize.width = 320;
    win.show();
    win.update();
    return {
        set: function (fraction, message) {
            var value = Math.round(fraction * 100);
            bar.value = (value < 0) ? 0 : (value > 100 ? 100 : value);
            if (message != null) { label.text = message; }
            win.update(); /* 同期ループ中でも再描画させる / force a repaint even inside a synchronous loop */
        },
        close: function () { win.close(); }
    };
}

// =========================================
// ダイアログ / Dialog
// =========================================
/* モーダルダイアログを表示。メインエンジンで動くので DOM 操作は直接呼ぶ（BridgeTalk 委譲なし）。
   ライブプレビューは workerApply がスナップショットで前回分を戻しつつ再適用する。
   OK でプレビューを確定、キャンセル／Esc／閉じるで元へ戻す
   Show the modal dialog. Running in the main engine, DOM work is called directly (no BridgeTalk).
   Live preview: workerApply reverts the previous preview via snapshot and reapplies.
   OK commits; Cancel / Esc / close reverts */
function showDialog() {
    if (app.documents.length === 0) { alert(L("alert.noDoc")); return; }

    /* 選択情報を取得（初期単位・チップ・スウォッチ名）/ Read selection info (default unit, chips, swatch names) */
    var info = readSelectionInfo();
    if (!info.ok || (!info.isText && info.itemCount === 0)) { alert(L("alert.noSel")); return; }

    var win = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    setupWindow(win);

    /* ドキュメントのスウォッチグループ（未分類は除外）/ Document swatch groups (uncategorized excluded) */
    var swatchGroups = readSwatchGroupsInfo();

    /* 適用に使うスウォッチ名（スウォッチ名で参照）。カラーソースに応じて差し替える
       Swatch names used for applying (referenced by name); swapped by the chosen color source */
    var loadedSwatchNames = info.swatchNames;

    /* 適用するカラー（配色チップ＋カラーソース選択）/ Color chips + color source */
    var colorPanel = win.add("panel", undefined, L("panel.colors"));
    setupPanel(colorPanel);
    colorPanel.helpTip = L("tooltip.colors");
    var chipHost = colorPanel.add("group");
    chipHost.orientation = "column";
    chipHost.alignChildren = ["left", "top"];
    buildColorChips(chipHost, info.chips);

    /* カラーソース：選択スウォッチ or スウォッチグループ（ラジオは手動で排他制御）。
       ラジオを group にまとめ、上マージンでチップとの間隔をとる
       Color source radios wrapped in a group; top margin adds space above them */
    var sourceGroup = colorPanel.add("group");
    sourceGroup.orientation = "column";
    sourceGroup.alignChildren = ["left", "top"];
    sourceGroup.margins = [0, 10, 0, 0];
    var selectedRadio = sourceGroup.add("radiobutton", undefined, L("source.selected"));
    selectedRadio.value = true;
    var groupRadio = sourceGroup.add("radiobutton", undefined, L("source.group"));
    var groupNames = [];
    for (var gi = 0; gi < swatchGroups.length; gi++) { groupNames.push(swatchGroups[gi].name); }
    /* ポップアップは次の行に置き、グループラジオの下へ少しインデントする
       Put the dropdown on the next line, slightly indented under the group radio */
    var groupDropdownRow = sourceGroup.add("group");
    setupRow(groupDropdownRow, "left");
    groupDropdownRow.margins = [16, 0, 0, 0];
    var groupDropdown = groupDropdownRow.add("dropdownlist", undefined, groupNames.length > 0 ? groupNames : [L("source.noGroup")]);
    groupDropdown.selection = 0;
    if (swatchGroups.length === 0) {
        /* グループが無ければグループ選択は無効 / disable the group option when there are none */
        groupRadio.enabled = false;
        groupDropdown.enabled = false;
    }

    /* カラーソースを適用してチップとプレビューを更新 / Apply the color source, refresh chips and preview */
    function applyColorSource() {
        if (groupRadio.value && swatchGroups.length > 0) {
            var group = swatchGroups[groupDropdown.selection.index];
            loadedSwatchNames = group.swatchNames;
            refreshChips(group.chips);
        } else {
            loadedSwatchNames = info.swatchNames;
            refreshChips(info.chips);
        }
        runPreview();
    }

    /* チップを描き直してレイアウトを取り直す / Rebuild chips and relayout */
    function refreshChips(chips) {
        clearChildren(chipHost);
        buildColorChips(chipHost, chips);
        win.layout.layout(true);
    }

    selectedRadio.onClick = function () {
        selectedRadio.value = true; groupRadio.value = false;
        applyColorSource();
    };
    groupRadio.onClick = function () {
        groupRadio.value = true; selectedRadio.value = false;
        applyColorSource();
    };
    groupDropdown.onChange = function () {
        if (!groupRadio.value) { groupRadio.value = true; selectedRadio.value = false; }
        applyColorSource();
    };

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
    var orderRadioButtons = addRadioPanel(unitsRow, L("panel.option"), buildOptionList("order", ["asis", "reverse", "random", "fullrandom"]), "asis", onOptionChange);
    setOptionTooltips(orderRadioButtons, {
        asis: L("tooltip.orderAsis"),
        reverse: L("tooltip.orderReverse"),
        random: L("tooltip.orderRandom"),
        fullrandom: L("tooltip.orderFullRandom")
    });

    /* OK / キャンセル / OK and Cancel */
    var buttonRow = win.add("group");
    buttonRow.orientation = "row";
    buttonRow.alignment = "right";
    buttonRow.spacing = PANEL_SPACING;
    var cancelButton = buttonRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
    var okButton = buttonRow.add("button", undefined, L("button.ok"), { name: "ok" });

    /* 現在の設定を取得 / Read current settings */
    function currentOptions() {
        return {
            unit: getSelectedOption(unitRadioButtons, "object"),
            order: getSelectedOption(orderRadioButtons, "asis")
        };
    }

    /* プレビュー適用（メインエンジンで直接実行）/ Apply the preview directly in the main engine */
    function runPreview() {
        workerApply(currentOptions(), loadedSwatchNames);
    }

    /* 単位・配色順の変更でプレビュー更新 / Refresh the preview on any change */
    function onOptionChange() {
        runPreview();
    }

    okButton.onClick = function () { win.close(1); };
    cancelButton.onClick = function () { win.close(2); };

    /* 初回プレビューはダイアログ表示前に済ませる。元状態の保存（全文字読み取り）が重いので、
       重い選択のときはプログレスバーを出して進捗を見せ、終わってから塗り済みのダイアログを開く
       Do the first preview before showing the dialog. Saving originals (reading every character) is
       heavy, so show a progress bar for heavy selections, then open the already-painted dialog */
    workerCommitPreview(); /* 旧ベースラインを破棄（現在の状態を基準にする）/ discard any stale baseline */
    runFirstPreviewW(currentOptions(), loadedSwatchNames);

    /* モーダル実行。OK 以外（キャンセル・Esc・クローズボックス）は元へ戻す
       Run modally; anything other than OK (Cancel, Esc, close box) reverts the preview */
    var result = win.show();
    if (result === 1) {
        /* 間引きプレビューだったときだけ、保存した配色で全文字をフル着色（プレビューと同じ結果になる）。
           間引きでなければプレビューが最終結果なので再適用しない（再適用するとランダムが引き直されて変わる）
           Only when the preview was decimated, apply fully using the saved colors (matches the preview);
           otherwise the preview is already final, so skip re-apply (it would re-randomize) */
        if ($.global.__aiApplyWasDecimated) {
            workerApply(currentOptions(), loadedSwatchNames, true);
        }
        workerCommitPreview();
    } else {
        restoreBaselineW();
        app.redraw();
    }
}

/* ラジオボタンのパネルを追加して配列を返す / Add a radio-button panel and return the buttons */
function addRadioPanel(parent, title, options, defaultKey, onChange) {
    var panel = parent.add("panel", undefined, title);
    setupPanel(panel);
    var radioButtons = [];
    for (var i = 0; i < options.length; i++) {
        var radioButton = panel.add("radiobutton", undefined, options[i].label);
        radioButton.optionKey = options[i].key;
        if (options[i].key === defaultKey) radioButton.value = true;
        radioButton.onClick = onChange;
        radioButtons.push(radioButton);
    }
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
/* 選択情報を取得してパースする（メインエンジンで直接実行）/ Read and parse selection info (called directly) */
function readSelectionInfo() {
    var info = { ok: false, defaultUnit: "object", chips: [], swatchNames: [], itemCount: 0, isText: false, paragraphCount: 0, textObjectCount: 0 };
    var raw = workerReadSelection();
    if (!raw || raw.indexOf("OK|") !== 0) return info;
    info.ok = true;

    var parts = raw.split("|");
    info.defaultUnit = parts[1] || "object";
    info.chips = parseChips(parts[2]);
    /* スウォッチ名は encodeURIComponent 済みのまま保持（適用時にそのまま worker へ渡す）
       Keep swatch names URI-encoded (passed to the worker as-is at apply time) */
    if (parts[3]) {
        info.swatchNames = parts[3].split(",");
    }
    parseSelectionStats(info, parts);
    return info;
}

/* ドキュメントのスウォッチグループを取得（未分類＝名前なしは除外、NoColor スウォッチも除外）。
   各グループは { name, swatchNames(encode済み), chips([[r,g,b]]) }。メインエンジンで直接実行
   Read the document's swatch groups (excluding the unnamed uncategorized group and NoColor swatches).
   Each group is { name, swatchNames (encoded), chips ([[r,g,b]]) }; called directly in the main engine */
function readSwatchGroupsInfo() {
    var groups = [];
    if (app.documents.length === 0) return groups;
    var doc = app.activeDocument;
    for (var i = 0; i < doc.swatchGroups.length; i++) {
        var sg = doc.swatchGroups[i];
        var name, swatches;
        /* 未分類グループの name は ""、getAllSwatches はまれに投げる。1つの try でまとめてスキップ判定
           The uncategorized group's name is ""; getAllSwatches occasionally throws — one try covers both */
        try {
            name = sg.name;
            if (name === "") { continue; } /* 未分類グループは除外 / skip the uncategorized group */
            swatches = sg.getAllSwatches();
        } catch (e) { continue; }
        var names = [];
        var chips = [];
        for (var j = 0; j < swatches.length; j++) {
            var color = swatches[j].color;
            if (color.typename === "NoColor") { continue; } /* [なし] は除外 / skip the None swatch */
            names.push(encodeURIComponent(swatches[j].name));
            chips.push(parseRGBTripletW(colorToRGB255W(color)));
        }
        if (names.length === 0) { continue; }
        groups.push({ name: name, swatchNames: names, chips: chips });
    }
    return groups;
}

/* 返却の選択統計（item数・テキスト有無・段落数）を info に取り込む / Parse selection stats into info */
function parseSelectionStats(info, parts) {
    info.itemCount = parts[4] ? parseInt(parts[4], 10) : 0;
    info.isText = (parts[5] === "1");
    info.paragraphCount = parts[6] ? parseInt(parts[6], 10) : 0;
    info.textObjectCount = parts[7] ? parseInt(parts[7], 10) : 0;
}

/* "r,g,b;r,g,b;..." を [[r,g,b],...] にパース / Parse a chip RGB string */
function parseChips(chipsField) {
    var chips = [];
    if (!chipsField) return chips;
    var chipStrings = chipsField.split(";");
    for (var i = 0; i < chipStrings.length; i++) {
        chips.push(parseRGBTripletW(chipStrings[i]));
    }
    return chips;
}

/* "r,g,b" を [r,g,b] にパース / Parse an "r,g,b" string into [r,g,b] */
function parseRGBTripletW(csv) {
    var p = csv.split(",");
    return [parseInt(p[0], 10), parseInt(p[1], 10), parseInt(p[2], 10)];
}

// =========================================
// DOM 処理 / DOM operations
// =========================================

/* 初回プレビューを実行。重い選択（長文・多数オブジェクト）のときだけプログレスバーを出す。
   軽い選択では一瞬で終わるためバーは出さない（点滅を避ける）
   Run the first preview; show a progress bar only for a heavy selection (long text / many objects).
   Light selections finish instantly, so no bar is shown (avoids a flash) */
function runFirstPreviewW(options, swatchNames) {
    var selection = (app.documents.length > 0) ? app.activeDocument.selection : null;
    var items = selection ? flattenSelectionW(selection) : [];
    var textRange = selection ? getSingleSelectedTextRangeW(selection) : null;
    var heavy = (baselineCharTotalW(items, textRange) > 1000) || (items.length > 300);
    if (!heavy) { workerApply(options, swatchNames); return; }
    var progress = createProgressW(L("progress.title"));
    try {
        workerApply(options, swatchNames, false, progress);
    } finally {
        progress.close();
    }
}

/* カラーを適用（swatchNames=適用するスウォッチ名の配列）。
   元の塗り/線/不透明度は開いた最初の1回だけ全選択ぶんスナップショット（ベースライン）し、
   以降のプレビューは上書きのみ＝復元も再スナップショットもしない。
   どの配色単位でも可視グリフ・全オブジェクトを必ず塗り直すので見た目は常に正しく、
   キャンセル時はベースラインから完全復元する（app.undo はグローバル履歴を巻き戻すため使わない）。
   これで単位切替のたびに文字を読み直していた重さを解消する
   Apply colors. Snapshot the whole selection's originals once (baseline); later previews only
   overwrite — no restore, no re-snapshot. Every unit repaints all visible glyphs / all objects,
   so the result is always correct; Cancel restores from the baseline (never app.undo) */
function workerApply(options, swatchNames, finalApply, progress) {
    if (app.documents.length === 0) return "NODOC";
    var doc = app.activeDocument;
    var selection = doc.selection;
    var items = flattenSelectionW(selection);
    var textRange = getSingleSelectedTextRangeW(selection);
    /* オブジェクト選択は適用後に選択フォーカスを戻すため控える（テキスト編集中は戻さない）
       Save the object selection to restore focus after applying (skip while editing text) */
    var savedSelection = textRange ? null : selection;
    if (progress) { progress.set(0.1, L("progress.analyze")); }
    /* 同一単位のプレビューでは対象集合をキャッシュ再利用（順序変更だけなら作り直さない）
       Reuse the cached target set across previews of the same unit (order-only changes skip the rebuild) */
    var targets = collectTargetsCachedW(items, textRange, options.unit);
    if (targets.length === 0) { restoreBaselineW(); app.redraw(); return "NOSEL"; }
    /* 元状態は初回だけ取得（重い文字読み取りは1回きり）。進捗はここが主コスト
       Snapshot originals only once (the heavy per-char read happens once); this is the main progress cost */
    if (progress) { progress.set(0.2, L("progress.snapshot")); }
    ensureBaselineW(items, textRange, progress ? function (frac) { progress.set(0.2 + frac * 0.6, L("progress.snapshot")); } : null);

    /* 着色前の準備（間引きの要否と着色件数を決める）/ Prepare before coloring (decide decimation and count) */
    var prep = prepareApplyW(items, textRange, options.unit, targets.length, finalApply);
    if (!finalApply) { $.global.__aiApplyWasDecimated = prep.decimated; }

    if (progress) { progress.set(0.85, L("progress.apply")); }
    /* ランダム系（並びシャッフル・完全ランダム抽選・自動CMYK生成）はプレビューと確定で結果が変わらないよう、
       プレビュー時に生成した配色と抽選結果を保存し、確定（finalApply）時はそれを再利用する
       Keep random results stable between preview and commit: generate & store the colors and the
       fully-random draws on preview, then reuse them on commit */
    var colors, groupIndexCache;
    if (finalApply && $.global.__aiApplyColors) {
        colors = $.global.__aiApplyColors;
        groupIndexCache = $.global.__aiApplyGroupIndex || {};
    } else {
        colors = orderColorsW(resolveAppliedColorsW(doc, targets.length, swatchNames), options.order);
        groupIndexCache = {};
        $.global.__aiApplyColors = colors;
        $.global.__aiApplyGroupIndex = groupIndexCache;
    }
    applyColorsToTargetsW(targets, prep.limit, colors, options.order === "fullrandom", prep.decimated, groupIndexCache);
    if (progress) { progress.set(1, L("progress.apply")); }

    /* 適用で変わった選択（フォーカス）を元に戻す / Restore the selection changed by applying */
    if (savedSelection) { try { app.selection = savedSelection; } catch (e) {} }
    app.redraw();
    return "OK";
}

/* 着色前の準備。1文字単位で対象が多いプレビューは間引き（先頭 PREVIEW_CHAR_CAP だけ着色し残りは元色）にし、
   それ以外は線なし・不透明度100 を範囲へ一括設定。戻り値 { decimated, limit } で着色ループを制御する
   Prepare before coloring. Long per-character previews are decimated (color the first PREVIEW_CHAR_CAP,
   leave the rest original); otherwise stroke/opacity is reset in bulk. Returns { decimated, limit } */
function prepareApplyW(items, textRange, unitMode, targetCount, finalApply) {
    var decimated = (!finalApply && unitMode === "character" && targetCount > PREVIEW_CHAR_CAP);
    if (decimated) {
        /* 前回プレビューの着色を元へ戻してから先頭だけ塗る（残りが元の色で見える）
           Repaint originals first, then color just the head (so the tail shows the original color) */
        paintBaselineW();
        return { decimated: true, limit: PREVIEW_CHAR_CAP };
    }
    /* 線なし・不透明度100 は全テキスト対象で共通。文字ごとに書かず範囲へ1回だけ書く
       No stroke / 100% opacity is common to all text targets; write it once per range, not per character */
    resetTextStrokeOpacityW(items, textRange);
    return { decimated: false, limit: targetCount };
}

/* 対象 targets の先頭 limit 件へ色を適用。完全ランダムはグループごとに1回抽選（groupIndexCache に保持）、
   間引き時は文字ごとに線・不透明度も設定
   Color the first `limit` targets. Fully random draws once per group (cached in groupIndexCache);
   when decimated, set stroke/opacity per char too */
function applyColorsToTargetsW(targets, limit, colors, isFullRandom, decimated, groupIndexCache) {
    for (var i = 0; i < limit; i++) {
        var colorIndex;
        if (isFullRandom) {
            var groupKey = (typeof targets[i].groupKey === "string") ? targets[i].groupKey : String(i);
            colorIndex = fullRandomColorIndexW(groupKey, groupIndexCache, colors.length);
        } else {
            colorIndex = (typeof targets[i].colorIndex === "number") ? targets[i].colorIndex : i;
        }
        var color = pickSwatchColorW(colorIndex, colors);
        if (decimated) { applyColorFullToCharW(targets[i], color); }
        else { applyColorToTargetW(targets[i], color); }
    }
}

/* 対象集合を単位ごとにキャッシュ。単位が同じなら作り直さない（順序変更だけのプレビューを速くする）。
   モーダル中は選択が変わらないので live 参照（Character/TextRange 等）も次回まで有効
   Cache the target set per unit; rebuild only when the unit changes (order-only previews stay fast).
   The selection is stable during the modal, so the live references remain valid until the next rebuild */
function collectTargetsCachedW(items, textRange, unitMode) {
    if ($.global.__aiApplyTargetsUnit === unitMode && $.global.__aiApplyTargets) {
        return $.global.__aiApplyTargets;
    }
    var targets = (items.length > 0 || textRange) ? collectColorTargetsByUnitW(items, textRange, unitMode) : [];
    $.global.__aiApplyTargets = targets;
    $.global.__aiApplyTargetsUnit = unitMode;
    return targets;
}

/* 元状態のベースラインを初回だけ取得（全選択ぶん・テキストは文字単位で忠実に保存）。
   単位に依存せず選択全体を控えるので、以後どの単位に切り替えても復元・再取得が不要
   Take the baseline of originals once (whole selection; text per character). Since it covers the
   entire selection regardless of unit, no later restore or re-snapshot is needed */
function ensureBaselineW(items, textRange, onProgress) {
    if ($.global.__aiApplyBaseline) { return; }
    var baseline = [];
    /* 進捗コンテキスト：処理済み文字数を総数で割って onProgress へ通知（テキストが無ければ null）
       Progress context: report processed chars / total to onProgress (null when there is no text) */
    var ctx = onProgress ? { done: 0, total: baselineCharTotalW(items, textRange), report: onProgress } : null;
    if (textRange) { snapshotTargetW({ kind: "textrange", node: textRange }, baseline, ctx); }
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it.typename === "TextFrame") { snapshotTargetW({ kind: "textframe", node: it }, baseline, ctx); }
        else if (it.typename === "PathItem") { snapshotTargetW({ kind: "path", node: it }, baseline, ctx); }
        else if (it.typename === "CompoundPathItem") { snapshotTargetW({ kind: "compound", node: it }, baseline, ctx); }
    }
    $.global.__aiApplyBaseline = baseline;
}

/* ベースライン取得で読む総文字数を見積もる（進捗の分母）/ Estimate total characters read for the baseline (progress denominator) */
function baselineCharTotalW(items, textRange) {
    var total = 0;
    if (textRange) { total += textRange.characters.length; }
    for (var i = 0; i < items.length; i++) {
        if (items[i].typename === "TextFrame") { total += items[i].textRange.characters.length; }
    }
    return total > 0 ? total : 1;
}

/* 現在のプレビューを確定（ベースラインを破棄＝以後 restore しない）。閉じる時・初回開始前に呼ぶ
   Commit the current preview: discard the baseline so it is never restored (called on close / before first preview) */
function workerCommitPreview() {
    $.global.__aiApplyBaseline = null;
    $.global.__aiApplySwatchMap = null;
    $.global.__aiApplyTargets = null;
    $.global.__aiApplyTargetsUnit = null;
    $.global.__aiApplyColors = null;
    $.global.__aiApplyGroupIndex = null;
    $.global.__aiApplyWasDecimated = null;
    return "OK";
}

/* ベースライン（元の塗り/線/不透明度）を画面へ書き戻すが、破棄はしない。
   間引きプレビューで着色しなかった残りを元の色に戻すのに使う
   Repaint the baseline (originals) without discarding it; used by the decimated preview to
   reset the characters it did not color back to their originals */
function paintBaselineW() {
    var baseline = $.global.__aiApplyBaseline;
    if (!baseline) { return; }
    for (var i = 0; i < baseline.length; i++) {
        try { restoreOneSnapshotW(baseline[i]); } catch (e) { }
    }
}

/* ベースラインを書き戻して破棄（キャンセル・Esc・閉じる時。以後 restore しない）
   Repaint the baseline and discard it (Cancel / Esc / close; never restored afterward) */
function restoreBaselineW() {
    paintBaselineW();
    $.global.__aiApplyBaseline = null;
}

/* スナップショット1件を復元 / Restore one snapshot entry */
function restoreOneSnapshotW(entry) {
    var node = entry.node;
    if (entry.kind === "textrun") {
        /* 同色ランを範囲1回で書き戻す（1文字ずつより桁違いに速い）/ Restore a run in one span write (far faster than per character) */
        var range = entry.story.textRange;
        range.start = entry.start; range.end = entry.end;
        range.fillColor = entry.fill; range.strokeColor = entry.stroke; range.opacity = entry.opacity;
    }
    else if (entry.kind === "path") { node.fillColor = entry.fill; node.stroked = entry.stroked; node.opacity = entry.opacity; }
    else if (entry.kind === "compound") {
        for (var i = 0; i < entry.subs.length; i++) {
            var s = entry.subs[i];
            try { s.sub.fillColor = s.fill; s.sub.stroked = s.stroked; s.sub.opacity = s.opacity; } catch (e) { }
        }
    }
}

/* 着色前の対象の元の状態を snapshot 配列へ保存（テキストは文字単位で忠実に保存）
   Snapshot a target's originals into the array (text is snapshotted per character for fidelity) */
function snapshotTargetW(target, out, ctx) {
    var node = target.node;
    if (target.kind === "span") { snapshotCharactersW(spanRangeW(target), out, ctx); }
    else if (target.kind === "textrange") { snapshotCharactersW(node, out, ctx); }
    else if (target.kind === "textframe") { snapshotCharactersW(node.textRange, out, ctx); }
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
function snapshotCharactersW(range, out, ctx) {
    var story = range.story;
    var chars = range.characters;
    var count = chars.length;
    /* 連続範囲なので i 文字目の story オフセットは range.start + i（ch.start の DOM 読み取りを省く）
       The range is contiguous, so character i's story offset is range.start + i (avoids reading ch.start) */
    var rangeStart = range.start;
    var started = false;
    var runFill = null, runStroke = null, runOpacity = 0, runKey = null, runStart = 0, runEnd = 0;
    for (var i = 0; i < count; i++) {
        var ch = chars[i];
        var fill = ch.fillColor;
        var stroke = ch.strokeColor;
        var opacity = ch.opacity;
        var chStart = rangeStart + i;
        /* 一定文字ごとに進捗を通知（毎回だと更新自体が重い）/ Report progress every N chars (updating each time is itself costly) */
        if (ctx) { ctx.done++; if ((ctx.done % 200) === 0) { ctx.report(ctx.done / ctx.total); } }
        var key = colorKeyW(fill) + "|" + colorKeyW(stroke) + "|" + opacity;
        if (started && key === runKey) {
            /* 見た目が同じ連続文字は run を伸ばす / extend the run for contiguous identical characters */
            runEnd = chStart + 1;
        } else {
            if (started) { out.push({ kind: "textrun", story: story, start: runStart, end: runEnd, fill: runFill, stroke: runStroke, opacity: runOpacity }); }
            started = true;
            runKey = key; runFill = fill; runStroke = stroke; runOpacity = opacity;
            runStart = chStart; runEnd = chStart + 1;
        }
    }
    if (started) { out.push({ kind: "textrun", story: story, start: runStart, end: runEnd, fill: runFill, stroke: runStroke, opacity: runOpacity }); }
}

/* 色の同一性キー（NoColor 対応。塗り・線の run 判定に使う）/ Identity key for a color (NoColor-aware; used for run detection) */
function colorKeyW(color) {
    if (color.typename === "NoColor") { return "N"; }
    return serializeColorW(color);
}

/* 選択情報を返す "OK|<defaultUnit>|<r,g,b;...>|<encName,...>|<itemCount>|<isText>|<paragraphCount>|<textObjectCount>" / Return selection info */
function workerReadSelection() {
    if (app.documents.length === 0) return "NODOC";
    var doc = app.activeDocument;
    var items = flattenSelectionW(app.selection);
    var textRange = getSingleSelectedTextRangeW(app.selection);
    var defaultUnit = "object";
    if (textRange) { defaultUnit = "character"; }
    else if (items.length === 1 && items[0].typename === "TextFrame") { defaultUnit = "character"; }
    var swatches = doc.swatches.getSelected();
    var usable = swatches && swatches.length >= 1 && !allWhiteSwatchesW(swatches);
    var chipParts = [];
    var nameParts = [];
    if (usable) {
        for (var i = 0; i < swatches.length; i++) {
            chipParts.push(colorToRGB255W(swatches[i].color));
            nameParts.push(encodeURIComponent(swatches[i].name));
        }
    }
    var stats = selectionTextStatsW(items, textRange);
    return "OK|" + defaultUnit + "|" + chipParts.join(";") + "|" + nameParts.join(",") + "|" + items.length + "|" + (stats.isText ? "1" : "0") + "|" + stats.paragraphCount + "|" + stats.textObjectCount;
}

/* 選択のテキスト統計（テキスト有無・段落数・テキストオブジェクト数）。
   段落数はテキストが1つのときだけ数える（ディム判定がその場合しか参照せず、長文では paragraphs.length が重いため）
   Text stats; paragraphs are counted only for a single text object (the only case the dim logic reads it; costly on long text) */
function selectionTextStatsW(items, textRange) {
    var isText = false;
    var textObjectCount = 0;
    var singleTextRange = null;
    if (textRange) { isText = true; textObjectCount++; singleTextRange = textRange; }
    for (var i = 0; i < items.length; i++) {
        if (items[i].typename === "TextFrame") { isText = true; textObjectCount++; singleTextRange = items[i].textRange; }
    }
    /* 段落数も \r 基準で数える（paragraphs.length は強制改行でも増えるので、
       強制改行入りの1段落が「2段落」に見えて段落単位が有効になってしまう）
       Count paragraphs by \r as well: paragraphs.length also counts forced line breaks,
       which would make a single paragraph look like two and wrongly enable the paragraph unit */
    var paragraphCount = (textObjectCount === 1) ? countParagraphsW(singleTextRange) : 0;
    return { isText: isText, paragraphCount: paragraphCount, textObjectCount: textObjectCount };
}

/* 選択をフラット化して色付け対象のみ収集 / Flatten selection to colorable items */
function flattenSelectionW(selection) {
    var collected = [];
    if (!selection || selection.length === 0) return collected;
    for (var i = 0; i < selection.length; i++) { collectColorableItemsW(selection[i], collected); }
    return collected;
}

/* 色付け可能な要素を再帰収集（グループ内も）/ Recursively collect colorable items */
function collectColorableItemsW(item, out) {
    if (!item) return;
    if (item.typename === "PathItem" || item.typename === "CompoundPathItem" || item.typename === "TextFrame") { out.push(item); return; }
    if (item.typename === "GroupItem") {
        var children = item.pageItems;
        for (var i = 0; i < children.length; i++) { collectColorableItemsW(children[i], out); }
    }
}

/* テキスト編集中の選択範囲を取得 / Get the active text range while editing */
function getSingleSelectedTextRangeW(selection) {
    if (!selection || selection.length !== 1) return null;
    if (selection[0] && selection[0].typename === "TextRange") return selection[0];
    return null;
}

/* 適用単位に応じた対象を収集 / Collect targets by apply unit */
function collectColorTargetsByUnitW(items, textRange, unitMode) {
    var targets = [];
    if (textRange) { pushTextRangeTargetsW(textRange, unitMode, targets); return targets; }
    if (items.length === 1 && items[0].typename === "TextFrame") {
        if (unitMode === "object") { targets.push({ kind: "textframe", node: items[0] }); }
        else { pushTextRangeTargetsW(items[0].textRange, unitMode, targets); }
        return targets;
    }
    var sorted = items.slice();
    sortByPositionW(sorted);
    for (var i = 0; i < sorted.length; i++) {
        var it = sorted[i];
        if (it.typename === "TextFrame" && unitMode !== "object") { pushTextRangeTargetsW(it.textRange, unitMode, targets); }
        else if (it.typename === "PathItem") { targets.push({ kind: "path", node: it }); }
        else if (it.typename === "CompoundPathItem") { targets.push({ kind: "compound", node: it }); }
        else if (it.typename === "TextFrame") { targets.push({ kind: "textframe", node: it }); }
    }
    return targets;
}

/* テキスト範囲を単位ごとに分割してターゲット化 / Split a text range into per-unit targets */
function pushTextRangeTargetsW(textRange, unitMode, out) {
    if (unitMode === "word") { pushStaggeredWordTargetsW(textRange, out); return; }
    if (unitMode === "paragraph") { pushParagraphTargetsW(textRange, out); return; }
    var ranges = getTextUnitRangesW(textRange, unitMode);
    if (!ranges) { out.push({ kind: "textrange", node: textRange }); return; }
    for (var i = 0; i < ranges.length; i++) { out.push({ kind: "textrange", node: ranges[i] }); }
}

/* 単語ごと：Illustrator の単語境界で単語スパンを作る（各行先頭は互い違い）。
   1文字ずつ塗ると文字数の二乗で重くなるため、単語スパン1回の書き込みにまとめている。
   スパンは次の単語の開始まで伸ばし、句読点・スペースを直前の単語色にする
   （行頭は先頭の単語色、行末の残りは最後の単語色。1文字ずつ塗っていた頃と同じ結果）
   Per word: build one span per word (staggered per line). Coloring per character costs
   roughly quadratic time, so each word is written once instead. Each span runs to the next
   word's start so punctuation and spaces take the preceding word's colour, matching the
   per-character behaviour this replaced */
function pushStaggeredWordTargetsW(textRange, out) {
    var story = textRange.story;
    var lines = textRange.lines;
    var lineCount = lines.length;
    for (var li = 0; li < lineCount; li++) {
        var line = lines[li];
        var words = line.words;
        var wordCount = words.length;
        for (var wi = 0; wi < wordCount; wi++) {
            /* 行頭の記号類は先頭の単語に、行末の残りは最後の単語に含める
               Leading symbols join the first word; the tail of the line joins the last word */
            var spanStart = (wi === 0) ? line.start : words[wi].start;
            var spanEnd = (wi === wordCount - 1) ? line.end : words[wi + 1].start;
            if (spanEnd <= spanStart) { continue; }
            /* groupKey は行内で一意（完全ランダムで単語ごとに別色を抽選するため）
               groupKey is unique per line+word so fully random draws a distinct color per word */
            out.push({ kind: "span", story: story, start: spanStart, end: spanEnd, colorIndex: wi + li, groupKey: li + ":" + wi });
        }
    }
}

/* 段落ごと：本文を \r で走査してスパンを作る。
   textRange.paragraphs は強制改行（）でも区切られてしまい、1段落が複数に割れるため使わない
   Per paragraph: scan the story text for \r. textRange.paragraphs also splits on a forced
   line break (), which would break one paragraph into several, so it is not used */
function pushParagraphTargetsW(textRange, out) {
    var story = textRange.story;
    var storyText = story.textRange.contents;
    var from = textRange.start;
    var to = textRange.end;
    var head = from;
    while (head < to) {
        var returnAt = storyText.indexOf("\r", head);
        /* 改行が無い、または選択範囲の外なら、そこが最後の段落 / No return, or one past the selection, ends the last paragraph */
        var tail = (returnAt === -1 || returnAt >= to) ? to : returnAt;
        /* 空段落（改行だけ）は塗る文字が無いので飛ばす / An empty paragraph has nothing to color */
        if (tail > head) { out.push({ kind: "span", story: story, start: head, end: tail }); }
        if (returnAt === -1 || returnAt >= to) break;
        head = returnAt + 1;
    }
}

/* 範囲内の段落数を \r 基準で数える（塗る文字が無い空段落は数えない）
   Count paragraphs in a range by \r (an empty paragraph has nothing to color, so it is not counted) */
function countParagraphsW(textRange) {
    var spans = [];
    pushParagraphTargetsW(textRange, spans);
    return spans.length;
}

/* 単位に対応するテキスト範囲コレクション（word/paragraph は pushTextRangeTargetsW で先に処理）
   Text range collection for a unit (word/paragraph are handled earlier in pushTextRangeTargetsW) */
function getTextUnitRangesW(textRange, unitMode) {
    if (unitMode === "character") return textRange.characters;
    if (unitMode === "line") return textRange.lines;
    return null;
}

/* スパン（story 内の start〜end）に対応する TextRange を作る。
   story.textRange は呼ぶたび独立したオブジェクトを返すので、start/end を書き換えて使える
   Build a TextRange for a span: story.textRange hands back an independent object each call,
   so its start/end can be rewritten safely */
function spanRangeW(target) {
    var range = target.story.textRange;
    range.start = target.start;
    range.end = target.end;
    return range;
}

/* テキストの線・不透明度をまとめて初期化（線なし・不透明度100）。
   全テキスト対象で同じ値なので、対象範囲へ1回だけ書き込む（1文字ずつ書くと重い）
   Reset stroke/opacity for text in bulk (no stroke, 100% opacity); the value is the same for every
   text target, so write it once per range instead of per character */
function resetTextStrokeOpacityW(items, textRange) {
    if (textRange) { applyNoStrokeFullOpacityW(textRange); }
    for (var i = 0; i < items.length; i++) {
        if (items[i].typename === "TextFrame") { applyNoStrokeFullOpacityW(items[i].textRange); }
    }
}

/* テキスト範囲に「線なし・不透明度100」を1回で設定 / Set no-stroke, 100% opacity on a range in one call */
function applyNoStrokeFullOpacityW(range) {
    range.strokeColor = new NoColor();
    range.opacity = 100;
}

/* 間引きプレビュー用：1文字に塗り・線なし・不透明度100 をまとめて設定（着色しない残りは元色のまま）。
   全体の resetTextStrokeOpacityW を使わず文字単位で設定するのは、着色範囲外の元の線・不透明度を保つため
   For decimated preview: set fill + no-stroke + 100% opacity on a single character. Done per character
   (not via the range-wide resetTextStrokeOpacityW) so the uncolored tail keeps its original stroke/opacity */
function applyColorFullToCharW(target, color) {
    var node = target.node;
    node.fillColor = color;
    applyNoStrokeFullOpacityW(node); /* Character も strokeColor/opacity を持つので流用 / a Character also has strokeColor/opacity */
}

/* ターゲットに塗りを設定（テキストの線・不透明度は resetTextStrokeOpacityW でまとめて処理済み）
   Set the fill on a target (text stroke/opacity are handled in bulk by resetTextStrokeOpacityW) */
function applyColorToTargetW(target, color) {
    var node = target.node;
    if (target.kind === "span") { spanRangeW(target).fillColor = color; }
    else if (target.kind === "textrange") { node.fillColor = color; }
    else if (target.kind === "textframe") { node.textRange.fillColor = color; }
    else if (target.kind === "path") { node.fillColor = color; node.stroked = false; node.opacity = 100; }
    else if (target.kind === "compound") {
        var subs = node.pathItems;
        for (var i = 0; i < subs.length; i++) { subs[i].fillColor = color; subs[i].stroked = false; subs[i].opacity = 100; }
    }
}

/* 位置順にソート / Sort by position */
function sortByPositionW(items) {
    var minLeft = Infinity, maxLeft = -Infinity, minTop = Infinity, maxTop = -Infinity;
    for (var i = 0; i < items.length; i++) {
        var left = items[i].left;
        var top = items[i].top;
        if (left < minLeft) minLeft = left;
        if (left > maxLeft) maxLeft = left;
        if (top < minTop) minTop = top;
        if (top > maxTop) maxTop = top;
    }
    if (maxLeft - minLeft > maxTop - minTop) { items.sort(function (a, b) { return comparePositionKeysW(a.left, b.left, b.top, a.top); }); }
    else { items.sort(function (a, b) { return comparePositionKeysW(b.top, a.top, a.left, b.left); }); }
}

/* ソート用比較（主キー→副キー）/ Compare by primary then secondary key */
function comparePositionKeysW(primaryA, primaryB, secondaryA, secondaryB) {
    return primaryA == primaryB ? secondaryA - secondaryB : primaryA - primaryB;
}

/* ランダムシャッフルして返す / Return a shuffled copy */
function shuffleArrayW(source) {
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
function randIntW(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

/* 使用するカラーを決定（取り込んだスウォッチ名を最優先、無ければ自動カラー）
   Resolve colors (captured swatch names take priority; otherwise auto colors) */
function resolveAppliedColorsW(doc, targetCount, swatchNames) {
    var colors = [];
    if (swatchNames && swatchNames.length > 0) {
        /* 名前→色マップはダイアログ中1回だけ構築してキャッシュ（モーダル中スウォッチは変わらない）。
           プレビューは頻繁に走るので毎回の全スウォッチ走査を避ける
           Build the name->color map once per dialog and cache it (swatches never change during the
           modal); the preview fires often, so avoid rescanning all swatches every time */
        var swatchColorsByName = $.global.__aiApplySwatchMap || ($.global.__aiApplySwatchMap = buildSwatchColorMapW(doc));
        for (var i = 0; i < swatchNames.length; i++) {
            var swatchColor = swatchColorsByName["$" + decodeURIComponent(swatchNames[i])];
            if (swatchColor) { colors.push(swatchColor); }
        }
    }
    if (colors.length >= 1) { return colors; }
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) { return generateRandomCMYPaletteUniqueW(targetCount, TMK_CMYK_FALLBACK_MAX_TOTAL); }
    return getDefaultRGBColorsW();
}

/* スウォッチ名→色のマップを生成（選択状態に依存しない）。キーは "$"+名前でプロトタイプ汚染を回避
   Build a name->color map (independent of selection state); keys are "$"+name to dodge prototype collisions */
function buildSwatchColorMapW(doc) {
    var map = {};
    var swatches = doc.swatches;
    for (var i = 0; i < swatches.length; i++) {
        var key = "$" + swatches[i].name;
        if (map[key] === undefined) { map[key] = swatches[i].color; }
    }
    return map;
}

/* カラーを同一性キー文字列にシリアライズ（colorKeyW の run 判定に使用）。
   SpotColor は tint も含める（tint 違いを別色として区別するため）
   Serialize a color to an identity key (used by colorKeyW for run detection);
   SpotColor includes tint so different tints are treated as distinct colors */
function serializeColorW(color) {
    var t = color.typename;
    if (t === "CMYKColor") { return "C," + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black; }
    if (t === "RGBColor") { return "R," + color.red + "," + color.green + "," + color.blue; }
    if (t === "GrayColor") { return "G," + color.gray; }
    if (t === "SpotColor") {
        var tint = (typeof color.tint === "number") ? color.tint : 100;
        return "S," + tint + "," + serializeColorW(color.spot.color);
    }
    return "R,128,128,128";
}

/* 適用順に並べ替え（完全ランダムは適用側でインデックスを抽選するので並びは変えない）
   Reorder colors (fully random draws indices at apply time, so the order is left as-is) */
function orderColorsW(colors, orderMode) {
    if (orderMode === "reverse") return colors.slice().reverse();
    if (orderMode === "random") return shuffleArrayW(colors);
    return colors;
}

/* 完全ランダム用：グループキーごとのカラーインデックスを抽選してキャッシュ
   Fully random: draw and cache a color index per group key */
function fullRandomColorIndexW(groupKey, cache, colorCount) {
    if (cache[groupKey] === undefined) { cache[groupKey] = randIntW(0, colorCount - 1); }
    return cache[groupKey];
}

/* インデックスに応じて色を取得（末尾を超えたらループ）。不透明度は適用側で設定する
   Pick a color by index (wraps past the end); opacity is set on the target when applying */
function pickSwatchColorW(index, colors) {
    return colors[index % colors.length];
}

/* CMYK カラーを生成 / Build a CMYKColor */
function buildCMYKColorW(cyan, magenta, yellow, black) {
    var color = new CMYKColor();
    color.cyan = cyan; color.magenta = magenta; color.yellow = yellow; color.black = black;
    return color;
}

/* RGB カラーを生成 / Build an RGBColor */
function buildRGBColorW(red, green, blue) {
    var color = new RGBColor();
    color.red = red; color.green = green; color.blue = blue;
    return color;
}

/* RGB ドキュメント用の定義済みカラー / Predefined colors for RGB documents */
function getDefaultRGBColorsW() {
    var defs = [[222, 84, 25], [245, 233, 40], [41, 163, 57], [53, 157, 209], [173, 127, 71], [238, 176, 51]];
    var colors = [];
    for (var i = 0; i < defs.length; i++) { colors.push(buildRGBColorW(defs[i][0], defs[i][1], defs[i][2])); }
    return colors;
}

/* CM/CY/MY の2チャンネル（K=0）をランダムに1色ぶん生成 / Pick a random 2-channel CMY color */
function pickTwoChannelCMYW(channelPair, maxTotal) {
    var amountA = randIntW(1, Math.min(100, maxTotal - 1));
    var amountBMax = Math.min(100, maxTotal - amountA);
    if (amountBMax < 1) return null;
    var amountB = randIntW(1, amountBMax);
    if (channelPair === "CM") return { cyan: amountA, magenta: amountB, yellow: 0 };
    if (channelPair === "CY") return { cyan: amountA, magenta: 0, yellow: amountB };
    return { cyan: 0, magenta: amountA, yellow: amountB };
}

/* CMYK 用：CM/CY/MY のみで可能な限り重複しない色を生成 / Generate CM/CY/MY colors, avoid duplicates */
function generateRandomCMYPaletteUniqueW(count, maxTotal) {
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
        if ((attempts % 37) === 0) { channelPairs = shuffleArrayW(channelPairs); }
        var channelPair = channelPairs[palette.length % channelPairs.length];
        var channels = pickTwoChannelCMYW(channelPair, maxTotal);
        if (!channels) continue;
        var key = channels.cyan + "," + channels.magenta + "," + channels.yellow;
        if (!allowDuplicate && seen[key]) continue;
        if (minDistance > 0 && !isFarEnoughCMYW(channels.cyan, channels.magenta, channels.yellow, accepted, minDistance)) continue;
        seen[key] = true;
        var color = buildCMYKColorW(channels.cyan, channels.magenta, channels.yellow, 0);
        palette.push(color);
        accepted.push(color);
    }
    return palette;
}

/* 既存の採用色から十分離れているか / Whether far enough from accepted colors */
function isFarEnoughCMYW(cyan, magenta, yellow, existing, minDistance) {
    for (var i = 0; i < existing.length; i++) {
        if (cmyDistanceW(cyan, magenta, yellow, existing[i].cyan, existing[i].magenta, existing[i].yellow) < minDistance) return false;
    }
    return true;
}

/* CMY のマンハッタン距離 / Manhattan distance in CMY */
function cmyDistanceW(c1, m1, y1, c2, m2, y2) {
    return Math.abs(c1 - c2) + Math.abs(m1 - m2) + Math.abs(y1 - y2);
}

/* 色が白か判定 / Whether a color is white */
function isWhiteColorW(color) {
    if (color.typename === "CMYKColor") return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    if (color.typename === "RGBColor") return color.red === 255 && color.green === 255 && color.blue === 255;
    return false;
}

/* 全スウォッチが白のみか判定 / Whether all swatches are white */
function allWhiteSwatchesW(swatches) {
    for (var i = 0; i < swatches.length; i++) {
        if (!isWhiteColorW(swatches[i].color)) return false;
    }
    return true;
}

/* Illustrator カラーを "r,g,b"(0..255) 文字列へ / Convert an Illustrator color to an "r,g,b" string */
function colorToRGB255W(color) {
    var t = color.typename;
    if (t === "RGBColor") return Math.round(color.red) + "," + Math.round(color.green) + "," + Math.round(color.blue);
    if (t === "CMYKColor") {
        var k = 1 - color.black / 100;
        return Math.round(255 * (1 - color.cyan / 100) * k) + "," + Math.round(255 * (1 - color.magenta / 100) * k) + "," + Math.round(255 * (1 - color.yellow / 100) * k);
    }
    if (t === "GrayColor") {
        var v = Math.round(255 * (1 - color.gray / 100));
        return v + "," + v + "," + v;
    }
    if (t === "SpotColor") {
        var base = colorToRGB255W(color.spot.color).split(",");
        var tint = (typeof color.tint === "number") ? color.tint / 100 : 1;
        return Math.round(255 - (255 - parseInt(base[0], 10)) * tint) + "," + Math.round(255 - (255 - parseInt(base[1], 10)) * tint) + "," + Math.round(255 - (255 - parseInt(base[2], 10)) * tint);
    }
    return "128,128,128";
}

showDialog();
