#target illustrator
#targetengine "AiApplySwatchesToSelection"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択したオブジェクトやテキストに、スウォッチや定義済みカラーを適用する常駐パレットです。
- 適用単位（オブジェクト／1文字／単語／行／段落）と適用順（そのまま／逆順／ランダム／完全ランダム）をラジオボタンで選択します。
- 「ランダム」はカラーの並びをシャッフルして繰り返し適用、「完全ランダム」は適用先ごとに毎回抽選するため繰り返しがありません。
- ラジオを変えるたびにライブプレビュー。適用ボタンはなく、閉じた時点の結果が確定します（取り消しは Cmd+Z）。
- 開いた時点の選択スウォッチを■で取り込み、以後はそれ（スウォッチ名で参照）を最優先で使用。ラジオ操作でスウォッチ選択が外れても影響しません。1色でも選択があれば優先します。
- スウォッチを選び直したら［アップデート］で再取り込みします。
- スウォッチ未選択時は自動カラー（CMYK は CM/CY/MY 生成、RGB は既定色）を使用。
- 「単語」は各行の先頭色が互い違いになるよう配色。
- DOM 操作はメインエンジンへ BridgeTalk 委譲（常駐パレットは表示中に DOM 接続を失うため）。

### 紹介記事（note）

https://note.com/dtp_tranist/n/n5602f3084d2b

### Overview

- A persistent palette that applies swatches or predefined colors to selected objects or text.
- Choose apply unit (object / character / word / line / paragraph) and order (as-is / reverse / random / fully random) with radio buttons.
- "Random" shuffles the color list once and cycles through it; "Fully random" draws a color per target, so no sequence repeats.
- Live preview on every change. There is no Apply button — the result is committed when the palette closes (undo via Cmd+Z).
- The swatches selected when the palette opens are captured as chips and used (by swatch name) with top priority, so losing the swatch selection when clicking radios has no effect. Even a single selected swatch is used.
- Use [Update] to re-capture after reselecting swatches.
- When no swatches are selected, auto colors are used (CMYK CM/CY/MY generation, RGB defaults).
- "Per word" staggers colors so each line starts on a different color.
- DOM work is delegated to the main engine via BridgeTalk (a persistent palette loses its DOM connection while shown).

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.8.0";

// =========================================
// ユーザー設定 / User settings
// =========================================
/* CMYK フォールバック生成の設定 / CMYK fallback generation settings */
var TMK_CMYK_FALLBACK_MAX_TOTAL = 200;    /* 合計上限 C+M / C+Y / M+Y / total channel limit */
var TMK_CMYK_FALLBACK_MIN_DISTANCE = 35;  /* 近似色回避の最小距離 / min distance to avoid similar colors */
var COLOR_CHIPS_PER_ROW = 12;             /* 1行あたりのチップ数 / chips per row */
var CHIP_SIZE = 18;                       /* チップの一辺(px) / chip size */
var CHIP_GAP = 4;                         /* チップ間の間隔(px) / gap between chips */
var PREVIEW_CHAR_CAP = 500;               /* 1文字単位プレビューで着色する最大文字数（超過分は閉じる時に着色）/ max chars colored in per-character preview */

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
        fromSwatches: { ja: "スウォッチを読込", en: "Load Swatches" },
        fromObjects:  { ja: "塗り色を読込", en: "Load Fills" },
        openSwatches: { ja: "「スウォッチ」パネルを開く", en: "Open Swatches Panel" }
    },
    tooltip: {
        fromSwatches:  { ja: "現在選択しているスウォッチを読み込み直します。パレットを閉じるには Esc キーを押します。", en: "Reload the currently selected swatches. Press Esc to close the palette." },
        fromObjects:   { ja: "現在選択しているオブジェクトの塗り色を読み込み直します。パレットを閉じるには Esc キーを押します。", en: "Reload the fill colors of the currently selected objects. Press Esc to close the palette." },
        colors:        { ja: "適用に使う配色。下のボタンで取り込み直せます。", en: "Colors used for applying; re-capture with the buttons below." },
        unitObject:    { ja: "選択オブジェクト単位でカラーを順番に適用します。", en: "Applies colors in order, one per selected object." },
        unitCharacter: { ja: "テキストを1文字ずつ分けてカラーを適用します。", en: "Applies a color to each character of the text." },
        unitWord:      { ja: "英文テキストを単語ごとに分けてカラーを適用します。英文以外では単語が正しく分割されないことがあります。", en: "Applies a color to each word of English text. Words may not split correctly for non-English text." },
        unitLine:      { ja: "テキストの行ごとにカラーを適用します。", en: "Applies a color to each line of the text." },
        unitParagraph: { ja: "テキストの段落ごとにカラーを適用します。", en: "Applies a color to each paragraph of the text." },
        orderAsis:     { ja: "取り込んだカラーの並び順で適用します（適用先は位置順・文字順）。", en: "Applies colors in the captured order (targets follow position / reading order)." },
        orderReverse:  { ja: "取り込んだカラーの並びを逆にして適用します。", en: "Applies the captured colors in reverse order." },
        orderRandom:     { ja: "取り込んだカラーの並びをランダムにして適用します（並びは繰り返します）。", en: "Applies the captured colors in a random order (the sequence repeats)." },
        orderFullRandom: { ja: "適用先ごとにカラーを毎回ランダムに選びます（並びは繰り返しません）。", en: "Picks a color at random for each target (no repeating sequence)." },
        openSwatches:    { ja: "Illustrator の「スウォッチ」パネルを表示します。", en: "Shows Illustrator's Swatches panel." }
    },
    progress: {
        title:    { ja: "読み込み中", en: "Loading" },
        reading:  { ja: "選択情報を読み込み中…", en: "Reading selection…" },
        applying: { ja: "プレビューを適用中…", en: "Applying preview…" }
    },
    note: {
        autoColor: {
            ja: "カラー未選択：自動カラーを使用します",
            en: "No colors selected: using auto colors"
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
// 常駐エンジンの永続状態 / Persistent engine state
// =========================================
/* 再実行でも消えないよう $.global に保持（GC回避・多重起動防止）。プレビューの復元情報（ベースライン）は
   メインエンジン側の $.global.__aiApplyBaseline に持つ（app.undo は使わない）
   Kept on $.global so re-runs don't reset it. Preview-restore data (baseline) lives on the main engine */
if (typeof $.global.__aiApplySwatchesState === "undefined") {
    $.global.__aiApplySwatchesState = { win: null, busy: false };
}
var STATE = $.global.__aiApplySwatchesState;

// =========================================
// パレット / Palette (UI engine)
// =========================================
/* 常駐パレットを表示（既存があれば閉じてから）/ Show the palette (closing any existing one first) */
function showPalette() {
    if (STATE.win) {
        try { STATE.win.close(); } catch (e) { }
        STATE.win = null;
    }
    stopDimPolling();

    /* 読み込み中プログレス（読み込み・初回適用が重い場合の表示）/ Loading progress (for slow read/first apply) */
    var progress = openProgressWindow();

    /* 選択情報を先に取得（初期単位・チップ・単語ディム判定）/ Read selection info first */
    var info = readSelectionInfo();
    setProgress(progress, 45, L("progress.reading"));

    var win = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
    setupWindow(win);
    STATE.win = win;

    /* 取り込んだ配色（適用に使用。ラジオ操作で選択が外れても影響しない）。名前 or 値のどちらか一方
       Captured colors for applying (unaffected by losing the selection): either swatch names OR serialized values */
    var loadedSwatchNames = info.swatchNames;
    var loadedColorValues = [];

    /* 適用するカラー（配色チップ）＋取り込みボタン / Color chips + capture buttons */
    var colorPanel = win.add("panel", undefined, L("panel.colors"));
    setupPanel(colorPanel);
    colorPanel.helpTip = L("tooltip.colors");
    var chipHost = colorPanel.add("group");
    chipHost.orientation = "column";
    chipHost.alignChildren = ["left", "top"];
    buildColorChips(chipHost, info.chips);
    var buttonRow = colorPanel.add("group");
    setupRow(buttonRow, "left");
    buttonRow.margins = [0, 5, 0, 0]; /* ボタン上に余白 / margin above the buttons */
    var fromSwatchesButton = buttonRow.add("button", undefined, L("button.fromSwatches"));
    fromSwatchesButton.helpTip = L("tooltip.fromSwatches");
    var fromObjectsButton = buttonRow.add("button", undefined, L("button.fromObjects"));
    fromObjectsButton.helpTip = L("tooltip.fromObjects");

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
    /* ポーリングから参照するために保持＋現在の署名を記録 / Expose radios and record current signature for polling */
    STATE.unitRadioButtons = unitRadioButtons;
    STATE.lastStatsSig = selectionStatsSignature(info);

    /* 配色順（カラム内は縦並び）/ Coloring order (vertical within its column) */
    var orderRadioButtons = addRadioPanel(unitsRow, L("panel.option"), buildOptionList("order", ["asis", "reverse", "random", "fullrandom"]), "asis", onOptionChange);
    setOptionTooltips(orderRadioButtons, {
        asis: L("tooltip.orderAsis"),
        reverse: L("tooltip.orderReverse"),
        random: L("tooltip.orderRandom"),
        fullrandom: L("tooltip.orderFullRandom")
    });

    /* パレット下部：スウォッチパネルを開く / Footer: open the Swatches panel */
    var footerRow = win.add("group");
    setupRow(footerRow, "left");
    footerRow.margins = [0, 5, 0, 0]; /* ボタン上に余白 / margin above the button */
    var openSwatchesButton = footerRow.add("button", undefined, L("button.openSwatches"));
    openSwatchesButton.helpTip = L("tooltip.openSwatches");
    openSwatchesButton.onClick = function () {
        runWorker("workerOpenSwatchesPanel()");
    };

    /* 現在の設定を取得 / Read current settings */
    function currentOptions() {
        return {
            unit: getSelectedOption(unitRadioButtons, "object"),
            order: getSelectedOption(orderRadioButtons, "asis")
        };
    }

    /* プレビュー適用をメインエンジンへ委譲（取り込み済みの配色＝名前 or 値を渡す）。
       前回プレビューの取り消しは worker がスナップショット復元で行うのでフラグ管理は不要
       Delegate a preview apply to the main engine; the worker reverts the previous preview via snapshot */
    function runPreview() {
        var call = "workerApply(" + optionsToLiteral(currentOptions()) +
            ", " + stringsToLiteral(loadedSwatchNames) + ", " + stringsToLiteral(loadedColorValues) + ")";
        runWorker(call);
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

    /* 選択スウォッチから配色を取り込んで適用（外部カラーなのでプレビュー表示）
       Capture from swatches and apply (external colors → show preview) */
    function onFromSwatches() {
        var fresh = readSelectionInfo();
        loadedSwatchNames = fresh.swatchNames;
        loadedColorValues = [];
        refreshColorUI(fresh);
        runPreview();
    }
    fromSwatchesButton.onClick = onFromSwatches;

    /* 選択オブジェクトの塗り色を取り込む。この時点では適用しない（＝オブジェクトの色を変えない）。
       配色は配色単位を選んだときに適用される
       Capture fills from objects; do NOT apply here (keep objects unchanged); applied when a unit is chosen */
    function onFromObjects() {
        var fresh = readObjectColorsInfo();
        loadedColorValues = fresh.colorValues;
        loadedSwatchNames = [];
        refreshColorUI(fresh);
    }
    fromObjectsButton.onClick = onFromObjects;

    /* Esc で閉じる / Close on Esc */
    win.addEventListener("keydown", function (k) {
        if (k.keyName === "Escape") win.close();
    });

    /* 選択を変えてパレットが前面に戻ったら、配色単位のディムを現在の選択で更新（配色は変えない）
       When focus returns to the palette after changing the selection, refresh the unit dims (colors unchanged) */
    win.onActivate = function () {
        if (STATE.busy) return;
        var freshStats = readSelectionStats();
        if (freshStats.ok) {
            updateUnitAvailability(unitRadioButtons, freshStats);
            STATE.lastStatsSig = selectionStatsSignature(freshStats);
        }
    };

    /* 適用ボタンなし：閉じたら現在の結果を確定。1文字間引きプレビューだった場合は全文字をフル着色してから確定する
       （プレビューと同じ配色を保存値から再利用）。スナップショットは破棄＝以後 restore しない（戻すのは Cmd+Z）
       No Apply button: closing commits. If the preview was decimated (per character), color everything
       first (reusing the saved colors so it matches the preview), then discard the snapshot (undo via Cmd+Z) */
    win.onClose = function () {
        stopDimPolling();
        var call = "workerFinalize(" + optionsToLiteral(currentOptions()) +
            ", " + stringsToLiteral(loadedSwatchNames) + ", " + stringsToLiteral(loadedColorValues) + ")";
        runWorker(call);
        STATE.win = null;
        STATE.unitRadioButtons = null;
        return true;
    };

    setProgress(progress, 75, L("progress.applying"));
    win.show();
    /* レイアウト確定後に取り込みボタンの天地を詰める / Trim the capture buttons' height after layout */
    trimButtonHeight(fromSwatchesButton, 2);
    trimButtonHeight(fromObjectsButton, 2);
    trimButtonHeight(openSwatchesButton, 2);
    /* 初回プレビュー前に前セッションのスナップショットを破棄（現在の状態を基準にする）
       Discard any previous-session snapshot before the first preview (current doc becomes the baseline) */
    runWorker("workerCommitPreview()");
    /* 表示後に初回プレビュー（パレットは非モーダルなので show 後でよい）/ First preview after show */
    runPreview();
    setProgress(progress, 100, L("progress.applying"));
    closeProgressWindow(progress);
    /* 選択変更に追従してディムを更新するポーリングを開始 / Start polling to follow selection changes */
    startDimPolling();
}

/* 読み込み中プログレスウィンドウを表示（表示は任意：レイアウト失敗時は null で続行）
   Show a loading-progress window (optional: return null and continue if layout fails) */
function openProgressWindow() {
    var progressWindow = null;
    try {
        progressWindow = new Window("palette", L("progress.title"), undefined, { closeButton: false });
        progressWindow.orientation = "column";
        progressWindow.alignChildren = "fill";
        progressWindow.margins = 16;
        progressWindow.spacing = 8;
        var message = progressWindow.add("statictext", undefined, L("progress.reading"));
        message.preferredSize.width = 240;
        var bar = progressWindow.add("progressbar", undefined, 0, 100);
        bar.preferredSize = [240, 8];
        progressWindow.progressMessage = message;
        progressWindow.progressBar = bar;
        /* show 前に明示レイアウトして "Window layout failed: size" を回避／捕捉
           Lay out explicitly before show to avoid/catch "Window layout failed: size" */
        progressWindow.layout.layout(true);
        progressWindow.show();
        forceWindowUpdate(progressWindow);
    } catch (e) {
        /* プログレスは装飾。表示に失敗しても本体パレットは開く / Progress is cosmetic; keep the main palette opening */
        if (progressWindow) { try { progressWindow.close(); } catch (e2) { } }
        progressWindow = null;
    }
    return progressWindow;
}

/* プログレスの値・メッセージを更新 / Update progress value and message */
function setProgress(progressWindow, value, message) {
    if (!progressWindow) return;
    progressWindow.progressBar.value = value;
    if (message) progressWindow.progressMessage.text = message;
    forceWindowUpdate(progressWindow);
}

/* プログレスウィンドウを閉じる / Close the progress window */
function closeProgressWindow(progressWindow) {
    if (progressWindow) {
        try { progressWindow.close(); } catch (e) { }
    }
}

/* ウィンドウを即時再描画（同期処理中でも表示を更新）/ Force an immediate repaint during sync work */
function forceWindowUpdate(win) {
    try { win.update(); } catch (e) { }
}

/* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
function trimButtonHeight(button, px) {
    try {
        button.size = [button.size.width, button.size.height - px];
    } catch (e) { }
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

/* 選択の統計から署名を作る（変化検出用）/ Build a signature from selection stats (to detect changes) */
function selectionStatsSignature(info) {
    return info.itemCount + "," + (info.isText ? "1" : "0") + "," + info.paragraphCount + "," + info.textObjectCount;
}

/* 定期ポーリング：選択が変わっていたら配色単位のディムだけ更新（読み取り専用・色は変えない）
   Periodic poll: refresh only the unit dims when the selection changed (read-only, colors unchanged) */
function pollDimsUpdate() {
    if (!STATE.win || STATE.busy || !STATE.unitRadioButtons) return;
    var info = readSelectionStats();
    if (!info.ok) return;
    var sig = selectionStatsSignature(info);
    if (sig === STATE.lastStatsSig) return;
    STATE.lastStatsSig = sig;
    updateUnitAvailability(STATE.unitRadioButtons, info);
}

/* ディム追従ポーリングの間隔（ms）。短くすると Illustrator 本体への割り込みが増えて重くなる
   Dim-polling interval (ms); shorter values interrupt Illustrator more often and slow it down */
var DIM_POLL_INTERVAL = 1200;

/* ポーリングの有効／無効。true にすると選択変更に常時追従するが、1.2秒ごとに同期 BridgeTalk で
   メインエンジンへ割り込むため作業中の操作が重くなる。false ではパレットが前面に戻ったとき
   （win.onActivate）と各ボタン操作でディムを更新する
   Enable dim polling. When true it follows selection changes continuously but interrupts the main
   engine via a synchronous BridgeTalk every 1.2s, which slows editing. When false, dims refresh on
   win.onActivate (palette returns to front) and on button actions instead */
var ENABLE_DIM_POLLING = false;

/* ディム追従ポーリングを開始（無効・scheduleTask 非対応環境では何もしない）/ Start dim polling (no-op when disabled or unavailable) */
function startDimPolling() {
    if (!ENABLE_DIM_POLLING) return;
    if (typeof app.scheduleTask !== "function") return;
    STATE.pollTaskId = app.scheduleTask("if(typeof pollDimsUpdate==='function')pollDimsUpdate();", DIM_POLL_INTERVAL, true);
}

/* ディム追従ポーリングを停止 / Stop dim polling */
function stopDimPolling() {
    if (STATE.pollTaskId && typeof app.cancelTask === "function") {
        try { app.cancelTask(STATE.pollTaskId); } catch (e) { }
    }
    STATE.pollTaskId = null;
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
// メインエンジンへの委譲 / Delegation to the main engine
// =========================================
/* worker 関数（メインエンジンで eval される DOM 処理）。追加時は必ずここへ登録
   Worker functions (DOM code eval'd in the main engine). Register every new one here. */
var WORKER_FUNCS = [
    workerApply, workerFinalize, workerCommitPreview, workerOpenSwatchesPanel,
    restoreBaselineW, paintBaselineW, restoreOneSnapshotW, ensureBaselineW, snapshotTargetW, snapshotCharactersW, colorKeyW, selectionKeyW,
    collectTargetsCachedW, prepareApplyW, applyColorsToTargetsW, resetTextStrokeOpacityW, applyNoStrokeFullOpacityW, applyColorFullToCharW,
    workerReadSelection, workerReadObjectColors, workerReadStats, selectionTextStatsW, countParagraphsW,
    flattenSelectionW, collectColorableItemsW, getSingleSelectedTextRangeW,
    collectColorTargetsByUnitW, pushTextRangeTargetsW, pushStaggeredWordTargetsW, pushParagraphTargetsW, getTextUnitRangesW, spanRangeW, applyColorToTargetW,
    sortByPositionW, comparePositionKeysW, shuffleArrayW, randIntW,
    resolveAppliedColorsW, findSwatchColorByNameW, collectFillColorsW, isApplicableFillColorW, serializeColorW, deserializeColorW,
    orderColorsW, fullRandomColorIndexW, pickSwatchColorW,
    buildCMYKColorW, buildRGBColorW, buildGrayColorW, getDefaultRGBColorsW,
    pickTwoChannelCMYW, generateRandomCMYPaletteUniqueW, isFarEnoughCMYW, cmyDistanceW,
    isWhiteColorW, allWhiteSwatchesW, colorToRGB255W
];

/* worker 関数定義部の encode 済みソース（toString/encode は一度だけ）
   Encoded worker-definitions source, computed once */
var gEncodedWorkerDefs = null;
function getEncodedWorkerDefs() {
    if (gEncodedWorkerDefs === null) {
        var source = "var TMK_CMYK_FALLBACK_MAX_TOTAL=" + TMK_CMYK_FALLBACK_MAX_TOTAL + ";var TMK_CMYK_FALLBACK_MIN_DISTANCE=" + TMK_CMYK_FALLBACK_MIN_DISTANCE + ";var PREVIEW_CHAR_CAP=" + PREVIEW_CHAR_CAP + ";";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            source += WORKER_FUNCS[i].toString() + ";";
        }
        gEncodedWorkerDefs = encodeURIComponent(source);
    }
    return gEncodedWorkerDefs;
}

/* ボディをメインエンジンで同期実行して結果を返す（低レベル）/ Run a raw body synchronously in the main engine */
function sendWorkerBody(body) {
    var holder = { result: "ERR:no-result" };
    try {
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = body;
        bridge.onResult = function (message) { holder.result = String(message.body); };
        bridge.onError = function (message) { holder.result = "ERR:" + message.body; };
        bridge.send(10);
    } catch (e) {
        holder.result = "ERR:" + e.message;
    }
    return holder.result;
}

/* メインエンジンが定義を永続保持できるか一度だけ判定 / Detect once whether defs persist across messages */
var gWorkerPersists = null;
function detectWorkerPersistence() {
    if (gWorkerPersists !== null) return;
    sendWorkerBody("eval(decodeURIComponent(\"" + encodeURIComponent("function __aiwProbe(){return 7;}") + "\"));'OK';");
    gWorkerPersists = (sendWorkerBody("(typeof __aiwProbe==='function')?String(__aiwProbe()):'NO';") === "7");
}

/* persist モード時に関数定義をメインエンジンへロード（コード更新時も上書き）/ Load defs into the main engine */
var gWorkerLoaded = false;
function loadWorkerDefs() {
    sendWorkerBody("eval(decodeURIComponent(\"" + getEncodedWorkerDefs() + "\"));'OK';");
    gWorkerLoaded = true;
}

/* メインエンジンに worker 関数が定義済みか（リトライ要否の判定）/ Whether worker funcs are defined (to decide retry) */
function workerFuncsPresent() {
    return sendWorkerBody("(typeof workerApply==='function')?'Y':'N';") === "Y";
}

/* 呼び出し式をメインエンジンで同期実行して結果マーカーを返す
   persist モード：関数は一度ロードし、以後は呼び出し式だけ送る（軽い）
   非対応環境：従来どおり毎回フル送信にフォールバック
   Run the call in the main engine; in persist mode load funcs once and send only tiny calls afterward */
function runWorker(callExpression) {
    if (STATE.busy) return "ERR:busy";
    STATE.busy = true;
    var result;
    try {
        detectWorkerPersistence();
        if (gWorkerPersists) {
            if (!gWorkerLoaded) loadWorkerDefs();
            result = sendWorkerBody(callExpression + ";");
            /* エラー時は「関数が実際に未定義のときだけ」再ロード＋リトライ。
               実行途中で失敗したものをリトライすると二重適用＋スナップショット基準が汚れるため避ける
               （未定義エラーは workerApply が実行前に落ちているので二重適用にならない）
               On error retry only if the funcs are genuinely undefined; retrying a mid-execution failure would double-apply and corrupt the snapshot baseline */
            if (result.indexOf("ERR:") === 0 && !workerFuncsPresent()) {
                loadWorkerDefs();
                result = sendWorkerBody(callExpression + ";");
            }
        } else {
            result = sendWorkerBody("eval(decodeURIComponent(\"" + getEncodedWorkerDefs() + encodeURIComponent(callExpression + ";") + "\"));");
        }
    } finally {
        STATE.busy = false;
    }
    return result;
}

/* 選択情報を委譲取得してパースする / Delegate and parse selection info */
function readSelectionInfo() {
    var info = { ok: false, defaultUnit: "object", chips: [], swatchNames: [], itemCount: 0, isText: false, paragraphCount: 0, textObjectCount: 0 };
    var raw = runWorker("workerReadSelection()");
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

/* 配色単位のディム判定に必要な統計だけを委譲取得（ポーリング用の軽量版）
   Delegate reading only the stats needed for the unit dims (lightweight poll path) */
function readSelectionStats() {
    var info = { ok: false, itemCount: 0, isText: false, paragraphCount: 0, textObjectCount: 0 };
    var raw = runWorker("workerReadStats()");
    if (!raw || raw.indexOf("OK|") !== 0) return info;
    info.ok = true;
    var parts = raw.split("|");
    info.itemCount = parts[1] ? parseInt(parts[1], 10) : 0;
    info.isText = (parts[2] === "1");
    info.paragraphCount = parts[3] ? parseInt(parts[3], 10) : 0;
    info.textObjectCount = parts[4] ? parseInt(parts[4], 10) : 0;
    return info;
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
        var rgb = chipStrings[i].split(",");
        chips.push([parseInt(rgb[0], 10), parseInt(rgb[1], 10), parseInt(rgb[2], 10)]);
    }
    return chips;
}

/* 選択オブジェクトの塗り色を委譲取得（チップRGB＋シリアライズ済みカラー値）
   Delegate reading the fill colors of selected objects (chip RGB + serialized color values) */
function readObjectColorsInfo() {
    var info = { chips: [], colorValues: [], itemCount: 0, isText: false, paragraphCount: 0 };
    var raw = runWorker("workerReadObjectColors()");
    if (!raw || raw.indexOf("OK|") !== 0) return info;

    var parts = raw.split("|");
    info.chips = parseChips(parts[2]);
    /* カラー値は ";" 区切り（各値は "C,c,m,y,k" 等で "," を含むため）/ Values are ";"-separated (each contains ",") */
    if (parts[3]) {
        info.colorValues = parts[3].split(";");
    }
    parseSelectionStats(info, parts);
    return info;
}

/* 文字列配列を worker 呼び出し用の配列リテラルに（encode済み名/シリアライズ値のどちらも安全）
   Serialize a string array as an array literal (safe for encoded names or serialized values) */
function stringsToLiteral(items) {
    if (!items || items.length === 0) return "[]";
    var quoted = [];
    for (var i = 0; i < items.length; i++) {
        quoted.push("\"" + items[i] + "\"");
    }
    return "[" + quoted.join(",") + "]";
}

/* オプションを worker へ渡すオブジェクトリテラル文字列に / Serialize options as an object literal for the worker */
function optionsToLiteral(options) {
    return "{unit:\"" + options.unit + "\",order:\"" + options.order + "\"}";
}

// =========================================
// worker 関数（メインエンジン・DOM）/ Worker functions (main engine, DOM)
// -----------------------------------------
// ※ toString() は改行を消すため、// 行コメント禁止・/* */ のみ・各文はセミコロンで終える
// -----------------------------------------

/* カラーを適用（swatchNames=スウォッチ名 / colorValues=シリアライズ済みカラー値、どちらか一方 / finalApply=確定時）。
   元の塗り/線/不透明度は開いた最初の1回だけ全選択ぶんスナップショット（ベースライン）し、以降のプレビューは上書きのみ。
   どの配色単位でも可視グリフ・全オブジェクトを必ず塗り直すので見た目は常に正しい。取り消しは Cmd+Z（app.undo は使わない）
   Apply colors. Snapshot the whole selection once (baseline); later previews only overwrite. */
function workerApply(options, swatchNames, colorValues, finalApply) {
    if (app.documents.length === 0) return "NODOC";
    var doc = app.activeDocument;
    var selection = doc.selection;
    var items = flattenSelectionW(selection);
    var textRange = getSingleSelectedTextRangeW(selection);
    /* オブジェクト選択は適用後にフォーカスを戻すため控える（テキスト編集中は戻さない）*/
    var savedSelection = textRange ? null : selection;
    /* パレットは非モーダルで選択が変わりうる。選択が変わったら前のプレビューを元へ戻し、全キャッシュを破棄して作り直す
       The palette is non-modal: on a selection change, revert the previous preview and drop all caches */
    var selKey = selectionKeyW(items, textRange);
    if ($.global.__aiApplySelKey !== selKey) {
        paintBaselineW();
        $.global.__aiApplyBaseline = null;
        $.global.__aiApplyTargets = null;
        $.global.__aiApplyTargetsUnit = null;
        $.global.__aiApplyColors = null;
        $.global.__aiApplyGroupIndex = null;
        $.global.__aiApplyWasDecimated = null;
        $.global.__aiApplySelKey = selKey;
    }
    /* 同一単位のプレビューでは対象集合をキャッシュ再利用（順序変更だけなら作り直さない）*/
    var targets = collectTargetsCachedW(items, textRange, options.unit);
    if (targets.length === 0) { restoreBaselineW(); app.redraw(); return "NOSEL"; }
    /* 元状態は初回だけ取得（重い文字読み取りは1回きり）*/
    ensureBaselineW(items, textRange);
    /* 着色前の準備（間引きの要否と着色件数）*/
    var prep = prepareApplyW(items, textRange, options.unit, targets.length, finalApply);
    if (!finalApply) { $.global.__aiApplyWasDecimated = prep.decimated; }
    /* ランダム系（並びシャッフル・完全ランダム抽選・自動CMYK生成）はプレビューと確定で変わらないよう、
       プレビュー時に生成した配色と抽選結果を保存し、確定（finalApply）時はそれを再利用する */
    var colors, groupIndexCache;
    if (finalApply && $.global.__aiApplyColors) {
        colors = $.global.__aiApplyColors;
        groupIndexCache = $.global.__aiApplyGroupIndex || {};
    } else {
        colors = orderColorsW(resolveAppliedColorsW(doc, targets.length, swatchNames, colorValues), options.order);
        groupIndexCache = {};
        $.global.__aiApplyColors = colors;
        $.global.__aiApplyGroupIndex = groupIndexCache;
    }
    applyColorsToTargetsW(targets, prep.limit, colors, options.order === "fullrandom", prep.decimated, groupIndexCache);
    /* 適用で変わった選択（フォーカス）を元に戻す */
    if (savedSelection) { try { app.selection = savedSelection; } catch (e) {} }
    app.redraw();
    return "OK";
}

/* 閉じる時の確定処理。間引きプレビューだったときだけ保存配色で全文字をフル着色し、その後ベースラインを破棄する
   Finalize on close: if the preview was decimated, apply fully using the saved colors, then discard the baseline */
function workerFinalize(options, swatchNames, colorValues) {
    if ($.global.__aiApplyWasDecimated) { workerApply(options, swatchNames, colorValues, true); }
    return workerCommitPreview();
}

/* 着色前の準備。1文字単位で対象が多いプレビューは間引き（先頭 PREVIEW_CHAR_CAP だけ着色し残りは元色）、
   それ以外は線なし・不透明度100 を範囲へ一括設定。戻り値 { decimated, limit } */
function prepareApplyW(items, textRange, unitMode, targetCount, finalApply) {
    var decimated = (!finalApply && unitMode === "character" && targetCount > PREVIEW_CHAR_CAP);
    if (decimated) { paintBaselineW(); return { decimated: true, limit: PREVIEW_CHAR_CAP }; }
    resetTextStrokeOpacityW(items, textRange);
    return { decimated: false, limit: targetCount };
}

/* 対象 targets の先頭 limit 件へ色を適用。完全ランダムはグループごとに1回抽選（groupIndexCache に保持）、
   間引き時は文字ごとに線・不透明度も設定 */
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

/* 対象集合を単位ごとにキャッシュ。単位が同じなら作り直さない（順序変更だけのプレビューを速くする）*/
function collectTargetsCachedW(items, textRange, unitMode) {
    if ($.global.__aiApplyTargetsUnit === unitMode && $.global.__aiApplyTargets) { return $.global.__aiApplyTargets; }
    var targets = (items.length > 0 || textRange) ? collectColorTargetsByUnitW(items, textRange, unitMode) : [];
    $.global.__aiApplyTargets = targets;
    $.global.__aiApplyTargetsUnit = unitMode;
    return targets;
}

/* 元状態のベースラインを初回だけ取得（全選択ぶん・テキストは文字単位で忠実に保存）*/
function ensureBaselineW(items, textRange) {
    if ($.global.__aiApplyBaseline) { return; }
    var baseline = [];
    if (textRange) { snapshotTargetW({ kind: "textrange", node: textRange }, baseline); }
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it.typename === "TextFrame") { snapshotTargetW({ kind: "textframe", node: it }, baseline); }
        else if (it.typename === "PathItem") { snapshotTargetW({ kind: "path", node: it }, baseline); }
        else if (it.typename === "CompoundPathItem") { snapshotTargetW({ kind: "compound", node: it }, baseline); }
    }
    $.global.__aiApplyBaseline = baseline;
}

/* 選択の同一性キー（選択の形だけ。塗り替えで変化しない値で作る）。前回と変われば選択が変わった＝作り直す
   Identity key for the selection shape (values coloring never changes); a change means re-capture */
function selectionKeyW(items, textRange) {
    var parts = [items.length];
    parts.push(textRange ? ("tr:" + textRange.start + ":" + textRange.end) : "tr:-");
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var key = item.typename;
        try {
            var b = item.geometricBounds;
            key += ":" + Math.round(b[0] * 100) + ":" + Math.round(b[1] * 100) + ":" + Math.round(b[2] * 100) + ":" + Math.round(b[3] * 100);
        } catch (e) { key += ":?"; }
        if (item.typename === "TextFrame") {
            try { key += ":" + item.textRange.end; } catch (e2) { key += ":?"; }
        }
        parts.push(key);
    }
    return parts.join("|");
}

/* 現在のプレビューを確定（各キャッシュを破棄＝以後 restore しない）。閉じる時・初回開始前に呼ぶ */
function workerCommitPreview() {
    $.global.__aiApplyBaseline = null;
    $.global.__aiApplyTargets = null;
    $.global.__aiApplyTargetsUnit = null;
    $.global.__aiApplyColors = null;
    $.global.__aiApplyGroupIndex = null;
    $.global.__aiApplyWasDecimated = null;
    $.global.__aiApplySelKey = null;
    return "OK";
}

/* 「スウォッチ」パネルを表示（メニューコマンドはメインエンジンでのみ実行可）
   Show the Swatches panel (menu commands only run in the main engine) */
function workerOpenSwatchesPanel() {
    app.executeMenuCommand('Adobe Swatches Menu Item');
    return "OK";
}

/* ベースライン（元の塗り/線/不透明度）を画面へ書き戻すが破棄はしない。間引きプレビューで着色しなかった残りを元色に戻す用
   Repaint the baseline (originals) without discarding it; used by the decimated preview to reset the tail */
function paintBaselineW() {
    var baseline = $.global.__aiApplyBaseline;
    if (!baseline) { return; }
    for (var i = 0; i < baseline.length; i++) {
        try { restoreOneSnapshotW(baseline[i]); } catch (e) { }
    }
}

/* ベースラインを書き戻して破棄（対象なし時など。以後 restore しない）
   Repaint the baseline and discard it (e.g. when there are no targets; never restored afterward) */
function restoreBaselineW() {
    paintBaselineW();
    $.global.__aiApplyBaseline = null;
}

/* スナップショット1件を復元。テキストは同色ランを範囲1回で書き戻す（1文字ずつより桁違いに速い）
   Restore one snapshot entry; text runs are written back in a single span write */
function restoreOneSnapshotW(entry) {
    var node = entry.node;
    if (entry.kind === "textrun") {
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
function snapshotTargetW(target, out) {
    var node = target.node;
    if (target.kind === "span") { snapshotCharactersW(spanRangeW(target), out); }
    else if (target.kind === "textrange") { snapshotCharactersW(node, out); }
    else if (target.kind === "textframe") { snapshotCharactersW(node.textRange, out); }
    else if (target.kind === "path") { out.push({ kind: "path", node: node, fill: node.fillColor, stroked: node.stroked, opacity: node.opacity }); }
    else if (target.kind === "compound") {
        var subs = node.pathItems;
        var arr = [];
        for (var i = 0; i < subs.length; i++) { arr.push({ sub: subs[i], fill: subs[i].fillColor, stroked: subs[i].stroked, opacity: subs[i].opacity }); }
        out.push({ kind: "compound", node: node, subs: arr });
    }
}

/* テキスト範囲を「同色ラン」単位でスナップショット（元が混色でも忠実に戻せる／文字ごと保存より桁違いに軽い）。
   連続範囲なので i 文字目の story オフセットは range.start + i（ch.start の DOM 読み取りを省く）
   Snapshot a text range as same-color runs (faithful for mixed originals, far lighter than per-character) */
function snapshotCharactersW(range, out) {
    var story = range.story;
    var chars = range.characters;
    var count = chars.length;
    var rangeStart = range.start;
    var started = false;
    var runFill = null, runStroke = null, runOpacity = 0, runKey = null, runStart = 0, runEnd = 0;
    for (var i = 0; i < count; i++) {
        var ch = chars[i];
        var fill = ch.fillColor;
        var stroke = ch.strokeColor;
        var opacity = ch.opacity;
        var chStart = rangeStart + i;
        var key = colorKeyW(fill) + "|" + colorKeyW(stroke) + "|" + opacity;
        if (started && key === runKey) {
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

/* 色の同一性キー（NoColor 対応・SpotColor は tint も含める）。run 判定に使う
   Identity key for a color (NoColor-aware; SpotColor includes tint) used for run detection */
function colorKeyW(color) {
    if (color.typename === "NoColor") { return "N"; }
    if (color.typename === "SpotColor") {
        var tint = (typeof color.tint === "number") ? color.tint : 100;
        return "S," + tint + "," + serializeColorW(color);
    }
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

/* 配色単位のディム判定に必要な統計だけを返す（ポーリング用の軽量版：スウォッチ取得・チップ生成をしない）
   "OK|<itemCount>|<isText>|<paragraphCount>|<textObjectCount>"
   Return only the stats needed for unit dims (lightweight poll path: no swatch read, no chips) */
function workerReadStats() {
    if (app.documents.length === 0) return "NODOC";
    var items = flattenSelectionW(app.selection);
    var stats = selectionTextStatsW(items, getSingleSelectedTextRangeW(app.selection));
    return "OK|" + items.length + "|" + (stats.isText ? "1" : "0") + "|" + stats.paragraphCount + "|" + stats.textObjectCount;
}

/* 選択オブジェクトの塗り色を返す "OK|<defaultUnit>|<r,g,b;...>|<serialized;...>|<itemCount>|<isText>|<paragraphCount>|<textObjectCount>" / Return fill colors of selected objects */
function workerReadObjectColors() {
    if (app.documents.length === 0) return "NODOC";
    var items = flattenSelectionW(app.selection);
    var textRange = getSingleSelectedTextRangeW(app.selection);
    var defaultUnit = "object";
    if (textRange) { defaultUnit = "character"; }
    else if (items.length === 1 && items[0].typename === "TextFrame") { defaultUnit = "character"; }
    var fillColors = collectFillColorsW(items, textRange);
    var chipParts = [];
    var valueParts = [];
    for (var i = 0; i < fillColors.length; i++) {
        chipParts.push(colorToRGB255W(fillColors[i]));
        valueParts.push(serializeColorW(fillColors[i]));
    }
    var stats = selectionTextStatsW(items, textRange);
    return "OK|" + defaultUnit + "|" + chipParts.join(";") + "|" + valueParts.join(";") + "|" + items.length + "|" + (stats.isText ? "1" : "0") + "|" + stats.paragraphCount + "|" + stats.textObjectCount;
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

/* テキストの線・不透明度をまとめて初期化（線なし・不透明度100）。全テキスト対象で同じ値なので範囲へ1回だけ書く
   Reset stroke/opacity for text in bulk; the value is the same for every text target, so write it once per range */
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

/* 間引きプレビュー用：1文字に塗り・線なし・不透明度100 をまとめて設定（着色しない残りは元色のまま）
   For decimated preview: set fill + no-stroke + 100% opacity on a single character (the rest stays original) */
function applyColorFullToCharW(target, color) {
    var node = target.node;
    node.fillColor = color;
    applyNoStrokeFullOpacityW(node);
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

/* 使用するカラーを決定（取り込んだスウォッチ名 or カラー値を最優先、無ければ自動カラー）
   Resolve colors (captured swatch names or serialized values take priority; otherwise auto colors) */
function resolveAppliedColorsW(doc, targetCount, swatchNames, colorValues) {
    var colors = [];
    if (swatchNames && swatchNames.length > 0) {
        for (var i = 0; i < swatchNames.length; i++) {
            var swatchColor = findSwatchColorByNameW(doc, decodeURIComponent(swatchNames[i]));
            if (swatchColor) { colors.push(swatchColor); }
        }
    } else if (colorValues && colorValues.length > 0) {
        for (var j = 0; j < colorValues.length; j++) {
            colors.push(deserializeColorW(colorValues[j]));
        }
    }
    if (colors.length >= 1) { return colors; }
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) { return generateRandomCMYPaletteUniqueW(targetCount, TMK_CMYK_FALLBACK_MAX_TOTAL); }
    return getDefaultRGBColorsW();
}

/* スウォッチ名から色を取得（選択状態に依存しない）/ Find a swatch color by name (independent of selection state) */
function findSwatchColorByNameW(doc, name) {
    var swatches = doc.swatches;
    for (var i = 0; i < swatches.length; i++) {
        if (swatches[i].name === name) { return swatches[i].color; }
    }
    return null;
}

/* 選択オブジェクトの塗り色を重複なく収集 / Collect distinct fill colors from selected objects */
function collectFillColorsW(items, textRange) {
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
        if (!isApplicableFillColorW(c)) { continue; }
        var key = serializeColorW(c);
        if (seen[key]) { continue; }
        seen[key] = true;
        colors.push(c);
    }
    return colors;
}

/* 単色として取り込める塗り色か（NoColor・グラデ・パターンは除外）
   Whether a fill is a solid color usable as a swatch (skip NoColor / gradient / pattern) */
function isApplicableFillColorW(color) {
    if (!color) { return false; }
    var t = color.typename;
    return (t === "CMYKColor" || t === "RGBColor" || t === "GrayColor" || t === "SpotColor");
}

/* カラーを文字列にシリアライズ / Serialize a color to a string */
function serializeColorW(color) {
    var t = color.typename;
    if (t === "CMYKColor") { return "C," + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black; }
    if (t === "RGBColor") { return "R," + color.red + "," + color.green + "," + color.blue; }
    if (t === "GrayColor") { return "G," + color.gray; }
    if (t === "SpotColor") { return serializeColorW(color.spot.color); }
    return "R,128,128,128";
}

/* 文字列からカラーを復元 / Reconstruct a color from a string */
function deserializeColorW(str) {
    var p = str.split(",");
    if (p[0] === "C") { return buildCMYKColorW(Number(p[1]), Number(p[2]), Number(p[3]), Number(p[4])); }
    if (p[0] === "R") { return buildRGBColorW(Number(p[1]), Number(p[2]), Number(p[3])); }
    if (p[0] === "G") { return buildGrayColorW(Number(p[1])); }
    return buildRGBColorW(128, 128, 128);
}

/* Gray カラーを生成 / Build a GrayColor */
function buildGrayColorW(gray) {
    var color = new GrayColor();
    color.gray = gray;
    return color;
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

showPalette();
