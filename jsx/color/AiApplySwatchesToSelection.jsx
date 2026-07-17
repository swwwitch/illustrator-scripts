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
var SCRIPT_VERSION = "v1.7.3";

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
        fromObjects:  { ja: "塗り色を読込", en: "Load Fills" }
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
        orderFullRandom: { ja: "適用先ごとにカラーを毎回ランダムに選びます（並びは繰り返しません）。", en: "Picks a color at random for each target (no repeating sequence)." }
    },
    status: {
        done:  { ja: "プレビューを更新しました。", en: "Preview updated." },
        noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSel: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
        busy:  { ja: "処理中です…", en: "Working…" },
        error: { ja: "エラー: ", en: "Error: " }
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
/* 再実行でも消えないよう $.global に保持（GC回避・多重起動防止）。プレビューの取り消し情報は
   メインエンジン側の $.global.__aiApplyPreviewSnapshot に持つ（app.undo は使わない）
   Kept on $.global so re-runs don't reset it. Preview-restore data lives on the main engine's snapshot */
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

    /* 状況表示 / Status line */
    var statusText = win.add("statictext", undefined, "");
    statusText.alignment = "left";

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
        statusText.text = markerToMessage(runWorker(call));
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

    /* 適用ボタンなし：閉じたら現在の適用結果を確定（スナップショットは破棄＝以後 restore しない。戻すのは Cmd+Z）
       No Apply button: closing keeps the current result and discards the snapshot (undo via Cmd+Z) */
    win.onClose = function () {
        stopDimPolling();
        runWorker("workerCommitPreview()");
        STATE.win = null;
        STATE.unitRadioButtons = null;
        return true;
    };

    setProgress(progress, 75, L("progress.applying"));
    win.show();
    /* レイアウト確定後に取り込みボタンの天地を詰める / Trim the capture buttons' height after layout */
    trimButtonHeight(fromSwatchesButton, 2);
    trimButtonHeight(fromObjectsButton, 2);
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

/* ディム追従ポーリングを開始（scheduleTask 非対応環境では何もしない）/ Start dim polling (no-op if scheduleTask is unavailable) */
function startDimPolling() {
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
    workerApply, workerCommitPreview, restorePreviewW, restoreOneSnapshotW, snapshotTargetW, snapshotCharactersW,
    workerReadSelection, workerReadObjectColors, workerReadStats, selectionTextStatsW,
    flattenSelectionW, collectColorableItemsW, getSingleSelectedTextRangeW,
    collectColorTargetsByUnitW, pushTextRangeTargetsW, pushStaggeredWordTargetsW, getTextUnitRangesW, applyColorToTargetW,
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
        var source = "var TMK_CMYK_FALLBACK_MAX_TOTAL=" + TMK_CMYK_FALLBACK_MAX_TOTAL + ";var TMK_CMYK_FALLBACK_MIN_DISTANCE=" + TMK_CMYK_FALLBACK_MIN_DISTANCE + ";";
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

/* 結果マーカーをローカライズした状況文へ / Map a result marker to a localized status message */
function markerToMessage(marker) {
    if (marker === "OK") return L("status.done");
    if (marker === "NODOC") return L("status.noDoc");
    if (marker === "NOSEL") return L("status.noSel");
    if (marker === "ERR:busy") return L("status.busy");
    if (marker.indexOf("ERR:") === 0) return L("status.error") + marker.substring(4);
    return String(marker);
}

// =========================================
// worker 関数（メインエンジン・DOM）/ Worker functions (main engine, DOM)
// -----------------------------------------
// ※ toString() は改行を消すため、// 行コメント禁止・/* */ のみ・各文はセミコロンで終える
// -----------------------------------------

/* カラーを適用（swatchNames=スウォッチ名 / colorValues=シリアライズ済みカラー値、どちらか一方）。
   app.undo は使わず、前回プレビューで書き換えた塗り/線/不透明度をスナップショットから復元してから再適用する。
   これで手動編集・選択変更・部分失敗に強くなる（app.undo はグローバル履歴を巻き戻すため危険）
   Apply colors; revert the previous preview via snapshot restore (never app.undo) */
function workerApply(options, swatchNames, colorValues) {
    if (app.documents.length === 0) return "NODOC";
    var doc = app.activeDocument;
    /* 前回プレビューを元の状態へ復元（スナップショットは restore 内で破棄）
       Restore the previous preview to its originals (snapshot is cleared inside restore) */
    restorePreviewW();
    /* 復元後の現在の選択で対象を集める（参照はすべて有効）/ Collect targets from the current selection (all refs valid) */
    var selection = doc.selection;
    var items = flattenSelectionW(selection);
    var textRange = getSingleSelectedTextRangeW(selection);
    /* オブジェクト選択は復元用に控える（テキスト編集中は復元しない）
       Save the object selection for restore (skip while editing text) */
    var savedSelection = textRange ? null : selection;
    var targets = (items.length > 0 || textRange) ? collectColorTargetsByUnitW(items, textRange, options.unit) : [];
    if (targets.length === 0) { app.redraw(); return "NOSEL"; }
    var colors = orderColorsW(resolveAppliedColorsW(doc, targets.length, swatchNames, colorValues), options.order);
    /* 各対象の元の状態を保存してから着色（次回プレビューで元に戻せるように）
       Snapshot each target's originals before coloring, so the next preview can revert it */
    var snapshot = [];
    /* 完全ランダム：グループ（単語など）ごとに一度だけ抽選し、同一グループ内は同じ色にする
       Fully random: draw once per group (e.g. a word) so a group keeps one color */
    var isFullRandom = (options.order === "fullrandom");
    var randomIndexByGroup = {};
    for (var i = 0; i < targets.length; i++) {
        snapshotTargetW(targets[i], snapshot);
        var colorIndex;
        if (isFullRandom) {
            var groupKey = (typeof targets[i].groupKey === "string") ? targets[i].groupKey : String(i);
            colorIndex = fullRandomColorIndexW(groupKey, randomIndexByGroup, colors.length);
        } else {
            colorIndex = (typeof targets[i].colorIndex === "number") ? targets[i].colorIndex : i;
        }
        applyColorToTargetW(targets[i], pickSwatchColorW(colorIndex, colors));
    }
    $.global.__aiApplyPreviewSnapshot = snapshot;
    /* 適用で変わった選択（フォーカス）を元に戻す / Restore the selection changed by applying */
    if (savedSelection) { try { app.selection = savedSelection; } catch (e) {} }
    app.redraw();
    return "OK";
}

/* 現在のプレビューを確定（スナップショットを破棄＝以後 restore しない）。閉じる時・初回開始時に呼ぶ
   Commit the current preview: discard the snapshot so it is never restored (called on close / before first preview) */
function workerCommitPreview() {
    $.global.__aiApplyPreviewSnapshot = null;
    return "OK";
}

/* 前回プレビューで書き換えた塗り/線/不透明度を元へ戻す（オブジェクト削除時などは個別に無視）
   Restore fills/strokes/opacity changed by the previous preview (skip individually if an object was deleted) */
function restorePreviewW() {
    var snapshot = $.global.__aiApplyPreviewSnapshot;
    $.global.__aiApplyPreviewSnapshot = null;
    if (!snapshot) { return; }
    for (var i = 0; i < snapshot.length; i++) {
        try { restoreOneSnapshotW(snapshot[i]); } catch (e) { }
    }
}

/* スナップショット1件を復元 / Restore one snapshot entry */
function restoreOneSnapshotW(entry) {
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

/* 着色前の対象の元の状態を snapshot 配列へ保存（テキストは文字単位で忠実に保存）
   Snapshot a target's originals into the array (text is snapshotted per character for fidelity) */
function snapshotTargetW(target, out) {
    var node = target.node;
    if (target.kind === "textrange") { snapshotCharactersW(node, out); }
    else if (target.kind === "textframe") { snapshotCharactersW(node.textRange, out); }
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
function snapshotCharactersW(range, out) {
    var chars = range.characters;
    for (var i = 0; i < chars.length; i++) {
        var ch = chars[i];
        out.push({ kind: "char", node: ch, fill: ch.fillColor, stroke: ch.strokeColor, opacity: ch.opacity });
    }
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
    var paragraphCount = (textObjectCount === 1) ? singleTextRange.paragraphs.length : 0;
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
    var ranges = getTextUnitRangesW(textRange, unitMode);
    if (!ranges) { out.push({ kind: "textrange", node: textRange }); return; }
    for (var i = 0; i < ranges.length; i++) { out.push({ kind: "textrange", node: ranges[i] }); }
}

/* 単語ごと：Illustrator の単語境界で文字を単語にマッピングし、句読点・スペースは直前の単語色にする
   （略語・小数は1単語のまま。各行先頭は互い違い）
   Per word: map characters to Illustrator's word boundaries; punctuation/spaces share the preceding word's color */
function pushStaggeredWordTargetsW(textRange, out) {
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
            var safeWordIndex = (wordIndex < 0) ? 0 : wordIndex;
            /* groupKey は行内で一意（完全ランダムで単語ごとに別色を抽選するため）
               groupKey is unique per line+word so fully random draws a distinct color per word */
            out.push({ kind: "textrange", node: chars[ci], colorIndex: safeWordIndex + li, groupKey: li + ":" + safeWordIndex });
        }
    }
}

/* 単位に対応するテキスト範囲コレクション（word は pushTextRangeTargetsW で先に処理）
   Text range collection for a unit (word is handled earlier in pushTextRangeTargetsW) */
function getTextUnitRangesW(textRange, unitMode) {
    if (unitMode === "character") return textRange.characters;
    if (unitMode === "line") return textRange.lines;
    if (unitMode === "paragraph") return textRange.paragraphs;
    return null;
}

/* ターゲットにカラーを設定 / Set the color on a target */
function applyColorToTargetW(target, color) {
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

/* インデックスに応じて色を取得（ループ・不透明度100%）/ Pick a color by index (wraps, opacity 100%) */
function pickSwatchColorW(index, colors) {
    var color = colors[index % colors.length];
    color.opacity = 100;
    return color;
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
