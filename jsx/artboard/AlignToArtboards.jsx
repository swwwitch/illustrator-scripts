#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
【概要 / Summary】
複数アートボードのドキュメントで、選択中のオブジェクトを各アートボード上の指定位置へ整列します。
ダイアログ表示中はライブプレビューで結果を確認でき、［OK］で確定、［キャンセル］または閉じる操作で元の位置へ戻します。

### 整列の基準
- すべてのアートボード：選択オブジェクトを中心点が属するアートボードごとに振り分け、各アートボード内の指定位置へ整列します。
- アクティブを基準：アクティブアートボード上の選択を基準に、他のアートボード上の選択を同じ相対位置へ整列します。アクティブ側のオブジェクトは移動しません。

### 整列先
- 3×3 の9点から整列位置を選択できます：左上／上中央／右上／左中央／中央／右中央／左下／下中央／右下
- 「すべてのアートボード」選択時は、選択した9点を各オブジェクトの整列アンカーとして使用します。
- 「アクティブを基準」選択時は、現在選択されている整列先を基準に相対位置を取得します。整列先パネルは無効表示になり、その時点の選択が維持されます。
- 「中央」選択時はマージンは無効（0扱い）です。上中央／下中央では左右マージン、左中央／右中央では上下マージンを使いません。

### マージン（左右／上下）
- 「すべてのアートボード」選択時、対応する辺から内側へオフセットできます。
- 左右・上下を別々に指定できます。「連動」ON のときは、左右の入力値を上下にも反映します（デフォルト：ON）。
- 入力単位は Illustrator のルーラー単位（rulerType）に連動し、内部では pt に変換して処理します。
- 「アクティブを基準」選択時は、マージンは無効（0扱い）です。

### 境界と操作
- 「プレビュー境界を使用」ON：線幅や効果を含む見た目の境界で整列します。OFF：図形本体の幾何境界で整列します。
- 整列先はキーボードでも選択できます：q=左上 / w=上中央 / e=右上 / a=左中央 / s=中央 / d=右中央 / z=左下 / x=下中央 / c=右下
- 整列の基準はキーボードでも選択できます：1=すべてのアートボード / 2=アクティブを基準
- マージン欄は ↑↓ で増減（Shift=±10、Option=±0.1）し、入力中もプレビューが更新されます。
- Enter/Return キーで［OK］を実行できます。

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AlignToArtboards";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-05-15";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在のロケールを判定 / Detect current locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ロケールに応じてラベルを切替 / Switch label by locale */
function _(ja, en) {
    return (lang === "ja") ? ja : en;
}

var LABELS = {
    dialogTitle: { ja: "各アートボードに整列 " + SCRIPT_VERSION, en: "Align to Artboards " + SCRIPT_VERSION },
    targetCorner: { ja: "整列先", en: "Target" },
    target: { ja: "整列の基準", en: "Align based on" },
    targetArtboard: { ja: "すべてのアートボード", en: "All Artboards" },
    targetSelection: { ja: "アクティブを基準", en: "Based on Active Artboard" },
    targetArtboardTip: { ja: "選択オブジェクトを中心点が属するアートボードごとに振り分け、各アートボード内の指定位置へ整列します。", en: "Groups selected objects by the artboard containing their center point, then aligns them to the selected position on each artboard." },
    targetSelectionTip: { ja: "アクティブアートボード上の選択を基準に、他のアートボード上の選択を同じ相対位置へ整列します。アクティブ側は動かしません。", en: "Uses the selection on the active artboard as the reference and aligns selections on other artboards to the same relative position. Objects on the active artboard are not moved." },
    targetCornerTip: { ja: "整列先の9点を選択します。アクティブを基準にする場合は、その時点の選択が維持されます。", en: "Choose one of the 9 alignment positions. When using Based on Active Artboard, the current choice is preserved." },
    marginTip: { ja: "対応する辺から内側へオフセットします。中央またはアクティブを基準にする場合は無効です。", en: "Offsets objects inward from the corresponding edge. Disabled for Center and Based on Active Artboard." },
    linkTip: { ja: "ONのときは左右の値を上下にも連動します。", en: "When enabled, the horizontal value is also used for the vertical margin." },
    usePreviewBoundsTip: { ja: "ON：線幅や効果を含む見た目の境界で整列。OFF：図形本体の幾何境界で整列。", en: "On: align by visual bounds including strokes and effects. Off: align by geometric bounds of the object shape." },
    leftTop: { ja: "左上", en: "Top-Left" },
    topCenter: { ja: "上中央", en: "Top-Center" },
    rightTop: { ja: "右上", en: "Top-Right" },
    leftMiddle: { ja: "左中央", en: "Middle-Left" },
    center: { ja: "中央", en: "Center" },
    rightMiddle: { ja: "右中央", en: "Middle-Right" },
    leftBottom: { ja: "左下", en: "Bottom-Left" },
    bottomCenter: { ja: "下中央", en: "Bottom-Center" },
    rightBottom: { ja: "右下", en: "Bottom-Right" },
    marginValue: { ja: "マージン", en: "Margin" },
    marginX: { ja: "左右", en: "Horizontal" },
    marginY: { ja: "上下", en: "Vertical" },
    link: { ja: "連動", en: "Linked" },
    usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
};

// =========================================
// 境界取得（プレビュー境界 or 幾何境界） / Bounds resolver
// プレビュー境界＝ストローク・効果を含む、幾何境界＝図形のみ / Preview bounds include strokes/effects; geometric bounds use the object shape only
// =========================================
var usePreviewBounds = true;

/* アイテムの整列用バウンディングボックスを返す / Return item bounds for alignment */
function getItemBounds(item) {
    try {
        return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
    } catch (e) {
        return null;
    }
}

// =========================================
// 単位（rulerType） / Units (rulerType)
// =========================================
var unitLabelMap = {
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

/* 現在のルーラー単位ラベルを取得 / Get current ruler unit label */
function getCurrentUnitLabel() {
    var unitCode = 2; /* 初期値はpt / Default unit is pt */
    try { unitCode = app.preferences.getIntegerPreference("rulerType"); } catch (e) { }
    return unitLabelMap[unitCode] || "pt";
}

/* 入力値（現在単位）を pt に変換 / Convert a value typed in current unit to points */
function toPt(valueNumber, unitLabel) {
    if (valueNumber === null || valueNumber === undefined) return null;
    if (isNaN(valueNumber)) return null;

    /* UnitValue互換の単位名へ正規化 / Normalize unit labels for UnitValue compatibility */
    var u = unitLabel;
    if (u === "pica") u = "pc";
    else if (u !== "in" && u !== "mm" && u !== "cm" && u !== "pt" && u !== "px") {
        /* 未対応または特殊な単位はptとして扱う / Treat unsupported or rare units as pt */
        u = "pt";
    }

    try {
        return new UnitValue(valueNumber, u).as("pt");
    } catch (eUV) {
        return valueNumber; /* フォールバック：ptとして扱う / Fallback: assume pt */
    }
}

// =========================================
// メイン / Main
// =========================================

/* エントリーポイント：ダイアログを表示して整列を実行 / Entry point: show dialog and run alignment */
function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) return;

    var input = showDialog();
    if (input === null) return;

    var corner = input.corner;   /* 9点コード / 9-point code: LT TC RT / LM C RM / LB BC RB */
    var targetMode = input.targetMode || "artboard"; /* 整列の基準 / Alignment base: "artboard" | "selection" */
    var margin = input.marginPt; /* pt単位：数値または{x, y} / In pt: number or {x, y} */

    if (targetMode === "selection") {
        alignByActiveArtboardReference(
            doc,
            doc.selection,
            corner,
            function (items, artboardRect, alignmentCorner, offset) {
                alignGroupToArtboardWithOffset(items, artboardRect, alignmentCorner, offset);
            }
        );

        try { app.redraw(); } catch (eRedrawMain1) { }

    } else {
        /* 通常モード：各アイテムの中心点を基準にアートボードへ振り分け / Default mode: group items by artboard using each item center */
        var map = groupSelectedItemsByArtboard(doc, doc.selection);
        if (!map) return;

        for (var abIndex in map) {
            if (!map.hasOwnProperty(abIndex)) continue;

            var items = map[abIndex];
            if (!items || items.length === 0) continue;

            var rect = doc.artboards[parseInt(abIndex, 10)].artboardRect; /* [左, 上, 右, 下] / [L, T, R, B] */
            alignItemsToCornerWithMargin(items, rect, corner, margin);
        }
        try { app.redraw(); } catch (eRedrawMain2) { }
    }
}

main();


// =========================================
// ダイアログ UI / Dialog UI
// =========================================

/* プレビュー状態を初期化 / Create preview state from current selection */
function createPreviewState(doc) {
    var previewState = {
        items: [],
        dx: [],
        dy: [],
        applied: false
    };

    try {
        if (doc && doc.selection && doc.selection.length > 0) {
            for (var i = 0; i < doc.selection.length; i++) {
                previewState.items.push(doc.selection[i]);
                previewState.dx.push(0);
                previewState.dy.push(0);
            }
        }
    } catch (e) { }

    return previewState;
}

/* プレビュー移動を巻き戻す / Revert preview translation */
function revertPreview(previewState, doRedraw) {
    if (doRedraw === undefined) doRedraw = true;
    if (!previewState || !previewState.applied) return;

    for (var i = 0; i < previewState.items.length; i++) {
        var item = previewState.items[i];
        var dx = previewState.dx[i];
        var dy = previewState.dy[i];
        if (!item) continue;

        try {
            if (dx !== 0 || dy !== 0) item.translate(-dx, -dy);
        } catch (e) { }

        previewState.dx[i] = 0;
        previewState.dy[i] = 0;
    }

    previewState.applied = false;

    if (doRedraw) {
        try { app.redraw(); } catch (eRedraw) { }
    }
}

/* プレビュー整列を適用 / Apply preview alignment */
function applyPreview(doc, previewState, corner, marginX, marginY, targetMode) {
    revertPreview(previewState, false);

    if (!doc || !previewState || !previewState.items || previewState.items.length === 0) return;

    if (targetMode === "selection") {
        alignByActiveArtboardReference(
            doc,
            previewState.items,
            corner,
            function (items, artboardRect, alignmentCorner, offset) {
                alignGroupToArtboardWithOffsetPreview(items, artboardRect, alignmentCorner, offset, previewState);
            }
        );
    } else {
        var map = groupSelectedItemsByArtboard(doc, previewState.items);
        if (!map) return;

        for (var abIndex in map) {
            if (!map.hasOwnProperty(abIndex)) continue;

            var items = map[abIndex];
            if (!items || items.length === 0) continue;

            var rect = doc.artboards[parseInt(abIndex, 10)].artboardRect;
            alignItemsToCornerWithMarginPreviewXY(items, rect, corner, marginX, marginY, previewState);
        }
    }

    previewState.applied = true;
    try { app.redraw(); } catch (eRedraw2) { }
}

/* 設定ダイアログを表示し、入力結果を返す / Show settings dialog and return user input */
function showDialog() {
    var w = new Window("dialog", _(LABELS.dialogTitle.ja, LABELS.dialogTitle.en));
    var unitLabel = getCurrentUnitLabel();

    var doc = app.activeDocument;
    var previewState = createPreviewState(doc);
    var closingByOK = false; /* OK経由で閉じるときの再描画を抑止 / Suppress redraw when closing via OK */

    /* UIから整列先コードを取得 / Read the selected corner code from UI */
    function getCornerFromUI() {
        if (rbLT.value) return "LT";
        if (rbTC.value) return "TC";
        if (rbRT.value) return "RT";
        if (rbLM.value) return "LM";
        if (rbC.value) return "C";
        if (rbRM.value) return "RM";
        if (rbLB.value) return "LB";
        if (rbBC.value) return "BC";
        if (rbRB.value) return "RB";
        return "LT";
    }

    /* 整列先ショートカットを登録 / Register corner shortcut keys */
    function addCornerKeyHandler(dialog) {
        dialog.addEventListener("keydown", function (event) {
            /* テキスト入力中はショートカットを無効化 / Skip shortcuts while typing in edittext */
            if (event.target && event.target.type === "edittext") return;
            /* 整列先パネルが無効のときはショートカットも無効 / Skip shortcuts when the corner panel is disabled */
            try { if (!pCorner.enabled) return; } catch (eP) { }

            var k = event.keyName;
            if (!k) return;
            k = ("" + k).toLowerCase();

            var target = null;
            if (k === "q") target = rbLT;
            else if (k === "w") target = rbTC;
            else if (k === "e") target = rbRT;
            else if (k === "a") target = rbLM;
            else if (k === "s") target = rbC;
            else if (k === "d") target = rbRM;
            else if (k === "z") target = rbLB;
            else if (k === "x") target = rbBC;
            else if (k === "c") target = rbRB;

            if (target) {
                selectCornerRb(target);
                event.preventDefault();
                updatePreviewFromUI();
            }
        });
    }

    w.orientation = "column";
    w.alignChildren = ["fill", "top"];

    /* 整列の基準パネル（全幅・縦並び） / Alignment base panel (full-width vertical radios) */
    var pTarget = w.add("panel", undefined, _(LABELS.target.ja, LABELS.target.en));
    pTarget.orientation = "column";
    pTarget.alignChildren = ["left", "center"];
    pTarget.margins = [15, 20, 15, 10];

    var rbTargetArtboard = pTarget.add("radiobutton", undefined, _(LABELS.targetArtboard.ja, LABELS.targetArtboard.en));
    var rbTargetSelection = pTarget.add("radiobutton", undefined, _(LABELS.targetSelection.ja, LABELS.targetSelection.en));
    rbTargetSelection.value = true;
    rbTargetArtboard.helpTip = _(LABELS.targetArtboardTip.ja, LABELS.targetArtboardTip.en);
    rbTargetSelection.helpTip = _(LABELS.targetSelectionTip.ja, LABELS.targetSelectionTip.en);

    /* 「アクティブを基準」選択中か判定 / Check whether “Based on Active Artboard” is selected */
    function isTargetSelection() {
        try { return rbTargetSelection.value === true; } catch (e) { return false; }
    }

    var gTop = w.add("group");
    gTop.orientation = "row";
    gTop.alignChildren = ["fill", "top"];

    var gLeft = gTop.add("group");
    gLeft.orientation = "column";
    gLeft.alignChildren = ["fill", "top"];

    var gRight = gTop.add("group");
    gRight.orientation = "column";
    gRight.alignChildren = ["fill", "top"];
    gRight.alignment = ["right", "top"];

    var pCorner = gLeft.add("panel", undefined, _(LABELS.targetCorner.ja, LABELS.targetCorner.en));
    pCorner.orientation = "column";
    pCorner.alignChildren = ["center", "top"];
    pCorner.margins = [15, 18, 15, 10];
    pCorner.spacing = 6;
    pCorner.helpTip = _(LABELS.targetCornerTip.ja, LABELS.targetCornerTip.en);

    var rowTop = pCorner.add("group");
    rowTop.orientation = "row";
    rowTop.alignChildren = ["center", "center"];
    rowTop.spacing = 12;
    var rbLT = rowTop.add("radiobutton", undefined, "");
    var rbTC = rowTop.add("radiobutton", undefined, "");
    var rbRT = rowTop.add("radiobutton", undefined, "");

    var rowMid = pCorner.add("group");
    rowMid.orientation = "row";
    rowMid.alignChildren = ["center", "center"];
    rowMid.spacing = 12;
    var rbLM = rowMid.add("radiobutton", undefined, "");
    var rbC = rowMid.add("radiobutton", undefined, "");
    var rbRM = rowMid.add("radiobutton", undefined, "");

    var rowBot = pCorner.add("group");
    rowBot.orientation = "row";
    rowBot.alignChildren = ["center", "center"];
    rowBot.spacing = 12;
    var rbLB = rowBot.add("radiobutton", undefined, "");
    var rbBC = rowBot.add("radiobutton", undefined, "");
    var rbRB = rowBot.add("radiobutton", undefined, "");

    rbLT.helpTip = _(LABELS.leftTop.ja, LABELS.leftTop.en) + " (q)";
    rbTC.helpTip = _(LABELS.topCenter.ja, LABELS.topCenter.en) + " (w)";
    rbRT.helpTip = _(LABELS.rightTop.ja, LABELS.rightTop.en) + " (e)";
    rbLM.helpTip = _(LABELS.leftMiddle.ja, LABELS.leftMiddle.en) + " (a)";
    rbC.helpTip = _(LABELS.center.ja, LABELS.center.en) + " (s)";
    rbRM.helpTip = _(LABELS.rightMiddle.ja, LABELS.rightMiddle.en) + " (d)";
    rbLB.helpTip = _(LABELS.leftBottom.ja, LABELS.leftBottom.en) + " (z)";
    rbBC.helpTip = _(LABELS.bottomCenter.ja, LABELS.bottomCenter.en) + " (x)";
    rbRB.helpTip = _(LABELS.rightBottom.ja, LABELS.rightBottom.en) + " (c)";

    var allCornerRadios = [rbLT, rbTC, rbRT, rbLM, rbC, rbRM, rbLB, rbBC, rbRB];

    /* コーナーラジオを排他選択 / Select exactly one corner radio */
    function selectCornerRb(target) {
        for (var i = 0; i < allCornerRadios.length; i++) {
            allCornerRadios[i].value = (allCornerRadios[i] === target);
        }
    }

    selectCornerRb(rbLT);
    addCornerKeyHandler(w);
    /* 整列の基準ショートカット：1=すべてのアートボード、2=アクティブを基準 / Target shortcuts: 1=All Artboards, 2=Based on Active Artboard */
    w.addEventListener("keydown", function (event) {
        /* テキスト入力中は整列の基準ショートカットを無効化 / Skip alignment-base shortcuts while typing in edittext */
        if (event.target && event.target.type === "edittext") return;
        var k = event.keyName;
        if (!k) return;
        k = ("" + k).toLowerCase();

        if (k === "1") {
            rbTargetArtboard.value = true;
            rbTargetSelection.value = false;
            event.preventDefault();
            updatePreviewFromUI();
        } else if (k === "2") {
            rbTargetSelection.value = true;
            rbTargetArtboard.value = false;
            event.preventDefault();
            updatePreviewFromUI();
        }
    });

    /* コーナーラジオ用クリックハンドラを作成（3行分を明示的に排他化）
       Create a corner radio click handler and enforce mutual exclusion across the 3 rows */
    function makeCornerClickHandler(rb) {
        return function () {
            selectCornerRb(rb);
            updatePreviewFromUI();
        };
    }
    for (var i = 0; i < allCornerRadios.length; i++) {
        allCornerRadios[i].onClick = makeCornerClickHandler(allCornerRadios[i]);
    }
    rbTargetArtboard.onClick = updatePreviewFromUI;
    rbTargetSelection.onClick = updatePreviewFromUI;

    var pMargin = gRight.add("panel", undefined, _(LABELS.marginValue.ja, LABELS.marginValue.en) + " (" + unitLabel + ")");
    pMargin.orientation = "row";
    pMargin.alignChildren = ["fill", "center"];
    pMargin.margins = [15, 20, 15, 10];
    pMargin.helpTip = _(LABELS.marginTip.ja, LABELS.marginTip.en);

    var leftColumn = pMargin.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["left", "center"];
    var rightColumn = pMargin.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["left", "center"];

    var gMX = leftColumn.add("group");
    gMX.orientation = "row";
    gMX.alignChildren = ["left", "center"];
    gMX.add("statictext", undefined, _(LABELS.marginX.ja, LABELS.marginX.en));
    var etMarginX = gMX.add("edittext", undefined, "0");
    etMarginX.characters = 4;

    var gMY = leftColumn.add("group");
    gMY.orientation = "row";
    gMY.alignChildren = ["left", "center"];
    gMY.add("statictext", undefined, _(LABELS.marginY.ja, LABELS.marginY.en));
    var etMarginY = gMY.add("edittext", undefined, "0");
    etMarginY.characters = 4;

    var cbLink = rightColumn.add("checkbox", undefined, _(LABELS.link.ja, LABELS.link.en));
    cbLink.value = true;
    cbLink.helpTip = _(LABELS.linkTip.ja, LABELS.linkTip.en);

    /* 左右マージン変更時の処理 / Handle horizontal margin edits */
    function onMarginXEdit() {
        if (cbLink.value) etMarginY.text = etMarginX.text;
        updatePreviewFromUI();
    }
    /* 上下マージン変更時の処理 / Handle vertical margin edits */
    function onMarginYEdit() {
        if (!cbLink.value) updatePreviewFromUI();
    }
    etMarginX.onChange = etMarginX.onChanging = onMarginXEdit;
    etMarginY.onChange = etMarginY.onChanging = onMarginYEdit;

    changeValueByArrowKey(etMarginX, true, onMarginXEdit);
    changeValueByArrowKey(etMarginY, true, onMarginYEdit);

    /* Enter/ReturnでOK実行 / Execute OK on Enter or Return */
    function bindEnterToOK(editText) {
        editText.addEventListener("keydown", function (ev) {
            if (ev.keyName === "Enter" || ev.keyName === "Return") {
                try { btnOK.notify(); } catch (eNotify) { }
                ev.preventDefault();
            }
        });
    }
    /* プレビュー境界オプション / Preview bounds option */
    var gOptions = w.add("group");
    gOptions.orientation = "row";
    gOptions.alignment = ["left", "top"];
    gOptions.margins = [4, 4, 4, 4];

    var cbUsePreviewBounds = gOptions.add("checkbox", undefined, _(LABELS.usePreviewBounds.ja, LABELS.usePreviewBounds.en));
    cbUsePreviewBounds.value = usePreviewBounds;
    cbUsePreviewBounds.helpTip = _(LABELS.usePreviewBoundsTip.ja, LABELS.usePreviewBoundsTip.en);
    cbUsePreviewBounds.onClick = function () {
        usePreviewBounds = cbUsePreviewBounds.value;
        updatePreviewFromUI();
    };

    var gButtons = w.add("group");
    gButtons.orientation = "row";
    gButtons.alignChildren = ["center", "center"];
    gButtons.alignment = ["center", "bottom"];

    gButtons.add("button", undefined, _(LABELS.cancel.ja, LABELS.cancel.en), { name: "cancel" });
    var btnOK = gButtons.add("button", undefined, _(LABELS.ok.ja, LABELS.ok.en), { name: "ok" });

    try { w.defaultElement = btnOK; } catch (eDef) { }

    bindEnterToOK(etMarginX);
    bindEnterToOK(etMarginY);

    /* 連動ON/OFF時のUI更新 / Update UI when link mode changes */
    function updateLinkStateUI() {
        if (cbLink.value) {
            /* 連動ON：左右を操作元に、上下入力を無効化 / Linked ON: use horizontal as source and disable vertical input */
            etMarginX.enabled = true;
            etMarginY.enabled = false;
            etMarginY.text = etMarginX.text;
        } else {
            /* 連動OFF：左右・上下を個別に操作 / Linked OFF: edit horizontal and vertical independently */
            etMarginX.enabled = true;
            etMarginY.enabled = true;
        }
        updatePreviewFromUI();
    }
    cbLink.onClick = updateLinkStateUI;

    /* モードに応じて整列先パネル／マージン欄の有効状態を同期
       Sync enabled state of the corner panel, margin panel, and link checkbox */
    function syncDialogEnabledState(targetIsSelection, isCenter) {
        try { pCorner.enabled = !targetIsSelection; } catch (eDimCorner) { }
        var enableMargin = (!isCenter && !targetIsSelection);
        try { pMargin.enabled = enableMargin; } catch (eDim) { }
        try { cbLink.enabled = enableMargin; } catch (eDim2) { }
    }

    /* 入力欄からマージンをpt単位で取得（中央・選択モード時は0）
       Read margin input as pt; use 0 in Center or Selection mode */
    function readMarginPt(isCenter, targetIsSelection) {
        if (isCenter || targetIsSelection) return { x: 0, y: 0 };
        var mxInput = safeToNumber(etMarginX.text, 0);
        var myInput = safeToNumber(etMarginY.text, 0);
        var sourceY = cbLink.value ? mxInput : myInput;
        var mxPt = toPt(mxInput, unitLabel);
        var myPt = toPt(sourceY, unitLabel);
        return { x: (mxPt === null ? 0 : mxPt), y: (myPt === null ? 0 : myPt) };
    }

    /* UI状態を読み取り、有効状態の同期とプレビュー反映を実行
       Read UI state, sync enabled states, and update the preview */
    function updatePreviewFromUI() {
        var targetIsSelection = isTargetSelection();
        var corner = getCornerFromUI();
        var isCenter = (corner === "C");

        syncDialogEnabledState(targetIsSelection, isCenter);

        var margin = readMarginPt(isCenter, targetIsSelection);
        applyPreview(doc, previewState, corner, margin.x, margin.y, targetIsSelection ? "selection" : "artboard");
    }

    /* ダイアログ終了時はプレビューを巻き戻す（OK後はmainで確定実行） / Revert preview on dialog close; OK is applied later in main() */
    w.onClose = function () {
        revertPreview(previewState, closingByOK ? false : true);
        return true;
    };

    /* 表示時に左右マージン欄へフォーカス / Focus and select the horizontal margin field on show */
    w.onShow = function () {
        try {
            etMarginX.active = true;
            etMarginX.selection = [0, (etMarginX.text || "").length];
        } catch (eFocus) { }
    };

    /* 初期表示でプレビュー適用（選択がある場合） / Apply initial preview when selection exists */
    updateLinkStateUI();

    btnOK.onClick = function () {
        closingByOK = true;
        revertPreview(previewState, false);
        w.close(1);
    };

    if (w.show() !== 1) return null;

    var corner = getCornerFromUI();
    var isCenter = (corner === "C");
    var targetIsSelection = isTargetSelection();
    var marginPt = readMarginPt(isCenter, targetIsSelection);
    return {
        corner: corner,
        targetMode: targetIsSelection ? "selection" : "artboard",
        /* 「アクティブを基準」または「中央」のときは0、それ以外は{x, y} / 0 when selection mode or center; otherwise {x, y} */
        marginPt: (isCenter || targetIsSelection) ? 0 : marginPt
    };
}

/* 文字列を安全に数値化、不正値は fallback / Safe string-to-number with fallback (default: null) */
function safeToNumber(s, fallback) {
    if (fallback === undefined) fallback = null;
    if (s === null || s === undefined) return fallback;
    s = ("" + s).replace(/^\s+|\s+$/g, "");
    if (s === "") return fallback;
    var n = Number(s);
    if (isNaN(n)) return fallback;
    return n;
}

/* 数値入力欄を矢印キーで増減（Shift=10、Option=0.1） / Arrow-key increment for number fields */
function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
    if (allowNegative === undefined) allowNegative = false;

    editText.addEventListener("keydown", function (event) {
        /* Up / Down以外は素通り（数値入力を妨げない） / Pass through non-arrow keys */
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        var value = Number(editText.text);
        if (isNaN(value)) value = 0;

        var keyboard = ScriptUI.environment.keyboardState;
        var direction = (event.keyName === "Up") ? 1 : -1;

        if (keyboard.shiftKey) {
            /* Shift：10の倍数にスナップ / Snap to multiples of 10 */
            var step = 10;
            value = (direction > 0)
                ? Math.ceil((value + 1) / step) * step
                : Math.floor((value - 1) / step) * step;
        } else if (keyboard.altKey) {
            /* Option：0.1単位 / Increment by 0.1 */
            value = Math.round((value + direction * 0.1) * 10) / 10;
        } else {
            /* 通常：1単位、整数丸め / Default: by 1, integer */
            value = Math.round(value + direction);
        }

        if (!allowNegative && value < 0) value = 0;
        editText.text = value;

        if (typeof onValueChanged === "function") {
            try { onValueChanged(); } catch (eCb) { }
        }

        event.preventDefault();
    });
}

// =========================================
// アートボード振り分けと判定 / Artboard mapping & predicates
// =========================================

/* 選択アイテムを中心点が属するアートボードごとに振り分け / Group selection by containing artboard (by center) */
function groupSelectedItemsByArtboard(doc, selection) {
    var result = {};

    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (isLockedOrHidden(item)) continue;

        var itemBounds = getItemBounds(item); /* [左, 上, 右, 下] / [L, T, R, B] */
        if (!itemBounds) continue;

        var cx = (itemBounds[0] + itemBounds[2]) / 2;
        var cy = (itemBounds[1] + itemBounds[3]) / 2;

        var abIndex = findArtboardIndexByPoint(doc, cx, cy);
        if (abIndex === -1) continue;

        if (!result[abIndex]) result[abIndex] = [];
        result[abIndex].push(item);
    }

    return result;
}

/* 座標(x,y)を含むアートボードのインデックスを返す / Return index of artboard containing (x, y) */
function findArtboardIndexByPoint(doc, x, y) {
    var n = doc.artboards.length;
    for (var i = 0; i < n; i++) {
        var r = doc.artboards[i].artboardRect; /* [左, 上, 右, 下] / [L, T, R, B] */
        if (x >= r[0] && x <= r[2] && y <= r[1] && y >= r[3]) return i;
    }
    return -1;
}

/* アイテムまたは親階層がロック/非表示か判定 / Check if item or any ancestor is locked/hidden */
function isLockedOrHidden(item) {
    try { if (item.locked) return true; } catch (e) { }
    try { if (item.hidden) return true; } catch (e2) { }

    var p = item.parent;
    while (p) {
        try { if (p.locked) return true; } catch (e3) { }
        try { if (p.hidden) return true; } catch (e4) { }
        if (p.typename === "Document") break;
        p = p.parent;
    }
    return false;
}

/* 複数アイテムを包含する最小バウンディングを返す / Return union bounds of multiple items */
function getUnionBounds(items) {
    if (!items || items.length === 0) return null;

    var l = null, t = null, r = null, b = null;
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;

        var itemBounds = getItemBounds(item);
        if (!itemBounds) continue;

        if (l === null) {
            l = itemBounds[0]; t = itemBounds[1]; r = itemBounds[2]; b = itemBounds[3];
        } else {
            if (itemBounds[0] < l) l = itemBounds[0];
            if (itemBounds[1] > t) t = itemBounds[1];
            if (itemBounds[2] > r) r = itemBounds[2];
            if (itemBounds[3] < b) b = itemBounds[3];
        }
    }

    if (l === null) return null;
    return [l, t, r, b];
}

// =========================================
// 整列処理 / Alignment
// =========================================

/* 整列先アンカーのアートボード上の座標を計算 / Compute target anchor point on the artboard */
function computeCornerTargetPoint(artboardRect, corner, marginX, marginY) {
    var abL = artboardRect[0];
    var abT = artboardRect[1];
    var abR = artboardRect[2];
    var abB = artboardRect[3];
    var abCx = (abL + abR) / 2;
    var abCy = (abT + abB) / 2;

    if (corner === "LT") return [abL + marginX, abT - marginY];
    if (corner === "TC") return [abCx, abT - marginY];
    if (corner === "RT") return [abR - marginX, abT - marginY];
    if (corner === "LM") return [abL + marginX, abCy];
    if (corner === "C") return [abCx, abCy];
    if (corner === "RM") return [abR - marginX, abCy];
    if (corner === "LB") return [abL + marginX, abB + marginY];
    if (corner === "BC") return [abCx, abB + marginY];
    return [abR - marginX, abB + marginY]; /* 右下 / RB */
}

/* 各アイテムをアートボード基準のアンカーへ整列（マージン付き） / Align each item to artboard anchor with margin */
function alignItemsToCornerWithMargin(items, artboardRect, corner, margin) {
    var marginXY = normalizeMarginXY(margin);

    alignItemsToCornerWithMarginCore(
        items,
        artboardRect,
        corner,
        marginXY.x,
        marginXY.y,
        translateItemDirect
    );
}

/* マージン指定を {x, y} 形式へ正規化 / Normalize margin value to {x, y} */
function normalizeMarginXY(margin) {
    if (typeof margin === "object" && margin !== null) {
        return {
            x: margin.x || 0,
            y: margin.y || 0
        };
    }

    return {
        x: margin || 0,
        y: margin || 0
    };
}

/* 整列処理の共通本体 / Shared item alignment core */
function alignItemsToCornerWithMarginCore(items, artboardRect, corner, marginX, marginY, translateItemForAlignment) {
    var target = computeCornerTargetPoint(artboardRect, corner, marginX, marginY);
    var targetX = target[0];
    var targetY = target[1];

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;

        var itemBounds = getItemBounds(item);
        if (!itemBounds) continue;

        var anchor = getBoundsAnchorPoint(itemBounds, corner);
        var dx = targetX - anchor[0];
        var dy = targetY - anchor[1];

        translateItemForAlignment(item, dx, dy);
    }
}

/* アイテムを直接移動 / Translate item directly */
function translateItemDirect(item, dx, dy) {
    try { item.translate(dx, dy); } catch (e) { }
}

/* プレビュー用：移動量を記録しながら整列（後で巻き戻し可能） / Preview alignment with recorded deltas for later revert */
function alignItemsToCornerWithMarginPreviewXY(items, artboardRect, corner, marginX, marginY, previewState) {
    alignItemsToCornerWithMarginCore(
        items,
        artboardRect,
        corner,
        marginX,
        marginY,
        function (item, dx, dy) {
            translateItemForPreview(item, dx, dy, previewState);
        }
    );
}

/* プレビュー用：アイテムを移動して移動量を記録 / Translate an item and record its preview delta */
function translateItemForPreview(item, dx, dy, previewState) {
    var itemIndex = findPreviewItemIndex(previewState, item);

    try { item.translate(dx, dy); } catch (e) { }

    if (itemIndex >= 0) {
        previewState.dx[itemIndex] += dx;
        previewState.dy[itemIndex] += dy;
    }
}

/* プレビュー状態内のアイテム位置を検索 / Find the item index in preview state */
function findPreviewItemIndex(previewState, targetItem) {
    if (!previewState || !previewState.items) return -1;

    for (var i = 0; i < previewState.items.length; i++) {
        if (previewState.items[i] === targetItem) return i;
    }

    return -1;
}

/* アートボード矩形上の9点アンカー座標を返す（マージンなし） / Return one of the 9 artboard anchor points without margin */
function getArtboardAnchorPoint(artboardRect, corner) {
    return computeCornerTargetPoint(artboardRect, corner, 0, 0);
}

/* バウンディングボックス上の9点アンカー座標を返す / Return one of the 9 bounds anchor points */
function getBoundsAnchorPoint(bounds, corner) {
    /* bounds: [左, 上, 右, 下] / [L, T, R, B] */
    var l = bounds[0], t = bounds[1], r = bounds[2], b = bounds[3];
    var cx = (l + r) / 2;
    var cy = (t + b) / 2;

    if (corner === "C") return [cx, cy];
    if (corner === "LT") return [l, t];
    if (corner === "TC") return [cx, t];
    if (corner === "RT") return [r, t];
    if (corner === "LM") return [l, cy];
    if (corner === "RM") return [r, cy];
    if (corner === "LB") return [l, b];
    if (corner === "BC") return [cx, b];
    return [r, b]; /* RB */
}

/* アクティブアートボード上の選択を基準に他アートボードへ整列 / Align other artboards by active-artboard reference */
function alignByActiveArtboardReference(doc, sourceItems, corner, moveGroupWithOffset) {
    var activeArtboardIndex = -1;
    try { activeArtboardIndex = doc.artboards.getActiveArtboardIndex(); } catch (e) { activeArtboardIndex = -1; }
    if (activeArtboardIndex < 0) return;

    var itemsByArtboard = groupSelectedItemsByArtboard(doc, sourceItems);
    if (!itemsByArtboard) return;

    var referenceItems = itemsByArtboard[activeArtboardIndex];
    if (!referenceItems || referenceItems.length === 0) return;

    var referenceRect = doc.artboards[activeArtboardIndex].artboardRect;
    var offset = computeGroupOffsetFromArtboard(referenceItems, referenceRect, corner);
    if (!offset) return;

    for (var artboardIndex in itemsByArtboard) {
        if (!itemsByArtboard.hasOwnProperty(artboardIndex)) continue;

        var targetArtboardIndex = parseInt(artboardIndex, 10);
        if (targetArtboardIndex === activeArtboardIndex) continue;

        var targetItems = itemsByArtboard[artboardIndex];
        if (!targetItems || targetItems.length === 0) continue;

        var targetRect = doc.artboards[targetArtboardIndex].artboardRect;
        moveGroupWithOffset(targetItems, targetRect, corner, offset);
    }
}

/* テンプレ群とアートボードアンカーの相対オフセットを計算 / Compute relative offset of template group from artboard anchor */
function computeGroupOffsetFromArtboard(items, artboardRect, corner) {
    /* 戻り値：{x, y} = groupAnchor - artboardAnchor / Returns {x, y} offset */
    var groupBounds = getUnionBounds(items);
    if (!groupBounds) return null;

    var artboardAnchor = getArtboardAnchorPoint(artboardRect, corner);
    var groupAnchor = getBoundsAnchorPoint(groupBounds, corner);

    return { x: (groupAnchor[0] - artboardAnchor[0]), y: (groupAnchor[1] - artboardAnchor[1]) };
}

/* アイテム群をオフセット付きでアートボードアンカーに整列 / Align a group to the artboard anchor with an offset */
function alignGroupToArtboardWithOffset(items, artboardRect, corner, offset) {
    var groupBounds = getUnionBounds(items);
    if (!groupBounds) return;

    var artboardAnchor = getArtboardAnchorPoint(artboardRect, corner);
    var groupAnchor = getBoundsAnchorPoint(groupBounds, corner);
    var dx = artboardAnchor[0] + (offset.x || 0) - groupAnchor[0];
    var dy = artboardAnchor[1] + (offset.y || 0) - groupAnchor[1];

    for (var i = 0; i < items.length; i++) {
        try { items[i].translate(dx, dy); } catch (e) { }
    }
}

/* プレビュー用：アイテム群を整列し、移動量を記録 / Preview group alignment with recorded deltas for later revert */
function alignGroupToArtboardWithOffsetPreview(items, artboardRect, corner, offset, previewState) {
    var groupBounds = getUnionBounds(items);
    if (!groupBounds) return;

    var artboardAnchor = getArtboardAnchorPoint(artboardRect, corner);
    var groupAnchor = getBoundsAnchorPoint(groupBounds, corner);
    var dx = artboardAnchor[0] + (offset.x || 0) - groupAnchor[0];
    var dy = artboardAnchor[1] + (offset.y || 0) - groupAnchor[1];

    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) continue;

        translateItemForPreview(item, dx, dy, previewState);
    }
}