#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボードを基準にガイドを整理・作成するツール。次の3系統をダイアログでまとめて設定できる。

- ルーラーガイドの変換：アートボード内に重なるルーラーガイドを検出し、アートボード基準の直線ガイドに引き直す（マスターでON/OFF）
- 中心ガイド：各アートボードの垂直・水平の中心にガイドを作成
- エッジガイド：各アートボードの上下左右にガイドを作成（マスターON/OFF、既定はOFF）
- ガイドが1本も無くても、中心・エッジの作成だけ実行可能
- 作成したガイドはすべて「_guide」レイヤーに集約（無ければ自動作成、ロック/非表示は解除して使用）

設定

- 「外側に延長」「外側へ延長」で、ガイドをアートボードの外側へ延ばす量を指定
- 「すべてのアートボード」は変換側（重なる全アートボード）と中心・エッジ側（全アートボード／アクティブのみ）で個別に指定
- 入力値は現在のルーラー単位として扱い、内部で pt に換算（単位は環境設定の rulerType を参照）

プレビュー

- 設定変更に追従するライブプレビュー（専用レイヤーに色付き線で仮表示し、確定時に本物のガイドへ置換）

### 値の変更

- ↑↓キーで±1増減
- shiftキーを併用すると±10増減

### 紹介記事（note）

https://note.com/dtp_tranist/n/n56d9c936a364

*/

/*

### Overview

A tool to organize/create guides relative to artboards. Three groups are configured together in one dialog:

- Convert ruler guides: detect ruler guides overlapping an artboard and redraw them as artboard-based straight guides (master toggle)
- Center guides: add vertical/horizontal guides at each artboard center
- Edge guides: add guides on each artboard's top/bottom/left/right edges (master toggle, off by default)
- Center/edge creation can run even when there are no guides at all
- All created guides are collected on a "_guide" layer (created if missing; unlocked/shown when reused)

Settings

- "Extend" fields control how far guides reach beyond the artboard edges
- "All artboards" is set independently for conversion (every overlapping artboard) and for center/edge (all artboards vs. the active one)
- Entered values are treated in the current ruler unit and converted to points (unit follows the rulerType preference)

Preview

- Live preview that follows the settings (colored overlay on a dedicated layer, replaced by real guides on commit)

### Changing the value

- Up/Down keys change by ±1
- Hold Shift for ±10

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiCreateArtboardGuides";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 既定の延長（現在のルーラー単位）/ Default extend amount (in current ruler unit) */
var DEFAULT_EXTEND = 0;

/* エッジ描画の既定の延長（mm）/ Default edge extend amount (mm) */
var DEFAULT_EDGE_EXTEND_MM = 10;

/* テキストフィールドの文字数幅 / Width of number fields (in characters) */
var FIELD_CHARACTERS = 4;

/* エッジ用チェックボックスの幅（px）/ Width of each edge checkbox (px) */
var EDGE_CHECKBOX_WIDTH = 48;

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在の言語を判定 / Detect current language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "アートボードガイドの作成", en: "Create Artboard Guides" }
    },
    field: {
        extend: { ja: "外側に延長", en: "Extend" },
        edgeExtend: { ja: "外側へ延長", en: "Extend Beyond Edge" }
    },
    panel: {
        convert: { ja: "ルーラーガイドを変換", en: "Convert ruler guides" },
        edge: { ja: "中心とエッジにガイドを描画", en: "Center & Edge Guides" }
    },
    edge: {
        top: { ja: "上", en: "Top" },
        left: { ja: "左", en: "Left" },
        right: { ja: "右", en: "Right" },
        bottom: { ja: "下", en: "Bottom" }
    },
    checkbox: {
        allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
        convertGuides: { ja: "ルーラーガイドを変換", en: "Convert ruler guides" },
        drawEdges: { ja: "エッジ", en: "Draw edge guides" },
        centerVertical: { ja: "中心（垂直）", en: "Draw vertical center guide" },
        centerHorizontal: { ja: "中心（水平）", en: "Draw horizontal center guide" },
        preview: { ja: "プレビュー", en: "Preview" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tip: {
        extend: {
            ja: "アートボードの端から外側へ延長する長さ（0で端ぴったり）。↑↓で増減、Shiftで±10",
            en: "How far to extend beyond the artboard edge (0 = flush with the edge). Arrow keys to step; Shift ±10"
        },
        edgeExtend: {
            ja: "エッジガイドをアートボードの角から外側へ延長する長さ",
            en: "How far the edge guides extend beyond the artboard corners"
        },
        edgeTop: { ja: "上辺にガイドを作成", en: "Create a guide on the top edge." },
        edgeBottom: { ja: "下辺にガイドを作成", en: "Create a guide on the bottom edge." },
        edgeLeft: { ja: "左辺にガイドを作成", en: "Create a guide on the left edge." },
        edgeRight: { ja: "右辺にガイドを作成", en: "Create a guide on the right edge." },
        target: {
            ja: "アートボード内にある、変換対象のガイドの数",
            en: "Number of guides inside artboards that will be converted"
        },
        allArtboards: {
            ja: "ON：ルーラーガイドが重なるすべてのアートボードを対象にします（OFF：最初の1つだけ）",
            en: "On: target every artboard the ruler guide overlaps (Off: only the first one)"
        },
        drawAllArtboards: {
            ja: "ON：すべてのアートボードに描画（OFF：アクティブなアートボードのみ）",
            en: "On: draw on every artboard (Off: the active artboard only)"
        },
        centerVertical: {
            ja: "アートボードを左右に分ける垂直の中心ガイドを作成",
            en: "Create a vertical guide at the horizontal center"
        },
        centerHorizontal: {
            ja: "アートボードを上下に分ける水平の中心ガイドを作成",
            en: "Create a horizontal guide at the vertical center"
        },
        drawEdges: {
            ja: "OFFにすると、このパネルのエッジ描画設定を無効化します",
            en: "Turn off to disable all edge settings in this panel"
        },
        convertGuides: {
            ja: "OFFにすると、ルーラーガイドの変換を行いません",
            en: "Turn off to skip converting ruler guides"
        }
    },
    hint: {
        noConvertTargets: {
            ja: "変換できるガイドがありません（エッジ描画のみ実行できます）",
            en: "No guides to convert (you can still draw edge guides)"
        }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
    }
};

/* LABELS からドット区切りのキーで文言を取得 / Resolve a label by dotted key */
function getLocalizedText(key) {
    var parts = key.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return key;
        node = node[parts[i]];
    }
    if (node == null) return key;
    return node[currentLanguage] || node.en || key;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLocalizedText(key) + (currentLanguage === "ja" ? "：" : ":");
}

/* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
function labelWithCount(key, count) {
    if (currentLanguage === "ja") {
        return getLocalizedText(key) + "（" + count + "）";
    }
    return getLocalizedText(key) + " (" + count + ")";
}

// =========================================
// 単位 / Units
// =========================================

/* 単位コード→ラベルと pt 換算係数のマップ / Map of unit code to label and points-per-unit */
var unitInfoMap = {
    0: { label: "in", points: 72.0 },
    1: { label: "mm", points: 72.0 / 25.4 },
    2: { label: "pt", points: 1.0 },
    3: { label: "pica", points: 12.0 },
    4: { label: "cm", points: 72.0 / 2.54 },
    5: { label: "Q/H", points: (72.0 / 25.4) * 0.25 },
    6: { label: "px", points: 1.0 },
    7: { label: "ft/in", points: 864.0 },
    8: { label: "m", points: (72.0 / 25.4) * 1000.0 },
    9: { label: "yd", points: 2592.0 },
    10: { label: "ft", points: 864.0 }
};

/* 既定の単位情報（pt）/ Fallback unit info (pt) */
var FALLBACK_UNIT_INFO = { label: "pt", points: 1.0 };

/* 現在のルーラー単位情報を取得 / Get current ruler unit info */
function getCurrentUnitInfo() {
    return unitInfoMap[app.preferences.getIntegerPreference("rulerType")] || FALLBACK_UNIT_INFO;
}

/* 現在の単位ラベルを取得 / Get current unit label */
function getCurrentUnitLabel() {
    return getCurrentUnitInfo().label;
}

/* 現在の単位 1 あたりの pt 数を取得 / Get points per current unit */
function getCurrentPointsPerUnit() {
    return getCurrentUnitInfo().points;
}

/* mm を現在のルーラー単位の値へ換算（小数第1位で丸め）/ Convert mm to current ruler unit (rounded to 1 decimal) */
function mmToCurrentUnit(millimeters) {
    var points = millimeters * (72.0 / 25.4);
    var value = points / getCurrentPointsPerUnit();
    return Math.round(value * 10) / 10;
}

// =========================================
// UI ヘルパー / UI helpers
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* テキストを数値へ（不正なら 0）/ Parse text to a number (0 when invalid) */
function parseNumberOrZero(text) {
    var value = parseFloat(text);
    return isNaN(value) ? 0 : value;
}

/* ↑↓キー操作後の次の値を計算（純粋関数）/ Compute the next value after an arrow key (pure) */
function computeArrowValue(currentValue, direction, keyboard) {
    // direction: 1 = Up, -1 = Down
    if (keyboard.shiftKey) {
        // Shift: 10の倍数にスナップ / snap to multiples of 10
        if (direction > 0) return Math.ceil((currentValue + 1) / 10) * 10;
        return Math.max(0, Math.floor((currentValue - 1) / 10) * 10);
    }
    // 通常: 1単位（下限0）/ step by 1 (clamped at 0)
    if (direction > 0) return currentValue + 1;
    return Math.max(0, currentValue - 1);
}

/* ↑↓キーで値を増減（shift で±10）/ Step value with arrow keys */
function changeValueByArrowKey(inputField) {
    inputField.addEventListener("keydown", function (event) {
        // 入れ子三項は括弧で右結合を明示（ExtendScriptは左結合に誤評価）/ Parenthesize: ExtendScript mis-parses nested ternary
        var direction = (event.keyName === "Up") ? 1 : ((event.keyName === "Down") ? -1 : 0);
        if (direction === 0) return;

        var currentValue = Number(inputField.text);
        if (isNaN(currentValue)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var nextValue = Math.round(computeArrowValue(currentValue, direction, keyboard));

        inputField.text = nextValue;
        // プログラム変更は onChanging を発火しないため明示的に呼ぶ / fire onChanging manually (programmatic change doesn't)
        if (typeof inputField.onChanging === "function") inputField.onChanging();
        event.preventDefault();
    });
}

/* 「ラベル＋数値入力＋単位」の行を追加 / Add a labeled number field with a unit suffix */
function addUnitField(parent, labelKey, defaultText, unitLabel, tooltipKey) {
    var fieldGroup = parent.add("group");
    fieldGroup.orientation = "row";
    fieldGroup.alignChildren = ["left", "center"];

    var fieldLabel = fieldGroup.add("statictext", undefined, labelText(labelKey));
    var inputField = fieldGroup.add("edittext", undefined, defaultText);
    inputField.characters = FIELD_CHARACTERS;
    fieldGroup.add("statictext", undefined, unitLabel);
    changeValueByArrowKey(inputField);

    if (tooltipKey) {
        var tooltip = getLocalizedText(tooltipKey);
        fieldLabel.helpTip = tooltip;
        inputField.helpTip = tooltip;
    }

    return inputField;
}

/* 中心とエッジのガイドパネルを構築 / Build the center & edge guides panel */
function buildEdgePanel(parent, unitLabel) {
    var edgePanel = parent.add("panel", undefined, getLocalizedText("panel.edge"));
    setupPanel(edgePanel);

    /* 中心ガイド（垂直・水平。マスターとは独立）/ Center guides (vertical/horizontal; independent of the edge master) */
    var verticalCheckbox = edgePanel.add("checkbox", undefined, getLocalizedText("checkbox.centerVertical"));
    verticalCheckbox.helpTip = getLocalizedText("tip.centerVertical");
    verticalCheckbox.value = false;

    var horizontalCheckbox = edgePanel.add("checkbox", undefined, getLocalizedText("checkbox.centerHorizontal"));
    horizontalCheckbox.helpTip = getLocalizedText("tip.centerHorizontal");
    horizontalCheckbox.value = false;

    /* エッジ描画のマスタースイッチ（既定OFF。OFFで十字・延長をディム）/ Edge master toggle (default OFF; dims the cross and extend) */
    var drawEdgesCheckbox = edgePanel.add("checkbox", undefined, getLocalizedText("checkbox.drawEdges"));
    drawEdgesCheckbox.helpTip = getLocalizedText("tip.drawEdges");
    drawEdgesCheckbox.value = false;

    /* 上・左右・下の十字配置（各行を中央寄せ）/ Cross layout (each row centered) */
    var topRow = edgePanel.add("group");
    topRow.alignment = "center";
    var topCheckbox = topRow.add("checkbox", undefined, getLocalizedText("edge.top"));

    var middleRow = edgePanel.add("group");
    middleRow.orientation = "row";
    middleRow.alignment = "center";
    middleRow.spacing = 24;
    var leftCheckbox = middleRow.add("checkbox", undefined, getLocalizedText("edge.left"));
    var rightCheckbox = middleRow.add("checkbox", undefined, getLocalizedText("edge.right"));

    var bottomRow = edgePanel.add("group");
    bottomRow.alignment = "center";
    var bottomCheckbox = bottomRow.add("checkbox", undefined, getLocalizedText("edge.bottom"));

    /* 各チェックボックスの幅を少し広げる / Slightly widen each checkbox */
    topCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
    leftCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
    rightCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
    bottomCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;

    /* 各辺の説明 tooltip / Per-edge tooltips */
    topCheckbox.helpTip = getLocalizedText("tip.edgeTop");
    bottomCheckbox.helpTip = getLocalizedText("tip.edgeBottom");
    leftCheckbox.helpTip = getLocalizedText("tip.edgeLeft");
    rightCheckbox.helpTip = getLocalizedText("tip.edgeRight");

    /* 既定はすべて ON / Default all ON */
    topCheckbox.value = leftCheckbox.value = rightCheckbox.value = bottomCheckbox.value = true;

    /* エッジの延長（メインとは別値、既定 10mm）/ Edge extend length (separate value, default 10mm) */
    var edgeExtendInput = addUnitField(edgePanel, "field.edgeExtend", String(mmToCurrentUnit(DEFAULT_EDGE_EXTEND_MM)), unitLabel, "tip.edgeExtend");

    /* 中心・エッジの描画スコープ（すべて / アクティブのみ。マスターとは独立）/ Center & edge drawing scope (independent of masters) */
    var allArtboardsCheckbox = edgePanel.add("checkbox", undefined, getLocalizedText("checkbox.allArtboards"));
    allArtboardsCheckbox.helpTip = getLocalizedText("tip.drawAllArtboards");
    allArtboardsCheckbox.value = false;

    /* 有効/無効の同期 / Sync enabled state */
    var edgeOnlyControls = [topRow, middleRow, bottomRow];
    function syncEdgeEnabled() {
        // 十字（上下左右）はエッジマスターのみで制御 / Cross is controlled by the edge master only
        for (var i = 0; i < edgeOnlyControls.length; i++) {
            edgeOnlyControls[i].enabled = drawEdgesCheckbox.value;
        }
        // 中心(垂直/水平) か エッジ のいずれかが有効か / Whether center or edge is on
        var anyDraw = drawEdgesCheckbox.value || verticalCheckbox.value || horizontalCheckbox.value;
        // 延長・すべてのアートボードは、何か1つでも描画ONなら使える / Extend and scope are usable when anything draws
        edgeExtendInput.parent.enabled = anyDraw;
        allArtboardsCheckbox.enabled = anyDraw;
    }
    drawEdgesCheckbox.onClick = syncEdgeEnabled;
    verticalCheckbox.onClick = syncEdgeEnabled;
    horizontalCheckbox.onClick = syncEdgeEnabled;
    syncEdgeEnabled(); // 初期状態を反映 / apply initial state

    return {
        verticalCheckbox: verticalCheckbox,
        horizontalCheckbox: horizontalCheckbox,
        drawEdgesCheckbox: drawEdgesCheckbox,
        topCheckbox: topCheckbox,
        leftCheckbox: leftCheckbox,
        rightCheckbox: rightCheckbox,
        bottomCheckbox: bottomCheckbox,
        extendInput: edgeExtendInput,
        allArtboardsCheckbox: allArtboardsCheckbox
    };
}

/* 変換パネル（タイトルに件数）を構築 / Build the convert panel (count shown in the title) */
function buildConvertPanel(parent, unitLabel, convertibleCount) {
    var convertPanel = parent.add("panel", undefined, labelWithCount("panel.convert", convertibleCount));
    setupPanel(convertPanel);
    convertPanel.helpTip = getLocalizedText("tip.target");

    /* 変換のマスタースイッチ（OFFで以下をディム）/ Master toggle (OFF dims the rest) */
    var convertGuidesCheckbox = convertPanel.add("checkbox", undefined, getLocalizedText("checkbox.convertGuides"));
    convertGuidesCheckbox.helpTip = getLocalizedText("tip.convertGuides");
    convertGuidesCheckbox.value = true;

    var extendInput = addUnitField(convertPanel, "field.extend", String(DEFAULT_EXTEND), unitLabel, "tip.extend");

    /* 変換のスコープ（重なるすべてのアートボード。一番下）/ Conversion scope (every overlapping artboard; bottom) */
    var allArtboardsCheckbox = convertPanel.add("checkbox", undefined, getLocalizedText("checkbox.allArtboards"));
    allArtboardsCheckbox.helpTip = getLocalizedText("tip.allArtboards");
    allArtboardsCheckbox.value = false;

    /* マスターOFF時にディムする要素 / Elements dimmed when the master toggle is OFF */
    var dimmableControls = [extendInput.parent, allArtboardsCheckbox];
    convertGuidesCheckbox.onClick = function () {
        for (var i = 0; i < dimmableControls.length; i++) {
            dimmableControls[i].enabled = convertGuidesCheckbox.value;
        }
    };

    /* 変換対象が無ければ説明を出してマスターごと無効化 / No targets: show a note and disable the whole section */
    if (convertibleCount === 0) {
        convertGuidesCheckbox.value = false;
        convertGuidesCheckbox.enabled = false;
        convertPanel.add("statictext", undefined, getLocalizedText("hint.noConvertTargets"));
    }
    convertGuidesCheckbox.onClick(); // 初期状態を反映 / apply initial state

    return {
        convertGuidesCheckbox: convertGuidesCheckbox,
        allArtboardsCheckbox: allArtboardsCheckbox,
        extendInput: extendInput
    };
}

/* 延長とエッジ描画を入力するダイアログを表示（ライブプレビュー付き）/ Show the dialog (with live preview) */
function showExtendDialog(doc, convertTargets) {
    var unitLabel = getCurrentUnitLabel();
    var convertibleCount = convertTargets.length;

    /* ダイアログ本体 / Dialog window */
    var dialog = new Window("dialog", getLocalizedText("dialog.title") + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];
    dialog.margins = 16;
    dialog.spacing = 10;

    /* 変換パネル（件数はタイトルに表示）/ Convert panel (count in the title) */
    var convertControls = buildConvertPanel(dialog, unitLabel, convertibleCount);
    var extendInput = convertControls.extendInput;

    /* 中心とエッジのガイドパネル / Center & edge guides panel */
    var edgeControls = buildEdgePanel(dialog, unitLabel);

    /* プレビュー切り替え（既定ON）/ Preview toggle (default ON) */
    var previewCheckbox = dialog.add("checkbox", undefined, getLocalizedText("checkbox.preview"));
    previewCheckbox.value = true;

    /* ボタン（左右中央・上にマージン5・Mac 順：Cancel → OK）/ Buttons (centered, 5px top margin, Mac order: Cancel → OK) */
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.margins = [0, 5, 0, 0];
    var cancelButton = buttonGroup.add("button", undefined, getLocalizedText("button.cancel"), { name: "cancel" });
    var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

    /* 現在のUIから設定を読み取る / Read options from the current UI */
    function readOptions() {
        return {
            convertGuides: convertControls.convertGuidesCheckbox.value,
            extend: parseNumberOrZero(extendInput.text),
            allArtboards: convertControls.allArtboardsCheckbox.value,
            center: {
                vertical: edgeControls.verticalCheckbox.value,
                horizontal: edgeControls.horizontalCheckbox.value
            },
            drawAllArtboards: edgeControls.allArtboardsCheckbox.value,
            drawEdges: edgeControls.drawEdgesCheckbox.value,
            edgeExtend: parseNumberOrZero(edgeControls.extendInput.text),
            edges: {
                top: edgeControls.topCheckbox.value,
                left: edgeControls.leftCheckbox.value,
                right: edgeControls.rightCheckbox.value,
                bottom: edgeControls.bottomCheckbox.value
            }
        };
    }

    /* プレビュー状態 / Preview state */
    var previewColor = makePreviewColor(doc);
    var hiddenGuides = [];   // 一時的に隠した元ガイド / originals temporarily hidden

    /* 隠した元ガイドを元に戻す / Restore originals that were hidden */
    function restoreHiddenGuides() {
        for (var k = 0; k < hiddenGuides.length; k++) {
            try { hiddenGuides[k].hidden = false; } catch (e) {}
        }
        hiddenGuides = [];
    }

    /* プレビューを消去（専用レイヤーごと削除）（app.undo は使わない）/ Clear preview (drop the whole layer; no app.undo) */
    function clearPreview() {
        removePreviewLayer(doc);
        restoreHiddenGuides();
    }

    /* プレビューを再描画 / Re-render the preview */
    function renderPreview() {
        clearPreview();
        if (previewCheckbox.value) {
            try {
                var layer = createPreviewLayer(doc);
                var opt = readOptions();
                var pointsPerUnit = getCurrentPointsPerUnit();
                var extendPoints = opt.extend * pointsPerUnit;
                var edgeExtendPoints = opt.edgeExtend * pointsPerUnit;

                /* 変換：元を隠して新規（色付き線）を描く / Conversion: hide originals, draw colored lines */
                if (opt.convertGuides) {
                    for (var i = 0; i < convertTargets.length; i++) {
                        var target = convertTargets[i];
                        try { target.guidePath.hidden = true; hiddenGuides.push(target.guidePath); } catch (e) {}
                        addPreviewSegments(layer, convertTargetSegments(target, opt.allArtboards, extendPoints), previewColor);
                    }
                }

                /* 中心・エッジ / Center & edge */
                var rects = getTargetArtboardRects(doc, opt.drawAllArtboards);
                var anyEdge = opt.drawEdges &&
                    (opt.edges.top || opt.edges.bottom || opt.edges.left || opt.edges.right);
                for (var r = 0; r < rects.length; r++) {
                    if (anyEdge) {
                        addPreviewSegments(layer, edgeSegments(rects[r], opt.edges, edgeExtendPoints), previewColor);
                    }
                    if (opt.center.vertical || opt.center.horizontal) {
                        addPreviewSegments(layer, centerSegments(rects[r], opt.center, edgeExtendPoints), previewColor);
                    }
                }

                layer.locked = true; // 選択不可に / make non-selectable
            } catch (e) {
                clearPreview();
            }
        }
        app.redraw();
    }

    /* 既存 onClick を保持しつつプレビュー更新を連結（addEventListener併用は不安定なため）/ Chain renderPreview after any existing onClick */
    function chainPreview(control) {
        var previousOnClick = control.onClick;
        control.onClick = function () {
            if (previousOnClick) previousOnClick();
            renderPreview();
        };
    }

    /* 変更を監視してプレビュー更新 / Update preview on any change */
    var previewTriggers = [
        convertControls.convertGuidesCheckbox,
        convertControls.allArtboardsCheckbox,
        edgeControls.verticalCheckbox,
        edgeControls.horizontalCheckbox,
        edgeControls.drawEdgesCheckbox,
        edgeControls.topCheckbox,
        edgeControls.leftCheckbox,
        edgeControls.rightCheckbox,
        edgeControls.bottomCheckbox,
        edgeControls.allArtboardsCheckbox,
        previewCheckbox
    ];
    for (var p = 0; p < previewTriggers.length; p++) {
        chainPreview(previewTriggers[p]);
    }
    extendInput.onChanging = renderPreview;
    edgeControls.extendInput.onChanging = renderPreview;

    extendInput.active = true;

    var dialogResult = null;
    okButton.onClick = function () {
        dialogResult = readOptions();
        dialog.close();
    };
    cancelButton.onClick = function () {
        dialogResult = null;
        dialog.close();
    };

    /* 閉じる時は必ずプレビューを後始末（本処理はクリーンな状態で実行）/ Always clean up on close */
    dialog.onClose = function () {
        clearPreview();
        app.redraw();
    };
    /* 表示時に初期プレビュー / Initial preview on show */
    dialog.onShow = function () {
        renderPreview();
    };

    dialog.show();
    return dialogResult;
}

// =========================================
// ガイド処理 / Guide processing
// =========================================

/* 描画対象のアートボード矩形を取得（ON：全部、OFF：アクティブのみ）/ Get target artboard rects (On: all, Off: active only) */
function getTargetArtboardRects(doc, allArtboards) {
    var rects = [];
    if (allArtboards) {
        for (var i = 0; i < doc.artboards.length; i++) {
            rects.push(doc.artboards[i].artboardRect);
        }
    } else {
        rects.push(doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect);
    }
    return rects;
}

/* 確定ガイドを作成するレイヤー名 / Layer name where committed guides are created */
var GUIDE_LAYER_NAME = "_guide";

/* 「_guide」レイヤーを取得（無ければ作成。ロック/非表示は解除）/ Get the "_guide" layer (create if missing; unlock/show) */
function getGuideLayer(doc) {
    var layer;
    try {
        layer = doc.layers.getByName(GUIDE_LAYER_NAME);
    } catch (e) {
        layer = doc.layers.add();
        layer.name = GUIDE_LAYER_NAME;
    }
    layer.locked = false;
    layer.visible = true;
    return layer;
}

/* 直線のガイドを 1 本作成 / Create a single straight guide */
function addGuideLine(doc, startPoint, endPoint) {
    var guidePath = getGuideLayer(doc).pathItems.add();
    guidePath.setEntirePath([startPoint, endPoint]);
    guidePath.stroked = false;
    guidePath.filled = false;
    guidePath.guides = true;
    return guidePath;
}

/* 変換ターゲットの線分配列を生成 / Build line segments for a convert target */
function convertTargetSegments(target, allArtboards, extendPoints) {
    var segments = [];
    var artboardRects = allArtboards ? target.artboardRects : [target.artboardRects[0]];

    for (var j = 0; j < artboardRects.length; j++) {
        var rect = artboardRects[j]; // [left, top, right, bottom]

        if (target.isVertical) {
            // 上方向は +、下方向は -（Illustrator の Y は上が大きい）/ Y grows upward in Illustrator
            segments.push([[target.position, rect[1] + extendPoints], [target.position, rect[3] - extendPoints]]);
        } else {
            // 左方向は -、右方向は + / Extend left and right
            segments.push([[rect[0] - extendPoints, target.position], [rect[2] + extendPoints, target.position]]);
        }
    }
    return segments;
}

/* エッジの線分配列を生成 / Build edge segments for an artboard rect */
function edgeSegments(rect, edges, edgeExtendPoints) {
    var segments = [];
    var abLeft = rect[0], abTop = rect[1], abRight = rect[2], abBottom = rect[3];

    if (edges.top)    segments.push([[abLeft - edgeExtendPoints, abTop],    [abRight + edgeExtendPoints, abTop]]);
    if (edges.bottom) segments.push([[abLeft - edgeExtendPoints, abBottom], [abRight + edgeExtendPoints, abBottom]]);
    if (edges.left)   segments.push([[abLeft, abTop + edgeExtendPoints],    [abLeft, abBottom - edgeExtendPoints]]);
    if (edges.right)  segments.push([[abRight, abTop + edgeExtendPoints],   [abRight, abBottom - edgeExtendPoints]]);
    return segments;
}

/* 中心の線分配列を生成 / Build center segments for an artboard rect */
function centerSegments(rect, center, edgeExtendPoints) {
    var segments = [];
    var centerX = (rect[0] + rect[2]) / 2;
    var centerY = (rect[1] + rect[3]) / 2;

    if (center.vertical) {
        segments.push([[centerX, rect[1] + edgeExtendPoints], [centerX, rect[3] - edgeExtendPoints]]);
    }
    if (center.horizontal) {
        segments.push([[rect[0] - edgeExtendPoints, centerY], [rect[2] + edgeExtendPoints, centerY]]);
    }
    return segments;
}

/* 線分配列からガイドを作成（collector があれば作成物を push）/ Create guides from segments (collect if given) */
function addGuideSegments(doc, segments, collector) {
    for (var i = 0; i < segments.length; i++) {
        var guidePath = addGuideLine(doc, segments[i][0], segments[i][1]);
        if (collector) collector.push(guidePath);
    }
}

/* プレビュー用レイヤー名 / Preview layer name */
var PREVIEW_LAYER_NAME = "__ArtboardGuidesPreview__";

/* プレビュー用の色（ドキュメントのカラースペースに合わせる）/ Preview color (matches doc color space) */
function makePreviewColor(doc) {
    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
        var cmyk = new CMYKColor();
        cmyk.cyan = 0; cmyk.magenta = 90; cmyk.yellow = 0; cmyk.black = 0;
        return cmyk;
    }
    var rgb = new RGBColor();
    rgb.red = 255; rgb.green = 0; rgb.blue = 255;
    return rgb;
}

/* プレビュー用レイヤーを削除 / Remove the preview layer if present */
function removePreviewLayer(doc) {
    try {
        var layer = doc.layers.getByName(PREVIEW_LAYER_NAME);
        layer.locked = false;
        layer.visible = true;
        layer.remove();
    } catch (e) {}
}

/* プレビュー用レイヤーを用意（既存は作り直し）/ Create a fresh preview layer */
function createPreviewLayer(doc) {
    removePreviewLayer(doc);
    var layer = doc.layers.add();
    layer.name = PREVIEW_LAYER_NAME;
    return layer;
}

/* 線分配列を色付きプレビュー線としてレイヤーへ描画 / Draw colored preview lines into a layer */
function addPreviewSegments(layer, segments, color) {
    for (var i = 0; i < segments.length; i++) {
        var path = layer.pathItems.add();
        path.setEntirePath([segments[i][0], segments[i][1]]);
        path.filled = false;
        path.stroked = true;
        path.strokeColor = color;
        path.strokeWidth = 1;
    }
}

/* 検出済みガイドをアートボード基準に引き直す / Redraw detected guides to their artboard(s) */
function convertGuidesToArtboards(doc, convertTargets, allArtboards, extendPoints) {
    for (var i = 0; i < convertTargets.length; i++) {
        var convertTarget = convertTargets[i];
        var segments = convertTargetSegments(convertTarget, allArtboards, extendPoints);

        /* 元ガイドは1回だけ削除（ロック中でも消せるよう先にアンロック）/ Remove the original guide once (unlock first so locked guides can be removed) */
        try { convertTarget.guidePath.locked = false; } catch (e) {}
        convertTarget.guidePath.remove();
        addGuideSegments(doc, segments, null);
    }
}

/* アートボードのエッジにガイドを描画 / Draw guides on the artboard edges */
function drawEdgeGuides(doc, artboardRects, edges, edgeExtendPoints) {
    for (var i = 0; i < artboardRects.length; i++) {
        addGuideSegments(doc, edgeSegments(artboardRects[i], edges, edgeExtendPoints), null);
    }
}

/* アートボードの中心にガイドを描画 / Draw guides at the artboard centers */
function drawCenterGuides(doc, artboardRects, center, edgeExtendPoints) {
    for (var i = 0; i < artboardRects.length; i++) {
        addGuideSegments(doc, centerSegments(artboardRects[i], center, edgeExtendPoints), null);
    }
}

/* ガイドが重なるアートボードの矩形をすべて返す / Return the rects of every artboard the guide overlaps */
function findOverlappingArtboards(doc, isVertical, guideLeft, guideTop, guideRight, guideBottom) {
    var overlappingRects = [];

    for (var i = 0; i < doc.artboards.length; i++) {
        var artboardRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]

        if (isVertical) {
            if (guideLeft >= artboardRect[0] && guideLeft <= artboardRect[2] &&
                Math.max(guideTop, guideBottom) >= artboardRect[3] &&
                Math.min(guideTop, guideBottom) <= artboardRect[1]) {
                overlappingRects.push(artboardRect);
            }
        } else {
            if (guideTop <= artboardRect[1] && guideTop >= artboardRect[3] &&
                Math.max(guideLeft, guideRight) >= artboardRect[0] &&
                Math.min(guideLeft, guideRight) <= artboardRect[2]) {
                overlappingRects.push(artboardRect);
            }
        }
    }

    return overlappingRects;
}

/* アートボードに重なる直線ガイドを変換対象として収集 / Collect straight guides that overlap an artboard */
function collectConvertTargets(doc, guidePaths) {
    var convertTargets = [];

    for (var i = 0; i < guidePaths.length; i++) {
        var guidePath = guidePaths[i];
        // 「ガイドをロック」ON だとルーラーガイドは locked=true になるが、対象から外さない（削除直前に個別アンロック）/ "Lock Guides" sets locked=true on ruler guides; keep them (unlocked right before removal)
        if (guidePath.hidden) continue;

        var guideBounds = guidePath.geometricBounds; // [left, top, right, bottom]
        var guideLeft = guideBounds[0];
        var guideTop = guideBounds[1];
        var guideRight = guideBounds[2];
        var guideBottom = guideBounds[3];

        var isVertical = Math.abs(guideLeft - guideRight) < 0.01;
        var isHorizontal = Math.abs(guideTop - guideBottom) < 0.01;
        if (!isVertical && !isHorizontal) continue;

        var overlappingArtboardRects = findOverlappingArtboards(doc, isVertical, guideLeft, guideTop, guideRight, guideBottom);
        if (overlappingArtboardRects.length === 0) continue;

        convertTargets.push({
            guidePath: guidePath,
            artboardRects: overlappingArtboardRects,
            isVertical: isVertical,
            // 縦ガイドは x（左端）、横ガイドは y（上端）/ x for vertical, y for horizontal
            position: isVertical ? guideLeft : guideTop
        });
    }

    return convertTargets;
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {

    /* ドキュメントの有無を確認 / Check that a document is open */
    if (app.documents.length === 0) {
        alert(getLocalizedText("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var guidePaths = [];

    /* 既存のガイドを収集 / Collect existing guides */
    for (var i = 0; i < doc.pathItems.length; i++) {
        if (doc.pathItems[i].guides) {
            guidePaths.push(doc.pathItems[i]);
        }
    }

    /* 変換対象（アートボード内の直線ガイド）を事前に検出 / Pre-detect convertible guides */
    /* ガイドが無くてもエッジ描画だけ実行できるよう、ここでは終了しない / Do not exit when empty: edge-only runs are allowed */
    var convertTargets = collectConvertTargets(doc, guidePaths);

    /* 延長・エッジ描画をダイアログで取得（プレビュー付き）/ Get options via the dialog (with preview) */
    var options = showExtendDialog(doc, convertTargets);
    if (options === null) {
        return; // キャンセル / Cancelled
    }

    var pointsPerUnit = getCurrentPointsPerUnit();
    var extendPoints = options.extend * pointsPerUnit;
    var edgeExtendPoints = options.edgeExtend * pointsPerUnit;

    /* 変換マスターOFFなら対象を空にしてスキップ / Empty the list to skip conversion when the master is OFF */
    if (!options.convertGuides) {
        convertTargets = [];
    }

    /* ルーラーガイドをアートボード基準に引き直す / Redraw ruler guides to their artboard(s) */
    convertGuidesToArtboards(doc, convertTargets, options.allArtboards, extendPoints);

    /* 中心・エッジの描画対象アートボード（OFFはアクティブのみ）/ Target artboards for center/edge drawing (active only when OFF) */
    var targetArtboardRects = getTargetArtboardRects(doc, options.drawAllArtboards);

    /* アートボードのエッジにガイドを描画（マスターON時のみ）/ Draw edge guides (only when the master toggle is ON) */
    if (options.drawEdges &&
        (options.edges.top || options.edges.bottom || options.edges.left || options.edges.right)) {
        drawEdgeGuides(doc, targetArtboardRects, options.edges, edgeExtendPoints);
    }

    /* アートボードの中心にガイドを描画 / Draw center guides */
    if (options.center.vertical || options.center.horizontal) {
        drawCenterGuides(doc, targetArtboardRects, options.center, edgeExtendPoints);
    }
})();
