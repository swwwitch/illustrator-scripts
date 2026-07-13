#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- アートボード内に置かれたルーラーガイドを検出し、普通の直線ガイドに引き直す
- ダイアログでアートボード端から外側へ伸ばす量を指定できる
- 入力値は現在のルーラー単位として扱い、内部で pt に換算して適用する
- 単位は環境設定の rulerType を参照

### 値の変更

- ↑↓キーで±1増減
- shiftキーを併用すると±10増減

*/

/*

### Overview

- Detects ruler guides placed inside artboards and redraws them as plain straight guides
- A dialog lets you specify how far to extend beyond the artboard edges
- The entered value is treated in the current ruler unit and converted to points internally
- The unit follows the rulerType preference

### Changing the value

- Up/Down keys change by ±1
- Hold Shift for ±10

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 既定の伸ばす量（現在のルーラー単位）/ Default extend amount (in current ruler unit) */
var DEFAULT_EXTEND = 0;

/* エッジ描画の既定の伸ばす量（mm）/ Default edge extend amount (mm) */
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
        title: { ja: "アートボードガイドを作成", en: "Create Artboard Guides" }
    },
    field: {
        extend: { ja: "伸ばす量", en: "Extend" },
        edgeExtend: { ja: "伸ばす量", en: "Extend" }
    },
    panel: {
        convert: { ja: "ルーラーガイドを変換", en: "Convert ruler guides" },
        edge: { ja: "エッジに描画", en: "Draw on edges" }
    },
    edge: {
        top: { ja: "上", en: "Top" },
        left: { ja: "左", en: "Left" },
        right: { ja: "右", en: "Right" },
        bottom: { ja: "下", en: "Bottom" }
    },
    checkbox: {
        allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
        drawEdges: { ja: "エッジにガイドを描画する", en: "Draw guides on edges" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tip: {
        extend: {
            ja: "アートボードの端から外側へ伸ばす長さ（0で端ぴったり）。↑↓で増減、Shiftで±10、Optionで±0.1",
            en: "How far to extend beyond the artboard edge (0 = flush with the edge). Arrow keys to step; Shift ±10, Option ±0.1"
        },
        edgeExtend: {
            ja: "エッジガイドをアートボードの角から外側へ伸ばす長さ",
            en: "How far the edge guides extend beyond the artboard corners"
        },
        edge: {
            ja: "チェックした辺に、アートボードのエッジ位置のガイドを描画します",
            en: "Draws guides at the artboard edges you check"
        },
        target: {
            ja: "アートボード内にある、変換対象のガイドの数",
            en: "Number of guides inside artboards that will be converted"
        },
        allArtboards: {
            ja: "ON：ルーラーガイドが重なるすべてのアートボードを対象にします（OFF：最初の1つだけ）",
            en: "On: target every artboard the ruler guide overlaps (Off: only the first one)"
        },
        drawEdges: {
            ja: "OFFにすると、このパネルのエッジ描画設定を無効化します",
            en: "Turn off to disable all edge settings in this panel"
        }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noGuides: { ja: "ガイドが見つかりません。", en: "No guides were found." }
    }
};

/* LABELS からドット区切りのキーで文言を取得 / Resolve a label by dotted key */
function L(key) {
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
    return L(key) + (currentLanguage === "ja" ? "：" : ":");
}

/* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
function labelWithCount(key, count) {
    if (currentLanguage === "ja") {
        return L(key) + "（" + count + "）";
    }
    return L(key) + " (" + count + ")";
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
    if (keyboard.altKey) {
        // Option: 0.1単位 / step by 0.1
        return currentValue + direction * 0.1;
    }
    // 通常: 1単位（下限0）/ step by 1 (clamped at 0)
    if (direction > 0) return currentValue + 1;
    return Math.max(0, currentValue - 1);
}

/* ↑↓キーで値を増減（shift で±10、option で±0.1）/ Step value with arrow keys */
function changeValueByArrowKey(inputField) {
    inputField.addEventListener("keydown", function (event) {
        // 入れ子三項は括弧で右結合を明示（ExtendScriptは左結合に誤評価）/ Parenthesize: ExtendScript mis-parses nested ternary
        var direction = (event.keyName === "Up") ? 1 : ((event.keyName === "Down") ? -1 : 0);
        if (direction === 0) return;

        var currentValue = Number(inputField.text);
        if (isNaN(currentValue)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var nextValue = computeArrowValue(currentValue, direction, keyboard);

        // Option時は小数第1位、それ以外は整数に丸め / round to 0.1 with Option, else integer
        nextValue = keyboard.altKey ? Math.round(nextValue * 10) / 10 : Math.round(nextValue);

        inputField.text = nextValue;
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
        var tooltip = L(tooltipKey);
        fieldLabel.helpTip = tooltip;
        inputField.helpTip = tooltip;
    }

    return inputField;
}

/* エッジ描画パネル（上・左右・下の十字配置＋伸ばす量）を構築 / Build the edge panel (cross layout + extend field) */
function buildEdgePanel(parent, unitLabel) {
    var edgePanel = parent.add("panel", undefined, L("panel.edge"));
    setupPanel(edgePanel);
    edgePanel.helpTip = L("tip.edge");

    /* エッジ描画のマスタースイッチ（OFFで以下をディム）/ Master toggle (OFF dims the rest) */
    var drawEdgesCheckbox = edgePanel.add("checkbox", undefined, L("checkbox.drawEdges"));
    drawEdgesCheckbox.helpTip = L("tip.drawEdges");
    drawEdgesCheckbox.value = true;

    /* 上・左右・下の十字配置（各行を中央寄せ）/ Cross layout (each row centered) */
    var topRow = edgePanel.add("group");
    topRow.alignment = "center";
    var topCheckbox = topRow.add("checkbox", undefined, L("edge.top"));

    var middleRow = edgePanel.add("group");
    middleRow.orientation = "row";
    middleRow.alignment = "center";
    middleRow.spacing = 24;
    var leftCheckbox = middleRow.add("checkbox", undefined, L("edge.left"));
    var rightCheckbox = middleRow.add("checkbox", undefined, L("edge.right"));

    var bottomRow = edgePanel.add("group");
    bottomRow.alignment = "center";
    var bottomCheckbox = bottomRow.add("checkbox", undefined, L("edge.bottom"));

    /* 各チェックボックスの幅を少し広げる / Slightly widen each checkbox */
    topCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
    leftCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
    rightCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;
    bottomCheckbox.preferredSize.width = EDGE_CHECKBOX_WIDTH;

    /* 既定はすべて ON / Default all ON */
    topCheckbox.value = leftCheckbox.value = rightCheckbox.value = bottomCheckbox.value = true;

    /* エッジの伸ばす量（メインとは別値、既定 10mm）/ Edge extend length (separate value, default 10mm) */
    var edgeExtendInput = addUnitField(edgePanel, "field.edgeExtend", String(mmToCurrentUnit(DEFAULT_EDGE_EXTEND_MM)), unitLabel, "tip.edgeExtend");

    /* マスターOFF時にディムする要素 / Elements dimmed when the master toggle is OFF */
    var dimmableControls = [topRow, middleRow, bottomRow, edgeExtendInput.parent];
    drawEdgesCheckbox.onClick = function () {
        for (var i = 0; i < dimmableControls.length; i++) {
            dimmableControls[i].enabled = drawEdgesCheckbox.value;
        }
    };
    drawEdgesCheckbox.onClick(); // 初期状態を反映 / apply initial state

    return {
        drawEdgesCheckbox: drawEdgesCheckbox,
        topCheckbox: topCheckbox,
        leftCheckbox: leftCheckbox,
        rightCheckbox: rightCheckbox,
        bottomCheckbox: bottomCheckbox,
        extendInput: edgeExtendInput
    };
}

/* 変換パネル（タイトルに件数）を構築 / Build the convert panel (count shown in the title) */
function buildConvertPanel(parent, unitLabel, convertibleCount) {
    var convertPanel = parent.add("panel", undefined, labelWithCount("panel.convert", convertibleCount));
    setupPanel(convertPanel);
    convertPanel.helpTip = L("tip.target");

    /* すべてのアートボードを対象にするか / Target every overlapping artboard */
    var allArtboardsCheckbox = convertPanel.add("checkbox", undefined, L("checkbox.allArtboards"));
    allArtboardsCheckbox.helpTip = L("tip.allArtboards");
    allArtboardsCheckbox.value = false;

    var extendInput = addUnitField(convertPanel, "field.extend", String(DEFAULT_EXTEND), unitLabel, "tip.extend");

    return {
        extendInput: extendInput,
        allArtboardsCheckbox: allArtboardsCheckbox
    };
}

/* 伸ばす量とエッジ描画を入力するダイアログを表示 / Show the dialog for extend amount and edge options */
function showExtendDialog(convertibleCount) {
    var unitLabel = getCurrentUnitLabel();

    /* ダイアログ本体 / Dialog window */
    var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = ["fill", "top"];
    dialog.margins = 16;
    dialog.spacing = 10;

    /* 変換パネル（件数はタイトルに表示）/ Convert panel (count in the title) */
    var convertControls = buildConvertPanel(dialog, unitLabel, convertibleCount);
    var extendInput = convertControls.extendInput;

    /* エッジ描画パネル / Edge panel */
    var edgeControls = buildEdgePanel(dialog, unitLabel);

    /* ボタン（左右中央・Mac 順：Cancel → OK）/ Buttons (centered, Mac order: Cancel → OK) */
    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var cancelButton = buttonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
    var okButton = buttonGroup.add("button", undefined, "OK", { name: "ok" });

    extendInput.active = true;

    var dialogResult = null;
    okButton.onClick = function () {
        dialogResult = {
            extend: parseNumberOrZero(extendInput.text),
            allArtboards: convertControls.allArtboardsCheckbox.value,
            drawEdges: edgeControls.drawEdgesCheckbox.value,
            edgeExtend: parseNumberOrZero(edgeControls.extendInput.text),
            edges: {
                top: edgeControls.topCheckbox.value,
                left: edgeControls.leftCheckbox.value,
                right: edgeControls.rightCheckbox.value,
                bottom: edgeControls.bottomCheckbox.value
            }
        };
        dialog.close();
    };
    cancelButton.onClick = function () {
        dialogResult = null;
        dialog.close();
    };

    dialog.show();
    return dialogResult;
}

// =========================================
// ガイド処理 / Guide processing
// =========================================

/* 直線のガイドを 1 本作成 / Create a single straight guide */
function addGuideLine(doc, startPoint, endPoint) {
    var guidePath = doc.pathItems.add();
    guidePath.setEntirePath([startPoint, endPoint]);
    guidePath.stroked = false;
    guidePath.filled = false;
    guidePath.guides = true;
    return guidePath;
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
        if (guidePath.locked || guidePath.hidden) continue;

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
        alert(L("alert.noDocument"));
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

    if (guidePaths.length === 0) {
        alert(L("alert.noGuides"));
        return;
    }

    /* 変換対象（アートボード内の直線ガイド）を事前に検出 / Pre-detect convertible guides */
    var convertTargets = collectConvertTargets(doc, guidePaths);

    /* 伸ばす量・エッジ描画をダイアログで取得 / Get extend amount and edge options */
    var options = showExtendDialog(convertTargets.length);
    if (options === null) {
        return; // キャンセル / Cancelled
    }

    var pointsPerUnit = getCurrentPointsPerUnit();
    var extendPoints = options.extend * pointsPerUnit;
    var edgeExtendPoints = options.edgeExtend * pointsPerUnit;

    /* 検出済みガイドをアートボード基準に引き直す / Redraw detected guides to the artboard */
    for (var i = 0; i < convertTargets.length; i++) {
        var convertTarget = convertTargets[i];

        /* 元ガイドは1回だけ削除 / Remove the original guide once */
        convertTarget.guidePath.remove();

        /* OFF：最初のアートボードのみ、ON：重なるすべて / Off: first artboard only, On: all overlapping */
        var artboardRects = options.allArtboards
            ? convertTarget.artboardRects
            : [convertTarget.artboardRects[0]];

        for (var j = 0; j < artboardRects.length; j++) {
            var artboardRect = artboardRects[j]; // [left, top, right, bottom]

            if (convertTarget.isVertical) {
                // 上方向は +、下方向は -（Illustrator の Y は上が大きい）/ Y grows upward in Illustrator
                addGuideLine(doc,
                    [convertTarget.position, artboardRect[1] + extendPoints],
                    [convertTarget.position, artboardRect[3] - extendPoints]);
            } else {
                // 左方向は -、右方向は + / Extend left and right
                addGuideLine(doc,
                    [artboardRect[0] - extendPoints, convertTarget.position],
                    [artboardRect[2] + extendPoints, convertTarget.position]);
            }
        }
    }

    /* アートボードのエッジにガイドを描画（マスターON時のみ）/ Draw edge guides (only when the master toggle is ON) */
    if (options.drawEdges &&
        (options.edges.top || options.edges.bottom || options.edges.left || options.edges.right)) {
        for (var i = 0; i < doc.artboards.length; i++) {
            var edgeRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
            var abLeft = edgeRect[0];
            var abTop = edgeRect[1];
            var abRight = edgeRect[2];
            var abBottom = edgeRect[3];

            if (options.edges.top) {
                addGuideLine(doc, [abLeft - edgeExtendPoints, abTop], [abRight + edgeExtendPoints, abTop]);
            }
            if (options.edges.bottom) {
                addGuideLine(doc, [abLeft - edgeExtendPoints, abBottom], [abRight + edgeExtendPoints, abBottom]);
            }
            if (options.edges.left) {
                addGuideLine(doc, [abLeft, abTop + edgeExtendPoints], [abLeft, abBottom - edgeExtendPoints]);
            }
            if (options.edges.right) {
                addGuideLine(doc, [abRight, abTop + edgeExtendPoints], [abRight, abBottom - edgeExtendPoints]);
            }
        }
    }
})();
