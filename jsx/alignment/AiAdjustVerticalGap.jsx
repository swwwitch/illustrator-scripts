#target illustrator
#targetengine "AdjustVerticalGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した2つのオブジェクトの上下の間隔を、指定した値にそろえる常駐パレットです。
ライブプレビュー対応で、設定を変えるたびに結果を確認できます。

- 「キーオブジェクト」（上／下）を基準に、もう一方を移動します
- 間隔値は定規の単位で指定でき、↑↓キー（Shiftで±10／Optionで±0.1）で増減できます
- 左右方向の整列（整列しない／左／中央／右）も同時に行えます
- テキストには段落の行揃え（変更しない／整列に連動／均等配置）を適用できます
- クリップグループはクリップパスを基準に、プレビュー境界（線幅・効果込み）の使用も切替可能
- ［適用］で確定、未確定のまま閉じると元に戻ります
- キー操作：T＝上／B＝下、Option+A＝適用、Esc＝閉じる

### Overview

A docking palette that sets the vertical gap between two selected objects,
with a live preview that updates as you change the settings.

- Keeps the chosen key object (top or bottom) in place and moves the other
- The gap value uses the document's ruler unit (arrow keys: Shift ±10 / Option ±0.1)
- Optional horizontal alignment (none / left / center / right)
- Optional paragraph alignment for text (keep / match align / justify)
- Clip groups measure by their clipping path; preview bounds (stroke/effects) can be toggled
- Apply commits; closing without applying reverts the preview
- Keys: T = top / B = bottom, Option+A = apply, Esc = close

### 実装メモ / Implementation note

常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理（runAdjustment と
その補助関数）はメインエンジンへ BridgeTalk で都度委譲する。委譲本文は関数を toString()
で連結し、encodeURIComponent + eval(decodeURIComponent()) で送って文字化けを防ぐ。
委譲する関数では行コメントを使わずブロックコメントのみ、必ずセミコロンで終える。
ライブプレビューは各プレビューを1トランザクションとし、次回送信時に worker 側で
app.undo() してから再適用する（確定時は取り消さない）。

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ユーザー設定 / User settings
// =========================================
var DEFAULT_GAP_VALUE = "20"; /* 間隔の初期値（定規の単位）/ Default gap (ruler unit) */

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;
var COLUMN_SPACING = 12; /* 2カラムの間隔 / Gap between the two columns */

/* パレットの参照を常駐エンジンに保持 / Keep the palette reference alive in the resident engine */
var paletteWindow = null;

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在の言語を判定（ja / en）/ Detect current language (ja / en) */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "上下の間隔を調整", en: "Adjust Vertical Gap" }
    },
    anchor: {
        title: { ja: "キーオブジェクト", en: "Key object" },
        top: { ja: "上", en: "Top" },
        bottom: { ja: "下", en: "Bottom" }
    },
    gap: {
        title: { ja: "間隔値", en: "Gap" }
    },
    checkbox: {
        previewBounds: { ja: "プレビュー境界", en: "Use preview bounds" }
    },
    align: {
        title: { ja: "左右の整列", en: "Horizontal align" },
        none: { ja: "整列しない", en: "Don't align" },
        left: { ja: "左", en: "Left" },
        center: { ja: "中央", en: "Center" },
        right: { ja: "右", en: "Right" }
    },
    justify: {
        title: { ja: "テキストの行揃え", en: "Text alignment" },
        none: { ja: "変更しない", en: "Keep current" },
        link: { ja: "整列に連動", en: "Match alignment" },
        full: { ja: "均等配置", en: "Justify" }
    },
    button: {
        apply: { ja: "適用", en: "Apply" }
    },
    status: {
        ready: { ja: "オブジェクトを2つ選択して［適用］", en: "Select two objects, then Apply" },
        previewed: { ja: "プレビュー中（［適用］で確定）", en: "Preview (Apply to commit)" },
        done: { ja: "適用しました", en: "Applied" },
        noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open" },
        selectTwo: { ja: "オブジェクトを2つ選択してください", en: "Please select two objects" },
        error: { ja: "エラー:", en: "Error:" }
    },
    tooltip: {
        anchorTop: {
            ja: "基準にする（動かさない）キーオブジェクト。上を選択（ショートカット: T）。",
            en: "Key object to keep in place. Select top (shortcut: T)."
        },
        anchorBottom: {
            ja: "基準にする（動かさない）キーオブジェクト。下を選択（ショートカット: B）。",
            en: "Key object to keep in place. Select bottom (shortcut: B)."
        },
        gap: {
            ja: "上下に並ぶ2つのオブジェクト間の距離です。↑↓キーで増減（Shiftで±10）。単位は定規に従います。",
            en: "Vertical distance between the two objects. Arrow keys change it (Shift: ±10). Unit follows the ruler."
        },
        previewBounds: {
            ja: "線幅や効果を含む見た目の境界で間隔を計算します。オフにすると線幅や効果を含まないパス境界で計算します。",
            en: "Calculate the gap using visual bounds, including strokes and effects. Turn off to use geometric path bounds."
        },
        align: {
            ja: "移動するオブジェクトを、キーオブジェクトの左・中央・右にそろえます。",
            en: "Align the moving object to the left, center, or right of the key object."
        },
        justify: {
            ja: "テキストの段落の行揃え。「整列に連動」は左右の整列（左／中央／右）に合わせます。テキスト以外には影響しません。",
            en: "Paragraph alignment of text. \"Match alignment\" follows the horizontal align. Non-text objects are unaffected."
        },
        justifyFull: {
            ja: "段落を均等配置します（最終行は左揃え）。",
            en: "Justify paragraphs (last line left-aligned)."
        },
        apply: {
            ja: "現在の選択（2つ）に設定を適用します（ショートカット: Option + A）。",
            en: "Apply the settings to the current selection (two objects) (shortcut: Option + A)."
        }
    }
};

/* ドット区切りパスでラベルを取得（途中欠落・null にも耐える）/ Look up a label by dotted path (tolerates missing or null nodes) */
function L(path) {
    var parts = path.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node === null || typeof node !== "object") {
            return path; /* 途中階層が辿れない / cannot descend further */
        }
        node = node[parts[i]];
    }
    if (node === null || typeof node !== "object") {
        return path; /* 葉が { ja, en } でない / leaf is not a localized object */
    }
    return node[currentLanguage] || node.en || path;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(path) {
    return L(path) + (currentLanguage === "ja" ? "：" : ":");
}

// =========================================
// 単位 / Units
// =========================================

/* 定規の単位からラベルと pt 換算係数を求める / Resolve ruler unit label and pt factor */
function getRulerUnitInfo() {
    var rulerUnit = app.preferences.getIntegerPreference("rulerType");
    var unitLabel = "pt";
    var unitFactor = 1.0;

    switch (rulerUnit) {
        case 0: // inch
            unitLabel = "inch";
            unitFactor = 72.0;
            break;
        case 1: // mm
            unitLabel = "mm";
            unitFactor = 72.0 / 25.4;
            break;
        case 2: // pt
            unitLabel = "pt";
            unitFactor = 1.0;
            break;
        case 3: // pica
            unitLabel = "pica";
            unitFactor = 12.0;
            break;
        case 4: // cm
            unitLabel = "cm";
            unitFactor = 72.0 / 2.54;
            break;
        case 5: // Q
            unitLabel = "Q";
            unitFactor = 72.0 / 25.4 * 0.25;
            break;
        case 6: // px
            unitLabel = "px";
            unitFactor = 1.0;
            break;
        default:
            unitLabel = "pt";
            unitFactor = 1.0;
    }

    return { label: unitLabel, factor: unitFactor };
}

// =========================================
// 委譲される処理 / Delegated worker functions
//   メインエンジンへ送って実行する。toString() で連結するため、
//   この節の関数では // コメントを使わず /* */ のみ・必ずセミコロンで終える。
//   Sent to the main engine; use only /* */ comments and explicit semicolons here.
// =========================================

/* クリップグループのクリップパスを返す（無ければ null）/ Return the clipping path of a clip group, or null */
function getClippingPath(groupItem) {
    var paths = groupItem.pathItems;
    for (var i = 0; i < paths.length; i++) {
        if (paths[i].clipping === true) {
            return paths[i];
        }
    }
    /* 複合パスでクリップしている場合 / When the clip path is a compound path */
    var compounds = groupItem.compoundPathItems;
    for (var j = 0; j < compounds.length; j++) {
        if (compounds[j].pathItems.length > 0 && compounds[j].pathItems[0].clipping === true) {
            return compounds[j];
        }
    }
    return null;
}

/* 計算に使う境界を取得（クリップグループはクリップパス基準）/ Bounds for calculation (clip group uses its clipping path) */
function getItemBounds(item, usePreviewBounds) {
    var boundsTarget = item;
    if (item.constructor.name === "GroupItem" && item.clipped === true) {
        var clipPath = getClippingPath(item);
        if (clipPath !== null) {
            boundsTarget = clipPath;
        }
    }
    return usePreviewBounds ? boundsTarget.visibleBounds : boundsTarget.geometricBounds;
}

/* UI の行揃え選択を実際のキーに解決（「整列に連動」は整列値を流用）/ Resolve the choice ("link" follows the align value) */
function resolveJustifyKey(justify, align) {
    if (justify === "link") {
        return (align === "none") ? "none" : align;
    }
    return justify;
}

/* 行揃えキーを Justification 列挙値に変換 / Map key to Justification enum */
function resolveJustification(key) {
    if (key === "left") {
        return Justification.LEFT;
    }
    if (key === "center") {
        return Justification.CENTER;
    }
    if (key === "right") {
        return Justification.RIGHT;
    }
    if (key === "full") {
        return Justification.FULLJUSTIFYLASTLINELEFT;
    }
    return null;
}

/* TextFrame なら段落の行揃えを設定 / Set justification when the object is a TextFrame */
function applyJustification(item, justification) {
    if (justification === null) {
        return;
    }
    if (item.constructor.name === "TextFrame") {
        item.textRange.paragraphAttributes.justification = justification;
    }
}

/* 直前のプレビューを取り消す（メインエンジンで実行）/ Undo the previous preview (runs in the main engine) */
function undoLast() {
    if (app.documents.length === 0) {
        return "NODOC";
    }
    app.undo();
    app.redraw();
    return "OK";
}

/* 選択2点に間隔・整列・行揃えを適用（メインエンジンで実行）/ Apply gap, align and justify to the two selected items */
function runAdjustment(options) {
    /* ドキュメント確認を最優先（undo より先）/ Check for a document first, before any undo */
    if (app.documents.length === 0) {
        return "NODOC";
    }
    /* ライブプレビュー：前回適用分を取り消してからやり直す / Live preview: undo the previous apply first */
    if (options.undoFirst === true) {
        app.undo();
    }
    var sel = app.activeDocument.selection;
    if (sel.length !== 2) {
        app.redraw();
        return "NOSEL";
    }

    var itemA = sel[0];
    var itemB = sel[1];

    /* 行揃えを先に適用（ポイント文字は揃えで境界が変わるため）/ Justify first; point-text bounds depend on it */
    var justifyKey = resolveJustifyKey(options.justify, options.align);
    if (justifyKey !== "none") {
        var justification = resolveJustification(justifyKey);
        applyJustification(itemA, justification);
        applyJustification(itemB, justification);
    }

    var boundsA = getItemBounds(itemA, options.usePreviewBounds);
    var boundsB = getItemBounds(itemB, options.usePreviewBounds);

    /* top が大きい方が上 / The object with the larger top is the upper one */
    var upperItem, lowerItem;
    if (boundsA[1] >= boundsB[1]) {
        upperItem = itemA;
        lowerItem = itemB;
    } else {
        upperItem = itemB;
        lowerItem = itemA;
    }

    var anchorItem = options.anchorTop ? upperItem : lowerItem;
    var movingItem = options.anchorTop ? lowerItem : upperItem;

    /* 上下方向の移動 / Vertical move */
    if (options.anchorTop) {
        var targetLowerTop = getItemBounds(upperItem, options.usePreviewBounds)[3] - options.gapPoints;
        movingItem.translate(0, targetLowerTop - getItemBounds(movingItem, options.usePreviewBounds)[1]);
    } else {
        var targetUpperBottom = getItemBounds(lowerItem, options.usePreviewBounds)[1] + options.gapPoints;
        movingItem.translate(0, targetUpperBottom - getItemBounds(movingItem, options.usePreviewBounds)[3]);
    }

    /* 左右方向の整列 / Horizontal alignment */
    if (options.align !== "none") {
        var anchorBounds = getItemBounds(anchorItem, options.usePreviewBounds);
        var movingBounds = getItemBounds(movingItem, options.usePreviewBounds);
        var dx = 0;
        if (options.align === "left") {
            dx = anchorBounds[0] - movingBounds[0];
        } else if (options.align === "right") {
            dx = anchorBounds[2] - movingBounds[2];
        } else if (options.align === "center") {
            dx = ((anchorBounds[0] + anchorBounds[2]) / 2) - ((movingBounds[0] + movingBounds[2]) / 2);
        }
        movingItem.translate(dx, 0);
    }

    app.redraw();
    return "OK";
}

// =========================================
// BridgeTalk 委譲 / BridgeTalk delegation
// =========================================

/* 委譲する worker 関数（追加したらここにも登録）/ Worker functions to delegate (register new ones here) */
var WORKER_FUNCS = [
    getClippingPath,
    getItemBounds,
    resolveJustifyKey,
    resolveJustification,
    applyJustification,
    undoLast,
    runAdjustment
];

/* options を JS リテラル文字列に変換 / Serialize options to a JS object literal */
function optionsToLiteral(options) {
    return "{"
        + "anchorTop:" + options.anchorTop + ","
        + "gapPoints:" + options.gapPoints + ","
        + "align:\"" + options.align + "\","
        + "justify:\"" + options.justify + "\","
        + "usePreviewBounds:" + options.usePreviewBounds + ","
        + "undoFirst:" + (options.undoFirst === true)
        + "}";
}

/* worker 関数群 + ディスパッチ式のコードを組み立てる / Build worker source + a dispatch expression */
function buildWorkerCode(dispatchExpr) {
    var bodies = [];
    for (var i = 0; i < WORKER_FUNCS.length; i++) {
        bodies.push(WORKER_FUNCS[i].toString());
    }
    return bodies.join("\n") + "\n" + dispatchExpr + ";";
}

/* メインエンジンへ同期委譲（%エンコードで文字化けを防ぐ）/ Delegate synchronously, %-encoded to avoid corruption */
function delegateToMainEngine(code) {
    var bridge = new BridgeTalk();
    bridge.target = "illustrator";
    bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
    var holder = { value: null };
    bridge.onResult = function (response) {
        holder.value = response.body;
    };
    bridge.onError = function (response) {
        holder.value = "ERR:" + response.body;
    };
    bridge.send(10); // 同期送信（最大10秒）/ Synchronous send (up to 10s)
    return (holder.value === null) ? "ERR:timeout" : holder.value;
}

/* プレビューを適用（前回分は worker 側で取り消し）/ Apply a preview (worker undoes the previous one) */
function runAdjustmentPreview(options) {
    return delegateToMainEngine(buildWorkerCode("runAdjustment(" + optionsToLiteral(options) + ")"));
}

/* 直前のプレビューを取り消す / Revert the last preview */
function revertLastPreview() {
    return delegateToMainEngine(buildWorkerCode("undoLast()"));
}

/* 委譲結果をローカライズした文言に変換 / Map a worker result to a localized message */
function describeResult(result) {
    if (result === "NODOC") {
        return L("status.noDocument");
    }
    if (result === "NOSEL") {
        return L("status.selectTwo");
    }
    return L("status.error") + " " + result;
}

// =========================================
// テキストフィールド操作 / Text field helpers
// =========================================

/* ↑↓キーで値を増減（Shiftで±10・10スナップ。Option修飾は非対応）/ Arrow keys change the value (Shift: ±10 snapped; Option is not supported) */
function changeValueByArrowKey(editText, onChangeCallback) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (keyboard.shiftKey) {
            var delta = 10;
            // Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift
            if (event.keyName === "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else {
            if (event.keyName === "Up") {
                value += 1;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value -= 1;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        // 整数に丸め / Round to integer
        value = Math.round(value);
        editText.text = value;

        /* 値変更を通知 / Notify the change */
        if (typeof onChangeCallback === "function") {
            onChangeCallback();
        }
    });
}

/* Option + A で適用を実行するキーハンドラ / Run apply with Option + A */
function addApplyKeyHandler(win, onApply) {
    win.addEventListener("keydown", function (event) {
        var keyboard = ScriptUI.environment.keyboardState;
        if (event.keyName === "A" && keyboard.altKey) {
            onApply();
            event.preventDefault();
        }
    });
}

/* Esc でパレットを閉じるキーハンドラ / Close the palette with Esc */
function addCloseKeyHandler(win) {
    win.addEventListener("keydown", function (event) {
        if (event.keyName === "Escape") {
            win.close();
            event.preventDefault();
        }
    });
}

/* T で上、B で下を選択するキーハンドラ（選択後にプレビュー更新）/ Select top with T, bottom with B (then refresh preview) */
function addAnchorKeyHandler(win, topRadio, bottomRadio, onPreview) {
    win.addEventListener("keydown", function (event) {
        if (event.keyName === "T") {
            topRadio.value = true;
            onPreview();
            event.preventDefault();
        } else if (event.keyName === "B") {
            bottomRadio.value = true;
            onPreview();
            event.preventDefault();
        }
    });
}

// =========================================
// レイアウト共通設定 / Shared layout helpers
// =========================================

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
function setupGroup(group, orientation, spacing) {
    var groupOrientation = orientation || "column";
    group.orientation = groupOrientation;
    /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
    group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// パネル生成 / Panel builders
// =========================================

/* 固定するオブジェクトパネル / Fixed-object panel */
function buildAnchorPanel(parent, onPreview) {
    var panel = parent.add("panel", undefined, L("anchor.title"));
    setupPanel(panel);

    /* 上下は縦並び（ショートカット T/B はラベル非表示）/ Stacked vertically (T/B shortcuts are not shown) */
    var anchorTopRadio = panel.add("radiobutton", undefined, L("anchor.top"));
    var anchorBottomRadio = panel.add("radiobutton", undefined, L("anchor.bottom"));
    anchorTopRadio.value = true;
    anchorTopRadio.helpTip = L("tooltip.anchorTop");
    anchorBottomRadio.helpTip = L("tooltip.anchorBottom");
    anchorTopRadio.onClick = onPreview;
    anchorBottomRadio.onClick = onPreview;

    return { anchorTopRadio: anchorTopRadio, anchorBottomRadio: anchorBottomRadio };
}

/* 間隔値＋プレビュー境界パネル / Gap + preview-bounds panel */
function buildGapPanel(parent, rulerUnit, onPreview) {
    var panel = parent.add("panel", undefined, L("gap.title"));
    setupPanel(panel);

    var row = panel.add("group");
    setupGroup(row, "row");
    var gapValueInput = row.add("edittext", undefined, DEFAULT_GAP_VALUE);
    gapValueInput.characters = 5;
    gapValueInput.helpTip = L("tooltip.gap");
    changeValueByArrowKey(gapValueInput, onPreview);
    gapValueInput.onChange = onPreview;
    row.add("statictext", undefined, rulerUnit.label);

    var previewBoundsCheckbox = panel.add("checkbox", undefined, L("checkbox.previewBounds"));
    previewBoundsCheckbox.value = true;
    previewBoundsCheckbox.helpTip = L("tooltip.previewBounds");
    previewBoundsCheckbox.onClick = onPreview;

    return {
        gapValueInput: gapValueInput,
        previewBoundsCheckbox: previewBoundsCheckbox
    };
}

/* 左右の整列パネル / Horizontal-alignment panel */
function buildAlignPanel(parent, onPreview) {
    var panel = parent.add("panel", undefined, L("align.title"));
    setupPanel(panel);
    panel.helpTip = L("tooltip.align");

    var radios = {
        none: panel.add("radiobutton", undefined, L("align.none")),
        left: panel.add("radiobutton", undefined, L("align.left")),
        center: panel.add("radiobutton", undefined, L("align.center")),
        right: panel.add("radiobutton", undefined, L("align.right"))
    };
    radios.none.value = true;
    radios.none.onClick = onPreview;
    radios.left.onClick = onPreview;
    radios.center.onClick = onPreview;
    radios.right.onClick = onPreview;
    radios.none.helpTip = L("tooltip.align");
    radios.left.helpTip = L("tooltip.align");
    radios.center.helpTip = L("tooltip.align");
    radios.right.helpTip = L("tooltip.align");
    return radios;
}

/* テキストの行揃えパネル / Paragraph-justification panel */
function buildJustifyPanel(parent, onPreview) {
    var panel = parent.add("panel", undefined, L("justify.title"));
    setupPanel(panel);
    panel.helpTip = L("tooltip.justify");

    var radios = {
        none: panel.add("radiobutton", undefined, L("justify.none")),
        link: panel.add("radiobutton", undefined, L("justify.link")),
        full: panel.add("radiobutton", undefined, L("justify.full"))
    };
    radios.link.value = true;
    radios.none.helpTip = L("tooltip.justify");
    radios.link.helpTip = L("tooltip.justify");
    radios.full.helpTip = L("tooltip.justifyFull"); /* （最終行左）はツールチップに / "(last line left)" lives in the tooltip */
    radios.none.onClick = onPreview;
    radios.link.onClick = onPreview;
    radios.full.onClick = onPreview;
    return radios;
}

/* 選択中のラジオに対応するキーを返す / Return the key of the selected radio */
function selectedRadioKey(radios, keys) {
    for (var i = 0; i < keys.length; i++) {
        if (radios[keys[i]].value) {
            return keys[i];
        }
    }
    return keys[0];
}

// =========================================
// パレット / Palette
// =========================================

/* UI 値を読み取り、間隔は pt に換算 / Read UI values, convert gap to points */
function readOptions(controls, rulerUnit) {
    var gapValue = parseFloat(controls.gap.gapValueInput.text);
    if (isNaN(gapValue)) {
        gapValue = parseFloat(DEFAULT_GAP_VALUE);
    }
    if (gapValue < 0) {
        gapValue = 0; /* 間隔は負にできない（手入力対策）/ Gap cannot be negative (guards manual input) */
    }
    /* 正規化した値を表示にも反映 / Reflect the normalized value back to the field */
    if (controls.gap.gapValueInput.text !== String(gapValue)) {
        controls.gap.gapValueInput.text = gapValue;
    }

    return {
        anchorTop: controls.anchor.anchorTopRadio.value,
        gapPoints: gapValue * rulerUnit.factor,
        align: selectedRadioKey(controls.align, ["none", "left", "center", "right"]),
        justify: selectedRadioKey(controls.justify, ["none", "link", "full"]),
        usePreviewBounds: controls.gap.previewBoundsCheckbox.value
    };
}

/* パレットを表示 / Show the palette */
function showPalette() {
    /* 多重起動を防ぐ：既存パレットがあれば閉じる / Prevent duplicates: close any existing palette */
    if (paletteWindow) {
        try {
            paletteWindow.close();
        } catch (e) {}
        paletteWindow = null;
    }

    var rulerUnit = getRulerUnitInfo();

    var win = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
    win.orientation = "column";
    win.alignChildren = "fill";

    var controls = {};
    var previewState = { active: false }; /* プレビューが反映中か / Whether a preview is currently applied */
    var isBusy = false; /* 同期委譲中の再入防止 / Guard against re-entry during a synchronous delegation */

    function setStatus(message) {
        statusText.text = message;
    }

    /* ライブプレビュー更新（前回分は worker 側で取り消し）/ Refresh live preview (worker undoes the previous one) */
    function updatePreview() {
        if (isBusy) return; /* 委譲中に発火した変更は無視 / ignore changes fired mid-delegation */
        isBusy = true;
        var options = readOptions(controls, rulerUnit);
        options.undoFirst = previewState.active;
        var result = runAdjustmentPreview(options);
        previewState.active = (result === "OK");
        setStatus(result === "OK" ? L("status.previewed") : describeResult(result));
        isBusy = false;
    }

    /* 確定（現在のプレビューを残す）/ Commit (keep the current preview) */
    function commitPreview() {
        if (!previewState.active) {
            updatePreview(); /* 未反映なら一度適用 / apply once if not previewed yet */
        }
        if (previewState.active) {
            previewState.active = false; /* 確定したので閉じても取り消さない / committed: do not revert on close */
            setStatus(L("status.done"));
        }
    }

    /* 行1：固定するオブジェクト ／ 間隔値＋プレビュー境界 / Row 1: anchor | gap + preview bounds */
    var topRow = win.add("group");
    setupGroup(topRow, "row", COLUMN_SPACING);
    topRow.alignChildren = ["fill", "top"];
    controls.anchor = buildAnchorPanel(topRow, updatePreview);
    controls.gap = buildGapPanel(topRow, rulerUnit, updatePreview);

    /* 行2：整列設定 ／ テキストの行揃え / Row 2: align settings | text alignment */
    var bottomRow = win.add("group");
    setupGroup(bottomRow, "row", COLUMN_SPACING);
    bottomRow.alignChildren = ["fill", "top"];
    controls.align = buildAlignPanel(bottomRow, updatePreview);
    controls.justify = buildJustifyPanel(bottomRow, updatePreview);

    /* 状況表示 / Status line */
    var statusText = win.add("statictext", undefined, L("status.ready"));
    statusText.alignment = "fill";

    /* ボタン / Buttons */
    var btnGroup = win.add("group");
    btnGroup.alignment = "right";
    var applyBtn = btnGroup.add("button", undefined, L("button.apply"));
    applyBtn.helpTip = L("tooltip.apply");

    applyBtn.onClick = commitPreview;

    /* 閉じる時：未確定のプレビューは取り消す（×・Esc 共通）/ On close: revert an uncommitted preview (X and Esc) */
    win.onClose = function () {
        if (previewState.active) {
            revertLastPreview();
            previewState.active = false;
        }
        return true;
    };

    /* キー操作：Option+A で確定、T/B で固定対象を選択、Esc で閉じる / Keys: Option+A commits, T/B pick the anchor, Esc closes */
    addApplyKeyHandler(win, commitPreview);
    addAnchorKeyHandler(win, controls.anchor.anchorTopRadio, controls.anchor.anchorBottomRadio, updatePreview);
    addCloseKeyHandler(win);

    win.center();
    win.show();
    updatePreview(); /* 初期プレビュー / Initial preview */
    return win;
}

paletteWindow = showPalette();
