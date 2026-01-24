#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * このスクリプトについて / About this script
 *
 * 選択したテキストフレームの見た目（位置/幅/行数）に合わせて、罫線・背景を自動生成します。
 * テキスト（文字/タブ/スタイル/タブストップ）には一切手を加えません。
 *
 * 生成モード / Modes
 * - 外枠は長方形：テキスト全体を囲む長方形の外枠＋横罫線（行区切り）を作成
 * - 外枠なし：横罫線（行区切り）のみを作成
 * - 行ごとに長方形：各行に背景（塗り）を作成（罫線は作りません）
 *
 * オプション / Options
 * - 縦罫：既存のタブストップ位置を参照して縦罫線を作成（テキストは変更しません）
 * - 見出し：指定モードで一部の線幅/背景を強調（モードにより有効/無効が変わります）
 */

/*
 * スクリプト名：TableMaker / Script name: TableMaker
 *
 * スクリプトの概要：選択中のテキストフレームの見た目（位置/幅/行数）に基づいて、
 *               外枠や横罫線、または行ごとの背景（塗り）を生成します。
 *               ※テキスト（タブ/文字/スタイル）には一切手を加えません。
 * 作成日：2026-01-24
 * 更新日：2026-01-24（v1.0） / Updated: 2026-01-24 (v1.0)
 * 最終更新日：2026-01-24 / Last updated: 2026-01-24
 * 行ごとに長方形を作成
 * 罫線ではなく「塗り」で作る（線はなし）
 * 上から奇数行はK10、偶数行はK30
 */

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "TableMaker" + ' ' + SCRIPT_VERSION,
        en: "TableMaker" + ' ' + SCRIPT_VERSION
    },
    strokeWidth: {
        ja: "線幅：",
        en: "Stroke width:"
    },
    shapePanel: {
        ja: "形状",
        en: "Shape"
    },
    shapeRect: {
        ja: "外枠は長方形",
        en: "Rectangle border"
    },
    shapeTopBottom: {
        ja: "外枠なし",
        en: "No outer border"
    },
    shapeRowRect: {
        ja: "行ごとに長方形",
        en: "Row rectangles"
    },
    verticalLines: {
        ja: "縦罫",
        en: "Vertical rules"
    },
    heading: {
        ja: "見出し",
        en: "Heading"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    }
};

function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    if (lang === 'ja' && entry.ja) return entry.ja;
    if (entry.en) return entry.en;
    if (entry.ja) return entry.ja;
    return key;
}

// ダイアログの表示位置と透明度 / Dialog position & opacity
var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY = 0.98;

/**
 * ダイアログの表示位置をずらす / Shift dialog position
 * @param {Window} dlg
 * @param {Number} offsetX
 * @param {Number} offsetY
 */
function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function () {
        var currentX = dlg.location[0];
        var currentY = dlg.location[1];
        dlg.location = [currentX + offsetX, currentY + offsetY];
    };
}

/**
 * ダイアログの透明度を設定 / Set dialog opacity
 * @param {Window} dlg
 * @param {Number} opacityValue
 */
function setDialogOpacity(dlg, opacityValue) {
    try {
        dlg.opacity = opacityValue;
    } catch (_) {
        // opacity をサポートしない環境では無視
    }
}

function createRowRectangles(textBounds, textWidth, halfGap, leading, paragraphCount, targetLayer, strokeWidthPt) {
    var leftX = textBounds[0] - halfGap;
    var width = textWidth + (halfGap * 2);

    var row;
    for (row = 0; row < paragraphCount; row++) {
        var topY = textBounds[1] + halfGap - (leading * row);
        var rect = targetLayer.pathItems.rectangle(topY, leftX, width, leading);

        // 線は使わず、塗りのみ
        rect.stroked = false;
        rect.filled = true;

        // row は 0始まりなので、row=0 が1行目（奇数行）
        var k = (row % 2 === 0) ? 10 : 30;
        rect.fillColor = createCmykColor(0, 0, 0, k);

        // 生成物は選択状態にしておく（元コード互換）
        try { rect.selected = true; } catch (_) { }

        // 生成順を末尾へ
        try { rect.move(rect.layer, ElementPlacement.PLACEATEND); } catch (_) { }
    }
}

function main() {
    try {
        // ドキュメントがない場合は静かに終了
        if (!app.documents.length) return;

        var doc = app.activeDocument;

        // 選択がない場合は静かに終了
        if (!doc.selection || doc.selection.length === 0) return;

        // 大量選択は処理を中断（例：1000個以上）
        if (doc.selection.length >= 1000) return;

        var selectionInfo = collectSelectionItems(doc.selection);
        if (selectionInfo.textFrames.length === 0) return;

        var rulerTypeIndex = app.preferences.getIntegerPreference("rulerType");
        var cursorKeyLengthPt = app.preferences.getRealPreference("cursorKeyLength");

        var unitNames = getRulerUnitNames();
        var displayUnit = unitNames[rulerTypeIndex];

        var defaultDisplayValue = convertUnit(cursorKeyLengthPt, "pt", displayUnit);
        var settings = showRuleSettingsDialog(defaultDisplayValue, displayUnit);
        if (settings === null) return;

        var strokeWidthPt = settings.strokeWidthPt;
        var shapeMode = settings.shapeMode;
        var verticalLines = settings.verticalLines;
        var heading = settings.heading;

        // 選択に含まれるパス（既存の線など）を削除
        removePathItems(selectionInfo.pathItems);

        // テキストごとに処理
        var i;
        for (i = 0; i < selectionInfo.textFrames.length; i++) {
            buildRulesForTextFrame(selectionInfo.textFrames[i], doc.activeLayer, strokeWidthPt, shapeMode, verticalLines, heading);
        }

    } catch (e) {
        // 仕様：不要な場面で alert は出さない
        try {
            $.writeln("エラー: " + e);
        } catch (_) { }
        return;
    }
}

/**
 * 選択から TextFrame と PathItem を抽出
 * @param {Array} selection
 * @returns {{textFrames:Array, pathItems:Array}}
 */
function collectSelectionItems(selection) {
    var result = { textFrames: [], pathItems: [] };
    var i;

    for (i = 0; i < selection.length; i++) {
        if (!selection[i]) continue;
        if (selection[i].typename === "TextFrame") result.textFrames.push(selection[i]);
        if (selection[i].typename === "PathItem") result.pathItems.push(selection[i]);
    }

    return result;
}

/**
 * PathItem を削除
 * @param {Array} pathItems
 */
function removePathItems(pathItems) {
    var i;
    for (i = 0; i < pathItems.length; i++) {
        try {
            pathItems[i].remove();
        } catch (_) { }
    }
}

/**
 * ルーラー単位の表示名（rulerType のインデックス対応）
 * 0:inch, 1:mm, 2:pt, 3:pica, 4:cm, 5:???（元コード互換で "H" を維持）, 6:px
 */
function getRulerUnitNames() {
    return ["inch", "mm", "pt", "pica", "cm", "H", "px"];
}

/**
 * 単位変換（小数第3位まで丸め）
 * @param {Number} value
 * @param {String} fromUnit
 * @param {String} toUnit
 * @returns {Number}
 */
function convertUnit(value, fromUnit, toUnit) {
    var uv = new UnitValue(value, fromUnit);
    var num = uv.as(toUnit);
    return Math.round(num * 1000) / 1000;
}

/**
 * ↑↓キーで値を増減（Shift/Option対応） / Change value by arrow keys (with Shift/Option)
 * - ↑↓: ±1
 * - Shift+↑↓: ±10（10の倍数へスナップ）
 * - Option+↑↓: ±0.1
 * @param {EditText} editText
 */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function (event) {
        if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;

            // Shiftキー押下時は10の倍数にスナップ / Snap to tens when Shift is held
            if (event.keyName === "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
            } else {
                value = Math.floor((value - 1) / delta) * delta;
            }

        } else if (keyboard.altKey) {
            delta = 0.1;

            if (event.keyName === "Up") {
                value += delta;
            } else {
                value -= delta;
            }

        } else {
            delta = 1;

            if (event.keyName === "Up") {
                value += delta;
            } else {
                value -= delta;
            }
        }

        // 下限を0に / Clamp to 0
        if (value < 0) value = 0;

        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
        } else {
            value = Math.round(value); /* 整数に丸め / Round to integer */
        }

        editText.text = String(value);
        event.preventDefault();
    });
}

/**
 * 罫線設定ダイアログを表示
 * @param {Number} defaultValue 表示単位での初期値
 * @param {String} displayUnit 表示単位文字列
 * @returns {{strokeWidthPt:Number, shapeMode:String, verticalLines:Boolean, heading:Boolean}|null} キャンセル時は null
 */
function showRuleSettingsDialog(defaultValue, displayUnit) {
    var dlg = new Window('dialog', L('dialogTitle'));

    // ダイアログの透明度と位置 / Dialog opacity & position
    setDialogOpacity(dlg, DIALOG_OPACITY);
    shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    /* 線幅 / Stroke width */
    var widthGroup = dlg.add('group');
    widthGroup.orientation = 'row';
    widthGroup.alignChildren = ['left', 'center'];

    widthGroup.add('statictext', undefined, L('strokeWidth'));
    var editText = widthGroup.add('edittext', undefined, String(defaultValue));
    editText.characters = 5;

    // ↑↓キーで値を増減 / Change value by arrow keys
    changeValueByArrowKey(editText);
    var unitLabel = widthGroup.add('statictext', undefined, displayUnit);

    /* 形状 / Shape */
    var shapePanel = dlg.add('panel', undefined, L('shapePanel'));
    shapePanel.orientation = 'column';
    shapePanel.alignChildren = ['left', 'top'];
    shapePanel.margins = [15, 20, 15, 10];

    var rbRect = shapePanel.add('radiobutton', undefined, L('shapeRect'));
    var rbTopBottom = shapePanel.add('radiobutton', undefined, L('shapeTopBottom'));
    var rbRowRect = shapePanel.add('radiobutton', undefined, L('shapeRowRect'));
    rbTopBottom.value = true;


    /* オプション / Options */
    var optionGroup = dlg.add('group');
    optionGroup.orientation = 'row';
    // ダイアログの左右中央に配置
    optionGroup.alignment = 'center';
    // グループ内の要素も中央揃え
    optionGroup.alignChildren = ['center', 'center'];

    var cbVerticalLines = optionGroup.add('checkbox', undefined, L('verticalLines'));
    cbVerticalLines.value = false;

    var cbHeading = optionGroup.add('checkbox', undefined, L('heading'));
    cbHeading.value = true;

    function updateLineWidthEnabled() {
        // 行ごとに長方形のときは線幅を使わないためディム表示
        var enabled = !rbRowRect.value;
        editText.enabled = enabled;
        unitLabel.enabled = enabled;

        // 行ごとに長方形が選ばれた場合は縦罫OFFかつディム
        if (rbRowRect.value) {
            cbVerticalLines.value = false;
            cbVerticalLines.enabled = false;
        } else {
            cbVerticalLines.enabled = true;
        }

        // 「見出し」は全モードで使用可能（外枠は長方形/外枠なし/行ごとに長方形）
        cbHeading.enabled = true;
    }

    rbRect.onClick = updateLineWidthEnabled;
    rbTopBottom.onClick = updateLineWidthEnabled;
    rbRowRect.onClick = updateLineWidthEnabled;
    updateLineWidthEnabled();

    /* ボタン / Buttons */
    var buttonGroup = dlg.add('group');
    buttonGroup.orientation = 'row';
    buttonGroup.alignment = 'right';

    buttonGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
    var okBtn = buttonGroup.add('button', undefined, L('ok'), { name: 'ok' });
    okBtn.active = true; // デフォルトボタンは右側（OK）

    if (dlg.show() !== 1) {
        return null;
    }

    var value = Number(editText.text);
    if (isNaN(value)) return null;

    var mode = 'rectangle';
    if (rbTopBottom.value) mode = 'topBottom';
    if (rbRowRect.value) mode = 'rowRectangles';

    return {
        strokeWidthPt: convertUnit(value, displayUnit, 'pt'),
        shapeMode: mode,
        verticalLines: cbVerticalLines.value,
        heading: cbHeading.value
    };
}

/**
 * テキストフレームの見た目に基づいて描画
 * ※テキストには一切手を加えない
 * @param {TextFrame} textFrame
 * @param {Layer} targetLayer
 * @param {Number} strokeWidthPt
 * @param {String} shapeMode 'rectangle' | 'topBottom' | 'rowRectangles'
 * @param {Boolean} verticalLines 縦罫線を描くかどうか
 * @param {Boolean} heading 見出しの強調を行うかどうか
 */
function buildRulesForTextFrame(textFrame, targetLayer, strokeWidthPt, shapeMode, verticalLines, heading) {
    try { textFrame.selected = false; } catch (_) { }

    var bounds = textFrame.geometricBounds; // [left, top, right, bottom]
    var paragraphs = textFrame.paragraphs;
    if (!paragraphs || paragraphs.length === 0) return;

    var fontSize = paragraphs[0].size;
    var leading = paragraphs[0].leading;
    var halfGap = (leading - fontSize) / 2;
    var totalHeight = leading * paragraphs.length;

    var width = textFrame.width;

    if (shapeMode === 'rowRectangles') {
        createRowFills(bounds, width, halfGap, leading, paragraphs.length, targetLayer, heading);
        // 行ごとに長方形の場合は縦罫は描かない
        return;
    }

    if (shapeMode === 'topBottom') {
        // 外枠なし：上下の罫線のみ作成（外枠の長方形は作らない）
        createOuterTopBottom(bounds, width, halfGap, totalHeight, targetLayer, strokeWidthPt, heading);
    } else {
        // 外枠は長方形：外枠（長方形）を作成
        var outerW = strokeWidthPt;
        if (heading === true && shapeMode === 'rectangle') outerW = strokeWidthPt * 3;
        createOuterRectangle(bounds, width, halfGap, totalHeight, targetLayer, outerW);
    }

    // 横線（行区切り）を生成（2行目以降）
    createHorizontalLines(bounds, halfGap, leading, paragraphs.length, targetLayer, strokeWidthPt, shapeMode, heading);

    // 縦罫線（参照のみ）：既存のタブストップ位置を使って縦線を引く（テキストは変更しない）
    if (verticalLines === true && shapeMode !== 'rowRectangles') {
        var tabPositions = getTabPositionsFromParagraphTabStops(paragraphs);
        if (tabPositions && tabPositions.length > 0) {
            createVerticalLines(bounds, tabPositions, halfGap, totalHeight, targetLayer, strokeWidthPt);
        }
    }
}

/**
 * 外枠（長方形）を作成
 */
function createOuterRectangle(textBounds, textWidth, halfGap, totalHeight, targetLayer, strokeWidthPt) {
    var topY = textBounds[1] + halfGap;
    var leftX = textBounds[0] - halfGap;

    var width = textWidth + (halfGap * 2);
    var height = totalHeight;

    var rect = targetLayer.pathItems.rectangle(topY, leftX, width, height);
    applyStrokeAttributes(rect, strokeWidthPt);
}

/**
 * 外側を「上下のみ」で作成
 */
function createOuterTopBottom(textBounds, textWidth, halfGap, totalHeight, targetLayer, strokeWidthPt, heading) {
    var leftX = textBounds[0] - halfGap;
    var rightX = leftX + (textWidth + (halfGap * 2));

    var topY = textBounds[1] + halfGap;
    var bottomY = topY - totalHeight;

    var w = (heading === true) ? (strokeWidthPt * 3) : strokeWidthPt;

    // 上
    createLine(targetLayer, [leftX, topY], [rightX, topY], w);

    // 下
    createLine(targetLayer, [leftX, bottomY], [rightX, bottomY], w);
}

/**
 * 行ごとの背景（塗り）を作成
 * 上から奇数行はK10、偶数行はK30
 * @param {Array} textBounds
 * @param {Number} textWidth
 * @param {Number} halfGap
 * @param {Number} leading
 * @param {Number} rowCount
 * @param {Layer} targetLayer
 * @param {Boolean} heading
 */
function createRowFills(textBounds, textWidth, halfGap, leading, rowCount, targetLayer, heading) {
    var leftX = textBounds[0] - halfGap;
    var width = textWidth + (halfGap * 2);

    var row;
    for (row = 0; row < rowCount; row++) {
        var topY = textBounds[1] + halfGap - (leading * row);
        var rect = targetLayer.pathItems.rectangle(topY, leftX, width, leading);

        rect.stroked = false;
        rect.filled = true;

        var k;
        // 「見出し」ON の場合：1つ目の長方形のみ K50
        if (heading === true && row === 0) {
            k = 50;
        } else {
            // 1行目（row=0）を K30、2行目（row=1）を K10 にする
            k = (row % 2 === 0) ? 30 : 10;
        }
        rect.fillColor = createCmykColor(0, 0, 0, k);

        try { rect.selected = true; } catch (_) { }
        try { rect.move(rect.layer, ElementPlacement.PLACEATEND); } catch (_) { }
    }
}

/**
 * 横罫線を作成（2行目以降）
 * @param {Array} textBounds
 * @param {Number} halfGap
 * @param {Number} leading
 * @param {Number} rowCount
 * @param {Layer} targetLayer
 * @param {Number} strokeWidthPt
 * @param {String} shapeMode
 * @param {Boolean} heading
 */
function createHorizontalLines(textBounds, halfGap, leading, rowCount, targetLayer, strokeWidthPt, shapeMode, heading) {
    var row;
    for (row = 1; row < rowCount; row++) {
        var y = textBounds[1] + halfGap - (leading * row);
        var leftX = textBounds[0] - halfGap;
        var rightX = textBounds[2] + halfGap;

        var w = strokeWidthPt;

        // 「外枠なし」＋「見出し」ON：
        // - 1本目（上罫）と最終（下罫）は createOuterTopBottom() 側で太くする
        // - 2本目（最初の行区切り横罫＝row=1）だけここで太くする
        if (heading === true && shapeMode === 'topBottom') {
            if (row === 1) {
                w = strokeWidthPt * 3;
            }
        }

        // 「外枠は長方形」＋「見出し」ON：上から2本目（最初の行区切り横罫＝row=1）を太くする
        if (heading === true && shapeMode === 'rectangle') {
            if (row === 1) {
                w = strokeWidthPt * 3;
            }
        }

        createLine(targetLayer, [leftX, y], [rightX, y], w);
    }
}

/**
 * 線分 PathItem を作成
 */
function createLine(targetLayer, startPoint, endPoint, strokeWidthPt) {
    var line = targetLayer.pathItems.add();
    line.setEntirePath([startPoint, endPoint]);
    applyStrokeAttributes(line, strokeWidthPt);
}

/**
 * 線・枠の共通属性を適用
 */
function applyStrokeAttributes(item, strokeWidthPt) {
    item.filled = false;
    item.stroked = true;
    item.strokeWidth = strokeWidthPt;
    item.strokeColor = createCmykColor(0, 0, 0, 100);

    try { item.selected = true; } catch (_) { }
    try { item.move(item.layer, ElementPlacement.PLACEATEND); } catch (_) { }
}

/**
 * CMYKColor を生成
 */
function createCmykColor(c, m, y, k) {
    var col = new CMYKColor();
    col.cyan = c;
    col.magenta = m;
    col.yellow = y;
    col.black = k;
    return col;
}

/**
 * アウトラインの子孫 PathItem を再帰的に取得し平坦化して配列で返す
 * @param {PageItem} parent
 * @returns {Array} PathItem の配列
 */
function collectOutlineItemsFlat(parent) {
    var items = [];
    collectOutlineItemsFlatRecursive(parent, items);
    return items;
}

function collectOutlineItemsFlatRecursive(parent, outItems) {
    if (!parent) return;

    // createOutline() は CompoundPathItem を多用するため、ここで拾う
    if (parent.typename === "PathItem" || parent.typename === "CompoundPathItem") {
        outItems.push(parent);
        return;
    }

    if (parent.pageItems && parent.pageItems.length > 0) {
        var i;
        for (i = 0; i < parent.pageItems.length; i++) {
            collectOutlineItemsFlatRecursive(parent.pageItems[i], outItems);
        }
    }
}

/**
 * 読み順（左から右、上から下）でアウトラインアイテムをソート
 * @param {Array} items PathItem 配列
 * @returns {Array} ソート済み配列
 */
function sortOutlineItemsReadingOrder(items) {
    var copied = items.slice(0);

    copied.sort(function (a, b) {
        var aBounds = a.geometricBounds; // [left, top, right, bottom]
        var bBounds = b.geometricBounds;

        // 上→下（topの降順）、同一行は左→右（leftの昇順）
        if (aBounds[1] !== bBounds[1]) return bBounds[1] - aBounds[1];
        return aBounds[0] - bBounds[0];
    });

    return copied;
}

/**
 * 既存のタブストップ位置を参照して、縦罫用の位置（左端からの距離）を取得
 * ※テキストは変更しない（参照のみ）
 * @param {Object} paragraphs textFrame.paragraphs
 * @returns {Array|null} タブ位置配列（positionの配列）。取得できない場合は null
 */
function getTabPositionsFromParagraphTabStops(paragraphs) {
    if (!paragraphs || paragraphs.length === 0) return null;

    // 行ごとにタブストップが異なる可能性があるため、列ごとに最大値を取る
    var maxPositions = [];
    var found = false;

    var p;
    for (p = 0; p < paragraphs.length; p++) {
        var ts = null;
        try { ts = paragraphs[p].tabStops; } catch (_) { ts = null; }

        if (!ts || ts.length === 0) continue;
        found = true;

        var i;
        for (i = 0; i < ts.length; i++) {
            var pos = null;
            try { pos = ts[i].position; } catch (_) { pos = null; }
            if (pos === null || pos === undefined) continue;
            if (pos <= 0) continue;

            if (i >= maxPositions.length) {
                maxPositions[i] = pos;
            } else {
                if (pos > maxPositions[i]) maxPositions[i] = pos;
            }
        }
    }

    if (!found) return null;

    // 末尾の未定義を詰める
    var out = [];
    var j;
    for (j = 0; j < maxPositions.length; j++) {
        if (maxPositions[j] !== undefined && maxPositions[j] !== null) out.push(maxPositions[j]);
    }

    if (out.length === 0) return null;
    return out;
}

/**
 * アウトラインと本文の文字列（タブ/改行）を突き合わせてタブ位置（累積幅）を推定
 * ※テキストには一切手を加えない
 * @param {Array} textBounds テキストフレームのジオメトリックバウンズ [left, top, right, bottom]
 * @param {Object} paragraphs textFrame.paragraphs
 * @param {Array} outlineItems 読み順ソート済みのアウトライン PathItem 配列
 * @param {Number} halfGap 行の上下余白の半分
 * @returns {Array} タブ位置（左端からの累積幅）配列
 */
function calculateTabPositions(textBounds, paragraphs, outlineItems, halfGap) {
    var result = [];
    var outlineIndex = 0;

    var paragraphIndex;
    for (paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
        var characters = paragraphs[paragraphIndex].characters;
        if (!characters || characters.length < 1) continue;

        var segmentStartX = textBounds[0];
        var tabIndex = 0;
        var accumulated = 0;

        var charIndex;
        for (charIndex = 0; charIndex < characters.length; charIndex++) {
            var ch = characters[charIndex].contents;

            // 改行はアウトラインが生成されない
            if (ch === "\r" || ch === "\n") {
                continue;
            }

            // タブはアウトラインが生成されない：直前文字のアウトライン右端で列幅を確定
            if (ch === "\t") {
                if (outlineIndex - 1 >= 0 && outlineIndex - 1 < outlineItems.length) {
                    var rightX = outlineItems[outlineIndex - 1].geometricBounds[2];
                    accumulated += (rightX - segmentStartX) + (halfGap * 2);

                    if (paragraphIndex === 0) {
                        // 1行目：そのまま登録
                        result.push(accumulated);
                    } else {
                        // 2行目以降：同列の最大値に更新
                        if (tabIndex < result.length) {
                            if (accumulated > result[tabIndex]) result[tabIndex] = accumulated;
                        }
                    }

                    // 次セグメント開始位置：タブ直後の最初のアウトライン左端
                    if (outlineIndex >= 0 && outlineIndex < outlineItems.length) {
                        segmentStartX = outlineItems[outlineIndex].geometricBounds[0];
                    }
                }

                tabIndex++;
                continue;
            }

            // 通常文字：アウトラインを1個消費
            if (outlineIndex >= outlineItems.length) {
                break;
            }

            outlineIndex++;
        }
    }

    return result;
}

/**
 * 縦罫線を描画（タブ位置＝左端からの累積幅）
 * @param {Array} textBounds テキストフレームのジオメトリックバウンズ [left, top, right, bottom]
 * @param {Array} tabPositions 左端からの累積幅の配列
 * @param {Number} halfGap
 * @param {Number} totalHeight
 * @param {Layer} targetLayer
 * @param {Number} strokeWidthPt
 */
function createVerticalLines(textBounds, tabPositions, halfGap, totalHeight, targetLayer, strokeWidthPt) {
    var leftX = textBounds[0] - halfGap;
    var topY = textBounds[1] + halfGap;
    var bottomY = topY - totalHeight;

    var i;
    for (i = 0; i < tabPositions.length; i++) {
        var x = leftX + tabPositions[i];
        createLine(targetLayer, [x, topY], [x, bottomY], strokeWidthPt);
    }
}

main();