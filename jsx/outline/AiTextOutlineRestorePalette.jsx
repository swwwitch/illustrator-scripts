#target illustrator
#targetengine "TextOutlineWithMemo"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name

TextOutlineRestorePalette.jsx

UI messages support Japanese/English. (Note format remains Japanese for compatibility.)

### 概要

- 常駐パレットから、選択テキストのアウトライン化（メモ付き）とその復元を行う
- 選択オブジェクトの note（メモ）をパネルに表示
- DOM 処理はすべてメインエンジンへ BridgeTalk 委譲

### Overview

- Outline selected text (saving its info to the note) and restore it back, from a persistent palette
- Show the selected object's note in the panel
- All DOM work is delegated to the main engine via BridgeTalk

### 主な機能 / Features

- アウトライン化（メモ付き）：各種プロパティを note に保存
- テキスト復元：note からテキストを再生成し、元アウトラインは outlined_text レイヤーへ退避
- 選択オブジェクトの note をパネルに表示（読み込みボタン）

### 保存・復元するプロパティ / Supported properties

- 文字列（本文・複数行対応）/ Text contents (multi-line)
- フォント（PostScript 名）/ Font (PostScript name)
- フォントサイズ / Font size
- 行送り / Leading
- カーニング方式（メトリクス／オプティカル／和文等幅／なし）/ Kerning method
- プロポーショナルメトリクス（カーニング方式に連動）/ Proportional metrics
- トラッキング / Tracking
- 文字ツメ（Tsume）/ Tsume
- 組み方向（縦組み／横組み）/ Orientation (vertical / horizontal)
- 文字カラー（CMYK／RGB／グレー／スポット）/ Fill color (CMYK / RGB / Gray / Spot)
- 位置（geometricBounds 基準）/ Position (by geometricBounds)

※ グラデーション／パターン塗り、混在属性は非対応（先頭文字の値を採用）
※ Gradient/pattern fills and mixed attributes are not supported (the first character's value is used)

### note

https://note.com/dtp_tranist/n/n3e0f241508db

### 更新履歴 / Change Log

- v1.0 (20240723) : 初期バージョン
- v1.9 (20260704) : 常駐パレット化＋テキスト復元・メモ表示・カーニング／文字カラー復元／複数行対応・堅牢化

*/

(function () {

// ==============================
// バージョン / Version
// ==============================
var SCRIPT_VERSION = "v1.9";

// ==============================
// ローカライズ / Localization
// ==============================
function getCurrentLang() {
    return (app.locale && app.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var CURRENT_LANG = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: 'アウトライン化（メモ付き）', en: 'Outline with Memo' }
    },
    panel: {
        selected: { ja: '先頭のメモ付きオブジェクト', en: 'First object with a note' },
        commands: { ja: '操作', en: 'Commands' }
    },
    button: {
        outline: { ja: 'アウトライン化（メモ付き）', en: 'Outline with Memo' },
        outlineTip: {
            ja: 'テキストを選択して実行（Esc で閉じる）',
            en: 'Select text and run (Esc to close)'
        },
        restore: { ja: 'テキストを復元', en: 'Restore Text' },
        restoreTip: {
            ja: 'アウトライン情報（note）からテキストを復元（Esc で閉じる）',
            en: 'Restore text from the outline note (Esc to close)'
        },
        load: { ja: 'メモを読み込み', en: 'Load Note' },
        loadTip: {
            ja: '選択オブジェクトの note を読み込んで表示',
            en: 'Load and show the selected object\'s note'
        }
    },
    memo: {
        empty: { ja: '（メモがありません）', en: '(No note)' }
    },
    listCol: {
        item: { ja: '項目', en: 'Item' },
        value: { ja: '値', en: 'Value' }
    },
    status: {
        ready: { ja: 'テキストまたはアウトラインを選択', en: 'Select text or an outline' },
        busy: { ja: '処理中…', en: 'Working…' },
        doneOutline: { ja: 'アウトライン化しました', en: 'Outlined' },
        doneRestore: { ja: 'テキストを復元しました', en: 'Text restored' },
        memoLoaded: { ja: 'メモを読み込みました', en: 'Note loaded' },
        fontWarn: { ja: '一部フォントは既定値を使用', en: 'Some fonts used defaults' },
        nodoc: { ja: 'ドキュメントがありません', en: 'No document is open' },
        nosel: { ja: 'オブジェクトが選択されていません', en: 'No objects are selected' },
        notgt: { ja: 'パス／グループを選択してください', en: 'Please select a path or group' },
        nonote: { ja: '有効なメモが見つかりません', en: 'No usable note found' },
        err: { ja: 'エラー', en: 'Error' }
    }
};

function L(path) {
    var parts = String(path).split('.');
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return path;
        node = node[parts[i]];
    }
    if (node == null) return path;
    if (node[CURRENT_LANG] != null) return node[CURRENT_LANG];
    if (node.en != null) return node.en;
    return path;
}

// ==============================
// ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing
// ==============================
var WINDOW_MARGINS = 16;
var WINDOW_SPACING = 12;
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

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

/* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// ==============================
// DOM 委譲用ワーカー関数 / Worker functions (run in main engine)
//   ※ 各 worker 関数の「本体内」は行コメント禁止／ブロックコメント /* */ のみ／必ずセミコロンで終える
//     （Function.toString() で送信する際に改行が落ちるため。関数外のこの見出しコメントは対象外）
// ==============================

/* --- アウトライン化 / Outline --- */
function workerRound(value) {
    return Math.round(value * 100) / 100;
}

/* --- note 文字列を組み立て / Build the memo text from gathered values --- */
function workerBuildMemoText(info) {
    return "文字列：\n" + info.content + "\n\n" +
        "フォント：\n" + info.fontName + "\n\n" +
        "フォントサイズ：\n" + info.fontSize + "\n\n" +
        "行送り：\n" + info.leading + "\n\n" +
        "カーニング：\n" + info.kerning + "\n\n" +
        "プロポーショナルメトリクス：\n" + info.proportionalMetrics + "\n\n" +
        "トラッキング：\n" + info.tracking + "\n\n" +
        "文字ツメ：\n" + info.tsume + "\n\n" +
        "組み方向：\n" + info.orientation + "\n\n" +
        "文字カラー：\n" + info.color + "\n\n" +
        "座標（geometricBounds）：\nL = " + info.left + ", T = " + info.top + ", R = " + info.right + ", B = " + info.bottom;
}

function workerProcessTextFrame(textFrame) {
    var textRange = textFrame.textRange;
    var kerningMethod = textRange.characterAttributes.kerningMethod;
    var kerningMethodText;
    if (kerningMethod == AutoKernType.AUTO) { kerningMethodText = "メトリクス"; }
    else if (kerningMethod == AutoKernType.METRICSROMANONLY) { kerningMethodText = "和文等幅"; }
    else if (kerningMethod == AutoKernType.OPTICAL) { kerningMethodText = "オプティカル"; }
    else { kerningMethodText = "なし"; }
    app.redraw();
    var bounds = textFrame.geometricBounds;
    var memoText = workerBuildMemoText({
        content: textRange.contents,
        fontName: textRange.characterAttributes.textFont.name,
        fontSize: workerRound(textRange.characterAttributes.size),
        leading: workerRound(textRange.characterAttributes.leading),
        kerning: kerningMethodText,
        proportionalMetrics: textRange.characterAttributes.proportionalMetrics ? "true" : "false",
        tracking: textRange.characterAttributes.tracking,
        tsume: textRange.characterAttributes.Tsume,
        orientation: (textFrame.orientation == TextOrientation.VERTICAL) ? "縦組み" : "横組み",
        color: workerColorToText(textRange.characterAttributes.fillColor),
        left: workerRound(bounds[0]),
        top: workerRound(bounds[1]),
        right: workerRound(bounds[2]),
        bottom: workerRound(bounds[3])
    });
    app.activeDocument.selection = null;
    textFrame.selected = true;
    textFrame.createOutline();
    /* createOutline 後の選択が常に期待どおりとは限らないので最低限チェック */
    if (app.activeDocument.selection && app.activeDocument.selection.length > 0) {
        app.activeDocument.selection[0].note = memoText;
    }
    return true;
}

function workerRunOutline() {
    if (app.documents.length < 1) { return "NODOC"; }
    var doc = app.activeDocument;
    var currentSelection = doc.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var selectedTextFrames = [];
    var loopIndex;
    for (loopIndex = 0; loopIndex < currentSelection.length; loopIndex++) {
        if (currentSelection[loopIndex].typename == "TextFrame") { selectedTextFrames.push(currentSelection[loopIndex]); }
    }
    if (selectedTextFrames.length < 1) { return "NOSEL"; }
    for (loopIndex = 0; loopIndex < selectedTextFrames.length; loopIndex++) {
        workerProcessTextFrame(selectedTextFrames[loopIndex]);
    }
    return "OK:" + selectedTextFrames.length;
}

/* --- note を表示用にコンパクト整形 / Format the note for compact display --- */
function workerFormatNoteForDisplay(noteText) {
    var noteLines = noteText.split("\n");
    var displayLabels = ["文字列", "フォント", "フォントサイズ", "行送り", "カーニング", "プロポーショナルメトリクス", "トラッキング", "文字ツメ", "組み方向", "文字カラー"];
    var displayLines = [];
    /* 本文（文字列）は複数行になりうるので "文字列：\n" 〜 "\n\nフォント：" を丸ごと取り出し、改行は ↵ で1行に畳む */
    var textStartMarker = "文字列：\n";
    var textEndMarker = "\n\nフォント：";
    var textStart = noteText.indexOf(textStartMarker);
    var textEnd = noteText.indexOf(textEndMarker);
    var bodyText = null;
    if (textStart >= 0 && textEnd > textStart) { bodyText = noteText.substring(textStart + textStartMarker.length, textEnd); }
    var labelIndex;
    for (labelIndex = 0; labelIndex < displayLabels.length; labelIndex++) {
        var currentLabel = displayLabels[labelIndex];
        if (currentLabel === "文字列" && bodyText != null) {
            displayLines.push("文字列： " + bodyText.replace(/[\r\n]/g, "↵"));
            continue;
        }
        var scanIndex;
        for (scanIndex = 0; scanIndex < noteLines.length; scanIndex++) {
            if (noteLines[scanIndex].indexOf(currentLabel + "：") === 0 && scanIndex + 1 < noteLines.length) {
                var displayValue = noteLines[scanIndex + 1];
                if (currentLabel === "文字ツメ") {
                    var tsumeNumber = parseFloat(displayValue);
                    if (!isNaN(tsumeNumber)) { displayValue = Math.round(tsumeNumber) + "%"; }
                }
                displayLines.push(currentLabel + "： " + displayValue);
                break;
            }
        }
    }
    return displayLines.join("\n");
}

/* --- 選択状態を検査（テキスト有無・メモ有無・表示用note） / Inspect selection state --- */
function workerInspectSelection() {
    if (app.documents.length < 1) { return "NODOC"; }
    var doc = app.activeDocument;
    var currentSelection = doc.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var noteHolder = null;
    var scanIndex;
    for (scanIndex = 0; scanIndex < currentSelection.length; scanIndex++) {
        var candidate = currentSelection[scanIndex];
        var candidateType = candidate.typename;
        /* 複数選択時は「先頭のメモ付きオブジェクト」を採用（メモを持つ最初の対象） */
        if (candidateType == "GroupItem" || candidateType == "PathItem" || candidateType == "TextFrame") {
            if (candidate.note && candidate.note.length > 0) { noteHolder = candidate; break; }
        }
    }
    if (!noteHolder) { return "NONOTE"; }
    var formattedNote = workerFormatNoteForDisplay(noteHolder.note);
    if (!formattedNote || formattedNote.length < 1) { formattedNote = noteHolder.note; }
    /* NOTE:<表示用note> / NONOTE の形式で返す */
    return "NOTE:" + formattedNote;
}

/* --- 復元：メモ解析 / Restore: parse note --- */
function workerExtractTextAttributes(noteText) {
    var attributes = {
        text: null,
        font: null,
        fontSize: null,
        leading: null,
        orientation: null,
        tracking: null,
        tsume: null,
        kerningText: null,
        colorText: null,
        proportionalMetrics: null,
        x: null,
        y: null,
        savedBounds: null
    };
    /* 本文（文字列）は複数行になりうるので "文字列：\n" 〜 "\n\nフォント：" を丸ごと本文にする（内部の改行・空行も保持） */
    var textStartMarker = "文字列：\n";
    var textEndMarker = "\n\nフォント：";
    var textStart = noteText.indexOf(textStartMarker);
    var textEnd = noteText.indexOf(textEndMarker);
    var restText = noteText;
    if (textStart >= 0 && textEnd > textStart) {
        attributes.text = noteText.substring(textStart + textStartMarker.length, textEnd);
        restText = noteText.substring(textEnd + 2);
    }
    /* フォント以降の単一行フィールドは、本文を除いた残りだけを走査（本文行の誤マッチ防止） */
    var noteLines = restText.split("\n");
    var lineIndex;
    for (lineIndex = 0; lineIndex < noteLines.length; lineIndex++) {
        if (attributes.text == null && noteLines[lineIndex].indexOf("文字列：") === 0 && lineIndex + 1 < noteLines.length) { attributes.text = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("フォント：") === 0 && lineIndex + 1 < noteLines.length) { attributes.font = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("フォントサイズ：") === 0 && lineIndex + 1 < noteLines.length) { attributes.fontSize = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("行送り：") === 0 && lineIndex + 1 < noteLines.length) { attributes.leading = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("組み方向：") === 0 && lineIndex + 1 < noteLines.length) { attributes.orientation = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("トラッキング：") === 0 && lineIndex + 1 < noteLines.length) { attributes.tracking = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("カーニング：") === 0 && lineIndex + 1 < noteLines.length) { attributes.kerningText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("文字ツメ：") === 0 && lineIndex + 1 < noteLines.length) { attributes.tsume = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("文字カラー：") === 0 && lineIndex + 1 < noteLines.length) { attributes.colorText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("プロポーショナルメトリクス：") === 0 && lineIndex + 1 < noteLines.length) { attributes.proportionalMetrics = noteLines[lineIndex + 1] === "true"; }
        if (noteLines[lineIndex].match(/^座標：\s*X\s*=\s*([-]?\d+(\.\d+)?),\s*Y\s*=\s*([-]?\d+(\.\d+)?)/)) {
            attributes.x = parseFloat(RegExp.$1);
            attributes.y = parseFloat(RegExp.$3);
        }
        if (noteLines[lineIndex].match(/L\s*=\s*([-]?\d+(?:\.\d+)?),\s*T\s*=\s*([-]?\d+(?:\.\d+)?),\s*R\s*=\s*([-]?\d+(?:\.\d+)?),\s*B\s*=\s*([-]?\d+(?:\.\d+)?)/)) {
            attributes.savedBounds = [parseFloat(RegExp.$1), parseFloat(RegExp.$2), parseFloat(RegExp.$3), parseFloat(RegExp.$4)];
        }
    }
    return attributes;
}

/* --- 復元：位置合わせ（geometricBounds ベース） / Restore: align by geometricBounds --- */
function workerAlignTextFrameByBounds(target, textFrame) {
    try {
        app.redraw();
        var targetBounds;
        if (target && target.length === 4 && typeof target[0] === 'number') { targetBounds = target; }
        else if (target && target.geometricBounds) { targetBounds = target.geometricBounds; }
        else { return; }
        var frameBounds = textFrame.geometricBounds;
        var dx = targetBounds[0] - frameBounds[0];
        var dy = targetBounds[1] - frameBounds[1];
        textFrame.translate(dx, dy);
    } catch (e) {
        /* 位置合わせに失敗しても処理は止めない / keep going on failure */
    }
}

/* --- 復元：レイヤーをロック解除＋表示（編集可能な状態に） / Make a layer usable --- */
function workerSetLayerUsable(layer) {
    try { layer.locked = false; } catch (eLock) {}
    try { layer.visible = true; } catch (eVis) {}
}

/* --- 復元：退避（アウトライン）レイヤーを毎回新規作成 / Restore: create a fresh stash layer --- */
function workerCreateOutlineStashLayer(doc) {
    if (!doc) { return null; }
    var stashLayer = doc.layers.add();
    workerSetLayerUsable(stashLayer);
    stashLayer.name = "__outlined_text_stash__";
    try {
        var markerGroup = stashLayer.groupItems.add();
        markerGroup.name = "__outlined_text_marker__";
    } catch (e2) {}
    return stashLayer;
}

/* --- 復元：restored_text レイヤーを用意（既存は統合） / Restore: get restored_text layer --- */
function workerCreateRestoredTextLayer(doc) {
    var baseName = "restored_text";
    var targetLayer = null;
    var findIndex;
    for (findIndex = 0; findIndex < doc.layers.length; findIndex++) {
        if (doc.layers[findIndex].name === baseName) { targetLayer = doc.layers[findIndex]; break; }
    }
    if (!targetLayer) { targetLayer = doc.layers.add(); targetLayer.name = baseName; }
    workerSetLayerUsable(targetLayer);
    /* 復元は常に同名 restored_text を再利用して集約するため、番号付きレイヤーの自動統合は行わない
       （ユーザーが手動で作った restored_text1 等を巻き込まない） */
    return targetLayer;
}

/* --- 復元：既存 outlined_text（過去の archive/dup 含む）を退避レイヤーに統合して1枚にまとめる
       Merge existing outlined_text layers (incl. old archive/dup) into the target so only one remains --- */
function workerMergeExistingOutlinedLayers(doc, targetLayer) {
    if (!doc || !targetLayer) { return; }
    var mergePattern = /^outlined_text(_archive\d+|_dup\d+)?$/;
    var mergeIndex;
    for (mergeIndex = doc.layers.length - 1; mergeIndex >= 0; mergeIndex--) {
        var mergeLayer = doc.layers[mergeIndex];
        if (mergeLayer === targetLayer) { continue; }
        if (!mergeLayer.name || !mergePattern.test(mergeLayer.name)) { continue; }
        workerSetLayerUsable(mergeLayer);
        try {
            while (mergeLayer.pageItems.length > 0) { mergeLayer.pageItems[0].move(targetLayer, ElementPlacement.PLACEATBEGINNING); }
        } catch (e1) {}
        try {
            while (mergeLayer.layers && mergeLayer.layers.length > 0) { mergeLayer.layers[0].remove(); }
        } catch (e2) {}
        try { mergeLayer.remove(); } catch (e3) {}
    }
}

/* --- 復元：outlined_text の重複名を解消 / Restore: dedupe outlined_text names --- */
function workerNormalizeOutlinedLayerNames(doc, keepLayer) {
    if (!doc || !keepLayer) { return; }
    var dupCounter = 1;
    var normalizeIndex;
    for (normalizeIndex = 0; normalizeIndex < doc.layers.length; normalizeIndex++) {
        var normalizeLayer = doc.layers[normalizeIndex];
        if (normalizeLayer !== keepLayer && normalizeLayer.name === "outlined_text") {
            workerSetLayerUsable(normalizeLayer);
            try { normalizeLayer.name = "outlined_text_dup" + dupCounter; } catch (e2) {}
            dupCounter++;
        }
    }
}

/* --- 復元：テンプレート適用後の後始末 / Restore: cleanup after template action --- */
function workerCleanupOutlinedDuplicateNames(doc) {
    if (!doc) { return null; }
    var keepLayer = null;
    var searchIndex;
    for (searchIndex = 0; searchIndex < doc.layers.length; searchIndex++) {
        var searchLayer = doc.layers[searchIndex];
        if (searchLayer.name !== "outlined_text") { continue; }
        try {
            var groupIndex;
            for (groupIndex = 0; groupIndex < searchLayer.groupItems.length; groupIndex++) {
                if (searchLayer.groupItems[groupIndex].name === "__outlined_text_marker__") { keepLayer = searchLayer; break; }
            }
        } catch (e0) {}
        if (keepLayer) { break; }
    }
    if (!keepLayer) {
        var bottomIndex;
        for (bottomIndex = doc.layers.length - 1; bottomIndex >= 0; bottomIndex--) {
            if (doc.layers[bottomIndex].name === "outlined_text") { keepLayer = doc.layers[bottomIndex]; break; }
        }
    }
    /* keep 以外の outlined_text をリネーム（normalize と同一処理を再利用） */
    workerNormalizeOutlinedLayerNames(doc, keepLayer);
    if (keepLayer) {
        workerSetLayerUsable(keepLayer);
        try {
            var removeIndex;
            for (removeIndex = keepLayer.groupItems.length - 1; removeIndex >= 0; removeIndex--) {
                if (keepLayer.groupItems[removeIndex].name === "__outlined_text_marker__") { keepLayer.groupItems[removeIndex].remove(); }
            }
        } catch (e4) {}
    }
    return keepLayer;
}

/* --- 復元：指定レイヤーを確実にアクティブ化 / Restore: force active layer --- */
function workerSetActiveLayerStrict(doc, layer) {
    if (!doc || !layer) { return false; }
    try { doc.activeLayer = layer; } catch (e0) {}
    try { if (doc.activeLayer === layer) { return true; } } catch (e1) {}
    try {
        var reFindIndex;
        for (reFindIndex = 0; reFindIndex < doc.layers.length; reFindIndex++) {
            if (doc.layers[reFindIndex] === layer) { doc.activeLayer = doc.layers[reFindIndex]; break; }
        }
    } catch (e2) {}
    try { return doc.activeLayer === layer; } catch (e3) { return false; }
}

/* --- 復元：テンプレートレイヤー属性をアクションで付与 / Restore: apply template-layer attribute via action --- */
function workerApplyTemplateLayerAttribute() {
    var actionString =
        '/version 3' +
        '/name [ 5' +
        ' 6c61796572' +
        ' ]' +
        '/isOpen 1' +
        '/actionCount 1' +
        '/action-1 {' +
        ' /name [ 24' +
        ' 6368616e67652d746f2d74656d706c6174652d6c61796572' +
        ' ]' +
        ' /keyIndex 0' +
        ' /colorIndex 0' +
        ' /isOpen 1' +
        ' /eventCount 1' +
        ' /event-1 {' +
        ' /useRulersIn1stQuadrant 0' +
        ' /internalName (ai_plugin_Layer)' +
        ' /localizedName [ 9' +
        ' e8a1a8e7a4ba203a20' +
        ' ]' +
        ' /isOpen 1' +
        ' /isOn 1' +
        ' /hasDialog 1' +
        ' /showDialog 0' +
        ' /parameterCount 10' +
        ' /parameter-1 {' +
        ' /key 1836411236' +
        ' /showInPalette 4294967295' +
        ' /type (integer)' +
        ' /value 4' +
        ' }' +
        ' /parameter-2 {' +
        ' /key 1851878757' +
        ' /showInPalette 4294967295' +
        ' /type (ustring)' +
        ' /value [ 36' +
        ' e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383' +
        ' a7e383b3' +
        ' ]' +
        ' }' +
        ' /parameter-3 {' +
        ' /key 1953068140' +
        ' /showInPalette 4294967295' +
        ' /type (ustring)' +
        ' /value [ 13' +
        ' 6f75746c696e65645f74657874' +
        ' ]' +
        ' }' +
        ' /parameter-4 {' +
        ' /key 1953329260' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-5 {' +
        ' /key 1936224119' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-6 {' +
        ' /key 1819239275' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-7 {' +
        ' /key 1886549623' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-8 {' +
        ' /key 1886547572' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 0' +
        ' }' +
        ' /parameter-9 {' +
        ' /key 1684630830' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-10 {' +
        ' /key 1885564532' +
        ' /showInPalette 4294967295' +
        ' /type (unit real)' +
        ' /value 50.0' +
        ' /unit 592474723' +
        ' }' +
        ' }' +
        '}';
    /* アクションはパレットを閉じるまで保持（初回のみ読み込む） */
    if (!$.global.__outlineTemplateActionLoaded) {
        var actionFile = new File('~/ScriptAction.aia');
        actionFile.encoding = 'UTF-8';
        actionFile.lineFeed = 'Unix';
        actionFile.open('w');
        actionFile.write(actionString);
        actionFile.close();
        app.loadAction(actionFile);
        actionFile.remove();
        $.global.__outlineTemplateActionLoaded = true;
    }
    app.doScript("change-to-template-layer", "layer", false);
    /* unloadAction はパレットを閉じるとき workerUnloadTemplateAction でまとめて実行 */
}

/* --- 復元：読み込んだテンプレートアクションを解放（パレットを閉じるとき） / Unload template action on palette close --- */
function workerUnloadTemplateAction() {
    if ($.global.__outlineTemplateActionLoaded) {
        try { app.unloadAction("layer", ""); } catch (e) {}
        $.global.__outlineTemplateActionLoaded = false;
    }
    return "OK";
}

/* --- 復元：テンプレートレイヤーを確定 / Restore: finalize template layer --- */
function workerFinalizeTemplateLayer(outlinedTextLayer) {
    if (!outlinedTextLayer) { return; }
    workerSetLayerUsable(outlinedTextLayer);
    workerMergeExistingOutlinedLayers(app.activeDocument, outlinedTextLayer);
    if (outlinedTextLayer.name !== "outlined_text") { outlinedTextLayer.name = "outlined_text"; }
    try { outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (eBack) {}
    workerNormalizeOutlinedLayerNames(app.activeDocument, outlinedTextLayer);
    try {
        var doc = app.activeDocument;
        var madeActive = workerSetActiveLayerStrict(doc, outlinedTextLayer);
        if (!madeActive) { return; }
        workerApplyTemplateLayerAttribute();
        outlinedTextLayer = workerCleanupOutlinedDuplicateNames(app.activeDocument) || outlinedTextLayer;
    } catch (eAct) {}
    try { outlinedTextLayer.locked = false; } catch (eUnlock) {}
    try {
        var markerIndex;
        for (markerIndex = outlinedTextLayer.groupItems.length - 1; markerIndex >= 0; markerIndex--) {
            if (outlinedTextLayer.groupItems[markerIndex].name === "__outlined_text_marker__") { outlinedTextLayer.groupItems[markerIndex].remove(); }
        }
    } catch (eMarker) {}
    try { outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (e9) {}
    try { outlinedTextLayer.locked = true; } catch (e10) {}
}

/* --- 塗り色を文字列へ（CMYK/RGB/Gray/Spot）/ Serialize a fill color to text --- */
/* ドキュメントカラーに応じて fillColor 型が CMYK/RGB になるので型を見て保存する */
function workerColorToText(color) {
    if (!color) { return ""; }
    var typeName = color.typename;
    if (typeName == "NoColor") { return "NONE"; }
    if (typeName == "CMYKColor") { return "CMYK " + workerRound(color.cyan) + " " + workerRound(color.magenta) + " " + workerRound(color.yellow) + " " + workerRound(color.black); }
    if (typeName == "RGBColor") { return "RGB " + Math.round(color.red) + " " + Math.round(color.green) + " " + Math.round(color.blue); }
    if (typeName == "GrayColor") { return "GRAY " + workerRound(color.gray); }
    if (typeName == "SpotColor") { return "SPOT " + color.spot.name; }
    /* グラデーション／パターン等は保存不可 → 空にして復元時は既定色のまま（不可視化を防ぐ） */
    return "";
}

/* --- 文字列から塗り色を再構築 / Rebuild a fill color from text --- */
function workerColorFromText(colorText) {
    if (!colorText) { return null; }
    if (colorText === "NONE") { return new NoColor(); }
    var parts = colorText.split(" ");
    var kind = parts[0];
    if (kind === "CMYK" && parts.length >= 5) {
        var cmyk = new CMYKColor();
        cmyk.cyan = parseFloat(parts[1]);
        cmyk.magenta = parseFloat(parts[2]);
        cmyk.yellow = parseFloat(parts[3]);
        cmyk.black = parseFloat(parts[4]);
        return cmyk;
    }
    if (kind === "RGB" && parts.length >= 4) {
        var rgb = new RGBColor();
        rgb.red = parseFloat(parts[1]);
        rgb.green = parseFloat(parts[2]);
        rgb.blue = parseFloat(parts[3]);
        return rgb;
    }
    if (kind === "GRAY" && parts.length >= 2) {
        var gray = new GrayColor();
        gray.gray = parseFloat(parts[1]);
        return gray;
    }
    if (kind === "SPOT") {
        /* スウォッチ名で既存スポットを再利用（名前に空白があるので prefix 以降を丸ごと名前に） */
        var spotName = colorText.substring(5);
        try {
            var spotColor = new SpotColor();
            spotColor.spot = app.activeDocument.spots.getByName(spotName);
            return spotColor;
        } catch (e) {
            return null;
        }
    }
    return null;
}

/* --- 復元：カーニング方式名を enum へ / Restore: resolve kerning label to AutoKernType --- */
/* AutoKerning.jsx の createAutoKernOptions と同じ対応（数値代入は不可・必ず enum） */
function workerResolveKernType(kerningText) {
    if (kerningText === "メトリクス") { return AutoKernType.AUTO; }
    if (kerningText === "オプティカル") { return AutoKernType.OPTICAL; }
    if (kerningText === "和文等幅") { return AutoKernType.METRICSROMANONLY; }
    return AutoKernType.NOAUTOKERN;
}

/* --- 復元：カーニング適用（AutoKerning.jsx の applyKerningToRanges と同一） / Apply kerning --- */
/* メトリクス（AUTO）のときだけ proportionalMetrics を ON、それ以外は OFF */
function workerApplyKerning(textRange, kerningMethod) {
    var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
    try {
        textRange.characterAttributes.kerningMethod = kerningMethod;
        textRange.characterAttributes.proportionalMetrics = useProportionalMetrics;
    } catch (e) {
        /* 適用できない範囲はスキップ / skip ranges that can't take these */
    }
}

/* --- 復元：テキストフレーム再生成 / Restore: recreate the text frame --- */
function workerCreateRestoredTextFrame(sourceItem, attributes, restoredLayer, restoreReport) {
    var targetLayer = restoredLayer || sourceItem.layer;
    var textFrame = targetLayer.textFrames.add();
    textFrame.contents = attributes.text;
    var posX = (attributes.x !== null) ? attributes.x : sourceItem.left;
    var posY = (attributes.y !== null) ? attributes.y : sourceItem.top;
    textFrame.position = [posX, posY];
    /* フォントだけ個別に：見つからなければ既定フォントのまま。他属性の適用は続行する */
    try {
        textFrame.textRange.characterAttributes.textFont = app.textFonts.getByName(attributes.font);
    } catch (eFont) {
        if (restoreReport) { restoreReport.fontFallback = true; }
    }
    /* 各属性は個別に適用（1つ失敗しても他の復元を止めない）。フォント取得の成否にも依存しない */
    var attrs = textFrame.textRange.characterAttributes;
    try { attrs.size = attributes.fontSize; } catch (eSize) {}
    try { attrs.autoLeading = false; attrs.leading = attributes.leading; } catch (eLead) {}
    try { attrs.tracking = attributes.tracking; } catch (eTrack) {}
    try { if (attributes.tsume !== null) { attrs.Tsume = attributes.tsume; } } catch (eTsume) {}
    try {
        if (attributes.kerningText != null) {
            /* カーニング方式を復元（AutoKerning ロジック：proportionalMetrics は方式に連動） */
            workerApplyKerning(textFrame.textRange, workerResolveKernType(attributes.kerningText));
        } else {
            /* 旧 note（カーニング未記録）は保存済みの proportionalMetrics を復元 */
            attrs.proportionalMetrics = attributes.proportionalMetrics;
        }
    } catch (eKern) {}
    try { textFrame.orientation = (attributes.orientation === "縦組み") ? TextOrientation.VERTICAL : TextOrientation.HORIZONTAL; } catch (eOri) {}
    try {
        if (attributes.colorText != null) {
            var restoredColor = workerColorFromText(attributes.colorText);
            if (restoredColor) { attrs.fillColor = restoredColor; }
        }
    } catch (eColor) {}
    return textFrame;
}

/* --- 復元：退避レイヤーへ移動（失敗時はレイヤー直下へフォールバック） / Restore: move source to stash layer --- */
function workerMoveToOutlinedLayer(sourceItem, outlinedTextLayer) {
    try {
        sourceItem.moveToBeginning(outlinedTextLayer.groupItems.add());
    } catch (e) {
        /* グループ生成／移動に失敗したらレイヤー直下へ退避 */
        try { sourceItem.moveToBeginning(outlinedTextLayer); } catch (e2) {}
    }
}

/* --- 復元：1オブジェクトを復元（Path / Group 共通） / Restore one item (path or group) --- */
function workerRestoreItem(sourceItem, outlinedTextLayer, restoredTextLayer, restoreReport) {
    var noteText = sourceItem.note;
    if (!noteText) { return false; }
    var attributes = workerExtractTextAttributes(noteText);
    if (!attributes || attributes.text == null) { return false; }
    var textFrame = workerCreateRestoredTextFrame(sourceItem, attributes, restoredTextLayer, restoreReport);
    workerAlignTextFrameByBounds(attributes.savedBounds || sourceItem, textFrame);
    /* 先に退避してから淡く＋ロック（ロック解除状態で移動する方が確実） */
    workerMoveToOutlinedLayer(sourceItem, outlinedTextLayer);
    /* 種別を問わず淡く＋ロック。失敗しても復元処理は止めない */
    try {
        sourceItem.opacity = 30;
        sourceItem.locked = true;
    } catch (eDim) {}
    if (restoreReport) { restoreReport.restored++; }
    return true;
}

/* --- 復元：エントリ / Restore: entry --- */
function workerRestoreText() {
    if (app.documents.length < 1) { return "NODOC"; }
    var doc = app.activeDocument;
    var currentSelection = doc.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var restorableItems = [];
    var pickIndex;
    for (pickIndex = 0; pickIndex < currentSelection.length; pickIndex++) {
        var pickType = currentSelection[pickIndex].typename;
        if (pickType === "GroupItem" || pickType === "PathItem") { restorableItems.push(currentSelection[pickIndex]); }
    }
    if (restorableItems.length < 1) { return "NOTGT"; }
    var outlinedTextLayer = workerCreateOutlineStashLayer(doc);
    var restoredTextLayer = workerCreateRestoredTextLayer(doc);
    var restoreReport = { restored: 0, fontFallback: false };
    /* ループ全体を try で囲み、想定外エラーでも空レイヤーを残さない */
    var runError = null;
    try {
        var restoreIndex;
        for (restoreIndex = 0; restoreIndex < restorableItems.length; restoreIndex++) {
            workerRestoreItem(restorableItems[restoreIndex], outlinedTextLayer, restoredTextLayer, restoreReport);
        }
    } catch (eRun) {
        runError = String(eRun);
    }
    if (restoreReport.restored < 1) {
        /* 1件も復元できなかったら（エラー時も含め）作成した退避／復元レイヤーを後始末 */
        try { outlinedTextLayer.remove(); } catch (eStash) {}
        try {
            if (restoredTextLayer.pageItems.length < 1 && (!restoredTextLayer.layers || restoredTextLayer.layers.length < 1)) { restoredTextLayer.remove(); }
        } catch (eRestored) {}
        return runError ? ("ERR:" + runError) : "NONOTE";
    }
    /* 一部でも成功していれば、その分を確定（途中エラーでも退避レイヤーを宙ぶらりんにしない） */
    workerFinalizeTemplateLayer(outlinedTextLayer);
    try { doc.activeLayer = restoredTextLayer; } catch (eActive) {}
    return "OK:" + restoreReport.restored + (restoreReport.fontFallback ? ":FONT" : "");
}

// ワーカー関数はすべてここに登録（追加漏れ防止）
var WORKER_FUNCS = [
    workerRound,
    workerSetLayerUsable,
    workerBuildMemoText,
    workerProcessTextFrame,
    workerRunOutline,
    workerFormatNoteForDisplay,
    workerInspectSelection,
    workerExtractTextAttributes,
    workerAlignTextFrameByBounds,
    workerCreateOutlineStashLayer,
    workerCreateRestoredTextLayer,
    workerMergeExistingOutlinedLayers,
    workerNormalizeOutlinedLayerNames,
    workerCleanupOutlinedDuplicateNames,
    workerSetActiveLayerStrict,
    workerApplyTemplateLayerAttribute,
    workerUnloadTemplateAction,
    workerFinalizeTemplateLayer,
    workerColorToText,
    workerColorFromText,
    workerResolveKernType,
    workerApplyKerning,
    workerCreateRestoredTextFrame,
    workerMoveToOutlinedLayer,
    workerRestoreItem,
    workerRestoreText
];

// ==============================
// BridgeTalk 委譲 / Delegation to main engine
// ==============================
function buildWorkerSource(funcs, entryCall) {
    var source = "";
    for (var i = 0; i < funcs.length; i++) {
        source += funcs[i].toString() + "\n";
    }
    return source + entryCall;
}

/* funcs 省略時は全 worker を送信。ポーリングなど軽量呼び出しは必要な関数だけ渡す */
function callWorker(entryCall, funcs) {
    var workerFuncs = funcs || WORKER_FUNCS;
    var resultHolder = { value: null };
    var bridge = new BridgeTalk();
    bridge.target = "illustrator";
    var code = buildWorkerSource(workerFuncs, entryCall);
    bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
    bridge.onResult = function (response) { resultHolder.value = String(response.body); };
    bridge.onError = function (errorResponse) { resultHolder.value = "ERR:" + String(errorResponse.body); };
    bridge.send(60); // 同期待ち上限（秒）。多数オブジェクトの復元やアクション実行に備えて長めに
    return resultHolder.value;
}

// ==============================
// パレット / Palette
// ==============================
// パレット参照は $.global に保持（IIFE をまたいで常駐させ GC・多重起動を防ぐ）
var isBusy = false;

function setStatus(win, message) {
    win.statusText.text = message;
    win.layout.layout(true);
}

function applyResultToStatus(win, result, doneKey) {
    if (result == null) { setStatus(win, L('status.err')); return; }
    if (result.indexOf("OK") === 0) {
        var parts = result.split(":");
        var count = (parts.length > 1) ? parts[1] : "";
        var message = L(doneKey) + (count ? " (" + count + ")" : "");
        if (result.indexOf("FONT") >= 0) { message += " / " + L('status.fontWarn'); }
        setStatus(win, message);
    } else if (result === "NODOC") { setStatus(win, L('status.nodoc'));
    } else if (result === "NOSEL") { setStatus(win, L('status.nosel'));
    } else if (result === "NOTGT") { setStatus(win, L('status.notgt'));
    } else if (result === "NONOTE") { setStatus(win, L('status.nonote'));
    } else if (result.indexOf("ERR") === 0) { setStatus(win, L('status.err') + ": " + result.substring(4));
    } else { setStatus(win, result); }
}

function onOutlineClick(win) {
    if (isBusy) return; // 再入防止
    isBusy = true;
    setStatus(win, L('status.busy'));
    var result = callWorker("workerRunOutline();");
    isBusy = false;
    applyResultToStatus(win, result, 'status.doneOutline');
    refreshSelectedNote(win, true); // アウトライン化直後の note を表示
}

function onRestoreClick(win) {
    if (isBusy) return;
    // 復元の直前に「読み込み」ロジックを強制実行し、対象のメモを表示（移動前の状態を見せる）
    refreshSelectedNote(win, true);
    isBusy = true;
    setStatus(win, L('status.busy'));
    var result = callWorker("workerRestoreText();");
    isBusy = false;
    applyResultToStatus(win, result, 'status.doneRestore');
}

function populateNoteList(win, formattedNote) {
    win.noteList.removeAll();
    var noteLines = formattedNote.split("\n");
    var lineIndex;
    for (lineIndex = 0; lineIndex < noteLines.length; lineIndex++) {
        var currentLine = noteLines[lineIndex];
        if (!currentLine) { continue; }
        var separatorPos = currentLine.indexOf("： ");
        var itemLabel, itemValue;
        if (separatorPos >= 0) {
            itemLabel = currentLine.substring(0, separatorPos);
            itemValue = currentLine.substring(separatorPos + 2);
        } else {
            itemLabel = currentLine;
            itemValue = "";
        }
        var row = win.noteList.add("item", itemLabel);
        row.subItems[0].text = itemValue;
    }
}

function refreshSelectedNote(win, keepStatus) {
    if (isBusy) return;
    isBusy = true;
    var result = callWorker("workerInspectSelection();");
    isBusy = false;
    win.noteList.removeAll();
    if (result != null && result.indexOf("NOTE:") === 0) {
        populateNoteList(win, result.substring(5));
        if (!keepStatus) { setStatus(win, L('status.memoLoaded')); }
    } else if (result === "NONOTE") {
        if (!keepStatus) { setStatus(win, L('status.nonote')); }
    } else if (result === "NOSEL") {
        if (!keepStatus) { setStatus(win, L('status.nosel')); }
    } else if (result === "NODOC") {
        if (!keepStatus) { setStatus(win, L('status.nodoc')); }
    } else if (!keepStatus) {
        setStatus(win, L('status.err'));
    }
}

/* パレットを閉じるときの後始末：読み込んだアクションを解放 */
function performCloseCleanup() {
    try { callWorker("workerUnloadTemplateAction();"); } catch (e) {}
    $.global.__textOutlineMemoPalette = null;
}

function showPalette() {
    // 多重起動防止：既存パレットがあれば閉じる（$.global で常駐参照して IIFE をまたいで保持）
    if ($.global.__textOutlineMemoPalette) {
        try { $.global.__textOutlineMemoPalette.close(); } catch (e) {}
        $.global.__textOutlineMemoPalette = null;
    }

    var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
    setupWindow(win);

    // 選択オブジェクトの note を表示 / Selected object's note
    var selectedObjectPanel = win.add("panel", undefined, L('panel.selected'));
    setupPanel(selectedObjectPanel, 6);

    win.noteList = selectedObjectPanel.add("listbox", undefined, [], {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: [L('listCol.item'), L('listCol.value')],
        columnWidths: [140, 170]
    });
    win.noteList.preferredSize = [320, 205]; // 9行分 / 9 rows

    var loadNoteButton = selectedObjectPanel.add("button", undefined, L('button.load'));
    loadNoteButton.helpTip = L('button.loadTip');
    loadNoteButton.alignment = "left"; // ボタンは幅いっぱいに広げない
    loadNoteButton.margins = [0, 6, 0, 0]; // ボタンの上に余白 +6
    loadNoteButton.onClick = function () { refreshSelectedNote(win, false); };

    // 操作ボタン / Command buttons
    var commandPanel = win.add("panel", undefined, L('panel.commands'));
    setupPanel(commandPanel, 6);

    var commandButtonRow = commandPanel.add("group");
    setupRow(commandButtonRow, "left", 8); // ボタン列は左寄せ（幅いっぱいに広げない）

    var outlineButton = commandButtonRow.add("button", undefined, L('button.outline'));
    outlineButton.helpTip = L('button.outlineTip');
    outlineButton.onClick = function () { onOutlineClick(win); };

    var restoreButton = commandButtonRow.add("button", undefined, L('button.restore'));
    restoreButton.helpTip = L('button.restoreTip');
    restoreButton.onClick = function () { onRestoreClick(win); };

    // ステータス表示 / Status
    win.statusText = win.add("statictext", undefined, L('status.ready'));
    win.statusText.alignment = "left";

    // Esc で閉じる
    win.addEventListener("keydown", function (ev) {
        if (ev.keyName == "Escape") {
            try { win.close(); } catch (e) {}
        }
    });

    // 閉じるとき（× / Esc）に読み込んだアクションを解放
    win.onClose = function () {
        performCloseCleanup();
        return true;
    };

    $.global.__textOutlineMemoPalette = win;
    refreshSelectedNote(win, true); // 起動時に選択オブジェクトの note を表示
    win.show();
    return win;
}

// ==============================
// 実行 / Run
// ==============================
showPalette();

})();
