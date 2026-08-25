#target illustrator
#targetengine "TextOutlineWithMemo"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストのアウトライン化と、アウトラインからのテキスト復元を常駐パレットから実行します。
アウトライン化の直前に文字・段落属性をオブジェクトのメモ（note）へ保存し、そのメモをもとにテキストフレームを再生成します。

詳細は README を参照してください。

### Overview

Outlines text and restores it back from the outlines, driven from a persistent palette.
Character and paragraph attributes are serialized into the object's note just before outlining, and the note is parsed to rebuild the text frame.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiTextOutlineRestorePalette";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-07-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiTextOutlineRestorePalette.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiTextOutlineRestorePalette.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc476be8ad43c"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

// ローカライズ / Localization
// ==============================
function getCurrentLang() {
    return (app.locale && app.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var CURRENT_LANG = getCurrentLang();
/* 英語環境では和文専用属性（組み方向・禁則・文字組みアキ量設定・文字ツメ）を保存／表示／復元しない */
var HANDLE_JP = (CURRENT_LANG === 'ja');

var LABELS = {
    dialog: {
        title: { ja: 'テキストのアウトライン化と復元', en: 'Outline & Restore Text' }
    },
    panel: {
        outline: { ja: 'アウトライン', en: 'Outline' },
        outlineTip: {
            ja: 'テキストをアウトライン化し、文字・段落属性をメモとして保存します',
            en: 'Outline text and save its character/paragraph attributes as a note'
        },
        selected: { ja: 'メモ付きオブジェクト', en: 'Object with a note' },
        selectedTip: {
            ja: '選択オブジェクトに保存されたメモ（属性）を一覧表示します',
            en: 'Lists the note (attributes) saved on the selected object'
        },
        restore: { ja: 'テキストを復元', en: 'Restore Text' },
        restoreTip: {
            ja: 'メモからテキストを復元します',
            en: 'Restore text from the note'
        }
    },
    button: {
        outline: { ja: 'アウトライン化（メモ付き）', en: 'Outline with Note' },
        outlineTip: {
            ja: 'テキストを選択して実行（Esc で閉じる）',
            en: 'Select text and run (Esc to close)'
        },
        restore: { ja: 'テキストを復元', en: 'Restore Text' },
        restoreTip: {
            ja: 'アウトライン情報（メモ）からテキストを復元（Esc で閉じる）',
            en: 'Restore text from the outline note (Esc to close)'
        },
        load: { ja: 'メモを読み込み', en: 'Load Note' },
        loadTip: {
            ja: '選択オブジェクトのメモを読み込んで表示',
            en: 'Load and show the selected object\'s note'
        },
        attributes: { ja: '属性パネル', en: 'Attributes' },
        attributesTip: {
            ja: '属性パネルを開く／閉じる（メモの確認・編集に使用）',
            en: 'Toggle the Attributes panel (used to view/edit the note)'
        }
    },
    option: {
        keepOutline: { ja: 'アウトラインデータを残す', en: 'Keep outline data' },
        keepOutlineTip: {
            ja: 'OFF にすると復元後にアウトラインを削除し、outlined_text レイヤーを作成しません',
            en: 'When off, outlines are deleted after restore and no outlined_text layer is created'
        },
        separateLayer: { ja: '復元したテキストを別レイヤーに', en: 'Restore text to a separate layer' },
        separateLayerTip: {
            ja: 'OFF にすると復元テキストをアウトライン情報と同じレイヤーに置き、restored_text レイヤーを作成しません',
            en: 'When off, restored text is placed on the same layer as the outline and no restored_text layer is created'
        }
    },
    memo: {
        empty: { ja: '（メモがありません）', en: '(No note)' }
    },
    listCol: {
        item: { ja: '項目', en: 'Item' },
        value: { ja: '値', en: 'Value' },
        hint: { ja: '選択オブジェクトのメモ（保存された属性）', en: 'The selected object\'s note (saved attributes)' }
    },
    status: {
        ready: { ja: 'テキストまたはアウトラインを選択', en: 'Select text or an outline' },
        busy: { ja: '処理中…', en: 'Working…' },
        doneOutline: { ja: 'アウトライン化しました', en: 'Outlined' },
        doneRestore: { ja: 'テキストを復元しました', en: 'Text restored' },
        memoLoaded: { ja: 'メモを読み込みました', en: 'Note loaded' },
        fontWarn: { ja: '一部フォントは既定値を使用', en: 'Some fonts used defaults' },
        partial: { ja: '一部は処理できませんでした', en: 'Some items could not be processed' },
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
// listbox 表示のローカライズ / Localize the listbox contents
//   note は互換性のため日本語で保存されるので、表示直前に項目名と列挙値だけ現在言語へ変換する
//   （数値・フォント名・カラー値・true/false はそのまま）
// ==============================

/* 項目名（左列）の英訳。日本語環境では note の表記をそのまま使うので en だけ持つ / Item labels (left column) */
var LIST_ITEM_LABELS_EN = {
    '文字列':                     'Text',
    '組み方向':                   'Orientation',
    'フォント':                   'Font',
    'フォントサイズ':             'Font size',
    '行送り':                     'Leading',
    '自動行送り':                 'Auto leading',
    '水平比率':                   'Horizontal scale',
    '垂直比率':                   'Vertical scale',
    'カーニング':                 'Kerning',
    'プロポーショナルメトリクス': 'Proportional metrics',
    'トラッキング':               'Tracking',
    '文字ツメ':                   'Tsume',
    '行揃え':                     'Alignment',
    '禁則':                       'Kinsoku',
    '文字組み':                   'Mojikumi',
    '文字カラー':                 'Fill color'
};

/* 列挙値（右列）の英訳。ここに無い値（数値・フォント名・カラー・true/false）は素通し / Enumerated values (right column) */
var LIST_VALUE_LABELS_EN = {
    '縦組み':                 'Vertical',
    '横組み':                 'Horizontal',
    'メトリクス':             'Metrics',
    'オプティカル':           'Optical',
    '和文等幅':               'Metrics (Roman Only)',
    'なし':                   'None',
    '左揃え':                 'Left',
    '中央揃え':               'Center',
    '右揃え':                 'Right',
    '均等配置':               'Justify all lines',
    '均等配置（最終行左）':   'Justify (last left)',
    '均等配置（最終行中央）': 'Justify (last center)',
    '均等配置（最終行右）':   'Justify (last right)',
    '強い禁則':               'Hard',
    '弱い禁則':               'Soft',
    '弱い禁則 v2':            'Soft v2',
    '行末約物全角/半角':      'Line end full/half-width',
    '約物半角':               'Half-width punctuation',
    '行末約物半角':           'Line end half-width',
    '行末約物全角':           'Line end full-width',
    '約物全角':               'Full-width punctuation',
    'ツメ組み':               'Tight',
    'ベタ組み':               'Solid'
};

/* 日本語表記を現在言語へ（未知＝数値・フォント名・カラー等はそのまま）/ Localize a Japanese label or value */
function localizeFromTable(table, jaText) {
    if (CURRENT_LANG === 'ja') { return jaText; }
    return (table[jaText] != null) ? table[jaText] : jaText;
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

/* --- 属性を1つだけ代入（null／代入不可なら黙って無視）/ Set one attribute, ignoring nulls and failures --- */
function workerSetAttr(target, key, value) {
    if (value == null) { return; }
    try { target[key] = value; } catch (e) {}
}

/* --- 属性パネルを開閉（メニューコマンド）/ Toggle the Attributes panel via menu command --- */
function workerToggleAttributesPanel() {
    try {
        app.executeMenuCommand("internal palettes posing as plug-in menus-attributes");
    } catch (e) {
    }
    return "OK";
}

/* --- note のフィールド定義（保存順）。組み立て・解析・一覧表示の共通ソース
       Note field table (in saved order): drives build, parse and list display --- */
/* type は body=複数行の本文／num=数値／bool=true,false／text=そのまま。jpOnly は英語環境で書き出さない属性 */
/* 本文は "文字列：\n" 〜 "\n\nフォント：" で切り出すので、文字列とフォントの並びは変更しないこと */
function workerNoteFields() {
    return [
        { label: "文字列",                     key: "text",                type: "body" },
        { label: "フォント",                   key: "font",                type: "text" },
        { label: "フォントサイズ",             key: "fontSize",            type: "num" },
        { label: "行送り",                     key: "leading",             type: "num" },
        { label: "カーニング",                 key: "kerningText",         type: "text" },
        { label: "プロポーショナルメトリクス", key: "proportionalMetrics", type: "bool" },
        { label: "トラッキング",               key: "tracking",            type: "num" },
        { label: "文字ツメ",                   key: "tsume",               type: "num",  jpOnly: true },
        { label: "組み方向",                   key: "orientation",         type: "text", jpOnly: true },
        { label: "文字カラー",                 key: "colorText",           type: "text" },
        { label: "行揃え",                     key: "justificationText",   type: "text" },
        { label: "禁則",                       key: "kinsokuText",         type: "text", jpOnly: true },
        { label: "文字組み",                   key: "mojikumiText",        type: "text", jpOnly: true },
        { label: "自動行送り",                 key: "autoLeading",         type: "bool" },
        { label: "水平比率",                   key: "horizontalScale",     type: "num" },
        { label: "垂直比率",                   key: "verticalScale",       type: "num" }
    ];
}

/* --- note 文字列を組み立て / Build the memo text from gathered values --- */
/* handleJp が false（英語環境）のときは和文専用属性（jpOnly）を書き出さない */
function workerBuildMemoText(info, handleJp) {
    var jp = (handleJp !== false);
    var fields = workerNoteFields();
    var memo = "";
    var i;
    for (i = 0; i < fields.length; i++) {
        if (fields[i].jpOnly && !jp) { continue; }
        memo += fields[i].label + "：\n" + info[fields[i].key] + "\n\n";
    }
    return memo + "座標：\nL = " + info.left + ", T = " + info.top + ", R = " + info.right + ", B = " + info.bottom;
}

/* --- カーニング方式（表示ラベル ↔ AutoKernType）/ Kerning method table --- */
/* AutoKerning.jsx の createAutoKernOptions と同じ対応（数値代入は不可・必ず enum） */
function workerKerningPairs() {
    return [
        { labelText: "メトリクス",   kernType: AutoKernType.AUTO },
        { labelText: "和文等幅",     kernType: AutoKernType.METRICSROMANONLY },
        { labelText: "オプティカル", kernType: AutoKernType.OPTICAL }
    ];
}

/* --- カーニング方式を保存用テキストへ / Serialize the kerning method to text --- */
function workerKerningToText(kerningMethod) {
    var pairs = workerKerningPairs();
    var i;
    for (i = 0; i < pairs.length; i++) {
        if (kerningMethod == pairs[i].kernType) { return pairs[i].labelText; }
    }
    return "なし";
}

/* --- 復元：カーニング方式名を enum へ / Restore: resolve kerning label to AutoKernType --- */
function workerResolveKernType(kerningText) {
    var pairs = workerKerningPairs();
    var i;
    for (i = 0; i < pairs.length; i++) {
        if (kerningText === pairs[i].labelText) { return pairs[i].kernType; }
    }
    return AutoKernType.NOAUTOKERN;
}

/* --- 1つのテキストフレームをアウトライン化し、生成されたグループに note を書く
       Outline one text frame and write the note onto the resulting group --- */
/* 戻り値は生成された GroupItem。作れなかったときは null（呼び出し側で失敗として数える） */
function workerProcessTextFrame(textFrame, handleJp) {
    var textRange = textFrame.textRange;
    var charAttrs = textRange.characterAttributes;
    var bounds = textFrame.geometricBounds;
    var memoText = workerBuildMemoText({
        text: textRange.contents,
        font: charAttrs.textFont.name,
        fontSize: workerRound(charAttrs.size),
        leading: workerRound(charAttrs.leading),
        kerningText: workerKerningToText(charAttrs.kerningMethod),
        proportionalMetrics: charAttrs.proportionalMetrics ? "true" : "false",
        tracking: charAttrs.tracking,
        tsume: charAttrs.Tsume,
        orientation: (textFrame.orientation == TextOrientation.VERTICAL) ? "縦組み" : "横組み",
        colorText: workerColorToText(charAttrs.fillColor),
        justificationText: workerJustificationToText(textRange.paragraphAttributes.justification),
        kinsokuText: workerKinsokuToText(textRange.paragraphAttributes),
        mojikumiText: workerMojikumiToText(textRange.paragraphAttributes),
        autoLeading: charAttrs.autoLeading ? "true" : "false",
        horizontalScale: workerRound(charAttrs.horizontalScale),
        verticalScale: workerRound(charAttrs.verticalScale),
        left: workerRound(bounds[0]),
        top: workerRound(bounds[1]),
        right: workerRound(bounds[2]),
        bottom: workerRound(bounds[3])
    }, handleJp);
    /* createOutline() の戻り値（生成された GroupItem）へ直接書く。選択に頼るとメモが付かないまま
       テキストだけ失われることがあるため */
    var outlineGroup = textFrame.createOutline();
    if (!outlineGroup) { return null; }
    outlineGroup.note = memoText;
    return outlineGroup;
}

/* --- 選択を取得（ドキュメントなし／選択なしはコード文字列を返す）/ Get the selection or an error code --- */
function workerGetSelection() {
    if (app.documents.length < 1) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    return currentSelection;
}

/* --- 選択から指定 typename のオブジェクトだけを取り出す / Filter a selection by typename --- */
function workerFilterByType(selection, typeNames) {
    var picked = [];
    var selectionIndex, typeIndex;
    for (selectionIndex = 0; selectionIndex < selection.length; selectionIndex++) {
        for (typeIndex = 0; typeIndex < typeNames.length; typeIndex++) {
            if (selection[selectionIndex].typename === typeNames[typeIndex]) { picked.push(selection[selectionIndex]); break; }
        }
    }
    return picked;
}

/* --- アウトライン化：エントリ / Outline: entry --- */
/* 1件失敗しても残りは処理し、生成したアウトラインを選択し直して直後のメモ表示につなげる */
function workerRunOutline(handleJp) {
    var currentSelection = workerGetSelection();
    if (typeof currentSelection === "string") { return currentSelection; }
    var selectedTextFrames = workerFilterByType(currentSelection, ["TextFrame"]);
    if (selectedTextFrames.length < 1) { return "NOSEL"; }
    /* geometricBounds を確定させるための再描画は1回だけ（件数分繰り返さない） */
    try { app.redraw(); } catch (eDraw) {}
    var outlineGroups = [];
    var errorText = null;
    var loopIndex;
    for (loopIndex = 0; loopIndex < selectedTextFrames.length; loopIndex++) {
        try {
            var outlineGroup = workerProcessTextFrame(selectedTextFrames[loopIndex], handleJp);
            if (outlineGroup) { outlineGroups.push(outlineGroup); }
            else if (errorText == null) { errorText = "createOutline failed"; }
        } catch (eItem) {
            if (errorText == null) { errorText = String(eItem); }
        }
    }
    if (outlineGroups.length < 1) { return errorText ? ("ERR:" + errorText) : "NOSEL"; }
    workerSetAttr(app.activeDocument, "selection", outlineGroups);
    return "OK:" + outlineGroups.length + (errorText ? ":PARTIAL" : "");
}

/* --- 表示用：禁則の内部名を日本語ラベルへ / Kinsoku internal name to Japanese label (display only) --- */
/* note には内部名（Soft/Hard/Soft_v2 等）を保存して復元はそれで行い、listbox 表示だけ日本語化する */
function workerKinsokuToDisplay(value) {
    var kinsokuPresets = [
        { kinsokuName: "None",    labelText: "なし" },
        { kinsokuName: "なし",    labelText: "なし" },
        { kinsokuName: "Hard",    labelText: "強い禁則" },
        { kinsokuName: "Soft",    labelText: "弱い禁則" },
        { kinsokuName: "Soft_v2", labelText: "弱い禁則 v2" }
    ];
    var i;
    for (i = 0; i < kinsokuPresets.length; i++) {
        if (kinsokuPresets[i].kinsokuName === value) { return kinsokuPresets[i].labelText; }
    }
    return value;
}

/* --- 表示用：文字組みアキ量設定を日本語ラベルへ / Mojikumi set name to Japanese label (display only) --- */
/* 新 note は日本語ラベルで保存されるので大半は素通し。旧 note の内部名（例: Gyomatsu Yakumono Zenkaku Hankaku）は
   (1) プリセット日本語ラベル一致 → (2) doc.mojikumiSet の位置で解決 → (3) 内部ローマ字名フォールバック の順で日本語化 */
function workerMojikumiToDisplay(value) {
    if (value == null || value === "" || value === "なし" || value === "None") { return "なし"; }
    var preset = workerFindMojikumiPreset(value);
    if (preset) { return preset.labelText; }
    var byCollection = workerMojikumiLabelFromApplied(value);
    if (byCollection != null) { return byCollection; }
    /* コレクションでも引けない環境向け：内部ローマ字名を前方一致（長い key を先に、zenkaku が zenkakuhankaku を食わないよう） */
    var romajiMap = [
        { key: "gyomatsuyakumonozenkakuhankaku", name: "行末約物全角/半角" },
        { key: "gyomatsuyakumonohankaku",        name: "行末約物半角" },
        { key: "gyomatsuyakumonozenkaku",        name: "行末約物全角" },
        { key: "yakumonohankaku",                name: "約物半角" },
        { key: "yakumonozenkaku",                name: "約物全角" },
        { key: "tsumegumi",                      name: "ツメ組み" },
        { key: "tsume",                          name: "ツメ組み" },
        { key: "betagumi",                       name: "ベタ組み" },
        { key: "beta",                           name: "ベタ組み" }
    ];
    var normalizedValue = workerNormalizeMojikumiName(value);
    var mapIndex;
    for (mapIndex = 0; mapIndex < romajiMap.length; mapIndex++) {
        if (normalizedValue.indexOf(romajiMap[mapIndex].key) === 0) { return romajiMap[mapIndex].name; }
    }
    return value;
}

/* --- note の本文（複数行になりうる「文字列」）を取り出す / Extract the multi-line body from a note --- */
/* "文字列：\n" 〜 "\n\nフォント：" を丸ごと本文とし、残りフィールドの走査開始位置もあわせて返す */
function workerExtractNoteBody(noteText) {
    var startMarker = "文字列：\n";
    var endMarker = "\n\nフォント：";
    var start = noteText.indexOf(startMarker);
    /* 本文自体が "フォント：" を含むことがあるので、最後の出現を本文の終端とする
       （本文より後ろのフィールドはすべて1行値なので、後方から探せば必ず本物に当たる） */
    var end = noteText.lastIndexOf(endMarker);
    if (start < 0 || end <= start) { return null; }
    return { bodyText: noteText.substring(start + startMarker.length, end), restIndex: end + 2 };
}

/* --- 一覧に出す項目名（表示順）/ Item labels to list, in display order --- */
/* handleJp が false（英語環境）のときは和文専用属性を一覧に出さない（旧 note に含まれていてもスキップ） */
function workerDisplayLabels(handleJp) {
    var displayOrder = ["文字列", "組み方向", "フォント", "フォントサイズ", "行送り", "自動行送り", "水平比率", "垂直比率", "カーニング", "プロポーショナルメトリクス", "トラッキング", "文字ツメ", "行揃え", "禁則", "文字組み", "文字カラー"];
    if (handleJp !== false) { return displayOrder; }
    var fields = workerNoteFields();
    var jpOnlyLabels = {};
    var fieldIndex;
    for (fieldIndex = 0; fieldIndex < fields.length; fieldIndex++) {
        if (fields[fieldIndex].jpOnly) { jpOnlyLabels[fields[fieldIndex].label] = true; }
    }
    var labels = [];
    var orderIndex;
    for (orderIndex = 0; orderIndex < displayOrder.length; orderIndex++) {
        if (!jpOnlyLabels[displayOrder[orderIndex]]) { labels.push(displayOrder[orderIndex]); }
    }
    return labels;
}

/* --- 一覧表示用に値を整える（文字ツメ＝%表記、禁則・文字組み＝日本語ラベル）/ Tidy a value for display --- */
function workerFormatDisplayValue(label, value) {
    if (label === "文字ツメ") {
        /* Tsume は 0.0〜1.0 で保存されているので %表記に直す */
        var tsumeNumber = parseFloat(value);
        return isNaN(tsumeNumber) ? value : (Math.round(tsumeNumber * 100) + "%");
    }
    if (label === "禁則") { return workerKinsokuToDisplay(value); }
    if (label === "文字組み") { return workerMojikumiToDisplay(value); }
    return value;
}

/* --- note を表示用にコンパクト整形 / Format the note for compact display --- */
function workerFormatNoteForDisplay(noteText, handleJp) {
    var displayLabels = workerDisplayLabels(handleJp);
    var parsedBody = workerExtractNoteBody(noteText);
    /* 単一行フィールドは本文を除いた残りだけを走査（本文行の誤マッチ防止。解析側と同じ方針） */
    var noteLines = (parsedBody ? noteText.substring(parsedBody.restIndex) : noteText).split("\n");
    var displayLines = [];
    var labelIndex;
    for (labelIndex = 0; labelIndex < displayLabels.length; labelIndex++) {
        var currentLabel = displayLabels[labelIndex];
        /* 本文の改行は ↵ に置き換えて1行に畳む */
        if (currentLabel === "文字列" && parsedBody) {
            displayLines.push("文字列： " + parsedBody.bodyText.replace(/[\r\n]/g, "↵"));
            continue;
        }
        var scanIndex;
        for (scanIndex = 0; scanIndex < noteLines.length; scanIndex++) {
            if (noteLines[scanIndex].indexOf(currentLabel + "：") === 0 && scanIndex + 1 < noteLines.length) {
                displayLines.push(currentLabel + "： " + workerFormatDisplayValue(currentLabel, noteLines[scanIndex + 1]));
                break;
            }
        }
    }
    return displayLines.join("\n");
}

/* --- 選択状態を検査（テキスト有無・メモ有無・表示用note） / Inspect selection state --- */
function workerInspectSelection(handleJp) {
    var currentSelection = workerGetSelection();
    if (typeof currentSelection === "string") { return currentSelection; }
    var candidates = workerFilterByType(currentSelection, ["GroupItem", "PathItem", "TextFrame"]);
    /* 複数選択時は「先頭のメモ付きオブジェクト」を採用（メモを持つ最初の対象） */
    var noteHolder = null;
    var scanIndex;
    for (scanIndex = 0; scanIndex < candidates.length; scanIndex++) {
        if (candidates[scanIndex].note && candidates[scanIndex].note.length > 0) { noteHolder = candidates[scanIndex]; break; }
    }
    if (!noteHolder) { return "NONOTE"; }
    var formattedNote = workerFormatNoteForDisplay(noteHolder.note, handleJp);
    if (!formattedNote || formattedNote.length < 1) { formattedNote = noteHolder.note; }
    /* NOTE:<表示用note> / NONOTE の形式で返す */
    return "NOTE:" + formattedNote;
}

/* --- 復元：メモ解析 / Restore: parse note --- */
function workerExtractTextAttributes(noteText) {
    var fields = workerNoteFields();
    var attributes = { x: null, y: null, savedBounds: null };
    var fieldIndex;
    for (fieldIndex = 0; fieldIndex < fields.length; fieldIndex++) { attributes[fields[fieldIndex].key] = null; }
    /* 本文（文字列）は複数行になりうるので丸ごと本文にする（内部の改行・空行も保持） */
    var parsedBody = workerExtractNoteBody(noteText);
    var restText = noteText;
    if (parsedBody) {
        attributes.text = parsedBody.bodyText;
        restText = noteText.substring(parsedBody.restIndex);
    }
    /* フォント以降の単一行フィールドは、本文を除いた残りだけを走査（本文行の誤マッチ防止） */
    var noteLines = restText.split("\n");
    var lineIndex;
    for (lineIndex = 0; lineIndex < noteLines.length; lineIndex++) {
        var currentLine = noteLines[lineIndex];
        var nextLine = (lineIndex + 1 < noteLines.length) ? noteLines[lineIndex + 1] : null;
        for (fieldIndex = 0; nextLine != null && fieldIndex < fields.length; fieldIndex++) {
            var field = fields[fieldIndex];
            /* 本文は切り出し済み。旧形式（1行だけの文字列）のときだけここで拾う */
            if (field.type === "body" && attributes.text != null) { continue; }
            if (currentLine.indexOf(field.label + "：") !== 0) { continue; }
            if (field.type === "num") { attributes[field.key] = parseFloat(nextLine); }
            else if (field.type === "bool") { attributes[field.key] = (nextLine === "true"); }
            else { attributes[field.key] = nextLine; }
            break;
        }
        if (currentLine.match(/^座標：\s*X\s*=\s*([-]?\d+(\.\d+)?),\s*Y\s*=\s*([-]?\d+(\.\d+)?)/)) {
            attributes.x = parseFloat(RegExp.$1);
            attributes.y = parseFloat(RegExp.$3);
        }
        if (currentLine.match(/L\s*=\s*([-]?\d+(?:\.\d+)?),\s*T\s*=\s*([-]?\d+(?:\.\d+)?),\s*R\s*=\s*([-]?\d+(?:\.\d+)?),\s*B\s*=\s*([-]?\d+(?:\.\d+)?)/)) {
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
    try {
        layer.locked = false;
        layer.visible = true;
    } catch (e) {}
}

/* --- 名前でレイヤーを探す（無ければ null）/ Find a layer by name --- */
function workerFindLayerByName(doc, layerName) {
    var findIndex;
    for (findIndex = 0; findIndex < doc.layers.length; findIndex++) {
        if (doc.layers[findIndex].name === layerName) { return doc.layers[findIndex]; }
    }
    return null;
}

/* --- 復元：退避（アウトライン）レイヤーを用意（既存 outlined_text があればロック解除して再利用）
       Restore: reuse the existing outlined_text layer (unlocked) if present, otherwise create a fresh stash layer --- */
/* 新規作成したレイヤーは一時名 __outlined_text_stash__ のままにし、確定時に outlined_text へ改名する。
   この一時名が「今回作ったレイヤーかどうか」の目印を兼ねる（ドキュメントに目印オブジェクトを置かない） */
function workerCreateOutlineStashLayer(doc) {
    if (!doc) { return null; }
    /* 既存の outlined_text レイヤーがあればロックを解除してそのまま退避先に使う（アウトラインを1枚に集約） */
    var stashLayer = workerFindLayerByName(doc, "outlined_text");
    if (!stashLayer) {
        stashLayer = doc.layers.add();
        stashLayer.name = "__outlined_text_stash__";
    }
    workerSetLayerUsable(stashLayer);
    return stashLayer;
}

/* --- 復元：restored_text レイヤーを用意 / Restore: get restored_text layer --- */
/* 復元は常に同名 restored_text を再利用して集約するため、番号付きレイヤーの自動統合は行わない
   （ユーザーが手動で作った restored_text1 等を巻き込まない） */
function workerCreateRestoredTextLayer(doc) {
    var targetLayer = workerFindLayerByName(doc, "restored_text");
    if (!targetLayer) {
        targetLayer = doc.layers.add();
        targetLayer.name = "restored_text";
    }
    workerSetLayerUsable(targetLayer);
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
            while (mergeLayer.pageItems.length > 0) {
                /* 退避済みアイテムはロックされていて動かせないので、解除して移動し元に戻す */
                var movingItem = mergeLayer.pageItems[0];
                var wasLocked = false;
                try { wasLocked = movingItem.locked; movingItem.locked = false; } catch (eUnlock) {}
                movingItem.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
                if (wasLocked) { workerSetAttr(movingItem, "locked", true); }
            }
        } catch (e1) {}
        /* 移動しきれなかったレイヤーは中身ごと消さない（統合できなかったアウトラインを守る） */
        if (mergeLayer.pageItems.length > 0) { continue; }
        try {
            while (mergeLayer.layers && mergeLayer.layers.length > 0) { mergeLayer.layers[0].remove(); }
            mergeLayer.remove();
        } catch (e2) {}
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
            workerSetAttr(normalizeLayer, "name", "outlined_text_dup" + dupCounter);
            dupCounter++;
        }
    }
}

/* --- 復元：指定レイヤーを確実にアクティブ化 / Restore: force active layer --- */
function workerSetActiveLayerStrict(doc, layer) {
    if (!doc || !layer) { return false; }
    try {
        doc.activeLayer = layer;
        if (doc.activeLayer === layer) { return true; }
        /* 参照が食い違う環境向けに、同一レイヤーをコレクションから引き直して再指定 */
        var reFindIndex;
        for (reFindIndex = 0; reFindIndex < doc.layers.length; reFindIndex++) {
            if (doc.layers[reFindIndex] === layer) { doc.activeLayer = doc.layers[reFindIndex]; break; }
        }
        return doc.activeLayer === layer;
    } catch (e) {
        return false;
    }
}

/* --- 一時アクションの名前とファイル / Names and file path of the temporary action --- */
/* セット名は固有名にする（"layer" のような一般名だとユーザーの同名アクションセットを消してしまう） */
function workerActionNames() {
    return {
        setName: "DynamicActionOutlineRestore",
        actionName: "change-to-template-layer",
        filePath: "~/AiTextOutlineRestoreAction.aia"
    };
}

/* --- 文字列を .aia の [ バイト長 16進 ] 形式へ / Encode a string as an .aia [ length hex ] token --- */
function workerEncodeActionText(sourceText) {
    var byteString = unescape(encodeURIComponent(sourceText));
    var hexText = "";
    var charIndex;
    for (charIndex = 0; charIndex < byteString.length; charIndex++) {
        var hexValue = byteString.charCodeAt(charIndex).toString(16);
        if (hexValue.length < 2) { hexValue = "0" + hexValue; }
        hexText += hexValue;
    }
    return "[ " + byteString.length + " " + hexText + " ]";
}

/* --- 復元：テンプレートレイヤー属性をアクションで付与 / Restore: apply template-layer attribute via action --- */
/* 読み込み → 実行 → 解放を1回で完結させ、finally で .aia とアクションセットを必ず片付ける
   （MakeTemplateLayer.jsx の playTemporaryAction と同じ方式）
   レイヤー名（titl）は決め打ちにせず対象レイヤーの実際の名前を注入する。決め打ちだと
   アクションがリネーム扱いになり、同名レイヤーが増える原因になるため */
function workerApplyTemplateLayerAttribute(layerName) {
    /* 「レイヤーオプション」イベントのパラメーター（key は FourCC：muid/name/titl/tmpl/show/lock/prvw/prnt/dim./pcnt） */
    var actionParams = [
        { key: 1836411236, valueType: "integer",   value: "4" },
        { key: 1851878757, valueType: "ustring",   value: "[ 36 e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383 a7e383b3 ]" },
        { key: 1953068140, valueType: "ustring",   value: workerEncodeActionText(layerName || "outlined_text") },
        { key: 1953329260, valueType: "boolean",   value: "1" },
        { key: 1936224119, valueType: "boolean",   value: "1" },
        { key: 1819239275, valueType: "boolean",   value: "1" },
        { key: 1886549623, valueType: "boolean",   value: "1" },
        { key: 1886547572, valueType: "boolean",   value: "0" },
        { key: 1684630830, valueType: "boolean",   value: "1" },
        { key: 1885564532, valueType: "unit real", value: "50.0", unit: "592474723" }
    ];
    var names = workerActionNames();
    /* アクション定義（.aia）は改行なしの1行。パラメーターだけ上表から組み立てる */
    var actionString = '/version 3/name ' + workerEncodeActionText(names.setName) + '/isOpen 1/actionCount 1/action-1 {' +
        ' /name ' + workerEncodeActionText(names.actionName) + ' /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1' +
        ' /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_Layer) /localizedName [ 9 e8a1a8e7a4ba203a20 ]' +
        ' /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount ' + actionParams.length;
    var paramIndex;
    for (paramIndex = 0; paramIndex < actionParams.length; paramIndex++) {
        var actionParam = actionParams[paramIndex];
        actionString += ' /parameter-' + (paramIndex + 1) + ' { /key ' + actionParam.key +
            ' /showInPalette 4294967295 /type (' + actionParam.valueType + ') /value ' + actionParam.value +
            (actionParam.unit ? (' /unit ' + actionParam.unit) : '') + ' }';
    }
    actionString += ' }}';
    var actionFile = new File(names.filePath);
    var isActionLoaded = false;
    var isFileOpen = false;
    /* 同名セットが残っていると doScript が別物を実行しかねないので、読み込む前に必ず解放しておく */
    try { app.unloadAction(names.setName, ""); } catch (eUnload) {}
    try {
        actionFile.encoding = 'UTF-8';
        actionFile.lineFeed = 'Unix';
        if (!actionFile.open('w')) { return; }
        isFileOpen = true;
        actionFile.write(actionString);
        actionFile.close();
        isFileOpen = false;
        app.loadAction(actionFile);
        isActionLoaded = true;
        app.doScript(names.actionName, names.setName, false);
    } catch (eAction) {
    } finally {
        if (isFileOpen) { try { actionFile.close(); } catch (eClose) {} }
        if (actionFile.exists) { try { actionFile.remove(); } catch (eRemove) {} }
        if (isActionLoaded) { try { app.unloadAction(names.setName, ""); } catch (eDone) {} }
    }
}

/* --- 復元：残っているテンプレートアクションを解放（パレットを閉じるとき） / Unload the template action on palette close --- */
/* 通常は workerApplyTemplateLayerAttribute の finally で解放済み。異常終了時の保険として実行する */
function workerUnloadTemplateAction() {
    var names = workerActionNames();
    try { app.unloadAction(names.setName, ""); } catch (e) {}
    return "OK";
}

/* --- 復元：テンプレートレイヤーを確定 / Restore: finalize template layer --- */
/* 退避レイヤーの参照はアクションを挟んでも有効なので、同名レイヤーが増えた場合も
   「この参照以外の outlined_text をリネームする」だけで1枚に保てる（目印オブジェクトは不要） */
function workerFinalizeTemplateLayer(outlinedTextLayer) {
    if (!outlinedTextLayer) { return; }
    var doc = app.activeDocument;
    workerSetLayerUsable(outlinedTextLayer);
    workerMergeExistingOutlinedLayers(doc, outlinedTextLayer);
    if (outlinedTextLayer.name !== "outlined_text") { outlinedTextLayer.name = "outlined_text"; }
    try { outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (eBack) {}
    workerNormalizeOutlinedLayerNames(doc, outlinedTextLayer);
    /* アクティブ化できたときだけテンプレート属性アクションを実行。
       アクティブ化に失敗しても、最背面とロックは必ず通す */
    try {
        if (workerSetActiveLayerStrict(doc, outlinedTextLayer)) {
            workerApplyTemplateLayerAttribute(outlinedTextLayer.name);
            /* 念のためアクション後にもう一度リネームを通す（増えた同名レイヤーは次回の統合で吸収される） */
            workerNormalizeOutlinedLayerNames(doc, outlinedTextLayer);
        }
    } catch (eAct) {}
    try {
        outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK);
        outlinedTextLayer.locked = true;
    } catch (eLock) {}
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

/* --- 行揃え（保存用テキスト ↔ Justification）/ Justification table --- */
function workerJustificationPairs() {
    return [
        { labelText: "中央揃え",               justification: Justification.CENTER },
        { labelText: "右揃え",                 justification: Justification.RIGHT },
        { labelText: "均等配置（最終行左）",   justification: Justification.FULLJUSTIFYLASTLINELEFT },
        { labelText: "均等配置（最終行中央）", justification: Justification.FULLJUSTIFYLASTLINECENTER },
        { labelText: "均等配置（最終行右）",   justification: Justification.FULLJUSTIFYLASTLINERIGHT },
        { labelText: "均等配置",               justification: Justification.FULLJUSTIFY }
    ];
}

/* --- 行揃えを保存用テキストへ / Serialize justification to text --- */
function workerJustificationToText(justification) {
    var pairs = workerJustificationPairs();
    var i;
    for (i = 0; i < pairs.length; i++) {
        if (justification == pairs[i].justification) { return pairs[i].labelText; }
    }
    return "左揃え";
}

/* --- 復元：行揃えテキストを enum へ / Restore: resolve alignment label to Justification --- */
/* 復元先は新規テキストフレームなので既定は LEFT。LEFT は代入が無視されることがあるが既定と一致するため実害なし */
function workerResolveJustification(justificationText) {
    var pairs = workerJustificationPairs();
    var i;
    for (i = 0; i < pairs.length; i++) {
        if (justificationText === pairs[i].labelText) { return pairs[i].justification; }
    }
    return Justification.LEFT;
}

/* --- 禁則を保存用テキストへ / Serialize kinsoku to text --- */
/* kinsoku はセット名の文字列（"Soft"／"Hard"／"Soft_v2" 等）。取得不能・なしは "なし" とする */
function workerKinsokuToText(paragraphAttributes) {
    try {
        var value = paragraphAttributes.kinsoku;
        if (value == null) { return "なし"; }
        if (typeof value === "string") {
            if (value === "" || value === "None") { return "なし"; }
            return value;
        }
        if (value.name != null) { return value.name; }
        return "なし";
    } catch (e) {
        return "なし";
    }
}

/* --- 復元：禁則を適用 / Restore: apply kinsoku --- */
/* 「なし」はスクリプトから代入できずエラーになるため、なし／空／不正値はスキップ（既定のまま） */
function workerApplyKinsoku(textRange, kinsokuText) {
    if (kinsokuText == null || kinsokuText === "" || kinsokuText === "なし" || kinsokuText === "None") { return; }
    try {
        textRange.paragraphAttributes.kinsoku = kinsokuText;
    } catch (e) {
    }
}

/* --- 文字組み名の正規化（表記ゆれ吸収）/ Normalize a mojikumi name for tolerant compare --- */
/* 内部名は環境で空白・区切り・大小が揺れる（例: "Gyomatsu Yakumono Zenkaku Hankaku"）ので畳んで比較 */
function workerNormalizeMojikumiName(raw) {
    return String(raw).replace(/[\s　_\/／・]/g, "").toLowerCase();
}

/* --- 文字組みプリセット（doc.mojikumiSet の index ↔ 日本語ラベル）/ Mojikumi presets (index <-> label) --- */
/* ApplyMojikumi.jsx と同一。組み込み7セットは doc.mojikumiSet[0..6] に固定対応（「なし」は index -1 で対象外） */
function workerMojikumiPresets() {
    return [
        { mojikumiIndex: 0, labelText: "行末約物全角/半角" },
        { mojikumiIndex: 1, labelText: "約物半角" },
        { mojikumiIndex: 2, labelText: "行末約物半角" },
        { mojikumiIndex: 3, labelText: "行末約物全角" },
        { mojikumiIndex: 4, labelText: "約物全角" },
        { mojikumiIndex: 5, labelText: "ツメ組み" },
        { mojikumiIndex: 6, labelText: "ベタ組み" }
    ];
}

/* --- 表記ゆれを吸収して文字組みプリセットを引く（無ければ null）/ Find a mojikumi preset by tolerant name compare --- */
function workerFindMojikumiPreset(name) {
    var normalizedName = workerNormalizeMojikumiName(name);
    var presets = workerMojikumiPresets();
    var presetIndex;
    for (presetIndex = 0; presetIndex < presets.length; presetIndex++) {
        if (workerNormalizeMojikumiName(presets[presetIndex].labelText) === normalizedName) { return presets[presetIndex]; }
    }
    return null;
}

/* --- 適用中の文字組み（オブジェクト or 内部名）を doc.mojikumiSet の位置で日本語ラベルへ / Resolve applied mojikumi to preset label --- */
/* 同一オブジェクト参照 → 正規化した内部名の順で突き合わせ、見つかった index のラベルを返す。引けなければ null */
function workerMojikumiLabelFromApplied(applied) {
    var presets = workerMojikumiPresets();
    var sets = null;
    try { sets = app.activeDocument.mojikumiSet; } catch (eSets) { sets = null; }
    if (sets == null) { return null; }
    var appliedName = null;
    try { appliedName = (typeof applied === "string") ? applied : applied.name; } catch (eName) { appliedName = null; }
    var normApplied = (appliedName != null) ? workerNormalizeMojikumiName(appliedName) : null;
    var i;
    for (i = 0; i < presets.length; i++) {
        var idx = presets[i].mojikumiIndex;
        if (idx < 0 || idx >= sets.length) { continue; }
        var matched = false;
        try { if (sets[idx] === applied) { matched = true; } } catch (eId) {}
        if (!matched && normApplied != null) {
            try { if (workerNormalizeMojikumiName(sets[idx].name) === normApplied) { matched = true; } } catch (eNm) {}
        }
        if (matched) { return presets[i].labelText; }
    }
    return null;
}

/* --- 文字組みアキ量設定を保存用テキストへ / Serialize mojikumi to text --- */
/* 可能なら日本語ラベル（doc.mojikumiSet の index で解決）で保存。引けない場合は内部名を素通し（後方互換） */
function workerMojikumiToText(paragraphAttributes) {
    try {
        var value = paragraphAttributes.mojikumi;
        if (value == null) { return "なし"; }
        if (typeof value === "string") {
            if (value === "" || value === "None") { return "なし"; }
            if (value === "なし") { return "なし"; }
            var labelFromName = workerMojikumiLabelFromApplied(value);
            return (labelFromName != null) ? labelFromName : value;
        }
        var label = workerMojikumiLabelFromApplied(value);
        if (label != null) { return label; }
        if (value.name != null) { return value.name; }
        return "なし";
    } catch (e) {
        return "なし";
    }
}

/* --- 復元：文字組みアキ量設定を適用 / Restore: apply mojikumi --- */
/* 「なし」は代入可。日本語ラベル → その index の doc.mojikumiSet を適用（ApplyMojikumi と同方式）。
   旧 note の内部名は doc.mojikumiSet 名との正規化一致でフォールバック */
function workerApplyMojikumi(textRange, mojikumiText) {
    if (mojikumiText == null || mojikumiText === "") { return; }
    try {
        if (mojikumiText === "なし" || mojikumiText === "None") {
            textRange.paragraphAttributes.mojikumi = "なし";
            return;
        }
        var sets = app.activeDocument.mojikumiSet;
        var normalizedText = workerNormalizeMojikumiName(mojikumiText);
        var preset = workerFindMojikumiPreset(mojikumiText);
        if (preset && preset.mojikumiIndex >= 0 && preset.mojikumiIndex < sets.length) {
            textRange.paragraphAttributes.mojikumi = sets[preset.mojikumiIndex];
            return;
        }
        var setIndex;
        for (setIndex = 0; setIndex < sets.length; setIndex++) {
            if (workerNormalizeMojikumiName(sets[setIndex].name) === normalizedText) {
                textRange.paragraphAttributes.mojikumi = sets[setIndex];
                return;
            }
        }
    } catch (e) {
    }
}

/* --- 復元：カーニング適用（AutoKerning.jsx の applyKerningToRanges と同一） / Apply kerning --- */
/* proportionalMetrics は note の保存値を優先。未記録の旧 note だけ、メトリクス（AUTO）連動で決める */
function workerApplyKerning(textRange, kerningMethod, proportionalMetrics) {
    var useProportionalMetrics = (proportionalMetrics != null) ? proportionalMetrics : (kerningMethod === AutoKernType.AUTO);
    try {
        textRange.characterAttributes.kerningMethod = kerningMethod;
        textRange.characterAttributes.proportionalMetrics = useProportionalMetrics;
    } catch (e) {
        /* 適用できない範囲はスキップ / skip ranges that can't take these */
    }
}

/* --- 復元：文字属性を適用 / Restore: apply character attributes --- */
/* 各属性は workerSetAttr 経由で個別に適用（1つ失敗しても他の復元を止めない）。null は未記録の旧 note なので触らない */
function workerApplyCharacterAttributes(textFrame, attributes, restoreReport, jp) {
    var attrs = textFrame.textRange.characterAttributes;
    /* フォントだけ個別に：見つからなければ既定フォントのまま。他属性の適用は続行する */
    try {
        attrs.textFont = app.textFonts.getByName(attributes.font);
    } catch (eFont) {
        if (restoreReport) { restoreReport.fontFallback = true; }
    }
    workerSetAttr(attrs, "size", attributes.fontSize);
    /* 自動行送り ON のときは行送りを明示せず自動計算に任せる（旧 note は autoLeading 未記録なので false 扱い） */
    if (attributes.autoLeading === true) {
        workerSetAttr(attrs, "autoLeading", true);
    } else {
        workerSetAttr(attrs, "autoLeading", false);
        workerSetAttr(attrs, "leading", attributes.leading);
    }
    workerSetAttr(attrs, "tracking", attributes.tracking);
    /* 水平比率・垂直比率（長体／平体） */
    workerSetAttr(attrs, "horizontalScale", attributes.horizontalScale);
    workerSetAttr(attrs, "verticalScale", attributes.verticalScale);
    /* 文字ツメ・組み方向は和文専用。英語環境（jp=false）はスキップ（旧 note に含まれていても適用しない） */
    if (jp) { workerSetAttr(attrs, "Tsume", attributes.tsume); }
    if (jp && attributes.orientation != null) {
        workerSetAttr(textFrame, "orientation", (attributes.orientation === "縦組み") ? TextOrientation.VERTICAL : TextOrientation.HORIZONTAL);
    }
    if (attributes.kerningText != null) {
        /* カーニング方式を復元。proportionalMetrics も note の保存値をそのまま渡す */
        workerApplyKerning(textFrame.textRange, workerResolveKernType(attributes.kerningText), attributes.proportionalMetrics);
    } else {
        /* 旧 note（カーニング未記録）は保存済みの proportionalMetrics を復元 */
        workerSetAttr(attrs, "proportionalMetrics", attributes.proportionalMetrics);
    }
    if (attributes.colorText != null) {
        workerSetAttr(attrs, "fillColor", workerColorFromText(attributes.colorText));
    }
}

/* --- 復元：段落属性を適用 / Restore: apply paragraph attributes --- */
function workerApplyParagraphAttributes(textFrame, attributes, jp) {
    /* 行揃え。新規フレームの既定は LEFT なので左揃えは実質そのまま */
    if (attributes.justificationText != null) {
        workerSetAttr(textFrame.textRange.paragraphAttributes, "justification", workerResolveJustification(attributes.justificationText));
    }
    /* 禁則・文字組みアキ量設定は和文専用。英語環境はスキップ（未記録・不正値は各 apply 側で無視される） */
    if (!jp) { return; }
    workerApplyKinsoku(textFrame.textRange, attributes.kinsokuText);
    workerApplyMojikumi(textFrame.textRange, attributes.mojikumiText);
}

/* --- 復元：テキストフレーム再生成 / Restore: recreate the text frame --- */
function workerCreateRestoredTextFrame(sourceItem, attributes, restoredLayer, restoreReport, handleJp) {
    var jp = (handleJp !== false);
    var targetLayer = restoredLayer || sourceItem.layer;
    var textFrame = targetLayer.textFrames.add();
    textFrame.contents = attributes.text;
    textFrame.position = [
        (attributes.x !== null) ? attributes.x : sourceItem.left,
        (attributes.y !== null) ? attributes.y : sourceItem.top
    ];
    workerApplyCharacterAttributes(textFrame, attributes, restoreReport, jp);
    workerApplyParagraphAttributes(textFrame, attributes, jp);
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
function workerRestoreItem(sourceItem, outlinedTextLayer, restoredTextLayer, restoreReport, stashOutline, handleJp) {
    var noteText = sourceItem.note;
    if (!noteText) { return false; }
    var attributes = workerExtractTextAttributes(noteText);
    if (!attributes || attributes.text == null) { return false; }
    var textFrame = workerCreateRestoredTextFrame(sourceItem, attributes, restoredTextLayer, restoreReport, handleJp);
    /* 位置合わせは元アウトラインを削除する前に行う（savedBounds が無い旧 note は sourceItem を参照するため） */
    workerAlignTextFrameByBounds(attributes.savedBounds || sourceItem, textFrame);
    if (stashOutline === false) {
        /* 「アウトラインデータを残す」OFF：元アウトラインは退避せず削除 */
        try { sourceItem.remove(); } catch (eDel) {}
    } else {
        /* 先に退避してから淡く＋ロック（ロック解除状態で移動する方が確実） */
        workerMoveToOutlinedLayer(sourceItem, outlinedTextLayer);
        /* 種別を問わず淡く＋ロック。失敗しても復元処理は止めない */
        try {
            sourceItem.opacity = 30;
            sourceItem.locked = true;
        } catch (eDim) {}
    }
    if (restoreReport) { restoreReport.restored++; }
    return true;
}

/* --- 復元：1件も復元できなかったときに作成したレイヤーを片付ける / Drop the layers created for a failed restore --- */
/* 退避レイヤーは既存の outlined_text を再利用していることがあるので、今回新規作成したもの
   （一時名のまま）だけを削除する。再利用した場合は過去のアウトラインごと消さずロックだけ戻す */
function workerDiscardUnusedLayers(outlinedTextLayer, restoredTextLayer) {
    try {
        if (outlinedTextLayer) {
            if (outlinedTextLayer.name === "__outlined_text_stash__") { outlinedTextLayer.remove(); }
            else { workerSetAttr(outlinedTextLayer, "locked", true); }
        }
    } catch (eStash) {}
    try {
        if (restoredTextLayer && restoredTextLayer.pageItems.length < 1 && (!restoredTextLayer.layers || restoredTextLayer.layers.length < 1)) { restoredTextLayer.remove(); }
    } catch (eRestored) {}
}

/* --- 復元：エントリ / Restore: entry --- */
function workerRestoreText(keepOutline, separateLayer, handleJp) {
    var currentSelection = workerGetSelection();
    if (typeof currentSelection === "string") { return currentSelection; }
    var restorableItems = workerFilterByType(currentSelection, ["GroupItem", "PathItem"]);
    if (restorableItems.length < 1) { return "NOTGT"; }
    var doc = app.activeDocument;
    /* 既定はアウトラインを残す。false のときだけ退避レイヤーを作らず元アウトラインを削除 */
    var stashOutline = (keepOutline !== false);
    var outlinedTextLayer = stashOutline ? workerCreateOutlineStashLayer(doc) : null;
    /* 既定は復元テキストを別レイヤー（restored_text）へ。false のときは元アウトラインと同じレイヤーへ置く（null で sourceItem.layer にフォールバック） */
    var restoredTextLayer = (separateLayer !== false) ? workerCreateRestoredTextLayer(doc) : null;
    var restoreReport = { restored: 0, fontFallback: false };
    /* 1件ずつ try で囲み、途中で失敗しても残りの復元を続ける */
    var runError = null;
    var restoreIndex;
    for (restoreIndex = 0; restoreIndex < restorableItems.length; restoreIndex++) {
        try {
            workerRestoreItem(restorableItems[restoreIndex], outlinedTextLayer, restoredTextLayer, restoreReport, stashOutline, handleJp);
        } catch (eRun) {
            if (runError == null) { runError = String(eRun); }
        }
    }
    if (restoreReport.restored < 1) {
        workerDiscardUnusedLayers(outlinedTextLayer, restoredTextLayer);
        return runError ? ("ERR:" + runError) : "NONOTE";
    }
    /* 一部でも成功していれば、その分を確定（途中エラーでも退避レイヤーを宙ぶらりんにしない） */
    if (stashOutline && outlinedTextLayer) { workerFinalizeTemplateLayer(outlinedTextLayer); }
    workerSetAttr(doc, "activeLayer", restoredTextLayer);
    return "OK:" + restoreReport.restored + (restoreReport.fontFallback ? ":FONT" : "") + (runError ? ":PARTIAL" : "");
}

// ワーカー関数はすべてここに登録（追加漏れ防止）
var WORKER_FUNCS = [
    workerRound,
    workerSetAttr,
    workerToggleAttributesPanel,
    workerSetLayerUsable,
    workerFindLayerByName,
    workerActionNames,
    workerEncodeActionText,
    workerNoteFields,
    workerBuildMemoText,
    workerKerningPairs,
    workerKerningToText,
    workerResolveKernType,
    workerProcessTextFrame,
    workerGetSelection,
    workerFilterByType,
    workerRunOutline,
    workerKinsokuToDisplay,
    workerMojikumiToDisplay,
    workerExtractNoteBody,
    workerDisplayLabels,
    workerFormatDisplayValue,
    workerFormatNoteForDisplay,
    workerInspectSelection,
    workerExtractTextAttributes,
    workerAlignTextFrameByBounds,
    workerCreateOutlineStashLayer,
    workerCreateRestoredTextLayer,
    workerMergeExistingOutlinedLayers,
    workerNormalizeOutlinedLayerNames,
    workerSetActiveLayerStrict,
    workerApplyTemplateLayerAttribute,
    workerUnloadTemplateAction,
    workerFinalizeTemplateLayer,
    workerColorToText,
    workerColorFromText,
    workerJustificationPairs,
    workerJustificationToText,
    workerResolveJustification,
    workerKinsokuToText,
    workerApplyKinsoku,
    workerNormalizeMojikumiName,
    workerMojikumiPresets,
    workerFindMojikumiPreset,
    workerMojikumiLabelFromApplied,
    workerMojikumiToText,
    workerApplyMojikumi,
    workerApplyKerning,
    workerApplyCharacterAttributes,
    workerApplyParagraphAttributes,
    workerCreateRestoredTextFrame,
    workerMoveToOutlinedLayer,
    workerRestoreItem,
    workerDiscardUnusedLayers,
    workerRestoreText
];

// メモの読み込み（選択状態の検査）だけに必要な worker。毎回フルバンドルを送らないための最小構成
var INSPECT_FUNCS = [
    workerGetSelection,
    workerFilterByType,
    workerNoteFields,
    workerExtractNoteBody,
    workerDisplayLabels,
    workerKinsokuToDisplay,
    workerNormalizeMojikumiName,
    workerMojikumiPresets,
    workerFindMojikumiPreset,
    workerMojikumiLabelFromApplied,
    workerMojikumiToDisplay,
    workerFormatDisplayValue,
    workerFormatNoteForDisplay,
    workerInspectSelection
];

// 属性パネルの開閉だけに必要な worker
var ATTRIBUTES_FUNCS = [workerToggleAttributesPanel];

// パレットを閉じるときの後始末だけに必要な worker
var CLEANUP_FUNCS = [workerActionNames, workerUnloadTemplateAction];

// ==============================
// BridgeTalk 委譲 / Delegation to main engine
// ==============================

/* Function.toString() は関数の前後に周辺コメントの断片（閉じ記号を欠いたもの）を含めて返すことがある。
   そのまま連結すると未終端コメントが次の関数を飲み込むので、宣言行から閉じ括弧行までを行単位で抜き出す */
function sliceFunctionSource(rawSource) {
    var lines = String(rawSource).replace(/\r\n?/g, "\n").split("\n");
    var first = 0;
    while (first < lines.length && lines[first].indexOf("function ") !== 0) { first++; }
    if (first >= lines.length) { return rawSource; } /* 想定外の形式：そのまま返す */
    var last = lines.length - 1;
    while (last > first && !/^\s*\}\s*$/.test(lines[last])) { last--; }
    return lines.slice(first, last + 1).join("\n");
}

/* 束ねたソースは呼び出しのたびに作り直さず、関数配列ごとにキャッシュする */
var WORKER_SOURCE_CACHE = [];

function buildWorkerSource(funcs, entryCall) {
    var cacheIndex;
    for (cacheIndex = 0; cacheIndex < WORKER_SOURCE_CACHE.length; cacheIndex++) {
        if (WORKER_SOURCE_CACHE[cacheIndex].funcs === funcs) { return WORKER_SOURCE_CACHE[cacheIndex].source + entryCall; }
    }
    var source = "";
    var funcIndex;
    for (funcIndex = 0; funcIndex < funcs.length; funcIndex++) {
        source += sliceFunctionSource(funcs[funcIndex].toString()) + "\n";
    }
    WORKER_SOURCE_CACHE.push({ funcs: funcs, source: source });
    return source + entryCall;
}

/* funcs 省略時は全 worker を送信。メモ読み込みなど軽量呼び出しは必要な関数だけ渡す */
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

// worker が返すコードとステータス文言の対応
var STATUS_CODE_KEYS = {
    NODOC: 'status.nodoc',
    NOSEL: 'status.nosel',
    NOTGT: 'status.notgt',
    NONOTE: 'status.nonote'
};

/* コード（NODOC 等）に対応する文言。未知のコードは null */
function statusTextForCode(result) {
    return (result != null && STATUS_CODE_KEYS[result]) ? L(STATUS_CODE_KEYS[result]) : null;
}

function applyResultToStatus(win, result, doneKey) {
    if (result == null) { setStatus(win, L('status.err')); return; }
    if (result.indexOf("OK") === 0) {
        var parts = result.split(":");
        var count = (parts.length > 1) ? parts[1] : "";
        var message = L(doneKey) + (count ? " (" + count + ")" : "");
        if (result.indexOf("FONT") >= 0) { message += " / " + L('status.fontWarn'); }
        if (result.indexOf("PARTIAL") >= 0) { message += " / " + L('status.partial'); }
        setStatus(win, message);
        return;
    }
    var codeText = statusTextForCode(result);
    if (codeText) { setStatus(win, codeText); return; }
    if (result.indexOf("ERR") === 0) { setStatus(win, L('status.err') + ": " + result.substring(4)); return; }
    setStatus(win, result);
}

/* 真偽値を worker 呼び出しの引数文字列へ / Boolean as a worker-call argument */
function argBool(value) {
    return value ? "true" : "false";
}

/* worker を呼び、送信自体が失敗してもエラー文字列に畳む（isBusy を立てたままにしない）
   Call a worker, folding a send failure into an error string */
function callWorkerSafely(entryCall, funcs) {
    try {
        return callWorker(entryCall, funcs);
    } catch (e) {
        return "ERR:" + e;
    }
}

/* 再入防止つきで worker を呼び、結果をステータスへ反映 / Call a worker behind the busy guard */
function runWorkerTask(win, entryCall, doneKey) {
    if (isBusy) { return null; }
    isBusy = true;
    setStatus(win, L('status.busy'));
    var result = callWorkerSafely(entryCall);
    isBusy = false;
    applyResultToStatus(win, result, doneKey);
    return result;
}

function onOutlineClick(win) {
    if (isBusy) return; // 再入防止
    runWorkerTask(win, "workerRunOutline(" + argBool(HANDLE_JP) + ");", 'status.doneOutline');
    refreshSelectedNote(win, true); // アウトライン化直後のメモを表示
}

function onRestoreClick(win) {
    if (isBusy) return;
    // 復元の直前に「読み込み」ロジックを強制実行し、対象のメモを表示（移動前の状態を見せる）
    refreshSelectedNote(win, true);
    // チェックボックスの状態を worker に渡す（残す／別レイヤー）
    var keepOutline = win.keepOutlineCheck ? win.keepOutlineCheck.value : true;
    var separateLayer = win.separateLayerCheck ? win.separateLayerCheck.value : true;
    runWorkerTask(win, "workerRestoreText(" + argBool(keepOutline) + ", " + argBool(separateLayer) + ", " + argBool(HANDLE_JP) + ");", 'status.doneRestore');
}

/* 属性パネルの開閉（メインエンジンにメニューコマンドを委譲）/ Toggle Attributes panel (delegated) */
function onAttributesClick(win) {
    if (isBusy) return;
    callWorkerSafely("workerToggleAttributesPanel();", ATTRIBUTES_FUNCS);
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
        var row = win.noteList.add("item", localizeFromTable(LIST_ITEM_LABELS_EN, itemLabel));
        row.subItems[0].text = localizeFromTable(LIST_VALUE_LABELS_EN, itemValue);
    }
}

function refreshSelectedNote(win, keepStatus) {
    if (isBusy) return;
    isBusy = true;
    var result = callWorkerSafely("workerInspectSelection(" + argBool(HANDLE_JP) + ");", INSPECT_FUNCS);
    isBusy = false;
    win.noteList.removeAll();
    if (result != null && result.indexOf("NOTE:") === 0) {
        populateNoteList(win, result.substring(5));
        if (!keepStatus) { setStatus(win, L('status.memoLoaded')); }
        return;
    }
    if (keepStatus) { return; }
    setStatus(win, statusTextForCode(result) || L('status.err'));
}

/* パレットを閉じるときの後始末：読み込んだアクションを解放 */
function performCloseCleanup() {
    callWorkerSafely("workerUnloadTemplateAction();", CLEANUP_FUNCS);
    $.global.__textOutlineMemoPalette = null;
}

/* パネル内のボタンを1つ追加 / Add a button with its help tip */
function addButton(parent, labelKey, tipKey, handler) {
    var button = parent.add("button", undefined, L(labelKey));
    button.helpTip = L(tipKey);
    button.onClick = handler;
    return button;
}

/* パネルを1つ追加 / Add a panel with its help tip */
function addPanel(win, labelKey, tipKey) {
    var panel = win.add("panel", undefined, L(labelKey));
    panel.helpTip = L(tipKey);
    setupPanel(panel, 6);
    return panel;
}

/* アウトライン化パネル（最上部） / Outline panel */
function buildOutlinePanel(win) {
    var outlinePanel = addPanel(win, 'panel.outline', 'panel.outlineTip');
    var outlineButtonRow = outlinePanel.add("group");
    setupRow(outlineButtonRow, "left", 8); // ボタン列は左寄せ
    addButton(outlineButtonRow, 'button.outline', 'button.outlineTip', function () { onOutlineClick(win); });
}

/* 選択オブジェクトのメモ一覧パネル / Selected object's note panel */
function buildNotePanel(win) {
    var selectedObjectPanel = addPanel(win, 'panel.selected', 'panel.selectedTip');

    win.noteList = selectedObjectPanel.add("listbox", undefined, [], {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: [L('listCol.item'), L('listCol.value')],
        columnWidths: [140, 170]
    });
    win.noteList.preferredSize = [320, 340]; // 16行分 / 16 rows
    win.noteList.helpTip = L('listCol.hint');

    // メモ操作の行：左＝属性パネル / 中央＝スペーサー / 右＝メモを読み込み
    var noteActionRow = selectedObjectPanel.add("group");
    noteActionRow.orientation = "row";
    noteActionRow.alignment = "fill"; // パネル幅いっぱいに広げて左右に振り分ける
    noteActionRow.alignChildren = ["fill", "center"];
    noteActionRow.margins = [0, 6, 0, 0]; // 行の上に余白 +6
    noteActionRow.spacing = 8;

    var attributesButton = addButton(noteActionRow, 'button.attributes', 'button.attributesTip', function () { onAttributesClick(win); });
    attributesButton.alignment = ["left", "center"];

    // 中央：フレキシブルスペーサー（左右のボタンを両端へ押し広げる）
    var noteActionSpacer = noteActionRow.add("group");
    noteActionSpacer.alignment = ["fill", "center"];

    var loadNoteButton = addButton(noteActionRow, 'button.load', 'button.loadTip', function () { refreshSelectedNote(win, false); });
    loadNoteButton.alignment = ["right", "center"];
}

/* 復元パネル（ボタン＋オプション） / Restore panel */
function buildRestorePanel(win) {
    var restorePanel = addPanel(win, 'panel.restore', 'panel.restoreTip');

    var restoreButtonRow = restorePanel.add("group");
    setupRow(restoreButtonRow, "left", 8); // ボタン列は左寄せ
    restoreButtonRow.margins = [0, 0, 0, 5]; // ボタンの下に余白 +5
    addButton(restoreButtonRow, 'button.restore', 'button.restoreTip', function () { onRestoreClick(win); });

    // 復元オプション：アウトラインを残す／別レイヤーに復元（いずれも既定 ON）
    win.keepOutlineCheck = restorePanel.add("checkbox", undefined, L('option.keepOutline'));
    win.keepOutlineCheck.helpTip = L('option.keepOutlineTip');
    win.keepOutlineCheck.value = true;

    win.separateLayerCheck = restorePanel.add("checkbox", undefined, L('option.separateLayer'));
    win.separateLayerCheck.helpTip = L('option.separateLayerTip');
    win.separateLayerCheck.value = true;
}

function showPalette() {
    // 多重起動防止：既存パレットがあれば閉じる（$.global で常駐参照して IIFE をまたいで保持）
    if ($.global.__textOutlineMemoPalette) {
        try { $.global.__textOutlineMemoPalette.close(); } catch (e) {}
        $.global.__textOutlineMemoPalette = null;
    }

    var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
    setupWindow(win);

    buildOutlinePanel(win);
    buildNotePanel(win);
    buildRestorePanel(win);

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
