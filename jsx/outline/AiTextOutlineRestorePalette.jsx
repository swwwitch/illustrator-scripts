#target illustrator
#targetengine "TextOutlineWithMemo"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name

TextOutlineRestorePalette.jsx

UI messages support Japanese/English. (Note format remains Japanese for compatibility.)

### 概要

- 常駐パレットから、選択テキストのアウトライン化（メモ付き）とその復元を行う
- アウトライン化の直前に、文字・段落属性を note（メモ）へテキスト形式で保存
- 復元時は note を解析してテキストフレームを再生成し、フォント・サイズ・行送り・カーニング・
  文字カラー・行揃え・禁則・文字組み・長体／平体・組み方向・座標などをまとめて復元
- 元のアウトラインは outlined_text レイヤーへ退避（淡く＋ロック）
- 選択オブジェクトの note をパネルに一覧表示（読み込みボタン）
- 旧 note（属性が未記録）は該当項目をスキップして従来動作を維持（後方互換）
- DOM 処理はすべてメインエンジンへ BridgeTalk 委譲

### Overview

- Outline selected text (saving its attributes to the note) and restore it back, from a persistent palette
- Before outlining, character/paragraph attributes are serialized into the object's note as text
- On restore, the note is parsed to recreate the text frame, bringing back font, size, leading,
  kerning, fill color, alignment, kinsoku, mojikumi, horizontal/vertical scale, orientation, position, etc.
- The original outlines are moved to the outlined_text layer (dimmed and locked)
- The selected object's note is listed in the panel (Load button)
- Older notes without newer fields fall back to previous behavior (backward compatible)
- All DOM work is delegated to the main engine via BridgeTalk

### 主な機能 / Features

- アウトライン化（メモ付き）：各種プロパティを note に保存
- テキスト復元：note からテキストを再生成し、元アウトラインは outlined_text レイヤーへ退避
- 選択オブジェクトの note をパネルに表示（読み込みボタン）

### 保存・復元するプロパティ / Supported properties

- 文字列（本文・複数行対応）/ Text contents (multi-line)
- 組み方向（縦組み／横組み）/ Orientation (vertical / horizontal)
- フォント（PostScript 名）/ Font (PostScript name)
- フォントサイズ / Font size
- 行送り / Leading
- 自動行送り（オンのときは行送り値を自動計算に戻す）/ Auto leading
- 水平比率・垂直比率（長体／平体）/ Horizontal & vertical scale
- カーニング方式（メトリクス／オプティカル／和文等幅／なし）/ Kerning method
- プロポーショナルメトリクス（カーニング方式に連動）/ Proportional metrics
- トラッキング / Tracking
- 文字ツメ（Tsume）/ Tsume
- 行揃え（左／中央／右／均等配置）/ Alignment (left / center / right / justify)
- 禁則（セット名。「なし」はスクリプト制約により復元スキップ）/ Kinsoku (set name; "none" is skipped)
- 文字組みアキ量設定（セット名で復元。「なし」対応）/ Mojikumi (restored by set name)
- 文字カラー（CMYK／RGB／グレー／スポット）/ Fill color (CMYK / RGB / Gray / Spot)
- 座標（geometricBounds 基準。listbox 非表示）/ Coordinates (by geometricBounds; not shown in the list)

※ グラデーション／パターン塗り、混在属性は非対応（先頭文字の値を採用）
※ Gradient/pattern fills and mixed attributes are not supported (the first character's value is used)

### note

https://note.com/dtp_transit/n/n3e0f241508db

### 更新履歴 / Change Log

- v1.0 (20240723) : 初期バージョン
- v1.9 (20260704) : 常駐パレット化＋テキスト復元・メモ表示・カーニング／文字カラー復元／複数行対応・堅牢化
- v1.10 (20260705) : 行揃え・自動行送りフラグ・水平比率／垂直比率（長体／平体）・禁則・文字組みアキ量設定の保存・復元に対応
- v2.0.0 (20260705) : 禁則・文字組みアキ量設定の listbox 日本語表示、表示順を整理、outlined_text レイヤーを再利用、タイトル変更
                      listbox の項目名・列挙値を英日ローカライズ、「アウトライン」「テキストを復元」パネルを分離（アウトライン化ボタンを最上部へ）、
                      復元オプション「アウトラインデータを残す」「復元したテキストを別レイヤーに」を追加（OFF で削除／同一レイヤーに復元）
                      パネル・listbox に helpTip を追加、UI 文言を「メモ」に統一（英語版は "note"）
                      英語環境では和文専用属性（組み方向・禁則・文字組みアキ量設定・文字ツメ）を保存／表示／復元しない（バイリンガル版は locale 判定、英語版は常時）、概要・コメントの「位置」を「座標」に統一

*/

(function () {

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiTextOutlineRestorePalette";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// ==============================
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

/* 項目名（左列）/ Item labels (left column) */
var LIST_ITEM_LABELS = {
    '文字列':                     { ja: '文字列', en: 'Text' },
    '組み方向':                   { ja: '組み方向', en: 'Orientation' },
    'フォント':                   { ja: 'フォント', en: 'Font' },
    'フォントサイズ':             { ja: 'フォントサイズ', en: 'Font size' },
    '行送り':                     { ja: '行送り', en: 'Leading' },
    '自動行送り':                 { ja: '自動行送り', en: 'Auto leading' },
    '水平比率':                   { ja: '水平比率', en: 'Horizontal scale' },
    '垂直比率':                   { ja: '垂直比率', en: 'Vertical scale' },
    'カーニング':                 { ja: 'カーニング', en: 'Kerning' },
    'プロポーショナルメトリクス': { ja: 'プロポーショナルメトリクス', en: 'Proportional metrics' },
    'トラッキング':               { ja: 'トラッキング', en: 'Tracking' },
    '文字ツメ':                   { ja: '文字ツメ', en: 'Tsume' },
    '行揃え':                     { ja: '行揃え', en: 'Alignment' },
    '禁則':                       { ja: '禁則', en: 'Kinsoku' },
    '文字組み':                   { ja: '文字組み', en: 'Mojikumi' },
    '文字カラー':                 { ja: '文字カラー', en: 'Fill color' }
};

/* 列挙値（右列）。ここに無い値（数値・フォント名・カラー・true/false）は素通し / Enumerated values (right column) */
var LIST_VALUE_LABELS = {
    '縦組み':                 { ja: '縦組み', en: 'Vertical' },
    '横組み':                 { ja: '横組み', en: 'Horizontal' },
    'メトリクス':             { ja: 'メトリクス', en: 'Metrics' },
    'オプティカル':           { ja: 'オプティカル', en: 'Optical' },
    '和文等幅':               { ja: '和文等幅', en: 'Metrics (Roman Only)' },
    'なし':                   { ja: 'なし', en: 'None' },
    '左揃え':                 { ja: '左揃え', en: 'Left' },
    '中央揃え':               { ja: '中央揃え', en: 'Center' },
    '右揃え':                 { ja: '右揃え', en: 'Right' },
    '均等配置':               { ja: '均等配置', en: 'Justify all lines' },
    '均等配置（最終行左）':   { ja: '均等配置（最終行左）', en: 'Justify (last left)' },
    '均等配置（最終行中央）': { ja: '均等配置（最終行中央）', en: 'Justify (last center)' },
    '均等配置（最終行右）':   { ja: '均等配置（最終行右）', en: 'Justify (last right)' },
    '強い禁則':               { ja: '強い禁則', en: 'Hard' },
    '弱い禁則':               { ja: '弱い禁則', en: 'Soft' },
    '弱い禁則 v2':            { ja: '弱い禁則 v2', en: 'Soft v2' },
    '行末約物全角/半角':      { ja: '行末約物全角/半角', en: 'Line end full/half-width' },
    '約物半角':               { ja: '約物半角', en: 'Half-width punctuation' },
    '行末約物半角':           { ja: '行末約物半角', en: 'Line end half-width' },
    '行末約物全角':           { ja: '行末約物全角', en: 'Line end full-width' },
    '約物全角':               { ja: '約物全角', en: 'Full-width punctuation' },
    'ツメ組み':               { ja: 'ツメ組み', en: 'Tight' },
    'ベタ組み':               { ja: 'ベタ組み', en: 'Solid' }
};

/* 項目名を現在言語へ（未知はそのまま）/ Localize a list item label */
function localizeListLabel(jaLabel) {
    var entry = LIST_ITEM_LABELS[jaLabel];
    if (entry && entry[CURRENT_LANG] != null) { return entry[CURRENT_LANG]; }
    return jaLabel;
}

/* 列挙値を現在言語へ（未知＝数値・フォント名・カラー等はそのまま）/ Localize a list value */
function localizeListValue(jaValue) {
    var entry = LIST_VALUE_LABELS[jaValue];
    if (entry && entry[CURRENT_LANG] != null) { return entry[CURRENT_LANG]; }
    return jaValue;
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

/* --- 属性パネルを開閉（メニューコマンド）/ Toggle the Attributes panel via menu command --- */
function workerToggleAttributesPanel() {
    try {
        app.executeMenuCommand("internal palettes posing as plug-in menus-attributes");
    } catch (e) {
    }
    return "OK";
}

/* --- note 文字列を組み立て / Build the memo text from gathered values --- */
/* handleJp が false（英語環境）のときは和文専用属性（組み方向・禁則・文字組み・文字ツメ）を書き出さない */
function workerBuildMemoText(info, handleJp) {
    var jp = (handleJp !== false);
    var memo = "文字列：\n" + info.content + "\n\n" +
        "フォント：\n" + info.fontName + "\n\n" +
        "フォントサイズ：\n" + info.fontSize + "\n\n" +
        "行送り：\n" + info.leading + "\n\n" +
        "カーニング：\n" + info.kerning + "\n\n" +
        "プロポーショナルメトリクス：\n" + info.proportionalMetrics + "\n\n" +
        "トラッキング：\n" + info.tracking + "\n\n";
    if (jp) { memo += "文字ツメ：\n" + info.tsume + "\n\n"; }
    if (jp) { memo += "組み方向：\n" + info.orientation + "\n\n"; }
    memo += "文字カラー：\n" + info.color + "\n\n" +
        "行揃え：\n" + info.justification + "\n\n";
    if (jp) { memo += "禁則：\n" + info.kinsoku + "\n\n"; }
    if (jp) { memo += "文字組み：\n" + info.mojikumi + "\n\n"; }
    memo += "自動行送り：\n" + info.autoLeading + "\n\n" +
        "水平比率：\n" + info.horizontalScale + "\n\n" +
        "垂直比率：\n" + info.verticalScale + "\n\n" +
        "座標：\nL = " + info.left + ", T = " + info.top + ", R = " + info.right + ", B = " + info.bottom;
    return memo;
}

function workerProcessTextFrame(textFrame, handleJp) {
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
        justification: workerJustificationToText(textRange.paragraphAttributes.justification),
        kinsoku: workerKinsokuToText(textRange.paragraphAttributes),
        mojikumi: workerMojikumiToText(textRange.paragraphAttributes),
        autoLeading: textRange.characterAttributes.autoLeading ? "true" : "false",
        horizontalScale: workerRound(textRange.characterAttributes.horizontalScale),
        verticalScale: workerRound(textRange.characterAttributes.verticalScale),
        left: workerRound(bounds[0]),
        top: workerRound(bounds[1]),
        right: workerRound(bounds[2]),
        bottom: workerRound(bounds[3])
    }, handleJp);
    app.activeDocument.selection = null;
    textFrame.selected = true;
    textFrame.createOutline();
    /* createOutline 後の選択が常に期待どおりとは限らないので最低限チェック */
    if (app.activeDocument.selection && app.activeDocument.selection.length > 0) {
        app.activeDocument.selection[0].note = memoText;
    }
    return true;
}

function workerRunOutline(handleJp) {
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
        workerProcessTextFrame(selectedTextFrames[loopIndex], handleJp);
    }
    return "OK:" + selectedTextFrames.length;
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
    var normalizedValue = workerNormalizeMojikumiName(value);
    var presets = workerMojikumiPresets();
    var presetIndex;
    for (presetIndex = 0; presetIndex < presets.length; presetIndex++) {
        if (workerNormalizeMojikumiName(presets[presetIndex].labelText) === normalizedValue) { return presets[presetIndex].labelText; }
    }
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
    var mapIndex;
    for (mapIndex = 0; mapIndex < romajiMap.length; mapIndex++) {
        if (normalizedValue.indexOf(romajiMap[mapIndex].key) === 0) { return romajiMap[mapIndex].name; }
    }
    return value;
}

/* --- note を表示用にコンパクト整形 / Format the note for compact display --- */
function workerFormatNoteForDisplay(noteText, handleJp) {
    var noteLines = noteText.split("\n");
    /* handleJp が false（英語環境）のときは和文専用属性（組み方向・禁則・文字組み・文字ツメ）を一覧に出さない（旧 note に含まれていてもスキップ） */
    var displayLabels = (handleJp !== false)
        ? ["文字列", "組み方向", "フォント", "フォントサイズ", "行送り", "自動行送り", "水平比率", "垂直比率", "カーニング", "プロポーショナルメトリクス", "トラッキング", "文字ツメ", "行揃え", "禁則", "文字組み", "文字カラー"]
        : ["文字列", "フォント", "フォントサイズ", "行送り", "自動行送り", "水平比率", "垂直比率", "カーニング", "プロポーショナルメトリクス", "トラッキング", "行揃え", "文字カラー"];
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
                if (currentLabel === "禁則") { displayValue = workerKinsokuToDisplay(displayValue); }
                if (currentLabel === "文字組み") { displayValue = workerMojikumiToDisplay(displayValue); }
                displayLines.push(currentLabel + "： " + displayValue);
                break;
            }
        }
    }
    return displayLines.join("\n");
}

/* --- 選択状態を検査（テキスト有無・メモ有無・表示用note） / Inspect selection state --- */
function workerInspectSelection(handleJp) {
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
    var formattedNote = workerFormatNoteForDisplay(noteHolder.note, handleJp);
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
        justificationText: null,
        kinsokuText: null,
        mojikumiText: null,
        autoLeading: null,
        horizontalScale: null,
        verticalScale: null,
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
        if (noteLines[lineIndex].indexOf("行揃え：") === 0 && lineIndex + 1 < noteLines.length) { attributes.justificationText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("禁則：") === 0 && lineIndex + 1 < noteLines.length) { attributes.kinsokuText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("文字組み：") === 0 && lineIndex + 1 < noteLines.length) { attributes.mojikumiText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("自動行送り：") === 0 && lineIndex + 1 < noteLines.length) { attributes.autoLeading = noteLines[lineIndex + 1] === "true"; }
        if (noteLines[lineIndex].indexOf("水平比率：") === 0 && lineIndex + 1 < noteLines.length) { attributes.horizontalScale = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("垂直比率：") === 0 && lineIndex + 1 < noteLines.length) { attributes.verticalScale = parseFloat(noteLines[lineIndex + 1]); }
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

/* --- 復元：退避（アウトライン）レイヤーを用意（既存 outlined_text があればロック解除して再利用）
       Restore: reuse the existing outlined_text layer (unlocked) if present, otherwise create a fresh stash layer --- */
function workerCreateOutlineStashLayer(doc) {
    if (!doc) { return null; }
    var stashLayer = null;
    /* 既存の outlined_text レイヤーがあればロックを解除してそのまま退避先に使う（アウトラインを1枚に集約） */
    var findIndex;
    for (findIndex = 0; findIndex < doc.layers.length; findIndex++) {
        if (doc.layers[findIndex].name === "outlined_text") { stashLayer = doc.layers[findIndex]; break; }
    }
    if (!stashLayer) {
        stashLayer = doc.layers.add();
        stashLayer.name = "__outlined_text_stash__";
    }
    workerSetLayerUsable(stashLayer);
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

/* --- 行揃えを保存用テキストへ / Serialize justification to text --- */
function workerJustificationToText(justification) {
    if (justification == Justification.CENTER) { return "中央揃え"; }
    if (justification == Justification.RIGHT) { return "右揃え"; }
    if (justification == Justification.FULLJUSTIFYLASTLINELEFT) { return "均等配置（最終行左）"; }
    if (justification == Justification.FULLJUSTIFYLASTLINECENTER) { return "均等配置（最終行中央）"; }
    if (justification == Justification.FULLJUSTIFYLASTLINERIGHT) { return "均等配置（最終行右）"; }
    if (justification == Justification.FULLJUSTIFY) { return "均等配置"; }
    return "左揃え";
}

/* --- 復元：行揃えテキストを enum へ / Restore: resolve alignment label to Justification --- */
/* 復元先は新規テキストフレームなので既定は LEFT。LEFT は代入が無視されることがあるが既定と一致するため実害なし */
function workerResolveJustification(justificationText) {
    if (justificationText === "中央揃え") { return Justification.CENTER; }
    if (justificationText === "右揃え") { return Justification.RIGHT; }
    if (justificationText === "均等配置（最終行左）") { return Justification.FULLJUSTIFYLASTLINELEFT; }
    if (justificationText === "均等配置（最終行中央）") { return Justification.FULLJUSTIFYLASTLINECENTER; }
    if (justificationText === "均等配置（最終行右）") { return Justification.FULLJUSTIFYLASTLINERIGHT; }
    if (justificationText === "均等配置") { return Justification.FULLJUSTIFY; }
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
        var presets = workerMojikumiPresets();
        var presetIndex;
        for (presetIndex = 0; presetIndex < presets.length; presetIndex++) {
            if (workerNormalizeMojikumiName(presets[presetIndex].labelText) === normalizedText) {
                var idx = presets[presetIndex].mojikumiIndex;
                if (idx >= 0 && idx < sets.length) {
                    textRange.paragraphAttributes.mojikumi = sets[idx];
                    return;
                }
            }
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
function workerCreateRestoredTextFrame(sourceItem, attributes, restoredLayer, restoreReport, handleJp) {
    var jp = (handleJp !== false);
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
    /* 自動行送り ON のときは行送りを明示せず自動計算に任せる（旧 note は autoLeading 未記録なので false 扱い） */
    try {
        if (attributes.autoLeading === true) {
            attrs.autoLeading = true;
        } else {
            attrs.autoLeading = false;
            if (attributes.leading !== null) { attrs.leading = attributes.leading; }
        }
    } catch (eLead) {}
    try { attrs.tracking = attributes.tracking; } catch (eTrack) {}
    /* 水平比率・垂直比率（長体／平体）。旧 note は未記録なので null のときは触らない */
    try { if (attributes.horizontalScale !== null) { attrs.horizontalScale = attributes.horizontalScale; } } catch (eHScale) {}
    try { if (attributes.verticalScale !== null) { attrs.verticalScale = attributes.verticalScale; } } catch (eVScale) {}
    /* 文字ツメは和文専用。英語環境（jp=false）はスキップ（旧 note に含まれていても適用しない） */
    try { if (jp && attributes.tsume !== null) { attrs.Tsume = attributes.tsume; } } catch (eTsume) {}
    try {
        if (attributes.kerningText != null) {
            /* カーニング方式を復元（AutoKerning ロジック：proportionalMetrics は方式に連動） */
            workerApplyKerning(textFrame.textRange, workerResolveKernType(attributes.kerningText));
        } else {
            /* 旧 note（カーニング未記録）は保存済みの proportionalMetrics を復元 */
            attrs.proportionalMetrics = attributes.proportionalMetrics;
        }
    } catch (eKern) {}
    /* 組み方向は和文専用。英語環境はスキップし新規フレーム既定（横組み）のまま */
    try { if (jp && attributes.orientation != null) { textFrame.orientation = (attributes.orientation === "縦組み") ? TextOrientation.VERTICAL : TextOrientation.HORIZONTAL; } } catch (eOri) {}
    /* 行揃え（段落属性）。新規フレームの既定は LEFT なので左揃えは実質そのまま */
    try {
        if (attributes.justificationText != null) {
            textFrame.textRange.paragraphAttributes.justification = workerResolveJustification(attributes.justificationText);
        }
    } catch (eJust) {}
    /* 禁則・文字組みアキ量設定（段落属性）は和文専用。英語環境はスキップ。旧 note は未記録なので null ならスキップ */
    try { if (jp && attributes.kinsokuText != null) { workerApplyKinsoku(textFrame.textRange, attributes.kinsokuText); } } catch (eKinsoku) {}
    try { if (jp && attributes.mojikumiText != null) { workerApplyMojikumi(textFrame.textRange, attributes.mojikumiText); } } catch (eMojikumi) {}
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

/* --- 復元：エントリ / Restore: entry --- */
function workerRestoreText(keepOutline, separateLayer, handleJp) {
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
    /* 既定はアウトラインを残す。false のときだけ退避レイヤーを作らず元アウトラインを削除 */
    var stashOutline = (keepOutline !== false);
    /* 既定は復元テキストを別レイヤー（restored_text）へ。false のときは元アウトラインと同じレイヤーへ置く（null で sourceItem.layer にフォールバック） */
    var useSeparateLayer = (separateLayer !== false);
    var outlinedTextLayer = stashOutline ? workerCreateOutlineStashLayer(doc) : null;
    var restoredTextLayer = useSeparateLayer ? workerCreateRestoredTextLayer(doc) : null;
    var restoreReport = { restored: 0, fontFallback: false };
    /* ループ全体を try で囲み、想定外エラーでも空レイヤーを残さない */
    var runError = null;
    try {
        var restoreIndex;
        for (restoreIndex = 0; restoreIndex < restorableItems.length; restoreIndex++) {
            workerRestoreItem(restorableItems[restoreIndex], outlinedTextLayer, restoredTextLayer, restoreReport, stashOutline, handleJp);
        }
    } catch (eRun) {
        runError = String(eRun);
    }
    if (restoreReport.restored < 1) {
        /* 1件も復元できなかったら（エラー時も含め）作成した退避／復元レイヤーを後始末 */
        if (outlinedTextLayer) { try { outlinedTextLayer.remove(); } catch (eStash) {} }
        if (restoredTextLayer) {
            try {
                if (restoredTextLayer.pageItems.length < 1 && (!restoredTextLayer.layers || restoredTextLayer.layers.length < 1)) { restoredTextLayer.remove(); }
            } catch (eRestored) {}
        }
        return runError ? ("ERR:" + runError) : "NONOTE";
    }
    /* 一部でも成功していれば、その分を確定（途中エラーでも退避レイヤーを宙ぶらりんにしない） */
    if (stashOutline && outlinedTextLayer) { workerFinalizeTemplateLayer(outlinedTextLayer); }
    if (restoredTextLayer) { try { doc.activeLayer = restoredTextLayer; } catch (eActive) {} }
    return "OK:" + restoreReport.restored + (restoreReport.fontFallback ? ":FONT" : "");
}

// ワーカー関数はすべてここに登録（追加漏れ防止）
var WORKER_FUNCS = [
    workerRound,
    workerToggleAttributesPanel,
    workerSetLayerUsable,
    workerBuildMemoText,
    workerProcessTextFrame,
    workerRunOutline,
    workerKinsokuToDisplay,
    workerMojikumiToDisplay,
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
    workerJustificationToText,
    workerResolveJustification,
    workerKinsokuToText,
    workerApplyKinsoku,
    workerNormalizeMojikumiName,
    workerMojikumiPresets,
    workerMojikumiLabelFromApplied,
    workerMojikumiToText,
    workerApplyMojikumi,
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
    var result = callWorker("workerRunOutline(" + (HANDLE_JP ? "true" : "false") + ");");
    isBusy = false;
    applyResultToStatus(win, result, 'status.doneOutline');
    refreshSelectedNote(win, true); // アウトライン化直後のメモを表示
}

function onRestoreClick(win) {
    if (isBusy) return;
    // 復元の直前に「読み込み」ロジックを強制実行し、対象のメモを表示（移動前の状態を見せる）
    refreshSelectedNote(win, true);
    isBusy = true;
    setStatus(win, L('status.busy'));
    // チェックボックスの状態を worker に渡す（残す／別レイヤー）
    var keepOutline = win.keepOutlineCheck ? win.keepOutlineCheck.value : true;
    var separateLayer = win.separateLayerCheck ? win.separateLayerCheck.value : true;
    var result = callWorker("workerRestoreText(" + (keepOutline ? "true" : "false") + ", " + (separateLayer ? "true" : "false") + ", " + (HANDLE_JP ? "true" : "false") + ");");
    isBusy = false;
    applyResultToStatus(win, result, 'status.doneRestore');
}

/* 属性パネルの開閉（メインエンジンにメニューコマンドを委譲）/ Toggle Attributes panel (delegated) */
function onAttributesClick(win) {
    if (isBusy) return;
    callWorker("workerToggleAttributesPanel();", [workerToggleAttributesPanel]);
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
        var row = win.noteList.add("item", localizeListLabel(itemLabel));
        row.subItems[0].text = localizeListValue(itemValue);
    }
}

function refreshSelectedNote(win, keepStatus) {
    if (isBusy) return;
    isBusy = true;
    var result = callWorker("workerInspectSelection(" + (HANDLE_JP ? "true" : "false") + ");");
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

    // アウトライン化 / Outline（最上部）
    var outlinePanel = win.add("panel", undefined, L('panel.outline'));
    outlinePanel.helpTip = L('panel.outlineTip');
    setupPanel(outlinePanel, 6);

    var outlineButtonRow = outlinePanel.add("group");
    setupRow(outlineButtonRow, "left", 8); // ボタン列は左寄せ

    var outlineButton = outlineButtonRow.add("button", undefined, L('button.outline'));
    outlineButton.helpTip = L('button.outlineTip');
    outlineButton.onClick = function () { onOutlineClick(win); };

    // 選択オブジェクトのメモを表示 / Selected object's note
    var selectedObjectPanel = win.add("panel", undefined, L('panel.selected'));
    selectedObjectPanel.helpTip = L('panel.selectedTip');
    setupPanel(selectedObjectPanel, 6);

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

    // 左：属性パネルの開閉
    var attributesButton = noteActionRow.add("button", undefined, L('button.attributes'));
    attributesButton.helpTip = L('button.attributesTip');
    attributesButton.alignment = ["left", "center"];
    attributesButton.onClick = function () { onAttributesClick(win); };

    // 中央：フレキシブルスペーサー（左右のボタンを両端へ押し広げる）
    var noteActionSpacer = noteActionRow.add("group");
    noteActionSpacer.alignment = ["fill", "center"];

    // 右：メモを読み込み
    var loadNoteButton = noteActionRow.add("button", undefined, L('button.load'));
    loadNoteButton.helpTip = L('button.loadTip');
    loadNoteButton.alignment = ["right", "center"];
    loadNoteButton.onClick = function () { refreshSelectedNote(win, false); };

    // テキストを復元 / Restore Text
    var restorePanel = win.add("panel", undefined, L('panel.restore'));
    restorePanel.helpTip = L('panel.restoreTip');
    setupPanel(restorePanel, 6);

    var restoreButtonRow = restorePanel.add("group");
    setupRow(restoreButtonRow, "left", 8); // ボタン列は左寄せ
    restoreButtonRow.margins = [0, 0, 0, 5]; // ボタンの下に余白 +5

    var restoreButton = restoreButtonRow.add("button", undefined, L('button.restore'));
    restoreButton.helpTip = L('button.restoreTip');
    restoreButton.onClick = function () { onRestoreClick(win); };

    // 復元オプション：アウトラインを残す／別レイヤーに復元
    win.keepOutlineCheck = restorePanel.add("checkbox", undefined, L('option.keepOutline'));
    win.keepOutlineCheck.helpTip = L('option.keepOutlineTip');
    win.keepOutlineCheck.value = true; // 既定 ON

    win.separateLayerCheck = restorePanel.add("checkbox", undefined, L('option.separateLayer'));
    win.separateLayerCheck.helpTip = L('option.separateLayerTip');
    win.separateLayerCheck.value = true; // 既定 ON

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
