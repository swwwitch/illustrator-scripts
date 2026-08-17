#targetengine "TextBreakSplitMergeEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

改行・分割・連結・整形に加えて、行の入れ替えや英字のケース変換までを1つのパレットに集約し、選択中のテキストへ即時に適用するツール。
選択状態に応じて、実行できる処理のボタンだけを有効化する。

詳細はREADMEを参照。

*/

/*

### Overview

Brings line-break, split, merge, cleanup, line reordering, and letter-case conversion into a single palette and
applies them to the current selection immediately. Only the buttons available for the current selection stay enabled.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextBreakSplitMergePallete";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-18";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextBreakSplitMergePallete.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextBreakSplitMergePallete.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf6f34559ba46"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 「指定文字で改行」の初期値 / Default characters for "break at specified characters" */
var DEFAULT_BREAK_CHARS = "、。，．｡､,.!?！？";

/* 「指定文字数で改行」の初期値 / Default character count for "break at character count" */
var DEFAULT_BREAK_COUNT = "35";

/* ケース変換プレビューに表示する最大文字数 / Maximum characters shown in the letter-case preview */
var CASE_PREVIEW_MAX_CHARS = 40;

/* 選択ステータスを取り直す最小間隔（ミリ秒）/ Minimum interval between selection status polls */
var STATUS_POLL_INTERVAL_MS = 400;

// =========================================
// 処理パラメーター / Processing parameters
// （メインエンジンでも同じ値を参照する。__WORKER_CONSTANTS 参照）
// =========================================

var LINE_Y_THRESHOLD     = 5;    /* 横連結で同じ行と見なすY座標差（pt）*/
var AUTO_LEADING_RATIO   = 1.2;  /* 行送りが取得できないときのフォントサイズ比 */
var MIN_LEADING_RATIO    = 1.2;  /* 連結後の行送りの下限（フォントサイズ比）*/
var DEFAULT_FONT_SIZE_PT = 12;   /* 文字サイズが取得できないときの既定値（pt）*/

// =========================================
// レイアウト / Layout
// =========================================

var PANEL_MARGINS       = [10, 18, 10, 8];   /* パネル余白 [左,上,右,下] */
var STATUS_MARGINS      = [50, 18, 50, 8];   /* ステータスパネル余白 */
var TAB_MARGINS         = [10, 20, 0, -10];  /* タブ余白 [左,上,右,下] */
var TAB_SPACING         = 15;                /* タブ内の要素間隔 */
var FOOTER_MARGINS      = [10, 10, 10, 0];   /* フッター余白 */
var BUTTON_ROW_MARGINS  = [10, 0, 10, 8];    /* パネル内ボタン列の余白 */
var STATUS_ROW_SPACING  = 16;                /* ステータス2カラムの間隔 */
var STATUS_LABEL_WIDTH  = 72;                /* ステータス右カラムのラベル幅 */
var LINE_LIST_SIZE      = [200, 460];        /* 行リストボックスのサイズ */
var LINE_LIST_FONT_SIZE = 18;                /* 行リストボックスの文字サイズ */
var CASE_BUTTON_SIZE    = [150, 24];         /* ケース変換ボタンのサイズ */
var CASE_PREVIEW_SIZE   = [120, 24];         /* ケース変換プレビューのサイズ */

/**
 * パネルへ共通のレイアウトを適用する
 * @param {Panel} panel - 対象パネル
 * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "center"]）
 * @returns {void}
 */
function setupPanel(panel, alignChildren) {
    panel.margins = PANEL_MARGINS;
    panel.alignment = ["fill", "top"];
    panel.alignChildren = alignChildren || ["fill", "center"];
}

/**
 * ラベル付きパネルを生成し、共通レイアウトを適用する
 * @param {Window|Panel|Group} parent - 追加先のコンテナ
 * @param {string} titleText - パネルのタイトル
 * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "center"]）
 * @returns {Panel} 生成したパネル
 */
function addPanel(parent, titleText, alignChildren) {
    var panel = parent.add("panel", undefined, titleText);
    setupPanel(panel, alignChildren);
    return panel;
}

/**
 * 縦積みのカラムグループを生成する
 * @param {Window|Panel|Group} parent - 追加先のコンテナ
 * @returns {Group} 生成したグループ
 */
function addColumnGroup(parent) {
    var columnGroup = parent.add("group");
    columnGroup.orientation = "column";
    columnGroup.alignment = ["fill", "top"];
    columnGroup.alignChildren = ["fill", "top"];
    return columnGroup;
}

/**
 * タブへ共通のレイアウトを適用する
 * @param {Object} tab - 対象タブ
 * @param {string} orientation - "row" または "column"
 * @param {Array<string>} alignment - タブ自身の整列指定（省略時は ["fill", "top"]）
 * @returns {void}
 */
function setupTab(tab, orientation, alignment) {
    tab.margins = TAB_MARGINS;
    tab.spacing = TAB_SPACING;
    tab.orientation = orientation;
    tab.alignment = alignment || ["fill", "top"];
    tab.alignChildren = ["fill", "top"];
}

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在の言語を判定（ロケールが ja 始まりなら日本語）/ Detect UI language (Japanese if locale starts with "ja") */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var uiLang = getCurrentLang();

/* 日英ラベル定義 / Bilingual labels */
var LABELS = {
    dialog: {
        title: { ja: "テキスト処理", en: "Text Processing" }
    },
    tab: {
        basic: { ja: "基本", en: "Basic" },
        cleanup: { ja: "整形", en: "Cleanup" },
        lineArrange: { ja: "行の編集", en: "Line Edit" },
        alnum: { ja: "英数字", en: "Alphanumeric" }
    },
    panel: {
        breakGroup: { ja: "改行", en: "Breaks" },
        removeBreak: { ja: "削除", en: "Remove" },
        insertBreak: { ja: "挿入", en: "Insert" },
        convertBreak: { ja: "切り換え", en: "Convert" },
        splitGroup: { ja: "分割", en: "Split" },
        splitByBreak: { ja: "改行で分割", en: "Split by Line Breaks" },
        splitByChar: { ja: "文字で分割", en: "Split by Character" },
        sort: { ja: "ソート", en: "Sort" },
        lineEdit: { ja: "編集", en: "Edit" },
        lineDelete: { ja: "行削除", en: "Delete Lines" },
        convert: { ja: "変換", en: "Convert" },
        list: { ja: "リストの除去", en: "Remove List" },
        concat: { ja: "連結", en: "Concatenate" },
        tab: { ja: "タブ", en: "Tab" },
        space: { ja: "スペース削除", en: "Remove Spaces" },
        status: { ja: "ステータス", en: "Status" },
        letterCase: { ja: "大文字{slash}小文字", en: "Letter Case" },
        symbolConvert: { ja: "スペースや記号の変換", en: "Spaces & Symbols" },
        symbolBefore: { ja: "変換前", en: "Before" },
        symbolAfter: { ja: "変換後", en: "After" },
        addSpace: { ja: "スペース追加", en: "Add Space" }
    },
    radio: {
        space: { ja: "スペース", en: "Space" },
        underscore: { ja: "アンダースコア", en: "Underscore" },
        hyphen: { ja: "ハイフン", en: "Hyphen" }
    },
    button: {
        flattenToOneLine: { ja: "すべて1行に", en: "Merge All into One Line" },
        removeLineBreaks: { ja: "改行のみ", en: "Line Breaks Only" },
        addLineBreaks: { ja: "1文字ごとに改行", en: "Insert Line Break After Each Character" },
        breakAtChars: { ja: "指定文字で改行", en: "At Specified Characters" },
        breakAtCount: { ja: "指定文字数で改行", en: "At Character Count" },
        convertBreaks: { ja: "強制改行→改行", en: "Forced Breaks to Paragraph Breaks" },
        convertToForcedBreaks: { ja: "改行→強制改行", en: "Paragraph Breaks to Forced Breaks" },
        splitByLine: { ja: "テキストばらし", en: "Split by Line Breaks" },
        splitByLineKeepStyle: { ja: "〃（書式保持）", en: "Split by Line Breaks (Keep Style)" },
        splitByTab: { ja: "タブで分割", en: "Split by Tabs" },
        splitKeepStyle: { ja: "書式を保持", en: "Keep Style" },
        splitIgnoreStyle: { ja: "書式を無視", en: "Ignore Style" },
        concatV: { ja: "縦方向に連結", en: "Vertical" },
        concatHOnly: { ja: "横連結（行維持）", en: "Merge Horizontally (Keep Rows)" },
        concatH: { ja: "横連結（行統合）", en: "Merge Horizontally (Merge Rows)" },
        concatToArea: { ja: "PDFテキスト整形", en: "Format PDF Text" },
        removeTabs: { ja: "タブを削除", en: "Remove Tabs" },
        tabsToSpaces: { ja: "タブ→スペース", en: "Tabs to Spaces" },
        trimSpaces: { ja: "行頭行末", en: "Leading{slash}Trailing Spaces" },
        cjkLatinSpaces: { ja: "和欧間", en: "Remove Spaces Between CJK and Latin" },
        collapseSpaces: { ja: "連続", en: "Collapse Spaces" },
        cleanupSpaces: { ja: "まとめて", en: "All at Once" },
        removeAllSpaces: { ja: "すべて", en: "Remove All" },
        fullToHalfAlnum: { ja: "全角英数字→半角", en: "Fullwidth to Halfwidth" },
        halfToFullKana: { ja: "半角カナ→全角", en: "Halfwidth Kana to Fullwidth" },
        bulletList: { ja: "箇条書き", en: "Bullet List" },
        numberList: { ja: "番号リスト", en: "Number List" },
        lineUp: { ja: "上へ", en: "Up" },
        lineDown: { ja: "下へ", en: "Down" },
        lineAdd: { ja: "追加", en: "Add" },
        lineEdit: { ja: "編集", en: "Edit" },
        lineDelete: { ja: "削除", en: "Delete" },
        sortByCharCode: { ja: "ソート", en: "Sort" },
        sortByLength: { ja: "文字数順", en: "Sort (Length)" },
        reverseOrder: { ja: "反転", en: "Reverse Order" },
        removeDuplicateLines: { ja: "重複行", en: "Remove Duplicates" },
        removeEmptyLines: { ja: "空行", en: "Remove Empty Lines" },
        caseUpper: { ja: "すべて大文字に", en: "UPPERCASE" },
        caseLower: { ja: "すべて小文字に", en: "lowercase" },
        caseWord: { ja: "単語の先頭を大文字", en: "Capitalize Words" },
        caseSentence: { ja: "文頭のみ大文字", en: "Sentence case" },
        caseTitle: { ja: "英語タイトル形式", en: "Title Case" },
        convertSymbol: { ja: "変換", en: "Convert" },
        spaceAfterPunct: { ja: ".と,の後", en: "Space After . and ," },
        showHiddenChar: { ja: "制御文字の表示{slash}非表示", en: "Show{slash}Hide Hidden Characters" }
    },
    checkbox: {
        includeForcedBreaks: { ja: "強制改行を含む", en: "Include Forced Breaks" },
        forcedBreak: { ja: "強制改行", en: "Forced Break" }
    },
    tooltip: {
        concatV: { ja: "上→下に連結", en: "Merge top to bottom" },
        concatHOnly: {
            ja: "行ごとに横連結して、行は維持します",
            en: "Merge horizontally within each row and keep the rows separate"
        },
        concatH: {
            ja: "横方向に連結した後、複数行を1つのテキストに統合します",
            en: "Merge horizontally and then combine multiple rows into a single text"
        },
        concatToArea: {
            ja: "横方向に連結し、エリア内文字として整形します",
            en: "Merge horizontally and format the result as area text"
        },
        removeLineBreaks: {
            ja: "段落改行を削除します（「強制改行を含む」ON で強制改行も対象）",
            en: "Remove paragraph breaks (also forced breaks when \"Include Forced Breaks\" is on)"
        },
        splitByLine: {
            ja: "改行ごとに別々のテキストフレームへ分割します",
            en: "Split into separate text frames at each line break"
        },
        splitByTab: {
            ja: "タブ位置で分割します（Option+クリックでグループ化しない）",
            en: "Split at tab positions (Option-click to leave the results ungrouped)"
        },
        splitByLineKeepStyle: {
            ja: "改行ごとに分割し、文字書式と位置を保持します",
            en: "Split at each line break while keeping character formatting and position"
        },
        splitKeepStyle: {
            ja: "1文字ずつ別フレームに分割し、書式を保持します",
            en: "Split into one frame per character, keeping formatting"
        },
        splitIgnoreStyle: {
            ja: "1文字ずつ別フレームに分割し、書式をリセットします",
            en: "Split into one frame per character, resetting formatting"
        },
        trimSpaces: {
            ja: "各行の行頭・行末のスペースを削除します",
            en: "Remove leading and trailing spaces on each line"
        },
        cjkLatinSpaces: {
            ja: "和文と欧文の間のスペースを削除します（欧文単語間は保持）",
            en: "Remove spaces between CJK and Latin (spaces within Latin words are kept)"
        },
        collapseSpaces: {
            ja: "連続したスペースを1つにまとめます",
            en: "Collapse consecutive spaces into a single space"
        },
        cleanupSpaces: {
            ja: "行頭行末・和欧間・連続スペースをまとめて処理します",
            en: "Apply trim, CJK/Latin, and collapse in one step"
        },
        removeAllSpaces: {
            ja: "すべてのスペース（半角・全角）を削除します",
            en: "Remove all spaces (half-width and full-width)"
        },
        spaceAfterPunct: {
            ja: "半角ピリオド・カンマの直後にスペースを挿入します",
            en: "Insert a space right after a period or comma"
        },
        bulletList: {
            ja: "行頭の箇条書き記号（・ ･ · • ◦ ● ○ ◎ □ ■ ◆ ◇ ✓ - *）を削除します",
            en: "Remove leading bullet markers (・ ･ · • ◦ ● ○ ◎ □ ■ ◆ ◇ ✓ - *)"
        },
        numberList: {
            ja: "行頭の番号（1. ① a. 一. など）を削除します",
            en: "Remove leading numbering (1. ① a. etc.)"
        }
    },
    prompt: {
        addLine: { ja: "追加する行を入力してください", en: "Enter the line to add" },
        editLine: { ja: "行を編集してください", en: "Edit the line" }
    },
    confirm: {
        deleteLine: { ja: "選択した行を削除しますか？", en: "Delete the selected line?" }
    },
    message: {
        processFailed: {
            ja: "処理中にエラーが発生しました。",
            en: "An error occurred while processing."
        },
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: {
            ja: "テキストフレーム、またはテキストを含むグループを選択してください。",
            en: "Please select a text frame or a group containing text."
        },
        noTextFrames: {
            ja: "選択内に対象のテキストフレームが見つかりません。テキストフレーム、またはテキストを含むグループを選択してください。",
            en: "No target text frames were found in the selection. Please select a text frame or a group containing text."
        }
    },
    info: {
        targetCount: { ja: "対象テキスト", en: "Target Texts" },
        pointCount: { ja: "ポイント文字", en: "Point Type" },
        areaCount: { ja: "エリア内文字", en: "Area Text" },
        paragraphBreak: { ja: "改行", en: "Paragraph Breaks" },
        forcedBreak: { ja: "強制改行", en: "Forced Breaks" },
        tab: { ja: "タブ", en: "Tabs" }
    }
};

/**
 * ラベルノードから現在言語の文言を返す（{slash} は / に展開）
 * @param {Object} labelNode - LABELS 内の { ja, en } ノード
 * @returns {string} 現在言語の文言
 */
function getLabel(labelNode) {
    if (!labelNode) return "";
    var labelString = labelNode[uiLang] || labelNode.ja || labelNode.en || "";
    return labelString.replace(/\{slash\}/g, "/");
}

/**
 * コロン付きラベルを返す（日本語は全角コロン、英語は半角コロン）
 * @param {Object} labelNode - LABELS 内の { ja, en } ノード
 * @returns {string} コロンを付けた文言
 */
function getColonLabel(labelNode) {
    return getLabel(labelNode) + (uiLang === "ja" ? "：" : ":");
}

/**
 * 処理エラーをアラートで通知する
 * @param {Error|Object} err - 例外オブジェクト、または message を持つオブジェクト
 * @returns {void}
 */
function showError(err) {
    var errorMessage = (err && err.message) ? err.message : String(err);
    alert(getLabel(LABELS.message.processFailed) + "\n\n" + errorMessage);
}

// =========================================
// テキスト処理ユーティリティ / Text utilities
// =========================================
/* ここから下の関数群は BridgeTalk でメインエンジンへ委譲される（__LIB_FUNCS 参照）。
   toString() の出力が壊れるため、JSDoc と // 形式のコメントは使わない。 */

function debugLog(context, err) {
    var logMessage = "[TextBreakSplitMerge] " + context;
    if (err) logMessage += " :: " + (err.message ? err.message : String(err));
    try {
        $.writeln(logMessage);
    } catch (e) { }
}

/* 段落改行の表記ゆれ（CRLF）だけを CR に揃える。
   強制改行（U+0003 / LF / U+2028）は行の途中の文字として保持し、段落改行に化けさせない */
function normalizeParagraphBreaks(txt) {
    return String(txt || "").replace(/\r\n/g, "\r");
}

function splitParagraphLines(txt) {
    return normalizeParagraphBreaks(txt).split("\r");
}

function trimLineSpaces(txt) {
    return String(txt || "").replace(/^[ \t　]+/, "").replace(/[ \t　]+$/, "");
}

function isBlankLine(txt) {
    return trimLineSpaces(txt) === "";
}

function stripTrailingBreaks(txt) {
    var result = String(txt || "");
    while (result.length > 0 && isAnyBreak(result.charAt(result.length - 1))) {
        result = result.substring(0, result.length - 1);
    }
    return result;
}

function isLatinLetterOrDigit(c) {
    if (!c) return false;
    var code = c.charCodeAt(0);
    if (code >= 0x41 && code <= 0x5A) return true;
    if (code >= 0x61 && code <= 0x7A) return true;
    if (code >= 0x30 && code <= 0x39) return true;
    return false;
}

function isAsciiTextOnly(txt) {
    return /^[\x00-\x7F]+$/.test(String(txt || ""));
}

function isSentenceEndingJP(txt) {
    return /[。！？]$/.test(String(txt || ""));
}

function isSentenceEndingEN(txt) {
    return /[.!?]$/.test(String(txt || ""));
}

function shouldInsertParagraphBreakBetweenLines(currentText) {
    var compact = String(currentText || "").replace(/[\s\r\n]/g, "");
    return isSentenceEndingJP(currentText) || (isSentenceEndingEN(currentText) && !isAsciiTextOnly(compact));
}

function getCharCodeSafe(ch) {
    if (ch == null || ch === "") return -1;
    return String(ch).charCodeAt(0);
}

function isParagraphBreak(codeOrChar) {
    var code = (typeof codeOrChar === "number") ? codeOrChar : getCharCodeSafe(codeOrChar);
    return code === 13;
}

/* 強制改行は contents の取り方によって U+0003 / LF / U+2028 のいずれかで現れる */
function isForcedBreak(codeOrChar) {
    var code = (typeof codeOrChar === "number") ? codeOrChar : getCharCodeSafe(codeOrChar);
    return code === 3 || code === 10 || code === 8232;
}

function isAnyBreak(codeOrChar) {
    return isParagraphBreak(codeOrChar) || isForcedBreak(codeOrChar);
}

function isTabChar(codeOrChar) {
    var code = (typeof codeOrChar === "number") ? codeOrChar : getCharCodeSafe(codeOrChar);
    return code === 9;
}

/* 改行・タブ以外の文字が残っている最後の位置を返す（無ければ -1）*/
function findLastVisibleIndex(txt) {
    var text = String(txt || "");
    for (var i = text.length - 1; i >= 0; i--) {
        var currentChar = text.charAt(i);
        if (!isAnyBreak(currentChar) && !isTabChar(currentChar)) return i;
    }
    return -1;
}

/* 英字ケース変換（純粋な文字列関数 / TextNormalize.jsx より移植）。
   プレビュー（パレット側）と適用（メインエンジン側）の両方で使う。 */

/* 単語の先頭のみ大文字。事前に小文字化しているので、すべて大文字の語も Negotiable のようになる */
function toWordCap(text) {
    return String(text).toLowerCase().replace(/\b([a-z])/g, function (matched, initial) {
        return initial.toUpperCase();
    });
}

/* 文頭のみ大文字。文区切りが無い（英単語が1つだけの）場合も先頭を大文字化する */
function toSentenceCase(text) {
    return String(text).toLowerCase().replace(/(^|[\.\!\?]\s+|[\r\n]+)([a-z])/g,
        function (matched, prefix, initial) { return prefix + initial.toUpperCase(); });
}

/* 英語タイトル形式（冠詞・前置詞などは小文字 / John Gruber の Title Caps 移植）*/
function toTitleCase(text) {
    var smallWords = "(a|abaft|aboard|about|above|absent|across|afore|after|against|along|alongside|amid|amidst|among|amongst|an|and|apropos|around|as|aside|astride|at|athwart|atop|barring|before|behind|below|beneath|beside|besides|between|betwixt|beyond|but|by|circa|concerning|despite|down|during|except|excluding|failing|following|for|from|given|in|including|inside|into|lest|like|mid|midst|minus|modulo|near|next|nor|notwithstanding|of|off|on|onto|opposite|or|out|outside|over|pace|per|plus|pro|qua|regarding|round|sans|save|than|that|the|through|throughout|till|times|to|toward|towards|under|underneath|unlike|until|unto|up|upon|versus|via|vice|with|within|without|worth|v[.]?|via|vs[.]?)";
    var punctuation = "([!\"#$%&'()*+,./:;<=>?@[\\\\\\]^_`{|}~-]*)";

    function toLowerWord(word) { return word.toLowerCase(); }
    function capitalizeWord(word) { return word.substr(0, 1).toUpperCase() + word.substr(1); }

    var title = String(text);
    var sentenceSplitter = /[:.;?!] |(?: |^)[\"Ò]/g;
    var segments = [];
    var segmentStart = 0;
    while (true) {
        var splitMatch = sentenceSplitter.exec(title);
        segments.push(
            title.substring(segmentStart, splitMatch ? splitMatch.index : title.length)
                .replace(/\b([A-Za-z][a-z.'Õ]*)\b/g, function (matched) {
                    return /[A-Za-z]\.[A-Za-z]/.test(matched) ? matched : capitalizeWord(matched);
                })
                .replace(RegExp("\\b" + smallWords + "\\b", "ig"), toLowerWord)
                .replace(RegExp("^" + punctuation + smallWords + "\\b", "ig"), function (matched, leadingPunct, word) {
                    return leadingPunct + capitalizeWord(word);
                })
                .replace(RegExp("\\b" + smallWords + punctuation + "$", "ig"), capitalizeWord)
        );
        segmentStart = sentenceSplitter.lastIndex;
        if (!splitMatch) break;
        segments.push(splitMatch[0]);
    }
    return segments.join("")
        .replace(/ V(s?)\. /ig, " v$1. ")
        .replace(/(['Õ])S\b/ig, "$1s")
        .replace(/\b(AT&T|Q&A)\b/ig, function (matched) { return matched.toUpperCase(); });
}

(function () {
    /* 重複起動ガード：既にパレットが開いていれば前面化して終了（二重起動を防ぐ）
       閉じると onClose で参照が null に戻るため、非 null＝表示中とみなす */
    try {
        var existingPalette = $.global.__TextBreakSplitMergePalette;
        if (existingPalette) {
            try { existingPalette.show(); } catch (_) { }
            try { existingPalette.active = true; } catch (_) { }
            return;
        }
    } catch (_) { }

    /* ドキュメントが開かれていない場合は処理を終了 / Abort when no document is open */
    if (app.documents.length === 0) {
        alert(getLabel(LABELS.message.noDocument));
        return;
    }

    /* 選択オブジェクトを取得 / Get the current selection */
    var selectedObjects = app.selection;
    if (!selectedObjects || !selectedObjects.length) {
        alert(getLabel(LABELS.message.noSelection));
        return;
    }

    /* 初期選択からテキストフレームを解決 / Resolve text frames from the initial selection */
    selectedObjects = getTextFrames(selectedObjects);
    if (selectedObjects.length === 0) {
        alert(getLabel(LABELS.message.noTextFrames));
        return;
    }

    /* 選択・テキストフレーム操作（メインエンジンへ委譲される関数群）*/

    /* テキストフレームのみ抽出（グループと TextRange も再帰的にたどる）*/
    function getTextFrames(objects) {
        var frames = [];

        function pushUnique(frame) {
            for (var i = 0; i < frames.length; i++) {
                if (frames[i] === frame) return;
            }
            frames.push(frame);
        }

        function collect(item) {
            if (!item) return;

            var typeName, textRangeOwner = null;
            try {
                if (item.isValid === false) return;
                typeName = item.typename || "";
                /* Illustrator ではテキスト選択が TextRange として返ることがある */
                if (typeName === "TextRange" && item.parent && item.parent.typename === "TextFrame") {
                    textRangeOwner = item.parent;
                }
            } catch (e) { debugLog("getTextFrames: inspect item", e); return; }

            if (typeName === "TextFrame") {
                pushUnique(item);
            } else if (textRangeOwner) {
                pushUnique(textRangeOwner);
            } else if (typeName === "GroupItem") {
                try {
                    for (var i = 0; i < item.pageItems.length; i++) collect(item.pageItems[i]);
                } catch (e) { debugLog("getTextFrames: walk group", e); }
            }
        }

        if (!objects) return frames;
        if (typeof objects.length === "number" && !objects.typename) {
            for (var k = 0; k < objects.length; k++) collect(objects[k]);
        } else {
            collect(objects);
        }
        return frames;
    }

    /* ポイント文字／エリア内文字の件数を集計（frames は解決済みのテキストフレーム配列）*/
    function countTextFrameTypes(frames) {
        var pointCount = 0;
        var areaCount = 0;

        for (var i = 0; i < frames.length; i++) {
            if (frames[i].kind === TextType.AREATEXT) {
                areaCount++;
            } else if (frames[i].kind === TextType.POINTTEXT) {
                pointCount++;
            }
        }

        return { total: frames.length, point: pointCount, area: areaCount };
    }

    /* テキストタイプ判定（混在は "mixed"）。frames は解決済みのテキストフレーム配列 */
    function detectTextFrameType(frames) {
        var counts = countTextFrameTypes(frames);
        if (counts.point > 0 && counts.area > 0) return "mixed";
        return (counts.area > 0) ? "area" : "point";
    }

    /* 改行数とタブ数を集計（frames は解決済みのテキストフレーム配列）。
       1文字ずつDOM経由で読むと選択が大きいときに待たされるため、contents を1回だけ読んで数える */
    function countBreakTypes(frames) {
        var paragraphBreakCount = 0;
        var forcedBreakCount = 0;
        var tabCount = 0;

        for (var i = 0; i < frames.length; i++) {
            var txt = String(frames[i].contents || "");
            for (var j = 0; j < txt.length; j++) {
                var code = txt.charCodeAt(j);
                if (isParagraphBreak(code)) {
                    paragraphBreakCount++;
                } else if (isForcedBreak(code)) {
                    forcedBreakCount++;
                } else if (isTabChar(code)) {
                    tabCount++;
                }
            }
        }

        return { paragraph: paragraphBreakCount, forced: forcedBreakCount, tab: tabCount };
    }

    /* 選択状態の集計ヘルパー（__dispatch / __stateDispatch / showPalette で共用）*/

    /* 2行以上（段落改行あり）のフレームが1つでもあるか */
    function hasMultipleLines(frames) {
        for (var i = 0; i < frames.length; i++) {
            if (splitParagraphLines(frames[i].contents).length >= 2) return true;
        }
        return false;
    }

    /* 半角/タブ/全角スペースを含むフレームが1つでもあるか */
    function hasSpacesOrTabs(frames) {
        for (var i = 0; i < frames.length; i++) {
            if (/[ \t　]/.test(frames[i].contents)) return true;
        }
        return false;
    }

    /* 選択状態をまとめたオブジェクトを返す（UI 反映用）。
       集計ごとに選択をたどり直すと遅いので、テキストフレームは一度だけ解決して使い回す */
    function computeSelectionState(objects) {
        var frames = getTextFrames(objects);
        var counts = countTextFrameTypes(frames);
        var breaks = countBreakTypes(frames);
        return {
            total: counts.total, point: counts.point, area: counts.area,
            para: breaks.paragraph, forced: breaks.forced, tab: breaks.tab,
            multiLines: hasMultipleLines(frames),
            multiFrames: frames.length >= 2,
            hasSpTab: hasSpacesOrTabs(frames)
        };
    }

    /* 選択状態を "|" 区切り文字列へエンコード（BridgeTalk の戻り値用）*/
    function encodeSelectionState(objects) {
        var state = computeSelectionState(objects);
        return [state.total, state.point, state.area, state.para, state.forced, state.tab,
        state.multiLines ? 1 : 0, state.multiFrames ? 1 : 0, state.hasSpTab ? 1 : 0].join("|");
    }

    /* 各テキストフレームのcontentsを変換する共通処理 */
    function transformContents(objects, transformFunc) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            frames[i].contents = transformFunc(frames[i].contents);
        }
    }

    /* 各テキストフレームの文字を末尾から走査し、matchesCharCode が真の文字を処理する共通ヘルパー。
       replacement が null なら削除、文字列なら差し替える（末尾走査なので remove でも index がずれない）*/
    function mutateMatchingChars(objects, matchesCharCode, replacement) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            var chars = frames[i].characters;
            for (var charIndex = chars.length - 1; charIndex >= 0; charIndex--) {
                if (!matchesCharCode(chars[charIndex].contents.charCodeAt(0))) continue;
                if (replacement === null) {
                    chars[charIndex].remove();
                } else {
                    chars[charIndex].contents = replacement;
                }
            }
        }
    }

    /* 強制改行（charCode 3 または 10）を削除する共通処理 */
    function removeForcedLineBreaks(objects) {
        mutateMatchingChars(objects, isForcedBreak, null);
    }

    /* 元配列を変更せずに並べ替えた配列を返す */
    function sortedCopy(items, comparator) {
        var sorted = items.slice(0);
        sorted.sort(comparator);
        return sorted;
    }

    /* 上から順にソート（Y降順、同じYならX昇順） */
    function sortByPosition(items) {
        return sortedCopy(items, function (a, b) {
            if (b.position[1] !== a.position[1]) return b.position[1] - a.position[1];
            return a.position[0] - b.position[0];
        });
    }

    /* Y座標で降順ソート */
    function sortByY(items) {
        return sortedCopy(items, function (a, b) { return b.position[1] - a.position[1]; });
    }

    /* X座標で昇順ソート */
    function sortByX(items) {
        return sortedCopy(items, function (a, b) { return a.position[0] - b.position[0]; });
    }

    /* Y位置で行グループ化 */
    function groupByLineY(sortedItems, threshold) {
        var lines = [];
        for (var i = 0; i < sortedItems.length; i++) {
            var y = sortedItems[i].position[1];
            var found = false;
            for (var j = 0; j < lines.length; j++) {
                if (Math.abs(lines[j][0].position[1] - y) <= threshold) {
                    lines[j].push(sortedItems[i]);
                    found = true;
                    break;
                }
            }
            if (!found) {
                lines.push([sortedItems[i]]);
            }
        }
        return lines;
    }

    /* 複数アイテムを内包するバウンディングボックス [左, 上, 右, 下] を返す */
    function getUnionBounds(items) {
        var left = items[0].visibleBounds[0];
        var top = items[0].visibleBounds[1];
        var right = items[0].visibleBounds[2];
        var bottom = items[0].visibleBounds[3];
        for (var i = 1; i < items.length; i++) {
            var bounds = items[i].visibleBounds;
            if (bounds[0] < left) left = bounds[0];
            if (bounds[1] > top) top = bounds[1];
            if (bounds[2] > right) right = bounds[2];
            if (bounds[3] < bottom) bottom = bounds[3];
        }
        return [left, top, right, bottom];
    }

    /* テキストフレームを1つのグループにまとめる */
    function groupTextFrames(frames, targetLayer) {
        var validFrames = getTextFrames(frames);
        if (validFrames.length === 0) return [];

        var targetGroup = (targetLayer || app.activeDocument.activeLayer).groupItems.add();
        for (var i = 0; i < validFrames.length; i++) {
            try {
                validFrames[i].move(targetGroup, ElementPlacement.PLACEATEND);
            } catch (e) { debugLog("groupTextFrames: move to group", e); }
        }
        return [targetGroup];
    }

    /* 改行系の関数 */

    /* 改行文字を削除する関数 */
    function removeLineBreaks(objects) {
        mutateMatchingChars(objects, isParagraphBreak, null);
    }

    /* 強制改行と改行を削除する関数 */
    function removeAllBreaks(objects) {
        mutateMatchingChars(objects, isAnyBreak, null);
    }

    /* 改行と強制改行を取り除いて1行にする関数。
       改行をはさんで欧文どうしが並んでいた場合は、語がくっつかないようスペースに置き換える */
    function joinBreaksToOneLine(objects) {
        transformContents(objects, function (txt) {
            var result = "";
            for (var i = 0; i < txt.length; i++) {
                var currentChar = txt.charAt(i);
                if (!isAnyBreak(currentChar)) {
                    result += currentChar;
                    continue;
                }

                /* 改行が続く場合はまとめて1か所として扱う */
                var nextIndex = i + 1;
                while (nextIndex < txt.length && isAnyBreak(txt.charAt(nextIndex))) nextIndex++;

                var prevChar = (result.length > 0) ? result.charAt(result.length - 1) : "";
                var nextChar = (nextIndex < txt.length) ? txt.charAt(nextIndex) : "";
                if (isLatinLetterOrDigit(prevChar) && isLatinLetterOrDigit(nextChar)) result += " ";
                i = nextIndex - 1;
            }
            return result;
        });
    }

    /* 複数フレームを連結し、改行を取り除いて1行に統合する関数 */
    function flattenToOneLine(objects) {
        var frames = getTextFrames(objects);
        if (frames.length < 2) {
            removeEmptyLines(frames);
            joinBreaksToOneLine(frames);
            return frames;
        }
        var result = concatVertical(objects);
        var targets = result && result.length ? result : frames;
        removeEmptyLines(targets);
        joinBreaksToOneLine(targets);
        return result;
    }

    /* 空行を削除する関数 */
    function removeEmptyLines(objects) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            var kept = [];
            for (var i = 0; i < lines.length; i++) {
                if (!isBlankLine(lines[i])) {
                    kept.push(lines[i]);
                }
            }
            return kept.join("\r");
        });
    }

    /* タブを削除する関数 */
    function removeTabs(objects) {
        mutateMatchingChars(objects, isTabChar, null);
    }

    /* タブをスペースに変換する関数 */
    function tabsToSpaces(objects) {
        transformContents(objects, function (txt) {
            return txt.replace(/\t/g, " ");
        });
    }

    /* 行頭行末のスペースを削除する関数 */
    function trimSpaces(objects) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            for (var i = 0; i < lines.length; i++) {
                lines[i] = trimLineSpaces(lines[i]);
            }
            return lines.join("\r");
        });
    }

    /* 連続スペースを1つにまとめる関数 */
    function collapseSpaces(objects) {
        transformContents(objects, function (txt) {
            /* 半角スペース連続 → 半角スペース1つ  */
            var result = txt.replace(/ {2,}/g, " ");
            /* 全角スペース連続 → 全角スペース1つ  */
            result = result.replace(/\u3000{2,}/g, "\u3000");
            return result;
        });
    }

    /* 各行の行頭を正規表現で除去する共通処理。
       prefixPatterns は配列で渡し、行ごとに最初に一致したものだけを適用する */
    function removeLinePrefix(objects, prefixPatterns) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            for (var i = 0; i < lines.length; i++) {
                for (var j = 0; j < prefixPatterns.length; j++) {
                    if (!prefixPatterns[j].test(lines[i])) continue;
                    lines[i] = lines[i].replace(prefixPatterns[j], "");
                    break;
                }
            }
            return lines.join("\r");
        });
    }

    /* 全角英数字を半角へ変換した文字列を返す
       対象は ０-９ (0xFF10-0xFF19) / Ａ-Ｚ (0xFF21-0xFF3A) / ａ-ｚ (0xFF41-0xFF5A) */
    function toHalfWidthAlnumText(txt) {
        var result = "";
        for (var i = 0; i < txt.length; i++) {
            var code = txt.charCodeAt(i);
            if (code >= 0xFF10 && code <= 0xFF19) {
                result += String.fromCharCode(code - 0xFF10 + 0x30);
            } else if (code >= 0xFF21 && code <= 0xFF3A) {
                result += String.fromCharCode(code - 0xFF21 + 0x41);
            } else if (code >= 0xFF41 && code <= 0xFF5A) {
                result += String.fromCharCode(code - 0xFF41 + 0x61);
            } else {
                result += txt.charAt(i);
            }
        }
        return result;
    }

    /* 半角カナを全角カナへ変換した文字列を返す（濁点・半濁点の合成と約物も対象）*/
    function toFullWidthKanaText(txt) {
        var halfKana = "ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ";
        var fullKana = "ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン";
        /* 濁点・半濁点の合成対応表 */
        var dakutenBase = "ｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾊﾋﾌﾍﾎ";
        var dakutenFull = "ガギグゲゴザジズゼゾダヂヅデドバビブベボ";
        var handakutenBase = "ﾊﾋﾌﾍﾎ";
        var handakutenFull = "パピプペポ";
        /* 単独で置き換える約物（濁点・半濁点・句読点・カギ括弧・中黒）*/
        var halfPunct = "ﾞﾟ｡｢｣､･";
        var fullPunct = "゛゜。「」、・";

        var result = "";
        for (var i = 0; i < txt.length; i++) {
            var currentChar = txt.charAt(i);
            var nextChar = (i + 1 < txt.length) ? txt.charAt(i + 1) : "";

            /* 直後が濁点・半濁点なら合成して1文字にする */
            var isDakuten = (nextChar === "ﾞ");
            var comboIndex = -1;
            if (isDakuten) comboIndex = dakutenBase.indexOf(currentChar);
            else if (nextChar === "ﾟ") comboIndex = handakutenBase.indexOf(currentChar);
            if (comboIndex >= 0) {
                result += isDakuten ? dakutenFull.charAt(comboIndex) : handakutenFull.charAt(comboIndex);
                i++;
                continue;
            }

            var kanaIndex = halfKana.indexOf(currentChar);
            var punctIndex = halfPunct.indexOf(currentChar);
            if (kanaIndex >= 0) {
                result += fullKana.charAt(kanaIndex);
            } else if (punctIndex >= 0) {
                result += fullPunct.charAt(punctIndex);
            } else {
                result += currentChar;
            }
        }
        return result;
    }

    /* 全角英数字を半角に変換する関数 */
    function fullToHalfAlnum(objects) {
        transformContents(objects, toHalfWidthAlnumText);
    }

    /* 半角カナを全角カナに変換する関数 */
    function halfToFullKana(objects) {
        transformContents(objects, toFullWidthKanaText);
    }

    /* 行頭の箇条書き記号を除去する関数。
       AddBulletsAndNumbers.jsx が付ける「タブ + 記号 + タブ」形式と手入力の両方に対応する。
       中黒は異体字（・ ･ ·）も対象。「-」「*」は直後が空白のときだけ除去する（-5℃ のような本文を守る）*/
    function removeBulletMarkers(objects) {
        removeLinePrefix(objects, [
            /^\t[・･·•◦●○◎□■◆◇✓\-\*]\t/,
            /^[\t 　]*(?:[・･·•◦●○◎□■◆◇✓]|[\-\*](?=[\t 　]))[\t 　]*/
        ]);
    }

    /* 行頭の番号を除去する関数。
       数字・全角数字・丸数字・ABC/abc・漢数字に対応し、区切りは . ． : ： | を認める。
       「12.5」のように区切りの直後が数字の場合は本文とみなして残す */
    function removeNumberMarkers(objects) {
        removeLinePrefix(objects, [
            /^\t(?:[①-⑳❶-❿⓫-⓴]|[A-Za-z]+|[〇一二三四五六七八九十百千]+|[0-9０-９]+)[.．:：|]?\t/,
            /^[\t 　]*[①-⑳❶-❿⓫-⓴][\t 　]*/,
            /^[\t 　]*(?:[A-Za-z]+|[〇一二三四五六七八九十百千]+|[0-9０-９]+)[.．:：|][\t 　]+/,
            /^[\t 　]*[0-9０-９]+[.．](?![0-9０-９])[\t 　]*/
        ]);
    }

    /* テキストフレームの順序を反転する関数 */
    /* テキストフレーム内の行の順序を反転する */
    function reverseOrder(objects) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            lines.reverse();
            return lines.join("\r");
        });
    }

    /* 重複行を削除する関数 */
    function removeDuplicateLines(objects) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            var seen = {};
            var kept = [];
            for (var i = 0; i < lines.length; i++) {
                /* "toString" などの組み込みプロパティ名と衝突しないよう接頭辞を付ける */
                var seenKey = "#" + lines[i];
                if (!seen[seenKey]) {
                    seen[seenKey] = true;
                    kept.push(lines[i]);
                }
            }
            return kept.join("\r");
        });
    }

    /* テキストフレーム内の行を文字コード順で並べ替える */
    function sortByCharCode(objects) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            lines.sort();
            return lines.join("\r");
        });
    }

    /* テキストフレーム内の行を文字数順で並べ替える */
    function sortByLength(objects) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            lines.sort(function (a, b) { return a.length - b.length; });
            return lines.join("\r");
        });
    }

    /* 和欧間のスペースを削除する関数
     * 欧文同士（英単語間）のスペースは残し、それ以外のスペースを削除する */

    function removeCjkLatinSpaces(objects) {
        transformContents(objects, function (txt) {
            var result = "";
            for (var i = 0; i < txt.length; i++) {
                var c = txt.charAt(i);
                if (c === " " || c === "\u3000") {
                    var prev = (i > 0) ? txt.charAt(i - 1) : "";
                    var next = (i < txt.length - 1) ? txt.charAt(i + 1) : "";
                    if (isLatinLetterOrDigit(prev) && isLatinLetterOrDigit(next)) {
                        result += c;
                    }
                    /* それ以外のスペースは削除（何も追加しない）  */
                } else {
                    result += c;
                }
            }
            return result;
        });
    }

    /* 1文字ごとに改行を挿入する関数 */
    function addLineBreakPerChar(objects) {
        transformContents(objects, function (txt) {
            var result = "";
            for (var i = 0; i < txt.length; i++) {
                var currentChar = txt.charAt(i);
                result += currentChar;
                if (isAnyBreak(currentChar) || i >= txt.length - 1) continue;
                if (!isAnyBreak(txt.charAt(i + 1))) result += "\r";
            }
            return result;
        });
    }

    /* 指定文字数ごとに改行を挿入する関数。
       折り返しだけを breakChar でつなぎ、元からあった段落改行はそのまま残す */
    function addLineBreakAtCount(objects, count, useForcedBreak) {
        var maxChars = parseInt(count, 10);
        if (!maxChars || maxChars <= 0) maxChars = parseInt(DEFAULT_BREAK_COUNT, 10);
        var breakChar = useForcedBreak ? String.fromCharCode(3) : "\r";
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            var wrappedLines = [];
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i];
                var segments = [];
                while (line.length > maxChars) {
                    segments.push(line.substring(0, maxChars));
                    line = line.substring(maxChars);
                }
                segments.push(line);
                wrappedLines.push(segments.join(breakChar));
            }
            return wrappedLines.join("\r");
        });
    }

    /* 強制改行を通常の改行に変換する関数 */
    function convertForcedLineBreaks(objects) {
        mutateMatchingChars(objects, isForcedBreak, String.fromCharCode(13));
    }

    /* 改行を強制改行に変換する関数 */
    function convertToForcedBreaks(objects) {
        mutateMatchingChars(objects, isParagraphBreak, String.fromCharCode(3));
    }

    /* 指定した記号の後に改行を挿入する関数（既定は和文・欧文の句読点と終止記号）*/
    function addLineBreakAtPunctuation(objects, punctuationChars) {
        var punctuation = punctuationChars || DEFAULT_BREAK_CHARS;
        transformContents(objects, function (txt) {
            /* 「後ろに実体のある文字が残っているか」は末尾位置を1回求めれば足りる */
            var lastVisibleIndex = findLastVisibleIndex(txt);
            var result = "";
            for (var i = 0; i < txt.length; i++) {
                var currentChar = txt.charAt(i);
                result += currentChar;
                if (punctuation.indexOf(currentChar) === -1 || i >= lastVisibleIndex) continue;
                if (!isAnyBreak(txt.charAt(i + 1))) result += "\r";
            }
            return result;
        });
    }

    /* 各テキストフレームの段落を上から順に走査し、段落ごとに
       makeFrames(sourceFrame, paragraph, originX, originY, fontSize) が返したフレームを集める。
       走査後に元フレームを削除し、生成したフレームを1つのグループにまとめて返す */
    function splitFramesByParagraph(objects, makeFrames) {
        var targetLayer = app.activeDocument.activeLayer;
        var sourceFrames = getTextFrames(objects);
        var resultFrames = [];

        for (var i = 0; i < sourceFrames.length; i++) {
            var sourceFrame = sourceFrames[i];
            var originX = sourceFrame.position[0];
            var originY = sourceFrame.position[1];
            var paragraphCount = sourceFrame.paragraphs.length;

            for (var j = 0; j < paragraphCount; j++) {
                var paragraph = sourceFrame.paragraphs[j];
                var metrics = getParagraphMetrics(paragraph, sourceFrame);
                var madeFrames = makeFrames(sourceFrame, paragraph, originX, originY, metrics.size, j);
                for (var k = 0; k < madeFrames.length; k++) resultFrames.push(madeFrames[k]);
                originY -= metrics.leading;
            }
            sourceFrame.remove();
        }

        return groupTextFrames(resultFrames, targetLayer);
    }

    /* 段落の文字サイズと行送りを返す。
       TextRange 自体に size / leading は無いので characterAttributes から読み、
       取得できないときはフレーム全体の書式、それも駄目なら既定値へ落とす */
    function getParagraphMetrics(paragraph, sourceFrame) {
        var fontSize = 0;
        var leading = 0;
        try {
            var attrs = paragraph.characterAttributes;
            fontSize = attrs.size;
            leading = attrs.leading;
        } catch (e) { debugLog("getParagraphMetrics: paragraph attributes", e); }

        if (!fontSize) {
            try { fontSize = sourceFrame.textRange.characterAttributes.size; } catch (e) { debugLog("getParagraphMetrics: frame attributes", e); }
        }
        if (!fontSize) fontSize = DEFAULT_FONT_SIZE_PT;
        if (!leading) leading = fontSize * AUTO_LEADING_RATIO;

        return { size: fontSize, leading: leading };
    }

    /* 分割対象になる文字（改行・タブ・スペース以外）の数を数える */
    function countSplittableChars(txt) {
        var text = String(txt || "");
        var count = 0;
        for (var i = 0; i < text.length; i++) {
            if (!isNonSplittableChar(text.charAt(i))) count++;
        }
        return count;
    }

    /* タブ直後の文字が、その段落の先頭文字からどれだけ右にあるかを段落ごとに返す。
       TextRange には座標が無いため、複製をアウトライン化してグリフの実測値から求める。
       戻り値は段落インデックスをキーにした配列（求められない段落は空配列）*/
    function collectTabOffsetsByParagraph(sourceFrame) {
        var offsetsByParagraph = [[]];

        var outlineInfo = null;
        try { outlineInfo = buildOutlineCharBounds(sourceFrame); } catch (e) { debugLog("collectTabOffsetsByParagraph: outline bounds", e); }
        if (outlineInfo && outlineInfo.outlinedRoot) removeItems([outlineInfo.outlinedRoot]);
        if (!outlineInfo || !outlineInfo.ok) return offsetsByParagraph;

        var boundsList = outlineInfo.boundsList;
        var txt = String(sourceFrame.contents || "");

        /* グリフ数と文字数がずれると対応が取れないので、その場合は呼び出し側の概算に任せる */
        if (countSplittableChars(txt) !== boundsList.length) {
            debugLog("collectTabOffsetsByParagraph: glyph count mismatch");
            return offsetsByParagraph;
        }

        var paragraphIndex = 0;
        var paragraphOriginX = null;
        var glyphIndex = 0;
        var afterTab = false;

        for (var i = 0; i < txt.length; i++) {
            var currentChar = txt.charAt(i);
            if (isParagraphBreak(currentChar)) {
                paragraphIndex++;
                offsetsByParagraph[paragraphIndex] = [];
                paragraphOriginX = null;
                afterTab = false;
                continue;
            }
            if (isTabChar(currentChar)) {
                afterTab = true;
                continue;
            }
            if (isNonSplittableChar(currentChar)) continue;

            var glyphLeft = boundsList[glyphIndex][0];
            if (paragraphOriginX === null) paragraphOriginX = glyphLeft;
            if (afterTab) {
                offsetsByParagraph[paragraphIndex].push(glyphLeft - paragraphOriginX);
                afterTab = false;
            }
            glyphIndex++;
        }
        return offsetsByParagraph;
    }

    /* タブで分解する関数（元のタブ位置に合わせて横へ並べ直す）*/
    function splitByTab(objects) {
        var targetLayer = app.activeDocument.activeLayer;
        var measuredFrame = null;
        var tabOffsets = [];

        return splitFramesByParagraph(objects, function (sourceFrame, paragraph, originX, originY, fontSize, paragraphIndex) {
            /* タブ位置の実測はフレーム単位で1回だけ行う */
            if (measuredFrame !== sourceFrame) {
                measuredFrame = sourceFrame;
                tabOffsets = collectTabOffsetsByParagraph(sourceFrame);
            }
            var offsetsX = tabOffsets[paragraphIndex] || [];
            var segments = stripTrailingBreaks(paragraph.contents).split("\t");
            var madeFrames = [];
            var prevFrame = null;

            for (var i = 0; i < segments.length; i++) {
                var segmentFrame = sourceFrame.duplicate(targetLayer);
                segmentFrame.contents = segments[i];

                var segmentX = originX;
                if (i > 0) {
                    if ((i - 1) < offsetsX.length) {
                        segmentX = originX + offsetsX[i - 1];
                    } else if (prevFrame) {
                        /* タブ位置が取れない分は直前フレームの右端から半角分あける */
                        segmentX = prevFrame.visibleBounds[2] + (fontSize * 0.5);
                    } else {
                        segmentX = originX + (fontSize * i);
                    }
                }
                segmentFrame.position = [segmentX, originY];
                prevFrame = segmentFrame;
                madeFrames.push(segmentFrame);
            }
            return madeFrames;
        });
    }

    /* 改行で分割する関数（段落ごとに別フレームへ）*/
    function splitByLineBreak(objects) {
        var targetLayer = app.activeDocument.activeLayer;
        return splitFramesByParagraph(objects, function (sourceFrame, paragraph, originX, originY) {
            var paragraphText = stripTrailingBreaks(paragraph.contents);
            if (paragraphText === "") return [];

            var lineFrame = sourceFrame.duplicate(targetLayer);
            lineFrame.contents = paragraphText;
            lineFrame.position = [originX, originY];
            return [lineFrame];
        });
    }

    /* フレーム末尾の改行文字（\r \n 強制改行）を除去し、除去数を返す。
       除去後に中身が空になったらフレームごと削除し Infinity を返す（呼び出し側で index を打ち切る）*/
    function trimTrailingBreaks(frame) {
        var removed = 0;
        for (var ci = frame.characters.length - 1; ci >= 0; ci--) {
            if (isAnyBreak(frame.characters[ci].contents)) {
                frame.characters[ci].remove();
                removed++;
            } else {
                break;
            }
        }
        if (/^\s*$/.test(frame.contents)) {
            frame.remove();
            return Infinity;
        }
        return removed;
    }

    /* 組み方向から、分割時に基準とする軸と geometricBounds の末尾側 index を返す。
       横組み → Y方向に伸びるので tail = bottom(3)、縦組み → X方向(左)に伸びるので tail = left(0) */
    function getTailAxis(frame) {
        var isHorizontal = (frame.orientation === TextOrientation.HORIZONTAL);
        return {
            axisIndex: isHorizontal ? 1 : 0,
            boundsIndex: isHorizontal ? 3 : 0
        };
    }

    /* 1フレームを段落ごとに分割し、書式と位置を保ったフレーム配列を返す。
       TextRange.duplicate で書式を維持し、geometricBounds の tail 座標で位置を合わせる。
       参考: Split Rows for Ai.jsx の splitRowsPoint 方式 */
    function splitFrameKeepStyle(sourceFrame) {
        var tail = getTailAxis(sourceFrame);

        /* 末尾改行を事前に除去。空になったフレームは trimTrailingBreaks 側で削除済み */
        if (trimTrailingBreaks(sourceFrame) === Infinity) return [];

        var paragraphs = sourceFrame.paragraphs;
        if (paragraphs.length <= 1) return [sourceFrame];

        var splitFrames = [];

        /* 後ろの段落から順に処理（参考スクリプトと同じ方式）*/
        for (var i = paragraphs.length - 1; i >= 1; i--) {
            var currentParagraph = paragraphs[i];

            /* 空行またはホワイトスペースのみの行は消して次へ */
            if (/^\s*$/.test(currentParagraph.contents)) {
                currentParagraph.remove();
                continue;
            }

            /* 分割前の tail 座標＝この段落があるべき位置の基準 */
            var oldTail = sourceFrame.geometricBounds[tail.boundsIndex];

            var newFrame = sourceFrame.duplicate(sourceFrame, ElementPlacement.PLACEAFTER);
            newFrame.contents = "";
            currentParagraph.duplicate(newFrame, ElementPlacement.INSIDE);
            trimTrailingBreaks(newFrame);

            /* 新フレームの tail を元フレームの旧 tail に揃えて元の位置へ戻す */
            var delta = [0, 0];
            delta[tail.axisIndex] = oldTail - newFrame.geometricBounds[tail.boundsIndex];
            newFrame.translate(delta[0], delta[1]);
            splitFrames.unshift(newFrame);

            currentParagraph.remove();

            /* 元フレームの末尾改行を除去し、除去した行数分 index を飛ばす。
               空になって削除された場合は元フレームを触れないのでここで打ち切る */
            var trimmedCount = trimTrailingBreaks(sourceFrame);
            if (trimmedCount === Infinity) return splitFrames;
            i -= trimmedCount;
        }

        /* 先頭段落が残った元フレームを先頭へ戻す */
        if (splitFrames.length > 0 && !/^\s*$/.test(sourceFrame.contents)) {
            sourceFrame.move(splitFrames[0], ElementPlacement.PLACEBEFORE);
            splitFrames.unshift(sourceFrame);
        }
        return splitFrames;
    }

    /* 改行で分割（書式保持）*/
    function splitByLineBreakKeepStyle(objects) {
        var sourceFrames = getTextFrames(objects);
        var resultFrames = [];

        for (var i = 0; i < sourceFrames.length; i++) {
            var sourceFrame = sourceFrames[i];

            /* 空フレームはスキップ */
            if (/^\s*$/.test(sourceFrame.contents)) {
                sourceFrame.remove();
                continue;
            }

            var splitFrames = splitFrameKeepStyle(sourceFrame);
            for (var j = 0; j < splitFrames.length; j++) resultFrames.push(splitFrames[j]);
        }

        return groupTextFrames(resultFrames, app.activeDocument.activeLayer);
    }

    /* =========================================
     * 1文字ごとにテキストフレームを分割（書式保持）
     * ========================================= */
    function splitByCharKeepStyle(objects) {
        return splitByChar(objects, true);
    }

    /* =========================================
     * 1文字ごとにテキストフレームを分割（書式無視）
     * =========================================  */
    function splitByCharIgnoreStyle(objects) {
        return splitByChar(objects, false);
    }

    /* 1文字分割の共通処理。keepStyle=false のときは先頭フォント以外の書式をリセットしてから分割 */
    function splitByChar(objects, keepStyle) {
        var frames = getTextFrames(objects);
        var resultFrames = [];
        for (var i = 0; i < frames.length; i++) {
            if (!keepStyle) stripStyleKeepFirstFont(frames[i]);
            var made = splitCharHighPrecision(frames[i], keepStyle);
            for (var j = 0; j < made.length; j++) resultFrames.push(made[j]);
        }
        return groupTextFrames(resultFrames, app.activeDocument.activeLayer);
    }

    /* 配列内のアイテムをまとめて削除（失敗は無視）*/
    function removeItems(items) {
        for (var i = 0; i < items.length; i++) {
            try { if (items[i]) items[i].remove(); } catch (e) { }
        }
    }

    /* 属性へ try 付きで代入（設定が失敗しても続行）*/
    function safeSet(target, propertyName, value) { try { target[propertyName] = value; } catch (e) { } }

    /* srcAttrs の1属性を dstAttrs へコピー（読み書きとも try で保護）*/
    function copyAttr(dstAttrs, srcAttrs, propertyName) { try { dstAttrs[propertyName] = srcAttrs[propertyName]; } catch (e) { } }

    /* characterAttributes を初期化（フォント・サイズは保持、色は黒・各種は既定へ）*/
    function applyResetStyle(charAttrs, keepFont, keepSize, blackColor) {
        if (keepFont) safeSet(charAttrs, "textFont", keepFont);
        if (keepSize != null) safeSet(charAttrs, "size", keepSize);
        safeSet(charAttrs, "fillColor", blackColor);
        safeSet(charAttrs, "baselineShift", 0);
        safeSet(charAttrs, "horizontalScale", 100);
        safeSet(charAttrs, "verticalScale", 100);
        safeSet(charAttrs, "rotation", 0);
        safeSet(charAttrs, "tracking", 0);
        safeSet(charAttrs, "kerningMethod", KerningMethod.METRICS);
        safeSet(charAttrs, "autoLeading", true);
    }

    /* 書式リセット（先頭文字のフォント情報のみ保持）*/
    function stripStyleKeepFirstFont(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;

        var textRange, chars, keepFont = null, keepSize = null;
        try {
            textRange = textFrame.textRange;
            chars = textRange.characters;
            if (!chars || chars.length < 1) return;
            keepFont = chars[0].characterAttributes.textFont;
            keepSize = chars[0].characterAttributes.size;
        } catch (e) { debugLog("stripStyleKeepFirstFont: read attributes", e); }
        if (!chars || chars.length < 1) return;

        var blackColor = new GrayColor();
        blackColor.gray = 100;

        try { applyResetStyle(textRange.characterAttributes, keepFont, keepSize, blackColor); } catch (e) { return; }
        for (var i = 0; i < chars.length; i++) {
            try { applyResetStyle(chars[i].characterAttributes, keepFont, keepSize, blackColor); } catch (e) { }
        }
    }

    /* 分割対象にならない文字（改行・タブ・スペース）か */
    function isNonSplittableChar(content) {
        return content === "" || content === " " || content === "　" ||
            isAnyBreak(content) || isTabChar(content);
    }

    /* 高精度分割（アウトライン化したグリフのバウンディングボックスに合わせて1文字ずつ配置）*/
    function splitCharHighPrecision(textFrame, keepStyle) {
        if (!textFrame || textFrame.typename !== "TextFrame") return [];

        var chars, charCount;
        try {
            chars = textFrame.textRange.characters;
            charCount = chars.length;
        } catch (e) { debugLog("splitCharHighPrecision: read characters", e); return []; }
        if (!charCount) return [];

        var outlineInfo = null;
        try { outlineInfo = buildOutlineCharBounds(textFrame); } catch (e) { debugLog("splitCharHighPrecision: outline bounds", e); }
        if (!outlineInfo || !outlineInfo.ok) return splitCharFallback(textFrame, keepStyle);

        var boundsList = outlineInfo.boundsList;

        /* 合字などでグリフ数と文字数がずれると途中から対応がずれるため、
           一致しない場合は作り始める前にフォールバックへ回す */
        if (countSplittableChars(textFrame.contents) !== boundsList.length) {
            debugLog("splitCharHighPrecision: glyph count mismatch");
            removeItems([outlineInfo.outlinedRoot]);
            return splitCharFallback(textFrame, keepStyle);
        }

        var targetLayer = textFrame.layer;
        var madeFrames = [];
        var boundsIndex = 0;

        /* 途中で失敗したら作りかけを片付けてフォールバックへ回す */
        function giveUpToFallback() {
            removeItems(madeFrames);
            removeItems([outlineInfo.outlinedRoot]);
            return splitCharFallback(textFrame, keepStyle);
        }

        for (var i = 0; i < charCount; i++) {
            var charItem, content;
            try { charItem = chars[i]; content = charItem.contents; } catch (e) { continue; }
            if (isNonSplittableChar(content)) continue;
            if (boundsIndex >= boundsList.length) return giveUpToFallback();

            var newFrame = null;
            try {
                newFrame = targetLayer.textFrames.add();
                newFrame.contents = content;
            } catch (e) { debugLog("splitCharHighPrecision: add frame", e); return giveUpToFallback(); }

            /* 書式と変形の引き継ぎは失敗しても続行する */
            try {
                if (keepStyle) copyCharacterAttributes(newFrame, charItem);
                else copyBaseFontAttributes(newFrame, textFrame);
                newFrame.matrix = textFrame.matrix;
                newFrame.left = textFrame.left;
                newFrame.top = textFrame.top;
            } catch (e) { debugLog("splitCharHighPrecision: copy style", e); }

            try {
                moveFrameToMatchBounds(newFrame, boundsList[boundsIndex]);
            } catch (e) {
                debugLog("splitCharHighPrecision: match bounds", e);
                removeItems([newFrame]);
                return giveUpToFallback();
            }

            madeFrames.push(newFrame);
            boundsIndex++;
        }

        removeItems([outlineInfo.outlinedRoot, textFrame]);
        return madeFrames;
    }

    /* テキストフレームをアウトライン化し、グリフごとの geometricBounds を読み順で返す */
    function buildOutlineCharBounds(textFrame) {
        var duplicatedFrame = null;
        try { duplicatedFrame = textFrame.duplicate(textFrame.parent, ElementPlacement.PLACEATBEGINNING); } catch (e) {
            try { duplicatedFrame = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATBEGINNING); } catch (err) { return { ok: false }; }
        }

        var outlinedGroup = null;
        try { outlinedGroup = duplicatedFrame.createOutline(); } catch (e) { debugLog("buildOutlineCharBounds: createOutline", e); }
        removeItems([duplicatedFrame]);
        if (!outlinedGroup) return { ok: false };

        /* アウトライン結果が単一グループなら、その中身をグリフとして扱う */
        var glyphItems = [];
        try {
            var topLevelItems = outlinedGroup.pageItems;
            var innerItems = (topLevelItems.length === 1 && topLevelItems[0].typename === "GroupItem")
                ? topLevelItems[0].pageItems : topLevelItems;
            for (var i = 0; i < innerItems.length; i++) glyphItems.push(innerItems[i]);
        } catch (e) { debugLog("buildOutlineCharBounds: collect glyphs", e); }

        sortOutlineItems(glyphItems);
        if (glyphItems.length === 0) {
            removeItems([outlinedGroup]);
            return { ok: false };
        }

        var boundsList = [];
        for (var j = 0; j < glyphItems.length; j++) {
            try { boundsList.push(glyphItems[j].geometricBounds); } catch (e) { }
        }
        return { ok: boundsList.length > 0, outlinedRoot: outlinedGroup, boundsList: boundsList };
    }

    /* アウトライン項目を読み順（上の行から、行内は左から）に並べ替える。items は破壊的に更新する */
    function sortOutlineItems(items) {
        if (!items || items.length <= 1) return;

        var glyphRecords = [];
        for (var i = 0; i < items.length; i++) {
            var bounds;
            try { bounds = items[i].geometricBounds; } catch (e) { continue; }
            if (!bounds || bounds.length !== 4) continue;
            glyphRecords.push({
                item: items[i],
                left: bounds[0],
                centerY: (bounds[1] + bounds[3]) / 2,
                height: Math.abs(bounds[1] - bounds[3]),
                order: i
            });
        }
        if (glyphRecords.length <= 1) return;

        var rowThreshold = estimateCharRowThreshold(glyphRecords);

        glyphRecords.sort(function (a, b) {
            var dy = b.centerY - a.centerY;
            if (Math.abs(dy) > 0.001) return (dy < 0) ? -1 : 1;
            var dx = a.left - b.left;
            if (Math.abs(dx) > 0.001) return (dx < 0) ? -1 : 1;
            return a.order - b.order;
        });

        /* Y中心が近いものを同じ行にまとめ、行の代表Yは平均で更新する */
        var rows = [];
        for (var j = 0; j < glyphRecords.length; j++) {
            var record = glyphRecords[j];
            var placed = false;
            for (var k = 0; k < rows.length; k++) {
                if (Math.abs(record.centerY - rows[k].centerY) > rowThreshold) continue;
                rows[k].items.push(record);
                rows[k].centerY = (rows[k].centerY * (rows[k].items.length - 1) + record.centerY) / rows[k].items.length;
                placed = true;
                break;
            }
            if (!placed) rows.push({ centerY: record.centerY, items: [record] });
        }

        rows.sort(function (a, b) { return b.centerY - a.centerY; });

        var sortedItems = [];
        for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            rows[rowIndex].items.sort(function (a, b) {
                var dx = a.left - b.left;
                if (Math.abs(dx) > 0.5) return (dx < 0) ? -1 : 1;
                return a.order - b.order;
            });
            for (var glyphIndex = 0; glyphIndex < rows[rowIndex].items.length; glyphIndex++) {
                sortedItems.push(rows[rowIndex].items[glyphIndex].item);
            }
        }

        /* バウンディング取得に失敗した要素は sortedItems に含まれないため、
           上書きではなく作り直して末尾に古い要素（重複）が残らないようにする */
        items.length = 0;
        for (var m = 0; m < sortedItems.length; m++) items.push(sortedItems[m]);
    }

    /* 行判定のしきい値をグリフ高さの中央値から推定する */
    function estimateCharRowThreshold(glyphRecords) {
        var heights = [];
        for (var i = 0; i < glyphRecords.length; i++) {
            if (glyphRecords[i].height > 0) heights.push(glyphRecords[i].height);
        }
        if (heights.length === 0) return 8;

        heights.sort(function (a, b) { return a - b; });
        var medianHeight = heights[Math.floor(heights.length / 2)] || heights[0];
        var rowThreshold = medianHeight * 0.6;
        return (rowThreshold < 2) ? 2 : rowThreshold;
    }

    /* 1文字分の文字属性を新しいフレームへコピーする */
    function copyCharacterAttributes(dstFrame, srcChar) {
        var srcAttrs, dstAttrs;
        try {
            srcAttrs = srcChar.characterAttributes;
            dstAttrs = dstFrame.textRange.characterAttributes;
        } catch (e) { debugLog("copyCharacterAttributes: read attributes", e); return; }

        var copiedProps = ["textFont", "size", "horizontalScale", "verticalScale",
            "tracking", "baselineShift", "rotation", "autoLeading", "kerningMethod"];
        for (var i = 0; i < copiedProps.length; i++) copyAttr(dstAttrs, srcAttrs, copiedProps[i]);

        /* 色と行送りは「無色」「自動」のときにコピーしない */
        try { if (srcAttrs.fillColor && srcAttrs.fillColor.typename !== "NoColor") dstAttrs.fillColor = srcAttrs.fillColor; } catch (e) { }
        try {
            if (srcAttrs.strokeColor && srcAttrs.strokeColor.typename !== "NoColor") {
                dstAttrs.strokeColor = srcAttrs.strokeColor;
                dstAttrs.strokeWeight = srcAttrs.strokeWeight;
            }
        } catch (e) { }
        try { if (!srcAttrs.autoLeading) dstAttrs.leading = srcAttrs.leading; } catch (e) { }
    }

    /* 書式を無視して分割する場合でも、フォントと文字サイズは元フレームから引き継ぐ
       （新規テキストフレームの既定書式に落ちると、元の見た目と大きさが変わってしまう）*/
    function copyBaseFontAttributes(dstFrame, srcFrame) {
        var srcAttrs, dstAttrs;
        try {
            srcAttrs = srcFrame.textRange.characterAttributes;
            dstAttrs = dstFrame.textRange.characterAttributes;
        } catch (e) { debugLog("copyBaseFontAttributes: read attributes", e); return; }

        copyAttr(dstAttrs, srcAttrs, "textFont");
        copyAttr(dstAttrs, srcAttrs, "size");
    }

    /* アウトライン化して求めた実寸バウンディングボックスの中心へフレームを移動する */
    function moveFrameToMatchBounds(newFrame, targetBounds) {
        if (!newFrame || !targetBounds || targetBounds.length !== 4) return;

        var duplicatedFrame = null;
        try { duplicatedFrame = newFrame.duplicate(newFrame.parent, ElementPlacement.PLACEATBEGINNING); } catch (e) {
            try { duplicatedFrame = newFrame.duplicate(newFrame.layer, ElementPlacement.PLACEATBEGINNING); } catch (err) { return; }
        }

        var outlinedGroup = null;
        try { outlinedGroup = duplicatedFrame.createOutline(); } catch (e) { debugLog("moveFrameToMatchBounds: createOutline", e); }
        removeItems([duplicatedFrame]);
        if (!outlinedGroup) return;

        var currentBounds = null;
        try { currentBounds = outlinedGroup.geometricBounds; } catch (e) { }
        removeItems([outlinedGroup]);
        if (!currentBounds || currentBounds.length !== 4) return;

        var dx = (targetBounds[0] + targetBounds[2]) / 2 - (currentBounds[0] + currentBounds[2]) / 2;
        var dy = (targetBounds[1] + targetBounds[3]) / 2 - (currentBounds[1] + currentBounds[3]) / 2;

        try { newFrame.left += dx; newFrame.top += dy; } catch (e) {
            try { newFrame.translate(dx, dy); } catch (err) { }
        }
    }

    /* 各文字の左端X方向オフセットを先頭から積算して配列で返す（フォールバック用）。
       文字ごとに先頭から測り直すと文字数の2乗に比例して遅くなるため、1回の走査で全件求める */
    function measureCharOffsetsX(textFrame, targetLayer) {
        var measuredProps = ["textFont", "size", "horizontalScale", "verticalScale", "tracking"];
        var chars = textFrame.textRange.characters;
        var offsets = [];
        var offsetX = 0;

        for (var i = 0; i < chars.length; i++) {
            offsets.push(offsetX);

            var currentChar = chars[i];
            if (isNonSplittableChar(currentChar.contents)) continue;

            var measureFrame = null;
            try {
                measureFrame = targetLayer.textFrames.add();
                measureFrame.contents = currentChar.contents;
                var measureAttrs = measureFrame.textRange.characterAttributes;
                var srcAttrs = currentChar.characterAttributes;
                for (var j = 0; j < measuredProps.length; j++) copyAttr(measureAttrs, srcAttrs, measuredProps[j]);
                offsetX += measureFrame.width;
            } catch (e) { debugLog("measureCharOffsetsX: measure char", e); }
            removeItems([measureFrame]);
        }
        return offsets;
    }

    /* フォールバック処理（アウトラインを使わず、文字幅の積算で位置を再構成）*/
    function splitCharFallback(textFrame, keepStyle) {
        if (!textFrame || textFrame.typename !== "TextFrame") return [];

        var charCount, targetLayer;
        try {
            charCount = textFrame.textRange.characters.length;
            targetLayer = textFrame.layer;
        } catch (e) { debugLog("splitCharFallback: read frame", e); return []; }
        if (!charCount) return [];

        /* 回転しているフレームでも並びが崩れないよう、回転角に沿ってオフセットを掛ける */
        var angleRad = 0;
        try {
            var matrix = textFrame.matrix;
            angleRad = Math.atan2(matrix.mValueB, matrix.mValueA);
        } catch (e) { debugLog("splitCharFallback: matrix", e); }

        var offsetsX = measureCharOffsetsX(textFrame, targetLayer);
        var madeFrames = [];
        for (var i = charCount - 1; i >= 0; i--) {
            try {
                var charItem = textFrame.textRange.characters[i];
                var content = charItem.contents;
                if (isNonSplittableChar(content)) continue;

                var newFrame = targetLayer.textFrames.add();
                newFrame.contents = content;
                if (keepStyle) copyCharacterAttributes(newFrame, charItem);
                else copyBaseFontAttributes(newFrame, textFrame);
                try { newFrame.matrix = textFrame.matrix; } catch (e) { debugLog("splitCharFallback: matrix", e); }

                var offsetX = offsetsX[i] || 0;
                newFrame.left = textFrame.left + offsetX * Math.cos(angleRad);
                newFrame.top = textFrame.top - offsetX * Math.sin(angleRad);

                madeFrames.push(newFrame);
            } catch (e) { debugLog("splitCharFallback: main loop", e); }
        }

        removeItems([textFrame]);
        return madeFrames;
    }

    /* =========================================
     * 連結（縦・横）
     *
     * いずれも見た目ベースの近似連結で、複雑な書式差・回転・厳密な段落属性までは保持しない。
     * - 縦連結：上から下、同じ高さでは左から右の順に、各行を \r でつないだ1つのテキストへ再構成
     * - 横連結（行維持）：同じ行のテキストを左から右へ連結し、行は別フレームのまま残す
     * - 横連結（行統合）：行ごとに連結したうえで、行間ルール（英文ハイフン除去／語間スペース／
     *   句点でのみ改行）に従って1つのテキストへまとめる
     * ========================================= */

    /* Y位置でフレームを行ごとにグループ化する（同じ行と見なすY座標差は LINE_Y_THRESHOLD）*/
    function groupFramesIntoRows(frames) {
        return groupByLineY(sortByY(frames), LINE_Y_THRESHOLD);
    }

    /* 行内のフレームをX順に連結した文字列と、その並びを返す */
    function concatRowText(rowFrames) {
        var sortedFrames = sortByX(rowFrames);
        var rowText = "";
        for (var i = 0; i < sortedFrames.length; i++) rowText += stripTrailingBreaks(sortedFrames[i].contents);
        return { sorted: sortedFrames, text: rowText };
    }

    /* 横連結（行維持）：同じ行を左から右へ連結し、行ごとに別テキストフレームとして残す */
    function concatHorizontalOnly(objects) {
        var textFrames = getTextFrames(objects);
        if (textFrames.length < 2) return textFrames;

        var rows = groupFramesIntoRows(textFrames);
        var resultFrames = [];

        for (var i = 0; i < rows.length; i++) {
            /* 1つしかない行はそのまま残す（内容に触れない）*/
            if (rows[i].length === 1) {
                resultFrames.push(rows[i][0]);
                continue;
            }

            var row = concatRowText(rows[i]);
            row.sorted[0].contents = row.text;
            resultFrames.push(row.sorted[0]);
            removeItems(row.sorted.slice(1));
        }

        return groupTextFrames(resultFrames, app.activeDocument.activeLayer);
    }

    /* 縦連結：各フレームを行単位へ分解し、位置順に並べ直して1つのテキストへ再構成する */
    function concatVertical(objects) {
        var sourceFrames = getTextFrames(objects);
        if (sourceFrames.length < 2) return sourceFrames;

        sourceFrames = sortByPosition(sourceFrames);

        /* 各テキストフレームを行単位に分解し、元の行送りで仮のY位置を与える */
        var lineFrames = [];
        for (var i = 0; i < sourceFrames.length; i++) {
            var sourceFrame = sourceFrames[i];
            var lines = splitParagraphLines(sourceFrame.contents);
            var baseX = sourceFrame.position[0];
            var baseY = sourceFrame.position[1];
            var baseLeading = getParagraphMetrics(sourceFrame.textRange, sourceFrame).leading;

            for (var j = 0; j < lines.length; j++) {
                var lineText = stripTrailingBreaks(lines[j]);
                if (lineText === "") continue;

                var lineFrame = sourceFrame.duplicate();
                lineFrame.contents = lineText;
                lineFrame.position = [baseX, baseY - (j * baseLeading)];
                lineFrames.push(lineFrame);
            }
        }
        if (lineFrames.length === 0) return [];

        /* 仮配置後に位置で並べ直して連結順を決める */
        lineFrames = sortByPosition(lineFrames);

        var mergedLines = [];
        for (var k = 0; k < lineFrames.length; k++) {
            var mergedLineText = stripTrailingBreaks(lineFrames[k].contents);
            if (mergedLineText !== "") mergedLines.push(mergedLineText);
        }
        if (mergedLines.length === 0) {
            removeItems(lineFrames);
            return [];
        }

        /* 最上段のフレームをベースに、残りは削除 */
        var baseFrame = lineFrames[0];
        baseFrame.contents = mergedLines.join("\r");
        removeItems(lineFrames.slice(1));

        var leftovers = [];
        for (var m = 0; m < sourceFrames.length; m++) {
            if (sourceFrames[m] !== baseFrame) leftovers.push(sourceFrames[m]);
        }
        removeItems(leftovers);

        return [baseFrame];
    }

    /* 塗り・線なしの矩形を作る（エリア内文字の枠用）*/
    function makeFramelessRect(bounds) {
        var rect = app.activeDocument.pathItems.rectangle(
            bounds[1], bounds[0], bounds[2] - bounds[0], bounds[1] - bounds[3]
        );
        rect.stroked = false;
        rect.filled = false;
        return rect;
    }

    /* Justification.LEFT は直接代入すると無視されるため、一時 resize（200%→50%＝実質等倍）で
       段落属性をリフレッシュしてから左揃えにする */
    function applyLeftJustification(frame) {
        if (!frame) return;
        try {
            frame.resize(200, 200);
            frame.resize(50, 50);
            frame.textRange.justification = Justification.LEFT;
        } catch (e) { debugLog("applyLeftJustification", e); }
    }

    /* 1行のみの横連結（area＝エリア内文字を新規作成 / それ以外＝先頭フレームへ集約）*/
    function concatHorizontalSingleLine(rowFrames, textMode) {
        var row = concatRowText(rowFrames);
        var sortedFrames = row.sorted;
        var baseAttrs = sortedFrames[0].textRange.characterAttributes;
        var resultFrame;

        if (textMode === "area") {
            resultFrame = createConcatOutputText("area", row.text, getUnionBounds(sortedFrames),
                baseAttrs.textFont, baseAttrs.size, true);
        } else {
            sortedFrames[0].contents = row.text;
            resultFrame = sortedFrames[0];
            applyLeftJustification(resultFrame);
        }

        var leftovers = [];
        for (var i = 0; i < sortedFrames.length; i++) {
            if (sortedFrames[i] !== resultFrame) leftovers.push(sortedFrames[i]);
        }
        removeItems(leftovers);

        return [resultFrame];
    }

    /* 連結済みの各行を、英文ハイフン除去／語間スペース／句点改行のルールで1つの文字列へ */
    function buildJoinedParagraphText(mergedFrames) {
        var joinedText = "";
        for (var i = 0; i < mergedFrames.length; i++) {
            var content = stripTrailingBreaks(mergedFrames[i].contents);
            joinedText += content;
            if (i >= mergedFrames.length - 1) continue;

            var nextContent = stripTrailingBreaks(mergedFrames[i + 1].contents);
            if (/[A-Za-z0-9)]-$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                /* 英単語がハイフンで分断 → ハイフン除去して結合 */
                joinedText = joinedText.replace(/-$/, "");
            } else if (/[A-Za-z0-9)]$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                /* 行末・次行頭がともに英単語 → 語間スペースを補う */
                joinedText += " ";
            }
            /* 句点等で終わる場合のみ改行を残す */
            if (shouldInsertParagraphBreakBetweenLines(content)) joinedText += "\r";
        }
        return joinedText;
    }

    /* 連結結果テキストを point/area で出力する（font/size/kinsoku/justification を設定）。
       forceLeft が真なら、エリア内文字でも均等配置ではなく左揃えにする */
    function createConcatOutputText(textMode, joinedText, bounds, baseFont, baseFontSize, forceLeft) {
        var isPointText = (textMode === "point");
        var outputFrame = isPointText
            ? app.activeDocument.textFrames.pointText([bounds[0], bounds[1]])
            : app.activeDocument.textFrames.areaText(makeFramelessRect(bounds));

        outputFrame.contents = joinedText;
        outputFrame.textRange.characterAttributes.textFont = baseFont;
        outputFrame.textRange.characterAttributes.size = baseFontSize;
        if (outputFrame.paragraphs.length > 0) outputFrame.paragraphs[0].paragraphAttributes.kinsoku = "Soft";

        if (isPointText || forceLeft) {
            applyLeftJustification(outputFrame);
        } else {
            outputFrame.textRange.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        }
        return outputFrame;
    }

    /* 横連結（行統合）：行ごとに連結したうえで、1つのテキストへまとめる。
       textMode が未指定または "mixed" のときは選択から自動判定し、混在時はエリア内文字を優先する */
    function concatHorizontal(objects, textMode) {
        var textFrames = getTextFrames(objects);
        if (textFrames.length < 2) return textFrames;

        if (!textMode || textMode === "mixed") {
            textMode = detectTextFrameType(textFrames);
            if (textMode === "mixed") textMode = "area";
        }

        /* 強制改行を削除してから行グループ化する */
        removeForcedLineBreaks(textFrames);
        var rows = groupFramesIntoRows(textFrames);
        if (rows.length === 1) return concatHorizontalSingleLine(rows[0], textMode);

        /* 行ごとにX順で連結して中間フレームを作る */
        var mergedFrames = [];
        for (var i = 0; i < rows.length; i++) {
            var row = concatRowText(rows[i]);
            var mergedFrame = row.sorted[0].duplicate();
            mergedFrame.contents = row.text;
            mergedFrame.position = row.sorted[0].position;
            mergedFrames.push(mergedFrame);
            removeItems(row.sorted);
        }

        var baseAttrs = mergedFrames[0].textRange.characterAttributes;
        var baseFont = baseAttrs.textFont;
        var baseFontSize = baseAttrs.size;
        var outputFrame = createConcatOutputText(textMode, buildJoinedParagraphText(mergedFrames),
            getUnionBounds(mergedFrames), baseFont, baseFontSize, false);

        /* 行送りは連結前の行間から復元する */
        if (mergedFrames.length >= 2) {
            var leading = Math.abs(mergedFrames[0].position[1] - mergedFrames[1].position[1]);
            if (leading < baseFontSize) leading = baseFontSize * MIN_LEADING_RATIO;
            outputFrame.textRange.characterAttributes.autoLeading = false;
            outputFrame.textRange.characterAttributes.leading = leading;
        }

        removeItems(mergedFrames);
        app.redraw();
        return [outputFrame];
    }

    /* =========================================
     * メインエンジン委譲（BridgeTalk）
     *
     * パレットは常駐エンジン（#targetengine）で動くが、その app は
     * パレット表示中に DOM 接続を失い "there is no document" を投げる。
     * そこで DOM を触る全処理は、生きた DOM を持つメインエンジン
     * （bridge.target = "illustrator"）へ BridgeTalk で都度委譲する。
     *
     * 上で定義済みの処理関数群を toString() で連結して本文に同梱し、
     * 末尾の __dispatch をメインエンジンで実行する。結果は
     * "OK:<state>" / "LINES:<encoded>" / "ERR:<msg>" の文字列で返す。
     * ========================================= */

    /* メインエンジンへ送る処理関数（上で定義済みのものを再利用）*/
    var __LIB_FUNCS = [
        debugLog, normalizeParagraphBreaks, splitParagraphLines, trimLineSpaces, isBlankLine,
        stripTrailingBreaks, isLatinLetterOrDigit, isAsciiTextOnly, isSentenceEndingJP, isSentenceEndingEN,
        shouldInsertParagraphBreakBetweenLines, getCharCodeSafe, isParagraphBreak, isForcedBreak, isAnyBreak,
        isTabChar, findLastVisibleIndex, isNonSplittableChar, countSplittableChars,
        getTextFrames, countTextFrameTypes, detectTextFrameType, countBreakTypes, transformContents,
        hasMultipleLines, hasSpacesOrTabs, computeSelectionState, encodeSelectionState,
        mutateMatchingChars, removeForcedLineBreaks, removeItems, sortedCopy, sortByPosition, sortByY, sortByX,
        groupByLineY, getUnionBounds, groupTextFrames,
        removeLineBreaks, removeAllBreaks, joinBreaksToOneLine, flattenToOneLine, removeEmptyLines, removeTabs, tabsToSpaces,
        trimSpaces, collapseSpaces, removeLinePrefix, toHalfWidthAlnumText, toFullWidthKanaText,
        fullToHalfAlnum, halfToFullKana, removeBulletMarkers, removeNumberMarkers,
        reverseOrder, removeDuplicateLines, sortByCharCode, sortByLength, removeCjkLatinSpaces,
        addLineBreakPerChar, addLineBreakAtCount, convertForcedLineBreaks, convertToForcedBreaks,
        addLineBreakAtPunctuation, splitFramesByParagraph, getParagraphMetrics,
        collectTabOffsetsByParagraph, splitByTab, splitByLineBreak,
        trimTrailingBreaks, getTailAxis, splitFrameKeepStyle, splitByLineBreakKeepStyle,
        splitByCharKeepStyle, splitByCharIgnoreStyle, splitByChar, stripStyleKeepFirstFont, splitCharHighPrecision,
        buildOutlineCharBounds, sortOutlineItems, estimateCharRowThreshold, copyCharacterAttributes,
        copyBaseFontAttributes, moveFrameToMatchBounds, measureCharOffsetsX, splitCharFallback,
        groupFramesIntoRows, concatRowText, concatHorizontalOnly, concatVertical, concatHorizontal,
        makeFramelessRect, applyLeftJustification, concatHorizontalSingleLine, buildJoinedParagraphText,
        createConcatOutputText,
        toWordCap, toSentenceCase, toTitleCase,
        ungroupResult, runStructureAction, runContentAction, runAction,
        runQueryAction, finalizeSelection,
        safeSet, copyAttr, applyResetStyle
    ];

    /* 処理関数がそのまま参照する定数（パレット側の値をメインエンジンへも渡す）*/
    var __WORKER_CONSTANTS = [
        ["DEFAULT_BREAK_CHARS", DEFAULT_BREAK_CHARS],
        ["DEFAULT_BREAK_COUNT", DEFAULT_BREAK_COUNT],
        ["LINE_Y_THRESHOLD", LINE_Y_THRESHOLD],
        ["AUTO_LEADING_RATIO", AUTO_LEADING_RATIO],
        ["MIN_LEADING_RATIO", MIN_LEADING_RATIO],
        ["DEFAULT_FONT_SIZE_PT", DEFAULT_FONT_SIZE_PT]
    ];

    /**
     * 定数を var 宣言の並びへ変換する（文字列はエスケープを避けて decodeURIComponent で復元）
     * @returns {string} "var NAME=...;" を並べたソース文字列
     */
    function buildConstSource() {
        var parts = [];
        for (var i = 0; i < __WORKER_CONSTANTS.length; i++) {
            var constName = __WORKER_CONSTANTS[i][0];
            var constValue = __WORKER_CONSTANTS[i][1];
            var literal = (typeof constValue === "number")
                ? String(constValue)
                : 'decodeURIComponent("' + encodeURIComponent(String(constValue)) + '")';
            parts.push("var " + constName + "=" + literal + ";");
        }
        return parts.join("\n");
    }

    /* Function.toString() の出力から関数本体だけを切り出す。
       ExtendScript は改行を CR で返し、前後のコメント断片（閉じていない \*\/ を含む）を
       混入させることがあるため、LF へ正規化したうえで先頭の function 行から
       閉じ括弧だけの行までを取り出す（そのままつなぐと構文エラーになる）*/
    function sliceFunctionSource(func) {
        var lines = String(func).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
        var firstIndex = -1;
        var lastIndex = -1;
        for (var i = 0; i < lines.length; i++) {
            if (firstIndex < 0 && /^\s*function\s/.test(lines[i])) firstIndex = i;
            if (/^\s*\}\s*$/.test(lines[i])) lastIndex = i;
        }
        if (firstIndex < 0) return String(func);
        if (lastIndex < firstIndex) {
            /* 1行で書かれた関数は、その行だけを取り出す */
            return /\}\s*$/.test(lines[firstIndex]) ? lines[firstIndex] : lines.slice(firstIndex).join("\n");
        }
        return lines.slice(firstIndex, lastIndex + 1).join("\n");
    }

    /* 定数と関数配列を1つのソース文字列にまとめる */
    function buildLibSource(funcs) {
        var sourceParts = [buildConstSource()];
        for (var i = 0; i < funcs.length; i++) sourceParts.push(sliceFunctionSource(funcs[i]));
        return sourceParts.join("\n");
    }

    var __LIB_SRC_CACHE = null;

    /* 全処理関数のソース（初回のみ組み立ててキャッシュ）*/
    function getLibSource() {
        if (__LIB_SRC_CACHE === null) __LIB_SRC_CACHE = buildLibSource(__LIB_FUNCS);
        return __LIB_SRC_CACHE;
    }

    /*
     * メインエンジンで実行されるディスパッチャ。
     * ここで参照する処理関数（getTextFrames 等）は getLibSource() で
     * 先に定義される。この関数自体はパレット側では呼ばず、toString() で
     * 本文に埋め込む用途のみ。
     */
    /* 分割結果のグループを解除し、中身を親へ出してバラの状態にする */
    function ungroupResult(result) {
        if (!result || !result.length) return result;
        var frames = getTextFrames(result);
        for (var i = 0; i < result.length; i++) {
            var groupItem = result[i];
            try {
                if (!groupItem || groupItem.typename !== "GroupItem") continue;
                var parentItem = groupItem.parent;
                for (var j = groupItem.pageItems.length - 1; j >= 0; j--) {
                    groupItem.pageItems[j].move(parentItem, ElementPlacement.PLACEATEND);
                }
                groupItem.remove();
            } catch (e) { debugLog("ungroupResult", e); }
        }
        return frames;
    }
    /* フレーム構成を変える処理（分割・連結・1行化）。該当しなければ null を返す */
    function runStructureAction(actionName, targets, params) {
        switch (actionName) {
            case "flatten": {
                var flattened = flattenToOneLine(targets);
                var flattenTargets = (flattened && flattened.length) ? flattened : targets;
                trimSpaces(flattenTargets);
                removeCjkLatinSpaces(flattenTargets);
                collapseSpaces(flattenTargets);
                return flattenTargets;
            }
            case "splitByLineBreak": return ungroupResult(splitByLineBreak(targets));
            case "splitByLineBreakKeepStyle": return splitByLineBreakKeepStyle(targets);
            case "splitByTab": {
                var tabFrames = splitByTab(targets);
                return params.ungroup ? ungroupResult(tabFrames) : tabFrames;
            }
            case "splitByCharKeepStyle": return splitByCharKeepStyle(targets);
            case "splitByCharIgnoreStyle": return splitByCharIgnoreStyle(targets);
            case "concatVertical": return concatVertical(targets);
            case "concatHorizontalOnly": return concatHorizontalOnly(targets);
            case "concatH": return concatHorizontal(targets, detectTextFrameType(targets));
            case "concatToArea": return concatHorizontal(targets, "area");
        }
        return null;
    }

    /* テキスト内容だけを書き換える処理（対象フレームはそのまま）*/
    function runContentAction(actionName, targets, params) {
        switch (actionName) {
            case "removeLineBreaks": if (params.forced) removeAllBreaks(targets); else removeLineBreaks(targets); return;
            case "addLineBreakPerChar": addLineBreakPerChar(targets); return;
            case "punctuation": addLineBreakAtPunctuation(targets, params.chars); return;
            case "breakAtCount": addLineBreakAtCount(targets, params.count, params.forced); return;
            case "convertForcedLineBreaks": convertForcedLineBreaks(targets); return;
            case "convertToForcedBreaks": convertToForcedBreaks(targets); return;
            case "removeTabs": removeTabs(targets); return;
            case "tabsToSpaces": tabsToSpaces(targets); return;
            case "trimSpaces": trimSpaces(targets); return;
            case "removeCjkLatinSpaces": removeCjkLatinSpaces(targets); return;
            case "collapseSpaces": collapseSpaces(targets); return;
            case "cleanupSpaces": trimSpaces(targets); removeCjkLatinSpaces(targets); collapseSpaces(targets); return;
            case "removeAllSpaces": transformContents(targets, function (txt) { return txt.replace(/[ 　]/g, ""); }); return;
            case "fullToHalfAlnum": fullToHalfAlnum(targets); return;
            case "halfToFullKana": halfToFullKana(targets); return;
            case "removeBulletMarkers": removeBulletMarkers(targets); return;
            case "removeNumberMarkers": removeNumberMarkers(targets); return;
            case "reverseOrder": reverseOrder(targets); return;
            case "removeDuplicateLines": removeDuplicateLines(targets); return;
            case "sortByCharCode": sortByCharCode(targets); return;
            case "sortByLength": sortByLength(targets); return;
            case "removeEmptyLines": removeEmptyLines(targets); return;
            case "caseUpper": transformContents(targets, function (txt) { return txt.toUpperCase(); }); return;
            case "caseLower": transformContents(targets, function (txt) { return txt.toLowerCase(); }); return;
            case "caseWord": transformContents(targets, toWordCap); return;
            case "caseSentence": transformContents(targets, toSentenceCase); return;
            case "caseTitle": transformContents(targets, toTitleCase); return;
            case "spaceAfterPunct": transformContents(targets, function (txt) { return txt.replace(/([.,])(?=[^\s\d.,])/g, "$1 "); }); return;
            case "convertSymbol": {
                /* ES の入れ子三項は左結合に誤評価されるため括弧で右結合を明示 */
                var fromPattern = (params.from === "underscore") ? /_/g : ((params.from === "hyphen") ? /-/g : /[ 　]/g);
                var toText = (params.to === "space") ? " " : ((params.to === "hyphen") ? "-" : "_");
                transformContents(targets, function (txt) { return txt.replace(fromPattern, toText); });
                return;
            }
        }
    }

    /* アクションIDに対応する処理を実行し、処理後の対象フレーム配列を返す */
    function runAction(actionName, targets, params) {
        var restructured = runStructureAction(actionName, targets, params);
        if (restructured) return restructured;
        runContentAction(actionName, targets, params);
        return targets;
    }

    /* 選択の取得・設定のみで完結するアクション。該当しなければ null を返す */
    function runQueryAction(actionId, params, doc, targets) {
        switch (actionId) {
            case "getState":
                return "OK:" + encodeSelectionState(targets);
            case "getLines":
                return (targets.length === 1) ? "LINES:" + encodeURIComponent(targets[0].contents) : "LINES:";
            case "getFirstText":
                return (targets.length >= 1) ? "TEXT:" + encodeURIComponent(targets[0].contents) : "TEXT:";
            case "setLines":
                if (targets.length === 1) targets[0].contents = params.text;
                app.redraw();
                return "OK:" + encodeSelectionState(getTextFrames(doc.selection));
            case "hiddenChar":
                try { app.executeMenuCommand("showHiddenChar"); } catch (e) { debugLog("showHiddenChar", e); }
                app.redraw();
                return "OK:" + encodeSelectionState(targets);
            case "finalizeClose":
                finalizeSelection(targets, params.turnOffHidden);
                return "OK:";
        }
        return null;
    }

    /* パレットを閉じるときの後始末：1要素だけのグループを解除し、選択を整えて制御文字表示を戻す */
    function finalizeSelection(targets, turnOffHidden) {
        for (var i = 0; i < targets.length; i++) {
            try {
                var parentGroup = targets[i].parent;
                if (!parentGroup || parentGroup.typename !== "GroupItem" || parentGroup.pageItems.length !== 1) continue;
                targets[i].move(parentGroup.parent, ElementPlacement.PLACEATEND);
                parentGroup.remove();
            } catch (e) { debugLog("finalizeSelection: ungroup", e); }
        }
        try { if (targets.length > 0) app.activeDocument.selection = targets; } catch (e) { debugLog("finalizeSelection: select", e); }
        if (turnOffHidden) {
            try { app.executeMenuCommand("showHiddenChar"); } catch (e) { debugLog("finalizeSelection: hidden char", e); }
        }
        app.redraw();
    }

    function __dispatch(actionId, params) {
        if (app.documents.length === 0) return "ERR:nodoc";
        var doc;
        try { doc = app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        var targets = getTextFrames(doc.selection);

        var queryResult = runQueryAction(actionId, params, doc, targets);
        if (queryResult !== null) return queryResult;

        /* 通常の処理アクション */
        if (!targets.length) return "OK:" + encodeSelectionState([]);

        var result;
        try {
            result = runAction(actionId, targets, params);
        } catch (err) {
            app.redraw();
            return "ERR:" + (err && err.message ? err.message : String(err));
        }

        var refreshed = getTextFrames((result && result.length) ? result : targets);
        if (refreshed.length > 0) {
            try { doc.selection = refreshed; } catch (e) { debugLog("__dispatch: select result", e); }
        }
        app.redraw();
        return "OK:" + encodeSelectionState(refreshed);
    }

    var __DISPATCH_SRC = "(" + sliceFunctionSource(__dispatch) + ")";

    /**
     * アクションのパラメーターを、そのまま eval できる JS リテラル文字列へ変換する
     * @param {Object} params - forced / turnOffHidden / ungroup / count / chars / text / from / to を持つオブジェクト
     * @returns {string} "{ ... }" 形式のソース文字列
     */
    function paramsToSource(params) {
        if (!params) return "{}";

        var parts = [];
        if (params.forced !== undefined) parts.push("forced:" + (params.forced ? "true" : "false"));
        if (params.turnOffHidden !== undefined) parts.push("turnOffHidden:" + (params.turnOffHidden ? "true" : "false"));
        if (params.ungroup !== undefined) parts.push("ungroup:" + (params.ungroup ? "true" : "false"));
        if (params.count !== undefined) parts.push("count:" + parseInt(params.count, 10));
        if (params.chars !== undefined) parts.push('chars:decodeURIComponent("' + encodeURIComponent(params.chars) + '")');
        if (params.text !== undefined) parts.push('text:decodeURIComponent("' + encodeURIComponent(params.text) + '")');
        if (params.from !== undefined) parts.push('from:"' + params.from + '"');
        if (params.to !== undefined) parts.push('to:"' + params.to + '"');
        return "{" + parts.join(",") + "}";
    }

    /**
     * 本文をメインエンジンへ送り、結果マーカーを解析して onDone を呼ぶ。
     * BridgeTalk は本文送信時にバックスラッシュをエスケープする（"\r" がターゲットで "\\r" になる）ため、
     * コード全体を encodeURIComponent で包んで送り（%エンコードに \ は出ない）、
     * ターゲット側で decodeURIComponent + eval して元ソースに復元してから実行する。
     * @param {string} code - メインエンジンで実行するソース
     * @param {Function} onDone - (status, payload) を受け取るコールバック。status は "ok" | "lines" | "text" | "error"
     * @returns {void}
     */
    function sendWorker(code, onDone) {
        var STATUS_BY_MARKER = { OK: "ok", LINES: "lines", TEXT: "text" };

        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
        bridge.onResult = function (bridgeResult) {
            var body = bridgeResult.body || "";
            var separatorIndex = body.indexOf(":");
            var marker = (separatorIndex >= 0) ? body.substring(0, separatorIndex) : body;
            var payload = (separatorIndex >= 0) ? body.substring(separatorIndex + 1) : "";
            onDone(STATUS_BY_MARKER[marker] || "error", payload);
        };
        bridge.onError = function (bridgeResult) {
            onDone("error", (bridgeResult && bridgeResult.body) ? bridgeResult.body : "BridgeTalk error");
        };
        bridge.send();
    }

    /**
     * メインエンジンへアクションを委譲する（非同期）
     * @param {string} actionId - アクションID
     * @param {Object} params - アクションのパラメーター
     * @param {Function} onDone - (status, payload) を受け取るコールバック
     * @returns {void}
     */
    function runWorker(actionId, params, onDone) {
        var code = getLibSource() + "\nvar __r=" + __DISPATCH_SRC + "(\"" + actionId + "\"," + paramsToSource(params) + ");__r;";
        sendWorker(code, onDone);
    }

    /**
     * state 文字列（"total|point|area|para|forced|tab|multiLines|multiFrames|hasSpTab"）をオブジェクトへ変換する
     * @param {string} encodedState - encodeSelectionState() が返した文字列
     * @returns {Object} 選択状態オブジェクト
     */
    function parseState(encodedState) {
        var fields = String(encodedState || "").split("|");
        return {
            total: parseInt(fields[0], 10) || 0,
            point: parseInt(fields[1], 10) || 0,
            area: parseInt(fields[2], 10) || 0,
            para: parseInt(fields[3], 10) || 0,
            forced: parseInt(fields[4], 10) || 0,
            tab: parseInt(fields[5], 10) || 0,
            multiLines: fields[6] === "1",
            multiFrames: fields[7] === "1",
            hasSpTab: fields[8] === "1"
        };
    }

    /**
     * Option(Alt)キーが押されているかを返す
     * @returns {boolean} 押されていれば true
     */
    function isAltPressed() {
        try {
            return !!(ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey);
        } catch (e) { debugLog("isAltPressed", e); }
        return false;
    }

    /* =========================================
     * 軽量ステータスポーラー（選択のリアルタイム反映用）
     *
     * 毎回 LIB 全体を送ると重いので、件数集計に必要な関数だけを同梱した
     * 軽量本文をメインエンジンへ送る。選択は変更せず、state 文字列だけを返す。
     * ========================================= */
    var __STATE_FUNCS = [
        debugLog, normalizeParagraphBreaks, splitParagraphLines,
        getCharCodeSafe, isParagraphBreak, isForcedBreak, isTabChar,
        getTextFrames, countTextFrameTypes, countBreakTypes,
        hasMultipleLines, hasSpacesOrTabs, computeSelectionState, encodeSelectionState
    ];
    var __STATE_LIB_CACHE = null;

    /* ステータス集計用の関数ソース（初回のみ組み立ててキャッシュ）*/
    function getStateLibSource() {
        if (__STATE_LIB_CACHE === null) __STATE_LIB_CACHE = buildLibSource(__STATE_FUNCS);
        return __STATE_LIB_CACHE;
    }

    function __stateDispatch() {
        if (app.documents.length === 0) return "ERR:nodoc";
        var doc;
        try { doc = app.activeDocument; } catch (e) { return "ERR:nodoc"; }
        return "OK:" + encodeSelectionState(getTextFrames(doc.selection));
    }
    var __STATE_DISPATCH_SRC = "(" + sliceFunctionSource(__stateDispatch) + ")";

    /**
     * 現在の選択をメインエンジンで集計し、state 文字列を受け取る（選択は変えない）
     * @param {Function} onDone - (status, payload) を受け取るコールバック
     * @returns {void}
     */
    function runStatePoll(onDone) {
        var code = getStateLibSource() + "\nvar __r=" + __STATE_DISPATCH_SRC + "();__r;";
        sendWorker(code, onDone);
    }

    /**
     * パレットを組み立てて表示する
     * @param {Array<TextFrame>} selectedObjects - 起動時に選択されていたテキストフレーム
     * @returns {void}
     */
    function showPalette(selectedObjects) {
        var paletteWindow = new Window("palette", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);

        /**
         * 処理をメインエンジンへ委譲し、結果のステータスで表示を更新する
         * @param {string} actionId - アクションID
         * @param {Object} params - アクションのパラメーター
         * @returns {void}
         */
        function executeAction(actionId, params) {
            executeActionThen(actionId, params, null);
        }

        /**
         * executeAction と同じだが、成功後に after() を呼ぶ（行リスト再読込などの連鎖用）
         * @param {string} actionId - アクションID
         * @param {Object} params - アクションのパラメーター
         * @param {Function} after - 成功後に呼ぶコールバック（不要なら null）
         * @returns {void}
         */
        function executeActionThen(actionId, params, after) {
            runWorker(actionId, params || {}, function (status, payload) {
                if (status === "error") {
                    showError({ message: payload });
                    return;
                }
                applySelectionState(parseState(payload));
                if (after) after();
            });
        }

        /**
         * ステータスパネルを構築し、件数表示用の statictext をまとめて返す
         * @returns {Object} 件数表示の statictext を持つオブジェクト
         */
        function buildStatusPanel() {
            var textFrameCounts = countTextFrameTypes(selectedObjects);
            var breakCounts = countBreakTypes(selectedObjects);

            var statusPanel = addPanel(paletteWindow, getLabel(LABELS.panel.status), ["left", "top"]);
            statusPanel.margins = STATUS_MARGINS;

            var statusRow = statusPanel.add("group");
            statusRow.orientation = "row";
            statusRow.alignment = ["fill", "top"];
            statusRow.alignChildren = ["left", "center"];
            statusRow.spacing = STATUS_ROW_SPACING;

            var frameCountColumn = addColumnGroup(statusRow);
            frameCountColumn.alignChildren = ["left", "top"];
            var breakCountColumn = addColumnGroup(statusRow);
            breakCountColumn.alignChildren = ["left", "top"];

            /**
             * 1行 = [ラベル：] [値] を作り、値側の statictext を返す
             * @param {Group} column - 追加先のカラム
             * @param {Object} labelNode - LABELS のラベルノード
             * @param {number} value - 初期表示する件数
             * @param {number} labelWidth - ラベル幅（省略時は自動）
             * @returns {StaticText} 値表示用の statictext
             */
            function addStatusRow(column, labelNode, value, labelWidth) {
                var statusFieldRow = column.add("group");
                statusFieldRow.orientation = "row";
                statusFieldRow.alignChildren = ["left", "center"];

                var fieldLabel = statusFieldRow.add("statictext", undefined, getColonLabel(labelNode));
                if (labelWidth) fieldLabel.preferredSize.width = labelWidth;
                return statusFieldRow.add("statictext", undefined, String(value));
            }

            return {
                target: addStatusRow(frameCountColumn, LABELS.info.targetCount, textFrameCounts.total),
                point: addStatusRow(frameCountColumn, LABELS.info.pointCount, textFrameCounts.point),
                area: addStatusRow(frameCountColumn, LABELS.info.areaCount, textFrameCounts.area),
                para: addStatusRow(breakCountColumn, LABELS.info.paragraphBreak, breakCounts.paragraph, STATUS_LABEL_WIDTH),
                forced: addStatusRow(breakCountColumn, LABELS.info.forcedBreak, breakCounts.forced, STATUS_LABEL_WIDTH),
                tab: addStatusRow(breakCountColumn, LABELS.info.tab, breakCounts.tab, STATUS_LABEL_WIDTH)
            };
        }

        var statusFields = buildStatusPanel();

        /**
         * 選択状態をステータス表示と各ボタンの有効・無効へ反映する
         * @param {Object} state - parseState() が返した選択状態
         * @returns {void}
         */
        function applySelectionState(state) {
            statusFields.target.text = String(state.total);
            statusFields.point.text = String(state.point);
            statusFields.area.text = String(state.area);
            statusFields.para.text = String(state.para);
            statusFields.forced.text = String(state.forced);
            statusFields.tab.text = String(state.tab);

            var hasParagraph = state.para > 0;
            var hasForced = state.forced > 0;
            var hasAnyBreaks = hasParagraph || hasForced;
            var hasTabs = state.tab > 0;

            /* 行操作系：2行以上あるときだけ */
            btnReverseOrder.enabled = state.multiLines;
            btnRemoveDuplicateLines.enabled = state.multiLines;
            btnRemoveEmptyLines.enabled = state.multiLines;
            btnSortByCharCode.enabled = state.multiLines;
            btnSortByLength.enabled = state.multiLines;
            btnSplitByLine.enabled = state.multiLines;
            btnSplitByLineKeepStyle.enabled = state.multiLines;

            /* 改行の削除・変換 */
            btnRemoveLineBreaks.enabled = hasAnyBreaks;
            chkIncludeForcedBreaks.enabled = hasForced;
            btnConvertBreaks.enabled = hasForced;
            btnConvertToForcedBreaks.enabled = hasParagraph;
            btnFlattenToOneLine.enabled = hasAnyBreaks || state.multiFrames;

            /* タブ */
            btnSplitByTab.enabled = hasTabs;
            btnRemoveTabs.enabled = hasTabs;
            btnTabsToSpaces.enabled = hasTabs;

            /* 連結：2つ以上のフレームが必要 */
            btnConcatV.enabled = state.multiFrames;
            btnConcatHOnly.enabled = state.multiFrames;
            btnConcatH.enabled = state.multiFrames;
            btnConcatToArea.enabled = state.multiFrames;

            /* スペース */
            btnTrimSpaces.enabled = state.hasSpTab;
            btnCjkLatinSpaces.enabled = state.hasSpTab;
            btnCollapseSpaces.enabled = state.hasSpTab;
            btnCleanupSpaces.enabled = state.hasSpTab;
            btnRemoveAllSpaces.enabled = state.hasSpTab;
        }

        /* 制御文字の表示状態 */
        var hiddenCharOn = false;
        var hiddenCharLabel = getLabel(LABELS.button.showHiddenChar);

        var tabbedPanel = paletteWindow.add("tabbedpanel");
        tabbedPanel.alignment = ["fill", "top"];
        tabbedPanel.alignChildren = ["fill", "top"];

        /* === タブ1: 基本 === */
        var tabBasic = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.basic));
        setupTab(tabBasic, "row", ["center", "top"]);

        /* 左カラム：改行 */
        var breakColumn = addColumnGroup(tabBasic);
        var panelBreakGroup = addPanel(breakColumn, getLabel(LABELS.panel.breakGroup), ["fill", "top"]);

        var flattenButtonRow = panelBreakGroup.add("group");
        flattenButtonRow.alignment = ["fill", "top"];
        flattenButtonRow.alignChildren = ["fill", "center"];
        flattenButtonRow.margins = BUTTON_ROW_MARGINS;

        var btnFlattenToOneLine = flattenButtonRow.add("button", undefined, getLabel(LABELS.button.flattenToOneLine));
        btnFlattenToOneLine.onClick = function () {
            executeAction("flatten");
        };

        /* 改行の削除 */
        var panelRemoveBreak = addPanel(panelBreakGroup, getLabel(LABELS.panel.removeBreak));

        var btnRemoveLineBreaks = panelRemoveBreak.add("button", undefined, getLabel(LABELS.button.removeLineBreaks));
        btnRemoveLineBreaks.helpTip = getLabel(LABELS.tooltip.removeLineBreaks);
        btnRemoveLineBreaks.onClick = function () {
            executeAction("removeLineBreaks", { forced: chkIncludeForcedBreaks.value });
        };

        var chkIncludeForcedBreaks = panelRemoveBreak.add("checkbox", undefined, getLabel(LABELS.checkbox.includeForcedBreaks));

        /* 改行の挿入 */
        var panelInsertBreak = addPanel(panelBreakGroup, getLabel(LABELS.panel.insertBreak));

        var btnAddLineBreaks = panelInsertBreak.add("button", undefined, getLabel(LABELS.button.addLineBreaks));
        btnAddLineBreaks.onClick = function () {
            executeAction("addLineBreakPerChar");
        };

        var btnBreakAtChars = panelInsertBreak.add("button", undefined, getLabel(LABELS.button.breakAtChars));
        btnBreakAtChars.onClick = function () {
            executeAction("punctuation", { chars: txtBreakChars.text });
        };

        var txtBreakChars = panelInsertBreak.add("edittext", undefined, DEFAULT_BREAK_CHARS);
        txtBreakChars.alignment = ["fill", "center"];

        var btnBreakAtCount = panelInsertBreak.add("button", undefined, getLabel(LABELS.button.breakAtCount));
        btnBreakAtCount.onClick = function () {
            executeAction("breakAtCount", { count: txtBreakCount.text, forced: chkForcedBreakAtCount.value });
        };

        var breakCountRow = panelInsertBreak.add("group");
        breakCountRow.orientation = "row";
        breakCountRow.alignment = ["fill", "center"];
        breakCountRow.alignChildren = ["left", "center"];
        var txtBreakCount = breakCountRow.add("edittext", undefined, DEFAULT_BREAK_COUNT);
        txtBreakCount.characters = 3;
        var chkForcedBreakAtCount = breakCountRow.add("checkbox", undefined, getLabel(LABELS.checkbox.forcedBreak));

        /* 改行の切り換え */
        var panelConvertBreak = addPanel(panelBreakGroup, getLabel(LABELS.panel.convertBreak));

        var btnConvertBreaks = panelConvertBreak.add("button", undefined, getLabel(LABELS.button.convertBreaks));
        btnConvertBreaks.onClick = function () {
            executeAction("convertForcedLineBreaks");
        };

        var btnConvertToForcedBreaks = panelConvertBreak.add("button", undefined, getLabel(LABELS.button.convertToForcedBreaks));
        btnConvertToForcedBreaks.onClick = function () {
            executeAction("convertToForcedBreaks");
        };

        /* 右カラム：分割・連結 */
        var splitConcatColumn = addColumnGroup(tabBasic);
        var panelSplitGroup = addPanel(splitConcatColumn, getLabel(LABELS.panel.splitGroup), ["fill", "top"]);

        /* 改行・タブで分割 */
        var panelSplitByBreak = addPanel(panelSplitGroup, getLabel(LABELS.panel.splitByBreak));

        var btnSplitByLine = panelSplitByBreak.add("button", undefined, getLabel(LABELS.button.splitByLine));
        btnSplitByLine.helpTip = getLabel(LABELS.tooltip.splitByLine);
        btnSplitByLine.onClick = function () {
            executeAction("splitByLineBreak");
        };

        var btnSplitByLineKeepStyle = panelSplitByBreak.add("button", undefined, getLabel(LABELS.button.splitByLineKeepStyle));
        btnSplitByLineKeepStyle.helpTip = getLabel(LABELS.tooltip.splitByLineKeepStyle);
        btnSplitByLineKeepStyle.onClick = function () {
            executeAction("splitByLineBreakKeepStyle");
        };

        var btnSplitByTab = panelSplitByBreak.add("button", undefined, getLabel(LABELS.button.splitByTab));
        btnSplitByTab.helpTip = getLabel(LABELS.tooltip.splitByTab);
        btnSplitByTab.onClick = function () {
            /* Option+クリックのときは分割結果をグループにまとめない */
            executeAction("splitByTab", { ungroup: isAltPressed() });
        };

        /* 1文字ずつ分割 */
        var panelSplitByChar = addPanel(panelSplitGroup, getLabel(LABELS.panel.splitByChar));

        var btnSplitKeepStyle = panelSplitByChar.add("button", undefined, getLabel(LABELS.button.splitKeepStyle));
        btnSplitKeepStyle.helpTip = getLabel(LABELS.tooltip.splitKeepStyle);
        btnSplitKeepStyle.onClick = function () {
            executeAction("splitByCharKeepStyle");
        };

        var btnSplitIgnoreStyle = panelSplitByChar.add("button", undefined, getLabel(LABELS.button.splitIgnoreStyle));
        btnSplitIgnoreStyle.helpTip = getLabel(LABELS.tooltip.splitIgnoreStyle);
        btnSplitIgnoreStyle.onClick = function () {
            executeAction("splitByCharIgnoreStyle");
        };

        /* 連結 */
        var panelConcat = addPanel(splitConcatColumn, getLabel(LABELS.panel.concat), ["center", "center"]);

        var concatButtonColumn = panelConcat.add("group");
        concatButtonColumn.margins = BUTTON_ROW_MARGINS;
        concatButtonColumn.orientation = "column";
        concatButtonColumn.alignment = ["fill", "top"];
        concatButtonColumn.alignChildren = ["fill", "center"];

        var btnConcatV = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatV));
        btnConcatV.helpTip = getLabel(LABELS.tooltip.concatV);
        btnConcatV.onClick = function () {
            executeAction("concatVertical");
        };

        var btnConcatHOnly = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatHOnly));
        btnConcatHOnly.helpTip = getLabel(LABELS.tooltip.concatHOnly);
        btnConcatHOnly.onClick = function () {
            executeAction("concatHorizontalOnly");
        };

        var btnConcatH = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatH));
        btnConcatH.helpTip = getLabel(LABELS.tooltip.concatH);
        btnConcatH.onClick = function () {
            executeAction("concatH");
        };

        var btnConcatToArea = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatToArea));
        btnConcatToArea.helpTip = getLabel(LABELS.tooltip.concatToArea);
        btnConcatToArea.onClick = function () {
            executeAction("concatToArea");
        };

        /* === タブ2: 整形 === */
        var tabCleanup = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.cleanup));
        setupTab(tabCleanup, "row");

        /* 左カラム：タブ・スペース */
        var spaceCleanupColumn = addColumnGroup(tabCleanup);

        var panelTabChar = addPanel(spaceCleanupColumn, getLabel(LABELS.panel.tab));

        var btnRemoveTabs = panelTabChar.add("button", undefined, getLabel(LABELS.button.removeTabs));
        btnRemoveTabs.onClick = function () {
            executeAction("removeTabs");
        };

        var btnTabsToSpaces = panelTabChar.add("button", undefined, getLabel(LABELS.button.tabsToSpaces));
        btnTabsToSpaces.onClick = function () {
            executeAction("tabsToSpaces");
        };

        var panelRemoveSpace = addPanel(spaceCleanupColumn, getLabel(LABELS.panel.space));

        var btnTrimSpaces = panelRemoveSpace.add("button", undefined, getLabel(LABELS.button.trimSpaces));
        btnTrimSpaces.helpTip = getLabel(LABELS.tooltip.trimSpaces);
        btnTrimSpaces.onClick = function () {
            executeAction("trimSpaces");
        };

        var btnCjkLatinSpaces = panelRemoveSpace.add("button", undefined, getLabel(LABELS.button.cjkLatinSpaces));
        btnCjkLatinSpaces.helpTip = getLabel(LABELS.tooltip.cjkLatinSpaces);
        btnCjkLatinSpaces.onClick = function () {
            executeAction("removeCjkLatinSpaces");
        };

        var btnCollapseSpaces = panelRemoveSpace.add("button", undefined, getLabel(LABELS.button.collapseSpaces));
        btnCollapseSpaces.helpTip = getLabel(LABELS.tooltip.collapseSpaces);
        btnCollapseSpaces.onClick = function () {
            executeAction("collapseSpaces");
        };

        var btnCleanupSpaces = panelRemoveSpace.add("button", undefined, getLabel(LABELS.button.cleanupSpaces));
        btnCleanupSpaces.helpTip = getLabel(LABELS.tooltip.cleanupSpaces);
        btnCleanupSpaces.onClick = function () {
            executeAction("cleanupSpaces");
        };

        var btnRemoveAllSpaces = panelRemoveSpace.add("button", undefined, getLabel(LABELS.button.removeAllSpaces));
        btnRemoveAllSpaces.helpTip = getLabel(LABELS.tooltip.removeAllSpaces);
        btnRemoveAllSpaces.onClick = function () {
            executeAction("removeAllSpaces");
        };

        var panelAddSpace = addPanel(spaceCleanupColumn, getLabel(LABELS.panel.addSpace));

        var btnSpaceAfterPunct = panelAddSpace.add("button", undefined, getLabel(LABELS.button.spaceAfterPunct));
        btnSpaceAfterPunct.helpTip = getLabel(LABELS.tooltip.spaceAfterPunct);
        btnSpaceAfterPunct.onClick = function () {
            executeAction("spaceAfterPunct");
        };

        /* 右カラム：変換・リスト */
        var convertListColumn = addColumnGroup(tabCleanup);

        /* スペースや記号の変換：変換前／変換後をラジオで選び、［変換］で置換する */
        var panelSymbolConvert = addPanel(convertListColumn, getLabel(LABELS.panel.symbolConvert), ["fill", "top"]);

        var symbolConvertColumn = panelSymbolConvert.add("group");
        symbolConvertColumn.orientation = "column";
        symbolConvertColumn.alignment = ["fill", "top"];
        symbolConvertColumn.alignChildren = ["fill", "top"];
        symbolConvertColumn.spacing = TAB_SPACING;

        var panelSymbolBefore = addPanel(symbolConvertColumn, getLabel(LABELS.panel.symbolBefore), ["left", "top"]);
        var rbBeforeSpace = panelSymbolBefore.add("radiobutton", undefined, getLabel(LABELS.radio.space));
        var rbBeforeUnderscore = panelSymbolBefore.add("radiobutton", undefined, getLabel(LABELS.radio.underscore));
        var rbBeforeHyphen = panelSymbolBefore.add("radiobutton", undefined, getLabel(LABELS.radio.hyphen));
        rbBeforeSpace.value = true;

        var panelSymbolAfter = addPanel(symbolConvertColumn, getLabel(LABELS.panel.symbolAfter), ["left", "top"]);
        var rbAfterSpace = panelSymbolAfter.add("radiobutton", undefined, getLabel(LABELS.radio.space));
        var rbAfterUnderscore = panelSymbolAfter.add("radiobutton", undefined, getLabel(LABELS.radio.underscore));
        var rbAfterHyphen = panelSymbolAfter.add("radiobutton", undefined, getLabel(LABELS.radio.hyphen));
        rbAfterUnderscore.value = true;

        /**
         * 変換前に選ばれている記号種別を返す
         * @returns {string} "space" | "underscore" | "hyphen"
         */
        function getBeforeSymbol() {
            if (rbBeforeUnderscore.value) return "underscore";
            if (rbBeforeHyphen.value) return "hyphen";
            return "space";
        }

        /**
         * 変換後に選ばれている記号種別を返す
         * @returns {string} "space" | "underscore" | "hyphen"
         */
        function getAfterSymbol() {
            if (rbAfterSpace.value) return "space";
            if (rbAfterHyphen.value) return "hyphen";
            return "underscore";
        }

        var btnConvertSymbol = panelSymbolConvert.add("button", undefined, getLabel(LABELS.button.convertSymbol));
        btnConvertSymbol.alignment = ["center", "top"];
        btnConvertSymbol.onClick = function () {
            executeAction("convertSymbol", { from: getBeforeSymbol(), to: getAfterSymbol() });
        };

        /**
         * 変換前と変換後が同じ記号なら［変換］をディムにする
         * @returns {void}
         */
        function updateConvertSymbolState() {
            btnConvertSymbol.enabled = (getBeforeSymbol() !== getAfterSymbol());
        }

        var symbolRadios = [rbBeforeSpace, rbBeforeUnderscore, rbBeforeHyphen, rbAfterSpace, rbAfterUnderscore, rbAfterHyphen];
        for (var radioIndex = 0; radioIndex < symbolRadios.length; radioIndex++) {
            symbolRadios[radioIndex].onClick = updateConvertSymbolState;
        }
        updateConvertSymbolState();

        /* 文字変換 */
        var panelConvert = addPanel(convertListColumn, getLabel(LABELS.panel.convert), ["center", "center"]);

        var btnFullToHalfAlnum = panelConvert.add("button", undefined, getLabel(LABELS.button.fullToHalfAlnum));
        btnFullToHalfAlnum.onClick = function () {
            executeAction("fullToHalfAlnum");
        };

        var btnHalfToFullKana = panelConvert.add("button", undefined, getLabel(LABELS.button.halfToFullKana));
        btnHalfToFullKana.onClick = function () {
            executeAction("halfToFullKana");
        };

        /* リストの除去 */
        var panelRemoveList = addPanel(convertListColumn, getLabel(LABELS.panel.list));
        panelRemoveList.orientation = "row";

        var btnBulletList = panelRemoveList.add("button", undefined, getLabel(LABELS.button.bulletList));
        btnBulletList.helpTip = getLabel(LABELS.tooltip.bulletList);
        btnBulletList.onClick = function () {
            executeAction("removeBulletMarkers");
        };

        var btnNumberList = panelRemoveList.add("button", undefined, getLabel(LABELS.button.numberList));
        btnNumberList.helpTip = getLabel(LABELS.tooltip.numberList);
        btnNumberList.onClick = function () {
            executeAction("removeNumberMarkers");
        };

        /* === タブ3: 行の編集 === */
        var tabLineArrange = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.lineArrange));
        setupTab(tabLineArrange, "row");

        /* 左カラム：行リスト */
        var lineListColumn = addColumnGroup(tabLineArrange);

        var lineListRow = lineListColumn.add("group");
        lineListRow.orientation = "row";
        lineListRow.alignChildren = ["fill", "fill"];
        lineListRow.spacing = 10;

        var lineListBox = lineListRow.add("listbox", undefined, [], { multiselect: false });
        lineListBox.preferredSize = LINE_LIST_SIZE;
        lineListBox.graphics.font = ScriptUI.newFont("dialog", "REGULAR", LINE_LIST_FONT_SIZE);

        /* リストボックスが保持する行データ */
        var lineArrangeLines = [];

        /**
         * リストの内容をテキストフレームへ書き戻す
         * @returns {void}
         */
        function applyLinesToTextFrame() {
            executeAction("setLines", { text: lineArrangeLines.join("\r") });
        }

        /**
         * リストボックスを再構築する
         * @param {number} selectIndex - 再構築後に選択する行番号
         * @param {boolean} skipApply - 真ならテキストフレームへの書き戻しを行わない
         * @returns {void}
         */
        function refreshLineList(selectIndex, skipApply) {
            lineListBox.removeAll();
            for (var i = 0; i < lineArrangeLines.length; i++) {
                lineListBox.add("item", lineArrangeLines[i]);
            }
            if (lineArrangeLines.length > 0) {
                if (selectIndex < 0) selectIndex = 0;
                if (selectIndex >= lineArrangeLines.length) selectIndex = lineArrangeLines.length - 1;
                lineListBox.selection = selectIndex;
            }
            updateLineListButtons();
            if (!skipApply) applyLinesToTextFrame();
        }

        /**
         * 選択中テキストフレームの内容を読み込んでリストへ反映する
         * @returns {void}
         */
        function loadLinesToList() {
            runWorker("getLines", {}, function (status, payload) {
                lineArrangeLines = [];
                if (status === "lines" && payload) {
                    lineArrangeLines = normalizeParagraphBreaks(decodeURIComponent(payload)).split("\r");
                }
                refreshLineList(0, true);
            });
        }

        /**
         * 行リストの操作ボタンの有効・無効を選択状態に合わせる
         * @returns {void}
         */
        function updateLineListButtons() {
            var hasSelection = lineListBox.selection !== null;
            btnLineUp.enabled = hasSelection && lineListBox.selection.index > 0;
            btnLineDown.enabled = hasSelection && lineListBox.selection.index < lineArrangeLines.length - 1;
            btnLineEdit.enabled = hasSelection;
            btnLineDelete.enabled = hasSelection;
        }

        /**
         * 選択行をダイアログで編集する
         * @returns {void}
         */
        function editSelectedLine() {
            if (!lineListBox.selection) return;
            var selectedIndex = lineListBox.selection.index;
            var editedText = prompt(getLabel(LABELS.prompt.editLine), lineArrangeLines[selectedIndex]);
            if (editedText === null) return;
            lineArrangeLines[selectedIndex] = editedText;
            refreshLineList(selectedIndex);
        }

        /**
         * 選択行を上下いずれかへ1行ぶん移動する
         * @param {number} offset - -1 なら上へ、1 なら下へ
         * @returns {void}
         */
        function moveSelectedLine(offset) {
            if (!lineListBox.selection) return;
            var selectedIndex = lineListBox.selection.index;
            var movedIndex = selectedIndex + offset;
            if (movedIndex < 0 || movedIndex >= lineArrangeLines.length) return;

            var movedLine = lineArrangeLines[selectedIndex];
            lineArrangeLines[selectedIndex] = lineArrangeLines[movedIndex];
            lineArrangeLines[movedIndex] = movedLine;
            refreshLineList(movedIndex);
        }

        lineListBox.onDoubleClick = editSelectedLine;
        lineListBox.onChange = updateLineListButtons;

        /* 右カラム：行への一括操作 */
        var lineActionColumn = addColumnGroup(tabLineArrange);

        var panelLineEdit = addPanel(lineActionColumn, getLabel(LABELS.panel.lineEdit));

        var btnLineUp = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineUp));
        btnLineUp.onClick = function () {
            moveSelectedLine(-1);
        };

        var btnLineDown = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineDown));
        btnLineDown.onClick = function () {
            moveSelectedLine(1);
        };

        var btnLineAdd = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineAdd));
        btnLineAdd.onClick = function () {
            var addedText = prompt(getLabel(LABELS.prompt.addLine), "");
            if (addedText === null) return;
            lineArrangeLines.push(addedText);
            refreshLineList(lineArrangeLines.length - 1);
        };

        var btnLineEdit = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineEdit));
        btnLineEdit.onClick = editSelectedLine;

        var btnLineDelete = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineDelete));
        btnLineDelete.onClick = function () {
            if (!lineListBox.selection) return;
            var selectedIndex = lineListBox.selection.index;
            if (!confirm(getLabel(LABELS.confirm.deleteLine))) return;
            lineArrangeLines.splice(selectedIndex, 1);
            refreshLineList(selectedIndex);
        };

        var panelSort = addPanel(lineActionColumn, getLabel(LABELS.panel.sort));

        var btnSortByCharCode = panelSort.add("button", undefined, getLabel(LABELS.button.sortByCharCode));
        btnSortByCharCode.onClick = function () {
            executeActionThen("sortByCharCode", {}, loadLinesToList);
        };

        var btnSortByLength = panelSort.add("button", undefined, getLabel(LABELS.button.sortByLength));
        btnSortByLength.onClick = function () {
            executeActionThen("sortByLength", {}, loadLinesToList);
        };

        var btnReverseOrder = panelSort.add("button", undefined, getLabel(LABELS.button.reverseOrder));
        btnReverseOrder.onClick = function () {
            executeActionThen("reverseOrder", {}, loadLinesToList);
        };

        var panelLineDelete = addPanel(lineActionColumn, getLabel(LABELS.panel.lineDelete));

        var btnRemoveDuplicateLines = panelLineDelete.add("button", undefined, getLabel(LABELS.button.removeDuplicateLines));
        btnRemoveDuplicateLines.onClick = function () {
            executeActionThen("removeDuplicateLines", {}, loadLinesToList);
        };

        var btnRemoveEmptyLines = panelLineDelete.add("button", undefined, getLabel(LABELS.button.removeEmptyLines));
        btnRemoveEmptyLines.onClick = function () {
            executeActionThen("removeEmptyLines", {}, loadLinesToList);
        };

        /* === タブ4: 英数字 === */
        var tabAlnum = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.alnum));
        setupTab(tabAlnum, "column");

        var panelLetterCase = addPanel(tabAlnum, getLabel(LABELS.panel.letterCase), ["fill", "top"]);

        /* モード名 → プレビュー用 statictext */
        var casePreviewFields = {};

        /**
         * 1行 = [変換ボタン] [変換結果プレビュー] を追加する
         * @param {string} modeKey - プレビューを引くためのモード名
         * @param {Object} labelNode - ボタンのラベルノード
         * @param {string} actionId - 実行するアクションID
         * @returns {void}
         */
        function addCaseRow(modeKey, labelNode, actionId) {
            var caseRow = panelLetterCase.add("group");
            caseRow.orientation = "row";
            caseRow.alignment = ["fill", "center"];
            caseRow.alignChildren = ["left", "center"];

            var caseButton = caseRow.add("button", undefined, getLabel(labelNode));
            caseButton.preferredSize = CASE_BUTTON_SIZE;
            caseButton.onClick = function () {
                executeActionThen(actionId, {}, refreshCasePreview);
            };

            var previewText = caseRow.add("statictext", undefined, "", { justify: "left" });
            previewText.preferredSize = CASE_PREVIEW_SIZE;
            casePreviewFields[modeKey] = previewText;
        }

        addCaseRow("upper", LABELS.button.caseUpper, "caseUpper");
        addCaseRow("lower", LABELS.button.caseLower, "caseLower");
        addCaseRow("word", LABELS.button.caseWord, "caseWord");
        addCaseRow("sentence", LABELS.button.caseSentence, "caseSentence");
        addCaseRow("title", LABELS.button.caseTitle, "caseTitle");

        /**
         * プレビュー表示用にテキストを1行へ詰めて短く整える
         * @param {string} text - 元テキスト
         * @returns {string} 整形後のテキスト
         */
        function normalizeCaseSample(text) {
            if (text == null) return "";
            var sample = String(text).replace(/[\r\n]+/g, " ").replace(/[ 　\t]+/g, " ").replace(/^\s+|\s+$/g, "");
            if (sample.length > CASE_PREVIEW_MAX_CHARS) sample = sample.substring(0, CASE_PREVIEW_MAX_CHARS) + "…";
            return sample;
        }

        /* プレビュー更新の多重実行ガード */
        var casePreviewRefreshing = false;

        /**
         * メインエンジンから先頭テキストを取得し、各モードのプレビューを更新する
         * @returns {void}
         */
        function refreshCasePreview() {
            if (casePreviewRefreshing) return;
            casePreviewRefreshing = true;

            runWorker("getFirstText", {}, function (status, payload) {
                casePreviewRefreshing = false;
                var sample = normalizeCaseSample((status === "text" && payload) ? decodeURIComponent(payload) : "");
                casePreviewFields.upper.text = normalizeCaseSample(sample.toUpperCase());
                casePreviewFields.lower.text = normalizeCaseSample(sample.toLowerCase());
                casePreviewFields.word.text = normalizeCaseSample(toWordCap(sample));
                casePreviewFields.sentence.text = normalizeCaseSample(toSentenceCase(sample));
                casePreviewFields.title.text = normalizeCaseSample(toTitleCase(sample));
            });
        }

        applySelectionState(computeSelectionState(selectedObjects));
        loadLinesToList();
        refreshCasePreview();

        /* === 選択のリアルタイム反映 ===
         * Illustrator 30.x には app.scheduleTask / setTimeout 等のタイマー API が無いため、
         * キャンバスにフォーカスがある間の連続ポーリングはできない。代わりに
         * 「ユーザーがパレットへ操作しに来た瞬間」（onActivate＝フォーカス復帰、mouseover）に
         * メインエンジンの選択を取り直してステータスへ反映する。
         * mouseover はパレット上でマウスを動かすたびに何度も発生するため、
         * 多重実行ガード（statusRefreshing）に加えて STATUS_POLL_INTERVAL_MS で間引く。*/
        var statusRefreshing = false;
        var lastFocusRefreshTime = 0;

        /**
         * メインエンジンの選択を取り直してステータス表示へ反映する
         * @returns {void}
         */
        function refreshStatusFromSelection() {
            if (statusRefreshing) return;
            statusRefreshing = true;

            runStatePoll(function (status, payload) {
                statusRefreshing = false;
                if (status !== "ok") return;
                applySelectionState(parseState(payload));
            });
        }

        /**
         * パレットがフォーカスを得たときの更新処理
         * @param {boolean} force - 真なら間引きを無視して必ず取り直す
         * @returns {void}
         */
        function onPaletteFocus(force) {
            var now = (new Date()).getTime();
            if (!force && (now - lastFocusRefreshTime) < STATUS_POLL_INTERVAL_MS) return;
            lastFocusRefreshTime = now;

            refreshStatusFromSelection();
            /* 英数字タブ表示中は選択変化に合わせてプレビューも更新 */
            if (tabbedPanel.selection === tabAlnum) refreshCasePreview();
        }

        paletteWindow.onActivate = function () { onPaletteFocus(true); };
        try {
            paletteWindow.addEventListener("mouseover", function () { onPaletteFocus(false); });
        } catch (e) { debugLog("addEventListener: mouseover", e); }

        tabbedPanel.onChange = function () {
            if (tabbedPanel.selection === tabLineArrange) {
                loadLinesToList();
            } else if (tabbedPanel.selection === tabAlnum) {
                refreshCasePreview();
            }
        };

        /**
         * パレットを閉じる（1要素だけのグループ解除・選択整理・制御文字OFF はメインエンジンで実行）
         * @returns {void}
         */
        function closePalette() {
            runWorker("finalizeClose", { turnOffHidden: hiddenCharOn }, function () {
                paletteWindow.close();
            });
        }

        /* フッター：制御文字の表示切り換え */
        var footerRow = paletteWindow.add("group");
        footerRow.orientation = "row";
        footerRow.alignment = ["fill", "bottom"];
        footerRow.alignChildren = ["left", "center"];
        footerRow.margins = FOOTER_MARGINS;

        var btnShowHiddenChar = footerRow.add("button", undefined, hiddenCharLabel);

        /**
         * 制御文字ボタンの文字色を状態に合わせる（ラベルは固定）
         * @returns {void}
         */
        function updateHiddenCharButton() {
            try {
                var buttonGraphics = btnShowHiddenChar.graphics;
                var textColor = hiddenCharOn ? [0.0, 0.5, 0.8] : [0.0, 0.0, 0.0];
                buttonGraphics.foregroundColor = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, textColor, 1);
            } catch (e) { debugLog("updateHiddenCharButton: set color", e); }
        }

        btnShowHiddenChar.onClick = function () {
            runWorker("hiddenChar", {}, function (status, payload) {
                if (status === "error") {
                    showError({ message: payload });
                    return;
                }
                hiddenCharOn = !hiddenCharOn;
                updateHiddenCharButton();
                applySelectionState(parseState(payload));
            });
        };

        /* パレットがアクティブなとき Esc で閉じる */
        paletteWindow.addEventListener("keydown", function (keyEvent) {
            if (keyEvent.keyName === "Escape") closePalette();
        });

        /* パレットが GC で破棄されないよう参照を保持 */
        $.global.__TextBreakSplitMergePalette = paletteWindow;
        paletteWindow.onClose = function () {
            $.global.__TextBreakSplitMergePalette = null;
        };

        paletteWindow.show();
    }

    try {
        showPalette(selectedObjects);
    } catch (err) {
        showError(err);
    }
})();