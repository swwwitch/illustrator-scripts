#targetengine "TextBreakSplitMergeEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustrator で分散しがちなテキスト処理（改行・分割・連結・整形）を、1つのパレットに統合したオールインワンのテキスト処理スクリプトです。
従来は個別スクリプトに分散しがちな処理（改行削除・空行整理・分割・連結など）を単一 UI に集約し、スクリプト管理やショートカット運用の負担を軽減します。

### 対象

- テキストフレーム
- テキストフレームを含むグループ
- Illustrator が TextRange として返すテキスト選択

### 主な機能

- 改行の削除／挿入／変換（段落改行・強制改行の相互変換、指定文字・指定文字数での改行、1文字ごとの改行）
- 段落／タブ／文字単位でのテキスト分割（書式保持／無視の両対応）
- 複数テキストの縦方向・横方向連結（見た目ベースでの再構成、PDF テキスト整形）
- スペース整理（行頭行末・和欧間・連続・まとめて・すべて削除）、タブ処理
- スペースや記号の相互変換（スペース／アンダースコア／ハイフンを Before→After で指定）、`.` `,` の後にスペース挿入
- 文字変換（全角英数→半角、半角カナ→全角）、リスト記号の除去
- 英字のケース変換（すべて大文字／小文字／単語の先頭のみ／文頭のみ／英語タイトル形式）＋変換プレビュー
- 行単位の編集・並べ替え・ソート・重複削除・空行削除
- テキスト構造の可視化（改行数・タブ数・文字種別の集計表示、選択に追従）
- 実行可能な処理だけを有効化する状態連動 UI

### UI タブ構成

- 基本：改行（削除・挿入・変換）、分割、連結
- クリーンアップ：タブ・スペース処理、スペースや記号の変換、その他、文字変換、リスト記号除去
- 行の編集/ソート：行単位の並べ替え・追加・削除・ソート・重複除去
- 英数字：英字のケース変換（変換結果のプレビュー付き）

### 設計方針

- 既存テキストを直接編集しつつ、結果を即座に確認できる即時実行型 UI
- 複数スクリプトの置き換えを前提とした「集約ツール」設計
- 見た目ベースの処理を優先し、実務での作業効率を重視
- 実行条件に合わない機能は無効化し、誤操作を防止
- パレットは常駐エンジン（#targetengine）で常駐表示。常駐エンジンの app はパレット表示中に DOM 接続を失うため、DOM 処理はメインエンジンへ BridgeTalk で都度委譲（コードは encodeURIComponent で包んで送信）
- ステータスはタイマー API が無いため、パレットへのフォーカス／マウスオーバー時に選択へ追従して更新
- Esc キーでパレットを閉じる

### 紹介記事

https://note.com/dtp_tranist/n/nf6f34559ba46

### メモ

- 作成日：2026-03-18 ／ 更新日：2026-06-15
- UI・メッセージは日本語／英語を自動切り替え（ロケール判定）

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.7.0";

// =========================================
// ローカライズ / Localization
// =========================================
/* 現在の言語を判定（ロケールが ja 始まりなら日本語）/ Detect UI language (Japanese if locale starts with "ja") */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Bilingual labels */
var LABELS = {
    dialog: {
        title: { ja: "テキスト処理", en: "Text Processing" }
    },
    tab: {
        basic: { ja: "基本", en: "Basic" },
        cleanup: { ja: "クリーンアップ", en: "Cleanup" },
        lineArrange: { ja: "行の編集{slash}ソート", en: "Line Edit{slash}Sort" },
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
        list: { ja: "リスト", en: "List" },
        concat: { ja: "連結", en: "Concatenate" },
        tab: { ja: "タブ", en: "Tab" },
        space: { ja: "スペース削除", en: "Remove Spaces" },
        status: { ja: "ステータス", en: "Status" },
        letterCase: { ja: "大文字{slash}小文字", en: "Letter Case" },
        symbolConvert: { ja: "スペースや記号の変換", en: "Spaces & Symbols" },
        symbolBefore: { ja: "変換前", en: "Before" },
        symbolAfter: { ja: "変換後", en: "After" },
        other: { ja: "その他", en: "Other" }
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
        punctuation: { ja: "指定文字で改行", en: "At Specified Characters" },
        breakAtCount: { ja: "指定文字数で改行", en: "At Character Count" },
        convertBreaks: { ja: "強制改行→改行", en: "Forced Breaks to Paragraph Breaks" },
        convertToForcedBreaks: { ja: "改行→強制改行", en: "Paragraph Breaks to Forced Breaks" },
        splitByLine: { ja: "テキストばらし", en: "Split by Line Breaks" },
        splitByLineKeepStyle: { ja: "〃（書式保持）", en: "Split by Line Breaks (Keep Style)" },
        splitByTab: { ja: "タブで分割", en: "Split by Tabs and Line Breaks" },
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
        caseWord: { ja: "単語の先頭のみ大文字", en: "Capitalize Each Word" },
        caseSentence: { ja: "文頭のみ大文字", en: "Sentence case" },
        caseTitle: { ja: "英語タイトル形式", en: "Title Case (English)" },
        convertSymbol: { ja: "変換", en: "Convert" },
        spaceAfterPunct: { ja: ".と,の後にスペース", en: "Space After . and ," },
        undo: { ja: "1つ戻す", en: "Undo" },
        close: { ja: "閉じる", en: "Close" }
    },
    checkbox: {
        showHiddenChar: { ja: "制御文字", en: "Hidden Characters" },
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

/* ラベルノードから現在言語の文言を返す（{slash}→/）/ Resolve a label node to the current language ({slash}→/) */
function getLabel(labelNode) {
    if (!labelNode) return "";
    var text = labelNode[lang] || labelNode.ja || labelNode.en || "";
    return text.replace(/\{slash\}/g, "/");
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with a trailing colon (full-width JA, half-width EN) */
function labelText(labelNode) {
    return getLabel(labelNode) + (lang === "ja" ? "：" : ":");
}

/* エラー表示補助 */
function showError(err) {
    var msg = getLabel(LABELS.message.processFailed) + "\n\n";
    if (err && err.message) {
        msg += err.message;
    } else {
        msg += String(err);
    }
    alert(msg);
}

/* 開発用ログ補助 */
function debugLog(context, err) {
    var msg = "[TextBreakSplitMerge] " + context;
    if (err) {
        if (err.message) {
            msg += " :: " + err.message;
        } else {
            msg += " :: " + String(err);
        }
    }
    try {
        $.writeln(msg);
    } catch (_) { }
}

/* 改行正規化ユーティリティ */
function normalizeParagraphBreaks(txt) {
    return String(txt || "").replace(/\r\n/g, "\r").replace(/\n/g, "\r");
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
    return String(txt || "").replace(/[\r\n]+$/g, "");
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

function isForcedBreak(codeOrChar) {
    var code = (typeof codeOrChar === "number") ? codeOrChar : getCharCodeSafe(codeOrChar);
    return code === 3 || code === 10;
}


function isAnyBreak(codeOrChar) {
    return isParagraphBreak(codeOrChar) || isForcedBreak(codeOrChar);
}


function isTabChar(codeOrChar) {
    var code = (typeof codeOrChar === "number") ? codeOrChar : getCharCodeSafe(codeOrChar);
    return code === 9;
}

function hasVisibleChars(txt) {
    return String(txt || "").replace(/[\r\n\t]/g, "").length > 0;
}

/* 英字ケース変換（純粋な文字列関数 / TextNormalize.jsx より移植）
   プレビュー（パレット側）と適用（メインエンジン側）の両方で使う */

/* 単語の先頭のみ大文字（各単語＝先頭大文字＋残りは小文字）
   事前に小文字化しておく必要はなく、すべて大文字の語も Negoticible のようになる */
function toWordCap(text) {
    return String(text).toLowerCase().replace(/\b([a-z])/g, function (m, c) { return c.toUpperCase(); });
}

/* 文頭のみ大文字（文単位＝先頭大文字＋残りは小文字）
   事前に小文字化しておく必要はなく、すべて大文字でも Negoticible のようになる。
   英単語が1つだけ（文区切りなし）の場合も、先頭の頭文字を大文字化する。*/
function toSentenceCase(text) {
    return String(text).toLowerCase().replace(/(^|[\.\!\?]\s+|[\r\n]+)([a-z])/g,
        function (m, prefix, c) { return prefix + c.toUpperCase(); });
}

/* 英語タイトル形式（冠詞・前置詞などは小文字 / John Gruber の Title Caps 移植）*/
function toTitleCase(text) {
    var small = "(a|abaft|aboard|about|above|absent|across|afore|after|against|along|alonside|amid|amidst|among|amongst|an|and|apopos|around|as|aside|astride|at|athwart|atop|barring|before|behind|below|beneath|beside|besides|between|betwixt|beyond|but|by|circa|concerning|despite|down|during|except|excluding|failing|following|for|from|given|in|including|inside|into|lest|like|mid|midst|minus|modula|near|next|nor|notwithstanding|of|off|on|onto|oppostie|or|out|outside|over|pace|per|plus|pro|qua|regarding|round|sans|save|than|that|the|through|throughout|till|times|to|toward|towards|under|underneath|unlike|until|unto|up|upon|versus|via|vice|with|within|without|worth|v[.]?|via|vs[.]?)";
    var punct = "([!\"#$%&'()*+,./:;<=>?@[\\\\\\]^_`{|}~-]*)";

    function lower(word) { return word.toLowerCase(); }
    function upper(word) { return word.substr(0, 1).toUpperCase() + word.substr(1); }

    function titleCaps(title) {
        var parts = [], split = /[:.;?!] |(?: |^)[\"Ò]/g, index = 0;
        while (true) {
            var m = split.exec(title);
            parts.push(
                title.substring(index, m ? m.index : title.length)
                    .replace(/\b([A-Za-z][a-z.'Õ]*)\b/g, function (all) {
                        return /[A-Za-z]\.[A-Za-z]/.test(all) ? all : upper(all);
                    })
                    .replace(RegExp("\\b" + small + "\\b", "ig"), lower)
                    .replace(RegExp("^" + punct + small + "\\b", "ig"), function (all, p, word) {
                        return p + upper(word);
                    })
                    .replace(RegExp("\\b" + small + punct + "$", "ig"), upper)
            );
            index = split.lastIndex;
            if (m) parts.push(m[0]);
            else break;
        }
        return parts.join("")
            .replace(/ V(s?)\. /ig, " v$1. ")
            .replace(/(['Õ])S\b/ig, "$1s")
            .replace(/\b(AT&T|Q&A)\b/ig, function (all) { return all.toUpperCase(); });
    }
    return titleCaps(String(text));
}

(function () {
    // ドキュメントが開かれていない場合は処理を終了
    if (app.documents.length === 0) {
        alert(getLabel(LABELS.message.noDocument));
        return;
    }

    // 選択オブジェクトを取得
    var selectedObjects = app.selection;

    // 選択オブジェクトがない場合は処理を終了
    try {
        if (!selectedObjects || (typeof selectedObjects.length === "number" && selectedObjects.length === 0)) {
            alert(getLabel(LABELS.message.noSelection));
            return;
        }
    } catch (e) {
        if (!selectedObjects) {
            alert(getLabel(LABELS.message.noSelection));
            return;
        }
    }

    // 初期選択からテキストフレームを解決
    selectedObjects = getTextFrames(selectedObjects);
    if (selectedObjects.length === 0) {
        alert(getLabel(LABELS.message.noTextFrames));
        return;
    }

    /* ユーティリティ関数 */

    /* テキストフレームのみ抽出 */
    function getTextFrames(objects) {
        var frames = [];

        function pushUnique(frame) {
            if (!frame) return;
            try {
                if (frame.isValid === false) return;
            } catch (e) { debugLog("getTextFrames: pushUnique/isValid", e); return; }

            for (var i = 0; i < frames.length; i++) {
                if (frames[i] === frame) return;
            }
            frames.push(frame);
        }

        function collect(item) {
            if (!item) return;

            try {
                if (item.isValid === false) return;
            } catch (e) { debugLog("getTextFrames: collect/isValid", e); return; }

            var typeName = "";
            try {
                typeName = item.typename || "";
            } catch (e2) { debugLog("getTextFrames: collect/typename", e2); return; }

            if (typeName === "TextFrame") {
                pushUnique(item);
                return;
            }

            /* Illustrator ではテキスト選択が TextRange として返ることがある */
            if (typeName === "TextRange") {
                try {
                    if (item.parent && item.parent.typename === "TextFrame") {
                        pushUnique(item.parent);
                        return;
                    }
                } catch (e3) { debugLog("getTextFrames: TextRange parent", e3); }
            }

            if (typeName === "GroupItem") {
                try {
                    for (var i = 0; i < item.pageItems.length; i++) {
                        collect(item.pageItems[i]);
                    }
                } catch (e4) { }
                return;
            }
        }

        if (!objects) return frames;

        try {
            if (typeof objects.length === "number" && !objects.typename) {
                for (var k = 0; k < objects.length; k++) {
                    collect(objects[k]);
                }
            } else {
                collect(objects);
            }
        } catch (e5) { debugLog("getTextFrames: iterate objects", e5); }

        return frames;
    }


    /* テキストタイプ判定 */
    function detectTextFrameType(objects) {
        var hasPointText = false;
        var hasAreaText = false;
        var frames = getTextFrames(objects);

        for (var i = 0; i < frames.length; i++) {
            if (frames[i].kind === TextType.AREATEXT) {
                hasAreaText = true;
            } else if (frames[i].kind === TextType.POINTTEXT) {
                hasPointText = true;
            }
        }

        if (hasPointText && hasAreaText) return "mixed";
        if (hasAreaText) return "area";
        return "point";
    }

    /* テキストタイプ件数を集計 */
    function countTextFrameTypes(objects) {
        var pointCount = 0;
        var areaCount = 0;
        var frames = getTextFrames(objects);

        for (var i = 0; i < frames.length; i++) {
            if (frames[i].kind === TextType.AREATEXT) {
                areaCount++;
            } else if (frames[i].kind === TextType.POINTTEXT) {
                pointCount++;
            }
        }

        return {
            total: frames.length,
            point: pointCount,
            area: areaCount
        };
    }

    /* 改行数とタブ数を集計 */
    function countBreakTypes(objects) {
        var paragraphBreakCount = 0;
        var forcedBreakCount = 0;
        var tabCount = 0;
        var frames = getTextFrames(objects);

        for (var i = 0; i < frames.length; i++) {
            var chars = frames[i].characters;
            for (var j = 0; j < chars.length; j++) {
                var code = chars[j].contents.charCodeAt(0);
                if (isParagraphBreak(code)) {
                    paragraphBreakCount++;
                } else if (isForcedBreak(code)) {
                    forcedBreakCount++;
                } else if (isTabChar(code)) {
                    tabCount++;
                }
            }
        }

        return {
            paragraph: paragraphBreakCount,
            forced: forcedBreakCount,
            tab: tabCount
        };
    }

    /* 各テキストフレームのcontentsを変換する共通処理 */
    function transformContents(objects, transformFunc) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            frames[i].contents = transformFunc(frames[i].contents);
        }
    }

    /* 強制改行（charCode 3 または 10）を削除する共通処理 */
    function removeForcedLineBreaks(frames) {
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (isForcedBreak(charCode)) {
                    frames[i].characters[c].remove();
                }
            }
        }
    }

    /* 上から順にソート（Y降順、同じYならX昇順） */
    function sortByPosition(items) {
        var arr = items.slice(0);
        arr.sort(function (a, b) {
            if (b.position[1] !== a.position[1]) return b.position[1] - a.position[1];
            return a.position[0] - b.position[0];
        });
        return arr;
    }

    /* Y座標で降順ソート */
    function sortByY(items) {
        var arr = items.slice(0);
        arr.sort(function (a, b) { return b.position[1] - a.position[1]; });
        return arr;
    }

    /* X座標で昇順ソート */
    function sortByX(items) {
        var arr = items.slice(0);
        arr.sort(function (a, b) { return a.position[0] - b.position[0]; });
        return arr;
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

    /* 選択範囲のバウンディングボックス取得 */
    function getSelBounds(selection) {
        var x1 = selection[0].visibleBounds[0];
        var y1 = selection[0].visibleBounds[1];
        var x2 = selection[0].visibleBounds[2];
        var y2 = selection[0].visibleBounds[3];
        for (var i = 1; i < selection.length; i++) {
            var b = selection[i].visibleBounds;
            if (b[0] < x1) x1 = b[0];
            if (b[1] > y1) y1 = b[1];
            if (b[2] > x2) x2 = b[2];
            if (b[3] < y2) y2 = b[3];
        }
        return [x1, y1, x2, y2];
    }

    /* テキストフレームを1つのグループにまとめる */
    function groupTextFrames(frames, targetLayer) {
        var validFrames = getTextFrames(frames);
        if (validFrames.length === 0) return [];

        var layer = targetLayer || app.activeDocument.activeLayer;
        var grp = layer.groupItems.add();

        for (var i = 0; i < validFrames.length; i++) {
            try {
                validFrames[i].move(grp, ElementPlacement.PLACEATEND);
            } catch (e) { debugLog("groupTextFrames: move to group", e); }
        }

        return [grp];
    }

    /* 改行系の関数 */

    /* 改行文字を削除する関数 */
    function removeLineBreaks(objects) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (isParagraphBreak(charCode)) {
                    frames[i].characters[c].remove();
                }
            }
        }
    }

    /* 強制改行と改行を削除する関数 */
    function removeAllBreaks(objects) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (isAnyBreak(charCode)) {
                    frames[i].characters[c].remove();
                }
            }
        }
    }

    /* 複数フレームを連結し、改行を取り除いて1行に統合する関数 */
    function flattenToOneLine(objects) {
        var frames = getTextFrames(objects);
        if (frames.length < 2) {
            removeEmptyLines(frames);
            removeAllBreaks(frames);
            return frames;
        }
        var result = concatVertical(objects);
        var targets = result && result.length ? result : frames;
        removeEmptyLines(targets);
        removeAllBreaks(targets);
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
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                if (isTabChar(frames[i].characters[c].contents)) {
                    frames[i].characters[c].remove();
                }
            }
        }
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

    /* 全角英数字を半角に変換する関数 */
    function fullToHalfAlnum(objects) {
        transformContents(objects, function (txt) {
            var result = "";
            for (var i = 0; i < txt.length; i++) {
                var code = txt.charCodeAt(i);
                // 全角数字 ０-９ (0xFF10-0xFF19) → 半角 0-9
                // 全角大文字 Ａ-Ｚ (0xFF21-0xFF3A) → 半角 A-Z
                // 全角小文字 ａ-ｚ (0xFF41-0xFF5A) → 半角 a-z
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
        });
    }

    /* 半角カナを全角カナに変換する関数 */
    function halfToFullKana(objects) {
        var halfKana = "\uFF66\uFF67\uFF68\uFF69\uFF6A\uFF6B\uFF6C\uFF6D\uFF6E\uFF6F\uFF70\uFF71\uFF72\uFF73\uFF74\uFF75\uFF76\uFF77\uFF78\uFF79\uFF7A\uFF7B\uFF7C\uFF7D\uFF7E\uFF7F\uFF80\uFF81\uFF82\uFF83\uFF84\uFF85\uFF86\uFF87\uFF88\uFF89\uFF8A\uFF8B\uFF8C\uFF8D\uFF8E\uFF8F\uFF90\uFF91\uFF92\uFF93\uFF94\uFF95\uFF96\uFF97\uFF98\uFF99\uFF9A\uFF9B\uFF9C\uFF9D";
        var fullKana = "\u30F2\u30A1\u30A3\u30A5\u30A7\u30A9\u30E3\u30E5\u30E7\u30C3\u30FC\u30A2\u30A4\u30A6\u30A8\u30AA\u30AB\u30AD\u30AF\u30B1\u30B3\u30B5\u30B7\u30B9\u30BB\u30BD\u30BF\u30C1\u30C4\u30C6\u30C8\u30CA\u30CB\u30CC\u30CD\u30CE\u30CF\u30D2\u30D5\u30D8\u30DB\u30DE\u30DF\u30E0\u30E1\u30E2\u30E4\u30E6\u30E8\u30E9\u30EA\u30EB\u30EC\u30ED\u30EF\u30F3";
        // 濁点・半濁点の対応
        var dakutenBase = "\uFF76\uFF77\uFF78\uFF79\uFF7A\uFF7B\uFF7C\uFF7D\uFF7E\uFF7F\uFF80\uFF81\uFF82\uFF83\uFF84\uFF8A\uFF8B\uFF8C\uFF8D\uFF8E";
        var dakutenFull = "\u30AC\u30AE\u30B0\u30B2\u30B4\u30B6\u30B8\u30BA\u30BC\u30BE\u30C0\u30C2\u30C5\u30C7\u30C9\u30D0\u30D3\u30D6\u30D9\u30DC";
        var handakutenBase = "\uFF8A\uFF8B\uFF8C\uFF8D\uFF8E";
        var handakutenFull = "\u30D1\u30D4\u30D7\u30DA\u30DD";
        transformContents(objects, function (txt) {
            var result = "";
            for (var i = 0; i < txt.length; i++) {
                var ch = txt.charAt(i);
                var next = (i + 1 < txt.length) ? txt.charAt(i + 1) : "";
                // 濁点結合 (ﾞ = \uFF9E)
                if (next === "\uFF9E") {
                    var di = dakutenBase.indexOf(ch);
                    if (di >= 0) {
                        result += dakutenFull.charAt(di);
                        i++;
                        continue;
                    }
                }
                // 半濁点結合 (ﾟ = \uFF9F)
                if (next === "\uFF9F") {
                    var hi = handakutenBase.indexOf(ch);
                    if (hi >= 0) {
                        result += handakutenFull.charAt(hi);
                        i++;
                        continue;
                    }
                }
                // 単独変換
                var ki = halfKana.indexOf(ch);
                if (ki >= 0) {
                    result += fullKana.charAt(ki);
                } else if (ch === "\uFF9E") {
                    result += "\u309B"; // 濁点単独
                } else if (ch === "\uFF9F") {
                    result += "\u309C"; // 半濁点単独
                } else if (ch === "\uFF61") {
                    result += "\u3002"; // 。
                } else if (ch === "\uFF62") {
                    result += "\u300C"; // 「
                } else if (ch === "\uFF63") {
                    result += "\u300D"; // 」
                } else if (ch === "\uFF64") {
                    result += "\u3001"; // 、
                } else if (ch === "\uFF65") {
                    result += "\u30FB"; // ・
                } else {
                    result += ch;
                }
            }
            return result;
        });
    }

    /* 箇条書き記号の除去 */
    function toggleBulletList(objects) {
        var bulletPat = /^[\u30FB\u2022\-\*]\s*/; // ・ • - *
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            for (var i = 0; i < lines.length; i++) {
                lines[i] = lines[i].replace(bulletPat, "");
            }
            return lines.join("\r");
        });
    }

    /* 番号リスト記号の除去 */
    function toggleNumberList(objects) {
        // 全角半角数字 + 全角半角ピリオド + スペース(複数可)
        var numPat = /^[0-9\uFF10-\uFF19]+[.\uFF0E]\s*/;
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            for (var i = 0; i < lines.length; i++) {
                lines[i] = lines[i].replace(numPat, "");
            }
            return lines.join("\r");
        });
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
                if (!seen[lines[i]]) {
                    seen[lines[i]] = true;
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
            var newTxt = "";
            for (var j = 0; j < txt.length; j++) {
                var c = txt.charAt(j);
                newTxt += c;
                if (!isAnyBreak(c) && j < txt.length - 1) {
                    var nextC = txt.charAt(j + 1);
                    if (!isAnyBreak(nextC)) {
                        newTxt += '\r';
                    }
                }
            }
            return newTxt;
        });
    }

    /* 指定文字数ごとに改行を挿入する関数 */
    function addLineBreakAtCount(objects, count, useForcedBreak) {
        var n = parseInt(count, 10);
        if (!n || n <= 0) n = 35;
        var br = useForcedBreak ? String.fromCharCode(3) : "\r";
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            var result = [];
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i];
                while (line.length > n) {
                    result.push(line.substring(0, n));
                    line = line.substring(n);
                }
                result.push(line);
            }
            return result.join(br);
        });
    }

    /* 強制改行を通常の改行に変換する関数 */
    function convertForcedLineBreaks(objects) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (isForcedBreak(charCode)) {
                    frames[i].characters[c].contents = String.fromCharCode(13);
                }
            }
        }
    }

    /* 改行を強制改行に変換する関数 */
    function convertToForcedBreaks(objects) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (isParagraphBreak(charCode)) {
                    frames[i].characters[c].contents = String.fromCharCode(3);
                }
            }
        }
    }

    /* 句読点の後に改行を挿入する関数 */
    function addLineBreakAtPunctuation(objects, punctuationChars) {
        /* 対象記号：和文・欧文の句読点と終止記号 */
        var punctuation = punctuationChars || "、。，．｡､,.!?！？";
        transformContents(objects, function (txt) {
            var newTxt = "";
            for (var j = 0; j < txt.length; j++) {
                var c = txt.charAt(j);
                newTxt += c;
                if (punctuation.indexOf(c) !== -1 && j < txt.length - 1) {
                    var remaining = txt.substring(j + 1);
                    if (hasVisibleChars(remaining)) {
                        var nextC = txt.charAt(j + 1);
                        if (!isAnyBreak(nextC)) {
                            newTxt += '\r';
                        }
                    }
                }
            }
            return newTxt;
        });
    }

    /* タブで分解する関数 */
    function splitByTab(objects) {
        var doc = app.activeDocument;
        var resultFrames = [];

        for (var i = 0; i < objects.length; i++) {
            var obj = objects[i];

            if (obj.typename === "TextFrame") {
                var paraCount = obj.paragraphs.length;
                var pos = obj.position;
                var myX = pos[0];
                var myY = pos[1];

                for (var p = 0; p < paraCount; p++) {
                    var para = obj.paragraphs[p];
                    var mySize = para.size;
                    var plusH = para.leading;
                    if (plusH === 0) { plusH = mySize * 1.2; }

                    /* 段落単位でタブ位置を取得 */
                    var tabCharPositions = [];
                    var paraChars = para.characters;
                    for (var ci = 0; ci < paraChars.length; ci++) {
                        if (isTabChar(paraChars[ci].contents)) {
                            if (ci + 1 < paraChars.length) {
                                var nextCharBounds = paraChars[ci + 1].visibleBounds;
                                tabCharPositions.push(nextCharBounds[0]); /* 左端X座標 */
                            }
                        }
                    }

                    var paraText = stripTrailingBreaks(para.contents);
                    var contAry = paraText.split("\t");
                    var prevFrame = null;
                    for (var c = 0; c < contAry.length; c++) {
                        var dupFrame = obj.duplicate(doc.activeLayer);
                        dupFrame.contents = contAry[c];
                        var moveX = myX;
                        if (c > 0) {
                            if ((c - 1) < tabCharPositions.length) {
                                moveX = tabCharPositions[c - 1];
                            } else if (prevFrame) {
                                var prevBounds = prevFrame.visibleBounds;
                                moveX = prevBounds[2] + (mySize * 0.5);
                            } else {
                                moveX = myX + (mySize * c);
                            }
                        }
                        dupFrame.position = [moveX, myY];
                        prevFrame = dupFrame;
                        resultFrames.push(dupFrame);
                    }
                    myY = myY - plusH;
                }

                obj.remove();
            }
        }

        return groupTextFrames(resultFrames, doc.activeLayer);
    }

    /* 改行で分割する関数 */
    function splitByLineBreak(objects) {
        var doc = app.activeDocument;
        var resultFrames = [];

        for (var i = 0; i < objects.length; i++) {
            var obj = objects[i];

            if (obj.typename === "TextFrame") {
                var paraCount = obj.paragraphs.length;
                var pos = obj.position;
                var myX = pos[0];
                var myY = pos[1];

                for (var p = 0; p < paraCount; p++) {
                    var para = obj.paragraphs[p];
                    var mySize = para.size;
                    var plusH = para.leading;
                    if (plusH === 0) { plusH = mySize * 1.2; }

                    var paraText = stripTrailingBreaks(para.contents);
                    if (paraText !== "") {
                        var dupFrame = obj.duplicate(doc.activeLayer);
                        dupFrame.contents = paraText;
                        dupFrame.position = [myX, myY];
                        resultFrames.push(dupFrame);
                    }
                    myY = myY - plusH;
                }

                obj.remove();
            }
        }

        return groupTextFrames(resultFrames, doc.activeLayer);
    }

    /* =========================================
     * 改行で分割（書式保持）
     * TextRange.duplicate で書式を維持し、geometricBounds の tail 座標で正確に配置
     * 参考: Split Rows for Ai.jsx の splitRowsPoint 方式
     * ========================================= */
    function splitByLineBreakKeepStyle(objects) {
        var resultFrames = [];
        var frames = getTextFrames(objects);

        for (var fi = 0; fi < frames.length; fi++) {
            var srcFrame = frames[fi];

            /* 空フレームはスキップ */
            if (/^\s*$/.test(srcFrame.contents)) {
                srcFrame.remove();
                continue;
            }

            var isHorizontal = (srcFrame.orientation === TextOrientation.HORIZONTAL);

            /*
             * direction: translate に渡す配列 [dx, dy] のどちらを操作するか
             * indexTail: geometricBounds [left, top, right, bottom] のうち末尾側の index
             *   横組み → Y方向に伸びる → tail = bottom(3)
             *   縦組み → X方向(左方向)に伸びる → tail = left(0)
             */
            var direction, indexTail;
            if (isHorizontal) {
                direction = 1; /* y */
                indexTail = 3; /* bottom */
            } else {
                direction = 0; /* x */
                indexTail = 0; /* left */
            }

            /* 末尾の改行文字を除去するヘルパー */
            function trimTrailingBreaks(frame) {
                var removed = 0;
                for (var ci = frame.characters.length - 1; ci >= 0; ci--) {
                    if (/[\r\n\x03]/.test(frame.characters[ci].contents)) {
                        frame.characters[ci].remove();
                        removed++;
                    } else {
                        break;
                    }
                }
                /* フレームが空になったら削除 */
                if (/^\s*$/.test(frame.contents)) {
                    frame.remove();
                    return Infinity;
                }
                return removed;
            }

            /* 末尾改行を事前に除去 */
            trimTrailingBreaks(srcFrame);
            if (/^\s*$/.test(srcFrame.contents)) {
                try { srcFrame.remove(); } catch (_) { }
                continue;
            }

            var rows = srcFrame.paragraphs;
            if (rows.length <= 1) {
                resultFrames.push(srcFrame);
                continue;
            }

            var splitResults = [];

            /* 後ろの段落から順に処理（参考スクリプトと同じ方式） */
            for (var i = rows.length - 1; i >= 1; i--) {
                var currentRow = rows[i];
                var currentText = currentRow.contents;

                /* 空行またはホワイトスペースのみの行は消して次へ */
                if (/^\s*$/.test(currentText)) {
                    currentRow.remove();
                    continue;
                }

                /*
                 * 分割前に元フレームの tail 座標を記録
                 * この値が「この段落より下（左）にあるべき位置」の基準になる
                 */
                var oldTail = srcFrame.geometricBounds[indexTail];

                /* 新フレームを生成し、TextRange.duplicate で書式ごとコピー */
                var newFrame = srcFrame.duplicate(srcFrame, ElementPlacement.PLACEAFTER);
                newFrame.contents = "";
                currentRow.duplicate(newFrame, ElementPlacement.INSIDE);

                /* 新フレームの末尾改行を除去 */
                trimTrailingBreaks(newFrame);

                /*
                 * 位置補正: 新フレームの tail が、元フレームの旧 tail に揃うように移動
                 * これにより各段落が元の位置に正確に配置される
                 */
                var newTail = newFrame.geometricBounds[indexTail];
                var delta = [0, 0];
                delta[direction] = oldTail - newTail;
                newFrame.translate(delta[0], delta[1]);

                splitResults.unshift(newFrame);

                /* 元フレームからこの段落を削除 */
                currentRow.remove();

                /* 元フレームの末尾改行を除去し、除去した行数分 index を飛ばす */
                i -= trimTrailingBreaks(srcFrame);
            }

            /* 元フレーム（先頭段落が残っている）を先頭に移動 */
            try {
                if (srcFrame && !/^\s*$/.test(srcFrame.contents)) {
                    srcFrame.move(splitResults[0], ElementPlacement.PLACEBEFORE);
                    splitResults.unshift(srcFrame);
                }
            } catch (_) { }

            for (var r = 0; r < splitResults.length; r++) {
                resultFrames.push(splitResults[r]);
            }
        }

        return groupTextFrames(resultFrames, app.activeDocument.activeLayer);
    }

    /* =========================================
     * 1文字ごとにテキストフレームを分割（書式保持）
     * ========================================= */
    function splitByCharKeepStyle(objects) {
        var frames = getTextFrames(objects);
        var resultFrames = [];
        for (var i = 0; i < frames.length; i++) {
            var made = splitCharHighPrecision(frames[i], true);
            for (var j = 0; j < made.length; j++) resultFrames.push(made[j]);
        }
        return groupTextFrames(resultFrames, app.activeDocument.activeLayer);
    }

    /* =========================================
     * 1文字ごとにテキストフレームを分割（書式無視）
     * =========================================  */
    function splitByCharIgnoreStyle(objects) {
        var frames = getTextFrames(objects);
        var resultFrames = [];
        for (var i = 0; i < frames.length; i++) {
            stripStyleKeepFirstFont(frames[i]);
            var made = splitCharHighPrecision(frames[i], false);
            for (var j = 0; j < made.length; j++) resultFrames.push(made[j]);
        }
        return groupTextFrames(resultFrames, app.activeDocument.activeLayer);
    }

    /* try 付き代入（属性設定が失敗しても続行）*/
    function safeSet(obj, prop, val) { try { obj[prop] = val; } catch (e) { } }

    /* src の属性を dst へコピー（読み書きとも try で保護）*/
    function copyAttr(dst, src, prop) { try { dst[prop] = src[prop]; } catch (e) { } }

    /* characterAttributes を初期化（フォント・サイズは保持、色は黒・各種は既定へ）*/
    function applyResetStyle(ca, keepFont, keepSize, black) {
        if (keepFont) safeSet(ca, "textFont", keepFont);
        if (keepSize != null) safeSet(ca, "size", keepSize);
        safeSet(ca, "fillColor", black);
        safeSet(ca, "baselineShift", 0);
        safeSet(ca, "horizontalScale", 100);
        safeSet(ca, "verticalScale", 100);
        safeSet(ca, "rotation", 0);
        safeSet(ca, "tracking", 0);
        safeSet(ca, "kerningMethod", KerningMethod.METRICS);
        safeSet(ca, "autoLeading", true);
    }

    /* 書式リセット（先頭文字のフォント情報のみ保持） */
    function stripStyleKeepFirstFont(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;
        var tr, chars;
        try { tr = textFrame.textRange; } catch (_) { return; }
        try { chars = tr.characters; } catch (_) { return; }
        if (!chars || chars.length < 1) return;

        var firstCA;
        try { firstCA = chars[0].characterAttributes; } catch (_) { return; }
        var keepFont = null, keepSize = null;
        try { keepFont = firstCA.textFont; } catch (_) { }
        try { keepSize = firstCA.size; } catch (_) { }

        var black = new GrayColor(); black.gray = 100;

        var ca;
        try { ca = tr.characterAttributes; } catch (_) { return; }
        applyResetStyle(ca, keepFont, keepSize, black);

        for (var i = 0; i < chars.length; i++) {
            var c2;
            try { c2 = chars[i].characterAttributes; } catch (_) { continue; }
            applyResetStyle(c2, keepFont, keepSize, black);
        }
    }

    /* 高精度分割（アウトラインのバウンディングボックスを使用） */
    function splitCharHighPrecision(textFrame, keepStyle) {
        if (!textFrame || textFrame.typename !== "TextFrame") return [];

        var tr, chars, n;
        try { tr = textFrame.textRange; } catch (_) { return []; }
        try { chars = tr.characters; } catch (_) { return []; }
        try { n = chars.length; } catch (_) { return []; }
        if (!n || n <= 0) return [];

        var outlineInfo = null;
        try { outlineInfo = buildOutlineCharBounds(textFrame); } catch (_) { outlineInfo = null; }

        if (outlineInfo && outlineInfo.ok && outlineInfo.boundsList && outlineInfo.boundsList.length > 0) {
            var boundsList = outlineInfo.boundsList;
            var layer = textFrame.layer;
            var made = [];
            var bi = 0;

            for (var ci = 0; ci < n; ci++) {
                var ch;
                try { ch = chars[ci]; } catch (_) { continue; }
                var content = "";
                try { content = ch.contents; } catch (_) { continue; }
                if (content === "") continue;

                var code = content.charCodeAt(0);
                if (code === 13 || code === 10 || code === 9 || content === " " || content === "\u3000") continue;

                if (bi >= boundsList.length) {
                    for (var x = 0; x < made.length; x++) { try { made[x].remove(); } catch (_) { } }
                    try { if (outlineInfo.outlinedRoot) outlineInfo.outlinedRoot.remove(); } catch (_) { }
                    return splitCharFallback(textFrame, keepStyle);
                }

                var nf;
                try {
                    nf = layer.textFrames.add();
                    nf.contents = content;
                } catch (_) {
                    for (var x2 = 0; x2 < made.length; x2++) { try { made[x2].remove(); } catch (_) { } }
                    try { if (outlineInfo.outlinedRoot) outlineInfo.outlinedRoot.remove(); } catch (_) { }
                    return splitCharFallback(textFrame, keepStyle);
                }

                if (keepStyle) {
                    try { copyCharAttrs(nf, ch); } catch (e) { debugLog("splitCharHighPrecision: copyCharAttrs", e); }
                }
                try { nf.matrix = textFrame.matrix; } catch (e) { debugLog("splitCharHighPrecision: matrix", e); }
                try { nf.left = textFrame.left; nf.top = textFrame.top; } catch (e) { debugLog("splitCharHighPrecision: initial position", e); }

                try {
                    moveFrameToMatchBounds(nf, boundsList[bi]);
                } catch (_) {
                    try { nf.remove(); } catch (_) { }
                    for (var x3 = 0; x3 < made.length; x3++) { try { made[x3].remove(); } catch (_) { } }
                    try { if (outlineInfo.outlinedRoot) outlineInfo.outlinedRoot.remove(); } catch (_) { }
                    return splitCharFallback(textFrame, keepStyle);
                }

                made.push(nf);
                bi++;
            }

            try { if (outlineInfo.outlinedRoot) outlineInfo.outlinedRoot.remove(); } catch (_) { }
            try { textFrame.remove(); } catch (_) { }
            return made;
        }

        return splitCharFallback(textFrame, keepStyle);
    }

    /* アウトラインのバウンディングボックス情報を構築 */
    function buildOutlineCharBounds(textFrame) {
        var dup = null;
        try { dup = textFrame.duplicate(textFrame.parent, ElementPlacement.PLACEATBEGINNING); } catch (_) {
            try { dup = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATBEGINNING); } catch (_) { return { ok: false }; }
        }

        var outlined = null;
        try { outlined = dup.createOutline(); } catch (_) {
            try { dup.remove(); } catch (_) { }
            return { ok: false };
        }
        try { dup.remove(); } catch (_) { }
        if (!outlined) return { ok: false };

        var items = [];
        try {
            var pi = outlined.pageItems;
            var direct = [];
            for (var i = 0; i < pi.length; i++) direct.push(pi[i]);
            if (direct.length === 1 && direct[0] && direct[0].typename === "GroupItem") {
                var inner = direct[0].pageItems;
                for (var j = 0; j < inner.length; j++) items.push(inner[j]);
            } else {
                for (var k = 0; k < direct.length; k++) items.push(direct[k]);
            }
        } catch (_) { }

        try { sortOutlineItems(items, textFrame); } catch (_) { }

        if (items.length === 0) {
            try { outlined.remove(); } catch (_) { }
            return { ok: false };
        }

        var boundsList = [];
        for (var m = 0; m < items.length; m++) {
            try { boundsList.push(items[m].geometricBounds); } catch (_) { }
        }

        return { ok: boundsList.length > 0, outlinedRoot: outlined, boundsList: boundsList };
    }

    /* アウトライン項目を読み順にソート */
    function sortOutlineItems(items) {
        if (!items || items.length <= 1) return;

        var arr = [];
        for (var i = 0; i < items.length; i++) {
            var bnd;
            try { bnd = items[i].geometricBounds; } catch (_) { continue; }
            if (!bnd || bnd.length !== 4) continue;
            var L = bnd[0], T = bnd[1], R = bnd[2], B = bnd[3];
            arr.push({ it: items[i], L: L, T: T, cx: (L + R) / 2, cy: (T + B) / 2, h: Math.abs(T - B), idx: i });
        }
        if (arr.length <= 1) return;

        var th = estimateCharRowThreshold(arr);

        arr.sort(function (p, q) {
            var dy = q.cy - p.cy;
            if (Math.abs(dy) > 0.001) return (dy < 0) ? -1 : 1;
            var dL = p.L - q.L;
            if (Math.abs(dL) > 0.001) return (dL < 0) ? -1 : 1;
            return p.idx - q.idx;
        });

        var rows = [];
        for (var a = 0; a < arr.length; a++) {
            var cur = arr[a];
            var placed = false;
            for (var r = 0; r < rows.length; r++) {
                if (Math.abs(cur.cy - rows[r].cy) <= th) {
                    rows[r].items.push(cur);
                    rows[r].cy = (rows[r].cy * (rows[r].items.length - 1) + cur.cy) / rows[r].items.length;
                    placed = true;
                    break;
                }
            }
            if (!placed) rows.push({ cy: cur.cy, items: [cur] });
        }

        rows.sort(function (p, q) { return (q.cy - p.cy < 0) ? -1 : (q.cy - p.cy > 0) ? 1 : 0; });

        var out = [];
        for (var ri = 0; ri < rows.length; ri++) {
            rows[ri].items.sort(function (p, q) {
                var dL2 = p.L - q.L;
                if (Math.abs(dL2) > 0.5) return (dL2 < 0) ? -1 : 1;
                return p.idx - q.idx;
            });
            for (var j = 0; j < rows[ri].items.length; j++) out.push(rows[ri].items[j].it);
        }
        for (var k = 0; k < out.length; k++) items[k] = out[k];
    }

    function estimateCharRowThreshold(arr) {
        var hs = [];
        for (var i = 0; i < arr.length; i++) {
            if (arr[i].h > 0) hs.push(arr[i].h);
        }
        if (hs.length === 0) return 8;
        hs.sort(function (a, b) { return a - b; });
        var mid = hs[Math.floor(hs.length / 2)];
        if (!mid || mid <= 0) mid = hs[0];
        var th = mid * 0.6;
        return th < 2 ? 2 : th;
    }

    /* 文字属性をコピー */
    function copyCharAttrs(dstFrame, srcChar) {
        var src;
        try { src = srcChar.characterAttributes; } catch (_) { return; }
        var dst;
        try { dst = dstFrame.textRange.characterAttributes; } catch (_) { return; }
        copyAttr(dst, src, "textFont");
        copyAttr(dst, src, "size");
        copyAttr(dst, src, "horizontalScale");
        copyAttr(dst, src, "verticalScale");
        copyAttr(dst, src, "tracking");
        copyAttr(dst, src, "baselineShift");
        copyAttr(dst, src, "rotation");
        try { if (src.fillColor && src.fillColor.typename !== "NoColor") dst.fillColor = src.fillColor; } catch (e) { }
        try { if (src.strokeColor && src.strokeColor.typename !== "NoColor") { dst.strokeColor = src.strokeColor; dst.strokeWeight = src.strokeWeight; } } catch (e) { }
        copyAttr(dst, src, "autoLeading");
        try { if (!src.autoLeading) dst.leading = src.leading; } catch (e) { }
        copyAttr(dst, src, "kerningMethod");
    }

    /* アウトラインのバウンディングボックスに合わせて移動 */
    function moveFrameToMatchBounds(nf, targetBounds) {
        if (!nf || !targetBounds || targetBounds.length !== 4) return;

        var dup;
        try { dup = nf.duplicate(nf.parent, ElementPlacement.PLACEATBEGINNING); } catch (_) {
            try { dup = nf.duplicate(nf.layer, ElementPlacement.PLACEATBEGINNING); } catch (_) { return; }
        }

        var outlined;
        try { outlined = dup.createOutline(); } catch (_) {
            try { dup.remove(); } catch (_) { }
            return;
        }
        try { dup.remove(); } catch (_) { }
        if (!outlined) return;

        var bNow;
        try { bNow = outlined.geometricBounds; } catch (_) { bNow = null; }
        try { outlined.remove(); } catch (_) { }
        if (!bNow || bNow.length !== 4) return;

        var dx = (targetBounds[0] + targetBounds[2]) / 2 - (bNow[0] + bNow[2]) / 2;
        var dy = (targetBounds[1] + targetBounds[3]) / 2 - (bNow[1] + bNow[3]) / 2;

        try { nf.left += dx; nf.top += dy; } catch (_) {
            try { nf.translate(dx, dy); } catch (_) { }
        }
    }

    /* フォールバック処理（文字幅の積算で位置を再構成） */
    function splitCharFallback(textFrame, keepStyle) {
        if (!textFrame || textFrame.typename !== "TextFrame") return [];

        var textLength;
        try { textLength = textFrame.textRange.characters.length; } catch (_) { return []; }
        if (!textLength) return [];

        var layer;
        try { layer = textFrame.layer; } catch (_) { return []; }

        var made = [];
        for (var i = textLength - 1; i >= 0; i--) {
            try {
                var character = textFrame.textRange.characters[i];
                var charContent = character.contents;
                var code = charContent.charCodeAt(0);
                if (isAnyBreak(code) || isTabChar(code) || charContent === " " || charContent === "\u3000") continue;

                var newFrame = layer.textFrames.add();
                newFrame.contents = charContent;

                if (keepStyle) {
                    copyCharAttrs(newFrame, character);
                }

                newFrame.top = textFrame.top;
                newFrame.left = textFrame.left;
                try { newFrame.matrix = textFrame.matrix; } catch (e) { debugLog("splitCharFallback: matrix", e); }

                var offsetX = 0;
                for (var j = 0; j < i; j++) {
                    var prevChar = textFrame.textRange.characters[j];
                    var pc = prevChar.contents;
                    var pcode = pc.charCodeAt(0);
                    if (isAnyBreak(pcode) || isTabChar(pcode) || pc === " " || pc === "\u3000") continue;

                    var tempFrame = layer.textFrames.add();
                    tempFrame.contents = pc;
                    try {
                        var tca = tempFrame.textRange.characterAttributes;
                        var sca = prevChar.characterAttributes;
                        tca.textFont = sca.textFont;
                        tca.size = sca.size;
                        tca.horizontalScale = sca.horizontalScale;
                        tca.verticalScale = sca.verticalScale;
                        tca.tracking = sca.tracking;
                    } catch (e) { debugLog("splitCharFallback: tempFrame attrs", e); }
                    offsetX += tempFrame.width;
                    tempFrame.remove();
                }

                var angle = 0;
                try {
                    var mx = textFrame.matrix;
                    angle = Math.atan2(mx.mValueB, mx.mValueA) * 180 / Math.PI;
                } catch (e) { debugLog("splitCharFallback: matrix", e); }
                var rad = angle * Math.PI / 180;
                newFrame.left = textFrame.left + offsetX * Math.cos(rad);
                newFrame.top = textFrame.top - offsetX * Math.sin(rad);

                made.push(newFrame);
            } catch (e) { debugLog("splitCharFallback: main loop", e); }
        }

        try { textFrame.remove(); } catch (_) { }
        return made;
    }

    /*
     * 縦連結
     *
     * 仕様：
     * - 選択されたテキストフレームを上から下、同じ高さでは左から右の順に処理します。
     * - 各フレーム内の複数行は、改行単位で分解し、元フレームの行送りを使って仮のY位置を再構成します。
     * - 連結順は仮配置後に再ソートして決定します。
     * - 出力は最上段要素をベースにし、各行を `\r` で連結した単一テキストへ再構成します。
     * - 複雑な段落属性や行ごとの差分書式は厳密には保持しません。
     */

    /* 横連結（行維持）
     * 同じ行のテキストを左から右へ連結し、行ごとに別テキストフレームとして残します。 */


    function concatHorizontalOnly(objects) {
        var LINE_Y_THRESHOLD = 5;
        var textItems = getTextFrames(objects);
        if (textItems.length < 2) return textItems;

        var textLines = groupByLineY(sortByY(textItems), LINE_Y_THRESHOLD);
        var resultFrames = [];

        for (var i = 0; i < textLines.length; i++) {
            var row = sortByX(textLines[i]);
            if (row.length === 1) {
                resultFrames.push(row[0]);
                continue;
            }

            var mergedText = "";
            for (var j = 0; j < row.length; j++) {
                mergedText += stripTrailingBreaks(row[j].contents);
            }

            row[0].contents = mergedText;
            resultFrames.push(row[0]);

            for (var k = 1; k < row.length; k++) {
                try { row[k].remove(); } catch (_) { }
            }
        }

        return groupTextFrames(resultFrames, app.activeDocument.activeLayer);
    }

    function concatVertical(objects) {
        var textFrames = getTextFrames(objects);
        if (textFrames.length < 2) return textFrames;

        textFrames = sortByPosition(textFrames);

        /* 各テキストフレームを行単位に分解して再ソート */
        var splitFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            var sourceFrame = textFrames[i];
            var lines = splitParagraphLines(sourceFrame.contents);
            var basePos = sourceFrame.position;
            var baseX = basePos[0];
            var baseY = basePos[1];
            var baseSize = sourceFrame.textRange.characterAttributes.size;
            var baseLeading = sourceFrame.textRange.characterAttributes.leading;
            if (!baseLeading || baseLeading === 0) {
                baseLeading = baseSize * 1.2;
            }

            for (var j = 0; j < lines.length; j++) {
                var lineText = stripTrailingBreaks(lines[j]);
                if (lineText !== "") {
                    var tf = sourceFrame.duplicate();
                    tf.contents = lineText;
                    tf.position = [baseX, baseY - (j * baseLeading)];
                    splitFrames.push(tf);
                }
            }
        }

        if (splitFrames.length === 0) return [];

        splitFrames = sortByPosition(splitFrames);

        /* 最上段のフレームをベースに文字列で再構成 */
        var mergedLines = [];
        for (var k = 0; k < splitFrames.length; k++) {
            var mergedLineText = stripTrailingBreaks(splitFrames[k].contents);
            if (mergedLineText !== "") {
                mergedLines.push(mergedLineText);
            }
        }

        if (mergedLines.length === 0) {
            for (var m = 0; m < splitFrames.length; m++) {
                splitFrames[m].remove();
            }
            return [];
        }

        var baseFrame = splitFrames[0];
        baseFrame.contents = mergedLines.join('\r');
        baseFrame.position = splitFrames[0].position;

        for (var n = 1; n < splitFrames.length; n++) {
            splitFrames[n].remove();
        }

        for (var t = 0; t < textFrames.length; t++) {
            try {
                if (textFrames[t] !== baseFrame) {
                    textFrames[t].remove();
                }
            } catch (removeErr) { }
        }
        return [baseFrame];
    }

    /*
 * 横連結（行統合）
 *
 * 仕様：
 * - 選択されたテキストフレームをY位置ベースで行グループ化し、各行を左から右へ連結します。
 * - 処理前に、強制改行（charCode 3 / 10）は削除します。
 * - 1行のみの場合は、その行の全テキストを連結して1つのテキストオブジェクトとして出力します。
 * - 複数行の場合は、まず各行を個別に連結し、その後、行間の連結ルールに従って1つのテキストへまとめます。
 * - 行末が英単語のハイフン区切りで、次行頭も英数字なら、ハイフンを削除して連結します。
 * - 行末と次行頭がともに英数字系なら、語間スペースを補います。
 * - 行末が句点・終止記号の場合のみ、行間に改行を残します。
 * - `textMode` が `point` の場合はポイント文字で出力し、それ以外はエリア内文字で出力します。
 * - `textMode` が未指定または `mixed` の場合は既存選択から自動判定し、混在時はエリア内文字を優先します。
 * - 出力書式は先頭要素のフォント・サイズを基準に再設定します。
 * - 行送りは、連結後の各行のY差分をもとに再設定します。
 * - これは見た目ベースの近似連結であり、複雑な書式差・回転・厳密な段落属性までは保持しません。
 */

    /* 行内のフレームをX順に連結した文字列と、その並びを返す */
    function concatRowText(rowFrames) {
        var sorted = sortByX(rowFrames);
        var text = "";
        for (var j = 0; j < sorted.length; j++) text += stripTrailingBreaks(sorted[j].contents);
        return { sorted: sorted, text: text };
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

    /* 1行のみの横連結（area＝エリア内文字を新規作成 / それ以外＝先頭フレームへ集約）*/
    function concatHorizontalSingleLine(rowFrames, textMode) {
        var row = concatRowText(rowFrames);
        var sorted = row.sorted;
        var resultFrame;
        if (textMode === "area") {
            resultFrame = app.activeDocument.textFrames.areaText(makeFramelessRect(getSelBounds(sorted)));
            resultFrame.contents = row.text;
            resultFrame.textRange.characterAttributes.textFont = sorted[0].textRange.characterAttributes.textFont;
            resultFrame.textRange.characterAttributes.size = sorted[0].textRange.characterAttributes.size;
            resultFrame.paragraphs[0].justification = Justification.LEFT;
        } else {
            sorted[0].contents = row.text;
            sorted[0].paragraphs[0].justification = Justification.LEFT;
            resultFrame = sorted[0];
        }
        for (var k = 0; k < sorted.length; k++) {
            if (sorted[k] !== resultFrame) sorted[k].remove();
        }
        return [resultFrame];
    }

    /* 連結済みの各行を、英文ハイフン除去／語間スペース／句点改行のルールで1つの文字列へ */
    function buildJoinedParagraphText(mergedFrames) {
        var finalText = "";
        for (var i = 0; i < mergedFrames.length; i++) {
            var content = stripTrailingBreaks(mergedFrames[i].contents);
            finalText += content;
            if (i < mergedFrames.length - 1) {
                var nextContent = stripTrailingBreaks(mergedFrames[i + 1].contents);
                /* 英単語がハイフンで分断 → ハイフン除去して結合 */
                if (/[A-Za-z0-9)]-$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                    finalText = finalText.replace(/-$/, "");
                }
                /* 末尾・次行頭がともに英単語 → 語間スペース */
                else if (/[A-Za-z0-9)]$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                    finalText += " ";
                }
                /* 句点等で終わる場合のみ改行を残す */
                if (shouldInsertParagraphBreakBetweenLines(content)) finalText += "\r";
            }
        }
        return finalText;
    }

    /* 連結結果テキストを point/area で出力（font/size/kinsoku/justification を設定）*/
    function createConcatOutputText(textMode, finalText, bounds, font, fontSize) {
        var newTF;
        if (textMode === "point") {
            newTF = app.activeDocument.textFrames.pointText([bounds[0], bounds[1]]);
            newTF.contents = finalText;
            newTF.textRange.characterAttributes.textFont = font;
            newTF.textRange.characterAttributes.size = fontSize;
            newTF.paragraphs[0].paragraphAttributes.kinsoku = "Soft";
            newTF.textRange.justification = Justification.LEFT;
        } else {
            newTF = app.activeDocument.textFrames.areaText(makeFramelessRect(bounds));
            newTF.contents = finalText;
            newTF.textRange.characterAttributes.textFont = font;
            newTF.textRange.characterAttributes.size = fontSize;
            newTF.paragraphs[0].paragraphAttributes.kinsoku = "Soft";
            newTF.textRange.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        }
        return newTF;
    }

    function concatHorizontal(objects, textMode) {
        var LINE_Y_THRESHOLD = 5;
        var MIN_LEADING_RATIO = 1.2;

        var textItems = getTextFrames(objects);
        if (textItems.length < 2) return textItems;

        if (!textMode || textMode === "mixed") {
            textMode = detectTextFrameType(objects);
            if (textMode === "mixed") textMode = "area";
        }

        /* 強制改行を削除 */
        removeForcedLineBreaks(textItems);

        /* Y位置で行ごとにグループ化 */
        var textLines = groupByLineY(sortByY(textItems), LINE_Y_THRESHOLD);

        /* 1行だけの場合 */
        if (textLines.length === 1) {
            return concatHorizontalSingleLine(textLines[0], textMode);
        }

        /* 複数行：行ごとにX順で連結して中間フレームを作る */
        var mergedFrames = [];
        for (var i = 0; i < textLines.length; i++) {
            var row = concatRowText(textLines[i]);
            var mf = row.sorted[0].duplicate();
            mf.contents = row.text;
            mf.position = row.sorted[0].position;
            mergedFrames.push(mf);
            for (var k = 0; k < row.sorted.length; k++) row.sorted[k].remove();
        }

        var font = mergedFrames[0].textRange.characterAttributes.textFont;
        var fontSize = mergedFrames[0].textRange.characterAttributes.size;
        var finalText = buildJoinedParagraphText(mergedFrames);
        var bounds = getSelBounds(mergedFrames);
        var newTF = createConcatOutputText(textMode, finalText, bounds, font, fontSize);

        /* 行間を設定 */
        if (mergedFrames.length >= 2) {
            var leading = Math.abs(mergedFrames[0].position[1] - mergedFrames[1].position[1]);
            if (leading < fontSize) leading = fontSize * MIN_LEADING_RATIO;
            newTF.textRange.characterAttributes.autoLeading = false;
            newTF.textRange.characterAttributes.leading = leading;
        }

        /* 元のフレームを削除 */
        for (var m = 0; m < mergedFrames.length; m++) mergedFrames[m].remove();

        app.redraw();
        return [newTF];
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
        isTabChar, hasVisibleChars,
        getTextFrames, detectTextFrameType, countTextFrameTypes, countBreakTypes, transformContents,
        removeForcedLineBreaks, sortByPosition, sortByY, sortByX, groupByLineY, getSelBounds, groupTextFrames,
        removeLineBreaks, removeAllBreaks, flattenToOneLine, removeEmptyLines, removeTabs, tabsToSpaces,
        trimSpaces, collapseSpaces, fullToHalfAlnum, halfToFullKana, toggleBulletList, toggleNumberList,
        reverseOrder, removeDuplicateLines, sortByCharCode, sortByLength, removeCjkLatinSpaces,
        addLineBreakPerChar, addLineBreakAtCount, convertForcedLineBreaks, convertToForcedBreaks,
        addLineBreakAtPunctuation, splitByTab, splitByLineBreak, splitByLineBreakKeepStyle,
        splitByCharKeepStyle, splitByCharIgnoreStyle, stripStyleKeepFirstFont, splitCharHighPrecision,
        buildOutlineCharBounds, sortOutlineItems, estimateCharRowThreshold, copyCharAttrs,
        moveFrameToMatchBounds, splitCharFallback, concatHorizontalOnly, concatVertical, concatHorizontal,
        concatRowText, makeFramelessRect, concatHorizontalSingleLine, buildJoinedParagraphText, createConcatOutputText,
        toWordCap, toSentenceCase, toTitleCase,
        safeSet, copyAttr, applyResetStyle
    ];

    /* 関数配列を toString() で連結してソース文字列にする */
    function buildLibSource(funcs) {
        var src = "";
        for (var i = 0; i < funcs.length; i++) src += funcs[i].toString() + "\n";
        return src;
    }

    var __LIB_SRC_CACHE = null;
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
    function __dispatch(ACTION, P) {
        if (app.documents.length === 0) return "ERR:nodoc";
        var doc;
        try { doc = app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        function hasMultipleLines(objs) {
            var f = getTextFrames(objs);
            for (var i = 0; i < f.length; i++) {
                if (splitParagraphLines(f[i].contents).length >= 2) return true;
            }
            return false;
        }
        function hasMultipleTextFrames(objs) { return getTextFrames(objs).length >= 2; }
        function hasSpacesOrTabs(objs) {
            var f = getTextFrames(objs);
            for (var i = 0; i < f.length; i++) {
                if (/[ \t　]/.test(f[i].contents)) return true;
            }
            return false;
        }
        function encState(objs) {
            var c = countTextFrameTypes(objs), b = countBreakTypes(objs);
            return [c.total, c.point, c.area, b.paragraph, b.forced, b.tab,
                hasMultipleLines(objs) ? 1 : 0, hasMultipleTextFrames(objs) ? 1 : 0, hasSpacesOrTabs(objs) ? 1 : 0].join("|");
        }
        /* 分割結果のグループを解除し、中身を親へ出してバラの状態にする */
        function ungroupResult(result) {
            if (!result || !result.length) return result;
            var frames = getTextFrames(result);
            for (var i = 0; i < result.length; i++) {
                var it = result[i];
                try {
                    if (it && it.typename === "GroupItem") {
                        var parent = it.parent;
                        for (var j = it.pageItems.length - 1; j >= 0; j--) {
                            it.pageItems[j].move(parent, ElementPlacement.PLACEATEND);
                        }
                        it.remove();
                    }
                } catch (e) { }
            }
            return frames;
        }
        function runAction(name, t) {
            switch (name) {
                case "flatten": { var r = flattenToOneLine(t); var tt = (r && r.length) ? r : t; trimSpaces(tt); removeCjkLatinSpaces(tt); collapseSpaces(tt); return tt; }
                case "removeLineBreaks": if (P.forced) { removeAllBreaks(t); } else { removeLineBreaks(t); } return t;
                case "addLineBreakPerChar": addLineBreakPerChar(t); return t;
                case "punctuation": addLineBreakAtPunctuation(t, P.chars); return t;
                case "breakAtCount": addLineBreakAtCount(t, P.count, P.forced); return t;
                case "convertForcedLineBreaks": convertForcedLineBreaks(t); return t;
                case "convertToForcedBreaks": convertToForcedBreaks(t); return t;
                case "splitByLineBreak": return ungroupResult(splitByLineBreak(t));
                case "splitByLineBreakKeepStyle": return splitByLineBreakKeepStyle(t);
                case "splitByTab": return splitByTab(t);
                case "splitByCharKeepStyle": return splitByCharKeepStyle(t);
                case "splitByCharIgnoreStyle": return splitByCharIgnoreStyle(t);
                case "concatVertical": return concatVertical(t);
                case "concatHorizontalOnly": return concatHorizontalOnly(t);
                case "concatH": return concatHorizontal(t, detectTextFrameType(t));
                case "concatToArea": return concatHorizontal(t, "area");
                case "removeTabs": removeTabs(t); return t;
                case "tabsToSpaces": tabsToSpaces(t); return t;
                case "trimSpaces": trimSpaces(t); return t;
                case "removeCjkLatinSpaces": removeCjkLatinSpaces(t); return t;
                case "collapseSpaces": collapseSpaces(t); return t;
                case "cleanupSpaces": trimSpaces(t); removeCjkLatinSpaces(t); collapseSpaces(t); return t;
                case "removeAllSpaces": transformContents(t, function (s) { return s.replace(/[ 　]/g, ""); }); return t;
                case "fullToHalfAlnum": fullToHalfAlnum(t); return t;
                case "halfToFullKana": halfToFullKana(t); return t;
                case "toggleBulletList": toggleBulletList(t); return t;
                case "toggleNumberList": toggleNumberList(t); return t;
                case "reverseOrder": reverseOrder(t); return t;
                case "removeDuplicateLines": removeDuplicateLines(t); return t;
                case "sortByCharCode": sortByCharCode(t); return t;
                case "sortByLength": sortByLength(t); return t;
                case "removeEmptyLines": removeEmptyLines(t); return t;
                case "caseUpper": transformContents(t, function (s) { return s.toUpperCase(); }); return t;
                case "caseLower": transformContents(t, function (s) { return s.toLowerCase(); }); return t;
                case "caseWord": transformContents(t, toWordCap); return t;
                case "caseSentence": transformContents(t, toSentenceCase); return t;
                case "caseTitle": transformContents(t, toTitleCase); return t;
                case "convertSymbol": {
                    /* ES の入れ子三項は左結合に誤評価されるため括弧で右結合を明示 */
                    var fromRe = (P.from === "underscore") ? /_/g : ((P.from === "hyphen") ? /-/g : /[ 　]/g);
                    var toStr = (P.to === "space") ? " " : ((P.to === "hyphen") ? "-" : "_");
                    transformContents(t, function (s) { return s.replace(fromRe, toStr); });
                    return t;
                }
                case "spaceAfterPunct": transformContents(t, function (s) { return s.replace(/([.,])(?=[^\s\d.,])/g, "$1 "); }); return t;
            }
            return t;
        }

        var targets = getTextFrames(doc.selection);

        if (ACTION === "getState") {
            return "OK:" + encState(targets);
        }
        if (ACTION === "getLines") {
            if (targets.length === 1) return "LINES:" + encodeURIComponent(targets[0].contents);
            return "LINES:";
        }
        if (ACTION === "getFirstText") {
            if (targets.length >= 1) return "TEXT:" + encodeURIComponent(targets[0].contents);
            return "TEXT:";
        }
        if (ACTION === "setLines") {
            if (targets.length === 1) { targets[0].contents = P.text; }
            app.redraw();
            return "OK:" + encState(getTextFrames(doc.selection));
        }
        if (ACTION === "undo") {
            try { app.executeMenuCommand("undo"); } catch (e2) { }
            app.redraw();
            return "OK:" + encState(getTextFrames(doc.selection));
        }
        if (ACTION === "hiddenChar") {
            try { app.executeMenuCommand("showHiddenChar"); } catch (e3) { }
            app.redraw();
            return "OK:" + encState(targets);
        }
        if (ACTION === "finalizeClose") {
            for (var i = 0; i < targets.length; i++) {
                try {
                    var p = targets[i].parent;
                    if (p && p.typename === "GroupItem" && p.pageItems.length === 1) {
                        var container = p.parent;
                        targets[i].move(container, ElementPlacement.PLACEATEND);
                        p.remove();
                    }
                } catch (e4) { }
            }
            try { if (targets.length > 0) doc.selection = targets; } catch (e5) { }
            if (P.turnOffHidden) { try { app.executeMenuCommand("showHiddenChar"); } catch (e6) { } }
            app.redraw();
            return "OK:";
        }

        /* 通常の処理アクション */
        if (!targets || targets.length === 0) return "OK:" + encState([]);
        var result;
        try {
            result = runAction(ACTION, targets);
        } catch (errProc) {
            try { app.redraw(); } catch (e7) { }
            return "ERR:" + (errProc && errProc.message ? errProc.message : String(errProc));
        }
        var refreshed = getTextFrames((result && result.length) ? result : targets);
        if (refreshed.length > 0) { try { doc.selection = refreshed; } catch (e8) { } targets = refreshed; }
        else { targets = []; }
        app.redraw();
        return "OK:" + encState(targets);
    }

    var __DISPATCH_SRC = "(" + __dispatch.toString() + ")";

    /* パラメータを安全な JS リテラル文字列へ */
    function paramsToSource(p) {
        if (!p) return "{}";
        var parts = [];
        if (p.forced !== undefined) parts.push("forced:" + (p.forced ? "true" : "false"));
        if (p.turnOffHidden !== undefined) parts.push("turnOffHidden:" + (p.turnOffHidden ? "true" : "false"));
        if (p.count !== undefined) parts.push("count:" + parseInt(p.count, 10));
        if (p.chars !== undefined) parts.push('chars:decodeURIComponent("' + encodeURIComponent(p.chars) + '")');
        if (p.text !== undefined) parts.push('text:decodeURIComponent("' + encodeURIComponent(p.text) + '")');
        if (p.from !== undefined) parts.push('from:"' + p.from + '"');
        if (p.to !== undefined) parts.push('to:"' + p.to + '"');
        return "{" + parts.join(",") + "}";
    }

    /*
     * 本文をメインエンジンへ送り、結果マーカーを解析して onDone(status, payload) を呼ぶ。
     * BridgeTalk は本文送信時にバックスラッシュをエスケープするため（"\r" → ターゲットで "\\r"）、
     * コード全体を encodeURIComponent で包んで送り（%エンコードに \ は出ない）、
     * ターゲットで decodeURIComponent + eval して元ソースに復元してから実行する。
     * status = "ok" | "lines" | "text" | "error"
     */
    function sendWorker(code, onDone) {
        var bt = new BridgeTalk();
        bt.target = "illustrator";
        bt.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
        bt.onResult = function (res) {
            var payload = res.body || "";
            var ci = payload.indexOf(":");
            var marker = ci >= 0 ? payload.substring(0, ci) : payload;
            var rest = ci >= 0 ? payload.substring(ci + 1) : "";
            if (marker === "OK") onDone("ok", rest);
            else if (marker === "LINES") onDone("lines", rest);
            else if (marker === "TEXT") onDone("text", rest);
            else onDone("error", rest);
        };
        bt.onError = function (res) { onDone("error", res && res.body ? res.body : "BridgeTalk error"); };
        bt.send();
    }

    /* メインエンジンへアクションを委譲（非同期）*/
    function runWorker(actionId, params, onDone) {
        var code = getLibSource() + "\nvar __r=" + __DISPATCH_SRC + "(\"" + actionId + "\"," + paramsToSource(params) + ");__r;";
        sendWorker(code, onDone);
    }

    /* state 文字列（"total|point|area|para|forced|tab|multiLines|multiFrames|hasSpTab"）をオブジェクトへ */
    function parseState(rest) {
        var a = String(rest || "").split("|");
        return {
            total: parseInt(a[0], 10) || 0,
            point: parseInt(a[1], 10) || 0,
            area: parseInt(a[2], 10) || 0,
            para: parseInt(a[3], 10) || 0,
            forced: parseInt(a[4], 10) || 0,
            tab: parseInt(a[5], 10) || 0,
            multiLines: a[6] === "1",
            multiFrames: a[7] === "1",
            hasSpTab: a[8] === "1"
        };
    }

    /* =========================================
     * 軽量ステータスポーラー（選択のリアルタイム反映用）
     *
     * 定期ポーリングで毎回 LIB 全体を送ると重いので、件数集計に
     * 必要な関数だけを同梱した軽量本文をメインエンジンへ送る。
     * 選択は変更せず、state 文字列だけを返す。
     * ========================================= */
    var __STATE_FUNCS = [
        debugLog, normalizeParagraphBreaks, splitParagraphLines,
        isParagraphBreak, isForcedBreak, isTabChar,
        getTextFrames, countTextFrameTypes, countBreakTypes
    ];
    var __STATE_LIB_CACHE = null;
    function getStateLibSource() {
        if (__STATE_LIB_CACHE === null) __STATE_LIB_CACHE = buildLibSource(__STATE_FUNCS);
        return __STATE_LIB_CACHE;
    }

    /* メインエンジンで現在の選択を集計して state を返す（選択は変えない）*/
    function __stateDispatch() {
        if (app.documents.length === 0) return "ERR:nodoc";
        var doc;
        try { doc = app.activeDocument; } catch (e) { return "ERR:nodoc"; }
        var t = getTextFrames(doc.selection);
        function hasMultipleLines(objs) { var f = getTextFrames(objs); for (var i = 0; i < f.length; i++) { if (splitParagraphLines(f[i].contents).length >= 2) return true; } return false; }
        function hasMultipleFrames(objs) { return getTextFrames(objs).length >= 2; }
        function hasSpaceOrTab(objs) { var f = getTextFrames(objs); for (var i = 0; i < f.length; i++) { if (/[ \t　]/.test(f[i].contents)) return true; } return false; }
        var c = countTextFrameTypes(t), b = countBreakTypes(t);
        return "OK:" + [c.total, c.point, c.area, b.paragraph, b.forced, b.tab, hasMultipleLines(t) ? 1 : 0, hasMultipleFrames(t) ? 1 : 0, hasSpaceOrTab(t) ? 1 : 0].join("|");
    }
    var __STATE_DISPATCH_SRC = "(" + __stateDispatch.toString() + ")";

    function runStatePoll(onDone) {
        var code = getStateLibSource() + "\nvar __r=" + __STATE_DISPATCH_SRC + "();__r;";
        sendWorker(code, onDone);
    }

    /* ダイアログボックスを作成・表示する関数 */
    function showDialog(selectedObjects) {
        var dialog = new Window("palette", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);

        function hasMultipleLines(objects) {
            var frames = getTextFrames(objects);
            for (var i = 0; i < frames.length; i++) {
                if (splitParagraphLines(frames[i].contents).length >= 2) {
                    return true;
                }
            }
            return false;
        }

        function hasMultipleTextFrames(objects) {
            return getTextFrames(objects).length >= 2;
        }

        /* 各処理をメインエンジンへ委譲し、結果のステータスで表示を更新 */
        function executeAction(actionId, params) {
            executeActionThen(actionId, params, null);
        }

        /* executeAction と同じだが、成功後に after() を呼ぶ（行リスト再読込などの連鎖用） */
        function executeActionThen(actionId, params, after) {
            runWorker(actionId, params || {}, function (status, payload) {
                if (status === "error") {
                    showError({ message: payload });
                    return;
                }
                var st = parseState(payload);
                updateStatusDisplay(st);
                updateActionAvailability(st);
                if (after) after();
            });
        }

        /* 起動直後（パレット表示前）はこのエンジンでも DOM を読めるので、
           初期ステータスはローカルで集計する */
        function computeStateLocal(objects) {
            var c = countTextFrameTypes(objects), b = countBreakTypes(objects);
            return {
                total: c.total, point: c.point, area: c.area,
                para: b.paragraph, forced: b.forced, tab: b.tab,
                multiLines: hasMultipleLines(objects),
                multiFrames: hasMultipleTextFrames(objects),
                hasSpTab: hasSpacesOrTabs(objects)
            };
        }

        /* 処理対象件数を集計 */
        var textFrameCounts = countTextFrameTypes(selectedObjects);
        var breakCounts = countBreakTypes(selectedObjects);

        /* ステータスパネル */
        var panelStatus = dialog.add("panel", undefined, getLabel(LABELS.panel.status));
        panelStatus.margins = [20, 20, 30, 10];
        panelStatus.alignment = ["fill", "top"];
        panelStatus.alignChildren = ["left", "top"];

        /* ステータス表示 */
        var statusRow = panelStatus.add("group");
        statusRow.orientation = "row";
        statusRow.alignment = ["fill", "top"];
        statusRow.alignChildren = ["left", "center"];
        statusRow.spacing = 30;

        var frameCountColumn = statusRow.add("group");
        frameCountColumn.orientation = "column";
        frameCountColumn.alignment = ["fill", "top"];
        frameCountColumn.alignChildren = ["left", "top"];

        var breakCountColumn = statusRow.add("group");
        breakCountColumn.orientation = "column";
        breakCountColumn.alignment = ["fill", "top"];
        breakCountColumn.alignChildren = ["left", "top"];


        var statusCenterLabelWidth = 70;

        /* 1行 = [ラベル：] [値] を作り、値の statictext を返す */
        function addStatusRow(column, labelNode, value, labelWidth) {
            var row = column.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            var lbl = row.add("statictext", undefined, labelText(labelNode));
            if (labelWidth) lbl.preferredSize.width = labelWidth;
            return row.add("statictext", undefined, String(value));
        }

        var valTargetCount = addStatusRow(frameCountColumn, LABELS.info.targetCount, textFrameCounts.total);
        var valPointCount = addStatusRow(frameCountColumn, LABELS.info.pointCount, textFrameCounts.point);
        var valAreaCount = addStatusRow(frameCountColumn, LABELS.info.areaCount, textFrameCounts.area);
        var valParagraphBreakCount = addStatusRow(breakCountColumn, LABELS.info.paragraphBreak, breakCounts.paragraph, statusCenterLabelWidth);
        var valForcedBreakCount = addStatusRow(breakCountColumn, LABELS.info.forcedBreak, breakCounts.forced, statusCenterLabelWidth);
        var valTabCount = addStatusRow(breakCountColumn, LABELS.info.tab, breakCounts.tab, statusCenterLabelWidth);

        /* ステータス表示を state で更新 */
        function updateStatusDisplay(state) {
            valTargetCount.text = String(state.total);
            valPointCount.text = String(state.point);
            valAreaCount.text = String(state.area);
            valParagraphBreakCount.text = String(state.para);
            valForcedBreakCount.text = String(state.forced);
            valTabCount.text = String(state.tab);
        }

        function hasSpacesOrTabs(objects) {
            var frames = getTextFrames(objects);
            for (var i = 0; i < frames.length; i++) {
                var txt = frames[i].contents;
                if (/[ \t\u3000]/.test(txt)) return true;
            }
            return false;
        }

        function updateActionAvailability(state) {
            var multiLines = state.multiLines;
            var multiFrames = state.multiFrames;
            var hasSpTab = state.hasSpTab;

            btnReverseOrder.enabled = multiLines;
            btnRemoveDuplicateLines.enabled = multiLines;
            btnRemoveEmptyLines.enabled = multiLines;
            btnSortByCharCode.enabled = multiLines;
            btnSortByLength.enabled = multiLines;
            btnSplitByLine.enabled = multiLines;
            btnSplitByLineKeepStyle.enabled = multiLines;
            var hasParagraph = state.para > 0;
            var hasForced = state.forced > 0;
            var hasAnyBreaks = hasParagraph || hasForced;

            /* 改行削除系 */
            btnRemoveLineBreaks.enabled = hasAnyBreaks;
            chkIncludeForcedBreaks.enabled = hasForced;

            /* 改行変換系 */
            btnConvertBreaks.enabled = hasForced;
            btnConvertToForcedBreaks.enabled = hasParagraph;

            /* スペシャル */
            btnFlattenToOneLine.enabled = hasAnyBreaks || multiFrames;

            /* 分割系 */
            btnSplitByTab.enabled = state.tab > 0;

            /* 連結系 */
            btnConcatV.enabled = multiFrames;
            btnConcatHOnly.enabled = multiFrames;
            btnConcatH.enabled = multiFrames;
            btnConcatToArea.enabled = multiFrames;

            /* スペース系 */
            btnTrimSpaces.enabled = hasSpTab;
            btnCjkLatinSpaces.enabled = hasSpTab;
            btnCollapseSpaces.enabled = hasSpTab;
            btnCleanupSpaces.enabled = hasSpTab;
            btnRemoveAllSpaces.enabled = hasSpTab;

            /* タブ系 */
            btnRemoveTabs.enabled = state.tab > 0;
            btnTabsToSpaces.enabled = state.tab > 0;
        }



        /* 制御文字の状態管理 */
        var hiddenCharOn = false;
        var hiddenCharLabel = getLabel(LABELS.checkbox.showHiddenChar);

        /* タブパネル（メイン） */
        var tabbedPanel = dialog.add("tabbedpanel");
        // tabbedPanel.margins = [20, 15, 0, -10];
        tabbedPanel.alignment = ["fill", "top"];
        tabbedPanel.alignChildren = ["fill", "top"];

        /* === タブ1: 基本 === */
        var tabBasic = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.basic));
        tabBasic.margins = [10, 20, 0, -10];
        tabBasic.spacing = 15;
        tabBasic.orientation = "row";
        tabBasic.alignment = ["center", "top"];
        tabBasic.alignChildren = ["fill", "top"];

        /* 左カラム：改行 */
        var breakColumn = tabBasic.add("group");
        breakColumn.orientation = "column";
        breakColumn.alignment = ["fill", "top"];
        breakColumn.alignChildren = ["fill", "top"];

        /* 改行グループパネル */
        var panelBreakGroup = breakColumn.add("panel", undefined, getLabel(LABELS.panel.breakGroup));
        panelBreakGroup.margins = [15, 20, 15, 10];
        panelBreakGroup.alignment = ["fill", "top"];
        panelBreakGroup.alignChildren = ["fill", "top"];

        /* すべて1行にボタン（中央配置） */
        var flattenButtonRow = panelBreakGroup.add("group");
        flattenButtonRow.alignment = ["fill", "top"];
        flattenButtonRow.alignChildren = ["fill", "center"];
        flattenButtonRow.margins = [15, 0, 15, 10];

        var btnFlattenToOneLine = flattenButtonRow.add("button", undefined, getLabel(LABELS.button.flattenToOneLine));
        btnFlattenToOneLine.onClick = function () {
            executeAction("flatten");
        };

        /* 改行削除パネル */
        var panelRemoveBreak = panelBreakGroup.add("panel", undefined, getLabel(LABELS.panel.removeBreak));
        panelRemoveBreak.margins = [15, 20, 15, 10];
        panelRemoveBreak.alignment = ["fill", "top"];
        panelRemoveBreak.alignChildren = ["fill", "center"];

        /* 改行削除ボタン */
        var btnRemoveLineBreaks = panelRemoveBreak.add("button", undefined, getLabel(LABELS.button.removeLineBreaks));
        btnRemoveLineBreaks.onClick = function () {
            executeAction("removeLineBreaks", { forced: chkIncludeForcedBreaks.value });
        };

        /* 強制改行を含むチェックボックス */
        var chkIncludeForcedBreaks = panelRemoveBreak.add("checkbox", undefined, getLabel(LABELS.checkbox.includeForcedBreaks));

        /* 改行パネル */
        var panelLineBreak = panelBreakGroup.add("panel", undefined, getLabel(LABELS.panel.insertBreak));
        panelLineBreak.margins = [15, 20, 15, 10];
        panelLineBreak.alignment = ["fill", "top"];
        panelLineBreak.alignChildren = ["fill", "center"];

        /* 1文字ごとに改行ボタン */
        var btnAddLineBreaks = panelLineBreak.add("button", undefined, getLabel(LABELS.button.addLineBreaks));
        btnAddLineBreaks.onClick = function () {
            executeAction("addLineBreakPerChar");
        };

        /* 句読点で改行ボタン */
        var btnPunctuation = panelLineBreak.add("button", undefined, getLabel(LABELS.button.punctuation));
        btnPunctuation.onClick = function () {
            executeAction("punctuation", { chars: txtPunctuation.text });
        };

        /* 句読点対象文字テキストフィールド */
        var txtPunctuation = panelLineBreak.add("edittext", undefined, "、。，．｡､,.!?！？");
        txtPunctuation.alignment = ["fill", "center"];

        /* 指定文字数で改行ボタン */
        var btnBreakAtCount = panelLineBreak.add("button", undefined, getLabel(LABELS.button.breakAtCount));
        btnBreakAtCount.onClick = function () {
            executeAction("breakAtCount", { count: txtBreakCount.text, forced: chkForcedBreakAtCount.value });
        };

        /* 文字数テキストフィールド + 強制改行チェックボックス */
        var breakCountRow = panelLineBreak.add("group");
        breakCountRow.orientation = "row";
        breakCountRow.alignment = ["fill", "center"];
        breakCountRow.alignChildren = ["left", "center"];
        var txtBreakCount = breakCountRow.add("edittext", undefined, "35");
        txtBreakCount.characters = 3;
        var chkForcedBreakAtCount = breakCountRow.add("checkbox", undefined, getLabel(LABELS.checkbox.forcedBreak));

        /* その他の改行パネル */
        var panelOtherBreak = panelBreakGroup.add("panel", undefined, getLabel(LABELS.panel.convertBreak));
        panelOtherBreak.margins = [15, 20, 15, 10];
        panelOtherBreak.alignment = ["fill", "top"];
        panelOtherBreak.alignChildren = ["fill", "center"];

        /* 強制改行→改行ボタン */
        var btnConvertBreaks = panelOtherBreak.add("button", undefined, getLabel(LABELS.button.convertBreaks));
        btnConvertBreaks.onClick = function () {
            executeAction("convertForcedLineBreaks");
        };

        /* 改行→強制改行ボタン */
        var btnConvertToForcedBreaks = panelOtherBreak.add("button", undefined, getLabel(LABELS.button.convertToForcedBreaks));
        btnConvertToForcedBreaks.onClick = function () {
            executeAction("convertToForcedBreaks");
        };

        /* 右カラム：分割・連結 */
        var splitConcatColumn = tabBasic.add("group");
        splitConcatColumn.orientation = "column";
        splitConcatColumn.alignment = ["fill", "top"];
        splitConcatColumn.alignChildren = ["fill", "top"];

        /* 分割グループパネル */
        var panelSplitGroup = splitConcatColumn.add("panel", undefined, getLabel(LABELS.panel.splitGroup));
        panelSplitGroup.margins = [15, 20, 15, 10];
        panelSplitGroup.alignment = ["fill", "top"];
        panelSplitGroup.alignChildren = ["fill", "top"];

        /* 分割パネル */
        var panelSplit = panelSplitGroup.add("panel", undefined, getLabel(LABELS.panel.splitByBreak));
        panelSplit.margins = [15, 20, 15, 10];
        panelSplit.alignment = ["fill", "top"];
        panelSplit.alignChildren = ["fill", "center"];

        /* 改行で分割ボタン */
        var btnSplitByLine = panelSplit.add("button", undefined, getLabel(LABELS.button.splitByLine));
        btnSplitByLine.onClick = function () {
            executeAction("splitByLineBreak");
        };

        /* 改行で分割（書式保持）ボタン */
        var btnSplitByLineKeepStyle = panelSplit.add("button", undefined, getLabel(LABELS.button.splitByLineKeepStyle));
        btnSplitByLineKeepStyle.onClick = function () {
            executeAction("splitByLineBreakKeepStyle");
        };

        /* タブで分解ボタン */
        var btnSplitByTab = panelSplit.add("button", undefined, getLabel(LABELS.button.splitByTab));
        btnSplitByTab.onClick = function () {
            executeAction("splitByTab");
        };

        /* 分割（文字）パネル */
        var panelSplitChar = panelSplitGroup.add("panel", undefined, getLabel(LABELS.panel.splitByChar));
        panelSplitChar.margins = [15, 20, 15, 10];
        panelSplitChar.alignment = ["fill", "top"];
        panelSplitChar.alignChildren = ["fill", "center"];

        /* 書式を保持ボタン */
        var btnSplitKeepStyle = panelSplitChar.add("button", undefined, getLabel(LABELS.button.splitKeepStyle));
        btnSplitKeepStyle.onClick = function () {
            executeAction("splitByCharKeepStyle");
        };

        /* 書式を無視ボタン */
        var btnSplitIgnoreStyle = panelSplitChar.add("button", undefined, getLabel(LABELS.button.splitIgnoreStyle));
        btnSplitIgnoreStyle.onClick = function () {
            executeAction("splitByCharIgnoreStyle");
        };

        /* 連結パネル */
        var panelConcat = splitConcatColumn.add("panel", undefined, getLabel(LABELS.panel.concat));
        panelConcat.margins = [15, 20, 15, 10];
        panelConcat.alignment = ["fill", "top"];
        panelConcat.alignChildren = ["center", "center"];

        var concatButtonColumn = panelConcat.add("group");
        concatButtonColumn.margins = [15, 0, 15, 10];
        concatButtonColumn.orientation = "column";
        concatButtonColumn.alignment = ["fill", "top"];
        concatButtonColumn.alignChildren = ["fill", "center"];

        /* 連結（縦）ボタン */
        var btnConcatV = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatV));
        btnConcatV.helpTip = getLabel(LABELS.tooltip.concatV);
        btnConcatV.onClick = function () {
            executeAction("concatVertical");
        };

        /* 横連結（行維持）ボタン */
        var btnConcatHOnly = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatHOnly));
        btnConcatHOnly.helpTip = getLabel(LABELS.tooltip.concatHOnly);
        btnConcatHOnly.onClick = function () {
            executeAction("concatHorizontalOnly");
        };

        /* 横連結ボタン */
        var btnConcatH = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatH));
        btnConcatH.helpTip = getLabel(LABELS.tooltip.concatH);
        btnConcatH.onClick = function () {
            executeAction("concatH");
        };

        /* PDFテキスト整形ボタン */
        var btnConcatToArea = concatButtonColumn.add("button", undefined, getLabel(LABELS.button.concatToArea));
        btnConcatToArea.helpTip = getLabel(LABELS.tooltip.concatToArea);
        btnConcatToArea.onClick = function () {
            executeAction("concatToArea");
        };

        /* === タブ2: クリーンアップ === */
        var tabCleanup = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.cleanup));
        tabCleanup.margins = [10, 20, 0, -10];
        tabCleanup.spacing = 15;
        tabCleanup.orientation = "row";
        tabCleanup.alignment = ["fill", "top"];
        tabCleanup.alignChildren = ["fill", "top"];

        /* 左カラム：タブ */
        var tabSpaceColumn = tabCleanup.add("group");
        tabSpaceColumn.orientation = "column";
        tabSpaceColumn.alignment = ["fill", "top"];
        tabSpaceColumn.alignChildren = ["fill", "top"];

        /* タブパネル */
        var panelTab = tabSpaceColumn.add("panel", undefined, getLabel(LABELS.panel.tab));
        panelTab.margins = [15, 20, 15, 10];
        panelTab.alignment = ["fill", "top"];
        panelTab.alignChildren = ["fill", "center"];

        /* タブを削除ボタン */
        var btnRemoveTabs = panelTab.add("button", undefined, getLabel(LABELS.button.removeTabs));
        btnRemoveTabs.onClick = function () {
            executeAction("removeTabs");
        };

        /* タブをスペースにボタン */
        var btnTabsToSpaces = panelTab.add("button", undefined, getLabel(LABELS.button.tabsToSpaces));
        btnTabsToSpaces.onClick = function () {
            executeAction("tabsToSpaces");
        };

        /* スペースパネル（タブパネルの下） */
        var panelSpace = tabSpaceColumn.add("panel", undefined, getLabel(LABELS.panel.space));
        panelSpace.margins = [15, 20, 15, 10];
        panelSpace.alignment = ["fill", "top"];
        panelSpace.alignChildren = ["fill", "center"];

        /* 行頭行末のスペースボタン */
        var btnTrimSpaces = panelSpace.add("button", undefined, getLabel(LABELS.button.trimSpaces));
        btnTrimSpaces.onClick = function () {
            executeAction("trimSpaces");
        };

        /* 和欧間のスペースボタン */
        var btnCjkLatinSpaces = panelSpace.add("button", undefined, getLabel(LABELS.button.cjkLatinSpaces));
        btnCjkLatinSpaces.onClick = function () {
            executeAction("removeCjkLatinSpaces");
        };

        /* 連続スペースボタン */
        var btnCollapseSpaces = panelSpace.add("button", undefined, getLabel(LABELS.button.collapseSpaces));
        btnCollapseSpaces.onClick = function () {
            executeAction("collapseSpaces");
        };

        /* スペース削除（一括）ボタン */
        var btnCleanupSpaces = panelSpace.add("button", undefined, getLabel(LABELS.button.cleanupSpaces));
        btnCleanupSpaces.onClick = function () {
            executeAction("cleanupSpaces");
        };

        /* すべてのスペースを削除ボタン */
        var btnRemoveAllSpaces = panelSpace.add("button", undefined, getLabel(LABELS.button.removeAllSpaces));
        btnRemoveAllSpaces.onClick = function () {
            executeAction("removeAllSpaces");
        };

        /* その他パネル（スペース削除パネルの下）*/
        var panelOther = tabSpaceColumn.add("panel", undefined, getLabel(LABELS.panel.other));
        panelOther.margins = [15, 20, 15, 10];
        panelOther.alignment = ["fill", "top"];
        panelOther.alignChildren = ["fill", "center"];

        var btnSpaceAfterPunct = panelOther.add("button", undefined, getLabel(LABELS.button.spaceAfterPunct));
        btnSpaceAfterPunct.onClick = function () {
            executeAction("spaceAfterPunct");
        };

        /* 右カラム：変換・リスト */
        var convertListColumn = tabCleanup.add("group");
        convertListColumn.orientation = "column";
        convertListColumn.alignment = ["fill", "top"];
        convertListColumn.alignChildren = ["fill", "top"];

        /* スペースや記号の変換パネル（右カラム先頭・変換の上）
         * 変換前（Before）／変換後（After）をラジオで選び、［変換］で置換。
         * 前後が同じ選択のときは［変換］をディムにする。*/
        var panelSymbolConvert = convertListColumn.add("panel", undefined, getLabel(LABELS.panel.symbolConvert));
        panelSymbolConvert.margins = [15, 20, 15, 10];
        panelSymbolConvert.alignment = ["fill", "top"];
        panelSymbolConvert.alignChildren = ["fill", "top"];

        var symbolRow = panelSymbolConvert.add("group");
        symbolRow.orientation = "row";
        symbolRow.alignment = ["fill", "top"];
        symbolRow.alignChildren = ["fill", "top"];
        symbolRow.spacing = 15;

        /* 変換前（Before）*/
        var beforePanel = symbolRow.add("panel", undefined, getLabel(LABELS.panel.symbolBefore));
        beforePanel.orientation = "column";
        beforePanel.alignChildren = ["left", "top"];
        beforePanel.margins = [15, 20, 15, 10];
        var rbBeforeSpace = beforePanel.add("radiobutton", undefined, getLabel(LABELS.radio.space));
        var rbBeforeUnderscore = beforePanel.add("radiobutton", undefined, getLabel(LABELS.radio.underscore));
        var rbBeforeHyphen = beforePanel.add("radiobutton", undefined, getLabel(LABELS.radio.hyphen));
        rbBeforeSpace.value = true;

        /* 変換後（After）*/
        var afterPanel = symbolRow.add("panel", undefined, getLabel(LABELS.panel.symbolAfter));
        afterPanel.orientation = "column";
        afterPanel.alignChildren = ["left", "top"];
        afterPanel.margins = [15, 20, 15, 10];
        var rbAfterSpace = afterPanel.add("radiobutton", undefined, getLabel(LABELS.radio.space));
        var rbAfterUnderscore = afterPanel.add("radiobutton", undefined, getLabel(LABELS.radio.underscore));
        var rbAfterHyphen = afterPanel.add("radiobutton", undefined, getLabel(LABELS.radio.hyphen));
        rbAfterUnderscore.value = true;

        function getBeforeSymbol() {
            if (rbBeforeUnderscore.value) return "underscore";
            if (rbBeforeHyphen.value) return "hyphen";
            return "space";
        }
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

        /* 前後が同じならディム */
        function updateConvertSymbolState() {
            btnConvertSymbol.enabled = (getBeforeSymbol() !== getAfterSymbol());
        }
        var symbolRadios = [rbBeforeSpace, rbBeforeUnderscore, rbBeforeHyphen, rbAfterSpace, rbAfterUnderscore, rbAfterHyphen];
        for (var rbi = 0; rbi < symbolRadios.length; rbi++) {
            symbolRadios[rbi].onClick = updateConvertSymbolState;
        }
        updateConvertSymbolState();

        /* 変換パネル */
        var panelConvert = convertListColumn.add("panel", undefined, getLabel(LABELS.panel.convert));
        panelConvert.margins = [15, 20, 15, 10];
        panelConvert.alignment = ["fill", "top"];
        panelConvert.alignChildren = ["center", "center"];

        /* 全角英数字→半角ボタン */
        var btnFullToHalfAlnum = panelConvert.add("button", undefined, getLabel(LABELS.button.fullToHalfAlnum));
        btnFullToHalfAlnum.onClick = function () {
            executeAction("fullToHalfAlnum");
        };

        /* 半角カナ→全角ボタン */
        var btnHalfToFullKana = panelConvert.add("button", undefined, getLabel(LABELS.button.halfToFullKana));
        btnHalfToFullKana.onClick = function () {
            executeAction("halfToFullKana");
        };

        /* リストパネル */
        var panelList = convertListColumn.add("panel", undefined, getLabel(LABELS.panel.list));
        panelList.margins = [15, 20, 15, 10];
        panelList.alignment = ["fill", "top"];
        panelList.alignChildren = ["center", "center"];

        /* 箇条書きボタン */
        var btnBulletList = panelList.add("button", undefined, getLabel(LABELS.button.bulletList));
        btnBulletList.onClick = function () {
            executeAction("toggleBulletList");
        };

        /* 番号リストボタン */
        var btnNumberList = panelList.add("button", undefined, getLabel(LABELS.button.numberList));
        btnNumberList.onClick = function () {
            executeAction("toggleNumberList");
        };

        /* === タブ3: 行の整理 === */
        var tabLineArrange = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.lineArrange));
        tabLineArrange.margins = [10, 20, 0, -10];
        tabLineArrange.spacing = 15;
        tabLineArrange.orientation = "row";
        tabLineArrange.alignment = ["fill", "top"];
        tabLineArrange.alignChildren = ["fill", "top"];

        /* 左カラム：行の並び替え（リストボックス＋操作ボタン） */
        var lineListColumn = tabLineArrange.add("group");
        lineListColumn.orientation = "column";
        lineListColumn.alignment = ["fill", "top"];
        lineListColumn.alignChildren = ["fill", "top"];

        var lineListRow = lineListColumn.add("group");
        lineListRow.orientation = "row";
        lineListRow.alignChildren = ["fill", "fill"];
        lineListRow.spacing = 10;

        var lineListBox = lineListRow.add("listbox", undefined, [], {
            multiselect: false
        });
        lineListBox.preferredSize = [200, 460];
        lineListBox.graphics.font = ScriptUI.newFont("dialog", "REGULAR", 18);

        /* リストボックスのデータ管理 */
        var lineArrangeLines = [];

        function applyLinesToTextFrame() {
            executeAction("setLines", { text: lineArrangeLines.join("\r") });
        }

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

        function loadLinesToList() {
            runWorker("getLines", {}, function (status, payload) {
                lineArrangeLines = [];
                if (status === "lines" && payload) {
                    var contents = decodeURIComponent(payload);
                    var normalized = contents.replace(/\r\n/g, "\r").replace(/\n/g, "\r");
                    lineArrangeLines = normalized.split("\r");
                }
                refreshLineList(0, true);
            });
        }

        function updateLineListButtons() {
            var hasSel = lineListBox.selection !== null;
            btnLineUp.enabled = hasSel && lineListBox.selection.index > 0;
            btnLineDown.enabled = hasSel && lineListBox.selection.index < lineArrangeLines.length - 1;
            btnLineEdit.enabled = hasSel;
            btnLineDelete.enabled = hasSel;
        }

        lineListBox.onDoubleClick = function () {
            if (!lineListBox.selection) return;
            var idx = lineListBox.selection.index;
            var result = prompt(getLabel(LABELS.prompt.editLine), lineArrangeLines[idx]);
            if (result === null) return;
            lineArrangeLines[idx] = result;
            refreshLineList(idx);
        };

        lineListBox.onChange = function () {
            updateLineListButtons();
        };

        /* 右カラム：一括操作ボタン */
        var lineActionColumn = tabLineArrange.add("group");
        lineActionColumn.orientation = "column";
        lineActionColumn.alignment = ["fill", "top"];
        lineActionColumn.alignChildren = ["fill", "top"];

        /* 編集パネル */
        var panelLineEdit = lineActionColumn.add("panel", undefined, getLabel(LABELS.panel.lineEdit));
        panelLineEdit.margins = [15, 20, 15, 10];
        panelLineEdit.alignment = ["fill", "top"];
        panelLineEdit.alignChildren = ["fill", "center"];

        /* 上へボタン */
        var btnLineUp = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineUp));
        btnLineUp.onClick = function () {
            if (!lineListBox.selection) return;
            var idx = lineListBox.selection.index;
            if (idx <= 0) return;
            var tmp = lineArrangeLines[idx];
            lineArrangeLines[idx] = lineArrangeLines[idx - 1];
            lineArrangeLines[idx - 1] = tmp;
            refreshLineList(idx - 1);
        };

        /* 下へボタン */
        var btnLineDown = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineDown));
        btnLineDown.onClick = function () {
            if (!lineListBox.selection) return;
            var idx = lineListBox.selection.index;
            if (idx >= lineArrangeLines.length - 1) return;
            var tmp = lineArrangeLines[idx];
            lineArrangeLines[idx] = lineArrangeLines[idx + 1];
            lineArrangeLines[idx + 1] = tmp;
            refreshLineList(idx + 1);
        };

        /* 追加ボタン */
        var btnLineAdd = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineAdd));
        btnLineAdd.onClick = function () {
            var result = prompt(getLabel(LABELS.prompt.addLine), "");
            if (result === null) return;
            lineArrangeLines.push(result);
            refreshLineList(lineArrangeLines.length - 1);
        };

        /* 編集ボタン */
        var btnLineEdit = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineEdit));
        btnLineEdit.onClick = function () {
            if (!lineListBox.selection) return;
            var idx = lineListBox.selection.index;
            var result = prompt(getLabel(LABELS.prompt.editLine), lineArrangeLines[idx]);
            if (result === null) return;
            lineArrangeLines[idx] = result;
            refreshLineList(idx);
        };

        /* 削除ボタン */
        var btnLineDelete = panelLineEdit.add("button", undefined, getLabel(LABELS.button.lineDelete));
        btnLineDelete.onClick = function () {
            if (!lineListBox.selection) return;
            var idx = lineListBox.selection.index;
            if (!confirm(getLabel(LABELS.confirm.deleteLine))) return;
            lineArrangeLines.splice(idx, 1);
            refreshLineList(idx);
        };

        /* ソートパネル */
        var panelSort = lineActionColumn.add("panel", undefined, getLabel(LABELS.panel.sort));
        panelSort.margins = [15, 20, 15, 10];
        panelSort.alignment = ["fill", "top"];
        panelSort.alignChildren = ["fill", "center"];

        /* ソート（文字コード）ボタン */
        var btnSortByCharCode = panelSort.add("button", undefined, getLabel(LABELS.button.sortByCharCode));
        btnSortByCharCode.onClick = function () {
            executeActionThen("sortByCharCode", {}, loadLinesToList);
        };

        /* ソート（文字数）ボタン */
        var btnSortByLength = panelSort.add("button", undefined, getLabel(LABELS.button.sortByLength));
        btnSortByLength.onClick = function () {
            executeActionThen("sortByLength", {}, loadLinesToList);
        };

        /* 順序を反転ボタン */
        var btnReverseOrder = panelSort.add("button", undefined, getLabel(LABELS.button.reverseOrder));
        btnReverseOrder.onClick = function () {
            executeActionThen("reverseOrder", {}, loadLinesToList);
        };

        /* 行削除パネル */
        var panelLineDelete = lineActionColumn.add("panel", undefined, getLabel(LABELS.panel.lineDelete));
        panelLineDelete.margins = [15, 20, 15, 10];
        panelLineDelete.alignment = ["fill", "top"];
        panelLineDelete.alignChildren = ["fill", "center"];

        /* 重複行の削除ボタン */
        var btnRemoveDuplicateLines = panelLineDelete.add("button", undefined, getLabel(LABELS.button.removeDuplicateLines));
        btnRemoveDuplicateLines.onClick = function () {
            executeActionThen("removeDuplicateLines", {}, loadLinesToList);
        };

        /* 空行削除ボタン */
        var btnRemoveEmptyLines = panelLineDelete.add("button", undefined, getLabel(LABELS.button.removeEmptyLines));
        btnRemoveEmptyLines.onClick = function () {
            executeActionThen("removeEmptyLines", {}, loadLinesToList);
        };

        /* === タブ4: 英数字 === */
        var tabAlnum = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab.alnum));
        tabAlnum.margins = [10, 20, 0, -10];
        tabAlnum.spacing = 15;
        tabAlnum.orientation = "column";
        tabAlnum.alignment = ["fill", "top"];
        tabAlnum.alignChildren = ["fill", "top"];

        /* 英字の大文字/小文字変換（ボタン＋プレビュー）*/
        var panelLetterCase = tabAlnum.add("panel", undefined, getLabel(LABELS.panel.letterCase));
        panelLetterCase.margins = [15, 20, 15, 10];
        panelLetterCase.alignment = ["fill", "top"];
        panelLetterCase.alignChildren = ["fill", "top"];

        /* mode -> プレビュー用 statictext */
        var casePreview = {};

        /* 1行 = [変換ボタン] [プレビュー] */
        function addCaseRow(modeKey, labelNode, actionId) {
            var row = panelLetterCase.add("group");
            row.orientation = "row";
            row.alignment = ["fill", "center"];
            row.alignChildren = ["left", "center"];

            var btn = row.add("button", undefined, getLabel(labelNode));
            btn.preferredSize = [180, 24];
            btn.onClick = function () {
                executeActionThen(actionId, {}, refreshCasePreview);
            };

            var preview = row.add("statictext", undefined, "");
            preview.preferredSize = [240, 24];
            try { preview.justify = "left"; } catch (e) { }
            casePreview[modeKey] = preview;
        }
        addCaseRow("upper", LABELS.button.caseUpper, "caseUpper");
        addCaseRow("lower", LABELS.button.caseLower, "caseLower");
        addCaseRow("word", LABELS.button.caseWord, "caseWord");
        addCaseRow("sentence", LABELS.button.caseSentence, "caseSentence");
        addCaseRow("title", LABELS.button.caseTitle, "caseTitle");

        /* プレビュー表示用にテキストを短く整える */
        function normalizeExample(s) {
            if (s == null) return "";
            s = String(s).replace(/[\r\n]+/g, " ").replace(/[ 　\t]+/g, " ").replace(/^\s+|\s+$/g, "");
            if (s.length > 40) s = s.substring(0, 40) + "…";
            return s;
        }

        /* メインエンジンから先頭テキストを取得し、各モードのプレビューを更新 */
        var casePreviewRefreshing = false;
        function refreshCasePreview() {
            if (casePreviewRefreshing) return;
            casePreviewRefreshing = true;
            runWorker("getFirstText", {}, function (status, payload) {
                casePreviewRefreshing = false;
                var base = (status === "text" && payload) ? decodeURIComponent(payload) : "";
                var src = normalizeExample(base);
                if (casePreview.upper) casePreview.upper.text = normalizeExample(src.toUpperCase());
                if (casePreview.lower) casePreview.lower.text = normalizeExample(src.toLowerCase());
                if (casePreview.word) casePreview.word.text = normalizeExample(toWordCap(src));
                if (casePreview.sentence) casePreview.sentence.text = normalizeExample(toSentenceCase(src));
                if (casePreview.title) casePreview.title.text = normalizeExample(toTitleCase(src));
            });
        }

        var initState = computeStateLocal(selectedObjects);
        updateStatusDisplay(initState);
        updateActionAvailability(initState);
        loadLinesToList();
        refreshCasePreview();

        /* === 選択のリアルタイム反映 ===
         * Illustrator 30.x には app.scheduleTask / setTimeout 等のタイマー API が
         * 無いため、キャンバスにフォーカスがある間の連続ポーリングはできない。
         * 代わりに「ユーザーがパレットへ操作しに来た瞬間」に選択を取り直す：
         *   - onActivate : フォーカス復帰（パレットをクリック）時
         *   - mouseover  : マウスがパレット上に乗った時
         * メインエンジンの選択を取り直してステータスへ反映する。
         * 多重実行ガード（statusRefreshing）で BridgeTalk の重複呼び出しを防ぐ。*/
        var statusRefreshing = false;
        function refreshStatusFromSelection() {
            if (statusRefreshing) return;
            statusRefreshing = true;
            runStatePoll(function (status, payload) {
                statusRefreshing = false;
                if (status !== "ok") return;
                var st = parseState(payload);
                updateStatusDisplay(st);
                updateActionAvailability(st);
            });
        }
        function onPaletteFocus() {
            refreshStatusFromSelection();
            /* 英数字タブ表示中は選択変化に合わせてプレビューも更新 */
            if (tabbedPanel.selection === tabAlnum) refreshCasePreview();
        }
        dialog.onActivate = onPaletteFocus;
        try { dialog.addEventListener("mouseover", onPaletteFocus); } catch (e) { }

        /* タブ切り替え時の再読み込み */
        tabbedPanel.onChange = function () {
            if (tabbedPanel.selection === tabLineArrange) {
                loadLinesToList();
            } else if (tabbedPanel.selection === tabAlnum) {
                refreshCasePreview();
            }
        };

        /* ボタン行（左：制御文字、右：1つ戻す・閉じる） */
        var footerRow = dialog.add("group");
        footerRow.alignment = ["fill", "bottom"];
        // footerRow.margins = [0, 15, 0, 0];

        /* 左側：制御文字ボタン */
        var footerLeftGroup = footerRow.add("group");
        footerLeftGroup.alignment = ["left", "center"];

        var btnShowHiddenChar = footerLeftGroup.add("button", undefined, hiddenCharLabel);
        btnShowHiddenChar.preferredSize = [120, -1];

        function updateHiddenCharButton() {
            if (hiddenCharOn) {
                btnShowHiddenChar.text = "\u2713 " + hiddenCharLabel;
                try {
                    var gfx = btnShowHiddenChar.graphics;
                    gfx.foregroundColor = gfx.newPen(gfx.PenType.SOLID_COLOR, [0.0, 0.5, 0.8], 1);
                } catch (_) { }
            } else {
                btnShowHiddenChar.text = hiddenCharLabel;
                try {
                    var gfx2 = btnShowHiddenChar.graphics;
                    gfx2.foregroundColor = gfx2.newPen(gfx2.PenType.SOLID_COLOR, [0.0, 0.0, 0.0], 1);
                } catch (_) { }
            }
        }

        btnShowHiddenChar.onClick = function () {
            runWorker("hiddenChar", {}, function (status, payload) {
                if (status === "error") { showError({ message: payload }); return; }
                hiddenCharOn = !hiddenCharOn;
                updateHiddenCharButton();
                var st = parseState(payload);
                updateStatusDisplay(st);
                updateActionAvailability(st);
            });
        };

        /* 右側：1つ戻す・閉じる */
        var footerRightGroup = footerRow.add("group");
        footerRightGroup.alignment = ["right", "center"];

        /* パレットを閉じる（1要素だけのグループ解除・選択整理・制御文字OFF はメインエンジンで）*/
        function closePalette() {
            runWorker("finalizeClose", { turnOffHidden: hiddenCharOn }, function () {
                dialog.close();
            });
        }

        var btnClose = footerRightGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "ok" });
        btnClose.onClick = closePalette;

        /* パレットがアクティブなとき Esc で閉じる */
        dialog.addEventListener("keydown", function (e) {
            if (e.keyName === "Escape") closePalette();
        });

        /* パレットが GC で破棄されないよう参照を保持 */
        $.global.__TextBreakSplitMergePalette = dialog;
        dialog.onClose = function () {
            try { $.global.__TextBreakSplitMergePalette = null; } catch (_) { }
        };

        dialog.show();
    }

    try {
        showDialog(selectedObjects);
    } catch (err) {
        showError(err);
    }
})();