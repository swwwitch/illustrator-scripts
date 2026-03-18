#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// TextBreakSplitMerge.jsx 
// 
// 作成日：2026年3月18日
// 更新日：2026年3月18日
//
// Illustratorで分散しがちなテキスト処理（改行・分割・連結・整形）を
// 1つのダイアログに統合したオールインワン・テキスト処理スクリプトです。
//
// 従来は個別スクリプトとして分散しがちなテキスト処理（改行削除・空行整理・分割・連結など）を、
// 単一UIに集約することで、スクリプト管理やショートカット運用の負担を軽減します。
//
// 対象は以下の選択状態に対応します：
// ・テキストフレーム
// ・テキストフレームを含むグループ
// ・Illustratorが TextRange として返すテキスト選択
//
// ［主な機能］
// ・改行の削除／挿入／変換（段落改行・強制改行の相互変換を含む）
// ・段落／タブ／文字単位でのテキスト分割（書式保持／無視の両対応）
// ・複数テキストの縦方向・横方向連結（見た目ベースでの再構成）
// ・テキスト内容のクリーンアップ（空行削除、スペース整理、タブ処理など）
// ・テキスト構造の可視化（改行数・タブ数・文字タイプの集計表示）
//
// ［設計方針］
// ・既存テキストを直接編集しつつ、結果を即座に確認できる即時実行型UI
// ・複数スクリプトの置き換えを前提とした「集約ツール」設計
// ・見た目ベースの処理を優先し、実務での作業効率を重視
//
// ダイアログUIは日本語／英語を自動切り替えし、タイトルバーにはバージョン番号を表示します。

var SCRIPT_VERSION = "v1.3";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 */
var LABELS = {
    dialogTitle: {
        ja: "テキスト処理",
        en: "Text Processing"
    },
    panelBreakGroup: {
        ja: "改行",
        en: "Breaks"
    },
    panelRemoveBreak: {
        ja: "削除",
        en: "Remove"
    },
    panelLineBreak: {
        ja: "挿入",
        en: "Insert"
    },
    panelOtherBreak: {
        ja: "切り換え",
        en: "Convert"
    },
    panelSplitGroup: {
        ja: "分割",
        en: "Split"
    },
    panelSplit: {
        ja: "分割（行）",
        en: "Split (Line)"
    },
    panelSplitChar: {
        ja: "分割（文字）",
        en: "Split (Char)"
    },
    panelOther: {
        ja: "その他",
        en: "Other"
    },
    btnReverseOrder: {
        ja: "順序を反転",
        en: "Reverse Order"
    },
    btnRemoveDuplicateLines: {
        ja: "重複行の削除",
        en: "Remove Duplicates"
    },
    btnSortByCharCode: {
        ja: "ソート",
        en: "Sort"
    },
    btnSortByLength: {
        ja: "ソート（文字数）",
        en: "Sort (Length)"
    },
    btnSplitKeepStyle: {
        ja: "書式を保持",
        en: "Keep Style"
    },
    btnSplitIgnoreStyle: {
        ja: "書式を無視",
        en: "Ignore Style"
    },
    panelConcat: {
        ja: "連結",
        en: "Concatenate"
    },
    panelCleanup: {
        ja: "クリーンアップ",
        en: "Cleanup"
    },
    panelTab: {
        ja: "タブ",
        en: "Tab"
    },
    panelSpace: {
        ja: "スペース",
        en: "Spaces"
    },
    btnCollapseSpaces: {
        ja: "連続スペース",
        en: "Collapse Spaces"
    },
    btnCjkLatinSpaces: {
        ja: "和欧間スペースを削除",
        en: "Remove CJK-Latin Spaces"
    },
    btnRemoveTabs: {
        ja: "タブを削除",
        en: "Remove Tabs"
    },
    btnTabsToSpaces: {
        ja: "タブをスペースに",
        en: "Tabs to Spaces"
    },
    btnTrimSpaces: {
        ja: "行頭行末のスペース",
        en: "Leading/Trailing Spaces"
    },
    panelStatus: {
        ja: "ステータス",
        en: "Status"
    },
    panelSpecial: {
        ja: "スペシャル",
        en: "Special"
    },
    btnFlattenToOneLine: {
        ja: "とにかく1行に",
        en: "Flatten to One Line"
    },
    btnRemoveLineBreaks: {
        ja: "改行のみ",
        en: "Line Breaks Only"
    },
    btnRemoveAllBreaks: {
        ja: "強制改行を含む",
        en: "Include Forced Breaks"
    },
    btnRemoveEmptyLines: {
        ja: "空行削除",
        en: "Remove Empty Lines"
    },
    btnAddLineBreaks: {
        ja: "1文字ごと",
        en: "Insert Line Break After Each Character"
    },
    btnPunctuation: {
        ja: "指定文字で",
        en: "At Specified Characters"
    },
    btnBreakAtCount: {
        ja: "指定文字数で",
        en: "At Character Count"
    },
    btnConvertBreaks: {
        ja: "強制改行→改行",
        en: "Forced Breaks to Paragraph Breaks"
    },
    btnConvertToForcedBreaks: {
        ja: "改行→強制改行",
        en: "Paragraph Breaks to Forced Breaks"
    },
    btnSplitByLine: {
        ja: "改行で分割",
        en: "Split by Line Breaks"
    },
    btnSplitByTab: {
        ja: "タブと改行",
        en: "Split by Tabs and Line Breaks"
    },
    btnConcatV: {
        ja: "縦連結",
        en: "Vertical"
    },
    btnConcatHOnly: {
        ja: "横連結（行維持）",
        en: "Horizontal (Keep Rows)"
    },
    btnConcatH: {
        ja: "横連結（行統合）",
        en: "Horizontal (Combine Rows)"
    },
    btnConcatToArea: {
        ja: "横連結→エリア",
        en: "To Area Text"
    },
    tipConcatV: {
        ja: "上→下に連結",
        en: "Merge top to bottom"
    },
    tipConcatHOnly: {
        ja: "横連結（行維持）",
        en: "Merge per row (keep rows)"
    },
    tipConcatH: {
        ja: "横連結（行統合）",
        en: "Merge horizontally (combine rows)"
    },
    tipConcatToArea: {
        ja: "横連結→エリア内文字",
        en: "Merge to area text"
    },
    chkShowHiddenChar: {
        ja: "制御文字",
        en: "Hidden Characters"
    },
    btnUndo: {
        ja: "1つ戻す",
        en: "Undo"
    },
    btnClose: {
        ja: "閉じる",
        en: "Close"
    },
    errProcessFailed: {
        ja: "処理中にエラーが発生しました。\n\n",
        en: "An error occurred while processing.\n\n"
    },
    msgNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    msgNoSelection: {
        ja: "テキストフレーム、またはテキストを含むグループを選択してください。",
        en: "Please select a text frame or a group containing text."
    },
    msgNoTextFrames: {
        ja: "選択内に対象のテキストフレームが見つかりません。テキストフレーム、またはテキストを含むグループを選択してください。",
        en: "No target text frames were found in the selection. Please select a text frame or a group containing text."
    },
    infoTargetCount: {
        ja: "対象テキスト: ",
        en: "Target Texts: "
    },
    infoPointAreaCount: {
        ja: "ポイント文字: ",
        en: "Point Type: "
    },
    infoAreaSeparator: {
        ja: "エリア内文字: ",
        en: "Area Type: "
    },
    infoParagraphBreakCount: {
        ja: "改行: ",
        en: "Paragraph Breaks: "
    },
    infoForcedBreakCount: {
        ja: "強制改行: ",
        en: "Forced Breaks: "
    },
    infoTabCount: {
        ja: "タブ: ",
        en: "Tabs: "
    },
};

function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    return entry[lang] || entry.ja || entry.en || key;
}

/* エラー表示補助 */
function showError(err) {
    var msg = L("errProcessFailed");
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

(function () {
    // ドキュメントが開かれていない場合は処理を終了
    if (app.documents.length === 0) {
        alert(L("msgNoDocument"));
        return;
    }

    // 選択オブジェクトを取得
    var selectedObjects = app.selection;

    // 選択オブジェクトがない場合は処理を終了
    try {
        if (!selectedObjects || (typeof selectedObjects.length === "number" && selectedObjects.length === 0)) {
            alert(L("msgNoSelection"));
            return;
        }
    } catch (e) {
        if (!selectedObjects) {
            alert(L("msgNoSelection"));
            return;
        }
    }

    // 初期選択からテキストフレームを解決
    selectedObjects = getTextFrames(selectedObjects);
    if (selectedObjects.length === 0) {
        alert(L("msgNoTextFrames"));
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

    /* テキストフレームをグループ化 */
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

    /* 複数フレームを連結し、全改行を削除して1行にする関数 */
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

    /* テキストフレーム内の行を文字コード順にソート */
    function sortByCharCode(objects) {
        transformContents(objects, function (txt) {
            var lines = splitParagraphLines(txt);
            lines.sort();
            return lines.join("\r");
        });
    }

    /* テキストフレーム内の行を文字数順にソート */
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
    function addLineBreakAtCount(objects, count) {
        var n = parseInt(count, 10);
        if (!n || n <= 0) n = 35;
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
            return result.join("\r");
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

    /* 書式リセット（先頭文字のフォント情報のみ保持） */
    function stripStyleKeepFirstFont(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;
        var tr, chars;
        try { tr = textFrame.textRange; } catch (_) { return; }
        try { chars = tr.characters; } catch (_) { return; }
        if (!chars || chars.length < 1) return;

        var firstCA;
        try { firstCA = chars[0].characterAttributes; } catch (_) { return; }
        var keepFont, keepSize;
        try { keepFont = firstCA.textFont; } catch (_) { keepFont = null; }
        try { keepSize = firstCA.size; } catch (_) { keepSize = null; }

        var ca;
        try { ca = tr.characterAttributes; } catch (_) { return; }
        try { if (keepFont) ca.textFont = keepFont; } catch (_) { }
        try { if (keepSize != null) ca.size = keepSize; } catch (_) { }
        try { var black = new GrayColor(); black.gray = 100; ca.fillColor = black; } catch (_) { }
        try { ca.baselineShift = 0; } catch (_) { }
        try { ca.horizontalScale = 100; } catch (_) { }
        try { ca.verticalScale = 100; } catch (_) { }
        try { ca.rotation = 0; } catch (_) { }
        try { ca.tracking = 0; } catch (_) { }
        try { ca.kerningMethod = KerningMethod.METRICS; } catch (_) { }
        try { ca.autoLeading = true; } catch (_) { }

        for (var i = 0; i < chars.length; i++) {
            var c2;
            try { c2 = chars[i].characterAttributes; } catch (_) { continue; }
            try { if (keepFont) c2.textFont = keepFont; } catch (_) { }
            try { if (keepSize != null) c2.size = keepSize; } catch (_) { }
            try { var bk = new GrayColor(); bk.gray = 100; c2.fillColor = bk; } catch (_) { }
            try { c2.baselineShift = 0; } catch (_) { }
            try { c2.horizontalScale = 100; } catch (_) { }
            try { c2.verticalScale = 100; } catch (_) { }
            try { c2.rotation = 0; } catch (_) { }
            try { c2.tracking = 0; } catch (_) { }
            try { c2.kerningMethod = KerningMethod.METRICS; } catch (_) { }
            try { c2.autoLeading = true; } catch (_) { }
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
        try { dst.textFont = src.textFont; } catch (e) { debugLog("copyCharAttrs: textFont", e); }
        try { dst.size = src.size; } catch (e) { debugLog("copyCharAttrs: size", e); }
        try { dst.horizontalScale = src.horizontalScale; } catch (e) { debugLog("copyCharAttrs: horizontalScale", e); }
        try { dst.verticalScale = src.verticalScale; } catch (e) { debugLog("copyCharAttrs: verticalScale", e); }
        try { dst.tracking = src.tracking; } catch (e) { debugLog("copyCharAttrs: tracking", e); }
        try { dst.baselineShift = src.baselineShift; } catch (e) { debugLog("copyCharAttrs: baselineShift", e); }
        try { dst.rotation = src.rotation; } catch (e) { debugLog("copyCharAttrs: rotation", e); }
        try { if (src.fillColor && src.fillColor.typename !== "NoColor") dst.fillColor = src.fillColor; } catch (e) { debugLog("copyCharAttrs: fillColor", e); }
        try { if (src.strokeColor && src.strokeColor.typename !== "NoColor") { dst.strokeColor = src.strokeColor; dst.strokeWeight = src.strokeWeight; } } catch (e) { debugLog("copyCharAttrs: strokeColor", e); }
        try { dst.autoLeading = src.autoLeading; } catch (e) { debugLog("copyCharAttrs: autoLeading", e); }
        try { if (!src.autoLeading) dst.leading = src.leading; } catch (e) { debugLog("copyCharAttrs: leading", e); }
        try { dst.kerningMethod = src.kerningMethod; } catch (e) { debugLog("copyCharAttrs: kerningMethod", e); }
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

    function concatHorizontal(objects, textMode) {
        var LINE_Y_THRESHOLD = 5;
        var MIN_LEADING_RATIO = 1.2;

        var textItems = getTextFrames(objects);
        if (textItems.length < 2) return textItems;

        if (!textMode || textMode === "mixed") {
            textMode = detectTextFrameType(objects);
            if (textMode === "mixed") {
                textMode = "area";
            }
        }

        /* 強制改行を削除 */
        removeForcedLineBreaks(textItems);

        /* Y位置で行ごとにグループ化 */
        var textLines = groupByLineY(
            sortByY(textItems), LINE_Y_THRESHOLD
        );

        /* 1行だけの場合：X順に連結して出力 */
        if (textLines.length === 1) {
            var sorted = sortByX(textLines[0]);
            var mergedText = "";
            var resultFrame;
            for (var j = 0; j < sorted.length; j++) {
                mergedText += stripTrailingBreaks(sorted[j].contents);
            }

            if (textMode === "area") {
                var oneLineBounds = getSelBounds(sorted);
                var oneLineWidth = oneLineBounds[2] - oneLineBounds[0];
                var oneLineHeight = oneLineBounds[1] - oneLineBounds[3];
                var oneLineRect = app.activeDocument.pathItems.rectangle(
                    oneLineBounds[1], oneLineBounds[0], oneLineWidth, oneLineHeight
                );
                oneLineRect.stroked = false;
                oneLineRect.filled = false;

                var oneLineAreaTF = app.activeDocument.textFrames.areaText(oneLineRect);
                oneLineAreaTF.contents = mergedText;
                oneLineAreaTF.textRange.characterAttributes.textFont = sorted[0].textRange.characterAttributes.textFont;
                oneLineAreaTF.textRange.characterAttributes.size = sorted[0].textRange.characterAttributes.size;
                oneLineAreaTF.paragraphs[0].justification = Justification.LEFT;
                resultFrame = oneLineAreaTF;
            } else {
                sorted[0].contents = mergedText;
                sorted[0].paragraphs[0].justification = Justification.LEFT;
                resultFrame = sorted[0];
            }

            for (var k = 0; k < sorted.length; k++) {
                if (sorted[k] !== resultFrame) {
                    sorted[k].remove();
                }
            }
            return [resultFrame];
        }

        /* 複数行：行ごとにX順で連結 */
        var mergedFrames = [];
        for (var i = 0; i < textLines.length; i++) {
            var sorted = sortByX(textLines[i]);
            var mergedText = "";
            for (var j = 0; j < sorted.length; j++) {
                mergedText += stripTrailingBreaks(sorted[j].contents);
            }
            var mf = sorted[0].duplicate();
            mf.contents = mergedText;
            mf.position = sorted[0].position;
            mergedFrames.push(mf);
            for (var k = 0; k < sorted.length; k++) {
                sorted[k].remove();
            }
        }

        /* 行間の連結処理（句点・ピリオド等で終わる場合のみ改行） */
        var finalText = "";
        var fontSize = mergedFrames[0].textRange.characterAttributes.size;

        for (var i = 0; i < mergedFrames.length; i++) {
            var content = stripTrailingBreaks(mergedFrames[i].contents);
            finalText += content;

            if (i < mergedFrames.length - 1) {
                var nextContent = stripTrailingBreaks(mergedFrames[i + 1].contents);

                /* 英単語がハイフンで分断されている場合、ハイフンを除去して結合 */
                if (/[A-Za-z0-9)]-$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                    finalText = finalText.replace(/-$/, "");
                }
                /* 末尾が英単語、次の行頭も英単語ならスペース挿入 */
                else if (/[A-Za-z0-9)]$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                    finalText += " ";
                }

                /* 句点等で終わる場合は改行 */
                if (shouldInsertParagraphBreakBetweenLines(content)) {
                    finalText += "\r";
                }
            }
        }

        /* 選択オブジェクトの幅を取得してエリアテキストに反映 */
        var bounds = getSelBounds(mergedFrames);
        var selWidth = bounds[2] - bounds[0];
        var selHeight = bounds[1] - bounds[3];

        var newTF;

        if (textMode === "point") {
            /* ポイントテキストを作成 */
            newTF = app.activeDocument.textFrames.pointText([bounds[0], bounds[1]]);
            newTF.contents = finalText;
            newTF.textRange.characterAttributes.textFont = mergedFrames[0].textRange.characterAttributes.textFont;
            newTF.textRange.characterAttributes.size = fontSize;
            newTF.paragraphs[0].paragraphAttributes.kinsoku = "Soft";
            newTF.textRange.justification = Justification.LEFT;
        } else {
            /* エリアテキストを作成 */
            var rect = app.activeDocument.pathItems.rectangle(
                bounds[1], bounds[0], selWidth, selHeight
            );
            rect.stroked = false;
            rect.filled = false;

            newTF = app.activeDocument.textFrames.areaText(rect);
            newTF.contents = finalText;
            newTF.textRange.characterAttributes.textFont = mergedFrames[0].textRange.characterAttributes.textFont;
            newTF.textRange.characterAttributes.size = fontSize;
            newTF.paragraphs[0].paragraphAttributes.kinsoku = "Soft";
            newTF.textRange.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        }

        /* 行間を設定 */
        if (mergedFrames.length >= 2) {
            var y1 = mergedFrames[0].position[1];
            var y2 = mergedFrames[1].position[1];
            var leading = Math.abs(y1 - y2);
            if (leading < fontSize) {
                leading = fontSize * MIN_LEADING_RATIO;
            }
            newTF.textRange.characterAttributes.autoLeading = false;
            newTF.textRange.characterAttributes.leading = leading;
        }

        /* 元のフレームを削除 */
        for (var i = 0; i < mergedFrames.length; i++) {
            mergedFrames[i].remove();
        }

        app.redraw();
        return [newTF];
    }

    /* ダイアログボックスを作成・表示する関数 */
    function showDialog(selectedObjects) {
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);

        /* 現在の処理対象を取得 */
        function getCurrentTargets() {
            return getTextFrames(selectedObjects);
        }

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

        /* 各処理を実行し、選択とステータス表示を更新 */
        function executeAction(actionFunc) {
            try {
                var targets = getCurrentTargets();
                if (!targets || targets.length === 0) return;

                var result = actionFunc(targets);
                var refreshedTargets = getTextFrames(result && result.length ? result : targets);

                if (refreshedTargets.length > 0) {
                    selectedObjects = refreshedTargets;
                    app.activeDocument.selection = refreshedTargets;
                    updateStatusDisplay(refreshedTargets);
                    updateActionAvailability(refreshedTargets);
                } else {
                    updateStatusDisplay([]);
                    updateActionAvailability([]);
                }

                app.redraw();
            } catch (err) {
                try { app.redraw(); } catch (redrawErr) { }
                showError(err);
            }
        }

        /* 処理対象件数を集計 */
        var textFrameCounts = countTextFrameTypes(selectedObjects);
        var breakCounts = countBreakTypes(selectedObjects);

        /* ステータスパネル */
        var panelStatus = dialog.add("panel", undefined, L("panelStatus"));
        panelStatus.margins = [15, 20, 15, 15];
        panelStatus.alignment = ["fill", "top"];
        panelStatus.alignChildren = ["left", "top"];

        /* ステータス表示 */
        var statusGroup = panelStatus.add("group");
        statusGroup.orientation = "row";
        statusGroup.alignment = ["fill", "top"];
        statusGroup.alignChildren = ["left", "center"];
        statusGroup.spacing = 20;

        var statusLeft = statusGroup.add("group");
        statusLeft.orientation = "column";
        statusLeft.alignChildren = ["left", "top"];

        var statusCenter = statusGroup.add("group");
        statusCenter.orientation = "column";
        statusCenter.alignChildren = ["left", "top"];

        var statusRight = statusGroup.add("group");
        statusRight.orientation = "column";
        statusRight.alignment = ["right", "center"];
        statusRight.alignChildren = ["right", "center"];
        statusRight.preferredSize = [100, -1];

        var statusLabelWidth = 120;

        var rowTargetCount = statusLeft.add("group");
        rowTargetCount.orientation = "row";
        rowTargetCount.alignChildren = ["left", "center"];
        var lblTargetCount = rowTargetCount.add("statictext", undefined, L("infoTargetCount"), { justify: "right" });
        lblTargetCount.preferredSize.width = statusLabelWidth;
        var valTargetCount = rowTargetCount.add("statictext", undefined, String(textFrameCounts.total));

        var rowPointCount = statusLeft.add("group");
        rowPointCount.orientation = "row";
        rowPointCount.alignChildren = ["left", "center"];
        var lblPointCount = rowPointCount.add("statictext", undefined, L("infoPointAreaCount"), { justify: "right" });
        lblPointCount.preferredSize.width = statusLabelWidth;
        var valPointCount = rowPointCount.add("statictext", undefined, String(textFrameCounts.point));

        var rowAreaCount = statusLeft.add("group");
        rowAreaCount.orientation = "row";
        rowAreaCount.alignChildren = ["left", "center"];
        var lblAreaCount = rowAreaCount.add("statictext", undefined, L("infoAreaSeparator"), { justify: "right" });
        lblAreaCount.preferredSize.width = statusLabelWidth;
        var valAreaCount = rowAreaCount.add("statictext", undefined, String(textFrameCounts.area));

        var statusRightLabelWidth = 80;

        var rowParagraphBreakCount = statusCenter.add("group");
        rowParagraphBreakCount.orientation = "row";
        rowParagraphBreakCount.alignChildren = ["left", "center"];
        var lblParagraphBreakCount = rowParagraphBreakCount.add("statictext", undefined, L("infoParagraphBreakCount"), { justify: "right" });
        lblParagraphBreakCount.preferredSize.width = statusRightLabelWidth;
        var valParagraphBreakCount = rowParagraphBreakCount.add("statictext", undefined, String(breakCounts.paragraph));

        var rowForcedBreakCount = statusCenter.add("group");
        rowForcedBreakCount.orientation = "row";
        rowForcedBreakCount.alignChildren = ["left", "center"];
        var lblForcedBreakCount = rowForcedBreakCount.add("statictext", undefined, L("infoForcedBreakCount"), { justify: "right" });
        lblForcedBreakCount.preferredSize.width = statusRightLabelWidth;
        var valForcedBreakCount = rowForcedBreakCount.add("statictext", undefined, String(breakCounts.forced));

        var rowTabCount = statusCenter.add("group");
        rowTabCount.orientation = "row";
        rowTabCount.alignChildren = ["left", "center"];
        var lblTabCount = rowTabCount.add("statictext", undefined, L("infoTabCount"), { justify: "right" });
        lblTabCount.preferredSize.width = statusRightLabelWidth;
        var valTabCount = rowTabCount.add("statictext", undefined, String(breakCounts.tab));

        lblTargetCount.enabled = true;
        valTargetCount.enabled = true;
        lblPointCount.enabled = true;
        valPointCount.enabled = true;
        lblAreaCount.enabled = true;
        valAreaCount.enabled = true;
        lblParagraphBreakCount.enabled = true;
        valParagraphBreakCount.enabled = true;
        lblForcedBreakCount.enabled = true;
        valForcedBreakCount.enabled = true;
        lblTabCount.enabled = true;
        valTabCount.enabled = true;

        /* ステータス表示を現在の選択状態で更新 */
        function updateStatusDisplay(objects) {
            var counts = countTextFrameTypes(objects);
            var breaks = countBreakTypes(objects);

            valTargetCount.text = String(counts.total);
            valPointCount.text = String(counts.point);
            valAreaCount.text = String(counts.area);
            valParagraphBreakCount.text = String(breaks.paragraph);
            valForcedBreakCount.text = String(breaks.forced);
            valTabCount.text = String(breaks.tab);
        }

        function updateActionAvailability(objects) {
            var multiLines = hasMultipleLines(objects);
            var multiFrames = hasMultipleTextFrames(objects);

            btnReverseOrder.enabled = multiLines;
            btnRemoveDuplicateLines.enabled = multiLines;
            btnSortByCharCode.enabled = multiLines;
            btnSortByLength.enabled = multiLines;
            btnSplitByLine.enabled = multiLines;

            btnConcatV.enabled = multiFrames;
            btnConcatHOnly.enabled = multiFrames;
            btnConcatH.enabled = multiFrames;
            btnConcatToArea.enabled = multiFrames;
        }



        /* 制御文字ボタン */
        var hiddenCharOn = false;
        var hiddenCharLabel = L("chkShowHiddenChar");
        var btnShowHiddenChar = statusRight.add("button", undefined, hiddenCharLabel);
        btnShowHiddenChar.preferredSize.width = 110;

        /* 制御文字ボタンの表示状態を更新 */

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

        /* 制御文字の表示切り替え */

        btnShowHiddenChar.onClick = function () {
            try {
                app.executeMenuCommand('showHiddenChar');
                app.redraw();
                hiddenCharOn = !hiddenCharOn;
                updateHiddenCharButton();
            } catch (err) {
                showError(err);
            }
        };

        /* スペシャルパネル */
        var panelSpecial = dialog.add("panel");
        panelSpecial.margins = [15, 15, 15, 15];
        panelSpecial.alignment = ["fill", "top"];
        panelSpecial.alignChildren = ["center", "top"];

        var specialBtnGroup = panelSpecial.add("group");
        specialBtnGroup.orientation = "row";
        specialBtnGroup.alignment = ["center", "center"];
        specialBtnGroup.alignChildren = ["center", "center"];

        /* とにかく1行にボタン */
        var btnFlattenToOneLine = specialBtnGroup.add("button", undefined, L("btnFlattenToOneLine"));
        btnFlattenToOneLine.onClick = function () {
            executeAction(flattenToOneLine);
        };

        /* 横連結→エリアボタン（連結パネルから移動） */
        var btnConcatToArea = specialBtnGroup.add("button", undefined, L("btnConcatToArea"));
        btnConcatToArea.helpTip = L("tipConcatToArea");
        btnConcatToArea.onClick = function () {
            executeAction(function (objects) {
                return concatHorizontal(objects, "area");
            });
            dialog.close();
        };

        /* 3カラムレイアウト */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignment = ["fill", "top"];
        columnsGroup.alignChildren = ["fill", "top"];

        /* 左カラム */
        var colLeft = columnsGroup.add("group");
        colLeft.orientation = "column";
        colLeft.alignChildren = ["fill", "top"];

        /* 改行グループパネル */
        var panelBreakGroup = colLeft.add("panel", undefined, L("panelBreakGroup"));
        panelBreakGroup.margins = [15, 20, 15, 10];
        panelBreakGroup.alignment = ["fill", "top"];
        panelBreakGroup.alignChildren = ["fill", "top"];

        /* 改行削除パネル */
        var panelRemoveBreak = panelBreakGroup.add("panel", undefined, L("panelRemoveBreak"));
        panelRemoveBreak.margins = [15, 20, 15, 10];
        panelRemoveBreak.alignment = ["fill", "top"];
        panelRemoveBreak.alignChildren = ["fill", "center"];

        /* 改行削除ボタン */
        var btnRemoveLineBreaks = panelRemoveBreak.add("button", undefined, L("btnRemoveLineBreaks"));
        btnRemoveLineBreaks.onClick = function () {
            executeAction(removeLineBreaks);
        };

        /* 強制改行を含むボタン */
        var btnRemoveAllBreaks = panelRemoveBreak.add("button", undefined, L("btnRemoveAllBreaks"));
        btnRemoveAllBreaks.onClick = function () {
            executeAction(removeAllBreaks);
        };

        /* 改行パネル */
        var panelLineBreak = panelBreakGroup.add("panel", undefined, L("panelLineBreak"));
        panelLineBreak.margins = [15, 20, 15, 10];
        panelLineBreak.alignment = ["fill", "top"];
        panelLineBreak.alignChildren = ["fill", "center"];

        /* 1文字ごとに改行ボタン */
        var btnAddLineBreaks = panelLineBreak.add("button", undefined, L("btnAddLineBreaks"));
        btnAddLineBreaks.onClick = function () {
            executeAction(addLineBreakPerChar);
        };

        /* 句読点で改行ボタン */
        var btnPunctuation = panelLineBreak.add("button", undefined, L("btnPunctuation"));
        btnPunctuation.onClick = function () {
            executeAction(function (objects) {
                return addLineBreakAtPunctuation(objects, txtPunctuation.text);
            });
        };

        /* 句読点対象文字テキストフィールド */
        var txtPunctuation = panelLineBreak.add("edittext", undefined, "、。，．｡､,.!?！？");
        txtPunctuation.alignment = ["fill", "center"];

        /* 指定文字数で改行ボタン */
        var btnBreakAtCount = panelLineBreak.add("button", undefined, L("btnBreakAtCount"));
        btnBreakAtCount.onClick = function () {
            executeAction(function (objects) {
                return addLineBreakAtCount(objects, txtBreakCount.text);
            });
        };

        /* 文字数テキストフィールド */
        var txtBreakCount = panelLineBreak.add("edittext", undefined, "35");
        txtBreakCount.alignment = ["fill", "center"];
        txtBreakCount.characters = 5;

        /* その他の改行パネル */
        var panelOtherBreak = panelBreakGroup.add("panel", undefined, L("panelOtherBreak"));
        panelOtherBreak.margins = [15, 20, 15, 10];
        panelOtherBreak.alignment = ["fill", "top"];
        panelOtherBreak.alignChildren = ["fill", "center"];

        /* 強制改行→改行ボタン */
        var btnConvertBreaks = panelOtherBreak.add("button", undefined, L("btnConvertBreaks"));
        btnConvertBreaks.onClick = function () {
            executeAction(convertForcedLineBreaks);
        };

        /* 改行→強制改行ボタン */
        var btnConvertToForcedBreaks = panelOtherBreak.add("button", undefined, L("btnConvertToForcedBreaks"));
        btnConvertToForcedBreaks.onClick = function () {
            executeAction(convertToForcedBreaks);
        };

        /* 中央カラム */
        var colCenter = columnsGroup.add("group");
        colCenter.orientation = "column";
        colCenter.alignChildren = ["fill", "top"];

        /* 右カラム */
        var colRight = columnsGroup.add("group");
        colRight.orientation = "column";
        colRight.alignChildren = ["fill", "top"];

        /* 分割グループパネル */
        var panelSplitGroup = colRight.add("panel", undefined, L("panelSplitGroup"));
        panelSplitGroup.margins = [15, 20, 15, 10];
        panelSplitGroup.alignment = ["fill", "top"];
        panelSplitGroup.alignChildren = ["fill", "top"];

        /* 分割パネル */
        var panelSplit = panelSplitGroup.add("panel", undefined, L("panelSplit"));
        panelSplit.margins = [15, 20, 15, 10];
        panelSplit.alignment = ["fill", "top"];
        panelSplit.alignChildren = ["fill", "center"];

        /* 改行で分割ボタン */
        var btnSplitByLine = panelSplit.add("button", undefined, L("btnSplitByLine"));
        btnSplitByLine.onClick = function () {
            executeAction(splitByLineBreak);
        };

        /* タブで分解ボタン */
        var btnSplitByTab = panelSplit.add("button", undefined, L("btnSplitByTab"));
        btnSplitByTab.onClick = function () {
            executeAction(splitByTab);
        };

        /* 分割（文字）パネル */
        var panelSplitChar = panelSplitGroup.add("panel", undefined, L("panelSplitChar"));
        panelSplitChar.margins = [15, 20, 15, 10];
        panelSplitChar.alignment = ["fill", "top"];
        panelSplitChar.alignChildren = ["fill", "center"];

        /* 書式を保持ボタン */
        var btnSplitKeepStyle = panelSplitChar.add("button", undefined, L("btnSplitKeepStyle"));
        btnSplitKeepStyle.onClick = function () {
            executeAction(splitByCharKeepStyle);
        };

        /* 書式を無視ボタン */
        var btnSplitIgnoreStyle = panelSplitChar.add("button", undefined, L("btnSplitIgnoreStyle"));
        btnSplitIgnoreStyle.onClick = function () {
            executeAction(splitByCharIgnoreStyle);
        };

        /* その他パネル */
        var panelOther = panelSplitGroup.add("panel", undefined, L("panelOther"));
        panelOther.margins = [15, 20, 15, 10];
        panelOther.alignment = ["fill", "top"];
        panelOther.alignChildren = ["fill", "center"];

        /* ソート（文字コード）ボタン */
        var btnSortByCharCode = panelOther.add("button", undefined, L("btnSortByCharCode"));
        btnSortByCharCode.onClick = function () {
            executeAction(sortByCharCode);
        };

        /* ソート（文字数）ボタン */
        var btnSortByLength = panelOther.add("button", undefined, L("btnSortByLength"));
        btnSortByLength.onClick = function () {
            executeAction(sortByLength);
        };

        /* 順序を反転ボタン */
        var btnReverseOrder = panelOther.add("button", undefined, L("btnReverseOrder"));
        btnReverseOrder.onClick = function () {
            executeAction(reverseOrder);
        };

        /* 重複行の削除ボタン */
        var btnRemoveDuplicateLines = panelOther.add("button", undefined, L("btnRemoveDuplicateLines"));
        btnRemoveDuplicateLines.onClick = function () {
            executeAction(removeDuplicateLines);
        };

        /* 連結パネル */
        var panelConcat = colCenter.add("panel", undefined, L("panelConcat"));
        panelConcat.margins = [15, 20, 15, 10];
        panelConcat.alignment = ["fill", "top"];
        panelConcat.alignChildren = ["fill", "center"];

        /* 連結（縦）ボタン */
        var btnConcatV = panelConcat.add("button", undefined, L("btnConcatV"));
        btnConcatV.helpTip = L("tipConcatV");
        btnConcatV.onClick = function () {
            executeAction(concatVertical);
        };

        /* 横連結（行維持）ボタン */
        var btnConcatHOnly = panelConcat.add("button", undefined, L("btnConcatHOnly"));
        btnConcatHOnly.helpTip = L("tipConcatHOnly");
        btnConcatHOnly.onClick = function () {
            executeAction(concatHorizontalOnly);
        };

        /* 横連結ボタン */
        var btnConcatH = panelConcat.add("button", undefined, L("btnConcatH"));
        btnConcatH.helpTip = L("tipConcatH");
        btnConcatH.onClick = function () {
            executeAction(function (objects) {
                return concatHorizontal(objects, detectTextFrameType(objects));
            });
        };

        /* クリーンアップパネル */
        var panelCleanup = colCenter.add("panel", undefined, L("panelCleanup"));
        panelCleanup.margins = [15, 20, 15, 10];
        panelCleanup.alignment = ["fill", "top"];
        panelCleanup.alignChildren = ["fill", "center"];

        /* 空行削除ボタン */
        var btnRemoveEmptyLines = panelCleanup.add("button", undefined, L("btnRemoveEmptyLines"));
        btnRemoveEmptyLines.onClick = function () {
            executeAction(removeEmptyLines);
        };

        /* タブパネル */
        var panelTab = panelCleanup.add("panel", undefined, L("panelTab"));
        panelTab.margins = [15, 20, 15, 10];
        panelTab.alignment = ["fill", "top"];
        panelTab.alignChildren = ["fill", "center"];

        /* タブを削除ボタン */
        var btnRemoveTabs = panelTab.add("button", undefined, L("btnRemoveTabs"));
        btnRemoveTabs.onClick = function () {
            executeAction(removeTabs);
        };

        /* タブをスペースにボタン */
        var btnTabsToSpaces = panelTab.add("button", undefined, L("btnTabsToSpaces"));
        btnTabsToSpaces.onClick = function () {
            executeAction(tabsToSpaces);
        };

        /* スペースパネル */
        var panelSpace = panelCleanup.add("panel", undefined, L("panelSpace"));
        panelSpace.margins = [15, 20, 15, 10];
        panelSpace.alignment = ["fill", "top"];
        panelSpace.alignChildren = ["fill", "center"];

        /* 行頭行末のスペースボタン */
        var btnTrimSpaces = panelSpace.add("button", undefined, L("btnTrimSpaces"));
        btnTrimSpaces.onClick = function () {
            executeAction(trimSpaces);
        };

        /* 和欧間のスペースボタン */
        var btnCjkLatinSpaces = panelSpace.add("button", undefined, L("btnCjkLatinSpaces"));
        btnCjkLatinSpaces.onClick = function () {
            executeAction(removeCjkLatinSpaces);
        };

        /* 連続スペースボタン */
        var btnCollapseSpaces = panelSpace.add("button", undefined, L("btnCollapseSpaces"));
        btnCollapseSpaces.onClick = function () {
            executeAction(collapseSpaces);
        };

        updateStatusDisplay(selectedObjects);
        updateActionAvailability(selectedObjects);

        /* ボタン行（戻す：左、閉じる：右） */
        var btnGroup = dialog.add("group");
        btnGroup.alignment = ["fill", "bottom"];
        btnGroup.margins = [0, 15, 0, 0];

        /* 直前の操作を取り消し、選択とステータスを再取得 */

        var btnUndo = btnGroup.add("button", undefined, L("btnUndo"));
        btnUndo.alignment = ["left", "center"];
        btnUndo.onClick = function () {
            try {
                app.executeMenuCommand('undo');
                app.redraw();
                var refreshed = getTextFrames(app.activeDocument.selection);
                if (refreshed.length > 0) {
                    selectedObjects = refreshed;
                    app.activeDocument.selection = refreshed;
                } else {
                    selectedObjects = [];
                }
                updateStatusDisplay(refreshed);
                updateActionAvailability(refreshed);
            } catch (e) {
                showError(e);
            }
        };

        var btnClose = btnGroup.add("button", undefined, L("btnClose"), { name: "ok" });
        btnClose.alignment = ["right", "center"];
        btnClose.onClick = function () {
            /* 1要素だけのグループは解除して選択を整える */
            try {
                var targets = getCurrentTargets();
                for (var i = 0; i < targets.length; i++) {
                    try {
                        var p = targets[i].parent;
                        if (p && p.typename === "GroupItem" && p.pageItems.length === 1) {
                            var container = p.parent;
                            targets[i].move(container, ElementPlacement.PLACEATEND);
                            p.remove();
                        }
                    } catch (e) { debugLog("btnClose: ungroup single-item group", e); }
                }
                if (targets.length > 0) {
                    app.activeDocument.selection = targets;
                }
            } catch (e) { debugLog("btnClose: finalize selection", e); }
            if (hiddenCharOn) {
                try { app.executeMenuCommand('showHiddenChar'); } catch (_) { }
            }
            dialog.close();
        };

        dialog.show();
    }

    try {
        showDialog(selectedObjects);
    } catch (err) {
        showError(err);
    }
})();