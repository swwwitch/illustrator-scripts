#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// TextBreakSplitMerge.jsx 
// 
// 作成日：2026年3月18日
// 更新日：2026年3月18日
//
// 選択したテキストフレームに対して、改行の整理、分割、連結を行うIllustrator用スクリプトです。
// 対象はテキストフレーム、テキストフレームを含むグループ、および Illustrator が TextRange として返す通常のテキスト選択です。
//
// ［改行］
// ・「改行削除」: 段落改行（\r）を削除して1つの連続テキストにまとめます。
// ・「空行削除」: 空行、または空白だけの行を削除します。
// ・「1文字ごとに改行」: 各文字の後ろに改行を挿入します（既存改行の直前直後は除く）。
// ・「句読点で改行」: 句読点や終止記号（、。,.!?！？ など）の後ろで改行します（次が既存改行でない場合）。
// ・「強制改行→改行」: 強制改行（charCode 3 / 10）を通常の段落改行（charCode 13）へ変換します。
//
// ［分割］
// ・「改行で分割」: 段落ごとにテキストフレームを複製し、空段落はスキップして分割します。
// ・「タブと改行」: 段落単位でタブ区切りに分解し、実測できるタブ位置を優先して個別のテキストフレームに分割します。
//
// ［連結］
// ・「縦」: 選択テキストを上から下へ連結します。複数行は行送りベースで並びを再構成し、最終的に1つのテキストへまとめます。
// ・「横連結」: 選択テキストを行単位で左から右へ連結します。英単語のハイフン結合や語間スペース補完を行い、必要に応じて改行を残します。
//
// ［ステータス表示］
// ・ダイアログ上部に、対象テキスト数、ポイント文字数、エリア内文字数、改行数、強制改行数、タブ数を表示します。
//
// ［その他］
// ・「連続実行」: モードパネル内に配置されています。UI上は残していますが、現在は動作を一時停止しています。チェック状態にかかわらず、各処理は選択中の元テキストに対して単独で実行されます。
// ・「制御文字」: Illustrator の制御文字表示を切り替えます。
// ・「キャンセル」: 現在はそのまま閉じます。
// ・「閉じる」: 現在の状態をそのまま確定して閉じます。
//
// ダイアログボックスのUIは日本語/英語を自動で切り替え、タイトルバーにはバージョン番号を表示します。
// なお、横連結は見た目ベースの近似処理であり、複雑な書式差や厳密な段落属性の保持は対象外です。
// Adobe Illustrator 2025に対応し、ExtendScript（ECMAScript 3）で記述されています。

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "テキスト処理",
        en: "Text Processing"
    },
    panelLineBreak: {
        ja: "改行",
        en: "Line Breaks"
    },
    panelSplit: {
        ja: "分割",
        en: "Split"
    },
    panelConcat: {
        ja: "連結",
        en: "Concatenate"
    },
    panelMode: {
        ja: "モード",
        en: "Mode"
    },
    panelStatus: {
        ja: "ステータス",
        en: "Status"
    },
    btnRemoveLineBreaks: {
        ja: "改行削除",
        en: "Remove Line Breaks"
    },
    btnRemoveEmptyLines: {
        ja: "空行削除",
        en: "Remove Empty Lines"
    },
    btnAddLineBreaks: {
        ja: "1文字ごとに改行",
        en: "Insert Line Break After Each Character"
    },
    btnPunctuation: {
        ja: "句読点で改行",
        en: "Insert Line Breaks at Punctuation"
    },
    btnConvertBreaks: {
        ja: "強制改行→改行",
        en: "Convert Forced Breaks to Paragraph Breaks"
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
        ja: "縦",
        en: "Vertical"
    },
    btnConcatH: {
        ja: "横連結",
        en: "Concatenate Horizontally"
    },
    chkStackActions: {
        ja: "連続実行",
        en: "Stack Actions"
    },
    chkShowHiddenChar: {
        ja: "制御文字",
        en: "Hidden Characters"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
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
    return LABELS[key][lang];
}

/* エラー表示補助 / Error display helper */
function showError(err) {
    var msg = L("errProcessFailed");
    if (err && err.message) {
        msg += err.message;
    } else {
        msg += String(err);
    }
    alert(msg);
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

    /* ユーティリティ関数 / Utility functions */

    /* テキストフレームのみ抽出 / Collect text frames only */
    function getTextFrames(objects) {
        var frames = [];

        function pushUnique(frame) {
            if (!frame) return;
            try {
                if (frame.isValid === false) return;
            } catch (e) { return; }

            for (var i = 0; i < frames.length; i++) {
                if (frames[i] === frame) return;
            }
            frames.push(frame);
        }

        function collect(item) {
            if (!item) return;

            try {
                if (item.isValid === false) return;
            } catch (e) { return; }

            var typeName = "";
            try {
                typeName = item.typename || "";
            } catch (e2) { return; }

            if (typeName === "TextFrame") {
                pushUnique(item);
                return;
            }

            /* Illustrator ではテキスト選択が TextRange として返ることがある / Illustrator may return a text selection as TextRange */
            if (typeName === "TextRange") {
                try {
                    if (item.parent && item.parent.typename === "TextFrame") {
                        pushUnique(item.parent);
                        return;
                    }
                } catch (e3) { }
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
        } catch (e5) { }

        return frames;
    }

    /* 初期選択からテキストフレームを解決 / Resolve text frames from the initial selection */
    function resolveInitialTextFrames(selection) {
        return getTextFrames(selection);
    }

    /* テキストタイプ判定 / Detect text frame type */
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

    /* テキストタイプ件数を集計 / Count text frame types */
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

    /* 改行数とタブ数を集計 / Count paragraph breaks, forced breaks, and tabs */
    function countBreakTypes(objects) {
        var paragraphBreakCount = 0;
        var forcedBreakCount = 0;
        var tabCount = 0;
        var frames = getTextFrames(objects);

        for (var i = 0; i < frames.length; i++) {
            var chars = frames[i].characters;
            for (var j = 0; j < chars.length; j++) {
                var code = chars[j].contents.charCodeAt(0);
                if (code === 13) {
                    paragraphBreakCount++;
                } else if (code === 3 || code === 10) {
                    forcedBreakCount++;
                } else if (code === 9) {
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

    /* 各テキストフレームのcontentsを変換する共通処理 / Shared process to transform each text frame contents */
    function transformContents(objects, transformFunc) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            frames[i].contents = transformFunc(frames[i].contents);
        }
    }

    /* 強制改行（charCode 3 または 10）を削除する共通処理 / Shared process to remove forced line breaks (charCode 3 or 10) */
    function removeForcedLineBreaks(frames) {
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (charCode === 3 || charCode === 10) {
                    frames[i].characters[c].remove();
                }
            }
        }
    }

    /* 上から順にソート（Y降順、同じYならX昇順） / Sort from top to bottom (Y descending, then X ascending) */
    function sortByPosition(items) {
        var arr = items.slice(0);
        arr.sort(function (a, b) {
            if (b.position[1] !== a.position[1]) return b.position[1] - a.position[1];
            return a.position[0] - b.position[0];
        });
        return arr;
    }

    /* Y座標で降順ソート / Sort by Y in descending order */
    function sortByY(items) {
        var arr = items.slice(0);
        arr.sort(function (a, b) { return b.position[1] - a.position[1]; });
        return arr;
    }

    /* X座標で昇順ソート / Sort by X in ascending order */
    function sortByX(items) {
        var arr = items.slice(0);
        arr.sort(function (a, b) { return a.position[0] - b.position[0]; });
        return arr;
    }

    /* Y位置で行グループ化 / Group into lines by Y position */
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

    /* 選択範囲のバウンディングボックス取得 / Get selection bounding box */
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

    /* テキストフレームをグループ化 / Group text frames */
    function groupTextFrames(frames, targetLayer) {
        var validFrames = getTextFrames(frames);
        if (validFrames.length === 0) return [];

        var layer = targetLayer || app.activeDocument.activeLayer;
        var grp = layer.groupItems.add();

        for (var i = 0; i < validFrames.length; i++) {
            try {
                validFrames[i].move(grp, ElementPlacement.PLACEATEND);
            } catch (e) { }
        }

        return [grp];
    }

    /* 改行系の関数 / Line break related functions */

    /* 改行文字を削除する関数 / Remove line break characters */
    function removeLineBreaks(objects) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (charCode === 13) {
                    frames[i].characters[c].remove();
                }
            }
        }
    }

    /* 空行を削除する関数 / Remove empty lines */
    function removeEmptyLines(objects) {
        transformContents(objects, function (txt) {
            var normalized = txt.replace(/\r\n/g, "\r").replace(/\n/g, "\r");
            var lines = normalized.split("\r");
            var kept = [];
            for (var i = 0; i < lines.length; i++) {
                if (lines[i].replace(/[ \t　]/g, "") !== "") {
                    kept.push(lines[i]);
                }
            }
            return kept.join("\r");
        });
    }

    /* 1文字ごとに改行を挿入する関数 / Insert a line break after each character */
    function addLineBreakPerChar(objects) {
        transformContents(objects, function (txt) {
            var newTxt = "";
            for (var j = 0; j < txt.length; j++) {
                var c = txt.charAt(j);
                newTxt += c;
                if (c !== '\r' && c !== '\n' && j < txt.length - 1) {
                    var nextC = txt.charAt(j + 1);
                    if (nextC !== '\r' && nextC !== '\n') {
                        newTxt += '\r';
                    }
                }
            }
            return newTxt;
        });
    }

    /* 強制改行を通常の改行に変換する関数 / Convert forced line breaks to normal paragraph breaks */
    function convertForcedLineBreaks(objects) {
        var frames = getTextFrames(objects);
        for (var i = 0; i < frames.length; i++) {
            for (var c = frames[i].characters.length - 1; c >= 0; c--) {
                var charCode = frames[i].characters[c].contents.charCodeAt(0);
                if (charCode === 3 || charCode === 10) {
                    frames[i].characters[c].contents = String.fromCharCode(13);
                }
            }
        }
    }

    /* 句読点の後に改行を挿入する関数 / Insert line breaks after punctuation */
    function addLineBreakAtPunctuation(objects) {
        /* 対象記号：和文/欧文の句読点と終止記号 / Target marks: Japanese/Western commas, periods, and sentence-ending punctuation */
        var punctuation = "、。，．｡､,.!?！？";
        transformContents(objects, function (txt) {
            var newTxt = "";
            for (var j = 0; j < txt.length; j++) {
                var c = txt.charAt(j);
                newTxt += c;
                if (punctuation.indexOf(c) !== -1 && j < txt.length - 1) {
                    var remaining = txt.substring(j + 1);
                    if (remaining.replace(/[\r\n]/g, "").length > 0) {
                        var nextC = txt.charAt(j + 1);
                        if (nextC !== '\r' && nextC !== '\n') {
                            newTxt += '\r';
                        }
                    }
                }
            }
            return newTxt;
        });
    }

    /* タブで分解する関数 / Split by tabs */
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

                    /* 段落単位でタブ位置を取得 / Collect tab positions per paragraph */
                    var tabCharPositions = [];
                    var paraChars = para.characters;
                    for (var ci = 0; ci < paraChars.length; ci++) {
                        if (paraChars[ci].contents === "\t") {
                            if (ci + 1 < paraChars.length) {
                                var nextCharBounds = paraChars[ci + 1].visibleBounds;
                                tabCharPositions.push(nextCharBounds[0]); /* 左端X座標 / Left X coordinate */
                            }
                        }
                    }

                    var contAry = para.contents.split("\t");
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

    /* 改行で分割する関数 / Split by line breaks */
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

                    var paraText = para.contents.replace(/[\r\n]+$/g, "");
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

    /* --- concatHorizontal specification moved below --- */

    /*
     * 連結（縦）する関数 / Concatenate vertically
     *
     * 仕様 / Specification:
     * - 選択されたテキストフレームを上から下、同じ高さでは左から右の順に処理します。
     *   Process selected text frames from top to bottom, and left to right when Y positions are equal.
     * - 各フレーム内の複数行は、改行単位で分解し、元フレームの行送りを使って仮のY位置を再構成します。
     *   Multi-line contents are split by paragraph breaks, and temporary Y positions are reconstructed using each source frame's leading.
     * - 連結順は仮配置後に再ソートして決定します。
     *   Final merge order is determined by sorting the temporary line items.
     * - 出力は最上段要素をベースにし、段落オブジェクト操作ではなく、各行を `\r` で連結した単一テキストへ再構成します。
     *   Output is rebuilt as a single text object based on the topmost item by joining lines with `\r`, instead of using fragile paragraph-object operations.
     * - 複雑な段落属性や行ごとの差分書式は厳密には保持しません。
     *   Complex paragraph attributes and per-line formatting differences are not preserved exactly.
     */
    function concatVertical(objects) {
        var textFrames = getTextFrames(objects);
        if (textFrames.length < 2) return textFrames;

        textFrames = sortByPosition(textFrames);

        /* 各テキストフレームを行単位に分解して再ソート / Split each text frame into lines and sort again */
        var splitFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            var sourceFrame = textFrames[i];
            var normalizedText = sourceFrame.contents.replace(/\r\n/g, "\r").replace(/\n/g, "\r");
            var lines = normalizedText.split('\r');
            var basePos = sourceFrame.position;
            var baseX = basePos[0];
            var baseY = basePos[1];
            var baseSize = sourceFrame.textRange.characterAttributes.size;
            var baseLeading = sourceFrame.textRange.characterAttributes.leading;
            if (!baseLeading || baseLeading === 0) {
                baseLeading = baseSize * 1.2;
            }

            for (var j = 0; j < lines.length; j++) {
                var lineText = lines[j].replace(/[\r\n]+$/g, "");
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

        /* 最上段のフレームをベースに文字列で再構成 / Rebuild as plain text using the topmost frame as the base */
        var mergedLines = [];
        for (var k = 0; k < splitFrames.length; k++) {
            var mergedLineText = splitFrames[k].contents.replace(/[\r\n]+$/g, "");
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
     * 連結（横：書式統一）する関数 / Concatenate horizontally and unify formatting
     *
     * 仕様 / Specification:
     * - 選択されたテキストフレームをY位置ベースで行グループ化し、各行を左から右へ連結します。
     *   Group selected text frames into lines by Y position, then concatenate each line from left to right.
     * - 処理前に、強制改行（charCode 3 / 10）は削除します。
     *   Forced line breaks (charCode 3 / 10) are removed before concatenation.
     * - 1行のみの場合は、その行の全テキストを連結して1つのテキストオブジェクトとして出力します。
     *   If only one line exists, all text in that line is merged into a single output text object.
     * - 複数行の場合は、まず各行を個別に連結し、その後、行間の連結ルールに従って1つのテキストへまとめます。
     *   If multiple lines exist, each line is merged first, then all merged lines are combined into one text object.
     * - 行末が英単語のハイフン区切りで、次行頭も英数字なら、ハイフンを削除して連結します。
     *   If a line ends with a hyphenated English word and the next line starts with an English letter/number, the hyphen is removed.
     * - 行末と次行頭がともに英数字系なら、語間スペースを補います。
     *   If the end of one line and the start of the next are both English word characters, a space is inserted.
     * - 行末が句点・終止記号の場合のみ、行間に改行を残します。
     *   A paragraph break is preserved between lines only when the line ends with sentence-ending punctuation.
     * - textMode が "point" の場合はポイント文字で出力し、それ以外はエリア内文字で出力します。
     *   When textMode is "point", the result is created as point text; otherwise it is created as area text.
     * - textMode が未指定または "mixed" の場合は既存選択から自動判定し、混在時はエリア内文字を優先します。
     *   When textMode is omitted or "mixed", the mode is auto-detected from the selection; mixed selections fall back to area text.
     * - 出力書式は先頭要素のフォント・サイズを基準に再設定します。
     *   Output formatting is rebuilt using the font and size of the first merged item.
     * - 行送りは、連結後の各行のY差分をもとに再設定します。
     *   Leading is recalculated from the Y-distance between merged lines.
     * - これは見た目ベースの近似連結であり、複雑な書式差・回転・厳密な段落属性までは保持しません。
     *   This is a visually approximated merge and does not fully preserve complex formatting differences, rotation, or exact paragraph attributes.
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

        /* 強制改行を削除 / Remove forced line breaks */
        removeForcedLineBreaks(textItems);

        /* Y位置で行ごとにグループ化 / Group text items into lines by Y position */
        var textLines = groupByLineY(
            sortByY(textItems), LINE_Y_THRESHOLD
        );

        /* 1行だけの場合：X順に連結して出力 / If there is only one line, concatenate in X order and output */
        if (textLines.length === 1) {
            var sorted = sortByX(textLines[0]);
            var mergedText = "";
            var resultFrame;
            for (var j = 0; j < sorted.length; j++) {
                mergedText += sorted[j].contents;
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

        /* 複数行：行ごとにX順で連結 / For multiple lines, concatenate each line in X order */
        var mergedFrames = [];
        for (var i = 0; i < textLines.length; i++) {
            var sorted = sortByX(textLines[i]);
            var mergedText = "";
            for (var j = 0; j < sorted.length; j++) {
                mergedText += sorted[j].contents;
            }
            var mf = sorted[0].duplicate();
            mf.contents = mergedText;
            mf.position = sorted[0].position;
            mergedFrames.push(mf);
            for (var k = 0; k < sorted.length; k++) {
                sorted[k].remove();
            }
        }

        /* 行間の連結処理（句点・ピリオド等で終わる場合のみ改行） / Merge between lines and insert breaks only when ending with punctuation */
        var finalText = "";
        var fontSize = mergedFrames[0].textRange.characterAttributes.size;

        for (var i = 0; i < mergedFrames.length; i++) {
            var content = mergedFrames[i].contents;
            finalText += content;

            if (i < mergedFrames.length - 1) {
                var nextContent = mergedFrames[i + 1].contents;

                /* 英単語がハイフンで分断されている場合、ハイフンを除去して結合 / Remove trailing hyphen when an English word is split across lines */
                if (/[A-Za-z0-9)]-$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                    finalText = finalText.replace(/-$/, "");
                }
                /* 末尾が英単語、次の行頭も英単語ならスペース挿入 / Insert a space when both adjacent line ends are English word characters */
                else if (/[A-Za-z0-9)]$/.test(content) && /^[A-Za-z0-9(]/.test(nextContent)) {
                    finalText += " ";
                }

                /* 句点等で終わる場合は改行 / Insert a line break when ending with sentence punctuation */
                var endsWithJP = /[。！？]$/.test(content);
                var endsWithEN = /[.!?]$/.test(content);
                var isEnglish = /^[\x00-\x7F]+$/.test(content.replace(/[\s\r\n]/g, ""));
                if (endsWithJP || (endsWithEN && !isEnglish)) {
                    finalText += "\r";
                }
            }
        }

        /* 選択オブジェクトの幅を取得してエリアテキストに反映 / Use the selection width and height for the area text */
        var bounds = getSelBounds(mergedFrames);
        var selWidth = bounds[2] - bounds[0];
        var selHeight = bounds[1] - bounds[3];

        var newTF;

        if (textMode === "point") {
            /* ポイントテキストを作成 / Create point text */
            newTF = app.activeDocument.textFrames.pointText([bounds[0], bounds[1]]);
            newTF.contents = finalText;
            newTF.textRange.characterAttributes.textFont = mergedFrames[0].textRange.characterAttributes.textFont;
            newTF.textRange.characterAttributes.size = fontSize;
            newTF.paragraphs[0].paragraphAttributes.kinsoku = "Soft";
            newTF.textRange.justification = Justification.LEFT;
        } else {
            /* エリアテキストを作成 / Create area text */
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

        /* 行間を設定 / Set leading */
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

        /* 元のフレームを削除 / Remove original frames */
        for (var i = 0; i < mergedFrames.length; i++) {
            mergedFrames[i].remove();
        }

        app.redraw();
        return [newTF];
    }

    /* ダイアログボックスを作成・表示する関数 / Create and show the dialog box */
    function showDialog(selectedObjects) {
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);

        /* 連続実行セッション管理（現在は停止中） / Stack-action session (currently disabled) */
        var currentAction = null;
        var stackSession = {
            started: false,
            workLayer: null,
            originalItems: [],
            originalLayers: [],
            workingItems: [],
            previousActiveLayer: null
        };

        /* 連続実行セッション開始（現在は停止中） / Start stack-action session (currently disabled) */
        function ensureStackSession() {
            return;
        }

        /* 作業レイヤー上のテキストを再取得（現在は停止中） / Refresh working-layer text frames (currently disabled) */
        function refreshWorkingItems() {
            stackSession.workingItems = [];
            return [];
        }

        /* 現在の処理対象を取得 / Get current action targets */
        function getCurrentTargets() {
            return getTextFrames(selectedObjects);
        }

        /* 連続実行セッションを破棄（現在は停止中） / Discard stack-action session (currently disabled) */
        function discardStackSession() {
            return;
        }

        /* 連続実行セッションを確定（現在は停止中） / Commit stack-action session (currently disabled) */
        function commitStackSession() {
            return;
        }

        /* 処理を実行する関数 / Execute an action */
        function executeAction(actionFunc) {
            currentAction = actionFunc;
            try {
                var targets = getCurrentTargets();
                if (!targets || targets.length === 0) return;

                var result = actionFunc(targets);
                var refreshedTargets = getTextFrames(result && result.length ? result : targets);

                if (refreshedTargets.length > 0) {
                    selectedObjects = refreshedTargets;
                    app.activeDocument.selection = refreshedTargets;
                    updateStatusDisplay(refreshedTargets);
                } else {
                    updateStatusDisplay([]);
                }

                app.redraw();
            } catch (err) {
                try { app.redraw(); } catch (redrawErr) { }
                showError(err);
            }
        }

        /* 処理対象件数を集計 / Count target text frames */
        var textFrameCounts = countTextFrameTypes(selectedObjects);
        var breakCounts = countBreakTypes(selectedObjects);

        /* ステータスパネル / Status panel */
        var panelStatus = dialog.add("panel", undefined, L("panelStatus"));
        panelStatus.margins = [15, 20, 15, 10];
        panelStatus.alignment = ["fill", "top"];
        panelStatus.alignChildren = ["left", "top"];

        /* ステータス表示 / Status display */
        var statusGroup = panelStatus.add("group");
        statusGroup.orientation = "row";
        statusGroup.alignment = ["fill", "top"];
        statusGroup.alignChildren = ["left", "top"];
        statusGroup.spacing = 30;

        var statusLeft = statusGroup.add("group");
        statusLeft.orientation = "column";
        statusLeft.alignChildren = ["left", "top"];

        var statusRight = statusGroup.add("group");
        statusRight.orientation = "column";
        statusRight.alignChildren = ["left", "top"];

        var statusLabelWidth = 100;

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

        var statusRightLabelWidth = 60;

        var rowParagraphBreakCount = statusRight.add("group");
        rowParagraphBreakCount.orientation = "row";
        rowParagraphBreakCount.alignChildren = ["left", "center"];
        var lblParagraphBreakCount = rowParagraphBreakCount.add("statictext", undefined, L("infoParagraphBreakCount"), { justify: "right" });
        lblParagraphBreakCount.preferredSize.width = statusRightLabelWidth;
        var valParagraphBreakCount = rowParagraphBreakCount.add("statictext", undefined, String(breakCounts.paragraph));

        var rowForcedBreakCount = statusRight.add("group");
        rowForcedBreakCount.orientation = "row";
        rowForcedBreakCount.alignChildren = ["left", "center"];
        var lblForcedBreakCount = rowForcedBreakCount.add("statictext", undefined, L("infoForcedBreakCount"), { justify: "right" });
        lblForcedBreakCount.preferredSize.width = statusRightLabelWidth;
        var valForcedBreakCount = rowForcedBreakCount.add("statictext", undefined, String(breakCounts.forced));

        var rowTabCount = statusRight.add("group");
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

        /* ステータス表示を更新 / Refresh status display */
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

        updateStatusDisplay(selectedObjects);

        /* モードパネル / Mode panel */
        var panelMode = dialog.add("panel", undefined, L("panelMode"));
        panelMode.margins = [15, 20, 15, 10];
        panelMode.alignment = ["fill", "top"];
        panelMode.alignChildren = ["left", "top"];

        /* モード行 / Mode row */
        var modeRow = panelMode.add("group");
        modeRow.orientation = "row";
        modeRow.alignment = ["center", "top"];
        modeRow.alignChildren = ["left", "center"];
        modeRow.spacing = 20;

        /* 連続実行チェックボックス / Stackable execution checkbox */
        var chkStackActions = modeRow.add("checkbox", undefined, L("chkStackActions"));
        chkStackActions.value = true;
        chkStackActions.enabled = false;

        /* 制御文字チェックボックス / Hidden characters checkbox */
        var chkShowHiddenChar = modeRow.add("checkbox", undefined, L("chkShowHiddenChar"));
        chkShowHiddenChar.value = false;
        chkShowHiddenChar.onClick = function () {
            try {
                app.executeMenuCommand('showHiddenChar');
                app.redraw();
            } catch (err) {
                showError(err);
            }
        };

        /* 2カラムレイアウト / Two-column layout */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignment = ["fill", "top"];
        columnsGroup.alignChildren = ["fill", "top"];

        /* 左カラム / Left column */
        var colLeft = columnsGroup.add("group");
        colLeft.orientation = "column";
        colLeft.alignChildren = ["fill", "top"];

        /* 改行パネル / Line break panel */
        var panelLineBreak = colLeft.add("panel", undefined, L("panelLineBreak"));
        panelLineBreak.margins = [15, 20, 15, 10];
        panelLineBreak.alignment = ["fill", "top"];
        panelLineBreak.alignChildren = ["fill", "center"];

        /* 改行削除ボタン / Remove line breaks button */
        var btnRemoveLineBreaks = panelLineBreak.add("button", undefined, L("btnRemoveLineBreaks"));
        btnRemoveLineBreaks.onClick = function () {
            executeAction(removeLineBreaks);
        };

        /* 空行削除ボタン / Remove empty lines button */
        var btnRemoveEmptyLines = panelLineBreak.add("button", undefined, L("btnRemoveEmptyLines"));
        btnRemoveEmptyLines.onClick = function () {
            executeAction(removeEmptyLines);
        };

        /* 1文字ごとに改行ボタン / Insert line break after each character button */
        var btnAddLineBreaks = panelLineBreak.add("button", undefined, L("btnAddLineBreaks"));
        btnAddLineBreaks.onClick = function () {
            executeAction(addLineBreakPerChar);
        };

        /* 句読点で改行ボタン / Insert line breaks at punctuation button */
        var btnPunctuation = panelLineBreak.add("button", undefined, L("btnPunctuation"));
        btnPunctuation.onClick = function () {
            executeAction(addLineBreakAtPunctuation);
        };

        /* 強制改行→改行ボタン / Convert forced breaks button */
        var btnConvertBreaks = panelLineBreak.add("button", undefined, L("btnConvertBreaks"));
        btnConvertBreaks.onClick = function () {
            executeAction(convertForcedLineBreaks);
        };

        /* 右カラム / Right column */
        var colRight = columnsGroup.add("group");
        colRight.orientation = "column";
        colRight.alignChildren = ["fill", "top"];

        /* 分割パネル / Split panel */
        var panelSplit = colRight.add("panel", undefined, L("panelSplit"));
        panelSplit.margins = [15, 20, 15, 10];
        panelSplit.alignment = ["fill", "top"];
        panelSplit.alignChildren = ["fill", "center"];

        /* 改行で分割ボタン / Split by line breaks button */
        var btnSplitByLine = panelSplit.add("button", undefined, L("btnSplitByLine"));
        btnSplitByLine.onClick = function () {
            executeAction(splitByLineBreak);
        };

        /* タブで分解ボタン / Split by tabs and line breaks button */
        var btnSplitByTab = panelSplit.add("button", undefined, L("btnSplitByTab"));
        btnSplitByTab.onClick = function () {
            executeAction(splitByTab);
        };

        /* 連結パネル / Concatenate panel */
        var panelConcat = colRight.add("panel", undefined, L("panelConcat"));
        panelConcat.margins = [15, 20, 15, 10];
        panelConcat.alignment = ["fill", "top"];
        panelConcat.alignChildren = ["fill", "center"];

        /* 連結（縦）ボタン / Vertical concatenate button */
        var btnConcatV = panelConcat.add("button", undefined, L("btnConcatV"));
        btnConcatV.onClick = function () {
            executeAction(concatVertical);
        };

        /* 横連結ボタン / Horizontal concatenate button */
        var btnConcatH = panelConcat.add("button", undefined, L("btnConcatH"));
        btnConcatH.onClick = function () {
            executeAction(function (objects) {
                return concatHorizontal(objects, detectTextFrameType(objects));
            });
        };


        /* ボタン行（キャンセル：左、閉じる：右） / Button row (Cancel left, Close right) */
        var btnGroup = dialog.add("group");
        btnGroup.alignment = ["fill", "bottom"];

        var btnCancel = btnGroup.add("button", undefined, L("btnCancel"));
        btnCancel.alignment = ["left", "center"];
        btnCancel.onClick = function () {
            dialog.close();
        };

        var btnClose = btnGroup.add("button", undefined, L("btnClose"));
        btnClose.alignment = ["right", "center"];
        btnClose.onClick = function () {
            /* グループ内オブジェクトが1つだけならグループ解除 / Ungroup single-item groups */
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
                    } catch (e) { }
                }
                if (targets.length > 0) {
                    app.activeDocument.selection = targets;
                }
            } catch (e) { }
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