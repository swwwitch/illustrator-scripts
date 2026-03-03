#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：
DetachPathText.jsx

### 更新日：
20260303

### 概要：
選択した「パス上文字（Path Text）」を、同じ文字内容・段落属性・文字属性をできるだけ維持したまま「ポイント文字」に変換します。
変換時に、「テキストの書式」およびパス設定を選択するダイアログが表示されます。
元のテキストパス形状は複製され、ダイアログの設定に応じて線属性（1pt黒／線なし）が適用されます。

### 使い方：
1) パス上文字を選択
2) スクリプトを実行
3) ダイアログでテキストの書式およびパスのオプションを選択

### 注意：
- 文字ごとの属性（フォント/サイズ/色/トラッキング等）は、可能な範囲で復元します。
- 復元時に設定できない属性は try/catch でスキップします。
*/

/* バージョン / Version */
var SCRIPT_VERSION = "v1.0.6";

/* 言語判定 / Language detection */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "パス上文字の解除", en: "Detach Path Text" },
    pnlKeep: { ja: "テキストの書式", en: "Text Formatting" },
    rbFull: { ja: "完全に保持", en: "Preserve Completely" },
    rbFast: { ja: "高速に保持", en: "Preserve Quickly" },
    rbNone: { ja: "削除", en: "Remove" },
    pnlPath: { ja: "パス", en: "Path" },
    rbPathBlack: { ja: "1pt黒に設定", en: "Set 1pt Black Stroke" },
    rbPathNone: { ja: "線なし", en: "No Stroke" },
    rbPathDelete: { ja: "削除", en: "Delete" },
    btnCancel: { ja: "閉じる", en: "Close" },
    btnOk: { ja: "OK", en: "OK" },
    alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoSel: { ja: "パス上文字を選択してください。", en: "Please select path text." },
    alertNoPath: { ja: "選択範囲にパス上文字が含まれていません。", en: "No path text found in selection." }
};

/* ラベル取得ヘルパー / Label lookup helper */
function L(key) {
    return LABELS[key][lang];
}

// ==========================================
// UI: ダイアログ / Dialog
// ==========================================
function showOptionsDialog() {
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 18;

    /* テキストパネル / Text panel */
    var pnlKeep = dlg.add("panel", undefined, L("pnlKeep"));
    pnlKeep.orientation = "column";
    pnlKeep.alignChildren = ["left", "top"];
    pnlKeep.margins = [15, 20, 15, 12];

    var rbFull = pnlKeep.add("radiobutton", undefined, L("rbFull"));
    var rbFast = pnlKeep.add("radiobutton", undefined, L("rbFast"));
    var rbNone = pnlKeep.add("radiobutton", undefined, L("rbNone"));
    rbFull.value = true;

    /* パスパネル / Path panel */
    var pnlPath = dlg.add("panel", undefined, L("pnlPath"));
    pnlPath.orientation = "column";
    pnlPath.alignChildren = ["left", "top"];
    pnlPath.margins = [15, 20, 15, 12];

    var rbPathBlack = pnlPath.add("radiobutton", undefined, L("rbPathBlack"));
    var rbPathNone = pnlPath.add("radiobutton", undefined, L("rbPathNone"));
    var rbPathDelete = pnlPath.add("radiobutton", undefined, L("rbPathDelete"));
    rbPathBlack.value = true; // デフォルト / Default

    /* ボタングループ / Button group */
    var grpBtns = dlg.add("group");
    grpBtns.orientation = "row";
    grpBtns.alignChildren = ["right", "center"];
    grpBtns.alignment = ["fill", "top"];

    var btnCancel = grpBtns.add("button", undefined, L("btnCancel"));
    var btnOk = grpBtns.add("button", undefined, L("btnOk"), { name: "ok" });

    var result = null;
    btnOk.onClick = function () {
        result = {
            mode: rbNone.value ? "none" : (rbFull.value ? "full" : "fast"),
            pathBlack1pt: rbPathBlack.value,
            pathNoStroke: rbPathNone.value,
            pathDelete: rbPathDelete.value
        };
        dlg.close(1);
    };
    btnCancel.onClick = function () {
        dlg.close(0);
    };

    var r = dlg.show();
    return (r === 1) ? result : null;
}

function main() {
    if (app.documents.length === 0) {
        alert(L("alertNoDoc"));
        return;
    }

    var opt = showOptionsDialog();
    if (!opt) return; // キャンセル / Cancelled

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert(L("alertNoSel"));
        return;
    }

    var pathTexts = [];

    /* 選択からパス上文字を抽出 / Extract path text from selection */
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "TextFrame" && sel[i].kind === TextType.PATHTEXT) {
            pathTexts.push(sel[i]);
        }
    }

    if (pathTexts.length === 0) {
        alert(L("alertNoPath"));
        return;
    }

    /* パス上文字の変換処理 / Convert path text items */
    for (var j = pathTexts.length - 1; j >= 0; j--) {
        var pathText = pathTexts[j];
        var originalPath = pathText.textPath;

        // ==========================================
        // 1. テキスト属性のスナップショットを取得 / Snapshot text attributes
        // ==========================================

        // full: 文字ごとの属性まで保持 / full: preserve per-character attributes
        // fast: textRange.duplicate() で高速に保持 / fast: duplicate textRange quickly
        // none: 書式を削除（デフォルト書式） / none: remove formatting (use defaults)
        var charAttrs = null;
        var firstCharAttr = null;
        if (opt.mode === "full") {
            charAttrs = [];
            for (var c = 0; c < pathText.characters.length; c++) {
                var ch = pathText.characters[c];
                var ca = ch.characterAttributes;
                charAttrs.push({
                    font: ca.textFont,
                    size: ca.size,
                    fillColor: ca.fillColor,
                    strokeColor: ca.strokeColor,
                    strokeWeight: ca.strokeWeight,
                    tracking: ca.tracking,
                    baselineShift: ca.baselineShift,
                    horizontalScale: ca.horizontalScale,
                    verticalScale: ca.verticalScale,
                    autoLeading: ca.autoLeading,
                    leading: ca.leading
                });
            }
        } else if (opt.mode === "fast") {
            // fast（代表値スナップショット：将来の拡張用 / representative snapshot: for future use）
            if (pathText.characters.length > 0) {
                var ca0 = pathText.characters[0].characterAttributes;
                firstCharAttr = {
                    font: ca0.textFont,
                    size: ca0.size,
                    fillColor: ca0.fillColor,
                    strokeColor: ca0.strokeColor,
                    strokeWeight: ca0.strokeWeight,
                    tracking: ca0.tracking,
                    baselineShift: ca0.baselineShift,
                    horizontalScale: ca0.horizontalScale,
                    verticalScale: ca0.verticalScale,
                    autoLeading: ca0.autoLeading,
                    leading: ca0.leading
                };
            }
        }

        /* 段落属性（先頭段落から取得）/ Paragraph attributes (from first paragraph) */
        var justification = null;
        if (pathText.paragraphs.length > 0) {
            var paraAttr = pathText.paragraphs[0].paragraphAttributes;
            justification = paraAttr.justification;
        }

        var textContents = pathText.contents; // 文字列を保存 / Save string contents

        // ==========================================
        // 2. パスを複製 / Duplicate path
        // ==========================================
        if (!opt.pathDelete) {
            var newPath = doc.pathItems.add();

            for (var k = 0; k < originalPath.pathPoints.length; k++) {
                var origPt = originalPath.pathPoints[k];
                var newPt = newPath.pathPoints.add();
                newPt.anchor = origPt.anchor;
                newPt.leftDirection = origPt.leftDirection;
                newPt.rightDirection = origPt.rightDirection;
                newPt.pointType = origPt.pointType;
            }
            newPath.closed = originalPath.closed;
            newPath.filled = false;

            if (opt.pathNoStroke) {
                newPath.stroked = false;
            } else {
                newPath.stroked = true;

                if (opt.pathBlack1pt) {
                    var blackColor = new CMYKColor();
                    blackColor.cyan = 0;
                    blackColor.magenta = 0;
                    blackColor.yellow = 0;
                    blackColor.black = 100;

                    newPath.strokeColor = blackColor;
                    newPath.strokeWidth = 1;
                }
            }
        }

        // ==========================================
        // 3. 新しいポイントテキストを作成し属性を復元 / Create new point text and restore attributes
        // ==========================================
        var newText = doc.textFrames.add();

        /* 位置をパス開始点に合わせる / Align position to path start point */
        var anchorPoint = originalPath.pathPoints[0].anchor;
        newText.position = [anchorPoint[0], anchorPoint[1]];

        if (opt.mode === "fast") {
            // fast: textRange を丸ごと複製（高速）/ fast: duplicate entire textRange
            // ※内容・書式（段落/文字）をまとめて複製する / ※duplicates content and formatting together
            pathText.textRange.duplicate(newText);

            /* 文字色を確実に引き継ぐ / Ensure character colors are carried over */
            for (var rr = 0; rr < pathText.textRanges.length; rr++) {
                try {
                    var srcAttr = pathText.textRanges[rr].characterAttributes;
                    var dstAttr = newText.textRanges[rr].characterAttributes;
                    dstAttr.fillColor = srcAttr.fillColor;
                    dstAttr.strokeColor = srcAttr.strokeColor;
                } catch (eRR) { }
            }
        } else if (opt.mode === "none") {
            // none: 内容だけ流し込み / none: insert content only, use default formatting
            newText.contents = textContents;
        } else {
            // full: 内容を流し込み、文字ごとの属性を復元 / full: insert content and restore per-character attributes
            newText.contents = textContents;

            /* 段落属性の復元 / Restore paragraph attributes */
            if (justification !== null && newText.paragraphs.length > 0) {
                newText.paragraphs[0].paragraphAttributes.justification = justification;
            }

            // ★重要：新規テキストのデフォルト線を先に消す / ★Important: clear default stroke on new text first
            // 後続の文字ごとの復元で必要な場合のみ線が復元される / stroke is only restored if the original had one
            try {
                var nc = new NoColor();
                newText.textRange.characterAttributes.strokeColor = nc;
                newText.textRange.characterAttributes.strokeWeight = 0;
            } catch (eClr) { }
        }

        /* 文字属性の復元（full モードのみ）/ Restore character attributes (full mode only) */
        if (opt.mode === "full" && charAttrs) {
            for (var c = 0; c < newText.characters.length; c++) {
                var targetCa = newText.characters[c].characterAttributes;
                var srcCa = charAttrs[c];
                if (!srcCa) break;

                try { targetCa.textFont = srcCa.font; } catch (e) { }
                try { targetCa.size = srcCa.size; } catch (e) { }
                try { targetCa.fillColor = srcCa.fillColor; } catch (e) { }
                /* stroke: 元が「なし」の場合は strokeWeight を 0 に / stroke: set weight to 0 if original had no stroke */
                try {
                    var sc = srcCa.strokeColor;
                    targetCa.strokeColor = sc;
                    if (sc && sc.typename === "NoColor") {
                        targetCa.strokeWeight = 0;
                    } else {
                        targetCa.strokeWeight = srcCa.strokeWeight;
                    }
                } catch (e) { }
                try { targetCa.tracking = srcCa.tracking; } catch (e) { }
                try { targetCa.baselineShift = srcCa.baselineShift; } catch (e) { }
                try { targetCa.horizontalScale = srcCa.horizontalScale; } catch (e) { }
                try { targetCa.verticalScale = srcCa.verticalScale; } catch (e) { }
                try { targetCa.autoLeading = srcCa.autoLeading; } catch (e) { }

                /* autoLeading が false の場合のみ leading を設定 / Set leading only when autoLeading is false */
                if (!srcCa.autoLeading) {
                    try { targetCa.leading = srcCa.leading; } catch (e) { }
                }
            }
        }

        // ==========================================
        // 4. 元のパス上文字を削除 / Remove original path text
        // ==========================================
        pathText.remove();
    }
}

main();
