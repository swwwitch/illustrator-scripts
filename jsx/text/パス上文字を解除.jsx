/*
### スクリプト名：
DetachPathText.jsx

### 更新日：
20260302

### 概要：
選択した「パス上文字（Path Text）」を、同じ文字内容・段落属性・文字属性をできるだけ維持したまま「ポイント文字」に変換します。
変換時に、書式保持モードおよびパス設定を選択するダイアログが表示されます。
元のテキストパス形状は複製され、ダイアログの設定に応じて線属性（1pt黒／線なし）が適用されます。

### 使い方：
1) パス上文字を選択
2) スクリプトを実行
3) ダイアログで書式保持のオプションを選択

### 注意：
- 文字ごとの属性（フォント/サイズ/色/トラッキング等）は、可能な範囲で復元します。
- 復元時に設定できない属性は try/catch でスキップします。
*/
#target illustrator

// ==========================================
// UI: Text Preserve Options
// ==========================================
function showOptionsDialog() {
    var dlg = new Window("dialog", "PathTextToPointText");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 18;

    var pnlKeep = dlg.add("panel", undefined, "テキスト保持");
    pnlKeep.orientation = "column";
    pnlKeep.alignChildren = ["left", "top"];
    pnlKeep.margins = [15, 20, 15, 12];

    var rbFull = pnlKeep.add("radiobutton", undefined, "書式保持（完全）");
    var rbFast = pnlKeep.add("radiobutton", undefined, "書式保持（高速）");
    var rbNone = pnlKeep.add("radiobutton", undefined, "なし（書式を削除）");
    rbFull.value = true;

    var pnlPath = dlg.add("panel", undefined, "パス");
    pnlPath.orientation = "column";
    pnlPath.alignChildren = ["left", "top"];
    pnlPath.margins = [15, 20, 15, 12];

    var rbPathBlack  = pnlPath.add("radiobutton", undefined, "1pt、黒に設定");
    var rbPathNone   = pnlPath.add("radiobutton", undefined, "線カラー：なし");
    var rbPathDelete = pnlPath.add("radiobutton", undefined, "削除");
    rbPathBlack.value = true; // デフォルト

    var grpBtns = dlg.add("group");
    grpBtns.orientation = "row";
    grpBtns.alignChildren = ["right", "center"];
    grpBtns.alignment = ["fill", "top"];

    var btnCancel = grpBtns.add("button", undefined, "キャンセル");
    var btnOk = grpBtns.add("button", undefined, "OK", { name: "ok" });

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
        alert("ドキュメントが開かれていません。");
        return;
    }

    var opt = showOptionsDialog();
    if (!opt) return; // キャンセル

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert("パス上文字を選択してください。");
        return;
    }

    var pathTexts = [];

    // 選択されたオブジェクトからパス上文字だけを抽出
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "TextFrame" && sel[i].kind === TextType.PATHTEXT) {
            pathTexts.push(sel[i]);
        }
    }

    if (pathTexts.length === 0) {
        alert("選択範囲にパス上文字が含まれていません。");
        return;
    }

    // パス上文字の処理
    for (var j = pathTexts.length - 1; j >= 0; j--) {
        var pathText = pathTexts[j];
        var originalPath = pathText.textPath;
        
        // ==========================================
        // 1. テキスト属性のスナップショットを取得
        // ==========================================
        
        // full: 文字ごとの属性まで保持
        // fast: textRange.duplicate() で高速に保持
        // none: 書式を削除（デフォルト書式）
        var charAttrs = null;
        var firstCharAttr = null;
        if (opt.mode === "full") {
            charAttrs = [];
            for (var c = 0; c < pathText.characters.length; c++) {
                var ch = pathText.characters[c];
                var ca = ch.characterAttributes;
                charAttrs.push({
                    font        : ca.textFont,
                    size        : ca.size,
                    fillColor   : ca.fillColor,
                    strokeColor : ca.strokeColor,
                    strokeWeight: ca.strokeWeight,
                    tracking    : ca.tracking,
                    baselineShift: ca.baselineShift,
                    horizontalScale: ca.horizontalScale,
                    verticalScale  : ca.verticalScale,
                    autoLeading    : ca.autoLeading,
                    leading        : ca.leading
                });
            }
        } else if (opt.mode === "fast") {
            // fast（以前の代表値スナップショットは残置：将来の拡張用）
            if (pathText.characters.length > 0) {
                var ca0 = pathText.characters[0].characterAttributes;
                firstCharAttr = {
                    font        : ca0.textFont,
                    size        : ca0.size,
                    fillColor   : ca0.fillColor,
                    strokeColor : ca0.strokeColor,
                    strokeWeight: ca0.strokeWeight,
                    tracking    : ca0.tracking,
                    baselineShift: ca0.baselineShift,
                    horizontalScale: ca0.horizontalScale,
                    verticalScale  : ca0.verticalScale,
                    autoLeading    : ca0.autoLeading,
                    leading        : ca0.leading
                };
            }
        }

        // 段落属性（先頭段落から取得）
        var justification = null;
        if (pathText.paragraphs.length > 0) {
            var paraAttr = pathText.paragraphs[0].paragraphAttributes;
            justification = paraAttr.justification;
        }

        var textContents = pathText.contents; // 文字列自体を保存

        // ==========================================
        // 2. パスを抽出（黒・1ptの線）
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
        // 3. 新しいポイントテキストを作成し、属性を復元
        // ==========================================
        var newText = doc.textFrames.add();

        // 位置をパス上文字の左端（開始点）あたりに合わせる
        var anchorPoint = originalPath.pathPoints[0].anchor;
        newText.position = [anchorPoint[0], anchorPoint[1]];

        if (opt.mode === "fast") {
            // fast: textRange を丸ごと複製（高速）
            // ※内容・書式（段落/文字）をまとめて複製する
            pathText.textRange.duplicate(newText);

            // ※文字色（塗り・線）を確実に引き継ぐ（特殊なカラースペース等は try/catch でスキップ）
            for (var rr = 0; rr < pathText.textRanges.length; rr++) {
                try {
                    var srcAttr = pathText.textRanges[rr].characterAttributes;
                    var dstAttr = newText.textRanges[rr].characterAttributes;
                    dstAttr.fillColor = srcAttr.fillColor;
                    dstAttr.strokeColor = srcAttr.strokeColor;
                } catch (eRR) { }
            }
        } else if (opt.mode === "none") {
            // none: 内容だけ流し込み（書式は新規テキストのデフォルトに任せる）
            newText.contents = textContents;
        } else {
            // full: contents を流し込み、文字ごとの属性を復元（完全）
            newText.contents = textContents; // 文字列を流し込む

            // 段落属性の復元
            if (justification !== null && newText.paragraphs.length > 0) {
                newText.paragraphs[0].paragraphAttributes.justification = justification;
            }

            // ★重要：新規テキストのデフォルト（黒1ptの線など）を先に消す
            // 後続の文字ごとの復元で必要な場合のみ線が復元される
            try {
                var nc = new NoColor();
                newText.textRange.characterAttributes.strokeColor = nc;
                newText.textRange.characterAttributes.strokeWeight = 0;
            } catch (eClr) { }
        }

        // 文字属性の復元
        if (opt.mode === "full" && charAttrs) {
            for (var c = 0; c < newText.characters.length; c++) {
                var targetCa = newText.characters[c].characterAttributes;
                var srcCa = charAttrs[c];
                if (!srcCa) break;

                try { targetCa.textFont = srcCa.font; } catch (e) { }
                try { targetCa.size = srcCa.size; } catch (e) { }
                try { targetCa.fillColor = srcCa.fillColor; } catch (e) { }
                // stroke: 元が「なし」の場合は strokeWeight を 0 にして確実に無効化
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

                // autoLeadingがfalseの場合のみleading（行送り）を設定する
                if (!srcCa.autoLeading) {
                    try { targetCa.leading = srcCa.leading; } catch (e) { }
                }
            }
        }

        // ==========================================
        // 4. 元のパス上文字を削除
        // ==========================================
        pathText.remove();
    }
}

main();
