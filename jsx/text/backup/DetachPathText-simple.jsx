#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/* バージョン / Version */
var SCRIPT_VERSION = "v1.0.9";

function safe(fn) { try { return fn(); } catch (e) { return undefined; } }

function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) return;

    var pathTexts = [];
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (it.typename === "TextFrame" && it.kind === TextType.PATHTEXT) pathTexts.push(it);
    }
    if (pathTexts.length === 0) return;

    /* パス上文字の変換処理 / Convert path text items */
    // 変換後のテキストだけを選択するため、いったん選択をクリア / Clear selection so only new texts are selected
    safe(function () { doc.selection = null; });

    for (var j = pathTexts.length - 1; j >= 0; j--) {
        var pathText = pathTexts[j];
        var originalPath = pathText.textPath;

        // ==========================================
        // 1. テキスト属性のスナップショットを取得 / Snapshot text attributes
        // ==========================================

        // full: 文字ごとの属性まで保持 / full: preserve per-character attributes
        var charAttrs = [];
        for (var c = 0; c < pathText.characters.length; c++) {
            var ca = pathText.characters[c].characterAttributes;
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

        /* 段落属性（先頭段落から取得）/ Paragraph attributes (from first paragraph) */
        var justification = null;
        if (pathText.paragraphs.length > 0) {
            var paraAttr = pathText.paragraphs[0].paragraphAttributes;
            justification = paraAttr.justification;
        }

        var textContents = pathText.contents; // 文字列を保存 / Save string contents

        // ==========================================
        // 2. 新しいポイントテキストを作成し属性を復元 / Create new point text and restore attributes
        // ==========================================
        var newText = doc.textFrames.add();

        /* 位置をパス開始点に合わせる / Align position to path start point */
        var anchorPoint = originalPath.pathPoints[0].anchor;
        newText.position = [anchorPoint[0], anchorPoint[1]];

        // full: 内容を流し込み、文字ごとの属性を復元 / full: insert content and restore per-character attributes
        newText.contents = textContents;

        /* 段落属性の復元 / Restore paragraph attributes */
        if (justification !== null && newText.paragraphs.length > 0) {
            newText.paragraphs[0].paragraphAttributes.justification = justification;
        }

        // ★重要：新規テキストのデフォルト線を先に消す / ★Important: clear default stroke on new text first
        // 後続の文字ごとの復元で必要な場合のみ線が復元される / stroke is only restored if the original had one
        safe(function () {
            var nc = new NoColor();
            newText.textRange.characterAttributes.strokeColor = nc;
            newText.textRange.characterAttributes.strokeWeight = 0;
        });

        /* 文字属性の復元 / Restore character attributes */
        var n = Math.min(newText.characters.length, charAttrs.length);
        for (var c = 0; c < n; c++) {
            var targetCa = newText.characters[c].characterAttributes;
            var srcCa = charAttrs[c];

            safe(function () { targetCa.textFont = srcCa.font; });
            safe(function () { targetCa.size = srcCa.size; });
            safe(function () { targetCa.fillColor = srcCa.fillColor; });
            safe(function () {
                var sc = srcCa.strokeColor;
                targetCa.strokeColor = sc;
                targetCa.strokeWeight = (sc && sc.typename === "NoColor") ? 0 : srcCa.strokeWeight;
            });
            safe(function () { targetCa.tracking = srcCa.tracking; });
            safe(function () { targetCa.baselineShift = srcCa.baselineShift; });
            safe(function () { targetCa.horizontalScale = srcCa.horizontalScale; });
            safe(function () { targetCa.verticalScale = srcCa.verticalScale; });
            safe(function () { targetCa.autoLeading = srcCa.autoLeading; });

            /* autoLeading が false の場合のみ leading を設定 / Set leading only when autoLeading is false */
            if (!srcCa.autoLeading) safe(function () { targetCa.leading = srcCa.leading; });
        }

        // ==========================================
        // 3. 元のパス上文字を削除（パスも同時に削除される） / Remove original path text (path is removed as well)
        // ==========================================
        pathText.remove();

        // ==========================================
        // 4. 新しいテキストを選択状態にする / Select the new text
        // ==========================================
        safe(function () { newText.selected = true; });
    }
}

main();