#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
現在のアートボードにかかっているテキスト、またはドキュメント全体のテキストを一覧表示し、
ダイアログ内に表示しているテキストをまとめてクリップボードにコピーします。
「ドキュメント全体」は、ドキュメント内のすべてのテキストフレームを対象にします。
「アートボード外を無視」を有効にすると、どのアートボードにもかかっていないテキストを除外します。
重複を有効にすると、前後の空白、連続する空白、全角スペース、改行コード差をならしたうえで、同一内容のテキストは1つにまとめて表示します。
空テキストは一覧に含めません。UIとメッセージは日本語 / 英語に対応します。
Lists text overlapping the current artboard or across the entire document,
and copies all text shown in the dialog to the clipboard.
"Document" scans all text frames in the document.
When "Ignore outside artboards" is enabled, text not overlapping any artboard is excluded.
When duplicate filtering is enabled, line endings, leading / trailing whitespace, repeated whitespace, and full-width spaces are normalized before duplicate checks.
Empty text is excluded. Supports Japanese / English UI and messages.

更新日 / Updated: 2026-03-31
バージョン / Version: v1.0.0
*/

var SCRIPT_VERSION = "v1.0.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "テキスト一覧",
        en: "Text"
    },
    scopeArtboard: {
        ja: "現在のアートボードにかかるもの",
        en: "Artboard"
    },
    scopeAll: {
        ja: "ドキュメント全体",
        en: "Document"
    },
    dedupeGroup: {
        ja: "オプション",
        en: "Options"
    },
    dedupeText: {
        ja: "同じテキストをまとめる",
        en: "Merge duplicates"
    },
    ignoreOutside: {
        ja: "アートボード外を無視",
        en: "Ignore outside artboards"
    },
    copyAll: {
        ja: "一覧をコピー",
        en: "Copy"
    },
    close: {
        ja: "閉じる",
        en: "Close"
    },
    noDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    noTextToCopy: {
        ja: "コピーできるテキストがありません。",
        en: "Nothing to copy."
    },
    copyDone: {
        ja: "コピーしました。",
        en: "Copied."
    },
    copyError: {
        ja: "コピー中にエラーが発生しました。",
        en: "Copy failed."
    },
    tempLayerPrefix: {
        ja: "__TextList_CopyTemp__",
        en: "__TextList_CopyTemp__"
    }
};

function L(key) {
    if (!LABELS[key]) {
        return key;
    }
    return LABELS[key][lang] || LABELS[key].ja || LABELS[key].en || key;
}

main();

function main() {
    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }

    var doc = app.activeDocument;

    // アートボード内にあるか判定する関数
    function isOnArtboard(item, abRect) {
        var gb = item.geometricBounds;
        return !(gb[2] < abRect[0] || gb[0] > abRect[2] || gb[1] < abRect[3] || gb[3] > abRect[1]);
    }

    function isOnAnyArtboard(item, artboardRects) {
        for (var i = 0; i < artboardRects.length; i++) {
            if (isOnArtboard(item, artboardRects[i])) {
                return true;
            }
        }
        return false;
    }

    function getAllArtboardRects() {
        var rects = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            rects.push(doc.artboards[i].artboardRect);
        }
        return rects;
    }

    function getTextFrameContents(textFrame) {
        // 強制改行を含めて1つのテキストとして扱う
        var text = textFrame.contents;
        // 改行コードを統一（\r → \n）
        text = text.replace(/\r/g, "\n");
        return text;
    }

    function normalizeLineBreaks(text) {
        return text.replace(/\r\n?/g, "\n");
    }

    function normalizeForDuplicateCheck(text) {
        var normalized = normalizeLineBreaks(text);
        normalized = normalized.replace(/\u3000/g, " ");
        normalized = normalized.replace(/[ \t\f\v]+/g, " ");
        normalized = normalized.replace(/ *\n */g, "\n");
        normalized = normalized.replace(/^\s+|\s+$/g, "");
        return normalized;
    }

    function isEmptyText(text) {
        return normalizeForDuplicateCheck(text) === "";
    }

    function createTempCopyLayer() {
        var prefix = L("tempLayerPrefix");
        var layer = doc.layers.add();
        layer.name = prefix + new Date().getTime() + "_" + Math.floor(Math.random() * 100000);
        layer.visible = true;
        layer.locked = false;
        return layer;
    }

    function collectTextFrames(textFrames, abRect, ignoreOutside, artboardRects) {
        var entries = [];
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var shouldInclude = false;

            if (abRect) {
                shouldInclude = isOnArtboard(textFrame, abRect);
            } else if (ignoreOutside) {
                shouldInclude = isOnAnyArtboard(textFrame, artboardRects || getAllArtboardRects());
            } else {
                shouldInclude = true;
            }

            if (shouldInclude) {
                var text = getTextFrameContents(textFrame);
                if (!isEmptyText(text)) {
                    entries.push(text);
                }
            }
        }
        return entries;
    }

    function dedupeEntries(entries) {
        var result = [];
        var seen = {};
        for (var i = 0; i < entries.length; i++) {
            var key = normalizeForDuplicateCheck(entries[i]);
            if (!Object.prototype.hasOwnProperty.call(seen, key)) {
                seen[key] = true;
                result.push(entries[i]);
            }
        }
        return result;
    }

    function joinEntries(entries) {
        return entries.length ? entries.join("\n") + "\n" : "";
    }

    function gatherText(useAll, dedupe, ignoreOutside) {
        var entries;
        var allArtboardRects = getAllArtboardRects();
        if (useAll) {
            entries = collectTextFrames(doc.textFrames, null, ignoreOutside, allArtboardRects);
        } else {
            var abIndex = doc.artboards.getActiveArtboardIndex();
            var abRect = doc.artboards[abIndex].artboardRect;
            entries = collectTextFrames(doc.textFrames, abRect, false, allArtboardRects);
        }

        if (dedupe) {
            entries = dedupeEntries(entries);
        }
        return joinEntries(entries);
    }

    function refreshText(editBox, useAll, dedupe, ignoreOutside) {
        editBox.text = gatherText(useAll, dedupe, ignoreOutside);
    }

    function copyAllText(text) {
        if (isEmptyText(text)) {
            alert(L("noTextToCopy"));
            return;
        }

        var prevSelection = [];
        var i;
        for (i = 0; i < doc.selection.length; i++) {
            prevSelection.push(doc.selection[i]);
        }

        var tempLayer = null;
        var tempFrame = null;
        try {
            tempLayer = createTempCopyLayer();
            tempFrame = tempLayer.textFrames.add();
            tempFrame.contents = text;
            doc.selection = null;
            tempFrame.selected = true;
            app.executeMenuCommand("copy");
            alert(L("copyDone"));
        } catch (e) {
            alert(L("copyError") + "\n" + e);
        } finally {
            try {
                if (tempLayer) {
                    tempLayer.remove();
                } else if (tempFrame) {
                    tempFrame.remove();
                }
            } catch (_) { }

            try {
                doc.selection = null;
                for (i = 0; i < prevSelection.length; i++) {
                    try {
                        prevSelection[i].selected = true;
                    } catch (_) { }
                }
            } catch (_) { }
        }
    }

    function bindDialogEvents(ui) {
        ui.rbArtboard.onClick = handleScopeChange;
        ui.rbAll.onClick = handleScopeChange;
        ui.chkDedupe.onClick = handleScopeChange;
        ui.copyBtn.onClick = handleCopy;

        reflectEnabledUI();

        function reflectEnabledUI() {
            ui.chkIgnoreOutside.enabled = ui.rbAll.value;
        }

        function handleScopeChange() {
            reflectEnabledUI();
            refreshText(ui.editBox, ui.rbAll.value, ui.chkDedupe.value, ui.chkIgnoreOutside.value);
        }

        function handleCopy() {
            copyAllText(ui.editBox.text);
        }
    }

    var ui = buildDialogUI();
    bindDialogEvents(ui);
    refreshText(ui.editBox, false, ui.chkDedupe.value, ui.chkIgnoreOutside.value);
    ui.dlg.show();

    function buildDialogUI() {
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        var grpScopeWrap = dlg.add("group");
        grpScopeWrap.orientation = "row";
        grpScopeWrap.alignment = ["center", "top"];
        grpScopeWrap.alignChildren = ["center", "center"];

        var grpScope = grpScopeWrap.add("group");
        grpScope.orientation = "row";
        grpScope.alignChildren = ["left", "center"];
        grpScope.alignment = ["center", "center"];

        var rbArtboard = grpScope.add("radiobutton", undefined, L("scopeArtboard"));
        var rbAll = grpScope.add("radiobutton", undefined, L("scopeAll"));
        rbArtboard.value = true;

        var pnlDedupe = dlg.add("panel", undefined, L("dedupeGroup"));
        pnlDedupe.orientation = "column";
        pnlDedupe.alignChildren = ["left", "top"];
        pnlDedupe.margins = [15, 20, 15, 10];
        var chkDedupe = pnlDedupe.add("checkbox", undefined, L("dedupeText"));
        chkDedupe.value = true;

        var chkIgnoreOutside = pnlDedupe.add("checkbox", undefined, L("ignoreOutside"));
        chkIgnoreOutside.value = false;

        var editBox = dlg.add("edittext", [0, 0, 400, 300], "", { multiline: true, scrolling: true });
        editBox.active = true;

        var btnGroup = dlg.add("group");
        btnGroup.alignment = ["fill", "top"];
        var copyBtn = btnGroup.add("button", undefined, L("copyAll"));
        copyBtn.alignment = ["left", "center"];

        var closeBtn = btnGroup.add("button", undefined, L("close"), { name: "cancel" });
        closeBtn.alignment = ["right", "center"];

        return {
            dlg: dlg,
            rbArtboard: rbArtboard,
            rbAll: rbAll,
            chkDedupe: chkDedupe,
            chkIgnoreOutside: chkIgnoreOutside,
            editBox: editBox,
            copyBtn: copyBtn,
            closeBtn: closeBtn
        };
    }
}
