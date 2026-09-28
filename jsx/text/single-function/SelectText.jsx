#target illustrator
#targetengine "SelectTextEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

現在のアートボードにかかっているテキスト、またはドキュメント全体のテキストを一覧表示し、まとめてクリップボードにコピーします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectText.md

### Overview

Lists the text on the current artboard, or in the whole document, and copies it all to the clipboard.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectText";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectText.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectText.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

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
        tipScopeArtboard: {
            ja: "現在のアートボードに重なるテキストだけを集めます。",
            en: "Collects only the text that overlaps the current artboard."
        },
        tipScopeAll: {
            ja: "ドキュメント内のすべてのテキストを集めます。",
            en: "Collects every text object in the document."
        },
        tipDedupe: {
            ja: "同じ内容のテキストを1行にまとめます。",
            en: "Merges identical text into a single line."
        },
        tipIgnoreOutside: {
            ja: "アートボードの外に置かれたテキストを対象から外します。",
            en: "Leaves out text placed outside the artboards."
        },
        tipTextList: {
            ja: "集めたテキストの一覧です。ここで選択した行のオブジェクトがアートボード上でも選択されます。",
            en: "The collected text. Selecting a line here also selects the matching object on the artboard."
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

    function getLabel(key) {
        if (!LABELS[key]) {
            return key;
        }
        return LABELS[key][uiLang] || LABELS[key].ja || LABELS[key].en || key;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    main();

    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("noDocument"));
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
            var prefix = getLabel("tempLayerPrefix");
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
                alert(getLabel("noTextToCopy"));
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
                alert(getLabel("copyDone"));
            } catch (e) {
                alert(getLabel("copyError") + "\n" + e);
            } finally {
                try {
                    if (tempLayer) {
                        tempLayer.remove();
                    } else if (tempFrame) {
                        tempFrame.remove();
                    }
                } catch (e) { }

                try {
                    doc.selection = null;
                    for (i = 0; i < prevSelection.length; i++) {
                        try {
                            prevSelection[i].selected = true;
                        } catch (e) { }
                    }
                } catch (e) { }
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
        prepareDialogWindow(ui.dlg, SCRIPT_NAME);
        ui.dlg.show();

        function buildDialogUI() {
            var dialog = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
            dialog.orientation = "column";
            dialog.alignChildren = ["fill", "top"];

            var grpScopeWrap = dialog.add("group");
            grpScopeWrap.orientation = "row";
            grpScopeWrap.alignment = ["center", "top"];
            grpScopeWrap.alignChildren = ["center", "center"];

            var grpScope = grpScopeWrap.add("group");
            grpScope.orientation = "row";
            grpScope.alignChildren = ["left", "center"];
            grpScope.alignment = ["center", "center"];

            var rbArtboard = grpScope.add("radiobutton", undefined, getLabel("scopeArtboard"));
            rbArtboard.helpTip = getLabel("tipScopeArtboard");
            var rbAll = grpScope.add("radiobutton", undefined, getLabel("scopeAll"));
            rbAll.helpTip = getLabel("tipScopeAll");
            rbArtboard.value = true;

            var pnlDedupe = dialog.add("panel", undefined, getLabel("dedupeGroup"));
            pnlDedupe.orientation = "column";
            pnlDedupe.alignChildren = ["left", "top"];
            pnlDedupe.margins = [15, 20, 15, 10];
            var chkDedupe = pnlDedupe.add("checkbox", undefined, getLabel("dedupeText"));
            chkDedupe.helpTip = getLabel("tipDedupe");
            chkDedupe.value = true;

            var chkIgnoreOutside = pnlDedupe.add("checkbox", undefined, getLabel("ignoreOutside"));
            chkIgnoreOutside.helpTip = getLabel("tipIgnoreOutside");
            chkIgnoreOutside.value = false;

            var editBox = dialog.add("edittext", [0, 0, 400, 300], "", { multiline: true, scrolling: true });
            editBox.helpTip = getLabel("tipTextList");
            editBox.active = true;

            var btnGroup = dialog.add("group");
            btnGroup.alignment = ["fill", "top"];
            var copyBtn = btnGroup.add("button", undefined, getLabel("copyAll"));
            copyBtn.alignment = ["left", "center"];

            var closeBtn = btnGroup.add("button", undefined, getLabel("close"), { name: "cancel" });
            closeBtn.alignment = ["right", "center"];

            return {
                dlg: dialog,
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

})();
