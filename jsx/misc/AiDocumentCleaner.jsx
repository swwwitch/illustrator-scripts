#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

未使用のパネル項目（スウォッチ・グラフィックスタイル・シンボル・ブラシ・段落スタイル・文字スタイル）、不要なオブジェクト、未使用アートボードをまとめて削除します。

- ダイアログで削除対象を選択（パネル項目 / オブジェクト / その他）
- スウォッチ・グラフィックスタイル・シンボル・ブラシは、各パネルの「未使用をすべて選択 → 削除」アクションを一時的に作成・再生して未使用のみ削除（Illustrator 本体の判定）
- 段落スタイル・文字スタイルは使用情報を取得できないため、強制 ON のときだけ既定以外を削除
- オブジェクトは「不透明度0%」「アートボード外」「アクティブアートボード外」を削除
- アートボードはアートワークが載っていない空のものを削除（強制 ON ですべて。最低1つは残す）
- 「使用中も削除（強制）」ON ではパネル項目を DOM で保護対象・既定以外まで総当たり削除
- 一時アクションは finally で必ず解放・削除
- 種類ごとに削除件数を集計して結果を表示

### Overview

Removes unused panel items (swatches, graphic styles, symbols, brushes, paragraph styles, character styles), stray objects, and unused artboards in one pass.

- Choose targets in the dialog (Panel items / Objects / Other)
- Swatches / graphic styles / symbols / brushes are pruned by building and playing each panel's "Select All Unused -> Delete" action (Illustrator's own judgment)
- Paragraph / character styles have no usage info, so they are removed only in force mode
- Objects: removes those at 0% opacity, outside all artboards, or outside the active artboard
- Artboards are removed when empty (all in force mode; at least one is kept)
- In force mode, panel items are brute-force removed via the DOM down to protected/default ones
- The temporary action is always unloaded and deleted in a finally block
- Reports the deleted count per type

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ユーザー設定 / User Settings
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
function setupGroup(group, orientation, spacing) {
    var groupOrientation = orientation || "column";
    group.orientation = groupOrientation;
    /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
    group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// 一時アクション設定 / Temporary action settings
// =========================================

var ACTION_SET_NAME = "TemporaryActionSet";
var ACTION_NAME = "TemporaryActionName";
var ACTION_FILE_NAME = "~/TemporaryAction.aia";

/* パネルごとの録画値（internalName・Select All Unused 値・Delete 値・各ラベルの16進）/ Recorded per-panel values (internalName, Select All Unused value, Delete value, label hex) */
var PRUNE_SPECS = {
    swatch:       { internalName: "ai_plugin_swatches",       localizedNameHex: "e382b9e382a6e382a9e38383e38381",                                 selectValue: 11, deleteValue: 3, deleteNameHex: "44656c65746520537761746368" },
    graphicstyle: { internalName: "ai_plugin_styles",         localizedNameHex: "e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab",   selectValue: 14, deleteValue: 3, deleteNameHex: "44656c657465205374796c65" },
    symbol:       { internalName: "ai_plugin_symbol_palette", localizedNameHex: "e382b7e383b3e3839ce383ab",                                       selectValue: 12, deleteValue: 5, deleteNameHex: "44656c6574652053796d626f6c" },
    brush:        { internalName: "ai_plugin_brush",          localizedNameHex: "e38396e383a9e382b7",                                             selectValue: 8,  deleteValue: 3, deleteNameHex: "44656c657465204272757368" }
};

/* 「Select All Unused」コマンド名の16進（全パネル共通）/ Hex for the "Select All Unused" command name (shared by all panels) */
var SELECT_ALL_UNUSED_HEX = "53656c65637420416c6c20556e75736564";

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在の UI 言語を取得（ja / en）/ Get current UI language (ja / en) */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "未使用項目を削除", en: "Delete Unused Items" }
    },
    panel: {
        panelItems: { ja: "パネル項目", en: "Panel items" },
        graphic: { ja: "グラフィック関連", en: "Graphic" },
        text: { ja: "テキスト関連", en: "Text" },
        object: { ja: "オブジェクト", en: "Objects" },
        artboard: { ja: "アートボード", en: "Artboards" }
    },
    checkbox: {
        swatches: { ja: "スウォッチ", en: "Swatches" },
        graphicStyles: { ja: "グラフィックスタイル", en: "Graphic styles" },
        symbols: { ja: "シンボル", en: "Symbols" },
        brushes: { ja: "ブラシ", en: "Brushes" },
        paragraphStyles: { ja: "段落スタイル", en: "Paragraph styles" },
        characterStyles: { ja: "文字スタイル", en: "Character styles" },
        zeroOpacity: { ja: "不透明度が0%のオブジェクト", en: "Objects at 0% opacity" },
        outsideAllArtboards: { ja: "アートボード外のオブジェクト", en: "Objects outside all artboards" },
        outsideActiveArtboard: { ja: "アクティブなアートボード外のオブジェクト", en: "Objects outside the active artboard" },
        emptyClipGroup: { ja: "空のクリップグループ", en: "Empty clip groups" },
        guides: { ja: "すべてのガイド", en: "All guides" },
        artboards: { ja: "アートボード", en: "Artboards" },
        force: { ja: "使用中の項目も削除（強制）", en: "Delete items even if in use (force)" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        run: { ja: "削除", en: "Delete" }
    },
    result: {
        swatches: { ja: "スウォッチ", en: "Swatches" },
        graphicStyles: { ja: "グラフィックスタイル", en: "Graphic styles" },
        symbols: { ja: "シンボル", en: "Symbols" },
        brushes: { ja: "ブラシ", en: "Brushes" },
        paragraphStyles: { ja: "段落スタイル", en: "Paragraph styles" },
        characterStyles: { ja: "文字スタイル", en: "Character styles" },
        zeroOpacity: { ja: "不透明度0%のオブジェクト", en: "0% opacity objects" },
        outsideAllArtboards: { ja: "アートボード外のオブジェクト", en: "Objects outside all artboards" },
        outsideActiveArtboard: { ja: "アクティブアートボード外のオブジェクト", en: "Objects outside the active artboard" },
        emptyClipGroup: { ja: "空のクリップグループ", en: "Empty clip groups" },
        guides: { ja: "ガイド", en: "Guides" },
        artboards: { ja: "アートボード", en: "Artboards" }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        done: { ja: "削除が完了しました。", en: "Deletion complete." }
    },
    tooltip: {
        swatches: {
            ja: "未使用のスウォッチをアクションで削除します（プロセスカラーも判定）。強制ONで保護対象以外をすべて削除します。",
            en: "Prunes unused swatches via an action (process colors included). Force mode removes all but protected ones."
        },
        graphicStyles: {
            ja: "未使用のグラフィックスタイルをアクションで削除します。強制ONで既定（最後の1つ）以外をすべて削除します。",
            en: "Prunes unused graphic styles via an action. Force mode removes all but the default (the last one)."
        },
        symbols: {
            ja: "未使用のシンボルをアクションで削除します。強制ONですべて削除します。",
            en: "Prunes unused symbols via an action. Force mode removes them all."
        },
        brushes: {
            ja: "未使用のブラシをアクションで削除します。",
            en: "Prunes unused brushes via an action."
        },
        zeroOpacity: {
            ja: "不透明度が0%（完全に透明）のオブジェクトを削除します。グループ内も対象です。",
            en: "Removes objects at 0% opacity (fully transparent), including those inside groups."
        },
        outsideAllArtboards: {
            ja: "どのアートボードにも載っていないオブジェクトを削除します。",
            en: "Removes objects that sit on none of the artboards."
        },
        outsideActiveArtboard: {
            ja: "アクティブなアートボードに載っていないオブジェクトを削除します。",
            en: "Removes objects that do not sit on the active artboard."
        },
        emptyClipGroup: {
            ja: "クリッピングマスクのみで中身が空（塗り・線なしのパスだけ）のグループを削除します。",
            en: "Removes clip groups whose contents are empty (only paths with no fill or stroke)."
        },
        guides: {
            ja: "ドキュメント内のすべてのガイドを削除します（ロックされたレイヤーも一時的に解除）。",
            en: "Removes all guides in the document (temporarily unlocking locked layers)."
        },
        paragraphStyles: {
            ja: "使用状況を取得できないため、強制ON時のみ [標準段落スタイル] 以外を削除します。",
            en: "Usage can't be detected, so all but [Normal Paragraph Style] are removed only in force mode."
        },
        characterStyles: {
            ja: "使用状況を取得できないため、強制ON時のみ [標準文字スタイル] 以外を削除します。",
            en: "Usage can't be detected, so all but [Normal Character Style] are removed only in force mode."
        },
        artboards: {
            ja: "アートワークが載っていない空のアートボードを削除します（最低1つは残します）。強制ONですべて。",
            en: "Removes empty artboards with no artwork (keeps at least one). Force mode removes them all."
        },
        force: {
            ja: "OFF では各パネルの「未使用を選択」で未使用のみ削除します。ON では保護対象・既定以外をすべて削除します。",
            en: "When off, only unused items are pruned via each panel's Select All Unused. When on, everything but protected/default is removed."
        }
    }
};

/* ドット区切りのキーからローカライズ文字列を取得 / Resolve a localized string from a dotted key path */
function L(path) {
    var parts = path.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) {
            return path;
        }
        node = node[parts[i]];
    }
    if (node == null) {
        return path;
    }
    return node[currentLanguage] || node.ja || path;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return L(key) + (currentLanguage === "ja" ? "：" : ":");
}

/* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
function labelWithCount(key, count) {
    if (currentLanguage === "ja") {
        return L(key) + "（" + count + "）";
    }
    return L(key) + " (" + count + ")";
}

// =========================================
// メイン処理 / Main
// =========================================
(function () {
    /* ドキュメントの有無を確認 / Ensure a document is open */
    if (app.documents.length === 0) {
        alert(L('alert.noDocument'));
        return;
    }

    var doc = app.activeDocument;

    /* 種類キーごとの削除処理。run(force) で件数を返す / Handler per type key; run(force) returns the count */
    var TARGET_RUNNERS = {
        swatches:              function (force) { return deleteUnusedSwatches(doc, force); },
        graphicStyles:         function (force) { return deleteUnusedGraphicStyles(doc, force); },
        symbols:               function (force) { return deleteUnusedSymbols(doc, force); },
        brushes:               function ()      { return deleteUnusedBrushes(doc); },
        paragraphStyles:       function (force) { return deleteUnusedParagraphStyles(doc, force); },
        characterStyles:       function (force) { return deleteUnusedCharacterStyles(doc, force); },
        zeroOpacity:           function ()      { return deleteZeroOpacityObjects(doc); },
        outsideAllArtboards:   function ()      { return deleteObjectsOutsideAllArtboards(doc); },
        outsideActiveArtboard: function ()      { return deleteObjectsOutsideActiveArtboard(doc); },
        emptyClipGroup:        function ()      { return deleteEmptyClipGroups(doc); },
        guides:                function ()      { return deleteAllGuides(doc); },
        artboards:             function (force) { return deleteUnusedArtboards(doc, force); }
    };

    /* ダイアログの構成（トップパネル → サブパネル/チェックボックス）。force はパネル項目パネルの末尾に置く / Dialog layout; the force option sits at the bottom of the panel-items panel */
    var DIALOG_LAYOUT = [
        { titleKey: "panel.panelItems", force: true, subPanels: [
            { titleKey: "panel.graphic", keys: ["swatches", "graphicStyles", "symbols", "brushes"] },
            { titleKey: "panel.text", keys: ["paragraphStyles", "characterStyles"] }
        ] },
        { titleKey: "panel.object", keys: ["zeroOpacity", "emptyClipGroup", "guides"] },
        { titleKey: "panel.artboard", keys: ["outsideAllArtboards", "outsideActiveArtboard", "artboards"] }
    ];

    /* レイアウトを順に辿って全キーを取り出す / Flatten all keys in layout order */
    var ALL_KEYS = (function (layout) {
        var keys = [];
        for (var i = 0; i < layout.length; i++) {
            if (layout[i].subPanels) {
                for (var s = 0; s < layout[i].subPanels.length; s++) {
                    keys = keys.concat(layout[i].subPanels[s].keys);
                }
            } else {
                keys = keys.concat(layout[i].keys);
            }
        }
        return keys;
    })(DIALOG_LAYOUT);

    /* 削除対象を選ぶダイアログを表示 / Show the dialog for choosing what to delete */
    var dialogChoices = showDeleteDialog();
    if (!dialogChoices) {
        return;
    }

    var force = dialogChoices.force;

    /* 選択された種類ごとに削除し、件数を集計 / Delete each selected type and tally the counts */
    var deletedCounts = {};
    for (var i = 0; i < ALL_KEYS.length; i++) {
        var key = ALL_KEYS[i];
        if (dialogChoices[key]) {
            deletedCounts[key] = TARGET_RUNNERS[key](force);
        }
    }

    /* パネル表示を更新 / Refresh the panels */
    app.redraw();

    /* 完了メッセージを組み立てて表示 / Build and show the completion message */
    var resultMessage = L('alert.done') + "\n\n";
    for (i = 0; i < ALL_KEYS.length; i++) {
        key = ALL_KEYS[i];
        if (dialogChoices[key]) {
            resultMessage += formatResultLine('result.' + key, deletedCounts[key]);
        }
    }
    alert(resultMessage);

    /* 集計1行を組み立て（日本語は「件」を付ける）/ Build one result line (JA appends a counter word) */
    function formatResultLine(key, count) {
        if (currentLanguage === "ja") {
            return L(key) + ": " + count + " 件\n";
        }
        return L(key) + ": " + count + "\n";
    }

    // ==================================================
    // ダイアログ生成 / Build the dialog
    // ==================================================

    /* 削除対象を選択するダイアログを構築し、選択結果を返す / Build the dialog and return the user's selection */
    function showDeleteDialog() {
        var dialog = new Window("dialog", L('dialog.title') + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        /* レイアウト定義どおりにパネル・サブパネル・チェックボックスを生成 / Build panels, sub-panels, and checkboxes per the layout */
        var checkboxes = {};
        var forceCheckbox = null;
        for (var i = 0; i < DIALOG_LAYOUT.length; i++) {
            var node = DIALOG_LAYOUT[i];
            var panel = dialog.add("panel", undefined, L(node.titleKey));
            setupPanel(panel, 6);

            if (node.subPanels) {
                /* サブパネルは横並び（2カラム）/ Lay sub-panels out side by side (two columns) */
                var columns = panel.add("group");
                columns.orientation = "row";
                columns.alignChildren = ["fill", "top"];
                columns.alignment = "fill";
                columns.spacing = PANEL_SPACING;
                for (var s = 0; s < node.subPanels.length; s++) {
                    var sub = node.subPanels[s];
                    var subPanel = columns.add("panel", undefined, L(sub.titleKey));
                    setupPanel(subPanel, 6);
                    addCheckboxes(subPanel, sub.keys, checkboxes);
                }
            } else {
                addCheckboxes(panel, node.keys, checkboxes);
            }

            /* 強制オプションは対象パネル（パネル項目）の末尾に追加 / The force option goes at the bottom of its panel (panel items) */
            if (node.force) {
                forceCheckbox = panel.add("checkbox", undefined, L('checkbox.force'));
                forceCheckbox.alignment = "left";
                forceCheckbox.value = false;
                forceCheckbox.helpTip = L('tooltip.force');
            }
        }

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";
        buttonGroup.add("button", undefined, L('button.cancel'), { name: "cancel" });
        buttonGroup.add("button", undefined, L('button.run'), { name: "ok" });

        if (dialog.show() !== 1) {
            return null;
        }

        /* チェック状態を種類キーごとにまとめて返す / Collect checkbox states keyed by type */
        var choices = { force: forceCheckbox ? forceCheckbox.value : false };
        for (i = 0; i < ALL_KEYS.length; i++) {
            choices[ALL_KEYS[i]] = checkboxes[ALL_KEYS[i]].value;
        }
        return choices;
    }

    /* キー配列ぶんのチェックボックスを親に追加し、参照を記録 / Add a checkbox per key to the parent and record the reference */
    function addCheckboxes(parent, keys, checkboxes) {
        for (var i = 0; i < keys.length; i++) {
            var checkbox = parent.add("checkbox", undefined, L('checkbox.' + keys[i]));
            checkbox.value = true;
            checkbox.helpTip = L('tooltip.' + keys[i]);
            checkboxes[keys[i]] = checkbox;
        }
    }

    // ==================================================
    // 一時アクション / Temporary action
    // ==================================================

    /* ASCII 文字列を16進に変換 / Convert an ASCII string to hex */
    function asciiToHex(text) {
        var hex = "";
        for (var i = 0; i < text.length; i++) {
            var code = text.charCodeAt(i).toString(16);
            if (code.length < 2) {
                code = "0" + code;
            }
            hex += code;
        }
        return hex;
    }

    /* メニューコマンド1件のイベントブロックを組み立て / Build one menu-command event block */
    function buildMenuEventBlock(eventIndex, internalName, localizedNameHex, commandNameHex, value, hasDialog) {
        var lines = [
            "\t/event-" + eventIndex + " {",
            "\t\t/useRulersIn1stQuadrant 1",
            "\t\t/internalName (" + internalName + ")",
            "\t\t/localizedName [ " + (localizedNameHex.length / 2),
            "\t\t\t" + localizedNameHex,
            "\t\t]",
            "\t\t/isOpen 0",
            "\t\t/isOn 1",
            "\t\t/hasDialog " + (hasDialog ? "1" : "0")
        ];
        if (hasDialog) {
            lines.push("\t\t/showDialog 0");
        }
        lines.push(
            "\t\t/parameterCount 1",
            "\t\t/parameter-1 {",
            "\t\t\t/key 1835363957",
            "\t\t\t/showInPalette 1",
            "\t\t\t/type (enumerated)",
            "\t\t\t/name [ " + (commandNameHex.length / 2),
            "\t\t\t\t" + commandNameHex,
            "\t\t\t]",
            "\t\t\t/value " + value,
            "\t\t}",
            "\t}"
        );
        return lines.join("\n");
    }

    /* 「未使用をすべて選択 → 削除」の一時アクション定義を録画値から組み立て / Build the temporary "Select All Unused -> Delete" action from recorded values */
    function buildActionSource(setName, actionName, spec) {
        return [
            "/version 3",
            "/name [ " + setName.length,
            "\t" + asciiToHex(setName),
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "\t/name [ " + actionName.length,
            "\t\t" + asciiToHex(actionName),
            "\t]",
            "\t/keyIndex 0",
            "\t/colorIndex 0",
            "\t/isOpen 1",
            "\t/eventCount 2",
            buildMenuEventBlock(1, spec.internalName, spec.localizedNameHex, SELECT_ALL_UNUSED_HEX, spec.selectValue, false),
            buildMenuEventBlock(2, spec.internalName, spec.localizedNameHex, spec.deleteNameHex, spec.deleteValue, true),
            "}"
        ].join("\n");
    }

    /* アクションを一時ファイルに書き出して再生。close/unload/remove は finally で必ず試みる / Write, load, and play the action; close/unload/remove are always attempted in finally */
    function playTemporaryAction(actionSource, setName, actionName, fileName) {
        var actionFile = new File(fileName);
        var played = false;
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) {
                return false;
            }
            actionFile.write(actionSource);
            actionFile.close();

            /* loadAction 前に同名セットを解放 / Unload any same-name set before loading */
            try { app.unloadAction(setName, ""); } catch (e1) {}

            app.loadAction(actionFile);
            app.doScript(actionName, setName);
            played = true;
        } catch (e) {
            /* 書き出し・読み込み・再生に失敗 / Failed to write, load, or play */
        } finally {
            try { actionFile.close(); } catch (e2) {}
            try { app.unloadAction(setName, ""); } catch (e3) {}
            try { actionFile.remove(); } catch (e4) {}
        }
        return played;
    }

    /* 指定コレクションの未使用項目を、対応するアクションを再生して削除し、件数を返す / Prune a collection's unused items via its action and return the count */
    function pruneUnusedViaAction(collection, spec) {
        var countBefore = collection.length;
        var actionSource = buildActionSource(ACTION_SET_NAME, ACTION_NAME, spec);
        playTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME, ACTION_FILE_NAME);
        return Math.max(0, countBefore - collection.length);
    }

    // ==================================================
    // アートボードの空判定 / Artboard emptiness check
    // ==================================================

    /* アートボード矩形にアートワークが載っているか判定 / Determine whether any artwork sits on the artboard rectangle */
    function isArtboardEmpty(doc, artboardRect) {
        for (var i = 0; i < doc.pageItems.length; i++) {
            var bounds;
            try {
                bounds = doc.pageItems[i].visibleBounds;
            } catch (e) {
                continue;
            }
            if (rectsIntersect(artboardRect, bounds)) {
                return false;
            }
        }
        return true;
    }

    /* 2つの矩形 [left, top, right, bottom] が重なるか判定（Illustrator は上が大きいY）/ Whether two [left, top, right, bottom] rects overlap (Illustrator: top has the larger Y) */
    function rectsIntersect(a, b) {
        if (b[2] < a[0]) { return false; } /* b 右端が a 左端より左 / b right is left of a left */
        if (b[0] > a[2]) { return false; } /* b 左端が a 右端より右 / b left is right of a right */
        if (b[3] > a[1]) { return false; } /* b 下端が a 上端より上 / b bottom is above a top */
        if (b[1] < a[3]) { return false; } /* b 上端が a 下端より下 / b top is below a bottom */
        return true;
    }

    /* 矩形がいずれかの矩形と重なるか / Whether a rect overlaps any of the rects */
    function intersectsAnyRect(bounds, rects) {
        for (var i = 0; i < rects.length; i++) {
            if (rectsIntersect(rects[i], bounds)) {
                return true;
            }
        }
        return false;
    }

    /* すべてのアートボードの矩形を取得 / Get the rectangles of all artboards */
    function getAllArtboardRects(doc) {
        var rects = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            rects.push(doc.artboards[i].artboardRect);
        }
        return rects;
    }

    // ==================================================
    // オブジェクト削除 / Object deletion
    // ==================================================

    /* 不透明度が0%のオブジェクトを削除し、件数を返す（グループ内も対象）/ Remove objects at 0% opacity and return the count (including inside groups) */
    function deleteZeroOpacityObjects(doc) {
        var removedCount = 0;

        /* doc.pageItems はグループ内も含む平坦なコレクション / doc.pageItems is a flat collection that includes items inside groups */
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            if (item.typename === "GroupItem") {
                continue;
            }
            try {
                if (item.opacity === 0) {
                    item.remove();
                    removedCount++;
                }
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    /* どのアートボードにも載っていないオブジェクトを削除し、件数を返す / Remove objects that sit on none of the artboards, return the count */
    function deleteObjectsOutsideAllArtboards(doc) {
        return deleteObjectsOutsideRects(doc, getAllArtboardRects(doc));
    }

    /* アクティブなアートボードに載っていないオブジェクトを削除し、件数を返す / Remove objects not on the active artboard, return the count */
    function deleteObjectsOutsideActiveArtboard(doc) {
        var activeIndex = doc.artboards.getActiveArtboardIndex();
        return deleteObjectsOutsideRects(doc, [doc.artboards[activeIndex].artboardRect]);
    }

    /* 指定矩形のいずれにも重ならないトップレベルオブジェクトを削除し、件数を返す / Remove top-level objects overlapping none of the given rects, return the count */
    function deleteObjectsOutsideRects(doc, rects) {
        var removedCount = 0;

        /* トップレベル（レイヤー直下）のオブジェクトだけをスナップショット / Snapshot only top-level objects (direct children of a layer) */
        var topLevelItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var candidate = doc.pageItems[i];
            if (candidate.parent && candidate.parent.typename === "Layer") {
                topLevelItems.push(candidate);
            }
        }

        for (var j = topLevelItems.length - 1; j >= 0; j--) {
            var item = topLevelItems[j];
            var bounds;
            try {
                bounds = item.visibleBounds;
            } catch (e) {
                continue;
            }
            if (!intersectsAnyRect(bounds, rects)) {
                try {
                    item.remove();
                    removedCount++;
                } catch (e2) {
                    /* 削除不可（ロック等）/ Not removable (locked, etc.) */
                }
            }
        }
        return removedCount;
    }

    /* 中身が空のクリップグループを削除し、件数を返す / Remove empty clip groups and return the count */
    function deleteEmptyClipGroups(doc) {
        var removedCount = 0;
        for (var i = 0; i < doc.layers.length; i++) {
            removedCount += removeEmptyClipGroupsIn(doc.layers[i]);
        }
        return removedCount;
    }

    /* コンテナ内を再帰的に探索し、空のクリップグループを削除して件数を返す / Recurse a container, removing empty clip groups, and return the count */
    function removeEmptyClipGroupsIn(container) {
        var removed = 0;
        /* 削除でインデックスがずれるため末尾から / Iterate from the end because removal shifts indices */
        for (var i = container.pageItems.length - 1; i >= 0; i--) {
            var item = container.pageItems[i];
            if (isEmptyClipGroup(item)) {
                try {
                    item.remove();
                    removed++;
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            } else if (item.typename === "GroupItem") {
                /* 通常グループは中を再帰探索 / Recurse into ordinary groups */
                removed += removeEmptyClipGroupsIn(item);
            }
        }
        return removed;
    }

    /* クリップグループで、子がすべて塗り・線なしの PathItem か判定 / Whether the item is a clip group whose children are all paths with no fill/stroke */
    function isEmptyClipGroup(item) {
        if (!(item.typename === "GroupItem" && item.clipped === true)) {
            return false;
        }

        var children = item.pageItems;
        if (children.length === 0) {
            return false;
        }

        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            if (child.typename !== "PathItem") {
                return false;
            }
            if (child.filled === true && child.fillColor.typename !== "NoColor") {
                return false;
            }
            if (child.stroked === true && child.strokeColor.typename !== "NoColor") {
                return false;
            }
        }
        return true;
    }

    /* すべてのガイドを削除し、件数を返す（レイヤーを一時アンロックし、後で元に戻す）/ Remove all guides and return the count (temporarily unlocking layers, then restoring) */
    function deleteAllGuides(doc) {
        var removedCount = 0;
        var layers = doc.layers;

        /* レイヤーのロック状態を保存して一時解除 / Save and temporarily clear layer locks */
        var lockStates = [];
        for (var i = 0; i < layers.length; i++) {
            lockStates[i] = layers[i].locked;
            if (layers[i].locked) {
                layers[i].locked = false;
            }
        }

        var guidesWereLocked = doc.guidesLocked;
        doc.guidesLocked = false;

        /* ガイド属性を持つパスを削除 / Remove paths flagged as guides */
        var paths = doc.pathItems;
        for (var p = paths.length - 1; p >= 0; p--) {
            if (paths[p].guides) {
                try {
                    paths[p].remove();
                    removedCount++;
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
        }

        /* ロック状態を元に戻す / Restore the original lock states */
        for (var j = 0; j < layers.length; j++) {
            layers[j].locked = lockStates[j];
        }
        doc.guidesLocked = guidesWereLocked;

        return removedCount;
    }

    // ==================================================
    // 関数：未使用スウォッチ削除 / Delete unused swatches
    // ==================================================

    /* 通常はアクションで未使用のみ、強制時は保護対象以外を全削除。件数を返す / Normally prune unused via action; force mode removes all but protected ones. Returns the count */
    function deleteUnusedSwatches(doc, force) {
        if (!force) {
            return pruneUnusedViaAction(doc.swatches, PRUNE_SPECS.swatch);
        }

        /* 強制時：削除してはいけない既定スウォッチ以外を総当たり削除 / Force: remove everything except the built-in swatches */
        var protectedNames = {
            "[None]": true,
            "[Registration]": true,
            "[Black]": true,
            "[White]": true
        };

        var removedCount = 0;
        for (var i = doc.swatches.length - 1; i >= 0; i--) {
            var swatch = doc.swatches[i];
            if (protectedNames[swatch.name]) {
                continue;
            }
            try {
                swatch.remove();
                removedCount++;
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用グラフィックスタイル削除 / Delete unused graphic styles
    // ==================================================

    /* 通常はアクションで未使用のみ、強制時は既定（最後の1つ）以外を全削除。件数を返す / Normally prune unused via action; force mode removes all but the default. Returns the count */
    function deleteUnusedGraphicStyles(doc, force) {
        if (!force) {
            return pruneUnusedViaAction(doc.graphicStyles, PRUNE_SPECS.graphicstyle);
        }

        var removedCount = 0;
        for (var i = doc.graphicStyles.length - 1; i >= 0; i--) {
            /* 最後の1つ（既定スタイル）は削除不可 / The last one (default style) cannot be removed */
            if (doc.graphicStyles.length <= 1) {
                break;
            }
            try {
                doc.graphicStyles[i].remove();
                removedCount++;
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用シンボル削除 / Delete unused symbols
    // ==================================================

    /* 通常はアクションで未使用のみ、強制時はすべて削除。件数を返す / Normally prune unused via action; force mode removes them all. Returns the count */
    function deleteUnusedSymbols(doc, force) {
        if (!force) {
            return pruneUnusedViaAction(doc.symbols, PRUNE_SPECS.symbol);
        }

        var removedCount = 0;
        for (var i = doc.symbols.length - 1; i >= 0; i--) {
            try {
                doc.symbols[i].remove();
                removedCount++;
            } catch (e) {
                /* 使用中、または削除不可 / In use or not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用ブラシ削除 / Delete unused brushes
    // ==================================================

    /* アクションで未使用ブラシを削除し、件数を返す / Prune unused brushes via action and return the count */
    function deleteUnusedBrushes(doc) {
        return pruneUnusedViaAction(doc.brushes, PRUNE_SPECS.brush);
    }

    // ==================================================
    // 関数：未使用段落スタイル削除 / Delete unused paragraph styles
    // ==================================================

    /* 使用情報を取得できないため強制時のみ、既定（先頭）以外を削除。件数を返す / No usage info, so only force mode removes all but the default (first). Returns the count */
    function deleteUnusedParagraphStyles(doc, force) {
        if (!force) {
            return 0;
        }

        var removedCount = 0;
        /* インデックス0は [標準段落スタイル] なので残す / Index 0 is [Normal Paragraph Style], keep it */
        for (var i = doc.paragraphStyles.length - 1; i >= 1; i--) {
            try {
                doc.paragraphStyles[i].remove();
                removedCount++;
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用文字スタイル削除 / Delete unused character styles
    // ==================================================

    /* 使用情報を取得できないため強制時のみ、既定（先頭）以外を削除。件数を返す / No usage info, so only force mode removes all but the default (first). Returns the count */
    function deleteUnusedCharacterStyles(doc, force) {
        if (!force) {
            return 0;
        }

        var removedCount = 0;
        /* インデックス0は [標準文字スタイル] なので残す / Index 0 is [Normal Character Style], keep it */
        for (var i = doc.characterStyles.length - 1; i >= 1; i--) {
            try {
                doc.characterStyles[i].remove();
                removedCount++;
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用アートボード削除 / Delete unused artboards
    // ==================================================

    /* 空の（強制時はすべての）アートボードを削除し、最低1つは残す。件数を返す / Remove empty (or all, when forced) artboards keeping at least one, return the count */
    function deleteUnusedArtboards(doc, force) {
        var removedCount = 0;

        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            /* アートボードは最低1つ必要 / At least one artboard must remain */
            if (doc.artboards.length <= 1) {
                break;
            }

            var artboard = doc.artboards[i];
            if (!force && !isArtboardEmpty(doc, artboard.artboardRect)) {
                continue;
            }

            try {
                doc.artboards.remove(i);
                removedCount++;
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }

        return removedCount;
    }

})();
