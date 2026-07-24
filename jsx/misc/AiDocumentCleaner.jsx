#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内の不要な要素をまとめて削除するクリーンアップツールです。ダイアログで対象を選びます。

- パネル項目（スウォッチ・シンボル・ブラシ・グラフィックスタイル・段落スタイル・文字スタイル）
  - スウォッチ・シンボル・ブラシ・グラフィックスタイルは、各パネルの「未使用をすべて選択 → 削除」アクションを一時生成・再生して未使用のみ削除（Illustrator 本体の判定）
  - 段落スタイル・文字スタイルは使用情報を取得できないため、強制 ON のときだけ既定（先頭）以外を削除
  - 「使用中も削除（強制）」ON では、パネル項目を DOM で保護対象・既定以外まで総当たり削除（ブラシは使用中・基本ブラシ以外を削除）
- パス/オブジェクト
  - 孤立点、空のテキスト、塗りも線もない（不可視）パス（ガイド・クリッピング・コンパウンドパスの構成は除外）、不透明度0%、非表示オブジェクト、リンク切れの配置画像
- グループ/レイヤー
  - 空のグループ（通常／クリップ）を再帰削除（空になった親もカスケード削除）
  - 空のレイヤー／サブレイヤーを再帰削除（ガイドだけのレイヤーは残す。_guide / _pasteboard は保護。トップレベルは最低1つ残す）
  - グループ／レイヤーの掃除は他の削除より後に実行（パスやオブジェクトの削除で空になった親も同じ実行で削除される）
- ガイド（ラジオで1つ選択。既定は「削除しない」）
  - ロックされていないガイド：メニュー「ガイドを消去」で削除（ロックされたレイヤー上のガイドは残る）
  - すべてのガイド（強制）：レイヤーのロックも一時解除して全削除
  - 現在のアートボード以外：アクティブなアートボード上にないガイドを削除
- アートボード
  - アートボード外／アクティブなアートボード外のオブジェクト、空のアートボード（強制 ON ですべて。最低1つは残す）
- 誤削除しやすい項目（非表示オブジェクト・リンク切れの配置画像・アートボード外／アクティブなアートボード外）は初期 OFF。ガイドは既定「削除しない」
- 一時アクションは finally で必ず解放・削除
- 種類ごとに削除件数を集計して表示（0件の種類は省略）

### Overview

A cleanup tool that removes unneeded elements from a document in one pass. Choose targets in the dialog.

- Panel items (swatches, symbols, brushes, graphic styles, paragraph styles, character styles)
  - Swatches / symbols / brushes / graphic styles are pruned by building and playing each panel's "Select All Unused -> Delete" action (Illustrator's own judgment)
  - Paragraph / character styles have no usage info, so they are removed only in force mode (all but the default/first)
  - In force mode, panel items are brute-force removed via the DOM down to protected/default ones (brushes: every removable one except in-use and basic brushes)
- Paths / Objects
  - Stray points, empty text frames, unpainted (invisible) paths (guides, clipping, and compound-path members excluded), 0% opacity, hidden objects, broken-link placed images
- Groups / Layers
  - Recursively removes empty groups (ordinary and clip; a parent that becomes empty is also removed)
  - Recursively removes empty layers and sublayers (guide-only layers are kept; _guide / _pasteboard are protected; at least one top-level layer remains)
  - Group / layer cleanup runs after the other deletions, so a parent emptied by path/object removal is cleaned in the same pass
- Guides (pick one radio option; default "Don't delete")
  - Guides on unlocked layers: remove via the Clear Guides menu command (guides on locked layers remain)
  - All guides (force): unlock layers temporarily and remove everything
  - Outside the active artboard: remove guides not on the active artboard
- Artboards
  - Objects outside all / the active artboard, and empty artboards (all in force mode; at least one is kept)
- Deletion-prone options (hidden objects, broken-link placed images, objects outside all / the active artboard) start unchecked; guides default to "Don't delete"
- The temporary action is always unloaded and deleted in a finally block
- Reports the deleted count per type (zero-count types are omitted)

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiDocumentCleaner";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-24";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiDocumentCleaner.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiDocumentCleaner.md
var SCRIPT_ARTICLE_URL      = "https://note.com/dtp_tranist/n/n0d70178f0f65"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 12;

/* システム管理レイヤー名（空でも削除しない）/ System-managed layer names (kept even when empty) */
var PROTECTED_LAYER_NAMES = { "_guide": true, "_pasteboard": true };

/* 初期状態でOFFにするチェックボックス（誤削除を招きやすい積極的な項目）/ Checkboxes that start unchecked (aggressive options prone to false positives) */
var UNCHECKED_BY_DEFAULT = { hiddenObjects: true, brokenLink: true, outsideAllArtboards: true, outsideActiveArtboard: true };

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// 一時アクション設定 / Temporary action settings
// =========================================

var ACTION_SET_NAME = "TemporaryActionSet";
var ACTION_NAME = "TemporaryActionName";
/* 一時ファイルはホーム直下ではなく OS の一時フォルダへ / Write the temp file to the OS temp folder, not the home directory */
var ACTION_FILE_NAME = Folder.temp.fsName + "/AiDocumentCleaner_TemporaryAction.aia";

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
        title: { ja: "不要な要素を削除", en: "Remove Unneeded Items" }
    },
    panel: {
        panelItems: { ja: "パネル項目", en: "Panel items" },
        object: { ja: "パス/オブジェクト", en: "Paths/Objects" },
        container: { ja: "グループ/レイヤー", en: "Groups/Layers" },
        guide: { ja: "ガイド", en: "Guides" },
        artboard: { ja: "アートボード", en: "Artboards" }
    },
    checkbox: {
        swatches: { ja: "スウォッチ", en: "Swatches" },
        graphicStyles: { ja: "グラフィックスタイル", en: "Graphic styles" },
        symbols: { ja: "シンボル", en: "Symbols" },
        brushes: { ja: "ブラシ", en: "Brushes" },
        paragraphStyles: { ja: "段落スタイル", en: "Paragraph styles" },
        characterStyles: { ja: "文字スタイル", en: "Character styles" },
        strayPoints: { ja: "孤立点", en: "Stray points" },
        emptyText: { ja: "空のテキスト", en: "Empty text frames" },
        noPaintPath: { ja: "塗りも線もないパス", en: "Paths with no fill or stroke" },
        zeroOpacity: { ja: "不透明度0%のオブジェクト", en: "Objects at 0% opacity" },
        hiddenObjects: { ja: "非表示オブジェクト", en: "Hidden objects" },
        brokenLink: { ja: "リンク切れの配置画像", en: "Broken-link placed images" },
        outsideAllArtboards: { ja: "アートボード外のオブジェクト", en: "Objects outside all artboards" },
        outsideActiveArtboard: { ja: "アクティブなアートボード外のオブジェクト", en: "Objects outside the active artboard" },
        emptyGroup: { ja: "空のグループ", en: "Empty groups" },
        guidesNone: { ja: "削除しない", en: "Don't delete" },
        clearGuides: { ja: "ロックされていないガイド", en: "Guides on unlocked layers" },
        guides: { ja: "すべてのガイド（強制）", en: "All guides (force)" },
        guidesOutsideActiveArtboard: { ja: "現在のアートボード以外", en: "Outside the active artboard" },
        emptyLayer: { ja: "空のレイヤー／サブレイヤー", en: "Empty layers / sublayers" },
        artboards: { ja: "空のアートボード", en: "Empty artboards" },
        force: { ja: "使用中の項目も削除", en: "Delete items even if in use" }
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
        strayPoints: { ja: "孤立点", en: "Stray points" },
        emptyText: { ja: "空のテキスト", en: "Empty text frames" },
        noPaintPath: { ja: "塗りも線もないパス", en: "Paths with no fill or stroke" },
        zeroOpacity: { ja: "不透明度0%のオブジェクト", en: "Objects at 0% opacity" },
        hiddenObjects: { ja: "非表示オブジェクト", en: "Hidden objects" },
        brokenLink: { ja: "リンク切れの配置画像", en: "Broken-link placed images" },
        outsideAllArtboards: { ja: "アートボード外のオブジェクト", en: "Objects outside all artboards" },
        outsideActiveArtboard: { ja: "アクティブなアートボード外のオブジェクト", en: "Objects outside the active artboard" },
        emptyGroup: { ja: "空のグループ", en: "Empty groups" },
        clearGuides: { ja: "ロックされていないガイド", en: "Guides on unlocked layers" },
        guides: { ja: "すべてのガイド（強制）", en: "All guides (force)" },
        guidesOutsideActiveArtboard: { ja: "現在のアートボード以外のガイド", en: "Guides outside the active artboard" },
        emptyLayer: { ja: "空のレイヤー／サブレイヤー", en: "Empty layers / sublayers" },
        artboards: { ja: "アートボード", en: "Artboards" }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        done: { ja: "削除が完了しました。", en: "Deletion complete." },
        noTarget: { ja: "削除対象はありませんでした。", en: "No items to delete." }
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
        strayPoints: {
            ja: "アンカーが1点だけで長さを持たないパス（孤立点）を削除します。",
            en: "Removes stray points (single-anchor paths with no length)."
        },
        emptyText: {
            ja: "文字が入っていない空のテキストを削除します（ポイント文字／塗り・線のないエリア内・パス上文字）。",
            en: "Removes empty text frames (point text, and area/path text whose path has no fill or stroke)."
        },
        noPaintPath: {
            ja: "塗りも線もない（画面に見えない）パスを削除します。ガイドとクリッピングパスは対象外です。",
            en: "Removes paths with no fill and no stroke (invisible). Guides and clipping paths are excluded."
        },
        zeroOpacity: {
            ja: "不透明度が0%（完全に透明）の個々のオブジェクトを削除します（グループ内の項目も対象）。グループ自体に設定した不透明度0%は対象外です。",
            en: "Removes individual objects at 0% opacity (fully transparent), including items inside groups. A group's own 0% opacity is not evaluated."
        },
        hiddenObjects: {
            ja: "非表示（隠した）オブジェクトを削除します。非表示グループはその中身ごと削除されます。",
            en: "Removes hidden objects. A hidden group is removed together with its contents."
        },
        brokenLink: {
            ja: "リンク先ファイルが見つからない配置画像（リンク切れ）を削除します。埋め込み画像や正常なリンクは対象外です。",
            en: "Removes placed images whose linked file is missing (broken links). Embedded images and valid links are kept."
        },
        outsideAllArtboards: {
            ja: "どのアートボードにも載っていないオブジェクトを削除します。",
            en: "Removes objects that sit on none of the artboards."
        },
        outsideActiveArtboard: {
            ja: "アクティブなアートボードに載っていないオブジェクトを削除します。",
            en: "Removes objects that do not sit on the active artboard."
        },
        emptyGroup: {
            ja: "中身のない空のグループを削除します。クリップグループは、マスク以外が塗り・線なしのパスだけの場合も対象です。",
            en: "Removes empty groups with no contents. Clip groups also count when their non-mask contents are only paths with no fill or stroke."
        },
        clearGuides: {
            ja: "メニューの「ガイドを消去」でガイドを削除します。ロックされたレイヤー上のガイドは残ることがあります（その場合は「すべてのガイド（強制）」を使用）。",
            en: "Removes guides via the Clear Guides menu command. Guides on locked layers may remain (use \"All guides (force)\" for those)."
        },
        guides: {
            ja: "ドキュメント内のすべてのガイドを削除します（ロックされたレイヤーも一時的に解除）。",
            en: "Removes all guides in the document (temporarily unlocking locked layers)."
        },
        guidesOutsideActiveArtboard: {
            ja: "現在（アクティブ）のアートボード上にないガイドを削除し、そのアートボード上のガイドだけを残します。",
            en: "Removes guides that are not on the active artboard, keeping only that artboard's guides."
        },
        emptyLayer: {
            ja: "中身が空のレイヤー／サブレイヤーを再帰的に削除します。ガイドだけが残っているレイヤーは残し、トップレベルは最低1つ残します。",
            en: "Recursively removes empty layers and sublayers. Layers holding only guides are kept, and at least one top-level layer remains."
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
            ja: "OFF では各パネルの「未使用を選択」で未使用のみ削除します。ON では保護対象・既定以外をすべて削除し、段落・文字スタイルの削除も有効になり、アートボードは空でなくてもすべて削除します（最低1つは残す）。",
            en: "When off, only unused items are pruned via each panel's Select All Unused. When on, everything but protected/default is removed, paragraph/character style removal is enabled, and artboards are removed even if not empty (at least one remains)."
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
        brushes:               function (force) { return deleteUnusedBrushes(doc, force); },
        paragraphStyles:       function (force) { return deleteUnusedParagraphStyles(doc, force); },
        characterStyles:       function (force) { return deleteUnusedCharacterStyles(doc, force); },
        strayPoints:           function ()      { return deleteStrayPoints(doc); },
        emptyText:             function ()      { return deleteEmptyTextFrames(doc); },
        noPaintPath:           function ()      { return deleteUnpaintedPaths(doc); },
        zeroOpacity:           function ()      { return deleteZeroOpacityObjects(doc); },
        hiddenObjects:         function ()      { return deleteHiddenObjects(doc); },
        brokenLink:            function ()      { return deleteBrokenLinkImages(doc); },
        outsideAllArtboards:   function ()      { return deleteObjectsOutsideAllArtboards(doc); },
        outsideActiveArtboard: function ()      { return deleteObjectsOutsideActiveArtboard(doc); },
        emptyGroup:            function ()      { return deleteEmptyGroups(doc); },
        clearGuides:           function ()      { return clearGuides(doc); },
        guides:                function ()      { return deleteAllGuides(doc); },
        guidesOutsideActiveArtboard: function () { return deleteGuidesOutsideActiveArtboard(doc); },
        emptyLayer:            function ()      { return deleteEmptyLayers(doc); },
        artboards:             function (force) { return deleteUnusedArtboards(doc, force); }
    };

    /* ダイアログの構成（トップパネル → サブパネル/チェックボックス）。force はパネル項目パネルの末尾に置く / Dialog layout; the force option sits at the bottom of the panel-items panel */
    var DIALOG_LAYOUT = [
        { row: [
            { column: [
                { titleKey: "panel.panelItems", force: true, keys: ["swatches", "symbols", "brushes", "graphicStyles", "paragraphStyles", "characterStyles"] },
                { titleKey: "panel.container", keys: ["emptyGroup", "emptyLayer"] }
            ] },
            { column: [
                { titleKey: "panel.object", keys: ["strayPoints", "emptyText", "noPaintPath", "zeroOpacity", "hiddenObjects", "brokenLink"] },
                { titleKey: "panel.guide", radio: true, noneKey: "guidesNone", keys: ["clearGuides", "guides", "guidesOutsideActiveArtboard"] }
            ] }
        ] },
        { titleKey: "panel.artboard", keys: ["outsideAllArtboards", "outsideActiveArtboard", "artboards"] }
    ];

    /* レイアウトを順に辿って全キーを取り出す / Flatten all keys in layout order */
    var ALL_KEYS = (function (layout) {
        var keys = [];
        function collectFromNode(node) {
            if (node.row) {
                for (var r = 0; r < node.row.length; r++) {
                    collectFromNode(node.row[r]);
                }
            } else if (node.column) {
                for (var c = 0; c < node.column.length; c++) {
                    collectFromNode(node.column[c]);
                }
            } else {
                keys = keys.concat(node.keys);
            }
        }
        for (var i = 0; i < layout.length; i++) {
            collectFromNode(layout[i]);
        }
        return keys;
    })(DIALOG_LAYOUT);

    /* 空グループ・空レイヤーの掃除は最後に回す（他の削除で空になった親も同じ実行で消せるように）/ Run container cleanup last so parents emptied by other deletions are removed in the same pass */
    var CONTAINER_CLEANUP_KEYS = { emptyGroup: true, emptyLayer: true };
    var EXECUTION_KEYS = (function (keys) {
        var head = [];
        var tail = [];
        for (var i = 0; i < keys.length; i++) {
            (CONTAINER_CLEANUP_KEYS[keys[i]] ? tail : head).push(keys[i]);
        }
        return head.concat(tail);
    })(ALL_KEYS);

    /* 削除対象を選ぶダイアログを表示 / Show the dialog for choosing what to delete */
    var dialogChoices = showDeleteDialog();
    if (!dialogChoices) {
        return;
    }

    var force = dialogChoices.force;

    /* 選択された種類ごとに削除し、件数を集計（実行順は EXECUTION_KEYS）/ Delete each selected type and tally the counts (in EXECUTION_KEYS order) */
    var deletedCounts = {};
    for (var i = 0; i < EXECUTION_KEYS.length; i++) {
        var key = EXECUTION_KEYS[i];
        if (dialogChoices[key]) {
            deletedCounts[key] = TARGET_RUNNERS[key](force);
        }
    }

    /* パネル表示を更新 / Refresh the panels */
    app.redraw();

    /* 完了メッセージを組み立てて表示（0件の項目は省略）/ Build and show the completion message (omit zero-count types) */
    var resultLines = "";
    for (i = 0; i < ALL_KEYS.length; i++) {
        key = ALL_KEYS[i];
        if (dialogChoices[key] && deletedCounts[key] > 0) {
            resultLines += formatResultLine('result.' + key, deletedCounts[key]);
        }
    }
    /* 1件も削除されなければ「対象なし」だけ表示 / If nothing was deleted, show only the no-target message */
    alert(resultLines === "" ? L('alert.noTarget') : L('alert.done') + "\n\n" + resultLines);

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

        /* 1ノードぶんのパネルを親コンテナに生成 / Build one node's panel inside a parent container */
        function buildPanelNode(parent, node) {
            /* column ノードは複数パネルを縦積みするグループ / A column node stacks several panels vertically */
            if (node.column) {
                var stack = parent.add("group");
                stack.orientation = "column";
                stack.alignChildren = ["fill", "top"];
                stack.alignment = "fill";
                stack.spacing = PANEL_SPACING;
                for (var k = 0; k < node.column.length; k++) {
                    buildPanelNode(stack, node.column[k]);
                }
                return;
            }

            var panel = parent.add("panel", undefined, L(node.titleKey));
            setupPanel(panel, 6);
            if (node.radio) {
                addRadioGroup(panel, node, checkboxes);
            } else {
                addCheckboxes(panel, node.keys, checkboxes);
            }

            /* 強制オプションはグループに入れて対象パネル（パネル項目）の末尾に追加、上にマージン6 / The force option sits in a group at the bottom of its panel (panel items), with a 6px top margin */
            if (node.force) {
                /* 使用中削除オプションの上に区切り線（上に余白+2）/ Divider above the force option, with a little (+2) space above it */
                var dividerWrap = panel.add("group");
                dividerWrap.orientation = "column";
                dividerWrap.alignChildren = ["fill", "top"];
                dividerWrap.alignment = ["fill", "top"];
                dividerWrap.margins = [0, 2, 0, 0];
                dividerWrap.spacing = 0;
                var forceDivider = dividerWrap.add("panel");
                forceDivider.alignment = ["fill", "top"];
                forceDivider.minimumSize.height = forceDivider.maximumSize.height = 1;

                var forceGroup = panel.add("group");
                forceGroup.orientation = "column";
                forceGroup.alignChildren = ["left", "top"];
                forceGroup.alignment = "left";
                forceGroup.margins = [0, 6, 0, 0];
                forceCheckbox = forceGroup.add("checkbox", undefined, L('checkbox.force'));
                forceCheckbox.value = false;
                forceCheckbox.helpTip = L('tooltip.force');
            }
        }

        for (var i = 0; i < DIALOG_LAYOUT.length; i++) {
            var node = DIALOG_LAYOUT[i];
            if (node.row) {
                /* 複数パネルを横並び / Lay multiple panels side by side */
                var panelRow = dialog.add("group");
                panelRow.orientation = "row";
                panelRow.alignChildren = ["fill", "top"];
                panelRow.alignment = "fill";
                panelRow.spacing = PANEL_SPACING;
                for (var r = 0; r < node.row.length; r++) {
                    buildPanelNode(panelRow, node.row[r]);
                }
            } else {
                buildPanelNode(dialog, node);
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
            checkbox.value = !UNCHECKED_BY_DEFAULT[keys[i]];
            checkbox.helpTip = L('tooltip.' + keys[i]);
            checkboxes[keys[i]] = checkbox;
        }
    }

    /* ラジオグループを生成。先頭に「削除しない」（既定で選択）を置き、各キーのラジオを記録 / Build a radio group with a "Don't delete" option (selected by default) first, recording each key's radio */
    function addRadioGroup(panel, node, checkboxes) {
        var noneRadio = panel.add("radiobutton", undefined, L('checkbox.' + node.noneKey));
        noneRadio.value = true;
        for (var i = 0; i < node.keys.length; i++) {
            var radio = panel.add("radiobutton", undefined, L('checkbox.' + node.keys[i]));
            radio.helpTip = L('tooltip.' + node.keys[i]);
            radio.value = false;
            checkboxes[node.keys[i]] = radio;
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

    /* 失敗しても無視してよい後始末を実行 / Run best-effort cleanup, ignoring any failure */
    function ignoringErrors(action) {
        try {
            action();
        } catch (e) {
            /* 後始末の失敗は無視 / Ignore cleanup failures */
        }
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
            ignoringErrors(function () { app.unloadAction(setName, ""); });

            app.loadAction(actionFile);
            app.doScript(actionName, setName);
            played = true;
        } catch (e) {
            /* 書き出し・読み込み・再生に失敗 / Failed to write, load, or play */
        } finally {
            ignoringErrors(function () { actionFile.close(); });
            ignoringErrors(function () { app.unloadAction(setName, ""); });
            ignoringErrors(function () { actionFile.remove(); });
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

    /* アートボード矩形にアートワークが載っているか判定（ガイドは対象外）/ Determine whether any artwork sits on the artboard rectangle (guides are ignored) */
    function isArtboardEmpty(doc, artboardRect) {
        for (var i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            var bounds;
            try {
                /* ガイドはアートワークとみなさない / Guides don't count as artwork */
                if (item.typename === "PathItem" && item.guides) {
                    continue;
                }
                bounds = item.visibleBounds;
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
    // 孤立点・空テキスト削除 / Stray points and empty text
    // 移植元 / Ported from: 不要なアイテムを削除.jsx (c) 2020 Toshiyuki Takahashi, MIT License
    // ==================================================

    /* 孤立点（アンカー1点・長さ0のパス）を削除し、件数を返す / Remove stray points (single-anchor, zero-length paths), return the count */
    function deleteStrayPoints(doc) {
        var removedCount = 0;
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            if (item.typename !== "PathItem") {
                continue;
            }
            try {
                if (item.pathPoints.length < 2 && item.length <= 0) {
                    item.remove();
                    removedCount++;
                }
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    /* 文字のない空テキストを削除し、件数を返す（エリア内/パス上は塗り・線のないパスのみ）/ Remove empty text frames, return the count (area/path text only when the path has no fill/stroke) */
    function deleteEmptyTextFrames(doc) {
        var removedCount = 0;
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            if (item.typename !== "TextFrame") {
                continue;
            }
            try {
                if (item.contents.length >= 1) {
                    continue;
                }
                var shouldRemove = false;
                if (item.kind === TextType.POINTTEXT) {
                    shouldRemove = true;
                } else if (item.kind === TextType.AREATEXT || item.kind === TextType.PATHTEXT) {
                    /* テキストパスが塗り・線なしのときだけ削除 / Remove only when the text path has no fill/stroke */
                    if (!item.textPath.stroked && !item.textPath.filled) {
                        shouldRemove = true;
                    }
                }
                if (shouldRemove) {
                    item.remove();
                    removedCount++;
                }
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // オブジェクト削除 / Object deletion
    // ==================================================

    /* 塗りも線もない（不可視の）パスを削除し、件数を返す。ガイド・クリッピングパスは除外 / Remove paths with no fill and no stroke (invisible) and return the count; guides and clipping paths are excluded */
    function deleteUnpaintedPaths(doc) {
        var removedCount = 0;
        var paths = doc.pathItems;
        for (var i = paths.length - 1; i >= 0; i--) {
            var path = paths[i];
            try {
                /* ガイドは別オプションで扱う / Guides are handled by a separate option */
                if (path.guides) {
                    continue;
                }
                /* クリッピングパスはグループの一部なので残す / Keep clipping paths (part of a clip group) */
                if (path.clipping) {
                    continue;
                }
                /* コンパウンドパスの構成パス（穴など）は単体で削除しない / Don't delete a compound path's member paths (holes, etc.) */
                if (path.parent && path.parent.typename === "CompoundPathItem") {
                    continue;
                }
                if (!path.filled && !path.stroked) {
                    path.remove();
                    removedCount++;
                }
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

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

    /* 非表示オブジェクトを削除し、件数を返す（非表示グループは中身ごと）/ Remove hidden objects and return the count (hidden groups go with their contents) */
    function deleteHiddenObjects(doc) {
        /* まず参照だけを収集（この間はコレクションを変更しない）。非表示グループを消すと子のインデックスがずれ、末尾からの走査でも取りこぼすため
           Collect references first without mutating the collection; removing a hidden group shifts child indices, so even a reverse scan would skip items */
        var targets = [];
        var items = doc.pageItems;
        for (var i = 0; i < items.length; i++) {
            try {
                if (items[i].hidden) {
                    targets.push(items[i]);
                }
            } catch (e) {
                /* 判定不可はスキップ / Skip items we can't test */
            }
        }

        var removedCount = 0;
        for (var j = 0; j < targets.length; j++) {
            try {
                targets[j].remove();
                removedCount++;
            } catch (e) {
                /* 親ごと削除済み、または削除不可 / Already removed with its parent, or not removable */
            }
        }
        return removedCount;
    }

    /* リンク切れ（リンク先が見つからない）の配置画像を削除し、件数を返す。埋め込み・正常リンクは対象外 / Remove placed images with a missing link and return the count; embedded images and valid links are kept */
    function deleteBrokenLinkImages(doc) {
        var removedCount = 0;
        var placed = doc.placedItems;
        for (var i = placed.length - 1; i >= 0; i--) {
            var item = placed[i];
            var isBroken = false;
            try {
                var linkedFile = item.file;
                isBroken = (!linkedFile || !linkedFile.exists);
            } catch (e) {
                /* .file 取得で例外＝リンク切れ扱い / A throwing .file access means a missing link */
                isBroken = true;
            }
            if (isBroken) {
                try {
                    item.remove();
                    removedCount++;
                } catch (e2) {
                    /* 削除不可 / Not removable */
                }
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

        /* トップレベル（レイヤー直下）のオブジェクトだけをスナップショット。ガイドは専用オプションで扱うため除外 / Snapshot only top-level objects (direct children of a layer); guides are excluded (handled by the dedicated guide option) */
        var topLevelItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var candidate = doc.pageItems[i];
            if (candidate.typename === "PathItem" && candidate.guides) {
                continue;
            }
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

    /* 空のグループ（通常グループ・クリップグループ）を削除し、件数を返す / Remove empty groups (ordinary and clip groups) and return the count */
    function deleteEmptyGroups(doc) {
        var removedCount = 0;
        for (var i = 0; i < doc.layers.length; i++) {
            removedCount += removeEmptyGroupsIn(doc.layers[i]);
        }
        return removedCount;
    }

    /* コンテナ内を再帰的に探索し、空のグループを削除して件数を返す。子を先に掃除するので、空になった親も同じパスで削除できる / Recurse a container removing empty groups; children are cleaned first so a parent that becomes empty is removed in the same pass */
    function removeEmptyGroupsIn(container) {
        var removed = 0;
        /* サブレイヤーの中身は layer.pageItems に含まれないため、先に再帰する / Sublayer contents aren't in layer.pageItems, so recurse into sublayers first */
        if (container.typename === "Layer") {
            for (var s = container.layers.length - 1; s >= 0; s--) {
                removed += removeEmptyGroupsIn(container.layers[s]);
            }
        }
        /* 削除でインデックスがずれるため末尾から / Iterate from the end because removal shifts indices */
        for (var i = container.pageItems.length - 1; i >= 0; i--) {
            var item = container.pageItems[i];
            if (item.typename === "GroupItem") {
                /* 先に中を掃除してから自身の空判定 / Clean inside first, then test this group */
                removed += removeEmptyGroupsIn(item);
                if (isEmptyGroup(item)) {
                    try {
                        item.remove();
                        removed++;
                    } catch (e) {
                        /* 削除不可 / Not removable */
                    }
                }
            }
        }
        return removed;
    }

    /* グループが空か判定。子が無いグループ、またはマスク以外が塗り・線なしのパスだけのクリップグループ / Whether a group is empty: no children, or a clip group whose non-mask contents are only paths with no fill/stroke */
    function isEmptyGroup(item) {
        if (item.typename !== "GroupItem") {
            return false;
        }

        var children = item.pageItems;
        /* 子のないグループは空 / A group with no children is empty */
        if (children.length === 0) {
            return true;
        }

        /* クリップグループのみ、中身が塗り・線なしパスだけなら空とみなす / Only for clip groups: empty when all contents are unpainted paths */
        if (item.clipped !== true) {
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

    /* ガイド属性を持つパスの数を数える / Count paths flagged as guides */
    function countGuidePaths(doc) {
        var count = 0;
        var paths = doc.pathItems;
        for (var i = 0; i < paths.length; i++) {
            if (paths[i].guides) {
                count++;
            }
        }
        return count;
    }

    /* メニューコマンド「ガイドを消去」でガイドを削除し、件数を返す / Remove guides via the Clear Guides menu command, return the count */
    function clearGuides(doc) {
        var countBefore = countGuidePaths(doc);
        var guidesWereLocked = doc.guidesLocked;
        try {
            /* ロックされたガイドも消去できるよう一時解除 / Temporarily unlock so locked guides can be cleared too */
            doc.guidesLocked = false;
            app.executeMenuCommand("clearguide");
        } catch (e) {
            /* 消去不可 / Could not clear */
        } finally {
            /* 例外時もロック状態を必ず元に戻す / Always restore the lock state, even on error */
            doc.guidesLocked = guidesWereLocked;
        }
        return Math.max(0, countBefore - countGuidePaths(doc));
    }

    /* ガイドのロック・レイヤーロックを一時解除し、判定関数が真のガイドを削除して件数を返す。ロック状態は finally で必ず復元 / Temporarily clear guide/layer locks, remove guides for which the predicate is true, return the count; locks are always restored in finally */
    function removeGuidesWhere(doc, shouldRemove) {
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

        try {
            doc.guidesLocked = false;

            var paths = doc.pathItems;
            for (var p = paths.length - 1; p >= 0; p--) {
                if (!paths[p].guides) {
                    continue;
                }
                try {
                    if (shouldRemove(paths[p])) {
                        paths[p].remove();
                        removedCount++;
                    }
                } catch (e) {
                    /* 判定不可・削除不可はスキップ / Skip guides we can't test or remove */
                }
            }
        } finally {
            /* 例外時もロック状態を必ず元に戻す / Always restore the lock states, even on error */
            for (var j = 0; j < layers.length; j++) {
                layers[j].locked = lockStates[j];
            }
            doc.guidesLocked = guidesWereLocked;
        }

        return removedCount;
    }

    /* すべてのガイドを削除し、件数を返す / Remove all guides and return the count */
    function deleteAllGuides(doc) {
        return removeGuidesWhere(doc, function () { return true; });
    }

    /* 現在（アクティブ）のアートボード上にないガイドを削除し、件数を返す / Remove guides not on the active artboard and return the count */
    function deleteGuidesOutsideActiveArtboard(doc) {
        var activeRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        return removeGuidesWhere(doc, function (guidePath) {
            return !rectsIntersect(activeRect, guidePath.geometricBounds);
        });
    }

    // ==================================================
    // 関数：空のレイヤー削除 / Delete empty layers
    // ==================================================

    /* 中身が空（pageItems もサブレイヤーも無い）のレイヤーを削除し、件数を返す。ガイドは pageItems に含まれるためガイドのみのレイヤーは残る。トップレベルは最低1つ残す
       Remove empty layers (no pageItems and no sublayers) and return the count; guides count as pageItems so guide-only layers stay, and at least one top-level layer remains */
    function deleteEmptyLayers(doc) {
        var removedCount = 0;
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            var topLayer = doc.layers[i];
            /* 先に空のサブレイヤーを削除 / Remove empty sublayers first */
            removedCount += removeEmptySublayers(topLayer);
            if (PROTECTED_LAYER_NAMES[topLayer.name]) {
                continue;
            }
            /* トップレベルは最低1つ必要 / At least one top-level layer must remain */
            if (topLayer.pageItems.length === 0 && topLayer.layers.length === 0 && doc.layers.length > 1) {
                try {
                    topLayer.locked = false;
                    topLayer.remove();
                    removedCount++;
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
        }
        return removedCount;
    }

    /* サブレイヤーを再帰的に処理し、空のものを削除して件数を返す / Recurse sublayers, removing empty ones, return the count */
    function removeEmptySublayers(parentLayer) {
        var removed = 0;
        for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
            var subLayer = parentLayer.layers[i];
            removed += removeEmptySublayers(subLayer);
            if (PROTECTED_LAYER_NAMES[subLayer.name]) {
                continue;
            }
            if (subLayer.pageItems.length === 0 && subLayer.layers.length === 0) {
                try {
                    subLayer.locked = false;
                    subLayer.remove();
                    removed++;
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
        }
        return removed;
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

    /* 通常はアクションで未使用のみ、強制時は削除できるものをすべて削除（使用中・基本ブラシは不可）。件数を返す / Normally prune unused via action; force mode removes every removable brush (in-use and basic brushes can't be removed). Returns the count */
    function deleteUnusedBrushes(doc, force) {
        if (!force) {
            return pruneUnusedViaAction(doc.brushes, PRUNE_SPECS.brush);
        }

        var removedCount = 0;
        for (var i = doc.brushes.length - 1; i >= 0; i--) {
            try {
                doc.brushes[i].remove();
                removedCount++;
            } catch (e) {
                /* 使用中・基本ブラシなど削除不可 / In use, basic brush, or otherwise not removable */
            }
        }
        return removedCount;
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
