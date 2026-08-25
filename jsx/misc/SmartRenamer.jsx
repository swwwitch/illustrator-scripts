#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード・シンボル・レイヤー・グラフィックスタイルの名前を、接頭辞／接尾辞／名前の基準／検索置換を組み合わせて一括リネームします。
ダイアログ上で対象の絞り込み・並び替え・個別の手動編集ができ、結果はプレビューで確認できます。

詳細は README を参照してください。

### Overview

Renames artboards, symbols, layers and graphic styles in bulk, combining a prefix, a suffix, a naming basis and find-and-replace.
The dialog filters, reorders and hand-edits individual entries, with a preview of the result.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartRenamer";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-24";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenamer.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartRenamer.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2db43c753c0b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログを開いたときに選ばれる種類 / Item type selected when the dialog opens */
    var DEFAULT_ITEM_TYPE = "artboard";

    /* 「正規表現」チェックボックスの初期状態 / Initial state of the Regex checkbox */
    var DEFAULT_USE_REGEX = true;

    /* リネーム対象外にするグラフィックスタイル名（角括弧で囲まれた予約スタイル）
       Graphic styles excluded from renaming (reserved names wrapped in brackets) */
    var RESERVED_STYLE_NAME = /^\[.*\]$/;

    /* 名前が重複したときに付ける区切り / Separator inserted when names collide */
    var COLLISION_SEPARATOR = "_";

    /* アートボードの並び替え中に使う一時名の接頭辞 / Prefix for temporary artboard names used while reordering */
    var TEMP_ARTBOARD_PREFIX = "__tmp_ab_";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS      = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING      = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS       = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING       = 8;                  /* パネル内の要素間隔 / panel spacing */
    var DENSE_SPACING       = 6;                  /* 密なパネル・一覧行の間隔 / spacing for dense panels and rows */
    var COLUMN_SPACING      = 12;                 /* 2カラムの間隔 / gap between columns */
    var TOKEN_SPACING       = 4;                  /* トークンボタン列の間隔 / gap between token buttons */
    var TOKEN_BUTTON_SIZE   = [28, 20];           /* トークンボタンの既定サイズ [幅,高さ] / default token button size */
    var NARROW_BUTTON_WIDTH = 22;                 /* 1文字ボタンの幅 / width of single-character buttons */
    var MOVE_BUTTON_SIZE    = [56, 22];           /* 並び替えボタンのサイズ [幅,高さ] / reorder button size */
    var FIELD_LABEL_WIDTH   = 36;                 /* 「検索」「置換」ラベルの幅 / width of the find/replace labels */
    var REGEX_GAP_WIDTH     = 20;                 /* 「正規表現」の手前に置く余白 / gap before the Regex checkbox */
    var LIST_MAX_HEIGHT     = 320;                /* 一覧の最大高さ / maximum height of the item list */
    var LIST_COLUMN_WIDTHS  = {                   /* 一覧の列幅 / column widths of the item list */
        order: 24,
        select: 28,
        currentName: 140,
        arrow: 14,
        newName: 160
    };
    var PREFIX_CHARS        = 14;                 /* 接頭辞入力欄の文字数 / width of the prefix field */
    var SUFFIX_CHARS        = 16;                 /* 接尾辞入力欄の文字数 / width of the suffix field */
    var CUSTOM_CHARS        = 12;                 /* 「指定」入力欄の文字数 / width of the custom text field */
    var FIND_CHARS          = 14;                 /* 検索・置換入力欄の文字数 / width of the find/replace fields */
    var FILTER_CHARS        = 10;                 /* フィルター入力欄の文字数 / width of the filter fields */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.toLowerCase().indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /* ラベル定義 / Label definitions (JA/EN) */
    var LABELS = {
        dialog: {
            title: { ja: "スマートリネーム", en: "Smart Renamer" }
        },
        panel: {
            renameRules: { ja: "リネーム条件", en: "Rename Rules" },
            prefix: { ja: "接頭辞", en: "Prefix" },
            suffix: { ja: "接尾辞", en: "Suffix" },
            nameSource: { ja: "名前の基準", en: "Name Source" },
            findReplace: { ja: "検索・置換", en: "Find / Replace" },
            filter: { ja: "フィルター", en: "Filter" },
            list: { ja: "リスト（並び替え／リネーム）", en: "List (Reorder / Rename)" }
        },
        radio: {
            itemTypeArtboard: { ja: "アートボード", en: "Artboard" },
            itemTypeSymbol: { ja: "シンボル", en: "Symbol" },
            itemTypeLayer: { ja: "レイヤー", en: "Layer" },
            itemTypeGraphicStyle: { ja: "グラフィックスタイル", en: "Graphic Style" },
            originalName: { ja: "元の名称", en: "Original Name" },
            frontmost: { ja: "最前面のテキスト", en: "Frontmost Text" },
            custom: { ja: "指定", en: "Custom" },
            allItems: { ja: "すべて", en: "All" },
            rangeItems: { ja: "指定範囲", en: "Range" }
        },
        checkbox: {
            searchFilter: { ja: "検索でフィルター", en: "Filter by search" },
            regex: { ja: "正規表現", en: "Regex" }
        },
        fieldLabel: {
            find: { ja: "検索", en: "Find" },
            replace: { ja: "置換", en: "Replace" }
        },
        listHeader: {
            order: { ja: "順", en: "#" },
            select: { ja: "選択", en: "Sel" },
            currentName: { ja: "現在の名前", en: "Current Name" },
            newName: { ja: "新しい名前", en: "New Name" }
        },
        button: {
            moveTop: { ja: "↑ 先頭へ", en: "↑ Top" },
            moveUp: { ja: "↑ 上へ", en: "↑ Up" },
            moveDown: { ja: "↓ 下へ", en: "↓ Down" },
            moveBottom: { ja: "↓ 末尾へ", en: "↓ Bottom" },
            refresh: { ja: "更新", en: "Refresh" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            needSettings: {
                ja: "接頭辞・接尾辞・検索文字列のいずれかを入力してください。",
                en: "Enter a prefix, suffix, or find text to rename."
            },
            emptyName: {
                ja: "{n} 番目の新しい名前が空です。名前を入力してください。",
                en: "Item {n}: new name is empty. Please enter a name."
            }
        },
        tooltip: {
            searchFilter: {
                ja: "現在の名前に指定文字列を含む項目だけをチェックします",
                en: "Check only items whose current names contain the specified text"
            }
        }
    };

    /**
     * ドット区切りのキーからUI言語のラベルを取得する
     * @param {string} labelPath - "panel.filter" のようなドット区切りのキー
     * @param {object} [params] - {n: 3} のような差し込み値
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath, params) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        var labelText = labelNode[uiLang] || labelNode["en"];
        if (!labelText) return labelPath;
        if (params) {
            for (var paramKey in params) {
                if (!params.hasOwnProperty(paramKey)) continue;
                labelText = labelText.replace(new RegExp("\\{" + paramKey + "\\}", "g"), params[paramKey]);
            }
        }
        return labelText;
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ダイアログウィンドウの共通設定を適用する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定を適用する（alignment と alignChildren は必ず対で指定する）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [alignment] - グループ自身の横方向の配置（省略時は "left"）
     * @param {number} [spacing] - グループ内の要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [alignment || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Group|Window} parentContainer - 追加先
     * @param {string} labelText - パネルのタイトル
     * @param {number} [spacing] - パネル内の要素間隔
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, labelText, spacing) {
        var createdPanel = parentContainer.add("panel", undefined, labelText);
        setupPanel(createdPanel, spacing);
        return createdPanel;
    }

    /**
     * 縦並びのカラムグループを生成する
     * @param {Group|Window} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addColumnGroup(parentContainer) {
        var columnGroup = parentContainer.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = PANEL_SPACING;
        return columnGroup;
    }

    /**
     * 幅を固定した statictext を追加する（一覧の列見出し・行ラベル用）
     * @param {Group} parentContainer - 追加先
     * @param {string} text - 表示文字列
     * @param {number} width - 列幅
     * @returns {StaticText} 生成したテキスト
     */
    function addFixedWidthText(parentContainer, text, width) {
        var staticText = parentContainer.add("statictext", undefined, text);
        staticText.preferredSize.width = width;
        return staticText;
    }

    /**
     * 伸縮しない空きスペースを追加する（3つのサイズを揃えないと潰れる）
     * @param {Group} parentContainer - 追加先
     * @param {number} width - 空ける幅
     * @returns {Group} 生成したスペーサー
     */
    function addFixedSpacer(parentContainer, width) {
        var spacer = parentContainer.add("group");
        spacer.minimumSize = [width, 1];
        spacer.preferredSize = [width, 1];
        spacer.maximumSize = [width, 1];
        return spacer;
    }

    /**
     * 入力欄へフォーカスを移す（環境によっては失敗するため握りつぶす）
     * @param {EditText} targetInput - 対象の入力欄
     * @returns {void}
     */
    function focusField(targetInput) {
        try { targetInput.active = true; } catch (focusError) { }
    }

    /**
     * Option（Alt）キーが押されているか
     * @returns {boolean} 押されていれば true
     */
    function isOptionKeyHeld() {
        return !!(ScriptUI.environment && ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey);
    }

    // =========================================
    // トークン挿入ボタン / Token insert buttons
    // =========================================

    /* 接頭辞・接尾辞に挿入するトークン / Tokens inserted into the prefix and suffix fields */
    var AFFIX_TOKENS = [
        { label: "1", value: "{#1}", width: NARROW_BUTTON_WIDTH },
        { label: "01", value: "{#01}" },
        { label: "-", value: "-", width: NARROW_BUTTON_WIDTH },
        { label: "_", value: "_", width: NARROW_BUTTON_WIDTH },
        { label: "#FN", value: "#FN", width: 40 },
        { label: "#DT", value: "#DT", width: 40 }
    ];

    /* 検索欄に挿入する正規表現ショートカット / Regex shortcuts inserted into the find field */
    var FIND_PATTERN_TOKENS = [
        { label: "#", value: "\\d", width: NARROW_BUTTON_WIDTH },
        { label: "##", value: "\\d+" },
        { label: "*", value: ".+", width: NARROW_BUTTON_WIDTH }
    ];

    /* 置換欄に挿入するトークン / Tokens inserted into the replace field */
    var REPLACE_TOKENS = [
        { label: "#", value: "{#1}", width: NARROW_BUTTON_WIDTH },
        { label: "##", value: "{#01}" },
        { label: "-", value: "-", width: NARROW_BUTTON_WIDTH },
        { label: "_", value: "_", width: NARROW_BUTTON_WIDTH }
    ];

    /**
     * トークン挿入ボタンの行を作る
     * @param {Panel|Group} parentContainer - 追加先
     * @param {EditText} targetInput - 挿入先の入力欄
     * @param {Array<object>} tokens - {label, value, width} の配列
     * @param {boolean} withClearButton - 末尾にクリアボタン（x）を置くか
     * @returns {Group} 生成した行グループ
     */
    function addTokenRow(parentContainer, targetInput, tokens, withClearButton) {
        var tokenRow = parentContainer.add("group");
        setupRow(tokenRow, "left", TOKEN_SPACING);
        tokenRow.margins = 0;
        for (var tokenIdx = 0; tokenIdx < tokens.length; tokenIdx++) {
            (function (token) {
                var tokenButton = tokenRow.add("button", undefined, token.label);
                tokenButton.preferredSize = [token.width || TOKEN_BUTTON_SIZE[0], TOKEN_BUTTON_SIZE[1]];
                tokenButton.onClick = function () {
                    targetInput.text = targetInput.text + token.value;
                    targetInput.notify("onChange");
                };
            })(tokens[tokenIdx]);
        }
        if (withClearButton) {
            var clearButton = tokenRow.add("button", undefined, "x");
            clearButton.preferredSize = [NARROW_BUTTON_WIDTH, TOKEN_BUTTON_SIZE[1]];
            clearButton.onClick = function () {
                targetInput.text = "";
                targetInput.notify("onChange");
            };
        }
        return tokenRow;
    }

    // =========================================
    // 状態の保存と復元 / State capture and restore
    // =========================================

    /**
     * 失敗内容を ExtendScript コンソールへ出力する
     * @param {string} context - どこで失敗したかを示す文字列
     * @param {object} error - 捕捉した例外
     * @returns {void}
     */
    function logFailure(context, error) {
        $.writeln("[" + SCRIPT_NAME + "] " + context + ": " + error);
    }

    /**
     * アイテムに名前を設定する（失敗しても処理を続ける）
     * @param {object} item - Artboard / SymbolItem / Layer / GraphicStyle
     * @param {string} name - 設定する名前
     * @param {string} context - ログ用の文字列
     * @returns {void}
     */
    function setItemName(item, name, context) {
        try {
            item.name = name;
        } catch (nameError) {
            logFailure(context, nameError);
        }
    }

    /**
     * アイテムをコレクションの先頭へ移動する（失敗しても処理を続ける）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} item - 移動するアイテム
     * @param {string} context - ログ用の文字列
     * @returns {void}
     */
    function moveToBeginning(doc, item, context) {
        try {
            item.move(doc, ElementPlacement.PLACEATBEGINNING);
        } catch (moveError) {
            logFailure(context, moveError);
        }
    }

    /**
     * コレクションの名前と参照を控える
     * @param {object} items - コレクションまたは配列
     * @returns {{names: Array<string>, refs: Array<object>}} 名前と参照の組
     */
    function captureNamesAndRefs(items) {
        var captured = { names: [], refs: [] };
        for (var i = 0; i < items.length; i++) {
            captured.names.push(items[i].name);
            captured.refs.push(items[i]);
        }
        return captured;
    }

    /**
     * ダイアログ表示前の状態（名前・アートボードrect・各コレクションの並び順）を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} 復元用の状態
     */
    function captureOriginalState(doc) {
        var artboardState = { names: [], rects: [] };
        for (var abIdx = 0; abIdx < doc.artboards.length; abIdx++) {
            artboardState.names.push(doc.artboards[abIdx].name);
            artboardState.rects.push(doc.artboards[abIdx].artboardRect);
        }
        return {
            artboard: artboardState,
            symbol: captureNamesAndRefs(doc.symbols),
            layer: captureNamesAndRefs(doc.layers),
            graphicStyle: captureNamesAndRefs(getRenamableGraphicStyles(doc))
        };
    }

    /**
     * アートボードへ一時名を割り当てて名前の衝突を避ける
     * @param {object} artboards - アートボードのコレクション
     * @param {number} count - 対象件数
     * @returns {void}
     */
    function assignTemporaryArtboardNames(artboards, count) {
        for (var i = 0; i < count; i++) {
            setItemName(artboards[i], TEMP_ARTBOARD_PREFIX + i + "__", "temporary artboard name at " + i);
        }
    }

    /**
     * 控えておいた参照の並び順と名前を復元する
     * @param {Document} doc - 対象ドキュメント
     * @param {{names: Array<string>, refs: Array<object>}} capturedState - 控えた状態
     * @param {string} context - ログ用の種類名
     * @returns {void}
     */
    function restoreOrderAndNames(doc, capturedState, context) {
        var refs = capturedState.refs;
        /* 末尾の要素から順に先頭へ送ると、控えた並び順どおりに戻る */
        for (var reverseIdx = refs.length - 1; reverseIdx >= 0; reverseIdx--) {
            moveToBeginning(doc, refs[reverseIdx], context + " order restore at " + reverseIdx);
        }
        for (var nameIdx = 0; nameIdx < refs.length; nameIdx++) {
            setItemName(refs[nameIdx], capturedState.names[nameIdx], context + " name restore at " + nameIdx);
        }
    }

    /**
     * キャンセル時にダイアログ表示前の状態を復元する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} originalState - captureOriginalState() の戻り値
     * @returns {void}
     */
    function restoreOriginalState(doc, originalState) {
        /* アートボード：一時名で衝突を避けてから rect と名前を戻す */
        var artboards = doc.artboards;
        var artboardCount = Math.min(artboards.length, originalState.artboard.names.length);
        assignTemporaryArtboardNames(artboards, artboardCount);
        for (var abIdx = 0; abIdx < artboardCount; abIdx++) {
            try {
                artboards[abIdx].artboardRect = originalState.artboard.rects[abIdx];
                artboards[abIdx].name = originalState.artboard.names[abIdx];
            } catch (artboardError) {
                logFailure("artboard restore at " + abIdx, artboardError);
            }
        }
        restoreOrderAndNames(doc, originalState.symbol, "symbol");
        restoreOrderAndNames(doc, originalState.layer, "layer");
        restoreOrderAndNames(doc, originalState.graphicStyle, "graphic style");
        invalidateFrontmostTextCache();
    }

    // =========================================
    // 設定の比較 / Settings comparison
    // =========================================

    /**
     * 現在設定＋一覧の状態から比較用の署名を作る（［更新］→ OK の差分判定に使う）
     * @param {object} settings - 現在の設定
     * @returns {string} 比較用の署名
     */
    function buildSettingsSignature(settings) {
        var parts = [];
        var settingsKeys = ["itemType", "mode", "prefix", "suffix", "customText", "rangeMode", "rangeText", "findText", "replaceText"];
        for (var keyIdx = 0; keyIdx < settingsKeys.length; keyIdx++) {
            parts.push(settingsKeys[keyIdx] + "=" + (settings[settingsKeys[keyIdx]] || ""));
        }
        parts.push("useRegex=" + (!!settings.useRegex));

        if (settings.itemEntries) {
            for (var entryIdx = 0; entryIdx < settings.itemEntries.length; entryIdx++) {
                var entry = settings.itemEntries[entryIdx];
                parts.push([
                    "entry",
                    entryIdx,
                    entry.originalIndex,
                    entry.checked ? "1" : "0",
                    entry.userEdited ? "1" : "0",
                    entry.newName || ""
                ].join(":"));
            }
        }
        return parts.join("\n");
    }

    // =========================================
    // ダイアログ構築 / Dialog build
    // =========================================

    /**
     * リネームダイアログのUIを構築する
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} ダイアログ本体・各コントロール・操作用メソッドをまとめたオブジェクト
     */
    function createRenameDialog(doc) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        // --- 種類（ダイアログ最上段） / Item type (top of the dialog) ---
        var itemTypeRow = dialog.add("group");
        itemTypeRow.orientation = "row";
        itemTypeRow.alignment = ["fill", "top"];
        itemTypeRow.alignChildren = ["center", "center"];   /* ラジオは行の中央にまとめて置く */
        itemTypeRow.margins = 0;

        var itemTypeArtboardRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeArtboard"));
        var itemTypeSymbolRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeSymbol"));
        var itemTypeLayerRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeLayer"));
        var itemTypeGraphicStyleRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeGraphicStyle"));

        // --- コンテンツ行（左：リネーム条件／右：フィルター＋一覧） ---
        var contentRow = dialog.add("group");
        contentRow.orientation = "row";
        contentRow.alignChildren = ["left", "top"];
        contentRow.spacing = COLUMN_SPACING;

        var leftColumn = addColumnGroup(contentRow);
        var rightColumn = addColumnGroup(contentRow);

        var renameRulesPanel = addPanel(leftColumn, getLabel("panel.renameRules"));

        // --- 接頭辞 / Prefix ---
        var prefixPanel = addPanel(renameRulesPanel, getLabel("panel.prefix"));
        var prefixInput = prefixPanel.add("edittext", undefined, "");
        prefixInput.characters = PREFIX_CHARS;
        prefixInput.active = true;
        addTokenRow(prefixPanel, prefixInput, AFFIX_TOKENS, true);

        // --- 名前の基準 / Name source ---
        var nameSourcePanel = addPanel(renameRulesPanel, getLabel("panel.nameSource"));
        var originalNameRadio = nameSourcePanel.add("radiobutton", undefined, getLabel("radio.originalName"));
        originalNameRadio.alignment = "left";   /* パネル幅いっぱいに広げない */

        var customRow = nameSourcePanel.add("group");
        setupRow(customRow);
        var customRadio = customRow.add("radiobutton", undefined, getLabel("radio.custom"));
        var customInput = customRow.add("edittext", undefined, "");
        customInput.characters = CUSTOM_CHARS;
        customInput.enabled = false;

        var frontmostRadio = nameSourcePanel.add("radiobutton", undefined, getLabel("radio.frontmost"));
        frontmostRadio.alignment = "left";

        /* ラジオを全部追加してから初期値を設定する（途中で設定すると後続の追加で解除される） */
        originalNameRadio.value = true;
        customRadio.value = false;
        frontmostRadio.value = false;

        // --- 検索・置換 / Find and replace ---
        var findReplacePanel = addPanel(renameRulesPanel, getLabel("panel.findReplace"));

        var findRow = findReplacePanel.add("group");
        setupRow(findRow);
        addFixedWidthText(findRow, getLabel("fieldLabel.find"), FIELD_LABEL_WIDTH);
        var findInput = findRow.add("edittext", undefined, "");
        findInput.characters = FIND_CHARS;
        addTokenRow(findReplacePanel, findInput, FIND_PATTERN_TOKENS, false);

        var replaceRow = findReplacePanel.add("group");
        setupRow(replaceRow);
        addFixedWidthText(replaceRow, getLabel("fieldLabel.replace"), FIELD_LABEL_WIDTH);
        var replaceInput = replaceRow.add("edittext", undefined, "");
        replaceInput.characters = FIND_CHARS;

        var replaceTokenRow = addTokenRow(findReplacePanel, replaceInput, REPLACE_TOKENS, true);
        addFixedSpacer(replaceTokenRow, REGEX_GAP_WIDTH);
        var regexCheckbox = replaceTokenRow.add("checkbox", undefined, getLabel("checkbox.regex"));
        regexCheckbox.value = DEFAULT_USE_REGEX;

        // --- 接尾辞 / Suffix ---
        var suffixPanel = addPanel(renameRulesPanel, getLabel("panel.suffix"));
        var suffixInput = suffixPanel.add("edittext", undefined, "");
        suffixInput.characters = SUFFIX_CHARS;
        addTokenRow(suffixPanel, suffixInput, AFFIX_TOKENS, true);

        // --- フィルター（右カラム上段） / Filter (top of the right column) ---
        var filterPanel = addPanel(rightColumn, getLabel("panel.filter"), DENSE_SPACING);

        var rangeFilterRow = filterPanel.add("group");
        setupRow(rangeFilterRow);
        var filterAllRadio = rangeFilterRow.add("radiobutton", undefined, getLabel("radio.allItems"));
        var filterRangeRadio = rangeFilterRow.add("radiobutton", undefined, getLabel("radio.rangeItems"));
        var rangeInput = rangeFilterRow.add("edittext", undefined, "");
        rangeInput.characters = FILTER_CHARS;
        rangeInput.enabled = false;
        filterAllRadio.value = true;
        filterRangeRadio.value = false;

        var searchFilterRow = filterPanel.add("group");
        setupRow(searchFilterRow);
        var searchFilterCheckbox = searchFilterRow.add("checkbox", undefined, getLabel("checkbox.searchFilter"));
        searchFilterCheckbox.value = false;
        var searchInput = searchFilterRow.add("edittext", undefined, "");
        searchInput.characters = FILTER_CHARS;
        searchInput.helpTip = getLabel("tooltip.searchFilter");
        searchInput.enabled = false;

        // --- 一覧（右カラム下段） / Item list (bottom of the right column) ---
        var listPanel = addPanel(rightColumn, getLabel("panel.list"), DENSE_SPACING);

        var entryRowsHost = listPanel.add("group");
        entryRowsHost.orientation = "column";
        entryRowsHost.alignChildren = ["fill", "top"];
        entryRowsHost.spacing = TOKEN_SPACING;
        entryRowsHost.maximumSize.height = LIST_MAX_HEIGHT;

        var moveButtonRow = listPanel.add("group");
        setupRow(moveButtonRow, "center", TOKEN_SPACING);
        moveButtonRow.margins = [0, 10, 0, 0];
        var moveToTopButton = moveButtonRow.add("button", undefined, getLabel("button.moveTop"));
        var moveUpButton = moveButtonRow.add("button", undefined, getLabel("button.moveUp"));
        var moveDownButton = moveButtonRow.add("button", undefined, getLabel("button.moveDown"));
        var moveToBottomButton = moveButtonRow.add("button", undefined, getLabel("button.moveBottom"));
        moveToTopButton.preferredSize = [MOVE_BUTTON_SIZE[0] + 4, MOVE_BUTTON_SIZE[1]];
        moveUpButton.preferredSize = MOVE_BUTTON_SIZE;
        moveDownButton.preferredSize = MOVE_BUTTON_SIZE;
        moveToBottomButton.preferredSize = [MOVE_BUTTON_SIZE[0] + 4, MOVE_BUTTON_SIZE[1]];

        // --- ボタンエリア（左：更新／右：キャンセル・OK） ---
        var buttonArea = dialog.add("group");
        setupRow(buttonArea, "fill");
        var refreshButton = buttonArea.add("button", undefined, getLabel("button.refresh"));
        refreshButton.alignment = ["left", "center"];
        var buttonSpacer = buttonArea.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var cancelButton = buttonArea.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        cancelButton.alignment = ["right", "center"];
        var okButton = buttonArea.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        okButton.alignment = ["right", "center"];

        // =====================================
        // 一覧の状態管理 / Item list state
        // =====================================

        var currentItemType = DEFAULT_ITEM_TYPE;
        var entryRows = [];

        /* bindDialogEvents から注入されるコールバック / Callbacks injected by bindDialogEvents */
        var requestPreviewUpdate = null;
        var itemTypeChangeCallback = null;
        var lastCommittedSignature = null;
        var skipApplyOnOk = false;

        /**
         * 種類に対応するエントリ配列を作る
         * @param {string} itemType - "artboard" / "symbol" / "layer" / "graphicstyle"
         * @returns {Array<object>} エントリ配列
         */
        function buildEntriesForItemType(itemType) {
            var items = getDocumentItems(doc, itemType);
            var entries = [];
            for (var i = 0; i < items.length; i++) {
                var entry = {
                    originalIndex: i,
                    name: items[i].name,
                    newName: items[i].name,
                    checked: false,
                    userEdited: false
                };
                if (itemType === "artboard") {
                    entry.rect = items[i].artboardRect;
                }
                entries.push(entry);
            }
            return entries;
        }

        var itemEntries = buildEntriesForItemType(currentItemType);

        /**
         * ダイアログ各コントロールから設定オブジェクトを作る
         * rangeMode / rangeText は一覧のチェック状態から実効値を導出する
         * @returns {object} 現在の設定
         */
        function readSettings() {
            var settings = {
                mode: frontmostRadio.value ? "frontmost" : (originalNameRadio.value ? "original" : "custom"),
                itemType: currentItemType,
                prefix: prefixInput.text,
                suffix: suffixInput.text,
                customText: customInput.text,
                findText: findInput.text,
                replaceText: replaceInput.text,
                useRegex: regexCheckbox.value
            };

            /* rangeMode / rangeText は「すべて／指定範囲」ラジオではなく、一覧のチェック状態から決める
               （検索フィルターや手動チェックの結果もそのまま実効範囲になる）
               リネーム実行側は canvas 順で動くので、ここは表示位置ではなく originalIndex を基準にする */
            var checkedIndices = [];
            for (var i = 0; i < itemEntries.length; i++) {
                if (itemEntries[i].checked) checkedIndices.push(itemEntries[i].originalIndex);
            }
            var range = deriveRangeSettings(checkedIndices, itemEntries.length);
            settings.rangeMode = range.rangeMode;
            settings.rangeText = range.rangeText;
            return settings;
        }

        /**
         * 編集中の一覧を反映してから設定を取得する
         * @returns {object} itemEntries 付きの設定
         */
        function readCommittedSettings() {
            syncEditingValues();
            var settings = readSettings();
            settings.itemEntries = itemEntries;
            return settings;
        }

        /**
         * 一覧のチェックと入力内容をエントリへ書き戻す
         * @returns {void}
         */
        function syncEditingValues() {
            for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                var row = entryRows[rowIdx];
                itemEntries[row.dataIndex].checked = row.checkbox.value;
                if (row.checkbox.value) {
                    itemEntries[row.dataIndex].newName = row.newNameField.text;
                }
            }
        }

        /**
         * エントリのチェック状態を更新する（OFF にしたら手動編集を破棄する）
         * @param {number} entryIdx - エントリの位置
         * @param {boolean} isChecked - 新しいチェック状態
         * @returns {void}
         */
        function setEntryChecked(entryIdx, isChecked) {
            var entry = itemEntries[entryIdx];
            entry.checked = isChecked;
            if (!isChecked) {
                entry.newName = entry.name;
                entry.userEdited = false;
            }
        }

        /**
         * 各行の表示をエントリの状態に合わせ直す
         * @returns {void}
         */
        function syncRowsToEntries() {
            for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                var row = entryRows[rowIdx];
                var entry = itemEntries[row.dataIndex];
                row.checkbox.value = entry.checked;
                row.newNameField.enabled = entry.checked;
                if (!entry.checked) row.newNameField.text = entry.name;
            }
        }

        /**
         * Option+クリック時の一括切り替え
         * 全行 ON のときはクリック行だけを残し（孤立化）、それ以外はクリック値で一括切替する
         * @param {number} clickedIdx - クリックされた行の位置
         * @param {boolean} newCheckedValue - クリック後のチェック値
         * @returns {void}
         */
        function applyOptionClickToggle(clickedIdx, newCheckedValue) {
            var allWereOn = itemEntries.length > 0;
            for (var preIdx = 0; preIdx < itemEntries.length; preIdx++) {
                if (!itemEntries[preIdx].checked) { allWereOn = false; break; }
            }
            for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                setEntryChecked(entryIdx, allWereOn ? (entryIdx === clickedIdx) : newCheckedValue);
            }
            syncRowsToEntries();
        }

        /**
         * チェック状態から「すべて／指定範囲」へ逆同期する
         * @returns {void}
         */
        function syncFilterFromCheckboxes() {
            var checkedPositions = [];
            for (var i = 0; i < itemEntries.length; i++) {
                if (itemEntries[i].checked) checkedPositions.push(i);
            }
            /* 「指定範囲」欄は一覧の表示位置を基準にする（applyFilterToCheckboxes も表示位置で読む） */
            var range = deriveRangeSettings(checkedPositions, itemEntries.length);
            var isAllChecked = (range.rangeMode === "all");
            filterAllRadio.value = isAllChecked;
            filterRangeRadio.value = !isAllChecked;
            rangeInput.enabled = !isAllChecked;
            if (!isAllChecked) rangeInput.text = range.rangeText;
        }

        // =====================================
        // 並び替え操作 / Reorder actions
        // =====================================

        /**
         * 2つのエントリを入れ替える
         * @param {number} indexA - 位置A
         * @param {number} indexB - 位置B
         * @returns {void}
         */
        function swapEntries(indexA, indexB) {
            var temp = itemEntries[indexA];
            itemEntries[indexA] = itemEntries[indexB];
            itemEntries[indexB] = temp;
        }

        /**
         * チェック済みと未チェックでエントリを2分割する
         * @returns {{checked: Array<object>, unchecked: Array<object>}} 分割結果
         */
        function partitionEntriesByChecked() {
            var partition = { checked: [], unchecked: [] };
            for (var i = 0; i < itemEntries.length; i++) {
                if (itemEntries[i].checked) partition.checked.push(itemEntries[i]);
                else partition.unchecked.push(itemEntries[i]);
            }
            return partition;
        }

        /**
         * チェック行を1つ上へ移動する
         * @returns {void}
         */
        function moveCheckedUp() {
            syncEditingValues();
            for (var i = 1; i < itemEntries.length; i++) {
                if (itemEntries[i].checked && !itemEntries[i - 1].checked) swapEntries(i, i - 1);
            }
            refreshReorderRows();
        }

        /**
         * チェック行を1つ下へ移動する
         * @returns {void}
         */
        function moveCheckedDown() {
            syncEditingValues();
            for (var i = itemEntries.length - 2; i >= 0; i--) {
                if (itemEntries[i].checked && !itemEntries[i + 1].checked) swapEntries(i, i + 1);
            }
            refreshReorderRows();
        }

        /**
         * チェック行を先頭へまとめる
         * @returns {void}
         */
        function moveCheckedToTop() {
            syncEditingValues();
            var partition = partitionEntriesByChecked();
            itemEntries = partition.checked.concat(partition.unchecked);
            refreshReorderRows();
        }

        /**
         * チェック行を末尾へまとめる
         * @returns {void}
         */
        function moveCheckedToBottom() {
            syncEditingValues();
            var partition = partitionEntriesByChecked();
            itemEntries = partition.unchecked.concat(partition.checked);
            refreshReorderRows();
        }

        /**
         * 上へ動かせる行があるか
         * @returns {boolean} 動かせれば true
         */
        function canMoveUp() {
            for (var i = 1; i < itemEntries.length; i++) {
                if (itemEntries[i].checked && !itemEntries[i - 1].checked) return true;
            }
            return false;
        }

        /**
         * 下へ動かせる行があるか
         * @returns {boolean} 動かせれば true
         */
        function canMoveDown() {
            for (var i = 0; i < itemEntries.length - 1; i++) {
                if (itemEntries[i].checked && !itemEntries[i + 1].checked) return true;
            }
            return false;
        }

        /**
         * 並び替えボタンの活性状態を更新する
         * @returns {void}
         */
        function updateMoveButtonsState() {
            var canMoveUpwards = canMoveUp();
            var canMoveDownwards = canMoveDown();
            moveToTopButton.enabled = canMoveUpwards;
            moveUpButton.enabled = canMoveUpwards;
            moveDownButton.enabled = canMoveDownwards;
            moveToBottomButton.enabled = canMoveDownwards;
        }

        moveToTopButton.onClick = function () { moveCheckedToTop(); };
        moveUpButton.onClick = function () { moveCheckedUp(); };
        moveDownButton.onClick = function () { moveCheckedDown(); };
        moveToBottomButton.onClick = function () { moveCheckedToBottom(); };

        // =====================================
        // 一覧の描画 / Item list rendering
        // =====================================

        /**
         * 一覧の見出し行を作る
         * @returns {void}
         */
        function addListHeaderRow() {
            var headerRow = entryRowsHost.add("group");
            setupRow(headerRow, "left", DENSE_SPACING);
            addFixedWidthText(headerRow, getLabel("listHeader.order"), LIST_COLUMN_WIDTHS.order);
            addFixedWidthText(headerRow, getLabel("listHeader.select"), LIST_COLUMN_WIDTHS.select);
            addFixedWidthText(headerRow, getLabel("listHeader.currentName"), LIST_COLUMN_WIDTHS.currentName);
            addFixedWidthText(headerRow, "→", LIST_COLUMN_WIDTHS.arrow);
            addFixedWidthText(headerRow, getLabel("listHeader.newName"), LIST_COLUMN_WIDTHS.newName);
        }

        /**
         * 一覧の1行（順・チェック・現在の名前・新しい名前）を作る
         * @param {number} idx - エントリの位置
         * @returns {void}
         */
        function addEntryRow(idx) {
            var row = entryRowsHost.add("group");
            setupRow(row, "left", DENSE_SPACING);

            addFixedWidthText(row, (idx + 1) + "", LIST_COLUMN_WIDTHS.order);

            var rowCheckbox = row.add("checkbox", undefined, "");
            rowCheckbox.value = itemEntries[idx].checked;
            rowCheckbox.preferredSize.width = LIST_COLUMN_WIDTHS.select;

            var currentNameLabel = addFixedWidthText(row, itemEntries[idx].name, LIST_COLUMN_WIDTHS.currentName);
            currentNameLabel.helpTip = itemEntries[idx].name;

            addFixedWidthText(row, "→", LIST_COLUMN_WIDTHS.arrow);

            var newNameField = row.add("edittext", undefined, itemEntries[idx].newName);
            newNameField.preferredSize.width = LIST_COLUMN_WIDTHS.newName;
            newNameField.enabled = itemEntries[idx].checked;

            rowCheckbox.onClick = function () {
                if (isOptionKeyHeld()) {
                    applyOptionClickToggle(idx, rowCheckbox.value);
                } else {
                    setEntryChecked(idx, rowCheckbox.value);
                    newNameField.enabled = rowCheckbox.value;
                    if (!rowCheckbox.value) newNameField.text = itemEntries[idx].name;
                }
                syncFilterFromCheckboxes();
                if (requestPreviewUpdate) requestPreviewUpdate();
                updateMoveButtonsState();
            };

            newNameField.onChange = function () {
                itemEntries[idx].newName = newNameField.text;
                itemEntries[idx].userEdited = true;
            };

            entryRows.push({
                group: row,
                checkbox: rowCheckbox,
                currentNameLabel: currentNameLabel,
                newNameField: newNameField,
                dataIndex: idx
            });
        }

        /**
         * 一覧を作り直す（並び替え・種類切替のあとに呼ぶ）
         * @returns {void}
         */
        function refreshReorderRows() {
            while (entryRowsHost.children.length > 0) {
                entryRowsHost.remove(entryRowsHost.children[0]);
            }
            entryRows = [];

            addListHeaderRow();
            for (var entryIndex = 0; entryIndex < itemEntries.length; entryIndex++) {
                addEntryRow(entryIndex);
            }

            updateMoveButtonsState();
            entryRowsHost.layout.layout(true);
            dialog.layout.layout(true);
        }

        /**
         * ［更新］確定後にエントリを canvas の現状へ再ベースライン化する
         * userEdited を解除して、以降の接頭辞・接尾辞の変更をプレビューへ反映できるようにする
         * @returns {void}
         */
        function rebaselineEntriesAfterCommit() {
            var items = getDocumentItems(doc, currentItemType);

            /* 確定した結果アイテム数が変わることがある（例：グラフィックスタイル名が `[...]` になり
               RESERVED_STYLE_NAME に引っかかって対象から外れる）。ずれたまま参照を続けると
               範囲外アクセスになるので、件数が変わったら一覧ごと作り直す */
            if (items.length !== itemEntries.length) {
                itemEntries = buildEntriesForItemType(currentItemType);
                refreshReorderRows();
                return;
            }

            for (var i = 0; i < itemEntries.length && i < items.length; i++) {
                itemEntries[i].originalIndex = i;
                itemEntries[i].name = items[i].name;
                itemEntries[i].newName = items[i].name;
                itemEntries[i].userEdited = false;
                if (currentItemType === "artboard") {
                    itemEntries[i].rect = items[i].artboardRect;
                }
            }
        }

        /**
         * 種類を切り替える：エントリを作り直し、フィルターと名前の基準をリセットする
         * @param {string} itemType - "artboard" / "symbol" / "layer" / "graphicstyle"
         * @returns {void}
         */
        function setItemType(itemType) {
            currentItemType = itemType;
            itemEntries = buildEntriesForItemType(itemType);
            lastCommittedSignature = null;
            skipApplyOnOk = false;

            /* 種類が変わると名前体系も変わるため、検索フィルターは OFF にしてクリアする
               （前の種類向けの語がそのまま効いてしまうのを防ぐ） */
            searchFilterCheckbox.value = false;
            searchInput.text = "";
            searchInput.enabled = false;

            filterAllRadio.value = true;
            filterRangeRadio.value = false;
            rangeInput.text = "";
            filterAllRadio.enabled = true;
            filterRangeRadio.enabled = true;
            rangeInput.enabled = false;

            /* 「最前面のテキスト」はアートボードのときだけ有効。ほかの種類では「指定」を使う */
            frontmostRadio.enabled = (itemType === "artboard");
            if (!frontmostRadio.enabled && frontmostRadio.value) {
                frontmostRadio.value = false;
                customRadio.value = true;
            }
            customInput.enabled = customRadio.value;

            dialog.layout.layout(true);
            refreshReorderRows();
            if (itemTypeChangeCallback) itemTypeChangeCallback();
        }

        okButton.onClick = function () {
            syncEditingValues();
            for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                if (itemEntries[entryIdx].checked && itemEntries[entryIdx].newName === "") {
                    alert(getLabel("alert.emptyName", { n: entryIdx + 1 }));
                    return;
                }
            }

            /* ［更新］以降に何も変わっていなければ、OK では何もしない（二重適用の防止） */
            skipApplyOnOk = false;
            if (lastCommittedSignature !== null) {
                var okSettings = readSettings();
                okSettings.itemEntries = itemEntries;
                if (buildSettingsSignature(okSettings) === lastCommittedSignature) {
                    skipApplyOnOk = true;
                }
            }
            dialog.close(1);
        };

        /* 全UI構築後に初期状態をそろえる。ラジオを立てるだけでなく setItemType() を通すことで、
           「最前面のテキスト」の有効・無効やフィルターの初期化も既定の種類に追随する */
        itemTypeArtboardRadio.value = (DEFAULT_ITEM_TYPE === "artboard");
        itemTypeSymbolRadio.value = (DEFAULT_ITEM_TYPE === "symbol");
        itemTypeLayerRadio.value = (DEFAULT_ITEM_TYPE === "layer");
        itemTypeGraphicStyleRadio.value = (DEFAULT_ITEM_TYPE === "graphicstyle");
        setItemType(DEFAULT_ITEM_TYPE);

        return {
            dialog: dialog,
            prefixInput: prefixInput,
            suffixInput: suffixInput,
            frontmostRadio: frontmostRadio,
            originalNameRadio: originalNameRadio,
            customRadio: customRadio,
            customInput: customInput,
            filterAllRadio: filterAllRadio,
            filterRangeRadio: filterRangeRadio,
            rangeInput: rangeInput,
            searchFilterCheckbox: searchFilterCheckbox,
            searchInput: searchInput,
            findInput: findInput,
            replaceInput: replaceInput,
            regexCheckbox: regexCheckbox,
            refreshButton: refreshButton,
            itemTypeArtboardRadio: itemTypeArtboardRadio,
            itemTypeSymbolRadio: itemTypeSymbolRadio,
            itemTypeLayerRadio: itemTypeLayerRadio,
            itemTypeGraphicStyleRadio: itemTypeGraphicStyleRadio,

            setItemType: setItemType,

            /**
             * 種類切替後に呼ぶコールバックを登録する
             * @param {function} callback - 切替後に実行する処理
             * @returns {void}
             */
            setItemTypeChangeCallback: function (callback) { itemTypeChangeCallback = callback; },

            /**
             * チェックボックス操作からプレビュー更新を呼ぶためのコールバックを登録する
             * @param {function} callback - プレビュー更新処理
             * @returns {void}
             */
            setRequestPreviewUpdate: function (callback) { requestPreviewUpdate = callback; },

            /**
             * ［更新］で確定した設定の署名を保持する（OK 時の差分判定に使う）
             * @param {string} signature - buildSettingsSignature() の戻り値
             * @returns {void}
             */
            setLastCommittedSignature: function (signature) { lastCommittedSignature = signature; },

            /**
             * OK 時に適用をスキップしてよいか（［更新］以降に差分がないか）
             * @returns {boolean} スキップしてよければ true
             */
            getSkipApplyOnOk: function () { return skipApplyOnOk; },

            /**
             * 現在のエントリ配列を返す
             * @returns {Array<object>} エントリ配列
             */
            getItemEntries: function () { return itemEntries; },

            readSettings: readSettings,
            readCommittedSettings: readCommittedSettings,
            syncEditingValues: syncEditingValues,
            rebaselineEntriesAfterCommit: rebaselineEntriesAfterCommit,

            /**
             * ［更新］確定後に、一覧の表示を canvas の現在名へそろえる
             * @returns {void}
             */
            syncReorderRowsToCurrentNames: function () {
                var items = getDocumentItems(doc, currentItemType);
                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var row = entryRows[rowIdx];
                    var entry = itemEntries[row.dataIndex];
                    /* 確定でコレクションが縮むことがあるため、範囲外は現在のエントリ名で代替する */
                    var currentItem = items[entry.originalIndex];
                    var currentName = currentItem ? currentItem.name : entry.name;

                    row.currentNameLabel.text = currentName;
                    row.currentNameLabel.helpTip = currentName;
                    row.newNameField.text = currentName;
                    row.newNameField.enabled = entry.checked;

                    entry.name = currentName;
                    entry.newName = currentName;
                    entry.userEdited = false;
                }
            },

            /**
             * 未確定のプレビュー名を一覧の「新しい名前」欄へ反映する
             * @param {Array<string>} previewNames - 元の並び順でのプレビュー名
             * @returns {void}
             */
            syncPreviewToReorderRows: function (previewNames) {
                var items = getDocumentItems(doc, currentItemType);
                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var row = entryRows[rowIdx];
                    var entry = itemEntries[row.dataIndex];
                    /* 確定でコレクションが縮むことがあるため、範囲外は現在のエントリ名で代替する */
                    var currentItem = items[entry.originalIndex];
                    var currentName = currentItem ? currentItem.name : entry.name;

                    /* 「現在の名前」列は canvas の現状（［更新］後は確定後の名前）を出す */
                    row.currentNameLabel.text = currentName;
                    row.currentNameLabel.helpTip = currentName;

                    /* 「新しい名前」列は未確定プレビュー。手動編集した行は上書きしない */
                    if (entry.userEdited) continue;
                    var previewName = (previewNames && previewNames[entry.originalIndex] != null)
                        ? previewNames[entry.originalIndex]
                        : currentName;
                    row.newNameField.text = previewName;
                    entry.newName = previewName;
                }
            },

            /**
             * フィルター条件を一覧のチェックボックスへ反映する
             * @param {string} rangeMode - "all" または "numbered"
             * @param {string} rangeText - "1-3,5" 形式の範囲文字列
             * @returns {void}
             */
            applyFilterToCheckboxes: function (rangeMode, rangeText) {
                /* 検索フィルターが ON かつテキストが空でないときは、検索条件に見合うものだけをチェックする
                   （「すべて／指定範囲」は無視する） */
                var filterActive = searchFilterCheckbox.value && searchInput.text !== "";
                var lowerQuery = filterActive ? searchInput.text.toLowerCase() : null;

                var isFilteredIndex = {};
                if (!filterActive) {
                    /* 「指定範囲」の番号は一覧の「順」列＝表示位置を指す。書き出す側
                       （syncFilterFromCheckboxes）が表示位置なので、読む側も originalIndex ではなく
                       表示位置で合わせる。合わせないと並び替え後にチェックが別の行へ飛ぶ */
                    var filteredIndices = getRangeItemIndices(itemEntries.length, rangeMode, rangeText);
                    for (var filteredIdx = 0; filteredIdx < filteredIndices.length; filteredIdx++) {
                        isFilteredIndex[filteredIndices[filteredIdx]] = true;
                    }
                }

                for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                    var isChecked = filterActive
                        ? itemEntries[entryIdx].name.toLowerCase().indexOf(lowerQuery) !== -1
                        : !!isFilteredIndex[entryIdx];
                    setEntryChecked(entryIdx, isChecked);
                }

                syncRowsToEntries();
                updateMoveButtonsState();
            }
        };
    }

    // =========================================
    // ダイアログイベント / Dialog events
    // =========================================

    /**
     * ダイアログ各コントロールにイベントハンドラを設定する
     * @param {object} dialogUI - createRenameDialog() の戻り値
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function bindDialogEvents(dialogUI, doc) {

        /* 複数フィールドをまとめて確定するあいだ、中間状態のプレビュー計算を止めるフラグ
           Suppresses intermediate preview passes while several fields are committed at once */
        var suppressPreview = false;

        /**
         * canvas には触れず、未確定のプレビュー名を計算して一覧を更新する
         * @returns {void}
         */
        function updatePreview() {
            if (suppressPreview) return;
            dialogUI.syncPreviewToReorderRows(computePreviewNames(doc, dialogUI.readSettings()));
        }

        /**
         * 現在のフィルター設定を一覧のチェックボックスへ反映する
         * @returns {void}
         */
        function applyCurrentFilter() {
            if (dialogUI.filterAllRadio.value) {
                dialogUI.applyFilterToCheckboxes("all", "");
            } else {
                dialogUI.applyFilterToCheckboxes("numbered", dialogUI.rangeInput.text);
            }
        }

        /**
         * 「すべて／指定範囲」の切り替えを反映する
         * @returns {void}
         */
        function syncFilterInput() {
            dialogUI.rangeInput.enabled = dialogUI.filterRangeRadio.value;
            applyCurrentFilter();
            updatePreview();
        }

        /**
         * ［更新］：現在の設定・並び替え・手動編集名を canvas へ確定する
         * @returns {void}
         */
        function commitCurrentSettings() {
            var settings = dialogUI.readCommittedSettings();
            var committed;
            if (hasReorderOrRename(settings.itemEntries)) {
                applyReorderAndRename(doc, settings.itemEntries, settings);
                committed = true;
            } else {
                committed = executeRename(doc, settings, { silent: true });
            }

            if (!committed) {
                updatePreview();
                return;
            }

            app.redraw();
            dialogUI.rebaselineEntriesAfterCommit();
            dialogUI.syncReorderRowsToCurrentNames();
            dialogUI.setLastCommittedSignature(buildSettingsSignature(dialogUI.readCommittedSettings()));
        }

        dialogUI.filterAllRadio.onClick = function () {
            dialogUI.filterRangeRadio.value = false;
            syncFilterInput();
        };
        dialogUI.filterRangeRadio.onClick = function () {
            dialogUI.filterAllRadio.value = false;
            syncFilterInput();
        };

        /* 種類切替後：「すべて」フィルターを適用してプレビューを更新する */
        dialogUI.setItemTypeChangeCallback(function () {
            dialogUI.applyFilterToCheckboxes("all", "");
            updatePreview();
        });

        /**
         * 種類のラジオを選び直して一覧を作り直す
         * @param {string} itemType - "artboard" / "symbol" / "layer" / "graphicstyle"
         * @returns {void}
         */
        function selectItemTypeRadio(itemType) {
            dialogUI.itemTypeArtboardRadio.value = (itemType === "artboard");
            dialogUI.itemTypeSymbolRadio.value = (itemType === "symbol");
            dialogUI.itemTypeLayerRadio.value = (itemType === "layer");
            dialogUI.itemTypeGraphicStyleRadio.value = (itemType === "graphicstyle");
            dialogUI.setItemType(itemType);
        }
        dialogUI.itemTypeArtboardRadio.onClick = function () { selectItemTypeRadio("artboard"); };
        dialogUI.itemTypeSymbolRadio.onClick = function () { selectItemTypeRadio("symbol"); };
        dialogUI.itemTypeLayerRadio.onClick = function () { selectItemTypeRadio("layer"); };
        dialogUI.itemTypeGraphicStyleRadio.onClick = function () { selectItemTypeRadio("graphicstyle"); };

        var sourceRadios = [dialogUI.originalNameRadio, dialogUI.customRadio, dialogUI.frontmostRadio];

        /**
         * 「名前の基準」のラジオを排他選択する
         * @param {RadioButton} selectedRadio - 選択されたラジオ
         * @returns {void}
         */
        function selectSourceRadio(selectedRadio) {
            for (var i = 0; i < sourceRadios.length; i++) {
                sourceRadios[i].value = (sourceRadios[i] === selectedRadio);
            }
            dialogUI.customInput.enabled = dialogUI.customRadio.value;
            if (dialogUI.customRadio.value) focusField(dialogUI.customInput);
            updatePreview();
        }
        dialogUI.originalNameRadio.onClick = function () { selectSourceRadio(dialogUI.originalNameRadio); };
        dialogUI.customRadio.onClick = function () { selectSourceRadio(dialogUI.customRadio); };
        dialogUI.frontmostRadio.onClick = function () { selectSourceRadio(dialogUI.frontmostRadio); };
        dialogUI.customInput.onChange = function () {
            if (dialogUI.customRadio.value) updatePreview();
        };

        dialogUI.prefixInput.onChange = updatePreview;
        dialogUI.suffixInput.onChange = updatePreview;
        dialogUI.findInput.onChange = updatePreview;
        dialogUI.replaceInput.onChange = updatePreview;
        dialogUI.regexCheckbox.onClick = updatePreview;

        dialogUI.rangeInput.onChange = function () {
            if (!dialogUI.filterRangeRadio.value) return;
            dialogUI.applyFilterToCheckboxes("numbered", dialogUI.rangeInput.text);
            updatePreview();
        };

        dialogUI.searchFilterCheckbox.onClick = function () {
            var filterOn = dialogUI.searchFilterCheckbox.value;
            dialogUI.searchInput.enabled = filterOn;
            /* フィルター ON 中は「すべて／指定範囲」を無効化し、どちらが効いているのか紛れないようにする */
            dialogUI.filterAllRadio.enabled = !filterOn;
            dialogUI.filterRangeRadio.enabled = !filterOn;
            dialogUI.rangeInput.enabled = !filterOn && dialogUI.filterRangeRadio.value;
            if (filterOn) focusField(dialogUI.searchInput);
            applyCurrentFilter();
            updatePreview();
        };
        dialogUI.searchInput.onChange = function () {
            if (!dialogUI.searchFilterCheckbox.value) return;
            applyCurrentFilter();
            updatePreview();
        };

        dialogUI.refreshButton.onClick = function () {
            /* フォーカスが外れていない edittext も確定させるため、関連フィールドの onChange を一括で発火する
               Force onChange on all edittexts so pending edits commit even without losing focus
               各 onChange はプレビューを走らせるので、確定前の中間状態では抑制する
               （抑制しないと1クリックで7回、「最前面のテキスト」では書類全走査が7回起きる） */
            var pendingFields = [
                dialogUI.prefixInput,
                dialogUI.suffixInput,
                dialogUI.customInput,
                dialogUI.findInput,
                dialogUI.replaceInput,
                dialogUI.rangeInput,
                dialogUI.searchInput
            ];
            suppressPreview = true;
            for (var pendingIdx = 0; pendingIdx < pendingFields.length; pendingIdx++) {
                try {
                    pendingFields[pendingIdx].notify("onChange");
                } catch (notifyError) {
                    logFailure("notify onChange", notifyError);
                }
            }
            suppressPreview = false;

            commitCurrentSettings();
        };

        /* チェックボックス操作からプレビュー更新を呼べるように注入する */
        dialogUI.setRequestPreviewUpdate(updatePreview);

        /* 初期同期：フィルター設定 → チェックボックス → プレビュー */
        applyCurrentFilter();
        updatePreview();
    }

    /**
     * ダイアログを表示して、確定された設定を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {object|null} 確定した設定（キャンセル時は null）
     */
    function showRenameDialog(doc) {
        var dialogUI = createRenameDialog(doc);
        bindDialogEvents(dialogUI, doc);

        if (dialogUI.dialog.show() !== 1) return null;

        var settings = dialogUI.readSettings();
        settings.itemEntries = dialogUI.getItemEntries();
        if (dialogUI.getSkipApplyOnOk()) settings.skipApplyOnOk = true;
        return settings;
    }

    // =========================================
    // 並び替えと適用 / Reorder and apply
    // =========================================

    /**
     * 並び替えまたは手動上書きが残っているか判定する
     * @param {Array<object>} itemEntries - エントリ配列
     * @returns {boolean} どちらかがあれば true
     */
    function hasReorderOrRename(itemEntries) {
        if (!itemEntries) return false;
        for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
            if (itemEntries[entryIdx].originalIndex !== entryIdx) return true;
            if (itemEntries[entryIdx].userEdited) return true;
        }
        return false;
    }

    /**
     * 参照を末尾から先頭へ送って、エントリの並び順に合わせる
     * @param {Document} doc - 対象ドキュメント
     * @param {object} items - 並び替え前のコレクション
     * @param {Array<object>} itemEntries - 新しい並び順のエントリ配列
     * @param {string} context - ログ用の種類名
     * @returns {void}
     */
    function reorderByMove(doc, items, itemEntries, context) {
        var refs = [];
        for (var initIdx = 0; initIdx < items.length; initIdx++) refs.push(items[initIdx]);
        for (var reverseIdx = itemEntries.length - 1; reverseIdx >= 0; reverseIdx--) {
            moveToBeginning(doc, refs[itemEntries[reverseIdx].originalIndex], context + " move at entry " + reverseIdx);
        }
    }

    /**
     * 並び替え後の位置を基準にチェック範囲を作り直した設定を返す
     * @param {object} settings - 元の設定
     * @param {Array<object>} itemEntries - 並び替え後のエントリ配列
     * @param {number} itemCount - アイテム総数
     * @returns {object} rangeMode / rangeText を差し替えた設定
     */
    function buildReorderedSettings(settings, itemEntries, itemCount) {
        var reorderedSettings = {};
        for (var settingsKey in settings) {
            if (settings.hasOwnProperty(settingsKey)) {
                reorderedSettings[settingsKey] = settings[settingsKey];
            }
        }

        var checkedPositions = [];
        for (var checkedPos = 0; checkedPos < itemEntries.length; checkedPos++) {
            if (itemEntries[checkedPos].checked) checkedPositions.push(checkedPos);
        }

        /* 並び替え済みなので表示位置＝canvas 順。ここは表示位置を基準にする */
        var range = deriveRangeSettings(checkedPositions, itemCount);
        reorderedSettings.rangeMode = range.rangeMode;
        reorderedSettings.rangeText = range.rangeText;
        return reorderedSettings;
    }

    /**
     * 並び替え結果と手動リネームをアイテムへ適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<object>} itemEntries - エントリ配列
     * @param {object} settings - 現在の設定
     * @returns {void}
     */
    function applyReorderAndRename(doc, itemEntries, settings) {
        if (!hasReorderOrRename(itemEntries)) return;

        /* アートボードの rect が入れ替わると最前面テキストの対応も変わるため、走査をやり直させる */
        invalidateFrontmostTextCache();

        var itemType = (settings && settings.itemType) || "artboard";
        var items = getDocumentItems(doc, itemType);
        var itemCount = items.length;

        /* ユーザーが手動で上書きした行を、新しい位置で控える */
        var userOverridesByNewPosition = {};
        for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
            if (itemEntries[entryIdx].userEdited && itemEntries[entryIdx].checked) {
                userOverridesByNewPosition[entryIdx] = itemEntries[entryIdx].newName;
            }
        }

        /* 並び替え前の canvas 名を originalIndex で控える（［更新］済みの名前を保つ） */
        var currentNamesByOriginalIndex = [];
        for (var origIdx = 0; origIdx < itemCount; origIdx++) {
            currentNamesByOriginalIndex.push(items[origIdx].name);
        }

        if (itemType === "artboard") {
            /* rect と現在名を新しい位置へ並べ替える（一時名で衝突を避ける）
               一時名を付ける件数と並べ替える件数は必ず同じにする。ずれると try の外にある
               artboardRect 代入で落ち、全アートボードが一時名のまま残る */
            var reorderCount = Math.min(itemEntries.length, itemCount);
            assignTemporaryArtboardNames(items, reorderCount);
            for (var newPos = 0; newPos < reorderCount; newPos++) {
                items[newPos].artboardRect = itemEntries[newPos].rect;
                setItemName(items[newPos], currentNamesByOriginalIndex[itemEntries[newPos].originalIndex], "artboard reorder at " + newPos);
            }
        } else {
            /* シンボル・レイヤー・グラフィックスタイルは安定参照を move() で並べ替える */
            reorderByMove(doc, items, itemEntries, itemType);
            items = getDocumentItems(doc, itemType);
        }

        /* 並び替え後の位置を基準にチェック範囲を作り直してからリネームする */
        if (settings) {
            executeRename(doc, buildReorderedSettings(settings, itemEntries, itemCount), { silent: true });
        }

        /* 手動上書きを適用する（move() 後の最新参照に追随するため items を取り直す） */
        items = getDocumentItems(doc, itemType);
        for (var posKey in userOverridesByNewPosition) {
            if (!userOverridesByNewPosition.hasOwnProperty(posKey)) continue;
            var positionIndex = parseInt(posKey, 10);
            if (positionIndex >= 0 && positionIndex < items.length) {
                setItemName(items[positionIndex], userOverridesByNewPosition[posKey], "manual override at " + positionIndex);
            }
        }
    }

    // =========================================
    // リネーム実行とプレビュー / Rename execution and preview
    // =========================================

    /**
     * リネームに使える入力が何もないか判定する
     * @param {object} settings - 現在の設定
     * @returns {boolean} 何も入力されていなければ true
     */
    function hasNoRenameInput(settings) {
        return settings.mode === "custom"
            && !settings.customText
            && !settings.prefix
            && !settings.suffix
            && !settings.findText;
    }

    /**
     * 名前の基準ごとに、各アイテムのベース文字列を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {object} items - 対象アイテムのコレクション
     * @param {object} settings - 現在の設定
     * @param {string} itemType - 正規化済みの種類（呼び出し側で "artboard" に既定化したもの）
     * @returns {object} インデックスをキーにした文字列配列のマップ
     */
    function buildItemTextMap(doc, items, settings, itemType) {
        var itemTextMap = {};
        if (settings.mode === "frontmost" && itemType === "artboard") {
            return mapTextFramesToArtboards(getCachedFrontmostTextFrames(doc), items);
        }
        for (var itemIdx = 0; itemIdx < items.length; itemIdx++) {
            if (settings.mode === "original") {
                itemTextMap[itemIdx] = [items[itemIdx].name];
            } else if (settings.mode === "custom") {
                itemTextMap[itemIdx] = [settings.customText];
            }
        }
        return itemTextMap;
    }

    /**
     * 設定に従ってアイテムをリネームし canvas を更新する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} settings - 現在の設定
     * @param {object} [options] - {silent: true} で警告を出さない
     * @returns {boolean} リネームを実行したら true
     */
    function executeRename(doc, settings, options) {
        if (hasNoRenameInput(settings)) {
            if (!(options && options.silent)) alert(getLabel("alert.needSettings"));
            return false;
        }

        var itemType = settings.itemType || "artboard";
        var items = getDocumentItems(doc, itemType);
        var itemTextMap = buildItemTextMap(doc, items, settings, itemType);
        var selectedIndices = getRangeItemIndices(items.length, settings.rangeMode, settings.rangeText);
        var renamePlan = buildRenamePlan(items, itemTextMap, settings, selectedIndices);

        for (var finalIdx = 0; finalIdx < renamePlan.indices.length; finalIdx++) {
            setItemName(
                items[renamePlan.indices[finalIdx]],
                renamePlan.names[finalIdx],
                "rename item index " + renamePlan.indices[finalIdx] + " to '" + renamePlan.names[finalIdx] + "'"
            );
        }

        invalidateFrontmostTextCache();
        return true;
    }

    /**
     * canvas を変更せずに、［更新］したらこうなるという名前を計算する
     * 連番トークン {#N} は現在の canvas 順で採番されるため、一覧だけ並び替えた直後の
     * プレビュー値は確定後の値と一時的にずれる（［更新］または OK で一致する）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} settings - 現在の設定
     * @returns {Array<string>} 元の並び順でのプレビュー名
     */
    function computePreviewNames(doc, settings) {
        var itemType = settings.itemType || "artboard";
        var items = getDocumentItems(doc, itemType);

        /* 既存の名前で初期化する（未選択アイテムは現在の名前を保つ） */
        var previewNames = [];
        for (var initIdx = 0; initIdx < items.length; initIdx++) {
            previewNames.push(items[initIdx].name);
        }
        if (hasNoRenameInput(settings)) return previewNames;

        var itemTextMap = buildItemTextMap(doc, items, settings, itemType);
        var selectedIndices = getRangeItemIndices(items.length, settings.rangeMode, settings.rangeText);
        var renamePlan = buildRenamePlan(items, itemTextMap, settings, selectedIndices);
        for (var finalIdx = 0; finalIdx < renamePlan.indices.length; finalIdx++) {
            previewNames[renamePlan.indices[finalIdx]] = renamePlan.names[finalIdx];
        }
        return previewNames;
    }

    /**
     * 選択アイテムの最終名プランを作る（プレビューと本番で共用）
     * @param {object} items - 対象アイテムのコレクション
     * @param {object} itemTextMap - インデックスごとのベース文字列
     * @param {object} settings - 現在の設定
     * @param {Array<number>} selectedIndices - 対象インデックス
     * @returns {{indices: Array<number>, names: Array<string>}} リネーム対象と最終名
     */
    function buildRenamePlan(items, itemTextMap, settings, selectedIndices) {
        var prefixTemplate = settings.prefix || "";
        var suffixTemplate = settings.suffix || "";
        var findText = settings.findText || "";
        var replaceText = settings.replaceText || "";
        var useRegex = !!settings.useRegex;

        var skipUniquification = hasSequenceToken(prefixTemplate)
            || hasSequenceToken(suffixTemplate)
            || (findText && hasSequenceToken(replaceText));
        var reservedNames = getReservedItemNames(items, selectedIndices);
        var selectedIndexSet = makeIndexSet(selectedIndices);
        var tokenContext = createTokenContext();
        var plannedBaseNames = [];
        var plannedIndices = [];
        var sequenceIndex = 1;

        for (var itemIdx = 0; itemIdx < items.length; itemIdx++) {
            if (!selectedIndexSet[itemIdx]) continue;
            var expandedPrefix = expandTemplateTokens(prefixTemplate, sequenceIndex, tokenContext);
            var expandedSuffix = expandTemplateTokens(suffixTemplate, sequenceIndex, tokenContext);
            var textPart = itemTextMap[itemIdx] ? itemTextMap[itemIdx].join(" ") : "";
            var baseName = expandedPrefix + textPart + expandedSuffix;

            /* 接頭辞・基準・接尾辞のどれも未指定なら、検索・置換だけで現在名を加工する */
            if (!baseName && findText) {
                baseName = items[itemIdx].name;
            }

            if (findText) {
                var expandedReplace = expandTemplateTokens(replaceText, sequenceIndex, tokenContext);
                baseName = applyFindReplace(baseName, findText, expandedReplace, useRegex);
            }

            if (baseName) {
                plannedBaseNames.push(baseName);
                plannedIndices.push(itemIdx);
                sequenceIndex++;
            } else {
                /* 選択済みだが結果が空になる行はスキップし、現在の名前を予約して衝突を防ぐ */
                reservedNames.push(items[itemIdx].name);
            }
        }

        return {
            indices: plannedIndices,
            names: resolveUniqueNames(plannedBaseNames, reservedNames, skipUniquification)
        };
    }

    /**
     * 検索・置換を適用する（正規表現対応）
     * @param {string} name - 元の名前
     * @param {string} findPattern - 検索文字列
     * @param {string} replaceText - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @returns {string} 置換後の名前（不正な正規表現なら元の名前）
     */
    function applyFindReplace(name, findPattern, replaceText, useRegex) {
        if (!findPattern) return name;
        try {
            var regex;
            var effectiveReplace = replaceText;
            if (useRegex) {
                regex = new RegExp(findPattern, "g");
            } else {
                var escapedPattern = findPattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                regex = new RegExp(escapedPattern, "g");
                /* リテラル置換では replaceText 中の $ も無効化する（$1 や $& を特殊解釈させない） */
                effectiveReplace = replaceText.replace(/\$/g, "$$$$");
            }
            return name.replace(regex, effectiveReplace);
        } catch (regexError) {
            logFailure("invalid find pattern '" + findPattern + "'", regexError);
            return name;
        }
    }

    // =========================================
    // 名前のユーティリティ / Name utilities
    // =========================================

    /**
     * 種類に応じたドキュメントのコレクションを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} itemType - "artboard" / "symbol" / "layer" / "graphicstyle"
     * @returns {object} コレクションまたは配列
     */
    function getDocumentItems(doc, itemType) {
        if (itemType === "symbol") return doc.symbols;
        if (itemType === "layer") return doc.layers;
        if (itemType === "graphicstyle") return getRenamableGraphicStyles(doc);
        return doc.artboards;
    }

    /**
     * リネームできるグラフィックスタイルだけを配列で返す（`[Default]` などの予約スタイルを除く）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<object>} リネーム可能なグラフィックスタイル
     */
    function getRenamableGraphicStyles(doc) {
        var renamable = [];
        for (var gsi = 0; gsi < doc.graphicStyles.length; gsi++) {
            var graphicStyle = doc.graphicStyles[gsi];
            if (RESERVED_STYLE_NAME.test(graphicStyle.name)) continue;
            renamable.push(graphicStyle);
        }
        return renamable;
    }

    /**
     * 未選択アイテムの名前を予約名として返す（衝突回避用）
     * @param {object} items - 対象アイテムのコレクション
     * @param {Array<number>} selectedIndices - 選択インデックス
     * @returns {Array<string>} 予約名
     */
    function getReservedItemNames(items, selectedIndices) {
        var selectedIndexSet = makeIndexSet(selectedIndices);
        var reserved = [];
        for (var i = 0; i < items.length; i++) {
            if (!selectedIndexSet[i]) reserved.push(items[i].name);
        }
        return reserved;
    }

    /**
     * 名前をハッシュのキーとして安全な形にする
     * 接頭辞を付けないと `toString` や `valueOf` が Object.prototype のメンバーと衝突し、
     * 重複判定やカウンターが壊れる
     * @param {string} name - アイテム名
     * @returns {string} ハッシュ用のキー
     */
    function nameKey(name) {
        return "name:" + name;
    }

    /**
     * 重複しているものだけに "_1", "_2" を付けて最終名を返す
     * 重複しない名前はそのまま。ハッシュ集合を使い O(N) で衝突判定する
     * @param {Array<string>} plannedBaseNames - 計画中のベース名
     * @param {Array<string>} reservedNames - 予約済みの名前
     * @param {boolean} skipUniquification - 連番トークンがあり一意が前提なら true
     * @returns {Array<string>} 最終名
     */
    function resolveUniqueNames(plannedBaseNames, reservedNames, skipUniquification) {
        var resolvedNames = [];
        if (skipUniquification) {
            for (var i = 0; i < plannedBaseNames.length; i++) {
                resolvedNames.push(plannedBaseNames[i]);
            }
            return resolvedNames;
        }

        /* plan 内での baseName 出現回数（2以上なら重複） */
        var baseNameCountInPlan = {};
        for (var countIdx = 0; countIdx < plannedBaseNames.length; countIdx++) {
            var countKey = nameKey(plannedBaseNames[countIdx]);
            baseNameCountInPlan[countKey] = (baseNameCountInPlan[countKey] || 0) + 1;
        }

        var reservedSet = {};
        var usedNameSet = {};
        for (var reservedIdx = 0; reservedIdx < reservedNames.length; reservedIdx++) {
            reservedSet[nameKey(reservedNames[reservedIdx])] = true;
            usedNameSet[nameKey(reservedNames[reservedIdx])] = true;
        }

        var collisionSeqByBase = {};
        for (var planIdx = 0; planIdx < plannedBaseNames.length; planIdx++) {
            var baseName = plannedBaseNames[planIdx];
            var baseKey = nameKey(baseName);
            var needsSuffix = baseNameCountInPlan[baseKey] > 1 || reservedSet[baseKey] === true;
            var finalName = baseName;

            /* 重複しない名前はそのまま使う。ただし先に確定した名前と当たったときは、
               空いている連番が見つかるまで "_1", "_2" … を試す
               （この当たり判定がないと、["A","A","A_1"] の3件目が1件目と同じ "A_1" になる） */
            if (needsSuffix || usedNameSet[nameKey(finalName)] === true) {
                do {
                    collisionSeqByBase[baseKey] = (collisionSeqByBase[baseKey] || 0) + 1;
                    finalName = baseName + COLLISION_SEPARATOR + collisionSeqByBase[baseKey];
                } while (usedNameSet[nameKey(finalName)] === true);
            }

            usedNameSet[nameKey(finalName)] = true;
            resolvedNames.push(finalName);
        }
        return resolvedNames;
    }

    /**
     * expandTemplateTokens 用の実行コンテキストを作る（fileName / dateString は呼び出しごとに同一）
     * @returns {{fileName: string, dateString: string}} トークン展開用の値
     */
    function createTokenContext() {
        var now = new Date();
        return {
            fileName: app.activeDocument.name.replace(/\.[^.]+$/, ""),
            dateString: now.getFullYear().toString() +
                ("0" + (now.getMonth() + 1)).slice(-2) +
                ("0" + now.getDate()).slice(-2)
        };
    }

    /**
     * テンプレート文字列の連番・#FN・#DT トークンを展開する
     * @param {string} template - テンプレート文字列
     * @param {number} index - 1 始まりの連番
     * @param {object} [context] - createTokenContext() の戻り値
     * @returns {string} 展開後の文字列
     */
    function expandTemplateTokens(template, index, context) {
        var tokenContext = context || createTokenContext();

        /* 連番トークン {#N} を展開する（ゼロパディング対応：{#01} → 01, 02, ...） */
        var result = template.replace(/\{#(\d+)\}/g, function (match, token) {
            var value = parseInt(token, 10) + index - 1;
            if (token.charAt(0) === "0" && token.length > 1) {
                return ("0000000000" + value).slice(-token.length);
            }
            return value.toString();
        });

        /* 差し込む値に `$&` や `$$` が含まれても置換パターンとして解釈されないよう、
           文字列ではなく関数を渡す（ファイル名が "A$$B" や "Price$&List" のケース） */
        result = result.replace(/#FN/g, function () { return tokenContext.fileName; });
        result = result.replace(/#DT/g, function () { return tokenContext.dateString; });
        return result;
    }

    /**
     * テンプレートに連番トークン {#N} が含まれるか判定する
     * @param {string} template - テンプレート文字列
     * @returns {boolean} 含まれていれば true
     */
    function hasSequenceToken(template) {
        return /\{#\d+\}/.test(template);
    }

    // =========================================
    // 範囲指定のユーティリティ / Range utilities
    // =========================================

    /**
     * "1-3,5" 形式の範囲文字列を 0 始まりのインデックス配列にする
     * @param {string} rangeText - 範囲文字列
     * @returns {Array<number>} 0 始まりのインデックス
     */
    function parseItemRangeString(rangeText) {
        var result = [];
        var parts = rangeText.split(",");
        for (var i = 0; i < parts.length; i++) {
            var part = parts[i].replace(/\s+/g, "");
            if (/^\d+$/.test(part)) {
                result.push(parseInt(part, 10) - 1);
            } else if (/^\d+-\d+$/.test(part)) {
                var range = part.split("-");
                var rangeStart = parseInt(range[0], 10);
                var rangeEnd = parseInt(range[1], 10);
                for (var j = rangeStart; j <= rangeEnd; j++) result.push(j - 1);
            }
        }
        return result;
    }

    /**
     * "all" / "numbered" に応じて対象インデックスを返す
     * @param {number} itemCount - アイテム総数
     * @param {string} rangeMode - "all" または "numbered"
     * @param {string} rangeText - 範囲文字列
     * @returns {Array<number>} 対象インデックス
     */
    function getRangeItemIndices(itemCount, rangeMode, rangeText) {
        if (rangeMode === "all") {
            var allIndices = [];
            for (var i = 0; i < itemCount; i++) allIndices.push(i);
            return allIndices;
        }
        return parseItemRangeString(rangeText);
    }

    /**
     * 0 始まりのインデックス配列から "1-3,5" 形式の文字列を作る
     * @param {Array<number>} zeroBasedIndices - 0 始まりのインデックス
     * @returns {string} 1 始まりの範囲文字列
     */
    function buildRangeString(zeroBasedIndices) {
        if (!zeroBasedIndices || zeroBasedIndices.length === 0) return "";
        var sorted = [];
        for (var i = 0; i < zeroBasedIndices.length; i++) sorted.push(zeroBasedIndices[i]);
        sorted.sort(function (a, b) { return a - b; });

        var parts = [];
        var rangeStart = sorted[0];
        var rangeEnd = sorted[0];

        /**
         * 連続した並びを "start" または "start-end" として書き出す
         * @returns {void}
         */
        function pushRange() {
            parts.push(rangeStart === rangeEnd
                ? (rangeStart + 1) + ""
                : (rangeStart + 1) + "-" + (rangeEnd + 1));
        }

        for (var j = 1; j < sorted.length; j++) {
            if (sorted[j] === rangeEnd + 1) {
                rangeEnd = sorted[j];
            } else {
                pushRange();
                rangeStart = sorted[j];
                rangeEnd = sorted[j];
            }
        }
        pushRange();
        return parts.join(",");
    }

    /**
     * チェック済みインデックスから範囲設定（すべて／指定範囲）を導出する
     * 呼び出し側でインデックスの基準（表示位置か originalIndex か）を決めて渡すこと
     * @param {Array<number>} checkedIndices - チェック済みのインデックス
     * @param {number} totalCount - 全体の件数
     * @returns {{rangeMode: string, rangeText: string}} 範囲設定
     */
    function deriveRangeSettings(checkedIndices, totalCount) {
        if (checkedIndices.length === totalCount) {
            return { rangeMode: "all", rangeText: "" };
        }
        return { rangeMode: "numbered", rangeText: buildRangeString(checkedIndices) };
    }

    /**
     * インデックス配列をハッシュ集合へ変換する
     * @param {Array<number>} indices - インデックス配列
     * @returns {object} メンバー判定用のハッシュ集合
     */
    function makeIndexSet(indices) {
        var set = {};
        for (var i = 0; i < indices.length; i++) set[indices[i]] = true;
        return set;
    }

    // =========================================
    // テキストフレームの探索 / Text frame lookup
    // =========================================

    /**
     * テキストフレームの可視バウンズの中心座標を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Array<number>} [x, y] の中心座標
     */
    function getTextCenter(textFrame) {
        var bounds = textFrame.visibleBounds;
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /**
     * 中心座標がアートボード矩形の内側かを判定する
     * @param {Array<number>} center - [x, y] の中心座標
     * @param {Array<number>} artboardBounds - artboardRect [左, 上, 右, 下]
     * @returns {boolean} 内側なら true
     */
    function isCenterInsideBounds(center, artboardBounds) {
        return center[0] >= artboardBounds[0] && center[0] <= artboardBounds[2] &&
            center[1] <= artboardBounds[1] && center[1] >= artboardBounds[3];
    }

    /**
     * テキストフレームを所属アートボードに紐付けてマッピングする
     * @param {Array<TextFrame>} textFrames - 対象のテキストフレーム
     * @param {object} artboards - アートボードのコレクション
     * @returns {object} アートボードのインデックスをキーにした文字列配列
     */
    function mapTextFramesToArtboards(textFrames, artboards) {
        var map = {};
        for (var i = 0; i < textFrames.length; i++) {
            var center = getTextCenter(textFrames[i]);
            for (var j = 0; j < artboards.length; j++) {
                if (!isCenterInsideBounds(center, artboards[j].artboardRect)) continue;
                if (!map[j]) map[j] = [];
                map[j].push(textFrames[i].contents.replace(/[\r\n\t]/g, ""));
                break;
            }
        }
        return map;
    }

    /* 最前面テキストの走査結果キャッシュ / Cached frontmost-text scan
       走査結果は設定ではなく書類にしか依存しないので、canvas を書き換えるまで使い回せる */
    var frontmostTextFrameCache = null;

    /**
     * 最前面 TextFrame の走査結果を返す（キャッシュがあれば再利用する）
     * 走査は O(アートボード数 × ページアイテム数) なので、入力のたびに回すと大きな書類で止まる
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<TextFrame>} 見つかったテキストフレーム
     */
    function getCachedFrontmostTextFrames(doc) {
        if (!frontmostTextFrameCache) {
            frontmostTextFrameCache = getFrontmostTextFramesPerArtboard(doc);
        }
        return frontmostTextFrameCache;
    }

    /**
     * 最前面テキストのキャッシュを捨てる（canvas を書き換えたあとに呼ぶ）
     * @returns {void}
     */
    function invalidateFrontmostTextCache() {
        frontmostTextFrameCache = null;
    }

    /**
     * 各アートボードの最前面 TextFrame を、レイヤー・グループ階層を再帰して取得する
     * 判定順はレイヤー順・pageItems 順に依存する（Illustrator の厳密な描画Z順ではない）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<TextFrame>} 見つかったテキストフレーム
     */
    function getFrontmostTextFramesPerArtboard(doc) {

        /**
         * コンテナ内を再帰的に走査して最初に見つかった TextFrame を返す
         * @param {Layer|GroupItem} container - 走査対象
         * @param {Array<number>} artboardBounds - artboardRect
         * @returns {TextFrame|null} 見つかったテキストフレーム
         */
        function findFrontmostTextFrameInContainer(container, artboardBounds) {
            var items = container.pageItems;
            for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
                var item = items[itemIndex];
                if (item.hidden || item.locked) continue;

                if (item.typename === "TextFrame") {
                    if (isCenterInsideBounds(getTextCenter(item), artboardBounds)) return item;
                } else if (item.typename === "GroupItem") {
                    var nestedTextFrame = findFrontmostTextFrameInContainer(item, artboardBounds);
                    if (nestedTextFrame) return nestedTextFrame;
                }
            }
            return null;
        }

        /**
         * レイヤーとサブレイヤーを再帰的に走査する
         * @param {Layer} layer - 走査対象のレイヤー
         * @param {Array<number>} artboardBounds - artboardRect
         * @returns {TextFrame|null} 見つかったテキストフレーム
         */
        function findFrontmostTextFrameInLayer(layer, artboardBounds) {
            if (!layer.visible || layer.locked) return null;

            var textFrame = findFrontmostTextFrameInContainer(layer, artboardBounds);
            if (textFrame) return textFrame;

            for (var layerIndex = 0; layerIndex < layer.layers.length; layerIndex++) {
                var nestedTextFrame = findFrontmostTextFrameInLayer(layer.layers[layerIndex], artboardBounds);
                if (nestedTextFrame) return nestedTextFrame;
            }
            return null;
        }

        var result = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboardBounds = doc.artboards[artboardIndex].artboardRect;
            var frontmostFrame = null;
            for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
                frontmostFrame = findFrontmostTextFrameInLayer(doc.layers[layerIndex], artboardBounds);
                if (frontmostFrame) break;
            }
            if (frontmostFrame) result.push(frontmostFrame);
        }
        return result;
    }

    // =========================================
    // メイン処理 / Main entry
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var originalState = captureOriginalState(doc);
        var dialogResult = showRenameDialog(doc);

        if (!dialogResult) {
            /* キャンセル：ダイアログを開く前の状態（名前・rect・並び順）まで戻す */
            restoreOriginalState(doc, originalState);
            return;
        }

        /* ［更新］以降に差分がなければ、二重適用を避けて何もしない */
        if (dialogResult.skipApplyOnOk) return;

        if (hasReorderOrRename(dialogResult.itemEntries)) {
            applyReorderAndRename(doc, dialogResult.itemEntries, dialogResult);
        } else {
            executeRename(doc, dialogResult);
        }
    }

    main();

})();
