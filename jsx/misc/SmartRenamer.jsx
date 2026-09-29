#target illustrator
#targetengine "SmartRenamerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード・シンボル・レイヤー・グラフィックスタイルの名前を、接頭辞／接尾辞／名前の基準／検索置換を組み合わせて一括リネームします。
ダイアログ上で対象の絞り込み・並び替え・個別の手動編集ができ、結果はプレビューで確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenamer.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n2db43c753c0b

### Overview

Renames artboards, symbols, layers and graphic styles in bulk, combining a prefix, a suffix, a naming basis and find-and-replace.
The dialog filters, reorders and hand-edits individual entries, with a preview of the result.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartRenamer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartRenamer";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartRenamer.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartRenamer.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2db43c753c0b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
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
    var LIST_VISIBLE_ROWS   = 12;                 /* 一覧に一度に表示する行数 / rows shown at once in the item list */
    var SCROLLBAR_WIDTH     = 16;                 /* 一覧のスクロールバーの幅 / width of the item list scrollbar */
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
                ja: "接頭辞・接尾辞・「指定」の文字列・検索文字列のいずれかを入力してください。",
                en: "Enter a prefix, suffix, custom text, or find text to rename."
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
            },
            custom: { ja: "右の欄に入力した文字列を名前にします", en: "Uses the text entered in the field on the right as the name" },
            frontmost: {
                ja: "各アートボードで最前面にあるテキストの内容を名前にします（アートボードのみ）",
                en: "Uses the contents of the frontmost text on each artboard (artboards only)"
            },
            rangeItems: {
                ja: "対象にする行を「順」列の番号で指定します（例：1-3,5）",
                en: "Rows to include, by their numbers in the # column (e.g. 1-3,5)"
            },
            tokenSequence: { ja: "連番（1, 2, 3…）を入れます", en: "Inserts a sequence number (1, 2, 3...)" },
            tokenSequencePadded: {
                ja: "ゼロ埋めの連番（01, 02, 03…）を入れます",
                en: "Inserts a zero-padded sequence number (01, 02, 03...)"
            },
            tokenFileName: { ja: "ドキュメント名（拡張子なし）を入れます", en: "Inserts the document name without its extension" },
            tokenDate: { ja: "今日の日付（YYYYMMDD）を入れます", en: "Inserts today's date (YYYYMMDD)" },
            tokenDigit: { ja: "正規表現の「数字1文字」（\\d）を入れます", en: "Inserts the regex for one digit (\\d)" },
            tokenDigits: { ja: "正規表現の「続いた数字」（\\d+）を入れます", en: "Inserts the regex for a run of digits (\\d+)" },
            tokenAnyText: { ja: "正規表現の「任意の文字列」（.+）を入れます", en: "Inserts the regex for any text (.+)" },
            clearField: { ja: "入力欄を空にします", en: "Clears the field" },
            moveTop: { ja: "チェックした行を先頭へ移動します", en: "Moves the checked rows to the top" },
            moveUp: { ja: "チェックした行を1つ上へ移動します", en: "Moves the checked rows up one place" },
            moveDown: { ja: "チェックした行を1つ下へ移動します", en: "Moves the checked rows down one place" },
            moveBottom: { ja: "チェックした行を末尾へ移動します", en: "Moves the checked rows to the bottom" },
            rowCheckbox: {
                ja: "Option＋クリックで全行を同じ状態にします（全行がオンのときはこの行だけを残します）",
                en: "Option-click sets every row the same way (when all rows are on, only this row stays on)"
            },
            refresh: {
                ja: "現在の設定をドキュメントに反映します。ダイアログは閉じず、キャンセルすると元に戻ります",
                en: "Applies the current settings to the document without closing. Cancel restores the original state"
            }
        }
    };

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
     * @param {string} panelTitle - パネルのタイトル
     * @param {number} [spacing] - パネル内の要素間隔
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle, spacing) {
        var createdPanel = parentContainer.add("panel", undefined, panelTitle);
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
     * @param {string} displayText - 表示文字列
     * @param {number} width - 列幅
     * @returns {StaticText} 生成したテキスト
     */
    function addFixedWidthText(parentContainer, displayText, width) {
        var staticText = parentContainer.add("statictext", undefined, displayText);
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
        { label: "1", value: "{#1}", width: NARROW_BUTTON_WIDTH, tooltip: "tooltip.tokenSequence" },
        { label: "01", value: "{#01}", tooltip: "tooltip.tokenSequencePadded" },
        { label: "-", value: "-", width: NARROW_BUTTON_WIDTH },
        { label: "_", value: "_", width: NARROW_BUTTON_WIDTH },
        { label: "#FN", value: "#FN", width: 40, tooltip: "tooltip.tokenFileName" },
        { label: "#DT", value: "#DT", width: 40, tooltip: "tooltip.tokenDate" }
    ];

    /* 検索欄に挿入する正規表現ショートカット / Regex shortcuts inserted into the find field */
    var FIND_PATTERN_TOKENS = [
        { label: "#", value: "\\d", width: NARROW_BUTTON_WIDTH, tooltip: "tooltip.tokenDigit" },
        { label: "##", value: "\\d+", tooltip: "tooltip.tokenDigits" },
        { label: "*", value: ".+", width: NARROW_BUTTON_WIDTH, tooltip: "tooltip.tokenAnyText" }
    ];

    /* 置換欄に挿入するトークン / Tokens inserted into the replace field */
    var REPLACE_TOKENS = [
        { label: "#", value: "{#1}", width: NARROW_BUTTON_WIDTH, tooltip: "tooltip.tokenSequence" },
        { label: "##", value: "{#01}", tooltip: "tooltip.tokenSequencePadded" },
        { label: "-", value: "-", width: NARROW_BUTTON_WIDTH },
        { label: "_", value: "_", width: NARROW_BUTTON_WIDTH }
    ];

    /**
     * トークン挿入ボタンの行を作る
     * @param {Panel|Group} parentContainer - 追加先
     * @param {EditText} targetInput - 挿入先の入力欄
     * @param {Array<object>} tokens - {label, value, width, tooltip} の配列（tooltip はラベルキー）
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
                if (token.tooltip) tokenButton.helpTip = getLabel(token.tooltip);
                tokenButton.onClick = function () {
                    targetInput.text = targetInput.text + token.value;
                    targetInput.notify("onChange");
                };
            })(tokens[tokenIdx]);
        }
        if (withClearButton) {
            var clearButton = tokenRow.add("button", undefined, "x");
            clearButton.preferredSize = [NARROW_BUTTON_WIDTH, TOKEN_BUTTON_SIZE[1]];
            clearButton.helpTip = getLabel("tooltip.clearField");
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
     * @param {string} logContext - どこで失敗したかを示す文字列
     * @param {object} error - 捕捉した例外
     * @returns {void}
     */
    function logFailure(logContext, error) {
        $.writeln("[" + SCRIPT_NAME + "] " + logContext + ": " + error);
    }

    /**
     * アイテムに名前を設定する（失敗しても処理を続ける）
     * @param {object} targetItem - Artboard / SymbolItem / Layer / GraphicStyle
     * @param {string} newName - 設定する名前
     * @param {string} logContext - ログ用の文字列
     * @returns {void}
     */
    function setItemName(targetItem, newName, logContext) {
        try {
            targetItem.name = newName;
        } catch (nameError) {
            logFailure(logContext, nameError);
        }
    }

    /**
     * アートボードの位置を設定する（失敗しても処理を続け、名前の復元を止めない）
     * @param {Artboard} targetArtboard - 対象のアートボード
     * @param {Array<number>} artboardBounds - artboardRect [左, 上, 右, 下]
     * @param {string} logContext - ログ用の文字列
     * @returns {void}
     */
    function setArtboardRect(targetArtboard, artboardBounds, logContext) {
        try {
            targetArtboard.artboardRect = artboardBounds;
        } catch (rectError) {
            logFailure(logContext, rectError);
        }
    }

    /**
     * アイテムをコレクションの先頭へ移動する（失敗しても処理を続ける）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} targetItem - 移動するアイテム
     * @param {string} logContext - ログ用の文字列
     * @returns {void}
     */
    function moveToBeginning(doc, targetItem, logContext) {
        try {
            targetItem.move(doc, ElementPlacement.PLACEATBEGINNING);
        } catch (moveError) {
            logFailure(logContext, moveError);
        }
    }

    /**
     * コレクションの名前と参照を控える
     * @param {object} sourceItems - コレクションまたは配列
     * @returns {{names: Array<string>, refs: Array<object>}} 名前と参照の組
     */
    function captureNamesAndRefs(sourceItems) {
        var capturedState = { names: [], refs: [] };
        for (var i = 0; i < sourceItems.length; i++) {
            capturedState.names.push(sourceItems[i].name);
            capturedState.refs.push(sourceItems[i]);
        }
        return capturedState;
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
     * @param {number} artboardCount - 対象件数
     * @returns {void}
     */
    function assignTemporaryArtboardNames(artboards, artboardCount) {
        for (var i = 0; i < artboardCount; i++) {
            setItemName(artboards[i], TEMP_ARTBOARD_PREFIX + i + "__", "temporary artboard name at " + i);
        }
    }

    /**
     * 種類ごとに並び替えできるか（Symbol / GraphicStyle には move() が無い。実測済み）
     * @param {string} itemType - "artboard" / "symbol" / "layer" / "graphicstyle"
     * @returns {boolean} 並び替えできれば true
     */
    function canReorderItemType(itemType) {
        return itemType === "artboard" || itemType === "layer";
    }

    /**
     * 控えておいた名前を復元する
     * @param {{names: Array<string>, refs: Array<object>}} capturedState - 控えた状態
     * @param {string} logContext - ログ用の種類名
     * @returns {void}
     */
    function restoreNames(capturedState, logContext) {
        for (var nameIdx = 0; nameIdx < capturedState.refs.length; nameIdx++) {
            setItemName(capturedState.refs[nameIdx], capturedState.names[nameIdx], logContext + " name restore at " + nameIdx);
        }
    }

    /**
     * 控えておいた参照の並び順と名前を復元する
     * @param {Document} doc - 対象ドキュメント
     * @param {{names: Array<string>, refs: Array<object>}} capturedState - 控えた状態
     * @param {string} logContext - ログ用の種類名
     * @returns {void}
     */
    function restoreOrderAndNames(doc, capturedState, logContext) {
        var itemRefs = capturedState.refs;
        /* 末尾の要素から順に先頭へ送ると、控えた並び順どおりに戻る */
        for (var reverseIdx = itemRefs.length - 1; reverseIdx >= 0; reverseIdx--) {
            moveToBeginning(doc, itemRefs[reverseIdx], logContext + " order restore at " + reverseIdx);
        }
        restoreNames(capturedState, logContext);
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
            setArtboardRect(artboards[abIdx], originalState.artboard.rects[abIdx], "artboard rect restore at " + abIdx);
            setItemName(artboards[abIdx], originalState.artboard.names[abIdx], "artboard name restore at " + abIdx);
        }
        /* シンボル・グラフィックスタイルは並び替えできないので名前だけ戻す */
        restoreNames(originalState.symbol, "symbol");
        restoreOrderAndNames(doc, originalState.layer, "layer");
        restoreNames(originalState.graphicStyle, "graphic style");
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
        var signatureParts = [];
        var settingsKeys = ["itemType", "mode", "prefix", "suffix", "customText", "rangeMode", "rangeText", "findText", "replaceText"];
        for (var keyIdx = 0; keyIdx < settingsKeys.length; keyIdx++) {
            signatureParts.push(settingsKeys[keyIdx] + "=" + (settings[settingsKeys[keyIdx]] || ""));
        }
        signatureParts.push("useRegex=" + (!!settings.useRegex));

        if (settings.itemEntries) {
            for (var entryIdx = 0; entryIdx < settings.itemEntries.length; entryIdx++) {
                var entry = settings.itemEntries[entryIdx];
                signatureParts.push([
                    "entry",
                    entryIdx,
                    entry.originalIndex,
                    entry.checked ? "1" : "0",
                    entry.userEdited ? "1" : "0",
                    entry.newName || ""
                ].join(":"));
            }
        }
        return signatureParts.join("\n");
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ構築 / Dialog build
    // =========================================

    /**
     * 種類のラジオボタン行（ダイアログ最上段）を作る
     * @param {Window} renameDialog - 追加先のダイアログ
     * @returns {{artboard: RadioButton, symbol: RadioButton, layer: RadioButton, graphicStyle: RadioButton}} 種類ごとのラジオ
     */
    function addItemTypeRow(renameDialog) {
        var itemTypeRow = renameDialog.add("group");
        itemTypeRow.orientation = "row";
        itemTypeRow.alignment = ["fill", "top"];
        itemTypeRow.alignChildren = ["center", "center"];   /* ラジオは行の中央にまとめて置く */
        itemTypeRow.margins = 0;

        var artboardRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeArtboard"));
        var symbolRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeSymbol"));
        var layerRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeLayer"));
        var graphicStyleRadio = itemTypeRow.add("radiobutton", undefined, getLabel("radio.itemTypeGraphicStyle"));
        return { artboard: artboardRadio, symbol: symbolRadio, layer: layerRadio, graphicStyle: graphicStyleRadio };
    }

    /**
     * 接頭辞・接尾辞のパネル（入力欄＋トークン挿入ボタン）を作る
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} titlePath - パネルタイトルのラベルキー
     * @param {number} chars - 入力欄の文字数
     * @returns {EditText} 作成した入力欄
     */
    function addAffixPanel(parentPanel, titlePath, chars) {
        var affixPanel = addPanel(parentPanel, getLabel(titlePath));
        var affixInput = affixPanel.add("edittext", undefined, "");
        affixInput.characters = chars;
        addTokenRow(affixPanel, affixInput, AFFIX_TOKENS, true);
        return affixInput;
    }

    /**
     * 「名前の基準」パネルを作る
     * @param {Panel} parentPanel - 追加先のパネル
     * @returns {{originalNameRadio: RadioButton, customRadio: RadioButton, customInput: EditText, frontmostRadio: RadioButton}} 作成したコントロール
     */
    function addNameSourcePanel(parentPanel) {
        var nameSourcePanel = addPanel(parentPanel, getLabel("panel.nameSource"));
        var originalNameRadio = nameSourcePanel.add("radiobutton", undefined, getLabel("radio.originalName"));
        originalNameRadio.alignment = "left";   /* パネル幅いっぱいに広げない */

        var customRow = nameSourcePanel.add("group");
        setupRow(customRow);
        var customRadio = customRow.add("radiobutton", undefined, getLabel("radio.custom"));
        customRadio.helpTip = getLabel("tooltip.custom");
        var customInput = customRow.add("edittext", undefined, "");
        customInput.characters = CUSTOM_CHARS;
        customInput.enabled = false;

        var frontmostRadio = nameSourcePanel.add("radiobutton", undefined, getLabel("radio.frontmost"));
        frontmostRadio.alignment = "left";
        frontmostRadio.helpTip = getLabel("tooltip.frontmost");

        /* ラジオを全部追加してから初期値を設定する（途中で設定すると後続の追加で解除される） */
        originalNameRadio.value = true;
        customRadio.value = false;
        frontmostRadio.value = false;

        return {
            originalNameRadio: originalNameRadio,
            customRadio: customRadio,
            customInput: customInput,
            frontmostRadio: frontmostRadio
        };
    }

    /**
     * 「検索・置換」パネルを作る
     * @param {Panel} parentPanel - 追加先のパネル
     * @returns {{findInput: EditText, replaceInput: EditText, regexCheckbox: Checkbox}} 作成したコントロール
     */
    function addFindReplacePanel(parentPanel) {
        var findReplacePanel = addPanel(parentPanel, getLabel("panel.findReplace"));

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

        return { findInput: findInput, replaceInput: replaceInput, regexCheckbox: regexCheckbox };
    }

    /**
     * 「フィルター」パネルを作る
     * @param {Group} parentColumn - 追加先のカラム
     * @returns {{filterAllRadio: RadioButton, filterRangeRadio: RadioButton, rangeInput: EditText, searchFilterCheckbox: Checkbox, searchInput: EditText}} 作成したコントロール
     */
    function addFilterPanel(parentColumn) {
        var filterPanel = addPanel(parentColumn, getLabel("panel.filter"), DENSE_SPACING);

        var rangeFilterRow = filterPanel.add("group");
        setupRow(rangeFilterRow);
        var filterAllRadio = rangeFilterRow.add("radiobutton", undefined, getLabel("radio.allItems"));
        var filterRangeRadio = rangeFilterRow.add("radiobutton", undefined, getLabel("radio.rangeItems"));
        var rangeInput = rangeFilterRow.add("edittext", undefined, "");
        rangeInput.characters = FILTER_CHARS;
        rangeInput.helpTip = getLabel("tooltip.rangeItems");
        rangeInput.enabled = false;
        filterAllRadio.value = true;
        filterRangeRadio.value = false;

        var searchFilterRow = filterPanel.add("group");
        setupRow(searchFilterRow);
        var searchFilterCheckbox = searchFilterRow.add("checkbox", undefined, getLabel("checkbox.searchFilter"));
        searchFilterCheckbox.value = false;
        searchFilterCheckbox.helpTip = getLabel("tooltip.searchFilter");
        var searchInput = searchFilterRow.add("edittext", undefined, "");
        searchInput.characters = FILTER_CHARS;
        searchInput.helpTip = getLabel("tooltip.searchFilter");
        searchInput.enabled = false;

        return {
            filterAllRadio: filterAllRadio,
            filterRangeRadio: filterRangeRadio,
            rangeInput: rangeInput,
            searchFilterCheckbox: searchFilterCheckbox,
            searchInput: searchInput
        };
    }

    /**
     * 一覧のパネル（行を並べる領域＋並び替えボタン）を作る
     * @param {Group} parentColumn - 追加先のカラム
     * @returns {{entryRowsHost: Group, listScrollbar: Scrollbar, moveToTopButton: Button, moveUpButton: Button, moveDownButton: Button, moveToBottomButton: Button}} 作成したコントロール
     */
    function addListPanel(parentColumn) {
        var listPanel = addPanel(parentColumn, getLabel("panel.list"), DENSE_SPACING);

        /* ScriptUI のグループはスクロールしないので、LIST_VISIBLE_ROWS 行だけを描き、スクロールバーで表示位置をずらす */
        var listBodyRow = listPanel.add("group");
        listBodyRow.orientation = "row";
        listBodyRow.alignment = ["fill", "top"];
        listBodyRow.alignChildren = ["left", "fill"];
        listBodyRow.spacing = DENSE_SPACING;

        var entryRowsHost = listBodyRow.add("group");
        entryRowsHost.orientation = "column";
        entryRowsHost.alignChildren = ["fill", "top"];
        entryRowsHost.spacing = TOKEN_SPACING;

        var listScrollbar = listBodyRow.add("scrollbar", undefined, 0, 0, 0);
        listScrollbar.preferredSize.width = SCROLLBAR_WIDTH;
        listScrollbar.alignment = ["right", "fill"];
        listScrollbar.stepdelta = 1;
        listScrollbar.jumpdelta = LIST_VISIBLE_ROWS;

        var moveButtonRow = listPanel.add("group");
        setupRow(moveButtonRow, "center", TOKEN_SPACING);
        moveButtonRow.margins = [0, 10, 0, 0];

        /**
         * 並び替えボタンを1つ追加する
         * @param {string} labelPath - ボタン名のラベルキー
         * @param {string} tooltipPath - tooltip のラベルキー
         * @param {number} extraWidth - 既定の幅に足す幅
         * @returns {Button} 追加したボタン
         */
        function addMoveButton(labelPath, tooltipPath, extraWidth) {
            var moveButton = moveButtonRow.add("button", undefined, getLabel(labelPath));
            moveButton.preferredSize = extraWidth ? [MOVE_BUTTON_SIZE[0] + extraWidth, MOVE_BUTTON_SIZE[1]] : MOVE_BUTTON_SIZE;
            moveButton.helpTip = getLabel(tooltipPath);
            return moveButton;
        }

        return {
            entryRowsHost: entryRowsHost,
            listScrollbar: listScrollbar,
            moveToTopButton: addMoveButton("button.moveTop", "tooltip.moveTop", 4),
            moveUpButton: addMoveButton("button.moveUp", "tooltip.moveUp", 0),
            moveDownButton: addMoveButton("button.moveDown", "tooltip.moveDown", 0),
            moveToBottomButton: addMoveButton("button.moveBottom", "tooltip.moveBottom", 4)
        };
    }

    /**
     * リネームダイアログのUIを構築する
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} ダイアログ本体・各コントロール・操作用メソッドをまとめたオブジェクト
     */
    function createRenameDialog(doc) {
        var renameDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(renameDialog);

        var itemTypeRadios = addItemTypeRow(renameDialog);

        // --- コンテンツ行（左：リネーム条件／右：フィルター＋一覧） ---
        var contentRow = renameDialog.add("group");
        contentRow.orientation = "row";
        contentRow.alignChildren = ["left", "top"];
        contentRow.spacing = COLUMN_SPACING;

        var leftColumn = addColumnGroup(contentRow);
        var rightColumn = addColumnGroup(contentRow);

        /* 左カラム：リネーム条件（接頭辞・名前の基準・検索置換・接尾辞）/ Left column: rename rules */
        var renameRulesPanel = addPanel(leftColumn, getLabel("panel.renameRules"));
        var prefixInput = addAffixPanel(renameRulesPanel, "panel.prefix", PREFIX_CHARS);
        focusField(prefixInput);
        var nameSourceControls = addNameSourcePanel(renameRulesPanel);
        var findReplaceControls = addFindReplacePanel(renameRulesPanel);
        var suffixInput = addAffixPanel(renameRulesPanel, "panel.suffix", SUFFIX_CHARS);

        /* 右カラム：フィルター（上段）と一覧（下段）/ Right column: filter and item list */
        var filterControls = addFilterPanel(rightColumn);
        var listControls = addListPanel(rightColumn);

        /* ボタンエリア（左：更新／右：キャンセル・OK）/ Button row: Refresh on the left, Cancel/OK on the right */
        var buttonRow = addButtonRow(renameDialog);
        var btnRefresh = buttonRow.leftGroup.add("button", undefined, getLabel("button.refresh"));
        btnRefresh.helpTip = getLabel("tooltip.refresh");
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /* 以降の処理で使うコントロール / Controls used by the closures below */
        var itemTypeArtboardRadio = itemTypeRadios.artboard;
        var itemTypeSymbolRadio = itemTypeRadios.symbol;
        var itemTypeLayerRadio = itemTypeRadios.layer;
        var itemTypeGraphicStyleRadio = itemTypeRadios.graphicStyle;
        var originalNameRadio = nameSourceControls.originalNameRadio;
        var customRadio = nameSourceControls.customRadio;
        var customInput = nameSourceControls.customInput;
        var frontmostRadio = nameSourceControls.frontmostRadio;
        var findInput = findReplaceControls.findInput;
        var replaceInput = findReplaceControls.replaceInput;
        var regexCheckbox = findReplaceControls.regexCheckbox;
        var filterAllRadio = filterControls.filterAllRadio;
        var filterRangeRadio = filterControls.filterRangeRadio;
        var rangeInput = filterControls.rangeInput;
        var searchFilterCheckbox = filterControls.searchFilterCheckbox;
        var searchInput = filterControls.searchInput;
        var entryRowsHost = listControls.entryRowsHost;
        var listScrollbar = listControls.listScrollbar;
        var moveToTopButton = listControls.moveToTopButton;
        var moveUpButton = listControls.moveUpButton;
        var moveDownButton = listControls.moveDownButton;
        var moveToBottomButton = listControls.moveToBottomButton;

        // =====================================
        // 一覧の状態管理 / Item list state
        // =====================================

        var currentItemType = DEFAULT_ITEM_TYPE;
        var entryRows = [];
        var listScrollOffset = 0;   /* 一覧の先頭に表示しているエントリの位置 / index of the first visible entry */

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
            var docItems = getDocumentItems(doc, itemType);
            var entries = [];
            for (var i = 0; i < docItems.length; i++) {
                var entry = {
                    originalIndex: i,
                    name: docItems[i].name,
                    newName: docItems[i].name,
                    checked: false,
                    userEdited: false
                };
                if (itemType === "artboard") {
                    entry.rect = docItems[i].artboardRect;
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
            var rangeSettings = deriveRangeSettings(checkedIndices, itemEntries.length);
            settings.rangeMode = rangeSettings.rangeMode;
            settings.rangeText = rangeSettings.rangeText;
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
         * フォーカスが残ったままの欄は onChange が発火しないので、欄と控えの差分でも手動編集とみなす
         * @returns {void}
         */
        function syncEditingValues() {
            for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                var entryRow = entryRows[rowIdx];
                var entry = itemEntries[entryRow.dataIndex];
                entry.checked = entryRow.checkbox.value;
                if (entry.checked && entryRow.newNameField.text !== entry.newName) {
                    entry.newName = entryRow.newNameField.text;
                    entry.userEdited = true;
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
                var entryRow = entryRows[rowIdx];
                var entry = itemEntries[entryRow.dataIndex];
                entryRow.checkbox.value = entry.checked;
                entryRow.newNameField.enabled = entry.checked;
                if (!entry.checked) entryRow.newNameField.text = entry.name;
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
            var rangeSettings = deriveRangeSettings(checkedPositions, itemEntries.length);
            var isAllChecked = (rangeSettings.rangeMode === "all");
            filterAllRadio.value = isAllChecked;
            filterRangeRadio.value = !isAllChecked;
            /* 検索フィルター ON の間は「指定範囲」欄を無効のまま保つ */
            rangeInput.enabled = !isAllChecked && !searchFilterCheckbox.value;
            if (!isAllChecked) rangeInput.text = rangeSettings.rangeText;
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
            var swappedEntry = itemEntries[indexA];
            itemEntries[indexA] = itemEntries[indexB];
            itemEntries[indexB] = swappedEntry;
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
            scrollToFirstCheckedEntry();
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
            scrollToFirstCheckedEntry();
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
            scrollToFirstCheckedEntry();
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
            scrollToFirstCheckedEntry();
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
            var isReorderable = canReorderItemType(currentItemType);
            var canMoveUpwards = isReorderable && canMoveUp();
            var canMoveDownwards = isReorderable && canMoveDown();
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
         * @param {number} entryIdx - エントリの位置
         * @returns {void}
         */
        function addEntryRow(entryIdx) {
            var entryRow = entryRowsHost.add("group");
            setupRow(entryRow, "left", DENSE_SPACING);

            addFixedWidthText(entryRow, (entryIdx + 1) + "", LIST_COLUMN_WIDTHS.order);

            var rowCheckbox = entryRow.add("checkbox", undefined, "");
            rowCheckbox.value = itemEntries[entryIdx].checked;
            rowCheckbox.preferredSize.width = LIST_COLUMN_WIDTHS.select;
            rowCheckbox.helpTip = getLabel("tooltip.rowCheckbox");

            var currentNameLabel = addFixedWidthText(entryRow, itemEntries[entryIdx].name, LIST_COLUMN_WIDTHS.currentName);
            currentNameLabel.helpTip = itemEntries[entryIdx].name;

            addFixedWidthText(entryRow, "→", LIST_COLUMN_WIDTHS.arrow);

            var newNameField = entryRow.add("edittext", undefined, itemEntries[entryIdx].newName);
            newNameField.preferredSize.width = LIST_COLUMN_WIDTHS.newName;
            newNameField.enabled = itemEntries[entryIdx].checked;

            rowCheckbox.onClick = function () {
                if (isOptionKeyHeld()) {
                    applyOptionClickToggle(entryIdx, rowCheckbox.value);
                } else {
                    setEntryChecked(entryIdx, rowCheckbox.value);
                    newNameField.enabled = rowCheckbox.value;
                    if (!rowCheckbox.value) newNameField.text = itemEntries[entryIdx].name;
                }
                syncFilterFromCheckboxes();
                if (requestPreviewUpdate) requestPreviewUpdate();
                updateMoveButtonsState();
            };

            newNameField.onChange = function () {
                itemEntries[entryIdx].newName = newNameField.text;
                itemEntries[entryIdx].userEdited = true;
            };

            entryRows.push({
                group: entryRow,
                checkbox: rowCheckbox,
                currentNameLabel: currentNameLabel,
                newNameField: newNameField,
                dataIndex: entryIdx
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

            /* 表示位置を件数に収めてから、見えている範囲の行だけを作る */
            var maxScrollOffset = Math.max(0, itemEntries.length - LIST_VISIBLE_ROWS);
            listScrollOffset = Math.max(0, Math.min(listScrollOffset, maxScrollOffset));
            listScrollbar.maxvalue = maxScrollOffset;
            listScrollbar.value = listScrollOffset;
            listScrollbar.enabled = (maxScrollOffset > 0);

            addListHeaderRow();
            var visibleEnd = Math.min(itemEntries.length, listScrollOffset + LIST_VISIBLE_ROWS);
            for (var entryIndex = listScrollOffset; entryIndex < visibleEnd; entryIndex++) {
                addEntryRow(entryIndex);
            }

            updateMoveButtonsState();
            entryRowsHost.layout.layout(true);
            renameDialog.layout.layout(true);
        }

        /**
         * 最初のチェック行が見える位置まで一覧をスクロールする（並び替えのあとに呼ぶ）
         * @returns {void}
         */
        function scrollToFirstCheckedEntry() {
            for (var i = 0; i < itemEntries.length; i++) {
                if (!itemEntries[i].checked) continue;
                if (i < listScrollOffset || i >= listScrollOffset + LIST_VISIBLE_ROWS) listScrollOffset = i;
                return;
            }
        }

        listScrollbar.onChanging = function () {
            var newScrollOffset = Math.round(listScrollbar.value);
            if (newScrollOffset === listScrollOffset) return;
            /* 行を作り直す前に、見えている行の入力を控えへ書き戻す */
            syncEditingValues();
            listScrollOffset = newScrollOffset;
            refreshReorderRows();
        };

        /**
         * ［更新］確定後にエントリを canvas の現状へ再ベースライン化する
         * userEdited を解除して、以降の接頭辞・接尾辞の変更をプレビューへ反映できるようにする
         * @returns {void}
         */
        function rebaselineEntriesAfterCommit() {
            var docItems = getDocumentItems(doc, currentItemType);

            /* 確定した結果アイテム数が変わることがある（例：グラフィックスタイル名が `[...]` になり
               RESERVED_STYLE_NAME に引っかかって対象から外れる）。ずれたまま参照を続けると
               範囲外アクセスになるので、件数が変わったら一覧ごと作り直す
               作り直した行は未チェックになるので、「すべて／指定範囲」もその状態にそろえる */
            if (docItems.length !== itemEntries.length) {
                itemEntries = buildEntriesForItemType(currentItemType);
                refreshReorderRows();
                syncFilterFromCheckboxes();
                return;
            }

            for (var i = 0; i < itemEntries.length; i++) {
                itemEntries[i].originalIndex = i;
                itemEntries[i].name = docItems[i].name;
                itemEntries[i].newName = docItems[i].name;
                itemEntries[i].userEdited = false;
                if (currentItemType === "artboard") {
                    itemEntries[i].rect = docItems[i].artboardRect;
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
            listScrollOffset = 0;
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

            renameDialog.layout.layout(true);
            refreshReorderRows();
            if (itemTypeChangeCallback) itemTypeChangeCallback();
        }

        /**
         * エントリに対応する canvas 上の現在名を返す
         * 確定でコレクションが縮むことがあるため、範囲外は現在のエントリ名で代替する
         * @param {object} docItems - 種類ごとのコレクション
         * @param {object} entry - 一覧のエントリ
         * @returns {string} 現在名
         */
        function getCurrentNameOfEntry(docItems, entry) {
            var currentItem = docItems[entry.originalIndex];
            return currentItem ? currentItem.name : entry.name;
        }

        /**
         * 「現在の名前」列の表示と tooltip（省略された全体名）を更新する
         * @param {object} entryRow - 一覧の行
         * @param {string} currentName - 表示する名前
         * @returns {void}
         */
        function setCurrentNameLabel(entryRow, currentName) {
            entryRow.currentNameLabel.text = currentName;
            entryRow.currentNameLabel.helpTip = currentName;
        }

        btnOK.onClick = function () {
            syncEditingValues();
            for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                if (itemEntries[entryIdx].checked && itemEntries[entryIdx].newName === "") {
                    alert(getLabel("alert.emptyName", { n: entryIdx + 1 }));
                    return;
                }
            }

            var okSettings = readSettings();
            okSettings.itemEntries = itemEntries;
            skipApplyOnOk = false;

            /* 入力がなく並び替え・手動編集もないときは、閉じる前に知らせる。
               ただし［更新］で確定済みなら、キャンセルでその確定を戻させないよう何もせずに閉じる */
            if (hasNoRenameInput(okSettings) && !hasReorderOrRename(itemEntries)) {
                if (lastCommittedSignature === null) {
                    alert(getLabel("alert.needSettings"));
                    return;
                }
                skipApplyOnOk = true;
            }

            /* ［更新］以降に何も変わっていなければ、OK では何もしない（二重適用の防止） */
            if (lastCommittedSignature !== null && buildSettingsSignature(okSettings) === lastCommittedSignature) {
                skipApplyOnOk = true;
            }
            renameDialog.close(1);
        };

        /* 全UI構築後に初期状態をそろえる。ラジオを立てるだけでなく setItemType() を通すことで、
           「最前面のテキスト」の有効・無効やフィルターの初期化も既定の種類に追随する */
        itemTypeArtboardRadio.value = (DEFAULT_ITEM_TYPE === "artboard");
        itemTypeSymbolRadio.value = (DEFAULT_ITEM_TYPE === "symbol");
        itemTypeLayerRadio.value = (DEFAULT_ITEM_TYPE === "layer");
        itemTypeGraphicStyleRadio.value = (DEFAULT_ITEM_TYPE === "graphicstyle");
        setItemType(DEFAULT_ITEM_TYPE);

        return {
            dialog: renameDialog,
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
            refreshButton: btnRefresh,
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
                /* 見えていない行のエントリもそろえる（スクロールで行を作り直すとエントリの値が出る） */
                var docItems = getDocumentItems(doc, currentItemType);
                for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                    var entry = itemEntries[entryIdx];
                    var currentName = getCurrentNameOfEntry(docItems, entry);
                    entry.name = currentName;
                    entry.newName = currentName;
                    entry.userEdited = false;
                }
                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var entryRow = entryRows[rowIdx];
                    var rowEntry = itemEntries[entryRow.dataIndex];
                    setCurrentNameLabel(entryRow, rowEntry.name);
                    entryRow.newNameField.text = rowEntry.newName;
                    entryRow.newNameField.enabled = rowEntry.checked;
                }
            },

            /**
             * 未確定のプレビュー名を一覧の「新しい名前」欄へ反映する
             * @param {Array<string>} previewNames - 元の並び順でのプレビュー名
             * @returns {void}
             */
            syncPreviewToReorderRows: function (previewNames) {
                /* 入力途中の行を手動編集として控えてから、見えていない行のエントリにもプレビュー名を入れる */
                syncEditingValues();
                var docItems = getDocumentItems(doc, currentItemType);
                for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                    var entry = itemEntries[entryIdx];
                    /* 「現在の名前」列は canvas の現状（［更新］後は確定後の名前）を出す */
                    var currentName = getCurrentNameOfEntry(docItems, entry);
                    entry.name = currentName;
                    /* 「新しい名前」列は未確定プレビュー。手動編集した行は上書きしない */
                    if (entry.userEdited) continue;
                    entry.newName = (previewNames && previewNames[entry.originalIndex] != null)
                        ? previewNames[entry.originalIndex]
                        : currentName;
                }
                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var entryRow = entryRows[rowIdx];
                    var rowEntry = itemEntries[entryRow.dataIndex];
                    setCurrentNameLabel(entryRow, rowEntry.name);
                    if (!rowEntry.userEdited) entryRow.newNameField.text = rowEntry.newName;
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
               （抑制しないと1クリックで5回、「最前面のテキスト」ではドキュメント全体の走査が5回起きる）
               フィルター欄（指定範囲・検索）は発火させない。発火させると一覧で手動で付け外ししたチェックが
               フィルター結果で上書きされ、見えているチェックと違う範囲が確定される */
            var pendingFields = [
                dialogUI.prefixInput,
                dialogUI.suffixInput,
                dialogUI.customInput,
                dialogUI.findInput,
                dialogUI.replaceInput
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

        prepareDialogWindow(dialogUI.dialog, SCRIPT_NAME);
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
     * @param {object} sourceItems - 並び替え前のコレクション
     * @param {Array<object>} itemEntries - 新しい並び順のエントリ配列
     * @param {string} logContext - ログ用の種類名
     * @returns {void}
     */
    function reorderByMove(doc, sourceItems, itemEntries, logContext) {
        var itemRefs = [];
        for (var initIdx = 0; initIdx < sourceItems.length; initIdx++) itemRefs.push(sourceItems[initIdx]);
        for (var reverseIdx = itemEntries.length - 1; reverseIdx >= 0; reverseIdx--) {
            moveToBeginning(doc, itemRefs[itemEntries[reverseIdx].originalIndex], logContext + " move at entry " + reverseIdx);
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
        var rangeSettings = deriveRangeSettings(checkedPositions, itemCount);
        reorderedSettings.rangeMode = rangeSettings.rangeMode;
        reorderedSettings.rangeText = rangeSettings.rangeText;
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
        var docItems = getDocumentItems(doc, itemType);
        var itemCount = docItems.length;

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
            currentNamesByOriginalIndex.push(docItems[origIdx].name);
        }

        if (itemType === "artboard") {
            /* rect と現在名を新しい位置へ並べ替える（一時名で衝突を避ける）
               一時名を付ける件数と並べ替える件数は必ず同じにする。rect の代入が失敗しても名前は戻す */
            var reorderCount = Math.min(itemEntries.length, itemCount);
            assignTemporaryArtboardNames(docItems, reorderCount);
            for (var newPos = 0; newPos < reorderCount; newPos++) {
                setArtboardRect(docItems[newPos], itemEntries[newPos].rect, "artboard rect reorder at " + newPos);
                setItemName(docItems[newPos], currentNamesByOriginalIndex[itemEntries[newPos].originalIndex], "artboard reorder at " + newPos);
            }
        } else if (canReorderItemType(itemType)) {
            /* レイヤーは安定参照を move() で並べ替える（シンボル・グラフィックスタイルは一覧でも並び替え不可） */
            reorderByMove(doc, docItems, itemEntries, itemType);
        }

        /* 並び替え後の位置を基準にチェック範囲を作り直してからリネームする */
        if (settings) {
            executeRename(doc, buildReorderedSettings(settings, itemEntries, itemCount), { silent: true });
        }

        /* 手動上書きを適用する（move() 後の最新参照に追随するため docItems を取り直す） */
        docItems = getDocumentItems(doc, itemType);
        for (var posKey in userOverridesByNewPosition) {
            if (!userOverridesByNewPosition.hasOwnProperty(posKey)) continue;
            var positionIndex = parseInt(posKey, 10);
            if (positionIndex >= 0 && positionIndex < docItems.length) {
                setItemName(docItems[positionIndex], userOverridesByNewPosition[posKey], "manual override at " + positionIndex);
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
        /* 「元の名称」で何も入力していなければ名前は変わらない。リネームを走らせると
           重複した名前にだけ "_1" が付いてしまうので、入力なしとして扱う */
        var hasNoBaseText = settings.mode === "original" || (settings.mode === "custom" && !settings.customText);
        return hasNoBaseText
            && !settings.prefix
            && !settings.suffix
            && !settings.findText;
    }

    /**
     * 名前の基準ごとに、各アイテムのベース文字列を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {object} docItems - 対象アイテムのコレクション
     * @param {object} settings - 現在の設定
     * @param {string} itemType - 正規化済みの種類（呼び出し側で "artboard" に既定化したもの）
     * @returns {object} インデックスをキーにした文字列配列のマップ
     */
    function buildItemTextMap(doc, docItems, settings, itemType) {
        var itemTextMap = {};
        if (settings.mode === "frontmost" && itemType === "artboard") {
            return buildFrontmostTextMap(getCachedFrontmostTextFrames(doc));
        }
        for (var itemIdx = 0; itemIdx < docItems.length; itemIdx++) {
            if (settings.mode === "original") {
                itemTextMap[itemIdx] = [docItems[itemIdx].name];
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
     * @param {object} [renameOptions] - {silent: true} で警告を出さない
     * @returns {boolean} リネームを実行したら true
     */
    function executeRename(doc, settings, renameOptions) {
        if (hasNoRenameInput(settings)) {
            if (!(renameOptions && renameOptions.silent)) alert(getLabel("alert.needSettings"));
            return false;
        }

        var itemType = settings.itemType || "artboard";
        var docItems = getDocumentItems(doc, itemType);
        var itemTextMap = buildItemTextMap(doc, docItems, settings, itemType);
        var selectedIndices = getRangeItemIndices(docItems.length, settings.rangeMode, settings.rangeText);
        var renamePlan = buildRenamePlan(doc, docItems, itemTextMap, settings, selectedIndices);

        for (var finalIdx = 0; finalIdx < renamePlan.indices.length; finalIdx++) {
            setItemName(
                docItems[renamePlan.indices[finalIdx]],
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
        var docItems = getDocumentItems(doc, itemType);

        /* 既存の名前で初期化する（未選択アイテムは現在の名前を保つ） */
        var previewNames = [];
        for (var initIdx = 0; initIdx < docItems.length; initIdx++) {
            previewNames.push(docItems[initIdx].name);
        }
        if (hasNoRenameInput(settings)) return previewNames;

        var itemTextMap = buildItemTextMap(doc, docItems, settings, itemType);
        var selectedIndices = getRangeItemIndices(docItems.length, settings.rangeMode, settings.rangeText);
        var renamePlan = buildRenamePlan(doc, docItems, itemTextMap, settings, selectedIndices);
        for (var finalIdx = 0; finalIdx < renamePlan.indices.length; finalIdx++) {
            previewNames[renamePlan.indices[finalIdx]] = renamePlan.names[finalIdx];
        }
        return previewNames;
    }

    /**
     * 選択アイテムの最終名プランを作る（プレビューと本番で共用）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} docItems - 対象アイテムのコレクション
     * @param {object} itemTextMap - インデックスごとのベース文字列
     * @param {object} settings - 現在の設定
     * @param {Array<number>} selectedIndices - 対象インデックス
     * @returns {{indices: Array<number>, names: Array<string>}} リネーム対象と最終名
     */
    function buildRenamePlan(doc, docItems, itemTextMap, settings, selectedIndices) {
        var prefixTemplate = settings.prefix || "";
        var suffixTemplate = settings.suffix || "";
        var findText = settings.findText || "";
        var replaceText = settings.replaceText || "";
        var useRegex = !!settings.useRegex;

        var skipUniquification = hasSequenceToken(prefixTemplate)
            || hasSequenceToken(suffixTemplate)
            || (findText && hasSequenceToken(replaceText));
        var reservedNames = getReservedItemNames(docItems, selectedIndices);
        var selectedIndexSet = makeIndexSet(selectedIndices);
        var tokenContext = createTokenContext(doc);
        var plannedBaseNames = [];
        var plannedIndices = [];
        var sequenceIndex = 1;

        for (var itemIdx = 0; itemIdx < docItems.length; itemIdx++) {
            if (!selectedIndexSet[itemIdx]) continue;
            var expandedPrefix = expandTemplateTokens(prefixTemplate, sequenceIndex, tokenContext);
            var expandedSuffix = expandTemplateTokens(suffixTemplate, sequenceIndex, tokenContext);
            var textPart = itemTextMap[itemIdx] ? itemTextMap[itemIdx].join(" ") : "";
            var baseName = expandedPrefix + textPart + expandedSuffix;

            /* 接頭辞・基準・接尾辞のどれも未指定なら、検索・置換だけで現在名を加工する */
            if (!baseName && findText) {
                baseName = docItems[itemIdx].name;
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
                reservedNames.push(docItems[itemIdx].name);
            }
        }

        return {
            indices: plannedIndices,
            names: resolveUniqueNames(plannedBaseNames, reservedNames, skipUniquification)
        };
    }

    /**
     * 検索・置換を適用する（正規表現対応）
     * @param {string} sourceName - 元の名前
     * @param {string} findPattern - 検索文字列
     * @param {string} replaceText - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @returns {string} 置換後の名前（不正な正規表現なら元の名前）
     */
    function applyFindReplace(sourceName, findPattern, replaceText, useRegex) {
        if (!findPattern) return sourceName;
        try {
            var findRegex;
            var effectiveReplace = replaceText;
            if (useRegex) {
                findRegex = new RegExp(findPattern, "g");
            } else {
                var escapedPattern = findPattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                findRegex = new RegExp(escapedPattern, "g");
                /* リテラル置換では replaceText 中の $ も無効化する（$1 や $& を特殊解釈させない） */
                effectiveReplace = replaceText.replace(/\$/g, "$$$$");
            }
            return sourceName.replace(findRegex, effectiveReplace);
        } catch (regexError) {
            logFailure("invalid find pattern '" + findPattern + "'", regexError);
            return sourceName;
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
        var renamableStyles = [];
        for (var styleIdx = 0; styleIdx < doc.graphicStyles.length; styleIdx++) {
            var graphicStyle = doc.graphicStyles[styleIdx];
            if (RESERVED_STYLE_NAME.test(graphicStyle.name)) continue;
            renamableStyles.push(graphicStyle);
        }
        return renamableStyles;
    }

    /**
     * 未選択アイテムの名前を予約名として返す（衝突回避用）
     * @param {object} docItems - 対象アイテムのコレクション
     * @param {Array<number>} selectedIndices - 選択インデックス
     * @returns {Array<string>} 予約名
     */
    function getReservedItemNames(docItems, selectedIndices) {
        var selectedIndexSet = makeIndexSet(selectedIndices);
        var reservedNames = [];
        for (var i = 0; i < docItems.length; i++) {
            if (!selectedIndexSet[i]) reservedNames.push(docItems[i].name);
        }
        return reservedNames;
    }

    /**
     * 名前をハッシュのキーとして安全な形にする
     * 接頭辞を付けないと `toString` や `valueOf` が Object.prototype のメンバーと衝突し、
     * 重複判定やカウンターが壊れる
     * @param {string} itemName - アイテム名
     * @returns {string} ハッシュ用のキー
     */
    function nameKey(itemName) {
        return "name:" + itemName;
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
     * @param {Document} doc - 対象ドキュメント
     * @returns {{fileName: string, dateString: string}} トークン展開用の値
     */
    function createTokenContext(doc) {
        var currentDate = new Date();
        return {
            fileName: doc.name.replace(/\.[^.]+$/, ""),
            dateString: currentDate.getFullYear().toString() +
                ("0" + (currentDate.getMonth() + 1)).slice(-2) +
                ("0" + currentDate.getDate()).slice(-2)
        };
    }

    /**
     * テンプレート文字列の連番・#FN・#DT トークンを展開する
     * @param {string} template - テンプレート文字列
     * @param {number} sequenceIndex - 1 始まりの連番
     * @param {object} tokenContext - createTokenContext() の戻り値
     * @returns {string} 展開後の文字列
     */
    function expandTemplateTokens(template, sequenceIndex, tokenContext) {

        /* 連番トークン {#N} を展開する（ゼロパディング対応：{#01} → 01, 02, ...） */
        var expandedText = template.replace(/\{#(\d+)\}/g, function (match, token) {
            var sequenceText = (parseInt(token, 10) + sequenceIndex - 1).toString();
            /* 足りない桁だけを埋める。slice で切り詰めると {#01} の100件目が "00" になる */
            if (token.charAt(0) === "0") {
                while (sequenceText.length < token.length) sequenceText = "0" + sequenceText;
            }
            return sequenceText;
        });

        /* 差し込む値に `$&` や `$$` が含まれても置換パターンとして解釈されないよう、
           文字列ではなく関数を渡す（ファイル名が "A$$B" や "Price$&List" のケース） */
        expandedText = expandedText.replace(/#FN/g, function () { return tokenContext.fileName; });
        expandedText = expandedText.replace(/#DT/g, function () { return tokenContext.dateString; });
        return expandedText;
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
        var rangeIndices = [];
        var rangeParts = rangeText.split(",");
        for (var i = 0; i < rangeParts.length; i++) {
            var rangePart = rangeParts[i].replace(/\s+/g, "");
            if (/^\d+$/.test(rangePart)) {
                rangeIndices.push(parseInt(rangePart, 10) - 1);
            } else if (/^\d+-\d+$/.test(rangePart)) {
                var rangeBounds = rangePart.split("-");
                var rangeStart = parseInt(rangeBounds[0], 10);
                var rangeEnd = parseInt(rangeBounds[1], 10);
                for (var j = rangeStart; j <= rangeEnd; j++) rangeIndices.push(j - 1);
            }
        }
        return rangeIndices;
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
        var sortedIndices = [];
        for (var i = 0; i < zeroBasedIndices.length; i++) sortedIndices.push(zeroBasedIndices[i]);
        sortedIndices.sort(function (a, b) { return a - b; });

        var rangeParts = [];
        var rangeStart = sortedIndices[0];
        var rangeEnd = sortedIndices[0];

        /**
         * 連続した並びを "start" または "start-end" として書き出す
         * @returns {void}
         */
        function pushRange() {
            rangeParts.push(rangeStart === rangeEnd
                ? (rangeStart + 1) + ""
                : (rangeStart + 1) + "-" + (rangeEnd + 1));
        }

        for (var j = 1; j < sortedIndices.length; j++) {
            if (sortedIndices[j] === rangeEnd + 1) {
                rangeEnd = sortedIndices[j];
            } else {
                pushRange();
                rangeStart = sortedIndices[j];
                rangeEnd = sortedIndices[j];
            }
        }
        pushRange();
        return rangeParts.join(",");
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
        var indexSet = {};
        for (var i = 0; i < indices.length; i++) indexSet[indices[i]] = true;
        return indexSet;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // テキストフレームの探索 / Text frame lookup
    // =========================================

    /**
     * テキストフレームの可視バウンズの中心座標を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Array<number>} [x, y] の中心座標
     */
    function getTextCenter(textFrame) {
        var textBounds = textFrame.visibleBounds;
        return [(textBounds[0] + textBounds[2]) / 2, (textBounds[1] + textBounds[3]) / 2];
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
     * アートボードごとの最前面 TextFrame から、名前の基準にする文字列のマップを作る
     * 走査で見つけたアートボードにそのまま結び付ける。中心座標で探し直すと、
     * アートボードが重なっているときに手前の番号のアートボードへ取られる
     * @param {Array<TextFrame|null>} frontmostFrames - アートボードの位置ごとの TextFrame（無ければ null）
     * @returns {object} アートボードのインデックスをキーにした文字列配列
     */
    function buildFrontmostTextMap(frontmostFrames) {
        var textMap = {};
        for (var i = 0; i < frontmostFrames.length; i++) {
            if (!frontmostFrames[i]) continue;
            textMap[i] = [frontmostFrames[i].contents.replace(/[\r\n\t]/g, "")];
        }
        return textMap;
    }

    /* 最前面テキストの走査結果キャッシュ / Cached frontmost-text scan
       走査結果は設定ではなくドキュメントにしか依存しないので、canvas を書き換えるまで使い回せる */
    var frontmostTextFrameCache = null;

    /**
     * 最前面 TextFrame の走査結果を返す（キャッシュがあれば再利用する）
     * 走査は O(アートボード数 × ページアイテム数) なので、入力のたびに回すと大きなドキュメントで止まる
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<TextFrame|null>} アートボードの位置ごとのテキストフレーム（無ければ null）
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
     * @returns {Array<TextFrame|null>} アートボードの位置ごとのテキストフレーム（無ければ null）
     */
    function getFrontmostTextFramesPerArtboard(doc) {
        /* 表示中・ロックなしのテキストを、レイヤー順・pageItems 順（サブレイヤーはページアイテムの後）で集める
           Visible, unlocked text in layer order and pageItems order (sublayers after page items) */
        var textFrames = collectSelectionTextFrames(doc.layers, { skipLocked: true, skipHidden: true, unique: false });

        var frontmostFrames = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboardBounds = doc.artboards[artboardIndex].artboardRect;
            var frontmostFrame = null;
            for (var frameIndex = 0; frameIndex < textFrames.length; frameIndex++) {
                if (isCenterInsideBounds(getTextCenter(textFrames[frameIndex]), artboardBounds)) {
                    frontmostFrame = textFrames[frameIndex];
                    break;
                }
            }
            frontmostFrames.push(frontmostFrame);   /* 見つからないアートボードも null で位置を詰めない */
        }
        return frontmostFrames;
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
