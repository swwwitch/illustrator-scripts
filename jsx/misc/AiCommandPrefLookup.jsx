#target illustrator
#targetengine "AiCommandPrefLookupEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

内蔵の一覧から項目を選んで、app.executeMenuCommand() / app.selectTool() / app.preferences の get・set のコードを出力します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCommandPrefLookup.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n0cf4826bf4a7

### Overview

Outputs app.executeMenuCommand() / app.selectTool() / app.preferences get/set code for items selected from the built-in list.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCommandPrefLookup.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiCommandPrefLookup";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-03";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCommandPrefLookup.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCommandPrefLookup.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0cf4826bf4a7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 種類ラジオボタンの順と初期値 / Kind radio order and default */
    var KIND_KEYS = ["menu", "pref", "tool"];
    var DEFAULT_KIND_INDEX = 0;

    /* 言語ラジオボタンの順と初期値 / Language radio order and default */
    var MENU_LANG_KEYS = ["ja", "en"];
    var DEFAULT_MENU_LANG_INDEX = 0;

    /* 名前をコメントで付ける（初期値）/ Add names as comments (default) */
    var DEFAULT_ADD_COMMENT = false;

    // =========================================
    // 一覧の書式 / List format
    // =========================================

    /* メニューコマンド・ツールの行：名前・ID・メモ（省略可、改行は \n）をタブで区切る
       Menu or tool line: name, ID and optional memo ("\n" for line breaks) */
    var MENU_LINE_PATTERN = /^([^\t]*)\t([^\t]*)(?:\t([^\t]*))?$/;

    /* 環境設定キーの行：名前・キー・型・値の例・メモ（省略可、改行は \n）をタブで区切る
       Preference line: name, key, type, sample value and optional memo ("\n" for line breaks) */
    var PREF_LINE_PATTERN = /^([^\t]*)\t([^\t]+)\t(Boolean|Integer|Real|String)\t([^\t]*)(?:\t([^\t]*))?$/;

    /* ID がこれに合えばツール（app.selectTool）/ IDs matching this are tools (app.selectTool) */
    var TOOL_ID_PATTERN = / Tool$/;

    /* 日本語の名前の判定（かな・漢字を含む）/ Japanese names contain kana or kanji */
    var JAPANESE_CHAR_PATTERN = /[\u3040-\u30FF\u3400-\u9FFF\uFF01-\uFF60]/;

    /* メニュー階層の区切り / Menu level separator */
    var MENU_LEVEL_SEPARATOR = " > ";

    /* 出力コードの書式（%ID% は ID、%TYPE% は型名、%VALUE% は値の例）/ Output code templates */
    var MENU_COMMAND_TEMPLATE = "app.executeMenuCommand('%ID%');";
    var TOOL_TEMPLATE         = "app.selectTool('%ID%');";
    var PREF_GET_TEMPLATE     = "app.preferences.get%TYPE%Preference('%ID%');";
    var PREF_SET_TEMPLATE     = "app.preferences.set%TYPE%Preference('%ID%', %VALUE%);";

    // =========================================
    // 再調査 / Recheck
    // =========================================

    /* 一覧を照合した Illustrator のメジャーバージョンと日付 / Version and date the list was last checked against */
    var DATA_CHECKED_VERSION = "30";
    var DATA_CHECKED_DATE    = "2026-09-27";

    /* 照合元の URL / Reference URLs */
    var REFERENCE_URLS = [
        { name: "Adobe Community", url: "https://community.adobe.com/questions-652/executemenucommand-command-list-797322" },
        { name: "Ai Command Palette", url: "https://github.com/joshbduncan/AiCommandPalette/tree/main/data" },
        { name: "sttk3 (Notion)", url: "https://judicious-night-bca.notion.site/app-executeMenuCommand-43b5a4b7a99d4ba2befd1798ba357b1a" },
        { name: "Ten A", url: "https://ten-artai.com/illustrator-ccver-22-menu-commands-list" }
    ];

    /* 環境設定ファイルの名前（英語版・日本語版）/ Preference file names (English and Japanese) */
    var PREF_FILE_PATTERN = /^Adobe Illustrator (?:Cloud Prefs|Prefs|環境設定)$/;

    /* 追加候補から外すキー（アクション・ダイアログ履歴・プラグインのファイル一覧など）/ Keys excluded from the candidates */
    var PREF_IGNORE_PATTERN = /^(?:plugin\/Action\/|artnewdialog\/|plugin\/(?:Mixed)?FileList\/|GenAI\/)/;

    /* 再調査の段階（進行状況バー）：ファイルを探す・.kys・環境設定・照合 / Recheck steps for the progress bar */
    var RECHECK_STEP_COUNT = 4;

    // =========================================
    // レイアウト / Layout
    // =========================================
    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
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
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    var LIST_SIZE             = [560, 270];      /* リストの寸法 [幅,高さ] / list size */
    var COLUMN_WIDTHS         = [220, 320];      /* 列幅（項目・ID とキー）/ column widths (item, ID / key) */
    var NAME_COLUMN_MAX_CHARS = 20;              /* 1列目に出す項目名の最大文字数 / max characters shown in the item column */
    var FILTER_FIELD_WIDTH    = 300;             /* 絞り込み欄の幅 / width of the filter fields */
    var ROW_LABEL_WIDTH       = 90;              /* 行ラベルの幅 / row label width */
    var CODE_FIELD_HEIGHT     = 56;              /* コード欄（1段ぶん）の高さ / height of each code field */
    var MEMO_FIELD_HEIGHT     = 70;              /* メモ欄の高さ / memo field height */
    var RECHECK_RESULT_SIZE   = [620, 420];      /* 再調査の結果欄の寸法 [幅,高さ] / recheck result field size */
    var PROGRESS_BAR_WIDTH    = 320;             /* 進行状況バーの幅 / progress bar width */
    var COPY_BUTTON_SIZE      = 22;              /* コピーボタンの一辺 / copy button size */

    /* コピーアイコンの色 [r, g, b, a]（UI の明るさで切り替え）/ Copy icon colors by UI brightness */
    var COPY_ICON_COLORS = isDarkUI()
        ? { normal: [0.85, 0.85, 0.85, 1], pressed: [1, 1, 1, 1], dimmed: [0.85, 0.85, 0.85, 0.3] }
        : { normal: [0.13, 0.19, 0.31, 1], pressed: [0.15, 0.5, 0.92, 1], dimmed: [0.13, 0.19, 0.31, 0.25] };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        dialog: {
            title:         { ja: "ExtendScript コード辞典", en: "ExtendScript Code Reference" },
            recheckResult: { ja: "再調査の結果", en: "Recheck Result" },
            recheckProgress: { ja: "再調査", en: "Recheck" }
        },
        fieldLabel: {
            kind:       { ja: "種類", en: "Kind" },
            menuLang:   { ja: "言語", en: "Language" },
            category:   { ja: "カテゴリ", en: "Category" },
            keyword:    { ja: "キーワード", en: "Keyword" },
            item:       { ja: "項目", en: "Item" },
            idOrKey:    { ja: "ID・キー", en: "ID / Key" },
            itemName:   { ja: "名前", en: "Name" },
            code:       { ja: "コード", en: "Code" },
            getCode:    { ja: "get", en: "get" },
            setCode:    { ja: "set", en: "set" },
            memo:       { ja: "メモ", en: "Memo" },
            references: { ja: "照合元", en: "References" }
        },
        dropdown: {
            allCategories: { ja: "［すべて］", en: "[All]" },
            tools:         { ja: "ツール", en: "Tools" },
            others:        { ja: "その他", en: "Others" }
        },
        radio: {
            menu: { ja: "メニューコマンド", en: "Menu commands" },
            pref: { ja: "環境設定", en: "Preferences" },
            tool: { ja: "ツール", en: "Tools" },
            ja:   { ja: "日本語", en: "Japanese" },
            en:   { ja: "英語", en: "English" }
        },
        checkbox: {
            addComment: { ja: "名前をコメントで付ける", en: "Add names as comments" }
        },
        button: {
            close:      { ja: "閉じる", en: "Close" },
            recheck:    { ja: "再調査", en: "Recheck" },
            openUrl:    { ja: "開く", en: "Open" },
            openFolder: { ja: "設定フォルダーを開く", en: "Open Settings Folder" },
            saveReport: { ja: "書き出し...", en: "Save..." }
        },
        tooltip: {
            kind: {
                ja: "メニューコマンドは executeMenuCommand、ツールは selectTool、環境設定は app.preferences の get／set のコードを出します。",
                en: "Menu commands give executeMenuCommand, tools give selectTool, and preferences give app.preferences get/set code."
            },
            menuLang: { ja: "リストに出す名前の言語です。", en: "Language of the names shown in the list." },
            category: {
                ja: "メニューの最上位（ファイル、編集...）や環境設定の区分で絞り込みます。",
                en: "Filters by top-level menu (File, Edit...) or preference section."
            },
            keyword: {
                ja: "名前と ID・キーに含まれる文字で絞り込みます（大文字小文字を区別しません）。正規表現も使えます。入力するたびに絞り込みます。",
                en: "Filters by text contained in the name or ID / key (case-insensitive). Regular expressions are allowed. The list is filtered as you type."
            },
            list: {
                ja: "複数選択すると、選んだ順ではなくリストの順にコードを並べます。",
                en: "With multiple items selected, the code is listed in list order."
            },
            code: {
                ja: "選択した項目のコードです。右のボタンでコピーできます。環境設定では上が get、下が set です。",
                en: "Code for the selected items. Copy it with the button on the right. For preferences, the upper field is get and the lower is set."
            },
            copyCode:   { ja: "この欄のコードをクリップボードにコピー", en: "Copy this code to the clipboard" },
            addComment: { ja: "各行の末尾に「// 名前」を付けます。", en: "Appends \"// name\" to each line." },
            memo:       { ja: "値の意味や、反映の条件などの注意点です。", en: "What the values mean and caveats such as when changes take effect." },
            recheck: {
                ja: "実行中の Illustrator のショートカットファイルと環境設定ファイルを読み、一覧との差分を表示します。バージョンアップ後の一覧の更新に使います。",
                en: "Reads the shortcut and preference files of the running Illustrator and shows how they differ from the list. Use it to update the list after an upgrade."
            },
            references: { ja: "照合元の一覧をブラウザーで開きます。", en: "Opens the reference list in a browser." },
            saveReport: {
                ja: "結果と、追加の候補を一覧の書式にした行をテキストに書き出します。Claude Code に渡すと ExtendScript.txt の更新と埋め込み直しに使えます。",
                en: "Saves the result plus the candidates as lines in the list format. Hand the file to Claude Code to update ExtendScript.txt and re-embed the list."
            }
        },
        report: {
            checkedWith: { ja: "一覧の照合", en: "List checked with" },
            running:     { ja: "実行中", en: "Running" },
            versionDiff: {
                ja: "※ バージョンが変わっています。下の候補と照合元を確認し、一覧を更新してください。",
                en: "* The version has changed. Review the candidates and references below, then update the list."
            },
            settingsDir: { ja: "設定フォルダー", en: "Settings folder" },
            noFolder:    { ja: "（見つかりません）", en: "(not found)" },
            kysFiles:    { ja: "■ ショートカットファイル（.kys）", en: "■ Keyboard shortcut files (.kys)" },
            prefFiles:   { ja: "■ 環境設定ファイル", en: "■ Preference files" },
            noFiles: {
                ja: "（ありません。［キーボードショートカット］でセットを保存すると作られます）",
                en: "(none; saving a set in Keyboard Shortcuts creates one)"
            },
            menuMissing: {
                ja: "■ 一覧にあるがショートカットファイルに無いメニューコマンド（改名・廃止の候補）",
                en: "■ Menu commands in the list but not in the shortcut files (renamed/removed?)"
            },
            menuNew: {
                ja: "■ ショートカットファイルにあって一覧に無いメニューコマンド（追加の候補）",
                en: "■ Menu commands in the shortcut files but not in the list (to add?)"
            },
            toolMissing:  { ja: "■ 一覧にあるがショートカットファイルに無いツール", en: "■ Tools in the list but not in the shortcut files" },
            toolNew:      { ja: "■ ショートカットファイルにあって一覧に無いツール", en: "■ Tools in the shortcut files but not in the list" },
            prefNew: {
                ja: "■ 環境設定ファイルにあって一覧に無いキー（追加の候補）",
                en: "■ Keys in the preference files but not in the list (to add?)"
            },
            prefTypeDiff: { ja: "■ 型が一覧と食い違うキー", en: "■ Keys whose type differs from the list" },
            candidateLines: {
                ja: "■ 追加の候補（ExtendScript.txt の書式。名前は空）",
                en: "■ Candidates in the ExtendScript.txt format (names empty)"
            },
            handoff: {
                ja: "※ このファイルを Claude Code に渡し、照合元と突き合わせて ExtendScript.txt を更新・埋め込み直してもらう。",
                en: "* Hand this file to Claude Code to check it against the references, update ExtendScript.txt and re-embed the list."
            },
            none:         { ja: "（なし）", en: "(none)" },
            note: {
                ja: "※ ショートカットファイル・環境設定ファイルに無いことは廃止の根拠になりません（ショートカット対象外のコマンドや、一度も変更していない設定は書かれないため）。",
                en: "* Absence from these files does not prove removal (commands outside the shortcut list and never-changed settings are not written)."
            }
        },
        progress: {
            findFiles: { ja: "設定ファイルを探しています...", en: "Looking for settings files..." },
            readKys:   { ja: "ショートカットファイルを読み込み中...", en: "Reading shortcut files..." },
            readPrefs: { ja: "環境設定ファイルを読み込み中...", en: "Reading preference files..." },
            comparing: { ja: "一覧と照合しています...", en: "Comparing with the list..." }
        },
        alert: {
            copied:     { ja: "コピーしました。", en: "Copied." },
            copyFailed: { ja: "クリップボードにコピーできませんでした。", en: "Could not copy to the clipboard." }
        },
        counter: {
            itemUnit: { ja: "件", en: " items" }
        }
    };

    /**
     * 文字列を言語別の括弧で囲む（日本語は全角「（）」、英語は半角「 ()」）
     * @param {string} innerText - 括弧の中の文字列
     * @returns {string} 括弧付きの文字列
     */
    function wrapInParentheses(innerText) {
        return (uiLang === "ja") ? "（" + innerText + "）" : " (" + innerText + ")";
    }

    /**
     * 件数を「1,234件」「1,234 items」の形にする
     * @param {number} itemCount - 件数
     * @returns {string} 件数の文字列
     */
    function formatItemCount(itemCount) {
        return String(itemCount).replace(/(\d)(?=(\d\d\d)+$)/g, "$1,") + getLabel(LABELS.counter.itemUnit);
    }

    // =========================================
    // データ / Data
    // =========================================

    /**
     * @typedef {Object} CommandEntry
     * @property {string} itemName - 名前（「ファイル > 開く...」など。空のこともある）
     * @property {string} commandId - コマンド ID・ツール ID・環境設定キー
     * @property {string} kind - 種類（"menu" / "pref" / "tool"）
     * @property {string} prefType - 環境設定キーの型（"Boolean" / "Integer" / "Real" / "String"。キー以外は ""）
     * @property {string} prefValue - 環境設定キーの値の例（キー以外は ""）
     * @property {string} memo - 値の説明や注意点（なければ ""）
     * @property {string} nameLang - 名前の言語（"ja" / "en"。名前が空なら ""＝両方に出す）
     * @property {string} category - 絞り込み用のカテゴリ
     */

    /**
     * 前後の空白を取り除く
     * @param {string} sourceText - 元の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimText(sourceText) {
        return sourceText.replace(/^\s+|\s+$/g, "");
    }

    /**
     * 一覧の行からコマンドと環境設定キーを取り出す（形式に合わない行・ID が空の行・重複は除く）
     * @param {string[]} lines - 一覧の行
     * @returns {CommandEntry[]} コマンドの一覧（元の順）
     */
    function parseCommandList(lines) {
        var seenKeys = {};
        var entries = [];
        for (var i = 0; i < lines.length; i++) {
            var commandEntry = parsePrefLine(lines[i]) || parseMenuLine(lines[i]);
            if (!commandEntry) {
                continue;
            }
            var entryKey = commandEntry.itemName + "\t" + commandEntry.commandId;
            if (seenKeys[entryKey]) {
                continue;
            }
            seenKeys[entryKey] = true;
            commandEntry.nameLang = getNameLang(commandEntry.itemName);
            commandEntry.category = getCategoryName(commandEntry.itemName, commandEntry.kind);
            entries.push(commandEntry);
        }
        return entries;
    }

    /**
     * 一覧の1項目を作る（nameLang・category は呼び出し側で補う）
     * @param {string} itemName - 名前
     * @param {string} commandId - ID・キー
     * @param {string} kind - 種類
     * @param {string} [prefType] - 環境設定キーの型
     * @param {string} [prefValue] - 環境設定キーの値の例
     * @param {string} [memo] - メモ
     * @returns {CommandEntry} 項目
     */
    function createEntry(itemName, commandId, kind, prefType, prefValue, memo) {
        return {
            itemName: trimText(itemName),
            commandId: commandId,
            kind: kind,
            prefType: prefType || "",
            prefValue: prefValue || "",
            memo: memo || ""
        };
    }

    /**
     * 一覧のメモを読める形にする（一覧では改行を「\n」の2文字で書いてある）
     * @param {string|undefined} memoText - 一覧のメモ欄
     * @returns {string} 改行を戻したメモ（なければ ""）
     */
    function decodeMemo(memoText) {
        return (memoText || "").split("\\n").join("\n");
    }

    /**
     * メニューコマンド・ツールの行（名前<タブ>ID<タブ>メモ）を読む
     * @param {string} line - 一覧の1行
     * @returns {CommandEntry|null} 読めなければ null
     */
    function parseMenuLine(line) {
        var matched = MENU_LINE_PATTERN.exec(line);
        /* 末尾の空白も ID の一部（'Live PSAdapter_plugin_Ct  ' など）なので削らない / Trailing spaces are part of some IDs */
        if (!matched || trimText(matched[2]) === "") {
            return null;
        }
        var commandId = matched[2];
        return createEntry(matched[1], commandId, TOOL_ID_PATTERN.test(commandId) ? "tool" : "menu", "", "", decodeMemo(matched[3]));
    }

    /**
     * 環境設定キーの行（名前<タブ>キー<タブ>型<タブ>値の例<タブ>メモ）を読む
     * @param {string} line - 一覧の1行
     * @returns {CommandEntry|null} 読めなければ null
     */
    function parsePrefLine(line) {
        var matched = PREF_LINE_PATTERN.exec(line);
        if (!matched) {
            return null;
        }
        return createEntry(matched[1], matched[2], "pref", matched[3], matched[4], decodeMemo(matched[5]));
    }

    /**
     * 名前の言語を決める
     * @param {string} itemName - 名前
     * @returns {string} "ja" / "en"。名前が空なら ""
     */
    function getNameLang(itemName) {
        if (itemName === "") {
            return "";
        }
        return JAPANESE_CHAR_PATTERN.test(itemName) ? "ja" : "en";
    }

    /**
     * 絞り込み用のカテゴリを決める（メニュー階層の先頭。階層がなければ「ツール」か「その他」）
     * @param {string} itemName - 名前
     * @param {string} kind - 種類（"menu" / "pref" / "tool"）
     * @returns {string} カテゴリ名
     */
    function getCategoryName(itemName, kind) {
        if (itemName.indexOf(MENU_LEVEL_SEPARATOR) > 0) {
            return itemName.split(MENU_LEVEL_SEPARATOR)[0];
        }
        return getLabel((kind === "tool") ? LABELS.dropdown.tools : LABELS.dropdown.others);
    }

    /**
     * 種類と言語が条件に合うか判定する（名前が空のものはどちらの言語にも出す）
     * @param {CommandEntry} commandEntry - 対象の項目
     * @param {string} kind - 種類（"menu" / "pref" / "tool"）
     * @param {string} nameLang - 言語（"ja" / "en"）
     * @returns {boolean} 合えば true
     */
    function matchesKindAndLang(commandEntry, kind, nameLang) {
        return commandEntry.kind === kind
            && (commandEntry.nameLang === "" || commandEntry.nameLang === nameLang);
    }

    /**
     * 指定した種類・言語のカテゴリ名を重複なく、一覧に出てきた順で返す
     * @param {CommandEntry[]} entries - 一覧
     * @param {string} kind - 種類（"menu" / "pref" / "tool"）
     * @param {string} nameLang - 言語（"ja" / "en"）
     * @returns {string[]} カテゴリ名の一覧
     */
    function collectCategoryNames(entries, kind, nameLang) {
        var seenNames = {};
        var categoryNames = [];
        for (var i = 0; i < entries.length; i++) {
            if (matchesKindAndLang(entries[i], kind, nameLang) && !seenNames[entries[i].category]) {
                seenNames[entries[i].category] = true;
                categoryNames.push(entries[i].category);
            }
        }
        return categoryNames;
    }

    /**
     * キーワードから照合用の正規表現を作る（正規表現として不正なら文字どおりに探す）
     * @param {string} keywordText - キーワード
     * @returns {RegExp|null} 正規表現。空欄なら null
     */
    function buildKeywordPattern(keywordText) {
        var trimmedText = trimText(keywordText);
        if (trimmedText === "") {
            return null;
        }
        /* 入力途中の不正な正規表現は例外になる / Incomplete regex input throws */
        try {
            return new RegExp(trimmedText, "i");
        } catch (e) {
            return new RegExp(trimmedText.replace(/[\\^$.*+?()[\]{}|\/]/g, "\\$&"), "i");
        }
    }

    /**
     * 種類・言語・カテゴリ・キーワードで絞り込む
     * @param {CommandEntry[]} entries - 一覧
     * @param {Object} filterState - 絞り込み条件 { kind, nameLang, keywordText, categoryName（null ならすべて） }
     * @returns {CommandEntry[]} 条件に合う項目
     */
    function filterCommands(entries, filterState) {
        var keywordPattern = buildKeywordPattern(filterState.keywordText);
        var matchedEntries = [];
        for (var i = 0; i < entries.length; i++) {
            var commandEntry = entries[i];
            if (!matchesKindAndLang(commandEntry, filterState.kind, filterState.nameLang)) {
                continue;
            }
            if (filterState.categoryName !== null && commandEntry.category !== filterState.categoryName) {
                continue;
            }
            if (keywordPattern && !keywordPattern.test(commandEntry.itemName) && !keywordPattern.test(commandEntry.commandId)) {
                continue;
            }
            matchedEntries.push(commandEntry);
        }
        return matchedEntries;
    }

    /**
     * 1列目に出す項目名を作る（メニュー階層の末尾を切り詰める。名前がなければ ID）
     * Mac の ScriptUI は1列目の幅を中身の最も長い文字列まで広げ、columnWidths を無視するため
     * @param {CommandEntry} commandEntry - 対象の項目
     * @returns {string} 項目名
     */
    function buildItemColumnText(commandEntry) {
        var menuLevels = commandEntry.itemName.split(MENU_LEVEL_SEPARATOR);
        var lastLevelName = menuLevels[menuLevels.length - 1] || commandEntry.commandId;
        return (lastLevelName.length > NAME_COLUMN_MAX_CHARS)
            ? lastLevelName.substring(0, NAME_COLUMN_MAX_CHARS - 1) + "..."
            : lastLevelName;
    }

    // =========================================
    // コード生成 / Code generation
    // =========================================

    /**
     * 文字列を単一引用符のリテラル用にエスケープする
     * @param {string} sourceText - 元の文字列
     * @returns {string} \ と ' をエスケープした文字列
     */
    function escapeSingleQuoted(sourceText) {
        return sourceText.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
    }

    /**
     * 書式の %名前% を置き換える（split/join で置換文字列の $ を特殊文字として扱わせない）
     * @param {string} codeTemplate - 書式
     * @param {Object} replacements - { "%ID%": "…" } 形式の置き換え
     * @returns {string} 置き換えた文字列
     */
    function fillTemplate(codeTemplate, replacements) {
        var filledText = codeTemplate;
        for (var placeholder in replacements) {
            if (replacements.hasOwnProperty(placeholder)) {
                filledText = filledText.split(placeholder).join(replacements[placeholder]);
            }
        }
        return filledText;
    }

    /**
     * set に渡す値のリテラルを作る（Boolean は 0/1 を false/true に、String は引用符で囲む）
     * @param {CommandEntry} prefEntry - 環境設定キー
     * @returns {string} 値のリテラル
     */
    function buildPrefValueLiteral(prefEntry) {
        var sampleValue = prefEntry.prefValue;
        if (prefEntry.prefType === "Boolean") {
            return (sampleValue === "0" || sampleValue === "false") ? "false" : "true";
        }
        if (prefEntry.prefType === "String") {
            return "'" + escapeSingleQuoted(sampleValue) + "'";
        }
        if (sampleValue === "") {
            return (prefEntry.prefType === "Real") ? "1.0" : "1";
        }
        return sampleValue;
    }

    /**
     * 項目のコードを作る。環境設定キーは get を primaryCode、set を setCode に分け、それ以外は primaryCode だけ
     * @param {CommandEntry} commandEntry - 対象の項目
     * @param {boolean} addComment - 名前を行末コメントで付けるなら true
     * @returns {{primaryCode: string, setCode: string}} コード（setCode は環境設定キー以外では ""）
     */
    function buildCodeParts(commandEntry, addComment) {
        var commentText = (addComment && commandEntry.itemName !== "") ? " // " + commandEntry.itemName : "";
        var replacements = { "%ID%": escapeSingleQuoted(commandEntry.commandId) };
        if (commandEntry.kind === "pref") {
            replacements["%TYPE%"] = commandEntry.prefType;
            replacements["%VALUE%"] = buildPrefValueLiteral(commandEntry);
            return {
                primaryCode: fillTemplate(PREF_GET_TEMPLATE, replacements) + commentText,
                setCode: fillTemplate(PREF_SET_TEMPLATE, replacements) + commentText
            };
        }
        var codeTemplate = (commandEntry.kind === "tool") ? TOOL_TEMPLATE : MENU_COMMAND_TEMPLATE;
        return { primaryCode: fillTemplate(codeTemplate, replacements) + commentText, setCode: "" };
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // ダイアログ部品 / Dialog parts
    // =========================================

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    /**
     * ダイアログを作る（縦並び・共通の余白）
     * @param {string} dialogTitle - タイトル
     * @returns {Window} ダイアログ
     */
    function createDialogWindow(dialogTitle) {
        var dialogWindow = new Window("dialog", dialogTitle);
        setupWindow(dialogWindow);
        return dialogWindow;
    }

    /**
     * ラベルと入力部品を1行に並べる
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @returns {Group} 行グループ（入力部品はここへ追加する。ラベルは children[0]）
     */
    function addFieldRow(parent, labelSet) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = ["fill", "top"];
        rowGroup.alignChildren = ["left", "center"];
        var rowLabel = rowGroup.add("statictext", undefined, labelText(labelSet));
        /* 幅を固定しないと右揃えが効かない / Right alignment needs a fixed width */
        rowLabel.minimumSize.width = ROW_LABEL_WIDTH;
        rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowGroup;
    }

    /**
     * ラジオボタンを横に並べる（同じ行グループに入れて排他にする）
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @param {string[]} radioKeys - LABELS.radio のキー（並べる順）
     * @param {number} defaultIndex - 最初に選んでおく添字
     * @param {Object} tooltipSet - ツールチップ
     * @returns {RadioButton[]} radioKeys の順のラジオボタン
     */
    function addRadioRow(parent, labelSet, radioKeys, defaultIndex, tooltipSet) {
        var radioRowGroup = addFieldRow(parent, labelSet);
        var radioButtons = [];
        for (var i = 0; i < radioKeys.length; i++) {
            var radioButton = radioRowGroup.add("radiobutton", undefined, getLabel(LABELS.radio[radioKeys[i]]));
            radioButton.helpTip = getLabel(tooltipSet);
            radioButtons.push(radioButton);
        }
        radioButtons[defaultIndex].value = true;
        return radioButtons;
    }

    /**
     * 複数行の欄を1行に置く（ラベル・欄）
     * @param {Group} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @param {number} fieldHeight - 欄の高さ
     * @param {boolean} isReadOnly - 読み取り専用なら true
     * @returns {{rowGroup: Group, textField: EditText}} 作成した部品
     */
    function addMultilineRow(parent, labelSet, fieldHeight, isReadOnly) {
        var rowGroup = addFieldRow(parent, labelSet);
        rowGroup.alignChildren = ["left", "top"];
        var textField = rowGroup.add("edittext", undefined, "", { multiline: true, scrolling: true, readonly: isReadOnly });
        textField.alignment = ["fill", "fill"];
        textField.preferredSize.height = fieldHeight;
        return { rowGroup: rowGroup, textField: textField };
    }

    /**
     * カテゴリのドロップダウンの項目を入れ替える（先頭は［すべて］。候補が1つ以下なら無効）
     * @param {DropDownList} categoryDropdown - カテゴリのドロップダウン
     * @param {string[]} categoryNames - カテゴリ名の一覧
     * @returns {void}
     */
    function fillCategoryDropdown(categoryDropdown, categoryNames) {
        categoryDropdown.removeAll();
        categoryDropdown.add("item", getLabel(LABELS.dropdown.allCategories));
        for (var i = 0; i < categoryNames.length; i++) {
            categoryDropdown.add("item", categoryNames[i]);
        }
        categoryDropdown.selection = 0;
        categoryDropdown.enabled = (categoryNames.length > 1);
    }

    /**
     * 項目・ID とキーの2列のリストを作る（複数選択可）
     * @param {Group} parent - 追加先
     * @returns {ListBox} 一覧
     */
    function addCommandList(parent) {
        var commandList = parent.add("listbox", [0, 0, LIST_SIZE[0], LIST_SIZE[1]], "", {
            multiselect: true,
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [
                getLabel(LABELS.fieldLabel.item),
                getLabel(LABELS.fieldLabel.idOrKey)
            ],
            columnWidths: COLUMN_WIDTHS
        });
        commandList.preferredSize = LIST_SIZE;
        commandList.helpTip = getLabel(LABELS.tooltip.list);
        return commandList;
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // ファイルビューアで表示（再利用パーツ） / Show in file viewer (reusable)

    /**
     * Path Finder（起動中のとき）か Finder で、フォルダーを開くかファイルを選択して表示する。
     * 補助アプリ /Applications/OpenInFileViewer.app に一時ファイルでパスを渡して起動する。
     * 補助アプリは illustrator-scripts の helpers/OpenInFileViewer.applescript から作る
     * @param {File|Folder} targetItem - 開くフォルダーか、選択して表示するファイル
     * @returns {boolean} 補助アプリを起動できたら true。無い・起動できない・macOS 以外のときは false
     */
    function openInFileViewer(targetItem) {
        /* 定数は巻き上げで未定義にならないよう関数内に置く / Kept local so hoisting never leaves them undefined */
        var viewerAppPath = "/Applications/OpenInFileViewer.app";
        var pathFilePath = "/tmp/open_in_file_viewer_path.txt";

        if ($.os.indexOf("Mac") === -1) return false;
        /* .app は実体がディレクトリなので Folder でも確かめる / An .app is a directory, so check it as a Folder too */
        if (!new Folder(viewerAppPath).exists && !new File(viewerAppPath).exists) return false;

        var pathFile = new File(pathFilePath);
        var written = false;
        try {
            pathFile.encoding = "UTF-8";
            pathFile.lineFeed = "Unix";
            if (pathFile.open("w")) {
                /* fsName で ~ ではなく絶対パスを渡す / fsName gives the absolute POSIX path */
                written = pathFile.write(targetItem.fsName);
            }
        } catch (e) {
        } finally {
            try { pathFile.close(); } catch (closeError) {}
        }
        return written && new File(viewerAppPath).execute();
    }

    // ファイルビューアで表示（再利用パーツ）ここまで / End of the reusable file viewer

    // =========================================
    // コピーボタン / Copy button
    // =========================================

    /**
     * group の onDraw を呼び直す（group には notify() が無いため、隠して再表示する）
     * @param {Group} drawnGroup - 描き直す group
     * @returns {void}
     */
    function redrawGroup(drawnGroup) {
        drawnGroup.hide();
        drawnGroup.show();
    }

    /**
     * 角丸の長方形を塗る。ScriptUI は多角形を塗れないため、十字の長方形2つと角の円4つを重ねる
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {Object} iconBrush - 塗りのブラシ
     * @param {number[]} rectBounds - [左, 上, 幅, 高さ]
     * @param {number} cornerRadius - 角の半径
     * @returns {void}
     */
    function fillRoundedRect(iconGraphics, iconBrush, rectBounds, cornerRadius) {
        var left = rectBounds[0];
        var top = rectBounds[1];
        var width = rectBounds[2];
        var height = rectBounds[3];
        var diameter = cornerRadius * 2;
        var crossRects = [
            [left + cornerRadius, top, width - diameter, height],
            [left, top + cornerRadius, width, height - diameter]
        ];
        for (var i = 0; i < crossRects.length; i++) {
            iconGraphics.newPath();
            iconGraphics.rectPath(crossRects[i][0], crossRects[i][1], crossRects[i][2], crossRects[i][3]);
            iconGraphics.fillPath(iconBrush);
        }
        var cornerOrigins = [
            [left, top], [left + width - diameter, top],
            [left, top + height - diameter], [left + width - diameter, top + height - diameter]
        ];
        for (var j = 0; j < cornerOrigins.length; j++) {
            iconGraphics.newPath();
            iconGraphics.ellipsePath(cornerOrigins[j][0], cornerOrigins[j][1], diameter, diameter);
            iconGraphics.fillPath(iconBrush);
        }
    }

    /**
     * コピーのアイコン（右上に塗りの角丸四角、左下に L 字の背面）を描く。寸法は一辺 22 を基準に拡大縮小する
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconSize - 描画域の一辺
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawCopyIcon(iconGraphics, iconSize, iconColor) {
        var iconBrush = iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, iconColor);
        var scaleUnit = iconSize / 22;

        /**
         * 基準寸法を描画域の寸法にする / Scale a base measurement
         * @param {number} baseValue - 一辺 22 のときの寸法
         * @returns {number} 描画域での寸法（1以上）
         */
        function scaled(baseValue) {
            return Math.max(1, Math.round(baseValue * scaleUnit));
        }

        /* 背面の L 字（左の縦棒と下の横棒）/ Back sheet: left and bottom bars */
        var backBars = [[4, 7, 2, 11], [4, 16, 11, 2]];
        for (var i = 0; i < backBars.length; i++) {
            iconGraphics.newPath();
            iconGraphics.rectPath(scaled(backBars[i][0]), scaled(backBars[i][1]), scaled(backBars[i][2]), scaled(backBars[i][3]));
            iconGraphics.fillPath(iconBrush);
        }
        /* 前面の角丸四角 / Front sheet */
        fillRoundedRect(iconGraphics, iconBrush, [scaled(8), scaled(3), scaled(11), scaled(11)], scaled(2));
    }

    /**
     * onDraw で描くコピーボタンを追加する（押している間は色を変え、無効なら薄く描く）
     * @param {Group} parent - 追加先
     * @returns {Group} ボタンとして使う group（onClick を設定して使う）
     */
    function addCopyIconButton(parent) {
        var copyButton = parent.add("group");
        copyButton.preferredSize = [COPY_BUTTON_SIZE, COPY_BUTTON_SIZE];
        copyButton.minimumSize = [COPY_BUTTON_SIZE, COPY_BUTTON_SIZE];
        copyButton.maximumSize = [COPY_BUTTON_SIZE, COPY_BUTTON_SIZE];
        copyButton.isPressed = false;
        /* group の enabled は描画に反映されないので自前の状態で持つ / Own flag; group.enabled does not affect drawing */
        copyButton.isEnabled = true;
        copyButton.onClick = null;

        copyButton.onDraw = function () {
            var iconColor = !copyButton.isEnabled ? COPY_ICON_COLORS.dimmed
                : (copyButton.isPressed ? COPY_ICON_COLORS.pressed : COPY_ICON_COLORS.normal);
            drawCopyIcon(copyButton.graphics, COPY_BUTTON_SIZE, iconColor);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function setPressed(isPressed) {
            if (copyButton.isPressed !== isPressed) {
                copyButton.isPressed = isPressed;
                redrawGroup(copyButton);
            }
        }
        copyButton.addEventListener("mousedown", function () {
            if (copyButton.isEnabled) {
                setPressed(true);
            }
        });
        /* ボタンの上で離したときだけ実行する / Fire only when released over the button */
        copyButton.addEventListener("mouseup", function () {
            var wasPressed = copyButton.isPressed;
            setPressed(false);
            if (wasPressed && copyButton.onClick) {
                copyButton.onClick();
            }
        });
        copyButton.addEventListener("mouseout", function () {
            setPressed(false);
        });
        return copyButton;
    }

    /**
     * コピーボタンの有効／無効を切り替えて描き直す（自作描画は自動で薄くならないため）
     * @param {Group} copyButton - addCopyIconButton() の戻り値
     * @param {boolean} isEnabled - 有効なら true
     * @returns {void}
     */
    function setCopyButtonEnabled(copyButton, isEnabled) {
        if (copyButton.isEnabled !== isEnabled) {
            copyButton.isEnabled = isEnabled;
            redrawGroup(copyButton);
        }
    }

    /**
     * コード欄の1段（ラベル・複数行の欄・右端のコピーボタン）を作る。ボタンを押すとその欄をコピーする
     * @param {Window} parent - 追加先
     * @param {Object} labelSet - 行ラベル
     * @returns {{rowLabel: StaticText, codeField: EditText, btnCopy: Group}} 作成した部品
     */
    function addCodeRow(parent, labelSet) {
        var codeRow = addMultilineRow(parent, labelSet, CODE_FIELD_HEIGHT, false);
        codeRow.textField.helpTip = getLabel(LABELS.tooltip.code);
        var btnCopy = addCopyIconButton(codeRow.rowGroup);
        btnCopy.alignment = ["right", "top"];
        btnCopy.helpTip = getLabel(LABELS.tooltip.copyCode);
        btnCopy.onClick = function () {
            copyCodeField(codeRow.textField);
        };
        return { rowLabel: codeRow.rowGroup.children[0], codeField: codeRow.textField, btnCopy: btnCopy };
    }

    /**
     * コード欄の内容をコピーし、コピーした内容を添えて知らせる
     * @param {EditText} codeField - コピーする欄
     * @returns {void}
     */
    function copyCodeField(codeField) {
        if (codeField.text === "") {
            return;
        }
        if (copyTextToClipboard(codeField.text)) {
            alert(getLabel(LABELS.alert.copied) + "\n\n" + codeField.text);
        } else {
            alert(getLabel(LABELS.alert.copyFailed));
        }
    }

    // =========================================
    // クリップボード / Clipboard
    // =========================================

    /**
     * 文字列をクリップボードにコピーする
     * Illustrator には文字列を直接クリップボードへ送る API が無いため、一時テキストフレームを作ってコピーし、すぐ削除する
     * （ドキュメントが無ければ一時ドキュメントを作って閉じる）
     * @param {string} copyText - コピーする文字列（改行は \n）
     * @returns {boolean} コピーできたら true
     */
    function copyTextToClipboard(copyText) {
        var usingTempDoc = (app.documents.length === 0);
        var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;
        /* 文字編集中の選択は TextRange なので控えない / A text selection is a TextRange; do not save it */
        var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];
        /* ロック・非表示のレイヤーでは textFrames.add() が失敗するため一時的に解除する / Unlock and show the layer temporarily */
        var editLayer = targetDoc.activeLayer;
        var layerWasLocked = editLayer.locked;
        var layerWasVisible = editLayer.visible;
        var tempFrame = null;
        var copySucceeded = false;
        /* 親レイヤーのロックなどで textFrames.add() が失敗しうる / textFrames.add() can fail, e.g. under a locked parent layer */
        try {
            editLayer.locked = false;
            editLayer.visible = true;
            tempFrame = editLayer.textFrames.add();
            /* テキストフレームの改行は \r / Text frames use \r for line breaks */
            tempFrame.contents = copyText.split("\n").join("\r");
            /* app.copy() は黙って無視されることがあるため、再描画を挟んでメニューコマンドでコピーする */
            app.redraw();
            app.executeMenuCommand("deselectall");
            tempFrame.selected = true;
            app.redraw();
            app.executeMenuCommand("copy");
            /* コピー確定前に削除すると空になる / Deleting before the copy settles leaves it empty */
            app.redraw();
            copySucceeded = true;
        } catch (e) {
            copySucceeded = false;
        }
        if (tempFrame) {
            tempFrame.remove();
        }
        editLayer.locked = layerWasLocked;
        editLayer.visible = layerWasVisible;
        if (usingTempDoc) {
            targetDoc.close(SaveOptions.DONOTSAVECHANGES);
        } else {
            restoreSelection(savedSelection);
        }
        return copySucceeded;
    }

    /**
     * 選択を元に戻す（ロック・非表示などで選べない項目は飛ばす）
     * @param {PageItem[]} savedSelection - 元の選択
     * @returns {void}
     */
    function restoreSelection(savedSelection) {
        app.executeMenuCommand("deselectall");
        for (var i = 0; i < savedSelection.length; i++) {
            /* ロック・非表示の項目は選択の代入で例外になる / Locked or hidden items throw on selection */
            try {
                savedSelection[i].selected = true;
            } catch (e) {
                // 選べない項目は飛ばす / skip items that cannot be selected
            }
        }
        app.redraw();
    }

    // =========================================
    // 設定ファイルの読み込み / Settings files
    // =========================================

    /**
     * @typedef {Object} SettingsEntry
     * @property {string[]} path - キーの階層（例 ["Menus", "new", "Key"]）
     * @property {string} valueType - 値の型（"Integer" / "Real" / "String" / "Array" / "Other"）
     * @property {string} valueText - 値の文字列（ファイルに書かれたまま）
     */

    /**
     * 実行中の Illustrator の設定フォルダーを返す
     * Mac は ~/Library/Preferences、Windows は %APPDATA%\Adobe の下（Windows は x64 サブフォルダー）
     * @returns {Folder} 設定フォルダー（存在しないこともある）
     */
    function getSettingsFolder() {
        var majorVersion = String(app.version).split(".")[0];
        var settingsName = "Adobe Illustrator " + majorVersion + " Settings/" + app.locale;
        if (File.fs === "Macintosh") {
            return new Folder("~/Library/Preferences/" + settingsName);
        }
        var windowsBasePath = Folder.userData.fsName + "/Adobe/" + settingsName;
        var windowsFolder = new Folder(windowsBasePath + "/x64");
        return windowsFolder.exists ? windowsFolder : new Folder(windowsBasePath);
    }

    /**
     * 設定ファイル（キー・{ } の入れ子・[ ] の配列）を読んで、キーと値の型を並べる
     * @param {File} settingsFile - 読むファイル（.kys や環境設定ファイル）
     * @returns {SettingsEntry[]} キーの一覧。読めなければ空
     */
    function readSettingsFile(settingsFile) {
        settingsFile.encoding = "UTF-8";
        if (!settingsFile.open("r")) {
            return [];
        }
        var fileText = settingsFile.read();
        settingsFile.close();

        var fileLines = fileText.split(/\r\n|\r|\n/);
        var keyStack = [];
        var entries = [];
        var inArray = false;
        for (var i = 0; i < fileLines.length; i++) {
            var line = fileLines[i].replace(/^\s+/, "");
            if (inArray) {
                /* 配列は「]」の行まで読み飛ばす / Skip array data up to the "]" line */
                inArray = !/\]\s*$/.test(line);
                continue;
            }
            if (/^\}/.test(line)) {
                keyStack.pop();
                continue;
            }
            /* キーは「/名前」、名前の中の空白は「\ 」。空の名前（「/ {」）もある / A key is "/name"; names may be empty */
            var matched = /^\/((?:\\.|[^\s\\])*)\s*(.*)$/.exec(line);
            if (!matched) {
                continue;
            }
            var keyName = matched[1].replace(/\\(.)/g, "$1");
            var valueText = matched[2];
            if (valueText === "{") {
                keyStack.push(keyName);
                continue;
            }
            if (/^\[/.test(valueText)) {
                inArray = !/\]\s*$/.test(valueText);
            }
            entries.push({ path: keyStack.concat([keyName]), valueType: getSettingsValueType(valueText), valueText: valueText });
        }
        return entries;
    }

    /**
     * 設定ファイルの値の文字列から型を決める
     * @param {string} valueText - 値の文字列
     * @returns {string} "Integer" / "Real" / "String" / "Array" / "Other"
     */
    function getSettingsValueType(valueText) {
        if (/^-?\d+$/.test(valueText)) {
            return "Integer";
        }
        if (/^-?\d*\.\d+(?:e-?\d+)?$/.test(valueText)) {
            return "Real";
        }
        if (/^[("]/.test(valueText)) {
            return "String";
        }
        return /^\[/.test(valueText) ? "Array" : "Other";
    }

    /**
     * 設定フォルダー内のファイルを名前の条件で集める
     * @param {Folder} settingsFolder - 設定フォルダー
     * @param {RegExp} namePattern - 対象にするファイル名（デコード後の名前で判定）
     * @returns {File[]} 対象のファイル
     */
    function collectSettingsFiles(settingsFolder, namePattern) {
        if (!settingsFolder.exists) {
            return [];
        }
        var matchedFiles = [];
        var folderItems = settingsFolder.getFiles();
        for (var i = 0; i < folderItems.length; i++) {
            if (folderItems[i] instanceof File && namePattern.test(File.decode(folderItems[i].name))) {
                matchedFiles.push(folderItems[i]);
            }
        }
        return matchedFiles;
    }

    /**
     * ショートカットファイルからメニューコマンドとツールの ID を集める
     * @param {File[]} kysFiles - .kys ファイル
     * @returns {{menu: Object, tool: Object}} ID をキーにした表（値は true）
     */
    function collectShortcutIds(kysFiles) {
        var shortcutIds = { menu: {}, tool: {} };
        for (var i = 0; i < kysFiles.length; i++) {
            var entries = readSettingsFile(kysFiles[i]);
            for (var j = 0; j < entries.length; j++) {
                var keyPath = entries[j].path;
                /* 「/Menus { /ID { /Key … } }」の2段目が ID / The ID is the second level */
                if (keyPath.length !== 3 || keyPath[1] === "") {
                    continue;
                }
                if (keyPath[0] === "Menus") {
                    shortcutIds.menu[keyPath[1]] = true;
                } else if (keyPath[0] === "Tools" && TOOL_ID_PATTERN.test(keyPath[1])) {
                    shortcutIds.tool[keyPath[1]] = true;
                }
            }
        }
        return shortcutIds;
    }

    /**
     * 環境設定ファイルからキー・型・値を集める（配列・その他の型は除く）
     * @param {File[]} prefFiles - 環境設定ファイル
     * @returns {Object} 「階層/キー」をキー、{ valueType, valueText } を値にした表
     */
    function collectPrefKeys(prefFiles) {
        var prefKeys = {};
        for (var i = 0; i < prefFiles.length; i++) {
            var entries = readSettingsFile(prefFiles[i]);
            for (var j = 0; j < entries.length; j++) {
                var valueType = entries[j].valueType;
                if (valueType === "Integer" || valueType === "Real" || valueType === "String") {
                    prefKeys[entries[j].path.join("/")] = { valueType: valueType, valueText: entries[j].valueText };
                }
            }
        }
        return prefKeys;
    }

    // =========================================
    // 再調査の照合 / Recheck comparison
    // =========================================

    /**
     * 一覧の ID を種類ごとの表にする
     * @param {CommandEntry[]} allCommands - すべての項目
     * @returns {{menu: Object, tool: Object, pref: Object}} ID をキーにした表（pref の値は型）
     */
    function collectListIds(allCommands) {
        var listIds = { menu: {}, tool: {}, pref: {} };
        for (var i = 0; i < allCommands.length; i++) {
            var commandEntry = allCommands[i];
            listIds[commandEntry.kind][commandEntry.commandId] = commandEntry.prefType || true;
        }
        return listIds;
    }

    /**
     * 元の表にあって比べる表に無いキーを並べる
     * @param {Object} sourceTable - 元の表
     * @param {Object} compareTable - 比べる表
     * @param {RegExp} [ignorePattern] - 除外するキー
     * @returns {string[]} 差分のキー（昇順）
     */
    function listKeysNotIn(sourceTable, compareTable, ignorePattern) {
        var missingKeys = [];
        for (var keyName in sourceTable) {
            if (sourceTable.hasOwnProperty(keyName) && !compareTable.hasOwnProperty(keyName)
                && !(ignorePattern && ignorePattern.test(keyName))) {
                missingKeys.push(keyName);
            }
        }
        return missingKeys.sort();
    }

    /**
     * 環境設定ファイルと一覧で型が食い違うキーを並べる（Boolean と Integer、Real と Integer は同じとみなす）
     * @param {Object} listPrefs - 一覧のキーと型
     * @param {Object} filePrefs - collectPrefKeys() の戻り値
     * @returns {string[]} 「キー（一覧の型 → ファイルの型）」の一覧
     */
    function listPrefTypeDiffs(listPrefs, filePrefs) {
        var typeDiffs = [];
        for (var keyName in listPrefs) {
            if (!listPrefs.hasOwnProperty(keyName) || !filePrefs.hasOwnProperty(keyName)) {
                continue;
            }
            var listType = listPrefs[keyName];
            var fileType = filePrefs[keyName].valueType;
            var isCompatible = (listType === fileType)
                || (fileType === "Integer" && (listType === "Boolean" || listType === "Real"));
            if (!isCompatible) {
                typeDiffs.push(keyName + wrapInParentheses(listType + " → " + fileType));
            }
        }
        return typeDiffs.sort();
    }

    /**
     * 見出しと項目の一覧を報告用の文字列にする
     * @param {Object} headingSet - 見出しのラベル
     * @param {string[]} reportItems - 項目
     * @returns {string} 見出し・件数・項目（1行1件）
     */
    function formatReportSection(headingSet, reportItems) {
        var sectionLines = [labelValueText(headingSet, formatItemCount(reportItems.length))];
        if (reportItems.length === 0) {
            sectionLines.push(getLabel(LABELS.report.none));
        }
        return sectionLines.concat(reportItems).join("\n");
    }

    /**
     * 見出しとファイル名の一覧を報告用の文字列にする
     * @param {Object} headingSet - 見出しのラベル
     * @param {File[]} files - ファイル
     * @returns {string} 見出しとファイル名（ない場合は案内文）
     */
    function formatFileSection(headingSet, files) {
        var sectionLines = [getLabel(headingSet)];
        for (var i = 0; i < files.length; i++) {
            sectionLines.push(File.decode(files[i].name));
        }
        if (files.length === 0) {
            sectionLines.push(getLabel(LABELS.report.noFiles));
        }
        return sectionLines.join("\n");
    }

    /**
     * 報告の冒頭（照合したバージョン・実行中のバージョン・設定フォルダー）を作る
     * @param {Folder} settingsFolder - 設定フォルダー
     * @returns {string[]} 冒頭の段落
     */
    function buildReportHeader(settingsFolder) {
        var runningVersion = String(app.version);
        var headerParts = [
            labelValueText(LABELS.report.checkedWith, "Ai " + DATA_CHECKED_VERSION + wrapInParentheses(DATA_CHECKED_DATE))
                + (uiLang === "ja" ? "／" : " / ") + labelValueText(LABELS.report.running, "Ai " + runningVersion)
        ];
        if (runningVersion.split(".")[0] !== DATA_CHECKED_VERSION) {
            headerParts.push(getLabel(LABELS.report.versionDiff));
        }
        headerParts.push(labelValueText(LABELS.report.settingsDir,
            settingsFolder.exists ? settingsFolder.fsName : getLabel(LABELS.report.noFolder)));
        return headerParts;
    }

    /**
     * 追加の候補を一覧の行（ExtendScript.txt と同じ書式、名前は空）にする
     * @param {string[]} menuIds - メニューコマンドの ID
     * @param {string[]} toolIds - ツールの ID
     * @param {string[]} prefKeyNames - 環境設定キー
     * @param {Object} filePrefs - collectPrefKeys() の戻り値（型と値の例に使う）
     * @returns {string[]} 一覧の行
     */
    function buildCandidateLines(menuIds, toolIds, prefKeyNames, filePrefs) {
        var candidateLines = [];
        var idLists = [menuIds, toolIds];
        for (var i = 0; i < idLists.length; i++) {
            for (var j = 0; j < idLists[i].length; j++) {
                candidateLines.push("\t" + idLists[i][j]);
            }
        }
        var addedNote = "Ai " + app.version + " の再調査で追加";
        for (var k = 0; k < prefKeyNames.length; k++) {
            var filePref = filePrefs[prefKeyNames[k]];
            /* 文字列の値は ( ) や " " で囲まれている / String values are wrapped in ( ) or " " */
            var sampleValue = (filePref.valueType === "String")
                ? filePref.valueText.replace(/^[("]([\s\S]*)[)"]$/, "$1")
                : filePref.valueText;
            candidateLines.push(["", prefKeyNames[k], filePref.valueType, sampleValue.replace(/\t/g, " "), addedNote].join("\t"));
        }
        return candidateLines;
    }

    /**
     * @typedef {Object} RecheckResult
     * @property {string} reportText - 報告（\n 区切り）
     * @property {string[]} candidateLines - 一覧に加えられる追加の候補（一覧の行の書式）
     */

    /**
     * 実行中の Illustrator の設定ファイルと一覧を照合した報告と、追加の候補を作る
     * @param {CommandEntry[]} allCommands - すべての項目
     * @param {Folder} settingsFolder - 設定フォルダー
     * @param {Object} progressWindow - createProgressWindow() の戻り値
     * @returns {RecheckResult} 照合の結果
     */
    function buildRecheckReport(allCommands, settingsFolder, progressWindow) {
        progressWindow.step(LABELS.progress.findFiles);
        var kysFiles = collectSettingsFiles(settingsFolder, /\.kys$/i);
        var prefFiles = collectSettingsFiles(settingsFolder, PREF_FILE_PATTERN);
        progressWindow.step(LABELS.progress.readKys);
        var shortcutIds = collectShortcutIds(kysFiles);
        progressWindow.step(LABELS.progress.readPrefs);
        var filePrefs = collectPrefKeys(prefFiles);
        progressWindow.step(LABELS.progress.comparing);
        var listIds = collectListIds(allCommands);

        var newMenuIds = listKeysNotIn(shortcutIds.menu, listIds.menu);
        var newToolIds = listKeysNotIn(shortcutIds.tool, listIds.tool);
        var newPrefKeys = listKeysNotIn(filePrefs, listIds.pref, PREF_IGNORE_PATTERN);

        var reportParts = buildReportHeader(settingsFolder);
        reportParts.push(formatFileSection(LABELS.report.kysFiles, kysFiles));
        if (kysFiles.length > 0) {
            reportParts.push(formatReportSection(LABELS.report.menuMissing, listKeysNotIn(listIds.menu, shortcutIds.menu)));
            reportParts.push(formatReportSection(LABELS.report.menuNew, newMenuIds));
            reportParts.push(formatReportSection(LABELS.report.toolMissing, listKeysNotIn(listIds.tool, shortcutIds.tool)));
            reportParts.push(formatReportSection(LABELS.report.toolNew, newToolIds));
        }
        reportParts.push(formatFileSection(LABELS.report.prefFiles, prefFiles));
        if (prefFiles.length > 0) {
            reportParts.push(formatReportSection(LABELS.report.prefNew, newPrefKeys));
            reportParts.push(formatReportSection(LABELS.report.prefTypeDiff, listPrefTypeDiffs(listIds.pref, filePrefs)));
        }
        reportParts.push(getLabel(LABELS.report.note));
        return {
            reportText: reportParts.join("\n\n"),
            candidateLines: buildCandidateLines(newMenuIds, newToolIds, newPrefKeys, filePrefs)
        };
    }

    // =========================================
    // 再調査の表示 / Recheck display
    // =========================================

    /**
     * 進行状況バーの小さなパレットを表示する
     * @param {number} stepCount - 段階の数（バーの最大値）
     * @returns {{step: function(Object): void, close: function(): void}} 操作用の関数（close は2回呼んでもよい）
     */
    function createProgressWindow(stepCount) {
        var progressPalette = new Window("palette", getLabel(LABELS.dialog.recheckProgress));
        setupWindow(progressPalette);

        var messageText = progressPalette.add("statictext", undefined, getLabel(LABELS.progress.findFiles));
        /* 文言が入れ替わっても切れないよう幅を確保 / Reserve width so longer messages are not cut off */
        messageText.preferredSize.width = PROGRESS_BAR_WIDTH;
        var progressBar = progressPalette.add("progressbar", undefined, 0, stepCount);
        progressBar.preferredSize.width = PROGRESS_BAR_WIDTH;

        progressPalette.show();
        progressPalette.update();
        var isClosed = false;

        return {
            step: function (messageSet) {
                messageText.text = getLabel(messageSet);
                progressBar.value = Math.min(progressBar.value + 1, stepCount);
                progressPalette.update();
            },
            close: function () {
                if (!isClosed) {
                    isClosed = true;
                    progressPalette.close();
                }
            }
        };
    }

    /**
     * 進行状況バーを出しながら照合し、報告を返す
     * @param {CommandEntry[]} allCommands - すべての項目
     * @param {Folder} settingsFolder - 設定フォルダー
     * @returns {RecheckResult} 照合の結果
     */
    function runRecheck(allCommands, settingsFolder) {
        var progressWindow = createProgressWindow(RECHECK_STEP_COUNT);
        /* 途中で例外が起きても進行状況ウィンドウを残さない / Never leave the progress window open on error */
        try {
            return buildRecheckReport(allCommands, settingsFolder, progressWindow);
        } finally {
            progressWindow.close();
        }
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openUrlInBrowser(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/" + SCRIPT_NAME + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    /**
     * 書き出す内容を作る（報告・追加の候補の行・Claude Code への引き継ぎの案内）
     * @param {RecheckResult} recheckResult - 照合の結果
     * @returns {string} 書き出す文字列
     */
    function buildExportText(recheckResult) {
        return [
            recheckResult.reportText,
            labelValueText(LABELS.report.candidateLines, formatItemCount(recheckResult.candidateLines.length))
                + "\n" + recheckResult.candidateLines.join("\n"),
            getLabel(LABELS.report.handoff)
        ].join("\n\n");
    }

    /**
     * 報告をテキストファイルに書き出す（保存先はダイアログで選ぶ）
     * @param {string} reportText - 報告
     * @returns {void}
     */
    function saveReportFile(reportText) {
        var defaultFile = new File(Folder.desktop.fsName + "/" + SCRIPT_NAME + "-recheck-" + app.version + ".txt");
        var targetFile = defaultFile.saveDlg(getLabel(LABELS.button.saveReport));
        if (!targetFile) {
            return;
        }
        targetFile.encoding = "UTF-8";
        targetFile.lineFeed = "Unix";
        targetFile.open("w");
        targetFile.write(reportText);
        targetFile.close();
    }

    /**
     * 照合元の URL を選んで開く行を作る
     * @param {Window} parent - 追加先
     * @returns {void}
     */
    function addReferenceRow(parent) {
        var referenceRow = parent.add("group");
        referenceRow.orientation = "row";
        referenceRow.alignChildren = ["left", "center"];
        referenceRow.add("statictext", undefined, labelText(LABELS.fieldLabel.references));
        var referenceDropdown = referenceRow.add("dropdownlist", undefined, []);
        for (var i = 0; i < REFERENCE_URLS.length; i++) {
            referenceDropdown.add("item", REFERENCE_URLS[i].name);
        }
        referenceDropdown.selection = 0;
        referenceDropdown.preferredSize.width = FILTER_FIELD_WIDTH;
        var btnOpenUrl = referenceRow.add("button", undefined, getLabel(LABELS.button.openUrl));
        btnOpenUrl.helpTip = getLabel(LABELS.tooltip.references);
        btnOpenUrl.onClick = function () {
            openUrlInBrowser(REFERENCE_URLS[referenceDropdown.selection.index].url);
        };
    }

    /**
     * 再調査を実行し、結果のダイアログ（結果・照合元の URL・設定フォルダー・書き出し）を表示する
     * @param {CommandEntry[]} allCommands - すべての項目
     * @returns {void}
     */
    function showRecheckDialog(allCommands) {
        var settingsFolder = getSettingsFolder();
        var recheckResult = runRecheck(allCommands, settingsFolder);
        var reportText = recheckResult.reportText;

        var recheckDialog = createDialogWindow(getLabel(LABELS.dialog.recheckResult) + " " + SCRIPT_VERSION);

        /* 結果を最初に大きく出す / Show the result first, large */
        var resultField = recheckDialog.add("edittext", undefined, reportText, { multiline: true, scrolling: true, readonly: true });
        resultField.alignment = ["fill", "fill"];
        resultField.preferredSize = RECHECK_RESULT_SIZE;

        /* 照合元の URL は補助として下に置く / Reference URLs sit below as a secondary tool */
        addReferenceRow(recheckDialog);

        var buttonRow = addButtonRow(recheckDialog);
        /* 右端に［閉じる］（デフォルトボタン兼キャンセルボタン）/ Close at the right end, both default and cancel */
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "ok" });
        recheckDialog.defaultElement = btnClose;
        recheckDialog.cancelElement = btnClose;
        var btnOpenFolder = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.openFolder));
        btnOpenFolder.enabled = settingsFolder.exists;
        btnOpenFolder.onClick = function () {
            if (!openInFileViewer(settingsFolder)) settingsFolder.execute();
        };
        var btnSaveReport = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.saveReport));
        btnSaveReport.helpTip = getLabel(LABELS.tooltip.saveReport);
        btnSaveReport.onClick = function () {
            saveReportFile(buildExportText(recheckResult));
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(recheckDialog, SCRIPT_NAME + "_recheck");
        recheckDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メインダイアログの部品を作る（上から絞り込み条件・リスト・名前・コード・メモ・ボタン）
     * @param {Window} mainDialog - 追加先のダイアログ
     * @param {Object} viewState - 表示の状態 { selectedKind, selectedNameLang, addComment }
     * @param {CommandEntry[]} allCommands - すべての項目
     * @returns {Object} 作成した部品
     */
    function buildMainControls(mainDialog, viewState, allCommands) {
        var mainControls = {};
        mainControls.kindRadios = addRadioRow(mainDialog, LABELS.fieldLabel.kind, KIND_KEYS, DEFAULT_KIND_INDEX, LABELS.tooltip.kind);
        mainControls.nameLangRadios = addRadioRow(mainDialog, LABELS.fieldLabel.menuLang, MENU_LANG_KEYS, DEFAULT_MENU_LANG_INDEX, LABELS.tooltip.menuLang);

        mainControls.categoryDropdown = addFieldRow(mainDialog, LABELS.fieldLabel.category).add("dropdownlist", undefined, []);
        mainControls.categoryDropdown.preferredSize.width = FILTER_FIELD_WIDTH;
        mainControls.categoryDropdown.helpTip = getLabel(LABELS.tooltip.category);
        fillCategoryDropdown(mainControls.categoryDropdown,
            collectCategoryNames(allCommands, viewState.selectedKind, viewState.selectedNameLang));

        mainControls.keywordInput = addFieldRow(mainDialog, LABELS.fieldLabel.keyword).add("edittext", undefined, "");
        mainControls.keywordInput.preferredSize.width = FILTER_FIELD_WIDTH;
        mainControls.keywordInput.helpTip = getLabel(LABELS.tooltip.keyword);
        mainControls.keywordInput.active = true;

        mainControls.commandList = addCommandList(mainDialog);

        mainControls.nameField = addFieldRow(mainDialog, LABELS.fieldLabel.itemName).add("edittext", undefined, "", { readonly: true });
        mainControls.nameField.alignment = ["fill", "center"];

        /* コード欄は2段。環境設定キーでは上が get・下が set、それ以外は上だけ使う
           Two code rows: get / set for preference keys, only the first one otherwise */
        mainControls.primaryCodeRow = addCodeRow(mainDialog, LABELS.fieldLabel.code);
        mainControls.setCodeRow = addCodeRow(mainDialog, LABELS.fieldLabel.setCode);

        /* ラベル幅ぶん字下げしてコードの欄にそろえる / Indent by the label width to line up with the code field */
        var optionGroup = mainDialog.add("group");
        optionGroup.margins = [ROW_LABEL_WIDTH + 10, 0, 0, 0];
        optionGroup.alignChildren = ["left", "center"];
        mainControls.addCommentCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.addComment));
        mainControls.addCommentCheckbox.value = viewState.addComment;
        mainControls.addCommentCheckbox.helpTip = getLabel(LABELS.tooltip.addComment);

        mainControls.memoField = addMultilineRow(mainDialog, LABELS.fieldLabel.memo, MEMO_FIELD_HEIGHT, true).textField;
        mainControls.memoField.helpTip = getLabel(LABELS.tooltip.memo);

        var buttonRow = addButtonRow(mainDialog);
        /* 右端に［閉じる］（デフォルトボタン兼キャンセルボタン）/ Close at the right end, both default and cancel */
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "ok" });
        mainDialog.defaultElement = btnClose;
        mainDialog.cancelElement = btnClose;
        mainControls.btnRecheck = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.recheck));
        mainControls.btnRecheck.helpTip = getLabel(LABELS.tooltip.recheck);
        return mainControls;
    }

    /**
     * 選択中の項目から、名前・コード（get / set）・メモの表示内容を作る
     * @param {CommandEntry[]} selectedCommands - 選択中の項目（リストの順）
     * @param {boolean} addComment - 名前をコメントで付けるなら true
     * @returns {{nameText: string, primaryCodeText: string, setCodeText: string, memoText: string}} 表示内容
     */
    function buildOutputTexts(selectedCommands, addComment) {
        var itemNames = [];
        var primaryCodeLines = [];
        var setCodeLines = [];
        var memoTexts = [];
        for (var i = 0; i < selectedCommands.length; i++) {
            var commandEntry = selectedCommands[i];
            var codeParts = buildCodeParts(commandEntry, addComment);
            itemNames.push(commandEntry.itemName);
            primaryCodeLines.push(codeParts.primaryCode);
            if (codeParts.setCode !== "") {
                setCodeLines.push(codeParts.setCode);
            }
            if (commandEntry.memo !== "") {
                /* 複数選択ではどの項目のメモか分かるよう ID を添える / Prefix the ID when several are selected */
                memoTexts.push((selectedCommands.length > 1 ? "[" + commandEntry.commandId + "]\n" : "") + commandEntry.memo);
            }
        }
        /* 複数行 edittext の改行は \n / Multiline edittext uses \n */
        return {
            nameText: itemNames.join(" / "),
            primaryCodeText: primaryCodeLines.join("\n"),
            setCodeText: setCodeLines.join("\n"),
            memoText: memoTexts.join("\n\n")
        };
    }

    /**
     * メインダイアログを作って表示する
     * @param {CommandEntry[]} allCommands - すべての項目
     * @returns {void}
     */
    function showMainDialog(allCommands) {
        var visibleCommands = [];
        var lastKeywordText = "";
        /* show() 前の checkbox・radiobutton の value は読み戻せないので変数で持つ / Values cannot be read back before show() */
        var viewState = {
            selectedKind: KIND_KEYS[DEFAULT_KIND_INDEX],
            selectedNameLang: MENU_LANG_KEYS[DEFAULT_MENU_LANG_INDEX],
            addComment: DEFAULT_ADD_COMMENT
        };

        var mainDialog = createDialogWindow(getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        var mainControls = buildMainControls(mainDialog, viewState, allCommands);

        /**
         * 選択中の項目をリストの順で返す
         * @returns {CommandEntry[]} 選択中の項目
         */
        function getSelectedCommands() {
            var listSelection = mainControls.commandList.selection;
            if (!listSelection) {
                return [];
            }
            /* selection は選んだ順なので添字で並べ直す / selection is in click order; reorder by index */
            var selectedIndexes = [];
            for (var i = 0; i < listSelection.length; i++) {
                selectedIndexes.push(listSelection[i].index);
            }
            selectedIndexes.sort(function (a, b) { return a - b; });
            var selectedCommands = [];
            for (var j = 0; j < selectedIndexes.length; j++) {
                selectedCommands.push(visibleCommands[selectedIndexes[j]]);
            }
            return selectedCommands;
        }

        /**
         * 選択中の項目の名前・コード・メモを欄に出し、コード欄のラベルと有効／無効を種類に合わせる
         * @returns {void}
         */
        function updateOutput() {
            var outputTexts = buildOutputTexts(getSelectedCommands(), viewState.addComment);
            var isPref = (viewState.selectedKind === "pref");
            mainControls.nameField.text = outputTexts.nameText;
            mainControls.memoField.text = outputTexts.memoText;
            mainControls.primaryCodeRow.codeField.text = outputTexts.primaryCodeText;
            mainControls.setCodeRow.codeField.text = outputTexts.setCodeText;
            /* 環境設定キーは上が get・下が set、それ以外は下の欄を無効 / get / set for preferences; the set row is disabled otherwise */
            mainControls.primaryCodeRow.rowLabel.text = labelText(isPref ? LABELS.fieldLabel.getCode : LABELS.fieldLabel.code);
            mainControls.setCodeRow.codeField.enabled = isPref;
            setCopyButtonEnabled(mainControls.primaryCodeRow.btnCopy, outputTexts.primaryCodeText !== "");
            setCopyButtonEnabled(mainControls.setCodeRow.btnCopy, isPref && outputTexts.setCodeText !== "");
        }

        /**
         * 絞り込みの結果で一覧を作り直す
         * @returns {void}
         */
        function refreshList() {
            var categoryDropdown = mainControls.categoryDropdown;
            lastKeywordText = mainControls.keywordInput.text;
            visibleCommands = filterCommands(allCommands, {
                kind: viewState.selectedKind,
                nameLang: viewState.selectedNameLang,
                keywordText: lastKeywordText,
                categoryName: (categoryDropdown.selection.index === 0) ? null : categoryDropdown.selection.text
            });

            var commandList = mainControls.commandList;
            commandList.removeAll();
            for (var i = 0; i < visibleCommands.length; i++) {
                var listItem = commandList.add("item", buildItemColumnText(visibleCommands[i]));
                listItem.subItems[0].text = visibleCommands[i].commandId;
            }
            mainDialog.text = getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION + wrapInParentheses(formatItemCount(visibleCommands.length));
            /* 1件に絞れたら選択してコードを出す / Select it when only one item is left */
            if (visibleCommands.length === 1) {
                commandList.selection = 0;
            }
            updateOutput();
        }

        /**
         * 種類・言語を切り替えたら、カテゴリの候補を入れ替えて一覧を作り直す
         * @returns {void}
         */
        function refreshCategories() {
            var categoryDropdown = mainControls.categoryDropdown;
            /* 項目の入れ替えで onChange が走らないよう外しておく / Detach so refilling does not fire onChange */
            categoryDropdown.onChange = null;
            fillCategoryDropdown(categoryDropdown, collectCategoryNames(allCommands, viewState.selectedKind, viewState.selectedNameLang));
            categoryDropdown.onChange = refreshList;
            refreshList();
        }

        /**
         * ラジオボタンのクリックで viewState の値を変える処理を作る
         * @param {string} stateName - viewState のプロパティ名
         * @param {string} stateValue - 設定する値
         * @returns {function(): void} クリック処理
         */
        function createStateClickHandler(stateName, stateValue) {
            return function () {
                viewState[stateName] = stateValue;
                refreshCategories();
            };
        }

        for (var i = 0; i < mainControls.kindRadios.length; i++) {
            mainControls.kindRadios[i].onClick = createStateClickHandler("selectedKind", KIND_KEYS[i]);
        }
        for (var j = 0; j < mainControls.nameLangRadios.length; j++) {
            mainControls.nameLangRadios[j].onClick = createStateClickHandler("selectedNameLang", MENU_LANG_KEYS[j]);
        }

        /* 入力するたびに絞り込む。onChanging と onChange の両方から呼ばれても1回で済ませる（Enter は［閉じる］に使う）
           Filter as you type; run once even when both events fire (Enter is for Close) */
        mainControls.keywordInput.onChanging = mainControls.keywordInput.onChange = function () {
            if (mainControls.keywordInput.text !== lastKeywordText) {
                refreshList();
            }
        };
        mainControls.categoryDropdown.onChange = refreshList;
        mainControls.commandList.onChange = updateOutput;
        mainControls.addCommentCheckbox.onClick = function () {
            viewState.addComment = mainControls.addCommentCheckbox.value;
            updateOutput();
        };
        mainControls.btnRecheck.onClick = function () {
            showRecheckDialog(allCommands);
        };

        refreshList();
        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        mainDialog.show();
    }

    showMainDialog(parseCommandList(getCommandListLines()));

    // =========================================
    // 一覧データ / Command list
    // =========================================

    /**
     * メニューコマンド・ツール・環境設定キーの一覧を返す
     * （1行＝「名前<タブ>ID<タブ>メモ」、環境設定キーは「名前<タブ>キー<タブ>型<タブ>値の例<タブ>メモ」。メモは省略可）
     * 元データ：AT-doc/ExtendScript.txt
     * 照合元（2026-09-27 照合）:
     * - Adobe Community のスレッド: https://community.adobe.com/questions-652/executemenucommand-command-list-797322
     * - Ai Command Palette（menu_commands.csv / tool_commands.csv）: https://github.com/joshbduncan/AiCommandPalette/tree/main/data
     * - sttk3 の Notion データベース（日本語名・対応バージョンあり）: https://judicious-night-bca.notion.site/app-executeMenuCommand-43b5a4b7a99d4ba2befd1798ba357b1a
     * - Ten A の一覧（CC ver.22）: https://ten-artai.com/illustrator-ccver-22-menu-commands-list
     * @returns {string[]} 一覧の行
     */
    function getCommandListLines() {
        return [
            "ファイル > 新規...\tnew",
            "ファイル > テンプレートから新規...\tnewFromTemplate",
            "ファイル > 開く...\topen",
            "ファイル > Bridge で参照...\tAdobe Bridge Browse",
            "ファイル > 閉じる\tclose",
            "ファイル > 保存\tsave",
            "ファイル > 別名で保存...\tsaveas",
            "ファイル > 複製を保存...\tsaveacopy",
            "ファイル > テンプレートとして保存...\tsaveastemplate\tAi 30 のショートカットファイルでは saveasTemplate（旧 ID が動くかは未確認）",
            "ファイル > 書き出し > Web 用に保存 (従来)...\tAdobe AI Save For Web\t廃止：Ai 19.9 まで（Ai 16 から）",
            "ファイル > 選択したスライスを保存...\tAdobe AI Save Selected Slices",
            "ファイル > 復帰\trevert",
            "ファイル > Adobe Stock を検索...\tSearch Adobe Stock",
            "ファイル > 配置...\tAI Place",
            "ファイル > 書き出し > 書き出し形式...\texport\t廃止：Ai 19.9 まで（Ai 16 から）",
            "ファイル > 選択範囲を書き出し...\texportSelection",
            "ファイル > パッケージ...\tPackage Menu Item",
            "ファイル > 書き出し > スクリーン用に書き出し...\texportForScreens",
            "ファイル > スクリプト > その他のスクリプト...\tai_browse_for_script",
            "ファイル > ドキュメント設定...\tdocument",
            "ファイル > ドキュメントのカラーモード > CMYK カラー\tdoc-color-cmyk",
            "ファイル > ドキュメントのカラーモード > RGB カラー\tdoc-color-rgb",
            "ファイル > ファイル情報...\tFile Info",
            "ファイル > プリント...\tPrint",
            "ファイル > 終了\tquit",
            "Illustrator > Illustrator を終了\tquit",
            "編集 > 取り消し\tundo",
            "編集 > やり直し\tredo",
            "編集 > カット\tcut",
            "編集 > コピー\tcopy",
            "編集 > ペースト\tpaste",
            "編集 > 前面へペースト\tpasteFront",
            "編集 > 背面へペースト\tpasteBack",
            "編集 > 同じ位置にペースト\tpasteInPlace",
            "編集 > すべてのアートボードにペースト\tpasteInAllArtboard",
            "編集 > 書式なしでペースト\tpasteWithoutFormatting",
            "編集 > 消去\tclear",
            "編集 > 検索と置換...\tFind and Replace",
            "編集 > 次を検索\tFind Next",
            "編集 > スペルチェック\tCheck Spelling\t廃止：Ai 24.9 まで（Ai 16 から）",
            "編集 > カスタム辞書を編集...\tEdit Custom Dictionary...",
            "編集 > カラーを編集 > オブジェクトを再配色...\tRecolor Art Dialog",
            "編集 > カラーを編集 > カラーバランス調整...\tAdjust3",
            "編集 > カラーを編集 > 前後にブレンド\tColors3",
            "編集 > カラーを編集 > 左右にブレンド\tColors4",
            "編集 > カラーを編集 > 上下にブレンド\tColors5",
            "編集 > カラーを編集 > CMYK に変換\tColors8",
            "編集 > カラーを編集 > グレースケールに変換\tColors7",
            "編集 > カラーを編集 > RGB に変換\tColors9",
            "編集 > カラーを編集 > カラー反転\tColors6",
            "編集 > カラーを編集 > オーバープリントブラック...\tOverprint2",
            "編集 > カラーを編集 > 彩度調整...\tSaturate3",
            "編集 > オリジナルを編集\tEditOriginal Menu Item",
            "編集 > 透明の分割・統合プリセット...\tTransparency Presets",
            "編集 > プリントプリセット...\tPrint Presets",
            "編集 > Adobe PDF プリセット...\tPDF Presets",
            "編集 > 遠近グリッドプリセット...\tPerspectiveGridPresets",
            "編集 > カラー設定...\tcolor",
            "編集 > プロファイルの指定...\tassignprofile",
            "編集 > キーボードショートカット...\tKBSC Menu Item",
            "編集 > 環境設定 > 一般...\tpreference",
            "編集 > 環境設定 > 選択範囲・アンカー表示...\tselectPref\tAi 30 のショートカットファイルでは selectionPref（旧 ID が動くかは未確認）",
            "編集 > 環境設定 > テキスト...\tkeyboardPref",
            "編集 > 環境設定 > 単位...\tunitundoPref",
            "編集 > 環境設定 > ガイド・グリッド...\tguidegridPref",
            "編集 > 環境設定 > スマートガイド...\tsnapPref",
            "編集 > 環境設定 > スライス...\tslicePref",
            "編集 > 環境設定 > ハイフネーション...\thyphenPref",
            "編集 > 環境設定 > プラグイン・仮想記憶ディスク...\tpluginPref",
            "編集 > 環境設定 > ユーザーインターフェイス...\tUIPref\tAi 30 のショートカットファイルでは userInterfacePref（旧 ID が動くかは未確認）",
            "編集 > 環境設定 > パフォーマンス...\tGPUPerformancePref",
            "編集 > 環境設定 > ブラックのアピアランス...\tBlackPref",
            "Illustrator > 環境設定 > 一般...\tpreference",
            "Illustrator > 環境設定 > 選択範囲・アンカー表示...\tselectPref\tAi 30 のショートカットファイルでは selectionPref（旧 ID が動くかは未確認）",
            "Illustrator > 環境設定 > テキスト...\tkeyboardPref",
            "Illustrator > 環境設定 > 単位...\tunitundoPref",
            "Illustrator > 環境設定 > ガイド・グリッド...\tguidegridPref",
            "Illustrator > 環境設定 > スマートガイド...\tsnapPref",
            "Illustrator > 環境設定 > スライス...\tslicePref",
            "Illustrator > 環境設定 > ハイフネーション...\thyphenPref",
            "Illustrator > 環境設定 > プラグイン・仮想記憶ディスク...\tpluginPref",
            "Illustrator > 環境設定 > ユーザーインターフェイス...\tUIPref\tAi 30 のショートカットファイルでは userInterfacePref（旧 ID が動くかは未確認）",
            "Illustrator > 環境設定 > パフォーマンス...\tGPUPerformancePref",
            "Illustrator > 環境設定 > ファイル管理...\tFilePref",
            "Illustrator > 環境設定 > クリップボード...\tClipboardPref",
            "Illustrator > 環境設定 > ブラックのアピアランス...\tBlackPref",
            "オブジェクト > 変形 > 変形の繰り返し\ttransformagain",
            "オブジェクト > 変形 > 移動...\ttransformmove",
            "オブジェクト > 変形 > 回転...\ttransformrotate",
            "オブジェクト > 変形 > リフレクト...\ttransformreflect",
            "オブジェクト > 変形 > 拡大・縮小...\ttransformscale",
            "オブジェクト > 変形 > シアー...\ttransformshear",
            "オブジェクト > 変形 > 個別に変形...\tTransform v23",
            "オブジェクト > 変形 > バウンディングボックスのリセット\tAI Reset Bounding Box",
            "オブジェクト > 重ね順 > 最前面へ\tsendToFront",
            "オブジェクト > 重ね順 > 前面へ\tsendForward",
            "オブジェクト > 重ね順 > 背面へ\tsendBackward",
            "オブジェクト > 重ね順 > 最背面へ\tsendToBack",
            "オブジェクト > 重ね順 > 選択しているレイヤーに移動\tSelection Hat 2",
            "オブジェクト > 整列 > 水平方向左に整列\tHorizontal Align Left",
            "オブジェクト > 整列 > 水平方向中央に整列\tHorizontal Align Center",
            "オブジェクト > 整列 > 水平方向右に整列\tHorizontal Align Right",
            "オブジェクト > 整列 > 垂直方向上に整列\tVertical Align Top",
            "オブジェクト > 整列 > 垂直方向中央に整列\tVertical Align Center",
            "オブジェクト > 整列 > 垂直方向下に整列\tVertical Align Bottom",
            "オブジェクト > グループ\tgroup",
            "オブジェクト > グループ解除\tungroup",
            "オブジェクト > すべてグループ解除\tungroupAll",
            "オブジェクト > ロック > 選択\tlock",
            "オブジェクト > ロック > 前面のすべてのアートワーク\tSelection Hat 5",
            "オブジェクト > ロック > その他のレイヤー\tSelection Hat 7",
            "オブジェクト > すべてをロック解除\tunlockAll",
            "オブジェクト > 隠す > 選択\thide",
            "オブジェクト > 隠す > 前面のすべてのアートワーク\tSelection Hat 4",
            "オブジェクト > 隠す > その他のレイヤー\tSelection Hat 6",
            "オブジェクト > すべてを表示\tshowAll",
            "オブジェクト > 分割・拡張...\tExpand3",
            "オブジェクト > アピアランスを分割\texpandStyle",
            "オブジェクト > 画像の切り抜き\tCrop Image",
            "オブジェクト > ラスタライズ...\tRasterize 8 menu item",
            "オブジェクト > グラデーションメッシュを作成...\tmake mesh",
            "オブジェクト > モザイクオブジェクトを作成...\tAI Object Mosaic Plug-in4",
            "オブジェクト > 透明部分を分割・統合...\tFlatten Transparency",
            "オブジェクト > ピクセルグリッドに最適化\tMake Pixel Perfect",
            "オブジェクト > スライス > 作成\tAISlice Make Slice",
            "オブジェクト > スライス > 解除\tAISlice Release Slice",
            "オブジェクト > スライス > ガイドから作成\tAISlice Create from Guides",
            "オブジェクト > スライス > 選択範囲から作成\tAISlice Create from Selection",
            "オブジェクト > スライス > スライスを複製\tAISlice Duplicate",
            "オブジェクト > スライス > スライスを結合\tAISlice Combine",
            "オブジェクト > スライス > スライスを分割...\tAISlice Divide",
            "オブジェクト > スライス > すべてを削除\tAISlice Delete All Slices",
            "オブジェクト > スライス > スライスオプション...\tAISlice Slice Options",
            "オブジェクト > スライス > アートボードサイズでクリップ\tAISlice Clip to Artboard",
            "オブジェクト > トリムマークを作成\tTrimMark v25",
            "トンボ\tTrimMark v25",
            "オブジェクト > パス > 連結\tjoin",
            "オブジェクト > パス > 平均...\taverage",
            "オブジェクト > パス > パスのアウトライン\tOffsetPath v22",
            "オブジェクト > パス > パスのオフセット...\tOffsetPath v23",
            "オブジェクト > パス > パスの方向反転\tReverse Path Direction",
            "オブジェクト > パス > 単純化...\tsimplify menu item",
            "オブジェクト > パス > アンカーポイントの追加\tAdd Anchor Points2",
            "オブジェクト > パス > アンカーポイントを削除\tRemove Anchor Points menu",
            "オブジェクト > パス > 背面のオブジェクトを分割\tKnife Tool2",
            "オブジェクト > パス > グリッドに分割...\tRows and Columns....",
            "オブジェクト > パス > パスの削除...\tcleanup menu item",
            "オブジェクト > シェイプ > シェイプに変換\tConvert to Shape",
            "オブジェクト > シェイプ > シェイプを拡張\tExpand Shape",
            "オブジェクト > パターン > 作成\tAdobe Make Pattern",
            "オブジェクト > パターン > パターンを編集\tAdobe Edit Pattern",
            "オブジェクト > パターン > タイルの境界線カラー...\tAdobe Pattern Tile Color",
            "オブジェクト > リピート > ラジアル\tMake Radial Repeat",
            "オブジェクト > リピート > グリッド\tMake Grid Repeat",
            "オブジェクト > リピート > ミラー\tMake Symmetry Repeat",
            "オブジェクト > リピート > 解除\tRelease Repeat Art",
            "オブジェクト > リピート > オプション...\tRepeat Art Options",
            "オブジェクト > ブレンド > 作成\tPath Blend Make",
            "オブジェクト > ブレンド > 解除\tPath Blend Release",
            "オブジェクト > ブレンド > 拡張\tPath Blend Expand",
            "オブジェクト > ブレンド > ブレンドオプション...\tPath Blend Options",
            "オブジェクト > ブレンド > ブレンド軸を置き換え\tPath Blend Replace Spine",
            "オブジェクト > ブレンド > ブレンド軸を反転\tPath Blend Reverse Spine",
            "オブジェクト > ブレンド > 前後を反転\tPath Blend Reverse Stack",
            "オブジェクト > エンベロープ > ワープで作成...\tMake Warp",
            "オブジェクト > エンベロープ > メッシュで作成...\tCreate Envelope Grid",
            "オブジェクト > エンベロープ > 最前面のオブジェクトで作成\tMake Envelope",
            "オブジェクト > エンベロープ > 解除\tRelease Envelope",
            "オブジェクト > エンベロープ > エンベロープオプション...\tEnvelope Options",
            "オブジェクト > エンベロープ > 拡張\tExpand Envelope",
            "オブジェクト > エンベロープ > オブジェクトを編集\tEdit Envelope Contents",
            "オブジェクト > 遠近 > 選択面の図形にする\tAttach to Active Plane",
            "オブジェクト > 遠近 > 遠近グリッド上から解除\tRelease with Perspective",
            "オブジェクト > 遠近 > オブジェクトに合わせて面を移動\tShow Object Grid Plane",
            "オブジェクト > 遠近 > テキストを編集\tEdit Original Object",
            "オブジェクト > ライブペイント > 作成\tMake Planet X",
            "オブジェクト > ライブペイント > 結合\tMarge Planet X",
            "オブジェクト > ライブペイント > 解除\tRelease Planet X",
            "オブジェクト > ライブペイント > 隙間オプション...\tPlanet X Options",
            "オブジェクト > ライブペイント > 拡張\tExpand Planet X",
            "オブジェクト > 画像トレース > 作成\tMake Image Tracing",
            "オブジェクト > 画像トレース > 作成して拡張\tMake and Expand Image Tracing",
            "オブジェクト > 画像トレース > 解除\tRelease Image Tracing",
            "オブジェクト > 画像トレース > 拡張\tExpand Image Tracing",
            "オブジェクト > テキストの回り込み > 作成\tMake Text Wrap",
            "オブジェクト > テキストの回り込み > 解除\tRelease Text Wrap",
            "オブジェクト > テキストの回り込み > テキストの回り込みオプション...\tText Wrap Options...",
            "オブジェクト > クリッピングマスク > 作成\tmakeMask",
            "オブジェクト > クリッピングマスク > 解除\treleaseMask",
            "オブジェクト > クリッピングマスク > マスクを編集\teditMask",
            "オブジェクト > 複合パス > 作成\tcompoundPath",
            "オブジェクト > 複合パス > 解除\tnoCompoundPath",
            "オブジェクト > アートボード > アートボードに変換\tsetCropMarks",
            "オブジェクト > アートボード > すべてのアートボードを再配置\tReArrange Artboards\t廃止：Ai 29.5 まで（Ai 16 から）",
            "オブジェクト > アートボード > オブジェクト全体に合わせる\tFit Artboard to artwork bounds",
            "オブジェクト > アートボード > 選択オブジェクトに合わせる\tFit Artboard to selected Art",
            "オブジェクト > グラフ > 設定...\tsetGraphStyle",
            "オブジェクト > グラフ > データ...\teditGraphData",
            "オブジェクト > グラフ > デザイン...\tgraphDesigns",
            "オブジェクト > グラフ > 棒グラフ...\tsetBarDesign",
            "オブジェクト > グラフ > マーカー...\tsetIconDesign",
            "書式 > Adobe Fonts のその他のフォント...\tBrowse Typekit Fonts Menu IllustratorUI",
            "書式 > 字形\talternate glyph palette plugin",
            "書式 > エリア内文字オプション...\tarea-type-options\tAi 30 のショートカットファイルでは areatextoptions（旧 ID が動くかは未確認）",
            "書式 > パス上文字オプション > 虹\tRainbow\tAi 30 のショートカットファイルでは textpathtypeRainbow（旧 ID が動くかは未確認）",
            "書式 > パス上文字オプション > 歪み\tSkew\tAi 30 のショートカットファイルでは textpathtypeSkew（旧 ID が動くかは未確認）",
            "書式 > パス上文字オプション > 3D リボン\t3D ribbon\tAi 30 のショートカットファイルでは textpathtype3d（旧 ID が動くかは未確認）",
            "書式 > パス上文字オプション > 階段状\tStair Step\tAi 30 のショートカットファイルでは textpathtypestairs（旧 ID が動くかは未確認）",
            "書式 > パス上文字オプション > 引力\tGravity\tAi 30 のショートカットファイルでは textpathtypeGravity（旧 ID が動くかは未確認）",
            "書式 > パス上文字オプション > パス上文字オプション...\ttypeOnPathOptions\tAi 30 のショートカットファイルでは textpathtypeOptions（旧 ID が動くかは未確認）",
            "書式 > パス上文字オプション > パス上文字を更新\tupdateLegacyTOP",
            "書式 > 合成フォント...\tAdobe internal composite font plugin",
            "書式 > 禁則処理設定...\tAdobe Kinsoku Settings",
            "書式 > 文字組みアキ量設定...\tAdobe MojiKumi Settings",
            "書式 > スレッドテキストオプション > 作成\tthreadTextCreate",
            "書式 > スレッドテキストオプション > 選択部分をスレッドから除外\treleaseThreadedTextSelection",
            "書式 > スレッドテキストオプション > スレッドのリンクを解除\tremoveThreading",
            "書式 > 箇条書き > テキストに変換\tconvert list style to text",
            "書式 > ヘッドラインを合わせる\tfitHeadline",
            "書式 > アウトラインを作成\toutline",
            "書式 > フォント検索...\tAdobe Illustrator Find Font Menu Item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "書式 > 環境に無いフォントを解決する...\tAdobe IllustratorUI Resolve Missing Font",
            "書式 > 大文字と小文字の変更 > すべて大文字\tUpperCase Change Case Item",
            "書式 > 大文字と小文字の変更 > すべて小文字\tLowerCase Change Case Item",
            "書式 > 大文字と小文字の変更 > 単語の先頭のみ大文字\tTitle Case Change Case Item",
            "書式 > 大文字と小文字の変更 > 文頭のみ大文字\tSentence case Change Case Item",
            "書式 > 句読点の自動調節...\tAdobe Illustrator Smart Punctuation Menu Item",
            "書式 > 最適なマージン揃え\tAdobe Optical Alignment Item",
            "書式 > 制御文字を表示\tshowHiddenChar",
            "書式 > 組み方向 > 横組み\ttype-horizontal",
            "書式 > 組み方向 > 縦組み\ttype-vertical",
            "選択 > すべてを選択\tselectall",
            "選択 > 作業アートボードのすべてを選択\tselectallinartboard",
            "選択 > 選択を解除\tdeselectall",
            "選択 > 再選択\tFind Reselect menu item",
            "選択 > 選択範囲を反転\tInverse menu item",
            "選択 > 前面のオブジェクト\tSelection Hat 8",
            "選択 > 背面のオブジェクト\tSelection Hat 9",
            "選択 > 共通 > アピアランス\tFind Appearance menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > アピアランス属性\tFind Appearance Attributes menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > 描画モード\tFind Blending Mode menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > 塗りと線\tFind Fill & Stroke menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > カラー (塗り)\tFind Fill Color menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > 不透明度\tFind Opacity menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > カラー (線)\tFind Stroke Color menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > 線幅\tFind Stroke Weight menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > グラフィックスタイル\tFind Style menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプ\tFind Live Shape menu item\t廃止：Ai 25.9 まで（Ai 17 から）",
            "選択 > 共通 > シンボルインスタンス\tFind Symbol Instance menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > 一連のリンクブロック\tFind Link Block Series menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > テキスト > フォントファミリー\tFind Text Font Family menu item",
            "選択 > 共通 > テキスト > フォントファミリー (スタイル)\tFind Text Font Family Style menu item",
            "選択 > 共通 > テキスト > フォントファミリー (スタイルとサイズ)\tFind Text Font Family Style Size menu item",
            "選択 > 共通 > テキスト > フォントサイズ\tFind Text Font Size menu item",
            "選択 > 共通 > テキスト > テキストカラー (塗り)\tFind Text Fill Color menu item",
            "選択 > 共通 > テキスト > テキストカラー (線)\tFind Text Stroke Color menu item",
            "選択 > 共通 > テキスト > テキストカラー (塗りと線)\tFind Text Fill Stroke Color menu item",
            "選択 > オブジェクト > 同一レイヤー上のすべて\tSelection Hat 3",
            "選択 > オブジェクト > セグメント\tSelection Hat 1",
            "選択 > オブジェクト > 絵筆ブラシストローク\tBristle Brush Strokes menu item",
            "選択 > オブジェクト > ブラシストローク\tBrush Strokes menu item",
            "選択 > オブジェクト > クリッピングマスク\tClipping Masks menu item",
            "選択 > オブジェクト > 孤立点\tStray Points menu item",
            "選択 > オブジェクト > すべてのテキストオブジェクト\tText Objects menu item",
            "選択 > オブジェクト > ポイント文字オブジェクト\tPoint Text Objects menu item",
            "選択 > オブジェクト > エリア内文字オブジェクト\tArea Text Objects menu item",
            "選択 > オブジェクトを一括選択|選択解除\tSmartEdit Menu Item",
            "選択 > 選択範囲を保存...\tSelection Hat 10",
            "選択 > 選択範囲を編集...\tSelection Hat 11",
            "効果 > 前回の効果を適用\tAdobe Apply Last Effect",
            "効果 > 前回の効果\tAdobe Last Effect",
            "効果 > ドキュメントのラスタライズ効果設定...\tLive Rasterize Effect Setting",
            "効果 > 3D とマテリアル > 押し出し・ベベル\tLive Adobe Geometry3D Extrude",
            "効果 > 3D とマテリアル > 回転体\tLive Adobe Geometry3D Revolve",
            "効果 > 3D とマテリアル > 膨張\tLive Adobe Geometry3D Inflate",
            "効果 > 3D とマテリアル > 回転\tLive Adobe Geometry3D Rotate",
            "効果 > 3D とマテリアル > マテリアル\tLive Adobe Geometry3D Materials",
            "効果 > 3D > 押し出し・ベベル...\tLive 3DExtrude\t廃止：Ai 25.9 まで（Ai 16 から）",
            "効果 > 3D > 回転体...\tLive 3DRevolve\t廃止：Ai 25.9 まで（Ai 16 から）",
            "効果 > 3D > 回転...\tLive 3DRotate\t廃止：Ai 25.9 まで（Ai 16 から）",
            "効果 > SVG フィルター > SVG フィルターを適用...\tLive SVG Filters",
            "効果 > SVG フィルター > SVG フィルターの読み込み...\tSVG Filter Import",
            "効果 > スタイライズ > ぼかし...\tLive Feather",
            "効果 > スタイライズ > ドロップシャドウ...\tLive Adobe Drop Shadow",
            "効果 > スタイライズ > 光彩 (内側)...\tLive Inner Glow",
            "効果 > スタイライズ > 光彩 (外側)...\tLive Outer Glow",
            "効果 > スタイライズ > 落書き...\tLive Scribble Fill",
            "効果 > スタイライズ > 角を丸くする...\tLive Adobe Round Corners",
            "効果 > トリムマーク\tLive Trim Marks",
            "効果 > パス > オブジェクトのアウトライン\tLive Outline Object",
            "効果 > パス > パスのアウトライン\tLive Outline Stroke",
            "効果 > パス > パスのオフセット...\tLive Offset Path",
            "効果 > パスの変形 > ジグザグ...\tLive Zig Zag",
            "効果 > パスの変形 > パスの自由変形...\tLive Free Distort",
            "効果 > パスの変形 > パンク・膨張...\tLive Pucker & Bloat",
            "効果 > パスの変形 > ラフ...\tLive Roughen",
            "効果 > パスの変形 > ランダム・ひねり...\tLive Scribble and Tweak",
            "効果 > パスの変形 > 変形...\tLive Transform",
            "効果 > パスの変形 > 旋回...\tLive Twist",
            "効果 > パスファインダー > 追加\tLive Pathfinder Add\t廃止：Ai 28.9 まで（Ai 16 から）",
            "効果 > パスファインダー > 合体\tLive Pathfinder Add\t廃止：Ai 28.9 まで（Ai 16 から）",
            "効果 > パスファインダー > 交差\tLive Pathfinder Intersect",
            "効果 > パスファインダー > 中マド\tLive Pathfinder Exclude",
            "効果 > パスファインダー > 前面オブジェクトで型抜き\tLive Pathfinder Subtract",
            "効果 > パスファインダー > 背面オブジェクトで型抜き\tLive Pathfinder Minus Back",
            "効果 > パスファインダー > 分割\tLive Pathfinder Divide",
            "効果 > パスファインダー > 刈り込み\tLive Pathfinder Trim",
            "効果 > パスファインダー > 合流\tLive Pathfinder Merge",
            "効果 > パスファインダー > 切り抜き\tLive Pathfinder Crop",
            "効果 > パスファインダー > アウトライン\tLive Pathfinder Outline",
            "効果 > パスファインダー > 濃い混色\tLive Pathfinder Hard Mix",
            "効果 > パスファインダー > 薄い混色...\tLive Pathfinder Soft Mix",
            "効果 > パスファインダー > トラップ...\tLive Pathfinder Trap",
            "効果 > ラスタライズ...\tLive Rasterize",
            "効果 > ワープ > 円弧...\tLive Deform Arc",
            "効果 > ワープ > 下弦...\tLive Deform Arc Lower",
            "効果 > ワープ > 上弦...\tLive Deform Arc Upper",
            "効果 > ワープ > アーチ...\tLive Deform Arch",
            "効果 > ワープ > でこぼこ...\tLive Deform Bulge",
            "効果 > ワープ > 貝殻 (下向き)...\tLive Deform Shell Lower",
            "効果 > ワープ > 貝殻 (上向き)...\tLive Deform Shell Upper",
            "効果 > ワープ > 旗...\tLive Deform Flag",
            "効果 > ワープ > 波形...\tLive Deform Wave",
            "効果 > ワープ > 魚形...\tLive Deform Fish",
            "効果 > ワープ > 上昇...\tLive Deform Rise",
            "効果 > ワープ > 魚眼レンズ...\tLive Deform Fisheye",
            "効果 > ワープ > 膨張...\tLive Deform Inflate",
            "効果 > ワープ > 絞り込み...\tLive Deform Squeeze",
            "効果 > ワープ > 旋回...\tLive Deform Twist",
            "効果 > 形状に変換 > 長方形...\tLive Rectangle",
            "効果 > 形状に変換 > 角丸長方形...\tLive Rounded Rectangle",
            "効果 > 形状に変換 > 楕円形...\tLive Ellipse",
            "効果 > 効果ギャラリー...\tLive PSAdapter_plugin_GEfc",
            "効果 > ぼかし > ぼかし (ガウス)...\tLive Adobe PSL Gaussian Blur",
            "効果 > ぼかし > ぼかし (放射状)...\tLive PSAdapter_plugin_RdlB",
            "効果 > ぼかし > ぼかし (詳細)...\tLive PSAdapter_plugin_SmrB",
            "効果 > アーティスティック > こする...\tLive PSAdapter_plugin_SmdS",
            "効果 > アーティスティック > エッジのポスタリゼーション...\tLive PSAdapter_plugin_PstE",
            "効果 > アーティスティック > カットアウト...\tLive PSAdapter_plugin_Ct  ",
            "効果 > アーティスティック > スポンジ...\tLive PSAdapter_plugin_Spng",
            "効果 > アーティスティック > ドライブラシ...\tLive PSAdapter_plugin_DryB",
            "効果 > アーティスティック > ネオン光彩...\tLive PSAdapter_plugin_NGlw",
            "効果 > アーティスティック > パレットナイフ...\tLive PSAdapter_plugin_PltK",
            "効果 > アーティスティック > フレスコ...\tLive PSAdapter_plugin_Frsc",
            "効果 > アーティスティック > ラップ...\tLive PSAdapter_plugin_PlsW",
            "効果 > アーティスティック > 塗料...\tLive PSAdapter_plugin_PntD",
            "効果 > アーティスティック > 水彩画...\tLive PSAdapter_plugin_Wtrc",
            "効果 > アーティスティック > 粒状フィルム...\tLive PSAdapter_plugin_FlmG",
            "効果 > アーティスティック > 粗いパステル画...\tLive PSAdapter_plugin_RghP",
            "効果 > アーティスティック > 粗描き...\tLive PSAdapter_plugin_Undr",
            "効果 > アーティスティック > 色鉛筆...\tLive PSAdapter_plugin_ClrP",
            "効果 > スケッチ > ぎざぎざのエッジ...\tLive PSAdapter_plugin_TrnE",
            "効果 > スケッチ > ちりめんじわ...\tLive PSAdapter_plugin_Rtcl",
            "効果 > スケッチ > ウォーターペーパー...\tLive PSAdapter_plugin_WtrP",
            "効果 > スケッチ > クレヨンのコンテ画...\tLive PSAdapter_plugin_CntC",
            "効果 > スケッチ > クロム...\tLive PSAdapter_plugin_Chrm",
            "効果 > スケッチ > グラフィックペン...\tLive PSAdapter_plugin_GraP",
            "効果 > スケッチ > コピー...\tLive PSAdapter_plugin_Phtc",
            "効果 > スケッチ > スタンプ...\tLive PSAdapter_plugin_Stmp",
            "効果 > スケッチ > チョーク・木炭画...\tLive PSAdapter_plugin_ChlC",
            "効果 > スケッチ > ノート用紙...\tLive PSAdapter_plugin_NtPr",
            "効果 > スケッチ > ハーフトーンパターン...\tLive PSAdapter_plugin_HlfS",
            "効果 > スケッチ > プラスター...\tLive PSAdapter_plugin_Plst",
            "効果 > スケッチ > 木炭画...\tLive PSAdapter_plugin_Chrc",
            "効果 > スケッチ > 浅浮彫り...\tLive PSAdapter_plugin_BsRl",
            "効果 > テクスチャ > クラッキング...\tLive PSAdapter_plugin_Crql",
            "効果 > テクスチャ > ステンドグラス...\tLive PSAdapter_plugin_StnG",
            "効果 > テクスチャ > テクスチャライザー...\tLive PSAdapter_plugin_Txtz",
            "効果 > テクスチャ > パッチワーク...\tLive PSAdapter_plugin_Ptch",
            "効果 > テクスチャ > モザイクタイル...\tLive PSAdapter_plugin_MscT",
            "効果 > テクスチャ > 粒状...\tLive PSAdapter_plugin_Grn ",
            "効果 > ビデオ > NTSC カラー\tLive PSAdapter_plugin_NTSC",
            "効果 > ビデオ > インターレース解除...\tLive PSAdapter_plugin_Dntr",
            "効果 > ピクセレート > カラーハーフトーン...\tLive PSAdapter_plugin_ClrH",
            "効果 > ピクセレート > メゾティント...\tLive PSAdapter_plugin_Mztn",
            "効果 > ピクセレート > 水晶...\tLive PSAdapter_plugin_Crst",
            "効果 > ピクセレート > 点描...\tLive PSAdapter_plugin_Pntl",
            "効果 > ブラシストローク > はね...\tLive PSAdapter_plugin_Spt ",
            "効果 > ブラシストローク > インク画 (外形)...\tLive PSAdapter_plugin_InkO",
            "効果 > ブラシストローク > エッジの強調...\tLive PSAdapter_plugin_AccE",
            "効果 > ブラシストローク > ストローク (スプレー)...\tLive PSAdapter_plugin_SprS",
            "効果 > ブラシストローク > ストローク (斜め)...\tLive PSAdapter_plugin_AngS",
            "効果 > ブラシストローク > ストローク (暗)...\tLive PSAdapter_plugin_DrkS",
            "効果 > ブラシストローク > 墨絵...\tLive PSAdapter_plugin_Smie",
            "効果 > ブラシストローク > 網目...\tLive PSAdapter_plugin_Crsh",
            "効果 > 変形 > ガラス...\tLive PSAdapter_plugin_Gls ",
            "効果 > 変形 > 光彩拡散...\tLive PSAdapter_plugin_DfsG",
            "効果 > 変形 > 海の波紋...\tLive PSAdapter_plugin_OcnR",
            "効果 > 表現手法 > エッジの光彩...\tLive PSAdapter_plugin_GlwE",
            "表示 > CPU|GPU で表示\tView using GPU",
            "表示 > アウトライン\tpreview",
            "表示 > オーバープリントプレビュー\tink",
            "表示 > ピクセルプレビュー\traster",
            "表示 > 校正設定 > 作業用 CMYK : Japan Color 2001 Coated\tproof-document",
            "表示 > 校正設定 > 以前の Macintosh RGB (ガンマ 1.8)\tproof-mac-rgb",
            "表示 > 校正設定 > インターネット標準 RGB (sRGB)\tproof-win-rgb",
            "表示 > 校正設定 > モニター RGB\tproof-monitor-rgb",
            "表示 > 校正設定 > P 型 (1 型) 色覚\tproof-colorblindp",
            "表示 > 校正設定 > D 型 (2 型) 色覚\tproof-colorblindd",
            "表示 > 校正設定 > カスタム...\tproof-custom",
            "表示 > 色の校正\tproofColors",
            "表示 > ズームイン\tzoomin",
            "表示 > ズームアウト\tzoomout",
            "表示 > アートボードを全体表示\tfitin",
            "表示 > すべてのアートボードを全体表示\tfitall",
            "表示 > スライスを表示|隠す\tAISlice Feedback Menu",
            "表示 > スライスをロック\tAISlice Lock Menu",
            "表示 > 100% 表示\tactualsize",
            "表示 > 境界線を表示|隠す\tedge",
            "表示 > アートボードを表示|隠す\tartboard",
            "表示 > プリント分割を表示|隠す\tpagetiling",
            "表示 > バウンディングボックスを表示|隠す\tAI Bounding Box Toggle",
            "表示 > 透明グリッドを表示|隠す\tTransparencyGrid Menu Item",
            "表示 > テンプレートを表示|隠す\tshowtemplate",
            "表示 > グラデーションガイドを表示|隠す\tGradient Feedback",
            "表示 > ライブペイントの隙間を表示|隠す\tShow Gaps Planet X",
            "表示 > コーナーウィジェットを表示|隠す\tLive Corner Annotator",
            "表示 > スマートガイド\tSnapomatic on-off menu item",
            "表示 > 遠近グリッド > グリッドを表示|隠す\tShow Perspective Grid",
            "表示 > 遠近グリッド > 定規を表示|隠す\tShow Ruler",
            "表示 > 遠近グリッド > グリッドにスナップ\tSnap to Grid",
            "表示 > 遠近グリッド > グリッドをロック|ロック解除\tLock Perspective Grid",
            "表示 > 遠近グリッド > 測点をロック|ロック解除\tLock Station Point",
            "表示 > 遠近グリッド > グリッドを定義...\tDefine Perspective Grid",
            "表示 > 遠近グリッド > グリッドをプリセットとして保存...\tSave Perspective Grid as Preset",
            "表示 > 定規 > 定規を表示|隠す\truler",
            "表示 > 定規 > アートボード定規に変更\trulerCoordinateSystem",
            "表示 > 定規 > ビデオ定規を表示|隠す\tvideoruler",
            "表示 > テキストのスレッドを表示|隠す\ttextthreads",
            "表示 > ガイド > ガイドを表示|隠す\tshowguide",
            "表示 > ガイド > ガイドをロック|ロック解除\tlockguide",
            "表示 > ガイド > ガイドを作成\tmakeguide",
            "表示 > ガイド > ガイドを解除\treleaseguide",
            "表示 > ガイド > ガイドを消去\tclearguide",
            "表示 > グリッドを表示|隠す\tshowgrid",
            "表示 > グリッドにスナップ\tsnapgrid",
            "表示 > ポイントにスナップ\tsnappoint",
            "表示 > 新規表示...\tnewview",
            "表示 > 表示の編集...\teditview",
            "ウィンドウ > 新規ウィンドウ\tnewwindow",
            "ウィンドウ > アレンジ > 重ねて表示\tcascade",
            "ウィンドウ > アレンジ > 並べて表示\ttile",
            "ウィンドウ > アレンジ > ウィンドウを分離\tfloatInWindow",
            "ウィンドウ > アレンジ > すべてのウィンドウを分離\tfloatAllInWindows",
            "ウィンドウ > アレンジ > すべてのウィンドウを統合\tconsolidateAllWindows",
            "ウィンドウ > Exchange でエクステンションを検索...\tBrowse Add-Ons Menu",
            "ウィンドウ > ワークスペース > 「現在のワークスペース」をリセット\tAdobe Reset Workspace",
            "ウィンドウ > ワークスペース > 新規ワークスペース...\tAdobe New Workspace",
            "ウィンドウ > ワークスペース > ワークスペースの管理...\tAdobe Manage Workspace",
            "ウィンドウ > コントロール\tdrover control palette plugin",
            "ウィンドウ > ツールバー > 基本\tAdobe Basic Toolbar Menu",
            "ウィンドウ > ツールバー > 詳細\tAdobe Advanced Toolbar Menu",
            "ウィンドウ > ツールバー > 新しいツールバー...\tNew Tools Panel",
            "ウィンドウ > ツールバー > ツールバーを管理...\tManage Tools Panel",
            "ウィンドウ > 3D とマテリアル\tAdobe 3D Panel",
            "ウィンドウ > CC ライブラリ\tAdobe CSXS Extension com.adobe.DesignLibraries.angularCC ライブラリ",
            "ウィンドウ > CSS プロパティ\tCSS Menu Item",
            "ウィンドウ > SVG インタラクティビティ\tAdobe SVG Interactivity Palette",
            "ウィンドウ > アクション\tAdobe Action Palette",
            "ウィンドウ > アセットの書き出し\tAdobe SmartExport Panel Menu Item",
            "ウィンドウ > アピアランス\tStyle Palette",
            "ウィンドウ > アートボード\tAdobe Artboard Palette",
            "ウィンドウ > カラー\tAdobe Color Palette",
            "ウィンドウ > カラーガイド\tAdobe Harmony Palette",
            "ウィンドウ > グラデーション\tAdobe Gradient Palette",
            "ウィンドウ > グラフィックスタイル\tAdobe Style Palette",
            "ウィンドウ > コメント\tAdobe Commenting Palette",
            "ウィンドウ > シンボル\tAdobe Symbol Palette",
            "ウィンドウ > スウォッチ\tAdobe Swatches Menu Item",
            "ウィンドウ > ドキュメント情報\tDocInfo1",
            "ウィンドウ > ナビゲーター\tAdobeNavigator",
            "ウィンドウ > パスファインダー\tAdobe PathfinderUI",
            "ウィンドウ > パターンオプション\tAdobe Pattern Panel Toggle",
            "ウィンドウ > ブラシ\tAdobe BrushManager Menu Item",
            "ウィンドウ > プロパティ\tAdobe Property Palette",
            "ウィンドウ > リンク\tAdobe LinkPalette Menu Item",
            "ウィンドウ > レイヤー\tAdobeLayerPalette1",
            "ウィンドウ > 分割・統合プレビュー\tAdobe Flattening Preview",
            "ウィンドウ > 分版プレビュー\tAdobe Separation Preview Panel",
            "ウィンドウ > 変形\tAdobeTransformObjects1",
            "ウィンドウ > 変数\tAdobe Variables Palette Menu Item",
            "ウィンドウ > 属性\tinternal palettes posing as plug-in menus-attributes",
            "ウィンドウ > 情報\tinternal palettes posing as plug-in menus-info",
            "ウィンドウ > 整列\tAdobeAlignObjects2",
            "ウィンドウ > 書式 > OpenType\tinternal palettes posing as plug-in menus-opentype",
            "ウィンドウ > 書式 > タブ\tinternal palettes posing as plug-in menus-tab",
            "ウィンドウ > 書式 > 字形\talternate glyph palette plugin 2",
            "ウィンドウ > 書式 > 文字\tinternal palettes posing as plug-in menus-character",
            "ウィンドウ > 書式 > 文字スタイル\tCharacter Styles",
            "ウィンドウ > 書式 > 段落\tinternal palettes posing as plug-in menus-paragraph",
            "ウィンドウ > 書式 > 段落スタイル\tAdobe Paragraph Styles Palette",
            "ウィンドウ > 線\tAdobe Stroke Palette",
            "ウィンドウ > 自動選択\tAI Magic Wand",
            "ウィンドウ > 透明\tAdobe Transparency Palette Menu Item",
            "ウィンドウ > 画像トレース\tAdobe Vectorize Panel",
            "ウィンドウ > グラフィックスタイルライブラリ > その他のライブラリ...\tAdobe Art Style Plugin Other libraries menu item",
            "ウィンドウ > シンボルライブラリ > その他のライブラリ...\tAdobe Symbol Palette Plugin Other libraries menu item",
            "ウィンドウ > ブラシライブラリ > その他のライブラリ...\tAdobeBrushMgrUI Other libraries menu item",
            "ヘルプ > Illustrator ヘルプ...\thelpcontent",
            "ヘルプ > Illustrator について...\tabout",
            "Illustrator > Illustrator について...\tabout",
            "ヘルプ > システム情報...\tSystem Info\tAi 30 のショートカットファイルでは systemInfo（旧 ID が動くかは未確認）",
            "その他のパネル > 新規シンボル\tAdobe New Symbol Shortcut",
            "その他のパネル > カラーパネルを表示 (2)\tAdobe Color Palette Secondary",
            "その他のパネル > アクションバッチ\tAdobe Actions Batch",
            "その他のパネル > 新規塗りを追加\tAdobe New Fill Shortcut",
            "その他のパネル > 新規線を追加\tAdobe New Stroke Shortcut",
            "その他のパネル > 新規グラフィックスタイル\tAdobe New Style Shortcut",
            "その他のパネル > 新規レイヤー\tAdobeLayerPalette2",
            "その他のパネル > 新規レイヤー (オプション表示)\tAdobeLayerPalette3",
            "その他のパネル > リンクを更新\tAdobe Update Link Shortcut",
            "その他のパネル > 新規スウォッチ\tAdobe New Swatch Shortcut Menu",
            "その他のパネル > 新規スウォッチ\tFind Appearance menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > アピアランス属性\tFind Appearance Attributes menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > 描画モード\tFind Blending Mode menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > 塗りと線\tFind Fill & Stroke menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > カラー (塗り)\tFind Fill Color menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > 不透明度\tFind Opacity menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > カラー (線)\tFind Stroke Color menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > 線幅\tFind Stroke Weight menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > グラフィックスタイル\tFind Style menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > シェイプ\tFind Live Shape menu item\t廃止：Ai 25.9 まで（Ai 17 から）",
            "選択 > 共通 > シェイプとテキスト > シンボルインスタンス\tFind Symbol Instance menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "選択 > 共通 > シェイプとテキスト > 一連のリンクブロック\tFind Link Block Series menu item\t廃止：Ai 25.9 まで（Ai 16 から）",
            "効果 > 3D とマテリアル > 3D (クラシック) > 押し出し・ベベル...\tLive 3DExtrude\t廃止：Ai 25.9 まで（Ai 16 から）",
            "効果 > 3D とマテリアル > 3D (クラシック) > 回転体...\tLive 3DRevolve\t廃止：Ai 25.9 まで（Ai 16 から）",
            "効果 > 3D とマテリアル > 3D (クラシック) > 回転...\tLive 3DRotate\t廃止：Ai 25.9 まで（Ai 16 から）",
            "表示 > トリミング表示\tTrimView",
            "文字ツール\tAdobe Type Tool",
            "エリア内文字ツール\tAdobe Area Type Tool",
            "パス上文字ツール\tAdobe Path Type Tool",
            "文字 (縦)ツール\tAdobe Vertical Type Tool",
            "エリア内文字 (縦)ツール\tAdobe Vertical Area Type Tool",
            "パス上文字 (縦)ツール\tAdobe Vertical Path Type Tool",
            "文字タッチツール\tAdobe Touch Type Tool",
            "\tAdobe Right-To-Left Type Tool",
            "\tAdobe Right-To-Left Area Type Tool",
            "\tAdobe Right-To-Left Path Type Tool",
            "角丸長方形ツール\tAdobe Rounded Rectangle Tool",
            "選択ツール\tAdobe Select Tool",
            "ダイレクト選択ツール\tAdobe Direct Select Tool",
            "グループ選択ツール\tAdobe Direct Object Select Tool",
            "拡大・縮小ツール\tAdobe Scale Tool",
            "回転ツール\tAdobe Rotate Tool",
            "リフレクトツール\tAdobe Reflect Tool",
            "シアーツール\tAdobe Shear Tool",
            "棒グラフツール\tAdobe Column Graph Tool",
            "積み上げ棒グラフツール\tAdobe Stacked Column Graph Tool",
            "折れ線グラフツール\tAdobe Line Graph Tool",
            "円グラフツール\tAdobe Pie Graph Tool",
            "階層グラフツール\tAdobe Area Graph Tool",
            "散布図ツール\tAdobe Scatter Graph Tool",
            "横向き棒グラフツール\tAdobe Bar Graph Tool",
            "横向き積み上げ棒グラフツール\tAdobe Stacked Bar Graph Tool",
            "レーダーチャートツール\tAdobe Radar Graph Tool",
            "手のひらツール\tAdobe Scroll Tool",
            "鉛筆ツール\tAdobe Freehand Tool",
            "スムーズツール\tAdobe Freehand Smooth Tool",
            "パス消しゴムツール\tAdobe Freehand Erase Tool",
            "連結ツール\tAdobe Corner Join Tool",
            "ペンツール\tAdobe Pen Tool",
            "ブレンドツール\tAdobe Blend Tool",
            "はさみツール\tAdobe Scissors Tool",
            "ものさしツール\tAdobe Measure Tool",
            "プリント分割ツール\tAdobe Page Tool",
            "ズームツール\tAdobe Zoom Tool",
            "アンカーポイントの追加ツール\tAdobe Add Anchor Point Tool",
            "アンカーポイントの削除ツール\tAdobe Delete Anchor Point Tool",
            "アンカーポイントツール\tAdobe Anchor Point Tool",
            "グラデーションツール\tAdobe Gradient Vector Tool",
            "ブラシツール\tAdobe Brush Tool",
            "スポイトツール\tAdobe Eyedropper Tool",
            "リシェイプツール\tAdobe Reshape Tool",
            "線幅ツール\tAdobe Width Tool",
            "ナイフツール\tAdobe Knife Tool",
            "多角形ツール\tAdobe Shape Construction Regular Polygon Tool",
            "スターツール\tAdobe Shape Construction Star Tool",
            "スパイラルツール\tAdobe Shape Construction Spiral Tool",
            "メッシュツール\tAdobe Mesh Editing Tool",
            "自由変形ツール\tAdobe Free Transform Tool",
            "なげなわツール\tAdobe Direct Lasso Tool",
            "楕円形ツール\tAdobe Ellipse Shape Tool",
            "長方形ツール\tAdobe Rectangle Shape Tool",
            "自動選択ツール\tAdobe Magic Wand Tool",
            "直線ツール\tAdobe Line Tool",
            "円弧ツール\tAdobe Arc Tool",
            "長方形グリッドツール\tAdobe Rectangular Grid Tool",
            "同心円グリッドツール\tAdobe Polar Grid Tool",
            "フレアツール\tAdobe Flare Tool",
            "ワープツール\tAdobe Warp Tool",
            "うねりツール\tAdobe New Twirl Tool",
            "収縮ツール\tAdobe Pucker Tool",
            "膨張ツール\tAdobe Bloat Tool",
            "ひだツール\tAdobe Scallop Tool",
            "クラウンツール\tAdobe Cyrstallize Tool",
            "リンクルツール\tAdobe Wrinkle Tool",
            "スライスツール\tAdobe Slice Tool",
            "スライス選択ツール\tAdobe Slice Select Tool",
            "シンボルスプレーツール\tAdobe Symbol Sprayer Tool",
            "シンボルシフトツール\tAdobe Symbol Shifter Tool",
            "シンボルスクランチツール\tAdobe Symbol Scruncher Tool",
            "シンボルリサイズツール\tAdobe Symbol Sizer Tool",
            "シンボルスピンツール\tAdobe Symbol Spinner Tool",
            "シンボルステインツール\tAdobe Symbol Stainer Tool",
            "シンボルスクリーンツール\tAdobe Symbol Screener Tool",
            "シンボルスタイルツール\tAdobe Symbol Styler Tool",
            "ライブペイントツール\tAdobe Planar Paintbucket Tool",
            "ライブペイント選択ツール\tAdobe Planar Face Select Tool",
            "消しゴムツール\tAdobe Eraser Tool",
            "アートボードツール\tAdobe Crop Tool",
            "塗りブラシツール\tAdobe Blob Brush Tool",
            "シェイプ形成ツール\tAdobe Shape Builder Tool",
            "遠近グリッドツール\tPerspective Grid Tool",
            "遠近図形選択ツール\tPerspective Selection Tool",
            "\tAdobe Pattern Tile Tool",
            "\tAdobe Place Gun Tool",
            "曲線ツール\tAdobe Curvature Tool",
            "Shaper ツール\tAdobe Shaper Tool",
            "\tAdobe Symmetry Tool",
            "パペットワープツール\tAdobe Puppet Warp Tool",
            "\tAdobe Diffusion Coloring Tool",
            "\tAdobe Smart Edit Tool",
            "カラーテーマピッカーツール\tAdobe Color Theme Picker Tool",
            "回転ビューツール\tAdobe Rotate Canvas Tool",
            "ステンシルツール\tAdobe Stencil Tool",
            "クロスと重なりを選択ツール\tAdobe Intertwine Zone Marker Tool",
            "寸法ツール\tAdobe Dimension Tool",
            "パス上オブジェクトツール\tAdobe Constraints Tool",
            "Object > Objects on Path > Attach\tAttach Objects on Path\t廃止：Ai 29.1 まで（Ai 29 から）",
            "Object > Objects on Path > Options\tOptions Objects on Path",
            "Object > Objects on Path > Expand\tExpand Objects on Path",
            "Window > Type > Reflow Viewer\tReflowWindowMenu",
            "Window > Contextual Task Bar\t_GenericPluginMenuItem 25",
            "Object > Path > Smooth\tsmooth menu item",
            "Object > Mockup (Beta) > Make\tMake Vector Edge\t廃止：Ai 28.9 まで（Ai 28.6 から）",
            "Object > Mockup (Beta) > Edit\tEdit Vector Edge\t廃止：Ai 28.9 まで（Ai 28 から）",
            "Window > Mockup (Beta)\tAdobe Vector Edge Panel\t廃止：Ai 28.9 まで（Ai 28 から）",
            "Window > Text to Vector Graphic (Beta)\tGenerate\t廃止：Ai 28.5 まで（Ai 28 から）",
            "Edit > Preferences > Touch Workspace\tTouchPref",
            "Select > Update Selection\tSelection Hat 14",
            "Object > Ungroup All\tungroup all",
            "Window > Toolbars > Getting Started\tAdobe Quick Toolbar Menu",
            "Type > Composite Fonts\tAdobe internal composite font plugin",
            "Type > Kinsoku Shori Settings\tAdobe Kinsoku Settings",
            "Type > Mojikumi Settings\tAdobe MojiKumi Settings",
            "Type > Bullets and Numbering > Convert to text\tconvert list style to text",
            "Window > Retype (Beta)\tReTypeWindowMenu\t廃止：Ai 29.2 まで（Ai 27.6 から）",
            "Edit > Edit Colors > Generative Recolor (Beta)\tGenerative Recolor Art Dialog",
            "Object > Mockup (Beta) > Release\tRelease Vector Edge\t廃止：Ai 28.9 まで（Ai 28 から）",
            "一般\tpreference",
            "選択範囲・アンカー表示\tselectPref\tAi 30 のショートカットファイルでは selectionPref（旧 ID が動くかは未確認）",
            "テキスト\tkeyboardPref",
            "単位\tunitundoPref",
            "ガイド・グリッド\tguidegridPref",
            "スマートガイド\tsnapPref",
            "スライス\tslicePref",
            "ハイフネーション\thyphenPref",
            "プラグイン・仮想記憶ディスク\tpluginPref",
            "ユーザーインターフェイス\tUIPref\tAi 30 のショートカットファイルでは userInterfacePref（旧 ID が動くかは未確認）",
            "パフォーマンス\tGPUPerformancePref",
            "ファイル管理\tFilePref",
            "クリップボードの処理\tClipboardPref",
            "ブラックのアピアランス\tBlackPref",
            "デバイス\tDevicesPref",
            "General\tpreference",
            "Selection & Anchor Display\tselectPref\tAi 30 のショートカットファイルでは selectionPref（旧 ID が動くかは未確認）",
            "Type\tkeyboardPref",
            "Units\tunitundoPref",
            "Guides & Grid\tguidegridPref",
            "Smart Guides\tsnapPref",
            "Slices\tslicePref",
            "Hyphenation\thyphenPref",
            "Plug-ins & Scratch Disks\tpluginPref",
            "User Interface\tUIPref\tAi 30 のショートカットファイルでは userInterfacePref（旧 ID が動くかは未確認）",
            "Performance\tGPUPerformancePref",
            "File Handling\tFilePref",
            "Clipboard Handling\tClipboardPref",
            "Appearance of Black\tBlackPref",
            "Devices\tDevicesPref",
            "Window > Libraries\tAdobe CSXS Extension com.adobe.DesignLibraries.angularLibraries",
            "Window > History\tAdobe History Panel Menu Item",
            "Window > Version History\tAdobe Version History File Menu Item",
            "Window > Swatch Libraries > Other Library...\tAdobeSwatch_ Other libraries menu item",
            "Edit > Spelling > Auto Spell Check\tAuto Spell Check",
            "Debug Panel\tDebug Panel",
            "View > GPU Preview / Preview on CPU\tGPU Preview",
            "Object > Distribute > Horizontal Distribute Center\tHorizontal Distribute Center",
            "Object > Distribute > Horizontal Distribute Left\tHorizontal Distribute Left",
            "Object > Distribute > Horizontal Distribute Right\tHorizontal Distribute Right",
            "Object > Intertwine > Edit\tPartial Rearrange Edit",
            "Object > Intertwine > Make\tPartial Rearrange Make",
            "Object > Intertwine > Release\tPartial Rearrange Release",
            "Help > Support Community\tsupportCommunity",
            "Object > Distribute > Vertical Distribute Bottom\tVertical Distribute Bottom",
            "Object > Distribute > Vertical Distribute Center\tVertical Distribute Center",
            "Object > Distribute > Vertical Distribute Top\tVertical Distribute Top",
            "Help > Submit Bug/Feature Request...\twishform",
            "ウィンドウ > ライブラリ\tAdobe CSXS Extension com.adobe.DesignLibraries.angularLibraries",
            "ウィンドウ > ヒストリー\tAdobe History Panel Menu Item",
            "ウィンドウ > バージョン履歴\tAdobe Version History File Menu Item",
            "ウィンドウ > スウォッチライブラリ > その他のライブラリ...\tAdobeSwatch_ Other libraries menu item",
            "編集 > スペルチェック > 自動スペルチェック\tAuto Spell Check",
            "デバッグパネル\tDebug Panel",
            "表示 > GPU プレビュー / CPU でプレビュー\tGPU Preview",
            "オブジェクト > 分布 > 水平方向中央に分布\tHorizontal Distribute Center",
            "オブジェクト > 分布 > 水平方向左に分布\tHorizontal Distribute Left",
            "オブジェクト > 分布 > 水平方向右に分布\tHorizontal Distribute Right",
            "オブジェクト > クロスと重なり > 編集\tPartial Rearrange Edit",
            "オブジェクト > クロスと重なり > 作成\tPartial Rearrange Make",
            "オブジェクト > クロスと重なり > 解除\tPartial Rearrange Release",
            "ヘルプ > サポートコミュニティ\tsupportCommunity",
            "オブジェクト > 分布 > 垂直方向下に分布\tVertical Distribute Bottom",
            "オブジェクト > 分布 > 垂直方向中央に分布\tVertical Distribute Center",
            "オブジェクト > 分布 > 垂直方向上に分布\tVertical Distribute Top",
            "ヘルプ > バグを報告 / 機能改善をリクエスト...\twishform",
            "オブジェクト > 整列 > キーオブジェクトに整列\tAlign To Key Object",
            "オブジェクト > 整列 > アートボードに整列\tAlign To Artboard",
            "オブジェクト > 分布 > 垂直方向等間隔に分布\tVertical Distribute Space",
            "オブジェクト > 分布 > 水平方向等間隔に分布\tHorizontal Distribute Space",
            "オブジェクト > 生成 > 書き直し > テキストを生成...\tGenAIConsolidatedGenerateTextGenerate",
            "オブジェクト > 生成 > 書き直し > 翻訳...\tGenAIConsolidatedGenerateTextTranslate\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "オブジェクト > 生成 > 書き直し > 校正\tGenAIConsolidatedGenerateTextProofread\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "オブジェクト > 生成 > 書き直し > テキストを調整\tGenAIConsolidatedGenerateTextRephraseToFit\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "オブジェクト > 整列 > 選択範囲に揃える\tAlign To Selection",
            "オブジェクト > 生成 > ターンテーブル (20 クレジット)\tGenAIConsolidatedTurntable\t廃止：Ai 30.4 まで（Ai 30.3 から）",
            "ウィンドウ > コンセプトからベクター生成\tSketchToVectorUnified",
            "書式 > 書き直し...\tGenerateTextTypeMenu\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "書式 > Retype\tReTypeTypeMenu",
            "オブジェクト > 生成 > コンセプトからベクター生成\tGenAIConsolidatedTrace",
            "オブジェクト > 整列 > 水平・垂直方向中央に整列\tHorizontal && Vertical Align Center",
            "編集 > プロンプトで編集\tEditGeneratedObjectEditMenu",
            "書式 > 文字を切り換え | エリア内文字に切り換え | ポイント文字に切り換え\tpoint-area\t廃止：Ai 30.4 まで（Ai 30.4 から）",
            "オブジェクト > 背景を削除\tRemove Background Object Menu",
            "その他 > 前のドキュメントグループに移動\tnavigateToPreviousDocumentGroup",
            "その他 > 次のドキュメントグループに移動\tnavigateToNextDocumentGroup",
            "その他 > 前のドキュメントに移動\tnavigateToPreviousDocument",
            "その他 > 次のドキュメントに移動\tnavigateToNextDocument",
            "その他 > 新規ファイル (ダイアログなし)\tnew2",
            "その他 > 単位を切り換え\tswitchSelTool",
            "その他のオブジェクト > 編集モードを終了\texitFocus",
            "その他のオブジェクト > 選択オブジェクト編集モード\tenterFocus",
            "その他のオブジェクト > 平均・連結\tavgAndJoin",
            "その他のオブジェクト > パスファインダーの繰り返し\trepeatPathfinder",
            "その他のオブジェクト > 他を隠す\thide2",
            "その他のオブジェクト > 他をロック\tlock2",
            "その他のテキスト > 下線\t~textUnderline",
            "その他のテキスト > 斜体\t~textItalic",
            "その他のテキスト > 太字\t~textBold",
            "その他のテキスト > 上付き文字\t~superScript",
            "その他のテキスト > 下付き文字\t~subscript",
            "その他のテキスト > コンポーザーを切り換え\ttoggleLineComposer",
            "その他のテキスト > 自動ハイフネーションを切り換え\ttoggleAutoHyphen",
            "その他のテキスト > 両端揃え\tjustifyAll",
            "その他のテキスト > 均等配置 (最終行右 / 下揃え)\tjustifyRight",
            "その他のテキスト > 均等配置 (最終行中央揃え)\tjustifyCenter",
            "その他のテキスト > 均等配置 (最終行左 / 上揃え)\tjustify",
            "その他のテキスト > 右 / 下揃え\trightAlign",
            "その他のテキスト > 中央揃え\tcenterAlign",
            "その他のテキスト > 左 / 上揃え\tleftAlign",
            "その他のテキスト > フォントを強調表示 (2)\thighlightFont2",
            "その他のテキスト > フォントを強調表示\thighlightFont",
            "その他のテキスト > 文字の縦横比をリセット\tclearTypeScale",
            "その他のテキスト > 空白\tspacing",
            "その他のテキスト > トラッキングを解除\tclearTrack",
            "その他のテキスト > トラッキング\ttracking",
            "その他のテキスト > 文字間隔を狭く\t~kernCloser",
            "その他のテキスト > 文字間隔を広く\t~kernFurther",
            "その他のテキスト > フォントサイズを小さく (設定値×5)\tsizeStepDown",
            "その他のテキスト > フォントサイズを大きく (設定値×5)\tsizeStepUp",
            "その他のテキスト > フォントサイズを小さく (設定値参照)\tfaceSizeDown",
            "その他のテキスト > フォントサイズを大きく (設定値参照)\tfaceSizeUp",
            "書式 > 分割文字を挿入 > 強制改行\t~forcedLineBreak",
            "書式 > 空白文字を挿入 > 細いスペース\t~thinSpace",
            "書式 > 空白文字を挿入 > 極細スペース\t~hairSpace",
            "書式 > 空白文字を挿入 > EN スペース\t~enSpace",
            "書式 > 空白文字を挿入 > EM スペース\t~emSpace",
            "書式 > 特殊文字を挿入 > 引用符 > 右一重引用符\t~singleRightQuote",
            "書式 > 特殊文字を挿入 > 引用符 > 左一重引用符\t~singleLeftQuote",
            "書式 > 特殊文字を挿入 > 引用符 > 右二重引用符\t~doubleRightQuote",
            "書式 > 特殊文字を挿入 > 引用符 > 左二重引用符\t~doubleLeftQuote",
            "書式 > 特殊文字を挿入 > ハイフンおよびダッシュ > 任意ハイフン\t~discretionaryHyphen",
            "書式 > 特殊文字を挿入 > ハイフンおよびダッシュ > EN ダッシュ\t~enDash",
            "書式 > 特殊文字を挿入 > ハイフンおよびダッシュ > EM ダッシュ\t~emDash",
            "書式 > 特殊文字を挿入 > 記号 > 商標記号\t~trademarkSymbol",
            "書式 > 特殊文字を挿入 > 記号 > セクション記号\t~sectionSymbol",
            "書式 > 特殊文字を挿入 > 記号 > 登録商標記号\t~registeredTrademark",
            "書式 > 特殊文字を挿入 > 記号 > 段落記号\t~paragraphSymbol",
            "書式 > 特殊文字を挿入 > 記号 > 省略記号\t~ellipsis",
            "書式 > 特殊文字を挿入 > 記号 > 著作権記号\t~copyright",
            "書式 > 特殊文字を挿入 > 記号 > ビュレット\t~bullet",
            "ファイル > すべてを閉じる\tcloseAll",
            "オブジェクト > 生成 > 生成塗りつぶし (シェイプ)...\tGenAIConsolidatedShapeFill",
            "オブジェクト > 生成 > 生成拡張 > 作成\tGen Expand Object Make\t廃止：Ai 29.8 まで（Ai 29.6 から）",
            "オブジェクト > 生成 > 生成拡張 > 結合\tGen Expand Object Combine\t廃止：Ai 29.8 まで（Ai 29.6 から）",
            "オブジェクト > 生成 > 裁ち落としを印刷\tGenAIConsolidatedBleed",
            "オブジェクト > 生成 > 生成再配色\tGenAIConsolidatedRecolor",
            "オブジェクト > 生成 > パターンを生成\tGenAIConsolidatedPatterns",
            "オブジェクト > 生成 > 生成履歴\tGenAIConsolidatedVariations",
            "オブジェクト > アートボード > 方向切り替え\tSwitch Orientation",
            "オブジェクト > 生成 > ベクターを生成...\tGenAIConsolidatedGenerateVectors",
            "ヘルプ > 新機能...\twhatsNewContent",
            "ファイル > レビュー用に共有...\tShare For Review\t廃止：Ai 27.3 まで（Ai 27 から）",
            "ファイル > 編集に招待...\tInvite People",
            "Object > Align > Align To Key Object\tAlign To Key Object",
            "Object > Align > Align To Artboard\tAlign To Artboard",
            "Object > Distribute > Vertical Distribute Space\tVertical Distribute Space",
            "Object > Distribute > Horizontal Distribute Space\tHorizontal Distribute Space",
            "Object > Generative > Rewrite > Generate Text...\tGenAIConsolidatedGenerateTextGenerate",
            "Object > Generative > Rewrite > Translate...\tGenAIConsolidatedGenerateTextTranslate\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "Object > Generative > Rewrite > Proofread\tGenAIConsolidatedGenerateTextProofread\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "Object > Generative > Rewrite > Fit text\tGenAIConsolidatedGenerateTextRephraseToFit\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "Object > Align > Align To Selection\tAlign To Selection",
            "Object > Generative > Turntable (20 credits)\tGenAIConsolidatedTurntable\t廃止：Ai 30.4 まで（Ai 30.3 から）",
            "Window > Concept to Vector\tSketchToVectorUnified",
            "Type > Rewrite...\tGenerateTextTypeMenu\t廃止：Ai 30.5 まで（Ai 30.5 から）",
            "Type > Retype\tReTypeTypeMenu",
            "Object > Generative > Concept to Vector\tGenAIConsolidatedTrace",
            "Object > Align > Horizontal & Vertical Align Center\tHorizontal && Vertical Align Center",
            "Edit > Prompt to edit\tEditGeneratedObjectEditMenu",
            "Type > Text Type Conversion | Convert To Area Type | Convert To Point Type\tpoint-area\t廃止：Ai 30.4 まで（Ai 30.4 から）",
            "Object > Remove Background\tRemove Background Object Menu",
            "Other Misc > Navigate to Previous Document Group\tnavigateToPreviousDocumentGroup",
            "Other Misc > Navigate to Next Document Group\tnavigateToNextDocumentGroup",
            "Other Misc > Navigate to Previous Document\tnavigateToPreviousDocument",
            "Other Misc > Navigate to Next Document\tnavigateToNextDocument",
            "Other Misc > New File (No Dialog)\tnew2",
            "Other Misc > Switch Units\tswitchSelTool",
            "Other Object > Exit Isolation Mode\texitFocus",
            "Other Object > Isolate Selected Object\tenterFocus",
            "Other Object > Average & Join\tavgAndJoin",
            "Other Object > Repeat Pathfinder\trepeatPathfinder",
            "Other Object > Hide Others\thide2",
            "Other Object > Lock Others\tlock2",
            "Other Text > Underline\t~textUnderline",
            "Other Text > Italic\t~textItalic",
            "Other Text > Bold\t~textBold",
            "Other Text > Superscript\t~superScript",
            "Other Text > Subscript\t~subscript",
            "Other Text > Toggle Line Composer\ttoggleLineComposer",
            "Other Text > Toggle Auto Hyphenation\ttoggleAutoHyphen",
            "Other Text > Justify All Lines\tjustifyAll",
            "Other Text > Justify Text Right\tjustifyRight",
            "Other Text > Justify Text Center\tjustifyCenter",
            "Other Text > Justify Text Left\tjustify",
            "Other Text > Right Align Text\trightAlign",
            "Other Text > Center Text\tcenterAlign",
            "Other Text > Left Align Text\tleftAlign",
            "Other Text > Highlight Font (Secondary)\thighlightFont2",
            "Other Text > Highlight Font\thighlightFont",
            "Other Text > Uniform Type\tclearTypeScale",
            "Other Text > Spacing\tspacing",
            "Other Text > Clear Tracking\tclearTrack",
            "Other Text > Tracking\ttracking",
            "Other Text > Kern Tighter\t~kernCloser",
            "Other Text > Kern Looser\t~kernFurther",
            "Other Text > Font Size Step Down\tsizeStepDown",
            "Other Text > Font Size Step Up\tsizeStepUp",
            "Other Text > Point Size Down\tfaceSizeDown",
            "Other Text > Point Size Up\tfaceSizeUp",
            "Type > Insert Break Character > Forced Line Break\t~forcedLineBreak",
            "Type > Insert WhiteSpace Character > Thin Space\t~thinSpace",
            "Type > Insert WhiteSpace Character > Hair Space\t~hairSpace",
            "Type > Insert WhiteSpace Character > En Space\t~enSpace",
            "Type > Insert WhiteSpace Character > Em Space\t~emSpace",
            "Type > Insert Special Character > Quotation Marks > Single Right Quotation Marks\t~singleRightQuote",
            "Type > Insert Special Character > Quotation Marks > Single Left Quotation Marks\t~singleLeftQuote",
            "Type > Insert Special Character > Quotation Marks > Double Right Quotation Marks\t~doubleRightQuote",
            "Type > Insert Special Character > Quotation Marks > Double Left Quotation Marks\t~doubleLeftQuote",
            "Type > Insert Special Character > Hyphens And Dashes > Discretionary Hyphen\t~discretionaryHyphen",
            "Type > Insert Special Character > Hyphens And Dashes > En Dash\t~enDash",
            "Type > Insert Special Character > Hyphens And Dashes > Em Dash\t~emDash",
            "Type > Insert Special Character > Symbols > Trademark Symbol\t~trademarkSymbol",
            "Type > Insert Special Character > Symbols > Section Symbol\t~sectionSymbol",
            "Type > Insert Special Character > Symbols > Registered Trademark Symbol\t~registeredTrademark",
            "Type > Insert Special Character > Symbols > Paragraph Symbol\t~paragraphSymbol",
            "Type > Insert Special Character > Symbols > Ellipsis\t~ellipsis",
            "Type > Insert Special Character > Symbols > Copyright Symbol\t~copyright",
            "Type > Insert Special Character > Symbols > Bullet\t~bullet",
            "File > Close All\tcloseAll",
            "Object > Generative > Gen Shape Fill...\tGenAIConsolidatedShapeFill",
            "Object > Generative > Generative Expand > Make\tGen Expand Object Make\t廃止：Ai 29.8 まで（Ai 29.6 から）",
            "Object > Generative > Generative Expand > Combine\tGen Expand Object Combine\t廃止：Ai 29.8 まで（Ai 29.6 から）",
            "Object > Generative > Print Bleed\tGenAIConsolidatedBleed",
            "Object > Generative > Generative Recolor\tGenAIConsolidatedRecolor",
            "Object > Generative > Generate Patterns\tGenAIConsolidatedPatterns",
            "Object > Generative > Generation History\tGenAIConsolidatedVariations",
            "Object > Artboards > Switch Orientation\tSwitch Orientation",
            "Object > Generative > Generate Vectors...\tGenAIConsolidatedGenerateVectors",
            "Help > What's New...\twhatsNewContent",
            "File > Share for Review...\tShare For Review\t廃止：Ai 27.3 まで（Ai 27 から）",
            "File > Invite to Edit…\tInvite People",
            "編集 > SWF プリセット...\tSWFPresets\t廃止：Ai 25.9 まで（Ai 16 から）",
            "編集 > 環境設定 > ファイル管理・クリップボード...\tFileClipboardPref\t廃止：Ai 24.9 まで（Ai 16 から）。FilePref / ClipboardPref に分割",
            "ウィンドウ > ツールバー > 初期設定\tDefault ToolBar\t廃止：Ai 22.9 まで（Ai 17 から）",
            "ウィンドウ > Adobe Color テーマ\tAdobe Illustrator Kuler Panel\t廃止：Ai 25.9 まで（Ai 16 から）",
            "ウィンドウ > ライブラリ\tAdobe CSXS Extension com.adobe.DesignLibraries.angularライブラリ\t廃止：Ai 22.9 まで（Ai 18.1 から）",
            "ウィンドウ > ラーニング\tAdobe Learn Panel Menu Item\t廃止：Ai 25.9 まで（Ai 17 から）",
            "Object > Pattern > Text to Pattern (Beta)\tText To Pattern\t廃止：Ai 28.5 まで（Ai 28 から）。Beta",
            "File > Generate Vectors (Beta)\tGenerate Modal File Menu \t廃止：Ai 29.8 まで（Beta、GenAIConsolidatedGenerateVectors に移行）。Beta。GenAIConsolidatedGenerateVectors に移行",
            "Object > Gen Shape Fill (Beta)\tShape Fill Object Menu\t廃止：Ai 29.8 まで（Ai 29.5 から）。Beta",
            "Window > Generate Patterns (Beta)\tAdobe Generative Patterns Panel\t廃止：Ai 29.8 まで（Ai 29.5 から）。Beta",
            "Illustrator > Preferences > File Handling & Clipboard\tFileClipboardPref\t廃止：Ai 24.9 まで（Ai 16 から）。FilePref / ClipboardPref に分割",
            "Window > History\tAdobe HistoryPanel Menu Item\t廃止：Ai 26.9 まで（Ai 26.4 から）。Adobe History Panel Menu Item に改名",
            "一般 > キー入力\tcursorKeyLength\tReal\t0.2835",
            "General > Keyboard Increment\tcursorKeyLength\tReal\t0.2835",
            "一般 > 角度の制限\tconstrain/angle\tReal\t0.0\tconstrain/sin・constrain/cos も合わせて書かないと拘束に反映されない",
            "General > Constrain Angle\tconstrain/angle\tReal\t0.0\tconstrain/sin・constrain/cos も合わせて書かないと拘束に反映されない",
            "\tconstrain/cos\tReal\t1.0\tconstrain/angleの値と矛盾が出ないようセットする必要がある。GUIでセットする限り自動で変更される",
            "\tconstrain/sin\tReal\t0.0\tconstrain/angleの値と矛盾が出ないようセットする必要がある。GUIでセットする限り自動で変更される",
            "一般 > 角丸の半径\tovalRadius\tReal\t2.8346456693",
            "General > Corner Radius\tovalRadius\tReal\t2.8346456693",
            "一般 > ペンツールでパス上にアンカーポイントを自動で追加 / 削除しない\tpen/disableAutoAddDelete\tBoolean\t0",
            "General > Disable Auto Add/Delete\tpen/disableAutoAddDelete\tBoolean\t0",
            "一般 > 十字カーソルを使用\tusePreciseCursors\tBoolean\t0",
            "General > Use Precise Cursors\tusePreciseCursors\tBoolean\t0",
            "一般 > ツールヒントを表示\tshowToolTips\tBoolean\t0",
            "General > Show Tool Tips\tshowToolTips\tBoolean\t0",
            "一般 > 詳細なツールヒントを表示\tshowRichToolTips\tBoolean\t0",
            "General > Show Rich Tool Tips\tshowRichToolTips\tBoolean\t0",
            "一般 > すべてのドキュメントで定規を表示 / 非表示\tuseGlobalRulers\tBoolean\t0",
            "General > Show/Hide Rulers\tuseGlobalRulers\tBoolean\t0",
            "一般 > アートワークのアンチエイリアス\tantialias/graphic\tBoolean\t1\tRead Only?",
            "General > Anti-aliased Artwork\tantialias/graphic\tBoolean\t1\tRead Only?",
            "\tantialias/text\tBoolean\t1",
            "\tantialias/image\tBoolean\t1",
            "一般 > 同じ濃度を選択\tselectSameTintPercentage\tBoolean\t0",
            "General > Select Same Tint %\tselectSameTintPercentage\tBoolean\t0",
            "一般 > ドキュメントを開いていないときにホーム画面を表示\tHello/ShowHomeScreenWS\tBoolean\t0\tドキュメントを開いていないときに「ホーム画面」を表示（PresetManager.jsx）",
            "General > Show The Home Screen When No Documents Are Open\tHello/ShowHomeScreenWS\tBoolean\t0\tドキュメントを開いていないときに「ホーム画面」を表示（PresetManager.jsx）",
            "一般 > 以前の「新規ドキュメント」インターフェイスを使用\tHello/NewDoc\tBoolean\t1\t{true: 以前のUIを使用, false: 新しいUIを使用}\\n以前の「新規ドキュメント」インターフェイスを使用（PresetManager.jsx）",
            "General > Use legacy “File New” interface\tHello/NewDoc\tBoolean\t1\t{true: 以前のUIを使用, false: 新しいUIを使用}\\n以前の「新規ドキュメント」インターフェイスを使用（PresetManager.jsx）",
            "一般 > 100% ズームで印刷サイズを表示\tEnableActualViewPreview\tBoolean\t1",
            "General > Display Print Size at 100% Zoom\tEnableActualViewPreview\tBoolean\t1",
            "一般 > 以前のバージョンファイルを開くときに [更新済み] をファイル名に追加\tfileFormatGetFile/ConvertedInFilename\tBoolean\t1",
            "General > Append [Converted] Upon Opening Legacy Files\tfileFormatGetFile/ConvertedInFilename\tBoolean\t1",
            "一般 > 起動時にシステム互換性の問題を表示\taiShowSystemCompatibilityIssuesAtStartup\tBoolean\t1",
            "General > Show system compatibility issues at startup\taiShowSystemCompatibilityIssuesAtStartup\tBoolean\t1",
            "一般 > ダブルクリックして編集モード\tdoubleClickToIsolate\tBoolean\t1",
            "General > Double Click To Isolate\tdoubleClickToIsolate\tBoolean\t1",
            "一般 > 日本式トンボを使用\tcropMarkStyle\tBoolean\t1\t{0: 欧文式, 1: 日本式} なので実質的にBoolean",
            "General > Use Japanese Crop Marks\tcropMarkStyle\tBoolean\t1\t{0: 欧文式, 1: 日本式} なので実質的にBoolean",
            "一般 > パターンを変形\ttransformPatterns\tBoolean\t1",
            "General > Transform Pattern Tiles\ttransformPatterns\tBoolean\t1",
            "一般 > 角を拡大・縮小\tpolicyForPreservingCorners\tBoolean\t2\t{true: ON, false: OFF} なので実質的にBoolean",
            "General > Scale Corners\tpolicyForPreservingCorners\tBoolean\t2\t{true: ON, false: OFF} なので実質的にBoolean",
            "一般 > 線幅と効果も拡大・縮小\tscaleLineWeight\tBoolean\t0",
            "General > Scale Strokes & Effects\tscaleLineWeight\tBoolean\t0",
            "一般 > コンテンツに応じた初期値を適用\tEnableContentAwareDefaults\tBoolean\t1",
            "General > Enable Content Aware Defaults\tEnableContentAwareDefaults\tBoolean\t1",
            "一般 > PDF を読み込むときに元のサイズを維持\tplugin/PDFImport/HonourPageScale\tBoolean\t0",
            "General > Honor Scale on PDF Import\tplugin/PDFImport/HonourPageScale\tBoolean\t0",
            "一般 > マウスホイールでズーム\tzoomWithMouseWheel\tBoolean\t0",
            "General > Zoom with Mouse Wheel\tzoomWithMouseWheel\tBoolean\t0",
            "一般 > トラックパッドジェスチャーでビューを回転\taiRotateWithTrackpad\tBoolean\t1",
            "General > Trackpad Gesture to Rotate View\taiRotateWithTrackpad\tBoolean\t1",
            "一般 > プレビュー境界を使用\tincludeStrokeInBounds\tBoolean\t0",
            "General > Use Preview Bounds\tincludeStrokeInBounds\tBoolean\t0",
            "選択範囲・アンカー表示 > 選択範囲 > 許容値\tselectionTolerance\tInteger\t6",
            "Selection & Anchor Display > Selection > Tolerance\tselectionTolerance\tInteger\t6",
            "選択範囲・アンカー表示 > 選択範囲 > カンバス上のオブジェクトを選択してロック解除\tshowLockIcon\tBoolean\t0\t未登録のキーは true を返すことがあるので、ON/OFF 両方で確認すること\\nカンバス上のオブジェクトとアートボードを選択してロック解除（PresetManager.jsx）",
            "Selection & Anchor Display > Selection > Select and Unlock objects on canvas\tshowLockIcon\tBoolean\t0\t未登録のキーは true を返すことがあるので、ON/OFF 両方で確認すること\\nカンバス上のオブジェクトとアートボードを選択してロック解除（PresetManager.jsx）",
            "選択範囲・アンカー表示 > 選択範囲 > 選択ツールおよびシェイプツールでアンカーポイントを表示\thideAnchorPointsInTools\tBoolean\t1\t{true: OFF, false: ON}",
            "Selection & Anchor Display > Selection > Show Anchor Points in Selection Tool and Shape Tools\thideAnchorPointsInTools\tBoolean\t1\t{true: OFF, false: ON}",
            "選択範囲・アンカー表示 > 選択範囲 > セグメントをドラッグしてリシェイプするときにハンドル方向を固定\tconstrainPathDragging\tBoolean\t0",
            "Selection & Anchor Display > Selection > Constrain Path Dragging on Segment Reshape\tconstrainPathDragging\tBoolean\t0",
            "選択範囲・アンカー表示 > 選択範囲 > ロックまたは非表示オブジェクトをアートボードと一緒に移動\tmoveLockedAndHiddenArt\tBoolean\t1",
            "Selection & Anchor Display > Selection > Move Locked and Hidden Artwork with Artboard\tmoveLockedAndHiddenArt\tBoolean\t1",
            "選択範囲・アンカー表示 > 選択範囲 > ポイントにスナップ «真偽値»\tsnapToPoint\tBoolean\t1",
            "Selection & Anchor Display > Selection > Snap to Point «Checked»\tsnapToPoint\tBoolean\t1",
            "選択範囲・アンカー表示 > 選択範囲 > ポイントにスナップ «数値»\tsnappingTolerance\tInteger\t1",
            "Selection & Anchor Display > Selection > Snap to Point «Integer»\tsnappingTolerance\tInteger\t1",
            "選択範囲・アンカー表示 > 選択範囲 > オブジェクトの選択範田をパスに制限\thitShapeOnPreview\tInteger\t1\t値：\\n0 ON\\n1 OFF\\n0＝ON・1＝OFF の反転（オブジェクトの選択範囲をパスに制限）（PresetManager.jsx）",
            "Selection & Anchor Display > Selection > Object Selection by Path Only\thitShapeOnPreview\tInteger\t1\t値：\\n0 ON\\n1 OFF\\n0＝ON・1＝OFF の反転（オブジェクトの選択範囲をパスに制限）（PresetManager.jsx）",
            "選択範囲・アンカー表示 > 選択範囲 > Cmd + クリックで背面のオブジェクトを選択\tselectBehind\tBoolean\t1",
            "Selection & Anchor Display > Selection > Command Click to Select Objects Behind\tselectBehind\tBoolean\t1",
            "選択範囲・アンカー表示 > 選択範囲 > 選択範囲へズーム\tzoomToSelection\tBoolean\t0",
            "Selection & Anchor Display > Selection > Zoom to Selection\tzoomToSelection\tBoolean\t0",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > サイズ\tanchorSizePref\tInteger\t7\t値：5 / 7 / 9 / 11（［選択範囲・アンカー表示］のスライダー4段階、既定 5）。アンカーポイント・ハンドル・バウンディングボックスの表示サイズ（PresetManager.jsx）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Size\tanchorSizePref\tInteger\t7\t値：5 / 7 / 9 / 11（［選択範囲・アンカー表示］のスライダー4段階、既定 5）。アンカーポイント・ハンドル・バウンディングボックスの表示サイズ（PresetManager.jsx）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > ハンドルスタイル\thandleTypePref\tInteger\t1\t値：\\n0 ● 塗り (Fill)\\n1 ◯ 線 (Stroke)",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Handle Style\thandleTypePref\tInteger\t1\t値：\\n0 ● 塗り (Fill)\\n1 ◯ 線 (Stroke)",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > カーソルを合わせたときにアンカーを強調表示\thighlightAnchorOnMouseOver\tBoolean\t1",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Highlight anchors on mouse over\thighlightAnchorOnMouseOver\tBoolean\t1",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > 複数アンカーを選択時にハンドルを表示\tshowDirectionHandles\tBoolean\t0",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Show handles when multiple anchors are selected\tshowDirectionHandles\tBoolean\t0",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > 次の角度より大きいときにコーナーウィジェットを隠す «真偽値»\tliveCorners/hideCornerWidgetBasedOnAngle\tBoolean\t1",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Hide Corner Widget for angles greater than «Checked»\tliveCorners/hideCornerWidgetBasedOnAngle\tBoolean\t1",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > 次の角度より大きいときにコーナーウィジェットを隠す «角度»\tliveCorners/cornerAngleLimit\tReal\t177.0",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Hide Corner Widget for angles greater than «Angle»\tliveCorners/cornerAngleLimit\tReal\t177.0",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > ラバーバンドを有効にする対象 > ペンツール\ttoolPen/rubberBandEnable\tBoolean\t1",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Enable Rubber Band for > Pen Tool\ttoolPen/rubberBandEnable\tBoolean\t1",
            "選択範囲・アンカー表示 > アンカーポイント、ハンドル、およびバウンディングボックスの表示 > ラバーバンドを有効にする対象 > 曲線ツール\tCurvatureTool/rubberBandEnable\tBoolean\t1",
            "Selection & Anchor Display > Anchor Points, Handle, and Bounding Box Display > Enable Rubber Band for > Curvature Tool\tCurvatureTool/rubberBandEnable\tBoolean\t1",
            "テキスト > サイズ / 行送り\ttext/sizeIncrement\tReal\t1.0",
            "Type > Size/Leading\ttext/sizeIncrement\tReal\t1.0",
            "テキスト > トラッキング\ttext/kernIncrement\tInteger\t20",
            "Type > Tracking\ttext/kernIncrement\tInteger\t20",
            "テキスト > ベースラインシフト\ttext/riseIncrement\tReal\t0.1",
            "Type > Baseline Shift\ttext/riseIncrement\tReal\t0.1",
            "テキスト > 言語オプション > 東アジア言語のオプションを表示\tshowAsianTextOptions\tBoolean\t1",
            "Type > Language Options > Show East Asian Options\tshowAsianTextOptions\tBoolean\t1",
            "テキスト > 言語オプション > インド言語のオプションを表示\tAI WorldReadiness Dict Key\tBoolean\t0\tテキスト > 言語オプション > 東アジア言語のオプションを表示 と択一",
            "Type > Language Options > Show Indic Options\tAI WorldReadiness Dict Key\tBoolean\t0\tテキスト > 言語オプション > 東アジア言語のオプションを表示 と択一",
            "テキスト > テキストオブジェクトの選択範囲をパスに制限\thitTypeShapeOnPreview\tInteger\t1\t値：\\n0 ON\\n1 OFF\\n0＝ON・1＝OFF の反転（テキストオブジェクトの選択範囲をパスに制限）（PresetManager.jsx）",
            "Type > Type Object Selection by Path Only\thitTypeShapeOnPreview\tInteger\t1\t値：\\n0 ON\\n1 OFF\\n0＝ON・1＝OFF の反転（テキストオブジェクトの選択範囲をパスに制限）（PresetManager.jsx）",
            "テキスト > フォント名を英語表記\ttext/useEnglishFontNames\tBoolean\t0",
            "Type > Show Font Names in English\ttext/useEnglishFontNames\tBoolean\t0",
            "テキスト > 新規エリア内文字の自動サイズ調整\ttext/autoSizing\tBoolean\t1",
            "Type > Auto Size New Area Type\ttext/autoSizing\tBoolean\t1",
            "テキスト > フォントメニュー内のフォントプレビューを表示\ttext/fontMenu/showInFace\tBoolean\t1",
            "Type > Enable in-menu font previews\ttext/fontMenu/showInFace\tBoolean\t1",
            "\ttext/fontMenu/faceSizeMultiplier\tReal\t1.0\tフォントプレビューの大きさ",
            "テキスト > 最近使用したフォントの表示数\ttext/recentFontMenu/showNEntries\tInteger\t15\t最近使用したフォントの表示数。0 で一覧を表示しない（ダイアログでの範囲は 1〜30）（PresetManager.jsx）",
            "Type > Number of Recent Fonts\ttext/recentFontMenu/showNEntries\tInteger\t15\t最近使用したフォントの表示数。0 で一覧を表示しない（ダイアログでの範囲は 1〜30）（PresetManager.jsx）",
            "テキスト > 「さらに検索」で日本語フォントを表示\ttext/fontMenu/japaneseFontPreview\tBoolean\t1",
            "Type > Enable Japanese Font Preview in ‘Find More’\ttext/fontMenu/japaneseFontPreview\tBoolean\t1",
            "テキスト > 見つからない字形の保護を有効にする\ttext/doFontLocking\tBoolean\t0",
            "Type > Enable Missing Glyph Protection\ttext/doFontLocking\tBoolean\t0",
            "テキスト > 代替フォントを強調表示\ttext/highlightMissingFonts\tBoolean\t1",
            "Type > Highlight Substituted Fonts\ttext/highlightMissingFonts\tBoolean\t1",
            "テキスト > 新規テキストオブジェクトにサンプルテキストを割り付け «グローバル»\ttext/fillWithDefaultText\tBoolean\t1",
            "Type > Fill New Type Objects With Placeholder Text\ttext/fillWithDefaultText\tBoolean\t1",
            "テキスト > 新規テキストオブジェクトにサンプルテキストを割り付け «日本語»\ttext/fillWithDefaultTextJP\tBoolean\t0",
            "Type > Fill New Type Objects With Placeholder Text «Japanese»\ttext/fillWithDefaultTextJP\tBoolean\t0",
            "テキスト > 選択された文字の異体字を表示\ttext/enableAlternateGlyph\tBoolean\t0",
            "Type > Show Character Alternates\ttext/enableAlternateGlyph\tBoolean\t0",
            "テキスト > 入力中に箇条書きリストのレベルを自動変更\ttext/enableListAutoDetection\tBoolean\t1",
            "Type > Automatic Bulleted and Numbered lists while typing\ttext/enableListAutoDetection\tBoolean\t1",
            "\ttext/fontMenu/menuNameLimit\tInteger\t35",
            "\ttext/fontMenu/showSubMenusInFace\tBoolean\t0",
            "\ttext/fontMenu/exploreModeOnbrdngShownCount\tInteger\t3",
            "\ttext/fontMenu/needsExploreModeOnboarding\tBoolean\t0",
            "\ttext/doNonLatinInlineInput\tInteger\t1",
            "\ttext/activateOnbrdngShownCount\tInteger\t2",
            "\ttext/enablePreciseBBox\tBoolean\t0",
            "\ttext/GreekingDisabled\tInteger\t1",
            "\ttext/greekingThreshold\tReal\t6",
            "\ttext/fontflyout/SampleTextOpt\tInteger\t0",
            "\ttext/faceSize\tReal\t12",
            "\ttext/groupTypeMenuByLanguage\tBoolean\t1\tfalseにするとPostscriptNameで並ぶ",
            "\ttext/debugFontMenu/writeDebugFile\tInteger\t0",
            "\ttext/debugFontMenu/lieAboutNativeScript\tInteger\t0",
            "単位 > 一般\trulerType\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n3 パイカ Picas p\\n4 センチメートル Centimeters cm\\n5 歯 Ha H\\n6 ピクセル Pixels px\\n7 フィートとインチ Feet & Inches ft in\\n8 メートル Meters m\\n9 ヤード Yards yd\\n10 フィート Feet ft\\n単位コード 5 は「歯（H）」",
            "Units > General\trulerType\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n3 パイカ Picas p\\n4 センチメートル Centimeters cm\\n5 歯 Ha H\\n6 ピクセル Pixels px\\n7 フィートとインチ Feet & Inches ft in\\n8 メートル Meters m\\n9 ヤード Yards yd\\n10 フィート Feet ft\\n単位コード 5 は「歯（H）」",
            "単位 > 線\tstrokeUnits\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n3 パイカ Picas p\\n4 センチメートル Centimeters cm\\n5 歯 Ha H\\n6 ピクセル Pixels px\\n単位コード 5 は「歯（H）」",
            "Units > Stroke\tstrokeUnits\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n3 パイカ Picas p\\n4 センチメートル Centimeters cm\\n5 歯 Ha H\\n6 ピクセル Pixels px\\n単位コード 5 は「歯（H）」",
            "単位 > 文字\ttext/units\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n5 級 Qs Q\\n6 ピクセル Pixels px\\n単位コード 5 は「級（Q）」（文字サイズだけ Q、ほかのキーは H）",
            "Units > Type\ttext/units\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n5 級 Qs Q\\n6 ピクセル Pixels px\\n単位コード 5 は「級（Q）」（文字サイズだけ Q、ほかのキーは H）",
            "単位 > 東アジア言語のオプション\ttext/asianunits\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n5 級 Qs Q\\n6 ピクセル Pixels px\\n単位コード 5 は「歯（H）」",
            "Units > East Asian Type\ttext/asianunits\tInteger\t2\t値：\\n0インチ Inches in\\n1 ミリメートル Millimeters mm\\n2 ポイント Points pt\\n5 級 Qs Q\\n6 ピクセル Pixels px\\n単位コード 5 は「歯（H）」",
            "\tnumbersArePoints\tInteger\t0\t単位 > 数値のみの入力はポイントを単位とする ?\\nUnits > Numbers Without Units Are Points ?\\n常にOFFになっている",
            "単位 > オブジェクトの識別方法\tartNamesAreXMLIDs\tInteger\t0\t値：\\n0 オブジェクト名 Object Name false\\n1 XML ID XML ID true",
            "Units > Identify Objects By\tartNamesAreXMLIDs\tInteger\t0\t値：\\n0 オブジェクト名 Object Name false\\n1 XML ID XML ID true",
            "ガイド・グリッド > ガイド > カラー «Red»\tGuide/Color/red\tReal\t0.29\tガイドカラー（RGB 0〜1）。例：シアン 0 / 1 / 1、ライトブルー 0.29 / 0.52 / 1.0。red・green・blue の3キーをそろえて書く（PresetManager.jsx）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "Guides & Grid > Guides > Color «Red»\tGuide/Color/red\tReal\t0.29\tガイドカラー（RGB 0〜1）。例：シアン 0 / 1 / 1、ライトブルー 0.29 / 0.52 / 1.0。red・green・blue の3キーをそろえて書く（PresetManager.jsx）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "ガイド・グリッド > ガイド > カラー «Green»\tGuide/Color/green\tReal\t0.52\tガイドカラー（RGB 0〜1）。red・green・blue の3キーをそろえて書く（PresetManager.jsx）",
            "Guides & Grid > Guides > Color «Green»\tGuide/Color/green\tReal\t0.52\tガイドカラー（RGB 0〜1）。red・green・blue の3キーをそろえて書く（PresetManager.jsx）",
            "ガイド・グリッド > ガイド > カラー «Blue»\tGuide/Color/blue\tReal\t1.0\tガイドカラー（RGB 0〜1）。red・green・blue の3キーをそろえて書く（PresetManager.jsx）",
            "Guides & Grid > Guides > Color «Blue»\tGuide/Color/blue\tReal\t1.0\tガイドカラー（RGB 0〜1）。red・green・blue の3キーをそろえて書く（PresetManager.jsx）",
            "ガイド・グリッド > ガイド > スタイル\tGuide/Style\tInteger\t0\t値：\\n0 ライン Lines\\n1 点線　ドット\\n値：0 ライン / 1 点線（PresetManager.jsx）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "Guides & Grid > Guides > Style\tGuide/Style\tInteger\t0\t値：\\n0 ライン Lines\\n1 点線　ドット\\n値：0 ライン / 1 点線（PresetManager.jsx）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "ガイド・グリッド > グリッド > カラー «Red»\tGrid/Color/Dark/r\tReal\t0.8000000119",
            "Guides & Grid > Grid > Color «Red»\tGrid/Color/Dark/r\tReal\t0.8000000119",
            "ガイド・グリッド > グリッド > カラー «Green»\tGrid/Color/Dark/g\tReal\t0.8000000119",
            "Guides & Grid > Grid > Color «Green»\tGrid/Color/Dark/g\tReal\t0.8000000119",
            "ガイド・グリッド > グリッド > カラー «Blue»\tGrid/Color/Dark/b\tReal\t0.8000000119",
            "Guides & Grid > Grid > Color «Blue»\tGrid/Color/Dark/b\tReal\t0.8000000119",
            "ガイド・グリッド > グリッド > カラー «Red»\tGrid/Color/Lite/r\tReal\t0.8999999762",
            "Guides & Grid > Grid > Color «Red»\tGrid/Color/Lite/r\tReal\t0.8999999762",
            "ガイド・グリッド > グリッド > カラー «Green»\tGrid/Color/Lite/g\tReal\t0.8999999762",
            "Guides & Grid > Grid > Color «Green»\tGrid/Color/Lite/g\tReal\t0.8999999762",
            "ガイド・グリッド > グリッド > カラー «Blue»\tGrid/Color/Lite/b\tReal\t0.8999999762",
            "Guides & Grid > Grid > Color «Blue»\tGrid/Color/Lite/b\tReal\t0.8999999762",
            "ガイド・グリッド > グリッド > スタイル\tGrid/Style\tInteger\t0\t値：\\n0 ライン Lines\\n1 点線　ドット\\n次回起動時に適用",
            "Guides & Grid > Grid > Style\tGrid/Style\tInteger\t0\t値：\\n0 ライン Lines\\n1 点線　ドット\\n次回起動時に適用",
            "ガイド・グリッド > グリッド > グリッド «横方向»\tGrid/Horizontal/Spacing\tReal\t72.0",
            "Guides & Grid > Grid > Gridline every «Horizontal»\tGrid/Horizontal/Spacing\tReal\t72.0",
            "ガイド・グリッド > グリッド > グリッド «縦方向»\tGrid/Vertical/Spacing\tReal\t72.0",
            "Guides & Grid > Grid > Gridline every «Vertical»\tGrid/Vertical/Spacing\tReal\t72.0",
            "ガイド・グリッド > グリッド > 分割数 «横方向»\tGrid/Horizontal/Ticks\tInteger\t8",
            "Guides & Grid > Grid > Subdivisions «Horizontal»\tGrid/Horizontal/Ticks\tInteger\t8",
            "ガイド・グリッド > グリッド > 分割数 «縦方向»\tGrid/Vertical/Ticks\tInteger\t8",
            "Guides & Grid > Grid > Subdivisions «Vertical»\tGrid/Vertical/Ticks\tInteger\t8",
            "ガイド・グリッド > グリッド > 背面にグリッドを表示\tGrid/Posn\tBoolean\t0\t次回起動時に適用",
            "Guides & Grid > Grid > Grids In Back\tGrid/Posn\tBoolean\t0\t次回起動時に適用",
            "ガイド・グリッド > グリッド > ピクセルグリッドを表示 (600% ズーム以上)\tGuide/ShowPixelGrid\tBoolean\t1",
            "Guides & Grid > Grid > Show Pixel Grid (Above 600% Zoom)\tGuide/ShowPixelGrid\tBoolean\t1",
            "スマートガイド > 表示オプション > オブジェクトガイド «Red»\tsnapomatic/Color/red_19_2\tReal\t1.0",
            "Smart Guides > Display Options > Object Guides «Red»\tsnapomatic/Color/red_19_2\tReal\t1.0",
            "スマートガイド > 表示オプション > オブジェクトガイド «Green»\tsnapomatic/Color/green_19_2\tReal\t0.2899976969",
            "Smart Guides > Display Options > Object Guides «Green»\tsnapomatic/Color/green_19_2\tReal\t0.2899976969",
            "スマートガイド > 表示オプション > オブジェクトガイド «Blue»\tsnapomatic/Color/blue_19_2\tReal\t1.0",
            "Smart Guides > Display Options > Object Guides «Blue»\tsnapomatic/Color/blue_19_2\tReal\t1.0",
            "スマートガイド > 表示オプション > グリフガイド «Red»\tsnapomatic/GlyphColor/red\tReal\t0.4313725531",
            "Smart Guides > Display Options > Glyph Guides «Red»\tsnapomatic/GlyphColor/red\tReal\t0.4313725531",
            "スマートガイド > 表示オプション > グリフガイド «Green»\tsnapomatic/GlyphColor/green\tReal\t0.8039215803",
            "Smart Guides > Display Options > Glyph Guides «Green»\tsnapomatic/GlyphColor/green\tReal\t0.8039215803",
            "スマートガイド > 表示オプション > グリフガイド «Blue»\tsnapomatic/GlyphColor/blue\tReal\t0.2941176593",
            "Smart Guides > Display Options > Glyph Guides «Blue»\tsnapomatic/GlyphColor/blue\tReal\t0.2941176593",
            "スマートガイド > 表示オプション > 整列ガイド\tsmartGuides/showAlignmentGuides\tBoolean\t1",
            "Smart Guides > Display Options > Alignment Guides\tsmartGuides/showAlignmentGuides\tBoolean\t1",
            "スマートガイド > 表示オプション > オブジェクトのハイライト表示\tsmartGuides/showObjectHighlighting\tBoolean\t0",
            "Smart Guides > Display Options > Object Highlighting\tsmartGuides/showObjectHighlighting\tBoolean\t0",
            "スマートガイド > 表示オプション > 変形ツール\tsmartGuides/showToolGuides\tBoolean\t0",
            "Smart Guides > Display Options > Transform Tools\tsmartGuides/showToolGuides\tBoolean\t0",
            "スマートガイド > 表示オプション > アンカーとパスのヒント表示\tsmartGuides/showLabels\tBoolean\t1",
            "Smart Guides > Display Options > Anchor/Path Labels\tsmartGuides/showLabels\tBoolean\t1",
            "スマートガイド > 表示オプション > 計測のヒント表示\tsmartGuides/showReadouts\tBoolean\t0",
            "Smart Guides > Display Options > Measurement Labels\tsmartGuides/showReadouts\tBoolean\t0",
            "スマートガイド > 表示オプション > 間隔ガイド\tsmartGuides/showSpacingGuides\tBoolean\t1",
            "Smart Guides > Display Options > Spacing Guides\tsmartGuides/showSpacingGuides\tBoolean\t1",
            "スマートガイド > 表示オプション > コンストラクションガイド «真偽値»\tsmartGuides/showConstructionGuides\tBoolean\t0",
            "Smart Guides > Display Options > Construction Guides «Checked»\tsmartGuides/showConstructionGuides\tBoolean\t0",
            "\tsmartGuides/anglesCount\tInteger\t4294967295\tスマートガイド > 表示オプション > コンストラクションガイドの入力してある角度の数",
            "\tsmartGuides/customAnglesCount\tInteger\t0",
            "\tsmartGuides/angles0\tReal\t0",
            "\tsmartGuides/angles1\tReal\t45",
            "\tsmartGuides/angles2\tReal\t90",
            "\tsmartGuides/angles3\tReal\t135",
            "\tsmartGuides/customAngles0\tReal\t0",
            "\tsmartGuides/customAngles1\tReal\t45",
            "\tsmartGuides/customAngles2\tReal\t90",
            "\tsmartGuides/customAngles3\tReal\t135",
            "\tsmartGuides/snapToActiveArtboardContent\tBoolean\t0",
            "スマートガイド > スナップの許容値\tsmartGuides/tolerance\tInteger\t6",
            "Smart Guides > Snapping Tolerance\tsmartGuides/tolerance\tInteger\t6",
            "\tsmartGuides/angularTolerance\tInteger\t2",
            "\tsmartGuides/rotationalSnapArcTolerance\tInteger\t6",
            "\tsmartGuides/showRotationalGuides\tBoolean\t1",
            "表示 > スマートガイド\tsmartGuides/isEnabled\tBoolean\t1\tメニューのチェックマークは更新されない",
            "View > Smart Guides\tsmartGuides/isEnabled\tBoolean\t1\tメニューのチェックマークは更新されない",
            "スライス > スライス番号を表示\tplugin/AdobeSlicingPlugin/showSliceNumbers\tBoolean\t1",
            "Slices > Show Slice Numbers\tplugin/AdobeSlicingPlugin/showSliceNumbers\tBoolean\t1",
            "スライス > 線のカラー «Red»\tplugin/AdobeSlicingPlugin/feedback/red\tInteger\t65535",
            "Slices > Line Color «Red»\tplugin/AdobeSlicingPlugin/feedback/red\tInteger\t65535",
            "スライス > 線のカラー «Green»\tplugin/AdobeSlicingPlugin/feedback/green\tInteger\t19005",
            "Slices > Line Color «Green»\tplugin/AdobeSlicingPlugin/feedback/green\tInteger\t19005",
            "スライス > 線のカラー «Blue»\tplugin/AdobeSlicingPlugin/feedback/blue\tInteger\t19005",
            "Slices > Line Color «Blue»\tplugin/AdobeSlicingPlugin/feedback/blue\tInteger\t19005",
            "\tplugin/AdobeSlicingPlugin/divideVerticallyPixelCount\tInteger\t0",
            "\tplugin/AdobeSlicingPlugin/divideVerticallyCount\tInteger\t1",
            "\tplugin/AdobeSlicingPlugin/divideHorizontallyPixelCount\tInteger\t0",
            "\tplugin/AdobeSlicingPlugin/divideHorizontallyCount\tInteger\t1",
            "\tplugin/AdobeSlicingPlugin/dividePreview\tInteger\t0",
            "\tplugin/AdobeSlicingPlugin/divideEvenVertically\tInteger\t1",
            "\tplugin/AdobeSlicingPlugin/divideEvenHorizontally\tInteger\t1",
            "\tplugin/AdobeSlicingPlugin/divideVertically\tInteger\t1",
            "\tplugin/AdobeSlicingPlugin/divideHorizontally\tInteger\t1",
            "\tplugin/AdobeSlicingPlugin/addSlicingTools\tInteger\t1",
            "\tplugin/AdobeSlicingPlugin/lockSlices\tInteger\t0",
            "\tplugin/AdobeSlicingPlugin/huggableText\tInteger\t1",
            "ハイフネーション > 言語\thyphenation/language\tInteger\t0",
            "Hyphenation > Default Language\thyphenation/language\tInteger\t0",
            "プラグイン・仮想記憶ディスク > 追加プラグインフォルダー «真偽値»\tloadAdditionalPlugins\tBoolean\t0",
            "Plug-ins & Scratch Disks > Additional Plug-ins Folder «Checked»\tloadAdditionalPlugins\tBoolean\t0",
            "プラグイン・仮想記憶ディスク > 追加プラグインフォルダー «フォルダ»\tplugins\tString\t/Users/«username»/Desktop/plugins",
            "Plug-ins & Scratch Disks > Additional Plug-ins Folder «Folder»\tplugins\tString\t/Users/«username»/Desktop/plugins",
            "プラグイン・仮想記憶ディスク > 仮想記憶ディスク > ディスク 1\trastersys/scratch/primary\tString\t",
            "Plug-ins & Scratch Disks > Scratch Disks > Primary\trastersys/scratch/primary\tString\t",
            "プラグイン・仮想記憶ディスク > 仮想記憶ディスク > ディスク 2\trastersys/scratch/secondary\tString\t",
            "Plug-ins & Scratch Disks > Scratch Disks > Secondary\trastersys/scratch/secondary\tString\t",
            "ユーザーインターフェイス > 明るさ\tuiBrightness\tReal\t1.0\t0.5 より大きければ明るい UI（スクリプトの明暗判定に使える）\\n値（4段階の離散値）：\\n0.0 暗\\n0.5 やや暗\\n0.50999999046326 やや明\\n1.0 明\\n0.5 と 0.50999999 は別の段階なので、比較は許容値 0.001 程度で（PresetManager.jsx）\\nset は値を書き込むだけでは画面に反映されない。書き込んだ後に環境設定（ユーザーインターフェイス、app.executeMenuCommand('UIPref')）を開き、矢印キー＋Return で確定する。環境設定はモーダルダイアログを閉じてから開く（PresetManager.jsx）",
            "User Interface > Brightness\tuiBrightness\tReal\t1.0\t0.5 より大きければ明るい UI（スクリプトの明暗判定に使える）\\n値（4段階の離散値）：\\n0.0 暗\\n0.5 やや暗\\n0.50999999046326 やや明\\n1.0 明\\n0.5 と 0.50999999 は別の段階なので、比較は許容値 0.001 程度で（PresetManager.jsx）\\nset は値を書き込むだけでは画面に反映されない。書き込んだ後に環境設定（ユーザーインターフェイス、app.executeMenuCommand('UIPref')）を開き、矢印キー＋Return で確定する。環境設定はモーダルダイアログを閉じてから開く（PresetManager.jsx）",
            "ユーザーインターフェイス > カンバスカラー\tuiCanvasIsWhite\tInteger\t0\t値：\\n0 ユーザーインターフェイスの明るさに一致させる Match User Interface Brightness\\n1 ホワイト White\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "User Interface > Canvas Color\tuiCanvasIsWhite\tInteger\t0\t値：\\n0 ユーザーインターフェイスの明るさに一致させる Match User Interface Brightness\\n1 ホワイト White\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "ユーザーインターフェイス > 自動的にアイコンパネル化\tuiPersistDrawers\tBoolean\t1\t{true: OFF, false: ON} 次回起動時に適用",
            "User Interface > Auto-Collapse Iconic Panels\tuiPersistDrawers\tBoolean\t1\t{true: OFF, false: ON} 次回起動時に適用",
            "ユーザーインターフェイス > タブでドキュメントを開く\tuiOpenDocumentsAsTabs\tBoolean\t1\t次回起動時に適用",
            "User Interface > Open Documents As Tabs\tuiOpenDocumentsAsTabs\tBoolean\t1\t次回起動時に適用",
            "ユーザーインターフェイス > 大きなタブ\tUIPreferences/workspaceTabsSize\tInteger\t2\t値：\\n1 Small\\n2 Large\\n次回起動時に適用",
            "User Interface > Large Tabs\tUIPreferences/workspaceTabsSize\tInteger\t2\t値：\\n1 Small\\n2 Large\\n次回起動時に適用",
            "ユーザーインターフェイス > UI スケール\tUIPreferences/appScaleFactor\tReal\t1.0",
            "User Interface > UI Scaling\tUIPreferences/appScaleFactor\tReal\t1.0",
            "ユーザーインターフェイス > UI スケール > 比率を保持してカーソルを拡大・縮小\tUIPreferences/scaleCursor\tBoolean\t1",
            "User Interface > UI Scaling > Scale Cursor Proportionately\tUIPreferences/scaleCursor\tBoolean\t1",
            "\tUIPreferences/defaultScaleFactorLaunch\tInteger\t1",
            "\tUIPreferences/scaleUI\tBoolean\t1",
            "\tUIPreferences/snapUIScaleFactor\tInteger\t1",
            "\tuiScrollButtonPosition\tInteger\t0",
            "パフォーマンス > GPU パフォーマンス > GPU パフォーマンス\tPerformance/EnableGPU_Ver19_2\tBoolean\t1",
            "Performance > GPU Performance > GPU Performance\tPerformance/EnableGPU_Ver19_2\tBoolean\t1",
            "パフォーマンス > GPU パフォーマンス > GPU パフォーマンス > アニメーションズーム\tPerformance/AnimZoom\tBoolean\t0",
            "Performance > GPU Performance > GPU Performance > Animated Zoom\tPerformance/AnimZoom\tBoolean\t0",
            "\tPerformance/AutoSwitchEngine\tInteger\t1",
            "\tPerformance/ResponsiveZoom\tInteger\t1",
            "\tPerformance/GPUSupported\tInteger\t1",
            "パフォーマンス > その他 > ヒストリー数\tmaximumUndoDepth\tInteger\t50\tヒストリー数（ダイアログでの範囲は 1〜1000、既定 100）（PresetManager.jsx）",
            "Performance > Others > History States\tmaximumUndoDepth\tInteger\t50\tヒストリー数（ダイアログでの範囲は 1〜1000、既定 100）（PresetManager.jsx）",
            "パフォーマンス > その他 > リアルタイムの描画と編集\tLiveEdit_State_Machine\tBoolean\t0",
            "Performance > Others > Real-Time Drawing and Editing\tLiveEdit_State_Machine\tBoolean\t0",
            "\tLiveEdit_Global_EPF_Limit\tReal\t5.0",
            "ファイル管理 > ファイル保存オプション > 復帰データを次の間隔で自動保存 «真偽値»\tCrashRecovery/AutomaticallySave\tBoolean\t0",
            "File Handling > File Save Options > Automatically Save Recovery Data Every «Checked»\tCrashRecovery/AutomaticallySave\tBoolean\t0",
            "ファイル管理 > ファイル保存オプション > 復帰データを次の間隔で自動保存 «時間»\tCrashRecovery/IdleLoopTimeInterval\tInteger\t0",
            "File Handling > File Save Options > Automatically Save Recovery Data Every «Time»\tCrashRecovery/IdleLoopTimeInterval\tInteger\t0",
            "ファイル管理 > ファイル保存オプション > 復帰データを次の間隔で自動保存 «フォルダ»\tCrashRecovery/RecoveryFolderLocation\tString\t\t値：\\n/Users/«username»/Library/Preferences/Adobe Illustrator 26 Settings/ja_JP/DataRecovery",
            "File Handling > File Save Options > Automatically Save Recovery Data Every «Folder»\tCrashRecovery/RecoveryFolderLocation\tString\t\t値：\\n/Users/«username»/Library/Preferences/Adobe Illustrator 26 Settings/ja_JP/DataRecovery",
            "ファイル管理 > ファイル保存オプション > 複雑なドキュメントではデータの復元を無効にする\tCrashRecovery/TurnOffForComplexDocument\tBoolean\t1",
            "File Handling > File Save Options > Turn off Data Recovery for complex documents\tCrashRecovery/TurnOffForComplexDocument\tBoolean\t1",
            "ファイル管理 > ファイル保存オプション > バックグラウンドで保存\tenableBackgroundSave\tInteger\t0\t動作しない?",
            "File Handling > File Save Options > Save in Background\tenableBackgroundSave\tInteger\t0\t動作しない?",
            "ファイル管理 > ファイル保存オプション > バックグラウンドで書き出し\tenableBackgroundExport\tInteger\t0\t動作しない?",
            "File Handling > File Save Options > Export in Background\tenableBackgroundExport\tInteger\t0\t動作しない?",
            "ファイル管理 > ファイル保存オプション > クラウドドキュメントを次の間隔で自動保存 «真偽値»\tcloudAIEnableAutoSave\tInteger\t0\t動作しない?",
            "File Handling > File Save Options > Automatically Save Cloud Documents Every «Checked»\tcloudAIEnableAutoSave\tInteger\t0\t動作しない?",
            "ファイル管理 > ファイル保存オプション > クラウドドキュメントを次の間隔で自動保存 «時間»\tcloudAIAutoSaveTimerInterval\tInteger\t300\t動作しない?",
            "File Handling > File Save Options > Automatically Save Cloud Documents Every «Time»\tcloudAIAutoSaveTimerInterval\tInteger\t300\t動作しない?",
            "ファイル管理 > ファイル保存オプション > 初期設定のファイルの場所\tAdobeSaveAsCloudDocumentPreference\tInteger\t0\t値：\\n0 コンピューター Local\\n1 Creative Cloud Creative Cloud\\n26.3.1のバグで有効になった。実際に動作する\\ntrue＝クラウドドキュメント／false＝コンピューターに保存（保存の既定）（PresetManager.jsx）",
            "\tplugin/FileClipboard/defaultCloudSave\tInteger\t1",
            "ファイル管理 > ファイル > 最近使用したファイルの表示数 (0 〜 30)\tRecentFileNumber\tInteger\t20",
            "File Handling > Files > Number of Recent Files to Display (0-30)\tRecentFileNumber\tInteger\t20",
            "ファイル管理 > ファイル > 低速ネットワーク上のファイルを開くおよび保存する時間を最適化\tplugin/FileClipboard/aiOptimizeNetworkOperations\tBoolean\t1",
            "File Handling > Files > Optimize File Open and Save Time on Slow Networks\tplugin/FileClipboard/aiOptimizeNetworkOperations\tBoolean\t1",
            "ファイル管理 > ファイル > リンクされた EPS に低解像度の表示用画像を使用\tuseLowResProxy\tBoolean\t0",
            "File Handling > Files > Use Low Resolution Proxy for Linked EPS\tuseLowResProxy\tBoolean\t0",
            "ファイル管理 > ファイル > ピクセルプレビューでビットマップ画像をアンチエイリアス処理した画像として表示\tDisplayBitmapsAsAntiAliasedPixelPreview\tBoolean\t0",
            "File Handling > Files > Display Bitmaps as Anti-Aliased images in Pixel Preview\tDisplayBitmapsAsAntiAliasedPixelPreview\tBoolean\t0",
            "ファイル管理 > ファイル > リンクを更新\tplugin/FileClipboard/linkoptions\tInteger\t0\t値：\\n0 自動 Automatically\\n1 Automatically Manually\\n2 修正されるときに確認 Ask When Modified\\n初期値：2\\n値（リンクを更新）：0 自動 / 1 手動 / 2 変更時に確認（PresetManager.jsx）",
            "File Handling > Files > Update Links\tplugin/FileClipboard/linkoptions\tInteger\t0\t値：\\n0 自動 Automatically\\n1 Automatically Manually\\n2 修正されるときに確認 Ask When Modified\\n初期値：2\\n値（リンクを更新）：0 自動 / 1 手動 / 2 変更時に確認（PresetManager.jsx）",
            "ファイル管理 > ファイル > 「オリジナルの編集」にシステムデフォルトを使用\tuseSysDefEdit\tBoolean\t1",
            "File Handling > Files > Use System Defaults for ‘Edit Original’\tuseSysDefEdit\tBoolean\t1",
            "ファイル管理 > フォント > Adobe Fonts を自動アクティベート\tAutoActivateMissingFont\tBoolean\t1",
            "File Handling > Fonts > Auto-activate Adobe Fonts\tAutoActivateMissingFont\tBoolean\t1",
            "クリップボードの処理 > クリップボード > コピー時 > SVG コードを含める\tplugin/FileClipboard/copySVGCode\tBoolean\t1",
            "Clipboard Handling > Clipboard > On Copy > Include SVG Code\tplugin/FileClipboard/copySVGCode\tBoolean\t1",
            "クリップボードの処理 > クリップボード > 終了時 > PDF\tplugin/FileClipboard/copyAsPDF\tBoolean\t1",
            "Clipboard Handling > Clipboard > On Quit > PDF\tplugin/FileClipboard/copyAsPDF\tBoolean\t1",
            "クリップボードの処理 > クリップボード > 終了時 > AICB (透明サポートなし) «真偽値»\tplugin/FileClipboard/copyAsAICB\tBoolean\t0",
            "Clipboard Handling > Clipboard > On Quit > AICB (no transparency support) «Checked»\tplugin/FileClipboard/copyAsAICB\tBoolean\t0",
            "クリップボードの処理 > クリップボード > 終了時 > AICB (透明サポートなし) «オプション»\tplugin/FileClipboard/AICBOption\tInteger\t1\t値：\\n0 パスを保持 Preserve Paths\\n1 アピアランスとオーバープリントを保持 Preserve Appearance And Overprints\\n初期値：1",
            "Clipboard Handling > Clipboard > On Quit > AICB (no transparency support) «Option»\tplugin/FileClipboard/AICBOption\tInteger\t1\t値：\\n0 パスを保持 Preserve Paths\\n1 アピアランスとオーバープリントを保持 Preserve Appearance And Overprints\\n初期値：1",
            "\tplugin/FileClipboard/flatten\tInteger\t1",
            "クリップボードの処理 > ペースト時 > 書式なしでテキストをペースト\tplugin/FileClipboard/pasteWithoutFormatting\tBoolean\t0",
            "Clipboard Handling > When Pasting > Paste text without Formatting\tplugin/FileClipboard/pasteWithoutFormatting\tBoolean\t0",
            "\tplugin/FileClipboard/retainClipboardData\tInteger\t0",
            "\tplugin/FileClipboard/copySVGCBFormat\tInteger\t1",
            "\tplugin/FileClipboard/copySVGForMuseCBFormat\tInteger\t1",
            "\tplugin/FileClipboard/appendExtension\tInteger\t0",
            "\tplugin/FileClipboard/lowerCase\tInteger\t1",
            "ブラックのアピアランス > RGB およびグレースケールデバイス上のブラックの表示オプション > スクリーン\tblackPreservation/Onscreen\tInteger\t0\t値：\\n0 すべてのブラックを正確に表示 Display All Blacks Accurately\\n1 すべてのブラックをリッチブラックとして表示 Display All Blacks as Rich Black\\n初期値：1",
            "Appearance of Black > Options for Black on RGB and Grayscale Devices > On Screen\tblackPreservation/Onscreen\tInteger\t0\t値：\\n0 すべてのブラックを正確に表示 Display All Blacks Accurately\\n1 すべてのブラックをリッチブラックとして表示 Display All Blacks as Rich Black\\n初期値：1",
            "ブラックのアピアランス > RGB およびグレースケールデバイス上のブラックの表示オプション > プリント / 書き出し\tblackPreservation/Export\tInteger\t0\t値：\\n0 すべてのブラックを正確に表示 Display All Blacks Accurately\\n1 すべてのブラックをリッチブラックとして表示 Display All Blacks as Rich Black\\n初期値：1",
            "Appearance of Black > Options for Black on RGB and Grayscale Devices > Printing / Exporting\tblackPreservation/Export\tInteger\t0\t値：\\n0 すべてのブラックを正確に表示 Display All Blacks Accurately\\n1 すべてのブラックをリッチブラックとして表示 Display All Blacks as Rich Black\\n初期値：1",
            "デバイス > Wacom の有効化\tDevices/EnableWacom\tBoolean\t1",
            "Devices > Enable Wacom\tDevices/EnableWacom\tBoolean\t1",
            "書式 > 合成フォント… > サンプルを表示\tshowSample\tBoolean\t1\tJapanese language features",
            "書式 > 合成フォント… > ズーム\tzoom\tInteger\t400\tJapanese language features",
            "書式 > 合成フォント… > x ハイト\txHeight\tBoolean\t0\tJapanese language features",
            "書式 > 合成フォント… > アセンダハイト\tmaxAscender\tBoolean\t0\tJapanese language features",
            "書式 > 合成フォント… > 最大のアセント / ディセント\tmaxAscentDesce\tBoolean\t0\tJapanese language features",
            "書式 > 合成フォント… > キャップハイト\tcapHeight\tBoolean\t0\tJapanese language features",
            "書式 > 合成フォント… > 欧文ベースライン\tbaseline\tBoolean\t0\tJapanese language features",
            "書式 > 合成フォント… > 仮想ボディ\temBox\tBoolean\t0\tJapanese language features",
            "書式 > 合成フォント… > 平均字面\ticfBox\tBoolean\t0\tJapanese language features",
            "文字パネル > グリフにスナップを表示\tshowSnapToGlyphOpt\tBoolean\t1",
            "Character Panel > Show Snap to Glyph Options\tshowSnapToGlyphOpt\tBoolean\t1",
            "文字パネル > フォントの高さを表示\tfontHeightOption\tBoolean\t0\t次回起動時に適用",
            "Character Panel > Show Font Height Options\tfontHeightOption\tBoolean\t0\t次回起動時に適用",
            "文字パネル > 文字タッチツール\ttext/showTextTouchUpButton\tBoolean\t1\t次回起動時に適用",
            "Character Panel > Touch Type Tool\ttext/showTextTouchUpButton\tBoolean\t1\t次回起動時に適用",
            "\tsnapToGlyph\tBoolean\t0\tグリフにスナップ機能全体を有効にするかどうか",
            "文字パネル > グリフにスナップ > 仮想ボディ\tjapTextEMBoxLineSnapping\tBoolean\t1\tJapanese language features. Em Box",
            "文字パネル > グリフにスナップ > 仮想ボディの中心\tjapTextCentreSnapping\tBoolean\t1\tJapanese language features. Em Box Center",
            "文字パネル > グリフにスナップ > 字形の境界\ttextLineBoundsSnapping\tBoolean\t1",
            "Character Panel > Snap to Glyph > Glyph Bounds\ttextLineBoundsSnapping\tBoolean\t1",
            "文字パネル > グリフにスナップ > ベースライン\ttextBaselineLineSnapping\tBoolean\t1",
            "Character Panel > Snap to Glyph > Baseline\ttextBaselineLineSnapping\tBoolean\t1",
            "文字パネル > グリフにスナップ > 角度ガイド\tangularGuides\tBoolean\t1",
            "Character Panel > Snap to Glyph > Angular Guides\tangularGuides\tBoolean\t1",
            "文字パネル > グリフにスナップ > アンカーポイント\ttextAnchorPointSnapping\tBoolean\t1",
            "Character Panel > Snap to Glyph > Anchor Point\ttextAnchorPointSnapping\tBoolean\t1",
            "Character Panel > Snap to Glyph > x-height\ttextXHeightSnapping\tBoolean\t1",
            "Character Panel > Snap to Glyph > Proximity Guides\ttextImportantVisualLinesSnapping\tBoolean\t0",
            "\ttextAllVisualLinesSnapping\tBoolean\t0",
            "\ttextFirstLineSnapping\tBoolean\t0",
            "\ttextLineSnapping\tBoolean\t1",
            "整列パネル > 字形の境界に整列 > ポイント文字\tEnableActualPointTextSpaceAlign\tBoolean\t0",
            "Align Panel > Align to Glyph Bounds > Point Text\tEnableActualPointTextSpaceAlign\tBoolean\t0",
            "整列パネル > 字形の境界に整列 > エリア内文字\tEnableActualAreaTextSpaceAlign\tBoolean\t0",
            "Align Panel > Align to Glyph Bounds > Area Text\tEnableActualAreaTextSpaceAlign\tBoolean\t0",
            "\tAi_SaveBackupFiles\tInteger\t1",
            "\tsaveTriggeredFromShareForReview\tInteger\t0",
            "\tDefaultSaveLocation\tInteger\t0",
            "\tfileHandling/RenderLinksInLowRes\tInteger\t0\tリンク画像を低画質プレビューで表示する。［ファイル管理 > ファイル > リンクされた EPS に低解像度の表示用画像を使用］とは別の概念",
            "\tadaptiveShareButtonClicked\tInteger\t0",
            "\tAIMacPluginsLoadingErrorDueToWrongArchitectureSHA\tString\t",
            "\tEPSResolution\tInteger\t300\t動作しない？ 常に0になる",
            "\tprint/gradient/useLevel3\tInteger\t1",
            "\tAi253_FastExit\tInteger\t1",
            "\tAIUserCancelledLogOff\tInteger\t0",
            "\tRotateViewToolForceEjected\tInteger\t0",
            "\tRotateViewToolForceInjected\tInteger\t1",
            "\ttemplatesBrowsed\tInteger\t1",
            "\tAdobeOpenCloudDocumentPreference\tInteger\t0",
            "\tCharPara/AddToCCLibraryCheckBox\tInteger\t0",
            "\tCloud AI Incremental Save Enabled Pref\tInteger\t1\t動作しない？",
            "\tCloudAIContentTypePref\tInteger\t1\t動作しない？",
            "\tfilenaming/style\tInteger\t1",
            "\taiOptimizedNetworkOnboardingShown\tInteger\t1",
            "\tDropLegacyScale9\tInteger\t0",
            "\tAlignToPixelGridDuringExport\tInteger\t1",
            "\tuseRealDWG\tInteger\t0",
            "\tdefaultPath\tString\t/Users/«username»/Desktop",
            "\tPlacedObject/DisabledNetworkLinkedObject\tInteger\t0",
            "\tDesignLibraryFTUEShown\tInteger\t1",
            "\tisHexInUpperCase\tInteger\t1",
            "\tGlobalEdit/MaximumSimilarArtsToProcess\tInteger\t1000\t［オブジェクトを一括選択］で選択できる最大個数",
            "\tperspectivegrid/expandAnswer\tInteger\t0",
            "変形パネル > 楕円形のプロパティ > 扇形の角度を制限\tLiveShapes/constrainPieAngles\tBoolean\t0",
            "Transform Panel > Ellipse Properties > Constrain Pie Angles\tLiveShapes/constrainPieAngles\tBoolean\t0",
            "変形パネル > 長方形のプロパティ > 角丸の半径値をリンク\tLiveShapes/constrainRadii\tBoolean\t1",
            "Transform Panel > Rectangle Properties > Link Corner Radius Values\tLiveShapes/constrainRadii\tBoolean\t1",
            "変形パネル > 長方形のプロパティ > 縦横比を固定\tLiveShapes/constrainDimensions\tBoolean\t1\t変形パネル > 楕円形のプロパティ > 縦横比を固定\\nTransform Panel > Ellipse Properties > Constrain Width and Height Proportions",
            "Transform Panel > Rectangle Properties > Constrain Width and Height Proportions\tLiveShapes/constrainDimensions\tBoolean\t1\t変形パネル > 楕円形のプロパティ > 縦横比を固定\\nTransform Panel > Ellipse Properties > Constrain Width and Height Proportions",
            "\tLiveShapes/maxLiveShapesToShowWidgetsOn\tInteger\t20",
            "変形パネル > シェイプの作成時に表示\tLiveShapes/autoShowPropertiesUIOnCreatingShape\tBoolean\t0",
            "Transform Panel > Show on Shape Creation\tLiveShapes/autoShowPropertiesUIOnCreatingShape\tBoolean\t0",
            "\tLiveShapes/createLiveShapes\tBoolean\t1",
            "\tLiveShapes/hideWidgetsForShapeTools\tBoolean\t0",
            "\twidthToolIsSelected\tInteger\t0",
            "\tuseSnapomatic\tInteger\t1",
            "\tDontShowMissingFontDialogPreference\tBoolean\t0",
            "\tmessageIndex\tInteger\t0",
            "\tFileNewFirstTime\tInteger\t0",
            "\tenableAGMThreadedRendering\tInteger\t0",
            "\tQuickPortImpl\tInteger\t1",
            "\tpageTiling/imageBoundsColor\tString\tr=0x00, g=0x00, b=0x00\t値：\\n(r=0x00, g=0x00, b=0x00)\\nその他の型。動作不可？",
            "\tpageTiling/mediaBoundsColor\tString\tr=0x00, g=0x00, b=0x00\t値：\\n(r=0x00, g=0x00, b=0x00)\\nその他の型。動作不可？",
            "\tenableHyperThreadedRendering\tInteger\t0",
            "\tRenderingTimerEnabled\tInteger\t0",
            "\tisRulerOriginTopLeft\tInteger\t1",
            "\tartboard/transform\tInteger\t0",
            "\tartboard/height\tReal\t1080.0",
            "\tartboard/width\tReal\t1920.0",
            "\txml/UID/autoBaseNameUS\tString\t",
            "\taiFileFormat/documentPSLevel\tInteger\t3",
            "Illustrator オプション > オプション > PDF 互換ファイルを作成\taiFileFormat/PDFCompatibility\tBoolean\t1",
            "Illustrator Options > Options > Create PDF Compatible File\taiFileFormat/PDFCompatibility\tBoolean\t1",
            "\taiFileFormat/enableATEWriteRecovery\tInteger\t0",
            "\taiFileFormat/enableATEReadRecovery\tInteger\t0",
            "\taiFileFormat/dumpPGFwithoutShortcut\tInteger\t0",
            "\taiFileFormat/enableContentRecovery\tInteger\t0",
            "\taiFileFormat/cmykPostScript\tInteger\t0",
            "\taiFileFormat/clipboardPSLevel\tInteger\t3",
            "\tIgnoreWordWithNumber\tBoolean\t0",
            "\tIgnoreRomanNumeral\tBoolean\t0",
            "\tIgnoreWordAllCap\tBoolean\t0",
            "\tIgnoreUncapSentenceStart\tBoolean\t0",
            "\tIgnoreRepeatedWord\tBoolean\t0",
            "\tcolorModel\tInteger\t2\t最後に新規で開いた書類のカラーモード。1: RGB, 2: CMYK",
            "\tprinterResolution\tReal\t800.0",
            "\tdoSplitPath\tInteger\t0",
            "\tshowPlacedImages\tInteger\t0",
            "\topen/legacyGradientMeshConversion\tInteger\t0",
            "\tLegacyArtboardOptions/artworkBounds\tInteger\t0",
            "\tLegacyArtboardOptions/pageTiles\tInteger\t0",
            "\tLegacyArtboardOptions/cropAreas\tInteger\t1",
            "\tLegacyArtboardOptions/artboard\tInteger\t1",
            "\tLegacyArtboardOptions/ShowDialog\tInteger\t1",
            "\twarning/legacyTextConversion\tInteger\t1",
            "\twarning/colorConversion\tInteger\t1",
            "\twarning/liveTrace\tInteger\t1",
            "\twarning/fontProblem\tInteger\t1",
            "\twarning/LegacyDictionaryChange\tInteger\t0",
            "\tartnewdialog/docppi\tReal\t72.0",
            "\tartnewdialog/pixelperfectobjects\tInteger\t0",
            "\tartnewdialog/rasterresolution\tInteger\t300",
            "\tartnewdialog/pixelaspectratio\tReal\t1.0",
            "\tartnewdialog/transparencygrid_color2_blue\tInteger\t43530",
            "\tartnewdialog/transparencygrid_color2_green\tInteger\t43530",
            "\tartnewdialog/transparencygrid_color2_red\tInteger\t43529",
            "\tartnewdialog/transparencygrid_color1_blue\tInteger\t43530",
            "\tartnewdialog/transparencygrid_color1_green\tInteger\t43530",
            "\tartnewdialog/transparencygrid_color1_red\tInteger\t43529",
            "\tartnewdialog/transparencygrid\tInteger\t1",
            "\tartnewdialog/previewmode\tInteger\t2",
            "\tartnewdialog/bleedsLocked\tInteger\t1",
            "\tartnewdialog/bottomBleedSize\tReal\t0.0",
            "\tartnewdialog/topBleedSize\tReal\t0.0",
            "\tartnewdialog/rightBleedSize\tReal\t0.0",
            "\tartnewdialog/leftBleedSize\tReal\t0.0",
            "\tartnewdialog/documentType\tInteger\t1",
            "\tartnewdialog/fileSize\tInteger\t84",
            "\tartnewdialog/unit\tInteger\t2",
            "\tartnewdialog/isArtboardLayoutLeftToRight\tInteger\t1",
            "\tartnewdialog/artboardLayout\tInteger\t0",
            "\tartnewdialog/artboardRowsOrCols\tInteger\t1",
            "\tartnewdialog/artboardSpacing\tReal\t60.0",
            "\tartnewdialog/numArtboards\tInteger\t1",
            "\tartnewdialog/advancedOptions\tInteger\t0",
            "\tartnewdialog/numRecentPresets\tInteger\t5",
            "\tWorkspace/4E4F524F4F4D53Key\tInteger\t0",
            "\tUserPreferencesMigrated\tInteger\t0",
            "\tselectedAnchorMarkType\tInteger\t9",
            "\tunselectedAnchorMarkType\tInteger\t1",
            "\tdirectionHandleMarkType\tInteger\t10",
            "\tdictionary/updateDocumentsToNewDicitonary\tInteger\t0",
            "\tattributePalette/currentUrlCount\tInteger\t0",
            "アピアランスパネル > 新規アートに基本アピアランスを適用\tAI New Art Basic Appearance\tBoolean\t1",
            "Appearance Panel > New Art Has Basic Appearance\tAI New Art Basic Appearance\tBoolean\t1",
            "リンクパネル > リンク情報を表示\tshowLinkInfo\tBoolean\t1\t次回起動時に適用",
            "Links Panel > Show Link Info\tshowLinkInfo\tBoolean\t1\t次回起動時に適用",
            "\tAI Container Overrides Object\tInteger\t1",
            "\tlayers/pastePreserveBackup\tBoolean\t0\tLayers Panel > Paste Remembers Layers (2)",
            "レイヤーパネル > コピー元のレイヤーにペースト\tlayers/pastePreserve\tBoolean\t0",
            "Layers Panel > Paste Remembers Layers\tlayers/pastePreserve\tBoolean\t0",
            "\tisRulerIn4thQuad\tBoolean\t1\t第4象限（原点左上）かどうか。falseのときは第1象限（原点左下）",
            "変形パネル > シンボル基準点を使用\tdontUseSymbolRegPoint\tBoolean\t0\t{true: OFF, false: ON}",
            "Transform Panel > Use Registration Point for Symbol\tdontUseSymbolRegPoint\tBoolean\t0\t{true: OFF, false: ON}",
            "変形パネル > パターンのみ変形\tonlyTransformPatterns\tBoolean\t0",
            "Transform Panel > Transform Pattern Only\tonlyTransformPatterns\tBoolean\t0",
            "変形パネル > 縦横比を固定\tlinkTransform\tBoolean\t0",
            "Transform Panel > Constrain Width and Height Proportions\tlinkTransform\tBoolean\t0",
            "\tInAppUpdateUI/LaunchSessionCount\tInteger\t0",
            "\taiSwapAltControlWithScrollWheel\tInteger\t0",
            "\toutlineTol\tReal\t0.0",
            "\tfreehandTol\tReal\t2.0",
            "\tprecision/userWithUnits\tInteger\t4",
            "\tprecision/percentageNoUnits\tInteger\t2",
            "\tprecision/angleNoUnits\tInteger\t3",
            "\tprecision/fixedNoUnits\tInteger\t3",
            "\tprecision/millimetersWithUnits\tInteger\t3",
            "\tprecision/centimetersWithUnits\tInteger\t4",
            "\tprecision/inchesWithUnits\tInteger\t4",
            "\tprecision/pixelsWithUnits\tInteger\t3",
            "\tprecision/pointsWithUnits\tInteger\t3",
            "\tshowFindMoreTabV2\tBoolean\t1",
            "\tshaper/constrainScaling\tBoolean\t0\tタッチワークスペースのShaper ツールで，Constrain Scalingを有効にするかどうか",
            "\tTouchPreferenceUI/SoftMessageDuration\tReal\t4.0",
            "\tTouchPreferenceUI/PreciseCursor\tBoolean\t1",
            "\tTouchPreferenceUI/TWSKeyboardDetach\tBoolean\t1",
            "\tSwatches/linkHtWidthInDlg\tInteger\t0",
            "\tSwatches/linkHVSpacingInDlg\tInteger\t0",
            "\tSwatches/moveTileBoundsWithArt\tInteger\t1",
            "\tSwatches/showTileBoundsInPDM\tInteger\t1",
            "\tSwatches/showSwatchBoundsInPDM\tInteger\t0",
            "\tSwatches/autoCloseOptionsPanelOnExitingPDM\tInteger\t1",
            "スウォッチパネル > 新規スウォッチ… > カラータイプ > グローバル\tSwatches/GlobalProcessCheckBox\tBoolean\t1",
            "Swatches Panel > New Swatch… > Color Type > Global\tSwatches/GlobalProcessCheckBox\tBoolean\t1",
            "\tSwatches/AddToCCLibraryCheckBoxNew\tBoolean\t0\t動作しない？",
            "\tDefaultPaintStyle/ActiveGradient/AIColor/Kind\tString\t\tその他の型。動作不可？",
            "\tDefaultPaintStyle/ActiveColor/AIColor/RGB/Blue\tReal\t0.0",
            "\tDefaultPaintStyle/ActiveColor/AIColor/RGB/Green\tReal\t0.0",
            "\tDefaultPaintStyle/ActiveColor/AIColor/RGB/Red\tReal\t0.0",
            "\tDefaultPaintStyle/ActiveColor/AIColor/CMYK/Black\tReal\t0.0",
            "\tDefaultPaintStyle/ActiveColor/AIColor/CMYK/Yellow\tReal\t0.0",
            "\tDefaultPaintStyle/ActiveColor/AIColor/CMYK/Magenta\tReal\t0.0",
            "\tDefaultPaintStyle/ActiveColor/AIColor/CMYK/Cyan\tReal\t0.0",
            "\tDefaultPaintStyle/ActiveColor/AIColor/Kind\tString\tCMYK\tその他の型。動作不可？",
            "\tDefaultPaintStyle/ActiveColor/AIColor/Gray/Gray\tReal\t1.0",
            "\tDefaultPaintStyle/PathStyle/Stroke/MiterLimit\tReal\t10.0",
            "\tDefaultPaintStyle/PathStyle/Stroke/Join\tString\tMiterJoin\tその他の型。動作不可？",
            "\tDefaultPaintStyle/PathStyle/Stroke/Cap\tString\tButtCap\tその他の型。動作不可？",
            "\tDefaultPaintStyle/PathStyle/Stroke/Dash/Offset\tReal\t0.0",
            "\tDefaultPaintStyle/PathStyle/Stroke/Dash/Length\tInteger\t0",
            "\tDefaultPaintStyle/PathStyle/Stroke/Width\tReal\t1.0",
            "\tDefaultPaintStyle/PathStyle/Stroke/Overprint\tInteger\t0",
            "\tDefaultPaintStyle/PathStyle/Stroke/AIColor/Kind\tString\t\tその他の型。動作不可？",
            "\tDefaultPaintStyle/PathStyle/StrokePaint\tInteger\t0",
            "\tDefaultPaintStyle/PathStyle/Fill/Overprint\tInteger\t0",
            "\tDefaultPaintStyle/PathStyle/Fill/AIColor/Gray/Gray\tReal\t1.0",
            "\tDefaultPaintStyle/PathStyle/Fill/AIColor/Kind\tString\tGray\tその他の型。動作不可？",
            "\tDefaultPaintStyle/PathStyle/FillPaint\tInteger\t1",
            "\tDefaultPaintStyle/StrokeActive\tInteger\t0",
            "\tIllustrator version\tInteger\t24",
            "\tEyedropperTool/graphicsStyleExpanded\tInteger\t1\t［スポイトツールオプション > スポイトの抽出 > アピアランス］を展開しているかどうか。次回起動時に適用",
            "\tEyedropperTool/paint/fillExpanded\tBoolean\t1\t［スポイトツールオプション > スポイトの抽出 > アピアランス > 塗り］を展開しているかどうか。次回起動時に適用",
            "\tEyedropperTool/paint/strokeExpanded\tBoolean\t1\t［スポイトツールオプション > スポイトの抽出 > アピアランス > 線］を展開しているかどうか。次回起動時に適用",
            "\tEyedropperTool/paint/objectExpanded\tBoolean\t1",
            "スポイトツールオプション > スポイトの抽出 > アピアランス\tEyedropperTool/graphicsStyle\tBoolean\t0\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance\tEyedropperTool/graphicsStyle\tBoolean\t0\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 透明\tEyedropperTool/paint/object/transparency\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Transparency\tEyedropperTool/paint/object/transparency\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 塗り\tEyedropperTool/paint/fill\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Fill\tEyedropperTool/paint/fill\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 塗り > カラー\tEyedropperTool/paint/fill/color\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Fill > Color\tEyedropperTool/paint/fill/color\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 塗り > 透明\tEyedropperTool/paint/fill/transparency\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Fill > Transparency\tEyedropperTool/paint/fill/transparency\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 塗り > オーバープリント\tEyedropperTool/paint/fill/overprint\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Fill > Overprint\tEyedropperTool/paint/fill/overprint\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線\tEyedropperTool/paint/stroke\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke\tEyedropperTool/paint/stroke\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > カラー\tEyedropperTool/paint/stroke/color\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Color\tEyedropperTool/paint/stroke/color\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > 透明\tEyedropperTool/paint/stroke/transparency\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Transparency\tEyedropperTool/paint/stroke/transparency\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > オーバープリント\tEyedropperTool/paint/stroke/overprint\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Overprint\tEyedropperTool/paint/stroke/overprint\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > 線幅\tEyedropperTool/paint/stroke/weight\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Weight\tEyedropperTool/paint/stroke/weight\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > 線端の形状\tEyedropperTool/paint/stroke/cap\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Cap\tEyedropperTool/paint/stroke/cap\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > 角の形状\tEyedropperTool/paint/stroke/join\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Join\tEyedropperTool/paint/stroke/join\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > 角の比率\tEyedropperTool/paint/stroke/miter\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Miter limit\tEyedropperTool/paint/stroke/miter\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > アピアランス > 線 > 破線パターン\tEyedropperTool/paint/stroke/dash\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Appearance > Focal Stroke > Dash pattern\tEyedropperTool/paint/stroke/dash\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > 文字スタイル\tEyedropperTool/text/characterStyle\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Character Style\tEyedropperTool/text/characterStyle\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの抽出 > 段落スタイル\tEyedropperTool/text/paragraphStyle\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Picks Up > Paragraph Style\tEyedropperTool/text/paragraphStyle\tBoolean\t1\t次回起動時に適用",
            "\tEyedropperTool/rasterSampleSize\tBoolean\t1",
            "\tBucketTool/graphicsStyleExpanded\tBoolean\t1\t［スポイトツールオプション > スポイトの適用 > アピアランス］を展開しているかどうか。次回起動時に適用",
            "\tBucketTool/paint/fillExpanded\tBoolean\t1\t［スポイトツールオプション > スポイトの適用 > アピアランス > 塗り］を展開しているかどうか。次回起動時に適用",
            "\tBucketTool/paint/strokeExpanded\tBoolean\t1\t［スポイトツールオプション > スポイトの適用 > アピアランス > 線］を展開しているかどうか。次回起動時に適用",
            "\tBucketTool/paint/objectExpanded\tInteger\t1",
            "スポイトツールオプション > スポイトの適用 > アピアランス\tBucketTool/graphicsStyle\tBoolean\t0\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance\tBucketTool/graphicsStyle\tBoolean\t0\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 透明\tBucketTool/paint/object/transparency\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Transparency\tBucketTool/paint/object/transparency\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 塗り\tBucketTool/paint/fill\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Fill\tBucketTool/paint/fill\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 塗り > カラー\tBucketTool/paint/fill/color\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Fill > Color\tBucketTool/paint/fill/color\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 塗り > 透明\tBucketTool/paint/fill/transparency\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Fill > Transparency\tBucketTool/paint/fill/transparency\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 塗り > オーバープリント\tBucketTool/paint/fill/overprint\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Fill > Overprint\tBucketTool/paint/fill/overprint\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線\tBucketTool/paint/stroke\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke\tBucketTool/paint/stroke\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > カラー\tBucketTool/paint/stroke/color\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Color\tBucketTool/paint/stroke/color\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > 透明\tBucketTool/paint/stroke/transparency\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Transparency\tBucketTool/paint/stroke/transparency\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > オーバープリント\tBucketTool/paint/stroke/overprint\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Overprint\tBucketTool/paint/stroke/overprint\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > 線幅\tBucketTool/paint/stroke/weight\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Weight\tBucketTool/paint/stroke/weight\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > 線端の形状\tBucketTool/paint/stroke/cap\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Cap\tBucketTool/paint/stroke/cap\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > 角の形状\tBucketTool/paint/stroke/join\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Join\tBucketTool/paint/stroke/join\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > 角の比率\tBucketTool/paint/stroke/miter\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Miter limit\tBucketTool/paint/stroke/miter\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > アピアランス > 線 > 破線パターン\tBucketTool/paint/stroke/dash\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Appearance > Focal Stroke > Dash pattern\tBucketTool/paint/stroke/dash\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > 文字スタイル\tBucketTool/text/characterStyle\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Character Style\tBucketTool/text/characterStyle\tBoolean\t1\t次回起動時に適用",
            "スポイトツールオプション > スポイトの適用 > 段落スタイル\tBucketTool/text/paragraphStyle\tBoolean\t1\t次回起動時に適用",
            "Eyedropper Options > Eyedropper Applies > Paragraph Style\tBucketTool/text/paragraphStyle\tBoolean\t1\t次回起動時に適用",
            "\tBucketTool/rasterSampleSize\tInteger\t1",
            "編集 > カラーを編集 > オブジェクトを再配色… > 編集 > 起動時に「詳細オブジェクトを再配色」ダイアログを開く\tColorHarmony/AdvanceDialogDefaultOn\tBoolean\t0",
            "Edit > Edit Colors > Recolor Artwork… > Edit > Open Advance Recolor Artwork dialog on launch\tColorHarmony/AdvanceDialogDefaultOn\tBoolean\t0",
            "\tColorHarmony/VariationTypeColorGuide\tInteger\t0",
            "\tColorHarmony/ShowHideColorGuide\tInteger\t1",
            "\tColorHarmony/combineTintsOnExtraction\tInteger\t1",
            "\tColorHarmony/mapMethodApplyToAll\tInteger\t1",
            "\tColorHarmony/mapMethodPreserveSpots\tInteger\t1",
            "\tColorHarmony/SliderColorSpace\tInteger\t3",
            "\tColorHarmony/ShowAdjustSliders\tInteger\t0",
            "\tColorHarmony/ShowStorage\tInteger\t1",
            "編集 > カラーを編集 > オブジェクトを再配色… > 編集 > オブジェクトを再配色\tColorHarmony/HarmonyDialogRecolorArtCheckbox\tBoolean\t1",
            "Edit > Edit Colors > Recolor Artwork… > Edit > Recolor Art\tColorHarmony/HarmonyDialogRecolorArtCheckbox\tBoolean\t1",
            "\tGradient Annotator View Option\tInteger\t1",
            "\tBBRotateTolerance\tReal\t-1.0",
            "\tBBScaleTolerance\tReal\t-1.0",
            "表示 > バウンディングボックスを表示\tshowBoundingBox\tBoolean\t1",
            "View > Show Bounding Box\tshowBoundingBox\tBoolean\t1",
            "編集 > スペルチェック > 自動スペルチェック\tdynamicspelling\tBoolean\t0",
            "Edit > Spelling > Auto Spell Check\tdynamicspelling\tBoolean\t0",
            "表示 > コーナーウィジェットを表示\tliveCorners/showWidget\tBoolean\t1",
            "View > Show Corner Widget\tliveCorners/showWidget\tBoolean\t1",
            "シェイプ形成ツールオプション > 隙間の検出\tPlanar/MergeTool/Gap/Detect\tBoolean\t0\t次回起動時に適用",
            "Shape Builder Tool Options > Gap Detection\tPlanar/MergeTool/Gap/Detect\tBoolean\t0\t次回起動時に適用",
            "シェイプ形成ツールオプション > 隙間の検出 > 隙間の長さ «プリセット»\tPlanar/MergeTool/Gap/Type\tInteger\t0\t次回起動時に適用",
            "Shape Builder Tool Options > Gap Detection > Gap Length «Preset»\tPlanar/MergeTool/Gap/Type\tInteger\t0\t次回起動時に適用",
            "シェイプ形成ツールオプション > 隙間の検出 > 隙間の長さ «長さ»\tPlanar/MergeTool/Gap/Length\tReal\t3.0\t次回起動時に適用",
            "Shape Builder Tool Options > Gap Detection > Gap Length «Length»\tPlanar/MergeTool/Gap/Length\tReal\t3.0\t次回起動時に適用",
            "シェイプ形成ツールオプション > オプション > 塗りつぶされたオープンパスをクローズパスとして処理\tPlanar/MergeTool/ClosedOpenFillPaths\tBoolean\t1\t次回起動時に適用",
            "Shape Builder Tool Options > Options > Consider Open Filled Path as Closed\tPlanar/MergeTool/ClosedOpenFillPaths\tBoolean\t1\t次回起動時に適用",
            "シェイプ形成ツールオプション > オプション > 結合モードで線をクリックしてパスを分割\tPlanar/MergeTool/BreakEdge\tBoolean\t0\t次回起動時に適用",
            "Shape Builder Tool Options > Options > In Merge Mode, Clicking Stroke Splits the Path\tPlanar/MergeTool/BreakEdge\tBoolean\t0\t次回起動時に適用",
            "シェイプ形成ツールオプション > オプション > 次のカラーを利用\tPlanar/MergeTool/PaintFills\tInteger\t0\t次回起動時に適用",
            "Shape Builder Tool Options > Options > Pick Color From\tPlanar/MergeTool/PaintFills\tInteger\t0\t次回起動時に適用",
            "シェイプ形成ツールオプション > オプション > カーソルスウォッチプレビュー\tPlanar/MergeTool/PreviewSwatchCursor\tBoolean\t0\t次回起動時に適用",
            "Shape Builder Tool Options > Options > Cursor Swatch Preview\tPlanar/MergeTool/PreviewSwatchCursor\tBoolean\t0\t次回起動時に適用",
            "シェイプ形成ツールオプション > 強調表示 > 塗り\tPlanar/MergeTool/Highlight/Fill\tBoolean\t1\t次回起動時に適用",
            "Shape Builder Tool Options > Highlight > Fill\tPlanar/MergeTool/Highlight/Fill\tBoolean\t1\t次回起動時に適用",
            "シェイプ形成ツールオプション > 強調表示 > 編集可能なパスを強調表示\tPlanar/MergeTool/Highlight/Stroke\tBoolean\t1\t次回起動時に適用",
            "Shape Builder Tool Options > Highlight > Highlight Stroke when Editable\tPlanar/MergeTool/Highlight/Stroke\tBoolean\t1\t次回起動時に適用",
            "シェイプ形成ツールオプション > 強調表示 > カラー\tPlanar/MergeTool/Highlight/StrokeColorIndex\tInteger\t3\t次回起動時に適用",
            "Shape Builder Tool Options > Highlight > Color\tPlanar/MergeTool/Highlight/StrokeColorIndex\tInteger\t3\t次回起動時に適用",
            "シェイプ形成ツールオプション > 強調表示 > カラー «Red»\tPlanar/MergeTool/Highlight/StrokeColor/Red\tInteger\t65535\t次回起動時に適用",
            "Shape Builder Tool Options > Highlight > Color «Red»\tPlanar/MergeTool/Highlight/StrokeColor/Red\tInteger\t65535\t次回起動時に適用",
            "シェイプ形成ツールオプション > 強調表示 > カラー «Green»\tPlanar/MergeTool/Highlight/StrokeColor/Green\tInteger\t20224\t次回起動時に適用",
            "Shape Builder Tool Options > Highlight > Color «Green»\tPlanar/MergeTool/Highlight/StrokeColor/Green\tInteger\t20224\t次回起動時に適用",
            "シェイプ形成ツールオプション > 強調表示 > カラー «Blue»\tPlanar/MergeTool/Highlight/StrokeColor/Blue\tInteger\t20224\t次回起動時に適用",
            "Shape Builder Tool Options > Highlight > Color «Blue»\tPlanar/MergeTool/Highlight/StrokeColor/Blue\tInteger\t20224\t次回起動時に適用",
            "ライブペイントオプション > オプション > 塗りをペイント\tPlanar/Paintbucket/PaintFills\tBoolean\t1\t次回起動時に適用",
            "Live Paint Bucket Options > Options > Paint Fills\tPlanar/Paintbucket/PaintFills\tBoolean\t1\t次回起動時に適用",
            "ライブペイントオプション > オプション > 線をペイント\tPlanar/Paintbucket/PaintStrokes\tBoolean\t0\t次回起動時に適用",
            "Live Paint Bucket Options > Options > Paint Strokes\tPlanar/Paintbucket/PaintStrokes\tBoolean\t0\t次回起動時に適用",
            "ライブペイントオプション > オプション > カーソルスウォッチプレビュー\tPlanar/Paintbucket/CursorSwatchPrevew\tBoolean\t1\t次回起動時に適用",
            "Live Paint Bucket Options > Cursor Swatch Preview\tPlanar/Paintbucket/CursorSwatchPrevew\tBoolean\t1\t次回起動時に適用",
            "ライブペイントオプション > 強調表示\tPlanar/Paintbucket/Highlight\tBoolean\t1\t次回起動時に適用",
            "Live Paint Bucket Options > Highlight\tPlanar/Paintbucket/Highlight\tBoolean\t1\t次回起動時に適用",
            "ライブペイントオプション > 強調表示 > カラー «Red»\tPlanar/Paintbucket/Highlight/Color/red\tInteger\t65535\t次回起動時に適用",
            "Live Paint Bucket Options > Highlight > Color «Red»\tPlanar/Paintbucket/Highlight/Color/red\tInteger\t65535\t次回起動時に適用",
            "ライブペイントオプション > 強調表示 > カラー «Green»\tPlanar/Paintbucket/Highlight/Color/green\tInteger\t20224\t次回起動時に適用",
            "Live Paint Bucket Options > Highlight > Color «Green»\tPlanar/Paintbucket/Highlight/Color/green\tInteger\t20224\t次回起動時に適用",
            "ライブペイントオプション > 強調表示 > カラー «Blue»\tPlanar/Paintbucket/Highlight/Color/blue\tInteger\t20224\t次回起動時に適用",
            "Live Paint Bucket Options > Highlight > Color «Blue»\tPlanar/Paintbucket/Highlight/Color/blue\tInteger\t20224\t次回起動時に適用",
            "ライブペイントオプション > 強調表示 > 幅\tPlanar/Paintbucket/Highlight/Stroke\tInteger\t4\t次回起動時に適用",
            "Live Paint Bucket Options > Highlight > Width\tPlanar/Paintbucket/Highlight/Stroke\tInteger\t4\t次回起動時に適用",
            "ライブペイント選択オプション > オプション > 塗りを選択\tPlanar/FaceSelect/SelectFills\tBoolean\t1\t次回起動時に適用",
            "Live Paint Selection Options > Options > Select Fills\tPlanar/FaceSelect/SelectFills\tBoolean\t1\t次回起動時に適用",
            "ライブペイント選択オプション > オプション > 線を選択\tPlanar/FaceSelect/SelectStrokes\tBoolean\t1\t次回起動時に適用",
            "Live Paint Selection Options > Options > Select Strokes\tPlanar/FaceSelect/SelectStrokes\tBoolean\t1\t次回起動時に適用",
            "ライブペイント選択オプション > 強調表示\tPlanar/FaceSelect/Highlight\tBoolean\t1\t次回起動時に適用",
            "Live Paint Selection Options > Highlight\tPlanar/FaceSelect/Highlight\tBoolean\t1\t次回起動時に適用",
            "ライブペイント選択オプション > 強調表示 > カラー «Red»\tPlanar/FaceSelect/Highlight/Color/red\tInteger\t65535\t次回起動時に適用",
            "Live Paint Selection Options > Highlight > Color «Red»\tPlanar/FaceSelect/Highlight/Color/red\tInteger\t65535\t次回起動時に適用",
            "ライブペイント選択オプション > 強調表示 > カラー «Green»\tPlanar/FaceSelect/Highlight/Color/green\tInteger\t20224\t次回起動時に適用",
            "Live Paint Selection Options > Highlight > Color «Green»\tPlanar/FaceSelect/Highlight/Color/green\tInteger\t20224\t次回起動時に適用",
            "ライブペイント選択オプション > 強調表示 > カラー «Blue»\tPlanar/FaceSelect/Highlight/Color/blue\tInteger\t20224\t次回起動時に適用",
            "Live Paint Selection Options > Highlight > Color «Blue»\tPlanar/FaceSelect/Highlight/Color/blue\tInteger\t20224\t次回起動時に適用",
            "ライブペイント選択オプション > 強調表示 > 幅\tPlanar/FaceSelect/Highlight/Stroke\tInteger\t4\t次回起動時に適用",
            "Live Paint Selection Options > Highlight > Width\tPlanar/FaceSelect/Highlight/Stroke\tInteger\t4\t次回起動時に適用",
            "\tPlanar/Resolution\tInteger\t100",
            "\tPlanar/ArtboardGapPreview\tInteger\t0",
            "オブジェクト > ライブペイント > 隙間オプション… > 隙間の検出\tPlanar/GapDetection\tBoolean\t1\t次回起動時に取得／適用",
            "Object > Live Paint > Gap Options… > Gap Detection\tPlanar/GapDetection\tBoolean\t1\t次回起動時に取得／適用",
            "オブジェクト > ライブペイント > 隙間オプション… > 隙間の検出 > 塗りの許容サイズ > カスタムの隙間\tPlanar/GapDetection/Custom\tBoolean\t0\t次回起動時に取得／適用",
            "Object > Live Paint > Gap Options… > Gap Detection > Custom Gaps\tPlanar/GapDetection/Custom\tBoolean\t0\t次回起動時に取得／適用",
            "オブジェクト > ライブペイント > 隙間オプション… > 隙間の検出 > 塗りの許容サイズ «隙間»\tPlanar/GapDetection/Size\tReal\t3.0\t次回起動時に取得／適用",
            "Object > Live Paint > Gap Options… > Gap Detection > Paint stops at «Size»\tPlanar/GapDetection/Size\tReal\t3.0\t次回起動時に取得／適用",
            "オブジェクト > ライブペイント > 隙間オプション… > 隙間の検出 > 隙間のプレビューカラー «Red»\tPlanar/GapDetection/GapColor/red\tInteger\t65535\t次回起動時に取得／適用",
            "Object > Live Paint > Gap Options… > Gap Detection > Gap Preview Color «Red»\tPlanar/GapDetection/GapColor/red\tInteger\t65535\t次回起動時に取得／適用",
            "オブジェクト > ライブペイント > 隙間オプション… > 隙間の検出 > 隙間のプレビューカラー «Green»\tPlanar/GapDetection/GapColor/green\tInteger\t20224\t次回起動時に取得／適用",
            "Object > Live Paint > Gap Options… > Gap Detection > Gap Preview Color «Green»\tPlanar/GapDetection/GapColor/green\tInteger\t20224\t次回起動時に取得／適用",
            "オブジェクト > ライブペイント > 隙間オプション… > 隙間の検出 > 隙間のプレビューカラー «Blue»\tPlanar/GapDetection/GapColor/blue\tInteger\t20224\t次回起動時に取得／適用",
            "Object > Live Paint > Gap Options… > Gap Detection > Gap Preview Color «Blue»\tPlanar/GapDetection/GapColor/blue\tInteger\t20224\t次回起動時に取得／適用",
            "\tPlanar/GapDetection/GapColor/Color/red\tInteger\t65535",
            "\tPlanar/GapDetection/GapColor/Color/green\tInteger\t20224",
            "\tPlanar/GapDetection/GapColor/Color/blue\tInteger\t20224",
            "オブジェクト > ライブペイント > 隙間オプション… > プレビュー\tPlanar/GapDetection/PreviewOn\tBoolean\t1\t次回起動時に取得／適用",
            "Object > Live Paint > Gap Options… > Preview\tPlanar/GapDetection/PreviewOn\tBoolean\t1\t次回起動時に取得／適用",
            "\thighlightLockedObjects\tBoolean\t0",
            "\tcheckRequiredPlugins\tInteger\t1",
            "\tBIBCacheMaxSizeInBytes\tInteger\t41943040",
            "\tBIBCacheSizeInPercentOfPhysicalRam\tReal\t2.0",
            "\tpluginCacheCheckCreateDate\tInteger\t3777586067",
            "\tpluginCacheCheckModDate\tInteger\t3782419273",
            "\tloadPluginsFromOutsideAppPackage\tInteger\t1",
            "\tmemory/memoryPercentage\tInteger\t0",
            "\tmemory/physicalRAMSize\tInteger\t131072",
            "\tmemory/usePhysicalRAMSize\tInteger\t0",
            "\tlayout/0/ApplicationBarOption\tInteger\t0",
            "\tlayout/0/ApplicationBarVisible\tInteger\t0",
            "\tuseProcessorSpecificCode\tInteger\t1",
            "\tEnableInternalMemoryPool\tInteger\t0",
            "\teditableGuides\tInteger\t0",
            "\tplugin/Cleanup_Preference/Cleanup_EmptyText\tInteger\t1",
            "\tplugin/Cleanup_Preference/Cleanup_NoPaintObjects\tInteger\t1",
            "\tplugin/Cleanup_Preference/Cleanup_StrayPoints\tInteger\t1",
            "\tplugin/PDF File Format UI/PDFImportFTUECount\tInteger\t3",
            "アクションパネル > バッチ… > ソース\tplugin/BatchAction/sourceMode\tInteger\t0",
            "Actions Panel > Batch… > Source\tplugin/BatchAction/sourceMode\tInteger\t0",
            "アクションパネル > バッチ… > ソース > フォルダー\tplugin/BatchAction/openDir\tString\t/Users/«username»/Desktop",
            "Actions Panel > Batch… > Source > Folder\tplugin/BatchAction/openDir\tString\t/Users/«username»/Desktop",
            "アクションパネル > バッチ… > ソース > フォルダー > アクションの「開く」コマンドを無視\tplugin/BatchAction/overrideOpen\tBoolean\t0",
            "Actions Panel > Batch… > Source > Folder > Override Action “Open” Commands\tplugin/BatchAction/overrideOpen\tBoolean\t0",
            "アクションパネル > バッチ… > ソース > フォルダー > サブディレクトリをすべて含める\tplugin/BatchAction/includeSubDir\tBoolean\t0",
            "Actions Panel > Batch… > Source > Folder > Include All Subdirectories\tplugin/BatchAction/includeSubDir\tBoolean\t0",
            "アクションパネル > バッチ… > ソース > フォルダー > オブジェクトを同じドキュメントに配置\tplugin/BatchAction/placeObject\tBoolean\t0",
            "Actions Panel > Batch… > Source > Folder > Place Object in Same Document\tplugin/BatchAction/placeObject\tBoolean\t0",
            "アクションパネル > バッチ… > 保存先\tplugin/BatchAction/saveMode\tInteger\t2",
            "Actions Panel > Batch… > Destination\tplugin/BatchAction/saveMode\tInteger\t2",
            "アクションパネル > バッチ… > 保存先 > フォルダー (保存)\tplugin/BatchAction/saveDir\tString\t/Users/«username»/Desktop",
            "Actions Panel > Batch… > Destination > Folder (Save)\tplugin/BatchAction/saveDir\tString\t/Users/«username»/Desktop",
            "アクションパネル > バッチ… > 保存先 > アクションの 「保存」コマンドを無視\tplugin/BatchAction/overrideSave\tBoolean\t1",
            "Actions Panel > Batch… > Destination > Override Action “Save” Commands\tplugin/BatchAction/overrideSave\tBoolean\t1",
            "アクションパネル > バッチ… > 保存先 > フォルダー (書き出し)\tplugin/BatchAction/exportDir\tString\t/Users/«username»/Desktop",
            "Actions Panel > Batch… > Destination > Folder (Export)\tplugin/BatchAction/exportDir\tString\t/Users/«username»/Desktop",
            "アクションパネル > バッチ… > 保存先 > アクションの「書き出し」コマンドを無視\tplugin/BatchAction/overrideExport\tBoolean\t0",
            "Actions Panel > Batch… > Destination > Override Action “Export” Commands\tplugin/BatchAction/overrideExport\tBoolean\t0",
            "アクションパネル > バッチ… > 保存先 > ファイル名\tplugin/BatchAction/fileNameAlgorithm\tInteger\t0",
            "Actions Panel > Batch… > Destination > File Name\tplugin/BatchAction/fileNameAlgorithm\tInteger\t0",
            "アクションパネル > バッチ… > エラー\tplugin/BatchAction/stopForErr\tInteger\t0",
            "Actions Panel > Batch… > Errors\tplugin/BatchAction/stopForErr\tInteger\t0",
            "アクションパネル > バッチ… > エラー > 別名で保存…\tplugin/BatchAction/errFile\tString\t/Users/«username»/Desktop/errorLog.txt",
            "Actions Panel > Batch… > Errors > Save As…\tplugin/BatchAction/errFile\tString\t/Users/«username»/Desktop/errorLog.txt",
            "アートボードパネル > すべてのアートボードを再配置… > オブジェクトと一緒に移動\tplugin/ArtboardRearrange/MoveArtwork\tBoolean\t1",
            "Artboards Panel > Rearrange All Artboards… > Move Artwork with Artboard\tplugin/ArtboardRearrange/MoveArtwork\tBoolean\t1",
            "アートボードパネル > すべてのアートボードを再配置… > レイアウト\tplugin/ArtboardRearrange/LayoutChosen\tInteger\t2",
            "Artboards Panel > Rearrange All Artboards… > Layout\tplugin/ArtboardRearrange/LayoutChosen\tInteger\t2",
            "アートボードパネル > すべてのアートボードを再配置… > レイアウトの順序\tplugin/ArtboardRearrange/IsLayoutRToL\tBoolean\t0\t{true: <—RtoL, false: —>LtoR}",
            "Artboards Panel > Rearrange All Artboards… > Layout order\tplugin/ArtboardRearrange/IsLayoutRToL\tBoolean\t0\t{true: <—RtoL, false: —>LtoR}",
            "アートボードパネル > すべてのアートボードを再配置… > 横列数 | 縦列数\tplugin/ArtboardRearrange/RowColumnDivision\tInteger\t1",
            "Artboards Panel > Rearrange All Artboards… > Columns | Rows\tplugin/ArtboardRearrange/RowColumnDivision\tInteger\t1",
            "アートボードパネル > すべてのアートボードを再配置… > 間隔\tplugin/ArtboardRearrange/ArtboardSpacing\tReal\t100",
            "Artboards Panel > Rearrange All Artboards… > Spacing\tplugin/ArtboardRearrange/ArtboardSpacing\tReal\t100",
            "\tplugin/AlignArtBoard/ArtboardPartialPathSelected\tInteger\t0",
            "\tplugin/AlignArtBoard/ArtBoardKeyObjectUserSelection\tInteger\t0",
            "\tplugin/AlignArtBoard/ArtboardToolSelected\tInteger\t1",
            "\tplugin/AlignArtBoard/ArtBoardDefaultSelection\tInteger\t2",
            "\tplugin/AlignArtBoard/ArtBoardUserSelection\tInteger\t0",
            "カラーピッカー > Web セーフカラーのみに制限\tplugin/ColorPickerDlg/WebSafe\tBoolean\t0",
            "Color Picker > Only Web Colors\tplugin/ColorPickerDlg/WebSafe\tBoolean\t0",
            "\tplugin/ColorPickerDlg/ItemID\tInteger\t0",
            "\tplugin/svgOMGOptionDlg/O_MinifySVG\tInteger\t1",
            "\tplugin/svgOMGOptionDlg/O_ResponsiveSVG\tInteger\t1",
            "\tplugin/svgOMGOptionDlg/O_IncludeColorSwatch\tInteger\t0",
            "\tplugin/svgOMGOptionDlg/O_EmbedRasterLoc\tInteger\t3",
            "\tplugin/svgOMGOptionDlg/O_SVGProfile\tInteger\t0",
            "\tplugin/svgOMGOptionDlg/O_IdType\tInteger\t0",
            "\tplugin/svgOMGOptionDlg/O_SVFontType\tInteger\t1",
            "\tplugin/SVGFormatOMGOptions/StyleType\tInteger\t3",
            "\tplugin/SVGFormatOMGOptions/CoordinatePrecision\tInteger\t2",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > 画像 > カラーモード\tplugin/JPEGFormat/ColorModel\tInteger\t2\t書類のカラーモードに依存し，セットしても適用されない",
            "File > Export > Export As… > JPEG Options > Image > Color Model\tplugin/JPEGFormat/ColorModel\tInteger\t2\t書類のカラーモードに依存し，セットしても適用されない",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > 画像 > 画質 «数値»\tplugin/JPEGFormat/Image Quality\tInteger\t10",
            "File > Export > Export As… > JPEG Options > Image > Quality «Integer»\tplugin/JPEGFormat/Image Quality\tInteger\t10",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > オプション > 圧縮形式\tplugin/JPEGFormat/Compression\tInteger\t1",
            "File > Export > Export As… > JPEG Options > Options > Compression Method\tplugin/JPEGFormat/Compression\tInteger\t1",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > オプション > 圧縮形式 > プログレッシブ > スキャン\tplugin/JPEGFormat/NumScans\tInteger\t3\t3〜5",
            "File > Export > Export As… > JPEG Options > Options > Compression Method > Progressive > Scans\tplugin/JPEGFormat/NumScans\tInteger\t3\t3〜5",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > オプション > 解像度\tplugin/JPEGFormat/DPI\tReal\t150",
            "File > Export > Export As… > JPEG Options > Options > Resolution\tplugin/JPEGFormat/DPI\tReal\t150",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > オプション > アンチエイリアス\tplugin/JPEGFormat/AntiAlias\tInteger\t3",
            "File > Export > Export As… > JPEG Options > Options > Anti-aliasing\tplugin/JPEGFormat/AntiAlias\tInteger\t3",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > オプション > イメージマップ «真偽値»\tplugin/JPEGFormat/DoImageMap\tBoolean\t0",
            "File > Export > Export As… > JPEG Options > Options > Imagemap «Checked»\tplugin/JPEGFormat/DoImageMap\tBoolean\t0",
            "ファイル > 書き出し > 書き出し形式… > JPEG オプション > オプション > イメージマップ «種類»\tplugin/JPEGFormat/ImageMapType\tInteger\t1",
            "File > Export > Export As… > JPEG Options > Options > Imagemap «Type»\tplugin/JPEGFormat/ImageMapType\tInteger\t1",
            "\tplugin/JPEGFormat/EmbedState\tInteger\t0",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 100 > オプション > 圧縮形式\tplugin/JPEGFormat100/Compression\tInteger\t1",
            "File > Export > Export for Screens… > Format Settings > JPG 100 > Options > Compression Method\tplugin/JPEGFormat100/Compression\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 100 > オプション > 圧縮形式 > プログレッシブ > スキャン\tplugin/JPEGFormat100/NumScans\tInteger\t3\t3〜5",
            "File > Export > Export for Screens… > Format Settings > JPG 100 > Options > Compression Method > Progressive > Scans\tplugin/JPEGFormat100/NumScans\tInteger\t3\t3〜5",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 100 > オプション > アンチエイリアス\tplugin/JPEGFormat100/AntiAlias\tInteger\t3",
            "File > Export > Export for Screens… > Format Settings > JPG 100 > Options > Anti-aliasing\tplugin/JPEGFormat100/AntiAlias\tInteger\t3",
            "\tplugin/JPEGFormat100/Image Quality\tInteger\t10",
            "\tplugin/JPEGFormat100/DPI\tReal\t144",
            "\tplugin/JPEGFormat100/DoImageMap\tInteger\t0",
            "\tplugin/JPEGFormat100/ImageMapType\tInteger\t1",
            "\tplugin/JPEGFormat100/EmbedState\tInteger\t1",
            "\tplugin/JPEGFormat100/ColorModel\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 80 > オプション > 圧縮形式\tplugin/JPEGFormat80/Compression\tInteger\t1",
            "File > Export > Export for Screens… > Format Settings > JPG 80 > Options > Compression Method\tplugin/JPEGFormat80/Compression\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 80 > オプション > 圧縮形式 > プログレッシブ > スキャン\tplugin/JPEGFormat80/NumScans\tInteger\t3",
            "File > Export > Export for Screens… > Format Settings > JPG 80 > Options > Compression Method > Progressive > Scans\tplugin/JPEGFormat80/NumScans\tInteger\t3",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 80 > オプション > アンチエイリアス\tplugin/JPEGFormat80/AntiAlias\tInteger\t3",
            "File > Export > Export for Screens… > Format Settings > JPG 80 > Options > Anti-aliasing\tplugin/JPEGFormat80/AntiAlias\tInteger\t3",
            "\tplugin/JPEGFormat80/Image Quality\tInteger\t8",
            "\tplugin/JPEGFormat80/DPI\tReal\t72",
            "\tplugin/JPEGFormat80/DoImageMap\tInteger\t0",
            "\tplugin/JPEGFormat80/ImageMapType\tInteger\t1",
            "\tplugin/JPEGFormat80/EmbedState\tInteger\t1",
            "\tplugin/JPEGFormat80/ColorModel\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 50 > オプション > 圧縮形式\tplugin/JPEGFormat50/Compression\tInteger\t1",
            "File > Export > Export for Screens… > Format Settings > JPG 50 > Options > Compression Method\tplugin/JPEGFormat50/Compression\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 50 > オプション > 圧縮形式 > プログレッシブ > スキャン\tplugin/JPEGFormat50/NumScans\tInteger\t3",
            "File > Export > Export for Screens… > Format Settings > JPG 50 > Options > Compression Method > Progressive > Scans\tplugin/JPEGFormat50/NumScans\tInteger\t3",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 50 > オプション > アンチエイリアス\tplugin/JPEGFormat50/AntiAlias\tInteger\t3",
            "File > Export > Export for Screens… > Format Settings > JPG 50 > Options > Anti-aliasing\tplugin/JPEGFormat50/AntiAlias\tInteger\t3",
            "\tplugin/JPEGFormat50/Image Quality\tInteger\t5",
            "\tplugin/JPEGFormat50/DPI\tReal\t72",
            "\tplugin/JPEGFormat50/DoImageMap\tInteger\t0",
            "\tplugin/JPEGFormat50/ImageMapType\tInteger\t1",
            "\tplugin/JPEGFormat50/EmbedState\tInteger\t1",
            "\tplugin/JPEGFormat50/ColorModel\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 20 > オプション > 圧縮形式\tplugin/JPEGFormat20/Compression\tInteger\t1",
            "File > Export > Export for Screens… > Format Settings > JPG 20 > Options > Compression Method\tplugin/JPEGFormat20/Compression\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 20 > オプション > 圧縮形式 > プログレッシブ > スキャン\tplugin/JPEGFormat20/NumScans\tInteger\t3",
            "File > Export > Export for Screens… > Format Settings > JPG 20 > Options > Compression Method > Progressive > Scans\tplugin/JPEGFormat20/NumScans\tInteger\t3",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > JPG 20 > オプション > アンチエイリアス\tplugin/JPEGFormat20/AntiAlias\tInteger\t3",
            "File > Export > Export for Screens… > Format Settings > JPG 20 > Options > Anti-aliasing\tplugin/JPEGFormat20/AntiAlias\tInteger\t3",
            "\tplugin/JPEGFormat20/Image Quality\tInteger\t2",
            "\tplugin/JPEGFormat20/DPI\tReal\t72",
            "\tplugin/JPEGFormat20/DoImageMap\tInteger\t0",
            "\tplugin/JPEGFormat20/ImageMapType\tInteger\t1",
            "\tplugin/JPEGFormat20/EmbedState\tInteger\t1",
            "\tplugin/JPEGFormat20/ColorModel\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PNG 8 > カラー\tplugin/PNG8/NumberOfColors\tInteger\t256",
            "File > Export > Export for Screens… > Format Settings > PNG 8 > Colors\tplugin/PNG8/NumberOfColors\tInteger\t256",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PNG 8 > 透明度\tplugin/PNG8/BackgroundTransparent\tBoolean\t1",
            "File > Export > Export for Screens… > Format Settings > PNG 8 > Transparent\tplugin/PNG8/BackgroundTransparent\tBoolean\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PNG 8 > インターレース\tplugin/PNG8/Interlaced\tBoolean\t0",
            "File > Export > Export for Screens… > Format Settings > PNG 8 > Interlace\tplugin/PNG8/Interlaced\tBoolean\t0",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PNG 8 > アンチエイリアス\tplugin/PNG8/AntiAlias\tInteger\t2",
            "File > Export > Export for Screens… > Format Settings > PNG 8 > Anti-aliasing\tplugin/PNG8/AntiAlias\tInteger\t2",
            "\tplugin/PNG8/ResolutionR\tReal\t72",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PNG 8 > マット «Red»\tplugin/PNG8/Background/red\tInteger\t255",
            "File > Export > Export for Screens… > Format Settings > PNG 8 > Matte «Red»\tplugin/PNG8/Background/red\tInteger\t255",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PNG 8 > マット «Green»\tplugin/PNG8/Background/green\tInteger\t255",
            "File > Export > Export for Screens… > Format Settings > PNG 8 > Matte «Green»\tplugin/PNG8/Background/green\tInteger\t255",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PNG 8 > マット «Blue»\tplugin/PNG8/Background/blue\tInteger\t255",
            "File > Export > Export for Screens… > Format Settings > PNG 8 > Matte «Blue»\tplugin/PNG8/Background/blue\tInteger\t255",
            "\tplugin/Adobe Vector Sculpting UI/ToolIntroFTUEShownCount\tInteger\t3",
            "\tplugin/Adobe Vector Sculpting UI/ToolIntroFTUEShown\tInteger\t0",
            "\tplugin/OnBoarding/PantoneLibOnBoardingShownCount\tInteger\t0",
            "\tplugin/OnBoarding/dontShowAgainPantoneLibOnBoarding\tInteger\t0",
            "ブラシパネル > ブラシオプション… > アートブラシオプション > プレビュー\tplugin/ArtBrushOptions/PreviewChanges\tBoolean\t0",
            "Brushes Panel > Brush Options… > Art Brush Options > Preview\tplugin/ArtBrushOptions/PreviewChanges\tBoolean\t0",
            "ブラシパネル > ブラシオプション… > パターンブラシオプション > プレビュー\tplugin/PatternBrushOptions/PreviewChanges\tBoolean\t0",
            "Brushes Panel > Brush Options… > Pattern Brush Options > Preview\tplugin/PatternBrushOptions/PreviewChanges\tBoolean\t0",
            "効果 > 3D とマテリアル > 3D (クラシック) > 回転体 (クラシック)… > プレビュー\tplugin/PreviewPref/3DRevolvePreviewPref\tBoolean\t1",
            "Effect > 3D and Materials > 3D (Classic) > Revolve (Classic)… > Preview\tplugin/PreviewPref/3DRevolvePreviewPref\tBoolean\t1",
            "オブジェクト > エンベロープ > エンベロープオプション… > プレビュー\tplugin/PreviewPref/EnvelopDistortPreviewPref\tBoolean\t1",
            "Object > Envelope Distort > Envelope Options… > Preview\tplugin/PreviewPref/EnvelopDistortPreviewPref\tBoolean\t1",
            "\tplugin/ShaperTool/MinAspectForSquareCircleSnapping\tReal\t0.75",
            "\tplugin/ShaperTool/snapShaperShapesTo45\tInteger\t1",
            "\tplugin/shaperTool/alwaysUseShaperDefaults\tInteger\t1",
            "\tplugin/VariablesPaletteUI/VariablesPaletteCoachOnboardingShown\tInteger\t0",
            "\tplugin/VariablesPaletteUI/VariablesPaletteCoachShownCount\tInteger\t1",
            "効果 > スタイライズ > 光彩 (外側)… > プレビュー\tplugin/PreviewPref/DropShadowUIOuterGlowPreviewPref\tBoolean\t1",
            "Effect > Stylize > Outer Glow… > Preview\tplugin/PreviewPref/DropShadowUIOuterGlowPreviewPref\tBoolean\t1",
            "オブジェクト > 変形 > リフレクト… > プレビュー\tplugin/PreviewPref/ReflectOptionPreviewPref\tBoolean\t1",
            "Object > Transform > Reflect… > Preview\tplugin/PreviewPref/ReflectOptionPreviewPref\tBoolean\t1",
            "オブジェクト > 変形 > 回転… > プレビュー\tplugin/PreviewPref/RotateOptionPreviewPref\tBoolean\t1",
            "Object > Transform > Rotate… > Preview\tplugin/PreviewPref/RotateOptionPreviewPref\tBoolean\t1",
            "オブジェクト > 変形 > 移動… > プレビュー\tplugin/PreviewPref/MoveOptionPreviewPref\tBoolean\t1",
            "Object > Transform > Move… > Preview\tplugin/PreviewPref/MoveOptionPreviewPref\tBoolean\t1",
            "効果 > 3D とマテリアル > 3D (クラシック) > 押し出しとベベル (クラシック)… > プレビュー\tplugin/PreviewPref/3DExtrude&BevelPreviewPref\tBoolean\t1",
            "Effect > 3D and Materials > 3D (Classic) > Extrude & Bevel (Classic)… > Preview\tplugin/PreviewPref/3DExtrude&BevelPreviewPref\tBoolean\t1",
            "効果 > スタイライズ > 光彩 (内側)… > プレビュー\tplugin/PreviewPref/DropShadowUIInnerGlowPreviewPref\tBoolean\t1",
            "Effect > Stylize > Inner Glow… > Preview\tplugin/PreviewPref/DropShadowUIInnerGlowPreviewPref\tBoolean\t1",
            "効果 > 形状に変換 > 長方形… > プレビュー\tplugin/PreviewPref/ShapeEffectRectanglePreviewPref\tBoolean\t1",
            "Effect > Convert to Shape > Rectangle… > Preview\tplugin/PreviewPref/ShapeEffectRectanglePreviewPref\tBoolean\t1",
            "効果 > パスの変形 > 変形… > プレビュー\tplugin/PreviewPref/TransformPreviewPref\tBoolean\t1\tオブジェクト > 変形 > 個別に変形… > プレビュー\\nObject > Transform > Transform Each… > Preview",
            "Effect > Distort & Transform > Transform… > Preview\tplugin/PreviewPref/TransformPreviewPref\tBoolean\t1\tオブジェクト > 変形 > 個別に変形… > プレビュー\\nObject > Transform > Transform Each… > Preview",
            "効果 > パスの変形 > ジグザグ… > プレビュー\tplugin/PreviewPref/ZigZagPreviewPref\tBoolean\t1",
            "Effect > Distort & Transform > Zig Zag… > Preview\tplugin/PreviewPref/ZigZagPreviewPref\tBoolean\t1",
            "効果 > ワープ > ワープオプション > プレビュー\tplugin/PreviewPref/WarpOptionsPreviewPref\tBoolean\t1",
            "Effect > Warp > Warp Options > Preview\tplugin/PreviewPref/WarpOptionsPreviewPref\tBoolean\t1",
            "効果 > パス > パスのオフセット… > プレビュー\tplugin/PreviewPref/OffsetPathPreviewPref\tBoolean\t1",
            "Effect > Path > Offset Path… > Preview\tplugin/PreviewPref/OffsetPathPreviewPref\tBoolean\t1",
            "効果 > スタイライズ > 角を丸くする… > プレビュー\tplugin/PreviewPref/RoundCornersPreviewPref\tBoolean\t1",
            "Effect > Stylize > Round Corners… > Preview\tplugin/PreviewPref/RoundCornersPreviewPref\tBoolean\t1",
            "効果 > パスの変形 > ラフ… > プレビュー\tplugin/PreviewPref/RoughenPreviewPref\tBoolean\t1",
            "Effect > Distort & Transform > Roughen… > Preview\tplugin/PreviewPref/RoughenPreviewPref\tBoolean\t1",
            "効果 > ぼかし > ぼかし (ガウス)… > プレビュー\tplugin/PreviewPref/GaussianBlurPreviewPref\tBoolean\t1",
            "Effect > Blur > Gaussian Blur… > Preview\tplugin/PreviewPref/GaussianBlurPreviewPref\tBoolean\t1",
            "効果 > スタイライズ > ぼかし… > プレビュー\tplugin/PreviewPref/FeatherEffectPreviewPref\tBoolean\t1",
            "Effect > Stylize > Feather… > Preview\tplugin/PreviewPref/FeatherEffectPreviewPref\tBoolean\t1",
            "効果 > スタイライズ > ドロップシャドウ… > プレビュー\tplugin/PreviewPref/DropShadowUIDropShadowPreviewPref\tBoolean\t1",
            "Effect > Stylize > Drop Shadow… > Preview\tplugin/PreviewPref/DropShadowUIDropShadowPreviewPref\tBoolean\t1",
            "オブジェクト > 変形 > 拡大・縮小… > プレビュー\tplugin/PreviewPref/ScaleOptionPreviewPref\tBoolean\t1",
            "Object > Transform > Scale… > Preview\tplugin/PreviewPref/ScaleOptionPreviewPref\tBoolean\t1",
            "オブジェクト > ブレンド > ブレンドオプション… > プレビュー\tplugin/PreviewPref/BlendOptionPreviewPref\tBoolean\t1",
            "Object > Blend > Blend Options… > Preview\tplugin/PreviewPref/BlendOptionPreviewPref\tBoolean\t1",
            "効果 > SVG フィルター > SVG フィルターを適用… > プレビュー\tplugin/PreviewPref/SVGFilterPreviewPref\tBoolean\t1",
            "Effect > SVG Filters > Apply SVG Filter… > Preview\tplugin/PreviewPref/SVGFilterPreviewPref\tBoolean\t1",
            "効果 > 3D とマテリアル > 3D (クラシック) > 回転 (クラシック)… > プレビュー\tplugin/PreviewPref/3DRotatePreviewPref\tBoolean\t1",
            "Effect > 3D and Materials > 3D (Classic) > Rotate (Classic)… > Preview\tplugin/PreviewPref/3DRotatePreviewPref\tBoolean\t1",
            "効果 > 形状に変換 > 角丸長方形… > プレビュー\tplugin/PreviewPref/ShapeEffectRoundedRectanglePreviewPref\tBoolean\t1",
            "Effect > Convert to Shape > Rounded Rectangle… > Preview\tplugin/PreviewPref/ShapeEffectRoundedRectanglePreviewPref\tBoolean\t1",
            "書式 > エリア内文字オプション… > プレビュー\tAreatypeOptionsPreview\tBoolean\t1",
            "Type > Area Type Options… > Preview\tAreatypeOptionsPreview\tBoolean\t1",
            "効果 > パスファインダー > パスファインダーオプション > プレビュー\tplugin/PathfinderEffectOptions/PreviewChanges\tBoolean\t1",
            "Effect > Pathfinder > Pathfinder Options > Preview\tplugin/PathfinderEffectOptions/PreviewChanges\tBoolean\t1",
            "ブラシパネル > ブラシオプション… > カリグラフィブラシオプション > プレビュー\tplugin/CalligraphicBrushOptions/PreviewChanges\tBoolean\t0",
            "Brushes Panel > Brush Options… > Calligraphic Brush Options > Preview\tplugin/CalligraphicBrushOptions/PreviewChanges\tBoolean\t0",
            "ブラシパネル > ブラシオプション… > 散布ブラシオプション > プレビュー\tplugin/ScatterBrushOptions/PreviewChanges\tBoolean\t1",
            "Brushes Panel > Brush Options… > Scatter Brush Options > Preview\tplugin/ScatterBrushOptions/PreviewChanges\tBoolean\t1",
            "オブジェクト > エンベロープ > ワープで作成… > プレビュー\tplugin/EnvelopeOptions/WarpShowPreview\tBoolean\t1",
            "Object > Envelope Distort > Make with Warp > Preview\tplugin/EnvelopeOptions/WarpShowPreview\tBoolean\t1",
            "オブジェクト > エンベロープ > エンベロープオプション… > アピアランスを変形\tplugin/EnvelopeOptions/ExpandAppearance\tBoolean\t1",
            "Object > Envelope Distort > Envelope Options… > Distort Appearance\tplugin/EnvelopeOptions/ExpandAppearance\tBoolean\t1",
            "オブジェクト > エンベロープ > エンベロープオプション… > パターンを変形\tplugin/EnvelopeOptions/ExpandPattern\tBoolean\t0",
            "Object > Envelope Distort > Envelope Options… > Distort Pattern Fills\tplugin/EnvelopeOptions/ExpandPattern\tBoolean\t0",
            "オブジェクト > エンベロープ > エンベロープオプション… > 線形グラデーションの塗りを変形\tplugin/EnvelopeOptions/ExpandGrdient\tBoolean\t0",
            "Object > Envelope Distort > Envelope Options… > Distort Linear Gradient Fills\tplugin/EnvelopeOptions/ExpandGrdient\tBoolean\t0",
            "オブジェクト > エンベロープ > エンベロープオプション… > 精度\tplugin/EnvelopeOptions/DeformFidelity\tInteger\t50",
            "Object > Envelope Distort > Envelope Options… > Fidelity\tplugin/EnvelopeOptions/DeformFidelity\tInteger\t50",
            "オブジェクト > エンベロープ > エンベロープオプション… > ラスタライズ > アンチエイリアス\tplugin/EnvelopeOptions/AntiAliasRasters\tBoolean\t1",
            "Object > Envelope Distort > Envelope Options… > Rasters > Anti-Alias\tplugin/EnvelopeOptions/AntiAliasRasters\tBoolean\t1",
            "オブジェクト > エンベロープ > エンベロープオプション… > ラスタライズ > シェイプの保持に使用\tplugin/EnvelopeOptions/AddAlphaToRasters\tInteger\t0",
            "Object > Envelope Distort > Envelope Options… > Rasters > Preserve Shape Using\tplugin/EnvelopeOptions/AddAlphaToRasters\tInteger\t0",
            "\tplugin/ai_navigation_get/tmpl\tInteger\t0",
            "\tplugin/ai_navigation_get/repl\tInteger\t0",
            "\tplugin/ai_navigation_get/link\tInteger\t1",
            "ブラシツールオプション > 精度\tplugin/GlobalBrushOptions/fairnessPreset\tInteger\t0\t0〜4",
            "Paintbrush Tool Options > Fidelity\tplugin/GlobalBrushOptions/fairnessPreset\tInteger\t0\t0〜4",
            "ブラシツールオプション > オプション > ブラシストロークに塗りを適用\tplugin/GlobalBrushOptions/fill\tBoolean\t1",
            "Paintbrush Tool Options > Options > Fill new brush strokes\tplugin/GlobalBrushOptions/fill\tBoolean\t1",
            "ブラシツールオプション > オプション > 選択を解除しない\tplugin/GlobalBrushOptions/select\tBoolean\t1",
            "Paintbrush Tool Options > Options > Keep Selected\tplugin/GlobalBrushOptions/select\tBoolean\t1",
            "ブラシツールオプション > オプション > 選択したパスを編集\tplugin/BRSPencilTool/edit_selection\tBoolean\t1",
            "Paintbrush Tool Options > Options > Edit Selected Paths\tplugin/BRSPencilTool/edit_selection\tBoolean\t1",
            "ブラシツールオプション > オプション > 範囲\tplugin/BRSTool30Mar9816Build/editing_distance\tReal\t12.0",
            "Paintbrush Tool Options > Options > Within\tplugin/BRSTool30Mar9816Build/editing_distance\tReal\t12.0",
            "\tplugin/BlobBrushOptions/deselectOtherArts\tInteger\t1",
            "塗りブラシツールオプション > 選択を解除しない\tplugin/BlobBrushOptions/keepSelected\tBoolean\t0",
            "Blob Brush Tool Options > Keep Selected\tplugin/BlobBrushOptions/keepSelected\tBoolean\t0",
            "塗りブラシツールオプション > 選択範囲のみ結合\tplugin/BlobBrushOptions/limitMerge\tBoolean\t0",
            "Blob Brush Tool Options > Merge Only with Selection\tplugin/BlobBrushOptions/limitMerge\tBoolean\t0",
            "塗りブラシツールオプション > 精度\tplugin/BlobBrushOptions/fairnessPreset\tInteger\t2\t0〜4",
            "Blob Brush Tool Options > Fidelity\tplugin/BlobBrushOptions/fairnessPreset\tInteger\t2\t0〜4",
            "塗りブラシツールオプション > サイズ «数値»\tplugin/BlobBrushTool/diameter\tReal\t10.0\tRead Only?",
            "Blob Brush Tool Options > Size «Real»\tplugin/BlobBrushTool/diameter\tReal\t10.0\tRead Only?",
            "塗りブラシツールオプション > サイズ «種類»\tplugin/BlobBrushTool/diameterDependsOn\tInteger\t0\tRead Only?",
            "Blob Brush Tool Options > Size «Type»\tplugin/BlobBrushTool/diameterDependsOn\tInteger\t0\tRead Only?",
            "塗りブラシツールオプション > サイズ > 変位\tplugin/BlobBrushTool/diameterVariation\tReal\t0.0\tRead Only?",
            "Blob Brush Tool Options > Size > Variation\tplugin/BlobBrushTool/diameterVariation\tReal\t0.0\tRead Only?",
            "塗りブラシツールオプション > 角度 «数値»\tplugin/BlobBrushTool/angle\tReal\t0.0\tRead Only?",
            "Blob Brush Tool Options > Angle «Real»\tplugin/BlobBrushTool/angle\tReal\t0.0\tRead Only?",
            "塗りブラシツールオプション > 角度 «種類»\tplugin/BlobBrushTool/angleDependsOn\tInteger\t0\tRead Only?",
            "Blob Brush Tool Options > Angle «Type»\tplugin/BlobBrushTool/angleDependsOn\tInteger\t0\tRead Only?",
            "塗りブラシツールオプション > 角度 > 変位\tplugin/BlobBrushTool/angleVariation\tReal\t0.0\tRead Only?",
            "Blob Brush Tool Options > Angle > Variation\tplugin/BlobBrushTool/angleVariation\tReal\t0.0\tRead Only?",
            "塗りブラシツールオプション > 真円率 «数値»\tplugin/BlobBrushTool/roundness\tReal\t100.0\tRead Only?",
            "Blob Brush Tool Options > Roundness «Real»\tplugin/BlobBrushTool/roundness\tReal\t100.0\tRead Only?",
            "塗りブラシツールオプション > 真円率 «種類»\tplugin/BlobBrushTool/roundnessDependsOn\tInteger\t0\tRead Only?",
            "Blob Brush Tool Options > Roundness «Type»\tplugin/BlobBrushTool/roundnessDependsOn\tInteger\t0\tRead Only?",
            "塗りブラシツールオプション > 真円率 > 変位\tplugin/BlobBrushTool/roundnessVariation\tReal\t0.0\tRead Only?",
            "Blob Brush Tool Options > Roundness > Variation\tplugin/BlobBrushTool/roundnessVariation\tReal\t0.0\tRead Only?",
            "\tplugin/BlobBrushTool/angleRelativeTo\tInteger\t0",
            "\tplugin/BlobBrushTool/angleRnd\tReal\t0.0",
            "\tplugin/BlobBrushTool/roundnessRnd\tReal\t0.0",
            "\tplugin/BlobBrushTool/diameterRnd\tReal\t0.0",
            "\tplugin/PhotoshopFileFormat/SpotOptionApplyToAll\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/SelectedLayerCompName\tString\t",
            "\tplugin/PhotoshopFileFormat/SelectedLayerCompID\tInteger\t4294967295",
            "\tplugin/PhotoshopFileFormat/ShowPreview\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/PixelAspectRatio\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/HiddenLayers\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/ImageMaps\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/Slices\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/PSD Import Option\tInteger\t2",
            "\tplugin/PhotoshopFileFormat/AI Script\tInteger\t1",
            "ファイル > 書き出し > 書き出し形式… > Photoshop 書き出しオプション > カラーモード\tplugin/PhotoshopFileFormat/ColorModel\tInteger\t2\t書類のカラーモードに依存し，セットしても適用されない",
            "File > Export > Export As… > Photoshop Export Options > Color Model\tplugin/PhotoshopFileFormat/ColorModel\tInteger\t2\t書類のカラーモードに依存し，セットしても適用されない",
            "ファイル > 書き出し > 書き出し形式… > Photoshop 書き出しオプション > 解像度\tplugin/PhotoshopFileFormat/DPI\tReal\t144",
            "File > Export > Export As… > Photoshop Export Options > Resolution\tplugin/PhotoshopFileFormat/DPI\tReal\t144",
            "ファイル > 書き出し > 書き出し形式… > Photoshop 書き出しオプション > オプション > 統合画像 | レイヤーを保持\tplugin/PhotoshopFileFormat/WriteLayers\tInteger\t0",
            "File > Export > Export As… > Photoshop Export Options > Options > Flat Image | Write Layers\tplugin/PhotoshopFileFormat/WriteLayers\tInteger\t0",
            "ファイル > 書き出し > 書き出し形式… > Photoshop 書き出しオプション > オプション > テキストの編集機能を保持\tplugin/PhotoshopFileFormat/LiveText\tBoolean\t1",
            "File > Export > Export As… > Photoshop Export Options > Options > Preserve Text Editability\tplugin/PhotoshopFileFormat/LiveText\tBoolean\t1",
            "ファイル > 書き出し > 書き出し形式… > Photoshop 書き出しオプション > オプション > 編集機能を最大限に保持\tplugin/PhotoshopFileFormat/Maximize Ediability\tBoolean\t0",
            "File > Export > Export As… > Photoshop Export Options > Options > Maximum Editability\tplugin/PhotoshopFileFormat/Maximize Ediability\tBoolean\t0",
            "ファイル > 書き出し > 書き出し形式… > Photoshop 書き出しオプション > オプション > アンチエイリアス\tplugin/PhotoshopFileFormat/AntiAlias\tInteger\t1",
            "File > Export > Export As… > Photoshop Export Options > Options > Anti-aliasing\tplugin/PhotoshopFileFormat/AntiAlias\tInteger\t1",
            "\tplugin/PhotoshopFileFormat/CustomDpi\tReal\t144",
            "\tplugin/PhotoshopFileFormat/ExportFormat\tInteger\t1",
            "\tplugin/PhotoshopFileFormat/WhichProfile\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/NestedLayers\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/CompoundShapes\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/PreserveSpotColors\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/NoHiddenLayersWarning\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/No100LayersWarning\tInteger\t0",
            "\tplugin/PhotoshopFileFormat/NoImageMapsWarning\tInteger\t0",
            "\tplugin/MultiEveDialogUI/height\tReal\t326",
            "\tplugin/MultiEveDialogUI/width\tReal\t645",
            "\tplugin/IllustratorUI/StylisticSetsFTUEShown\tInteger\t1",
            "\tplugin/IllustratorUI/VariableFontsFTUEShown\tInteger\t1",
            "\tplugin/AI Toolbox UI Plugin/kToolBarDrawerIntroShown\tInteger\t1",
            "ファイル > 別名で保存… > SVG オプション > SVG プロファイル\tplugin/svgOptionDlg/O_SVGDTD\tInteger\t0",
            "File > Save As… > SVG Options > SVG Profiles\tplugin/svgOptionDlg/O_SVGDTD\tInteger\t0",
            "ファイル > 別名で保存… > SVG オプション > フォント > 文字\tplugin/svgOptionDlg/O_SVFontType\tInteger\t1",
            "File > Save As… > SVG Options > Fonts > Type\tplugin/svgOptionDlg/O_SVFontType\tInteger\t1",
            "ファイル > 別名で保存… > SVG オプション > フォント > サブセット\tplugin/SVGFormat/fontSubsetting\tInteger\t1",
            "File > Save As… > SVG Options > Fonts > Subsetting\tplugin/SVGFormat/fontSubsetting\tInteger\t1",
            "ファイル > 別名で保存… > SVG オプション > オプション > 画像の場所\tplugin/svgOptionDlg/O_EmbedRasterLoc\tInteger\t2",
            "File > Save As… > SVG Options > Options > Image Location\tplugin/svgOptionDlg/O_EmbedRasterLoc\tInteger\t2",
            "ファイル > 別名で保存… > SVG オプション > オプション > Illustrator の編集機能を保持\tplugin/svgOptionDlg/O_IncludePGF\tBoolean\t0",
            "File > Save As… > SVG Options > Options > Preserve Illustrator Editing Capabilities\tplugin/svgOptionDlg/O_IncludePGF\tBoolean\t0",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > CSS プロパティ\tplugin/SVGFormat/StyleType\tInteger\t3",
            "File > Save As… > SVG Options > Advanced Options > CSS Properties\tplugin/SVGFormat/StyleType\tInteger\t3",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > 未使用グラフィックスタイルを含める\tplugin/svgOptionDlg/O_IncludeUnusedStyles\tBoolean\t0",
            "File > Save As… > SVG Options > Advanced Options > Include Unused Graphic Styles\tplugin/svgOptionDlg/O_IncludeUnusedStyles\tBoolean\t0",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > 小数点以下の桁数\tplugin/SVGFormat/CoordinatePrecision2\tInteger\t3",
            "File > Save As… > SVG Options > Advanced Options > Decimal Places\tplugin/SVGFormat/CoordinatePrecision2\tInteger\t3",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > エンコーディング\tplugin/SVGFormat/Encoding\tInteger\t2",
            "File > Save As… > SVG Options > Advanced Options > Encoding\tplugin/SVGFormat/Encoding\tInteger\t2",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > <tspan> エレメントの出力を制御\tplugin/svgOptionDlg/O_DisbleAutoKerning\tBoolean\t1",
            "File > Save As… > SVG Options > Advanced Options > Output fewer <tspan> elements\tplugin/svgOptionDlg/O_DisbleAutoKerning\tBoolean\t1",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > パス上テキストに <textPath> エレメントを使用\tplugin/svgOptionDlg/O_UseSVGTextOnPath\tBoolean\t1",
            "File > Save As… > SVG Options > Advanced Options > Use <textPath> element for Text on Path\tplugin/svgOptionDlg/O_UseSVGTextOnPath\tBoolean\t1",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > レスポンシブ\tplugin/svgOptionDlg/O_ResponsiveSVG\tBoolean\t1",
            "File > Save As… > SVG Options > Advanced Options > Responsive\tplugin/svgOptionDlg/O_ResponsiveSVG\tBoolean\t1",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > スライスデータを含める\tplugin/svgOptionDlg/O_IncludeSlices\tBoolean\t0",
            "File > Save As… > SVG Options > Advanced Options > Include Slicing Data\tplugin/svgOptionDlg/O_IncludeSlices\tBoolean\t0",
            "ファイル > 別名で保存… > SVG オプション > 詳細オプション > XMP を含める\tplugin/svgOptionDlg/O_IncludeXAP\tBoolean\t0",
            "File > Save As… > SVG Options > Advanced Options > Include XMP\tplugin/svgOptionDlg/O_IncludeXAP\tBoolean\t0",
            "ファイル > 別名で保存… > SVG オプション > 基本オプション | 詳細オプション\tplugin/svgOptionDlg/O_DialogExpanded\tBoolean\t1\t{true: 詳細オプションを表示, false: 基本オプションのみ表示}",
            "File > Save As… > SVG Options > Less Options | More Options\tplugin/svgOptionDlg/O_DialogExpanded\tBoolean\t1\t{true: 詳細オプションを表示, false: 基本オプションのみ表示}",
            "\tplugin/svgOptionDlg/O_IncludeTemplate\tBoolean\t0\tDOCTYPEやデータ駆動型グラフィックスの変数を含むSVGになる。何かそういった専門機能で使うのかもしれない",
            "\tplugin/svgOptionDlg/O_IncludeAdobeNameSpace\tBoolean\t0\txmlns:x=\"&ns_extend;\" xmlns:i=\"&ns_ai;\" xmlns:graph=\"&ns_graphs;\" xmlns:a=\"http://ns.adobe.com/AdobeSVGViewerExtensions/3.0/\"\\nなどが追加される",
            "\tplugin/svgOptionDlg/O_RoundTrip\tInteger\t0",
            "\tplugin/svgOptionDlg/O_TextOnPath\tInteger\t1",
            "\tplugin/svgOptionDlg/O_GradientTolerance\tReal\t0.25",
            "\tplugin/svgOptionDlg/O_RasterResolution\tInteger\t72",
            "\tplugin/svgOptionDlg/O_EmbedFontFormats\tInteger\t2\t動作しない？ CEFフォーマットなどの話だろうか",
            "\tplugin/svgOptionDlg/O_Clip_To_Artboard\tInteger\t0",
            "\tplugin/svgOptionDlg/O_Constrain_Proportions\tInteger\t1",
            "\tplugin/svgOptionDlg/O_Also_Export_Compressed\tInteger\t0",
            "\tplugin/svgOptionDlg/O_Export_LayerAsTitle\tInteger\t0",
            "\tplugin/svgOptionDlg/O_ExportHiddenObjects\tInteger\t0",
            "\tplugin/svgOptionDlg/O_Anti_Alias_Artwork\tInteger\t1",
            "\tplugin/svgOptionDlg/O_Resolution_Unit\tInteger\t0",
            "\tplugin/svgOptionDlg/O_Height_Unit\tInteger\t2",
            "\tplugin/svgOptionDlg/O_Width_Unit\tInteger\t2",
            "\tplugin/SVGFormat/CoordinatePrecision\tInteger\t1",
            "\tplugin/SVGFormat/FontLocation\tInteger\t1",
            "\tplugin/SVGFormat/RestrictLinksToDocDirectory\tInteger\t0",
            "\tplugin/SVGFormat/FormatForFonts\tInteger\t1",
            "\tplugin/SVGFormat/FileCompression\tInteger\t1",
            "\tplugin/SVGFormat/systemFont\tInteger\t1",
            "\tplugin/SVGFormat/zoom\tInteger\t1",
            "\tplugin/SVGFormat/Rendering\tInteger\t1",
            "\tplugin/SVGFormat/FormatForImages\tInteger\t1",
            "\tplugin/SVGFormat/allowedSVGLinkingDepth\tInteger\t1",
            "\tplugin/PDF/CompressPGF\tInteger\t1",
            "\tplugin/sangam/readers/sTEXT\tInteger\t0",
            "\tplugin/sangam/readers/RTF\tString\t",
            "\tplugin/sangam/readers/sRTF\tInteger\t12",
            "\tplugin/sangam/readers/InPlayMode\tInteger\t0",
            "\tplugin/Illustrator/WritingStartupFile\tInteger\t0",
            "\tplugin/Illustrator/IsGB18030\tInteger\t0",
            "\tplugin/Vector Sculpting/Mesh Visibility\tInteger\t1",
            "\tplugin/Adobe Smart Edit/SmartEditOnBoardingCount\tInteger\t3",
            "\tplugin/Adobe Smart Edit/SmartEditOnBoardingShown\tInteger\t0",
            "\tplugin/com.adobe.illustrator.WorkspaceBasicsOnBoarding.extension/Location and Size/r\tInteger\t1130",
            "\tplugin/com.adobe.illustrator.WorkspaceBasicsOnBoarding.extension/Location and Size/b\tInteger\t738",
            "\tplugin/com.adobe.illustrator.WorkspaceBasicsOnBoarding.extension/Location and Size/t\tInteger\t288",
            "\tplugin/com.adobe.illustrator.WorkspaceBasicsOnBoarding.extension/Location and Size/l\tInteger\t550",
            "\tplugin/Properties Panel/WorkspaceCardCount\tInteger\t1",
            "\tplugin/Properties Panel/WorkspaceCardShown\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > アートボード | アセット\tplugin/SmartExportUI/SmartExportActiveTab\tString\tSmartExportActiveTabArtboard",
            "File > Export > Export for Screens… > Artboards | Assets\tplugin/SmartExportUI/SmartExportActiveTab\tString\tSmartExportActiveTabArtboard",
            "ファイル > 書き出し > スクリーン用に書き出し… > アートボード > 選択 > すべて | 範囲 | ドキュメント全体\tplugin/SmartExportUI/SmartExportContentTypeForArtboard\tString\tall",
            "File > Export > Export for Screens… > Artboards > Select > All | Range | Full Document\tplugin/SmartExportUI/SmartExportContentTypeForArtboard\tString\tall",
            "ファイル > 書き出し > スクリーン用に書き出し… > アートボード > 選択 > 範囲 «文字列»\tplugin/SmartExportUI/SmartExportDialogRange\tString\t45659",
            "File > Export > Export for Screens… > Artboards > Select > Range «String»\tplugin/SmartExportUI/SmartExportDialogRange\tString\t45659",
            "ファイル > 書き出し > スクリーン用に書き出し… > アートボード > 選択 > 裁ち落としを含める\tplugin/SmartExportUI/IncludeBleedInExport\tBoolean\t0",
            "File > Export > Export for Screens… > Artboards > Select > Include Bleed\tplugin/SmartExportUI/IncludeBleedInExport\tBoolean\t0",
            "ファイル > 書き出し > スクリーン用に書き出し… > アセット > 選択 > すべてのアセット\tplugin/SmartExportUI/SmartExportContentTypeForAssets\tString\tall",
            "File > Export > Export for Screens… > Assets > Select > All Assets\tplugin/SmartExportUI/SmartExportContentTypeForAssets\tString\tall",
            "ファイル > 書き出し > スクリーン用に書き出し… > 書き出し先\tplugin/SmartExportUI/SmartExportFolderPath\tString\t/Users/«username»/Desktop",
            "File > Export > Export for Screens… > Export to\tplugin/SmartExportUI/SmartExportFolderPath\tString\t/Users/«username»/Desktop",
            "ファイル > 書き出し > スクリーン用に書き出し… > 書き出し後に場所を開く\tplugin/SmartExportUI/OpenLocationAfterExportPreference\tBoolean\t1",
            "File > Export > Export for Screens… > Open Location after Export\tplugin/SmartExportUI/OpenLocationAfterExportPreference\tBoolean\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > サブフォルダーを作成\tplugin/SmartExportUI/CreateFoldersPreference\tBoolean\t0",
            "File > Export > Export for Screens… > Create Sub-folders\tplugin/SmartExportUI/CreateFoldersPreference\tBoolean\t0",
            "ファイル > 書き出し > スクリーン用に書き出し… > サブフォルダーを作成 > 拡大・縮小 | 形式\tplugin/SmartExportUI/CreateFoldersByScalePreference\tInteger\t1",
            "File > Export > Export for Screens… > Create Sub-folders > Scale | Format\tplugin/SmartExportUI/CreateFoldersByScalePreference\tInteger\t1",
            "ファイル > 書き出し > スクリーン用に書き出し… > プレフィックス\tplugin/SmartExportUI/SmartExportFileNamePrefix\tString\timg_",
            "File > Export > Export for Screens… > Prefix\tplugin/SmartExportUI/SmartExportFileNamePrefix\tString\timg_",
            "ファイル > 書き出し > スクリーン用に書き出し… > サムネール (大) を表示 | サムネール (小) を表示\tplugin/SmartExportUI/SmartExportPreviewType\tString\tPreviewTypeGrid",
            "File > Export > Export for Screens… > Large thumbnail view | Small thumbnail view\tplugin/SmartExportUI/SmartExportPreviewType\tString\tPreviewTypeGrid",
            "\tplugin/SmartExportUI/ExportSelection/ShowExportPanel/Count\tInteger\t1",
            "\tplugin/SmartExportUI/ExportSelection/Count\tInteger\t1",
            "\tplugin/SmartExportUI/ExportPanel/AssetAdded/ExcludingExpSelRoute\tInteger\t1",
            "\tplugin/SmartExportUI/ShowLayersPanelAddAssetFTUE\tInteger\t0",
            "\tplugin/SmartExportUI/SmartExportDialogBaseWidth\tReal\t840\tサムネールを表示する領域の幅",
            "\tplugin/SmartExportUI/SmartExportDialogBaseHeight\tReal\t542\tサムネールを表示する領域の高さ",
            "\tplugin/SmartExportUI/exportFolderPath\tString\t",
            "\tplugin/SmartExportUI/OpenLocationAfterExportFromPanelPreference\tInteger\t0",
            "\tplugin/NamedStylePopupPanel/ThumbnailType\tInteger\t0",
            "\tplugin/NamedStylePopupPanel/StylesView\tInteger\t0",
            "\tplugin/StatusBarView/FontSizeReduction\tReal\t2.0",
            "\tplugin/DontShowWarningAgain/Missing_AutoActivate_Font\tInteger\t0",
            "\tplugin/DontShowWarningAgain/text/dontshowAutoActivateOnboarding\tInteger\t1",
            "\tplugin/DontShowWarningAgain/AIMacPluginsLoadingErrorDueToWrongArchitectureDialog\tInteger\t0",
            "\tplugin/DontShowWarningAgain/AIPluginsLoadingErrorDialog\tInteger\t0",
            "\tplugin/DontShowWarningAgain/PDFExportDialog.NoPGFBecauseOfPFDX\tInteger\t0",
            "\tplugin/DontShowWarningAgain/lockedHiddenObjectsNotMoved\tInteger\t0",
            "\tplugin/DontShowWarningAgain/illegalJoin\tInteger\t0",
            "\tplugin/DontShowWarningAgain/illegalScissor\tInteger\t0",
            "\tplugin/DontShowWarningAgain/LegacyReflowWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/legacyCSSaveWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/DropLegacyScale9\tInteger\t0",
            "\tplugin/DontShowWarningAgain/ShowPathfinderGroupWarning\tInteger\t1",
            "\tplugin/DontShowWarningAgain/AISpotTransparencyAlertDlg\tInteger\t0",
            "\tplugin/DontShowWarningAgain/MissingLinkFileDummyKey\tInteger\t0",
            "\tplugin/DontShowWarningAgain/notOkToInteract\tInteger\t1",
            "\tplugin/DontShowWarningAgain/showClipboardWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/applyAllLinkOption\tInteger\t0",
            "\tplugin/DontShowWarningAgain/AI9ClrMgmtNoProfInFile\tInteger\t0",
            "\tplugin/DontShowWarningAgain/noSelection\tInteger\t0",
            "\tplugin/DontShowWarningAgain/illegalDeleteKnot\tInteger\t0",
            "\tplugin/DontShowWarningAgain/Swatches/PDM_ExpansionAlert\tInteger\t0",
            "\tplugin/DontShowWarningAgain/Swatches/PDM_NewPattern\tInteger\t0",
            "\tplugin/DontShowWarningAgain/ApplyEffectWarningDlg\tInteger\t0",
            "\tplugin/DontShowWarningAgain/-Vectorize Warning For Big Image Trace\tInteger\t0",
            "\tplugin/DontShowWarningAgain/SaveForWebLatinCharacterDontShowAgain\tInteger\t0",
            "\tplugin/DontShowWarningAgain/illegalAddKnot\tInteger\t0",
            "\tplugin/DontShowWarningAgain/AttributeOverprintRGBWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/NoHiddenLayersWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/NoImageMapsWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/NoSelectiveMergeWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/illegalAverage\tInteger\t0",
            "\tplugin/DontShowWarningAgain/VariablesWarning\tInteger\t0",
            "\tplugin/DontShowWarningAgain/svgSaveWarningSupress\tInteger\t0",
            "\tplugin/PreferenceUI/dialogHeight\tReal\t621.0",
            "\tplugin/PreferenceUI/dialogWidth\tReal\t765.0",
            "\tplugin/PreferenceUI/PrefDialogDimensionVersion\tString\t30.8.2",
            "\tplugin/AssetManagement/vcLinkStatusUpdateCycleInTicks\tInteger\t120",
            "\tplugin/AssetManagement/vcReplicaStatusUpdateCycleInTicks\tInteger\t30",
            "\tplugin/AssetManagement/bridgeTalkPumpCycleInTicks\tInteger\t30",
            "\tplugin/WorkspacePrefix/Workspace Essentials Handled\tInteger\t1",
            "\tplugin/WorkspacePrefix/Previous Workspace Name\tString\tEssentials\t前回使ったワークスペース名",
            "\tplugin/WorkspacePrefix/Last Actual Used Workspace Name\tString\tEssentials\t現在のワークスペース名",
            "\tplugin/WorkspacePrefix/Last Used Workspace Mode\tInteger\t1",
            "\tplugin/WorkspacePrefix/Last Used Workspace Name\tString\tEssentials\t現在のワークスペース名。Last Actual Used Workspace Nameとの違いはわからない",
            "\tplugin/Navigator/Color/blue\tReal\t0.31",
            "\tplugin/Navigator/Color/green\tReal\t0.31",
            "\tplugin/Navigator/Color/red\tReal\t1.0",
            "\tplugin/Navigator/Artboard Mode\tInteger\t1",
            "\tplugin/Navigator/DontDrawDash\tInteger\t1",
            "\tplugin/Navigator/Greeking\tReal\t72.0",
            "\tplugin/Align/StickyAlignOption\tInteger\t2\t次回起動時に適用",
            "\tplugin/KBSCShortcutFile/IsPreset\tBoolean\t0\tキーボードショートカット設定が用意されたプリセットかどうか",
            "\tplugin/KBSCShortcutFile/FileName\tString\t\tkysファイルの名前",
            "\tplugin/NamedStyle/ThumbnailType\tInteger\t0",
            "\tplugin/NamedStyle/StylesView\tInteger\t0",
            "\tplugin/Symbols/VisibleKind\tInteger\t0",
            "変形パネル > 基準点\tplugin/Transform/AnchorPoint\tInteger\t4\tパネルの見た目は更新されないので，再表示が必要",
            "Transform Panel > Reference Point\tplugin/Transform/AnchorPoint\tInteger\t4\tパネルの見た目は更新されないので，再表示が必要",
            "\tplugin/AdobeBrush/ThumbnailView\tInteger\t1",
            "CSS プロパティパネル > 書き出しオプション… > CSS 単位\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportUnit\tInteger\t4294967295",
            "CSS Properties Panel > Export Options… > CSS Units\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportUnit\tInteger\t4294967295",
            "CSS プロパティパネル > 書き出しオプション… > オブジェクトのアピアランス > 塗りを含める\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportFill\tBoolean\t1",
            "CSS Properties Panel > Export Options… > Object Appearance > Include Fill\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportFill\tBoolean\t1",
            "CSS プロパティパネル > 書き出しオプション… > オブジェクトのアピアランス > 線を含める\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportStroke\tBoolean\t1",
            "CSS Properties Panel > Export Options… > Object Appearance > Include Stroke\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportStroke\tBoolean\t1",
            "CSS プロパティパネル > 書き出しオプション… > オブジェクトのアピアランス > 不透明度を含める\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportOpacity\tBoolean\t1",
            "CSS Properties Panel > Export Options… > Object Appearance > Include Opacity\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportOpacity\tBoolean\t1",
            "CSS プロパティパネル > 書き出しオプション… > 位置とサイズ > 絶対位置を含める\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportPosition\tBoolean\t0",
            "CSS Properties Panel > Export Options… > Position and Size > Include Absolute Position\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportPosition\tBoolean\t0",
            "CSS プロパティパネル > 書き出しオプション… > 位置とサイズ > サイズを含める\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportDimension\tBoolean\t0",
            "CSS Properties Panel > Export Options… > Position and Size > Include Dimensions\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportDimension\tBoolean\t0",
            "CSS プロパティパネル > 書き出しオプション… > オプション > 名称未設定オブジェクト用に CSS を生成\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportUnnamedObjPref\tBoolean\t0",
            "CSS Properties Panel > Export Options… > Options > Generate CSS for Unnamed Objects\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportUnnamedObjPref\tBoolean\t0",
            "CSS プロパティパネル > 書き出しオプション… > オプション > サポートされていないアートをラスタライズ\tplugin/CSSExportPreference/CSSExtract/CSSExport/UnsupportedArt\tBoolean\t1",
            "CSS Properties Panel > Export Options… > Options > Rasterize Unsupported Art\tplugin/CSSExportPreference/CSSExtract/CSSExport/UnsupportedArt\tBoolean\t1",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/OtherResolution\tReal\t300.0",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/ResolutionOption\tInteger\t300",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/Opera\tBoolean\t1\tCreateCSS.jsx内で，Opera用コードをつけるかどうかを制御する",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/IExplorer\tBoolean\t1\tCreateCSS.jsx内で，IE用コードをつけるかどうかを制御する",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/FireFox\tBoolean\t1\tCreateCSS.jsx内で，Firefox用コードをつけるかどうかを制御する",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/Webkit\tBoolean\t1\tCreateCSS.jsx内で，Webkit用コードをつけるかどうかを制御する",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/VendorPref\tBoolean\t1\tCreateCSS.jsx内で，ベンダープレフィックスをつけるかどうかを制御する",
            "\tplugin/CSSExportPreference/CSSExtract/CSSExport/ExportMode\tInteger\t3\tCreateCSS.jsx内で，モードを切り替えるのに使用",
            "\tplugin/AdobeSwatchPopup_Stroke/GroupsListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Stroke/PatternListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Stroke/GradientListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Stroke/ColorListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Stroke/ShowAllListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Stroke/VisibleKind\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Stroke/CurrentListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Stroke/ShowFindField\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/GroupsListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/PatternListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/GradientListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/ColorListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/ShowAllListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/VisibleKind\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/CurrentListStyle\tInteger\t0",
            "\tplugin/AdobeSwatchPopup_Fill/ShowFindField\tInteger\t0",
            "\tplugin/AdobeSwatch_/GroupsListStyle\tInteger\t0",
            "\tplugin/AdobeSwatch_/PatternListStyle\tInteger\t0",
            "\tplugin/AdobeSwatch_/GradientListStyle\tInteger\t0",
            "\tplugin/AdobeSwatch_/ColorListStyle\tInteger\t0",
            "\tplugin/AdobeSwatch_/ShowAllListStyle\tInteger\t0",
            "\tplugin/AdobeSwatch_/VisibleKind\tInteger\t0",
            "\tplugin/AdobeSwatch_/CurrentListStyle\tInteger\t0",
            "\tplugin/AdobeSwatch_/ShowFindField\tInteger\t0",
            "\tplugin/ICCColorMgmt3.0/settingsFileValid\tInteger\t0",
            "\tplugin/ICCColorMgmt3.0/useAdvanced\tInteger\t0",
            "\tplugin/ICCColorMgmt3.0/useAI6\tInteger\t0",
            "\tplugin/ICCColorMgmt3.0/settingsFile\tString\t/Library/Application Support/Adobe/Color/Settings/Recommended/North America General Purpose.csf",
            "ファイル > 配置… > PDF を配置 > «ドキュメントに配置する PDF のページを選択してください。»\tplugin/PDFImport/PageNumber\tInteger\t0\tRead Only?",
            "File > Place… > Place PDF > «Select a page from the PDF to place into the document.»\tplugin/PDFImport/PageNumber\tInteger\t0\tRead Only?",
            "ファイル > 配置… > PDF を配置 > トリミング\tplugin/PDFImport/CropTo\tInteger\t4\tRead Only?",
            "File > Place… > Place PDF > Crop to\tplugin/PDFImport/CropTo\tInteger\t4\tRead Only?",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > PDF > Adobe PDF プリセット\tplugin/PDFExport/SmartExportDefaultPDFPresetName\tString\t\t値：\\n[Smallest File Size (PDF 1.6)]",
            "File > Export > Export for Screens… > Format Settings > PDF > Adobe PDF Preset\tplugin/PDFExport/SmartExportDefaultPDFPresetName\tString\t\t値：\\n[Smallest File Size (PDF 1.6)]",
            "\tplugin/PDFExport/AI12PDF_RegistryName\tString\t",
            "\tplugin/PDFExport/AI12PDF_OutputConditionIdentifier\tString\t",
            "\tplugin/PDFExport/AI12PDF_OutputCondition\tString\t",
            "\tplugin/PDFExport/AI12PDF_OutputIntentProfileName\tString\t",
            "\tplugin/PDFExport/AI12PDF_DestinationName\tString\t",
            "\tplugin/PDFExport/AI12PDF_Description\tString\t",
            "\tplugin/PDFExport/AI12PDF_Trapped\tInteger\t0",
            "\tplugin/PDFExport/AI12PDF_ColorConversionPolicy\tInteger\t1",
            "\tplugin/PDFExport/AI12PDF_OutputIntentProfileNamePolicy\tInteger\t1",
            "\tplugin/PDFExport/AI12PDF_DestinationPolicy\tInteger\t1",
            "\tplugin/PDFExport/AI12PDF_ProfileInclusionPolicy\tInteger\t1",
            "\tplugin/PDFExport/AI12PDF_Standard\tInteger\t1",
            "\tplugin/PDFExport/AI12PDF_DontShowSecurityWarning\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetPrinterResolution\tReal\t800.0",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetOutlineText\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetOutlineStroke\tInteger\t1",
            "\tplugin/PDFExport/AI16PDF_FlatteningPresetAntiAlias\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetClipComplexRegion\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetVectorBalance\tInteger\t75",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetMinResolution\tReal\t150.0",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetMaxResolution\tReal\t300.0",
            "\tplugin/PDFExport/AI11PDF_FlatteningPresetName\tString\t",
            "\tplugin/PDFExport/AI11PDF_FlattenTransparency\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_PreserveAcrobatLayers\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_Overprint\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_SubsetFontRatio\tInteger\t1",
            "\tplugin/PDFExport/AI14PDF_DocBleed\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_BleedBottom\tReal\t0.0",
            "\tplugin/PDFExport/AI11PDF_BleedTop\tReal\t0.0",
            "\tplugin/PDFExport/AI11PDF_BleedRight\tReal\t0.0",
            "\tplugin/PDFExport/AI11PDF_BleedLeft\tReal\t0.0",
            "\tplugin/PDFExport/AI11PDF_BleedLink\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_OffsetFromArtworkBottom\tReal\t6.0",
            "\tplugin/PDFExport/AI11PDF_OffsetFromArtworkTop\tReal\t6.0",
            "\tplugin/PDFExport/AI11PDF_OffsetFromArtworkRight\tReal\t6.0",
            "\tplugin/PDFExport/AI11PDF_OffsetFromArtworkLeft\tReal\t6.0",
            "\tplugin/PDFExport/AI11PDF_TrimMarkWeight\tReal\t0.25",
            "\tplugin/PDFExport/AI11PDF_PrinterMarkType\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_PageInfo\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_ColorBars\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_RegMarks\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_TrimMarks\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_CompressArt\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_MonoDownsampleImageAbove\tInteger\t450",
            "\tplugin/PDFExport/AI11PDF_MonochromeDownsampleResolution\tInteger\t300",
            "\tplugin/PDFExport/AI11PDF_MonochromeDownsampleKind\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_MonochromeCompressionKind\tInteger\t4",
            "\tplugin/PDFExport/AI11PDF_GrayTileSize\tInteger\t256",
            "\tplugin/PDFExport/AI11PDF_GrayDownsampleImageAbove\tInteger\t225",
            "\tplugin/PDFExport/AI11PDF_GrayDownsampleResolution\tInteger\t150",
            "\tplugin/PDFExport/AI11PDF_GrayDownsampleKind\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_GrayCompressionQuality\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_GrayCompressionKind\tInteger\t6",
            "\tplugin/PDFExport/AI11PDF_ColorTileSize\tInteger\t256",
            "\tplugin/PDFExport/AI11PDF_ColorDownsampleImageAbove\tInteger\t225",
            "\tplugin/PDFExport/AI11PDF_ColorDownsampleResolution\tInteger\t150",
            "\tplugin/PDFExport/AI11PDF_ColorDownsampleKind\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_ColorCompressionQuality\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_ColorCompressionKind\tInteger\t6",
            "\tplugin/PDFExport/AI11PDF_FastWebView\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_ViewPDFFile\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_GenerateThumbnails\tInteger\t1",
            "\tplugin/PDFExport/AI11PDF_PreserveIllustratorEditingCapabilities\tInteger\t1",
            "\tplugin/PDFExport/AI12PDF_UsePrintTiling\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_Compatibility\tInteger\t3",
            "\tplugin/PDFExport/OptionSetNameSE\tString\t",
            "\tplugin/PDFExport/OptionSetName\tString\t",
            "\tplugin/PDFExport/OptionSet\tInteger\t2",
            "\tplugin/PDFExport/FlattenerListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/MonoImageListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/GrayImageListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/ColorImageListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/SecurityListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/AdvancedListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/OutputPDFXExpanded\tInteger\t0",
            "\tplugin/PDFExport/OutputColorExpanded\tInteger\t0",
            "\tplugin/PDFExport/OutputListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/MarksBleedsListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/CompressionListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/GeneralListEntryExpanded\tInteger\t0",
            "\tplugin/PDFExport/DescriptionListEntryExpanded\tInteger\t0",
            "消しゴムツールオプション > 角度 «数値»\tplugin/EraserTool/angle\tReal\t0.0\tRead Only?",
            "Eraser Tool Options > Angle «Real»\tplugin/EraserTool/angle\tReal\t0.0\tRead Only?",
            "消しゴムツールオプション > 角度 «種類»\tplugin/EraserTool/angleDependsOn\tInteger\t0\tRead Only?",
            "Eraser Tool Options > Angle «Type»\tplugin/EraserTool/angleDependsOn\tInteger\t0\tRead Only?",
            "消しゴムツールオプション > 角度 > 変位\tplugin/EraserTool/angleVariation\tReal\t0.0\tRead Only?",
            "Eraser Tool Options > Angle > Variation\tplugin/EraserTool/angleVariation\tReal\t0.0\tRead Only?",
            "消しゴムツールオプション > 真円率 «数値»\tplugin/EraserTool/roundness\tReal\t100.0\tRead Only?",
            "Eraser Tool Options > Roundness «Real»\tplugin/EraserTool/roundness\tReal\t100.0\tRead Only?",
            "消しゴムツールオプション > 真円率 «種類»\tplugin/EraserTool/roundnessDependsOn\tInteger\t0\tRead Only?",
            "Eraser Tool Options > Roundness «Type»\tplugin/EraserTool/roundnessDependsOn\tInteger\t0\tRead Only?",
            "消しゴムツールオプション > 真円率 > 変位\tplugin/EraserTool/roundnessVariation\tReal\t0.0\tRead Only?",
            "Eraser Tool Options > Roundness > Variation\tplugin/EraserTool/roundnessVariation\tReal\t0.0\tRead Only?",
            "消しゴムツールオプション > サイズ «数値»\tplugin/EraserTool/diameter\tReal\t10.0\tRead Only?",
            "Eraser Tool Options > Size «Real»\tplugin/EraserTool/diameter\tReal\t10.0\tRead Only?",
            "消しゴムツールオプション > サイズ «種類»\tplugin/EraserTool/diameterDependsOn\tInteger\t0\tRead Only?",
            "Eraser Tool Options > Size «Type»\tplugin/EraserTool/diameterDependsOn\tInteger\t0\tRead Only?",
            "消しゴムツールオプション > サイズ > 変位\tplugin/EraserTool/diameterVariation\tReal\t0.0\tRead Only?",
            "Eraser Tool Options > Size > Variation\tplugin/EraserTool/diameterVariation\tReal\t0.0\tRead Only?",
            "\tplugin/EraserTool/angleRelativeTo\tInteger\t0",
            "\tplugin/EraserTool/angleRnd\tReal\t0.0",
            "\tplugin/EraserTool/roundnessRnd\tReal\t0.0",
            "\tplugin/EraserTool/diameterRnd\tReal\t0.0",
            "\tplugin/dBrush/OptionsDialog/PreviewState\tInteger\t0",
            "\tplugin/dBrush/ShowCursorAnnotatorForMouse\tInteger\t0",
            "\tplugin/dBrush/ShowGeneralBrushAnnotation\tInteger\t0",
            "\tplugin/dBrush/ShowPoseAnnotation\tInteger\t0",
            "\tplugin/dBrush/CountBeforeSavePrint\tInteger\t1",
            "\tplugin/dBrush/MousePosePressure\tReal\t0.5",
            "\tplugin/dBrush/MousePoseTwist\tReal\t0.0",
            "\tplugin/dBrush/MousePoseAltitude\tReal\t45.0000012522",
            "\tplugin/dBrush/MousePoseAzimuth\tReal\t45.0000012522",
            "\tplugin/dBrush/MousePoseFollowPath\tInteger\t1",
            "\tplugin/dBrush/UseSampledStepping\tInteger\t1",
            "\tplugin/dBrush/ShowSteppingTrajectories\tInteger\t0",
            "\tplugin/dBrush/PhysicsThreads\tInteger\t0",
            "\tplugin/dBrush/UseODEPhysics\tInteger\t0",
            "\tplugin/dBrush/UseSubSteppingInFeedback\tInteger\t0",
            "スムーズツールオプション > 精度\tplugin/Adobe Freehand Smooth Tool/fairness_preset\tInteger\t2\t次回起動時に取得／適用",
            "Smooth Tool Options > Fidelity\tplugin/Adobe Freehand Smooth Tool/fairness_preset\tInteger\t2\t次回起動時に取得／適用",
            "鉛筆ツールオプション > 精度\tplugin/Adobe Freehand Tool/fairness_preset\tInteger\t2\tRead Only?",
            "Pencil Tool Options > Fidelity\tplugin/Adobe Freehand Tool/fairness_preset\tInteger\t2\tRead Only?",
            "鉛筆ツールオプション > オプション > 鉛筆の線に塗りを適用\tplugin/Adobe Freehand Tool/fill\tBoolean\t0\tRead Only?",
            "Pencil Tool Options > Options > Fill new pencil strokes\tplugin/Adobe Freehand Tool/fill\tBoolean\t0\tRead Only?",
            "鉛筆ツールオプション > オプション > 選択を解除しない\tplugin/Adobe Freehand Tool/select\tBoolean\t1\tRead Only?",
            "Pencil Tool Options > Options > Keep selected\tplugin/Adobe Freehand Tool/select\tBoolean\t1\tRead Only?",
            "鉛筆ツールオプション > オプション > Option キーでスムーズツールを使用\tplugin/Adobe Freehand Tool/toggle_to_smooth\tBoolean\t0\tRead Only?",
            "Pencil Tool Options > Options > Option key toggles to Smooth Tool\tplugin/Adobe Freehand Tool/toggle_to_smooth\tBoolean\t0\tRead Only?",
            "鉛筆ツールオプション > オプション > 両端が次の範囲内のときにパスを閉じる «真偽値»\tplugin/Adobe Freehand Tool/auto_close\tBoolean\t1\tRead Only?",
            "Pencil Tool Options > Options > Close paths when ends are within «Checked»\tplugin/Adobe Freehand Tool/auto_close\tBoolean\t1\tRead Only?",
            "鉛筆ツールオプション > オプション > 両端が次の範囲内のときにパスを閉じる «Pixel»\tplugin/Adobe Freehand Tool/closing_distance\tReal\t15.0\tRead Only?",
            "Pencil Tool Options > Options > Close paths when ends are within «Pixel»\tplugin/Adobe Freehand Tool/closing_distance\tReal\t15.0\tRead Only?",
            "鉛筆ツールオプション > オプション > 選択したパスを編集\tplugin/Adobe Freehand Tool/edit_selection\tBoolean\t1\tRead Only?",
            "Pencil Tool Options > Options > Edit selected paths\tplugin/Adobe Freehand Tool/edit_selection\tBoolean\t1\tRead Only?",
            "鉛筆ツールオプション > オプション > 選択したパスを編集 > 範囲\tplugin/Adobe Freehand Tool/editing_distance\tReal\t6.0\tRead Only?",
            "Pencil Tool Options > Options > Edit selected paths > Within\tplugin/Adobe Freehand Tool/editing_distance\tReal\t6.0\tRead Only?",
            "効果 > ドキュメントのラスタライズ効果設定… > カラーモード/背景\tplugin/Rasterize/Defaults/Type\tInteger\t4\tRead Only?",
            "Effect >  Document Raster Effects Settings… > Color Model/Background\tplugin/Rasterize/Defaults/Type\tInteger\t4\tRead Only?",
            "効果 > ドキュメントのラスタライズ効果設定… > 解像度\tplugin/Rasterize/Defaults/DPI\tInteger\t72",
            "Effect >  Document Raster Effects Settings… > Resolution\tplugin/Rasterize/Defaults/DPI\tInteger\t72",
            "効果 > ドキュメントのラスタライズ効果設定… > オプション > アンチエイリアス\tplugin/Rasterize/Defaults/Alias\tBoolean\t0\tRead Only?",
            "Effect > Document Raster Effects Settings… > Options > Anti-alias\tplugin/Rasterize/Defaults/Alias\tBoolean\t0\tRead Only?",
            "効果 > ドキュメントのラスタライズ効果設定… > オプション > クリッピングマスクを作成\tplugin/Rasterize/Defaults/Mask\tBoolean\t0\tRead Only?",
            "Effect > Document Raster Effects Settings… > Options > Create Clipping Mask\tplugin/Rasterize/Defaults/Mask\tBoolean\t0\tRead Only?",
            "効果 > ドキュメントのラスタライズ効果設定… > オプション > オブジェクトの周囲に追加\tplugin/Rasterize/Defaults/Padding\tReal\t36.0\tRead Only?",
            "Effect > Document Raster Effects Settings… > Options > Add Around Object\tplugin/Rasterize/Defaults/Padding\tReal\t36.0\tRead Only?",
            "効果 > ドキュメントのラスタライズ効果設定… > オプション > 特色を保持\tplugin/Rasterize/Defaults/Spot\tBoolean\t1\tRead Only?",
            "Effect > Document Raster Effects Settings… > Options > Preserve spot colors\tplugin/Rasterize/Defaults/Spot\tBoolean\t1\tRead Only?",
            "\tplugin/Rasterize/rasterizeMemory\tInteger\t20",
            "オブジェクト > 分割・拡張… > グラデーションの分割・拡張 > 指定 オブジェクト\tplugin/objectExpand/gradientSteps\tInteger\t255\t次回起動時に適用",
            "Object > Expand… > Expand Gradient To > Specify Objects\tplugin/objectExpand/gradientSteps\tInteger\t255\t次回起動時に適用",
            "\tplugin/PG/PlaneSwitcherPosition\tInteger\t0",
            "\tplugin/PG/PlaneSwitcherVisible\tInteger\t1",
            "\tplugin/PG/infGridLines\tInteger\t1",
            "\tplugin/PG/infObjects\tInteger\t1",
            "パスファインダーパネル > パスファインダーオプション… > 余分なポイントを削除\tplugin/PathFinder/RemoveRedundantPoints\tBoolean\t1\t次回起動時に適用",
            "Pathfinder Panel > Pathfinder Options… > Remove Redundant Points\tplugin/PathFinder/RemoveRedundantPoints\tBoolean\t1\t次回起動時に適用",
            "\tplugin/SmartExport/SVGFormatSmartExport\tInteger\t0",
            "\tplugin/SmartExport/PDFFormatSmartExport\tInteger\t0",
            "\tplugin/SmartExport/JPEGFormatToExport\tString\tJPG File Format\tその他の型。動作不可？",
            "\tplugin/AIDefaultPersistentLibraryPrefix/_count\tInteger\t0",
            "\tplugin/Workspaces/PrefCommentsPanelShowSetInWorkspace\tInteger\t1",
            "\tplugin/Workspaces/PrefPropertiesPanelShowSetInWorkspace\tInteger\t1",
            "\tplugin/Workspaces/PrefDLPanelShowSetInWorkspace\tInteger\t1",
            "\tplugin/AISwatchLibPersistentPrefix/_count\tInteger\t0",
            "\tplugin/Debug/timeActions\tInteger\t0",
            "\tplugin/Debug/useDebugMenu\tInteger\t0",
            "\tplugin/FavoriteFamily/FavoriteFamilyCount\tInteger\t0\tお気に入りに追加したフォントファミリーの数。次回起動時に更新される",
            "\tplugin/FavoriteFamily/0/Name\tString\tMyriad Pro\tお気に入りに追加したフォントファミリーの名前。plugin/FavoriteFamily/1/Nameのように増える。次回起動時に更新される",
            "\tplugin/RecentFont/1/FontTechnology\tInteger\t100",
            "\tplugin/RecentFont/RecentFontCount\tInteger\t0",
            "プロパティパネル > アートボード > 変形 > 基準点\tcropareatool/CropAreaControlPalette/9PointRef\tInteger\t0\t次回起動時に適用。\\nコントロール > アートボード > 基準点\\nControl > Artboard > Reference Point",
            "Properties Panel > Artboard > Transform > Set the Reference Point\tcropareatool/CropAreaControlPalette/9PointRef\tInteger\t0\t次回起動時に適用。\\nコントロール > アートボード > 基準点\\nControl > Artboard > Reference Point",
            "プロパティパネル > アートボード > アートボード > オブジェクトと一緒に移動\tplugin/CropAreaPrefix/MoveContentWithArtbrd\tBoolean\t1\tコントロール > アートボード > オブジェクトと一緒に移動またはコピー\\nControl > Artboard > Move/Copy Artwork with Artboard",
            "Properties Panel > Artboard > Artboards > Move artwork with Artboard\tplugin/CropAreaPrefix/MoveContentWithArtbrd\tBoolean\t1\tコントロール > アートボード > オブジェクトと一緒に移動またはコピー\\nControl > Artboard > Move/Copy Artwork with Artboard",
            "プロパティパネル > アートボード > アートボード > オブジェクトと一緒に拡大・縮小\tplugin/CropAreaPrefix/ScaleContentWithArtbrd\tBoolean\t0\tコントロール > アートボード > オブジェクトと一緒に拡大・縮小\\nControl > Artboard > Scale Artwork with Artboard",
            "Properties Panel > Artboard > Artboards > Scale artwork with Artboard\tplugin/CropAreaPrefix/ScaleContentWithArtbrd\tBoolean\t0\tコントロール > アートボード > オブジェクトと一緒に拡大・縮小\\nControl > Artboard > Scale Artwork with Artboard",
            "\tplugin/CropAreaPrefix/RealTimeFading\tInteger\t1",
            "\tplugin/CropAreaPrefix/ShowFading\tInteger\t1",
            "透明パネル > 新規不透明マスクにクリッピングを適用\tplugin/TransparencyPalette/NewMasksClip\tBoolean\t1\t透明パネル > クリップ\\nTransparency Panel > Clip",
            "Transparency Panel > New Opacity Masks Are Clipping\tplugin/TransparencyPalette/NewMasksClip\tBoolean\t1\t透明パネル > クリップ\\nTransparency Panel > Clip",
            "透明パネル > 新規不透明マスクに反転を適用\tplugin/TransparencyPalette/InvertNewMasks\tBoolean\t0\t透明パネル > マスクを反転\\nTransparency Panel > Invert Mask",
            "Transparency Panel > New Opacity Masks Are Inverted\tplugin/TransparencyPalette/InvertNewMasks\tBoolean\t0\t透明パネル > マスクを反転\\nTransparency Panel > Invert Mask",
            "ファイル > 書き出し > 書き出し形式… > PNG オプション > オプション > 解像度\tplugin/PNGFileFormat/ResolutionR\tReal\t72",
            "File > Export > Export As… > PNG Options > Options > Resolution\tplugin/PNGFileFormat/ResolutionR\tReal\t72",
            "ファイル > 書き出し > 書き出し形式… > PNG オプション > オプション > アンチエイリアス\tplugin/PNGFileFormat/AntiAlias\tInteger\t1",
            "File > Export > Export As… > PNG Options > Options > Anti-aliasing\tplugin/PNGFileFormat/AntiAlias\tInteger\t1",
            "ファイル > 書き出し > 書き出し形式… > PNG オプション > オプション > インターレース\tplugin/PNGFileFormat/Interlaced\tBoolean\t0",
            "File > Export > Export As… > PNG Options > Options > Interlaced\tplugin/PNGFileFormat/Interlaced\tBoolean\t0",
            "ファイル > 書き出し > 書き出し形式… > PNG オプション > プレビュー > 背景色 > 透明\tplugin/PNGFileFormat/BackgroundTransparent\tBoolean\t1\t{true: 透明, false: それ以外の色}",
            "File > Export > Export As… > PNG Options > Preview > Background Color > Transparent\tplugin/PNGFileFormat/BackgroundTransparent\tBoolean\t1\t{true: 透明, false: それ以外の色}",
            "ファイル > 書き出し > 書き出し形式… > PNG オプション > プレビュー > 背景色 «Red»\tplugin/PNGFileFormat/Background/red\tInteger\t65535\tplugin/PNGFileFormat/BackgroundTransparentがfalseのとき有効",
            "File > Export > Export As… > PNG Options > Preview > Background Color «Red»\tplugin/PNGFileFormat/Background/red\tInteger\t65535\tplugin/PNGFileFormat/BackgroundTransparentがfalseのとき有効",
            "ファイル > 書き出し > 書き出し形式… > PNG オプション > プレビュー > 背景色 «Green»\tplugin/PNGFileFormat/Background/green\tInteger\t65535\tplugin/PNGFileFormat/BackgroundTransparentがfalseのとき有効",
            "File > Export > Export As… > PNG Options > Preview > Background Color «Green»\tplugin/PNGFileFormat/Background/green\tInteger\t65535\tplugin/PNGFileFormat/BackgroundTransparentがfalseのとき有効",
            "ファイル > 書き出し > 書き出し形式… > PNG オプション > プレビュー > 背景色 «Blue»\tplugin/PNGFileFormat/Background/blue\tInteger\t65535\tplugin/PNGFileFormat/BackgroundTransparentがfalseのとき有効",
            "File > Export > Export As… > PNG Options > Preview > Background Color «Blue»\tplugin/PNGFileFormat/Background/blue\tInteger\t65535\tplugin/PNGFileFormat/BackgroundTransparentがfalseのとき有効",
            "\tplugin/PNGFileFormat/NumberOfColors\tInteger\t0",
            "\tplugin/MTRasterize/rasterizeMemory\tInteger\t30",
            "\tplugin/ai_navigation_put/apnd\tInteger\t0",
            "\tplugin/ai_navigation_put/lowr\tInteger\t1",
            "\tplugin/ai_navigation_put/exto\tInteger\t0",
            "\tplugin/ai_navigation_put/alla\tInteger\t1",
            "\tplugin/ai_navigation_put/rstr\tString\t6",
            "\tplugin/Luckyworker Pref/ReadPDFContent\tInteger\t0",
            "\tplugin/AdobePaintStyle/GradientMeshFTUEShown\tInteger\t0",
            "\tplugin/AdobePaintStyle/ToolIntroFTUEShown\tInteger\t0",
            "\tplugin/AdobePaintStyle/ToolIntroFTUEShownCount\tInteger\t2",
            "\tplugin/3D Plugin Flattener Pref/DontReorderClipGroups\tInteger\t0",
            "\tplugin/SymbolsPopupPanel/VisibleKind\tInteger\t0",
            "\tplugin/Character Style/proxy\tInteger\t0",
            "プロパティパネル > 切り抜き > 変形 > 基準点\tplugin/Adobe Crop UI/CropAnchorPoint\tInteger\t4\t次回起動時に適用。画像の切り抜き発動時に使用。\\nコントロール > 切り抜き > 変形 > 基準点\\nControl > Cropping > Transform > Reference Point",
            "Properties Panel > Cropping > Transform > Reference Point\tplugin/Adobe Crop UI/CropAnchorPoint\tInteger\t4\t次回起動時に適用。画像の切り抜き発動時に使用。\\nコントロール > 切り抜き > 変形 > 基準点\\nControl > Cropping > Transform > Reference Point",
            "パターンオプションパネル > タイルの境界線カラー… > カラー «Red»\tplugin/AIPattern_Default_Tile_Color/Red\tInteger\t0",
            "Pattern Options Panel > Tile Edge Color… > Color «Red»\tplugin/AIPattern_Default_Tile_Color/Red\tInteger\t0",
            "パターンオプションパネル > タイルの境界線カラー… > カラー «Green»\tplugin/AIPattern_Default_Tile_Color/Green\tInteger\t0",
            "Pattern Options Panel > Tile Edge Color… > Color «Green»\tplugin/AIPattern_Default_Tile_Color/Green\tInteger\t0",
            "パターンオプションパネル > タイルの境界線カラー… > カラー «Blue»\tplugin/AIPattern_Default_Tile_Color/Blue\tInteger\t65535",
            "Pattern Options Panel > Tile Edge Color… > Color «Blue»\tplugin/AIPattern_Default_Tile_Color/Blue\tInteger\t65535",
            "効果 > ぼかし > ぼかし (ガウス)… > 半径\tplugin/PSLFilterAdapterPlugin/Radius\tReal\t30",
            "Effect > Blur > Gaussian Blur… > Radius\tplugin/PSLFilterAdapterPlugin/Radius\tReal\t30",
            "効果 > スタイライズ > ぼかし… > 半径\tplugin/FuzzyMask/Radius\tReal\t6",
            "Effect > Stylize > Feather… Radius\tplugin/FuzzyMask/Radius\tReal\t6",
            "\tplugin/FuzzyMask/Preview\tInteger\t1\t廃止？ plugin/PreviewPref/FeatherEffectPreviewPrefで制御できる",
            "オブジェクト > グラデーションメッシュを作成… > プレビュー\tplugin/CreateMesh/Preview\tBoolean\t0",
            "Object > Create Gradient Mesh… > Preview\tplugin/CreateMesh/Preview\tBoolean\t0",
            "オブジェクト > グラデーションメッシュを作成… > 行数\tplugin/CreateMesh/Rows\tInteger\t4",
            "Object > Create Gradient Mesh… > Rows\tplugin/CreateMesh/Rows\tInteger\t4",
            "オブジェクト > グラデーションメッシュを作成… > 列数\tplugin/CreateMesh/Columns\tInteger\t4",
            "Object > Create Gradient Mesh… > Columns\tplugin/CreateMesh/Columns\tInteger\t4",
            "オブジェクト > グラデーションメッシュを作成… > 種類\tplugin/CreateMesh/Shade\tInteger\t0",
            "Object > Create Gradient Mesh… > Appearance\tplugin/CreateMesh/Shade\tInteger\t0",
            "オブジェクト > グラデーションメッシュを作成… > ハイライト\tplugin/CreateMesh/Hilite\tReal\t100",
            "Object > Create Gradient Mesh… > Highlight\tplugin/CreateMesh/Hilite\tReal\t100",
            "\tplugin/com.adobe.ccx.fnft/Location and Size/l\tInteger\t305",
            "\tplugin/com.adobe.ccx.fnft/Location and Size/t\tInteger\t206",
            "\tplugin/com.adobe.ccx.fnft/Location and Size/b\tInteger\t821",
            "\tplugin/com.adobe.ccx.fnft/Location and Size/r\tInteger\t1375",
            "自動選択パネル > すべてのレイヤーを適用\tplugin/Magic Wand/UseAllLayers\tBoolean\t1",
            "Magic Wand Panel > Use All Layers\tplugin/Magic Wand/UseAllLayers\tBoolean\t1",
            "\tshowPDFOptionDialog\tInteger\t1",
            "\tmaxBridgeMRUFiles\tInteger\t30",
            "\tThreshold/DbrushPath/Print\tInteger\t30",
            "\tThreshold/DbrushPath/Save\tInteger\t30",
            "\tThreshold/DbrushPath/Clipboard\tInteger\t30",
            "\tactionwarning/oldFileFormat\tInteger\t1",
            "\tprinting/PostScript/preserveImageOrMeshFilledText\tInteger\t0",
            "\tforceExtension\tInteger\t1",
            "\tenableThreadedRendering\tInteger\t1",
            "\tNewRenderingEnabled\tInteger\t1",
            "\tFileFormatGetFile/Link\tInteger\t1",
            "\tPeriodicCopyDelay\tReal\t0.5",
            "\tExperimentalPixelMerge\tInteger\t0",
            "\tdefaultPathShown\tInteger\t1",
            "\tphotoshopFilter/excludeSpots\tInteger\t0",
            "\tphotoshopFilter/separateProcess\tInteger\t0",
            "\tImagePerformance/DisableCache\tInteger\t0",
            "オブジェクト > パス > 単純化… > 最新の設定を保持し、このダイアログを直接開く\tSimplify/skipHUD\tBoolean\t0",
            "Object > Path > Simplify… > Retain my latest settings and directly open this dialog\tSimplify/skipHUD\tBoolean\t0",
            "オブジェクト > パス > 単純化… > アンカーポイントを削減 > 曲線の単純化\tSimplify/curvePrecision\tInteger\t100",
            "Object > Path > Simplify… > Reduce Anchor Points > Simplify Curve\tSimplify/curvePrecision\tInteger\t100",
            "オブジェクト > パス > 単純化… > アンカーポイントを削減 > コーナーポイント角度のしきい値\tSimplify/angleThreshold\tInteger\t1",
            "Object > Path > Simplify… > Reduce Anchor Points > Corner Point Angle Threshold\tSimplify/angleThreshold\tInteger\t1",
            "\tresetPreferencesSet\tInteger\t0",
            "\tGPPreferences/resetPreferencesSet\tInteger\t0",
            "\tpattern/automaticBoundingBox\tBoolean\t1\tfalseにすると，パターン生成のとき枠になる透明の四角が必須になる。trueのときはアイテムから自動でサイズを取得する",
            "\tGraphDimDialogConstrainPref\tInteger\t0",
            "\tdontShowAgainOnboarding\tBoolean\t0",
            "\tsavingAsCopy\tBoolean\t0",
            "\tcloudAIAutoSaveCoachMarkAcknowledged\tInteger\t1",
            "\tAIPattern_Default_Dimming_On\tInteger\t1",
            "\tAIPattern_Default_Dim_Percent\tReal\t30",
            "\tcountRecentPPDS\tInteger\t0",
            "\tVectorize/WriteToSVgFile\tInteger\t0",
            "\tmesh/c1Tolerance\tReal\t3.0000000000000001E-3",
            "\tmesh/Flatness\tReal\t0.1",
            "\tmesh/c1Filter\tInteger\t1",
            "\trulerTyle\tInteger\t28108440",
            "\tShowExternalJSXWarning\tBoolean\t0\t外部ExtendScript実行時，警告を表示する",
            "\tSavingFromInviteToEdit\tInteger\t0",
            "コントロール > 作成および変形時にアートをピクセルグリッドに整合\tsnapToPixelOnUserAction\tBoolean\t0\t新規書類のときに切り替わる",
            "Control > Align art to pixel grid on creation and transformation\tsnapToPixelOnUserAction\tBoolean\t0\t新規書類のときに切り替わる",
            "\tAdobePlaceCloudDocumentPreference\tInteger\t0",
            "\tReplacingLinks\tBoolean\t0",
            "\tAdobeReplacingLinkWithLocalDocumentPreference\tInteger\t0",
            "\tElevatedSaveForCloudDocument\tInteger\t0",
            "\tElevatedSaveToCloud\tInteger\t0",
            "\tSavingBeforeInviteToEdit\tBoolean\t0",
            "\tSwitchingNativeToCDPDlg\tBoolean\t0",
            "\tenhanced3DEffectsOnboardingCoachMarkAcknowledged\tInteger\t0",
            "\tglobalRulersVisible\tBoolean\t1",
            "\tchanged\tInteger\t0",
            "\tAIEnableAnnotationWithAsyncPan\tInteger\t0",
            "\tAIIsModifiedPreset\tInteger\t1",
            "\tAILastUsedPDFSaveSettings\tString\t[Smallest File Size (PDF 1.6)]\t値：\\n[Smallest File Size (PDF 1.6)]",
            "\tAIShouldSuspendIdleInBackgroud\tInteger\t1",
            "\tConsistenFontReccomendation\tInteger\t0",
            "ウィンドウ > コンテキストタスクバー\tContextualTaskBarEnabled\tBoolean\t0\tv27.9では再起動時，v28からはリアルタイムで適用される",
            "Window > Contextual Task Bar\tContextualTaskBarEnabled\tBoolean\t0\tv27.9では再起動時，v28からはリアルタイムで適用される",
            "\tDimensionToolForceInjected\tInteger\t1",
            "\tEnableJSONForClipboardCopy\tInteger\t1",
            "\tEnableJSONForClipboardPaste\tInteger\t1",
            "\tExportSettings/browselocation\tString\t",
            "\tGenAI/AILegalTermsAccepted_V2\tInteger\t1",
            "\tGradientColorStopDialogPref\tInteger\t0",
            "\tGradientToolUsedOnce\tInteger\t1",
            "\tHistoryPanelOpenedOnce\tInteger\t0",
            "\tIntertwineToolUsedOnce\tInteger\t0",
            "\tNumOfTimesSFROnboardingShownPref\tInteger\t0",
            "\tOnboarding/OnShapeCreation/FirstCard\tInteger\t1",
            "\tOnboarding/OnShapeCreation/SecondCard\tInteger\t0",
            "\tOnboarding/OnShapeCreation/ThirdCard\tInteger\t0",
            "\tOnboarding/SampleFile/FirstCard\tInteger\t0",
            "\tOnboarding/UserFile/FirstCard\tInteger\t0",
            "\tRecolorArtworkBlueDotShownCount\tInteger\t3",
            "\tResponsive_Zoom_EPF_Limit\tReal\t6.6666666667",
            "\tS4RSharesheetOnboardingPref\tInteger\t0",
            "\tSFRDontShowModalIntercepts\tInteger\t0",
            "\tSFRInterceptExportAsStatePref\tInteger\t0",
            "\tSFRInterceptRestartCountSuffix\tInteger\t0",
            "\tSFRInterceptSaveAsPDFStatePref\tInteger\t0",
            "\tSFRInterceptShownTotalCountPref\tInteger\t0",
            "\tSFRInterceptStatePref\tInteger\t0",
            "\tSFRJPGPNGInterceptRandomContentSeqToShow\tInteger\t0",
            "\tSFRLastInterceptShownTS\tInteger\t0",
            "\tSFRLastReviewCreatedUpdatedTS\tInteger\t0",
            "\tSFRLinkCreatedPref\tInteger\t0",
            "\tSFRPDFInterceptRandomContentSeqToShow\tInteger\t0",
            "\tSVGWriterCPPSave\tInteger\t0",
            "\tSettingsToolTip\tInteger\t1",
            "\tSimplify/smoothingPercentage\tReal\t0",
            "オブジェクト > 変形 > 個別に変形… > 縦横比を固定\tTransformEachConstrain\tBoolean\t0",
            "Object > Transform > Transform Each… > Constrain Width and Height Proportions\tTransformEachConstrain\tBoolean\t0",
            "\tenableConcurrentEditing\tInteger\t0",
            "\tgenerateView\tInteger\t1",
            "\tisCurvatureToolApplied\tInteger\t0",
            "\tplugin/Adobe ReType/ContextualCoachShownCount\tInteger\t3",
            "\tplugin/Adobe ReType/OnDemandPopupFirstLaunchFromMode\tInteger\t1",
            "\tplugin/Adobe ReType/RetypePanelFirstLaunch\tInteger\t1",
            "オブジェクト > リピート > リピートオプション… > ラジアル > インスタンス数\tplugin/Adobe Repeat Arts UI/RadialInstances\tInteger\t8",
            "Object > Repeat > Options… > Radial > Number of instances\tplugin/Adobe Repeat Arts UI/RadialInstances\tInteger\t8",
            "オブジェクト > リピート > リピートオプション… > ラジアル > 重なりを反転\tplugin/Adobe Repeat Arts UI/ReverseOverlapCheckbox\tBoolean\t0",
            "Object > Repeat > Options… > Radial > Reverse Overlap\tplugin/Adobe Repeat Arts UI/ReverseOverlapCheckbox\tBoolean\t0",
            "オブジェクト > リピート > リピートオプション… > グリッド > グリッドの水平方向の間隔\tplugin/Adobe Repeat Arts UI/HorizontalSpacing\tReal\t40",
            "Object > Repeat > Options… > Grid > Horizontal spacing in grid\tplugin/Adobe Repeat Arts UI/HorizontalSpacing\tReal\t40",
            "オブジェクト > リピート > リピートオプション… > グリッド > グリッドの垂直方向の間隔\tplugin/Adobe Repeat Arts UI/VerticalSpacing\tReal\t10",
            "Object > Repeat > Options… > Grid > Vertical spacing in grid\tplugin/Adobe Repeat Arts UI/VerticalSpacing\tReal\t10",
            "オブジェクト > リピート > リピートオプション… > グリッド > グリッドの種類\tplugin/Adobe Repeat Arts UI/GridPatternType\tInteger\t1",
            "Object > Repeat > Options… > Grid > Grid Type\tplugin/Adobe Repeat Arts UI/GridPatternType\tInteger\t1",
            "オブジェクト > リピート > リピートオプション… > グリッド > 行を反転 > 水平方向に反転\tplugin/Adobe Repeat Arts UI/FlipRowY\tBoolean\t0",
            "Object > Repeat > Options… > Grid > Flip Rows > Flip horizontal\tplugin/Adobe Repeat Arts UI/FlipRowY\tBoolean\t0",
            "オブジェクト > リピート > リピートオプション… > グリッド > 行を反転 > 垂直方向に反転\tplugin/Adobe Repeat Arts UI/FlipRowX\tBoolean\t0",
            "Object > Repeat > Options… > Grid > Flip Rows > Flip vertical\tplugin/Adobe Repeat Arts UI/FlipRowX\tBoolean\t0",
            "オブジェクト > リピート > リピートオプション… > グリッド > 列を反転 > 水平方向に反転\tplugin/Adobe Repeat Arts UI/FlipColumnY\tBoolean\t0",
            "Object > Repeat > Options… > Grid > Flip Column > Flip horizontal\tplugin/Adobe Repeat Arts UI/FlipColumnY\tBoolean\t0",
            "オブジェクト > リピート > リピートオプション… > グリッド > 列を反転 > 垂直方向に反転\tplugin/Adobe Repeat Arts UI/FlipColumnX\tBoolean\t0",
            "Object > Repeat > Options… > Grid > Flip Column > Flip vertical\tplugin/Adobe Repeat Arts UI/FlipColumnX\tBoolean\t0",
            "オブジェクト > リピート > リピートオプション… > ミラー > ミラー軸の角度\tplugin/Adobe Repeat Arts UI/AxisRotationAngle\tReal\t90",
            "Object > Repeat > Options… > Mirror > Angle of mirror axis\tplugin/Adobe Repeat Arts UI/AxisRotationAngle\tReal\t90",
            "\tplugin/AdobePaintStyle/Color Dialog Color Space\tInteger\t3",
            "\tplugin/DimensionUI/AssisstiveCount\tInteger\t11",
            "\tplugin/DimensionUI/DimensionCount\tInteger\t16",
            "\tplugin/DimensionUI/DimensionFeedbackShown\tInteger\t1",
            "\tplugin/DimensionUI/DimensionToolInvokedCount\tInteger\t1",
            "\tplugin/DimensionUI/ManualCount\tInteger\t1",
            "\tplugin/DimensionUI/RadialInvokedCount\tInteger\t1",
            "\tplugin/DimensionUI/ToolOptionsInvoked\tInteger\t1",
            "\tplugin/ExperimentationPrefix/AILightHandled\tInteger\t0",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > 画像圧縮 > 可逆圧縮 «真偽値»\tplugin/ImageDecoderEncoder/AIWebPEncodeLosslessCompression\tBoolean\t1\t動作しない？ {true: 可逆圧縮, false: 非可逆圧縮}\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Image Compression > Lossless «Boolean»\tplugin/ImageDecoderEncoder/AIWebPEncodeLosslessCompression\tBoolean\t1\t動作しない？ {true: 可逆圧縮, false: 非可逆圧縮}\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > 画像圧縮 > 非可逆圧縮 > 画質\tplugin/ImageDecoderEncoder/AIWebPEncodeImageQuality\tInteger\t100\t動作しない？ \\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Image Compression > Lossy > Quality\tplugin/ImageDecoderEncoder/AIWebPEncodeImageQuality\tInteger\t100\t動作しない？ \\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > オプション > アンチエイリアス\tplugin/ImageDecoderEncoder/AIWebPEncodeAntiAlias\tInteger\t2\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Options > Anti-aliasing\tplugin/ImageDecoderEncoder/AIWebPEncodeAntiAlias\tInteger\t2\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > オプション > 背景色 > 透明 «真偽値»\tplugin/ImageDecoderEncoder/AIWebPEncodeIsTransparent\tBoolean\t1\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Options > Background Color > Transparent «Boolean»\tplugin/ImageDecoderEncoder/AIWebPEncodeIsTransparent\tBoolean\t1\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > オプション > 背景色 «Red»\tplugin/ImageDecoderEncoder/AIWebPEncodeMatteColorRed\tInteger\t65535\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Options > Background Color «Red»\tplugin/ImageDecoderEncoder/AIWebPEncodeMatteColorRed\tInteger\t65535\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > オプション > 背景色 «Green»\tplugin/ImageDecoderEncoder/AIWebPEncodeMatteColorGreen\tInteger\t65535\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Options > Background Color «Green»\tplugin/ImageDecoderEncoder/AIWebPEncodeMatteColorGreen\tInteger\t65535\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > オプション > 背景色 «Blue»\tplugin/ImageDecoderEncoder/AIWebPEncodeMatteColorBlue\tInteger\t65535\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Options > Background Color «Blue»\tplugin/ImageDecoderEncoder/AIWebPEncodeMatteColorBlue\tInteger\t65535\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > ICC プロファイルを埋め込む\tplugin/ImageDecoderEncoder/AIWebPEncodeEmbedICCProfile\tBoolean\t1\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Options > Embed ICC Profile\tplugin/ImageDecoderEncoder/AIWebPEncodeEmbedICCProfile\tBoolean\t1\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > スクリーン用に書き出し… > 形式の設定 > WebP > XMP メタデータを含める\tplugin/ImageDecoderEncoder/AIWebPEncodeIncludeXMPMetadata\tBoolean\t0\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "File > Export > Export for Screens… > Format Settings > WebP > Options > Include XMP Metadata\tplugin/ImageDecoderEncoder/AIWebPEncodeIncludeXMPMetadata\tBoolean\t0\t動作しない？\\nファイル > 書き出し > 書き出し形式… と設定を共有する",
            "ファイル > 書き出し > 書き出し形式… > WebP オプション > オプション > 解像度\tplugin/ImageDecoderEncoder/AIWebPEncodePPI\tReal\t72.0\t動作しない？",
            "File > Export > Export As… > WebP Options > Options > Resolution\tplugin/ImageDecoderEncoder/AIWebPEncodePPI\tReal\t72.0\t動作しない？",
            "\tplugin/OnBoarding/IntertwineOnBoardingShownCount\tInteger\t2",
            "\tplugin/OnBoarding/PantoneLibOnBoardingShownCountAugust2023\tInteger\t0",
            "\tplugin/OnBoarding/PantoneLibOnBoardingShownCountNovember2022\tInteger\t0",
            "\tplugin/OnBoarding/RecolorArtworkTooltipShownCount\tInteger\t1",
            "\tplugin/OnBoarding/SelectSaveTooltipShownCount\tInteger\t1",
            "\tplugin/OnBoarding/VectorEdgeOnBoardingShownCount\tInteger\t1",
            "\tplugin/OnBoarding/VectorEdgeRichMediaCardShownCount\tInteger\t1",
            "\tplugin/OnBoarding/dontShowAgainPantoneLibOnBoardingAugust2023\tInteger\t0",
            "\tplugin/OnBoarding/dontShowAgainPantoneLibOnBoardingNovember2022\tInteger\t0",
            "\tplugin/PDFExport/AI11PDF_IncludeHyperlinks\tBoolean\t1",
            "\tplugin/PDFImport/UseInMemoryPDFLRepair\tInteger\t0",
            "ファイル > 書き出し > 書き出し形式… > TIFF オプション > カラーモード\tplugin/TIFFFileFormat/ColorModel\tInteger\t1",
            "File > Export > Export As… > TIFF Options > Color Model\tplugin/TIFFFileFormat/ColorModel\tInteger\t1",
            "ファイル > 書き出し > 書き出し形式… > TIFF オプション > 解像度\tplugin/TIFFFileFormat/DPI\tReal\t300",
            "File > Export > Export As… > TIFF Options > Resolution\tplugin/TIFFFileFormat/DPI\tReal\t300",
            "ファイル > 書き出し > 書き出し形式… > TIFF オプション > アンチエイリアス\tplugin/TIFFFileFormat/AntiAlias\tInteger\t2",
            "File > Export > Export As… > TIFF Options > Anti-aliasing\tplugin/TIFFFileFormat/AntiAlias\tInteger\t2",
            "ファイル > 書き出し > 書き出し形式… > TIFF オプション > LZW 圧縮\tplugin/TIFFFileFormat/LZWCompression\tBoolean\t1",
            "File > Export > Export As… > TIFF Options > LZW Compression\tplugin/TIFFFileFormat/LZWCompression\tBoolean\t1",
            "\tplugin/TIFFFileFormat/ByteOrder\tInteger\t2\t無視される？ 常にTIFFByteOrder.IBMPC。ファイルの先頭がIIならIBMPC、MMならMACINTOSH",
            "\tplugin/TIFFFileFormat/PreserveSpotColors\tInteger\t0",
            "\tplugin/ai_navigation_put/suffixstr\tInteger\t0",
            "\tplugin/ai_navigation_put/unembedAll\tInteger\t0",
            "\tplugin/snapping/disableSelfSnapping\tInteger\t1",
            "\ttypeConversion/NumConversionMessageShown\tInteger\t1",
            "一般 > 裁ち落とし部分に「裁ち落としを印刷」生成 AI ボタンを表示\tenablePrintBleedWidget\tBoolean\t0",
            "General > Show 'Print Bleed' generative AI buttons on Bleed\tenablePrintBleedWidget\tBoolean\t0",
            "\tArtboardBBColorRed\tReal\t0.0\tアートボードの境界線の色（RGB 0〜1）。例：ライトブルー 0.29 / 0.52 / 1.0、ライトレッド 1.0 / 0.29 / 0.29、グリーン 0 / 0.65 / 0.31。Red・Green・Blue の3キーをそろえて書く（PresetManager.jsx）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "\tArtboardBBColorGreen\tReal\t0.0\tアートボードの境界線の色（RGB 0〜1）。Red・Green・Blue の3キーをそろえて書く（PresetManager.jsx）",
            "\tArtboardBBColorBlue\tReal\t0.0\tアートボードの境界線の色（RGB 0〜1）。Red・Green・Blue の3キーをそろえて書く（PresetManager.jsx）",
            "\tArtboardBBWidth\tReal\t1.0\tアートボードの境界線の幅（PresetManager.jsx の選択肢は 1 / 2 / 3 / 4）\\n書き込み後は app.redraw() だけでは画面に反映されないことがある（PresetManager.jsx は zoomout → zoomin で再描画）",
            "\tshowArtboardLabelOnCanvas\tInteger\t1\tアートボード名をカンバスに表示（PresetManager.jsx）",
            "\tenableEnclosedMode\tInteger\t0",
            "\texcludeEffectRegionsFromHitTest\tInteger\t0",
            "\tuiShareButtonIsBlue\tInteger\t0",
            "\tshowColorSamplingRing\tInteger\t1",
            "\tsmartGuides/sensitivity\tInteger\t1",
            "\tsmartGuides/cursorSnapping\tInteger\t0",
            "\tsmartGuides/snapToIsolatedObjects\tInteger\t1",
            "\tsmartGuides/invokeDistanceGuides\tInteger\t1",
            "\talignmentGuides/showSnapToGuide\tInteger\t1",
            "\talignmentGuides/showArtboardGuides\tInteger\t1",
            "\talignmentGuides/centerpoint\tInteger\t1",
            "\talignmentGuides/midpoint\tInteger\t1",
            "\talignmentGuides/endpoint\tInteger\t1",
            "\tsnapomatic/ConstructionGuideColor/red\tReal\t0.180392161",
            "\tsnapomatic/ConstructionGuideColor/green\tReal\t0.1921568662",
            "\tsnapomatic/ConstructionGuideColor/blue\tReal\t0.5725490451",
            "\tPerformance/FlickPan\tInteger\t0",
            "\tPerformance/TurnOffGPUDueToCrash\tInteger\t0",
            "\tfSnapToGrid\tInteger\t1",
            "\tAIDisableAsyncRendering\tInteger\t0",
            "\tplugin/Adobe Constraints UI/PivotPointPreference\tInteger\t4",
            "\tplugin/Adobe Constraints UI/RotationAnglePreference\tReal\t0.0",
            "\tplugin/Adobe Constraints UI/CopyPasteToastShown\tInteger\t0",
            "\tplugin/Magic Wand/BlendModeEnabled\tInteger\t0",
            "\tplugin/Magic Wand/OpacityEnabled\tInteger\t0",
            "\tplugin/Magic Wand/StrokeWeightEnabled\tInteger\t0",
            "\tplugin/Magic Wand/StrokeColorEnabled\tInteger\t0",
            "\tplugin/Magic Wand/FillColorEnabled\tInteger\t1",
            "\tplugin/Magic Wand/CMYKStrokeColorTolerance\tInteger\t20",
            "\tplugin/Magic Wand/CMYKFillColorTolerance\tInteger\t20",
            "\tplugin/Magic Wand/RGBStrokeColorTolerance\tInteger\t32",
            "\tplugin/Magic Wand/RGBFillColorTolerance\tInteger\t32",
            "\tplugin/Magic Wand/OpacityTolerance\tReal\t0.0500000007",
            "\tplugin/Magic Wand/StrokeWeightTolerance\tReal\t5.0",
            "\tplugin/Adobe Measure Tool/MeasureTool_Precision\tInteger\t2",
            "\tplugin/Adobe Measure Tool/MeasureTool_CustomScale_Count\tInteger\t0",
            "\tplugin/Adobe Measure Tool/MeasureTool_CustomScale_ReplacementIndex\tInteger\t0",
            "\tplugin/Adobe Measure Tool/MeasureTool_Scale\tInteger\t0",
            "\tplugin/Adobe Measure Tool/MeasureTool_Units\tInteger\t0",
            "書式 > エリア内文字オプション...\tareatextoptions\t旧 ID：area-type-options",
            "その他 > コピー (2)\tcopy2",
            "その他 > カット (2)\tcut2",
            "その他 > ヘルプ (2)\thelpcontent2",
            "その他 > ペースト (2)\tpaste2",
            "ファイル > テンプレートとして保存...\tsaveasTemplate\t旧 ID：saveastemplate",
            "編集 > 環境設定 > 選択範囲・アンカー表示...\tselectionPref\t旧 ID：selectPref",
            "Illustrator > 環境設定 > 選択範囲・アンカー表示...\tselectionPref\t旧 ID：selectPref",
            "選択範囲・アンカー表示\tselectionPref\t旧 ID：selectPref",
            "ヘルプ > システム情報...\tsystemInfo\t旧 ID：System Info",
            "書式 > パス上文字オプション > 3D リボン\ttextpathtype3d\t旧 ID：3D ribbon",
            "書式 > パス上文字オプション > 引力\ttextpathtypeGravity\t旧 ID：Gravity",
            "書式 > パス上文字オプション > パス上文字オプション...\ttextpathtypeOptions\t旧 ID：typeOnPathOptions",
            "書式 > パス上文字オプション > 虹\ttextpathtypeRainbow\t旧 ID：Rainbow",
            "書式 > パス上文字オプション > 歪み\ttextpathtypeSkew\t旧 ID：Skew",
            "書式 > パス上文字オプション > 階段状\ttextpathtypestairs\t旧 ID：Stair Step",
            "その他 > 取り消し (2)\tundo2",
            "編集 > 環境設定 > ユーザーインターフェイス...\tuserInterfacePref\t旧 ID：UIPref",
            "Illustrator > 環境設定 > ユーザーインターフェイス...\tuserInterfacePref\t旧 ID：UIPref",
            "ユーザーインターフェイス\tuserInterfacePref\t旧 ID：UIPref",
            "その他 > ズームイン (2)\tzoomin2",
            "その他 > 下付き文字 (2)\t~subscript2",
            "その他 > 上付き文字 (2)\t~superScript2",
            "Other Misc > Copy (Secondary)\tcopy2",
            "Other Misc > Cut (Secondary)\tcut2",
            "Other Misc > Help (Secondary)\thelpcontent2",
            "Other Misc > Paste (Secondary)\tpaste2",
            "Selection & Anchor Display\tselectionPref\t旧 ID：selectPref",
            "Other Misc > Undo (Secondary)\tundo2",
            "User Interface\tuserInterfacePref\t旧 ID：UIPref",
            "Other Misc > Zoom In (Secondary)\tzoomin2",
            "Other Misc > Subscript (Secondary)\t~subscript2",
            "Other Misc > Superscript (Secondary)\t~superScript2",
            "\tAdobe Presentation Mode",
            "\tAdobeAlignHorizVertCenterToArtboardOther",
            "\tApply Last Filter",
            "\tExportSettings",
            "\tImportSettings",
            "\tLast Filter",
            "\tOpenGLCompositorPreview",
            "\tRotateView120",
            "\tRotateView135",
            "\tRotateView15",
            "\tRotateView150",
            "\tRotateView180",
            "\tRotateView30",
            "\tRotateView45",
            "\tRotateView60",
            "\tRotateView90",
            "\tRotateViewNegative120",
            "\tRotateViewNegative135",
            "\tRotateViewNegative15",
            "\tRotateViewNegative150",
            "\tRotateViewNegative180",
            "\tRotateViewNegative30",
            "\tRotateViewNegative45",
            "\tRotateViewNegative60",
            "\tRotateViewNegative90",
            "\tRotateViewZero",
            "\tappframe",
            "\tapplicationbar",
            "\tbringAllToFront",
            "\tcloseAll2",
            "\tcollectForExportMultipleAsset",
            "\tcollectForExportSingleAsset",
            "\tconvertlegacyText",
            "\tconvertlegacyText1",
            "\tconvertlegacyText2",
            "\tconvertlegacyText3",
            "\tconvertlegacyText4",
            "\tdebugPalette",
            "\tglyphSnapping",
            "\thideApp",
            "\thideOthers",
            "\tminimizeWindow",
            "\topenInFFBoards",
            "\tpixelconstraints",
            "\tresetRotationView",
            "\trotateViewToSelection",
            "\tshowAllWindows",
            "\tsupportContent",
            "\tswitchUnits",
            "\tview1",
            "\tview10",
            "\tview2",
            "\tview3",
            "\tview4",
            "\tview5",
            "\tview6",
            "\tview7",
            "\tview8",
            "\tview9",
            "\t~nonBreakingSpace",
            "\t~placeHolderText",
            "\tAdobe Ignore Color Tool",
            "\tRearrange Artboard Tool",
            "\tAutoInstantSave/IdleLoopTimeInterval\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tDisableHoverScrollOnUnfocused\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tforceSnapToGrid\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\thighContrastEnabled\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tShow/HideRecentFonts\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tShapeCoreUI/LiveShape/NumStarPoints\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/Adobe Freehand Tool/round_caps\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/AdobePaintStyle/gradientPresetPopupViewType\tInteger\t3\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/AdobePaintStyle/gradientPresetViewType\tInteger\t3\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/AdobeSwatchPopup_Fill/ShowRecentColors\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/AdobeSwatchPopup_Stroke/ShowRecentColors\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/AdobeSwatch_/ShowRecentColors\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/ArtboardColor/Kind\tInteger\t6\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/ContextualTaskBar/Pinned\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/ContextualTaskBar/PositionH\tReal\t0.0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/ContextualTaskBar/PositionV\tReal\t0.0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SmartExport/AIFormatSmartExport\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SwatchPalettePrefix/HideBackground\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SwatchPalettePrefix/LinkDimension\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SwatchPalettePrefix/ShowColorInfo\tInteger\t255\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SwatchPalettePrefix/SwatchChipHeight\tReal\t100.0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SwatchPalettePrefix/SwatchChipMarginHorizontal\tReal\t10.0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SwatchPalettePrefix/SwatchChipMarginVertical\tReal\t10.0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SwatchPalettePrefix/SwatchChipWidth\tReal\t100.0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/SymbolPalette/NewSymbol/SymbolType\tInteger\t2\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/svgOMGOptionDlg/O_RasterFormat\tInteger\t4\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tplugin/svgOMGOptionDlg/O_RasterResolution\tInteger\t72\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tshowHelpBar\tInteger\t0\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tshowSnapping\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tshowVisualGuidesForSnapToGrid\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\tsnapToTangentPerpendicularParallel\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\ttext/enableAutoFontDownloadForGenAI\tInteger\t1\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "\ttolerance\tInteger\t6\tAi 30.8.2 の再調査で追加（名前・値の意味は未確認）",
            "オブジェクト > 裁ち落としを印刷...\tGen Bleed Object Menu\t廃止：Ai 29.8 まで（Ai 29.6 から）",
            "Object > Print Bleed...\tGen Bleed Object Menu\t廃止：Ai 29.8 まで（Ai 29.6 から）",
            "ファイル > ベクターを生成\tGenerate Modal File Menu \t廃止：Ai 29.8 まで（Ai 29.5 から）",
            "File > Generate Vectors\tGenerate Modal File Menu \t廃止：Ai 29.8 まで（Ai 29.5 から）"
        ];
    }

}());
