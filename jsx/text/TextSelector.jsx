#target illustrator
#targetengine "TextSelectorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.TextSelector
Meta-Default-language=ja
Name=複数の条件でテキストフレームを一括選択
Name[en]=Select text frames by a combination of conditions
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.4
Release-Date=2026-05-09
Tags=text,テキスト
X-com.swwwitch-Script-Name=TextSelector
X-com.swwwitch-Updated=2026-11-20
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextSelector.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextSelector.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/n76f1e0937088
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextSelector.jsx
Description-BEGIN
[ja]
ドキュメント内のテキストフレームを、複数の条件で一括選択します。
フォント属性・テキストの種類・文字列を組み合わせて絞り込み、選択後に非表示やレイヤー移動、一括編集も行えます。
[en]
Selects text frames across the document by a combination of conditions.
Combine font attributes, text type and string to narrow it down, then hide, move to a layer, or bulk-edit what was selected.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「選択後にレイヤーへ移動」先のレイヤー名の初期値 / Default destination layer for "move to layer after selection" */
    var DEFAULT_MOVE_LAYER_NAME = "_text";

    /* 一括編集の欄での強制改行の代替文字 / Placeholder for forced line breaks in the bulk edit field */
    var SOFT_BREAK = "@#";

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

    var ATTRIBUTE_LABEL_WIDTH = 120;               /* 属性チェックボックスの幅 / width of the attribute checkboxes */
    var PREVIEW_CHARACTERS = 16;                   /* 属性プレビューの文字数 / width of an attribute preview */
    var KEYWORD_CHARACTERS = 15;                   /* 検索文字列欄の文字数 / width of the keyword field */
    var MOVE_LAYER_FIELD_CHARACTERS = 12;          /* 移動先レイヤー名の欄の文字数 / width of the destination layer field */
    var MOVE_LAYER_FIELD_INDENT = 20;              /* 移動先レイヤー名の欄の字下げ / indent of the destination layer field */
    var KEYWORD_FIELD_HEIGHT = 60;                 /* 検索文字列欄の高さ / height of the keyword field */
    var TARGET_COUNT_CHARACTERS = 20;              /* 対象テキスト数の表示幅 / width of the target count text */
    var BULK_INPUT_SIZE = [300, 70];               /* 一括編集の入力欄 / size of the bulk edit field */
    var INSERT_BUTTON_HEIGHT = 20;                 /* 改行の挿入ボタンの高さ / height of the break insert buttons */
    var INSERT_FONT_SHRINK = 2;                    /* 挿入ボタンの文字を小さくする量（pt）/ how much smaller the insert button font is (pt) */

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

    var LABELS = {
        dialog: {
            title: { ja: "テキストを選択", en: "Select Text" },
            bulkEditTitle: { ja: "テキストを一括編集", en: "Bulk Edit Text" }
        },
        panel: {
            selection: { ja: "選択条件", en: "Selection Criteria" },
            artboardScope: { ja: "アートボード", en: "Artboard" },
            attribute: { ja: "フォント属性", en: "Font Attributes" },
            textType: { ja: "テキストの種類", en: "Text Type" },
            textMatch: { ja: "文字列", en: "String" },
            postProcess: { ja: "選択後の処理", en: "After Selection" }
        },
        radio: {
            artboardAll: { ja: "すべて", en: "All" },
            artboardCurrent: { ja: "現在のアートボードのみ", en: "Current Artboard Only" },
            noTextMatch: { ja: "なし", en: "None" },
            exactMatch: { ja: "完全一致", en: "Exact Match" },
            containsMatch: { ja: "部分一致", en: "Partial Match" },
            startsWith: { ja: "先頭一致", en: "Starts With" },
            endsWith: { ja: "末尾一致", en: "Ends With" },
            regexMatch: { ja: "正規表現", en: "Regular Expression" },
            noPostProcess: { ja: "なし（選択のみ）", en: "None (Select Only)" },
            hideAfterSelection: { ja: "選択したテキストを非表示", en: "Hide Selected" },
            hideOthers: { ja: "選択した以外を非表示", en: "Hide All but Selected" },
            moveToLayer: { ja: "レイヤーへ移動", en: "Move to Layer" },
            bulkEdit: { ja: "一括編集", en: "Bulk Edit" }
        },
        checkbox: {
            fontFamily: { ja: "ファミリー", en: "Family" },
            fontStyle: { ja: "スタイル", en: "Style" },
            fontSize: { ja: "フォントサイズ", en: "Font Size" },
            textFillColor: { ja: "塗りカラー", en: "Fill Color" },
            pointText: { ja: "ポイント文字", en: "Point Text" },
            areaText: { ja: "エリア内文字", en: "Area Text" },
            pathText: { ja: "パス上文字", en: "Path Text" }
        },
        button: {
            paragraphBreak: { ja: "改行", en: "Paragraph Break" },
            lineBreak: { ja: "強制改行", en: "Forced Line Break" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        attributeName: {
            fontFamily: { ja: "フォントファミリー", en: "font family" },
            fontStyle: { ja: "スタイル", en: "font style" },
            fontSize: { ja: "フォントサイズ", en: "font size" },
            textFillColor: { ja: "塗りカラー", en: "fill color" }
        },
        fieldLabel: {
            targetCount: { ja: "対象テキスト数", en: "Target texts" }
        },
        tooltip: {
            bulkInput: {
                ja: "Shift+Enter で強制改行を入れます（{softBreak} と表示されます）",
                en: "Shift+Enter inserts a forced line break (shown as {softBreak})"
            },
            paragraphBreak: { ja: "カーソルの位置に改行を入れます", en: "Inserts a paragraph break at the cursor" },
            lineBreak: {
                ja: "カーソルの位置に強制改行（{softBreak}）を入れます（Shift+Enter）",
                en: "Inserts a forced line break ({softBreak}) at the cursor (Shift+Enter)"
            },
            attributeMatch: {
                ja: "選択中のテキストと{attribute}が同じテキストを選択します。オンにした属性はすべて一致するものだけを選びます。Option＋クリックですべてオン、もう一度でそれ以外をオン。",
                en: "Select text with the same {attribute} as the selected text. Every checked attribute must match. Option-click to check all; Option-click again to check all but this one."
            },
            pointText: { ja: "ポイント文字を対象にします。［フォント属性］［文字列］と組み合わせられます。", en: "Include point text. Can be combined with Font Attributes and String." },
            areaText: { ja: "エリア内文字を対象にします。［フォント属性］［文字列］と組み合わせられます。", en: "Include area text. Can be combined with Font Attributes and String." },
            pathText: { ja: "パス上文字を対象にします。［フォント属性］［文字列］と組み合わせられます。", en: "Include path text. Can be combined with Font Attributes and String." },
            keywordInput: {
                ja: "検索対象にする文字列を入力します。［なし］以外の条件を選んでいるときは、空欄では実行できません。テキストを複数選択しているときは使えません。",
                en: "Enter the search string. It cannot be empty when a condition other than None is chosen. Unavailable when several texts are selected."
            },
            noTextMatch: { ja: "文字列では絞り込みません。", en: "Do not filter by string." },
            targetCount: { ja: "今の条件で選ばれるテキストの数です。OK で選択されるのはこの数です。", en: "Number of texts the current conditions select. OK selects exactly these." },
            attributePanel: {
                ja: "テキストを選択してから実行すると使えます。選択中のテキストのどれかと、オンにした属性がすべて一致するテキストを選びます。",
                en: "Available when text is selected before running. Selects text whose checked attributes all match one of the selected texts."
            },
            textTypePanel: { ja: "オンにした種類のテキストだけを対象にします。すべてオフにすると対象は 0 件です。", en: "Only the checked kinds of text are targeted. With all off, nothing is targeted." },
            exactMatch: {
                ja: "入力した文字列と完全に一致するテキストを選択します。大文字と小文字を区別します。",
                en: "Select text that exactly matches the entered string. Case-sensitive."
            },
            startsWith: {
                ja: "入力した文字列で始まるテキストを選択します。大文字と小文字を区別します。",
                en: "Select text that starts with the entered string. Case-sensitive."
            },
            endsWith: {
                ja: "入力した文字列で終わるテキストを選択します。大文字と小文字を区別します。",
                en: "Select text that ends with the entered string. Case-sensitive."
            },
            containsMatch: {
                ja: "入力した文字列を含むテキストを選択します。大文字と小文字を区別します。",
                en: "Select text that contains the entered string. Case-sensitive."
            },
            regexMatch: {
                ja: "入力した正規表現に一致するテキストを選択します。大文字と小文字を区別します。",
                en: "Select text that matches the entered regular expression. Case-sensitive."
            },
            artboardAll: { ja: "ドキュメント内のすべてのテキストを対象にします。", en: "Target text across the whole document." },
            artboardCurrent: {
                ja: "現在のアートボードに重なるテキストだけを対象にします。",
                en: "Target only text that overlaps the active artboard."
            },
            noPostProcess: { ja: "テキストを選択するだけで、ほかの処理は行いません。", en: "Only select the text; do nothing else." },
            hideAfterSelection: { ja: "選択されたテキストを非表示にします。", en: "Hide the selected text." },
            hideOthers: {
                ja: "選択されたテキスト以外のオブジェクトを非表示にします。ロックされたオブジェクトは除きます。",
                en: "Hide all objects except the selected text. Locked objects are left untouched."
            },
            moveToLayer: {
                ja: "選択されたテキストを、下の欄に入力したレイヤーへ移動します。",
                en: "Move the selected text to the layer named below."
            },
            moveLayerName: {
                ja: "移動先のレイヤー名です。同じ名前のレイヤーがあればそこへ移動し、なければ作ります。空欄なら「_text」になります。",
                en: "Name of the destination layer. Moves to an existing layer with this name, or creates one. Uses “_text” when empty."
            },
            bulkEdit: {
                ja: "選択されたテキストの内容をまとめて置き換えます。書式は元の文字の位置ごとに引き継ぎます。［文字列］で［なし］以外の条件を選んでいるときに使えます。",
                en: "Replace the contents of all selected text frames at once. Formatting is carried over by character position. Available when a String condition other than None is chosen."
            }
        },
        preview: {
            mixed: { ja: "混在", en: "Mixed" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Open a document before running this script." },
            emptyKeyword: { ja: "検索文字列を入力してください。", en: "Enter a search string." },
            invalidRegex: { ja: "正規表現が正しくありません。", en: "The regular expression is invalid." },
            noMatch: { ja: "条件に一致するテキストが見つかりませんでした。", en: "No text matched the criteria." }
        }
    };

    // =========================================
    // 選択中のテキストの情報 / Current text selection
    // =========================================

    /**
     * 選択（配列・単体・TextRange）を配列にそろえる
     * @param {Object} docSelection - doc.selection の値
     * @returns {Array} 選択中のオブジェクトの配列
     */
    function collectSelectionAsArray(docSelection) {
        var selectedItems = [];
        if (!docSelection) {
            return selectedItems;
        }

        if (docSelection.typename) {
            selectedItems.push(docSelection);
            return selectedItems;
        }

        if (typeof docSelection.length === "number") {
            for (var i = 0; i < docSelection.length; i++) {
                selectedItems.push(docSelection[i]);
            }
            return selectedItems;
        }

        selectedItems.push(docSelection);
        return selectedItems;
    }

    /**
     * 選択中のオブジェクトからテキスト範囲を取り出す
     * @param {Object} selectionItem - 選択中のオブジェクト
     * @returns {TextRange|null} テキスト範囲（無ければ null）
     */
    function getTextRangeFromSelectionItem(selectionItem) {
        if (!selectionItem) {
            return null;
        }

        /* テキスト以外ではプロパティの参照で例外になることがある / non-text items may throw on these properties */
        try {
            if (selectionItem.typename === "TextFrame") {
                return selectionItem.textRange;
            }
            if (selectionItem.characterAttributes) {
                return selectionItem;
            }
            if (selectionItem.textRange) {
                return selectionItem.textRange;
            }
        } catch (textRangeError) {
            return null;
        }

        return null;
    }


    /**
     * 選択中の最初のテキストの文字列を返す（検索文字列の初期値）
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {string} 文字列（無ければ空文字）
     */
    function getSelectedTextString(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            var textRange = getTextRangeFromSelectionItem(selectedItems[i]);
            if (textRange && textRange.contents) {
                return textRange.contents;
            }
        }
        return "";
    }

    /**
     * 選択中のオブジェクトのうち、テキストとして読めるものの数を返す
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {number} テキストの数
     */
    function countSelectedTexts(selectedItems) {
        var textCount = 0;
        for (var i = 0; i < selectedItems.length; i++) {
            if (getTextRangeFromSelectionItem(selectedItems[i])) {
                textCount++;
            }
        }
        return textCount;
    }

    /**
     * 属性プレビューの初期値（未取得を表す「—」）を返す
     * @returns {Object} family / style / size
     */
    function makeEmptyAttributePreview() {
        return {
            family: "—",
            style: "—",
            size: "—"
        };
    }

    /**
     * テキスト範囲から文字属性のプレビューを作る
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @returns {Object} makeEmptyAttributePreview() と同じ形のプレビュー
     */
    function getTextAttributePreviewFromTextRange(textRange) {
        var attributePreview = makeEmptyAttributePreview();

        /* フォントが見つからないなどで属性が読めないことがある / attributes may be unreadable, e.g. a missing font */
        try {
            var characterAttributes = textRange.characterAttributes;
            var textFont = characterAttributes.textFont;
            var fontSize = characterAttributes.size;
            var familyName = textFont.family || "";
            var styleName = textFont.style || "";
            var sizeText = fontSize ? String(Math.round(fontSize * 100) / 100) + " pt" : "—";

            attributePreview.family = familyName || "—";
            attributePreview.style = styleName || "—";
            attributePreview.size = sizeText;
        } catch (attributePreviewError) {
            return attributePreview;
        }

        return attributePreview;
    }

    /* プレビューの項目名 / Keys of the attribute preview */
    var ATTRIBUTE_PREVIEW_KEYS = ["family", "style", "size"];

    /**
     * 選択中のテキストから文字属性のプレビューを作る（値が食い違う項目は「混在」）
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {Object} makeEmptyAttributePreview() と同じ形のプレビュー
     */
    function getSelectedTextAttributePreview(selectedItems) {
        var mergedPreview = null;
        var mixedLabel = getLabel("preview.mixed");

        for (var i = 0; i < selectedItems.length; i++) {
            var textRange = getTextRangeFromSelectionItem(selectedItems[i]);
            if (!textRange) {
                continue;
            }
            var rangePreview = getTextAttributePreviewFromTextRange(textRange);

            if (!mergedPreview) {
                mergedPreview = rangePreview;
                continue;
            }
            for (var j = 0; j < ATTRIBUTE_PREVIEW_KEYS.length; j++) {
                var previewKey = ATTRIBUTE_PREVIEW_KEYS[j];
                if (mergedPreview[previewKey] !== rangePreview[previewKey]) {
                    mergedPreview[previewKey] = mixedLabel;
                }
            }
        }

        return mergedPreview || makeEmptyAttributePreview();
    }

    // =========================================
    // 条件による選択 / Selection by criteria
    // =========================================

    /**
     * 文字列が検索文字列に一致するかを返す
     * @param {string} frameText - テキストフレームの文字列
     * @param {string} keyword - 検索文字列
     * @param {string} matchMode - "exact" / "startsWith" / "endsWith" / "contains" / "regex"
     * @param {RegExp|null} keywordPattern - 正規表現モードのときの正規表現
     * @returns {boolean} 一致すれば true
     */
    function textMatches(frameText, keyword, matchMode, keywordPattern) {
        if (matchMode === "regex") {
            return keywordPattern ? keywordPattern.test(frameText) : false;
        }
        if (matchMode === "exact") {
            return frameText === keyword;
        }
        if (matchMode === "startsWith") {
            return frameText.indexOf(keyword) === 0;
        }
        if (matchMode === "endsWith") {
            return frameText.lastIndexOf(keyword) === frameText.length - keyword.length;
        }
        if (matchMode === "contains") {
            return frameText.indexOf(keyword) !== -1;
        }
        return false;
    }

    /**
     * 検索文字列を確かめ、正規表現モードなら RegExp にする（メッセージは呼び出し側で出す）
     * @param {string} keyword - 検索文字列
     * @param {string} matchMode - 一致のモード
     * @returns {Object} { regex: RegExp|null, errorKey: string|null }。errorKey は LABELS のパス
     */
    function parseKeywordPattern(keyword, matchMode) {
        if (!keyword) {
            return { regex: null, errorKey: "alert.emptyKeyword" };
        }
        if (matchMode !== "regex") {
            return { regex: null, errorKey: null };
        }
        /* 不正なパターンは RegExp が例外を投げる / RegExp throws on an invalid pattern */
        try {
            return { regex: new RegExp(keyword), errorKey: null };
        } catch (regexError) {
            return { regex: null, errorKey: "alert.invalidRegex" };
        }
    }

    /**
     * テキストの種類に対応する判定関数を作る
     * @param {Array|null} textKinds - 対象にする TextType の配列（null なら絞り込まない）
     * @returns {Function} テキストフレームを受け取って true/false を返す関数
     */
    function buildTextTypePredicate(textKinds) {
        if (!textKinds) {
            return function () { return true; };
        }
        return function (textFrame) {
            return indexOfItem(textKinds, textFrame.kind) !== -1;
        };
    }


    /**
     * オブジェクトが矩形に重なるかを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @param {number[]} artboardRect - [左, 上, 右, 下]
     * @returns {boolean} 重なれば true
     */
    function isItemWithinArtboardRect(pageItem, artboardRect) {
        var itemBounds = pageItem.visibleBounds; /* [left, top, right, bottom] */
        var overlapsHorizontally = itemBounds[2] >= artboardRect[0] && itemBounds[0] <= artboardRect[2];
        var overlapsVertically = itemBounds[1] >= artboardRect[3] && itemBounds[3] <= artboardRect[1];
        return overlapsHorizontally && overlapsVertically;
    }

    /**
     * 現在のアートボードに重なるオブジェクトだけに絞り込む
     * @param {Array} pageItems - 対象のオブジェクト
     * @returns {Array} 絞り込んだオブジェクト
     */
    function filterItemsByActiveArtboard(pageItems) {
        var activeArtboards = app.activeDocument.artboards;
        var artboardRect = activeArtboards[activeArtboards.getActiveArtboardIndex()].artboardRect; /* [left, top, right, bottom] */
        var itemsOnArtboard = [];
        for (var i = 0; i < pageItems.length; i++) {
            if (isItemWithinArtboardRect(pageItems[i], artboardRect)) {
                itemsOnArtboard.push(pageItems[i]);
            }
        }
        return itemsOnArtboard;
    }

    /**
     * 判定関数に一致するテキストフレームを集める
     * @param {Function} predicate - テキストフレームを受け取って true/false を返す関数
     * @param {string} artboardScope - "all" / "current"
     * @returns {TextFrame[]} 一致したテキストフレーム
     */
    function collectMatchingTextFrames(predicate, artboardScope) {
        var allTextFrames = app.activeDocument.textFrames;
        var matchedFrames = [];

        for (var i = 0; i < allTextFrames.length; i++) {
            if (predicate(allTextFrames[i])) {
                matchedFrames.push(allTextFrames[i]);
            }
        }

        if (artboardScope === "current") {
            matchedFrames = filterItemsByActiveArtboard(matchedFrames);
        }
        return matchedFrames;
    }

    /**
     * 判定関数に一致するテキストフレームを選択する
     * @param {Function} predicate - テキストフレームを受け取って true/false を返す関数
     * @param {string} artboardScope - "all" / "current"
     * @returns {TextFrame[]} 選択したテキストフレーム
     */
    function selectTextFrames(predicate, artboardScope) {
        var matchedFrames = collectMatchingTextFrames(predicate, artboardScope);
        app.activeDocument.selection = matchedFrames;
        return matchedFrames;
    }

    /**
     * 文字列条件に一致するテキストフレームの判定関数を作る
     * @param {string} keyword - 検索文字列
     * @param {string} matchMode - 一致のモード
     * @param {RegExp|null} keywordPattern - 正規表現モードのときの正規表現
     * @returns {Function} テキストフレームを受け取って true/false を返す関数
     */
    function buildTextMatchPredicate(keyword, matchMode, keywordPattern) {
        return function (textFrame) {
            return textMatches(textFrame.contents || "", keyword, matchMode, keywordPattern);
        };
    }

    /**
     * 2つの判定関数の両方に一致するかを返す判定関数を作る
     * @param {Function} firstPredicate - 1つ目の判定関数
     * @param {Function} secondPredicate - 2つ目の判定関数
     * @returns {Function} 両方が true なら true を返す関数
     */
    function combinePredicates(firstPredicate, secondPredicate) {
        return function (textFrame) {
            return firstPredicate(textFrame) && secondPredicate(textFrame);
        };
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA から読んだ SCRIPT_*（スクリプト情報の部品）を読み、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * SCRIPT_* の値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値
     */
    function getVersionInfoValue(variableKey) {
        var infoValue;
        if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
        else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
        else if (variableKey === "readme") {
            if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
            else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
        }
        else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
        else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
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

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // =========================================
    // 後処理 / Post-processing
    // =========================================

    /**
     * 編集欄の文字列を、書き戻す文字列にする
     * 編集欄は改行を \n（Windows は \r\n）で返すので、段落の改行 \r にそろえる。強制改行の代替文字は \x03 に戻す
     * @param {string} editText - 編集欄の文字列
     * @returns {string} 書き戻す文字列
     */
    function toFrameText(editText) {
        return editText.replace(/\r\n|\n/g, "\r").split(SOFT_BREAK).join("\x03");
    }

    /**
     * ほかのボタンよりひとまわり小さいボタンを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} labelKey - LABELS.button と LABELS.tooltip のキー
     * @returns {Button} 作ったボタン
     */
    function addSmallButton(parentGroup, labelKey) {
        var smallButton = parentGroup.add("button", undefined, getLabel("button." + labelKey));
        setHelpTip(smallButton, getLabel("tooltip." + labelKey, { softBreak: SOFT_BREAK }));
        var buttonFont = smallButton.graphics.font;
        smallButton.graphics.font = ScriptUI.newFont(buttonFont.name, buttonFont.style, buttonFont.size - INSERT_FONT_SHRINK);
        smallButton.preferredSize.height = INSERT_BUTTON_HEIGHT;
        return smallButton;
    }

    /**
     * 押すと入力欄のカーソルの位置に文字を入れるボタンにする
     * ボタンを押すと入力欄のカーソルが失われるので、押し下げた時点の位置を控えておく
     * @param {Button} insertButton - 対象のボタン
     * @param {EditText} targetInput - 入れる先の入力欄
     * @param {string} insertedText - 入れる文字（複数行の欄の改行は \n）
     * @returns {void}
     */
    function bindInsertButton(insertButton, targetInput, insertedText) {
        var savedCaret = null;
        /* クリックが確定する前（押し下げた時点）にカーソル位置を読む / Read the caret on mouse down, before the click completes */
        insertButton.addEventListener("mousedown", function () {
            savedCaret = captureCaret(targetInput);
        });
        insertButton.onClick = function () {
            insertTextAtCaret(targetInput, insertedText, savedCaret);
            savedCaret = null;
        };
    }

    /**
     * 控えたカーソルの位置に文字を入れる。位置が取れていなければ末尾に足す
     * @param {EditText} targetInput - 入れる先の入力欄
     * @param {string} insertedText - 入れる文字
     * @param {Object|null} caretPosition - captureCaret() の結果
     * @returns {void}
     */
    function insertTextAtCaret(targetInput, insertedText, caretPosition) {
        var currentText = targetInput.text;
        if (caretPosition && caretPosition.start + caretPosition.length <= currentText.length) {
            targetInput.text = currentText.substring(0, caretPosition.start) + insertedText + currentText.substring(caretPosition.start + caretPosition.length);
        } else {
            targetInput.text = currentText + insertedText;
        }
    }

    /**
     * 入力欄のカーソル位置を読む。目印の文字を選択範囲に差し込んで位置を測り、元の文字列に戻す
     * （ScriptUI にはカーソル位置を返すプロパティが無いため）
     * @param {EditText} targetInput - 対象の入力欄（カーソルがあるうちに呼ぶ）
     * @returns {Object|null} { start: number, length: number }。読めなければ null
     */
    function captureCaret(targetInput) {
        var CARET_MARKER = "\u0001";
        var originalText = targetInput.text;
        var selectedLength = targetInput.textselection.length;
        targetInput.textselection = CARET_MARKER;
        var markerIndex = targetInput.text.indexOf(CARET_MARKER);
        targetInput.text = originalText;
        /* カーソルを失った入力欄では目印が先頭に入るので、先頭は信用しない（末尾に足す側に倒す）
           A field that lost its caret puts the marker at the start, so treat the start as unknown */
        if (markerIndex < 0 || (markerIndex === 0 && originalText.length > 0)) return null;
        return { start: markerIndex, length: selectedLength };
    }

    /**
     * 一括編集の文字列を入力するダイアログを表示する（改行・強制改行はボタンか Shift+Enter で入れる）
     * @returns {string|null} 書き戻す文字列（キャンセル時は null）
     */
    function showBulkEditDialog() {
        var bulkDialog = new Window("dialog", getLabel("dialog.bulkEditTitle"));
        setupWindow(bulkDialog);

        /* 入力欄の上に改行の挿入ボタンを右寄せで並べる / Break insert buttons, right-aligned above the field */
        var insertButtonRow = bulkDialog.add("group");
        insertButtonRow.orientation = "row";
        insertButtonRow.alignment = ["right", "top"];
        insertButtonRow.alignChildren = ["right", "center"];
        insertButtonRow.spacing = 4;
        var btnParagraphBreak = addSmallButton(insertButtonRow, "paragraphBreak");
        var btnLineBreak = addSmallButton(insertButtonRow, "lineBreak");

        var replacementInput = bulkDialog.add("edittext", undefined, "", {
            multiline: true,
            scrolling: true
        });
        replacementInput.preferredSize = BULK_INPUT_SIZE;
        setHelpTip(replacementInput, getLabel("tooltip.bulkInput", { softBreak: SOFT_BREAK }));
        replacementInput.active = true;

        /* Shift+Enter で強制改行の代替文字を入れる / Insert the soft-break placeholder with Shift+Enter */
        replacementInput.addEventListener("keydown", function (keyEvent) {
            if (ScriptUI.environment.keyboardState.shiftKey && keyEvent.keyName === "Enter") {
                this.textselection = SOFT_BREAK;
                keyEvent.preventDefault();
            }
        });
        bindInsertButton(btnParagraphBreak, replacementInput, "\n");
        bindInsertButton(btnLineBreak, replacementInput, SOFT_BREAK);

        var bulkButtonRow = addButtonRow(bulkDialog);
        var btnBulkCancel = bulkButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnBulkOK = bulkButtonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnBulkOK.onClick = function () {
            bulkDialog.close(1);
        };
        btnBulkCancel.onClick = function () {
            bulkDialog.close(0);
        };

        alignRightOnlyButtonRow(bulkButtonRow);
        prepareDialogWindow(bulkDialog, SCRIPT_NAME + "_bulkEdit");
        if (bulkDialog.show() !== 1) {
            return null;
        }
        return toFrameText(replacementInput.text);
    }

    /**
     * テキストフレームの内容を、文字ごとの書式を残したまま置き換える。
     * contents をまとめて代入すると全体が先頭の文字の書式になるので、文字（characters）単位で書き換える。
     * 番号が変わらないよう右から処理する。増えた文字は元の最後の文字の書式を引き継ぐ
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} replacementText - 置き換える文字列（段落の改行は \r、強制改行は \x03）
     * @returns {void}
     */
    function replaceTextKeepingFormat(textFrame, replacementText) {
        var frameCharacters = textFrame.textRange.characters;
        var originalLength = frameCharacters.length;
        var replacementLength = replacementText.length;
        if (originalLength === 0 || replacementLength === 0) {
            textFrame.contents = replacementText;
            return;
        }

        /* 余る文字を末尾から消す / Remove the surplus characters from the end */
        for (var i = originalLength - 1; i >= replacementLength; i--) {
            frameCharacters[i].remove();
        }
        /* 足りない分は、共通部分の最後の文字に続けて入れる（その文字の書式になる）/ Append the extra text to the last shared character */
        var sharedLength = Math.min(originalLength, replacementLength);
        var lastSharedIndex = sharedLength - 1;
        frameCharacters[lastSharedIndex].contents = replacementText.substring(lastSharedIndex);
        /* 残りは1文字ずつ、変わった文字だけ書き換える / Rewrite the rest one character at a time, only where it changed */
        for (var j = lastSharedIndex - 1; j >= 0; j--) {
            var replacementCharacter = replacementText.charAt(j);
            if (frameCharacters[j].contents !== replacementCharacter) {
                frameCharacters[j].contents = replacementCharacter;
            }
        }
    }

    /**
     * テキストフレームの内容をまとめて置き換える（文字ごとの書式は元の文字の位置ごとに引き継ぐ）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function bulkEditTextFrames(textFrames) {
        if (textFrames.length === 0) {
            return;
        }

        var replacementText = showBulkEditDialog();
        if (replacementText === null) {
            return;
        }

        for (var i = 0; i < textFrames.length; i++) {
            /* 空のテキストは書式の手がかりが無いので変えない / Leave empty text alone; it has no formatting to carry */
            if (textFrames[i].textRange.length === 0) {
                continue;
            }
            replaceTextKeepingFormat(textFrames[i], replacementText);
        }

        app.redraw();
    }

    /**
     * レイヤーを名前で取得し、無ければ作る。ロック解除・表示にして、元の状態も返す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Object} { layer, originalLocked, originalVisible }
     */
    function getOrCreateLayerByName(doc, layerName) {
        var targetLayer;
        /* getByName は見つからないと例外 / getByName throws when the layer is missing */
        try {
            targetLayer = doc.layers.getByName(layerName);
        } catch (layerFindError) {
            targetLayer = doc.layers.add();
            targetLayer.name = layerName;
        }

        var originalLocked = targetLayer.locked;
        var originalVisible = targetLayer.visible;

        targetLayer.locked = false;
        targetLayer.visible = true;

        return {
            layer: targetLayer,
            originalLocked: originalLocked,
            originalVisible: originalVisible
        };
    }

    /**
     * 選択中の子孫を持つかを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 子孫に選択中のものがあれば true
     */
    function containsSelectedDescendant(pageItem) {
        var childItems;
        /* pageItems を持たないオブジェクトでは例外になる / items without pageItems throw */
        try {
            childItems = pageItem.pageItems;
        } catch (childAccessError) {
            return false;
        }
        if (!childItems || childItems.length === 0) {
            return false;
        }
        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];
            if (childItem.selected || containsSelectedDescendant(childItem)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 選択中以外のオブジェクトを非表示にする（ロック中と、選択を含むグループは残す）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function hideItemsExceptSelection(doc) {
        var allPageItems = doc.pageItems;
        for (var i = 0; i < allPageItems.length; i++) {
            var pageItem = allPageItems[i];
            /* ロックされたレイヤー上などで非表示化できない場合はスキップ / Skip items that cannot be hidden because of parent layer state */
            try {
                /* 選択中・ロック中・選択を内包するグループは残す / Keep selected, locked, or selection-containing items */
                if (pageItem.selected || pageItem.locked) {
                    continue;
                }
                if (containsSelectedDescendant(pageItem)) {
                    continue;
                }
                pageItem.hidden = true;
            } catch (hideOtherError) {
            }
        }
    }

    /**
     * 選択中のオブジェクトを非表示にする
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function hideSelectedItems(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            /* ロックされたレイヤー上などで非表示化できない場合はスキップ / Skip items that cannot be hidden because of parent layer state */
            try {
                selectedItems[i].hidden = true;
            } catch (hideItemError) {
            }
        }
    }

    /**
     * 選択中のオブジェクトを指定したレイヤーへ移動する（レイヤーのロック・表示状態は元に戻す）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {string} layerName - 移動先のレイヤー名
     * @returns {void}
     */
    function moveItemsToLayer(doc, selectedItems, layerName) {
        var layerInfo = getOrCreateLayerByName(doc, layerName);
        var destinationLayer = layerInfo.layer;
        for (var i = 0; i < selectedItems.length; i++) {
            /* ロック状態や親レイヤーの状態によって移動できない場合はスキップ / Skip items that cannot be moved because of lock state or parent layer state */
            try {
                selectedItems[i].locked = false;
                selectedItems[i].move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (moveItemError) {
            }
        }
        /* 元のロック・可視状態へ戻す（新規作成時は既定値に戻るだけ）/ Restore the original lock and visibility (a no-op for a new layer) */
        destinationLayer.locked = layerInfo.originalLocked;
        destinationLayer.visible = layerInfo.originalVisible;
    }

    /**
     * 選択後の処理を行う
     * @param {TextFrame[]} selectedFrames - 選択したテキストフレーム
     * @param {string} postProcessMode - "" / "hide" / "hideOthers" / "moveToLayer" / "bulkEdit"
     * @param {string} moveLayerName - 「レイヤーへ移動」の移動先
     * @returns {void}
     */
    function applyPostProcessToSelection(selectedFrames, postProcessMode, moveLayerName) {
        var doc = app.activeDocument;
        if (postProcessMode === "hide") {
            hideSelectedItems(selectedFrames);
        } else if (postProcessMode === "hideOthers") {
            hideItemsExceptSelection(doc);
        } else if (postProcessMode === "bulkEdit") {
            bulkEditTextFrames(selectedFrames);
        } else if (postProcessMode === "moveToLayer") {
            moveItemsToLayer(doc, selectedFrames, moveLayerName);
        }
    }

    /**
     * 選択件数を確かめてから後処理を行う（0件ならメッセージ）
     * @param {TextFrame[]} selectedFrames - 選択したテキストフレーム
     * @param {string} postProcessMode - 後処理のモード
     * @param {string} moveLayerName - 「レイヤーへ移動」の移動先
     * @returns {void}
     */
    function finalizeSelection(selectedFrames, postProcessMode, moveLayerName) {
        if (selectedFrames.length === 0) {
            alert(getLabel("alert.noMatch"));
            return;
        }
        applyPostProcessToSelection(selectedFrames, postProcessMode, moveLayerName);
    }

    // =========================================
    // フォント属性の照合 / Font attribute matching
    // =========================================

    /**
     * 色を比べるための文字列にする
     * @param {Color} color - 塗りの色
     * @returns {string} 色の種類と値をつないだ文字列
     */
    function getColorKey(color) {
        if (!color) return "";
        switch (color.typename) {
            case "RGBColor": return "RGB:" + [color.red, color.green, color.blue].join(",");
            case "CMYKColor": return "CMYK:" + [color.cyan, color.magenta, color.yellow, color.black].join(",");
            case "GrayColor": return "Gray:" + color.gray;
            case "SpotColor": return "Spot:" + color.spot.name + ":" + color.tint;
            default: return color.typename;
        }
    }

    /**
     * テキストの属性を、比較に使う文字列にする
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @param {string} attributeKey - "family" / "style" / "size" / "fillColor"
     * @returns {string|null} 比較用の文字列（読めなければ null）
     */
    function getTextAttributeKey(textRange, attributeKey) {
        /* フォントが見つからないなどで属性が読めないことがある / attributes may be unreadable, e.g. a missing font */
        try {
            var characterAttributes = textRange.characterAttributes;
            if (attributeKey === "fillColor") return getColorKey(characterAttributes.fillColor);
            if (attributeKey === "size") return String(characterAttributes.size);
            if (attributeKey === "family") return characterAttributes.textFont.family;
            return characterAttributes.textFont.style;
        } catch (attributeKeyError) {
            return null;
        }
    }

    /**
     * オンにした属性をまとめて、比較用の文字列にする
     * @param {Object} textItem - テキストフレームまたは選択中のテキスト
     * @param {string[]} attributeKeys - 比べる属性
     * @returns {string|null} 比較用の文字列（テキストでない・読めない属性があれば null）
     */
    function getCombinedAttributeKey(textItem, attributeKeys) {
        var textRange = getTextRangeFromSelectionItem(textItem);
        if (!textRange) return null;
        var keyParts = [];
        for (var i = 0; i < attributeKeys.length; i++) {
            var attributeValue = getTextAttributeKey(textRange, attributeKeys[i]);
            if (attributeValue === null) return null;
            keyParts.push(attributeValue);
        }
        return keyParts.join("\n");
    }

    /**
     * 選択中のテキストのどれかと、オンにした属性がすべて一致するかの判定関数を作る
     * @param {Array} selectedItems - 開いた時点の選択
     * @param {string[]} attributeKeys - 比べる属性
     * @returns {Function} テキストフレームを受け取って true/false を返す関数
     */
    function buildAttributePredicate(selectedItems, attributeKeys) {
        var wantedKeys = {};
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedKey = getCombinedAttributeKey(selectedItems[i], attributeKeys);
            if (selectedKey !== null) wantedKeys[selectedKey] = true;
        }
        return function (textFrame) {
            var candidateKey = getCombinedAttributeKey(textFrame, attributeKeys);
            return candidateKey !== null && wantedKeys.hasOwnProperty(candidateKey);
        };
    }

    /**
     * 設定から、テキストの種類・文字列・フォント属性をすべて満たす判定関数を作る
     * @param {Object} selectorSettings - readDialogSettings() の結果
     * @param {Array} initialSelection - 開いた時点の選択（フォント属性の基準）
     * @param {RegExp|null} keywordPattern - 正規表現モードのときの正規表現
     * @returns {Function} テキストフレームを受け取って true/false を返す関数
     */
    function buildSettingsPredicate(selectorSettings, initialSelection, keywordPattern) {
        /* 軽い判定から順に重ねる（属性の読み取りがいちばん重い）/ Cheapest checks first; reading attributes is the slowest */
        var settingsPredicate = buildTextTypePredicate(selectorSettings.textKinds);
        if (selectorSettings.textMatchMode) {
            settingsPredicate = combinePredicates(settingsPredicate, buildTextMatchPredicate(selectorSettings.keyword, selectorSettings.textMatchMode, keywordPattern));
        }
        if (selectorSettings.attributeKeys.length > 0) {
            settingsPredicate = combinePredicates(settingsPredicate, buildAttributePredicate(initialSelection, selectorSettings.attributeKeys));
        }
        return settingsPredicate;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ラジオ・チェックボックスの対応表。labelKey は LABELS.radio / checkbox / tooltip のキー
       Tables for the radios and checkboxes; labelKey indexes LABELS.radio / checkbox / tooltip */

    /* テキストの種類のチェックボックス / Text type checkboxes */
    var TEXT_TYPE_CHECKBOXES = [
        { controlKey: "cbPointText", labelKey: "pointText", textKind: TextType.POINTTEXT, shortcutKey: "Alt+W" },
        { controlKey: "cbAreaText", labelKey: "areaText", textKind: TextType.AREATEXT, shortcutKey: "Alt+E" },
        { controlKey: "cbPathText", labelKey: "pathText", textKind: TextType.PATHTEXT, shortcutKey: "Alt+T" }
    ];

    /* フォント属性のチェックボックス（previewKey は見本の項目、null なら見本なし）/ Font attribute checkboxes */
    var ATTRIBUTE_CHECKBOXES = [
        { controlKey: "cbFontFamily", labelKey: "fontFamily", attributeKey: "family", previewKey: "family" },
        { controlKey: "cbFontStyle", labelKey: "fontStyle", attributeKey: "style", previewKey: "style" },
        { controlKey: "cbFontSize", labelKey: "fontSize", attributeKey: "size", previewKey: "size" },
        { controlKey: "cbTextFillColor", labelKey: "textFillColor", attributeKey: "fillColor", previewKey: null }
    ];

    /* 文字列の条件のラジオ（column は 2列のどちらに置くか）/ String condition radios; column picks one of two columns */
    var TEXT_MATCH_RADIOS = [
        { controlKey: "rbNoTextMatch", labelKey: "noTextMatch", mode: "", column: 0 },
        { controlKey: "rbExactMatch", labelKey: "exactMatch", mode: "exact", column: 0, shortcutKey: "Alt+A" },
        { controlKey: "rbContainsMatch", labelKey: "containsMatch", mode: "contains", column: 0, shortcutKey: "Alt+I" },
        { controlKey: "rbStartsWith", labelKey: "startsWith", mode: "startsWith", column: 1, shortcutKey: "Alt+B" },
        { controlKey: "rbEndsWith", labelKey: "endsWith", mode: "endsWith", column: 1, shortcutKey: "Alt+D" },
        { controlKey: "rbRegexMatch", labelKey: "regexMatch", mode: "regex", column: 1, shortcutKey: "Alt+R" }
    ];

    /* 選択後の処理のラジオ / Post-process radios */
    var POST_PROCESS_RADIOS = [
        { controlKey: "rbNoPostProcess", labelKey: "noPostProcess", mode: "", column: 0 },
        { controlKey: "rbMoveToLayer", labelKey: "moveToLayer", mode: "moveToLayer", column: 0 },
        { controlKey: "rbHideSelected", labelKey: "hideAfterSelection", mode: "hide", column: 1 },
        { controlKey: "rbHideOthers", labelKey: "hideOthers", mode: "hideOthers", column: 1 },
        { controlKey: "rbBulkEdit", labelKey: "bulkEdit", mode: "bulkEdit", column: 1 }
    ];

    /**
     * ヘルプチップを設定する（文言が空なら何もしない）
     * @param {Object} targetControl - 対象のコントロール
     * @param {string} tipText - ヘルプチップの文言
     * @returns {void}
     */
    function setHelpTip(targetControl, tipText) {
        if (targetControl && tipText) {
            targetControl.helpTip = tipText;
        }
    }

    /**
     * 縦並び・左揃えの列グループを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加した列グループ
     */
    function addColumnGroup(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "left";
        columnGroup.spacing = 6;
        return columnGroup;
    }

    /**
     * 配列の中の位置を返す（ES3 に Array.indexOf は無い）
     * @param {Array} items - 探す配列
     * @param {*} targetItem - 探す値
     * @returns {number} 位置（無ければ -1）
     */
    function indexOfItem(items, targetItem) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === targetItem) return i;
        }
        return -1;
    }

    /**
     * 対応表のラジオを列グループに作り、列をまたいで排他にする（ラジオは同じ親の中でしか排他にならない）
     * @param {Object} selectorControls - コントロールの格納先
     * @param {Object[]} radioTable - controlKey / labelKey / column を持つ行の配列
     * @param {Group[]} columnGroups - 列グループ（column の番号で選ぶ）
     * @param {Function} onSelect - 選んだあとに呼ぶ関数
     * @returns {RadioButton[]} 作ったラジオ
     */
    function addExclusiveRadios(selectorControls, radioTable, columnGroups, onSelect) {
        var radioButtons = [];
        for (var i = 0; i < radioTable.length; i++) {
            var radioRow = radioTable[i];
            var radioButton = columnGroups[radioRow.column].add("radiobutton", undefined, getLabel("radio." + radioRow.labelKey));
            setHelpTip(radioButton, getLabel("tooltip." + radioRow.labelKey));
            selectorControls[radioRow.controlKey] = radioButton;
            radioButtons.push(radioButton);
        }
        for (var j = 0; j < radioButtons.length; j++) {
            radioButtons[j].onClick = function () {
                for (var k = 0; k < radioButtons.length; k++) {
                    if (radioButtons[k] !== this) radioButtons[k].value = false;
                }
                onSelect();
            };
        }
        radioButtons[0].value = true;
        return radioButtons;
    }

    /**
     * 対応表から、オンになっている最初の行の値を返す
     * @param {Object} selectorControls - buildDialog() の結果
     * @param {Object[]} optionTable - controlKey と値を持つ行の配列
     * @param {string} valueKey - 返す値のキー
     * @returns {*} 値（どれもオフなら最初の行の値）
     */
    function readRadioChoice(selectorControls, optionTable, valueKey) {
        for (var i = 0; i < optionTable.length; i++) {
            if (selectorControls[optionTable[i].controlKey].value) return optionTable[i][valueKey];
        }
        return optionTable[0][valueKey];
    }

    /**
     * 対応表のうち、オンになっている行の値を集める
     * @param {Object} selectorControls - buildDialog() の結果
     * @param {Object[]} optionTable - controlKey と値を持つ行の配列
     * @param {string} valueKey - 集める値のキー
     * @returns {Array} オンの行の値
     */
    function readCheckedValues(selectorControls, optionTable, valueKey) {
        var checkedValues = [];
        for (var i = 0; i < optionTable.length; i++) {
            if (selectorControls[optionTable[i].controlKey].value) checkedValues.push(optionTable[i][valueKey]);
        }
        return checkedValues;
    }

    /**
     * ［アートボード］パネルを追加する
     * @param {Panel} selectionPanel - 追加先
     * @param {Object} selectorControls - コントロールの格納先
     * @returns {void}
     */
    function addArtboardScopePanel(selectionPanel, selectorControls) {
        var artboardScopePanel = selectionPanel.add("panel", undefined, getLabel("panel.artboardScope"));
        setupPanel(artboardScopePanel, COLUMN_SPACING);
        artboardScopePanel.orientation = "row";
        artboardScopePanel.alignChildren = ["left", "center"];

        selectorControls.rbArtboardAll = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.artboardAll"));
        selectorControls.rbArtboardCurrent = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.artboardCurrent"));
        selectorControls.rbArtboardAll.value = true;
        setHelpTip(selectorControls.rbArtboardAll, getLabel("tooltip.artboardAll"));
        setHelpTip(selectorControls.rbArtboardCurrent, getLabel("tooltip.artboardCurrent"));
        /* 同じ親の中なので排他は自動。切り替えで数え直す / Same parent, so exclusive already; recount on change */
        selectorControls.rbArtboardAll.onClick = selectorControls.rbArtboardCurrent.onClick = function () {
            updateTargetState(selectorControls);
        };
    }

    /**
     * ［テキストの種類］パネルを追加する（初期値はすべてオン）
     * @param {Panel} selectionPanel - 追加先
     * @param {Object} selectorControls - コントロールの格納先
     * @returns {void}
     */
    function addTextTypePanel(selectionPanel, selectorControls) {
        var textTypePanel = selectionPanel.add("panel", undefined, getLabel("panel.textType"));
        setupPanel(textTypePanel, COLUMN_SPACING);
        textTypePanel.orientation = "row";
        textTypePanel.alignChildren = ["left", "center"];
        setHelpTip(textTypePanel, getLabel("tooltip.textTypePanel"));

        for (var i = 0; i < TEXT_TYPE_CHECKBOXES.length; i++) {
            var textTypeRow = TEXT_TYPE_CHECKBOXES[i];
            var textTypeCheckbox = textTypePanel.add("checkbox", undefined, getLabel("checkbox." + textTypeRow.labelKey));
            setHelpTip(textTypeCheckbox, getLabel("tooltip." + textTypeRow.labelKey));
            textTypeCheckbox.value = true;
            textTypeCheckbox.onClick = function () {
                updateTargetState(selectorControls);
            };
            selectorControls[textTypeRow.controlKey] = textTypeCheckbox;
        }
    }

    /**
     * チェックボックスと、右に添える値の見本を1行に並べる
     * @param {Panel} parentPanel - 追加先
     * @param {string} checkboxLabel - チェックボックスのラベル
     * @param {string|null} previewText - 見本の文字列（null なら添えない）
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckboxRow(parentPanel, checkboxLabel, previewText) {
        var checkboxRowGroup = parentPanel.add("group");
        checkboxRowGroup.orientation = "row";
        checkboxRowGroup.alignChildren = ["left", "center"];
        checkboxRowGroup.alignment = "fill";

        var rowCheckbox = checkboxRowGroup.add("checkbox", undefined, checkboxLabel);
        rowCheckbox.preferredSize.width = ATTRIBUTE_LABEL_WIDTH;
        if (previewText !== null) {
            var previewLabel = checkboxRowGroup.add("statictext", undefined, previewText);
            previewLabel.alignment = ["fill", "center"];
            previewLabel.characters = PREVIEW_CHARACTERS;
        }
        return rowCheckbox;
    }

    /**
     * ［フォント属性］パネルを追加する（選択があればファミリー・スタイル・サイズをオン）
     * @param {Panel} selectionPanel - 追加先
     * @param {Object} selectorControls - コントロールの格納先
     * @param {Object} initialState - 開いた時点の選択の情報
     * @returns {void}
     */
    function addAttributePanel(selectionPanel, selectorControls, initialState) {
        var attributePanel = selectionPanel.add("panel", undefined, getLabel("panel.attribute"));
        setupPanel(attributePanel, 6);
        attributePanel.enabled = initialState.hasSelection;
        setHelpTip(attributePanel, getLabel("tooltip.attributePanel"));
        selectorControls.attributePanel = attributePanel;

        var attributeCheckboxes = [];
        for (var i = 0; i < ATTRIBUTE_CHECKBOXES.length; i++) {
            var attributeRow = ATTRIBUTE_CHECKBOXES[i];
            var previewText = attributeRow.previewKey ? initialState.attributePreview[attributeRow.previewKey] : null;
            var attributeCheckbox = addCheckboxRow(attributePanel, getLabel("checkbox." + attributeRow.labelKey), previewText);
            setHelpTip(attributeCheckbox, getLabel("tooltip.attributeMatch", { attribute: getLabel("attributeName." + attributeRow.labelKey) }));
            attributeCheckbox.value = initialState.hasSelection && attributeRow.attributeKey !== "fillColor";
            attributeCheckbox.onClick = function () {
                if (ScriptUI.environment.keyboardState.altKey) {
                    applyAttributeOptionClick(attributeCheckboxes, this);
                }
                updateTargetState(selectorControls);
            };
            selectorControls[attributeRow.controlKey] = attributeCheckbox;
            attributeCheckboxes.push(attributeCheckbox);
        }
    }

    /**
     * 属性のチェックボックスの Option＋クリック：すべてオンでなければすべてオン、すべてオンならクリックしたもの以外をオン
     * @param {Checkbox[]} attributeCheckboxes - 属性のチェックボックス
     * @param {Checkbox} clickedCheckbox - クリックしたチェックボックス（クリックで値は反転済み）
     * @returns {void}
     */
    function applyAttributeOptionClick(attributeCheckboxes, clickedCheckbox) {
        /* クリックで反転する前の状態で、すべてオンだったかを判定する / Judge using the state before this click toggled it */
        var wereAllChecked = true;
        for (var i = 0; i < attributeCheckboxes.length; i++) {
            var previousValue = (attributeCheckboxes[i] === clickedCheckbox) ? !attributeCheckboxes[i].value : attributeCheckboxes[i].value;
            if (!previousValue) wereAllChecked = false;
        }
        for (var j = 0; j < attributeCheckboxes.length; j++) {
            attributeCheckboxes[j].value = wereAllChecked ? (attributeCheckboxes[j] !== clickedCheckbox) : true;
        }
    }

    /**
     * ［文字列］パネルを追加する（左に検索文字列、右に条件のラジオ2列）
     * @param {Panel} selectionPanel - 追加先
     * @param {Object} selectorControls - コントロールの格納先
     * @param {Object} initialState - 開いた時点の選択の情報
     * @returns {void}
     */
    function addTextMatchPanel(selectionPanel, selectorControls, initialState) {
        var textMatchPanel = selectionPanel.add("panel", undefined, getLabel("panel.textMatch"));
        setupPanel(textMatchPanel, COLUMN_SPACING);
        textMatchPanel.orientation = "row";
        textMatchPanel.alignChildren = ["left", "fill"];
        /* テキストを複数選択しているときはディム / Disabled while several texts are selected */
        textMatchPanel.enabled = !initialState.hasMultipleTexts;
        selectorControls.textMatchPanel = textMatchPanel;

        /* 高さを取るので複数行にし、文字を上から表示する / Multiline so text starts at the top of the taller field */
        var keywordInput = textMatchPanel.add("edittext", undefined, initialState.keyword, { multiline: true, scrolling: false });
        keywordInput.characters = KEYWORD_CHARACTERS;
        keywordInput.preferredSize.height = KEYWORD_FIELD_HEIGHT;
        /* 幅は文字数で決め、高さはラジオの列に合わせて伸ばす / Width from characters; height follows the radio column */
        keywordInput.alignment = ["left", "fill"];
        setHelpTip(keywordInput, getLabel("tooltip.keywordInput"));
        keywordInput.onChanging = function () {
            updateTargetState(selectorControls);
        };
        selectorControls.keywordInput = keywordInput;

        var textMatchColumnsGroup = textMatchPanel.add("group");
        textMatchColumnsGroup.orientation = "row";
        textMatchColumnsGroup.alignment = ["left", "top"];
        textMatchColumnsGroup.alignChildren = ["left", "top"];
        textMatchColumnsGroup.spacing = COLUMN_SPACING;
        var textMatchColumns = [addColumnGroup(textMatchColumnsGroup), addColumnGroup(textMatchColumnsGroup)];
        addExclusiveRadios(selectorControls, TEXT_MATCH_RADIOS, textMatchColumns, function () {
            updateTargetState(selectorControls);
        });
    }

    /**
     * ［選択条件］パネル（アートボード・テキストの種類・フォント属性・文字列）を追加する
     * @param {Object} selectorControls - コントロールの格納先
     * @param {Object} initialState - 開いた時点の選択の情報（selectedItems / hasSelection / hasMultipleTexts / keyword / attributePreview）
     * @returns {void}
     */
    function addSelectionPanel(selectorControls, initialState) {
        var selectionPanel = selectorControls.selectorDialog.add("panel", undefined, getLabel("panel.selection"));
        setupPanel(selectionPanel);
        addArtboardScopePanel(selectionPanel, selectorControls);
        addTextTypePanel(selectionPanel, selectorControls);
        addAttributePanel(selectionPanel, selectorControls, initialState);
        addTextMatchPanel(selectionPanel, selectorControls, initialState);
    }

    /**
     * ［選択後の処理］パネルを追加する（左右2列。移動先のレイヤー名は［レイヤーへ移動］の下）
     * @param {Object} selectorControls - コントロールの格納先
     * @returns {void}
     */
    function addPostProcessPanel(selectorControls) {
        var postProcessPanel = selectorControls.selectorDialog.add("panel", undefined, getLabel("panel.postProcess"));
        setupPanel(postProcessPanel, COLUMN_SPACING);
        postProcessPanel.orientation = "row";
        postProcessPanel.alignChildren = ["left", "top"];
        /* ドキュメント内にテキストが無ければ後処理は無意味なのでディム / Disable post-process when the document has no text */
        postProcessPanel.enabled = app.activeDocument.textFrames.length > 0;

        var postProcessColumns = [addColumnGroup(postProcessPanel), addColumnGroup(postProcessPanel)];
        addExclusiveRadios(selectorControls, POST_PROCESS_RADIOS, postProcessColumns, function () {
            selectorControls.moveLayerInput.enabled = selectorControls.rbMoveToLayer.value;
        });

        /* 移動先のレイヤー名（左の列の末尾＝［レイヤーへ移動］の下に字下げ）/ Destination layer, indented under Move to Layer */
        var moveLayerGroup = postProcessColumns[0].add("group");
        moveLayerGroup.margins = [MOVE_LAYER_FIELD_INDENT, 0, 0, 0];
        selectorControls.moveLayerInput = moveLayerGroup.add("edittext", undefined, DEFAULT_MOVE_LAYER_NAME);
        selectorControls.moveLayerInput.characters = MOVE_LAYER_FIELD_CHARACTERS;
        selectorControls.moveLayerInput.enabled = false;
        setHelpTip(selectorControls.moveLayerInput, getLabel("tooltip.moveLayerName"));
    }

    /**
     * Option＋キーのショートカットを付ける（検索文字列の入力中は効かない。ツールチップにキーを足す）
     * @param {Object} selectorControls - buildDialog() の結果
     * @returns {void}
     */
    function addSelectionKeyShortcuts(selectorControls) {
        var shortcutMap = {};
        var shortcutTables = [TEXT_TYPE_CHECKBOXES, TEXT_MATCH_RADIOS];
        for (var i = 0; i < shortcutTables.length; i++) {
            for (var j = 0; j < shortcutTables[i].length; j++) {
                var shortcutRow = shortcutTables[i][j];
                if (shortcutRow.shortcutKey) shortcutMap[shortcutRow.shortcutKey] = selectorControls[shortcutRow.controlKey];
            }
        }
        addKeyShortcuts(selectorControls.selectorDialog, shortcutMap, { showInTip: true });
    }

    /**
     * ダイアログの状態から設定を読み取る
     * @param {Object} selectorControls - buildDialog() の結果
     * @returns {Object} artboardScope / textKinds / attributeKeys / textMatchMode / keyword / postProcessMode / moveLayerName
     */
    function readDialogSettings(selectorControls) {
        var checkedKinds = readCheckedValues(selectorControls, TEXT_TYPE_CHECKBOXES, "textKind");
        return {
            artboardScope: selectorControls.rbArtboardCurrent.value ? "current" : "all",
            /* すべてオンなら null（絞り込まない）/ null when all are on (no filter) */
            textKinds: (checkedKinds.length === TEXT_TYPE_CHECKBOXES.length) ? null : checkedKinds,
            attributeKeys: selectorControls.attributePanel.enabled ? readCheckedValues(selectorControls, ATTRIBUTE_CHECKBOXES, "attributeKey") : [],
            textMatchMode: selectorControls.textMatchPanel.enabled ? readRadioChoice(selectorControls, TEXT_MATCH_RADIOS, "mode") : "",
            /* 複数行の欄の改行は \n なので、テキストの contents に合わせて \r にする / Field newlines are \n; contents use \r */
            keyword: (selectorControls.keywordInput.text || "").replace(/\r\n|\n/g, "\r"),
            postProcessMode: readRadioChoice(selectorControls, POST_PROCESS_RADIOS, "mode"),
            moveLayerName: selectorControls.moveLayerInput.text || DEFAULT_MOVE_LAYER_NAME
        };
    }

    /**
     * 今の設定で選ばれるテキストの数を返す
     * @param {Object} selectorControls - buildDialog() の結果
     * @param {Object} selectorSettings - readDialogSettings() の結果
     * @returns {number} 一致する数
     */
    function countTargetTexts(selectorControls, selectorSettings) {
        if (selectorSettings.textMatchMode) {
            /* 入力途中の空欄・不正な正規表現は 0 件として、メッセージは出さない / Treat a field still being typed as 0, quietly */
            var parsedKeyword = parseKeywordPattern(selectorSettings.keyword, selectorSettings.textMatchMode);
            if (parsedKeyword.errorKey) return 0;
            return collectMatchingTextFrames(buildSettingsPredicate(selectorSettings, selectorControls.initialSelection, parsedKeyword.regex), selectorSettings.artboardScope).length;
        }
        /* 文字列で絞らないときの数は、組み合わせごとに控える（属性の読み取りが重い）/ Cache per combination when not filtering by string */
        var cacheKey = selectorSettings.attributeKeys.join(",") + "@" + (selectorSettings.textKinds ? selectorSettings.textKinds.join(",") : "all") + "@" + selectorSettings.artboardScope;
        if (!selectorControls.targetCountCache.hasOwnProperty(cacheKey)) {
            selectorControls.targetCountCache[cacheKey] = collectMatchingTextFrames(buildSettingsPredicate(selectorSettings, selectorControls.initialSelection, null), selectorSettings.artboardScope).length;
        }
        return selectorControls.targetCountCache[cacheKey];
    }

    /**
     * 最上部の対象テキスト数を更新し、［一括編集］を［文字列］の条件を選んでいるときだけ使えるようにする
     * @param {Object} selectorControls - buildDialog() の結果
     * @returns {void}
     */
    function updateTargetState(selectorControls) {
        var selectorSettings = readDialogSettings(selectorControls);
        selectorControls.targetCountText.text = labelValueText("fieldLabel.targetCount", countTargetTexts(selectorControls, selectorSettings));

        var isBulkEditAvailable = selectorSettings.textMatchMode !== "";
        selectorControls.rbBulkEdit.enabled = isBulkEditAvailable;
        if (!isBulkEditAvailable && selectorControls.rbBulkEdit.value) {
            selectorControls.rbBulkEdit.value = false;
            selectorControls.rbNoPostProcess.value = true;
        }
    }

    /**
     * ダイアログを組み立てる（OK/キャンセルの処理は main() で付ける）
     * @param {Object} initialState - 開いた時点の選択の情報（selectedItems / hasSelection / hasMultipleTexts / keyword / attributePreview）
     * @returns {Object} ダイアログ本体（selectorDialog）と各コントロール
     */
    function buildDialog(initialState) {
        var selectorControls = {
            initialSelection: initialState.selectedItems,
            targetCountCache: {}
        };
        selectorControls.selectorDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(selectorControls.selectorDialog);

        /* 対象テキスト数（左右中央）/ Number of target texts, centered */
        var targetCountGroup = selectorControls.selectorDialog.add("group");
        targetCountGroup.alignment = "center";
        selectorControls.targetCountText = targetCountGroup.add("statictext", undefined, "");
        selectorControls.targetCountText.characters = TARGET_COUNT_CHARACTERS;
        selectorControls.targetCountText.justify = "center";
        setHelpTip(selectorControls.targetCountText, getLabel("tooltip.targetCount"));

        addSelectionPanel(selectorControls, initialState);
        addPostProcessPanel(selectorControls);
        addSelectionKeyShortcuts(selectorControls);

        /* ボタンエリア / Button area */
        var buttonRow = addButtonRow(selectorControls.selectorDialog);
        selectorControls.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        selectorControls.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        updateTargetState(selectorControls);
        return selectorControls;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで条件を選び、テキストフレームを選択して後処理を行う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var initialSelection = collectSelectionAsArray(app.activeDocument.selection);
        /* 複数のテキストの文字列は初期値にしない（欄が広がる）/ Don't prefill from several texts (the field would grow) */
        var hasMultipleTexts = countSelectedTexts(initialSelection) > 1;
        var selectorControls = buildDialog({
            selectedItems: initialSelection,
            hasSelection: initialSelection.length > 0,
            hasMultipleTexts: hasMultipleTexts,
            keyword: hasMultipleTexts ? "" : getSelectedTextString(initialSelection),
            attributePreview: getSelectedTextAttributePreview(initialSelection)
        });

        selectorControls.btnOK.onClick = function () {
            var selectorSettings = readDialogSettings(selectorControls);

            /* 文字列で絞るときは入力を確かめる / Validate the keyword when filtering by string */
            var keywordPattern = null;
            if (selectorSettings.textMatchMode) {
                var parsedKeyword = parseKeywordPattern(selectorSettings.keyword, selectorSettings.textMatchMode);
                if (parsedKeyword.errorKey) {
                    alert(getLabel(parsedKeyword.errorKey));
                    return;
                }
                keywordPattern = parsedKeyword.regex;
            }

            selectorControls.selectorDialog.close();
            var settingsPredicate = buildSettingsPredicate(selectorSettings, selectorControls.initialSelection, keywordPattern);
            finalizeSelection(selectTextFrames(settingsPredicate, selectorSettings.artboardScope), selectorSettings.postProcessMode, selectorSettings.moveLayerName);
        };

        selectorControls.btnCancel.onClick = function () {
            selectorControls.selectorDialog.close();
        };

        addVersionFooter(selectorControls.selectorDialog);
        prepareDialogWindow(selectorControls.selectorDialog, SCRIPT_NAME);
        selectorControls.selectorDialog.show();
    }

    main();

})();
