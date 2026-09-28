#target illustrator
#targetengine "ImportStylesEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

あらかじめ登録しておいたAIファイルを一覧から選び、その中身を現在のドキュメントへ取り込みます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportStyles.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n0b929db4a4ad

### Overview

Picks one of the registered AI files from a list and imports its contents into the current document.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportStyles.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportStyles";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportStyles.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportStyles.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0b929db4a4ad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* スタイル用AIファイルのベースフォルダ（初期状態は未設定。初回起動時に選択して保存）
       Base folder for style AI files (empty until the user picks one on first launch) */
    var styleLibraryFolder = "";

    /* 以前のフォルダー設定のキー（Illustrator環境設定。読み継ぎ用）/ Former preference key for the folder setting (migration) */
    var LEGACY_PREF_KEY_LIBRARY_FOLDER = "ImportStyles.libraryFolder";

    /* 登録リストの外部ファイル名（TSV）/ External library TSV
       フォーマット: label \t path \t category  ※旧形式の check(0|1) は読み込み時に無視（後方互換） */
    var LIBRARY_TSV_NAME = "ImportStyles_candidates.tsv";

    /* 貼り付け先レイヤー名 / Destination layer name */
    var IMPORT_LAYER_NAME = "// _imported";

    /* ダイアログの初期表示位置を右へずらす量（px）/ Horizontal offset for the dialog position */
    var DIALOG_OFFSET_X = 300;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（uiLang の定義より後、ダイアログを作る関数より前）に貼る。
    //    識別子はすべて KEY_SHORTCUT_* / *KeyShortcut* の名前。uiLang はコピー先のものをそのまま使う
    // 2. コントロールをすべて作り、onClick を付けたあとで1回だけ呼ぶ（keydown はウィンドウに1つ）
    //      addKeyShortcuts(dialog, {
    //          "L": alignLeftRadio,                    … ラジオ：選んで onClick
    //          "P": previewCheckbox,                   … チェックボックス：反転して onClick
    //          "Shift+R": btnReset,                    … ボタン：onClick（無ければ notify）
    //          "G": function () { toggleGuides(); },   … 関数：呼ぶだけ
    //          "Escape": { target: function () { palette.close(); }, inFields: true }
    //      }, { numericFields: [widthInput, heightInput], afterKey: updatePreview });
    //    キーは keyName と同じ綴り（"A"〜"Z"・"1"・"Semicolon"・"Escape" など。大小文字は区別しない）。
    //    修飾キーは "Shift+" / "Alt+"（option）/ "Cmd+"（⌘、Windows は Ctrl）を前に付ける
    // 3. 修飾キーは完全一致。"R" は Shift・option・⌘ を押しながらでは効かない（⌘C などを横取りしない）。
    //    Shift＋R に別の動作を付けるときは "Shift+R" を並べる
    // 4. 入力欄（edittext）・ドロップダウン・リストにフォーカスがあるときは効かない（文字は普通に入る）。
    //    数値だけの欄で効かせたいときは options.numericFields に並べる（押した文字は欄に入らない）。
    //    入力中でも効かせたいキーは { target: …, inFields: true } にする（Esc で閉じる、option＋数字など）
    // 5. 無効・非表示のコントロールは、親のパネルやグループが無効なときも含めて何もしない
    //    （親を無効にしても子の enabled は true のまま、のため親までたどる）
    // 6. 関数の戻り値：false はこのキーを使わない（文字をそのまま通す）。コントロールを返すと、そのコントロールを
    //    押したことにする（向きによってラジオが変わるときなど）。それ以外は処理済み
    // 7. ツールチップへのキー表記は options.showInTip: true で「…（L）」「… (L)」を末尾に足す。
    //    LABELS の tooltip にすでにキーを書いてあるスクリプトでは付けない（同じキーが書いてあれば二重には足さない）
    // 8. 既存の keydown 処理（bindKeyboardShortcuts・addAlignKeyHandler など）と入力欄の focus／blur による抑止は消して、これに寄せる。
    //    ↑↓キー（StepperButtons の bindSteppedArrowKeys）はそのまま残す
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "スタイル読み込み：ファイル選択", en: "Import Styles: Choose File" }
        },
        /* リスト列見出し / List column headers */
        column: {
            content: { ja: "コンテンツ", en: "Content" },
            filename: { ja: "ファイル名", en: "Filename" }
        },
        /* 検索 / Search */
        search: {
            label: { ja: "検索", en: "Search" },
            placeholder: {
                ja: "コンテンツ / ファイル名で絞り込み",
                en: "Filter by Content / Filename"
            }
        },
        /* カテゴリ / Category */
        category: {
            all: { ja: "すべて", en: "All" },
            styleBrushSymbol: { ja: "スタイル／ブラシ／シンボル", en: "Style/Brush/Symbol" },
            font: { ja: "フォント", en: "Fonts" },
            panelTitle: { ja: "カテゴリー", en: "Category" }
        },
        /* 貼り付け先 / Destination */
        destination: {
            label: { ja: "ペースト先", en: "Paste into" },
            currentLayer: { ja: "現在のレイヤー", en: "Current layer" },
            importLayer: { ja: "指定レイヤー", en: "Dedicated layer" }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            add: { ja: "追加", en: "Add" },
            load: { ja: "読み込み", en: "Load" },
            register: { ja: "登録", en: "Register" },
            folder: { ja: "フォルダー…", en: "Folder…" }
        },
        /* 入力プロンプト / Prompts */
        prompt: {
            pickAi: { ja: "追加するAIファイルを選択", en: "Choose an AI file to add" },
            pickLibraryFolder: {
                ja: "スタイル用AIファイルを置くフォルダーを選択",
                en: "Choose the folder that holds your style AI files"
            },
            enterLabel: { ja: "表示される項目名", en: "Display name for the dialog" },
            registerTitle: { ja: "読み込み設定", en: "Import Settings" }
        },
        /* 実行時メッセージ / Runtime messages */
        message: {
            openDocFirst: {
                ja: "元のドキュメントを開いてから実行してください。",
                en: "Please open the destination document first."
            },
            folderRequired: {
                ja: "スタイル用AIファイルのフォルダーが設定されていないため、終了します。",
                en: "No folder for style AI files was set, so the script has stopped."
            },
            nothingToCopy: {
                ja: "作業アートボード内にオブジェクトがないため、中止しました：\n",
                en: "Nothing on the active artboard, so the import was cancelled:\n"
            },
            layerNotEditable: {
                ja: "現在のレイヤーがロックまたは非表示のため、ペーストできません：\n",
                en: "The current layer is locked or hidden, so nothing can be pasted into it:\n"
            },
            fileNotFoundTitle: { ja: "ファイルが見つかりません", en: "File Not Found" },
            fileNotFoundBody: {
                ja: "指定したファイルが見つかりません：\n",
                en: "The specified file was not found:\n"
            },
            added: { ja: "候補を追加しました：", en: "Candidate added:" },
            overwritten: { ja: "候補を上書きしました：", en: "Candidate overwritten:" },
            synced: {
                ja: "\n\nImportStyles_candidates.tsv と現在のダイアログに反映しました。",
                en: "\n\nReflected in ImportStyles_candidates.tsv and the current dialog."
            },
            folderHint: {
                ja: "現在のフォルダー",
                en: "Current folder"
            }
        },
        /* エラー / Errors */
        error: {
            createFolder: {
                ja: "保存先フォルダを作成できませんでした。\n",
                en: "Could not create folder:\n"
            },
            copyFile: {
                ja: "ファイルをコピーできませんでした。\n",
                en: "Could not copy file to:\n"
            },
            tsvFolder: {
                ja: "候補TSVの保存先を作成できませんでした。\n",
                en: "Could not create folder for the candidates TSV:\n"
            },
            tsvOpenWrite: {
                ja: "候補TSVを書き込み用に開けませんでした。\n",
                en: "Could not open candidates TSV for writing:\n"
            }
        }
    };

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ボタン列など）/ Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // フォルダー設定 / Library Folder Setting
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は SETTINGS_STORE_* / createSettingsStore / readSettingsLegacyFile / readSettingsLegacyPreference / settingsStore*
    // 2. 寿命は今のスクリプトに合わせて選ぶ。
    //      "session"    … $.global に置く。Illustrator を終了するまで残る。#targetengine が必須（無いと毎回消える）
    //      "persistent" … Folder.userData/illustrator-scripts/<storeName>.json に書く。再起動しても残る
    //    storeName はふつう SCRIPT_NAME。ダイアログの位置は DialogPosition の部品が持つので、ここには入れない
    // 3. 既定値を1か所にまとめ、load で受け取る。戻り値は毎回新しいオブジェクト（書き換えても保存されない）
    //      var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
    //      var DEFAULT_SETTINGS = { widthPt: 10, addFrame: true, modeKey: "fit", corners: { tl: 0, tr: 0 } };
    //      var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
    //      …OK で閉じたら…
    //      settingsStore.save({ widthPt: …, addFrame: …, modeKey: …, corners: { tl: …, tr: … } });
    //    型は既定値に合わせる（数値の既定値には "12" も 12 として読む。真偽は "1"/"0"/"true"/"false" も読む）。
    //    合わない値・既定値に無い項目は捨てて既定値を使う。{} と null の既定値は中身を問わずそのまま受け取る
    //    （名前をキーにしたプリセット集など）。配列は配列ならそのまま受け取る
    // 4. 保存できるのは文字列・数値・真偽・null と、その配列・入れ子のオブジェクトだけ。
    //    DOM オブジェクト・File・関数は入れない（パスは fsName の文字列で持つ）。長さは pt で持つ
    // 5. 旧形式の設定を読み継ぐときは、3つ目の引数に legacy 関数を渡す。
    //    新しい保存が1度も無いとき（ファイルが無い・$.global に無い）だけ呼ばれ、戻り値を保存値として既定値と突き合わせる。
    //    旧ファイル・旧キーは消さない。キー名が変わったときは legacy の中で詰め替える
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");  … key=value / toSource / JSON を自動判別
    //      } });
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyPreference("SmartTextFindReplace/settings");  … app.preferences の文字列
    //      } });
    // 6. clear() は保存を消す。legacy を渡したストアでは空の保存（{}）を書き、旧設定が戻ってこないようにする
    // 7. 失敗は例外にせず、load は既定値、save は false を返す（$.writeln に理由を出す）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^﻿/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* フォルダー設定の保存先（Folder.userData/illustrator-scripts/ImportStyles.json）。
       以前は Illustrator の環境設定に保存していたので、新しい保存が無いときだけそこから読み継ぐ
       Folder setting store; the value formerly in Illustrator's preferences is read only until the first save */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent", {
        legacy: function () {
            var legacyFolder = app.preferences.getStringPreference(LEGACY_PREF_KEY_LIBRARY_FOLDER);
            return legacyFolder ? { libraryFolder: String(legacyFolder) } : null;
        }
    });
    var DEFAULT_SETTINGS = { libraryFolder: "" }; /* 既定値 / defaults */

    /**
     * 保存済みのフォルダー設定を読み込む
     * @returns {string} フォルダーのパス。未設定なら空文字
     */
    function loadSavedLibraryFolder() {
        return settingsStore.load(DEFAULT_SETTINGS).libraryFolder;
    }

    /**
     * フォルダーを選択して設定を保存する
     * @returns {string} フォルダーのパス（末尾に /）。キャンセルなら空文字
     */
    function pickAndSaveLibraryFolder() {
        var pickedFolder = Folder.selectDialog(getLabel("prompt.pickLibraryFolder"));
        if (!pickedFolder) return "";
        var folderPath = pickedFolder.fsName + "/";
        settingsStore.save({ libraryFolder: folderPath });
        return folderPath;
    }

    /* 有効なフォルダー設定を確定（未設定・不在なら選択させる）/ Resolve a usable folder (ask when unset or missing) */
    function resolveLibraryFolder() {
        var savedPath = loadSavedLibraryFolder();
        if (savedPath !== "" && Folder(savedPath).exists) return savedPath;
        return pickAndSaveLibraryFolder();
    }

    /* 登録リストのTSVファイル / The library TSV file */
    function getLibraryTsvFile() {
        return File(Folder(styleLibraryFolder).fsName + "/" + LIBRARY_TSV_NAME);
    }

    // =========================================
    // スタイルライブラリ / Style Library
    // =========================================

    /* カテゴリ名の正規化値 / Canonical category names */
    var CATEGORY_STYLE = "スタイル／ブラシ／シンボル";
    var CATEGORY_FONT = "フォント";

    /* カテゴリ名の正規化（フォント系だけ判定し、他はすべてスタイル扱い）/ Normalize category (font-like → Font, else Style) */
    function normalizeCategory(rawCategory) {
        var text = String(rawCategory != null ? rawCategory : "");
        var compact = text.toLowerCase().replace(/\s+/g, "");
        var isFont = /font/.test(compact) || text.indexOf("フォント") !== -1 || text.indexOf("フォンツ") !== -1;
        return isFont ? CATEGORY_FONT : CATEGORY_STYLE;
    }

    /* 保存済みファイル名を実ファイルに合わせて正規化（旧版が書いた%エンコード名の救済）
       Resolve a stored file name against the actual file (recovers legacy percent-encoded names) */
    function resolveStoredFileName(storedName) {
        if (!/%[0-9A-Fa-f]{2}/.test(storedName)) return storedName;
        var decoded = decodeFileName(storedName);
        // %を含む実在の名前を壊さないよう、復号後に実在する場合だけ採用
        // Adopt the decoded form only when it actually exists, so real names with % survive
        return (decoded !== storedName && File(styleLibraryFolder + decoded).exists) ? decoded : storedName;
    }

    /* TSVの1行をライブラリ項目へ変換（不正行は null）/ Parse one TSV line into a library item (null if invalid) */
    function parseLibraryLine(line) {
        if (!line || line.charAt(0) === "#") return null;
        var columns = line.split("\t");
        if (columns.length < 2 || columns[1] === "") return null;

        // 旧形式: label, path, check(0|1), category / 新形式: label, path, category
        var categoryIndex = (columns.length >= 3 && /^(0|1)$/.test(columns[2])) ? 3 : 2;
        var rawCategory = (columns.length > categoryIndex && columns[categoryIndex] !== "") ? columns[categoryIndex] : CATEGORY_STYLE;
        var fileName = resolveStoredFileName(columns[1]);

        return {
            label: (columns[0] !== "") ? columns[0] : fileName.replace(/\.[^\.]+$/, ""),
            fileName: fileName,
            category: normalizeCategory(rawCategory)
        };
    }

    /* フォルダー内のAIファイルを走査して項目化（TSV未提供時のフォールバック）
       Scan the folder for AI files (fallback when no TSV is available) */
    function scanFolderForAiFiles() {
        var libraryFolder = Folder(styleLibraryFolder);
        if (!libraryFolder.exists) return [];

        // 拡張子の大小を問わず拾う（文字列マスクはmacOSで大小を区別）
        // Match the extension case-insensitively (string masks are case-sensitive on macOS)
        var aiFiles = libraryFolder.getFiles(function(entry) {
            return (entry instanceof File) && /\.ai$/i.test(decodeFileName(entry.name));
        });

        var scanned = [];
        for (var i = 0; i < aiFiles.length; i++) {
            var fileName = decodeFileName(aiFiles[i].name);
            scanned.push({
                label: fileName.replace(/\.[^\.]+$/, ""),
                fileName: fileName,
                category: normalizeCategory(fileName)
            });
        }
        return scanned;
    }

    /* ライブラリをTSVから読み込み（無ければフォルダー内のAIファイルを列挙）
       Load the library from TSV (falls back to scanning the folder) */
    function loadStyleLibrary() {
        var tsvFile = getLibraryTsvFile();
        if (!tsvFile.exists) return scanFolderForAiFiles();

        tsvFile.encoding = "UTF-8";
        if (!tsvFile.open("r")) return scanFolderForAiFiles();

        var loaded = [];
        while (!tsvFile.eof) {
            var styleItem = parseLibraryLine(tsvFile.readln());
            if (styleItem) loaded.push(styleItem);
        }
        tsvFile.close();
        return loaded.length ? loaded : scanFolderForAiFiles();
    }

    /* 現在のライブラリをTSVに保存（上書き）/ Save the current library to TSV (overwrite) */
    function saveStyleLibrary(styleItems) {
        var tsvFile = getLibraryTsvFile();
        var parentFolder = tsvFile.parent;
        if (parentFolder && !parentFolder.exists && !parentFolder.create()) {
            alert(getLabel("error.tsvFolder") + parentFolder.fsName);
            return false;
        }
        tsvFile.encoding = "UTF-8";
        if (!tsvFile.open("w")) {
            alert(getLabel("error.tsvOpenWrite") + tsvFile.fsName);
            return false;
        }

        tsvFile.writeln("# label\tpath\tcategory");
        for (var i = 0; i < styleItems.length; i++) {
            var styleItem = styleItems[i];
            if (!styleItem || !styleItem.fileName) continue;
            tsvFile.writeln(styleItem.label + "\t" + styleItem.fileName + "\t" + styleItem.category);
        }
        tsvFile.close();
        return true;
    }

    /* 現在のライブラリ（フォルダー確定後に読み込む）/ Current library (loaded once the folder is resolved) */
    var styleLibrary = [];

    /* ファイル名でライブラリ項目を検索し index を返す（無ければ -1）/ Find item index by file name (-1 if none) */
    function findLibraryIndex(fileName) {
        for (var i = 0; i < styleLibrary.length; i++) {
            if (styleLibrary[i] && styleLibrary[i].fileName === fileName) return i;
        }
        return -1;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* ファイル名を表示用にデコード / Decode a file name for display */
    function decodeFileName(fileName) {
        try {
            return decodeURI(fileName);
        } catch (e) {
            return fileName;
        }
    }

    /* 大文字小文字を無視した部分一致 / Case-insensitive contains */
    function containsIgnoreCase(haystack, needle) {
        return String(haystack).toLowerCase().indexOf(String(needle).toLowerCase()) !== -1;
    }

    /* フォルダを確保（存在しなければ作成）/ Ensure a folder exists (create if missing) */
    function ensureFolder(folder) {
        if (!folder) return false;
        if (folder.exists || folder.create()) return true;
        alert(getLabel("error.createFolder") + folder.fsName);
        return false;
    }

    /* 同名上書きでファイルをコピー / Copy a file with overwrite */
    function copyFileOverwrite(sourceFile, destFile) {
        if (!sourceFile || !destFile) return false;
        if (sourceFile.fsName === destFile.fsName) return true; // 同一ファイルなら何もしない / Same file
        if (destFile.exists) destFile.remove();
        if (sourceFile.copy(destFile.fsName)) return true;
        alert(getLabel("error.copyFile") + destFile.fsName);
        return false;
    }

    /* 取り込み先レイヤーを取得（無ければ作成）/ Get or create the destination layer */
    function getOrCreateImportLayer(doc) {
        var importLayer;
        try {
            importLayer = doc.layers.getByName(IMPORT_LAYER_NAME);
        } catch (e) {
            importLayer = doc.layers.add();
            importLayer.name = IMPORT_LAYER_NAME;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
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
    // ライブラリ登録フロー / Register Flow
    // =========================================

    /* 表示名とカテゴリを入力させる（キャンセルで null）/ Ask for label + category (null on cancel) */
    function askLabelAndCategory(baseName) {
        var registerWindow = new Window("dialog", getLabel("prompt.registerTitle") + " " + SCRIPT_VERSION);
        setupWindow(registerWindow);

        // 表示名入力 / Label field
        var labelColumn = registerWindow.add("group");
        labelColumn.orientation = "column";
        labelColumn.alignChildren = ["fill", "top"];
        labelColumn.add("statictext", undefined, labelText("prompt.enterLabel"));
        var labelField = labelColumn.add("edittext", undefined, baseName);
        labelField.characters = 20;

        // カテゴリ選択（パネル＋ラジオ）/ Category panel with radios
        var categoryPanel = registerWindow.add("panel", undefined, getLabel("category.panelTitle"));
        setupPanel(categoryPanel, 6);
        var styleRadio = categoryPanel.add("radiobutton", undefined, getLabel("category.styleBrushSymbol"));
        var fontRadio = categoryPanel.add("radiobutton", undefined, getLabel("category.font"));

        // ファイル名からカテゴリを自動推定 / Auto-detect category from the file name
        var looksLikeFont = containsIgnoreCase(baseName, "font") || baseName.indexOf("フォント") !== -1;
        fontRadio.value = looksLikeFont;
        styleRadio.value = !looksLikeFont;

        // ボタン行（パネル幅いっぱいに広げない）/ Button row (buttons must not stretch)
        var registerButtonRow = addButtonRow(registerWindow, { centered: true });
        var btnRegisterCancel = registerButtonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnRegisterOK = registerButtonRow.rowGroup.add("button", undefined, getLabel("button.register"), { name: "ok" });

        prepareDialogWindow(registerWindow, SCRIPT_NAME + "_register");
        if (registerWindow.show() !== 1) return null;

        // 入力値整形 / Sanitize the label
        var cleanLabel = String(labelField.text || "").replace(/[\r\n\t]/g, " ").replace(/^\s+|\s+$/g, "");

        return {
            label: (cleanLabel !== "") ? cleanLabel : baseName,
            category: fontRadio.value ? CATEGORY_FONT : CATEGORY_STYLE
        };
    }

    /* AIファイルをフォルダへコピーしてライブラリに登録（登録したファイル名を返す／中断で null）
       Copy an AI file into the library folder and register it (returns the file name, null if cancelled) */
    function registerStyleFile() {
        // 1) ファイル選択 / Pick an AI file
        var pickedFile = File.openDialog(getLabel("prompt.pickAi"), "*.ai");
        if (!pickedFile) return null;

        // 2) 保存先フォルダを確保してコピー / Ensure the folder, then copy
        var libraryFolder = Folder(styleLibraryFolder);
        if (!ensureFolder(libraryFolder)) return null;

        var destFile = File(libraryFolder.fsName + "/" + pickedFile.name);
        if (!copyFileOverwrite(pickedFile, destFile)) return null;

        // 3) 表示名とカテゴリを入力 / Ask for label + category
        var destFileName = decodeFileName(destFile.name); // TSVには生のファイル名で保存 / Store the plain name in the TSV
        var userInput = askLabelAndCategory(destFileName.replace(/\.[^\.]+$/, ""));
        if (!userInput) return null;

        // 4) 既存項目を上書き、無ければ追加 / Overwrite the existing item or append
        var styleItem = { label: userInput.label, fileName: destFileName, category: userInput.category };
        var existingIndex = findLibraryIndex(destFileName);
        if (existingIndex >= 0) {
            styleLibrary[existingIndex] = styleItem;
        } else {
            styleLibrary.push(styleItem);
        }

        // 5) TSVへ保存 / Save to TSV
        if (!saveStyleLibrary(styleLibrary)) return null;

        alert((existingIndex >= 0 ? getLabel("message.overwritten") : getLabel("message.added")) + "\n" +
            userInput.label + "  (" + destFileName + ")" + getLabel("message.synced"));
        return destFileName;
    }

    // =========================================
    // ダイアログの部品 / Dialog Parts
    // =========================================

    /* カテゴリ絞り込みラジオ（初期は「すべて」）/ Category filter radios (defaults to All) */
    function buildCategoryFilterRow(parent) {
        var categoryFilterRow = parent.add("group");
        setupRow(categoryFilterRow, "center");

        var allRadio = categoryFilterRow.add("radiobutton", undefined, getLabel("category.all"));
        var styleRadio = categoryFilterRow.add("radiobutton", undefined, getLabel("category.styleBrushSymbol"));
        var fontRadio = categoryFilterRow.add("radiobutton", undefined, getLabel("category.font"));
        allRadio.value = true;

        return { allRadio: allRadio, styleRadio: styleRadio, fontRadio: fontRadio };
    }

    /* 検索欄と検索ボタン / Search field and Search button */
    function buildSearchRow(parent) {
        var searchRow = parent.add("group");
        setupRow(searchRow, "center");
        searchRow.margins = [0, 5, 0, 15];

        var searchField = searchRow.add("edittext", undefined, "");
        searchField.characters = 24;
        searchField.helpTip = getLabel("search.placeholder");

        var searchButton = searchRow.add("button", undefined, getLabel("search.label"));
        searchButton.alignment = "left";

        return { field: searchField, button: searchButton };
    }

    /* ヘッダ付き2列のListBox / Headered 2-column ListBox */
    function buildStyleListBox(parent) {
        var contentColumnWidth = 140;
        var filenameColumnWidth = 180;

        var styleListBox = parent.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [getLabel("column.content"), getLabel("column.filename")],
            columnWidths: [contentColumnWidth, filenameColumnWidth]
        });
        styleListBox.alignment = ["fill", "fill"];
        styleListBox.preferredSize = [contentColumnWidth + filenameColumnWidth + 20, 200];
        return styleListBox;
    }

    /* 貼り付け先ラジオ（初期は現在のレイヤー）/ Destination radios (defaults to the current layer) */
    function buildDestinationRow(parent) {
        var destinationRow = parent.add("group");
        setupRow(destinationRow, "center");
        destinationRow.add("statictext", undefined, labelText("destination.label"));

        var currentLayerRadio = destinationRow.add("radiobutton", undefined, getLabel("destination.currentLayer"));
        var importLayerLabel = (uiLang === "ja") ?
            getLabel("destination.importLayer") + "（" + IMPORT_LAYER_NAME + "）" :
            getLabel("destination.importLayer") + " (" + IMPORT_LAYER_NAME + ")";
        var importLayerRadio = destinationRow.add("radiobutton", undefined, importLayerLabel);
        currentLayerRadio.value = true;

        return { currentLayerRadio: currentLayerRadio, importLayerRadio: importLayerRadio };
    }

    /* 下部ボタン行：左（フォルダー）－スペーサー－右（キャンセル／追加／読み込み）/ Footer buttons */
    function buildFooterRow(parent) {
        var footerRow = addButtonRow(parent);

        var btnFolder = footerRow.leftGroup.add("button", undefined, getLabel("button.folder"));

        var btnCancel = footerRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnAdd = footerRow.rightGroup.add("button", undefined, getLabel("button.add"));
        var btnLoad = footerRow.rightGroup.add("button", undefined, getLabel("button.load"), { name: "ok" });

        return {
            cancelButton: btnCancel,
            folderButton: btnFolder,
            addButton: btnAdd,
            loadButton: btnLoad
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ライブラリから1つ選択し、AIファイルのパスと貼り付け先を返す（キャンセルで null）
       Choose one item and return the AI file path plus the destination (null if cancelled) */
    function showStyleLibraryDialog() {
        var libraryWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(libraryWindow);
        libraryWindow.onShow = function() {
            libraryWindow.location = [libraryWindow.location[0] + DIALOG_OFFSET_X, libraryWindow.location[1]];
        };

        var categoryFilter = buildCategoryFilterRow(libraryWindow);
        var search = buildSearchRow(libraryWindow);
        var styleListBox = buildStyleListBox(libraryWindow);
        var destination = buildDestinationRow(libraryWindow);
        var footer = buildFooterRow(libraryWindow);

        // 絞り込みの状態 / Filter state
        var activeCategory = null; // null = 全件 / null = all
        var activeQuery = "";

        /* 絞り込み条件に一致するか / Does the item match the current filter? */
        function matchesFilter(styleItem) {
            if (activeCategory != null && styleItem.category !== activeCategory) return false;
            if (activeQuery === "") return true;
            return containsIgnoreCase(styleItem.label, activeQuery) ||
                containsIgnoreCase(styleItem.fileName, activeQuery);
        }

        /* 絞り込み結果でリストを作り直す / Rebuild the list from the current filter */
        function refreshList() {
            styleListBox.removeAll();
            for (var i = 0; i < styleLibrary.length; i++) {
                var styleItem = styleLibrary[i];
                if (!styleItem || !styleItem.fileName || !matchesFilter(styleItem)) continue;

                var row = styleListBox.add("item", styleItem.label);
                row.subItems[0].text = styleItem.fileName;
                row.styleItem = styleItem; // 参照を保持 / Keep the reference
            }
            if (styleListBox.items.length > 0) styleListBox.selection = 0;
            libraryWindow.layout.layout(true);
        }

        /* 選択行を確定してダイアログを閉じる / Confirm the selected row */
        function confirmSelection() {
            if (styleListBox.selection) libraryWindow.close(1);
        }

        /* カテゴリ絞り込みを切り替え / Switch the category filter */
        function applyCategory(categoryName) {
            return function() {
                activeCategory = categoryName;
                refreshList();
            };
        }
        categoryFilter.allRadio.onClick = applyCategory(null);
        categoryFilter.styleRadio.onClick = applyCategory(CATEGORY_STYLE);
        categoryFilter.fontRadio.onClick = applyCategory(CATEGORY_FONT);

        /* 検索欄の内容を絞り込みに反映 / Apply the search field to the filter */
        function applySearch() {
            activeQuery = String(search.field.text || "");
            refreshList();
        }

        // 検索はボタン押下時と検索欄の確定時のみ適用（ライブ絞り込みはしない）
        // Apply the query on the button click and when the field is committed (no live filtering)
        search.button.onClick = applySearch;
        search.field.onChange = applySearch;

        // キーボード：Enterで決定（検索欄では検索）、⌘（Ctrl）+ Fで検索欄へ
        // Keyboard: Enter confirms (or searches while in the field), Cmd/Ctrl+F focuses the search field
        styleListBox.onDoubleClick = confirmSelection;
        // 入力欄・リストにフォーカスがあっても効かせる（inFields）。数値だけの欄は無い
        // Both keys work while the field or list has focus (inFields); there are no numeric-only fields
        addKeyShortcuts(libraryWindow, {
            "Enter": {
                target: function(keyEvent) {
                    if (keyEvent.target === search.field) {
                        applySearch();
                        return true;
                    }
                    /* 決定後もキーは既定の処理へ通す（従来どおり）/ Let the key reach the default handling, as before */
                    confirmSelection();
                    return false;
                },
                inFields: true
            },
            "Cmd+F": {
                target: function() { search.field.active = true; },
                inFields: true
            }
        });

        // 追加：AIファイルを登録してリストへ反映 / Add: register an AI file and reflect it
        footer.addButton.onClick = function() {
            var addedFileName = registerStyleFile();
            if (!addedFileName) return;

            // 追加した項目が絞り込みで隠れないよう「すべて」に戻す / Reset to All so the new item stays visible
            categoryFilter.allRadio.value = true;
            activeCategory = null;
            refreshList();
            for (var i = 0; i < styleListBox.items.length; i++) {
                if (styleListBox.items[i].styleItem.fileName === addedFileName) {
                    styleListBox.selection = i;
                    break;
                }
            }
        };

        // フォルダー：保存先を選び直してライブラリを読み込み直す / Folder: re-pick the folder and reload the library
        function updateFolderHint() {
            footer.folderButton.helpTip = labelValueText("message.folderHint", styleLibraryFolder);
        }
        updateFolderHint();
        footer.folderButton.onClick = function() {
            var pickedPath = pickAndSaveLibraryFolder();
            if (pickedPath === "") return;
            styleLibraryFolder = pickedPath;
            styleLibrary = loadStyleLibrary();
            updateFolderHint();
            refreshList();
        };

        footer.loadButton.onClick = confirmSelection;

        // 初期表示 → ダイアログ表示 → 選択を返却 / Init, show, return the selection
        refreshList();
        prepareDialogWindow(libraryWindow, SCRIPT_NAME);
        if (libraryWindow.show() !== 1 || !styleListBox.selection) return null;
        return {
            filePath: styleLibraryFolder + styleListBox.selection.styleItem.fileName,
            useImportLayer: destination.importLayerRadio.value === true
        };
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================

    /* ライブラリ選択 → 対象AIをコピー → 元ドキュメントへ貼り付け / Choose, copy from the AI, paste into the original */
    function main() {
        // 貼り付け先のドキュメントを先に確認 / Check for the destination document first
        if (app.documents.length === 0) {
            alert(getLabel("message.openDocFirst"));
            return;
        }
        var originalDoc = app.activeDocument;

        // フォルダー設定を確定し、ライブラリを読み込む / Resolve the folder setting, then load the library
        styleLibraryFolder = resolveLibraryFolder();
        if (styleLibraryFolder === "") {
            alert(getLabel("message.folderRequired"));
            return;
        }
        styleLibrary = loadStyleLibrary();

        var choice = showStyleLibraryDialog();
        if (!choice) return;

        // 現在のレイヤーへペーストする場合は、編集可能かを確認 / When pasting into the current layer, make sure it is editable
        var currentLayer = originalDoc.activeLayer;
        if (!choice.useImportLayer && (currentLayer.locked || !currentLayer.visible)) {
            alert(getLabel("message.layerNotEditable") + currentLayer.name);
            return;
        }

        // 選択したファイルを開く / Open the chosen file
        var styleFile = new File(choice.filePath);
        if (!styleFile.exists) {
            alert(getLabel("message.fileNotFoundTitle") + "\n" + getLabel("message.fileNotFoundBody") + decodeFileName(styleFile.name));
            return;
        }
        var styleDoc = app.open(styleFile);

        // 作業アートボード内をすべてコピーし、保存せずに閉じる / Copy in-artboard objects, close without saving
        app.executeMenuCommand("selectallinartboard");
        if (styleDoc.selection.length === 0) {
            // 選択が空のままコピーすると、以前のクリップボードが貼り付けられるため中止
            // Copying an empty selection would paste whatever was in the clipboard before
            styleDoc.close(SaveOptions.DONOTSAVECHANGES);
            app.activeDocument = originalDoc;
            alert(getLabel("message.nothingToCopy") + decodeFileName(styleFile.name));
            return;
        }
        app.executeMenuCommand("copy");
        styleDoc.close(SaveOptions.DONOTSAVECHANGES);

        // 貼り付け先を決めてからペースト（現在のレイヤーはそのまま）/ Set the destination, then paste
        app.activeDocument = originalDoc;
        if (choice.useImportLayer) {
            originalDoc.activeLayer = getOrCreateImportLayer(originalDoc);
        }
        app.executeMenuCommand("paste");
    }

    main();

})();
