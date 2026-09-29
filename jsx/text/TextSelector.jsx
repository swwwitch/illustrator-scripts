#target illustrator
#targetengine "TextSelectorEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストフレームを、複数の条件で一括選択します。
属性・テキストの種類・文字列で絞り込み、選択後に非表示やレイヤー移動、一括編集も行えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextSelector.md

### Overview

Selects text frames across the document by a combination of conditions.
You can filter by attribute, text kind and string, then hide, move to a layer, or bulk-edit what was selected.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextSelector.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextSelector";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextSelector.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextSelector.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「選択後に移動」先のレイヤー名 / Destination layer name for "move after selection" */
    var TEXT_LAYER_NAME = "_text";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];          /* パネルの余白 / panel margins */
    var OUTER_PANEL_MARGINS = [15, 20, 15, 15];    /* 選択条件・対象アートボードパネルの余白 / margins of the outer panels */
    var ATTRIBUTE_LABEL_WIDTH = 150;               /* 属性ラジオの幅 / width of the attribute radios */
    var PREVIEW_CHARACTERS = 24;                   /* 属性プレビューの文字数 / width of an attribute preview */
    var KEYWORD_CHARACTERS = 22;                   /* 検索文字列欄の文字数 / width of the keyword field */
    var BULK_INPUT_SIZE = [300, 70];               /* 一括編集の入力欄 / size of the bulk edit field */

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
            bulkEditTitle: { ja: "テキストを一括変更", en: "Bulk Edit Text" }
        },
        panel: {
            selection: { ja: "選択条件", en: "Selection Criteria" },
            artboardScope: { ja: "対象とするアートボード", en: "Target Artboard" },
            attribute: { ja: "属性で選択", en: "Select by Attribute" },
            textType: { ja: "テキストの種類", en: "Text Type" },
            textMatch: { ja: "文字列で選択", en: "Select by String" },
            postProcess: { ja: "選択後の処理", en: "After Selection" }
        },
        radio: {
            artboardAll: { ja: "すべて", en: "All" },
            artboardCurrent: { ja: "現在のアートボードのみ", en: "Current Artboard Only" },
            fontFamily: { ja: "フォントファミリー", en: "Font Family" },
            fontFamilyStyle: { ja: "+ スタイル", en: "+ Style" },
            fontFamilyStyleSize: { ja: "+ スタイルとサイズ", en: "+ Style and Size" },
            fontSize: { ja: "フォントサイズ", en: "Font Size" },
            textFillColor: { ja: "テキストカラー", en: "Text Color" },
            opacity: { ja: "不透明度", en: "Opacity" },
            allText: { ja: "すべて", en: "All" },
            pointText: { ja: "ポイント文字", en: "Point Text" },
            areaText: { ja: "エリア内文字", en: "Area Text" },
            pathText: { ja: "パス上文字", en: "Path Text" },
            exactMatch: { ja: "完全一致", en: "Exact" },
            containsMatch: { ja: "部分一致", en: "Contains" },
            startsWith: { ja: "先頭一致", en: "Starts With" },
            endsWith: { ja: "末尾一致", en: "Ends With" },
            regexMatch: { ja: "正規表現", en: "Regex" },
            noPostProcess: { ja: "なし", en: "None" },
            hideAfterSelection: { ja: "選択後に非表示", en: "Hide after selection" },
            hideOthers: { ja: "選択したテキスト以外を非表示", en: "Hide all but selected" },
            moveToTextLayer: { ja: "選択後に「_text」レイヤーへ移動", en: "Move to “_text” layer after selection" },
            bulkEdit: { ja: "一括編集", en: "Bulk edit" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            fontFamily: {
                ja: "選択中のテキストと同じフォントファミリーのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font family as the selected text using Illustrator’s built-in command."
            },
            fontFamilyStyle: {
                ja: "選択中のテキストと同じフォントファミリー＋スタイルのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font family and style as the selected text using Illustrator’s built-in command."
            },
            fontFamilyStyleSize: {
                ja: "選択中のテキストと同じフォントファミリー＋スタイル＋サイズのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font family, style, and size as the selected text using Illustrator’s built-in command."
            },
            fontSize: {
                ja: "選択中のテキストと同じフォントサイズのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same font size as the selected text using Illustrator’s built-in command."
            },
            textFillColor: {
                ja: "選択中のテキストと同じテキストカラーのテキストを、Illustrator標準コマンドで検索します。",
                en: "Find text with the same text color as the selected text using Illustrator’s built-in command."
            },
            opacity: {
                ja: "選択中のオブジェクトと同じ不透明度のオブジェクトを、Illustrator標準コマンドで検索します。テキスト以外のオブジェクトも対象になります。",
                en: "Find objects with the same opacity as the selected object using Illustrator’s built-in command. Non-text objects are also matched."
            },
            allText: {
                ja: "ドキュメント内のすべてのテキストフレームを選択します。ショートカット：Option + Q",
                en: "Select all text frames in the document. Shortcut: Option + Q"
            },
            pointText: { ja: "ポイント文字だけを選択します。ショートカット：Option + W", en: "Select point text only. Shortcut: Option + W" },
            areaText: { ja: "エリア内文字だけを選択します。ショートカット：Option + E", en: "Select area text only. Shortcut: Option + E" },
            pathText: { ja: "パス上文字だけを選択します。", en: "Select path text only." },
            keywordInput: { ja: "検索対象にする文字列を入力します。空欄では実行できません。", en: "Enter the search string. This cannot be empty." },
            exactMatch: {
                ja: "入力した文字列と完全に一致するテキストを選択します。大小区別あり。ショートカット：Option + A",
                en: "Select text that exactly matches the entered string. Case-sensitive. Shortcut: Option + A"
            },
            startsWith: {
                ja: "入力した文字列で始まるテキストを選択します。大小区別あり。ショートカット：Option + B",
                en: "Select text that starts with the entered string. Case-sensitive. Shortcut: Option + B"
            },
            endsWith: {
                ja: "入力した文字列で終わるテキストを選択します。大小区別あり。ショートカット：Option + D",
                en: "Select text that ends with the entered string. Case-sensitive. Shortcut: Option + D"
            },
            containsMatch: {
                ja: "入力した文字列を含むテキストを選択します。大小区別あり。ショートカット：Option + I",
                en: "Select text that contains the entered string. Case-sensitive. Shortcut: Option + I"
            },
            regexMatch: {
                ja: "入力した正規表現に一致するテキストを選択します。大小区別あり。ショートカット：Option + R",
                en: "Select text that matches the entered regular expression. Case-sensitive. Shortcut: Option + R"
            },
            artboardAll: { ja: "ドキュメント内のすべてのアートボードを対象に選択します。", en: "Select across all artboards in the document." },
            artboardCurrent: {
                ja: "現在のアートボードに重なるオブジェクトだけを対象に選択します。",
                en: "Select only objects that overlap the active artboard."
            },
            noPostProcess: { ja: "選択後に追加処理を行いません。", en: "Do not apply any additional processing after selection." },
            hideAfterSelection: { ja: "選択されたテキストを非表示にします。", en: "Hide the selected text." },
            hideOthers: {
                ja: "選択されたテキスト以外のオブジェクトを非表示にします。ロックされたオブジェクトは除きます。",
                en: "Hide all objects except the selected text. Locked objects are left untouched."
            },
            moveToTextLayer: {
                ja: "選択されたテキストを「_text」レイヤーへ移動します。レイヤーがない場合は作成します。",
                en: "Move the selected text to the “_text” layer. The layer will be created if it does not exist."
            },
            bulkEdit: {
                ja: "選択されたテキストの内容をまとめて置き換えます。書式は維持されます。",
                en: "Replace contents of all selected text frames at once. Formatting is preserved."
            }
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
     * 選択中のオブジェクトの文字列を取り出す
     * @param {Object} selectionItem - 選択中のオブジェクト
     * @returns {string} 文字列（無ければ空文字）
     */
    function getTextContentsFromSelectionItem(selectionItem) {
        if (!selectionItem) {
            return "";
        }

        /* contents を持たないオブジェクトでは例外になることがある / may throw on items without contents */
        try {
            if (typeof selectionItem.contents === "string") {
                return selectionItem.contents;
            }
        } catch (contentsError) {
        }

        var textRange = getTextRangeFromSelectionItem(selectionItem);
        if (textRange && typeof textRange.contents === "string") {
            return textRange.contents;
        }

        return "";
    }

    /**
     * 選択中の最初のテキストの文字列を返す（検索文字列の初期値）
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {string} 文字列（無ければ空文字）
     */
    function getSelectedTextString(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            var textContents = getTextContentsFromSelectionItem(selectedItems[i]);
            if (textContents) {
                return textContents;
            }
        }
        return "";
    }

    /**
     * 属性プレビューの初期値（未取得を表す「—」）を返す
     * @returns {Object} family / familyStyle / familyStyleSize / size / opacity
     */
    function makeEmptyAttributePreview() {
        return {
            family: "—",
            familyStyle: "—",
            familyStyleSize: "—",
            size: "—",
            opacity: "—"
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
            attributePreview.familyStyle = familyName ? familyName + (styleName ? " / " + styleName : "") : "—";
            attributePreview.familyStyleSize = attributePreview.familyStyle !== "—" ? attributePreview.familyStyle + " / " + sizeText : "—";
            attributePreview.size = sizeText;
        } catch (attributePreviewError) {
            return attributePreview;
        }

        return attributePreview;
    }

    /**
     * 選択中のオブジェクトから不透明度のプレビューを作る
     * @param {Object} selectionItem - 選択中のオブジェクト
     * @returns {string} 不透明度の表示（無ければ「—」）
     */
    function getOpacityPreviewFromSelectionItem(selectionItem) {
        /* TextRange などでは opacity の参照で例外になることがある / may throw on a TextRange and the like */
        try {
            if (selectionItem && typeof selectionItem.opacity === "number") {
                return String(Math.round(selectionItem.opacity * 100) / 100) + " %";
            }
        } catch (opacityPreviewError) {
        }
        return "—";
    }

    /**
     * 選択中の最初のテキストから文字属性と不透明度のプレビューを作る
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {Object} makeEmptyAttributePreview() と同じ形のプレビュー
     */
    function getSelectedTextAttributePreview(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            var textRange = getTextRangeFromSelectionItem(selectedItems[i]);
            if (textRange) {
                var rangePreview = getTextAttributePreviewFromTextRange(textRange);
                rangePreview.opacity = getOpacityPreviewFromSelectionItem(selectedItems[i]);
                return rangePreview;
            }
        }

        return makeEmptyAttributePreview();
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
     * 文字列条件の入力を確かめる（空欄・不正な正規表現はメッセージを出す）
     * @param {string} keyword - 検索文字列
     * @param {string} matchMode - 一致のモード
     * @returns {Object|null} { regex: RegExp|null }。入力が不正なら null
     */
    function validateTextMatchInput(keyword, matchMode) {
        if (!keyword) {
            alert(getLabel("alert.emptyKeyword"));
            return null;
        }
        if (matchMode === "regex") {
            /* 不正なパターンは RegExp が例外を投げる / RegExp throws on an invalid pattern */
            try {
                return { regex: new RegExp(keyword) };
            } catch (regexError) {
                alert(getLabel("alert.invalidRegex"));
                return null;
            }
        }
        return { regex: null };
    }

    /**
     * テキストの種類に対応する判定関数を作る
     * @param {string} textType - "all" / "point" / "area" / "path"
     * @returns {Function} テキストフレームを受け取って true/false を返す関数
     */
    function buildTextTypePredicate(textType) {
        if (textType === "point") {
            return function (textFrame) { return textFrame.kind === TextType.POINTTEXT; };
        }
        if (textType === "area") {
            return function (textFrame) { return textFrame.kind === TextType.AREATEXT; };
        }
        if (textType === "path") {
            return function (textFrame) { return textFrame.kind === TextType.PATHTEXT; };
        }
        /* "all" は全件一致 / "all" matches everything */
        return function () { return true; };
    }

    /**
     * 現在のアートボードの矩形を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {number[]|null} [左, 上, 右, 下]（取れなければ null）
     */
    function getActiveArtboardRect(doc) {
        /* アートボードの取得に失敗したら絞り込まない / skip the filter when the artboard cannot be read */
        try {
            var activeIndex = doc.artboards.getActiveArtboardIndex();
            return doc.artboards[activeIndex].artboardRect; /* [left, top, right, bottom] */
        } catch (artboardRectError) {
            return null;
        }
    }

    /**
     * オブジェクトが矩形に重なるかを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @param {number[]|null} artboardRect - [左, 上, 右, 下]（null なら常に true）
     * @returns {boolean} 重なれば true
     */
    function isItemWithinArtboardRect(pageItem, artboardRect) {
        if (!artboardRect) {
            return true;
        }
        /* 境界が読めないオブジェクトは残す / keep items whose bounds cannot be read */
        try {
            var itemBounds = pageItem.visibleBounds; /* [left, top, right, bottom] */
            var overlapsHorizontally = itemBounds[2] >= artboardRect[0] && itemBounds[0] <= artboardRect[2];
            var overlapsVertically = itemBounds[1] >= artboardRect[3] && itemBounds[3] <= artboardRect[1];
            return overlapsHorizontally && overlapsVertically;
        } catch (boundsError) {
            return true;
        }
    }

    /**
     * 現在のアートボードに重なるオブジェクトだけに絞り込む
     * @param {Array} pageItems - 対象のオブジェクト
     * @returns {Array} 絞り込んだオブジェクト
     */
    function filterItemsByActiveArtboard(pageItems) {
        var artboardRect = getActiveArtboardRect(app.activeDocument);
        var itemsOnArtboard = [];
        for (var i = 0; i < pageItems.length; i++) {
            if (isItemWithinArtboardRect(pageItems[i], artboardRect)) {
                itemsOnArtboard.push(pageItems[i]);
            }
        }
        return itemsOnArtboard;
    }

    /**
     * 判定関数に一致するテキストフレームを選択する
     * @param {Function} predicate - テキストフレームを受け取って true/false を返す関数
     * @param {string} artboardScope - "all" / "current"
     * @returns {number} 選択した件数
     */
    function selectTextFrames(predicate, artboardScope) {
        var doc = app.activeDocument;
        var allTextFrames = doc.textFrames;
        var matchedFrames = [];

        for (var i = 0; i < allTextFrames.length; i++) {
            if (predicate(allTextFrames[i])) {
                matchedFrames.push(allTextFrames[i]);
            }
        }

        if (artboardScope === "current") {
            matchedFrames = filterItemsByActiveArtboard(matchedFrames);
        }

        doc.selection = matchedFrames;
        return matchedFrames.length;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // 後処理 / Post-processing
    // =========================================

    // ボタン行（再利用パーツ） / Button row (reusable)

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

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /**
     * 一括編集の文字列を入力するダイアログを表示する
     * @returns {string|null} 入力した文字列（キャンセル時は null）
     */
    function showBulkEditDialog() {
        var bulkDialog = new Window("dialog", getLabel("dialog.bulkEditTitle"));
        bulkDialog.orientation = "column";
        bulkDialog.alignChildren = "fill";
        bulkDialog.spacing = 10;
        bulkDialog.margins = 16;

        var replacementInput = bulkDialog.add("edittext", undefined, "", {
            multiline: true,
            scrolling: true
        });
        replacementInput.preferredSize = BULK_INPUT_SIZE;
        replacementInput.active = true;

        var bulkButtonRow = addButtonRow(bulkDialog);
        var btnBulkCancel = bulkButtonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnBulkOK = bulkButtonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnBulkOK.onClick = function () {
            bulkDialog.close(1);
        };
        btnBulkCancel.onClick = function () {
            bulkDialog.close(0);
        };

        prepareDialogWindow(bulkDialog, SCRIPT_NAME + "_bulkEdit");
        if (bulkDialog.show() !== 1) {
            return null;
        }
        return replacementInput.text;
    }

    /**
     * テキストフレームの内容をまとめて置き換える（文字ごとの書式は元の文字数ぶん引き継ぐ）
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
            var textFrame = textFrames[i];
            var originalRange = textFrame.textRange;
            var originalLength = originalRange.length;
            if (originalLength === 0) {
                continue;
            }

            textFrame.contents = replacementText;

            var newRange = textFrame.textRange;
            var sharedLength = Math.min(originalLength, newRange.length);
            for (var j = 0; j < sharedLength; j++) {
                /* 書式を写せない文字は飛ばす / skip characters whose formatting cannot be copied */
                try {
                    newRange.characters[j].characterAttributes = originalRange.characters[j].characterAttributes;
                } catch (attrCopyError) {
                }
            }
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
            /* selected を読めないオブジェクトは飛ばす / skip items whose selected state cannot be read */
            try {
                if (childItem.selected) {
                    return true;
                }
            } catch (selectedReadError) {
            }
            if (containsSelectedDescendant(childItem)) {
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
     * 選択中のオブジェクトを「_text」レイヤーへ移動する（レイヤーのロック・表示状態は元に戻す）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function moveItemsToTextLayer(doc, selectedItems) {
        var layerInfo = getOrCreateLayerByName(doc, TEXT_LAYER_NAME);
        var textLayer = layerInfo.layer;
        for (var j = 0; j < selectedItems.length; j++) {
            /* ロック状態や親レイヤーの状態によって移動できない場合はスキップ / Skip items that cannot be moved because of lock state or parent layer state */
            try {
                selectedItems[j].locked = false;
                selectedItems[j].move(textLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (moveItemError) {
            }
        }
        /* 元のロック・可視状態へ戻す（新規作成時はデフォルト値の戻りで実質 no-op） / Restore original lock/visibility (no-op when newly created) */
        try {
            textLayer.locked = layerInfo.originalLocked;
            textLayer.visible = layerInfo.originalVisible;
        } catch (restoreLayerError) {
        }
    }

    /**
     * 選択後の処理を行う
     * @param {string} postProcessMode - "" / "hide" / "hideOthers" / "moveToTextLayer" / "bulkEdit"
     * @returns {void}
     */
    function applyPostProcessToSelection(postProcessMode) {
        if (!postProcessMode) {
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = collectSelectionAsArray(doc.selection);
        if (selectedItems.length === 0) {
            return;
        }

        if (postProcessMode === "hide") {
            hideSelectedItems(selectedItems);
        } else if (postProcessMode === "hideOthers") {
            hideItemsExceptSelection(doc);
        } else if (postProcessMode === "bulkEdit") {
            var textFramesOnly = [];
            for (var k = 0; k < selectedItems.length; k++) {
                if (selectedItems[k].typename === "TextFrame") {
                    textFramesOnly.push(selectedItems[k]);
                }
            }
            bulkEditTextFrames(textFramesOnly);
        } else if (postProcessMode === "moveToTextLayer") {
            moveItemsToTextLayer(doc, selectedItems);
        }
    }

    /**
     * 選択件数を確かめてから後処理を行う（0件ならメッセージ）
     * @param {number} selectedCount - 選択した件数
     * @param {string} postProcessMode - 後処理のモード
     * @returns {void}
     */
    function finalizeSelection(selectedCount, postProcessMode) {
        if (selectedCount === 0) {
            alert(getLabel("alert.noMatch"));
            return;
        }
        applyPostProcessToSelection(postProcessMode);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネルに共通のレイアウトを設定する
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔
     * @returns {void}
     */
    function setupPanelLayout(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
    }

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
     * ラジオボタンの行を追加する。previewText を渡すと右に値のプレビューを添える
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelText - ラジオボタンのラベル
     * @param {string|null} previewText - プレビューの文字列（null なら添えない）
     * @param {number} [labelWidth] - ラジオボタンの幅
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadioRow(parentPanel, labelText, previewText, labelWidth) {
        var rowGroup = parentPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.alignment = "fill";

        var radioButton = rowGroup.add("radiobutton", undefined, labelText);
        if (typeof labelWidth === "number") {
            radioButton.preferredSize.width = labelWidth;
        }
        if (previewText !== null) {
            var previewLabel = rowGroup.add("statictext", undefined, previewText);
            previewLabel.alignment = ["fill", "center"];
            previewLabel.characters = PREVIEW_CHARACTERS;
        }
        return radioButton;
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
        return columnGroup;
    }

    /**
     * ラジオボタンを、親をまたいで排他にする
     * @param {RadioButton[]} radioButtons - 排他にするラジオボタン
     * @returns {void}
     */
    function setupExclusiveRadioButtons(radioButtons) {
        for (var i = 0; i < radioButtons.length; i++) {
            (function (currentRadioButton) {
                currentRadioButton.onClick = function () {
                    for (var j = 0; j < radioButtons.length; j++) {
                        if (radioButtons[j] !== currentRadioButton) {
                            radioButtons[j].value = false;
                        }
                    }
                };
            })(radioButtons[i]);
        }
    }

    /**
     * 指定したラジオボタンだけを選択する
     * @param {RadioButton[]} radioButtons - 同じ組のラジオボタン
     * @param {RadioButton} targetRadioButton - 選択するラジオボタン
     * @returns {void}
     */
    function selectExclusiveRadioButton(radioButtons, targetRadioButton) {
        for (var i = 0; i < radioButtons.length; i++) {
            radioButtons[i].value = (radioButtons[i] === targetRadioButton);
        }
    }

    /**
     * Option＋キーで選択条件を切り替えるハンドラーを付ける（検索文字列の入力中は無効）
     * @param {Object} selectorControls - buildDialog() の結果
     * @returns {void}
     */
    function addSelectionKeyHandler(selectorControls) {
        var shortcutRadios = {
            Q: selectorControls.rbAllText,
            W: selectorControls.rbPointText,
            E: selectorControls.rbAreaText,
            A: selectorControls.rbExactMatch,
            B: selectorControls.rbStartsWith,
            D: selectorControls.rbEndsWith,
            I: selectorControls.rbContainsMatch,
            R: selectorControls.rbRegexMatch
        };

        /**
         * ラジオを、親をまたいだ組の中で1つだけ選ぶショートカットを作る
         * @param {RadioButton} targetRadioButton - 選ぶラジオボタン
         * @returns {Function} addKeyShortcuts に渡す関数
         */
        function makeSelectShortcut(targetRadioButton) {
            return function () {
                selectExclusiveRadioButton(selectorControls.selectionRadios, targetRadioButton);
            };
        }

        /* 検索文字列の入力中は効かない（文字の欄なので numericFields に入れない）/ Off while typing the keyword */
        var shortcutMap = {};
        for (var keyLetter in shortcutRadios) {
            if (!shortcutRadios.hasOwnProperty(keyLetter)) continue;
            shortcutMap["Alt+" + keyLetter] = makeSelectShortcut(shortcutRadios[keyLetter]);
        }
        addKeyShortcuts(selectorControls.selectorDialog, shortcutMap);
    }

    /**
     * 選択条件パネル（対象アートボード・属性・テキストの種類・文字列）を追加する
     * @param {Object} selectorControls - コントロールの格納先
     * @param {Object} initialState - 開いた時点の選択の情報（hasSelection / keyword / attributePreview）
     * @returns {void}
     */
    function addSelectionPanel(selectorControls, initialState) {
        var selectionPanel = selectorControls.selectorDialog.add("panel", undefined, getLabel("panel.selection"));
        setupPanelLayout(selectionPanel, 10);
        selectionPanel.margins = OUTER_PANEL_MARGINS;

        /* 対象アートボードパネル / Target artboard panel */
        var artboardScopePanel = selectionPanel.add("panel", undefined, getLabel("panel.artboardScope"));
        artboardScopePanel.orientation = "row";
        artboardScopePanel.alignChildren = ["left", "center"];
        artboardScopePanel.alignment = "fill";
        artboardScopePanel.margins = OUTER_PANEL_MARGINS;
        artboardScopePanel.spacing = 20;

        selectorControls.rbArtboardAll = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.artboardAll"));
        selectorControls.rbArtboardCurrent = artboardScopePanel.add("radiobutton", undefined, getLabel("radio.artboardCurrent"));
        selectorControls.rbArtboardAll.value = true;
        setHelpTip(selectorControls.rbArtboardAll, getLabel("tooltip.artboardAll"));
        setHelpTip(selectorControls.rbArtboardCurrent, getLabel("tooltip.artboardCurrent"));
        setupExclusiveRadioButtons([selectorControls.rbArtboardAll, selectorControls.rbArtboardCurrent]);

        /* 属性選択パネル / Attribute selection panel */
        var attributePanel = selectionPanel.add("panel", undefined, getLabel("panel.attribute"));
        setupPanelLayout(attributePanel, 6);
        attributePanel.enabled = initialState.hasSelection;

        var attributePreview = initialState.attributePreview;
        selectorControls.rbFontFamily = addRadioRow(attributePanel, getLabel("radio.fontFamily"), attributePreview.family, ATTRIBUTE_LABEL_WIDTH);
        selectorControls.rbFontFamilyStyle = addRadioRow(attributePanel, getLabel("radio.fontFamilyStyle"), attributePreview.familyStyle, ATTRIBUTE_LABEL_WIDTH);
        selectorControls.rbFontFamilyStyleSize = addRadioRow(attributePanel, getLabel("radio.fontFamilyStyleSize"), attributePreview.familyStyleSize, ATTRIBUTE_LABEL_WIDTH);
        selectorControls.rbFontSize = addRadioRow(attributePanel, getLabel("radio.fontSize"), attributePreview.size, ATTRIBUTE_LABEL_WIDTH);
        selectorControls.rbTextFillColor = addRadioRow(attributePanel, getLabel("radio.textFillColor"), null, ATTRIBUTE_LABEL_WIDTH);
        selectorControls.rbOpacity = addRadioRow(attributePanel, getLabel("radio.opacity"), attributePreview.opacity, ATTRIBUTE_LABEL_WIDTH);

        setHelpTip(selectorControls.rbFontFamily, getLabel("tooltip.fontFamily"));
        setHelpTip(selectorControls.rbFontFamilyStyle, getLabel("tooltip.fontFamilyStyle"));
        setHelpTip(selectorControls.rbFontFamilyStyleSize, getLabel("tooltip.fontFamilyStyleSize"));
        setHelpTip(selectorControls.rbFontSize, getLabel("tooltip.fontSize"));
        setHelpTip(selectorControls.rbTextFillColor, getLabel("tooltip.textFillColor"));
        setHelpTip(selectorControls.rbOpacity, getLabel("tooltip.opacity"));

        /* テキスト種類と文字列条件を横並びに配置 / Arrange text type and string condition panels side by side */
        var textConditionGroup = selectionPanel.add("group");
        textConditionGroup.orientation = "row";
        textConditionGroup.alignChildren = "fill";
        textConditionGroup.alignment = "fill";

        /* テキスト種類パネル / Text type panel */
        var textTypePanel = textConditionGroup.add("panel", undefined, getLabel("panel.textType"));
        setupPanelLayout(textTypePanel, 6);

        selectorControls.rbAllText = textTypePanel.add("radiobutton", undefined, getLabel("radio.allText"));
        selectorControls.rbPointText = textTypePanel.add("radiobutton", undefined, getLabel("radio.pointText"));
        selectorControls.rbAreaText = textTypePanel.add("radiobutton", undefined, getLabel("radio.areaText"));
        selectorControls.rbPathText = textTypePanel.add("radiobutton", undefined, getLabel("radio.pathText"));

        setHelpTip(selectorControls.rbAllText, getLabel("tooltip.allText"));
        setHelpTip(selectorControls.rbPointText, getLabel("tooltip.pointText"));
        setHelpTip(selectorControls.rbAreaText, getLabel("tooltip.areaText"));
        setHelpTip(selectorControls.rbPathText, getLabel("tooltip.pathText"));

        /* 初期選択を設定（選択があれば「＋スタイルとサイズ」、なければ「すべて」） / Set initial selection */
        if (initialState.hasSelection) {
            selectorControls.rbFontFamilyStyleSize.value = true;
        } else {
            selectorControls.rbAllText.value = true;
        }

        /* 文字列条件パネル / String condition panel */
        var textMatchPanel = textConditionGroup.add("panel", undefined, getLabel("panel.textMatch"));
        setupPanelLayout(textMatchPanel, 6);

        selectorControls.keywordInput = textMatchPanel.add("edittext", undefined, initialState.keyword);
        selectorControls.keywordInput.characters = KEYWORD_CHARACTERS;
        setHelpTip(selectorControls.keywordInput, getLabel("tooltip.keywordInput"));

        var textMatchOptionsGroup = textMatchPanel.add("group");
        textMatchOptionsGroup.orientation = "row";
        textMatchOptionsGroup.alignChildren = "top";
        textMatchOptionsGroup.margins = [0, 10, 0, 0];
        textMatchOptionsGroup.spacing = 22;

        var textMatchLeftColumnGroup = addColumnGroup(textMatchOptionsGroup);
        var textMatchCenterColumnGroup = addColumnGroup(textMatchOptionsGroup);
        var textMatchRightColumnGroup = addColumnGroup(textMatchOptionsGroup);

        selectorControls.rbExactMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, getLabel("radio.exactMatch"));
        selectorControls.rbContainsMatch = textMatchLeftColumnGroup.add("radiobutton", undefined, getLabel("radio.containsMatch"));
        selectorControls.rbStartsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, getLabel("radio.startsWith"));
        selectorControls.rbEndsWith = textMatchCenterColumnGroup.add("radiobutton", undefined, getLabel("radio.endsWith"));
        selectorControls.rbRegexMatch = textMatchRightColumnGroup.add("radiobutton", undefined, getLabel("radio.regexMatch"));

        setHelpTip(selectorControls.rbExactMatch, getLabel("tooltip.exactMatch"));
        setHelpTip(selectorControls.rbStartsWith, getLabel("tooltip.startsWith"));
        setHelpTip(selectorControls.rbEndsWith, getLabel("tooltip.endsWith"));
        setHelpTip(selectorControls.rbContainsMatch, getLabel("tooltip.containsMatch"));
        setHelpTip(selectorControls.rbRegexMatch, getLabel("tooltip.regexMatch"));

        /* 選択条件のラジオは、パネルをまたいで排他にする / The criteria radios are exclusive across panels */
        selectorControls.selectionRadios = [
            selectorControls.rbFontFamily, selectorControls.rbFontFamilyStyle, selectorControls.rbFontFamilyStyleSize, selectorControls.rbFontSize, selectorControls.rbTextFillColor, selectorControls.rbOpacity,
            selectorControls.rbAllText, selectorControls.rbPointText, selectorControls.rbAreaText, selectorControls.rbPathText,
            selectorControls.rbExactMatch, selectorControls.rbStartsWith, selectorControls.rbEndsWith, selectorControls.rbContainsMatch, selectorControls.rbRegexMatch
        ];
        setupExclusiveRadioButtons(selectorControls.selectionRadios);
    }

    /**
     * 選択後の処理パネルを追加する
     * @param {Object} selectorControls - コントロールの格納先
     * @returns {void}
     */
    function addPostProcessPanel(selectorControls) {
        var postProcessPanel = selectorControls.selectorDialog.add("panel", undefined, getLabel("panel.postProcess"));
        setupPanelLayout(postProcessPanel, 6);
        /* ドキュメント内に TextFrame が 0 件なら後処理は無意味なのでディム / Disable post-process when document has no text frames */
        postProcessPanel.enabled = app.activeDocument.textFrames.length > 0;

        selectorControls.rbNoPostProcess = postProcessPanel.add("radiobutton", undefined, getLabel("radio.noPostProcess"));
        selectorControls.rbHide = postProcessPanel.add("radiobutton", undefined, getLabel("radio.hideAfterSelection"));
        selectorControls.rbHideOthers = postProcessPanel.add("radiobutton", undefined, getLabel("radio.hideOthers"));
        selectorControls.rbMove = postProcessPanel.add("radiobutton", undefined, getLabel("radio.moveToTextLayer"));
        selectorControls.rbBulkEdit = postProcessPanel.add("radiobutton", undefined, getLabel("radio.bulkEdit"));

        selectorControls.rbNoPostProcess.value = true;

        setHelpTip(selectorControls.rbNoPostProcess, getLabel("tooltip.noPostProcess"));
        setHelpTip(selectorControls.rbHide, getLabel("tooltip.hideAfterSelection"));
        setHelpTip(selectorControls.rbHideOthers, getLabel("tooltip.hideOthers"));
        setHelpTip(selectorControls.rbMove, getLabel("tooltip.moveToTextLayer"));
        setHelpTip(selectorControls.rbBulkEdit, getLabel("tooltip.bulkEdit"));

        setupExclusiveRadioButtons([selectorControls.rbNoPostProcess, selectorControls.rbHide, selectorControls.rbHideOthers, selectorControls.rbMove, selectorControls.rbBulkEdit]);
    }

    /**
     * ダイアログを組み立てる（OK/キャンセルの処理は main() で付ける）
     * @param {Object} initialState - 開いた時点の選択の情報（hasSelection / keyword / attributePreview）
     * @returns {Object} ダイアログ本体（selectorDialog）と各コントロール
     */
    function buildDialog(initialState) {
        var selectorControls = {};
        selectorControls.selectorDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);

        addSelectionPanel(selectorControls, initialState);
        addSelectionKeyHandler(selectorControls);
        addPostProcessPanel(selectorControls);

        /* ボタンエリア / Button area */
        var buttonRow = addButtonRow(selectorControls.selectorDialog);
        selectorControls.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        selectorControls.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return selectorControls;
    }

    /* 属性ラジオと、選択に使う Illustrator 標準コマンド / Attribute radios and the built-in commands they run */
    var ATTRIBUTE_COMMANDS = [
        { radioKey: "rbFontFamily", command: "Find Text Font Family menu item" },
        { radioKey: "rbFontFamilyStyle", command: "Find Text Font Family Style menu item" },
        { radioKey: "rbFontFamilyStyleSize", command: "Find Text Font Family Style Size menu item" },
        { radioKey: "rbFontSize", command: "Find Text Font Size menu item" },
        { radioKey: "rbTextFillColor", command: "Find Text Fill Color menu item" },
        { radioKey: "rbOpacity", command: "Find Opacity menu item" }
    ];

    /* 文字列の一致モードのラジオ（判定順） / String match radios, in checking order */
    var TEXT_MATCH_MODES = [
        { radioKey: "rbExactMatch", mode: "exact" },
        { radioKey: "rbStartsWith", mode: "startsWith" },
        { radioKey: "rbEndsWith", mode: "endsWith" },
        { radioKey: "rbContainsMatch", mode: "contains" },
        { radioKey: "rbRegexMatch", mode: "regex" }
    ];

    /* テキストの種類のラジオ（判定順） / Text type radios, in checking order */
    var TEXT_TYPES = [
        { radioKey: "rbAllText", textType: "all" },
        { radioKey: "rbPointText", textType: "point" },
        { radioKey: "rbAreaText", textType: "area" },
        { radioKey: "rbPathText", textType: "path" }
    ];

    /* 後処理のラジオ（判定順） / Post-process radios, in checking order */
    var POST_PROCESS_MODES = [
        { radioKey: "rbNoPostProcess", mode: "" },
        { radioKey: "rbHide", mode: "hide" },
        { radioKey: "rbHideOthers", mode: "hideOthers" },
        { radioKey: "rbMove", mode: "moveToTextLayer" },
        { radioKey: "rbBulkEdit", mode: "bulkEdit" }
    ];

    /**
     * ラジオの対応表から、オンになっている最初の行の値を返す
     * @param {Object} selectorControls - buildDialog() の結果
     * @param {Object[]} radioTable - radioKey と値を持つ行の配列
     * @param {string} valueKey - 返す値のキー
     * @param {string} fallbackValue - どれもオフのときの値
     * @returns {string} 値
     */
    function readRadioChoice(selectorControls, radioTable, valueKey, fallbackValue) {
        for (var i = 0; i < radioTable.length; i++) {
            if (selectorControls[radioTable[i].radioKey].value) return radioTable[i][valueKey];
        }
        return fallbackValue;
    }

    /**
     * ダイアログの状態から設定を読み取る
     * @param {Object} selectorControls - buildDialog() の結果
     * @returns {Object} postProcessMode / artboardScope / attributeCommand / textMatchMode / keyword / textType
     */
    function readDialogSettings(selectorControls) {
        return {
            postProcessMode: readRadioChoice(selectorControls, POST_PROCESS_MODES, "mode", ""),
            artboardScope: selectorControls.rbArtboardCurrent.value ? "current" : "all",
            attributeCommand: readRadioChoice(selectorControls, ATTRIBUTE_COMMANDS, "command", ""),
            textMatchMode: readRadioChoice(selectorControls, TEXT_MATCH_MODES, "mode", ""),
            keyword: selectorControls.keywordInput.text || "",
            textType: readRadioChoice(selectorControls, TEXT_TYPES, "textType", "all")
        };
    }

    /**
     * 属性の標準コマンドで選択し、必要なら現在のアートボードに絞って後処理を行う
     * @param {Object} selectorSettings - readDialogSettings() の結果
     * @returns {void}
     */
    function selectByAttribute(selectorSettings) {
        app.executeMenuCommand(selectorSettings.attributeCommand);
        if (selectorSettings.artboardScope === "current") {
            var filteredAttributeSelection = filterItemsByActiveArtboard(collectSelectionAsArray(app.activeDocument.selection));
            app.activeDocument.selection = filteredAttributeSelection;
            finalizeSelection(filteredAttributeSelection.length, selectorSettings.postProcessMode);
            return;
        }
        applyPostProcessToSelection(selectorSettings.postProcessMode);
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
        var selectorControls = buildDialog({
            hasSelection: initialSelection.length > 0,
            keyword: getSelectedTextString(initialSelection),
            attributePreview: getSelectedTextAttributePreview(initialSelection)
        });

        /* OKボタン実行処理 / Handle OK button action */
        selectorControls.btnOK.onClick = function () {
            var selectorSettings = readDialogSettings(selectorControls);

            /* 属性で選択：Illustrator標準コマンドに選択を任せる / By attribute: let Illustrator's command perform the selection */
            if (selectorSettings.attributeCommand) {
                selectorControls.selectorDialog.close();
                selectByAttribute(selectorSettings);
                return;
            }

            /* 文字列で選択 / By string */
            if (selectorSettings.textMatchMode) {
                var keywordValidation = validateTextMatchInput(selectorSettings.keyword, selectorSettings.textMatchMode);
                if (!keywordValidation) {
                    return;
                }
                selectorControls.selectorDialog.close();
                var stringPredicate = function (textFrame) {
                    return textMatches(textFrame.contents || "", selectorSettings.keyword, selectorSettings.textMatchMode, keywordValidation.regex);
                };
                finalizeSelection(selectTextFrames(stringPredicate, selectorSettings.artboardScope), selectorSettings.postProcessMode);
                return;
            }

            /* テキストの種類で選択 / By text type */
            selectorControls.selectorDialog.close();
            var typePredicate = buildTextTypePredicate(selectorSettings.textType);
            finalizeSelection(selectTextFrames(typePredicate, selectorSettings.artboardScope), selectorSettings.postProcessMode);
        };

        /* キャンセルボタン処理 / Handle Cancel button action */
        selectorControls.btnCancel.onClick = function () {
            selectorControls.selectorDialog.close();
        };

        prepareDialogWindow(selectorControls.selectorDialog, SCRIPT_NAME);
        selectorControls.selectorDialog.show();
    }

    main();

})();
