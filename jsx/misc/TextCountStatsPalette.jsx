#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のテキストの文字数などを集計して表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextCountStatsPalette.md

### Overview

Tallies the character counts and related statistics of the text in the document.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextCountStatsPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextCountStatsPalette";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextCountStatsPalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextCountStatsPalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    var LABEL_WIDTH = 110;                  /* 項目名の幅 / Row label width */
    var VALUE_WIDTH = 100;                  /* 値の幅 / Value width */
    var PALETTE_OPACITY = 0.97;             /* パレットの不透明度 / Palette opacity */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "文字数カウント", en: "Text Count Stats" }
        },
        panel: {
            charPara: { ja: "文字・段落", en: "Characters & Paragraphs" },
            check: { ja: "チェック項目", en: "Check Items" },
            kinds: { ja: "種別", en: "Type" },
            other: { ja: "その他", en: "Other" }
        },
        fieldLabel: {
            chars: { ja: "文字", en: "Characters" },
            paras: { ja: "段落", en: "Paragraphs" },
            lines: { ja: "行", en: "Lines" },
            words: { ja: "英単語", en: "English Words" },
            fullwidth: { ja: "全角文字", en: "Fullwidth Chars" },
            hankakuKana: { ja: "半角カナ", en: "Half-width Kana" },
            pointText: { ja: "ポイント文字", en: "Point Type" },
            areaText: { ja: "エリア内文字", en: "Area Type" },
            pathText: { ja: "パス上文字", en: "Type on a Path" },
            fonts: { ja: "使用フォント", en: "Fonts Used" }
        },
        button: {
            refresh: { ja: "更新", en: "Refresh" }
        },
        status: {
            ready: { ja: "準備完了", en: "Ready" },
            noDoc: { ja: "ドキュメントが開かれていません", en: "No document open" },
            wholeDoc: { ja: "選択なし（全体を集計）", en: "No selection (counting all)" },
            selectedPrefix: { ja: "選択 ", en: "Selected " },
            selectedSuffix: { ja: " 件を集計", en: " object(s)" },
            timeout: { ja: "Illustrator から応答がありません", en: "No response from Illustrator" },
            busy: { ja: "処理中です", en: "Busy" },
            error: { ja: "エラー", en: "Error" }
        },
        tooltip: {
            refresh: { ja: "選択内容を再集計（⌘R / Enter）", en: "Recount selection (Cmd+R / Enter)" },
            esc: { ja: "Esc で閉じる", en: "Press Esc to close" },
            statValue: {
                ja: "左が選択範囲、右がドキュメント全体の値です",
                en: "Shows the selection on the left and the whole document on the right"
            },
            paras: { ja: "空の段落や空白だけの段落は数えません", en: "Empty paragraphs and paragraphs with only spaces are not counted" },
            lines: { ja: "自動で折り返された行も1行として数えます", en: "Lines created by automatic wrapping are counted too" },
            words: { ja: "半角英字だけの並びを1語として数えます", en: "Counts each run of ASCII letters as one word" },
            fullwidth: {
                ja: "全角の英数字・記号（！〜｠、￠〜￦）を数えます。漢字・かな・句読点は含みません",
                en: "Counts fullwidth letters, digits and symbols (U+FF01-FF60, U+FFE0-FFE6); kanji, kana and Japanese punctuation are not included"
            },
            fonts: {
                ja: "選択に関係なく、ドキュメント全体で使われているフォントの数を表示します",
                en: "Shows the number of fonts used in the whole document, regardless of the selection"
            }
        }
    };

    /* ============================================================
       worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
       ------------------------------------------------------------
       注意 / Notes:
       - toString() で送るので JSDoc を付けない。// 行コメント禁止・/* *\/ のみ・
         各文は必ずセミコロンで終える
       - Sent via toString(): no JSDoc, no // comments; use block comments and
         always terminate statements with a semicolon
       ============================================================ */
    function wkCountTextStats() {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection) { currentSelection = []; }
        var fontSet = {};

        /* テキストフレームを種別・文字数などで集計し、フォント名を fontSet に集める / Tally text frames and collect font names */
        function tallyTextFrames(items) {
            var tally = { chars: 0, paras: 0, lines: 0, words: 0, fullwidth: 0, kana: 0, point: 0, area: 0, path: 0 };
            for (var i = 0; i < items.length; i++) {
                var frame = items[i];
                if (frame.typename !== "TextFrame") { continue; }
                if (frame.kind === TextType.POINTTEXT) { tally.point++; }
                else if (frame.kind === TextType.AREATEXT) { tally.area++; }
                else if (frame.kind === TextType.PATHTEXT) { tally.path++; }
                try {
                    var frameRange = frame.textRange || frame.textRanges[0];
                    fontSet[frameRange.characterAttributes.textFont.name] = true;
                } catch (e) {}
                try { tally.chars += frame.characters.length; } catch (e2) {}
                /* 末尾の空段落は読むと例外になることがある / Reading a trailing empty paragraph can throw */
                try {
                    var paragraphs = frame.paragraphs;
                    var filledParaCount = 0;
                    for (var p = 0; p < paragraphs.length; p++) {
                        if (paragraphs[p].contents.replace(/[\s　]/g, "").length > 0) { filledParaCount++; }
                    }
                    tally.paras += filledParaCount;
                } catch (e3) {}
                try { tally.lines += frame.lines.length; } catch (e4) {}
                try {
                    var frameContents = frame.contents;
                    if (typeof frameContents === "string") {
                        var wordMatches = frameContents.match(/\b[a-zA-Z]+\b/g); if (wordMatches) { tally.words += wordMatches.length; }
                        var fullwidthMatches = frameContents.match(/[！-｠￠-￦]/g); if (fullwidthMatches) { tally.fullwidth += fullwidthMatches.length; }
                        var kanaMatches = frameContents.match(/[･-ﾟ]/g); if (kanaMatches) { tally.kana += kanaMatches.length; }
                    }
                } catch (e5) {}
            }
            return tally;
        }

        var selectionTally = tallyTextFrames(currentSelection);
        var allTally = tallyTextFrames(doc.pageItems);

        var fontCount = 0;
        for (var fontName in fontSet) { if (fontSet.hasOwnProperty(fontName)) { fontCount++; } }

        var resultPairs = [
            "selCount=" + currentSelection.length,
            "charSel=" + selectionTally.chars,
            "charAll=" + allTally.chars,
            "paraSel=" + selectionTally.paras,
            "paraAll=" + allTally.paras,
            "lineSel=" + selectionTally.lines,
            "lineAll=" + allTally.lines,
            "wordSel=" + selectionTally.words,
            "wordAll=" + allTally.words,
            "fwSel=" + selectionTally.fullwidth,
            "fwAll=" + allTally.fullwidth,
            "kanaSel=" + selectionTally.kana,
            "kanaAll=" + allTally.kana,
            "pointSel=" + selectionTally.point,
            "pointAll=" + allTally.point,
            "areaSel=" + selectionTally.area,
            "areaAll=" + allTally.area,
            "pathSel=" + selectionTally.path,
            "pathAll=" + allTally.path,
            "fontCount=" + fontCount
        ];
        return "OK|" + resultPairs.join("|");
    }

    /* worker 関数は全登録（追加漏れ防止） / Register every worker function */
    var WORKER_FUNCS = [wkCountTextStats];

    // =========================================
    // BridgeTalk 委譲 / Delegation to the main engine
    // =========================================
    var isBridgeBusy = false;

    /**
     * worker 関数を連結してメインエンジンで評価し、結果の文字列を返す
     * @param {string} callExpression - 評価させる呼び出し式（例: "wkCountTextStats()"）
     * @returns {string} worker の戻り値、または "ERR:BUSY" / "ERR:TIMEOUT" / "ERR:…"
     */
    function callMainEngine(callExpression) {
        /* 再入防止 / Re-entrancy guard */
        if (isBridgeBusy) { return "ERR:BUSY"; }
        isBridgeBusy = true;

        var resultHolder = { value: null };
        /* BridgeTalk の生成・送信 / Creating and sending BridgeTalk */
        try {
            /* worker 群を連結し、末尾に呼び出し式を付与 / Concatenate workers + call */
            var workerSource = "";
            for (var i = 0; i < WORKER_FUNCS.length; i++) { workerSource += WORKER_FUNCS[i].toString(); }
            workerSource += callExpression + ";";

            var bridgeTalk = new BridgeTalk();
            bridgeTalk.target = "illustrator";
            /* encodeURIComponent で多バイト・改行・特殊文字の破損を回避 / Avoid corrupting multibyte text and newlines */
            bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerSource) + "\"));";
            bridgeTalk.onResult = function (response) {
                resultHolder.value = (response && response.body != null) ? String(response.body) : "";
            };
            bridgeTalk.onError = function (errorResponse) {
                resultHolder.value = "ERR:" + ((errorResponse && errorResponse.body) ? errorResponse.body : "bridge");
            };
            bridgeTalk.send(10);
        } catch (e) {
            resultHolder.value = "ERR:" + e;
        } finally {
            isBridgeBusy = false;
        }

        if (resultHolder.value === null) { return "ERR:TIMEOUT"; }
        return resultHolder.value;
    }

    /**
     * worker の戻り値（OK|key=value|...）を解析する
     * @param {string} response - worker の戻り値
     * @returns {Object|null} キーと値の対応表（形式が違うときは null）
     */
    function parseStats(response) {
        if (!response || response.indexOf("OK|") !== 0) return null;
        var statsMap = {};
        var pairs = response.substring(3).split("|");
        for (var i = 0; i < pairs.length; i++) {
            var keyValue = pairs[i].split("=");
            if (keyValue.length === 2) { statsMap[keyValue[0]] = keyValue[1]; }
        }
        return statsMap;
    }

    // =========================================
    // パレット構築 / Build palette
    // =========================================

    /**
     * 集計項目をまとめるパネルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} titlePath - パネルタイトルの LABELS パス
     * @returns {Panel} 追加したパネル
     */
    function addStatPanel(parentGroup, titlePath) {
        var statPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        setupPanel(statPanel);
        return statPanel;
    }

    /**
     * 項目名と値の行を追加し、値の statictext を返す
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.fieldLabel のキー（LABELS.tooltip に同じキーがあれば項目名の tooltip にする）
     * @returns {StaticText} 値を表示する statictext
     */
    function addStatRow(parentPanel, labelKey) {
        var statRow = parentPanel.add("group");
        statRow.orientation = "row";
        statRow.alignChildren = ["left", "center"];

        var rowLabel = statRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        if (LABELS.tooltip[labelKey]) rowLabel.helpTip = getLabel("tooltip." + labelKey);

        var valueText = statRow.add("statictext", undefined, "-");
        valueText.preferredSize.width = VALUE_WIDTH;
        valueText.justify = "left";
        return valueText;
    }

    /**
     * 集計結果を値の欄に書き込む
     * @param {Object} valueFields - 項目キーと値の statictext の対応表
     * @param {Object} stats - parseStats() の結果
     * @returns {void}
     */
    function showStats(valueFields, stats) {
        valueFields.chars.text = stats.charSel + " / " + stats.charAll;
        valueFields.paras.text = stats.paraSel + " / " + stats.paraAll;
        valueFields.lines.text = stats.lineSel + " / " + stats.lineAll;
        valueFields.words.text = stats.wordSel + " / " + stats.wordAll;
        valueFields.fullwidth.text = stats.fwSel + " / " + stats.fwAll;
        valueFields.hankakuKana.text = stats.kanaSel + " / " + stats.kanaAll;
        valueFields.pointText.text = stats.pointSel + " / " + stats.pointAll;
        valueFields.areaText.text = stats.areaSel + " / " + stats.areaAll;
        valueFields.pathText.text = stats.pathSel + " / " + stats.pathAll;
        valueFields.fonts.text = stats.fontCount;
    }

    /**
     * 集計パレットを組み立てる
     * @returns {Window} 組み立てたパレット
     */
    function buildPalette() {
        var statsPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        setupWindow(statsPalette);

        var panelColumnGroup = statsPalette.add("group");
        panelColumnGroup.orientation = "column";
        panelColumnGroup.alignChildren = ["fill", "top"];

        var charParaPanel = addStatPanel(panelColumnGroup, "panel.charPara");
        var checkPanel = addStatPanel(panelColumnGroup, "panel.check");
        var kindsPanel = addStatPanel(panelColumnGroup, "panel.kinds");
        var otherPanel = addStatPanel(panelColumnGroup, "panel.other");

        /* 値の statictext 参照を保持 / Keep references to value fields */
        var valueFields = {
            chars: addStatRow(charParaPanel, "chars"),
            paras: addStatRow(charParaPanel, "paras"),
            lines: addStatRow(charParaPanel, "lines"),
            words: addStatRow(charParaPanel, "words"),
            fullwidth: addStatRow(checkPanel, "fullwidth"),
            hankakuKana: addStatRow(checkPanel, "hankakuKana"),
            pointText: addStatRow(kindsPanel, "pointText"),
            areaText: addStatRow(kindsPanel, "areaText"),
            pathText: addStatRow(kindsPanel, "pathText"),
            fonts: addStatRow(otherPanel, "fonts")
        };
        /* 「選択 / 全体」の並びを説明（フォント数は全体のみ） / Explain the "selection / all" pair (fonts show the total only) */
        for (var fieldKey in valueFields) {
            if (fieldKey !== "fonts") valueFields[fieldKey].helpTip = getLabel("tooltip.statValue");
        }

        /* ステータス表示 / Status line */
        var statusText = statsPalette.add("statictext", undefined, getLabel("status.ready"));
        statusText.alignment = ["fill", "bottom"];

        /**
         * ステータス行の表示を差し替える
         * @param {string} message - 表示する文字列
         * @returns {void}
         */
        function setStatus(message) {
            statusText.text = message;
        }

        /**
         * メインエンジンで集計し直して表示を更新する
         * @returns {void}
         */
        function refreshStats() {
            setStatus(getLabel("status.busy"));
            var response = callMainEngine("wkCountTextStats()");

            if (response === "ERR:BUSY") { setStatus(getLabel("status.busy")); return; }
            if (response === null || response === "ERR:TIMEOUT") { setStatus(getLabel("status.timeout")); return; }
            if (response === "NODOC") { setStatus(getLabel("status.noDoc")); return; }
            if (response.indexOf("ERR:") === 0) { setStatus(labelValueText("status.error", response.substring(4))); return; }

            var stats = parseStats(response);
            if (!stats) { setStatus(getLabel("status.error")); return; }

            showStats(valueFields, stats);

            var selectedCount = parseInt(stats.selCount, 10) || 0;
            if (selectedCount > 0) {
                setStatus(getLabel("status.selectedPrefix") + selectedCount + getLabel("status.selectedSuffix"));
            } else {
                setStatus(getLabel("status.wholeDoc"));
            }
        }

        /* ボタン（更新のみ。閉じるは × / Esc に任せる） / Button (refresh only) */
        var btnRowGroup = statsPalette.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.alignChildren = ["right", "center"];

        var btnRefresh = btnRowGroup.add("button", undefined, getLabel("button.refresh"));
        btnRefresh.helpTip = getLabel("tooltip.refresh") + "\n" + getLabel("tooltip.esc");
        /* onClick 連結（addEventListener('click') は不発の環境がある） / addEventListener('click') does not fire in some environments */
        btnRefresh.onClick = refreshStats;

        /* キー操作：Esc で閉じる、⌘R / Enter で更新 / Keys: Esc close, Cmd+R / Enter refresh */
        addKeyShortcuts(statsPalette, {
            "Escape": { target: function () { statsPalette.close(); }, inFields: true },
            "R": refreshStats,
            "Cmd+R": refreshStats,
            "Enter": refreshStats,
            "Return": refreshStats
        });

        /* opacity に対応しない環境がある / Some environments do not support opacity */
        try { statsPalette.opacity = PALETTE_OPACITY; } catch (e) {}

        /* 表示直後に一度集計 / Count once right after showing */
        statsPalette.onShow = function () {
            refreshStats();
        };

        return statsPalette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * パレットを開く（開いているものがあれば閉じてから開き直す）
     * @returns {void}
     */
    function showPalette() {
        /* 多重起動防止：既存パレットがあれば閉じる（前回の実行で無効になっていることがある）
           Prevent duplicates: close the existing palette (it may already be invalid) */
        if ($.global.__TextCountStatsPalette) {
            try { $.global.__TextCountStatsPalette.close(); } catch (e) {}
            $.global.__TextCountStatsPalette = null;
        }

        var statsPalette = buildPalette();

        /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
        $.global.__TextCountStatsPalette = statsPalette;
        statsPalette.onClose = function () {
            $.global.__TextCountStatsPalette = null;
        };

        statsPalette.show();
    }

    showPalette();

})();
