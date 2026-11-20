#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.PathInspectorPalette
Name=パスの統計を常駐パレットで表示
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.9
Release-Date=2026-07-31
Description-BEGIN
選択中またはドキュメント全体のパス統計を集計し、常駐パレットで表示します。

Counts path statistics for the selection, or for the whole document, and shows them in a persistent palette.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択中またはドキュメント全体のパス統計を集計し、常駐パレットで表示します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathInspectorPalette.md

### Overview

Counts path statistics for the selection, or for the whole document, and shows them in a persistent palette.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathInspectorPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathInspectorPalette";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "パスの統計を常駐パレットで表示"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathInspectorPalette.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathInspectorPalette.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/single-function/PathInspectorPalette.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    /* 日英ラベル定義（カテゴリ構造） / Japanese-English labels (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "パスのカウント", en: "Path Count" }
        },
        report: {
            title: { ja: "Path Inspector Report", en: "Path Inspector Report" },
            document: { ja: "ドキュメント", en: "Document" },
            date: { ja: "日付", en: "Date" },
            valueNote: {
                ja: "※ 値は『選択 / 全体』の形式です",
                en: "Note: values are formatted as 'Selection / All'"
            }
        },
        section: {
            paths: { ja: "パス", en: "Paths" }
        },
        panel: {
            path: { ja: "パス", en: "Paths" }
        },
        row: {
            pathCount: { ja: "パス", en: "Paths" },
            openPath: { ja: "オープンパス", en: "Open Paths" },
            closedPath: { ja: "クローズパス", en: "Closed Paths" },
            anchors: { ja: "アンカーポイント", en: "Anchor Points" },
            handles: { ja: "ハンドル", en: "Handles" },
            compoundPath: { ja: "複合パス", en: "Compound Paths" },
            compoundShape: { ja: "複合シェイプ", en: "Compound Shapes" }
        },
        button: {
            refresh: { ja: "更新", en: "Refresh" },
            exportPreset: { ja: "書き出し", en: "Export" }
        },
        status: {
            ready: { ja: "準備完了", en: "Ready" },
            noDoc: { ja: "ドキュメントが開かれていません", en: "No document open" },
            wholeDoc: { ja: "選択なし（全体を集計）", en: "No selection (counting all)" },
            selectedPrefix: { ja: "選択 ", en: "Selected " },
            selectedSuffix: { ja: " 件を集計", en: " object(s)" },
            timeout: { ja: "Illustrator から応答がありません", en: "No response from Illustrator" },
            busy: { ja: "処理中です", en: "Busy" },
            error: { ja: "エラー", en: "Error" },
            exportedPrefix: { ja: "書き出しました: ", en: "Exported: " },
            exportFailOpen: { ja: "ファイルを開けませんでした", en: "Failed to open the file" }
        },
        hint: {
            refresh: { ja: "選択内容を再集計（⌘R）", en: "Recount selection (Cmd+R)" },
            esc: { ja: "Esc で閉じる", en: "Press Esc to close" }
        }
    };

    /**
     * 書き出し用にラベル末尾のコロンを半角へ正規化する
     * @param {string} path ラベルのドットパス
     * @returns {string} 正規化済み文字列
     */
    function LX(path) {
        return getLabel(path) + ":";
    }

    (function () {

        /* ============================================================
           定数 / Constants
           ============================================================ */
        var LABEL_WIDTH = 130;
        var VALUE_WIDTH = 90;
        var PALETTE_OPACITY = 0.97;

        /* ============================================================
           レイアウト / Layout
           ============================================================ */
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
        /* 左下のスクリプト名・バージョンをクリックすると開く。基本情報ブロックの SCRIPT_* を読み、
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
         * 基本情報ブロックの値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
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
            /* 余りの幅を取り、コピーボタンを右端に置く / Take the extra width so the copy button sits at the right */
            valueLabel.alignment = ["fill", "center"];
            if (copyFromPalette !== null) addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
            return valueLabel;
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
         * ファイルを Finder で表示する。RevealInFinder.app があればファイルを選択した状態で表示し、無ければ入っているフォルダーを開く
         * @param {File} scriptFile - 表示するファイル
         * @returns {void}
         */
        function revealVersionInfoScriptFile(scriptFile) {
            /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
            var hasRevealApp = (File.fs === "Macintosh")
                && (new Folder(VERSION_INFO_REVEAL_APP_PATH).exists || new File(VERSION_INFO_REVEAL_APP_PATH).exists);
            /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
            if (hasRevealApp && writeVersionInfoRevealPath(scriptFile.fsName) && new File(VERSION_INFO_REVEAL_APP_PATH).execute()) return;
            /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
            scriptFile.parent.execute();
        }

        /**
         * RevealInFinder.app に渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
         * @param {string} targetPath - 絶対パス
         * @returns {boolean} 書き出せたら true
         */
        function writeVersionInfoRevealPath(targetPath) {
            var pathFile = new File(VERSION_INFO_REVEAL_PATH_FILE);
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

        /* ============================================================
           worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
           ------------------------------------------------------------
           注意 / Notes:
           - toString() は改行を全削除するため、// 行コメント禁止・/* *\/ のみ・
             各文は必ずセミコロンで終える
           ============================================================ */
        function wkIsGuidePath(pi) {
            try { return (pi && pi.typename === "PathItem" && pi.guides === true); } catch (e) { return false; }
        }

        function wkCountHandles(pi) {
            var c = 0;
            try {
                var pts = pi.pathPoints;
                for (var i = 0; i < pts.length; i++) {
                    var p = pts[i];
                    var a = p.anchor;
                    var l = p.leftDirection;
                    var rr = p.rightDirection;
                    if (l[0] !== a[0] || l[1] !== a[1]) { c++; }
                    if (rr[0] !== a[0] || rr[1] !== a[1]) { c++; }
                }
            } catch (e) {}
            return c;
        }

        function wkCountPathStats(it, stats) {
            if (it.typename === "GroupItem") {
                for (var gi = 0; gi < it.pageItems.length; gi++) { wkCountPathStats(it.pageItems[gi], stats); }
            } else if (it.typename === "PathItem") {
                if (!wkIsGuidePath(it)) {
                    stats.pathCount++;
                    stats.anchorCount += it.pathPoints.length;
                    stats.handleCount += wkCountHandles(it);
                    if (it.closed) { stats.closedPath++; } else { stats.openPath++; }
                }
            } else if (it.typename === "CompoundPathItem") {
                for (var ci = 0; ci < it.pathItems.length; ci++) {
                    if (wkIsGuidePath(it.pathItems[ci])) { continue; }
                    stats.pathCount++;
                    stats.anchorCount += it.pathItems[ci].pathPoints.length;
                    stats.handleCount += wkCountHandles(it.pathItems[ci]);
                    if (it.pathItems[ci].closed) { stats.closedPath++; } else { stats.openPath++; }
                }
            }
        }

        function wkCollect() {
            if (app.documents.length === 0) { return "NODOC"; }
            var doc = app.activeDocument;
            var currentSelection = doc.selection;
            if (!currentSelection) { currentSelection = []; }
            var selCount = currentSelection.length;

            var allItems = doc.pageItems;

            var cpathSel = 0, cpathAll = 0, cshapeSel = 0, cshapeAll = 0;

            for (var i = 0; i < currentSelection.length; i++) {
                if (currentSelection[i].typename === "CompoundPathItem") { cpathSel++; }
                if (currentSelection[i].typename === "PluginItem") {
                    try { if (currentSelection[i].name && currentSelection[i].name.indexOf("Compound Shape") !== -1) { cshapeSel++; } } catch (e) {}
                }
            }

            for (var k = 0; k < allItems.length; k++) {
                var obj = allItems[k];
                if (obj.typename === "CompoundPathItem") { cpathAll++; }
                if (obj.typename === "PluginItem") {
                    try { if (obj.name && obj.name.indexOf("Compound Shape") !== -1) { cshapeAll++; } } catch (e2) {}
                }
            }

            var pathStatsSel = { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 };
            for (var i2 = 0; i2 < currentSelection.length; i2++) { wkCountPathStats(currentSelection[i2], pathStatsSel); }

            var pathStatsAll = { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 };
            for (var k2 = 0; k2 < allItems.length; k2++) { wkCountPathStats(allItems[k2], pathStatsAll); }

            var docName = "";
            try { docName = doc.name; } catch (e3) { docName = ""; }

            var out = [];
            out.push("selCount=" + selCount);
            out.push("docName=" + encodeURIComponent(docName));
            out.push("pathCountSel=" + pathStatsSel.pathCount);
            out.push("pathCountAll=" + pathStatsAll.pathCount);
            out.push("openSel=" + pathStatsSel.openPath);
            out.push("openAll=" + pathStatsAll.openPath);
            out.push("closedSel=" + pathStatsSel.closedPath);
            out.push("closedAll=" + pathStatsAll.closedPath);
            out.push("anchorSel=" + pathStatsSel.anchorCount);
            out.push("anchorAll=" + pathStatsAll.anchorCount);
            out.push("handleSel=" + pathStatsSel.handleCount);
            out.push("handleAll=" + pathStatsAll.handleCount);
            out.push("cpathSel=" + cpathSel);
            out.push("cpathAll=" + cpathAll);
            out.push("cshapeSel=" + cshapeSel);
            out.push("cshapeAll=" + cshapeAll);

            return "OK|" + out.join("|");
        }

        /* worker 関数は全登録（追加漏れ防止） / Register every worker function */
        var WORKER_FUNCS = [
            wkIsGuidePath,
            wkCountHandles,
            wkCountPathStats,
            wkCollect
        ];

        /* ============================================================
           BridgeTalk 委譲 / Delegation to the main engine
           ============================================================ */
        var isBusy = false;

        /**
         * worker 関数群をメインエンジンへ送り、指定した式を評価して結果を得る
         * @param {string} callExpr メインエンジンで評価する式（例: "wkCollect()"）
         * @returns {string} 戻り値文字列、またはエラー文字列（"ERR:..."）
         */
        function callMainEngine(callExpr) {
            if (isBusy) { return "ERR:BUSY"; }
            isBusy = true;

            var holder = { value: null };
            try {
                var src = "";
                for (var i = 0; i < WORKER_FUNCS.length; i++) { src += WORKER_FUNCS[i].toString(); }
                src += callExpr + ";";

                var bt = new BridgeTalk();
                bt.target = BridgeTalk.appSpecifier;
                bt.body = "eval(decodeURIComponent(\"" + encodeURIComponent(src) + "\"));";
                bt.onResult = function (res) {
                    holder.value = (res && res.body != null) ? String(res.body) : "";
                };
                bt.onError = function (err) {
                    holder.value = "ERR:" + ((err && err.body) ? err.body : "bridge");
                };
                bt.send(10);
            } catch (e) {
                holder.value = "ERR:" + e;
            } finally {
                isBusy = false;
            }

            if (holder.value === null) { return "ERR:TIMEOUT"; }
            return holder.value;
        }

        /**
         * 集計結果（OK|key=value|...）を解析する
         * @param {string} resp メインエンジンからの戻り値
         * @returns {object} キーと値のマップ（解析できない場合は null）
         */
        function parseCollect(resp) {
            if (!resp || resp.indexOf("OK|") !== 0) return null;
            var statPart = resp.substring(3);

            var map = {};
            var pairs = statPart.split("|");
            for (var i = 0; i < pairs.length; i++) {
                var eq = pairs[i].indexOf("=");
                if (eq > 0) { map[pairs[i].substring(0, eq)] = pairs[i].substring(eq + 1); }
            }
            if (map.docName != null) { try { map.docName = decodeURIComponent(map.docName); } catch (e) {} }

            return map;
        }

        /* ============================================================
           状態保持（常駐エンジン） / Session state (resident engine)
           ============================================================ */
        if (!$.global.__pathInspectorState) {
            $.global.__pathInspectorState = { location: null };
        }

        /* ============================================================
           パレット構築 / Build palette
           ============================================================ */

        /**
         * ラベルと値のペアを 1 行追加する
         * @param {object} panel 追加先のパネル
         * @param {string} labelText ラベル文字列
         * @param {number} labelWidth ラベルの幅（px）
         * @returns {object} 値表示用の statictext
         */
        function addStatRow(panel, labelText, labelWidth) {
            var row = panel.add("group");
            row.orientation = "row";
            var lbl = row.add("statictext", undefined, labelText);
            lbl.justify = "right";
            lbl.preferredSize.width = labelWidth;
            var val = row.add("statictext", undefined, "-");
            val.preferredSize.width = VALUE_WIDTH;
            return val;
        }

        /**
         * パレットを構築する
         * @returns {object} 構築済みの Window（palette）
         */
        function buildPalette() {
            var win = new Window("palette", getLabel('dialog.title'), undefined, { resizeable: false });
            setupWindow(win);

            var content = win.add("group");
            content.orientation = "column";
            content.alignChildren = ["fill", "top"];

            var v = {};

            var panelPath = content.add("panel", undefined, getLabel('panel.path'));
            setupPanel(panelPath);

            v.pathCount = addStatRow(panelPath, labelText('row.pathCount'), LABEL_WIDTH);
            v.openPath = addStatRow(panelPath, labelText('row.openPath'), LABEL_WIDTH);
            v.closedPath = addStatRow(panelPath, labelText('row.closedPath'), LABEL_WIDTH);
            v.anchors = addStatRow(panelPath, labelText('row.anchors'), LABEL_WIDTH);
            v.handles = addStatRow(panelPath, labelText('row.handles'), LABEL_WIDTH);
            v.compoundPath = addStatRow(panelPath, labelText('row.compoundPath'), LABEL_WIDTH);
            v.compoundShape = addStatRow(panelPath, labelText('row.compoundShape'), LABEL_WIDTH);

            /* ステータス / Status line */
            var statusText = win.add("statictext", undefined, getLabel('status.ready'));
            statusText.alignment = ["fill", "bottom"];

            /**
             * ステータス行を更新する
             * @param {string} msg 表示メッセージ
             * @returns {void}
             */
            function setStatus(msg) { statusText.text = msg; }

            /**
             * 集計値をパネルへ反映する
             * @param {object} m 集計結果のマップ
             * @returns {void}
             */
            function applyValues(m) {
                v.pathCount.text = m.pathCountSel + " / " + m.pathCountAll;
                v.openPath.text = m.openSel + " / " + m.openAll;
                v.closedPath.text = m.closedSel + " / " + m.closedAll;
                v.anchors.text = m.anchorSel + " / " + m.anchorAll;
                v.handles.text = m.handleSel + " / " + m.handleAll;
                v.compoundPath.text = m.cpathSel + " / " + m.cpathAll;
                v.compoundShape.text = m.cshapeSel + " / " + m.cshapeAll;
            }

            /**
             * 表示中の値をクリアする
             * @returns {void}
             */
            function clearValues() {
                for (var kk in v) { if (v.hasOwnProperty(kk)) { try { v[kk].text = "-"; } catch (e) {} } }
            }

            /**
             * メインエンジンへ集計を委譲し、結果を表示に反映する
             * @returns {void}
             */
            function refresh() {
                setStatus(getLabel('status.busy'));
                var resp = callMainEngine("wkCollect()");

                if (resp === "ERR:BUSY") { setStatus(getLabel('status.busy')); return; }
                if (resp === null || resp === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
                if (resp === "NODOC") { setStatus(getLabel('status.noDoc')); clearValues(); return; }
                if (resp.indexOf("ERR:") === 0) { setStatus(labelValueText('status.error', resp.substring(4))); return; }

                var map = parseCollect(resp);
                if (!map) { setStatus(getLabel('status.error')); return; }

                applyValues(map);

                var selN = parseInt(map.selCount, 10) || 0;
                if (selN > 0) {
                    setStatus(getLabel('status.selectedPrefix') + selN + getLabel('status.selectedSuffix'));
                } else {
                    setStatus(getLabel('status.wholeDoc'));
                }
            }

            /**
             * 集計結果をテキストファイルとしてデスクトップへ書き出す
             * @returns {void}
             */
            function exportReport() {
                setStatus(getLabel('status.busy'));
                var resp = callMainEngine("wkCollect()");
                if (resp === "NODOC") { setStatus(getLabel('status.noDoc')); return; }
                if (resp === null || resp === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
                if (typeof resp === "string" && resp.indexOf("ERR:") === 0) { setStatus(labelValueText('status.error', resp.substring(4))); return; }

                var m = parseCollect(resp);
                if (!m) { setStatus(getLabel('status.error')); return; }

                try {
                    var fullName = m.docName || "";
                    var baseName = fullName.replace(/\.[^\.]+$/, "");
                    var today = new Date();
                    var yyyy = today.getFullYear();
                    var mm = ("0" + (today.getMonth() + 1)).slice(-2);
                    var dd = ("0" + today.getDate()).slice(-2);
                    var dateStr = yyyy + mm + dd;

                    var path = Folder.desktop + "/path-" + baseName + "-" + dateStr + ".txt";
                    var file = new File(path);

                    /**
                     * 「選択 / 全体」形式で 1 行書き出す
                     * @param {string} path2 ラベルのドットパス
                     * @param {string} selVal 選択側の値
                     * @param {string} allVal 全体側の値
                     * @returns {void}
                     */
                    function wPair(path2, selVal, allVal) { file.writeln(LX(path2) + " " + selVal + " / " + allVal); }

                    /**
                     * セクション見出しを書き出す
                     * @param {string} path2 見出しのドットパス
                     * @returns {void}
                     */
                    function wSection(path2) { file.writeln(""); file.writeln(getLabel(path2)); }

                    if (file.open("w")) {
                        file.writeln(getLabel('report.title'));
                        file.writeln(labelText('report.document') + " " + fullName);
                        file.writeln(labelText('report.date') + " " + yyyy + "-" + mm + "-" + dd);
                        file.writeln("");
                        file.writeln(getLabel('report.valueNote'));

                        wSection('section.paths');
                        wPair('row.pathCount', m.pathCountSel, m.pathCountAll);
                        wPair('row.openPath', m.openSel, m.openAll);
                        wPair('row.closedPath', m.closedSel, m.closedAll);
                        wPair('row.anchors', m.anchorSel, m.anchorAll);
                        wPair('row.handles', m.handleSel, m.handleAll);
                        wPair('row.compoundPath', m.cpathSel, m.cpathAll);
                        wPair('row.compoundShape', m.cshapeSel, m.cshapeAll);

                        file.close();
                        setStatus(getLabel('status.exportedPrefix') + path);
                    } else {
                        setStatus(getLabel('status.exportFailOpen'));
                    }
                } catch (err) {
                    setStatus(labelValueText('status.error', err));
                }
            }

            /* --- ボタン行 / Button row --- */
            var btnRow = win.add("group");
            btnRow.orientation = "row";
            btnRow.alignChildren = ["fill", "center"];
            btnRow.alignment = ["fill", "bottom"];

            var btnLeft = btnRow.add("group");
            btnLeft.alignChildren = ["left", "center"];
            var btnExport = btnLeft.add("button", undefined, getLabel('button.exportPreset'));
            btnExport.helpTip = getLabel('hint.esc');

            var spacer = btnRow.add("statictext", undefined, "");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            var btnRight = btnRow.add("group");
            btnRight.alignChildren = ["right", "center"];
            var btnRefresh = btnRight.add("button", undefined, getLabel('button.refresh'));
            btnRefresh.helpTip = getLabel('hint.refresh') + "\n" + getLabel('hint.esc');

            btnExport.onClick = exportReport;
            btnRefresh.onClick = refresh;

            /* キー操作 / Key handling
               Esc: 閉じる / close
               ⌘R: 更新 / Cmd+R refresh */
            addKeyShortcuts(win, {
                "Escape": { target: function () { try { win.close(); } catch (e1) {} }, inFields: true },
                "Cmd+R": { target: refresh, inFields: true }
            });

            try { win.opacity = PALETTE_OPACITY; } catch (e) {}

            /* 表示直後に一度集計 / Count once right after showing */
            win.onShow = function () {
                refresh();
            };

            return win;
        }

        /* ============================================================
           位置の記憶・復元 / Remember & restore location
           ============================================================ */

        /**
         * 記憶した位置へパレットを復元する（未記憶なら中央）
         * @param {object} win 対象の Window
         * @returns {void}
         */
        function restoreLocation(win) {
            try {
                var loc = $.global.__pathInspectorState.location;
                if (loc && loc.length === 2) {
                    win.location = [loc[0], loc[1]];
                } else {
                    win.center();
                }
            } catch (e) {
                win.center();
            }
        }

        /**
         * パレットの現在位置を記憶する
         * @param {object} win 対象の Window
         * @returns {void}
         */
        function rememberLocation(win) {
            try {
                if (win.location && win.location.length === 2) {
                    $.global.__pathInspectorState.location = [win.location[0], win.location[1]];
                }
            } catch (e) {}
        }

        /* ============================================================
           起動 / Entry point
           ============================================================ */

        /**
         * パレットを表示する（多重起動時は既存を閉じてから再表示）
         * @returns {void}
         */
        function showPalette() {
            if ($.global.__PathInspectorPalette) {
                try { $.global.__PathInspectorPalette.close(); } catch (e) {}
                $.global.__PathInspectorPalette = null;
            }

            var win = buildPalette();

            /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
            $.global.__PathInspectorPalette = win;
            win.onClose = function () {
                rememberLocation(win);
                app.redraw();
                $.global.__PathInspectorPalette = null;
            };

            restoreLocation(win);
            addVersionFooter(win);
            win.show();
        }

        showPalette();

    })();

})();
