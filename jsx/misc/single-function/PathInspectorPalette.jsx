#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

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
var SCRIPT_VERSION  = "v1.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-31";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathInspectorPalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathInspectorPalette.md"; /* README (English) */

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
                bt.target = "illustrator";
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
            var win = new Window("palette", getLabel('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
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
            win.show();
        }

        showPalette();

    })();

})();
