#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中またはドキュメント全体のオブジェクトを集計し、2カラムの常駐パレットで表示します。
テキスト・配置画像・透明・グループ・パス・ガイドの内訳を確認でき、選択オブジェクトのメモの閲覧・編集にも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionInspectorPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nefcb1ce828ce

### Overview

Tallies the objects in the selection, or in the whole document, and shows them in a two-column persistent palette.
It breaks down text, placed images, transparency, groups, paths and guides, and can view and edit the note on the selected object.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectionInspectorPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectionInspectorPalette";    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-10";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionInspectorPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectionInspectorPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nefcb1ce828ce"; /* 紹介記事 / article URL */

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

    var LABEL_WIDTH_LEFT = 100;               /* 左カラムの項目名の幅 / Label width, left column */
    var LABEL_WIDTH_RIGHT = 130;              /* 右カラムの項目名の幅 / Label width, right column */
    var VALUE_WIDTH = 90;                     /* 値の幅 / Value width */
    var VIEW_MARGINS = [10, 15, 10, 10];      /* 表示切り替え（情報／メモ）の各ビューの余白 / Margins of the Info / Notes views */
    var MEMO_PREVIEW_SIZE = [160, 50];        /* 情報タブのメモ表示の寸法 / Note preview size on the Info tab */
    var MEMO_FIELD_SIZE = [340, 44];          /* メモタブの入力欄の寸法 / Note field size on the Notes tab */
    var PALETTE_OPACITY = 0.97;               /* パレットの不透明度 / Palette opacity */
    var BRIDGE_TIMEOUT_SECONDS = 300;         /* 集計の応答を待つ秒数（大きなドキュメント向け）/ Seconds to wait for the count, for large documents */

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
            title: { ja: "選択／全体オブジェクトのカウント", en: "Selection / All Objects Count" }
        },
        report: {
            title: { ja: "Selection Inspector Report", en: "Selection Inspector Report" },
            document: { ja: "ドキュメント", en: "Document" },
            date: { ja: "日付", en: "Date" },
            valueNote: {
                ja: "※ 値は『選択 / 全体』の形式です（アートボードのみ全体）",
                en: "Note: values are formatted as 'Selection / All' (Artboards is All only)"
            }
        },
        panel: {
            basics: { ja: "基本", en: "Basics" },
            texts: { ja: "テキスト", en: "Text Frames" },
            charPara: { ja: "文字・段落", en: "Characters & Paragraphs" },
            images: { ja: "配置画像", en: "Images" },
            memo: { ja: "メモ", en: "Notes" },
            group: { ja: "グループ", en: "Groups" },
            transparency: { ja: "透明", en: "Transparency" },
            path: { ja: "パス", en: "Paths" },
            guide: { ja: "ガイド", en: "Guides" }
        },
        radio: {
            info: { ja: "情報", en: "Info" },
            memo: { ja: "メモ", en: "Notes" }
        },
        fieldLabel: {
            artboards: { ja: "アートボード", en: "Artboards" },
            objects: { ja: "オブジェクト", en: "Objects" },
            texts: { ja: "テキスト", en: "Text Frames" },
            pointText: { ja: "ポイント文字", en: "Point Type" },
            areaText: { ja: "エリア内文字", en: "Area Type" },
            pathText: { ja: "パス上文字", en: "Path Text" },
            chars: { ja: "文字数", en: "Characters" },
            paras: { ja: "段落数", en: "Paragraphs" },
            forcedBreaks: { ja: "強制改行", en: "Line Breaks" },
            linked: { ja: "リンク", en: "Linked Images" },
            embed: { ja: "埋め込み", en: "Embedded Images" },
            broken: { ja: "リンク切れ", en: "Broken Links" },
            group: { ja: "グループ", en: "Groups" },
            clipGroup: { ja: "クリップグループ", en: "Clipping Groups" },
            opacityLt100: { ja: "不透明度<100", en: "Opacity < 100" },
            blendNotNormal: { ja: "描画モード≠通常", en: "Blend Mode != Normal" },
            pathCount: { ja: "パス", en: "Paths" },
            openPath: { ja: "オープンパス", en: "Open Paths" },
            closedPath: { ja: "クローズパス", en: "Closed Paths" },
            anchors: { ja: "アンカーポイント", en: "Anchor Points" },
            handles: { ja: "ハンドル", en: "Handles" },
            compoundPath: { ja: "複合パス", en: "Compound Paths" },
            compoundShape: { ja: "複合シェイプ", en: "Compound Shapes" },
            rulerGuides: { ja: "ルーラーガイド", en: "Ruler Guides" },
            artboardGuides: { ja: "アートボードガイド", en: "Artboard Guides" },
            otherGuides: { ja: "その他のガイド", en: "Other Guides" }
        },
        button: {
            refresh: { ja: "更新", en: "Refresh" },
            exportReport: { ja: "書き出し", en: "Export" },
            applyMemo: { ja: "適用", en: "Apply" }
        },
        memo: {
            multiple: {
                ja: "複数のメモがあります。\n「メモ」タブで確認",
                en: "Multiple notes found.\nSee the \"Notes\" tab."
            },
            none: { ja: "選択オブジェクトがありません", en: "No objects selected" }
        },
        status: {
            ready: { ja: "準備完了", en: "Ready" },
            noDoc: { ja: "ドキュメントが開かれていません", en: "No document open" },
            wholeDoc: { ja: "選択なし（全体を集計）", en: "No selection (counting all)" },
            selectedPrefix: { ja: "選択 ", en: "Selected " },
            selectedSuffix: { ja: " 件を集計", en: " object(s)" },
            timeout: { ja: "Illustrator から応答がありません", en: "No response from Illustrator" },
            busy: { ja: "処理中です", en: "Busy" },
            counting: { ja: "集計しています...", en: "Counting..." },
            error: { ja: "エラー", en: "Error" },
            memoApplied: { ja: "メモを適用しました", en: "Note applied" },
            selChanged: { ja: "選択が変わりました。更新してください", en: "Selection changed. Please refresh." },
            exportedPrefix: { ja: "書き出しました: ", en: "Exported: " },
            exportFailOpen: { ja: "ファイルを開けませんでした", en: "Failed to open the file" }
        },
        tooltip: {
            shortcut: { ja: "⌥I: 情報  ⌥M: メモ  ⌘R: 更新", en: "⌥I: Info  ⌥M: Notes  ⌘R: Refresh" },
            refresh: { ja: "選択内容を再集計（⌘R）", en: "Recount selection (Cmd+R)" },
            esc: { ja: "Esc で閉じる", en: "Press Esc to close" },
            exportReport: {
                ja: "集計結果をデスクトップにテキストファイル（count-ドキュメント名-日付.txt）で書き出します",
                en: "Writes the tallies to a text file on the desktop (count-<document>-<date>.txt)"
            },
            applyMemo: {
                ja: "この欄の内容を、対応するオブジェクトのメモに設定します（上から、同じ高さなら左から順）",
                en: "Sets this text as the note of the matching object (ordered top to bottom, then left to right)"
            },
            statValue: {
                ja: "選択 / 全体 の数です（アートボードは全体のみ）",
                en: "Selection / All (Artboards shows All only)"
            },
            handles: {
                ja: "アンカーポイントから伸びている方向線の数です",
                en: "Number of direction handles that stick out of their anchor points"
            },
            rulerGuides: {
                ja: "どのアートボードの幅・高さよりも長い水平／垂直のガイドです",
                en: "Horizontal or vertical guides longer than every artboard's width and height"
            },
            artboardGuides: {
                ja: "長さがいずれかのアートボードの幅または高さと同じ（±0.5pt）水平／垂直のガイドです",
                en: "Horizontal or vertical guides as long as an artboard's width or height (±0.5 pt)"
            },
            otherGuides: {
                ja: "ルーラーガイドにもアートボードガイドにも当たらないガイドです（斜め・閉じたパスなど）",
                en: "Guides that are neither ruler guides nor artboard guides (diagonal, closed, and so on)"
            }
        }
    };

    /**
     * 書き出し用の項目名を返す（言語にかかわらず半角コロン）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} 半角コロン付きの項目名
     */
    function reportLabelText(labelPath) {
        return getLabel(labelPath) + ":";
    }

    // =========================================
    // 集計項目の構成 / Tally layout
    // =========================================

    /* パネルごとの行。label は項目名（LABELS.fieldLabel）、key は集計結果のキー（<key>Sel / <key>All）、
       tip は項目名の helpTip（LABELS.tooltip）。アートボードは全体の値だけ、オブジェクト数はキー名が別形式
       Rows per panel: label = row label, key = result key (<key>Sel / <key>All), tip = helpTip on the label */
    var STAT_ROWS = {
        basics: [
            { label: "artboards", allKey: "artboards" },
            { label: "objects", selKey: "selCount", allKey: "allCount" }
        ],
        texts: [
            { label: "texts", key: "text" },
            { label: "pointText", key: "point" },
            { label: "areaText", key: "area" },
            { label: "pathText", key: "path" }
        ],
        charPara: [
            { label: "chars", key: "char" },
            { label: "paras", key: "para" },
            { label: "forcedBreaks", key: "fb" }
        ],
        images: [
            { label: "linked", key: "linked" },
            { label: "embed", key: "embed" },
            { label: "broken", key: "broken" }
        ],
        group: [
            { label: "group", key: "group" },
            { label: "clipGroup", key: "clip" }
        ],
        transparency: [
            { label: "opacityLt100", key: "opacity" },
            { label: "blendNotNormal", key: "blend" }
        ],
        path: [
            { label: "pathCount", key: "pathCount" },
            { label: "openPath", key: "open" },
            { label: "closedPath", key: "closed" },
            { label: "anchors", key: "anchor" },
            { label: "handles", key: "handle", tip: "handles" },
            { label: "compoundPath", key: "cpath" },
            { label: "compoundShape", key: "cshape" }
        ],
        guide: [
            { label: "rulerGuides", key: "ruler", tip: "rulerGuides" },
            { label: "artboardGuides", key: "abguide", tip: "artboardGuides" },
            { label: "otherGuides", key: "otherguide", tip: "otherGuides" }
        ]
    };

    /* パレットの2カラムに並べるパネル（memo はメモの表示欄）/ Panels in the palette's two columns */
    var PALETTE_COLUMNS = [
        ["basics", "texts", "charPara", "images", "memo"],
        ["group", "transparency", "path", "guide"]
    ];

    /* 書き出すセクションの順番 / Section order in the exported report */
    var REPORT_SECTIONS = ["basics", "texts", "charPara", "images", "memo", "transparency", "group", "path", "guide"];

    /**
     * 集計結果から1行ぶんの表示値を作る（「選択 / 全体」、全体のみの行は全体の値）
     * @param {Object} statRow - STAT_ROWS の1行
     * @param {Object} statMap - parseCollect() の map
     * @returns {string} 表示値
     */
    function formatStatValue(statRow, statMap) {
        if (!statRow.key && !statRow.selKey) return statMap[statRow.allKey];
        var selKey = statRow.selKey || (statRow.key + "Sel");
        var allKey = statRow.allKey || (statRow.key + "All");
        return statMap[selKey] + " / " + statMap[allKey];
    }

    /**
     * 空でないメモだけを取り出す
     * @param {string[]} memoList - メモの一覧
     * @returns {string[]} 空でないメモ
     */
    function collectNonEmptyNotes(memoList) {
        var nonEmptyNotes = [];
        for (var i = 0; i < memoList.length; i++) {
            if (memoList[i] && memoList[i] !== "") { nonEmptyNotes.push(memoList[i]); }
        }
        return nonEmptyNotes;
    }

    // =========================================
    // worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
    // -----------------------------------------
    // toString() で文字列にして BridgeTalk で送るため、JSDoc は付けない（構文エラーになる）
    // toString() は改行を全削除するため、関数内は // 行コメント禁止・/* */ のみ・各文は必ずセミコロンで終える
    // 関数を足したら WORKER_FUNCS にも登録する
    // Serialized with toString() and sent through BridgeTalk: no JSDoc, no // comments inside,
    // and every statement ends with a semicolon. Register new functions in WORKER_FUNCS.
    // =========================================

    /* アートボードの辺の長さと最大値を1度だけ控える / Read artboard side lengths and the largest one once */
    function wkReadArtboardSides(doc) {
        var artboardSides = { lengths: [], maxSpan: 0 };
        try {
            var artboards = doc.artboards;
            var artboardCount = artboards.length;
            for (var i = 0; i < artboardCount; i++) {
                var artboardRect = artboards[i].artboardRect;
                var artboardWidth = Math.abs(artboardRect[2] - artboardRect[0]);
                var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);
                artboardSides.lengths.push(artboardWidth, artboardHeight);
                if (artboardWidth > artboardSides.maxSpan) { artboardSides.maxSpan = artboardWidth; }
                if (artboardHeight > artboardSides.maxSpan) { artboardSides.maxSpan = artboardHeight; }
            }
        } catch (e) {}
        return artboardSides;
    }

    /* 水平・垂直の2点のガイドなら長さを、それ以外は -1 を返す / Length of a straight 2-point guide, otherwise -1 */
    function wkStraightGuideLength(pathItem) {
        if (pathItem.closed) { return -1; }
        var pathPoints = pathItem.pathPoints;
        if (pathPoints.length !== 2) { return -1; }
        var anchor0 = pathPoints[0].anchor;
        var anchor1 = pathPoints[1].anchor;
        var straightTolerance = 0.01;
        var isVertical = Math.abs(anchor0[0] - anchor1[0]) <= straightTolerance;
        var isHorizontal = Math.abs(anchor0[1] - anchor1[1]) <= straightTolerance;
        if (!(isVertical || isHorizontal)) { return -1; }
        var dx = anchor0[0] - anchor1[0];
        var dy = anchor0[1] - anchor1[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    /* 長さがいずれかのアートボードの幅または高さと同じか / Whether a length matches an artboard side */
    function wkMatchesArtboardSide(guideLength, artboardSides) {
        var lengthTolerance = 0.5;
        for (var i = 0; i < artboardSides.lengths.length; i++) {
            if (Math.abs(guideLength - artboardSides.lengths[i]) <= lengthTolerance) { return true; }
        }
        return false;
    }

    /* ガイド1本を種類別に集計に足す（ルーラー・アートボード・その他）/ Classify one guide into the tally */
    function wkAddGuideCount(pathItem, tally, artboardSides) {
        var guideLength = -1;
        try { guideLength = wkStraightGuideLength(pathItem); } catch (e) {}
        if (guideLength >= 0 && (!(artboardSides.maxSpan > 0) || guideLength > artboardSides.maxSpan)) { tally.ruler++; }
        else if (guideLength >= 0 && wkMatchesArtboardSide(guideLength, artboardSides)) { tally.abguide++; }
        else { tally.otherguide++; }
    }

    /* 不透明度が100未満・描画モードが通常以外なら集計に足す / Count opacity below 100 and non-normal blend mode */
    function wkAddTransparency(pageItem, tally) {
        try { if (typeof pageItem.opacity === "number" && pageItem.opacity < 100) { tally.opacity++; } } catch (e) {}
        try { if (pageItem.blendingMode !== undefined && pageItem.blendingMode !== BlendModes.NORMAL) { tally.blend++; } } catch (e2) {}
    }

    /* アンカーから伸びている方向線の数を数える / Count direction handles that stick out of their anchors */
    function wkCountHandles(pathPoints, pointCount) {
        var handleCount = 0;
        try {
            for (var i = 0; i < pointCount; i++) {
                var pathPoint = pathPoints[i];
                var anchor = pathPoint.anchor;
                var leftDirection = pathPoint.leftDirection;
                var rightDirection = pathPoint.rightDirection;
                if (leftDirection[0] !== anchor[0] || leftDirection[1] !== anchor[1]) { handleCount++; }
                if (rightDirection[0] !== anchor[0] || rightDirection[1] !== anchor[1]) { handleCount++; }
            }
        } catch (e) {}
        return handleCount;
    }

    /* パス1本を集計に足す（ガイドはガイドとして数える）/ Add one path into the tally, guides counted separately */
    function wkAddPath(pathItem, tally, artboardSides) {
        var isGuide = false;
        try { isGuide = (pathItem.guides === true); } catch (e) {}
        if (isGuide) {
            if (artboardSides) { wkAddGuideCount(pathItem, tally, artboardSides); }
            return;
        }
        var pathStats = tally.path;
        var pathPoints = pathItem.pathPoints;
        var pointCount = pathPoints.length;
        pathStats.pathCount++;
        pathStats.anchorCount += pointCount;
        pathStats.handleCount += wkCountHandles(pathPoints, pointCount);
        if (pathItem.closed) { pathStats.closedPath++; } else { pathStats.openPath++; }
    }

    /* 文字列に含まれる改行（\n）を数える / Count \n in a string */
    function wkCountForcedBreaks(contentsText) {
        if (!contentsText || !contentsText.length) { return 0; }
        var breakMatches = contentsText.match(/[\n]/g);
        return breakMatches ? breakMatches.length : 0;
    }

    /* テキスト1つの文字数・段落数・種類を集計に足す / Add one text frame into the text stats */
    function wkAddText(textFrame, textStats) {
        textStats.textCount++;
        try { textStats.charCount += textFrame.characters.length; } catch (e) {}
        try { textStats.paraCount += textFrame.paragraphs.length; } catch (e2) {}
        try { textStats.forcedBreakCount += wkCountForcedBreaks(textFrame.contents); } catch (e3) {}
        var textKind = textFrame.kind;
        if (textKind === TextType.POINTTEXT) { textStats.pointText++; }
        else if (textKind === TextType.AREATEXT) { textStats.areaText++; }
        else if (textKind === TextType.PATHTEXT) { textStats.pathText++; }
    }

    /* 進捗の状態を作る（ウィンドウは時間がかかったときだけ開く）/ Create the progress state; the window opens only when counting runs long */
    function wkCreateProgress(totalCount, encodedTitle) {
        var progressTitle = "";
        try { progressTitle = decodeURIComponent(encodedTitle); } catch (e) {}
        return { totalCount: totalCount, doneCount: 0, startTime: new Date().getTime(), title: progressTitle, progressWindow: null, progressBar: null, progressText: null, failed: false };
    }

    /* 1件進め、0.5秒を過ぎていたら進捗ウィンドウを開いて更新する / Advance by one; open and update the window once 0.5 s have passed */
    function wkStepProgress(progress) {
        progress.doneCount++;
        if (progress.failed || progress.doneCount % 50 !== 0) { return; }
        if (!progress.progressWindow) {
            if (new Date().getTime() - progress.startTime < 500) { return; }
            try {
                var progressWindow = new Window("palette", progress.title);
                progressWindow.orientation = "column";
                progressWindow.alignChildren = ["fill", "center"];
                progressWindow.margins = [15, 15, 15, 15];
                progressWindow.spacing = 8;
                var progressText = progressWindow.add("statictext", undefined, progress.title);
                progressText.preferredSize.width = 300;
                var progressBar = progressWindow.add("progressbar", undefined, 0, progress.totalCount);
                progressBar.minvalue = 0;
                progressBar.maxvalue = progress.totalCount;
                progressBar.preferredSize = [300, 8];
                progressWindow.center();
                progressWindow.show();
                progress.progressWindow = progressWindow;
                progress.progressBar = progressBar;
                progress.progressText = progressText;
            } catch (e) { progress.failed = true; return; }
        }
        try {
            progress.progressBar.value = Math.min(progress.doneCount, progress.totalCount);
            progress.progressText.text = progress.title + " (" + Math.min(progress.doneCount, progress.totalCount) + " / " + progress.totalCount + ")";
            progress.progressWindow.update();
        } catch (e2) {}
    }

    /* 進捗ウィンドウを閉じる / Close the progress window */
    function wkCloseProgress(progress) {
        if (!progress || !progress.progressWindow) { return; }
        try { progress.progressWindow.close(); } catch (e) {}
        progress.progressWindow = null;
    }

    /* グループ・複合パスの中のパスとテキストを集計に足す（選択用）/ Add paths and text inside groups and compound paths (selection only) */
    function wkAddNestedStats(pageItem, itemType, tally, progress) {
        var i;
        if (itemType === "GroupItem") {
            var childItems = pageItem.pageItems;
            var childCount = childItems.length;
            for (i = 0; i < childCount; i++) {
                var childItem = childItems[i];
                wkStepProgress(progress);
                wkAddNestedStats(childItem, childItem.typename, tally, progress);
            }
        } else if (itemType === "CompoundPathItem") {
            var memberPaths = pageItem.pathItems;
            var memberCount = memberPaths.length;
            for (i = 0; i < memberCount; i++) { wkAddPath(memberPaths[i], tally, null); }
        } else if (itemType === "PathItem") {
            wkAddPath(pageItem, tally, null);
        } else if (itemType === "TextFrame") {
            wkAddText(pageItem, tally.text);
        }
    }

    /* オブジェクトの一覧を種類別に集計する。isFlat は入れ子が一覧に含まれる doc.pageItems 用 / Tally a list of items; isFlat for doc.pageItems, which already lists nested items */
    function wkTallyItems(pageItems, artboardSides, isFlat, progress) {
        var tally = {
            cpath: 0, cshape: 0, opacity: 0, blend: 0, ruler: 0, abguide: 0, otherguide: 0,
            linked: 0, embed: 0, broken: 0, group: 0, clip: 0,
            path: { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 },
            text: { textCount: 0, charCount: 0, paraCount: 0, forcedBreakCount: 0, pointText: 0, areaText: 0, pathText: 0 }
        };
        var itemCount = pageItems.length;
        for (var i = 0; i < itemCount; i++) {
            var pageItem = pageItems[i];
            var itemType = pageItem.typename;
            wkStepProgress(progress);
            wkAddTransparency(pageItem, tally);
            if (itemType === "PathItem") {
                wkAddPath(pageItem, tally, artboardSides);
            } else if (itemType === "CompoundPathItem") {
                tally.cpath++;
                if (!isFlat) {
                    var memberPaths = pageItem.pathItems;
                    var memberCount = memberPaths.length;
                    for (var j = 0; j < memberCount; j++) { wkAddPath(memberPaths[j], tally, artboardSides); }
                }
            } else if (itemType === "TextFrame") {
                wkAddText(pageItem, tally.text);
            } else if (itemType === "GroupItem") {
                tally.group++;
                if (pageItem.clipped) { tally.clip++; }
                if (!isFlat) { wkAddNestedStats(pageItem, itemType, tally, progress); }
            } else if (itemType === "PlacedItem") {
                if (pageItem.embedded) { tally.embed++; }
                else {
                    tally.linked++;
                    try { var linkedFile = pageItem.file; if (!linkedFile || !linkedFile.exists) { tally.broken++; } } catch (e) { tally.broken++; }
                }
            } else if (itemType === "PluginItem") {
                try { if (pageItem.name && pageItem.name.indexOf("Compound Shape") !== -1) { tally.cshape++; } } catch (e2) {}
            }
        }
        return tally;
    }

    /* 0以上の整数を12桁の0埋め文字列にする / Zero-pad a non-negative integer to 12 digits */
    function wkPadSortKey(intValue) {
        var keyText = String(intValue);
        while (keyText.length < 12) { keyText = "0" + keyText; }
        return keyText;
    }

    /* 選択を上から、同じ高さなら左から並べ替える（比較関数なしの sort で）/ Sort the selection top to bottom, then left to right, without a comparator */
    function wkSortSelection(selectedItems) {
        var itemCount = selectedItems.length;
        var coordOffset = 100000000;
        var sortKeys = [];
        for (var i = 0; i < itemCount; i++) {
            var itemTop = 0;
            var itemLeft = 0;
            try { var itemBounds = selectedItems[i].geometricBounds; itemTop = itemBounds[1]; itemLeft = itemBounds[0]; } catch (e) {}
            sortKeys.push(wkPadSortKey(Math.round((coordOffset - itemTop) * 1000)) + wkPadSortKey(Math.round((coordOffset + itemLeft) * 1000)) + wkPadSortKey(i));
        }
        sortKeys.sort();
        var sortedItems = [];
        for (var k = 0; k < itemCount; k++) { sortedItems.push(selectedItems[parseInt(sortKeys[k].substring(24), 10)]); }
        return sortedItems;
    }

    /* 「key+Sel=値」「key+All=値」の組を足す / Push a Sel/All pair */
    function wkPushPair(resultPairs, statKey, selValue, allValue) {
        resultPairs.push(statKey + "Sel=" + selValue);
        resultPairs.push(statKey + "All=" + allValue);
    }

    /* 集計して「OK|key=value|...|MEMO|件数|メモ...」の文字列で返す / Collect and return OK|key=value|...|MEMO|count|notes */
    function wkCollect(encodedProgressTitle) {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        if (!selectedItems) { selectedItems = []; }
        var selCount = selectedItems.length;
        var docItems = doc.pageItems;
        var allCount = docItems.length;
        var artboardSides = wkReadArtboardSides(doc);

        /* 選択にグループがあると中まで数えるので、進捗の分母は全体の件数を上限として見込む / Selected groups are counted inside, so budget up to the document count for them */
        var selWeight = selCount;
        for (var si = 0; si < selCount; si++) {
            if (selectedItems[si].typename === "GroupItem") { selWeight = allCount; break; }
        }
        var progress = wkCreateProgress(allCount + selWeight, encodedProgressTitle);
        var selTally, allTally;
        try {
            allTally = wkTallyItems(docItems, artboardSides, true, progress);
            selTally = wkTallyItems(selectedItems, artboardSides, false, progress);
        } finally {
            wkCloseProgress(progress);
        }

        var sortedItems = wkSortSelection(selectedItems);
        var memoParts = [];
        for (var i = 0; i < sortedItems.length; i++) {
            var noteText = "";
            try { noteText = sortedItems[i].note || ""; } catch (e) { noteText = ""; }
            memoParts.push(encodeURIComponent(noteText));
        }

        var docName = "";
        try { docName = doc.name; } catch (e2) { docName = ""; }

        var resultPairs = [];
        resultPairs.push("selCount=" + selCount);
        resultPairs.push("allCount=" + allCount);
        resultPairs.push("artboards=" + doc.artboards.length);
        resultPairs.push("docName=" + encodeURIComponent(docName));
        wkPushPair(resultPairs, "text", selTally.text.textCount, allTally.text.textCount);
        wkPushPair(resultPairs, "point", selTally.text.pointText, allTally.text.pointText);
        wkPushPair(resultPairs, "area", selTally.text.areaText, allTally.text.areaText);
        wkPushPair(resultPairs, "path", selTally.text.pathText, allTally.text.pathText);
        wkPushPair(resultPairs, "char", selTally.text.charCount, allTally.text.charCount);
        wkPushPair(resultPairs, "para", selTally.text.paraCount, allTally.text.paraCount);
        wkPushPair(resultPairs, "fb", selTally.text.forcedBreakCount, allTally.text.forcedBreakCount);
        wkPushPair(resultPairs, "linked", selTally.linked, allTally.linked);
        wkPushPair(resultPairs, "embed", selTally.embed, allTally.embed);
        wkPushPair(resultPairs, "broken", selTally.broken, allTally.broken);
        wkPushPair(resultPairs, "group", selTally.group, allTally.group);
        wkPushPair(resultPairs, "clip", selTally.clip, allTally.clip);
        wkPushPair(resultPairs, "opacity", selTally.opacity, allTally.opacity);
        wkPushPair(resultPairs, "blend", selTally.blend, allTally.blend);
        wkPushPair(resultPairs, "pathCount", selTally.path.pathCount, allTally.path.pathCount);
        wkPushPair(resultPairs, "open", selTally.path.openPath, allTally.path.openPath);
        wkPushPair(resultPairs, "closed", selTally.path.closedPath, allTally.path.closedPath);
        wkPushPair(resultPairs, "anchor", selTally.path.anchorCount, allTally.path.anchorCount);
        wkPushPair(resultPairs, "handle", selTally.path.handleCount, allTally.path.handleCount);
        wkPushPair(resultPairs, "cpath", selTally.cpath, allTally.cpath);
        wkPushPair(resultPairs, "cshape", selTally.cshape, allTally.cshape);
        wkPushPair(resultPairs, "ruler", selTally.ruler, allTally.ruler);
        wkPushPair(resultPairs, "abguide", selTally.abguide, allTally.abguide);
        wkPushPair(resultPairs, "otherguide", selTally.otherguide, allTally.otherguide);

        return "OK|" + resultPairs.join("|") + "|MEMO|" + selCount + "|" + memoParts.join("|");
    }

    /* 並べ替えた選択の itemIndex 番目にメモを設定する / Set the note on the itemIndex-th item of the sorted selection */
    function wkApplyMemo(itemIndex, encodedNote) {
        if (app.documents.length === 0) { return "NODOC"; }
        var selectedItems = app.activeDocument.selection;
        if (!selectedItems) { selectedItems = []; }
        if (itemIndex < 0 || itemIndex >= selectedItems.length) { return "IDX"; }
        var sortedItems = wkSortSelection(selectedItems);
        try { sortedItems[itemIndex].note = decodeURIComponent(encodedNote); } catch (e) { return "ERR:" + e; }
        app.redraw();
        return "OK";
    }

    /* worker 関数は全登録（追加漏れ防止） / Register every worker function */
    var WORKER_FUNCS = [
        wkReadArtboardSides,
        wkStraightGuideLength,
        wkMatchesArtboardSide,
        wkAddGuideCount,
        wkAddTransparency,
        wkCountHandles,
        wkAddPath,
        wkCountForcedBreaks,
        wkAddText,
        wkAddNestedStats,
        wkCreateProgress,
        wkStepProgress,
        wkCloseProgress,
        wkTallyItems,
        wkPadSortKey,
        wkSortSelection,
        wkPushPair,
        wkCollect,
        wkApplyMemo
    ];

    // =========================================
    // BridgeTalk 委譲 / Delegation to the main engine
    // =========================================

    var isBusy = false;

    /**
     * worker 関数一式をメインエンジンへ送り、指定の呼び出しを実行して結果を受け取る
     * @param {string} callExpression - メインエンジンで評価する呼び出し式（"wkCollect()" など）
     * @returns {string} 戻り値。失敗時は "ERR:" で始まる文字列
     */
    function callMainEngine(callExpression) {
        if (isBusy) { return "ERR:BUSY"; }
        isBusy = true;

        var resultHolder = { value: null };
        try {
            var workerSource = "";
            for (var i = 0; i < WORKER_FUNCS.length; i++) { workerSource += WORKER_FUNCS[i].toString(); }
            workerSource += callExpression + ";";

            var bridgeMessage = new BridgeTalk();
            bridgeMessage.target = BridgeTalk.appSpecifier;
            bridgeMessage.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerSource) + "\"));";
            bridgeMessage.onResult = function (bridgeResult) {
                resultHolder.value = (bridgeResult && bridgeResult.body != null) ? String(bridgeResult.body) : "";
            };
            bridgeMessage.onError = function (bridgeError) {
                resultHolder.value = "ERR:" + ((bridgeError && bridgeError.body) ? bridgeError.body : "bridge");
            };
            bridgeMessage.send(BRIDGE_TIMEOUT_SECONDS);
        } catch (e) {
            resultHolder.value = "ERR:" + e;
        } finally {
            isBusy = false;
        }

        if (resultHolder.value === null) { return "ERR:TIMEOUT"; }
        return resultHolder.value;
    }

    /**
     * wkCollect() の戻り値（OK|key=value|...|MEMO|count|note...）を解析する
     * @param {string} response - 戻り値
     * @returns {{map: Object, memoList: string[]}|null} 集計値とメモ。形式が違えば null
     */
    function parseCollect(response) {
        if (!response || response.indexOf("OK|") !== 0) return null;
        var responseBody = response.substring(3);
        var memoMarkerIndex = responseBody.indexOf("|MEMO|");
        if (memoMarkerIndex < 0) return null;
        var statPart = responseBody.substring(0, memoMarkerIndex);
        var memoPart = responseBody.substring(memoMarkerIndex + 6);

        var statMap = {};
        var statPairs = statPart.split("|");
        for (var i = 0; i < statPairs.length; i++) {
            var separatorIndex = statPairs[i].indexOf("=");
            if (separatorIndex > 0) { statMap[statPairs[i].substring(0, separatorIndex)] = statPairs[i].substring(separatorIndex + 1); }
        }
        if (statMap.docName != null) { try { statMap.docName = decodeURIComponent(statMap.docName); } catch (e) {} }

        var memoFieldsText = memoPart.split("|");
        var memoCount = parseInt(memoFieldsText[0], 10) || 0;
        var memoList = [];
        for (var k = 1; k <= memoCount && k < memoFieldsText.length; k++) {
            var noteText = memoFieldsText[k];
            try { noteText = decodeURIComponent(noteText); } catch (e2) {}
            memoList.push(noteText);
        }
        return { map: statMap, memoList: memoList };
    }

    // =========================================
    // 状態保持（常駐エンジン） / Session state (resident engine)
    // =========================================

    if (!$.global.__selectionInspectorState) {
        $.global.__selectionInspectorState = { location: null };
    }

    // =========================================
    // レポート書き出し / Report export
    // =========================================

    /**
     * 集計結果をデスクトップのテキストファイルに書き出す
     * @param {{map: Object, memoList: string[]}} collected - parseCollect() の結果
     * @returns {string|null} 書き出したファイルのパス。ファイルを開けなければ null
     */
    function writeReportFile(collected) {
        var statMap = collected.map;
        var fullName = statMap.docName || "";
        var baseName = fullName.replace(/\.[^\.]+$/, "");
        var today = new Date();
        var yyyy = today.getFullYear();
        var mm = ("0" + (today.getMonth() + 1)).slice(-2);
        var dd = ("0" + today.getDate()).slice(-2);

        var reportPath = Folder.desktop + "/count-" + baseName + "-" + yyyy + mm + dd + ".txt";
        var reportFile = new File(reportPath);
        if (!reportFile.open("w")) return null;

        reportFile.writeln(getLabel('report.title'));
        reportFile.writeln(labelText('report.document') + " " + fullName);
        reportFile.writeln(labelText('report.date') + " " + yyyy + "-" + mm + "-" + dd);
        reportFile.writeln("");
        reportFile.writeln(getLabel('report.valueNote'));

        for (var sectionIndex = 0; sectionIndex < REPORT_SECTIONS.length; sectionIndex++) {
            var sectionKey = REPORT_SECTIONS[sectionIndex];
            reportFile.writeln("");
            reportFile.writeln(getLabel('panel.' + sectionKey));
            if (sectionKey === "memo") {
                var nonEmptyNotes = collectNonEmptyNotes(collected.memoList);
                if (nonEmptyNotes.length > 0) { reportFile.writeln(nonEmptyNotes.join("\n")); }
                continue;
            }
            var statRows = STAT_ROWS[sectionKey];
            for (var i = 0; i < statRows.length; i++) {
                reportFile.writeln(reportLabelText('fieldLabel.' + statRows[i].label) + " " + formatStatValue(statRows[i], statMap));
            }
        }

        reportFile.close();
        return reportPath;
    }

    // =========================================
    // パレット構築 / Build palette
    // =========================================

    /**
     * 項目名と値を1行追加する
     * @param {Panel} statPanel - 追加先のパネル
     * @param {Object} statRow - STAT_ROWS の1行
     * @param {number} labelWidth - 項目名の幅
     * @returns {StaticText} 値を表示する statictext
     */
    function addStatRow(statPanel, statRow, labelWidth) {
        var rowGroup = statPanel.add("group");
        rowGroup.orientation = "row";
        var rowLabel = rowGroup.add("statictext", undefined, labelText('fieldLabel.' + statRow.label));
        rowLabel.justify = "right";
        rowLabel.preferredSize.width = labelWidth;
        if (statRow.tip) rowLabel.helpTip = getLabel('tooltip.' + statRow.tip);
        var valueText = rowGroup.add("statictext", undefined, "-");
        valueText.preferredSize.width = VALUE_WIDTH;
        valueText.helpTip = getLabel('tooltip.statValue');
        return valueText;
    }

    /**
     * 見出し付きのパネルを追加する
     * @param {Group} parentColumn - 追加先のカラム
     * @param {string} titlePath - 見出しのラベルのパス
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentColumn, titlePath) {
        var statPanel = parentColumn.add("panel", undefined, getLabel(titlePath));
        setupPanel(statPanel);
        return statPanel;
    }

    /**
     * 情報タブ（2カラムの集計パネルとメモの表示欄）を組み立てる
     * @param {Group} infoTab - 情報タブのグループ
     * @returns {{valueTexts: Object, memoPreview: StaticText}} 行ラベル名 → 値の statictext と、メモの表示欄
     */
    function buildInfoTab(infoTab) {
        var twoColGroup = infoTab.add("group");
        twoColGroup.orientation = "row";
        twoColGroup.alignChildren = ["fill", "top"];
        twoColGroup.spacing = COLUMN_SPACING;

        var valueTexts = {};
        var memoPreview = null;
        for (var columnIndex = 0; columnIndex < PALETTE_COLUMNS.length; columnIndex++) {
            var columnGroup = twoColGroup.add("group");
            columnGroup.orientation = "column";
            columnGroup.alignChildren = ["fill", "top"];
            var labelWidth = (columnIndex === 0) ? LABEL_WIDTH_LEFT : LABEL_WIDTH_RIGHT;

            for (var panelIndex = 0; panelIndex < PALETTE_COLUMNS[columnIndex].length; panelIndex++) {
                var panelKey = PALETTE_COLUMNS[columnIndex][panelIndex];
                var statPanel = addPanel(columnGroup, 'panel.' + panelKey);
                if (panelKey === "memo") {
                    memoPreview = statPanel.add("statictext", undefined, "", { multiline: true });
                    memoPreview.preferredSize = MEMO_PREVIEW_SIZE;
                    continue;
                }
                var statRows = STAT_ROWS[panelKey];
                for (var i = 0; i < statRows.length; i++) {
                    valueTexts[statRows[i].label] = addStatRow(statPanel, statRows[i], labelWidth);
                }
            }
        }
        return { valueTexts: valueTexts, memoPreview: memoPreview };
    }

    /**
     * 集計値をすべての行に反映する
     * @param {Object} valueTexts - 行ラベル名 → 値の statictext
     * @param {Object} statMap - parseCollect() の map
     * @returns {void}
     */
    function applyStatValues(valueTexts, statMap) {
        for (var panelKey in STAT_ROWS) {
            if (!STAT_ROWS.hasOwnProperty(panelKey)) continue;
            var statRows = STAT_ROWS[panelKey];
            for (var i = 0; i < statRows.length; i++) {
                valueTexts[statRows[i].label].text = formatStatValue(statRows[i], statMap);
            }
        }
    }

    /**
     * すべての行の値を「-」に戻す
     * @param {Object} valueTexts - 行ラベル名 → 値の statictext
     * @returns {void}
     */
    function clearStatValues(valueTexts) {
        for (var rowLabel in valueTexts) {
            if (valueTexts.hasOwnProperty(rowLabel)) { valueTexts[rowLabel].text = "-"; }
        }
    }

    /**
     * パレットのキー操作を組み込む
     * Esc: 閉じる / ⌥I・⌥M: タブ切替 / ⌘R: 更新（Enter はメモ編集と衝突するため使わない）
     * @param {Window} inspectorPalette - パレット
     * @param {{showInfo: Function, showMemo: Function, refresh: Function}} keyActions - キーに割り当てる処理
     * @returns {void}
     */
    function bindPaletteKeys(inspectorPalette, keyActions) {
        /* メモ欄の編集中も効かせる（どれも修飾キー付きか Esc なので入力とぶつからない）
           / Active while editing a memo too: every key is Esc or carries a modifier */
        addKeyShortcuts(inspectorPalette, {
            "Escape": { target: function () { try { inspectorPalette.close(); } catch (e) {} }, inFields: true },
            "Alt+I": { target: function () { keyActions.showInfo(); }, inFields: true },
            "Alt+M": { target: function () { keyActions.showMemo(); }, inFields: true },
            "Cmd+R": { target: function () { keyActions.refresh(); }, inFields: true }
        });
    }

    /**
     * パレットを組み立てる
     * @returns {Window} 組み立てたパレット（未表示）
     */
    function buildPalette() {
        var inspectorPalette = new Window("palette", getLabel('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
        setupWindow(inspectorPalette);

        var memoFields = [];

        /* 表示切り替え（ラジオボタン＋stack）/ View switcher (radio buttons + stack) */
        var switchRow = inspectorPalette.add("group");
        switchRow.orientation = "row";
        switchRow.alignChildren = ["center", "center"];
        switchRow.alignment = ["center", "center"];
        switchRow.helpTip = getLabel('tooltip.shortcut');

        var infoRadio = switchRow.add("radiobutton", undefined, getLabel('radio.info'));
        var memoRadio = switchRow.add("radiobutton", undefined, getLabel('radio.memo'));
        infoRadio.helpTip = getLabel('tooltip.shortcut');
        memoRadio.helpTip = getLabel('tooltip.shortcut');
        infoRadio.value = true;

        var stackWrap = inspectorPalette.add("group");
        stackWrap.orientation = "stack";
        stackWrap.alignChildren = ["fill", "fill"];

        var infoTab = stackWrap.add("group");
        infoTab.orientation = "column";
        infoTab.alignChildren = ["fill", "top"];
        infoTab.margins = VIEW_MARGINS;

        var memoTab = stackWrap.add("group");
        memoTab.orientation = "column";
        memoTab.alignChildren = ["fill", "top"];
        memoTab.margins = VIEW_MARGINS;
        memoTab.visible = false;

        /* 情報タブ（2カラム） / Info tab (two columns) */
        var infoParts = buildInfoTab(infoTab);
        var valueTexts = infoParts.valueTexts;
        var memoPreview = infoParts.memoPreview;

        /* ステータス / Status line */
        var statusText = inspectorPalette.add("statictext", undefined, getLabel('status.ready'));
        statusText.alignment = ["fill", "bottom"];

        /**
         * ステータス行に表示する
         * @param {string} statusMessage - 表示する文言
         * @returns {void}
         */
        function setStatus(statusMessage) { statusText.text = statusMessage; }

        /**
         * タブの中身が変わったあとにレイアウトを組み直す
         * @returns {void}
         */
        function relayout() {
            /* 表示前など、layout を呼べない状態がある / layout may be unavailable, e.g. before the palette is shown */
            try { if (stackWrap.layout) { stackWrap.layout.layout(true); } } catch (e) {}
            try { if (inspectorPalette.layout) { inspectorPalette.layout.layout(true); } } catch (e) {}
        }

        /**
         * 情報タブとメモタブを切り替える
         * @param {string} viewMode - "info" または "memo"
         * @returns {void}
         */
        function switchView(viewMode) {
            var showInfo = (viewMode === "info");
            infoRadio.value = showInfo;
            memoRadio.value = !showInfo;
            infoTab.visible = showInfo;
            memoTab.visible = !showInfo;
            relayout();
        }

        /**
         * メモタブを開き、先頭の入力欄にフォーカスする
         * @returns {void}
         */
        function showMemoView() {
            switchView("memo");
            try { if (memoFields.length > 0) { memoFields[0].active = true; } } catch (e) {}
        }

        /**
         * 情報タブのメモ欄を更新する（1件ならその内容、複数なら案内）
         * @param {string[]} memoList - 選択オブジェクトのメモ
         * @returns {void}
         */
        function updateMemoPreview(memoList) {
            var nonEmptyNotes = collectNonEmptyNotes(memoList);
            memoPreview.text = (nonEmptyNotes.length === 1) ? nonEmptyNotes[0] : (nonEmptyNotes.length > 1 ? getLabel('memo.multiple') : "");
        }

        /**
         * メモタブを作り直す（選択オブジェクトごとに入力欄と［適用］ボタン）
         * @param {string[]} memoList - 選択オブジェクトのメモ（並べ替え済み）
         * @returns {void}
         */
        function rebuildMemo(memoList) {
            while (memoTab.children.length > 0) {
                try { memoTab.remove(memoTab.children[0]); } catch (e) { break; }
            }
            memoFields = [];

            if (!memoList || memoList.length === 0) {
                memoTab.add("statictext", undefined, getLabel('memo.none'));
                relayout();
                return;
            }

            for (var i = 0; i < memoList.length; i++) {
                (function (itemIndex, noteText) {
                    var memoRow = memoTab.add("group");
                    memoRow.orientation = "row";
                    memoRow.alignChildren = ["fill", "center"];
                    var memoField = memoRow.add("edittext", undefined, noteText, { multiline: true });
                    memoField.preferredSize = MEMO_FIELD_SIZE;
                    memoFields.push(memoField);
                    var btnApply = memoRow.add("button", undefined, getLabel('button.applyMemo'));
                    btnApply.helpTip = getLabel('tooltip.applyMemo');
                    btnApply.onClick = function () { applyMemo(itemIndex, memoField.text); };
                })(i, memoList[i]);
            }
            relayout();
        }

        /**
         * メインエンジンで集計し直し、パレットに反映する
         * @returns {void}
         */
        function refresh() {
            setStatus(getLabel('status.busy'));
            var response = callMainEngine("wkCollect(\"" + encodeURIComponent(getLabel('status.counting')) + "\")");

            if (response === "ERR:BUSY") { setStatus(getLabel('status.busy')); return; }
            if (response === null || response === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
            if (response === "NODOC") { setStatus(getLabel('status.noDoc')); clearStatValues(valueTexts); updateMemoPreview([]); rebuildMemo([]); return; }
            if (response.indexOf("ERR:") === 0) { setStatus(labelValueText('status.error', response.substring(4))); return; }

            var collected = parseCollect(response);
            if (!collected) { setStatus(getLabel('status.error')); return; }

            applyStatValues(valueTexts, collected.map);
            updateMemoPreview(collected.memoList);
            rebuildMemo(collected.memoList);

            var selectedCount = parseInt(collected.map.selCount, 10) || 0;
            if (selectedCount > 0) {
                setStatus(getLabel('status.selectedPrefix') + selectedCount + getLabel('status.selectedSuffix'));
            } else {
                setStatus(getLabel('status.wholeDoc'));
            }
        }

        /**
         * メモをオブジェクトに設定し、集計し直す
         * @param {number} itemIndex - 並べ替えた選択での番号
         * @param {string} noteText - 設定するメモ
         * @returns {void}
         */
        function applyMemo(itemIndex, noteText) {
            setStatus(getLabel('status.busy'));
            var response = callMainEngine("wkApplyMemo(" + itemIndex + ",\"" + encodeURIComponent(noteText) + "\")");
            if (response === "OK") { setStatus(getLabel('status.memoApplied')); refresh(); }
            else if (response === "NODOC") { setStatus(getLabel('status.noDoc')); }
            else if (response === "IDX") { setStatus(getLabel('status.selChanged')); }
            else { setStatus(labelValueText('status.error', response)); }
        }

        /**
         * 集計し直してレポートをデスクトップに書き出す（ファイルはパレット側で作る）
         * @returns {void}
         */
        function exportReport() {
            setStatus(getLabel('status.busy'));
            var response = callMainEngine("wkCollect(\"" + encodeURIComponent(getLabel('status.counting')) + "\")");
            if (response === "NODOC") { setStatus(getLabel('status.noDoc')); return; }
            if (response === null || response === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
            if (typeof response === "string" && response.indexOf("ERR:") === 0) { setStatus(labelValueText('status.error', response.substring(4))); return; }

            var collected = parseCollect(response);
            if (!collected) { setStatus(getLabel('status.error')); return; }

            try {
                var reportPath = writeReportFile(collected);
                if (reportPath !== null) {
                    setStatus(getLabel('status.exportedPrefix') + reportPath);
                } else {
                    setStatus(getLabel('status.exportFailOpen'));
                }
            } catch (err) {
                setStatus(labelValueText('status.error', err));
            }
        }

        /* ボタン行 / Button row */
        var btnRowGroup = inspectorPalette.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnExport = btnLeftGroup.add("button", undefined, getLabel('button.exportReport'));
        btnExport.helpTip = getLabel('tooltip.exportReport') + "\n" + getLabel('tooltip.esc');

        var spacer = btnRowGroup.add("statictext", undefined, "");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnRefresh = btnRightGroup.add("button", undefined, getLabel('button.refresh'));
        btnRefresh.helpTip = getLabel('tooltip.refresh') + "\n" + getLabel('tooltip.esc');

        btnExport.onClick = exportReport;
        btnRefresh.onClick = refresh;

        infoRadio.onClick = function () { switchView("info"); };
        memoRadio.onClick = showMemoView;

        bindPaletteKeys(inspectorPalette, {
            showInfo: function () { switchView("info"); },
            showMemo: showMemoView,
            refresh: refresh
        });

        try { inspectorPalette.opacity = PALETTE_OPACITY; } catch (e) {}

        switchView("info");

        /* 表示直後に一度集計 / Count once right after showing */
        inspectorPalette.onShow = function () {
            refresh();
        };

        return inspectorPalette;
    }

    // =========================================
    // 位置の記憶・復元 / Remember & restore location
    // =========================================

    /**
     * 前回の位置にパレットを置く（記録が無ければ中央）
     * @param {Window} inspectorPalette - パレット
     * @returns {void}
     */
    function restoreLocation(inspectorPalette) {
        try {
            var savedLocation = $.global.__selectionInspectorState.location;
            if (savedLocation && savedLocation.length === 2) {
                inspectorPalette.location = [savedLocation[0], savedLocation[1]];
            } else {
                inspectorPalette.center();
            }
        } catch (e) {
            inspectorPalette.center();
        }
    }

    /**
     * パレットの位置を常駐エンジンに控える
     * @param {Window} inspectorPalette - パレット
     * @returns {void}
     */
    function rememberLocation(inspectorPalette) {
        try {
            if (inspectorPalette.location && inspectorPalette.location.length === 2) {
                $.global.__selectionInspectorState.location = [inspectorPalette.location[0], inspectorPalette.location[1]];
            }
        } catch (e) {}
    }

    // =========================================
    // 起動 / Entry point
    // =========================================

    /**
     * パレットを表示する（開いているものがあれば閉じてから）
     * @returns {void}
     */
    function showPalette() {
        /* 多重起動防止：既存パレットがあれば閉じる / Prevent duplicates */
        if ($.global.__SelectionInspectorPalette) {
            try { $.global.__SelectionInspectorPalette.close(); } catch (e) {}
            $.global.__SelectionInspectorPalette = null;
        }

        var inspectorPalette = buildPalette();

        /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
        $.global.__SelectionInspectorPalette = inspectorPalette;
        inspectorPalette.onClose = function () {
            rememberLocation(inspectorPalette);
            app.redraw();
            $.global.__SelectionInspectorPalette = null;
        };

        restoreLocation(inspectorPalette);
        inspectorPalette.show();
    }

    showPalette();

})();
