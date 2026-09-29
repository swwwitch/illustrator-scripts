#target illustrator
#targetengine "SmartDrawArtboardRectangleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブまたはすべてのアートボードと同じサイズの長方形を、オフセットを考慮して描画します。
カラー・配置位置・対象をライブプレビューで確かめながら指定でき、描画後に「ガイドに変換」「ライブシェイプ化」を適用できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDrawArtboardRectangle.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1ba88513a9c8

### Overview

Draws a rectangle the size of the active artboard, or of every artboard, taking an offset into account.
Color, placement and target scope are set with a live preview, and the result can be converted to guides or to a live shape.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDrawArtboardRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartDrawArtboardRectangle";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDrawArtboardRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDrawArtboardRectangle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1ba88513a9c8"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 入力中のプレビュー遅延（タイプしやすさ優先）/ Preview delay while typing */
    var PREVIEW_DELAY_TYPING_MS = 110; /* 推奨 100–120ms / recommend 100–120ms */

    /* K100モードの不透明度（%）/ Opacity (%) used by the K100 color mode */
    var K100_OPACITY = 15;

    /* 「bgレイヤー」配置で使うレイヤー名 / Layer name used by the "bg layer" placement */
    var BG_LAYER_NAME = 'bg';

    /* 描画先が見つからないときに作るレイヤー名 / Layer created when no writable layer is found */
    var FALLBACK_LAYER_NAME = '_auto_draw';

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* 余白と間隔 / Margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING = 8;                /* パネル内の要素間隔 */
    var COLUMN_SPACING = 12;              /* 2カラムの間隔 */
    var STACK_SPACING = 10;               /* カラム内のパネル間隔・広めの行間 */
    var TIGHT_SPACING = 6;                /* 詰めた行間 */

    /* カスタムの色見本の大きさ / size of the Custom swatch */
    var CUSTOM_SWATCH_SIZE = [36, 18];

    /**
     * パネルの共通設定
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループの共通設定（row/column で整列を切り替え）
     * @param {Group} group - 対象グループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // カラーモード / Color modes
    // =========================================

    /* 塗りの決め方を表す定数 / How the fill color is decided */
    var ColorMode = {
        NONE: 'none',
        K100: 'k100',
        CUSTOM: 'custom'
    };

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

    /* ラベル定義（カテゴリ別）/ Label definitions (by category) */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードサイズの長方形を描画", en: "Draw Artboard Size Rectangle" }
        },
        panel: {
            offset: { ja: "オフセット", en: "Offset" },
            color: { ja: "カラー", en: "Color" },
            placement: { ja: "配置位置", en: "Placement" },
            target: { ja: "対象", en: "Target" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            colorNone: { ja: "なし", en: "None" },
            colorK100: { ja: "K100、不透明度15%", en: "K100, Opacity 15%" },
            colorCustom: { ja: "カスタム", en: "Custom" },
            placeFront: { ja: "最前面", en: "Front" },
            placeBack: { ja: "最背面", en: "Back" },
            placeBgLayer: { ja: "bgレイヤー", en: "bg Layer" },
            currentArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All Artboards" }
        },
        checkbox: {
            bleed: { ja: "裁ち落とし", en: "Bleed" },
            makeGuide: { ja: "ガイドに変換", en: "Convert to Guides" },
            convertToLiveShape: { ja: "ライブシェイプ化", en: "Convert to Live Shape" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            previewOutline: { ja: "アウトライン表示", en: "Outline" },
            previewPreview: { ja: "プレビュー表示", en: "Preview" }
        },
        tooltip: {
            offsetInput: {
                ja: "アートボード境界から外側へ広げる量を指定します。負の値で内側へ縮めます。",
                en: "Set how far the bounds expand outward from the artboard. Use a negative value to shrink inward."
            },
            bleed: {
                ja: "現在の単位に応じて、裁ち落とし相当の値を自動入力します。",
                en: "Automatically fills a bleed-equivalent offset based on the current unit."
            },
            colorNone: { ja: "塗りも線もない長方形を描画します。", en: "Draws the rectangle with no fill and no stroke." },
            colorCustom: {
                ja: "色見本をクリックすると、Illustrator 標準のカラーピッカーで塗りの色を選べます。",
                en: "Click the swatch to choose the fill color in Illustrator's standard color picker."
            },
            placeFront: { ja: "現在のレイヤー内で最前面に配置します。", en: "Places the rectangle at the front of the current layer." },
            placeBack: { ja: "現在のレイヤー内で最背面に配置します。", en: "Places the rectangle at the back of the current layer." },
            bgLayer: {
                ja: "bgレイヤーを作成または使用し、レイヤーの最背面へ配置します。",
                en: "Creates or uses the bg layer and places it at the back of the layer stack."
            },
            previewToggle: {
                ja: "Illustratorのアウトライン表示／プレビュー表示を切り替えます。",
                en: "Toggles Illustrator's Outline and Preview display modes."
            },
            convertToLiveShape: {
                ja: "Illustratorのメニューコマンドで長方形をライブシェイプ化します。中心点も表示されます。",
                en: "Uses Illustrator's menu command to convert rectangles to Live Shapes. The center point is also shown."
            },
            makeGuide: { ja: "描画した長方形をガイドに変換します。", en: "Converts the drawn rectangles to guides." },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        warning: {
            singleArtboard: { ja: "アートボードが1つのため選択できません", en: "Disabled: only one artboard exists" }
        },
        objectName: {
            previewLayer: { ja: "_preview", en: "_preview" },
            rect: { ja: "<長方形>", en: "<Rectangle>" },
            guide: { ja: "<ガイド>", en: "<Guide>" },
            previewRect: { ja: "__プレビュー_アートボードサイズの長方形", en: "__Preview_ArtboardSizeRectangle" }
        }
    };

    // =========================================
    // ダイアログ共通ユーティリティ / Dialog utilities
    // =========================================

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

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 入力欄の表示値と内部pt値を一元的に解決する（裁ち落としプリセットを含む）
     * @param {string} offsetText - 入力欄の現在のテキスト
     * @param {number} unitCode - rulerType の単位コード
     * @param {boolean} bleedEnabled - 裁ち落としがONかどうか
     * @returns {object} { pt: number, displayText: string, disabled: boolean }
     */
    function resolveOffsetToPt(offsetText, unitCode, bleedEnabled) {
        var displayText = String(offsetText == null ? '' : offsetText);

        if (bleedEnabled) {
            /* 単位ごとの裁ち落とし相当値。表示値と pt 値を必ず同じ量にする
               Bleed preset per unit; the shown value and the pt value always describe the same amount */
            var bleedAmount = 3;      /* 既定は 3mm 相当 / defaults to 3mm */
            var bleedUnitCode = 1;
            if (unitCode === 5) {          /* Q/H */
                bleedAmount = 12;
                bleedUnitCode = 5;
            } else if (unitCode === 2) {   /* pt（0.125in = 9pt）*/
                bleedAmount = 9;
                bleedUnitCode = 2;
            } else if (unitCode === 1) {   /* mm */
                bleedAmount = 3;
                bleedUnitCode = 1;
            } else {
                /* mm・Q/H・pt 以外は 3mm 相当を現在の単位へ換算して表示
                   For other units, convert the 3mm equivalent into the current unit */
                bleedAmount = 3 * UNITS[1].pointsPerUnit / (UNITS[unitCode] ? UNITS[unitCode].pointsPerUnit : 1);
                bleedAmount = Math.round(bleedAmount * 1000) / 1000;
                bleedUnitCode = unitCode;
            }
            return {
                pt: bleedAmount * (UNITS[bleedUnitCode] ? UNITS[bleedUnitCode].pointsPerUnit : 1),
                displayText: String(bleedAmount),
                disabled: true
            };
        }

        /* 通常時は現在の単位の係数を掛ける / Normal case: multiply by the current unit factor */
        var offsetValue = parseFloat(displayText);
        if (isNaN(offsetValue)) offsetValue = 0;
        return {
            pt: offsetValue * (UNITS[unitCode] ? UNITS[unitCode].pointsPerUnit : 1),
            displayText: displayText,
            disabled: false
        };
    }

    // =========================================
    // カラー / Color
    // =========================================

    /**
     * 数値を指定範囲に収める
     * @param {number} value - 対象の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clampValue(value, minValue, maxValue) {
        return value < minValue ? minValue : (value > maxValue ? maxValue : value);
    }

    /**
     * RGBColor を生成する（0–255にクランプ）
     * @param {number} red - 赤（0–255）
     * @param {number} green - 緑（0–255）
     * @param {number} blue - 青（0–255）
     * @returns {RGBColor} 生成した色
     */
    function makeRgbColor(red, green, blue) {
        var rgbColor = new RGBColor();
        rgbColor.red = clampValue(Math.round(red), 0, 255);
        rgbColor.green = clampValue(Math.round(green), 0, 255);
        rgbColor.blue = clampValue(Math.round(blue), 0, 255);
        return rgbColor;
    }

    /**
     * CMYKColor を生成する（0–100にクランプ）
     * @param {number} cyan - シアン（0–100）
     * @param {number} magenta - マゼンタ（0–100）
     * @param {number} yellow - イエロー（0–100）
     * @param {number} black - ブラック（0–100）
     * @returns {CMYKColor} 生成した色
     */
    function makeCmykColor(cyan, magenta, yellow, black) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = clampValue(cyan, 0, 100);
        cmykColor.magenta = clampValue(magenta, 0, 100);
        cmykColor.yellow = clampValue(yellow, 0, 100);
        cmykColor.black = clampValue(black, 0, 100);
        return cmykColor;
    }

    /**
     * CMYK値をRGB値へ変換する
     * @param {number} cyan - シアン（0–100）
     * @param {number} magenta - マゼンタ（0–100）
     * @param {number} yellow - イエロー（0–100）
     * @param {number} black - ブラック（0–100）
     * @returns {number[]} [R, G, B]（0–255）
     */
    function cmykToRgb(cyan, magenta, yellow, black) {
        var cyanRatio = clampValue(cyan, 0, 100) / 100;
        var magentaRatio = clampValue(magenta, 0, 100) / 100;
        var yellowRatio = clampValue(yellow, 0, 100) / 100;
        var blackRatio = clampValue(black, 0, 100) / 100;
        return [
            Math.round(255 * (1 - cyanRatio) * (1 - blackRatio)),
            Math.round(255 * (1 - magentaRatio) * (1 - blackRatio)),
            Math.round(255 * (1 - yellowRatio) * (1 - blackRatio))
        ];
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を生成する
     * @param {Document} doc - 対象ドキュメント
     * @returns {RGBColor|CMYKColor} 黒
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) return makeRgbColor(0, 0, 0);
        return makeCmykColor(0, 0, 0, 100);
    }

    /**
     * カラーモードに応じた塗りを適用する（プレビューと本描画で共通）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} targetRectangle - 塗りを適用する長方形
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function applyFillByMode(doc, targetRectangle, drawSettings) {
        var fillColor = null;
        var fillOpacity = 100;

        if (drawSettings.colorMode === ColorMode.K100) {
            fillColor = createBlackColor(doc);
            fillOpacity = K100_OPACITY;
        } else if (drawSettings.colorMode === ColorMode.CUSTOM) {
            fillColor = drawSettings.customColor;
        }

        /* 「なし」は塗りなし。線は呼び出し側（プレビュー）で付け直す
           Unparsable values and "None" mean no fill; the caller re-applies any stroke */
        targetRectangle.stroked = false;
        targetRectangle.filled = !!fillColor;
        if (fillColor) targetRectangle.fillColor = fillColor;
        targetRectangle.opacity = fillOpacity;
    }

    // =========================================
    // 長方形とレイヤーの共通処理 / Rectangle and layer helpers
    // =========================================

    /**
     * 座標系をドキュメント座標へ切り替える
     * @returns {CoordinateSystem|null} 切り替え前の座標系（取得できなければ null）
     */
    function switchToDocumentCoordinates() {
        var previousCoordinateSystem = null;
        try {
            previousCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }
        return previousCoordinateSystem;
    }

    /**
     * switchToDocumentCoordinates() で控えた座標系へ戻す
     * @param {CoordinateSystem|null} previousCoordinateSystem - 切り替え前の座標系
     * @returns {void}
     */
    function restoreCoordinateSystem(previousCoordinateSystem) {
        try {
            if (previousCoordinateSystem !== null) app.coordinateSystem = previousCoordinateSystem;
        } catch (e) { }
    }

    /**
     * アートボードをオフセットぶん広げた長方形の位置と寸法を求める
     * @param {number[]} artboardRect - アートボードの [left, top, right, bottom]
     * @param {number} offsetPt - 外側へ広げる量（pt、負の値で内側）
     * @returns {{top: number, left: number, width: number, height: number}} 長方形の位置と寸法
     */
    function getOffsetRectangleBounds(artboardRect, offsetPt) {
        return {
            top: artboardRect[1] + offsetPt,
            left: artboardRect[0] - offsetPt,
            width: (artboardRect[2] - artboardRect[0]) + offsetPt * 2,
            height: (artboardRect[1] - artboardRect[3]) + offsetPt * 2
        };
    }

    /**
     * 配置位置の設定どおりにレイヤー内の重ね順を変える（bgレイヤーは作成順のまま）
     * @param {PathItem} targetRectangle - 対象の長方形
     * @param {string} zOrder - "front" / "back" / "bg"
     * @returns {void}
     */
    function applyZOrder(targetRectangle, zOrder) {
        if (zOrder === 'front') targetRectangle.zOrder(ZOrderMethod.BRINGTOFRONT);
        else if (zOrder === 'back') targetRectangle.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 名前が一致する最上位レイヤーを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー。無ければ null
     */
    function findLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return doc.layers[i];
        }
        return null;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 使い方:
     *   PreviewHistory.start();      // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();       // プレビュー描画ごとにカウント(+1)
     *   PreviewHistory.undo();       // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */
    (function (globalObject) {
        if (!globalObject.PreviewHistory) {
            globalObject.PreviewHistory = {
                start: function () {
                    globalObject.__previewUndoCount = 0;
                },
                bump: function () {
                    globalObject.__previewUndoCount = (globalObject.__previewUndoCount | 0) + 1;
                },
                undo: function () {
                    var undoCount = globalObject.__previewUndoCount | 0;
                    try {
                        for (var i = 0; i < undoCount; i++) app.executeMenuCommand('undo');
                    } catch (e) { }
                    globalObject.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

    /* デバウンス中のプレビュータスクID / Task id of the pending debounced preview */
    var previewDebounceTaskId = null;

    /**
     * プレビュー描画を遅延スケジュールする（デバウンス）
     * @param {object} drawSettings - ダイアログの設定値
     * @param {number} delayMs - 遅延ミリ秒
     * @returns {void}
     */
    function schedulePreview(drawSettings, delayMs) {
        PreviewHistory.cancelTask(previewDebounceTaskId);
        /* scheduleTask の文字列はグローバルスコープで評価されるため、IIFE内の関数を $.global 経由で渡す
           scheduleTask runs its string in global scope, so expose the renderer via $.global */
        $.global.__previewSettings = drawSettings;
        $.global.__previewRenderer = renderPreview;
        var scheduledCode = 'try{$.global.__previewRenderer(app.activeDocument, $.global.__previewSettings);}catch(e){}';
        try {
            previewDebounceTaskId = app.scheduleTask(scheduledCode, Math.max(0, delayMs | 0), false);
        } catch (e) {
            try { renderPreview(app.activeDocument, drawSettings); } catch (err) { }
        }
    }

    /**
     * プレビュー専用レイヤーの名前かどうかを判定する
     * @param {string} layerName - レイヤー名
     * @returns {boolean} プレビュー専用レイヤーなら true
     */
    function isPreviewLayerName(layerName) {
        return layerName === getLabel('objectName.previewLayer') || layerName === '_preview';
    }

    /**
     * プレビューを片付ける
     * @param {boolean} removeLayer - true でレイヤーごと削除、false ではプレビュー長方形を隠すだけ
     * @returns {void}
     */
    function clearPreview(removeLayer) {
        try {
            var doc = app.activeDocument;
            var previewItemPrefix = getLabel('objectName.previewRect') + "#";
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var layer = doc.layers[i];
                if (!isPreviewLayerName(layer.name)) continue;
                if (removeLayer) {
                    layer.remove();
                    continue;
                }
                /* 入力中は削除せず、このスクリプトが作った長方形だけ隠す
                   While typing, hide only the items this script created instead of deleting them */
                for (var k = layer.pathItems.length - 1; k >= 0; k--) {
                    if (String(layer.pathItems[k].name || "").indexOf(previewItemPrefix) === 0) {
                        layer.pathItems[k].hidden = true;
                    }
                }
            }
        } catch (e) { }
    }

    /**
     * プレビュー専用レイヤーを取得する（なければ作成し最前面へ）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} プレビュー専用レイヤー
     */
    function getOrCreatePreviewLayer(doc) {
        var previewLayerName = getLabel('objectName.previewLayer');
        var previewLayer = findLayerByName(doc, previewLayerName);
        if (!previewLayer) {
            previewLayer = doc.layers.add();
            previewLayer.name = previewLayerName;
        }
        previewLayer.visible = true;
        previewLayer.locked = false;
        try {
            previewLayer.move(doc, ElementPlacement.PLACEATBEGINNING);
        } catch (e) { }
        return previewLayer;
    }

    /**
     * アートボード番号に対応するプレビュー長方形を取得する（なければ作成、あれば再利用）
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {number} artboardIndex - アートボード番号
     * @param {{top: number, left: number, width: number, height: number}} rectangleBounds - 長方形の位置と寸法
     * @returns {PathItem} プレビュー長方形
     */
    function getOrCreatePreviewRectangle(previewLayer, artboardIndex, rectangleBounds) {
        var previewItemName = getLabel('objectName.previewRect') + "#" + artboardIndex;
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var existingRectangle = previewLayer.pathItems[i];
            if (existingRectangle.name !== previewItemName) continue;
            /* 既存を使い回して再作成のヒストリーを増やさない / Reuse in place to avoid extra history entries */
            existingRectangle.top = rectangleBounds.top;
            existingRectangle.left = rectangleBounds.left;
            existingRectangle.width = rectangleBounds.width;
            existingRectangle.height = rectangleBounds.height;
            existingRectangle.hidden = false;
            return existingRectangle;
        }
        var previewRectangle = previewLayer.pathItems.rectangle(
            rectangleBounds.top, rectangleBounds.left, rectangleBounds.width, rectangleBounds.height);
        previewRectangle.name = previewItemName;
        return previewRectangle;
    }

    /**
     * プレビュー用の50%グレー（ドキュメントのカラースペースに合わせる）
     * @param {Document} doc - 対象ドキュメント
     * @returns {RGBColor|CMYKColor} 線色
     */
    function getPreviewStrokeColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) return makeRgbColor(128, 128, 128);
        return makeCmykColor(0, 0, 0, 50);
    }

    /**
     * 1つのアートボードぶんのプレビュー長方形を描く
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {number} artboardIndex - アートボード番号
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function drawPreviewRectangle(doc, previewLayer, artboardIndex, drawSettings) {
        var rectangleBounds = getOffsetRectangleBounds(doc.artboards[artboardIndex].artboardRect, drawSettings.offset || 0);
        var previewRectangle = getOrCreatePreviewRectangle(previewLayer, artboardIndex, rectangleBounds);

        applyFillByMode(doc, previewRectangle, drawSettings);

        /* 塗りなしでも位置が分かるよう、プレビューは常に破線で縁取る
           Always outline the preview so it stays visible even with no fill */
        previewRectangle.stroked = true;
        previewRectangle.strokeWidth = 1;
        previewRectangle.strokeDashes = [6, 4];
        previewRectangle.strokeColor = getPreviewStrokeColor(doc);
        previewRectangle.selected = false;

        applyZOrder(previewRectangle, drawSettings.zOrder);
    }

    /**
     * 対象外のアートボードに対応するプレビュー長方形を隠す
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function hidePreviewItemsOutOfScope(doc, previewLayer, drawSettings) {
        /* すべてのアートボードが対象のときは -1（範囲外の番号だけ隠す）/ -1 means "every artboard is in scope" */
        var visibleArtboardIndex = (drawSettings.target === 'all') ? -1 : doc.artboards.getActiveArtboardIndex();
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var previewItem = previewLayer.pathItems[i];
            /* アイテム名は "__Preview_...#<アートボード番号>" 形式 */
            var indexMatch = /#(\d+)$/.exec(previewItem.name || "");
            if (!indexMatch) continue;
            var itemArtboardIndex = parseInt(indexMatch[1], 10);
            var keepVisible = (visibleArtboardIndex < 0) ?
                (itemArtboardIndex < doc.artboards.length) :
                (itemArtboardIndex === visibleArtboardIndex);
            if (!keepVisible) previewItem.hidden = true;
        }
    }

    /**
     * プレビューを描画する（専用レイヤーへ一時オブジェクトを生成）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function renderPreview(doc, drawSettings) {
        /* レイヤーごと消さずに既存プレビューを隠す / Hide existing preview items instead of deleting the layer */
        clearPreview(false);
        if (!doc || !drawSettings) return;

        var previousCoordinateSystem = switchToDocumentCoordinates();
        var previewLayer = getOrCreatePreviewLayer(doc);
        if (drawSettings.target === 'all') {
            for (var i = 0; i < doc.artboards.length; i++) drawPreviewRectangle(doc, previewLayer, i, drawSettings);
        } else {
            drawPreviewRectangle(doc, previewLayer, doc.artboards.getActiveArtboardIndex(), drawSettings);
        }
        hidePreviewItemsOutOfScope(doc, previewLayer, drawSettings);
        restoreCoordinateSystem(previousCoordinateSystem);

        PreviewHistory.bump();
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * オフセットパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} { offsetInput, bleedCheckbox, initFieldState }
     */
    function buildOffsetPanel(parentGroup, previewHooks) {
        var offsetPanel = parentGroup.add('panel', undefined, getLabel('panel.offset'));
        setupPanel(offsetPanel);

        var offsetRow = offsetPanel.add('group');
        setupGroup(offsetRow, 'row');
        offsetRow.alignChildren = 'center';
        offsetRow.alignment = 'center';

        /* ∧∨と入力欄は隙間0で突き合わせる。負の値（内側へ縮小）も可 / butt the stepper against the field; negatives allowed */
        var offsetFieldGroup = offsetRow.add('group');
        offsetFieldGroup.orientation = 'row';
        offsetFieldGroup.alignChildren = ['left', 'center'];
        offsetFieldGroup.spacing = 0;
        offsetFieldGroup.margins = 0;
        var offsetInput;
        var offsetStepper = addStepper(offsetFieldGroup, function () { return offsetInput; }, {
            onStep: function () { previewHooks.deferred(); }
        });
        offsetInput = offsetFieldGroup.add('edittext', undefined, '0');
        offsetInput.characters = 4;
        offsetInput.helpTip = getLabel('tooltip.offsetInput');
        offsetRow.add('statictext', undefined, getUnitInfo().label);

        var bleedRow = offsetPanel.add('group');
        setupGroup(bleedRow, 'row');
        bleedRow.alignChildren = 'center';
        bleedRow.alignment = 'center';

        var bleedCheckbox = bleedRow.add('checkbox', undefined, getLabel('checkbox.bleed'));
        bleedCheckbox.alignment = 'center';
        bleedCheckbox.value = false; /* デフォルトOFF / default OFF */
        bleedCheckbox.helpTip = getLabel('tooltip.bleed');

        /* 裁ち落としON/OFFの往復で手入力値を失わないよう控えておく
           Remember the manual offset so toggling Bleed does not lose it */
        var manualOffsetText = '0';

        /* 裁ち落としの状態を入力欄へ反映 / Reflect the current Bleed state in the field */
        function applyBleedState(refreshPreview) {
            if (bleedCheckbox.value) {
                var resolvedOffset = resolveOffsetToPt(offsetInput.text, getUnitInfo().code, true);
                offsetInput.text = resolvedOffset.displayText;
                offsetInput.enabled = !resolvedOffset.disabled;
            } else {
                offsetInput.text = manualOffsetText;
                offsetInput.enabled = true;
            }
            /* ∧∨も入力欄にそろえて切り替え、描き直す / sync and redraw the stepper */
            offsetStepper.enabled = offsetInput.enabled;
            redrawSteppersIn(offsetStepper);
            if (refreshPreview) previewHooks.deferred();
        }

        bleedCheckbox.onClick = function () {
            if (bleedCheckbox.value) manualOffsetText = String(offsetInput.text);
            applyBleedState(true);
        };

        offsetInput.onChanging = previewHooks.deferred;
        offsetInput.onChange = previewHooks.immediate;
        offsetInput.addEventListener('keydown', function (event) {
            /* Enterでも即座に反映 / Enter refreshes the preview immediately */
            if (event.keyName == 'Enter') previewHooks.immediate();
        });
        bindSteppedArrowKeys(offsetInput, offsetStepper);

        return {
            offsetInput: offsetInput,
            bleedCheckbox: bleedCheckbox,
            initFieldState: function () {
                manualOffsetText = String(offsetInput.text);
                applyBleedState(false);
            }
        };
    }

    /**
     * 色を色見本に描くための RGB（0〜1）にする
     * @param {RGBColor|CMYKColor|GrayColor} fillColor - 色
     * @returns {number[]} [r, g, b]（0〜1）
     */
    function colorToScreenRgb(fillColor) {
        if (fillColor.typename === "RGBColor") return [fillColor.red / 255, fillColor.green / 255, fillColor.blue / 255];
        if (fillColor.typename === "GrayColor") return [1 - fillColor.gray / 100, 1 - fillColor.gray / 100, 1 - fillColor.gray / 100];
        var rgbValues = cmykToRgb(fillColor.cyan, fillColor.magenta, fillColor.yellow, fillColor.black);
        return [rgbValues[0] / 255, rgbValues[1] / 255, rgbValues[2] / 255];
    }

    /**
     * ドキュメントのカラースペースに合わせた中間のグレーを作る（カスタムの初期色）
     * @param {Document} doc - 対象ドキュメント
     * @returns {RGBColor|CMYKColor} グレー
     */
    function createDefaultCustomColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) return makeRgbColor(128, 128, 128);
        return makeCmykColor(0, 0, 0, 50);
    }

    /**
     * カラーパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} 各ラジオと、カスタムの色を返す getCustomColor をまとめたオブジェクト
     */
    function buildColorPanel(parentGroup, previewHooks) {
        var colorPanel = parentGroup.add('panel', undefined, getLabel('panel.color'));
        setupPanel(colorPanel, STACK_SPACING); /* やや広めの行間 / a bit more vertical gap */

        var noneRadio = colorPanel.add('radiobutton', undefined, getLabel('radio.colorNone'));
        noneRadio.helpTip = getLabel('tooltip.colorNone');
        var k100Radio = colorPanel.add('radiobutton', undefined, getLabel('radio.colorK100'));

        /* カスタムはラジオと色見本を同じ行に / the Custom radio and its swatch share one row */
        var customRow = colorPanel.add('group');
        setupGroup(customRow, 'row', TIGHT_SPACING);
        var customRadio = customRow.add('radiobutton', undefined, getLabel('radio.colorCustom'));
        customRadio.helpTip = getLabel('tooltip.colorCustom');
        var customSwatch = customRow.add('group');
        customSwatch.preferredSize = CUSTOM_SWATCH_SIZE;
        customSwatch.helpTip = getLabel('tooltip.colorCustom');

        var customColor = createDefaultCustomColor(app.activeDocument);
        customSwatch.onDraw = function () {
            var swatchGraphics = customSwatch.graphics;
            var swatchWidth = customSwatch.size[0];
            var swatchHeight = customSwatch.size[1];
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, colorToScreenRgb(customColor)));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1));
        };

        /* 色見本をクリックしたら標準のカラーピッカーを開き、カスタムを選ぶ。
           OK なら選んだ色（ドキュメントのカラーモードの型）、キャンセルなら渡した色がそのまま返る
           Clicking the swatch opens the standard color picker and selects Custom */
        customSwatch.addEventListener('click', function () {
            customColor = app.showColorPicker(customColor);
            customSwatch.hide(); /* group には notify() が無いので、隠して再表示して描き直す / redraw the group */
            customSwatch.show();
            selectColorMode(ColorMode.CUSTOM);
        });

        /* カラーモードを排他選択してプレビューを更新 / Select a color mode exclusively, then refresh the preview */
        function selectColorMode(colorMode) {
            noneRadio.value = (colorMode === ColorMode.NONE);
            k100Radio.value = (colorMode === ColorMode.K100);
            customRadio.value = (colorMode === ColorMode.CUSTOM);
            previewHooks.immediate();
        }

        var colorRadioModes = [
            [noneRadio, ColorMode.NONE],
            [k100Radio, ColorMode.K100],
            [customRadio, ColorMode.CUSTOM]
        ];
        for (var j = 0; j < colorRadioModes.length; j++) {
            (function (radio, colorMode) {
                radio.onClick = radio.onChanging = function () { selectColorMode(colorMode); };
            })(colorRadioModes[j][0], colorRadioModes[j][1]);
        }

        k100Radio.value = true; /* デフォルトはK100 / default to K100 */

        return {
            noneRadio: noneRadio,
            k100Radio: k100Radio,
            customRadio: customRadio,
            getCustomColor: function () { return customColor; }
        };
    }

    /**
     * 配置位置（重ね順）パネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} { frontRadio, backRadio, bgLayerRadio }
     */
    function buildPlacementPanel(parentGroup, previewHooks) {
        var placementPanel = parentGroup.add('panel', undefined, getLabel('panel.placement'));
        setupPanel(placementPanel, TIGHT_SPACING);

        var frontRadio = placementPanel.add('radiobutton', undefined, getLabel('radio.placeFront'));
        var backRadio = placementPanel.add('radiobutton', undefined, getLabel('radio.placeBack'));
        var bgLayerRadio = placementPanel.add('radiobutton', undefined, getLabel('radio.placeBgLayer'));
        frontRadio.helpTip = getLabel('tooltip.placeFront');
        backRadio.helpTip = getLabel('tooltip.placeBack');
        bgLayerRadio.helpTip = getLabel('tooltip.bgLayer');

        frontRadio.value = true; /* デフォルトは最前面 / default to Bring to Front */

        var placementRadios = [frontRadio, backRadio, bgLayerRadio];
        for (var i = 0; i < placementRadios.length; i++) {
            placementRadios[i].onClick = previewHooks.immediate;
        }

        return { frontRadio: frontRadio, backRadio: backRadio, bgLayerRadio: bgLayerRadio };
    }

    /**
     * 対象アートボードのパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} { currentArtboardRadio, allArtboardsRadio }
     */
    function buildTargetPanel(parentGroup, previewHooks) {
        var targetPanel = parentGroup.add('panel', undefined, getLabel('panel.target'));
        setupPanel(targetPanel);

        var currentArtboardRadio = targetPanel.add('radiobutton', undefined, getLabel('radio.currentArtboard'));
        var allArtboardsRadio = targetPanel.add('radiobutton', undefined, getLabel('radio.allArtboards'));

        /* 常に「現在のアートボード」をデフォルト選択 / Always default to the current artboard */
        currentArtboardRadio.value = true;
        allArtboardsRadio.value = false;

        /* 1枚しかない場合は「すべてのアートボード」をディム / Dim "All Artboards" when there is only one */
        var artboardCount = app.documents.length ? app.activeDocument.artboards.length : 0;
        if (artboardCount <= 1) {
            allArtboardsRadio.enabled = false;
            allArtboardsRadio.helpTip = getLabel('warning.singleArtboard');
        }

        /* クリックだけでなくキーボード操作（onChanging）でも更新 / Refresh on click and on keyboard change */
        currentArtboardRadio.onClick = currentArtboardRadio.onChanging = previewHooks.immediate;
        allArtboardsRadio.onClick = allArtboardsRadio.onChanging = previewHooks.immediate;

        return { currentArtboardRadio: currentArtboardRadio, allArtboardsRadio: allArtboardsRadio };
    }

    /**
     * オプションパネル（ガイド化／ライブシェイプ変換）を構築する
     * 各オプションは独立。必要なものだけ描画後に適用（applyDrawOptions）
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @returns {object} { makeGuideCheckbox, convertToLiveShapeCheckbox }
     */
    function buildOptionsPanel(parentGroup) {
        var optionsPanel = parentGroup.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel);

        var makeGuideCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.makeGuide'));
        makeGuideCheckbox.value = false; /* デフォルトOFF / default OFF */
        makeGuideCheckbox.helpTip = getLabel('tooltip.makeGuide');

        var convertToLiveShapeCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.convertToLiveShape'));
        convertToLiveShapeCheckbox.value = true; /* デフォルトON / default ON */
        convertToLiveShapeCheckbox.helpTip = getLabel('tooltip.convertToLiveShape');

        return { makeGuideCheckbox: makeGuideCheckbox, convertToLiveShapeCheckbox: convertToLiveShapeCheckbox };
    }

    /**
     * ダイアログのショートカットキーを登録する（F/B/L=重ね順、C/A=対象、G=ガイド化）
     * 数値の欄でも効かせる
     * @param {Window} settingsDialog - 対象ダイアログ
     * @param {object} dialogControls - 各パネルのコントロール
     * @returns {void}
     */
    function addDialogShortcutKeys(settingsDialog, dialogControls) {
        var numericFields = [dialogControls.offset.offsetInput];
        addKeyShortcuts(settingsDialog, {
            /* ガイド化は描画後の処理なのでプレビューには反映しない / Make-guides is a post-draw option and is not previewed */
            "G": dialogControls.options.makeGuideCheckbox,
            "F": dialogControls.placement.frontRadio,
            "B": dialogControls.placement.backRadio,
            "L": dialogControls.placement.bgLayerRadio,
            "C": dialogControls.target.currentArtboardRadio,
            "A": dialogControls.target.allArtboardsRadio
        }, { numericFields: numericFields });
    }

    /**
     * ダイアログの入力内容から描画設定を組み立てる（プレビューと確定値で共通）
     * @param {object} dialogControls - 各パネルのコントロール
     * @returns {object} 描画設定
     */
    function collectDrawSettings(dialogControls) {
        var colorControls = dialogControls.color;
        var placementControls = dialogControls.placement;
        var offsetControls = dialogControls.offset;

        var colorMode = ColorMode.NONE;
        if (colorControls.k100Radio.value) colorMode = ColorMode.K100;
        else if (colorControls.customRadio.value) colorMode = ColorMode.CUSTOM;

        var zOrder = placementControls.frontRadio.value ? 'front' :
            (placementControls.bgLayerRadio.value ? 'bg' : 'back');

        /* オフセット計算は resolveOffsetToPt に一元化 / All offset math lives in resolveOffsetToPt */
        var resolvedOffset = resolveOffsetToPt(offsetControls.offsetInput.text, getUnitInfo().code, !!offsetControls.bleedCheckbox.value);

        return {
            colorMode: colorMode,
            customColor: colorControls.getCustomColor(),
            offset: resolvedOffset.pt,
            zOrder: zOrder,
            target: dialogControls.target.allArtboardsRadio.value ? 'all' : 'current',
            makeGuide: !!dialogControls.options.makeGuideCheckbox.value,
            convertToLiveShape: !!dialogControls.options.convertToLiveShapeCheckbox.value
        };
    }

    /**
     * ボタン行（左：表示モード切り替え、右：キャンセル／OK）を構築する
     * @param {Window} settingsDialog - 追加先のダイアログ
     * @returns {{btnOK: Button, btnCancel: Button}} OK・キャンセルボタン
     */
    function buildButtonRow(settingsDialog) {
        var btnRowGroup = settingsDialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['fill', 'center'];
        btnRowGroup.alignment = 'fill';

        var btnLeftGroup = btnRowGroup.add('group');
        setupGroup(btnLeftGroup, 'row');

        var isPreviewDisplayMode = true;
        var btnDisplayToggle = btnLeftGroup.add('button', undefined, getLabel('button.previewOutline'));
        btnDisplayToggle.helpTip = getLabel('tooltip.previewToggle');

        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.orientation = 'row';
        btnRightGroup.alignment = ['right', 'center']; /* 右カラムは右揃え / right-align the right column */

        var btnCancel = btnRightGroup.add('button', undefined, getLabel('button.cancel'));
        var btnOK = btnRightGroup.add('button', undefined, getLabel('button.ok'));

        btnDisplayToggle.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                isPreviewDisplayMode = !isPreviewDisplayMode;
                btnDisplayToggle.text = isPreviewDisplayMode ? getLabel('button.previewOutline') : getLabel('button.previewPreview');
            } catch (e) { }
        };

        return { btnOK: btnOK, btnCancel: btnCancel };
    }

    /**
     * 設定ダイアログを構築して結果を返す
     * @returns {object|null} 描画設定。キャンセル時は null
     */
    function showDialog() {
        var settingsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        settingsDialog.alignChildren = 'left';

        /* 各パネルより先に定義してコールバックとして配る（実行はパネル構築後）
           Declared before the panels so they can be handed out as callbacks */
        var dialogControls = null;

        function updatePreviewImmediately() {
            PreviewHistory.cancelTask(previewDebounceTaskId);
            try {
                renderPreview(app.activeDocument, collectDrawSettings(dialogControls));
            } catch (e) { }
        }

        function updatePreviewDeferred() {
            try {
                schedulePreview(collectDrawSettings(dialogControls), PREVIEW_DELAY_TYPING_MS);
            } catch (e) { }
        }

        var previewHooks = {
            immediate: updatePreviewImmediately,
            deferred: updatePreviewDeferred
        };

        /* 2カラム構成 / Two-column layout */
        var mainColumnsGroup = settingsDialog.add('group');
        setupGroup(mainColumnsGroup, 'row', COLUMN_SPACING);
        mainColumnsGroup.alignChildren = ['fill', 'top']; /* 2カラムを上揃え・横いっぱいに */

        var leftColumnGroup = mainColumnsGroup.add('group');
        setupGroup(leftColumnGroup, 'column', STACK_SPACING);
        leftColumnGroup.alignChildren = 'fill'; /* パネルを列幅いっぱいに / panels fill the column */

        var rightColumnGroup = mainColumnsGroup.add('group');
        setupGroup(rightColumnGroup, 'column', STACK_SPACING);
        rightColumnGroup.alignChildren = 'fill';

        dialogControls = {
            offset: buildOffsetPanel(leftColumnGroup, previewHooks),
            placement: buildPlacementPanel(leftColumnGroup, previewHooks),
            options: buildOptionsPanel(leftColumnGroup),
            color: buildColorPanel(rightColumnGroup, previewHooks),
            target: buildTargetPanel(rightColumnGroup, previewHooks)
        };

        addDialogShortcutKeys(settingsDialog, dialogControls);

        var dialogButtons = buildButtonRow(settingsDialog);

        /* プレビューを片付けてから閉じる / Clean the preview up, then close */
        function closeWithCleanup(resultCode) {
            PreviewHistory.cancelTask(previewDebounceTaskId);
            PreviewHistory.undo();
            clearPreview(true); /* undo回数に依存せず _preview レイヤーを確実に削除 */
            settingsDialog.close(resultCode);
        }

        dialogButtons.btnOK.onClick = function () { closeWithCleanup(1); };
        dialogButtons.btnCancel.onClick = function () { closeWithCleanup(0); };

        settingsDialog.onShow = function () {
            dialogControls.offset.initFieldState();
            try { dialogControls.offset.offsetInput.active = true; } catch (e) { }
            PreviewHistory.start(); /* プレビューのUndoカウンタを初期化 */
            updatePreviewImmediately();
        };

        prepareDialogWindow(settingsDialog, SCRIPT_NAME);
        if (settingsDialog.show() != 1) return null;

        /* 確定値もプレビューと同じ計算経路から取る / Final values come from the same computation as the preview */
        return collectDrawSettings(dialogControls);
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 編集可能なレイヤーを取得する（なければ作成）
     * テンプレートレイヤーは locked が false でも編集できない（Error 8705）ので除外し、
     * プレビュー用レイヤーは本番の描画先にしない（残存時の誤描画防止）。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 描画先レイヤー
     */
    function getWritableLayer(doc) {
        function isWritableLayer(layer) {
            return !!layer && !layer.locked && layer.visible && !layer.template && !isPreviewLayerName(layer.name);
        }
        try {
            if (isWritableLayer(doc.activeLayer)) return doc.activeLayer;
            for (var i = 0; i < doc.layers.length; i++) {
                if (isWritableLayer(doc.layers[i])) return doc.layers[i];
            }
            var newLayer = doc.layers.add();
            newLayer.name = FALLBACK_LAYER_NAME;
            return newLayer;
        } catch (e) { }
        return doc.activeLayer;
    }

    /**
     * 「bg」レイヤーを取得する（なければ作成し、最背面へ移動）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} bgレイヤー
     */
    function getOrCreateBgLayer(doc) {
        var bgLayer = findLayerByName(doc, BG_LAYER_NAME);
        if (!bgLayer) {
            bgLayer = doc.layers.add();
            bgLayer.name = BG_LAYER_NAME;
        }
        /* 見える＆編集可能にしてから最背面へ / Make it visible and editable, then send it to the back */
        bgLayer.visible = true;
        bgLayer.locked = false;
        try {
            bgLayer.printable = true;
            bgLayer.move(doc, ElementPlacement.PLACEATEND);
        } catch (e) { }
        return bgLayer;
    }

    /**
     * アートボードと同サイズ（オフセット込み）の長方形を1枚描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {Artboard} artboard - 対象アートボード
     * @param {object} drawSettings - 描画設定
     * @returns {PathItem} 描画した長方形
     */
    function drawRectangleForArtboard(doc, artboard, drawSettings) {
        var rectangleBounds = getOffsetRectangleBounds(artboard.artboardRect, drawSettings.offset);
        var targetLayer = (drawSettings.zOrder === 'bg') ? getOrCreateBgLayer(doc) : getWritableLayer(doc);
        /* アクティブレイヤーがロックされたままだと、別の編集可能レイヤーへ作成しても
           Illustrator が Error 8705（対象レイヤーは編集できません）を投げる。
           Make the target layer editable AND active before creating, or a locked active layer
           triggers Error 8705 "Target layer cannot be modified". */
        try {
            targetLayer.locked = false;
            targetLayer.visible = true;
            doc.activeLayer = targetLayer;
        } catch (e) { }

        var artboardRectangle = targetLayer.pathItems.rectangle(
            rectangleBounds.top, rectangleBounds.left, rectangleBounds.width, rectangleBounds.height);

        applyFillByMode(doc, artboardRectangle, drawSettings);
        artboardRectangle.name = getLabel('objectName.rect');
        artboardRectangle.selected = true;
        applyZOrder(artboardRectangle, drawSettings.zOrder);

        return artboardRectangle;
    }

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /**
     * 中心の○（属性パネル「中心点を表示」）を選択オブジェクトへ適用する
     * API・メニューコマンドからは設定できないため、記録済みアクション(.aia)を一時ファイルへ書き出して
     * loadAction→doScript で再生する。呼び出し側で「対象だけを選択した状態」にしてから実行すること。
     * @returns {void}
     */
    function showShapeCenterWidget() {
        var ACTION_SET_NAME = 'SmartDrawArtboardRectangle';
        var ACTION_NAME = 'CenterPoint';
        var ACTION_BODY = [
            '/version 3',
            '/name [ 26',
            '\t536d61727444726177417274626f61726452656374616e676c65',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '\t/name [ 11',
            '\t\t43656e746572506f696e74',
            '\t]',
            '\t/keyIndex 0',
            '\t/colorIndex 0',
            '\t/isOpen 1',
            '\t/eventCount 1',
            '\t/event-1 {',
            '\t\t/useRulersIn1stQuadrant 0',
            '\t\t/internalName (adobe_attributePalette)',
            '\t\t/localizedName [ 12',
            '\t\t\te5b19ee680a7e8a8ade5ae9a',
            '\t\t]',
            '\t\t/isOpen 1',
            '\t\t/isOn 1',
            '\t\t/hasDialog 0',
            '\t\t/parameterCount 1',
            '\t\t/parameter-1 {',
            '\t\t\t/key 1668183154',
            '\t\t\t/showInPalette 4294967295',
            '\t\t\t/type (boolean)',
            '\t\t\t/value 1',
            '\t\t}',
            '\t}',
            '}',
            ''
        ].join('\n');

        /* 失敗しても描画は続ける（中心の○が付かないだけ）/ Keep going on failure: only the center widget is missing */
        runTemporaryAction(ACTION_BODY, ACTION_SET_NAME, ACTION_NAME);
    }

    /**
     * 指定アイテムだけを選択状態にする
     * @param {PathItem[]} items - 選択したいアイテム
     * @returns {void}
     */
    function selectOnly(items) {
        try { app.executeMenuCommand('deselectall'); } catch (e) { }
        try {
            for (var i = 0; i < items.length; i++) items[i].selected = true;
        } catch (e) { }
    }

    /**
     * 描画後のオプション（ライブシェイプ化／中心の○表示／ガイド化）を適用する
     * @param {PathItem[]} createdRectangles - 描画した長方形
     * @param {object} drawSettings - 描画設定
     * @returns {void}
     */
    function applyDrawOptions(createdRectangles, drawSettings) {
        if (!createdRectangles || !createdRectangles.length) return;

        if (drawSettings.convertToLiveShape) {
            /* 選択ベースのメニューコマンドなので、対象だけを選択してから実行 */
            selectOnly(createdRectangles);
            try { app.executeMenuCommand('Convert to Shape'); } catch (e) { }
            /* 中心の○はライブシェイプにのみ表示される。変換後に選択し直してから適用
               The center widget only renders on live shapes, so re-select after converting */
            selectOnly(createdRectangles);
            showShapeCenterWidget();
        }

        if (drawSettings.makeGuide) {
            /* PathItem.guides を直接立てる（選択・メニュー状態に依存せず確実）
               Set PathItem.guides directly - robust, independent of selection and menu state */
            for (var i = 0; i < createdRectangles.length; i++) {
                createdRectangles[i].name = getLabel('objectName.guide');
                createdRectangles[i].guides = true;
            }
            try { app.executeMenuCommand('deselectall'); } catch (e) { }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント：ダイアログ→描画→オプション適用
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var drawSettings = showDialog();
        if (drawSettings === null) return;

        var doc = app.activeDocument;
        var previousCoordinateSystem = switchToDocumentCoordinates();

        app.executeMenuCommand('deselectall'); /* 既存選択を解除 / clear any existing selection */

        var createdRectangles = [];
        if (drawSettings.target === 'all') {
            for (var i = 0; i < doc.artboards.length; i++) {
                createdRectangles.push(drawRectangleForArtboard(doc, doc.artboards[i], drawSettings));
            }
        } else {
            var currentArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            createdRectangles.push(drawRectangleForArtboard(doc, currentArtboard, drawSettings));
        }

        applyDrawOptions(createdRectangles, drawSettings);
        restoreCoordinateSystem(previousCoordinateSystem);
    }

    main();

})();
