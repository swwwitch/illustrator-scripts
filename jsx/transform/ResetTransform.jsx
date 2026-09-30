#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### 概要

配置画像・テキスト・長方形・クリップグループ・直線パスに対して、回転／シアー／スケール／縦横比を安全にリセットします。
バウンディングボックスをリセットしたあと元の中心位置へ戻すため、見た目の位置は保たれます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetTransform.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n52f6b645bc70

### Overview

Safely resets rotation, shear, scale and aspect ratio on placed images, text, rectangles, clipping groups and straight paths.
The bounding box is reset and the item is moved back to its original center, so its apparent position is preserved.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetTransform.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResetTransform";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.7";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetTransform.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetTransform.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n52f6b645bc70"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 軸スナップの許容範囲（度）/ Angle range that counts as "near an axis" */
    var AXIS_SNAP_MIN_DEG = 0.5;    /* これ未満はすでに正立とみなす / below this the path is treated as upright */
    var AXIS_SNAP_MAX_DEG = 44;     /* これを超えると意図的な傾きとみなす / above this the tilt is treated as intentional */

    /* スケール入力の下限（%）/ Minimum scale percent allowed in the input */
    var SCALE_MIN_PERCENT = 20;

    /* 判定用の微小値 / Numerical tolerances */
    var MATRIX_EPSILON = 1e-8;         /* 行列演算のゼロ判定 / zero threshold for matrix math */
    var SCALE_EPSILON = 1e-6;          /* スケール・シアーの残差判定 / residual threshold for scale and shear */
    var ROTATION_EPSILON_DEG = 0.0001; /* 回転の残差判定（度）/ residual threshold for rotation (deg) */

    // =========================================
    // レイアウト設定 / Layout settings
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

    var PANEL_SPACING_COMPACT = 6;          /* チェックボックスを並べるときの間隔 / spacing for checkbox stacks */

    // =========================================
    // 単位 / Units
    // =========================================
    /* ルーラー単位の換算は使用しません（スケールは % 指定）/ No ruler-unit conversion (scale is percent-based) */

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

    /* UIラベル（カテゴリ別）/ UI labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "リセット（回転・比率）", en: "Reset (Rotate / Scale)" }
        },
        panel: {
            placedImage: { ja: "配置画像", en: "Placed Images" },
            clippedGroup: { ja: "クリップグループ", en: "Clip Group" },
            textFrame: { ja: "テキスト", en: "Text" },
            rectanglePath: { ja: "長方形（パス）", en: "Rectangle (Path)" },
            straightLine: { ja: "パス（直線）", en: "Path (Line)" }
        },
        tooltip: {
            rotate:      { ja: "掛かっている回転を元に戻します。", en: "Clears the rotation." },
            shear:       { ja: "掛かっているシアー（傾き）を元に戻します。", en: "Clears the shear." },
            aspectRatio: { ja: "変形でくずれた縦横比を元に戻します。", en: "Restores the original aspect ratio." },
            flip:        { ja: "反転を元に戻します。", en: "Clears the flip." },
            scale:       { ja: "拡大・縮小率を、下の欄の値にそろえます。", en: "Sets the scale to the value in the field below." },
            scalePercent:{ ja: "そろえる拡大・縮小率（％）です。", en: "The scale, in percent, every object is set to." },
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
        checkbox: {
            rotate: { ja: "回転", en: "Rotate" },
            shear: { ja: "シアー", en: "Shear" },
            aspectRatio: { ja: "縦横比", en: "Aspect Ratio" },
            flip: { ja: "反転", en: "Flip" },
            scale: { ja: "スケール", en: "Scale" },
            textScaleRatio: { ja: "垂直比率／水平比率", en: "Horizontal & Vertical Scale" },
            tracking: { ja: "トラッキング", en: "Tracking" }
        },
        button: {
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectFirst: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noTarget: { ja: "リセットできる対象が選択されていません。", en: "No resettable objects are selected." }
        }
    };

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

    // =========================================
    // UIヘルパー / UI helpers
    // =========================================

    /**
     * カラム用のグループを作成する。
     * @param {object} parentGroup - 親グループ
     * @returns {object} 作成した縦並びグループ
     */
    function addColumnGroup(parentGroup) {
        var columnGroup = parentGroup.add('group');
        columnGroup.orientation = 'column';
        columnGroup.alignChildren = 'left';
        columnGroup.alignment = 'fill';
        return columnGroup;
    }

    /**
     * パネルにチェックボックスを追加する。
     * @param {object} targetPanel - 追加先のパネル
     * @param {object} labelEntry - ラベル定義
     * @param {boolean} initialValue - 初期のオン／オフ
     * @param {object} [tooltipEntry] - ツールチップのラベル定義
     * @returns {object} 追加したチェックボックス
     */
    function addCheckbox(targetPanel, labelEntry, initialValue, tooltipEntry) {
        var checkbox = targetPanel.add('checkbox', undefined, getLabel(labelEntry));
        checkbox.value = initialValue;
        if (tooltipEntry) checkbox.helpTip = getLabel(tooltipEntry);
        return checkbox;
    }

    /**
     * 入力欄にフォーカスして内容を全選択する。
     * @param {object} editText - 対象の入力欄
     * @returns {void}
     */
    function focusAndSelectAll(editText) {
        editText.active = true;
        editText.textselection = editText.text;
    }

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
    // 選択状態の判定 / Selection analysis
    // =========================================

    /**
     * 4点の閉じたパス（長方形とみなす）かどうかを判定する。
     * @param {object} pathItem - 判定するパス
     * @returns {boolean} 長方形とみなせるなら true
     */
    function isRectanglePath(pathItem) {
        return !!(pathItem.closed && pathItem.pathPoints && pathItem.pathPoints.length === 4);
    }

    /**
     * 2点の開いたパス（直線）かどうかを判定する。
     * @param {object} pathItem - 判定するパス
     * @returns {boolean} 直線とみなせるなら true
     */
    function isStraightLinePath(pathItem) {
        return !!(!pathItem.closed && pathItem.pathPoints && pathItem.pathPoints.length === 2);
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * 同じページアイテムを指しているかを判定する。
     * uuid が読めない環境では参照の一致で判定する。
     * @param {object} itemA - 比較するページアイテム
     * @param {object} itemB - 比較するページアイテム
     * @returns {boolean} 同じアイテムなら true
     */
    function isSameItem(itemA, itemB) {
        if (itemA === itemB) return true;
        try {
            if (itemA.uuid && itemB.uuid) return itemA.uuid === itemB.uuid;
        } catch (e) {}
        return false;
    }

    /**
     * まだ収集していなければ配列に追加する。
     * @param {array} collectedItems - 収集先の配列
     * @param {object} pageItem - 追加するページアイテム
     * @returns {void}
     */
    function pushUniqueItem(collectedItems, pageItem) {
        for (var i = 0; i < collectedItems.length; i++) {
            if (isSameItem(collectedItems[i], pageItem)) return;
        }
        collectedItems.push(pageItem);
    }

    /**
     * 親をたどって最上位のクリップグループを探す。
     * @param {object} pageItem - 起点のページアイテム
     * @returns {object} 最上位のクリップグループ。見つからなければ null
     */
    function findTopmostClippedAncestor(pageItem) {
        var topmostClippedGroup = null;
        var parentItem = pageItem.parent;
        /* 親がグループでなくなった時点（レイヤーに到達）で打ち切り / stop once the parent is no longer a group */
        while (parentItem && parentItem.typename === 'GroupItem') {
            if (parentItem.clipped === true) topmostClippedGroup = parentItem;
            parentItem = parentItem.parent;
        }
        return topmostClippedGroup;
    }

    /**
     * 選択からリセット対象を再帰的に集める。
     * グループ・複合パスは中身をたどり、クリップグループ内のオブジェクトは最上位のクリップグループにまとめる。
     * @param {array} items - たどるページアイテムの配列
     * @param {array} collectedItems - 収集先の配列
     * @returns {array} 収集した対象の配列
     */
    function collectResetTargets(items, collectedItems) {
        for (var i = 0; i < items.length; i++) {
            var pageItem = items[i];
            if (!pageItem || !pageItem.typename) continue;

            /* クリップグループの中身は、最上位のクリップグループとして一度だけ処理 / inside a clip group, target the topmost clip group once */
            var clippedAncestor = findTopmostClippedAncestor(pageItem);
            if (clippedAncestor) {
                pushUniqueItem(collectedItems, clippedAncestor);
                continue;
            }

            if (pageItem.typename === 'GroupItem') {
                if (pageItem.clipped === true) {
                    pushUniqueItem(collectedItems, pageItem);
                } else {
                    collectResetTargets(pageItem.pageItems, collectedItems);
                }
                continue;
            }

            if (pageItem.typename === 'CompoundPathItem') {
                collectResetTargets(pageItem.pathItems, collectedItems);
                continue;
            }

            pushUniqueItem(collectedItems, pageItem);
        }
        return collectedItems;
    }

    /**
     * 選択内容から、どの種別のリセットが使えるかを調べる。
     * @param {array} selectedItems - 選択中のページアイテム
     * @returns {object} 種別ごとの可否フラグ
     */
    function getSelectionCapabilities(selectedItems) {
        var capabilities = {
            hasPlacedOrRaster: false,
            hasClippedGroup: false,
            hasTextFrame: false,
            hasRectanglePath: false,
            hasStraightLine: false
        };
        if (!selectedItems || !selectedItems.length) return capabilities;

        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!selectedItem || !selectedItem.typename) continue;
            var typeName = selectedItem.typename;

            if (typeName === 'PlacedItem' || typeName === 'RasterItem') {
                capabilities.hasPlacedOrRaster = true;
            } else if (typeName === 'GroupItem' && selectedItem.clipped === true) {
                capabilities.hasClippedGroup = true;
            } else if (typeName === 'TextFrame') {
                capabilities.hasTextFrame = true;
            } else if (typeName === 'PathItem') {
                if (isRectanglePath(selectedItem)) capabilities.hasRectanglePath = true;
                if (isStraightLinePath(selectedItem)) capabilities.hasStraightLine = true;
            }
        }
        return capabilities;
    }

    // =========================================
    // ダイアログの各パネル / Dialog panels
    // =========================================

    /**
     * 配置画像パネルを作成する。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {boolean} isEnabled - 選択内容に配置画像が含まれるか
     * @returns {object} パネル内のコントロール一式
     */
    function buildPlacedImagePanel(parentGroup, isEnabled) {
        var pnlPlacedImage = parentGroup.add('panel', undefined, getLabel(LABELS.panel.placedImage));
        setupPanel(pnlPlacedImage, PANEL_SPACING_COMPACT);

        var cbRotate = addCheckbox(pnlPlacedImage, LABELS.checkbox.rotate, true, LABELS.tooltip.rotate);
        var cbShear = addCheckbox(pnlPlacedImage, LABELS.checkbox.shear, true, LABELS.tooltip.shear);
        var cbAspectRatio = addCheckbox(pnlPlacedImage, LABELS.checkbox.aspectRatio, true, LABELS.tooltip.aspectRatio);
        var cbFlip = addCheckbox(pnlPlacedImage, LABELS.checkbox.flip, true, LABELS.tooltip.flip);
        var cbScale = addCheckbox(pnlPlacedImage, LABELS.checkbox.scale, false, LABELS.tooltip.scale);

        var scaleInputGroup = pnlPlacedImage.add('group');
        scaleInputGroup.orientation = 'row';
        scaleInputGroup.alignChildren = 'center';
        scaleInputGroup.alignment = 'left'; /* 入力欄はパネル幅いっぱいに広げない / keep the scale input compact */

        /* ∧∨と入力欄は隙間0で突き合わせる。整数・下限 SCALE_MIN_PERCENT / butt the stepper against the field; integer, min SCALE_MIN_PERCENT */
        var scaleFieldGroup = scaleInputGroup.add('group');
        scaleFieldGroup.orientation = 'row';
        scaleFieldGroup.alignChildren = ['left', 'center'];
        scaleFieldGroup.spacing = 0;
        scaleFieldGroup.margins = 0;
        var etScalePercent;
        var scaleStepper = addStepper(scaleFieldGroup, function () { return etScalePercent; }, { integer: true, min: SCALE_MIN_PERCENT });
        etScalePercent = scaleFieldGroup.add('edittext', undefined, '100');
        etScalePercent.helpTip = getLabel(LABELS.tooltip.scalePercent);
        etScalePercent.characters = 5;
        bindSteppedArrowKeys(etScalePercent, scaleStepper);
        var stPercentUnit = scaleInputGroup.add('statictext', undefined, '%');

        /**
         * スケール入力欄の有効・無効をチェックボックスに連動させる。
         * @param {boolean} isScaleOn - スケールがオンかどうか
         * @returns {void}
         */
        function syncScaleInput(isScaleOn) {
            etScalePercent.enabled = isScaleOn;
            scaleStepper.enabled = isScaleOn;
            redrawSteppersIn(scaleStepper);
            stPercentUnit.enabled = isScaleOn;
            if (isScaleOn) focusAndSelectAll(etScalePercent);
        }

        etScalePercent.enabled = cbScale.value;
        scaleStepper.enabled = cbScale.value;
        stPercentUnit.enabled = cbScale.value;
        cbScale.onClick = function () {
            syncScaleInput(cbScale.value);
        };

        /* パネルを無効化すれば子コントロールもまとめて無効になる / disabling the panel disables its children */
        pnlPlacedImage.enabled = isEnabled;
        redrawSteppersIn(pnlPlacedImage); /* ∧∨は自作描画なので描き直す / redraw the custom-drawn stepper */

        return {
            rotate: cbRotate,
            shear: cbShear,
            aspectRatio: cbAspectRatio,
            flip: cbFlip,
            scale: cbScale,
            scalePercent: etScalePercent,
            syncScaleInput: syncScaleInput
        };
    }

    /**
     * クリップグループパネルを作成する。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {boolean} isEnabled - 選択内容にクリップグループが含まれるか
     * @returns {object} パネル内のコントロール一式
     */
    function buildClippedGroupPanel(parentGroup, isEnabled) {
        var pnlClippedGroup = parentGroup.add('panel', undefined, getLabel(LABELS.panel.clippedGroup));
        setupPanel(pnlClippedGroup, PANEL_SPACING_COMPACT);

        var controls = {
            rotate: addCheckbox(pnlClippedGroup, LABELS.checkbox.rotate, true, LABELS.tooltip.rotate),
            aspectRatio: addCheckbox(pnlClippedGroup, LABELS.checkbox.aspectRatio, true, LABELS.tooltip.aspectRatio),
            flip: addCheckbox(pnlClippedGroup, LABELS.checkbox.flip, true, LABELS.tooltip.flip)
        };
        pnlClippedGroup.enabled = isEnabled;
        return controls;
    }

    /**
     * テキストパネルを作成する。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {boolean} isEnabled - 選択内容にテキストが含まれるか
     * @returns {object} パネル内のコントロール一式
     */
    function buildTextFramePanel(parentGroup, isEnabled) {
        var pnlTextFrame = parentGroup.add('panel', undefined, getLabel(LABELS.panel.textFrame));
        setupPanel(pnlTextFrame, PANEL_SPACING_COMPACT);

        var controls = {
            rotate: addCheckbox(pnlTextFrame, LABELS.checkbox.rotate, true, LABELS.tooltip.rotate),
            shear: addCheckbox(pnlTextFrame, LABELS.checkbox.shear, true, LABELS.tooltip.shear),
            scaleRatio: addCheckbox(pnlTextFrame, LABELS.checkbox.textScaleRatio, true),
            tracking: addCheckbox(pnlTextFrame, LABELS.checkbox.tracking, true)
        };
        pnlTextFrame.enabled = isEnabled;
        return controls;
    }

    /**
     * 回転チェックボックスだけを持つパネルを作成する（長方形・直線用）。
     * @param {object} parentGroup - 追加先のカラムグループ
     * @param {object} titleEntry - パネル見出しのラベル定義
     * @param {boolean} isEnabled - 選択内容に対象が含まれるか
     * @returns {object} 追加した回転チェックボックス
     */
    function buildRotateOnlyPanel(parentGroup, titleEntry, isEnabled) {
        var rotateOnlyPanel = parentGroup.add('panel', undefined, getLabel(titleEntry));
        setupPanel(rotateOnlyPanel, PANEL_SPACING_COMPACT);

        var cbRotate = addCheckbox(rotateOnlyPanel, LABELS.checkbox.rotate, true, LABELS.tooltip.rotate);
        rotateOnlyPanel.enabled = isEnabled;
        return cbRotate;
    }

    /**
     * 入力された倍率を整数％に整えて下限でクランプする。
     * @param {string} scaleText - 入力欄の文字列
     * @returns {number} 適用する倍率（%）
     */
    function parseScalePercent(scaleText) {
        var scalePercent = parseFloat(scaleText);
        if (isNaN(scalePercent)) scalePercent = 100;
        scalePercent = Math.round(scalePercent); /* 整数％に統一（例：16.3 → 16）/ keep it an integer percent */
        return (scalePercent < SCALE_MIN_PERCENT) ? SCALE_MIN_PERCENT : scalePercent;
    }

    /**
     * リセットオプションのダイアログを表示する。
     * @param {array} selectedItems - 選択中のページアイテム
     * @returns {object} 選択されたオプション。キャンセル時は null
     */
    function showResetOptionsDialog(selectedItems) {
        var capabilities = getSelectionCapabilities(selectedItems);

        var mainDialog = new Window('dialog', getLabel(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        setupWindow(mainDialog);

        /* 2カラムレイアウト / two-column layout */
        var columnsGroup = mainDialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignment = 'fill';
        columnsGroup.alignChildren = 'top';
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = addColumnGroup(columnsGroup);
        var rightColumnGroup = addColumnGroup(columnsGroup);

        var placedImageControls = buildPlacedImagePanel(leftColumnGroup, capabilities.hasPlacedOrRaster);
        var clippedGroupControls = buildClippedGroupPanel(leftColumnGroup, capabilities.hasClippedGroup);
        var textFrameControls = buildTextFramePanel(rightColumnGroup, capabilities.hasTextFrame);
        var cbRectangleRotate = buildRotateOnlyPanel(rightColumnGroup, LABELS.panel.rectanglePath, capabilities.hasRectanglePath);
        var cbStraightLineRotate = buildRotateOnlyPanel(rightColumnGroup, LABELS.panel.straightLine, capabilities.hasStraightLine);

        /* Sキーでスケール、Fキーで反転を切り替え / 'S' toggles Scale, 'F' toggles Flip */
        addKeyShortcuts(mainDialog, {
            "S": placedImageControls.scale,
            "F": placedImageControls.flip
        }, { numericFields: [placedImageControls.scalePercent] });

        var buttonRow = addButtonRow(mainDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel(LABELS.button.cancel), { name: 'cancel' });
        var btnReset = buttonRow.rightGroup.add('button', undefined, getLabel(LABELS.button.reset), { name: 'ok' });
        btnReset.onClick = function () {
            mainDialog.close(1);
        };
        btnCancel.onClick = function () {
            mainDialog.close(0);
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        if (mainDialog.show() !== 1) return null; /* キャンセル / cancelled */

        return {
            placedRotate: placedImageControls.rotate.value,
            placedShear: placedImageControls.shear.value,
            placedAspectRatio: placedImageControls.aspectRatio.value,
            placedFlip: placedImageControls.flip.value,
            placedScale: placedImageControls.scale.value,
            placedScalePercent: parseScalePercent(placedImageControls.scalePercent.text),
            clipRotate: clippedGroupControls.rotate.value,
            clipAspectRatio: clippedGroupControls.aspectRatio.value,
            clipFlip: clippedGroupControls.flip.value,
            textRotate: textFrameControls.rotate.value,
            textShear: textFrameControls.shear.value,
            textScaleRatio: textFrameControls.scaleRatio.value,
            textTracking: textFrameControls.tracking.value,
            rectangleRotate: cbRectangleRotate.value,
            straightLineRotate: cbStraightLineRotate.value
        };
    }

    // =========================================
    // 行列ユーティリティ / Matrix utilities
    // =========================================

    /**
     * 行列を持つオブジェクトかどうかを判定する。
     * PathItem や GroupItem は matrix を持たないため、ここで弾く。
     * @param {object} pageItem - 判定するページアイテム
     * @returns {boolean} matrix を参照できるなら true
     */
    function hasMatrix(pageItem) {
        try {
            return !!(pageItem && pageItem.matrix && typeof pageItem.matrix.mValueA !== 'undefined');
        } catch (e) {
            return false;
        }
    }

    /**
     * 2×2行列を掛け合わせる。
     * @param {object} leftComponents - 左側の成分 { a, b, c, d }
     * @param {object} rightComponents - 右側の成分 { a, b, c, d }
     * @returns {object} 積の成分 { a, b, c, d }
     */
    function multiply2x2(leftComponents, rightComponents) {
        return {
            a: leftComponents.a * rightComponents.a + leftComponents.c * rightComponents.b,
            b: leftComponents.b * rightComponents.a + leftComponents.d * rightComponents.b,
            c: leftComponents.a * rightComponents.c + leftComponents.c * rightComponents.d,
            d: leftComponents.b * rightComponents.c + leftComponents.d * rightComponents.d
        };
    }

    /**
     * 2×2行列の逆行列を求める（行列式が0に近い場合は微小値で代用）。
     * @param {object} components - 成分 { a, b, c, d }
     * @returns {object} 逆行列の成分 { a, b, c, d }
     */
    function invert2x2(components) {
        var determinant = components.a * components.d - components.b * components.c;
        if (Math.abs(determinant) < MATRIX_EPSILON) {
            determinant = (determinant < 0 ? -1 : 1) * MATRIX_EPSILON;
        }
        var inverseDeterminant = 1.0 / determinant;
        return {
            a: components.d * inverseDeterminant,
            b: -components.b * inverseDeterminant,
            c: -components.c * inverseDeterminant,
            d: components.a * inverseDeterminant
        };
    }

    /**
     * 成分から平行移動なしの Matrix を作る。
     * @param {object} components - 成分 { a, b, c, d }
     * @returns {object} Illustrator の Matrix
     */
    function buildMatrixFromComponents(components) {
        var matrix = new Matrix();
        matrix.mValueA = components.a;
        matrix.mValueB = components.b;
        matrix.mValueC = components.c;
        matrix.mValueD = components.d;
        matrix.mValueTX = 0;
        matrix.mValueTY = 0;
        return matrix;
    }

    /**
     * 行列をQR分解し、向き・スケール・シアーに分ける。
     * @param {object} itemMatrix - 対象の Matrix
     * @returns {object} { scaleX, scaleY, shear, axis1X, axis1Y, axis2X, axis2Y }
     */
    function decomposeMatrix(itemMatrix) {
        var matrixA = itemMatrix.mValueA;
        var matrixB = itemMatrix.mValueB;
        var matrixC = itemMatrix.mValueC;
        var matrixD = itemMatrix.mValueD;

        var scaleX = Math.sqrt(matrixA * matrixA + matrixB * matrixB);
        if (scaleX === 0) scaleX = MATRIX_EPSILON;

        var axis1X = matrixA / scaleX;
        var axis1Y = matrixB / scaleX;

        var projection = axis1X * matrixC + axis1Y * matrixD;
        var residualX = matrixC - projection * axis1X;
        var residualY = matrixD - projection * axis1Y;

        var scaleY = Math.sqrt(residualX * residualX + residualY * residualY);
        if (scaleY === 0) {
            scaleY = MATRIX_EPSILON;
            residualX = -axis1Y;
            residualY = axis1X;
        }

        return {
            scaleX: scaleX,
            scaleY: scaleY,
            shear: projection / scaleX,
            axis1X: axis1X,
            axis1Y: axis1Y,
            axis2X: residualX / scaleY,
            axis2Y: residualY / scaleY
        };
    }

    /**
     * 分解結果の一部を差し替えて、目標となる2×2成分を組み立てる。
     * @param {object} decomposed - decomposeMatrix の戻り値
     * @param {number} scaleX - 目標の水平スケール
     * @param {number} scaleY - 目標の垂直スケール
     * @param {number} shear - 目標のシアー量
     * @returns {object} 成分 { a, b, c, d }
     */
    function buildTargetComponents(decomposed, scaleX, scaleY, shear) {
        var shearTerm = shear * scaleX;
        return {
            a: decomposed.axis1X * scaleX,
            b: decomposed.axis1Y * scaleX,
            c: decomposed.axis1X * shearTerm + decomposed.axis2X * scaleY,
            d: decomposed.axis1Y * shearTerm + decomposed.axis2Y * scaleY
        };
    }

    /**
     * 現在の行列を目標成分にするための差分行列を作る。
     * @param {object} itemMatrix - 現在の Matrix
     * @param {object} targetComponents - 目標の成分 { a, b, c, d }
     * @returns {object} 差分の Matrix
     */
    function buildDeltaMatrix(itemMatrix, targetComponents) {
        var currentComponents = {
            a: itemMatrix.mValueA,
            b: itemMatrix.mValueB,
            c: itemMatrix.mValueC,
            d: itemMatrix.mValueD
        };
        return buildMatrixFromComponents(multiply2x2(invert2x2(currentComponents), targetComponents));
    }

    /**
     * 行列を分解し、目標成分との差分だけを適用する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {function} buildTargetFn - 分解結果を受け取り目標成分を返す関数
     * @returns {void}
     */
    function applyDecomposedTransform(pageItem, buildTargetFn) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        pageItem.transform(buildDeltaMatrix(itemMatrix, buildTargetFn(decomposeMatrix(itemMatrix))));
    }

    /**
     * 縦横のスケールを大きいほうに揃えた目標成分を求める（整数％にスナップ）。
     * @param {object} decomposed - decomposeMatrix の戻り値
     * @returns {object} 成分 { a, b, c, d }
     */
    function getUniformScaleTarget(decomposed) {
        var uniformScale = Math.max(decomposed.scaleX, decomposed.scaleY);
        /* 整数％にスナップ（例：16.321% → 16%）/ snap to an integer percent */
        var uniformPercent = Math.round(uniformScale * 100);
        /* 0% に丸めるとオブジェクトが潰れるため最小1%を保証 / never collapse the item to 0% */
        if (uniformPercent < 1) uniformPercent = 1;
        uniformScale = uniformPercent / 100;
        return buildTargetComponents(decomposed, uniformScale, uniformScale, decomposed.shear);
    }

    // =========================================
    // 変形ヘルパー / Transform helpers
    // =========================================

    /**
     * バウンディングボックスをリセットする（位置は変えない）。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function resetBoundingBox(pageItem) {
        try {
            app.selection = null;
            app.selection = [pageItem];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (e) { }
    }

    /**
     * 変形前の中心位置に戻す。
     * @param {object} pageItem - 対象のページアイテム
     * @param {array} originalPosition - 変形前の position（左上座標）
     * @param {number} widthBefore - 変形前の幅
     * @param {number} heightBefore - 変形前の高さ
     * @returns {void}
     */
    function recenterToOriginalCenter(pageItem, originalPosition, widthBefore, heightBefore) {
        pageItem.position = [
            originalPosition[0] + widthBefore / 2 - pageItem.width / 2,
            originalPosition[1] - heightBefore / 2 + pageItem.height / 2
        ];
    }

    /**
     * 変形 → バウンディングボックスのリセット → 元の中心へ再配置、をまとめて行う。
     * @param {object} pageItem - 対象のページアイテム
     * @param {function} transformFn - 実際の変形処理
     * @returns {void}
     */
    function withBoundsResetAndRecenter(pageItem, transformFn) {
        var originalPosition = pageItem.position;
        var widthBefore = pageItem.width;
        var heightBefore = pageItem.height;

        transformFn();
        resetBoundingBox(pageItem);
        recenterToOriginalCenter(pageItem, originalPosition, widthBefore, heightBefore);
    }

    /**
     * ロック・非表示を一時解除して処理を実行し、元の状態に戻す。
     * @param {object} pageItem - 対象のページアイテム
     * @param {function} actionFn - 実行する処理
     * @returns {void}
     */
    function withUnlockedVisible(pageItem, actionFn) {
        var wasLocked = pageItem.locked;
        var wasHidden = pageItem.hidden;
        pageItem.locked = false;
        pageItem.hidden = false;
        try {
            actionFn();
        } finally {
            pageItem.locked = wasLocked;
            pageItem.hidden = wasHidden;
        }
    }

    /**
     * ロック・非表示を解除したうえで、中心基準・全オプション有効で変形を適用する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {object} transformMatrix - 適用する Matrix
     * @returns {void}
     */
    function transformItemUnlocked(pageItem, transformMatrix) {
        withUnlockedVisible(pageItem, function () {
            pageItem.transform(transformMatrix, true, true, true, true, true, Transformation.CENTER);
        });
    }

    /**
     * 行列の成分から回転角を求める。
     * @param {number} matrixA - 行列の a 成分
     * @param {number} matrixB - 行列の b 成分
     * @param {number} angleSign - 符号（RasterItem は -1、それ以外は 1）
     * @returns {number} 回転角（度）
     */
    function getRotationAngleDeg(matrixA, matrixB, angleSign) {
        var angleDeg = Math.atan2(matrixB, matrixA) * 180 / Math.PI;
        return (angleSign < 0) ? -angleDeg : angleDeg;
    }

    /**
     * アイテムを指定角度だけ回転する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {number} degrees - 回転角（度）
     * @returns {void}
     */
    function rotateItemBy(pageItem, degrees) {
        pageItem.transform(app.getRotationMatrix(degrees));
    }

    /**
     * テキストフレームを中心基準で回転する。
     * @param {object} textFrame - 対象のテキストフレーム
     * @param {number} degrees - 回転角（度）
     * @returns {void}
     */
    function rotateTextFrameBy(textFrame, degrees) {
        textFrame.rotate(degrees, true, true, true, true, Transformation.CENTER);
    }

    /**
     * 行列の符号規則に合わせて回転を打ち消す（配置画像・ラスター用）。
     * @param {object} pageItem - 対象のページアイテム
     * @param {number} angleSign - 符号（RasterItem は -1、それ以外は 1）
     * @returns {void}
     */
    function cancelRotationBySign(pageItem, angleSign) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        rotateItemBy(pageItem, getRotationAngleDeg(itemMatrix.mValueA, itemMatrix.mValueB, angleSign));
    }

    /**
     * 現在の角度の逆回転を掛けて 0° に戻す（テキスト用）。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function cancelRotationToZero(pageItem) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        var angleDeg = Math.atan2(itemMatrix.mValueB, itemMatrix.mValueA) * 180 / Math.PI;
        if (pageItem.typename === 'TextFrame') {
            rotateTextFrameBy(pageItem, -angleDeg);
        } else {
            rotateItemBy(pageItem, -angleDeg);
        }
    }

    /**
     * シアーだけを取り除く。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function removeShear(pageItem) {
        applyDecomposedTransform(pageItem, function (decomposed) {
            return buildTargetComponents(decomposed, decomposed.scaleX, decomposed.scaleY, 0);
        });
    }

    /**
     * シアーを取り除き、丸め誤差で残った微小シアーをもう一度取り除く。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function removeShearWithRetry(pageItem) {
        removeShear(pageItem);
        if (!hasMatrix(pageItem)) return;
        if (Math.abs(decomposeMatrix(pageItem.matrix).shear) > SCALE_EPSILON) removeShear(pageItem);
    }

    /**
     * スケールを100%に正規化する（向きとシアーは保つ）。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function normalizeScaleTo100(pageItem) {
        applyDecomposedTransform(pageItem, function (decomposed) {
            return buildTargetComponents(decomposed, 1, 1, decomposed.shear);
        });
    }

    /**
     * 縦横のスケールを大きいほうに揃える。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function equalizeScaleToLarger(pageItem) {
        applyDecomposedTransform(pageItem, getUniformScaleTarget);
    }

    /**
     * 指定した倍率で等比拡大・縮小する。
     * @param {object} pageItem - 対象のページアイテム
     * @param {number} scalePercent - 倍率（100 = 100%）
     * @returns {void}
     */
    function applyUniformScalePercent(pageItem, scalePercent) {
        if (!(scalePercent > 0)) return;
        pageItem.resize(scalePercent, scalePercent, true, true, true, true, true, Transformation.CENTER);
    }

    // =========================================
    // 反転の解除 / Flip handling
    // =========================================

    /**
     * 左右反転しているかを判定する（回転・シアーなしが前提）。
     * @param {object} itemMatrix - 対象の Matrix
     * @returns {boolean} 左右反転していれば true
     */
    function isFlippedHorizontal(itemMatrix) {
        return itemMatrix.mValueA < 0;
    }

    /**
     * 上下反転しているかを判定する（回転・シアーなしが前提）。
     * 配置画像・ラスターは既定で mValueD が負のため、正のときを反転とみなす。
     * @param {object} itemMatrix - 対象の Matrix
     * @returns {boolean} 上下反転していれば true
     */
    function isFlippedVertical(itemMatrix) {
        return itemMatrix.mValueD > 0;
    }

    /**
     * 反転フラグに応じて反転を打ち消す。
     * @param {object} pageItem - 対象のページアイテム
     * @param {boolean} hasHorizontalFlip - 左右反転しているか
     * @param {boolean} hasVerticalFlip - 上下反転しているか
     * @returns {void}
     */
    function undoFlipByFlags(pageItem, hasHorizontalFlip, hasVerticalFlip) {
        if (!pageItem || (!hasHorizontalFlip && !hasVerticalFlip)) return;
        pageItem.transform(
            app.getScaleMatrix(hasHorizontalFlip ? -100 : 100, hasVerticalFlip ? -100 : 100),
            true, /* changePositions */
            true, /* changeFillPatterns */
            true, /* changeFillGradients */
            true, /* changeStrokePattern */
            true, /* changeLineWidths */
            Transformation.CENTER
        );
    }

    /**
     * 自身の行列から反転を判定して打ち消す。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {void}
     */
    function undoItemFlip(pageItem) {
        if (!hasMatrix(pageItem)) return;
        var itemMatrix = pageItem.matrix;
        undoFlipByFlags(pageItem, isFlippedHorizontal(itemMatrix), isFlippedVertical(itemMatrix));
    }

    // =========================================
    // 軸へのスナップ / Axis snapping
    // =========================================

    /**
     * 角度を -180〜180 の範囲に正規化する。
     * @param {number} angleDeg - 角度（度）
     * @returns {number} 正規化した角度（度）
     */
    function normalizeTo180(angleDeg) {
        while (angleDeg > 180) angleDeg -= 360;
        while (angleDeg < -180) angleDeg += 360;
        return angleDeg;
    }

    /**
     * 角度を -90〜90 の範囲に畳み込む。
     * @param {number} angleDeg - 角度（度）
     * @returns {number} 畳み込んだ角度（度）
     */
    function clampTo90(angleDeg) {
        angleDeg = normalizeTo180(angleDeg);
        if (angleDeg > 90) angleDeg -= 180;
        if (angleDeg < -90) angleDeg += 180;
        return angleDeg;
    }

    /**
     * 最も近い軸（0°／90°）へ向かう最小の回転量を求める。
     * @param {number} angleDeg - 現在の角度（度）
     * @returns {number} 最小の回転量（度）
     */
    function getRotationToNearestAxis(angleDeg) {
        var toHorizontal = clampTo90(-angleDeg);
        var toVertical = clampTo90(90 - angleDeg);
        return (Math.abs(toHorizontal) <= Math.abs(toVertical)) ? toHorizontal : toVertical;
    }

    /**
     * パスの最初の2アンカーが作る辺の角度を求める。
     * @param {object} pathItem - 対象のパス
     * @returns {number} 水平からの角度（度）。2点が同一なら null
     */
    function getFirstEdgeAngleDeg(pathItem) {
        var anchorA = pathItem.pathPoints[0].anchor;
        var anchorB = pathItem.pathPoints[1].anchor;
        var dx = anchorB[0] - anchorA[0];
        var dy = anchorB[1] - anchorA[1];
        if (dx === 0 && dy === 0) return null;
        return Math.atan2(dy, dx) * 180 / Math.PI;
    }

    /**
     * わずかに傾いたパスを最も近い軸へスナップする（長方形・直線で共用）。
     * 許容範囲を外れた傾きは意図的とみなして何もしない。
     * @param {object} pathItem - 対象のパス
     * @returns {boolean} スナップしたら true
     */
    function snapPathToNearestAxis(pathItem) {
        var edgeAngleDeg = getFirstEdgeAngleDeg(pathItem);
        if (edgeAngleDeg === null) return false;

        var rotationDeg = getRotationToNearestAxis(edgeAngleDeg);
        var distanceToAxis = Math.abs(rotationDeg);
        if (distanceToAxis < AXIS_SNAP_MIN_DEG || distanceToAxis > AXIS_SNAP_MAX_DEG) return false;

        withBoundsResetAndRecenter(pathItem, function () {
            rotateItemBy(pathItem, rotationDeg);
        });
        return true;
    }

    // =========================================
    // 配置画像・ラスター / Placed & raster items
    // =========================================

    /**
     * 配置画像・ラスターの変形をリセットする（回転→反転→縦横比→スケール→シアーの順）。
     * @param {object} pageItem - 対象のページアイテム
     * @param {string} itemTypeName - typename（"PlacedItem" または "RasterItem"）
     * @param {object} resetOptions - ダイアログで選択したオプション
     * @returns {void}
     */
    function resetPlacedOrRasterTransforms(pageItem, itemTypeName, resetOptions) {
        var rotationSign = (itemTypeName === "RasterItem") ? -1 : 1;

        withBoundsResetAndRecenter(pageItem, function () {
            /* 1) 回転（向きを安定させるため最初に）/ rotation first */
            if (resetOptions.placedRotate) cancelRotationBySign(pageItem, rotationSign);

            /* 2) 反転（回転補正後に判定）/ flip, after the rotation is cancelled */
            if (resetOptions.placedFlip) undoItemFlip(pageItem);

            /* 3) 縦横比の等比化（絶対スケールの前に）/ equalize the aspect ratio before absolute scaling */
            if (resetOptions.placedAspectRatio) equalizeScaleToLarger(pageItem);

            /* 4) 絶対スケール（100%へ正規化してから指定%）/ normalize to 100%, then apply the requested percent */
            if (resetOptions.placedScale) {
                normalizeScaleTo100(pageItem);
                applyUniformScalePercent(pageItem, resetOptions.placedScalePercent);
            }

            /* 5) シアー除去は最後（上記で混入した微小シアーも取り除く）/ shear removal last */
            if (resetOptions.placedShear) removeShearWithRetry(pageItem);
        });
    }

    // =========================================
    // テキスト / Text frames
    // =========================================

    /**
     * テキストの水平比率・垂直比率を100%に戻す。
     * @param {object} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function resetTextFrameScaleRatio(textFrame) {
        if (!textFrame.textRange) return;
        var textRange = textFrame.textRange;
        textRange.scaling = [1, 1];

        /* 丸め誤差が残った場合だけもう一度適用 / re-apply only when a rounding residual remains */
        var currentScaling = textRange.scaling;
        if (Math.abs(currentScaling[0] - 1) > SCALE_EPSILON || Math.abs(currentScaling[1] - 1) > SCALE_EPSILON) {
            textRange.scaling = [1, 1];
        }
    }

    /**
     * テキストのトラッキングを0に戻す。
     * @param {object} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function resetTextFrameTracking(textFrame) {
        if (!textFrame.textRange) return;
        textFrame.textRange.characterAttributes.tracking = 0;
    }

    /**
     * テキストの回転・シアー・比率・トラッキングをまとめてリセットする。
     * @param {object} textFrame - 対象のテキストフレーム
     * @param {object} resetOptions - ダイアログで選択したオプション
     * @returns {void}
     */
    function resetTextFrameTransforms(textFrame, resetOptions) {
        withBoundsResetAndRecenter(textFrame, function () {
            /* 1) まず回転を0°へ（ポイント文字・エリア内文字の双方に有効）/ rotation first */
            if (resetOptions.textRotate) cancelRotationToZero(textFrame);
            /* 2) シアー除去 / shear removal */
            if (resetOptions.textShear) removeShearWithRetry(textFrame);
            /* 3) 比率 / ratio */
            if (resetOptions.textScaleRatio) resetTextFrameScaleRatio(textFrame);
            /* 4) トラッキングは幅が変わるため最後 / tracking last, it changes the width */
            if (resetOptions.textTracking) resetTextFrameTracking(textFrame);
        });
    }

    // =========================================
    // クリップグループ / Clipped groups
    // =========================================

    /**
     * バウンディングボックスから面積を求める。
     * @param {object} pageItem - 対象のページアイテム
     * @returns {number} 面積
     */
    function getBoundsArea(pageItem) {
        return Math.abs(pageItem.width * pageItem.height);
    }

    /**
     * 複合パスがマスクとして使われているかを判定する。
     * @param {object} compoundPath - 対象の CompoundPathItem
     * @returns {boolean} 子パスのいずれかがマスクなら true
     */
    function isClippingCompoundPath(compoundPath) {
        var childPaths = compoundPath.pathItems || [];
        for (var i = 0; i < childPaths.length; i++) {
            if (childPaths[i] && childPaths[i].clipping) return true;
        }
        return false;
    }

    /**
     * クリップグループから、代表となる配置画像とマスクパスを再帰的に探す。
     * どちらも面積が最大のものを採用する。
     * @param {object} container - 探索するグループ
     * @returns {object} { image, clipPath }（見つからなければ null）
     */
    function findLargestImageAndClipPath(container) {
        var largestImage = null;
        var largestImageArea = -1;
        var largestClipPath = null;
        var largestClipPathArea = -1;

        /**
         * 候補のほうが面積が大きければ採用する。
         * @param {object} candidate - 候補のページアイテム
         * @param {boolean} isClipPath - マスクパスとして扱うか
         * @returns {void}
         */
        function keepIfLarger(candidate, isClipPath) {
            if (!candidate) return;
            var candidateArea = getBoundsArea(candidate);
            if (isClipPath) {
                if (candidateArea > largestClipPathArea) {
                    largestClipPath = candidate;
                    largestClipPathArea = candidateArea;
                }
            } else if (candidateArea > largestImageArea) {
                largestImage = candidate;
                largestImageArea = candidateArea;
            }
        }

        var childItems = container.pageItems || [];
        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];
            if (!childItem) continue;
            var typeName = childItem.typename;

            if (typeName === 'PlacedItem' || typeName === 'RasterItem') {
                keepIfLarger(childItem, false);
            } else if (typeName === 'PathItem' && childItem.clipping) {
                keepIfLarger(childItem, true);
            } else if (typeName === 'CompoundPathItem' && isClippingCompoundPath(childItem)) {
                /* 複合パス自身の面積を代用値として使う / use the compound's own area as a proxy */
                keepIfLarger(childItem, true);
            } else if (typeName === 'GroupItem') {
                var nestedResult = findLargestImageAndClipPath(childItem);
                keepIfLarger(nestedResult.image, false);
                keepIfLarger(nestedResult.clipPath, true);
            }
        }

        return { image: largestImage, clipPath: largestClipPath };
    }

    /**
     * マスクパスのうち、実際に変形を掛ける対象を決める。
     * 複合パスなら、マスク指定されている子パスを優先する。
     * @param {object} clipPathCandidate - 見つかったマスクパス
     * @returns {object} 変形対象。マスクパスがなければ null
     */
    function resolveClippingTransformTarget(clipPathCandidate) {
        if (!clipPathCandidate) return null;
        if (clipPathCandidate.typename !== 'CompoundPathItem') return clipPathCandidate;

        var childPaths = clipPathCandidate.pathItems || [];
        for (var i = 0; i < childPaths.length; i++) {
            if (childPaths[i] && childPaths[i].clipping) return childPaths[i];
        }
        /* マスク指定の子が見つからなければ複合パス自体を変形 / fall back to the compound itself */
        return clipPathCandidate;
    }

    /**
     * クリップグループの構成要素を一度だけ集める。
     * @param {object} groupItem - 対象のクリップグループ
     * @returns {object} { image, clipPath, clipTarget }
     */
    function collectClippedGroupParts(groupItem) {
        var clipParts = findLargestImageAndClipPath(groupItem);
        clipParts.clipTarget = resolveClippingTransformTarget(clipParts.clipPath);
        return clipParts;
    }

    /**
     * クリップグループの回転をリセットする。
     * 子同士の位置関係を保つため、グループごと回転する。
     * @param {object} groupItem - 対象のクリップグループ
     * @param {object} clipParts - collectClippedGroupParts の戻り値
     * @returns {boolean} 回転したら true
     */
    function resetClippedGroupRotation(groupItem, clipParts) {
        var representativeImage = clipParts.image;
        if (!representativeImage || !hasMatrix(representativeImage)) return false;

        var imageMatrix = representativeImage.matrix;
        var rotationSign = (representativeImage.typename === 'RasterItem') ? -1 : 1;
        var rotationDeg = getRotationAngleDeg(imageMatrix.mValueA, imageMatrix.mValueB, rotationSign);
        if (Math.abs(rotationDeg) <= ROTATION_EPSILON_DEG) return false;

        withBoundsResetAndRecenter(groupItem, function () {
            rotateItemBy(groupItem, rotationDeg);
        });
        return true;
    }

    /**
     * クリップグループの反転を、配置画像とマスクパスに同じだけ適用して打ち消す。
     * @param {object} clipParts - collectClippedGroupParts の戻り値
     * @returns {boolean} 反転を打ち消したら true
     */
    function undoClippedGroupFlip(clipParts) {
        var representativeImage = clipParts.image;
        if (!representativeImage || !hasMatrix(representativeImage)) return false;

        var imageMatrix = representativeImage.matrix;
        var hasHorizontalFlip = isFlippedHorizontal(imageMatrix);
        var hasVerticalFlip = isFlippedVertical(imageMatrix);
        if (!hasHorizontalFlip && !hasVerticalFlip) return false;

        withUnlockedVisible(representativeImage, function () {
            undoFlipByFlags(representativeImage, hasHorizontalFlip, hasVerticalFlip);
        });
        if (clipParts.clipTarget) {
            withUnlockedVisible(clipParts.clipTarget, function () {
                undoFlipByFlags(clipParts.clipTarget, hasHorizontalFlip, hasVerticalFlip);
            });
        }
        return true;
    }

    /**
     * クリップグループの縦横比を等比に戻す。
     * 配置画像から求めた差分を、マスクパスにも同じだけ適用する。
     * @param {object} clipParts - collectClippedGroupParts の戻り値
     * @returns {boolean} 等比化したら true
     */
    function resetClippedGroupAspectRatio(clipParts) {
        var representativeImage = clipParts.image;
        if (!representativeImage || !hasMatrix(representativeImage)) return false;

        var imageMatrix = representativeImage.matrix;
        var deltaMatrix = buildDeltaMatrix(imageMatrix, getUniformScaleTarget(decomposeMatrix(imageMatrix)));

        transformItemUnlocked(representativeImage, deltaMatrix);
        if (clipParts.clipTarget) transformItemUnlocked(clipParts.clipTarget, deltaMatrix);

        /* 丸め誤差で等比になりきらなかった場合の再調整 / re-equalize if a rounding residual remains */
        var decomposedAfter = decomposeMatrix(representativeImage.matrix);
        if (Math.abs(decomposedAfter.scaleX - decomposedAfter.scaleY) > SCALE_EPSILON) {
            withUnlockedVisible(representativeImage, function () {
                equalizeScaleToLarger(representativeImage);
            });
        }
        return true;
    }

    // =========================================
    // 種別ごとの振り分け / Dispatch by typename
    // =========================================

    /**
     * typename をキーにしたハンドラー一覧を作る。
     * 各ハンドラーは「リセット対象として処理したか」を返す。
     * @param {object} resetOptions - ダイアログで選択したオプション
     * @returns {object} typename → ハンドラー関数の対応表
     */
    function makeItemHandlers(resetOptions) {

        /**
         * テキストを処理する。
         * @param {object} textFrame - 対象のテキストフレーム
         * @returns {boolean} 処理したら true
         */
        function handleTextFrame(textFrame) {
            if (!resetOptions.textRotate && !resetOptions.textShear &&
                !resetOptions.textScaleRatio && !resetOptions.textTracking) return false;
            resetTextFrameTransforms(textFrame, resetOptions);
            return true;
        }

        /**
         * パス（長方形・直線）を処理する。
         * すでに正立している場合も「対象として処理済み」として扱う。
         * @param {object} pathItem - 対象のパス
         * @returns {boolean} 処理したら true
         */
        function handlePathItem(pathItem) {
            if (resetOptions.straightLineRotate && isStraightLinePath(pathItem)) {
                snapPathToNearestAxis(pathItem);
                return true;
            }
            if (resetOptions.rectangleRotate && isRectanglePath(pathItem)) {
                snapPathToNearestAxis(pathItem);
                return true;
            }
            return false;
        }

        /**
         * クリップグループを処理する。
         * @param {object} groupItem - 対象のグループ
         * @returns {boolean} 処理したら true
         */
        function handleClippedGroup(groupItem) {
            if (groupItem.clipped !== true) return false;

            /* 行列を読む前にバウンディングボックスを更新 / refresh the bounds before reading matrices */
            resetBoundingBox(groupItem);
            var clipParts = collectClippedGroupParts(groupItem);

            var didRotate = resetOptions.clipRotate && resetClippedGroupRotation(groupItem, clipParts);
            /* 反転判定の前に、回転後の状態を反映させる / let flip detection see the rotated state */
            if (didRotate && resetOptions.clipFlip) resetBoundingBox(groupItem);

            var didFlip = resetOptions.clipFlip && undoClippedGroupFlip(clipParts);
            var didAspectRatio = resetOptions.clipAspectRatio && resetClippedGroupAspectRatio(clipParts);

            /* 配置画像があれば無変更でも対象とみなす（誤アラート防止）/ an image means it was a valid target */
            return !!(didRotate || didFlip || didAspectRatio || clipParts.image);
        }

        /**
         * 配置画像・ラスターを処理する。
         * @param {object} pageItem - 対象のページアイテム
         * @param {string} itemTypeName - typename
         * @returns {boolean} 処理したら true
         */
        function handlePlacedOrRaster(pageItem, itemTypeName) {
            if (!resetOptions.placedRotate && !resetOptions.placedShear && !resetOptions.placedScale &&
                !resetOptions.placedAspectRatio && !resetOptions.placedFlip) return false;
            resetPlacedOrRasterTransforms(pageItem, itemTypeName, resetOptions);
            return true;
        }

        return {
            TextFrame: handleTextFrame,
            PathItem: handlePathItem,
            GroupItem: handleClippedGroup,
            PlacedItem: function (pageItem) {
                return handlePlacedOrRaster(pageItem, 'PlacedItem');
            },
            RasterItem: function (pageItem) {
                return handlePlacedOrRaster(pageItem, 'RasterItem');
            }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリーポイント。
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var currentDocument = app.activeDocument;
        var currentSelection = currentDocument.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel(LABELS.alert.selectFirst));
            return;
        }

        /* 処理中に app.selection を張り替えるため、開始時の選択を控える / snapshot the selection */
        var originalSelection = [];
        for (var i = 0; i < currentSelection.length; i++) originalSelection.push(currentSelection[i]);

        /* グループ・複合パスの中身までたどって対象を集める / dig into groups and compound paths */
        var resetTargets = collectResetTargets(originalSelection, []);

        var resetOptions = showResetOptionsDialog(resetTargets);
        if (!resetOptions) return;

        var itemHandlers = makeItemHandlers(resetOptions);
        var processedCount = 0;

        for (var j = 0; j < resetTargets.length; j++) {
            var targetItem = resetTargets[j];
            var itemHandler = itemHandlers[targetItem.typename];
            if (itemHandler && itemHandler(targetItem)) processedCount++;
        }

        /* 元の選択に戻す / restore the original selection */
        currentDocument.selection = originalSelection;

        if (processedCount === 0) alert(getLabel(LABELS.alert.noTarget));
    }

    main();

})();
