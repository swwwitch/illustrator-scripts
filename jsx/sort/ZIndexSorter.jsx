#target illustrator
#targetengine "ZIndexSorterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの重ね順を、指定した基準で並べ替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZIndexSorter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n3230f35844b2

### Overview

Reorders the stacking order of the selected objects according to a chosen criterion.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZIndexSorter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ZIndexSorter";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZIndexSorter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ZIndexSorter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n3230f35844b2"; /* 紹介記事 / article URL */

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
     * 横並びの行グループの共通設定（ボタン列など）
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - alignment（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = rowAlignment || "left";
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

    var RADIO_SPACING = 6; /* ラジオボタンの縦の間隔 / spacing between radio buttons */

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
            title: { ja: "重ね順ソート", en: "Z-Index Sorter" }
        },
        panel: {
            sortBasis: { ja: "並べ替えの基準", en: "Sort By" },
            sortOrder: { ja: "並び順", en: "Order" }
        },
        radio: {
            zOrder: { ja: "元の重ね順", en: "Original Stacking Order" },
            xAxis: { ja: "X座標（左端）", en: "X Position (Left Edge)" },
            yAxis: { ja: "Y座標（上端）", en: "Y Position (Top Edge)" },
            asc: { ja: "昇順", en: "Ascending" },
            desc: { ja: "降順", en: "Descending" },
            random: { ja: "ランダム", en: "Random" }
        },
        tooltip: {
            zOrder: {
                ja: "ダイアログを開いたときの重ね順を基準にします。昇順で元のまま、降順で反転します。",
                en: "Uses the stacking order at the time the dialog opened. Ascending keeps it; Descending reverses it."
            },
            xAxis: { ja: "左端の位置で重ね順を組み直します。", en: "Restacks the objects by their left edge." },
            yAxis: { ja: "上端の位置で重ね順を組み直します。", en: "Restacks the objects by their top edge." },
            asc: {
                ja: "X座標は左、Y座標は上にあるものほど背面にします。",
                en: "Objects further left (X) or higher up (Y) go further back."
            },
            desc: {
                ja: "X座標は右、Y座標は下にあるものほど背面にします。",
                en: "Objects further right (X) or lower down (Y) go further back."
            },
            random: { ja: "基準に関係なく、重ね順をシャッフルします。", en: "Shuffles the stacking order regardless of the basis." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            selectMore: { ja: "2つ以上のオブジェクトを選択してください。", en: "Please select two or more objects." },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    // =========================================
    // 重ね順 / Stacking order
    // =========================================

    var SORT_KEY_SCALE = 1000; /* 座標を 0.001pt 単位の整数にして並べ替えキーにする / coordinates become integer keys in 0.001pt */
    var SORT_KEY_DIGITS = 12;  /* 並べ替えキーの桁数 / digits of a sort key */

    /**
     * 先頭のオブジェクトの前面へ、末尾から順に移す（配列の先頭ほど背面になる）
     * @param {PageItem[]} backToFrontItems - 背面から前面の順に並べたオブジェクト
     * @returns {void}
     */
    function restackItems(backToFrontItems) {
        var baseItem = backToFrontItems[0];
        for (var j = backToFrontItems.length - 1; j >= 1; j--) {
            backToFrontItems[j].move(baseItem, ElementPlacement.PLACEBEFORE);
        }
    }

    /**
     * 各オブジェクトのすぐ背面に空のグループを目印として置く。
     * 選択外のオブジェクトとの前後や、レイヤー・グループの位置までキャンセルで戻せるようにする
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {GroupItem[]} 目印（targetItems と同じ順）
     */
    function placeStackMarkers(targetItems) {
        var stackMarkers = [];
        for (var i = 0; i < targetItems.length; i++) {
            var stackMarker = targetItems[i].parent.groupItems.add();
            stackMarker.move(targetItems[i], ElementPlacement.PLACEAFTER);
            stackMarkers.push(stackMarker);
        }
        return stackMarkers;
    }

    /**
     * 各オブジェクトを目印のすぐ前面へ戻す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {GroupItem[]} stackMarkers - placeStackMarkers() の目印
     * @returns {void}
     */
    function restoreFromStackMarkers(targetItems, stackMarkers) {
        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].move(stackMarkers[i], ElementPlacement.PLACEBEFORE);
        }
    }

    /**
     * 目印を削除する
     * @param {GroupItem[]} stackMarkers - placeStackMarkers() の目印
     * @returns {void}
     */
    function removeStackMarkers(stackMarkers) {
        for (var i = 0; i < stackMarkers.length; i++) {
            stackMarkers[i].remove();
        }
    }

    /**
     * 選択を JavaScript の配列に写す（選択は前面から背面の順）
     * @param {Object} docSelection - ドキュメントの選択
     * @returns {PageItem[]} 前面から背面の順の配列
     */
    function selectionToArray(docSelection) {
        var selectedItems = [];
        for (var i = 0; i < docSelection.length; i++) {
            selectedItems.push(docSelection[i]);
        }
        return selectedItems;
    }

    /**
     * 配列の順序をシャッフルする（Fisher–Yates、直接書き換える）
     * @param {PageItem[]} targetItems - 対象の配列
     * @returns {void}
     */
    function shuffleItems(targetItems) {
        for (var i = targetItems.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapItem = targetItems[i];
            targetItems[i] = targetItems[j];
            targetItems[j] = swapItem;
        }
    }

    /**
     * 数値を指定の桁数になるよう左をゼロで埋める
     * @param {number} value - 0以上の整数
     * @param {number} digits - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function padWithZeros(value, digits) {
        var paddedText = String(value);
        while (paddedText.length < digits) paddedText = "0" + paddedText;
        return paddedText;
    }

    /**
     * 左端または上端の値が小さい順に並べた配列を返す。
     * 比較関数つきの sort() は並びが狂うことがあるため、ゼロ埋めした文字列キーを引数なしの sort() で並べる
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {number} boundsIndex - geometricBounds の添字（0 = 左端、1 = 上端）
     * @returns {PageItem[]} 並べ替えた新しい配列
     */
    function sortItemsByBounds(targetItems, boundsIndex) {
        var edgeValues = [];
        var minValue = Infinity;
        for (var i = 0; i < targetItems.length; i++) {
            edgeValues.push(targetItems[i].geometricBounds[boundsIndex]);
            if (edgeValues[i] < minValue) minValue = edgeValues[i];
        }
        /* 最小値を0にずらして負の座標もそのまま並ぶようにし、末尾に元の番号を付ける
           Shift so the minimum is 0 (negative coordinates sort correctly) and append the original index */
        var sortKeys = [];
        for (var j = 0; j < targetItems.length; j++) {
            var scaledValue = Math.round((edgeValues[j] - minValue) * SORT_KEY_SCALE);
            sortKeys.push(padWithZeros(scaledValue, SORT_KEY_DIGITS) + ":" + j);
        }
        sortKeys.sort();
        var sortedItems = [];
        for (var k = 0; k < sortKeys.length; k++) {
            sortedItems.push(targetItems[Number(sortKeys[k].split(":")[1])]);
        }
        return sortedItems;
    }

    /**
     * 基準と並び順から、背面から前面の順に並べた配列を作る
     * @param {PageItem[]} originalItems - ダイアログを開いたときの選択（前面から背面の順）
     * @param {string} sortBasis - "zOrder" / "xAxis" / "yAxis"
     * @param {string} sortOrder - "asc" / "desc" / "random"
     * @returns {PageItem[]} 背面から前面の順の配列
     */
    function buildStackOrder(originalItems, sortBasis, sortOrder) {
        var orderedItems = originalItems.slice();
        if (sortOrder === "random") {
            shuffleItems(orderedItems);
            return orderedItems;
        }
        var isDescending = (sortOrder === "desc");
        if (sortBasis === "zOrder") {
            /* 元の配列は前面から背面の順なので、昇順（元のまま）は反転して背面からにする
               The original array runs front to back, so Ascending (unchanged) reverses it to run back to front */
            return isDescending ? orderedItems : orderedItems.reverse();
        }
        var isYAxis = (sortBasis === "yAxis");
        orderedItems = sortItemsByBounds(orderedItems, isYAxis ? 1 : 0);
        /* Y は上ほど値が大きいので、X と昇順・降順を逆にする / Y grows upward, so flip the direction for Y */
        return (isYAxis !== isDescending) ? orderedItems.reverse() : orderedItems;
    }

    /**
     * 有効な選択を返す。ドキュメントが無い、または2つ未満なら知らせて null
     * @returns {PageItem[]|null} 前面から背面の順の選択オブジェクト
     */
    function getValidSelection() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return null;
        }
        var docSelection = app.activeDocument.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / A text selection returns a TextRange, which has no [0] */
        if (!docSelection || docSelection.typename === "TextRange" || docSelection.length < 2) {
            alert(getLabel("alert.selectMore"));
            return null;
        }
        return selectionToArray(docSelection);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンのパネルを作る
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} titlePath - パネル名の LABELS パス
     * @param {string[]} radioKeys - radio / tooltip の LABELS キー（ラジオボタンの .optionKey にも入れる）
     * @returns {RadioButton[]} 作ったラジオボタン（先頭を選択済み）
     */
    function addRadioPanel(parentWindow, titlePath, radioKeys) {
        var radioPanel = parentWindow.add("panel", undefined, getLabel(titlePath));
        setupPanel(radioPanel, RADIO_SPACING);
        var panelRadios = [];
        for (var i = 0; i < radioKeys.length; i++) {
            var panelRadio = radioPanel.add("radiobutton", undefined, getLabel("radio." + radioKeys[i]));
            panelRadio.helpTip = getLabel("tooltip." + radioKeys[i]);
            panelRadio.optionKey = radioKeys[i];
            panelRadios.push(panelRadio);
        }
        panelRadios[0].value = true;
        return panelRadios;
    }

    /**
     * 選ばれているラジオボタンのキーを返す
     * @param {RadioButton[]} panelRadios - addRadioPanel() で作ったラジオボタン
     * @returns {string} 選ばれているラジオボタンの .optionKey（無ければ先頭のキー）
     */
    function getCheckedKey(panelRadios) {
        for (var i = 0; i < panelRadios.length; i++) {
            if (panelRadios[i].value) return panelRadios[i].optionKey;
        }
        return panelRadios[0].optionKey;
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
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを出し、ラジオボタンを押すたびに重ね順を組み直す
     * @param {PageItem[]} originalItems - ダイアログを開いたときの選択（前面から背面の順）
     * @returns {number} show() の戻り値（OK は 1）
     */
    function showSortDialog(originalItems) {
        var sortDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(sortDialog);

        var basisRadios = addRadioPanel(sortDialog, "panel.sortBasis", ["zOrder", "xAxis", "yAxis"]);
        var orderRadios = addRadioPanel(sortDialog, "panel.sortOrder", ["asc", "desc", "random"]);

        /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
        var buttonRow = addButtonRow(sortDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 選んだ基準と並び順で重ね順を組み直す（ラジオボタンを押すたびに呼ぶ）
         * @returns {void}
         */
        function applySelectedSort() {
            restackItems(buildStackOrder(originalItems, getCheckedKey(basisRadios), getCheckedKey(orderRadios)));
            app.redraw();
        }

        var allRadios = basisRadios.concat(orderRadios);
        for (var i = 0; i < allRadios.length; i++) {
            allRadios[i].onClick = applySelectedSort;
        }

        btnOK.onClick = function () {
            sortDialog.close(1);
        };
        btnCancel.onClick = function () {
            sortDialog.close(0);
        };

        addKeyShortcuts(sortDialog, {
            "Z": basisRadios[0],
            "X": basisRadios[1],
            "Y": basisRadios[2],
            "A": orderRadios[0],
            "D": orderRadios[1],
            "R": orderRadios[2]
        }, { showInTip: true });

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(sortDialog, SCRIPT_NAME);
        return sortDialog.show();
    }

    /**
     * 選択を確かめてからダイアログを出す。OK 以外（キャンセル・Esc・閉じるボタン）で閉じたら元の位置へ戻す
     * @returns {void}
     */
    function main() {
        var originalItems = getValidSelection();
        if (!originalItems) return;
        var stackMarkers = placeStackMarkers(originalItems);
        try {
            if (showSortDialog(originalItems) !== 1) {
                restoreFromStackMarkers(originalItems, stackMarkers);
                app.redraw();
            }
        } finally {
            /* 途中で失敗しても目印を残さない / Never leave the markers behind, even on failure */
            removeStackMarkers(stackMarkers);
        }
    }

    main();

})();
