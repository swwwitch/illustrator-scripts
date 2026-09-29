#target illustrator
#targetengine "ConvertFontInfoEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストを、そのテキストで使われているフォント情報の文字列に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertFontInfo.md

### Overview

Replaces the selected text with a string describing the font information it uses.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertFontInfo.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertFontInfo";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertFontInfo.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertFontInfo.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
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
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
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
        dialog: {
            title: { ja: "フォント情報に変換", en: "Convert Font Info" }
        },
        panel: {
            title: { ja: "変換形式", en: "Conversion Format" }
        },
        radio: {
            family: { ja: "フォント名", en: "Font Family" },
            style: { ja: "スタイル", en: "Style" },
            familyStyle: { ja: "フォント名＋スタイル", en: "Font Family + Style" },
            postscript: { ja: "PostScript 名", en: "PostScript Name" },
            fullNameSize: { ja: "フルネーム＋サイズ", en: "Full Name + Size" },
            detail: { ja: "詳細", en: "Labeled Detail" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        help: {
            fontFamily: { ja: "ショートカット: F", en: "Shortcut: F" },
            style: { ja: "ショートカット: S", en: "Shortcut: S" },
            fontFamilyStyle: { ja: "ショートカット: B", en: "Shortcut: B" },
            postScriptName: { ja: "ショートカット: P", en: "Shortcut: P" },
            fullNameSize: { ja: "ショートカット: M", en: "Shortcut: M" },
            detailLines: { ja: "ショートカット: D", en: "Shortcut: D" }
        },
        detail: {
            familyLabel: { ja: "フォント名", en: "Font Family" },
            styleLabel: { ja: "スタイル", en: "Style" },
            nameLabel: { ja: "PostScript 名", en: "PostScript Name" }
        }
    };

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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * pt のフォントサイズを文字サイズの単位（text/units）に換算し、小数第2位までの単位付き文字列にする
     * @param {number} sizePt - フォントサイズ（pt）
     * @returns {string} 「12 pt」「17 Q」のような文字列
     */
    function getFontSizeWithUnit(sizePt) {
        var sizeUnit = getUnitInfo("text/units");
        var convertedSize = sizePt / sizeUnit.pointsPerUnit;
        return (Math.round(convertedSize * 100) / 100) + " " + sizeUnit.label;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // UI設定 / UI Settings
    // =========================================
    var FORMAT_ITEMS = [
        { labelKey: 'radio.family', format: 'name', helpKey: 'help.fontFamily', shortcut: 'F', showPreview: true },
        { labelKey: 'radio.style', format: 'style', helpKey: 'help.style', shortcut: 'S', showPreview: true },
        { labelKey: 'radio.familyStyle', format: 'family+style', helpKey: 'help.fontFamilyStyle', shortcut: 'B', showPreview: true },
        { labelKey: 'radio.postscript', format: 'postscript', helpKey: 'help.postScriptName', shortcut: 'P', showPreview: true },
        { labelKey: 'radio.fullNameSize', format: 'fullName+size', helpKey: 'help.fullNameSize', shortcut: 'M', showPreview: true },
        { labelKey: 'radio.detail', format: 'detailLines', helpKey: 'help.detailLines', shortcut: 'D', showPreview: false }
    ];

    // 文字ツールで文字を選択すると、app.selection は配列ではなく TextRange を返す
    // （配列でなく .story を持つオブジェクトを TextRange とみなす）
    function isTextRangeSelection(selection) {
        return selection && !(selection instanceof Array) && selection.story;
    }

    // TextRange（文字選択）から、その文字が属する親テキストフレームを取得
    // contents はストーリー単位なので、連結テキストでも先頭フレーム1つで足りる
    function textFramesFromTextRange(textRange) {
        var storyTextFrames = textRange.story && textRange.story.textFrames;
        return (storyTextFrames && storyTextFrames.length >= 1) ? [storyTextFrames[0]] : [];
    }

    // 選択内容から対象テキストフレームの配列を解決する（doc.selection は書き換えない）
    // ・オブジェクト選択：選択内のテキストフレーム
    // ・文字ツールでの部分選択（TextRange）：その文字の親テキストフレーム
    function resolveTargetTextFrames(doc) {
        var selection = doc.selection;
        if (!selection) return [];

        if (isTextRangeSelection(selection)) {
            return textFramesFromTextRange(selection);
        }
        if (!(selection instanceof Array)) return [];

        var textFrames = [];
        for (var selectedIndex = 0; selectedIndex < selection.length; selectedIndex++) {
            var selectedItem = selection[selectedIndex];
            if (selectedItem.typename === "TextFrame") {
                textFrames.push(selectedItem);
            } else if (isTextRangeSelection(selectedItem)) {
                // 配列内に TextRange が入るケースにも対応
                var parentFrames = textFramesFromTextRange(selectedItem);
                for (var parentIndex = 0; parentIndex < parentFrames.length; parentIndex++) {
                    textFrames.push(parentFrames[parentIndex]);
                }
            }
        }
        return textFrames;
    }

    function collectTextFrameInfo(textFrame) {
        try {
            var textRange = textFrame.textRange;
            return {
                textFrame: textFrame,
                font: textRange.characterAttributes.textFont,
                originalSize: textRange.characterAttributes.size,
                originalJustification: textRange.paragraphAttributes.justification,
                originalText: textFrame.contents
            };
        } catch (e) {
            return null;
        }
    }

    function collectSelectedTextFrameInfos(selectedItems) {
        var textFrameInfos = [];

        // 配列風オブジェクト以外（TextRange など）は対象外
        if (!selectedItems || typeof selectedItems.length !== 'number') {
            return textFrameInfos;
        }

        for (var selectedIndex = 0; selectedIndex < selectedItems.length; selectedIndex++) {
            var selectedItem = selectedItems[selectedIndex];
            if (selectedItem.typename !== "TextFrame") continue;

            var textFrameInfo = collectTextFrameInfo(selectedItem);
            if (textFrameInfo) {
                textFrameInfos.push(textFrameInfo);
            }
        }

        return textFrameInfos;
    }

    function restoreTextFrameInfo(originalTextFrameInfo) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            textFrame.contents = originalTextFrameInfo.originalText;
            var textRange = textFrame.textRange;
            textRange.characterAttributes.textFont = originalTextFrameInfo.font;
            textRange.characterAttributes.size = originalTextFrameInfo.originalSize;
            textRange.paragraphAttributes.justification = originalTextFrameInfo.originalJustification;
        } catch (e) {
            return false;
        }
        return true;
    }

    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var targetTextFrames = resolveTargetTextFrames(doc);
        if (targetTextFrames.length === 0 || targetTextFrames.length >= 1000) return;

        var originalTextFrameInfos = collectSelectedTextFrameInfos(targetTextFrames);
        if (originalTextFrameInfos.length === 0) return;

        showDialog(originalTextFrameInfos);
    }

    function getConvertedFontInfoText(originalTextFrameInfo, format) {
        var sourceFont = originalTextFrameInfo.font;
        var sourceFontSize = originalTextFrameInfo.originalSize;

        if (!sourceFont) return "";

        switch (format) {
            case "style":
                return sourceFont.style;
            case "family+style":
                return sourceFont.family + " " + sourceFont.style;
            case "postscript":
                return sourceFont.name;
            case "fullName+size":
                var fontSizeText = getFontSizeWithUnit(sourceFontSize);
                return (sourceFont.fullName && sourceFont.fullName !== "")
                    ? sourceFont.fullName + "\t" + fontSizeText
                    : sourceFont.family + " " + sourceFont.style + "\t" + fontSizeText;
            default:
                return sourceFont.family;
        }
    }

    function buildDetailLineItems(originalTextFrameInfo) {
        var sourceFont = originalTextFrameInfo.font;

        return [
            { label: getLabel('detail.familyLabel'), value: sourceFont.family },
            { label: getLabel('detail.styleLabel'), value: sourceFont.style },
            { label: getLabel('detail.nameLabel'), value: sourceFont.name }
        ];
    }

    function findTextFontByName(fontName) {
        var textFonts = app.textFonts;
        for (var fontIndex = 0; fontIndex < textFonts.length; fontIndex++) {
            if (textFonts[fontIndex].name === fontName) {
                return textFonts[fontIndex];
            }
        }
        return null;
    }

    // 詳細表示のラベル用フォント（セッション中変わらないので一度だけ走査してキャッシュ）
    var cachedDetailLabelFont; // 未取得は undefined、取得済みで未発見は null
    function getDetailLabelFont() {
        if (cachedDetailLabelFont === undefined) {
            cachedDetailLabelFont = findTextFontByName("HiraginoSans-W3");
        }
        return cachedDetailLabelFont;
    }

    function buildDetailLinesText(originalTextFrameInfo) {
        var lineBreak = String.fromCharCode(13);
        var detailLineItems = buildDetailLineItems(originalTextFrameInfo);
        var detailTextParts = [];

        for (var detailItemIndex = 0; detailItemIndex < detailLineItems.length; detailItemIndex++) {
            detailTextParts.push(detailLineItems[detailItemIndex].label);
            detailTextParts.push(detailLineItems[detailItemIndex].value);
        }

        return detailTextParts.join(lineBreak);
    }

    function applyDetailFontInfoPreview(originalTextFrameInfo, detailLabelFont) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            var sourceFont = originalTextFrameInfo.font;
            var sourceFontSize = originalTextFrameInfo.originalSize;

            textFrame.contents = buildDetailLinesText(originalTextFrameInfo);
            textFrame.textRange.paragraphAttributes.justification = Justification.LEFT;

            var previewLines = textFrame.textRange.lines;
            if (previewLines.length > 0 && detailLabelFont) {
                for (var lineIndex = 0; lineIndex < previewLines.length; lineIndex++) {
                    var isLabelLine = (lineIndex % 2 === 0);
                    var lineAttributes = previewLines[lineIndex].characterAttributes;
                    lineAttributes.textFont = isLabelLine ? detailLabelFont : sourceFont;
                    lineAttributes.size = isLabelLine ? 10 : sourceFontSize;
                }
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    function applySimpleFontInfoPreview(originalTextFrameInfo, format) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            var sourceFont = originalTextFrameInfo.font;
            var sourceFontSize = originalTextFrameInfo.originalSize;
            var textRange = textFrame.textRange;

            textRange.characterAttributes.textFont = sourceFont;
            textRange.characterAttributes.size = sourceFontSize;
            textFrame.contents = getConvertedFontInfoText(originalTextFrameInfo, format);
        } catch (e) {
            return false;
        }
        return true;
    }

    function updateFontInfoPreview(originalTextFrameInfos, format) {
        var isDetailLinesFormat = (format === "detailLines");
        var detailLabelFont = isDetailLinesFormat ? getDetailLabelFont() : null;

        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            if (isDetailLinesFormat) {
                applyDetailFontInfoPreview(originalTextFrameInfos[textFrameIndex], detailLabelFont);
            } else {
                applySimpleFontInfoPreview(originalTextFrameInfos[textFrameIndex], format);
            }
        }

        app.redraw();
    }

    function restoreOriginalText(originalTextFrameInfos) {
        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            restoreTextFrameInfo(originalTextFrameInfos[textFrameIndex]);
        }

        app.redraw();
    }

    function showDialog(originalTextFrameInfos) {
        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);

        var dialogContainer = dialog.add("group");
        dialogContainer.orientation = "column";
        dialogContainer.alignChildren = "left";
        dialogContainer.alignment = "fill";

        buildFormatPanel(dialog, dialogContainer, originalTextFrameInfos);

        /* キャンセル／OK のボタン行。OK をデフォルトボタンにする / Cancel/OK row; OK is the default button */
        var buttonRow = addButtonRow(dialogContainer);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
        dialog.defaultElement = btnOK;

        // 選択中の形式はライブプレビューで既にドキュメントへ適用済み。
        // OK時は再適用せずそのまま確定し（二重適用によるクラッシュを回避）、
        // キャンセル時のみ開始時の内容・フォント・サイズ・行揃えに戻す。
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();
        if (dialogResult !== 1) {
            restoreOriginalText(originalTextFrameInfos);
        }
    }

    // 変換形式パネル（ラジオ生成・選択・キー操作）を組み立て、選択中フォーマットを返す関数を返す
    function buildFormatPanel(dialog, dialogContainer, originalTextFrameInfos) {
        var DEFAULT_FORMAT = 'family+style';
        var previewSourceInfo = originalTextFrameInfos[0];
        var radioColumnWidth = 180;
        var formatRadioButtons = [];

        var formatPanel = dialogContainer.add("panel", undefined, getLabel('panel.title'));
        formatPanel.orientation = "column";
        formatPanel.alignChildren = "left";
        formatPanel.alignment = "fill";
        formatPanel.margins = [15, 20, 15, 10];

        var formatRowsGroup = formatPanel.add("group");
        formatRowsGroup.orientation = "column";
        formatRowsGroup.alignChildren = ["fill", "top"];
        formatRowsGroup.spacing = 6;

        function addFormatRow(labelKey, previewText, helpTipKey) {
            var formatRow = formatRowsGroup.add("group");
            formatRow.orientation = "row";
            formatRow.alignChildren = ["left", "center"];

            var radioColumn = formatRow.add("group");
            radioColumn.preferredSize.width = radioColumnWidth;
            var radioButton = radioColumn.add("radiobutton", undefined, getLabel(labelKey));

            if (helpTipKey) {
                radioButton.helpTip = getLabel(helpTipKey);
            }

            if (previewText) {
                var previewStaticText = formatRow.add("statictext", undefined, previewText);
                previewStaticText.helpTip = previewText;
            }

            return radioButton;
        }

        function selectFormat(format) {
            for (var itemIndex = 0; itemIndex < FORMAT_ITEMS.length; itemIndex++) {
                formatRadioButtons[itemIndex].value = (FORMAT_ITEMS[itemIndex].format === format);
            }
            updateFontInfoPreview(originalTextFrameInfos, format);
        }

        for (var formatIndex = 0; formatIndex < FORMAT_ITEMS.length; formatIndex++) {
            (function (formatItem, itemIndex) {
                var previewText = formatItem.showPreview ? getConvertedFontInfoText(previewSourceInfo, formatItem.format) : "";
                var radioButton = addFormatRow(formatItem.labelKey, previewText, formatItem.helpKey);
                formatRadioButtons[itemIndex] = radioButton;
                radioButton.onClick = function () {
                    selectFormat(formatItem.format);
                };
            })(FORMAT_ITEMS[formatIndex], formatIndex);
        }

        /* 各行のキーで形式を選ぶ（onClick が selectFormat を通る）/ Pick a format by its key (onClick goes through selectFormat) */
        var formatShortcutMap = {};
        for (var shortcutIndex = 0; shortcutIndex < FORMAT_ITEMS.length; shortcutIndex++) {
            formatShortcutMap[FORMAT_ITEMS[shortcutIndex].shortcut] = formatRadioButtons[shortcutIndex];
        }
        addKeyShortcuts(dialog, formatShortcutMap);

        selectFormat(DEFAULT_FORMAT);
    }

    main();

})();
