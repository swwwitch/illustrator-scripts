#target illustrator
#targetengine "ImageLinkManagerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

配置画像に対する「埋め込み」「埋め込み解除」「リセット」「ケイ線」「リンク」を、1つのダイアログでまとめて扱います。
ダイアログ上部のモードで処理を切り替え、対応するパネルだけが有効になります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageLinkManager.md

### Overview

Handles Embed, Unembed, Reset, Stroke and Relink for placed images from a single dialog.
The mode selector at the top switches the operation, and only the matching panel stays enabled.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageLinkManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImageLinkManager";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageLinkManager.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageLinkManager.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
      dialogTitle: {
        ja: "まるっと配置画像 " + SCRIPT_VERSION,
        en: "Image Link Manager " + SCRIPT_VERSION
      },

      // ===== Mode (top) =====
      modeEmbed: { ja: "埋め込み", en: "Embed" },
      modeRelease: { ja: "解除", en: "Unembed" },
      modeKeisen: { ja: "ケイ", en: "Stroke" },
      modeReset: { ja: "リセット", en: "Reset" },
      modeLink: { ja: "リンク", en: "Link" },

      // ===== Tooltips =====
      tooltip: {
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
      tipModeEmbed: { ja: "配置画像をドキュメントに埋め込みます。", en: "Embeds the placed images into the document." },
      tipModeRelease: { ja: "埋め込み画像をファイルへ書き出し、リンクに戻します。", en: "Writes embedded images out to files and links to them instead." },
      tipModeKeisen: { ja: "配置画像にケイ線を追加します。", en: "Adds a rule around the placed images." },
      tipModeReset: { ja: "配置画像に掛かっている変形（回転・シアーなど）を元に戻します。", en: "Undoes the transforms applied to placed images, such as rotation and shear." },
      tipModeLink: { ja: "リンクの更新やさしかえを行います。", en: "Updates or replaces the links." },
      tipEmbedSelection: { ja: "選択している配置画像だけを埋め込みます。", en: "Embeds only the selected placed images." },
      tipEmbedAll: { ja: "ドキュメント内のすべての配置画像を埋め込みます。", en: "Embeds every placed image in the document." },
      tipReleaseSelection: { ja: "選択している埋め込み画像だけをリンクに戻します。", en: "Unembeds only the selected images." },
      tipReleaseAll: { ja: "ドキュメント内のすべての埋め込み画像をリンクに戻します。", en: "Unembeds every embedded image in the document." },
      tipKeisenStrokeOnly: { ja: "配置画像にそのまま線を追加します。画像の形は変わりません。", en: "Adds a stroke straight to the placed image, leaving its shape alone." },
      tipKeisenClipGroup: { ja: "配置画像を長方形でクリップしてから、その長方形に線を追加します。角丸にできます。", en: "Clips the image with a rectangle and strokes that rectangle, so the corners can be rounded." },
      tipKeisenRoundCorners: { ja: "クリップした長方形の角を丸めます。半径は右の欄で指定します。", en: "Rounds the corners of the clipping rectangle. The field on the right sets the radius." },
      tipLinkUpdate: { ja: "リンク切れや更新のある画像を、いまのリンク先で読み直します。", en: "Reloads the images from their current link paths." },
      tipLinkRelinkAll: { ja: "すべてのリンクを、選んだファイルへ張り替えます。", en: "Relinks every image to a file you choose." },
      tipResetReplace: { ja: "画像を同じ位置に置き直して、変形を落とします。", en: "Re-places the image at the same spot, dropping its transforms." },
      tipResetRotate: { ja: "回転を元に戻します。", en: "Clears the rotation." },
      tipResetSkew: { ja: "シアー（傾き）を元に戻します。", en: "Clears the shear." },
      tipResetRatio: { ja: "縦横比を元に戻します。", en: "Restores the original aspect ratio." },
      tipResetFlip: { ja: "反転を元に戻します。", en: "Clears the flip." },
      tipResetScale: { ja: "拡大・縮小率を右の欄の値にそろえます。", en: "Sets the scale to the value in the field on the right." },
      tipScale: { ja: "そろえる拡大・縮小率（％）です。", en: "The scale, in percent, every image is set to." },

      // ===== Panels =====
      panelEmbed: { ja: "埋め込み", en: "Embed" },
      panelRelease: { ja: "解除", en: "Unembed" },
      panelKeisen: { ja: "ケイ", en: "Stroke" },
      panelReset: { ja: "リセット", en: "Reset" },
      panelLink: { ja: "リンク", en: "Link" },
      linkUpdate: { ja: "リンクを更新", en: "Update links" },
      linkRelinkAll: { ja: "すべてさしかえ", en: "Relink all" },

      // ===== Embed options =====
      embedSelection: { ja: "選択した配置画像のみ", en: "Selected placed images only" },
      embedAll: { ja: "すべての配置画像", en: "All placed images" },

      // ===== Release options =====
      releaseSelection: { ja: "選択した画像のみ", en: "Selected images only" },
      releaseAll: { ja: "すべての画像", en: "All images" },

      // ===== Keisen options =====
      keisenStrokeOnly: { ja: "ケイ線のみを追加", en: "Add rules only" },
      keisenClipGroup: { ja: "クリップグループ", en: "Clipping group" },
      keisenRoundCorners: { ja: "角丸", en: "Round corners" },

      // ===== Reset options =====
      resetReplace: { ja: "再配置", en: "Re-place" },
      resetRotate: { ja: "回転", en: "Rotation" },
      resetSkew: { ja: "シアー", en: "Shear" },
      resetRatio: { ja: "縦横比", en: "Aspect ratio" },
      resetFlip: { ja: "反転", en: "Flip" },
      resetScale: { ja: "スケール", en: "Scale" },

      // ===== Buttons =====
      ok: { ja: "OK", en: "OK" },
      cancel: { ja: "キャンセル", en: "Cancel" },

      // ===== Alerts / Errors =====
      alertSelectObject: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
      alertKeisenError: { ja: "ケイ線の処理中にエラーが発生しました。\n", en: "An error occurred while processing rules.\n" },
      alertClipCannot: {
        ja: "クリップグループ化できる選択ではありません。\n（例：画像1つ、または 画像1つ+パス1つ、または 複数オブジェクト+パス）",
        en: "The selection cannot be converted into a clipping group.\n(Example: one image, one image + one path, or multiple objects + one path)"
      },
      errInvalidRoundRadius: {
        ja: "角丸の値が不正です。0以上の数値を入力してください。",
        en: "Invalid round radius. Please enter a number 0 or greater."
      },
      alertLinkedCannotUnembed: {
        ja: "リンク画像は解除できません。\n埋め込み画像を選択して実行してください。",
        en: "Linked images cannot be unembedded.\nPlease select embedded images and run again."
      },
      alertDocNotSaved: {
        ja: "Illustratorドキュメントが保存されていないため、「埋め込み解除」は実行できません。\n先にドキュメントを保存してください。",
        en: "The Illustrator document is not saved. Unembed cannot be executed.\nPlease save the document first."
      },
      // ===== Link dialogs / alerts =====
      dialogSelectReplaceFile: { ja: "置換するファイルを選択してください", en: "Select a file to relink to" },
      alertSelectPlacedItem: { ja: "配置画像（リンク画像）を1つ選択してください。", en: "Please select one placed (linked) image." },
      alertSelectedNotPlacedItem: { ja: "選択されたアイテムは配置画像ではありません。", en: "The selected item is not a placed image." },
      alertFileSelectionCanceled: { ja: "ファイルの選択がキャンセルされました。", en: "File selection was canceled." },
      alertEmbeddedDoneSuffix: { ja: "個の画像を埋め込みました", en: " image(s) embedded." },
      alertReleasedDoneSuffix: { ja: "個の画像を解除しました", en: " image(s) unembedded." },
      alertLinkedCreatedPrefix: { ja: "（リンク作成: ", en: "(Links created: " },
      alertLinkedCreatedSuffix: { ja: "）", en: ")" },
      alertLinkedCreateZeroNote: {
        ja: "\n\n※リンク作成が確認できませんでした。書き出し先フォルダや権限、ドキュメントの保存状態をご確認ください。",
        en: "\n\nNote: No links were detected as created. Please check the export folder, permissions, and whether the document is saved."
      },
      alertUnembedFailed: {
        ja: "解除できませんでした。\n\n",
        en: "Unembed failed.\n\n"
      }
    };

    /**
     * ===== Units (from rulerType) =====
     */
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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

    function showDialog() {
        var dlg = new Window('dialog', LABELS.dialogTitle[uiLang]);
        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];
        dlg.margins = [15, 20, 15, 10];

        // ===== モード選択（上部）/ Mode (top) =====
        var modeWrap = dlg.add('group');
        modeWrap.orientation = 'row';
        modeWrap.alignment = 'center';

        var modeGroup = modeWrap.add('group');
        modeGroup.orientation = 'row';
        modeGroup.alignChildren = ['left', 'center'];

        var rbModeEmbed = modeGroup.add('radiobutton', undefined, getLabel('modeEmbed'));
        rbModeEmbed.helpTip = getLabel('tipModeEmbed');
        var rbModeRelease = modeGroup.add('radiobutton', undefined, getLabel('modeRelease'));
        rbModeRelease.helpTip = getLabel('tipModeRelease');
        var rbModeKeisen = modeGroup.add('radiobutton', undefined, getLabel('modeKeisen'));
        rbModeKeisen.helpTip = getLabel('tipModeKeisen');
        var rbModeReset = modeGroup.add('radiobutton', undefined, getLabel('modeReset'));
        rbModeReset.helpTip = getLabel('tipModeReset');
        var rbModeLink = modeGroup.add('radiobutton', undefined, getLabel('modeLink'));
        rbModeLink.helpTip = getLabel('tipModeLink');

        rbModeEmbed.value = true; // デフォルト

        // ===== 2 columns wrapper =====
        var columns = dlg.add('group');
        columns.orientation = 'row';
        columns.alignChildren = ['fill', 'top'];

        var leftCol = columns.add('group');
        leftCol.orientation = 'column';
        leftCol.alignChildren = ['fill', 'top'];

        var rightCol = columns.add('group');
        rightCol.orientation = 'column';
        rightCol.alignChildren = ['fill', 'top'];

        // ===== 埋め込み / Embed =====
        var panelEmbed = leftCol.add('panel', undefined, getLabel('panelEmbed'));
        panelEmbed.orientation = 'column';
        panelEmbed.alignChildren = 'left';
        panelEmbed.margins = [15, 20, 15, 10];

        var rbEmbedSelection = panelEmbed.add('radiobutton', undefined, getLabel('embedSelection'));
        rbEmbedSelection.helpTip = getLabel('tipEmbedSelection');
        var rbEmbedAll = panelEmbed.add('radiobutton', undefined, getLabel('embedAll'));
        rbEmbedAll.helpTip = getLabel('tipEmbedAll');
        rbEmbedSelection.value = true; // デフォルト

        // 選択がない場合は「すべての配置画像」を自動選択
        var hasSelection = false;
        try {
            hasSelection = (app.documents.length > 0 && app.activeDocument.selection && app.activeDocument.selection.length > 0);
        } catch (e) {
            hasSelection = false;
        }
        if (!hasSelection) {
            rbEmbedSelection.value = false;
            rbEmbedAll.value = true;
        }

        // ===== 解除 / Release =====
        var panelRelease = leftCol.add('panel', undefined, getLabel('panelRelease'));
        panelRelease.orientation = 'column';
        panelRelease.alignChildren = 'left';
        panelRelease.margins = [15, 20, 15, 10];

        var rbReleaseSelection = panelRelease.add('radiobutton', undefined, getLabel('releaseSelection'));
        rbReleaseSelection.helpTip = getLabel('tipReleaseSelection');
        var rbReleaseAll = panelRelease.add('radiobutton', undefined, getLabel('releaseAll'));
        rbReleaseAll.helpTip = getLabel('tipReleaseAll');
        rbReleaseSelection.value = true; // デフォルト

        // 選択がない場合は「すべての画像」を自動選択（解除）
        if (!hasSelection) {
            rbReleaseSelection.value = false;
            rbReleaseAll.value = true;
        }

        // ===== ケイ線 / Rules =====
        var panelKeisen = leftCol.add('panel', undefined, getLabel('panelKeisen'));
        panelKeisen.orientation = 'column';
        panelKeisen.alignChildren = ['left', 'top'];
        panelKeisen.margins = [15, 20, 15, 10];

        var rbKeisenStrokeOnly = panelKeisen.add('radiobutton', undefined, getLabel('keisenStrokeOnly'));
        rbKeisenStrokeOnly.helpTip = getLabel('tipKeisenStrokeOnly');
        var rbKeisenClipGroup = panelKeisen.add('radiobutton', undefined, getLabel('keisenClipGroup'));
        rbKeisenClipGroup.helpTip = getLabel('tipKeisenClipGroup');

        // 選択オブジェクト全体の外接矩形から角丸のデフォルト値を算出
        function calcDefaultRoundRadiusFromSelection(currentSelection) {
            try {
                if (!currentSelection || currentSelection.length === 0) return 10;

                var top = -Infinity, left = Infinity, bottom = Infinity, right = -Infinity;
                for (var i = 0, n = currentSelection.length; i < n; i++) {
                    var b = currentSelection[i].geometricBounds; // [top, left, bottom, right]
                    if (!b || b.length !== 4) continue;
                    if (b[0] > top) top = b[0];
                    if (b[1] < left) left = b[1];
                    if (b[2] < bottom) bottom = b[2];
                    if (b[3] > right) right = b[3];
                }

                if (!isFinite(top) || !isFinite(left) || !isFinite(bottom) || !isFinite(right)) return 10;

                var height = Math.abs(top - bottom);
                var width = Math.abs(right - left);
                var A = height + width;
                var B = A / 2;
                var r = Math.max(0, B / 20);
                return r;
            } catch (e) {
                return 10;
            }
        }

        var defaultRoundRadius = 10;
        try {
            if (app.documents.length > 0) {
                defaultRoundRadius = calcDefaultRoundRadiusFromSelection(app.activeDocument.selection);
            }
        } catch (eSel) {
            defaultRoundRadius = 10;
        }
        defaultRoundRadius = Math.round(defaultRoundRadius * 100) / 100;

        // 角丸（クリップグループ用）
        var roundRow = panelKeisen.add('group');
        roundRow.orientation = 'row';
        roundRow.alignChildren = ['left', 'center'];
        roundRow.margins = [20, 0, 0, 0]; // インデント

        var cbKeisenRoundCorners = roundRow.add('checkbox', undefined, getLabel('keisenRoundCorners'));
        cbKeisenRoundCorners.helpTip = getLabel('tipKeisenRoundCorners');
        cbKeisenRoundCorners.value = true;

        /* ∧∨と入力欄は隙間0で突き合わせる。角丸は0以上 / stepper butts the field; radius is 0 or more */
        var roundFieldGroup = roundRow.add('group');
        roundFieldGroup.orientation = 'row';
        roundFieldGroup.alignChildren = ['left', 'center'];
        roundFieldGroup.spacing = 0;
        roundFieldGroup.margins = 0;
        var etKeisenRoundRadius;
        var roundRadiusStepper = addStepper(roundFieldGroup, function () { return etKeisenRoundRadius; }, { min: 0 });
        etKeisenRoundRadius = roundFieldGroup.add('edittext', undefined, String(defaultRoundRadius));
        etKeisenRoundRadius.stepperGroup = roundRadiusStepper;
        bindSteppedArrowKeys(etKeisenRoundRadius, roundRadiusStepper);
        etKeisenRoundRadius.helpTip = getLabel('tipKeisenRoundCorners');
        etKeisenRoundRadius.characters = 6;
        var stKeisenRoundUnit = roundRow.add('statictext', undefined, getUnitInfo().label);

        // 初期状態
        cbKeisenRoundCorners.enabled = false;
        setFieldEnabled(etKeisenRoundRadius, false);
        stKeisenRoundUnit.enabled = false;

        rbKeisenStrokeOnly.value = true;

        function updateKeisenRoundUI() {
            var clipEnabled = (rbKeisenClipGroup.value === true);
            cbKeisenRoundCorners.enabled = clipEnabled;

            var radiusEnabled = clipEnabled && (cbKeisenRoundCorners.value === true);
            setFieldEnabled(etKeisenRoundRadius, radiusEnabled);
            stKeisenRoundUnit.enabled = radiusEnabled;
        }

        rbKeisenStrokeOnly.onClick = updateKeisenRoundUI;
        rbKeisenClipGroup.onClick = updateKeisenRoundUI;
        cbKeisenRoundCorners.onClick = updateKeisenRoundUI;
        updateKeisenRoundUI();
        try {
            stKeisenRoundUnit.text = getUnitInfo().label;
        } catch (e) { }

        etKeisenRoundRadius.onChange = function () {
            var v = Number(etKeisenRoundRadius.text);
            if (isNaN(v)) return;
            etKeisenRoundRadius.text = String(v);
        };

        // ===== リセット / Reset (right column) =====
        var panelReset = rightCol.add('panel', undefined, getLabel('panelReset'));
        panelReset.orientation = 'column';
        panelReset.alignChildren = ['left', 'top'];
        panelReset.margins = [15, 20, 15, 10];

        // ===== リンク / Link (right column) =====
        var panelLink = rightCol.add('panel', undefined, getLabel('panelLink'));
        panelLink.orientation = 'column';
        panelLink.alignChildren = ['left', 'top'];
        panelLink.margins = [15, 20, 15, 10];

        // リンク：操作（ロジックは追って）
        var gLinkMode = panelLink.add('group');
        gLinkMode.orientation = 'column';
        gLinkMode.alignChildren = ['left', 'center'];

        var rbLinkUpdate = gLinkMode.add('radiobutton', undefined, getLabel('linkUpdate'));
        rbLinkUpdate.helpTip = getLabel('tipLinkUpdate');
        var rbLinkRelinkAll = gLinkMode.add('radiobutton', undefined, getLabel('linkRelinkAll'));
        rbLinkRelinkAll.helpTip = getLabel('tipLinkRelinkAll');

        // デフォルト：リンクを更新
        rbLinkUpdate.value = true;

        // （ResetTransform のUIを右カラムに移植：ロジックは後で接続）
        var cbReplaceReset = panelReset.add('checkbox', undefined, getLabel('resetReplace'));
        cbReplaceReset.helpTip = getLabel('tipResetReplace');
        cbReplaceReset.value = false;

        var cbRotate = panelReset.add('checkbox', undefined, getLabel('resetRotate'));
        cbRotate.helpTip = getLabel('tipResetRotate');
        var cbSkew = panelReset.add('checkbox', undefined, getLabel('resetSkew'));
        cbSkew.helpTip = getLabel('tipResetSkew');
        var cbRatio = panelReset.add('checkbox', undefined, getLabel('resetRatio'));
        cbRatio.helpTip = getLabel('tipResetRatio');
        var cbFlip = panelReset.add('checkbox', undefined, getLabel('resetFlip'));
        cbFlip.helpTip = getLabel('tipResetFlip');
        var cbScale = panelReset.add('checkbox', undefined, getLabel('resetScale'));
        cbScale.helpTip = getLabel('tipResetScale');

        cbRotate.value = true;
        cbSkew.value = true;
        cbRatio.value = true;
        cbFlip.value = true;
        cbScale.value = false;

        // スケール（%）
        var gScale = panelReset.add('group');
        gScale.orientation = 'row';
        gScale.alignChildren = 'center';

        /* ∧∨と入力欄は隙間0で突き合わせる。倍率は1%以上 / stepper butts the field; scale is 1% or more */
        var scaleFieldGroup = gScale.add('group');
        scaleFieldGroup.orientation = 'row';
        scaleFieldGroup.alignChildren = ['left', 'center'];
        scaleFieldGroup.spacing = 0;
        scaleFieldGroup.margins = 0;
        var etScale;
        var scaleStepper = addStepper(scaleFieldGroup, function () { return etScale; }, { min: 1 });
        etScale = scaleFieldGroup.add('edittext', undefined, '100');
        etScale.stepperGroup = scaleStepper;
        bindSteppedArrowKeys(etScale, scaleStepper);
        etScale.helpTip = getLabel('tipScale');
        etScale.characters = 5;
        var stPercent = gScale.add('statictext', undefined, '%');

        /**
         * 入力欄の有効／無効を∧∨ごと切り替える（∧∨は自作描画なので描き直してディム表示を変える）
         * @param {EditText} numberInput - stepperGroup を持つ入力欄
         * @param {boolean} isEnabled - 有効にするなら true
         * @returns {void}
         */
        function setFieldEnabled(numberInput, isEnabled) {
            numberInput.enabled = isEnabled;
            numberInput.stepperGroup.enabled = isEnabled;
            redrawSteppersIn(numberInput.stepperGroup);
        }

        // スケールUIの有効/無効
        setFieldEnabled(etScale, cbScale.value);
        stPercent.enabled = cbScale.value;
        cbScale.onClick = function () {
            var on = cbScale.value;
            setFieldEnabled(etScale, on);
            stPercent.enabled = on;
            if (on) {
                try {
                    etScale.active = true;
                    if (typeof etScale.select === 'function') {
                        etScale.select();
                    } else if (typeof etScale.textselection !== 'undefined') {
                        etScale.textselection = etScale.text;
                    }
                } catch (e) { }
            }
        };

        // 再配置ONのときは他をディム（ただし「スケール」は指定可能）
        function updateResetPanelEnabled() {
            var replaceOn = cbReplaceReset.value;

            // 再配置ONのときは「スケール」を自動ON（デフォルト100%のまま指定可能に）
            if (replaceOn) {
                cbScale.value = true;
            }

            // 再配置ONのときは回転/シアー/縦横比/反転を無効化
            cbRotate.enabled = !replaceOn;
            cbSkew.enabled = !replaceOn;
            cbRatio.enabled = !replaceOn;
            cbFlip.enabled = !replaceOn;

            // スケールは再配置ONでも指定可能
            cbScale.enabled = true;

            // スケール入力欄は cbScale に追従（再配置ONでも可）
            setFieldEnabled(etScale, cbScale.value);
            stPercent.enabled = cbScale.value;
        }
        cbReplaceReset.onClick = function () { updateResetPanelEnabled(); };
        updateResetPanelEnabled();

        // UI：パネルのディム切替
        function updatePanelState() {
            panelEmbed.enabled = rbModeEmbed.value;
            panelRelease.enabled = rbModeRelease.value;
            panelReset.enabled = rbModeReset.value;
            panelKeisen.enabled = rbModeKeisen.value;
            panelLink.enabled = rbModeLink.value;
            /* ∧∨のディム表示を切り替える / update stepper dimming */
            redrawSteppersIn(panelReset);
            redrawSteppersIn(panelKeisen);
        }
        rbModeEmbed.onClick = updatePanelState;
        rbModeRelease.onClick = updatePanelState;
        rbModeReset.onClick = updatePanelState;
        rbModeKeisen.onClick = updatePanelState;
        rbModeLink.onClick = updatePanelState;
        updatePanelState();

        /* モードを 埋め込み(E) / 解除(U) / リセット(R) / ケイ線(S) / リンク(L) で切り替える（数値欄の入力中も効く）
           Switch the mode with E / U / R / S / L (also while a numeric field has focus) */
        addKeyShortcuts(dlg, {
            'E': rbModeEmbed,
            'U': rbModeRelease,
            'R': rbModeReset,
            'S': rbModeKeisen,
            'L': rbModeLink
        }, { numericFields: [etKeisenRoundRadius, etScale] });

        // ===== Buttons (center) =====
        var buttonRow = addButtonRow(dlg, { centered: true });

        var btnCancel = buttonRow.rowGroup.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rowGroup.add('button', undefined, getLabel('ok'), { name: 'ok' });

        btnOK.onClick = function () {
            dlg.close(1);
        };
        btnCancel.onClick = function () {
            dlg.close(0);
        };

        prepareDialogWindow(dlg, SCRIPT_NAME);
        var result = dlg.show();
        if (result !== 1) return null;

        return {
            mode: rbModeEmbed.value ? 'embed' : (rbModeRelease.value ? 'release' : (rbModeReset.value ? 'reset' : (rbModeLink.value ? 'link' : 'keisen'))),
            embedMode: rbEmbedAll.value ? 'all' : 'selection',
            releaseMode: rbReleaseAll.value ? 'all' : 'selection',
            keisenOptions: {
                mode: rbKeisenStrokeOnly.value ? 'strokeOnly' : 'clipGroup',
                roundCorners: (cbKeisenRoundCorners.value === true),
                roundRadius: (function () {
                    var n = parseFloat(etKeisenRoundRadius.text, 10);
                    if (isNaN(n)) n = 0;
                    if (n < 0) n = 0;
                    return n;
                })()
            },
            linkOptions: {
                mode: rbLinkRelinkAll.value ? 'relinkAll' : 'update'
            },
            resetOptions: {
                rotate: cbRotate.value,
                skew: cbSkew.value,
                ratio: cbRatio.value,
                flip: cbFlip.value,
                replaceReset: cbReplaceReset.value,
                scale: cbScale.value,
                scalePercent: (function () {
                    var n = parseFloat(etScale.text, 10);
                    if (isNaN(n)) n = 100;
                    if (n <= 0) n = 100;
                    n = Math.round(n);
                    return n;
                })()
            }
        };
    }

    // ===== Collectors =====
    function collectPlacedItemsFromSelection(selection) {
        var results = [];

        function walk(item) {
            if (!item) return;
            if (item.typename === 'PlacedItem') {
                results.push(item);
                return;
            }
            if (item.typename === 'GroupItem') {
                for (var i = 0; i < item.pageItems.length; i++) {
                    walk(item.pageItems[i]);
                }
            }
        }

        for (var i = 0; i < selection.length; i++) {
            walk(selection[i]);
        }

        return results;
    }

    function collectAllPlacedItems(doc) {
        var results = [];
        var items = doc.placedItems;
        for (var i = 0; i < items.length; i++) {
            results.push(items[i]);
        }
        return results;
    }

    function collectRasterItemsFromSelection(selection) {
        var results = [];

        function isEmbeddedPlacedItem(it) {
            if (!it || it.typename !== 'PlacedItem') return false;
            try {
                return (it.file == null);
            } catch (e) {
                return true;
            }
        }

        function walk(item) {
            if (!item) return;

            if (item.typename === 'RasterItem') {
                results.push(item);
                return;
            }

            if (isEmbeddedPlacedItem(item)) {
                results.push(item);
                return;
            }

            if (item.typename === 'GroupItem') {
                for (var i = 0; i < item.pageItems.length; i++) {
                    walk(item.pageItems[i]);
                }
            }
        }

        for (var i = 0; i < selection.length; i++) {
            walk(selection[i]);
        }

        return results;
    }

    function collectAllRasterItems(doc) {
        var results = [];

        // RasterItem（埋め込み画像）
        try {
            var r = doc.rasterItems;
            for (var i = 0; i < r.length; i++) {
                results.push(r[i]);
            }
        } catch (e) { }

        // 念のため：埋め込み後も PlacedItem のまま残るケース
        try {
            var p = doc.placedItems;
            for (var j = 0; j < p.length; j++) {
                try {
                    if (p[j].file == null) {
                        results.push(p[j]);
                    }
                } catch (e2) {
                    // file 参照で例外の場合も embedded 扱い
                    results.push(p[j]);
                }
            }
        } catch (e3) { }

        return results;
    }

    function selectionHasLinkedPlacedItem(selection) {
        function isLinkedPlacedItem(it) {
            if (!it || it.typename !== 'PlacedItem') return false;
            try {
                return (it.file != null && it.file.exists);
            } catch (e) {
                // file 参照で例外の場合は「リンク」とは断定しない
                return false;
            }
        }

        function walk(item) {
            if (!item) return false;
            if (isLinkedPlacedItem(item)) return true;
            if (item.typename === 'GroupItem') {
                for (var i = 0; i < item.pageItems.length; i++) {
                    if (walk(item.pageItems[i])) return true;
                }
            }
            return false;
        }

        for (var i = 0; i < selection.length; i++) {
            if (walk(selection[i])) return true;
        }
        return false;
    }

    // ===== Reset Transform (placed/raster) =====
    function collectPlacedOrRasterFromSelection(selection) {
        var results = [];
        function walk(item) {
            if (!item) return;
            var itemTypeName = item.typename;
            if (itemTypeName === 'PlacedItem' || itemTypeName === 'RasterItem') {
                results.push(item);
                return;
            }
            if (itemTypeName === 'GroupItem') {
                for (var i = 0; i < item.pageItems.length; i++) {
                    walk(item.pageItems[i]);
                }
            }
        }
        for (var i = 0; i < selection.length; i++) {
            walk(selection[i]);
        }
        return results;
    }

    function applyResetTransformToItems(opts, items) {
        if (!items || items.length === 0) return;

        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        try {
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                if (!it || !it.typename) continue;
                if (it.typename !== 'PlacedItem' && it.typename !== 'RasterItem') continue;
                resetTransformOne(it, it.typename, opts || {});
            }
        } finally {
            app.userInteractionLevel = previousInteractionLevel;
            app.redraw();
        }
    }

    function resetTransformOne(item, objectType, opts) {
        var sign = (objectType === 'RasterItem') ? -1 : 1;

        if (!opts) opts = {};
        if (!opts.rotate && !opts.skew && !opts.scale && !opts.ratio && !opts.flip && !opts.replaceReset) return;

        // Re-place reset (PlacedItem only)
        // ※再配置後にスケール（%）指定がある場合は、その値を適用できるようにする
        if (opts.replaceReset && objectType === 'PlacedItem') {
            var replaced = replacePlacedItemToReset(item); // old item is removed inside
            if (replaced && opts.scale) {
                // 念のため：基準を正規化してから指定%を適用
                try { normalizeScaleOnly(replaced); } catch (eNS) { }
                try { applyUniformScalePercent(replaced, opts.scalePercent); } catch (eSP) { }
            }
            return;
        }

        withBBoxResetAndRecenter(item, function () {
            if (opts.rotate && hasMatrix(item)) {
                cancelRotation(item, sign);
            }

            if (opts.flip && hasMatrix(item)) {
                try {
                    var mNow = item.matrix;
                    var fh = isFlippedHorizontal(mNow);
                    var fv = isFlippedVertical(mNow);
                    if (fh && fv) unflipBoth(item);
                    else if (fh) unflipHorizontal(item);
                    else if (fv) unflipVertical(item);
                } catch (eFlip) { }
            }

            if (opts.ratio) {
                equalizeScaleToMax(item);
            }

            if (opts.scale) {
                normalizeScaleOnly(item);
                applyUniformScalePercent(item, opts.scalePercent);
            }

            if (opts.skew) {
                removeSkewOnly(item);
                // retry
                try {
                    if (hasMatrix(item)) {
                        var m = item.matrix;
                        var d = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
                        if (Math.abs(d.shear) > 1e-6) {
                            removeSkewOnly(item);
                        }
                    }
                } catch (eSkew2) { }
            }
        });

        try {
            app.selection = null;
            app.selection = [item];
            app.executeMenuCommand('AI Reset Bounding Box');
        } catch (eBB) { }
    }

    function withBBoxResetAndRecenter(item, opFn) {
        var tl = item.position;
        var w1 = item.width;
        var h1 = item.height;
        if (typeof opFn === 'function') opFn();
        try {
            app.selection = null;
            app.selection = [item];
            app.executeMenuCommand('AI Reset Bounding Box');
        } catch (e) { }
        recenterToTopLeft(item, tl, w1, h1);
    }

    function recenterToTopLeft(item, tl, w1, h1) {
        var w2 = item.width;
        var h2 = item.height;
        item.position = [tl[0] + w1 / 2 - w2 / 2, tl[1] - h1 / 2 + h2 / 2];
    }

    function hasMatrix(obj) {
        try {
            return !!(obj && obj.matrix && typeof obj.matrix.mValueA !== 'undefined');
        } catch (e) { return false; }
    }

    function normalizeZero(n) { return (n === 0) ? 0 : n; }

    function getRotationAngleDeg(a, b, sign) {
        var ang = Math.atan2(b, a) * 180 / Math.PI;
        ang = (sign < 0) ? -ang : ang;
        return normalizeZero(ang);
    }

    function getRotationMatrixSafe(deg) {
        try { if (app && typeof app.getRotationMatrix === 'function') return app.getRotationMatrix(deg); } catch (e) { }
        var rad = deg * Math.PI / 180.0;
        var cosv = Math.cos(rad), sinv = Math.sin(rad);
        var M = new Matrix();
        M.mValueA = cosv; M.mValueB = sinv; M.mValueC = -sinv; M.mValueD = cosv;
        M.mValueTX = 0; M.mValueTY = 0;
        return M;
    }
    function rotateBy(item, deg) { item.transform(getRotationMatrixSafe(deg)); }

    function cancelRotation(item, sign) {
        if (!hasMatrix(item)) return 0;
        var a = item.matrix.mValueA, b = item.matrix.mValueB;
        var rot = getRotationAngleDeg(a, b, sign);
        rotateBy(item, rot);
        return rot;
    }

    function multiply2x2(a1, b1, c1, d1, a2, b2, c2, d2) { return { a: a1 * a2 + c1 * b2, b: b1 * a2 + d1 * b2, c: a1 * c2 + c1 * d2, d: b1 * c2 + d1 * d2 }; }
    function invert2x2(a, b, c, d) {
        var det = a * d - b * c; if (Math.abs(det) < 1e-8) det = (det < 0 ? -1 : 1) * 1e-8;
        var invDet = 1.0 / det; return { a: d * invDet, b: -b * invDet, c: -c * invDet, d: a * invDet };
    }
    function toMatrix(obj) { var M = new Matrix(); M.mValueA = obj.a; M.mValueB = obj.b; M.mValueC = obj.c; M.mValueD = obj.d; M.mValueTX = 0; M.mValueTY = 0; return M; }

    function decomposeQR(a, b, c, d) {
        var sx = Math.sqrt(a * a + b * b); if (sx === 0) sx = 1e-8;
        var q1x = a / sx, q1y = b / sx;
        var r12 = q1x * c + q1y * d;
        var u2x = c - r12 * q1x, u2y = d - r12 * q1y;
        var sy = Math.sqrt(u2x * u2x + u2y * u2y);
        if (sy === 0) { sy = 1e-8; u2x = -q1y; u2y = q1x; }
        var q2x = u2x / sy, q2y = u2y / sy;
        var shear = r12 / sx;
        return { sx: sx, sy: sy, shear: shear, q1x: q1x, q1y: q1y, q2x: q2x, q2y: q2y };
    }
    function buildFromQR(q1x, q1y, q2x, q2y, sx, sy, shear) {
        var r11 = sx, r12 = shear * sx, r21 = 0, r22 = sy;
        return { a: q1x * r11 + q2x * r21, b: q1y * r11 + q2y * r21, c: q1x * r12 + q2x * r22, d: q1y * r12 + q2y * r22 };
    }
    function applyDeltaToMatch(item, target2x2) {
        var m = item.matrix;
        var cur = { a: m.mValueA, b: m.mValueB, c: m.mValueC, d: m.mValueD };
        var inv = invert2x2(cur.a, cur.b, cur.c, cur.d);
        var delta = multiply2x2(inv.a, inv.b, inv.c, inv.d, target2x2.a, target2x2.b, target2x2.c, target2x2.d);
        item.transform(toMatrix(delta));
    }

    function removeSkewOnly(item) {
        if (!item || !hasMatrix(item)) return;
        var m = item.matrix;
        var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, dec.sx, dec.sy, 0);
        applyDeltaToMatch(item, target);
    }
    function normalizeScaleOnly(item) {
        if (!item || !hasMatrix(item)) return;
        var m = item.matrix;
        var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, 1, 1, dec.shear);
        applyDeltaToMatch(item, target);
    }
    function equalizeScaleToMax(item) {
        if (!item || !hasMatrix(item)) return;
        var m = item.matrix;
        var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
        var u = Math.max(dec.sx, dec.sy);
        var uPercent = Math.round(u * 100); u = uPercent / 100;
        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, u, u, dec.shear);
        applyDeltaToMatch(item, target);
    }
    function applyUniformScalePercent(item, percent) {
        var p = Number(percent); if (!(p > 0)) return;
        try { item.resize(p, p, true, true, true, true, true); } catch (e) { }
    }

    function isFlippedHorizontal(mat) { return mat && mat.mValueA < 0; }
    // Placed/Raster は上下判定が特殊
    function isFlippedVertical(mat) { return mat && mat.mValueD > 0; }
    function unflipHorizontal(item) { item.transform(app.getScaleMatrix(-100, 100), true, true, true, true, true, Transformation.CENTER); }
    function unflipVertical(item) { item.transform(app.getScaleMatrix(100, -100), true, true, true, true, true, Transformation.CENTER); }
    function unflipBoth(item) { item.transform(app.getScaleMatrix(-100, -100), true, true, true, true, true, Transformation.CENTER); }

    function replacePlacedItemToReset(placedItem) {
        if (!placedItem || !placedItem.file) return null;
        try {
            var srcPos = placedItem.position;
            var srcW = placedItem.width;
            var srcH = placedItem.height;
            var srcCX = srcPos[0] + srcW / 2;
            var srcCY = srcPos[1] - srcH / 2;

            var doc = app.activeDocument;
            var newItem = doc.placedItems.add();
            newItem.file = placedItem.file;
            newItem.move(placedItem, ElementPlacement.PLACEBEFORE);

            try { normalizeScaleOnly(newItem); } catch (e) { }

            try {
                app.selection = null;
                app.selection = [newItem];
                app.executeMenuCommand('AI Reset Bounding Box');
            } catch (e2) { }

            try { if (placedItem.name) newItem.name = placedItem.name; } catch (e3) { }
            try { if (typeof placedItem.opacity !== 'undefined') newItem.opacity = placedItem.opacity; } catch (e4) { }
            try { if (typeof placedItem.blendingMode !== 'undefined') newItem.blendingMode = placedItem.blendingMode; } catch (e5) { }

            var newW = newItem.width;
            var newH = newItem.height;
            newItem.position = [srcCX - newW / 2, srcCY + newH / 2];

            try { placedItem.remove(); } catch (e6) { }
            return newItem;
        } catch (e7) { }
        return null;
    }

    // ===== Keisen (Rules) =====
    function applyKeisenToSelection(doc, opts) {
        if (!doc || !doc.selection || doc.selection.length === 0) return;

        var mode = (opts && opts.mode) ? opts.mode : 'strokeOnly';
        var doClipGroup = (mode === 'clipGroup');
        var doRound = (opts && opts.roundCorners === true);
        var roundRadius = (opts && typeof opts.roundRadius !== 'undefined') ? Number(opts.roundRadius) : 0;

        if (doRound) {
            if (isNaN(roundRadius) || roundRadius < 0) {
                throw new Error(getLabel('errInvalidRoundRadius'));
            }
        }

        // クリップグループ化（必要な場合）
        if (doClipGroup) {
            var g = makeClippingFromSelection(doc);
            if (!g) {
                alert(getLabel('alertClipCannot'));
                return;
            }
        }

        // ケイ線のみ
        if (!doClipGroup) {
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            app.executeMenuCommand('Live Outline Object');
            return;
        }

        // クリップグループ後：各グループに適用
        var targets = doc.selection;
        if (targets && targets.length) {
            for (var i = 0; i < targets.length; i++) {
                var gi = targets[i];
                if (!gi || gi.typename !== 'GroupItem' || !gi.clipped) continue;

                if (doRound) {
                    applyRoundCornersLiveEffect(gi, roundRadius);
                }

                // メニューコマンドは単体選択で実行
                doc.selection = [gi];
                app.executeMenuCommand('Adobe New Stroke Shortcut');
                app.executeMenuCommand('Live Pathfinder Exclude');
            }
            // 選択を戻す
            doc.selection = targets;
        }
    }

    function isImageItem(item) {
        return item && (item.typename === 'PlacedItem' || item.typename === 'RasterItem');
    }

    function getFrontmostPath(pathArray) {
        var topPath = null;
        var highestZ = -1;
        for (var i = 0; i < pathArray.length; i++) {
            var item = pathArray[i];
            if (item && item.typename === 'PathItem' && item.zOrderPosition > highestZ) {
                highestZ = item.zOrderPosition;
                topPath = item;
            }
        }
        return topPath;
    }

    // 画像1つ → 画像外接の矩形でクリップ
    function createClippingMaskGroup(imageItem) {
        var targetLayer = imageItem.layer;
        var wasLocked = targetLayer.locked;
        var wasVisible = targetLayer.visible;
        var wasTemplate = targetLayer.isTemplate;

        if (wasLocked) targetLayer.locked = false;
        if (!wasVisible) targetLayer.visible = true;
        if (wasTemplate) targetLayer.isTemplate = false;

        var rect = targetLayer.pathItems.rectangle(
            imageItem.top,
            imageItem.left,
            imageItem.width,
            imageItem.height
        );
        rect.stroked = false;
        rect.filled = false;

        var groupItem = targetLayer.groupItems.add();
        imageItem.moveToBeginning(groupItem);
        rect.moveToBeginning(groupItem);
        groupItem.clipped = true;

        if (wasLocked) targetLayer.locked = true;
        if (!wasVisible) targetLayer.visible = false;
        if (wasTemplate) targetLayer.isTemplate = true;

        return groupItem;
    }

    // 画像1つ + パス1つ → パスでクリップ
    function createMaskWithPath(imageItem, pathItem) {
        var targetLayer = imageItem.layer;
        if (pathItem.layer != targetLayer) {
            pathItem.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
        }

        var groupItem = targetLayer.groupItems.add();
        imageItem.moveToBeginning(groupItem);
        pathItem.moveToBeginning(groupItem);
        groupItem.clipped = true;

        return groupItem;
    }

    // 現在の選択から「クリップグループ」を作成し、選択を更新
    // 戻り値: 作成した groupItem（複数作成時は先頭を返す） / 作成できない場合は null
    function makeClippingFromSelection(doc) {
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) return null;

        // すでにクリップグループが1つ選択されているなら、そのまま
        if (currentSelection.length === 1 && currentSelection[0].typename === 'GroupItem' && currentSelection[0].clipped) {
            return currentSelection[0];
        }

        var images = [];
        var paths = [];
        var i;

        for (i = 0; i < currentSelection.length; i++) {
            if (!currentSelection[i]) continue;
            if (isImageItem(currentSelection[i])) {
                images.push(currentSelection[i]);
            } else if (currentSelection[i].typename === 'PathItem') {
                paths.push(currentSelection[i]);
            }
        }

        // 画像1 + パス1（選択がちょうど2つ）
        if (currentSelection.length === 2 && images.length === 1 && paths.length === 1) {
            var g1 = createMaskWithPath(images[0], paths[0]);
            doc.selection = [g1];
            return g1;
        }

        // 画像1のみ
        if (currentSelection.length === 1 && images.length === 1) {
            var g2 = createClippingMaskGroup(images[0]);
            doc.selection = [g2];
            return g2;
        }

        // パスが含まれている場合は、最前面パスをマスクとして全体をクリップ
        if (paths.length > 0) {
            var maskPath = getFrontmostPath(paths);
            if (!maskPath) return null;

            var targetLayer = maskPath.layer;
            var wasLocked = targetLayer.locked;
            var wasVisible = targetLayer.visible;
            var wasTemplate = targetLayer.isTemplate;

            if (wasLocked) targetLayer.locked = false;
            if (!wasVisible) targetLayer.visible = true;
            if (wasTemplate) targetLayer.isTemplate = false;

            var groupItem = targetLayer.groupItems.add();

            // マスク以外を先に移動（末尾から）
            for (i = currentSelection.length - 1; i >= 0; i--) {
                if (!currentSelection[i]) continue;
                if (currentSelection[i] === maskPath) continue;
                currentSelection[i].moveToBeginning(groupItem);
            }
            maskPath.moveToBeginning(groupItem);
            groupItem.clipped = true;

            if (wasLocked) targetLayer.locked = true;
            if (!wasVisible) targetLayer.visible = false;
            if (wasTemplate) targetLayer.isTemplate = true;

            doc.selection = [groupItem];
            return groupItem;
        }

        // 複数画像のみ → それぞれ外接矩形で個別にクリップ
        if (images.length > 0) {
            var groups = [];
            for (i = 0; i < images.length; i++) {
                try {
                    groups.push(createClippingMaskGroup(images[i]));
                } catch (e) { }
            }
            if (groups.length > 0) {
                doc.selection = groups;
                return groups[0];
            }
        }

        return null;
    }

    // 角丸 LiveEffect XML を生成
    function createRoundCornersEffectXML(radius) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radius);
    }

    // 角丸（LiveEffect: Adobe Round Corners）を適用
    function applyRoundCornersLiveEffect(targetItem, radius) {
        if (!targetItem) return false;
        var r = Number(radius);
        if (isNaN(r) || r < 0) return false;

        var xml = createRoundCornersEffectXML(r);

        try {
            targetItem.applyEffect(xml);
            return true;
        } catch (e) {
            return false;
        }
    }

    // ===== Embed =====
    function embedPlacedItems(placedItems) {
        var count = 0;

        for (var i = 0; i < placedItems.length; i++) {
            var item = placedItems[i];
            if (!item) continue;

            // 親がクリップグループかどうか
            var parent = item.parent;
            var isClipGroup = false;
            if (parent && parent.typename === 'GroupItem') {
                isClipGroup = parent.clipped;
            }

            // PSDファイルかどうか
            if (item.file && item.file.name.match(/\.psd$/i)) {
                try {
                    item.embed();
                    count++;
                } catch (e) {
                    continue;
                }

                if (!isClipGroup) {
                    try {
                        app.executeMenuCommand('ungroup');
                    } catch (e2) { }
                }
            } else {
                try {
                    item.embed();
                    count++;
                } catch (e3) {
                    continue;
                }
            }
        }

        return count;
    }

    function sanitizeFileBaseName(name) {
        if (!name) return '';
        // strip any path
        name = String(name).replace(/^.*[\\\/]/, '');
        // strip extension
        name = name.replace(/\.[^.]+$/, '');
        // replace invalid characters for mac/windows
        name = name.replace(/[\/:*?"<>|\r\n]+/g, '_');
        name = name.replace(/^\s+|\s+$/g, '');
        if (name.length > 120) name = name.substring(0, 120);
        return name;
    }

    function getFileNameFromXMPForEmbeddedItem(item) {
        try {
            var xmpStr = item.XMPString;
            if (!xmpStr) return '';

            // dc:title
            var titleMatch = xmpStr.match(/<dc:title>\s*<rdf:Alt>\s*<rdf:li[^>]*>(.*?)<\/rdf:li>/);
            if (titleMatch && titleMatch[1]) return titleMatch[1];

            // tiff:ImageDescription
            var descMatch = xmpStr.match(/tiff:ImageDescription>(.*?)<\/tiff:ImageDescription>/);
            if (descMatch && descMatch[1]) return descMatch[1];

            // photoshop:Source (sometimes contains a path)
            var srcMatch = xmpStr.match(/photoshop:Source="(.*?)"/);
            if (srcMatch && srcMatch[1]) return srcMatch[1];

            return '';
        } catch (e) {
            return '';
        }
    }

    function getPreferredBaseNameForUnembed(item, fallbackBaseName) {
        var name = '';

        // 1) item.name (レイヤーパネルの名称にファイル名が残っていることが多い)
        try {
            if (item.name && item.name !== '') name = item.name;
        } catch (e1) { }

        // 2) item.file.name（埋め込み後は null が多いが念のため）
        if (!name) {
            try {
                if (item.file && item.file.name) name = item.file.name;
            } catch (e2) { }
        }

        // 3) XMP
        if (!name) {
            name = getFileNameFromXMPForEmbeddedItem(item);
        }

        name = sanitizeFileBaseName(name);
        if (!name) name = fallbackBaseName;

        return name;
    }

    // ===== Release (Unembed) =====
    function releaseRasterItemsToLinks(doc, rasterItems) {
        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var exportErrors = [];
        var counter = 0;
        var linkedCount = 0;

        // export folder
        var exportFolder;
        try {
            exportFolder = Folder(doc.fullName.parent.fsName + '/Links/');
        } catch (e) {
            exportFolder = Folder(Folder.desktop.fsName + '/Links/');
        }
        if (!exportFolder.exists) {
            try { exportFolder.create(); } catch (e2) { }
        }

        try {
            // 後ろから処理（順序変化に強く）
            for (var i = rasterItems.length - 1; i >= 0; i--) {
                var oldImage = rasterItems[i];
                if (!oldImage) continue;

                var fallbackTitle = 'image' + (i + 1);
                var imageTitle = getPreferredBaseNameForUnembed(oldImage, fallbackTitle);

                // colorSpace
                var colorSpace;
                try {
                    // RasterItem
                    colorSpace = oldImage.imageColorSpace;
                } catch (eCS) {
                    // embedded PlacedItem など
                    try {
                        colorSpace = (doc.documentColorSpace === DocumentColorSpace.CMYK) ? ImageColorSpace.CMYK : ImageColorSpace.RGB;
                    } catch (eCS2) {
                        colorSpace = ImageColorSpace.RGB;
                    }
                }

                // script can't handle other colorSpaces
                if (
                    colorSpace != ImageColorSpace.CMYK &&
                    colorSpace != ImageColorSpace.RGB &&
                    colorSpace != ImageColorSpace.GrayScale
                ) {
                    exportErrors.push(imageTitle + ' has unsupported color space. (' + colorSpace + ')');
                    continue;
                }

                var position;
                try { position = oldImage.position; } catch (ePos) { position = [0, 0]; }

                // get current scale and rotation
                var sr;
                try {
                    sr = getLinkScaleAndRotation(oldImage);
                } catch (eSR) {
                    sr = [100, 100, 0];
                }
                var scale = [sr[0], sr[1]];
                var rotation = sr[2];

                // move to new document for exporting
                var temp;
                try {
                    temp = newDocument(imageTitle, colorSpace, 1000, 1000);
                } catch (eDoc) {
                    exportErrors.push('"' + imageTitle + '" failed to create temp doc. (' + eDoc.message + ')');
                    continue;
                }

                // duplicate to new document
                var workingImage;
                try {
                    workingImage = oldImage.duplicate(temp.layers[0], ElementPlacement.PLACEATBEGINNING);
                } catch (eDup) {
                    exportErrors.push('"' + imageTitle + '" failed to duplicate. (' + eDup.message + ')');
                    try { temp.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose1) { }
                    continue;
                }

                // set image to 100% scale 0° rotation and position
                try {
                    var tm = app.getRotationMatrix(-rotation);
                    tm = app.concatenateScaleMatrix(tm, 100 / scale[0] * 100, 100 / scale[1] * 100);
                    workingImage.transform(tm, true, true, true, true, true);
                    workingImage.position = [0, workingImage.height];
                    temp.artboards[0].artboardRect = [0, workingImage.height, workingImage.width, 0];
                } catch (eTx) { }

                // export
                var file;
                try {
                    var path = getFilePathWithOverwriteProtectionSuffix(exportFolder.fsName + '/' + imageTitle + '.psd');
                    file = exportAsPSD(temp, path, colorSpace, 72);
                } catch (error) {
                    exportErrors.push('"' + imageTitle + '" failed to export. (' + error.message + ')');
                }

                // close temp doc
                try { temp.close(SaveOptions.DONOTSAVECHANGES); } catch (eClose2) { }

                if (!file || !file.exists) {
                    continue;
                }

                // place link in active layer then move to oldImage position in stack
                var newImage;
                try {
                    var targetLayer = doc.activeLayer || oldImage.layer;
                    newImage = targetLayer.placedItems.add();

                    // まず file を設定（基本）
                    newImage.file = file;

                    // 環境によっては file 代入だけでは反映されないことがあるため、relink/update を試行
                    try {
                        if (newImage.relink) newImage.relink(file);
                    } catch (eRelink) { }
                    try {
                        if (newImage.update) newImage.update();
                    } catch (eUpdate) { }

                    // リンクできているか簡易検証
                    try {
                        if (newImage.file != null && newImage.file.exists) {
                            linkedCount++;
                        }
                    } catch (eCheck) { }

                } catch (eLink) {
                    exportErrors.push('"' + imageTitle + '" failed to place link. (' + eLink.message + ')');
                    try { if (file && file.exists) file.remove(); } catch (eRm) { }
                    continue;
                }

                // scale, rotate, position to match original
                try {
                    var tm2 = app.getScaleMatrix(scale[0], scale[1]);
                    tm2 = app.concatenateRotationMatrix(tm2, rotation);
                    newImage.transform(tm2, true, true, true, true, true);
                } catch (eFit1) { }

                // stacking + position
                try {
                    newImage.move(oldImage, ElementPlacement.PLACEAFTER);
                } catch (eMove) { }
                try {
                    newImage.position = position;
                } catch (eFit2) { }

                // remove old embedded image
                try {
                    oldImage.remove();
                } catch (eDel) {
                    exportErrors.push('Warning: Could not remove embedded image for "' + imageTitle + '".');
                }

                counter++;
            }
        } finally {
            app.userInteractionLevel = previousInteractionLevel;
            app.redraw();
        }

        return {
            count: counter,
            linkedCount: linkedCount,
            errors: exportErrors
        };
    }

    // ===== Helpers (from reference, simplified) =====
    function exportAsPSD(doc, path, colorSpace, resolution) {
        var file = File(path);
        var options = new ExportOptionsPhotoshop();

        options.antiAliasing = false;
        options.artBoardClipping = true;
        options.imageColorSpace = colorSpace;
        options.editableText = false;
        options.flatten = true;
        options.maximumEditability = false;
        options.resolution = (resolution || 72);
        options.warnings = false;
        options.writeLayers = false;

        doc.exportFile(file, ExportType.PHOTOSHOP, options);
        return file;
    }

    function getLinkScaleAndRotation(item) {
        if (item == undefined) return;

        var m = item.matrix;
        var rotatedAmount;
        var unrotatedMatrix;
        var scaledAmount;

        var flipPlacedItem = (item.typename == 'PlacedItem') ? 1 : -1;

        try {
            rotatedAmount = item.tags.getByName('BBAccumRotation').value * 180 / Math.PI;
        } catch (error) {
            rotatedAmount = 0;
        }

        unrotatedMatrix = app.concatenateRotationMatrix(m, rotatedAmount * flipPlacedItem);
        scaledAmount = [unrotatedMatrix.mValueA * 100, unrotatedMatrix.mValueD * -100 * flipPlacedItem];

        return [scaledAmount[0], scaledAmount[1], rotatedAmount];
    }

    function newDocument(name, colorSpace, width, height) {
        var myDocPreset = new DocumentPreset();
        var myDocPresetType;

        myDocPreset.title = name;
        myDocPreset.width = width || 1000;
        myDocPreset.height = height || 1000;

        if (
            colorSpace == ImageColorSpace.CMYK ||
            colorSpace == ImageColorSpace.GrayScale ||
            colorSpace == ImageColorSpace.DeviceN
        ) {
            myDocPresetType = DocumentPresetType.BasicCMYK;
            myDocPreset.colorMode = DocumentColorSpace.CMYK;
        } else {
            myDocPresetType = DocumentPresetType.BasicRGB;
            myDocPreset.colorMode = DocumentColorSpace.RGB;
        }

        return app.documents.addDocument(myDocPresetType, myDocPreset);
    }

    function getFilePathWithOverwriteProtectionSuffix(path) {
        var index = 1;
        var parts = path.split(/(\.[^\.]+)$/);

        while (File(path).exists) {
            path = parts[0] + '(' + (++index) + ')' + parts[1];
        }

        return path;
    }

    // ===== Main =====
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;

        var dialogResult = showDialog();
        if (!dialogResult) return;

        // リセット
        if (dialogResult.mode === 'reset') {
            if (!doc.selection || doc.selection.length === 0) {
                doc.selection = null;
                return;
            }
            var targets = collectPlacedOrRasterFromSelection(doc.selection);
            if (targets.length === 0) {
                doc.selection = null;
                return;
            }
            applyResetTransformToItems(dialogResult.resetOptions || {}, targets);
            doc.selection = null;
            return;
        }

        // リンク
        if (dialogResult.mode === 'link') {
            var linkMode = (dialogResult.linkOptions && dialogResult.linkOptions.mode) ? dialogResult.linkOptions.mode : 'update';

            // リンクを更新：ドキュメント内のすべてのリンク配置画像を対象
            if (linkMode === 'update') {
                var targetsAll = [];
                try {
                    var itemsAll = doc.placedItems;
                    for (var ai = 0; ai < itemsAll.length; ai++) {
                        var itAll = itemsAll[ai];
                        if (!itAll || itAll.typename !== 'PlacedItem') continue;
                        try {
                            if (itAll.file && itAll.file.exists) {
                                targetsAll.push(itAll);
                            }
                        } catch (eFile) { }
                    }
                } catch (eCollectAll) { }

                if (targetsAll.length > 0) {
                    try {
                        doc.selection = targetsAll;
                        app.executeMenuCommand('Adobe Update Link Shortcut');
                    } catch (eUpdAll) { }
                }

                doc.selection = null;
                return;
            }

            // すべてさしかえ
            // 選択した配置画像と「同じファイル名」のリンク配置画像をすべて、新しいファイルに差し替える
            if (linkMode === 'relinkAll') {
                var currentSelection = doc.selection;
                if (!currentSelection || currentSelection.length === 0) {
                    doc.selection = null;
                    alert(getLabel('alertSelectPlacedItem'));
                    return;
                }

                var selectedItem = currentSelection[0];
                if (!selectedItem || selectedItem.typename !== 'PlacedItem') {
                    doc.selection = null;
                    alert(getLabel('alertSelectedNotPlacedItem'));
                    return;
                }

                // 選択された配置画像のファイル名
                var fileName = '';
                try {
                    if (selectedItem.file && selectedItem.file.name) {
                        fileName = selectedItem.file.name;
                    }
                } catch (eFN) { }

                if (!fileName) {
                    doc.selection = null;
                    alert(getLabel('alertSelectPlacedItem'));
                    return;
                }

                // 同じ名前のリンク配置画像をすべて収集
                var targets = [];
                try {
                    var linkedItems = doc.placedItems;
                    for (var i = 0; i < linkedItems.length; i++) {
                        var it = linkedItems[i];
                        if (!it || it.typename !== 'PlacedItem') continue;

                        try {
                            if (it.file && it.file.name === fileName) {
                                targets.push(it);
                            }
                        } catch (eIt) { }
                    }
                } catch (eCollect) { }

                if (targets.length === 0) {
                    doc.selection = null;
                    return;
                }

                // 置換先ファイルを選択
                var fileToReplace = File.openDialog(getLabel('dialogSelectReplaceFile'));
                if (!fileToReplace) {
                    doc.selection = null;
                    alert(getLabel('alertFileSelectionCanceled'));
                    return;
                }

                // 差し替え（relink/update を優先）
                for (var j = 0; j < targets.length; j++) {
                    try {
                        if (targets[j].relink) {
                            targets[j].relink(fileToReplace);
                        } else {
                            targets[j].file = fileToReplace;
                        }
                        try { if (targets[j].update) targets[j].update(); } catch (eUpd2) { }
                    } catch (eRel) { }
                }

                // 画面更新
                app.redraw();

                // 既存の選択をクリア
                doc.selection = null;
                return;
            }

            doc.selection = null;
            return;
        }

        // ケイ線
        if (dialogResult.mode === 'keisen') {
            if (!doc.selection || doc.selection.length === 0) {
                alert(getLabel('alertSelectObject'));
                doc.selection = null;
                return;
            }

            try {
                applyKeisenToSelection(doc, dialogResult.keisenOptions || {});
            } catch (eKeisen) {
                alert(getLabel('alertKeisenError') + eKeisen);
            }

            doc.selection = null;
            return;
        }

        // モード分岐
        if (dialogResult.mode === 'embed') {
            var placedItems = [];

            if (dialogResult.embedMode === 'selection') {
                if (!doc.selection || doc.selection.length === 0) {
                    doc.selection = null;
                    return;
                }
                placedItems = collectPlacedItemsFromSelection(doc.selection);
            } else {
                placedItems = collectAllPlacedItems(doc);
            }

            if (placedItems.length === 0) {
                doc.selection = null;
                return;
            }

            var embeddedCount = embedPlacedItems(placedItems);
            if (embeddedCount > 0) {
                alert(embeddedCount + getLabel('alertEmbeddedDoneSuffix'));
            }

            doc.selection = null;
            return;
        }

        // ドキュメント未保存の場合は埋め込み解除を禁止
        try {
            if (!doc.saved) {
                alert(getLabel('alertDocNotSaved'));
                doc.selection = null;
                return;
            }
        } catch (eSave) { }

        // 解除（埋め込み解除）
        var rasterItems = [];

        if (dialogResult.releaseMode === 'selection') {
            if (!doc.selection || doc.selection.length === 0) {
                doc.selection = null;
                return;
            }
            // 選択に「リンク」画像（PlacedItem）が含まれている場合は解除できないため通知
            if (selectionHasLinkedPlacedItem(doc.selection)) {
                alert(getLabel('alertLinkedCannotUnembed'));
                doc.selection = null;
                return;
            }
            rasterItems = collectRasterItemsFromSelection(doc.selection);
        } else {
            rasterItems = collectAllRasterItems(doc);
        }

        if (rasterItems.length === 0) {
            doc.selection = null;
            return;
        }

        var result = releaseRasterItemsToLinks(doc, rasterItems);
        if (result) {
            if (result.count > 0) {
                var msg = result.count + getLabel('alertReleasedDoneSuffix');
                if (typeof result.linkedCount === 'number') {
                    msg += '\n' + getLabel('alertLinkedCreatedPrefix') + result.linkedCount + getLabel('alertLinkedCreatedSuffix');
                    if (result.linkedCount === 0) {
                        msg += getLabel('alertLinkedCreateZeroNote');
                    }
                }
                if (result.errors && result.errors.length > 0) {
                    msg += '\n\n' + result.errors.join('\n');
                }
                alert(msg);
            } else if (result.errors && result.errors.length > 0) {
                alert(getLabel('alertUnembedFailed') + result.errors.join('\n'));
            }
        }

        doc.selection = null;
    }

    main();

})();
