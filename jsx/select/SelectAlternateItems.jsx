#target illustrator
#targetengine "SelectAlternateItemsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトを並び順で数え、奇数番目または偶数番目だけを互い違いに選択し直します。数える方向は垂直・水平から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectAlternateItems.md

### Overview

Counts the selected objects in order and reselects only the odd- or even-numbered ones. The counting direction can be set to vertical or horizontal.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectAlternateItems.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectAlternateItems";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectAlternateItems.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectAlternateItems.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

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
      dialog: {
        title: { ja: "互い違いに選択", en: "Alternate Select" }
      },
      panel: {
        select: { ja: "選択", en: "Selection" },
        direction: { ja: "方向", en: "Direction" }
      },
      radio: {
        odd: { ja: "奇数", en: "Odd" },
        even: { ja: "偶数", en: "Even" },
        vertical: { ja: "垂直", en: "Vertical" },
        horizontal: { ja: "水平", en: "Horizontal" },
        zOrder: { ja: "重ね順", en: "Z-order" }
      },
      tooltip: {
        odd: { ja: "並びの1番目から1つおきに選びます。", en: "Selects every other object starting from the first." },
        even: { ja: "並びの2番目から1つおきに選びます。", en: "Selects every other object starting from the second." },
        vertical: { ja: "上から下の並び順で数えます。", en: "Counts the objects from top to bottom." },
        horizontal: { ja: "左から右の並び順で数えます。", en: "Counts the objects from left to right." },
        zOrder: { ja: "重ね順（背面から前面）で数えます。位置ではなく前後関係で選びます。", en: "Counts the objects by stacking order, from back to front, rather than by position." }
      },
      button: {
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" }
      },
      /* エラー／ログ文言 / Error and log messages */
      alert: {
        noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
        noValidItems: { ja: "有効なオブジェクトが選択されていません。", en: "No valid objects are selected." },
        preview: { ja: "プレビューエラー：", en: "Preview Error: " },
        prefix: { ja: "エラー：", en: "Error: " },
        debugPrefix: { ja: "デバッグ：", en: "Debug: " }
      }
    };


    function getErrorMessage(key, detail) {
      var message = getLabel("alert." + key);
      if (detail !== undefined && detail !== null && String(detail) !== "") {
        message += String(detail);
      }
      return message;
    }

    function showError(key, detail) {
      alert(getErrorMessage(key, detail));
    }

    /* プレビュー管理 / Preview manager
       - 選択のプレビューは通常 Undo 履歴に乗らないため、選択はスナップショット復元
       - ドキュメント変更を伴う場合のみ app.undo() を使って巻き戻せるようにする
    */
    function PreviewManager() {
      this.undoDepth = 0; // Undoable step count during preview
      this.selectionSnapshot = null; // Original selection snapshot

      /* 選択状態を退避 / Capture selection */
      this.captureSelection = function (doc) {
        var snap = [];
        try {
          var currentSelection = doc.selection;
          if (currentSelection && currentSelection.length) {
            for (var selectionIndex = 0; selectionIndex < currentSelection.length; selectionIndex++) {
              snap.push(currentSelection[selectionIndex]);
            }
          }
        } catch (e) { }
        this.selectionSnapshot = snap;
      };

      /* 選択状態を復元 / Restore selection */
      this.restoreSelection = function (doc) {
        try {
          doc.selection = null;
          if (this.selectionSnapshot && this.selectionSnapshot.length) {
            for (var snapshotIndex = 0; snapshotIndex < this.selectionSnapshot.length; snapshotIndex++) {
              try { this.selectionSnapshot[snapshotIndex].selected = true; } catch (e) { }
            }
          }
          app.redraw();
        } catch (e) { }
      };

      /**
       * 変更操作を実行し、必要なら履歴としてカウント / Run step and optionally count as undoable
       * @param {Function} func - 実行処理 / action
       * @param {Boolean} [undoable=false] - Undo 対象なら true / true if it creates undo history
       */
      this.addStep = function (func, undoable) {
        try {
          func();
          if (undoable) this.undoDepth++;
          app.redraw();
        } catch (e) {
          showError("preview", e);
        }
      };

      /* プレビュー分の変更を巻き戻し / Rollback preview changes */
      this.rollback = function (doc) {
        // Undoable steps rollback
        while (this.undoDepth > 0) {
          try { app.undo(); } catch (e) { break; }
          this.undoDepth--;
        }
        // Selection rollback (always)
        this.restoreSelection(doc);
      };

      /**
       * 確定 / Confirm
       * @param {Document} doc
       * @param {Function} [finalAction] - 一度戻してから本番処理 / optional final action
       */
      this.confirm = function (doc, finalAction) {
        if (finalAction) {
          this.rollback(doc);
          finalAction();
          this.undoDepth = 0;
          this.captureSelection(doc); // keep current as baseline
        } else {
          this.undoDepth = 0;
          this.captureSelection(doc); // keep current as baseline
        }
      };
    }

    /* 方向推定の設定 / Direction detection settings */
    var DIRECTION_THRESHOLD = 1.20; // しきい値（1.20 = 20%差） / Threshold ratio
    var PREF_KEY_LAST_DIR = "AlternateSelect.LastDirectionMode";

    /* 前回の方向モードを取得 / Load last direction mode (custom options)
       - 取得できない場合は null を返す
       - mode: "vertical" | "horizontal" | "zorder"
    */
    function loadLastDirectionMode() {
      try {
        var desc = app.getCustomOptions(PREF_KEY_LAST_DIR);
        if (desc && desc.hasKey(stringIDToTypeID("mode"))) {
          return desc.getString(stringIDToTypeID("mode"));
        }
      } catch (e) { }
      return null;
    }

    /* 前回の方向モードを保存 / Save last direction mode (custom options) */
    function saveLastDirectionMode(mode) {
      try {
        var m = String(mode);
        if (m !== "vertical" && m !== "horizontal" && m !== "zorder") return;
        var desc = new ActionDescriptor();
        desc.putString(stringIDToTypeID("mode"), m);
        app.putCustomOptions(PREF_KEY_LAST_DIR, desc, true);
      } catch (e) { }
    }

    /* 方向の自動推定 / Auto detect direction (initial)
       - 中心点のX/Yレンジで判定（外れ値耐性あり）
       - rangeX が rangeY の DIRECTION_THRESHOLD 倍より大きい → "horizontal"
       - rangeY が rangeX の DIRECTION_THRESHOLD 倍より大きい → "vertical"
       - それ以外（僅差/グリッド等）は「前回mode優先（fallback）」を返す
    */
    function guessInitialDirectionMode(items, fallbackMode) {
      var fallbackDirectionMode = (fallbackMode !== undefined && fallbackMode !== null) ? String(fallbackMode) : "vertical";
      if (fallbackDirectionMode !== "vertical" && fallbackDirectionMode !== "horizontal" && fallbackDirectionMode !== "zorder") fallbackDirectionMode = "vertical";

      if (!items || items.length < 2) return fallbackDirectionMode;

      var xs = [];
      var ys = [];

      for (var i = 0; i < items.length; i++) {
        var candidateItem = items[i];
        if (!candidateItem) continue;

        // geometricBounds: [left, top, right, bottom]
        var geometricBounds;
        try { geometricBounds = candidateItem.geometricBounds; } catch (e) { continue; }
        if (!geometricBounds || geometricBounds.length < 4) continue;

        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;
        xs.push(centerX);
        ys.push(centerY);
      }

      if (xs.length < 2 || ys.length < 2) return fallbackDirectionMode;

      xs.sort(function (firstItem, secondItem) { return firstItem - secondItem; });
      ys.sort(function (firstItem, secondItem) { return firstItem - secondItem; });

      // 外れ値耐性：両端を少し落としてレンジを取る（nが大きいほど効果）
      var n = xs.length;
      var trim = 0;
      if (n >= 10) trim = Math.floor(n * 0.10); // 10% trimming
      else if (n >= 6) trim = 1;

      var minX = xs[trim];
      var maxX = xs[n - 1 - trim];
      var minY = ys[trim];
      var maxY = ys[n - 1 - trim];

      var rangeX = maxX - minX;
      var rangeY = maxY - minY;

      // しきい値判定（明確な差があるときだけ自動切替）
      if (rangeX > rangeY * DIRECTION_THRESHOLD) return "horizontal";
      if (rangeY > rangeX * DIRECTION_THRESHOLD) return "vertical";

      // 僅差（グリッド等）は前回mode優先
      return fallbackDirectionMode;
    }

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

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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

    (function () {
      /* ドキュメント確認 / Check document */
      if (app.documents.length === 0) {
        showError("noDocument");
        return;
      }

      var doc = app.activeDocument;

      /* プレビューマネージャ / Preview manager */
      var previewMgr = new PreviewManager();
      previewMgr.captureSelection(doc);

      var originalSelection = doc.selection;

      /* 選択確認 / Check selection */
      if (!originalSelection || originalSelection.length === 0) {
        showError("noSelection");
        return;
      }

      /* 有効なオブジェクト抽出 / Collect valid items */
      var items = [];
      for (var selectionIndex = 0; selectionIndex < originalSelection.length; selectionIndex++) {
        var selectedItem = originalSelection[selectionIndex];
        if (!selectedItem.locked && !selectedItem.hidden) {
          items.push(selectedItem);
        }
      }

      if (items.length === 0) {
        showError("noValidItems");
        return;
      }

      function applySelectionPreview() {
        // まず前回プレビューを巻き戻す（選択はスナップショットで復元）
        previewMgr.rollback(doc);

        // 今回のプレビューを適用（選択変更は通常 Undoable ではないので undoable=false）
        previewMgr.addStep(function () {
          var selectOdd = oddRadio.value;
          var mode = verticalRadio.value ? "vertical" : (horizontalRadio.value ? "horizontal" : "zorder");

          /* 並び順ソート / Sort order */
          if (mode === "vertical") {
            /* 垂直：Y降順（上→下） / Vertical: Y desc (top to bottom) */
            items.sort(function (firstItem, secondItem) { return secondItem.position[1] - firstItem.position[1]; });
          } else if (mode === "horizontal") {
            /* 水平：X昇順（左→右） / Horizontal: X asc (left to right) */
            items.sort(function (firstItem, secondItem) { return firstItem.position[0] - secondItem.position[0]; });
          } else {
            /* 重ね順：zOrderPosition 昇順 / Z-order: zOrderPosition asc */
            items.sort(function (firstItem, secondItem) {
              var firstZOrderPosition = 0, secondZOrderPosition = 0;
              try { firstZOrderPosition = firstItem.zOrderPosition; } catch (e) { firstZOrderPosition = 0; }
              try { secondZOrderPosition = secondItem.zOrderPosition; } catch (e) { secondZOrderPosition = 0; }
              return firstZOrderPosition - secondZOrderPosition;
            });
          }

          /* 互い違い選択 / Alternate selection */
          doc.selection = null;
          for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var isOddIndex = (itemIndex % 2 === 0); // 0,2,4... => 1,3,5...
            if ((selectOdd && isOddIndex) || (!selectOdd && !isOddIndex)) {
              items[itemIndex].selected = true;
            }
          }
        }, false);
      }

      /* ダイアログボックス / Dialog box */
      var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
      dialog.orientation = "column";
      dialog.alignChildren = "left";

      var selectionPanel = dialog.add("panel", undefined, getLabel("panel.select"));
      selectionPanel.margins = [15, 20, 15, 10];
      selectionPanel.orientation = "column";
      selectionPanel.alignChildren = "left";

      var selectGroup = selectionPanel.add("group");
      selectGroup.orientation = "row";
      selectGroup.alignChildren = "left";

      var oddRadio = selectGroup.add("radiobutton", undefined, getLabel("radio.odd"));
      oddRadio.helpTip = getLabel("tooltip.odd");
      var evenRadio = selectGroup.add("radiobutton", undefined, getLabel("radio.even"));
      evenRadio.helpTip = getLabel("tooltip.even");
      oddRadio.value = true; // デフォルトは奇数 / Default is odd

      function setAlternateSelectionMode(selectionMode) {
        oddRadio.value = (selectionMode === "odd");
        evenRadio.value = (selectionMode === "even");
        applySelectionPreview();
      }

      /* 方向パネル / Direction panel */
      var dirPanel = dialog.add("panel", undefined, getLabel("panel.direction"));
      dirPanel.margins = [15, 20, 15, 10];
      dirPanel.orientation = "column";
      dirPanel.alignChildren = "left";

      var dirGroup = dirPanel.add("group");
      dirGroup.orientation = "row";
      dirGroup.alignChildren = "left";

      var verticalRadio = dirGroup.add("radiobutton", undefined, getLabel("radio.vertical"));
      verticalRadio.helpTip = getLabel("tooltip.vertical");
      var horizontalRadio = dirGroup.add("radiobutton", undefined, getLabel("radio.horizontal"));
      horizontalRadio.helpTip = getLabel("tooltip.horizontal");
      var zOrderRadio = dirGroup.add("radiobutton", undefined, getLabel("radio.zOrder"));
      zOrderRadio.helpTip = getLabel("tooltip.zOrder");

      function setDirectionMode(directionMode) {
        verticalRadio.value = (directionMode === "vertical");
        horizontalRadio.value = (directionMode === "horizontal");
        zOrderRadio.value = (directionMode === "zorder");
        saveLastDirectionMode(directionMode);
        applySelectionPreview();
      }
      // 前回値を基本にし、差が明確なときだけ自動切替 / Prefer last value; auto-switch only if clear
      var lastMode = loadLastDirectionMode();
      if (lastMode === null) lastMode = "vertical";

      var initialMode = guessInitialDirectionMode(items, lastMode);
      verticalRadio.value = (initialMode === "vertical");
      horizontalRadio.value = (initialMode === "horizontal");
      zOrderRadio.value = (initialMode === "zorder");

      var buttonGroup = dialog.add("group");
      buttonGroup.alignment = "right";
      var cancelBtn = buttonGroup.add("button", undefined, getLabel("button.cancel"));
      var okBtn = buttonGroup.add("button", undefined, getLabel("button.ok"));

      oddRadio.onClick = function () {
        setAlternateSelectionMode("odd");
      };
      evenRadio.onClick = function () {
        setAlternateSelectionMode("even");
      };
      verticalRadio.onClick = function () {
        setDirectionMode("vertical");
      };
      horizontalRadio.onClick = function () {
        setDirectionMode("horizontal");
      };
      zOrderRadio.onClick = function () {
        setDirectionMode("zorder");
      };

      /* キー入力でラジオ切替 / Keyboard shortcuts for the radio buttons
         Odd: O / Even: E / Vertical: V / Horizontal: H / Z-Order: A */
      addKeyShortcuts(dialog, {
        "O": oddRadio,
        "E": evenRadio,
        "V": verticalRadio,
        "H": horizontalRadio,
        "A": zOrderRadio
      });

      /* ボタン動作 / Button handlers */
      cancelBtn.onClick = function () {
        // キャンセル：プレビューを巻き戻して閉じる / Cancel: rollback and close
        previewMgr.rollback(doc);
        dialog.close(0);
      };

      okBtn.onClick = function () {
        // OK：Undo を綺麗にするため「一度戻して再実行」 / OK: rollback then re-apply once
        previewMgr.confirm(doc, function () {
          // 最終適用（1回） / Final apply (single step)
          var selectOdd = oddRadio.value;
          var mode = verticalRadio.value ? "vertical" : (horizontalRadio.value ? "horizontal" : "zorder");

          if (mode === "vertical") {
            items.sort(function (firstItem, secondItem) { return secondItem.position[1] - firstItem.position[1]; });
          } else if (mode === "horizontal") {
            items.sort(function (firstItem, secondItem) { return firstItem.position[0] - secondItem.position[0]; });
          } else {
            items.sort(function (firstItem, secondItem) {
              var firstZOrderPosition = 0, secondZOrderPosition = 0;
              try { firstZOrderPosition = firstItem.zOrderPosition; } catch (e) { firstZOrderPosition = 0; }
              try { secondZOrderPosition = secondItem.zOrderPosition; } catch (e) { secondZOrderPosition = 0; }
              return firstZOrderPosition - secondZOrderPosition;
            });
          }

          doc.selection = null;
          for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
            var isOddIndex = (itemIndex % 2 === 0);
            if ((selectOdd && isOddIndex) || (!selectOdd && !isOddIndex)) {
              items[itemIndex].selected = true;
            }
          }
          app.redraw();
        });
        var modeToSave = verticalRadio.value ? "vertical" : (horizontalRadio.value ? "horizontal" : "zorder");
        saveLastDirectionMode(modeToSave);
        dialog.close(1);
      };

      // 初期状態をプレビュー反映（ダイアログ表示直後に選択を更新） / Initial preview
      applySelectionPreview();

      // ダイアログ表示 / Show dialog
      prepareDialogWindow(dialog, SCRIPT_NAME);
      dialog.show();
    })();

})();
