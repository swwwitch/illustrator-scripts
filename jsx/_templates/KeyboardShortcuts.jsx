#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログやパレットに、文字キーでラジオボタンを選ぶ・チェックボックスを切り替える・ボタンを押すショートカットを付ける再利用テンプレートです。
入力欄の編集中、修飾キーを押しているとき、無効なコントロール（親ごと無効なものを含む）では何もしません。

### Overview

A reusable template that adds letter-key shortcuts to a dialog or palette: select a radio button, toggle a checkbox or press a button.
Nothing happens while a text field is being edited, while modifier keys are held, or when the control (or any of its parents) is disabled.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "KeyboardShortcuts";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    // 【移植手順 / How to port】
    // 1. 「（再利用パーツ）」の行から「ここまで」の行までをまるごと、コピー先の IIFE 内（uiLang の定義より後、ダイアログを作る関数より前）に貼る。
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
    // 9. パレット（new Window("palette")）は、ダイアログと違って Esc では閉じない。パレットには必ず Esc を割り当てる。
    //    ［閉じる］ボタンがあればそれと同じ処理（プレビューの片付けなど）を呼び、入力中も効かせる
    //      addKeyShortcuts(palette, {
    //          "Escape": { target: function () { btnClose.onClick(); }, inFields: true }
    //      });
    //    後片付けを onClose に置くときは DOM に触れない（常駐パレットの onClose で DOM に触ると落ちる）

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
            title: { ja: "キーボードショートカット", en: "Keyboard Shortcuts" }
        },
        panel: {
            align: { ja: "揃え", en: "Align" },
            options: { ja: "オプション", en: "Options" },
            corner: { ja: "角（無効のパネル）", en: "Corners (disabled panel)" }
        },
        radio: {
            left: { ja: "左", en: "Left" },
            center: { ja: "中央", en: "Center" },
            right: { ja: "右", en: "Right" },
            round: { ja: "丸", en: "Round" },
            square: { ja: "角", en: "Square" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            width: { ja: "幅", en: "Width" },
            name: { ja: "名前", en: "Name" }
        },
        tooltip: {
            left: { ja: "左に揃えます", en: "Aligns to the left" },
            center: { ja: "中央に揃えます", en: "Aligns to the center" },
            right: { ja: "右に揃えます", en: "Aligns to the right" },
            preview: { ja: "プレビューを切り替えます", en: "Toggles the preview" },
            reset: { ja: "初期値に戻します", en: "Restores the defaults" },
            width: { ja: "数値の欄。フォーカスがあってもショートカットが効きます", en: "Numeric field; shortcuts still work while it has focus" },
            name: { ja: "文字の欄。フォーカスがある間はショートカットが効きません", en: "Text field; shortcuts are off while it has focus" }
        },
        button: {
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        status: {
            changed: { ja: "揃え：%1 / プレビュー：%2", en: "Align: %1 / Preview: %2" }
        }
    };

    /**
     * 現在の UI 言語のラベルを返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} ラベル
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelEntry) {
        return getLabel(labelEntry) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // デモ / Demo
    // =========================================
    /**
     * デモのダイアログを表示する
     * @returns {void}
     */
    function showDemoDialog() {
        var demoDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        demoDialog.orientation = "column";
        demoDialog.alignChildren = ["fill", "top"];
        demoDialog.margins = 16;
        demoDialog.spacing = 12;

        var alignPanel = demoDialog.add("panel", undefined, getLabel(LABELS.panel.align));
        alignPanel.orientation = "row";
        alignPanel.margins = [15, 20, 15, 10];
        var alignLeftRadio = alignPanel.add("radiobutton", undefined, getLabel(LABELS.radio.left));
        var alignCenterRadio = alignPanel.add("radiobutton", undefined, getLabel(LABELS.radio.center));
        var alignRightRadio = alignPanel.add("radiobutton", undefined, getLabel(LABELS.radio.right));
        alignLeftRadio.helpTip = getLabel(LABELS.tooltip.left);
        alignCenterRadio.helpTip = getLabel(LABELS.tooltip.center);
        alignRightRadio.helpTip = getLabel(LABELS.tooltip.right);
        alignLeftRadio.value = true;

        var optionsPanel = demoDialog.add("panel", undefined, getLabel(LABELS.panel.options));
        optionsPanel.orientation = "column";
        optionsPanel.alignChildren = ["left", "top"];
        optionsPanel.margins = [15, 20, 15, 10];
        var previewCheckbox = optionsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.helpTip = getLabel(LABELS.tooltip.preview);

        var widthRow = optionsPanel.add("group");
        widthRow.add("statictext", undefined, labelText(LABELS.fieldLabel.width));
        var widthInput = widthRow.add("edittext", undefined, "100");
        widthInput.characters = 6;
        widthInput.helpTip = getLabel(LABELS.tooltip.width);

        var nameRow = optionsPanel.add("group");
        nameRow.add("statictext", undefined, labelText(LABELS.fieldLabel.name));
        var nameInput = nameRow.add("edittext", undefined, "");
        nameInput.characters = 12;
        nameInput.helpTip = getLabel(LABELS.tooltip.name);

        /* 親ごと無効のパネル。R / S キーを押しても切り替わらない / Disabled as a whole: R / S do nothing */
        var cornerPanel = demoDialog.add("panel", undefined, getLabel(LABELS.panel.corner));
        cornerPanel.orientation = "row";
        cornerPanel.margins = [15, 20, 15, 10];
        var cornerRoundRadio = cornerPanel.add("radiobutton", undefined, getLabel(LABELS.radio.round));
        var cornerSquareRadio = cornerPanel.add("radiobutton", undefined, getLabel(LABELS.radio.square));
        cornerRoundRadio.value = true;
        cornerPanel.enabled = false;

        var statusText = demoDialog.add("statictext", undefined, "");
        statusText.characters = 30;

        var btnRow = demoDialog.add("group");
        btnRow.alignment = ["fill", "bottom"];
        var btnReset = btnRow.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);
        var spacer = btnRow.add("group");
        spacer.alignment = ["fill", "fill"];
        var btnRightGroup = btnRow.add("group");
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        /**
         * 現在の揃えとプレビューの状態を表示する
         * @returns {void}
         */
        function refreshStatus() {
            var alignName = alignLeftRadio.value ? LABELS.radio.left : (alignCenterRadio.value ? LABELS.radio.center : LABELS.radio.right);
            statusText.text = getLabel(LABELS.status.changed)
                .replace("%1", getLabel(alignName))
                .replace("%2", previewCheckbox.value ? "ON" : "OFF");
        }

        alignLeftRadio.onClick = refreshStatus;
        alignCenterRadio.onClick = refreshStatus;
        alignRightRadio.onClick = refreshStatus;
        previewCheckbox.onClick = refreshStatus;
        btnReset.onClick = function () {
            alignLeftRadio.value = true;
            alignCenterRadio.value = false;
            alignRightRadio.value = false;
            previewCheckbox.value = false;
            widthInput.text = "100";
            refreshStatus();
        };

        addKeyShortcuts(demoDialog, {
            "L": alignLeftRadio,
            "C": alignCenterRadio,
            "R": alignRightRadio,
            "P": previewCheckbox,
            "Shift+R": btnReset,
            "S": cornerSquareRadio,
            /* ⌘＋1 は入力中でも数値欄を空にする / Cmd+1 clears the width even while typing */
            "Cmd+1": { target: function () { widthInput.text = ""; }, inFields: true }
        }, { numericFields: [widthInput], showInTip: true });

        refreshStatus();
        demoDialog.show();
    }

    showDemoDialog();

})();
