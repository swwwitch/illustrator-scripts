#target illustrator
#targetengine "GradientFromFillEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した塗りオブジェクトに対して、元の塗り色を始点にした線形グラデーションを作成します。
終点カラー（黒・白・透明・補色・淡色）と角度を指定でき、セパレートグラデーションや反転にも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GradientFromFill.md

### Overview

Creates a linear gradient on the selected filled objects, starting from their original fill color.
The end color (black, white, transparent, complementary or tint) and the angle are selectable, and separate gradients and reversing are supported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GradientFromFill.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GradientFromFill";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GradientFromFill.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GradientFromFill.md"; /* README (English) */

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

    /* 角度の選択肢（ラジオボタンの並び順） / Angle choices in radio order */
    var ANGLE_CHOICES = [0, 30, 45, 60, 90];

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グラデーション作成", en: "Create Gradient" }
        },
        panel: {
            endColor: { ja: "終点のカラー", en: "End Color" },
            angle: { ja: "角度", en: "Angle" },
            sourceColor: { ja: "始点カラー", en: "Source Color" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            black: { ja: "黒", en: "Black" },
            white: { ja: "白", en: "White" },
            transparent: { ja: "透明", en: "Transparent" },
            complementary: { ja: "補色", en: "Complementary" },
            tint: { ja: "淡色", en: "Tint" }
        },
        dropdown: {
            auto: { ja: "自動（先頭）", en: "Auto (First)" }
        },
        colorName: {
            gray: { ja: "グレー", en: "Gray" },
            spot: { ja: "特色", en: "Spot" }
        },
        checkbox: {
            separateGradient: { ja: "セパレートグラデーション", en: "Separate Gradient" },
            reverse: { ja: "反転", en: "Reverse" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            black: { ja: "終点を黒にします。", en: "Ends the gradient in black." },
            white: { ja: "終点を白にします。", en: "Ends the gradient in white." },
            transparent: {
                ja: "終点の不透明度を0にして、透明へ抜けるグラデーションにします。",
                en: "Fades the gradient out to fully transparent."
            },
            complementary: { ja: "始点カラーの補色を終点にします。", en: "Uses the complement of the source color as the end color." },
            tint: {
                ja: "始点カラーを薄くした色を終点にします。濃度はスライダーで決めます。",
                en: "Ends in a lighter tint of the source color. The slider sets how light."
            },
            tintSlider: {
                ja: "終点に使う濃度（％）です。小さいほど薄くなります。",
                en: "Tint percentage used for the end color. Lower is lighter."
            },
            angle: { ja: "グラデーションの角度です。", en: "Angle of the gradient." },
            sourceDropdown: {
                ja: "始点に使うカラーを選びます。「自動（先頭）」は選択の先頭オブジェクトの塗りを使います。",
                en: "Color used as the gradient start. Auto (First) takes the fill of the first selected object."
            },
            separateGradient: {
                ja: "選択したオブジェクトごとに、それぞれの塗りからグラデーションを作ります。",
                en: "Builds a separate gradient for each selected object from its own fill."
            },
            reverse: { ja: "始点と終点を入れ替えます。", en: "Swaps the start and end colors." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original fills."
            }
        },
        alert: {
            selectObject: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            unsupportedSpotComplementary: {
                ja: "スポットカラーの補色計算には未対応です。",
                en: "Complementary color calculation is not supported for spot colors."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した塗りオブジェクトに、元の塗り色を始点にしたグラデーションを適用する
     * キャンセル時は塗りを元に戻し、作ったグラデーションを削除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            return;
        }

        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        var originalSelection = [];
        for (var i = 0; i < currentSelection.length; i++) {
            originalSelection.push(currentSelection[i]);
        }

        if (currentSelection.length === 0) {
            alert(getLabel("alert.selectObject"));
            return;
        }

        var isCMYK = isDocumentCMYK(doc);
        var targetObjects = [];
        var shouldRevertChanges = false;

        try {
            /* 選択オブジェクトの情報を保持 / Store selected object data */
            targetObjects = collectGradientTargets(currentSelection);

            if (targetObjects.length === 0) {
                return;
            }

            var dialogControls = buildDialogUI();
            populateSourceDropdown(dialogControls.sourceDropdown, targetObjects, isCMYK);
            updateSourcePanelEnabled(dialogControls.sourcePanel, dialogControls.sourceDropdown, targetObjects);

            var updatePreview = createPreviewUpdater(doc, targetObjects, originalSelection, dialogControls, isCMYK);
            bindDialogEvents(dialogControls, updatePreview);

            prepareDialogWindow(dialogControls.gradientDialog, SCRIPT_NAME);
            if (dialogControls.gradientDialog.show() === 1) {
                if (!dialogControls.previewCheckbox.value) {
                    applyGradient(doc, targetObjects, dialogControls, isCMYK);
                }
            } else {
                shouldRevertChanges = true;
            }
        } finally {
            if (shouldRevertChanges) {
                try {
                    restoreOriginal(targetObjects);
                } catch (e) { }
                try {
                    cleanupAppliedGradients(targetObjects);
                } catch (e) { }
            }
            try {
                restoreSelection(doc, originalSelection);
            } catch (e) { }
            app.redraw();
        }
    }

    // =========================================
    // グラデーションの適用と取り消し / Apply and revert
    // =========================================

    /**
     * 終点カラーのラジオボタンから、始点カラーに対する終点の色と不透明度を決める
     * @param {Color} sourceColor - 始点カラー
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {{color: Color, opacity: number}} 終点の色と不透明度
     */
    function resolveEndpoint(sourceColor, dialogControls, isCMYK) {
        var endpointColor = sourceColor;
        var endpointOpacity = 100.0;

        if (dialogControls.transparentRadio.value) {
            endpointOpacity = 0.0;
        } else if (dialogControls.blackRadio.value) {
            endpointColor = createBlackColor(isCMYK);
        } else if (dialogControls.whiteRadio.value) {
            endpointColor = createWhiteColor(isCMYK);
        } else if (dialogControls.complementaryRadio.value) {
            endpointColor = createComplementaryColor(sourceColor, isCMYK);
        } else if (dialogControls.tintRadio.value) {
            endpointColor = createTintColor(sourceColor, dialogControls.tintSlider.value, isCMYK);
        }
        return { color: endpointColor, opacity: endpointOpacity };
    }

    /**
     * グラデーションのストップを設定する
     * @param {GradientStop} gradientStop - 対象のストップ
     * @param {number} rampPoint - 位置（0〜100）
     * @param {Color} stopColor - 色
     * @param {number} stopOpacity - 不透明度（0〜100）
     * @returns {void}
     */
    function setStop(gradientStop, rampPoint, stopColor, stopOpacity) {
        gradientStop.rampPoint = rampPoint;
        gradientStop.color = stopColor;
        gradientStop.opacity = stopOpacity;
    }

    /**
     * 対象オブジェクトすべてにダイアログの設定でグラデーションを適用する（2回目以降は作ったグラデーションを使い回す）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} targetObjects - collectGradientTargets() の結果
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {void}
     */
    function applyGradient(doc, targetObjects, dialogControls, isCMYK) {
        for (var i = 0; i < targetObjects.length; i++) {
            var targetRecord = targetObjects[i];
            var targetItem = targetRecord.item;
            var sourceColor = getSourceColorForFill(targetRecord.originalFillColor, isCMYK, dialogControls.sourceDropdown);
            var endpoint = resolveEndpoint(sourceColor, dialogControls, isCMYK);

            var startColor = sourceColor;
            var startOpacity = 100.0;
            var endColor = endpoint.color;
            var endOpacity = endpoint.opacity;

            if (dialogControls.reverseCheckbox.value) {
                startColor = endpoint.color;
                startOpacity = endpoint.opacity;
                endColor = sourceColor;
                endOpacity = 100.0;
            }

            /* グラデーションの新規作成または再利用 / Create or reuse gradient */
            if (targetRecord.appliedGrad === null) {
                targetRecord.appliedGrad = doc.gradients.add();
                targetRecord.appliedGrad.type = GradientType.LINEAR;
            }
            var activeGradient = targetRecord.appliedGrad;

            var isSeparate = dialogControls.separateCheckbox.value;
            var requiredStops = isSeparate ? 4 : 2;

            while (activeGradient.gradientStops.length < requiredStops) {
                activeGradient.gradientStops.add();
            }
            while (activeGradient.gradientStops.length > requiredStops) {
                activeGradient.gradientStops[activeGradient.gradientStops.length - 1].remove();
            }

            if (isSeparate) {
                /* セパレート（0, 50, 50, 100） / Separate stops (0, 50, 50, 100) */
                setStop(activeGradient.gradientStops[0], 0, startColor, startOpacity);
                setStop(activeGradient.gradientStops[1], 50.0, startColor, startOpacity);
                setStop(activeGradient.gradientStops[2], 50.0, endColor, endOpacity);
                setStop(activeGradient.gradientStops[3], 100.0, endColor, endOpacity);
            } else {
                /* 通常（0, 100） / Standard stops (0, 100) */
                setStop(activeGradient.gradientStops[0], 0, startColor, startOpacity);
                setStop(activeGradient.gradientStops[1], 100.0, endColor, endOpacity);
            }

            var gradientFill = new GradientColor();
            gradientFill.gradient = activeGradient;
            setTargetFillColor(targetItem, gradientFill);

            doc.selection = null;
            targetItem.selected = true;
            applyGradientAngle(targetRecord, gradientFill, getSelectedAngle(dialogControls.angleRadios));
        }
    }

    /**
     * 対象オブジェクトの塗りを元に戻す
     * @param {Object[]} targetObjects - collectGradientTargets() の結果
     * @returns {void}
     */
    function restoreOriginal(targetObjects) {
        for (var i = 0; i < targetObjects.length; i++) {
            var targetRecord = targetObjects[i];
            setTargetFillColor(targetRecord.item, targetRecord.originalFillColor);
            targetRecord.lastAngle = 0;
        }
    }

    /**
     * プレビュー・適用で作ったグラデーションを削除する
     * @param {Object[]} targetObjects - collectGradientTargets() の結果
     * @returns {void}
     */
    function cleanupAppliedGradients(targetObjects) {
        for (var i = 0; i < targetObjects.length; i++) {
            var targetRecord = targetObjects[i];
            if (targetRecord.appliedGrad !== null) {
                try { targetRecord.appliedGrad.remove(); } catch (e) { }
            }
            targetRecord.appliedGrad = null;
            targetRecord.lastAngle = 0;
        }
    }

    /**
     * 選択を元に戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} originalSelection - 元の選択
     * @returns {void}
     */
    function restoreSelection(doc, originalSelection) {
        doc.selection = null;
        for (var i = 0; i < originalSelection.length; i++) {
            try {
                originalSelection[i].selected = true;
            } catch (e) { /* ロック・非表示は選択できない / locked or hidden items cannot be selected */ }
        }
    }

    /**
     * プレビューを更新する関数を作る（ON なら適用、OFF なら元に戻す。失敗したら元に戻して片付ける）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} targetObjects - collectGradientTargets() の結果
     * @param {PageItem[]} originalSelection - 元の選択
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Function} プレビューを更新する関数
     */
    function createPreviewUpdater(doc, targetObjects, originalSelection, dialogControls, isCMYK) {
        return function () {
            try {
                if (dialogControls.previewCheckbox.value) {
                    applyGradient(doc, targetObjects, dialogControls, isCMYK);
                } else {
                    restoreOriginal(targetObjects);
                }
            } catch (e) {
                try {
                    restoreOriginal(targetObjects);
                } catch (restoreErr) { }
                try {
                    cleanupAppliedGradients(targetObjects);
                } catch (cleanupErr) { }
            } finally {
                restoreSelection(doc, originalSelection);
                app.redraw();
            }
        };
    }

    // =========================================
    // UI構築 / UI construction
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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 余白をそろえたパネルを追加する
     * @param {Object} parentContainer - 追加先（Window / Group）
     * @param {string} titlePath - パネル名のラベルのパス
     * @param {string} orientation - "column" / "row"
     * @param {string[]} alignChildren - 子の揃え
     * @returns {Panel} 追加したパネル
     */
    function addStyledPanel(parentContainer, titlePath, orientation, alignChildren) {
        var styledPanel = parentContainer.add("panel", undefined, getLabel(titlePath));
        setupPanel(styledPanel, 6);
        styledPanel.orientation = orientation;
        styledPanel.alignChildren = alignChildren;
        return styledPanel;
    }

    /**
     * tooltip 付きのコントロールを追加する
     * @param {Object} parentContainer - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox" など
     * @param {string} textPath - 表示テキストのラベルのパス
     * @param {string} tipPath - tooltip のラベルのパス
     * @returns {Object} 追加したコントロール
     */
    function addTippedControl(parentContainer, controlType, textPath, tipPath) {
        var tippedControl = parentContainer.add(controlType, undefined, getLabel(textPath));
        tippedControl.helpTip = getLabel(tipPath);
        return tippedControl;
    }

    /**
     * ダイアログを組み立てる
     * @returns {Object} ダイアログと各コントロール
     */
    function buildDialogUI() {
        var gradientDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(gradientDialog);

        var topGroup = gradientDialog.add("group");
        topGroup.orientation = "row";
        topGroup.alignChildren = ["fill", "top"];
        topGroup.spacing = COLUMN_SPACING;

        /* 終点カラーの設定 / Set end color options */
        var endColorPanel = addStyledPanel(topGroup, "panel.endColor", "column", ["left", "top"]);

        var blackRadio = addTippedControl(endColorPanel, "radiobutton", "radio.black", "tooltip.black");
        var whiteRadio = addTippedControl(endColorPanel, "radiobutton", "radio.white", "tooltip.white");
        var transparentRadio = addTippedControl(endColorPanel, "radiobutton", "radio.transparent", "tooltip.transparent");
        var complementaryRadio = addTippedControl(endColorPanel, "radiobutton", "radio.complementary", "tooltip.complementary");

        /* 淡色ラジオ＋スライダー（ラジオは別グループなので排他は手動） / Tint radio + slider (exclusivity handled by hand) */
        var tintLabelGroup = endColorPanel.add("group");
        tintLabelGroup.orientation = "row";
        tintLabelGroup.alignChildren = ["left", "center"];
        tintLabelGroup.spacing = 4;
        var tintRadio = addTippedControl(tintLabelGroup, "radiobutton", "radio.tint", "tooltip.tint");
        var tintValueText = tintLabelGroup.add("statictext", undefined, "50%");
        tintValueText.characters = 5;

        var tintSlider = endColorPanel.add("slider", undefined, 50, 0, 100);
        tintSlider.helpTip = getLabel("tooltip.tintSlider");
        tintSlider.alignment = ["fill", "top"];
        tintSlider.enabled = false;

        transparentRadio.value = true; /* デフォルト / Default */

        /* 角度 / Angle */
        var anglePanel = addStyledPanel(topGroup, "panel.angle", "column", ["left", "top"]);
        var angleRadios = [];
        for (var i = 0; i < ANGLE_CHOICES.length; i++) {
            var angleRadio = anglePanel.add("radiobutton", undefined, String(ANGLE_CHOICES[i]));
            angleRadio.helpTip = getLabel("tooltip.angle");
            angleRadios.push(angleRadio);
        }
        angleRadios[0].value = true; /* デフォルト / Default */

        /* 始点カラー / Source color */
        var sourcePanel = addStyledPanel(gradientDialog, "panel.sourceColor", "row", ["left", "center"]);
        var sourceDropdown = sourcePanel.add("dropdownlist", undefined, [getLabel("dropdown.auto")]);
        sourceDropdown.helpTip = getLabel("tooltip.sourceDropdown");
        sourceDropdown.selection = 0; /* デフォルト / Default */

        /* オプションの設定 / Set options */
        var optionsPanel = addStyledPanel(gradientDialog, "panel.options", "column", ["left", "top"]);
        var separateCheckbox = addTippedControl(optionsPanel, "checkbox", "checkbox.separateGradient", "tooltip.separateGradient");
        var reverseCheckbox = addTippedControl(optionsPanel, "checkbox", "checkbox.reverse", "tooltip.reverse");
        var previewCheckbox = addTippedControl(optionsPanel, "checkbox", "checkbox.preview", "tooltip.preview");

        /* ボタンの設定 / Set button layout */
        var buttonRow = addButtonRow(gradientDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        return {
            gradientDialog: gradientDialog,
            blackRadio: blackRadio,
            whiteRadio: whiteRadio,
            transparentRadio: transparentRadio,
            complementaryRadio: complementaryRadio,
            tintRadio: tintRadio,
            tintSlider: tintSlider,
            tintValueText: tintValueText,
            angleRadios: angleRadios,
            sourcePanel: sourcePanel,
            sourceDropdown: sourceDropdown,
            separateCheckbox: separateCheckbox,
            reverseCheckbox: reverseCheckbox,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * ダイアログのコントロールにイベントをつなぐ
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @param {Function} updatePreview - プレビューを更新する関数
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, updatePreview) {
        var tintRadio = dialogControls.tintRadio;
        var tintSlider = dialogControls.tintSlider;
        var otherEndpointRadios = [
            dialogControls.blackRadio,
            dialogControls.whiteRadio,
            dialogControls.transparentRadio,
            dialogControls.complementaryRadio
        ];

        dialogControls.previewCheckbox.onClick = updatePreview;
        dialogControls.separateCheckbox.onClick = updatePreview;
        dialogControls.reverseCheckbox.onClick = updatePreview;
        for (var i = 0; i < dialogControls.angleRadios.length; i++) {
            dialogControls.angleRadios[i].onClick = updatePreview;
        }
        dialogControls.sourceDropdown.onChange = updatePreview;

        /* 淡色ラジオは別グループなので、ほかの終点ラジオを手動で OFF に / Tint radio sits in another group */
        tintRadio.onClick = function () {
            for (var j = 0; j < otherEndpointRadios.length; j++) {
                otherEndpointRadios[j].value = false;
            }
            tintSlider.enabled = tintRadio.value;
            updatePreview();
        };

        /* 他のラジオ選択時は淡色を OFF にしてスライダーを無効化 / Disable slider when other radios selected */
        for (var k = 0; k < otherEndpointRadios.length; k++) {
            otherEndpointRadios[k].onClick = function () {
                tintRadio.value = false;
                tintSlider.enabled = false;
                updatePreview();
            };
        }

        /**
         * スライダーの値を表示に反映する（Shift で 10% 刻み）
         * @returns {void}
         */
        function onTintSliderChange() {
            if (ScriptUI.environment.keyboardState.shiftKey) {
                tintSlider.value = Math.round(tintSlider.value / 10) * 10;
            }
            dialogControls.tintValueText.text = Math.round(tintSlider.value) + "%";
            updatePreview();
        }
        tintSlider.onChanging = onTintSliderChange;
        tintSlider.onChange = onTintSliderChange;

        /* キーで終点カラーを切り替える（B: 黒、W: 白、T: 透明、C: 補色、L: 淡色）。onClick が淡色との排他を受け持つ
           Switch the end color by key (B / W / T / C / L); the onClick handlers keep the tint radio exclusive */
        addKeyShortcuts(dialogControls.gradientDialog, {
            "B": dialogControls.blackRadio,
            "W": dialogControls.whiteRadio,
            "T": dialogControls.transparentRadio,
            "C": dialogControls.complementaryRadio,
            "L": tintRadio
        });
    }

    // =========================================
    // 始点カラーの候補 / Source color choices
    // =========================================

    /**
     * 始点カラーのドロップダウンを作り直す（1つだけ選択したグラデーションの塗りから、黒・白・透明以外のストップ色を並べる）
     * @param {DropDownList} sourceDropdown - 対象のドロップダウン
     * @param {Object[]} targetObjects - collectGradientTargets() の結果
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {void}
     */
    function populateSourceDropdown(sourceDropdown, targetObjects, isCMYK) {
        if (!sourceDropdown) {
            return;
        }

        removeAllDropdownItems(sourceDropdown);
        sourceDropdown.add("item", getLabel("dropdown.auto"));

        var gradientFill = getSingleGradientFillColor(targetObjects);
        var sourceColors = gradientFill ? collectSourceColorsFromGradient(gradientFill, isCMYK) : [];
        for (var i = 0; i < sourceColors.length; i++) {
            sourceDropdown.add("item", formatSourceColorLabel(sourceColors[i], i));
        }
        sourceDropdown.selection = 0;
    }

    /**
     * ドロップダウンの項目をすべて削除する
     * @param {DropDownList} targetDropdown - 対象のドロップダウン
     * @returns {void}
     */
    function removeAllDropdownItems(targetDropdown) {
        while (targetDropdown.items.length > 0) {
            targetDropdown.remove(targetDropdown.items[0]);
        }
    }

    /**
     * 対象が1つだけで塗りがグラデーションなら、その塗りを返す
     * @param {Object[]} targetObjects - collectGradientTargets() の結果
     * @returns {GradientColor|null} グラデーションの塗り（該当しなければ null）
     */
    function getSingleGradientFillColor(targetObjects) {
        if (!targetObjects || targetObjects.length !== 1) {
            return null;
        }

        var fillColor = targetObjects[0].originalFillColor;
        if (fillColor && fillColor.typename === "GradientColor") {
            return fillColor;
        }
        return null;
    }

    /**
     * 始点カラーのパネルは、グラデーションの塗りを1つだけ選択したときだけ有効にする
     * @param {Panel} sourcePanel - 始点カラーのパネル
     * @param {DropDownList} sourceDropdown - 始点カラーのドロップダウン
     * @param {Object[]} targetObjects - collectGradientTargets() の結果
     * @returns {void}
     */
    function updateSourcePanelEnabled(sourcePanel, sourceDropdown, targetObjects) {
        var isEnabled = !!getSingleGradientFillColor(targetObjects);

        if (sourcePanel) {
            sourcePanel.enabled = isEnabled;
        }
        if (sourceDropdown) {
            sourceDropdown.enabled = isEnabled;
        }
    }

    /**
     * 塗りから始点カラーを決める（グラデーションならドロップダウンで選んだストップ色）
     * @param {Color} fillColor - 元の塗り
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @param {DropDownList} sourceDropdown - 始点カラーのドロップダウン
     * @returns {Color} 始点カラー
     */
    function getSourceColorForFill(fillColor, isCMYK, sourceDropdown) {
        if (!fillColor || fillColor.typename !== "GradientColor") {
            return cloneSimpleColor(fillColor, isCMYK);
        }

        var sourceColors = collectSourceColorsFromGradient(fillColor, isCMYK);
        if (sourceColors.length === 0) {
            return getFallbackGradientColor(fillColor, isCMYK);
        }

        var selectedIndex = getSelectedSourceColorIndex(sourceDropdown);
        if (selectedIndex < 0 || selectedIndex >= sourceColors.length) {
            selectedIndex = 0;
        }

        return cloneSimpleColor(sourceColors[selectedIndex], isCMYK);
    }

    /**
     * グラデーションのストップ色から、黒・白・透明を除いて重複なく集める
     * @param {GradientColor} fillColor - グラデーションの塗り
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Color[]} 候補の色
     */
    function collectSourceColorsFromGradient(fillColor, isCMYK) {
        var sourceColors = [];
        var sourceGradient = fillColor.gradient;
        if (!sourceGradient || !sourceGradient.gradientStops) {
            return sourceColors;
        }

        for (var i = 0; i < sourceGradient.gradientStops.length; i++) {
            var gradientStop = sourceGradient.gradientStops[i];
            if (!isExcludedSourceStop(gradientStop)) {
                var clonedColor = cloneSimpleColor(gradientStop.color, isCMYK);
                if (!containsEquivalentColor(sourceColors, clonedColor)) {
                    sourceColors.push(clonedColor);
                }
            }
        }
        return sourceColors;
    }

    /**
     * 同じ値の色が一覧にあるかを調べる
     * @param {Color[]} colorList - 色の一覧
     * @param {Color} targetColor - 探す色
     * @returns {boolean} あれば true
     */
    function containsEquivalentColor(colorList, targetColor) {
        for (var i = 0; i < colorList.length; i++) {
            if (isSameColorValue(colorList[i], targetColor)) {
                return true;
            }
        }
        return false;
    }

    /**
     * 2つの色が同じ値かを調べる（CMYK / RGB / グレー / 特色）
     * @param {Color} colorA - 1つ目の色
     * @param {Color} colorB - 2つ目の色
     * @returns {boolean} 同じなら true
     */
    function isSameColorValue(colorA, colorB) {
        if (!colorA || !colorB) {
            return false;
        }

        if (colorA.typename !== colorB.typename) {
            return false;
        }

        if (colorA.typename === "CMYKColor") {
            return isSameNumber(colorA.cyan, colorB.cyan) &&
                isSameNumber(colorA.magenta, colorB.magenta) &&
                isSameNumber(colorA.yellow, colorB.yellow) &&
                isSameNumber(colorA.black, colorB.black);
        }

        if (colorA.typename === "RGBColor") {
            return isSameNumber(colorA.red, colorB.red) &&
                isSameNumber(colorA.green, colorB.green) &&
                isSameNumber(colorA.blue, colorB.blue);
        }

        if (colorA.typename === "GrayColor") {
            return isSameNumber(colorA.gray, colorB.gray);
        }

        if (colorA.typename === "SpotColor") {
            var spotNameA = (colorA.spot && colorA.spot.name) ? colorA.spot.name : "";
            var spotNameB = (colorB.spot && colorB.spot.name) ? colorB.spot.name : "";
            return spotNameA === spotNameB && isSameNumber(colorA.tint, colorB.tint);
        }

        return false;
    }

    /**
     * 2つの値が数値として（0.001 未満の差で）等しいかを調べる
     * @param {number} valueA - 1つ目の値
     * @param {number} valueB - 2つ目の値
     * @returns {boolean} 等しければ true（数値でなければ false）
     */
    function isSameNumber(valueA, valueB) {
        var numberA = Number(valueA);
        var numberB = Number(valueB);
        if (isNaN(numberA) || isNaN(numberB)) {
            return false;
        }
        return Math.abs(numberA - numberB) < 0.001;
    }

    /**
     * ドロップダウンで選んだ候補の番号を返す（先頭の「自動」は 0 番目の候補と同じ）
     * @param {DropDownList} sourceDropdown - 始点カラーのドロップダウン
     * @returns {number} 候補の番号
     */
    function getSelectedSourceColorIndex(sourceDropdown) {
        if (!sourceDropdown || !sourceDropdown.selection) {
            return 0;
        }

        if (sourceDropdown.selection.index <= 0) {
            return 0;
        }

        return Math.max(0, sourceDropdown.selection.index - 1);
    }

    /**
     * ドロップダウンに出す候補の表示名を作る（"1: C0 M50 Y100 K0" など）
     * @param {Color} sourceColor - 候補の色
     * @param {number} index - 候補の番号
     * @returns {string} 表示名
     */
    function formatSourceColorLabel(sourceColor, index) {
        return String(index + 1) + ': ' + formatColorLabel(sourceColor);
    }

    /**
     * 色の値を短い文字列にする
     * @param {Color} color - 対象の色
     * @returns {string} "C0 M50 Y100 K0" / "R255 G0 B0" / "グレー 50" / "特色名 100%" など
     */
    function formatColorLabel(color) {
        if (!color) {
            return '-';
        }

        if (color.typename === "CMYKColor") {
            return 'C' + formatColorNumber(color.cyan) + ' M' + formatColorNumber(color.magenta) + ' Y' + formatColorNumber(color.yellow) + ' K' + formatColorNumber(color.black);
        }

        if (color.typename === "RGBColor") {
            return 'R' + formatColorNumber(color.red) + ' G' + formatColorNumber(color.green) + ' B' + formatColorNumber(color.blue);
        }

        if (color.typename === "GrayColor") {
            return getLabel("colorName.gray") + ' ' + formatColorNumber(color.gray);
        }

        if (color.typename === "SpotColor") {
            var spotName = (color.spot && color.spot.name) ? color.spot.name : getLabel("colorName.spot");
            return spotName + ' ' + formatColorNumber(color.tint) + '%';
        }

        return color.typename;
    }

    /**
     * 色の値を小数1桁までの文字列にする（整数なら小数点なし）
     * @param {number} value - 値
     * @returns {string} 表示用の文字列（数値でなければ "0"）
     */
    function formatColorNumber(value) {
        if (typeof value !== "number") {
            return '0';
        }

        var rounded = Math.round(value * 10) / 10;
        if (Math.abs(rounded - Math.round(rounded)) < 0.001) {
            return String(Math.round(rounded));
        }
        return String(rounded);
    }

    /**
     * 始点カラーの候補から外すストップか（不透明度 0、黒、白）
     * @param {GradientStop} gradientStop - 対象のストップ
     * @returns {boolean} 外すなら true
     */
    function isExcludedSourceStop(gradientStop) {
        if (!gradientStop) {
            return true;
        }
        if (typeof gradientStop.opacity === "number" && gradientStop.opacity <= 0) {
            return true;
        }
        return isBlackOrWhiteColor(gradientStop.color);
    }

    /**
     * 黒または白かを判定する
     * @param {Color} color - 対象の色
     * @returns {boolean} 黒・白なら true
     */
    function isBlackOrWhiteColor(color) {
        if (!color) {
            return false;
        }

        if (color.typename === "RGBColor") {
            var isBlackRgb = (color.red === 0 && color.green === 0 && color.blue === 0);
            var isWhiteRgb = (color.red === 255 && color.green === 255 && color.blue === 255);
            return isBlackRgb || isWhiteRgb;
        }

        if (color.typename === "CMYKColor") {
            var isBlackCmyk = (color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 100);
            var isWhiteCmyk = (color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0);
            return isBlackCmyk || isWhiteCmyk;
        }

        if (color.typename === "GrayColor") {
            return color.gray === 0 || color.gray === 100;
        }

        return false;
    }

    /**
     * 候補が無いグラデーションの始点カラー（先頭ストップの色、ストップが無ければ黒）
     * @param {GradientColor} fillColor - グラデーションの塗り
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Color} 始点カラー
     */
    function getFallbackGradientColor(fillColor, isCMYK) {
        var sourceGradient = fillColor.gradient;
        if (sourceGradient && sourceGradient.gradientStops && sourceGradient.gradientStops.length > 0) {
            return cloneSimpleColor(sourceGradient.gradientStops[0].color, isCMYK);
        }
        return createBlackColor(isCMYK);
    }

    // =========================================
    // 対象オブジェクト / Target objects
    // =========================================

    /**
     * ドキュメントが CMYK かを調べる
     * @param {Document} doc - 対象ドキュメント
     * @returns {boolean} CMYK なら true
     */
    function isDocumentCMYK(doc) {
        return doc.documentColorSpace === DocumentColorSpace.CMYK;
    }

    /**
     * グラデーションを適用できるオブジェクトか（塗りのあるパス、塗りのあるパスを含む複合パス。クリップパスは除く）
     * @param {PageItem} candidateItem - 対象のオブジェクト
     * @returns {boolean} 適用できるなら true
     */
    function isGradientTarget(candidateItem) {
        if (!candidateItem) {
            return false;
        }

        if (candidateItem.typename === "PathItem") {
            return candidateItem.filled && !candidateItem.clipping;
        }

        if (candidateItem.typename === "CompoundPathItem") {
            return !!findFirstFilledPath(candidateItem);
        }

        return false;
    }

    /**
     * 複合パスの中で、塗りがありクリップパスでない最初のパスを返す
     * @param {CompoundPathItem} compoundPath - 対象の複合パス
     * @returns {PathItem|null} 見つかったパス（無ければ null）
     */
    function findFirstFilledPath(compoundPath) {
        if (!compoundPath || !compoundPath.pathItems || compoundPath.pathItems.length === 0) {
            return null;
        }

        for (var i = 0; i < compoundPath.pathItems.length; i++) {
            var pathItem = compoundPath.pathItems[i];
            if (pathItem.filled && !pathItem.clipping) {
                return pathItem;
            }
        }
        return null;
    }

    /**
     * 対象オブジェクトの記録を作る
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {{item: PageItem, originalFillColor: Color, appliedGrad: Gradient, lastAngle: number}} 記録
     */
    function createGradientTargetRecord(targetItem) {
        return {
            item: targetItem,
            originalFillColor: getTargetFillColor(targetItem),
            appliedGrad: null,
            lastAngle: 0
        };
    }

    /**
     * 選択からグラデーションを適用するオブジェクトを集める（グループの中もたどる）
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {Object[]} 対象オブジェクトの記録
     */
    function collectGradientTargets(selectedItems) {
        var targetObjects = [];
        for (var i = 0; i < selectedItems.length; i++) {
            collectGradientTargetsFromItem(selectedItems[i], targetObjects);
        }
        return targetObjects;
    }

    /**
     * 1つのオブジェクトから対象を集める（グループは再帰）
     * @param {PageItem} sourceItem - 対象のオブジェクト
     * @param {Object[]} targetObjects - 記録を追加する配列
     * @returns {void}
     */
    function collectGradientTargetsFromItem(sourceItem, targetObjects) {
        if (!sourceItem) {
            return;
        }

        if (sourceItem.typename === "GroupItem") {
            for (var i = 0; i < sourceItem.pageItems.length; i++) {
                collectGradientTargetsFromItem(sourceItem.pageItems[i], targetObjects);
            }
            return;
        }

        if (isGradientTarget(sourceItem)) {
            targetObjects.push(createGradientTargetRecord(sourceItem));
        }
    }

    /**
     * 対象オブジェクトの塗りを読む（複合パスは最初の塗りのあるパス）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {Color|null} 塗り
     */
    function getTargetFillColor(targetItem) {
        if (!targetItem) {
            return null;
        }

        if (targetItem.typename === "CompoundPathItem") {
            return getCompoundPathFillColor(targetItem);
        }

        return targetItem.fillColor;
    }

    /**
     * 対象オブジェクトに塗りを設定する（複合パスはクリップパス以外のすべてのパス）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {Color} fillColor - 設定する塗り
     * @returns {void}
     */
    function setTargetFillColor(targetItem, fillColor) {
        if (!targetItem) {
            return;
        }

        if (targetItem.typename === "CompoundPathItem") {
            setCompoundPathFillColor(targetItem, fillColor);
            return;
        }

        targetItem.fillColor = fillColor;
    }

    /**
     * 複合パスの塗りを読む（最初の塗りのあるパス、無ければ先頭のパス）
     * @param {CompoundPathItem} compoundPath - 対象の複合パス
     * @returns {Color|null} 塗り
     */
    function getCompoundPathFillColor(compoundPath) {
        if (!compoundPath || !compoundPath.pathItems || compoundPath.pathItems.length === 0) {
            return null;
        }

        var filledPath = findFirstFilledPath(compoundPath);
        return filledPath ? filledPath.fillColor : compoundPath.pathItems[0].fillColor;
    }

    /**
     * 複合パスのクリップパス以外のすべてのパスに塗りを設定する
     * @param {CompoundPathItem} compoundPath - 対象の複合パス
     * @param {Color} fillColor - 設定する塗り
     * @returns {void}
     */
    function setCompoundPathFillColor(compoundPath, fillColor) {
        if (!compoundPath || !compoundPath.pathItems || compoundPath.pathItems.length === 0) {
            return;
        }

        for (var i = 0; i < compoundPath.pathItems.length; i++) {
            var pathItem = compoundPath.pathItems[i];
            if (!pathItem.clipping) {
                pathItem.fillColor = fillColor;
            }
        }
    }

    // =========================================
    // 角度 / Angle
    // =========================================

    /**
     * 選択中の角度を返す
     * @param {RadioButton[]} angleRadios - 角度のラジオボタン（ANGLE_CHOICES と同じ順）
     * @returns {number} 角度（度）
     */
    function getSelectedAngle(angleRadios) {
        for (var i = 0; i < angleRadios.length; i++) {
            if (angleRadios[i].value) {
                return ANGLE_CHOICES[i];
            }
        }
        return 0;
    }

    /**
     * グラデーションの角度を設定する（前回の角度との差だけグラデーションを回転する）
     * @param {Object} targetRecord - 対象オブジェクトの記録（lastAngle を更新する）
     * @param {GradientColor} gradientFill - 適用した塗り
     * @param {number} angle - 角度（度）
     * @returns {void}
     */
    function applyGradientAngle(targetRecord, gradientFill, angle) {
        var previousAngle = (typeof targetRecord.lastAngle === "number") ? targetRecord.lastAngle : 0;
        var deltaAngle = angle - previousAngle;

        gradientFill.angle = angle;

        if (deltaAngle !== 0) {
            rotateTargetGradient(targetRecord.item, deltaAngle);
        }

        targetRecord.lastAngle = angle;
    }

    /**
     * オブジェクトは動かさず、塗りのグラデーションだけを中心で回転する（パスの無い複合パスは何もしない）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {number} deltaAngle - 回転角（度）
     * @returns {void}
     */
    function rotateTargetGradient(targetItem, deltaAngle) {
        if (!targetItem || deltaAngle === 0) {
            return;
        }

        if (targetItem.typename === "CompoundPathItem" && (!targetItem.pathItems || targetItem.pathItems.length === 0)) {
            return;
        }

        targetItem.rotate(deltaAngle, false, false, true, false, Transformation.CENTER);
    }

    // =========================================
    // 色の作成 / Color creation
    // =========================================

    /**
     * CMYK カラーを作る
     * @param {number} cyan - C
     * @param {number} magenta - M
     * @param {number} yellow - Y
     * @param {number} black - K
     * @returns {CMYKColor} CMYK カラー
     */
    function makeCMYKColor(cyan, magenta, yellow, black) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cyan;
        cmykColor.magenta = magenta;
        cmykColor.yellow = yellow;
        cmykColor.black = black;
        return cmykColor;
    }

    /**
     * RGB カラーを作る
     * @param {number} red - R
     * @param {number} green - G
     * @param {number} blue - B
     * @returns {RGBColor} RGB カラー
     */
    function makeRGBColor(red, green, blue) {
        var rgbColor = new RGBColor();
        rgbColor.red = red;
        rgbColor.green = green;
        rgbColor.blue = blue;
        return rgbColor;
    }

    /**
     * グレーカラーを作る
     * @param {number} grayValue - 濃度
     * @returns {GrayColor} グレーカラー
     */
    function makeGrayColor(grayValue) {
        var grayColor = new GrayColor();
        grayColor.gray = grayValue;
        return grayColor;
    }

    /**
     * 特色を作る
     * @param {Spot} sourceSpot - スポット
     * @param {number} tint - 濃度
     * @returns {SpotColor} 特色
     */
    function makeSpotColor(sourceSpot, tint) {
        var spotColor = new SpotColor();
        spotColor.spot = sourceSpot;
        spotColor.tint = tint;
        return spotColor;
    }

    /**
     * 単色（CMYK / RGB / グレー / 特色）を複製する。それ以外や空なら黒
     * @param {Color} color - 元の色
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Color} 複製した色
     */
    function cloneSimpleColor(color, isCMYK) {
        if (!color) {
            return createBlackColor(isCMYK);
        }

        if (color.typename === "CMYKColor") {
            return makeCMYKColor(color.cyan, color.magenta, color.yellow, color.black);
        }

        if (color.typename === "RGBColor") {
            return makeRGBColor(color.red, color.green, color.blue);
        }

        if (color.typename === "GrayColor") {
            return makeGrayColor(color.gray);
        }

        if (color.typename === "SpotColor") {
            return makeSpotColor(color.spot, color.tint);
        }

        return createBlackColor(isCMYK);
    }

    /**
     * 黒を作る
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Color} K100 または RGB 0,0,0
     */
    function createBlackColor(isCMYK) {
        return isCMYK ? makeCMYKColor(0, 0, 0, 100) : makeRGBColor(0, 0, 0);
    }

    /**
     * 白を作る
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Color} CMYK すべて 0 または RGB 255,255,255
     */
    function createWhiteColor(isCMYK) {
        return isCMYK ? makeCMYKColor(0, 0, 0, 0) : makeRGBColor(255, 255, 255);
    }

    /**
     * 始点カラーを薄くした色を作る
     * @param {Color} sourceColor - 始点カラー
     * @param {number} amount - 濃度（％。0〜100）
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Color} 淡色（未対応の種類は白）
     */
    function createTintColor(sourceColor, amount, isCMYK) {
        var ratio = Math.max(0, Math.min(100, amount)) / 100;
        if (sourceColor.typename === "CMYKColor") {
            return makeCMYKColor(sourceColor.cyan * ratio, sourceColor.magenta * ratio, sourceColor.yellow * ratio, sourceColor.black * ratio);
        } else if (sourceColor.typename === "RGBColor") {
            return makeRGBColor(
                Math.round(sourceColor.red + (255 - sourceColor.red) * (1 - ratio)),
                Math.round(sourceColor.green + (255 - sourceColor.green) * (1 - ratio)),
                Math.round(sourceColor.blue + (255 - sourceColor.blue) * (1 - ratio))
            );
        } else if (sourceColor.typename === "GrayColor") {
            return makeGrayColor(sourceColor.gray * ratio);
        } else if (sourceColor.typename === "SpotColor") {
            return makeSpotColor(sourceColor.spot, sourceColor.tint * ratio);
        }
        return createWhiteColor(isCMYK);
    }

    /**
     * 始点カラーの補色を作る（特色は例外を投げる）
     * @param {Color} sourceColor - 始点カラー
     * @param {boolean} isCMYK - CMYK ドキュメントか
     * @returns {Color} 補色（未対応の種類は黒）
     */
    function createComplementaryColor(sourceColor, isCMYK) {
        if (sourceColor.typename === "CMYKColor") {
            return makeCMYKColor(100 - sourceColor.cyan, 100 - sourceColor.magenta, 100 - sourceColor.yellow, sourceColor.black);
        } else if (sourceColor.typename === "RGBColor") {
            return makeRGBColor(255 - sourceColor.red, 255 - sourceColor.green, 255 - sourceColor.blue);
        } else if (sourceColor.typename === "GrayColor") {
            return makeGrayColor(100 - sourceColor.gray);
        } else if (sourceColor.typename === "SpotColor") {
            throw new Error(getLabel("alert.unsupportedSpotComplementary"));
        }
        return createBlackColor(isCMYK);
    }

    main();

})();
