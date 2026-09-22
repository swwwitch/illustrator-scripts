#target illustrator
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
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GradientFromFill.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GradientFromFill.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白 [左, 上, 右, 下] / Panel margins [left, top, right, bottom] */
    var PANEL_MARGINS = [15, 20, 15, 10];

    /* 角度の選択肢（ラジオボタンの並び順） / Angle choices in radio order */
    var ANGLE_CHOICES = [0, 30, 45, 60, 90];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "radio.black" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
        }
        return labelNode[uiLang];
    }

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
        styledPanel.orientation = orientation;
        styledPanel.alignChildren = alignChildren;
        styledPanel.margins = PANEL_MARGINS;
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
        gradientDialog.orientation = "column";
        gradientDialog.alignChildren = ["fill", "top"];

        var topGroup = gradientDialog.add("group");
        topGroup.orientation = "row";
        topGroup.alignChildren = ["fill", "top"];

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
        var btnRowGroup = gradientDialog.add("group");
        btnRowGroup.alignment = ["center", "center"]; /* 中央揃え / Center align */
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

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

        addColorKeyHandler(dialogControls, updatePreview);
    }

    /* キー（B / W / T / C / L）と終点カラーの対応 / Shortcut keys for the end color */
    var ENDPOINT_KEY_MODES = { B: "black", W: "white", T: "transparent", C: "complementary", L: "tint" };

    /**
     * キー操作で終点カラーを切り替える（B: 黒、W: 白、T: 透明、C: 補色、L: 淡色）
     * @param {Object} dialogControls - buildDialogUI() の戻り値
     * @param {Function} onChange - 切り替えたあとに呼ぶ関数
     * @returns {void}
     */
    function addColorKeyHandler(dialogControls, onChange) {
        /**
         * 終点カラーのラジオボタンを1つだけ ON にする
         * @param {string} endpointMode - "black" / "white" / "transparent" / "complementary" / "tint"
         * @returns {void}
         */
        function setEndpointMode(endpointMode) {
            dialogControls.blackRadio.value = (endpointMode === "black");
            dialogControls.whiteRadio.value = (endpointMode === "white");
            dialogControls.transparentRadio.value = (endpointMode === "transparent");
            dialogControls.complementaryRadio.value = (endpointMode === "complementary");
            dialogControls.tintRadio.value = (endpointMode === "tint");
            dialogControls.tintSlider.enabled = (endpointMode === "tint");
        }

        dialogControls.gradientDialog.addEventListener("keydown", function (event) {
            if (!ENDPOINT_KEY_MODES.hasOwnProperty(event.keyName)) {
                return;
            }
            setEndpointMode(ENDPOINT_KEY_MODES[event.keyName]);
            event.preventDefault();
            onChange();
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
