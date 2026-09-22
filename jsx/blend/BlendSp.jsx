#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択内容に応じて、ブレンドの作成・設定・調整を1つのダイアログで行います。
ステップ数と方向はライブプレビューで確認でき、解除／拡張／ブレンド軸の置き換えは［OK］で確定したときだけ実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendSp.md

### Overview

Creates, configures and adjusts a blend from a single dialog, depending on what is selected.
Steps and orientation are shown as a live preview, while Release, Expand and Replace Spine run only when the dialog is confirmed.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BlendSp.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "BlendSp";                      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BlendSp.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BlendSp.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php
(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ブレンド以外を選択して開いたときのステップ数 / Steps used when no blend is selected at launch */
    var DEFAULT_BLEND_STEPS = 8;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;              /* ダイアログの余白 / dialog margins */
    var DIALOG_SPACING = 10;              /* ダイアログ内の要素間隔 / dialog spacing */
    var STEP_AREA_SPACING = 6;            /* ステップ数の行とスライダーの間隔 / gap between the Steps row and the slider */
    var COLUMN_SPACING = 12;              /* 2カラムの間隔 / gap between the two columns */
    var COLUMN_STACK_SPACING = 10;        /* カラム内のパネル間隔 / gap between panels in a column */
    var PANEL_MARGINS = [12, 18, 12, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var OPTION_LIST_SPACING = 4;          /* ラジオ・チェックボックスの行間 / gap between radios and checkboxes */

    /**
     * オプションパネルの共通設定
     * @param {Panel} optionPanel - 対象パネル
     * @param {string} horizontalAlign - 子の横方向の揃え（"fill" / "left"）
     * @param {number} spacing - 要素間隔
     * @returns {void}
     */
    function setupOptionPanel(optionPanel, horizontalAlign, spacing) {
        optionPanel.orientation = 'column';
        optionPanel.alignChildren = [horizontalAlign, 'top'];
        optionPanel.spacing = spacing;
        optionPanel.margins = PANEL_MARGINS;
    }

    /**
     * パネルを縦に積むカラムを追加する
     * @param {Group} columnsGroup - 追加先の横並びグループ
     * @returns {Group} 追加したカラム
     */
    function addColumnGroup(columnsGroup) {
        var columnGroup = columnsGroup.add('group');
        columnGroup.orientation = 'column';
        columnGroup.alignChildren = ['fill', 'top'];
        columnGroup.spacing = COLUMN_STACK_SPACING;
        return columnGroup;
    }

    /**
     * ラジオ・チェックボックスを縦に並べるグループを追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @returns {Group} 追加したグループ
     */
    function addOptionList(parentContainer) {
        var optionList = parentContainer.add('group');
        optionList.orientation = 'column';
        optionList.alignChildren = ['left', 'top'];
        optionList.spacing = OPTION_LIST_SPACING;
        return optionList;
    }

    /**
     * ラベルと tooltip 付きのコントロールを追加する
     * @param {Group} optionList - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox"
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - tooltip のパス
     * @returns {RadioButton|Checkbox} 追加したコントロール
     */
    function addOptionControl(optionList, controlType, labelPath, tooltipPath) {
        var optionControl = optionList.add(controlType, undefined, getLabel(labelPath));
        optionControl.helpTip = getLabel(tooltipPath);
        return optionControl;
    }

    // =========================================
    // ブレンドの設定 / Blend settings
    // =========================================

    /* ブレンドオプション > 方向（0: Align to Page〈垂直方向〉/ 1: Align to Path〈パスに沿う〉）
       Blend Options > Orientation */
    var BlendOrientation = {
        'Align to Page': 0,
        'Align to Path': 1
    };

    /* ステップ数の上限 / Maximum number of steps */
    var MAX_BLEND_STEPS = 1000;

    /* ステップ数スライダーの上限（通常／Option 併用／Shift 併用）/ Slider maximum: plain / with Option / with Shift */
    var STEP_SLIDER_MAX = { normal: 32, option: 128, shift: 1000 };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "ブレンドSpecial", en: "Blend Special" }
        },
        panel: {
            orientation: { ja: "方向", en: "Orientation" },
            adjust: { ja: "その他", en: "Misc" },
            reverse: { ja: "反転", en: "Reverse" }
        },
        fieldLabel: {
            steps: { ja: "ステップ数", en: "Steps" }
        },
        radio: {
            alignToPage: { ja: "垂直方向", en: "Align to Page" },
            alignToPath: { ja: "パスに沿う", en: "Align to Path" },
            adjustNone: { ja: "なし", en: "None" },
            adjustRelease: { ja: "解除", en: "Release" },
            adjustExpand: { ja: "拡張", en: "Expand" },
            adjustReplace: { ja: "ブレンド軸を置き換え", en: "Replace Spine" }
        },
        checkbox: {
            reverseSpine: { ja: "ブレンド軸を反転", en: "Reverse Spine" },
            reverseStack: { ja: "前後を反転", en: "Reverse Front to Back" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            step: {
                ja: "ブレンドの中間オブジェクト数です。0 で中間なし。スライダーでも変えられます。",
                en: "How many intermediate objects the blend creates. 0 means none. The slider changes it too."
            },
            alignToPage: {
                ja: "中間オブジェクトの向きを、ページの垂直方向に固定します。",
                en: "Keeps the intermediate objects upright with respect to the page."
            },
            alignToPath: {
                ja: "中間オブジェクトの向きを、スパイン（軸のパス）の傾きに合わせます。",
                en: "Rotates the intermediate objects to follow the spine."
            },
            adjustNone: { ja: "ブレンドをそのまま残します。", en: "Leaves the blend as a live blend." },
            adjustRelease: {
                ja: "ブレンドを解除して、元のオブジェクトとスパインに戻します。",
                en: "Releases the blend back into the original objects and the spine."
            },
            adjustExpand: {
                ja: "ブレンドを分割・拡張して、中間オブジェクトを実体のあるパスにします。",
                en: "Expands the blend so the intermediate steps become real paths."
            },
            adjustReplace: {
                ja: "スパインを、選択しておいた別のパスに置き換えます。",
                en: "Replaces the spine with another path you selected."
            },
            reverseSpine: {
                ja: "スパインの向きを反転し、始点と終点を入れ替えます。",
                en: "Reverses the spine so the start and the end swap places."
            },
            reverseStack: { ja: "ブレンド内の重ね順を逆にします。", en: "Reverses the stacking order inside the blend." }
        },
        alert: {
            invalidStep: { ja: "ステップ数は 0〜1000 の整数で入力してください。", en: "Please enter an integer from 0 to 1000." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.orientation" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split('.');
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (labelNode === undefined || labelNode === null) return labelPath;
        }
        var labelValue = labelNode[uiLang];
        if (typeof labelValue !== 'string') labelValue = labelNode.en;
        return (typeof labelValue === 'string') ? labelValue : labelPath;
    }

    // =========================================
    // 選択とブレンドの判定 / Selection and blend helpers
    // =========================================

    /**
     * 選択から最初の BlendItem または PluginItem を返す（BlendItem を優先）
     * @param {PageItem[]} selection - 探索する選択オブジェクト
     * @returns {PageItem|null} 見つかったブレンドオブジェクト。無ければ null
     */
    function findFirstBlendObject(selection) {
        try {
            if (!selection || selection.length <= 0) {
                return null;
            }
            /* BlendItem を優先（blendOptions.steps を持つことが多い）/ Prefer BlendItem; it usually exposes blendOptions.steps */
            for (var i = 0; i < selection.length; i++) {
                if (selection[i] && selection[i].typename === 'BlendItem') {
                    return selection[i];
                }
            }
            for (var j = 0; j < selection.length; j++) {
                if (selection[j] && selection[j].typename === 'PluginItem') {
                    return selection[j];
                }
            }
        } catch (e) { }
        return null;
    }

    /**
     * 選択がブレンドオブジェクトのみで構成されているかを判定する
     *
     * 「ブレンド軸を置き換え」の有効／無効制御に使う。ブレンドとパスの混在選択では false になる。
     * @param {PageItem[]} selection - 判定する選択オブジェクト
     * @returns {boolean} すべてブレンドなら true
     */
    function isOnlyBlendObjects(selection) {
        try {
            if (!selection || selection.length <= 0) {
                return false;
            }
            /* 1つ以上あり、全てが BlendItem / PluginItem のとき「ブレンドのみ」/ Every item is a blend */
            for (var i = 0; i < selection.length; i++) {
                var selectedItem = selection[i];
                if (!selectedItem) {
                    return false;
                }
                var itemType = selectedItem.typename;
                if (itemType !== 'PluginItem' && itemType !== 'BlendItem') {
                    return false;
                }
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 指定したアイテムだけを選択状態にする（失敗しても例外は投げない）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} targetItem - 選択したいアイテム
     * @returns {void}
     */
    function selectOnlyItem(doc, targetItem) {
        try {
            if (!doc || !targetItem) {
                return;
            }
            doc.selection = null;
            targetItem.selected = true;
        } catch (e) { }
    }

    /**
     * 現在の選択を配列に控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択していたアイテム
     */
    function snapshotSelection(doc) {
        var selectedItems = [];
        try {
            var currentSelection = doc.selection;
            if (currentSelection && currentSelection.length) {
                for (var i = 0; i < currentSelection.length; i++) {
                    selectedItems.push(currentSelection[i]);
                }
            }
        } catch (e) { }
        return selectedItems;
    }

    /**
     * snapshotSelection() で控えた選択へ戻す（選択できないアイテムは飛ばす）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 控えておいたアイテム
     * @returns {void}
     */
    function restoreSelection(doc, selectedItems) {
        try {
            doc.selection = null;
            for (var i = 0; i < selectedItems.length; i++) {
                /* ロック・非表示などで選択できないアイテムは飛ばす / Skip items that cannot be selected */
                try { selectedItems[i].selected = true; } catch (err) { }
            }
        } catch (e) { }
    }

    /**
     * 選択のうち最前面のパスを削除する
     * @param {PageItem[]} selection - 対象の選択オブジェクト
     * @returns {boolean} 削除できた場合は true、対象が無い場合は false
     */
    function deleteFrontmostPath(selection) {
        try {
            if (!selection || selection.length <= 0) {
                return false;
            }

            var frontmostPath = null;
            var frontmostZ = null;

            for (var i = 0; i < selection.length; i++) {
                var selectedItem = selection[i];
                if (!selectedItem || selectedItem.typename !== 'PathItem') {
                    continue;
                }

                var zPosition = null;
                try {
                    zPosition = selectedItem.zOrderPosition;
                } catch (err) {
                    zPosition = null;
                }

                /* zOrderPosition が大きいものを優先。取れなければ先に見つかったもの（比較できる値が無いうちは後のもので置き換え）
                   Prefer the higher zOrderPosition; while no comparable value exists, the later path takes over */
                var hasFrontmostZ = (frontmostZ !== null && frontmostZ !== undefined);
                var hasZ = (zPosition !== null && zPosition !== undefined);
                if (!hasFrontmostZ || (hasZ && zPosition > frontmostZ)) {
                    frontmostPath = selectedItem;
                    frontmostZ = zPosition;
                }
            }

            if (frontmostPath) {
                frontmostPath.remove();
                return true;
            }
        } catch (e) { }

        return false;
    }

    /**
     * ブレンドのメニューコマンドを実行し、選択に残った最前面のパスを削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} menuCommand - 実行するメニューコマンド
     * @returns {void}
     */
    function runBlendCommandAndDeleteFrontmostPath(doc, menuCommand) {
        app.executeMenuCommand(menuCommand);
        /* コマンド実行後に選択が変わる可能性があるため、doc.selection を参照する
           The command may change the selection, so read doc.selection afresh */
        try {
            deleteFrontmostPath(doc.selection);
        } catch (e) { }
    }

    /**
     * ステップ数を 0〜MAX_BLEND_STEPS に収める
     * @param {number} stepValue - 対象の値
     * @returns {number} 範囲に収めた値
     */
    function clampBlendSteps(stepValue) {
        if (stepValue < 0) return 0;
        if (stepValue > MAX_BLEND_STEPS) return MAX_BLEND_STEPS;
        return stepValue;
    }

    /**
     * PluginItem のブレンドから中間ステップ数を取得する
     *
     * 元データを壊さないよう、複製してから分割・拡張してカウントし、後片付けする。
     * @param {PageItem} blendItem - ステップ数を調べるブレンド（PluginItem）
     * @returns {number} 中間ステップ数。取得できない場合は -1
     */
    function getBlendStepsFromPluginItem(blendItem) {
        /* PluginItem 以外は対象外 / Only PluginItem is supported */
        if (!blendItem || blendItem.typename !== 'PluginItem') {
            return -1;
        }

        var doc = app.activeDocument;
        var stepCount = -1;
        var duplicatedBlend = null;

        /* 選択状態を退避（この関数内で選択を変更するため）/ Save the selection; this function changes it */
        var originalSelection = snapshotSelection(doc);

        try {
            /* 1) 複製（元データ保護）/ Duplicate to protect the original */
            duplicatedBlend = blendItem.duplicate();

            /* 2) 複製したものだけを選択 / Select only the duplicate */
            doc.selection = null;
            duplicatedBlend.selected = true;

            /* 3) 分割・拡張 / Expand */
            app.executeMenuCommand('Path Blend Expand');

            /* 4) 分割・拡張後の選択（展開結果）を確認 / Check the expanded result */
            if (doc.selection && doc.selection.length > 0) {
                /* 5) トップレベルのグループを解除（ネストまでは無理に追わない）/ Ungroup the top level only */
                try { app.executeMenuCommand('ungroup'); } catch (err) { }

                /* 6) 数を数え、7) 始点・終点を除外（全要素数 - 2）/ Count, excluding the start and end objects */
                var expandedCount = doc.selection.length;
                stepCount = (expandedCount >= 2) ? (expandedCount - 2) : 0;
            }
        } catch (e) {
            stepCount = -1;
        } finally {
            /* 展開残骸を削除 / Remove what the expansion left */
            try {
                var expandedItems = doc.selection;
                if (expandedItems && expandedItems.length) {
                    for (var i = 0; i < expandedItems.length; i++) {
                        try { expandedItems[i].remove(); } catch (removeError) { }
                    }
                }
            } catch (selectionError) { }

            /* 複製が残っていた場合の保険 / In case the duplicate survived */
            try {
                if (duplicatedBlend) { duplicatedBlend.remove(); }
            } catch (duplicateError) { }

            /* 選択状態を復元 / Restore the selection */
            restoreSelection(doc, originalSelection);
        }

        return stepCount;
    }

    /**
     * ダイアログのステップ数の初期値を決める
     *
     * ブレンドを選択して開いたときは、そのステップ数を読めれば使う（読めなければ既定値）。
     * @param {PageItem[]} selection - 実行時の選択オブジェクト
     * @param {boolean} wasBlendSelectedAtOpen - ダイアログを開いた時点でブレンドを選択していたか
     * @returns {number} ステップ数の初期値
     */
    function readDefaultStep(selection, wasBlendSelectedAtOpen) {
        var defaultStep = DEFAULT_BLEND_STEPS;
        if (!wasBlendSelectedAtOpen) {
            return defaultStep;
        }
        try {
            var blendObject = findFirstBlendObject(selection);
            /* BlendItem は blendOptions.steps を持つことが多い。PluginItem は持たないことがある
               BlendItem typically exposes blendOptions.steps; PluginItem may not */
            if (blendObject && blendObject.blendOptions && typeof blendObject.blendOptions.steps === 'number') {
                defaultStep = blendObject.blendOptions.steps;
            } else if (blendObject && blendObject.typename === 'PluginItem') {
                var countedSteps = getBlendStepsFromPluginItem(blendObject);
                if (typeof countedSteps === 'number' && countedSteps >= 0) {
                    defaultStep = countedSteps;
                }
            }
        } catch (e) { }
        return defaultStep;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ステップ数の入力欄とスライダーを組み立てる
     * @param {Window} blendDialog - 追加先のダイアログ
     * @param {number} defaultStep - ステップ数の初期値
     * @returns {{stepInput: EditText, stepSlider: Slider}} 入力欄とスライダー
     */
    function buildStepArea(blendDialog, defaultStep) {
        var stepArea = blendDialog.add('group');
        stepArea.orientation = 'column';
        stepArea.alignChildren = ['fill', 'top'];
        stepArea.spacing = STEP_AREA_SPACING;

        var stepRow = stepArea.add('group');
        stepRow.orientation = 'row';
        stepRow.alignChildren = ['left', 'center'];
        stepRow.add('statictext', undefined, getLabel('fieldLabel.steps'));

        var stepInput = stepRow.add('edittext', undefined, String(defaultStep));
        stepInput.helpTip = getLabel('tooltip.step');
        stepInput.characters = 4;

        /* スライダー（ステップ数の下、横いっぱい）/ Slider under Steps, full width */
        var stepSlider = stepArea.add('slider', undefined, defaultStep, 0, STEP_SLIDER_MAX.normal);
        stepSlider.helpTip = getLabel('tooltip.step');
        stepSlider.alignment = ['fill', 'center'];

        return { stepInput: stepInput, stepSlider: stepSlider };
    }

    /**
     * 「方向」「反転」（左カラム）と「その他」（右カラム）のパネルを組み立てる
     * @param {Window} blendDialog - 追加先のダイアログ
     * @returns {Object} 各パネルとラジオ・チェックボックス
     */
    function buildOptionColumns(blendDialog) {
        var columnsGroup = blendDialog.add('group');
        columnsGroup.orientation = 'row';
        columnsGroup.alignChildren = ['fill', 'top'];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = addColumnGroup(columnsGroup);
        var rightColumnGroup = addColumnGroup(columnsGroup);

        /* 方向 / Orientation */
        var orientationPanel = leftColumnGroup.add('panel', undefined, getLabel('panel.orientation'));
        setupOptionPanel(orientationPanel, 'fill', 8);

        var orientationRow = orientationPanel.add('group');
        orientationRow.orientation = 'row';
        orientationRow.alignChildren = ['left', 'top'];

        var orientationList = addOptionList(orientationRow);
        var alignToPageRadio = addOptionControl(orientationList, 'radiobutton', 'radio.alignToPage', 'tooltip.alignToPage');
        var alignToPathRadio = addOptionControl(orientationList, 'radiobutton', 'radio.alignToPath', 'tooltip.alignToPath');
        alignToPageRadio.value = true;

        /* その他（解除・拡張・ブレンド軸の置き換え）/ Misc: release, expand, replace spine */
        var adjustPanel = rightColumnGroup.add('panel', undefined, getLabel('panel.adjust'));
        setupOptionPanel(adjustPanel, 'left', 6);

        var adjustList = addOptionList(adjustPanel);
        var adjustNoneRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustNone', 'tooltip.adjustNone');
        var adjustReleaseRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustRelease', 'tooltip.adjustRelease');
        var adjustExpandRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustExpand', 'tooltip.adjustExpand');
        var adjustReplaceRadio = addOptionControl(adjustList, 'radiobutton', 'radio.adjustReplace', 'tooltip.adjustReplace');
        adjustNoneRadio.value = true;

        /* 反転 / Reverse */
        var reversePanel = leftColumnGroup.add('panel', undefined, getLabel('panel.reverse'));
        setupOptionPanel(reversePanel, 'left', 6);

        var reverseList = addOptionList(reversePanel);
        var reverseSpineCheckbox = addOptionControl(reverseList, 'checkbox', 'checkbox.reverseSpine', 'tooltip.reverseSpine');
        var reverseStackCheckbox = addOptionControl(reverseList, 'checkbox', 'checkbox.reverseStack', 'tooltip.reverseStack');
        reverseSpineCheckbox.value = false;
        reverseStackCheckbox.value = false;

        return {
            orientationPanel: orientationPanel,
            alignToPageRadio: alignToPageRadio,
            alignToPathRadio: alignToPathRadio,
            adjustNoneRadio: adjustNoneRadio,
            adjustReleaseRadio: adjustReleaseRadio,
            adjustExpandRadio: adjustExpandRadio,
            adjustReplaceRadio: adjustReplaceRadio,
            reversePanel: reversePanel,
            reverseSpineCheckbox: reverseSpineCheckbox,
            reverseStackCheckbox: reverseStackCheckbox
        };
    }

    /**
     * ボタン行（キャンセル／OK）を組み立てる
     * @param {Window} blendDialog - 追加先のダイアログ
     * @returns {{btnOK: Button, btnCancel: Button}} OK・キャンセルボタン
     */
    function buildButtonRow(blendDialog) {
        var btnRowGroup = blendDialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignment = 'center';
        var btnCancel = btnRowGroup.add('button', undefined, getLabel('button.cancel'), {
            name: 'cancel'
        });
        var btnOK = btnRowGroup.add('button', undefined, getLabel('button.ok'), {
            name: 'ok'
        });
        return { btnOK: btnOK, btnCancel: btnCancel };
    }

    /**
     * 選択中の方向をブレンドオプションの値で返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {number} BlendOrientation の値
     */
    function readOrientation(dialogControls) {
        return (dialogControls.alignToPathRadio.value) ? BlendOrientation['Align to Path'] : BlendOrientation['Align to Page'];
    }

    /**
     * 選択中の「その他」の調整モードを返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "none" / "release" / "expand" / "replaceSpine"
     */
    function readAdjustMode(dialogControls) {
        if (dialogControls.adjustReleaseRadio.value) return 'release';
        if (dialogControls.adjustExpandRadio.value) return 'expand';
        if (dialogControls.adjustReplaceRadio.value) return 'replaceSpine';
        return 'none';
    }

    /**
     * ステップ数の入力欄とスライダーを連動させ、値が変わるたびにプレビューを更新する
     * @param {EditText} stepInput - ステップ数の入力欄
     * @param {Slider} stepSlider - ステップ数のスライダー
     * @param {function} refreshPreview - プレビューを更新するコールバック
     * @returns {void}
     */
    function bindStepControls(stepInput, stepSlider, refreshPreview) {
        /* 範囲を切り替えるとき（32 ⇔ 128 ⇔ 1000）につまみの位置を保つための控え
           Last slider maximum, used to keep the thumb position when the range switches */
        var lastSliderMax = STEP_SLIDER_MAX.normal;

        /**
         * 修飾キーの状態に応じてステップ数スライダーの上限を切り替える
         *
         * 通常0〜32、Option併用で0〜128、Shift併用で0〜1000。つまみの相対位置は保つ。
         * @returns {void}
         */
        function updateStepSliderRangeFromKeyboard() {
            try {
                var keyboardState = ScriptUI.environment.keyboardState;

                /* 優先順位: Shift (0-1000) > Option (0-128) > 通常 (0-32) / Priority: Shift > Option > default */
                var nextMax = STEP_SLIDER_MAX.normal;
                if (keyboardState && keyboardState.shiftKey) {
                    nextMax = STEP_SLIDER_MAX.shift;
                } else if (keyboardState && keyboardState.altKey) {
                    nextMax = STEP_SLIDER_MAX.option;
                }

                if (lastSliderMax !== nextMax) {
                    /* 範囲を変えてもつまみの相対位置（比率）を保つ / Keep the thumb's relative position */
                    var thumbRatio = (lastSliderMax > 0) ? (stepSlider.value / lastSliderMax) : 0;
                    if (!isFinite(thumbRatio) || isNaN(thumbRatio)) thumbRatio = 0;

                    stepSlider.maxvalue = nextMax;

                    var nextValue = Math.round(thumbRatio * nextMax);
                    if (nextValue < 0) nextValue = 0;
                    if (nextValue > nextMax) nextValue = nextMax;

                    stepSlider.value = nextValue;
                    /* 入力欄も新しい値にそろえる / Keep the field in step with the new value */
                    stepInput.text = String(nextValue);

                    lastSliderMax = nextMax;
                } else if (stepSlider.maxvalue !== nextMax) {
                    /* 上限が他所で変わっていても範囲をそろえる / Re-apply the range if it drifted */
                    stepSlider.maxvalue = nextMax;
                }
            } catch (e) { }
        }

        /**
         * ステップ数を整数に丸め、0以上スライダー上限以下に収める
         * @param {number} stepValue - 丸める前の値
         * @returns {number} 0〜スライダー上限に収めた整数
         */
        function clampToSliderRange(stepValue) {
            var roundedValue = Math.round(stepValue);
            if (roundedValue < 0) roundedValue = 0;
            if (roundedValue > stepSlider.maxvalue) roundedValue = stepSlider.maxvalue;
            return roundedValue;
        }

        /**
         * 入力欄の値をスライダーへ反映する
         * @returns {void}
         */
        function syncSliderFromEditText() {
            var stepValue = parseInt(stepInput.text, 10);
            if (isNaN(stepValue)) return;
            stepSlider.value = clampToSliderRange(stepValue);
        }

        /**
         * スライダーの値を入力欄へ反映する
         * @returns {void}
         */
        function syncEditTextFromSlider() {
            stepInput.text = String(clampToSliderRange(stepSlider.value));
        }

        /* ↑↓キー：スライダーとプレビューを更新 / Arrow keys update the slider and the preview */
        changeValueByArrowKey(stepInput, function () {
            syncSliderFromEditText();
            refreshPreview();
        });

        /* スライダー：入力欄とプレビューを更新 / The slider updates the field and the preview */
        stepSlider.onChanging = stepSlider.onChange = function () {
            updateStepSliderRangeFromKeyboard();
            syncEditTextFromSlider();
            refreshPreview();
        };

        /* 手入力でもプレビュー（数値として読めるときだけ。入力中は書き換えない）
           Preview while typing, only when the text parses; the field is not rewritten per keystroke */
        stepInput.onChanging = function () {
            if (isNaN(parseInt(stepInput.text, 10))) {
                return;
            }
            syncSliderFromEditText();
            refreshPreview();
        };

        /* 確定時（Enter・フォーカス移動）に 0〜上限の整数へ整える / Sanitize on commit */
        stepInput.onChange = function () {
            var stepValue = parseInt(stepInput.text, 10);
            if (isNaN(stepValue)) {
                stepValue = 0;
            }
            stepInput.text = String(clampBlendSteps(stepValue));
            syncSliderFromEditText();
            refreshPreview();
        };
    }

    /**
     * ダイアログのライブプレビューを用意する
     *
     * キャンセルで戻せるよう、プレビューで実行した操作の数を数える。
     * 反転コマンドはトグルなので、前回適用時から状態が変わったときだけ実行する。
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {Object} applyStep / applyReverse / undoAll / keepChanges / wasReverseApplied を持つオブジェクト
     */
    function createBlendPreview(dialogControls) {
        /* プレビューで実行した操作の数（キャンセル時に取り消す）/ Operations to undo on Cancel */
        var previewUndoDepth = 0;

        /* プレビュー対象のブレンド（コマンド実行時に選択されている必要がある）/ Blend the preview acts on */
        var targetBlendItem = null;
        try {
            targetBlendItem = findFirstBlendObject(app.activeDocument.selection);
        } catch (e) { }

        /* 反転の前回適用時の状態と、プレビューで反転を実行したか / Last applied reverse state */
        var lastReverseSpine = dialogControls.reverseSpineCheckbox.value;
        var lastReverseStack = dialogControls.reverseStackCheckbox.value;
        var reversePreviewTouched = false;

        /**
         * 対象のブレンドを特定し、それだけを選択状態にする
         * @returns {void}
         */
        function ensureTargetBlendSelected() {
            try {
                if (!targetBlendItem) {
                    targetBlendItem = findFirstBlendObject(app.activeDocument.selection);
                }
                if (targetBlendItem) {
                    selectOnlyItem(app.activeDocument, targetBlendItem);
                }
            } catch (e) { }
        }

        /**
         * UIの現在値（ステップ数と方向）でブレンド設定を適用し、プレビューを更新する
         * @returns {void}
         */
        function applyStepPreview() {
            var stepValue = parseInt(dialogControls.stepInput.text, 10);
            if (isNaN(stepValue)) {
                return;
            }
            try {
                ensureTargetBlendSelected();
                setBlendOption(clampBlendSteps(stepValue), readOrientation(dialogControls));
                previewUndoDepth++;
                app.redraw();
            } catch (e) { }
        }

        /**
         * 反転チェックボックスの状態をブレンドへ即時プレビューする
         * @returns {void}
         */
        function applyReversePreview() {
            var wantSpine = !!dialogControls.reverseSpineCheckbox.value;
            var wantStack = !!dialogControls.reverseStackCheckbox.value;

            ensureTargetBlendSelected();

            try {
                if (wantSpine !== lastReverseSpine) {
                    app.executeMenuCommand('Path Blend Reverse Spine');
                    previewUndoDepth++;
                    lastReverseSpine = wantSpine;
                    reversePreviewTouched = true;
                }
            } catch (e) { }

            try {
                if (wantStack !== lastReverseStack) {
                    app.executeMenuCommand('Path Blend Reverse Stack');
                    previewUndoDepth++;
                    lastReverseStack = wantStack;
                    reversePreviewTouched = true;
                }
            } catch (e) { }

            app.redraw();
        }

        /**
         * プレビューで実行した操作をすべて取り消す
         * @returns {void}
         */
        function undoAll() {
            try {
                while (previewUndoDepth > 0) {
                    app.executeMenuCommand('undo');
                    previewUndoDepth--;
                }
                app.redraw();
            } catch (e) { }
        }

        return {
            applyStep: applyStepPreview,
            applyReverse: applyReversePreview,
            undoAll: undoAll,
            /* 確定時はプレビューの結果を残す / Keep the preview result on OK */
            keepChanges: function () { previewUndoDepth = 0; },
            wasReverseApplied: function () { return reversePreviewTouched; }
        };
    }

    /**
     * 「その他」「反転」のラジオ・チェックボックスの連動を登録し、初期状態を反映する
     * @param {Window} blendDialog - 対象ダイアログ（淡色表示のペンを作る）
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {PageItem[]} selection - 実行時の選択オブジェクト
     * @param {Object} blendPreview - createBlendPreview() の戻り値
     * @returns {function} 調整モードに合わせて各パネルの状態を更新する関数
     */
    function bindAdjustControls(blendDialog, dialogControls, selection, blendPreview) {
        var adjustRadios = {
            none: dialogControls.adjustNoneRadio,
            release: dialogControls.adjustReleaseRadio,
            expand: dialogControls.adjustExpandRadio,
            replaceSpine: dialogControls.adjustReplaceRadio
        };

        /**
         * 調整モードのラジオボタンを排他的に切り替える
         *
         * ScriptUI は同じ親の中でしか自動排他しないため、パネルをまたぐ排他をここで行う。
         * @param {string} adjustMode - "none" / "release" / "expand" / "replaceSpine" のいずれか
         * @returns {void}
         */
        function selectAdjustMode(adjustMode) {
            adjustRadios.none.value = (adjustMode === 'none');
            adjustRadios.release.value = (adjustMode === 'release');
            adjustRadios.expand.value = (adjustMode === 'expand');
            adjustRadios.replaceSpine.value = (adjustMode === 'replaceSpine');
        }

        var dimPen = null;
        var normalPen = null;

        /**
         * 「その他」が「なし」のとき、他の選択肢のラベルを淡色にする
         *
         * `enabled = false` にすると選択できなくなるため、文字色だけを変える。
         * @returns {void}
         */
        function syncAdjustOptionDimming() {
            try {
                if (!dimPen) {
                    dimPen = blendDialog.graphics.newPen(PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1);
                }
                if (!normalPen) {
                    normalPen = blendDialog.graphics.newPen(PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
                }

                var useDim = !!adjustRadios.none.value;

                /* 無効化中のラジオ（ブレンドのみ選択時の置き換えなど）は常に淡色。有効なものだけ切り替える
                   Disabled radios stay dim; only enabled ones follow the None state */
                var dimmableRadios = [adjustRadios.release, adjustRadios.expand, adjustRadios.replaceSpine];
                for (var i = 0; i < dimmableRadios.length; i++) {
                    var adjustRadio = dimmableRadios[i];
                    if (!adjustRadio) continue;

                    if (adjustRadio.enabled === false) {
                        adjustRadio.graphics.foregroundColor = dimPen;
                        continue;
                    }

                    adjustRadio.graphics.foregroundColor = useDim ? dimPen : normalPen;
                }
            } catch (e) { }
        }

        /**
         * 選択内容に応じて「ブレンド軸を置き換え」の有効／無効を更新する
         *
         * 選択がブレンドのみのときは無効にし、選択済みなら「なし」へ戻す。
         * @returns {void}
         */
        function updateAdjustAvailability() {
            try {
                adjustRadios.replaceSpine.enabled = !isOnlyBlendObjects(selection);
                if (!adjustRadios.replaceSpine.enabled && adjustRadios.replaceSpine.value) {
                    selectAdjustMode('none');
                }
                syncAdjustOptionDimming();
            } catch (e) { }
        }

        /**
         * 「その他」が「なし」以外のとき、方向パネルを無効化し、フォーカスを移す
         * @returns {void}
         */
        function syncBlendPanelEnabled() {
            var isAdjustNone = !!adjustRadios.none.value;
            dialogControls.orientationPanel.enabled = isAdjustNone;

            /* フォーカスを分かりやすい位置へ / Keep focus sensible */
            try {
                if (isAdjustNone) {
                    dialogControls.stepInput.active = true;
                    dialogControls.stepInput.setSelection(0, dialogControls.stepInput.text.length);
                } else {
                    /* 選択中の選択肢にフォーカス / Focus the selected option */
                    if (adjustRadios.release.value) adjustRadios.release.active = true;
                    else if (adjustRadios.expand.value) adjustRadios.expand.active = true;
                    else if (adjustRadios.replaceSpine.value) adjustRadios.replaceSpine.active = true;
                    else adjustRadios.none.active = true;
                }
            } catch (e) { }
        }

        /**
         * 「その他」が「なし」以外のとき、反転パネルを無効化する
         *
         * チェックボックスの値は書き換えない（プレビュー済みの反転は OK／キャンセルの処理まで残す）。
         * @returns {void}
         */
        function syncReversePanelEnabled() {
            dialogControls.reversePanel.enabled = !!adjustRadios.none.value;
        }

        /**
         * 調整モードに合わせて各パネルの状態を更新する
         * @returns {void}
         */
        function refreshAdjustState() {
            updateAdjustAvailability();
            syncBlendPanelEnabled();
            syncReversePanelEnabled();
        }

        /* 「なし」以外を選ぶと方向・反転パネルを無効化 / Choosing an adjustment disables the other panels */
        for (var adjustMode in adjustRadios) {
            if (!adjustRadios.hasOwnProperty(adjustMode)) continue;
            (function (radioMode) {
                adjustRadios[radioMode].onClick = function () {
                    selectAdjustMode(radioMode);
                    refreshAdjustState();
                };
            })(adjustMode);
        }

        /* 反転は独立した操作。混乱を避けるため「その他」は「なし」に戻す
           Reverse actions are independent; keep Misc on None to avoid mixed modes */
        dialogControls.reverseSpineCheckbox.onClick = dialogControls.reverseStackCheckbox.onClick = function () {
            selectAdjustMode('none');
            refreshAdjustState();
            blendPreview.applyReverse();
        };

        refreshAdjustState();
        return refreshAdjustState;
    }

    /**
     * 設定ダイアログを組み立てて表示し、確定した入力値を返す
     * @param {PageItem[]} selection - 実行時の選択オブジェクト
     * @param {boolean} wasBlendSelectedAtOpen - ダイアログを開いた時点でブレンドを選択していたか
     * @returns {Object|null} ステップ数・方向・反転・調整モードを持つ入力値。キャンセル時は null
     */
    function showBlendDialog(selection, wasBlendSelectedAtOpen) {
        var doc = app.activeDocument;

        /* ダイアログ開始時の選択を保存（Replace Spine などで「ブレンド＋パス」の混在選択を維持するため）
           Keep the launch selection; Replace Spine needs the blend and the path selected together */
        var selectionAtOpen = snapshotSelection(doc);
        var defaultStep = readDefaultStep(selection, wasBlendSelectedAtOpen);

        var blendDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        blendDialog.orientation = 'column';
        blendDialog.alignChildren = ['fill', 'top'];
        blendDialog.spacing = DIALOG_SPACING;
        blendDialog.margins = DIALOG_MARGINS;

        var stepControls = buildStepArea(blendDialog, defaultStep);
        var dialogControls = buildOptionColumns(blendDialog);
        dialogControls.stepInput = stepControls.stepInput;
        dialogControls.stepSlider = stepControls.stepSlider;

        var blendPreview = createBlendPreview(dialogControls);
        bindStepControls(dialogControls.stepInput, dialogControls.stepSlider, blendPreview.applyStep);
        dialogControls.alignToPageRadio.onClick = blendPreview.applyStep;
        dialogControls.alignToPathRadio.onClick = blendPreview.applyStep;
        var refreshAdjustState = bindAdjustControls(blendDialog, dialogControls, selection, blendPreview);

        var dialogButtons = buildButtonRow(blendDialog);
        var dialogResult = null;

        blendDialog.onShow = function () {
            blendPreview.applyStep();
            refreshAdjustState();
        };

        dialogButtons.btnOK.onClick = function () {
            var reverseSpine = !!dialogControls.reverseSpineCheckbox.value;
            var reverseStack = !!dialogControls.reverseStackCheckbox.value;

            var stepValue = parseInt(dialogControls.stepInput.text, 10);
            if (isNaN(stepValue) || stepValue < 0 || stepValue > MAX_BLEND_STEPS) {
                alert(getLabel('alert.invalidStep'));
                return;
            }
            var adjustMode = readAdjustMode(dialogControls);

            /* Replace Spine は「ブレンド＋置き換え用パス」の同時選択が必要。
               プレビュー中にブレンド単体選択へ切り替わっていることがあるため、ここで元の選択を復元する。
               Replace Spine needs the blend and the new path selected; restore the launch selection */
            if (adjustMode === 'replaceSpine') {
                restoreSelection(doc, selectionAtOpen);
            }

            /* プレビューで反転済みなら文書は既に目的の状態。true を渡すと main() で再度トグルして打ち消してしまう
               If the reverse preview already ran, passing true would toggle it back in main() */
            var reverseApplied = blendPreview.wasReverseApplied();
            dialogResult = {
                step: stepValue,
                orientation: readOrientation(dialogControls),
                adjustMode: adjustMode,
                reverseSpine: reverseApplied ? false : reverseSpine,
                reverseStack: reverseApplied ? false : reverseStack
            };
            blendPreview.keepChanges();
            blendDialog.close(1);
        };

        dialogButtons.btnCancel.onClick = function () {
            /* このダイアログで適用したプレビューをすべて戻す / Revert every preview change */
            blendPreview.undoAll();
            /* 取り消し時も、ダイアログ開始時の選択に戻す / Restore the launch selection */
            restoreSelection(doc, selectionAtOpen);
            blendDialog.close(0);
        };

        if (blendDialog.show() !== 1) {
            return null;
        }
        return dialogResult;
    }

    /**
     * 入力欄で↑↓キーによる増減を行えるようにする
     *
     * Shift併用で±10。値は 0〜MAX_BLEND_STEPS の整数に収める。
     * @param {EditText} editText - 対象の入力欄
     * @param {function} onValueChanged - 値が変わったときに呼ぶコールバック
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onValueChanged) {
        editText.addEventListener("keydown", function (event) {
            /* 上下キーだけを扱う / Only handle Up/Down */
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = parseInt(editText.text, 10);
            if (isNaN(value)) value = 0;

            /* Shiftキー押下時は10刻み（整数のみ）/ Shift steps by 10 */
            var delta = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
            value += (event.keyName === "Up") ? delta : -delta;
            event.preventDefault();

            /* 整数・0〜上限に収める / Integer within 0..MAX_BLEND_STEPS */
            editText.text = String(clampBlendSteps(Math.round(value)));

            if (typeof onValueChanged === 'function') {
                onValueChanged();
            }
        });
    }

    // =========================================
    // 一時アクション / Temporary action
    // =========================================

    /**
     * アクションを文字列から生成し実行するブロック構文。終了時・エラー発生時の後片付けは自動
     * @param {string} actionCode - アクションのソースコード
     * @param {function} actionCallback - 読み込んだアクションを受け取って実行する処理
     * @returns {void}
     */
    function tempAction(actionCode, actionCallback) {
        /**
         * UTF-8の16進数文字コードを文字列に変換する
         * @param {string} hex - 16進数で表した文字コード列
         * @returns {string} 変換した文字列
         */
        var hexToString = function (hexText) {
            return decodeURIComponent(hexText.replace(/(.{2})/g, '%$1'));
        };

        // ActionItemのconstructor。ActionItem.exec()を使えばわざわざ名前を直接指定しなくても実行できる
        var ActionItem = function ActionItem(index, name, parent) {
            this.index = index;
            this.name = name; // actionName
            this.parent = parent; // setName
        };
        ActionItem.prototype.exec = function(showDialog) {
            doScript(this.name, this.parent, showDialog);
        };

        // ActionItemsのconstructor。
        // ActionItems['actionName'],  ActionItems.getByName('actionName'),
        // ActionItems[0],  ActionItems.index(-1)
        // などの形式で中身のアクションを取得できる
        var ActionItems = function ActionItems() {
            this.length = 0;
        };
        ActionItems.prototype.getByName = function(nameStr) {
            for (var i = 0, len = this.length; i < len; i++) {
                if (this[i].name == nameStr) {
                    return this[i];
                }
            }
        };
        ActionItems.prototype.index = function(keyNumber) {
            var res;
            if (keyNumber >= 0) {
                res = this[keyNumber];
            } else {
                res = this[this.length + keyNumber];
            }
            return res;
        };

        // アクションセット名を取得
        var regExpSetName = /^\/name\s+\[\s+\d+\s+([^\]]+?)\s+\]/m;
        var setName = hexToString(actionCode.match(regExpSetName)[1].replace(/\s+/g, ''));

        // セット内のアクションを取得
        var regExpActionNames = /^\/action-\d+\s+\{\s+\/name\s+\[\s+\d+\s+([^\]]+?)\s+\]/mg;
        var actionItemsObj = new ActionItems();
        var i = 0;
        var matchObj;
        while (matchObj = regExpActionNames.exec(actionCode)) {
            var actionName = hexToString(matchObj[1].replace(/\s+/g, ''));
            var actionObj = new ActionItem(i, actionName, setName);
            actionItemsObj[actionName] = actionObj;
            actionItemsObj[i] = actionObj;
            i++;
            if (i > 1000) {
                break;
            } // limiter
        }
        actionItemsObj.length = i;

        // aiaファイルとして書き出し
        var failed = false;
        var aiaFileObj = new File(Folder.temp + '/tempActionSet.aia');
        try {
            aiaFileObj.open('w');
            aiaFileObj.write(actionCode);
        } catch (e) {
            failed = true;
            alert(e);
            return;
        } finally {
            aiaFileObj.close();
            if (failed) {
                try {
                    aiaFileObj.remove();
                } catch (e) {}
            }
        }

        // 同名アクションセットがあったらunloadする。これは余計なお世話かもしれない
        try {
            app.unloadAction(setName, '');
        } catch (e) {}

        // アクションを読み込み実行する
        var actionLoaded = false;
        try {
            app.loadAction(aiaFileObj);
            actionLoaded = true;
            actionCallback.call(actionCallback, actionItemsObj);
        } catch (e) {
            alert(e);
        } finally {
            // 読み込んだアクションと，そのaiaファイルを削除
            if (actionLoaded) {
                app.unloadAction(setName, '');
            }
            aiaFileObj.remove();
        }
    }

    /**
     * ステップ数と方向を指定して、ブレンドオプションを一時アクションで適用する
     * @param {number} step - 中間ステップ数
     * @param {number} orientationValue - ブレンドの方向を表す値
     * @returns {void}
     */
    function setBlendOption(step, orientationValue) {
        var actionCode = [
            "/version 3",
            "/name [ 5",
            "	426c656e64",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "	/name [ 7",
            "		73657453746570",
            "	]",
            "	/keyIndex 0",
            "	/colorIndex 0",
            "	/isOpen 1",
            "	/eventCount 1",
            "	/event-1 {",
            "		/useRulersIn1stQuadrant 0",
            "		/internalName (ai_plugin_liveblend)",
            "		/localizedName [ 12",
            "			e38396e383ace383b3e38389",
            "		]",
            "		/isOpen 0",
            "		/isOn 1",
            "		/hasDialog 1",
            "		/showDialog 0",
            "		/parameterCount 3",
            "		/parameter-1 {",
            "			/key 1835363957",
            "			/showInPalette 4294967295",
            "			/type (enumerated)",
            "			/name [ 15",
            "				e382aae38397e382b7e383a7e383b3",
            "			]",
            "			/value 5",
            "		}",
            "		/parameter-2 {",
            "			/key 1937007984",
            "			/showInPalette 4294967295",
            "			/type (integer)",
            "			/value " + String(step),
            "		}",
            "		/parameter-3 {",
            "			/key 1919906913",
            "			/showInPalette 4294967295",
            "			/type (enumerated)",
            "			/name [ 12",
            "				e59e82e79bb4e696b9e59091",
            "			]",
            "			/value " + String(orientationValue),
            "		}",
            "	}",
            "}"
        ].join("\n");

        tempAction(actionCode, function(actionItems) {
            actionItems[0].exec(false);
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確認し、ダイアログを開いてブレンドの作成・設定・調整を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length <= 0) {
            return;
        }
        var doc = app.activeDocument;
        var selection = doc.selection;

        /* ダイアログを開いた時点で「ブレンドオブジェクトを選択していたか」を記録
           （この後 Path Blend Make を実行して選択がブレンドに変わることがあるため）
           Remember whether a blend was selected at launch; Path Blend Make may change the selection */
        var wasBlendSelectedAtOpen = !!findFirstBlendObject(selection);

        /* 選択にブレンド（PluginItem/BlendItem）が含まれない場合のみ、ブレンドを作成
           （既に含まれる場合は、パスが混ざっていても選択はそのまま。作成せずオプション/調整のみ行う）
           Make a blend only when none is selected; otherwise keep the selection as is */
        if (!wasBlendSelectedAtOpen) {
            try {
                app.executeMenuCommand('Path Blend Make');
            } catch (e) { }

            /* コマンド実行後に選択が変わる可能性があるため、再取得 / The command may change the selection */
            selection = doc.selection;
            if (!selection || selection.length <= 0) {
                return;
            }

            /* 生成された PluginItem（ブレンド）だけを選択状態にする / Select only the new blend */
            var createdBlend = findFirstBlendObject(selection);
            if (createdBlend) {
                selectOnlyItem(doc, createdBlend);
                selection = doc.selection;
            }
        }

        /* ダイアログを表示してユーザー入力を取得（UI操作専用。引数実行は行わない）
           Show the dialog; this script is UI-only */
        var dialogInput = showBlendDialog(selection, wasBlendSelectedAtOpen);
        if (!dialogInput) {
            return; /* キャンセル / cancelled */
        }

        var adjustMode = dialogInput.adjustMode || 'none';
        if (adjustMode === 'release') {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Release');
            return;
        }
        if (adjustMode === 'expand') {
            app.executeMenuCommand('Path Blend Expand');
            return;
        }
        if (adjustMode === 'replaceSpine') {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Replace Spine');
            return;
        }

        var reverseSpine = !!dialogInput.reverseSpine;
        var reverseStack = !!dialogInput.reverseStack;
        if (reverseSpine) {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Reverse Spine');
        }
        if (reverseStack) {
            runBlendCommandAndDeleteFrontmostPath(doc, 'Path Blend Reverse Stack');
        }
        if (reverseSpine || reverseStack) {
            return;
        }

        /* オプションだけを適用（新しいブレンドは作らない）/ Apply options only; no new blend */
        setBlendOption(dialogInput.step, dialogInput.orientation);
    }

    main();

})();
