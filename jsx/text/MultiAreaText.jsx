#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のテキストを1つのエリア内文字にまとめたり、逆に分割したりします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MultiAreaText.md

### Overview

Merges several text objects into a single area text, or splits one back out.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MultiAreaText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MultiAreaText";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-03-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MultiAreaText.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MultiAreaText.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];      /* パネルの余白 / panel margins */
    var SPACING_FIELD_CHARACTERS = 4;          /* 外枠からの間隔の入力欄の文字数 / width of the inset spacing field */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "複数のテキスト", en: "Multiple Text Objects" }
        },
        panel: {
            operation: { ja: "操作", en: "Action" },
            order: { ja: "順序", en: "Order" },
            threadText: { ja: "スレッドテキスト", en: "Threaded Text" },
            mergeSettings: { ja: "マージの設定", en: "Merge Settings" },
            areaTextHeight: { ja: "エリア内文字の高さ", en: "Area Text Height" }
        },
        radio: {
            merge: { ja: "マージ", en: "Merge" },
            thread: { ja: "スレッドテキスト", en: "Threaded Text" },
            swap: { ja: "交換（文字列のみ）", en: "Swap (Text Only)" },
            topToBottom: { ja: "上から", en: "Top to Bottom" },
            leftToRight: { ja: "左から", en: "Left to Right" },
            threadLink: { ja: "作成", en: "Create" },
            threadUnlink: { ja: "スレッドのリンクを解除", en: "Remove Threading" },
            threadAdd: { ja: "スレッドに追加", en: "Add to Thread" },
            threadRelease: { ja: "スレッドから除外", en: "Release from Thread" },
            releaseByDuplicate: { ja: "複製してスレッドから除外", en: "Duplicate and Release" },
            heightNone: { ja: "何もしない", en: "None" },
            heightFit: { ja: "フィット", en: "Fit" },
            heightAuto: { ja: "自動サイズ調整", en: "Auto Size" }
        },
        checkbox: {
            removeLineBreaks: { ja: "改行を削除", en: "Remove Line Breaks" },
            preserveFormatting: { ja: "書式を保持（段落単位）", en: "Preserve Formatting (Per Paragraph)" },
            insetSpacing: { ja: "外枠からの間隔", en: "Inset Spacing" },
            justify: { ja: "均等配置（最終行左揃え）", en: "Justify (Last Line Left)" },
            addBorder: { ja: "長方形の枠を付ける", en: "Add Rectangle Border" }
        },
        tooltip: {
            merge: {
                ja: "選択したテキストを連結し、全体を囲む1つのエリア内文字にまとめます",
                en: "Joins the selected text into one area text that covers them all."
            },
            thread: {
                ja: "［スレッドテキスト］パネルで選んだ操作を実行します",
                en: "Runs the action chosen in the Threaded Text panel."
            },
            swap: {
                ja: "2つのテキストの文字列を入れ替えます。位置と書式はそのままです",
                en: "Swaps the text of the two objects; their position and formatting stay put."
            },
            topToBottom: { ja: "上にあるテキストから順に連結します", en: "Joins the text from top to bottom." },
            leftToRight: { ja: "左にあるテキストから順に連結します", en: "Joins the text from left to right." },
            threadLink: {
                ja: "選択したエリア内文字をスレッドでつなぎます。すでにつながっているものがあれば、いったん解除してからつなぎ直します",
                en: "Threads the selected area text, removing any existing threading first."
            },
            threadUnlink: {
                ja: "［スレッドのリンクを解除］を実行します",
                en: "Runs Remove Threading."
            },
            threadAdd: {
                ja: "スレッドをいったん解除し、選択したエリア内文字をまとめてつなぎ直します",
                en: "Removes the threading, then threads the selected area text again."
            },
            threadRelease: {
                ja: "選択したエリア内文字をスレッドから外し、中の文字を同じ位置に独立したエリア内文字として残します（文字属性は段落ごと）",
                en: "Releases the selected area text and keeps its text as separate area text in place (character attributes per paragraph)."
            },
            releaseByDuplicate: {
                ja: "選択したエリア内文字を複製して独立させ、元のエリア内文字とその文字をスレッドから削除します",
                en: "Duplicates each selected area text as a standalone copy, then deletes the original and its text from the thread."
            },
            removeLineBreaks: {
                ja: "連結したテキストの改行をすべて削除します（［書式を保持］がオンのときは使えません）",
                en: "Removes every line break from the joined text (unavailable while Preserve Formatting is on)."
            },
            preserveFormatting: {
                ja: "段落ごとに文字属性を復元します（段落設定は完全ではありません）",
                en: "Restores character attributes paragraph by paragraph (paragraph settings are not fully kept)."
            },
            insetSpacing: {
                ja: "まとめたエリア内文字の［外枠からの間隔］を、定規の単位で設定します",
                en: "Sets the inset spacing of the merged area text, in ruler units."
            },
            addBorder: {
                ja: "新しい線を追加し、［形状に変換］の長方形効果でテキストを枠で囲みます",
                en: "Adds a new stroke and a Convert to Rectangle effect to frame the text."
            },
            heightFit: {
                ja: "テキストに合わせて高さを一度だけ調整します（自動サイズ調整はオフに戻します）",
                en: "Fits the height to the text once, then turns Auto Size off again."
            },
            heightAuto: {
                ja: "エリア内文字の自動サイズ調整をオンにします",
                en: "Turns on Auto Size for the area text."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            noTextFrame: { ja: "選択にテキストが含まれていません。", en: "Selection does not contain any text." },
            notThreaded: { ja: "選択されたテキストはスレッドテキストではありません。", en: "The selected text is not threaded text." },
            needTwo: { ja: "テキストを2つ以上選択してください。", en: "Please select 2 or more text objects." },
            releaseFailed: {
                ja: "スレッドから除外できなかったため、内容は変更していません。",
                en: "Could not release the text from the thread, so nothing was changed."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 自動サイズ調整（一時アクション） / Auto size via a temporary action
    // =========================================

    /**
     * 一時アクションで、選択中のエリア内文字の自動サイズ調整を切り替える
     * @param {number} autoSizeValue - 1 = オン、2 = オフ
     * @returns {void}
     */
    function runAutoSizeAction(autoSizeValue) {
        var actionSource = '/version 3' + '/name [ 8' + ' 4172656154797065' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 8' + ' 4175746f53697a65' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_SLOAreaTextDialog)' + ' /localizedName [ 33' + ' e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1952539754' + ' /showInPalette 4294967295' + ' /type (integer)' + ' /value ' + autoSizeValue + ' }' + ' }' + '}';

        var actionFile = new File('~/ScriptAction.aia');
        var actionLoaded = false;
        /* 書き出し・読み込み・実行のどこで失敗しても、一時ファイルとアクションは片付ける
           Clean up the temp file and the action whichever step fails */
        try {
            actionFile.open('w');
            actionFile.write(actionSource);
            actionFile.close();
            app.loadAction(actionFile);
            actionLoaded = true;
            app.doScript("AutoSize", "AreaType", false);
        } finally {
            try { actionFile.remove(); } catch (e) { }
            if (actionLoaded) {
                try { app.unloadAction("AreaType", ""); } catch (e) { }
            }
        }
    }

    /**
     * エリア内文字を選択し、自動サイズ調整をオン／オフする
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @param {boolean} autoSizeOn - オンにするなら true
     * @returns {void}
     */
    function setAreaTextAutoSize(textFrame, autoSizeOn) {
        app.activeDocument.selection = [textFrame];
        runAutoSizeAction(autoSizeOn ? 1 : 2);
    }

    // =========================================
    // 形状に変換（ライブ効果） / Convert to Shape live effect
    // =========================================

    /* ［形状に変換］の長方形：相対指定で幅・高さに 10pt 追加（CornerRadius は角丸長方形用で、長方形では使われない）
       Convert to Rectangle, relative: 10 pt extra width and height (CornerRadius only matters for rounded rectangles) */
    var RECTANGLE_SHAPE_EFFECT_XML = '<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 ' +
        'R RelWidth 10 R RelHeight 10 R AbsWidth 10 R AbsHeight 10 R Absolute 0 R CornerRadius 9 "/></LiveEffect>';

    /**
     * ［形状に変換］の長方形効果を適用する。失敗したらメッセージを出す
     * @param {PageItem} targetItem - 適用先
     * @returns {void}
     */
    function applyRectangleShapeEffect(targetItem) {
        try {
            targetItem.applyEffect(RECTANGLE_SHAPE_EFFECT_XML);
        } catch (error) {
            alert(error.message);
        }
    }

    // =========================================
    // 選択と文字属性 / Selection and character attributes
    // =========================================

    /**
     * テキストフレームがスレッドでつながっているかを返す
     * @param {TextFrame} textFrame - 判定するフレーム
     * @returns {boolean} 前後どちらかにフレームがつながっていれば true
     */
    function isThreadedFrame(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return false;
        /* つながりが無いときや無効なフレームで例外になることがある / may throw when unthreaded or invalid */
        try { if (textFrame.nextFrame) return true; } catch (e) { }
        try { if (textFrame.previousFrame) return true; } catch (e) { }
        return false;
    }

    /**
     * グループの中から、ロック・非表示でないテキストフレームを再帰的に集める
     * クリップグループのクリッピングパス（pageItems[0]）は対象外
     * @param {GroupItem} groupItem - 対象のグループ
     * @param {Array} collectedItems - 見つけたテキストフレームを追加する配列
     * @returns {void}
     */
    function collectTextFramesInGroup(groupItem, collectedItems) {
        var startIndex = groupItem.clipped ? 1 : 0;
        for (var i = startIndex; i < groupItem.pageItems.length; i++) {
            var childItem = groupItem.pageItems[i];
            if (childItem.locked || childItem.hidden) continue;
            if (childItem.typename === "GroupItem") {
                collectTextFramesInGroup(childItem, collectedItems);
            } else if (childItem.typename === "TextFrame") {
                collectedItems.push(childItem);
            }
        }
    }

    /**
     * 選択を処理対象の並びに展開する。グループは中のテキストフレームに置き換え、それ以外はそのまま残す
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {Array} 展開した処理対象
     */
    function expandGroupsInSelection(selectedItems) {
        var expandedItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem") {
                collectTextFramesInGroup(selectedItems[i], expandedItems);
            } else {
                expandedItems.push(selectedItems[i]);
            }
        }
        return expandedItems;
    }

    /**
     * 選択にテキストフレームが1つでも含まれるかを返す
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 含まれていれば true
     */
    function containsTextFrame(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame") return true;
        }
        return false;
    }

    /**
     * 選択にポイント文字またはパス上文字が含まれるかを返す
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 含まれていれば true
     */
    function containsNonAreaText(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame" &&
                (selectedItems[i].kind === TextType.POINTTEXT || selectedItems[i].kind === TextType.PATHTEXT)) {
                return true;
            }
        }
        return false;
    }

    /* 引き継ぐ文字属性（読み書きの順） / Character attributes carried over, in read/write order */
    var COPIED_CHARACTER_ATTRIBUTES = ["textFont", "size", "leading", "fillColor", "tracking", "kerningMethod"];

    /**
     * 文字属性のうち引き継ぐものを控える
     * @param {CharacterAttributes} sourceAttributes - 読み取り元の文字属性
     * @returns {Object} 属性名をキーにした控え
     */
    function readCharacterAttributes(sourceAttributes) {
        var attributeSnapshot = {};
        for (var attrIndex = 0; attrIndex < COPIED_CHARACTER_ATTRIBUTES.length; attrIndex++) {
            attributeSnapshot[COPIED_CHARACTER_ATTRIBUTES[attrIndex]] = sourceAttributes[COPIED_CHARACTER_ATTRIBUTES[attrIndex]];
        }
        return attributeSnapshot;
    }

    /**
     * 控えた文字属性を書き込む
     * @param {CharacterAttributes} targetAttributes - 書き込み先の文字属性
     * @param {Object} attributeSnapshot - readCharacterAttributes() の結果
     * @returns {void}
     */
    function writeCharacterAttributes(targetAttributes, attributeSnapshot) {
        for (var attrIndex = 0; attrIndex < COPIED_CHARACTER_ATTRIBUTES.length; attrIndex++) {
            targetAttributes[COPIED_CHARACTER_ATTRIBUTES[attrIndex]] = attributeSnapshot[COPIED_CHARACTER_ATTRIBUTES[attrIndex]];
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 縦並びの列グループを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加した列グループ
     */
    function addColumnGroup(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        return columnGroup;
    }

    /**
     * 共通レイアウトのパネルを追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} titlePath - パネル見出しのラベルパス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentGroup, titlePath) {
        var optionPanel = parentGroup.add("panel", undefined, getLabel(titlePath));
        optionPanel.orientation = "column";
        optionPanel.alignment = ["fill", "top"];
        optionPanel.alignChildren = ["left", "center"];
        optionPanel.margins = PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * ツールチップを付けてコントロールを追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} controlType - "radiobutton" / "checkbox"
     * @param {string} labelPath - ラベルのパス
     * @param {string} [tooltipPath] - ツールチップのラベルパス
     * @returns {Object} 追加したコントロール
     */
    function addControl(parentContainer, controlType, labelPath, tooltipPath) {
        var addedControl = parentContainer.add(controlType, undefined, getLabel(labelPath));
        if (tooltipPath) addedControl.helpTip = getLabel(tooltipPath);
        return addedControl;
    }

    /**
     * 左カラム（操作・順序・スレッドテキスト）を組み立てる
     * @param {Group} leftColumn - 追加先の列グループ
     * @param {Object} dialogControls - コントロールを登録するオブジェクト
     * @param {boolean} isSingle - テキストが1つだけなら true
     * @returns {void}
     */
    function addOperationColumn(leftColumn, dialogControls, isSingle) {
        dialogControls.operationPanel = addOptionPanel(leftColumn, "panel.operation");
        dialogControls.rbMerge = addControl(dialogControls.operationPanel, "radiobutton", "radio.merge", "tooltip.merge");
        dialogControls.rbThread = addControl(dialogControls.operationPanel, "radiobutton", "radio.thread", "tooltip.thread");
        dialogControls.rbSwap = addControl(dialogControls.operationPanel, "radiobutton", "radio.swap", "tooltip.swap");
        dialogControls.rbMerge.value = !isSingle;
        dialogControls.rbThread.value = isSingle;

        dialogControls.orderPanel = addOptionPanel(leftColumn, "panel.order");
        dialogControls.rbTopToBottom = addControl(dialogControls.orderPanel, "radiobutton", "radio.topToBottom", "tooltip.topToBottom");
        dialogControls.rbLeftToRight = addControl(dialogControls.orderPanel, "radiobutton", "radio.leftToRight", "tooltip.leftToRight");
        dialogControls.rbTopToBottom.value = true;

        dialogControls.threadPanel = addOptionPanel(leftColumn, "panel.threadText");
        dialogControls.rbThreadLink = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadLink", "tooltip.threadLink");
        dialogControls.rbThreadUnlink = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadUnlink", "tooltip.threadUnlink");
        dialogControls.rbThreadAdd = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadAdd", "tooltip.threadAdd");
        dialogControls.rbThreadRelease = addControl(dialogControls.threadPanel, "radiobutton", "radio.threadRelease", "tooltip.threadRelease");
        dialogControls.rbReleaseByDuplicate = addControl(dialogControls.threadPanel, "radiobutton", "radio.releaseByDuplicate", "tooltip.releaseByDuplicate");
        dialogControls.rbThreadRelease.value = isSingle;
        dialogControls.rbThreadLink.value = !isSingle;
    }

    /**
     * ↑↓キーで数値欄の値を増減する（shift で10刻み、option で0.1刻み）
     * @param {EditText} editText - 対象の数値欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != 'Up' && event.keyName != 'Down') return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = (event.keyName == 'Up') ? 1 : -1;

            if (keyboard.shiftKey) {
                /* 10刻み（10の倍数にスナップ）/ Step by 10, snapping to multiples of 10 */
                value = (delta > 0) ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
                if (value < 0) value = 0;
            } else if (keyboard.altKey) {
                /* 0.1刻み（小数第1位に丸め）/ Step by 0.1, rounded to one decimal */
                value = Math.round((value + delta * 0.1) * 10) / 10;
            } else {
                value = Math.round(value + delta);
                if (value < 0) value = 0;
            }

            editText.text = value;
            event.preventDefault();
        });
    }

    /**
     * 右カラム（マージの設定・エリア内文字の高さ）を組み立てる
     * @param {Group} rightColumn - 追加先の列グループ
     * @param {Object} dialogControls - コントロールを登録するオブジェクト
     * @param {string} rulerLabel - 定規の単位の表示名
     * @returns {void}
     */
    function addMergeSettingsColumn(rightColumn, dialogControls, rulerLabel) {
        var settingsPanel = addOptionPanel(rightColumn, "panel.mergeSettings");
        dialogControls.mergeSettingsPanel = settingsPanel;
        dialogControls.cbRemoveLineBreaks = addControl(settingsPanel, "checkbox", "checkbox.removeLineBreaks", "tooltip.removeLineBreaks");
        dialogControls.cbPreserveFormatting = addControl(settingsPanel, "checkbox", "checkbox.preserveFormatting", "tooltip.preserveFormatting");
        dialogControls.cbPreserveFormatting.value = true;
        dialogControls.cbRemoveLineBreaks.enabled = false;
        dialogControls.cbPreserveFormatting.onClick = function () {
            dialogControls.cbRemoveLineBreaks.enabled = !dialogControls.cbPreserveFormatting.value;
        };

        var insetSpacingGroup = settingsPanel.add("group");
        dialogControls.cbInsetSpacing = addControl(insetSpacingGroup, "checkbox", "checkbox.insetSpacing", "tooltip.insetSpacing");
        dialogControls.insetSpacingInput = insetSpacingGroup.add("edittext", undefined, "1");
        dialogControls.insetSpacingInput.characters = SPACING_FIELD_CHARACTERS;
        dialogControls.insetSpacingInput.enabled = false;
        changeValueByArrowKey(dialogControls.insetSpacingInput);
        insetSpacingGroup.add("statictext", undefined, rulerLabel);
        dialogControls.cbInsetSpacing.onClick = function () {
            dialogControls.insetSpacingInput.enabled = dialogControls.cbInsetSpacing.value;
        };

        dialogControls.cbJustify = addControl(settingsPanel, "checkbox", "checkbox.justify");
        dialogControls.cbJustify.value = true;
        dialogControls.cbAddBorder = addControl(settingsPanel, "checkbox", "checkbox.addBorder", "tooltip.addBorder");

        dialogControls.heightPanel = addOptionPanel(rightColumn, "panel.areaTextHeight");
        dialogControls.rbHeightNone = addControl(dialogControls.heightPanel, "radiobutton", "radio.heightNone");
        dialogControls.rbHeightFit = addControl(dialogControls.heightPanel, "radiobutton", "radio.heightFit", "tooltip.heightFit");
        dialogControls.rbHeightAuto = addControl(dialogControls.heightPanel, "radiobutton", "radio.heightAuto", "tooltip.heightAuto");
        dialogControls.rbHeightFit.value = true;
    }

    /**
     * 選んだ操作に合わせて、パネルとラジオボタンの有効／無効を切り替える
     * 子の有効／無効は各チェックボックスの onClick が保つので、ここではパネル単位で切り替える
     * @param {Object} dialogControls - buildDialog() のコントロール
     * @param {Array} selectedItems - 処理対象
     * @returns {void}
     */
    function updatePanelStates(dialogControls, selectedItems) {
        var isMerge = dialogControls.rbMerge.value;
        /* 1つのときはマージ不可、交換は2つのときだけ / Merge needs two or more, swap exactly two */
        dialogControls.rbMerge.enabled = (selectedItems.length > 1);
        dialogControls.rbSwap.enabled = (selectedItems.length === 2);
        /* ポイント文字・パス上文字はスレッドにできない / Point and path text cannot be threaded */
        dialogControls.rbThread.enabled = !containsNonAreaText(selectedItems);
        dialogControls.orderPanel.enabled = isMerge;
        dialogControls.threadPanel.enabled = dialogControls.rbThread.value;
        dialogControls.mergeSettingsPanel.enabled = isMerge;
        dialogControls.heightPanel.enabled = isMerge;
    }

    /**
     * ダイアログを組み立てる（表示はしない）
     * @param {Array} selectedItems - 処理対象
     * @param {string} rulerLabel - 定規の単位の表示名
     * @param {boolean} heightOnly - ［エリア内文字の高さ］だけを有効にするなら true
     * @returns {Object} ダイアログ本体（window）と各コントロール
     */
    function buildDialog(selectedItems, rulerLabel, heightOnly) {
        var dialogControls = {};

        dialogControls.window = new Window("dialog", getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        var columnsGroup = dialogControls.window.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];

        addOperationColumn(addColumnGroup(columnsGroup), dialogControls, selectedItems.length === 1);
        addMergeSettingsColumn(addColumnGroup(columnsGroup), dialogControls, rulerLabel);

        if (heightOnly) {
            /* スレッドでないエリア内文字1つ：エリア内文字の高さ以外を無効化 / A single unthreaded area text: disable everything but Area Text Height */
            dialogControls.operationPanel.enabled = false;
            dialogControls.orderPanel.enabled = false;
            dialogControls.threadPanel.enabled = false;
            dialogControls.mergeSettingsPanel.enabled = false;
        } else {
            var onOperationClick = function () { updatePanelStates(dialogControls, selectedItems); };
            dialogControls.rbMerge.onClick = onOperationClick;
            dialogControls.rbThread.onClick = onOperationClick;
            dialogControls.rbSwap.onClick = onOperationClick;
            onOperationClick();
        }

        var btnRowGroup = dialogControls.window.add("group");
        btnRowGroup.alignment = ["right", "center"];
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return dialogControls;
    }

    /**
     * ダイアログのコントロールから処理設定を読み取る
     * @param {Object} dialogControls - buildDialog() の結果
     * @returns {Object} 処理設定
     */
    function readDialogSettings(dialogControls) {
        var threadAction = "link";
        if (dialogControls.rbThreadUnlink.value) threadAction = "unlink";
        else if (dialogControls.rbThreadRelease.value) threadAction = "release";
        else if (dialogControls.rbReleaseByDuplicate.value) threadAction = "releaseByDuplicate";
        else if (dialogControls.rbThreadAdd.value) threadAction = "add";

        var frameHeightMode = "none";
        if (dialogControls.rbHeightFit.value) frameHeightMode = "fit";
        else if (dialogControls.rbHeightAuto.value) frameHeightMode = "auto";

        return {
            operation: dialogControls.rbThread.value ? "thread" : (dialogControls.rbSwap.value ? "swap" : "merge"),
            threadAction: threadAction,
            leftToRight: dialogControls.rbLeftToRight.value,
            removeLineBreaks: dialogControls.cbRemoveLineBreaks.value,
            preserveFormatting: dialogControls.cbPreserveFormatting.value,
            insetSpacingOn: dialogControls.cbInsetSpacing.value,
            insetSpacingText: dialogControls.insetSpacingInput.text,
            justify: dialogControls.cbJustify.value,
            addBorder: dialogControls.cbAddBorder.value,
            frameHeightMode: frameHeightMode
        };
    }

    // =========================================
    // スレッドテキスト / Threaded text
    // =========================================

    /**
     * 選択したフレームをスレッドでつなぐ（つながっているものがあれば先に解除する）
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function linkThreadFrames(selectedItems) {
        /* スレッドテキストが混在しているか判定 / Check if selection has threaded frames */
        var hasThreaded = false;
        for (var i = 0; i < selectedItems.length; i++) {
            if (isThreadedFrame(selectedItems[i])) {
                hasThreaded = true;
                break;
            }
        }
        if (hasThreaded) {
            app.executeMenuCommand('removeThreading');
        }
        app.executeMenuCommand('threadTextCreate');
    }

    /**
     * スレッドから外す前に、フレームの位置・文字列・段落ごとの文字属性を控える
     * @param {TextFrame} textFrame - 対象のフレーム
     * @returns {Object} 控え（frame / left / top / width / height / text / paragraphAttributes）
     */
    function snapshotFrameForRelease(textFrame) {
        var frameBounds = textFrame.geometricBounds;
        var frameText = textFrame.contents;

        var paragraphAttributes = [];
        for (var paragraphIndex = 0; paragraphIndex < textFrame.paragraphs.length; paragraphIndex++) {
            /* 読めない段落（末尾の空段落など）で打ち切る / stop at a paragraph that cannot be read, such as a trailing empty one */
            try {
                paragraphAttributes.push(readCharacterAttributes(textFrame.paragraphs[paragraphIndex].characterAttributes));
            } catch (e) {
                break;
            }
        }

        return {
            frame: textFrame,
            left: frameBounds[0],
            top: frameBounds[1],
            width: frameBounds[2] - frameBounds[0],
            height: frameBounds[1] - frameBounds[3],
            text: frameText,
            paragraphAttributes: paragraphAttributes
        };
    }

    /**
     * 控えた内容から、同じ位置・大きさのエリア内文字を作り直す
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} frameSnapshot - snapshotFrameForRelease() の結果
     * @returns {void}
     */
    function recreateAreaText(doc, frameSnapshot) {
        var newRect = doc.pathItems.rectangle(frameSnapshot.top, frameSnapshot.left, frameSnapshot.width, frameSnapshot.height);
        var newTextFrame = doc.textFrames.areaText(newRect);
        newTextFrame.contents = frameSnapshot.text;

        var paragraphAttributes = frameSnapshot.paragraphAttributes;
        /* まず先頭段落の書式をtextRange全体に適用 / Apply first para attrs to entire textRange */
        if (paragraphAttributes.length > 0) {
            writeCharacterAttributes(newTextFrame.textRange.characterAttributes, paragraphAttributes[0]);
        }

        /* 段落ごとに書式を適用。段落が足りなければ打ち切る / Apply attributes per paragraph; stop when paragraphs run out */
        for (var paragraphIndex = 1; paragraphIndex < paragraphAttributes.length; paragraphIndex++) {
            try {
                writeCharacterAttributes(newTextFrame.paragraphs[paragraphIndex].characterAttributes, paragraphAttributes[paragraphIndex]);
            } catch (e) {
                break;
            }
        }
    }

    /**
     * 選択を控える（あとで元に戻すため）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択されていたオブジェクト
     */
    function copySelection(doc) {
        var savedSelection = [];
        for (var i = 0; i < doc.selection.length; i++) savedSelection.push(doc.selection[i]);
        return savedSelection;
    }

    /**
     * 控えた選択を戻す。削除済みのオブジェクトは飛ばす
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} savedSelection - copySelection() の結果
     * @returns {void}
     */
    function restoreSelection(doc, savedSelection) {
        var itemsToSelect = [];
        for (var i = 0; i < savedSelection.length; i++) {
            /* 削除済みのオブジェクトは typename の参照で例外になる / a removed item throws on typename */
            try {
                if (savedSelection[i] && savedSelection[i].typename) itemsToSelect.push(savedSelection[i]);
            } catch (e) { }
        }
        if (itemsToSelect.length === 0) return;
        try {
            doc.selection = itemsToSelect;
        } catch (e) { }
    }

    /**
     * 選択したフレームをスレッドから外し、中の文字を同じ位置の独立したエリア内文字として残す
     * 先に contents を消すと、コマンド失敗時にテキストが失われるため、release の成功を確かめてから消す
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function releaseFromThread(doc, selectedItems) {
        var frameSnapshots = [];
        var releaseTargets = [];

        /* 選択を控えて、最後に戻す / Save the selection to restore it at the end */
        var savedSelection = copySelection(doc);

        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename !== "TextFrame") continue;
            frameSnapshots.push(snapshotFrameForRelease(selectedItems[i]));
            releaseTargets.push(selectedItems[i]);
        }

        /* 先に release を実行（まだ内容は変えない） / Execute release first, before any destructive change */
        var releaseOk = false;
        try {
            doc.selection = releaseTargets;
            app.executeMenuCommand('releaseThreadedTextSelection');

            /* すべての対象がスレッドから外れていれば成功 / Success when no target is threaded any more */
            releaseOk = true;
            for (var checkIndex = 0; checkIndex < frameSnapshots.length; checkIndex++) {
                if (isThreadedFrame(frameSnapshots[checkIndex].frame)) {
                    releaseOk = false;
                    break;
                }
            }
        } catch (e) {
            releaseOk = false;
        }

        if (!releaseOk) {
            alert(getLabel("alert.releaseFailed"));
            return;
        }

        /* スレッドから外れたので、元の内容を消しても連結先に影響しない / Safe to clear now that the chain is broken */
        for (var clearIndex = 0; clearIndex < frameSnapshots.length; clearIndex++) {
            frameSnapshots[clearIndex].frame.contents = "";
        }

        /* 保存した情報から新しいエリア内文字を作成 / Create new area text from saved info */
        for (var createIndex = 0; createIndex < frameSnapshots.length; createIndex++) {
            recreateAreaText(doc, frameSnapshots[createIndex]);
        }

        /* 空になり、スレッドから外れた元のフレームだけを削除 / Remove only the originals that are empty and unthreaded */
        for (var removeIndex = 0; removeIndex < frameSnapshots.length; removeIndex++) {
            /* 無効になったフレームは読み取りや削除で例外になる / an invalid frame throws on read or remove */
            try {
                var originalFrame = frameSnapshots[removeIndex].frame;
                if (!originalFrame || originalFrame.typename !== "TextFrame") continue;
                if (originalFrame.contents !== "") continue;
                if (isThreadedFrame(originalFrame)) continue;
                originalFrame.remove();
            } catch (e) { }
        }

        restoreSelection(doc, savedSelection);
    }

    /**
     * 選択したフレームを複製して独立させ、元のフレームを削除する
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function releaseByDuplicate(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename !== "TextFrame") continue;
            var originalFrame = selectedItems[i];
            /* 複製（スレッドに属さない独立コピーが作られる） / The duplicate is a standalone copy */
            var duplicatedFrame = originalFrame.duplicate();
            /* 複製がスレッドから独立しているか簡易検証 / Make sure the duplicate is not threaded */
            if (isThreadedFrame(duplicatedFrame)) {
                duplicatedFrame.remove();
                continue;
            }
            /* 元フレームのテキストをスレッドから削除し、元フレームも削除 / Clear the original text from the thread, then remove the frame */
            originalFrame.contents = "";
            originalFrame.remove();
        }
    }

    /**
     * ［スレッドテキスト］パネルで選んだ操作を実行する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {string} threadAction - "link" / "unlink" / "add" / "release" / "releaseByDuplicate"
     * @returns {void}
     */
    function runThreadAction(doc, selectedItems, threadAction) {
        /* メニューコマンドは選択に効くので、グループを展開した並びを選び直す / Menu commands act on the selection, so reselect the expanded items */
        doc.selection = selectedItems;
        if (threadAction === "unlink") {
            app.executeMenuCommand('removeThreading');
        } else if (threadAction === "release") {
            releaseFromThread(doc, selectedItems);
        } else if (threadAction === "releaseByDuplicate") {
            releaseByDuplicate(selectedItems);
        } else if (threadAction === "add") {
            app.executeMenuCommand('removeThreading');
            app.executeMenuCommand('threadTextCreate');
        } else {
            linkThreadFrames(selectedItems);
        }
    }

    // =========================================
    // マージ / Merge
    // =========================================

    /**
     * マージの対象になるテキストフレームを集め、指定の順に並べる
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {boolean} leftToRight - 左から並べるなら true（false なら上から）
     * @returns {TextFrame[]} 並べ替えたテキストフレーム
     */
    function collectMergeTargets(selectedItems, leftToRight) {
        var textFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame") textFrames.push(selectedItems[i]);
        }

        /* ソート（左から右、または上から下） / Sort (left-to-right or top-to-bottom) */
        if (leftToRight) {
            textFrames.sort(function (firstFrame, secondFrame) {
                return firstFrame.geometricBounds[0] - secondFrame.geometricBounds[0];
            });
        } else {
            textFrames.sort(function (firstFrame, secondFrame) {
                return secondFrame.geometricBounds[1] - firstFrame.geometricBounds[1];
            });
        }
        return textFrames;
    }

    /**
     * 複数のフレーム全体を囲む矩形を求める
     * @param {TextFrame[]} textFrames - 対象のフレーム
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getCombinedBounds(textFrames) {
        var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        for (var i = 0; i < textFrames.length; i++) {
            var frameBounds = textFrames[i].geometricBounds;
            minX = Math.min(minX, frameBounds[0]);
            maxY = Math.max(maxY, frameBounds[1]);
            maxX = Math.max(maxX, frameBounds[2]);
            minY = Math.min(minY, frameBounds[3]);
        }
        return [minX, maxY, maxX, minY];
    }

    /**
     * まとめたエリア内文字に、元のフレームごとの文字属性を段落単位で書き戻す
     * @param {TextFrame} newTextFrame - まとめたエリア内文字
     * @param {Object[]} frameAttributes - フレームごとの文字属性の控え
     * @param {number[]} frameParagraphCounts - フレームごとの段落数
     * @returns {void}
     */
    function applyFormattingPerFrame(newTextFrame, frameAttributes, frameParagraphCounts) {
        /* 先頭フレームの書式をtextRange全体に適用後、段落ごとに上書き / Apply the first frame to the whole range, then override per paragraph */
        writeCharacterAttributes(newTextFrame.textRange.characterAttributes, frameAttributes[0]);
        var paragraphIndex = 0;
        for (var frameIndex = 0; frameIndex < frameAttributes.length; frameIndex++) {
            for (var framePara = 0; framePara < frameParagraphCounts[frameIndex]; framePara++) {
                /* 末尾の空段落は Error 1302 になるのでスキップ / skip the trailing empty paragraph (Error 1302) */
                try {
                    writeCharacterAttributes(newTextFrame.paragraphs[paragraphIndex].characterAttributes, frameAttributes[frameIndex]);
                } catch (e) { }
                paragraphIndex++;
            }
        }
    }

    /**
     * 行揃えと外枠からの間隔を適用する
     * @param {TextFrame} textFrame - 対象のフレーム
     * @param {Object} mergeSettings - readDialogSettings() の結果
     * @param {number} rulerToPoint - 定規の単位 1 あたりのポイント数
     * @returns {void}
     */
    function applyStyle(textFrame, mergeSettings, rulerToPoint) {
        if (mergeSettings.justify) {
            textFrame.textRange.paragraphAttributes.justification = Justification.FULLJUSTIFYLASTLINELEFT;
        }
        if (mergeSettings.insetSpacingOn) {
            var spacingValue = parseFloat(mergeSettings.insetSpacingText);
            if (!isNaN(spacingValue)) textFrame.spacing = spacingValue * rulerToPoint;
        }
    }

    /**
     * ［エリア内文字の高さ］の設定を適用する
     * @param {TextFrame} textFrame - 対象のフレーム
     * @param {string} frameHeightMode - "none" / "fit" / "auto"
     * @returns {void}
     */
    function applyFrameHeight(textFrame, frameHeightMode) {
        if (frameHeightMode === "none") return;
        if (textFrame.kind !== TextType.AREATEXT) return;
        setAreaTextAutoSize(textFrame, true);
        if (frameHeightMode === "fit") {
            /* 一度だけ合わせて、自動サイズ調整はオフに戻す / Fit once, then turn Auto Size back off */
            setAreaTextAutoSize(textFrame, false);
        }
    }

    /**
     * 選択したテキストを連結し、全体を囲む1つのエリア内文字にまとめる
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {Object} mergeSettings - readDialogSettings() の結果
     * @param {number} rulerToPoint - 定規の単位 1 あたりのポイント数
     * @returns {void}
     */
    function mergeTextFrames(doc, selectedItems, mergeSettings, rulerToPoint) {
        var textFrames = collectMergeTargets(selectedItems, mergeSettings.leftToRight);
        if (textFrames.length < 2) {
            alert(getLabel("alert.needTwo"));
            return;
        }

        var combinedText = [];
        for (var i = 0; i < textFrames.length; i++) {
            combinedText.push(textFrames[i].contents);
        }

        /* 書式保持用：フレームごとの属性と段落数を収集 / Collect frame-level attributes and paragraph counts */
        var frameAttributes = [];
        var frameParagraphCounts = [];
        if (mergeSettings.preserveFormatting) {
            for (var frameIndex = 0; frameIndex < textFrames.length; frameIndex++) {
                frameAttributes.push(readCharacterAttributes(textFrames[frameIndex].textRange.characterAttributes));
                frameParagraphCounts.push(textFrames[frameIndex].paragraphs.length);
            }
        }

        var combinedBounds = getCombinedBounds(textFrames);
        var areaRect = doc.pathItems.rectangle(combinedBounds[1], combinedBounds[0],
            combinedBounds[2] - combinedBounds[0], combinedBounds[1] - combinedBounds[3]);
        var newTextFrame = doc.textFrames.areaText(areaRect);
        var joinedText = combinedText.join("\r");
        if (mergeSettings.removeLineBreaks) {
            joinedText = joinedText.replace(/[\r\n]/g, '');
        }
        newTextFrame.contents = joinedText;

        /* 書式の適用 / Apply character attributes */
        if (mergeSettings.preserveFormatting && !mergeSettings.removeLineBreaks) {
            applyFormattingPerFrame(newTextFrame, frameAttributes, frameParagraphCounts);
        } else {
            /* 先頭フレームの書式を一括適用 / Apply first frame attributes uniformly */
            writeCharacterAttributes(newTextFrame.textRange.characterAttributes,
                readCharacterAttributes(textFrames[0].textRange.characterAttributes));
        }

        for (var removeIndex = textFrames.length - 1; removeIndex >= 0; removeIndex--) {
            textFrames[removeIndex].remove();
        }

        /* スタイルを適用 / Apply style */
        applyStyle(newTextFrame, mergeSettings, rulerToPoint);

        /* 新しいエリア内文字を選択 / Select new area text frame */
        newTextFrame.selected = true;

        /* エリア内文字の高さを適用 / Apply area text height */
        applyFrameHeight(newTextFrame, mergeSettings.frameHeightMode);

        /* 長方形の枠を付ける：新規線を追加し、長方形効果で囲む / Add a border: a new stroke plus a rectangle effect */
        if (mergeSettings.addBorder) {
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            applyRectangleShapeEffect(newTextFrame);
        }
        app.redraw();
    }

    // =========================================
    // 交換 / Swap
    // =========================================

    /**
     * 2つのテキストの文字列を入れ替える（位置と書式はそのまま）
     * @param {TextFrame} firstFrame - 1つ目のテキスト
     * @param {TextFrame} secondFrame - 2つ目のテキスト
     * @returns {void}
     */
    function swapTextContents(firstFrame, secondFrame) {
        var firstContents = firstFrame.contents;
        firstFrame.contents = secondFrame.contents;
        secondFrame.contents = firstContents;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを表示し、選んだ操作（マージ／スレッド／交換）を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;

        if (!doc.selection || doc.selection.length < 1) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* グループは中のテキストフレームに展開 / Expand groups into the text frames inside */
        var selectedItems = expandGroupsInSelection(doc.selection);

        /* テキストフレームが含まれているか検証 / Verify selection contains text frames */
        if (!containsTextFrame(selectedItems)) {
            alert(getLabel("alert.noTextFrame"));
            return;
        }

        /* 1つ選択時：スレッドでなければエリア内文字の高さだけ。ポイント文字などは対象外 / Single selection: unthreaded area text gets Area Text Height only; other kinds exit */
        var heightOnly = false;
        if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame" && !isThreadedFrame(selectedItems[0])) {
            if (selectedItems[0].kind !== TextType.AREATEXT) {
                alert(getLabel("alert.notThreaded"));
                return;
            }
            heightOnly = true;
        }

        /* ルーラー単位 / Ruler units */
        var rulerUnit = getUnitInfo("rulerType");

        var dialogControls = buildDialog(selectedItems, rulerUnit.label, heightOnly);
        if (dialogControls.window.show() !== 1) return;
        var dialogSettings = readDialogSettings(dialogControls);

        /* エリア内文字の高さだけ / Area Text Height only */
        if (heightOnly) {
            applyFrameHeight(selectedItems[0], dialogSettings.frameHeightMode);
            app.redraw();
            return;
        }

        /* スレッドテキスト / Thread text */
        if (dialogSettings.operation === "thread") {
            runThreadAction(doc, selectedItems, dialogSettings.threadAction);
            return;
        }

        /* 交換 / Swap */
        if (dialogSettings.operation === "swap") {
            swapTextContents(selectedItems[0], selectedItems[1]);
            return;
        }

        /* マージ / Merge */
        mergeTextFrames(doc, selectedItems, dialogSettings, rulerUnit.pointsPerUnit);
    }

    main();

})();
