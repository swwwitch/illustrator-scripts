#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のテキストフレームを1つのエリア内文字にまとめたり、逆に分割したりします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MultiAreaText.md

### Overview

Merges several text frames into a single area text, or splits one back out.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MultiAreaText.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MultiAreaText";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-03-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

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
            title: { ja: "複数のテキストフレーム", en: "Multiple Text Frames" }
        },
        panel: {
            option: { ja: "オプション", en: "Option" },
            order: { ja: "順序", en: "Order" },
            threadText: { ja: "スレッドテキスト", en: "Thread Text" },
            style: { ja: "スタイル", en: "Style" },
            frameHeight: { ja: "フレームの高さ", en: "Frame Height" }
        },
        radio: {
            merge: { ja: "マージ", en: "Merge" },
            thread: { ja: "スレッドテキスト", en: "Thread Text" },
            swap: { ja: "交換（文字列のみ）", en: "Swap (Text Only)" },
            topToBottom: { ja: "上から", en: "Top to Bottom" },
            leftToRight: { ja: "左から", en: "Left to Right" },
            threadLink: { ja: "リンク", en: "Link" },
            threadUnlink: { ja: "リンクを解除", en: "Unlink" },
            threadAdd: { ja: "スレッドに追加", en: "Add to Thread" },
            threadRelease: { ja: "スレッドから除外", en: "Release from Thread" },
            threadRelease2: { ja: "スレッドから除外（2）", en: "Release from Thread (2)" },
            heightNone: { ja: "何もしない", en: "None" },
            heightFit: { ja: "フィット", en: "Fit" },
            heightAuto: { ja: "自動サイズ調整", en: "Auto Size" }
        },
        checkbox: {
            removeLineBreaks: { ja: "改行削除", en: "Remove Line Breaks" },
            preserveFormatting: { ja: "書式保持（段落単位）", en: "Preserve Formatting (per paragraph)" },
            insetSpacing: { ja: "外枠からの間隔", en: "Inset Spacing" },
            justify: { ja: "均等配置（最終行左揃え）", en: "Justify (Last Line Left)" },
            appearance: { ja: "アピアランス", en: "Appearance" }
        },
        tooltip: {
            merge: {
                ja: "選択したテキストを連結し、全体を囲む1つのエリア内文字にまとめます",
                en: "Joins the selected text into one area text that covers them all."
            },
            thread: {
                ja: "［スレッドテキスト］パネルで選んだ操作を実行します",
                en: "Runs the action chosen in the Thread Text panel."
            },
            swap: {
                ja: "2つのテキストの文字列を入れ替えます。フレームと書式はそのままです",
                en: "Swaps the text of the two frames; the frames and formatting stay put."
            },
            threadLink: {
                ja: "選択したフレームをスレッドでつなぎます。すでにつながっているフレームがあれば、いったん解除してからつなぎ直します",
                en: "Threads the selected frames, removing any existing threading first."
            },
            threadUnlink: {
                ja: "［スレッドのリンクを解除］を実行します",
                en: "Runs Remove Threading."
            },
            threadAdd: {
                ja: "スレッドをいったん解除し、選択したフレームをまとめてつなぎ直します",
                en: "Removes the threading, then threads the selected frames again."
            },
            threadRelease: {
                ja: "選択したフレームをスレッドから外し、中の文字をそのフレームの位置に独立したエリア内文字として残します（文字属性は段落ごと）",
                en: "Releases the selected frames and keeps their text as separate area text in place (character attributes per paragraph)."
            },
            threadRelease2: {
                ja: "選択したフレームを複製して独立させ、元のフレームとその文字をスレッドから削除します",
                en: "Duplicates each selected frame as a standalone copy, then deletes the original and its text from the thread."
            },
            removeLineBreaks: {
                ja: "連結したテキストの改行をすべて削除します（書式保持がオンのときは使えません）",
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
            appearance: {
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
            noSelection: { ja: "テキストフレームを選択してください。", en: "Please select a text frame." },
            noTextFrame: { ja: "選択にテキストフレームが含まれていません。", en: "Selection does not contain any text frames." },
            notThreaded: { ja: "選択されたテキストフレームはスレッドテキストではありません。", en: "The selected text frame is not threaded text." },
            needTwo: { ja: "テキストフレームを2つ以上選択してください。", en: "Please select 2 or more text frames." }
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
        for (var n = 0; n < COPIED_CHARACTER_ATTRIBUTES.length; n++) {
            attributeSnapshot[COPIED_CHARACTER_ATTRIBUTES[n]] = sourceAttributes[COPIED_CHARACTER_ATTRIBUTES[n]];
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
        for (var n = 0; n < COPIED_CHARACTER_ATTRIBUTES.length; n++) {
            targetAttributes[COPIED_CHARACTER_ATTRIBUTES[n]] = attributeSnapshot[COPIED_CHARACTER_ATTRIBUTES[n]];
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
     * ダイアログを組み立てる（表示はしない）
     * @param {Array} selectedItems - 選択中のオブジェクト
     * @param {string} rulerLabel - 定規の単位の表示名
     * @returns {Object} ダイアログ本体（window）と各コントロール
     */
    function buildDialog(selectedItems, rulerLabel) {
        var isSingle = (selectedItems.length === 1);
        var hasNonAreaText = containsNonAreaText(selectedItems);
        var ui = {};

        ui.window = new Window("dialog", getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        var columnsGroup = ui.window.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];

        /* 左カラム：オプション、順序、スレッドテキスト / Left column: Option, Order, Thread Text */
        var leftColumn = addColumnGroup(columnsGroup);

        var optionPanel = addOptionPanel(leftColumn, "panel.option");
        ui.rbMerge = addControl(optionPanel, "radiobutton", "radio.merge", "tooltip.merge");
        ui.rbThread = addControl(optionPanel, "radiobutton", "radio.thread", "tooltip.thread");
        ui.rbSwap = addControl(optionPanel, "radiobutton", "radio.swap", "tooltip.swap");
        if (isSingle) {
            ui.rbMerge.value = false;
            ui.rbThread.value = true;
        } else {
            ui.rbMerge.value = true;
        }

        var orderPanel = addOptionPanel(leftColumn, "panel.order");
        ui.rbTopToBottom = addControl(orderPanel, "radiobutton", "radio.topToBottom");
        ui.rbLeftToRight = addControl(orderPanel, "radiobutton", "radio.leftToRight");
        ui.rbTopToBottom.value = true;

        var threadPanel = addOptionPanel(leftColumn, "panel.threadText");
        ui.rbThreadLink = addControl(threadPanel, "radiobutton", "radio.threadLink", "tooltip.threadLink");
        ui.rbThreadUnlink = addControl(threadPanel, "radiobutton", "radio.threadUnlink", "tooltip.threadUnlink");
        ui.rbThreadAdd = addControl(threadPanel, "radiobutton", "radio.threadAdd", "tooltip.threadAdd");
        ui.rbThreadRelease = addControl(threadPanel, "radiobutton", "radio.threadRelease", "tooltip.threadRelease");
        ui.rbThreadRelease2 = addControl(threadPanel, "radiobutton", "radio.threadRelease2", "tooltip.threadRelease2");
        if (isSingle) {
            ui.rbThreadRelease.value = true;
        } else {
            ui.rbThreadLink.value = true;
        }

        /* 右カラム：スタイル、フレームの高さ / Right column: Style, Frame Height */
        var rightColumn = addColumnGroup(columnsGroup);

        var stylePanel = addOptionPanel(rightColumn, "panel.style");
        ui.cbRemoveLineBreaks = addControl(stylePanel, "checkbox", "checkbox.removeLineBreaks", "tooltip.removeLineBreaks");
        ui.cbRemoveLineBreaks.value = false;
        ui.cbPreserveFormatting = addControl(stylePanel, "checkbox", "checkbox.preserveFormatting", "tooltip.preserveFormatting");
        ui.cbPreserveFormatting.value = true;
        ui.cbPreserveFormatting.onClick = function () { ui.cbRemoveLineBreaks.enabled = !ui.cbPreserveFormatting.value; };
        var insetSpacingGroup = stylePanel.add("group");
        ui.cbInsetSpacing = addControl(insetSpacingGroup, "checkbox", "checkbox.insetSpacing", "tooltip.insetSpacing");
        ui.cbInsetSpacing.value = false;
        ui.insetSpacingInput = insetSpacingGroup.add("edittext", undefined, "1");
        ui.insetSpacingInput.characters = SPACING_FIELD_CHARACTERS;
        insetSpacingGroup.add("statictext", undefined, rulerLabel);
        ui.cbInsetSpacing.onClick = function () { ui.insetSpacingInput.enabled = ui.cbInsetSpacing.value; };
        ui.cbJustify = addControl(stylePanel, "checkbox", "checkbox.justify");
        ui.cbJustify.value = true;
        ui.cbAppearance = addControl(stylePanel, "checkbox", "checkbox.appearance", "tooltip.appearance");
        ui.cbAppearance.value = false;

        var heightPanel = addOptionPanel(rightColumn, "panel.frameHeight");
        ui.rbHeightNone = addControl(heightPanel, "radiobutton", "radio.heightNone");
        ui.rbHeightFit = addControl(heightPanel, "radiobutton", "radio.heightFit", "tooltip.heightFit");
        ui.rbHeightAuto = addControl(heightPanel, "radiobutton", "radio.heightAuto", "tooltip.heightAuto");
        ui.rbHeightFit.value = true;

        /* パネルの有効/無効を切り替え / Toggle panel enabled state */
        var updatePanels = function () {
            /* 1つのときマージ・交換を無効化 / Disable merge and swap when single object */
            ui.rbMerge.enabled = !isSingle;
            ui.rbSwap.enabled = (selectedItems.length === 2);
            /* ポイント文字/パステキスト含有時スレッドを無効化 / Disable thread when non-area text present */
            ui.rbThread.enabled = !hasNonAreaText;
            /* 交換・スレッドのとき順序を無効化 / Disable order when swap or thread is selected */
            var orderEnabled = !ui.rbSwap.value && !ui.rbThread.value;
            ui.rbTopToBottom.enabled = orderEnabled;
            ui.rbLeftToRight.enabled = orderEnabled;
            /* スレッドテキストパネルはスレッド選択時のみ有効 / Thread text panel enabled only when thread is selected */
            var threadPanelEnabled = ui.rbThread.value;
            ui.rbThreadLink.enabled = threadPanelEnabled;
            ui.rbThreadUnlink.enabled = threadPanelEnabled;
            ui.rbThreadRelease.enabled = threadPanelEnabled;
            ui.rbThreadAdd.enabled = threadPanelEnabled;
            ui.rbThreadRelease2.enabled = threadPanelEnabled;
            /* 交換・1つ選択のときスタイルを無効化 / Disable style when swap is selected or single object */
            var styleEnabled = !ui.rbSwap.value && !isSingle;
            ui.cbRemoveLineBreaks.enabled = styleEnabled && !ui.cbPreserveFormatting.value;
            ui.cbPreserveFormatting.enabled = styleEnabled;
            ui.cbInsetSpacing.enabled = styleEnabled;
            ui.insetSpacingInput.enabled = styleEnabled && ui.cbInsetSpacing.value;
            ui.cbJustify.enabled = styleEnabled;
            ui.cbAppearance.enabled = styleEnabled;
            /* 交換・スレッド・1つ選択のときフレームの高さを無効化 / Disable frame height when swap, thread, or single object */
            var heightEnabled = !ui.rbThread.value && !ui.rbSwap.value && !isSingle;
            ui.rbHeightNone.enabled = heightEnabled;
            ui.rbHeightFit.enabled = heightEnabled;
            ui.rbHeightAuto.enabled = heightEnabled;
        };
        updatePanels();
        ui.rbMerge.onClick = updatePanels;
        ui.rbThread.onClick = updatePanels;
        ui.rbSwap.onClick = updatePanels;

        var btnRowGroup = ui.window.add("group");
        btnRowGroup.alignment = ["right", "center"];
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return ui;
    }

    /**
     * ダイアログのコントロールから処理設定を読み取る
     * @param {Object} ui - buildDialog() の結果
     * @returns {Object} 処理設定
     */
    function readDialogSettings(ui) {
        var threadAction = "link";
        if (ui.rbThreadUnlink.value) threadAction = "unlink";
        else if (ui.rbThreadRelease.value) threadAction = "release";
        else if (ui.rbThreadRelease2.value) threadAction = "releaseByDuplicate";
        else if (ui.rbThreadAdd.value) threadAction = "add";

        var frameHeightMode = "none";
        if (!ui.rbHeightNone.value) {
            if (ui.rbHeightFit.value) frameHeightMode = "fit";
            else if (ui.rbHeightAuto.value) frameHeightMode = "auto";
        }

        return {
            operation: ui.rbThread.value ? "thread" : (ui.rbSwap.value ? "swap" : "merge"),
            threadAction: threadAction,
            leftToRight: ui.rbLeftToRight.value,
            removeLineBreaks: ui.cbRemoveLineBreaks.value,
            preserveFormatting: ui.cbPreserveFormatting.value,
            insetSpacingOn: ui.cbInsetSpacing.value,
            insetSpacingText: ui.insetSpacingInput.text,
            justify: ui.cbJustify.value,
            appearance: ui.cbAppearance.value,
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
        for (var p = 0; p < textFrame.paragraphs.length; p++) {
            /* 読めない段落（末尾の空段落など）で打ち切る / stop at a paragraph that cannot be read, such as a trailing empty one */
            try {
                paragraphAttributes.push(readCharacterAttributes(textFrame.paragraphs[p].characterAttributes));
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
        for (var p = 1; p < paragraphAttributes.length; p++) {
            try {
                writeCharacterAttributes(newTextFrame.paragraphs[p].characterAttributes, paragraphAttributes[p]);
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
        /* 選択の読み取りに失敗しても処理は続ける / keep going even if the selection cannot be read */
        try {
            for (var i = 0; i < doc.selection.length; i++) savedSelection.push(doc.selection[i]);
        } catch (e) { }
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
            for (var ck = 0; ck < frameSnapshots.length; ck++) {
                if (isThreadedFrame(frameSnapshots[ck].frame)) {
                    releaseOk = false;
                    break;
                }
            }
        } catch (e) {
            releaseOk = false;
        }

        if (!releaseOk) {
            alert("スレッド解除に失敗したため、内容は変更していません。\n（releaseThreadedTextSelection が実行できませんでした）");
            return;
        }

        /* スレッドから外れたので、元の内容を消しても連結先に影響しない / Safe to clear now that the chain is broken */
        for (var clr = 0; clr < frameSnapshots.length; clr++) {
            try {
                frameSnapshots[clr].frame.contents = "";
            } catch (e) { }
        }

        /* 保存した情報から新しいエリア内文字を作成 / Create new area text from saved info */
        for (var rj = 0; rj < frameSnapshots.length; rj++) {
            recreateAreaText(doc, frameSnapshots[rj]);
        }

        /* 空になり、スレッドから外れた元のフレームだけを削除 / Remove only the originals that are empty and unthreaded */
        for (var rm = 0; rm < frameSnapshots.length; rm++) {
            /* 無効になったフレームは読み取りや削除で例外になる / an invalid frame throws on read or remove */
            try {
                var originalFrame = frameSnapshots[rm].frame;
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
            if (selectedItems[i].typename === "TextFrame" &&
                (selectedItems[i].kind === TextType.AREATEXT || selectedItems[i].kind === TextType.POINTTEXT || selectedItems[i].kind === TextType.PATHTEXT)) {
                textFrames.push(selectedItems[i]);
            }
        }

        /* ソート（左から右、または上から下） / Sort (left-to-right or top-to-bottom) */
        if (leftToRight) {
            textFrames.sort(function (a, b) {
                return a.geometricBounds[0] - b.geometricBounds[0];
            });
        } else {
            textFrames.sort(function (a, b) {
                return b.geometricBounds[1] - a.geometricBounds[1];
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
        for (var fi = 0; fi < frameAttributes.length; fi++) {
            for (var pi = 0; pi < frameParagraphCounts[fi]; pi++) {
                /* 末尾の空段落は Error 1302 になるのでスキップ / skip the trailing empty paragraph (Error 1302) */
                try {
                    writeCharacterAttributes(newTextFrame.paragraphs[paragraphIndex].characterAttributes, frameAttributes[fi]);
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
        if (textFrame.typename !== "TextFrame") return;
        /* フレームが無効な状態の場合スキップ / skip when the frame is in an invalid state */
        try {
            if (mergeSettings.justify) {
                textFrame.textRange.paragraphAttributes.justification = Justification.FULLJUSTIFYLASTLINELEFT;
            }
            if (mergeSettings.insetSpacingOn && textFrame.kind === TextType.AREATEXT) {
                var spacingValue = parseFloat(mergeSettings.insetSpacingText);
                if (!isNaN(spacingValue)) {
                    textFrame.spacing = spacingValue * rulerToPoint;
                }
            }
        } catch (e) { }
    }

    /**
     * フレームの高さを適用する（エリア内文字のみ）
     * @param {TextFrame} textFrame - 対象のフレーム
     * @param {string} frameHeightMode - "none" / "fit" / "auto"
     * @returns {void}
     */
    function applyFrameHeight(textFrame, frameHeightMode) {
        if (frameHeightMode === "none") return;
        if (textFrame.typename !== "TextFrame" || textFrame.kind !== TextType.AREATEXT) return;
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
            for (var ci = 0; ci < textFrames.length; ci++) {
                frameAttributes.push(readCharacterAttributes(textFrames[ci].textRange.characterAttributes));
                frameParagraphCounts.push(textFrames[ci].paragraphs.length);
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

        for (var r = textFrames.length - 1; r >= 0; r--) {
            textFrames[r].remove();
        }

        /* スタイルを適用 / Apply style */
        applyStyle(newTextFrame, mergeSettings, rulerToPoint);

        /* 新しいエリア内文字を選択 / Select new area text frame */
        newTextFrame.selected = true;

        /* フレームの高さを適用 / Apply frame height */
        applyFrameHeight(newTextFrame, mergeSettings.frameHeightMode);

        /* アピアランスを適用：新規線を追加し、長方形効果で囲む / Apply appearance: add a new stroke and a rectangle effect */
        if (mergeSettings.appearance) {
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            applyRectangleShapeEffect(newTextFrame);
        }
        app.redraw();
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
        var selectedItems = doc.selection;

        if (!selectedItems || selectedItems.length < 1) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* テキストフレームが含まれているか検証 / Verify selection contains text frames */
        if (!containsTextFrame(selectedItems)) {
            alert(getLabel("alert.noTextFrame"));
            return;
        }

        /* 1つ選択時：スレッドテキストでなければ終了 / Single selection: exit if not threaded */
        if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame" && !isThreadedFrame(selectedItems[0])) {
            alert(getLabel("alert.notThreaded"));
            return;
        }

        /* ルーラー単位 / Ruler units */
        var rulerUnit = getUnitInfo("rulerType");

        var dialogUI = buildDialog(selectedItems, rulerUnit.label);
        if (dialogUI.window.show() !== 1) return;
        var dialogSettings = readDialogSettings(dialogUI);

        /* スレッドテキスト / Thread text */
        if (dialogSettings.operation === "thread") {
            runThreadAction(doc, selectedItems, dialogSettings.threadAction);
            return;
        }

        /* 交換 / Swap */
        if (dialogSettings.operation === "swap") {
            if (selectedItems.length !== 2) {
                alert(getLabel("alert.needTwo"));
                return;
            }
            var firstContents = selectedItems[0].contents;
            selectedItems[0].contents = selectedItems[1].contents;
            selectedItems[1].contents = firstContents;
            return;
        }

        /* マージ / Merge */
        mergeTextFrames(doc, selectedItems, dialogSettings, rulerUnit.pointsPerUnit);
    }

    main();

})();
