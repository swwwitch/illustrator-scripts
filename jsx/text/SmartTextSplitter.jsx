#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームを、書式を保ったまま1文字ごとのテキストフレームへ分割します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartTextSplitter.md

### Overview

Splits the selected text frame into one text frame per character, preserving the formatting.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartTextSplitter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartTextSplitter";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartTextSplitter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartTextSplitter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログを開いたときの初期状態 / Initial state of the dialog */
    var DEFAULT_KEEP_STYLE           = true;   /* 書式を保持 / Keep style */
    var DEFAULT_KEEP_SPACES          = false;  /* スペースを残す / Keep spaces */
    var DEFAULT_CONVERT_TO_AREA_TEXT = false;  /* エリア内文字に変換 / Convert to area type */
    var DEFAULT_MERGE_AREA_TEXT      = false;  /* エリア内文字を連結 / Merge area text */
    var DEFAULT_GROUP_MODE           = "all";  /* グループ化："none" / "line" / "all" / Grouping */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 15;                /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS  = [15, 20, 15, 10];  /* パネルの余白 / panel margins */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "テキスト分割", en: "Text Split" }
        },
        panel: {
            options: { ja: "オプション", en: "Options" },
            grouping: { ja: "グループ化", en: "Grouping" }
        },
        checkbox: {
            keepStyle: { ja: "書式を保持", en: "Keep style" },
            keepSpaces: { ja: "スペースを残す", en: "Keep spaces" },
            convertToAreaText: { ja: "エリア内文字に変換", en: "Convert to area type" },
            mergeAreaText: { ja: "エリア内文字を連結", en: "Merge area text" }
        },
        radio: {
            groupNone: { ja: "グループ化しない", en: "Do not group" },
            groupLine: { ja: "行ごとにグループ化", en: "Group each line" },
            groupAll: { ja: "全体をグループ化", en: "Group all" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            keepStyle: { ja: "分割前の文字の書式を、分けたあとのテキストにも引き継ぎます。", en: "Carries the character formatting over to the split text." },
            keepSpaces: { ja: "区切りに使ったスペースを、分けたテキストに残します。", en: "Keeps the separating spaces in the split text." },
            convertToAreaText: { ja: "分けたテキストをポイント文字ではなくエリア内文字にします。", en: "Makes the split text area type instead of point type." },
            mergeAreaText: { ja: "エリア内文字どうしを1つのテキストにつなげます。", en: "Joins area text frames into a single text object." },
            groupNone: { ja: "分けたテキストをそのまま並べます。", en: "Leaves the split text ungrouped." },
            groupLine: { ja: "同じ行から分けたテキストを、行ごとにグループにします。", en: "Groups the pieces from each line together." },
            groupAll: { ja: "分けたテキストをすべて1つのグループにします。", en: "Puts all the split text into one group." }
        }
    };

    /**
     * "category.key" 形式のキーからラベルを取得する
     * @param {string} labelPath - ラベルキー（例: "checkbox.keepStyle"）
     * @returns {string} 現在の言語のラベル文字列（見つからない場合は labelPath をそのまま返す）
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
    // 分割の設定 / Split options
    // =========================================

    /* ダイアログの［OK］で上書きする分割の設定 / Split options, overwritten by the dialog's OK */
    var splitOptions = {
        keepStyle: DEFAULT_KEEP_STYLE,
        keepSpaces: DEFAULT_KEEP_SPACES,
        convertToAreaText: DEFAULT_CONVERT_TO_AREA_TEXT,
        mergeAreaText: DEFAULT_MERGE_AREA_TEXT,
        groupMode: DEFAULT_GROUP_MODE
    };

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} titlePath - パネル見出しのラベルキー
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentWindow, titlePath) {
        var addedPanel = parentWindow.add("panel", undefined, getLabel(titlePath));
        addedPanel.orientation = "column";
        addedPanel.alignChildren = ["left", "top"];
        addedPanel.margins = PANEL_MARGINS;
        return addedPanel;
    }

    /**
     * チェックボックスかラジオボタンを、LABELS のキーで tooltip 付きで追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} controlType - "checkbox" / "radiobutton"
     * @param {string} labelKey - LABELS.checkbox（または LABELS.radio）と LABELS.tooltip のキー
     * @returns {Checkbox|RadioButton} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var labelCategory = (controlType === "radiobutton") ? "radio." : "checkbox.";
        var addedControl = parentContainer.add(controlType, undefined, getLabel(labelCategory + labelKey));
        addedControl.helpTip = getLabel("tooltip." + labelKey);
        return addedControl;
    }

    /**
     * ダイアログを組み立てる（イベントは showSplitDialog() で結び付ける）
     * @returns {Object} ダイアログと各コントロール
     */
    function buildDialog() {
        var splitDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        splitDialog.orientation = "column";
        splitDialog.alignChildren = ["fill", "top"];
        splitDialog.margins = DIALOG_MARGINS;

        /* パネルの並び：オプション → グループ化 / Panel order: Options → Grouping */
        var optionsPanel = addPanel(splitDialog, "panel.options");
        var keepStyleCheckbox = addLabeledControl(optionsPanel, "checkbox", "keepStyle");
        keepStyleCheckbox.value = DEFAULT_KEEP_STYLE;
        var keepSpacesCheckbox = addLabeledControl(optionsPanel, "checkbox", "keepSpaces");
        keepSpacesCheckbox.value = DEFAULT_KEEP_SPACES;
        var convertToAreaTextCheckbox = addLabeledControl(optionsPanel, "checkbox", "convertToAreaText");
        convertToAreaTextCheckbox.value = DEFAULT_CONVERT_TO_AREA_TEXT;
        var mergeAreaTextCheckbox = addLabeledControl(optionsPanel, "checkbox", "mergeAreaText");
        mergeAreaTextCheckbox.value = DEFAULT_MERGE_AREA_TEXT;

        var groupingPanel = addPanel(splitDialog, "panel.grouping");
        var groupNoneRadio = addLabeledControl(groupingPanel, "radiobutton", "groupNone");
        var groupLineRadio = addLabeledControl(groupingPanel, "radiobutton", "groupLine");
        var groupAllRadio = addLabeledControl(groupingPanel, "radiobutton", "groupAll");
        groupNoneRadio.value = (DEFAULT_GROUP_MODE === "none");
        groupLineRadio.value = (DEFAULT_GROUP_MODE === "line");
        groupAllRadio.value = (DEFAULT_GROUP_MODE === "all");

        var btnRowGroup = splitDialog.add("group");
        btnRowGroup.alignment = "right";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            splitDialog: splitDialog,
            keepStyleCheckbox: keepStyleCheckbox,
            keepSpacesCheckbox: keepSpacesCheckbox,
            convertToAreaTextCheckbox: convertToAreaTextCheckbox,
            mergeAreaTextCheckbox: mergeAreaTextCheckbox,
            groupNoneRadio: groupNoneRadio,
            groupLineRadio: groupLineRadio,
            btnOK: btnOK,
            btnCancel: btnCancel
        };
    }

    /**
     * ダイアログの状態を分割の設定にまとめる
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {Object} 分割の設定（keepStyle / keepSpaces / convertToAreaText / mergeAreaText / groupMode）
     */
    function readSplitOptions(dialogControls) {
        var convertToAreaText = !!dialogControls.convertToAreaTextCheckbox.value;
        return {
            keepStyle: !!dialogControls.keepStyleCheckbox.value,
            keepSpaces: !!dialogControls.keepSpacesCheckbox.value,
            convertToAreaText: convertToAreaText,
            mergeAreaText: convertToAreaText && !!dialogControls.mergeAreaTextCheckbox.value,
            groupMode: dialogControls.groupNoneRadio.value ? "none" : (dialogControls.groupLineRadio.value ? "line" : "all")
        };
    }

    /**
     * ダイアログを表示し、［OK］で選択中のテキストを分割する（分割はダイアログを閉じる前に行う）
     * @returns {void}
     */
    function showSplitDialog() {
        var dialogControls = buildDialog();
        var splitDialog = dialogControls.splitDialog;
        var convertToAreaTextCheckbox = dialogControls.convertToAreaTextCheckbox;
        var mergeAreaTextCheckbox = dialogControls.mergeAreaTextCheckbox;

        /**
         * 連結は「エリア内文字に変換」がONのときだけ有効にする
         * @returns {void}
         */
        function syncAreaOptionsEnabled() {
            mergeAreaTextCheckbox.enabled = !!convertToAreaTextCheckbox.value;
            if (!mergeAreaTextCheckbox.enabled) mergeAreaTextCheckbox.value = false;
        }

        convertToAreaTextCheckbox.onClick = syncAreaOptionsEnabled;
        syncAreaOptionsEnabled();

        dialogControls.btnOK.onClick = function () {
            splitOptions = readSplitOptions(dialogControls);
            /* 分割中の失敗は黙って閉じる / Any failure during the split is ignored */
            try { splitSelectedTextFrames(); } catch (e) { }
            splitDialog.close(1);
        };
        dialogControls.btnCancel.onClick = function () {
            splitDialog.close(0);
        };

        splitDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のテキストフレームを1文字ずつに分割する
     * @returns {void}
     */
    function splitSelectedTextFrames() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) return;

        /* TextFrame のみ抽出 / Collect TextFrames only */
        var targetFrames = [];
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem && selectedItem.typename === "TextFrame") targetFrames.push(selectedItem);
        }

        for (var k = 0; k < targetFrames.length; k++) {
            // 「書式を保持」OFF: 分割前に書式を整理（先頭文字のフォント情報のみ保持）
            if (!splitOptions.keepStyle) {
                try { stripStyleKeepFirstFont(targetFrames[k]); } catch (e) { }
            }
            /* 1つのフレームで失敗しても残りを処理する / One failing frame must not stop the rest */
            try { splitTextFrameHighPrecision(targetFrames[k]); } catch (e2) { }
        }
    }

    // =========================================
    // 属性の読み書き / Attribute helpers
    // =========================================

    /**
     * 文字属性を1つ書き込む（失敗しても続ける）
     * @param {CharacterAttributes} targetAttributes - 書き込み先
     * @param {string} propertyName - 属性名
     * @param {*} value - 書き込む値
     * @returns {void}
     */
    function setAttributeSafely(targetAttributes, propertyName, value) {
        try { targetAttributes[propertyName] = value; } catch (e) { }
    }

    /**
     * 文字属性を1つコピーする（失敗しても続ける）
     * @param {CharacterAttributes} targetAttributes - コピー先
     * @param {CharacterAttributes} sourceAttributes - コピー元
     * @param {string} propertyName - 属性名
     * @returns {void}
     */
    function copyAttributeSafely(targetAttributes, sourceAttributes, propertyName) {
        try { targetAttributes[propertyName] = sourceAttributes[propertyName]; } catch (e) { }
    }

    /**
     * 文字属性を1つ読む
     * @param {CharacterAttributes} sourceAttributes - 読み取り元
     * @param {string} propertyName - 属性名
     * @returns {*} 属性の値（読めなければ null）
     */
    function readAttributeSafely(sourceAttributes, propertyName) {
        try { return sourceAttributes[propertyName]; } catch (e) { return null; }
    }

    /**
     * エリア内文字かを返す（種類が読めなければ false）
     * @param {TextFrame} textFrame - 調べるテキストフレーム
     * @returns {boolean} エリア内文字なら true
     */
    function isAreaTextSafely(textFrame) {
        try { return textFrame.kind === TextType.AREATEXT; } catch (e) { return false; }
    }

    // =========================================
    // 書式削除（先頭文字のフォント情報だけ保持）/ Style reset
    // =========================================

    /**
     * 文字属性を先頭文字のフォント・サイズ、黒（K100）にそろえ、それ以外を初期化する（できる範囲で）
     * @param {CharacterAttributes} targetAttributes - 対象の文字属性
     * @param {TextFont|null} keepFont - そろえるフォント
     * @param {number|null} keepSize - そろえるサイズ
     * @returns {void}
     */
    function resetCharacterFormatting(targetAttributes, keepFont, keepSize) {
        if (keepFont) setAttributeSafely(targetAttributes, "textFont", keepFont);
        if (keepSize != null) setAttributeSafely(targetAttributes, "size", keepSize);

        // テキストカラーを黒に統一
        try {
            var blackColor = new GrayColor();
            blackColor.gray = 100; // K100
            targetAttributes.fillColor = blackColor;
        } catch (e) { }

        setAttributeSafely(targetAttributes, "baselineShift", 0);
        setAttributeSafely(targetAttributes, "horizontalScale", 100);
        setAttributeSafely(targetAttributes, "verticalScale", 100);
        setAttributeSafely(targetAttributes, "rotation", 0);
        setAttributeSafely(targetAttributes, "tracking", 0);

        // カーニング/行送り
        /* KerningMethod は環境によって定義されていないことがある / KerningMethod may be undefined */
        try { targetAttributes.kerningMethod = KerningMethod.METRICS; } catch (e2) { }
        setAttributeSafely(targetAttributes, "autoLeading", true);
    }

    /**
     * 書式を削除する（先頭文字のフォントとサイズだけ残す）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function stripStyleKeepFirstFont(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;

        var textCharacters = null;
        var firstAttributes = null;
        var frameAttributes = null;
        /* 読めなければ何もしない / Give up when the text cannot be read */
        try {
            var fullRange = textFrame.textRange;
            textCharacters = fullRange.characters;
            if (!textCharacters || textCharacters.length < 1) return;
            firstAttributes = textCharacters[0].characterAttributes;
            frameAttributes = fullRange.characterAttributes;
        } catch (e) { return; }
        if (!firstAttributes || !frameAttributes) return;

        // 先頭文字のフォント情報を保持（実用上、サイズも保持）
        var keepFont = readAttributeSafely(firstAttributes, "textFont");
        var keepSize = readAttributeSafely(firstAttributes, "size");

        resetCharacterFormatting(frameAttributes, keepFont, keepSize);

        // 文字単位の回転/変形が残るケースがあるので、個別にも念押し
        for (var i = 0; i < textCharacters.length; i++) {
            var characterAttributes = null;
            try { characterAttributes = textCharacters[i].characterAttributes; } catch (e2) { characterAttributes = null; }
            if (!characterAttributes) continue;
            resetCharacterFormatting(characterAttributes, keepFont, keepSize);
        }
    }

    // =========================================
    // 高精度分割 / High precision split
    // =========================================

    /**
     * 作ったテキストフレームを後処理する（エリア内文字に変換・連結・グループ化）し、計算用アウトラインと元のフレームを消す
     * @param {TextFrame[]} createdFrames - 分割で作ったテキストフレーム
     * @param {TextFrame} sourceFrame - 分割元のテキストフレーム
     * @param {Object} [outlineInfo] - buildOutlineCharBounds() の戻り値（高精度分割のときだけ）
     * @returns {void}
     */
    function finishSplit(createdFrames, sourceFrame, outlineInfo) {
        var finishedFrames = createdFrames;

        // エリア内文字に変換（UI設定 or マージ時）
        try { if (splitOptions.convertToAreaText || splitOptions.mergeAreaText) finishedFrames = convertTextFramesToAreaText(finishedFrames); } catch (e) { }

        // エリア内文字を連結（スレッド化）
        try { if (splitOptions.mergeAreaText) finishedFrames = threadAreaTextFrames(finishedFrames); } catch (e2) { }

        // グループ化（UI設定）
        try { applyGroupingByMode(finishedFrames, splitOptions.groupMode); } catch (e3) { }

        // 計算用アウトラインを削除
        removeOutlineInfo(outlineInfo);

        // 元のTextFrameを削除
        try { sourceFrame.remove(); } catch (e4) { }
    }

    /**
     * 高精度分割をやめて、作りかけのフレームと計算用アウトラインを片付け、幅の積算による分割に切り替える
     * @param {TextFrame[]} createdFrames - ここまでに作ったテキストフレーム
     * @param {Object} outlineInfo - buildOutlineCharBounds() の戻り値
     * @param {TextFrame} textFrame - 分割元のテキストフレーム
     * @returns {void}
     */
    function fallBackToWidthSplit(createdFrames, outlineInfo, textFrame) {
        removeItems(createdFrames);
        removeOutlineInfo(outlineInfo);
        splitTextFrameFallback(textFrame);
    }

    /**
     * テキストフレームを、アウトラインで測った文字の位置に合わせて1文字ずつに分割する
     * （スペースを残すとき・アウトラインが取れないときは幅の積算による分割）
     * @param {TextFrame} textFrame - 分割するテキストフレーム
     * @returns {void}
     */
    function splitTextFrameHighPrecision(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;

        // スペースを残す場合、アウトラインbounds方式ではスペース位置を保持できないためフォールバックに切替
        if (splitOptions.keepSpaces) {
            splitTextFrameFallback(textFrame);
            return;
        }

        var textCharacters = null;
        var characterCount = 0;
        /* 読めなければ何もしない / Give up when the text cannot be read */
        try {
            textCharacters = textFrame.textRange.characters;
            characterCount = textCharacters.length;
        } catch (e) { return; }
        if (!characterCount || characterCount <= 0) return;

        /* アウトライン bounds を利用 / Use outline bounds */
        var outlineInfo = null;
        try {
            outlineInfo = buildOutlineCharBounds(textFrame);
        } catch (e2) {
            outlineInfo = null;
        }

        if (!(outlineInfo && outlineInfo.ok && outlineInfo.boundsList && outlineInfo.boundsList.length > 0)) {
            /* フォールバック / Fallback */
            splitTextFrameFallback(textFrame);
            return;
        }

        var boundsList = outlineInfo.boundsList;
        var targetLayer = textFrame.layer;

        // スペースはアウトライン側に存在しないため、boundsList は「可視文字のみ」として扱う
        // textCharacters は元テキストの全文字（スペース含む）なので、インデックスを分けて走査する
        var createdFrames = [];
        var boundsIndex = 0; // boundsList index（可視文字用）

        for (var i = 0; i < characterCount; i++) {
            var sourceCharacter = null;
            var characterText = "";
            /* 読めない文字は飛ばす / Skip characters that cannot be read */
            try {
                sourceCharacter = textCharacters[i];
                characterText = sourceCharacter.contents;
            } catch (e3) { continue; }
            if (!sourceCharacter || characterText === "") continue;

            // スペース/タブ/改行などは完全に無視（生成しない / boundsも消費しない）
            if (isIgnoredSpaceChar(characterText)) continue;

            // 可視文字に対して bounds が不足したら、この方式は成立しないのでフォールバック
            if (boundsIndex >= boundsList.length) {
                fallBackToWidthSplit(createdFrames, outlineInfo, textFrame);
                return;
            }

            // 1文字TextFrameを作成
            var characterFrame = null;
            try {
                characterFrame = targetLayer.textFrames.add();
                characterFrame.contents = characterText;
            } catch (e4) {
                fallBackToWidthSplit(createdFrames, outlineInfo, textFrame);
                return;
            }

            // 文字属性をコピー
            copyCharacterAttributes(characterFrame, sourceCharacter);

            // 元TextFrameの変形（回転/拡縮など）を適用（TextFrame.matrix は読み取り専用のことがある / matrix may be read-only）
            try { characterFrame.matrix = textFrame.matrix; } catch (e5) { }

            // まず元フレーム近傍に置く（大外れ回避）
            try { characterFrame.left = textFrame.left; characterFrame.top = textFrame.top; } catch (e6) { }

            // 目標 bounds に一致するように characterFrame を移動
            try {
                moveTextFrameToMatchBounds(characterFrame, boundsList[boundsIndex]);
            } catch (e7) {
                try { characterFrame.remove(); } catch (e8) { }
                fallBackToWidthSplit(createdFrames, outlineInfo, textFrame);
                return;
            }

            createdFrames.push(characterFrame);
            boundsIndex++;
        }

        finishSplit(createdFrames, textFrame, outlineInfo);
    }

    // =========================================
    // アウトラインboundsの取得 / Outline bounds builder
    // =========================================

    /**
     * テキストフレームを複製してアウトライン化する（複製は消費されるので後片付けする）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {GroupItem|null} アウトライン（作れなければ null）
     */
    function createOutlineOfDuplicate(textFrame) {
        var duplicateFrame = null;
        try {
            duplicateFrame = textFrame.duplicate(textFrame.parent, ElementPlacement.PLACEATBEGINNING);
        } catch (e) {
            try { duplicateFrame = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATBEGINNING); } catch (e2) { duplicateFrame = null; }
        }
        if (!duplicateFrame) return null;

        // アウトライン化
        var outlinedGroup = null;
        try {
            outlinedGroup = duplicateFrame.createOutline();
        } catch (e3) {
            try { duplicateFrame.remove(); } catch (e4) { }
            return null;
        }

        // createOutline 後に複製自体が残る環境があるので消す（消費済みなら例外）
        try { duplicateFrame.remove(); } catch (e5) { }

        return outlinedGroup || null;
    }

    /**
     * アウトラインから1文字ずつの bounds を読み順で取り出す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{ok: boolean, outlinedRoot: GroupItem, boundsList: number[][]}} 取り出した bounds（失敗時は ok: false）
     */
    function buildOutlineCharBounds(textFrame) {
        var outlinedGroup = createOutlineOfDuplicate(textFrame);
        if (!outlinedGroup) return { ok: false };

        // outlinedGroup は GroupItem で返ることが多い
        var charItems = [];
        collectCharLikeItems(outlinedGroup, charItems);

        // 取得順は不定なので、座標から読み順に並べ替える
        sortItemsByTextDirection(charItems);

        if (charItems.length === 0) {
            try { outlinedGroup.remove(); } catch (e) { }
            return { ok: false };
        }

        var boundsList = [];
        for (var i = 0; i < charItems.length; i++) {
            /* 字形の無い文字は空のグループで、bounds が例外になる / Glyphless characters throw on geometricBounds */
            try { boundsList.push(charItems[i].geometricBounds); } catch (e2) { }
        }

        return {
            ok: boundsList.length > 0,
            outlinedRoot: outlinedGroup,
            boundsList: boundsList
        };
    }

    /**
     * アウトラインから文字単位のオブジェクトを集める
     * @param {GroupItem} outlinedRoot - アウトラインのグループ
     * @param {PageItem[]} collectedItems - 集めたオブジェクトの入れ物（破壊的に追加）
     * @returns {void}
     */
    function collectCharLikeItems(outlinedRoot, collectedItems) {
        if (!outlinedRoot) return;

        var directItems = [];
        try {
            var rootItems = outlinedRoot.pageItems;
            for (var i = 0; i < rootItems.length; i++) directItems.push(rootItems[i]);
        } catch (e) { }

        if (directItems.length === 1 && directItems[0] && directItems[0].typename === "GroupItem") {
            try {
                var innerItems = directItems[0].pageItems;
                if (innerItems && innerItems.length > 0) {
                    for (var j = 0; j < innerItems.length; j++) collectedItems.push(innerItems[j]);
                    return;
                }
            } catch (e2) { }
        }

        for (var k = 0; k < directItems.length; k++) collectedItems.push(directItems[k]);
    }

    // =========================================
    // 読み順の並べ替え / Reading order
    // =========================================

    /**
     * オブジェクトの bounds と中心・高さをまとめる（bounds が読めないものは除く）
     * @param {PageItem[]} pageItems - 対象のオブジェクト
     * @returns {Object[]} it / L / T / R / B / cx / cy / h / idx を持つ配列
     */
    function collectBoundsEntries(pageItems) {
        var boundsEntries = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (!pageItem) continue;
            var itemBounds = null;
            try { itemBounds = pageItem.geometricBounds; } catch (e) { itemBounds = null; }
            if (!itemBounds || itemBounds.length !== 4) continue;

            var L = itemBounds[0], T = itemBounds[1], R = itemBounds[2], B = itemBounds[3];
            var cx = (L + R) / 2;
            var cy = (T + B) / 2;
            var h = Math.abs(T - B);
            boundsEntries.push({ it: pageItem, L: L, T: T, R: R, B: B, cx: cx, cy: cy, h: h, idx: i });
        }
        return boundsEntries;
    }

    /**
     * bounds のまとまりを Y 方向でクラスタリングし、上→下の行に分ける
     * @param {Object[]} boundsEntries - collectBoundsEntries() の戻り値
     * @returns {{cy: number, items: Object[]}[]} 上→下に並べた行（行内は未整列）
     */
    function clusterEntriesIntoRows(boundsEntries) {
        var rowThreshold = estimateRowThreshold(boundsEntries);

        // 上→下（cy降順）
        boundsEntries.sort(function (p, q) {
            var dy = q.cy - p.cy;
            if (Math.abs(dy) > 0.001) return (dy < 0) ? -1 : 1;
            var dL = p.L - q.L;
            if (Math.abs(dL) > 0.001) return (dL < 0) ? -1 : 1;
            return p.idx - q.idx;
        });

        var rows = [];
        for (var a = 0; a < boundsEntries.length; a++) {
            var currentEntry = boundsEntries[a];
            var placed = false;
            for (var r = 0; r < rows.length; r++) {
                if (Math.abs(currentEntry.cy - rows[r].cy) <= rowThreshold) {
                    rows[r].items.push(currentEntry);
                    rows[r].cy = (rows[r].cy * (rows[r].items.length - 1) + currentEntry.cy) / rows[r].items.length;
                    placed = true;
                    break;
                }
            }
            if (!placed) rows.push({ cy: currentEntry.cy, items: [currentEntry] });
        }

        // 行を上→下
        rows.sort(function (p, q) {
            var dy2 = q.cy - p.cy;
            if (Math.abs(dy2) > 0.001) return (dy2 < 0) ? -1 : 1;
            return 0;
        });
        return rows;
    }

    /**
     * アウトライン側の items を「読み順」に並べ替える（複数行対応。items を書き換える）
     * - まずY（行）でクラスタリングして上→下に並べる
     * - 各行の中は「左端(L)」で左→右
     * - TextFrame.matrix は参照せず、見た目の座標だけで決める
     * @param {PageItem[]} charItems - 並べ替えるオブジェクト
     * @returns {void}
     */
    function sortItemsByTextDirection(charItems) {
        if (!charItems || charItems.length <= 1) return;

        var boundsEntries = collectBoundsEntries(charItems);
        if (boundsEntries.length <= 1) return;

        var rows = clusterEntriesIntoRows(boundsEntries);

        var sortedItems = [];
        for (var ri = 0; ri < rows.length; ri++) {
            var rowItems = rows[ri].items;
            rowItems.sort(function (p, q) {
                var dL2 = p.L - q.L;
                if (Math.abs(dL2) > 0.5) return (dL2 < 0) ? -1 : 1;

                var dT = p.T - q.T;
                if (Math.abs(dT) > 0.5) return (dT < 0) ? 1 : -1;

                var dcx = p.cx - q.cx;
                if (Math.abs(dcx) > 0.5) return (dcx < 0) ? -1 : 1;

                var dcy = p.cy - q.cy;
                if (Math.abs(dcy) > 0.5) return (dcy < 0) ? 1 : -1;

                return p.idx - q.idx;
            });

            for (var j = 0; j < rowItems.length; j++) sortedItems.push(rowItems[j].it);
        }

        for (var k = 0; k < sortedItems.length; k++) charItems[k] = sortedItems[k];
    }

    /**
     * 行分けのしきい値を、高さの中央値の 0.6 倍（最小 2）として求める
     * @param {Object[]} boundsEntries - collectBoundsEntries() の戻り値
     * @returns {number} しきい値（pt）
     */
    function estimateRowThreshold(boundsEntries) {
        var heights = [];
        for (var i = 0; i < boundsEntries.length; i++) {
            if (boundsEntries[i].h && boundsEntries[i].h > 0) heights.push(boundsEntries[i].h);
        }
        if (heights.length === 0) return 8;
        heights.sort(function (firstHeight, secondHeight) { return firstHeight - secondHeight; });
        var medianHeight = heights[Math.floor(heights.length / 2)];
        if (!medianHeight || medianHeight <= 0) medianHeight = heights[0];
        var rowThreshold = medianHeight * 0.6;
        if (rowThreshold < 2) rowThreshold = 2;
        return rowThreshold;
    }

    /**
     * テキストフレームを読み順の行に分ける
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {TextFrame[][]} 上→下の行ごとに、左→右に並べたテキストフレーム
     */
    function framesToRowsInReadingOrder(textFrames) {
        var boundsEntries = collectBoundsEntries(textFrames);
        if (boundsEntries.length === 0) return [];

        var rows = clusterEntriesIntoRows(boundsEntries);

        // 行内 左→右
        var outRows = [];
        for (var ri = 0; ri < rows.length; ri++) {
            var rowItems = rows[ri].items;
            rowItems.sort(function (p, q) {
                var dL2 = p.L - q.L;
                if (Math.abs(dL2) > 0.5) return (dL2 < 0) ? -1 : 1;

                var dT = p.T - q.T;
                if (Math.abs(dT) > 0.5) return (dT < 0) ? 1 : -1;

                return p.idx - q.idx;
            });

            var rowFrames = [];
            for (var j = 0; j < rowItems.length; j++) rowFrames.push(rowItems[j].it);
            outRows.push(rowFrames);
        }

        return outRows;
    }

    // =========================================
    // 文字の判定 / Character checks
    // =========================================

    /**
     * 半角／全角スペースかを返す
     * @param {string} characterText - 調べる文字
     * @returns {boolean} スペースなら true
     */
    function isSpaceChar(characterText) {
        return (characterText === " " || characterText === "\u3000"); // 半角/全角スペース
    }

    /**
     * 分割で無視する文字かを返す（改行は常に、スペース・タブは「スペースを残す」OFFのとき）
     * @param {string} characterText - 調べる文字
     * @returns {boolean} 無視するなら true
     */
    function isIgnoredSpaceChar(characterText) {
        // 改行（段落終端）は常に無視
        if (characterText === "\r" || characterText === "\n") return true;

        // 「スペースを残す」ONのときは半角/全角スペースは無視しない
        if (splitOptions.keepSpaces && (characterText === " " || characterText === "\u3000")) return false;

        // 既定（OFF）: 半角/全角スペース/タブを無視
        return (characterText === " " || characterText === "\t" || characterText === "\u3000");
    }

    // =========================================
    // ユーティリティ / Utils
    // =========================================

    /**
     * 計算用アウトラインを削除する
     * @param {Object|null} outlineInfo - buildOutlineCharBounds() の戻り値
     * @returns {void}
     */
    function removeOutlineInfo(outlineInfo) {
        if (!outlineInfo) return;
        try { if (outlineInfo.outlinedRoot) outlineInfo.outlinedRoot.remove(); } catch (e) { }
    }

    /**
     * オブジェクトをまとめて削除する（消せないものは飛ばす）
     * @param {PageItem[]} pageItems - 削除するオブジェクト
     * @returns {void}
     */
    function removeItems(pageItems) {
        for (var i = 0; i < pageItems.length; i++) {
            try { pageItems[i].remove(); } catch (e) { }
        }
    }

    /**
     * アウトライン化して測った中心が目標 bounds の中心に来るよう、テキストフレームを移動する
     * @param {TextFrame} characterFrame - 移動するテキストフレーム
     * @param {number[]} targetBounds - 目標の bounds [左, 上, 右, 下]
     * @returns {void}
     */
    function moveTextFrameToMatchBounds(characterFrame, targetBounds) {
        if (!characterFrame || characterFrame.typename !== "TextFrame") return;
        if (!targetBounds || targetBounds.length !== 4) return;

        var outlinedGroup = createOutlineOfDuplicate(characterFrame);
        if (!outlinedGroup) return;

        var currentBounds = null;
        try { currentBounds = outlinedGroup.geometricBounds; } catch (e) { currentBounds = null; }
        try { outlinedGroup.remove(); } catch (e2) { }

        if (!currentBounds || currentBounds.length !== 4) return;

        var currentCenter = boundsCenter(currentBounds);
        var targetCenter = boundsCenter(targetBounds);

        var dx = targetCenter[0] - currentCenter[0];
        var dy = targetCenter[1] - currentCenter[1];

        try {
            characterFrame.left += dx;
            characterFrame.top += dy;
        } catch (e3) {
            try { characterFrame.translate(dx, dy); } catch (e4) { }
        }
    }

    /**
     * bounds の中心を返す
     * @param {number[]} itemBounds - [左, 上, 右, 下]
     * @returns {number[]} 中心 [x, y]
     */
    function boundsCenter(itemBounds) {
        var cx = (itemBounds[0] + itemBounds[2]) / 2;
        var cy = (itemBounds[1] + itemBounds[3]) / 2;
        return [cx, cy];
    }

    /* そのままコピーする文字属性 / Character attributes copied as-is */
    var COPIED_CHARACTER_ATTRIBUTES = ["textFont", "size", "horizontalScale", "verticalScale", "tracking", "baselineShift", "rotation"];

    /**
     * 文字属性をコピーする（属性ごとに失敗しても続ける）
     * @param {TextFrame} targetFrame - コピー先のテキストフレーム
     * @param {TextRange} sourceRange - コピー元の文字（またはテキスト範囲）
     * @returns {void}
     */
    function copyCharacterAttributes(targetFrame, sourceRange) {
        if (!targetFrame || targetFrame.typename !== "TextFrame") return;
        if (!sourceRange) return;

        var sourceAttributes = null;
        try { sourceAttributes = sourceRange.characterAttributes; } catch (e) { sourceAttributes = null; }
        if (!sourceAttributes) return;

        var targetAttributes = null;
        try { targetAttributes = targetFrame.textRange.characterAttributes; } catch (e2) { targetAttributes = null; }
        if (!targetAttributes) return;

        for (var i = 0; i < COPIED_CHARACTER_ATTRIBUTES.length; i++) {
            copyAttributeSafely(targetAttributes, sourceAttributes, COPIED_CHARACTER_ATTRIBUTES[i]);
        }

        try {
            if (sourceAttributes.fillColor && sourceAttributes.fillColor.typename !== "NoColor") targetAttributes.fillColor = sourceAttributes.fillColor;
        } catch (e3) { }
        try {
            if (sourceAttributes.strokeColor && sourceAttributes.strokeColor.typename !== "NoColor") {
                targetAttributes.strokeColor = sourceAttributes.strokeColor;
                targetAttributes.strokeWeight = sourceAttributes.strokeWeight;
            }
        } catch (e4) { }

        copyAttributeSafely(targetAttributes, sourceAttributes, "autoLeading");
        try { if (!sourceAttributes.autoLeading) targetAttributes.leading = sourceAttributes.leading; } catch (e5) { }
        copyAttributeSafely(targetAttributes, sourceAttributes, "kerningMethod");
    }

    // =========================================
    // フォールバック：幅を積算して配置 / Fallback: accumulate widths
    // =========================================

    /**
     * 文字の属性を控える（フォールバック分割で新しいフレームに書き戻す用）
     * @param {TextRange} sourceCharacter - 対象の文字
     * @returns {Object} 文字属性のスナップショット（塗り・線がなしのときは null）
     */
    function snapshotCharacterAttributes(sourceCharacter) {
        var sourceAttributes = sourceCharacter.characterAttributes;
        return {
            textFont: sourceAttributes.textFont,
            size: sourceAttributes.size,
            horizontalScale: sourceAttributes.horizontalScale,
            verticalScale: sourceAttributes.verticalScale,
            tracking: sourceAttributes.tracking,
            baselineShift: sourceAttributes.baselineShift,
            rotation: sourceAttributes.rotation,
            fillColor: (sourceAttributes.fillColor.typename === "NoColor") ? null : sourceAttributes.fillColor,
            strokeColor: (sourceAttributes.strokeColor.typename === "NoColor") ? null : sourceAttributes.strokeColor,
            strokeWeight: sourceAttributes.strokeWeight,
            autoLeading: sourceAttributes.autoLeading,
            leading: sourceAttributes.leading,
            kerningMethod: sourceAttributes.kerningMethod
        };
    }

    /**
     * 1文字分のテキストフレームを作り、先行する文字の幅を積算した位置に置く（失敗すると例外）
     * @param {TextFrame} textFrame - 分割元のテキストフレーム
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @param {number} charIndex - 文字の位置
     * @param {TextRange} sourceCharacter - 対象の文字
     * @param {string} characterText - 対象の文字の内容
     * @returns {TextFrame} 作成したテキストフレーム
     */
    function createCharacterFrame(textFrame, targetLayer, charIndex, sourceCharacter, characterText) {
        var characterAttributes = snapshotCharacterAttributes(sourceCharacter);
        var sourceMatrix = textFrame.matrix;

        var characterFrame = targetLayer.textFrames.add();
        characterFrame.contents = characterText;
        applyAttributes(characterFrame, characterAttributes);

        characterFrame.top = textFrame.top;
        characterFrame.left = textFrame.left;
        characterFrame.matrix = sourceMatrix;

        // 先行文字の幅を積算してオフセットを作る（改行/タブなど無視する文字は数えない）
        var offsetX = 0;
        for (var j = 0; j < charIndex; j++) {
            var precedingCharacter = textFrame.textRange.characters[j];
            var precedingText = precedingCharacter.contents;
            if (isIgnoredSpaceChar(precedingText)) continue;

            var measureFrame = targetLayer.textFrames.add();
            measureFrame.contents = precedingText;
            applyAttributes(measureFrame, {
                textFont: precedingCharacter.characterAttributes.textFont,
                size: precedingCharacter.characterAttributes.size,
                horizontalScale: precedingCharacter.characterAttributes.horizontalScale,
                verticalScale: precedingCharacter.characterAttributes.verticalScale,
                tracking: precedingCharacter.characterAttributes.tracking
            });
            offsetX += measureFrame.width;
            measureFrame.remove();
        }

        var angle = getRotationFromMatrix(sourceMatrix);
        var radians = angle * Math.PI / 180;
        var cosA = Math.cos(radians);
        var sinA = Math.sin(radians);

        characterFrame.left = textFrame.left + offsetX * cosA;
        characterFrame.top = textFrame.top - offsetX * sinA;
        return characterFrame;
    }

    /**
     * 文字の幅を積算して1文字ずつのテキストフレームに分割する
     * @param {TextFrame} textFrame - 分割するテキストフレーム
     * @returns {void}
     */
    function splitTextFrameFallback(textFrame) {
        if (!textFrame || textFrame.typename !== "TextFrame") return;

        var textLength = 0;
        try { textLength = textFrame.textRange.characters.length; } catch (e) { textLength = 0; }
        if (!textLength) return;

        var targetLayer = null;
        try { targetLayer = textFrame.layer; } catch (e2) { targetLayer = null; }
        if (!targetLayer) return;

        var createdFrames = [];
        var i, sourceCharacter, characterText;

        if (splitOptions.keepSpaces) {
            // 「スペースを残す」ONのときは、左→右に走査してスペースを直前の文字に結合する
            var lastFrame = null;
            for (i = 0; i < textLength; i++) {
                /* 1文字で失敗しても残りを処理する / One failing character must not stop the rest */
                try {
                    sourceCharacter = textFrame.textRange.characters[i];
                    characterText = sourceCharacter.contents;

                    // 改行/タブは常に無視
                    if (isIgnoredSpaceChar(characterText)) continue;

                    // スペースは直前の文字に結合
                    if (isSpaceChar(characterText)) {
                        if (lastFrame) {
                            try { lastFrame.contents += characterText; } catch (e3) { }
                        }
                        continue;
                    }

                    // 通常文字: TextFrameを作成して配置
                    lastFrame = createCharacterFrame(textFrame, targetLayer, i, sourceCharacter, characterText);
                    createdFrames.push(lastFrame);
                } catch (e4) { }
            }
        } else {
            for (i = textLength - 1; i >= 0; i--) {
                /* 1文字で失敗しても残りを処理する / One failing character must not stop the rest */
                try {
                    sourceCharacter = textFrame.textRange.characters[i];
                    characterText = sourceCharacter.contents;
                    if (isIgnoredSpaceChar(characterText)) continue;
                    createdFrames.push(createCharacterFrame(textFrame, targetLayer, i, sourceCharacter, characterText));
                } catch (e5) { }
            }
        }

        finishSplit(createdFrames, textFrame);
    }

    /**
     * 控えた文字属性をテキストフレームに書き戻す
     * @param {TextFrame} textFrame - 書き戻し先
     * @param {Object} attributeSnapshot - 文字属性（省略した比率・トラッキングなどは既定値）
     * @returns {void}
     */
    function applyAttributes(textFrame, attributeSnapshot) {
        /* 途中で失敗したらそこまで / Stop at the first failure */
        try {
            var targetAttributes = textFrame.textRange.characterAttributes;
            if (!targetAttributes) return;

            if (attributeSnapshot.textFont) targetAttributes.textFont = attributeSnapshot.textFont;
            if (attributeSnapshot.size) targetAttributes.size = attributeSnapshot.size;
            targetAttributes.horizontalScale = (attributeSnapshot.horizontalScale !== undefined) ? attributeSnapshot.horizontalScale : 100;
            targetAttributes.verticalScale = (attributeSnapshot.verticalScale !== undefined) ? attributeSnapshot.verticalScale : 100;
            targetAttributes.tracking = (attributeSnapshot.tracking !== undefined) ? attributeSnapshot.tracking : 0;
            targetAttributes.baselineShift = (attributeSnapshot.baselineShift !== undefined) ? attributeSnapshot.baselineShift : 0;
            targetAttributes.rotation = (attributeSnapshot.rotation !== undefined) ? attributeSnapshot.rotation : 0;

            if (attributeSnapshot.fillColor) targetAttributes.fillColor = attributeSnapshot.fillColor;
            if (attributeSnapshot.strokeColor) {
                targetAttributes.strokeColor = attributeSnapshot.strokeColor;
                targetAttributes.strokeWeight = attributeSnapshot.strokeWeight;
            }

            if (attributeSnapshot.autoLeading !== undefined) {
                targetAttributes.autoLeading = attributeSnapshot.autoLeading;
                if (!attributeSnapshot.autoLeading && attributeSnapshot.leading) {
                    targetAttributes.leading = attributeSnapshot.leading;
                }
            }

            if (attributeSnapshot.kerningMethod) targetAttributes.kerningMethod = attributeSnapshot.kerningMethod;
        } catch (e) { }
    }

    /**
     * 変形行列から回転角を求める
     * @param {Matrix} sourceMatrix - 変形行列
     * @returns {number} 回転角（度。求められなければ 0）
     */
    function getRotationFromMatrix(sourceMatrix) {
        /* 行列が取れないことがある / The matrix may be unavailable */
        try {
            return Math.atan2(sourceMatrix.mValueB, sourceMatrix.mValueA) * 180 / Math.PI;
        } catch (e) {
            return 0;
        }
    }

    // =========================================
    // エリア内文字への変換 / Area text conversion
    // =========================================

    /**
     * ポイント文字の bounds を枠にしてエリア内文字を作り、内容と主な属性を移す（元は削除）
     * @param {TextFrame} sourceFrame - 変換するテキストフレーム
     * @returns {TextFrame} 作成したエリア内文字（作れなければ sourceFrame）
     */
    function rebuildAsAreaText(sourceFrame) {
        /* 失敗したらそのまま / Keep the original on failure */
        try {
            var frameBounds = sourceFrame.geometricBounds; // [L,T,R,B]
            var L = frameBounds[0], T = frameBounds[1], R = frameBounds[2], B = frameBounds[3];
            var frameWidth = R - L;
            var frameHeight = T - B;
            if (frameWidth <= 0 || frameHeight <= 0) return sourceFrame;

            var sourceLayer = sourceFrame.layer;
            var frameRect = sourceLayer.pathItems.rectangle(T, L, frameWidth, frameHeight);
            frameRect.stroked = false;
            frameRect.filled = false;

            var areaFrame = null;
            try { areaFrame = sourceLayer.textFrames.areaText(frameRect); } catch (e) { areaFrame = null; }
            if (!areaFrame) {
                try { frameRect.remove(); } catch (e2) { }
                return sourceFrame;
            }

            // 内容と主要属性を移植
            try { areaFrame.contents = sourceFrame.contents; } catch (e3) { }
            try { areaFrame.matrix = sourceFrame.matrix; } catch (e4) { }

            // フォント/サイズ/スケール/トラッキング等（可能な範囲）
            var sourceRange = null;
            try { sourceRange = sourceFrame.textRange; } catch (e5) { sourceRange = null; }
            copyCharacterAttributes(areaFrame, sourceRange);

            // 元を削除（frameRect は areaText の枠として保持される）
            try { sourceFrame.remove(); } catch (e6) { }

            return areaFrame;
        } catch (e7) {
            return sourceFrame;
        }
    }

    /**
     * テキストフレームをエリア内文字に変換する（ネイティブAPI → メニューコマンド → bounds から作り直し の順に試す）
     * @param {TextFrame[]} textFrames - 変換するテキストフレーム
     * @returns {TextFrame[]} 変換後のテキストフレーム（変換できなかったものは元のまま）
     */
    function convertTextFramesToAreaText(textFrames) {
        if (!textFrames || textFrames.length === 0) return textFrames;

        // 変換結果のTextFrame配列を返す（失敗時は元を返す）
        var convertedFrames = [];

        // selection を汚さないよう退避
        var doc = null;
        var previousSelection = null;
        try {
            doc = app.activeDocument;
            previousSelection = doc.selection;
        } catch (e) { }

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!textFrame || textFrame.typename !== "TextFrame") {
                continue;
            }

            // すでにエリア内文字ならそのまま
            if (isAreaTextSafely(textFrame)) {
                convertedFrames.push(textFrame);
                continue;
            }

            // ポイントテキストなら、ネイティブAPIでエリア内文字へ変換（最優先）
            try {
                if (textFrame.kind === TextType.POINTTEXT && textFrame.convertPointObjectToAreaObject) {
                    textFrame.convertPointObjectToAreaObject();
                }
            } catch (e2) { }

            // 変換できたらそのまま採用
            if (isAreaTextSafely(textFrame)) {
                convertedFrames.push(textFrame);
                continue;
            }

            // まずはメニューコマンドで変換を試す（最も見た目を保持できる）
            var menuConvertedFrame = null;
            try {
                if (doc) {
                    doc.selection = [textFrame];
                    try { app.executeMenuCommand('ConvertToAreaType'); } catch (e3) {
                        // 環境差のため別名も試す
                        try { app.executeMenuCommand('ConvertToAreaText'); } catch (e4) { }
                    }
                    if (doc.selection && doc.selection.length === 1 && doc.selection[0].typename === "TextFrame") {
                        menuConvertedFrame = doc.selection[0];
                    }
                }
            } catch (e5) {
                menuConvertedFrame = null;
            }

            if (menuConvertedFrame && menuConvertedFrame.typename === "TextFrame") {
                convertedFrames.push(menuConvertedFrame);
                continue;
            }

            // うまくいかなければ簡易フォールバック: bounds を枠にして areaText を作る
            convertedFrames.push(rebuildAsAreaText(textFrame));
        }

        // selection 復元
        try { if (doc) doc.selection = previousSelection; } catch (e6) { }

        return convertedFrames;
    }

    // =========================================
    // エリア内文字のスレッド化 / Area text threading
    // =========================================

    /**
     * エリア内文字を読み順に連結（スレッド化）する
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {TextFrame[]} 読み順に並べたエリア内文字（連結しなかったときは textFrames）
     */
    function threadAreaTextFrames(textFrames) {
        if (!textFrames || textFrames.length < 2) return textFrames;

        var doc = null;
        try { doc = app.activeDocument; } catch (e) { doc = null; }
        if (!doc) return textFrames;

        // 読み順で並べ替え（複数行対応）
        var rows = framesToRowsInReadingOrder(textFrames);

        var orderedFrames = [];
        if (rows.length > 0) {
            for (var r = 0; r < rows.length; r++) {
                for (var i = 0; i < rows[r].length; i++) orderedFrames.push(rows[r][i]);
            }
        } else {
            // フォールバック: 入力順
            for (var j = 0; j < textFrames.length; j++) orderedFrames.push(textFrames[j]);
        }

        // エリア内文字だけに絞る（念のため）
        var areaFrames = [];
        for (var k = 0; k < orderedFrames.length; k++) {
            var orderedFrame = orderedFrames[k];
            if (!orderedFrame || orderedFrame.typename !== "TextFrame") continue;
            if (isAreaTextSafely(orderedFrame)) areaFrames.push(orderedFrame);
        }
        if (areaFrames.length < 2) return textFrames;

        // selection を退避して threadTextCreate を実行
        var previousSelection = null;
        try { previousSelection = doc.selection; } catch (e2) { previousSelection = null; }

        try {
            doc.selection = areaFrames;
            // スレッド（連結）
            app.executeMenuCommand('threadTextCreate');
        } catch (e3) {
            // 失敗しても無視
        }

        try { doc.selection = previousSelection; } catch (e4) { }

        // 連結後も参照はそのまま使えるため、並び順（areaFrames）を返す
        return areaFrames;
    }

    // =========================================
    // グループ化 / Grouping
    // =========================================

    /**
     * 設定に応じて、分けたテキストを行ごと・全体でグループにする
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} groupMode - "none" / "line" / "all"
     * @returns {void}
     */
    function applyGroupingByMode(textFrames, groupMode) {
        if (!textFrames || textFrames.length < 2) return;
        if (!groupMode || groupMode === "none") return;

        // 読み順で並べ替えた行リストを作る
        var rows = framesToRowsInReadingOrder(textFrames);
        if (rows.length === 0) return;

        if (groupMode === "all") {
            // 全体を1グループ
            var allFrames = [];
            for (var r = 0; r < rows.length; r++) {
                for (var i = 0; i < rows[r].length; i++) allFrames.push(rows[r][i]);
            }
            groupItems(allFrames);
            return;
        }

        if (groupMode === "line") {
            // 各行ごとにグループ
            for (var rr = 0; rr < rows.length; rr++) {
                if (rows[rr].length >= 2) groupItems(rows[rr]);
            }
        }
    }

    /**
     * オブジェクトを同じ階層に作ったグループへ、並び順のまま移す
     * @param {PageItem[]} pageItems - グループにするオブジェクト
     * @returns {GroupItem|null} 作成したグループ（作れなければ null）
     */
    function groupItems(pageItems) {
        if (!pageItems || pageItems.length < 2) return null;

        var parentContainer = null;
        try { parentContainer = pageItems[0].parent; } catch (e) { parentContainer = null; }
        if (!parentContainer) return null;

        var newGroup = null;
        try {
            // 同階層にグループを作る
            if (parentContainer.typename === "Layer") newGroup = parentContainer.groupItems.add();
            else if (parentContainer.groupItems) newGroup = parentContainer.groupItems.add();
            else newGroup = pageItems[0].layer.groupItems.add();
        } catch (e2) {
            try { newGroup = pageItems[0].layer.groupItems.add(); } catch (e3) { newGroup = null; }
        }
        if (!newGroup) return null;

        // 読み順のままグループへ移動
        for (var i = 0; i < pageItems.length; i++) {
            try { pageItems[i].move(newGroup, ElementPlacement.PLACEATEND); } catch (e4) { }
        }

        return newGroup;
    }

    showSplitDialog();

})();
