#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字のうち、あふれているものだけ自動サイズ調整をONにして解消してから、ポイント文字に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md

### Overview

Turns on auto-sizing for whichever of the selected area texts are overset, resolves the overflow, and then converts them to point text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitAndConvertToPointType";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 「強制改行を削除」の初期状態 / Initial state of "Remove forced line breaks" */
    var DEFAULT_REMOVE_LINE_BREAKS = false;

    /* 強制改行（ソフトリターン）と見なす文字。段落改行（\r）は残す
       Characters treated as forced line breaks (soft returns); paragraph returns (\r) are kept */
    var FORCED_BREAK_PATTERN = /[\u0003\n\u2028]/;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;            /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;            /* ウィンドウ内の要素間隔 / window spacing */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0]; /* ボタンバーの余白 / margins of the button bar */
    var BUTTON_BAR_SPACING = 10;            /* ボタンバー内の要素間隔 / spacing inside the button bar */

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

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "ポイント文字に変換", en: "Convert to Point Type" }
        },
        checkbox: {
            removeLineBreaks: { ja: "強制改行を削除", en: "Remove forced line breaks" }
        },
        tooltip: {
            removeLineBreaks: {
                ja: "変換後に残る強制改行（ソフトリターン）を取り除き、行をつなげます。段落改行は残ります。",
                en: "Strips the forced line breaks (soft returns) left after the conversion, joining the lines. Paragraph returns are kept."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectAreaText: { ja: "エリア内文字を選択してください。", en: "Please select area type." },
            noTarget: {
                ja: "変換できるエリア内文字がありませんでした。ロックや非表示になっていないか確認してください。",
                en: "No area type could be converted. Check whether the frames are locked or hidden."
            },
            notSupported: {
                ja: "お使いのIllustratorはポイント文字への変換に対応していません。",
                en: "This version of Illustrator cannot convert area type to point text."
            },
            partialFailure: {
                ja: "{done}件を変換しました。{failed}件は変換できませんでした（連結されたテキストなどは対象外です）。",
                en: "Converted {done}. {failed} could not be converted (threaded text and the like are not supported)."
            }
        }
    };

    /**
     * "category.key" 形式のキーからラベルを取得する
     * @param {string} labelPath - ラベルキー（例: "alert.noDocument"）
     * @returns {string} 現在の言語のラベル文字列（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[labelPathKeys[i]];
        }
        if (!labelNode) return labelPath;
        if (typeof labelNode[uiLang] === "string") return labelNode[uiLang];
        return (typeof labelNode.en === "string") ? labelNode.en : labelPath;
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   自動サイズ調整はDOMから設定できないため、アクション経由で切り替える
    //   Auto-sizing cannot be set from the DOM, so it is toggled through an action
    // =========================================

    /**
     * 文字列を ASCII 16進に変換する
     * @param {string} text - 変換する文字列
     * @returns {string} 16進文字列
     */
    function asciiToHex(text) {
        var hexText = "";
        for (var i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) hexPair = "0" + hexPair;
            hexText += hexPair;
        }
        return hexText;
    }

    /**
     * アクション名ブロック /name [ <len> <hex> ] を生成する
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(actionName) {
        return "/name [ " + actionName.length + " " + asciiToHex(actionName).toUpperCase() + " ]";
    }

    /**
     * 自動サイズ調整アクションセットの定義（.aia 文字列）を組み立てる
     * @param {string} setName - アクションセット名
     * @returns {string} .aia 形式のアクションセット定義
     */
    function buildAutoSizeActionSetAia(setName) {
        return "/version 3" +
            buildActionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount 1" +
            "/action-1 {" +
            " " + buildActionNameBlock("AutoSizeOn") +
            " /keyIndex 0" +
            " /colorIndex 0" +
            " /isOpen 1" +
            " /eventCount 1" +
            " /event-1 {" +
            " /useRulersIn1stQuadrant 0" +
            " /internalName (adobe_SLOAreaTextDialog)" +
            " /localizedName [ 33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3 ]" +
            " /isOpen 0" +
            " /isOn 1" +
            " /hasDialog 0" +
            " /parameterCount 1" +
            " /parameter-1 {" +
            " /key 1952539754" +
            " /showInPalette 4294967295" +
            " /type (integer)" +
            " /value 1" +
            " }" +
            " }" +
            "}";
    }

    var ACTION_SET_AUTO_SIZE = SCRIPT_NAME + "_AutoSize";

    /**
     * 自動サイズ調整アクションを一時ファイル経由で読み込む（スクリプト開始時に1回。既存があれば先に外す）
     * @returns {void}
     */
    function loadAutoSizeAction() {
        unloadAutoSizeAction();
        var tempFile = new File(Folder.temp + "/" + SCRIPT_NAME + "_action.aia");
        tempFile.open("w");
        tempFile.write(buildAutoSizeActionSetAia(ACTION_SET_AUTO_SIZE));
        tempFile.close();
        app.loadAction(tempFile);
        try { tempFile.remove(); } catch (e) { /* 一時ファイルの削除失敗は無視 / Ignore a failed temp cleanup */ }
    }

    /**
     * 読み込んだアクションセットを破棄する（スクリプト終了時）
     * @returns {void}
     */
    function unloadAutoSizeAction() {
        /* 読み込まれていなければ例外になる / Throws when the set is not loaded */
        try { app.unloadAction(ACTION_SET_AUTO_SIZE, ""); } catch (e) { }
    }

    /**
     * 選択中のエリア内文字に自動サイズ調整をかける（アクションは選択に効く）
     * @returns {void}
     */
    function runAutoSizeAction() {
        try { app.doScript("AutoSizeOn", ACTION_SET_AUTO_SIZE, false); } catch (e) { }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * グループを横並びの行にする（揃えは横と天地を対で指定する）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [horizontalAlign] - 行の横方向の揃え（既定は "left"）
     * @param {number} [spacing] - 要素間隔（既定は WINDOW_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * オプションダイアログを表示する
     * @returns {{removeLineBreaks: boolean}|null} 設定（キャンセル時は null）
     */
    function showOptionsDialog() {
        var optionsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        optionsDialog.orientation = "column";
        optionsDialog.alignChildren = ["fill", "top"];
        optionsDialog.spacing = WINDOW_SPACING;
        optionsDialog.margins = WINDOW_MARGINS;

        var optionRow = optionsDialog.add("group");
        setupRow(optionRow);
        var removeLineBreaksCheckbox = optionRow.add("checkbox", undefined, getLabel("checkbox.removeLineBreaks"));
        removeLineBreaksCheckbox.value = DEFAULT_REMOVE_LINE_BREAKS;
        removeLineBreaksCheckbox.helpTip = getLabel("tooltip.removeLineBreaks");

        /* ボタンバー（Mac 規約で キャンセル → OK）/ Button bar (Cancel → OK per macOS) */
        var btnRowGroup = optionsDialog.add("group");
        setupRow(btnRowGroup, "right", BUTTON_BAR_SPACING);
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        if (optionsDialog.show() !== 1) return null;
        return { removeLineBreaks: removeLineBreaksCheckbox.value };
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /**
     * エリア内文字があふれているかを返す
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @returns {boolean} あふれていれば true
     */
    function isFrameOverset(areaTextFrame) {
        /* overflows を持たない環境がある / Some versions lack overflows */
        try {
            if (typeof areaTextFrame.overflows !== "undefined") return !!areaTextFrame.overflows;
        } catch (e) { }
        // overflows が読めない環境では、行に入っている文字数と全文字数を突き合わせる
        // Where overflows cannot be read, the characters inside the lines are counted against the total
        try {
            var visibleCount = 0;
            for (var i = 0; i < areaTextFrame.lines.length; i++) {
                visibleCount += areaTextFrame.lines[i].characters.length;
            }
            return visibleCount < areaTextFrame.characters.length;
        } catch (e0) { return false; }
    }

    /**
     * 変換できるエリア内文字かを返す（ロック・非表示のものは対象外）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 変換できれば true
     */
    function isConvertibleAreaTextFrame(pageItem) {
        try {
            if (!pageItem || pageItem.typename !== "TextFrame" || pageItem.kind !== TextType.AREATEXT) return false;
            if (pageItem.locked || pageItem.hidden) return false;
            var ownerLayer = pageItem.layer;
            if (ownerLayer && (ownerLayer.locked || !ownerLayer.visible)) return false;
        } catch (e) { return false; }
        return true;
    }

    /**
     * 強制改行を1文字ずつ取り除く（contents の一括置換は文字ごとの書式を失うため使わない）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function removeForcedLineBreaks(textFrame) {
        var textCharacters;
        try { textCharacters = textFrame.characters; } catch (e) { return; }
        // 削除するとインデックスがずれるので後ろから処理する / Deleting shifts the indices, so walk from the back
        for (var i = textCharacters.length - 1; i >= 0; i--) {
            /* 読めない文字・消せない文字は飛ばす / Skip characters that cannot be read or removed */
            try {
                if (FORCED_BREAK_PATTERN.test(textCharacters[i].contents)) textCharacters[i].remove();
            } catch (eChar) { }
        }
    }

    /**
     * 選択からエリア内文字だけを拾う
     * @param {Array} selectedItems - ドキュメントの選択内容
     * @returns {TextFrame[]} 変換できるエリア内文字
     */
    function collectAreaTextFrames(selectedItems) {
        var areaTextFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (isConvertibleAreaTextFrame(selectedItems[i])) { areaTextFrames.push(selectedItems[i]); }
        }
        return areaTextFrames;
    }

    /**
     * あふれているフレームだけ自動サイズ調整で解消してから、ポイント文字へ変換する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} areaTextFrames - 変換するエリア内文字
     * @param {boolean} removeLineBreaks - 変換後に強制改行を取り除くか
     * @returns {{converted: TextFrame[], failedCount: number}} 変換できたポイント文字と、変換できなかった数
     */
    function convertAreaTextFramesToPointText(doc, areaTextFrames, removeLineBreaks) {
        /* あふれていないフレームに自動サイズ調整をかけると、テキストの配置が中央・下のときに文字が動くため
           Auto-sizing a frame with no overset would move the text when it is centred or bottom-aligned
           convertAreaObjectToPointObject() は「その場変換・戻り値 null」で、変換後は元の参照が stale になり
           kind を AREATEXT のまま報告することがある。そこで変換前に目印の名前を付け、変換後に
           doc.textFrames を1回だけ走査して回収する
           (the API converts in place and returns null, and the old wrappers can go stale and still report
            AREATEXT, so the frames are tagged with a marker name first and collected afterwards
            by a single fresh scan of doc.textFrames) */
        var markerPrefix = "__" + SCRIPT_NAME + "_marker_";
        var previousNames = [];

        for (var i = 0; i < areaTextFrames.length; i++) {
            var previousName = "";
            try { previousName = areaTextFrames[i].name; } catch (eRead) { }
            previousNames.push(previousName);
            try { areaTextFrames[i].name = markerPrefix + i; } catch (eTag) { }
        }

        // 変換するとオブジェクトが置き換わるので後ろから処理する
        // Each frame is replaced as it goes, so the list is walked from the back
        for (var j = areaTextFrames.length - 1; j >= 0; j--) {
            // 1フレームで失敗しても残りを処理できるようにする
            // One failing frame must not stop the rest
            try {
                if (isFrameOverset(areaTextFrames[j])) {
                    doc.selection = null;
                    areaTextFrames[j].selected = true;
                    app.redraw(); // Illustratorに選択状態を確定させる / Let Illustrator settle the selection
                    runAutoSizeAction();
                }
                areaTextFrames[j].convertAreaObjectToPointObject();
            } catch (e) { }
        }

        // 目印で回収して名前を戻す（走査は1回だけ）/ Collect by marker and restore the names (a single scan)
        var convertedFrames = [], failedCount = 0;
        var allTextFrames = doc.textFrames;
        for (var k = 0; k < allTextFrames.length; k++) {
            var frameName = "";
            try { frameName = allTextFrames[k].name; } catch (eName) { continue; }
            if (frameName.indexOf(markerPrefix) !== 0) continue;

            var markerIndex = parseInt(frameName.substring(markerPrefix.length), 10);
            var restoredName = (!isNaN(markerIndex) && typeof previousNames[markerIndex] === "string")
                ? previousNames[markerIndex] : "";
            try { allTextFrames[k].name = restoredName; } catch (eRestore) { }
            if (allTextFrames[k].kind === TextType.POINTTEXT) {
                // 変換すると折り返し位置に強制改行が残るので、ポイント文字になってから取り除く
                // The conversion leaves a forced break at every wrap, so they are stripped once it is point text
                if (removeLineBreaks) removeForcedLineBreaks(allTextFrames[k]);
                convertedFrames.push(allTextFrames[k]);
            } else { failedCount++; }
        }
        return { converted: convertedFrames, failedCount: failedCount };
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    /**
     * 選択を確かめて変換する（ダイナミックアクションは終了時に必ずアンロードする）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        /* 文字ツールで文字を選択中は TextRange が返り、length が文字数になるため配列かどうかで判定する
           With the type tool the selection is a TextRange whose length counts characters */
        if (!(selectedItems instanceof Array) || selectedItems.length === 0) {
            alert(getLabel("alert.selectAreaText"));
            return;
        }

        var areaTextFrames = collectAreaTextFrames(selectedItems);
        if (!areaTextFrames.length) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        // 古いIllustratorには変換APIが無いので、何も触らずに知らせる
        // Older Illustrator has no conversion API, so nothing is touched
        var supportsConversion = false;
        try { supportsConversion = !!areaTextFrames[0].convertAreaObjectToPointObject; } catch (eApi) { }
        if (!supportsConversion) {
            alert(getLabel("alert.notSupported"));
            return;
        }

        var dialogOptions = showOptionsDialog();
        if (!dialogOptions) return;

        loadAutoSizeAction();
        var conversionResult;
        try {
            conversionResult = convertAreaTextFramesToPointText(doc, areaTextFrames, dialogOptions.removeLineBreaks);
        } finally {
            unloadAutoSizeAction();
        }

        // 変換後のポイント文字を選び直す（削除済みの参照を選択に残さない）
        // Re-select the resulting point text, so no stale reference lingers in the selection
        try { doc.selection = conversionResult.converted.length ? conversionResult.converted : null; } catch (eSelect) { }
        app.redraw();

        if (!conversionResult.converted.length) {
            alert(getLabel("alert.noTarget"));
        } else if (conversionResult.failedCount > 0) {
            alert(getLabel("alert.partialFailure")
                .replace("{done}", conversionResult.converted.length)
                .replace("{failed}", conversionResult.failedCount));
        }
    }

    main();

})();
