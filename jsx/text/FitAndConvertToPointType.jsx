#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字のうち、あふれているものだけ自動サイズ調整をONにして解消してから、ポイント文字に変換します。

詳細はREADMEを参照。

*/

/*

### Overview

Clears the overset of the selected Area Type with auto-sizing, then converts every frame to point text.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitAndConvertToPointType";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-20";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md

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

    /* 現在の言語（ja / en）/ Current language (ja / en) */
    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

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
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            selectAreaText: {
                ja: "エリア内文字を選択してください。",
                en: "Please select area type."
            },
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

    /* ラベル取得（"category.key" 形式）/ Resolve a "category.key" label */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[keyParts[i]];
        }
        if (!labelNode) return key;
        if (typeof labelNode[currentLanguage] === "string") return labelNode[currentLanguage];
        return (typeof labelNode.en === "string") ? labelNode.en : key;
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   自動サイズ調整はDOMから設定できないため、アクション経由で切り替える
    //   Auto-sizing cannot be set from the DOM, so it is toggled through an action
    // =========================================

    /* 文字列を ASCII 16進に変換 / Convert a string to ASCII hex */
    function asciiToHex(text) {
        var hex = "";
        for (var i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) hexPair = "0" + hexPair;
            hex += hexPair;
        }
        return hex;
    }

    /* アクション名ブロック /name [ <len> <hex> ] を生成 / Build the /name [ <len> <hex> ] block */
    function buildActionNameBlock(name) {
        return "/name [ " + name.length + " " + asciiToHex(name).toUpperCase() + " ]";
    }

    /* 自動サイズ調整アクションセットの定義（.aia 文字列）を組み立てる
       Build the action set definition (.aia string) for auto-sizing */
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

    /* アクションセット名。ユーザーのアクションを消さないよう、スクリプト名を冠して衝突を避ける
       Action set name, prefixed with the script name so a user's own sets are never unloaded */
    var ACTION_SET_AUTO_SIZE = SCRIPT_NAME + "_AutoSize";

    /* 自動サイズ調整アクションを読み込む（スクリプト開始時に1回）/ Load the auto-size action set (once at startup) */
    function loadAutoSizeAction() {
        try { app.unloadAction(ACTION_SET_AUTO_SIZE, ""); } catch (e0) { }
        var tempFile = new File(Folder.temp + "/" + SCRIPT_NAME + "_action.aia");
        tempFile.open("w");
        tempFile.write(buildAutoSizeActionSetAia(ACTION_SET_AUTO_SIZE));
        tempFile.close();
        app.loadAction(tempFile);
        try { tempFile.remove(); } catch (e1) { }
    }

    /* 読み込んだアクションセットを破棄する（スクリプト終了時）/ Discard the loaded action set (on exit) */
    function unloadAutoSizeAction() {
        try { app.unloadAction(ACTION_SET_AUTO_SIZE, ""); } catch (e) { }
    }

    /* 選択中のエリア内文字に自動サイズ調整をかける（アクションは選択に効く）
       Turn auto-sizing on for the current selection (the action works on the selection) */
    function runAutoSizeAction() {
        try { app.doScript("AutoSizeOn", ACTION_SET_AUTO_SIZE, false); } catch (e) { }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* グループを横並びの行にする（揃えは横と天地を対で指定する）
       Lay a group out as a row (both axes are always paired) */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* オプションダイアログ。キャンセルなら null を返す / Options dialog; returns null when cancelled */
    function showOptionsDialog() {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = WINDOW_SPACING;
        dialog.margins = WINDOW_MARGINS;

        var optionRow = dialog.add("group");
        setupRow(optionRow);
        var removeLineBreaksCheckbox = optionRow.add("checkbox", undefined, getLabel("checkbox.removeLineBreaks"));
        removeLineBreaksCheckbox.value = DEFAULT_REMOVE_LINE_BREAKS;
        removeLineBreaksCheckbox.helpTip = getLabel("tooltip.removeLineBreaks");

        /* ボタンバー（Mac 規約で キャンセル → OK）/ Button bar (Cancel → OK per macOS) */
        var buttonBarGroup = dialog.add("group");
        setupRow(buttonBarGroup, "right", BUTTON_BAR_SPACING);
        buttonBarGroup.margins = BUTTON_BAR_MARGINS;
        buttonBarGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonBarGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        if (dialog.show() !== 1) return null;
        return { removeLineBreaks: removeLineBreaksCheckbox.value };
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /* エリア内文字があふれているか / Whether the Area Type is overset */
    function isFrameOverset(areaTextFrame) {
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

    /* 変換できるエリア内文字か（ロック・非表示のものは対象外）
       Whether the frame can be converted (locked and hidden ones are skipped) */
    function isConvertibleAreaTextFrame(item) {
        try {
            if (!item || item.typename !== "TextFrame" || item.kind !== TextType.AREATEXT) return false;
            if (item.locked || item.hidden) return false;
            var ownerLayer = item.layer;
            if (ownerLayer && (ownerLayer.locked || !ownerLayer.visible)) return false;
        } catch (e) { return false; }
        return true;
    }

    /* 強制改行を1文字ずつ取り除く（contents の一括置換は文字ごとの書式を失うため使わない）
       Remove the forced line breaks one character at a time
       (rewriting contents in one go would flatten the per-character formatting) */
    function removeForcedLineBreaks(textFrame) {
        var textCharacters;
        try { textCharacters = textFrame.characters; } catch (e) { return; }
        // 削除するとインデックスがずれるので後ろから処理する / Deleting shifts the indices, so walk from the back
        for (var i = textCharacters.length - 1; i >= 0; i--) {
            var characterText = "";
            try { characterText = textCharacters[i].contents; } catch (eRead) { continue; }
            if (!FORCED_BREAK_PATTERN.test(characterText)) continue;
            try { textCharacters[i].remove(); } catch (eRemove) { }
        }
    }

    /* 選択からエリア内文字だけを拾う / Pick just the Area Type frames out of the selection */
    function collectAreaTextFrames(selection) {
        var areaTextFrames = [];
        for (var i = 0; i < selection.length; i++) {
            if (isConvertibleAreaTextFrame(selection[i])) { areaTextFrames.push(selection[i]); }
        }
        return areaTextFrames;
    }

    /* あふれているフレームだけ自動サイズ調整で解消してから、ポイント文字へ変換する
       Clear the overset with auto-sizing where there is one, then convert to point text
       あふれていないフレームに自動サイズ調整をかけると、テキストの配置が中央・下のときに文字が動くため
       Auto-sizing a frame with no overset would move the text when it is centred or bottom-aligned
       convertAreaObjectToPointObject() は「その場変換・戻り値 null」で、変換後は元の参照が stale になり
       kind を AREATEXT のまま報告することがある。そこで変換前に目印の名前を付け、変換後に
       doc.textFrames を1回だけ走査して回収する
       (the API converts in place and returns null, and the old wrappers can go stale and still report
        AREATEXT, so the frames are tagged with a marker name first and collected afterwards
        by a single fresh scan of doc.textFrames) */
    function convertAreaTextFramesToPointText(doc, areaTextFrames, removeLineBreaks) {
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
        var textFrames = doc.textFrames;
        for (var k = 0; k < textFrames.length; k++) {
            var frameName = "";
            try { frameName = textFrames[k].name; } catch (eName) { continue; }
            if (frameName.indexOf(markerPrefix) !== 0) continue;

            var markerIndex = parseInt(frameName.substring(markerPrefix.length), 10);
            var restoredName = (!isNaN(markerIndex) && typeof previousNames[markerIndex] === "string")
                ? previousNames[markerIndex] : "";
            try { textFrames[k].name = restoredName; } catch (eRestore) { }
            if (textFrames[k].kind === TextType.POINTTEXT) {
                // 変換すると折り返し位置に強制改行が残るので、ポイント文字になってから取り除く
                // The conversion leaves a forced break at every wrap, so they are stripped once it is point text
                if (removeLineBreaks) removeForcedLineBreaks(textFrames[k]);
                convertedFrames.push(textFrames[k]);
            } else { failedCount++; }
        }
        return { converted: convertedFrames, failedCount: failedCount };
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    /* 選択を確かめて変換する（ダイナミックアクションは終了時に必ずアンロードする）
       Check the selection and convert (the dynamic action is always unloaded on exit) */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selection = doc.selection;

        /* 文字ツールで文字を選択中は TextRange が返り、length が文字数になるため配列かどうかで判定する
           With the type tool the selection is a TextRange whose length counts characters */
        if (!(selection instanceof Array) || selection.length === 0) {
            alert(getLabel("alert.selectAreaText"));
            return;
        }

        var areaTextFrames = collectAreaTextFrames(selection);
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
        var result;
        try {
            result = convertAreaTextFramesToPointText(doc, areaTextFrames, dialogOptions.removeLineBreaks);
        } finally {
            unloadAutoSizeAction();
        }

        // 変換後のポイント文字を選び直す（削除済みの参照を選択に残さない）
        // Re-select the resulting point text, so no stale reference lingers in the selection
        try { doc.selection = result.converted.length ? result.converted : null; } catch (eSelect) { }
        app.redraw();

        if (!result.converted.length) {
            alert(getLabel("alert.noTarget"));
        } else if (result.failedCount > 0) {
            alert(getLabel("alert.partialFailure")
                .replace("{done}", result.converted.length)
                .replace("{failed}", result.failedCount));
        }
    }

    main();

})();
