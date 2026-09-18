#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字の垂直方向の配置と行揃えを、まとめて中央にそろえます。閉じたパスを選択している場合はエリア内文字に変換し、サンプルテキスト、または一緒に選択した／グループ化したテキストの内容を流し込みます（グループは流し込み後に解除）。

詳細は README を参照してください。

### Overview

Sets both the vertical alignment and the justification of the selected Area Type frames to center in one pass. Selected closed paths are converted to Area Type and filled with sample text, or with the contents of a text object selected alongside or grouped with the path (such a group is released once the text is poured).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AreaTypeCenterMiddle";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeCenterMiddle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AreaTypeCenterMiddle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 選択したパスに流し込むサンプルテキスト / Sample text poured into the selected path */
    var DUMMY_TEXT_JA = "山路を登りながら";
    var DUMMY_TEXT_EN = "Typography";

    /* サンプルテキストの優先フォント候補とサイズ / Preferred fonts and size for the sample text */
    var DUMMY_FONT_JA = ["HiraginoSans-W3", "Hiragino Sans W3"];
    var DUMMY_FONT_EN = ["MyriadPro-Regular", "Myriad Pro Regular", "MyriadPro", "Myriad"];
    var DUMMY_FONT_SIZE = 10;
    // ============================================================

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の言語（ja / en）を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLanguage();

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTarget: { ja: "エリア内文字またはパスを選択してください。", en: "Please select area text or a path." }
        }
    };

    /**
     * "category.key" 形式のキーからラベルを取得する
     * @param {string} key - ラベルキー（例: "alert.noDocument"）
     * @returns {string} 現在の言語のラベル文字列
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[keyParts[i]];
        }
        if (labelNode) {
            if (typeof labelNode[currentLanguage] === "string") return labelNode[currentLanguage];
            if (typeof labelNode.en === "string") return labelNode.en;
        }
        return key;
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   垂直方向の配置はDOMから設定できないため、
    //   スクリプト実行時にアクションを読み込み、終了時にアンロードする
    //   Vertical alignment has no DOM API, so an action is loaded at
    //   startup and unloaded on exit
    // =========================================

    var ACTION_SET_ALIGNMENT = "AreaTypeCenterMiddle_Alignment";
    var ACTION_ALIGN_CENTER = "AlignCenter";

    /**
     * 文字列をASCII16進に変換する
     * @param {string} text - 変換する文字列
     * @returns {string} 16進文字列
     */
    function asciiToHex(text) {
        var hex = "";
        for (var i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) hexPair = "0" + hexPair;
            hex += hexPair;
        }
        return hex;
    }

    /**
     * アクション名ブロック /name [ <len> <hex> ] を生成する
     * @param {string} name - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(name) {
        return "/name [ " + name.length + " " + asciiToHex(name).toUpperCase() + " ]";
    }

    /**
     * 垂直方向の配置（中央）アクションセット定義（.aia文字列）を組み立てる
     * @returns {string} .aia形式のアクションセット定義
     */
    function buildAlignCenterAia() {
        return "/version 3" +
            buildActionNameBlock(ACTION_SET_ALIGNMENT) +
            "/isOpen 1" +
            "/actionCount 1" +
            "/action-1 {" +
            " " + buildActionNameBlock(ACTION_ALIGN_CENTER) +
            " /keyIndex 0" +
            " /colorIndex 0" +
            " /isOpen 1" +
            " /eventCount 1" +
            " /event-1 {" +
            " /useRulersIn1stQuadrant 0" +
            " /internalName (adobe_frameAlignment)" +
            " /localizedName [ 39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897 ]" +
            " /isOpen 0" +
            " /isOn 1" +
            " /hasDialog 0" +
            " /parameterCount 1" +
            " /parameter-1 {" +
            " /key 1717660782" +
            " /showInPalette 4294967295" +
            " /type (integer)" +
            " /value 1" +
            " }" +
            " }" +
            "}";
    }

    /**
     * アクションセットを一時ファイル経由で読み込む（既存があれば先に外す）
     * @returns {void}
     */
    function loadAlignmentAction() {
        try { app.unloadAction(ACTION_SET_ALIGNMENT, ""); } catch (e0) { }
        var tempFile = new File(Folder.temp + "/" + ACTION_SET_ALIGNMENT + ".aia");
        tempFile.open("w");
        tempFile.write(buildAlignCenterAia());
        tempFile.close();
        app.loadAction(tempFile);
        tempFile.remove();
    }

    /**
     * 読み込んだアクションセットを破棄する（スクリプト終了時）
     * @returns {void}
     */
    function unloadAlignmentAction() {
        try { app.unloadAction(ACTION_SET_ALIGNMENT, ""); } catch (e) { }
    }

    // =========================================
    // テキストの流し込み / Text pouring
    // =========================================

    /**
     * 候補名から利用できるフォントを返す
     * @param {Array<string>} candidateFontNames - フォント名の候補
     * @returns {TextFont} 見つかったフォント（無ければnull）
     */
    function findAvailableTextFont(candidateFontNames) {
        for (var i = 0; i < candidateFontNames.length; i++) {
            try {
                var font = app.textFonts.getByName(candidateFontNames[i]);
                if (font) return font;
            } catch (e) { }
        }
        return null;
    }

    /**
     * 閉じたパスを取り出す（複合パスは先頭のパスを見る）
     * @param {object} item - 選択オブジェクト
     * @returns {PathItem} 閉じたパス（無ければnull）
     */
    function getClosedPathItem(item) {
        if (item.typename === "PathItem") return item.closed ? item : null;
        if (item.typename === "CompoundPathItem" && item.pathItems.length > 0) {
            var firstPath = item.pathItems[0];
            return firstPath.closed ? firstPath : null;
        }
        return null;
    }

    /**
     * 流し込んだテキストに書式を設定する
     * @param {TextFrame} areaFrame - 設定先のエリア内文字
     * @param {TextFrame} sourceFrame - 書式の引き継ぎ元（nullならサンプルテキスト用の書式）
     * @param {TextFont} sampleFont - サンプルテキストのフォント（無ければnull）
     * @returns {void}
     */
    function applyTextStyle(areaFrame, sourceFrame, sampleFont) {
        var targetAttrs = areaFrame.textRange.characterAttributes;
        /* 未インストールのフォントやテキストに使えない色は適用に失敗しうる / An uninstalled font or an unusable color can fail to apply */
        try {
            if (sourceFrame) {
                var sourceAttrs = sourceFrame.textRange.characterAttributes;
                targetAttrs.size = sourceAttrs.size;
                targetAttrs.textFont = sourceAttrs.textFont;
                targetAttrs.fillColor = sourceAttrs.fillColor;
            } else {
                targetAttrs.size = DUMMY_FONT_SIZE;
                if (sampleFont) targetAttrs.textFont = sampleFont;
            }
        } catch (e) { }
    }

    /**
     * 閉じたパスをエリア内文字に変換してテキストを流し込む
     * @param {Document} doc - 対象ドキュメント
     * @param {object} shapeItem - 変換するパス（複合パスも可）
     * @param {string} bodyText - 流し込むテキスト
     * @returns {TextFrame} 作成したエリア内文字（変換できなければnull）
     */
    function convertShapeToAreaText(doc, shapeItem, bodyText) {
        var closedPath = getClosedPathItem(shapeItem);
        if (!closedPath) return null;
        var areaFrame = null;
        /* 種類によってはエリア内文字にできない / Some shapes cannot become Area Type */
        try {
            closedPath.filled = false;
            closedPath.stroked = false;
            areaFrame = doc.textFrames.areaText(closedPath);
            areaFrame.contents = bodyText;
        } catch (e) {
            return null;
        }
        /* 複合パスは先頭のパスだけを枠にするので、空になった殻を残さない / Drop the compound shell left empty */
        if (shapeItem.typename === "CompoundPathItem" && shapeItem.pathItems.length === 0) shapeItem.remove();
        return areaFrame;
    }

    /**
     * 組み合わせごとに、閉じたパスをエリア内文字にしてテキストを流し込む
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} fillJobs - shapeItem（変換するパス）と sourceTextFrame（流し込み元。nullならサンプルテキスト）の組み合わせ
     * @returns {Array<TextFrame>} 作成したエリア内文字
     */
    function fillShapesWithText(doc, fillJobs) {
        var sampleText = (currentLanguage === "ja") ? DUMMY_TEXT_JA : DUMMY_TEXT_EN;
        var sampleFont = findAvailableTextFont((currentLanguage === "ja") ? DUMMY_FONT_JA : DUMMY_FONT_EN);
        var createdFrames = [];

        for (var i = 0; i < fillJobs.length; i++) {
            var sourceTextFrame = fillJobs[i].sourceTextFrame;
            var areaFrame = convertShapeToAreaText(doc, fillJobs[i].shapeItem, sourceTextFrame ? sourceTextFrame.contents : sampleText);
            if (!areaFrame) continue;
            applyTextStyle(areaFrame, sourceTextFrame, sampleFont);
            createdFrames.push(areaFrame);
            if (!sourceTextFrame) continue;
            var pairGroup = (sourceTextFrame.parent.typename === "GroupItem") ? sourceTextFrame.parent : null;
            /* 流し込みが済んだ元のテキストは残さない / Remove the source text once it has been poured */
            sourceTextFrame.remove();
            if (!pairGroup) continue;
            /* 組み合わせのグループは、エリア内文字を外に出して解除する / Ungroup the pair by moving the frame out and dropping the empty group */
            try {
                areaFrame.move(pairGroup, ElementPlacement.PLACEBEFORE);
                if (pairGroup.pageItems.length === 0) pairGroup.remove();
            } catch (e) { }
        }
        return createdFrames;
    }

    // =========================================
    // 適用 / Apply
    // =========================================

    /**
     * グループが「閉じたパス1つ＋テキスト1つ」なら、その組み合わせを返す
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {object} shapeItem / textFrame を持つオブジェクト（該当しなければnull）
     */
    function getShapeTextPair(groupItem) {
        /* クリップグループの枠はマスクなので対象にしない / The frame of a clipping group is a mask, not a shape to convert */
        if (groupItem.clipped) return null;
        var members = groupItem.pageItems;
        if (members.length !== 2) return null;
        var pair = { shapeItem: null, textFrame: null };
        for (var i = 0; i < members.length; i++) {
            if (members[i].typename === "TextFrame") pair.textFrame = members[i];
            else if (getClosedPathItem(members[i])) pair.shapeItem = members[i];
        }
        return (pair.shapeItem && pair.textFrame) ? pair : null;
    }

    /**
     * グループから「閉じたパス1つ＋テキスト1つ」の組み合わせを集める
     * @param {GroupItem} groupItem - 対象のグループ
     * @param {Array} pairs - 集めた組み合わせの入れ物
     * @returns {void}
     */
    function collectShapeTextPairs(groupItem, pairs) {
        var pair = getShapeTextPair(groupItem);
        if (pair) {
            pairs.push(pair);
            return;
        }
        /* 該当しないグループは、入れ子になっているグループを見る / Look into nested groups when the group itself is not a pair */
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            if (groupItem.pageItems[i].typename === "GroupItem") collectShapeTextPairs(groupItem.pageItems[i], pairs);
        }
    }

    /**
     * 選択オブジェクトを、エリア内文字・閉じたパス・それ以外のテキスト・グループの組み合わせに仕分ける
     * @param {Array} selection - ドキュメントの選択内容
     * @returns {object} areaTextFrames / shapeItems / otherTextFrames / shapePairs を持つオブジェクト
     */
    function classifySelection(selection) {
        var picked = { areaTextFrames: [], shapeItems: [], otherTextFrames: [], shapePairs: [] };
        /* 文字編集中は選択がTextRangeになり、ページアイテムが取り出せない / While editing text the selection is a TextRange, not page items */
        if (!selection || !selection.length) return picked;
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (!item || !item.typename) continue;
            if (item.typename === "TextFrame") {
                if (item.kind === TextType.AREATEXT) picked.areaTextFrames.push(item);
                else picked.otherTextFrames.push(item);
            } else if (item.typename === "GroupItem") {
                collectShapeTextPairs(item, picked.shapePairs);
            } else if (getClosedPathItem(item)) {
                picked.shapeItems.push(item);
            }
        }
        return picked;
    }

    /**
     * 流し込む組み合わせ（パスと流し込み元のテキスト）を作る
     * @param {object} picked - classifySelection() の戻り値
     * @param {number} selectionLength - 選択オブジェクトの数
     * @returns {Array<object>} shapeItem / sourceTextFrame を持つオブジェクトの配列
     */
    function buildFillJobs(picked, selectionLength) {
        /* 「閉じたパス1つ＋テキスト1つ」の選択なら、そのテキストを流し込む / Pour the selected text when it is a single path plus a single text */
        var sourceTextFrame = (selectionLength === 2 && picked.shapeItems.length === 1 && picked.otherTextFrames.length === 1) ? picked.otherTextFrames[0] : null;
        var fillJobs = [];
        for (var i = 0; i < picked.shapeItems.length; i++) {
            fillJobs.push({ shapeItem: picked.shapeItems[i], sourceTextFrame: sourceTextFrame });
        }
        /* グループはそれぞれの中のテキストを流し込む / Each group pours the text it holds */
        for (var j = 0; j < picked.shapePairs.length; j++) {
            fillJobs.push({ shapeItem: picked.shapePairs[j].shapeItem, sourceTextFrame: picked.shapePairs[j].textFrame });
        }
        return fillJobs;
    }

    /**
     * エリア内文字を中央揃え・天地中央にする
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function applyCenterAlignment(doc, textFrame) {
        /* 空のフレームなどでは行揃えを設定できない / Justification can fail, e.g. on an empty frame */
        try { textFrame.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
        /* 垂直方向の配置はアクション経由なので、対象だけを選択してから実行する / Vertical centering runs as an action, so select just this frame */
        doc.selection = null;
        doc.selection = [textFrame];
        app.redraw(); /* Illustratorに選択状態を確定させる / Let Illustrator commit the selection */
        app.doScript(ACTION_ALIGN_CENTER, ACTION_SET_ALIGNMENT, false);
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var picked = classifySelection(doc.selection);
    var fillJobs = buildFillJobs(picked, doc.selection.length);
    var targetFrames = picked.areaTextFrames;
    if (!targetFrames.length && !fillJobs.length) {
        alert(getLabel("alert.selectTarget"));
        return;
    }

    /* パスはエリア内文字に変換してテキストを流し込む / Turn paths into Area Type and pour text into them */
    if (fillJobs.length) targetFrames = targetFrames.concat(fillShapesWithText(doc, fillJobs));

    loadAlignmentAction();
    try {
        for (var i = 0; i < targetFrames.length; i++) {
            applyCenterAlignment(doc, targetFrames[i]);
        }
    } finally {
        unloadAlignmentAction();
        /* 元の選択に戻す / Restore the original selection */
        if (targetFrames.length) doc.selection = targetFrames;
    }

})();
