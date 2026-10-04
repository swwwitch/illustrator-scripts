#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字の垂直方向の配置と行揃えを、まとめて中央にそろえます。閉じたパスを選択している場合はエリア内文字に変換し、サンプルテキスト、または一緒に選択した／グループ化したテキストの内容を流し込みます（グループは流し込み後に解除）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeCenterMiddle.md

### Overview

Sets both the vertical alignment and the justification of the selected Area Type frames to center in one pass. Selected closed paths are converted to Area Type and filled with sample text, or with the contents of a text object selected alongside or grouped with the path (such a group is released once the text is poured).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AreaTypeCenterMiddle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AreaTypeCenterMiddle";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

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

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTarget: { ja: "エリア内文字またはパスを選択してください。", en: "Please select area text or a path." },
            actionFailed: { ja: "アクションを実行できませんでした。", en: "Could not run the action." }
        }
    };

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

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
     * アクション名ブロック /name [ <len> <hex> ] を生成する
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(actionName) {
        var nameHex = toActionHex(actionName);
        return "/name [ " + (nameHex.length / 2) + " " + nameHex.toUpperCase() + " ]";
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

    // =========================================
    // テキストの流し込み / Text pouring
    // =========================================

    /**
     * 候補名から利用できるフォントを返す
     * @param {string[]} candidateFontNames - フォント名の候補
     * @returns {TextFont} 見つかったフォント（無ければnull）
     */
    function findAvailableTextFont(candidateFontNames) {
        for (var i = 0; i < candidateFontNames.length; i++) {
            try {
                var candidateFont = app.textFonts.getByName(candidateFontNames[i]);
                if (candidateFont) return candidateFont;
            } catch (e) { /* 無いフォント名は例外になる / getByName throws for a missing font */ }
        }
        return null;
    }

    /**
     * 閉じたパスを取り出す（複合パスは先頭のパスを見る）
     * @param {PageItem} pageItem - 選択オブジェクト
     * @returns {PathItem} 閉じたパス（無ければnull）
     */
    function getClosedPathItem(pageItem) {
        if (pageItem.typename === "PathItem") return pageItem.closed ? pageItem : null;
        if (pageItem.typename === "CompoundPathItem" && pageItem.pathItems.length > 0) {
            var firstPath = pageItem.pathItems[0];
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
     * @param {PathItem|CompoundPathItem} shapeItem - 変換するパス（複合パスも可）
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
     * @param {Object[]} fillJobs - shapeItem（変換するパス）と sourceTextFrame（流し込み元。nullならサンプルテキスト）の組み合わせ
     * @returns {TextFrame[]} 作成したエリア内文字
     */
    function fillShapesWithText(doc, fillJobs) {
        var sampleText = (uiLang === "ja") ? DUMMY_TEXT_JA : DUMMY_TEXT_EN;
        var sampleFont = findAvailableTextFont((uiLang === "ja") ? DUMMY_FONT_JA : DUMMY_FONT_EN);
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
     * @returns {Object|null} shapeItem / textFrame を持つオブジェクト（該当しなければnull）
     */
    function getShapeTextPair(groupItem) {
        /* クリップグループの枠はマスクなので対象にしない / The frame of a clipping group is a mask, not a shape to convert */
        if (groupItem.clipped) return null;
        var groupMembers = groupItem.pageItems;
        if (groupMembers.length !== 2) return null;
        var shapeTextPair = { shapeItem: null, textFrame: null };
        for (var i = 0; i < groupMembers.length; i++) {
            if (groupMembers[i].typename === "TextFrame") shapeTextPair.textFrame = groupMembers[i];
            else if (getClosedPathItem(groupMembers[i])) shapeTextPair.shapeItem = groupMembers[i];
        }
        return (shapeTextPair.shapeItem && shapeTextPair.textFrame) ? shapeTextPair : null;
    }

    /**
     * グループから「閉じたパス1つ＋テキスト1つ」の組み合わせを集める
     * @param {GroupItem} groupItem - 対象のグループ
     * @param {Object[]} shapeTextPairs - 集めた組み合わせの入れ物
     * @returns {void}
     */
    function collectShapeTextPairs(groupItem, shapeTextPairs) {
        var shapeTextPair = getShapeTextPair(groupItem);
        if (shapeTextPair) {
            shapeTextPairs.push(shapeTextPair);
            return;
        }
        /* 該当しないグループは、入れ子になっているグループを見る / Look into nested groups when the group itself is not a pair */
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            if (groupItem.pageItems[i].typename === "GroupItem") collectShapeTextPairs(groupItem.pageItems[i], shapeTextPairs);
        }
    }

    /**
     * 選択オブジェクトを、エリア内文字・閉じたパス・それ以外のテキスト・グループの組み合わせに仕分ける
     * @param {PageItem[]} selectedItems - ドキュメントの選択内容
     * @returns {Object} areaTextFrames / shapeItems / otherTextFrames / shapePairs を持つオブジェクト
     */
    function classifySelection(selectedItems) {
        var classifiedItems = { areaTextFrames: [], shapeItems: [], otherTextFrames: [], shapePairs: [] };
        /* 文字編集中は選択がTextRangeになり、ページアイテムが取り出せない / While editing text the selection is a TextRange, not page items */
        if (!selectedItems || !selectedItems.length) return classifiedItems;
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!selectedItem || !selectedItem.typename) continue;
            if (selectedItem.typename === "TextFrame") {
                if (selectedItem.kind === TextType.AREATEXT) classifiedItems.areaTextFrames.push(selectedItem);
                else classifiedItems.otherTextFrames.push(selectedItem);
            } else if (selectedItem.typename === "GroupItem") {
                collectShapeTextPairs(selectedItem, classifiedItems.shapePairs);
            } else if (getClosedPathItem(selectedItem)) {
                classifiedItems.shapeItems.push(selectedItem);
            }
        }
        return classifiedItems;
    }

    /**
     * 流し込む組み合わせ（パスと流し込み元のテキスト）を作る
     * @param {Object} classifiedItems - classifySelection() の戻り値
     * @param {number} selectionLength - 選択オブジェクトの数
     * @returns {Object[]} shapeItem / sourceTextFrame を持つオブジェクトの配列
     */
    function buildFillJobs(classifiedItems, selectionLength) {
        /* 「閉じたパス1つ＋テキスト1つ」の選択なら、そのテキストを流し込む / Pour the selected text when it is a single path plus a single text */
        var sourceTextFrame = (selectionLength === 2 && classifiedItems.shapeItems.length === 1 && classifiedItems.otherTextFrames.length === 1) ? classifiedItems.otherTextFrames[0] : null;
        var fillJobs = [];
        for (var i = 0; i < classifiedItems.shapeItems.length; i++) {
            fillJobs.push({ shapeItem: classifiedItems.shapeItems[i], sourceTextFrame: sourceTextFrame });
        }
        /* グループはそれぞれの中のテキストを流し込む / Each group pours the text it holds */
        for (var j = 0; j < classifiedItems.shapePairs.length; j++) {
            fillJobs.push({ shapeItem: classifiedItems.shapePairs[j].shapeItem, sourceTextFrame: classifiedItems.shapePairs[j].textFrame });
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
    var classifiedItems = classifySelection(doc.selection);
    var fillJobs = buildFillJobs(classifiedItems, doc.selection.length);
    var targetFrames = classifiedItems.areaTextFrames;
    if (!targetFrames.length && !fillJobs.length) {
        alert(getLabel("alert.selectTarget"));
        return;
    }

    /* パスはエリア内文字に変換してテキストを流し込む / Turn paths into Area Type and pour text into them */
    if (fillJobs.length) targetFrames = targetFrames.concat(fillShapesWithText(doc, fillJobs));

    if (!loadTemporaryActionSet(buildAlignCenterAia(), ACTION_SET_ALIGNMENT)) {
        alert(getLabel("alert.actionFailed"));
        return;
    }
    try {
        for (var i = 0; i < targetFrames.length; i++) {
            applyCenterAlignment(doc, targetFrames[i]);
        }
    } finally {
        unloadTemporaryActionSet(ACTION_SET_ALIGNMENT);
        /* 元の選択に戻す / Restore the original selection */
        if (targetFrames.length) doc.selection = targetFrames;
    }

})();
