#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクトを一時的にグループ化したうえで、整列パネルの「水平方向中央に整列」「垂直方向中央に整列」をダイナミックアクション経由で実行します。
実行中だけ「字形の境界に整列」をONにし、終了時に元の状態へ戻します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterAlignAsGroup.md

note記事も参照してください。
https://note.com/dtp_tranist/n/xxxxxxxx

### Overview

Groups the selection temporarily, then runs Align Horizontal Center and Align Vertical Center from the Align panel through a dynamic action.
"Align to glyph bounds" is turned on only while it runs and restored afterwards.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CenterAlignAsGroup.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CenterAlignAsGroup";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CenterAlignAsGroup.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CenterAlignAsGroup.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // アクション設定 / Action settings
    // =========================================
    var ACTION_SET_NAME = "CenterAlignAsGroup"; /* アクションセット名 / action set name */
    var ACTION_NAME     = "Center";             /* アクション名 / action name */

    /* 名前は .aia 内の /name（16進とバイト数）と一致させること / Keep the names in sync with the hex in the definition */

    /* アクション定義（.aia 形式）/ Action definition */
    var ACTION_CODE = [
        "/version 3",
        "/name [ 18",
        "\t43656e746572416c69676e417347726f7570",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        "\t/name [ 6",
        "\t\t43656e746572",
        "\t]",
        "\t/keyIndex 0",
        "\t/colorIndex 0",
        "\t/isOpen 1",
        "\t/eventCount 2",
        "\t/event-1 {",
        "\t\t/useRulersIn1stQuadrant 0",
        "\t\t/internalName (ai_plugin_alignPalette)",
        "\t\t/localizedName [ 6",
        "\t\t\te695b4e58897",
        "\t\t]",
        "\t\t/isOpen 1",
        "\t\t/isOn 1",
        "\t\t/hasDialog 0",
        "\t\t/parameterCount 1",
        "\t\t/parameter-1 {",
        "\t\t\t/key 1954115685",
        "\t\t\t/showInPalette 4294967295",
        "\t\t\t/type (enumerated)",
        "\t\t\t/name [ 27",
        "\t\t\t\te6b0b4e5b9b3e696b9e59091e4b8ade5a4aee381abe695b4e58897",
        "\t\t\t]",
        "\t\t\t/value 2",
        "\t\t}",
        "\t}",
        "\t/event-2 {",
        "\t\t/useRulersIn1stQuadrant 0",
        "\t\t/internalName (ai_plugin_alignPalette)",
        "\t\t/localizedName [ 6",
        "\t\t\te695b4e58897",
        "\t\t]",
        "\t\t/isOpen 1",
        "\t\t/isOn 1",
        "\t\t/hasDialog 0",
        "\t\t/parameterCount 1",
        "\t\t/parameter-1 {",
        "\t\t\t/key 1954115685",
        "\t\t\t/showInPalette 4294967295",
        "\t\t\t/type (enumerated)",
        "\t\t\t/name [ 27",
        "\t\t\t\te59e82e79bb4e696b9e59091e4b8ade5a4aee381abe695b4e58897",
        "\t\t\t]",
        "\t\t\t/value 5",
        "\t\t}",
        "\t}",
        "}",
        ""
    ].join("\n");

    // =========================================
    // 日英ラベル定義 / Japanese-English label definitions
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:  { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
            genericError: { ja: "エラーが発生しました：", en: "An error occurred: " },
            multipleLayers: {
                ja: "複数のレイヤーにまたがって選択されています。\nグループ化するとレイヤーが1つにまとまり、解除しても元に戻らないため中止しました。",
                en: "The selection spans multiple layers.\nGrouping would merge them into one layer and ungrouping cannot undo that, so nothing was changed."
            }
        }
    };

    // =========================================
    // 環境設定 / Preferences
    // =========================================

    /**
     * 「字形の境界に整列」の現在の状態を取得する
     * @returns {{point: boolean, area: boolean}} ポイント文字・エリア内文字それぞれのON/OFF
     */
    function getGlyphBoundsAlign() {
        return {
            point: app.preferences.getBooleanPreference("EnableActualPointTextSpaceAlign") === true,
            area: app.preferences.getBooleanPreference("EnableActualAreaTextSpaceAlign") === true
        };
    }

    /**
     * 「字形の境界に整列」をポイント文字・エリア内文字それぞれに設定する
     * @param {{point: boolean, area: boolean}} glyphBoundsState - 設定するON/OFF
     * @returns {void}
     */
    function setGlyphBoundsAlign(glyphBoundsState) {
        app.preferences.setBooleanPreference("EnableActualPointTextSpaceAlign", glyphBoundsState.point);
        app.preferences.setBooleanPreference("EnableActualAreaTextSpaceAlign", glyphBoundsState.area);
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は runTemporaryAction / loadTemporaryActionSet / unloadTemporaryActionSet / toActionHex / buildActionNameLines
    // 2. アクション定義は配列＋join("\n") で組み立てる（''' は ES3 の構文エラー）。
    //    セット名・アクション名は英数字にする。/name [ n 16進 ] は buildActionNameLines で作るとバイト数がずれない
    //      var actionSource = [
    //          "/version 3"
    //      ].concat(buildActionNameLines("", "MySet"), [
    //          "/isOpen 1", "/actionCount 1", "/action-1 {"
    //      ], buildActionNameLines("\t", "myAction"), [ … ]).join("\n");
    // 3. 1回だけ実行するとき:
    //      if (!runTemporaryAction(actionSource, "MySet", "myAction")) alert(getLabel("alert.actionFailed"));
    //    何度も実行するとき（オブジェクトごとなど）は、読み込み・解除を1回ずつにする:
    //      if (!loadTemporaryActionSet(actionSource, "MySet")) { alert(…); return; }
    //      try { for (…) app.doScript("myAction", "MySet"); } finally { unloadTemporaryActionSet("MySet"); }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。警告を出すかはコピー先で決める
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // アートボード / Artboards
    // =========================================

    /**
     * 選択オブジェクト全体を囲む矩形を求める
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {number[]} [左, 上, 右, 下] の座標
     */
    function getSelectionBounds(selectedItems) {
        var bounds = selectedItems[0].visibleBounds;
        var left = bounds[0];
        var top = bounds[1];
        var right = bounds[2];
        var bottom = bounds[3];
        for (var i = 1; i < selectedItems.length; i++) {
            var itemBounds = selectedItems[i].visibleBounds;
            if (itemBounds[0] < left) left = itemBounds[0];
            if (itemBounds[1] > top) top = itemBounds[1];
            if (itemBounds[2] > right) right = itemBounds[2];
            if (itemBounds[3] < bottom) bottom = itemBounds[3];
        }
        return [left, top, right, bottom];
    }

    /**
     * 2つの矩形が重なっている面積を求める
     * @param {number[]} boundsA - [左, 上, 右, 下] の座標
     * @param {number[]} boundsB - [左, 上, 右, 下] の座標
     * @returns {number} 重なっている面積（重ならない場合は 0）
     */
    function getOverlapArea(boundsA, boundsB) {
        var overlapWidth = Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]);
        var overlapHeight = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
        if (overlapWidth <= 0 || overlapHeight <= 0) {
            return 0;
        }
        return overlapWidth * overlapHeight;
    }

    /**
     * 選択範囲と最も広く重なるアートボードを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} selectionBounds - [左, 上, 右, 下] の座標
     * @param {number} currentIndex - 重なりが同じときに優先するアートボード番号
     * @returns {number} アートボード番号（どこにも重ならない場合は -1）
     */
    function findOverlappingArtboardIndex(doc, selectionBounds, currentIndex) {
        /* 現在のアートボードを先に見て、重なりが同じなら切り替えない / Check the current artboard first so ties keep it */
        var searchOrder = [currentIndex];
        for (var i = 0; i < doc.artboards.length; i++) {
            if (i !== currentIndex) {
                searchOrder.push(i);
            }
        }
        var largestOverlapIndex = -1;
        var largestOverlapArea = 0;
        for (var j = 0; j < searchOrder.length; j++) {
            var overlapArea = getOverlapArea(selectionBounds, doc.artboards[searchOrder[j]].artboardRect);
            if (overlapArea > largestOverlapArea) {
                largestOverlapArea = overlapArea;
                largestOverlapIndex = searchOrder[j];
            }
        }
        return largestOverlapIndex;
    }

    /**
     * 選択範囲の中心に最も近いアートボードを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} selectionBounds - [左, 上, 右, 下] の座標
     * @returns {number} アートボード番号
     */
    function findNearestArtboardIndex(doc, selectionBounds) {
        var centerX = (selectionBounds[0] + selectionBounds[2]) / 2;
        var centerY = (selectionBounds[1] + selectionBounds[3]) / 2;
        var nearestIndex = 0;
        var nearestDistance = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardRect = doc.artboards[i].artboardRect;
            var offsetX = centerX - (artboardRect[0] + artboardRect[2]) / 2;
            var offsetY = centerY - (artboardRect[1] + artboardRect[3]) / 2;
            var distance = offsetX * offsetX + offsetY * offsetY;
            if (nearestDistance === null || distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /**
     * 選択が現在のアートボード上にないとき、選択を含むアートボードを現在のアートボードにする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function activateArtboardForSelection(doc, selectedItems) {
        if (doc.artboards.length < 2) {
            return;
        }
        var selectionBounds = getSelectionBounds(selectedItems);
        var currentIndex = doc.artboards.getActiveArtboardIndex();
        var targetIndex = findOverlappingArtboardIndex(doc, selectionBounds, currentIndex);
        /* どのアートボードにも重ならないときは一番近いアートボードを使う / Fall back to the nearest artboard */
        if (targetIndex < 0) {
            targetIndex = findNearestArtboardIndex(doc, selectionBounds);
        }
        if (targetIndex !== currentIndex) {
            doc.artboards.setActiveArtboardIndex(targetIndex);
        }
    }

    // =========================================
    // 選択の検査 / Selection checks
    // =========================================

    /**
     * 文字を部分選択している場合に、その文字を含むテキストオブジェクトを選択し直す
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択し直したあとの選択内容
     */
    function selectTextFrameFromTextRange(doc) {
        var storyFrames = doc.selection.story.textFrames;
        var targetFrames = [];
        for (var i = 0; i < storyFrames.length; i++) {
            targetFrames.push(storyFrames[i]);
        }
        /* 文字編集を抜けてからテキストオブジェクトを選択 / Leave text editing, then select the frames */
        app.executeMenuCommand("deselectall");
        for (var j = 0; j < targetFrames.length; j++) {
            targetFrames[j].selected = true;
        }
        return doc.selection;
    }

    /**
     * レイヤーを一意に識別するキーを作る（サブレイヤーの親子関係も含める）
     * @param {Layer} layer - 対象レイヤー
     * @returns {string} 識別用のキー
     */
    function getLayerKey(layer) {
        var keyParts = [];
        var layerNode = layer;
        while (layerNode && layerNode.typename === "Layer") {
            keyParts.push(layerNode.zOrderPosition + ":" + layerNode.name);
            layerNode = layerNode.parent;
        }
        return keyParts.join("/");
    }

    /**
     * 選択が複数のレイヤーにまたがっているか判定する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {boolean} またがっていれば true
     */
    function spansMultipleLayers(selectedItems) {
        var firstKey = getLayerKey(selectedItems[0].layer);
        for (var i = 1; i < selectedItems.length; i++) {
            if (getLayerKey(selectedItems[i].layer) !== firstKey) {
                return true;
            }
        }
        return false;
    }

    // =========================================
    // テキスト / Text
    // =========================================

    /**
     * 選択が1行だけのテキストオブジェクト1つか判定する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 1行だけのテキストオブジェクト1つなら true
     */
    function isSingleLineTextFrame(selectedItems) {
        if (selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") {
            return false;
        }
        /* 折り返しも含めた実際の行数で判定 / Count the rendered lines, wrapping included */
        return selectedItems[0].lines.length === 1;
    }

    /**
     * テキストオブジェクト全体の行揃えを中央揃えにする
     * @param {TextFrame} textFrame - 対象のテキストオブジェクト
     * @returns {void}
     */
    function setCenterJustification(textFrame) {
        textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確認し、選択を含むアートボードに切り替えてから、複数選択時は一時的にグループ化してアクションを実行する
     * 1行だけのテキストオブジェクト1つの選択では、行揃えも中央揃えにする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字を部分選択しているときは selection が TextRange になるため、テキストオブジェクトに置き換える
           A partial text selection comes back as a TextRange; promote it to the text object */
        if (selectedItems && !(selectedItems instanceof Array)) {
            selectedItems = selectTextFrameFromTextRange(doc);
        }
        if (!(selectedItems instanceof Array) || selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 複数選択時はまとめて動かすため一時的にグループ化 / Group temporarily so the selection moves as one */
        var needsGroup = selectedItems.length > 1;
        /* レイヤーをまたぐ選択はグループ化でレイヤーが移動してしまうため中止 / Grouping across layers is not reversible */
        if (needsGroup && spansMultipleLayers(selectedItems)) {
            alert(getLabel("alert.multipleLayers"));
            return;
        }
        /* 選択が現在のアートボード外にあるときは、選択を含むアートボードに切り替える
           Aligning to the artboard uses the active one, so switch to the one holding the selection */
        activateArtboardForSelection(doc, selectedItems);

        /* 実行中だけONにして、終了時に元の状態へ戻す / Turn on for this run only, then restore */
        var previousGlyphBounds = getGlyphBoundsAlign();
        try {
            setGlyphBoundsAlign({ point: true, area: true });
            /* 1行だけのテキスト1つの選択は行揃えも中央揃えにする / A lone single-line text object gets centered justification too */
            if (isSingleLineTextFrame(selectedItems)) {
                setCenterJustification(selectedItems[0]);
            }
            if (needsGroup) {
                app.executeMenuCommand("group");
            }
            if (!runTemporaryAction(ACTION_CODE, ACTION_SET_NAME, ACTION_NAME)) {
                throw new Error("Could not run the action \"" + ACTION_NAME + "\".");
            }
        } catch (e) {
            alert(getLabel("alert.genericError") + e);
        } finally {
            if (needsGroup) {
                app.executeMenuCommand("ungroup");
            }
            setGlyphBoundsAlign(previousGlyphBounds);
        }
    }

    main();

})();
