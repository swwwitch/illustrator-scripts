#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.AddTrimMarkToCurrentArtboard
Meta-Default-language=ja
Name=現在のアートボードに日本式トンボを作成
Name[en]=Create Japanese-style trim marks on the current artboard
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.1
Release-Date=2025-02-05
Tags=misc,その他
X-com.swwwitch-Script-Name=AddTrimMarkToCurrentArtboard
X-com.swwwitch-Updated=2026-09-27
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddTrimMarkToCurrentArtboard.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddTrimMarkToCurrentArtboard.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/n40e3e39cf9f2
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/single-function/AddTrimMarkToCurrentArtboard.jsx
Description-BEGIN
[ja]
現在のアートボードを対象に、トンボを作成します。
[en]
Creates trim marks for the current artboard.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* トンボを作るレイヤー名（無ければ作る） / Layer that receives the trim marks; created when missing */
    var TRIM_LAYER_NAME = "トンボ";

    // =========================================
    // トンボ / Trim marks
    // =========================================

    /**
     * トンボ用のレイヤーを返す（無ければ作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} トンボ用のレイヤー
     */
    function getOrCreateTrimLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === TRIM_LAYER_NAME) return doc.layers[i];
        }
        var trimLayer = doc.layers.add();
        trimLayer.name = TRIM_LAYER_NAME;
        return trimLayer;
    }

    /**
     * アートボードと同じ大きさの矩形を作り、そこからトンボを作成して矩形をガイドにする
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} trimLayer - トンボ用のレイヤー（ロック解除済み）
     * @param {Artboard} artboard - 対象のアートボード
     * @returns {void}
     */
    function createTrimMarks(doc, trimLayer, artboard) {
        /* 日本式トンボをONに設定 / Enable Japanese-style trim marks */
        app.preferences.setBooleanPreference("cropMarkStyle", 1);

        /* 「トンボ」レイヤー上にアートボード矩形を作成 / Create the artboard rectangle on the trim layer */
        var artboardRect = artboard.artboardRect; /* [左, 上, 右, 下] / [left, top, right, bottom] */
        var artboardFrame = trimLayer.pathItems.rectangle(artboardRect[1], artboardRect[0], artboardRect[2] - artboardRect[0], artboardRect[1] - artboardRect[3]);
        artboardFrame.filled = false;
        artboardFrame.stroked = false;

        /* 矩形を選択してトリムマークを作成 / Select the rectangle and create trim marks */
        doc.selection = [artboardFrame];
        app.executeMenuCommand("TrimMark v25");

        /* 矩形はガイドにする / Turn the rectangle into a guide */
        artboardFrame.guides = true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * アクティブなアートボードにトンボを作成する（「トンボ」レイヤーのロックは一時的に外して戻す）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var trimLayer = getOrCreateTrimLayer(doc);
        var layerWasLocked = trimLayer.locked;
        if (layerWasLocked) trimLayer.locked = false;

        /* 途中で失敗しても選択解除とロックの復元は行う / Always clear the selection and restore the lock */
        try {
            createTrimMarks(doc, trimLayer, doc.artboards[doc.artboards.getActiveArtboardIndex()]);
        } finally {
            doc.selection = null;
            trimLayer.locked = layerWasLocked;
        }
    }

    main();

})();
