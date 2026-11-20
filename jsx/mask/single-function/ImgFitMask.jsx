#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ImgFitMask
Meta-Default-language=ja
Name=図形の大きさに画像を合わせてクリッピングマスクを作成
Name[en]=Fit an image to a shape and clip it
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.3
Tags=mask,マスク
X-com.swwwitch-Script-Name=ImgFitMask
X-com.swwwitch-Updated=2026-10-04
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImgFitMask.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImgFitMask.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/single-function/ImgFitMask.jsx
Description-BEGIN
[ja]
1つの図形と1つの画像を選択して実行すると、図形の大きさに合わせて画像を拡大・縮小してからクリッピングマスクを作成します。
[en]
With one shape and one image selected, scales the image to the shape and then creates a clipping mask.
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo: {
                ja: "エラー: 1つの図形（マスク用）と1つの画像を選択してください。",
                en: "Error: Select one shape (for the mask) and one image."
            },
            needShapeAndImage: {
                ja: "エラー: 「パス（図形）」と「配置画像」をそれぞれ1つずつ選択してください。",
                en: "Error: Select one path (shape) and one placed image."
            }
        }
    };

    // =========================================
    // 画像とマスク / Image and mask
    // =========================================

    /**
     * マスクに使える図形（パス／複合パス）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} 図形なら true
     */
    function isMaskShape(item) {
        return item.typename === "PathItem" || item.typename === "CompoundPathItem";
    }

    /**
     * 画像（配置画像／埋め込み画像）か
     * @param {PageItem} item - 判定するアイテム
     * @returns {boolean} 画像なら true
     */
    function isImageItem(item) {
        return item.typename === "PlacedItem" || item.typename === "RasterItem";
    }

    /**
     * アイテムの中心座標を返す
     * @param {PageItem} item - 対象アイテム
     * @returns {number[]} [x, y]
     */
    function getCenter(item) {
        return [
            item.left + item.width / 2,
            item.top - item.height / 2
        ];
    }

    /**
     * 画像を図形を隙間なく覆う大きさに拡大・縮小して中央に合わせ、図形でクリッピングマスクを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacedItem|RasterItem} image - マスクされる画像
     * @param {PathItem|CompoundPathItem} maskShape - マスクになる図形
     * @returns {GroupItem} 作成したクリップグループ
     */
    function fitImageToMask(doc, image, maskShape) {
        /* 隙間が出ないよう、倍率の大きい方を使う（全体を収めるなら Math.min）/ Use the larger ratio so no gap remains */
        var scaleFactor = Math.max(maskShape.width / image.width, maskShape.height / image.height);
        /* resize はパーセント指定 / resize() takes percentages */
        image.resize(scaleFactor * 100, scaleFactor * 100, true, true, true, true);

        /* リサイズ後の中心を図形の中心に合わせる / Center the resized image on the shape */
        var maskCenter = getCenter(maskShape);
        var imageCenter = getCenter(image);
        image.translate(maskCenter[0] - imageCenter[0], maskCenter[1] - imageCenter[1]);

        /* 図形の重ね順の位置にグループを作り、図形を最前面に入れる / Group at the shape's stacking position, shape on top */
        var clipGroup = doc.groupItems.add();
        clipGroup.move(maskShape, ElementPlacement.PLACEBEFORE);
        maskShape.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        image.move(clipGroup, ElementPlacement.PLACEATEND);
        clipGroup.clipped = true;
        return clipGroup;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した図形と画像から、画像を図形に合わせたクリッピングマスクを作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var currentSelection = doc.selection;

        if (currentSelection.length !== 2) {
            alert(getLabel("alert.selectTwo"));
            return;
        }

        var maskShape = null;   /* マスクになる図形 / shape used as the mask */
        var targetImage = null; /* マスクされる画像 / image being masked */
        for (var i = 0; i < currentSelection.length; i++) {
            var item = currentSelection[i];
            if (isMaskShape(item)) {
                maskShape = item;
            } else if (isImageItem(item)) {
                targetImage = item;
            }
        }

        if (!maskShape || !targetImage) {
            alert(getLabel("alert.needShapeAndImage"));
            return;
        }

        var clipGroup = fitImageToMask(doc, targetImage, maskShape);

        /* 作ったクリップグループを選択 / Select the new clipping group */
        doc.selection = null;
        clipGroup.selected = true;
    }

    main();

})();
