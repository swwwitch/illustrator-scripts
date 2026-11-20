#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.RegisterAndApplySwatches
Meta-Default-language=ja
Name=塗りと線のカラーをスウォッチに登録して再適用
Name[en]=Register fill and stroke colors as swatches and reapply them
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.3
Release-Date=2025-06-26
Tags=color,カラー
X-com.swwwitch-Script-Name=RegisterAndApplySwatches
X-com.swwwitch-Updated=2026-10-04
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RegisterAndApplySwatches.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegisterAndApplySwatches.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/RegisterAndApplySwatches.jsx
Description-BEGIN
[ja]
選択オブジェクトの塗りおよび線のカラーをスウォッチ（スポットカラー）として登録し、その場で再適用します。
RGBとCMYKに対応し、既存の同名スウォッチは再利用します。
[en]
Registers the fill and stroke colors of the selected objects as spot-color swatches and reapplies them in place.
RGB and CMYK are supported, and an existing swatch with the same name is reused.
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

    var LABELS = {
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." },
            noTextFill: { ja: "塗りが設定されていません。", en: "The text has no fill color." },
            applyError: { ja: "カラーの適用中にエラーが発生しました。", en: "An error occurred while applying colors." }
        }
    };

    // =========================================
    // スウォッチ登録 / Swatch registration
    // =========================================

    /**
     * RGB／CMYK カラーの複製を作る
     * @param {Color} color - 元のカラー
     * @returns {RGBColor|CMYKColor|null} 複製（RGB／CMYK 以外は null）
     */
    function createColorCopy(color) {
        var colorCopy = null;
        switch (color.typename) {
            case "RGBColor":
                colorCopy = new RGBColor();
                colorCopy.red = color.red;
                colorCopy.green = color.green;
                colorCopy.blue = color.blue;
                break;
            case "CMYKColor":
                colorCopy = new CMYKColor();
                colorCopy.cyan = color.cyan;
                colorCopy.magenta = color.magenta;
                colorCopy.yellow = color.yellow;
                colorCopy.black = color.black;
                break;
        }
        return colorCopy;
    }

    /**
     * カラーの値からスウォッチ名を作る（例: "C=0 M=100 Y=100 K=0"）
     * @param {Color} color - 元のカラー
     * @returns {string} スウォッチ名（RGB／CMYK 以外は空文字）
     */
    function buildSwatchName(color) {
        if (color.typename === "RGBColor") {
            return "R=" + Math.round(color.red) + " G=" + Math.round(color.green) + " B=" + Math.round(color.blue);
        }
        if (color.typename === "CMYKColor") {
            return "C=" + Math.round(color.cyan) +
                " M=" + Math.round(color.magenta) +
                " Y=" + Math.round(color.yellow) +
                " K=" + Math.round(color.black);
        }
        return "";
    }

    /**
     * 同名のスポットカラーを探し、無ければ新規に登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} color - 登録するカラー
     * @param {string} swatchName - スウォッチ名
     * @returns {Spot|null} スポットカラー（登録できないカラーは null）
     */
    function findOrAddSpot(doc, color, swatchName) {
        for (var i = 0; i < doc.spots.length; i++) {
            if (doc.spots[i].name === swatchName) {
                return doc.spots[i];
            }
        }
        var colorCopy = createColorCopy(color);
        if (colorCopy === null) {
            return null;
        }
        var spot = doc.spots.add();
        spot.colorType = ColorModel.SPOT;
        spot.color = colorCopy;
        spot.name = swatchName;
        return spot;
    }

    /**
     * 1色をスウォッチ登録し、そのスポットカラーをオブジェクトの塗りまたは線に適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のオブジェクト
     * @param {Color} color - 登録するカラー
     * @param {boolean} isStroke - 線に適用するなら true
     * @returns {void}
     */
    function registerAndApplyColor(doc, item, color, isStroke) {
        /* RGB／CMYK 以外（なし・スポット・グレー・グラデーション・パターン）は対象外 / Only RGB and CMYK are registered */
        if (!color) return;
        var swatchName = buildSwatchName(color);
        if (swatchName === "") return;

        var spot = findOrAddSpot(doc, color, swatchName);
        if (!spot) return;

        var spotColor = new SpotColor();
        spotColor.spot = spot;
        spotColor.tint = 100;

        if (isStroke) {
            item.strokeColor = spotColor;
        } else if (item.typename === "TextFrame") {
            item.textRange.characterAttributes.fillColor = spotColor;
        } else {
            item.fillColor = spotColor;
        }
    }

    /**
     * パスの塗りと線、またはテキストの塗りをスウォッチ登録して適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem|TextFrame|CompoundPathItem} targetItem - 対象のオブジェクト
     * @returns {void}
     */
    function registerItemColors(doc, targetItem) {
        if (targetItem.typename === "PathItem") {
            registerAndApplyColor(doc, targetItem, targetItem.fillColor, false);
            registerAndApplyColor(doc, targetItem, targetItem.strokeColor, true);
        } else if (targetItem.typename === "TextFrame") {
            var textFillColor = targetItem.textRange.characterAttributes.fillColor;
            if (!textFillColor || textFillColor.typename === "NoColor") {
                alert(getLabel("alert.noTextFill"));
                return;
            }
            registerAndApplyColor(doc, targetItem, textFillColor, false);
        }
    }

    /**
     * 選択中のオブジェクトを処理する（グループは再帰的にたどる）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {void}
     */
    function processItem(doc, item) {
        if ((item.typename === "PathItem" && item.closed) || item.typename === "TextFrame") {
            registerItemColors(doc, item);
        } else if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                processItem(doc, item.pageItems[i]);
            }
        } else if (item.typename === "CompoundPathItem") {
            if (item.pathItems.length > 0) {
                registerItemColors(doc, item);
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中のオブジェクトのカラーをスウォッチ登録して適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        /* DOM の予期しない失敗を1回の通知にまとめる / Report unexpected DOM failures once */
        try {
            for (var i = 0; i < selectedItems.length; i++) {
                processItem(doc, selectedItems[i]);
            }
        } catch (e) {
            alert(getLabel("alert.applyError") + "\n" + e);
        }
    }

    main();

})();
