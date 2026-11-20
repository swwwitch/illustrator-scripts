#target illustrator
#targetengine "UnembedToLinksEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.UnembedToLinks
Meta-Default-language=ja
Name=埋め込み画像をリンク画像に変換（アクション方式）
Name[en]=Unembed images back into links
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.14
Release-Date=2026-07-27
Tags=link,リンク
X-com.swwwitch-Script-Name=UnembedToLinks
X-com.swwwitch-Updated=2026-11-20
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/UnembedToLinks.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/UnembedToLinks.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/n6d9e2dabb054
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/UnembedToLinks.jsx
Description-BEGIN
[ja]
埋め込み画像を、一時アクションを動的に生成して実行することでリンク画像に置き換えます。
「選択オブジェクトと置換」で配置するため、位置・サイズ・回転角・重ね順はそのまま引き継がれます。
[en]
Replaces embedded images with linked images by generating and running a temporary action on the fly.
Because the placement replaces the selected object, position, size, rotation and stacking order all carry over unchanged.
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

    var ACTION_SET_NAME    = "UnembedToLinksTempSet";    /* 一時アクションセット名 / temporary action set name */
    var ACTION_NAME        = "UnembedToLinksPlace";      /* 一時アクション名 / temporary action name */
    var LINKS_FOLDER_NAME  = "Links";                    /* 収集先フォルダー名 / collect destination folder */

    /* 元ファイルが不明なときのPSD書き出し設定 / PSD export fallback */
    var EXPORT_RESOLUTION  = 72;                         /* 書き出し解像度（ppi） / export resolution */
    var USE_XMP_NAMES      = true;                       /* XMPマニフェストの元ファイル名を使う / use names from XMP manifest */

    /* Dropboxのローカルマウントパス。空文字にするとホーム直下から自動検出 / Local Dropbox mount path ("" = auto detect) */
    var DROPBOX_PREFIX     = resolveDropboxPrefix("");

    // =========================================
    // 文字列エンコード / String encoding
    // =========================================

    /**
     * URIエンコードされた文字列をデコードする。
     * File.name のように「%」が含まれうる値でも例外を投げない。
     * @param {string} sourceText - デコードする文字列
     * @returns {string} デコード結果。不正なエスケープを含む場合は元の文字列
     */
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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "埋め込み画像をリンクに変換", en: "Unembed Images to Links" }
        },
        panel: {
            scope: { ja: "対象", en: "Target" },
            list:  { ja: "対象ファイル", en: "Files" }
        },
        column: {
            fileName: { ja: "ファイル名", en: "File name" },
            path:     { ja: "パス", en: "Path" }
        },
        radio: {
            selection: { ja: "選択している画像のみ", en: "Selected images only" },
            artboard:  { ja: "現在のアートボード上の埋め込み画像", en: "Embedded images on the current artboard" },
            all:       { ja: "すべての埋め込み画像", en: "All embedded images" }
        },
        checkbox: {
            fullPath: { ja: "フルパス", en: "Full path" },
            dropbox:  { ja: "Dropboxパスを短縮", en: "Shorten Dropbox paths" },
            collect:  { ja: "再リンク後に収集（同階層の「{folder}」フォルダーへコピー）",
                        en: "Collect after relinking (copy into a \"{folder}\" folder beside the document)" }
        },
        tooltip: {
            selection: { ja: "いま選択している埋め込み画像だけを変換します。", en: "Converts only the embedded images that are selected." },
            artboard:  { ja: "現在のアートボードに載っている埋め込み画像を変換します。", en: "Converts the embedded images on the current artboard." },
            all:       { ja: "ドキュメント内のすべての埋め込み画像を変換します。", en: "Converts every embedded image in the document." },
            fullPath:  { ja: "一覧にファイルの絶対パスを表示します。オフだとファイル名だけになります。", en: "Shows the full path in the list. Off shows just the file name." },
            dropbox:   { ja: "Dropbox のパスを短い表記に置き換えて表示します。", en: "Shows Dropbox paths in a shortened form." },
            collect:   { ja: "書き出したリンク画像を、ドキュメントと同じ階層のフォルダーへコピーしてまとめます。", en: "Copies the unembedded images into a folder beside the document." }
        },
        count: { ja: "（{n} 件）", en: " ({n})" }
    };

    /**
     * 件数付きのラベルを作る
     * @param {string} labelPath - ラベルのドット区切りキー
     * @param {number} itemCount - 表示する件数
     * @returns {string} 件数を添えたラベル
     */
    function labelWithCount(labelPath, itemCount) {
        return getLabel(labelPath) + getLabel("count", { n: itemCount });
    }

    function safeDecodeURI(sourceText) {
        try {
            return decodeURI(sourceText);
        } catch (e) {
            return String(sourceText);
        }
    }

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

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

    /**
     * .aia のテキスト値（`[ バイト数 16進 ]`）を組み立てる。記録された.aiaと同じく16進は32バイトごとに改行する。
     * @param {string} sourceText - 埋め込む文字列
     * @param {string} indentText - 閉じ括弧と16進行のインデント
     * @returns {string} .aia形式のテキスト値
     */
    function buildTextValue(sourceText, indentText) {
        var valueHex = toActionHex(sourceText);
        if (valueHex.length === 0) {
            return "[ 0 ]";
        }
        var lineList = [];
        for (var start = 0; start < valueHex.length; start += 64) {
            lineList.push(indentText + "\t" + valueHex.substr(start, 64));
        }
        return "[ " + (valueHex.length / 2) + "\n"
            + lineList.join("\n") + "\n"
            + indentText + "]";
    }

    /**
     * adobe_placeDocument のパラメーター定義。
     * キーと型は、実際に記録した.aiaから採取した値をそのまま使う。
     * ustring（ファイルパス）の値だけは実行時に差し込む。
     */
    var PLACE_PARAMETERS = [
        { key: 1851878757, type: "ustring", value: null,  note: "ファイルパス / file path (name)" },
        { key: 1818848875, type: "boolean", value: "1",   note: "リンクとして配置 / place as link (link)" },
        { key: 1919970403, type: "boolean", value: "1",   note: "選択オブジェクトと置換 / replace selection (rplc)" },
        { key: 1953329260, type: "boolean", value: "0",   note: "テンプレート / template (tmpl)" },
        { key: 1768779887, type: "boolean", value: "0",   note: "読み込みオプション / import options (impo)" },
        { key: 1885828462, type: "boolean", value: "0",   note: "ページ番号指定 / page option (pgun)" },
        { key: 1935895653, type: "real",    value: "1.0", note: "拡大縮小率 / scale (scle)" },
        { key: 1953656440, type: "real",    value: "0.0", note: "水平移動量 / translate X (trnx)" },
        { key: 1953656441, type: "real",    value: "0.0", note: "垂直移動量 / translate Y (trny)" }
    ];

    /**
     * パラメーター1件分の .aia ブロックを組み立てる。
     * @param {number} index - パラメーター番号（1始まり）
     * @param {object} parameter - PLACE_PARAMETERS の1要素
     * @param {string} filePath - 配置する画像ファイルのフルパス（fsName）
     * @returns {string} .aia形式のパラメーターブロック
     */
    function buildParameterBlock(index, parameter, filePath) {
        var indentText = "\t\t\t";
        var valueText = (parameter.value !== null)
            ? parameter.value
            : buildTextValue(filePath, indentText);

        return '\t\t/parameter-' + index + ' {\n'
            + indentText + '/key ' + parameter.key + '\n'
            + indentText + '/showInPalette 4294967295\n'
            + indentText + '/type (' + parameter.type + ')\n'
            + indentText + '/value ' + valueText + '\n'
            + '\t\t}\n';
    }

    /**
     * 配置（adobe_placeDocument）を実行する一時アクションのソースを生成する。
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} filePath - 配置する画像ファイルのフルパス（fsName）
     * @returns {string} .aia形式のアクションソース
     */
    function buildActionSource(setName, actionName, filePath) {
        var parameterText = "";
        for (var i = 0; i < PLACE_PARAMETERS.length; i++) {
            parameterText += buildParameterBlock(i + 1, PLACE_PARAMETERS[i], filePath);
        }

        return ''
            + '/version 3\n'
            + '/name ' + buildTextValue(setName, '') + '\n'
            + '/isOpen 1\n'
            + '/actionCount 1\n'
            + '/action-1 {\n'
            + '\t/name ' + buildTextValue(actionName, '\t') + '\n'
            + '\t/keyIndex 0\n'
            + '\t/colorIndex 0\n'
            + '\t/isOpen 1\n'
            + '\t/eventCount 1\n'
            + '\t/event-1 {\n'
            + '\t\t/useRulersIn1stQuadrant 0\n'
            + '\t\t/internalName (adobe_placeDocument)\n'
            + '\t\t/localizedName [ 0 ]\n'
            + '\t\t/isOpen 1\n'
            + '\t\t/isOn 1\n'
            + '\t\t/hasDialog 1\n'
            + '\t\t/showDialog 0\n'
            + '\t\t/parameterCount ' + PLACE_PARAMETERS.length + '\n'
            + parameterText
            + '\t}\n'
            + '}\n';
    }

    // =========================================
    // 再リンク処理 / Relink processing
    // =========================================

    /**
     * 選択できない状態になっている理由を返す。
     * 親グループがロック・非表示の場合も選択できないため、祖先を辿って調べる。
     * @param {PageItem} item - 調べるアイテム
     * @returns {string} 選択できない理由。選択できる場合は空文字
     */
    function findSelectionBlocker(item) {
        var node = item;

        while (node && node.typename !== "Document") {

            if (node.typename === "Layer") {
                if (node.locked)   return "レイヤー「" + node.name + "」がロックされています。";
                if (!node.visible) return "レイヤー「" + node.name + "」が非表示です。";

            } else {
                var target = (node === item) ? "対象の画像" : "親グループ";
                if (node.locked) return target + "がロックされています。";
                if (node.hidden) return target + "が非表示です。";
            }

            node = node.parent;
        }
        return "";
    }

    /**
     * 2つのアイテムが同じアートオブジェクトを指しているかを返す。
     * @param {PageItem} itemA - 比較するアイテム
     * @param {PageItem} itemB - 比較するアイテム
     * @returns {boolean} 同じアイテムとみなせる場合はtrue
     */
    function isSameArtItem(itemA, itemB) {
        try {
            return itemA.uuid === itemB.uuid;
        } catch (e) {
            /* uuidを参照できない場合は判定しない */
            return true;
        }
    }

    /**
     * 埋め込み画像（RasterItem）を、ダイナミックアクション経由でリンク画像に置き換える。
     * 選択オブジェクトの置換（rplc）で実行するため、位置・サイズ・角度・重なり順は
     * アクション側が引き継ぐ。スクリプト側での復元処理は行わない。
     * @param {Document} doc - 対象ドキュメント
     * @param {RasterItem} item - 置き換え対象の埋め込み画像
     * @param {File} targetFile - リンク先の画像ファイル
     * @param {object} status - 実行状況を書き戻すオブジェクト（{ actionPlayed: boolean }）
     * @returns {PlacedItem} 置換後のリンク画像
     */
    function relinkByAction(doc, item, targetFile, status) {
        var blockerReason = findSelectionBlocker(item);
        if (blockerReason) {
            throw new Error(blockerReason);
        }

        /* 置換対象だけを選択した状態で実行 / Select only the replacement target */
        doc.activeLayer = item.layer;
        doc.selection = null;
        item.selected = true;

        /* 選択できていないとアクションが単なる配置になるため、ここで中止 */
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length !== 1 || !isSameArtItem(currentSelection[0], item)) {
            doc.selection = null;
            throw new Error("対象の画像を選択できませんでした。");
        }

        var actionSource = buildActionSource(ACTION_SET_NAME, ACTION_NAME, targetFile.fsName);

        /* この時点以降はアクションが実行済みとして扱う / The action may have run from here on */
        status.actionPlayed = true;
        if (!runTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME)) {
            throw new Error("アクションによる置換結果を取得できませんでした。");
        }

        /* 置換直後の選択がリンク画像 / The replaced link image is selected right after the action */
        var newSelection = doc.selection;
        if (!newSelection || newSelection.length === 0 || newSelection[0].typename !== "PlacedItem") {
            throw new Error("アクションによる置換結果を取得できませんでした。");
        }

        return newSelection[0];
    }

    // =========================================
    // 収集処理 / Collect processing
    // =========================================

    /**
     * 2つのファイルを同一とみなせるか判定する。
     * @param {File} fileA - 比較対象1
     * @param {File} fileB - 比較対象2
     * @returns {boolean} 同一とみなせる場合はtrue
     */
    function isSameFile(fileA, fileB) {
        if (fileA.fsName === fileB.fsName) return true;
        return fileA.length === fileB.length
            && fileA.modified.getTime() === fileB.modified.getTime();
    }

    /**
     * 収集先のファイルを決定する。同名かつ内容が異なるファイルがある場合は連番を付ける。
     * @param {Folder} linksFolder - 収集先フォルダー
     * @param {File} sourceFile - 収集元のファイル
     * @returns {File} 収集先のファイル
     */
    function resolveCollectDestination(linksFolder, sourceFile) {
        var fileName  = safeDecodeURI(sourceFile.name);
        var dotIndex  = fileName.lastIndexOf(".");
        var baseName  = (dotIndex > 0) ? fileName.substring(0, dotIndex) : fileName;
        var extension = (dotIndex > 0) ? fileName.substring(dotIndex) : "";

        var destFile = new File(linksFolder.fsName + "/" + fileName);
        var counter = 1;
        while (destFile.exists && !isSameFile(destFile, sourceFile)) {
            destFile = new File(linksFolder.fsName + "/" + baseName + "-" + counter + extension);
            counter++;
        }
        return destFile;
    }

    /**
     * ドキュメントと同階層の「Links」フォルダーを返す。存在しない場合は作成する。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Folder} 収集先フォルダー
     */
    function getLinksFolder(doc) {
        var docFile = doc.fullName;
        if (!docFile || !docFile.exists) {
            throw new Error("ドキュメントが保存されていないため、「" + LINKS_FOLDER_NAME + "」フォルダーの場所を決定できません。");
        }

        var linksFolder = new Folder(docFile.parent.fsName + "/" + LINKS_FOLDER_NAME);
        if (!linksFolder.exists && !linksFolder.create()) {
            throw new Error("「" + LINKS_FOLDER_NAME + "」フォルダーを作成できませんでした。");
        }
        return linksFolder;
    }

    /**
     * リンク先ファイルをドキュメントと同階層の「Links」フォルダーへ複製し、リンクを張り替える。
     * @param {Document} doc - 対象ドキュメント
     * @param {PlacedItem} placedItem - 張り替え対象のリンク画像
     * @param {File} sourceFile - 収集元のファイル
     * @returns {File} 収集後のリンク先ファイル
     */
    function collectLink(doc, placedItem, sourceFile) {
        var linksFolder = getLinksFolder(doc);

        var destFile = resolveCollectDestination(linksFolder, sourceFile);
        if (!destFile.exists && !sourceFile.copy(destFile.fsName)) {
            throw new Error("リンクファイルを複製できませんでした。");
        }

        placedItem.file = destFile;
        return destFile;
    }

    // =========================================
    // 元ファイル名の推定 / Original file name resolution
    // =========================================

    /**
     * 埋め込み画像自身が持つ名前を返す。
     * レイヤー名、無ければ拡張子付きの親グループ名（配置時にファイル名が残ることがある）。
     * @param {RasterItem} item - 対象の埋め込み画像
     * @returns {string} 拡張子付きの名前。取得できない場合は空文字
     */
    function getImageNameFromItem(item) {
        if (item.name) return item.name;

        var parentItem = item.parent;
        if (parentItem == undefined || parentItem.typename !== "GroupItem") return "";

        var parentName = parentItem.name || "";
        return /\.[a-z][a-z0-9]{1,4}\s*$/i.test(parentName) ? parentName : "";
    }

    /**
     * ドキュメントのXMPマニフェストから、埋め込み前の元ファイル名を出現順に取得する。
     * @param {Document} doc - 対象ドキュメント
     * @returns {string[]} 重複を除いた元ファイル名。取得できない場合は空配列
     */
    function getManifestFileNames(doc) {
        var nameList = [];
        var foundNames = {};
        var filePaths;

        try {
            var documentXMP = new XML(doc.XMPString);

            /* 埋め込み参照のみを対象にし、取得できなければすべての参照を見る */
            filePaths = documentXMP.xpath("//stMfs:reference/stRef:filePath");
            if (filePaths == null || filePaths.length() === 0) {
                filePaths = documentXMP.xpath("//stRef:filePath");
            }
        } catch (e) {
            return nameList;
        }

        if (filePaths == null) return nameList;

        for (var i = 0; i < filePaths.length(); i++) {
            var fileName = safeDecodeURI(String(filePaths[i])).replace(/^.*[\/\\]/, "");
            var nameKey = fileName.toLowerCase();

            if (fileName === "" || foundNames[nameKey]) continue;

            foundNames[nameKey] = true;
            nameList.push(fileName);
        }
        return nameList;
    }

    /**
     * 名前が取得できなかった画像の名前を、XMPマニフェストの元ファイル名で補う。
     * 候補の件数が名前未定の画像数と一致するときだけ割り当て、一致しない場合は取り違えを避けて何もしない。
     * @param {Document} doc - 対象ドキュメント
     * @param {string[]} nameList - 画像ごとの名前。空文字の要素が埋められる
     * @returns {void}
     */
    function fillNamesFromManifest(doc, nameList) {
        var manifestNames = getManifestFileNames(doc);
        if (manifestNames.length === 0) return;

        /* すでに名前が判明している画像は、その名前を候補から除く */
        var knownNames = {};
        var missingCount = 0;
        var i;

        for (i = 0; i < nameList.length; i++) {
            if (nameList[i] === "") {
                missingCount++;
                continue;
            }
            knownNames[nameList[i].toLowerCase()] = true;
        }

        if (missingCount === 0) return;

        var remainingNames = [];
        for (i = 0; i < manifestNames.length; i++) {
            if (!knownNames[manifestNames[i].toLowerCase()]) remainingNames.push(manifestNames[i]);
        }

        /* 件数が一致しないときは、取り違えを避けるため使わない */
        if (remainingNames.length !== missingCount) return;

        var nameIndex = 0;
        for (i = 0; i < nameList.length; i++) {
            if (nameList[i] === "") nameList[i] = remainingNames[nameIndex++];
        }
    }

    /**
     * 名前から拡張子と使用できない文字を取り除き、ファイル名として整える。
     * @param {string} fileName - 元の名前（拡張子付きでも可）
     * @param {string} fallbackName - 名前が空になる場合に使う代替名
     * @returns {string} ファイル名に使用できる文字列
     */
    function toSafeBaseName(fileName, fallbackName) {
        /* 拡張子と前後の空白を除去（「2026.07.27」のような名前を壊さないよう英字始まりの2〜5文字に限定） */
        var baseName = (fileName || "").replace(/\.[a-z][a-z0-9]{1,4}$/i, "").replace(/^\s+|\s+$/g, "");

        /* ファイル名に使えない文字を置換 */
        baseName = baseName.replace(/[\\\/:*?"<>|]/g, "_");

        return baseName === "" ? fallbackName : baseName;
    }

    /**
     * ドキュメント内の埋め込み画像ごとに、書き出し用の名前（拡張子なし）を決める。
     * レイヤー名 → 拡張子付きの親グループ名 → XMPマニフェスト → 連番 の順に探す。
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} uuidをキー、拡張子を除いた名前を値とするオブジェクト
     */
    function buildExportNameMap(doc) {
        var allItems = doc.rasterItems;
        var nameList = [];
        var i;

        for (i = 0; i < allItems.length; i++) {
            nameList.push(getImageNameFromItem(allItems[i]));
        }

        if (USE_XMP_NAMES) fillNamesFromManifest(doc, nameList);

        /* 処理対象が選択範囲だけの場合もあるため、uuidで引けるようにする */
        var nameMap = {};
        for (i = 0; i < allItems.length; i++) {
            nameMap[allItems[i].uuid] = toSafeBaseName(nameList[i], "image" + (i + 1));
        }
        return nameMap;
    }

    // =========================================
    // PSD書き出し / PSD export
    // =========================================

    /**
     * このスクリプトで扱えるカラースペースかどうかを返す。
     * @param {ImageColorSpace} imageColorSpace - 判定するカラースペース
     * @returns {boolean} CMYK／RGB／グレースケールの場合はtrue
     */
    function isSupportedColorSpace(imageColorSpace) {
        return imageColorSpace == ImageColorSpace.CMYK
            || imageColorSpace == ImageColorSpace.RGB
            || imageColorSpace == ImageColorSpace.GrayScale;
    }

    /**
     * 画像アイテムに蓄積された回転角を度で返す。
     * @param {PlacedItem|RasterItem} item - 対象の画像アイテム
     * @returns {number} 回転角（度）。取得できない場合は0
     */
    function getAccumulatedRotation(item) {
        var itemTags = item.tags;

        for (var i = 0; i < itemTags.length; i++) {
            if (itemTags[i].name === "BBAccumRotation") return itemTags[i].value * 180 / Math.PI;
        }
        return 0;
    }

    /**
     * PlacedItemまたはRasterItemの拡大率と回転角を返す。
     * @param {PlacedItem|RasterItem} item - 対象の画像アイテム
     * @returns {object} 拡大率と回転角（{ scaleX: number, scaleY: number, rotation: number }）
     */
    function getScaleAndRotation(item) {
        /* RasterItemは行列のY方向が反転している */
        var placedItemFlip = (item.typename === "PlacedItem") ? 1 : -1;
        var rotationAngle = getAccumulatedRotation(item);

        var unrotatedMatrix = app.concatenateRotationMatrix(item.matrix, rotationAngle * placedItemFlip);

        return {
            scaleX: unrotatedMatrix.mValueA * 100,
            scaleY: unrotatedMatrix.mValueD * -100 * placedItemFlip,
            rotation: rotationAngle
        };
    }

    /**
     * 既存ファイルを上書きしないパスを返す。上書きを避けるため連番の接尾辞を付ける。
     * @param {string} filePath - 調べるパス
     * @returns {string} 既存ファイルを上書きしないパス
     */
    function getNonOverwritingFilePath(filePath) {
        var suffixIndex = 1;
        var pathParts = filePath.split(/(\.[^\.]+)$/);

        while (File(filePath).exists) {
            filePath = pathParts[0] + "(" + (++suffixIndex) + ")" + pathParts[1];
        }
        return filePath;
    }

    /**
     * 書き出し用の新規ドキュメントを作成する。
     * @param {string} documentTitle - ドキュメントの名前
     * @param {ImageColorSpace} imageColorSpace - 画像のカラースペース
     * @returns {Document} 作成したドキュメント
     */
    function createExportDocument(documentTitle, imageColorSpace) {
        var documentPreset = new DocumentPreset();
        var documentPresetType;

        documentPreset.title = documentTitle;
        documentPreset.width = 1000;
        documentPreset.height = 1000;

        if (imageColorSpace == ImageColorSpace.RGB) {
            documentPresetType = DocumentPresetType.BasicRGB;
            documentPreset.colorMode = DocumentColorSpace.RGB;
        } else {
            documentPresetType = DocumentPresetType.BasicCMYK;
            documentPreset.colorMode = DocumentColorSpace.CMYK;
        }

        return app.documents.addDocument(documentPresetType, documentPreset);
    }

    /**
     * ドキュメントをPSDとして書き出す。
     * @param {Document} exportDocument - 書き出すドキュメント
     * @param {string} exportFilePath - 書き出し先のパス
     * @param {ImageColorSpace} imageColorSpace - 画像のカラースペース
     * @param {number} resolution - 書き出し解像度（ppi）
     * @returns {File} 書き出したPSDファイル
     */
    function exportDocumentAsPSD(exportDocument, exportFilePath, imageColorSpace, resolution) {
        var exportedFile = File(exportFilePath);
        var psdOptions = new ExportOptionsPhotoshop();

        psdOptions.antiAliasing = false;
        psdOptions.artBoardClipping = true;
        psdOptions.imageColorSpace = imageColorSpace;
        psdOptions.editableText = false;
        psdOptions.flatten = true;
        psdOptions.maximumEditability = false;
        psdOptions.resolution = (resolution || 72);
        psdOptions.warnings = false;
        psdOptions.writeLayers = false;

        exportDocument.exportFile(exportedFile, ExportType.PHOTOSHOP, psdOptions);

        return exportedFile;
    }

    /**
     * 埋め込み画像を一時ドキュメントへ複製し、等倍・回転なしの状態でPSDに書き出す。
     * 元ファイルが不明な画像のフォールバックとして使う。
     * 再配置はアクションの置換（rplc）が行うため、ここでは変形を戻した素の画像だけを書き出す。
     * @author m1b
     * @discussion https://community.adobe.com/t5/illustrator-discussions/is-it-possible-to-convert-rasteritem-to-placeditem/m-p/13081172
     * @param {Document} doc - 対象ドキュメント
     * @param {RasterItem} item - 書き出す埋め込み画像
     * @param {string} baseName - 拡張子を除いたファイル名
     * @returns {File} 書き出したPSDファイル
     */
    function exportEmbeddedImageAsPSD(doc, item, baseName) {
        var linksFolder = getLinksFolder(doc);
        var imageColorSpace = item.imageColorSpace;
        var scaleAndRotation = getScaleAndRotation(item);

        /* 0で割ると変形行列が壊れるため、等倍に戻せない画像はここで中止 */
        if (!scaleAndRotation.scaleX || !scaleAndRotation.scaleY) {
            throw new Error("拡大率を取得できませんでした。");
        }

        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var exportDocument = createExportDocument(baseName, imageColorSpace);
        var exportedFile;

        try {
            var workingImage = item.duplicate(exportDocument.layers[0], ElementPlacement.PLACEATBEGINNING);

            /* 拡大率100%、回転角0°に戻してからアートボードを合わせる */
            var transformMatrix = app.getRotationMatrix(-scaleAndRotation.rotation);
            transformMatrix = app.concatenateScaleMatrix(transformMatrix,
                100 / scaleAndRotation.scaleX * 100,
                100 / scaleAndRotation.scaleY * 100);
            workingImage.transform(transformMatrix, true, true, true, true, true);
            workingImage.position = [0, workingImage.height];
            exportDocument.artboards[0].artboardRect = [0, workingImage.height, workingImage.width, 0];

            var exportFilePath = getNonOverwritingFilePath(linksFolder.fsName + "/" + baseName + ".psd");
            exportedFile = exportDocumentAsPSD(exportDocument, exportFilePath, imageColorSpace, EXPORT_RESOLUTION);

        } finally {
            exportDocument.close(SaveOptions.DONOTSAVECHANGES);
            app.userInteractionLevel = previousInteractionLevel;
            /* 一時ドキュメントを閉じたあと、確実に元のドキュメントへ戻す */
            app.activeDocument = doc;
        }

        if (!exportedFile.exists) {
            throw new Error("PSDを書き出せませんでした。");
        }
        return exportedFile;
    }

    // =========================================
    // パス表示 / Path display
    // =========================================

    /**
     * ホーム直下から「Dropbox」を含むフォルダーを探す。
     * チームフォルダー（「sw Dropbox」など）を優先する。
     * 個人用の「Dropbox」は「~/Dropbox/」として短縮できるため、優先度を下げている。
     * @returns {Folder|null} 見つかったフォルダー。なければnull
     */
    function findDropboxFolder() {
        var homeFolder = Folder("~");
        if (!homeFolder.exists) return null;

        var entryList = [];
        try { entryList = homeFolder.getFiles(); } catch (e) { return null; }

        var personalFolder = null;
        var teamFolder     = null;

        for (var i = 0; i < entryList.length; i++) {
            var entry = entryList[i];
            if (!(entry instanceof Folder)) continue;

            var entryName = safeDecodeURI(entry.name);
            if (entryName.charAt(0) === ".") continue;
            if (entryName.indexOf("Dropbox") === -1) continue;

            if (entryName === "Dropbox") {
                if (!personalFolder) personalFolder = entry;
            } else if (!teamFolder) {
                teamFolder = entry;
            }
        }
        return teamFolder ? teamFolder : personalFolder;
    }

    /**
     * フォルダー直下に表示用のサブフォルダーが1つだけあるとき、そのフォルダーを返す。
     * チームDropboxのメンバーフォルダー（「takano masahiro」など）の判定に使う。
     * @param {Folder} parentFolder - 探索するフォルダー
     * @returns {Folder|null} 唯一のサブフォルダー。0個または2個以上のときはnull
     */
    function findSingleSubFolder(parentFolder) {
        var entryList = [];
        try { entryList = parentFolder.getFiles(); } catch (e) { return null; }

        var foundFolder = null;
        for (var i = 0; i < entryList.length; i++) {
            var entry = entryList[i];
            if (!(entry instanceof Folder)) continue;
            if (safeDecodeURI(entry.name).charAt(0) === ".") continue;

            if (foundFolder) return null;
            foundFolder = entry;
        }
        return foundFolder;
    }

    /**
     * Dropboxのローカルマウントパスを決める。
     * 手動指定が空のときは、ホーム直下の「Dropbox」を含むフォルダーを探し、
     * その中にメンバーフォルダーが1つだけあれば、そこまでをプレフィックスとする。
     * @param {string} manualPath - 手動で指定するパス。空文字なら自動検出
     * @returns {string} 末尾に「/」を付けたプレフィックス。見つからない場合は空文字
     */
    function resolveDropboxPrefix(manualPath) {
        if (manualPath) {
            return (manualPath.charAt(manualPath.length - 1) === "/") ? manualPath : manualPath + "/";
        }

        var dropboxFolder = findDropboxFolder();
        if (!dropboxFolder) return "";

        var memberFolder = findSingleSubFolder(dropboxFolder);
        return (memberFolder ? memberFolder : dropboxFolder).fsName + "/";
    }

    /**
     * ホームフォルダー配下のパスを「~」始まりに置き換える。
     * @param {string} absolutePath - 絶対パス
     * @returns {string} 短縮したパス
     */
    function toTildePath(absolutePath) {
        if (!absolutePath) return absolutePath;

        var homePath = "";
        try { homePath = Folder("~").fsName; } catch (e) { }
        if (!homePath) return absolutePath;

        if (absolutePath === homePath) return "~";
        if (absolutePath.indexOf(homePath + "/") === 0) {
            return "~" + absolutePath.substring(homePath.length);
        }
        return absolutePath;
    }

    /**
     * 表示用のパス文字列を組み立てる。
     * @param {string} absolutePath - 絶対パス
     * @param {boolean} useFullPath - 絶対パスのまま表示するか
     * @param {boolean} useDropbox - Dropboxのプレフィックスを取り除くか
     * @returns {string} 表示用のパス
     */
    function formatDisplayPath(absolutePath, useFullPath, useDropbox) {
        if (!absolutePath) return absolutePath;
        if (useFullPath) return absolutePath;

        if (useDropbox && DROPBOX_PREFIX && absolutePath.indexOf(DROPBOX_PREFIX) === 0) {
            return absolutePath.substring(DROPBOX_PREFIX.length);
        }
        return toTildePath(absolutePath);
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA から読んだ SCRIPT_*（スクリプト情報の部品）を読み、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * SCRIPT_* の値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値
     */
    function getVersionInfoValue(variableKey) {
        var infoValue;
        if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
        else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
        else if (variableKey === "readme") {
            if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
            else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
        }
        else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
        else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 埋め込み画像の一覧をリストボックスへ反映する。
     * @param {ListBox} listBox - 反映先のリストボックス
     * @param {RasterItem[]} itemList - 表示する埋め込み画像
     * @param {boolean} useFullPath - 絶対パスのまま表示するか
     * @param {boolean} useDropbox - Dropboxのプレフィックスを取り除くか
     * @param {object} exportNameMap - uuidをキーとする書き出し名の対応表
     * @returns {void}
     */
    function updateFileList(listBox, itemList, useFullPath, useDropbox, exportNameMap) {
        listBox.removeAll();

        for (var i = 0; i < itemList.length; i++) {
            var item = itemList[i];
            var sourceFile = getEmbeddedSourceFile(item);

            var nameText, pathText;

            if (sourceFile) {
                nameText = safeDecodeURI(sourceFile.name);
                /* fsNameはデコード済みのパスなので、そのまま表示する / fsName is not URI-encoded */
                pathText = formatDisplayPath(sourceFile.parent.fsName, useFullPath, useDropbox);

            } else {
                /* 元ファイルが不明な画像は書き出し予定のファイル名を見せる / Show the planned export name */
                var exportName = exportNameMap[item.uuid];
                nameText = exportName ? (exportName + ".psd") : "（元ファイル不明）";
                pathText = "（PSDに書き出し）";
            }

            var row = listBox.add("item", nameText);
            row.subItems[0].text = pathText;
        }
    }

    /**
     * 処理オプションを選択するダイアログを表示する。
     * @param {RasterItem[]} selectedItems - 選択範囲内の埋め込み画像
     * @param {RasterItem[]} artboardItems - 現在のアートボード上の埋め込み画像
     * @param {RasterItem[]} allItems - ドキュメント内すべての埋め込み画像
     * @param {object} exportNameMap - uuidをキーとする書き出し名の対応表
     * @returns {object|null} 選択結果（{ scope: string, collect: boolean }）。キャンセル時はnull
     */
    function showOptionDialog(selectedItems, artboardItems, allItems, exportNameMap) {
        var hasSelectedRaster = (selectedItems.length > 0);

        var dialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(dialog);

        var scopePanel = dialog.add("panel", undefined, getLabel("panel.scope"));
        setupPanel(scopePanel, 6);

        var selectionRadio = scopePanel.add("radiobutton", undefined,
            labelWithCount("radio.selection", selectedItems.length));
        selectionRadio.helpTip = getLabel("tooltip.selection");
        var artboardRadio  = scopePanel.add("radiobutton", undefined,
            labelWithCount("radio.artboard", artboardItems.length));
        artboardRadio.helpTip = getLabel("tooltip.artboard");
        var allRadio       = scopePanel.add("radiobutton", undefined,
            labelWithCount("radio.all", allItems.length));
        allRadio.helpTip = getLabel("tooltip.all");

        /* 選択がなければ「すべて」を既定にする / Fall back to "all" when nothing is selected */
        selectionRadio.enabled = hasSelectedRaster;
        selectionRadio.value   = hasSelectedRaster;
        artboardRadio.enabled  = (artboardItems.length > 0);
        allRadio.value         = !hasSelectedRaster;

        var listPanel = dialog.add("panel", undefined, getLabel("panel.list"));
        setupPanel(listPanel);

        var fileList = listPanel.add("listbox", undefined, [], {
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [getLabel("column.fileName"), getLabel("column.path")],
            columnWidths: [200, 340]
        });
        fileList.preferredSize = [560, 220];

        var pathOptionGroup = listPanel.add("group");
        pathOptionGroup.alignment = "left";
        pathOptionGroup.spacing = 16;

        var fullPathCheck = pathOptionGroup.add("checkbox", undefined, getLabel("checkbox.fullPath"));
        fullPathCheck.helpTip = getLabel("tooltip.fullPath");
        var dropboxCheck  = pathOptionGroup.add("checkbox", undefined, getLabel("checkbox.dropbox"));
        dropboxCheck.helpTip = getLabel("tooltip.dropbox");

        fullPathCheck.value  = false;
        dropboxCheck.value   = (DROPBOX_PREFIX !== "");
        dropboxCheck.enabled = (DROPBOX_PREFIX !== "");

        var collectCheck = dialog.add("checkbox", undefined,
            getLabel("checkbox.collect").replace("{folder}", LINKS_FOLDER_NAME));
        collectCheck.helpTip = getLabel("tooltip.collect");
        collectCheck.value = true;

        /* 選択中の対象に対応する画像を返す / Return the items for the current scope */
        function getScopeItems() {
            if (selectionRadio.value) return selectedItems;
            if (artboardRadio.value)  return artboardItems;
            return allItems;
        }

        /* 一覧を現在の設定で描き直す / Redraw the list with the current settings */
        function refreshFileList() {
            /* Dropbox短縮中はフルパス表示を無効化 / Full path is meaningless while shortening */
            fullPathCheck.enabled = !dropboxCheck.value;
            if (!fullPathCheck.enabled) fullPathCheck.value = false;

            updateFileList(fileList,
                getScopeItems(),
                fullPathCheck.value,
                dropboxCheck.value,
                exportNameMap);
        }

        selectionRadio.onClick = refreshFileList;
        artboardRadio.onClick  = refreshFileList;
        allRadio.onClick       = refreshFileList;
        fullPathCheck.onClick  = refreshFileList;
        dropboxCheck.onClick   = refreshFileList;
        refreshFileList();

        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, "キャンセル", { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });

        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(dialog);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        if (dialog.show() !== 1) return null;
        return {
            scope: selectionRadio.value ? "selection" : (artboardRadio.value ? "artboard" : "all"),
            collect: collectCheck.value
        };
    }

    // =========================================
    // 選択・リンク先の判定 / Selection and target resolution
    // =========================================

    /**
     * 配列やコレクションから埋め込み画像を再帰的に集める。
     * @param {Array|PageItems} itemList - 走査対象のアイテム群
     * @param {RasterItem[]} resultList - 収集先の配列
     * @returns {void}
     */
    function collectRasterItems(itemList, resultList) {
        for (var i = 0; i < itemList.length; i++) {
            var item = itemList[i];
            if (item.typename === "RasterItem") {
                resultList.push(item);
            } else if (item.typename === "GroupItem") {
                collectRasterItems(item.pageItems, resultList);
            }
        }
    }

    /**
     * 選択範囲に含まれる埋め込み画像を取得する。
     * @param {Document} doc - 対象ドキュメント
     * @returns {RasterItem[]} 埋め込み画像の配列
     */
    function getSelectedRasterItems(doc) {
        var sel = doc.selection;
        if (!sel || sel.length === 0) return [];

        var resultList = [];
        collectRasterItems(sel, resultList);
        return resultList;
    }

    /**
     * ドキュメント内のすべての埋め込み画像を取得する。
     * 処理中にコレクションが変化するため、配列へ写し取ってから返す。
     * @param {Document} doc - 対象ドキュメント
     * @returns {RasterItem[]} 埋め込み画像の配列
     */
    function getAllRasterItems(doc) {
        var resultList = [];
        for (var i = 0; i < doc.rasterItems.length; i++) {
            resultList.push(doc.rasterItems[i]);
        }
        return resultList;
    }

    /**
     * 2つの矩形が重なっているかを判定する。
     * @param {number[]} rectA - 矩形1（[left, top, right, bottom]）
     * @param {number[]} rectB - 矩形2（[left, top, right, bottom]）
     * @returns {boolean} 重なっている場合はtrue
     */
    function rectsOverlap(rectA, rectB) {
        return rectA[0] < rectB[2]
            && rectA[2] > rectB[0]
            && rectA[1] > rectB[3]
            && rectA[3] < rectB[1];
    }

    /**
     * 現在のアートボードと重なる埋め込み画像を取得する。
     * 判定には効果を含まないgeometricBoundsを使う。
     * @param {Document} doc - 対象ドキュメント
     * @returns {RasterItem[]} 埋め込み画像の配列
     */
    function getArtboardRasterItems(doc) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var resultList = [];

        for (var i = 0; i < doc.rasterItems.length; i++) {
            var item = doc.rasterItems[i];
            if (rectsOverlap(item.geometricBounds, artboardRect)) resultList.push(item);
        }
        return resultList;
    }

    /**
     * 埋め込み画像から元ファイルを取得する。
     * @param {RasterItem} item - 対象の埋め込み画像
     * @returns {File|null} 元ファイル。参照できない場合はnull
     */
    function getEmbeddedSourceFile(item) {
        /* 埋め込み画像では file プロパティの参照自体が失敗することがある */
        try {
            if (item.file && item.file.exists) return item.file;
        } catch (e) { }
        return null;
    }

    /**
     * 元ファイルが不明なとき、ユーザーに再リンク先を選ばせる。
     * @returns {File|null} 選ばれたファイル。選択されなかった場合はnull
     */
    function promptForTargetFile() {
        var userChoice = confirm(
            "元のファイル情報が残っていないか、ファイルが見つかりません。\n" +
            "手動でファイルを選択して再リンクしますか？"
        );
        return userChoice ? File.openDialog("再リンクする画像ファイルを選択してください") : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 埋め込み画像1件を再リンクし、必要なら収集する。
     * 再リンクの失敗は例外、収集の失敗は戻り値の collectError で伝える。
     * @param {Document} doc - 対象ドキュメント
     * @param {RasterItem} item - 対象の埋め込み画像
     * @param {File} targetFile - リンク先の画像ファイル
     * @param {boolean} shouldCollect - 収集するかどうか
     * @param {object} status - 実行状況を書き戻すオブジェクト（{ actionPlayed: boolean }）
     * @returns {object} 処理結果（{ placedItem: PlacedItem, linkedFile: File, collectError: string }）
     */
    function processRasterItem(doc, item, targetFile, shouldCollect, status) {
        var newPlaced;

        try {
            newPlaced = relinkByAction(doc, item, targetFile, status);
        } catch (e) {
            throw new Error("再リンクに失敗: " + e.message);
        }

        var result = { placedItem: newPlaced, linkedFile: targetFile, collectError: "" };

        if (shouldCollect) {
            /* 再リンク自体は成功しているため、収集の失敗で結果を失わない */
            try {
                result.linkedFile = collectLink(doc, newPlaced, targetFile);
            } catch (e) {
                result.collectError = "収集に失敗: " + e.message;
            }
        }

        return result;
    }

    (function () {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var doc = app.activeDocument;

        var selectedItems = getSelectedRasterItems(doc);
        var artboardItems = getArtboardRasterItems(doc);
        var allItems      = getAllRasterItems(doc);

        if (allItems.length === 0) {
            alert("ドキュメントに埋め込み画像が見つかりません。");
            return;
        }

        /* 元ファイル不明時の書き出し名は、置き換えを始める前にまとめて決めておく */
        var exportNameMap = buildExportNameMap(doc);

        var options = showOptionDialog(selectedItems, artboardItems, allItems, exportNameMap);
        if (!options) return;

        /* 入れ子の三項演算子は ExtendScript で誤評価されるため if/else で書く / nested ternaries misevaluate in ExtendScript */
        var targetItems = allItems;
        if (options.scope === "selection") {
            targetItems = selectedItems;
        } else if (options.scope === "artboard") {
            targetItems = artboardItems;
        }

        var isSingle     = (targetItems.length === 1);
        var placedList   = [];
        var successCount = 0;
        var skipList     = [];
        var warningList  = [];
        var errorList    = [];

        for (var i = 0; i < targetItems.length; i++) {
            var item = targetItems[i];
            var itemName = item.name || ("画像 " + (i + 1));

            var targetFile = getEmbeddedSourceFile(item);
            var isExported = false;

            /* 元ファイルが不明なら、埋め込み画像自体をPSDに書き出してリンク先にする */
            if (!targetFile) {
                var failureText = "";
                var isSkipped   = false;

                if (!isSupportedColorSpace(item.imageColorSpace)) {
                    failureText = "未対応のカラースペース";
                    isSkipped   = true;

                } else {
                    try {
                        targetFile = exportEmbeddedImageAsPSD(doc, item,
                            exportNameMap[item.uuid] || ("image" + (i + 1)));
                        isExported = true;
                    } catch (e) {
                        failureText = "PSD書き出しに失敗: " + e.message;
                    }
                }

                /* 書き出せず1件のみのときは、手動選択を促す */
                if (!targetFile && isSingle) targetFile = promptForTargetFile();

                if (!targetFile) {
                    if (isSkipped) skipList.push(itemName + "（" + failureText + "）");
                    else           errorList.push(itemName + "：" + failureText);
                    continue;
                }
            }

            var status = { actionPlayed: false };

            try {
                /* 書き出し済みのPSDはすでに収集先にあるため、収集処理は不要 */
                var result = processRasterItem(doc, item, targetFile, options.collect && !isExported, status);
                placedList.push(result.placedItem);
                successCount++;

                if (result.collectError) warningList.push(itemName + "：" + result.collectError);

            } catch (e) {
                /* 配置アクションの実行前に失敗した場合だけ、未使用のPSDを片付ける */
                if (isExported && !status.actionPlayed && targetFile.exists) targetFile.remove();
                errorList.push(itemName + "：" + e.message);
            }
        }

        /* 新しいリンク画像を選択状態にし、画面を更新してから結果を表示 */
        doc.selection = null;
        for (var j = 0; j < placedList.length; j++) {
            placedList[j].selected = true;
        }
        app.redraw();

        var messageList = ["リンク画像への切り替えが完了しました。",
            "",
            "成功: " + successCount + " 件",
            "スキップ: " + skipList.length + " 件",
            "失敗: " + errorList.length + " 件"];

        if (skipList.length > 0)    messageList.push("", "【スキップ】", skipList.join("\n"));
        if (warningList.length > 0) messageList.push("", "【警告】", warningList.join("\n"));
        if (errorList.length > 0)   messageList.push("", "【失敗】", errorList.join("\n"));

        alert(messageList.join("\n"));
    })();

})();
