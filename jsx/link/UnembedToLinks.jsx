#target illustrator
#targetengine "UnembedToLinksEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

埋め込み画像を、一時アクションを動的に生成して実行することでリンク画像に置き換えます。
「選択オブジェクトと置換」で配置するため、位置・サイズ・回転角・重ね順はそのまま引き継がれます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/UnembedToLinks.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n6d9e2dabb054

### Overview

Replaces embedded images with linked images by generating and running a temporary action on the fly.
Because the placement replaces the selected object, position, size, rotation and stacking order all carry over unchanged.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/UnembedToLinks.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "UnembedToLinks";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/UnembedToLinks.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/UnembedToLinks.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6d9e2dabb054"; /* 紹介記事 / article URL */

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
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
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

        var targetItems = (options.scope === "selection") ? selectedItems
            : (options.scope === "artboard") ? artboardItems
            : allItems;

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
