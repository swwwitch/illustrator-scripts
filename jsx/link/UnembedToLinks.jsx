#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

埋め込み画像（RasterItem）を、一時アクション（adobe_placeDocument）を動的に生成して実行し、
リンク画像に置き換えるスクリプトです。「選択オブジェクトと置換」で配置するため、
位置・サイズ・回転角・重ね順はIllustrator側がそのまま引き継ぎます。

### 主な機能

- 処理対象をダイアログで選択
  - 選択している画像のみ（グループの中も再帰的に探索）
  - 現在のアートボード上の埋め込み画像（アクティブなアートボードと重なる画像）
  - すべての埋め込み画像
- 対象ファイルの一覧表示（ファイル名／パス）
  - ［フルパス］絶対パスのまま表示
  - ［Dropboxパスを短縮］DROPBOX_PREFIX以下を相対表示。オフのときはホーム以下を「~」に短縮
- 元ファイルが判明している画像はそのファイルへリンク
- 元ファイルが不明な画像は「Links」フォルダーへPSDで書き出してリンク
  - 書き出し名は、レイヤー名 → 拡張子付きの親グループ名 → XMPマニフェスト → image1、image2… の順に探索
  - 同名ファイルがある場合は連番の接尾辞を付けるため、既存ファイルを上書きしない
- ［再リンク後に収集］リンク先をドキュメントと同階層の「Links」フォルダーへコピー
  - 同名で内容が異なるファイルには連番を付け、同一とみなせるファイルは再コピーしない
- 処理結果（成功・スキップ・失敗の件数と明細）をまとめて表示

### 処理の流れ

1. 対象の埋め込み画像を集め、元ファイル不明時の書き出し名をまとめて決める
2. 元ファイルを取得。取得できない場合は一時ドキュメントへ複製し、
   拡大率100%・回転角0°に戻してPSDとして書き出す
3. 対象の画像だけを選択し、一時アクションを生成して実行（実行後にアクションは必ず削除）
4. 収集がオンなら「Links」フォルダーへコピーしてリンクを張り替える
5. 置き換え後のリンク画像を選択状態にして結果を表示

### 注意

- ドキュメントが未保存の場合、「Links」フォルダーの場所を決められないため収集・書き出しに失敗します。
- ロックまたは非表示のレイヤー・グループ・画像は置換できません（理由を明示して失敗として報告します）。
- PSD書き出しはCMYK／RGB／グレースケールのみ対応です。
- 効果（ドロップシャドウなど）はPSD書き出し時にラスタライズされて含まれます。

*/

/*

### Overview

Replaces embedded raster images with linked images by generating and playing a temporary
action (adobe_placeDocument). Because the action places the file with "replace selection",
Illustrator itself preserves the position, size, rotation and stacking order.

### Features

- Scope chosen in a dialog
  - Selected images only (groups are searched recursively)
  - Embedded images on the current artboard
  - All embedded images
- Target file list (file name / path)
  - [Full path] shows the absolute path
  - [Shorten Dropbox path] hides DROPBOX_PREFIX; when off, the home folder becomes "~"
- Images whose original file is known are linked to that file
- Images with an unknown original are exported as a PSD into the "Links" folder and linked
  - The name is looked up in the layer name, the parent group name, the XMP manifest,
    then falls back to image1, image2…
  - An incremental suffix is added when a file of the same name exists
- [Collect after relinking] copies the linked file into the "Links" folder next to the document
  - A numbered suffix is added for different files of the same name; identical files are not copied again
- Reports the result (success / skipped / failed counts and details) in a single alert

### Flow

1. Collect the target images and resolve the export names up front
2. Get the original file; when unavailable, duplicate the image into a temporary document,
   reset the scale to 100% and the rotation to 0°, then export a PSD
3. Select only the target image and play the temporary action (always removed afterwards)
4. Copy the file into the "Links" folder and repoint the link when collecting is enabled
5. Select the resulting linked images and show the summary

### Notes

- Collecting and exporting fail when the document has not been saved.
- Locked or hidden layers, groups and items cannot be replaced (the reason is reported).
- PSD export supports CMYK / RGB / Grayscale only.
- Effects such as drop shadows are rasterized into the exported PSD.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "UnembedToLinks";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/UnembedToLinks.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/UnembedToLinks.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6d9e2dabb054"; /* 紹介記事 / article URL */

// 埋め込み画像のPSD書き出し処理は次のコードを参照 / The PSD export logic is based on:
// @author m1b
// @discussion https://community.adobe.com/t5/illustrator-discussions/is-it-possible-to-convert-rasteritem-to-placeditem/m-p/13081172

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// 設定 / Settings
// =========================================

var ACTION_SET_NAME    = "UnembedToLinksTempSet";    /* 一時アクションセット名 / temporary action set name */
var ACTION_NAME        = "UnembedToLinksPlace";      /* 一時アクション名 / temporary action name */
var ACTION_FILE_NAME   = "~/UnembedToLinksTemp.aia"; /* 一時アクションファイル / temporary action file */
var LINKS_FOLDER_NAME  = "Links";                    /* 収集先フォルダー名 / collect destination folder */

/* 元ファイルが不明なときのPSD書き出し設定 / PSD export fallback */
var EXPORT_RESOLUTION  = 72;                         /* 書き出し解像度（ppi） / export resolution */
var USE_XMP_NAMES      = true;                       /* XMPマニフェストの元ファイル名を使う / use names from XMP manifest */

/* Dropboxのローカルマウントパス。空文字にすると短縮機能は無効 / Local Dropbox mount path */
var DROPBOX_PREFIX     = "/Users/takano/sw Dropbox/takano masahiro/";

// =========================================
// 文字列エンコード / String encoding
// =========================================

/**
 * URIエンコードされた文字列をデコードする。
 * File.name のように「%」が含まれうる値でも例外を投げない。
 * @param {string} sourceText - デコードする文字列
 * @returns {string} デコード結果。不正なエスケープを含む場合は元の文字列
 */
function safeDecodeURI(sourceText) {
    try {
        return decodeURI(sourceText);
    } catch (e) {
        return String(sourceText);
    }
}

/**
 * 文字列をUTF-8バイト列に変換する。
 * @param {string} sourceText - 変換元の文字列
 * @returns {number[]} UTF-8バイト値の配列
 */
function stringToUtf8Bytes(sourceText) {
    var encodedText = encodeURIComponent(sourceText);
    var byteList = [];

    for (var i = 0; i < encodedText.length; i++) {
        var currentChar = encodedText.charAt(i);
        if (currentChar === "%") {
            byteList.push(parseInt(encodedText.substr(i + 1, 2), 16));
            i += 2;
        } else {
            byteList.push(currentChar.charCodeAt(0));
        }
    }
    return byteList;
}

/**
 * バイト列を16進文字列に変換し、記録された.aiaと同じく32バイトごとに改行する。
 * @param {number[]} byteList - バイト値の配列
 * @param {string} indentText - 各行の先頭に付けるインデント
 * @returns {string} 改行区切りの16進文字列
 */
function bytesToHexLines(byteList, indentText) {
    var hexText = "";
    for (var i = 0; i < byteList.length; i++) {
        hexText += (byteList[i] < 16 ? "0" : "") + byteList[i].toString(16);
    }

    var lineList = [];
    for (var start = 0; start < hexText.length; start += 64) {
        lineList.push(indentText + hexText.substr(start, 64));
    }
    return lineList.join("\n");
}

/**
 * .aia のテキスト値（`[ バイト数 16進 ]`）を組み立てる。
 * @param {string} sourceText - 埋め込む文字列
 * @param {string} indentText - 閉じ括弧と16進行のインデント
 * @returns {string} .aia形式のテキスト値
 */
function buildTextValue(sourceText, indentText) {
    var byteList = stringToUtf8Bytes(sourceText);
    if (byteList.length === 0) {
        return "[ 0 ]";
    }
    return "[ " + byteList.length + "\n"
        + bytesToHexLines(byteList, indentText + "\t") + "\n"
        + indentText + "]";
}

// =========================================
// 一時アクション生成 / Temporary action generation
// =========================================

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
// 一時アクション実行 / Temporary action playback
// =========================================

/**
 * 一時アクションを書き出して読み込み、実行後に必ず後片付けする。
 * File の close / remove は失敗時に false を返すだけなので try で囲まない。
 * @param {string} actionSource - .aia形式のアクションソース
 * @param {string} setName - アクションセット名
 * @param {string} actionName - アクション名
 * @param {string} actionFilePath - 一時アクションファイルのパス
 * @returns {void}
 */
function playTemporaryAction(actionSource, setName, actionName, actionFilePath) {
    var actionFile = new File(actionFilePath);

    /* 同名セットが残っていれば先に破棄 / Unload a leftover set with the same name */
    try { app.unloadAction(setName, ""); } catch (e) { }

    try {
        actionFile.lineFeed = "Unix";
        if (!actionFile.open("w")) {
            throw new Error("一時アクションファイルを作成できませんでした。");
        }
        actionFile.write(actionSource);
        actionFile.close();

        app.loadAction(actionFile);
        app.doScript(actionName, setName, false);

    } finally {
        actionFile.close();
        if (actionFile.exists) actionFile.remove();
        try { app.unloadAction(setName, ""); } catch (e) { }
    }
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
    playTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME, ACTION_FILE_NAME);

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

    var dialog = new Window("dialog", "埋め込み画像をリンクに変換");
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = 16;
    dialog.spacing = 12;

    var scopePanel = dialog.add("panel", undefined, "対象");
    scopePanel.orientation = "column";
    scopePanel.alignChildren = "left";
    scopePanel.margins = [12, 16, 12, 12];
    scopePanel.spacing = 8;

    var selectionRadio = scopePanel.add("radiobutton", undefined,
        "選択している画像のみ（" + selectedItems.length + " 件）");
    var artboardRadio  = scopePanel.add("radiobutton", undefined,
        "現在のアートボード上の埋め込み画像（" + artboardItems.length + " 件）");
    var allRadio       = scopePanel.add("radiobutton", undefined,
        "すべての埋め込み画像（" + allItems.length + " 件）");

    /* 選択がなければ「すべて」を既定にする / Fall back to "all" when nothing is selected */
    selectionRadio.enabled = hasSelectedRaster;
    selectionRadio.value   = hasSelectedRaster;
    artboardRadio.enabled  = (artboardItems.length > 0);
    allRadio.value         = !hasSelectedRaster;

    var listPanel = dialog.add("panel", undefined, "対象ファイル");
    listPanel.orientation = "column";
    listPanel.alignChildren = "fill";
    listPanel.margins = [12, 16, 12, 12];

    var fileList = listPanel.add("listbox", undefined, [], {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: ["ファイル名", "パス"],
        columnWidths: [200, 340]
    });
    fileList.preferredSize = [560, 220];

    var pathOptionGroup = listPanel.add("group");
    pathOptionGroup.alignment = "left";
    pathOptionGroup.spacing = 16;

    var fullPathCheck = pathOptionGroup.add("checkbox", undefined, "フルパス");
    var dropboxCheck  = pathOptionGroup.add("checkbox", undefined, "Dropboxパスを短縮");

    fullPathCheck.value  = false;
    dropboxCheck.value   = (DROPBOX_PREFIX !== "");
    dropboxCheck.enabled = (DROPBOX_PREFIX !== "");

    var collectCheck = dialog.add("checkbox", undefined,
        "再リンク後に収集（同階層の「" + LINKS_FOLDER_NAME + "」フォルダーへコピー）");
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

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

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
