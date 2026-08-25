#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// 設定 / Settings
// =========================================

var ACTION_SET_NAME    = "Embed2RelinkTempSet";      /* 一時アクションセット名 / temporary action set name */
var ACTION_NAME        = "Embed2RelinkPlace";        /* 一時アクション名 / temporary action name */
var ACTION_FILE_NAME   = "~/Embed2RelinkTemp.aia";   /* 一時アクションファイル / temporary action file */
var LINKS_FOLDER_NAME  = "Links";                    /* 収集先フォルダー名 / collect destination folder */

/* Dropboxのローカルマウントパス。空文字にすると短縮機能は無効 / Local Dropbox mount path */
var DROPBOX_PREFIX     = "/Users/takano/sw Dropbox/takano masahiro/";

// =========================================
// 文字列エンコード / String encoding
// =========================================

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
 * 埋め込み画像（RasterItem）を、ダイナミックアクション経由でリンク画像に置き換える。
 * 選択オブジェクトの置換（rplc）で実行するため、位置・サイズ・角度・重なり順は
 * アクション側が引き継ぐ。スクリプト側での復元処理は行わない。
 * @param {Document} doc - 対象ドキュメント
 * @param {RasterItem} item - 置き換え対象の埋め込み画像
 * @param {File} targetFile - リンク先の画像ファイル
 * @returns {PlacedItem} 置換後のリンク画像
 */
function relinkByAction(doc, item, targetFile) {
    var layer = item.layer;

    if (layer.locked || !layer.visible) {
        throw new Error("対象のレイヤーがロックまたは非表示です。");
    }
    if (item.locked || item.hidden) {
        throw new Error("対象の画像がロックまたは非表示です。");
    }

    /* 置換対象だけを選択した状態で実行 / Select only the replacement target */
    doc.activeLayer = layer;
    doc.selection = null;
    item.selected = true;

    var actionSource = buildActionSource(ACTION_SET_NAME, ACTION_NAME, targetFile.fsName);
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
    var fileName  = decodeURI(sourceFile.name);
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
 * リンク先ファイルをドキュメントと同階層の「Links」フォルダーへ複製し、リンクを張り替える。
 * @param {Document} doc - 対象ドキュメント
 * @param {PlacedItem} placedItem - 張り替え対象のリンク画像
 * @param {File} sourceFile - 収集元のファイル
 * @returns {File} 収集後のリンク先ファイル
 */
function collectLink(doc, placedItem, sourceFile) {
    var docFile = doc.fullName;
    if (!docFile || !docFile.exists) {
        throw new Error("ドキュメントが保存されていないため、収集先を決定できません。");
    }

    var linksFolder = new Folder(docFile.parent.fsName + "/" + LINKS_FOLDER_NAME);
    if (!linksFolder.exists && !linksFolder.create()) {
        throw new Error("「" + LINKS_FOLDER_NAME + "」フォルダーを作成できませんでした。");
    }

    var destFile = resolveCollectDestination(linksFolder, sourceFile);
    if (!destFile.exists && !sourceFile.copy(destFile.fsName)) {
        throw new Error("リンクファイルを複製できませんでした。");
    }

    placedItem.file = destFile;
    return destFile;
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
 * @returns {void}
 */
function updateFileList(listBox, itemList, useFullPath, useDropbox) {
    listBox.removeAll();

    for (var i = 0; i < itemList.length; i++) {
        var sourceFile = getEmbeddedSourceFile(itemList[i]);
        var row = listBox.add("item",
            sourceFile ? decodeURI(sourceFile.name) : "（元ファイル不明）");
        row.subItems[0].text = sourceFile
            ? formatDisplayPath(decodeURI(sourceFile.parent.fsName), useFullPath, useDropbox)
            : "-";
    }
}

/**
 * 処理オプションを選択するダイアログを表示する。
 * @param {RasterItem[]} selectedItems - 選択範囲内の埋め込み画像
 * @param {RasterItem[]} allItems - ドキュメント内すべての埋め込み画像
 * @returns {object|null} 選択結果（{ scope: string, collect: boolean }）。キャンセル時はnull
 */
function showOptionDialog(selectedItems, allItems) {
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
    var allRadio       = scopePanel.add("radiobutton", undefined,
        "すべての埋め込み画像（" + allItems.length + " 件）");

    /* 選択がなければ「すべて」を既定にする / Fall back to "all" when nothing is selected */
    selectionRadio.enabled = hasSelectedRaster;
    selectionRadio.value   = hasSelectedRaster;
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

    /* 一覧を現在の設定で描き直す / Redraw the list with the current settings */
    function refreshFileList() {
        /* Dropbox短縮中はフルパス表示を無効化 / Full path is meaningless while shortening */
        fullPathCheck.enabled = !dropboxCheck.value;
        if (!fullPathCheck.enabled) fullPathCheck.value = false;

        updateFileList(fileList,
            selectionRadio.value ? selectedItems : allItems,
            fullPathCheck.value,
            dropboxCheck.value);
    }

    selectionRadio.onClick = refreshFileList;
    allRadio.onClick       = refreshFileList;
    fullPathCheck.onClick  = refreshFileList;
    dropboxCheck.onClick   = refreshFileList;
    refreshFileList();

    var collectCheck = dialog.add("checkbox", undefined,
        "再リンク後に収集（同階層の「" + LINKS_FOLDER_NAME + "」フォルダーへコピー）");
    collectCheck.value = true;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (dialog.show() !== 1) return null;
    return {
        scope: selectionRadio.value ? "selection" : "all",
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
 * 再リンク先のファイルを決定する。元ファイルが不明な場合はダイアログで選択させる。
 * @param {RasterItem} item - 対象の埋め込み画像
 * @returns {File|null} 再リンク先のファイル。選択されなかった場合はnull
 */
function resolveTargetFile(item) {
    var sourceFile = getEmbeddedSourceFile(item);
    if (sourceFile) return sourceFile;

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
 * @param {Document} doc - 対象ドキュメント
 * @param {RasterItem} item - 対象の埋め込み画像
 * @param {File} targetFile - リンク先の画像ファイル
 * @param {boolean} shouldCollect - 収集するかどうか
 * @returns {object} 処理結果（{ placedItem: PlacedItem, linkedFile: File }）
 */
function processRasterItem(doc, item, targetFile, shouldCollect) {
    var isRelinked = false;

    try {
        var newPlaced = relinkByAction(doc, item, targetFile);
        isRelinked = true;

        var linkedFile = shouldCollect
            ? collectLink(doc, newPlaced, targetFile)
            : targetFile;

        return { placedItem: newPlaced, linkedFile: linkedFile };

    } catch (e) {
        throw new Error((isRelinked ? "収集" : "再リンク") + "に失敗: " + e.message);
    }
}

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    var selectedItems = getSelectedRasterItems(doc);
    var allItems      = getAllRasterItems(doc);

    if (allItems.length === 0) {
        alert("ドキュメントに埋め込み画像が見つかりません。");
        return;
    }

    var options = showOptionDialog(selectedItems, allItems);
    if (!options) return;

    var targetItems = (options.scope === "all") ? allItems : selectedItems;

    var isSingle    = (targetItems.length === 1);
    var placedList  = [];
    var successList = [];
    var skipList    = [];
    var errorList   = [];

    for (var i = 0; i < targetItems.length; i++) {
        var item = targetItems[i];
        var itemName = decodeURI(item.name) || ("画像 " + (i + 1));

        /* 1件のみのときは、元ファイル不明でも手動選択を促す */
        var targetFile = isSingle ? resolveTargetFile(item) : getEmbeddedSourceFile(item);
        if (!targetFile) {
            if (isSingle) return;
            skipList.push(itemName + "（元ファイル不明）");
            continue;
        }

        try {
            var result = processRasterItem(doc, item, targetFile, options.collect);
            placedList.push(result.placedItem);
            successList.push(itemName + " → " + result.linkedFile.fsName);
        } catch (e) {
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
        "成功: " + successList.length + " 件",
        "スキップ: " + skipList.length + " 件",
        "失敗: " + errorList.length + " 件"];

    if (successList.length > 0) messageList.push("", "【再リンク先】", successList.join("\n"));
    if (skipList.length > 0)    messageList.push("", "【スキップ】", skipList.join("\n"));
    if (errorList.length > 0)   messageList.push("", "【失敗】", errorList.join("\n"));

    alert(messageList.join("\n"));
})();
