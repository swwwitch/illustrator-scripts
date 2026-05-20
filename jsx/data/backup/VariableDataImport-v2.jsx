#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  VariableDataImport_v12.jsx

  更新日：20260122

  仕様：
  1. 実行時にドキュメントを複製（別名保存）し、複製したファイルに対して流し込みを実行します。アートボード指定時：複製したドキュメント内でアートボードを増やして流し込み。選択グループ指定時：アートボードは増やさず、選択中のグループを同一アートボード内に複製して流し込み。
  2. CSV(.csv)とTSV(.txt)の両方に対応（カンマや引用符も解析）。
  3. 1件目のデータを最後に処理することで、タグ消失バグを回避。
  4. 進捗バーを表示。
  5. 流し込み対象（アートボード / 選択グループ）を選べる（起動時にグループ選択中なら自動で「選択グループ」を選択）。選択グループ時は「アートボード名に使用する項目」をディム表示。
  6. 「プレビュー」をオンにすると、複製ドキュメントへ流し込んだ結果を表示（設定変更のたび自動更新、ダイアログは開いたまま）。
*/

// =========================================
// バージョンとローカライズ
// =========================================

var SCRIPT_VERSION = "v1.3";

/* ロケール判定 / Detect locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "データ流し込み", en: "Data Import" },
    fileLabel: { ja: "ファイル:", en: "File:" },
    settingsTitle: { ja: "流し込み設定", en: "Import Settings" },
    abNameLabel: { ja: "アートボード名に使用する項目:", en: "Field for artboard name:" },
    colsLabel: { ja: "横に並べる数:", en: "Columns:" },
    wrapLabel: { ja: "列で折り返し", en: "wrap" },
    targetLabel: { ja: "流し込み対象:", en: "Target:" },
    targetArtboard: { ja: "アートボード", en: "Artboards" },
    targetSelection: { ja: "選択グループ", en: "Selected group(s)" },
    previewLabel: { ja: "プレビュー（複製ドキュメントに表示）", en: "Preview (in a duplicate document)" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    run: { ja: "実行", en: "Run" },
    processing: { ja: "処理中...", en: "Processing..." },
    done: { ja: "完了！\n#count# 件処理しました。", en: "Done!\nProcessed #count# rows." },
    noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    needSave: { ja: "ドキュメントを保存してから実行してください。", en: "Please save the document before running." },
    noDataFiles: { ja: "同じ階層に.txtまたは.csvファイルが見つかりませんでした。", en: "No .txt or .csv files found in the same folder." },
    noTemplate: { ja: "雛形にオブジェクトがありません。", en: "No template objects found." },
    needGroupSelect: { ja: "選択グループを対象にするには、グループを選択してください。", en: "To use Selected group(s), please select a group." },
    onlyGroupSelect: { ja: "『選択グループ』を選んだ場合は、グループ（GroupItem）のみを選択してください。\n現在の選択にグループ以外が含まれています。", en: "When using Selected group(s), select GroupItem only.\nYour selection contains non-group items." },
    dupFailed: { ja: "ドキュメントの複製（別名保存）に失敗しました。\n\n#detail#", en: "Failed to duplicate (Save As) the document.\n\n#detail#" },
    artboardLimit: { ja: "アートボードをこれ以上追加できませんでした。\nIllustratorのカンバス上限に達した可能性があります。\n「横に並べる数」を減らして実行し直してください。\n\n#detail#", en: "Could not add any more artboards.\nThe Illustrator canvas limit may have been reached.\nReduce \"Columns\" and run again.\n\n#detail#" }
};

/* ラベル取得 / Get localized label */
function L(key) {
    try {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
    } catch (e) { }
    return key;
}

(function () {
    // =========================================
    // 事前チェック
    // =========================================

    if (app.documents.length === 0) {
        alert(L("noDoc"));
        return;
    }

    var doc = app.activeDocument;
    var docFolderPath;
    try {
        docFolderPath = doc.path;
    } catch (e) {
        alert(L("needSave"));
        return;
    }

    var docFolder = new Folder(docFolderPath);
    var dataFiles = docFolder.getFiles(/\.(txt|csv)$/i);
    if (dataFiles.length === 0) {
        alert(L("noDataFiles"));
        return;
    }

    // =========================================
    // メインダイアログ
    // =========================================

    var mainDialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    mainDialog.orientation = "column";
    mainDialog.alignChildren = ["fill", "top"];
    mainDialog.spacing = 10;
    mainDialog.margins = 15;

    /* ファイル選択行 / File selection row */
    var fileSelectGroup = mainDialog.add("group");
    fileSelectGroup.add("statictext", undefined, L("fileLabel"));
    var fileDropdown = fileSelectGroup.add("dropdownlist", undefined, []);
    fileDropdown.size = [350, 25];
    for (var i = 0; i < dataFiles.length; i++) fileDropdown.add("item", decodeURI(dataFiles[i].name));

    /* プレビュー用リストボックスの置き場 / Container for the preview listbox */
    var previewListContainer = mainDialog.add("group");
    previewListContainer.alignment = ["fill", "fill"];
    var previewListBox = null;

    /* 流し込み設定パネル / Import settings panel */
    var settingsPanel = mainDialog.add("panel", undefined, L("settingsTitle"));
    settingsPanel.orientation = "column";
    settingsPanel.alignChildren = ["left", "top"];
    settingsPanel.margins = 15;

    var artboardNameFieldGroup = settingsPanel.add("group");
    artboardNameFieldGroup.add("statictext", undefined, L("abNameLabel"));
    var artboardNameFieldDropdown = artboardNameFieldGroup.add("dropdownlist", undefined, []);
    artboardNameFieldDropdown.size = [200, 25];

    /* 対象モードに応じてUIを更新 / Update UI for the current target mode */
    function updateUiByTargetMode() {
        // 選択グループ時は、アートボード名の指定は不要なのでディム表示
        artboardNameFieldGroup.enabled = (targetMode !== "selection");
    }

    var wrapColsGroup = settingsPanel.add("group");
    wrapColsGroup.add("statictext", undefined, L("colsLabel"));
    var wrapColsInput = wrapColsGroup.add("edittext", undefined, "4");
    wrapColsInput.size = [50, 25];
    wrapColsGroup.add("statictext", undefined, L("wrapLabel"));

    /* 流し込み対象（アートボード / 選択グループ）/ Target mode radios */
    var targetModeGroup = settingsPanel.add("group");
    targetModeGroup.add("statictext", undefined, L("targetLabel"));
    var targetArtboardRadio = targetModeGroup.add("radiobutton", undefined, L("targetArtboard"));
    var targetSelectionRadio = targetModeGroup.add("radiobutton", undefined, L("targetSelection"));

    /* プレビュー切り替え / Preview toggle */
    var previewGroup = settingsPanel.add("group");
    var previewCheckbox = previewGroup.add("checkbox", undefined, L("previewLabel"));

    var targetMode = "artboard"; // "artboard" | "selection"

    /* 起動時の選択状態から対象モードを決定 / Initialize target mode from the current selection */
    (function initTargetModeBySelection() {
        try {
            var selectedItems = doc.selection;
            if (!selectedItems || selectedItems.length === 0) {
                targetArtboardRadio.value = true;
                targetMode = "artboard";
                updateUiByTargetMode();
                return;
            }

            var allSelectedAreGroups = true;
            for (var i = 0; i < selectedItems.length; i++) {
                if (!selectedItems[i] || selectedItems[i].typename !== "GroupItem") { allSelectedAreGroups = false; break; }
            }

            if (allSelectedAreGroups) {
                targetSelectionRadio.value = true;
                targetMode = "selection";
                updateUiByTargetMode();
            } else {
                targetArtboardRadio.value = true;
                targetMode = "artboard";
                updateUiByTargetMode();
            }
        } catch (e) {
            targetArtboardRadio.value = true;
            targetMode = "artboard";
            updateUiByTargetMode();
        }
    })();

    targetArtboardRadio.onClick = function () { targetMode = "artboard"; updateUiByTargetMode(); refreshPreview(); };
    targetSelectionRadio.onClick = function () { targetMode = "selection"; updateUiByTargetMode(); refreshPreview(); };

    updateUiByTargetMode();

    /* ダイアログ下部のボタン / Dialog footer buttons */
    var dialogButtonsGroup = mainDialog.add("group");
    dialogButtonsGroup.alignment = "right";
    var cancelButton = dialogButtonsGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    var runButton = dialogButtonsGroup.add("button", undefined, L("run"), { name: "ok" });

    var dataRows = [];
    var headerNames = [];

    var previewDoc = null;              // プレビュー用に開いた複製ドキュメント
    var previewTempFile = null;         // プレビュー用の一時ファイル
    var suppressPreviewRefresh = false; // ファイル読み込み中などに自動更新を抑止

    // =========================================
    // ドキュメント複製（別名保存）
    // =========================================

    /* 複製ファイルのパスを生成 / Build the path for the duplicated file */
    function buildDuplicateFilePath(originalFile) {
        var now = new Date();
        /* 2桁ゼロ埋め / Zero-pad to 2 digits */
        function pad2(n) { return (n < 10 ? "0" : "") + String(n); }
        var timestamp = String(now.getFullYear()) + pad2(now.getMonth() + 1) + pad2(now.getDate()) + "_" + pad2(now.getHours()) + pad2(now.getMinutes()) + pad2(now.getSeconds());

        var parentFolder = originalFile.parent;
        var originalName = originalFile.name;
        var dotIndex = originalName.lastIndexOf(".");
        var baseName = (dotIndex >= 0) ? originalName.substring(0, dotIndex) : originalName;
        var extension = (dotIndex >= 0) ? originalName.substring(dotIndex) : ".ai";

        var duplicateName = baseName + "_import_" + timestamp + extension;
        return new File(parentFolder.fsName + "/" + duplicateName);
    }

    /* 別名保存でドキュメントを複製 / Duplicate the document via Save As */
    function duplicateDocumentBySaveAs(targetDoc) {
        var originalFile = targetDoc.fullName;
        var duplicateFile = buildDuplicateFilePath(originalFile);
        targetDoc.saveAs(duplicateFile); // 現在のドキュメントが複製ファイルに切り替わる
        return duplicateFile;
    }

    // =========================================
    // CSV/TSV 解析とリストプレビュー
    // =========================================

    /* 1行を区切り文字で分割（CSVは引用符も解析）/ Split a line by delimiter (CSV is quote-aware) */
    function parseDelimitedLine(line, delimiter) {
        if (delimiter === "\t") return line.split("\t");
        var fields = [], csvPattern = new RegExp("(\\" + delimiter + "|\\r?\\n|\\r|^)(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|([^\"\\" + delimiter + "\\r\\n]*))", "gi");
        var regexMatch = null;
        while (regexMatch = csvPattern.exec(line)) {
            fields.push(regexMatch[2] ? regexMatch[2].replace(/\"\"/g, "\"") : regexMatch[3]);
        }
        return fields;
    }

    /* データファイルを読み込み、リストボックスと内部データを更新 / Load a data file and refresh the list */
    function loadAndDisplayDataFile(dataFile) {
        if (previewListBox !== null) previewListContainer.remove(previewListBox);
        dataRows = [];
        dataFile.open("r");
        dataFile.encoding = "UTF-8";
        var fileText = dataFile.read();
        dataFile.close();

        var lines = fileText.split(/[\r\n]+/);
        if (lines.length === 0) return;

        var delimiter = dataFile.name.toLowerCase().match(/\.csv$/) ? "," : "\t";
        headerNames = parseDelimitedLine(lines[0], delimiter);

        // ヘッダーの正規化（BOM/前後空白を除去）
        for (var i = 0; i < headerNames.length; i++) {
            headerNames[i] = String(headerNames[i]).replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
        }

        artboardNameFieldDropdown.removeAll();
        for (var i = 0; i < headerNames.length; i++) artboardNameFieldDropdown.add("item", headerNames[i]);
        artboardNameFieldDropdown.selection = 0;

        previewListBox = previewListContainer.add("listbox", [0, 0, 550, 180], [], {
            numberOfColumns: headerNames.length, showHeaders: true, columnTitles: headerNames
        });

        for (var i = 1; i < lines.length; i++) {
            if (lines[i] === "") continue;
            var fieldValues = parseDelimitedLine(lines[i], delimiter);

            // 値の正規化（前後空白を除去）
            for (var j = 0; j < fieldValues.length; j++) {
                fieldValues[j] = String(fieldValues[j]).replace(/^\s+|\s+$/g, "");
            }

            dataRows.push(fieldValues);
            var listItem = previewListBox.add("item", fieldValues[0] || "");
            for (var j = 1; j < fieldValues.length; j++) {
                if (j < headerNames.length) listItem.subItems[j - 1].text = fieldValues[j] || "";
            }
        }
        mainDialog.layout.layout(true);
    }

    // =========================================
    // 流し込み処理（共通）
    // =========================================

    /* 指定ドキュメントへ全データを流し込む / Flow all data rows into the given document
       戻り値: 成功なら true、中断なら false（!isPreview のときはエラー表示し progressDialog を閉じる） */
    function flowDataIntoDocument(targetDoc, fieldIndex, wrapColumns, progressBar, progressDialog, isPreview) {
        // 雛形の情報を取得（レイアウトの基準はアートボード0）
        var sourceArtboard = targetDoc.artboards[0];
        var sourceArtboardRect = sourceArtboard.artboardRect;
        var itemSpacing = 50;

        var templateItems = [];  // 対象アイテム（雛形）
        var templateWidth = 0;   // 複製配置の幅（雛形のサイズ）
        var templateHeight = 0;  // 複製配置の高さ

        /* アイテム群の外接バウンディングを取得 / Get the union bounds of items */
        function getUnionBounds(pageItems) {
            // returns [left, top, right, bottom]
            var left = null, top = null, right = null, bottom = null;
            for (var i = 0; i < pageItems.length; i++) {
                var currentItem = pageItems[i];
                if (!currentItem) continue;
                var itemBounds = currentItem.visibleBounds ? currentItem.visibleBounds : currentItem.geometricBounds;
                if (!itemBounds) continue;
                var itemLeft = itemBounds[0], itemTop = itemBounds[1], itemRight = itemBounds[2], itemBottom = itemBounds[3];
                if (left === null || itemLeft < left) left = itemLeft;
                if (top === null || itemTop > top) top = itemTop;
                if (right === null || itemRight > right) right = itemRight;
                if (bottom === null || itemBottom < bottom) bottom = itemBottom;
            }
            return [left, top, right, bottom];
        }

        if (targetMode === "selection") {
            // 選択グループ：選択中のグループだけを雛形として複製
            var selectedItems = targetDoc.selection;
            if (!selectedItems || selectedItems.length === 0) {
                if (progressDialog) progressDialog.close();
                if (!isPreview) alert(L("needGroupSelect"));
                return false;
            }

            // グループ以外が混ざっていたら中断（安全側）
            for (var i = 0; i < selectedItems.length; i++) {
                if (!selectedItems[i] || selectedItems[i].typename !== "GroupItem") {
                    if (progressDialog) progressDialog.close();
                    if (!isPreview) alert(L("onlyGroupSelect"));
                    return false;
                }
            }

            templateItems = selectedItems;
            var templateBounds = getUnionBounds(templateItems);
            templateWidth = templateBounds[2] - templateBounds[0];
            templateHeight = templateBounds[1] - templateBounds[3];
        } else {
            // アートボード：現在のアートボード上の選択オブジェクトを雛形として使用
            targetDoc.selectObjectsOnActiveArtboard();
            templateItems = targetDoc.selection;
            if (!templateItems || templateItems.length === 0) {
                if (progressDialog) progressDialog.close();
                if (!isPreview) alert(L("noTemplate"));
                return false;
            }

            // アートボードサイズを基準に配置
            templateWidth = sourceArtboardRect[2] - sourceArtboardRect[0];
            templateHeight = sourceArtboardRect[1] - sourceArtboardRect[3];
        }

        // グリッド全体のサイズから、カンバス中央へ寄せるシフト量を算出
        // （アートボードモードのみ。雛形アートボード0と全セルをまとめて移動して、
        // 　右下方向だけでなく上下左右へ均等展開し、カンバス上限に当たりにくくする）
        var gridShiftX = 0;
        var gridShiftY = 0;
        if (targetMode !== "selection") {
            var totalCount = dataRows.length;
            var gridCols = Math.min(wrapColumns, totalCount);
            var gridRows = Math.ceil(totalCount / wrapColumns);
            var gridWidth = (gridCols - 1) * (templateWidth + itemSpacing) + templateWidth;
            var gridHeight = (gridRows - 1) * (templateHeight + itemSpacing) + templateHeight;
            // グリッド中心が雛形アートボード0の元の中心に重なるよう全体をずらす
            gridShiftX = (templateWidth - gridWidth) / 2;
            gridShiftY = (gridHeight - templateHeight) / 2;
        }

        // 重要：2件目(i=1)から先に処理する（タグを維持するため）
        for (var i = 1; i < dataRows.length; i++) {
            if (progressBar) { progressBar.value = i; progressDialog.update(); }

            var colIndex = i % wrapColumns;
            var rowIndex = Math.floor(i / wrapColumns);
            var dx = (templateWidth + itemSpacing) * colIndex + gridShiftX;
            var dy = -(templateHeight + itemSpacing) * rowIndex + gridShiftY;

            if (targetMode !== "selection") {
                // カンバス上限を超えるとadd()が例外を投げるため、捕捉してエラー表示
                var newArtboard;
                try {
                    newArtboard = targetDoc.artboards.add([
                        sourceArtboardRect[0] + dx, sourceArtboardRect[1] + dy,
                        sourceArtboardRect[2] + dx, sourceArtboardRect[3] + dy
                    ]);
                } catch (e) {
                    if (progressDialog) progressDialog.close();
                    if (!isPreview) alert(L("artboardLimit").replace("#detail#", String(e)));
                    return false;
                }
                newArtboard.name = (dataRows[i][fieldIndex] || "Data_" + (i + 1)).toString();
            }

            // 選択グループ指定時は、アートボードを増やさず同一アートボード内に複製して流し込み
            for (var j = 0; j < templateItems.length; j++) {
                var duplicatedItem = templateItems[j].duplicate();
                duplicatedItem.translate(dx, dy);
                replacePlaceholdersRecursive(duplicatedItem, headerNames, dataRows[i]);
            }
        }

        // 最後に1件目(i=0)を書き換える。雛形アートボード0と雛形オブジェクトもグリッド中央へシフト
        if (progressBar) { progressBar.value = dataRows.length; progressDialog.update(); }
        if (targetMode !== "selection") {
            sourceArtboard.artboardRect = [
                sourceArtboardRect[0] + gridShiftX, sourceArtboardRect[1] + gridShiftY,
                sourceArtboardRect[2] + gridShiftX, sourceArtboardRect[3] + gridShiftY
            ];
            sourceArtboard.name = (dataRows[0][fieldIndex] || "Data_1").toString();
        }
        for (var k = 0; k < templateItems.length; k++) {
            if (targetMode !== "selection") templateItems[k].translate(gridShiftX, gridShiftY);
            replacePlaceholdersRecursive(templateItems[k], headerNames, dataRows[0]);
        }

        targetDoc.selection = null;
        return true;
    }

    // =========================================
    // プレビュー
    // =========================================

    /* プレビュー用ドキュメントと一時ファイルを破棄 / Discard the preview document and its temp file */
    function closePreviewDoc() {
        var hadPreview = (previewDoc !== null);
        try {
            if (previewDoc !== null) previewDoc.close(SaveOptions.DONOTSAVECHANGES);
        } catch (e) { }
        previewDoc = null;
        try {
            if (previewTempFile !== null && previewTempFile.exists) previewTempFile.remove();
        } catch (e) { }
        previewTempFile = null;
        // プレビューを閉じたときだけ元のドキュメントへ戻す
        if (hadPreview) {
            try { app.activeDocument = doc; } catch (e) { }
        }
    }

    /* プレビューを再生成（複製ドキュメントへ全件流し込み）/ Rebuild the preview into a fresh duplicate document */
    function refreshPreview() {
        if (suppressPreviewRefresh) return;
        closePreviewDoc();
        if (!previewCheckbox.value || dataRows.length === 0) return;

        var fieldIndex = artboardNameFieldDropdown.selection ? artboardNameFieldDropdown.selection.index : 0;
        var wrapColumns = parseInt(wrapColsInput.text) || 4;

        try {
            // 最後に保存された状態の元ファイルを一時ファイルへコピーして開く
            var originalFile = doc.fullName;
            var dotIndex = originalFile.name.lastIndexOf(".");
            var extension = (dotIndex >= 0) ? originalFile.name.substring(dotIndex) : ".ai";
            previewTempFile = new File(Folder.temp.fsName + "/VDImport_preview_" + (new Date()).getTime() + extension);
            originalFile.copy(previewTempFile);
            previewDoc = app.open(previewTempFile);
        } catch (e) {
            closePreviewDoc();
            return;
        }

        // プレビューは中断時もエラーを出さない（isPreview = true）
        if (!flowDataIntoDocument(previewDoc, fieldIndex, wrapColumns, null, null, true)) {
            closePreviewDoc();
        }
    }

    // =========================================
    // UIイベント
    // =========================================

    fileDropdown.onChange = function () {
        if (!fileDropdown.selection) return;
        // ファイル読み込み中はプレビュー自動更新を抑止し、最後に一度だけ更新
        suppressPreviewRefresh = true;
        loadAndDisplayDataFile(dataFiles[fileDropdown.selection.index]);
        suppressPreviewRefresh = false;
        refreshPreview();
    };
    artboardNameFieldDropdown.onChange = function () { refreshPreview(); };
    wrapColsInput.onChange = function () { refreshPreview(); };
    previewCheckbox.onClick = function () { refreshPreview(); };
    cancelButton.onClick = function () { closePreviewDoc(); };
    mainDialog.onClose = function () { closePreviewDoc(); };

    fileDropdown.selection = 0;

    // =========================================
    // 実行処理
    // =========================================

    /* 実行ボタン押下時のメイン処理 / Main routine triggered by the Run button */
    runButton.onClick = function () {
        if (dataRows.length === 0) return;
        var fieldIndex = artboardNameFieldDropdown.selection ? artboardNameFieldDropdown.selection.index : 0;
        var wrapColumns = parseInt(wrapColsInput.text) || 4;

        // プレビュー用ドキュメントが残っていれば閉じる
        closePreviewDoc();

        // 実行前にドキュメントを複製（別名保存）し、複製したファイルに対して流し込み
        try {
            duplicateDocumentBySaveAs(doc);
        } catch (e) {
            alert(L("dupFailed").replace("#detail#", String(e)));
            return;
        }

        mainDialog.close();

        var progressDialog = new Window("palette", L("processing"), undefined);
        var progressBar = progressDialog.add("progressbar", undefined, 0, dataRows.length);
        progressBar.preferredSize.width = 300;
        progressDialog.show();

        if (flowDataIntoDocument(doc, fieldIndex, wrapColumns, progressBar, progressDialog, false)) {
            progressDialog.close();
            alert(L("done").replace("#count#", String(dataRows.length)));
        }
        // 中断時は flowDataIntoDocument 内で progressDialog を閉じ、エラー表示済み
    };

    // =========================================
    // テキスト置換
    // =========================================

    /* オブジェクトを再帰探索してプレースホルダーを置換 / Recursively replace placeholders in an item */
    function replacePlaceholdersRecursive(pageItem, headerNames, rowValues) {
        if (!pageItem) return;

        if (pageItem.typename === "TextFrame") {
            var textContent = pageItem.contents;
            if (textContent == null) return;

            for (var i = 0; i < headerNames.length; i++) {
                var headerName = String(headerNames[i]).replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
                if (headerName === "") continue;

                var placeholder = "<" + headerName + ">";
                if (textContent.indexOf(placeholder) !== -1) {
                    var value = (rowValues && rowValues.length > i && rowValues[i] != null) ? String(rowValues[i]) : "";
                    // 全出現を置換（ExtendScript安全）
                    textContent = textContent.split(placeholder).join(value);
                }
            }

            pageItem.contents = textContent;
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                replacePlaceholdersRecursive(pageItem.pageItems[i], headerNames, rowValues);
            }
        }
    }

    mainDialog.show();
})();
