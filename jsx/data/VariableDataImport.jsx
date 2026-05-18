#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
VariableDataImport.jsx

概要 / Overview:
CSV / タブ区切りテキスト のデータファイルを Illustrator のテンプレートに流し込むデータ結合スクリプト。
テキストフレーム内の <ヘッダー名> 形式のタグをデータ値に置換し、
データ行数ぶんのアートボードとバリエーションを生成する。
A data-merge script for Illustrator. It imports a CSV / TSV data file,
replaces <header> placeholder tags inside text frames with row values,
and generates one artboard variation per data row.

主な動作 / Behavior:
- データファイルは、開いているドキュメントと同じフォルダーの .csv / .txt から選ぶ。
- 実行時に元ドキュメントを別名保存で複製し、複製側に流し込む（元ファイルは無変更）。
- アートボード0を雛形とし、データ件数ぶんアートボードを複製する。
  正方形に近いグリッドを計算し、カンバスの天地・左右中央に配置する。
- アートボード名は、選択したデータ列の値から設定する。
- 「プレビュー」をオンにすると、元ファイルを変更せず複製ファイルで結果を確認できる。
  設定変更時は元ドキュメントからプレビュー用ファイルを作り直して反映する。
  「プレビュー」をオフにすると、開いているプレビュー用ドキュメントを保存せず閉じる。
- 1件目のデータを最後に処理し、雛形のタグ消失を回避する。
*/

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.4.1";

/* 実行環境のロケールから表示言語を判定 / Detect the display language from the runtime locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "データ流し込み", en: "Data Import" },
    infoTitle: { ja: "データファイル", en: "Data File" },
    fileLabel: { ja: "ファイル", en: "File" },
    settingsTitle: { ja: "流し込み設定", en: "Import Settings" },
    abNameLabel: { ja: "アートボード名の参照列", en: "Artboard name column" },
    preview: { ja: "プレビュー", en: "Preview" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    run: { ja: "実行", en: "Run" },
    processing: { ja: "処理中...", en: "Processing..." },
    done: { ja: "完了！\n#count# 件処理しました。", en: "Done!\nProcessed #count# rows." },
    noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    needSave: { ja: "ドキュメントを保存してから実行してください。", en: "Please save the document before running." },
    noDataFiles: { ja: "同じ階層に.txtまたは.csvファイルが見つかりませんでした。", en: "No .txt or .csv files found in the same folder." },
    noTemplate: { ja: "雛形にオブジェクトがありません。", en: "No template objects found." },
    dupFailed: { ja: "ドキュメントの複製（別名保存）に失敗しました。\n\n#detail#", en: "Failed to duplicate (Save As) the document.\n\n#detail#" },
    tooManyData: { ja: "データ件数（#count# 件）がカンバスに収まりません。\n現在の設定では最大 #max# 件まで配置できます。\n間隔を小さくしてください。", en: "The number of rows (#count#) does not fit on the canvas.\nUp to #max# artboards can be placed with the current settings.\nReduce the gap." },
    gapLabel: { ja: "アートボード間隔", en: "Artboard gap" },
    fileHelp: { ja: "開いているドキュメントと同じフォルダーにある CSV / TSV ファイルを選びます。", en: "Choose a CSV / TSV file in the same folder as the open document." },
    artboardNameHelp: { ja: "各アートボード名に使うデータ列を選びます。", en: "Choose the data column used for each artboard name." },
    gapHelp: { ja: "複製するアートボード同士の間隔です。入力値は10pt単位に丸められます。", en: "The gap between duplicated artboards. Values are rounded to 10 pt increments." },
    previewHelp: { ja: "元ファイルを変更せず、複製ファイルで流し込み結果を確認します。", en: "Preview the import result in a duplicate file without changing the original." },
    runHelp: { ja: "元ドキュメントを別名保存で複製し、複製側にデータを流し込みます。", en: "Duplicate the original document with Save As and import the data into the copy." },
    emptyFile: { ja: "選択したファイルにヘッダー行がありません。", en: "The selected file has no header row." },
    fileOpenFailed: { ja: "選択したファイルを開けませんでした。", en: "Could not open the selected file." }
};

/* キーから現在の言語のラベルを取得（無ければキー名をそのまま返す）/ Get the localized label for a key (falls back to the key itself) */
function L(key) {
    try {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
    } catch (e) { }
    return key;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with a colon (full-width for JA, half-width for EN) */
function labelText(key) {
    return L(key) + (lang === "ja" ? "：" : ":");
}

/* 文字列先頭のBOM（U+FEFF / 65279）と前後の空白を除去 / Strip a leading BOM (U+FEFF / 65279) and surrounding whitespace */
function trimAndStripBom(text) {
    text = String(text);
    if (text.length && text.charCodeAt(0) === 65279) text = text.substring(1);
    return text.replace(/^\s+|\s+$/g, "");
}

(function () {

    var dataRows = [];      // 読み込んだデータ行 / Loaded data rows
    var headers = [];       // ヘッダー（列名）/ Header column names
    var previewFile = null; // プレビュー用に開く複製ファイル / Duplicate file opened for preview

    // =========================================
    // 初期チェックとデータファイル収集 / Initial checks & data file discovery
    // =========================================

    if (app.documents.length === 0) {
        alert(L("noDoc"));
        return;
    }

    var doc = app.activeDocument;
    var docPath;
    try {
        docPath = doc.path;
    } catch (e) {
        alert(L("needSave"));
        return;
    }

    var documentFolder = new Folder(docPath);
    var dataFiles = documentFolder.getFiles(/\.(txt|csv)$/i);
    if (dataFiles.length === 0) {
        alert(L("noDataFiles"));
        return;
    }

    /* カンバス範囲と雛形アートボードのサイズを取得 / Canvas bounds & template artboard size */
    var MAX_ARTBOARDS = 1000;
    var canvasBounds = getCanvasBounds();
    var templateRect = doc.artboards[0].artboardRect;
    var templateWidth = templateRect[2] - templateRect[0];
    var templateHeight = templateRect[1] - templateRect[3];

    // =========================================
    // ダイアログUIの構築 / Build the dialog UI
    // =========================================

    var LABEL_WIDTH = (lang === "ja") ? 165 : 195; // 流し込み設定パネルのラベル幅 / Label width in the import-settings panel

    /* 設定パネル用：一定幅のコロン付きラベルを追加し、項目を縦に揃える / Add a fixed-width colon label so settings rows line up */
    function addFieldLabel(parentGroup, key) {
        var labelControl = parentGroup.add("statictext", undefined, labelText(key));
        labelControl.preferredSize.width = LABEL_WIDTH;
        return labelControl;
    }

    var mainDialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    mainDialog.orientation = "column";
    mainDialog.alignChildren = ["fill", "top"];
    mainDialog.spacing = 10;
    mainDialog.margins = 15;

    /* データファイルパネル：ファイル選択とデータプレビュー / Data-file panel: file selector & data preview */
    var infoPanel = mainDialog.add("panel", undefined, L("infoTitle"));
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["fill", "top"];
    infoPanel.margins = 15;
    infoPanel.spacing = 10;

    var fileSelectGroup = infoPanel.add("group");
    fileSelectGroup.add("statictext", undefined, labelText("fileLabel"));
    var fileDropdown = fileSelectGroup.add("dropdownlist", undefined, []);
    fileDropdown.size = [350, 25];
    fileDropdown.helpTip = L("fileHelp");
    for (var i = 0; i < dataFiles.length; i++) fileDropdown.add("item", decodeURI(dataFiles[i].name));

    var listContainer = infoPanel.add("group");
    listContainer.alignment = ["fill", "fill"];
    var dataListBox = null;

    var settingsPanel = mainDialog.add("panel", undefined, L("settingsTitle"));
    settingsPanel.orientation = "column";
    settingsPanel.alignChildren = ["left", "top"];
    settingsPanel.margins = 15;

    var artboardNameGroup = settingsPanel.add("group");
    addFieldLabel(artboardNameGroup, "abNameLabel");
    var artboardNameDropdown = artboardNameGroup.add("dropdownlist", undefined, []);
    artboardNameDropdown.size = [200, 25];
    artboardNameDropdown.helpTip = L("artboardNameHelp");

    /* グリッド配置設定（アートボード間隔）/ Grid layout settings (artboard gap) */
    var defaultGap = Math.round(templateWidth / 5 / 10) * 10;

    var rowGap = settingsPanel.add("group");
    addFieldLabel(rowGap, "gapLabel");
    var gapInput = rowGap.add("edittext", undefined, String(defaultGap));
    gapInput.size = [60, 25];
    gapInput.helpTip = L("gapHelp");
    rowGap.add("statictext", undefined, "pt");

    // =========================================
    // グリッド配置の計算 / Grid layout calculations
    // =========================================

    /* 横方向に並べられる最大数（カンバス幅から算出） / Max columns by canvas width */
    function calcMaxColumns(gapSize) {
        var availableWidth = canvasBounds[2] - canvasBounds[0];
        var maxCount = Math.floor((availableWidth - templateWidth) / (gapSize + templateWidth)) + 1;
        return (maxCount < 1) ? 1 : maxCount;
    }

    /* 縦方向に並べられる最大数（カンバス高さから算出） / Max rows by canvas height */
    function calcMaxRows(gapSize) {
        var availableHeight = canvasBounds[1] - canvasBounds[3];
        var maxCount = Math.floor((availableHeight - templateHeight) / (gapSize + templateHeight)) + 1;
        return (maxCount < 1) ? 1 : maxCount;
    }

    /* アートボード間隔の入力値を取得（負値は0に補正し、10の倍数へ四捨五入）/ Read the gap input (clamp negatives to 0, round to the nearest multiple of 10) */
    function getGapValue() {
        var gap = parseFloat(gapInput.text);
        if (isNaN(gap) || gap < 0) gap = 0;
        return Math.round(gap / 10) * 10;
    }

    /*
       アートボードのグリッドを算出 / Resolve the artboard grid.
       グリッド全体の幅と高さがなるべく等しく（正方形に近く）なる列数を選び、
       カンバスの天地・左右中央に配置する。
    */
    function computeArtboardLayout() {
        var gap = getGapValue();
        var maxFitColumns = calcMaxColumns(gap); // 横に収まる最大列数 / max columns that fit
        var maxFitRows = calcMaxRows(gap);       // 縦に収まる最大行数 / max rows that fit
        var dataCount = dataRows ? dataRows.length : 0;
        var layout = { cols: maxFitColumns, rows: 0, gap: gap, fits: false };
        if (dataCount < 1) return layout; // 未読込：表示用に最大列数を返す

        var bestColumns = 0, smallestDiff = -1;
        for (var columns = 1; columns <= maxFitColumns; columns++) {
            var rows = Math.ceil(dataCount / columns);
            if (rows > maxFitRows) continue;                       // 縦に収まらない
            if (columns * rows > MAX_ARTBOARDS) continue;          // アートボード数の上限
            var occupiedColumns = (columns < dataCount) ? columns : dataCount; // 実際に使う列数
            var gridWidth = occupiedColumns * templateWidth + (occupiedColumns - 1) * gap;
            var gridHeight = rows * templateHeight + (rows - 1) * gap;
            var widthHeightDiff = Math.abs(gridWidth - gridHeight);
            if (smallestDiff < 0 || widthHeightDiff < smallestDiff) {
                smallestDiff = widthHeightDiff;
                bestColumns = columns;
            }
        }
        if (bestColumns < 1) return layout; // どの列数でも収まらない（fits:false）

        layout.cols = bestColumns;
        layout.rows = Math.ceil(dataCount / bestColumns);
        layout.fits = true;
        return layout;
    }

    /* 入力確定時に間隔を10の倍数へ四捨五入し、入力欄へ反映 / On commit, round the gap to the nearest 10 and write it back to the field */
    function roundGapInput() {
        gapInput.text = String(getGapValue());
    }

    gapInput.onChange = function () {
        roundGapInput();
        if (previewCheckbox.value) runPreview(); // プレビュー中なら複製ファイルを開き直して更新 / refresh the preview while active
    };

    // ボタンエリア：3カラム（左＝プレビュー / 中央＝スペーサー / 右＝キャンセル・実行）
    var buttonArea = mainDialog.add("group");
    buttonArea.orientation = "row";
    buttonArea.alignment = ["fill", "top"];

    var buttonAreaLeft = buttonArea.add("group");
    buttonAreaLeft.alignment = ["left", "center"];
    var previewCheckbox = buttonAreaLeft.add("checkbox", undefined, L("preview"));
    previewCheckbox.helpTip = L("previewHelp");

    var buttonAreaCenter = buttonArea.add("group"); // 中央スペーサー（伸縮）/ flexible spacer
    buttonAreaCenter.alignment = ["fill", "center"];

    var buttonAreaRight = buttonArea.add("group");
    buttonAreaRight.alignment = ["right", "center"];
    var cancelButton = buttonAreaRight.add("button", undefined, L("cancel"), { name: "cancel" });
    var runButton = buttonAreaRight.add("button", undefined, L("run"), { name: "ok" });
    runButton.helpTip = L("runHelp");

    // =========================================
    // ドキュメントの複製（別名保存）/ Document duplication (Save As)
    // =========================================

    /* 複製ファイルのパスを生成（元名_タグ_日時.拡張子）/ Build the duplicate file path (origName_tag_timestamp.ext) */
    function buildDuplicateFilePath(originalFile, tag) {
        var now = new Date();
        function padTwoDigits(num) { return (num < 10 ? "0" : "") + String(num); }
        var timestamp = String(now.getFullYear()) + padTwoDigits(now.getMonth() + 1) + padTwoDigits(now.getDate()) + "_" + padTwoDigits(now.getHours()) + padTwoDigits(now.getMinutes()) + padTwoDigits(now.getSeconds());

        var parentFolder = originalFile.parent;
        var originalName = originalFile.name;
        var dotIndex = originalName.lastIndexOf(".");
        var baseName = (dotIndex >= 0) ? originalName.substring(0, dotIndex) : originalName;
        var extension = (dotIndex >= 0) ? originalName.substring(dotIndex) : ".ai";

        var duplicateName = baseName + "_" + (tag || "import") + "_" + timestamp + extension;
        return new File(parentFolder.fsName + "/" + duplicateName);
    }

    /* 指定ドキュメントを別名保存し、操作対象として複製後ドキュメントを返す / Save the document under a new name and return the duplicated document */
    function duplicateDocumentBySaveAs(sourceDocument) {
        var originalFile = sourceDocument.fullName; // File
        var duplicateFile = buildDuplicateFilePath(originalFile);
        sourceDocument.saveAs(duplicateFile); // Illustrator switches this document to the duplicated file
        return sourceDocument;
    }

    // =========================================
    // CSV / TSV の解析と読み込み / CSV / TSV parsing & loading
    // =========================================

    /* 1行を区切り文字で分割（CSVは引用符・エスケープも解釈）/ Split one line by delimiter (CSV also honors quotes and escapes) */
    function parseLine(line, delimiter) {
        if (delimiter === "\t") return line.split("\t");
        var fields = [], pattern = new RegExp("(\\" + delimiter + "|\\r?\\n|\\r|^)(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|([^\"\\" + delimiter + "\\r\\n]*))", "gi");
        var matches = null;
        while (matches = pattern.exec(line)) {
            fields.push(matches[2] ? matches[2].replace(/\"\"/g, "\"") : matches[3]);
        }
        return fields;
    }

    /* Illustrator の最大カンバス範囲を取得（一時レイヤーで原点を測定）/ Get Illustrator's max canvas bounds (measures the origin via a temp layer) */
    function getCanvasBounds() {
        var CANVAS_MAX_SIZE = 16383;
        var targetDoc = app.activeDocument;
        var wasModified = targetDoc.modified; // 計測前の変更フラグを退避 / remember the modified flag before measuring
        var tempLayer = targetDoc.layers.add();
        var tempTextFrame = tempLayer.textFrames.add();
        var left = tempTextFrame.matrix.mValueTX;
        var top = tempTextFrame.matrix.mValueTY;
        tempLayer.remove();
        targetDoc.modified = wasModified; // 一時レイヤー追加で立った変更フラグを元に戻す / restore the modified flag
        // [left, top, right, bottom]
        return [left, top, left + CANVAS_MAX_SIZE, top - CANVAS_MAX_SIZE];
    }

    /* 選択ファイルを読み込み、ヘッダー・データ・データリストを再構築 / Load the selected file and rebuild headers, rows and the data list */
    function loadDataFile(dataFile) {
        if (dataListBox !== null) { listContainer.remove(dataListBox); dataListBox = null; }
        dataRows = [];
        dataFile.encoding = "UTF-8";
        if (!dataFile.open("r")) {
            headers = [];
            artboardNameDropdown.removeAll();
            mainDialog.layout.layout(true);
            alert(L("fileOpenFailed"));
            return;
        }
        var fileContent = dataFile.read();
        dataFile.close();

        var lines = fileContent.split(/[\r\n]+/);
        if (lines.length === 0) return;

        var delimiter = dataFile.name.toLowerCase().match(/\.csv$/) ? "," : "\t";
        headers = parseLine(lines[0], delimiter);

        // ヘッダーの正規化（BOM/前後空白を除去）/ Normalize headers (strip BOM & surrounding spaces)
        for (var headerIndex = 0; headerIndex < headers.length; headerIndex++) {
            headers[headerIndex] = trimAndStripBom(headers[headerIndex]);
        }

        // ヘッダーが実質空（空ファイル等）なら中断 / Abort when there is effectively no header (e.g. empty file)
        var hasHeader = false;
        for (var headerCheckIndex = 0; headerCheckIndex < headers.length; headerCheckIndex++) {
            if (headers[headerCheckIndex] !== "") { hasHeader = true; break; }
        }
        if (!hasHeader) {
            headers = [];
            artboardNameDropdown.removeAll();
            mainDialog.layout.layout(true);
            alert(L("emptyFile"));
            return;
        }

        artboardNameDropdown.removeAll();
        for (var dropdownHeaderIndex = 0; dropdownHeaderIndex < headers.length; dropdownHeaderIndex++) {
            artboardNameDropdown.add("item", headers[dropdownHeaderIndex]);
        }
        artboardNameDropdown.selection = 0;

        dataListBox = listContainer.add("listbox", [0, 0, 550, 180], [], {
            numberOfColumns: headers.length, showHeaders: true, columnTitles: headers
        });

        for (var lineIndex = 1; lineIndex < lines.length; lineIndex++) {
            if (lines[lineIndex] === "") continue;
            var cells = parseLine(lines[lineIndex], delimiter);

            // 値の正規化（前後空白を除去）/ Normalize cell values (strip surrounding spaces)
            for (var cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                cells[cellIndex] = String(cells[cellIndex]).replace(/^\s+|\s+$/g, "");
            }

            dataRows.push(cells);
            var listItem = dataListBox.add("item", cells[0] || "");
            for (var listCellIndex = 1; listCellIndex < cells.length; listCellIndex++) {
                if (listCellIndex < headers.length) {
                    listItem.subItems[listCellIndex - 1].text = cells[listCellIndex] || "";
                }
            }
        }

        mainDialog.layout.layout(true);
    }

    fileDropdown.onChange = function () {
        if (fileDropdown.selection) loadDataFile(dataFiles[fileDropdown.selection.index]);
    };
    fileDropdown.selection = 0;

    // =========================================
    // 実行処理（流し込み）/ Run handler (data import)
    // =========================================

    /* 配置レイアウトを確定（収まらなければ警告して null を返す）/ Resolve the layout (alerts and returns null when it does not fit) */
    function prepareLayout() {
        var nameColumnIndex = artboardNameDropdown.selection ? artboardNameDropdown.selection.index : 0;
        var gridLayout = computeArtboardLayout();
        if (!gridLayout.fits) {
            var maxFitColumns = calcMaxColumns(gridLayout.gap);
            var maxCapacity = Math.min(maxFitColumns * calcMaxRows(gridLayout.gap), MAX_ARTBOARDS);
            alert(L("tooManyData").replace("#count#", String(dataRows.length)).replace("#max#", String(maxCapacity)));
            return null;
        }
        return {
            cols: gridLayout.cols,
            rows: gridLayout.rows,
            gap: gridLayout.gap,
            nameColumnIndex: nameColumnIndex
        };
    }

    /* 雛形を複製してデータを流し込む（targetDoc に対して実行）/ Duplicate the template and merge data into targetDoc */
    function performImport(targetDoc, layout, isPreview) {
        targetDoc.activate();
        var nameColumnIndex = layout.nameColumnIndex;
        var columnCount = layout.cols;
        var gap = layout.gap;

        var progressWindow = new Window("palette", L("processing"), undefined);
        var progressBar = progressWindow.add("progressbar", undefined, 0, dataRows.length);
        progressBar.preferredSize.width = 300;
        progressWindow.show();

        // 雛形：アートボード0と、その上に乗っているオブジェクト
        var templateArtboard = targetDoc.artboards[0];
        var placementRect = templateArtboard.artboardRect;
        targetDoc.artboards.setActiveArtboardIndex(0);
        targetDoc.selectObjectsOnActiveArtboard();
        var templateItems = targetDoc.selection;
        if (!templateItems || templateItems.length === 0) {
            progressWindow.close();
            alert(L("noTemplate"));
            return;
        }

        // 雛形1セルのサイズ（アートボード0のサイズ）
        var cellWidth = placementRect[2] - placementRect[0];
        var cellHeight = placementRect[1] - placementRect[3];

        // グリッド全体の幅・高さを求め、カンバスの天地・左右中央へ配置
        var dataCount = dataRows.length;
        var gridRows = Math.ceil(dataCount / columnCount);
        var occupiedColumns = (columnCount < dataCount) ? columnCount : dataCount;
        var gridWidth = occupiedColumns * cellWidth + (occupiedColumns - 1) * gap;
        var gridHeight = gridRows * cellHeight + (gridRows - 1) * gap;
        var canvasWidth = canvasBounds[2] - canvasBounds[0];
        var canvasHeight = canvasBounds[1] - canvasBounds[3];
        var leftMargin = Math.round((canvasWidth - gridWidth) / 2);
        var topMargin = Math.round((canvasHeight - gridHeight) / 2);
        var gridLeft = canvasBounds[0] + leftMargin;
        var gridTop = canvasBounds[1] - topMargin;
        var deltaX = gridLeft - placementRect[0];
        var deltaY = gridTop - placementRect[1];
        templateArtboard.artboardRect = [gridLeft, gridTop, gridLeft + cellWidth, gridTop - cellHeight];
        for (var i = 0; i < templateItems.length; i++) {
            templateItems[i].translate(deltaX, deltaY);
        }
        placementRect = templateArtboard.artboardRect; // 移動後の位置をグリッド配置の基準にする

        // 重要：2件目(i=1)から先に処理する（タグを維持するため）
        for (var i = 1; i < dataRows.length; i++) {
            progressBar.value = i;
            progressWindow.update();

            var col = i % columnCount;
            var row = Math.floor(i / columnCount);
            var offsetX = (cellWidth + gap) * col;
            var offsetY = -(cellHeight + gap) * row;

            var newArtboard = targetDoc.artboards.add([
                placementRect[0] + offsetX, placementRect[1] + offsetY,
                placementRect[2] + offsetX, placementRect[3] + offsetY
            ]);
            newArtboard.name = (dataRows[i][nameColumnIndex] || "Data_" + (i + 1)).toString();

            // 雛形オブジェクトを複製し、対応するアートボードへ移動して流し込み
            for (var j = 0; j < templateItems.length; j++) {
                var duplicatedItem = templateItems[j].duplicate();
                duplicatedItem.translate(offsetX, offsetY);
                replaceTextRecursive(duplicatedItem, headers, dataRows[i]);
            }
        }

        // 最後に1件目(i=0)を書き換える
        progressBar.value = dataRows.length;
        progressWindow.update();
        templateArtboard.name = (dataRows[0][nameColumnIndex] || "Data_1").toString();
        for (var k = 0; k < templateItems.length; k++) {
            replaceTextRecursive(templateItems[k], headers, dataRows[0]);
        }

        targetDoc.selection = null;
        progressWindow.close();
        app.redraw();
        if (!isPreview) {
            alert(L("done").replace("#count#", String(dataRows.length)));
        }
    }

    /* 「実行」：ドキュメントを別名保存で複製し、流し込んでダイアログを閉じる / Run: duplicate via Save As, merge, then close the dialog */
    runButton.onClick = function () {
        if (dataRows.length === 0) return;
        var layout = prepareLayout();
        if (!layout) return;

        // プレビュー用ドキュメント・ファイルが残っていれば破棄してから実行 / Discard any leftover preview document & file before running
        cleanupPreviewFile();

        var importDocument;
        try {
            importDocument = duplicateDocumentBySaveAs(doc);
        } catch (e) {
            alert(L("dupFailed").replace("#detail#", String(e)));
            return;
        }
        mainDialog.close();
        performImport(importDocument, layout, false);
    };

    /* 既に開いているプレビュー用ドキュメントを探す（無ければ null）/ Find the already-open preview document (null if none) */
    function getOpenPreviewDocument() {
        if (!previewFile) return null;
        for (var i = 0; i < app.documents.length; i++) {
            try {
                if (app.documents[i].fullName.fsName === previewFile.fsName) return app.documents[i];
            } catch (e) { }
        }
        return null;
    }

    /* プレビュー用ドキュメントを閉じ、プレビュー用ファイルを削除 / Close the preview document and remove the preview file */
    function cleanupPreviewFile() {
        var openedPreview = getOpenPreviewDocument();
        if (openedPreview) {
            try { openedPreview.close(SaveOptions.DONOTSAVECHANGES); } catch (e) { }
        }
        if (previewFile && previewFile.exists) {
            try { previewFile.remove(); } catch (e) { }
        }
        previewFile = null;
    }

    /*
       プレビューを生成・更新 / Build or refresh the preview.
       既存のプレビュードキュメントがあれば一旦閉じ、元ドキュメントから複製ファイルを作り直して開く。
       流し込み結果は保存しないため、毎回クリーンな雛形状態から始まる。
    */
    function runPreview() {
        if (dataRows.length === 0) return false;
        var layout = prepareLayout();
        if (!layout) return false;

        // 既存のプレビュードキュメントがあれば閉じる（流し込み結果は破棄）
        var openedPreview = getOpenPreviewDocument();
        if (openedPreview) {
            try { openedPreview.close(SaveOptions.DONOTSAVECHANGES); } catch (e) { }
        }

        // 元ドキュメントからプレビュー用の複製ファイルを毎回作り直す
        var previewDocument;
        try {
            var sourceFile = doc.fullName;
            if (!previewFile) {
                previewFile = buildDuplicateFilePath(sourceFile, "preview");
            }
            if (previewFile.exists) {
                previewFile.remove();
            }
            if (!sourceFile.copy(previewFile)) throw new Error("File copy failed");
            previewDocument = app.open(previewFile);
        } catch (e) {
            alert(L("dupFailed").replace("#detail#", String(e)));
            previewFile = null;
            return false;
        }

        performImport(previewDocument, layout, true);
        return true;
    }

    /* 「プレビュー」：オンで表示、オフでプレビュー用ドキュメントを閉じる / Preview checkbox: show on, close the preview document off */
    previewCheckbox.onClick = function () {
        if (!previewCheckbox.value) {
            cleanupPreviewFile();
            return;
        }
        if (!runPreview()) previewCheckbox.value = false;
    };

    /* 「キャンセル」：プレビュー用ファイルを削除して閉じる / Cancel: remove the preview file and close the dialog */
    cancelButton.onClick = function () {
        cleanupPreviewFile();
        mainDialog.close();
    };

    /* アートボード名の参照項目を変更したら、プレビュー中は複製ファイルを開き直して更新 / Refresh the preview when the artboard-name field changes */
    artboardNameDropdown.onChange = function () {
        if (previewCheckbox.value) runPreview();
    };

    // =========================================
    // テキストの置換 / Text replacement
    // =========================================

    /* オブジェクトを再帰的に走査し、テキスト内の <タグ> をデータ値へ置換 / Walk objects recursively and replace <tag> placeholders in text with data values */
    function replaceTextRecursive(pageItem, headerKeys, rowValues) {
        if (!pageItem) return;

        if (pageItem.typename === "TextFrame") {
            var textContent = pageItem.contents;
            if (textContent == null) return;

            for (var k = 0; k < headerKeys.length; k++) {
                var headerKey = trimAndStripBom(headerKeys[k]);
                if (headerKey === "") continue;

                var placeholderTag = "<" + headerKey + ">";
                if (textContent.indexOf(placeholderTag) !== -1) {
                    var value = (rowValues && rowValues.length > k && rowValues[k] != null) ? String(rowValues[k]) : "";
                    // 全出現を置換（ExtendScript安全）
                    textContent = textContent.split(placeholderTag).join(value);
                }
            }

            pageItem.contents = textContent;
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                replaceTextRecursive(pageItem.pageItems[i], headerKeys, rowValues);
            }
        }
    }

    mainDialog.show();
})();
