#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

CSV / タブ区切りテキストのデータを、Illustratorのテンプレートに流し込むデータ結合スクリプトです。
テキストフレーム内の `<ヘッダー名>` タグをデータ値に置換し、データ件数ぶんのアートボードを生成します。

- 元ドキュメントを別名保存で複製し、複製側に流し込みます（元ファイルは無変更）
- アートボード0を雛形に、正方形に近いグリッドでカンバス中央へ配置します
- ［プレビュー］で、元ファイルを変更せず結果を確認できます

### 注意

- 雛形として複製されるのは、アートボード0上のロックも非表示もされていないオブジェクトだけです。

詳しい機能・使い方はREADMEを参照してください。

*/

/*

### Overview

A data-merge script for Illustrator. It replaces `<header>` tags inside text frames with CSV / TSV
values and generates one artboard per data row.

- The original document is duplicated with Save As and the data is merged into the copy
- Artboard 0 is the template; the grid is sized close to a square and centred on the canvas
- Preview shows the result without touching the original file

### Notes

- Only unlocked, visible objects on artboard 0 are duplicated as the template.

See the README for the full feature list and usage.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "VariableDataImport";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VariableDataImport.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/VariableDataImport.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n741c9f28d0fd"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

var MAX_ARTBOARD_COUNT = 1000;              /* 生成できるアートボードの上限 / artboard count limit */
var CANVAS_MAX_SIZE = 16383;                /* Illustratorのカンバス最大寸法（pt） / max canvas size in pt */
var ARTBOARD_GAP_STEP = 10;                 /* アートボード間隔の丸め単位（pt） / gap rounding step in pt */
var ARTBOARD_GAP_DIVISOR = 5;               /* 間隔の初期値＝雛形幅/この値 / gap default divisor */
var DATA_FILE_PATTERN = /\.(txt|csv)$/i;    /* データファイルとして拾う拡張子 / data file extensions */
var ARTBOARD_NAME_PREFIX = "Data_";         /* 名前が空のときのアートボード名 / fallback artboard name */
var MAX_TAG_REPLACEMENTS = 1000;            /* 1フレーム内で同一タグを置換する上限 / replacement guard */

// =========================================
// レイアウト / Layout
// =========================================

var DIALOG_MARGINS = 15;                        /* ダイアログの余白 / dialog margins */
var DIALOG_SPACING = 10;                        /* ダイアログの行間 / dialog spacing */
var PANEL_MARGINS = 15;                         /* パネルの余白 / panel margins */
var PANEL_SPACING = 10;                         /* パネルの行間 / panel spacing */
var FIELD_LABEL_WIDTH = { ja: 165, en: 195 };   /* 流し込み設定パネルのラベル幅 / settings label width */
var FILE_DROPDOWN_SIZE = [350, 25];             /* ファイル選択ドロップダウン / file dropdown */
var COLUMN_DROPDOWN_SIZE = [200, 25];           /* 列選択ドロップダウン / column dropdown */
var ARTBOARD_GAP_INPUT_SIZE = [60, 25];         /* 間隔入力欄 / gap input field */
var DATA_LIST_BOUNDS = [0, 0, 550, 180];        /* データ一覧リスト / data list box */
var PROGRESS_BAR_WIDTH = 300;                   /* 進捗バーの幅 / progress bar width */

// =========================================
// ローカライズ / Localization
// =========================================

/**
 * 実行環境のロケールから表示言語を判定する
 * @returns {string} "ja" または "en"
 */
function detectUILang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var uiLang = detectUILang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    /* ダイアログ / Dialog */
    dialog: {
        title: { ja: "データ結合", en: "Data Merge" }
    },
    /* パネル見出し / Panel titles */
    panel: {
        dataFile: { ja: "データファイル", en: "Data File" },
        settings: { ja: "アートボード設定", en: "Artboard Settings" }
    },
    /* フィールド見出し（コロンは labelText で付与）/ Field labels (colon added by labelText) */
    fieldLabel: {
        file: { ja: "ファイル", en: "File" },
        artboardNameColumn: { ja: "アートボード名に使う列", en: "Artboard name column" },
        gap: { ja: "アートボード間隔", en: "Artboard gap" }
    },
    /* チェックボックス / Checkboxes */
    checkbox: {
        preview: { ja: "プレビュー", en: "Preview" }
    },
    /* ボタン / Buttons */
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        run: { ja: "複製して実行", en: "Duplicate and Run" }
    },
    /* 進捗表示 / Progress window */
    progress: {
        title: { ja: "処理中…", en: "Processing…" }
    },
    /* ヘルプチップ / Tooltips */
    tooltip: {
        file: {
            ja: "開いているドキュメントと同じフォルダーにある CSV / TSV ファイルを選びます。",
            en: "Choose a CSV / TSV file in the same folder as the open document."
        },
        artboardNameColumn: {
            ja: "各アートボード名に使うデータ列を選びます。",
            en: "Choose the data column used for each artboard name."
        },
        gap: {
            ja: "複製するアートボード同士の間隔です。縦横とも同じ間隔で、単位はptです。\n入力値は10pt単位に丸められます。",
            en: "The gap between duplicated artboards, applied both horizontally and vertically, in points.\nValues are rounded to 10 pt increments."
        },
        preview: {
            ja: "元ファイルを変更せず、複製ファイルで流し込み結果を確認します。\n保存していない編集は反映されません。",
            en: "Preview the import result in a duplicate file without changing the original.\nUnsaved edits are not reflected."
        },
        cancel: {
            ja: "プレビュー用に作成したファイルを削除して閉じます。",
            en: "Remove the preview file and close the dialog."
        },
        run: {
            ja: "元ドキュメントを別名保存で複製し、複製側にデータを流し込みます。",
            en: "Duplicate the original document with Save As and import the data into the copy."
        }
    },
    /* 警告・完了メッセージ / Alerts */
    alert: {
        noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        needSave: { ja: "ドキュメントを保存してから実行してください。", en: "Please save the document before running." },
        noDataFiles: { ja: "同じフォルダーに.txtまたは.csvファイルが見つかりませんでした。", en: "No .txt or .csv files found in the same folder." },
        noTemplate: { ja: "アートボード1にオブジェクトがありません。", en: "No objects found on artboard 1." },
        emptyFile: { ja: "選択したファイルにヘッダー行がありません。", en: "The selected file has no header row." },
        fileOpenFailed: { ja: "選択したファイルを開けませんでした。", en: "Could not open the selected file." },
        done: {
            ja: "完了しました。\n#count# 件を処理し、次のファイルに保存しました。\n\n#filename#",
            en: "Done.\nProcessed #count# rows and saved to:\n\n#filename#"
        },
        dupFailed: {
            ja: "ドキュメントの複製に失敗しました。\n\n#detail#",
            en: "Failed to duplicate the document.\n\n#detail#"
        },
        tooManyData: {
            ja: "データ件数（#count# 件）がカンバスに収まりません。\n現在の設定では最大 #max# 件まで配置できます。\nアートボード間隔を小さくするか、データ件数を減らしてください。",
            en: "The number of rows (#count#) does not fit on the canvas.\nUp to #max# artboards can be placed with the current settings.\nReduce the artboard gap or the number of rows."
        }
    }
};

/**
 * ドット区切りキーから現在の言語のラベルを取得する（見つからなければキー名を返す）
 * @param {string} labelKey - "alert.noDoc" のような階層キー
 * @returns {string} 表示用の文字列
 */
function getLabel(labelKey) {
    var keyParts = labelKey.split(".");
    var labelNode = LABELS;
    for (var i = 0; i < keyParts.length; i++) {
        if (labelNode && labelNode[keyParts[i]] !== undefined) {
            labelNode = labelNode[keyParts[i]];
        } else {
            return labelKey;
        }
    }
    return (labelNode && labelNode[uiLang] !== undefined) ? String(labelNode[uiLang]) : labelKey;
}

/**
 * コロン付きラベルを組み立てる（日本語は全角、英語は半角）
 * @param {string} labelKey - ラベルの階層キー
 * @returns {string} コロンを付けたラベル
 */
function labelText(labelKey) {
    return getLabel(labelKey) + (uiLang === "ja" ? "：" : ":");
}

/**
 * 文字列先頭のBOM（U+FEFF / 65279）と前後の空白を除去する
 * @param {string} text - 対象の文字列
 * @returns {string} 整形後の文字列
 */
function trimAndStripBom(text) {
    text = String(text);
    if (text.length && text.charCodeAt(0) === 65279) text = text.substring(1);
    return text.replace(/^\s+|\s+$/g, "");
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {

    var dataRows = [];               // 読み込んだデータ行 / Loaded data rows
    var headerNames = [];            // ヘッダー（列名）/ Header column names
    var previewFile = null;          // プレビュー用に開く複製ファイル / Duplicate file opened for preview
    var artboardNameColumnIndex = 0; // アートボード名に使う列の番号 / Column used for artboard names

    // =========================================
    // 初期チェックとデータファイル収集 / Initial checks & data file discovery
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDoc"));
        return;
    }

    var originalDocument = app.activeDocument;
    var documentFolderPath;
    try {
        documentFolderPath = originalDocument.path;
    } catch (e) {
        alert(getLabel("alert.needSave"));
        return;
    }

    var documentFolder = new Folder(documentFolderPath);
    var dataFiles = documentFolder.getFiles(DATA_FILE_PATTERN);
    if (dataFiles.length === 0) {
        alert(getLabel("alert.noDataFiles"));
        return;
    }

    /* カンバス範囲と雛形アートボードのサイズを取得 / Canvas bounds & template artboard size */
    var canvasBounds = getCanvasBounds();
    var templateArtboardRect = originalDocument.artboards[0].artboardRect;
    var templateArtboardWidth = templateArtboardRect[2] - templateArtboardRect[0];
    var templateArtboardHeight = templateArtboardRect[1] - templateArtboardRect[3];

    // =========================================
    // ダイアログUIの構築 / Build the dialog UI
    // =========================================

    /**
     * 設定パネル用に、一定幅のコロン付きラベルを追加して項目を縦に揃える
     * @param {Group} parentGroup - ラベルを追加するグループ
     * @param {string} labelKey - ラベルの階層キー
     * @returns {StaticText} 追加したラベル
     */
    function addFieldLabel(parentGroup, labelKey) {
        var fieldLabel = parentGroup.add("statictext", undefined, labelText(labelKey));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH[uiLang] || FIELD_LABEL_WIDTH.en;
        return fieldLabel;
    }

    var mainDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    mainDialog.orientation = "column";
    mainDialog.alignChildren = ["fill", "top"];
    mainDialog.spacing = DIALOG_SPACING;
    mainDialog.margins = DIALOG_MARGINS;

    /* データファイルパネル：ファイル選択とデータ一覧 / Data-file panel: file selector & data list */
    var dataFilePanel = mainDialog.add("panel", undefined, getLabel("panel.dataFile"));
    dataFilePanel.orientation = "column";
    dataFilePanel.alignChildren = ["fill", "top"];
    dataFilePanel.margins = PANEL_MARGINS;
    dataFilePanel.spacing = PANEL_SPACING;

    var fileSelectGroup = dataFilePanel.add("group");
    fileSelectGroup.add("statictext", undefined, labelText("fieldLabel.file"));
    var dataFileDropdown = fileSelectGroup.add("dropdownlist", undefined, []);
    dataFileDropdown.size = FILE_DROPDOWN_SIZE;
    dataFileDropdown.helpTip = getLabel("tooltip.file");
    for (var i = 0; i < dataFiles.length; i++) dataFileDropdown.add("item", decodeURI(dataFiles[i].name));

    var dataListGroup = dataFilePanel.add("group");
    dataListGroup.alignment = ["fill", "fill"];
    var dataListBox = null;

    var importSettingsPanel = mainDialog.add("panel", undefined, getLabel("panel.settings"));
    importSettingsPanel.orientation = "column";
    importSettingsPanel.alignChildren = ["left", "top"];
    importSettingsPanel.margins = PANEL_MARGINS;

    var artboardNameColumnGroup = importSettingsPanel.add("group");
    addFieldLabel(artboardNameColumnGroup, "fieldLabel.artboardNameColumn");
    var artboardNameColumnDropdown = artboardNameColumnGroup.add("dropdownlist", undefined, []);
    artboardNameColumnDropdown.size = COLUMN_DROPDOWN_SIZE;
    artboardNameColumnDropdown.helpTip = getLabel("tooltip.artboardNameColumn");

    /* グリッド配置設定（アートボード間隔）/ Grid layout settings (artboard gap) */
    var defaultArtboardGap = Math.round(templateArtboardWidth / ARTBOARD_GAP_DIVISOR / ARTBOARD_GAP_STEP) * ARTBOARD_GAP_STEP;

    var artboardGapGroup = importSettingsPanel.add("group");
    addFieldLabel(artboardGapGroup, "fieldLabel.gap");
    var artboardGapInput = artboardGapGroup.add("edittext", undefined, String(defaultArtboardGap));
    artboardGapInput.size = ARTBOARD_GAP_INPUT_SIZE;
    artboardGapInput.helpTip = getLabel("tooltip.gap");
    artboardGapGroup.add("statictext", undefined, "pt");

    // =========================================
    // グリッド配置の計算 / Grid layout calculations
    // =========================================

    /**
     * 横方向に並べられるアートボードの最大数をカンバス幅から算出する
     * @param {number} artboardGap - アートボード間隔（pt）
     * @returns {number} 収まる列数（収まらなければ0）
     */
    function calcMaxArtboardColumns(artboardGap) {
        var availableWidth = canvasBounds[2] - canvasBounds[0];
        var maxColumnCount = Math.floor((availableWidth - templateArtboardWidth) / (artboardGap + templateArtboardWidth)) + 1;
        return (maxColumnCount < 0) ? 0 : maxColumnCount;
    }

    /**
     * 縦方向に並べられるアートボードの最大数をカンバス高さから算出する
     * @param {number} artboardGap - アートボード間隔（pt）
     * @returns {number} 収まる行数（収まらなければ0）
     */
    function calcMaxArtboardRows(artboardGap) {
        var availableHeight = canvasBounds[1] - canvasBounds[3];
        var maxRowCount = Math.floor((availableHeight - templateArtboardHeight) / (artboardGap + templateArtboardHeight)) + 1;
        return (maxRowCount < 0) ? 0 : maxRowCount;
    }

    /**
     * アートボード間隔の入力値を取得する（負値は0に補正し、10の倍数へ四捨五入）
     * @returns {number} 丸めた間隔（pt）
     */
    function getArtboardGap() {
        var gapValue = parseFloat(artboardGapInput.text);
        if (isNaN(gapValue) || gapValue < 0) gapValue = 0;
        return Math.round(gapValue / ARTBOARD_GAP_STEP) * ARTBOARD_GAP_STEP;
    }

    /**
     * グリッド全体の外形サイズを求める
     * @param {number} columnCount - 列数
     * @param {number} rowCount - 行数
     * @param {number} cellWidth - 1セルの幅（pt）
     * @param {number} cellHeight - 1セルの高さ（pt）
     * @param {number} artboardGap - アートボード間隔（pt）
     * @returns {{width: number, height: number}} グリッド全体の幅と高さ
     */
    function calcGridSize(columnCount, rowCount, cellWidth, cellHeight, artboardGap) {
        return {
            width: columnCount * cellWidth + (columnCount - 1) * artboardGap,
            height: rowCount * cellHeight + (rowCount - 1) * artboardGap
        };
    }

    /**
     * アートボードのグリッドを算出する
     * グリッド全体の幅と高さがなるべく等しく（正方形に近く）なる列数を選ぶ
     * @returns {{columnCount: number, rowCount: number, artboardGap: number, fits: boolean}} 配置情報
     */
    function computeArtboardLayout() {
        var artboardGap = getArtboardGap();
        var maxFitColumns = calcMaxArtboardColumns(artboardGap); // 横に収まる最大列数 / max columns that fit
        var maxFitRows = calcMaxArtboardRows(artboardGap);       // 縦に収まる最大行数 / max rows that fit
        var dataCount = dataRows ? dataRows.length : 0;
        var artboardLayout = { columnCount: maxFitColumns, rowCount: 0, artboardGap: artboardGap, fits: false };
        if (dataCount < 1) return artboardLayout;                     // 未読込：表示用に最大列数を返す
        if (dataCount > MAX_ARTBOARD_COUNT) return artboardLayout;    // アートボード数の上限を超過

        var bestColumnCount = 0, smallestDiff = -1;
        for (var columnCount = 1; columnCount <= maxFitColumns; columnCount++) {
            var rowCount = Math.ceil(dataCount / columnCount);
            if (rowCount > maxFitRows) continue;                                       // 縦に収まらない
            var occupiedColumns = (columnCount < dataCount) ? columnCount : dataCount; // 実際に使う列数
            var gridSize = calcGridSize(occupiedColumns, rowCount, templateArtboardWidth, templateArtboardHeight, artboardGap);
            var widthHeightDiff = Math.abs(gridSize.width - gridSize.height);
            if (smallestDiff < 0 || widthHeightDiff < smallestDiff) {
                smallestDiff = widthHeightDiff;
                bestColumnCount = columnCount;
            }
        }
        if (bestColumnCount < 1) return artboardLayout; // どの列数でも収まらない（fits:false）

        artboardLayout.columnCount = bestColumnCount;
        artboardLayout.rowCount = Math.ceil(dataCount / bestColumnCount);
        artboardLayout.fits = true;
        return artboardLayout;
    }

    /**
     * 入力確定時に間隔を10の倍数へ四捨五入し、入力欄へ反映する
     * @returns {void}
     */
    function roundArtboardGapInput() {
        artboardGapInput.text = String(getArtboardGap());
    }

    artboardGapInput.onChange = function () {
        roundArtboardGapInput();
        refreshPreviewIfActive();
    };

    /* ボタンバー：3カラム（左＝プレビュー / 中央＝スペーサー / 右＝キャンセル・実行） */
    var buttonBarGroup = mainDialog.add("group");
    buttonBarGroup.orientation = "row";
    buttonBarGroup.alignment = ["fill", "top"];

    var previewToggleGroup = buttonBarGroup.add("group");
    previewToggleGroup.alignment = ["left", "center"];
    var previewCheckbox = previewToggleGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
    previewCheckbox.helpTip = getLabel("tooltip.preview");

    var buttonSpacerGroup = buttonBarGroup.add("group"); // 中央スペーサー（伸縮）/ flexible spacer
    buttonSpacerGroup.alignment = ["fill", "center"];

    var actionButtonGroup = buttonBarGroup.add("group");
    actionButtonGroup.alignment = ["right", "center"];
    var cancelButton = actionButtonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    cancelButton.helpTip = getLabel("tooltip.cancel");
    var runButton = actionButtonGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });
    runButton.helpTip = getLabel("tooltip.run");

    // =========================================
    // ドキュメントの複製（別名保存）/ Document duplication (Save As)
    // =========================================

    /**
     * 複製ファイルのパスを生成する（元名_用途_日時.拡張子）
     * @param {File} originalFile - 元ファイル
     * @param {string} [purposeTag] - ファイル名に挟む用途識別子（既定は "import"）
     * @returns {File} 複製先のファイル
     */
    function buildDuplicateFilePath(originalFile, purposeTag) {
        var now = new Date();
        function padTwoDigits(numberValue) { return (numberValue < 10 ? "0" : "") + String(numberValue); }
        var timestamp = String(now.getFullYear()) + padTwoDigits(now.getMonth() + 1) + padTwoDigits(now.getDate()) + "_" + padTwoDigits(now.getHours()) + padTwoDigits(now.getMinutes()) + padTwoDigits(now.getSeconds());

        var parentFolder = originalFile.parent;
        var originalFileName = originalFile.name;
        var dotIndex = originalFileName.lastIndexOf(".");
        var baseName = (dotIndex >= 0) ? originalFileName.substring(0, dotIndex) : originalFileName;
        var fileExtension = (dotIndex >= 0) ? originalFileName.substring(dotIndex) : ".ai";

        var duplicateFileName = baseName + "_" + (purposeTag || "import") + "_" + timestamp + fileExtension;
        return new File(parentFolder.fsName + "/" + duplicateFileName);
    }

    /**
     * 指定ドキュメントを別名保存し、操作対象として複製後ドキュメントを返す
     * @param {Document} sourceDocument - 複製元のドキュメント
     * @returns {Document} 別名保存後のドキュメント
     */
    function duplicateDocumentBySaveAs(sourceDocument) {
        var originalFile = sourceDocument.fullName;
        var duplicateFile = buildDuplicateFilePath(originalFile);
        sourceDocument.saveAs(duplicateFile); // Illustrator switches this document to the duplicated file
        return sourceDocument;
    }

    // =========================================
    // CSV / TSV の解析と読み込み / CSV / TSV parsing & loading
    // =========================================

    /**
     * データ1行を区切り文字で分割する（CSVは引用符・""エスケープも解釈）
     * @param {string} dataLine - 1行分の文字列
     * @param {string} delimiter - 区切り文字（"," または "\t"）
     * @returns {string[]} 分割したフィールド
     */
    function parseDataLine(dataLine, delimiter) {
        if (delimiter === "\t") return dataLine.split("\t");
        var parsedFields = [], currentField = "", inQuotes = false;
        for (var i = 0; i < dataLine.length; i++) {
            var currentChar = dataLine.charAt(i);
            if (inQuotes) {
                if (currentChar !== '"') currentField += currentChar;
                else if (dataLine.charAt(i + 1) === '"') { currentField += '"'; i++; } // "" は引用符1つ / escaped quote
                else inQuotes = false;
            } else if (currentChar === '"') {
                inQuotes = true;
            } else if (currentChar === delimiter) {
                parsedFields.push(currentField);
                currentField = "";
            } else {
                currentField += currentChar;
            }
        }
        parsedFields.push(currentField);
        return parsedFields;
    }

    /**
     * バイト列がUTF-8として妥当か判定する
     * @param {string} byteString - 1バイト＝1文字で読み込んだ文字列（BINARY読み）
     * @returns {boolean} UTF-8として解釈できればtrue
     */
    function isValidUtf8(byteString) {
        var i = 0, byteCount = byteString.length;
        while (i < byteCount) {
            var leadByte = byteString.charCodeAt(i);
            if (leadByte <= 0x7F) { i++; continue; }
            var followByteCount;
            if (leadByte >= 0xC2 && leadByte <= 0xDF) followByteCount = 1;
            else if (leadByte >= 0xE0 && leadByte <= 0xEF) followByteCount = 2;
            else if (leadByte >= 0xF0 && leadByte <= 0xF4) followByteCount = 3;
            else return false;
            if (i + followByteCount >= byteCount) return false;
            for (var j = 1; j <= followByteCount; j++) {
                var followByte = byteString.charCodeAt(i + j);
                if (followByte < 0x80 || followByte > 0xBF) return false;
            }
            i += followByteCount + 1;
        }
        return true;
    }

    /**
     * データファイルを読み込む（UTF-8として解釈できなければShift-JISで読み直す）
     * @param {File} dataFile - 読み込むファイル
     * @returns {string} ファイル全体の文字列（読み込めなければnull）
     */
    function readDataFileText(dataFile) {
        /* まずバイト列として読み、エンコーディングを判定 / Read raw bytes to detect the encoding */
        var rawByteString = readFileAs(dataFile, "BINARY");
        if (rawByteString === null) return null;
        return readFileAs(dataFile, isValidUtf8(rawByteString) ? "UTF-8" : "SJIS");
    }

    /**
     * 指定のエンコーディングでファイル全体を読む
     * @param {File} dataFile - 読み込むファイル
     * @param {string} encoding - ExtendScriptのエンコーディング名（"BINARY" / "UTF-8" / "SJIS"）
     * @returns {string} ファイル全体の文字列（開けなければnull）
     */
    function readFileAs(dataFile, encoding) {
        dataFile.encoding = encoding;
        if (!dataFile.open("r")) return null;
        var fileContent = dataFile.read();
        dataFile.close();
        return fileContent;
    }

    /**
     * Illustratorの最大カンバス範囲を取得する（一時レイヤーで原点を測定）
     * @returns {number[]} [left, top, right, bottom]
     */
    function getCanvasBounds() {
        var measuredDocument = app.activeDocument;
        var wasModified = measuredDocument.modified; // 計測前の変更フラグを退避 / remember the flag before measuring
        var originLayer = measuredDocument.layers.add();
        var originTextFrame = originLayer.textFrames.add();
        var originLeft = originTextFrame.matrix.mValueTX;
        var originTop = originTextFrame.matrix.mValueTY;
        originLayer.remove();
        measuredDocument.modified = wasModified; // 一時レイヤーで立った変更フラグを元に戻す / restore the flag
        return [originLeft, originTop, originLeft + CANVAS_MAX_SIZE, originTop - CANVAS_MAX_SIZE];
    }

    /**
     * ヘッダー行を解析し、BOM・前後空白を除いた列名を返す
     * @param {string} headerLine - ヘッダー行の文字列
     * @param {string} delimiter - 区切り文字
     * @returns {string[]} 正規化した列名
     */
    function parseColumnNames(headerLine, delimiter) {
        var columnNames = parseDataLine(headerLine, delimiter);
        for (var i = 0; i < columnNames.length; i++) {
            columnNames[i] = trimAndStripBom(columnNames[i]);
        }
        return columnNames;
    }

    /**
     * 列名がひとつでも入っているか判定する（空ファイル対策）
     * @param {string[]} columnNames - 列名の配列
     * @returns {boolean} ひとつでも空でなければtrue
     */
    function hasAnyColumnName(columnNames) {
        for (var i = 0; i < columnNames.length; i++) {
            if (columnNames[i] !== "") return true;
        }
        return false;
    }

    /**
     * ヘッダー行より後を解析し、前後空白を除いたデータ行の配列を返す
     * @param {string[]} fileLines - ファイルを行分割した配列
     * @param {string} delimiter - 区切り文字
     * @returns {Array<string[]>} データ行の配列
     */
    function parseDataRows(fileLines, delimiter) {
        var parsedRows = [];
        for (var i = 1; i < fileLines.length; i++) {
            if (fileLines[i] === "") continue;
            var cellValues = parseDataLine(fileLines[i], delimiter);
            for (var j = 0; j < cellValues.length; j++) {
                cellValues[j] = String(cellValues[j]).replace(/^\s+|\s+$/g, "");
            }
            parsedRows.push(cellValues);
        }
        return parsedRows;
    }

    /**
     * アートボード名の参照列ドロップダウンを作り直す
     * @returns {void}
     */
    function rebuildColumnDropdown() {
        artboardNameColumnIndex = 0;
        /* removeAll() は項目の表示が壊れることがあるため末尾から個別に削除する / removeAll() can corrupt item display */
        while (artboardNameColumnDropdown.items.length > 0) {
            artboardNameColumnDropdown.remove(artboardNameColumnDropdown.items[artboardNameColumnDropdown.items.length - 1]);
        }
        for (var i = 0; i < headerNames.length; i++) {
            artboardNameColumnDropdown.add("item", headerNames[i]);
        }
        if (headerNames.length > 0) artboardNameColumnDropdown.selection = 0;
    }

    /**
     * データ行から表示用のセル文字列を取り出す（"0" を空扱いしない）
     * @param {string[]} cellValues - 1行分のデータ値
     * @param {number} cellIndex - 取り出す列の番号
     * @returns {string} セルの文字列（無ければ空文字）
     */
    function cellText(cellValues, cellIndex) {
        return (cellValues[cellIndex] != null) ? String(cellValues[cellIndex]) : "";
    }

    /**
     * データ一覧リストを作り直す
     * @returns {void}
     */
    function rebuildDataListBox() {
        if (dataListBox !== null) { dataListGroup.remove(dataListBox); dataListBox = null; }
        if (headerNames.length === 0) return;

        /* ScriptUIに渡す配列は複製する（headerNamesを共有するとUI側から書き換わりうる）/ Pass copies to ScriptUI */
        var columnTitles = [];
        for (var titleIndex = 0; titleIndex < headerNames.length; titleIndex++) {
            columnTitles.push(headerNames[titleIndex]);
        }
        dataListBox = dataListGroup.add("listbox", [DATA_LIST_BOUNDS[0], DATA_LIST_BOUNDS[1], DATA_LIST_BOUNDS[2], DATA_LIST_BOUNDS[3]], [], {
            numberOfColumns: columnTitles.length, showHeaders: true, columnTitles: columnTitles
        });
        for (var i = 0; i < dataRows.length; i++) {
            var cellValues = dataRows[i];
            var dataListItem = dataListBox.add("item", cellText(cellValues, 0));
            for (var j = 1; j < cellValues.length && j < headerNames.length; j++) {
                dataListItem.subItems[j - 1].text = cellText(cellValues, j);
            }
        }
    }

    /**
     * 読み込み済みのヘッダー・データと、それを表示しているUIを空にする
     * @returns {void}
     */
    function clearLoadedData() {
        headerNames = [];
        dataRows = [];
        rebuildColumnDropdown();
        rebuildDataListBox();
        mainDialog.layout.layout(true);
    }

    /**
     * 選択ファイルを読み込み、ヘッダー・データ・データ一覧を再構築する
     * @param {File} dataFile - 読み込むデータファイル
     * @returns {void}
     */
    function loadDataFile(dataFile) {
        var fileContent = readDataFileText(dataFile);
        if (fileContent === null) {
            clearLoadedData();
            alert(getLabel("alert.fileOpenFailed"));
            return;
        }

        var fileLines = fileContent.split(/[\r\n]+/);
        var delimiter = dataFile.name.toLowerCase().match(/\.csv$/) ? "," : "\t";
        var columnNames = parseColumnNames(fileLines[0], delimiter);
        if (!hasAnyColumnName(columnNames)) {
            clearLoadedData();
            alert(getLabel("alert.emptyFile"));
            return;
        }

        headerNames = columnNames;
        dataRows = parseDataRows(fileLines, delimiter);
        rebuildColumnDropdown();
        rebuildDataListBox();
        mainDialog.layout.layout(true);
    }

    dataFileDropdown.onChange = function () {
        if (dataFileDropdown.selection) loadDataFile(dataFiles[dataFileDropdown.selection.index]);
    };
    dataFileDropdown.selection = 0;

    // =========================================
    // 実行処理（流し込み）/ Run handler (data import)
    // =========================================

    /**
     * 流し込み用の配置情報を確定する（収まらなければ警告してnullを返す）
     * @returns {{columnCount: number, rowCount: number, artboardGap: number, nameColumnIndex: number}} 配置情報（収まらなければnull）
     */
    function resolveImportLayout() {
        var artboardLayout = computeArtboardLayout();
        if (!artboardLayout.fits) {
            var maxArtboardCapacity = Math.min(
                calcMaxArtboardColumns(artboardLayout.artboardGap) * calcMaxArtboardRows(artboardLayout.artboardGap),
                MAX_ARTBOARD_COUNT
            );
            alert(getLabel("alert.tooManyData").replace("#count#", String(dataRows.length)).replace("#max#", String(maxArtboardCapacity)));
            return null;
        }
        artboardLayout.nameColumnIndex = artboardNameColumnIndex;
        return artboardLayout;
    }

    /**
     * アートボード0上の雛形オブジェクトを集める
     * 選択は複製操作で変化しうるので、配列に控えてから返す
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {PageItem[]} 雛形オブジェクト（見つからなければ空配列）
     */
    function collectTemplateItems(targetDocument) {
        targetDocument.artboards.setActiveArtboardIndex(0);
        targetDocument.selectObjectsOnActiveArtboard();
        var selectedItems = targetDocument.selection;
        var templateItems = [];
        if (!selectedItems) return templateItems;
        for (var i = 0; i < selectedItems.length; i++) {
            templateItems.push(selectedItems[i]);
        }
        return templateItems;
    }

    /**
     * 雛形アートボードと中身を、グリッドの左上（カンバス中央寄せ）へ移動する
     * @param {Artboard} templateArtboard - 雛形のアートボード
     * @param {PageItem[]} templateItems - 雛形オブジェクト
     * @param {{columnCount: number, artboardGap: number}} importLayout - 配置情報
     * @returns {number[]} 移動後の雛形アートボード矩形（グリッド配置の基準）
     */
    function moveTemplateToGridOrigin(templateArtboard, templateItems, importLayout) {
        var placementRect = templateArtboard.artboardRect;
        var cellWidth = placementRect[2] - placementRect[0];
        var cellHeight = placementRect[1] - placementRect[3];
        var dataCount = dataRows.length;
        var occupiedColumns = (importLayout.columnCount < dataCount) ? importLayout.columnCount : dataCount;
        var gridRowCount = Math.ceil(dataCount / importLayout.columnCount);
        var gridSize = calcGridSize(occupiedColumns, gridRowCount, cellWidth, cellHeight, importLayout.artboardGap);

        var gridLeft = canvasBounds[0] + Math.round(((canvasBounds[2] - canvasBounds[0]) - gridSize.width) / 2);
        var gridTop = canvasBounds[1] - Math.round(((canvasBounds[1] - canvasBounds[3]) - gridSize.height) / 2);
        var dx = gridLeft - placementRect[0];
        var dy = gridTop - placementRect[1];

        templateArtboard.artboardRect = [gridLeft, gridTop, gridLeft + cellWidth, gridTop - cellHeight];
        for (var i = 0; i < templateItems.length; i++) {
            templateItems[i].translate(dx, dy);
        }
        return templateArtboard.artboardRect;
    }

    /**
     * データ行からアートボード名を決める（値が空なら連番の既定名）
     * @param {number} dataIndex - データ行の番号（0始まり）
     * @param {number} nameColumnIndex - アートボード名に使う列の番号
     * @returns {string} アートボード名
     */
    function resolveArtboardName(dataIndex, nameColumnIndex) {
        var rowValues = dataRows[dataIndex];
        /* "0" を空扱いしないよう、|| ではなく明示的に空文字を判定する / "0" must not be treated as empty */
        var nameValue = (rowValues && rowValues.length > nameColumnIndex && rowValues[nameColumnIndex] != null)
            ? String(rowValues[nameColumnIndex]) : "";
        return (nameValue !== "") ? nameValue : (ARTBOARD_NAME_PREFIX + (dataIndex + 1));
    }

    /**
     * 1データ行ぶんのアートボードと、流し込み済みオブジェクトを生成する
     * @param {Document} targetDocument - 流し込み先のドキュメント
     * @param {PageItem[]} templateItems - 雛形オブジェクト
     * @param {number[]} originRect - グリッド左上の基準矩形
     * @param {{columnCount: number, artboardGap: number, nameColumnIndex: number}} importLayout - 配置情報
     * @param {number} dataIndex - データ行の番号（0始まり）
     * @returns {void}
     */
    function buildVariation(targetDocument, templateItems, originRect, importLayout, dataIndex) {
        var cellWidth = originRect[2] - originRect[0];
        var cellHeight = originRect[1] - originRect[3];
        var offsetX = (cellWidth + importLayout.artboardGap) * (dataIndex % importLayout.columnCount);
        var offsetY = -(cellHeight + importLayout.artboardGap) * Math.floor(dataIndex / importLayout.columnCount);

        var variationArtboard = targetDocument.artboards.add([
            originRect[0] + offsetX, originRect[1] + offsetY,
            originRect[2] + offsetX, originRect[3] + offsetY
        ]);
        variationArtboard.name = resolveArtboardName(dataIndex, importLayout.nameColumnIndex);

        /* 雛形オブジェクトを複製し、対応するアートボードへ移動して流し込み / Duplicate and merge */
        for (var i = 0; i < templateItems.length; i++) {
            var duplicatedItem = templateItems[i].duplicate();
            duplicatedItem.translate(offsetX, offsetY);
            replaceTagsRecursive(duplicatedItem, headerNames, dataRows[dataIndex]);
        }
    }

    /**
     * 雛形を複製してデータを流し込む
     * @param {Document} targetDocument - 流し込み先のドキュメント
     * @param {{columnCount: number, rowCount: number, artboardGap: number, nameColumnIndex: number}} importLayout - 配置情報
     * @param {boolean} isPreview - プレビュー実行ならtrue（完了メッセージを出さない）
     * @returns {boolean} 流し込めたらtrue
     */
    function mergeDataIntoDocument(targetDocument, importLayout, isPreview) {
        targetDocument.activate();

        var templateArtboard = targetDocument.artboards[0];
        var templateItems = collectTemplateItems(targetDocument);
        if (templateItems.length === 0) {
            alert(getLabel("alert.noTemplate"));
            return false;
        }
        var originRect = moveTemplateToGridOrigin(templateArtboard, templateItems, importLayout);

        var progressWindow = new Window("palette", getLabel("progress.title"), undefined);
        var progressBar = progressWindow.add("progressbar", undefined, 0, dataRows.length);
        progressBar.preferredSize.width = PROGRESS_BAR_WIDTH;
        progressWindow.show();

        try {
            /* 重要：2件目(i=1)から先に処理する（雛形のタグを維持するため）/ Start at row 1 to keep the template tags */
            for (var i = 1; i < dataRows.length; i++) {
                progressBar.value = i;
                progressWindow.update();
                buildVariation(targetDocument, templateItems, originRect, importLayout, i);
            }

            /* 最後に1件目を雛形自身へ書き込む / Finally, merge row 0 into the template itself */
            progressBar.value = dataRows.length;
            progressWindow.update();
            templateArtboard.name = resolveArtboardName(0, importLayout.nameColumnIndex);
            for (var j = 0; j < templateItems.length; j++) {
                replaceTagsRecursive(templateItems[j], headerNames, dataRows[0]);
            }

            targetDocument.selection = null;
        } finally {
            progressWindow.close(); // 途中でエラーが出てもパレットを残さない / never leave the palette on screen
        }

        app.redraw();
        if (!isPreview) {
            alert(getLabel("alert.done")
                .replace("#count#", String(dataRows.length))
                .replace("#filename#", String(targetDocument.name)));
        }
        return true;
    }

    /* 「実行」：ドキュメントを別名保存で複製し、流し込んでダイアログを閉じる / Run: duplicate via Save As, merge, then close */
    runButton.onClick = function () {
        if (dataRows.length === 0) return;
        var importLayout = resolveImportLayout();
        if (!importLayout) return;

        /* プレビュー用ドキュメント・ファイルが残っていれば破棄してから実行 / Discard any leftover preview first */
        discardPreview();

        var importDocument;
        try {
            importDocument = duplicateDocumentBySaveAs(originalDocument);
        } catch (e) {
            alertDuplicateFailure(e);
            return;
        }
        mainDialog.close();
        mergeDataIntoDocument(importDocument, importLayout, false);
    };

    /**
     * 複製（別名保存・ファイルコピー）の失敗を通知する
     * @param {Error} err - 捕捉した例外
     * @returns {void}
     */
    function alertDuplicateFailure(err) {
        alert(getLabel("alert.dupFailed").replace("#detail#", String(err)));
    }

    /**
     * 既に開いているプレビュー用ドキュメントを探す
     * @returns {Document} 見つかったドキュメント（無ければnull）
     */
    function getOpenPreviewDocument() {
        if (!previewFile) return null;
        for (var i = 0; i < app.documents.length; i++) {
            try {
                if (app.documents[i].fullName.fsName === previewFile.fsName) return app.documents[i];
            } catch (e) { }
        }
        return null;
    }

    /**
     * 開いているプレビュー用ドキュメントを、保存せずに閉じる
     * @returns {void}
     */
    function closePreviewDocument() {
        var openedPreviewDocument = getOpenPreviewDocument();
        if (!openedPreviewDocument) return;
        try { openedPreviewDocument.close(SaveOptions.DONOTSAVECHANGES); } catch (e) { }
    }

    /**
     * プレビュー用ドキュメントを閉じ、プレビュー用ファイルを削除する
     * @returns {void}
     */
    function discardPreview() {
        closePreviewDocument();
        try {
            if (previewFile && previewFile.exists) previewFile.remove();
        } catch (e) { }
        previewFile = null;
    }

    /**
     * プレビューを作り直す
     * 既存のプレビュードキュメントがあれば一旦閉じ、元ドキュメントから複製ファイルを作り直して開く。
     * 流し込み結果は保存しないため、毎回クリーンな雛形状態から始まる。
     * @returns {boolean} プレビューを表示できたらtrue
     */
    function rebuildPreview() {
        if (dataRows.length === 0) return false;
        var importLayout = resolveImportLayout();
        if (!importLayout) return false;

        closePreviewDocument(); // 前回の流し込み結果は破棄 / discard the previous merge

        /* 元ドキュメントからプレビュー用の複製ファイルを毎回作り直す / Rebuild the preview copy every time */
        var previewDocument;
        try {
            var originalFile = originalDocument.fullName;
            if (!previewFile) previewFile = buildDuplicateFilePath(originalFile, "preview");
            if (previewFile.exists) previewFile.remove();
            if (!originalFile.copy(previewFile)) throw new Error("File copy failed");
            previewDocument = app.open(previewFile);
        } catch (e) {
            alertDuplicateFailure(e);
            previewFile = null;
            return false;
        }

        if (!mergeDataIntoDocument(previewDocument, importLayout, true)) {
            discardPreview(); // 雛形が空などで流し込めなかった / nothing was merged
            return false;
        }
        return true;
    }

    /**
     * プレビュー表示中のときだけ、プレビューを作り直す
     * @returns {void}
     */
    function refreshPreviewIfActive() {
        if (previewCheckbox.value) rebuildPreview();
    }

    /* 「プレビュー」：オンで表示、オフでプレビュー用ドキュメントを閉じる / Preview checkbox */
    previewCheckbox.onClick = function () {
        if (!previewCheckbox.value) {
            discardPreview();
            return;
        }
        if (!rebuildPreview()) previewCheckbox.value = false;
    };

    /* 「キャンセル」：プレビュー用ファイルを削除して閉じる / Cancel: remove the preview file and close */
    cancelButton.onClick = function () {
        discardPreview();
        mainDialog.close();
    };

    /* ×ボタンやESCで閉じたときもプレビュー用ファイルを残さない / Clean up on any close path */
    mainDialog.onClose = function () {
        discardPreview();
        return true; // falseを返すと閉じられなくなる / returning false would block the close
    };

    /*
       アートボード名に使う列を変更したら、選択された列番号を控えてプレビューを更新。
       selection は再描画などで null に戻ることがあるため、値そのものを保持する。
       Keep the chosen index: selection can fall back to null and silently reset the column.
    */
    artboardNameColumnDropdown.onChange = function () {
        if (artboardNameColumnDropdown.selection) {
            artboardNameColumnIndex = artboardNameColumnDropdown.selection.index;
        }
        refreshPreviewIfActive();
    };

    // =========================================
    // テキストの置換 / Text replacement
    // =========================================

    /**
     * テキストフレーム内のタグを、文字書式を保ったまま置換する
     * タグの1文字目にデータ値を流し込み、残りのタグ文字を削除することで書式を引き継ぐ
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} placeholderTag - 置換する `<ヘッダー名>` 形式のタグ
     * @param {string} replacementValue - 流し込む値
     * @returns {void}
     */
    function replaceTagKeepingStyle(textFrame, placeholderTag, replacementValue) {
        var guardCount = 0;
        var tagPosition = textFrame.contents.indexOf(placeholderTag);
        while (tagPosition !== -1 && guardCount++ < MAX_TAG_REPLACEMENTS) {
            if (replacementValue === "") {
                for (var i = 0; i < placeholderTag.length; i++) {
                    textFrame.characters[tagPosition].remove();
                }
            } else {
                textFrame.characters[tagPosition].contents = replacementValue; // 1文字目の書式を引き継ぐ / inherit the style
                for (var j = 1; j < placeholderTag.length; j++) {
                    textFrame.characters[tagPosition + replacementValue.length].remove();
                }
            }
            tagPosition = textFrame.contents.indexOf(placeholderTag);
        }
    }

    /**
     * オブジェクトを再帰的に走査し、テキスト内の `<タグ>` をデータ値へ置換する
     * @param {PageItem} pageItem - 走査対象のオブジェクト
     * @param {string[]} columnNames - ヘッダー（列名）の配列
     * @param {string[]} rowValues - 1行分のデータ値
     * @returns {void}
     */
    function replaceTagsRecursive(pageItem, columnNames, rowValues) {
        if (!pageItem) return;

        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                replaceTagsRecursive(pageItem.pageItems[i], columnNames, rowValues);
            }
            return;
        }
        if (pageItem.typename !== "TextFrame") return;

        var originalContents = pageItem.contents;
        if (originalContents == null) return;

        var tagReplacements = collectTagReplacements(String(originalContents), columnNames, rowValues);
        if (tagReplacements.length === 0) return; // タグが無いフレームには触らない / leave untagged frames alone

        try {
            for (var j = 0; j < tagReplacements.length; j++) {
                replaceTagKeepingStyle(pageItem, tagReplacements[j].tag, tagReplacements[j].value);
            }
        } catch (e) {
            /* 文字単位で置換できない場合は一括代入にフォールバック（書式は失われる）/ Fallback: whole-contents assignment */
            pageItem.contents = applyReplacementsToText(originalContents, tagReplacements);
        }
    }

    /**
     * テキストに含まれているタグと、その置換値を集める
     * @param {string} textContents - テキストフレームの内容
     * @param {string[]} columnNames - ヘッダー（列名）の配列
     * @param {string[]} rowValues - 1行分のデータ値
     * @returns {Array<{tag: string, value: string}>} 置換の組（含まれていないタグは返さない）
     */
    function collectTagReplacements(textContents, columnNames, rowValues) {
        var tagReplacements = [];
        for (var i = 0; i < columnNames.length; i++) {
            var columnName = trimAndStripBom(columnNames[i]);
            if (columnName === "") continue;

            var placeholderTag = "<" + columnName + ">";
            if (textContents.indexOf(placeholderTag) === -1) continue;

            tagReplacements.push({ tag: placeholderTag, value: cellText(rowValues, i) });
        }
        return tagReplacements;
    }

    /**
     * 集めた置換をすべて文字列に適用する（書式を保てないときのフォールバック用）
     * @param {string} textContents - 置換前のテキスト
     * @param {Array<{tag: string, value: string}>} tagReplacements - 置換の組
     * @returns {string} 置換後のテキスト
     */
    function applyReplacementsToText(textContents, tagReplacements) {
        var mergedContents = String(textContents);
        for (var i = 0; i < tagReplacements.length; i++) {
            /* 全出現を置換（ExtendScript安全）/ Replace every occurrence */
            mergedContents = mergedContents.split(tagReplacements[i].tag).join(tagReplacements[i].value);
        }
        return mergedContents;
    }

    mainDialog.show();
})();
