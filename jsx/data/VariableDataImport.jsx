#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

CSV / タブ区切りテキストのデータを、Illustratorのテンプレートに流し込むデータ結合スクリプトです。
テキストフレーム内の `<変数名>` タグをデータ列の値に置換し、データ件数ぶんのアートボードを生成します。

詳しい機能・使い方はREADMEを参照してください。

*/

/*

### Overview

A data-merge script for Illustrator. It replaces `<tag>` placeholders inside text frames with values
from a CSV / TSV column and generates one artboard per data row.

See the README for the full feature list and usage.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "VariableDataImport";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-25";                   /* 更新日 / last updated */

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
var DEFAULT_FILE_SUFFIX = "";               /* ファイル名に挟む既定の文字列（既定は挟まない）/ default file-name suffix (none) */
/* ファイル名に使えない文字 / Characters not allowed in a file name */
var FILE_NAME_FORBIDDEN_PATTERN = /[\\\/:*?"<>|]/g;
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
var TAG_MAPPING_ROW_SPACING = 6;                /* 変数と列の対応行の行間 / mapping row spacing */
var SAMPLE_VALUE_WIDTH = 150;                   /* 対応行に出す実データの表示幅 / sample value width */
var SAMPLE_VALUE_MAX_CHARS = 18;                /* 対応行に出す実データの表示文字数 / sample value length */
var SAMPLE_VALUE_COLOR = [0.45, 0.45, 0.45, 1]; /* 実データの文字色（補助表示）/ sample value colour */
var TAX_RATE = 0.1;                             /* 消費税率（税込金額からの逆算に使う）/ consumption tax rate */
/* アートボード名に使いたい列名（先に書いたものほど優先）/ Preferred artboard-name columns, most preferred first */
var ARTBOARD_NAME_KEYWORDS = ["名前", "御中", "宛先", "様", "会社名"];
/* 税込金額が入っていそうな列名 / Column names that look like a tax-included amount */
var TAX_INCLUDED_COLUMN_PATTERN = /税込|金額|価格|料金|定価|合計|price|amount|total|cost|fee/i;
var ARTBOARD_GAP_INPUT_SIZE = [60, 25];         /* 間隔入力欄 / gap input field */
var ARTBOARD_COLUMN_INPUT_SIZE = [60, 25];      /* 列数入力欄 / column count input field */
var FILE_SUFFIX_INPUT_SIZE = [120, 25];         /* 接尾辞の入力欄 / file suffix input field */
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
        tagMapping: { ja: "変数とデータ列の対応", en: "Variable Mapping" },
        fileName: { ja: "ファイル名", en: "File Name" },
        settings: { ja: "アートボード設定", en: "Artboard Settings" },
        settingsCount: { ja: "アートボード設定（#count#）", en: "Artboard Settings (#count#)" }
    },
    /* フィールド見出し（コロンは labelText で付与）/ Field labels (colon added by labelText) */
    fieldLabel: {
        file: { ja: "ファイル", en: "File" },
        artboardNameColumn: { ja: "アートボード名", en: "Artboard name" },
        gap: { ja: "アートボード間隔", en: "Artboard gap" },
        columnCount: { ja: "列数", en: "Columns" },
        fileBaseName: { ja: "ベース名", en: "Base name" },
        fileSuffix: { ja: "接尾辞", en: "Suffix" },
        fileSeparator: { ja: "区切り文字", en: "Separator" },
        fileNameResult: { ja: "保存名", en: "Saved as" }
    },
    /* モード切替のラジオボタン / Mode radio buttons */
    radio: {
        autoMatch: { ja: "列名で自動対応", en: "Match by column name" },
        manualMapping: { ja: "手動で対応づけ", en: "Map manually" },
        hyphen: { ja: "ハイフン", en: "Hyphen" },
        underscore: { ja: "アンダースコア", en: "Underscore" }
    },
    /* ドロップダウン項目 / Dropdown items */
    dropdown: {
        noColumn: { ja: "（対応なし）", en: "(none)" },
        netSuffix: { ja: "（税抜）", en: " (excl. tax)" },
        taxSuffix: { ja: "（税額）", en: " (tax)" }
    },
    /* パネル内の補足表示 / Inline messages */
    message: {
        noTags: { ja: "アートボード1に <変数名> が見つかりません。", en: "No <tag> found on artboard 1." }
    },
    /* チェックボックス / Checkboxes */
    checkbox: {
        preview: { ja: "プレビュー", en: "Preview" },
        taxCalc: { ja: "消費税を自動計算", en: "Calculate consumption tax" },
        fileDate: { ja: "日付を付ける", en: "Append date" },
        fileTime: { ja: "時刻を付ける", en: "Append time" }
    },
    /* ボタン / Buttons */
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        run: { ja: "複製して流し込む", en: "Duplicate and Merge" }
    },
    /* 進捗表示 / Progress window */
    progress: {
        title: { ja: "流し込み中…", en: "Merging…" }
    },
    /* ヘルプチップ / Tooltips */
    tooltip: {
        file: {
            ja: "開いているドキュメントと同じフォルダーにある CSV / TSV ファイルを選びます。",
            en: "Choose a CSV / TSV file in the same folder as the open document."
        },
        tagMapping: {
            ja: "カンバス上の <変数名> に流し込むデータ列を選びます。\n名前が一致する列は自動で選ばれます。\n変更するとプレビューは解除されます。",
            en: "Choose the data column merged into each <tag> on the canvas.\nColumns with a matching name are selected automatically.\nChanging it turns the preview off."
        },
        mappingMode: {
            ja: "列名で自動対応：<変数名> と同じ名前のデータ列をそのまま使います。\n手動で対応づけ：変数ごとに使うデータ列を選びます。",
            en: "Match by column name: use the data column whose name matches each <tag>.\nMap manually: choose the data column for each variable."
        },
        taxCalc: {
            ja: "税込金額の列から、税抜価格と税額を計算した項目をポップアップメニューに足します。\n税抜＝税込÷1.1の四捨五入、税額＝税込−税抜です。\n税込金額の列が1つに絞れないときは使えません。",
            en: "Adds derived items to the popup menus: the amount excluding tax, and the tax itself.\nExcl. tax = round(total / 1.1); tax = total - excl. tax.\nUnavailable unless a single tax-included column can be identified."
        },
        dataList: {
            ja: "読み込んだデータの一覧です。\n1行目がヘッダー行、2行目以降が1件ずつアートボードになります。",
            en: "The imported data.\nLine 1 is the header row; each line after it becomes one artboard."
        },
        artboardNameSample: {
            ja: "各アートボードに実際に付く名前です。\n値が空の行は Data_1、Data_2… になります。",
            en: "The name each artboard actually gets.\nRows with an empty value fall back to Data_1, Data_2, and so on."
        },
        sampleValue: {
            ja: "選んだ列に入っている、データの1件目（ファイルの2行目）の値です。",
            en: "The value of the first data row (line 2 of the file) in the selected column."
        },
        artboardNameColumn: {
            ja: "各アートボード名に使うデータ列を選びます。\n「名前」「御中」「宛先」「様」「会社名」を含む列があれば、それを初期値にします。",
            en: "Choose the data column used for each artboard name.\nA column whose name contains 名前 / 御中 / 宛先 / 様 / 会社名 is preselected."
        },
        columnCount: {
            ja: "横に並べるアートボードの数です。\n空欄にして確定すると、グリッドが正方形に近くなる列数が入ります。\n↑↓キーで増減できます（shift併用で10単位）。\nカンバスに収まらない値は、収まる最大数に丸めます。",
            en: "How many artboards to place side by side.\nLeave it empty and commit to fill in the count that makes the grid closest to a square.\nStep it with the arrow keys (hold shift for tens).\nValues that do not fit the canvas are clamped to the maximum that does."
        },
        fileBaseName: {
            ja: "複製して保存するファイル名の本体です。\n空にすると元のファイル名を使います。",
            en: "The base of the duplicated file's name.\nLeave it empty to reuse the original file name."
        },
        fileSuffix: {
            ja: "ベース名の後ろに、区切り文字を挟んで付ける文字列です。\n空にすると付けません。",
            en: "Appended after the base name, joined with the separator.\nLeave it empty to omit it."
        },
        fileDate: {
            ja: "ファイル名の末尾に保存日（YYYYMMDD）を付けます。",
            en: "Appends the save date (YYYYMMDD) to the end of the file name."
        },
        fileTime: {
            ja: "ファイル名の末尾に保存時刻（HHMMSS）を付けます。",
            en: "Appends the save time (HHMMSS) to the end of the file name."
        },
        fileSeparator: {
            ja: "ベース名・接尾辞・日付・時刻をつなぐ文字です。",
            en: "The character that joins the base name, suffix, date, and time."
        },
        fileNameResult: {
            ja: "この名前で、元ファイルと同じフォルダーに複製して保存します。",
            en: "The duplicate is saved under this name, in the original file's folder."
        },
        gap: {
            ja: "複製するアートボード同士の間隔です。縦横とも同じ間隔で、単位はptです。\n↑↓キーで増減できます（shift併用で10単位）。\n確定時に10pt単位へ丸めます。",
            en: "The gap between duplicated artboards, applied both horizontally and vertically, in points.\nStep it with the arrow keys (hold shift for tens).\nThe value is rounded to 10 pt increments when committed."
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
        noTemplate: { ja: "アートボード1に、ロックも非表示もされていないオブジェクトがありません。", en: "No unlocked, visible objects found on artboard 1." },
        emptyFile: { ja: "選択したファイルにヘッダー行がありません。", en: "The selected file has no header row." },
        fileOpenFailed: { ja: "選択したファイルを開けませんでした。", en: "Could not open the selected file." },
        done: {
            ja: "完了しました。\n#count# 件を処理し、次のファイルに保存しました。\n\n#filename#",
            en: "Done.\nProcessed #count# rows and saved to:\n\n#filename#"
        },
        sameAsOriginal: {
            ja: "保存名が元のファイルと同じです。\n元ファイルを上書きしてしまうため実行できません。\n［ファイル名］で接尾辞・日付・時刻のいずれかを設定してください。",
            en: "The saved name matches the original file.\nRunning would overwrite the original, so it is blocked.\nSet a suffix, date, or time under File Name."
        },
        overwrite: {
            ja: "次のファイルはすでに存在します。上書きしますか？\n\n#filename#",
            en: "This file already exists. Overwrite it?\n\n#filename#"
        },
        dupFailed: {
            ja: "ドキュメントの複製に失敗しました。\n\n#detail#",
            en: "Failed to duplicate the document.\n\n#detail#"
        },
        tooManyData: {
            ja: "データ件数（#count# 件）がカンバスに収まりません。\n現在の設定では最大 #max# 件まで配置できます。\nアートボード間隔を小さくするか、［列数］を空欄に戻すか、データ件数を減らしてください。",
            en: "The number of rows (#count#) does not fit on the canvas.\nUp to #max# artboards can be placed with the current settings.\nReduce the artboard gap, clear the Columns field, or reduce the number of rows."
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
/**
 * 入力欄で↑↓キーによる値の増減を有効にする
 * ↑↓で±1、shift併用で10の倍数へスナップ、option併用で±0.1
 * @param {EditText} editText - 対象の入力欄
 * @returns {void}
 */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (event.keyName === "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
            if (event.keyName === "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName === "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            // 小数第1位までに丸め
            value = Math.round(value * 10) / 10;
        } else {
            // 整数に丸め
            value = Math.round(value);
        }

        editText.text = value;
    });
}

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
    var templateTagNames = [];       // 雛形に含まれる変数名 / Tag names found in the template
    var columnSources = [];          // 選べるデータ列（実データ列と計算列）/ Selectable columns (raw & derived)
    var tagSourceIndexes = [];       // 変数ごとの対応列の番号（-1は対応なし）/ Source per tag (-1 = none)
    var tagSourceDropdowns = [];     // 変数ごとの対応列ドロップダウン / Column dropdown per tag
    var tagSampleLabels = [];        // 変数ごとの実データ表示 / Sample value label per tag
    var taxCalcEnabled = false;      // 消費税を自動計算するか / Whether to derive the tax columns
    var taxCalcAvailable = false;    // 税込金額の列を特定できたか / Whether a tax-included column exists

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
    var originalFileName = decodeURI(originalDocument.fullName.name);
    var extensionDotIndex = originalFileName.lastIndexOf(".");
    var originalBaseName = (extensionDotIndex >= 0) ? originalFileName.substring(0, extensionDotIndex) : originalFileName;

    var templateArtboardWidth = templateArtboardRect[2] - templateArtboardRect[0];
    var templateArtboardHeight = templateArtboardRect[1] - templateArtboardRect[3];

    /* 雛形に含まれる <変数名> を集める / Collect the <tag> names in the template */
    templateTagNames = collectTemplateTagNames(originalDocument);

    // =========================================
    // ダイアログUIの構築 / Build the dialog UI
    // =========================================

    /**
     * 補助表示用に、長い文字列を切り詰める
     * @param {string} text - 対象の文字列
     * @returns {string} 表示用の文字列
     */
    function truncateForDisplay(text) {
        var displayText = String(text);
        return (displayText.length > SAMPLE_VALUE_MAX_CHARS) ? (displayText.substring(0, SAMPLE_VALUE_MAX_CHARS) + "…") : displayText;
    }

    /**
     * 実データ（データの1件目）を出す補助表示を、行の右端に足す
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} sampleText - 表示する文字列
     * @param {string} tooltipKey - ヘルプチップの階層キー
     * @returns {StaticText} 追加した補助表示
     */
    function addSampleValueLabel(parentGroup, sampleText, tooltipKey) {
        var sampleValueLabel = parentGroup.add("statictext", undefined, sampleText);
        sampleValueLabel.preferredSize.width = SAMPLE_VALUE_WIDTH;
        sampleValueLabel.helpTip = getLabel(tooltipKey);
        applySampleValueColor(sampleValueLabel);
        return sampleValueLabel;
    }

    /**
     * 補助表示であることが分かるよう、文字色を淡くする
     * @param {StaticText} targetLabel - 対象の表示
     * @returns {void}
     */
    function applySampleValueColor(targetLabel) {
        var labelGraphics = targetLabel.graphics;
        labelGraphics.foregroundColor = labelGraphics.newPen(labelGraphics.PenType.SOLID_COLOR, SAMPLE_VALUE_COLOR, 1);
    }

    /**
     * 設定パネル用に、一定幅のコロン付きラベルを追加して項目を縦に揃える
     * @param {Group} parentGroup - ラベルを追加するグループ
     * @param {string} labelKey - ラベルの階層キー
     * @returns {StaticText} 追加したラベル
     */
    function addFieldLabel(parentGroup, labelKey) {
        /* justify は生成時にしか効かない / justify only takes effect at creation time */
        var fieldLabel = parentGroup.add("statictext", undefined, labelText(labelKey), { justify: "right" });
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH[uiLang] || FIELD_LABEL_WIDTH.en;
        return fieldLabel;
    }

    var mainDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    mainDialog.orientation = "column";
    mainDialog.alignChildren = ["fill", "top"];
    mainDialog.spacing = DIALOG_SPACING;
    mainDialog.margins = DIALOG_MARGINS;

    /* 対応づけの方法：自動認識か手動照合か / Mapping mode: automatic or manual */
    var mappingModeGroup = mainDialog.add("group");
    mappingModeGroup.orientation = "row";
    mappingModeGroup.alignment = ["center", "top"];
    mappingModeGroup.alignChildren = ["left", "center"];
    var autoMatchRadio = mappingModeGroup.add("radiobutton", undefined, getLabel("radio.autoMatch"));
    autoMatchRadio.helpTip = getLabel("tooltip.mappingMode");
    var manualMappingRadio = mappingModeGroup.add("radiobutton", undefined, getLabel("radio.manualMapping"));
    manualMappingRadio.helpTip = getLabel("tooltip.mappingMode");
    autoMatchRadio.value = true;

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

    /*
       変数と列の対応パネルの置き場所。
       自動認識のときはパネルごと外すので、枠だけを先に作って場所を押さえておく。
       Keep an empty host here: the panel itself is removed in automatic mode so it takes no height.
    */
    var tagMappingHost = mainDialog.add("group");
    tagMappingHost.orientation = "column";
    tagMappingHost.alignment = ["fill", "top"];
    tagMappingHost.alignChildren = ["fill", "top"];
    tagMappingHost.margins = 0;
    tagMappingHost.spacing = 0;

    var tagMappingPanel = null;
    var tagMappingGroup = null;
    var taxCalcCheckbox = null;

    var artboardSettingsPanel = mainDialog.add("panel", undefined, getLabel("panel.settings"));
    artboardSettingsPanel.orientation = "column";
    artboardSettingsPanel.alignChildren = ["left", "top"];
    artboardSettingsPanel.margins = PANEL_MARGINS;

    var artboardNameColumnGroup = artboardSettingsPanel.add("group");
    addFieldLabel(artboardNameColumnGroup, "fieldLabel.artboardNameColumn");
    var artboardNameColumnDropdown = artboardNameColumnGroup.add("dropdownlist", undefined, []);
    artboardNameColumnDropdown.size = COLUMN_DROPDOWN_SIZE;
    artboardNameColumnDropdown.helpTip = getLabel("tooltip.artboardNameColumn");
    var artboardNameSampleLabel = addSampleValueLabel(artboardNameColumnGroup, "", "tooltip.artboardNameSample");

    /* グリッド配置設定（アートボード間隔）/ Grid layout settings (artboard gap) */
    var defaultArtboardGap = Math.round(templateArtboardWidth / ARTBOARD_GAP_DIVISOR / ARTBOARD_GAP_STEP) * ARTBOARD_GAP_STEP;

    var artboardGapGroup = artboardSettingsPanel.add("group");
    addFieldLabel(artboardGapGroup, "fieldLabel.gap");
    var artboardGapInput = artboardGapGroup.add("edittext", undefined, String(defaultArtboardGap));
    artboardGapInput.size = ARTBOARD_GAP_INPUT_SIZE;
    artboardGapInput.helpTip = getLabel("tooltip.gap");
    changeValueByArrowKey(artboardGapInput);
    artboardGapGroup.add("statictext", undefined, "pt");

    var artboardColumnCountGroup = artboardSettingsPanel.add("group");
    addFieldLabel(artboardColumnCountGroup, "fieldLabel.columnCount");
    var artboardColumnCountInput = artboardColumnCountGroup.add("edittext", undefined, "");
    artboardColumnCountInput.size = ARTBOARD_COLUMN_INPUT_SIZE;
    artboardColumnCountInput.helpTip = getLabel("tooltip.columnCount");
    changeValueByArrowKey(artboardColumnCountInput);

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
     * 列数の入力値を取得する（空欄や不正な値は0＝自動、収まらない値は最大数に丸める）
     * @returns {number} 列数（自動なら0）
     */
    function getArtboardColumnCount() {
        var columnCount = parseInt(artboardColumnCountInput.text, 10);
        if (isNaN(columnCount) || columnCount < 1) return 0;
        var maxFitColumns = calcMaxArtboardColumns(getArtboardGap());
        return (columnCount > maxFitColumns) ? maxFitColumns : columnCount;
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

        /* 列数の指定があれば、正方形に近づける探索はせずそのまま使う / an explicit column count wins */
        var requestedColumnCount = getArtboardColumnCount();
        if (requestedColumnCount >= 1) {
            var requestedRowCount = Math.ceil(dataCount / requestedColumnCount);
            if (requestedRowCount <= maxFitRows) {
                artboardLayout.columnCount = requestedColumnCount;
                artboardLayout.rowCount = requestedRowCount;
                artboardLayout.fits = true;
            }
            return artboardLayout;
        }

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

    /**
     * アートボード設定パネルの見出しに、作られるアートボード数を出す
     * @returns {void}
     */
    function refreshArtboardSettingsPanelTitle() {
        artboardSettingsPanel.text = (dataRows.length > 0)
            ? getLabel("panel.settingsCount").replace("#count#", String(dataRows.length))
            : getLabel("panel.settings");
    }

    /**
     * データを読み込み直したとき、列数の入力欄を自動計算の値に戻す
     * @returns {void}
     */
    function resetArtboardColumnCountInput() {
        artboardColumnCountInput.text = "";
        if (dataRows.length === 0) return;
        var autoLayout = computeArtboardLayout();
        if (autoLayout.fits) artboardColumnCountInput.text = String(autoLayout.columnCount);
    }

    /**
     * 入力確定時に、列数を収まる範囲へ丸めて入力欄へ反映する（空欄なら自動計算の値に戻す）
     * @returns {void}
     */
    function normalizeArtboardColumnCountInput() {
        var columnCount = getArtboardColumnCount();
        if (columnCount >= 1) {
            artboardColumnCountInput.text = String(columnCount);
            return;
        }
        resetArtboardColumnCountInput();
    }

    artboardGapInput.onChange = function () {
        roundArtboardGapInput();
        normalizeArtboardColumnCountInput();
        refreshPreviewIfActive();
    };

    artboardColumnCountInput.onChange = function () {
        normalizeArtboardColumnCountInput();
        refreshPreviewIfActive();
    };

    /* ファイル名パネル：複製して保存する名前を決める / File-name panel: name of the duplicated file */
    var fileNamePanel = mainDialog.add("panel", undefined, getLabel("panel.fileName"));
    fileNamePanel.orientation = "column";
    fileNamePanel.alignChildren = ["fill", "top"];
    fileNamePanel.margins = PANEL_MARGINS;
    fileNamePanel.spacing = PANEL_SPACING;

    var fileBaseNameGroup = fileNamePanel.add("group");
    fileBaseNameGroup.alignment = ["fill", "top"];
    fileBaseNameGroup.alignChildren = ["left", "center"];
    addFieldLabel(fileBaseNameGroup, "fieldLabel.fileBaseName");
    var fileBaseNameInput = fileBaseNameGroup.add("edittext", undefined, originalBaseName);
    /* 幅は指定せず、パネル幅に合わせて伸ばす / stretch instead of widening the dialog */
    fileBaseNameInput.alignment = ["fill", "center"];
    fileBaseNameInput.helpTip = getLabel("tooltip.fileBaseName");

    var fileSuffixGroup = fileNamePanel.add("group");
    fileSuffixGroup.alignment = ["left", "top"];
    fileSuffixGroup.alignChildren = ["left", "center"];
    addFieldLabel(fileSuffixGroup, "fieldLabel.fileSuffix");
    var fileSuffixInput = fileSuffixGroup.add("edittext", undefined, DEFAULT_FILE_SUFFIX);
    fileSuffixInput.size = FILE_SUFFIX_INPUT_SIZE;
    fileSuffixInput.helpTip = getLabel("tooltip.fileSuffix");
    var fileDateCheckbox = fileSuffixGroup.add("checkbox", undefined, getLabel("checkbox.fileDate"));
    fileDateCheckbox.helpTip = getLabel("tooltip.fileDate");
    fileDateCheckbox.value = true;
    var fileTimeCheckbox = fileSuffixGroup.add("checkbox", undefined, getLabel("checkbox.fileTime"));
    fileTimeCheckbox.helpTip = getLabel("tooltip.fileTime");
    fileTimeCheckbox.value = true;

    var fileSeparatorGroup = fileNamePanel.add("group");
    fileSeparatorGroup.alignment = ["left", "top"];
    fileSeparatorGroup.alignChildren = ["left", "center"];
    addFieldLabel(fileSeparatorGroup, "fieldLabel.fileSeparator");
    var hyphenSeparatorRadio = fileSeparatorGroup.add("radiobutton", undefined, getLabel("radio.hyphen"));
    hyphenSeparatorRadio.helpTip = getLabel("tooltip.fileSeparator");
    var underscoreSeparatorRadio = fileSeparatorGroup.add("radiobutton", undefined, getLabel("radio.underscore"));
    underscoreSeparatorRadio.helpTip = getLabel("tooltip.fileSeparator");
    underscoreSeparatorRadio.value = true;

    var fileNameResultGroup = fileNamePanel.add("group");
    fileNameResultGroup.alignment = ["fill", "top"];
    fileNameResultGroup.alignChildren = ["left", "center"];
    addFieldLabel(fileNameResultGroup, "fieldLabel.fileNameResult");
    var fileNameResultLabel = fileNameResultGroup.add("statictext", undefined, "");
    fileNameResultLabel.alignment = ["fill", "center"];
    fileNameResultLabel.helpTip = getLabel("tooltip.fileNameResult");
    applySampleValueColor(fileNameResultLabel);

    /**
     * 保存されるファイル名の表示を更新する
     * @returns {void}
     */
    function refreshFileNameResult() {
        fileNameResultLabel.text = decodeURI(buildDuplicateFilePath(originalDocument.fullName, fileSuffixInput.text, fileDateCheckbox.value, fileTimeCheckbox.value).name);
    }

    fileBaseNameInput.onChanging = refreshFileNameResult;
    fileSuffixInput.onChanging = refreshFileNameResult;
    hyphenSeparatorRadio.onClick = refreshFileNameResult;
    underscoreSeparatorRadio.onClick = refreshFileNameResult;
    fileDateCheckbox.onClick = refreshFileNameResult;
    fileTimeCheckbox.onClick = refreshFileNameResult;
    refreshFileNameResult();

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
     * 1桁の数値を0詰めで2桁にする
     * @param {number} numberValue - 対象の数値
     * @returns {string} 2桁の文字列
     */
    function padTwoDigits(numberValue) {
        return (numberValue < 10 ? "0" : "") + String(numberValue);
    }

    /**
     * 現在の日付を YYYYMMDD の形にする
     * @returns {string} 日付の文字列
     */
    function buildDateStamp() {
        var now = new Date();
        return String(now.getFullYear()) + padTwoDigits(now.getMonth() + 1) + padTwoDigits(now.getDate());
    }

    /**
     * 現在の時刻を HHMMSS の形にする
     * @returns {string} 時刻の文字列
     */
    function buildTimeStamp() {
        var now = new Date();
        return padTwoDigits(now.getHours()) + padTwoDigits(now.getMinutes()) + padTwoDigits(now.getSeconds());
    }

    /**
     * ファイル名に使えない文字と前後の空白を取り除く
     * @param {string} text - 対象の文字列
     * @returns {string} ファイル名に使える文字列
     */
    function sanitizeFileNamePart(text) {
        return String(text).replace(FILE_NAME_FORBIDDEN_PATTERN, "").replace(/^\s+|\s+$/g, "");
    }

    /**
     * 保存するファイル名の本体（拡張子なし）を組み立てる
     * ベース名が空のときは元のファイル名を使う
     * @param {string} nameSuffix - ベース名の後ろに挟む文字列（空なら挟まない）
     * @param {boolean} useDate - 末尾に日付を付けるか
     * @param {boolean} useTime - 末尾に時刻を付けるか
     * @returns {string} 拡張子を除いたファイル名
     */
    function buildOutputBaseName(nameSuffix, useDate, useTime) {
        var partSeparator = hyphenSeparatorRadio.value ? "-" : "_";
        var baseName = sanitizeFileNamePart(fileBaseNameInput.text);
        if (baseName === "") baseName = originalBaseName;

        var nameParts = [baseName];
        var sanitizedSuffix = sanitizeFileNamePart(nameSuffix);
        if (sanitizedSuffix !== "") nameParts.push(sanitizedSuffix);
        if (useDate) nameParts.push(buildDateStamp());
        if (useTime) nameParts.push(buildTimeStamp());
        return nameParts.join(partSeparator);
    }

    /**
     * 複製ファイルのパスを生成する（元ファイルと同じフォルダー・同じ拡張子）
     * @param {File} originalFile - 元ファイル
     * @param {string} nameSuffix - ベース名の後ろに挟む文字列
     * @param {boolean} useDate - 末尾に日付を付けるか
     * @param {boolean} useTime - 末尾に時刻を付けるか
     * @returns {File} 複製先のファイル
     */
    function buildDuplicateFilePath(originalFile, nameSuffix, useDate, useTime) {
        var sourceFileName = decodeURI(originalFile.name);
        var dotIndex = sourceFileName.lastIndexOf(".");
        var fileExtension = (dotIndex >= 0) ? sourceFileName.substring(dotIndex) : ".ai";
        /* 名前側だけURIエンコードする（fsNameは生のパス）/ encode only the name part; fsName is a raw path */
        return new File(originalFile.parent.fsName + "/" + File.encode(buildOutputBaseName(nameSuffix, useDate, useTime) + fileExtension));
    }

    /**
     * 指定ドキュメントを別名保存し、操作対象として複製後ドキュメントを返す
     * @param {Document} sourceDocument - 複製元のドキュメント
     * @param {File} duplicateFile - 保存先のファイル
     * @returns {Document} 別名保存後のドキュメント
     */
    function duplicateDocumentBySaveAs(sourceDocument, duplicateFile) {
        sourceDocument.saveAs(duplicateFile); // Illustrator switches this document to the duplicated file
        return sourceDocument;
    }

    /**
     * 保存先が元ファイルや既存ファイルを壊さないか確かめる
     * 元ファイルと同じ名前になる場合は中止し、別の既存ファイルなら上書きの可否を尋ねる
     * @param {File} outputFile - 保存先のファイル
     * @returns {boolean} 保存してよければtrue
     */
    function confirmOutputFile(outputFile) {
        if (outputFile.fsName === originalDocument.fullName.fsName) {
            alert(getLabel("alert.sameAsOriginal"));
            return false;
        }
        if (!outputFile.exists) return true;
        return confirm(getLabel("alert.overwrite").replace("#filename#", decodeURI(outputFile.name)));
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
        var rawByteString = readFileWithEncoding(dataFile, "BINARY");
        if (rawByteString === null) return null;
        return readFileWithEncoding(dataFile, isValidUtf8(rawByteString) ? "UTF-8" : "SJIS");
    }

    /**
     * 指定のエンコーディングでファイル全体を読む
     * @param {File} dataFile - 読み込むファイル
     * @param {string} encoding - ExtendScriptのエンコーディング名（"BINARY" / "UTF-8" / "SJIS"）
     * @returns {string} ファイル全体の文字列（開けなければnull）
     */
    function readFileWithEncoding(dataFile, encoding) {
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
        var probeLayer = measuredDocument.layers.add();
        var probeTextFrame = probeLayer.textFrames.add();
        var canvasLeft = probeTextFrame.matrix.mValueTX;
        var canvasTop = probeTextFrame.matrix.mValueTY;
        probeLayer.remove();
        measuredDocument.modified = wasModified; // 一時レイヤーで立った変更フラグを元に戻す / restore the flag
        return [canvasLeft, canvasTop, canvasLeft + CANVAS_MAX_SIZE, canvasTop - CANVAS_MAX_SIZE];
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
        /* removeAll() は項目の表示が壊れることがあるため末尾から個別に削除する / removeAll() can corrupt item display */
        while (artboardNameColumnDropdown.items.length > 0) {
            artboardNameColumnDropdown.remove(artboardNameColumnDropdown.items[artboardNameColumnDropdown.items.length - 1]);
        }
        for (var i = 0; i < headerNames.length; i++) {
            artboardNameColumnDropdown.add("item", headerNames[i]);
        }
        artboardNameColumnIndex = findPreferredArtboardNameColumnIndex();
        if (headerNames.length > 0) artboardNameColumnDropdown.selection = artboardNameColumnIndex;
        refreshArtboardNameSample();
    }

    /**
     * アートボード名に使う列の初期値を決める
     * 「名前」「御中」などを含む列名を優先し、無ければ先頭の列を使う
     * @returns {number} 列の番号
     */
    function findPreferredArtboardNameColumnIndex() {
        for (var i = 0; i < ARTBOARD_NAME_KEYWORDS.length; i++) {
            for (var j = 0; j < headerNames.length; j++) {
                if (String(headerNames[j]).indexOf(ARTBOARD_NAME_KEYWORDS[i]) !== -1) return j;
            }
        }
        return 0;
    }

    /**
     * アートボード名に使う列の実データ表示を、実際に付く名前で更新する
     * @returns {void}
     */
    function refreshArtboardNameSample() {
        artboardNameSampleLabel.text = (dataRows.length > 0 && headerNames.length > 0)
            ? truncateForDisplay(resolveArtboardName(0, artboardNameColumnIndex)) : "";
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
        dataListBox.helpTip = getLabel("tooltip.dataList");
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
        rebuildColumnSources();
        applyAutoTagMapping();
        rebuildTagMappingRows();
        resetArtboardColumnCountInput();
        refreshArtboardSettingsPanelTitle();
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
        rebuildColumnSources();
        applyAutoTagMapping();
        rebuildTagMappingRows();
        resetArtboardColumnCountInput();
        refreshArtboardSettingsPanelTitle();
        mainDialog.layout.layout(true);
    }

    dataFileDropdown.onChange = function () {
        if (dataFileDropdown.selection) loadDataFile(dataFiles[dataFileDropdown.selection.index]);
    };
    dataFileDropdown.selection = 0;

    // =========================================
    // 変数と列の対応 / Tag-to-column mapping
    // =========================================

    /**
     * オブジェクトを再帰的に走査し、テキスト内の `<変数名>` を重複なく集める
     * @param {PageItem} pageItem - 走査対象のオブジェクト
     * @param {string[]} tagNames - 見つかった変数名を追加する配列
     * @param {Object} seenTagNames - 既出判定に使う辞書
     * @returns {void}
     */
    function collectTagNamesFromItem(pageItem, tagNames, seenTagNames) {
        if (!pageItem) return;

        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                collectTagNamesFromItem(pageItem.pageItems[i], tagNames, seenTagNames);
            }
            return;
        }
        if (pageItem.typename !== "TextFrame") return;

        var textContents = pageItem.contents;
        if (textContents == null) return;

        var tagPattern = /<([^<>\r\n]+)>/g;
        var tagMatch;
        while ((tagMatch = tagPattern.exec(String(textContents))) !== null) {
            var tagName = trimAndStripBom(tagMatch[1]);
            if (tagName === "") continue;
            /* Objectの既定プロパティと衝突しないよう接頭辞を付ける / prefix to avoid built-in keys */
            var seenKey = "tag:" + tagName;
            if (seenTagNames[seenKey]) continue;
            seenTagNames[seenKey] = true;
            tagNames.push(tagName);
        }
    }

    /**
     * 雛形（アートボード1）のテキストに含まれる変数名を集める
     * 選択とアクティブアートボードは、走査前の状態へ戻す
     * @param {Document} targetDocument - 走査するドキュメント
     * @returns {string[]} 出現順に並べた変数名
     */
    function collectTemplateTagNames(targetDocument) {
        var savedSelection = [];
        try {
            var currentSelection = targetDocument.selection;
            for (var i = 0; i < currentSelection.length; i++) savedSelection.push(currentSelection[i]);
        } catch (e) { }
        var savedArtboardIndex = targetDocument.artboards.getActiveArtboardIndex();
        var wasModified = targetDocument.modified; // 走査で立つ変更フラグを退避 / remember the flag

        var tagNames = [];
        var seenTagNames = {};
        var templateItems = collectTemplateItems(targetDocument);
        for (var j = 0; j < templateItems.length; j++) {
            collectTagNamesFromItem(templateItems[j], tagNames, seenTagNames);
        }

        try {
            targetDocument.artboards.setActiveArtboardIndex(savedArtboardIndex);
            targetDocument.selection = (savedSelection.length > 0) ? savedSelection : null;
        } catch (e) { }
        targetDocument.modified = wasModified; // 走査で立った変更フラグを元に戻す / restore the flag
        return tagNames;
    }

    /**
     * 名前を照合用に正規化する（空白・アンダースコア・ハイフンを除いて小文字化）
     * @param {string} text - 対象の文字列
     * @returns {string} 正規化した文字列
     */
    function normalizeNameForMatch(text) {
        return String(text).replace(/[\s\u3000_\-]/g, "").toLowerCase();
    }

    /**
     * 金額の文字列を数値にする（3桁区切り・通貨記号・全角数字を許容する）
     * @param {string} text - 対象の文字列
     * @returns {number} 金額（数値として読めなければnull）
     */
    function parseAmount(text) {
        var normalizedText = String(text).replace(/[０-９]/g, function (fullWidthDigit) {
            return String.fromCharCode(fullWidthDigit.charCodeAt(0) - 0xFEE0);
        });
        normalizedText = normalizedText.replace(/[,\s\u3000\u00A5\uFFE5$円]/g, "");
        if (!/^-?\d+(\.\d+)?$/.test(normalizedText)) return null;
        return Number(normalizedText);
    }

    /**
     * 金額を文字列にする（元の値が3桁区切りなら、区切りも付け直す）
     * @param {number} amount - 金額
     * @param {string} sourceText - 元になった文字列
     * @returns {string} 表示用の文字列
     */
    function formatAmount(amount, sourceText) {
        var amountText = String(amount);
        if (String(sourceText).indexOf(",") === -1) return amountText;

        var sign = (amountText.charAt(0) === "-") ? "-" : "";
        var digits = sign ? amountText.substring(1) : amountText;
        var groupedDigits = "";
        while (digits.length > 3) {
            groupedDigits = "," + digits.substring(digits.length - 3) + groupedDigits;
            digits = digits.substring(0, digits.length - 3);
        }
        return sign + digits + groupedDigits;
    }

    /**
     * 列に金額が入っているかを、最初に見つかった空でない値で判定する
     * @param {number} columnIndex - 列の番号
     * @returns {boolean} 金額として読めればtrue
     */
    function isAmountColumn(columnIndex) {
        for (var i = 0; i < dataRows.length; i++) {
            var cellValue = cellText(dataRows[i], columnIndex);
            if (cellValue === "") continue;
            return parseAmount(cellValue) !== null;
        }
        return false;
    }

    /**
     * 税込金額が入っている列を1つだけ特定する
     * 列名が金額らしいものを優先し、それが無ければ金額の列が1つだけのときにその列を使う。
     * 候補が複数あるとどれが税込か決められないため、そのときは特定しない。
     * @returns {number} 税込金額の列の番号（特定できなければ-1）
     */
    function findTaxIncludedColumnIndex() {
        var amountColumns = [];
        for (var i = 0; i < headerNames.length; i++) {
            if (isAmountColumn(i)) amountColumns.push(i);
        }
        if (amountColumns.length === 0) return -1;

        var namedColumns = [];
        for (var j = 0; j < amountColumns.length; j++) {
            if (TAX_INCLUDED_COLUMN_PATTERN.test(headerNames[amountColumns[j]])) namedColumns.push(amountColumns[j]);
        }
        if (namedColumns.length > 0) return (namedColumns.length === 1) ? namedColumns[0] : -1;
        return (amountColumns.length === 1) ? amountColumns[0] : -1;
    }

    /**
     * ドロップダウンに並べる列を作り直す
     * 税込金額の列を特定できて、かつ「消費税を自動計算」がオンのときだけ「本体」「税」も足す
     * @returns {void}
     */
    function rebuildColumnSources() {
        var taxIncludedColumnIndex = findTaxIncludedColumnIndex();
        taxCalcAvailable = (taxIncludedColumnIndex >= 0);
        if (taxCalcCheckbox) taxCalcCheckbox.enabled = taxCalcAvailable;
        columnSources = [];
        for (var i = 0; i < headerNames.length; i++) {
            columnSources.push({ label: headerNames[i], columnIndex: i, valueKind: "raw" });
            if (i !== taxIncludedColumnIndex || !taxCalcEnabled) continue;
            columnSources.push({ label: headerNames[i] + getLabel("dropdown.netSuffix"), columnIndex: i, valueKind: "net" });
            columnSources.push({ label: headerNames[i] + getLabel("dropdown.taxSuffix"), columnIndex: i, valueKind: "tax" });
        }
    }

    /**
     * 列の値を取り出す（本体・税は税込金額から逆算する）
     * 本体は端数を四捨五入し、税は税込との差にするため、本体＋税は必ず税込に一致する
     * @param {{columnIndex: number, valueKind: string}} columnSource - 対応づけた列
     * @param {string[]} rowValues - 1行分のデータ値
     * @returns {string} 流し込む文字列
     */
    function resolveSourceValue(columnSource, rowValues) {
        var rawText = cellText(rowValues, columnSource.columnIndex);
        if (columnSource.valueKind === "raw") return rawText;

        var taxIncludedAmount = parseAmount(rawText);
        if (taxIncludedAmount === null) return rawText; // 金額として読めなければそのまま / leave as-is

        var netAmount = Math.round(taxIncludedAmount / (1 + TAX_RATE));
        var resolvedAmount = (columnSource.valueKind === "tax") ? (taxIncludedAmount - netAmount) : netAmount;
        return formatAmount(resolvedAmount, rawText);
    }

    /**
     * 変数名に対応する列を探す（完全一致を優先し、無ければ正規化して照合）
     * 計算列は自動では選ばない
     * @param {string} tagName - カンバス上の変数名
     * @returns {number} 列の番号（見つからなければ-1）
     */
    function findMatchingSourceIndex(tagName) {
        for (var i = 0; i < columnSources.length; i++) {
            if (columnSources[i].valueKind === "raw" && columnSources[i].label === tagName) return i;
        }
        var normalizedTagName = normalizeNameForMatch(tagName);
        if (normalizedTagName === "") return -1;
        for (var j = 0; j < columnSources.length; j++) {
            if (columnSources[j].valueKind !== "raw") continue;
            if (normalizeNameForMatch(columnSources[j].label) === normalizedTagName) return j;
        }
        return -1;
    }

    /**
     * 読み込んだ列名をもとに、各変数の対応列を自動で選び直す
     * @returns {void}
     */
    function applyAutoTagMapping() {
        tagSourceIndexes = [];
        for (var i = 0; i < templateTagNames.length; i++) {
            tagSourceIndexes.push(findMatchingSourceIndex(templateTagNames[i]));
        }
    }

    /**
     * 選んだ列の実データ（データの1件目）を、表示用に切り詰めて返す
     * @param {number} sourceIndex - 列の番号（-1は対応なし）
     * @returns {string} 表示用の文字列（値が無ければ空文字）
     */
    function sampleValueText(sourceIndex) {
        if (sourceIndex == null || sourceIndex < 0 || sourceIndex >= columnSources.length) return "";
        if (dataRows.length === 0) return "";
        return truncateForDisplay(resolveSourceValue(columnSources[sourceIndex], dataRows[0]));
    }

    /**
     * その列が、ほかの変数ですでに使われているかを調べる
     * @param {number} sourceIndex - 調べる列の番号
     * @param {number} tagIndex - 判定の対象外にする変数の番号（自分自身）
     * @returns {boolean} ほかの変数で使われていればtrue
     */
    function isSourceUsedByOtherTag(sourceIndex, tagIndex) {
        for (var i = 0; i < tagSourceIndexes.length; i++) {
            if (i === tagIndex) continue;
            if (tagSourceIndexes[i] === sourceIndex) return true;
        }
        return false;
    }

    /**
     * ほかの変数で使っている列を、ポップアップ内で淡くして選べなくする
     * @returns {void}
     */
    function refreshUsedColumnItems() {
        for (var i = 0; i < tagSourceDropdowns.length; i++) {
            var dropdownItems = tagSourceDropdowns[i].items;
            for (var j = 0; j < columnSources.length && (j + 1) < dropdownItems.length; j++) {
                /* 先頭の「対応なし」は常に選べる / the "none" item stays selectable */
                dropdownItems[j + 1].enabled = !isSourceUsedByOtherTag(j, i);
            }
        }
    }

    /**
     * 「変数とデータ列の対応」パネルを組み立てる（すでにあれば何もしない）
     * @returns {void}
     */
    function buildTagMappingPanel() {
        if (tagMappingPanel) return;

        tagMappingPanel = tagMappingHost.add("panel", undefined, getLabel("panel.tagMapping"));
        tagMappingPanel.orientation = "column";
        tagMappingPanel.alignChildren = ["fill", "top"];
        tagMappingPanel.margins = PANEL_MARGINS;
        tagMappingPanel.spacing = TAG_MAPPING_ROW_SPACING;

        tagMappingGroup = tagMappingPanel.add("group");
        tagMappingGroup.orientation = "column";
        tagMappingGroup.alignment = ["fill", "top"];
        tagMappingGroup.alignChildren = ["left", "top"];
        tagMappingGroup.spacing = TAG_MAPPING_ROW_SPACING;

        taxCalcCheckbox = tagMappingPanel.add("checkbox", undefined, getLabel("checkbox.taxCalc"));
        taxCalcCheckbox.alignment = ["left", "top"];
        taxCalcCheckbox.helpTip = getLabel("tooltip.taxCalc");
        taxCalcCheckbox.value = taxCalcEnabled;
        taxCalcCheckbox.enabled = taxCalcAvailable;
        taxCalcCheckbox.onClick = applyTaxCalculationToggle;

        tagMappingHost.visible = true;
    }

    /**
     * 「変数とデータ列の対応」パネルを丸ごと外し、置き場所も隠して高さを0にする
     * @returns {void}
     */
    function removeTagMappingPanel() {
        if (!tagMappingPanel) return;
        tagMappingHost.remove(tagMappingPanel);
        tagMappingPanel = null;
        tagMappingGroup = null;
        taxCalcCheckbox = null;
        tagSourceDropdowns = [];
        tagSampleLabels = [];
        tagMappingHost.visible = false;
    }

    /**
     * 変数ごとの対応行（変数名＋データ列ドロップダウン＋実データ）を作り直す
     * パネルが外れている（自動認識）ときは何もしない
     * @returns {void}
     */
    function rebuildTagMappingRows() {
        tagSourceDropdowns = [];
        tagSampleLabels = [];
        if (!tagMappingGroup) return;
        while (tagMappingGroup.children.length > 0) {
            tagMappingGroup.remove(tagMappingGroup.children[0]);
        }
        if (templateTagNames.length === 0) {
            tagMappingGroup.add("statictext", undefined, getLabel("message.noTags"));
            return;
        }

        for (var i = 0; i < templateTagNames.length; i++) {
            var tagRowGroup = tagMappingGroup.add("group");
            tagRowGroup.orientation = "row";
            tagRowGroup.alignment = ["left", "top"];
            tagRowGroup.alignChildren = ["left", "center"];

            var tagNameLabel = tagRowGroup.add("statictext", undefined, "<" + templateTagNames[i] + ">", { justify: "right" });
            tagNameLabel.preferredSize.width = FIELD_LABEL_WIDTH[uiLang] || FIELD_LABEL_WIDTH.en;

            var tagSourceDropdown = tagRowGroup.add("dropdownlist", undefined, []);
            tagSourceDropdown.size = COLUMN_DROPDOWN_SIZE;
            tagSourceDropdown.helpTip = getLabel("tooltip.tagMapping");
            tagSourceDropdown.add("item", getLabel("dropdown.noColumn"));
            for (var j = 0; j < columnSources.length; j++) {
                tagSourceDropdown.add("item", columnSources[j].label);
            }
            /* 先頭が「対応なし」なので、項目番号は列番号より1つ大きい / index 0 is the "none" item */
            var sourceIndex = tagSourceIndexes[i];
            tagSourceDropdown.selection = (sourceIndex >= 0 && sourceIndex < columnSources.length) ? (sourceIndex + 1) : 0;
            tagSourceDropdowns.push(tagSourceDropdown);

            /* 選んだ列に実際に入っている値（データの1件目）を並べて出す / Show the first data row's value */
            tagSampleLabels.push(addSampleValueLabel(tagRowGroup, sampleValueText(sourceIndex), "tooltip.sampleValue"));

            /*
               行番号はコントロール自身に持たせる（ループ変数を参照すると全行が最後の値を見てしまう）。
               Store the row index on the control: a closure over the loop variable would share the last value.
            */
            tagSourceDropdown.tagIndex = i;
            tagSourceDropdown.onChange = function () {
                if (!this.selection) return;
                tagSourceIndexes[this.tagIndex] = this.selection.index - 1;
                tagSampleLabels[this.tagIndex].text = sampleValueText(tagSourceIndexes[this.tagIndex]);
                refreshUsedColumnItems();
                dropPreviewForMappingChange();
            };
        }

        refreshUsedColumnItems();
    }

    /**
     * 作り直した一覧から、以前と同じ列を探す
     * 計算列が無くなったときは対応なしに戻す（税抜の変数へ税込の値が黙って入るのを避けるため）
     * @param {object} previousSource - 以前に対応づけていた列
     * @returns {number} 列の番号（見つからなければ-1）
     */
    function findSameSourceIndex(previousSource) {
        if (!previousSource) return -1;
        for (var i = 0; i < columnSources.length; i++) {
            if (columnSources[i].columnIndex === previousSource.columnIndex && columnSources[i].valueKind === previousSource.valueKind) return i;
        }
        return -1;
    }

    /**
     * 「消費税を自動計算」の切り替えを反映する
     * 計算列が増減して番号がずれるので、対応づけは列そのもので引き継ぐ
     * @returns {void}
     */
    function applyTaxCalculationToggle() {
        taxCalcEnabled = taxCalcCheckbox.value;

        var previousSources = [];
        for (var i = 0; i < tagSourceIndexes.length; i++) {
            var sourceIndex = tagSourceIndexes[i];
            previousSources.push((sourceIndex >= 0 && sourceIndex < columnSources.length) ? columnSources[sourceIndex] : null);
        }

        rebuildColumnSources();
        tagSourceIndexes = [];
        for (var j = 0; j < previousSources.length; j++) {
            tagSourceIndexes.push(findSameSourceIndex(previousSources[j]));
        }

        rebuildTagMappingRows();
        mainDialog.layout.layout(true);
        dropPreviewForMappingChange();
    }

    /**
     * 自動認識か手動照合かを反映する
     * 自動認識では名前が一致する列に戻し、対応パネルは隠す
     * @returns {void}
     */
    function applyMappingMode() {
        if (autoMatchRadio.value) applyAutoTagMapping();
        if (manualMappingRadio.value) {
            buildTagMappingPanel();
            rebuildTagMappingRows();
        } else {
            removeTagMappingPanel();
        }
        mainDialog.layout.layout(true);
        dropPreviewForMappingChange();
    }

    autoMatchRadio.onClick = applyMappingMode;
    manualMappingRadio.onClick = applyMappingMode;
    applyMappingMode();

    /**
     * 対応づけ済みの変数から、置換に使う組を作る
     * @returns {Array<{tag: string, source: object}>} タグ文字列と対応列の組
     */
    function buildActiveTagMappings() {
        var tagMappings = [];
        for (var i = 0; i < templateTagNames.length; i++) {
            var sourceIndex = tagSourceIndexes[i];
            if (sourceIndex == null || sourceIndex < 0 || sourceIndex >= columnSources.length) continue;
            tagMappings.push({ tag: "<" + templateTagNames[i] + ">", source: columnSources[sourceIndex] });
        }
        return tagMappings;
    }

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
     * @param {number[]} gridOriginRect - グリッド左上の基準矩形
     * @param {{columnCount: number, artboardGap: number, nameColumnIndex: number}} importLayout - 配置情報
     * @param {Array<{tag: string, source: object}>} tagMappings - 変数と列の対応
     * @param {number} dataIndex - データ行の番号（0始まり）
     * @returns {void}
     */
    function buildVariation(targetDocument, templateItems, gridOriginRect, importLayout, tagMappings, dataIndex) {
        var cellWidth = gridOriginRect[2] - gridOriginRect[0];
        var cellHeight = gridOriginRect[1] - gridOriginRect[3];
        var offsetX = (cellWidth + importLayout.artboardGap) * (dataIndex % importLayout.columnCount);
        var offsetY = -(cellHeight + importLayout.artboardGap) * Math.floor(dataIndex / importLayout.columnCount);

        var variationArtboard = targetDocument.artboards.add([
            gridOriginRect[0] + offsetX, gridOriginRect[1] + offsetY,
            gridOriginRect[2] + offsetX, gridOriginRect[3] + offsetY
        ]);
        variationArtboard.name = resolveArtboardName(dataIndex, importLayout.nameColumnIndex);

        /* 雛形オブジェクトを複製し、対応するアートボードへ移動して流し込み / Duplicate and merge */
        for (var i = 0; i < templateItems.length; i++) {
            var duplicatedItem = templateItems[i].duplicate();
            duplicatedItem.translate(offsetX, offsetY);
            replaceTagsRecursive(duplicatedItem, tagMappings, dataRows[dataIndex]);
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
        var tagMappings = buildActiveTagMappings();
        var templateItems = collectTemplateItems(targetDocument);
        if (templateItems.length === 0) {
            alert(getLabel("alert.noTemplate"));
            return false;
        }
        var gridOriginRect = moveTemplateToGridOrigin(templateArtboard, templateItems, importLayout);

        var progressWindow = new Window("palette", getLabel("progress.title"), undefined);
        var progressBar = progressWindow.add("progressbar", undefined, 0, dataRows.length);
        progressBar.preferredSize.width = PROGRESS_BAR_WIDTH;
        progressWindow.show();

        try {
            /* 重要：2件目(i=1)から先に処理する（雛形のタグを維持するため）/ Start at row 1 to keep the template tags */
            for (var i = 1; i < dataRows.length; i++) {
                progressBar.value = i;
                progressWindow.update();
                buildVariation(targetDocument, templateItems, gridOriginRect, importLayout, tagMappings, i);
            }

            /* 最後に1件目を雛形自身へ書き込む / Finally, merge row 0 into the template itself */
            progressBar.value = dataRows.length;
            progressWindow.update();
            templateArtboard.name = resolveArtboardName(0, importLayout.nameColumnIndex);
            for (var j = 0; j < templateItems.length; j++) {
                replaceTagsRecursive(templateItems[j], tagMappings, dataRows[0]);
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

        var outputFile = buildDuplicateFilePath(originalDocument.fullName, fileSuffixInput.text, fileDateCheckbox.value, fileTimeCheckbox.value);
        if (!confirmOutputFile(outputFile)) return;

        /* プレビュー用ドキュメント・ファイルが残っていれば破棄してから実行 / Discard any leftover preview first */
        discardPreview();

        var importDocument;
        try {
            importDocument = duplicateDocumentBySaveAs(originalDocument, outputFile);
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
            /* プレビュー用は必ず日時付きにして、保存名とぶつからないようにする / always timestamped so it never collides */
            if (!previewFile) previewFile = buildDuplicateFilePath(originalFile, "preview", true, true);
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
     * ドロップダウンの表示を、onChangeを発火させずに選び直す
     * @param {object} dropdown - 対象のドロップダウン
     * @param {number} itemIndex - 選び直す項目の番号
     * @returns {void}
     */
    function setDropdownSelectionSilently(dropdown, itemIndex) {
        if (itemIndex == null || itemIndex < 0 || itemIndex >= dropdown.items.length) return;
        var savedOnChange = dropdown.onChange;
        dropdown.onChange = null;
        if (!dropdown.selection || dropdown.selection.index !== itemIndex) {
            dropdown.selection = itemIndex;
        }
        dropdown.onChange = savedOnChange;
    }

    /**
     * ドロップダウンの表示を、控えてある番号から選び直す
     * ドキュメントの開閉でダイアログが再描画されると、表示だけが空欄になるため。
     * The dialog redraw that follows opening/closing a document blanks the shown item.
     * @returns {void}
     */
    function restoreDropdownSelections() {
        setDropdownSelectionSilently(artboardNameColumnDropdown, artboardNameColumnIndex);
        for (var i = 0; i < tagSourceDropdowns.length; i++) {
            /* 先頭が「対応なし」なので、項目番号は列番号より1つ大きい / index 0 is the "none" item */
            var sourceIndex = tagSourceIndexes[i];
            var itemIndex = (sourceIndex != null && sourceIndex >= 0 && sourceIndex < columnSources.length) ? (sourceIndex + 1) : 0;
            setDropdownSelectionSilently(tagSourceDropdowns[i], itemIndex);
        }
    }

    /**
     * プレビュー表示中のときだけ、プレビューを作り直す
     * @returns {void}
     */
    function refreshPreviewIfActive() {
        if (previewCheckbox.value) rebuildPreview();
        restoreDropdownSelections();
    }

    /**
     * 対応列を変更したときは、プレビューを作り直さずに外す
     * 作り直すとドキュメントの開閉でダイアログが再描画され、選んだ項目が空欄のままになるため。
     * Rebuilding reopens a document, and the dialog repaint that follows blanks the item just picked.
     * @returns {void}
     */
    function dropPreviewForMappingChange() {
        if (!previewCheckbox.value) return;
        previewCheckbox.value = false;
        discardPreview();
        restoreDropdownSelections();
    }

    /* 「プレビュー」：オンで表示、オフでプレビュー用ドキュメントを閉じる / Preview checkbox */
    previewCheckbox.onClick = function () {
        if (!previewCheckbox.value) {
            discardPreview();
        } else if (!rebuildPreview()) {
            previewCheckbox.value = false;
        }
        restoreDropdownSelections();
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
        refreshArtboardNameSample();
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
     * @param {Array<{tag: string, source: object}>} tagMappings - 変数と列の対応
     * @param {string[]} rowValues - 1行分のデータ値
     * @returns {void}
     */
    function replaceTagsRecursive(pageItem, tagMappings, rowValues) {
        if (!pageItem) return;

        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                replaceTagsRecursive(pageItem.pageItems[i], tagMappings, rowValues);
            }
            return;
        }
        if (pageItem.typename !== "TextFrame") return;

        var originalContents = pageItem.contents;
        if (originalContents == null) return;

        var tagReplacements = collectTagReplacements(String(originalContents), tagMappings, rowValues);
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
     * @param {Array<{tag: string, source: object}>} tagMappings - 変数と列の対応
     * @param {string[]} rowValues - 1行分のデータ値
     * @returns {Array<{tag: string, value: string}>} 置換の組（含まれていないタグは返さない）
     */
    function collectTagReplacements(textContents, tagMappings, rowValues) {
        var tagReplacements = [];
        for (var i = 0; i < tagMappings.length; i++) {
            if (textContents.indexOf(tagMappings[i].tag) === -1) continue;
            tagReplacements.push({ tag: tagMappings[i].tag, value: resolveSourceValue(tagMappings[i].source, rowValues) });
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
