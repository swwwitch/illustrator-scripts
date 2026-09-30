#target illustrator
#targetengine "ArtboardExporterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログで選んだアートボードを PNG・JPEG・PDF で書き出します。
アートボードごとに形式・倍率（複数可）・背景（白・黒・透明）を指定でき、ファイル名・保存先・除外も設定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardExporter.md

### Overview

Exports the artboards chosen in a dialog as PNG, JPEG or PDF.
Each artboard can have its own format, scales (one or more) and background (white, black or transparent); file names, destination and exclusions are configurable too.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardExporter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArtboardExporter";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardExporter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardExporter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var DEFAULT_FORMAT      = "png";   /* 形式の初期値（"png" / "jpeg" / "pdf"）。2回目からは前回選んだ形式 / Initial format; later runs use the last one chosen */
    var DEFAULT_SCALES      = [100];   /* 倍率の初期値（%）/ Initial scales (%) */
    var DEFAULT_BACKGROUND  = "white"; /* 背景の初期値（"transparent" / "white" / "black"）/ Initial background */
    var JPEG_QUALITY        = 90;      /* JPEG の品質の初期値（0〜100）/ Initial JPEG quality (0-100) */
    var JPEG_MAX_SCALE      = 776;     /* JPEG で書き出せる倍率の上限（%）/ Largest scale JPEG export accepts (%) */
    /* PDF プリセットの初期値（上から順に、あるものを使う。どれも無ければ一覧の先頭）/ Initial PDF preset: the first one found, else the first in the list */
    var DEFAULT_PDF_PRESETS = ["[Illustrator Default]", "[Illustrator 初期設定]"];

    var DEFAULT_REVEAL_FOLDER           = true;    /* ［書き出し後にフォルダーを表示］の初期値 / Initial "Show Folder After Export" */
    var DEFAULT_SUBFOLDER_NAME          = "{doc}"; /* ［フォルダーを作成］の名前の初期値 / Initial name for "Create Folder" */
    var DEFAULT_EXCLUDE_ARTBOARD_PREFIX = "#";     /* 書き出さないアートボード名の頭文字 / Prefix of artboard names to skip */
    var DEFAULT_EXCLUDE_LAYER_PREFIX    = "//";    /* 書き出し中に隠すレイヤー名の頭文字 / Prefix of layer names to hide while exporting */

    var DEFAULT_USE_DOCUMENT_NAME        = true;  /* ファイル名にドキュメント名を入れるか / Include the document name */
    var DEFAULT_ARTBOARD_PATTERN         = 4;     /* アートボード部分の初期値（ARTBOARD_NAME_PATTERNS の番号。4 は「-名称」）/ Initial artboard part (index; 4 = "-Name") */
    var DEFAULT_USE_DATE                 = false; /* ファイル名に日付を入れるか / Include the date */
    var DEFAULT_DATE_FORMAT              = 0;     /* 日付の書式（DATE_FORMATS の番号）/ Date format (index into DATE_FORMATS) */
    var DEFAULT_SUFFIX_SEPARATOR         = "_";   /* 倍率・日付・文字列の前の区切り（"_" / "-"）/ Separator before the scale, date and text */
    var DEFAULT_USE_TEMPLATE             = false; /* 初めは［書式］ではなく［組み合わせ］/ Start with Combine rather than Pattern */
    var DEFAULT_FILE_NAME_TEMPLATE       = "{doc}-{name}"; /* ［書式］の初期値 / Initial pattern */
    var DEFAULT_INVALID_CHAR_REPLACEMENT = "_";   /* ファイル名に使えない文字の置き換え先 / Replacement for characters file names cannot hold */
    var DEFAULT_REPLACE_SPACES           = false; /* スペースも置き換えるか / Also replace spaces */

    // =========================================
    // 書き出し形式 / Export formats
    // =========================================

    /* 形式ごとの表示名と拡張子 / Display label and extension per format */
    var EXPORT_FORMATS = [
        { key: "png",  label: "PNG",  extension: "png" },
        { key: "jpeg", label: "JPEG", extension: "jpg" },
        { key: "pdf",  label: "PDF",  extension: "pdf" }
    ];

    /**
     * 形式のキーから形式の情報を返す
     * @param {string} formatKey - "png" / "jpeg" / "pdf"
     * @returns {{key: string, label: string, extension: string}} 形式の情報（未知のキーは PNG）
     */
    function getFormatInfo(formatKey) {
        for (var i = 0; i < EXPORT_FORMATS.length; i++) {
            if (EXPORT_FORMATS[i].key === formatKey) return EXPORT_FORMATS[i];
        }
        return EXPORT_FORMATS[0];
    }

    /**
     * Illustrator にある PDF プリセットの名前を返す
     * @returns {string[]} プリセット名（取れなければ空）
     */
    function getPdfPresetNames() {
        var presetNames = [];
        try {
            var presetList = app.PDFPresetsList;
            for (var i = 0; i < presetList.length; i++) presetNames.push(String(presetList[i]));
        } catch (e) {}
        return presetNames;
    }

    /**
     * 使う PDF プリセットを決める（希望の名前 → DEFAULT_PDF_PRESETS → 一覧の先頭の順）
     * @param {string[]} presetNames - getPdfPresetNames() の結果
     * @param {string} [preferredName] - 希望のプリセット名
     * @returns {string} プリセット名（プリセットが1つも無ければ空文字）
     */
    function resolvePdfPreset(presetNames, preferredName) {
        var candidates = [preferredName].concat(DEFAULT_PDF_PRESETS);
        for (var i = 0; i < candidates.length; i++) {
            for (var j = 0; j < presetNames.length; j++) {
                if (candidates[i] && presetNames[j] === candidates[i]) return presetNames[j];
            }
        }
        return presetNames.length > 0 ? presetNames[0] : "";
    }

    // =========================================
    // ファイル名 / File names
    // =========================================

    /* アートボード部分の書式。{n} は番号、{a} はアートボード名 / Artboard part: {n} = number, {a} = artboard name */
    var ARTBOARD_NAME_PATTERNS = [
        "{n}", "-{n}", "_{n}",
        "{a}", "-{a}", "_{a}",
        "{n}{a}", "{n}-{a}", "{n}_{a}",
        "-{n}{a}", "_{n}{a}", "-{n}-{a}", "_{n}_{a}",
        ""
    ];

    /* 日付の書式 / Date formats */
    var DATE_FORMATS = ["YYYYMMDD", "YYMMDD", "YYYY-MM-DD", "MMDD"];

    /* 区切り・置き換え先に選べる文字 / Characters offered as separators and replacements */
    var SEPARATOR_CHOICES = ["_", "-"];

    /* ファイル名に使えない文字 / Characters file names cannot hold */
    var INVALID_FILE_NAME_CHARS = /[\/\\:\*\?"<>\|¥]/g;

    /**
     * 今日の日付を指定の書式で返す
     * @param {number} dateFormatIndex - DATE_FORMATS の番号
     * @returns {string} 日付
     */
    function formatDateStamp(dateFormatIndex) {
        var today = new Date();
        var yearText = String(today.getFullYear());
        var monthText = (today.getMonth() < 9 ? "0" : "") + (today.getMonth() + 1);
        var dayText = (today.getDate() < 10 ? "0" : "") + today.getDate();
        return (DATE_FORMATS[dateFormatIndex] || DATE_FORMATS[0])
            .replace("YYYY", yearText)
            .replace("YY", yearText.substring(2))
            .replace("MM", monthText)
            .replace("DD", dayText);
    }

    /**
     * アートボード番号を、アートボードの数の桁数に 0 でそろえる（10枚以上なら 01, 02 …）
     * @param {number} artboardNumber - 1 から始まる番号
     * @param {number} artboardCount - アートボードの数
     * @returns {string} そろえた番号
     */
    function padArtboardNumber(artboardNumber, artboardCount) {
        var numberText = String(artboardNumber);
        while (numberText.length < String(artboardCount).length) numberText = "0" + numberText;
        return numberText;
    }

    /**
     * ファイル名に使えない文字（と、指定があればスペース）を置き換える
     * @param {string} rawName - 元の名前
     * @param {{invalidCharReplacement: string, replaceSpaces: boolean, spaceReplacement: string}} fileNameOptions - ファイル名の設定
     * @returns {string} 置き換えた名前
     */
    function cleanFileName(rawName, fileNameOptions) {
        var cleanedName = String(rawName).replace(INVALID_FILE_NAME_CHARS, fileNameOptions.invalidCharReplacement || DEFAULT_INVALID_CHAR_REPLACEMENT);
        /* 半角・全角のスペース / Half- and full-width spaces */
        if (fileNameOptions.replaceSpaces) cleanedName = cleanedName.replace(/[ 　]/g, fileNameOptions.spaceReplacement || DEFAULT_INVALID_CHAR_REPLACEMENT);
        return cleanedName;
    }

    /**
     * ［書式］の記号を置き換える
     * @param {string} fileNameTemplate - 書式（例 "{doc}_{name}_{date}"）
     * @param {{baseFileName: string, artboardName: string, artboardNumber: string, scale: number, dateStamp: string}} nameContext - 置き換える値
     * @returns {string} 置き換えた文字列
     */
    function fillFileNameTemplate(fileNameTemplate, nameContext) {
        return String(fileNameTemplate)
            .replace(/\{doc\}/g, nameContext.baseFileName)
            .replace(/\{name\}/g, nameContext.artboardName)
            .replace(/\{num\}/g, nameContext.artboardNumber)
            .replace(/\{scale\}/g, String(nameContext.scale))
            .replace(/\{date\}/g, nameContext.dateStamp);
    }

    /**
     * 書き出すファイル名（拡張子なし）を組み立てる。
     * ［書式］は記号を置き換え、{scale} が無く倍率が 100% 以外なら末尾に「-倍率」を付ける。
     * ［組み合わせ］は [ドキュメント名][アートボード部分][区切り倍率][区切り日付][区切り文字列] の順
     * @param {Object} fileNameOptions - ファイル名の設定（loadOutputSettings() の形）と documentName
     * @param {{baseFileName: string, artboardState: Object, artboardCount: number, scale: number, dateStamp: string}} stemContext - 名前の材料
     * @returns {string} ファイル名（空になることもある）
     */
    function buildFileStem(fileNameOptions, stemContext) {
        var artboardState = stemContext.artboardState;
        var nameContext = {
            baseFileName: getDocumentNamePart(fileNameOptions, stemContext.baseFileName, true),
            artboardName: artboardState.name,
            artboardNumber: padArtboardNumber(artboardState.index + 1, stemContext.artboardCount),
            scale: stemContext.scale,
            dateStamp: stemContext.dateStamp
        };
        var fileStem;
        if (fileNameOptions.useFileNameTemplate) {
            fileStem = fillFileNameTemplate(fileNameOptions.fileNameTemplate, nameContext);
            /* 倍率ごとのファイルが重ならないよう、{scale} が無ければ付ける / Keep per-scale files apart when {scale} is missing */
            if (String(fileNameOptions.fileNameTemplate).indexOf("{scale}") < 0 && stemContext.scale !== 100) fileStem += (fileStem === "" ? "" : "-") + stemContext.scale;
            return cleanFileName(fileStem, fileNameOptions);
        }
        var suffixSeparator = fileNameOptions.suffixSeparator || DEFAULT_SUFFIX_SEPARATOR;
        fileStem = getDocumentNamePart(fileNameOptions, stemContext.baseFileName, false);
        var artboardPart = (ARTBOARD_NAME_PATTERNS[fileNameOptions.artboardPatternIndex] || "")
            .replace("{n}", nameContext.artboardNumber)
            .replace("{a}", nameContext.artboardName);
        /* 前に何も無いときは先頭の区切りを外す / Drop the leading separator when nothing precedes it */
        if (fileStem === "") artboardPart = artboardPart.replace(/^[-_]/, "");
        fileStem += artboardPart;
        /* 100% 以外は倍率を付けて見分ける / Scales other than 100% get a suffix */
        if (stemContext.scale !== 100) fileStem += (fileStem === "" ? "" : suffixSeparator) + stemContext.scale;
        fileStem = appendDateAndText(fileStem, fileNameOptions, stemContext.dateStamp);
        return cleanFileName(fileStem, fileNameOptions);
    }

    /**
     * PDF を1つにまとめるときのファイル名（拡張子なし）を作る。アートボード部分と倍率は入れない
     * @param {Object} fileNameOptions - ファイル名の設定
     * @param {string} baseFileName - 拡張子を除いたドキュメント名
     * @param {string} dateStamp - 日付
     * @returns {string} ファイル名（空ならドキュメント名）
     */
    function buildCombinedPdfStem(fileNameOptions, baseFileName, dateStamp) {
        var fileStem;
        if (fileNameOptions.useFileNameTemplate) {
            fileStem = fillFileNameTemplate(fileNameOptions.fileNameTemplate, {
                baseFileName: getDocumentNamePart(fileNameOptions, baseFileName, true),
                artboardName: "",
                artboardNumber: "",
                scale: "",
                dateStamp: dateStamp
            });
            /* 空になった記号のまわりの区切りを詰める / Collapse separators left around emptied tokens */
            fileStem = fileStem.replace(/([-_])[-_]+/g, "$1").replace(/^[-_]+|[-_]+$/g, "");
        } else {
            fileStem = appendDateAndText(getDocumentNamePart(fileNameOptions, baseFileName, false), fileNameOptions, dateStamp);
        }
        return cleanFileName(fileStem === "" ? baseFileName : fileStem, fileNameOptions);
    }

    /**
     * ファイル名に使うドキュメント名を返す（［ファイル名］欄を書き換えていればその文字列）
     * @param {Object} fileNameOptions - ファイル名の設定
     * @param {string} baseFileName - 拡張子を除いたドキュメント名
     * @param {boolean} forTemplate - ［書式］の {doc} 用なら true（チェックに関係なく返す）
     * @returns {string} ドキュメント名（入れないときは空文字）
     */
    function getDocumentNamePart(fileNameOptions, baseFileName, forTemplate) {
        if (!forTemplate && !fileNameOptions.useDocumentName) return "";
        var documentName = fileNameOptions.documentName;
        return (documentName === undefined || documentName === "") ? baseFileName : documentName;
    }

    /**
     * ［組み合わせ］の末尾に日付と文字列を足す
     * @param {string} fileStem - ここまでのファイル名
     * @param {Object} fileNameOptions - ファイル名の設定
     * @param {string} dateStamp - 日付
     * @returns {string} 足したファイル名
     */
    function appendDateAndText(fileStem, fileNameOptions, dateStamp) {
        var suffixSeparator = fileNameOptions.suffixSeparator || DEFAULT_SUFFIX_SEPARATOR;
        if (fileNameOptions.useDate) fileStem += (fileStem === "" ? "" : suffixSeparator) + dateStamp;
        if (fileNameOptions.useSuffixText && fileNameOptions.suffixText) fileStem += (fileStem === "" ? "" : suffixSeparator) + fileNameOptions.suffixText;
        return fileStem;
    }

    /**
     * 書き出しジョブのうち、ファイル名が重なるものを返す
     * @param {Object[]} exportJobs - 書き出しジョブ
     * @returns {string[]} 重なったファイル名
     */
    function findDuplicateFileNames(exportJobs) {
        var seenNames = {};
        var duplicateNames = [];
        for (var i = 0; i < exportJobs.length; i++) {
            var nameKey = "f:" + exportJobs[i].file.fsName.toLowerCase(); /* macOS は大文字・小文字を区別しない / macOS is case-insensitive */
            if (seenNames[nameKey] === 1) duplicateNames.push(decodeURI(exportJobs[i].file.name));
            seenNames[nameKey] = (seenNames[nameKey] || 0) + 1;
        }
        return duplicateNames;
    }

    /**
     * ［フォルダーを作成］の名前を作る（{doc} と {date} を置き換える）
     * @param {string} subfolderTemplate - フォルダー名の書式
     * @param {string} baseFileName - 拡張子を除いたドキュメント名
     * @param {string} dateStamp - 日付
     * @param {Object} fileNameOptions - 置き換えの設定
     * @returns {string} フォルダー名（空なら作らない）
     */
    function resolveSubfolderName(subfolderTemplate, baseFileName, dateStamp, fileNameOptions) {
        var folderName = String(subfolderTemplate || "").replace(/\{doc\}/g, baseFileName).replace(/\{date\}/g, dateStamp);
        return cleanFileName(folderName, fileNameOptions).replace(/^\s+|\s+$/g, "");
    }

    // =========================================
    // 書き出し対象 / Export targets
    // =========================================

    /**
     * 「1-3, 5」のような番号の指定を読む（全角・「〜」「~」も可）
     * @param {string} rangeText - 番号の指定
     * @param {number} artboardCount - アートボードの数
     * @returns {number[]|null} 0 から始まる番号の配列。読めないときや範囲外があれば null
     */
    function parseArtboardRange(rangeText, artboardCount) {
        var normalizedText = String(rangeText).replace(/[０-９]/g, function (fullWidthChar) {
            return String.fromCharCode(fullWidthChar.charCodeAt(0) - 0xFEE0);
        }).replace(/[〜~－ー―]/g, "-").replace(/[、，]/g, ",");
        var rangeTokens = normalizedText.split(/[,\s]+/);
        var pickedIndexes = [];
        var pickedFlags = {};
        for (var i = 0; i < rangeTokens.length; i++) {
            if (rangeTokens[i] === "") continue;
            var rangeMatch = rangeTokens[i].match(/^(\d+)(?:-(\d+))?$/);
            if (!rangeMatch) return null;
            var firstNumber = Number(rangeMatch[1]);
            var lastNumber = rangeMatch[2] ? Number(rangeMatch[2]) : firstNumber;
            if (firstNumber > lastNumber) {
                var swapNumber = firstNumber;
                firstNumber = lastNumber;
                lastNumber = swapNumber;
            }
            if (firstNumber < 1 || lastNumber > artboardCount) return null;
            for (var n = firstNumber; n <= lastNumber; n++) {
                if (pickedFlags["n" + n]) continue;
                pickedFlags["n" + n] = true;
                pickedIndexes.push(n - 1);
            }
        }
        return (pickedIndexes.length > 0) ? pickedIndexes : null;
    }

    /**
     * アートボードが除外の対象かを返す（名前の前方一致）
     * @param {string} artboardName - アートボード名
     * @param {{excludeArtboards: boolean, excludeArtboardPrefix: string}} exclusionOptions - 除外の設定
     * @returns {boolean} 除外するなら true
     */
    function isArtboardExcluded(artboardName, exclusionOptions) {
        var namePrefix = exclusionOptions.excludeArtboardPrefix;
        return !!(exclusionOptions.excludeArtboards && namePrefix && String(artboardName).indexOf(namePrefix) === 0);
    }

    /**
     * 名前が指定の文字で始まる表示中のレイヤーを、サブレイヤーまでたどって隠す
     * @param {Document|Layer} layerParent - ドキュメントまたはレイヤー
     * @param {string} namePrefix - レイヤー名の頭文字
     * @param {Layer[]} [hiddenLayers] - 隠したレイヤーを足していく配列
     * @returns {Layer[]} 隠したレイヤー（元に戻すときに使う）
     */
    function hideExcludedLayers(layerParent, namePrefix, hiddenLayers) {
        hiddenLayers = hiddenLayers || [];
        for (var i = 0; i < layerParent.layers.length; i++) {
            var targetLayer = layerParent.layers[i];
            if (targetLayer.name.indexOf(namePrefix) === 0) {
                if (targetLayer.visible) {
                    targetLayer.visible = false;
                    hiddenLayers.push(targetLayer);
                }
                continue; /* 隠したレイヤーの中は見なくてよい / No need to look inside a hidden layer */
            }
            hideExcludedLayers(targetLayer, namePrefix, hiddenLayers);
        }
        return hiddenLayers;
    }

    /**
     * hideExcludedLayers() で隠したレイヤーを表示に戻す
     * @param {Layer[]} hiddenLayers - 隠したレイヤー
     * @returns {void}
     */
    function restoreHiddenLayers(hiddenLayers) {
        for (var i = 0; i < hiddenLayers.length; i++) {
            try {
                hiddenLayers[i].visible = true;
            } catch (e) {}
        }
    }

    // =========================================
    // パスの表示 / Path display
    // =========================================

    /**
     * ホームフォルダー以下のパスを ~ 表記に短縮する（/Users/takano/Desktop → ~/Desktop）
     * @param {string} fsPath - 短縮するパス
     * @returns {string} 短縮したパス。ホーム以下でなければそのまま
     */
    function abbreviateHomePath(fsPath) {
        var homePath = "";
        try {
            homePath = Folder("~").fsName;
        } catch (e) {}
        if (!homePath) return fsPath;
        if (fsPath === homePath) return "~";
        if (fsPath.indexOf(homePath) === 0) {
            var pathAfterHome = fsPath.substring(homePath.length);
            /* 直後が区切り文字のときだけ短縮（/Users/tak が /Users/takano に一致しないように）/ Abbreviate only on a separator boundary */
            if (pathAfterHome.charAt(0) === "/" || pathAfterHome.charAt(0) === "\\") return "~" + pathAfterHome;
        }
        return fsPath;
    }

    // =========================================
    // セッション中の記憶 / Session memory
    // =========================================

    /* 常駐エンジンの $.global に置き、Illustrator を終了するまで残す / Kept on $.global of the persistent engine until Illustrator quits */
    var FORMAT_MEMORY_KEY = "__" + SCRIPT_NAME + "_Format";            /* 最後に選んだ形式 / Last format chosen */
    var PDF_PRESET_MEMORY_KEY = "__" + SCRIPT_NAME + "_PdfPreset";     /* 最後に選んだ PDF プリセット / Last PDF preset chosen */
    var OUTPUT_MEMORY_KEY = "__" + SCRIPT_NAME + "_Output";            /* 保存先・ファイル名・除外などの設定 / Destination, file name, exclusion and other settings */
    var ARTBOARD_MEMORY_KEY = "__" + SCRIPT_NAME + "_ArtboardSettings"; /* ドキュメント・アートボードごとの設定 / Settings per document and artboard */

    /**
     * ドキュメントのアートボードごとの設定を記憶から取り出す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} アートボード名をキーにした設定（無ければ空のオブジェクト）
     */
    function loadRememberedArtboardSettings(doc) {
        var memory = $.global[ARTBOARD_MEMORY_KEY];
        var documentKey = getDocumentMemoryKey(doc);
        if (!memory || !memory[documentKey]) return {};
        return memory[documentKey];
    }

    /**
     * アートボードごとの設定を記憶する（同じ名前のアートボードは上書き）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} artboardStates - アートボードの状態
     * @returns {void}
     */
    function rememberArtboardSettings(doc, artboardStates) {
        if (!$.global[ARTBOARD_MEMORY_KEY]) $.global[ARTBOARD_MEMORY_KEY] = {};
        var documentKey = getDocumentMemoryKey(doc);
        var documentSettings = $.global[ARTBOARD_MEMORY_KEY][documentKey] || {};
        for (var i = 0; i < artboardStates.length; i++) {
            var artboardState = artboardStates[i];
            /* 名前をキーにするので、並べ替えや追加があっても同じアートボードに戻る / Keyed by name so the settings follow the artboard after reordering */
            documentSettings["ab:" + artboardState.name] = {
                enabled: artboardState.enabled,
                formatKey: artboardState.formatKey,
                pdfPreset: artboardState.pdfPreset,
                jpegQuality: artboardState.jpegQuality,
                scales: artboardState.scales.slice(0),
                backgroundKey: artboardState.backgroundKey
            };
        }
        $.global[ARTBOARD_MEMORY_KEY][documentKey] = documentSettings;
    }

    /**
     * 記憶のキーにするドキュメントの識別子を返す（保存先のパス）
     * @param {Document} doc - 対象ドキュメント
     * @returns {string} 識別子
     */
    function getDocumentMemoryKey(doc) {
        var sourceFolder = getSourceFolder(doc);
        /* 未保存のドキュメントは名前で見分ける / Unsaved documents are told apart by name */
        return sourceFolder ? "doc:" + doc.fullName.fsName : "unsaved:" + doc.name;
    }

    /**
     * アートボードごと以外の設定を記憶から取り出す（無い項目は初期値）
     * @returns {Object} 保存先・ファイル名・除外・PDF のまとめなどの設定
     */
    function loadOutputSettings() {
        var remembered = $.global[OUTPUT_MEMORY_KEY];
        var outputSettings = {
            useCustomFolder: false,
            customFolderPath: "",
            createSubfolder: false,
            subfolderName: DEFAULT_SUBFOLDER_NAME,
            revealAfterExport: DEFAULT_REVEAL_FOLDER,
            excludeArtboards: false,
            excludeArtboardPrefix: DEFAULT_EXCLUDE_ARTBOARD_PREFIX,
            excludeLayers: false,
            excludeLayerPrefix: DEFAULT_EXCLUDE_LAYER_PREFIX,
            combinePdf: false,
            useFileNameTemplate: DEFAULT_USE_TEMPLATE,
            fileNameTemplate: DEFAULT_FILE_NAME_TEMPLATE,
            useDocumentName: DEFAULT_USE_DOCUMENT_NAME,
            artboardPatternIndex: DEFAULT_ARTBOARD_PATTERN,
            suffixSeparator: DEFAULT_SUFFIX_SEPARATOR,
            useDate: DEFAULT_USE_DATE,
            dateFormatIndex: DEFAULT_DATE_FORMAT,
            useSuffixText: false,
            suffixText: "",
            invalidCharReplacement: DEFAULT_INVALID_CHAR_REPLACEMENT,
            replaceSpaces: DEFAULT_REPLACE_SPACES,
            spaceReplacement: DEFAULT_INVALID_CHAR_REPLACEMENT
        };
        if (remembered) {
            for (var settingKey in outputSettings) {
                if (remembered[settingKey] !== undefined) outputSettings[settingKey] = remembered[settingKey];
            }
        }
        return outputSettings;
    }

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PROGRESS_WIDTH      = 360;              /* 状況表示とバーの幅 / Width of the status text and bar */
    var PROGRESS_BAR_HEIGHT = 14;               /* バーの高さ / Bar height */

    var ARTBOARD_LIST_WIDTH      = 440;                         /* アートボード一覧の幅 / Artboard list width */
    var ARTBOARD_LIST_HEIGHT     = 250;                         /* アートボード一覧の高さ / Artboard list height */
    var ARTBOARD_LIST_COLUMNS    = [32, 56, 150, 50, 90, 60];   /* 一覧の列幅 [番号,書き出し,名前,形式,倍率,背景] / Column widths [no., export, name, format, scale, background] */
    var LIST_COLUMN_SPACING      = 8;                           /* 一覧まわりの間隔 / Spacing around the list */
    var TARGET_RANGE_CHARS       = 8;                           /* 番号指定欄の文字数 / Range field width in characters */
    var PREFIX_FIELD_CHARS       = 5;                           /* 除外の頭文字欄の文字数 / Prefix field width in characters */
    var JPEG_QUALITY_FIELD_CHARS = 4;                           /* 品質欄の文字数 / Quality field width in characters */
    var SCALE_FIELD_CHARS        = 8;                           /* 倍率欄の文字数（「100, 200」が収まる幅）/ Scale field width (fits "100, 200") */
    var PDF_PRESET_WIDTH         = 200;                         /* PDF プリセットの幅 / PDF preset dropdown width */
    var SWATCH_SIZE              = 16;                          /* 背景の色見本の大きさ / Background swatch size */
    var SWATCH_CHECKER_SIZE      = 4;                           /* 透明の市松模様の1マス / Checker cell size for transparent */
    var DOCUMENT_NAME_CHARS      = 12;                          /* ［ファイル名］欄の文字数 / Document name field width */
    var SUFFIX_TEXT_CHARS        = 8;                           /* 文字列欄の文字数 / Text field width */
    var FILE_NAME_TEMPLATE_CHARS = 20;                          /* ［書式］欄の文字数 / Pattern field width in characters */
    var SUBFOLDER_NAME_CHARS     = 14;                          /* フォルダー名欄の文字数 / Folder name field width */
    var FOLDER_PATH_WIDTH        = 420;                         /* 保存先のパス表示の幅 / Width of the destination path */
    var OVERWRITE_LIST_MAX       = 10;                          /* 確認に並べるファイル名の上限 / Max file names listed in prompts */

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

    /**
     * 複数のラベルの幅を最長のものへそろえる
     * @param {StaticText[]} labelControls - 幅をそろえる statictext
     * @returns {void}
     */
    function alignLabelWidths(labelControls) {
        var maxLabelWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            if (labelControls[i].preferredSize.width > maxLabelWidth) maxLabelWidth = labelControls[i].preferredSize.width;
        }
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxLabelWidth;
            labelControls[j].justify = "right";
        }
    }

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title:         { ja: "アートボードを書き出し", en: "Export Artboards" },
            progressTitle: { ja: "書き出し中…", en: "Exporting…" }
        },
        panel: {
            selectedArtboards: { ja: "選択したアートボードの書き出し設定", en: "Export Settings for Selected Artboards" },
            destination:       { ja: "保存先", en: "Destination" },
            fileName:          { ja: "ファイル名", en: "File Name" }
        },
        fieldLabel: {
            format:         { ja: "形式", en: "Format" },
            scale:          { ja: "倍率", en: "Scale" },
            background:     { ja: "背景", en: "Background" },
            jpegQuality:    { ja: "品質", en: "Quality" },
            pdfPreset:      { ja: "PDF プリセット", en: "PDF Preset" },
            example:        { ja: "例", en: "Example" },
            tokens:         { ja: "使える記号", en: "Tokens" },
            targetRange:    { ja: "番号で指定", en: "By Number" },
            prefixMatch:    { ja: "前方一致", en: "prefix" },
            replaceChars:   { ja: "置き換え", en: "Replace" },
            jpegScaleLimit: { ja: "（JPEG は %1 まで）", en: "(JPEG up to %1)" }
        },
        column: {
            number:       { ja: "#", en: "#" },
            exportMark:   { ja: "書き出し", en: "Export" },
            name:         { ja: "アートボード名", en: "Artboard" },
            format:       { ja: "形式", en: "Format" },
            scale:        { ja: "倍率", en: "Scale" },
            background:   { ja: "背景", en: "Background" },
            excludedMark: { ja: "除外", en: "Excluded" }
        },
        checkbox: {
            exportTarget:     { ja: "書き出す", en: "Export" },
            revealFolder:     { ja: "書き出し後にフォルダーを表示", en: "Show Folder After Export" },
            documentName:     { ja: "ファイル名", en: "File Name" },
            date:             { ja: "日付", en: "Date" },
            suffixText:       { ja: "文字列", en: "Text" },
            replaceSpaces:    { ja: "スペース", en: "Spaces" },
            excludeArtboards: { ja: "除外アートボード名", en: "Exclude Artboards" },
            excludeLayers:    { ja: "除外レイヤー名", en: "Exclude Layers" },
            createSubfolder:  { ja: "フォルダーを作成", en: "Create Folder" },
            combinePdf:       { ja: "PDF を1つにまとめる", en: "Combine PDFs into One File" }
        },
        dropdown: {
            artboardNumber: { ja: "番号", en: "Number" },
            artboardName:   { ja: "名称", en: "Name" },
            noArtboardPart: { ja: "（なし）", en: "(None)" }
        },
        radio: {
            transparent:      { ja: "透明", en: "Transparent" },
            white:            { ja: "白", en: "White" },
            black:            { ja: "黒", en: "Black" },
            sameFolder:       { ja: "元ファイルと同じ", en: "Same as Source File" },
            customFolder:     { ja: "指定", en: "Custom" },
            combineParts:     { ja: "組み合わせ", en: "Combine" },
            fileNameTemplate: { ja: "書式", en: "Pattern" }
        },
        tooltip: {
            artboardList:     { ja: "ダブルクリックで書き出す・書き出さないを切り替え", en: "Double-click to switch export on or off" },
            scaleField:       { ja: "カンマで区切ると複数の倍率で書き出します（例：100, 200）", en: "Separate with commas to export at several scales (e.g. 100, 200)" },
            suffixSeparator:  { ja: "倍率・日付・文字列の前の区切り", en: "Separator before the scale, date and text" },
            dateFormat:       { ja: "日付の書式（［書式］の {date} にも使う）", en: "Date format (also used for {date} in Pattern)" },
            targetRange:      { ja: "例：1-3, 5（Return キーでも適用）", en: "e.g. 1-3, 5 (Return also applies)" },
            targetAll:        { ja: "除外以外のすべてのアートボードを書き出す", en: "Export every artboard that is not excluded" },
            targetActive:     { ja: "作業中のアートボードだけを書き出す", en: "Export only the active artboard" },
            excludeArtboards: { ja: "名前がこの文字で始まるアートボードは書き出さない", en: "Skip artboards whose names start with this" },
            excludeLayers:    { ja: "名前がこの文字で始まるレイヤーを隠して書き出す（書き出し後に元に戻す）", en: "Hide layers whose names start with this while exporting (shown again afterwards)" },
            replaceChars:     { ja: "ファイル名に使えない文字を置き換える文字", en: "Replacement for characters file names cannot hold" },
            combinePdf:       { ja: "書き出す PDF を1つのファイルにまとめる（プリセットは先頭のアートボードのもの）", en: "Put every PDF into one file (uses the first artboard's preset)" },
            subfolderName:    { ja: "{doc}：ドキュメント名\n{date}：日付", en: "{doc}: document name\n{date}: date" },
            fileNameTemplate: {
                ja: "{doc}：ドキュメント名\n{name}：アートボード名\n{num}：アートボード番号\n{scale}：倍率（書式に無いときは、100% 以外なら末尾に「-倍率」）\n{date}：日付",
                en: "{doc}: document name\n{name}: artboard name\n{num}: artboard number\n{scale}: scale (when missing, \"-scale\" is added for anything other than 100%)\n{date}: date"
            }
        },
        button: {
            selectAll:    { ja: "すべて選択", en: "Select All" },
            targetAll:    { ja: "すべて", en: "All" },
            targetActive: { ja: "作業中のアートボード", en: "Active Artboard" },
            applyRange:   { ja: "適用", en: "Apply" },
            chooseFolder: { ja: "選択…", en: "Choose…" },
            cancel:       { ja: "キャンセル", en: "Cancel" },
            ok:           { ja: "書き出し", en: "Export" }
        },
        status: {
            preparing:  { ja: "準備中…", en: "Preparing…" },
            cancelling: { ja: "キャンセル中…", en: "Cancelling…" },
            cancelled:  { ja: "キャンセルしました", en: "Cancelled" },
            done:       { ja: "完了", en: "Done" }
        },
        alert: {
            chooseFolder:   { ja: "保存先のフォルダーを選択", en: "Choose the destination folder" },
            folderMissing:  { ja: "保存先のフォルダーが見つかりません。", en: "The destination folder was not found." },
            folderNotCreated: { ja: "フォルダーを作成できませんでした：\n%1", en: "Could not create the folder:\n%1" },
            emptyFileName:  { ja: "ファイル名が空になります。いずれかを指定してください。", en: "The file name would be empty. Choose at least one part." },
            invalidRange:   { ja: "番号の指定を読み取れません。例：1-3, 5", en: "Could not read the numbers. Example: 1-3, 5" },
            noExportTarget: { ja: "書き出すアートボードがありません。", en: "There are no artboards to export." },
            duplicateFileNames: {
                ja: "次のファイル名が重なっています。ファイル名の設定を見直してください。\n\n%1",
                en: "These file names are used more than once. Review the file name settings.\n\n%1"
            },
            jpegScaleTooLarge: {
                ja: "JPEG の倍率は %1% までです。次のアートボードの倍率を見直してください。\n\n%2",
                en: "JPEG scales go up to %1%. Review the scale of these artboards.\n\n%2"
            },
            pdfNeedsSavedDocument: {
                ja: "PDF を書き出すには、先にドキュメントを保存してください。",
                en: "Save the document before exporting PDF."
            },
            invalidScale: { ja: "倍率には 0 より大きい数値を入力してください。", en: "Enter a scale greater than 0." },
            saveBeforePdf: {
                ja: "PDF は保存されている内容から書き出します。\nドキュメントを保存してから書き出しますか？",
                en: "PDF files are made from the saved document.\nSave the document and continue?"
            },
            overwrite: {
                ja: "次の %1 個のファイルはすでにあります。上書きしますか？\n\n%2",
                en: "%1 file(s) already exist. Overwrite them?\n\n%2"
            },
            exportError: {
                ja: "アートボード「%1」の書き出し中にエラーが発生しました：\n",
                en: "An error occurred while exporting the artboard \"%1\":\n"
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで選んだアートボードを、選んだ形式で書き出す
     * @returns {void}
     */
    function exportArtboards() {
        if (app.documents.length === 0) {
            return;
        }

        var activeDoc = app.activeDocument;
        /* 一度も保存していないドキュメントは「元ファイルと同じ」を使えない / An unsaved document cannot use "Same as Source File" */
        var sourceFolder = getSourceFolder(activeDoc);
        var baseFileName = activeDoc.name.replace(/\.[^.]+$/, "");

        var pdfPresetNames = getPdfPresetNames();
        var artboardStates = buildArtboardStates(activeDoc, $.global[FORMAT_MEMORY_KEY] || DEFAULT_FORMAT, loadRememberedArtboardSettings(activeDoc), pdfPresetNames);
        var outputSettings = showExportDialog(artboardStates, pdfPresetNames, loadOutputSettings(), {
            sourceFolder: sourceFolder,
            baseFileName: baseFileName,
            activeArtboardIndex: activeDoc.artboards.getActiveArtboardIndex()
        });
        if (!outputSettings) {
            return;
        }
        rememberArtboardSettings(activeDoc, artboardStates);
        /* ［ファイル名］欄の文字列はドキュメントごとなので記憶しない / The edited document name belongs to this document, so it is not remembered */
        var rememberedOutput = {};
        for (var settingKey in outputSettings) {
            if (settingKey !== "documentName") rememberedOutput[settingKey] = outputSettings[settingKey];
        }
        $.global[OUTPUT_MEMORY_KEY] = rememberedOutput;

        var dateStamp = formatDateStamp(outputSettings.dateFormatIndex);
        var outputFolder = outputSettings.useCustomFolder ? new Folder(outputSettings.customFolderPath) : sourceFolder;
        if (outputSettings.createSubfolder) {
            var subfolderName = resolveSubfolderName(outputSettings.subfolderName, baseFileName, dateStamp, outputSettings);
            if (subfolderName) outputFolder = new Folder(outputFolder.fsName + "/" + subfolderName);
        }

        var exportJobs = buildExportJobs(artboardStates, outputFolder, baseFileName, outputSettings, dateStamp);
        if (exportJobs.length === 0) {
            alert(getLabel("alert.noExportTarget"));
            return;
        }
        /* 同じ名前に書き出すと後のファイルが前のファイルを上書きするので止める / Stop when files would overwrite each other */
        var duplicateNames = findDuplicateFileNames(exportJobs);
        if (duplicateNames.length > 0) {
            alert(getLabel("alert.duplicateFileNames", [duplicateNames.slice(0, OVERWRITE_LIST_MAX).join("\n")]));
            return;
        }
        /* PDF は保存済みファイルの複製から作るので、一度も保存していないと書き出せない / PDF is made from a copy of the saved file */
        if (hasPdfJob(exportJobs) && !sourceFolder) {
            alert(getLabel("alert.pdfNeedsSavedDocument"));
            return;
        }
        if (!confirmOverwrite(exportJobs)) {
            return;
        }

        /* 未保存の変更があれば先に保存する / Save pending changes first */
        if (hasPdfJob(exportJobs) && !activeDoc.saved) {
            if (!confirm(getLabel("alert.saveBeforePdf"))) {
                return;
            }
            activeDoc.save();
        }

        if (!outputFolder.exists && !outputFolder.create()) {
            alert(getLabel("alert.folderNotCreated", [decodeURI(outputFolder.fsName)]));
            return;
        }

        runExportJobs(activeDoc, exportJobs, outputSettings);

        if (outputSettings.revealAfterExport) {
            outputFolder.execute();
        }
    }

    /**
     * ドキュメントが保存されているフォルダーを返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Folder|null} フォルダー。一度も保存していなければ null（fullName が例外になることがある）
     */
    function getSourceFolder(doc) {
        try {
            var sourceFolder = doc.fullName.parent;
            if (doc.fullName.exists && sourceFolder && sourceFolder.exists) return sourceFolder;
        } catch (e) {}
        return null;
    }

    /**
     * ドキュメントの全アートボードについて、一覧に出す状態を作る（記憶があればそれを使う）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} initialFormatKey - 形式の初期値（記憶の無いアートボード用）
     * @param {Object} rememberedSettings - loadRememberedArtboardSettings() の結果
     * @param {string[]} pdfPresetNames - PDF プリセット名
     * @returns {Array<{index: number, name: string, enabled: boolean, formatKey: string, scales: number[], backgroundKey: string, pdfPreset: string, jpegQuality: number}>} アートボードの状態
     */
    function buildArtboardStates(doc, initialFormatKey, rememberedSettings, pdfPresetNames) {
        var artboardStates = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardName = doc.artboards[i].name;
            var remembered = rememberedSettings["ab:" + artboardName];
            artboardStates.push({
                index: i,
                name: artboardName,
                enabled: remembered ? remembered.enabled : true,
                formatKey: getFormatInfo(remembered ? remembered.formatKey : initialFormatKey).key,
                scales: (remembered ? remembered.scales : DEFAULT_SCALES).slice(0),
                backgroundKey: (remembered && remembered.backgroundKey) || DEFAULT_BACKGROUND,
                pdfPreset: resolvePdfPreset(pdfPresetNames, (remembered && remembered.pdfPreset) || $.global[PDF_PRESET_MEMORY_KEY]),
                jpegQuality: (remembered && remembered.jpegQuality !== undefined) ? remembered.jpegQuality : JPEG_QUALITY
            });
        }
        return artboardStates;
    }

    /**
     * アートボードを書き出すかを返す（［書き出す］が ON で、除外に当たらない）
     * @param {Object} artboardState - アートボードの状態
     * @param {Object} exclusionOptions - 除外の設定
     * @returns {boolean} 書き出すなら true
     */
    function isArtboardExported(artboardState, exclusionOptions) {
        return artboardState.enabled && !isArtboardExcluded(artboardState.name, exclusionOptions);
    }

    /**
     * 書き出すアートボードの状態から、1ファイルずつの書き出しジョブを作る
     * @param {Object[]} artboardStates - アートボードの状態
     * @param {Folder} outputFolder - 保存先フォルダー
     * @param {string} baseFileName - 拡張子を除いたドキュメント名
     * @param {Object} outputSettings - ファイル名・除外・PDF のまとめの設定
     * @param {string} dateStamp - 日付
     * @returns {Array<{artboardIndexes: number[], formatKey: string, displayName: string, scale: number, backgroundKey: string, pdfPreset: string, jpegQuality: number, file: File}>} 書き出しジョブ
     */
    function buildExportJobs(artboardStates, outputFolder, baseFileName, outputSettings, dateStamp) {
        var exportJobs = [];
        var combinedPdfStates = [];
        for (var i = 0; i < artboardStates.length; i++) {
            var artboardState = artboardStates[i];
            if (!isArtboardExported(artboardState, outputSettings)) continue;
            var formatInfo = getFormatInfo(artboardState.formatKey);
            /* まとめる PDF はあとで1つのジョブにする / Combined PDFs become one job afterwards */
            if (formatInfo.key === "pdf" && outputSettings.combinePdf) {
                combinedPdfStates.push(artboardState);
                continue;
            }
            /* PDF は倍率を持たないので1ファイルだけ / PDF has no scale, so one file per artboard */
            var jobScales = (formatInfo.key === "pdf") ? [100] : artboardState.scales;
            for (var j = 0; j < jobScales.length; j++) {
                var fileStem = buildFileStem(outputSettings, {
                    baseFileName: baseFileName,
                    artboardState: artboardState,
                    artboardCount: artboardStates.length,
                    scale: jobScales[j],
                    dateStamp: dateStamp
                });
                exportJobs.push({
                    artboardIndexes: [artboardState.index],
                    formatKey: formatInfo.key,
                    displayName: artboardState.name + ((jobScales[j] === 100) ? "" : " " + jobScales[j] + "%"),
                    scale: jobScales[j],
                    backgroundKey: getEffectiveBackgroundKey(artboardState),
                    pdfPreset: artboardState.pdfPreset,
                    jpegQuality: artboardState.jpegQuality,
                    file: new File(outputFolder.fsName + "/" + fileStem + "." + formatInfo.extension)
                });
            }
        }
        if (combinedPdfStates.length > 0) {
            var combinedIndexes = [];
            for (var k = 0; k < combinedPdfStates.length; k++) combinedIndexes.push(combinedPdfStates[k].index);
            var combinedStem = buildCombinedPdfStem(outputSettings, baseFileName, dateStamp);
            exportJobs.push({
                artboardIndexes: combinedIndexes,
                formatKey: "pdf",
                displayName: combinedStem + ".pdf",
                scale: 100,
                backgroundKey: DEFAULT_BACKGROUND,
                pdfPreset: combinedPdfStates[0].pdfPreset,
                jpegQuality: JPEG_QUALITY,
                file: new File(outputFolder.fsName + "/" + combinedStem + ".pdf")
            });
        }
        return exportJobs;
    }

    /**
     * PDF の書き出しジョブがあるかを返す
     * @param {Object[]} exportJobs - 書き出しジョブ
     * @returns {boolean} PDF が1つでもあれば true
     */
    function hasPdfJob(exportJobs) {
        for (var i = 0; i < exportJobs.length; i++) {
            if (exportJobs[i].formatKey === "pdf") return true;
        }
        return false;
    }

    /**
     * すでにあるファイルを上書きしてよいか、まとめて確認する
     * @param {Object[]} exportJobs - 書き出しジョブ
     * @returns {boolean} 書き出してよいとき true（既存ファイルが無いときも true）
     */
    function confirmOverwrite(exportJobs) {
        var existingNames = [];
        for (var i = 0; i < exportJobs.length; i++) {
            if (exportJobs[i].file.exists) existingNames.push(decodeURI(exportJobs[i].file.name));
        }
        if (existingNames.length === 0) {
            return true;
        }
        var listedNames = existingNames.slice(0, OVERWRITE_LIST_MAX);
        if (existingNames.length > OVERWRITE_LIST_MAX) listedNames.push("…");
        return confirm(getLabel("alert.overwrite", [existingNames.length, listedNames.join("\n")]));
    }

    /**
     * 書き出しジョブを進捗表示つきで順に実行する
     * @param {Document} activeDoc - 書き出すドキュメント
     * @param {Object[]} exportJobs - 書き出しジョブ
     * @param {{excludeLayers: boolean, excludeLayerPrefix: string}} outputSettings - レイヤーの除外の設定
     * @returns {void}
     */
    function runExportJobs(activeDoc, exportJobs, outputSettings) {
        var originalArtboardIndex = activeDoc.artboards.getActiveArtboardIndex();
        var progress = createProgressWindow(exportJobs.length);
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var layerPrefix = outputSettings.excludeLayers ? outputSettings.excludeLayerPrefix : "";
        var hiddenLayers = [];
        var pdfCopyFile = null;
        var pdfSourceDoc = null;
        var cancelled = false;
        /* 途中で失敗しても、レイヤー・複製・警告表示の設定・進捗ウィンドウは戻す
           Always restore layers, clean up the copy, restore alerts and close the progress window */
        try {
            if (layerPrefix) hiddenLayers = hideExcludedLayers(activeDoc, layerPrefix);
            /* saveAs() は保存先をそのドキュメントに結び付けるので、PDF は複製から作る / saveAs() rebinds the document to the new file, so PDF is made from a copy */
            if (hasPdfJob(exportJobs)) {
                pdfCopyFile = copyDocumentFileToTemp(activeDoc);
                pdfSourceDoc = app.open(pdfCopyFile);
                /* 複製は保存しないので、隠したまま閉じてよい / The copy is never saved, so it can close with the layers hidden */
                if (layerPrefix) hideExcludedLayers(pdfSourceDoc, layerPrefix);
            }
            for (var i = 0; i < exportJobs.length; i++) {
                progress.update(i, exportJobs[i].displayName);
                /* キャンセルボタンが押されていれば中断 / Stop if the cancel button was pressed */
                if (progress.isCancelled()) {
                    cancelled = true;
                    break;
                }
                if (exportJobs[i].formatKey === "pdf") {
                    saveArtboardsAsPdf(pdfSourceDoc, exportJobs[i]);
                } else {
                    /* PDF 用の複製を開いていても、画像は元のドキュメントから書き出す / Images come from the original even while the PDF copy is open */
                    if (pdfSourceDoc) app.activeDocument = activeDoc;
                    activeDoc.artboards.setActiveArtboardIndex(exportJobs[i].artboardIndexes[0]);
                    exportArtboardAsImage(activeDoc, exportJobs[i]);
                }
            }
            progress.update(cancelled ? i : exportJobs.length, getLabel(cancelled ? "status.cancelled" : "status.done"));
        } finally {
            if (pdfSourceDoc) {
                pdfSourceDoc.close(SaveOptions.DONOTSAVECHANGES);
                app.activeDocument = activeDoc;
            }
            if (pdfCopyFile) {
                pdfCopyFile.remove();
            }
            restoreHiddenLayers(hiddenLayers);
            activeDoc.artboards.setActiveArtboardIndex(originalArtboardIndex);
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            progress.close();
        }
    }

    // =========================================
    // 書き出しダイアログ / Export dialog
    // =========================================

    /**
     * 一覧の「倍率」列の文字列を作る
     * @param {Object} artboardState - アートボードの状態
     * @returns {string} 倍率の表示（PDF は「—」）
     */
    function formatScaleColumn(artboardState) {
        if (artboardState.formatKey === "pdf") return "—";
        return artboardState.scales.join("% / ") + "%";
    }

    /**
     * 形式に合わせた実際の背景を返す（JPEG は透明にできないので白）
     * @param {Object} artboardState - アートボードの状態
     * @returns {string} "transparent" / "white" / "black"
     */
    function getEffectiveBackgroundKey(artboardState) {
        if (artboardState.formatKey === "jpeg" && artboardState.backgroundKey === "transparent") return "white";
        return artboardState.backgroundKey;
    }

    /**
     * 一覧の「背景」列の文字列を作る
     * @param {Object} artboardState - アートボードの状態
     * @returns {string} 背景の表示（PDF は「—」）
     */
    function formatBackgroundColumn(artboardState) {
        if (artboardState.formatKey === "pdf") return "—";
        return getLabel("radio." + getEffectiveBackgroundKey(artboardState));
    }

    /**
     * 倍率欄の文字列を数値の配列にする（カンマ・空白・読点で区切る。全角も可。重複は1つに）
     * @param {string} scaleText - 倍率欄の文字列
     * @returns {number[]|null} 倍率の配列。空・0 以下・数値でないものがあれば null
     */
    function parseScales(scaleText) {
        /* 全角の数字・小数点は半角にそろえる / Normalize full-width digits and dots */
        var halfWidthText = String(scaleText).replace(/[０-９．]/g, function (fullWidthChar) {
            return String.fromCharCode(fullWidthChar.charCodeAt(0) - 0xFEE0);
        });
        var scaleTokens = halfWidthText.replace(/[%％]/g, " ").replace(/[、，]/g, ",").split(/[,\s]+/);
        var scales = [];
        var seenScales = {};
        for (var i = 0; i < scaleTokens.length; i++) {
            if (scaleTokens[i] === "") continue;
            var scaleValue = Number(scaleTokens[i]);
            if (isNaN(scaleValue) || scaleValue <= 0) return null;
            if (seenScales["s" + scaleValue]) continue;
            seenScales["s" + scaleValue] = true;
            scales.push(scaleValue);
        }
        return (scales.length > 0) ? scales : null;
    }

    /**
     * JPEG で上限を超える倍率のアートボード名を返す
     * @param {Object[]} artboardStates - アートボードの状態
     * @param {Object} exclusionOptions - 除外の設定
     * @returns {string[]} アートボード名
     */
    function findTooLargeJpegScales(artboardStates, exclusionOptions) {
        var artboardNames = [];
        for (var i = 0; i < artboardStates.length; i++) {
            var artboardState = artboardStates[i];
            if (!isArtboardExported(artboardState, exclusionOptions) || artboardState.formatKey !== "jpeg") continue;
            for (var j = 0; j < artboardState.scales.length; j++) {
                if (artboardState.scales[j] > JPEG_MAX_SCALE) {
                    artboardNames.push(artboardState.name + " (" + artboardState.scales[j] + "%)");
                    break;
                }
            }
        }
        return artboardNames;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * 背景の色見本（白・黒・透明の市松模様）を追加する。クリックすると onSelect を呼ぶ
     * @param {Group} parent - 追加先
     * @param {string} backgroundKey - "transparent" / "white" / "black"
     * @param {RadioButton} linkedRadio - 対になるラジオボタン（無効のときは薄く描く）
     * @param {Function} onSelect - クリック時の処理
     * @returns {Group} 色見本
     */
    function addBackgroundSwatch(parent, backgroundKey, linkedRadio, onSelect) {
        var swatchBox = parent.add("group");
        swatchBox.preferredSize = [SWATCH_SIZE, SWATCH_SIZE];
        swatchBox.minimumSize = [SWATCH_SIZE, SWATCH_SIZE];
        swatchBox.maximumSize = [SWATCH_SIZE, SWATCH_SIZE];
        swatchBox.onDraw = function () {
            var swatchGraphics = swatchBox.graphics;
            var alpha = isEnabledInTree(linkedRadio) ? 1 : 0.35;
            /* 地の色（透明は明るい灰色の上に濃い灰色の市松）/ Base color; transparent is a light gray checker */
            var baseColor = (backgroundKey === "black") ? [0, 0, 0, alpha] : (backgroundKey === "white") ? [1, 1, 1, alpha] : [0.85, 0.85, 0.85, alpha];
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, SWATCH_SIZE, SWATCH_SIZE);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, baseColor));
            if (backgroundKey === "transparent") {
                var checkerBrush = swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, [0.55, 0.55, 0.55, alpha]);
                for (var row = 0; row * SWATCH_CHECKER_SIZE < SWATCH_SIZE; row++) {
                    for (var column = 0; column * SWATCH_CHECKER_SIZE < SWATCH_SIZE; column++) {
                        if ((row + column) % 2 === 0) continue;
                        /* rectPath の前に newPath しないとパスが重なっていく / Without newPath the paths accumulate */
                        swatchGraphics.newPath();
                        swatchGraphics.rectPath(column * SWATCH_CHECKER_SIZE, row * SWATCH_CHECKER_SIZE, SWATCH_CHECKER_SIZE, SWATCH_CHECKER_SIZE);
                        swatchGraphics.fillPath(checkerBrush);
                    }
                }
            }
            /* 白が地に溶けないよう枠を付ける / Frame so white does not melt into the background */
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, SWATCH_SIZE - 1, SWATCH_SIZE - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, alpha], 1));
        };
        swatchBox.addEventListener("mousedown", function () {
            if (!isEnabledInTree(linkedRadio)) return;
            onSelect();
        });
        return swatchBox;
    }

    /**
     * アートボードの一覧と各種設定を持つダイアログを開く。アートボードごとの設定は artboardStates に書き戻す
     * @param {Object[]} artboardStates - アートボードの状態（書き換えられる）
     * @param {string[]} pdfPresetNames - PDF プリセット名
     * @param {Object} initialOutput - アートボードごと以外の設定の初期値（loadOutputSettings() の形）
     * @param {{sourceFolder: Folder|null, baseFileName: string, activeArtboardIndex: number}} documentInfo - ドキュメントの情報
     * @returns {Object|null} アートボードごと以外の設定（loadOutputSettings() の形に documentName を足したもの）。キャンセル時は null
     */
    function showExportDialog(artboardStates, pdfPresetNames, initialOutput, documentInfo) {
        var sourceFolder = documentInfo.sourceFolder;
        var baseFileName = documentInfo.baseFileName;
        var exportDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(exportDialog);

        /* 上段：左に一覧、右に選んだ行の設定 / Top: the list on the left, settings for the selected rows on the right */
        var topRow = exportDialog.add("group");
        topRow.orientation = "row";
        topRow.alignChildren = ["fill", "top"];
        topRow.spacing = COLUMN_SPACING;

        var listColumn = topRow.add("group");
        listColumn.orientation = "column";
        listColumn.alignChildren = ["fill", "top"];
        listColumn.spacing = LIST_COLUMN_SPACING;

        /* 書き出し対象をまとめて決める / Set the export targets at once */
        var targetRow = listColumn.add("group");
        setupRow(targetRow);
        var btnTargetAll = targetRow.add("button", undefined, getLabel("button.targetAll"));
        btnTargetAll.helpTip = getLabel("tooltip.targetAll");
        var btnTargetActive = targetRow.add("button", undefined, getLabel("button.targetActive"));
        btnTargetActive.helpTip = getLabel("tooltip.targetActive");
        targetRow.add("statictext", undefined, labelText("fieldLabel.targetRange"));
        var targetRangeField = targetRow.add("edittext", undefined, "");
        targetRangeField.characters = TARGET_RANGE_CHARS;
        targetRangeField.helpTip = getLabel("tooltip.targetRange");
        var btnApplyRange = targetRow.add("button", undefined, getLabel("button.applyRange"));

        /* アートボード一覧（Mac では1列目が長い文字列で広がるため、先頭は短い番号にする）
           Artboard list (the first column widens with long text on Mac, so it holds the short number) */
        var artboardList = listColumn.add("listbox", undefined, [], {
            multiselect: true,
            numberOfColumns: 6,
            showHeaders: true,
            columnTitles: [getLabel("column.number"), getLabel("column.exportMark"), getLabel("column.name"), getLabel("column.format"), getLabel("column.scale"), getLabel("column.background")],
            columnWidths: ARTBOARD_LIST_COLUMNS
        });
        artboardList.preferredSize = [ARTBOARD_LIST_WIDTH, ARTBOARD_LIST_HEIGHT];
        artboardList.helpTip = getLabel("tooltip.artboardList");

        var allIndexes = [];
        for (var j = 0; j < artboardStates.length; j++) {
            var listItem = artboardList.add("item", String(artboardStates[j].index + 1));
            listItem.subItems[1].text = artboardStates[j].name;
            allIndexes.push(j);
        }

        /* 除外（前方一致）/ Exclusion by prefix */
        var excludeArtboardRow = listColumn.add("group");
        setupRow(excludeArtboardRow);
        var chkExcludeArtboards = excludeArtboardRow.add("checkbox", undefined, labelText("checkbox.excludeArtboards"));
        chkExcludeArtboards.helpTip = getLabel("tooltip.excludeArtboards");
        var excludeArtboardPrefixField = excludeArtboardRow.add("edittext", undefined, initialOutput.excludeArtboardPrefix);
        excludeArtboardPrefixField.characters = PREFIX_FIELD_CHARS;
        excludeArtboardRow.add("statictext", undefined, getLabel("fieldLabel.prefixMatch"));

        var excludeLayerRow = listColumn.add("group");
        setupRow(excludeLayerRow);
        var chkExcludeLayers = excludeLayerRow.add("checkbox", undefined, labelText("checkbox.excludeLayers"));
        chkExcludeLayers.helpTip = getLabel("tooltip.excludeLayers");
        var excludeLayerPrefixField = excludeLayerRow.add("edittext", undefined, initialOutput.excludeLayerPrefix);
        excludeLayerPrefixField.characters = PREFIX_FIELD_CHARS;
        excludeLayerRow.add("statictext", undefined, getLabel("fieldLabel.prefixMatch"));

        chkExcludeArtboards.value = initialOutput.excludeArtboards;
        chkExcludeLayers.value = initialOutput.excludeLayers;

        /* 選んだ行の設定 / Settings for the selected rows */
        var settingsPanel = topRow.add("panel", undefined, getLabel("panel.selectedArtboards"));
        setupPanel(settingsPanel);

        var chkExportTarget = settingsPanel.add("checkbox", undefined, getLabel("checkbox.exportTarget"));

        var formatRow = settingsPanel.add("group");
        setupRow(formatRow);
        var formatLabel = formatRow.add("statictext", undefined, labelText("fieldLabel.format"));
        var formatRadios = [];
        for (var i = 0; i < EXPORT_FORMATS.length; i++) {
            formatRadios.push(formatRow.add("radiobutton", undefined, EXPORT_FORMATS[i].label));
        }
        var lastChosenFormatKey = null;

        var scaleRow = settingsPanel.add("group");
        setupRow(scaleRow);
        var scaleLabel = scaleRow.add("statictext", undefined, labelText("fieldLabel.scale"));
        var scaleField = scaleRow.add("edittext", undefined, "");
        scaleField.characters = SCALE_FIELD_CHARS;
        scaleField.helpTip = getLabel("tooltip.scaleField");
        scaleRow.add("statictext", undefined, "% " + getLabel("fieldLabel.jpegScaleLimit", [JPEG_MAX_SCALE]));

        /* 背景：ラジオボタンの右に色見本 / Background: a swatch right of each radio button */
        var backgroundRow = settingsPanel.add("group");
        setupRow(backgroundRow);
        var backgroundLabel = backgroundRow.add("statictext", undefined, labelText("fieldLabel.background"));
        var backgroundChoices = [];
        var backgroundKeys = ["white", "black", "transparent"];
        for (var b = 0; b < backgroundKeys.length; b++) {
            var backgroundRadio = backgroundRow.add("radiobutton", undefined, "");
            backgroundRadio.helpTip = getLabel("radio." + backgroundKeys[b]);
            var backgroundSwatch = addBackgroundSwatch(backgroundRow, backgroundKeys[b], backgroundRadio, (function (choiceIndex) {
                return function () {
                    selectBackgroundChoice(choiceIndex);
                };
            })(b));
            backgroundSwatch.helpTip = getLabel("radio." + backgroundKeys[b]);
            backgroundChoices.push({ key: backgroundKeys[b], radio: backgroundRadio, swatch: backgroundSwatch });
        }

        var jpegQualityRow = settingsPanel.add("group");
        setupRow(jpegQualityRow);
        var jpegQualityLabel = jpegQualityRow.add("statictext", undefined, labelText("fieldLabel.jpegQuality"));
        var jpegQualityField = jpegQualityRow.add("edittext", undefined, "");
        jpegQualityField.characters = JPEG_QUALITY_FIELD_CHARS;
        jpegQualityRow.add("statictext", undefined, "(0-100)");

        var pdfPresetRow = settingsPanel.add("group");
        setupRow(pdfPresetRow);
        var pdfPresetLabel = pdfPresetRow.add("statictext", undefined, labelText("fieldLabel.pdfPreset"));
        var pdfPresetDropdown = pdfPresetRow.add("dropdownlist");
        for (var p = 0; p < pdfPresetNames.length; p++) pdfPresetDropdown.add("item", pdfPresetNames[p]);
        pdfPresetDropdown.preferredSize.width = PDF_PRESET_WIDTH;
        var lastChosenPdfPreset = null;
        var isLoadingFields = false;

        /* PDF のまとめは全体の設定（行ごとではない）/ Combining PDFs is a global setting, not per row */
        var chkCombinePdf = settingsPanel.add("checkbox", undefined, getLabel("checkbox.combinePdf"));
        chkCombinePdf.helpTip = getLabel("tooltip.combinePdf");
        chkCombinePdf.value = initialOutput.combinePdf;

        alignLabelWidths([formatLabel, scaleLabel, backgroundLabel, jpegQualityLabel, pdfPresetLabel]);

        /* ファイル名 / File name */
        var fileNamePanel = exportDialog.add("panel", undefined, getLabel("panel.fileName"));
        setupPanel(fileNamePanel, 6);
        var fileNamePartsRow = fileNamePanel.add("group");
        setupRow(fileNamePartsRow);
        var rbCombineParts = fileNamePartsRow.add("radiobutton", undefined, getLabel("radio.combineParts"));
        var chkDocumentName = fileNamePartsRow.add("checkbox", undefined, "");
        chkDocumentName.helpTip = getLabel("checkbox.documentName");
        var documentNameField = fileNamePartsRow.add("edittext", undefined, baseFileName);
        documentNameField.characters = DOCUMENT_NAME_CHARS;
        documentNameField.helpTip = getLabel("checkbox.documentName");
        fileNamePartsRow.add("statictext", undefined, "+");
        var artboardPatternDropdown = fileNamePartsRow.add("dropdownlist");
        for (var q = 0; q < ARTBOARD_NAME_PATTERNS.length; q++) {
            var patternLabel = ARTBOARD_NAME_PATTERNS[q] === "" ? getLabel("dropdown.noArtboardPart")
                : ARTBOARD_NAME_PATTERNS[q].replace("{n}", getLabel("dropdown.artboardNumber")).replace("{a}", getLabel("dropdown.artboardName"));
            artboardPatternDropdown.add("item", patternLabel);
        }
        fileNamePartsRow.add("statictext", undefined, "+");
        var suffixSeparatorDropdown = fileNamePartsRow.add("dropdownlist", undefined, SEPARATOR_CHOICES);
        suffixSeparatorDropdown.helpTip = getLabel("tooltip.suffixSeparator");
        var chkDate = fileNamePartsRow.add("checkbox", undefined, "");
        chkDate.helpTip = getLabel("checkbox.date");
        var dateFormatDropdown = fileNamePartsRow.add("dropdownlist");
        for (var d = 0; d < DATE_FORMATS.length; d++) dateFormatDropdown.add("item", formatDateStamp(d));
        dateFormatDropdown.helpTip = getLabel("tooltip.dateFormat");
        fileNamePartsRow.add("statictext", undefined, "+");
        var chkSuffixText = fileNamePartsRow.add("checkbox", undefined, "");
        chkSuffixText.helpTip = getLabel("checkbox.suffixText");
        var suffixTextField = fileNamePartsRow.add("edittext", undefined, initialOutput.suffixText);
        suffixTextField.characters = SUFFIX_TEXT_CHARS;
        suffixTextField.helpTip = getLabel("checkbox.suffixText");

        /* ［組み合わせ］と［書式］は別の行なので、排他は手で管理する / The radios sit in different rows, so exclusivity is handled by hand */
        var fileNameTemplateRow = fileNamePanel.add("group");
        setupRow(fileNameTemplateRow);
        var rbFileNameTemplate = fileNameTemplateRow.add("radiobutton", undefined, getLabel("radio.fileNameTemplate"));
        var fileNameTemplateField = fileNameTemplateRow.add("edittext", undefined, "");
        fileNameTemplateField.characters = FILE_NAME_TEMPLATE_CHARS;
        fileNameTemplateField.helpTip = getLabel("tooltip.fileNameTemplate");
        var fileNameTokensText = fileNameTemplateRow.add("statictext", undefined, labelValueText("fieldLabel.tokens", "{doc} {name} {num} {scale} {date}"));
        fileNameTokensText.helpTip = getLabel("tooltip.fileNameTemplate");

        /* 置き換え / Replacement */
        var replaceRow = fileNamePanel.add("group");
        setupRow(replaceRow);
        replaceRow.add("statictext", undefined, labelText("fieldLabel.replaceChars") + " / \\ : * ? \" < > | ¥ →");
        var invalidReplacementDropdown = replaceRow.add("dropdownlist", undefined, SEPARATOR_CHOICES);
        invalidReplacementDropdown.helpTip = getLabel("tooltip.replaceChars");
        var chkReplaceSpaces = replaceRow.add("checkbox", undefined, getLabel("checkbox.replaceSpaces") + " →");
        var spaceReplacementDropdown = replaceRow.add("dropdownlist", undefined, SEPARATOR_CHOICES);

        var fileNameExampleText = fileNamePanel.add("statictext", undefined, "", { truncate: "middle" });
        fileNameExampleText.preferredSize.width = FOLDER_PATH_WIDTH;

        /* 選択肢の番号を値から探す / Find the index of a value among the choices */
        function indexOfChoice(choiceValues, value) {
            for (var k = 0; k < choiceValues.length; k++) {
                if (choiceValues[k] === value) return k;
            }
            return 0;
        }

        chkDocumentName.value = initialOutput.useDocumentName;
        artboardPatternDropdown.selection = (initialOutput.artboardPatternIndex >= 0 && initialOutput.artboardPatternIndex < ARTBOARD_NAME_PATTERNS.length) ? initialOutput.artboardPatternIndex : DEFAULT_ARTBOARD_PATTERN;
        suffixSeparatorDropdown.selection = indexOfChoice(SEPARATOR_CHOICES, initialOutput.suffixSeparator);
        chkDate.value = initialOutput.useDate;
        dateFormatDropdown.selection = (initialOutput.dateFormatIndex >= 0 && initialOutput.dateFormatIndex < DATE_FORMATS.length) ? initialOutput.dateFormatIndex : DEFAULT_DATE_FORMAT;
        chkSuffixText.value = initialOutput.useSuffixText;
        fileNameTemplateField.text = initialOutput.fileNameTemplate;
        rbCombineParts.value = !initialOutput.useFileNameTemplate;
        rbFileNameTemplate.value = !!initialOutput.useFileNameTemplate;
        invalidReplacementDropdown.selection = indexOfChoice(SEPARATOR_CHOICES, initialOutput.invalidCharReplacement);
        chkReplaceSpaces.value = initialOutput.replaceSpaces;
        spaceReplacementDropdown.selection = indexOfChoice(SEPARATOR_CHOICES, initialOutput.spaceReplacement);

        /* 保存先 / Destination */
        var destinationPanel = exportDialog.add("panel", undefined, getLabel("panel.destination"));
        setupPanel(destinationPanel, 6);
        var destinationRadioRow = destinationPanel.add("group");
        setupRow(destinationRadioRow);
        var rbSameFolder = destinationRadioRow.add("radiobutton", undefined, getLabel("radio.sameFolder"));
        var rbCustomFolder = destinationRadioRow.add("radiobutton", undefined, getLabel("radio.customFolder"));
        var destinationPathRow = destinationPanel.add("group");
        setupRow(destinationPathRow);
        var destinationPathText = destinationPathRow.add("statictext", undefined, "", { truncate: "middle" });
        destinationPathText.preferredSize.width = FOLDER_PATH_WIDTH;
        var btnChooseFolder = destinationPathRow.add("button", undefined, getLabel("button.chooseFolder"));
        var subfolderRow = destinationPanel.add("group");
        setupRow(subfolderRow);
        var chkCreateSubfolder = subfolderRow.add("checkbox", undefined, labelText("checkbox.createSubfolder"));
        var subfolderNameField = subfolderRow.add("edittext", undefined, initialOutput.subfolderName);
        subfolderNameField.characters = SUBFOLDER_NAME_CHARS;
        subfolderNameField.helpTip = getLabel("tooltip.subfolderName");
        var subfolderTokensText = subfolderRow.add("statictext", undefined, labelValueText("fieldLabel.tokens", "{doc} {date}"));
        subfolderTokensText.helpTip = getLabel("tooltip.subfolderName");
        var chkRevealFolder = destinationPanel.add("checkbox", undefined, getLabel("checkbox.revealFolder"));

        var customFolderPath = initialOutput.customFolderPath || "";
        /* 未保存のドキュメントは「元ファイルと同じ」を使えない / An unsaved document cannot use "Same as Source File" */
        rbSameFolder.enabled = !!sourceFolder;
        var useCustomFolder = initialOutput.useCustomFolder || !sourceFolder;
        rbSameFolder.value = !useCustomFolder;
        rbCustomFolder.value = useCustomFolder;
        chkCreateSubfolder.value = initialOutput.createSubfolder;
        chkRevealFolder.value = initialOutput.revealAfterExport;

        var buttonRow = addButtonRow(exportDialog);
        var btnSelectAll = buttonRow.leftGroup.add("button", undefined, getLabel("button.selectAll"));
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var scaleFieldEdited = false;

        /* 今の除外の設定 / Current exclusion settings */
        function getExclusionOptions() {
            return { excludeArtboards: chkExcludeArtboards.value, excludeArtboardPrefix: excludeArtboardPrefixField.text };
        }

        /* 選んでいる行の番号 / Indexes of the selected rows */
        function getSelectedIndexes() {
            var selectedIndexes = [];
            if (!artboardList.selection) return selectedIndexes;
            for (var k = 0; k < artboardList.selection.length; k++) {
                selectedIndexes.push(artboardList.selection[k].index);
            }
            return selectedIndexes;
        }

        /* 1行の表示を状態に合わせる / Refresh one row from its state */
        function refreshRow(rowIndex) {
            var rowState = artboardStates[rowIndex];
            var rowItem = artboardList.items[rowIndex];
            var isExcluded = isArtboardExcluded(rowState.name, getExclusionOptions());
            rowItem.subItems[0].text = isExcluded ? getLabel("column.excludedMark") : (rowState.enabled ? "✓" : "");
            rowItem.subItems[2].text = getFormatInfo(rowState.formatKey).label;
            rowItem.subItems[3].text = formatScaleColumn(rowState);
            rowItem.subItems[4].text = formatBackgroundColumn(rowState);
            /* Mac では subItems を変えただけでは描き直されないため、1列目を書き換えて行を描き直させる
               On Mac changing subItems alone does not repaint, so rewrite the first column to force it */
            var numberText = rowItem.text;
            rowItem.text = "";
            rowItem.text = numberText;
        }

        /* すべての行を描き直す / Refresh every row */
        function refreshAllRows() {
            for (var k = 0; k < artboardStates.length; k++) refreshRow(k);
            updateOKButton();
            updateFileNameExample();
        }

        /* 書き出す行が1つも無いときは書き出せない / Export needs at least one row to export */
        function updateOKButton() {
            var hasTarget = false;
            for (var k = 0; k < artboardStates.length; k++) {
                if (isArtboardExported(artboardStates[k], getExclusionOptions())) hasTarget = true;
            }
            btnOK.enabled = hasTarget;
        }

        /* 色見本は自作描画なので、有効・無効が変わったら描き直す / Swatches are custom-drawn, so redraw them when the enabled state changes */
        function redrawSwatches() {
            if (!exportDialog.visible) return;
            for (var k = 0; k < backgroundChoices.length; k++) {
                backgroundChoices[k].swatch.hide();
                backgroundChoices[k].swatch.show();
            }
        }

        /* 形式で使わない設定を無効にする（選んだ先頭の行の形式で決める）/ Disable settings the format of the first selected row does not use */
        function updateFormatControls() {
            var selectedIndexes = getSelectedIndexes();
            if (selectedIndexes.length === 0) return;
            var formatKey = artboardStates[selectedIndexes[0]].formatKey;
            scaleRow.enabled = (formatKey !== "pdf");
            backgroundRow.enabled = (formatKey !== "pdf");
            /* JPEG は透明にできない / JPEG cannot be transparent */
            for (var k = 0; k < backgroundChoices.length; k++) {
                backgroundChoices[k].radio.enabled = (backgroundChoices[k].key !== "transparent" || formatKey === "png");
            }
            jpegQualityRow.enabled = (formatKey === "jpeg");
            pdfPresetRow.enabled = (formatKey === "pdf" && pdfPresetNames.length > 0);
            redrawSwatches();
        }

        /* 選んだ先頭の行の設定を欄に読み込む / Load the first selected row into the fields */
        function loadSelectedSettings() {
            var selectedIndexes = getSelectedIndexes();
            settingsPanel.enabled = (selectedIndexes.length > 0);
            scaleFieldEdited = false;
            if (selectedIndexes.length === 0) {
                redrawSwatches();
                return;
            }
            var firstState = artboardStates[selectedIndexes[0]];
            chkExportTarget.value = firstState.enabled;
            for (var k = 0; k < EXPORT_FORMATS.length; k++) {
                formatRadios[k].value = (EXPORT_FORMATS[k].key === firstState.formatKey);
            }
            scaleField.text = firstState.scales.join(", ");
            var backgroundKey = getEffectiveBackgroundKey(firstState);
            for (var c = 0; c < backgroundChoices.length; c++) {
                backgroundChoices[c].radio.value = (backgroundChoices[c].key === backgroundKey);
            }
            jpegQualityField.text = String(firstState.jpegQuality);
            isLoadingFields = true;
            pdfPresetDropdown.selection = null;
            for (var n = 0; n < pdfPresetNames.length; n++) {
                if (pdfPresetNames[n] === firstState.pdfPreset) pdfPresetDropdown.selection = n;
            }
            isLoadingFields = false;
            updateFormatControls();
            updateFileNameExample();
        }

        /* 今のファイル名の設定 / Current file name settings */
        function getFileNameOptions() {
            return {
                useFileNameTemplate: rbFileNameTemplate.value,
                fileNameTemplate: fileNameTemplateField.text,
                useDocumentName: chkDocumentName.value,
                documentName: documentNameField.text,
                artboardPatternIndex: artboardPatternDropdown.selection ? artboardPatternDropdown.selection.index : DEFAULT_ARTBOARD_PATTERN,
                suffixSeparator: suffixSeparatorDropdown.selection ? SEPARATOR_CHOICES[suffixSeparatorDropdown.selection.index] : DEFAULT_SUFFIX_SEPARATOR,
                useDate: chkDate.value,
                dateFormatIndex: dateFormatDropdown.selection ? dateFormatDropdown.selection.index : DEFAULT_DATE_FORMAT,
                useSuffixText: chkSuffixText.value,
                suffixText: suffixTextField.text,
                invalidCharReplacement: invalidReplacementDropdown.selection ? SEPARATOR_CHOICES[invalidReplacementDropdown.selection.index] : DEFAULT_INVALID_CHAR_REPLACEMENT,
                replaceSpaces: chkReplaceSpaces.value,
                spaceReplacement: spaceReplacementDropdown.selection ? SEPARATOR_CHOICES[spaceReplacementDropdown.selection.index] : DEFAULT_INVALID_CHAR_REPLACEMENT
            };
        }

        /* 選んでいない方式・チェックの外れた欄を無効にする / Disable fields of the method not chosen and of unchecked parts */
        function updateFileNameControls() {
            var useTemplate = rbFileNameTemplate.value;
            chkDocumentName.enabled = !useTemplate;
            documentNameField.enabled = !useTemplate && chkDocumentName.value;
            artboardPatternDropdown.enabled = !useTemplate;
            suffixSeparatorDropdown.enabled = !useTemplate;
            chkDate.enabled = !useTemplate;
            /* 日付の書式は［書式］の {date} にも使う / The date format also serves {date} in Pattern */
            dateFormatDropdown.enabled = useTemplate || chkDate.value;
            chkSuffixText.enabled = !useTemplate;
            suffixTextField.enabled = !useTemplate && chkSuffixText.value;
            fileNameTemplateField.enabled = useTemplate;
            fileNameTokensText.enabled = useTemplate;
            spaceReplacementDropdown.enabled = chkReplaceSpaces.value;
            excludeArtboardPrefixField.enabled = chkExcludeArtboards.value;
            excludeLayerPrefixField.enabled = chkExcludeLayers.value;
            subfolderNameField.enabled = chkCreateSubfolder.value;
            subfolderTokensText.enabled = chkCreateSubfolder.value;
        }

        /* 選んだ先頭の行（無ければ先頭の行）でファイル名の例を出す / Show an example file name for the first selected (or first) row */
        function updateFileNameExample() {
            var selectedIndexes = getSelectedIndexes();
            var exampleState = selectedIndexes.length > 0 ? artboardStates[selectedIndexes[0]] : artboardStates[0];
            if (!exampleState) return;
            var fileNameOptions = getFileNameOptions();
            var exampleDateStamp = formatDateStamp(fileNameOptions.dateFormatIndex);
            var exampleName;
            if (exampleState.formatKey === "pdf" && chkCombinePdf.value) {
                exampleName = buildCombinedPdfStem(fileNameOptions, baseFileName, exampleDateStamp) + ".pdf";
            } else {
                var exampleScale = (exampleState.formatKey === "pdf") ? 100 : exampleState.scales[0];
                exampleName = buildFileStem(fileNameOptions, {
                    baseFileName: baseFileName,
                    artboardState: exampleState,
                    artboardCount: artboardStates.length,
                    scale: exampleScale,
                    dateStamp: exampleDateStamp
                }) + "." + getFormatInfo(exampleState.formatKey).extension;
            }
            if (chkCreateSubfolder.value) {
                var subfolderName = resolveSubfolderName(subfolderNameField.text, baseFileName, exampleDateStamp, fileNameOptions);
                if (subfolderName) exampleName = subfolderName + "/" + exampleName;
            }
            fileNameExampleText.text = labelValueText("fieldLabel.example", exampleName);
        }

        /* 選んでいる行すべてに設定を書き込む / Apply a change to every selected row */
        function applyToSelectedRows(applyChange) {
            var selectedIndexes = getSelectedIndexes();
            for (var k = 0; k < selectedIndexes.length; k++) {
                applyChange(artboardStates[selectedIndexes[k]]);
                refreshRow(selectedIndexes[k]);
            }
            updateFileNameExample();
            updateOKButton();
        }

        /* 書き出す行を番号の配列で決める（それ以外は書き出さない）/ Set the rows to export by index; others are turned off */
        function setExportTargets(targetIndexes) {
            var targetFlags = {};
            for (var k = 0; k < targetIndexes.length; k++) targetFlags["i" + targetIndexes[k]] = true;
            for (var m = 0; m < artboardStates.length; m++) artboardStates[m].enabled = !!targetFlags["i" + m];
            refreshAllRows();
            loadSelectedSettings();
        }

        /* 倍率欄を確定する。不正な値なら元に戻して false / Commit the scale field; revert and return false when invalid */
        function commitScaleField() {
            if (!scaleFieldEdited) return true;
            scaleFieldEdited = false;
            var scales = parseScales(scaleField.text);
            if (!scales) {
                alert(getLabel("alert.invalidScale"));
                loadSelectedSettings();
                return false;
            }
            applyToSelectedRows(function (artboardState) {
                artboardState.scales = scales.slice(0);
            });
            scaleField.text = scales.join(", ");
            return true;
        }

        /* 背景を選ぶ（色見本のクリックからも呼ぶ）/ Choose a background (also from swatch clicks) */
        function selectBackgroundChoice(choiceIndex) {
            for (var k = 0; k < backgroundChoices.length; k++) {
                backgroundChoices[k].radio.value = (k === choiceIndex);
            }
            var backgroundKey = backgroundChoices[choiceIndex].key;
            applyToSelectedRows(function (artboardState) {
                artboardState.backgroundKey = backgroundKey;
            });
        }

        for (var r = 0; r < formatRadios.length; r++) {
            formatRadios[r].onClick = (function (clickedIndex) {
                return function () {
                    lastChosenFormatKey = EXPORT_FORMATS[clickedIndex].key;
                    applyToSelectedRows(function (artboardState) {
                        artboardState.formatKey = lastChosenFormatKey;
                    });
                    loadSelectedSettings();
                };
            })(r);
        }
        for (var s = 0; s < backgroundChoices.length; s++) {
            backgroundChoices[s].radio.onClick = (function (choiceIndex) {
                return function () {
                    selectBackgroundChoice(choiceIndex);
                };
            })(s);
        }

        artboardList.onChange = loadSelectedSettings;
        artboardList.onDoubleClick = function () {
            var selectedIndexes = getSelectedIndexes();
            if (selectedIndexes.length === 0) return;
            var newEnabled = !artboardStates[selectedIndexes[0]].enabled;
            applyToSelectedRows(function (artboardState) {
                artboardState.enabled = newEnabled;
            });
            chkExportTarget.value = newEnabled;
        };
        chkExportTarget.onClick = function () {
            var newEnabled = chkExportTarget.value;
            applyToSelectedRows(function (artboardState) {
                artboardState.enabled = newEnabled;
            });
        };
        scaleField.onChanging = function () {
            scaleFieldEdited = true;
        };
        scaleField.onChange = commitScaleField;
        /* 品質は 0〜100 の整数に丸める。数値でなければ元に戻す / Round quality into 0-100; revert non-numbers */
        jpegQualityField.onChange = function () {
            var qualityText = String(jpegQualityField.text).replace(/[０-９]/g, function (fullWidthChar) {
                return String.fromCharCode(fullWidthChar.charCodeAt(0) - 0xFEE0);
            }).replace(/\s/g, "");
            var qualityValue = Number(qualityText);
            if (qualityText === "" || isNaN(qualityValue)) {
                loadSelectedSettings();
                return;
            }
            qualityValue = Math.max(0, Math.min(100, Math.round(qualityValue)));
            jpegQualityField.text = String(qualityValue);
            applyToSelectedRows(function (artboardState) {
                artboardState.jpegQuality = qualityValue;
            });
        };
        pdfPresetDropdown.onChange = function () {
            /* 欄の読み込みで選択を書き換えたときにも呼ばれるので、そのときは何もしない / Also fires when the fields are loaded; ignore that */
            if (isLoadingFields || !pdfPresetDropdown.selection) return;
            var presetName = pdfPresetNames[pdfPresetDropdown.selection.index];
            lastChosenPdfPreset = presetName;
            applyToSelectedRows(function (artboardState) {
                artboardState.pdfPreset = presetName;
            });
        };
        chkCombinePdf.onClick = updateFileNameExample;

        btnSelectAll.onClick = function () {
            artboardList.selection = allIndexes;
            loadSelectedSettings();
        };
        btnTargetAll.onClick = function () {
            setExportTargets(allIndexes);
        };
        btnTargetActive.onClick = function () {
            setExportTargets([documentInfo.activeArtboardIndex]);
        };
        /* 番号の指定を適用する / Apply the numbers typed */
        function applyTargetRange() {
            if (String(targetRangeField.text).replace(/\s/g, "") === "") return;
            var targetIndexes = parseArtboardRange(targetRangeField.text, artboardStates.length);
            if (!targetIndexes) {
                alert(getLabel("alert.invalidRange"));
                return;
            }
            setExportTargets(targetIndexes);
        }
        btnApplyRange.onClick = applyTargetRange;
        targetRangeField.onChange = applyTargetRange;

        chkExcludeArtboards.onClick = function () {
            updateFileNameControls();
            refreshAllRows();
        };
        excludeArtboardPrefixField.onChanging = refreshAllRows;
        chkExcludeLayers.onClick = updateFileNameControls;

        /* 今の保存先をパス欄に出す / Show the current destination */
        function updateDestinationPath() {
            var shownFolder = rbCustomFolder.value ? (customFolderPath ? new Folder(customFolderPath) : null) : sourceFolder;
            var shownPath = shownFolder ? decodeURI(shownFolder.fsName) : "";
            /* 表示はホームを ~ に略し、ツールチップにはフルパスを出す / Show ~ for the home folder; the tooltip keeps the full path */
            destinationPathText.text = abbreviateHomePath(shownPath);
            destinationPathText.helpTip = shownPath;
        }

        /* フォルダーを選ばせる。選ばなければ false / Ask for a folder; false when cancelled */
        function chooseCustomFolder() {
            var startFolder = customFolderPath ? new Folder(customFolderPath) : sourceFolder;
            var chosenFolder = (startFolder && startFolder.exists) ? startFolder.selectDlg(getLabel("alert.chooseFolder")) : Folder.selectDialog(getLabel("alert.chooseFolder"));
            if (!chosenFolder) return false;
            customFolderPath = chosenFolder.fsName;
            return true;
        }

        /* ファイル名の欄を変えたら、有効・無効と例を更新する / Update the enabled states and the example when a file name field changes */
        function onFileNameSettingChanged() {
            updateFileNameControls();
            updateFileNameExample();
        }
        chkDocumentName.onClick = onFileNameSettingChanged;
        documentNameField.onChanging = updateFileNameExample;
        artboardPatternDropdown.onChange = updateFileNameExample;
        suffixSeparatorDropdown.onChange = updateFileNameExample;
        chkDate.onClick = onFileNameSettingChanged;
        dateFormatDropdown.onChange = updateFileNameExample;
        chkSuffixText.onClick = onFileNameSettingChanged;
        suffixTextField.onChanging = updateFileNameExample;
        fileNameTemplateField.onChanging = updateFileNameExample;
        invalidReplacementDropdown.onChange = updateFileNameExample;
        chkReplaceSpaces.onClick = onFileNameSettingChanged;
        spaceReplacementDropdown.onChange = updateFileNameExample;
        chkCreateSubfolder.onClick = onFileNameSettingChanged;
        subfolderNameField.onChanging = updateFileNameExample;
        rbCombineParts.onClick = function () {
            rbFileNameTemplate.value = false;
            onFileNameSettingChanged();
        };
        rbFileNameTemplate.onClick = function () {
            rbCombineParts.value = false;
            onFileNameSettingChanged();
        };

        rbSameFolder.onClick = updateDestinationPath;
        rbCustomFolder.onClick = function () {
            /* まだフォルダーが無ければ選ばせ、選ばなければ元に戻す / Ask for a folder if none; revert when cancelled */
            if (!customFolderPath && !chooseCustomFolder() && sourceFolder) {
                rbCustomFolder.value = false;
                rbSameFolder.value = true;
            }
            updateDestinationPath();
        };
        btnChooseFolder.onClick = function () {
            if (!chooseCustomFolder()) return;
            rbSameFolder.value = false;
            rbCustomFolder.value = true;
            updateDestinationPath();
        };

        /* 入力中の値を確定し、書き出せるかを確かめてから閉じる / Commit typed values and check before closing */
        btnOK.onClick = function () {
            if (!commitScaleField()) return;
            var selectedIndexes = getSelectedIndexes();
            if (selectedIndexes.length > 0 && jpegQualityField.text !== String(artboardStates[selectedIndexes[0]].jpegQuality)) jpegQualityField.onChange();
            var fileNameOptions = getFileNameOptions();
            var isFileNameEmpty = fileNameOptions.useFileNameTemplate
                ? String(fileNameOptions.fileNameTemplate).replace(/\s/g, "") === ""
                : (!(fileNameOptions.useDocumentName && fileNameOptions.documentName) && ARTBOARD_NAME_PATTERNS[fileNameOptions.artboardPatternIndex] === "" && !fileNameOptions.useDate && !(fileNameOptions.useSuffixText && fileNameOptions.suffixText));
            if (isFileNameEmpty) {
                alert(getLabel("alert.emptyFileName"));
                return;
            }
            var tooLargeNames = findTooLargeJpegScales(artboardStates, getExclusionOptions());
            if (tooLargeNames.length > 0) {
                alert(getLabel("alert.jpegScaleTooLarge", [JPEG_MAX_SCALE, tooLargeNames.slice(0, OVERWRITE_LIST_MAX).join("\n")]));
                return;
            }
            if (rbCustomFolder.value && (!customFolderPath || !new Folder(customFolderPath).exists)) {
                alert(getLabel("alert.folderMissing"));
                if (!chooseCustomFolder()) return;
                updateDestinationPath();
            }
            exportDialog.close(1);
        };

        /* 初期状態は全行を選び、設定がまとめて変わるようにする / Start with every row selected so a change applies to all */
        for (var m = 0; m < artboardStates.length; m++) refreshRow(m);
        artboardList.selection = allIndexes;
        loadSelectedSettings();
        updateOKButton();
        updateDestinationPath();
        updateFileNameControls();
        updateFileNameExample();

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(exportDialog, SCRIPT_NAME);
        if (exportDialog.show() !== 1) {
            return null;
        }
        /* 最後に選んだ形式・プリセットを次回の初期値にする / The last format and preset chosen become the next initial values */
        if (lastChosenFormatKey) $.global[FORMAT_MEMORY_KEY] = lastChosenFormatKey;
        if (lastChosenPdfPreset) $.global[PDF_PRESET_MEMORY_KEY] = lastChosenPdfPreset;
        var chosenSettings = getFileNameOptions();
        chosenSettings.useCustomFolder = rbCustomFolder.value;
        chosenSettings.customFolderPath = customFolderPath;
        chosenSettings.createSubfolder = chkCreateSubfolder.value;
        chosenSettings.subfolderName = subfolderNameField.text;
        chosenSettings.revealAfterExport = chkRevealFolder.value;
        chosenSettings.excludeArtboards = chkExcludeArtboards.value;
        chosenSettings.excludeArtboardPrefix = excludeArtboardPrefixField.text;
        chosenSettings.excludeLayers = chkExcludeLayers.value;
        chosenSettings.excludeLayerPrefix = excludeLayerPrefixField.text;
        chosenSettings.combinePdf = chkCombinePdf.value;
        return chosenSettings;
    }

    // =========================================
    // 書き出しヘルパー / Export helpers
    // =========================================

    /**
     * アクティブなアートボードを PNG または JPEG で書き出す
     * @param {Document} sourceDoc - 書き出すドキュメント
     * @param {{formatKey: string, displayName: string, scale: number, backgroundKey: string, jpegQuality: number, file: File}} exportJob - 書き出しジョブ
     * @returns {void}
     */
    function exportArtboardAsImage(sourceDoc, exportJob) {
        var exportOptions;
        var exportType;
        if (exportJob.formatKey === "jpeg") {
            exportOptions = new ExportOptionsJPEG();
            exportOptions.qualitySetting = exportJob.jpegQuality;
            exportType = ExportType.JPEG;
        } else {
            exportOptions = new ExportOptionsPNG24();
            exportOptions.transparency = (exportJob.backgroundKey === "transparent");
            exportType = ExportType.PNG24;
        }
        /* 白・黒の背景はマット色で塗る / White and black backgrounds are filled with the matte color */
        if (exportJob.backgroundKey !== "transparent") {
            var matteValue = (exportJob.backgroundKey === "black") ? 0 : 255;
            var matteColor = new RGBColor();
            matteColor.red = matteValue;
            matteColor.green = matteValue;
            matteColor.blue = matteValue;
            exportOptions.matte = true;
            exportOptions.matteColor = matteColor;
        }
        exportOptions.artBoardClipping = true;
        exportOptions.antiAliasing = true;
        exportOptions.horizontalScale = exportJob.scale;
        exportOptions.verticalScale = exportJob.scale;

        /* 1枚の失敗で全体を止めない / One failed export does not stop the rest */
        try {
            sourceDoc.exportFile(exportJob.file, exportType, exportOptions);
        } catch (e) {
            alert(getLabel("alert.exportError", [exportJob.displayName]) + e.message);
        }
    }

    /**
     * 指定したアートボードを1つの PDF として保存する（まとめるときは複数ページ）
     * @param {Document} sourceDoc - 保存に使うドキュメント（元ファイルの複製）
     * @param {{artboardIndexes: number[], displayName: string, pdfPreset: string, file: File}} exportJob - 書き出しジョブ
     * @returns {void}
     */
    function saveArtboardsAsPdf(sourceDoc, exportJob) {
        var pageNumbers = [];
        for (var i = 0; i < exportJob.artboardIndexes.length; i++) pageNumbers.push(exportJob.artboardIndexes[i] + 1);
        var pdfOptions = new PDFSaveOptions();
        /* プリセットを先に読み込み、範囲などはそのあとで上書きする / Load the preset first, then override the range and viewing */
        if (exportJob.pdfPreset) pdfOptions.pDFPreset = exportJob.pdfPreset;
        pdfOptions.artboardRange = pageNumbers.join(",");
        pdfOptions.viewAfterSaving = false;

        /* 1枚の失敗で全体を止めない / One failed export does not stop the rest */
        try {
            sourceDoc.saveAs(exportJob.file, pdfOptions);
        } catch (e) {
            alert(getLabel("alert.exportError", [exportJob.displayName]) + e.message);
        }
    }

    /**
     * ドキュメントのファイルを一時フォルダーへ複製する
     * @param {Document} doc - 複製するドキュメント（保存済み）
     * @returns {File} 複製したファイル
     */
    function copyDocumentFileToTemp(doc) {
        var extensionMatch = doc.fullName.name.match(/\.[^.]+$/);
        var tempFile = new File(Folder.temp.fsName + "/" + SCRIPT_NAME + "-" + new Date().getTime() + (extensionMatch ? extensionMatch[0] : ".ai"));
        doc.fullName.copy(tempFile);
        return tempFile;
    }

    // =========================================
    // 進捗表示 / Progress UI
    // =========================================

    /**
     * 状況表示・プログレスバー・キャンセルボタンを持つ進捗パレットを開く
     * @param {number} totalJobs - ジョブの総数
     * @returns {{isCancelled: Function, update: Function, close: Function}} 進捗パレットの操作
     */
    function createProgressWindow(totalJobs) {
        var progressWin = new Window("palette", getLabel("dialog.progressTitle") + " " + SCRIPT_VERSION, undefined, { closeButton: false });
        setupWindow(progressWin);

        var statusText = progressWin.add("statictext", undefined, getLabel("status.preparing"));
        statusText.preferredSize.width = PROGRESS_WIDTH;

        var progressBar = progressWin.add("progressbar", undefined, 0, totalJobs);
        progressBar.preferredSize = [PROGRESS_WIDTH, PROGRESS_BAR_HEIGHT];

        /* キャンセルボタン（押下でフラグを立て、ループ側が中断）/ Cancel button (sets a flag that the export loop checks) */
        var cancelled = false;
        var buttonRow = addButtonRow(progressWin);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnCancel.onClick = function () {
            cancelled = true;
            btnCancel.enabled = false;
            statusText.text = getLabel("status.cancelling");
            progressWin.update();
        };
        alignRightOnlyButtonRow(buttonRow);

        progressWin.show();

        return {
            isCancelled: function () {
                return cancelled;
            },
            update: function (value, statusLabel) {
                statusText.text = statusLabel + "  (" + value + " / " + totalJobs + ")";
                progressBar.value = value;
                progressWin.update();
            },
            close: function () {
                progressWin.close();
            }
        };
    }

    exportArtboards();

})();
