#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

フォルダー内の Illustrator ファイル（.ai / .svg）を、指定したバージョン形式でまとめて保存します。
上書きモードとカスタム保存先を切り替えられ、対象の拡張子やパス表示の形式も指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AIBatchVersionSave.md

### Overview

Batch-saves the Illustrator files (.ai / .svg) in a folder in a chosen file-format version.
You can switch between overwrite mode and a custom destination, and choose the target extensions and how paths are displayed.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AIBatchVersionSave.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AIBatchVersionSave";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AIBatchVersionSave.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AIBatchVersionSave.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 保存オプション（保存バージョンと PDF 互換はダイアログの設定を使う）
       Save options (the version and PDF compatibility come from the dialog) */
    var SAVE_OPTS = {
        /* ファイル形式 / File format */
        compressed: true,                /* 圧縮を使用 / Use compression */

        /* 埋め込みリソース / Embedded resources */
        embedLinkedFiles: false,         /* 配置画像を埋め込む【7 or later】/ Embed linked images */
        embedICCProfile: false,          /* ICC プロファイルを埋め込む【10 or later】/ Embed ICC profile */
        fontSubsetThreshold: 100,        /* フォントサブセット閾値（％）/ Font subset threshold (%) */

        /* 出力処理 / Output rendering */
        /* PRESERVEPDFOVERPRINT（保持）/ DISCARDPDFOVERPRINT（破棄） */
        overprint: PDFOverprint.PRESERVEPDFOVERPRINT,
        /* PRESERVEAPPEARANCE（アピアランス保持）/ PRESERVEPATHS（パス保持） */
        flattenOutput: OutputFlattening.PRESERVEAPPEARANCE
    };

    /* Dropbox のローカルマウントパス接頭辞（環境に合わせて書き換える）/ Local Dropbox mount prefix — adjust to your env */
    var DROPBOX_PREFIX = "/Users/takano/sw Dropbox/takano masahiro/";

    /* ファイル名の区切り文字の候補と初期値 / Separator choices and the default one */
    var SEPARATOR_OPTIONS = ["-", "_"];
    var SEPARATOR_DEFAULT_INDEX = 1;

    // =========================================
    // 環境設定キー / Preference key
    // =========================================

    /* ［更新済み］をファイル名に付ける環境設定 / Preference that appends [Converted] to file names */
    var CONVERTED_PREF_KEY = "fileFormatGetFile/ConvertedInFilename";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 12;

    /* フォルダー指定のボタン幅とパス表示の幅 / Folder button width and path label width */
    var FOLDER_BUTTON_WIDTH = 90;
    var FOLDER_LABEL_WIDTH = 320;

    /* ファイル名に付ける文字列の入力欄（文字数） / Width of the suffix fields in characters */
    var SUFFIX_FIELD_CHARACTERS = 12;

    /* 進捗ウィンドウの幅 / Progress window width */
    var PROGRESS_WIDTH = 420;

    /**
     * パネルの共通レイアウトを設定する
     * @param {Panel} targetPanel - 設定するパネル
     * @param {number} [spacing] - 子の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function applyPanelLayout(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "バージョン指定で一括保存", en: "Batch Save As Version" }
        },
        prompt: {
            selectSourceFolder: {
                ja: "対象ファイル（.ai / .svg）の入っているフォルダを選択してください",
                en: "Select the folder containing target files (.ai / .svg)"
            },
            selectDestFolder: { ja: "保存するフォルダを選択してください", en: "Select the destination folder" }
        },
        panel: {
            folders: { ja: "フォルダー指定", en: "Folders" },
            targets: { ja: "対象", en: "Targets" },
            saveOptions: { ja: "保存設定", en: "Save Settings" },
            filename: { ja: "ファイル名の調整", en: "Filename Adjustments" }
        },
        radio: {
            overwrite: { ja: "上書きモード", en: "Overwrite mode" },
            customMode: { ja: "カスタム", en: "Custom" }
        },
        checkbox: {
            sameAsSource: { ja: "保存先を対象と同じにする", en: "Use same folder as source" },
            fullPath: { ja: "フルパス", en: "Full path" },
            dropbox: { ja: "Dropboxパスを短縮", en: "Shorten Dropbox path" },
            appendConverted: { ja: "Illustratorの［更新済み］付与を使用", en: "Use Illustrator's [Converted] suffix" },
            pdfCompatible: { ja: "PDF互換ファイルを作成", en: "Create PDF Compatible File" },
            filenameSuffix: { ja: "保存バージョンを付加", en: "Append save version" },
            customSuffix: { ja: "任意の文字列を付加", en: "Append custom text" }
        },
        fieldLabel: {
            saveVersion: { ja: "保存バージョン", en: "Save version" },
            separator: { ja: "区切り文字", en: "Separator" }
        },
        button: {
            setSource: { ja: "対象...", en: "Source..." },
            setDestination: { ja: "保存先...", en: "Destination..." },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        status: {
            sourceNotSet: { ja: "未指定", en: "Not set" },
            destinationNotSet: { ja: "未指定", en: "Not set" },
            destinationSameAsSource: { ja: "〈対象と同じ〉", en: "〈Same as source〉" },
            processing: { ja: "処理中", en: "Processing" }
        },
        tooltip: {
            sourceFolder: {
                ja: "処理対象のファイル（.ai / .svg）が入っているフォルダーを指定します。未指定の場合はOK後に選択します。",
                en: "Selects the folder containing target files (.ai / .svg). If not set, you will be asked after clicking OK."
            },
            destinationFolder: {
                ja: "別名保存先のフォルダーを指定します。未指定の場合は対象フォルダー選択後に確認します。",
                en: "Selects the destination folder for saved copies. If not set, you will be asked after choosing the source folder."
            },
            sameAsSource: {
                ja: "保存先フォルダーを指定せず、対象と同じフォルダーへ保存します。ファイル名にサフィックスや任意文字列を付けて別名保存できます。",
                en: "Save into the source folder without picking a separate destination. Add a suffix or custom text to keep filenames distinct."
            },
            fullPath: {
                ja: "ONにすると ~ 短縮を無効にしてフルパス表示。Dropboxパス短縮がONの場合は無効。",
                en: "When ON, disables ~ abbreviation and shows the full path. Ignored when Dropbox shortening is ON."
            },
            dropbox: { ja: "Dropbox ローカルマウントの接頭辞を取り除いて表示します。", en: "Strips the local Dropbox mount prefix from the displayed path." },
            targetAi: { ja: ".ai ファイルを処理対象に含めます。", en: "Include .ai files in the batch." },
            targetSvg: {
                ja: ".svg ファイルを処理対象に含めます。開いた後、選択した Illustrator バージョンの .ai として保存します。",
                en: "Include .svg files in the batch. Opened and saved as .ai using the selected Illustrator version."
            },
            saveVersion: { ja: "実行中のIllustratorで保存できるバージョンのみ表示します。", en: "Lists only versions supported by the running Illustrator." },
            appendConverted: {
                ja: "Illustrator標準の［更新済み］付与設定を一時的に切り替えます。処理後は元の設定に戻します。",
                en: "Temporarily changes Illustrator's built-in [Converted] filename setting, then restores it after processing."
            },
            pdfCompatible: {
                ja: "PDF互換ファイルを作成します。互換性は上がりますが、ファイルサイズが大きくなる場合があります。",
                en: "Creates a PDF-compatible file. This improves compatibility but may increase file size."
            },
            overwrite: {
                ja: "対象フォルダー内のファイルを、同じファイル名で .ai として保存します（.ai 入力は上書き、.svg 入力は同名の .ai を生成）。保存先フォルダーやファイル名への付加設定は無効になります。",
                en: "Saves files in the source folder as .ai using the same base name (overwrites .ai inputs; .svg inputs generate a same-named .ai). Destination folder and filename append options are disabled."
            },
            customMode: {
                ja: "保存先フォルダーとファイル名の調整を指定して .ai として保存します。",
                en: "Saves as .ai using the destination folder and filename settings below."
            },
            filenameSuffix: {
                ja: "選択した保存バージョンに対応する文字列をファイル名の末尾に付加します。",
                en: "Appends text matching the selected save version to the end of the filename."
            },
            customSuffix: { ja: "任意の文字列を、保存バージョンの文字列の後ろに追加します。", en: "Adds custom text after the save-version suffix." },
            separator: {
                ja: "元のファイル名内のスペース、および付加文字列との区切りに使う文字です。",
                en: "Used to replace spaces in the original filename and to separate appended text."
            }
        },
        alert: {
            processingComplete: { ja: "処理が完了しました。スクリプトを終了します。", en: "Processing complete. Exiting script." },
            noFilesFound: { ja: "対象のファイルが見つかりませんでした。", en: "No matching files were found." },
            noTargetSelected: { ja: "処理対象が選択されていません。", en: "No target file types are selected." },
            noVersionsAvailable: {
                ja: "このIllustratorでは選択可能な保存バージョンがありません。",
                en: "No save-version targets are available in this Illustrator."
            },
            confirmSameFolder: {
                ja: "対象フォルダーと保存先フォルダーが同じで、ファイル名への追加もありません。\n同名で上書きされます（明示的に上書きしたい場合は「上書きモード」をご利用ください）。続行しますか？",
                en: "Source and destination folders are the same, and no filename suffix is set.\nFiles will be overwritten with the same name (use the “Overwrite” mode for explicit overwriting). Continue?"
            },
            confirmOverwriteSvg: {
                ja: "上書きモードで .svg を対象にしています。\n.svg ファイル自体は上書きされませんが、同じフォルダー内に同名の .ai がある場合は上書きされます。続行しますか？",
                en: "Overwrite mode includes .svg files.\nThe .svg files themselves will not be overwritten, but same-named .ai files in the source folder will be overwritten. Continue?"
            }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.folders" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 複数のコントロールに同じ tooltip を設定する
     * @param {Object[]} controls - 設定するコントロール
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @returns {void}
     */
    function setHelpTips(controls, tooltipPath) {
        for (var i = 0; i < controls.length; i++) {
            controls[i].helpTip = getLabel(tooltipPath);
        }
    }

    // =========================================
    // ファイルとパス / Files and paths
    // =========================================

    /**
     * Illustrator ファイル（.ai）か
     * @param {File|Folder} fileEntry - 調べるファイル
     * @returns {boolean} .ai ファイルなら true
     */
    function isIllustratorFile(fileEntry) {
        return (fileEntry instanceof File) && (/\.ai$/i).test(fileEntry.name);
    }

    /**
     * SVG ファイルか
     * @param {File|Folder} fileEntry - 調べるファイル
     * @returns {boolean} .svg ファイルなら true
     */
    function isSvgFile(fileEntry) {
        return (fileEntry instanceof File) && (/\.svg$/i).test(fileEntry.name);
    }

    /**
     * 選んだ種類（.ai / .svg）のファイルだけを通すフィルターを作る
     * @param {boolean} includeAi - .ai を含めるか
     * @param {boolean} includeSvg - .svg を含めるか
     * @returns {Function} Folder.getFiles() に渡すフィルター
     */
    function buildTargetFileFilter(includeAi, includeSvg) {
        return function (fileEntry) {
            if (includeAi && isIllustratorFile(fileEntry)) return true;
            if (includeSvg && isSvgFile(fileEntry)) return true;
            return false;
        };
    }

    /**
     * ファイル一覧に指定の種類が含まれるか
     * @param {File[]} fileList - ファイル一覧
     * @param {Function} matchesType - isIllustratorFile / isSvgFile
     * @returns {boolean} 1つでもあれば true
     */
    function containsFileType(fileList, matchesType) {
        for (var i = 0; i < fileList.length; i++) {
            if (matchesType(fileList[i])) return true;
        }
        return false;
    }

    /**
     * フォルダー内の .ai / .svg ファイル数を数える
     * @param {Folder} folder - 調べるフォルダー
     * @returns {{ai: number, svg: number}} 種類ごとのファイル数
     */
    function countTargetFiles(folder) {
        var fileCounts = { ai: 0, svg: 0 };
        /* フォルダーの読み取りに失敗したら 0 件 / count nothing when the folder cannot be read */
        try {
            var folderFiles = folder.getFiles(function (fileEntry) { return fileEntry instanceof File; });
            if (!folderFiles) return fileCounts;
            for (var i = 0; i < folderFiles.length; i++) {
                var fileName = folderFiles[i].name;
                if ((/\.ai$/i).test(fileName)) fileCounts.ai++;
                else if ((/\.svg$/i).test(fileName)) fileCounts.svg++;
            }
        } catch (countError) { }
        return fileCounts;
    }

    /**
     * URI エンコードされたファイル名をデコードする
     * @param {string} rawName - File.name の値
     * @returns {string} デコードした名前（失敗時は元の文字列）
     */
    function decodeFileName(rawName) {
        /* 不正なエスケープは URIError / malformed escapes throw URIError */
        try {
            return decodeURI(rawName);
        } catch (decodeError) {
            return rawName;
        }
    }

    /**
     * 区切り文字をはさんでファイル名の部品をつなぐ（空の部品は足さない）
     * @param {string} currentName - ここまでの名前
     * @param {string} namePart - 足す部品
     * @param {string} separator - 区切り文字
     * @returns {string} つないだ名前
     */
    function appendNamePart(currentName, namePart, separator) {
        if (namePart === "") return currentName;
        if (currentName === "") return namePart;
        return currentName + separator + namePart;
    }

    /**
     * ホームディレクトリ部分を ~ に短縮する
     * @param {string} fsPath - パス
     * @returns {string} 短縮したパス
     */
    function toTildePath(fsPath) {
        /* ホームフォルダーが取れなければそのまま / keep the path when the home folder cannot be resolved */
        try {
            var homePath = Folder("~").fsName;
            if (homePath && homePath.length > 0) {
                if (fsPath === homePath) return "~";
                if (fsPath.indexOf(homePath + "/") === 0) return "~" + fsPath.substr(homePath.length);
                if (fsPath.indexOf(homePath + "\\") === 0) return "~" + fsPath.substr(homePath.length);
            }
        } catch (homeError) { }
        return fsPath;
    }

    /**
     * 表示用にパスを整形する（Dropbox 接頭辞の除去が優先、次に ~ 短縮）
     * @param {string} fsPath - パス
     * @param {boolean} useTilde - ~ に短縮するか
     * @param {boolean} useDropbox - Dropbox 接頭辞を取り除くか
     * @returns {string} 表示用のパス
     */
    function formatDisplayPath(fsPath, useTilde, useDropbox) {
        if (!fsPath) return fsPath;
        if (useDropbox && DROPBOX_PREFIX && fsPath.indexOf(DROPBOX_PREFIX) === 0) {
            return fsPath.substr(DROPBOX_PREFIX.length);
        }
        if (useTilde) return toTildePath(fsPath);
        return fsPath;
    }

    // =========================================
    // バージョン定義 / Version Options
    // =========================================

    /* ドロップダウンの並び順＝ここでの定義順。デフォルトは isDefault:true の項目。
       実行中のIllustratorが対応していないバージョンは自動的に除外される。
       Dropdown order follows this list; the isDefault item is preselected. Versions the running Illustrator lacks are skipped. */
    var VERSION_OPTIONS = [];

    /**
     * 実行中の Illustrator に Compatibility の値があるバージョンだけを候補に加える
     * @param {string} label - ドロップダウンの表示名
     * @param {string} suffix - ファイル名に付ける文字列
     * @param {string} enumKey - Compatibility のキー
     * @param {boolean} [isDefault] - 初期選択にするか
     * @returns {void}
     */
    function addVersionOption(label, suffix, enumKey, isDefault) {
        /* 古い版の列挙値が無い環境もある / older enum values may be missing */
        try {
            var compatibilityValue = Compatibility[enumKey];
            if (compatibilityValue == null) return;
            VERSION_OPTIONS.push({
                label: label,
                suffix: suffix,
                compatibility: compatibilityValue,
                isDefault: (isDefault === true)
            });
        } catch (e) { }
    }

    /* Illustrator CC (v17) 以降は内部ファイル形式が変わっていないため v17 に集約 / CC (v17) onward share the same .ai format */
    addVersionOption("Illustrator CC以降 (v17)", "v17", "ILLUSTRATOR17", true);
    addVersionOption("Illustrator CS6 (v16)", "v16", "ILLUSTRATOR16");
    addVersionOption("Illustrator CS5 (v15)", "v15", "ILLUSTRATOR15");
    addVersionOption("Illustrator CS4 (v14)", "v14", "ILLUSTRATOR14");
    addVersionOption("Illustrator CS3 (v13)", "v13", "ILLUSTRATOR13");
    addVersionOption("Illustrator CS2 (v12)", "v12", "ILLUSTRATOR12");
    addVersionOption("Illustrator CS (v11)", "v11", "ILLUSTRATOR11");
    addVersionOption("Illustrator 10 (v10)", "v10", "ILLUSTRATOR10");
    addVersionOption("Illustrator 9 (v9)", "v9", "ILLUSTRATOR9");
    addVersionOption("Illustrator 8 (v8)", "v8", "ILLUSTRATOR8");

    /**
     * 初期選択の版が除外されていたら、先頭の版を初期選択にする
     * @returns {void}
     */
    function ensureDefaultVersion() {
        if (VERSION_OPTIONS.length === 0) return;
        for (var i = 0; i < VERSION_OPTIONS.length; i++) {
            if (VERSION_OPTIONS[i].isDefault) return;
        }
        VERSION_OPTIONS[0].isDefault = true;
    }
    ensureDefaultVersion();

    // =========================================
    // 保存 / Saving
    // =========================================

    /**
     * 保存先のファイルを決める
     * @param {File} sourceFile - 元のファイル
     * @param {Folder} destinationFolder - 保存先フォルダー
     * @param {Object} batchSettings - ダイアログの設定
     * @returns {File} 保存先のファイル
     */
    function getOutputFile(sourceFile, destinationFolder, batchSettings) {
        /* File.name はURIエンコード文字列なので先にデコード / File.name is URI-encoded — decode first */
        var decodedName = decodeFileName(sourceFile.name);
        var dotIndex = decodedName.lastIndexOf(".");
        var baseName = (dotIndex >= 0) ? decodedName.substr(0, dotIndex) : decodedName;

        if (batchSettings.overwrite) {
            /* 上書きモード：ファイル名は加工せず、拡張子のみ小文字 .ai に統一 / In overwrite mode, keep filename but force lowercase .ai */
            return new File(sourceFile.parent.fsName + "/" + baseName + ".ai");
        }

        var separator = batchSettings.separator;
        /* 半角/全角スペース・タブを区切り文字に変換 / Convert half/full-width spaces and tabs to the chosen separator */
        var outputName = baseName.replace(/[ \t　]+/g, separator);
        /* 末尾の区切り文字を除去して二重連結を防止 / Strip trailing separator to avoid double joins */
        while (outputName.length > 0 && outputName.charAt(outputName.length - 1) === separator) {
            outputName = outputName.substr(0, outputName.length - 1);
        }
        outputName = appendNamePart(outputName, batchSettings.suffix, separator);
        outputName = appendNamePart(outputName, batchSettings.customSuffix, separator);
        return new File(destinationFolder.fsName + "/" + outputName + ".ai");
    }

    /**
     * 保存オプションを作る
     * @param {Object} batchSettings - ダイアログの設定
     * @returns {IllustratorSaveOptions} 保存オプション
     */
    function createSaveOptions(batchSettings) {
        var saveOptions = new IllustratorSaveOptions();

        /* 保存バージョン・PDF互換はダイアログ設定から / Version & PDF compatibility from dialog */
        saveOptions.compatibility = batchSettings.compatibility;
        saveOptions.pdfCompatible = batchSettings.pdfCompatible;

        /* それ以外は SAVE_OPTS を参照 / Others from SAVE_OPTS */
        saveOptions.compressed = SAVE_OPTS.compressed;
        saveOptions.embedLinkedFiles = SAVE_OPTS.embedLinkedFiles;
        saveOptions.embedICCProfile = SAVE_OPTS.embedICCProfile;
        saveOptions.fontSubsetThreshold = SAVE_OPTS.fontSubsetThreshold;
        saveOptions.overprint = SAVE_OPTS.overprint;
        saveOptions.flattenOutput = SAVE_OPTS.flattenOutput;

        return saveOptions;
    }

    /**
     * 1ファイルを開いて .ai で保存し、閉じる
     * @param {File} sourceFile - 元のファイル
     * @param {Folder} destinationFolder - 保存先フォルダー
     * @param {Object} batchSettings - ダイアログの設定
     * @returns {void}
     */
    function saveOneFile(sourceFile, destinationFolder, batchSettings) {
        var sourceDocument = null;

        try {
            sourceDocument = app.open(sourceFile);
            sourceDocument.saveAs(getOutputFile(sourceFile, destinationFolder, batchSettings), createSaveOptions(batchSettings));
        } finally {
            /* 保存中にエラーが出ても、開いたドキュメントを閉じる / Close opened document even if saving fails */
            if (sourceDocument != null) {
                try {
                    sourceDocument.close(SaveOptions.DONOTSAVECHANGES);
                } catch (closeError) { }
            }
        }
    }

    /**
     * 進捗ウィンドウを作って表示する
     * @param {number} totalCount - ファイルの総数
     * @returns {{update: Function, close: Function}} 進捗の更新と閉じる操作
     */
    function createProgressWindow(totalCount) {
        var progressWindow = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        progressWindow.alignChildren = "fill";
        progressWindow.margins = 16;
        progressWindow.spacing = 8;

        var statusText = progressWindow.add("statictext", undefined, "");
        statusText.preferredSize.width = PROGRESS_WIDTH;

        var progressBar = progressWindow.add("progressbar", undefined, 0, totalCount);
        progressBar.preferredSize.width = PROGRESS_WIDTH;
        progressBar.preferredSize.height = 8;

        progressWindow.show();

        return {
            update: function (currentIndex, filename) {
                progressBar.value = currentIndex;
                statusText.text = getLabel("status.processing") + " (" + currentIndex + "/" + totalCount + "): " + filename;
                progressWindow.update();
            },
            close: function () {
                /* 既に閉じていても続行 / ignore when already closed */
                try { progressWindow.close(); } catch (closeError) { }
            }
        };
    }

    /**
     * ［更新済み］の設定と警告の抑制を切り替えて、全ファイルを保存する（終わったら元に戻す）
     * @param {File[]} targetFileList - 保存するファイル
     * @param {Folder} destinationFolder - 保存先フォルダー
     * @param {Object} batchSettings - ダイアログの設定
     * @returns {void}
     */
    function runBatchSave(targetFileList, destinationFolder, batchSettings) {
        /* ［更新済み］付与の環境設定を一時的に変更し、処理後に元へ戻す / Toggle [Converted] pref then restore */
        var originalConvertedPref;
        try {
            originalConvertedPref = app.preferences.getBooleanPreference(CONVERTED_PREF_KEY);
        } catch (e) {
            originalConvertedPref = false;
        }
        app.preferences.setBooleanPreference(CONVERTED_PREF_KEY, batchSettings.appendConverted);

        /* Illustrator の各種警告ダイアログを抑制（カラープロファイル、フォント置換、上書き確認など）/ Suppress Illustrator interaction dialogs during batch */
        var originalInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var progressUI = null;
        try {
            progressUI = createProgressWindow(targetFileList.length);
            for (var i = 0; i < targetFileList.length; i++) {
                progressUI.update(i + 1, decodeFileName(targetFileList[i].name));
                saveOneFile(targetFileList[i], destinationFolder, batchSettings);
            }
        } finally {
            /* 中断されても進捗ウィンドウを閉じ、環境設定とユーザー操作レベルを元に戻す / Close progress window, restore preference and interaction level even if interrupted */
            if (progressUI != null) progressUI.close();
            app.preferences.setBooleanPreference(CONVERTED_PREF_KEY, originalConvertedPref);
            app.userInteractionLevel = originalInteractionLevel;
        }
        /* 完了通知は元に戻してから / Notify after restoring */
        alert(getLabel("alert.processingComplete"));
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを表示する
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showSettingsDialog() {
        var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        settingsDialog.alignChildren = "fill";
        settingsDialog.margins = 16;
        settingsDialog.spacing = 12;

        var dialogUi = {};
        buildModeControls(settingsDialog, dialogUi);
        buildFolderPickerControls(settingsDialog, dialogUi);
        buildTargetPanel(settingsDialog, dialogUi);
        buildSaveOptionsPanel(settingsDialog, dialogUi);
        buildFilenamePanel(settingsDialog, dialogUi);
        buildDialogButtons(settingsDialog);
        bindDialogEvents(dialogUi);

        if (settingsDialog.show() !== 1) return null;
        return readDialogSettings(dialogUi);
    }

    /**
     * モード切り替え（ダイアログ最上部、中央寄せ）を作る
     * @param {Window} settingsDialog - ダイアログ
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function buildModeControls(settingsDialog, dialogUi) {
        var modeGroup = settingsDialog.add("group");
        modeGroup.orientation = "row";
        modeGroup.alignChildren = ["center", "center"];
        modeGroup.alignment = ["center", "top"];
        dialogUi.overwriteRadio = modeGroup.add("radiobutton", undefined, getLabel("radio.overwrite"));
        dialogUi.customRadio = modeGroup.add("radiobutton", undefined, getLabel("radio.customMode"));
        dialogUi.customRadio.value = true;
        dialogUi.overwriteRadio.value = false;
        dialogUi.overwriteRadio.helpTip = getLabel("tooltip.overwrite");
        dialogUi.customRadio.helpTip = getLabel("tooltip.customMode");
    }

    /**
     * フォルダー指定ボタンとパス表示の1行を作る
     * @param {Panel} folderPanel - 追加先
     * @param {string} buttonLabelPath - ボタンの LABELS パス
     * @param {string} pathLabelPath - パス表示の初期文字の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @returns {{button: Button, pathLabel: StaticText}} ボタンとパス表示
     */
    function addFolderPickerRow(folderPanel, buttonLabelPath, pathLabelPath, tooltipPath) {
        var pickerRow = folderPanel.add("group");
        var pickerButton = pickerRow.add("button", undefined, getLabel(buttonLabelPath));
        pickerButton.preferredSize.width = FOLDER_BUTTON_WIDTH;
        var pathLabel = pickerRow.add("statictext", undefined, getLabel(pathLabelPath));
        pathLabel.preferredSize.width = FOLDER_LABEL_WIDTH;
        setHelpTips([pickerButton, pathLabel], tooltipPath);
        return { button: pickerButton, pathLabel: pathLabel };
    }

    /**
     * フォルダー指定パネルを作る
     * @param {Window} settingsDialog - ダイアログ
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function buildFolderPickerControls(settingsDialog, dialogUi) {
        var folderPanel = settingsDialog.add("panel", undefined, getLabel("panel.folders"));
        applyPanelLayout(folderPanel, 6);

        var sourceRow = addFolderPickerRow(folderPanel, "button.setSource", "status.sourceNotSet", "tooltip.sourceFolder");
        dialogUi.sourceButton = sourceRow.button;
        dialogUi.sourcePathLabel = sourceRow.pathLabel;
        dialogUi.pickedSourceFolder = null;

        var destinationRow = addFolderPickerRow(folderPanel, "button.setDestination", "status.destinationSameAsSource", "tooltip.destinationFolder");
        dialogUi.destinationButton = destinationRow.button;
        dialogUi.destinationPathLabel = destinationRow.pathLabel;
        dialogUi.pickedDestinationFolder = null;

        /* 保存先＝対象フォルダー切替 / Same-as-source toggle */
        var sameAsSourceGroup = folderPanel.add("group");
        sameAsSourceGroup.orientation = "row";
        sameAsSourceGroup.alignChildren = ["left", "center"];
        dialogUi.sameAsSourceCheckbox = sameAsSourceGroup.add("checkbox", undefined, getLabel("checkbox.sameAsSource"));
        dialogUi.sameAsSourceCheckbox.value = true;
        dialogUi.sameAsSourceCheckbox.helpTip = getLabel("tooltip.sameAsSource");

        /* パス表示オプション / Path display options */
        var pathOptionGroup = folderPanel.add("group");
        pathOptionGroup.orientation = "row";
        pathOptionGroup.alignChildren = ["left", "center"];
        pathOptionGroup.margins = [0, 10, 0, 0];
        dialogUi.fullPathCheckbox = pathOptionGroup.add("checkbox", undefined, getLabel("checkbox.fullPath"));
        dialogUi.fullPathCheckbox.value = false;
        dialogUi.dropboxCheckbox = pathOptionGroup.add("checkbox", undefined, getLabel("checkbox.dropbox"));
        dialogUi.dropboxCheckbox.value = true;
        dialogUi.fullPathCheckbox.helpTip = getLabel("tooltip.fullPath");
        dialogUi.dropboxCheckbox.helpTip = getLabel("tooltip.dropbox");
        /* Dropbox 短縮 ON のときフルパスは意味がないので無効化 / Disable full-path toggle when Dropbox shortening overrides */
        dialogUi.fullPathCheckbox.enabled = !dialogUi.dropboxCheckbox.value;
    }

    /**
     * 対象ファイルの種類（.ai / .svg）のパネルを作る
     * @param {Window} settingsDialog - ダイアログ
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function buildTargetPanel(settingsDialog, dialogUi) {
        var targetPanel = settingsDialog.add("panel", undefined, getLabel("panel.targets"));
        applyPanelLayout(targetPanel, 6);

        var targetTypeGroup = targetPanel.add("group");
        targetTypeGroup.orientation = "row";
        targetTypeGroup.alignChildren = ["left", "center"];

        dialogUi.aiCheckbox = targetTypeGroup.add("checkbox", undefined, ".ai");
        dialogUi.aiCheckbox.value = true;
        dialogUi.aiCheckbox.helpTip = getLabel("tooltip.targetAi");

        dialogUi.svgCheckbox = targetTypeGroup.add("checkbox", undefined, ".svg");
        dialogUi.svgCheckbox.value = true;
        dialogUi.svgCheckbox.helpTip = getLabel("tooltip.targetSvg");
    }

    /**
     * 保存設定パネル（保存バージョン・［更新済み］・PDF互換）を作る
     * @param {Window} settingsDialog - ダイアログ
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function buildSaveOptionsPanel(settingsDialog, dialogUi) {
        var saveOptionsPanel = settingsDialog.add("panel", undefined, getLabel("panel.saveOptions"));
        applyPanelLayout(saveOptionsPanel, 6);

        var versionGroup = saveOptionsPanel.add("group");
        dialogUi.versionLabel = versionGroup.add("statictext", undefined, labelText("fieldLabel.saveVersion"));
        dialogUi.versionDropdown = versionGroup.add("dropdownlist");
        setHelpTips([dialogUi.versionLabel, dialogUi.versionDropdown], "tooltip.saveVersion");

        dialogUi.defaultVersionIndex = 0;
        for (var i = 0; i < VERSION_OPTIONS.length; i++) {
            dialogUi.versionDropdown.add("item", VERSION_OPTIONS[i].label);
            if (VERSION_OPTIONS[i].isDefault) dialogUi.defaultVersionIndex = i;
        }
        dialogUi.versionDropdown.selection = dialogUi.defaultVersionIndex;

        dialogUi.convertedCheckbox = saveOptionsPanel.add("checkbox", undefined, getLabel("checkbox.appendConverted"));
        dialogUi.convertedCheckbox.value = true;
        dialogUi.convertedCheckbox.helpTip = getLabel("tooltip.appendConverted");

        dialogUi.pdfCheckbox = saveOptionsPanel.add("checkbox", undefined, getLabel("checkbox.pdfCompatible"));
        dialogUi.pdfCheckbox.value = true;
        dialogUi.pdfCheckbox.helpTip = getLabel("tooltip.pdfCompatible");
    }

    /**
     * チェックボックスと入力欄の1行（ファイル名に付ける文字列）を作る
     * @param {Panel} filenamePanel - 追加先
     * @param {string} checkboxLabelPath - チェックボックスの LABELS パス
     * @param {string} initialText - 入力欄の初期値
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @returns {{checkbox: Checkbox, input: EditText}} チェックボックスと入力欄
     */
    function addSuffixRow(filenamePanel, checkboxLabelPath, initialText, tooltipPath) {
        var suffixRow = filenamePanel.add("group");
        var suffixCheckbox = suffixRow.add("checkbox", undefined, labelText(checkboxLabelPath));
        var suffixInput = suffixRow.add("edittext", undefined, initialText);
        suffixInput.characters = SUFFIX_FIELD_CHARACTERS;
        setHelpTips([suffixCheckbox, suffixInput], tooltipPath);
        return { checkbox: suffixCheckbox, input: suffixInput };
    }

    /**
     * ファイル名の調整パネル（保存バージョン・任意の文字列・区切り文字）を作る
     * @param {Window} settingsDialog - ダイアログ
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function buildFilenamePanel(settingsDialog, dialogUi) {
        var filenamePanel = settingsDialog.add("panel", undefined, getLabel("panel.filename"));
        applyPanelLayout(filenamePanel, 6);

        var versionSuffixRow = addSuffixRow(filenamePanel, "checkbox.filenameSuffix", VERSION_OPTIONS[dialogUi.defaultVersionIndex].suffix, "tooltip.filenameSuffix");
        dialogUi.suffixCheckbox = versionSuffixRow.checkbox;
        dialogUi.suffixCheckbox.value = true;
        dialogUi.suffixInput = versionSuffixRow.input;

        var customSuffixRow = addSuffixRow(filenamePanel, "checkbox.customSuffix", "", "tooltip.customSuffix");
        dialogUi.customCheckbox = customSuffixRow.checkbox;
        dialogUi.customCheckbox.value = false;
        dialogUi.customInput = customSuffixRow.input;
        dialogUi.customInput.enabled = dialogUi.customCheckbox.value;

        var separatorGroup = filenamePanel.add("group");
        dialogUi.separatorLabel = separatorGroup.add("statictext", undefined, labelText("fieldLabel.separator"));
        dialogUi.separatorLabel.helpTip = getLabel("tooltip.separator");

        dialogUi.separatorRadios = [];
        for (var i = 0; i < SEPARATOR_OPTIONS.length; i++) {
            var separatorRadio = separatorGroup.add("radiobutton", undefined, SEPARATOR_OPTIONS[i]);
            separatorRadio.value = (i === SEPARATOR_DEFAULT_INDEX);
            separatorRadio.helpTip = getLabel("tooltip.separator");
            dialogUi.separatorRadios.push(separatorRadio);
        }
    }

    /**
     * OK／キャンセルボタンを作る
     * @param {Window} settingsDialog - ダイアログ
     * @returns {void}
     */
    function buildDialogButtons(settingsDialog) {
        var btnRowGroup = settingsDialog.add("group");
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, "OK", { name: "ok" });
    }

    /**
     * フォルダー選択ダイアログを開き、選ばれたフォルダーを返す
     * @param {string} promptPath - 案内文の LABELS パス
     * @returns {Folder|null} 選ばれたフォルダー（キャンセル時は null）
     */
    function selectFolder(promptPath) {
        return Folder.selectDialog(getLabel(promptPath));
    }

    /**
     * ダイアログのイベントを設定する
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function bindDialogEvents(dialogUi) {
        dialogUi.sourceButton.onClick = function () {
            var selectedFolder = selectFolder("prompt.selectSourceFolder");
            if (selectedFolder != null) {
                dialogUi.pickedSourceFolder = selectedFolder;
                /* ツールチップは常にフルパスで確認可能 / Tooltip always shows the full path */
                dialogUi.sourcePathLabel.helpTip = selectedFolder.fsName;
                refreshPathDisplays(dialogUi);
                updateTargetAvailability(dialogUi);
            }
        };

        dialogUi.destinationButton.onClick = function () {
            var selectedFolder = selectFolder("prompt.selectDestFolder");
            if (selectedFolder != null) {
                dialogUi.pickedDestinationFolder = selectedFolder;
                dialogUi.destinationPathLabel.helpTip = selectedFolder.fsName;
                refreshPathDisplays(dialogUi);
            }
        };

        dialogUi.fullPathCheckbox.onClick = function () {
            refreshPathDisplays(dialogUi);
        };
        dialogUi.dropboxCheckbox.onClick = function () {
            /* Dropbox 短縮 ON のときフルパスは意味がないので disable / Full-path toggle has no effect when Dropbox shortening is ON */
            dialogUi.fullPathCheckbox.enabled = !dialogUi.dropboxCheckbox.value;
            refreshPathDisplays(dialogUi);
        };

        dialogUi.suffixCheckbox.onClick = function () {
            dialogUi.suffixInput.enabled = dialogUi.suffixCheckbox.value;
        };

        dialogUi.versionDropdown.onChange = function () {
            dialogUi.suffixInput.text = VERSION_OPTIONS[dialogUi.versionDropdown.selection.index].suffix;
        };

        dialogUi.customCheckbox.onClick = function () {
            dialogUi.customInput.enabled = dialogUi.customCheckbox.value;
        };

        var onModeChange = function () {
            updateOverwriteState(dialogUi);
        };
        dialogUi.overwriteRadio.onClick = onModeChange;
        dialogUi.customRadio.onClick = onModeChange;
        dialogUi.sameAsSourceCheckbox.onClick = onModeChange;
        updateOverwriteState(dialogUi);
    }

    /**
     * 対象の種類のチェックを、フォルダー内のファイルの有無に合わせる（無い種類は無効＋OFF）
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function updateTargetAvailability(dialogUi) {
        var hasAi = true;
        var hasSvg = true;
        /* フォルダー未選択時は両方を初期状態（有効＋ON）に戻す / Reset to enabled + ON when no folder is picked */
        if (dialogUi.pickedSourceFolder) {
            var fileCounts = countTargetFiles(dialogUi.pickedSourceFolder);
            hasAi = fileCounts.ai > 0;
            hasSvg = fileCounts.svg > 0;
        }
        dialogUi.aiCheckbox.enabled = hasAi;
        dialogUi.svgCheckbox.enabled = hasSvg;
        /* ファイル有無に合わせて value も同期（有ればON、無ければOFF）/ Sync value to availability */
        dialogUi.aiCheckbox.value = hasAi;
        dialogUi.svgCheckbox.value = hasSvg;
    }

    /**
     * フルパス／Dropbox 短縮の状態に合わせて、フォルダーの表示用パスを返す
     * @param {Object} dialogUi - コントロールの控え
     * @param {Folder} folder - 表示するフォルダー
     * @returns {string} 表示用のパス
     */
    function getFolderDisplayPath(dialogUi, folder) {
        return formatDisplayPath(folder.fsName, !dialogUi.fullPathCheckbox.value, dialogUi.dropboxCheckbox.value);
    }

    /**
     * パス表示を今の設定で更新する
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function refreshPathDisplays(dialogUi) {
        if (dialogUi.pickedSourceFolder) {
            dialogUi.sourcePathLabel.text = getFolderDisplayPath(dialogUi, dialogUi.pickedSourceFolder);
        }
        refreshDestinationLabel(dialogUi);
    }

    /**
     * 保存先の表示を更新する（「保存先を対象と同じにする」がオンなら〈対象と同じ〉）
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function refreshDestinationLabel(dialogUi) {
        var destinationPathLabel = dialogUi.destinationPathLabel;
        if (dialogUi.sameAsSourceCheckbox.value && !dialogUi.overwriteRadio.value) {
            destinationPathLabel.text = getLabel("status.destinationSameAsSource");
            destinationPathLabel.helpTip = dialogUi.pickedSourceFolder ? dialogUi.pickedSourceFolder.fsName : getLabel("tooltip.sameAsSource");
            return;
        }
        if (dialogUi.pickedDestinationFolder) {
            destinationPathLabel.text = getFolderDisplayPath(dialogUi, dialogUi.pickedDestinationFolder);
            destinationPathLabel.helpTip = dialogUi.pickedDestinationFolder.fsName;
        } else {
            destinationPathLabel.text = getLabel("status.destinationNotSet");
            destinationPathLabel.helpTip = getLabel("tooltip.destinationFolder");
        }
    }

    /**
     * モード（上書き／カスタム）と「保存先を対象と同じにする」に合わせて、各コントロールの状態を更新する
     * @param {Object} dialogUi - コントロールの控え
     * @returns {void}
     */
    function updateOverwriteState(dialogUi) {
        var isOverwrite = dialogUi.overwriteRadio.value;
        var hideDestination = isOverwrite || dialogUi.sameAsSourceCheckbox.value;
        if (isOverwrite) {
            dialogUi.convertedCheckbox.value = false;
            dialogUi.suffixCheckbox.value = false;
            dialogUi.customCheckbox.value = false;
        } else {
            dialogUi.suffixCheckbox.value = true;
        }
        dialogUi.convertedCheckbox.enabled = !isOverwrite;
        dialogUi.suffixCheckbox.enabled = !isOverwrite;
        dialogUi.suffixInput.enabled = !isOverwrite && dialogUi.suffixCheckbox.value;
        dialogUi.customCheckbox.enabled = !isOverwrite;
        dialogUi.customInput.enabled = !isOverwrite && dialogUi.customCheckbox.value;
        dialogUi.separatorLabel.enabled = !isOverwrite;
        for (var i = 0; i < dialogUi.separatorRadios.length; i++) {
            dialogUi.separatorRadios[i].enabled = !isOverwrite;
        }
        /* 上書きモードでは sameAsSource は暗黙的に true なので操作不可 / In overwrite mode, sameAsSource is implicit */
        dialogUi.sameAsSourceCheckbox.enabled = !isOverwrite;
        dialogUi.destinationButton.enabled = !hideDestination;
        dialogUi.destinationPathLabel.enabled = !hideDestination;
        refreshDestinationLabel(dialogUi);
    }

    /**
     * 選ばれている区切り文字を返す
     * @param {Object} dialogUi - コントロールの控え
     * @returns {string} 区切り文字
     */
    function getSelectedSeparator(dialogUi) {
        for (var i = 0; i < dialogUi.separatorRadios.length; i++) {
            if (dialogUi.separatorRadios[i].value) return SEPARATOR_OPTIONS[i];
        }
        return SEPARATOR_OPTIONS[SEPARATOR_DEFAULT_INDEX];
    }

    /**
     * ダイアログの設定を読み取る
     * @param {Object} dialogUi - コントロールの控え
     * @returns {Object} 設定
     */
    function readDialogSettings(dialogUi) {
        var selectedVersion = VERSION_OPTIONS[dialogUi.versionDropdown.selection.index];
        return {
            compatibility: selectedVersion.compatibility,
            suffix: dialogUi.suffixCheckbox.value ? dialogUi.suffixInput.text : "",
            customSuffix: dialogUi.customCheckbox.value ? dialogUi.customInput.text : "",
            separator: getSelectedSeparator(dialogUi),
            pdfCompatible: dialogUi.pdfCheckbox.value,
            appendConverted: dialogUi.convertedCheckbox.value,
            overwrite: dialogUi.overwriteRadio.value,
            sameAsSource: dialogUi.sameAsSourceCheckbox.value,
            targetAi: dialogUi.aiCheckbox.value,
            targetSvg: dialogUi.svgCheckbox.value,
            sourceFolder: dialogUi.pickedSourceFolder,
            destinationFolder: dialogUi.pickedDestinationFolder
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 保存先を決める（上書き・「対象と同じ」→対象フォルダー／ダイアログで指定済み→それ／未指定→ここで聞く）
     * @param {Object} batchSettings - ダイアログの設定
     * @param {Folder} sourceFolder - 対象フォルダー
     * @returns {Folder|null} 保存先（キャンセル時は null）
     */
    function resolveDestinationFolder(batchSettings, sourceFolder) {
        if (batchSettings.overwrite || batchSettings.sameAsSource) return sourceFolder;
        if (batchSettings.destinationFolder != null) return batchSettings.destinationFolder;
        return selectFolder("prompt.selectDestFolder");
    }

    /**
     * 上書きの危険がある組み合わせなら確認する
     * @param {Object} batchSettings - ダイアログの設定
     * @param {File[]} targetFileList - 保存するファイル
     * @param {Folder} sourceFolder - 対象フォルダー
     * @param {Folder} destinationFolder - 保存先フォルダー
     * @returns {boolean} 続行するなら true
     */
    function confirmOverwriteRisks(batchSettings, targetFileList, sourceFolder, destinationFolder) {
        /* 上書き事故防止：実際に .ai 入力があり、同一フォルダかつサフィックスなしなら確認 / Prevent accidental overwrite only when actual .ai inputs exist */
        if (!batchSettings.overwrite &&
            containsFileType(targetFileList, isIllustratorFile) &&
            String(sourceFolder.fsName) === String(destinationFolder.fsName) &&
            batchSettings.suffix === "" &&
            batchSettings.customSuffix === "" &&
            !batchSettings.appendConverted) {
            if (!confirm(getLabel("alert.confirmSameFolder"))) return false;
        }

        /* 上書きモードで実際に .svg を処理する場合のみ、同名 .ai の上書き可能性を確認 / Confirm only when SVG files actually exist in the batch */
        if (batchSettings.overwrite && batchSettings.targetSvg && containsFileType(targetFileList, isSvgFile)) {
            if (!confirm(getLabel("alert.confirmOverwriteSvg"))) return false;
        }
        return true;
    }

    /**
     * フォルダー内のファイルを指定のバージョンで一括保存する
     * @returns {void}
     */
    function main() {
        if (VERSION_OPTIONS.length === 0) {
            alert(getLabel("alert.noVersionsAvailable"));
            return;
        }

        var batchSettings = showSettingsDialog();
        if (batchSettings == null) return;

        /* 対象が一つも選択されていない場合は早期終了（フォルダー選択ダイアログを開く前に弾く）/ Early exit when no target type is selected (before any folder picker) */
        if (!batchSettings.targetAi && !batchSettings.targetSvg) {
            alert(getLabel("alert.noTargetSelected"));
            return;
        }

        /* 入力フォルダ決定：ダイアログで指定済→それ／未指定→ここで聞く / Resolve source: picked-in-dialog → use it / otherwise → ask now */
        var sourceFolder = (batchSettings.sourceFolder != null) ? batchSettings.sourceFolder : selectFolder("prompt.selectSourceFolder");
        if (sourceFolder == null) return;

        var destinationFolder = resolveDestinationFolder(batchSettings, sourceFolder);
        if (destinationFolder == null) return;

        /* 対象パネルで選択された拡張子（.ai / .svg）のファイルを収集 / Collect files matching the selected target extensions */
        var targetFileList = sourceFolder.getFiles(buildTargetFileFilter(batchSettings.targetAi, batchSettings.targetSvg));
        if (targetFileList.length === 0) {
            alert(getLabel("alert.noFilesFound"));
            return;
        }

        if (!confirmOverwriteRisks(batchSettings, targetFileList, sourceFolder, destinationFolder)) return;

        runBatchSave(targetFileList, destinationFolder, batchSettings);
    }

    main();

})();
