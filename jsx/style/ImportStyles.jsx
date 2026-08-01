#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- あらかじめ登録しておいたAIファイルを一覧から選び、その中身を現在のドキュメントへ取り込みます。
- 機能の詳細と使い方はREADME（readme-ja/ImportStyles.md）を参照してください。

*/

/*

### Overview

- Pick one of the AI files you registered in advance and import its contents into the current document.
- See the README (readme-en/ImportStyles.md) for the full feature list and usage.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportStyles";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportStyles.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportStyles.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0b929db4a4ad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* スタイル用AIファイルのベースフォルダ（初期状態は未設定。初回起動時に選択して保存）
       Base folder for style AI files (empty until the user picks one on first launch) */
    var styleLibraryFolder = "";

    /* フォルダー設定の保存先キー（Illustrator環境設定に永続化）/ Preference key for the folder setting */
    var PREF_KEY_LIBRARY_FOLDER = "ImportStyles.libraryFolder";

    /* 登録リストの外部ファイル名（TSV）/ External library TSV
       フォーマット: label \t path \t category  ※旧形式の check(0|1) は読み込み時に無視（後方互換） */
    var LIBRARY_TSV_NAME = "ImportStyles_candidates.tsv";

    /* 貼り付け先レイヤー名 / Destination layer name */
    var IMPORT_LAYER_NAME = "// _imported";

    /* ダイアログの初期表示位置を右へずらす量（px）/ Horizontal offset for the dialog position */
    var DIALOG_OFFSET_X = 300;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のロケールから言語コードを取得 / Get language code from locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "スタイル読み込み：ファイル選択", en: "Import Styles: Choose File" }
        },
        /* リスト列見出し / List column headers */
        column: {
            content: { ja: "コンテンツ", en: "Content" },
            filename: { ja: "ファイル名", en: "Filename" }
        },
        /* 検索 / Search */
        search: {
            label: { ja: "検索", en: "Search" },
            placeholder: {
                ja: "コンテンツ / ファイル名で絞り込み",
                en: "Filter by Content / Filename"
            }
        },
        /* カテゴリ / Category */
        category: {
            all: { ja: "すべて", en: "All" },
            styleBrushSymbol: { ja: "スタイル／ブラシ／シンボル", en: "Style/Brush/Symbol" },
            font: { ja: "フォント", en: "Fonts" },
            panelTitle: { ja: "カテゴリー", en: "Category" }
        },
        /* 貼り付け先 / Destination */
        destination: {
            label: { ja: "ペースト先", en: "Paste into" },
            currentLayer: { ja: "現在のレイヤー", en: "Current layer" },
            importLayer: { ja: "指定レイヤー", en: "Dedicated layer" }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            add: { ja: "追加", en: "Add" },
            load: { ja: "読み込み", en: "Load" },
            register: { ja: "登録", en: "Register" },
            folder: { ja: "フォルダー…", en: "Folder…" }
        },
        /* 入力プロンプト / Prompts */
        prompt: {
            pickAi: { ja: "追加するAIファイルを選択", en: "Choose an AI file to add" },
            pickLibraryFolder: {
                ja: "スタイル用AIファイルを置くフォルダーを選択",
                en: "Choose the folder that holds your style AI files"
            },
            enterLabel: { ja: "表示される項目名", en: "Display name for the dialog" },
            registerTitle: { ja: "読み込み設定", en: "Import Settings" }
        },
        /* 実行時メッセージ / Runtime messages */
        message: {
            openDocFirst: {
                ja: "元のドキュメントを開いてから実行してください。",
                en: "Please open the destination document first."
            },
            folderRequired: {
                ja: "スタイル用AIファイルのフォルダーが設定されていないため、終了します。",
                en: "No folder for style AI files was set, so the script has stopped."
            },
            nothingToCopy: {
                ja: "作業アートボード内にオブジェクトがないため、中止しました：\n",
                en: "Nothing on the active artboard, so the import was cancelled:\n"
            },
            layerNotEditable: {
                ja: "現在のレイヤーがロックまたは非表示のため、ペーストできません：\n",
                en: "The current layer is locked or hidden, so nothing can be pasted into it:\n"
            },
            fileNotFoundTitle: { ja: "ファイルが見つかりません", en: "File Not Found" },
            fileNotFoundBody: {
                ja: "指定したファイルが見つかりません：\n",
                en: "The specified file was not found:\n"
            },
            added: { ja: "候補を追加しました：", en: "Candidate added:" },
            overwritten: { ja: "候補を上書きしました：", en: "Candidate overwritten:" },
            synced: {
                ja: "\n\nImportStyles_candidates.tsv と現在のダイアログに反映しました。",
                en: "\n\nReflected in ImportStyles_candidates.tsv and the current dialog."
            },
            folderHint: {
                ja: "現在のフォルダー：",
                en: "Current folder: "
            }
        },
        /* エラー / Errors */
        error: {
            createFolder: {
                ja: "保存先フォルダを作成できませんでした。\n",
                en: "Could not create folder:\n"
            },
            copyFile: {
                ja: "ファイルをコピーできませんでした。\n",
                en: "Could not copy file to:\n"
            },
            tsvFolder: {
                ja: "候補TSVの保存先を作成できませんでした。\n",
                en: "Could not create folder for the candidates TSV:\n"
            },
            tsvOpenWrite: {
                ja: "候補TSVを書き込み用に開けませんでした。\n",
                en: "Could not open candidates TSV for writing:\n"
            }
        }
    };

    /* 指定キーのローカライズ文字列を取得（ドット区切りパス対応）/ Resolve localized string by dotted key path */
    function L(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) break;
            node = node[parts[i]];
        }
        return (node && node[currentLanguage] != null) ? node[currentLanguage] : ("[" + key + "]");
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ボタン列など）/ Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // フォルダー設定 / Library Folder Setting
    // =========================================

    /* 保存済みのフォルダー設定を読み込む（未設定なら空文字）/ Read the saved folder setting ("" if unset) */
    function loadSavedLibraryFolder() {
        var saved = app.preferences.getStringPreference(PREF_KEY_LIBRARY_FOLDER);
        return (saved != null) ? String(saved) : "";
    }

    /* フォルダーを選択して設定を保存（キャンセルで空文字）/ Pick a folder and store it ("" if cancelled) */
    function pickAndSaveLibraryFolder() {
        var pickedFolder = Folder.selectDialog(L("prompt.pickLibraryFolder"));
        if (!pickedFolder) return "";
        var folderPath = pickedFolder.fsName + "/";
        app.preferences.setStringPreference(PREF_KEY_LIBRARY_FOLDER, folderPath);
        return folderPath;
    }

    /* 有効なフォルダー設定を確定（未設定・不在なら選択させる）/ Resolve a usable folder (ask when unset or missing) */
    function resolveLibraryFolder() {
        var savedPath = loadSavedLibraryFolder();
        if (savedPath !== "" && Folder(savedPath).exists) return savedPath;
        return pickAndSaveLibraryFolder();
    }

    /* 登録リストのTSVファイル / The library TSV file */
    function getLibraryTsvFile() {
        return File(Folder(styleLibraryFolder).fsName + "/" + LIBRARY_TSV_NAME);
    }

    // =========================================
    // スタイルライブラリ / Style Library
    // =========================================

    /* カテゴリ名の正規化値 / Canonical category names */
    var CATEGORY_STYLE = "スタイル／ブラシ／シンボル";
    var CATEGORY_FONT = "フォント";

    /* カテゴリ名の正規化（フォント系だけ判定し、他はすべてスタイル扱い）/ Normalize category (font-like → Font, else Style) */
    function normalizeCategory(rawCategory) {
        var text = String(rawCategory != null ? rawCategory : "");
        var compact = text.toLowerCase().replace(/\s+/g, "");
        var isFont = /font/.test(compact) || text.indexOf("フォント") !== -1 || text.indexOf("フォンツ") !== -1;
        return isFont ? CATEGORY_FONT : CATEGORY_STYLE;
    }

    /* 保存済みファイル名を実ファイルに合わせて正規化（旧版が書いた%エンコード名の救済）
       Resolve a stored file name against the actual file (recovers legacy percent-encoded names) */
    function resolveStoredFileName(storedName) {
        if (!/%[0-9A-Fa-f]{2}/.test(storedName)) return storedName;
        var decoded = decodeFileName(storedName);
        // %を含む実在の名前を壊さないよう、復号後に実在する場合だけ採用
        // Adopt the decoded form only when it actually exists, so real names with % survive
        return (decoded !== storedName && File(styleLibraryFolder + decoded).exists) ? decoded : storedName;
    }

    /* TSVの1行をライブラリ項目へ変換（不正行は null）/ Parse one TSV line into a library item (null if invalid) */
    function parseLibraryLine(line) {
        if (!line || line.charAt(0) === "#") return null;
        var columns = line.split("\t");
        if (columns.length < 2 || columns[1] === "") return null;

        // 旧形式: label, path, check(0|1), category / 新形式: label, path, category
        var categoryIndex = (columns.length >= 3 && /^(0|1)$/.test(columns[2])) ? 3 : 2;
        var rawCategory = (columns.length > categoryIndex && columns[categoryIndex] !== "") ? columns[categoryIndex] : CATEGORY_STYLE;
        var fileName = resolveStoredFileName(columns[1]);

        return {
            label: (columns[0] !== "") ? columns[0] : fileName.replace(/\.[^\.]+$/, ""),
            fileName: fileName,
            category: normalizeCategory(rawCategory)
        };
    }

    /* フォルダー内のAIファイルを走査して項目化（TSV未提供時のフォールバック）
       Scan the folder for AI files (fallback when no TSV is available) */
    function scanFolderForAiFiles() {
        var libraryFolder = Folder(styleLibraryFolder);
        if (!libraryFolder.exists) return [];

        // 拡張子の大小を問わず拾う（文字列マスクはmacOSで大小を区別）
        // Match the extension case-insensitively (string masks are case-sensitive on macOS)
        var aiFiles = libraryFolder.getFiles(function(entry) {
            return (entry instanceof File) && /\.ai$/i.test(decodeFileName(entry.name));
        });

        var scanned = [];
        for (var i = 0; i < aiFiles.length; i++) {
            var fileName = decodeFileName(aiFiles[i].name);
            scanned.push({
                label: fileName.replace(/\.[^\.]+$/, ""),
                fileName: fileName,
                category: normalizeCategory(fileName)
            });
        }
        return scanned;
    }

    /* ライブラリをTSVから読み込み（無ければフォルダー内のAIファイルを列挙）
       Load the library from TSV (falls back to scanning the folder) */
    function loadStyleLibrary() {
        var tsvFile = getLibraryTsvFile();
        if (!tsvFile.exists) return scanFolderForAiFiles();

        tsvFile.encoding = "UTF-8";
        if (!tsvFile.open("r")) return scanFolderForAiFiles();

        var loaded = [];
        while (!tsvFile.eof) {
            var styleItem = parseLibraryLine(tsvFile.readln());
            if (styleItem) loaded.push(styleItem);
        }
        tsvFile.close();
        return loaded.length ? loaded : scanFolderForAiFiles();
    }

    /* 現在のライブラリをTSVに保存（上書き）/ Save the current library to TSV (overwrite) */
    function saveStyleLibrary(styleItems) {
        var tsvFile = getLibraryTsvFile();
        var parentFolder = tsvFile.parent;
        if (parentFolder && !parentFolder.exists && !parentFolder.create()) {
            alert(L("error.tsvFolder") + parentFolder.fsName);
            return false;
        }
        tsvFile.encoding = "UTF-8";
        if (!tsvFile.open("w")) {
            alert(L("error.tsvOpenWrite") + tsvFile.fsName);
            return false;
        }

        tsvFile.writeln("# label\tpath\tcategory");
        for (var i = 0; i < styleItems.length; i++) {
            var styleItem = styleItems[i];
            if (!styleItem || !styleItem.fileName) continue;
            tsvFile.writeln(styleItem.label + "\t" + styleItem.fileName + "\t" + styleItem.category);
        }
        tsvFile.close();
        return true;
    }

    /* 現在のライブラリ（フォルダー確定後に読み込む）/ Current library (loaded once the folder is resolved) */
    var styleLibrary = [];

    /* ファイル名でライブラリ項目を検索し index を返す（無ければ -1）/ Find item index by file name (-1 if none) */
    function findLibraryIndex(fileName) {
        for (var i = 0; i < styleLibrary.length; i++) {
            if (styleLibrary[i] && styleLibrary[i].fileName === fileName) return i;
        }
        return -1;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* ファイル名を表示用にデコード / Decode a file name for display */
    function decodeFileName(fileName) {
        try {
            return decodeURI(fileName);
        } catch (e) {
            return fileName;
        }
    }

    /* 大文字小文字を無視した部分一致 / Case-insensitive contains */
    function containsIgnoreCase(haystack, needle) {
        return String(haystack).toLowerCase().indexOf(String(needle).toLowerCase()) !== -1;
    }

    /* フォルダを確保（存在しなければ作成）/ Ensure a folder exists (create if missing) */
    function ensureFolder(folder) {
        if (!folder) return false;
        if (folder.exists || folder.create()) return true;
        alert(L("error.createFolder") + folder.fsName);
        return false;
    }

    /* 同名上書きでファイルをコピー / Copy a file with overwrite */
    function copyFileOverwrite(sourceFile, destFile) {
        if (!sourceFile || !destFile) return false;
        if (sourceFile.fsName === destFile.fsName) return true; // 同一ファイルなら何もしない / Same file
        if (destFile.exists) destFile.remove();
        if (sourceFile.copy(destFile.fsName)) return true;
        alert(L("error.copyFile") + destFile.fsName);
        return false;
    }

    /* 取り込み先レイヤーを取得（無ければ作成）/ Get or create the destination layer */
    function getOrCreateImportLayer(doc) {
        var importLayer;
        try {
            importLayer = doc.layers.getByName(IMPORT_LAYER_NAME);
        } catch (e) {
            importLayer = doc.layers.add();
            importLayer.name = IMPORT_LAYER_NAME;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    // =========================================
    // ライブラリ登録フロー / Register Flow
    // =========================================

    /* 表示名とカテゴリを入力させる（キャンセルで null）/ Ask for label + category (null on cancel) */
    function askLabelAndCategory(baseName) {
        var registerWindow = new Window("dialog", L("prompt.registerTitle") + " " + SCRIPT_VERSION);
        setupWindow(registerWindow);

        // 表示名入力 / Label field
        var labelColumn = registerWindow.add("group");
        labelColumn.orientation = "column";
        labelColumn.alignChildren = ["fill", "top"];
        labelColumn.add("statictext", undefined, labelText("prompt.enterLabel"));
        var labelField = labelColumn.add("edittext", undefined, baseName);
        labelField.characters = 20;

        // カテゴリ選択（パネル＋ラジオ）/ Category panel with radios
        var categoryPanel = registerWindow.add("panel", undefined, L("category.panelTitle"));
        setupPanel(categoryPanel, 6);
        var styleRadio = categoryPanel.add("radiobutton", undefined, L("category.styleBrushSymbol"));
        var fontRadio = categoryPanel.add("radiobutton", undefined, L("category.font"));

        // ファイル名からカテゴリを自動推定 / Auto-detect category from the file name
        var looksLikeFont = containsIgnoreCase(baseName, "font") || baseName.indexOf("フォント") !== -1;
        fontRadio.value = looksLikeFont;
        styleRadio.value = !looksLikeFont;

        // ボタン行（パネル幅いっぱいに広げない）/ Button row (buttons must not stretch)
        var registerButtonRow = registerWindow.add("group");
        setupRow(registerButtonRow, "center");
        var registerCancelButton = registerButtonRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var registerOkButton = registerButtonRow.add("button", undefined, L("button.register"), { name: "ok" });
        registerCancelButton.alignment = "left";
        registerOkButton.alignment = "left";

        if (registerWindow.show() !== 1) return null;

        // 入力値整形 / Sanitize the label
        var cleanLabel = String(labelField.text || "").replace(/[\r\n\t]/g, " ").replace(/^\s+|\s+$/g, "");

        return {
            label: (cleanLabel !== "") ? cleanLabel : baseName,
            category: fontRadio.value ? CATEGORY_FONT : CATEGORY_STYLE
        };
    }

    /* AIファイルをフォルダへコピーしてライブラリに登録（登録したファイル名を返す／中断で null）
       Copy an AI file into the library folder and register it (returns the file name, null if cancelled) */
    function registerStyleFile() {
        // 1) ファイル選択 / Pick an AI file
        var pickedFile = File.openDialog(L("prompt.pickAi"), "*.ai");
        if (!pickedFile) return null;

        // 2) 保存先フォルダを確保してコピー / Ensure the folder, then copy
        var libraryFolder = Folder(styleLibraryFolder);
        if (!ensureFolder(libraryFolder)) return null;

        var destFile = File(libraryFolder.fsName + "/" + pickedFile.name);
        if (!copyFileOverwrite(pickedFile, destFile)) return null;

        // 3) 表示名とカテゴリを入力 / Ask for label + category
        var destFileName = decodeFileName(destFile.name); // TSVには生のファイル名で保存 / Store the plain name in the TSV
        var userInput = askLabelAndCategory(destFileName.replace(/\.[^\.]+$/, ""));
        if (!userInput) return null;

        // 4) 既存項目を上書き、無ければ追加 / Overwrite the existing item or append
        var styleItem = { label: userInput.label, fileName: destFileName, category: userInput.category };
        var existingIndex = findLibraryIndex(destFileName);
        if (existingIndex >= 0) {
            styleLibrary[existingIndex] = styleItem;
        } else {
            styleLibrary.push(styleItem);
        }

        // 5) TSVへ保存 / Save to TSV
        if (!saveStyleLibrary(styleLibrary)) return null;

        alert((existingIndex >= 0 ? L("message.overwritten") : L("message.added")) + "\n" +
            userInput.label + "  (" + destFileName + ")" + L("message.synced"));
        return destFileName;
    }

    // =========================================
    // ダイアログの部品 / Dialog Parts
    // =========================================

    /* カテゴリ絞り込みラジオ（初期は「すべて」）/ Category filter radios (defaults to All) */
    function buildCategoryFilterRow(parent) {
        var categoryFilterRow = parent.add("group");
        setupRow(categoryFilterRow, "center");

        var allRadio = categoryFilterRow.add("radiobutton", undefined, L("category.all"));
        var styleRadio = categoryFilterRow.add("radiobutton", undefined, L("category.styleBrushSymbol"));
        var fontRadio = categoryFilterRow.add("radiobutton", undefined, L("category.font"));
        allRadio.value = true;

        return { allRadio: allRadio, styleRadio: styleRadio, fontRadio: fontRadio };
    }

    /* 検索欄と検索ボタン / Search field and Search button */
    function buildSearchRow(parent) {
        var searchRow = parent.add("group");
        setupRow(searchRow, "center");
        searchRow.margins = [0, 5, 0, 15];

        var searchField = searchRow.add("edittext", undefined, "");
        searchField.characters = 24;
        searchField.helpTip = L("search.placeholder");

        var searchButton = searchRow.add("button", undefined, L("search.label"));
        searchButton.alignment = "left";

        return { field: searchField, button: searchButton };
    }

    /* ヘッダ付き2列のListBox / Headered 2-column ListBox */
    function buildStyleListBox(parent) {
        var contentColumnWidth = 140;
        var filenameColumnWidth = 180;

        var styleListBox = parent.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [L("column.content"), L("column.filename")],
            columnWidths: [contentColumnWidth, filenameColumnWidth]
        });
        styleListBox.alignment = ["fill", "fill"];
        styleListBox.preferredSize = [contentColumnWidth + filenameColumnWidth + 20, 200];
        return styleListBox;
    }

    /* 貼り付け先ラジオ（初期は現在のレイヤー）/ Destination radios (defaults to the current layer) */
    function buildDestinationRow(parent) {
        var destinationRow = parent.add("group");
        setupRow(destinationRow, "center");
        destinationRow.add("statictext", undefined, labelText("destination.label"));

        var currentLayerRadio = destinationRow.add("radiobutton", undefined, L("destination.currentLayer"));
        var importLayerLabel = (currentLanguage === "ja") ?
            L("destination.importLayer") + "（" + IMPORT_LAYER_NAME + "）" :
            L("destination.importLayer") + " (" + IMPORT_LAYER_NAME + ")";
        var importLayerRadio = destinationRow.add("radiobutton", undefined, importLayerLabel);
        currentLayerRadio.value = true;

        return { currentLayerRadio: currentLayerRadio, importLayerRadio: importLayerRadio };
    }

    /* 下部ボタン行：左（キャンセル／フォルダー）－スペーサー－右（追加／読み込み）/ Footer buttons */
    function buildFooterRow(parent) {
        var footerRow = parent.add("group");
        setupRow(footerRow, "fill");
        footerRow.margins = [0, 10, 0, 0];

        var cancelButton = footerRow.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var folderButton = footerRow.add("button", undefined, L("button.folder"));
        cancelButton.alignment = "left";
        folderButton.alignment = "left";

        var flexSpacer = footerRow.add("group");
        flexSpacer.alignment = ["fill", "fill"];
        flexSpacer.minimumSize = [0, 0];

        var addButton = footerRow.add("button", undefined, L("button.add"));
        var loadButton = footerRow.add("button", undefined, L("button.load"), { name: "ok" });
        addButton.alignment = "left";
        loadButton.alignment = "left";

        return {
            cancelButton: cancelButton,
            folderButton: folderButton,
            addButton: addButton,
            loadButton: loadButton
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* ライブラリから1つ選択し、AIファイルのパスと貼り付け先を返す（キャンセルで null）
       Choose one item and return the AI file path plus the destination (null if cancelled) */
    function showStyleLibraryDialog() {
        var libraryWindow = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(libraryWindow);
        libraryWindow.opacity = 0.97;
        libraryWindow.onShow = function() {
            libraryWindow.location = [libraryWindow.location[0] + DIALOG_OFFSET_X, libraryWindow.location[1]];
        };

        var categoryFilter = buildCategoryFilterRow(libraryWindow);
        var search = buildSearchRow(libraryWindow);
        var styleListBox = buildStyleListBox(libraryWindow);
        var destination = buildDestinationRow(libraryWindow);
        var footer = buildFooterRow(libraryWindow);

        // 絞り込みの状態 / Filter state
        var activeCategory = null; // null = 全件 / null = all
        var activeQuery = "";

        /* 絞り込み条件に一致するか / Does the item match the current filter? */
        function matchesFilter(styleItem) {
            if (activeCategory != null && styleItem.category !== activeCategory) return false;
            if (activeQuery === "") return true;
            return containsIgnoreCase(styleItem.label, activeQuery) ||
                containsIgnoreCase(styleItem.fileName, activeQuery);
        }

        /* 絞り込み結果でリストを作り直す / Rebuild the list from the current filter */
        function refreshList() {
            styleListBox.removeAll();
            for (var i = 0; i < styleLibrary.length; i++) {
                var styleItem = styleLibrary[i];
                if (!styleItem || !styleItem.fileName || !matchesFilter(styleItem)) continue;

                var row = styleListBox.add("item", styleItem.label);
                row.subItems[0].text = styleItem.fileName;
                row.styleItem = styleItem; // 参照を保持 / Keep the reference
            }
            if (styleListBox.items.length > 0) styleListBox.selection = 0;
            libraryWindow.layout.layout(true);
        }

        /* 選択行を確定してダイアログを閉じる / Confirm the selected row */
        function confirmSelection() {
            if (styleListBox.selection) libraryWindow.close(1);
        }

        /* カテゴリ絞り込みを切り替え / Switch the category filter */
        function applyCategory(categoryName) {
            return function() {
                activeCategory = categoryName;
                refreshList();
            };
        }
        categoryFilter.allRadio.onClick = applyCategory(null);
        categoryFilter.styleRadio.onClick = applyCategory(CATEGORY_STYLE);
        categoryFilter.fontRadio.onClick = applyCategory(CATEGORY_FONT);

        /* 検索欄の内容を絞り込みに反映 / Apply the search field to the filter */
        function applySearch() {
            activeQuery = String(search.field.text || "");
            refreshList();
        }

        // 検索はボタン押下時と検索欄の確定時のみ適用（ライブ絞り込みはしない）
        // Apply the query on the button click and when the field is committed (no live filtering)
        search.button.onClick = applySearch;
        search.field.onChange = applySearch;

        // キーボード：Enterで決定（検索欄では検索）、⌘（Ctrl）+ Fで検索欄へ
        // Keyboard: Enter confirms (or searches while in the field), Cmd/Ctrl+F focuses the search field
        styleListBox.onDoubleClick = confirmSelection;
        libraryWindow.addEventListener("keydown", function(keyEvent) {
            if (keyEvent.keyName === "Enter") {
                if (keyEvent.target === search.field) {
                    keyEvent.preventDefault();
                    applySearch();
                } else {
                    confirmSelection();
                }
            } else if (keyEvent.keyName === "F" && (keyEvent.metaKey || keyEvent.ctrlKey)) {
                keyEvent.preventDefault();
                search.field.active = true;
            }
        });

        // 追加：AIファイルを登録してリストへ反映 / Add: register an AI file and reflect it
        footer.addButton.onClick = function() {
            var addedFileName = registerStyleFile();
            if (!addedFileName) return;

            // 追加した項目が絞り込みで隠れないよう「すべて」に戻す / Reset to All so the new item stays visible
            categoryFilter.allRadio.value = true;
            activeCategory = null;
            refreshList();
            for (var i = 0; i < styleListBox.items.length; i++) {
                if (styleListBox.items[i].styleItem.fileName === addedFileName) {
                    styleListBox.selection = i;
                    break;
                }
            }
        };

        // フォルダー：保存先を選び直してライブラリを読み込み直す / Folder: re-pick the folder and reload the library
        function updateFolderHint() {
            footer.folderButton.helpTip = L("message.folderHint") + styleLibraryFolder;
        }
        updateFolderHint();
        footer.folderButton.onClick = function() {
            var pickedPath = pickAndSaveLibraryFolder();
            if (pickedPath === "") return;
            styleLibraryFolder = pickedPath;
            styleLibrary = loadStyleLibrary();
            updateFolderHint();
            refreshList();
        };

        footer.loadButton.onClick = confirmSelection;

        // 初期表示 → ダイアログ表示 → 選択を返却 / Init, show, return the selection
        refreshList();
        if (libraryWindow.show() !== 1 || !styleListBox.selection) return null;
        return {
            filePath: styleLibraryFolder + styleListBox.selection.styleItem.fileName,
            useImportLayer: destination.importLayerRadio.value === true
        };
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================

    /* ライブラリ選択 → 対象AIをコピー → 元ドキュメントへ貼り付け / Choose, copy from the AI, paste into the original */
    function main() {
        // 貼り付け先のドキュメントを先に確認 / Check for the destination document first
        if (app.documents.length === 0) {
            alert(L("message.openDocFirst"));
            return;
        }
        var originalDoc = app.activeDocument;

        // フォルダー設定を確定し、ライブラリを読み込む / Resolve the folder setting, then load the library
        styleLibraryFolder = resolveLibraryFolder();
        if (styleLibraryFolder === "") {
            alert(L("message.folderRequired"));
            return;
        }
        styleLibrary = loadStyleLibrary();

        var choice = showStyleLibraryDialog();
        if (!choice) return;

        // 現在のレイヤーへペーストする場合は、編集可能かを確認 / When pasting into the current layer, make sure it is editable
        var currentLayer = originalDoc.activeLayer;
        if (!choice.useImportLayer && (currentLayer.locked || !currentLayer.visible)) {
            alert(L("message.layerNotEditable") + currentLayer.name);
            return;
        }

        // 選択したファイルを開く / Open the chosen file
        var styleFile = new File(choice.filePath);
        if (!styleFile.exists) {
            alert(L("message.fileNotFoundTitle") + "\n" + L("message.fileNotFoundBody") + decodeFileName(styleFile.name));
            return;
        }
        var styleDoc = app.open(styleFile);

        // 作業アートボード内をすべてコピーし、保存せずに閉じる / Copy in-artboard objects, close without saving
        app.executeMenuCommand("selectallinartboard");
        if (styleDoc.selection.length === 0) {
            // 選択が空のままコピーすると、以前のクリップボードが貼り付けられるため中止
            // Copying an empty selection would paste whatever was in the clipboard before
            styleDoc.close(SaveOptions.DONOTSAVECHANGES);
            app.activeDocument = originalDoc;
            alert(L("message.nothingToCopy") + decodeFileName(styleFile.name));
            return;
        }
        app.executeMenuCommand("copy");
        styleDoc.close(SaveOptions.DONOTSAVECHANGES);

        // 貼り付け先を決めてからペースト（現在のレイヤーはそのまま）/ Set the destination, then paste
        app.activeDocument = originalDoc;
        if (choice.useImportLayer) {
            originalDoc.activeLayer = getOrCreateImportLayer(originalDoc);
        }
        app.executeMenuCommand("paste");
    }

    main();

})();
