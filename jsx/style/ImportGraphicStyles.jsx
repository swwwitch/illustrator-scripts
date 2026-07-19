#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- ダイアログの「スタイルを読み込み」ボタンで AI ファイルを指定し、そのファイル内のグラフィックスタイルを取り込む
- 取り込んだスタイル名でラジオボタンを自動生成し、選んだスタイルを選択オブジェクトへ適用する
- 指定したファイルは記憶され、次回以降は自動で参照（ボタンで別ファイルを選ぶまで保持）

### 処理の流れ

1. 対象オブジェクトを選択した状態でスクリプトを実行
2. 記憶しているファイルがあれば取り込み、取り込んだスタイル名でラジオを表示
3. 「スタイルを読み込み」ボタンで別ファイルを選ぶと、その場で取り込み直してラジオを更新
4. スタイルを選んで「適用」→ 選択オブジェクトへグラフィックスタイルを適用

### note

- 取り込み時は「// _imported」レイヤーへ一時的に貼り付け、アセット登録後にレイヤーごと削除します（見た目は変わりません）
- 参照ファイルのパスは Folder.userData（swwwitch_ImportGraphicStyles.txt）に記憶し、スクリプト内には残しません
- ラジオは元ファイルのグラフィックスタイル名から自動生成されます（index 0 の既定スタイルは除外）

### GitHub

https://github.com/swwwitch/illustrator-scripts

### 更新履歴

- v1.7.0 (20260701) : 参照ファイルをスクリプト固定から「スタイルを読み込み」ボタンでの選択式に変更（Folder.userData に記憶）。取り込んだスタイル名からラジオを自動生成する構成に変更。固定スタイル名（文字白抜き／枠のみ）を撤去
- v1.6.0 (20260701) : 検索・追加・候補リスト（ListBox／カテゴリ選択）を撤去。ラジオ（文字白抜き／枠のみ）で選んだグラフィックスタイルを、必要に応じて AI ファイルから取り込み、選択オブジェクトへ適用する構成に変更
- v1.5.0 (20260701) : ローカライズを構造化（ネスト LABELS ＋ ドット区切り L()）、全体を IIFE 化、パネル共通設定（setupPanel）を追加、変数名・関数名を整理、追加フローを関数分割、重複コード・不要な try を削減
- v1.4 (20250815) : 標準 ListBox（2 列ヘッダ）へ刷新、削除オプションを廃止、カテゴリラジオ＆検索ボタンを追加、貼り付け先を「// _imported」に統一、ドキュメント刷新
- v1.3 (20250815) : カテゴリ列（スタイル／ブラシ／シンボル／フォント）対応、追加時にカテゴリ選択（ドロップダウン）
- v1.2 (20250815) : CANDIDATES を外部 TSV から読み込み可能に
- v1.1 (20250815) : CANDIDATES に削除オプションを記録
- v1.0 (20250814) : 初期バージョン

*/

/*

### Overview

- Pick an AI file via the dialog's "Load Styles" button and import the graphic styles it contains
- Radio buttons are generated automatically from the imported style names; the chosen style is applied to the selection
- The picked file is remembered and reused on later runs (until you pick another via the button)

### Process Flow

1. Run the script with the target objects selected
2. If a file is remembered, import it and show radios built from the imported style names
3. Clicking "Load Styles" lets you pick another file, re-importing and rebuilding the radios in place
4. Choose a style, click Apply, and the graphic style is applied to the selection

### note

- Import pastes temporarily onto the "// _imported" layer and removes the whole layer after the assets are registered (no visual change)
- The source file path is remembered in Folder.userData (swwwitch_ImportGraphicStyles.txt); no path is hardcoded in the script
- Radios are generated from the source file's graphic-style names (the default style at index 0 is excluded)

### GitHub

https://github.com/swwwitch/illustrator-scripts

### Changelog

- v1.7.0 (2026-07-01): Replaced the hardcoded source path with a "Load Styles" file picker (remembered in Folder.userData); radios are now generated automatically from the imported style names; removed the fixed style names (White text / Frame only)
- v1.6.0 (2026-07-01): Removed search, add-candidate, and the candidate list (ListBox/category selection); reworked into a radio dialog (White text / Frame only) that imports the chosen graphic style from the AI file when needed and applies it to the selection
- v1.5.0 (2026-07-01): Structured localization (nested LABELS + dotted L()), wrapped everything in an IIFE, added shared panel setup (setupPanel), tidied variable/function names, split the add flow into functions, reduced duplication and unnecessary try blocks
- v1.4 (2025-08-15): Switched to standard 2-column ListBox, removed delete option, added category radios & Search button, unified destination layer to "// _imported", refreshed docs
- v1.3 (2025-08-15): Added Category column (Style/Brush/Font) and category dropdown on Add
- v1.2 (2025-08-15): Load CANDIDATES from external TSV file
- v1.1 (2025-08-15): Added delete option in CANDIDATES to remember after loading
- v1.0 (2025-08-14): Initial release

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportGraphicStyles";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
            title: { ja: "スタイルを適用", en: "Apply Style" },
            pickFile: { ja: "スタイルの AI ファイルを選択", en: "Select a style AI file" }
        },
        /* スタイル選択 / Style choices */
        style: {
            panelTitle: { ja: "スタイル", en: "Style" }
        },
        /* スタイルの読み込み / Load styles */
        load: {
            panelTitle: { ja: "スタイルの読み込み", en: "Load Styles" }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" },
            loadFile: { ja: "読み込み", en: "Load" }
        },
        /* 実行時メッセージ / Runtime messages */
        message: {
            openDocFirst: {
                ja: "元のドキュメントを開いてから実行してください。",
                en: "Please open the destination document first."
            },
            selectObjectFirst: {
                ja: "スタイルを適用するオブジェクトを選択してから実行してください。",
                en: "Please select the object(s) to apply the style to."
            },
            fileNotFoundTitle: { ja: "ファイルが見つかりません", en: "File Not Found" },
            fileNotFoundBody: {
                ja: "指定したファイルが見つかりません：\n",
                en: "The specified file was not found:\n"
            },
            styleNotFound: {
                ja: "指定したグラフィックスタイルが見つかりません：\n",
                en: "The graphic style was not found:\n"
            },
            noFileSelected: { ja: "ファイル未選択", en: "No file selected" },
            noStylesHint: {
                ja: "「スタイルを読み込み」でファイルを選択してください。",
                en: "Click “Load Styles” to choose a file."
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

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* 取り込み先レイヤーを取得（無ければ作成）/ Get or create the destination layer "// _imported" */
    function getOrCreateImportLayer(destinationDoc) {
        var importLayerName = "// _imported";
        var importLayer;
        try {
            importLayer = destinationDoc.layers.getByName(importLayerName);
        } catch (e) {
            importLayer = destinationDoc.layers.add();
            importLayer.name = importLayerName;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    /* パスから表示用ファイル名を取得（文字化け対策で decodeURI）/ Get a display-friendly filename from a path */
    function getDisplayFileName(filePath) {
        try {
            return decodeURI(new File(filePath).name);
        } catch (e) {
            return filePath;
        }
    }

    /* 名前でグラフィックスタイルを取得（無ければ null）/ Get a graphic style by name (null if none) */
    function findGraphicStyle(destinationDoc, styleName) {
        try {
            return destinationDoc.graphicStyles.getByName(styleName);
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // 設定の保存・読込 / Preferences (remember last file)
    // =========================================

    /* 設定ファイル（前回のスタイルファイルのパスを記憶）/ Prefs file that remembers the last style file */
    function getPrefsFile() {
        return new File(Folder.userData + "/swwwitch_ImportGraphicStyles.txt");
    }

    /* 記憶しているスタイルファイルのパスを読み込む（無ければ空文字）/ Load the remembered style-file path */
    function loadSavedFilePath() {
        var prefsFile = getPrefsFile();
        if (!prefsFile.exists) return "";
        var savedPath = "";
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("r");
            var content = prefsFile.read();
            prefsFile.close();
            var lines = content.split(/\r\n|\r|\n/);
            for (var i = 0; i < lines.length; i++) {
                var separatorIndex = lines[i].indexOf("=");
                if (separatorIndex < 0) continue;
                if (lines[i].substring(0, separatorIndex) === "styleFilePath") {
                    savedPath = lines[i].substring(separatorIndex + 1);
                }
            }
        } catch (e) {}
        return savedPath;
    }

    /* スタイルファイルのパスを記憶する（key=value 形式）/ Remember the style-file path (key=value) */
    function saveFilePath(filePath) {
        var prefsFile = getPrefsFile();
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("w");
            prefsFile.write("styleFilePath=" + filePath + "\n");
            prefsFile.close();
        } catch (e) {}
    }

    /* スタイル用 AI ファイルを選ばせる（キャンセルで空文字）/ Let the user pick a style AI file */
    function pickStyleFile() {
        var picked = File.openDialog(L("dialog.pickFile"), function (candidate) {
            return (candidate instanceof Folder) || /\.ai$/i.test(candidate.name);
        });
        return picked ? picked.fsName : "";
    }

    // =========================================
    // スタイル取り込み / Import Styles
    // =========================================

    /* 対象AIを開いてスタイル名を取得→コピー→一時レイヤーへ貼り付け→レイヤーごと削除（アセットのみ登録）
       戻り値: 取り込んだグラフィックスタイル名の配列（ファイル未検出時は警告して null）
       Open the AI, read its style names, copy, paste to a temp layer, then remove it (assets only stay registered).
       Returns imported graphic-style names (null if the file is missing) */
    function importStylesFrom(destinationDoc, filePath) {
        var styleFile = new File(filePath);
        if (!styleFile.exists) {
            alert(L("message.fileNotFoundTitle") + "\n" + L("message.fileNotFoundBody") + getDisplayFileName(filePath));
            return null;
        }
        var styleSourceDoc = app.open(styleFile);

        // 元ファイルのグラフィックスタイル名を取得（index 0 の既定スタイルは除外）
        // Collect style names from the source (skip the default style at index 0)
        var sourceStyleNames = [];
        for (var i = 1; i < styleSourceDoc.graphicStyles.length; i++) {
            sourceStyleNames.push(styleSourceDoc.graphicStyles[i].name);
        }

        // 作業アートボード内の全てをコピーして保存せず閉じる / Copy in-artboard objects, close without saving
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);

        // 一時レイヤーへ貼り付け→パネル登録後にレイヤーごと削除 / Paste to temp layer, then remove it
        app.activeDocument = destinationDoc;
        var importLayer = getOrCreateImportLayer(destinationDoc);
        destinationDoc.activeLayer = importLayer;
        app.executeMenuCommand("paste");
        try {
            importLayer.remove();
        } catch (e) {}

        // 実際に書類へ登録されたスタイル名だけを返す / Keep only names actually registered in the destination
        var importedStyleNames = [];
        for (var k = 0; k < sourceStyleNames.length; k++) {
            if (findGraphicStyle(destinationDoc, sourceStyleNames[k])) importedStyleNames.push(sourceStyleNames[k]);
        }
        return importedStyleNames;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* スタイルを選ばせ、選んだグラフィックスタイル名を返す（キャンセルで null）
       「スタイルを読み込み」ボタンで別ファイルを指定すると、その場でラジオを組み直す
       Ask which style, return its name. The "Load Styles" button re-imports and rebuilds the radios in place */
    function runStyleDialog(destinationDoc, initialFilePath, initialStyleNames) {
        var state = {
            filePath: initialFilePath || "",
            styleNames: initialStyleNames || []
        };

        var styleDialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        styleDialog.orientation = "column";
        styleDialog.alignChildren = ["fill", "top"];

        // スタイル選択パネル / Style panel
        var stylePanel = styleDialog.add("panel", undefined, L("style.panelTitle"));
        stylePanel.orientation = "column";
        stylePanel.alignChildren = ["left", "top"];
        stylePanel.margins = [16, 20, 16, 12];
        stylePanel.spacing = 8;

        // ラジオを差し替えるためのコンテナ / Container whose radios get rebuilt on reload
        var radioGroup = stylePanel.add("group");
        radioGroup.orientation = "column";
        radioGroup.alignChildren = ["left", "top"];
        radioGroup.spacing = 8;
        var styleRadios = [];

        // スタイルの読み込みパネル（ボタンの下にファイルパスを表示）/ Load-styles panel (path shown below the button)
        var loadPanel = styleDialog.add("panel", undefined, L("load.panelTitle"));
        loadPanel.orientation = "column";
        loadPanel.alignChildren = ["left", "top"];
        loadPanel.margins = [16, 20, 16, 12];
        loadPanel.spacing = 8;
        var loadFileButton = loadPanel.add("button", undefined, L("button.loadFile"));
        var fileNameText = loadPanel.add("statictext", undefined, "", { truncate: "middle" });
        fileNameText.preferredSize.width = 240;

        // ボタン行（Mac 規約: キャンセル → 適用）/ Buttons (Mac order: Cancel → Apply)
        var dialogButtonGroup = styleDialog.add("group");
        dialogButtonGroup.alignment = ["right", "bottom"];
        dialogButtonGroup.alignChildren = ["right", "center"];
        dialogButtonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var applyButton = dialogButtonGroup.add("button", undefined, L("button.apply"), { name: "ok" });

        /* 選択中のファイル名表示を更新 / Update the file-name label */
        function refreshFileLabel() {
            fileNameText.text = state.filePath ? getDisplayFileName(state.filePath) : L("message.noFileSelected");
        }

        /* ラジオを現在のスタイル名で組み直す / Rebuild radios from the current style names */
        function rebuildRadios() {
            for (var i = radioGroup.children.length - 1; i >= 0; i--) radioGroup.remove(radioGroup.children[i]);
            styleRadios = [];
            if (state.styleNames.length === 0) {
                radioGroup.add("statictext", undefined, L("message.noStylesHint"));
                applyButton.enabled = false;
            } else {
                for (var j = 0; j < state.styleNames.length; j++) {
                    styleRadios.push(radioGroup.add("radiobutton", undefined, state.styleNames[j]));
                }
                styleRadios[0].value = true; // 既定 / Default
                applyButton.enabled = true;
            }
            styleDialog.layout.layout(true);
            styleDialog.layout.resize();
        }

        // onClick で連結（addEventListener は発火しない環境があるため）/ Use onClick, not addEventListener
        loadFileButton.onClick = function () {
            var pickedPath = pickStyleFile();
            if (!pickedPath) return;
            var importedStyleNames = importStylesFrom(destinationDoc, pickedPath);
            if (importedStyleNames === null) return; // ファイル未検出は importStylesFrom 側で警告済み
            state.filePath = pickedPath;
            state.styleNames = importedStyleNames;
            saveFilePath(pickedPath); // 次回以降このファイルを参照 / Remember for next runs
            refreshFileLabel();
            rebuildRadios();
        };

        refreshFileLabel();
        rebuildRadios();

        if (styleDialog.show() !== 1) return null;
        for (var k = 0; k < styleRadios.length; k++) {
            if (styleRadios[k].value) return state.styleNames[k];
        }
        return null;
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================

    /* 記憶したファイルから取り込み→スタイル選択→選択オブジェクトへ適用 / Import from the remembered file, choose, apply */
    function main() {
        // ドキュメントの存在を確認 / Check for an open document
        if (app.documents.length === 0) {
            alert(L("message.openDocFirst"));
            return;
        }
        var destinationDoc = app.activeDocument;

        // 選択オブジェクトを確保 / Capture current selection
        var currentSelection = destinationDoc.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(L("message.selectObjectFirst"));
            return;
        }
        var selectedItems = [];
        for (var i = 0; i < currentSelection.length; i++) selectedItems.push(currentSelection[i]);

        // 記憶しているファイルがあればダイアログ前に取り込む / If a file is remembered, import it before the dialog
        var savedFilePath = loadSavedFilePath();
        var initialFilePath = "";
        var initialStyleNames = [];
        if (savedFilePath) {
            initialFilePath = savedFilePath;
            if (new File(savedFilePath).exists) {
                var importedStyleNames = importStylesFrom(destinationDoc, savedFilePath);
                if (importedStyleNames !== null) initialStyleNames = importedStyleNames;
            }
            // ファイルが見つからない場合はダイアログの「スタイルを読み込み」で選び直してもらう
            // If the file is missing, the user re-picks it via "Load Styles"
        }

        // スタイルを選択（読み込みボタンで別ファイルにも切替可）/ Choose the style (Load Styles can switch files)
        var styleName = runStyleDialog(destinationDoc, initialFilePath, initialStyleNames);
        if (!styleName) return;

        // スタイルを取得 / Get the style
        var graphicStyle = findGraphicStyle(destinationDoc, styleName);
        if (!graphicStyle) {
            alert(L("message.styleNotFound") + styleName);
            return;
        }

        // 選択オブジェクトへ適用 / Apply to the selected objects
        for (var j = 0; j < selectedItems.length; j++) {
            try {
                graphicStyle.applyTo(selectedItems[j]);
            } catch (e) {}
        }

        // 選択状態を復帰 / Restore selection
        destinationDoc.selection = selectedItems;
    }

    main();

})();
