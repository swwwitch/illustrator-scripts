#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- ダイアログのラジオボタン（文字白抜き／枠のみ）でグラフィックスタイルを選択
- 選択したスタイルがドキュメントに無ければ、定義済みの AI ファイル（TARGET_FILE_PATH）から取り込む
- 選択中のオブジェクトへ、選んだグラフィックスタイルを適用する

### 処理の流れ

1. 対象オブジェクトを選択した状態でスクリプトを実行
2. ダイアログでスタイル（文字白抜き／枠のみ）を選択
3. スタイルが未登録なら対象 AI を開いてコピー→貼り付けで取り込み、一時オブジェクト・レイヤーを削除
4. 選択オブジェクトへグラフィックスタイルを適用

### note

- 取り込み時は「// _imported」レイヤーへ一時的に貼り付け、アセット登録後にレイヤーごと削除します（見た目は変わりません）
- 読み込むファイルは TARGET_FILE_PATH、スタイル名は STYLE_NAME_WHITE_TEXT / STYLE_NAME_FRAME_ONLY で切り替えられます

### GitHub

https://github.com/swwwitch/illustrator-scripts

### 更新履歴

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

- Choose a graphic style via dialog radio buttons (White text / Frame only)
- If the chosen style is not in the document, import it from a predefined AI file (TARGET_FILE_PATH)
- Apply the chosen graphic style to the currently selected objects

### Process Flow

1. Run the script with the target objects selected
2. Choose a style (White text / Frame only) in the dialog
3. If the style is not registered, open the target AI, copy/paste to import, then remove the temp objects/layer
4. Apply the graphic style to the selected objects

### note

- Import pastes temporarily onto the "// _imported" layer and removes the whole layer after the assets are registered (no visual change)
- Switch the source file via TARGET_FILE_PATH and style names via STYLE_NAME_WHITE_TEXT / STYLE_NAME_FRAME_ONLY

### GitHub

https://github.com/swwwitch/illustrator-scripts

### Changelog

- v1.6.0 (2026-07-01): Removed search, add-candidate, and the candidate list (ListBox/category selection); reworked into a radio dialog (White text / Frame only) that imports the chosen graphic style from the AI file when needed and applies it to the selection
- v1.5.0 (2026-07-01): Structured localization (nested LABELS + dotted L()), wrapped everything in an IIFE, added shared panel setup (setupPanel), tidied variable/function names, split the add flow into functions, reduced duplication and unnecessary try blocks
- v1.4 (2025-08-15): Switched to standard 2-column ListBox, removed delete option, added category radios & Search button, unified destination layer to "// _imported", refreshed docs
- v1.3 (2025-08-15): Added Category column (Style/Brush/Font) and category dropdown on Add
- v1.2 (2025-08-15): Load CANDIDATES from external TSV file
- v1.1 (2025-08-15): Added delete option in CANDIDATES to remember after loading
- v1.0 (2025-08-14): Initial release

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.6.0";

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* スタイルを取り込む AI ファイル（環境に合わせて変更）/ AI file to import styles from */
    var TARGET_FILE_PATH = "/Users/takano/sw Dropbox/takano masahiro/sync-setting/ai/styles/StyleForAreaType.ai";

    /* グラフィックスタイル名（AIファイル内の登録名に合わせる）/ Graphic style names as registered in the AI file */
    var STYLE_NAME_WHITE_TEXT = "文字白抜き"; // ラジオ「文字白抜き」/ Radio "White text"
    var STYLE_NAME_FRAME_ONLY = "枠のみ";     // ラジオ「枠のみ」/ Radio "Frame only"

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
            title: { ja: "スタイルを適用", en: "Apply Style" }
        },
        /* スタイル選択 / Style choices */
        style: {
            panelTitle: { ja: "スタイル", en: "Style" },
            whiteText: { ja: "文字白抜き", en: "White text" },
            frameOnly: { ja: "枠のみ", en: "Frame only" }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" }
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
    // ダイアログ / Dialog
    // =========================================

    /* スタイルを選ばせ、対応するグラフィックスタイル名を返す（キャンセルで null）/ Ask which style, return its graphic-style name */
    function chooseGraphicStyleName() {
        var styleDialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        styleDialog.orientation = "column";
        styleDialog.alignChildren = ["fill", "top"];

        // スタイル選択（パネル＋ラジオ）/ Style panel with radios
        var stylePanel = styleDialog.add("panel", undefined, L("style.panelTitle"));
        stylePanel.orientation = "column";
        stylePanel.alignChildren = ["left", "top"];
        stylePanel.margins = [16, 20, 16, 12];
        stylePanel.spacing = 8;
        var whiteTextRadio = stylePanel.add("radiobutton", undefined, L("style.whiteText"));
        var frameOnlyRadio = stylePanel.add("radiobutton", undefined, L("style.frameOnly"));
        whiteTextRadio.value = true; // 既定 / Default

        // ボタン行（Mac 規約: キャンセル → 適用）/ Buttons (Mac order: Cancel → Apply)
        var dialogButtonGroup = styleDialog.add("group");
        dialogButtonGroup.alignment = ["right", "bottom"];
        dialogButtonGroup.alignChildren = ["right", "center"];
        dialogButtonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        dialogButtonGroup.add("button", undefined, L("button.apply"), { name: "ok" });

        if (styleDialog.show() !== 1) return null;
        return frameOnlyRadio.value ? STYLE_NAME_FRAME_ONLY : STYLE_NAME_WHITE_TEXT;
    }

    // =========================================
    // スタイル取り込み / Import Styles
    // =========================================

    /* 対象AIを開いてコピー→一時レイヤーへ貼り付け→レイヤーごと削除（アセットのみ登録）/ Import assets from the target AI */
    function importStyles(destinationDoc) {
        var styleFile = new File(encodeURI(TARGET_FILE_PATH));
        if (!styleFile.exists) {
            alert(L("message.fileNotFoundTitle") + "\n" + L("message.fileNotFoundBody") + getDisplayFileName(TARGET_FILE_PATH));
            return false;
        }
        var styleSourceDoc = app.open(styleFile);

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
        return true;
    }

    /* スタイルを取得（未登録なら取り込んでから取得）／見つからなければ警告して null / Get the style, importing it if needed */
    function ensureGraphicStyle(destinationDoc, styleName) {
        var graphicStyle = findGraphicStyle(destinationDoc, styleName);
        if (graphicStyle) return graphicStyle;

        if (!importStyles(destinationDoc)) return null; // importStyles 側でファイル未検出を警告 / importStyles alerts on missing file
        graphicStyle = findGraphicStyle(destinationDoc, styleName);
        if (!graphicStyle) alert(L("message.styleNotFound") + styleName);
        return graphicStyle;
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================

    /* スタイル選択→（必要なら取り込み）→選択オブジェクトへ適用 / Choose style, import if needed, apply to selection */
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

        // 適用するスタイルを選択 / Choose the style to apply
        var styleName = chooseGraphicStyleName();
        if (!styleName) return;

        // スタイルを取得（未登録なら取り込む）/ Get the style (import if needed)
        var graphicStyle = ensureGraphicStyle(destinationDoc, styleName);
        if (!graphicStyle) return;

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
