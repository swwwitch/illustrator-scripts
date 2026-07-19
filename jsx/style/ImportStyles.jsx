#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- ダイアログで候補リストから AI ファイルを 1 つ選択
- 指定 AI を開き、作業アートボード内のオブジェクトをコピーして閉じる
- 元ドキュメントへ貼り付け（貼り付け先は新規レイヤー「// _imported」／なければ自動作成）
- カテゴリ（「スタイル／ブラシ／シンボル」「フォント」）を上部ラジオで切替
- 検索欄＋［検索］ボタンで「コンテンツ／ファイル名」を対象に絞り込み
- ListBox（2 列・ヘッダ付き、スクロール可）で候補を表示／ダブルクリックまたは Enter で決定

### 処理の流れ

1. ダイアログで候補から AI ファイルを選択
2. 対象 AI を開いて作業アートボード内をコピー
3. 対象 AI を保存せずに閉じる
4. 元ドキュメントへ復帰し、取り込み先レイヤー（// _imported）を用意
5. 貼り付けを実行

### note

- Illustrator の仕様により、グラフィックスタイル／ブラシは最低 1 件が残ります
- 候補は `ImportStyles_candidates.tsv`（UTF-8 / label, path, category）で永続化されます（旧 check 欄は読み込み時に無視）

### GitHub

https://github.com/swwwitch/illustrator-scripts

### 更新履歴

- v1.5.0 (20260701) : ローカライズを構造化（ネスト LABELS ＋ ドット区切り L()）、全体を IIFE 化、パネル共通設定（setupPanel）を追加、変数名・関数名を整理、追加フローを関数分割、重複コード・不要な try を削減
- v1.4 (20250815) : 標準 ListBox（2 列ヘッダ）へ刷新、削除オプションを廃止、カテゴリラジオ＆検索ボタンを追加、貼り付け先を「// _imported」に統一、ドキュメント刷新
- v1.3 (20250815) : カテゴリ列（スタイル／ブラシ／シンボル／フォント）対応、追加時にカテゴリ選択（ドロップダウン）
- v1.2 (20250815) : CANDIDATES を外部 TSV から読み込み可能に
- v1.1 (20250815) : CANDIDATES に削除オプションを記録
- v1.0 (20250814) : 初期バージョン

*/

/*

### Overview

- Choose one AI file from a dialog list
- Open the chosen AI, copy objects from the working artboard, then close it (without saving)
- Paste into the original document (destination layer: "// _imported"; auto-created if missing)
- Switch category via top radio buttons ("Style/Brush/Symbol" or "Font")
- Filter by both Content / Filename using the search field + [Search] button
- Display candidates in a headered 2-column ListBox (scrollable); confirm with double-click or Enter

### Process Flow

1. Select an AI file from the dialog
2. Open it and copy objects from the working artboard
3. Close the AI without saving
4. Return to the original document and prepare the destination layer (// _imported)
5. Paste

### note

- Due to Illustrator specs, at least one Graphic Style / Brush will remain
- Candidates are persisted in `ImportStyles_candidates.tsv` (UTF-8 / label, path, category). Legacy `check` column is ignored on load

### GitHub

https://github.com/swwwitch/illustrator-scripts

### Changelog

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
var SCRIPT_NAME     = "ImportStyles";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* スタイル用 AI ファイルのベースディレクトリ（環境に合わせて変更）/ Base directory for style AI files */
    var STYLES_DIRECTORY = "~/sw Dropbox/takano masahiro/sync-setting/ai/styles/";

    /* 候補リストの外部ファイル（TSV）/ External candidates TSV
       フォーマット: label \t path \t category  ※旧形式の check(0|1) は読み込み時に無視（後方互換） */
    var CANDIDATES_FILE = File(Folder(STYLES_DIRECTORY).fsName + "/ImportStyles_candidates.tsv");

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
            styleBrushSymbol: { ja: "スタイル／ブラシ／シンボル", en: "Style/Brush/Symbol" },
            font: { ja: "フォント", en: "Fonts" },
            panelTitle: { ja: "カテゴリー", en: "Category" }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            add: { ja: "追加", en: "Add" },
            load: { ja: "読み込み", en: "Load" }
        },
        /* 入力プロンプト / Prompts */
        prompt: {
            pickAi: { ja: "追加するAIファイルを選択", en: "Choose an AI file to add" },
            enterLabel: { ja: "表示される項目名", en: "Display name for the dialog" },
            pickCategory: { ja: "読み込み設定", en: "Import Settings" }
        },
        /* 実行時メッセージ / Runtime messages */
        message: {
            openDocFirst: {
                ja: "元のドキュメントを開いてから実行してください。",
                en: "Please open the destination document first."
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
            }
        },
        /* エラー / Errors */
        error: {
            createDir: {
                ja: "保存先フォルダを作成できませんでした。\n",
                en: "Could not create directory:\n"
            },
            copyFile: {
                ja: "ファイルをコピーできませんでした。\n",
                en: "Could not copy file to:\n"
            },
            candSaveFolder: {
                ja: "候補TSVの保存先を作成できませんでした。\n",
                en: "Could not create folder:\n"
            },
            candOpenWrite: {
                ja: "候補TSVを書き込み用に開けませんでした。\n",
                en: "Could not open candidates TSV for writing:\n"
            },
            candSaveError: {
                ja: "候補TSVの保存中にエラーが発生しました。\n",
                en: "Error saving candidates TSV:\n"
            },
            addGeneral: {
                ja: "候補の追加中にエラーが発生しました。\n",
                en: "An error occurred while adding a candidate.\n"
            },
            uiRefresh: {
                ja: "UIの更新に失敗しました：\n",
                en: "UI refresh failed:\n"
            },
            dialogOpen: {
                ja: "ダイアログを開けませんでした：\n",
                en: "Dialog failed to open:\n"
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
    // レイアウト / Layout
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // 候補データ / Candidate Data
    // =========================================

    /* カテゴリ名の正規化値 / Canonical category names */
    var CATEGORY_STYLE = "スタイル／ブラシ／シンボル";
    var CATEGORY_FONT = "フォント";

    /* 既定候補（TSV 未提供時のフォールバック）/ Default candidates (fallback when no TSV) */
    var DEFAULT_CANDIDATES = [
        { label: "オープンパス", path: "オープンパス.ai", category: CATEGORY_STYLE },
        { label: "矢印", path: "矢印.ai", category: CATEGORY_STYLE },
        { label: "丸数字", path: "丸数字.ai", category: CATEGORY_STYLE },
        { label: "定番フォント（欧文）", path: "fonts.ai", category: CATEGORY_FONT }
    ];

    /* カテゴリ名の正規化（フォント系だけ判定し、他はすべてスタイル扱い）/ Normalize category (font-like → Font, else Style) */
    function normalizeCategory(raw) {
        var text = String(raw != null ? raw : "");
        var compact = text.toLowerCase().replace(/\s+/g, "");
        var isFont = /font/.test(compact) || text.indexOf("フォント") !== -1 || text.indexOf("フォンツ") !== -1;
        return isFont ? CATEGORY_FONT : CATEGORY_STYLE;
    }

    /* 候補リストをTSVから読み込み（無ければ既定を返す）/ Load candidates from TSV (fallback to defaults) */
    function loadCandidates(defaults) {
        var file = CANDIDATES_FILE;
        if (!file.exists) return defaults.slice(0);

        file.encoding = "UTF-8";
        if (!file.open("r")) return defaults.slice(0);

        var loaded = [];
        while (!file.eof) {
            var line = file.readln();
            if (!line || line.charAt(0) === "#") continue;
            var parts = line.split("\t");
            if (parts.length < 2) continue;

            // 旧形式: label, path, check(0|1), category / 新形式: label, path, category
            var categoryIndex = (parts.length >= 3 && /^(0|1)$/.test(parts[2])) ? 3 : 2;
            var rawCategory = (parts.length > categoryIndex && parts[categoryIndex] !== "") ? parts[categoryIndex] : CATEGORY_STYLE;

            loaded.push({
                label: parts[0],
                path: parts[1],
                category: normalizeCategory(rawCategory)
            });
        }
        file.close();
        return loaded.length ? loaded : defaults.slice(0);
    }

    /* 現在の候補をTSVに保存（上書き）/ Save current candidates to TSV (overwrite) */
    function saveCandidates(candidates) {
        var file = CANDIDATES_FILE;
        var parent = file.parent;
        if (parent && !parent.exists && !parent.create()) {
            alert(L("error.candSaveFolder") + toDisplayPath(parent.fsName));
            return false;
        }
        file.encoding = "UTF-8";
        if (!file.open("w")) {
            alert(L("error.candOpenWrite") + toDisplayPath(file.fsName));
            return false;
        }
        try {
            file.writeln("# label\tpath\tcategory");
            for (var i = 0; i < candidates.length; i++) {
                var candidate = candidates[i];
                if (!candidate || !candidate.path) continue;
                var label = (candidate.label != null) ? candidate.label : "";
                var category = (candidate.category != null && candidate.category !== "") ? candidate.category : CATEGORY_STYLE;
                file.writeln(label + "\t" + candidate.path + "\t" + category);
            }
        } catch (e) {
            alert(L("error.candSaveError") + e);
            return false;
        } finally {
            file.close();
        }
        return true;
    }

    /* 候補を読み込み、旧称カテゴリを統一 / Load candidates and normalize legacy category labels */
    var CANDIDATES = loadCandidates(DEFAULT_CANDIDATES);
    for (var candidateIndex = 0; candidateIndex < CANDIDATES.length; candidateIndex++) {
        if (CANDIDATES[candidateIndex]) {
            CANDIDATES[candidateIndex].category = normalizeCategory(CANDIDATES[candidateIndex].category);
        }
    }

    /* パスで候補を検索し index を返す（無ければ -1）/ Find candidate index by path (-1 if none) */
    function indexOfCandidateByPath(path) {
        for (var i = 0; i < CANDIDATES.length; i++) {
            if (CANDIDATES[i] && CANDIDATES[i].path === path) return i;
        }
        return -1;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* 取り込み先レイヤーを取得（無ければ作成）/ Get or create the destination layer "// _imported" */
    function getOrCreateImportLayer(doc) {
        var layerName = "// _imported";
        var importLayer;
        try {
            importLayer = doc.layers.getByName(layerName);
        } catch (e) {
            importLayer = doc.layers.add();
            importLayer.name = layerName;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    /* パスからファイル名を取得（文字化け対策で decodeURI）/ Get decoded filename from path */
    function toFilename(path) {
        try {
            return decodeURI(new File(path).name);
        } catch (e) {
            return path;
        }
    }

    /* ファイル名だけを decodeURI（拡張子付きの生ファイル名向け）/ Decode a raw filename */
    function decodeFileName(name) {
        try {
            return decodeURI(name);
        } catch (e) {
            return name;
        }
    }

    /* 表示用にパスをデコード / Decode path for display */
    function toDisplayPath(path) {
        var text = String(path);
        // %E3%81%AA といった %xx を含む場合のみ decode / Decode only when percent-encoded
        if (/%[0-9A-Fa-f]{2}/.test(text)) {
            try {
                return decodeURI(text);
            } catch (e) {}
        }
        return text;
    }

    /* ファイル操作ヘルパ / File operations helper */
    var fileOps = {
        // ディレクトリを確保（存在しなければ作成）/ Ensure directory exists (create if missing)
        ensureDir: function(folder) {
            if (!folder) return false;
            if (folder.exists) return true;
            if (folder.create()) return true;
            alert(L("error.createDir") + toDisplayPath(folder.fsName));
            return false;
        },
        // src を dest に同名上書きコピー / Copy with overwrite
        copyOverwrite: function(src, dest) {
            if (!src || !dest) return false;
            if (src.fsName === dest.fsName) return true; // 同一ファイルなら何もしない / Same file
            if (dest.exists) dest.remove();
            if (src.copy(dest.fsName)) return true;
            alert(L("error.copyFile") + toDisplayPath(dest.fsName));
            return false;
        }
    };

    // =========================================
    // 候補追加フロー / Add-candidate Flow
    // =========================================

    /* 表示名とカテゴリを入力させる（キャンセルで null）/ Ask for label + category (null on cancel) */
    function promptLabelAndCategory(baseName) {
        var addDialog = new Window("dialog", L("prompt.pickCategory") + " " + SCRIPT_VERSION);
        addDialog.orientation = "column";
        addDialog.alignChildren = ["fill", "top"];

        // ラベル入力 / Label field
        var labelFieldGroup = addDialog.add("group");
        labelFieldGroup.orientation = "column";
        labelFieldGroup.alignChildren = ["fill", "top"];
        labelFieldGroup.add("statictext", undefined, labelText("prompt.enterLabel"));
        var labelField = labelFieldGroup.add("edittext", undefined, baseName);
        labelField.characters = 20;

        // カテゴリ選択（パネル＋ラジオ）/ Category panel with radios
        var categoryPanel = addDialog.add("panel", undefined, L("category.panelTitle"));
        setupPanel(categoryPanel, 6);
        var styleRadio = categoryPanel.add("radiobutton", undefined, L("category.styleBrushSymbol"));
        var fontRadio = categoryPanel.add("radiobutton", undefined, L("category.font"));
        // ファイル名からカテゴリを自動推定 / Auto-detect category from filename
        var isFontFile = (baseName.toLowerCase().indexOf("font") !== -1) || (baseName.indexOf("フォント") !== -1);
        fontRadio.value = isFontFile;
        styleRadio.value = !isFontFile;

        // ボタン行 / Buttons (centered)
        var buttonGroup = addDialog.add("group");
        buttonGroup.alignment = ["center", "bottom"];
        buttonGroup.alignChildren = ["center", "center"];
        buttonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, L("button.load"), { name: "ok" });

        if (addDialog.show() !== 1) return null;

        // 入力値整形 / Sanitize label
        var cleanLabel = String(labelField.text || "").replace(/\r?\n/g, " ").replace(/\t/g, " ").replace(/^\s+|\s+$/g, "");
        if (cleanLabel === "") cleanLabel = baseName;

        return {
            label: cleanLabel,
            category: (fontRadio.value === true) ? CATEGORY_FONT : CATEGORY_STYLE
        };
    }

    /* AIファイルを取り込んで候補に登録（追加した path を返す／中断で null）/ Import an AI file and register it */
    function importCandidateFromFile() {
        try {
            // 1) ファイル選択 / Pick an AI file
            var pickedFile = File.openDialog(L("prompt.pickAi"), "*.ai");
            if (!pickedFile) return null;

            // 2) 保存先フォルダを確保 / Ensure destination folder
            var stylesFolder = Folder(STYLES_DIRECTORY);
            if (!fileOps.ensureDir(stylesFolder)) return null;

            // 3) 配下へコピー（同名は上書き）/ Copy into folder (overwrite)
            var destFile = File(stylesFolder.fsName + "/" + pickedFile.name);
            if (!fileOps.copyOverwrite(pickedFile, destFile)) return null;

            // 4) 表示名とカテゴリを入力 / Ask for label + category
            var baseName = decodeFileName(pickedFile.name).replace(/\.[^\.]+$/, "");
            var input = promptLabelAndCategory(baseName);
            if (!input) return null;

            // 5) 既存候補を上書き or 追加 / Overwrite existing or append
            var candidate = { label: input.label, path: destFile.name, category: input.category };
            var existingIndex = indexOfCandidateByPath(destFile.name);
            if (existingIndex >= 0) {
                CANDIDATES[existingIndex] = candidate;
            } else {
                CANDIDATES.push(candidate);
            }

            // 6) TSVへ保存 / Save to TSV
            if (!saveCandidates(CANDIDATES)) return null;

            alert((existingIndex >= 0 ? L("message.overwritten") : L("message.added")) + "\n" +
                input.label + "  (" + toFilename(destFile.name) + ")" + L("message.synced"));
            return destFile.name;
        } catch (e) {
            alert(L("error.addGeneral") + "\n" + e);
            return null;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 候補から1つ選択し、フルパスなどを返す / Show a dialog to choose one candidate, return its full path or null */
    function chooseFileFromList(candidates) {
        var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);

        // 位置＆透明度 / Position & opacity
        var offsetX = 300;
        try {
            dialog.opacity = 0.97;
        } catch (e) {}
        dialog.onShow = function() {
            try {
                dialog.location = [dialog.location[0] + offsetX, dialog.location[1]];
            } catch (e) {}
        };

        // 本体レイアウト / Main layout
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        // ---- 上部カテゴリ選択ラジオ / Category filter (top) ----
        var categoryFilterGroup = dialog.add("group");
        categoryFilterGroup.orientation = "row";
        categoryFilterGroup.alignChildren = ["left", "center"];
        categoryFilterGroup.alignment = ["center", "top"]; // ダイアログ中央揃え / Center within dialog
        var styleFilterRadio = categoryFilterGroup.add("radiobutton", undefined, L("category.styleBrushSymbol"));
        var fontFilterRadio = categoryFilterGroup.add("radiobutton", undefined, L("category.font"));
        // 既定ではどちらも未選択（全件表示）/ No filter by default

        // ---- 上部検索フィルタ / Search filter (top) ----
        var searchGroup = dialog.add("group");
        searchGroup.orientation = "row";
        searchGroup.alignChildren = ["center", "center"];
        searchGroup.alignment = ["center", "top"]; // ダイアログ左右中央に配置 / Center horizontally
        searchGroup.margins = [0, 5, 0, 15]; // 上下マージンを追加 / Add vertical margins

        var searchField = searchGroup.add("edittext", undefined, "");
        searchField.characters = 24;
        searchField.helpTip = L("search.placeholder");

        // 検索ボタン（クリック時のみフィルタ適用）/ Search button (apply filter only on click)
        var searchButton = searchGroup.add("button", undefined, L("search.label"));
        searchButton.onClick = function() {
            currentSearchText = searchField.text; // 現在のテキストを適用 / Apply current text
            relayout();
        };

        // 標準 ListBox（ヘッダ付き2列）/ Standard ListBox with 2 columns
        var listColumnWidths = { content: 140, filename: 180 };
        var candidateList = dialog.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [L("column.content"), L("column.filename")],
            columnWidths: [listColumnWidths.content, listColumnWidths.filename]
        });
        candidateList.alignment = ["fill", "fill"];
        candidateList.preferredSize = [listColumnWidths.content + listColumnWidths.filename + 20, 200];

        // ダブルクリックで決定 / Double-click to confirm
        candidateList.onDoubleClick = function() {
            activateRow();
        };

        /* 選択行を確定してダイアログを閉じる / Confirm the selected row */
        function activateRow() {
            if (candidateList && candidateList.selection) {
                dialog.close(1);
            }
        }

        // ===== 検索・カテゴリフィルタの状態 / Filter state =====
        var currentCategory = null; // null = 全件 / null = all
        var currentSearchText = "";

        /* カテゴリ反映ヘルパ / Category apply helper */
        function applyCategoryFilter(categoryName) {
            currentCategory = categoryName;
            relayout();
        }
        styleFilterRadio.onClick = function() {
            applyCategoryFilter(CATEGORY_STYLE);
        };
        fontFilterRadio.onClick = function() {
            applyCategoryFilter(CATEGORY_FONT);
        };

        // 検索欄はボタン押下で適用（ライブでは反映しない）/ Apply only on button click
        searchField.onChanging = function() {
            /* no live filtering */
        };

        /* レイアウト一括再計算 / Unified relayout */
        function relayout() {
            try {
                recomputeRowVisibility();
            } catch (e) {
                alert(L("error.uiRefresh") + e);
            }
            try {
                dialog.layout.layout(true);
            } catch (e2) {}
        }

        /* 指定パスの行を選択 / Select row by candidate path */
        function selectByPath(path) {
            var items = candidateList.items;
            for (var i = 0; i < items.length; i++) {
                if (items[i] && items[i]._candidate && items[i]._candidate.path === path) {
                    candidateList.selection = i;
                    return;
                }
            }
        }

        // キーボード：Enterで決定（上下移動はListBox標準）/ Keyboard: Enter confirms
        if (dialog.addEventListener) {
            dialog.addEventListener("keydown", function(evt) {
                try {
                    if (evt.keyName === "Enter") activateRow();
                } catch (e) {}
            });
        }

        /* 文字列一致（大文字小文字を無視）/ Case-insensitive contains */
        function containsCI(haystack, needle) {
            return String(haystack).toLowerCase().indexOf(String(needle).toLowerCase()) !== -1;
        }

        /* 行の表示/非表示を再計算（カテゴリ＋テキスト）/ Recompute row visibility (category + text) */
        function recomputeRowVisibility() {
            candidateList.removeAll();
            if (!candidates || !candidates.length) return;

            var query = currentSearchText;
            for (var i = 0; i < candidates.length; i++) {
                var candidate = candidates[i];
                if (!candidate) continue;
                var labelValue = (candidate.label != null) ? String(candidate.label) : "";
                var pathValue = (candidate.path != null) ? String(candidate.path) : "";
                var categoryValue = (candidate.category != null && candidate.category !== "") ? String(candidate.category) : CATEGORY_STYLE;

                if (currentCategory != null && categoryValue !== currentCategory) continue;

                if (query && query.length > 0) {
                    var matched = containsCI(labelValue, query) || containsCI(toFilename(pathValue), query);
                    if (!matched) continue;
                }

                var item = candidateList.add("item", labelValue);
                try {
                    item.subItems[0].text = toFilename(pathValue);
                } catch (eSub) {}
                item._candidate = candidate; // 参照を保持 / Keep reference
                item._path = STYLES_DIRECTORY + pathValue; // 絶対パス / Absolute path
            }
            if (candidateList.items.length > 0) candidateList.selection = 0;
        }

        // ボタン行：左(キャンセル) - スペーサー - 右(追加 / 読み込み) / Button row
        var buttonRow = dialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.margins = [0, 10, 0, 0];

        var cancelGroup = buttonRow.add("group");
        cancelGroup.orientation = "row";
        var cancelButton = cancelGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        cancelButton.onClick = function() {
            dialog.close(0);
        };

        var spacer = buttonRow.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize = [0, 0];

        var actionGroup = buttonRow.add("group");
        actionGroup.alignChildren = ["right", "center"];

        // 候補を追加（AIファイルをコピーしてTSVへ登録）/ Add a candidate
        var addButton = actionGroup.add("button", undefined, L("button.add"), "");
        addButton.onClick = function() {
            var addedPath = importCandidateFromFile();
            if (!addedPath) return;
            relayout();
            selectByPath(addedPath);
        };

        // 読み込み（選択を確定）/ Load (confirm selection)
        var loadButton = actionGroup.add("button", undefined, L("button.load"), { name: "ok" });
        loadButton.onClick = function() {
            dialog.close(1);
        };

        // 初期表示 → ダイアログ表示 → 選択を返却 / Init, show, return selection
        relayout();
        if (dialog.show() !== 1) return null;
        var selectedItem = candidateList.selection;
        if (!selectedItem) return null;
        return { path: selectedItem._path };
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================

    /* 候補選択→対象AIをコピー→元ドキュメントへ貼り付け / Choose, copy from AI, paste into original */
    function main() {
        var choice;
        try {
            choice = chooseFileFromList(CANDIDATES);
        } catch (e) {
            alert(L("error.dialogOpen") + e);
            return;
        }
        if (!choice) return;

        // ドキュメントの存在を確認 / Check for an open document
        if (app.documents.length === 0) {
            alert(L("message.openDocFirst"));
            return;
        }
        var originalDoc = app.activeDocument;

        // 指定ファイルを開く / Open the chosen file
        var styleFile = new File(choice.path);
        if (!styleFile.exists) {
            alert(L("message.fileNotFoundTitle") + "\n" + L("message.fileNotFoundBody") + toFilename(choice.path));
            return;
        }
        var styleDoc = app.open(styleFile);

        // 作業アートボード内の全てをコピーして保存せず閉じる / Copy in-artboard objects, close without saving
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        styleDoc.close(SaveOptions.DONOTSAVECHANGES);

        // 取り込み先レイヤーを用意してからペースト / Prepare destination layer then paste
        app.activeDocument = originalDoc;
        originalDoc.activeLayer = getOrCreateImportLayer(originalDoc);
        app.executeMenuCommand("paste");
    }

    main();

})();
