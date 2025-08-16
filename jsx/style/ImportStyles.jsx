#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

ImportStyles.jsx — 作業アートボードからコピーペースト（ダイアログ付）

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- ダイアログで候補リストから AI ファイルを 1 つ選択
- 指定 AI を開き、作業アートボード内のオブジェクトをコピーして閉じる
- 元ドキュメントへ貼り付け（貼り付け先は新規レイヤー「// _imported」／なければ自動作成）
- カテゴリ（「スタイル／ブラシ／シンボル」「フォント」）を上部ラジオで切替
- 検索欄＋［検索］ボタンで「コンテンツ／ファイル名」を対象に絞り込み
- ListBox（2 列・ヘッダ付き、スクロール可）で候補を表示／ダブルクリックまたは Enter で決定

### 主な機能：

- 候補ファイルのダイアログ選択（外部 TSV から読み込み・保存）
- 取り込み先レイヤーの自動準備（「// _imported」）
- カテゴリ切替（ラジオボタン）と検索によるフィルタリング
- キーボード操作（↑↓：ListBox 標準、Enter：決定）

### 処理の流れ：

1. ダイアログで候補から AI ファイルを選択
2. 対象 AI を開いて作業アートボード内をコピー
3. 対象 AI を保存せずに閉じる
4. 元ドキュメントへ復帰し、取り込み先レイヤー（// _imported）を用意
5. 貼り付けを実行

### note：

- Illustrator の仕様により、グラフィックスタイル／ブラシは最低 1 件が残ります
- 候補は `ImportStyles_candidates.tsv`（UTF-8 / label, path, category）で永続化されます（旧 check 欄は読み込み時に無視）

### 更新履歴：

- v1.4 (20250815) : 標準 ListBox（2 列ヘッダ）へ刷新、削除オプションを廃止、カテゴリラジオ＆検索ボタンを追加、貼り付け先を「// _imported」に統一、ドキュメント刷新
- v1.3 (20250815) : カテゴリ列（スタイル／ブラシ／シンボル／フォント）対応、追加時にカテゴリ選択（ドロップダウン）
- v1.2 (20250815) : CANDIDATES を外部 TSV から読み込み可能に
- v1.1 (20250815) : CANDIDATES に削除オプションを記録
- v1.0 (20250814) : 初期バージョン

----

### Script Name:

ImportStyles.jsx — Copy from Working Artboard and Import Styles (with Dialog)

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Choose one AI file from a dialog list
- Open the chosen AI, copy objects from the working artboard, then close it (without saving)
- Paste into the original document (destination layer: "// _imported"; auto-created if missing)
- Switch category via top radio buttons ("Style/Brush/Symbol" or "Font")
- Filter by both Content / Filename using the search field + [Search] button
- Display candidates in a headered 2-column ListBox (scrollable); confirm with double-click or Enter

### Key Features:

- Dialog selection from candidates (load / save via external TSV)
- Auto-prepare destination layer ("// _imported")
- Category toggle (radio) and search filtering
- Keyboard: Up/Down by ListBox default, Enter to confirm

### Process Flow:

1. Select an AI file from the dialog
2. Open it and copy objects from the working artboard
3. Close the AI without saving
4. Return to the original document and prepare the destination layer (// _imported)
5. Paste

### note:

- Due to Illustrator specs, at least one Graphic Style / Brush will remain
- Candidates are persisted in `ImportStyles_candidates.tsv` (UTF-8 / label, path, category). Legacy `check` column is ignored on load

### Changelog:

- v1.4 (2025-08-15): Switched to standard 2-column ListBox, removed delete option, added category radios & Search button, unified destination layer to "// _imported", refreshed docs
- v1.3 (2025-08-15): Added Category column (Style/Brush/Font) and category dropdown on Add
- v1.2 (2025-08-15): Load CANDIDATES from external TSV file
- v1.1 (2025-08-15): Added delete option in CANDIDATES to remember after loading
- v1.0 (2025-08-14): Initial release

*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.4";

/* =========================
   設定 / Settings
   ========================= */

// ベースディレクトリ / Base directory
// ※ここはユーザーの環境に合わせて変更してください / Change this to your environment
var DIRECTORY = "~/sw Dropbox/takano masahiro/sync-setting/ai/styles/";

// 候補リストの外部ファイル（TSV）/ External candidates TSV
// フォーマット: label \t path \t category  ※旧形式の check(0|1) は読み込み時に無視（後方互換）
var CANDIDATES_FILE = File(Folder(DIRECTORY).fsName + "/ImportStyles_candidates.tsv");

// 現在のロケールを取得 / Get current locale
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var LANG = getCurrentLang();

var LABELS = {
    /* ===== Dialog & Header / ダイアログ & ヘッダ ===== */
    dialogTitle: {
        ja: "スタイル読み込み：ファイル選択 " + SCRIPT_VERSION,
        en: "Import Styles: Choose File " + SCRIPT_VERSION
    },
    columnContent: {
        ja: "コンテンツ",
        en: "Content"
    },
    columnFilename: {
        ja: "ファイル名",
        en: "Filename"
    },

    /* ===== Search (top) / 検索（上部） ===== */
    searchLabel: {
        ja: "検索",
        en: "Search"
    },
    searchPlaceholder: {
        ja: "コンテンツ / ファイル名で絞り込み",
        en: "Filter by Content / Filename"
    },

    /* ===== Category Radios / カテゴリ選択ラジオ ===== */
    optCategoryStyle: {
        ja: "スタイル／ブラシ／シンボル",
        en: "Style/Brush/Symbol"
    },
    optCategoryFont: {
        ja: "フォント",
        en: "Fonts"
    },

    /* ===== Buttons (footer) / フッターボタン ===== */
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    btnAdd: {
        ja: "追加",
        en: "Add"
    },
    btnLoad: {
        ja: "読み込み",
        en: "Load"
    },

    /* ===== Prompts (Add flow) / 追加フローのプロンプト ===== */
    dlgPickAi: {
        ja: "追加するAIファイルを選択",
        en: "Choose an AI file to add"
    },
    promptEnterLabel: {
        ja: "表示される項目名",
        en: "Display name for the dialog"
    },
    dlgPickCategory: {
        ja: "読み込み設定",
        en: "Import Settings"
    },

    /* ===== Runtime messages / 実行時メッセージ ===== */
    msgOpenDocFirst: {
        ja: "元のドキュメントを開いてから実行してください。",
        en: "Please open the destination document first."
    },
    msgFileNotFoundTitle: {
        ja: "ファイルが見つかりません",
        en: "File Not Found"
    },
    msgFileNotFoundBody: {
        ja: "指定したファイルが見つかりません：\n",
        en: "The specified file was not found:\n"
    },

    /* ===== Success (Add) / 成功（追加・上書き） ===== */
    msgAdded: {
        ja: "候補を追加しました：",
        en: "Candidate added:"
    },
    msgOverwritten: {
        ja: "候補を上書きしました：",
        en: "Candidate overwritten:"
    },
    msgSynced: {
        ja: "\n\nImportStyles_candidates.tsv と現在のダイアログに反映しました。",
        en: "\n\nReflected in ImportStyles_candidates.tsv and the current dialog."
    },

    /* ===== Errors (file/dir ops) / エラー（ファイル・フォルダ操作） ===== */
    errCreateDir: {
        ja: "保存先フォルダを作成できませんでした。\n",
        en: "Could not create directory:\n"
    },
    errCopyFile: {
        ja: "ファイルをコピーできませんでした。\n",
        en: "Could not copy file to:\n"
    },
    errCopying: {
        ja: "ファイルのコピー中にエラーが発生しました。\n",
        en: "Error copying file:\n"
    },
    errCandSaveFolder: {
        ja: "候補TSVの保存先を作成できませんでした。\n",
        en: "Could not create folder:\n"
    },
    errCandOpenWrite: {
        ja: "候補TSVを書き込み用に開けませんでした。\n",
        en: "Could not open candidates TSV for writing:\n"
    },
    errCandSaveError: {
        ja: "候補TSVの保存中にエラーが発生しました。\n",
        en: "Error saving candidates TSV:\n"
    },
    errAddGeneral: {
        ja: "候補の追加中にエラーが発生しました。\n",
        en: "An error occurred while adding a candidate.\n"
    }
};

// 既定候補（ファイル未提供時のフォールバック）/ Default candidates (fallback)

var DEFAULT_CANDIDATES = [{
        label: "オープンパス",
        path: "オープンパス.ai",
        category: "スタイル／ブラシ／シンボル"
    },
    {
        label: "矢印",
        path: "矢印.ai",
        category: "スタイル／ブラシ／シンボル"
    },
    {
        label: "丸数字",
        path: "丸数字.ai",
        category: "スタイル／ブラシ／シンボル"
    },
    {
        label: "定番フォント（欧文）",
        path: "fonts.ai",
        category: "フォント"
    }
];

/* カテゴリ名の正規化 / Normalize category label (JP/EN legacy variants) */
function normalizeCategory(raw) {
    // Defensive normalization
    var s = String(raw != null ? raw : '').toLowerCase();
    // Remove spaces and common separators for comparison
    var compact = s.replace(/\s+/g, '').replace(/[／\/|]/g, '/');

    // Heuristics: treat anything that looks like "font(s)" or contains the JP word as Fonts
    var isFonts = /font/.test(compact) || /\u30d5\u30a9\u30f3\u30c8/.test(raw) || /\u30d5\u30a9\u30f3\u30c4/.test(raw) || /\u30d5\u30a9\u30f3\u30c8\u985e/.test(raw) || /\u30d5\u30a9\u30f3\u30c8s?/.test(compact);
    if (isFonts) return 'フォント';

    // English legacy labels for the other bucket → unify
    // e.g., "style", "styles", "brush", "symbol", any combination
    var isStyleBrushSymbol = /style|brush|symbol/.test(compact);
    if (isStyleBrushSymbol) return 'スタイル／ブラシ／シンボル';

    // Japanese variants that should collapse to Style/Brush/Symbol
    if (/\u30b9\u30bf\u30a4\u30eb/.test(raw) || /\u30d6\u30e9\u30b7/.test(raw) || /\u30b7\u30f3\u30dc\u30eb/.test(raw)) {
        return 'スタイル／ブラシ／シンボル';
    }

    // Fallback: default bucket is Style/Brush/Symbol
    return 'スタイル／ブラシ／シンボル';
}

/* 取り込み先レイヤーを取得 / Get or create the destination layer (create if missing) */
// - 名前: "// _imported"

function getOrCreateImportLayer(doc) {
    var LNAME = "// _imported";
    try {
        var lyr = null;
        try {
            lyr = doc.layers.getByName(LNAME);
        } catch (eFind) {
            lyr = doc.layers.add();
            lyr.name = LNAME;
        }
        try {
            lyr.locked = false;
        } catch (e1) {}
        try {
            lyr.visible = true;
        } catch (e2) {}
        return lyr;
    } catch (e) {
        // フォールバック：新規レイヤー作成
        var fallback = doc.layers.add();
        try {
            fallback.name = LNAME;
        } catch (_) {}
        try {
            fallback.locked = false;
        } catch (_) {}
        try {
            fallback.visible = true;
        } catch (_) {}
        return fallback;
    }
}

function loadCandidates(defaults) {
    try {
        var arr = [];
        var f = CANDIDATES_FILE;
        if (!f.exists) {
            return defaults.slice(0); // ファイルが無ければ既定を返す
        }
        f.encoding = "UTF-8";
        if (!f.open('r')) return defaults.slice(0);
        while (!f.eof) {
            var line = f.readln();
            if (!line || line.charAt(0) === '#') continue;
            var parts = line.split('\t');
            if (parts.length >= 2) {
                var label = parts[0];
                var path = parts[1];

                // 旧形式: label, path, check(0|1), category
                // 新形式: label, path, category
                var idxCat = 2;
                if (parts.length >= 3 && /^(0|1)$/.test(parts[2])) {
                    // parts[2] は旧 check フィールド。category はその次。
                    idxCat = 3;
                }
                var rawCat = (parts.length > idxCat && parts[idxCat] !== "") ? parts[idxCat] : "スタイル／ブラシ／シンボル";
                var cat = normalizeCategory(rawCat);

                // check はファイルには保存しないため、カテゴリから導出（フォント=0 / それ以外=1）
                var chk = (cat === 'フォント') ? 0 : 1;

                arr.push({
                    label: label,
                    path: path,
                    check: chk,
                    category: cat
                });
            }
        }
        f.close();
        return arr.length ? arr : defaults.slice(0);
    } catch (e) {
        return defaults.slice(0);
    }
}

// 現在の候補をTSVに保存（上書き）/ Save current candidates to TSV (overwrite)
function saveCandidates(cands) {
    try {
        var cf = CANDIDATES_FILE;
        var parent = cf.parent;
        if (parent && !parent.exists && !parent.create()) {
            alert(T('errCandSaveFolder') + toDisplayPath(parent.fsName));
            return false;
        }
        cf.encoding = "UTF-8";
        if (!cf.open('w')) {
            alert(T('errCandOpenWrite') + toDisplayPath(cf.fsName));
            return false;
        }
        cf.writeln("# label\tpath\tcategory");
        for (var i = 0; i < cands.length; i++) {
            var ci = cands[i];
            if (!ci || !ci.path) continue;
            var lab = (ci.label != null) ? ci.label : "";
            var cat = (ci.category != null && ci.category !== "") ? ci.category : "スタイル／ブラシ／シンボル";
            cf.writeln(lab + "\t" + ci.path + "\t" + cat);
        }
        cf.close();
        return true;
    } catch (e) {
        try {
            alert(T('errCandSaveError') + e);
        } catch (_) {}
        return false;
    }
}

// フルパス化（~ を展開）/ Expand '~' and build absolute path

var CANDIDATES = loadCandidates(DEFAULT_CANDIDATES);
// 旧称を統一（TSVに残っていても表示は統一）/ Normalize legacy category labels
for (var i = 0; i < CANDIDATES.length; i++) {
    if (CANDIDATES[i]) {
        CANDIDATES[i].category = normalizeCategory(CANDIDATES[i].category);
    }
}


// 指定キーのローカライズ文字列を取得 / Get localized label for key
function T(key) {
    return (LABELS[key] && LABELS[key][LANG]) ? LABELS[key][LANG] : ("[" + key + "]");
}

/* =========================
   ユーティリティ / Utilities
   ========================= */


/* パスからファイル名を取得 / Get filename from path */
function toFilename(p) {
    try {
        return decodeURI(new File(p).name);
    } catch (e) {
        return p;
    }
}

/* 表示用にパスをデコード / Decode path for display */
function toDisplayPath(p) {
    try {
        var s = String(p);
        // %E3%81%AAといった %xx を含む場合は decode
        if (/%[0-9A-Fa-f]{2}/.test(s)) {
            try {
                return decodeURI(s);
            } catch (e1) {}
        }
        return s;
    } catch (e) {
        return String(p);
    }
}

/* 表示名（ファイル名のみ）/ Display name (filename only) */
function toDisplayName(n) {
    // ファイル名（パス不要）は既存の toFilename を使用
    return toFilename(n);
}

/* ファイル操作ヘルパ / File operations helper */
var fileOps = {
    // ディレクトリを確保（存在しなければ作成）/ Ensure directory exists (create if missing)
    ensureDir: function(folder) {
        try {
            if (!folder) return false;
            var p = folder.fsName;
            if (folder.exists) return true;
            if (folder.create()) return true;
            alert(T('errCreateDir') + toDisplayPath(p));
            return false;
        } catch (e) {
            try {
                alert(T('errCreateDir') + toDisplayPath(folder && folder.fsName ? folder.fsName : ''));
            } catch (_) {}
            return false;
        }
    },
    // src を dest に同名上書きコピー / Copy with overwrite
    copyOverwrite: function(src, dest) {
        try {
            if (!src || !dest) return false;
            if (src.fsName === dest.fsName) return true; // same file, nothing to do
            if (dest.exists) {
                try {
                    dest.remove();
                } catch (_) {}
            }
            if (src.copy(dest.fsName)) return true;
            alert(T('errCopyFile') + toDisplayPath(dest.fsName));
            return false;
        } catch (e) {
            try {
                alert(T('errCopying') + e);
            } catch (_) {}
            return false;
        }
    }
};

/*
  ダイアログで候補から1つ選択し、フルパスなどを返す
  Show a dialog to choose one candidate, return its full path or null.
*/
function chooseFileFromList(candidates) {
    var dlg = new Window('dialog', T('dialogTitle'));

    // 位置＆透明度 / Position & opacity
    var offsetX = 300;
    var dialogOpacity = 0.97;
    // 既存のヘルパを踏襲（なければインラインで実装）
    try {
        dlg.opacity = dialogOpacity;
    } catch (e) {}
    dlg.onShow = function() {
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + 0];
        } catch (e) {}
    };

    // 本体レイアウト / Main layout
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    // ---- Category filter (top) / 上部カテゴリ選択ラジオ ----
    var filterGroup = dlg.add('group');
    filterGroup.orientation = 'row';
    filterGroup.alignChildren = ['left', 'center'];
    filterGroup.alignment = ['center', 'top']; // ダイアログ中央揃え / Center within dialog
    var rbFilterStyle = filterGroup.add('radiobutton', undefined, T('optCategoryStyle')); // スタイル／ブラシ／シンボル
    var rbFilterFont = filterGroup.add('radiobutton', undefined, T('optCategoryFont')); // フォント
    // 既定ではどちらも未選択（全件表示）/ No filter by default

    // ---- Search filter (top) / 上部検索フィルタ ----
    var searchGroup = dlg.add('group');
    searchGroup.orientation = 'row';
    searchGroup.alignChildren = ['center', 'center'];
    searchGroup.alignment = ['center', 'top']; // ダイアログ左右中央に配置
    searchGroup.margins = [0, 5, 0, 15]; // 上下マージンを追加

    var etSearch = searchGroup.add('edittext', undefined, '');
    etSearch.characters = 24;
    etSearch.helpTip = T('searchPlaceholder');

    // 検索ボタン（クリック時のみフィルタ適用）/ Search button (apply filter only on click)
    var btnSearch = searchGroup.add('button', undefined, T('searchLabel'));
    btnSearch.onClick = function() {
        currentSearchText = etSearch.text; // Apply the current text
        relayout(); // Recompute visibility & layout
    };

    // 標準 ListBox（ヘッダ付き2列）/ Standard ListBox with 2 columns
    var colW = {
        content: 140,
        fname: 180
    };
    var list = dlg.add('listbox', undefined, [], {
        multiselect: false,
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: [T('columnContent'), T('columnFilename')],
        columnWidths: [colW.content, colW.fname]
    });
    list.alignment = ['fill', 'fill'];
    list.preferredSize = [colW.content + colW.fname + 20, 200];

    // ダブルクリックで決定 / Double-click to confirm
    list.onDoubleClick = function() {
        activateRow();
    };

    function activateRow() {
        if (list && list.selection) {
            dlg.close(1);
        }
    }

    // ===== 検索・カテゴリフィルタの状態と再計算 =====
    // 現在のフィルタ状態 / Current filter state
    var currentCategory = null; // null = 全件
    var currentSearchText = '';

    // カテゴリ反映ヘルパ / Category apply helper
    function applyCategoryFilter(categoryName) {
        currentCategory = categoryName; // 'スタイル／ブラシ／シンボル' or 'フォント'
        relayout();
    }
    rbFilterStyle.onClick = function() {
        applyCategoryFilter('スタイル／ブラシ／シンボル');
    };
    rbFilterFont.onClick = function() {
        applyCategoryFilter('フォント');
    };

    // 検索欄はボタン押下で適用（ライブでは反映しない）/ Apply only on button click
    etSearch.onChanging = function() {
        /* no live filtering */
    };

    // レイアウト一括 / Unified relayout
    function relayout() {
        try {
            recomputeRowVisibility();
        } catch (e) {
            try {
                alert("UI refresh failed:\n" + e);
            } catch (_) {}
        }
        try {
            dlg.layout.layout(true);
        } catch (e2) {}
    }

    // 指定パスの行を選択 / Select row by candidate path
    function selectByPath(p) {
        try {
            if (!list || !list.items || list.items.length === 0) return;
            for (var ii = 0; ii < list.items.length; ii++) {
                var it = list.items[ii];
                if (it && it._cand && it._cand.path === p) {
                    list.selection = ii;
                    return;
                }
            }
        } catch (e) {}
    }

    // キーボード：Enterで決定（上下移動はListBox標準）/ Keyboard: Enter confirms
    if (dlg.addEventListener) {
        dlg.addEventListener('keydown', function(evt) {
            try {
                if (evt.keyName === 'Enter') activateRow();
            } catch (e) {}
        });
    }

    // 文字列一致（大文字小文字を無視）/ Case-insensitive contains
    function containsCI(haystack, needle) {
        try {
            return String(haystack).toLowerCase().indexOf(String(needle).toLowerCase()) !== -1;
        } catch (e) {
            return false;
        }
    }

    // 行の表示/非表示を再計算 / Recompute row visibility (category + text)
    function recomputeRowVisibility() {
        try {
            list.removeAll();
        } catch (e) {}

        if (!candidates || !candidates.length) return;
        for (var i = 0; i < candidates.length; i++) {
            var c = candidates[i];
            if (!c) continue;
            var lab = (c.label != null) ? String(c.label) : "";
            var pth = (c.path != null) ? String(c.path) : "";
            var cat = (c.category != null && c.category !== "") ? String(c.category) : "スタイル／ブラシ／シンボル";

            var catOK = (currentCategory == null) ? true : (cat === currentCategory);
            if (!catOK) continue;

            var q = currentSearchText;
            if (q && q.length > 0) {
                var hit = false;
                hit = hit || containsCI(lab, q);
                hit = hit || containsCI(toFilename(pth), q);
                if (!hit) continue;
            }

            var item;
            try {
                item = list.add('item', lab);
            } catch (eAdd) {
                continue;
            }
            try {
                item.subItems[0].text = toFilename(pth);
            } catch (eSub) {}
            item._cand = c; // keep reference
            item._path = DIRECTORY + pth; // absolute path
        }
        try {
            if (list.items.length > 0) list.selection = 0;
        } catch (eSel) {}
    }

    // ボタン行：左(キャンセル) - スペーサー - 右(追加 / 読み込み)
    var btnRowGroup = dlg.add('group');
    btnRowGroup.orientation = 'row';
    btnRowGroup.alignment = ['fill', 'bottom'];
    btnRowGroup.margins = [0, 10, 0, 0];

    var btnLeftGroup = btnRowGroup.add('group');
    btnLeftGroup.orientation = 'row';
    var btnCancel = btnLeftGroup.add('button', undefined, T('btnCancel'), {
        name: 'cancel'
    });
    btnCancel.onClick = function() {
        dlg.close(0);
    };

    var spacer = btnRowGroup.add('group');
    spacer.alignment = ['fill', 'fill'];
    spacer.minimumSize = [0, 0];

    var btnRightGroup = btnRowGroup.add('group');
    btnRightGroup.alignChildren = ['right', 'center'];

    var btnAdd = btnRightGroup.add('button', undefined, T('btnAdd'), '');
    btnAdd.onClick = function() {
        try {
            // 1) ファイル選択
            var picked = File.openDialog(T('dlgPickAi'), '*.ai');
            if (!picked) return;

            // 2) 保存先フォルダ（DIRECTORY）を確保 / Ensure destination folder
            var dirFolder = Folder(DIRECTORY);
            if (!fileOps.ensureDir(dirFolder)) return;

            // 3) DIRECTORY 配下へコピー（同名は上書き）/ Copy into DIRECTORY (overwrite)
            var dest = File(dirFolder.fsName + '/' + picked.name);
            if (!fileOps.copyOverwrite(picked, dest)) return;

            // 4) 表示名とカテゴリを同一ダイアログで入力 / One dialog for label + category
            //    → ファイル名は decodeURI してから拡張子を外す（文字化け対策）
            var base = (function(fname) {
                try {
                    return decodeURI(fname);
                } catch (e) {
                    return fname;
                }
            })(picked.name).replace(/\.[^\.]+$/, '');

            var addDlg = new Window('dialog', T('dlgPickCategory'));
            addDlg.orientation = 'column';
            addDlg.alignChildren = ['fill', 'top'];

            // ラベル入力 / Label field
            var labelGroup = addDlg.add('group');
            labelGroup.orientation = 'column';
            labelGroup.alignChildren = ['fill', 'top'];
            labelGroup.add('statictext', undefined, T('promptEnterLabel'));
            var etLabel = labelGroup.add('edittext', undefined, base);
            etLabel.characters = 20;

            // カテゴリ選択（パネル＋ラジオ）/ Category panel with radios
            var pnlCategory = addDlg.add('panel', undefined, "カテゴリー");
            pnlCategory.orientation = 'column';
            pnlCategory.alignChildren = ['left', 'top'];
            pnlCategory.margins = [15, 20, 15, 10];
            var rbStyle = pnlCategory.add('radiobutton', undefined, T('optCategoryStyle'));
            var rbFont = pnlCategory.add('radiobutton', undefined, T('optCategoryFont'));
            // ファイル名からカテゴリを自動推定 / Auto-detect category from filename
            var fnameForDetect = base.toLowerCase();
            var isFontFile = (fnameForDetect.indexOf('font') !== -1) || (/\u30d5\u30a9\u30f3\u30c8/.test(base)) || (base.indexOf('フォント') !== -1);
            rbFont.value = !!isFontFile;
            rbStyle.value = !rbFont.value; // デフォルトはスタイル、ただしフォント候補ならフォント

            // ボタン行 / Buttons (centered)
            var g = addDlg.add('group');
            g.alignment = ['center', 'bottom'];
            g.alignChildren = ['center', 'center'];
            g.add('button', undefined, T('btnCancel'), {
                name: 'cancel'
            });
            g.add('button', undefined, T('btnLoad'), {
                name: 'ok'
            });

            if (addDlg.show() !== 1) return; // ユーザーがキャンセル

            // 入力値整形 / Sanitize label
            var userLabel = String(etLabel.text || '').replace(/\r?\n/g, ' ').replace(/\t/g, ' ').replace(/^\s+|\s+$/g, '');
            if (userLabel === '') userLabel = base;

            // カテゴリ決定 / Decide category
            var selectedCat = (rbFont.value === true) ? 'フォント' : 'スタイル／ブラシ／シンボル';

            // 5) 既存候補の上書き or 追加
            var updated = {
                label: userLabel,
                path: dest.name,
                category: selectedCat
            };
            var found = -1;
            for (var i = 0; i < CANDIDATES.length; i++) {
                if (CANDIDATES[i] && CANDIDATES[i].path === dest.name) {
                    found = i;
                    break;
                }
            }
            if (found >= 0) {
                CANDIDATES[found] = updated;
            } else {
                CANDIDATES.push(updated);
            }

            // 6) TSVへ保存 → UI再構築 → 追加/上書き対象を選択
            if (saveCandidates(CANDIDATES)) {
                relayout();
                selectByPath(dest.name);
                alert((found >= 0 ? T('msgOverwritten') : T('msgAdded')) + "\n" +
                    userLabel + "  (" + toDisplayName(dest.name) + ")" + T('msgSynced'));
            }
        } catch (e) {
            alert(T('errAddGeneral') + "\n" + e);
        }
    };

    var btnOk = btnRightGroup.add('button', undefined, T('btnLoad'), {
        name: 'ok'
    });
    btnOk.onClick = function() {
        dlg.close(1);
    };

    // 初期表示 → ダイアログ表示 → 選択を返却
    relayout();
    var result = dlg.show();
    if (result !== 1) return null;
    var chosenItem = list && list.selection ? list.selection : null;
    if (!chosenItem) return null;
    return {
        path: chosenItem._path
    };
}

/* =========================
   メイン処理 / Main Process
   ========================= */
function main() {
    var choice = null;
    try {
        choice = chooseFileFromList(CANDIDATES);
    } catch (e) {
        try {
            alert("Dialog failed to open:\n" + e);
        } catch (_) {}
        return;
    }
    if (!choice) return;

    // ここでドキュメントの存在を確認 / Check for an open document now
    if (app.documents.length === 0) {
        alert(T('msgOpenDocFirst'));
        return;
    }

    var originalDoc = app.activeDocument;

    // 2) 指定ファイルを開く / Open the chosen file
    var fileToOpen = new File(choice.path);
    if (!fileToOpen.exists) {
        alert(T('msgFileNotFoundTitle') + "\n" + T('msgFileNotFoundBody') + toDisplayName(choice.path));
        return;
    }
    var tempDoc = app.open(fileToOpen);

    // 3) 作業アートボード内の全てをコピー / Copy in current artboard
    app.executeMenuCommand('selectallinartboard');
    app.executeMenuCommand('copy');

    // 4) 閉じる（保存しない）/ Close without saving
    tempDoc.close(SaveOptions.DONOTSAVECHANGES);

    // 5) 取り込み先レイヤーを用意してからペースト / Prepare destination layer then paste
    app.activeDocument = originalDoc;
    var importLayer = getOrCreateImportLayer(originalDoc);
    try {
        originalDoc.activeLayer = importLayer;
    } catch (eSet) {}
    app.executeMenuCommand('paste');
}

main();