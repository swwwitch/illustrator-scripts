#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

ImportStyles.jsx
指定ファイルの作業アートボードからコピーペースト（ダイアログ付き）

### GitHub：
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportStyles.jsx

### 概要：
- ダイアログで候補リストからAIファイルを1つ選択
- 指定AIを開き、作業アートボード内のオブジェクトをコピーして閉じる
- 元ドキュメントにペーストし、削除チェックONならオブジェクト削除（スタイルのみ残す）、OFFなら残す
- 「スタイル削除」ボタンで既存グラフィックスタイルを一括削除（最低1件は残る）
- キーボード操作で行選択（↑↓）、決定（Enter）、削除チェック切替（Space）

### 主な機能：
- 候補ファイルのダイアログ選択
- グラフィックスタイルのみの取り込み（削除オプション付き）
- グラフィックスタイル一括削除
- キーボード操作対応（↑↓、Enter、Space）

### 処理の流れ：
1. ダイアログで候補からAIファイルを選択
2. 指定ファイルを開き、作業アートボード内をコピー
3. ファイルを閉じる（保存なし）
4. 元ドキュメントにペースト
5. 削除チェックONならオブジェクトを削除

### note：
- Illustratorの仕様によりグラフィックスタイルは最低1件残ります
- 候補リストは最大8行表示（超える場合はスクロール）

### 更新履歴：
- v1.3 (20250815) : カテゴリ列（グラフィックスタイル／フォント）対応、追加時にカテゴリ選択（ドロップダウン）
- v1.2 (20250815) : CANDIDATESを外部ファイル（TSV）から読み込み可能に
- v1.1 (20250815) : CANDIDATESに読み込み後に削除のON/OFFを記録
- v1.0 (20250814) : 初期バージョン

---

### Script Name:
Copy from Working Artboard and Import Styles (with Dialog)

### GitHub:
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportStyles.jsx

### Overview:
- Choose an AI file from the dialog list
- Open the chosen AI, copy objects from the working artboard, and close it
- Paste into the original document; if "Delete" is checked, remove objects (keep styles only), otherwise keep them
- "Delete Styles" button: bulk delete existing graphic styles (at least one must remain)
- Keyboard navigation: ↑↓ to move, Enter to confirm, Space to toggle delete

### Key Features:
- Dialog selection from candidate files
- Import graphic styles only (with optional object deletion)
- Bulk delete of graphic styles
- Keyboard shortcuts support

### Process Flow:
1. Select AI file from dialog list
2. Open file, copy objects from working artboard
3. Close file without saving
4. Paste into original document
5. Delete objects if delete checkbox is ON

### note:
- At least one graphic style must remain due to Illustrator specs
- Candidate list shows up to 8 rows, scrolls if exceeded

### Changelog:
- v1.3 (20250815) : Added Category column (Graphic Style/Font) and category dropdown on Add
- v1.2 (20250815) : Load CANDIDATES from external TSV file
- v1.1 (20250815) : Added delete option in CANDIDATES to remember after loading
- v1.0 (20250814) : Initial release

*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.3";

/* =========================
   設定 / Settings
   ========================= */

// ベースディレクトリ / Base directory
// ※ここはユーザーの環境に合わせて変更してください / Change this to your environment
var DIRECTORY = "~/sw Dropbox/takano masahiro/sync-setting/ai/styles/";

// 既定候補（ファイル未提供時のフォールバック）/ Default candidates (fallback)
var DEFAULT_CANDIDATES = [
    { label: "オープンパス", path: "オープンパス.ai", check: 1, category: "グラフィックスタイル" },
    { label: "矢印",       path: "矢印.ai",       check: 0, category: "グラフィックスタイル" },
    { label: "丸数字",     path: "丸数字.ai",     check: 1, category: "グラフィックスタイル" },
    { label: "定番フォント（欧文）",     path: "fonts.ai",     check: 0, category: "フォント" }
];

// 候補リストの外部ファイル（TSV）/ External candidates TSV
// フォーマット: label \t path \t check(0|1)  ※checkは省略可
var CANDIDATES_FILE = File(Folder(DIRECTORY).fsName + "/ImportStyles_candidates.tsv");

// 現在のロケールを取得 / Get current locale
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var LANG = getCurrentLang();

/*
  〈ローカライズ〉LABELS は UI 出現順に定義
  Define LABELS in the order they appear in the UI
*/
var LABELS = {
    // ===== Dialog & Header =====
    dialogTitle: { ja: "スタイル読み込み：ファイル選択 " + SCRIPT_VERSION, en: "Import Styles: Choose File " + SCRIPT_VERSION },
    columnDelete: { ja: "削除", en: "Delete" },
    columnCategory: { ja: "カテゴリ", en: "Category" },
    columnContent: { ja: "コンテンツ", en: "Content" },
    columnFilename: { ja: "ファイル名", en: "Filename" },

    // ===== Buttons (footer) =====
    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnAdd: { ja: "追加", en: "Add" },
    btnLoad: { ja: "読み込み", en: "Load" },

    // ===== Prompts (Add flow) =====
    dlgPickAi: { ja: "追加するAIファイルを選択", en: "Choose an AI file to add" },
    promptEnterLabel: { ja: "ダイアログボックスに表示される項目名を指定", en: "Enter the display name for the dialog" },
    dlgPickCategory: { ja: "カテゴリを選択", en: "Choose a category" },
    optCategoryStyle: { ja: "グラフィックスタイル", en: "Graphic Style" },
    optCategoryFont: { ja: "フォント", en: "Font" },

    // ===== Runtime messages =====
    msgOpenDocFirst: { ja: "元のドキュメントを開いてから実行してください。", en: "Please open the destination document first." },
    msgFileNotFoundTitle: { ja: "ファイルが見つかりません", en: "File Not Found" },
    msgFileNotFoundBody: { ja: "指定したファイルが見つかりません：\n", en: "The specified file was not found:\n" },

    // Success (Add)
    msgAdded: { ja: "候補を追加しました：", en: "Candidate added:" },
    msgOverwritten: { ja: "候補を上書きしました：", en: "Candidate overwritten:" },
    msgSynced: { ja: "\n\nImportStyles_candidates.tsv と現在のダイアログに反映しました。", en: "\n\nReflected in ImportStyles_candidates.tsv and the current dialog." },

    // Errors (file/dir ops)
    errCreateDir: { ja: "保存先フォルダを作成できませんでした。\n", en: "Could not create directory:\n" },
    errCopyFile: { ja: "ファイルをコピーできませんでした。\n", en: "Could not copy file to:\n" },
    errCopying: { ja: "ファイルのコピー中にエラーが発生しました。\n", en: "Error copying file:\n" },
    errCandCreate: { ja: "候補TSVを作成できませんでした。\n", en: "Could not create candidates TSV:\n" },
    errCandOpen: { ja: "候補TSVを開けませんでした。\n", en: "Could not open candidates TSV:\n" },
    errCandSaveFolder: { ja: "候補TSVの保存先を作成できませんでした。\n", en: "Could not create folder:\n" },
    errCandOpenWrite: { ja: "候補TSVを書き込み用に開けませんでした。\n", en: "Could not open candidates TSV for writing:\n" },
    errAddGeneral: { ja: "候補の追加中にエラーが発生しました。\n", en: "An error occurred while adding a candidate.\n" }
};

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
                var path  = parts[1];
                var check = (parts.length >= 3) ? (parseInt(parts[2], 10) === 1 ? 1 : 0) : 1; // 省略時は1
                arr.push({ label: label, path: path, check: check, category: (parts.length >= 4 && parts[3] !== "" ? parts[3] : "グラフィックスタイル") });
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
            alert("候補TSVの保存先を作成できませんでした.\nCould not create folder:\n" + toDisplayPath(parent.fsName));
            return false;
        }
        cf.encoding = "UTF-8";
        if (!cf.open('w')) {
            alert("候補TSVを書き込み用に開けませんでした.\nCould not open candidates TSV for writing:\n" + toDisplayPath(cf.fsName));
            return false;
        }
        cf.writeln("# label\tpath\tcheck\tcategory");
        for (var i = 0; i < cands.length; i++) {
            var ci = cands[i];
            if (!ci || !ci.path) continue;
            var lab = (ci.label != null) ? ci.label : "";
            var chk = (ci.check ? 1 : 0);
            var cat = (ci.category != null && ci.category !== "") ? ci.category : "グラフィックスタイル";
            cf.writeln(lab + "\t" + ci.path + "\t" + chk + "\t" + cat);
        }
        cf.close();
        return true;
    } catch (e) {
        try { alert("候補TSVの保存中にエラーが発生しました。\nError saving candidates TSV:\n" + e); } catch(_){}
        return false;
    }
}

// フルパス化（~ を展開）/ Expand '~' and build absolute path
var STATE_FILE = File(Folder(DIRECTORY).fsName + "/ImportStyles_state.tsv");

// 状態を読み込み（path -> check）して CANDIDATES に反映
function loadState(cands) {
    try {
        var f = STATE_FILE;
        if (!f.exists) return;
        f.encoding = "UTF-8";
        if (!f.open('r')) return;
        var map = {};
        while (!f.eof) {
            var line = f.readln();
            if (!line || line.charAt(0) === '#') continue;
            var parts = line.split('\t');
            if (parts.length >= 2) {
                var p = parts[0];
                var v = parseInt(parts[1], 10);
                map[p] = (v === 1) ? 1 : 0;
            }
        }
        f.close();
        for (var i = 0; i < cands.length; i++) {
            var ci = cands[i];
            if (map.hasOwnProperty(ci.path)) {
                ci.check = map[ci.path];
            }
        }
    } catch (e) {}
}

// 現在の CANDIDATES の check 値を保存（TSV: path \t check）
function saveState(cands) {
    try {
        var f = STATE_FILE;
        f.encoding = "UTF-8";
        // 親フォルダが無ければ作成 / Ensure parent folder exists
        try {
            var parent = f.parent;
            if (parent && !parent.exists) parent.create();
        } catch (e) {}
        if (!f.open('w')) return;
        f.writeln("# ImportStyles state");
        f.writeln("# path\\tcheck (1=delete after paste, 0=keep)");
        for (var i = 0; i < cands.length; i++) {
            var ci = cands[i];
            var chk = (ci && ci.check) ? 1 : 0;
            f.writeln(ci.path + "\t" + chk);
        }
        f.close();
        try {
            if (!f.exists) {
                // 書き込み失敗時のサイレントフォールバック（必要なら alert に変更可）
            }
        } catch (e) {}
    } catch (e) {}
}

var CANDIDATES = loadCandidates(DEFAULT_CANDIDATES);
loadState(CANDIDATES);
// フォントカテゴリは常に削除OFF / Enforce OFF for Font category
for (var i = 0; i < CANDIDATES.length; i++) {
    if (CANDIDATES[i] && CANDIDATES[i].category === 'フォント') {
        CANDIDATES[i].check = 0;
    }
}

// 指定キーのローカライズ文字列を取得 / Get localized label for key
function T(key) {
    return (LABELS[key] && LABELS[key][LANG]) ? LABELS[key][LANG] : ("[" + key + "]");
}

/* =========================
   ユーティリティ / Utilities
   ========================= */

/*
  指定ドキュメントの全グラフィックスタイルを削除
  Delete all graphic styles in the given document.
  ※Illustrator仕様により最低1件は残る
  At least one must remain by Illustrator spec.
*/
function deleteAllGraphicStyles(doc) {
    var gs = doc && doc.graphicStyles;
    if (!gs || gs.length === 0) return;
    for (var i = gs.length - 1; i >= 0; i--) {
        if (gs.length <= 1) break;
        try {
            gs[i].remove();
        } catch (e) {}
    }
    app.redraw();
}

/*
  パスからファイル名のみ取得
  Get filename from path
*/
function toFilename(p) {
    try {
        return decodeURI(new File(p).name);
    } catch (e) {
        return p;
    }
}

/*
  表示用にパスをデコード
  Decode path for display
*/
function toDisplayPath(p) {
    try {
        var s = String(p);
        // %E3%81%AAといった %xx を含む場合は decode
        if (/%[0-9A-Fa-f]{2}/.test(s)) {
            try { return decodeURI(s); } catch (e1) {}
        }
        return s;
    } catch (e) {
        return String(p);
    }
}

/*
  表示名取得（ファイル名のみ）
  Get display name (filename only)
*/
function toDisplayName(n) {
    // ファイル名（パス不要）は既存の toFilename を使用
    return toFilename(n);
}

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
    var rbFilterStyle = filterGroup.add('radiobutton', undefined, T('optCategoryStyle')); // グラフィックスタイル
    var rbFilterFont  = filterGroup.add('radiobutton', undefined, T('optCategoryFont'));  // フォント
    // 既定ではどちらも未選択（全件表示）/ No filter by default

    // ヘッダ行 / Header
    var header = dlg.add('group');
    header.orientation = 'row';
    header.alignChildren = ['left', 'center'];

    var colW = {
        del: 30,
        cat: 140,
        content: 150,
        fname: 160
    };

    var hDel = header.add('statictext', undefined, T('columnDelete'));
    hDel.preferredSize = [colW.del, -1];
    var hCat = header.add('statictext', undefined, T('columnCategory'));
    hCat.preferredSize = [colW.cat, -1];
    var hContent = header.add('statictext', undefined, T('columnContent'));
    hContent.preferredSize = [colW.content, -1];
    var hFile = header.add('statictext', undefined, T('columnFilename'));
    hFile.preferredSize = [colW.fname, -1];

    // 行リスト（スクロール）/ Rows area (scrollable)
    var listPanel = dlg.add('panel', undefined, '');
    listPanel.alignment = ['fill', 'fill'];
    var panelWidth = colW.del + colW.cat + colW.content + colW.fname + 20;
    // 高さは後で候補数に応じて自動計算（最大行数で上限）/ Height computed after rows are built
    listPanel.preferredSize = [panelWidth, 10];
    listPanel.margins = 8;

    var rowsContainer = listPanel.add('group');
    rowsContainer.orientation = 'column';
    rowsContainer.alignChildren = ['left', 'fill'];

    // 行データ保持 / Keep references
    var rows = []; // { group, cb, rb, st, path }
    var selectedIndex = -1; // 現在選択

    // ダイアログ高さの自動調整用 / Dynamic panel height helper
    var MAX_VISIBLE_ROWS = 8; // これ以上はスクロール
    var ROW_H = 22;           // 行の目安高さ（checkbox+radio+label）
    var V_PAD = 16;           // panel の上下マージン分

    function getVisibleRowCount() {
        var cnt = 0;
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].group.visible !== false) cnt++;
        }
        return cnt;
    }
    function isRowVisible(idx) {
        return rows[idx] && rows[idx].group.visible !== false;
    }
    function selectFirstVisible() {
        for (var i = 0; i < rows.length; i++) {
            if (isRowVisible(i)) {
                selectIndex(i);
                return;
            }
        }
    }

    function updatePanelHeight() {
        var visibleRows = Math.min(getVisibleRowCount(), MAX_VISIBLE_ROWS);
        var dynamicHeight = (visibleRows * ROW_H) + V_PAD;
        listPanel.preferredSize = [panelWidth, dynamicHeight];
    }

function selectIndex(i) {
    if (i < 0 || i >= rows.length) return;
    for (var r = 0; r < rows.length; r++) {
        rows[r].rb.value = (r === i);
    }
    selectedIndex = i;
    updateRowHighlights(); // ← 選択変更時にハイライト更新
}

    // 行のハイライト更新
function updateRowHighlights() {
    for (var r = 0; r < rows.length; r++) {
        var rowPanel = rows[r].group; // 行コンテナ（panel）
        try {
            if (r === selectedIndex && isRowVisible(r)) {
                // 選択行のみ淡いブルー
                rowPanel.graphics.backgroundColor =
                    rowPanel.graphics.newBrush(rowPanel.graphics.BrushType.SOLID_COLOR, [0.85, 0.90, 1, 1]);
            } else {
                // 既定に戻す（undefinedでOK）
                rowPanel.graphics.backgroundColor = undefined;
            }
            rowPanel.graphics.invalidate(); // 再描画
        } catch (e) {}
    }
}

    function activateRow(i) {
        selectIndex(i);
        dlg.close(1);
    }

    for (var i = 0; i < candidates.length; i++) {
        var c = candidates[i];

// 変更前
// var row = rowsContainer.add('group');
// row.orientation = 'row';
// row.alignChildren = ['left', 'center'];

// 変更後
var row = rowsContainer.add('panel', undefined, ''); // panelなら背景色が効く
row.margins = [0, 0, 0, 0];
row.orientation = 'row';
row.alignChildren = ['left', 'center'];
row.minimumSize.height = ROW_H; // 行高さの目安

        // 1列目：本物のチェックボックス / Real checkbox
        var cb = row.add('checkbox', undefined, '');
        cb.preferredSize = [colW.del, -1];
        cb.value = (c.category === 'フォント') ? false : !!c.check; // Font: always OFF
        if (c.category === 'フォント') {
            cb.enabled = false; // 固定OFF（操作不可）
        } else {
            (function(cand, checkbox){
                checkbox.onClick = function () {
                    cand.check = checkbox.value ? 1 : 0;
                };
            })(c, cb);
        }

        // 2列目：カテゴリ表示 / Category
        var sc = row.add('statictext', undefined, String(c.category || 'グラフィックスタイル'));
        sc.preferredSize = [colW.cat, -1];

        // 3列目：選択用ラジオボタン（コンテンツ名をラベルに）/ Radio for selection
        var rb = row.add('radiobutton', undefined, c.label);
        rb.preferredSize = [colW.content, -1];

        // 4列目：ファイル名表示 / Filename
        var st = row.add('statictext', undefined, toFilename(c.path));
        st.preferredSize = [colW.fname, -1];

        // 行クリックでラジオ選択 / Click row selects
        (function(idx) {
            if (row.addEventListener) {
                row.addEventListener('click', function() {
                    selectIndex(idx);
                });
                // ダブルクリックで決定 / Double-click to confirm
                row.addEventListener('doubleclick', function() {
                    activateRow(idx);
                });
            }
            rb.onClick = function() {
                selectIndex(idx);
            };
            // ラジオのダブルクリックでも決定
            rb.onDoubleClick = function() {
                activateRow(idx);
            };
            // ファイル名のダブルクリックでも決定
            if (st.addEventListener) {
                st.addEventListener('doubleclick', function() {
                    activateRow(idx);
                });
            }
        })(i);

        rows.push({
            group: row,
            cb: cb,
            cat: sc,
            rb: rb,
            st: st,
            _labelBase: c.label,
            path: DIRECTORY + c.path,
            candidate: c
        });
    }

    // カテゴリフィルター適用 / Apply category filter
    function applyCategoryFilter(categoryName) {
        for (var i = 0; i < rows.length; i++) {
            var show = (rows[i].candidate && rows[i].candidate.category === categoryName);
            rows[i].group.visible = show;
            // ラジオ自体の状態は保持（表示/非表示のみ）/ keep rb state
        }
        updatePanelHeight();
        dlg.layout.layout(true);
        selectFirstVisible();
        updateRowHighlights();
    }
    rbFilterStyle.onClick = function () { applyCategoryFilter('グラフィックスタイル'); };
    rbFilterFont.onClick  = function () { applyCategoryFilter('フォント'); };

    // === ダイナミック高さ計算 / Dynamic height with max cap ===
    updatePanelHeight();

    // 既定選択（先頭） / Default selection
    if (rows.length > 0) selectFirstVisible();
    updateRowHighlights();

    // キーボード操作：↑↓で選択行移動、Enterで決定 / Keyboard navigation
    if (dlg.addEventListener) {
        dlg.addEventListener('keydown', function(evt) {
            try {
                var k = evt.keyName;
                if (k === 'Up') {
                    var u = selectedIndex;
                    do { u--; } while (u >= 0 && !isRowVisible(u));
                    if (u >= 0) selectIndex(u);
                    evt.preventDefault && evt.preventDefault();
                } else if (k === 'Down') {
                    var d = selectedIndex;
                    do { d++; } while (d < rows.length && !isRowVisible(d));
                    if (d < rows.length) selectIndex(d);
                    evt.preventDefault && evt.preventDefault();
                } else if (k === 'Enter') {
                    activateRow(selectedIndex);
                } else if (k === 'Space') {
                    // Spaceで削除チェックをトグル（選択行に対して）
                    if (selectedIndex >= 0) {
                        var rowRef = rows[selectedIndex];
                        if (rowRef.candidate && rowRef.candidate.category === 'フォント') {
                            // フォントは常にOFF固定
                            rowRef.cb.value = false;
                            rowRef.candidate.check = 0;
                        } else {
                            rowRef.cb.value = !rowRef.cb.value;
                            rowRef.candidate.check = rowRef.cb.value ? 1 : 0;
                        }
                    }
                    evt.preventDefault && evt.preventDefault();
                }
            } catch (e) {}
        });
    }

    // ボタン行：左(キャンセル) - スペーサー - 右(スタイル削除 / 読み込み)
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
    btnAdd.onClick = function () {
        try {
            // 1) ファイル選択
            var picked = File.openDialog(T('dlgPickAi'), '*.ai');
            if (!picked) return;

            // 2) 保存先フォルダ（DIRECTORY）を確保
            var dirFolder = Folder(DIRECTORY);
            var dirPath = dirFolder.fsName;
            if (!dirFolder.exists && !dirFolder.create()) {
                alert("保存先フォルダを作成できませんでした.\nCould not create directory:\n" + toDisplayPath(dirPath));
                return;
            }

            // 3) DIRECTORY 配下へコピー（同名は上書き）
            var dest = File(dirPath + '/' + picked.name);
            try {
                if (picked.fsName !== dest.fsName) {
                    if (dest.exists) dest.remove();
                    if (!picked.copy(dest.fsName)) {
                        alert("ファイルをコピーできませんでした。\nCould not copy file to:\n" + toDisplayPath(dest.fsName));
                        return;
                    }
                }
            } catch (eCopy) {
                alert("ファイルのコピー中にエラーが発生しました。\nError copying file:\n" + eCopy);
                return;
            }

            // 4) 表示名の入力（既定は拡張子抜きのファイル名）/ Ask display name
            var base = picked.name.replace(/\.[^\.]+$/, ''); // 拡張子抜き
            var userLabel = prompt(T('promptEnterLabel'), base);
            if (userLabel === null) {
                // ユーザーがキャンセルした場合は中断 / Cancel addition if user cancels
                return;
            }
            // タブや改行はTSVに不向きなので除去・整形 / Sanitize for TSV
            userLabel = String(userLabel).replace(/\r?\n/g, ' ').replace(/\t/g, ' ').replace(/^\s+|\s+$/g, '');
            if (userLabel === "") userLabel = base;

            // カテゴリ選択ダイアログ / Category dropdown
            var catDlg = new Window('dialog', T('dlgPickCategory'));
            catDlg.orientation = 'column';
            catDlg.alignChildren = ['fill', 'top'];
            var dd = catDlg.add('dropdownlist', undefined, [T('optCategoryStyle'), T('optCategoryFont')]);
            dd.selection = 0; // default to グラフィックスタイル
            var g = catDlg.add('group');
            g.alignment = ['fill', 'bottom'];
            var btnC = g.add('button', undefined, T('btnCancel'), {name:'cancel'});
            var btnO = g.add('button', undefined, T('btnLoad'), {name:'ok'});
            if (catDlg.show() !== 1) { return; }
            var selectedCat = (dd.selection && dd.selection.text === T('optCategoryFont')) ? 'フォント' : 'グラフィックスタイル';

            // 5) 既存候補の上書き or 追加 / Overwrite if duplicate, else append in memory
            var chkVal = (selectedCat === 'フォント') ? 0 : 1;
            var updated = { label: userLabel, path: dest.name, check: chkVal, category: selectedCat };
            var found = -1;
            for (var i = 0; i < CANDIDATES.length; i++) {
                if (CANDIDATES[i] && CANDIDATES[i].path === dest.name) { found = i; break; }
            }
            if (found >= 0) {
                CANDIDATES[found] = updated; // 上書き
            } else {
                CANDIDATES.push(updated);    // 追加
            }

            // 6) TSVをフル書き出し（上書き保存）/ Rewrite TSV
            if (saveCandidates(CANDIDATES)) {
                // 既存行の上書き or 新規行の追加でUIに反映
                if (found >= 0) {
                    // 既存行を更新
                    var rowRef = rows[found];
                    rowRef.rb.text = userLabel;
                    rowRef.st.text = toFilename(dest.name);
                    rowRef.cat.text = selectedCat;
                    rowRef.candidate = CANDIDATES[found];
                    if (selectedCat === 'フォント') {
                        rowRef.cb.value = false;
                        rowRef.cb.enabled = false;
                        rowRef.candidate.check = 0;
                    } else {
                        rowRef.cb.enabled = true;
                        rowRef.cb.value = true; // 追加直後はON既定
                        rowRef.candidate.check = 1;
                    }
                    selectIndex(found);
                } else {
                    // 新規行を作成して追加
                    var chkVal2 = (selectedCat === 'フォント') ? 0 : 1;
                    var c = { label: userLabel, path: dest.name, check: chkVal2, category: selectedCat };
                    var row = rowsContainer.add('panel', undefined, '');
                    row.margins = [0, 0, 0, 0];
                    row.orientation = 'row';
                    row.alignChildren = ['left', 'center'];
                    row.minimumSize.height = ROW_H;

                    var cb = row.add('checkbox', undefined, '');
                    cb.preferredSize = [colW.del, -1];
                    cb.value = (c.category === 'フォント') ? false : !!c.check;
                    if (c.category === 'フォント') {
                        cb.enabled = false;
                    } else {
                        (function(cand, checkbox){
                            checkbox.onClick = function () {
                                cand.check = checkbox.value ? 1 : 0;
                            };
                        })(c, cb);
                    }

                    var sc2 = row.add('statictext', undefined, String(c.category || 'グラフィックスタイル'));
                    sc2.preferredSize = [colW.cat, -1];

                    var rb = row.add('radiobutton', undefined, c.label);
                    rb.preferredSize = [colW.content, -1];

                    var st = row.add('statictext', undefined, toFilename(c.path));
                    st.preferredSize = [colW.fname, -1];

                    (function(idx) {
                        if (row.addEventListener) {
                            row.addEventListener('click', function () { selectIndex(idx); });
                            row.addEventListener('doubleclick', function () { activateRow(idx); });
                        }
                        rb.onClick = function () { selectIndex(idx); };
                        rb.onDoubleClick = function () { activateRow(idx); };
                        if (st.addEventListener) {
                            st.addEventListener('doubleclick', function () { activateRow(idx); });
                        }
                    })(rows.length);

                    rows.push({
                        group: row,
                        cb: cb,
                        cat: sc2,
                        rb: rb,
                        st: st,
                        _labelBase: c.label,
                        path: DIRECTORY + c.path,
                        candidate: c
                    });
                    updatePanelHeight();      // 高さ再計算
                    selectIndex(rows.length - 1); // 追加した行を選択
                    dlg.layout.layout(true);  // レイアウトを再計算
                }

                alert((found >= 0 ? "候補を上書きしました：" : "候補を追加しました：") + "\n" + userLabel + "  (" + toDisplayName(dest.name) + ")\n\nImportStyles_candidates.tsv と現在のダイアログに反映しました。");
            }
        } catch (e) {
            alert("候補の追加中にエラーが発生しました。\nAn error occurred while adding a candidate.\n\n" + e);
        }
    };
    // OK（読み込み）
    var btnOk = btnRightGroup.add('button', undefined, T('btnLoad'), {
        name: 'ok'
    });
    btnOk.onClick = function() {
        dlg.close(1);
    };

    var result = dlg.show();
    if (result !== 1) return null;

    // 現在のチェック値をCANDIDATESへ反映し保存
    for (var k = 0; k < rows.length; k++) {
        if (rows[k].candidate && rows[k].candidate.category === 'フォント') {
            rows[k].cb.value = false;
            rows[k].candidate.check = 0;
        } else {
            rows[k].candidate.check = rows[k].cb.value ? 1 : 0;
        }
    }
    saveState(CANDIDATES);

    // 選択された行の取得
    var chosen = null;
    for (var j = 0; j < rows.length; j++) {
        if (rows[j].rb.value) {
            chosen = rows[j];
            break;
        }
    }
    if (!chosen) return null;

    return {
        path: chosen.path,
        deleteAfter: !!chosen.cb.value
    };
}

/*
  ペースト後の処理を関数化
  Apply post-paste action
  - deleteAfter = true  -> 即削除（スタイルのみ残す）
  - deleteAfter = false -> そのまま残す
*/
function applyPostPasteAction(deleteAfter) {
    if (deleteAfter) {
        // ON: 読み込み後に削除（スタイルのみ残す） / Clear objects after paste (keep styles)
        app.executeMenuCommand('clear');
    } else {
        // OFF: 読み込み後にオブジェクトを残す / Keep pasted objects
    }
}

/* =========================
   メイン処理 / Main Process
   ========================= */
function main() {
    if (app.documents.length === 0) {
        alert(T('msgOpenDocFirst'));
        return;
    }

    var originalDoc = app.activeDocument;

    // 1) 候補から選択 / Choose a file
    var choice = chooseFileFromList(CANDIDATES);
    if (!choice) return;

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

    // 5) ペースト→即削除（スタイルのみ残す）/ Paste then clear
    app.activeDocument = originalDoc;
    app.executeMenuCommand('paste');
    applyPostPasteAction(choice.deleteAfter);
    /*
      ※不要なロジックや整理できる箇所があれば下記コメントにて提案
      // 例: 「deleteAfter」判定後の分岐は関数化できる
    */
}

main();