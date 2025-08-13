#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

指定ファイルの作業アートボードからコピーペースト（ダイアログ付き）

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportGraphicStyle.jsx

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

- v1.0 (20250814) : 初期バージョン
- v1.1 (20250815) : CANDIDATESに読み込み後に削除のON/OFFを記録

---

### Script Name:

Copy from Working Artboard and Import Styles (with Dialog)

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportGraphicStyle.jsx

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

- v1.0 (20250814) : Initial release
- v1.1 (20250815) : Added delete option in CANDIDATES to remember after loading
*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.1";

/* =========================
   設定 / Settings
   ========================= */
// ベースディレクトリ / Base directory
var DIRECTORY = "~/sw Dropbox/takano masahiro/sync-setting/ai/graphic_style/";

// ダイアログに表示する候補（ラベルとファイル名）/ Candidates (label & filename only)
var CANDIDATES = [{
        label: "オープンパス",
        path: "オープンパス.ai",
        check: 1
    },
    {
        label: "矢印",
        path: "矢印.ai",
        check: 0
    },
    {
        label: "丸数字",
        path: "丸数字.ai",
        check: 1
    }
];

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var LANG = getCurrentLang();

/* 〈ローカライズ〉LABELS は UI 出現順に定義 / Define in the order they appear in the UI */
var LABELS = {
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
    columnDelete: {
        ja: "削除",
        en: "Delete"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    btnDeleteStyles: {
        ja: "スタイル削除",
        en: "Delete Styles"
    },
    btnLoad: {
        ja: "読み込み",
        en: "Load"
    },
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
    }
};

function T(key) {
    return (LABELS[key] && LABELS[key][LANG]) ? LABELS[key][LANG] : ("[" + key + "]");
}

/* =========================
   ユーティリティ / Utilities
   ========================= */
/**
 * Delete all graphic styles in the given document.
 * At least one must remain by Illustrator spec.
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

function toFilename(p) {
    try {
        return decodeURI(new File(p).name);
    } catch (e) {
        return p;
    }
}

/**
 * Show a dialog to choose one candidate, return its full path or null.
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

    // ヘッダ行 / Header
    var header = dlg.add('group');
    header.orientation = 'row';
    header.alignChildren = ['left', 'center'];

    var colW = {
        del: 30,
        content: 150,
        fname: 160
    };

    var hDel = header.add('statictext', undefined, T('columnDelete'));
    hDel.preferredSize = [colW.del, -1];
    var hContent = header.add('statictext', undefined, T('columnContent'));
    hContent.preferredSize = [colW.content, -1];
    var hFile = header.add('statictext', undefined, T('columnFilename'));
    hFile.preferredSize = [colW.fname, -1];

    // 行リスト（スクロール）/ Rows area (scrollable)
    var listPanel = dlg.add('panel', undefined, '');
    listPanel.alignment = ['fill', 'fill'];
    var panelWidth = colW.del + colW.content + colW.fname + 20;
    // 高さは後で候補数に応じて自動計算（最大行数で上限）/ Height computed after rows are built
    listPanel.preferredSize = [panelWidth, 10];
    listPanel.margins = 8;

    var rowsContainer = listPanel.add('group');
    rowsContainer.orientation = 'column';
    rowsContainer.alignChildren = ['left', 'fill'];

    // 行データ保持 / Keep references
    var rows = []; // { group, cb, rb, st, path }
    var selectedIndex = -1; // 現在選択

    function selectIndex(i) {
        if (i < 0 || i >= rows.length) return;
        for (var r = 0; r < rows.length; r++) {
            rows[r].rb.value = (r === i);
        }
        selectedIndex = i;
    }

    function activateRow(i) {
        selectIndex(i);
        dlg.close(1);
    }

    for (var i = 0; i < candidates.length; i++) {
        var c = candidates[i];

        var row = rowsContainer.add('group');
        row.orientation = 'row';
        row.alignChildren = ['left', 'center'];

        // 1列目：本物のチェックボックス / Real checkbox
        var cb = row.add('checkbox', undefined, '');
        cb.preferredSize = [colW.del, -1];
        cb.value = !!c.check; // Set initial value from candidate's check property
        cb.onClick = function () {
            c.check = cb.value ? 1 : 0;
        };

        // 2列目：選択用ラジオボタン（コンテンツ名をラベルに）/ Radio for selection
        var rb = row.add('radiobutton', undefined, c.label);
        rb.preferredSize = [colW.content, -1];

        // 3列目：ファイル名表示 / Filename
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
            rb: rb,
            st: st,
            _labelBase: c.label,
            path: DIRECTORY + c.path
        });
    }

    // === ダイナミック高さ計算 / Dynamic height with max cap ===
    var MAX_VISIBLE_ROWS = 8; // これ以上はスクロール
    var ROW_H = 22; // 行の目安高さ（checkbox+radio+label）
    var V_PAD = 16; // panel の上下マージン分
    var visibleRows = Math.min(rows.length, MAX_VISIBLE_ROWS);
    var dynamicHeight = (visibleRows * ROW_H) + V_PAD;
    listPanel.preferredSize = [panelWidth, dynamicHeight];

    // 既定選択（先頭） / Default selection
    if (rows.length > 0) selectIndex(0);

    // キーボード操作：↑↓で選択行移動、Enterで決定 / Keyboard navigation
    if (dlg.addEventListener) {
        dlg.addEventListener('keydown', function(evt) {
            try {
                var k = evt.keyName;
                if (k === 'Up') {
                    if (selectedIndex > 0) selectIndex(selectedIndex - 1);
                    evt.preventDefault && evt.preventDefault();
                } else if (k === 'Down') {
                    if (selectedIndex < rows.length - 1) selectIndex(selectedIndex + 1);
                    evt.preventDefault && evt.preventDefault();
                } else if (k === 'Enter') {
                    activateRow(selectedIndex);
                } else if (k === 'Space') {
                    // Spaceで削除チェックをトグル（選択行に対して）
                    if (selectedIndex >= 0) rows[selectedIndex].cb.value = !rows[selectedIndex].cb.value;
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
    var btnDelete = btnRightGroup.add('button', undefined, T('btnDeleteStyles'), '');
    btnDelete.onClick = function() {
        deleteAllGraphicStyles(app.activeDocument);
    };
    var btnOk = btnRightGroup.add('button', undefined, T('btnLoad'), {
        name: 'ok'
    });
    btnOk.onClick = function() {
        dlg.close(1);
    };

    var result = dlg.show();
    if (result !== 1) return null;

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
        alert(T('msgFileNotFoundTitle') + "\n" + T('msgFileNotFoundBody') + choice.path);
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
    if (choice.deleteAfter) {
        app.executeMenuCommand('clear'); // ON: 読み込み後に削除（スタイルのみ残す）
    } else {
        // OFF: 読み込み後にオブジェクトを残す
    }
}

main();