#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.3";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
  dialogTitle: { ja: "データ流し込み", en: "Data Import" },
  fileLabel: { ja: "ファイル:", en: "File:" },
  settingsTitle: { ja: "流し込み設定", en: "Import Settings" },
  abNameLabel: { ja: "アートボード名に使用する項目:", en: "Field for artboard name:" },
  colsLabel: { ja: "横に並べる数:", en: "Columns:" },
  wrapLabel: { ja: "列で折り返し", en: "wrap" },
  targetLabel: { ja: "流し込み対象:", en: "Target:" },
  targetArtboard: { ja: "アートボード", en: "Artboards" },
  targetSelection: { ja: "選択グループ", en: "Selected group(s)" },
  cancel: { ja: "キャンセル", en: "Cancel" },
  run: { ja: "実行", en: "Run" },
  processing: { ja: "処理中...", en: "Processing..." },
  done: { ja: "完了！\n#count# 件処理しました。", en: "Done!\nProcessed #count# rows." },
  noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
  needSave: { ja: "ドキュメントを保存してから実行してください。", en: "Please save the document before running." },
  noDataFiles: { ja: "同じ階層に.txtまたは.csvファイルが見つかりませんでした。", en: "No .txt or .csv files found in the same folder." },
  noTemplate: { ja: "雛形にオブジェクトがありません。", en: "No template objects found." },
  needGroupSelect: { ja: "選択グループを対象にするには、グループを選択してください。", en: "To use Selected group(s), please select a group." },
  onlyGroupSelect: { ja: "『選択グループ』を選んだ場合は、グループ（GroupItem）のみを選択してください。\n現在の選択にグループ以外が含まれています。", en: "When using Selected group(s), select GroupItem only.\nYour selection contains non-group items." },
  dupFailed: { ja: "ドキュメントの複製（別名保存）に失敗しました。\n\n#detail#", en: "Failed to duplicate (Save As) the document.\n\n#detail#" }
};

function L(key) {
  try {
    if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
  } catch (e) {}
  return key;
}

/*
  VariableDataImport_v12.jsx（更新）
  更新日：20260122
  仕様：
  1. 実行時にドキュメントを複製（別名保存）し、複製したファイルに対して流し込みを実行します。アートボード指定時：複製したドキュメント内でアートボードを増やして流し込み。選択グループ指定時：アートボードは増やさず、選択中のグループを同一アートボード内に複製して流し込み。
  2. CSV(.csv)とTSV(.txt)の両方に対応（カンマや引用符も解析）。
  3. 1件目のデータを最後に処理することで、タグ消失バグを回避。
  4. 進捗バーを表示。
  5. 流し込み対象（アートボード / 選択グループ）を選べる（起動時にグループ選択中なら自動で「選択グループ」を選択）。選択グループ時は「アートボード名に使用する項目」をディム表示。
*/

(function() {
    if (app.documents.length === 0) {
        alert(L("noDoc"));
        return;
    }

    var doc = app.activeDocument;
    var docPath;
    try {
        docPath = doc.path;
    } catch (e) {
        alert(L("needSave"));
        return;
    }

    var folder = new Folder(docPath);
    var files = folder.getFiles(/\.(txt|csv)$/i);
    if (files.length === 0) {
        alert(L("noDataFiles"));
        return;
    }

    // --- メインダイアログ ---
    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 10;
    win.margins = 15;

    var groupSelect = win.add("group");
    groupSelect.add("statictext", undefined, L("fileLabel"));
    var dropdown = groupSelect.add("dropdownlist", undefined, []);
    dropdown.size = [350, 25];
    for (var i = 0; i < files.length; i++) dropdown.add("item", decodeURI(files[i].name));

    var listContainer = win.add("group");
    listContainer.alignment = ["fill", "fill"];
    var listBox = null; 

    var settingsPanel = win.add("panel", undefined, L("settingsTitle"));
    settingsPanel.orientation = "column";
    settingsPanel.alignChildren = ["left", "top"];
    settingsPanel.margins = 15;

    var groupAbName = settingsPanel.add("group");
    groupAbName.add("statictext", undefined, L("abNameLabel"));
    var dropdownAbName = groupAbName.add("dropdownlist", undefined, []);
    dropdownAbName.size = [200, 25];

    function updateUiByTargetMode() {
        // 選択グループ時は、アートボード名の指定は不要なのでディム表示
        groupAbName.enabled = (targetMode !== "selection");
    }

    var groupLayout = settingsPanel.add("group");
    groupLayout.add("statictext", undefined, L("colsLabel"));
    var editCols = groupLayout.add("edittext", undefined, "10");
    editCols.size = [50, 25];
    groupLayout.add("statictext", undefined, L("wrapLabel"));

    // 流し込み対象（UIのみ。ロジックは後で実装）
    var groupTarget = settingsPanel.add("group");
    groupTarget.add("statictext", undefined, L("targetLabel"));
    var rbTargetArtboard = groupTarget.add("radiobutton", undefined, L("targetArtboard"));
    var rbTargetSelection = groupTarget.add("radiobutton", undefined, L("targetSelection"));

    // 現時点ではUI状態を保持するだけ（実行ロジックは後で反映）
    var targetMode = "artboard"; // "artboard" | "selection"

    // 起動時：グループ（GroupItem）のみが選択されていたら、自動で「選択グループ」を選ぶ
    (function initTargetModeBySelection() {
        try {
            var sel = doc.selection;
            if (!sel || sel.length === 0) {
                rbTargetArtboard.value = true;
                targetMode = "artboard";
                updateUiByTargetMode();
                return;
            }

            var allGroup = true;
            for (var i = 0; i < sel.length; i++) {
                if (!sel[i] || sel[i].typename !== "GroupItem") { allGroup = false; break; }
            }

            if (allGroup) {
                rbTargetSelection.value = true;
                targetMode = "selection";
                updateUiByTargetMode();
            } else {
                rbTargetArtboard.value = true;
                targetMode = "artboard";
                updateUiByTargetMode();
            }
        } catch (e) {
            rbTargetArtboard.value = true;
            targetMode = "artboard";
            updateUiByTargetMode();
        }
    })();

    rbTargetArtboard.onClick = function() { targetMode = "artboard"; updateUiByTargetMode(); };
    rbTargetSelection.onClick = function() { targetMode = "selection"; updateUiByTargetMode(); };

    updateUiByTargetMode();

    var groupBtn = win.add("group");
    groupBtn.alignment = "right";
    var btnCancel = groupBtn.add("button", undefined, L("cancel"), {name: "cancel"});
    var btnRun = groupBtn.add("button", undefined, L("run"), {name: "ok"});

    var tsvData = [];
    var headers = [];

    /* ドキュメント複製（別名保存） / Duplicate document by Save As */
    function buildDuplicateFilePath(originalFile) {
        var d = new Date();
        function pad2(n) { return (n < 10 ? "0" : "") + String(n); }
        var stamp = String(d.getFullYear()) + pad2(d.getMonth() + 1) + pad2(d.getDate()) + "_" + pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());

        var parent = originalFile.parent;
        var name = originalFile.name;
        var dot = name.lastIndexOf(".");
        var base = (dot >= 0) ? name.substring(0, dot) : name;
        var ext = (dot >= 0) ? name.substring(dot) : ".ai";

        var newName = base + "_import_" + stamp + ext;
        return new File(parent.fsName + "/" + newName);
    }

    function duplicateDocumentBySaveAs(doc) {
        var originalFile = doc.fullName; // File
        var dupFile = buildDuplicateFilePath(originalFile);
        doc.saveAs(dupFile); // current doc becomes the duplicated file
        return dupFile;
    }

    // --- CSV/TSV 解析関数 ---
    function parseLine(line, delimiter) {
        if (delimiter === "\t") return line.split("\t");
        var parts = [], pattern = new RegExp("(\\" + delimiter + "|\\r?\\n|\\r|^)(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|([^\"\\" + delimiter + "\\r\\n]*))", "gi");
        var matches = null;
        while (matches = pattern.exec(line)) {
            parts.push(matches[2] ? matches[2].replace(/\"\"/g, "\"") : matches[3]);
        }
        return parts;
    }

    function updateListBox(fileObj) {
        if (listBox !== null) listContainer.remove(listBox);
        tsvData = [];
        fileObj.open("r");
        fileObj.encoding = "UTF-8";
        var content = fileObj.read();
        fileObj.close();
        
        var lines = content.split(/[\r\n]+/);
        if (lines.length === 0) return;

        var delimiter = fileObj.name.toLowerCase().match(/\.csv$/) ? "," : "\t";
        headers = parseLine(lines[0], delimiter);

        // ヘッダーの正規化（BOM/前後空白を除去）
        for (var hh = 0; hh < headers.length; hh++) {
            headers[hh] = String(headers[hh]).replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        }

        dropdownAbName.removeAll();
        for (var h = 0; h < headers.length; h++) dropdownAbName.add("item", headers[h]);
        dropdownAbName.selection = 0;

        listBox = listContainer.add("listbox", [0, 0, 550, 180], [], {
            numberOfColumns: headers.length, showHeaders: true, columnTitles: headers
        });

        for (var i = 1; i < lines.length; i++) {
            if (lines[i] === "") continue;
            var cells = parseLine(lines[i], delimiter);

            // 値の正規化（前後空白を除去）
            for (var cc = 0; cc < cells.length; cc++) {
                cells[cc] = String(cells[cc]).replace(/^\s+|\s+$/g, "");
            }

            tsvData.push(cells);
            var item = listBox.add("item", cells[0] || "");
            for (var j = 1; j < cells.length; j++) {
                if (j < headers.length) item.subItems[j-1].text = cells[j] || "";
            }
        }
        win.layout.layout(true);
    }

    dropdown.onChange = function() { if (dropdown.selection) updateListBox(files[dropdown.selection.index]); };
    dropdown.selection = 0;

    // --- 実行処理 ---
    btnRun.onClick = function() {
        if (tsvData.length === 0) return;
        var nameIndex = dropdownAbName.selection.index;
        var maxCols = parseInt(editCols.text) || 10;

        // 実行前にドキュメントを複製（別名保存）し、複製したファイルに対して流し込み
        try {
            duplicateDocumentBySaveAs(doc);
        } catch (e) {
            alert(L("dupFailed").replace("#detail#", String(e)));
            return;
        }

        win.close();

        var progWin = new Window("palette", L("processing"), undefined);
        var progBar = progWin.add("progressbar", undefined, 0, tsvData.length);
        progBar.preferredSize.width = 300;
        progWin.show();

        // 雛形の情報を取得（レイアウトの基準はアートボード0）
        var sourceAb = doc.artboards[0];
        var abRect = sourceAb.artboardRect;
        var margin = 50;

        // 対象アイテム（雛形）
        var sourceItems = [];

        // 複製配置の幅・高さ（雛形のサイズ）
        var w = 0;
        var h = 0;

        // 選択中アイテムの外接バウンディングを取得
        function getUnionBounds(items) {
            // returns [left, top, right, bottom]
            var left = null, top = null, right = null, bottom = null;
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                if (!it) continue;
                var b = it.visibleBounds ? it.visibleBounds : it.geometricBounds;
                if (!b) continue;
                var l = b[0], t = b[1], r = b[2], bt = b[3];
                if (left === null || l < left) left = l;
                if (top === null || t > top) top = t;
                if (right === null || r > right) right = r;
                if (bottom === null || bt < bottom) bottom = bt;
            }
            return [left, top, right, bottom];
        }

        if (targetMode === "selection") {
            // 選択グループ：選択中のグループだけを雛形として複製
            var sel = doc.selection;
            if (!sel || sel.length === 0) {
                progWin.close();
                alert(L("needGroupSelect"));
                return;
            }

            // グループ以外が混ざっていたら中断（安全側）
            for (var si = 0; si < sel.length; si++) {
                if (!sel[si] || sel[si].typename !== "GroupItem") {
                    progWin.close();
                    alert(L("onlyGroupSelect"));
                    return;
                }
            }

            sourceItems = sel;
            var gb = getUnionBounds(sourceItems);
            w = gb[2] - gb[0];
            h = gb[1] - gb[3];
        } else {
            // アートボード：現在のアートボード上の選択オブジェクトを雛形として使用
            doc.selectObjectsOnActiveArtboard();
            sourceItems = doc.selection;
            if (!sourceItems || sourceItems.length === 0) {
                progWin.close();
                alert(L("noTemplate"));
                return;
            }

            // アートボードサイズを基準に配置
            w = abRect[2] - abRect[0];
            h = abRect[1] - abRect[3];
        }

        // 重要：2件目(i=1)から先に処理する（タグを維持するため）
        for (var i = 1; i < tsvData.length; i++) {
            progBar.value = i;
            progWin.update();

            var col = i % maxCols;
            var row = Math.floor(i / maxCols);
            var offsetX = (w + margin) * col;
            var offsetY = -(h + margin) * row;

            if (targetMode !== "selection") {
                var newAb = doc.artboards.add([
                    abRect[0] + offsetX, abRect[1] + offsetY,
                    abRect[2] + offsetX, abRect[3] + offsetY
                ]);
                newAb.name = (tsvData[i][nameIndex] || "Data_" + (i+1)).toString();
            }

            // 選択グループ指定時は、アートボードを増やさず同一アートボード内に複製して流し込み
            for (var j = 0; j < sourceItems.length; j++) {
                var newItem = sourceItems[j].duplicate();
                newItem.translate(offsetX, offsetY);
                replaceTextRecursive(newItem, headers, tsvData[i]);
            }
        }

        // 最後に1件目(i=0)を書き換える
        progBar.value = tsvData.length;
        progWin.update();
        if (targetMode !== "selection") {
            sourceAb.name = (tsvData[0][nameIndex] || "Data_1").toString();
        }
        for (var k = 0; k < sourceItems.length; k++) {
            replaceTextRecursive(sourceItems[k], headers, tsvData[0]);
        }

        doc.selection = null;
        progWin.close();
        alert(L("done").replace("#count#", String(tsvData.length)));
    };

    function replaceTextRecursive(obj, keys, values) {
        if (!obj) return;

        if (obj.typename === "TextFrame") {
            var txt = obj.contents;
            if (txt == null) return;

            for (var k = 0; k < keys.length; k++) {
                var key = String(keys[k]).replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
                if (key === "") continue;

                var tag = "<" + key + ">";
                if (txt.indexOf(tag) !== -1) {
                    var val = (values && values.length > k && values[k] != null) ? String(values[k]) : "";
                    // 全出現を置換（ExtendScript安全）
                    txt = txt.split(tag).join(val);
                }
            }

            obj.contents = txt;
        } else if (obj.typename === "GroupItem") {
            for (var g = 0; g < obj.pageItems.length; g++) {
                replaceTextRecursive(obj.pageItems[g], keys, values);
            }
        }
    }

    win.show();
})();
