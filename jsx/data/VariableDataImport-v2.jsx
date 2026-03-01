/*
  VariableDataImport_v12.jsx
  仕様：
  1. 元のドキュメント内で直接アートボードを増やして流し込みます。
  2. CSV(.csv)とTSV(.txt)の両方に対応（カンマや引用符も解析）。
  3. 1件目のデータを最後に処理することで、タグ消失バグを回避。
  4. 進捗バーを表示。
*/

(function() {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var docPath;
    try {
        docPath = doc.path;
    } catch (e) {
        alert("ドキュメントを保存してから実行してください。");
        return;
    }

    var folder = new Folder(docPath);
    var files = folder.getFiles(/\.(txt|csv)$/i);
    if (files.length === 0) {
        alert("同じ階層に.txtまたは.csvファイルが見つかりませんでした。");
        return;
    }

    // --- メインダイアログ ---
    var win = new Window("dialog", "データ流し込み");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 10;
    win.margins = 15;

    var groupSelect = win.add("group");
    groupSelect.add("statictext", undefined, "ファイル:");
    var dropdown = groupSelect.add("dropdownlist", undefined, []);
    dropdown.size = [350, 25];
    for (var i = 0; i < files.length; i++) dropdown.add("item", decodeURI(files[i].name));

    var listContainer = win.add("group");
    listContainer.alignment = ["fill", "fill"];
    var listBox = null; 

    var settingsPanel = win.add("panel", undefined, "流し込み設定");
    settingsPanel.orientation = "column";
    settingsPanel.alignChildren = ["left", "top"];
    settingsPanel.margins = 15;

    var groupAbName = settingsPanel.add("group");
    groupAbName.add("statictext", undefined, "アートボード名に使用する項目:");
    var dropdownAbName = groupAbName.add("dropdownlist", undefined, []);
    dropdownAbName.size = [200, 25];

    var groupLayout = settingsPanel.add("group");
    groupLayout.add("statictext", undefined, "横に並べる数:");
    var editCols = groupLayout.add("edittext", undefined, "10");
    editCols.size = [50, 25];
    groupLayout.add("statictext", undefined, "列で折り返し");

    var groupBtn = win.add("group");
    groupBtn.alignment = "right";
    var btnCancel = groupBtn.add("button", undefined, "キャンセル", {name: "cancel"});
    var btnRun = groupBtn.add("button", undefined, "実行", {name: "ok"});

    var tsvData = [];
    var headers = [];

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

        dropdownAbName.removeAll();
        for (var h = 0; h < headers.length; h++) dropdownAbName.add("item", headers[h]);
        dropdownAbName.selection = 0;

        listBox = listContainer.add("listbox", [0, 0, 550, 180], [], {
            numberOfColumns: headers.length, showHeaders: true, columnTitles: headers
        });

        for (var i = 1; i < lines.length; i++) {
            if (lines[i] === "") continue;
            var cells = parseLine(lines[i], delimiter);
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
        win.close();

        var progWin = new Window("palette", "処理中...", undefined);
        var progBar = progWin.add("progressbar", undefined, 0, tsvData.length);
        progBar.preferredSize.width = 300;
        progWin.show();

        // 雛形の情報を取得
        var sourceAb = doc.artboards[0];
        var abRect = sourceAb.artboardRect;
        var w = abRect[2] - abRect[0];
        var h = abRect[1] - abRect[3];
        var margin = 50;

        doc.selectObjectsOnActiveArtboard();
        var sourceItems = doc.selection;
        if (sourceItems.length === 0) {
            progWin.close();
            alert("雛形にオブジェクトがありません。");
            return;
        }

        // 重要：2件目(i=1)から先に処理する（タグを維持するため）
        for (var i = 1; i < tsvData.length; i++) {
            progBar.value = i;
            progWin.update();

            var col = i % maxCols;
            var row = Math.floor(i / maxCols);
            var offsetX = (w + margin) * col;
            var offsetY = -(h + margin) * row;

            var newAb = doc.artboards.add([
                abRect[0] + offsetX, abRect[1] + offsetY,
                abRect[2] + offsetX, abRect[3] + offsetY
            ]);
            newAb.name = (tsvData[i][nameIndex] || "Data_" + (i+1)).toString();

            for (var j = 0; j < sourceItems.length; j++) {
                var newItem = sourceItems[j].duplicate();
                newItem.translate(offsetX, offsetY);
                replaceTextRecursive(newItem, headers, tsvData[i]);
            }
        }

        // 最後に1件目(i=0)を書き換える
        progBar.value = tsvData.length;
        progWin.update();
        sourceAb.name = (tsvData[0][nameIndex] || "Data_1").toString();
        for (var k = 0; k < sourceItems.length; k++) {
            replaceTextRecursive(sourceItems[k], headers, tsvData[0]);
        }

        doc.selection = null;
        progWin.close();
        alert("完了！\n" + tsvData.length + " 件処理しました。");
    };

    function replaceTextRecursive(obj, keys, values) {
        if (obj.typename === "TextFrame") {
            for (var k = 0; k < keys.length; k++) {
                var tag = "<" + keys[k] + ">";
                if (obj.contents.indexOf(tag) !== -1) {
                    obj.contents = obj.contents.replace(tag, values[k] || "");
                }
            }
        } else if (obj.typename === "GroupItem") {
            for (var g = 0; g < obj.pageItems.length; g++) {
                replaceTextRecursive(obj.pageItems[g], keys, values);
            }
        }
    }

    win.show();
})();
