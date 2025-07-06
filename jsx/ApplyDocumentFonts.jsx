#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ApplyDocumentFonts.jsx

### 概要

- ドキュメント内で使用されているフォントを集計し、使用数順に一覧表示するスクリプトです。
- 選択テキストに即座にフォントを適用でき、一覧はテキストファイルとして書き出し可能です。

### 主な機能

- ドキュメント内のフォントを使用数順に収集・表示
- 検索フィルターによるフォント絞り込み
- 選択中のテキストオブジェクトに選択したフォントを即時適用
- フォント一覧をデスクトップにテキストファイルとして書き出し
- 日本語／英語インターフェース対応

### 処理の流れ

1. ドキュメントからフォント情報を収集
2. ダイアログで一覧を表示、検索フィルターで絞り込み
3. フォントを選択すると即座にテキストに適用
4. 必要に応じてフォント一覧をテキストファイルに書き出し

### 更新履歴

- v1.0.0 (20250225) : 初期バージョン
- v1.1.0 (20250228) : 書き出し機能を追加
- v1.1.1 (20250301) : グループ内テキストへの適用に対応
- v1.1.2 (20250302) : グループ内フォント数のカウント方法を調整

---

### Script Name:

ApplyDocumentFonts.jsx

### Overview

- A script that collects fonts used in the document and displays them sorted by usage count.
- You can immediately apply a selected font to text objects and export the list as a text file.

### Main Features

- Collect and display document fonts sorted by usage count
- Filter fonts using a search filter
- Instantly apply selected font to currently selected text objects
- Export font list as a text file on the desktop
- Japanese and English UI support

### Process Flow

1. Collect font information from the document
2. Display fonts in a dialog, filterable via search
3. Apply font instantly when selected
4. Optionally export font list to a text file

### Update History

- v1.0.0 (20250225): Initial version
- v1.1.0 (20250228): Added export feature
- v1.1.1 (20250301): Supported applying to text inside groups
- v1.1.2 (20250302): Adjusted font count method for groups
*/

(function() {
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();
    var LABELS = {
        dialogTitle: {
            ja: "ドキュメントフォントを適用",
            en: "Apply Document Fonts"
        },
        searchLabel: {
            ja: "検索フィルター（フォント名・スタイル名・PostScript名に対応）",
            en: "Search Filter (Font Name, Style, PostScript Name)"
        },
        listLabel: {
            ja: "ドキュメントフォント（使用数順）",
            en: "Document Fonts (by Usage)"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        exportLabel: {
            ja: "書き出し",
            en: "Export"
        },
        noDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        errorApplyFont: {
            ja: "フォントの適用に失敗しました。",
            en: "Failed to apply the font."
        },
        exportSuccess: {
            ja: "書き出しました：",
            en: "Font list saved:\n"
        },
        exportFail: {
            ja: "ファイルの書き出しに失敗しました。",
            en: "Failed to save the file."
        },
        exportHeader: {
            ja: "Adobe Illustrator ドキュメント情報",
            en: "Adobe Illustrator Document Info"
        },
        exportDocLabel: {
            ja: "ドキュメント : ",
            en: "Document: "
        },
        exportFontCount: {
            ja: "このドキュメントに使用されているフォント : ",
            en: "Fonts used in this document: "
        }
    };

    function main() {
        if (app.documents.length === 0) {
            alert(LABELS.noDocument[lang]);
            return;
        }

        var doc = app.activeDocument;
        var fontMap = {};
        var currentFontDisplayName = null;
        var originalFonts = [];

        function getFontDisplayName(textFont) {
            return textFont.style ? (textFont.family + " - " + textFont.style) : textFont.family;
        }

        function collectFontsFromTextFrame(textFrame) {
            if (!textFrame.contents || textFrame.contents.length === 0) return;
            try {
                var textFont = textFrame.textRange.characterAttributes.textFont;
                var displayName = getFontDisplayName(textFont);
                if (!fontMap[displayName]) {
                    fontMap[displayName] = {
                        count: 1,
                        postScriptName: textFont.name
                    };
                } else {
                    fontMap[displayName].count++;
                }
            } catch (e) {}
        }

        function collectFontsFromGroup(groupItem) {
            for (var i = 0; i < groupItem.pageItems.length; i++) {
                var item = groupItem.pageItems[i];
                if (item.typename === "TextFrame") {
                    collectFontsFromTextFrame(item);
                } else if (item.typename === "GroupItem") {
                    collectFontsFromGroup(item);
                }
            }
        }

        // グループ再帰走査を削除し、textFramesのみでフォント情報を収集
        for (var i = 0; i < doc.textFrames.length; i++) {
            collectFontsFromTextFrame(doc.textFrames[i]);
        }

        for (var s = 0; s < app.selection.length; s++) {
            if (app.selection[s].typename === "TextFrame") {
                try {
                    var selFont = app.selection[s].textRange.characterAttributes.textFont;
                    currentFontDisplayName = getFontDisplayName(selFont);
                    originalFonts.push({
                        item: app.selection[s],
                        font: selFont
                    });
                } catch (e) {}
            }
        }

        var sortedFonts = [];
        for (var name in fontMap) {
            if (fontMap.hasOwnProperty(name)) {
                sortedFonts.push({
                    displayName: name,
                    postScriptName: fontMap[name].postScriptName
                });
            }
        }
        sortedFonts.sort(function(a, b) {
            return a.displayName.localeCompare(b.displayName);
        });

        function getListBoxHeight(fontList) {
            return Math.min(300, Math.max(100, fontList.length * 20));
        }

        var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
        dialog.orientation = "column";
        dialog.alignChildren = ["left", "top"];
        dialog.margins = 20;

        dialog.add("statictext", undefined, LABELS.searchLabel[lang]);
        var filterInput = dialog.add("edittext", undefined, "");
        filterInput.preferredSize = [400, 24];

        dialog.add("statictext", undefined, LABELS.listLabel[lang]);
        var listBox = dialog.add("listbox", undefined, [], { multiselect: false });
        listBox.preferredSize = [400, getListBoxHeight(sortedFonts)];

        function updateListBox(filterText) {
            listBox.removeAll();
            var filter = filterText.toLowerCase();
            for (var i = 0; i < sortedFonts.length; i++) {
                var font = sortedFonts[i];
                if (filter === "" || font.displayName.toLowerCase().indexOf(filter) !== -1 || font.postScriptName.toLowerCase().indexOf(filter) !== -1) {
                    var count = fontMap[font.displayName].count;
                    var label = font.displayName + " (" + count + ")";
                    var item = listBox.add("item", label);
                    item.fontDisplayName = font.displayName;
                    item.postScriptName = font.postScriptName;
                }
            }
            if (listBox.items.length > 0) listBox.selection = 0;
        }

        updateListBox("");

        filterInput.onChanging = function() {
            updateListBox(filterInput.text);
        };

        // 再帰的にTextFrameへフォントを適用する関数
        function applyFontToItem(item, targetFont) {
            if (item.typename === "TextFrame") {
                item.textRange.characterAttributes.textFont = targetFont;
            } else if (item.typename === "GroupItem") {
                for (var i = 0; i < item.pageItems.length; i++) {
                    applyFontToItem(item.pageItems[i], targetFont);
                }
            } else if (item.typename === "CompoundPathItem" && item.pageItems && item.pageItems.length > 0) {
                for (var i = 0; i < item.pageItems.length; i++) {
                    applyFontToItem(item.pageItems[i], targetFont);
                }
            } else if (item.typename === "PathItem" || item.typename === "MeshItem") {
                // 何もしない
            } else if (item.typename === "ClipGroup" || (item.clipped && item.pageItems)) {
                for (var i = 0; i < item.pageItems.length; i++) {
                    applyFontToItem(item.pageItems[i], targetFont);
                }
            }
        }

        listBox.onChange = function() {
            if (!listBox.selection) return;
            try {
                var targetFont = app.textFonts.getByName(listBox.selection.postScriptName);
                for (var i = 0; i < app.selection.length; i++) {
                    applyFontToItem(app.selection[i], targetFont);
                }
                app.redraw();
            } catch (e) {
                alert(LABELS.errorApplyFont[lang]);
            }
        };

        var outerGroup = dialog.add("group");
        outerGroup.orientation = "row";
        outerGroup.alignChildren = ["fill", "center"];
        outerGroup.alignment = "fill";
        outerGroup.margins = [0, 10, 0, 0];
        outerGroup.spacing = 0;

        var leftGroup = outerGroup.add("group");
        leftGroup.orientation = "row";
        leftGroup.alignChildren = "left";
        leftGroup.alignment = ["left", "center"];
        var exportButton = leftGroup.add("button", undefined, LABELS.exportLabel[lang]);

        var spacer = outerGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 10;

        var rightGroup = outerGroup.add("group");
        rightGroup.orientation = "row";
        rightGroup.alignChildren = ["right", "center"];
        rightGroup.alignment = ["right", "center"];
        rightGroup.spacing = 10;
        var cancelButton = rightGroup.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
        var okButton = rightGroup.add("button", undefined, LABELS.ok[lang], { name: "ok" });

        okButton.onClick = function() {
            dialog.close();
        };

        cancelButton.onClick = function() {
            for (var i = 0; i < originalFonts.length; i++) {
                try {
                    originalFonts[i].item.textRange.characterAttributes.textFont = originalFonts[i].font;
                } catch (e) {}
            }
            app.redraw();
            dialog.close();
        };

        exportButton.onClick = function() {
            var docPath = doc.fullName.fsName;
            var docName = doc.name.replace(/\.[^\.]+$/, "");
            var defaultFileName = docName + ((lang === "ja") ? "-ドキュメントフォント一覧.txt" : "-Document-Font-List.txt");

            var desktopFolder = Folder.desktop;
            var saveFile = new File(desktopFolder.fsName + "/" + defaultFileName);

            var fontNames = [];
            for (var i = 0; i < listBox.items.length; i++) {
                fontNames.push(listBox.items[i].fontDisplayName);
            }

            var output = "";
            output += LABELS.exportHeader[lang] + "\n\n";
            output += LABELS.exportDocLabel[lang] + docPath + "\n\n";
            output += LABELS.exportFontCount[lang] + fontNames.length + "\n\n";
            output += fontNames.join("\n") + "\n";

            try {
                saveFile.encoding = "UTF-8";
                saveFile.open("w");
                saveFile.write(output);
                saveFile.close();
                alert(LABELS.exportSuccess[lang] + saveFile.fsName);
            } catch (e) {
                alert(LABELS.exportFail[lang]);
            }
        };

        dialog.show();
        listBox.active = true;
    }

    main();
})();