#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：SmartBatchImporter.jsx
作成日：2025-05-29
更新日：2025-05-29
- 1.0.1：フォルダー指定で読む込むときの挙動を変更、ラベルテキストを「_label」レイヤーに
- 1.0.2：読み込むファイル数を表示、プログレスバーとキャンセル機能を追加
- 1.0.3：プログレスバーに進捗数表示（n/N）を追加

スクリプトの概要：
複数の Illustrator ファイル（.ai/.svg）を一括で読み込み、表示中かつロック解除されたオブジェクトを新規ドキュメントに貼り付けて整列配置します。
ファイル名ラベル（オプション）や拡張子表示、読み込み後に元ファイルを閉じるか保持するかなどのオプションも指定可能です。
フォルダー読み込み時は対象ファイルを1件ずつ開いて処理し、メモリ負荷を軽減します。
実行中はプログレスバーと進捗数（n/N）を表示し、キャンセルも可能です。
*/

// -------------------------------
// 日英ラベル定義 / Japanese-English label definitions
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var LABELS = {
    dialogTitle: {
        ja: "読み込みオプション",
        en: "New Document Settings"
    },
    colorSpace: {
        ja: "カラースペース",
        en: "Color Space"
    },
    rgb: {
        ja: "RGB",
        en: "RGB"
    },
    cmyk: {
        ja: "CMYK",
        en: "CMYK"
    },
    docSize: {
        ja: "ドキュメントサイズ",
        en: "Document Size"
    },
    presetCustom: {
        ja: "カスタム",
        en: "Custom"
    },
    presetA4: {
        ja: "A4（210mm × 297mm）",
        en: "A4 (210mm × 297mm)"
    },
    presetSquare: {
        ja: "スクエア（100mm × 100mm）",
        en: "Square (100mm × 100mm)"
    },
    presetFullHD: {
        ja: "フルHD（1920px × 1080px）",
        en: "Full HD (1920px × 1080px)"
    },
    width: {
        ja: "幅:",
        en: "Width:"
    },
    height: {
        ja: "高さ:",
        en: "Height:"
    },
    labelTitle: {
        ja: "ラベル",
        en: "File Name Label"
    },
    showLabel: {
        ja: "ファイル名",
        en: "Attach Label"
    },
    showExt: {
        ja: "拡張子",
        en: "Include Extension"
    },
    afterImport: {
        ja: "読み込み後の動作",
        en: "After Import"
    },
    close: {
        ja: "ファイルを閉じる",
        en: "Close"
    },
    keepOpen: {
        ja: "閉じない",
        en: "Keep Open"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    selectFolder: {
        ja: "フォルダを選択してください",
        en: "Select a folder"
    },
    noValidFile: {
        ja: "有効なファイルが見つかりませんでした。",
        en: "No valid files found."
    },
    invalidNumber: {
        ja: "数値が正しくありません。",
        en: "Invalid numeric input."
    },
    pasteFail: {
        ja: "ペーストに失敗しました：",
        en: "Paste failed: "
    },
    progressCount: {
        ja: "進捗：",
        en: "Progress: "
    },
    processing: {
        ja: "読み込み中...",
        en: "Processing..."
    },
    cancelledMessage: {
        ja: "ユーザーによって読み込みがキャンセルされました。",
        en: "Import cancelled by user."
    }
};

function positionLabelBelowGroup(label, group) {
    label.top = group.visibleBounds[3] - 10;
    label.left = group.visibleBounds[0];
}

function getOrCreateLabelLayer(doc) {
    var layer;
    try {
        layer = doc.layers.getByName("_label");
    } catch (e) {
        layer = doc.layers.add();
        layer.name = "_label";
    }
    return layer;
}

function main() {
    var lang = getCurrentLang();
    var originalDocs = [];
    var importFromFolder = false;

    if (app.documents.length <= 1) {
        importFromFolder = true;
        var folder = Folder.selectDialog(LABELS.selectFolder[lang]);
        if (!folder) return;

        var files = folder.getFiles(function(f) {
            return f instanceof File && (f.name.match(/\.ai$/i) || f.name.match(/\.svg$/i));
        }).sort(function(a, b) {
            return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
        });

        if (files.length < 1) {
            alert(LABELS.noValidFile[lang]);
            return;
        }

        originalDocs = files; // ファイル一覧を保持
    } else {
        for (var i = 0; i < app.documents.length; i++) {
            originalDocs.push(app.documents[i]);
        }
    }

    var originalDocsLength = originalDocs.length;
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    var fileCountText = dialog.add("statictext", undefined, (lang === "ja" ? "対象ファイル数：" : "Files to Import: ") + originalDocsLength);
    fileCountText.alignment = ["center", "top"];

    dialog.orientation = "column";
    dialog.alignChildren = ["left", "top"];

    var colorPanel = dialog.add("panel", undefined, LABELS.colorSpace[lang]);
    colorPanel.margins = [10, 20, 10, 10];
    colorPanel.orientation = "row";
    colorPanel.alignChildren = ["left", "center"];
    var rgbRadio = colorPanel.add("radiobutton", undefined, LABELS.rgb[lang]);
    var cmykRadio = colorPanel.add("radiobutton", undefined, LABELS.cmyk[lang]);
    rgbRadio.value = true;

    var sizePanel = dialog.add("panel", undefined, LABELS.docSize[lang]);
    sizePanel.margins = [10, 20, 10, 10];
    sizePanel.orientation = "column";
    sizePanel.alignChildren = ["left", "top"];

    var presetDropdown = sizePanel.add("dropdownlist", undefined, [
        LABELS.presetCustom[lang],
        LABELS.presetA4[lang],
        LABELS.presetSquare[lang],
        LABELS.presetFullHD[lang]
    ]);
    presetDropdown.selection = 0;

    var sizeGroup = sizePanel.add("group");
    sizeGroup.orientation = "row";
    sizeGroup.add("statictext", undefined, LABELS.width[lang]);
    var widthInput = sizeGroup.add("edittext", undefined, "1000");
    widthInput.characters = 6;
    sizeGroup.add("statictext", undefined, "  " + LABELS.height[lang]);
    var heightInput = sizeGroup.add("edittext", undefined, "1000");
    heightInput.characters = 6;

    presetDropdown.onChange = function() {
        switch (presetDropdown.selection.index) {
            case 1: // A4
                widthInput.text = Math.round(210 * 2.8346).toString();
                heightInput.text = Math.round(297 * 2.8346).toString();
                break;
            case 2: // Square 100mm
                widthInput.text = Math.round(100 * 2.8346).toString();
                heightInput.text = Math.round(100 * 2.8346).toString();
                break;
            case 3: // Full HD px
                widthInput.text = Math.round(1920).toString();
                heightInput.text = Math.round(1080).toString();
                break;
        }
    };

    var labelPanel = dialog.add("panel", undefined, LABELS.labelTitle[lang]);
    labelPanel.orientation = "row";
    labelPanel.alignChildren = ["left", "center"];
    labelPanel.margins = [10, 20, 10, 10];
    var showLabelCheckbox = labelPanel.add("checkbox", undefined, LABELS.showLabel[lang]);
    showLabelCheckbox.value = true;
    var showExtensionCheckbox = labelPanel.add("checkbox", undefined, LABELS.showExt[lang]);
    showExtensionCheckbox.value = false;
    showExtensionCheckbox.enabled = showLabelCheckbox.value;
    showLabelCheckbox.onClick = function() {
        showExtensionCheckbox.enabled = this.value;
    };

    var closePanel = dialog.add("panel", undefined, LABELS.afterImport[lang]);
    closePanel.orientation = "row";
    closePanel.alignChildren = ["left", "center"];
    closePanel.margins = [10, 20, 10, 10];
    var closeRadio = closePanel.add("radiobutton", undefined, LABELS.close[lang]);
    var keepOpenRadio = closePanel.add("radiobutton", undefined, LABELS.keepOpen[lang]);
    closeRadio.value = true;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang]);

    if (dialog.show() !== 1) return;

    // プログレスバーのダイアログを表示
    var progressWin = new Window("palette", LABELS.processing[lang]);
    progressWin.orientation = "column";
    progressWin.alignChildren = ["fill", "top"];
    progressWin.margins = 20;
    var progressTextGroup = progressWin.add("group");
    progressTextGroup.alignment = ["center", "top"];
    var processedCountStatic = progressTextGroup.add("statictext", undefined, LABELS.progressCount[lang] + "0/" + originalDocsLength);
    processedCountStatic.preferredSize = [100, 30];

    var progressBar = progressWin.add("progressbar", undefined, 0, originalDocsLength);
    progressBar.preferredSize = [300, 6];

    var cancelGroup = progressWin.add("group");
    cancelGroup.alignment = "right";
    var cancelBtn = cancelGroup.add("button", undefined, LABELS.cancel[lang]);

    var userCancelled = false;

    cancelBtn.onClick = function() {
        userCancelled = true;
    };
    progressWin.addEventListener("keydown", function(e) {
        if (e.keyName === "Escape") {
            userCancelled = true;
        }
    });
    progressWin.show();

    var docWidthPx = parseFloat(widthInput.text);
    var docHeightPx = parseFloat(heightInput.text);
    if (isNaN(docWidthPx) || isNaN(docHeightPx)) {
        alert(LABELS.invalidNumber[lang]);
        return;
    }

    const spacingX = 100;
    const spacingY = 100;

    var docWidthPt = docWidthPx * 0.75;
    var docHeightPt = docHeightPx * 0.75;
    var colorSpace = rgbRadio.value ? DocumentColorSpace.RGB : DocumentColorSpace.CMYK;

    var newDoc = app.documents.add(colorSpace, docWidthPt, docHeightPt);
    app.activeDocument = newDoc;

    var artboardRect = newDoc.artboards[0].artboardRect;
    var artboardLeft = artboardRect[0];
    var artboardTop = artboardRect[1];

    var margin = 8.5;
    var layoutLeft = artboardLeft + margin;
    var layoutRight = artboardRect[2] - margin;
    var layoutTop = artboardTop - margin;

    var rowY = layoutTop;
    var rowMaxHeight = 0;

    for (var j = 0; j < originalDocs.length; j++) {
        $.sleep(0);
        app.redraw();
        if (userCancelled) {
            progressWin.close();
            throw new Error(LABELS.cancelledMessage[lang]);
        }

        var srcDoc;
        if (importFromFolder) {
            srcDoc = app.open(originalDocs[j]);
        } else {
            srcDoc = originalDocs[j];
        }
        app.activeDocument = srcDoc;

        app.executeMenuCommand("selectall");

        var filteredSelection = [];
        for (var s = 0; s < app.selection.length; s++) {
            var obj = app.selection[s];
            if (!obj.locked && !obj.hidden) {
                filteredSelection.push(obj);
            }
        }
        if (filteredSelection.length === 0) {
            if (closeRadio.value) {
                srcDoc.close(SaveOptions.DONOTSAVECHANGES);
            }
            continue;
        }

        app.selection = filteredSelection;

        app.copy();
        app.activeDocument = newDoc;

        app.paste();

        var pastedItems = newDoc.selection;
        if (pastedItems.length === 0) {
            alert(LABELS.pasteFail[lang] + srcDoc.name);
            if (closeRadio.value) {
                srcDoc.close(SaveOptions.DONOTSAVECHANGES);
            }
            continue;
        }

        var targetLayer = newDoc.activeLayer;
        var pastedGroup = targetLayer.groupItems.add();
        for (var m = pastedItems.length - 1; m >= 0; m--) {
            pastedItems[m].moveToBeginning(pastedGroup);
        }

        newDoc.activeLayer = targetLayer; // 貼り付け直後にアクティブレイヤーを明示的に戻す

        var bounds = pastedGroup.visibleBounds;
        var width = bounds[2] - bounds[0];
        var groupHeight = bounds[1] - bounds[3];

        var dx = 0;
        var dy = 0;

        if (j === 0) {
            dx = layoutLeft - bounds[0];
            dy = rowY - bounds[1];
            rowMaxHeight = groupHeight;
        } else {
            var nextX = previousBounds[2] + spacingX;
            var projectedRight = nextX + width;

            if (projectedRight > layoutRight) {
                rowY -= rowMaxHeight + spacingY;
                rowMaxHeight = groupHeight;

                dx = layoutLeft - bounds[0];
                dy = rowY - bounds[1];
            } else {
                dx = nextX - bounds[0];
                dy = rowY - bounds[1];
                if (groupHeight > rowMaxHeight) rowMaxHeight = groupHeight;
            }
        }

        pastedGroup.translate(dx, dy);

        if (showLabelCheckbox.value) {
            var labelLayer = getOrCreateLabelLayer(newDoc);

            var label = newDoc.textFrames.add();
            label.contents = showExtensionCheckbox.value ? srcDoc.name : srcDoc.name.replace(/\.[^\.]+$/, "");
            label.textRange.characterAttributes.textFont = app.textFonts.getByName("HiraginoSans-W3");
            label.textRange.characterAttributes.size = 9;
            positionLabelBelowGroup(label, pastedGroup);
            if (label.layer != labelLayer) label.layer = labelLayer;
            label.move(labelLayer, ElementPlacement.PLACEATBEGINNING);
        }

        var previousBounds = pastedGroup.visibleBounds;

        if (closeRadio.value) {
            srcDoc.close(SaveOptions.DONOTSAVECHANGES);
        }
        progressBar.value = j + 1;
        processedCountStatic.text = LABELS.progressCount[lang] + (j + 1) + "/" + originalDocsLength;
        progressWin.update();
    }
    app.activeDocument = newDoc;
    progressWin.close();
}

main();